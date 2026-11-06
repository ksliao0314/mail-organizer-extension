// Outlook REST wrapper regressions (2026-11-05 audit):
//   item 1  — folder listings follow @odata.nextLink (child folders 201+ were lost)
//   item 2  — well-known non-filing folders' subtrees leave fetchFolderTree
//   item 13 — no default $orderby next to a non-ReceivedDateTime $filter
//   item 14 — abort while a POST is on the wire = uncertain, not plain cancel
//   item 17 — delete = move into Deleted Items
//   item 6  — mailbox-wide modified-since listing for the harvester

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  OutlookApi,
  OutlookError,
  OutlookTokenUnavailableError,
  filterLeadsWithReceivedDateTime,
  flattenAiVisibleFolders,
} from '@/shared/outlook-api'
import { TokenBucket } from '@/shared/throttle'
import type { MailFolderNode } from '@/shared/types'

let fetchSpy: ReturnType<typeof vi.fn>

beforeEach(() => {
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
})
afterEach(() => vi.restoreAllMocks())

const BASE = 'https://outlook.office.com/api/v2.0'
function makeApi(): OutlookApi {
  // Generous private bucket: these tests issue many requests per file.
  return new OutlookApi(async () => 'tok', undefined, { bucket: new TokenBucket(1000, 1000) })
}
function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status })
}
function folder(id: string, name: string, childCount = 0, parent = 'root') {
  return { Id: id, DisplayName: name, ParentFolderId: parent, ChildFolderCount: childCount }
}
function urlOf(call: unknown[]): string {
  return String(call[0])
}

describe('folder listings follow @odata.nextLink (item 1)', () => {
  it('listChildFolders returns every page, not just the first 200', async () => {
    const page1 = Array.from({ length: 200 }, (_, i) => folder(`c${i}`, `案件${i}`, 0, 'p'))
    const page2 = Array.from({ length: 200 }, (_, i) => folder(`c${200 + i}`, `案件${200 + i}`, 0, 'p'))
    const page3 = [folder('c400', '案件400', 0, 'p')]
    fetchSpy
      .mockResolvedValueOnce(
        json({ value: page1, '@odata.nextLink': `${BASE}/me/MailFolders/p/childFolders?$top=200&$skip=200` }),
      )
      .mockResolvedValueOnce(
        json({ value: page2, '@odata.nextLink': `${BASE}/me/MailFolders/p/childFolders?$top=200&$skip=400` }),
      )
      .mockResolvedValueOnce(json({ value: page3 }))
    const out = await makeApi().listChildFolders('p')
    expect(out).toHaveLength(401)
    expect(out.at(-1)!.Id).toBe('c400')
    expect(urlOf(fetchSpy.mock.calls[1]!)).toBe(`${BASE}/me/MailFolders/p/childFolders?$top=200&$skip=200`)
  })

  it('never sends the bearer token to a nextLink on a foreign host', async () => {
    fetchSpy.mockResolvedValueOnce(
      json({ value: [folder('a', 'A')], '@odata.nextLink': 'https://evil.example/api/v2.0/me/MailFolders?$skip=1' }),
    )
    await expect(makeApi().listTopFolders()).rejects.toBeInstanceOf(OutlookError)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it('a server echoing the same nextLink is an error, never a silent truncation', async () => {
    const loop = `${BASE}/me/MailFolders?$top=100&$skip=100`
    fetchSpy.mockImplementation(async () => json({ value: [folder('a', 'A')], '@odata.nextLink': loop }))
    await expect(makeApi().listTopFolders()).rejects.toThrow(/分頁未完成/)
  })
})

// O1 (2026-11-06 r2). v1.0.8 hid Deleted Items, Junk, Drafts, Outbox and Sent
// Items with their whole subtrees — rules filing into Junk / Sent Items (user
// rules included) got silently orphaned. Round 1 brought Junk / Sent back but
// kept the trash in the tree and LISTED its whole subtree (and Drafts') on
// every fetch, and one failed lookup listed everything, cached for a day.
describe('fetchFolderTree: trash / Drafts / Outbox left out unlisted; Junk / Sent tagged (O1)', () => {
  const DISPLAY: Record<string, string> = {
    deleteditems: '刪除的郵件',
    drafts: '草稿',
    outbox: '寄件匣',
    junkemail: '垃圾郵件',
    sentitems: '寄件備份',
  }
  function routeTree(opts: { wellKnownStatus?: (role: string) => number | undefined } = {}) {
    fetchSpy.mockImplementation(async (input: unknown) => {
      const url = String(input)
      const wk = url.match(/\/me\/MailFolders\/(deleteditems|junkemail|drafts|outbox|sentitems)\?/)
      if (wk) {
        const status = opts.wellKnownStatus?.(wk[1]!)
        if (status) return new Response('nope', { status })
        return json({ Id: `WK-${wk[1]}`, DisplayName: DISPLAY[wk[1]!] })
      }
      if (url.includes('/me/MailFolders?$top=100')) {
        return json({
          value: [
            folder('INBOX', '收件匣'),
            folder('CASES', '案件', 1),
            folder('WK-deleteditems', '刪除的郵件', 1),
            folder('WK-sentitems', '寄件備份', 1),
            folder('WK-junkemail', '垃圾郵件'),
            folder('WK-drafts', '草稿', 1),
            folder('WK-outbox', '寄件匣'),
          ],
        })
      }
      if (url.includes('/MailFolders/CASES/childFolders')) {
        return json({ value: [folder('JIA', '甲', 0, 'CASES')] })
      }
      if (url.includes('/MailFolders/WK-deleteditems/childFolders')) {
        // The deleted case folder keeps its id inside Deleted Items.
        return json({ value: [folder('OLD-CASE', '乙', 1, 'WK-deleteditems')] })
      }
      if (url.includes('/MailFolders/OLD-CASE/childFolders')) {
        return json({ value: [folder('OLD-CASE-SUB', '函', 0, 'OLD-CASE')] })
      }
      if (url.includes('/MailFolders/WK-sentitems/childFolders')) {
        return json({ value: [folder('SENT-A', '客戶A', 0, 'WK-sentitems')] })
      }
      if (url.includes('/MailFolders/WK-drafts/childFolders')) {
        return json({ value: [folder('DRAFT-SUB', '範本', 0, 'WK-drafts')] })
      }
      throw new Error(`unexpected ${url}`)
    })
  }
  const pathsOf = (tree: Awaited<ReturnType<OutlookApi['fetchFolderTree']>>) =>
    tree.flatMap((n) => [n.path, ...n.children.map((c) => c.path)])
  const listed = (re: RegExp) => fetchSpy.mock.calls.some((c) => re.test(urlOf(c)))

  it('leaves the trash, Drafts and Outbox out WITHOUT listing their subtrees; tags Junk / Sent', async () => {
    routeTree()
    const { tree, wellKnown } = await makeApi().fetchFolderTreeWithMeta()
    expect(pathsOf(tree)).toEqual(['收件匣', '案件', '案件/甲', '寄件備份', '寄件備份/客戶A', '垃圾郵件'])
    expect(tree.find((n) => n.id === 'WK-sentitems')!.wellKnown).toBe('sentitems')
    expect(tree.find((n) => n.id === 'WK-junkemail')!.wellKnown).toBe('junkemail')
    expect(tree.find((n) => n.id === 'CASES')!.wellKnown).toBeUndefined()
    expect(wellKnown).toEqual({
      complete: true,
      roots: [
        { id: 'WK-deleteditems', role: 'deleteditems', displayName: '刪除的郵件' },
        { id: 'WK-drafts', role: 'drafts', displayName: '草稿' },
        { id: 'WK-outbox', role: 'outbox', displayName: '寄件匣' },
        { id: 'WK-junkemail', role: 'junkemail', displayName: '垃圾郵件' },
        { id: 'WK-sentitems', role: 'sentitems', displayName: '寄件備份' },
      ],
    })
    // (d) Not one request for an omitted subtree.
    expect(listed(/WK-deleteditems\/childFolders|WK-drafts\/childFolders|OLD-CASE/)).toBe(false)
  })

  it('(e) a TRANSIENT lookup failure: what resolved is still used, the result marked incomplete', async () => {
    routeTree({ wellKnownStatus: (role) => (role === 'deleteditems' ? 401 : undefined) })
    const { tree, wellKnown } = await makeApi().fetchFolderTreeWithMeta()
    expect(wellKnown.complete).toBe(false)
    expect(wellKnown.roots.map((r) => r.role)).toEqual(['drafts', 'outbox', 'junkemail', 'sentitems'])
    // The unrecognised trash is listed like any folder; Drafts / Outbox are
    // still left out and Junk / Sent still tagged — no all-or-nothing.
    expect(pathsOf(tree)).toContain('刪除的郵件/乙')
    expect(pathsOf(tree)).not.toContain('草稿')
    expect(tree.find((n) => n.id === 'WK-junkemail')!.wellKnown).toBe('junkemail')
  })

  it('a STABLE 4xx for a name is a definitive "no such folder": complete, nothing special about it', async () => {
    routeTree({ wellKnownStatus: (role) => (role === 'outbox' ? 400 : undefined) })
    const { tree, wellKnown } = await makeApi().fetchFolderTreeWithMeta()
    expect(wellKnown.complete).toBe(true)
    expect(wellKnown.roots.map((r) => r.role)).not.toContain('outbox')
    expect(pathsOf(tree)).toContain('寄件匣')
    expect(pathsOf(tree)).not.toContain('刪除的郵件')
  })

  // O1 r3 (2026-11-06): a tree built with NO root recognised listed the trash
  // subtree and Drafts, left Junk / Sent untagged — and the same preflight
  // handed it to the AI (an AI move to 垃圾郵件 resolved as move/ai, learned).
  it('every lookup failing fails the fetch — no unrecognised tree is built', async () => {
    routeTree({ wellKnownStatus: () => 401 })
    await expect(makeApi().fetchFolderTreeWithMeta()).rejects.toBeInstanceOf(OutlookError)
    expect(listed(/\/me\/MailFolders\?\$top=100/)).toBe(false)
  })

  it('…and a token outage surfaces as itself (the caller serves its cache quietly)', async () => {
    routeTree({ wellKnownStatus: () => 401 })
    const api = new OutlookApi(
      async (opts) => {
        if (opts?.force) throw Object.assign(new Error('請先打開並登入 outlook.office.com 分頁'), { code: 'NOT_ON_OWA' })
        return 'tok'
      },
      undefined,
      { bucket: new TokenBucket(1000, 1000) },
    )
    await expect(api.fetchFolderTreeWithMeta()).rejects.toBeInstanceOf(OutlookTokenUnavailableError)
  })

  it('a failed role keeps the previous fetch’s root while it is still at the top level (sticky)', async () => {
    routeTree({ wellKnownStatus: (role) => (role === 'junkemail' || role === 'deleteditems' ? 401 : undefined) })
    const previous = {
      complete: true,
      roots: [
        { id: 'WK-deleteditems', role: 'deleteditems' as const, displayName: '刪除的郵件' },
        { id: 'WK-junkemail', role: 'junkemail' as const },
      ],
    }
    const { tree, wellKnown } = await makeApi().fetchFolderTreeWithMeta({ previous })
    expect(wellKnown.complete).toBe(true)
    expect(wellKnown.roots.map((r) => r.role)).toEqual(['deleteditems', 'drafts', 'outbox', 'junkemail', 'sentitems'])
    expect(tree.find((n) => n.id === 'WK-junkemail')!.wellKnown).toBe('junkemail')
    expect(pathsOf(tree)).not.toContain('刪除的郵件')
    expect(listed(/WK-deleteditems\/childFolders/)).toBe(false)
  })

  it('…but never another mailbox’s root (not at this top level): incomplete, listed as is', async () => {
    routeTree({ wellKnownStatus: (role) => (role === 'junkemail' ? 401 : undefined) })
    const previous = { complete: true, roots: [{ id: 'OTHER-ACCOUNT-JUNK', role: 'junkemail' as const }] }
    const { tree, wellKnown } = await makeApi().fetchFolderTreeWithMeta({ previous })
    expect(wellKnown.complete).toBe(false)
    expect(wellKnown.roots.map((r) => r.id)).not.toContain('OTHER-ACCOUNT-JUNK')
    expect(tree.find((n) => n.id === 'WK-junkemail')!.wellKnown).toBeUndefined()
  })
})

describe('flattenAiVisibleFolders (O1c)', () => {
  it('drops the Junk / Sent subtrees, keeps everything else', () => {
    const node = (
      id: string,
      path: string,
      children: MailFolderNode[] = [],
      extra: Partial<MailFolderNode> = {},
    ): MailFolderNode => ({ id, displayName: path.split('/').at(-1)!, path, children, ...extra })
    const tree: MailFolderNode[] = [
      node('C', '案件', [node('J', '案件/甲')]),
      node('S', '寄件備份', [node('SA', '寄件備份/客戶A')], { wellKnown: 'sentitems' }),
      node('K', '垃圾郵件', [], { wellKnown: 'junkemail' }),
    ]
    expect(flattenAiVisibleFolders(tree).map((n) => n.path)).toEqual(['案件', '案件/甲'])
  })
})

describe('listFolderMessages default $orderby (item 13)', () => {
  it('omits the default ReceivedDateTime orderby next to a From/... filter (InefficientFilter)', async () => {
    fetchSpy.mockResolvedValueOnce(json({ value: [] }))
    await makeApi().listFolderMessages('fid', { filter: "From/EmailAddress/Address eq 'a@b.example'" })
    const url = decodeURIComponent(urlOf(fetchSpy.mock.calls[0]!))
    expect(url).toContain("$filter=From/EmailAddress/Address eq 'a@b.example'")
    expect(url).not.toContain('$orderby')
  })

  it('keeps the default orderby without a filter (inbox scans unchanged)', async () => {
    fetchSpy.mockResolvedValueOnce(json({ value: [] }))
    await makeApi().listFolderMessages('inbox', { top: 5 })
    expect(decodeURIComponent(urlOf(fetchSpy.mock.calls[0]!))).toContain('$orderby=ReceivedDateTime desc')
  })

  it('keeps it when the filter leads with ReceivedDateTime; an explicit orderby always wins', async () => {
    expect(filterLeadsWithReceivedDateTime('ReceivedDateTime ge 2026-01-01T00:00:00Z')).toBe(true)
    expect(filterLeadsWithReceivedDateTime("from/x eq 'y' and ReceivedDateTime ge 2026-01-01")).toBe(false)
    fetchSpy.mockResolvedValueOnce(json({ value: [] }))
    await makeApi().listFolderMessages('fid', {
      filter: 'LastModifiedDateTime gt 2026-01-01T00:00:00Z',
      orderby: 'LastModifiedDateTime desc',
    })
    expect(decodeURIComponent(urlOf(fetchSpy.mock.calls[0]!))).toContain('$orderby=LastModifiedDateTime desc')
  })
})

describe('delete = move into Deleted Items (item 17)', () => {
  it('POSTs /move with the well-known DeletedItems destination — never DELETE', async () => {
    fetchSpy.mockResolvedValueOnce(json({ Id: 'new-id', ParentFolderId: 'DELETED-ID' }))
    const res = await makeApi().moveToDeletedItems('msg-1')
    expect(res).toEqual({ Id: 'new-id', ParentFolderId: 'DELETED-ID' })
    const [url, init] = fetchSpy.mock.calls[0]!
    expect(String(url)).toBe(`${BASE}/me/messages/msg-1/move`)
    expect((init as RequestInit).method).toBe('POST')
    expect(JSON.parse(String((init as RequestInit).body))).toEqual({ DestinationId: 'DeletedItems' })
  })

  it('a 400 on the well-known name retries ONCE with the resolved folder id', async () => {
    fetchSpy.mockImplementation(async (input: unknown, init?: RequestInit) => {
      const url = String(input)
      if (url.endsWith('/move')) {
        const dest = JSON.parse(String(init?.body)).DestinationId
        return dest === 'DeletedItems'
          ? new Response('ErrorInvalidIdMalformed', { status: 400 })
          : json({ Id: 'new-id' })
      }
      const wk = url.match(/\/me\/MailFolders\/(\w+)\?/)
      if (wk) return json({ Id: `WK-${wk[1]}` })
      throw new Error(url)
    })
    const res = await makeApi().moveToDeletedItems('msg-1')
    expect(res.Id).toBe('new-id')
    const moves = fetchSpy.mock.calls.filter((c) => urlOf(c).endsWith('/move'))
    expect(moves.map((c) => JSON.parse(String((c[1] as RequestInit).body)).DestinationId)).toEqual([
      'DeletedItems',
      'WK-deleteditems',
    ])
  })
})

describe('cancel while a request is on the wire (item 14)', () => {
  function hangUntilAbort() {
    fetchSpy.mockImplementation(
      (_input: unknown, init?: RequestInit) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')))
        }),
    )
  }

  it('POST /move aborted mid-flight → OutlookError uncertain + aborted (may have committed)', async () => {
    hangUntilAbort()
    const ctl = new AbortController()
    const p = makeApi().moveMessage('msg-1', 'fid-1234567890123456789', ctl.signal)
    await vi.waitFor(() => expect(fetchSpy).toHaveBeenCalledTimes(1))
    ctl.abort()
    await expect(p).rejects.toMatchObject({ uncertain: true, aborted: true })
  })

  it('GET aborted mid-flight stays a plain AbortError (nothing to be uncertain about)', async () => {
    hangUntilAbort()
    const ctl = new AbortController()
    const p = makeApi().listFolderMessages('inbox', { signal: ctl.signal })
    await vi.waitFor(() => expect(fetchSpy).toHaveBeenCalledTimes(1))
    ctl.abort()
    await expect(p).rejects.toMatchObject({ name: 'AbortError' })
  })

  it('an abort BEFORE the POST is sent is a plain AbortError (definitely not sent)', async () => {
    const ctl = new AbortController()
    ctl.abort()
    await expect(makeApi().moveMessage('m', 'f', ctl.signal)).rejects.toMatchObject({ name: 'AbortError' })
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})

describe('listMessagesModifiedSince (harvester, item 6; keyset O5)', () => {
  type Row = { Id: string; LastModifiedDateTime: string; ParentFolderId?: string }
  const at = (sec: number) => new Date(Date.UTC(2026, 10, 2, 0, 0, sec)).toISOString().replace(/\.\d{3}Z$/, 'Z')
  /**
   * A server honouring `LastModifiedDateTime gt|ge X`, ascending order (ties in
   * insertion order), $top and $skip — like Exchange. `rows` may be mutated
   * between requests (a message modified mid-scan).
   */
  function serve(rows: Row[], onRequest: (n: number) => void = () => {}) {
    let n = 0
    fetchSpy.mockImplementation(async (input: unknown) => {
      onRequest(n++)
      const url = decodeURIComponent(String(input))
      const f = /LastModifiedDateTime (gt|ge) (\S+?)(&|$)/.exec(url)!
      const bound = Date.parse(f[2]!)
      const top = Number(/\$top=(\d+)/.exec(url)![1])
      const skip = Number(/\$skip=(\d+)/.exec(url)?.[1] ?? 0)
      const value = rows
        .map((r, i) => ({ r, i }))
        .filter(({ r }) => {
          const t = Date.parse(r.LastModifiedDateTime)
          return f[1] === 'gt' ? t > bound : t >= bound
        })
        .sort((a, b) => Date.parse(a.r.LastModifiedDateTime) - Date.parse(b.r.LastModifiedDateTime) || a.i - b.i)
        .slice(skip, skip + top)
        .map(({ r }) => ({ ...r }))
      // Exchange's own continuation: the same query, $skip-based — the offset
      // paging that loses a row when one moves mid-scan. Never followed now.
      const nextLink =
        `${BASE}/me/messages?$filter=${encodeURIComponent(`LastModifiedDateTime ${f[1]} ${f[2]}`)}` +
        `&$orderby=${encodeURIComponent('LastModifiedDateTime asc')}&$top=${top}&$skip=${skip + top}`
      return json({ value, ...(value.length === top ? { '@odata.nextLink': nextLink } : {}) })
    })
  }

  it('one mailbox-wide query: filter + orderby on LastModifiedDateTime, ParentFolderId selected, paged', async () => {
    serve([
      { Id: 'm1', LastModifiedDateTime: at(1), ParentFolderId: 'f1' },
      { Id: 'm2', LastModifiedDateTime: at(2), ParentFolderId: 'f2' },
    ])
    const r = await makeApi().listMessagesModifiedSince('2026-11-01T00:00:00Z', { pageSize: 1 })
    expect(r.complete).toBe(true)
    expect(r.messages.map((m) => m.Id)).toEqual(['m1', 'm2'])
    const first = decodeURIComponent(urlOf(fetchSpy.mock.calls[0]!))
    expect(first).toContain('/me/messages?')
    expect(first).toContain('$filter=LastModifiedDateTime gt 2026-11-01T00:00:00Z')
    // Oldest first (C7, 2026-11-06): a page-capped run covers a prefix of the
    // window, so the next run can resume where it stopped.
    expect(first).toContain('$orderby=LastModifiedDateTime asc')
    expect(first).toContain('ParentFolderId')
    // Keyset: the next page starts at the previous page's last timestamp.
    expect(decodeURIComponent(urlOf(fetchSpy.mock.calls[1]!))).toContain(`LastModifiedDateTime ge ${at(1)}`)
  })

  it('a message modified mid-scan no longer makes the next page skip a row', async () => {
    const rows: Row[] = ['a', 'b', 'c', 'd', 'e'].map((id, i) => ({ Id: id, LastModifiedDateTime: at(i + 1), ParentFolderId: 'old' }))
    // After the first page, `a` is moved (its LastModifiedDateTime jumps to the end).
    serve(rows, (n) => {
      if (n === 1) rows[0] = { Id: 'a', LastModifiedDateTime: at(9), ParentFolderId: 'new' }
    })
    const r = await makeApi().listMessagesModifiedSince(at(0), { pageSize: 2 })
    expect(r.complete).toBe(true)
    // $skip paging read [a,b] then [d,e] — `c` was never seen.
    expect(r.messages.map((m) => m.Id).sort()).toEqual(['a', 'b', 'c', 'd', 'e'])
    // De-duplicated, the latest copy winning (it says where the mail is now).
    expect(r.messages.find((m) => m.Id === 'a')!.ParentFolderId).toBe('new')
  })

  it('a page entirely within one timestamp still makes progress (no endless re-read)', async () => {
    const rows: Row[] = [
      ...Array.from({ length: 5 }, (_, i) => ({ Id: `t${i}`, LastModifiedDateTime: at(5) })),
      { Id: 'later', LastModifiedDateTime: at(6) },
    ]
    serve(rows)
    const r = await makeApi().listMessagesModifiedSince(at(0), { pageSize: 2 })
    expect(r.complete).toBe(true)
    expect(r.messages.map((m) => m.Id)).toEqual(['t0', 't1', 't2', 't3', 't4', 'later'])
    expect(r.pages).toBeLessThanOrEqual(5)
  })

  it('reports complete=false when the page cap stops the walk', async () => {
    serve(Array.from({ length: 50 }, (_, i) => ({ Id: `m${i}`, LastModifiedDateTime: at(i + 1) })))
    const r = await makeApi().listMessagesModifiedSince(at(0), { pageSize: 2, maxPages: 3 })
    expect(r.complete).toBe(false)
    expect(r.pages).toBe(3)
    expect(fetchSpy).toHaveBeenCalledTimes(3)
  })
})
