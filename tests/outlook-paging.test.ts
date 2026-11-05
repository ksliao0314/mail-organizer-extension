// Outlook REST wrapper regressions (2026-11-05 audit):
//   item 1  — folder listings follow @odata.nextLink (child folders 201+ were lost)
//   item 2  — well-known non-filing folders' subtrees leave fetchFolderTree
//   item 13 — no default $orderby next to a non-ReceivedDateTime $filter
//   item 14 — abort while a POST is on the wire = uncertain, not plain cancel
//   item 17 — delete = move into Deleted Items
//   item 6  — mailbox-wide modified-since listing for the harvester

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { OutlookApi, OutlookError, filterLeadsWithReceivedDateTime } from '@/shared/outlook-api'
import { TokenBucket } from '@/shared/throttle'

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

describe('fetchFolderTree drops well-known non-filing subtrees (item 2)', () => {
  function routeTree(opts: { failWellKnown?: boolean } = {}) {
    fetchSpy.mockImplementation(async (input: unknown) => {
      const url = String(input)
      const wk = url.match(/\/me\/MailFolders\/(deleteditems|junkemail|drafts|outbox|sentitems)\?/)
      if (wk) {
        if (opts.failWellKnown) return new Response('nope', { status: 400 })
        return json({ Id: `WK-${wk[1]}` })
      }
      if (url.includes('/me/MailFolders?$top=100')) {
        return json({
          value: [
            folder('INBOX', '收件匣'),
            folder('CASES', '案件', 1),
            folder('WK-deleteditems', '刪除的郵件', 1),
            folder('WK-sentitems', '寄件備份'),
          ],
        })
      }
      if (url.includes('/MailFolders/CASES/childFolders')) {
        return json({ value: [folder('JIA', '甲', 0, 'CASES')] })
      }
      if (url.includes('/MailFolders/WK-deleteditems/childFolders')) {
        // The deleted case folder keeps its id inside Deleted Items.
        return json({ value: [folder('OLD-CASE', '乙', 0, 'WK-deleteditems')] })
      }
      throw new Error(`unexpected ${url}`)
    })
  }

  it('excludes Deleted Items / Sent Items subtrees and reports their ids', async () => {
    routeTree()
    const { tree, excludedFolderIds } = await makeApi().fetchFolderTreeWithMeta()
    const paths = tree.flatMap((n) => [n.path, ...n.children.map((c) => c.path)])
    expect(paths).toEqual(['收件匣', '案件', '案件/甲'])
    expect(excludedFolderIds.sort()).toEqual(
      ['WK-deleteditems', 'WK-drafts', 'WK-junkemail', 'WK-outbox', 'WK-sentitems'].sort(),
    )
    // Never even listed the trash's children.
    expect(fetchSpy.mock.calls.some((c) => urlOf(c).includes('WK-deleteditems/childFolders'))).toBe(false)
  })

  it('a failed well-known lookup never breaks the tree', async () => {
    routeTree({ failWellKnown: true })
    const { tree, excludedFolderIds } = await makeApi().fetchFolderTreeWithMeta()
    expect(excludedFolderIds).toEqual([])
    expect(tree.map((n) => n.path)).toEqual(['收件匣', '案件', '刪除的郵件', '寄件備份'])
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

describe('listMessagesModifiedSince (harvester, item 6)', () => {
  it('one mailbox-wide query: filter + orderby on LastModifiedDateTime, ParentFolderId selected, paged', async () => {
    fetchSpy
      .mockResolvedValueOnce(
        json({
          value: [{ Id: 'm1', ParentFolderId: 'f1' }],
          '@odata.nextLink': `${BASE}/me/messages?$skip=1`,
        }),
      )
      .mockResolvedValueOnce(json({ value: [{ Id: 'm2', ParentFolderId: 'f2' }] }))
    const r = await makeApi().listMessagesModifiedSince('2026-11-01T00:00:00Z', { pageSize: 1 })
    expect(r.complete).toBe(true)
    expect(r.messages.map((m) => m.Id)).toEqual(['m1', 'm2'])
    const first = decodeURIComponent(urlOf(fetchSpy.mock.calls[0]!))
    expect(first).toContain('/me/messages?')
    expect(first).toContain('$filter=LastModifiedDateTime gt 2026-11-01T00:00:00Z')
    expect(first).toContain('$orderby=LastModifiedDateTime desc')
    expect(first).toContain('ParentFolderId')
  })

  it('reports complete=false when the page cap stops the walk', async () => {
    fetchSpy.mockImplementation(async () =>
      json({ value: [{ Id: `m${Math.random()}` }], '@odata.nextLink': `${BASE}/me/messages?$skip=${Math.random()}` }),
    )
    const r = await makeApi().listMessagesModifiedSince('2026-11-01T00:00:00Z', { maxPages: 3 })
    expect(r.complete).toBe(false)
    expect(r.pages).toBe(3)
    expect(fetchSpy).toHaveBeenCalledTimes(3)
  })
})
