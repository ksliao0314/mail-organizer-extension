// Folder tree / trash rules, round 3 (2026-11-06 r3), through the REAL
// service worker:
//   1 — every well-known lookup failing (e.g. a failed token refresh) built a
//       tree with nothing recognised — trash subtree and Drafts listed, Junk /
//       Sent untagged — and the same preflight handed it to the AI.
//   2 — a pre-v1.0.8 cache (no well-known info at all) wasn't legacy: served
//       for up to a day, trash subtree and Junk / Sent offered to the AI.
//   6 — an invalidation made while a refetch ran was overwritten by it.
//   8 — the AI's top-level "new" 「刪除的郵件」 looked valid in the plan.
//   9 — a rule routing to 刪除的郵件 showed as a move the popup marked
//       「執行時會失敗」 (the tree it has no longer lists the trash).

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { newRule } from '@/shared/rules'
import { getFolderCache, invalidateFolderCache, setFolderCache, setRules } from '@/shared/storage'
import type { FolderCache, MailFolderNode, PlanItem } from '@/shared/types'
import { answerEvents, sseResponse, sseText } from './helpers/claude-sse'
import {
  folderTree,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  seedClassifyEnv,
  waitFor,
  type FetchRouter,
} from './helpers/classify-env'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const TRASH = 'AAMk-WELLKNOWN-DELETEDITEMS-0000000000000'
const JUNK = 'AAMk-WELLKNOWN-JUNKEMAIL-00000000000000000'
const DRAFTS = 'AAMk-WELLKNOWN-DRAFTS-0000000000000000000'
const ROLES: Record<string, { Id: string; DisplayName: string }> = {
  deleteditems: { Id: TRASH, DisplayName: '刪除的郵件' },
  drafts: { Id: DRAFTS, DisplayName: '草稿' },
  outbox: { Id: 'AAMk-WELLKNOWN-OUTBOX-000000000000000000', DisplayName: '寄件匣' },
  junkemail: { Id: JUNK, DisplayName: '垃圾郵件' },
  sentitems: { Id: 'AAMk-WELLKNOWN-SENTITEMS-0000000000000000', DisplayName: '寄件備份' },
}
const WELL_KNOWN: FolderCache['wellKnown'] = {
  complete: true,
  roots: Object.entries(ROLES).map(([role, f]) => ({
    id: f.Id,
    role: role as 'deleteditems',
    displayName: f.DisplayName,
  })),
}

let sw: SwHarness
let router: FetchRouter
let urls: string[]
let wellKnownStatus: number | null
let onListing: () => Promise<void>

beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  // No Outlook tab unless a test opens one (the harness mocks persist).
  sw.tabsQuery.mockResolvedValue([])
  sw.tabsSendMessage.mockResolvedValue(undefined)
  await seedClassifyEnv()
  await setFolderCache({ updatedAt: new Date().toISOString(), tree: folderTree(), wellKnown: WELL_KNOWN })
  router = installFetchRouter()
  urls = []
  wellKnownStatus = null
  onListing = async () => {}
  const inner = globalThis.fetch as unknown as (url: string, init?: RequestInit) => Promise<Response>
  const json = (v: unknown, status = 200) => new Response(JSON.stringify(v), { status })
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string, init?: RequestInit) => {
      urls.push(url)
      const wk = url.match(/\/me\/MailFolders\/(deleteditems|drafts|outbox|junkemail|sentitems)\?/)
      if (wk) return wellKnownStatus ? new Response('nope', { status: wellKnownStatus }) : json(ROLES[wk[1]!])
      if (url.includes('/me/MailFolders?$top=100')) {
        await onListing()
        return json({
          value: [
            { Id: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE0', DisplayName: '內部資料', ChildFolderCount: 0 },
            { Id: JUNK, DisplayName: '垃圾郵件', ChildFolderCount: 0 },
            { Id: TRASH, DisplayName: '刪除的郵件', ChildFolderCount: 1 },
            { Id: DRAFTS, DisplayName: '草稿', ChildFolderCount: 0 },
          ],
        })
      }
      return inner(url, init)
    }),
  )
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

const listingRequested = () => urls.some((u) => u.includes('/me/MailFolders?$top=100'))
const paths = (tree: MailFolderNode[]) => tree.flatMap((n) => [n.path, ...n.children.map((c) => c.path)])

describe('folder tree fetch (r3 items 1, 2, 6)', () => {
  const staleCache = (): FolderCache => ({
    updatedAt: new Date(Date.now() - 2 * 86_400_000).toISOString(), // stale → refetch
    tree: folderTree(),
    wellKnown: WELL_KNOWN,
  })

  it('1 — every well-known lookup failing serves the cached tree: nothing unrecognised is built', async () => {
    const old = staleCache()
    await setFolderCache(old)
    // The token refresh works, but every lookup keeps failing — while the
    // listing itself would succeed (and list the trash, untagged Junk…).
    sw.tabsQuery.mockResolvedValue([
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
    ])
    sw.tabsSendMessage.mockResolvedValue({ ok: true, secret: 'tok', expiresOn: Date.now() / 1000 + 3600 })
    wellKnownStatus = 401
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    expect(paths((r.data as { tree: MailFolderNode[] }).tree)).toEqual(paths(folderTree()))
    expect(listingRequested()).toBe(false)
    expect((await getFolderCache())?.updatedAt).toBe(old.updatedAt)
  })

  it('1 — …the same with no token at all (a failed token refresh)', async () => {
    const old = staleCache()
    await setFolderCache(old)
    wellKnownStatus = 401 // → the forced refresh finds no Outlook tab
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    expect(listingRequested()).toBe(false)
    expect((await getFolderCache())?.updatedAt).toBe(old.updatedAt)
  })

  it('2 — a pre-v1.0.8 cache (no well-known info at all) is refetched on first use', async () => {
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: [
        ...folderTree(),
        { id: TRASH, displayName: '刪除的郵件', path: '刪除的郵件', children: [] },
        { id: JUNK, displayName: '垃圾郵件', path: '垃圾郵件', children: [] },
      ],
    })
    const r = await sw.send({ type: 'getFolderTree' })
    expect(listingRequested()).toBe(true)
    const tree = (r.data as { tree: MailFolderNode[] }).tree
    expect(paths(tree)).not.toContain('刪除的郵件')
    expect(tree.find((n) => n.id === JUNK)?.wellKnown).toBe('junkemail')
    expect((await getFolderCache())?.wellKnown?.complete).toBe(true)
  })

  it('6 — an invalidation made while the refetch ran is kept (the tree may predate it)', async () => {
    await setFolderCache({ updatedAt: new Date(Date.now() - 2 * 86_400_000).toISOString(), tree: folderTree(), wellKnown: WELL_KNOWN })
    onListing = () => invalidateFolderCache() // execute finds a trashed folder mid-fetch
    await sw.send({ type: 'getFolderTree' })
    expect(listingRequested()).toBe(true)
    expect((await getFolderCache())?.invalidatedAt).toBeTruthy()
  })
})

describe('the plan (r3 items 8, 9)', () => {
  const answer = (actions: unknown[]) => () => Promise.resolve(sseResponse(sseText(answerEvents(JSON.stringify({ actions }))), { chunkBytes: 64 }))
  type Full = { done: boolean; aiPlan: PlanItem[] }
  const full = async () => ((await sw.send({ type: 'getAiClassifyProgress', view: 'full' })).data ?? null) as Full | null

  it('8 — the AI’s top-level new folder named like the trash is unresolved, not a doomed create', async () => {
    router.inboxPages = [{ messages: [inboxMail(1)] }]
    router.claude = answer([
      { emailIndex: 0, action: 'new_folder', suggestedFolderName: '刪除的郵件', confidence: 0.9, reason: 'x' },
    ])
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    await markBodiesFetched()
    await sw.send({ type: 'classifyAi', runId: (pf.data as { runId: string }).runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    expect((await full())!.aiPlan[0]).toMatchObject({ action: 'skip', source: 'unresolved', aiOriginalAction: 'new_folder' })
  })

  it('9 — a user rule on Deleted Items itself is a DELETE row, rule kept (preflight)', async () => {
    await setRules([
      {
        ...newRule({ type: 'sender', signal: 'spam@shop.example', targetFolderId: TRASH, targetFolderPath: '刪除的郵件', confidence: 0.75, source: 'user_manual' }),
        id: 'R-trash',
      },
    ])
    router.inboxPages = [{ messages: [inboxMail(1, { From: { EmailAddress: { Address: 'spam@shop.example', Name: 'S' } } })] }]
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    const [row] = (pf.data as { rulePlan: PlanItem[] }).rulePlan
    expect(row).toMatchObject({ action: 'delete', source: 'rule', ruleId: 'R-trash', originalRuleId: 'R-trash' })
    expect(row!.matchedRule?.id).toBe('R-trash')
    expect(row!.targetFolderId).toBeUndefined()
    expect(row!.targetFolderPath).toBeUndefined()
  })

  it('9 — …and a body re-match hit on it too', async () => {
    await setRules([
      {
        ...newRule({ type: 'case_code', signal: '25A0067A', targetFolderId: TRASH, targetFolderPath: '刪除的郵件', confidence: 0.9, source: 'user_manual' }),
        id: 'R-trash-case',
      },
    ])
    router.inboxPages = [{ messages: [inboxMail(1)] }] // the code is only in the body
    const inner = globalThis.fetch as unknown as (url: string, init?: RequestInit) => Promise<Response>
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init?: RequestInit) =>
        url.includes('$select=UniqueBody')
          ? new Response(JSON.stringify({ UniqueBody: { ContentType: 'text', Content: '本案代號 25A0067A,請查照。' } }), { status: 200 })
          : inner(url, init),
      ),
    )
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    expect((pf.data as { rulePlan: PlanItem[] }).rulePlan).toEqual([])
    await sw.send({ type: 'classifyAi', runId: (pf.data as { runId: string }).runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    expect((await full())!.aiPlan[0]).toMatchObject({ action: 'delete', source: 'rule', ruleId: 'R-trash-case' })
    expect(router.claudeBodies).toHaveLength(0) // routed by the rule, never sent to Claude
  })
})
