// Non-filing folders, end to end through the real service worker (C2,
// 2026-11-06; redesigned O1, 2026-11-06 r2).
//
// v1.0.8 left Deleted Items, Junk Email, Drafts, Outbox and Sent Items OUT of
// the folder tree with their whole subtrees: reconcile orphaned every rule
// filing into Junk / Sent Items, user_manual included. Round 1 put Junk / Sent
// (and the trash root) back, but stopped orphaning anything "excluded but
// existing" — so a case folder deleted in OWA (moved into the trash with the
// same id) kept routing mail into an execute refusal every batch — and listed
// the whole trash subtree on every fetch. Now: Deleted Items (root + subtree),
// Drafts and Outbox are left out unlisted; only Deleted Items ITSELF stays a
// valid rule target; everything else missing from a complete tree is orphaned.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { listRules, newRule } from '@/shared/rules'
import { getFolderCache, setFolderCache, setRules } from '@/shared/storage'
import type { FolderCache, MailFolderNode, Rule } from '@/shared/types'
import { seedClassifyEnv } from './helpers/classify-env'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const JUNK = 'WK-junkemail-000000000000000'
const SENT = 'WK-sentitems-000000000000000'
const SENT_A = 'FOLDER-SENT-CLIENT-A-0000000'
const TRASH = 'WK-deleteditems-0000000000000'
const DRAFTS = 'WK-drafts-0000000000000000000'
const OUTBOX = 'WK-outbox-0000000000000000000'
const TRASHED_CASE = 'FOLDER-TRASHED-CASE-000000000'
const CASES = 'FOLDER-CASES-000000000000000'
const PURGED = 'FOLDER-PURGED-0000000000000'

const WELL_KNOWN: Record<string, { Id: string; DisplayName: string }> = {
  deleteditems: { Id: TRASH, DisplayName: '刪除的郵件' },
  drafts: { Id: DRAFTS, DisplayName: '草稿' },
  outbox: { Id: OUTBOX, DisplayName: '寄件匣' },
  junkemail: { Id: JUNK, DisplayName: '垃圾郵件' },
  sentitems: { Id: SENT, DisplayName: '寄件備份' },
}

function folder(id: string, name: string, childCount = 0) {
  return { Id: id, DisplayName: name, ChildFolderCount: childCount }
}
const json = (v: unknown, status = 200) =>
  new Response(JSON.stringify(v), { status, headers: { 'content-type': 'application/json' } })

let sw: SwHarness
let urls: string[]

beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await seedClassifyEnv()
  urls = []
  vi.stubGlobal(
    'fetch',
    vi.fn(async (input: string) => {
      const url = String(input)
      urls.push(url)
      const wk = url.match(/\/me\/MailFolders\/(deleteditems|drafts|outbox|junkemail|sentitems)\?/)
      if (wk) return json(WELL_KNOWN[wk[1]!])
      if (url.includes('/me/MailFolders?$top=100')) {
        return json({
          value: [
            folder(CASES, '案件'),
            folder(JUNK, '垃圾郵件'),
            folder(SENT, '寄件備份', 1),
            folder(TRASH, '刪除的郵件', 1),
            folder(DRAFTS, '草稿'),
            folder(OUTBOX, '寄件匣'),
          ],
        })
      }
      if (url.includes(`/MailFolders/${SENT}/childFolders`)) return json({ value: [folder(SENT_A, '客戶A')] })
      if (url.includes(`/MailFolders/${TRASH}/childFolders`)) return json({ value: [folder(TRASHED_CASE, '甲')] })
      if (url.includes('/me/MailFolders/inbox/messages')) return json({ value: [] })
      if (/\/me(\?|$)/.test(url)) return json({ EmailAddress: 'me@firm.example' })
      return json({ value: [] })
    }),
  )
})
afterEach(() => vi.unstubAllGlobals())

/** A fresh-looking v1.0.8 cache: the five well-known subtrees missing, roots "excluded". */
async function seedV108Cache() {
  const tree: MailFolderNode[] = [{ id: CASES, displayName: '案件', path: '案件', children: [] }]
  await setFolderCache({
    updatedAt: new Date().toISOString(),
    tree,
    excludedFolderIds: [TRASH, JUNK, SENT, DRAFTS, OUTBOX],
  })
}

const rule = (id: string, over: Partial<Rule>): Rule => ({
  ...newRule({
    type: 'sender',
    signal: `${id}@x.example`,
    targetFolderId: CASES,
    targetFolderPath: '案件',
    confidence: 0.7,
    source: 'user_manual',
  }),
  id,
  ...over,
})

const treePaths = (tree: MailFolderNode[]) => tree.flatMap((n) => [n.path, ...n.children.map((c) => c.path)])

describe('non-filing folders (O1)', () => {
  it('a v1.0.8 cache is refetched on first use, even while "fresh"; the trash is never listed', async () => {
    await seedV108Cache()
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r.ok).toBe(true)
    expect(treePaths((r.data as { tree: MailFolderNode[] }).tree)).toEqual(['案件', '垃圾郵件', '寄件備份', '寄件備份/客戶A'])
    const cache = await getFolderCache()
    expect(cache?.excludedFolderIds).toBeUndefined()
    expect(cache?.wellKnown?.complete).toBe(true)
    expect(cache?.wellKnown?.roots).toContainEqual({ id: TRASH, role: 'deleteditems', displayName: '刪除的郵件' })
    // (d) The trash's subtree is not even listed.
    expect(urls.some((u) => u.includes(`/MailFolders/${TRASH}/childFolders`))).toBe(false)
    // …and a new-format cache is served from storage again.
    urls = []
    await sw.send({ type: 'getFolderTree' })
    expect(urls).toEqual([])
  })

  it('a round-1 cache (trash in the tree, nonFiling) is refetched once too', async () => {
    const roundOne: FolderCache = {
      updatedAt: new Date().toISOString(),
      tree: [
        { id: CASES, displayName: '案件', path: '案件', children: [] },
        { id: TRASH, displayName: '刪除的郵件', path: '刪除的郵件', children: [] },
      ],
      nonFiling: { trashFolderId: TRASH, byId: { [TRASHED_CASE]: 'deleteditems', [DRAFTS]: 'drafts' } },
    }
    await setFolderCache(roundOne)
    const r = await sw.send({ type: 'getFolderTree' })
    expect(treePaths((r.data as { tree: MailFolderNode[] }).tree)).not.toContain('刪除的郵件')
    const cache = await getFolderCache()
    expect(cache?.nonFiling).toBeUndefined()
    expect(cache?.wellKnown?.complete).toBe(true)
  })

  it('preflight reconcile: Deleted Items itself stays valid, a folder trashed in OWA is orphaned', async () => {
    await seedV108Cache()
    await setRules([
      rule('junk-user', { targetFolderId: JUNK, targetFolderPath: '垃圾郵件', orphaned: true }),
      rule('sent-ai', {
        type: 'domain',
        signal: 'client-a.example',
        targetFolderId: SENT_A,
        targetFolderPath: '寄件備份/客戶A',
        source: 'ai_confirmed',
        orphaned: true,
      }),
      rule('trash-user', { targetFolderId: TRASH, targetFolderPath: '刪除的郵件', orphaned: true }),
      rule('trashed-user', { targetFolderId: TRASHED_CASE, targetFolderPath: '案件/甲' }),
      rule('drafts-ai', { targetFolderId: DRAFTS, targetFolderPath: '草稿', source: 'ai_confirmed' }),
      rule('purged-ai', { targetFolderId: PURGED, targetFolderPath: '案件/乙', source: 'ai_confirmed' }),
    ])
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    const byId = new Map((await listRules()).map((x) => [x.id, x]))
    expect(byId.get('junk-user')!.orphaned).toBe(false)
    expect(byId.get('sent-ai')!.orphaned).toBe(false)
    // A deliberate auto-trash rule: valid (v1.0.8's orphan flag cleared).
    expect(byId.get('trash-user')!.orphaned).toBe(false)
    // Round 1 kept this one live: every batch then routed mail into an
    // execute refusal (red row, failing retry, the AI never saw the mail).
    expect(byId.get('trashed-user')!.orphaned).toBe(true)
    expect(byId.get('trashed-user')!.targetFolderPath).toBe('案件/甲') // never re-pathed into the trash
    expect(byId.get('drafts-ai')!.orphaned).toBe(true)
    expect(byId.get('purged-ai')!.orphaned).toBe(true)
  })

  // Gap 1 (2026-11-06 r2): execute found a folder of the cached tree in the
  // trash (deleted in OWA after the fetch) and invalidated the cache.
  it('a cache execute invalidated is refetched by the next preflight, which orphans the rule', async () => {
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: [
        {
          id: CASES,
          displayName: '案件',
          path: '案件',
          children: [{ id: TRASHED_CASE, displayName: '甲', path: '案件/甲', children: [] }],
        },
      ],
      wellKnown: { roots: [{ id: TRASH, role: 'deleteditems' }], complete: true },
      // Invalidated by an earlier batch — before this refetch started.
      invalidatedAt: new Date(Date.now() - 1_000).toISOString(),
    })
    await setRules([rule('trashed-user', { targetFolderId: TRASHED_CASE, targetFolderPath: '案件/甲' })])
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(urls.some((u) => u.includes('/me/MailFolders?$top=100'))).toBe(true)
    expect((await getFolderCache())?.invalidatedAt).toBeUndefined()
    expect((await listRules())[0]!.orphaned).toBe(true)
  })

  it('(e) a tree from a partial lookup is reconciled non-destructively and only briefly fresh', async () => {
    // What a degraded fetch caches: the trash unrecognised → listed with its
    // subtree, where the case folder deleted in OWA now sits.
    const degraded: FolderCache = {
      updatedAt: new Date().toISOString(),
      tree: [
        { id: CASES, displayName: '案件', path: '案件', children: [] },
        {
          id: TRASH,
          displayName: '刪除的郵件',
          path: '刪除的郵件',
          children: [{ id: TRASHED_CASE, displayName: '甲', path: '刪除的郵件/甲', children: [] }],
        },
      ],
      wellKnown: { roots: [{ id: DRAFTS, role: 'drafts' }], complete: false },
    }
    await setFolderCache(degraded)
    await setRules([rule('trashed-user', { targetFolderId: TRASHED_CASE, targetFolderPath: '案件/甲' })])
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    const kept = (await listRules())[0]!
    expect(kept.targetFolderPath).toBe('案件/甲') // round 1 rewrote it to 刪除的郵件/甲
    expect(kept.orphaned ?? false).toBe(false)
    expect(urls.some((u) => u.includes('/me/MailFolders?$top=100'))).toBe(false) // still fresh

    // A minute later it is stale: the next preflight refetches — a complete
    // tree now — and the rule is orphaned.
    await setFolderCache({ ...degraded, updatedAt: new Date(Date.now() - 2 * 60_000).toISOString() })
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(urls.some((u) => u.includes('/me/MailFolders?$top=100'))).toBe(true)
    expect((await getFolderCache())?.wellKnown?.complete).toBe(true)
    expect((await listRules())[0]!.orphaned).toBe(true)
  })
})
