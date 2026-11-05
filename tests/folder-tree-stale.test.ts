// getOrFetchFolderTree stale-cache visibility (2026-11-05 audit, item 16).
//
// A failing refresh used to fall back to the cached tree with only a
// console.warn — silently, forever. It still serves the cache, but now logs
// to the central error log ONCE per failure streak and records the streak
// (session key 'folderTreeStale') so getStatus / the UI can surface it; a
// successful refresh clears it. The fresh tree also persists the excluded
// well-known folder ids next to the cache.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { getErrorLog } from '@/shared/error-log'
import { getFolderCache, setFolderCache } from '@/shared/storage'
import type { MailFolderNode } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const OLD_TREE: MailFolderNode[] = [
  { id: 'FOLDER-OLD-00000000000000', displayName: '舊', path: '舊', children: [] },
]

let sw: SwHarness
let fetchSpy: ReturnType<typeof vi.fn>

beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  sw.tabsQuery.mockResolvedValue([
    { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
  ])
  sw.tabsSendMessage.mockResolvedValue({ ok: true, secret: 'tok', expiresOn: Date.now() / 1000 + 3600 })
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
})
afterEach(() => vi.restoreAllMocks())

const stale = async () =>
  ((await chrome.storage.session.get('folderTreeStale')) as Record<string, unknown>).folderTreeStale

describe('folder tree refresh failure streak', () => {
  it('serves the cache, logs ONCE per streak, records + clears staleness', async () => {
    await setFolderCache({ updatedAt: '2026-11-01T00:00:00.000Z', tree: OLD_TREE })
    const wellKnown = (url: string) => {
      const wk = url.match(/\/me\/MailFolders\/(deleteditems|junkemail|drafts|outbox|sentitems)\?/)
      return wk ? new Response(JSON.stringify({ Id: `WK-${wk[1]}` }), { status: 200 }) : null
    }
    // The folder listing fails with a definitive 400 (no retry backoff).
    fetchSpy.mockImplementation(async (input: unknown) => wellKnown(String(input)) ?? new Response('nope', { status: 400 }))

    const r1 = await sw.send({ type: 'getFolderTree', forceFresh: true })
    const r2 = await sw.send({ type: 'getFolderTree', forceFresh: true })
    expect((r1.data as { tree: MailFolderNode[] }).tree).toEqual(OLD_TREE)
    expect((r2.data as { tree: MailFolderNode[] }).tree).toEqual(OLD_TREE)
    const logged = (await getErrorLog()).filter((e) => e.source === 'outlook:folder-tree')
    expect(logged).toHaveLength(1)
    expect(await stale()).toMatchObject({ cacheUpdatedAt: '2026-11-01T00:00:00.000Z' })

    // Recovery: a working refresh clears the streak and persists excluded ids.
    fetchSpy.mockImplementation(async (input: unknown) => {
      const url = String(input)
      const wk = wellKnown(url)
      if (wk) return wk
      if (url.includes('/me/MailFolders?$top=100')) {
        return new Response(
          JSON.stringify({
            value: [
              { Id: 'FOLDER-NEW-00000000000000', DisplayName: '新', ChildFolderCount: 0 },
              { Id: 'WK-deleteditems', DisplayName: '刪除的郵件', ChildFolderCount: 3 },
            ],
          }),
          { status: 200 },
        )
      }
      throw new Error(`unexpected ${url}`)
    })
    const r3 = await sw.send({ type: 'getFolderTree', forceFresh: true })
    expect((r3.data as { tree: MailFolderNode[] }).tree.map((n) => n.path)).toEqual(['新'])
    expect(await stale()).toBeUndefined()
    expect((await getFolderCache())?.excludedFolderIds).toContain('WK-deleteditems')
  })
})
