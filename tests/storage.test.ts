import { describe, expect, it, vi, type Mock } from 'vitest'
import { clearErrorLog, getErrorLog, logError } from '@/shared/error-log'
import {
  addFolderToCache,
  addToSkipHistory,
  getFolderActivity,
  getFolderCache,
  getSkipHistory,
  getStorageTopKeys,
  getStorageUsage,
  getSubjectMemory,
  mergeFolderActivityScan,
  recordFolderActivityFromBatch,
  recordSubjectFilings,
  setFolderCache,
} from '@/shared/storage'

const baseTree = () => [
  {
    id: 'top1',
    displayName: '內部資料',
    path: '內部資料',
    children: [
      { id: 'top1a', displayName: '工時審閱', path: '內部資料/工時審閱', children: [] },
    ],
  },
  {
    id: 'top2',
    displayName: '案件',
    path: '案件',
    children: [],
  },
]

describe('addFolderToCache', () => {
  it('returns false when no cache exists', async () => {
    const ok = await addFolderToCache(
      { Id: 'new1', DisplayName: 'AA' },
      undefined,
    )
    expect(ok).toBe(false)
  })

  it('splices a top-level folder into the cache', async () => {
    await setFolderCache({ updatedAt: '2026-01-01T00:00:00Z', tree: baseTree() })
    const ok = await addFolderToCache(
      { Id: 'new1', DisplayName: 'NewTop', ParentFolderId: undefined },
      undefined,
    )
    expect(ok).toBe(true)
    const cache = await getFolderCache()
    const top = cache!.tree.find((n) => n.id === 'new1')
    expect(top).toBeDefined()
    expect(top!.path).toBe('NewTop')
  })

  it('splices a child folder under an existing parent', async () => {
    await setFolderCache({ updatedAt: '2026-01-01T00:00:00Z', tree: baseTree() })
    const ok = await addFolderToCache(
      { Id: 'newchild', DisplayName: 'AA', ParentFolderId: 'top2' },
      'top2',
    )
    expect(ok).toBe(true)
    const cache = await getFolderCache()
    const parent = cache!.tree.find((n) => n.id === 'top2')
    const child = parent!.children.find((c) => c.id === 'newchild')
    expect(child).toBeDefined()
    expect(child!.path).toBe('案件/AA')
  })

  it('returns false when parent ID not found in cached tree', async () => {
    await setFolderCache({ updatedAt: '2026-01-01T00:00:00Z', tree: baseTree() })
    const ok = await addFolderToCache(
      { Id: 'x', DisplayName: 'X' },
      'nonexistent',
    )
    expect(ok).toBe(false)
  })

  it('does not duplicate when folder already in target list', async () => {
    await setFolderCache({ updatedAt: '2026-01-01T00:00:00Z', tree: baseTree() })
    await addFolderToCache({ Id: 'dup', DisplayName: 'D' }, undefined)
    await addFolderToCache({ Id: 'dup', DisplayName: 'D' }, undefined)
    const cache = await getFolderCache()
    expect(cache!.tree.filter((n) => n.id === 'dup')).toHaveLength(1)
  })
})

// ---- skipHistory cap + TTL (added 2026-05-22) ------------------------------
//
// Pre-2026-05-22 the skip history was append-only, grew forever. Now caps at
// 5000 entries OR 60-day age, whichever hits first. These tests use the
// real `addToSkipHistory` against a mocked storage to verify the prune
// runs on every write.
describe('skipHistory cap + TTL', () => {
  it('drops entries older than 60 days on next write', async () => {
    // Seed storage with an "ancient" entry by faking Date.now during write.
    const ancient = Date.now() - 70 * 24 * 60 * 60 * 1000 // 70 days ago
    vi.spyOn(Date, 'now').mockReturnValueOnce(ancient)
    await addToSkipHistory(['old-1'])
    // New write — uses real Date.now, triggers prune.
    vi.restoreAllMocks()
    await addToSkipHistory(['fresh-1'])
    const hist = await getSkipHistory()
    expect('old-1' in hist).toBe(false)
    expect('fresh-1' in hist).toBe(true)
  })

  it('caps total entries at 5000, dropping oldest first', async () => {
    // addToSkipHistory captures Date.now() once per call — to get different
    // timestamps per id we have to call it multiple times. Batch writes
    // of size 1000 keeps the test fast while still differentiating "early"
    // from "late" batches in the prune ordering.
    let t = Date.now() - 60_000 // start a minute ago
    const spy = vi.spyOn(Date, 'now').mockImplementation(() => t)
    for (let batch = 0; batch < 6; batch++) {
      t += 1000 // advance 1 sec per batch
      const ids = Array.from({ length: 1000 }, (_, i) => `e-${batch}-${i}`)
      await addToSkipHistory(ids)
    }
    spy.mockRestore()
    const hist = await getSkipHistory()
    expect(Object.keys(hist).length).toBeLessThanOrEqual(5000)
    // Batch 0 ids are oldest → should be dropped. Batch 5 ids are
    // newest → should survive.
    expect('e-0-0' in hist).toBe(false)
    expect('e-5-999' in hist).toBe(true)
  })

  it('preserves entries within TTL even at high count', async () => {
    // 100 entries, all within TTL → all preserved.
    const ids = Array.from({ length: 100 }, (_, i) => `recent-${i}`)
    await addToSkipHistory(ids)
    const hist = await getSkipHistory()
    expect(Object.keys(hist).length).toBe(100)
  })
})

// ---- subject memory conflictCount decay (added 2026-05-22) -----------------
//
// One accidental cross-folder filing used to poison a subject forever.
// New behavior: streak counter on entry; after DECAY_AFTER_STABLE (5)
// consecutive same-folder filings, each subsequent same-folder filing
// drops conflictCount by 1.
describe('subjectMemory conflictCount decay', () => {
  it('increments conflictCount when filed to different folder', async () => {
    await recordSubjectFilings([
      { normalizedSubject: 'foo', folderId: 'f1', folderPath: 'A' },
    ])
    await recordSubjectFilings([
      { normalizedSubject: 'foo', folderId: 'f2', folderPath: 'B' },
    ])
    const mem = await getSubjectMemory()
    expect(mem['foo']?.conflictCount).toBe(1)
    expect(mem['foo']?.folderId).toBe('f2')
  })

  it('decays conflictCount after 5+ consecutive same-folder filings', async () => {
    // Seed conflict: file to A then B (conflict=1), then 6× same-folder C.
    // Wait — once we file to C after B, that's a new conflict not a decay.
    // The decay scenario is: hit one conflict to B, then 6+ subsequent
    // filings to that SAME B build up the streak past threshold.
    await recordSubjectFilings([
      { normalizedSubject: 'bar', folderId: 'f1', folderPath: 'A' },
    ])
    await recordSubjectFilings([
      { normalizedSubject: 'bar', folderId: 'f2', folderPath: 'B' },
    ])
    // Now 6 consecutive to f2 should push streak past 5 and start decaying.
    for (let i = 0; i < 6; i++) {
      await recordSubjectFilings([
        { normalizedSubject: 'bar', folderId: 'f2', folderPath: 'B' },
      ])
    }
    const mem = await getSubjectMemory()
    // After streak crosses threshold (5), each extra same-folder filing
    // decrements conflictCount. With 6 same-folder filings post-conflict,
    // streak goes 1,2,3,4,5,6 — only the 6th triggers decay (streak > 5).
    expect(mem['bar']?.conflictCount).toBeLessThan(1)
  })

  it('resets streak when next filing conflicts', async () => {
    await recordSubjectFilings([
      { normalizedSubject: 'baz', folderId: 'f1', folderPath: 'A' },
    ])
    // Build up a streak on f1
    for (let i = 0; i < 6; i++) {
      await recordSubjectFilings([
        { normalizedSubject: 'baz', folderId: 'f1', folderPath: 'A' },
      ])
    }
    // Now conflict — streak resets, conflictCount increments
    await recordSubjectFilings([
      { normalizedSubject: 'baz', folderId: 'f2', folderPath: 'B' },
    ])
    const mem = await getSubjectMemory()
    expect(mem['baz']?.stableStreak).toBe(0)
    expect(mem['baz']?.conflictCount).toBeGreaterThanOrEqual(1)
  })

  // 優化 2026-07: return-to-nest must not double-penalize.
  it('returning to the pre-conflict home does NOT count as a second conflict', async () => {
    // A → B (conflict, cc=1, home remembered = A) → A (回巢, NOT a new conflict)
    await recordSubjectFilings([{ normalizedSubject: 'k', folderId: 'A', folderPath: 'PA' }])
    await recordSubjectFilings([{ normalizedSubject: 'k', folderId: 'B', folderPath: 'PB' }])
    await recordSubjectFilings([{ normalizedSubject: 'k', folderId: 'A', folderPath: 'PA' }])
    const mem = await getSubjectMemory()
    expect(mem['k']?.conflictCount).toBe(1) // NOT 2
    expect(mem['k']?.folderId).toBe('A') // back home
    expect(mem['k']?.stableStreak).toBe(1) // streak resumed, not reset
  })

  it('recovers to conflictCount 0 after ~5 filings once returned home (was ~8)', async () => {
    await recordSubjectFilings([{ normalizedSubject: 'm', folderId: 'A', folderPath: 'PA' }])
    await recordSubjectFilings([{ normalizedSubject: 'm', folderId: 'B', folderPath: 'PB' }]) // cc=1
    await recordSubjectFilings([{ normalizedSubject: 'm', folderId: 'A', folderPath: 'PA' }]) // 回巢, streak=1
    // 5 more home filings: streak 2,3,4,5,6 — the 6th crosses threshold → cc 1→0.
    for (let i = 0; i < 5; i++) {
      await recordSubjectFilings([{ normalizedSubject: 'm', folderId: 'A', folderPath: 'PA' }])
    }
    const mem = await getSubjectMemory()
    expect(mem['m']?.conflictCount).toBe(0)
    expect(mem['m']?.previousFolderId).toBeUndefined()
  })

  it('oscillating A/B/A/B keeps accumulating conflicts (stays disqualified)', async () => {
    await recordSubjectFilings([{ normalizedSubject: 'o', folderId: 'A', folderPath: 'PA' }])
    await recordSubjectFilings([{ normalizedSubject: 'o', folderId: 'B', folderPath: 'PB' }]) // cc=1, home=A
    await recordSubjectFilings([{ normalizedSubject: 'o', folderId: 'A', folderPath: 'PA' }]) // 回巢, cc=1
    await recordSubjectFilings([{ normalizedSubject: 'o', folderId: 'B', folderPath: 'PB' }]) // new conflict, cc=2
    const mem = await getSubjectMemory()
    expect(mem['o']?.conflictCount).toBe(2)
  })
})

// ---- folderActivity.latestMessage preservation (added 2026-05-22) ---------
//
// The bug: pre-2026-05-22, `recordFolderActivityFromBatch` rebuilt each
// entry from scratch on every batch — wiping `latestMessage` that
// `mergeFolderActivityScan` (the 重新整理 button) had populated. After
// the user pressed 重新整理 once, the very next classify batch erased
// the subject preview.
describe('folderActivity.latestMessage preservation', () => {
  // Timestamps MUST be relative to now (time-bomb regression, 2026-06):
  // writeFolderActivity prunes entries older than FOLDER_ACTIVITY_MAX_AGE_MS
  // (30 days) against the REAL clock, so hard-coded absolute dates made
  // these tests start failing the day the fixtures aged past the window —
  // green on 2026-06-18, red on 2026-06-23, no code change in between.
  const HOUR = 60 * 60 * 1000
  const t0 = new Date(Date.now() - 12 * HOUR).toISOString() // scan time
  const t1 = new Date(Date.now() - 1 * HOUR).toISOString() // batch time (newer)

  it('preserves prev.latestMessage when batch supplies none', async () => {
    // Step 1: scan populates latestMessage
    await mergeFolderActivityScan([
      {
        folderId: 'f1',
        folderPath: 'P/A',
        latestMessageAt: t0,
        latestMessage: {
          subject: 'foo',
          from: 'a@b.com',
          receivedAt: t0,
        },
      },
    ])
    // Step 2: batch without latestMessage (simulates pre-fix call site).
    await recordFolderActivityFromBatch(
      [{ folderId: 'f1', folderPath: 'P/A', count: 3 }],
      t1,
    )
    const activity = await getFolderActivity()
    const row = activity.find((r) => r.folderId === 'f1')
    expect(row?.latestMessage?.subject).toBe('foo')
    expect(row?.recentCount).toBe(3)
  })

  it('updates latestMessage when batch supplies a fresher one', async () => {
    await mergeFolderActivityScan([
      {
        folderId: 'f2',
        folderPath: 'P/B',
        latestMessageAt: t0,
        latestMessage: {
          subject: 'old subject',
          from: 'old@example.com',
          receivedAt: t0,
        },
      },
    ])
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'f2',
          folderPath: 'P/B',
          count: 1,
          latestMessage: {
            subject: 'new subject',
            from: 'new@example.com',
            receivedAt: t1,
          },
        },
      ],
      t1,
    )
    const activity = await getFolderActivity()
    const row = activity.find((r) => r.folderId === 'f2')
    expect(row?.latestMessage?.subject).toBe('new subject')
  })
})

// ---- error log: clear vs in-flight append (2026-11-05) ---------------------
//
// clearErrorLog did a bare remove() outside the append lock. A clear landing
// between an in-flight logError's read and its set() was undone: the append
// re-wrote the whole pre-clear list plus its new entry.
describe('clearErrorLog takes the append lock', () => {
  it('a clear issued while an append is mid read-modify-write is not undone', async () => {
    await chrome.storage.local.set({
      errorLog: [
        { at: '2026-11-01T00:00:00.000Z', source: 'old:a', message: 'old A' },
        { at: '2026-11-01T00:00:01.000Z', source: 'old:b', message: 'old B' },
      ],
    })
    // Hold logError's read open so the clear lands between its read and set.
    const get = chrome.storage.local.get as unknown as Mock<
      (keys?: unknown) => Promise<Record<string, unknown>>
    >
    const realGet = get.getMockImplementation()!
    let releaseRead!: () => void
    const readGate = new Promise<void>((r) => {
      releaseRead = r
    })
    let readStarted = false
    get.mockImplementationOnce(async (keys) => {
      readStarted = true
      const v = await realGet(keys)
      await readGate
      return v
    })
    try {
      const appending = logError('new:c', 'new C')
      while (!readStarted) await new Promise((r) => setTimeout(r, 0))
      const clearing = clearErrorLog()
      releaseRead()
      await Promise.all([appending, clearing])
    } finally {
      get.mockImplementation(realGet)
    }
    // The user's clear came after the append began → nothing survives, in
    // particular not the pre-clear entries the append had already read.
    expect(await getErrorLog()).toEqual([])
  })
})

// ---- storage quota: read the real one (2026-11-05) -------------------------
describe('getStorageUsage / getStorageTopKeys', () => {
  type RuntimeStub = { getManifest?: () => { permissions?: string[] } }
  const g = globalThis as unknown as { chrome: { runtime?: RuntimeStub } }

  it('measures against chrome.storage.local.QUOTA_BYTES, not a hard-coded 5 MB', async () => {
    const area = chrome.storage.local as unknown as { QUOTA_BYTES: number }
    const saved = area.QUOTA_BYTES
    area.QUOTA_BYTES = 10 * 1024 * 1024
    try {
      await chrome.storage.local.set({ blob: 'x'.repeat(5 * 1024 * 1024) })
      const u = await getStorageUsage()
      expect(u.quotaBytes).toBe(10 * 1024 * 1024)
      // 5 MB of 10 MB: ~50%, NOT "100% / approaching" as the 5 MB constant said.
      expect(u.percentUsed).toBeGreaterThan(0.49)
      expect(u.percentUsed).toBeLessThan(0.51)
      expect(u.approaching).toBe(false)

      await chrome.storage.local.set({ blob2: 'y'.repeat(3.5 * 1024 * 1024) })
      expect((await getStorageUsage()).approaching).toBe(true)
    } finally {
      area.QUOTA_BYTES = saved
    }
  })

  it('reports no quota (and never warns) under unlimitedStorage', async () => {
    const saved = g.chrome.runtime
    g.chrome.runtime = { getManifest: () => ({ permissions: ['storage', 'unlimitedStorage'] }) }
    try {
      await chrome.storage.local.set({ blob: 'x'.repeat(200_000) })
      const u = await getStorageUsage()
      expect(u.quotaBytes).toBeNull()
      expect(u.percentUsed).toBeNull()
      expect(u.approaching).toBe(false)
      expect(u.bytesInUse).toBeGreaterThan(200_000)
    } finally {
      g.chrome.runtime = saved
    }
  })

  it('lists the biggest keys by name only, biggest first, capped', async () => {
    await chrome.storage.local.set({
      settings: { claudeApiKey: 'sk-ant-api03-SECRET' },
      rules: Array.from({ length: 50 }, (_, i) => ({ id: `r${i}`, signal: 'x'.repeat(40) })),
      metrics: { moved: 1 },
    })
    const top = await getStorageTopKeys(2)
    expect(top.map((t) => t.key)).toEqual(['rules', 'settings'])
    expect(top[0]!.bytes).toBeGreaterThan(top[1]!.bytes)
    expect(JSON.stringify(await getStorageTopKeys())).not.toContain('sk-ant')
  })
})
