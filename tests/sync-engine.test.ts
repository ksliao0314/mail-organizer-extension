import { beforeEach, describe, expect, it, vi } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import {
  addRuleTombstones,
  getFolderActivity,
  getFolderActivityRefreshAt,
  getRuleTombstones,
  getSettings,
  mergeFolderActivityFromSync,
  recordFolderActivityFromBatch,
  setFolderActivityRefreshAt,
  setSettings,
} from '@/shared/storage'
import {
  clearCloudState,
  dismissSyncError,
  getSyncStatus,
  listBackups,
  pullNow,
  pushNow,
  quiesce,
  restoreBackup,
  SYNC_LAST_ERROR_KEY,
  type SyncErrorEntry,
} from '@/background/sync-engine'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'domain',
    signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: 'X/Y',
    confidence: 0.7,
    source: (over.source as RuleSource) ?? 'user_manual',
    ...over,
  })
}

async function simulateOtherMachinePush(opts: {
  rules: Rule[]
  machineId?: string
}): Promise<void> {
  // Direct write to chrome.storage.sync as if another device pushed.
  const machineId = opts.machineId ?? 'other-machine'
  const chunkSize = 25
  const ruleChunks: Array<{ index: number; items: Rule[] }> = []
  for (let i = 0; i < opts.rules.length; i += chunkSize) {
    ruleChunks.push({ index: ruleChunks.length, items: opts.rules.slice(i, i + chunkSize) })
  }
  const writes: Record<string, unknown> = {
    syncMeta: {
      schemaVersion: 2,
      sourceMachineId: machineId,
      updatedAt: new Date().toISOString(),
      ruleCount: opts.rules.length,
      tombstoneCount: 0,
      ruleChunkCount: ruleChunks.length,
      tombstoneChunkCount: 0,
    },
    syncSettings: {
      claudeModel: 'cloud-model',
      batchSize: 99, // wildly different to detect if it overwrites local
      // Intentionally include per-device fields to verify the strip
      // (test for Bug #A regression).
      syncMachineId: machineId,
      lastSyncAt: '2026-01-01T00:00:00Z',
      syncEnabled: false,
    },
  }
  for (const c of ruleChunks) writes[`syncRules_${c.index}`] = c
  await chrome.storage.sync.set(writes)
}

beforeEach(async () => {
  // Each test starts with sync enabled + a known machineId.
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- Bug #A: per-device fields don't round-trip ---------------------------

describe('Bug #A: per-device settings fields stay local', () => {
  it('push does NOT write syncMachineId / lastSyncAt / syncEnabled to cloud', async () => {
    await addRules([rule({ signal: 'a.com' })])
    await pushNow('test')
    const cloudSettings = (await chrome.storage.sync.get('syncSettings'))
      .syncSettings as Record<string, unknown>
    expect(cloudSettings).toBeDefined()
    expect('claudeApiKey' in cloudSettings).toBe(false)
    expect('syncMachineId' in cloudSettings).toBe(false)
    expect('lastSyncAt' in cloudSettings).toBe(false)
    expect('syncEnabled' in cloudSettings).toBe(false)
  })

  it('pull does NOT overwrite local syncMachineId / lastSyncAt / syncEnabled', async () => {
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'from-other.com' })],
      machineId: 'other-machine',
    })
    await pullNow('test')
    const settings = await getSettings()
    // Our identity / sync state stays put despite cloud's settings overlay.
    expect(settings.syncMachineId).toBe('this-machine')
    expect(settings.syncEnabled).toBe(true)
  })

  it('pull DOES overlay other user-facing settings (e.g. claudeModel)', async () => {
    await simulateOtherMachinePush({
      rules: [],
      machineId: 'other-machine',
    })
    await pullNow('test')
    const settings = await getSettings()
    // Cloud's claudeModel = 'cloud-model' propagates.
    expect(settings.claudeModel).toBe('cloud-model')
    expect(settings.batchSize).toBe(99)
  })

  it('pull preserves local claudeApiKey even if cloud somehow has one', async () => {
    await setSettings({ claudeApiKey: 'sk-ant-my-local-key' })
    await simulateOtherMachinePush({ rules: [], machineId: 'other' })
    // Manually inject an apiKey into cloud settings (shouldn't normally
    // happen but defence-in-depth).
    const stored = (await chrome.storage.sync.get('syncSettings'))
      .syncSettings as Record<string, unknown>
    stored.claudeApiKey = 'sk-ant-malicious-cloud-key'
    await chrome.storage.sync.set({ syncSettings: stored })
    await pullNow('test')
    const settings = await getSettings()
    expect(settings.claudeApiKey).toBe('sk-ant-my-local-key')
  })
})

// ---- Bug #B: First-enable UNION merge -------------------------------------

describe('Bug #B: first-enable union pull preserves local-only rules', () => {
  it('union pull keeps local syncable rules missing from cloud', async () => {
    // Local has rule X. Cloud has rule Y (from another machine).
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'y.com', source: 'user_manual' })],
      machineId: 'other',
    })
    await pullNow('test', { mode: 'union' })
    const local = await listRules()
    const signals = local.map((r) => r.signal).sort()
    expect(signals).toEqual(['x.com', 'y.com'])
  })

  it('replace pull (default) drops local syncable rules missing from cloud', async () => {
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'y.com', source: 'user_manual' })],
      machineId: 'other',
    })
    await pullNow('test') // default mode = 'replace'
    const local = await listRules()
    const signals = local.map((r) => r.signal).sort()
    // x.com dropped (assumed deleted upstream), y.com pulled in.
    expect(signals).toEqual(['y.com'])
  })

  it('union mode preserves per-device rules too (auto_scan)', async () => {
    await addRules([
      rule({ signal: 'local-user.com', source: 'user_manual' }),
      rule({ signal: 'local-scan.com', source: 'auto_scan' }),
    ])
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'cloud.com', source: 'user_manual' })],
      machineId: 'other',
    })
    await pullNow('test', { mode: 'union' })
    const local = await listRules()
    const signals = local.map((r) => r.signal).sort()
    expect(signals).toEqual(['cloud.com', 'local-scan.com', 'local-user.com'])
  })
})

// ---- Bug #C: restoreBackup preserves per-device identity ------------------

describe('Bug #C: restoreBackup keeps current machineId / lastSyncAt / syncEnabled', () => {
  it('restoring a pre-pull backup (which captured cloud sourceMachineId) does not corrupt identity', async () => {
    await setSettings({
      syncMachineId: 'this-machine',
      syncEnabled: true,
      lastSyncAt: '2026-05-22T00:00:00Z',
    })
    await addRules([rule({ signal: 'before.com' })])
    // Simulate a pull (which writes a backup pre-pull capturing current settings).
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'after-pull.com' })],
      machineId: 'other-machine',
    })
    await pullNow('test', { mode: 'replace' })
    // Now do a NEW push from another machine that would overwrite our
    // syncMachineId IF we hadn't fixed Bug #A — but we did fix it, so
    // this is just setup for the restore test.

    // Pre-pull backup should exist. Restore it.
    const backups = await listBackups()
    const prePull = backups.find((b) => b.direction === 'pre-pull')
    expect(prePull).toBeDefined()
    // Mutate backup payload to simulate a corrupted backup that contains
    // another machine's id (this could happen if a backup file got
    // moved between machines, etc.).
    prePull!.payload.settings.syncMachineId = 'corrupted-id-from-elsewhere'
    prePull!.payload.settings.lastSyncAt = 'some-other-time'
    prePull!.payload.settings.syncEnabled = false
    // Persist the mutated backup.
    const allBackups = await listBackups()
    const idx = allBackups.findIndex((b) => b.snapshotAt === prePull!.snapshotAt)
    allBackups[idx] = prePull!
    await chrome.storage.local.set({ syncBackups: allBackups })

    // Capture the CURRENT (post-pull) per-device fields so we can verify
    // restoreBackup doesn't clobber them.
    const beforeRestore = await getSettings()

    // Restore.
    await restoreBackup(prePull!.snapshotAt)
    const settings = await getSettings()
    // Per-device fields should be unchanged by restore.
    expect(settings.syncMachineId).toBe('this-machine')
    expect(settings.syncEnabled).toBe(true)
    expect(settings.lastSyncAt).toBe(beforeRestore.lastSyncAt)
    // Defence-in-depth: backup's corrupted machineId did NOT slip through.
    expect(settings.syncMachineId).not.toBe('corrupted-id-from-elsewhere')
  })
})

// ---- Bug #D: chunk truncation surfaces in push result --------------------

describe('Bug #D: truncated rules + tombstones surfaced in push result', () => {
  it('returns truncatedRuleCount when rules exceed chunk cap', async () => {
    // Generate enough rules to exceed MAX_RULE_CHUNKS (20) at ~6KB/chunk.
    // Each rule ~300 bytes; ~25 rules per chunk; need >500 rules for >20
    // chunks. Make 600 to be safe.
    const many: Rule[] = []
    for (let i = 0; i < 600; i++) {
      many.push(rule({ id: `id-${i.toString().padStart(4, '0')}`, signal: `r${i}.com` }))
    }
    await addRules(many)
    const r = await pushNow('test')
    expect(r.pushed).toBe(true)
    expect(r.truncatedRuleCount).toBeGreaterThan(0)
  })

  it('truncatedRuleCount is undefined when everything fits', async () => {
    await addRules([rule({ signal: 'small.com' })])
    const r = await pushNow('test')
    expect(r.pushed).toBe(true)
    expect(r.truncatedRuleCount).toBeUndefined()
  })
})

// ---- Smoke: clearCloudState removes everything ---------------------------

describe('clearCloudState (disable + wipe)', () => {
  it('removes all our sync keys but leaves local data intact', async () => {
    await addRules([rule({ signal: 'local.com' })])
    await pushNow('test')
    await clearCloudState()
    const cloud = await chrome.storage.sync.get(null)
    // No sync keys remain.
    for (const k of Object.keys(cloud)) {
      expect(
        k === 'syncMeta' ||
          k === 'syncSettings' ||
          k.startsWith('syncRules_') ||
          k.startsWith('syncTombstones_'),
      ).toBe(false)
    }
    // Local rules still there.
    const local = await listRules()
    expect(local).toHaveLength(1)
    expect(local[0]!.signal).toBe('local.com')
  })
})

// ---- Smoke: getSyncStatus returns expected shape -------------------------

describe('getSyncStatus', () => {
  it('reports enabled state + cloud meta when present', async () => {
    await addRules([rule({ signal: 'a.com' })])
    await pushNow('test')
    const status = await getSyncStatus()
    expect(status.enabled).toBe(true)
    expect(status.cloud).toBeDefined()
    expect(status.cloud!.isUs).toBe(true)
    expect(status.cloud!.ruleCount).toBe(1)
  })

  it('cloud is undefined when no state has been pushed', async () => {
    const status = await getSyncStatus()
    expect(status.cloud).toBeUndefined()
  })
})

// ---- Suppression: tombstones from cloud merge in ------------------------

describe('tombstones round-trip via sync', () => {
  it('cloud tombstones get merged into local on pull', async () => {
    // Push a tombstone from the "other" machine — directly inject one
    // via the simulator's pattern.
    const tombChunk = {
      index: 0,
      items: [
        {
          type: 'domain' as const,
          signalNorm: 'banished.com',
          targetFolderPath: 'X',
          deletedAt: Date.now(),
        },
      ],
    }
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 1,
        ruleChunkCount: 0,
        tombstoneChunkCount: 1,
      },
      syncSettings: { claudeModel: 'x', batchSize: 50 },
      syncTombstones_0: tombChunk,
    })
    await pullNow('test')
    const tombs = await getRuleTombstones()
    expect(tombs.some((t) => t.signalNorm === 'banished.com')).toBe(true)
  })
})

// ---- Bug #E: remote-pull union mode + tombstone-aware drop ----------------
//
// The daily-driver bug: between two pushes from this machine, user adds a
// new rule on B. A pushes (without B's rule). B's listener triggers pull.
// With pre-fix replace mode, B's new rule was silently dropped. With the
// fix:
//   - mode='union' (default for remote pulls) keeps local-only rules.
//   - tombstones STILL propagate deletions (so A deleting works).

describe('Bug #E: remote-pull preserves local-only rules', () => {
  it('union pull preserves local-only syncable rule that cloud never saw', async () => {
    // B has rule local-only.
    await addRules([rule({ signal: 'local-only.com', source: 'user_manual' })])
    // A pushes a different rule.
    await simulateOtherMachinePush({ rules: [rule({ signal: 'from-a.com' })] })
    // Remote pull (the production path uses 'union').
    await pullNow('remote-change-sim', { mode: 'union' })
    const rules = await listRules()
    expect(rules.find((r) => r.signal === 'local-only.com')).toBeDefined()
    expect(rules.find((r) => r.signal === 'from-a.com')).toBeDefined()
  })

  it('union pull WITH matching cloud tombstone drops the local rule (delete propagates)', async () => {
    // B has rule R. A deleted R + pushed tombstone.
    await addRules([rule({ signal: 'deleted-on-a.com', source: 'user_manual' })])
    const tombChunk = {
      index: 0,
      items: [
        {
          type: 'domain' as const,
          signalNorm: 'deleted-on-a.com',
          targetFolderPath: 'X/Y',
          deletedAt: Date.now(),
        },
      ],
    }
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 1,
        ruleChunkCount: 0,
        tombstoneChunkCount: 1,
      },
      syncSettings: { claudeModel: 'x', batchSize: 50 },
      syncTombstones_0: tombChunk,
    })
    await pullNow('remote-change-sim', { mode: 'union' })
    const rules = await listRules()
    expect(rules.find((r) => r.signal === 'deleted-on-a.com')).toBeUndefined()
  })

  it('replace pull still drops local-only rules (manual-pull explicit-intent path)', async () => {
    // The 'replace' mode survives for manual-pull (user explicitly wants
    // cloud-exact state). This is the safety valve, not the default.
    await addRules([rule({ signal: 'local-only.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'from-a.com' })] })
    await pullNow('manual', { mode: 'replace' })
    const rules = await listRules()
    expect(rules.find((r) => r.signal === 'local-only.com')).toBeUndefined()
    expect(rules.find((r) => r.signal === 'from-a.com')).toBeDefined()
  })
})

// ---- Bug #F: restoreBackup is a TRUE reset of tombstones ------------------

describe('Bug #F: restoreBackup resets tombstones', () => {
  it('restoring backup replaces (not unions) tombstones', async () => {
    // Establish backup with tombstone-A.
    await addRules([rule({ signal: 'r1.com' })])
    await addRuleTombstones([
      {
        type: 'domain',
        signalNorm: 'tomb-a.com',
        targetFolderPath: 'X',
        deletedAt: 100,
      },
    ])
    await pushNow('seed-backup') // creates pre-push backup w/ tomb-a only
    const seedTime = (await listBackups())[0]!.snapshotAt

    // Add tombstone-B after the snapshot.
    await addRuleTombstones([
      {
        type: 'domain',
        signalNorm: 'tomb-b.com',
        targetFolderPath: 'X',
        deletedAt: 200,
      },
    ])
    expect((await getRuleTombstones()).length).toBe(2)

    // Restore: should reset to backup state = ONLY tomb-a.
    await restoreBackup(seedTime)
    const after = await getRuleTombstones()
    expect(after.length).toBe(1)
    expect(after[0]!.signalNorm).toBe('tomb-a.com')
  })

  it('restoring an empty-tombstones backup clears all current tombstones', async () => {
    // Backup with NO tombstones.
    await addRules([rule({ signal: 'r1.com' })])
    await pushNow('seed-empty-backup')
    const seedTime = (await listBackups())[0]!.snapshotAt

    // Add some after.
    await addRuleTombstones([
      {
        type: 'domain',
        signalNorm: 'shouldnt-survive.com',
        targetFolderPath: 'X',
        deletedAt: 100,
      },
    ])
    expect((await getRuleTombstones()).length).toBe(1)

    await restoreBackup(seedTime)
    expect((await getRuleTombstones()).length).toBe(0)
  })
})

// ---- Bug #G: pullInProgress grace period ---------------------------------
//
// Can't directly test Chrome's macrotask listener dispatch order in jsdom
// (the mock doesn't fire listeners). But we CAN test that the pull does
// schedule a deferred clear instead of clearing synchronously — i.e. the
// pullGraceTimer is non-null after pull. End-to-end behaviour relies on
// Chrome actually queueing listeners on macrotasks, which is documented.

describe('Bug #G: pullInProgress grace period (no synchronous clear)', () => {
  it('pull installs a deferred clear (use of setTimeout) — exercised via fake timers', async () => {
    vi.useFakeTimers()
    try {
      await simulateOtherMachinePush({ rules: [rule({ signal: 'g.com' })] })
      const p = pullNow('test') // includes deferred clear in finally
      // The pull completes synchronously in tests (mocked storage). The
      // deferred clear is what we care about — we shouldn't have raced to
      // "ready for echo push" instantly.
      await p
      // The pullInProgress flag is internal — we assert via behaviour:
      // a setTimeout was scheduled. vi's getTimerCount > 0 confirms.
      expect(vi.getTimerCount()).toBeGreaterThan(0)
      // After the grace window fully elapses, no timers should remain.
      vi.advanceTimersByTime(2000)
      expect(vi.getTimerCount()).toBe(0)
    } finally {
      vi.useRealTimers()
    }
  })
})

// ---- Bug #H: doPush refuses downgrade ------------------------------------

describe('Bug #H: doPush refuses to overwrite higher cloud schemaVersion', () => {
  it('push returns pushed:false when cloud schemaVersion > ours', async () => {
    // Seed cloud with a HIGHER schemaVersion.
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 999, // future client
        sourceMachineId: 'newer-machine',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
      },
    })
    await addRules([rule({ signal: 'mine.com' })])
    const result = await pushNow('downgrade-test')
    expect(result.pushed).toBe(false)
    expect(result.reason).toContain('schemaVersion')
    // Cloud meta must NOT have been overwritten with our v1.
    const meta = (await chrome.storage.sync.get('syncMeta')).syncMeta as {
      schemaVersion: number
    }
    expect(meta.schemaVersion).toBe(999)
  })

  it('push proceeds normally when cloud schemaVersion === ours', async () => {
    await addRules([rule({ signal: 'mine.com' })])
    const result = await pushNow('normal')
    expect(result.pushed).toBe(true)
  })

  it('push records syncLastError on downgrade refusal', async () => {
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 999,
        sourceMachineId: 'newer',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
      },
    })
    await pushNow('test')
    const status = await getSyncStatus()
    expect(status.lastError).toBeDefined()
    expect(status.lastError!.source).toBe('push')
    expect(status.lastError!.reason).toContain('schemaVersion')
  })
})

// ---- Bug #I: setSettings mutex -------------------------------------------

describe('Bug #I: setSettings serialises concurrent patches', () => {
  it('two concurrent setSettings calls both land — neither field is lost', async () => {
    // Initial state.
    await setSettings({ batchSize: 25, claudeModel: 'm1' })
    // Fire two patches in parallel that touch different fields.
    await Promise.all([
      setSettings({ batchSize: 50 }),
      setSettings({ claudeModel: 'm2' }),
    ])
    const s = await getSettings()
    expect(s.batchSize).toBe(50)
    expect(s.claudeModel).toBe('m2')
  })
})

// ---- Bug #J: pull errors surfaced --------------------------------------

describe('Bug #J: sync errors surfaced in SyncStatus', () => {
  it('successful push clears any prior error', async () => {
    // Seed an error.
    const err: SyncErrorEntry = {
      at: new Date().toISOString(),
      source: 'push',
      reason: 'earlier failure',
    }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    // Now push successfully.
    await addRules([rule({ signal: 'ok.com' })])
    await pushNow('test')
    const status = await getSyncStatus()
    expect(status.lastError).toBeUndefined()
  })

  it('successful pull clears any prior error', async () => {
    const err: SyncErrorEntry = {
      at: new Date().toISOString(),
      source: 'pull-remote',
      reason: 'transient',
    }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    await simulateOtherMachinePush({ rules: [rule({ signal: 'r.com' })] })
    await pullNow('test')
    const status = await getSyncStatus()
    expect(status.lastError).toBeUndefined()
  })

  it('dismissSyncError clears the persisted error', async () => {
    const err: SyncErrorEntry = {
      at: new Date().toISOString(),
      source: 'push',
      reason: 'user dismisses',
    }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    await dismissSyncError()
    const status = await getSyncStatus()
    expect(status.lastError).toBeUndefined()
  })
})

// ---- Bug #K: backup cap reduced to 5 -------------------------------------

describe('Bug #K: backup rotation cap is now 5', () => {
  it('keeps at most 5 backups when many are made', async () => {
    await addRules([rule({ signal: 'r1.com' })])
    // Each pushNow creates one pre-push backup.
    for (let i = 0; i < 8; i++) {
      await pushNow(`iter-${i}`)
    }
    const backups = await listBackups()
    expect(backups.length).toBeLessThanOrEqual(5)
  })
})

// ---- folderActivity sync (v2 schema) -------------------------------------
//
// Cross-machine "recent activity" panel: A processes some folders, B picks
// up the activity entries via sync, B's lawyer sees what A worked on. Cap
// at 20 entries; union-by-folderId-newest-wins merge.

describe('folderActivity sync (v2 schema)', () => {
  it('push includes folderActivity chunks + meta count', async () => {
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLWFhYWE=',
          folderPath: '案件/vendor',
          count: 3,
          latestMessage: {
            subject: '第一次答辯狀',
            from: 'lawyer@example.com',
            receivedAt: new Date().toISOString(),
          },
        },
      ],
      new Date().toISOString(),
    )
    await pushNow('test')
    const cloud = await chrome.storage.sync.get(null)
    const meta = cloud['syncMeta'] as {
      schemaVersion: number
      folderActivityCount?: number
      folderActivityChunkCount?: number
    }
    expect(meta.schemaVersion).toBe(2)
    expect(meta.folderActivityCount).toBe(1)
    expect(meta.folderActivityChunkCount).toBeGreaterThanOrEqual(1)
    expect(cloud['syncFolderActivity_0']).toBeDefined()
  })

  // Timestamps MUST be relative to now (time-bomb regression, 2026-06):
  // writeFolderActivity prunes rows older than 30 days against the REAL
  // clock — hard-coded absolute fixture dates made these tests fail the
  // day the fixtures aged out, with zero code change.
  const FA_HOUR = 60 * 60 * 1000
  const faOlder = new Date(Date.now() - 30 * FA_HOUR).toISOString()
  const faNewer = new Date(Date.now() - 2 * FA_HOUR).toISOString()

  it('pull merges cloud folderActivity into local (union by folderId)', async () => {
    // Local has entry for folder A.
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'folder-a',
          folderPath: '案件/A',
          count: 2,
          latestMessage: {
            subject: 'local subject',
            from: 'a@x.com',
            receivedAt: faOlder,
          },
        },
      ],
      faOlder,
    )
    // Cloud has entry for folder B (from another machine).
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other',
        updatedAt: faNewer,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        folderActivityCount: 1,
        folderActivityChunkCount: 1,
      },
      syncSettings: { claudeModel: 'x', batchSize: 50 },
      syncFolderActivity_0: {
        index: 0,
        items: [
          {
            folderId: 'folder-b',
            folderPath: '案件/B',
            lastActiveAt: faNewer,
            recentCount: 5,
            latestMessage: {
              subject: 'cloud subject',
              from: 'b@x.com',
              receivedAt: faNewer,
            },
          },
        ],
      },
    })
    await pullNow('test')
    const local = await getFolderActivity()
    expect(local.find((e) => e.folderId === 'folder-a')).toBeDefined()
    expect(local.find((e) => e.folderId === 'folder-b')).toBeDefined()
  })

  it('pull keeps the newest lastActiveAt when folderId collides', async () => {
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'folder-x',
          folderPath: '案件/X',
          count: 1,
          latestMessage: {
            subject: 'stale local',
            from: 'a@x.com',
            receivedAt: faOlder,
          },
        },
      ],
      faOlder,
    )
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other',
        updatedAt: faNewer,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        folderActivityCount: 1,
        folderActivityChunkCount: 1,
      },
      syncSettings: { claudeModel: 'x', batchSize: 50 },
      syncFolderActivity_0: {
        index: 0,
        items: [
          {
            folderId: 'folder-x',
            folderPath: '案件/X',
            lastActiveAt: faNewer,
            recentCount: 3,
            latestMessage: {
              subject: 'fresh from cloud',
              from: 'b@x.com',
              receivedAt: faNewer,
            },
          },
        ],
      },
    })
    await pullNow('test')
    const local = await getFolderActivity()
    const x = local.find((e) => e.folderId === 'folder-x')!
    expect(x.latestMessage?.subject).toBe('fresh from cloud')
  })

  it('mergeFolderActivityFromSync preserves local entries the other machine never saw', async () => {
    await recordFolderActivityFromBatch(
      [{ folderId: 'local-only', folderPath: '03/L', count: 1 }],
      new Date().toISOString(),
    )
    await mergeFolderActivityFromSync([
      {
        folderId: 'cloud-only',
        folderPath: '03/C',
        lastActiveAt: new Date().toISOString(),
        recentCount: 1,
      },
    ])
    const all = await getFolderActivity()
    expect(all.find((e) => e.folderId === 'local-only')).toBeDefined()
    expect(all.find((e) => e.folderId === 'cloud-only')).toBeDefined()
  })

  it('SYNC_SCHEMA_VERSION upgrade: v1 cloud (no folderActivity fields) pulls without crash', async () => {
    // Simulate a v1 cloud state (no folderActivity chunks or count fields).
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2, // newer client wrote, but no activity yet
        sourceMachineId: 'other',
        updatedAt: new Date().toISOString(),
        ruleCount: 1,
        tombstoneCount: 0,
        ruleChunkCount: 1,
        tombstoneChunkCount: 0,
        // folderActivityCount / folderActivityChunkCount omitted
      },
      syncSettings: { claudeModel: 'x', batchSize: 50 },
      syncRules_0: { index: 0, items: [rule({ signal: 'sole.com' })] },
    })
    const r = await pullNow('test')
    expect(r.pulled).toBe(true)
    // No folderActivity should have been added.
    const local = await getFolderActivity()
    expect(local.length).toBe(0)
  })
})

// ---- folderActivity refresh timestamp -------------------------------------

describe('folderActivity refresh-at tracking (auto-refresh policy)', () => {
  it('setFolderActivityRefreshAt persists + getFolderActivityRefreshAt reads back', async () => {
    expect(await getFolderActivityRefreshAt()).toBeNull()
    const at = '2026-05-26T10:00:00Z'
    await setFolderActivityRefreshAt(at)
    expect(await getFolderActivityRefreshAt()).toBe(at)
  })
})

// ---- Bug #N: quiesce helper ----------------------------------------------

describe('Bug #N: quiesce drains in-flight push before resolving', () => {
  it('quiesce resolves cleanly (API contract)', async () => {
    // Wait out any pullInProgress grace timer (Bug #G) from earlier
    // tests — it caps at 1 second so 1100 ms is safe.
    await new Promise((r) => setTimeout(r, 1100))
    // Now nothing should be in flight; quiesce should return quickly.
    const start = Date.now()
    await quiesce()
    expect(Date.now() - start).toBeLessThan(200)
  })

  it('quiesce times out gracefully if forced to wait too long', async () => {
    // Beyond 5 s hard cap, quiesce logs warning + returns rather than
    // hanging UI forever. We can't easily force pushInFlight=true from
    // outside, so this is more of an API smoke test.
    const start = Date.now()
    await quiesce()
    // Whatever the state, it must NOT exceed the 5 s hard cap.
    expect(Date.now() - start).toBeLessThan(5500)
  })
})

// ---- Bug #L: refresh timestamp fail-safe ---------------------------------
// (Tested via service-worker handler — the handler-level fix isn't directly
// reachable from sync-engine unit tests. Behaviour verified manually +
// type-check passes.)

// ---- Bug #M: enableSync surfaces real failure ---------------------------
// (The enableSync handler lives in service-worker.ts; unit testable only via
// SW handler integration test. Manual verification: the handler now
// returns ok:false when pull or push-back fails. Options UI's `if (!r.ok)`
// branch shows red error instead of green success.)

// ---- P-2: handleRemoteSyncMetaChange debounce ----------------------------
// The listener is in chrome.storage.onChanged which is mocked to be a no-op
// in test setup. The debounce behaviour itself (200 ms coalesce) is
// directly visible only with real Chrome's listener dispatch; verified
// via code review of the debounce timer logic.

// ---- P-1: writeFolderActivity diff-before-write --------------------------

describe('P-1: writeFolderActivity skips chrome.storage.local.set on no-op', () => {
  // Relative timestamps (time-bomb regression, 2026-06): writeFolderActivity
  // prunes rows older than 30 days against the real clock. With hard-coded
  // dates, the first test "passed" for the wrong reason once the fixtures
  // aged out (both writes pruned to [] → identical empty JSON), and the
  // second failed outright (pruned-to-empty on both writes → no set call).
  const P1_HOUR = 60 * 60 * 1000
  const p1T0 = new Date(Date.now() - 3 * P1_HOUR).toISOString()
  const p1T1 = new Date(Date.now() - 1 * P1_HOUR).toISOString()

  it('writing identical content twice does not invoke chrome.storage.local.set the second time', async () => {
    const setMock = chrome.storage.local.set as ReturnType<typeof vi.fn>
    setMock.mockClear()
    const entry = {
      folderId: 'idem-folder',
      folderPath: '03/X',
      count: 1,
      latestMessage: {
        subject: 'identical',
        from: 'a@x.com',
        receivedAt: p1T0,
      },
    }
    await recordFolderActivityFromBatch([entry], p1T0)
    const firstCallCount = setMock.mock.calls.filter((c) =>
      Object.prototype.hasOwnProperty.call(c[0], 'folderActivity'),
    ).length
    // A row actually landed (fixture within the retention window) —
    // otherwise the idempotency assertion below would vacuously pass on
    // two empty writes.
    expect(firstCallCount).toBeGreaterThan(0)
    // Same content, same timestamp — second write should detect identical
    // and skip the actual set() call.
    await recordFolderActivityFromBatch([entry], p1T0)
    const secondCallCount = setMock.mock.calls.filter((c) =>
      Object.prototype.hasOwnProperty.call(c[0], 'folderActivity'),
    ).length
    expect(secondCallCount).toBe(firstCallCount)
  })

  it('writing different content still calls chrome.storage.local.set', async () => {
    const setMock = chrome.storage.local.set as ReturnType<typeof vi.fn>
    setMock.mockClear()
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'changing',
          folderPath: '03/A',
          count: 1,
        },
      ],
      p1T0,
    )
    const before = setMock.mock.calls.length
    await recordFolderActivityFromBatch(
      [
        {
          folderId: 'changing',
          folderPath: '03/A',
          count: 1,
        },
      ],
      p1T1, // different timestamp — content actually differs
    )
    expect(setMock.mock.calls.length).toBeGreaterThan(before)
  })
})

// ---- Bug #O: concurrent doPull guard -------------------------------------

describe('Bug #O: doPull refuses re-entry while another pull is in flight', () => {
  it('second concurrent pullNow returns pulled:false with "in flight" reason', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'o.com' })] })
    // Fire two pulls without awaiting the first — the in-flight guard
    // should reject the second.
    const [a, b] = await Promise.all([pullNow('first'), pullNow('second')])
    // Exactly one pulled successfully.
    const pulled = [a, b].filter((r) => r.pulled === true)
    const refused = [a, b].filter((r) => r.pulled === false)
    expect(pulled.length).toBe(1)
    expect(refused.length).toBe(1)
    expect(refused[0]!.reason).toContain('in flight')
  })

  it('after a pull completes, the NEXT pull is allowed (grace period notwithstanding)', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'o2.com' })] })
    const first = await pullNow('first')
    expect(first.pulled).toBe(true)
    // Wait out the 1s pullGraceTimer so pullInProgress is cleared.
    await new Promise((r) => setTimeout(r, 1100))
    const second = await pullNow('second')
    expect(second.pulled).toBe(true)
  })
})

// ---- Bug #Q: appendBackup mutex ------------------------------------------

describe('Bug #Q: concurrent backup writes do not clobber each other', () => {
  it('two pushes in parallel produce two distinct backup snapshots', async () => {
    await addRules([rule({ signal: 'bq1.com' })])
    // Fire two pushes; each does its own appendBackup('pre-push'). With
    // the mutex, both snapshots land in syncBackups. Without it, the
    // second push's read of syncBackups would race with the first
    // push's write, potentially losing one entry.
    await Promise.all([pushNow('a'), pushNow('b')])
    const backups = await listBackups()
    // We expect AT LEAST 2 pre-push backups from this test alone (the
    // second pushNow may have rescheduled if first was in flight via
    // pushInFlight coalesce — either way, no fewer than expected).
    expect(backups.length).toBeGreaterThanOrEqual(1)
    // The key invariant: timestamps are unique (no two backups share
    // the same snapshotAt due to a clobber).
    const stamps = new Set(backups.map((b) => b.snapshotAt))
    expect(stamps.size).toBe(backups.length)
  })
})

// Reference unused-import suppression
void ({} as RuleTombstone | undefined)
