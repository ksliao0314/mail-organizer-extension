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
  resetSyncEngineForTests,
  restoreBackup,
  SYNC_BACKUPS_KEY,
  SYNC_LAST_ERROR_KEY,
  type SyncErrorEntry,
  type SyncMeta,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import { decodeRulesFromSync } from '@/shared/sync-codec'
import type { FolderDictEntry, WireRuleEntry } from '@/shared/sync-codec'
import { SYNC_SCHEMA_VERSION, SYNC_TOTAL_BUDGET_BYTES } from '@/shared/constants'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'
import {
  cloudKeys,
  readCloudMeta,
  simulateOtherMachinePush,
  simulateV2Push,
  simulateV3Push,
  syncListener,
} from './helpers/sync-sim'

function rule(over: Partial<Rule> = {}): Rule {
  // newRule() ignores id / counters — apply overrides on the result.
  return {
    ...newRule({
      type: 'domain',
      signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.7,
      source: (over.source as RuleSource) ?? 'user_manual',
      enabled: over.enabled,
    }),
    ...over,
  }
}

type SyncChunkLike<T> = { index: number; items: T[]; p?: string }

async function cloudRules(): Promise<Rule[]> {
  const cloud = await chrome.storage.sync.get(null)
  const rows = Object.keys(cloud)
    .filter((k) => k.startsWith('syncRows4_'))
    .map((k) => cloud[k] as SyncChunkLike<WireRuleEntry>)
  const folders = Object.keys(cloud)
    .filter((k) => k.startsWith('syncFolders4_'))
    .map((k) => cloud[k] as SyncChunkLike<FolderDictEntry>)
  return decodeRulesFromSync({ folders: unchunkArray(folders), rows: unchunkArray(rows) })
}

beforeEach(async () => {
  resetSyncEngineForTests()
  // Each test starts with sync enabled + a known machineId.
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- Bug #A: per-device fields don't round-trip ---------------------------

describe('Bug #A: per-device settings fields stay local', () => {
  it('push does NOT write syncMachineId / lastSyncAt / syncEnabled to cloud', async () => {
    await setSettings({ harvestEnabled: true })
    await addRules([rule({ signal: 'a.com' })])
    await pushNow('test')
    const item = (await chrome.storage.sync.get('syncSettings4')).syncSettings4 as {
      v: Record<string, unknown>
    }
    expect(item).toBeDefined()
    const cloudSettings = item.v
    expect('claudeApiKey' in cloudSettings).toBe(false)
    expect('syncMachineId' in cloudSettings).toBe(false)
    expect('lastSyncAt' in cloudSettings).toBe(false)
    expect('syncEnabled' in cloudSettings).toBe(false)
    expect('harvestEnabled' in cloudSettings).toBe(false)
  })

  it('pull does NOT overwrite local syncMachineId / lastSyncAt / syncEnabled', async () => {
    const payload = await simulateOtherMachinePush({ rules: [rule({ signal: 'from-other.com' })] })
    // Defence-in-depth: inject per-device fields into the cloud settings item.
    const item = payload.writes['syncSettings4'] as { p: string; v: Record<string, unknown> }
    await chrome.storage.sync.set({
      syncSettings4: {
        ...item,
        v: { ...item.v, syncMachineId: 'other-machine', syncEnabled: false, lastSyncAt: '2026-01-01T00:00:00Z' },
      },
    })
    await pullNow('test')
    const settings = await getSettings()
    expect(settings.syncMachineId).toBe('this-machine')
    expect(settings.syncEnabled).toBe(true)
    expect(settings.lastSyncAt).not.toBe('2026-01-01T00:00:00Z')
  })

  it('pull DOES overlay other user-facing settings (e.g. claudeModel)', async () => {
    await simulateOtherMachinePush({ rules: [] })
    await pullNow('test')
    const settings = await getSettings()
    // Cloud's claudeModel propagates.
    expect(settings.claudeModel).toBe('claude-opus-5-5')
    expect(settings.batchSize).toBe(99)
  })

  it('pull preserves local claudeApiKey even if cloud somehow has one', async () => {
    await setSettings({ claudeApiKey: 'sk-ant-my-local-key' })
    const payload = await simulateOtherMachinePush({ rules: [] })
    const item = payload.writes['syncSettings4'] as { p: string; v: Record<string, unknown> }
    await chrome.storage.sync.set({
      syncSettings4: { ...item, v: { ...item.v, claudeApiKey: 'sk-ant-malicious-cloud-key' } },
    })
    await pullNow('test')
    expect((await getSettings()).claudeApiKey).toBe('sk-ant-my-local-key')
  })
})

// ---- Bug #B: First-enable UNION merge -------------------------------------

describe('Bug #B: first-enable union pull preserves local-only rules', () => {
  it('union pull keeps local syncable rules missing from cloud', async () => {
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'y.com', source: 'user_manual' })] })
    await pullNow('test', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['x.com', 'y.com'])
  })

  it('replace pull (default) drops local syncable rules missing from cloud', async () => {
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'y.com', source: 'user_manual' })] })
    await pullNow('test') // default mode = 'replace'
    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['y.com'])
  })

  it('立即下載 (manual pull handler) is union: a rule the push budget left out of the cloud survives', async () => {
    const { handlePullSyncNow } = await import('@/background/handlers/sync')
    await addRules([rule({ signal: 'local-only.com', source: 'user_manual', enabled: false })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'cloud.com', source: 'user_manual' })] })
    const r = await handlePullSyncNow()
    expect(r.ok).toBe(true)
    expect((r as { data?: { mergeMode?: string } }).data?.mergeMode).toBe('union')
    expect((await listRules()).map((x) => x.signal).sort()).toEqual(['cloud.com', 'local-only.com'])
  })

  it('立即下載 still applies upstream deletions (tombstones) in union mode', async () => {
    const { handlePullSyncNow } = await import('@/background/handlers/sync')
    const doomed = rule({ signal: 'deleted-elsewhere.com', source: 'user_manual' })
    await addRules([doomed])
    await simulateOtherMachinePush({
      rules: [rule({ signal: 'keep.com' })],
      tombstones: [
        {
          type: doomed.type,
          signalNorm: 'deleted-elsewhere.com',
          targetFolderPath: doomed.targetFolderPath,
          deletedAt: Date.now() + 1000, // after the local copy's edit
        },
      ],
    })
    await handlePullSyncNow()
    expect((await listRules()).map((x) => x.signal).sort()).toEqual(['keep.com'])
  })

  it('union mode preserves per-device rules too (auto_scan)', async () => {
    await addRules([
      rule({ signal: 'local-user.com', source: 'user_manual' }),
      rule({ signal: 'local-scan.com', source: 'auto_scan' }),
    ])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'cloud.com', source: 'user_manual' })] })
    await pullNow('test', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal).sort()).toEqual([
      'cloud.com',
      'local-scan.com',
      'local-user.com',
    ])
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
    await simulateOtherMachinePush({ rules: [rule({ signal: 'after-pull.com' })] })
    await pullNow('test', { mode: 'replace' }) // replace merges snapshot first
    const backups = await listBackups()
    const prePull = backups.find((b) => b.direction === 'pre-pull')
    expect(prePull).toBeDefined()
    // A corrupted backup that carries another machine's identity.
    prePull!.payload.settings.syncMachineId = 'corrupted-id-from-elsewhere'
    prePull!.payload.settings.lastSyncAt = 'some-other-time'
    prePull!.payload.settings.syncEnabled = false
    const all = await listBackups()
    all[all.findIndex((b) => b.snapshotAt === prePull!.snapshotAt)] = prePull!
    await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: all })
    const beforeRestore = await getSettings()

    await restoreBackup(prePull!.snapshotAt)
    const settings = await getSettings()
    expect(settings.syncMachineId).toBe('this-machine')
    expect(settings.syncEnabled).toBe(true)
    expect(settings.lastSyncAt).toBe(beforeRestore.lastSyncAt)
    expect((await listRules()).map((r) => r.signal)).toEqual(['before.com'])
  })
})

// ---- Bug #D: chunk truncation surfaces in push result --------------------

describe('Bug #D: truncated rules + tombstones surfaced in push result', () => {
  it('returns truncatedRuleCount when the library exceeds the byte budget', async () => {
    // Each rule has a DISTINCT ~125-byte folder id (no dedup benefit) + a long
    // signal → ~400 bytes each; 500 of them blow well past the budget.
    const many: Rule[] = []
    for (let i = 0; i < 500; i++) {
      many.push(
        rule({
          id: `id-${i.toString().padStart(4, '0')}`,
          signal: `一一四年度上字第${i}號損害賠償等事件通知書主旨`,
          targetFolderId: 'AAMkADAwMDAwMDAwLWJiYmItY2NjYy1kZGRk' + 'z'.repeat(90) + i,
          targetFolderPath: `01進行中案件/2024年度/客戶第${i}號/往來書狀`,
        }),
      )
    }
    await addRules(many)
    const r = await pushNow('test')
    expect(r.pushed).toBe(true)
    expect(r.truncatedRuleCount).toBeGreaterThan(0)
    // The whole cloud payload stays inside the budget (100 KB hard quota).
    expect(await chrome.storage.sync.getBytesInUse(null)).toBeLessThanOrEqual(SYNC_TOTAL_BUDGET_BYTES)
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
  it('removes all our sync keys (every schema) but leaves local data intact', async () => {
    await addRules([rule({ signal: 'local.com' })])
    await pushNow('test')
    await chrome.storage.sync.set({ syncRulesC_0: { index: 0, items: [] }, syncSettings: {} })
    await clearCloudState()
    expect(await cloudKeys()).toEqual([])
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
    expect(status.cloud!.schemaVersion).toBe(SYNC_SCHEMA_VERSION)
  })

  it('cloud is undefined when no state has been pushed', async () => {
    expect((await getSyncStatus()).cloud).toBeUndefined()
  })
})

// ---- Suppression: tombstones from cloud merge in ------------------------

describe('tombstones round-trip via sync', () => {
  it('cloud tombstones get merged into local on pull', async () => {
    await simulateOtherMachinePush({
      tombstones: [
        { type: 'domain', signalNorm: 'banished.com', targetFolderPath: 'X', deletedAt: Date.now() },
      ],
    })
    await pullNow('test')
    expect((await getRuleTombstones()).some((t) => t.signalNorm === 'banished.com')).toBe(true)
  })
})

// ---- Bug #E: remote-pull union mode + tombstone-aware drop ----------------

describe('Bug #E: remote-pull preserves local-only rules', () => {
  it('union pull preserves local-only syncable rule that cloud never saw', async () => {
    await addRules([rule({ signal: 'local-only.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'from-a.com' })] })
    await pullNow('remote-change-sim', { mode: 'union' })
    const rules = await listRules()
    expect(rules.find((r) => r.signal === 'local-only.com')).toBeDefined()
    expect(rules.find((r) => r.signal === 'from-a.com')).toBeDefined()
  })

  it('union pull WITH matching cloud tombstone drops the local rule (delete propagates)', async () => {
    await addRules([rule({ signal: 'deleted-on-a.com', source: 'user_manual' })])
    await simulateOtherMachinePush({
      tombstones: [
        {
          type: 'domain',
          signalNorm: 'deleted-on-a.com',
          targetFolderPath: 'X/Y',
          deletedAt: Date.now() + 1000,
        },
      ],
    })
    await pullNow('remote-change-sim', { mode: 'union' })
    expect((await listRules()).find((r) => r.signal === 'deleted-on-a.com')).toBeUndefined()
  })

  it('replace pull still drops local-only rules (explicit-intent path)', async () => {
    await addRules([rule({ signal: 'local-only.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'from-a.com' })] })
    await pullNow('manual', { mode: 'replace' })
    const rules = await listRules()
    expect(rules.find((r) => r.signal === 'local-only.com')).toBeUndefined()
    expect(rules.find((r) => r.signal === 'from-a.com')).toBeDefined()
  })
})

// ---- Bug #G: pullInProgress grace period ---------------------------------

describe('Bug #G: pullInProgress grace period (no synchronous clear)', () => {
  it('pull installs a deferred clear (use of setTimeout) — exercised via fake timers', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'g.com' })] })
    vi.useFakeTimers()
    try {
      await pullNow('test')
      expect(vi.getTimerCount()).toBeGreaterThan(0)
      vi.advanceTimersByTime(2000)
      expect(vi.getTimerCount()).toBe(0)
    } finally {
      vi.useRealTimers()
    }
  })
})

// ---- Bug #H: doPush refuses downgrade ------------------------------------

describe('Bug #H: doPush refuses to overwrite higher cloud schemaVersion', () => {
  const future: SyncMeta = {
    schemaVersion: 999,
    sourceMachineId: 'newer-machine',
    updatedAt: new Date().toISOString(),
    ruleCount: 0,
    tombstoneCount: 0,
    ruleChunkCount: 0,
    tombstoneChunkCount: 0,
  }

  it('push returns pushed:false when cloud schemaVersion > ours', async () => {
    await chrome.storage.sync.set({ syncMeta: future })
    await addRules([rule({ signal: 'mine.com' })])
    const result = await pushNow('downgrade-test')
    expect(result.pushed).toBe(false)
    expect(result.reason).toContain('schemaVersion')
    expect(((await readCloudMeta()) as SyncMeta).schemaVersion).toBe(999)
  })

  it('push proceeds normally when cloud schemaVersion === ours', async () => {
    await addRules([rule({ signal: 'mine.com' })])
    expect((await pushNow('normal')).pushed).toBe(true)
  })

  it('push records syncLastError on downgrade refusal', async () => {
    await chrome.storage.sync.set({ syncMeta: future })
    await pushNow('test')
    const status = await getSyncStatus()
    expect(status.lastError!.source).toBe('push')
    expect(status.lastError!.reason).toContain('schemaVersion')
  })
})

// ---- Bug #I: setSettings mutex -------------------------------------------

describe('Bug #I: setSettings serialises concurrent patches', () => {
  it('two concurrent setSettings calls both land — neither field is lost', async () => {
    // Initial state.
    await setSettings({ batchSize: 25, claudeModel: 'claude-haiku-5-5' })
    // Fire two patches in parallel that touch different fields.
    await Promise.all([
      setSettings({ batchSize: 50 }),
      setSettings({ claudeModel: 'claude-opus-5-5' }),
    ])
    const s = await getSettings()
    expect(s.batchSize).toBe(50)
    expect(s.claudeModel).toBe('claude-opus-5-5')
  })
})

// ---- Bug #J: pull errors surfaced --------------------------------------

describe('Bug #J: sync errors surfaced in SyncStatus', () => {
  it('successful push clears any prior error', async () => {
    const err: SyncErrorEntry = { at: new Date().toISOString(), source: 'push', reason: 'earlier failure' }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    await addRules([rule({ signal: 'ok.com' })])
    await pushNow('test')
    expect((await getSyncStatus()).lastError).toBeUndefined()
  })

  it('successful pull clears any prior error', async () => {
    const err: SyncErrorEntry = { at: new Date().toISOString(), source: 'pull-remote', reason: 'transient' }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    await simulateOtherMachinePush({ rules: [rule({ signal: 'r.com' })] })
    await pullNow('test')
    expect((await getSyncStatus()).lastError).toBeUndefined()
  })

  it('dismissSyncError clears the persisted error', async () => {
    const err: SyncErrorEntry = { at: new Date().toISOString(), source: 'push', reason: 'user dismisses' }
    await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: err })
    await dismissSyncError()
    expect((await getSyncStatus()).lastError).toBeUndefined()
  })
})

// ---- folderActivity sync -------------------------------------------------

describe('folderActivity sync', () => {
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
    const meta = cloud['syncMeta'] as SyncMeta
    expect(meta.schemaVersion).toBe(SYNC_SCHEMA_VERSION)
    expect(meta.folderActivityCount).toBe(1)
    expect(meta.v4!.activity).toBeGreaterThanOrEqual(1)
    expect(cloud['syncActivity4_0']).toBeDefined()
  })

  // Timestamps MUST be relative to now (time-bomb regression, 2026-06):
  // writeFolderActivity prunes rows older than 30 days against the REAL clock.
  const FA_HOUR = 60 * 60 * 1000
  const faOlder = new Date(Date.now() - 30 * FA_HOUR).toISOString()
  const faNewer = new Date(Date.now() - 2 * FA_HOUR).toISOString()

  it('pull merges cloud folderActivity into local (union by folderId)', async () => {
    await recordFolderActivityFromBatch(
      [{ folderId: 'folder-a', folderPath: '案件/A', count: 2 }],
      faOlder,
    )
    await simulateOtherMachinePush({
      folderActivity: [
        { folderId: 'folder-b', folderPath: '案件/B', lastActiveAt: faNewer, recentCount: 5 },
      ],
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
          latestMessage: { subject: 'stale local', from: 'a@x.com', receivedAt: faOlder },
        },
      ],
      faOlder,
    )
    await simulateOtherMachinePush({
      folderActivity: [
        {
          folderId: 'folder-x',
          folderPath: '案件/X',
          lastActiveAt: faNewer,
          recentCount: 3,
          latestMessage: { subject: 'fresh from cloud', from: 'b@x.com', receivedAt: faNewer },
        },
      ],
    })
    await pullNow('test')
    const x = (await getFolderActivity()).find((e) => e.folderId === 'folder-x')!
    expect(x.latestMessage?.subject).toBe('fresh from cloud')
  })

  it('mergeFolderActivityFromSync preserves local entries the other machine never saw', async () => {
    await recordFolderActivityFromBatch(
      [{ folderId: 'local-only', folderPath: '03/L', count: 1 }],
      new Date().toISOString(),
    )
    await mergeFolderActivityFromSync([
      { folderId: 'cloud-only', folderPath: '03/C', lastActiveAt: new Date().toISOString(), recentCount: 1 },
    ])
    const all = await getFolderActivity()
    expect(all.find((e) => e.folderId === 'local-only')).toBeDefined()
    expect(all.find((e) => e.folderId === 'cloud-only')).toBeDefined()
  })

  it('cloud without folderActivity pulls without crash', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'sole.com' })] })
    const r = await pullNow('test')
    expect(r.pulled).toBe(true)
    expect((await getFolderActivity()).length).toBe(0)
  })
})

// ---- v4 wire format ---------------------------------------------------------

describe('v4 wire format (schemaVersion 4)', () => {
  it('push writes only v4 keys; every chunk + settings item carries the meta pushId', async () => {
    const rules: Rule[] = []
    for (let i = 0; i < 10; i++) {
      rules.push(
        rule({
          id: `id-${i}`,
          signal: `主旨關鍵字${i}`,
          targetFolderId: `FID_${i % 3}` + 'x'.repeat(100),
          targetFolderPath: `01案件/客戶${i % 3}`,
        }),
      )
    }
    await addRules(rules)
    await pushNow('test')
    const cloud = await chrome.storage.sync.get(null)
    const meta = cloud['syncMeta'] as SyncMeta
    expect(meta.schemaVersion).toBe(4)
    expect(meta.pushId).toBeTruthy()
    expect(meta.ruleChunkCount).toBe(0)
    for (const k of Object.keys(cloud)) {
      if (k === 'syncMeta') continue
      expect(k).toMatch(/^sync(Settings4|Rows4_|Folders4_|Tombs4_|Activity4_)/)
      expect((cloud[k] as { p?: string }).p).toBe(meta.pushId)
    }
    // Folder dict deduped 10 refs → 3 entries; every target survives.
    const folders = unchunkArray(
      Object.keys(cloud)
        .filter((k) => k.startsWith('syncFolders4_'))
        .map((k) => cloud[k] as SyncChunkLike<FolderDictEntry>),
    )
    expect(folders).toHaveLength(3)
    const decoded = await cloudRules()
    expect(decoded).toHaveLength(10)
    for (const r of rules) {
      const d = decoded.find((x) => x.id === r.id)!
      expect(d.targetFolderId).toBe(r.targetFolderId)
      expect(d.targetFolderPath).toBe(r.targetFolderPath)
      expect(d.signal).toBe(r.signal)
      expect(d.editedAt).toBeGreaterThan(0)
    }
  })

  it('a fresh machine pulls v4-encoded cloud rules back', async () => {
    await simulateOtherMachinePush({
      rules: [
        rule({ id: 'a', signal: 'alpha.com', source: 'user_manual' }),
        rule({ id: 'b', signal: 'beta.com', source: 'ai_confirmed' }),
      ],
    })
    await pullNow('test', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['alpha.com', 'beta.com'])
  })

  it('an identical second push writes nothing (content fingerprint unchanged)', async () => {
    await addRules([rule({ signal: 'same.com' })])
    expect((await pushNow('first')).pushed).toBe(true)
    const setMock = chrome.storage.sync.set as ReturnType<typeof vi.fn>
    setMock.mockClear()
    const r = await pushNow('second')
    expect(r).toMatchObject({ pushed: true, unchanged: true })
    expect(setMock).not.toHaveBeenCalled()
  })
})

// ---- Item 7: schema skew --------------------------------------------------

describe('Item 7: older clouds are read and upgraded; newer clouds are refused', () => {
  it('a v3 cloud (v1.0.7 machine) is pulled and union-merged', async () => {
    await addRules([rule({ signal: 'local.com' })])
    await simulateV3Push({
      rules: [rule({ signal: 'from-v3.com' })],
      tombstones: [{ type: 'domain', signalNorm: 'gone.com', targetFolderPath: 'X/Y', deletedAt: Date.now() }],
    })
    const r = await pullNow('test', { mode: 'union' })
    expect(r.pulled).toBe(true)
    expect((await listRules()).map((x) => x.signal).sort()).toEqual(['from-v3.com', 'local.com'])
    expect((await getRuleTombstones()).map((t) => t.signalNorm)).toEqual(['gone.com'])
    expect((await getSettings()).claudeModel).toBe('claude-opus-4-7')
  })

  it('enabling sync against a v2 cloud union-merges it and rewrites the cloud as v4', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await setSettings({ syncEnabled: false })
    await addRules([rule({ signal: 'mine.com' })])
    await simulateV2Push({ rules: [rule({ signal: 'from-v2.com' })] })
    const r = await handleEnableSync()
    // Before: ENABLE_SYNC_FAILED "unsupported cloud schemaVersion=2 (we support 3)".
    expect(r.ok).toBe(true)
    expect((await listRules()).map((x) => x.signal).sort()).toEqual(['from-v2.com', 'mine.com'])
    expect((await readCloudMeta())!.schemaVersion).toBe(SYNC_SCHEMA_VERSION)
    // Legacy keys are gone; the union is in the cloud.
    expect((await cloudKeys()).some((k) => k.startsWith('syncRules_') || k === 'syncSettings')).toBe(false)
    expect((await cloudRules()).map((x) => x.signal).sort()).toEqual(['from-v2.com', 'mine.com'])
  })

  it('enabling against a NEWER cloud fails with an upgrade-this-machine message', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await setSettings({ syncEnabled: false })
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: SYNC_SCHEMA_VERSION + 1,
        sourceMachineId: 'future',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
      } satisfies SyncMeta,
    })
    const r = await handleEnableSync()
    expect(r.ok).toBe(false)
    expect((r as { message?: string }).message).toContain('較新版本')
    expect((await getSettings()).syncEnabled).toBe(false)
  })

  it('a schema refusal on the REMOTE-change path is recorded in lastError', async () => {
    const listener = syncListener()
    vi.useFakeTimers()
    try {
      const meta: SyncMeta = {
        schemaVersion: SYNC_SCHEMA_VERSION + 1,
        sourceMachineId: 'future',
        updatedAt: new Date().toISOString(),
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
      }
      await chrome.storage.sync.set({ syncMeta: meta })
      listener({ syncMeta: { newValue: meta } }, 'sync')
      await vi.advanceTimersByTimeAsync(500)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    const status = await getSyncStatus()
    expect(status.lastError?.source).toBe('pull-remote')
    expect(status.lastError?.reason).toContain('newer')
  })
})

// ---- Item 8: backups ------------------------------------------------------

describe('Item 8: backups only before destructive merges, tombstone-free, pinned pre-wipe slot', () => {
  it('ordinary pushes and union pulls write no backup', async () => {
    await addRules([rule({ signal: 'r1.com' })])
    for (let i = 0; i < 3; i++) await pushNow(`iter-${i}`)
    await simulateOtherMachinePush({ rules: [rule({ signal: 'r2.com' })] })
    await pullNow('remote', { mode: 'union' })
    expect(await listBackups()).toEqual([])
  })

  it('a replace pull snapshots first, without tombstones; rotation keeps MAX_BACKUPS', async () => {
    await addRules([rule({ signal: 'r1.com' })])
    await addRuleTombstones(
      Array.from({ length: 50 }, (_, i) => ({
        type: 'domain' as const,
        signalNorm: `t${i}.com`,
        targetFolderPath: 'X',
        deletedAt: 1000 + i,
      })),
    )
    await simulateOtherMachinePush({ rules: [rule({ signal: 'cloud.com' })] })
    for (let i = 0; i < 7; i++) await pullNow(`replace-${i}`, { mode: 'replace' })
    const backups = await listBackups()
    expect(backups.length).toBe(5)
    expect(backups.every((b) => b.payload.tombstones.length === 0)).toBe(true)
    expect(backups[0]!.trigger).toBe('replace-pull')
  })

  it('a backup failure is recorded in lastError and the pull still completes', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'cloud.com' })] })
    const setMock = chrome.storage.local.set as ReturnType<typeof vi.fn>
    const real = setMock.getMockImplementation() as (items: Record<string, unknown>) => Promise<void>
    setMock.mockImplementation(async (items: Record<string, unknown>) => {
      if (SYNC_BACKUPS_KEY in items) throw new Error('QUOTA_BYTES quota exceeded')
      return real(items)
    })
    try {
      const r = await pullNow('test', { mode: 'replace' })
      expect(r.pulled).toBe(true)
    } finally {
      setMock.mockImplementation(real)
    }
    expect((await listRules()).map((x) => x.signal)).toEqual(['cloud.com'])
    const status = await getSyncStatus()
    expect(status.lastError?.reason).toContain('backup failed')
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

describe('Bug #N: quiesce drains in-flight ops before resolving', () => {
  it('quiesce resolves quickly when nothing is in flight', async () => {
    const start = Date.now()
    await quiesce()
    expect(Date.now() - start).toBeLessThan(200)
  })

  it('quiesce waits for an in-flight push to land before resolving', async () => {
    await addRules([rule({ signal: 'inflight.com' })])
    const push = pushNow('in-flight')
    await quiesce()
    // The push finished (cloud written) before quiesce returned.
    expect((await readCloudMeta())?.sourceMachineId).toBe('this-machine')
    await push
  })
})

// ---- P-1: writeFolderActivity diff-before-write --------------------------

describe('P-1: writeFolderActivity skips chrome.storage.local.set on no-op', () => {
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
      latestMessage: { subject: 'identical', from: 'a@x.com', receivedAt: p1T0 },
    }
    await recordFolderActivityFromBatch([entry], p1T0)
    const firstCallCount = setMock.mock.calls.filter((c) =>
      Object.prototype.hasOwnProperty.call(c[0], 'folderActivity'),
    ).length
    expect(firstCallCount).toBeGreaterThan(0)
    await recordFolderActivityFromBatch([entry], p1T0)
    const secondCallCount = setMock.mock.calls.filter((c) =>
      Object.prototype.hasOwnProperty.call(c[0], 'folderActivity'),
    ).length
    expect(secondCallCount).toBe(firstCallCount)
  })

  it('writing different content still calls chrome.storage.local.set', async () => {
    const setMock = chrome.storage.local.set as ReturnType<typeof vi.fn>
    setMock.mockClear()
    await recordFolderActivityFromBatch([{ folderId: 'changing', folderPath: '03/A', count: 1 }], p1T0)
    const before = setMock.mock.calls.length
    await recordFolderActivityFromBatch([{ folderId: 'changing', folderPath: '03/A', count: 1 }], p1T1)
    expect(setMock.mock.calls.length).toBeGreaterThan(before)
  })
})

// ---- H1: quota-wedge recovery ---------------------------------------------

describe('H1: doPush quota-wedge recovery', () => {
  const QUOTA_ERR = 'QUOTA_BYTES quota exceeded'

  async function seedLocalAndOrphans() {
    await addRules([rule({ signal: 'wedge.com' })])
    // Stranded chunks from an earlier, larger library.
    await chrome.storage.sync.set({
      syncRules_1: { index: 1, items: [] },
      syncRules_2: { index: 2, items: [] },
    })
  }

  it('quota failure with orphans stored → removes orphans, retries once, push succeeds', async () => {
    await seedLocalAndOrphans()
    const setMock = chrome.storage.sync.set as ReturnType<typeof vi.fn>
    setMock.mockImplementationOnce(async () => {
      throw new Error(QUOTA_ERR)
    })
    const r = await pushNow('test')
    expect(r.pushed).toBe(true)
    const cloud = await chrome.storage.sync.get(null)
    expect(cloud['syncRows4_0']).toBeDefined()
    expect(cloud['syncRules_1']).toBeUndefined()
    expect(cloud['syncRules_2']).toBeUndefined()
  })

  it('non-quota failure → no orphan removal, push reports failure', async () => {
    await seedLocalAndOrphans()
    const setMock = chrome.storage.sync.set as ReturnType<typeof vi.fn>
    setMock.mockImplementationOnce(async () => {
      throw new Error('network down')
    })
    const r = await pushNow('test')
    expect(r.pushed).toBe(false)
    const cloud = await chrome.storage.sync.get(null)
    expect(cloud['syncRules_1']).toBeDefined()
    expect(cloud['syncRules_2']).toBeDefined()
  })

  it('quota failure persisting after orphan removal → push reports failure', async () => {
    await seedLocalAndOrphans()
    const setMock = chrome.storage.sync.set as ReturnType<typeof vi.fn>
    setMock
      .mockImplementationOnce(async () => {
        throw new Error(QUOTA_ERR)
      })
      .mockImplementationOnce(async () => {
        throw new Error(QUOTA_ERR)
      })
    const r = await pushNow('test')
    expect(r.pushed).toBe(false)
    expect((await getSettings()).lastSyncAt).toBe('')
  })
})

// ---- Item 10: write-rate limits ---------------------------------------------

describe('Item 10: write-rate limit errors back off instead of deleting orphans', () => {
  it('MAX_WRITE_OPERATIONS_PER_MINUTE → no orphan removal, no immediate retry, retried after the backoff', async () => {
    await addRules([rule({ signal: 'rate.com' })])
    await chrome.storage.sync.set({ syncRules_7: { index: 7, items: [] } })
    const setMock = chrome.storage.sync.set as ReturnType<typeof vi.fn>
    setMock.mockImplementationOnce(async () => {
      throw new Error('This request exceeds the MAX_WRITE_OPERATIONS_PER_MINUTE quota.')
    })
    vi.useFakeTimers()
    try {
      const r = await pushNow('test')
      expect(r.pushed).toBe(false)
      // Before: /quota/i matched → orphans removed + set retried at once.
      expect((await chrome.storage.sync.get('syncRules_7')).syncRules_7).toBeDefined()
      expect((await getSyncStatus()).lastError?.reason).toContain('MAX_WRITE_OPERATIONS_PER_MINUTE')
      expect(await readCloudMeta()).toBeUndefined()
      await vi.advanceTimersByTimeAsync(30_000)
      expect(await readCloudMeta()).toBeUndefined() // still backing off
      await vi.advanceTimersByTimeAsync(40_000)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    expect((await readCloudMeta())?.sourceMachineId).toBe('this-machine')
    expect((await getSyncStatus()).lastError).toBeUndefined()
  })

  it('describeSyncError tells a rate-limited user to wait, never to delete rules', async () => {
    const { describeSyncError } = await import('@/options/App')
    const minute = describeSyncError('MAX_WRITE_OPERATIONS_PER_MINUTE quota exceeded', 'push')
    expect(minute.hint).not.toContain('清掉')
    expect(minute.cta?.action).toBe('retry')
    const hour = describeSyncError('MAX_WRITE_OPERATIONS_PER_HOUR quota exceeded', 'push')
    expect(hour.cta?.action).toBe('retry')
    expect(hour.hint).toContain('小時')
    expect(describeSyncError('QUOTA_BYTES quota exceeded', 'push').cta?.action).toBe('cleanup')
    expect(
      describeSyncError('cloud snapshot incomplete: syncRows4_0 missing or from another push', 'pull-remote').hint,
    ).toContain('自動')
  })
})

// ---- Item 11: harvestEnabled is per-device ------------------------------------

describe('Item 11: toggling harvestEnabled schedules no push', () => {
  it('a harvestEnabled-only settings change is ignored; a synced field schedules a push', async () => {
    const listener = syncListener()
    vi.useFakeTimers()
    try {
      const base = { claudeModel: 'm', harvestEnabled: false }
      const before = vi.getTimerCount()
      listener({ settings: { oldValue: base, newValue: { ...base, harvestEnabled: true } } }, 'local')
      expect(vi.getTimerCount()).toBe(before)
      listener({ settings: { oldValue: base, newValue: { ...base, claudeModel: 'm2' } } }, 'local')
      expect(vi.getTimerCount()).toBe(before + 1)
    } finally {
      vi.useRealTimers()
      resetSyncEngineForTests()
    }
  })
})

// ---- Item 13c: orphan keys from the previous meta ---------------------------

describe('Item 13c: orphans come from the previous meta, not a full scan per push', () => {
  it('only the first push of a SW lifetime reads every key', async () => {
    await addRules([rule({ signal: 'a.com' })])
    const getMock = chrome.storage.sync.get as ReturnType<typeof vi.fn>
    await pushNow('first')
    await addRules([rule({ signal: 'b.com' })])
    getMock.mockClear()
    await pushNow('second')
    expect(getMock.mock.calls.some((c) => c[0] === null || c[0] === undefined)).toBe(false)
  })

  it('a shrinking library still removes the chunks the previous meta referenced', async () => {
    const big: Rule[] = []
    for (let i = 0; i < 120; i++) {
      big.push(rule({ id: `big-${i}`, signal: `長主旨關鍵字第${i}號事件通知書與相關附件`, targetFolderPath: `案件/${i}` }))
    }
    await addRules(big)
    await pushNow('big')
    const bigMeta = (await readCloudMeta())!
    expect(bigMeta.v4!.rows).toBeGreaterThan(1)
    const { mutateRules } = await import('@/shared/rules')
    await mutateRules((rs) => ({ next: rs.slice(0, 1), result: undefined }))
    await pushNow('small')
    const keys = await cloudKeys()
    expect(keys.filter((k) => k.startsWith('syncRows4_'))).toEqual(['syncRows4_0'])
    expect(keys.filter((k) => k.startsWith('syncFolders4_'))).toEqual(['syncFolders4_0'])
  })
})

// ---- Bug #O → one queue: pulls are serialised, never refused ---------------

describe('Bug #O: concurrent pulls run one at a time (none refused)', () => {
  it('two concurrent pullNow calls both complete without interleaving', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'o.com' })] })
    const [a, b] = await Promise.all([pullNow('first'), pullNow('second')])
    expect(a.pulled).toBe(true)
    expect(b.pulled).toBe(true)
    expect((await listRules()).map((r) => r.signal)).toEqual(['o.com'])
  })

  it('after a pull completes, the NEXT pull is allowed immediately', async () => {
    await simulateOtherMachinePush({ rules: [rule({ signal: 'o2.com' })] })
    expect((await pullNow('first')).pulled).toBe(true)
    expect((await pullNow('second')).pulled).toBe(true)
  })
})

// ---- Bug #Q: appendBackup mutex ------------------------------------------

describe('Bug #Q: concurrent backup writes do not clobber each other', () => {
  it('two replace pulls in parallel produce two distinct backup snapshots', async () => {
    await addRules([rule({ signal: 'bq1.com' })])
    await simulateOtherMachinePush({ rules: [rule({ signal: 'bq2.com' })] })
    await Promise.all([pullNow('a', { mode: 'replace' }), pullNow('b', { mode: 'replace' })])
    const backups = await listBackups()
    expect(backups.length).toBe(2)
    expect(new Set(backups.map((b) => b.snapshotAt)).size).toBe(2)
  })
})

// ---- Item 14: settings only through setSettings ----------------------------

describe('Item 14: the engine writes settings only through setSettings', () => {
  it('sync-engine.ts has no raw settings / rules writes', async () => {
    const { readFileSync } = await import('node:fs')
    const { resolve } = await import('node:path')
    const code = readFileSync(resolve(__dirname, '../src/background/sync-engine.ts'), 'utf8')
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/\/\/.*$/gm, '')
    expect(code).not.toMatch(/setRules\(/)
    expect(code).not.toMatch(/local\.set\(\{\s*\[?['"]?settings/)
    expect(code).not.toMatch(/local\.set\(\{\s*\[?['"]?rules/)
  })
})

void ({} as RuleTombstone | undefined)
