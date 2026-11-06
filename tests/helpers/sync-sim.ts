// Cross-machine sync test helpers (2026-11-05).
//
// "Another machine" is simulated by writing chrome.storage.sync directly in
// the layout that machine's build would produce: the CURRENT (v4) layout via
// the engine's own pure buildPushPayload, or hand-built v3 / v2 layouts for
// schema-skew tests. tests/setup.ts's onChanged mock records listeners but
// never fires them; `syncListener()` digs out the engine's listener so tests
// can deliver change events by hand.

import { vi } from 'vitest'
import {
  buildPushPayload,
  installSyncListener,
  resetSyncEngineForTests,
  type PushPayload,
  type SyncMeta,
} from '@/background/sync-engine'
import { chunkCompactRows, chunkFolderDict, chunkRules, chunkTombstones } from '@/shared/sync-chunks'
import { encodeRulesForSync } from '@/shared/sync-codec'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { FolderActivity, Rule, RuleTombstone, Settings } from '@/shared/types'

let seq = 0

export async function readCloudMeta(): Promise<SyncMeta | undefined> {
  return (await chrome.storage.sync.get('syncMeta')).syncMeta as SyncMeta | undefined
}

export async function cloudKeys(): Promise<string[]> {
  return Object.keys(await chrome.storage.sync.get(null)).sort()
}

async function replaceCloud(writes: Record<string, unknown>): Promise<void> {
  // Like a real push: keys the new state doesn't use are removed.
  const stale = (await cloudKeys()).filter((k) => k.startsWith('sync') && !(k in writes))
  if (stale.length > 0) await chrome.storage.sync.remove(stale)
  await chrome.storage.sync.set(writes)
}

/** Another machine on THIS build pushes its full state (v4 layout). */
export async function simulateOtherMachinePush(opts: {
  rules?: Rule[]
  tombstones?: RuleTombstone[]
  folderActivity?: FolderActivity[]
  settings?: Partial<Settings>
  machineId?: string
  /** Stamp a fresh wipe marker (id defaults to a unique value). */
  wipe?: { id?: string; at?: string }
  pushId?: string
  /** The pusher's active-day ledger (2026-11-06). */
  activeDays?: string[]
}): Promise<PushPayload> {
  const machineId = opts.machineId ?? 'other-machine'
  const updatedAt = new Date().toISOString()
  const payload = buildPushPayload({
    machineId,
    rules: opts.rules ?? [],
    tombstones: opts.tombstones ?? [],
    folderActivity: opts.folderActivity ?? [],
    settings: { ...DEFAULT_SETTINGS, claudeModel: 'claude-opus-5-5', batchSize: 99, ...opts.settings },
    activeDays: opts.activeDays,
    prevMeta: await readCloudMeta(),
    wipeMarker: opts.wipe
      ? { id: opts.wipe.id ?? `wipe-${++seq}`, at: opts.wipe.at ?? updatedAt, byMachineId: machineId }
      : undefined,
    updatedAt,
    pushId: opts.pushId ?? `push${String(++seq).padStart(8, '0')}`,
    now: Date.now(),
  })
  await replaceCloud(payload.writes)
  return payload
}

/** A v1.0.7 machine (schemaVersion 3, compact rows) pushes. */
export async function simulateV3Push(opts: {
  rules?: Rule[]
  tombstones?: RuleTombstone[]
  folderActivity?: FolderActivity[]
  settings?: Record<string, unknown>
  machineId?: string
  wipeMarker?: { at: string; byMachineId: string }
  updatedAt?: string
}): Promise<void> {
  const rules = opts.rules ?? []
  const encoded = encodeRulesForSync(rules)
  // v3 rows: string signals + ISO times (this build's encoder compacts them).
  const rows = encoded.rows.map((row, i) => {
    const r = rules[i]!
    const v3: Record<string, unknown> = { ...(row as object), s: r.signal, ca: r.createdAt }
    delete v3.ea
    if (r.lastUsedAt) v3.lu = r.lastUsedAt
    if (r.autoDisabledAt) v3.da = r.autoDisabledAt
    return v3
  })
  const rowChunks = chunkCompactRows(rows as never)
  const folderChunks = chunkFolderDict(encoded.folders)
  const tombChunks = chunkTombstones(opts.tombstones ?? [])
  const activity = opts.folderActivity ?? []
  const updatedAt = opts.updatedAt ?? new Date().toISOString()
  const writes: Record<string, unknown> = {
    syncMeta: {
      schemaVersion: 3,
      sourceMachineId: opts.machineId ?? 'v3-machine',
      updatedAt,
      ruleCount: rules.length,
      tombstoneCount: (opts.tombstones ?? []).length,
      ruleChunkCount: 0,
      ruleV3ChunkCount: rowChunks.length,
      folderChunkCount: folderChunks.length,
      tombstoneChunkCount: tombChunks.length,
      folderActivityCount: activity.length,
      folderActivityChunkCount: activity.length > 0 ? 1 : 0,
      ...(opts.wipeMarker ? { wipeMarker: opts.wipeMarker } : {}),
    } satisfies SyncMeta,
    syncSettings: opts.settings ?? { claudeModel: 'claude-opus-4-7', batchSize: 33 },
  }
  for (const c of rowChunks) writes[`syncRulesC_${c.index}`] = c
  for (const c of folderChunks) writes[`syncFolders_${c.index}`] = c
  for (const c of tombChunks) writes[`syncTombstones_${c.index}`] = c
  if (activity.length > 0) writes['syncFolderActivity_0'] = { index: 0, items: activity }
  await replaceCloud(writes)
}

/** A v1.0.6 machine (schemaVersion 2, full Rule objects) pushes. */
export async function simulateV2Push(opts: {
  rules?: Rule[]
  tombstones?: RuleTombstone[]
  machineId?: string
}): Promise<void> {
  const rules = opts.rules ?? []
  const ruleChunks = chunkRules(rules)
  const tombChunks = chunkTombstones(opts.tombstones ?? [])
  const writes: Record<string, unknown> = {
    syncMeta: {
      schemaVersion: 2,
      sourceMachineId: opts.machineId ?? 'v2-machine',
      updatedAt: new Date().toISOString(),
      ruleCount: rules.length,
      tombstoneCount: (opts.tombstones ?? []).length,
      ruleChunkCount: ruleChunks.length,
      tombstoneChunkCount: tombChunks.length,
      folderActivityCount: 0,
      folderActivityChunkCount: 0,
    } satisfies SyncMeta,
    syncSettings: { claudeModel: 'claude-haiku-4-5-20251001', batchSize: 22 },
  }
  for (const c of ruleChunks) writes[`syncRules_${c.index}`] = c
  for (const c of tombChunks) writes[`syncTombstones_${c.index}`] = c
  await replaceCloud(writes)
}

/**
 * Several machines on one mock (2026-11-06): chrome.storage.sync is shared,
 * each machine has its own chrome.storage.local. `use(name)` parks the
 * current machine's local storage (JSON round-trip, like Chrome's own
 * serialisation) and loads `name`'s; the engine's in-memory state is reset,
 * as if that machine's service worker had just woken.
 */
export function machines(): { use: (name: string) => Promise<void>; current: () => string | undefined } {
  const parked = new Map<string, string>()
  let current: string | undefined
  return {
    current: () => current,
    async use(name: string) {
      if (current === name) return
      if (current !== undefined) parked.set(current, JSON.stringify(await chrome.storage.local.get(null)))
      await chrome.storage.local.clear()
      const next = parked.get(name)
      if (next) await chrome.storage.local.set(JSON.parse(next) as Record<string, unknown>)
      resetSyncEngineForTests()
      current = name
    },
  }
}

type Listener = (changes: Record<string, chrome.storage.StorageChange>, area: string) => void

/** The engine's onChanged listener (installed fresh; the mock never fires it). */
export function syncListener(): Listener {
  const add = chrome.storage.onChanged.addListener as unknown as ReturnType<typeof vi.fn>
  const before = add.mock.calls.length
  installSyncListener()
  // Drop the startup-reconcile timer the installer arms (it would push into
  // the middle of the test); the listener itself stays registered.
  resetSyncEngineForTests()
  return add.mock.calls[before]![0] as Listener
}
