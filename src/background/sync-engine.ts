// Cross-machine sync engine via chrome.storage.sync.
//
// What gets synced:
//   - Rules where source ∈ {user_manual, ai_confirmed, ai_overridden}
//     and !orphaned. (auto_scan rules are cheap to re-derive per
//     machine via initial scan.)
//   - Recent tombstones (capped to TOMBSTONE_SYNC_CAP for quota).
//   - Settings minus claudeApiKey (never leaves the device).
//
// What stays per-device:
//   - claudeApiKey (security)
//   - folderCache (different Outlook IDs even if same mailbox)
//   - folderActivity (per-device execution log)
//   - skipHistory (per-device preview choices)
//   - conversationMemory / subjectMemory (per-device threading)
//
// Conflict semantics: chrome.storage.sync is per-item last-writer-wins.
// We accept that. Before every push/pull we snapshot local state into
// `syncBackups` (chrome.storage.local, last 10 rotation) so user can
// rollback if something looks wrong.
//
// Detection: every push writes `syncMeta.sourceMachineId = ourId`. The
// chrome.storage.onChanged listener ignores events where the changed
// `syncMeta` carries our own id (= our echo), pulls otherwise.

import {
  addRuleTombstones,
  clearAllRuleTombstones,
  getFolderActivity,
  getRuleTombstones,
  getSettings,
  mergeFolderActivityFromSync,
  setRules,
  setSettings,
} from '@/shared/storage'
import { applyConfidenceCap, listRules, mutateRules, normalizeSignal } from '@/shared/rules'
import {
  chunkFolderActivity,
  chunkRules,
  chunkTombstones,
  shouldSyncRule,
  unchunkArray,
  type SyncChunk,
} from '@/shared/sync-chunks'
import {
  MAX_BACKUPS,
  MAX_FOLDER_ACTIVITY_CHUNKS,
  MAX_RULE_CHUNKS,
  MAX_TOMBSTONE_CHUNKS,
  PULL_DEBOUNCE_MS,
  PULL_GRACE_MS,
  PUSH_DEBOUNCE_MS,
  QUIESCE_TIMEOUT_MS,
  SYNC_RECENT_PUSHES_CAP,
  SYNC_SCHEMA_VERSION,
} from '@/shared/constants'
import { logError } from '@/shared/error-log'
import type { FolderActivity, Rule, RuleTombstone, Settings } from '@/shared/types'

/**
 * Fields that must NEVER round-trip through chrome.storage.sync — each
 * device has its own value and importing another device's value would
 * cause identity corruption / echo loops / UX bugs:
 *
 *   - claudeApiKey: security; never leaves the device.
 *   - syncMachineId: identifies "us" for echo prevention. If A's id
 *     overwrites B's via settings sync, B starts pushing under A's id
 *     and the onChanged listener can't tell them apart.
 *   - lastSyncAt: each device tracks its own last-sync time. The
 *     cloud's view ≠ this device's view.
 *   - syncEnabled: per-device toggle. A enabling doesn't mean B opted in.
 *   - onboardingDismissed (Bug #T): each device has its own "have I
 *     finished setup HERE yet" answer. A's user finishing the wizard
 *     should NOT make B's user skip it — B still needs to set its own
 *     API key (which is per-device too) and the wizard is the
 *     intended guide. Without this, B silently shows an empty popup
 *     with a "缺 Claude API key" warning and no guidance, defeating
 *     the wizard's purpose.
 */
const PER_DEVICE_SETTINGS_FIELDS: ReadonlySet<keyof Settings> = new Set([
  'claudeApiKey',
  'syncMachineId',
  'lastSyncAt',
  'syncEnabled',
  'onboardingDismissed',
])

type PerDeviceFieldKey =
  | 'claudeApiKey'
  | 'syncMachineId'
  | 'lastSyncAt'
  | 'syncEnabled'
  | 'onboardingDismissed'

/** Project a Settings record down to only the fields safe to sync. */
function stripPerDeviceSettings(s: Settings): Omit<Settings, PerDeviceFieldKey> {
  const out: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(s)) {
    if (!PER_DEVICE_SETTINGS_FIELDS.has(k as keyof Settings)) {
      out[k] = v
    }
  }
  return out as Omit<Settings, PerDeviceFieldKey>
}

// ---- sync-storage keys -----------------------------------------------------

const SYNC_META_KEY = 'syncMeta'
const SYNC_SETTINGS_KEY = 'syncSettings'
const SYNC_RULES_CHUNK_PREFIX = 'syncRules_'
const SYNC_TOMBSTONES_CHUNK_PREFIX = 'syncTombstones_'
const SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX = 'syncFolderActivity_'

// MAX_RULE_CHUNKS / MAX_TOMBSTONE_CHUNKS / MAX_FOLDER_ACTIVITY_CHUNKS:
// see src/shared/constants.ts. Chunk-count caps prevent the key namespace
// from blowing up when a single push has many chunks.

// ---- local-only meta -------------------------------------------------------

export const SYNC_BACKUPS_KEY = 'syncBackups'
// MAX_BACKUPS: see src/shared/constants.ts. 5 snapshots × full payload
// keeps us under chrome.storage.local 5 MB even with large rule libs.

// Local key tracking the last sync failure so the options UI can surface
// "sync silently broken" to the user instead of just logging to console.
export const SYNC_LAST_ERROR_KEY = 'syncLastError'

// Local key for the cross-machine wipe notice — set by doPull when it
// detects a fresh wipeMarker from another machine and applies it
// locally. The Options UI reads this to surface a banner explaining
// "machine X cleared everything on YYYY-MM-DD; we did the same here.
// Roll back from backup if this wasn't intended." Dismissed via
// acknowledgeRemoteWipeNotice() which clears the entry.
export const REMOTE_WIPE_NOTICE_KEY = 'remoteWipeNotice'

export type RemoteWipeNotice = {
  /** Machine that issued the wipe (from SyncMeta.wipeMarker.byMachineId). */
  byMachineId: string
  /** When the wipe was issued (cloud-side ISO). */
  at: string
  /** When THIS machine applied the wipe. */
  appliedAt: string
  /**
   * Number of syncable rules dropped during the apply step — helps
   * the user judge "what did we lose?" before deciding whether to
   * rollback from backup.
   */
  droppedRuleCount: number
}

export type SyncErrorEntry = {
  at: string
  /** 'push' | 'pull' | 'pull-remote' — which path failed. */
  source: 'push' | 'pull' | 'pull-remote'
  reason: string
}

export type SyncBackup = {
  snapshotAt: string
  direction: 'pre-push' | 'pre-pull' | 'manual'
  /** Local snapshot of what would be overwritten — minus per-machine state. */
  payload: {
    rules: Rule[]
    tombstones: RuleTombstone[]
    settings: Omit<Settings, 'claudeApiKey'>
  }
}

// ---- cloud envelope --------------------------------------------------------

/** One entry in syncMeta.recentPushes — see SyncMeta. */
export type SyncRecentPush = {
  machineId: string
  at: string
}

export type SyncMeta = {
  /** Bump on schema change so old clients can refuse. */
  schemaVersion: number
  /** Machine that produced this cloud state. Used to recognise our echoes. */
  sourceMachineId: string
  /** ISO timestamp of when this state was written. */
  updatedAt: string
  /** Total rule + tombstone counts — quick health check before reading chunks. */
  ruleCount: number
  tombstoneCount: number
  /** Number of chunks each category was split into (for fan-out reads). */
  ruleChunkCount: number
  tombstoneChunkCount: number
  /**
   * folderActivity sync was added in schemaVersion 2. Optional on the type
   * because v1 cloud snapshots predate it — readers default to 0 when
   * absent (no chunks to fetch, no entries to merge).
   */
  folderActivityCount?: number
  folderActivityChunkCount?: number
  /**
   * Rolling log of recent push events from ANY machine that has pushed
   * to this cloud state (cap SYNC_RECENT_PUSHES_CAP, newest first). Lets
   * the Options UI show "A 推了 5 分鐘前 · B 推了 2 小時前" so the user
   * understands multi-machine activity at a glance. Each push appends
   * its own (machineId, at) entry and inherits the others from the
   * previous syncMeta read. Optional for backward compat with v1 / v2
   * snapshots that predated this field.
   */
  recentPushes?: SyncRecentPush[]
  /**
   * Cross-machine wipe propagation marker. Set by `wipeAllRules` →
   * `pushNow('post-wipe', { wipeMarker: true })`. Subsequent pulls on
   * OTHER machines compare `wipeMarker.at` to their local
   * `settings.lastSyncAt`; when fresher, they:
   *   1. Clear their own syncable rules + tombstones (auto_scan stays
   *      because those re-derive per machine via initial scan).
   *   2. Record a notice in chrome.storage.local
   *      (`remoteWipeNotice`) for the Options UI to surface "machine
   *      X cleared the library on YYYY-MM-DD; rolled back from
   *      backup if needed".
   *
   * The marker persists across subsequent pushes (preserved from
   * cloudMeta.wipeMarker) so a third machine that connects later
   * still trips the wipe. Replaced only by a fresh wipe from any
   * machine.
   *
   * Echo guard: a machine never trips its own wipeMarker
   * (byMachineId !== syncMachineId). Staleness guard: a machine
   * never re-trips a marker it already applied
   * (marker.at <= lastSyncAt).
   *
   * Optional for backward compat with pre-2026-05-27 cloud snapshots.
   */
  wipeMarker?: {
    /** ISO of when the wipe was issued. */
    at: string
    /** Issuing machine — pulls echo-guard on this. */
    byMachineId: string
  }
}

// SYNC_SCHEMA_VERSION: see src/shared/constants.ts.
// v2 (2026-05-26): folderActivity sync.
//   v1 → v2: added folderActivity chunks + meta fields. Old v1 clients
//   refuse to pull from v2 cloud (correct — they wouldn't apply the new
//   chunks). New v2 clients refuse to push over v2 cloud only if cloud's
//   version is HIGHER (downgrade protection, Bug #H); pushing v2 over v1
//   is fine (upgrade).

// ---- state -----------------------------------------------------------------

// Bug #Q: appendBackup does read-modify-write on syncBackups in
// chrome.storage.local. Two concurrent appendBackup calls (e.g. a
// doPush's 'pre-push' overlapping with a doPull's 'pre-pull', or two
// remote-pull dispatches that slipped through before Bug #O's guard
// landed) would each read the same N-entry list, prepend their own,
// slice to MAX_BACKUPS, and write — last writer wins, losing the
// other's snapshot. Serialize the whole RMW via this chain.
let backupWriteChain: Promise<void> = Promise.resolve()

async function withBackupLock<T>(fn: () => Promise<T>): Promise<T> {
  const release = backupWriteChain
  let resolve: () => void = () => {}
  backupWriteChain = new Promise<void>((r) => {
    resolve = r
  })
  try {
    await release
    return await fn()
  } finally {
    resolve()
  }
}

// PUSH_DEBOUNCE_MS: see src/shared/constants.ts. Coalesces rapid
// rule/tombstone/folderActivity writes into a single push.

let pushTimer: ReturnType<typeof setTimeout> | null = null
let pushInFlight = false
// Debounce timer for remote-pull dispatch. If another machine pushes
// multiple times in rapid succession (e.g. burst of rule edits, or
// initial-scan generating ai_confirmed across many emails), we collapse
// the resulting onChanged events into a single pull. 200 ms covers
// typical bursts without delaying the user-visible "another machine
// just updated" feedback noticeably.
let pullDebounceTimer: ReturnType<typeof setTimeout> | null = null
let pendingPullMeta: SyncMeta | undefined
// Two pull flags, separating different concerns (Bug #O fix):
//
//   - pullActive: true ONLY while doPull's try block (writes) is running.
//     New doPull calls bail when this is true — that's the re-entry
//     guard preventing concurrent pulls from racing pullGraceTimer.
//
//   - pullInProgress: true from pull start through the grace period
//     (1 s after pull's writes complete). The local-storage listener
//     uses this to suppress echo pushes for the pull's own writes,
//     including listeners that fire as macrotasks AFTER the pull's
//     finally has run.
//
// Why both: the re-entry guard needs to clear quickly so the next
// genuine pull (e.g. a manual click 1.5s later) isn't blocked. The
// echo-suppression needs to last past macrotask drain. Different
// lifetimes → different flags.
let pullActive = false
let pullInProgress = false
// Pending timer that will clear pullInProgress AFTER macrotask-dispatched
// onChanged listeners from the pull's own writes have fired. See
// PULL_GRACE_MS comment + the doPull finally block for why we don't just
// clear synchronously.
let pullGraceTimer: ReturnType<typeof setTimeout> | null = null

// PULL_GRACE_MS: see src/shared/constants.ts. Window during which the
// local-mutation listener treats writes as the pull's own (suppresses
// echo push). Outlives macrotask dispatch of pull's onChanged events.

// ---- public API ------------------------------------------------------------

/**
 * Installer — called on module load AND on `chrome.runtime.onStartup` so
 * the listener survives service-worker restarts.
 *
 * Two listeners:
 *   1. `chrome.storage.onChanged` for the 'sync' area → detect REMOTE
 *      changes from another machine and pull them in.
 *   2. `chrome.storage.onChanged` for the 'local' area → detect LOCAL
 *      rule / tombstone / settings mutations (which all flow through
 *      chrome.storage.local) and schedule a debounced push. This way
 *      we don't need to wire schedulePush() into every individual
 *      mutation site (upsertRule / deleteRule / autoDisableStaleRules /
 *      initial-scan / ai_confirmed generation / etc.) — they all write
 *      to local storage eventually, so a single listener catches all.
 */
export function installSyncListener(): void {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync') {
      // Remote change pulled in. Filter to syncMeta — chunk writes
      // always come paired with a syncMeta bump.
      if (!changes[SYNC_META_KEY]) return
      // Synchronous debounce-schedule; the actual async work runs in
      // dispatchRemotePull after the timer.
      handleRemoteSyncMetaChange(changes[SYNC_META_KEY].newValue as SyncMeta | undefined)
      return
    }
    if (area === 'local') {
      // Don't push echoes from our own pull (we'd write the same data
      // back up). pullInProgress is set/cleared inside doPull.
      if (pullInProgress) return
      // Local mutation. Schedule push if it touched data we sync.
      // Settings changes that ONLY touch lastSyncAt are our own
      // bookkeeping — filter them out to avoid push loops.
      //
      // folderActivity changes too: a batch execution or the user
      // clicking 重新整理 updates folderActivity, which the OTHER machine
      // wants to see. Same 5s debounce coalesces these with any
      // accompanying rule mutations.
      if (
        changes['rules'] ||
        changes['ruleTombstones'] ||
        changes['folderActivity'] ||
        isUserFacingSettingsChange(changes['settings'])
      ) {
        schedulePush('local-mutation')
      }
      return
    }
  })
}

/**
 * Returns true when the settings change touched fields the user cares
 * about syncing (vs internal bookkeeping like lastSyncAt). Without this
 * guard, every `setSettings({ lastSyncAt: ... })` triggers another
 * push, which triggers another settings write, which... loop.
 */
function isUserFacingSettingsChange(change?: chrome.storage.StorageChange): boolean {
  if (!change) return false
  const oldS = (change.oldValue as Partial<Settings> | undefined) ?? {}
  const newS = (change.newValue as Partial<Settings> | undefined) ?? {}
  // Compare every settings field EXCEPT per-device / bookkeeping ones.
  // If any actually-synced field changed, return true.
  //
  // syncEnabled stays IN the comparison: flipping it should trigger
  // the enable-flow push (which also explicitly calls pushNow). The
  // OTHER per-device fields are filtered because their changes
  // shouldn't cause sync churn — they don't affect cloud at all.
  const IGNORED_FIELDS = new Set<keyof Settings>([
    'lastSyncAt',
    'syncMachineId',
    'claudeApiKey',
    'onboardingDismissed', // Bug #T: per-device wizard state
  ])
  const allKeys = new Set([...Object.keys(oldS), ...Object.keys(newS)])
  for (const k of allKeys) {
    if (IGNORED_FIELDS.has(k as keyof Settings)) continue
    if (JSON.stringify((oldS as Record<string, unknown>)[k]) !== JSON.stringify((newS as Record<string, unknown>)[k])) {
      return true
    }
  }
  return false
}

/**
 * Schedule a debounced push. Safe to call on every mutation — coalesces
 * into a single push per debounce window. No-op if sync disabled.
 */
export function schedulePush(reason: string): void {
  if (pushTimer) clearTimeout(pushTimer)
  pushTimer = setTimeout(() => {
    pushTimer = null
    void doPush(reason).catch((e) =>
      console.warn(`[mail-organizer] sync push failed (${reason})`, e),
    )
  }, PUSH_DEBOUNCE_MS)
}

/**
 * Force-push now (manual button / first-enable / post-wipe). Bypasses
 * debounce.
 *
 * `options.wipeMarker = true` stamps `wipeMarker` into the pushed
 * SyncMeta — used by wipeAllRules so other machines pull, see the
 * fresh marker, and apply the wipe locally. Without this, "全部刪除"
 * on machine A would only clean A; B's next push would re-acquire A
 * the very rules A just deleted.
 */
export async function pushNow(
  reason: string,
  options: { wipeMarker?: boolean } = {},
): Promise<{
  pushed: boolean
  reason?: string
  truncatedRuleCount?: number
  truncatedTombstoneCount?: number
}> {
  if (pushTimer) {
    clearTimeout(pushTimer)
    pushTimer = null
  }
  return doPush(reason, options)
}

/**
 * Pull cloud state into local — manual button / startup / remote-change
 * trigger. Always snapshots local state to backup first.
 *
 * Default mode is 'replace' (steady-state). Pass mode='union' for
 * first-enable scenarios where local has rules the cloud machine
 * never saw — dropping those would be data loss.
 */
export async function pullNow(
  reason: string,
  options: { mode?: 'replace' | 'union' } = {},
): Promise<{
  pulled: boolean
  reason?: string
  ruleCount?: number
  tombstoneCount?: number
  mergeMode?: 'replace' | 'union'
}> {
  return doPull(reason, options)
}

/**
 * Wait for any in-flight push / pull to drain. Used by disableSync
 * before clearing cloud, so an in-flight doPush (which already
 * computed its chunks and is about to call chrome.storage.sync.set)
 * can't race and re-write the keys we're trying to remove.
 *
 * Bug #N: previously disableSync called setSettings(syncEnabled:false)
 * then clearCloudState() without waiting, leaving a window where a
 * mid-flight push could repopulate cloud right after the wipe.
 *
 * Implementation: cancel the debounced timer (so future pushes don't
 * fire), then poll for the in-flight flags. 50 ms poll is a balance
 * between responsiveness and CPU. Max wait ~5 s — beyond that a push
 * is hung and we give up rather than block UI forever.
 */
export async function quiesce(): Promise<void> {
  if (pushTimer) {
    clearTimeout(pushTimer)
    pushTimer = null
  }
  const start = Date.now()
  // Wait on pullActive (writes happening), NOT pullInProgress (echo
  // suppression). The grace period doesn't matter for clearCloudState —
  // it only suppresses listener-triggered echoes, which clearCloud
  // doesn't care about.
  while (pushInFlight || pullActive) {
    if (Date.now() - start > QUIESCE_TIMEOUT_MS) {
      console.warn(
        '[mail-organizer] quiesce timed out — push or pull still in flight after 5s',
      )
      return
    }
    await new Promise((r) => setTimeout(r, 50))
  }
}

/**
 * Wipe everything we put in chrome.storage.sync. Called by
 * "停用並清除雲端資料". Doesn't touch local state.
 */
export async function clearCloudState(): Promise<void> {
  const all = await chrome.storage.sync.get(null)
  const ourKeys = Object.keys(all).filter(
    (k) =>
      k === SYNC_META_KEY ||
      k === SYNC_SETTINGS_KEY ||
      k.startsWith(SYNC_RULES_CHUNK_PREFIX) ||
      k.startsWith(SYNC_TOMBSTONES_CHUNK_PREFIX) ||
      k.startsWith(SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX),
  )
  if (ourKeys.length > 0) await chrome.storage.sync.remove(ourKeys)
}

export type SyncStatus = {
  enabled: boolean
  lastSyncAt: string
  machineId: string
  /** Cloud state summary; undefined when no cloud state present (never synced). */
  cloud?: {
    sourceMachineId: string
    updatedAt: string
    ruleCount: number
    tombstoneCount: number
    isUs: boolean
    /** All-machine push log (cap 20), newest first. Each entry shows
     *  which machine pushed when. Multi-machine visibility (#6). */
    recentPushes: SyncRecentPush[]
  }
  /** Approx bytes our keys use in chrome.storage.sync (vs 100 KB total quota). */
  bytesInUse: number
  /** Hard quota for chrome.storage.sync. */
  bytesQuota: number
  /**
   * Last sync failure, if any. Set when push / pull / remote-pull throws or
   * returns a refusal (quota, schema mismatch, network blip). Cleared on
   * the next successful push or pull. Surfaced in the options UI as a red
   * banner so silently-broken sync doesn't look "working".
   */
  lastError?: SyncErrorEntry
}

export async function getSyncStatus(): Promise<SyncStatus> {
  const settings = await getSettings()
  const meta = await readSyncMeta()
  const bytesInUse = await getOurBytesInUse()
  const lastError = await readSyncError()
  return {
    enabled: settings.syncEnabled,
    lastSyncAt: settings.lastSyncAt,
    machineId: settings.syncMachineId,
    cloud: meta
      ? {
          sourceMachineId: meta.sourceMachineId,
          updatedAt: meta.updatedAt,
          ruleCount: meta.ruleCount,
          tombstoneCount: meta.tombstoneCount,
          isUs: meta.sourceMachineId === settings.syncMachineId,
          // Bug #V: guard against malformed cloud payload.
          recentPushes: Array.isArray(meta.recentPushes) ? meta.recentPushes : [],
        }
      : undefined,
    bytesInUse,
    bytesQuota: chrome.storage.sync.QUOTA_BYTES ?? 102_400,
    lastError,
  }
}

// ---- backups ---------------------------------------------------------------

export async function listBackups(): Promise<SyncBackup[]> {
  const r = await chrome.storage.local.get(SYNC_BACKUPS_KEY)
  return (r[SYNC_BACKUPS_KEY] as SyncBackup[] | undefined) ?? []
}

async function appendBackup(direction: SyncBackup['direction']): Promise<void> {
  // Bug #Q: serialize the read-list / prepend / slice / write sequence
  // so concurrent appendBackup calls (e.g. doPush + doPull racing, or
  // restoreBackup landing during a push) can't clobber each other's
  // snapshot. The data captured (rules / tombstones / settings) is read
  // inside the lock too, so each backup reflects a consistent moment.
  return withBackupLock(async () => {
    const [rules, tombstones, settings] = await Promise.all([
      listRules(),
      getRuleTombstones(),
      getSettings(),
    ])
    const { claudeApiKey: _omit, ...settingsNoKey } = settings
    void _omit
    const snapshot: SyncBackup = {
      snapshotAt: new Date().toISOString(),
      direction,
      payload: { rules, tombstones, settings: settingsNoKey },
    }
    const existing = await listBackups()
    // Newest first, cap to MAX_BACKUPS.
    const next = [snapshot, ...existing].slice(0, MAX_BACKUPS)
    await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: next })
  })
}

/**
 * Restore a backup by its snapshotAt timestamp. Overwrites rules +
 * tombstones; settings overlay preserves the local API key. Does NOT
 * trigger a push (the user is explicitly rolling back local — they
 * may or may not want this to propagate yet).
 */
export async function restoreBackup(snapshotAt: string): Promise<{ restored: boolean }> {
  const backups = await listBackups()
  const backup = backups.find((b) => b.snapshotAt === snapshotAt)
  if (!backup) return { restored: false }
  // Save a "manual-rollback" snapshot of current state before
  // overwriting. Gives the user one more level of undo.
  await appendBackup('manual')
  // Pre-cap restored rules (older snapshots may carry uncapped confidence).
  await setRules(backup.payload.rules.map(applyConfidenceCap))
  // Tombstones: TRUE restore — clear current first, then write backup's.
  //   Bug #F: previously this was addRuleTombstones-only, which UNIONed
  //   the backup's tombstones with whatever was currently in local
  //   storage. Result: restoring a backup didn't actually revert
  //   tombstones, it just added the backup's on top. Now restore is a
  //   real reset: state after restore matches state at backup time
  //   (for tombstones; rules; settings minus per-device).
  await clearAllRuleTombstones()
  if (backup.payload.tombstones.length > 0) {
    await addRuleTombstones(backup.payload.tombstones)
  }
  // Settings overlay — strip per-device fields (machineId / lastSyncAt /
  // syncEnabled) so a backup taken pre-pull (which captured the cloud
  // device's id) doesn't change our identity on restore.
  const safeBackupSettings: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(backup.payload.settings)) {
    if (!PER_DEVICE_SETTINGS_FIELDS.has(k as keyof Settings)) {
      safeBackupSettings[k] = v
    }
  }
  await setSettings(safeBackupSettings as Partial<Settings>)
  return { restored: true }
}

// ---- internals -------------------------------------------------------------

async function ensureMachineId(): Promise<string> {
  const settings = await getSettings()
  if (settings.syncMachineId) return settings.syncMachineId
  const id =
    typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? (crypto as { randomUUID: () => string }).randomUUID()
      : Math.random().toString(36).slice(2) + Date.now().toString(36)
  await setSettings({ syncMachineId: id })
  return id
}

async function readSyncMeta(): Promise<SyncMeta | undefined> {
  const r = await chrome.storage.sync.get(SYNC_META_KEY)
  return r[SYNC_META_KEY] as SyncMeta | undefined
}

async function getOurBytesInUse(): Promise<number> {
  // chrome.storage.sync.getBytesInUse(null) returns bytes used by ALL
  // keys, including any from other extensions or earlier versions.
  // For our purposes, we list and filter to known prefixes.
  const all = await chrome.storage.sync.get(null)
  let bytes = 0
  for (const [key, value] of Object.entries(all)) {
    if (
      key !== SYNC_META_KEY &&
      key !== SYNC_SETTINGS_KEY &&
      !key.startsWith(SYNC_RULES_CHUNK_PREFIX) &&
      !key.startsWith(SYNC_TOMBSTONES_CHUNK_PREFIX) &&
      !key.startsWith(SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX)
    ) {
      continue
    }
    bytes += key.length + JSON.stringify(value).length
  }
  return bytes
}

async function doPush(
  reason: string,
  options: { wipeMarker?: boolean } = {},
): Promise<{
  pushed: boolean
  reason?: string
  /** Rules dropped because chunk count exceeded MAX_RULE_CHUNKS. */
  truncatedRuleCount?: number
  /** Tombstones dropped because chunk count exceeded MAX_TOMBSTONE_CHUNKS. */
  truncatedTombstoneCount?: number
}> {
  const settings = await getSettings()
  if (!settings.syncEnabled) return { pushed: false, reason: 'sync disabled' }
  if (pushInFlight) {
    // Coalesce — schedule another push after the current finishes.
    schedulePush('coalesce')
    return { pushed: false, reason: 'already in flight; rescheduled' }
  }
  pushInFlight = true
  try {
    // Downgrade protection (Bug #H): if cloud has a HIGHER schemaVersion
    // than us, refuse to push — overwriting would silently corrupt the
    // other machine's data with a lower-version payload. Only happens
    // when one machine is on an older extension version than another;
    // surface the version mismatch so the user knows to upgrade.
    const cloudMeta = await readSyncMeta()
    if (cloudMeta && cloudMeta.schemaVersion > SYNC_SCHEMA_VERSION) {
      const reason = `cloud schemaVersion=${cloudMeta.schemaVersion} > ours=${SYNC_SCHEMA_VERSION}; upgrade this browser's extension before pushing`
      await recordSyncError({ source: 'push', reason })
      return { pushed: false, reason }
    }
    const machineId = await ensureMachineId()
    const [rules, tombstones, folderActivity] = await Promise.all([
      listRules(),
      getRuleTombstones(),
      getFolderActivity(),
    ])
    await appendBackup('pre-push')

    const allRuleChunks = chunkRules(rules)
    const ruleChunks = allRuleChunks.slice(0, MAX_RULE_CHUNKS)
    const truncatedRuleChunks = allRuleChunks.length - ruleChunks.length
    const truncatedRuleCount = allRuleChunks
      .slice(MAX_RULE_CHUNKS)
      .reduce((sum, c) => sum + c.items.length, 0)

    const allTombChunks = chunkTombstones(tombstones)
    const tombChunks = allTombChunks.slice(0, MAX_TOMBSTONE_CHUNKS)
    const truncatedTombCount = allTombChunks
      .slice(MAX_TOMBSTONE_CHUNKS)
      .reduce((sum, c) => sum + c.items.length, 0)

    // chunkFolderActivity already truncates to FOLDER_ACTIVITY_SYNC_CAP=20.
    // The MAX_FOLDER_ACTIVITY_CHUNKS cap here is defense-in-depth against
    // an unexpected blow-up (e.g. one entry with a 5KB latestMessage subject).
    const allActivityChunks = chunkFolderActivity(folderActivity)
    const activityChunks = allActivityChunks.slice(0, MAX_FOLDER_ACTIVITY_CHUNKS)
    const truncatedActivityCount = allActivityChunks
      .slice(MAX_FOLDER_ACTIVITY_CHUNKS)
      .reduce((sum, c) => sum + c.items.length, 0)

    if (
      truncatedRuleChunks > 0 ||
      truncatedTombCount > 0 ||
      truncatedActivityCount > 0
    ) {
      console.warn(
        `[mail-organizer] sync push: truncated ${truncatedRuleCount} rules` +
          ` + ${truncatedTombCount} tombstones` +
          ` + ${truncatedActivityCount} folderActivity entries (chunk cap reached).` +
          ` Clean up auto-disabled / sleeping rules to reduce sync set.`,
      )
    }

    // Settings stripped of all per-device fields (apiKey, machineId,
    // lastSyncAt, syncEnabled) — see PER_DEVICE_SETTINGS_FIELDS for why.
    const settingsToSync = stripPerDeviceSettings(settings)

    const updatedAt = new Date().toISOString()
    // #6: extend the recentPushes log with this push's entry. Keep the
    // existing entries (from any machine that pushed before this), drop
    // any older-than-cap entries. cloudMeta might be undefined on first
    // push or missing the field if cloud was written by an older version
    // — default to [] in those cases.
    //
    // Bug #V: defend against malformed cloud payload. `?? []` only
    // catches null/undefined; if a buggy client wrote `recentPushes:
    // "string"` or `{}`, `.filter` would crash and break the entire
    // push. Array.isArray is the only safe guard.
    const prevRecent = Array.isArray(cloudMeta?.recentPushes)
      ? (cloudMeta!.recentPushes as SyncRecentPush[])
      : []
    const recentPushes: SyncRecentPush[] = [
      { machineId, at: updatedAt },
      ...prevRecent.filter(
        // Deduplicate: drop ANY older entry from the same machineId so
        // each machine has at most one entry, the latest. Without this,
        // a single chatty machine would dominate the log and we'd never
        // see entries from less-active machines. Also defend against
        // entries with wrong shape via the optional-chain.
        (p) => p && typeof p === 'object' && p.machineId !== machineId,
      ),
    ].slice(0, SYNC_RECENT_PUSHES_CAP)
    // wipeMarker handling:
    //   - When THIS push is a wipe (options.wipeMarker = true): stamp a
    //     new marker keyed to our machineId + now. Replaces any prior
    //     marker; only the latest wipe is honoured downstream.
    //   - When THIS push is normal: preserve cloudMeta.wipeMarker so the
    //     marker keeps living until the next wipe. Without this preservation,
    //     a normal push from any machine right after a wipe would erase
    //     the marker before the OTHER machine ever had a chance to pull
    //     and trip it.
    const carriedWipeMarker = options.wipeMarker
      ? { at: updatedAt, byMachineId: machineId }
      : (cloudMeta?.wipeMarker &&
          typeof cloudMeta.wipeMarker === 'object' &&
          typeof cloudMeta.wipeMarker.at === 'string' &&
          typeof cloudMeta.wipeMarker.byMachineId === 'string')
        ? cloudMeta.wipeMarker
        : undefined

    const meta: SyncMeta = {
      schemaVersion: SYNC_SCHEMA_VERSION,
      sourceMachineId: machineId,
      updatedAt,
      ruleCount: ruleChunks.reduce((sum, c) => sum + c.items.length, 0),
      tombstoneCount: tombChunks.reduce((sum, c) => sum + c.items.length, 0),
      ruleChunkCount: ruleChunks.length,
      tombstoneChunkCount: tombChunks.length,
      folderActivityCount: activityChunks.reduce(
        (sum, c) => sum + c.items.length,
        0,
      ),
      folderActivityChunkCount: activityChunks.length,
      recentPushes,
      ...(carriedWipeMarker ? { wipeMarker: carriedWipeMarker } : {}),
    }

    // Build the write payload. First clear any orphan chunks from a
    // previous push that produced more chunks than this one — otherwise
    // a shrinking library would leave stale chunks behind that get
    // included on the next pull.
    const allKeys = Object.keys(await chrome.storage.sync.get(null))
    const orphanRuleKeys = allKeys.filter(
      (k) =>
        k.startsWith(SYNC_RULES_CHUNK_PREFIX) &&
        chunkIndex(k, SYNC_RULES_CHUNK_PREFIX) >= ruleChunks.length,
    )
    const orphanTombKeys = allKeys.filter(
      (k) =>
        k.startsWith(SYNC_TOMBSTONES_CHUNK_PREFIX) &&
        chunkIndex(k, SYNC_TOMBSTONES_CHUNK_PREFIX) >= tombChunks.length,
    )
    const orphanActivityKeys = allKeys.filter(
      (k) =>
        k.startsWith(SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX) &&
        chunkIndex(k, SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX) >= activityChunks.length,
    )
    const writes: Record<string, unknown> = {
      [SYNC_META_KEY]: meta,
      [SYNC_SETTINGS_KEY]: settingsToSync,
    }
    for (const c of ruleChunks) writes[`${SYNC_RULES_CHUNK_PREFIX}${c.index}`] = c
    for (const c of tombChunks) writes[`${SYNC_TOMBSTONES_CHUNK_PREFIX}${c.index}`] = c
    for (const c of activityChunks)
      writes[`${SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX}${c.index}`] = c

    try {
      // Audit G3: write the new meta + chunks FIRST, then remove orphans.
      // The new meta carries a SMALLER chunk count when the library shrank;
      // a puller only ever reads chunk indices < that count, so the orphan
      // chunks (old indices at/after the new count) are harmless dead data
      // once the new meta lands. Removing them BEFORE the set was a silent-
      // data-loss hazard: if this set failed (quota/network) or the SW died
      // in between, the OLD meta (larger count) survived in cloud while the
      // chunks it referenced were already deleted, so a replace-mode pull on
      // another machine read undefined chunks and dropped those rules with
      // no tombstone. Set-first means a partial failure leaves a fully valid
      // cloud (just some extra dead chunks).
      await chrome.storage.sync.set(writes)
    } catch (e) {
      const errMsg = e instanceof Error ? e.message : String(e)
      // Quota-wedge recovery (audit H1 — a failure branch the G3 reorder
      // opened): the sync quota is enforced on the post-set bucket TOTAL,
      // i.e. new payload + still-stored orphan chunks. Near quota, a
      // shrinking library can make that union exceed the limit forever —
      // and the orphan cleanup that would free the space sat AFTER this
      // set, unreachable. Pruning more rules doesn't help (payload shrinks,
      // orphan set grows one-for-one), so the user was permanently wedged
      // until 停用並清除雲端資料. Recovery: ONLY on a quota error, remove
      // the orphans first and retry the set once. This briefly reopens the
      // pre-G3 window (old meta referencing removed chunks) — acceptable
      // because (a) it triggers only when the push is already failing
      // permanently, and (b) remote pulls are union-mode + tombstone-aware
      // (Bug #E), so missing chunks no longer silently drop rules.
      const orphanKeys = [...orphanRuleKeys, ...orphanTombKeys, ...orphanActivityKeys]
      const isQuota = /quota/i.test(errMsg)
      if (isQuota && orphanKeys.length > 0) {
        try {
          console.warn(
            `[mail-organizer] sync push (${reason}): quota exceeded with ` +
              `${orphanKeys.length} orphan chunk(s) still stored — removing orphans and retrying once`,
          )
          await chrome.storage.sync.remove(orphanKeys)
          await chrome.storage.sync.set(writes)
          // Recovered — fall through to the normal post-set path (the
          // orphan cleanup below re-removes already-gone keys, a no-op).
        } catch (e2) {
          const retryMsg = e2 instanceof Error ? e2.message : String(e2)
          console.warn(
            `[mail-organizer] sync push (${reason}): retry after orphan removal also failed`,
            e2,
          )
          await recordSyncError({ source: 'push', reason: retryMsg })
          await logError('sync:push', retryMsg, { reason })
          return { pushed: false, reason: retryMsg }
        }
      } else {
        // Non-quota errors (network etc.) or nothing to free. Surface in
        // console + persistent error so the options UI can show "sync
        // broken" (Bug #J). Don't update lastSyncAt so a retry remains
        // possible.
        console.warn(`[mail-organizer] sync push (${reason}): chrome.storage.sync.set failed`, e)
        await recordSyncError({ source: 'push', reason: errMsg })
        await logError('sync:push', errMsg, { reason })
        return { pushed: false, reason: errMsg }
      }
    }

    // Orphan cleanup AFTER the authoritative write — non-fatal. If it fails
    // or the SW dies here, cloud stays valid (extra unreferenced chunks are
    // never read); the next successful push recomputes and removes them.
    if (
      orphanRuleKeys.length > 0 ||
      orphanTombKeys.length > 0 ||
      orphanActivityKeys.length > 0
    ) {
      await chrome.storage.sync
        .remove([...orphanRuleKeys, ...orphanTombKeys, ...orphanActivityKeys])
        .catch((e) =>
          console.warn('[mail-organizer] orphan chunk cleanup failed (non-fatal)', e),
        )
    }

    await setSettings({ lastSyncAt: meta.updatedAt })
    // Push succeeded — clear any prior error so the UI banner goes away.
    await clearSyncError()
    return {
      pushed: true,
      truncatedRuleCount: truncatedRuleCount > 0 ? truncatedRuleCount : undefined,
      truncatedTombstoneCount:
        truncatedTombCount > 0 ? truncatedTombCount : undefined,
    }
  } finally {
    pushInFlight = false
  }
}

function chunkIndex(key: string, prefix: string): number {
  const n = Number.parseInt(key.slice(prefix.length), 10)
  return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER
}

/**
 * Pull merge mode:
 *   - 'replace': delete local syncable rules that are missing from
 *     cloud (assumed deleted upstream). Used ONLY for explicit
 *     user-clicked manual pulls where the intent is "make my local
 *     state match cloud exactly". Risky outside that context —
 *     between two pushes any rule the user just added locally would
 *     get dropped.
 *   - 'union': keep local rules that aren't in cloud. Apply cloud
 *     tombstones to detect upstream deletions and drop those
 *     specifically (Bug #E). This is the default for remote-triggered
 *     pulls AND first-enable — both cases where local-only rules
 *     might legitimately exist that the cloud machine never saw.
 */
type PullMergeMode = 'replace' | 'union'

async function doPull(
  _reason: string,
  options: { mode?: PullMergeMode } = {},
): Promise<{
  pulled: boolean
  reason?: string
  ruleCount?: number
  tombstoneCount?: number
  /** Whether local-only rules were preserved during merge (union mode). */
  mergeMode?: PullMergeMode
}> {
  // Bug #O: concurrent doPull guard. Set pullActive=true SYNCHRONOUSLY
  // before any await — otherwise two doPull calls reach the `if
  // (pullActive)` check before either assignment, both see false, both
  // proceed (classic TOCTOU). The check + set must be atomic from JS's
  // single-threaded perspective: no awaits between them.
  if (pullActive) {
    return { pulled: false, reason: 'pull already in flight' }
  }
  pullActive = true
  pullInProgress = true
  if (pullGraceTimer) {
    clearTimeout(pullGraceTimer)
    pullGraceTimer = null
  }
  try {
  const settings = await getSettings()
  if (!settings.syncEnabled) {
    // Caller flipped sync off between our entry and the awaited read.
    // Bail cleanly — the flags get released by finally.
    return { pulled: false, reason: 'sync disabled' }
  }
  const meta = await readSyncMeta()
  if (!meta) return { pulled: false, reason: 'no cloud state' }
  if (meta.schemaVersion !== SYNC_SCHEMA_VERSION) {
    return {
      pulled: false,
      reason: `unsupported cloud schemaVersion=${meta.schemaVersion} (we support ${SYNC_SCHEMA_VERSION})`,
    }
  }
  // Default merge mode: replace (the steady-state assumption). First-
  // enable explicitly passes 'union' via the enable handler.
  const mode = options.mode ?? 'replace'

  // Cross-machine wipe detection. If cloud carries a wipeMarker from
  // another machine that's newer than our last sync, this machine
  // needs to drop its syncable rules + tombstones BEFORE the normal
  // merge. Without this, union-mode pull would preserve our local-only
  // syncable rules (which the issuing machine intentionally wiped),
  // defeating "全部刪除" across machines.
  //
  // Echo guard: a machine never trips its own wipeMarker
  // (byMachineId === ours).
  // Staleness guard: don't re-trip a marker we already applied
  // (marker.at <= our lastSyncAt — we already advanced past it).
  // First-participation guard (F1, 2026-06-03): a machine that has
  // NEVER synced (lastSyncAt === '') must NOT be wiped by a marker
  // that predates its participation. Otherwise enabling sync on a
  // machine with a locally-built rule library — whose first pull is
  // union-mode specifically to PRESERVE those local rules — would trip
  // an old wipe left in cloud by another machine and silently nuke the
  // library at the exact moment the user opted in. The marker is a
  // signal to machines ALREADY in the sync set, not to newcomers.
  // Both checked here BEFORE the mutation runs.
  let remoteWipeApplied:
    | { byMachineId: string; at: string }
    | undefined
  if (
    meta.wipeMarker &&
    typeof meta.wipeMarker === 'object' &&
    typeof meta.wipeMarker.at === 'string' &&
    typeof meta.wipeMarker.byMachineId === 'string' &&
    meta.wipeMarker.byMachineId !== settings.syncMachineId &&
    settings.lastSyncAt !== '' && // F1: never wipe a first-time participant
    meta.wipeMarker.at > settings.lastSyncAt
  ) {
    remoteWipeApplied = {
      byMachineId: meta.wipeMarker.byMachineId,
      at: meta.wipeMarker.at,
    }
  }

  await appendBackup('pre-pull')

  // Apply the remote wipe (if any) BEFORE merging cloud chunks. We
  // keep this inside the pullActive critical section so the local
  // mutation listener treats the writes as pull-internal (no echo
  // push). The pre-pull backup we just took preserves the user's
  // chance to roll back via the syncBackups list.
  let droppedRuleCount = 0
  if (remoteWipeApplied) {
    await mutateRules((existing) => {
      const next = existing.filter((r) => !shouldSyncRule(r))
      droppedRuleCount = existing.length - next.length
      return { next, result: undefined }
    })
    // F3 (2026-06-03): do NOT clear local tombstones on a REMOTE wipe
    // apply. Tombstones are this machine's own deletion intents; the
    // user didn't click "全部刪除" here, so wiping them would let
    // rules this machine's user specifically deleted resurrect later
    // (e.g. a third machine pushes the same triple back). The LOCAL
    // wipeAllRules handler still clears tombstones — there the user
    // explicitly chose a fresh start on this device. Keeping tombstones
    // is strictly safer (they're capped and only block resurrection).
    // F2 (2026-06-03): advance lastSyncAt to the marker time IMMEDIATELY
    // after the destructive wipe commits, BEFORE the chunk reads / merge
    // below (which can throw on network/quota). Without this, a throw
    // anywhere downstream leaves the wipe applied but lastSyncAt
    // un-advanced, so the marker stays "fresh" and re-trips on the next
    // remote push — repeatedly clobbering any rules re-created in between.
    try {
      await setSettings({ lastSyncAt: remoteWipeApplied.at })
    } catch (e) {
      console.warn(
        '[mail-organizer] advancing lastSyncAt after wipe failed (non-fatal)',
        e,
      )
    }
    try {
      const notice: RemoteWipeNotice = {
        byMachineId: remoteWipeApplied.byMachineId,
        at: remoteWipeApplied.at,
        appliedAt: new Date().toISOString(),
        droppedRuleCount,
      }
      await chrome.storage.local.set({ [REMOTE_WIPE_NOTICE_KEY]: notice })
    } catch (e) {
      // Non-fatal — the wipe itself is done; only the UI banner notice
      // is missing. Log so the user can still find evidence of what
      // happened.
      console.warn(
        '[mail-organizer] persisting remoteWipeNotice failed (non-fatal)',
        e,
      )
    }
    console.warn(
      `[mail-organizer] remote wipe applied: ${droppedRuleCount} syncable rules dropped (issued by machine ${remoteWipeApplied.byMachineId} at ${remoteWipeApplied.at})`,
    )
  }

  // Read all chunks. Plus the settings overlay.
  const ruleKeys: string[] = []
  for (let i = 0; i < meta.ruleChunkCount; i++) ruleKeys.push(`${SYNC_RULES_CHUNK_PREFIX}${i}`)
  const tombKeys: string[] = []
  for (let i = 0; i < meta.tombstoneChunkCount; i++)
    tombKeys.push(`${SYNC_TOMBSTONES_CHUNK_PREFIX}${i}`)
  // folderActivity chunks: v2 schema. v1 cloud has 0 of these (the field
  // is optional on SyncMeta), so the for-loop is a no-op and we read no
  // activity from cloud — local stays untouched.
  const activityKeys: string[] = []
  const activityChunkCount = meta.folderActivityChunkCount ?? 0
  for (let i = 0; i < activityChunkCount; i++)
    activityKeys.push(`${SYNC_FOLDER_ACTIVITY_CHUNK_PREFIX}${i}`)

  const stored = await chrome.storage.sync.get([
    SYNC_SETTINGS_KEY,
    ...ruleKeys,
    ...tombKeys,
    ...activityKeys,
  ])

  const ruleChunks = ruleKeys
    .map((k) => stored[k] as SyncChunk<Rule> | undefined)
    .filter((c): c is SyncChunk<Rule> => !!c)
  const tombChunks = tombKeys
    .map((k) => stored[k] as SyncChunk<RuleTombstone> | undefined)
    .filter((c): c is SyncChunk<RuleTombstone> => !!c)
  const activityChunksFromCloud = activityKeys
    .map((k) => stored[k] as SyncChunk<FolderActivity> | undefined)
    .filter((c): c is SyncChunk<FolderActivity> => !!c)

  const cloudRules = unchunkArray(ruleChunks)
  const cloudTombstones = unchunkArray(tombChunks)
  const cloudFolderActivity = unchunkArray(activityChunksFromCloud)

  // Build the tombstone key set ONCE before the mutex — checking each
  // local rule against this catches upstream deletions even in union
  // mode (Bug #E: without this, deleted-on-A rules would survive on B
  // because B's pull keeps them, defeating the deletion).
  const cloudTombstoneKeys = new Set(
    cloudTombstones.map(
      (t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`,
    ),
  )

  // Merge into local rules.
  //   For each existing local rule:
  //     - same id in cloud → take cloud version (Q4: full sync incl stats)
  //     - missing from cloud:
  //         - per-device (auto_scan / orphaned) → keep
  //         - syncable + tombstoned in cloud → drop (deleted upstream,
  //           regardless of mode — tombstone trumps mode)
  //         - syncable + mode='replace' → drop (caller wants exact match
  //           with cloud)
  //         - syncable + mode='union' → KEEP (no tombstone = local-only
  //           addition we don't want to silently lose)
  //   Cloud rules not in local → always add.
  await mutateRules((existing) => {
    // Pre-cap incoming cloud rules. Cross-machine sync used to be a
    // bypass for applyConfidenceCap — a machine on an old build could
    // push plain-domain rules at conf=1.0 that landed on every other
    // machine unchanged, inverting the per-type priority the new
    // design relies on.
    const cappedCloudRules = cloudRules.map(applyConfidenceCap)
    const cloudById = new Map(cappedCloudRules.map((r) => [r.id, r]))
    const next: Rule[] = []
    for (const r of existing) {
      const fromCloud = cloudById.get(r.id)
      if (fromCloud) {
        next.push(fromCloud)
        cloudById.delete(r.id)
      } else if (!shouldSyncRule(r)) {
        // Per-device rule (auto_scan / orphaned) — never touched by sync.
        next.push(r)
      } else {
        // Syncable rule, not in cloud. Check upstream deletion first.
        const key = `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
        if (cloudTombstoneKeys.has(key)) {
          // Tombstoned upstream — drop in EITHER mode. This is how
          // deletes propagate across machines under union semantics.
        } else if (mode === 'union') {
          // Local-only rule, no tombstone → user added on this machine
          // between the last push from cloud and now. Preserve it.
          next.push(r)
        }
        // mode === 'replace' + syncable + missing from cloud + no tombstone
        // → drop (caller explicitly wants cloud-exact state).
      }
    }
    for (const r of cloudById.values()) next.push(r)
    return { next, result: undefined }
  })

  // Tombstones: union (addRuleTombstones is idempotent by triple).
  if (cloudTombstones.length > 0) {
    await addRuleTombstones(cloudTombstones)
  }

  // folderActivity: union by folderId, most-recent lastActiveAt wins.
  // Additive merge — never drops local entries cloud doesn't have, since
  // each device legitimately has its own per-machine activity history.
  if (cloudFolderActivity.length > 0) {
    await mergeFolderActivityFromSync(cloudFolderActivity)
  }

  // Settings overlay. Strip per-device fields from incoming (defence-in-
  // depth — push side already strips them, but a future schema change
  // might leak them through). Local apiKey + machineId etc. stay put.
  const incomingSettings = stored[SYNC_SETTINGS_KEY] as Partial<Settings> | undefined
  if (incomingSettings && typeof incomingSettings === 'object') {
    const safeIncoming: Record<string, unknown> = {}
    for (const [k, v] of Object.entries(incomingSettings)) {
      if (!PER_DEVICE_SETTINGS_FIELDS.has(k as keyof Settings)) {
        safeIncoming[k] = v
      }
    }
    await setSettings(safeIncoming as Partial<Settings>)
  }

  await setSettings({ lastSyncAt: meta.updatedAt })
  // Pull succeeded — clear any prior error so the UI banner goes away.
  await clearSyncError()
  return {
    pulled: true,
    ruleCount: cloudRules.length,
    tombstoneCount: cloudTombstones.length,
    mergeMode: mode,
  }
  } finally {
    // Writes are done — release the re-entry guard immediately so the
    // next genuine pull (manual, debounced, or first-enable) isn't
    // blocked. Echo suppression continues via pullInProgress for the
    // grace window.
    pullActive = false
    // Bug #G: don't clear pullInProgress synchronously. Chrome's
    // chrome.storage.onChanged listeners for the writes inside this try
    // block are dispatched as MACROTASKS — they fire AFTER this finally
    // runs. Synchronous clear would let those listeners see the flag as
    // false and schedule an echo push. Delay by PULL_GRACE_MS to outlive
    // the listener queue.
    pullGraceTimer = setTimeout(() => {
      pullInProgress = false
      pullGraceTimer = null
    }, PULL_GRACE_MS)
  }
}

function handleRemoteSyncMetaChange(newMeta: SyncMeta | undefined): void {
  if (!newMeta) return
  // P-2: debounce 200 ms so a burst of remote-meta changes from the
  // other machine (e.g. rule edits N times in 500 ms) collapses into
  // ONE pull, not N sequential pulls. Each pull does a full backup +
  // read + merge so back-to-back was meaningful waste.
  //
  // Replace any pending pull with the latest meta — the freshest is
  // always the right thing to chase (intermediate snapshots are
  // subsumed by it).
  pendingPullMeta = newMeta
  if (pullDebounceTimer) clearTimeout(pullDebounceTimer)
  pullDebounceTimer = setTimeout(() => {
    pullDebounceTimer = null
    const meta = pendingPullMeta
    pendingPullMeta = undefined
    if (!meta) return
    void dispatchRemotePull(meta)
  }, PULL_DEBOUNCE_MS)
}

async function dispatchRemotePull(newMeta: SyncMeta): Promise<void> {
  const settings = await getSettings()
  if (!settings.syncEnabled) return
  // Echo of our own push? Skip.
  if (newMeta.sourceMachineId === settings.syncMachineId) return
  // Pull. Bug #E: use 'union' mode (with tombstone-aware drop), not
  // 'replace' — between two pushes the user may have added local rules
  // that the remote machine never saw. 'replace' would silently drop
  // them. 'union' preserves them; tombstones still propagate deletions.
  //
  // Bug #J: errors here previously only console.warn'd. Now we also
  // persist them so the options UI can surface "sync silently broken".
  try {
    await doPull('remote-change', { mode: 'union' })
  } catch (e) {
    const errMsg = e instanceof Error ? e.message : String(e)
    console.warn('[mail-organizer] sync pull on remote-change failed', e)
    await recordSyncError({ source: 'pull-remote', reason: errMsg }).catch(
      () => {
        /* avoid feedback loop if storage itself is broken */
      },
    )
    await logError('sync:pull-remote', errMsg, {
      sourceMachineId: newMeta.sourceMachineId,
    }).catch(() => {})
  }
}

// ---- error tracking (Bug #J) ---------------------------------------------

async function recordSyncError(
  entry: Omit<SyncErrorEntry, 'at'>,
): Promise<void> {
  const payload: SyncErrorEntry = { ...entry, at: new Date().toISOString() }
  await chrome.storage.local.set({ [SYNC_LAST_ERROR_KEY]: payload })
}

async function clearSyncError(): Promise<void> {
  await chrome.storage.local.remove(SYNC_LAST_ERROR_KEY)
}

async function readSyncError(): Promise<SyncErrorEntry | undefined> {
  const r = await chrome.storage.local.get(SYNC_LAST_ERROR_KEY)
  return r[SYNC_LAST_ERROR_KEY] as SyncErrorEntry | undefined
}

/** Manually dismiss the persistent sync error banner. */
export async function dismissSyncError(): Promise<void> {
  await clearSyncError()
}

// ---- remote wipe notice ---------------------------------------------------

/** Read the most recent remote-wipe-applied notice, if any. */
export async function readRemoteWipeNotice(): Promise<RemoteWipeNotice | undefined> {
  const r = await chrome.storage.local.get(REMOTE_WIPE_NOTICE_KEY)
  return r[REMOTE_WIPE_NOTICE_KEY] as RemoteWipeNotice | undefined
}

/** Dismiss the remote-wipe banner. */
export async function dismissRemoteWipeNotice(): Promise<void> {
  await chrome.storage.local.remove(REMOTE_WIPE_NOTICE_KEY)
}
