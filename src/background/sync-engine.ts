// Cross-machine sync engine via chrome.storage.sync.
//
// What gets synced:
//   - Rules where source ∈ {user_manual, ai_confirmed, ai_overridden}
//     and !orphaned. (auto_scan rules are cheap to re-derive per
//     machine via initial scan.) Disabled rules whose full row doesn't
//     fit the byte budget still travel as a compact stub (v4).
//   - Recent tombstones, newest first, inside a bounded byte share.
//   - Settings minus PER_DEVICE_SETTINGS_FIELDS (claudeApiKey never leaves
//     the device).
//
// What stays per-device:
//   - claudeApiKey (security)
//   - folderCache (different Outlook IDs even if same mailbox)
//   - skipHistory (per-device preview choices)
//   - conversationMemory / subjectMemory (per-device threading)
//
// Conflict semantics (v4, 2026-11-05): every push is the full state of one
// machine, but a pull MERGES it per rule id — semantic fields from the copy
// with the newer `editedAt` (tie → cloud), counters max, createdAt min — so a
// stale copy can no longer revert a newer edit. Tombstones suppress a rule
// only if the rule was created/edited BEFORE the deletion. Push and pull run
// through ONE queue, every push pulls first when the cloud holds a state this
// machine hasn't merged, and a pull accepts only a snapshot whose chunks all
// carry the meta's pushId (torn reads abort and retry). Destructive merges
// (replace, remote wipe, first-enable, the first v4 merge of an upgraded
// machine) and restores snapshot local state into `syncBackups` first.
//
// 2026-11-06 additions:
//   - adoption: a machine's FIRST v4 merge keeps (and re-stamps) live rules
//     that v3 kept next to an older tombstone (adoptionProtects), after an
//     'upgrade' backup.
//   - `orphaned` is per-device: never stamps editedAt, never wins a merge.
//   - counters merge by epoch (Rule.statsResetAt); counter-only changes are
//     pushed at most once per SYNC_COUNTER_PUSH_COALESCE_MS (meta.semanticHash).
//   - a cloud left torn under one pushId is repaired by its owner, or after
//     repeated torn reads rewritten from local (after a backup).
//   - the stale-sweep active-day ledger rides in the settings item (`d`).
//   - a synced setting this build can't represent (an unknown claudeModel) is
//     neither adopted nor bounced back (foreignSettings).
//
// Detection: every push writes `syncMeta.sourceMachineId = ourId` and a fresh
// `pushId`. The chrome.storage.onChanged listener ignores our own echoes and
// pulls otherwise.

import {
  addRuleTombstones,
  clearMatchingTombstones,
  getActiveDays,
  getFolderActivity,
  getRuleTombstones,
  getSettings,
  mergeActiveDaysFromSync,
  mergeFolderActivityFromSync,
  setSettings,
} from '@/shared/storage'
import {
  applyConfidenceCap,
  legacySignalNorm,
  listRules,
  mutateRules,
  normalizeSignal,
  ruleTombstone,
  semanticFieldsEqual,
  tombstoneKey,
} from '@/shared/rules'
import {
  chunkFolderActivity,
  jsonByteSize,
  planSyncPayload,
  shouldSyncRule,
  unchunkArray,
  utf8ByteLength,
  writeBytes,
  type SyncChunk,
  type SyncPayloadPlan,
} from '@/shared/sync-chunks'
import {
  contentHash,
  decodeActiveDays,
  decodeRulePayload,
  decodeTombstone,
  encodeActiveDays,
  sanitizeLegacyRule,
  sanitizeLegacyTombstone,
  spellingFingerprint,
  type CompactRuleRow,
  type CompactTombstone,
  type DecodedStub,
  type WireRuleEntry,
} from '@/shared/sync-codec'
import { isKnownClaudeModel } from '@/shared/claude-models'
import {
  ACTIVE_DAYS_CAP,
  CANONICAL_TOMBSTONE_KEYS_SINCE_MS,
  CHUNK_BYTE_TARGET,
  MAX_BACKUPS,
  MAX_FOLDER_ACTIVITY_CHUNKS,
  PULL_DEBOUNCE_MS,
  PULL_GRACE_MS,
  PUSH_DEBOUNCE_MS,
  QUIESCE_TIMEOUT_MS,
  SYNC_COUNTER_PUSH_COALESCE_MS,
  SYNC_MIN_READABLE_SCHEMA_VERSION,
  SYNC_PUSH_RETRY_BASE_MS,
  SYNC_PUSH_RETRY_MAX_MS,
  SYNC_RATE_LIMIT_HOUR_RETRY_MS,
  SYNC_RATE_LIMIT_MINUTE_RETRY_MS,
  SYNC_RECENT_PUSHES_CAP,
  SYNC_RECENT_TOMBSTONE_BUDGET_SHARE,
  SYNC_SCHEMA_VERSION,
  SYNC_STARTUP_RECONCILE_DELAY_MS,
  SYNC_TOMBSTONE_BUDGET_SHARE,
  SYNC_TORN_RECOVERY_MIN_MS,
  SYNC_TORN_RECOVERY_MIN_READS,
  SYNC_TORN_RETRY_BASE_MS,
  SYNC_TORN_RETRY_MAX_ATTEMPTS,
  SYNC_TORN_RETRY_MAX_MS,
  SYNC_TOTAL_BUDGET_BYTES,
} from '@/shared/constants'
import { logError } from '@/shared/error-log'
import type { FolderActivity, Rule, RuleSource, RuleTombstone, Settings } from '@/shared/types'

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
  // Harvesting reads THIS machine's Outlook manual-filing history — a
  // privacy-sensitive opt-in that each device must grant independently. Its
  // watermark (lastHarvestAt) is already chrome.storage.local, so keep the
  // toggle local too; a cloud pull must never silently enable it elsewhere.
  'harvestEnabled',
])

type PerDeviceFieldKey =
  | 'claudeApiKey'
  | 'syncMachineId'
  | 'lastSyncAt'
  | 'syncEnabled'
  | 'onboardingDismissed'
  | 'harvestEnabled'

/** Project a Settings record down to only the fields safe to sync. */
function stripPerDeviceSettings(s: Partial<Settings>): Omit<Settings, PerDeviceFieldKey> {
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
// v1–v3 layouts. A v4 pull still READS them (union-merged, then the next push
// rewrites the cloud as v4 and removes them) — schema skew used to refuse any
// non-v3 cloud, so enabling sync against a v2 cloud failed and told THIS
// machine to upgrade (2026-11-05).
const LEGACY_SETTINGS_KEY = 'syncSettings'
const LEGACY_RULES_PREFIX = 'syncRules_' // v1/v2: full Rule objects
const V3_ROWS_PREFIX = 'syncRulesC_' // v3: compact rows
const V3_FOLDERS_PREFIX = 'syncFolders_' // v3: folder dictionary
const LEGACY_TOMBS_PREFIX = 'syncTombstones_' // v1–v3: full tombstone objects
const LEGACY_ACTIVITY_PREFIX = 'syncFolderActivity_' // v2/v3
// v4 (2026-11-05). New key names on purpose: a v3 reader that slipped past
// its own version check must never mis-decode v4 rows (compact compounds
// would decode to empty signals there).
const V4_SETTINGS_KEY = 'syncSettings4'
const V4_ROWS_PREFIX = 'syncRows4_'
const V4_FOLDERS_PREFIX = 'syncFolders4_'
const V4_TOMBS_PREFIX = 'syncTombs4_'
const V4_ACTIVITY_PREFIX = 'syncActivity4_'

const CHUNK_PREFIXES = [
  LEGACY_RULES_PREFIX,
  V3_ROWS_PREFIX,
  V3_FOLDERS_PREFIX,
  LEGACY_TOMBS_PREFIX,
  LEGACY_ACTIVITY_PREFIX,
  V4_ROWS_PREFIX,
  V4_FOLDERS_PREFIX,
  V4_TOMBS_PREFIX,
  V4_ACTIVITY_PREFIX,
] as const

function isOurChunkKey(k: string): boolean {
  return CHUNK_PREFIXES.some((p) => k.startsWith(p))
}

function isOurSyncKey(k: string): boolean {
  return (
    k === SYNC_META_KEY || k === LEGACY_SETTINGS_KEY || k === V4_SETTINGS_KEY || isOurChunkKey(k)
  )
}

function rangeKeys(prefix: string, count: unknown): string[] {
  const n = typeof count === 'number' && Number.isInteger(count) && count > 0 ? Math.min(count, 64) : 0
  const out: string[] = []
  for (let i = 0; i < n; i++) out.push(`${prefix}${i}`)
  return out
}

// MAX_RULE_CHUNKS / MAX_TOMBSTONE_CHUNKS / MAX_FOLDER_ACTIVITY_CHUNKS:
// see src/shared/constants.ts. Chunk-count caps prevent the key namespace
// from blowing up when a single push has many chunks.

// ---- local-only meta -------------------------------------------------------

export const SYNC_BACKUPS_KEY = 'syncBackups'
// MAX_BACKUPS: see src/shared/constants.ts. Rotating slots (+ one pinned
// pre-wipe slot, never rotated out).

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

/**
 * Per-device sync bookkeeping (2026-11-05), chrome.storage.local only —
 * never synced, never in backups / exports. Only touched inside the sync
 * queue.
 */
export const SYNC_LOCAL_STATE_KEY = 'syncLocalState'

type SyncLocalState = {
  /**
   * Has this machine completed a pull or push since sync was (re-)enabled?
   * A non-participant records a cloud wipe marker as seen WITHOUT applying
   * it (F1: enabling sync must never wipe a locally-built library).
   */
  participating: boolean
  /** Id of the newest wipe marker this machine has acted on (applied or recorded). */
  lastSeenWipeId?: string
  /** pushId of the cloud state this machine last merged or wrote. */
  lastMergedPushId?: string
  /**
   * The synced settings as of the last successful pull / push — the common
   * base for a three-way settings merge. Without it a pull (now also run
   * before every push) overlaid the cloud's settings onto a local edit that
   * was still waiting to be pushed, silently reverting it.
   */
  settingsBase?: Record<string, unknown>
  /**
   * The first-participation merge ("adoption", 2026-11-06) has run since sync
   * was (re-)enabled — see adoptionProtects. Absent on states written by
   * v1.0.8: a participating one with lastMergedPushId counts as adopted.
   */
  v4Adopted?: boolean
  /**
   * lastSyncAt of a machine that was syncing under v1.0.7 when this build
   * first ran, kept until its first v4 merge: decides the transition's wipe
   * marker and which tombstones v3 had already processed.
   */
  legacyLastSyncAt?: string
  /**
   * Synced settings values this build cannot represent (A13: a claudeModel id
   * from a newer peer). `cloud` is pushed back unchanged while the local value
   * is still `local` — neither adopted (sanitize would map it to a default)
   * nor overwritten with ours.
   */
  foreignSettings?: Record<string, { cloud: unknown; local: unknown }>
  /** Counter-only changes are waiting for a coalesced push due at this time (ms). */
  counterPushDueAt?: number
}

/** Persisted torn-read streak (per-device): same meta pushId seen torn `count` times since `firstAt`. */
const SYNC_TORN_STREAK_KEY = 'syncTornStreak'
type TornStreak = { pushId: string; count: number; firstAt: number }

const COUNTER_PUSH_ALARM = 'mail-organizer:sync-counter-push'

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
  /**
   * Why the snapshot was taken (2026-11-05; absent on older backups).
   * 'upgrade' (2026-11-06): before an upgraded v1.0.7 machine's first v4
   * merge. 'torn-recovery': before a push rewrote a cloud left torn.
   */
  trigger?: 'replace-pull' | 'remote-wipe' | 'first-enable' | 'restore' | 'upgrade' | 'torn-recovery'
  /**
   * Exempt from rotation — one slot per trigger: the pre-wipe snapshot the
   * remote-wipe banner points to, and the pre-upgrade snapshot (2026-11-06).
   * The old 5-slot rotation, refilled by every push and pull, evicted the
   * pre-wipe one within minutes.
   */
  pinned?: boolean
  /** Local snapshot of what would be overwritten — minus per-machine state. */
  payload: {
    rules: Rule[]
    /**
     * [] for routine snapshots since 2026-11-05: a full tombstone list made
     * every snapshot multi-MB (measured 3.7 MB per push at 2000 tombstones).
     * The one-off 'upgrade' and 'first-enable' snapshots carry the list
     * (2026-11-06) — restore unions it back.
     */
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

export type SyncWipeMarker = {
  /**
   * v4: random id of this wipe. A machine applies a marker at most once —
   * by id, never by comparing timestamps written by different machines'
   * clocks (2026-11-05: a machine whose clock ran 10 min behind re-applied a
   * wipe; one that pushed before pulling never applied it). Absent on v3
   * markers; readers synthesise `legacy:<byMachineId>:<at>`.
   */
  id?: string
  /** ISO of when the wipe was issued (display only). */
  at: string
  /** Issuing machine — pulls echo-guard on this. */
  byMachineId: string
}

export type SyncMeta = {
  /** Bump on schema change so old clients can refuse. */
  schemaVersion: number
  /** Machine that produced this cloud state. Used to recognise our echoes. */
  sourceMachineId: string
  /** ISO timestamp of when this state was written. */
  updatedAt: string
  /** Full rule rows + tombstone counts — quick health check before reading chunks. */
  ruleCount: number
  tombstoneCount: number
  /** v1/v2 full-rule chunk count (syncRules_*). v3/v4 write 0. */
  ruleChunkCount: number
  /** v1–v3 tombstone chunk count (syncTombstones_*). v4 writes 0 (see `v4`). */
  tombstoneChunkCount: number
  /** v3 compact layout (syncRulesC_* rows + syncFolders_* dictionary). */
  ruleV3ChunkCount?: number
  folderChunkCount?: number
  /** v2/v3 folderActivity (syncFolderActivity_*). Also the entry count in v4. */
  folderActivityCount?: number
  folderActivityChunkCount?: number
  /**
   * Rolling log of recent push events from ANY machine that has pushed
   * to this cloud state (cap SYNC_RECENT_PUSHES_CAP, newest first). Lets
   * the Options UI show "A 推了 5 分鐘前 · B 推了 2 小時前" so the user
   * understands multi-machine activity at a glance.
   */
  recentPushes?: SyncRecentPush[]
  /**
   * Cross-machine wipe propagation marker. Set by `wipeAllRules` →
   * `pushNow('post-wipe', { wipeMarker: true })`; preserved by every later
   * push so a machine that connects later still sees it, replaced only by a
   * fresh wipe. Other machines that already participate in sync apply an
   * UNSEEN marker once: drop their syncable rules (auto_scan stays), keep
   * their own tombstones (F3), pin a pre-wipe backup, show a notice.
   */
  wipeMarker?: SyncWipeMarker
  /**
   * v4: random id of this push, stamped into EVERY chunk (`p`) and the
   * settings item. A pull accepts only a snapshot whose items all carry it.
   */
  pushId?: string
  /** v4: fingerprint of the pushed content (no ids/times) — lets an identical push be skipped. */
  contentHash?: string
  /**
   * v4 additive (2026-11-06; v1.0.8 ignores it): fingerprint of the content
   * WITHOUT hit counters (matchCount / overrideCount / lastUsedAt). Equal
   * semanticHash + different contentHash = a counter-only change → coalesced
   * (see SYNC_COUNTER_PUSH_COALESCE_MS).
   */
  semanticHash?: string
  /** v4 additive (2026-11-06): fingerprint of the active-day ledger (settings item `d`). */
  daysHash?: string
  /** v4: disabled rules present only as stubs. */
  disabledStubCount?: number
  /** v4 chunk counts (syncRows4_* / syncFolders4_* / syncTombs4_* / syncActivity4_*). */
  v4?: { rows: number; folders: number; tombs: number; activity: number }
}

// ---- state -----------------------------------------------------------------

// Bug #Q: appendBackup does read-modify-write on syncBackups in
// chrome.storage.local. Serialize the whole RMW via this chain.
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

// ONE queue for every push, pull, restore and cloud clear (2026-11-05). The
// old engine guarded push and pull with separate flags, so a push could
// snapshot local state while a pull was rewriting it (and upload the
// half-merged result), and a remote pull arriving during another pull was
// simply dropped. Ops now run strictly one at a time; nothing is refused.
let syncChain: Promise<unknown> = Promise.resolve()

function enqueue<T>(fn: () => Promise<T>): Promise<T> {
  const run = syncChain.then(fn)
  syncChain = run.catch(() => undefined)
  return run
}

let pushTimer: ReturnType<typeof setTimeout> | null = null
// The debounced push must not defer counter-only changes (a flush was asked
// for while it was pending: 立即上傳, the coalesced counter push, …).
let pendingPushFlush = false
// In-memory twin of the COUNTER_PUSH_ALARM (the alarm survives SW restarts).
let counterPushTimer: ReturnType<typeof setTimeout> | null = null
let counterPushArmedFor: number | undefined
// Debounce timer for remote-pull dispatch: a burst of remote onChanged
// events (meta + late chunks, several pushes) collapses into one pull.
let pullDebounceTimer: ReturnType<typeof setTimeout> | null = null
// True while a remote pull is waiting in the queue: further remote changes
// need no new entry — the queued pull reads the newest meta when it runs. It
// is cleared when the pull STARTS, so a change landing mid-pull queues
// another (the old engine dropped it: "pull already in flight").
let remotePullQueued = false
// Echo window for the local listener: true from pull start until
// PULL_GRACE_MS after it ends (chrome.storage.onChanged is dispatched as a
// macrotask, after the pull's finally). Local changes seen in the window are
// DEFERRED to a push after it, never dropped (2026-11-05: they used to be
// lost). The pull's own writes defer a push too — harmless, because a push
// whose content equals the cloud's is skipped (meta.contentHash).
let pullInProgress = false
let deferredLocalChange = false
let pullGraceTimer: ReturnType<typeof setTimeout> | null = null
// Torn-read retry (item 4).
let tornAttempts = 0
let tornRetryTimer: ReturnType<typeof setTimeout> | null = null
// Failed-push retry (2026-11-05: failed pushes used to wait for the next
// local mutation; rate-limit errors were retried immediately).
let pushRetryAttempts = 0
let pushRetryTimer: ReturnType<typeof setTimeout> | null = null
let startupTimer: ReturnType<typeof setTimeout> | null = null
// pushId of our latest push (recognises our own chunk echoes in the listener).
let lastPushedPushId: string | undefined
// First push of each SW lifetime scans every key for orphans; later pushes
// derive orphans from the previous meta's counts (no sync.get(null) per push).
let fullOrphanScanDone = false

/**
 * Test-only: drop timers, the queue and in-memory flags so one test's
 * scheduled retry can't fire into the next test's storage.
 */
export function resetSyncEngineForTests(): void {
  for (const t of [
    pushTimer,
    pullDebounceTimer,
    pullGraceTimer,
    tornRetryTimer,
    pushRetryTimer,
    startupTimer,
    counterPushTimer,
  ]) {
    if (t) clearTimeout(t)
  }
  pushTimer = pullDebounceTimer = pullGraceTimer = tornRetryTimer = pushRetryTimer = startupTimer = null
  counterPushTimer = null
  counterPushArmedFor = undefined
  pendingPushFlush = false
  syncChain = Promise.resolve()
  remotePullQueued = false
  pullInProgress = false
  deferredLocalChange = false
  tornAttempts = 0
  pushRetryAttempts = 0
  lastPushedPushId = undefined
  fullOrphanScanDone = false
}

// ---- public API ------------------------------------------------------------

/**
 * Installer — called on module load so the listener survives service-worker
 * restarts.
 *
 * Two listeners:
 *   1. `chrome.storage.onChanged` for the 'sync' area → detect REMOTE
 *      changes from another machine and pull them in.
 *   2. `chrome.storage.onChanged` for the 'local' area → detect LOCAL
 *      rule / tombstone / settings mutations (which all flow through
 *      chrome.storage.local) and schedule a debounced push.
 *
 * Plus a startup retry (2026-11-05): a failed push's backoff timer dies with
 * the service worker, so a push that failed in a previous SW lifetime (its
 * error is still recorded) is retried shortly after start. Remote changes
 * need no such help — chrome.storage.onChanged itself wakes the SW.
 */
export function installSyncListener(): void {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync') {
      onSyncAreaChanged(changes)
      return
    }
    if (area === 'local') {
      onLocalAreaChanged(changes)
    }
  })
  // The coalesced counter push (2026-11-06) is armed as an alarm too: an
  // in-memory timer dies with the service worker long before 20 minutes.
  if (typeof chrome.alarms?.onAlarm?.addListener === 'function') {
    chrome.alarms.onAlarm.addListener((alarm) => {
      if (alarm.name === COUNTER_PUSH_ALARM) schedulePush('counter-flush', 0, { flush: true })
    })
  }
  if (startupTimer) clearTimeout(startupTimer)
  startupTimer = setTimeout(() => {
    startupTimer = null
    void readSyncError()
      .then((err) => {
        if (err?.source === 'push') schedulePush('startup-retry', 0)
      })
      .catch(() => {})
    // A counter push that came due while no service worker was alive.
    void readLocalSyncState()
      .then(({ state }) => {
        if (state.counterPushDueAt !== undefined) armCounterPush(state.counterPushDueAt)
      })
      .catch(() => {})
  }, SYNC_STARTUP_RECONCILE_DELAY_MS)
}

/** (Re-)arm the coalesced counter push for `dueAt` (timer + alarm). */
function armCounterPush(dueAt: number): void {
  if (counterPushTimer && counterPushArmedFor === dueAt) return // already armed this SW lifetime
  if (counterPushTimer) clearTimeout(counterPushTimer)
  counterPushArmedFor = dueAt
  counterPushTimer = setTimeout(() => {
    counterPushTimer = null
    counterPushArmedFor = undefined
    schedulePush('counter-flush', 0, { flush: true })
  }, Math.max(0, dueAt - Date.now()))
  try {
    void Promise.resolve(
      chrome.alarms?.create?.(COUNTER_PUSH_ALARM, { when: Math.max(dueAt, Date.now() + 1000) }),
    ).catch(() => {})
  } catch {
    // alarms unavailable — the in-memory timer and the next push cover it
  }
}

function disarmCounterPush(): void {
  if (counterPushTimer) {
    clearTimeout(counterPushTimer)
    counterPushTimer = null
  }
  counterPushArmedFor = undefined
  try {
    void Promise.resolve(chrome.alarms?.clear?.(COUNTER_PUSH_ALARM)).catch(() => {})
  } catch {
    // ignore
  }
}

function onSyncAreaChanged(changes: Record<string, chrome.storage.StorageChange>): void {
  const metaChange = changes[SYNC_META_KEY]
  if (metaChange) {
    const nv = metaChange.newValue as SyncMeta | undefined
    if (!nv) return // cloud cleared elsewhere — nothing to pull
    if (nv.pushId && nv.pushId === lastPushedPushId) return // our own echo
    requestRemotePull('remote-change')
    return
  }
  // Item 4 (2026-11-05): a push's chunks may reach us AFTER its meta (the
  // pull that saw the meta first aborted as torn). The late chunks arrive as
  // their own change event — re-pull on it instead of only on syncMeta.
  for (const [k, ch] of Object.entries(changes)) {
    if (!isOurChunkKey(k) && k !== V4_SETTINGS_KEY) continue
    if (ch.newValue === undefined) continue // orphan cleanup elsewhere
    const p = (ch.newValue as { p?: unknown } | null)?.p
    if (typeof p === 'string' && p === lastPushedPushId) continue
    requestRemotePull('remote-chunk')
    return
  }
}

function onLocalAreaChanged(changes: Record<string, chrome.storage.StorageChange>): void {
  // Settings changes that ONLY touch lastSyncAt are our own bookkeeping —
  // filtered out to avoid push loops. folderActivity changes too: a batch
  // execution or the user clicking 重新整理 updates folderActivity, which the
  // OTHER machine wants to see.
  if (
    !changes['rules'] &&
    !changes['ruleTombstones'] &&
    !changes['folderActivity'] &&
    !isUserFacingSettingsChange(changes['settings'])
  ) {
    return
  }
  if (pullInProgress) {
    // Can't tell the pull's own writes from a concurrent genuine change
    // (a batch finishing, an Options edit). Defer — never drop — and push
    // once the echo window closes; an echo-only push is skipped as a no-op.
    deferredLocalChange = true
    return
  }
  schedulePush('local-mutation')
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
    'harvestEnabled', // per-device opt-in (2026-11-05: toggling it pushed for nothing)
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
export function schedulePush(
  reason: string,
  delayMs: number = PUSH_DEBOUNCE_MS,
  opts: { flush?: boolean } = {},
): void {
  if (pushTimer) clearTimeout(pushTimer)
  if (opts.flush) pendingPushFlush = true
  pushTimer = setTimeout(() => {
    pushTimer = null
    const flush = pendingPushFlush
    pendingPushFlush = false
    void enqueue(() => doPush(reason, { flush })).catch((e) =>
      console.warn(`[mail-organizer] sync push failed (${reason})`, e),
    )
  }, delayMs)
}

export type PushResult = {
  pushed: boolean
  reason?: string
  /** Syncable rules with NO cloud representation (no full row, no stub). */
  truncatedRuleCount?: number
  /** Tombstones from the last 30 days that did not fit the cloud budget. */
  truncatedTombstoneCount?: number
  /** True when the cloud already held exactly this content (nothing written). */
  unchanged?: boolean
  /**
   * Only hit counters differed from the cloud: the write was postponed to a
   * coalesced push (2026-11-06). Never set for a flushing push (pushNow).
   */
  deferred?: boolean
}

/**
 * Force-push now (manual button / first-enable / post-wipe). Bypasses
 * debounce; waits its turn in the sync queue.
 *
 * `options.wipeMarker = true` stamps a fresh `wipeMarker` (new id) into the
 * pushed SyncMeta — used by wipeAllRules so other machines pull, see the
 * unseen marker, and apply the wipe locally. A wipe push never pulls first:
 * it is meant to overwrite the cloud.
 *
 * pushNow always FLUSHES (2026-11-06): pending counter-only changes are
 * written now instead of waiting for the coalesced push.
 */
export async function pushNow(
  reason: string,
  options: { wipeMarker?: boolean } = {},
): Promise<PushResult> {
  if (pushTimer) {
    clearTimeout(pushTimer)
    pushTimer = null
  }
  pendingPushFlush = false
  return enqueue(() => doPush(reason, { ...options, flush: true }))
}

/**
 * Write any counter-only changes still waiting for their coalesced push
 * (2026-11-06) — called before sync is switched off, so the last hit counts
 * of this machine aren't stranded locally. No-op when nothing is pending.
 */
export async function flushPendingPush(reason: string): Promise<PushResult | undefined> {
  const { state } = await readLocalSyncState()
  if (state.counterPushDueAt === undefined && !pushTimer) return undefined
  return pushNow(reason)
}

export type PullResult = {
  pulled: boolean
  reason?: string
  ruleCount?: number
  tombstoneCount?: number
  mergeMode?: PullMergeMode
  /** The cloud snapshot was incomplete / mid-delivery; a retry is scheduled. */
  torn?: boolean
}

/**
 * Pull cloud state into local — manual button / first-enable / remote change.
 *
 * Default mode is 'replace' (cloud-exact; only for explicit callers). Pass
 * mode='union' wherever local may hold rules the cloud never saw.
 * `backup: true` snapshots local state first (first-enable merges); replace
 * pulls and remote wipes always do.
 */
export async function pullNow(
  reason: string,
  options: { mode?: PullMergeMode; backup?: boolean } = {},
): Promise<PullResult> {
  return enqueue(() =>
    doPull(reason, { mode: options.mode ?? 'replace', source: 'pull', backup: options.backup }),
  )
}

/**
 * Wait for queued / in-flight sync ops to drain. Used by disableSync and the
 * wipe flow before clearing cloud, so an in-flight push can't re-write the
 * keys we're about to remove (Bug #N). Cancels pending timers first. Max wait
 * QUIESCE_TIMEOUT_MS — beyond that an op is hung and we give up rather than
 * block UI forever.
 */
export async function quiesce(): Promise<void> {
  for (const t of [pushTimer, pullDebounceTimer, tornRetryTimer, pushRetryTimer, startupTimer, counterPushTimer]) {
    if (t) clearTimeout(t)
  }
  pushTimer = pullDebounceTimer = tornRetryTimer = pushRetryTimer = startupTimer = null
  counterPushTimer = null
  counterPushArmedFor = undefined
  pendingPushFlush = false
  let timer: ReturnType<typeof setTimeout> | undefined
  const timedOut = new Promise<'timeout'>((r) => {
    timer = setTimeout(() => r('timeout'), QUIESCE_TIMEOUT_MS)
  })
  const which = await Promise.race([enqueue(async () => 'drained' as const), timedOut])
  if (timer) clearTimeout(timer)
  if (which === 'timeout') {
    console.warn('[mail-organizer] quiesce timed out — a sync op is still in flight after 5s')
  }
}

/**
 * Wipe everything we put in chrome.storage.sync (every schema's keys).
 * Called by "停用並清除雲端資料" and the wipe flow. Doesn't touch local state.
 */
export async function clearCloudState(): Promise<void> {
  return enqueue(async () => {
    const all = await chrome.storage.sync.get(null)
    const ourKeys = Object.keys(all).filter(isOurSyncKey)
    if (ourKeys.length > 0) await chrome.storage.sync.remove(ourKeys)
  })
}

/**
 * Forget this machine's sync participation (2026-11-05). Called when sync is
 * disabled, so re-enabling counts as FIRST participation again: a wipe marker
 * still in the cloud is then recorded as seen, never applied to the library
 * this machine may have built while sync was off.
 *
 * Also the rollback of a failed enable (S6, 2026-11-06), with
 * `keepTornStreak`: a cloud left torn by a vanished machine is only rewritten
 * once the same torn pushId has been seen for a minute (tornRecoveryAllowed),
 * so the streak must survive the user's repeated 啟用同步 attempts (S7).
 * A machine upgraded from v1.0.7 that hasn't completed its first v4 merge
 * keeps legacyLastSyncAt: that merge is still an upgrade (A1/A4).
 */
export async function resetSyncParticipation(opts: { keepTornStreak?: boolean } = {}): Promise<void> {
  return enqueue(async () => {
    const { state } = await readLocalSyncState()
    await writeLocalSyncState({
      participating: false,
      ...(state.legacyLastSyncAt !== undefined ? { legacyLastSyncAt: state.legacyLastSyncAt } : {}),
    })
    if (!opts.keepTornStreak) await chrome.storage.local.remove(SYNC_TORN_STREAK_KEY)
    tornAttempts = 0
    resetPushRetry()
    disarmCounterPush()
  })
}

/**
 * Sync is being switched ON (2026-11-06, A12): this machine participates
 * only from now on. A machine that switched sync off under v1.0.7 has no v4
 * state and a stale settings.lastSyncAt, which used to count as
 * participation — re-enabling then applied a wipe marker issued while it was
 * away to the library it built meanwhile. Records non-participation unless a
 * v4 state already exists (then disable already reset it, or sync is on).
 * Its old lastSyncAt is kept for one purpose only: its first v4 merge is
 * still an upgrade (pinned backup; tombstones v3 had processed by then are
 * not news — see adoptionProtects).
 */
export async function beginSyncParticipation(): Promise<void> {
  return enqueue(async () => {
    const r = await chrome.storage.local.get(SYNC_LOCAL_STATE_KEY)
    if (!r[SYNC_LOCAL_STATE_KEY] || typeof r[SYNC_LOCAL_STATE_KEY] !== 'object') {
      const { lastSyncAt } = await getSettings()
      await writeLocalSyncState({ participating: false, ...(lastSyncAt ? { legacyLastSyncAt: lastSyncAt } : {}) })
    }
  })
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
    schemaVersion?: number
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
          schemaVersion: meta.schemaVersion,
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
  const list = r[SYNC_BACKUPS_KEY]
  return Array.isArray(list) ? (list as SyncBackup[]) : []
}

/**
 * Snapshot local rules + settings (2026-11-05: ONLY before a replace merge,
 * a remote-wipe apply, a first-enable merge, or a restore — no longer on
 * every push and pull, which rewrote multiple MB each time and rotated the
 * pre-wipe snapshot out within minutes). Throws on storage failure; callers
 * record it in lastError and carry on.
 */
async function appendBackup(
  direction: SyncBackup['direction'],
  trigger: NonNullable<SyncBackup['trigger']>,
  opts: { pinned?: boolean; withTombstones?: boolean } = {},
): Promise<void> {
  // Bug #Q: serialize the read-list / prepend / slice / write sequence so
  // concurrent appendBackup calls can't clobber each other's snapshot.
  return withBackupLock(async () => {
    const [rules, settings, existing, tombstones] = await Promise.all([
      listRules(),
      getSettings(),
      listBackups(),
      opts.withTombstones ? getRuleTombstones() : Promise.resolve([] as RuleTombstone[]),
    ])
    const { claudeApiKey: _omit, ...settingsNoKey } = settings
    void _omit
    // snapshotAt doubles as the restore handle → keep it unique.
    let at = Date.now()
    const taken = new Set(existing.map((b) => b.snapshotAt))
    while (taken.has(new Date(at).toISOString())) at++
    const snapshot: SyncBackup = {
      snapshotAt: new Date(at).toISOString(),
      direction,
      trigger,
      ...(opts.pinned ? { pinned: true } : {}),
      payload: { rules, tombstones, settings: settingsNoKey },
    }
    // One pinned slot per trigger (a pinned backup from before 2026-11-06
    // carries no trigger — it was always the pre-wipe one).
    const slot = (b: SyncBackup) => b.trigger ?? 'remote-wipe'
    const pinned = [
      ...(opts.pinned ? [snapshot] : []),
      ...existing.filter((b) => b.pinned && !(opts.pinned && slot(b) === trigger)),
    ]
    const rotating = [...(opts.pinned ? [] : [snapshot]), ...existing.filter((b) => !b.pinned)].slice(
      0,
      MAX_BACKUPS,
    )
    const next = [...pinned, ...rotating].sort((a, b) =>
      a.snapshotAt < b.snapshotAt ? 1 : a.snapshotAt > b.snapshotAt ? -1 : 0,
    )
    await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: next })
  })
}

/**
 * The pinned pre-upgrade snapshot (A1/A4). Taken once per machine: a retry
 * after a first merge that threw must not replace it with a later state.
 */
async function appendUpgradeBackupOnce(direction: SyncBackup['direction']): Promise<void> {
  if ((await listBackups()).some((b) => b.trigger === 'upgrade')) return
  await appendBackup(direction, 'upgrade', { pinned: true, withTombstones: true })
}

/** User-facing name of a backup's trigger (Options backup list). */
export function backupLabel(b: Pick<SyncBackup, 'trigger' | 'direction'>): string {
  switch (b.trigger) {
    case 'upgrade':
      return '升級同步格式前'
    case 'remote-wipe':
      return '套用他機全部刪除前'
    case 'first-enable':
      return '啟用同步前'
    case 'replace-pull':
      return '下載覆蓋前'
    case 'restore':
      return '回復備份前'
    case 'torn-recovery':
      return '修復雲端資料前'
    default:
      return b.direction === 'pre-push' ? '上傳前' : b.direction === 'pre-pull' ? '下載前' : '手動'
  }
}

/**
 * Restore a backup by its snapshotAt timestamp. Rules are replaced by the
 * snapshot's (through mutateRules — invariant 8 — with `preserveEditedAt`:
 * the stamps below are restore's own); settings overlay keeps the local
 * per-device fields.
 *
 * Propagation (2026-11-06 — A5): a restore is a rollback the user expects to
 * hold on EVERY machine, so it is written as a new edit:
 *   - a restored version that differs from the current one (or whose rule is
 *     gone here) gets a fresh editedAt, above every known deletion of its
 *     triple — it wins the next merge everywhere. Keeping the snapshot's own
 *     editedAt let any newer copy anywhere revert the "已回復" rollback at the
 *     next pull. A rule identical to the current one keeps that version.
 *   - a syncable rule that exists now but not in the snapshot is removed
 *     with a tombstone (tombstoneDeletedAt), so the removal propagates —
 *     the other machines used to push it right back.
 *   - tombstones the snapshot carries ('upgrade' / 'first-enable') are
 *     unioned back, except where a restored rule stands for the triple.
 * A restored rule whose triple was deleted after the snapshot is an explicit
 * un-delete: its local tombstones are cleared.
 *
 * Restore restores SEMANTIC content only (S2, 2026-11-06): hit counters,
 * their epoch and lastUsedAt are merged with the current copy like a sync
 * merge (assignCounters) — never rolled back. The snapshot's lower counts
 * read as a counter reset against the local baseline, so the sweep judged the
 * rule's whole pre-v1.0.8 record again and disabled it for high error rate.
 */
export async function restoreBackup(
  snapshotAt: string,
): Promise<{ restored: boolean; undeletedCount?: number; removedCount?: number }> {
  return enqueue(async () => {
    const backups = await listBackups()
    const backup = backups.find((b) => b.snapshotAt === snapshotAt)
    if (!backup) return { restored: false }
    // One more level of undo: snapshot the current state first.
    let backupError: string | undefined
    try {
      await appendBackup('manual', 'restore')
    } catch (e) {
      backupError = errorMessage(e)
    }
    const restored = (Array.isArray(backup.payload?.rules) ? backup.payload.rules : [])
      .filter((r): r is Rule => !!r && typeof r === 'object' && typeof r.id === 'string')
      .map(applyConfidenceCap)
    const snapshotTombs = (Array.isArray(backup.payload?.tombstones) ? backup.payload.tombstones : [])
      .map(sanitizeLegacyTombstone)
      .filter((t): t is RuleTombstone => t !== null)
    let undeletedCount = 0
    let removedCount = 0
    await mutateRules(
      async (current) => {
        const localTombs = await getRuleTombstones()
        const newest = newestDeletionByKey(localTombs)
        const now = Date.now()
        const restoredKeys = new Set(restored.map(ruleKey))
        const restoredIds = new Set(restored.map((r) => r.id))
        const currentById = new Map(current.map((r) => [r.id, r]))
        const next = restored.map((r) => {
          const key = ruleKey(r)
          const d = newest.get(key)
          const cur = currentById.get(r.id)
          if (cur && semanticFieldsEqual(cur, r) && (d === undefined || d < ruleTimeMs(cur))) return cur
          // Same predicate as the pull merge's suppression (d >= version time).
          if (d !== undefined && d >= ruleTimeMs(r)) undeletedCount++
          const out: Rule = { ...r }
          if (cur) assignCounters(out, cur, r)
          // A restored orphaned rule carries a restore stamp, never a v1.0.8
          // orphan stamp (S3): record when this machine saw it orphaned.
          const sameDeadTarget =
            cur?.orphaned === true &&
            cur.targetFolderId === out.targetFolderId &&
            cur.targetFolderPath === out.targetFolderPath
          if (out.orphaned) out.orphanedAt = sameDeadTarget ? (cur!.orphanedAt ?? now) : now
          else delete out.orphanedAt
          out.editedAt = Math.max(now, (cur?.editedAt ?? 0) + 1, (r.editedAt ?? 0) + 1, (d ?? 0) + 1)
          return out
        })
        // Un-delete: every local tombstone of a restored triple goes (legacy
        // spellings included — matched by canonical key).
        const clear = localTombs.filter((t) => restoredKeys.has(tombstoneKey(t)))
        if (clear.length > 0) await clearMatchingTombstones(clear)
        // Removals + the snapshot's own tombstones, never for a triple a
        // restored rule stands for — each kept unless a known deletion
        // already deletes everything it deletes, at least as late (round 3:
        // per spelling; newest-per-triple dropped the other twin's spelling).
        const known = groupTombstonesByKey(localTombs)
        const add: RuleTombstone[] = []
        const offer = (t: RuleTombstone) => {
          const k = tombstoneKey(t)
          if (restoredKeys.has(k) || coversDeletion(known.get(k), t)) return
          add.push(t)
          known.set(k, [...(known.get(k) ?? []), t])
        }
        for (const t of [...snapshotTombs].sort((a, b) => b.deletedAt - a.deletedAt)) offer(t)
        for (const r of current) {
          if (restoredIds.has(r.id) || !SYNCABLE_SOURCES.has(r.source)) continue
          if (restoredKeys.has(ruleKey(r))) continue
          removedCount++
          offer(ruleTombstone(r, now))
        }
        if (add.length > 0) await addRuleTombstones(add)
        return { next, result: undefined }
      },
      { preserveEditedAt: true },
    )
    // Settings overlay — strip per-device fields so a backup taken pre-pull
    // can't change our identity on restore. Through setSettings (sanitised).
    if (backup.payload?.settings && typeof backup.payload.settings === 'object') {
      await setSettings(stripPerDeviceSettings(backup.payload.settings as Partial<Settings>) as Partial<Settings>)
    }
    if (backupError) {
      await recordSyncError({ source: 'pull', reason: `backup failed before restore (restore continued): ${backupError}` })
    }
    return { restored: true, undeletedCount, removedCount }
  })
}

// ---- internals: local state ------------------------------------------------

async function readLocalSyncState(): Promise<{ state: SyncLocalState; legacyLastSyncAt?: string }> {
  const r = await chrome.storage.local.get(SYNC_LOCAL_STATE_KEY)
  const raw = r[SYNC_LOCAL_STATE_KEY] as Partial<SyncLocalState> | undefined
  if (raw && typeof raw === 'object') {
    const foreign =
      raw.foreignSettings && typeof raw.foreignSettings === 'object' ? raw.foreignSettings : undefined
    return {
      state: {
        participating: raw.participating === true,
        ...(typeof raw.lastSeenWipeId === 'string' ? { lastSeenWipeId: raw.lastSeenWipeId } : {}),
        ...(typeof raw.lastMergedPushId === 'string' ? { lastMergedPushId: raw.lastMergedPushId } : {}),
        ...(raw.settingsBase && typeof raw.settingsBase === 'object'
          ? { settingsBase: raw.settingsBase }
          : {}),
        ...(raw.v4Adopted === true ? { v4Adopted: true } : {}),
        ...(typeof raw.legacyLastSyncAt === 'string' ? { legacyLastSyncAt: raw.legacyLastSyncAt } : {}),
        ...(foreign && Object.keys(foreign).length > 0 ? { foreignSettings: foreign } : {}),
        ...(typeof raw.counterPushDueAt === 'number' && Number.isFinite(raw.counterPushDueAt)
          ? { counterPushDueAt: raw.counterPushDueAt }
          : {}),
      },
      ...(typeof raw.legacyLastSyncAt === 'string' ? { legacyLastSyncAt: raw.legacyLastSyncAt } : {}),
    }
  }
  // Absent → first run of this build with sync on (enabling sync writes a
  // state first — beginSyncParticipation). A machine that was syncing under
  // v1.0.7 is a participant; its lastSyncAt decides (once) whether the wipe
  // marker currently in the cloud is one it has already applied — the v3
  // rule, kept only for this one transition — and which tombstones v3 had
  // already processed (adoptionProtects). It is carried in the state until
  // the first v4 merge.
  const settings = await getSettings()
  return settings.lastSyncAt
    ? {
        state: { participating: true, legacyLastSyncAt: settings.lastSyncAt },
        legacyLastSyncAt: settings.lastSyncAt,
      }
    : { state: { participating: false } }
}

/**
 * Has this machine's library been through its first-participation merge?
 * States written by v1.0.8 carry no flag: a participating one that already
 * merged or pushed (lastMergedPushId) has been merging under v4 all along.
 */
function isAdopted(state: SyncLocalState): boolean {
  return state.v4Adopted === true || (state.participating && state.lastMergedPushId !== undefined)
}

/** State after the adoption merge: adopted, the v3 transition data dropped. */
function adoptedState(state: SyncLocalState): SyncLocalState {
  const { legacyLastSyncAt: _drop, ...rest } = state
  void _drop
  return { ...rest, v4Adopted: true }
}

async function writeLocalSyncState(state: SyncLocalState): Promise<void> {
  await chrome.storage.local.set({ [SYNC_LOCAL_STATE_KEY]: state })
}

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

function randomId(len = 12): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789'
  let out = ''
  if (typeof crypto !== 'undefined' && 'getRandomValues' in crypto) {
    const bytes = crypto.getRandomValues(new Uint8Array(len))
    for (const b of bytes) out += alphabet[b % alphabet.length]
    return out
  }
  while (out.length < len) out += Math.random().toString(36).slice(2)
  return out.slice(0, len)
}

async function readSyncMeta(): Promise<SyncMeta | undefined> {
  const r = await chrome.storage.sync.get(SYNC_META_KEY)
  const m = r[SYNC_META_KEY]
  return m && typeof m === 'object' ? (m as SyncMeta) : undefined
}

async function getOurBytesInUse(): Promise<number> {
  // chrome.storage.sync is isolated PER EXTENSION — getBytesInUse(null)
  // returns only our own keys, measured in the real UTF-8 bytes Chrome bills
  // the quota in. That is the authoritative number.
  try {
    const authoritative = await chrome.storage.sync.getBytesInUse(null)
    if (typeof authoritative === 'number' && authoritative >= 0) return authoritative
  } catch {
    // getBytesInUse unavailable (e.g. test mock) — fall through to manual sum.
  }
  // Byte-accurate manual fallback. Mirrors Chrome's accounting: key bytes +
  // JSON-value bytes, both in UTF-8 (NOT String.length, which under-counts CJK).
  const all = await chrome.storage.sync.get(null)
  let bytes = 0
  for (const [key, value] of Object.entries(all)) {
    if (!isOurSyncKey(key)) continue
    bytes += utf8ByteLength(key) + jsonByteSize(value)
  }
  return bytes
}

function errorMessage(e: unknown): string {
  return e instanceof Error ? e.message : String(e)
}

// ---- internals: wire helpers -------------------------------------------------

type NormalizedWipeMarker = { id: string; at: string; byMachineId: string }

function normalizeWipeMarker(raw: unknown): NormalizedWipeMarker | undefined {
  if (!raw || typeof raw !== 'object') return undefined
  const m = raw as Record<string, unknown>
  if (typeof m.at !== 'string' || typeof m.byMachineId !== 'string') return undefined
  const id = typeof m.id === 'string' && m.id ? m.id : `legacy:${m.byMachineId}:${m.at}`
  return { id, at: m.at, byMachineId: m.byMachineId }
}

export type PushPayloadInput = {
  machineId: string
  rules: Rule[]
  tombstones: RuleTombstone[]
  folderActivity: FolderActivity[]
  /** Full settings (per-device fields are stripped here). */
  settings: Partial<Settings>
  /** Cloud meta this push replaces (recentPushes / wipe marker carry-over). */
  prevMeta?: SyncMeta
  /** A FRESH marker for a wipe push; otherwise the previous one is carried. */
  wipeMarker?: NormalizedWipeMarker
  /** The stale-sweep active-day ledger ('YYYY-MM-DD'); optional (older callers / tests). */
  activeDays?: readonly string[]
  updatedAt: string
  pushId: string
  now: number
}

export type PushPayload = {
  meta: SyncMeta
  writes: Record<string, unknown>
  plan: SyncPayloadPlan
  contentHash: string
  /** Content minus hit counters (see SyncMeta.semanticHash). */
  semanticHash: string
  /**
   * contentHash as v1.0.8 computes it for the same content (S10): rows
   * without `sr`, tombstones without their spelling entry — the fields a
   * v1.0.8 machine drops when it re-pushes what it merged. Equal to
   * contentHash when this push carries none of them.
   */
  v108ContentHash: string
  daysHash: string
  truncatedActivityCount: number
}

/** Full rows without their hit counters (m / o / lu) — the semanticHash input. */
function withoutCounters(rows: readonly WireRuleEntry[]): unknown[] {
  return rows.map((e) => {
    if (Array.isArray(e)) return e
    const { m: _m, o: _o, lu: _lu, ...rest } = e
    void _m
    void _o
    void _lu
    return rest
  })
}

/**
 * Folder activity as semanticHash sees it (S9, 2026-11-06): which folders
 * are in the synced list, not their metrics. Every batch rewrites
 * lastActiveAt / recentCount / lastBatchAt / latestMessage of the folders it
 * filed into, which kept batch pushes immediate (~88 KB each) despite the
 * counter coalescing; such metric-only changes now wait for the coalesced
 * push like hit counters. A folder entering or leaving the list (or renamed)
 * still pushes at once.
 */
function activityMembership(items: readonly FolderActivity[]): Array<[string, string]> {
  return items.map((e) => [e.folderId, e.folderPath])
}

function buildMeta(
  input: PushPayloadInput,
  hash: string,
  counts: { rows: number; folders: number; tombs: number; activity: number },
  totals: { rules: number; stubs: number; tombs: number; activity: number },
  extra: { semanticHash: string; daysHash: string },
): SyncMeta {
  // #6: extend the recentPushes log with this push's entry; each machine
  // keeps one (its latest). Bug #V: Array.isArray guards a malformed cloud.
  const prevRecent = Array.isArray(input.prevMeta?.recentPushes)
    ? (input.prevMeta!.recentPushes as SyncRecentPush[])
    : []
  const recentPushes: SyncRecentPush[] = [
    { machineId: input.machineId, at: input.updatedAt },
    ...prevRecent.filter(
      (p) => p && typeof p === 'object' && p.machineId !== input.machineId,
    ),
  ].slice(0, SYNC_RECENT_PUSHES_CAP)
  // The marker lives on until the next wipe (a machine connecting later
  // must still see it) — carry the previous one unless this push IS a wipe.
  const wipeMarker = input.wipeMarker ?? normalizeWipeMarker(input.prevMeta?.wipeMarker)
  return {
    schemaVersion: SYNC_SCHEMA_VERSION,
    sourceMachineId: input.machineId,
    updatedAt: input.updatedAt,
    pushId: input.pushId,
    contentHash: hash,
    semanticHash: extra.semanticHash,
    daysHash: extra.daysHash,
    ruleCount: totals.rules,
    disabledStubCount: totals.stubs,
    tombstoneCount: totals.tombs,
    // Legacy counts stay 0: a v1–v3 reader that slipped past its version
    // check reads no rules / tombstones (keeps its local set) — no corruption.
    ruleChunkCount: 0,
    tombstoneChunkCount: 0,
    ruleV3ChunkCount: 0,
    folderChunkCount: 0,
    folderActivityChunkCount: 0,
    folderActivityCount: totals.activity,
    v4: counts,
    recentPushes,
    ...(wipeMarker ? { wipeMarker } : {}),
  }
}

/**
 * Everything one v4 push writes — pure. doPush uses it; tests use it to
 * simulate another machine running this build.
 *
 * Budget: the 100 KB quota is a HARD ceiling (one over-budget set fails the
 * whole batch). Settings + meta + folderActivity are measured exactly and the
 * remainder of SYNC_TOTAL_BUDGET_BYTES goes to planSyncPayload (rows, stubs,
 * folder dictionary, tombstones). Rules that don't fit stay local — this only
 * limits cross-machine propagation — and are counted.
 */
export function buildPushPayload(input: PushPayloadInput): PushPayload {
  const { pushId } = input
  const settingsToSync = stripPerDeviceSettings(input.settings)
  // The active-day ledger rides in the settings item as `d` (2026-11-06): it
  // carries the pushId check like every other item, and a v1.0.8 reader only
  // looks at `p` / `v`. Dropped (never the settings) if the item would grow
  // past one chunk.
  let days = encodeActiveDays(input.activeDays ?? [], ACTIVE_DAYS_CAP)
  let settingsValue: { p: string; v: typeof settingsToSync; d?: number[] } =
    days.length > 0 ? { p: pushId, v: settingsToSync, d: days } : { p: pushId, v: settingsToSync }
  if (days.length > 0 && writeBytes(V4_SETTINGS_KEY, settingsValue) > CHUNK_BYTE_TARGET) {
    days = []
    settingsValue = { p: pushId, v: settingsToSync }
  }
  // chunkFolderActivity already truncates to FOLDER_ACTIVITY_SYNC_CAP=20; the
  // chunk cap is defense-in-depth against one huge entry.
  const allActivityChunks = chunkFolderActivity(input.folderActivity, pushId)
  const activityChunks = allActivityChunks.slice(0, MAX_FOLDER_ACTIVITY_CHUNKS)
  const activityItems = activityChunks.flatMap((c) => c.items)
  const truncatedActivityCount =
    allActivityChunks.reduce((s, c) => s + c.items.length, 0) - activityItems.length

  // Size the meta with worst-case counts (digits only grow it slightly).
  const big = { rows: 99, folders: 99, tombs: 99, activity: 99 }
  const metaUpperBound = buildMeta(
    input,
    'x'.repeat(12),
    big,
    { rules: 99999, stubs: 99999, tombs: 99999, activity: 99999 },
    { semanticHash: 'x'.repeat(12), daysHash: 'x'.repeat(12) },
  )
  let fixedBytes =
    writeBytes(V4_SETTINGS_KEY, settingsValue) + writeBytes(SYNC_META_KEY, metaUpperBound)
  for (const c of activityChunks) fixedBytes += writeBytes(`${V4_ACTIVITY_PREFIX}${c.index}`, c)

  const plan = planSyncPayload({
    rules: input.rules,
    tombstones: input.tombstones,
    budgetBytes: SYNC_TOTAL_BUDGET_BYTES - fixedBytes,
    tombstoneShareBytes: Math.floor(SYNC_TOMBSTONE_BUDGET_SHARE * SYNC_TOTAL_BUDGET_BYTES),
    recentTombstoneShareBytes: Math.floor(SYNC_RECENT_TOMBSTONE_BUDGET_SHARE * SYNC_TOTAL_BUDGET_BYTES),
    pushId,
    keyPrefix: { rows: V4_ROWS_PREFIX, folders: V4_FOLDERS_PREFIX, tombs: V4_TOMBS_PREFIX },
    now: input.now,
  })

  // Fingerprint of the CONTENT only (no pushId / times): equal content →
  // equal hash, so an echo push after a pull can be skipped. Same formula as
  // v1.0.8 (the ledger has its own daysHash), so echo detection keeps working
  // across the two versions.
  const hash = contentHash(
    JSON.stringify([plan.rows, plan.folders, plan.tombstones, activityItems, settingsToSync]),
  )
  const semanticHash = contentHash(
    JSON.stringify([
      withoutCounters(plan.rows),
      plan.folders,
      plan.tombstones,
      activityMembership(activityItems),
      settingsToSync,
    ]),
  )
  // S10: what a v1.0.8 peer's push of this very content would hash to.
  const hasSr = plan.rows.some((e) => !Array.isArray(e) && (e as CompactRuleRow).sr !== undefined)
  const hasSpelling = plan.tombstones.some((t) => t.length > 4)
  const v108ContentHash =
    hasSr || hasSpelling
      ? contentHash(
          JSON.stringify([
            hasSr ? plan.rows.map(withoutV109RowFields) : plan.rows,
            plan.folders,
            hasSpelling ? tombstonesAsV108Keeps(plan.tombstones) : plan.tombstones,
            activityItems,
            settingsToSync,
          ]),
        )
      : hash
  const daysHash = contentHash(JSON.stringify(days))
  const meta = buildMeta(
    input,
    hash,
    {
      rows: plan.rowChunks.length,
      folders: plan.folderChunks.length,
      tombs: plan.tombChunks.length,
      activity: activityChunks.length,
    },
    {
      rules: plan.fullRowCount,
      stubs: plan.stubCount,
      tombs: plan.tombstones.length,
      activity: activityItems.length,
    },
    { semanticHash, daysHash },
  )
  const writes: Record<string, unknown> = {
    [SYNC_META_KEY]: meta,
    [V4_SETTINGS_KEY]: settingsValue,
  }
  for (const c of plan.rowChunks) writes[`${V4_ROWS_PREFIX}${c.index}`] = c
  for (const c of plan.folderChunks) writes[`${V4_FOLDERS_PREFIX}${c.index}`] = c
  for (const c of plan.tombChunks) writes[`${V4_TOMBS_PREFIX}${c.index}`] = c
  for (const c of activityChunks) writes[`${V4_ACTIVITY_PREFIX}${c.index}`] = c
  return { meta, writes, plan, contentHash: hash, semanticHash, v108ContentHash, daysHash, truncatedActivityCount }
}

/**
 * Tombstone rows as a v1.0.8 machine re-pushes them after merging: no
 * spelling entry, and ONE row per stored triple — the newest; it keeps one
 * tombstone per triple (round 3: this build may push one per spelling).
 * Chunk order is preserved (a triple's rows are adjacent).
 */
function tombstonesAsV108Keeps(rows: readonly CompactTombstone[]): CompactTombstone[] {
  const out: CompactTombstone[] = []
  let lastKey: string | undefined
  for (const t of rows) {
    const key = JSON.stringify([t[0], t[1], t[2]])
    const row: CompactTombstone = [t[0], t[1], t[2], t[3]]
    if (key === lastKey) {
      if (t[3] > out[out.length - 1]![3]) out[out.length - 1] = row
    } else {
      out.push(row)
      lastKey = key
    }
  }
  return out
}

/** A row as a v1.0.8 encoder writes the same rule: no `sr` (key order otherwise identical). */
function withoutV109RowFields(e: WireRuleEntry): WireRuleEntry {
  if (Array.isArray(e) || (e as CompactRuleRow).sr === undefined) return e
  const { sr: _sr, ...rest } = e as CompactRuleRow
  void _sr
  return rest
}

/** Every key the cloud state described by `meta` occupies. */
function keysOfMeta(meta: SyncMeta | undefined): string[] {
  if (!meta) return []
  if (meta.schemaVersion >= 4 && meta.v4 && typeof meta.v4 === 'object') {
    return [
      V4_SETTINGS_KEY,
      ...rangeKeys(V4_ROWS_PREFIX, meta.v4.rows),
      ...rangeKeys(V4_FOLDERS_PREFIX, meta.v4.folders),
      ...rangeKeys(V4_TOMBS_PREFIX, meta.v4.tombs),
      ...rangeKeys(V4_ACTIVITY_PREFIX, meta.v4.activity),
    ]
  }
  return [
    LEGACY_SETTINGS_KEY,
    ...rangeKeys(LEGACY_RULES_PREFIX, meta.ruleChunkCount),
    ...rangeKeys(V3_ROWS_PREFIX, meta.ruleV3ChunkCount),
    ...rangeKeys(V3_FOLDERS_PREFIX, meta.folderChunkCount),
    ...rangeKeys(LEGACY_TOMBS_PREFIX, meta.tombstoneChunkCount),
    ...rangeKeys(LEGACY_ACTIVITY_PREFIX, meta.folderActivityChunkCount),
  ]
}

type StorageErrorKind = 'rate-minute' | 'rate-hour' | 'quota-bytes' | 'other'

/**
 * Classify a chrome.storage.sync failure (2026-11-05). `/quota/i` used to
 * match the write-RATE limits too ("MAX_WRITE_OPERATIONS_PER_MINUTE quota
 * exceeded"), so a rate-limited push deleted orphan chunks and retried at
 * once — burning more of the very write budget it had run out of.
 */
function classifyStorageError(msg: string): StorageErrorKind {
  if (/MAX_WRITE_OPERATIONS_PER_HOUR/i.test(msg)) return 'rate-hour'
  if (/WRITE_OPERATIONS/i.test(msg)) return 'rate-minute'
  if (/QUOTA_BYTES|quota/i.test(msg)) return 'quota-bytes'
  return 'other'
}

// ---- push ----------------------------------------------------------------------

async function doPush(
  reason: string,
  options: { wipeMarker?: boolean; flush?: boolean } = {},
): Promise<PushResult> {
  try {
    return await doPushInner(reason, options)
  } catch (e) {
    // Anything unexpected (local storage read failure, …) used to escape to a
    // console-only catch in schedulePush — sync then died silently.
    const msg = errorMessage(e)
    console.warn(`[mail-organizer] sync push (${reason}) failed`, e)
    await recordSyncError({ source: 'push', reason: msg }).catch(() => {})
    await logError('sync:push', msg, { reason }).catch(() => {})
    schedulePushRetry('other')
    return { pushed: false, reason: msg }
  }
}

function newerThanUs(meta: SyncMeta): boolean {
  return typeof meta.schemaVersion === 'number' && meta.schemaVersion > SYNC_SCHEMA_VERSION
}

function sameLegacyMeta(a: SyncMeta, b: SyncMeta): boolean {
  return (
    a.schemaVersion === b.schemaVersion &&
    a.updatedAt === b.updatedAt &&
    a.sourceMachineId === b.sourceMachineId
  )
}

/** Why a cloud meta can't be read by this build, or undefined if it can. */
function schemaRefusal(meta: SyncMeta): string | undefined {
  const v = meta.schemaVersion
  if (typeof v !== 'number' || !Number.isInteger(v) || v < SYNC_MIN_READABLE_SCHEMA_VERSION) {
    return `unrecognized cloud data (schemaVersion=${String(v)})`
  }
  if (v > SYNC_SCHEMA_VERSION) {
    return `cloud schemaVersion=${v} is newer than this extension supports (${SYNC_SCHEMA_VERSION}); upgrade this browser's extension`
  }
  return undefined
}

async function refuseNewerCloud(meta: SyncMeta): Promise<PushResult> {
  // Downgrade protection (Bug #H): overwriting a newer schema would silently
  // corrupt the other machine's data with a lower-version payload.
  const reason = `cloud schemaVersion=${meta.schemaVersion} > ours=${SYNC_SCHEMA_VERSION}; upgrade this browser's extension before pushing`
  await recordSyncError({ source: 'push', reason })
  return { pushed: false, reason }
}

async function doPushInner(
  reason: string,
  options: { wipeMarker?: boolean; flush?: boolean },
): Promise<PushResult> {
  const settings = await getSettings()
  if (!settings.syncEnabled) return { pushed: false, reason: 'sync disabled' }
  let cloudMeta = await readSyncMeta()
  if (cloudMeta && newerThanUs(cloudMeta)) return refuseNewerCloud(cloudMeta)
  const machineId = await ensureMachineId()
  let { state } = await readLocalSyncState()
  // A6 (2026-11-06): the cloud stayed torn under one pushId — rewrite it.
  let recovering = false

  // Pull before push (item 2, 2026-11-05) whenever the cloud holds a state
  // this machine hasn't merged — the first push after SW start, after a
  // refused / failed pull, or after another machine pushed. Pushing without
  // it overwrote the other machine's unmerged edits (and a wipe marker a
  // machine pushed past was never applied). Skipped only when the cloud is
  // provably what we last merged or wrote (same pushId: the pull would be a
  // no-op), and for a wipe push (it overwrites the cloud on purpose).
  if (cloudMeta && !options.wipeMarker) {
    for (let attempt = 0; ; attempt++) {
      if (cloudMeta.pushId && cloudMeta.pushId === state.lastMergedPushId) break
      if (attempt >= 3) {
        // The cloud kept changing under us — try again shortly.
        schedulePushRetry('other')
        return { pushed: false, reason: 'cloud kept changing during pull-before-push; retry scheduled' }
      }
      // A first-participation merge (an enable's push path, S7) snapshots
      // local state first, like the enable's own pull does.
      const pr = await doPull(`pre-push:${reason}`, { mode: 'union', source: 'pull', backup: !state.participating })
      if (!pr.pulled) {
        if (pr.reason === 'no cloud state') {
          cloudMeta = undefined
          break
        }
        // A6: a snapshot torn under the SAME pushId for many reads and over a
        // minute is not a push in flight — its owner is gone or can't repair
        // it, and every other machine was locked out of pushing (立即上傳
        // included). Back up, then rewrite the cloud from local; the other
        // machines union-merge it and push their own rules back.
        if (pr.torn && (await tornRecoveryAllowed(machineId))) {
          try {
            await appendBackup('pre-push', 'torn-recovery')
          } catch (e) {
            console.warn('[mail-organizer] backup before torn-cloud recovery failed (recovery continues)', e)
          }
          recovering = true
          cloudMeta = (await readSyncMeta()) ?? cloudMeta
          state = (await readLocalSyncState()).state
          break
        }
        schedulePushRetry('other')
        return { pushed: false, reason: `pull before push failed: ${pr.reason ?? 'unknown'}` }
      }
      const merged: SyncMeta = cloudMeta
      cloudMeta = await readSyncMeta()
      state = (await readLocalSyncState()).state
      if (!cloudMeta) break
      if (newerThanUs(cloudMeta)) return refuseNewerCloud(cloudMeta)
      // A legacy (v1–v3) cloud carries no pushId: one merge of an unchanged
      // meta is all it needs before we overwrite it as v4.
      if (!cloudMeta.pushId && sameLegacyMeta(merged, cloudMeta)) break
    }
  }

  // Adoption (2026-11-06) when this machine's first v4 operation is a push
  // (nothing in the cloud to merge, or a torn one): protect its own library
  // the way a first merge would — before it is uploaded and then merged back.
  if (!isAdopted(state)) {
    if (!options.wipeMarker) await adoptLocalLibrary(state)
    state = adoptedState(state)
    await writeLocalSyncState(state)
  }

  // Read local state AFTER the pull: it may have merged rules and overlaid
  // the other machine's settings, which must not be pushed back stale.
  const [rules, tombstones, folderActivity, pushSettings, activeDays] = await Promise.all([
    listRules(),
    getRuleTombstones(),
    getFolderActivity(),
    getSettings(),
    getActiveDays(),
  ])
  if (!pushSettings.syncEnabled) return { pushed: false, reason: 'sync disabled' }
  // A13: a synced value this build can't represent goes back out unchanged.
  const outgoing = applyForeignSettings(pushSettings, state.foreignSettings)
  const updatedAt = new Date().toISOString()
  const pushId = randomId()
  const freshMarker = options.wipeMarker
    ? { id: randomId(16), at: updatedAt, byMachineId: machineId }
    : undefined
  const payload = buildPushPayload({
    machineId,
    rules,
    tombstones,
    folderActivity,
    settings: outgoing.settings,
    activeDays,
    prevMeta: cloudMeta,
    wipeMarker: freshMarker,
    updatedAt,
    pushId,
    now: Date.now(),
  })
  const { plan } = payload
  const truncatedRuleCount = plan.truncatedRuleCount
  const truncatedTombCount = plan.truncatedTombstoneCount
  if (truncatedRuleCount > 0 || plan.omittedTombstoneCount > 0 || payload.truncatedActivityCount > 0) {
    console.warn(
      `[mail-organizer] sync push: ${truncatedRuleCount} rules not represented in the cloud` +
        ` (${plan.stubCount} disabled rules sent as stubs), ${plan.omittedTombstoneCount} older tombstones` +
        ` and ${payload.truncatedActivityCount} folderActivity entries left out by the byte budget.` +
        ` They remain in local storage.`,
    )
  }
  const result = (extra: Partial<PushResult> = {}): PushResult => ({
    pushed: true,
    truncatedRuleCount: truncatedRuleCount > 0 ? truncatedRuleCount : undefined,
    truncatedTombstoneCount: truncatedTombCount > 0 ? truncatedTombCount : undefined,
    ...extra,
  })

  const localBase = stripPerDeviceSettings(pushSettings) as Record<string, unknown>
  // S10 (2026-11-06): the cloud was last written by a v1.0.8 peer (its meta
  // has no semanticHash) and holds exactly this content minus what v1.0.8
  // drops when it re-pushes a merge (rows' `sr`, tombstone spellings) —
  // i.e. our own last push as that peer echoed it. Re-pushing changes
  // nothing for it and only provokes its next echo.
  const sameAsV108Echo =
    cloudMeta !== undefined &&
    cloudMeta.semanticHash === undefined &&
    cloudMeta.contentHash === payload.v108ContentHash
  // The cloud's own items must actually carry its meta's pushId before we
  // trust its hashes (A6, 2026-11-06): the owner of a meta whose chunks were
  // overwritten by another push used to skip its repair push as "unchanged".
  const comparable =
    !freshMarker &&
    !recovering &&
    cloudMeta !== undefined &&
    cloudMeta.schemaVersion === SYNC_SCHEMA_VERSION &&
    (cloudMeta.contentHash === payload.contentHash ||
      cloudMeta.semanticHash === payload.semanticHash ||
      sameAsV108Echo) &&
    (await snapshotIntact(cloudMeta))
  // A first-time participant that writes without having pulled (torn-cloud
  // recovery, S7) records the wipe marker its push carries forward as seen —
  // F1: never applied to the library it joined with.
  const carriedMarker = state.participating ? undefined : normalizeWipeMarker(cloudMeta?.wipeMarker)

  // The active-day ledger (A14) forces a push only when this machine knows a
  // day the cloud lacks: a day it merely LACKS (e.g. a far-future day from a
  // skewed peer that its decoder drops) must never make every follow-up push
  // differ — that would ping-pong. A meta from an older build has no ledger:
  // the days then ride along with the next real push.
  const daysNews = comparable && (await ledgerHasNews(cloudMeta!, payload.daysHash, activeDays))

  // No-op push: the cloud already holds exactly this content (typically the
  // echo of a pull). Skipping it is what makes deferring every local change
  // during a pull safe — echoes cost nothing and can't ping-pong.
  if (comparable && !daysNews && (cloudMeta!.contentHash === payload.contentHash || sameAsV108Echo)) {
    if (
      !state.participating ||
      state.lastMergedPushId !== cloudMeta!.pushId ||
      !state.settingsBase ||
      state.counterPushDueAt !== undefined ||
      !sameForeign(state.foreignSettings, outgoing.foreign)
    ) {
      await writeLocalSyncState({
        ...withoutCounterPush(state),
        participating: true,
        lastMergedPushId: cloudMeta!.pushId,
        settingsBase: localBase,
        ...foreignField(outgoing.foreign),
        ...(carriedMarker ? { lastSeenWipeId: carriedMarker.id } : {}),
      })
    }
    if (state.counterPushDueAt !== undefined || counterPushTimer) disarmCounterPush()
    await clearSyncError()
    resetPushRetry()
    return result({ unchanged: true })
  }

  // Counter-only change (2026-11-06): the cloud differs from local only in
  // hit counters — postpone the ~88 KB rewrite to one coalesced push per
  // SYNC_COUNTER_PUSH_COALESCE_MS. Anything semantic (rules, tombstones,
  // settings, folder activity) changes semanticHash and pushes now, as does
  // a new active day; so do 立即上傳 and every other pushNow (flush).
  if (comparable && !daysNews && !options.flush && cloudMeta!.semanticHash === payload.semanticHash) {
    const dueAt = state.counterPushDueAt ?? Date.now() + SYNC_COUNTER_PUSH_COALESCE_MS
    if (state.counterPushDueAt === undefined) await writeLocalSyncState({ ...state, counterPushDueAt: dueAt })
    armCounterPush(dueAt)
    await clearSyncError()
    resetPushRetry()
    return { pushed: false, deferred: true, reason: 'counter-only changes deferred' }
  }

  // Orphans: keys the previous cloud state occupied that this push doesn't
  // rewrite (13c: derived from the previous meta's counts; the first push of
  // a SW lifetime scans everything once to catch strays).
  const keep = new Set(Object.keys(payload.writes))
  let orphanCandidates: string[]
  let didFullScan = false
  if (!fullOrphanScanDone) {
    orphanCandidates = Object.keys(await chrome.storage.sync.get(null)).filter(isOurSyncKey)
    didFullScan = true
  } else {
    orphanCandidates = keysOfMeta(cloudMeta)
  }
  const orphanKeys = orphanCandidates.filter((k) => !keep.has(k))

  // Set before writing: the listener must recognise this push's own echo.
  lastPushedPushId = pushId
  try {
    // Audit G3: write the new meta + chunks FIRST, then remove orphans. A
    // reader only reads what the meta it sees references, so orphans are dead
    // data once the new meta lands; removing them first left an old meta
    // pointing at deleted chunks if this set failed.
    await chrome.storage.sync.set(payload.writes)
  } catch (e) {
    const errMsg = errorMessage(e)
    const kind = classifyStorageError(errMsg)
    if (kind === 'quota-bytes') {
      // Quota-wedge recovery (audit H1): the byte quota is enforced on the
      // post-set bucket TOTAL — new payload + still-stored orphans. Free every
      // key of ours this push doesn't write, then retry the set once.
      try {
        const all = Object.keys(await chrome.storage.sync.get(null)).filter(
          (k) => isOurSyncKey(k) && !keep.has(k),
        )
        if (all.length === 0) throw e
        console.warn(
          `[mail-organizer] sync push (${reason}): quota exceeded with ${all.length} stale key(s) — removing and retrying once`,
        )
        await chrome.storage.sync.remove(all)
        await chrome.storage.sync.set(payload.writes)
      } catch (e2) {
        const retryMsg = errorMessage(e2)
        console.warn(`[mail-organizer] sync push (${reason}): quota recovery failed`, e2)
        await recordSyncError({ source: 'push', reason: retryMsg })
        await logError('sync:push', retryMsg, { reason })
        fullOrphanScanDone = false
        schedulePushRetry(classifyStorageError(retryMsg))
        return { pushed: false, reason: retryMsg }
      }
    } else {
      // Rate limits back off on their own schedule; network etc. retry with
      // exponential backoff. Don't touch lastSyncAt.
      console.warn(`[mail-organizer] sync push (${reason}): chrome.storage.sync.set failed`, e)
      await recordSyncError({ source: 'push', reason: errMsg })
      await logError('sync:push', errMsg, { reason, kind })
      schedulePushRetry(kind)
      return { pushed: false, reason: errMsg }
    }
  }

  // Orphan cleanup AFTER the authoritative write — non-fatal. If it fails,
  // cloud stays valid (unreferenced keys are never read); the next SW
  // lifetime's full scan removes them.
  if (orphanKeys.length > 0) {
    try {
      await chrome.storage.sync.remove(orphanKeys)
      if (didFullScan) fullOrphanScanDone = true
    } catch (e) {
      fullOrphanScanDone = false
      console.warn('[mail-organizer] orphan chunk cleanup failed (non-fatal)', e)
    }
  } else if (didFullScan) {
    fullOrphanScanDone = true
  }

  const seenMarker = freshMarker ?? carriedMarker
  await writeLocalSyncState({
    ...withoutCounterPush(state),
    participating: true,
    lastMergedPushId: pushId,
    settingsBase: localBase,
    ...foreignField(outgoing.foreign),
    ...(seenMarker ? { lastSeenWipeId: seenMarker.id } : {}),
  })
  if (state.counterPushDueAt !== undefined || counterPushTimer) disarmCounterPush()
  if (recovering) await chrome.storage.local.remove(SYNC_TORN_STREAK_KEY)
  await setSettings({ lastSyncAt: updatedAt })
  await clearSyncError()
  resetPushRetry()
  return result()
}

/** `state` minus a pending counter push (a push just carried the counters). */
function withoutCounterPush(state: SyncLocalState): SyncLocalState {
  const { counterPushDueAt: _drop, ...rest } = state
  void _drop
  return rest
}

type ForeignSettings = NonNullable<SyncLocalState['foreignSettings']>

/** `{ foreignSettings }` when there is any, else an explicit removal. */
function foreignField(f: ForeignSettings | undefined): { foreignSettings?: ForeignSettings } {
  return f && Object.keys(f).length > 0 ? { foreignSettings: f } : { foreignSettings: undefined }
}

function sameForeign(a: ForeignSettings | undefined, b: ForeignSettings | undefined): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {})
}

/**
 * Can this build hold the synced value `v` of setting `k` as is? (A13,
 * 2026-11-06) sanitizeSettings maps an unknown claudeModel — e.g. one a newer
 * peer offers — to the default, so adopting it would really adopt the
 * default, and the next push would bounce that default back to the peer.
 */
function isAdoptableSetting(k: string, v: unknown): boolean {
  return k !== 'claudeModel' || isKnownClaudeModel(v)
}

/**
 * The settings a push sends: local values, except a synced value this build
 * couldn't adopt (foreignSettings) goes back out unchanged while the local
 * value is still the one it had then. A local change since drops the record —
 * the user's own choice is pushed instead.
 */
function applyForeignSettings(
  local: Settings,
  foreign: ForeignSettings | undefined,
): { settings: Partial<Settings>; foreign?: ForeignSettings } {
  if (!foreign || Object.keys(foreign).length === 0) return { settings: local }
  const out: Record<string, unknown> = { ...local }
  const kept: ForeignSettings = {}
  for (const [k, f] of Object.entries(foreign)) {
    if (!f || typeof f !== 'object') continue
    if (JSON.stringify(out[k]) === JSON.stringify(f.local)) {
      out[k] = f.cloud
      kept[k] = f
    }
  }
  return { settings: out as Partial<Settings>, ...(Object.keys(kept).length > 0 ? { foreign: kept } : {}) }
}

/**
 * Does every item `meta` references carry its pushId? (A6.) Legacy metas
 * can't be checked (no pushId) and count as intact; so does a meta that
 * changed meanwhile — the change brings its own pull.
 */
async function snapshotIntact(meta: SyncMeta): Promise<boolean> {
  if (meta.schemaVersion < 4 || !meta.v4 || typeof meta.pushId !== 'string' || !meta.pushId) return true
  const keys = keysOfMeta(meta)
  const stored = await chrome.storage.sync.get([SYNC_META_KEY, ...keys])
  const fresh = stored[SYNC_META_KEY] as SyncMeta | undefined
  if (!fresh || fresh.pushId !== meta.pushId) return true
  return keys.every((k) => {
    const item = stored[k] as { p?: unknown } | undefined
    return !!item && typeof item === 'object' && item.p === meta.pushId
  })
}

/**
 * Does this machine know an active day the cloud's ledger lacks? (A14)
 * Free when the fingerprints agree; otherwise reads the cloud's settings
 * item. An older writer (no daysHash) never forces a push for the ledger.
 */
async function ledgerHasNews(meta: SyncMeta, daysHash: string, localDays: readonly string[]): Promise<boolean> {
  if (meta.daysHash === undefined || meta.daysHash === daysHash || localDays.length === 0) return false
  const item = (await chrome.storage.sync.get(V4_SETTINGS_KEY))[V4_SETTINGS_KEY] as
    | { p?: unknown; d?: unknown }
    | undefined
  if (!item || item.p !== meta.pushId) return false // changed meanwhile — that change brings a pull
  const cloudDays = new Set(decodeActiveDays(item.d, Number.POSITIVE_INFINITY))
  const ours = decodeActiveDays(encodeActiveDays(localDays, ACTIVE_DAYS_CAP), Number.POSITIVE_INFINITY)
  return ours.some((d) => !cloudDays.has(d))
}

async function readTornStreak(): Promise<TornStreak | undefined> {
  const r = await chrome.storage.local.get(SYNC_TORN_STREAK_KEY)
  const s = r[SYNC_TORN_STREAK_KEY] as Partial<TornStreak> | undefined
  return s && typeof s.pushId === 'string' && typeof s.count === 'number' && typeof s.firstAt === 'number'
    ? (s as TornStreak)
    : undefined
}

/** Count one more torn read of `pushId` (a different pushId restarts the streak). */
async function noteTornRead(pushId: string): Promise<void> {
  const cur = await readTornStreak()
  const next: TornStreak =
    cur && cur.pushId === pushId
      ? { ...cur, count: cur.count + 1 }
      : { pushId, count: 1, firstAt: Date.now() }
  await chrome.storage.local.set({ [SYNC_TORN_STREAK_KEY]: next })
}

/**
 * May a push rewrite the torn cloud? Only when the CURRENT meta's pushId has
 * read torn at least SYNC_TORN_RECOVERY_MIN_READS times over at least
 * SYNC_TORN_RECOVERY_MIN_MS — a push still being delivered heals in seconds —
 * and the meta carries no wipe this machine still has to act on (the
 * rewrite would re-upload what that wipe removes).
 */
async function tornRecoveryAllowed(machineId: string): Promise<boolean> {
  const meta = await readSyncMeta()
  if (!meta || typeof meta.pushId !== 'string' || !meta.pushId) return false
  const streak = await readTornStreak()
  if (!streak || streak.pushId !== meta.pushId) return false
  if (streak.count < SYNC_TORN_RECOVERY_MIN_READS) return false
  if (Date.now() - streak.firstAt < SYNC_TORN_RECOVERY_MIN_MS) return false
  const { state } = await readLocalSyncState()
  const marker = normalizeWipeMarker(meta.wipeMarker)
  if (marker && marker.id !== state.lastSeenWipeId && marker.byMachineId !== machineId && state.participating) {
    return false
  }
  return true
}

/**
 * Adoption without a cloud to merge (push path): the first-participation
 * protections of mergeCloudIntoLocal against this machine's own tombstones,
 * after the 'upgrade' backup for a machine that synced under v1.0.7.
 */
async function adoptLocalLibrary(state: SyncLocalState): Promise<void> {
  if (state.legacyLastSyncAt !== undefined) {
    try {
      await appendUpgradeBackupOnce('pre-push')
    } catch (e) {
      console.warn('[mail-organizer] upgrade backup failed (sync continues)', e)
    }
  }
  await mutateRules(
    async (existing) => {
      const res = mergeCloudIntoLocal({
        local: existing,
        localTombstones: await getRuleTombstones(),
        cloudRules: [],
        cloudStubs: [],
        cloudTombstones: [],
        mode: 'union',
        adoption: { processedBeforeMs: parseTime(state.legacyLastSyncAt) },
      })
      return { next: res.next, result: undefined }
    },
    { preserveEditedAt: true },
  )
}

// ---- pull ------------------------------------------------------------------------

/**
 * Pull merge mode:
 *   - 'replace': drop local syncable rules the cloud doesn't represent at
 *     all. ONLY for explicit callers that want cloud-exact state — the push
 *     budget deliberately leaves low-priority rules out of the cloud.
 *   - 'union': keep local rules the cloud doesn't have. Deletions still
 *     propagate via tombstones. The default for remote-change pulls, the
 *     pull before a push, first-enable and 立即下載.
 */
type PullMergeMode = 'replace' | 'union'

async function doPull(
  reason: string,
  opts: { mode: PullMergeMode; source: 'pull' | 'pull-remote'; backup?: boolean },
): Promise<PullResult> {
  pullInProgress = true
  if (pullGraceTimer) {
    clearTimeout(pullGraceTimer)
    pullGraceTimer = null
  }
  try {
    return await doPullInner(opts)
  } catch (e) {
    const msg = errorMessage(e)
    console.warn(`[mail-organizer] sync pull (${reason}) failed`, e)
    await recordSyncError({ source: opts.source, reason: msg }).catch(() => {})
    await logError(opts.source === 'pull-remote' ? 'sync:pull-remote' : 'sync:pull', msg, {
      reason,
    }).catch(() => {})
    return { pulled: false, reason: msg }
  } finally {
    // Bug #G: don't clear pullInProgress synchronously — onChanged listeners
    // for this pull's writes fire as macrotasks after this finally. When the
    // window closes, push whatever local change arrived during it.
    pullGraceTimer = setTimeout(() => {
      pullInProgress = false
      pullGraceTimer = null
      if (deferredLocalChange) {
        deferredLocalChange = false
        schedulePush('deferred-after-pull')
      }
    }, PULL_GRACE_MS)
  }
}

type CloudSnapshot = {
  rules: Rule[]
  stubs: DecodedStub[]
  tombstones: RuleTombstone[]
  activity: FolderActivity[]
  settings?: Record<string, unknown>
  /** The pusher's active-day ledger (v4 settings item `d`, 2026-11-06); [] when absent. */
  activeDays: string[]
  /** Rows dropped by the decoder, or chunks missing from a partial legacy read. */
  imperfect: boolean
}

type SnapshotRead = { ok: true; value: CloudSnapshot } | { ok: false; reason: string }

function sanitizeActivity(items: unknown[]): FolderActivity[] {
  return items.filter(
    (e): e is FolderActivity =>
      !!e &&
      typeof e === 'object' &&
      typeof (e as FolderActivity).folderId === 'string' &&
      (e as FolderActivity).folderId !== '' &&
      typeof (e as FolderActivity).folderPath === 'string' &&
      typeof (e as FolderActivity).lastActiveAt === 'string',
  )
}

/**
 * Read meta + every item it references in ONE storage read, and verify they
 * belong together (item 4). v4: every chunk and the settings item must carry
 * meta.pushId — otherwise abort without writing anything. Legacy (v1–v3,
 * no pushId): every referenced chunk must be present; after repeated misses
 * a partial read is accepted, forced to union (it can then only ADD).
 */
async function readCloudSnapshot(meta: SyncMeta, allowPartialLegacy: boolean): Promise<SnapshotRead> {
  const v = meta.schemaVersion
  if (v >= 4) {
    const c = meta.v4
    const pushId = meta.pushId
    if (!c || typeof c !== 'object' || typeof pushId !== 'string' || !pushId) {
      return { ok: false, reason: 'cloud snapshot incomplete: malformed v4 meta' }
    }
    const rowKeys = rangeKeys(V4_ROWS_PREFIX, c.rows)
    const folderKeys = rangeKeys(V4_FOLDERS_PREFIX, c.folders)
    const tombKeys = rangeKeys(V4_TOMBS_PREFIX, c.tombs)
    const activityKeys = rangeKeys(V4_ACTIVITY_PREFIX, c.activity)
    const stored = await chrome.storage.sync.get([
      SYNC_META_KEY,
      V4_SETTINGS_KEY,
      ...rowKeys,
      ...folderKeys,
      ...tombKeys,
      ...activityKeys,
    ])
    const fresh = stored[SYNC_META_KEY] as SyncMeta | undefined
    if (!fresh || fresh.pushId !== pushId) {
      return { ok: false, reason: 'cloud snapshot incomplete: cloud changed during read' }
    }
    const chunkOf = <T,>(k: string): SyncChunk<T> | null => {
      const ch = stored[k] as SyncChunk<T> | undefined
      return ch && ch.p === pushId && Array.isArray(ch.items) ? ch : null
    }
    const collect = <T,>(keys: string[]): SyncChunk<T>[] | string => {
      const out: SyncChunk<T>[] = []
      for (const k of keys) {
        const ch = chunkOf<T>(k)
        if (!ch) return k
        out.push(ch)
      }
      return out
    }
    const rows = collect<unknown>(rowKeys)
    const folders = collect<unknown>(folderKeys)
    const tombs = collect<unknown>(tombKeys)
    const activity = collect<unknown>(activityKeys)
    const settingsItem = stored[V4_SETTINGS_KEY] as { p?: unknown; v?: unknown; d?: unknown } | undefined
    for (const part of [rows, folders, tombs, activity]) {
      if (typeof part === 'string') {
        return { ok: false, reason: `cloud snapshot incomplete: ${part} missing or from another push` }
      }
    }
    if (!settingsItem || settingsItem.p !== pushId) {
      return { ok: false, reason: `cloud snapshot incomplete: ${V4_SETTINGS_KEY} missing or from another push` }
    }
    const decoded = decodeRulePayload({
      folders: unchunkArray(folders as SyncChunk<unknown>[]),
      rows: unchunkArray(rows as SyncChunk<unknown>[]),
    })
    const tombstones: RuleTombstone[] = []
    let badTombs = 0
    for (const row of unchunkArray(tombs as SyncChunk<unknown>[])) {
      const t = decodeTombstone(row)
      if (t) tombstones.push(t)
      else badTombs++
    }
    return {
      ok: true,
      value: {
        rules: decoded.rules,
        stubs: decoded.stubs,
        tombstones,
        activity: sanitizeActivity(unchunkArray(activity as SyncChunk<unknown>[])),
        settings:
          settingsItem.v && typeof settingsItem.v === 'object'
            ? (settingsItem.v as Record<string, unknown>)
            : undefined,
        activeDays: decodeActiveDays(settingsItem.d),
        imperfect: decoded.skipped > 0 || badTombs > 0,
      },
    }
  }

  // v1–v3.
  const legacyRuleKeys = v <= 2 ? rangeKeys(LEGACY_RULES_PREFIX, meta.ruleChunkCount) : []
  const rowKeys = v === 3 ? rangeKeys(V3_ROWS_PREFIX, meta.ruleV3ChunkCount) : []
  const folderKeys = v === 3 ? rangeKeys(V3_FOLDERS_PREFIX, meta.folderChunkCount) : []
  const tombKeys = rangeKeys(LEGACY_TOMBS_PREFIX, meta.tombstoneChunkCount)
  const activityKeys = v >= 2 ? rangeKeys(LEGACY_ACTIVITY_PREFIX, meta.folderActivityChunkCount) : []
  const allKeys = [...legacyRuleKeys, ...rowKeys, ...folderKeys, ...tombKeys, ...activityKeys]
  const stored = await chrome.storage.sync.get([SYNC_META_KEY, LEGACY_SETTINGS_KEY, ...allKeys])
  const fresh = stored[SYNC_META_KEY] as SyncMeta | undefined
  if (
    !fresh ||
    fresh.updatedAt !== meta.updatedAt ||
    fresh.sourceMachineId !== meta.sourceMachineId ||
    fresh.schemaVersion !== meta.schemaVersion
  ) {
    return { ok: false, reason: 'cloud snapshot incomplete: cloud changed during read' }
  }
  const missing = allKeys.filter((k) => {
    const ch = stored[k] as SyncChunk<unknown> | undefined
    return !ch || !Array.isArray(ch.items)
  })
  if (missing.length > 0 && !allowPartialLegacy) {
    return { ok: false, reason: `cloud snapshot incomplete: ${missing[0]} missing` }
  }
  const chunks = (keys: string[]) =>
    keys.map((k) => stored[k] as SyncChunk<unknown> | undefined).filter((c): c is SyncChunk<unknown> => !!c)
  const rules: Rule[] = []
  let skipped = 0
  for (const raw of unchunkArray(chunks(legacyRuleKeys))) {
    const r = sanitizeLegacyRule(raw)
    if (r) rules.push(r)
    else skipped++
  }
  let stubs: DecodedStub[] = []
  if (v === 3) {
    const decoded = decodeRulePayload({
      folders: unchunkArray(chunks(folderKeys)),
      rows: unchunkArray(chunks(rowKeys)),
    })
    rules.push(...decoded.rules)
    stubs = decoded.stubs
    skipped += decoded.skipped
  }
  const tombstones: RuleTombstone[] = []
  for (const raw of unchunkArray(chunks(tombKeys))) {
    const t = sanitizeLegacyTombstone(raw)
    if (t) tombstones.push(t)
    else skipped++
  }
  const settingsRaw = stored[LEGACY_SETTINGS_KEY]
  return {
    ok: true,
    value: {
      rules,
      stubs,
      tombstones,
      activity: sanitizeActivity(unchunkArray(chunks(activityKeys))),
      settings:
        settingsRaw && typeof settingsRaw === 'object' ? (settingsRaw as Record<string, unknown>) : undefined,
      activeDays: [],
      imperfect: skipped > 0 || missing.length > 0,
    },
  }
}

async function doPullInner(opts: {
  mode: PullMergeMode
  source: 'pull' | 'pull-remote'
  backup?: boolean
}): Promise<PullResult> {
  const settings = await getSettings()
  if (!settings.syncEnabled) {
    // Caller flipped sync off between our entry and the awaited read.
    return { pulled: false, reason: 'sync disabled' }
  }
  let meta = await readSyncMeta()
  if (!meta) return { pulled: false, reason: 'no cloud state' }
  const refuse = async (reason: string): Promise<PullResult> => {
    // Item 7b: refusals are recorded on EVERY path (the remote-change path
    // used to swallow them, so a schema-refused machine looked healthy).
    await recordSyncError({ source: opts.source, reason })
    return { pulled: false, reason }
  }
  const refusal = schemaRefusal(meta)
  if (refusal) return refuse(refusal)

  // Item 4: verified snapshot; one immediate re-read if a new push landed
  // between the two reads, otherwise back off and retry.
  let snap: SnapshotRead = { ok: false, reason: 'not read' }
  for (let attempt = 0; attempt < 2; attempt++) {
    snap = await readCloudSnapshot(meta, tornAttempts >= SYNC_TORN_RETRY_MAX_ATTEMPTS)
    if (snap.ok) break
    const again = await readSyncMeta()
    if (!again) return { pulled: false, reason: 'no cloud state' }
    const againRefusal = schemaRefusal(again)
    if (againRefusal) return refuse(againRefusal)
    meta = again
  }
  if (!snap.ok) {
    tornAttempts++
    // A6: the streak survives SW restarts — see tornRecoveryAllowed.
    if (typeof meta.pushId === 'string' && meta.pushId) {
      await noteTornRead(meta.pushId).catch(() => {})
    }
    if (tornAttempts >= SYNC_TORN_RETRY_MAX_ATTEMPTS) {
      await recordSyncError({ source: opts.source, reason: snap.reason })
    }
    scheduleTornRetry()
    return { pulled: false, reason: snap.reason, torn: true }
  }
  tornAttempts = 0
  if (tornRetryTimer) {
    clearTimeout(tornRetryTimer)
    tornRetryTimer = null
  }
  if (await readTornStreak()) await chrome.storage.local.remove(SYNC_TORN_STREAK_KEY)
  const cloud = snap.value
  // An imperfect snapshot (skipped rows, missing legacy chunks) may not
  // delete anything a full one would have kept.
  const mode: PullMergeMode = cloud.imperfect ? 'union' : opts.mode

  const { state, legacyLastSyncAt } = await readLocalSyncState()
  const nextState: SyncLocalState = { ...state }
  // First merge since sync was (re-)enabled — or the first v4 merge of a
  // machine upgraded from v1.0.7 (A1, 2026-11-06).
  const adopting = !isAdopted(state)

  // Item 6: a wipe marker is acted on at most once, by id.
  //   - our own marker → record only (echo guard)
  //   - first participation → record only (F1)
  //   - otherwise → apply. The first v4 run of a machine that synced under
  //     v1.0.7 decides its first marker with the old lastSyncAt rule, once.
  const marker = normalizeWipeMarker(meta.wipeMarker)
  let applyWipe = false
  if (marker && marker.id !== state.lastSeenWipeId) {
    if (marker.byMachineId !== settings.syncMachineId && state.participating) {
      applyWipe = legacyLastSyncAt !== undefined ? marker.at > legacyLastSyncAt : true
    }
    nextState.lastSeenWipeId = marker.id
  }

  // Item 8: back up only before merges that can destroy local state. A1/A4
  // (2026-11-06): the first v4 merge of a machine that synced under v1.0.7
  // is one of them — its rules carry no editedAt (ties go to the cloud) and
  // there is no settings base (every cloud value wins) — so it gets a pinned
  // 「升級同步格式前」 snapshot, tombstones included; so does a first-enable
  // merge (the other first contact).
  let backupError: string | undefined
  const upgradeBackup = adopting && legacyLastSyncAt !== undefined
  if (applyWipe || mode === 'replace' || opts.backup || upgradeBackup) {
    try {
      if (upgradeBackup) await appendUpgradeBackupOnce('pre-pull')
      if (applyWipe) await appendBackup('pre-pull', 'remote-wipe', { pinned: true })
      else if (mode === 'replace') await appendBackup('pre-pull', 'replace-pull')
      else if (opts.backup && !upgradeBackup) {
        await appendBackup('pre-pull', 'first-enable', { withTombstones: true })
      }
    } catch (e) {
      backupError = errorMessage(e)
      console.warn('[mail-organizer] sync backup failed (sync continues)', e)
    }
  }

  let droppedRuleCount = 0
  if (applyWipe && marker) {
    await mutateRules(
      (existing) => {
        const next = existing.filter((r) => !shouldSyncRule(r))
        droppedRuleCount = existing.length - next.length
        return { next: droppedRuleCount > 0 ? next : existing, result: undefined }
      },
      { preserveEditedAt: true },
    )
    // F3: local tombstones are NOT cleared on a remote wipe — they are this
    // machine's own deletion intents. F2: record the marker IMMEDIATELY after
    // the destructive step commits, before the merge (which can throw) — a
    // throw must not leave it "unseen" and re-trip on the next pull.
    await writeLocalSyncState(nextState)
    try {
      const notice: RemoteWipeNotice = {
        byMachineId: marker.byMachineId,
        at: marker.at,
        appliedAt: new Date().toISOString(),
        droppedRuleCount,
      }
      await chrome.storage.local.set({ [REMOTE_WIPE_NOTICE_KEY]: notice })
    } catch (e) {
      console.warn('[mail-organizer] persisting remoteWipeNotice failed (non-fatal)', e)
    }
    console.warn(
      `[mail-organizer] remote wipe applied: ${droppedRuleCount} syncable rules dropped (issued by machine ${marker.byMachineId} at ${marker.at})`,
    )
  } else if (nextState.lastSeenWipeId !== state.lastSeenWipeId) {
    await writeLocalSyncState(nextState)
  }

  // Merge (items 1–3) under the rules lock; tombstones are read and written
  // inside it too, so a concurrent deleteRule (which tombstones under the
  // same lock) can't interleave.
  const cappedCloud = cloud.rules.map(applyConfidenceCap)
  await mutateRules(
    async (existing) => {
      const res = mergeCloudIntoLocal({
        local: existing,
        localTombstones: await getRuleTombstones(),
        cloudRules: cappedCloud,
        cloudStubs: cloud.stubs,
        cloudTombstones: cloud.tombstones,
        mode,
        ...(adopting ? { adoption: { processedBeforeMs: parseTime(legacyLastSyncAt) } } : {}),
      })
      if (res.newTombstones.length > 0) await addRuleTombstones(res.newTombstones)
      return { next: res.next, result: undefined }
    },
    { preserveEditedAt: true },
  )

  // folderActivity: union by folderId, most-recent lastActiveAt wins.
  if (cloud.activity.length > 0) {
    await mergeFolderActivityFromSync(cloud.activity)
  }
  // A14 (2026-11-06): the stale-sweep clock is the union of every machine's
  // active days, so the no-tombstone stale / dormant deletes converge.
  if (cloud.activeDays.length > 0) {
    await mergeActiveDaysFromSync(cloud.activeDays)
  }

  // Settings overlay through setSettings (item 14) — per-device fields
  // stripped (defence-in-depth: the push side strips them too) — plus our
  // own lastSyncAt, in one write. Three-way against the last-synced base: a
  // field edited HERE since the last sync is a pending local edit (its push
  // follows) and is kept; every other field takes the cloud's value. On the
  // first merge there is no base: every cloud value wins (A4 — the
  // 'upgrade' / 'first-enable' backup above holds the local ones).
  // A13: a value this build can't represent (an unknown claudeModel) is
  // neither adopted — sanitize would turn it into the default — nor later
  // pushed over; foreignSettings sends it back unchanged.
  const incoming = cloud.settings
    ? (stripPerDeviceSettings(cloud.settings as Partial<Settings>) as Record<string, unknown>)
    : undefined
  const overlay: Record<string, unknown> = {}
  let settingsBase: Record<string, unknown> | undefined
  const foreign: ForeignSettings = { ...(state.foreignSettings ?? {}) }
  if (incoming) {
    const local = stripPerDeviceSettings(await getSettings()) as Record<string, unknown>
    const base = state.settingsBase
    settingsBase = { ...incoming }
    for (const [k, v] of Object.entries(incoming)) {
      const editedHere = base !== undefined && JSON.stringify(local[k]) !== JSON.stringify(base[k])
      if (!isAdoptableSetting(k, v)) {
        if (editedHere) {
          // A local edit is waiting to be pushed — it goes out instead.
          delete foreign[k]
          settingsBase[k] = base?.[k]
        } else {
          foreign[k] = { cloud: v, local: local[k] }
          settingsBase[k] = local[k]
        }
        continue
      }
      delete foreign[k]
      if (!editedHere) overlay[k] = v
    }
  }
  await setSettings({ ...(overlay as Partial<Settings>), lastSyncAt: new Date().toISOString() })

  await writeLocalSyncState({
    ...(adopting ? adoptedState(nextState) : nextState),
    participating: true,
    lastMergedPushId: meta.pushId,
    ...(settingsBase ? { settingsBase } : {}),
    ...(incoming ? foreignField(foreign) : {}),
  })
  if (backupError) {
    await recordSyncError({
      source: opts.source,
      reason: `backup failed before merge (sync continued): ${backupError}`,
    })
  } else {
    await clearSyncError()
  }
  return {
    pulled: true,
    ruleCount: cloud.rules.length,
    tombstoneCount: cloud.tombstones.length,
    mergeMode: mode,
  }
}

// ---- merge (pure) --------------------------------------------------------------

// Tombstone keys come from rules.ts `tombstoneKey` (imported above), which
// re-canonicalizes the stored signalNorm (2026-11-05): normalizeSignal now
// folds court-case forms / full-width digits, so a tombstone written by an
// older build in pre-canonical form must still match ruleKey() below.

function ruleKey(r: Pick<Rule, 'type' | 'signal' | 'targetFolderPath'>): string {
  return `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
}

function newestDeletionByKey(tombstones: RuleTombstone[]): Map<string, number> {
  const out = new Map<string, number>()
  for (const t of tombstones) {
    const k = tombstoneKey(t)
    const cur = out.get(k)
    if (cur === undefined || t.deletedAt > cur) out.set(k, t.deletedAt)
  }
  return out
}

function parseTime(s: string | undefined): number | undefined {
  if (typeof s !== 'string') return undefined
  const t = Date.parse(s)
  return Number.isFinite(t) ? t : undefined
}

/** When this version of the rule came to be: max(createdAt, editedAt). */
function ruleTimeMs(r: Rule): number {
  return Math.max(parseTime(r.createdAt) ?? 0, r.editedAt ?? 0)
}

function earlierIso(a: string, b: string): string {
  const ta = parseTime(a)
  const tb = parseTime(b)
  if (ta === undefined) return tb === undefined ? a || b : b
  if (tb === undefined) return a
  return tb < ta ? b : a
}

function laterIso(a: string | undefined, b: string | undefined): string | undefined {
  if (a === undefined) return b
  if (b === undefined) return a
  const ta = parseTime(a)
  const tb = parseTime(b)
  if (ta === undefined) return tb === undefined ? a : b
  if (tb === undefined) return a
  return tb > ta ? b : a
}

const RULE_FIELDS: ReadonlyArray<keyof Rule> = [
  'id',
  'type',
  'signal',
  'targetFolderId',
  'targetFolderPath',
  'confidence',
  'matchCount',
  'overrideCount',
  'enabled',
  'createdAt',
  'lastUsedAt',
  'source',
  'orphaned',
  'autoDisabledAt',
  'autoDisabledReason',
  'editedAt',
  'statsResetAt',
  'orphanedAt',
]

function rulesEqual(a: Rule, b: Rule): boolean {
  for (const f of RULE_FIELDS) {
    const x = a[f]
    const y = b[f]
    if (f === 'orphaned') {
      if (!!x !== !!y) return false
    } else if (f === 'overrideCount') {
      if ((x ?? 0) !== (y ?? 0)) return false
    } else if ((x ?? undefined) !== (y ?? undefined)) {
      return false
    }
  }
  return true
}

/**
 * Merge two versions of one rule (same id). Semantic fields come from the
 * newer editedAt (tie → cloud); lastUsedAt takes the latest, createdAt the
 * earliest. On a tie an auto rule keeps the higher confidence: auto-promotion
 * (bumpRuleHits) deliberately doesn't stamp editedAt, so a promotion seen by
 * only one side survives the merge.
 *
 * 2026-11-06:
 *   - matchCount / overrideCount follow the counter epoch (statsResetAt): the
 *     copy with the newer epoch supplies both, equal epochs take the max. Max
 *     alone brought a revived rule's pre-reset counters back from the other
 *     machine, and the next sweep disabled it again (A3).
 *   - `orphaned` is per-device (A2): it stays orphaned only while the merged
 *     target is that same dead one — adopting another target clears it so
 *     reconcile re-checks it. A local orphaned copy WITHOUT orphanedAt may
 *     carry a v1.0.8 orphan stamp in editedAt, so it never supplies the
 *     semantic fields (orphanStampSuspect). S3: one orphaned under this build
 *     (orphanedAt recorded; orphaning no longer stamps) merges by plain
 *     editedAt — forcing the cloud copy reverted a genuine local edit (e.g.
 *     the user disabling the orphaned rule) while keeping its newer stamp,
 *     and the un-orphan later published the reverted content under it.
 */
function mergeRuleVersions(local: Rule, cloud: Rule): Rule {
  const le = local.editedAt ?? 0
  const ce = cloud.editedAt ?? 0
  const base = orphanStampSuspect(local) || ce >= le ? cloud : local
  const merged: Rule = { ...base }
  if (ce === le && base.source !== 'user_manual') {
    merged.confidence = Math.max(local.confidence, cloud.confidence)
  }
  assignCounters(merged, local, cloud)
  const ea = Math.max(le, ce)
  if (ea > 0) merged.editedAt = ea
  else delete merged.editedAt
  if (
    local.orphaned &&
    merged.targetFolderId === local.targetFolderId &&
    merged.targetFolderPath === local.targetFolderPath
  ) {
    merged.orphaned = true
    if (local.orphanedAt !== undefined) merged.orphanedAt = local.orphanedAt
    else delete merged.orphanedAt
  } else {
    delete merged.orphaned
    delete merged.orphanedAt
  }
  const capped = applyConfidenceCap(merged)
  return rulesEqual(capped, local) ? local : capped
}

/**
 * Non-semantic fields of one rule merged from two copies of it, onto `out`
 * (sync merge; restore, S2): matchCount / overrideCount by counter epoch
 * (statsResetAt — the newer epoch supplies both, equal epochs take the max),
 * lastUsedAt the latest, createdAt the earliest.
 */
function assignCounters(out: Rule, a: Rule, b: Rule): void {
  const as = a.statsResetAt ?? 0
  const bs = b.statsResetAt ?? 0
  const counters = as === bs ? undefined : as > bs ? a : b
  out.matchCount = counters ? counters.matchCount || 0 : Math.max(a.matchCount || 0, b.matchCount || 0)
  const oc = counters
    ? (counters.overrideCount ?? 0)
    : Math.max(a.overrideCount ?? 0, b.overrideCount ?? 0)
  if (oc > 0) out.overrideCount = oc
  else delete out.overrideCount
  const sr = Math.max(as, bs)
  if (sr > 0) out.statsResetAt = sr
  else delete out.statsResetAt
  out.createdAt = earlierIso(a.createdAt, b.createdAt)
  const lu = laterIso(a.lastUsedAt, b.lastUsedAt)
  if (lu !== undefined) out.lastUsedAt = lu
  else delete out.lastUsedAt
}

/**
 * An orphaned local copy whose editedAt may be a v1.0.8 orphan stamp: it was
 * orphaned before this build recorded orphanedAt, and not edited since (S3).
 */
function orphanStampSuspect(r: Rule): boolean {
  return !!r.orphaned && r.orphanedAt === undefined
}

/**
 * Apply a disabled stub to the local copy (newer-or-equal stub wins; an
 * orphan-stamp suspect loses to it like it loses to a full row — S3).
 */
function applyStub(local: Rule, stub: DecodedStub): Rule {
  const le = local.editedAt ?? 0
  const se = stub.editedAt ?? 0
  if (se < le && !orphanStampSuspect(local)) return local // the local edit (e.g. a re-enable) is newer
  const merged: Rule = { ...local, enabled: false }
  if (stub.autoDisabledReason) merged.autoDisabledReason = stub.autoDisabledReason
  else delete merged.autoDisabledReason
  if (stub.autoDisabledAt) merged.autoDisabledAt = stub.autoDisabledAt
  else delete merged.autoDisabledAt
  const ea = Math.max(le, se)
  if (ea > 0) merged.editedAt = ea
  return rulesEqual(merged, local) ? local : merged
}

/**
 * First-participation merge ("adoption", 2026-11-06 — A1). Present on the
 * first merge since sync was (re-)enabled, which for a machine that synced
 * under v1.0.7 is its first v4 merge.
 */
export type MergeAdoption = {
  /**
   * The v1.0.7 machine's lastSyncAt (ms): cloud tombstones deleted at or
   * before it were already processed by v3. Undefined for a machine that
   * never synced before (only its own tombstones count as processed then).
   */
  processedBeforeMs?: number
}

export type MergeInput = {
  local: Rule[]
  localTombstones: RuleTombstone[]
  cloudRules: Rule[]
  cloudStubs: DecodedStub[]
  cloudTombstones: RuleTombstone[]
  mode: PullMergeMode
  adoption?: MergeAdoption
}

export type MergeOutput = {
  /** === input.local when nothing changed (so mutateRules skips the write). */
  next: Rule[]
  /** Cloud tombstones newer than the local one for their triple (to persist). */
  newTombstones: RuleTombstone[]
  stats: {
    added: number
    updated: number
    dropped: number
    suppressedIncoming: number
    /** Local rules adoption kept (and re-stamped) despite a suppressing tombstone. */
    protected: number
  }
}

const SYNCABLE_SOURCES: ReadonlySet<RuleSource> = new Set<RuleSource>([
  'user_manual',
  'ai_confirmed',
  'ai_overridden',
])

/**
 * Fingerprint of the spelling (legacySignalNorm form) of the rule a tombstone
 * was written for, when it can be known (S1, 2026-11-06):
 *   - the spellingHash this build records on every tombstone it writes;
 *   - a tombstone deleted before the canonical keys shipped
 *     (CANONICAL_TOMBSTONE_KEYS_SINCE_MS) was written by v1.0.7 or older, in
 *     the deleted rule's own spelling: its signalNorm IS that spelling.
 * Otherwise (v1.0.8 stored the canonical form and nothing else) it is
 * unknown — and an unknown spelling never counts as a collision: v1.0.8
 * wrote 「112訴304」 for the user's own deleted 「112年度訴字第304號」.
 */
function deletedSpellingHash(t: RuleTombstone): string | undefined {
  if (typeof t.spellingHash === 'string') return t.spellingHash
  if (t.deletedAt < CANONICAL_TOMBSTONE_KEYS_SINCE_MS) return spellingFingerprint(t.signalNorm)
  return undefined
}

/** One known deletion as the merge applies it; `own` = this machine's local tombstone. */
type Deletion = { t: RuleTombstone; own: boolean }

/** Same deletion, whatever slot it sits in (spelling stripped or not). */
function echoKey(t: RuleTombstone): string {
  return `${t.type}::${t.signalNorm}::${t.targetFolderPath}@${t.deletedAt}`
}

/**
 * Every known deletion per canonical triple, local ∪ cloud (round 3,
 * 2026-11-06). A spelling-less entry with the same stored triple and
 * deletedAt as a spelled one is that deletion re-pushed by v1.0.8 (which
 * drops spellings) and is left out — see storage.ts mergeTombstoneSlots.
 */
function deletionsByKey(
  local: readonly RuleTombstone[],
  cloud: readonly RuleTombstone[],
): { byKey: Map<string, Deletion[]>; spelledEchoes: Set<string> } {
  const byKey = new Map<string, Deletion[]>()
  const spelledEchoes = new Set<string>()
  const add = (t: RuleTombstone, own: boolean) => {
    if (t.spellingHash) spelledEchoes.add(echoKey(t))
    const k = tombstoneKey(t)
    const list = byKey.get(k)
    if (list) list.push({ t, own })
    else byKey.set(k, [{ t, own }])
  }
  for (const t of local) add(t, true)
  for (const t of cloud) add(t, false)
  if (spelledEchoes.size > 0) {
    for (const [k, list] of byKey) {
      const kept = list.filter(({ t }) => t.spellingHash || !spelledEchoes.has(echoKey(t)))
      if (kept.length !== list.length) byKey.set(k, kept)
    }
  }
  return { byKey, spelledEchoes }
}

/**
 * Does deletion `t` delete the version of a rule whose max(createdAt,
 * editedAt) is `rt` and whose spelling fingerprint `spelling()` returns?
 * Round 3 (2026-11-06): a tombstone that records the deleted rule's spelling
 * (spellingHash, v1.0.9+) deletes only that spelling — the canonical twin of
 * a deleted rule (dedupe keeps a disabled user_manual 「112年度訴字第304號」
 * next to an enabled auto 「112訴304」) is not touched, wherever its copy
 * is and whether or not this push carries it. One without (v1.0.8, v1.0.7,
 * imports) deletes every spelling of the triple, as it always did — and a
 * v1.0.8 machine, which ignores spellings, still applies every tombstone
 * that way.
 */
function deletes(t: RuleTombstone, rt: number, spelling: () => string): boolean {
  return t.deletedAt >= rt && (!t.spellingHash || t.spellingHash === spelling())
}

/**
 * Pure pull merge (items 1–3, 2026-11-05).
 *
 *   - Same id on both sides → mergeRuleVersions (newer editedAt wins the
 *     semantic fields; the old engine took the cloud copy wholesale, so a
 *     stale copy reverted newer edits).
 *   - Local only, with a cloud disabled stub → the stub's disable applies if
 *     it is at least as new as the local edit (item 1).
 *   - Per-device local rules (auto_scan / orphaned) are untouched.
 *   - Tombstones (local ∪ cloud) suppress a rule ONLY if its version
 *     predates the deletion: deletedAt >= max(createdAt, editedAt) — in BOTH
 *     directions (item 3). Incoming cloud rules are now checked against LOCAL
 *     tombstones too (a delete here used to be undone by the next pull and
 *     re-published), and a rule re-created after an old deletion is no longer
 *     killed by that tombstone. Round 3: a tombstone recording the deleted
 *     rule's spelling suppresses only that spelling (see `deletes`).
 *   - mode 'replace' additionally drops local syncable rules the cloud does
 *     not represent at all.
 *   - adoption (first participation, 2026-11-06): see adoptionProtects.
 */
export function mergeCloudIntoLocal(input: MergeInput): MergeOutput {
  const { byKey: deletions, spelledEchoes } = deletionsByKey(input.localTombstones, input.cloudTombstones)
  const spellings = new Map<Rule, string>()
  const spellingOf = (r: Rule): string => {
    let s = spellings.get(r)
    if (s === undefined) {
      s = spellingFingerprint(legacySignalNorm(r.type, r.signal))
      spellings.set(r, s)
    }
    return s
  }
  // `>=` (in `deletes`): a deletion always happens after the version it
  // deletes, but both can land in the same millisecond; a re-creation in
  // that same ms is not a human-reachable case.
  const suppressed = (r: Rule): boolean => {
    const list = deletions.get(ruleKey(r))
    if (!list) return false
    const rt = ruleTimeMs(r)
    return list.some(({ t }) => deletes(t, rt, () => spellingOf(r)))
  }
  const newTombstones = cloudTombstonesToKeep(input, spelledEchoes)
  const stats = { added: 0, updated: 0, dropped: 0, suppressedIncoming: 0, protected: 0 }
  const protects = input.adoption ? adoptionProtects(deletions, input.adoption, spellingOf) : undefined
  /**
   * A suppressed rule adoption keeps: re-stamped strictly after every known
   * deletion of its canonical triple (any spelling — a v1.0.8 machine applies
   * them all), so it also survives on every other machine (convergent) —
   * undefined when it really goes.
   */
  const keepAdopted = (m: Rule, incoming = false): Rule | undefined => {
    if (!protects?.(m, incoming)) return undefined
    const d = Math.max(...deletions.get(ruleKey(m))!.map(({ t }) => t.deletedAt))
    stats.protected++
    return { ...m, editedAt: Math.max(m.editedAt ?? 0, d + 1) }
  }
  const cloudById = new Map(input.cloudRules.map((r) => [r.id, r]))
  const stubById = new Map(input.cloudStubs.map((s) => [s.id, s]))
  const next: Rule[] = []
  let changed = false
  for (const r of input.local) {
    const c = cloudById.get(r.id)
    if (c) {
      cloudById.delete(r.id)
      let m = mergeRuleVersions(r, c)
      if (shouldSyncRule(m) && suppressed(m)) {
        const kept = keepAdopted(m)
        if (!kept) {
          changed = true
          stats.dropped++
          continue
        }
        m = kept
      }
      if (m !== r) {
        changed = true
        stats.updated++
      }
      next.push(m)
      continue
    }
    if (!shouldSyncRule(r)) {
      // Per-device: auto_scan never syncs. An orphaned copy (its target is
      // dead HERE) takes the cloud's edits — a full row above, a disable stub
      // here (S3: stubs used to be ignored while full rows applied) — but is
      // never suppressed or dropped while orphaned.
      const stub = r.orphaned && SYNCABLE_SOURCES.has(r.source) ? stubById.get(r.id) : undefined
      const m = stub ? applyStub(r, stub) : r
      if (m !== r) {
        changed = true
        stats.updated++
      }
      next.push(m)
      continue
    }
    const stub = stubById.get(r.id)
    let m = stub ? applyStub(r, stub) : r
    if (!stub && input.mode === 'replace') {
      changed = true
      stats.dropped++
      continue
    }
    if (suppressed(m)) {
      const kept = keepAdopted(m)
      if (!kept) {
        changed = true
        stats.dropped++
        continue
      }
      m = kept
    }
    if (m !== r) {
      changed = true
      stats.updated++
    }
    next.push(m)
  }
  for (const cloudRule of cloudById.values()) {
    let c = cloudRule
    if (suppressed(c)) {
      const kept = keepAdopted(c, true)
      if (!kept) {
        stats.suppressedIncoming++
        continue
      }
      c = kept
    }
    next.push(c)
    changed = true
    stats.added++
  }
  return { next: changed ? next : input.local, newTombstones, stats }
}

/**
 * The cloud tombstones worth persisting locally (round 3, 2026-11-06): each
 * unless a local (or already chosen) tombstone of its canonical triple
 * deletes everything it deletes, at least as late — a spelling-less one
 * covers every spelling, a spelled one only its own. Keyed by triple alone,
 * a deletion of the canonical twin used to be dropped as "older" (or, once
 * stored, to overwrite the other spelling). A spelling-less copy of a spelled
 * deletion (a v1.0.8 re-push) is never news.
 */
function cloudTombstonesToKeep(input: MergeInput, spelledEchoes: ReadonlySet<string>): RuleTombstone[] {
  if (input.cloudTombstones.length === 0) return []
  // A local spelling-less copy of a spelled deletion (stored before the
  // original reached this machine) covers nothing: the original is stored,
  // and storage folds the copy into it.
  const known = groupTombstonesByKey(
    input.localTombstones.filter((t) => t.spellingHash || !spelledEchoes.has(echoKey(t))),
  )
  const out: RuleTombstone[] = []
  // Newest first, so an older copy of a deletion already chosen is covered.
  for (const c of [...input.cloudTombstones].sort((a, b) => b.deletedAt - a.deletedAt)) {
    if (!c.spellingHash && spelledEchoes.has(echoKey(c))) continue
    const k = tombstoneKey(c)
    if (coversDeletion(known.get(k), c)) continue
    out.push(c)
    known.set(k, [...(known.get(k) ?? []), c])
  }
  return out
}

function groupTombstonesByKey(tombstones: readonly RuleTombstone[]): Map<string, RuleTombstone[]> {
  const out = new Map<string, RuleTombstone[]>()
  for (const t of tombstones) {
    const k = tombstoneKey(t)
    const list = out.get(k)
    if (list) list.push(t)
    else out.set(k, [t])
  }
  return out
}

/**
 * Does a known deletion of the same canonical triple delete everything `t`
 * deletes, at least as late? A spelling-less one covers every spelling, a
 * spelled one only its own (round 3).
 */
function coversDeletion(known: readonly RuleTombstone[] | undefined, t: RuleTombstone): boolean {
  return (known ?? []).some(
    (x) => x.deletedAt >= t.deletedAt && (!x.spellingHash || x.spellingHash === t.spellingHash),
  )
}

/**
 * Adoption (A1, 2026-11-06). v1.0.7 never let a tombstone remove a rule the
 * cloud still had, and a machine's own tombstones never removed its own live
 * rules, so "live rule + newer tombstone of its triple" pairs are ordinary
 * pre-v4 state: a duplicate that outlived its deleted twin, a rule the
 * push-before-pull race brought back, one a v3 restore kept — or, under the
 * 2026-11-05 canonical keys, a rule that merely COLLIDES with a deleted
 * spelling (a user_manual 「112年度訴字第304號」 vs a deleted auto
 * 「112訴304」). v1.0.8's first merge deleted every one of them, on every
 * machine. On the first merge a LOCAL rule is kept when each tombstone that
 * would suppress it is
 *   - this machine's own (it has always coexisted with this library), or
 *   - one v3 had already processed (deletedAt <= the v1.0.7 machine's
 *     lastSyncAt),
 * and a user_manual rule also when a tombstone only collides canonically
 * (the spelling it was written for is known and isn't the rule's own — see
 * deletedSpellingHash; S1, 2026-11-06: v1.0.8+ tombstones store the canonical
 * form, so comparing signalNorm read the user's own deletion of
 * 「112年度訴字第304號」 as a collision and resurrected the rule on every
 * adoption — the v1.0.7 upgrade and every sync re-enable). An INCOMING cloud
 * rule — v3 added every cloud rule — gets the same verdict, except that this
 * machine's own newer tombstone does count (a delete made here). The caller
 * re-stamps a kept rule above the deletion, so the verdict holds on every
 * machine.
 */
function adoptionProtects(
  deletions: ReadonlyMap<string, Deletion[]>,
  adoption: MergeAdoption,
  spellingOf: (r: Rule) => string,
): (m: Rule, incoming: boolean) => boolean {
  const processedBefore = adoption.processedBeforeMs
  return (m, incoming) => {
    const rt = ruleTimeMs(m)
    for (const { t, own } of deletions.get(ruleKey(m)) ?? []) {
      // Doesn't delete this version — or (round 3) records another spelling.
      if (!deletes(t, rt, () => spellingOf(m))) continue
      if (own && !incoming) continue
      if (processedBefore !== undefined && t.deletedAt <= processedBefore) continue
      if (m.source === 'user_manual') {
        // Only a spelling-less tombstone gets here with another spelling
        // (a spelled one that deletes `m` records m's own): known when v1.0.7
        // wrote it, unknown — never a collision — when v1.0.8 did.
        const deleted = deletedSpellingHash(t)
        if (deleted !== undefined && deleted !== spellingOf(m)) continue // a canonical collision
      }
      return false
    }
    return true
  }
}

// ---- remote pulls / retries --------------------------------------------------------

function requestRemotePull(reason: string): void {
  // P-2: debounce so a burst of remote changes collapses into one pull.
  if (pullDebounceTimer) clearTimeout(pullDebounceTimer)
  pullDebounceTimer = setTimeout(() => {
    pullDebounceTimer = null
    void dispatchRemotePull(reason)
  }, PULL_DEBOUNCE_MS)
}

async function dispatchRemotePull(reason: string): Promise<void> {
  // Already waiting in the queue: it will read the newest meta when it runs.
  if (remotePullQueued) return
  remotePullQueued = true
  try {
    await enqueue(async () => {
      remotePullQueued = false
      const settings = await getSettings()
      if (!settings.syncEnabled) return
      const meta = await readSyncMeta()
      if (!meta) return
      // Echo of our own push? Skip — we pulled before writing it. Unless
      // another push's chunks now sit under OUR meta (A6, 2026-11-06): the
      // cloud is torn for everyone else and only its owner can see that its
      // content hasn't changed — rewrite it (the no-op check verifies the
      // items too, so this push really writes).
      if (meta.sourceMachineId === settings.syncMachineId) {
        if (!(await snapshotIntact(meta))) schedulePush('repair-own-torn-meta', 0, { flush: true })
        return
      }
      const { state } = await readLocalSyncState()
      if (meta.pushId && meta.pushId === state.lastMergedPushId) return // already merged
      // Bug #E: union (with tombstone-aware drop), never replace.
      await doPull(reason, { mode: 'union', source: 'pull-remote' })
    })
  } catch (e) {
    const errMsg = errorMessage(e)
    console.warn('[mail-organizer] sync pull on remote-change failed', e)
    await recordSyncError({ source: 'pull-remote', reason: errMsg }).catch(() => {})
    await logError('sync:pull-remote', errMsg, { reason }).catch(() => {})
  }
}

function scheduleTornRetry(): void {
  if (tornRetryTimer) return
  const delay = Math.min(
    SYNC_TORN_RETRY_BASE_MS * 2 ** Math.max(0, tornAttempts - 1),
    SYNC_TORN_RETRY_MAX_MS,
  )
  tornRetryTimer = setTimeout(() => {
    tornRetryTimer = null
    void dispatchRemotePull('torn-retry')
  }, delay)
}

function schedulePushRetry(kind: StorageErrorKind): void {
  pushRetryAttempts++
  const n = pushRetryAttempts - 1
  const delay =
    kind === 'rate-minute'
      ? SYNC_RATE_LIMIT_MINUTE_RETRY_MS
      : kind === 'rate-hour'
        ? Math.min(SYNC_RATE_LIMIT_HOUR_RETRY_MS * 2 ** n, 60 * 60_000)
        : Math.min(SYNC_PUSH_RETRY_BASE_MS * 2 ** n, SYNC_PUSH_RETRY_MAX_MS)
  if (pushRetryTimer) clearTimeout(pushRetryTimer)
  pushRetryTimer = setTimeout(() => {
    pushRetryTimer = null
    schedulePush('retry', 0)
  }, delay)
}

function resetPushRetry(): void {
  pushRetryAttempts = 0
  if (pushRetryTimer) {
    clearTimeout(pushRetryTimer)
    pushRetryTimer = null
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
