import {
  CONVERSATION_MEMORY_CAP,
  DECAY_AFTER_STABLE,
  FOLDER_ACTIVITY_CAP,
  FOLDER_ACTIVITY_MAX_AGE_MS,
  LOCAL_STORAGE_QUOTA_BYTES,
  LOCAL_STORAGE_WARN_THRESHOLD,
  RECENTLY_PROCESSED_CAP,
  RECENTLY_PROCESSED_TTL_MS,
  RULE_HISTORY_CAP,
  SKIP_HISTORY_CAP,
  SKIP_HISTORY_TTL_MS,
  SUBJECT_MEMORY_CAP,
  THREAD_MEMORY_MAX_AGE_MS,
  TOMBSTONE_CAP,
} from './constants'
import { joinFolderPath } from './outlook-api'
import {
  DEFAULT_METRICS,
  DEFAULT_SETTINGS,
  type FolderActivity,
  type FolderCache,
  type MailFolderNode,
  type Metrics,
  type Rule,
  type RuleEvent,
  type RuleTombstone,
  type Settings,
  type ThreadMemoryEntry,
  type UndoSnapshot,
  type WeeklyDigestState,
} from './types'

const KEY_RULES = 'rules'
const KEY_FOLDER_CACHE = 'folderCache'
const KEY_SETTINGS = 'settings'
const KEY_METRICS = 'metrics'
const KEY_SKIP_HISTORY = 'skipHistory'
const KEY_UNDO_SNAPSHOT = 'undoSnapshot'
const KEY_RULE_TOMBSTONES = 'ruleTombstones'
const KEY_RULE_HISTORY = 'ruleHistory'
const KEY_FOLDER_ACTIVITY = 'folderActivity'
const KEY_WEEKLY_DIGEST = 'weeklyDigest'
const KEY_CONVERSATION_MEMORY = 'conversationMemory'
const KEY_SUBJECT_MEMORY = 'subjectMemory'
const KEY_RECENTLY_PROCESSED = 'recentlyProcessed'

// FOLDER_ACTIVITY_CAP / FOLDER_ACTIVITY_MAX_AGE_MS / TOMBSTONE_CAP /
// RULE_HISTORY_CAP: see src/shared/constants.ts.

/**
 * Map of emailId → first-recorded-at unix-ms. Emails the user has already
 * decided to keep ("skip" action executed) are tracked here so next classify
 * pass can pre-filter them out of the inbox fetch.
 */
export type SkipHistory = Record<string, number>

/**
 * Folder IDs that should NOT be stored on a rule. Path-lookup fallback covers
 * these gracefully at match time, but leaving them in storage is misleading.
 */
function sanitizeStoredFolderId(id: string | undefined | null): string {
  if (!id) return ''
  if (id.startsWith('PLACEHOLDER')) return ''
  if (id.startsWith('pending:')) return ''
  if (id.length < 20) return ''
  return id
}

export async function getRules(): Promise<Rule[]> {
  const stored = await chrome.storage.local.get(KEY_RULES)
  const rules: Rule[] = stored[KEY_RULES] ?? []

  // Auto-heal: prior versions saved `pending:` sentinels and other invalid
  // IDs into rules. Lazily clean them on read so existing users get fixed
  // without manual action. Persist back only if anything actually changed.
  let dirty = false
  const cleaned = rules.map((r) => {
    const next = sanitizeStoredFolderId(r.targetFolderId)
    if (next !== r.targetFolderId) {
      dirty = true
      return { ...r, targetFolderId: next }
    }
    return r
  })
  if (dirty) {
    await chrome.storage.local.set({ [KEY_RULES]: cleaned })
  }
  return cleaned
}

export async function setRules(rules: Rule[]): Promise<void> {
  await chrome.storage.local.set({ [KEY_RULES]: rules })
}

/**
 * Pick only known fields, drop legacy/unknown keys (e.g. removed
 * `confidenceThreshold`). Type-guards each so a junk value falls back to
 * default instead of crashing reads downstream.
 */
function sanitizeSettings(raw: Record<string, unknown>): Settings {
  return {
    claudeApiKey:
      typeof raw.claudeApiKey === 'string' ? raw.claudeApiKey : DEFAULT_SETTINGS.claudeApiKey,
    claudeModel:
      typeof raw.claudeModel === 'string' && raw.claudeModel.length > 0
        ? raw.claudeModel
        : DEFAULT_SETTINGS.claudeModel,
    batchSize:
      typeof raw.batchSize === 'number' && Number.isFinite(raw.batchSize)
        ? raw.batchSize
        : DEFAULT_SETTINGS.batchSize,
    excludeFolderPrefixes:
      Array.isArray(raw.excludeFolderPrefixes) &&
      raw.excludeFolderPrefixes.every((s) => typeof s === 'string')
        ? (raw.excludeFolderPrefixes as string[])
        : DEFAULT_SETTINGS.excludeFolderPrefixes,
    aiConfidenceThreshold:
      typeof raw.aiConfidenceThreshold === 'number' &&
      Number.isFinite(raw.aiConfidenceThreshold)
        ? raw.aiConfidenceThreshold
        : DEFAULT_SETTINGS.aiConfidenceThreshold,
    skipFlagged:
      typeof raw.skipFlagged === 'boolean' ? raw.skipFlagged : DEFAULT_SETTINGS.skipFlagged,
    showOwaFab:
      typeof raw.showOwaFab === 'boolean' ? raw.showOwaFab : DEFAULT_SETTINGS.showOwaFab,
    prefetchNextBatch:
      typeof raw.prefetchNextBatch === 'boolean'
        ? raw.prefetchNextBatch
        : DEFAULT_SETTINGS.prefetchNextBatch,
    recentActivityIncludePrefixes:
      Array.isArray(raw.recentActivityIncludePrefixes) &&
      raw.recentActivityIncludePrefixes.every((s) => typeof s === 'string')
        ? (raw.recentActivityIncludePrefixes as string[])
        : DEFAULT_SETTINGS.recentActivityIncludePrefixes,
    recentActivityIncludeLeafNames:
      Array.isArray(raw.recentActivityIncludeLeafNames) &&
      raw.recentActivityIncludeLeafNames.every((s) => typeof s === 'string')
        ? (raw.recentActivityIncludeLeafNames as string[])
        : DEFAULT_SETTINGS.recentActivityIncludeLeafNames,
    internalDomains:
      Array.isArray(raw.internalDomains) &&
      raw.internalDomains.every((s) => typeof s === 'string')
        ? (raw.internalDomains as string[]).map((s) => s.toLowerCase().trim()).filter(Boolean)
        : DEFAULT_SETTINGS.internalDomains,
    primaryRootPath:
      typeof raw.primaryRootPath === 'string' ? raw.primaryRootPath : DEFAULT_SETTINGS.primaryRootPath,
    internalSubjectCategories:
      Array.isArray(raw.internalSubjectCategories) &&
      raw.internalSubjectCategories.every((s) => typeof s === 'string')
        ? (raw.internalSubjectCategories as string[])
            .map((s) => s.trim())
            .filter(Boolean)
        : DEFAULT_SETTINGS.internalSubjectCategories,
    aiIncludeFewShotExamples:
      typeof raw.aiIncludeFewShotExamples === 'boolean'
        ? raw.aiIncludeFewShotExamples
        : DEFAULT_SETTINGS.aiIncludeFewShotExamples,
    syncEnabled:
      typeof raw.syncEnabled === 'boolean'
        ? raw.syncEnabled
        : DEFAULT_SETTINGS.syncEnabled,
    syncMachineId:
      typeof raw.syncMachineId === 'string'
        ? raw.syncMachineId
        : DEFAULT_SETTINGS.syncMachineId,
    lastSyncAt:
      typeof raw.lastSyncAt === 'string'
        ? raw.lastSyncAt
        : DEFAULT_SETTINGS.lastSyncAt,
    onboardingDismissed:
      typeof raw.onboardingDismissed === 'boolean'
        ? raw.onboardingDismissed
        : DEFAULT_SETTINGS.onboardingDismissed,
  }
}

const ALLOWED_SETTINGS_KEYS: ReadonlySet<keyof Settings> = new Set([
  'claudeApiKey',
  'claudeModel',
  'batchSize',
  'excludeFolderPrefixes',
  'aiConfidenceThreshold',
  'skipFlagged',
  'showOwaFab',
  'prefetchNextBatch',
  'recentActivityIncludePrefixes',
  'recentActivityIncludeLeafNames',
  'internalDomains',
  'primaryRootPath',
  'internalSubjectCategories',
  'aiIncludeFewShotExamples',
  'syncEnabled',
  'syncMachineId',
  'lastSyncAt',
  'onboardingDismissed',
])

export async function getSettings(): Promise<Settings> {
  const stored = await chrome.storage.local.get(KEY_SETTINGS)
  const raw = (stored[KEY_SETTINGS] ?? {}) as Record<string, unknown>
  // Migration (G1 generification, 2026-05-22): existing users — those
  // with ANY prior settings in storage but no `internalDomains` field —
  // get the previously-hardcoded values backfilled. Fresh installs (raw
  // is `{}`) skip backfill and land on empty defaults, which triggers
  // the onboarding banner in the popup. Once the field exists in raw,
  // this branch is a no-op forever.
  let migrated = false
  if (Object.keys(raw).length > 0 && !('internalDomains' in raw)) {
    raw.internalDomains = ['example.com']
    raw.primaryRootPath = '案件'
    raw.internalSubjectCategories = ['工時', '薪資', '利衝', '行政', '公告']
    migrated = true
  }
  const cleaned = sanitizeSettings(raw)
  // Persist back if storage had legacy fields, so the diagnostic export
  // (and any other dump) stops surfacing them — OR if we just backfilled
  // the migration, so the next read is fast and the values appear in
  // the options UI as the current canonical values.
  const hasExtraKeys = Object.keys(raw).some(
    (k) => !ALLOWED_SETTINGS_KEYS.has(k as keyof Settings),
  )
  if (hasExtraKeys || migrated) {
    await chrome.storage.local.set({ [KEY_SETTINGS]: cleaned })
  }
  return cleaned
}

export async function setSettings(patch: Partial<Settings>): Promise<Settings> {
  return withSettingsLock(async () => {
    const next = { ...(await getSettings()), ...patch }
    await chrome.storage.local.set({ [KEY_SETTINGS]: next })
    return next
  })
}

export async function getFolderCache(): Promise<FolderCache | undefined> {
  const stored = await chrome.storage.local.get(KEY_FOLDER_CACHE)
  return stored[KEY_FOLDER_CACHE]
}

export async function setFolderCache(cache: FolderCache): Promise<void> {
  await chrome.storage.local.set({ [KEY_FOLDER_CACHE]: cache })
}

/**
 * Splice a newly created Outlook folder into the cached tree without
 * invalidating the rest. parentId === undefined means it's a top-level folder.
 * Returns true if the splice succeeded (parent found or it's top-level).
 *
 * This keeps the cache usable after `executeItem` runs new_folder actions —
 * otherwise the next classify run wouldn't see the new folder until the 24h
 * TTL expires.
 */
/**
 * Serialize concurrent splices so the read/mutate/write of folderCache is
 * atomic. Without this, two new_folder operations completing close in time
 * can both read the same initial tree, each splice their own copy, and the
 * later write clobbers the earlier folder entry.
 */
let folderCacheWriteChain: Promise<void> = Promise.resolve()

async function withFolderCacheLock<T>(fn: () => Promise<T>): Promise<T> {
  const release = folderCacheWriteChain
  let resolve: () => void = () => {}
  folderCacheWriteChain = new Promise<void>((r) => {
    resolve = r
  })
  try {
    await release
    return await fn()
  } finally {
    resolve()
  }
}

// Generic single-resource write-serialization. Each call to createWriteMutex
// returns its own scoped chain — independent resources don't block each other.
function createWriteMutex(): <T>(fn: () => Promise<T>) => Promise<T> {
  let chain: Promise<void> = Promise.resolve()
  return async <T>(fn: () => Promise<T>): Promise<T> => {
    const release = chain
    let resolve: () => void = () => {}
    chain = new Promise<void>((r) => {
      resolve = r
    })
    try {
      await release
      return await fn()
    } finally {
      resolve()
    }
  }
}

const withFolderActivityLock = createWriteMutex()
const withSkipHistoryLock = createWriteMutex()
const withMetricsLock = createWriteMutex()
// Tombstones share data with rules — rules.ts callers already hold withRulesLock
// when they touch tombstones, but defensive in case other callers appear later.
const withTombstoneLock = createWriteMutex()
const withConversationMemoryLock = createWriteMutex()
const withSubjectMemoryLock = createWriteMutex()
const withRecentlyProcessedLock = createWriteMutex()
// F11 (2026-06-03): the rule-history audit log did an unserialized
// read-append-write. Callers invoke recordRuleEvents AFTER their
// mutateRules critical section has released (so the rules lock doesn't
// cover it), and several independent paths append concurrently
// (reconcile-on-preflight, a user delete in Options, import). Two
// overlapping appends both read the same N-entry log and the later
// write clobbers the earlier's events. Serialize the whole RMW.
const withHistoryLock = createWriteMutex()
// Bug #I: settings did read-modify-write without serialisation. Concurrent
// patches (e.g. sync engine writing lastSyncAt while the options page
// toggles skipFlagged) raced on the read step — both saw old state, both
// merged their patch, last write wins, the other patch's field lost. With
// this mutex each patch reads the FRESH state before merging.
const withSettingsLock = createWriteMutex()

export async function addFolderToCache(
  folder: { Id: string; DisplayName: string; ParentFolderId?: string },
  parentId: string | undefined,
): Promise<boolean> {
  return withFolderCacheLock(async () => {
    const cache = await getFolderCache()
    if (!cache) return false

    // Find parent path (if any) by walking the tree once.
    let parentPath: string | undefined
    let parentChildren: MailFolderNode[] | undefined
    if (parentId) {
      const walk = (nodes: MailFolderNode[]): boolean => {
        for (const n of nodes) {
          if (n.id === parentId) {
            parentPath = n.path
            parentChildren = n.children
            return true
          }
          if (walk(n.children)) return true
        }
        return false
      }
      if (!walk(cache.tree)) return false
    }

    const newNode: MailFolderNode = {
      id: folder.Id,
      displayName: folder.DisplayName,
      parentFolderId: folder.ParentFolderId,
      path: joinFolderPath(parentPath, folder.DisplayName),
      children: [],
    }

    // Don't duplicate if (somehow) already there
    const target = parentChildren ?? cache.tree
    if (target.some((n) => n.id === folder.Id)) return true
    target.push(newNode)

    await setFolderCache(cache)
    return true
  })
}

export async function getMetrics(): Promise<Metrics> {
  const stored = await chrome.storage.local.get(KEY_METRICS)
  return { ...DEFAULT_METRICS, ...(stored[KEY_METRICS] ?? {}) }
}

// ---- Skip history ----------------------------------------------------------

// SKIP_HISTORY_TTL_MS / SKIP_HISTORY_CAP: see src/shared/constants.ts.
// Both bounds: TTL (skipped 60 days ago is unlikely to recur) AND count
// (cap blast radius of a degenerate 10k-skip session within the window).

function pruneSkipHistory(history: SkipHistory): SkipHistory {
  const now = Date.now()
  const cutoff = now - SKIP_HISTORY_TTL_MS
  // Drop TTL-expired first, then sort remaining by recency and keep top N.
  const fresh: Array<[string, number]> = []
  for (const [id, ts] of Object.entries(history)) {
    if (typeof ts === 'number' && ts >= cutoff) fresh.push([id, ts])
  }
  if (fresh.length > SKIP_HISTORY_CAP) {
    fresh.sort((a, b) => b[1] - a[1]) // newest first
    fresh.length = SKIP_HISTORY_CAP
  }
  return Object.fromEntries(fresh)
}

export async function getSkipHistory(): Promise<SkipHistory> {
  const stored = await chrome.storage.local.get(KEY_SKIP_HISTORY)
  return stored[KEY_SKIP_HISTORY] ?? {}
}

/**
 * Add a batch of email IDs to the skip history. Returns how many were
 * actually new (idempotent on duplicates).
 */
export async function addToSkipHistory(emailIds: string[]): Promise<number> {
  if (emailIds.length === 0) return 0
  return withSkipHistoryLock(async () => {
    const existing = await getSkipHistory()
    const now = Date.now()
    let added = 0
    for (const id of emailIds) {
      if (!id) continue
      if (!(id in existing)) {
        existing[id] = now
        added++
      }
    }
    if (added > 0) {
      // Opportunistic prune on every write — cheap (Object.entries scan)
      // and keeps storage from drifting if the user goes months without
      // a manual clear.
      const pruned = pruneSkipHistory(existing)
      await chrome.storage.local.set({ [KEY_SKIP_HISTORY]: pruned })
    }
    return added
  })
}

export async function clearSkipHistory(): Promise<number> {
  return withSkipHistoryLock(async () => {
    const existing = await getSkipHistory()
    const count = Object.keys(existing).length
    await chrome.storage.local.remove(KEY_SKIP_HISTORY)
    return count
  })
}

export async function getSkipHistoryCount(): Promise<number> {
  const existing = await getSkipHistory()
  return Object.keys(existing).length
}

// ---- Recently-processed ledger --------------------------------------------
//
// emailId → unix-ms when we successfully moved / deleted it. Used to
// pre-filter the NEXT batch's inbox fetch so an email we just handled
// can't reappear via Outlook's read-after-write lag (its now-dead id
// would 404 on the re-move). Distinct from skipHistory:
//   - skipHistory  = user DELIBERATELY kept it in inbox → exclude 60 days
//   - recentlyProcessed = we MOVED/DELETED it → exclude ~15 min (just
//     long enough to outlast Outlook propagation; after that it's
//     genuinely gone from inbox and won't be re-listed anyway)

/** Map of emailId → unix-ms recorded when moved/deleted. */
export type RecentlyProcessed = Record<string, number>

function pruneRecentlyProcessed(rec: RecentlyProcessed): RecentlyProcessed {
  const cutoff = Date.now() - RECENTLY_PROCESSED_TTL_MS
  const fresh: Array<[string, number]> = []
  for (const [id, ts] of Object.entries(rec)) {
    if (typeof ts === 'number' && ts >= cutoff) fresh.push([id, ts])
  }
  if (fresh.length > RECENTLY_PROCESSED_CAP) {
    fresh.sort((a, b) => b[1] - a[1]) // newest first
    fresh.length = RECENTLY_PROCESSED_CAP
  }
  return Object.fromEntries(fresh)
}

export async function getRecentlyProcessed(): Promise<RecentlyProcessed> {
  const stored = await chrome.storage.local.get(KEY_RECENTLY_PROCESSED)
  return (stored[KEY_RECENTLY_PROCESSED] as RecentlyProcessed | undefined) ?? {}
}

/**
 * Returns the set of email IDs processed within the TTL window (TTL-pruned
 * on read so callers always get a current view). Cheap helper for the
 * inbox-filter call sites.
 */
export async function getRecentlyProcessedIds(): Promise<Set<string>> {
  const rec = await getRecentlyProcessed()
  const cutoff = Date.now() - RECENTLY_PROCESSED_TTL_MS
  const ids = new Set<string>()
  for (const [id, ts] of Object.entries(rec)) {
    if (typeof ts === 'number' && ts >= cutoff) ids.add(id)
  }
  return ids
}

/** Record email IDs we just moved/deleted. Idempotent on duplicates. */
export async function addToRecentlyProcessed(emailIds: string[]): Promise<void> {
  if (emailIds.length === 0) return
  await withRecentlyProcessedLock(async () => {
    const existing = await getRecentlyProcessed()
    const now = Date.now()
    let changed = false
    for (const id of emailIds) {
      if (!id) continue
      // Always refresh the timestamp — re-touching an id extends its
      // exclusion window, which is the desired behaviour if it somehow
      // gets re-processed.
      existing[id] = now
      changed = true
    }
    if (changed) {
      await chrome.storage.local.set({
        [KEY_RECENTLY_PROCESSED]: pruneRecentlyProcessed(existing),
      })
    }
  })
}

// ---- Rule tombstones -------------------------------------------------------

export async function getRuleTombstones(): Promise<RuleTombstone[]> {
  const r = await chrome.storage.local.get(KEY_RULE_TOMBSTONES)
  return (r[KEY_RULE_TOMBSTONES] as RuleTombstone[] | undefined) ?? []
}

async function setRuleTombstones(list: RuleTombstone[]): Promise<void> {
  // F13 (2026-06-03): cap by AGE, not array position. addRuleTombstones
  // builds the list from a Map keyed by triple — when an existing triple
  // is re-tombstoned, its value (deletedAt) updates but its insertion
  // POSITION stays put. A position-based tail-slice could therefore drop
  // a freshly re-tombstoned entry sitting near the front while keeping a
  // never-touched older one at the tail — the opposite of "drop oldest",
  // letting a recently-deleted rule lose its tombstone and resurrect.
  // Sort by deletedAt ascending and keep the newest TOMBSTONE_CAP.
  if (list.length <= TOMBSTONE_CAP) {
    await chrome.storage.local.set({ [KEY_RULE_TOMBSTONES]: list })
    return
  }
  const byAge = [...list].sort((a, b) => a.deletedAt - b.deletedAt)
  const capped = byAge.slice(byAge.length - TOMBSTONE_CAP)
  await chrome.storage.local.set({ [KEY_RULE_TOMBSTONES]: capped })
}

export async function addRuleTombstones(items: RuleTombstone[]): Promise<void> {
  if (items.length === 0) return
  await withTombstoneLock(async () => {
    const existing = await getRuleTombstones()
    // Dedup by triple — if the same triple is already tombstoned, keep the
    // newer deletedAt (so cleanup by age behaves correctly).
    const key = (t: RuleTombstone) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`
    const byKey = new Map<string, RuleTombstone>()
    for (const t of existing) byKey.set(key(t), t)
    for (const t of items) byKey.set(key(t), t)
    await setRuleTombstones([...byKey.values()])
  })
}

/**
 * Remove any tombstones matching the given triples — call when the user
 * explicitly re-creates a rule with the same (type, signal, target), so
 * the rule isn't bizarrely auto-disabled by its own ghost.
 */
export async function clearMatchingTombstones(
  triples: Array<{ type: RuleTombstone['type']; signalNorm: string; targetFolderPath: string }>,
): Promise<number> {
  if (triples.length === 0) return 0
  return withTombstoneLock(async () => {
    const existing = await getRuleTombstones()
    const keys = new Set(
      triples.map((t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`),
    )
    const next = existing.filter(
      (t) => !keys.has(`${t.type}::${t.signalNorm}::${t.targetFolderPath}`),
    )
    if (next.length === existing.length) return 0
    await setRuleTombstones(next)
    return existing.length - next.length
  })
}

export async function clearAllRuleTombstones(): Promise<void> {
  await withTombstoneLock(async () => {
    await chrome.storage.local.remove(KEY_RULE_TOMBSTONES)
  })
}

// ---- Rule history (audit log of all mutations) ----------------------------

export async function getRuleEvents(limit?: number): Promise<RuleEvent[]> {
  const r = await chrome.storage.local.get(KEY_RULE_HISTORY)
  const all = (r[KEY_RULE_HISTORY] as RuleEvent[] | undefined) ?? []
  if (limit && limit < all.length) {
    // History is appended chronologically, so most-recent are at the tail.
    return all.slice(all.length - limit)
  }
  return all
}

/**
 * Append one or more events to the rule history. Trims FIFO from the front
 * once over RULE_HISTORY_CAP so unbounded churn doesn't blow storage.
 *
 * Single helper for one-or-many because the bulk paths (addRules from
 * generators, import) want to record N events atomically.
 */
export async function recordRuleEvents(events: RuleEvent[]): Promise<void> {
  if (events.length === 0) return
  // F11: serialize the read-append-write so concurrent appends (e.g. a
  // preflight reconcile and an Options-page delete firing together)
  // can't both read the same log and clobber each other's events.
  await withHistoryLock(async () => {
    const current = await getRuleEvents()
    const next = [...current, ...events]
    // Trim oldest if over cap.
    const trimmed = next.length > RULE_HISTORY_CAP ? next.slice(next.length - RULE_HISTORY_CAP) : next
    await chrome.storage.local.set({ [KEY_RULE_HISTORY]: trimmed })
  })
}

export async function clearRuleHistory(): Promise<void> {
  await chrome.storage.local.remove(KEY_RULE_HISTORY)
}

/**
 * Wipe both AI memory layers — conversation thread memory (routes by
 * ConversationId) and subject memory (routes by normalized subject).
 *
 * These layers feed AI's per-email decisions in the next batch ("this
 * thread previously went to X" / "this normalized subject previously
 * went to X"). Wiping them is part of the "全部刪除規則 + clean slate"
 * workflow so the next batch starts with no inherited routing bias —
 * AI decisions then become pure signals for future rule learning,
 * unpolluted by pre-wipe state.
 *
 * Other state intentionally left untouched by this clear:
 *   - skipHistory: emails the user deliberately kept in inbox; not
 *     routing memory, more like "don't re-ask about these".
 *   - folderActivity: per-folder mail counts for the IdleScreen
 *     "recent activity" panel; UI display only, never feeds learning.
 *   - rule events log: cleared by clearRuleHistory if the caller wants
 *     a true clean slate, but the wipe action itself doesn't touch
 *     audit history.
 */
export async function clearAllAiMemory(): Promise<void> {
  await chrome.storage.local.remove([KEY_CONVERSATION_MEMORY, KEY_SUBJECT_MEMORY])
}

// ---- Folder activity (recent-activity quick-jump panel) -------------------
//
// Records which target folders are seeing recent classify activity so the
// IdleScreen can show a "近日活動" list — one click jumps the user to the
// folder in OWA. Two write paths:
//   1. recordFolderActivityFromBatch — called after every execute batch
//      completes. Merges per-folder counts + recency.
//   2. mergeFolderActivityScan — called by the optional refresh button,
//      which uses Graph API to scan case folders for messages the user
//      dragged in manually (bypassing the extension's classify flow).
//
// Read path: getFolderActivity() returns the list sorted by lastActiveAt
// desc, capped at FOLDER_ACTIVITY_CAP entries (oldest pruned on write).

export async function getFolderActivity(): Promise<FolderActivity[]> {
  const r = await chrome.storage.local.get(KEY_FOLDER_ACTIVITY)
  const list = (r[KEY_FOLDER_ACTIVITY] as FolderActivity[] | undefined) ?? []
  // Filter at read time too — handles legacy data written before the
  // 30-day retention policy was added. Sort defensively in case a
  // malformed write landed out-of-order.
  const cutoff = Date.now() - FOLDER_ACTIVITY_MAX_AGE_MS
  return [...list]
    .filter((e) => {
      const t = Date.parse(e.lastActiveAt)
      return Number.isFinite(t) ? t >= cutoff : false
    })
    .sort((a, b) => b.lastActiveAt.localeCompare(a.lastActiveAt))
}

async function writeFolderActivity(list: FolderActivity[]): Promise<void> {
  // Dedup by folderId, keep most-recent variant. Then sort desc by
  // lastActiveAt, drop anything older than the retention window, and
  // cap at FOLDER_ACTIVITY_CAP. The age filter is the primary policy
  // (panel only shows current work); the cap is belt-and-suspenders.
  const byId = new Map<string, FolderActivity>()
  for (const entry of list) {
    const prev = byId.get(entry.folderId)
    if (!prev || prev.lastActiveAt < entry.lastActiveAt) {
      byId.set(entry.folderId, entry)
    }
  }
  const cutoff = Date.now() - FOLDER_ACTIVITY_MAX_AGE_MS
  const fresh = [...byId.values()].filter((e) => {
    const t = Date.parse(e.lastActiveAt)
    return Number.isFinite(t) ? t >= cutoff : false
  })
  const sorted = fresh.sort((a, b) => b.lastActiveAt.localeCompare(a.lastActiveAt))
  const capped =
    sorted.length > FOLDER_ACTIVITY_CAP ? sorted.slice(0, FOLDER_ACTIVITY_CAP) : sorted
  // P-1: diff before write. Two parallel scans landing on the same
  // results, or a refresh that finds the same latestMessage as last
  // time, would otherwise fire onChanged → schedulePush → wasted sync
  // round-trip. Compare JSON serialisation (small payload, ~few KB).
  const existing = await chrome.storage.local.get(KEY_FOLDER_ACTIVITY)
  const currentJson = JSON.stringify(existing[KEY_FOLDER_ACTIVITY] ?? [])
  const nextJson = JSON.stringify(capped)
  if (currentJson === nextJson) return
  await chrome.storage.local.set({ [KEY_FOLDER_ACTIVITY]: capped })
}

/**
 * Record folder activity from a freshly completed execute batch.
 * `targets` is the per-folder tally that execute.ts builds from results.
 * - new folderIds appended; existing entries updated with the larger
 *   lastActiveAt + the batch count overwriting recentCount (we want the
 *   panel to show "+N from latest batch" not cumulative).
 */
export async function recordFolderActivityFromBatch(
  targets: Array<{
    folderId: string
    folderPath: string
    count: number
    /** Subject of the most recent email moved to this folder in this batch.
     * Lets the panel show "what's the latest" without requiring a separate
     * Graph API scan via the 重新整理 button. */
    latestMessage?: {
      subject: string
      from: string
      receivedAt: string
    }
  }>,
  batchAt: string,
): Promise<void> {
  if (targets.length === 0) return
  await withFolderActivityLock(async () => {
    const existing = await getFolderActivity()
    const byId = new Map(existing.map((e) => [e.folderId, e]))
    for (const t of targets) {
      if (!t.folderId) continue // skip pending: / placeholder ids
      const prev = byId.get(t.folderId)
      // Preserve prev.latestMessage when the batch didn't supply one
      // (defense-in-depth: the older code path always wiped this field,
      // erasing the result of any prior 重新整理 scan on the next batch).
      byId.set(t.folderId, {
        folderId: t.folderId,
        folderPath: t.folderPath,
        lastActiveAt: batchAt,
        recentCount: t.count,
        lastBatchAt: batchAt,
        latestMessage: t.latestMessage ?? prev?.latestMessage,
      })
    }
    await writeFolderActivity([...byId.values()])
  })
}

/**
 * Merge results from a Graph API folder-scan (Phase 3 refresh). Each entry
 * has a folder's latest message timestamp from the live mailbox — used to
 * catch mail the user dragged into the folder outside our classify flow.
 * recentCount is NOT updated here (we don't know it from a top-1 scan);
 * only lastActiveAt is bumped if the scan found a newer message.
 */
export async function mergeFolderActivityScan(
  scans: Array<{
    folderId: string
    folderPath: string
    latestMessageAt: string
    latestMessage?: {
      subject: string
      from: string
      receivedAt: string
    }
  }>,
): Promise<void> {
  if (scans.length === 0) return
  await withFolderActivityLock(async () => {
    const existing = await getFolderActivity()
    const byId = new Map(existing.map((e) => [e.folderId, e]))
    for (const s of scans) {
      if (!s.folderId) continue
      const prev = byId.get(s.folderId)
      // Always update latestMessage on scan — it's the freshest snapshot
      // regardless of whether lastActiveAt itself is newer.
      const nextLatestMessage = s.latestMessage ?? prev?.latestMessage
      if (prev && prev.lastActiveAt >= s.latestMessageAt) {
        // Activity timestamp not newer, but maybe latestMessage updated.
        if (nextLatestMessage !== prev.latestMessage) {
          byId.set(s.folderId, { ...prev, latestMessage: nextLatestMessage })
        }
        continue
      }
      byId.set(s.folderId, {
        folderId: s.folderId,
        folderPath: s.folderPath,
        lastActiveAt: s.latestMessageAt,
        // No batch-count info from a scan — preserve prev count if known.
        recentCount: prev?.recentCount ?? 0,
        lastBatchAt: prev?.lastBatchAt,
        latestMessage: nextLatestMessage,
      })
    }
    await writeFolderActivity([...byId.values()])
  })
}

export async function clearFolderActivity(): Promise<void> {
  await withFolderActivityLock(async () => {
    await chrome.storage.local.remove(KEY_FOLDER_ACTIVITY)
  })
}

/**
 * Merge an incoming list of FolderActivity entries (from chrome.storage.sync
 * pull) into the local list. Conflict resolution: per folderId, keep the
 * entry with the most-recent lastActiveAt. Other-machine activity that's
 * stale loses to local; local activity that's stale loses to other-machine.
 *
 * Does NOT delete local entries that are missing from incoming — sync is
 * additive ("here's what the other machine processed"), not authoritative.
 * Old local entries that the other machine never saw stay on the device.
 */
export async function mergeFolderActivityFromSync(
  incoming: FolderActivity[],
): Promise<void> {
  if (incoming.length === 0) return
  await withFolderActivityLock(async () => {
    const existing = await getFolderActivity()
    const byId = new Map(existing.map((e) => [e.folderId, e]))
    for (const inc of incoming) {
      if (!inc.folderId) continue
      const prev = byId.get(inc.folderId)
      // If incoming wins (newer lastActiveAt), take its full snapshot —
      // including its latestMessage / recentCount, which are the freshest
      // facts the other machine saw. If prev wins, keep it intact.
      if (!prev || prev.lastActiveAt < inc.lastActiveAt) {
        byId.set(inc.folderId, inc)
      }
    }
    await writeFolderActivity([...byId.values()])
  })
}

// ---- folderActivity refresh timestamp -------------------------------------
//
// Tracks when refreshFolderActivity (Graph API scan) was last run. Popup's
// IdleScreen uses this to decide whether to auto-trigger refresh on open
// after a sync pull (B machine just received A's activity entries that
// might be hours old; latestMessage on those entries reflects A's pull-
// time view, not the current state of the mailbox).
const KEY_FOLDER_ACTIVITY_REFRESH_AT = 'folderActivityRefreshAt'

export async function getFolderActivityRefreshAt(): Promise<string | null> {
  const r = await chrome.storage.local.get(KEY_FOLDER_ACTIVITY_REFRESH_AT)
  return (r[KEY_FOLDER_ACTIVITY_REFRESH_AT] as string | undefined) ?? null
}

export async function setFolderActivityRefreshAt(at: string): Promise<void> {
  await chrome.storage.local.set({ [KEY_FOLDER_ACTIVITY_REFRESH_AT]: at })
}

// ---- Undo snapshot ---------------------------------------------------------

/**
 * Read the current undo snapshot. Returns null if absent or expired —
 * expired snapshots are also evicted in passing so we don't accumulate
 * dead state if the alarm-based cleanup missed (e.g. browser was closed).
 */
export async function getUndoSnapshot(): Promise<UndoSnapshot | null> {
  const r = await chrome.storage.local.get(KEY_UNDO_SNAPSHOT)
  const snap = r[KEY_UNDO_SNAPSHOT] as UndoSnapshot | undefined
  if (!snap) return null
  if (Date.now() >= snap.expiresAt) {
    await chrome.storage.local.remove(KEY_UNDO_SNAPSHOT)
    return null
  }
  return snap
}

export async function setUndoSnapshot(snap: UndoSnapshot): Promise<void> {
  await chrome.storage.local.set({ [KEY_UNDO_SNAPSHOT]: snap })
}

export async function clearUndoSnapshot(): Promise<void> {
  await chrome.storage.local.remove(KEY_UNDO_SNAPSHOT)
}

// ---- Metrics ---------------------------------------------------------------

export async function bumpMetrics(patch: Partial<Metrics>): Promise<Metrics> {
  return withMetricsLock(async () => {
    const current = await getMetrics()
    const next: Metrics = {
      moved: current.moved + (patch.moved ?? 0),
      deleted: current.deleted + (patch.deleted ?? 0),
      foldersCreated: current.foldersCreated + (patch.foldersCreated ?? 0),
      errors: current.errors + (patch.errors ?? 0),
    }
    await chrome.storage.local.set({ [KEY_METRICS]: next })
    return next
  })
}

// ---- Storage usage -----------------------------------------------------

// ---- Weekly digest --------------------------------------------------------

export async function getWeeklyDigestState(): Promise<WeeklyDigestState | null> {
  const r = await chrome.storage.local.get(KEY_WEEKLY_DIGEST)
  return (r[KEY_WEEKLY_DIGEST] as WeeklyDigestState | undefined) ?? null
}

export async function setWeeklyDigestState(state: WeeklyDigestState): Promise<void> {
  await chrome.storage.local.set({ [KEY_WEEKLY_DIGEST]: state })
}

// ---- Thread memory --------------------------------------------------------
//
// Two parallel maps remember "the user previously filed this thread to
// folder X" so that subsequent emails in the same conversation (including
// internal replies with vague subjects) can be auto-routed.
//
//   conversationMemory:  ConversationId → ThreadMemoryEntry (exact match)
//   subjectMemory:        normalizedSubject → ThreadMemoryEntry (loose
//                         fallback, gated on conflictCount === 0)
//
// Cap at LRU size + age-based retention. Pruned on every write.

// CONVERSATION_MEMORY_CAP / SUBJECT_MEMORY_CAP / THREAD_MEMORY_MAX_AGE_MS:
// see src/shared/constants.ts.

type ThreadMemoryMap = Record<string, ThreadMemoryEntry>

function pruneThreadMemory(
  map: ThreadMemoryMap,
  cap: number,
): ThreadMemoryMap {
  const cutoff = Date.now() - THREAD_MEMORY_MAX_AGE_MS
  // Filter expired
  const entries = Object.entries(map).filter(([, v]) => {
    const t = Date.parse(v.lastFiledAt)
    return Number.isFinite(t) && t >= cutoff
  })
  // Sort by lastFiledAt desc, keep top `cap`
  entries.sort(([, a], [, b]) => b.lastFiledAt.localeCompare(a.lastFiledAt))
  if (entries.length > cap) entries.length = cap
  return Object.fromEntries(entries)
}

export async function getConversationMemory(): Promise<ThreadMemoryMap> {
  const r = await chrome.storage.local.get(KEY_CONVERSATION_MEMORY)
  return (r[KEY_CONVERSATION_MEMORY] as ThreadMemoryMap | undefined) ?? {}
}

export async function getSubjectMemory(): Promise<ThreadMemoryMap> {
  const r = await chrome.storage.local.get(KEY_SUBJECT_MEMORY)
  return (r[KEY_SUBJECT_MEMORY] as ThreadMemoryMap | undefined) ?? {}
}

/**
 * Record one batch of conversation-id filings. Each filing means "the
 * user filed an email with this convId into this folder" — we increment
 * `timesFiled` and overwrite the target (latest wins, since a single
 * conversation should converge to one folder over time).
 *
 * Conflict tracking + decay (mirrors subjectMemory logic): if a filing
 * targets a DIFFERENT folder from the previous one for the same convId,
 * conflictCount increments. The matchThreadMemory check in classify
 * pipeline only uses convId-memory when conflictCount===0 — so a
 * previously-good convId that recently got mis-filed temporarily stops
 * being trusted until DECAY_AFTER_STABLE consecutive same-folder filings
 * earn the trust back.
 *
 * Why this matters: a conversation occasionally drifts (subject changes
 * mid-thread, project pivots). Without decay, a single accidental
 * cross-folder filing locks the convId into "always trusted" forever
 * pointing at the wrong place.
 */
export async function recordConversationFilings(
  filings: Array<{ convId: string; folderId: string; folderPath: string }>,
): Promise<void> {
  if (filings.length === 0) return
  await withConversationMemoryLock(async () => {
    const map = await getConversationMemory()
    const now = new Date().toISOString()
    for (const f of filings) {
      if (!f.convId) continue
      const prev = map[f.convId]
      const sameFolder = prev?.folderId === f.folderId
      const conflict = !!prev && !sameFolder
      // Streak / decay book-keeping, parallel to recordSubjectFilings:
      let nextStreak: number
      let nextConflict = prev?.conflictCount ?? 0
      if (conflict) {
        nextStreak = 0
        nextConflict += 1
      } else {
        nextStreak = (prev?.stableStreak ?? 0) + 1
        if (nextStreak > DECAY_AFTER_STABLE && nextConflict > 0) {
          nextConflict -= 1
        }
      }
      map[f.convId] = {
        folderId: f.folderId,
        folderPath: f.folderPath,
        lastFiledAt: now,
        timesFiled: (prev?.timesFiled ?? 0) + 1,
        conflictCount: nextConflict,
        stableStreak: nextStreak,
      }
    }
    const pruned = pruneThreadMemory(map, CONVERSATION_MEMORY_CAP)
    await chrome.storage.local.set({ [KEY_CONVERSATION_MEMORY]: pruned })
  })
}

/**
 * Record one batch of normalized-subject filings. Differs from
 * conversation map by tracking `conflictCount`: a single subject can
 * legitimately span multiple cases (think "報告" / "請示" / repeated
 * subject lines), so pre-filter only trusts entries with conflictCount
 * === 0. The latest folderId wins for display, but conflictCount
 * accumulates across the lifetime to disqualify ambiguous entries.
 *
 * Decay: a single accidental cross-folder filing used to poison the
 * subject forever. Now: when an entry has been hit ≥ DECAY_AFTER_STABLE
 * times consecutively on the same folder, conflictCount drops by 1 per
 * extra same-folder filing. Lets the subject "earn back" trust over
 * time rather than staying dead from a stray CC. Stored in
 * `stableStreak` so we don't have to scan history.
 */
// DECAY_AFTER_STABLE: see src/shared/constants.ts.

export async function recordSubjectFilings(
  filings: Array<{ normalizedSubject: string; folderId: string; folderPath: string }>,
): Promise<void> {
  if (filings.length === 0) return
  await withSubjectMemoryLock(async () => {
    const map = await getSubjectMemory()
    const now = new Date().toISOString()
    for (const f of filings) {
      if (!f.normalizedSubject) continue
      const prev = map[f.normalizedSubject]
      const sameFolder = prev?.folderId === f.folderId
      const conflict = !!prev && !sameFolder
      // Streak / decay book-keeping. Conflict resets the streak; same
      // folder extends it. Decay only kicks in past the threshold so a
      // newly-conflicted entry stays disqualified for a meaningful
      // re-validation period (5 consecutive same-folder filings).
      let nextStreak: number
      let nextConflict = prev?.conflictCount ?? 0
      if (conflict) {
        nextStreak = 0
        nextConflict += 1
      } else {
        nextStreak = (prev?.stableStreak ?? 0) + 1
        if (nextStreak > DECAY_AFTER_STABLE && nextConflict > 0) {
          nextConflict -= 1
        }
      }
      map[f.normalizedSubject] = {
        folderId: f.folderId,
        folderPath: f.folderPath,
        lastFiledAt: now,
        timesFiled: (prev?.timesFiled ?? 0) + 1,
        conflictCount: nextConflict,
        stableStreak: nextStreak,
      }
    }
    const pruned = pruneThreadMemory(map, SUBJECT_MEMORY_CAP)
    await chrome.storage.local.set({ [KEY_SUBJECT_MEMORY]: pruned })
  })
}

export async function getThreadMemoryCounts(): Promise<{
  conversationEntries: number
  subjectEntries: number
}> {
  const [c, s] = await Promise.all([getConversationMemory(), getSubjectMemory()])
  return {
    conversationEntries: Object.keys(c).length,
    subjectEntries: Object.keys(s).length,
  }
}

// ---- Storage usage --------------------------------------------------------

// LOCAL_STORAGE_QUOTA_BYTES / LOCAL_STORAGE_WARN_THRESHOLD: see
// src/shared/constants.ts. We don't declare 'unlimitedStorage', so the
// 5 MB cap is real — folderCache + ruleHistory + tombstones can reach it.

export type StorageUsage = {
  bytesInUse: number
  quotaBytes: number
  percentUsed: number
  approaching: boolean // true once we cross the warn threshold
}

export async function getStorageUsage(): Promise<StorageUsage> {
  let bytesInUse = 0
  try {
    bytesInUse = await chrome.storage.local.getBytesInUse(null)
  } catch (e) {
    console.warn('[mail-organizer] getBytesInUse failed', e)
  }
  const percentUsed = bytesInUse / LOCAL_STORAGE_QUOTA_BYTES
  return {
    bytesInUse,
    quotaBytes: LOCAL_STORAGE_QUOTA_BYTES,
    percentUsed,
    approaching: percentUsed >= LOCAL_STORAGE_WARN_THRESHOLD,
  }
}
