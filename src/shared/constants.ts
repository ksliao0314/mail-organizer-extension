// Centralized magic numbers. One place to tune timing / caps / retention.
//
// Organized by domain:
//   - sync       : cross-machine sync engine
//   - storage    : chrome.storage local + sync quotas
//   - retention  : how long different data types live
//   - threading  : conversation + subject memory
//   - classify   : Claude API batching
//   - folders    : folderCache + folderActivity
//
// Keep VALUES here, not behaviour. If a number's meaning changes, fix the
// constant; if its USAGE changes, fix the call site.

// ---- sync engine ----------------------------------------------------------

/** Debounce window collapsing rapid local mutations into one sync push. */
export const PUSH_DEBOUNCE_MS = 5000

/**
 * Echo suppression window after a pull. chrome.storage.onChanged listeners
 * for the pull's own writes dispatch as MACROTASKS — they fire AFTER our
 * finally block runs. Without this delay the listener would schedule a
 * push that re-uploads what we just downloaded. 1 s gives margin for
 * macrotask drain under load (typical < 1 ms).
 */
export const PULL_GRACE_MS = 1000

/**
 * Debounce for remote-pull dispatch: when the other machine pushes
 * several times in a burst (rule edits, initial-scan generating
 * ai_confirmed), collapse those onChanged events into one pull.
 */
export const PULL_DEBOUNCE_MS = 200

/** Hard cap on quiesce()'s wait — beyond this a push is hung, don't block UI. */
export const QUIESCE_TIMEOUT_MS = 5000

/** Per-item ceiling chrome.storage.sync enforces is 8 KB; we target 6 KB
 *  to leave envelope + JSON-overhead headroom. */
export const CHUNK_BYTE_TARGET = 6 * 1024

/** Max chunks per category before truncation kicks in. */
export const MAX_RULE_CHUNKS = 20
export const MAX_TOMBSTONE_CHUNKS = 8
export const MAX_FOLDER_ACTIVITY_CHUNKS = 4
/**
 * Folder-dictionary chunks (schemaVersion 3). The ~88 KB byte budget can hold
 * at most ~15 chunks of folders anyway (it shrinks the folder set by dropping
 * rules before this cap bites), so 16 makes the byte budget the sole limiter
 * in practice and leaves this cap as a pure defensive backstop. ~40 distinct
 * folders fit in ~2 chunks; a firm with a folder-per-case has ample headroom.
 */
export const MAX_FOLDER_DICT_CHUNKS = 16

/**
 * Total-byte budget for OUR sync payload, well under the 100 KB hard quota.
 * doPush measures the non-rule writes (settings + meta + tombstones +
 * folderActivity) and gives compact rule rows + folder dict the remainder up
 * to this budget, dropping the lowest-priority rules if they don't fit. The
 * ~12 KB gap to 100 KB absorbs Chrome's per-item bookkeeping and leaves room
 * so a push never fails the whole batch on `kQuotaBytes`.
 */
export const SYNC_TOTAL_BUDGET_BYTES = 88 * 1024

/** Tombstone sync cap — older deletions are unlikely to be re-learned. */
export const TOMBSTONE_SYNC_CAP = 500

/** folderActivity sync cap — only the most-recent entries propagate. */
export const FOLDER_ACTIVITY_SYNC_CAP = 20

/** Schema version of the cloud envelope. Bump on breaking change. */
export const SYNC_SCHEMA_VERSION = 3

/** Local backup-snapshot rotation. */
export const MAX_BACKUPS = 5

/**
 * Two-phase closed-case retirement (2026-07-16): days a rule stays
 * soft-disabled (autoDisabledReason 'case_closed') before the sweep
 * hard-deletes it. 使用者實務: 結案後信件尾巴約 15 天 → 取 2 倍保險。
 * Phase 1 的軟停用已同步到所有機器，30 天只是等每台機器都看過。
 */
export const CASE_CLOSED_PURGE_DAYS = 30

/** Recent-pushes log carried in syncMeta for multi-machine visibility. */
export const SYNC_RECENT_PUSHES_CAP = 20

// ---- stale sweep ----------------------------------------------------------

/**
 * Stale-sweep clock (2026-11-05). Rule staleness is measured in ACTIVE days —
 * UTC days on which a batch was actually executed — not calendar days. A rule
 * can only "fail to match" on a day the inbox went through the matcher; the
 * old wall-clock predicate hard-deleted the whole learned library after ~100
 * days of simply not using the tool (vacation, a busy quarter), and re-deleted
 * any restored backup on the next wake. Only the most recent ~100 active days
 * ever matter to the predicate, so this cap is generous headroom.
 */
export const ACTIVE_DAYS_CAP = 400

/**
 * Minimum spacing between non-forced sweeps. The alarm is daily; this gate
 * stops the module-load sweep from re-running on every SW wake (the SW wakes
 * on most tab events) and from double-running on the alarm's own wake.
 */
export const SWEEP_MIN_INTERVAL_MS = 20 * 60 * 60 * 1000 // 20 h

// ---- storage --------------------------------------------------------------

/** chrome.storage.local default quota. */
export const LOCAL_STORAGE_QUOTA_BYTES = 5 * 1024 * 1024
/** Warn at this fraction of LOCAL_STORAGE_QUOTA_BYTES. */
export const LOCAL_STORAGE_WARN_THRESHOLD = 0.8

/** Max tombstones to keep locally before FIFO eviction. */
export const TOMBSTONE_CAP = 2000

/** Max rule-history events. ~300-500 B each, so 500 ≈ 250 KB. */
export const RULE_HISTORY_CAP = 500

/** Max entries in centralised SW error log (#2). */
export const ERROR_LOG_CAP = 200

// ---- retention ------------------------------------------------------------

export const SKIP_HISTORY_TTL_MS = 60 * 24 * 60 * 60 * 1000 // 60 days
export const SKIP_HISTORY_CAP = 5000

// recentlyProcessed: emails we just moved / deleted in a batch, excluded
// from the NEXT batch's candidate list. Outlook's message store is
// eventually consistent — for a short window after a move/delete the
// inbox listing can still return the just-handled email (with its now-
// dead id), so the next batch tries to re-move it and gets a 404
// ErrorItemNotFound. This ledger filters those out client-side. Short
// TTL: it only needs to outlive Outlook's propagation lag (seconds to a
// minute or two in practice), not user sessions.
export const RECENTLY_PROCESSED_TTL_MS = 15 * 60 * 1000 // 15 min
export const RECENTLY_PROCESSED_CAP = 2000

// Moved-by-us ledger (harvester provenance). Records the DESTINATION-side new
// message id of every mail WE moved, so the manual-filing harvester can exclude
// our own moves. Unlike recentlyProcessed (source-side ids, 15-min TTL to
// outlast propagation), the harvester may run days after a batch, so this
// keeps entries for weeks.
export const MOVED_BY_US_LEDGER_TTL_MS = 21 * 24 * 60 * 60 * 1000 // 21 days
export const MOVED_BY_US_LEDGER_CAP = 5000

/** Recent-activity panel data store cap. */
export const FOLDER_ACTIVITY_CAP = 200
/** Retention for folderActivity entries. */
export const FOLDER_ACTIVITY_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000 // 30 days

// ---- threading ------------------------------------------------------------

export const CONVERSATION_MEMORY_CAP = 5000
export const SUBJECT_MEMORY_CAP = 3000
export const THREAD_MEMORY_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000 // 90 days

/**
 * Threshold for conversationMemory / subjectMemory decay: after this many
 * consecutive same-folder filings on a previously-conflicted entry,
 * conflictCount decrements by 1 per extra same-folder filing.
 */
export const DECAY_AFTER_STABLE = 5

// ---- classify -------------------------------------------------------------

/** Emails per Claude API call. */
export const AI_CHUNK_SIZE = 25

// ---- folders --------------------------------------------------------------

/** folderCache freshness window. */
export const FOLDER_CACHE_TTL_MS = 24 * 60 * 60 * 1000 // 24 h

// ---- popup ---------------------------------------------------------------

/** Auto-trigger refresh in popup if last successful refresh was longer ago. */
export const POPUP_AUTO_REFRESH_STALE_MS = 30 * 60 * 1000 // 30 min
