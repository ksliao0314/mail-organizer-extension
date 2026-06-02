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

/** Tombstone sync cap — older deletions are unlikely to be re-learned. */
export const TOMBSTONE_SYNC_CAP = 500

/** folderActivity sync cap — only the most-recent entries propagate. */
export const FOLDER_ACTIVITY_SYNC_CAP = 20

/** Schema version of the cloud envelope. Bump on breaking change. */
export const SYNC_SCHEMA_VERSION = 2

/** Local backup-snapshot rotation. */
export const MAX_BACKUPS = 5

/** Recent-pushes log carried in syncMeta for multi-machine visibility. */
export const SYNC_RECENT_PUSHES_CAP = 20

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
