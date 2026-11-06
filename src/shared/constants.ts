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

/**
 * Schema version of the cloud envelope. Bump on breaking change.
 * v4 (2026-11-05): per-rule editedAt, pushId-stamped chunks (torn-read
 * detection), disabled-rule stubs, compact compound / tombstone rows, wipe
 * ids. A v4 client READS v1–v4 and WRITES v4; v3 clients refuse a v4 cloud
 * (their own downgrade protection), which is the intended upgrade prompt.
 */
export const SYNC_SCHEMA_VERSION = 4

/** Oldest cloud schema a pull can still read (union-merged, then rewritten as current). */
export const SYNC_MIN_READABLE_SCHEMA_VERSION = 1

/**
 * Tombstones' first claim on the sync byte budget, as a share of
 * SYNC_TOTAL_BUDGET_BYTES (2026-11-05). Newest deletions go first; they may
 * spill past this share only into budget the rules leave unused — so a big
 * deletion history can never crowd rules out of the cloud again.
 */
export const SYNC_TOMBSTONE_BUDGET_SHARE = 0.15

/** Torn-read retry: a push's chunks can land after its meta (2026-11-05). */
export const SYNC_TORN_RETRY_BASE_MS = 2000
export const SYNC_TORN_RETRY_MAX_MS = 60_000
/** Consecutive torn reads before the failure is surfaced in lastError. */
export const SYNC_TORN_RETRY_MAX_ATTEMPTS = 6

/** Failed-push retry backoff (doubling from base, capped). */
export const SYNC_PUSH_RETRY_BASE_MS = 30_000
export const SYNC_PUSH_RETRY_MAX_MS = 30 * 60_000
/** chrome.storage.sync write-rate limits: per-minute and per-hour buckets. */
export const SYNC_RATE_LIMIT_MINUTE_RETRY_MS = 65_000
export const SYNC_RATE_LIMIT_HOUR_RETRY_MS = 10 * 60_000

/** Delay before retrying, after SW start, a push that failed in a previous SW lifetime. */
export const SYNC_STARTUP_RECONCILE_DELAY_MS = 3000

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

/**
 * Sync budget claims (2026-11-06). Deletions newer than this are "recent":
 * they may take up to SYNC_RECENT_TOMBSTONE_BUDGET_SHARE of the budget, but
 * only by displacing low-priority rule rows (never-matched auto rules,
 * disabled rules) — a 300-rule bulk delete used to leave ~200 deletions out
 * of a 15 % share while the peer kept routing mail with the deleted rules.
 */
export const SYNC_RECENT_TOMBSTONE_MS = 30 * 24 * 60 * 60 * 1000
export const SYNC_RECENT_TOMBSTONE_BUDGET_SHARE = 0.4
/**
 * Disabled rules whose disable is at most this old get a guaranteed stub
 * (2026-11-06). Every disabled rule used to get one before any enabled row was
 * placed, so a big pile of long-settled disables (legacy_token, user-disabled
 * — never purged) crowded live rules out of the cloud.
 */
export const SYNC_STUB_RECENT_MS = 30 * 24 * 60 * 60 * 1000

/**
 * Counter-only changes (matchCount / overrideCount / lastUsedAt) are pushed at
 * most once per this window (2026-11-06): every push rewrites every chunk
 * (~88 KB, each carries the pushId), and hit counters change on every batch.
 * Semantic edits, tombstones, settings and folder activity still push at once.
 */
export const SYNC_COUNTER_PUSH_COALESCE_MS = 20 * 60_000

/**
 * Torn-cloud recovery (2026-11-06): after this many consecutive torn reads of
 * the SAME cloud pushId, spanning at least SYNC_TORN_RECOVERY_MIN_MS, a push
 * may rewrite the cloud from local (after a local backup). A torn snapshot
 * that never heals used to lock every other machine out of pushing at all.
 */
export const SYNC_TORN_RECOVERY_MIN_READS = 3
export const SYNC_TORN_RECOVERY_MIN_MS = 60_000

/**
 * Release of the canonical tombstone keys (v1.0.8, 2026-11-05): from then on
 * signalNorm is normalizeSignal output (court-case forms folded — a deleted
 * 「112年度訴字第304號」 is stored as 「112訴304」). A tombstone deleted before
 * it was written by v1.0.7 or older in the deleted rule's own spelling — the
 * only ones the first-participation merge may read as a mere canonical
 * collision when that spelling differs from a user_manual rule's (S1).
 */
export const CANONICAL_TOMBSTONE_KEYS_SINCE_MS = Date.parse('2026-11-05T00:00:00.000Z')

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

/**
 * chrome.storage.local quota FALLBACK — only used when the runtime doesn't
 * expose `chrome.storage.local.QUOTA_BYTES` (2026-11-05). Chrome 114+ (and
 * Chromium Edge) default to 10 MB; this used to say 5 MB, so the usage meter
 * reported twice the real fill and warned at 4 MB. With the `unlimitedStorage`
 * permission there is no enforced quota at all (getStorageUsage → null).
 */
export const LOCAL_STORAGE_QUOTA_BYTES = 10 * 1024 * 1024
/** Warn at this fraction of the effective local quota. */
export const LOCAL_STORAGE_WARN_THRESHOLD = 0.8
/** How many of the largest chrome.storage.local keys getStorageUsage lists. */
export const STORAGE_USAGE_TOP_KEYS = 10

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
// keeps entries for weeks. MUST be ≥ THREAD_MEMORY_MAX_AGE_MS (2026-11-05):
// at 21 days a 22-day-old own move was re-harvested as a "manual filing"
// against a thread-memory entry that was still alive (90 days).
export const MOVED_BY_US_LEDGER_TTL_MS = 90 * 24 * 60 * 60 * 1000 // 90 days
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

/**
 * Valid range for settings.batchSize (2026-11-05: enforced in storage on every
 * read AND every merged write, not just in the SW setSettings handler — sync
 * pull / backup restore bypass the handler). 0 meant every batch came back
 * empty. Ceiling raised 200 → 500 (2026-11-05, 放寬並加速) together with the
 * parallel AI chunks / parallel moves / faster Outlook rate below, so a
 * 500-email batch finishes in minutes rather than a quarter of an hour.
 */
export const BATCH_SIZE_MIN = 1
export const BATCH_SIZE_MAX = 500

/**
 * Claude calls in flight at once while classifying (2026-11-05). Each call
 * carries one AI_CHUNK_SIZE chunk; results are still merged into the plan in
 * chunk order. Drops to 1 for the rest of the run after any 429 / 529, so an
 * account with a low rate-limit tier degrades to the old sequential pace
 * instead of retrying in a storm.
 */
export const AI_PARALLEL_CHUNKS = 3

/**
 * Refusal isolation (2026-11-06). When Claude refuses a chunk, the chunk is
 * bisected — sequentially, inside its own lane — to find the email(s) that
 * trip the safety classifier; the rest are classified normally. Extra calls
 * are capped per chunk (one culprit in a 25-email chunk takes 10) and per run,
 * so a refusal triggered by something every request shares (folder names,
 * rule text) can't multiply the calls.
 */
export const AI_REFUSAL_ISOLATION_MAX_CALLS = 12
export const AI_REFUSAL_ISOLATION_MAX_CALLS_PER_RUN = 36

/**
 * Emails Claude refused, remembered (local, never synced) so later batches
 * don't send them again: they take no batch slot and, unless a rule or thread
 * memory routes them, are left for manual handling (counted next to the
 * flagged mail). Per model — switching models retries them.
 */
export const REFUSED_EMAILS_TTL_MS = 7 * 24 * 60 * 60 * 1000 // 7 days
export const REFUSED_EMAILS_CAP = 500

/**
 * A refusal of the prompt every request shares (2026-11-06): once a one-email
 * probe with that prompt was refused too, later runs with the same model and
 * prompt prefix don't probe / bisect again for this long — unless a call is
 * answered first, which clears it.
 */
export const REFUSAL_SUSPICION_TTL_MS = 6 * 60 * 60 * 1000 // 6 hours

/**
 * Move / delete requests in flight at once during execute (2026-11-05).
 * new_folder items still run first and one at a time (later moves may target
 * them). Below MAX_CONCURRENT_REQUESTS (4) on purpose: the token is OWA's
 * own, so the Outlook tab's traffic shares the same per-mailbox budget.
 */
export const EXECUTE_PARALLEL_ITEMS = 3

/**
 * Outlook REST pacing (2026-11-05): steady requests/second and burst for the
 * shared TokenBucket. Was 1/s (burst 10): once the bucket became a real FIFO
 * reservation queue, every fan-out (25 body fetches per AI chunk, the folder
 * tree's sibling GETs) ran ~10× slower than before. 8/s sustained is 4,800
 * requests / 10 min — under half of Exchange's ~10,000 per mailbox — and the
 * per-mailbox concurrency cap (4 in flight, outlook-api.ts) stays the guard.
 */
export const OUTLOOK_REQUESTS_PER_SEC = 8
export const OUTLOOK_REQUEST_BURST = 40

/**
 * Inbox scan sizing for one batch (2026-11-05). Pages hold 100–200 messages
 * (200 = the size already proven against Outlook REST: batch 100 + 100
 * ghosts; the 100 floor keeps a small batch from walking 20 tiny pages when
 * the inbox top is all kept mail) and the scan follows @odata.nextLink; the
 * total scanned scales with the batch (3×, so kept / flagged mail can't
 * starve a big batch) within a hard cap.
 */
export const INBOX_PAGE_SIZE_MIN = 100
export const INBOX_PAGE_SIZE_MAX = 200
export const INBOX_SCAN_MIN_MESSAGES = 500
export const INBOX_SCAN_MAX_MESSAGES = 2000

// ---- folders --------------------------------------------------------------

/** folderCache freshness window. */
export const FOLDER_CACHE_TTL_MS = 24 * 60 * 60 * 1000 // 24 h

/**
 * Freshness of a tree built from a PARTIAL well-known folder lookup
 * (2026-11-06 r2): short, so the next preflight refetches it, yet long enough
 * that one popup / Options session doesn't re-walk the tree on every call.
 */
export const FOLDER_CACHE_INCOMPLETE_TTL_MS = 60 * 1000

/**
 * A cached folder tree older than this is double-checked live before execute
 * files mail into one of its folders (2026-11-06 r2): a folder deleted in OWA
 * after the fetch is still in the tree, yet sits in 已刪除的郵件. A
 * just-fetched tree is trusted as is.
 */
export const FOLDER_TREE_VERIFY_AFTER_MS = 10 * 60 * 1000

// ---- popup ---------------------------------------------------------------

/** Auto-trigger refresh in popup if last successful refresh was longer ago. */
export const POPUP_AUTO_REFRESH_STALE_MS = 30 * 60 * 1000 // 30 min
