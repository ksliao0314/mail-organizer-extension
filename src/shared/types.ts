// Outlook REST v2.0 entity types ---------------------------------------------

export type MailFolder = {
  Id: string
  DisplayName: string
  ParentFolderId?: string
  ChildFolderCount?: number
  TotalItemCount?: number
  UnreadItemCount?: number
}

export type MailFolderNode = {
  id: string
  displayName: string
  parentFolderId?: string
  path: string
  children: MailFolderNode[]
  /**
   * Set on the ROOT of Junk Email / Sent Items (2026-11-06 r2). Those two and
   * their subfolders are ordinary targets for rules, the picker and execute,
   * but never offered to the AI — see flattenAiVisibleFolders.
   */
  wellKnown?: AiHiddenFolderRole
}

export type EmailAddress = { Name?: string; Address: string }
export type Recipient = { EmailAddress: EmailAddress }

export type Email = {
  Id: string
  Subject: string
  BodyPreview: string
  /**
   * Normalised plain-text body (batch-3). Fetched on demand via
   * `getMessageBody` (UniqueBody, Prefer:text) ONLY for the AI-bucket subset —
   * emails that miss rules/thread and go to Claude. Absent on list results and
   * rule/thread-routed mail; consumers use `bodyText ?? BodyPreview`.
   */
  bodyText?: string
  /**
   * Outlook's conversation grouping — all messages in the same email
   * thread share this. Smart batching groups by this so Claude only
   * classifies the representative of each conversation.
   */
  ConversationId?: string
  From: Recipient
  ToRecipients: Recipient[]
  CcRecipients?: Recipient[]
  BccRecipients?: Recipient[]
  ReceivedDateTime: string
  ParentFolderId: string
  HasAttachments?: boolean
  IsRead?: boolean
  /** Outlook follow-up flag (旗標). Used to skip emails the user marked
   * "待處理" during classify. */
  Flag?: {
    FlagStatus?: 'NotFlagged' | 'Flagged' | 'Complete'
  }
}

// Rule engine ------------------------------------------------------------------

export type RuleType =
  | 'case_code'
  | 'domain'
  | 'compound'
  | 'subject_keyword'
  | 'sender'

export type RuleSource = 'auto_scan' | 'ai_confirmed' | 'ai_overridden' | 'user_manual'

export type Rule = {
  id: string
  type: RuleType
  signal: string
  targetFolderId: string
  targetFolderPath: string
  /**
   * Configured / declared confidence. Treat as the rule's "intent" or
   * authority signal (user_manual = 0.95, ai_confirmed = 0.7 etc.). The
   * actual track record is `overrideCount / matchCount` — call
   * `effectiveConfidence(rule)` to get a blended number.
   */
  confidence: number
  matchCount: number
  /**
   * Times this rule fired AND the user then edited the resulting PlanItem
   * before execute (i.e. they overrode the rule's suggestion). Bumped at
   * execute time via PlanItem.originalRuleId. Used to compute the rule's
   * empirical accuracy.
   */
  overrideCount?: number
  enabled: boolean
  createdAt: string
  lastUsedAt?: string
  source: RuleSource
  /**
   * Set by reconcileRulesAgainstTree when the rule's targetFolderId no longer
   * resolves against Outlook's folder tree (folder deleted or Id rotated).
   * Orphaned rules are skipped during matching and surface in the rules UI
   * with a "重新指定" prompt. Auto-cleared once the user re-targets.
   */
  orphaned?: boolean
  /**
   * Timestamp (ISO) when this rule was auto-disabled by the daily
   * sweep. Set when `enabled` flips to false via the auto-disable
   * path; cleared if the user re-enables (toggleRule unsets it).
   *
   * Surfaces in the rules UI as a "已自動休眠" badge so the user can
   * distinguish auto-disabled rules from ones they manually disabled.
   */
  autoDisabledAt?: string
  /**
   * Why the sweep disabled this rule:
   *   - 'stale': matchCount === 0 + older than 100 days, OR matchCount > 0
   *     + lastUsedAt older than 100 days.
   *   - 'high-error-rate': matchCount ≥ 20 + overrideCount/matchCount ≥ 0.5.
   *   - 'legacy_token': subject_keyword rule from the pre-2026-05-27
   *     token-tokenization design (signal is a 3–8 char Chinese fragment
   *     extracted from a subject). The new design uses the entire
   *     normalized subject; legacy token rules are over-broad (fire on
   *     unrelated mail containing the fragment) and get retired.
   *   - 'case_closed' (2026-07-16): the rule's target folder sits under a
   *     user-designated closed-case archive tree (settings.
   *     closedCasePrefixes). Phase 1 of two-phase retirement: soft-disable
   *     here propagates via sync to every machine; the sweep hard-deletes
   *     (no tombstone) after CASE_CLOSED_PURGE_DAYS — except user_manual
   *     rules, which stay soft-disabled forever. SYMMETRIC: moving the
   *     folder back OUT of the archive tree auto-re-enables the rule on
   *     the next preflight / sweep.
   *   - 'override' (2026-11-05): one-strike disable of a structural rule
   *     after the user re-targeted the mail it routed.
   *   - 'superseded' (2026-11-05): a broad rule silenced because a narrower
   *     rule now covers its traffic.
   * Cleared along with `autoDisabledAt` on user re-enable.
   * Sync wire codes are append-only (sync-codec.ts REASON_TO_CODE).
   */
  autoDisabledReason?:
    | 'stale'
    | 'high-error-rate'
    | 'legacy_token'
    | 'case_closed'
    | 'override'
    | 'superseded'
  /**
   * Logical last-edit time (epoch ms) of this rule's SEMANTIC fields — type,
   * signal, target, confidence, enabled, source, auto-disable state
   * (2026-11-05). Stamped ONLY by `mutateRules` (rules.ts) when one of those
   * fields actually changes, or when the rule is new; hit counters
   * (matchCount / overrideCount / lastUsedAt) never stamp it, and neither
   * does `orphaned` (2026-11-06: a per-device observation of this machine's
   * folder tree, not an edit). Cross-machine
   * sync merges per rule id by it: the copy with the newer editedAt supplies
   * the semantic fields (tie → cloud), so a stale copy can no longer revert a
   * newer edit. Absent on rules last written before v1.0.8 (treated as 0).
   */
  editedAt?: number
  /**
   * Accuracy-counter epoch (epoch ms, 2026-11-06). Set when the engine
   * deliberately restarts matchCount / overrideCount (revivedRule after a
   * 'high-error-rate' disable). Sync merges counters by it: the copy with the
   * newer epoch supplies matchCount / overrideCount, equal epochs take the
   * max — otherwise the other machine's pre-reset counters came straight back
   * and the next sweep re-disabled the rule. Absent = never reset.
   */
  statsResetAt?: number
  /**
   * LOCAL ONLY (2026-11-06, S3): when THIS machine's reconcile set `orphaned`
   * (epoch ms), maintained by mutateRules alongside editedAt; never synced.
   * An orphaned copy without it may carry a v1.0.8 orphan stamp (v1.0.8
   * stamped editedAt when orphaning), so the sync merge never lets it supply
   * the semantic fields (A2); with it, editedAt is a genuine edit and the
   * merge is plain last-writer-wins.
   */
  orphanedAt?: number
}

// Plan items ------------------------------------------------------------------

export type PlanAction = 'move' | 'delete' | 'new_folder' | 'skip'

/**
 * Lightweight snapshot of the rule that produced a rule-sourced PlanItem,
 * embedded so the popup can show "為什麼歸到這裡" trace without an extra
 * round-trip to fetch the full rule. Captured at classifyPreflight time;
 * if the rule changes after that, trace shows the as-of-classify state.
 */
export type MatchedRuleSummary = {
  id: string
  type: RuleType
  signal: string
  source: RuleSource
  matchCount: number
  lastUsedAt?: string
  enabled: boolean
}

export type PlanItem = {
  emailId: string
  emailSubject: string
  emailFrom: string
  /**
   * Outlook's `BodyPreview` — first ~250 chars of the email body. Already
   * fetched as part of the message list, just propagated here so the popup
   * can show inline preview on row expand without an extra round-trip.
   */
  bodyPreview?: string
  /**
   * Case identifiers extracted from the email BODY at classify time (batch-3),
   * from the 800-char bodyText — kept separate from subject identifiers so the
   * learning loop can apply the stricter body gate (subject-first, distinct===1,
   * lower confidence). Empty/absent when the subject already carried a case
   * number (body is then ignored) or none was found.
   */
  bodyCaseNumbers?: string[]
  bodyCaseCodes?: string[]
  /**
   * Outlook's `ConversationId` — used by smart batching to send only one
   * representative email per thread to Claude. Siblings inherit the AI's
   * decision, saving Claude calls (and matching the user's mental model
   * that a thread belongs to one case).
   */
  conversationId?: string
  action: PlanAction
  targetFolderId?: string
  targetFolderPath?: string
  suggestedFolderName?: string
  suggestedParentPath?: string
  confidence: number
  reason: string
  source: 'rule' | 'ai' | 'unresolved' | 'thread'
  ruleId?: string
  matchedRule?: MatchedRuleSummary
  /**
   * Source detail for `thread`-routed items — surfaces in the UI
   * trace block. `convId` means we matched by ConversationId (exact
   * thread); `subject` means we matched by normalized subject (looser
   * fallback when no convId match was available).
   */
  threadMatch?: { kind: 'convId' | 'subject'; previousFolderPath: string }
  /**
   * Set on a `thread`-routed item when a broad rule (domain / sender /
   * subject_keyword) ALSO matched and agreed on the same destination (優化
   * 2026-07). Thread precedence hides that rule, so its matchCount /
   * lastUsedAt would freeze and the 100-day stale sweep would eventually
   * hard-delete it — even though it's been silently correct all along.
   * execute credits a hit to this rule (when the item is un-edited) so a
   * consistently-agreeing rule survives.
   */
  agreeingRuleId?: string
  /**
   * Captured at rule-match time and preserved across user edits even
   * after `source` flips to 'ai'. Lets execute attribute "user overrode
   * this rule" back to the original rule, so we can track empirical
   * accuracy (overrideCount / matchCount) separately from configured
   * confidence.
   */
  originalRuleId?: string
  /**
   * True for sibling items inside a conversation thread that inherited
   * the AI's decision from the thread's representative (not classified
   * independently). UI surfaces this so the user knows why every member
   * of a 5-email thread has identical action.
   */
  threadInherited?: boolean
  /**
   * Marks rows the user explicitly edited (action change / target pick /
   * keyboard d/s). Used to gate auto-propagation to same-subject siblings —
   * we don't want to clobber an item the user already made a deliberate
   * different decision on.
   */
  userTouched?: boolean
  /**
   * True for placeholder rows of mail that was NEVER classified — the Claude
   * call failed (429 / 529 / network…) or the response omitted this email —
   * as opposed to an AI or confidence-gate "skip" decision (2026-11-05).
   * Executing a batch must not record these in skipHistory: that hid mail
   * that merely hit a transient API error from every batch for 60 days,
   * labelled 「先前選擇保留」. Unless the user explicitly chose 保留
   * (userTouched), they come back next batch.
   */
  unclassified?: boolean
  /**
   * On an `unclassified` placeholder Claude refused (2026-11-06): `email` =
   * this email (or its thread's representative) was refused on its own — or
   * refused before and not re-sent — and needs manual handling; `batch` = the
   * request carrying it was refused without being pinned on one email (tried
   * again next batch). `category` = stop_details.category, null when the API
   * gave none (UI labels: refusalCategoryLabel). Absent on every other row.
   */
  refusal?: { scope: 'email' | 'batch'; category: string | null }
  /**
   * Snapshot of what the AI originally decided, captured at classifier time
   * before any user edits. Used after execute to detect "user overrode AI
   * to point X" cases and promote them into ai_overridden rules so the same
   * mistake doesn't recur. Set only on AI-sourced items.
   */
  aiOriginalAction?: PlanAction
  aiOriginalTargetFolderPath?: string
  aiOriginalSuggestedFolderName?: string
  aiOriginalSuggestedParentPath?: string
}

// Stored state ----------------------------------------------------------------

export type Token = {
  secret: string
  expiresOn: number
}

export type Settings = {
  claudeApiKey: string
  claudeModel: string
  batchSize: number
  excludeFolderPrefixes: string[]
  /**
   * Folder-path prefixes the user marked as CLOSED-CASE archive trees
   * (2026-07-16). Distinct from excludeFolderPrefixes ("don't scan here"):
   * a prefix listed here additionally triggers two-phase rule retirement —
   * rules whose target folder moves under one of these trees soft-disable
   * with autoDisabledReason 'case_closed' (and re-enable if the folder
   * moves back out). Chosen in Options from among excludeFolderPrefixes.
   * SYNCED (not per-device): the retirement sweep predicate must see the
   * same value on every machine to converge.
   */
  closedCasePrefixes: string[]
  /**
   * Auto-gate AI plan items whose confidence falls below this threshold.
   * Gated items become `action: 'skip'` so they stay in inbox for manual
   * triage rather than triggering low-confidence moves.
   */
  aiConfidenceThreshold: number
  /**
   * When true, classify automatically excludes inbox messages that the user
   * has flagged (Outlook 「待處理」 / follow-up flag). Default true — the
   * assumption is that flagged messages are something the user is actively
   * tracking and wants to keep in inbox.
   */
  skipFlagged: boolean
  /**
   * When true (default), inject a floating launcher button into OWA pages
   * (lower-right, near Microsoft's Copilot button). Click it to open the
   * classify panel inline without leaving the inbox. Independent of the
   * Chrome toolbar icon, which always works.
   */
  showOwaFab: boolean
  /**
   * Pipeline mode — when true, the popup silently triggers a background
   * classify for the next batch as soon as the current execute finishes.
   * Saves 30-60 seconds per continuation but spends Claude tokens
   * speculatively. If the user walks away after one batch without
   * continuing, those tokens are wasted. Default false so the lawyer
   * opts in deliberately after understanding the trade-off.
   */
  prefetchNextBatch: boolean
  /**
   * Recent Activity panel filter — folders are shown only if their path
   * starts with one of these prefixes OR the leaf-name matches one of
   * `recentActivityIncludeLeafNames`. Trailing slash on prefix is
   * recommended (e.g. `<primaryRootPath>/`) so the parent itself doesn't
   * surface, just its descendants. Empty array means no prefix match —
   * only leaf-name allowlist applies.
   */
  recentActivityIncludePrefixes: string[]
  /**
   * Recent Activity panel filter — leaf folder names that are always
   * shown regardless of where in the hierarchy they sit. e.g. add
   * `文章審閱` to surface that folder regardless of its parent path.
   */
  recentActivityIncludeLeafNames: string[]
  /**
   * Law firm internal email domains (e.g. `['example.com']`). Drives:
   *   1. initial-scan: skip these domains when learning sender rules —
   *      colleagues cross-work many cases so per-folder uniqueness is
   *      misleading.
   *   2. execute.ts `chooseLearningSignal`: treat these as "unusable
   *      domain" — don't auto-generate domain rules pointing internal
   *      mail to a specific folder.
   *   3. classifier system prompt: hint that internal emails should be
   *      categorised by subject keywords, not by sender.
   *
   * Empty array → solo / no-firm-domain mode. The lawyer can still
   * classify external mail; the internal-skip / hint logic is bypassed.
   */
  internalDomains: string[]
  /**
   * Path to the lawyer's primary case-tracking root (e.g.
   * `案件`). Initial scan starts here, options-page UI uses it
   * as placeholder, the AI prompt embeds it in example folder paths.
   *
   * Empty string → user hasn't completed onboarding; UIs / scans
   * requiring a root path should prompt the user to set it instead of
   * silently using a wrong default.
   */
  primaryRootPath: string
  /**
   * Internal-email category hints fed to the AI prompt — short Chinese
   * keywords identifying the firm's common categories of internal
   * correspondence (e.g. `['工時','薪資','利衝','行政','公告']`).
   *
   * These are SUGGESTIONS for the classifier, not rules. Empty array →
   * the prompt omits the category hint entirely.
   */
  internalSubjectCategories: string[]
  /**
   * Whether classifier should include user-validated rules as few-shot
   * examples in the prompt. Default true — examples meaningfully boost
   * accuracy. Set false if rule target paths contain sensitive client
   * data the user doesn't want round-tripping through the LLM (e.g.
   * lawyer with paths like `客戶X 離婚案`). The classifier still uses
   * the user's rules for direct matching either way; only the prompt
   * mining is gated.
   */
  aiIncludeFewShotExamples: boolean
  /**
   * Manual-filing harvester (opt-in, default off). When on, the tool scans
   * case folders for mail you filed by hand in Outlook and feeds the routing
   * into thread-memory (soft, non-authoritative — surfaced as a hint, never an
   * auto-route or a rule). Local-only, never synced.
   */
  harvestEnabled?: boolean
  /**
   * Cross-machine sync via the browser's account sync (chrome.storage.sync).
   * Disabled by default — opt-in via Options UI. Sync covers:
   *   - rules where source !== 'auto_scan' (auto_scan re-derived per machine
   *     by running initial scan)
   *   - recent tombstones (capped to fit chrome.storage.sync quota)
   *   - settings excluding `claudeApiKey` (never leaves local storage)
   *
   * Edge syncs via Microsoft account; Chrome via Google account.
   * Cross-browser doesn't auto-sync (the sync backends are isolated).
   */
  syncEnabled: boolean
  /**
   * Unique-per-device id used to recognise our own writes in the
   * chrome.storage.onChanged listener (so we don't pull on echoes of
   * our own pushes). Generated lazily on first sync enable; persists
   * for the device lifetime. Never synced to cloud.
   */
  syncMachineId: string
  /**
   * ISO timestamp of last successful sync (push OR pull, whichever ran
   * last). Empty string when sync has never run. Surfaced in the UI so
   * the user can see when their state was last propagated.
   */
  lastSyncAt: string
  /**
   * Has the user dismissed (or completed) the first-run onboarding
   * wizard? Drives whether IdleScreen shows the wizard overlay.
   *   - true: wizard already shown / completed / skipped — don't show
   *   - false: show wizard when other "looks empty" signals are also true
   * Per-device because each machine's "have I been onboarded here yet"
   * answer is independent.
   */
  onboardingDismissed: boolean
}

export type Metrics = {
  moved: number
  deleted: number
  foldersCreated: number
  errors: number
}

export type FolderCache = {
  updatedAt: string
  tree: MailFolderNode[]
  /**
   * The Outlook distinguished folders resolved with `tree` (2026-11-06 r2) —
   * see WellKnownFolders. A cache without it (and without the legacy fields
   * below) predates v1.0.8 and lists every folder.
   */
  wellKnown?: WellKnownFolders
  /**
   * Set when execute found the tree out of date (2026-11-06 r2): a folder in
   * it now sits in the trash / Drafts / Outbox, or is gone. The cache is then
   * no longer fresh — the next preflight refetches it, and reconcile orphans
   * the rules that pointed there.
   */
  invalidatedAt?: string
  /**
   * v1.0.8 only — no longer written. Ids of the well-known folders whose
   * whole subtrees that build left OUT of `tree`. Refetched on first use.
   */
  excludedFolderIds?: string[]
  /**
   * The 2026-11-06 round-1 build only — no longer written (that build kept
   * the trash in `tree` and listed its subfolders here). Refetched on first
   * use.
   */
  nonFiling?: NonFilingFolders
}

/** Outlook REST distinguished folder names the folder tree treats specially. */
export type WellKnownFolderRole = 'deleteditems' | 'drafts' | 'outbox' | 'junkemail' | 'sentitems'

/** The well-known roots that stay in the tree but are hidden from the AI. */
export type AiHiddenFolderRole = Extract<WellKnownFolderRole, 'junkemail' | 'sentitems'>

export type WellKnownFolderRoot = { id: string; role: WellKnownFolderRole; displayName?: string }

/**
 * Well-known folders (2026-11-06 r2). Deleted Items (root and subtree),
 * Drafts and Outbox are left out of the tree — their subtrees are never even
 * listed. A rule targeting Deleted Items ITSELF stays valid (a deliberate
 * auto-trash rule; execute runs it as a delete); anything else missing from
 * a complete tree is orphaned, including a folder deleted in OWA (it moves
 * into the trash with the same id). Junk Email and Sent Items stay in the
 * tree as ordinary targets, their roots tagged `MailFolderNode.wellKnown`.
 */
export type WellKnownFolders = {
  roots: WellKnownFolderRoot[]
  /**
   * False when some lookup failed transiently: `roots` holds what resolved,
   * the cache is only briefly fresh (refetched by the next preflight), and
   * reconcile leaves every rule as it is.
   */
  complete: boolean
}

/** Well-known folders mail is never filed into. */
export type NonFilingRoot = 'deleteditems' | 'drafts' | 'outbox'

/**
 * Round-1 (2026-11-06) cache shape, read only to upgrade it: the trash stayed
 * in `tree`; its subfolders and Drafts / Outbox (+ subtrees) were recorded here.
 */
export type NonFilingFolders = {
  trashFolderId?: string
  byId: Record<string, NonFilingRoot>
}

/**
 * Per-folder "recent activity" record — used to populate the IdleScreen's
 * 「近日活動」quick-jump panel. Written every time an execute batch
 * successfully moves mail into the folder; also refreshable on-demand via
 * a Graph API scan (catches mail the user dragged into the folder manually,
 * bypassing the extension).
 */
export type FolderActivity = {
  folderId: string
  folderPath: string
  /** ISO timestamp of the most recent mail known to have entered this folder. */
  lastActiveAt: string
  /** Number of mails moved into this folder in the most recent execute batch. */
  recentCount: number
  /** ISO timestamp of the batch that produced recentCount. */
  lastBatchAt?: string
  /**
   * Most-recent message in the folder, captured during refresh scan so the
   * popup row can show subject + sender on hover without an extra Graph
   * roundtrip. Subject truncated to ~120 chars for storage; full subject
   * is one click away in OWA.
   */
  latestMessage?: {
    subject: string
    from: string
    receivedAt: string
  }
}

export type StoredState = {
  rules: Rule[]
  folderCache?: FolderCache
  settings: Settings
  metrics: Metrics
}

/**
 * Thread-memory entry: "the user previously filed this conversation (or
 * subject) into folder X". Pre-filter step in classifyPreflight uses
 * this to route follow-up emails (internal replies, fwd, etc.) to the
 * same case folder without falling through to AI.
 *
 * Two parallel maps share this entry shape:
 *   - ConversationFolderMap: keyed by Outlook ConversationId (exact)
 *   - SubjectFolderMap: keyed by normalized subject (looser fallback)
 *
 * `conflictCount` only on subject map — increments when the same
 * normalized subject was filed to a DIFFERENT folder than current.
 * Pre-filter ignores entries with conflictCount > 0 (subject is
 * ambiguous, let rules / AI decide instead).
 */
export type ThreadMemoryEntry = {
  folderId: string
  folderPath: string
  /** ISO timestamp of the most recent filing into this folder for this key. */
  lastFiledAt: string
  /** Lifetime count of times this key was filed (to ANY folder). */
  timesFiled: number
  /** Only on SubjectFolderMap entries. Times this subject was filed to a folder OTHER than `folderId`. */
  conflictCount?: number
  /**
   * Only on SubjectFolderMap entries. Consecutive filings to the same
   * folder without a conflict. Once it crosses a threshold, each extra
   * same-folder filing decays `conflictCount` by 1 — lets a subject
   * earn back its routing fallback after an accidental cross-filing.
   */
  stableStreak?: number
  /**
   * The folder this key routed to BEFORE the most recent conflict (優化
   * 2026-07). When a subsequent filing returns to this folder ("回巢"), we
   * treat it as the thread reverting to its established home rather than a
   * SECOND conflict — so a single legitimate cross-folder filing doesn't
   * double-penalize (A→B counts once; the A→…→A return no longer counts
   * again). Cleared once conflictCount decays back to 0.
   */
  previousFolderId?: string
  /**
   * True when EVERY filing for this key so far was derived from FOLDER
   * CONTENTS (the manual-filing harvester, or the initial-scan seed), never
   * from a plan the lawyer reviewed. Graph can't prove a human filed it, so
   * this weak-evidence class is NOT trusted for auto-route even at
   * conflictCount===0 — it only surfaces as a soft "依你既有的 Outlook 歸檔"
   * hint, until a real plan-reviewed execute filing confirms the route and
   * permanently clears this flag.
   */
  harvested?: boolean
}

/**
 * Snapshot used by the weekly digest. We store one snapshot per dismiss —
 * next digest reports the delta between current state and this snapshot.
 * "Weekly" is a soft target; if the user doesn't open the popup for 3
 * weeks, they get a 3-week summary on the next dismiss.
 */
export type WeeklyDigestState = {
  /** ISO timestamp of the last digest dismiss (or initial seed). */
  lastShownAt: string
  snapshot: {
    metrics: Metrics
    rulesCount: number
  }
}

// ---- Execute state (shared between SW and popup) --------------------------
//
// Defined here so the popup's local view (which doesn't carry `plan`) is
// derived from the same source of truth as the SW's full state. Adding
// fields here propagates to both sides; the popup uses Omit to strip what
// it doesn't need.

export type ExecuteItemStatus =
  | 'queued'
  | 'processing'
  | 'moved'
  | 'deleted'
  | 'folder_created'
  | 'skipped'
  | 'cancelled'
  | 'error'

export type ExecuteItemResult = {
  emailId: string
  subject: string
  action: PlanAction
  status: ExecuteItemStatus
  message?: string
  // Move / folder_created / deleted outcomes. Outlook returns a NEW Id after
  // the move (a delete is a move into Deleted Items since 2026-11-05); the
  // undo snapshot needs this new Id.
  newMessageId?: string
  destinationFolderId?: string
  destinationFolderPath?: string
  /**
   * Outlook answered 404 ErrorItemNotFound — the id is provably dead
   * (2026-11-05: persisted so SW-death recovery can replay the
   * recentlyProcessed ledger for it exactly like the normal path).
   */
  alreadyGone?: boolean
  /**
   * Moved into a folder outside the tree while Deleted Items couldn't be
   * identified (2026-11-06 r3) — it may have been the trash itself. The mail
   * moved (and 撤回 can bring it back), but nothing is learned, remembered
   * or logged as folder activity from it.
   */
  unverifiedTarget?: boolean
}

export type ExecuteSummary = {
  moved: number
  deleted: number
  foldersCreated: number
  skipped: number
  cancelled: number
  errors: number
}

// One rule the learning loop minted during an execute batch. Surfaced to the
// DoneScreen so the reviewer can see EXACTLY which corrections became durable
// rules (覆核 P1-3) — closing the "will my override stick?" transparency gap.
export type RuleLearnedDetail = {
  type: RuleType
  signal: string
  targetFolderPath: string
  source: RuleSource
}

export type ExecuteState = {
  inProgress: boolean
  cancelRequested: boolean
  startedAt: number
  finishedAt?: number
  total: number
  current: number
  // Original plan items, persisted so "retry failed" can rebuild without
  // asking the popup to re-send. Popup itself never reads this field.
  plan: PlanItem[]
  results: ExecuteItemResult[]
  summary: ExecuteSummary
  rulesAdded: number
  /**
   * Details of the rules minted this batch (same count as rulesAdded).
   * Optional for backward-compat with persisted states from older versions
   * that lack it. Consumed by the DoneScreen learned-rules list.
   */
  rulesAddedDetail?: RuleLearnedDetail[]
}

// Defaults are intentionally empty / generic for new installs. Existing
// users' previously-hardcoded values (example.com / 案件 / 等)
// are preserved via the migration block in `getSettings` — see
// storage.ts. The empty defaults also drive the onboarding banner: when
// `internalDomains` is empty AND `primaryRootPath` is empty, popup
// surfaces a prompt to fill them in.
export const DEFAULT_SETTINGS: Settings = {
  claudeApiKey: '',
  // New installs only (2026-11-05); a stored known id is never migrated.
  // Keep in sync with DEFAULT_CLAUDE_MODEL in claude-models.ts.
  claudeModel: 'claude-sonnet-5-5',
  batchSize: 50,
  excludeFolderPrefixes: [],
  closedCasePrefixes: [],
  aiConfidenceThreshold: 0.5,
  skipFlagged: true,
  showOwaFab: true,
  prefetchNextBatch: false,
  recentActivityIncludePrefixes: [],
  recentActivityIncludeLeafNames: [],
  internalDomains: [],
  primaryRootPath: '',
  internalSubjectCategories: [],
  aiIncludeFewShotExamples: true,
  harvestEnabled: false,
  syncEnabled: false,
  syncMachineId: '',
  lastSyncAt: '',
  onboardingDismissed: false,
}

export const DEFAULT_METRICS: Metrics = {
  moved: 0,
  deleted: 0,
  foldersCreated: 0,
  errors: 0,
}

/**
 * Records what an execute batch moved so it can be reversed within a short
 * undo window. Captured when the batch's bookkeeping runs — BEFORE the
 * executeState flips to inProgress=false (2026-11-05), so the DoneScreen's
 * one-shot snapshot read on mount always finds it.
 *
 * Deletes ARE recorded (2026-11-05): a delete is now a move into Deleted
 * Items (POST /move, DestinationId 'DeletedItems' — it used to be
 * DELETE /me/messages/{id}, which lands in Recoverable Items, not the
 * 「已刪除的郵件」 folder the UI promised), so 撤回 moves them back to the
 * Inbox like any other move; those entries carry `deleted: true`. Folder
 * creations are still not rolled back (more disruptive than leaving an
 * empty folder).
 *
 * Stored in chrome.storage.local with a chrome.alarms expiry — local
 * persistence is more robust than session storage if the SW restarts during
 * the 30s window.
 */
export type UndoSnapshot = {
  batchId: string
  createdAt: number
  expiresAt: number
  moves: Array<{
    /**
     * Outlook returns a NEW message ID after move. This is the ID we need
     * to address the message at, not the pre-move ID.
     */
    newMessageId: string
    subject: string
    destinationFolderId: string
    destinationFolderPath?: string
    /** True for a delete (moved into Deleted Items) — no learning to reverse. */
    deleted?: true
    /**
     * Learning-reversal keys (H2). Captured from the plan item so 撤回 can
     * also revert what the batch TAUGHT the system — without these, the
     * misfiling was deterministically re-proposed next batch and grew
     * stickier each cycle (memory streak + rule hit counts kept climbing).
     */
    conversationId?: string
    normalizedSubject?: string
    /** Rule that drove this move (plan source === 'rule') — or this delete,
     *  for an auto-trash rule (2026-11-06 r3). Undo counts as an override
     *  against it, feeding the empirical-accuracy demotion. */
    ruleId?: string
  }>
  /**
   * Rules minted by THIS batch's learning — AI-confirmed and (2026-11-06)
   * override. Deleted on undo (with tombstones, so the same wrong mapping
   * isn't auto-re-learned).
   */
  mintedRuleIds?: string[]
  /**
   * Auto-disabled rules THIS batch's learning switched back on — the
   * AI-confirmed and (since R7) the override learner (2026-11-06). Undo puts
   * each back to sleep with its old provenance — otherwise the revived rule
   * re-proposed the undone misfiling next batch.
   */
  revivedRules?: RevivedRuleUndo[]
  /**
   * Enabled rules THIS batch switched off (2026-11-06): the learners'
   * 'superseded' / 'override' demotions and the one-strike disable of a case
   * rule whose row was re-filed elsewhere. Undo switches each back on with
   * its pre-batch provenance. One entry per rule across both lists: a rule
   * the batch touched more than once carries its state BEFORE the batch and
   * the fingerprint AFTER it.
   */
  disabledRules?: DisabledRuleUndo[]
  /**
   * Deletes in the batch (user messaging). Since 2026-11-05 the ones with a
   * known new id are ALSO in `moves` (deleted: true) and are undoable.
   */
  deletedCount: number
  /** Folders created by the batch — NOT rolled back by undo. */
  newFolderCount: number
}

/** A learning revive, as undo needs it (UndoSnapshot.revivedRules, 2026-11-06). */
export type RevivedRuleUndo = {
  ruleId: string
  /** Provenance before the revive — what undo restores. */
  autoDisabledReason?: Rule['autoDisabledReason']
  autoDisabledAt?: string
  /** The rule's semantic state right after the revive. Undo only acts while
   *  the stored rule still matches it — an edit since then (user, sync,
   *  reconcile) wins. */
  fingerprint: string
}

/**
 * A rule the batch switched off, as undo needs it (UndoSnapshot.disabledRules,
 * 2026-11-06) — same shape as RevivedRuleUndo, but the rule was ENABLED
 * before the batch (an enabled rule may still carry the closed-case
 * exemption marker as its provenance).
 */
export type DisabledRuleUndo = RevivedRuleUndo

export const UNDO_WINDOW_MS = 30_000

/**
 * Lightweight snapshot of a rule's essential identity at a point in time —
 * used by RuleEvent history records so an event is self-contained even if
 * the underlying rule is later deleted/edited.
 */
export type RuleSnapshot = {
  type: RuleType
  signal: string
  targetFolderPath: string
  confidence: number
  source: RuleSource
  enabled: boolean
  /**
   * Captured so reconcile audit events can show the user when a folder
   * deletion / reappearance flipped this flag. Absent in older history
   * entries (treat falsy as "not orphaned").
   */
  orphaned?: boolean
}

/**
 * Audit log entry for any mutation that touched the rule library. Captured
 * by every CRUD helper in rules.ts so the user can review what changed,
 * when, and who caused it (their own action vs. the system auto-generating
 * something).
 *
 * `actor`:
 *   - 'user' — the user explicitly triggered this (options page edit /
 *     toggle / delete, plan-screen "停用此規則", rule import).
 *   - 'system' — auto-derived (initial-scan, ai_confirmed, ai_overridden,
 *     reconcile, ai_overridden's auto-demote of conflicting rules).
 *
 * The `kind` discriminates payload shape — kept as a union so consumers
 * (UI history list, future analytics) can switch on it cleanly.
 */
export type RuleEvent =
  | {
      kind: 'create'
      ruleId: string
      at: number
      actor: 'user' | 'system'
      after: RuleSnapshot
    }
  | {
      kind: 'edit'
      ruleId: string
      at: number
      actor: 'user' | 'system'
      before: RuleSnapshot
      after: RuleSnapshot
      /** Names of RuleSnapshot fields whose value changed. */
      changedFields: string[]
    }
  | {
      kind: 'toggle'
      ruleId: string
      at: number
      actor: 'user' | 'system'
      enabled: boolean
      /** Identifying signal so the UI can render without re-fetching the rule. */
      signal: string
      type: RuleType
      targetFolderPath: string
    }
  | {
      kind: 'delete'
      ruleId: string
      at: number
      actor: 'user' | 'system'
      before: RuleSnapshot
    }

/**
 * Records that the user explicitly deleted a (type, signal, target) triple
 * so we don't auto-regenerate it later. Without this, deleting an
 * AI-confirmed domain rule just makes it reappear on the next batch (the
 * AI will suggest the same target, dedup against current storage misses,
 * and the rule is re-added — user thought they deleted it).
 *
 * Match logic uses normalized signal + exact target path. If the user
 * later manually re-creates a rule with the same triple, we clear the
 * matching tombstone (CRUD ops in rules.ts).
 */
export type RuleTombstone = {
  type: RuleType
  /** Lowercased and stripped (e.g. `@company-a.example` → `company-a.example`). */
  signalNorm: string
  targetFolderPath: string
  deletedAt: number
  /**
   * v1.0.9+ (2026-11-06, S1): fingerprint (rules.ts spellingFingerprint) of
   * the deleted rule's own signal in v1.0.7's normal form (legacySignalNorm)
   * — e.g. of 「112年度訴字第304號」 under the canonical signalNorm 「112訴304」.
   * Lets the first-participation merge tell "this exact rule was deleted"
   * from a canonical collision. Every tombstone this build writes carries
   * it; older builds' never do (and ignore it on the wire). A fingerprint, not
   * the spelling itself: a compound's legacy spelling can be 3× its
   * canonical key, and tombstones compete for the sync budget.
   */
  spellingHash?: string
}
