// Execute flow.
//
// Drives a PlanItem[] through Outlook REST + persists progress to
// chrome.storage.session so the popup can poll and resume across reopens.
//
// Per-item handling:
//   move        → POST /messages/{id}/move (uses targetFolderId, or resolves from path);
//                 Deleted Items ITSELF as the target runs as a delete (2026-11-06 r2)
//   delete      → POST /messages/{id}/move into Deleted Items (2026-11-05 —
//                 was DELETE, which lands in Recoverable Items, not 已刪除的郵件)
//   new_folder  → POST createChildFolder (or top) + POST /messages/{id}/move
//   skip        → no-op
//
// Rule feedback:
//   - rule-sourced success → bumpRuleHit(ruleId)
//   - ai-sourced successful move → auto-create domain rule (confidence 0.7,
//     source: 'ai_confirmed') if no equivalent rule already covers it.
//     Skips senders on settings.internalDomains (firm-internal mail —
//     colleagues send across many cases; domain-routing is misleading).
//
// Post-loop bookkeeping (ledgers, thread memory, rule hits / overrides,
// learning, skipHistory, metrics, undo) lives in finalizeExecuteBatch, which
// the SW-death recovery path replays from the persisted plan + results
// (2026-11-05).

import {
  findFolderByPath,
  flattenFolderTree,
  joinFolderPath,
  OutlookApi,
  OutlookError,
  OutlookTokenUnavailableError,
} from '@/shared/outlook-api'
import {
  addRulesFilteringTombstones,
  applyConfidenceCap,
  bumpRuleHits,
  bumpRuleOverrides,
  caseSignalsForLearning,
  courtCaseSignal,
  decodeCompound,
  deleteRule,
  encodeCompound,
  extractSubjectSignal,
  isAnchorNarrowingCompound,
  isConflictRetired,
  isGenericProviderDomain,
  isRevivableAutoDisabled,
  listRules,
  mutateRules,
  newRule,
  normalizeSignal,
  revivedRule,
  ruleAnchorKeys,
  tombstoneKey,
} from '@/shared/rules'
import {
  addFolderToCache,
  invalidateFolderCache,
  addToMovedByUsLedger,
  addToRecentlyProcessed,
  type MovedByUsEntry,
  appendBatchStat,
  addToSkipHistory,
  bumpMetrics,
  clearUndoSnapshot,
  getFolderCache,
  type MailboxIdentity,
  getRuleTombstones,
  getSettings,
  getSubjectMemory,
  getUndoSnapshot,
  markThreadMemoryConflicts,
  recordActiveDay,
  recordConversationFilings,
  recordFolderActivityFromBatch,
  recordRuleEvents,
  recordSubjectFilings,
  restoreSubjectMemoryEntries,
  setUndoSnapshot,
} from '@/shared/storage'
// Exactly-once finalize counting (2026-11-06) reads these back.
import { getBatchStats, getConversationMemory, getMetrics } from '@/shared/storage'
import { EXECUTE_PARALLEL_ITEMS, FOLDER_TREE_VERIFY_AFTER_MS } from '@/shared/constants'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from '@/shared/normalize'
import {
  type DisabledRuleUndo,
  type ExecuteItemResult,
  type ExecuteItemStatus,
  type ExecuteState,
  type FolderCache,
  type MailFolder,
  type MailFolderNode,
  type NonFilingRoot,
  type PlanItem,
  type RevivedRuleUndo,
  type Rule,
  type RuleEvent,
  type RuleLearnedDetail,
  type RuleSnapshot,
  type ThreadMemoryEntry,
  type UndoSnapshot,
  UNDO_WINDOW_MS,
} from '@/shared/types'
import { holdKeepAlive, releaseKeepAlive } from './keep-alive'
import { getAccountCheckedIdentity } from './mailbox-identity'

function isAbortError(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError'
}

// In-memory controller for the currently-running execute. Survives within a
// single service-worker lifetime; SW restart kills both the execute Promise
// AND this controller together, so no consistency issue.
let currentController: AbortController | null = null

// The live batch's state object + its throttled writer (2026-11-05). The
// loop reads the cancel flag from memory instead of re-reading storage per
// item; requestCancel flips it here AND persists it.
let runningState: ExecuteState | null = null
let runningWriter: ExecuteStateWriter | null = null

// Synchronous start guard (2026-11-05, CLAUDE.md invariant 6). The old
// `await isExecuteRunning()` check-then-act let two overlapping starts both
// pass before either persisted inProgress — repro: a 3-item plan sent 6
// POST /move and ended "moved 0, skipped 3". Set before the first await; SW
// idle-kill resets it, which is fine — the persisted inProgress flag (plus
// recoverStaleExecuteState) covers restarts.
let executeClaimed = false
// Same idea for 撤回: a double trigger (click + `u` key) must not move the
// snapshot's messages twice, and undo must not interleave with a batch.
let undoInFlight = false
// Concurrent recovery calls (module load + onInstalled/onStartup) share one
// run — replaying the bookkeeping twice would double-count rule hits.
let recoveryInFlight: Promise<{ recovered: boolean; stuck: number }> | null = null

// Local helper type — not part of the persisted ExecuteState, just an
// internal carrier for the result of executing one item.
type ItemOutcome = {
  status: ExecuteItemStatus
  newMessageId?: string
  destinationFolderId?: string
  destinationFolderPath?: string
  /**
   * Optional human-readable note attached on non-fatal outcomes —
   * e.g. "訊息已不在原位置(已成功移動或被其他規則處理)" when a 404
   * on move was treated as skipped rather than a hard error.
   * Surfaced into ExecuteItemResult.message.
   */
  message?: string
  /**
   * True when Outlook answered 404 ErrorItemNotFound — POSITIVE
   * confirmation that this id is dead (already moved / deleted).
   * Distinct from the uncertain network soft-skip (which gives no
   * confirmation either way): already-gone ids ARE recorded into the
   * recentlyProcessed ledger so they stop reappearing in the next
   * batch while Outlook's inbox listing lags; uncertain ones are NOT,
   * so a genuinely-unmoved email can still be retried.
   */
  alreadyGone?: boolean
  /** Moved into a destination that couldn't be verified — see ExecuteItemResult. */
  unverifiedTarget?: boolean
}

// Re-export for backwards compatibility with existing call sites.
export type { ExecuteItemResult, ExecuteItemStatus, ExecuteState } from '@/shared/types'

const STATE_KEY = 'executeState'

export async function getExecuteState(): Promise<ExecuteState | null> {
  const r = await chrome.storage.session.get(STATE_KEY)
  return (r[STATE_KEY] as ExecuteState | undefined) ?? null
}

async function saveExecuteState(state: ExecuteState): Promise<void> {
  await chrome.storage.session.set({ [STATE_KEY]: state })
}

export async function clearExecuteState(): Promise<void> {
  await chrome.storage.session.remove([STATE_KEY, FINALIZE_KEY, IN_FLIGHT_KEY])
}

// ---- In-flight marker (2026-11-06) ---------------------------------------
//
// The whole-state write is throttled (below), so after an SW death the
// persisted row of an item whose move was on the wire could still read
// 'queued' — recovery said 「此筆未開始」 for mail that had in fact moved, and
// skipped its bookkeeping. With EXECUTE_PARALLEL_ITEMS workers, several items
// are in flight at once. So before EVERY non-idempotent request (move,
// delete-as-move, folder create) the item's index goes into a tiny separate
// session record — awaited, so it is durable before the request leaves.
// Recovery reads a marked row that never reached a final status as
// "may have completed". Kept until the batch's final rows are persisted.

const IN_FLIGHT_KEY = 'executeInFlight'

type InFlightRecord = { startedAt: number; indexes: number[] }

class InFlightMarker {
  private readonly indexes = new Set<number>()
  private readonly persisted = new Set<number>()
  /** A write that hasn't snapshotted `indexes` yet — later marks can ride on it. */
  private pending: Promise<void> | null = null
  private chain: Promise<void> = Promise.resolve()

  constructor(private readonly startedAt: number) {}

  /** Resolves once a record containing `i` is stored. Never rejects. */
  mark(i: number): Promise<void> {
    if (this.persisted.has(i)) return Promise.resolve()
    this.indexes.add(i)
    if (!this.pending) {
      const write = this.chain.then(async () => {
        this.pending = null
        const snapshot = [...this.indexes]
        try {
          const rec: InFlightRecord = { startedAt: this.startedAt, indexes: snapshot }
          await chrome.storage.session.set({ [IN_FLIGHT_KEY]: rec })
          for (const x of snapshot) this.persisted.add(x)
        } catch (e) {
          console.warn('[mail-organizer] in-flight marker write failed (non-fatal)', e)
        }
      })
      this.pending = write
      this.chain = write
    }
    return this.pending
  }

  /** Drop the record once every row's final status is persisted. */
  async clear(): Promise<void> {
    if (this.indexes.size === 0) return
    await this.chain
    await clearInFlightRecord(this.startedAt)
  }
}

async function readInFlightIndexes(startedAt: number): Promise<Set<number>> {
  try {
    const rec = (await chrome.storage.session.get(IN_FLIGHT_KEY))[IN_FLIGHT_KEY] as InFlightRecord | undefined
    if (rec && rec.startedAt === startedAt && Array.isArray(rec.indexes)) {
      return new Set(rec.indexes.filter((x) => typeof x === 'number'))
    }
  } catch {
    /* unknown → nothing marked */
  }
  return new Set()
}

/**
 * Drop the finalize / in-flight records of any batch but `startedAt`
 * (2026-11-06 r3) — the previous batch's, once this one has started.
 */
async function clearStaleBatchRecords(startedAt: number): Promise<void> {
  try {
    const r = await chrome.storage.session.get([FINALIZE_KEY, IN_FLIGHT_KEY])
    const stale = [FINALIZE_KEY, IN_FLIGHT_KEY].filter((k) => {
      const rec = r[k] as { startedAt?: unknown } | undefined
      return !!rec && rec.startedAt !== startedAt
    })
    if (stale.length > 0) await chrome.storage.session.remove(stale)
  } catch {
    /* best effort: readers ignore records of another batch anyway */
  }
}

async function clearInFlightRecord(startedAt: number): Promise<void> {
  try {
    const rec = (await chrome.storage.session.get(IN_FLIGHT_KEY))[IN_FLIGHT_KEY] as InFlightRecord | undefined
    if (rec && rec.startedAt === startedAt) await chrome.storage.session.remove(IN_FLIGHT_KEY)
  } catch {
    /* ignore */
  }
}

const SW_DIED_IN_FLIGHT_MESSAGE =
  'Service worker 中斷時這封正在處理:可能已完成,請到 Outlook 確認;若仍在收件匣,下次歸類會再處理。'

// Throttled executeState persistence (2026-11-05). The loop used to do one
// session read + two whole-state writes PER ITEM, each re-serialising the
// full plan — O(n²) bytes per batch. Now writes are coalesced to at most one
// per STATE_WRITE_INTERVAL_MS (trailing edge, so a long item still gets its
// 'processing' row out within the interval) plus explicit flushes at the
// points that matter (start, loop end, final). Field names are unchanged, so
// getExecuteState consumers see the same shape.
const STATE_WRITE_INTERVAL_MS = 500

class ExecuteStateWriter {
  private lastWriteAt = 0
  private timer: ReturnType<typeof setTimeout> | null = null
  private chain: Promise<void> = Promise.resolve()

  constructor(private readonly state: ExecuteState) {}

  /** Ask for a write soon (throttled, trailing edge). */
  touch(): void {
    if (this.timer) return
    const wait = this.lastWriteAt + STATE_WRITE_INTERVAL_MS - Date.now()
    if (wait <= 0) {
      void this.flush().catch((e) =>
        console.warn('[mail-organizer] executeState write failed (non-fatal)', e),
      )
      return
    }
    this.timer = setTimeout(() => {
      this.timer = null
      void this.flush().catch((e) =>
        console.warn('[mail-organizer] executeState write failed (non-fatal)', e),
      )
    }, wait)
  }

  /** Write now (after any write already in flight); rejects on failure. */
  flush(): Promise<void> {
    if (this.timer) {
      clearTimeout(this.timer)
      this.timer = null
    }
    this.lastWriteAt = Date.now()
    // Writes are chained so they land in order; each serialises the state as
    // of when it actually runs (chrome.storage snapshots at call time).
    const write = this.chain.then(() => saveExecuteState(this.state))
    this.chain = write.catch(() => {})
    return write
  }

  dispose(): void {
    if (this.timer) {
      clearTimeout(this.timer)
      this.timer = null
    }
  }
}

/**
 * Detect an `inProgress = true` state that has no live executor and recover.
 * Called on SW module load to handle MV3 idle-kill mid-run. Items in
 * `processing` become `error` (ambiguous — the operation may have partially
 * applied server-side); items in `queued` become `cancelled` (never started,
 * shouldn't be counted as a failure). Without this, `isExecuteRunning()`
 * would block the next batch forever.
 *
 * 2026-11-05: the dead batch's post-loop bookkeeping (recentlyProcessed +
 * moved-by-us ledgers, thread memory, rule hits / overrides, learning,
 * skipHistory, folder activity, metrics, active day, undo snapshot) is now
 * replayed through the SAME finalizeExecuteBatch the normal path uses, from
 * the persisted plan + results (which carry newMessageId / destination).
 * Previously only recentlyProcessed, metrics and the active day survived an
 * idle-kill.
 */
export function recoverStaleExecuteState(): Promise<{ recovered: boolean; stuck: number }> {
  if (!recoveryInFlight) {
    recoveryInFlight = doRecoverStaleExecuteState().finally(() => {
      recoveryInFlight = null
    })
  }
  return recoveryInFlight
}

async function doRecoverStaleExecuteState(): Promise<{ recovered: boolean; stuck: number }> {
  // A batch running in THIS worker is alive, not stale.
  if (executeClaimed) return { recovered: false, stuck: 0 }
  const state = await getExecuteState()
  if (!state?.inProgress || executeClaimed) return { recovered: false, stuck: 0 }

  let cancelledCount = 0
  let errorCount = 0
  let uncertainCount = 0
  // A row whose move / delete / folder create was already sent (in-flight
  // marker, 2026-11-06) MAY have completed even if it still reads 'queued' —
  // the throttled state write just hadn't landed. Same honest soft skip as a
  // lost response in the live loop (and, like it, kept out of
  // recentlyProcessed so mail still in the inbox is re-offered).
  const inFlight = await readInFlightIndexes(state.startedAt)
  const results = state.results.map((r, i) => {
    if ((r.status === 'processing' || r.status === 'queued') && inFlight.has(i)) {
      uncertainCount++
      return { ...r, status: 'skipped' as const, message: SW_DIED_IN_FLIGHT_MESSAGE }
    }
    if (r.status === 'processing') {
      errorCount++
      return { ...r, status: 'error' as const, message: 'Service worker 中斷、批次未完成' }
    }
    if (r.status === 'queued') {
      cancelledCount++
      return { ...r, status: 'cancelled' as const, message: 'Service worker 中斷、此筆未開始' }
    }
    return r
  })
  const stuck = errorCount + cancelledCount + uncertainCount
  const recovered: ExecuteState = {
    ...state,
    cancelRequested: false,
    results,
    summary: {
      ...state.summary,
      skipped: state.summary.skipped + uncertainCount,
      cancelled: state.summary.cancelled + cancelledCount,
      errors: state.summary.errors + errorCount,
    },
  }
  const plan = Array.isArray(state.plan) ? state.plan : []

  // Safe against double-apply: the normal path only flips inProgress=false
  // AFTER finalizeExecuteBatch completes, each finished step is recorded in
  // the finalize marker (so a death mid-bookkeeping replays only the rest),
  // and this recovery only fires on inProgress===true.
  try {
    await finalizeExecuteBatch(recovered, plan)
  } catch (e) {
    // Never let a bookkeeping bug strand the executor in inProgress=true.
    console.warn('[mail-organizer] recovery bookkeeping failed (non-fatal)', e)
  }

  await saveExecuteState({
    ...recovered,
    inProgress: false,
    finishedAt: state.finishedAt ?? Date.now(),
  })
  await clearFinalizeMarker(state.startedAt)
  await clearInFlightRecord(state.startedAt)
  console.warn(
    `[mail-organizer] recovered stuck executeState: ${errorCount} error, ${cancelledCount} cancelled, ${uncertainCount} may have completed`,
  )
  return { recovered: true, stuck }
}

export async function requestCancel(): Promise<void> {
  // In-memory first (2026-11-05): the loop checks this object's flag, so the
  // cancel takes effect at the next item boundary without a storage read;
  // the in-flight request aborts via the controller. Persisting through the
  // batch's own writer keeps writes ordered (no stale write can land after).
  const live = runningState
  if (live && live.inProgress) {
    live.cancelRequested = true
    currentController?.abort()
    const writer = runningWriter
    await (writer ? writer.flush() : saveExecuteState(live)).catch((e) =>
      console.warn('[mail-organizer] persisting cancel flag failed (non-fatal)', e),
    )
    return
  }
  // No batch in this worker: only touch a persisted state that still claims
  // to run (a dead batch awaiting recovery) — never clobber anything else.
  const cur = await getExecuteState()
  if (cur && cur.inProgress) {
    await saveExecuteState({ ...cur, cancelRequested: true })
  }
  currentController?.abort()
}

/**
 * Synchronous: a batch / retry has claimed the executor, or a 撤回 is
 * replaying, in THIS worker — startExecute would refuse right now.
 */
export function isExecuteBusy(): boolean {
  return executeClaimed || undoInFlight
}

export async function isExecuteRunning(): Promise<boolean> {
  if (executeClaimed) return true
  const state = await getExecuteState()
  // Re-check after the await — a start may have claimed meanwhile.
  return executeClaimed || !!state?.inProgress
}

function initialState(plan: PlanItem[]): ExecuteState {
  return {
    inProgress: true,
    cancelRequested: false,
    startedAt: Date.now(),
    total: plan.length,
    current: 0,
    plan,
    results: plan.map((p) => ({
      emailId: p.emailId,
      subject: p.emailSubject,
      action: p.action,
      status: 'queued' as const,
    })),
    summary: { moved: 0, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
    rulesAdded: 0,
  }
}

export async function startExecute(
  plan: PlanItem[],
  tree: MailFolderNode[],
  api: OutlookApi,
): Promise<void> {
  const { done } = await launchExecute(plan, tree, api)
  await done
}

/**
 * The batch's initial state could not be persisted (2026-11-06 r2) — nothing
 * ran. Its message is user-facing.
 */
export class ExecuteStartError extends Error {
  constructor(cause: unknown) {
    super(
      `無法寫入執行狀態(瀏覽器儲存空間出錯),這批沒有開始、信件都沒有移動。請再試一次;若持續發生,請重新開啟瀏覽器。(${
        cause instanceof Error ? cause.message : String(cause)
      })`,
    )
    this.name = 'ExecuteStartError'
  }
}

/**
 * Claims the executor and starts the batch; resolves once its initial state
 * is durably persisted (2026-11-06 r2), with `done` settling when the batch
 * finishes. The SW handler awaits this before answering: the old
 * fire-and-forget start answered ok even when that first session write
 * failed, and the popup then polled a null state forever. Rejects — with
 * nothing run and the claim released — on a claim conflict or an
 * ExecuteStartError.
 */
export async function launchExecute(
  plan: PlanItem[],
  tree: MailFolderNode[],
  api: OutlookApi,
): Promise<{ done: Promise<void> }> {
  // Claim synchronously, BEFORE the first await (invariant 6).
  if (executeClaimed) throw new Error('已有批次執行中')
  if (undoInFlight) throw new Error('撤回進行中,請稍候再執行')
  executeClaimed = true
  let handedOff = false
  try {
    // Storage check kept for the SW-restart case (a persisted inProgress
    // from a previous worker that recovery hasn't swept yet).
    if ((await getExecuteState())?.inProgress) throw new Error('已有批次執行中')
    handedOff = true
  } finally {
    if (!handedOff) executeClaimed = false
  }
  let markStarted: () => void = () => {}
  const started = new Promise<void>((resolve) => (markStarted = resolve))
  const done = runClaimedExecute(plan, tree, api, markStarted)
  // Whichever settles first: the state is durable, or the batch failed
  // before that point (its rejection propagates here).
  await Promise.race([started, done])
  return { done }
}

export type RetryFailedOutcome =
  | { ok: true; started: number }
  | {
      ok: false
      code: 'ALREADY_RUNNING' | 'NO_PREV' | 'NO_FAILED' | 'NO_TREE' | 'EXECUTE_START_FAILED'
      message: string
    }

/**
 * 「重試失敗項目」(2026-11-05: moved here from the SW handler so the
 * executor is claimed synchronously BEFORE the handler's first await — two
 * overlapping retries used to both pass `isExecuteRunning()`, and the second
 * one's clearExecuteState could wipe the first one's freshly started state).
 * Resolves once the retry batch has STARTED (fire-and-forget, like
 * startExecute from the SW); the popup polls getExecuteState.
 */
export async function startRetryFailed(
  api: OutlookApi,
  loadTree: () => Promise<MailFolderNode[] | undefined>,
): Promise<RetryFailedOutcome> {
  if (executeClaimed || undoInFlight) {
    return { ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' }
  }
  executeClaimed = true
  let handedOff = false
  try {
    const prev = await getExecuteState()
    if (prev?.inProgress) {
      return { ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' }
    }
    if (!prev || !prev.plan) {
      return { ok: false, code: 'NO_PREV', message: '沒有可重試的批次' }
    }
    const failedItems: PlanItem[] = []
    for (let i = 0; i < prev.results.length; i++) {
      if (prev.results[i]?.status === 'error' && prev.plan[i]) failedItems.push(prev.plan[i]!)
    }
    if (failedItems.length === 0) {
      return { ok: false, code: 'NO_FAILED', message: '沒有失敗項目' }
    }
    const tree = await loadTree()
    if (!tree) {
      return { ok: false, code: 'NO_TREE', message: '資料夾樹不存在，請先重新分類' }
    }
    // The previous batch's undo snapshot and records are let go by
    // runClaimedExecute once the retry's first write has landed
    // (2026-11-06 r3) — clearing them up front lost the failed rows (and
    // 撤回) whenever that write then failed.
    handedOff = true
    // Answer only once the retry's state is persisted (2026-11-06 r2) — see
    // launchExecute.
    let markStarted: () => void = () => {}
    const started = new Promise<void>((resolve) => (markStarted = resolve))
    const done = runClaimedExecute(failedItems, tree, api, markStarted)
    try {
      await Promise.race([started, done])
    } catch (e) {
      return {
        ok: false,
        code: 'EXECUTE_START_FAILED',
        message: e instanceof Error ? e.message : String(e),
      }
    }
    void done.catch((err) => console.error('[mail-organizer] retryFailed failed', err))
    return { ok: true, started: failedItems.length }
  } finally {
    if (!handedOff) executeClaimed = false
  }
}

/**
 * The batch itself. Caller MUST hold the executor claim; this releases it.
 */
async function runClaimedExecute(
  plan: PlanItem[],
  tree: MailFolderNode[],
  api: OutlookApi,
  /** Called once the initial state is persisted (2026-11-06 r2). */
  onStarted: () => void = () => {},
): Promise<void> {
  let state: ExecuteState
  try {
    state = initialState(plan)
  } catch (e) {
    // A malformed plan (2026-11-06 r2): nothing ran — release the claim
    // instead of leaving the executor claimed for this worker's lifetime.
    executeClaimed = false
    throw e
  }
  const controller = new AbortController()
  currentController = controller
  runningState = state
  const writer = new ExecuteStateWriter(state)
  runningWriter = writer
  const inFlight = new InFlightMarker(state.startedAt)
  let held = false

  try {
    // The first write decides whether this batch starts at all. Only once it
    // is durable is the PREVIOUS batch let go (2026-11-06 r3): a failed start
    // used to have dismissed its 撤回 already (and, for 重試, wiped the
    // failed rows the user was retrying) while the message said to retry.
    try {
      await writer.flush()
    } catch (e) {
      throw new ExecuteStartError(e)
    }
    // Dismiss any still-unexpired undo snapshot from the PREVIOUS batch
    // (audit P3): captureUndoSnapshot early-returns when the new batch has
    // zero moves, so without this a quick zero-move follow-up batch left
    // the old banner alive — and clicking 撤回 would revert the WRONG
    // batch's moves. Its finalize / in-flight records go too (they are
    // keyed by startedAt; this batch hasn't written its own yet).
    await dismissUndo().catch(() => {})
    await clearStaleBatchRecords(state.startedAt)
    onStarted()

    // Execute new_folder items before everything else. This lets the popup
    // expose pending folders as targets for later `move` items in the same plan
    // — by the time those moves run, the real folder has been spliced into the
    // in-memory tree, so path-based resolution finds it.
    // state.results stays in original index order so the UI is unaffected.
    const indexOrder: number[] = []
    for (let i = 0; i < plan.length; i++) {
      if (plan[i]!.action === 'new_folder') indexOrder.push(i)
    }
    for (let i = 0; i < plan.length; i++) {
      if (plan[i]!.action !== 'new_folder') indexOrder.push(i)
    }

    await holdKeepAlive()
    held = true
    const guard = new DestinationGuard(tree, api)
    // Circuit breaker (2026-11-06 r2): set by the first item whose request
    // could not get an Outlook token (no / frozen OWA tab, signed out…). No
    // request can go out, so the batch stops right there — each item used to
    // wait out its own ~10 s token timeout (≈ 28 minutes of error rows for a
    // 500-item batch). The rows that never ran get the same message and stay
    // retryable (「重試失敗項目」).
    let tokenOutage: string | null = null
    let tokenOutageError: OutlookTokenUnavailableError | null = null

    // One item; true = stop the batch (cancel requested / aborted / no token).
    const runOne = async (i: number): Promise<boolean> => {
      // Cancellation check between items — in memory (2026-11-05); the
      // in-flight call below is interrupted by the controller.
      if (state.cancelRequested || controller.signal.aborted) {
        state.results[i] = { ...state.results[i]!, status: 'cancelled' }
        state.summary.cancelled++
        return true
      }
      if (tokenOutage) return true // left 'queued' → marked below with the outage message

      const item = plan[i]!
      state.current = i
      state.results[i] = { ...state.results[i]!, status: 'processing' }
      writer.touch()

      // Right before each non-idempotent request (2026-11-06 r3): an item
      // already past runOne's entry check when another worker tripped the
      // breaker must not send its POST anyway.
      const beforeSend = async (): Promise<void> => {
        if (tokenOutageError) throw tokenOutageError
        await inFlight.mark(i)
        if (tokenOutageError) throw tokenOutageError
      }
      try {
        const outcome = await executeItem(item, tree, api, controller.signal, guard, beforeSend)
        const { status, newMessageId, destinationFolderId, destinationFolderPath, message } = outcome
        state.results[i] = {
          ...state.results[i]!,
          status,
          newMessageId,
          destinationFolderId,
          destinationFolderPath,
          // Propagate the explanatory note (set on 404-as-skipped path,
          // possibly other graceful-degradation cases later).
          ...(message ? { message } : {}),
          // 404 ErrorItemNotFound = positive confirmation this id is dead.
          // Persisted (not just tallied) so both the normal bookkeeping and
          // SW-death recovery feed it to the recentlyProcessed ledger —
          // without that, a ghost entry Outlook keeps listing loops forever.
          ...(outcome.alreadyGone ? { alreadyGone: true } : {}),
          ...(outcome.unverifiedTarget ? { unverifiedTarget: true } : {}),
        }

        if (status === 'moved') state.summary.moved++
        else if (status === 'deleted') state.summary.deleted++
        else if (status === 'folder_created') {
          state.summary.foldersCreated++
          state.summary.moved++ // also counts as a move
        } else if (status === 'skipped') state.summary.skipped++
      } catch (e) {
        if (isAbortError(e)) {
          state.results[i] = { ...state.results[i]!, status: 'cancelled' }
          state.summary.cancelled++
          return true
        }
        // F5 (2026-06-03): a non-idempotent op (move/delete) that failed
        // with `uncertain` — the request reached Outlook and MAY have
        // committed server-side before the response was lost (network
        // blip / 5xx / cancel while the POST was on the wire). outlook-api
        // deliberately does NOT retry these (retrying would 404 on the
        // now-stale id). Marking them as a hard `error` is wrong: it shows a
        // scary red row AND makes the item a 重試 candidate; marking a
        // mid-POST cancel 'cancelled' is wrong too when the move committed
        // (2026-11-05). Soft skip with an honest "needs confirmation" note —
        // the email is either at its destination or still in inbox
        // (re-classifiable next run); both are safe.
        if (e instanceof OutlookError && e.uncertain) {
          state.results[i] = {
            ...state.results[i]!,
            status: 'skipped',
            message: uncertainOutcomeMessage(item, e),
          }
          state.summary.skipped++
          writer.touch()
          return false
        }
        if (e instanceof OutlookTokenUnavailableError) {
          tokenOutageError ??= e
          tokenOutage ??= tokenOutageMessage(e)
          state.results[i] = { ...state.results[i]!, status: 'error', message: tokenOutage }
          state.summary.errors++
          writer.touch()
          return true
        }
        const message =
          e instanceof OutlookError
            ? `[${e.status}] ${e.message}`
            : e instanceof Error
            ? e.message
            : String(e)
        state.results[i] = { ...state.results[i]!, status: 'error', message }
        state.summary.errors++
      }

      writer.touch()
      return false
    }

    // new_folder items first and one at a time (later moves may target the
    // folders they create); then everything else with up to
    // EXECUTE_PARALLEL_ITEMS moves / deletes in flight (2026-11-05, 放寬並加速).
    // Each worker pulls the next index, so results land in their own rows
    // whatever order they finish in; the first stop ends every worker at its
    // next item boundary, and never-started rows become 'cancelled' below.
    const firstOther = indexOrder.findIndex((i) => plan[i]!.action !== 'new_folder')
    const folderPhase = firstOther < 0 ? indexOrder : indexOrder.slice(0, firstOther)
    const parallelPhase = firstOther < 0 ? [] : indexOrder.slice(firstOther)
    let stopped = false
    for (const i of folderPhase) {
      if (await runOne(i)) {
        stopped = true
        break
      }
    }
    let nextItem = 0
    const worker = async (): Promise<void> => {
      while (!stopped && nextItem < parallelPhase.length) {
        if (await runOne(parallelPhase[nextItem++]!)) stopped = true
      }
    }
    await Promise.all(
      Array.from(
        { length: Math.max(1, Math.min(EXECUTE_PARALLEL_ITEMS, parallelPhase.length)) },
        worker,
      ),
    )

    // Mark any never-started items (after break) as cancelled — or, after a
    // token outage the user didn't cancel, as retryable errors that say why.
    for (let i = 0; i < state.results.length; i++) {
      if (state.results[i]!.status === 'queued') {
        if (tokenOutage && !state.cancelRequested) {
          state.results[i] = { ...state.results[i]!, status: 'error', message: tokenOutage }
          state.summary.errors++
        } else {
          state.results[i] = { ...state.results[i]!, status: 'cancelled' }
          state.summary.cancelled++
        }
      }
    }
    // Persist the final per-item results BEFORE bookkeeping: if the worker
    // dies during it, recovery replays from exactly these rows. Every row is
    // final now, so the in-flight record has done its job.
    await writer.flush()
    await inFlight.clear()

    // Everything the DoneScreen reads (learned-rule details, the undo
    // snapshot, recentlyProcessed for peekNextBatch) is written BEFORE
    // inProgress flips to false (2026-11-05): the popup switches to the
    // DoneScreen on the first poll that sees inProgress=false and reads the
    // undo snapshot exactly once on mount — the old order (flag first,
    // snapshot ~last) made some batches show no 撤回 banner.
    await finalizeExecuteBatch(state, plan)

    state.inProgress = false
    state.finishedAt = Date.now()
    await writer.flush()
    await clearFinalizeMarker(state.startedAt)
  } catch (e) {
    state.inProgress = false
    state.finishedAt = Date.now()
    // A failed first write changed nothing: the previous batch's state, its
    // 撤回 and (for 重試) its failed rows all stay as they were.
    if (!(e instanceof ExecuteStartError)) await saveExecuteState(state).catch(() => {})
    throw e
  } finally {
    writer.dispose()
    currentController = null
    runningState = null
    runningWriter = null
    executeClaimed = false
    if (held) await releaseKeepAlive()
  }
}

/** The one message a token outage leaves on every row it stopped (2026-11-06 r2). */
function tokenOutageMessage(e: OutlookTokenUnavailableError): string {
  const reason = e.reason.trim().replace(/[。.！!\s]+$/u, '')
  return `無法取得 Outlook 授權,這批已停止,這封信沒有移動${reason ? `(${reason})` : ''}。處理好 Outlook 分頁後按「重試失敗項目」繼續。`
}

function uncertainOutcomeMessage(item: PlanItem, e: OutlookError): string {
  const where = item.action === 'delete' ? '移到「已刪除的郵件」' : '移動'
  const tail = '請到 Outlook 確認;若仍在收件匣,下次歸類會再處理。'
  return e.aborted
    ? `已取消,但取消時這封正在${where},可能已完成。${tail}`
    : `網路中斷,這封可能已成功${where}、也可能仍在收件匣。${tail}`
}

// ---- Post-loop bookkeeping (shared by the normal path and recovery) ------
//
// Every step derives its input from (plan, results) alone, so recovery can
// replay a killed batch exactly. Steps are non-fatal; each completed step is
// recorded in a small session marker so a worker death DURING bookkeeping
// replays only the remaining steps (no double-counted hits / filings / stats).

type FinalizeStep =
  | 'rules'
  | 'ledgers'
  | 'convMemory'
  | 'subjectMemory'
  | 'learning'
  | 'skipHistory'
  | 'folderActivity'
  | 'metrics'
  | 'undo'

type LearnedOutput = {
  added: number
  /** Rules both learners minted (the override learner's too since 2026-11-06) — deleted by undo. */
  mintedRuleIds: string[]
  details: RuleLearnedDetail[]
  /** Rules the learning revived (both learners since R7) — reversed by undo (2026-11-06). */
  revived?: RevivedRuleUndo[]
  /** Rules the learning switched off ('superseded' / 'override') — reversed by undo (2026-11-06). */
  disabled?: DisabledRuleUndo[]
}

/**
 * What ONE learner's rule write adds / revives / switches off, persisted in
 * the finalize marker under the rules lock right BEFORE the write
 * (2026-11-06, R8). New rule ids are random, so their presence in the
 * library proves the write landed; so does a revived / switched-off rule no
 * longer in its pre-write state. See runLearning / landedIntent.
 */
type LearnerIntent = {
  added: Array<{ id: string; detail: RuleLearnedDetail }>
  revived: RevivedRuleUndo[]
  disabled?: DisabledRuleUndo[]
}

/** Counter values before a counting write + this batch's increments (2026-11-06). */
type CounterIntent = { pre: Record<string, number>; inc: Record<string, number> }

type FinalizeMarker = {
  startedAt: number
  done: FinalizeStep[]
  learned?: LearnedOutput
  /**
   * Subject memory as it was BEFORE this batch's filings, for this batch's
   * subject keys only (null = absent) — persisted before they are recorded
   * and dropped once learning has run (2026-11-06). Makes the subjectMemory
   * step + birth gate replay-safe: see finalizeExecuteBatch.
   */
  subjectPrior?: Record<string, ThreadMemoryEntry | null>
  /** Per counting write — see applyCountsOnce. */
  intents?: Record<string, CounterIntent>
  /** `at` of this batch's stats row, recorded before it is appended. */
  statAt?: string
  /** Per learner, what its write adds / revives — see LearnerIntent (R8).
   *  Dropped once `learned` is recorded. */
  learningIntent?: { confirmed?: LearnerIntent; override?: LearnerIntent }
  /**
   * Case rules the 'rules' step's one-strike disable switched off, recorded
   * before its write (2026-11-06) and kept for the undo step. A replay whose
   * write never landed re-disables and replaces the entries of those rules.
   */
  ruleDisables?: DisabledRuleUndo[]
}

const FINALIZE_KEY = 'executeFinalize'

async function readFinalizeMarker(startedAt: number): Promise<FinalizeMarker> {
  try {
    const r = await chrome.storage.session.get(FINALIZE_KEY)
    const m = r[FINALIZE_KEY] as FinalizeMarker | undefined
    if (m && m.startedAt === startedAt && Array.isArray(m.done)) {
      return {
        startedAt,
        done: [...m.done],
        ...(m.learned ? { learned: m.learned } : {}),
        ...(m.subjectPrior ? { subjectPrior: m.subjectPrior } : {}),
        ...(m.intents && typeof m.intents === 'object' ? { intents: { ...m.intents } } : {}),
        ...(typeof m.statAt === 'string' ? { statAt: m.statAt } : {}),
        ...(m.learningIntent && typeof m.learningIntent === 'object'
          ? { learningIntent: { ...m.learningIntent } }
          : {}),
        ...(Array.isArray(m.ruleDisables) ? { ruleDisables: [...m.ruleDisables] } : {}),
      }
    }
  } catch {
    /* fall through to a fresh marker */
  }
  return { startedAt, done: [] }
}

/**
 * Exactly-once counter bumps across a worker death (2026-11-06). A step's
 * counting write and the marker's `done` are two writes; dying between them
 * replayed the step and counted the batch twice (rule hits / overrides,
 * metrics, timesFiled). Now the counters' current values and this batch's
 * increments are stored in the marker BEFORE the write. A replay compares:
 * every counter still at its recorded value → the write never landed → apply
 * the recorded increments; anything moved → it landed (or can't be told
 * apart) → skip. Undercounting a death-window batch beats double-counting.
 */
async function applyCountsOnce(
  marker: FinalizeMarker,
  key: string,
  inc: ReadonlyMap<string, number>,
  read: () => Promise<ReadonlyMap<string, number>>,
  apply: (inc: Map<string, number>) => Promise<unknown>,
): Promise<void> {
  const recorded = marker.intents?.[key]
  if (recorded) {
    const now = await read()
    const untouched = Object.keys(recorded.inc).every((k) => !now.has(k) || now.get(k) === (recorded.pre[k] ?? 0))
    if (!untouched) return
    await apply(new Map(Object.entries(recorded.inc)))
    return
  }
  const positive = [...inc].filter(([, n]) => n > 0)
  if (positive.length === 0) return
  const before = await read()
  marker.intents = {
    ...marker.intents,
    [key]: {
      pre: Object.fromEntries(positive.map(([k]) => [k, before.get(k) ?? 0])),
      inc: Object.fromEntries(positive),
    },
  }
  await writeFinalizeMarker(marker)
  await apply(new Map(positive))
}

async function writeFinalizeMarker(marker: FinalizeMarker): Promise<void> {
  try {
    await chrome.storage.session.set({ [FINALIZE_KEY]: marker })
  } catch (e) {
    console.warn('[mail-organizer] finalize marker write failed (non-fatal)', e)
  }
}

async function clearFinalizeMarker(startedAt: number): Promise<void> {
  try {
    const r = await chrome.storage.session.get(FINALIZE_KEY)
    const m = r[FINALIZE_KEY] as FinalizeMarker | undefined
    if (m && m.startedAt === startedAt) await chrome.storage.session.remove(FINALIZE_KEY)
  } catch {
    /* ignore */
  }
}

function isFiled(r: ExecuteItemResult | undefined): boolean {
  return r?.status === 'moved' || r?.status === 'folder_created'
}

/**
 * A filing the batch may learn from and remember (2026-11-06 r3): not one
 * into a destination that couldn't be verified (unverifiedTarget — it may
 * have been the trash itself).
 */
function isLearnableFiling(r: ExecuteItemResult | undefined): boolean {
  return isFiled(r) && r?.unverifiedTarget !== true
}

/**
 * A rule-routed row that ran as a delete — a delete row from an auto-trash
 * rule (2026-11-06 r3), or a 'move' whose target was Deleted Items itself
 * (r2). Not a filing (nothing learned or remembered), but the rule fired:
 * its hit counts, and 撤回 counts against it.
 */
function isRuleRoutedDelete(item: PlanItem | undefined, r: ExecuteItemResult | undefined): boolean {
  return r?.status === 'deleted' && item?.source === 'rule' && !!item.ruleId
}

/** Normalized subjects of this batch's filings — every key the subjectMemory
 *  step may write and the birth gate may read. */
function batchSubjectKeys(plan: readonly PlanItem[], results: readonly ExecuteItemResult[]): Set<string> {
  const keys = new Set<string>()
  for (let i = 0; i < plan.length; i++) {
    if (!isLearnableFiling(results[i])) continue
    const key = normalizeSubject(plan[i]?.emailSubject ?? '')
    if (key.length >= MIN_NORMALIZED_SUBJECT_LEN) keys.add(key)
  }
  return keys
}

function presentEntries(prior: Record<string, ThreadMemoryEntry | null>): SubjectMemoryMap {
  const out: SubjectMemoryMap = {}
  for (const [k, v] of Object.entries(prior)) if (v) out[k] = v
  return out
}

async function finalizeExecuteBatch(state: ExecuteState, plan: PlanItem[]): Promise<void> {
  const marker = await readFinalizeMarker(state.startedAt)
  const done = new Set(marker.done)
  const results = state.results
  let mintedRuleIds = marker.learned?.mintedRuleIds ?? []
  let revivedRules = marker.learned?.revived ?? []
  let disabledRules = marker.learned?.disabled ?? []
  if (marker.learned) {
    state.rulesAdded = marker.learned.added
    state.rulesAddedDetail = marker.learned.details
  }

  const step = async (name: FinalizeStep, fn: () => Promise<void>): Promise<void> => {
    if (done.has(name)) return
    try {
      await fn()
    } catch (e) {
      console.warn(`[mail-organizer] batch bookkeeping "${name}" failed (non-fatal)`, e)
    }
    done.add(name)
    marker.done = [...done]
    await writeFinalizeMarker(marker)
  }

  // Order matters: overrides before hits (F9, inside flushRuleFeedback);
  // thread memory before learning (the birth gate reads this batch's filing).
  await step('rules', () => flushRuleFeedback(plan, results, marker))
  await step('ledgers', () => flushLedgers(plan, results))
  await step('convMemory', () => flushConversationMemory(plan, results, marker))
  // Birth gate (2026-11-05): generateAiConfirmedRules needs subject memory as
  // it was BEFORE this batch's filings. The pre-batch values of this batch's
  // keys go into the marker BEFORE they are recorded (2026-11-06): a worker
  // death after recordSubjectFilings wrote but before the step was marked
  // done used to replay with a snapshot that already held this batch —
  // filings counted twice and a full_subject rule minted on first sight. The
  // replay now puts those keys back to the pre-batch values and records once
  // more. A marker without them (written by an older build) falls back to
  // the learner subtracting this batch's own filings.
  let priorSubjectMemory: SubjectMemoryMap | undefined
  if (marker.subjectPrior) {
    priorSubjectMemory = presentEntries(marker.subjectPrior)
    if (!done.has('subjectMemory')) {
      await restoreSubjectMemoryEntries(marker.subjectPrior).catch((e) =>
        console.warn('[mail-organizer] subject memory replay restore failed (non-fatal)', e),
      )
    }
  } else if (!done.has('subjectMemory')) {
    const current = await getSubjectMemory().catch(() => undefined)
    if (current) {
      const prior: Record<string, ThreadMemoryEntry | null> = {}
      for (const key of batchSubjectKeys(plan, results)) prior[key] = current[key] ?? null
      marker.subjectPrior = prior
      await writeFinalizeMarker(marker)
      priorSubjectMemory = presentEntries(prior)
    }
  }
  await step('subjectMemory', () => flushSubjectMemory(plan, results))
  await step('learning', async () => {
    const out = await runLearning(plan, results, priorSubjectMemory, marker)
    state.rulesAdded = out.added
    state.rulesAddedDetail = out.details
    mintedRuleIds = out.mintedRuleIds
    revivedRules = out.revived ?? []
    disabledRules = out.disabled ?? []
    marker.learned = out
    // No reader after learning; keep later writes small.
    delete marker.subjectPrior
    delete marker.learningIntent
  })
  await step('skipHistory', () => flushSkipHistory(plan, results))
  await step('folderActivity', () => flushFolderActivity(plan, results))
  await step('metrics', () => flushMetricsAndStats(state, plan, marker))
  // Capture undo snapshot — only when there's actually something to undo.
  // Folder creates aren't reversed here (folder cleanup would amplify
  // undo's blast radius beyond what users expect from "撤回"). Rule state
  // changes in the order they happened: the 'rules' step's one-strike
  // disables, then the learners' revives (AI-confirmed before override),
  // then the override learner's disables (see foldRuleStateChanges).
  await step('undo', () =>
    captureUndoSnapshot(state, plan, mintedRuleIds, [
      ...(marker.ruleDisables ?? []).map((rec) => ({ wasEnabled: true, rec })),
      ...revivedRules.map((rec) => ({ wasEnabled: false, rec })),
      ...disabledRules.map((rec) => ({ wasEnabled: true, rec })),
    ]),
  )
}

/**
 * Did the user send a rule-routed email to a DIFFERENT folder than the rule
 * targets (2026-11-05)? Only that is evidence the rule was wrong. Deleting
 * or keeping the email is not — the old tally counted deletes, and the
 * one-strike structural recovery then disabled a case rule because the user
 * deleted one notification it had routed.
 */
function isRetargetAwayFromRule(
  rule: Rule | undefined,
  item: PlanItem,
  res: ExecuteItemResult,
): boolean {
  if (!rule) return false
  const destId = res.destinationFolderId
  const ruleTargetId = sanitizeRuleTargetFolderId(rule.targetFolderId)
  if (destId && ruleTargetId) return destId !== ruleTargetId
  const destPath = res.destinationFolderPath ?? finalTargetPath(item)
  if (!destPath || !rule.targetFolderPath) return false // can't tell → no penalty
  return destPath !== rule.targetFolderPath
}

function ruleOverrideCounts(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  rulesById: ReadonlyMap<string, Rule>,
): Map<string, number> {
  const overrides = new Map<string, number>()
  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]
    const res = results[i]
    if (!item || !res || !isFiled(res)) continue
    if (item.source === 'rule' || !item.originalRuleId) continue
    if (!isRetargetAwayFromRule(rulesById.get(item.originalRuleId), item, res)) continue
    overrides.set(item.originalRuleId, (overrides.get(item.originalRuleId) ?? 0) + 1)
  }
  return overrides
}

async function flushRuleFeedback(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  marker: FinalizeMarker,
): Promise<void> {
  const rules = await listRules()
  const rulesById = new Map(rules.map((r) => [r.id, r]))
  // Rule overrides: a rule fired at classify time but the user re-targeted
  // the resulting PlanItem before execute. Detected via item.source !==
  // 'rule' with item.originalRuleId still present (PlanRow.tsx preserves it
  // on edit via {...item} spread). Feeds the rule's empirical accuracy.
  const overrides = ruleOverrideCounts(plan, results, rulesById)
  const hits = new Map<string, number>()
  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]
    const res = results[i]
    if (!item || !res || !(isLearnableFiling(res) || isRuleRoutedDelete(item, res))) continue
    if (item.source === 'rule' && item.ruleId) {
      // Accumulated and flushed once (one mutateRules for the whole batch).
      hits.set(item.ruleId, (hits.get(item.ruleId) ?? 0) + 1)
    }
    if (!isFiled(res)) continue // an auto-trash rule's hit, nothing more
    // Thread-shadowed rule keep-alive (優化 2026-07). A broad rule that
    // agreed with the thread route (agreeingRuleId) gets a hit credited
    // — but only when the user DIDN'T edit the row (an edit means the
    // rule/thread target was wrong, so don't reward it). Prevents the
    // stale sweep from hard-deleting a silently-correct rule.
    if (item.source === 'thread' && item.agreeingRuleId && !item.userTouched) {
      hits.set(item.agreeingRuleId, (hits.get(item.agreeingRuleId) ?? 0) + 1)
    }
  }

  // F9 (2026-06-03): flush OVERRIDES before HITS. bumpRuleHits's
  // confidence-promotion gate reads `overrideCount / matchCount`; applying
  // overrides first means the gate sees this batch's corrections. Both are
  // applied exactly once per batch even across a worker death (2026-11-06,
  // applyCountsOnce); the structural disable is idempotent by itself.
  const ruleCounts = (field: 'matchCount' | 'overrideCount') => async () =>
    new Map((await listRules()).map((r) => [r.id, r[field] ?? 0] as const))
  if (overrides.size > 0 || marker.intents?.ruleOverrides) {
    try {
      await applyCountsOnce(marker, 'ruleOverrides', overrides, ruleCounts('overrideCount'), bumpRuleOverrides)
    } catch (e) {
      console.warn('[mail-organizer] bumpRuleOverrides flush failed (non-fatal)', e)
    }
    try {
      // What it switches off goes into the marker BEFORE the write, for 撤回
      // (2026-11-06): a replay finds those rules already off and records
      // nothing — the entries of a write that never landed are replaced.
      await disableStructuralOverriddenRules([...overrides.keys()], async (records) => {
        const ids = new Set(records.map((r) => r.ruleId))
        marker.ruleDisables = [
          ...(marker.ruleDisables ?? []).filter((r) => !ids.has(r.ruleId)),
          ...records,
        ]
        await writeFinalizeMarker(marker)
      })
    } catch (e) {
      console.warn('[mail-organizer] structural override recovery failed (non-fatal)', e)
    }
  }
  if (hits.size > 0 || marker.intents?.ruleHits) {
    try {
      await applyCountsOnce(marker, 'ruleHits', hits, ruleCounts('matchCount'), bumpRuleHits)
    } catch (e) {
      console.warn('[mail-organizer] bumpRuleHits flush failed (non-fatal)', e)
    }
  }
}

/**
 * Fast structural-override recovery (優化 2026-07). A structural identifier
 * (case_code, or a court-case-shaped subject/compound rule) can't
 * legitimately map to two folders — so ONE re-target is sufficient evidence
 * the rule's target is wrong; disable it immediately. user_manual rules are
 * exempt (sacred).
 *
 * 2026-11-05: written through mutateRules with autoDisabledReason 'override'
 * + autoDisabledAt. The old reason-less toggleRule(false) looked exactly like
 * a USER disable: it never revived and blocked relearning of the same case.
 *
 * `beforeWrite` (2026-11-06) receives what undo needs to switch them back on,
 * under the rules lock before the write.
 */
async function disableStructuralOverriddenRules(
  ids: string[],
  beforeWrite?: (records: DisabledRuleUndo[]) => Promise<void>,
): Promise<void> {
  const idSet = new Set(ids)
  const now = Date.now()
  const events: RuleEvent[] = []
  await mutateRules(async (rules) => {
    const records: DisabledRuleUndo[] = []
    const next = rules.map((r) => {
      if (!idSet.has(r.id) || !r.enabled || r.source === 'user_manual') return r
      const structural =
        r.type === 'case_code' ||
        ((r.type === 'subject_keyword' || r.type === 'compound') && isStructuralCourtCaseRule(r))
      if (!structural) return r
      events.push({
        kind: 'toggle',
        ruleId: r.id,
        at: now,
        actor: 'system',
        enabled: false,
        type: r.type,
        signal: r.signal,
        targetFolderPath: r.targetFolderPath,
      })
      const off: Rule = {
        ...r,
        enabled: false,
        autoDisabledReason: 'override' as const,
        autoDisabledAt: new Date(now).toISOString(),
      }
      records.push(stateUndoRecord(r, off))
      return off
    })
    if (records.length === 0) return { next: rules, result: undefined }
    if (beforeWrite) await beforeWrite(records)
    return { next, result: undefined }
  })
  if (events.length > 0) {
    await recordRuleEvents(events).catch((e) =>
      console.warn('[mail-organizer] recordRuleEvents (override disable) failed', e),
    )
  }
}

async function flushLedgers(plan: PlanItem[], results: ExecuteItemResult[]): Promise<void> {
  // Record the emails we actually moved/deleted into the recently-
  // processed ledger so the NEXT batch's inbox fetch filters them out.
  // Outlook's store is eventually consistent: for a short window after a
  // move/delete the inbox listing can still return the just-handled email
  // with its now-dead id, and the next batch would try to re-move it and
  // hit 404 ErrorItemNotFound. We record only DEFINITE outcomes (moved /
  // deleted / folder_created) plus ids Outlook confirmed dead via 404
  // (alreadyGone — re-recording refreshes their TTL so a long-lagging
  // listing keeps them excluded). The uncertain-network soft-skip is left
  // out so a genuinely-unmoved email can still be retried next batch.
  const processedIds = results
    .filter(
      (r) =>
        r.status === 'moved' ||
        r.status === 'deleted' ||
        r.status === 'folder_created' ||
        r.alreadyGone === true,
    )
    .map((r) => r.emailId)
    .filter(Boolean)
  if (processedIds.length > 0) {
    try {
      await addToRecentlyProcessed(processedIds)
    } catch (e) {
      console.warn('[mail-organizer] addToRecentlyProcessed failed (non-fatal)', e)
    }
  }

  // Moved-by-us ledger (harvester provenance): the DESTINATION-side new id
  // of every mail we filed. Only while the harvester is enabled (2026-11-05)
  // — it's the ledger's only reader, and at a 90-day TTL it is not small.
  let harvestEnabled = false
  try {
    harvestEnabled = (await getSettings()).harvestEnabled === true
  } catch {
    /* treat as disabled */
  }
  if (!harvestEnabled) return
  const movedAt = Date.now()
  const entries: MovedByUsEntry[] = []
  for (let i = 0; i < plan.length; i++) {
    const res = results[i]
    if (!isFiled(res) || !res!.newMessageId || !res!.destinationFolderId) continue
    entries.push({
      newMessageId: res!.newMessageId,
      destinationFolderId: res!.destinationFolderId,
      conversationId: plan[i]?.conversationId,
      movedAt,
    })
  }
  if (entries.length > 0) {
    try {
      await addToMovedByUsLedger(entries)
    } catch (e) {
      console.warn('[mail-organizer] addToMovedByUsLedger failed (non-fatal)', e)
    }
  }
}

// Thread memory writeback. Record convId + normalized subject for every
// successful move into a real folder. We do this even for thread-sourced
// rows so the entry's timesFiled keeps climbing. destinationFolderId comes
// from the move result because Outlook assigns a NEW id after move — using
// item.targetFolderId would record the source-side id and pre-filter
// wouldn't match next batch.
async function flushConversationMemory(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  marker: FinalizeMarker,
): Promise<void> {
  const filings: Array<{ convId: string; folderId: string; folderPath: string }> = []
  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]
    const res = results[i]
    if (!item?.conversationId || !isLearnableFiling(res)) continue
    if (!res!.destinationFolderId || !res!.destinationFolderPath) continue
    filings.push({
      convId: item.conversationId,
      folderId: res!.destinationFolderId,
      folderPath: res!.destinationFolderPath,
    })
  }
  if (filings.length === 0) return
  // Once per batch (2026-11-06): a replay used to file the batch twice
  // (timesFiled / stableStreak).
  const perConversation = new Map<string, number>()
  for (const f of filings) perConversation.set(f.convId, (perConversation.get(f.convId) ?? 0) + 1)
  await applyCountsOnce(
    marker,
    'convMemory',
    perConversation,
    async () => {
      const mem = await getConversationMemory()
      return new Map(Object.entries(mem).map(([k, e]) => [k, e.timesFiled] as const))
    },
    () => recordConversationFilings(filings),
  )
}

async function flushSubjectMemory(plan: PlanItem[], results: ExecuteItemResult[]): Promise<void> {
  const filings: Array<{ normalizedSubject: string; folderId: string; folderPath: string }> = []
  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]
    const res = results[i]
    if (!item || !isLearnableFiling(res)) continue
    if (!res!.destinationFolderId || !res!.destinationFolderPath) continue
    const normalized = normalizeSubject(item.emailSubject ?? '')
    if (normalized.length < MIN_NORMALIZED_SUBJECT_LEN) continue
    filings.push({
      normalizedSubject: normalized,
      folderId: res!.destinationFolderId,
      folderPath: res!.destinationFolderPath,
    })
  }
  if (filings.length > 0) await recordSubjectFilings(filings)
}

// Generate ai_confirmed + ai_overridden rules from successful runs —
// both non-fatal. ai_confirmed captures "AI was right, lock it in";
// ai_overridden captures "user changed AI's verdict, learn the new
// mapping AND demote any existing rule that pointed to the wrong
// target". They're mutually exclusive per item via wasUserOverride.
// Conflict prevention lives in chooseLearningSignal (2026-05-27 redesign).
//
// Exactly-once across a worker death (2026-11-06, R8): each learner hands
// its write's plan (LearnerIntent) to the marker under the rules lock BEFORE
// writing. A replay whose learner write already landed rebuilds that
// learner's output from it instead of re-running — the re-run dedups against
// its own landed rules, so the minted ids / revives (撤回) and the learned
// list (DoneScreen) used to come back empty. A planned write that never
// landed is simply re-run.
async function runLearning(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  priorSubjectMemory: SubjectMemoryMap | undefined,
  marker: FinalizeMarker,
): Promise<LearnedOutput> {
  // A filing into an unverifiable destination (maybe the trash itself)
  // teaches nothing (2026-11-06 r3): the learners see it as not filed.
  results = results.map((r) => (r.unverifiedTarget ? { ...r, status: 'skipped' as const } : r))
  let added = 0
  // Everything THIS batch's learning did to the library, for 撤回 (H2
  // learning reversal): the rules both learners minted — deleted — and the
  // ones they revived or switched off — put back as they were (2026-11-06:
  // the override learner's mints and disables too).
  const mintedRuleIds: string[] = []
  const revived: RevivedRuleUndo[] = []
  const disabled: DisabledRuleUndo[] = []
  const details: RuleLearnedDetail[] = []
  const persistIntent =
    (learner: 'confirmed' | 'override') =>
    async (intent: LearnerIntent): Promise<void> => {
      marker.learningIntent = { ...marker.learningIntent, [learner]: intent }
      await writeFinalizeMarker(marker)
    }
  const library = marker.learningIntent ? await listRules().catch(() => null) : null
  const landed = (intent: LearnerIntent | undefined): LearnerIntent | null =>
    intent && library ? landedIntent(intent, library) : null
  try {
    const replay = landed(marker.learningIntent?.confirmed)
    if (replay) {
      added += replay.added.length
      mintedRuleIds.push(...replay.added.map((a) => a.id))
      revived.push(...replay.revived)
      details.push(...replay.added.map((a) => a.detail))
    } else {
      const gen = await generateAiConfirmedRules(plan, results, {
        priorSubjectMemory,
        onIntent: persistIntent('confirmed'),
      })
      added += gen.added
      mintedRuleIds.push(...gen.addedIds)
      revived.push(...gen.revivedRules)
      details.push(...gen.details)
    }
  } catch (e) {
    console.warn('[mail-organizer] generateAiConfirmedRules failed (non-fatal)', e)
  }
  try {
    const replay = landed(marker.learningIntent?.override)
    if (replay) {
      added += replay.added.length
      mintedRuleIds.push(...replay.added.map((a) => a.id))
      revived.push(...replay.revived)
      disabled.push(...(replay.disabled ?? []))
      details.push(...replay.added.map((a) => a.detail))
    } else {
      const r = await generateAiOverrideRules(plan, results, { onIntent: persistIntent('override') })
      added += r.added
      mintedRuleIds.push(...r.addedIds)
      revived.push(...r.revivedRules)
      disabled.push(...r.disabledRules)
      details.push(...r.details)
      if (r.disabled > 0) {
        console.info(
          `[mail-organizer] ai_overridden disabled ${r.disabled} stale rule(s) pointing at the AI's old wrong target`,
        )
      }
    }
  } catch (e) {
    console.warn('[mail-organizer] generateAiOverrideRules failed (non-fatal)', e)
  }
  return {
    added,
    mintedRuleIds,
    details,
    ...(revived.length > 0 ? { revived } : {}),
    ...(disabled.length > 0 ? { disabled } : {}),
  }
}

/**
 * A persisted learner intent as output, when its write landed; null when it
 * never happened (re-run the learner). The write is atomic, so any trace of
 * it means all of it landed: a planned rule id in the library, or a rule no
 * longer in the state the write was about to change (enabled / provenance) —
 * not "still in the written state": a later change in the same batch (the
 * override learner superseding what the AI-confirmed one revived) must not
 * hide it. Rules gone since drop out; a later change is undo's fingerprint
 * guard's business.
 */
function landedIntent(intent: LearnerIntent, library: readonly Rule[]): LearnerIntent | null {
  const byId = new Map(library.map((r) => [r.id, r]))
  const movedOff = (rec: RevivedRuleUndo, wasEnabled: boolean) => {
    const r = byId.get(rec.ruleId)
    return (
      !!r &&
      !(
        r.enabled === wasEnabled &&
        r.autoDisabledReason === rec.autoDisabledReason &&
        r.autoDisabledAt === rec.autoDisabledAt
      )
    )
  }
  const disabled = intent.disabled ?? []
  const landed =
    intent.added.some((a) => byId.has(a.id)) ||
    intent.revived.some((rec) => movedOff(rec, false)) ||
    disabled.some((rec) => movedOff(rec, true))
  if (!landed) return null
  return {
    added: intent.added.filter((a) => byId.has(a.id)),
    revived: intent.revived.filter((rec) => byId.has(rec.ruleId)),
    disabled: disabled.filter((rec) => byId.has(rec.ruleId)),
  }
}

/**
 * What undo needs to put a rule back the way it was: its provenance BEFORE
 * this batch changed it and its fingerprint AFTER (2026-11-06). Whether it
 * was on or off before is known from the list the record goes in.
 */
function stateUndoRecord(before: Rule, after: Rule): RevivedRuleUndo {
  return {
    ruleId: after.id,
    ...(before.autoDisabledReason ? { autoDisabledReason: before.autoDisabledReason } : {}),
    ...(before.autoDisabledAt ? { autoDisabledAt: before.autoDisabledAt } : {}),
    fingerprint: ruleStateFingerprint(after),
  }
}

/** What undo needs to put each revived rule back to sleep (2026-11-06). */
function reviveRecords(reEnabled: readonly Rule[], revivedFrom: readonly Rule[]): RevivedRuleUndo[] {
  return reEnabled.map((after, i) => stateUndoRecord(revivedFrom[i]!, after))
}

async function flushSkipHistory(plan: PlanItem[], results: ExecuteItemResult[]): Promise<void> {
  // Remember which emails the user chose to keep in inbox — next classify
  // pass auto-excludes them so they don't show up again to be re-decided.
  //
  // CRITICAL (audit G1): gate on the plan item's action === 'skip', NOT
  // just status === 'skipped'. 'skipped' is overloaded: it also covers the
  // uncertain-network soft-skip and the 404 alreadyGone soft-skip. Those
  // must NOT land in skipHistory's 60-day suppression. Only a deliberate
  // action==='skip' is a real "user kept this in inbox" signal.
  // Never-classified placeholders (Claude call failed / omitted the email)
  // are NOT a keep decision unless the user explicitly chose 保留 on them —
  // recording them hid transient-error mail for 60 days (2026-11-05).
  const skippedIds = results
    .filter((r, i) => {
      const item = plan[i]
      if (r.status !== 'skipped' || item?.action !== 'skip') return false
      return !(item.unclassified === true && item.userTouched !== true)
    })
    .map((r) => r.emailId)
    .filter((id) => id)
  if (skippedIds.length > 0) await addToSkipHistory(skippedIds)
}

async function flushFolderActivity(plan: PlanItem[], results: ExecuteItemResult[]): Promise<void> {
  // Tally per-folder mail counts from this batch and feed the IdleScreen's
  // 「近日活動」quick-jump panel.
  const batchAt = new Date().toISOString()
  const folderTally = new Map<
    string,
    {
      folderId: string
      folderPath: string
      count: number
      latestMessage?: { subject: string; from: string; receivedAt: string }
    }
  >()
  for (let i = 0; i < results.length; i++) {
    const r = results[i]!
    if (!isLearnableFiling(r)) continue
    const planItem = plan[i]
    if (!planItem) continue
    // Prefer the move result's resolved destination (covers pending:
    // targets and in-batch created folders); fall back to the plan.
    const folderId =
      sanitizeRuleTargetFolderId(r.destinationFolderId) ||
      sanitizeRuleTargetFolderId(planItem.targetFolderId)
    const folderPath = r.destinationFolderPath ?? finalTargetPath(planItem)
    if (!folderId || !folderPath) continue
    const prev = folderTally.get(folderId)
    const latestMessage = {
      subject: (planItem.emailSubject ?? '').slice(0, 120),
      from: planItem.emailFrom ?? '',
      receivedAt: batchAt,
    }
    if (prev) {
      prev.count++
      prev.latestMessage = latestMessage
    } else {
      folderTally.set(folderId, { folderId, folderPath, count: 1, latestMessage })
    }
  }
  if (folderTally.size > 0) {
    await recordFolderActivityFromBatch([...folderTally.values()], batchAt)
  }
}

async function flushMetricsAndStats(state: ExecuteState, plan: PlanItem[], marker: FinalizeMarker): Promise<void> {
  try {
    // Once per batch (2026-11-06) — see applyCountsOnce.
    await applyCountsOnce(
      marker,
      'metrics',
      new Map([
        ['moved', state.summary.moved],
        ['deleted', state.summary.deleted],
        ['foldersCreated', state.summary.foldersCreated],
        ['errors', state.summary.errors],
      ]),
      async () => new Map(Object.entries(await getMetrics())),
      (inc) => bumpMetrics(Object.fromEntries(inc)),
    )
  } catch (e) {
    console.warn('[mail-organizer] metrics flush failed (non-fatal)', e)
  }

  // Stale-sweep clock (2026-11-05): the batch's day counts as an ACTIVE
  // day — the matcher ran over real inbox mail, so rules that didn't fire
  // have genuinely aged one day. Idle calendar days never reach this ledger.
  if (plan.length > 0) {
    await recordActiveDay(state.startedAt).catch((e) =>
      console.warn('[mail-organizer] recordActiveDay failed (non-fatal)', e),
    )
  }

  // Per-source stats ledger (優化 2026-07) — the instrument for telling
  // WHICH axis (rule / AI / thread) is erring over time.
  try {
    const rulesById = new Map((await listRules()).map((r) => [r.id, r]))
    const planned = { rule: 0, thread: 0, ai: 0, unresolved: 0 }
    const overrides = { rule: 0, ai: 0, thread: 0 }
    let gateResolved = 0
    for (let i = 0; i < plan.length; i++) {
      const item = plan[i]!
      const res = state.results[i]
      const terminal =
        res?.status === 'moved' || res?.status === 'deleted' || res?.status === 'folder_created'
      // planned: bucket by how the ITEM was sourced when the plan was built.
      // A user edit flips a rule/thread item's source to 'ai', so this
      // counts post-edit source.
      if (item.source === 'rule') planned.rule++
      else if (item.source === 'thread') planned.thread++
      else if (item.source === 'unresolved') planned.unresolved++
      else planned.ai++
      if (!terminal) continue
      // overrides.rule: a rule fired and the user re-targeted it to a
      // different folder — same definition the override tally uses
      // (2026-11-05; deletes no longer count as rule errors).
      if (
        item.source !== 'rule' &&
        item.originalRuleId &&
        isFiled(res) &&
        isRetargetAwayFromRule(rulesById.get(item.originalRuleId), item, res!)
      ) {
        overrides.rule++
      }
      // overrides.ai: user changed the AI's own verdict.
      if (item.source === 'ai' && wasUserOverride(item)) overrides.ai++
      // overrides.thread: user re-targeted away from the thread's home.
      if (
        item.source === 'thread' &&
        item.userTouched &&
        item.threadMatch &&
        finalTargetPath(item) !== item.threadMatch.previousFolderPath
      ) {
        overrides.thread++
      }
      // gateResolved: AI had an answer that the confidence gate demoted to
      // unresolved, and the user resolved it (aiOriginalAction survives).
      if (item.source === 'unresolved' && item.aiOriginalAction !== undefined) gateResolved++
    }
    // Once per batch (2026-11-06): the row's `at` is recorded before it is
    // appended; a replay that finds a row with that `at` doesn't add another.
    if (marker.statAt && (await getBatchStats()).some((s) => s.at === marker.statAt)) return
    const at = marker.statAt ?? new Date(state.finishedAt ?? Date.now()).toISOString()
    if (!marker.statAt) {
      marker.statAt = at
      await writeFinalizeMarker(marker)
    }
    await appendBatchStat({
      at,
      planned,
      moved: state.summary.moved,
      deleted: state.summary.deleted,
      overrides,
      gateResolved,
    })
  } catch (e) {
    console.warn('[mail-organizer] batch stats ledger failed (non-fatal)', e)
  }
}

/**
 * Recognises the "message no longer at this Id" 404 returned by Outlook
 * when:
 *   - The message was already moved (its Id changed; the OLD id is gone)
 *   - The message was deleted elsewhere
 *   - An earlier non-idempotent POST succeeded but its response was
 *     lost, and we (correctly) didn't retry — but the user reopened
 *     the popup and the result row still showed the failed attempt,
 *     prompting a retry which then hits 404.
 *
 * All three cases share an observable: from this batch's perspective,
 * the message is gone from its source location. There's no point
 * marking the row as a hard `error`:
 *   - It usually IS at its intended destination (case 3 above)
 *   - Even if it's not, retrying would 404 again
 *
 * Treat these as `skipped` with an explanatory message so the UI shows
 * a soft "noted, moved on" rather than a red error inviting another
 * retry — which is the user-visible loop that prompted this fix.
 *
 * (2026-11-05) A move to a PURGED destination can produce the same 404 — the
 * caller verifies the destination still exists before soft-skipping.
 */
export function isAlreadyMovedError(e: unknown): boolean {
  if (!(e instanceof OutlookError)) return false
  if (e.status !== 404) return false
  // The text "ErrorItemNotFound" is what Outlook embeds in the error
  // body and bubbles into OutlookError.message. Treat absence of the
  // marker as "some OTHER 404" (path typo, folder gone — different bug)
  // and let it surface normally.
  return /ErrorItemNotFound/i.test(e.message)
}

const ALREADY_MOVED_MESSAGE =
  '訊息已不在原位置(可能已成功移動、或被其他規則 / 帳號處理過)'

const DESTINATION_GONE_MESSAGE =
  '目標資料夾已不存在(可能已被刪除或清空)——這封信沒有移動、仍在收件匣。請重新選擇資料夾。'

// Refusals of non-filing destinations (2026-11-06). Each says WHY, instead of
// v1.0.8's catch-all 「(可能已被刪除)」, which also covered Junk / Sent Items.
const TRASHED_DESTINATION_MESSAGE =
  '目標資料夾已在 Outlook 被刪除(目前在「已刪除的郵件」裡,清空時會一起消失),不能當歸檔目標——這封信沒有移動。請把資料夾移回原處,或改選其他資料夾;若是規則指定的目標,也請修改該規則。'
const DRAFTS_OUTBOX_DESTINATION_MESSAGE =
  '目標資料夾是「草稿」或「寄件匣」(或在它們底下),不能當歸檔目標——這封信沒有移動。請改選其他資料夾;若是規則指定的目標,也請修改該規則。'
const LEGACY_EXCLUDED_DESTINATION_MESSAGE =
  '目標資料夾在 Outlook 的特殊資料夾(已刪除的郵件 / 草稿 / 寄件匣)底下,不能當歸檔目標——這封信沒有移動。請改選其他資料夾。'
const TRASH_PARENT_MESSAGE =
  '不能在「已刪除的郵件」裡建立新資料夾(清空時會一起刪除)——這封信沒有移動。請改選其他上層資料夾。'
// new_folder whose create 409'd against — or whose in-batch twin is — the
// trash itself (2026-11-06 r2): a top-level 「刪除的郵件」 is not a new folder.
const TRASH_AS_NEW_FOLDER_MESSAGE =
  '新資料夾和「已刪除的郵件」同名,不能當歸檔目標——這封信沒有移動。請改用其他名稱或改選資料夾。'

/** Where a non-filing folder sits; 'legacy' = a v1.0.8 cache id of unknown kind. */
type NonFilingKind = NonFilingRoot | 'legacy'

/**
 * A destination's standing (2026-11-06 r2): null = an ordinary target;
 * 'trash' = Deleted Items ITSELF (a move there runs as a delete); otherwise
 * the non-filing area it is in (refused). 'unverified' (2026-11-06 r3) = a
 * target outside the tree while Deleted Items can't be identified: it may
 * BE the trash — the move runs, but nothing is learned from it.
 */
type DestinationKind = 'trash' | 'unverified' | NonFilingKind | null

function nonFilingTargetMessage(kind: NonFilingKind): string {
  if (kind === 'deleteditems') return TRASHED_DESTINATION_MESSAGE
  if (kind === 'legacy') return LEGACY_EXCLUDED_DESTINATION_MESSAGE
  return DRAFTS_OUTBOX_DESTINATION_MESSAGE
}

type GuardInfo = {
  trashId?: string
  byId: Map<string, NonFilingKind>
  /**
   * The cached tree may be out of date (2026-11-06 r2): older than
   * FOLDER_TREE_VERIFY_AFTER_MS, invalidated, or undatable. Its folders are
   * then verified live before mail is filed into them. No cache at all =
   * nothing to date the tree by → trusted, as before.
   */
  treeOld: boolean
}

/** One live GET of a folder: its parent, or why there is none. */
type FolderHop = { kind: 'ok'; parent?: string } | { kind: 'gone' } | { kind: 'unknown' }

/** What a live parent walk found: a non-filing area, an ordinary folder (null), or can't tell. */
type LiveVerdict = NonFilingKind | null | 'unknown'

// Deeper than any real folder tree; a walk that gets here can't tell.
const LIVE_WALK_MAX_HOPS = 12

/**
 * Classifies destinations against the non-filing folders (2026-11-05,
 * redesigned 2026-11-06 r2): Deleted Items itself (a move there becomes a
 * delete), anything under it (deleting folder 案件/甲 in OWA re-parents it
 * there with the SAME id, so a rule, an old cache or a resumed plan could
 * otherwise keep filing into the trash), and Drafts / Outbox with their
 * subtrees (refused). Junk Email and Sent Items are ordinary targets.
 *
 * Sources, cheapest first: the roots stored with the folder cache
 * (FolderCache.wellKnown — zero requests when complete); otherwise the roots
 * resolved live once per batch (or, failing everything, a v1.0.8 cache's ids
 * with their old meaning). A target outside the tree gets a live parent walk.
 * So does a target IN the tree when the cached tree is old (2026-11-06 r2): a
 * folder deleted in OWA after the fetch is still in it, and mail filed there
 * sat in the trash until the trash was emptied. Walks go up to a known root,
 * are bounded, and are memoized per batch — the guard is per batch and the
 * parallel workers share it — so each folder, target or ancestor, is fetched
 * at most once (through the rate limiter, honouring the batch's signal).
 * Every lookup is defensive — failing to verify never blocks a move; only a
 * POSITIVE finding does (or a 404 for the target itself: DESTINATION_GONE).
 */
class DestinationGuard {
  private infoPromise: Promise<GuardInfo> | null = null
  private index: Map<string, MailFolderNode> | null = null
  private rootParents: Set<string> | null = null
  /** GET /MailFolders/{id}, once per folder per batch. */
  private readonly hops = new Map<string, Promise<FolderHop>>()
  /** Live verdict per in-tree target of an old tree, once per batch. */
  private readonly verdicts = new Map<string, Promise<DestinationKind>>()
  /** Folders this batch created (or adopted after a check) — trusted. */
  private readonly trusted = new Set<string>()
  private cacheInvalidated = false

  constructor(
    private readonly tree: MailFolderNode[],
    private readonly api: OutlookApi,
  ) {}

  private info(): Promise<GuardInfo> {
    if (!this.infoPromise) {
      this.infoPromise = (async () => {
        const byId = new Map<string, NonFilingKind>()
        let trashId: string | undefined
        let complete = false
        let cache: FolderCache | undefined
        try {
          cache = await getFolderCache()
        } catch {
          /* ignore */
        }
        if (cache?.wellKnown) {
          for (const r of cache.wellKnown.roots) {
            if (r.role === 'deleteditems') trashId = r.id
            else if (r.role === 'drafts' || r.role === 'outbox') byId.set(r.id, r.role)
          }
          complete = cache.wellKnown.complete
        } else if (cache?.nonFiling) {
          // Round-1 cache (refetched on first use; read until then).
          trashId = cache.nonFiling.trashFolderId
          for (const [id, kind] of Object.entries(cache.nonFiling.byId)) byId.set(id, kind)
          complete = true
        }
        if (!complete && typeof this.api.getWellKnownFolderId === 'function') {
          // Partial / missing → look up just the roots still unknown, each on
          // its own (2026-11-06 r3: a retry no longer hinges on all five
          // lookups succeeding together), once per batch.
          const knownRoles = new Set(byId.values())
          const missing = (['deleteditems', 'drafts', 'outbox'] as const).filter((role) =>
            role === 'deleteditems' ? !trashId : !knownRoles.has(role),
          )
          await Promise.all(
            missing.map(async (role) => {
              try {
                const id = await this.api.getWellKnownFolderId(role)
                if (!id) return
                if (role === 'deleteditems') trashId ??= id
                else if (!byId.has(id)) byId.set(id, role)
              } catch (e) {
                // A token outage stops the batch; anything else: unknown.
                if (e instanceof OutlookTokenUnavailableError) throw e
              }
            }),
          )
        }
        if (!trashId && byId.size === 0) {
          // Nothing resolved: a v1.0.8 cache's ids keep their old meaning
          // (root and subtree refused) rather than nothing at all.
          for (const id of cache?.excludedFolderIds ?? []) byId.set(id, 'legacy')
        }
        const updated = cache ? Date.parse(cache.updatedAt) : Number.NaN
        const treeOld =
          !!cache && (!!cache.invalidatedAt || !(Date.now() - updated < FOLDER_TREE_VERIFY_AFTER_MS))
        return { trashId, byId, treeOld }
      })()
    }
    return this.infoPromise
  }

  /** Lookup by id; rebuilds once on a miss (new_folder splices mid-batch). */
  lookup(id: string, rebuildOnMiss = true): MailFolderNode | undefined {
    if (!this.index || (rebuildOnMiss && !this.index.has(id))) {
      this.index = new Map(flattenFolderTree(this.tree).map((n) => [n.id, n]))
    }
    return this.index.get(id)
  }

  /** A folder this batch created, or adopted after checking it: trusted from now on. */
  trust(folderId: string): void {
    this.trusted.add(folderId)
  }

  /** What `folderId` is as a destination — see DestinationKind. */
  async destinationKind(folderId: string, signal: AbortSignal): Promise<DestinationKind> {
    const info = await this.info()
    const self = info.byId.get(folderId)
    if (self) return self
    if (folderId === info.trashId) return 'trash'
    const node = this.lookup(folderId)
    if (node) {
      const fromTree = this.fromTreeAncestors(node, info)
      if (fromTree || !info.treeOld || this.trusted.has(folderId)) return fromTree
      return this.verifiedInTree(folderId, signal, info)
    }
    // Not in the tree: walk the live parent chain — trusting the tree's
    // folders as ancestors only while the tree is fresh.
    const live = await this.liveKind(folderId, signal, info, !info.treeOld)
    if (live !== null && live !== 'unknown') return live
    // With Deleted Items unidentifiable even after a retry, a folder outside
    // the tree may be the trash itself (2026-11-06 r3): a rule on it used to
    // run as an ordinary filing — learned, remembered, 近日活動.
    if (!info.trashId) return 'unverified'
    return null
  }

  private ancestorKind(id: string, info: GuardInfo): NonFilingKind | null {
    return id === info.trashId ? 'deleteditems' : (info.byId.get(id) ?? null)
  }

  /** A STRICT ancestor in the tree decides: the trash, or a non-filing folder. */
  private fromTreeAncestors(node: MailFolderNode, info: GuardInfo): NonFilingKind | null {
    let parent = node.parentFolderId
    for (let hops = 0; parent && hops < 64; hops++) {
      const kind = this.ancestorKind(parent, info)
      if (kind) return kind
      parent = this.lookup(parent, false)?.parentFolderId
    }
    return null
  }

  /** Parents of the tree's top-level folders: the mailbox root (msgfolderroot). */
  private rootParentIds(): Set<string> {
    this.rootParents ??= new Set(this.tree.map((n) => n.parentFolderId).filter((p): p is string => !!p))
    return this.rootParents
  }

  private hop(id: string, signal: AbortSignal): Promise<FolderHop> {
    let p = this.hops.get(id)
    if (!p) {
      p = this.api.getFolder(id, signal).then(
        (f): FolderHop => ({ kind: 'ok', parent: f?.ParentFolderId || undefined }),
        (e): FolderHop => {
          // No token is not "can't tell" (2026-11-06 r3): the move would hit
          // it next, a full token wait later. The memoized rejection trips
          // every worker waiting on this folder at once.
          if (isAbortError(e) || e instanceof OutlookTokenUnavailableError) throw e
          return e instanceof OutlookError && e.status === 404 ? { kind: 'gone' } : { kind: 'unknown' }
        },
      )
      this.hops.set(id, p)
    }
    return p
  }

  /**
   * Walks `folderId`'s live parent chain up to a known root: the trash /
   * Drafts / Outbox (→ that area), the mailbox root (→ null, an ordinary
   * folder) or — with `trustTree` — a folder of the tree, whose tree
   * ancestors then decide. 404 for `folderId` itself = DESTINATION_GONE; any
   * other failure, or no root within the bound = 'unknown' (can't tell).
   */
  private async liveKind(
    folderId: string,
    signal: AbortSignal,
    info: GuardInfo,
    trustTree: boolean,
  ): Promise<LiveVerdict> {
    if (typeof this.api.getFolder !== 'function') return 'unknown'
    const roots = this.rootParentIds()
    let id = folderId
    for (let hop = 0; hop < LIVE_WALK_MAX_HOPS; hop++) {
      const h = await this.hop(id, signal)
      if (h.kind === 'gone') {
        if (hop === 0) throw new Error(DESTINATION_GONE_MESSAGE)
        return 'unknown'
      }
      if (h.kind === 'unknown' || !h.parent) return 'unknown'
      const kind = this.ancestorKind(h.parent, info)
      if (kind) return kind
      if (roots.has(h.parent)) return null
      if (trustTree) {
        const parentNode = this.lookup(h.parent, false)
        if (parentNode) return this.fromTreeAncestors(parentNode, info)
      }
      id = h.parent
    }
    return 'unknown'
  }

  /** An in-tree target of an old tree, verified live — once per batch. */
  private verifiedInTree(folderId: string, signal: AbortSignal, info: GuardInfo): Promise<DestinationKind> {
    let p = this.verdicts.get(folderId)
    if (!p) {
      p = (async (): Promise<DestinationKind> => {
        let live: LiveVerdict
        try {
          live = await this.liveKind(folderId, signal, info, false)
        } catch (e) {
          // Gone since the fetch: the cached tree is wrong as well.
          if (e instanceof Error && e.message === DESTINATION_GONE_MESSAGE) await this.invalidateCache()
          throw e
        }
        if (live === 'unknown') return null // can't tell → allowed, as before this check
        if (live) await this.invalidateCache()
        return live
      })()
      this.verdicts.set(folderId, p)
    }
    return p
  }

  /** The cached tree is out of date: the next preflight refetches it (and reconcile orphans its rules). */
  private async invalidateCache(): Promise<void> {
    if (this.cacheInvalidated) return
    this.cacheInvalidated = true
    await invalidateFolderCache().catch((e) =>
      console.warn('[mail-organizer] folder cache invalidation failed (non-fatal)', e),
    )
  }

  /**
   * A folder this new_folder item would file into WITHOUT creating it — its
   * in-batch twin, or the sibling a 409 / lost response recovered
   * (2026-11-06 r2). Must be an ordinary target: a top-level 「草稿」 /
   * 「寄件匣」 / 「刪除的郵件」 suggestion 409s against the well-known folder
   * itself, which used to be adopted (and spliced into the cached tree).
   * 'unverified' (2026-11-06 r3) = it can't be told apart from the trash:
   * the caller may file into it, but learns nothing and splices nothing.
   */
  async assertAdoptable(folderId: string, signal: AbortSignal): Promise<'unverified' | null> {
    const kind = await this.destinationKind(folderId, signal)
    if (kind === 'trash') throw new Error(TRASH_AS_NEW_FOLDER_MESSAGE)
    if (kind === 'unverified') return kind
    if (kind) throw new Error(nonFilingTargetMessage(kind))
    return null
  }

  /** new_folder's parent: an ordinary target (never the trash or anything under it). */
  async assertNewFolderParent(folderId: string, signal: AbortSignal): Promise<void> {
    const kind = await this.destinationKind(folderId, signal)
    if (kind === 'trash' || kind === 'deleteditems') throw new Error(TRASH_PARENT_MESSAGE)
    if (kind && kind !== 'unverified') throw new Error(nonFilingTargetMessage(kind))
  }

  /** Positive (404) confirmation that a folder no longer exists. */
  async isFolderGone(folderId: string, signal: AbortSignal): Promise<boolean> {
    if (typeof this.api.getFolder !== 'function') return false
    try {
      await this.api.getFolder(folderId, signal)
      return false
    } catch (e) {
      if (isAbortError(e)) throw e
      return e instanceof OutlookError && e.status === 404
    }
  }
}

/**
 * A move answered 404 ErrorItemNotFound. Usually the MESSAGE is gone
 * (already moved — soft skip). But a move into a purged destination folder
 * probably answers the same (2026-11-05), and soft-skipping that left the
 * mail in the inbox labelled "already moved". Verify the destination first.
 */
async function alreadyMovedOrDestinationGone(
  folderId: string,
  guard: DestinationGuard,
  signal: AbortSignal,
  extra: Partial<ItemOutcome> = {},
): Promise<ItemOutcome> {
  if (await guard.isFolderGone(folderId, signal)) throw new Error(DESTINATION_GONE_MESSAGE)
  return { status: 'skipped', message: ALREADY_MOVED_MESSAGE, alreadyGone: true, ...extra }
}

/** Exchange folder names are case-insensitive (and we trim like the UI). */
function sameFolderName(a: string | undefined, b: string): boolean {
  return (a ?? '').trim().toLowerCase() === b.trim().toLowerCase()
}

/**
 * `beforeSend` (2026-11-06) is awaited right before every non-idempotent
 * request — move, delete-as-move, folder create — so the in-flight marker is
 * durable before anything can change on the server.
 */
async function executeItem(
  item: PlanItem,
  tree: MailFolderNode[],
  api: OutlookApi,
  signal: AbortSignal,
  guard: DestinationGuard,
  beforeSend: () => Promise<void> = async () => {},
): Promise<ItemOutcome> {
  // Each network sub-call checks signal first. new_folder does
  // create + move (sometimes also list-siblings on 409) — without these
  // checks an in-flight cancel would let createFolder + moveMessage both
  // complete before the next item-boundary cancel test caught up.
  signal.throwIfAborted()

  if (item.action === 'skip') return { status: 'skipped' }

  // A move into Deleted Items (2026-11-05), so it shows up in OWA's
  // 已刪除的郵件 and 撤回 can bring it back.
  const moveToTrash = async (): Promise<ItemOutcome> => {
    try {
      await beforeSend()
      const res = await api.moveToDeletedItems(item.emailId, signal)
      return {
        status: 'deleted',
        newMessageId: res.Id,
        destinationFolderId: res.ParentFolderId || 'deleteditems',
      }
    } catch (e) {
      if (isAlreadyMovedError(e)) {
        // Delete target gone = deletion's already done by something
        // else (or by an earlier lost-response delete from us). Soft
        // success — don't surface as red error. (The destination is the
        // well-known Deleted Items, which can't be purged away.)
        return { status: 'skipped', message: ALREADY_MOVED_MESSAGE, alreadyGone: true }
      }
      throw e
    }
  }

  if (item.action === 'delete') return moveToTrash()

  if (item.action === 'move') {
    const targetId = resolveTargetFolderId(item, tree)
    if (!targetId) throw new Error('目標資料夾未指定或無法在資料夾樹中找到')
    const kind = await guard.destinationKind(targetId, signal)
    // Deleted Items itself (a rule that auto-trashes, 2026-11-06 r2) runs as
    // a delete: status 'deleted' — no learning, thread memory or folder
    // activity (trash moves used to be learned as filings) — and 撤回 returns
    // the mail to the inbox. A rule-routed one still counts its rule hit.
    if (kind === 'trash') return moveToTrash()
    if (kind && kind !== 'unverified') throw new Error(nonFilingTargetMessage(kind))
    try {
      await beforeSend()
      const res = await api.moveMessage(item.emailId, targetId, signal)
      return {
        status: 'moved',
        newMessageId: res.Id,
        destinationFolderId: targetId,
        destinationFolderPath: item.targetFolderPath ?? guard.lookup(targetId)?.path,
        // May have been the trash itself (2026-11-06 r3): moved, not learned.
        ...(kind === 'unverified' ? { unverifiedTarget: true } : {}),
      }
    } catch (e) {
      if (isAlreadyMovedError(e)) return alreadyMovedOrDestinationGone(targetId, guard, signal)
      throw e
    }
  }

  if (item.action === 'new_folder') {
    const name = item.suggestedFolderName?.trim()
    if (!name) throw new Error('未指定新資料夾名稱')
    let parentId: string | undefined
    let siblingsInTree: MailFolderNode[] = tree
    if (item.suggestedParentPath) {
      const parent = findFolderByPath(tree, item.suggestedParentPath)
      if (parent) {
        parentId = parent.id
        siblingsInTree = parent.children
      } else throw new Error(`找不到父資料夾「${item.suggestedParentPath}」`)
      await guard.assertNewFolderParent(parentId, signal)
    }

    // In-batch dedup: prior new_folder items have already created what's
    // needed (reorder above guarantees those run first). Case-insensitive
    // (2026-11-05) — Exchange treats "ACME Corp" and "Acme Corp" as the
    // same name, so creating it would just 409.
    const existing = siblingsInTree.find((n) => sameFolderName(n.displayName, name))
    if (existing) {
      // A tree that listed a well-known root (degraded fetch, legacy cache)
      // must not turn 「草稿」 into a filing target (2026-11-06 r2).
      const existingKind = await guard.assertAdoptable(existing.id, signal)
      signal.throwIfAborted()
      try {
        await beforeSend()
        const res = await api.moveMessage(item.emailId, existing.id, signal)
        return {
          status: 'moved',
          newMessageId: res.Id,
          destinationFolderId: existing.id,
          destinationFolderPath: existing.path,
          ...(existingKind === 'unverified' ? { unverifiedTarget: true } : {}),
        }
      } catch (e) {
        if (isAlreadyMovedError(e)) return alreadyMovedOrDestinationGone(existing.id, guard, signal)
        throw e
      }
    }

    let created: { Id: string; DisplayName: string; ParentFolderId?: string }
    let recoveredFromExisting = false
    // The recovered folder can't be told apart from the trash (2026-11-06 r3).
    let unverified = false
    try {
      signal.throwIfAborted()
      await beforeSend()
      created = parentId
        ? await api.createChildFolder(parentId, name, signal)
        : await api.createTopFolder(name, signal)
    } catch (e) {
      if (isAbortError(e)) throw e
      const conflict = e instanceof OutlookError && e.status === 409
      // uncertain = our POST may have created the folder before its
      // response was lost (network / 5xx / cancel mid-POST) (2026-11-05).
      const lost = e instanceof OutlookError && e.uncertain
      if (!conflict && !lost) throw e
      console.warn(
        `[mail-organizer] createFolder ${conflict ? '409' : 'response lost'} for ${name}; re-listing siblings`,
      )
      // Cancel mid-create: the mail was never moved → plain 'cancelled'.
      signal.throwIfAborted()
      let match: MailFolder | undefined
      try {
        // Paged listing (2026-11-05) — the old single page missed siblings 201+.
        const siblings = parentId
          ? await api.listChildFolders(parentId, signal)
          : await api.listTopFolders(signal)
        match = siblings.find((f) => sameFolderName(f.DisplayName, name))
      } catch (listErr) {
        if (isAbortError(listErr)) throw listErr
        throw new Error(
          conflict
            ? `Outlook 回 409(同名資料夾已存在)但無法列出同層資料夾:${listErr instanceof Error ? listErr.message : String(listErr)}`
            : `建立資料夾「${name}」時網路中斷,也無法確認是否已建立;這封信沒有移動、仍在收件匣,可重試。`,
        )
      }
      if (!match) {
        // The create definitely didn't take (or it's a different 409) —
        // and the mail was never moved: the move is only issued after a
        // folder id is in hand. So never say "may have moved" here.
        throw new Error(
          conflict
            ? `Outlook 回 409 但同層找不到「${name}」資料夾`
            : `建立資料夾「${name}」時網路中斷,資料夾未建立;這封信沒有移動、仍在收件匣,可重試。`,
        )
      }
      // Never adopt — or splice into the cached tree — a non-filing match
      // (2026-11-06 r2): a top-level 「草稿」 409s against Drafts itself.
      unverified = (await guard.assertAdoptable(match.Id, signal)) === 'unverified'
      created = match
      // Found after a lost response → most likely OUR create; report it so.
      recoveredFromExisting = conflict
    }

    // Use the folder's real DisplayName for the path (a case-insensitive 409
    // recovery may differ in casing from the suggested name).
    const createdPath = joinFolderPath(item.suggestedParentPath, created.DisplayName || name)

    if (!unverified) {
      // Created under a checked parent (or checked when adopted): later
      // moves into it this batch need no live check (2026-11-06 r2).
      guard.trust(created.Id)
      // Splice into in-memory tree + persist (idempotent). Never an
      // unverifiable match: it may be the trash itself (2026-11-06 r3).
      spliceFolderIntoTree(tree, created, parentId)
      addFolderToCache(created, parentId).catch((e) =>
        console.warn('[mail-organizer] addFolderToCache failed (non-fatal)', e),
      )
    }

    signal.throwIfAborted()
    try {
      await beforeSend()
      const res = await api.moveMessage(item.emailId, created.Id, signal)
      return {
        status: recoveredFromExisting ? 'moved' : 'folder_created',
        newMessageId: res.Id,
        destinationFolderId: created.Id,
        destinationFolderPath: createdPath,
        ...(unverified ? { unverifiedTarget: true } : {}),
      }
    } catch (e) {
      if (isAlreadyMovedError(e)) {
        // Folder was already created (or recovered from 409) — that
        // bookkeeping stands. We just couldn't deposit this particular
        // message because its source Id is no longer valid. Tag the
        // outcome as skipped so the row doesn't appear as a hard error.
        // Keep the folder visible to downstream code (sibling items in
        // the same batch can use it via the spliced-into-tree path).
        return alreadyMovedOrDestinationGone(created.Id, guard, signal, {
          destinationFolderId: created.Id,
          destinationFolderPath: createdPath,
        })
      }
      throw e
    }
  }

  throw new Error(`未知 action: ${item.action}`)
}

function spliceFolderIntoTree(
  tree: MailFolderNode[],
  folder: { Id: string; DisplayName: string; ParentFolderId?: string },
  parentId: string | undefined,
): void {
  let parentChildren: MailFolderNode[] = tree
  let parentPath = ''
  if (parentId) {
    const findParent = (nodes: MailFolderNode[]): MailFolderNode | undefined => {
      for (const n of nodes) {
        if (n.id === parentId) return n
        const found = findParent(n.children)
        if (found) return found
      }
      return undefined
    }
    const parent = findParent(tree)
    if (!parent) return
    parentChildren = parent.children
    parentPath = parent.path
  }
  if (parentChildren.some((n) => n.id === folder.Id)) return
  parentChildren.push({
    id: folder.Id,
    displayName: folder.DisplayName,
    parentFolderId: folder.ParentFolderId,
    path: joinFolderPath(parentPath || undefined, folder.DisplayName),
    children: [],
  })
}

function isValidFolderId(id: string | undefined | null): boolean {
  if (!id) return false
  if (id.startsWith('PLACEHOLDER')) return false
  // `pending:<emailId>` is the popup sentinel for "user picked a folder that
  // will be created by another item in this same plan". Force a path lookup
  // at execute time instead — by then the real folder has been spliced into
  // the tree by the new_folder handler.
  if (id.startsWith('pending:')) return false
  // Real Outlook folder IDs are long base64-like strings. Defensively reject
  // anything implausibly short.
  if (id.length < 20) return false
  return true
}

function resolveTargetFolderId(item: PlanItem, tree: MailFolderNode[]): string | undefined {
  if (isValidFolderId(item.targetFolderId)) return item.targetFolderId
  if (item.targetFolderPath) {
    return findFolderByPath(tree, item.targetFolderPath)?.id
  }
  return undefined
}

/**
 * Strip non-real folder IDs before saving into a rule. `pending:<emailId>` is
 * an in-batch sentinel for "user picked a folder that will exist after this
 * batch's new_folder runs" — it must NOT leak into persisted rules, where
 * it'd just be opaque garbage. Empty string is fine: path-based lookup
 * fallback handles match time.
 */
function sanitizeRuleTargetFolderId(id: string | undefined | null): string {
  if (!id) return ''
  if (id.startsWith('PLACEHOLDER')) return ''
  if (id.startsWith('pending:')) return ''
  if (id.length < 20) return ''
  return id
}

// ---- Rule feedback ---------------------------------------------------------

/**
 * Compute what the AI originally pointed this item at — collapses 'move' +
 * 'new_folder' to a single target path so we can compare against the user's
 * final choice with one equality check.
 */
export function aiOriginalTargetPath(item: PlanItem): string | undefined {
  if (item.aiOriginalAction === 'move') return item.aiOriginalTargetFolderPath
  if (
    item.aiOriginalAction === 'new_folder' &&
    item.aiOriginalSuggestedParentPath &&
    item.aiOriginalSuggestedFolderName
  ) {
    return joinFolderPath(item.aiOriginalSuggestedParentPath, item.aiOriginalSuggestedFolderName)
  }
  return undefined
}

export function finalTargetPath(item: PlanItem): string | undefined {
  if (item.action === 'move') return item.targetFolderPath
  if (
    item.action === 'new_folder' &&
    item.suggestedParentPath &&
    item.suggestedFolderName
  ) {
    return joinFolderPath(item.suggestedParentPath, item.suggestedFolderName.trim())
  }
  return undefined
}

/**
 * Did the user override what the AI suggested? Returns true only when the AI
 * had a verdict (aiOriginalAction defined) AND the final action OR target
 * differs from it. Items the AI gave up on (unresolved → user fills in) are
 * NOT overrides — those flow through ai_confirmed instead.
 */
export function wasUserOverride(item: PlanItem): boolean {
  if (item.aiOriginalAction === undefined) return false
  if (item.action !== item.aiOriginalAction) return true
  return finalTargetPath(item) !== aiOriginalTargetPath(item)
}

/**
 * Pick the best signal to attach a learned rule to.
 *
 * Priority (most specific → least specific). Future emails get routed
 * correctly even when the same domain serves multiple cases — the
 * dominant pattern in any organisation that sends mail about distinct
 * projects / clients from a shared address.
 *
 *   1. domain + Taiwan court case number → compound  (Taiwan legal)
 *      e.g. `telco.example + 112訴204` — uniquely identifies a case-client pair.
 *   2. Taiwan court case number alone (no usable domain) → subject_keyword
 *      Court case numbers are globally unique per court; safe.
 *   3. Latin case code alone → case_code
 *      Structurally unique within a firm's internal coding system.
 *   4. usable domain (non-generic-provider, non-internal, non-own) → domain,
 *      upgraded to compound (domain + 整段主旨) on a conflict
 *   5. generic-provider sender → sender (exact email address), upgraded
 *      to compound (sender + 整段主旨) on a conflict
 *      gmail.com / yahoo.com etc. as a domain is useless; a specific
 *      address like andy@gmail.com IS a precise routing signal.
 *   6. internal / own-domain sender → compound (sender + 整段主旨)
 *
 * The `featureKind` discriminator on compound / subject_keyword variants
 * lets `generateAiConfirmedRules` apply differentiated confidence:
 * tight identifiers (court case) get higher confidence than generic
 * subject features.
 *
 * Returns null when no usable signal exists (internal domain only, no
 * usable subject, no @, etc.).
 */
/**
 * Design principle (2026-05-27 v2): "先廣後窄" — prefer the broadest
 * signal that works (a single plain-domain or plain-sender rule covers
 * many emails), and only upgrade to compound (domain + 整段主旨) when a
 * conflict surfaces (same domain genuinely routing to different
 * folders). Prevents rule library bloat from per-subject compound
 * rules accumulating one-per-email forever.
 *
 * Structural identifiers (court case numbers, Latin case codes) are
 * exempt — they're naturally unique-per-case, so building a compound
 * rule for them doesn't cause growth and gives stronger cross-domain
 * routing (a court case number is recognized regardless of which
 * lawyer's domain forwards the message).
 *
 * `demoteOnly` flag — set when chooseLearningSignal detected a domain /
 * sender conflict with an existing rule but couldn't extract any subject
 * to upgrade with. The override pipeline still uses this signal to
 * DEMOTE the stale rule, but SKIPS creating a new rule. Fires for
 * subjects too short to learn (< MIN_NORMALIZED_SUBJECT_LEN, 2026-11-05).
 *
 * `featureKind` on subject_keyword / compound variants distinguishes:
 *   - 'court_case': signal is a Taiwan court case number — structural,
 *     safe at high confidence.
 *   - 'full_subject': signal is the entire normalized subject — used
 *     by the conflict-upgrade path and the internal-domain fallback.
 */
// `fromBody` (batch-3): the structural case identifier came from the email
// BODY, not the subject. Body cases are noisier (quote other parties' cases),
// so downstream drops their confidence and tightens the source gate.
type LearningSignal =
  | { type: 'compound'; signal: string; descr: string; featureKind: 'court_case' | 'full_subject'; demoteOnly?: false; fromBody?: boolean }
  | { type: 'case_code'; signal: string; descr: string; demoteOnly?: false; fromBody?: boolean }
  | { type: 'subject_keyword'; signal: string; descr: string; featureKind: 'court_case' | 'full_subject'; demoteOnly?: false; fromBody?: boolean }
  | { type: 'domain'; signal: string; descr: string; demoteOnly?: boolean }
  | { type: 'sender'; signal: string; descr: string; demoteOnly?: boolean }

/** `${type}::${normalizeSignal}` — key of LearningLookup.bySignal. */
function signalKey(type: Rule['type'], signal: string): string {
  return `${type}::${normalizeSignal(type, signal)}`
}

/**
 * Every stored rule bucketed by (type, normalized signal), built once per
 * learning pass (2026-11-05). Replaces the per-item `existing.some(…
 * normalizeSignal …)` scans — several per item, each JSON-parsing every
 * compound rule (≈30k normalizations per 50-item batch on a 300-rule
 * library).
 *
 * `narrowedByAnchor` (2026-11-06): live AUTO full-subject compounds by the
 * domain / sender anchor they carry (`domain::x` / `sender::y`) — the trace
 * a 先廣後窄 upgrade leaves. While one routes elsewhere, a new filing on the
 * anchor learns the narrow form too (chooseLearningSignal), even after the
 * superseded broad rule itself has been purged. Hand-written compounds don't
 * count: a broad rule plus the user's own narrow exception is a legitimate
 * library.
 */
type LearningLookup = {
  bySignal: ReadonlyMap<string, readonly Rule[]>
  narrowedByAnchor: ReadonlyMap<string, readonly Rule[]>
}

function buildLearningLookup(rules: readonly Rule[]): LearningLookup {
  const bySignal = new Map<string, Rule[]>()
  const narrowedByAnchor = new Map<string, Rule[]>()
  const push = (map: Map<string, Rule[]>, key: string, r: Rule) => {
    const bucket = map.get(key)
    if (bucket) bucket.push(r)
    else map.set(key, [r])
  }
  for (const r of rules) {
    push(bySignal, signalKey(r.type, r.signal), r)
    // Same predicate as the central revive guard (rules.ts reviveWouldConflict).
    if (r.enabled && !r.orphaned && isAnchorNarrowingCompound(r)) {
      for (const anchor of ruleAnchorKeys(r)) push(narrowedByAnchor, anchor, r)
    }
  }
  return { bySignal, narrowedByAnchor }
}

/**
 * The anchor a learner would route this item's sender on — `domain::<d>` for
 * a usable non-generic domain (P4), `sender::<addr>` for a generic-provider
 * address (P5) — or null (internal / own / no address). Mirrors
 * chooseLearningSignal's usableDomain logic.
 */
function learningAnchorOf(item: PlanItem, internalDomainSet: ReadonlySet<string>): string | null {
  const from = (item.emailFrom ?? '').toLowerCase().trim()
  const at = from.lastIndexOf('@')
  const domain = at >= 0 ? from.slice(at + 1).trim() : ''
  if (!domain || internalDomainSet.has(domain) || internalDomainSet.has(from)) return null
  return isGenericProviderDomain(domain) ? signalKey('sender', from) : signalKey('domain', domain)
}

/**
 * The rule that routed a rule-sourced row — the stored copy, else the plan's
 * snapshot of it — or null when the row wasn't (or no longer is) rule-routed.
 */
function routingRuleOf(
  item: PlanItem,
  rulesById: ReadonlyMap<string, Rule>,
): Pick<Rule, 'type' | 'signal' | 'source'> | null {
  if (item.source !== 'rule') return null
  return (item.ruleId ? rulesById.get(item.ruleId) : undefined) ?? item.matchedRule ?? null
}

/**
 * A hand-written rule that itself routes by the subject or the case — a
 * subject_keyword, a case_code, or a compound with a subject condition
 * (R6, narrowed 2026-11-06). The mail it routes teaches no case rule: the
 * user already decides that mail by its subject, and a case rule learned
 * from it would turn the user's exception inside the case into the whole
 * case's route. A hand-written plain domain / sender rule (or a compound
 * of those) says nothing about the case, so its mail still teaches it.
 */
function isSubjectBearingUserRule(r: Pick<Rule, 'type' | 'signal' | 'source'>): boolean {
  if (r.source !== 'user_manual') return false
  if (r.type === 'subject_keyword' || r.type === 'case_code') return true
  if (r.type !== 'compound') return false
  return (decodeCompound(r.signal)?.conditions ?? []).some((c) => c.type === 'subject_keyword')
}

/**
 * Does `r` route ahead of every AUTO domain rule — case_code, compound,
 * court-case subject, or the user_manual tier (matchEmailWithIndex 1–4)?
 */
function outranksAutoDomain(r: Pick<Rule, 'type' | 'signal' | 'source'>): boolean {
  return (
    r.source === 'user_manual' ||
    r.type === 'case_code' ||
    r.type === 'compound' ||
    (r.type === 'subject_keyword' && courtCaseSignal(r.signal) !== null)
  )
}

/**
 * Folders THIS batch filed each anchor's mail into (2026-11-06) — every
 * successful filing, whatever routed it. Two folders for one anchor in one
 * batch is the same evidence a stored conflict is: a broad rule for either
 * side would misroute the other (the same-batch variant of "an A filing
 * revives broad→A while a B filing mints domain→B). Mail whose SUBJECT
 * carries a case number / case code doesn't count: the structural tiers
 * route it ahead of any domain rule (as court-case compounds don't count in
 * narrowedByAnchor). Nor does mail a rule that outranks auto domain rules
 * routed (R3) — a hand-written exception, a compound: it keeps routing that
 * mail whatever broad rule is learned next to it. Counting it, one such
 * filing per client per batch (near-certain at 500 mails) kept the broad
 * rule from ever being learned. An auto full-subject compound's routing
 * still counts — through narrowedByAnchor, from the library.
 */
function batchAnchorTargets(
  plan: readonly PlanItem[],
  results: readonly ExecuteItemResult[],
  internalDomainSet: ReadonlySet<string>,
  rulesById: ReadonlyMap<string, Rule>,
): Map<string, Set<string>> {
  const out = new Map<string, Set<string>>()
  for (let i = 0; i < plan.length; i++) {
    if (!isFiled(results[i])) continue
    const item = plan[i]!
    if (item.action !== 'move' && item.action !== 'new_folder') continue
    if (caseSignalsForLearning(item).source === 'subject') continue
    const via = routingRuleOf(item, rulesById)
    if (via && outranksAutoDomain(via)) continue
    const target = finalTargetPath(item)
    const anchor = learningAnchorOf(item, internalDomainSet)
    if (!target || !anchor) continue
    const set = out.get(anchor) ?? new Set<string>()
    set.add(target)
    out.set(anchor, set)
  }
  return out
}

/**
 * Domains the learners must never mint a domain rule for: the configured
 * internal domains plus the signed-in mailbox's own domain (2026-11-05).
 * With internalDomains empty (the default), the first AI filing of a
 * colleague's mail learned `domain: <own domain>` — and since the owner is
 * on To/Cc of every inbound mail, that rule claimed nearly the whole inbox.
 *
 * A generic-provider mailbox (me@outlook.com, 2026-11-06) shares its domain
 * with every stranger on it: excluding the DOMAIN stopped friend@outlook.com
 * from ever learning its sender rule. There only the owner's own ADDRESS is
 * excluded (entries containing '@' are exact addresses — chooseLearningSignal
 * checks the From address against them too).
 */
function learningExclusionDomains(
  internalDomains: readonly string[],
  identity: MailboxIdentity | null,
): Set<string> {
  const out = new Set(internalDomains.map((d) => d.trim().toLowerCase().replace(/^@/, '')))
  if (identity?.domain) {
    if (isGenericProviderDomain(identity.domain)) out.add(identity.address)
    else out.add(identity.domain)
  }
  out.delete('')
  return out
}

/**
 * Existing rules (bucketed by signal) passed alongside the email being
 * learned-from. Used by the domain (P4) and sender (P5) branches to detect
 * whether the email would create a conflict with a stored
 * same-signal-different-target rule — that's the moment when "先廣後窄"
 * upgrades to compound.
 *
 * When the user just routed an email and the existing broad rule was
 * about to fire elsewhere, that's the signal "the broad rule isn't
 * enough anymore". We use THIS email's subject to build a precise
 * compound rule, demoting the broad rule via the override pipeline.
 *
 * `correction` (2026-11-05): the item is an explicit correction of the rule
 * that routed it (ruleCorrectionOf). That rule pointed this very email
 * elsewhere, so the conflict exists by definition — whatever its type —
 * and P4 / P5 upgrade straight to the narrow form.
 *
 * The conflict check also covers (2026-11-06) the anchor's HISTORY — a plain
 * rule the engine retired for routing it wrong ('superseded' / 'override',
 * isConflictRetired), whatever folder it pointed at — live auto narrow rules
 * on the anchor that route elsewhere (lookup.narrowedByAnchor), and this
 * batch's own filings of the anchor elsewhere (`batchAnchorTargets`). Without
 * them the broad rule came straight back: the next A filing revived the
 * superseded domain→A (or a B filing minted a fresh domain→B), so 先廣後窄
 * silencing lasted about one batch.
 */
function chooseLearningSignal(
  item: PlanItem,
  internalDomainSet: ReadonlySet<string>,
  lookup: LearningLookup,
  opts: {
    correction?: boolean
    batchAnchorTargets?: ReadonlyMap<string, ReadonlySet<string>>
  } = {},
): LearningSignal | null {
  const subject = item.emailSubject ?? ''
  const target = finalTargetPath(item)
  // Case identifiers via the shared gate (batch-3): subject-first, body only
  // as a single-distinct fallback. `fromBody` flows into the confidence and
  // source gate downstream. An ambiguous body (>1 distinct) yields no case
  // signal at all → we fall through to the domain path (寧缺勿濫).
  const caseGate = caseSignalsForLearning(item)
  const fromBody = caseGate.source === 'body'
  const cc = fromBody && caseGate.bodyAmbiguous ? undefined : caseGate.courtCases[0]
  const codeAlpha = fromBody && caseGate.bodyAmbiguous ? undefined : caseGate.caseCodes[0]
  // Normalized full subject (reply/forward prefixes stripped, lower-
  // cased, whitespace collapsed). Returns '' for subjects too short to
  // learn (< MIN_NORMALIZED_SUBJECT_LEN since 2026-11-05).
  const subjectSignal = extractSubjectSignal(subject)

  const fromAddr = item.emailFrom ?? ''
  const at = fromAddr.lastIndexOf('@')
  const domain =
    at >= 0 ? fromAddr.slice(at + 1).toLowerCase().trim() : ''
  // "Usable" excludes internal domains (configurable via
  // settings.internalDomains) and the mailbox's own domain — colleagues
  // forwarding multi-case mail should not produce per-folder domain rules —
  // or, for a generic-provider mailbox, the owner's own address.
  // An empty set means solo mode, every external domain is usable.
  const usableDomain =
    domain && !internalDomainSet.has(domain) && !internalDomainSet.has(fromAddr.toLowerCase().trim())
      ? domain
      : ''
  // The anchor already routes (or routed, or this batch routes) elsewhere —
  // see the function doc.
  const conflictsWith = (type: 'domain' | 'sender', value: string): boolean => {
    if (opts.correction === true) return true
    const key = signalKey(type, value)
    for (const r of lookup.bySignal.get(key) ?? []) {
      if (r.enabled && !r.orphaned && r.targetFolderPath !== target) return true
      if (isConflictRetired(r)) return true
    }
    if ((lookup.narrowedByAnchor.get(key) ?? []).some((r) => r.targetFolderPath !== target)) {
      return true
    }
    for (const t of opts.batchAnchorTargets?.get(key) ?? []) if (t !== target) return true
    return false
  }

  // ---- Priority 1: compound with court case (structural ID fast-path)
  // Court case numbers are unique per case → no bloat risk. Build
  // compound + domain immediately for strongest routing. Confidence 0.9.
  // ONLY for SUBJECT cases (!fromBody): a compound's subject_keyword
  // condition is matched against the SUBJECT (matchCompound), so a
  // body-derived case number in a compound would never fire → the stale
  // sweep would hard-delete it. Body court cases fall to P2, which mints a
  // pure-case subject_keyword rule that C5's internal-body pass CAN match —
  // this is the closed-loop symmetry guarantee.
  if (usableDomain && cc && !fromBody) {
    const signal = encodeCompound([
      { type: 'domain', value: usableDomain },
      { type: 'subject_keyword', value: cc },
    ])
    return {
      type: 'compound',
      signal,
      descr: `compound: @${usableDomain} + 主旨含 ${cc}`,
      featureKind: 'court_case',
    }
  }
  // ---- Priority 2: court case alone (structural ID fast-path)
  // Court case numbers are globally unique; safe to route on the identifier
  // alone even without a domain anchor. Confidence 0.85 (subject) / 0.75 (body).
  if (cc) {
    return {
      type: 'subject_keyword',
      signal: cc,
      descr: `${fromBody ? '內文' : '主旨'}含 ${cc}`,
      featureKind: 'court_case',
      fromBody,
    }
  }
  // ---- Priority 3: Latin case code (structural ID fast-path)
  // Firm-internal case codes (25A0067A etc.) — also unique per case.
  // Confidence 0.9 (subject) / 0.8 (body).
  if (codeAlpha) {
    return {
      type: 'case_code',
      signal: codeAlpha,
      descr: `${fromBody ? '內文' : ''}案件代號 ${codeAlpha}`,
      fromBody,
    }
  }
  // ---- Priority 4: usable domain (non-generic-provider)
  //
  // "先廣後窄" — start with a plain-domain rule. It covers all mail
  // from this client in one shot. Only when a CONFLICT exists (same
  // domain already routes to a DIFFERENT folder) do we upgrade to
  // compound (domain + 整段主旨), encoding what made the user pick THIS
  // folder over the existing one.
  if (usableDomain && !isGenericProviderDomain(usableDomain)) {
    // Compared via normalizeSignal (the lookup key) so a stored rule
    // signal of `@company-b.example` or ` company-b.example ` (manual /
    // import paths) still matches `company-b.example` from the address.
    if (conflictsWith('domain', usableDomain)) {
      // Upgrade path: this email's normalized subject becomes the
      // discriminator. The override pipeline will also demote the
      // stale plain-domain rule, so future mail of this domain whose
      // subject doesn't match either compound rule falls back to AI.
      if (subjectSignal) {
        const signal = encodeCompound([
          { type: 'domain', value: usableDomain },
          { type: 'subject_keyword', value: subjectSignal },
        ])
        return {
          type: 'compound',
          signal,
          descr: `compound: @${usableDomain} + 主旨「${subjectSignal}」(網域已衝突、升級為主旨規則)`,
          featureKind: 'full_subject',
        }
      }
      // No subject to upgrade with — just demote the stale rule.
      return {
        type: 'domain',
        signal: usableDomain,
        descr: `@${usableDomain} (僅停用既有衝突規則、主旨過短無法升級)`,
        demoteOnly: true,
      }
    }
    // Common case: no conflict → broad plain-domain rule. Will catch
    // every subsequent email from this domain until the user routes
    // one elsewhere (then P4 conflict-upgrade fires).
    return { type: 'domain', signal: usableDomain, descr: `@${usableDomain}` }
  }
  // ---- Priority 5: generic-provider sender (gmail / yahoo / outlook…)
  //
  // Generic-provider DOMAIN is useless (millions of unrelated humans
  // share it), so skip the plain-domain rule and anchor on the exact
  // sender address from the start. Same "先廣後窄" pattern applied to
  // sender: plain-sender broadly, upgrade to compound on conflict.
  if (fromAddr && usableDomain && isGenericProviderDomain(usableDomain)) {
    const normalized = fromAddr.toLowerCase().trim()
    if (normalized) {
      if (conflictsWith('sender', normalized)) {
        if (subjectSignal) {
          const signal = encodeCompound([
            { type: 'sender', value: normalized },
            { type: 'subject_keyword', value: subjectSignal },
          ])
          return {
            type: 'compound',
            signal,
            descr: `compound: ${normalized} + 主旨「${subjectSignal}」(寄件人已衝突、升級為主旨規則)`,
            featureKind: 'full_subject',
          }
        }
        return {
          type: 'sender',
          signal: normalized,
          descr: `寄件人 ${normalized} (僅停用既有衝突規則、主旨過短無法升級)`,
          demoteOnly: true,
        }
      }
      return {
        type: 'sender',
        signal: normalized,
        descr: `寄件人 ${normalized}`,
      }
    }
  }
  // ---- Priority 6: internal-domain fallback (subject anchored on sender)
  //
  // Reached when there's no usable domain (internal forwarding from
  // colleagues — including the mailbox's own domain — anonymous From,
  // etc.). Domain rule isn't appropriate here — internal colleagues mail
  // about many cases.
  //
  // Subject alone is also dangerous: a bare `subject_keyword: 「請審閱
  // 三井租賃合約」` rule would substring-match an external client's reply
  // quoting the same subject. To narrow scope, anchor on the sender
  // address when available — `compound (sender + full_subject)` matches
  // only when BOTH conditions hold. Confidence 0.85 (same as the P4
  // conflict-upgrade form).
  //
  // Only fall back to plain subject_keyword when there's no sender to
  // anchor on (e.g., system-generated mail with empty From). Even then
  // it's labeled clearly so users can spot it in the rules UI.
  if (subjectSignal) {
    const normalizedFrom = fromAddr ? fromAddr.toLowerCase().trim() : ''
    if (normalizedFrom) {
      const signal = encodeCompound([
        { type: 'sender', value: normalizedFrom },
        { type: 'subject_keyword', value: subjectSignal },
      ])
      return {
        type: 'compound',
        signal,
        descr: `compound: ${normalizedFrom} + 主旨「${subjectSignal}」(內部信件)`,
        featureKind: 'full_subject',
      }
    }
    return {
      type: 'subject_keyword',
      signal: subjectSignal,
      descr: `主旨「${subjectSignal}」(無寄件人,僅依主旨)`,
      featureKind: 'full_subject',
    }
  }
  return null
}

/**
 * The rule a plan item EXPLICITLY corrected, or null (2026-11-05).
 *
 * A rule-routed row the user re-targeted reaches execute as source 'ai' with
 * originalRuleId + userTouched but no aiOriginalAction (PlanRow flips the
 * source when the row is edited and the rule never produced an AI verdict).
 * wasUserOverride() therefore says "no", so the correction was treated as a
 * passive AI confirmation: the birth gate held it back, and
 * generateAiOverrideRules skipped it — nothing was learned and the broad
 * rule kept misrouting (correcting client.example→A to B: 0 added, 0
 * disabled). It is a correction only when the originating rule still exists
 * and pointed somewhere other than where the user sent the mail.
 * Auto-propagated siblings (no userTouched) are not corrections.
 */
export function ruleCorrectionOf(
  item: PlanItem,
  rulesById: ReadonlyMap<string, Rule>,
): Rule | null {
  if (!item.originalRuleId || item.userTouched !== true || item.source === 'rule') return null
  if (item.action !== 'move' && item.action !== 'new_folder') return null
  const rule = rulesById.get(item.originalRuleId)
  if (!rule) return null
  const target = finalTargetPath(item)
  if (!target || target === rule.targetFolderPath) return null
  return rule
}

type SubjectMemoryMap = Record<string, ThreadMemoryEntry>

/**
 * Birth-gate evidence for a full_subject signal (2026-11-05): has this
 * normalized subject been filed to the same folder BEFORE — by a real
 * (plan-reviewed, non-harvested) filing, with no conflict — or, inside this
 * batch, by a different conversation?
 *
 * `prior` is subject memory as it was BEFORE this batch's filings were
 * recorded (startExecute snapshots it). Reading it afterwards — the old
 * behaviour — saw the current filing already counted and, worse, a
 * harvested-only entry with its `harvested` flag already cleared by that
 * very filing, so the harvested check was dead. Without a snapshot (older
 * callers) the post-recording map is used with this batch's own filings
 * subtracted from timesFiled; that fallback cannot recover the cleared
 * harvested flag.
 */
function makeObservationCheck(
  plan: readonly PlanItem[],
  results: readonly ExecuteItemResult[],
  prior: SubjectMemoryMap | undefined,
  postRecording: SubjectMemoryMap | undefined,
): (index: number, key: string, targetPath: string) => boolean {
  // This batch's successful filings by normalized subject — the same
  // eligibility execute uses when it records subject filings.
  const batch = new Map<string, Array<{ index: number; target: string; convId?: string }>>()
  for (let i = 0; i < plan.length; i++) {
    const status = results[i]?.status
    if (status !== 'moved' && status !== 'folder_created') continue
    const item = plan[i]!
    const key = normalizeSubject(item.emailSubject ?? '')
    const target = finalTargetPath(item)
    if (key.length < MIN_NORMALIZED_SUBJECT_LEN || !target) continue
    const list = batch.get(key) ?? []
    list.push({ index: i, target, convId: item.conversationId })
    batch.set(key, list)
  }
  return (index, key, targetPath) => {
    if (key.length < MIN_NORMALIZED_SUBJECT_LEN) return false
    const item = plan[index]!
    const entry = prior ? prior[key] : postRecording?.[key]
    if (entry) {
      const priorFilings = prior
        ? entry.timesFiled
        : entry.timesFiled - (batch.get(key)?.length ?? 0)
      const sameFolder =
        entry.folderPath === targetPath ||
        (!!entry.folderId && entry.folderId === item.targetFolderId)
      if (
        priorFilings >= 1 &&
        sameFolder &&
        // 主旨曾在資料夾間搖擺 → 壞規則候選，比照 matchThreadMemory 的
        // auto-route 門檻: 無衝突且非 harvested-only 才信任。
        (entry.conflictCount ?? 0) === 0 &&
        entry.harvested !== true
      ) {
        return true
      }
    }
    // Cross-thread recurrence inside this batch: another conversation with
    // the same subject went to the same folder. Same-conversation replies
    // don't count — they are one thread (thread memory's job).
    const self = item.conversationId
    if (!self) return false
    return (batch.get(key) ?? []).some(
      (o) => o.index !== index && o.target === targetPath && !!o.convId && o.convId !== self,
    )
  }
}

/**
 * True when a rule is a court-case structural identifier: a subject_keyword
 * whose signal IS a case number, or a compound whose subject condition is.
 * Used by the fast structural-override recovery (優化 2026-07).
 */
function isStructuralCourtCaseRule(r: Rule): boolean {
  if (r.type === 'subject_keyword') return courtCaseSignal(r.signal) != null
  if (r.type === 'compound') {
    const parsed = decodeCompound(r.signal)
    return (
      parsed?.conditions.some(
        (c) => c.type === 'subject_keyword' && courtCaseSignal(c.value) != null,
      ) ?? false
    )
  }
  return false
}

/** Semantic state of a rule — what undo compares before un-reviving it. */
function ruleStateFingerprint(r: Rule): string {
  return JSON.stringify([
    r.enabled,
    r.type,
    r.signal,
    r.targetFolderId,
    r.targetFolderPath,
    r.confidence,
    r.source,
    r.autoDisabledReason ?? null,
    r.autoDisabledAt ?? null,
    !!r.orphaned,
  ])
}

export async function generateAiConfirmedRules(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  opts: {
    /** Subject memory BEFORE this batch's filings were recorded (birth gate). */
    priorSubjectMemory?: SubjectMemoryMap
    /** Receives the write's plan under the rules lock, before it lands (R8). */
    onIntent?: (intent: LearnerIntent) => Promise<void>
  } = {},
): Promise<{
  added: number
  addedIds: string[]
  details: RuleLearnedDetail[]
  revived: number
  /** What undo needs to put each revived rule back to sleep (2026-11-06). */
  revivedRules: RevivedRuleUndo[]
}> {
  // Snapshot of current rules — only used for dedup / conflict checks.
  // The atomic addRulesFilteringTombstones call below ensures we don't
  // clobber concurrent writers and respects user tombstones.
  const [existing, settings, identity, postRecording] = await Promise.all([
    listRules(),
    getSettings(),
    getAccountCheckedIdentity(), // 2026-11-06 r2: never another account's
    opts.priorSubjectMemory ? Promise.resolve(undefined) : getSubjectMemory(),
  ])
  const internalDomainSet = learningExclusionDomains(settings.internalDomains, identity)
  const lookup = buildLearningLookup(existing)
  const rulesById = new Map(existing.map((r) => [r.id, r]))
  const observedBefore = makeObservationCheck(plan, results, opts.priorSubjectMemory, postRecording)
  const anchorsInBatch = batchAnchorTargets(plan, results, internalDomainSet, rulesById)
  const newRules: Rule[] = []
  const newKeys = new Set<string>()
  // Candidates subject to the same-signal conflict gate (see below).
  const conflictProne = new Set<Rule>()

  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]!
    const result = results[i]!
    if (result.status !== 'moved' && result.status !== 'folder_created') continue
    if (item.action !== 'move' && item.action !== 'new_folder') continue
    // User overrode the AI's verdict, or explicitly corrected the rule that
    // routed the row — handled by generateAiOverrideRules. Skip here so we
    // don't mis-attribute the user's choice to AI confidence (and so the
    // birth gate can't swallow a correction, 2026-11-05).
    if (wasUserOverride(item) || ruleCorrectionOf(item, rulesById)) continue
    // Routed by a hand-written rule that speaks about the subject / the case
    // (R6, 2026-11-06; narrowed the same day — isSubjectBearingUserRule): a
    // rule-routed row only learns structural signals — the WHOLE case. A
    // hand-written exception inside a case (domain + 「112訴204 帳單」 →
    // 帳務) taught (domain + 112訴204) → 帳務 here, and every later mail of
    // the case followed the exception into 帳務. Mail a hand-written PLAIN
    // domain / sender rule routed still teaches its case number (below).
    const via = routingRuleOf(item, rulesById)
    if (via && isSubjectBearingUserRule(via)) continue

    const sig = chooseLearningSignal(item, internalDomainSet, lookup, {
      batchAnchorTargets: anchorsInBatch,
    })
    if (!sig) continue
    // An anchor conflict with no subject to narrow to (2026-11-06): the
    // override path demotes on it; a passive confirmation learns nothing —
    // minting (or reviving) the plain rule would re-create the conflict.
    if ((sig.type === 'domain' || sig.type === 'sender') && sig.demoteOnly) continue

    // Structural unique identifier = court case number or Latin case code.
    // These can't legitimately map to two folders, so they're safe to learn
    // on FIRST sight — even from a rule- or thread-routed success, not only
    // AI-accepted mail (優化 2026-07, items「首見即學」+「thread→rule」).
    const isStructural =
      sig.type === 'case_code' ||
      ('featureKind' in sig && sig.featureKind === 'court_case')
    // A BODY-derived structural id (batch-3) is NOT trusted enough to earn the
    // rule/thread bypass below — bodies quote other parties' cases. Only a
    // SUBJECT structural id keeps the "first-sight learn from any route" power.
    const structuralTrusted = isStructural && !('fromBody' in sig && sig.fromBody)
    const fullSubject = 'featureKind' in sig && sig.featureKind === 'full_subject'
    // Thread-routed SUBJECT repeat (2026-11-05): subject memory routed this
    // row because the same normalized subject was filed here before, and the
    // user let it stand — that IS the second observation the birth gate
    // waits for. The source gate used to drop every thread-routed
    // non-structural row, and the second same-subject email is exactly what
    // thread memory routes — so in normal flow the gate never opened and
    // passive full_subject rules could not be minted at all. convId repeats
    // stay excluded: a reply in the same conversation is the same thread.
    const threadSubjectRepeat =
      item.source === 'thread' && item.threadMatch?.kind === 'subject' && item.userTouched !== true

    // Source gate:
    //   - ai / unresolved: always eligible (AI was confident & accepted, or
    //     the user resolved it — human-validated intent).
    //   - rule / thread: eligible for SUBJECT structural IDs. This lets a
    //     case number routed by thread memory (or by a DIFFERENT rule, e.g. a
    //     plain-domain rule — a hand-written one included, R6 above) mint a
    //     durable, syncable, cross-sender case rule — closing the
    //     thread-memory learning dead-end. Non-structural
    //     and body-derived rule/thread hits stay excluded (they'd overfit a
    //     broad subject, or trust a noisy body citation from an unvalidated
    //     route) — except the thread-subject repeat above, which still has
    //     to pass the birth gate.
    if (
      item.source !== 'ai' &&
      item.source !== 'unresolved' &&
      !structuralTrusted &&
      !(threadSubjectRepeat && fullSubject)
    ) {
      continue
    }

    const targetPath =
      item.action === 'move'
        ? item.targetFolderPath
        : item.suggestedParentPath && item.suggestedFolderName
        ? joinFolderPath(item.suggestedParentPath, item.suggestedFolderName.trim())
        : undefined
    if (!targetPath) continue

    // 兩次觀察出生閘門 (2026-07-16): full_subject 訊號不再首見即鑄。
    // 實測 325 條規則庫中 73% 是 compound/full_subject、其中 71% 從未
    // 命中過 — 大多數主旨一輩子只出現一次，同 thread 的回信本來就由
    // thread memory 接住。改為: 本批「之前」SubjectMemory 中同主旨已有
    // 至少一次真實歸檔到同一資料夾（或本批內另一個對話的同主旨信）才鑄
    // 永久規則（見 makeObservationCheck）。第一與第二次之間的跨 thread
    // 同主旨信多花一次 AI 呼叫，屬預期成本。
    //
    // 範圍: 被動 AI 確認 (source === 'ai') 與 thread 主旨重複。使用者主動
    // 解決 (unresolved) 與 override / 更正路徑 (generateAiOverrideRules)
    // 表達了明確人為意圖，維持立即學習。結構化案號 (court_case /
    // case_code) 不是 full_subject、不受影響。full_subject 訊號必然 ≥
    // MIN_NORMALIZED_SUBJECT_LEN（extractSubjectSignal 的門檻，2026-11-05
    // 起與 SubjectMemory 一致），所以閘門對每個訊號都可運作、不會無限封鎖。
    if ((item.source === 'ai' || item.source === 'thread') && fullSubject) {
      const key = normalizeSubject(item.emailSubject ?? '')
      if (!observedBefore(i, key, targetPath)) continue
    }

    // Dedup against ALL rules: an enabled twin, or a twin the USER disabled
    // (or one parked as case_closed / legacy_token), covers this triple. A
    // twin the engine auto-disabled for a revivable reason does NOT — the
    // candidate goes on to addRulesFilteringTombstones, which switches that
    // twin back on instead of adding a row (2026-11-05: treating every
    // disabled copy as covered made the revive branch unreachable).
    const sameSignal = lookup.bySignal.get(signalKey(sig.type, sig.signal)) ?? []
    const tripleKey = `${signalKey(sig.type, sig.signal)}::${targetPath}`
    const alreadyCovered =
      newKeys.has(tripleKey) ||
      sameSignal.some(
        (r) => r.targetFolderPath === targetPath && !isRevivableAutoDisabled(r),
      )
    if (alreadyCovered) continue

    // Gate: don't auto-create a rule when the same signal is already
    // routed elsewhere — prevents conflicting rules accumulating.
    //
    // Applies to signals where the same key could legitimately map to
    // different folders depending on context:
    //   - domain: same-client-multiple-cases (most common)
    //   - sender: one human emails about multiple topics
    //   - subject_keyword (full_subject): rare but possible if user
    //     treats two same-subject threads differently
    //   - compound (full_subject): same, with domain attached
    //
    // NOT applied to court_case / case_code variants — those are
    // structurally unique identifiers (an entire court case number
    // can't legitimately route to multiple folders). (Reviving a dormant
    // twin into a conflict is refused centrally in
    // addRulesFilteringTombstones, for every signal type.)
    const isConflictProneSignal =
      sig.type === 'domain' ||
      sig.type === 'sender' ||
      (sig.type === 'subject_keyword' && sig.featureKind === 'full_subject') ||
      (sig.type === 'compound' && sig.featureKind === 'full_subject')
    if (isConflictProneSignal) {
      // An orphaned rule routes nothing — no conflict (round 3, as in
      // chooseLearningSignal and the central revive guard).
      const hasConflict = sameSignal.some(
        (r) => r.enabled && !r.orphaned && r.targetFolderPath !== targetPath,
      )
      if (hasConflict) continue
    }

    // When the user resolved an "unresolved" item, they actively chose
    // even when AI was unsure — that's stronger evidence than passive
    // confirmation. Boost confidence +0.05. But the source label still
    // depends on whether the user agreed with AI's hint or overrode it:
    //
    //   - source === 'unresolved' + final target === aiOriginalTargetPath
    //     → user agreed with AI's low-confidence hint → 'ai_confirmed'
    //   - source === 'unresolved' + final target !== aiOriginalTargetPath
    //     → user picked a different folder than AI suggested (or AI gave
    //       no hint at all) → 'ai_overridden'
    //   - source === 'ai' (regular high-confidence AI suggestion, user
    //     didn't change) → 'ai_confirmed'
    //
    // wasUserOverride() is already filtered out at the top of the loop,
    // so by the time we reach here the user has NOT changed AI's hint
    // when there was one. For unresolved items where AI HAD a hint
    // (aiOriginalAction set) and user kept the same target → confirmed.
    // For unresolved items where AI gave up (no hint) but user picked a
    // folder → user-validated, so 'ai_overridden'.
    const userResolved = item.source === 'unresolved'
    const userOverroseAbsentAiHint =
      userResolved && item.aiOriginalAction === undefined
    // Confidence tier (post-2026-05-27 redesign):
    //   - case_code alone: 0.9 — structural identifier.
    //   - compound + court_case: 0.9 — structural identifier + domain.
    //   - compound + full_subject: 0.85 — exact subject + domain. Specific
    //     enough to safely auto-route; subject equality means same case.
    //   - subject_keyword + court_case: 0.85 — unique identifier even
    //     without domain.
    //   - subject_keyword + full_subject: 0.8 — same subject without
    //     domain still very specific (substring match by lowercased
    //     normalized subject).
    //   - sender (generic-provider exact address): 0.7
    //   - domain alone: 0.55 — broadest signal, kept below typical AI
    //     thresholds so subject signals can override.
    const baseConfidence =
      sig.type === 'case_code'
        ? 0.9
        : sig.type === 'compound'
          ? sig.featureKind === 'court_case'
            ? 0.9
            : 0.85 // full_subject
          : sig.type === 'subject_keyword'
            ? sig.featureKind === 'court_case'
              ? 0.85
              : 0.8 // full_subject
            : sig.type === 'sender'
              ? 0.7
              : 0.55
    // Body-derived case identifiers are noisier than subject ones — dock 0.1
    // so a bad body-learned rule crosses bumpRuleHits' error-rate disable
    // threshold sooner (batch-3). case_code 0.9→0.8, court-case 0.85→0.75.
    const bodyPenalty = 'fromBody' in sig && sig.fromBody ? 0.1 : 0
    const confidence =
      (userResolved ? Math.min(0.95, baseConfidence + 0.05) : baseConfidence) - bodyPenalty

    newKeys.add(tripleKey)
    newRules.push(
      newRule({
        type: sig.type,
        signal: sig.signal,
        // Strip sentinel/placeholder IDs so they don't pollute persisted rules.
        // Empty string is fine — match time falls back to path lookup.
        targetFolderId: sanitizeRuleTargetFolderId(item.targetFolderId),
        targetFolderPath: targetPath,
        // Confidence scales with signal specificity. Compound and case_code
        // are unique-per-case → high. Sender (exact From address) sits
        // between domain and subject_keyword — more specific than domain
        // but less authoritative than a subject pattern. Domain alone is
        // broad — kept intentionally below the AI threshold so
        // subject-specific rules and AI per-email judgements still win.
        confidence,
        // Source: see comment above. ai_overridden ONLY when user picked
        // a folder with no AI hint to compare against; otherwise
        // ai_confirmed (the user agreed with AI's verdict).
        source: userOverroseAbsentAiHint ? 'ai_overridden' : 'ai_confirmed',
      }),
    )
    if (isConflictProneSignal) conflictProne.add(newRules[newRules.length - 1]!)
  }

  // The gate above only sees STORED rules (2026-11-06): one batch learning
  // the same conflict-prone signal toward two folders shows the signal is
  // ambiguous by itself — learn neither side.
  const conflictTargets = new Map<string, Set<string>>()
  for (const r of conflictProne) {
    const key = signalKey(r.type, r.signal)
    const set = conflictTargets.get(key) ?? new Set<string>()
    set.add(r.targetFolderPath)
    conflictTargets.set(key, set)
  }
  const toAdd = newRules.filter(
    (r) => !conflictProne.has(r) || conflictTargets.get(signalKey(r.type, r.signal))!.size === 1,
  )

  // Filter against tombstones + write atomically. Closes the race where
  // a concurrent deleteRule could land between filter and write, letting
  // us auto-resurrect a rule the user just deleted. Revived twins are
  // reported separately and are NOT in addedIds: undo deletes (and
  // tombstones) addedIds, which must never hit a rule older than the batch
  // — revives are reversed instead (revivedRules, 2026-11-06).
  const onIntent = opts.onIntent
  const { added, reEnabled, revivedFrom } = await addRulesFilteringTombstones(
    toAdd,
    onIntent
      ? {
          beforeWrite: (p) =>
            onIntent({
              added: p.added.map((r) => ({ id: r.id, detail: ruleLearnedDetail(r) })),
              revived: reviveRecords(p.reEnabled, p.revivedFrom),
            }),
        }
      : undefined,
  )
  // Ids are threaded into the undo snapshot (H2) so 撤回 can delete the
  // rules this batch just minted — otherwise the reverted misfiling is
  // deterministically re-proposed next batch. `details` feed the DoneScreen
  // learned-rules list (覆核 P1-3).
  return {
    added: added.length,
    addedIds: added.map((r) => r.id),
    details: added.map((r) => ruleLearnedDetail(r)),
    revived: reEnabled.length,
    revivedRules: reviveRecords(reEnabled, revivedFrom),
  }
}

// Project a freshly-minted Rule down to the display-only fields the DoneScreen
// needs. Keeps ExecuteState small (no confidence / stats / ids the UI ignores).
function ruleLearnedDetail(r: Rule): RuleLearnedDetail {
  return {
    type: r.type,
    signal: r.signal,
    targetFolderPath: r.targetFolderPath,
    source: r.source,
  }
}

/**
 * Mirror of generateAiConfirmedRules for the "user disagreed" case.
 *
 * When the AI suggested move/new_folder → A but the user changed it to B and
 * the move to B executed successfully — or the user re-targeted a row a RULE
 * had routed to A (ruleCorrectionOf, 2026-11-05) — the user has just taught
 * us something concrete. We:
 *   1. Add a new ai_overridden rule (confidence 0.95, capped per type —
 *      user-validated > AI-validated), or switch an auto-disabled twin of it
 *      back on.
 *   2. Disable any *non-user_manual* existing rule that points the same
 *      (type, signal) at a different target, and — on a "先廣後窄" compound
 *      upgrade — the broad plain-domain / plain-sender rule it narrows,
 *      recording autoDisabledReason 'superseded' (or 'override' when nothing
 *      replaces it) so the rule can revive / be purged later.
 *
 * User_manual rules are sacred — never auto-disable them; the user knew what
 * they wanted when they hand-built that rule.
 *
 * Revives are reported like the AI-confirmed learner's (`revivedRules`, R7):
 * 撤回 puts them back to sleep under the same fingerprint guard.
 */
export async function generateAiOverrideRules(
  plan: PlanItem[],
  results: ExecuteItemResult[],
  opts: {
    /** Receives the write's plan under the rules lock, before it lands (R8). */
    onIntent?: (intent: LearnerIntent) => Promise<void>
  } = {},
): Promise<{
  added: number
  /** Ids of the rules minted — 撤回 deletes them, like the AI-confirmed ones (2026-11-06). */
  addedIds: string[]
  disabled: number
  details: RuleLearnedDetail[]
  revived: number
  revivedRules: RevivedRuleUndo[]
  /** Rules this pass switched off — 撤回 switches them back on (2026-11-06). */
  disabledRules: DisabledRuleUndo[]
}> {
  type Override = {
    type: LearningSignal['type']
    signal: string
    signalNorm: string
    targetPath: string
    targetId: string
    // Propagated from chooseLearningSignal — when true the demote loop
    // still retires the stale rule but we MUST NOT create a new rule
    // for the same signal (no usable subject feature was extractable,
    // so a new plain-domain / plain-sender rule would just re-create
    // the same conflict).
    demoteOnly: boolean
    // batch-3: structural id came from the body, not the subject → mint at a
    // lower confidence even on the (user-validated) override path.
    fromBody: boolean
    // Plain anchors of a compound signal (decoded ONCE here instead of per
    // rule inside the demote scan — 2026-11-05).
    anchors: Array<{ type: 'domain' | 'sender'; norm: string }>
  }
  const overrides: Override[] = []
  // Existing rules needed for chooseLearningSignal's conflict-prevention
  // check (priority 4 / 5) — if learning a plain-domain rule would
  // collide with an existing different-target rule, force compound.
  const [existing, settings, identity] = await Promise.all([
    listRules(),
    getSettings(),
    getAccountCheckedIdentity(), // 2026-11-06 r2: never another account's
  ])
  const internalDomainSet = learningExclusionDomains(settings.internalDomains, identity)
  const lookup = buildLearningLookup(existing)
  const rulesById = new Map(existing.map((r) => [r.id, r]))
  const anchorsInBatch = batchAnchorTargets(plan, results, internalDomainSet, rulesById)

  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]!
    const result = results[i]!
    if (result.status !== 'moved' && result.status !== 'folder_created') continue
    const aiOverride = wasUserOverride(item)
    const correction = !aiOverride && ruleCorrectionOf(item, rulesById) !== null
    if (!aiOverride && !correction) continue
    // We can only codify move-style overrides as rules. delete / skip are
    // valid user decisions but our Rule type has no action field — a rule
    // always means "move to target".
    if (item.action !== 'move' && item.action !== 'new_folder') continue
    const sig = chooseLearningSignal(item, internalDomainSet, lookup, {
      correction,
      batchAnchorTargets: anchorsInBatch,
    })
    if (!sig) continue
    const target = finalTargetPath(item)
    if (!target) continue
    const anchors: Override['anchors'] = []
    if (sig.type === 'compound') {
      for (const c of decodeCompound(sig.signal)?.conditions ?? []) {
        if (c.type === 'domain' || c.type === 'sender') {
          anchors.push({ type: c.type, norm: normalizeSignal(c.type, c.value) })
        }
      }
    }
    overrides.push({
      type: sig.type,
      signal: sig.signal,
      signalNorm: normalizeSignal(sig.type, sig.signal),
      targetPath: target,
      targetId: sanitizeRuleTargetFolderId(item.targetFolderId),
      // Only domain / sender variants of LearningSignal carry demoteOnly.
      // For other types this property is absent → default false.
      demoteOnly:
        (sig.type === 'domain' || sig.type === 'sender') && sig.demoteOnly === true,
      fromBody: 'fromBody' in sig && sig.fromBody === true,
      anchors,
    })
  }
  if (overrides.length === 0) {
    return {
      added: 0,
      addedIds: [],
      disabled: 0,
      details: [] as RuleLearnedDetail[],
      revived: 0,
      revivedRules: [],
      disabledRules: [],
    }
  }

  // Track the mutations so we can record audit events AFTER the mutex
  // releases. Doing it inside mutateRules would mean nested storage writes
  // before the lock yielded — annoying but not broken; we keep it outside
  // for cleaner ordering.
  const auditEvents: RuleEvent[] = []
  // Rules minted this pass, collected for the DoneScreen learned-rules list
  // (覆核 P1-3). Mirrors auditEvents' outer-scope + push-inside pattern; the
  // mutateRules closure runs once under the lock so no double-collection.
  const createdRules: Rule[] = []
  // Rules this pass revived (R7) / switched off (2026-11-06), for undo —
  // same pattern.
  let revivedRecords: RevivedRuleUndo[] = []
  let disabledRecords: DisabledRuleUndo[] = []
  const snapshotRule = (r: Rule): RuleSnapshot => ({
    type: r.type,
    signal: r.signal,
    targetFolderPath: r.targetFolderPath,
    confidence: r.confidence,
    source: r.source,
    enabled: r.enabled,
  })
  const outcome = await mutateRules(async (rules) => {
    // Tombstone read INSIDE the rules lock — closes the race where a
    // concurrent deleteRule could write a tombstone between our read
    // and our write, letting us auto-resurrect the deleted rule. Keys are
    // canonical (tombstoneKey) so tombstones from older builds still block.
    const tombKeys = new Set((await getRuleTombstones()).map(tombstoneKey))
    const allowedOverrides = overrides.filter(
      (o) => !tombKeys.has(`${o.type}::${o.signalNorm}::${o.targetPath}`),
    )
    if (allowedOverrides.length === 0) {
      return { next: rules, result: { added: 0, disabled: 0, revived: 0 } }
    }
    let added = 0
    let disabled = 0
    let revived = 0
    const next = [...rules]
    // Normalized once per rule (2026-11-05) — the scan below runs per
    // override, and compound normalization is a JSON parse.
    const norms = next.map((r) => normalizeSignal(r.type, r.signal))
    const now = Date.now()
    const nowIso = new Date(now).toISOString()

    for (const o of allowedOverrides) {
      // Demote sibling rules with same (type, signal) but different target —
      // they're the "AI's wrong answer" we want to retire — and, on a
      // "先廣後窄" compound upgrade (e.g. compound: vendor.com + 請款通知單),
      // any enabled plain-domain (or plain-sender) rule for the same anchor
      // that points elsewhere: the user's mental model is "the domain rule
      // wasn't enough anymore", so the broad rule should step aside. Future
      // mail of this domain whose subject doesn't match either compound rule
      // falls back to AI / conversation memory, which is the correct
      // behavior. user_manual is sacred — the user knew what they were
      // doing when they handbuilt it.
      //
      // Reason (2026-11-05): 'superseded' when a narrower / replacement rule
      // takes over, 'override' for a demoteOnly correction with nothing to
      // replace it. Either way it is an AUTO disable: revivable by learning,
      // purged after DORMANT_PURGE_ACTIVE_DAYS — before, the bare
      // enabled=false read as user-disabled and blocked the triple forever.
      for (let i = 0; i < next.length; i++) {
        const r = next[i]!
        const sameSignal = r.type === o.type && norms[i] === o.signalNorm
        const narrowedAnchor = o.anchors.some((a) => a.type === r.type && a.norm === norms[i])
        if (!sameSignal && !narrowedAnchor) continue
        if (r.targetFolderPath === o.targetPath) continue
        if (r.source === 'user_manual') continue
        if (!r.enabled) continue
        next[i] = {
          ...r,
          enabled: false,
          autoDisabledAt: nowIso,
          autoDisabledReason: o.demoteOnly ? 'override' : 'superseded',
        }
        disabled++
        auditEvents.push({
          kind: 'toggle',
          ruleId: r.id,
          at: now,
          actor: 'system',
          enabled: false,
          signal: r.signal,
          type: r.type,
          targetFolderPath: r.targetFolderPath,
        })
      }
      // Is the user's preferred mapping already represented? An enabled
      // twin covers it; a twin the user disabled (or one parked as
      // case_closed / legacy) still blocks it.
      let revivableAt = -1
      let covered = false
      for (let i = 0; i < next.length; i++) {
        const r = next[i]!
        if (r.type !== o.type || norms[i] !== o.signalNorm || r.targetFolderPath !== o.targetPath) {
          continue
        }
        if (isRevivableAutoDisabled(r)) {
          if (revivableAt < 0) revivableAt = i
        } else {
          covered = true
        }
      }
      if (covered) continue
      // Gate: don't auto-create plain-domain rules from a single user
      // override. Same domain often serves multiple cases — one override
      // isn't enough evidence to reroute every email from this domain.
      // Sender rules (specific From address) DO pass through: one
      // human-validated routing for an exact address is strong signal.
      if (o.type === 'domain') continue
      // demoteOnly: chooseLearningSignal detected a conflict but the
      // subject is too short to learn. We already demoted the stale rule
      // above; creating (or reviving) a rule for the same plain signal
      // would just re-create the same conflict pattern next batch. Applies
      // to the sender variant (domain already skipped above).
      if (o.demoteOnly) continue
      if (revivableAt >= 0) {
        // The engine had auto-disabled exactly this mapping (one-strike
        // 'override', 'superseded' …) and the user now chose it explicitly:
        // switch it back on rather than minting a twin (2026-11-05).
        const r = next[revivableAt]!
        next[revivableAt] = revivedRule(r, nowIso)
        revived++
        auditEvents.push({
          kind: 'toggle',
          ruleId: r.id,
          at: now,
          actor: 'system',
          enabled: true,
          signal: r.signal,
          type: r.type,
          targetFolderPath: r.targetFolderPath,
        })
        continue
      }
      // User-validated → higher confidence than ai_confirmed. Domain
      // rule creation is blocked above so we only get here for
      // compound / case_code / subject_keyword / sender — all of which
      // either uniquely identify a case or pinpoint a specific human.
      // A body-derived case id (batch-3) is docked to 0.85 even here: the
      // user validated the DESTINATION, but the body citation itself is
      // noisier than a subject one, so let the error-rate gate cull it
      // sooner if it misfires. applyConfidenceCap (2026-11-05): this path
      // used to write 0.95 straight through — above the sender (0.75) and
      // subject_keyword (0.9) ceilings every other write enforces.
      const created = applyConfidenceCap(
        newRule({
          type: o.type,
          signal: o.signal,
          targetFolderId: o.targetId,
          targetFolderPath: o.targetPath,
          confidence: o.fromBody ? 0.85 : 0.95,
          source: 'ai_overridden',
        }),
      )
      next.push(created)
      norms.push(o.signalNorm)
      createdRules.push(created)
      auditEvents.push({
        kind: 'create',
        ruleId: created.id,
        at: now,
        actor: 'system',
        after: snapshotRule(created),
      })
      added++
    }

    if (added === 0 && disabled === 0 && revived === 0) {
      return { next: rules, result: { added, disabled, revived } }
    }
    // For undo, the NET effect of the pass per rule (2026-11-06): a rule one
    // override demoted and a later one revived (or the reverse) is recorded
    // once, by where it started and where it ended, in the list of its
    // starting state (off → revivedRecords, on → disabledRecords). `next`
    // keeps the stored rules at their indexes; the minted ones follow them.
    revivedRecords = []
    disabledRecords = []
    for (let i = 0; i < rules.length; i++) {
      const before = rules[i]!
      const after = next[i]!
      if (
        before.enabled === after.enabled &&
        before.autoDisabledReason === after.autoDisabledReason &&
        before.autoDisabledAt === after.autoDisabledAt
      ) {
        continue
      }
      ;(before.enabled ? disabledRecords : revivedRecords).push(stateUndoRecord(before, after))
    }
    if (opts.onIntent && (createdRules.length > 0 || revivedRecords.length > 0 || disabledRecords.length > 0)) {
      await opts.onIntent({
        added: createdRules.map((r) => ({ id: r.id, detail: ruleLearnedDetail(r) })),
        revived: revivedRecords,
        disabled: disabledRecords,
      })
    }
    return { next, result: { added, disabled, revived } }
  })
  if (auditEvents.length > 0) {
    await recordRuleEvents(auditEvents).catch((e) =>
      console.warn('[mail-organizer] recordRuleEvents (ai-override) failed', e),
    )
  }
  return {
    ...outcome,
    addedIds: createdRules.map((r) => r.id),
    details: createdRules.map(ruleLearnedDetail),
    revivedRules: revivedRecords,
    disabledRules: disabledRecords,
  }
}

// ---- Undo ------------------------------------------------------------------

const UNDO_ALARM_NAME = 'undo-expire'

/** One rule state change of the batch, in the order it happened; `wasEnabled` = on before it. */
type RuleStateChange = { wasEnabled: boolean; rec: RevivedRuleUndo }

/**
 * One undo record per rule (2026-11-06): its state before its FIRST change of
 * the batch and its fingerprint after its LAST — a case rule the one-strike
 * disable switched off, the AI-confirmed learner revived and the override
 * learner superseded again is put back as it was before the batch, not
 * half-way. Filed by that starting state: off → revived, on → disabled.
 */
function foldRuleStateChanges(changes: readonly RuleStateChange[]): {
  revived: RevivedRuleUndo[]
  disabled: DisabledRuleUndo[]
} {
  const byId = new Map<string, { first: RuleStateChange; fingerprint: string }>()
  for (const c of changes) {
    const seen = byId.get(c.rec.ruleId)
    if (seen) seen.fingerprint = c.rec.fingerprint
    else byId.set(c.rec.ruleId, { first: c, fingerprint: c.rec.fingerprint })
  }
  const revived: RevivedRuleUndo[] = []
  const disabled: DisabledRuleUndo[] = []
  for (const { first, fingerprint } of byId.values()) {
    ;(first.wasEnabled ? disabled : revived).push({ ...first.rec, fingerprint })
  }
  return { revived, disabled }
}

async function captureUndoSnapshot(
  state: ExecuteState,
  plan: PlanItem[],
  mintedRuleIds: string[],
  stateChanges: readonly RuleStateChange[] = [],
): Promise<void> {
  // state.results is index-aligned with plan (initialState maps plan →
  // results 1:1 and the item loop writes results[i] for plan[i]) — walk by
  // index so each move can carry its learning-reversal keys (H2).
  const moves: UndoSnapshot['moves'] = []
  for (let i = 0; i < state.results.length; i++) {
    const r = state.results[i]!
    // Deletes are moves into Deleted Items now (2026-11-05) → undoable.
    // No learning keys: a delete teaches nothing (no thread memory) — except
    // that an auto-trash RULE's delete carries its ruleId (2026-11-06 r3), so
    // 撤回 counts as an override against a wrong auto-trash rule.
    if (r.status === 'deleted' && r.newMessageId) {
      const item = plan[i]
      moves.push({
        newMessageId: r.newMessageId,
        subject: r.subject,
        destinationFolderId: r.destinationFolderId ?? 'deleteditems',
        deleted: true,
        ...(isRuleRoutedDelete(item, r) ? { ruleId: item!.ruleId } : {}),
      })
      continue
    }
    if (
      (r.status !== 'moved' && r.status !== 'folder_created') ||
      !r.newMessageId ||
      !r.destinationFolderId
    ) {
      continue
    }
    const item = plan[i]
    if (r.unverifiedTarget) {
      // Nothing was learned from it (2026-11-06 r3) — nothing to reverse.
      moves.push({
        newMessageId: r.newMessageId,
        subject: r.subject,
        destinationFolderId: r.destinationFolderId,
        destinationFolderPath: r.destinationFolderPath,
      })
      continue
    }
    const normalized = normalizeSubject(item?.emailSubject ?? '')
    moves.push({
      newMessageId: r.newMessageId,
      subject: r.subject,
      destinationFolderId: r.destinationFolderId,
      destinationFolderPath: r.destinationFolderPath,
      conversationId: item?.conversationId,
      normalizedSubject:
        normalized.length >= MIN_NORMALIZED_SUBJECT_LEN ? normalized : undefined,
      ruleId: item?.source === 'rule' ? item.ruleId : undefined,
    })
  }

  if (moves.length === 0) return

  const now = Date.now()
  const { revived: revivedRules, disabled: disabledRules } = foldRuleStateChanges(stateChanges)
  const snap: UndoSnapshot = {
    batchId: String(now),
    createdAt: now,
    expiresAt: now + UNDO_WINDOW_MS,
    moves,
    ...(mintedRuleIds.length > 0 ? { mintedRuleIds } : {}),
    ...(revivedRules.length > 0 ? { revivedRules } : {}),
    ...(disabledRules.length > 0 ? { disabledRules } : {}),
    deletedCount: state.summary.deleted,
    newFolderCount: state.summary.foldersCreated,
  }
  await setUndoSnapshot(snap)

  // Schedule expiry — chrome.alarms guarantees firing even if SW idles.
  // delayInMinutes is fractional-allowed below 1.0 in MV3.
  try {
    await chrome.alarms.create(UNDO_ALARM_NAME, {
      when: snap.expiresAt,
    })
  } catch (e) {
    // Non-fatal — getUndoSnapshot() also evicts expired snapshots in passing.
    console.warn('[mail-organizer] undo alarm create failed (non-fatal)', e)
  }
}

/**
 * Undo of the batch's rule state changes (2026-11-06): each rule goes back
 * to its pre-batch enabled state and provenance — a revive back to sleep
 * (`revived`), a 'superseded' / 'override' / one-strike disable back on
 * (`disabled`) — unless it changed since (fingerprint mismatch: the user
 * edited / toggled it, sync replaced it, reconcile re-pointed it). Counters
 * (matchCount / overrideCount / statsResetAt) stay as they are.
 */
async function restoreRuleStates(
  revived: readonly RevivedRuleUndo[],
  disabled: readonly DisabledRuleUndo[],
): Promise<void> {
  const byId = new Map<string, { enabled: boolean; rec: RevivedRuleUndo }>()
  for (const rec of revived) byId.set(rec.ruleId, { enabled: false, rec })
  for (const rec of disabled) byId.set(rec.ruleId, { enabled: true, rec })
  const now = Date.now()
  const events: RuleEvent[] = []
  await mutateRules((rules) => {
    let changed = false
    const next = rules.map((r) => {
      const want = byId.get(r.id)
      if (!want || ruleStateFingerprint(r) !== want.rec.fingerprint) return r
      changed = true
      if (r.enabled !== want.enabled) {
        events.push({
          kind: 'toggle',
          ruleId: r.id,
          at: now,
          actor: 'system',
          enabled: want.enabled,
          type: r.type,
          signal: r.signal,
          targetFolderPath: r.targetFolderPath,
        })
      }
      return {
        ...r,
        enabled: want.enabled,
        autoDisabledReason: want.rec.autoDisabledReason,
        autoDisabledAt: want.rec.autoDisabledAt,
      }
    })
    return { next: changed ? next : rules, result: undefined }
  })
  if (events.length > 0) {
    await recordRuleEvents(events).catch((e) =>
      console.warn('[mail-organizer] recordRuleEvents (undo rule states) failed', e),
    )
  }
}

export function installUndoExpireListener(): void {
  chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === UNDO_ALARM_NAME) {
      void clearUndoSnapshot().catch((e) =>
        console.warn('[mail-organizer] undo clearSnapshot failed (non-fatal)', e),
      )
    }
  })
}

export type UndoResult = {
  attempted: number
  restored: number
  failed: number
  errors: Array<{ subject: string; message: string }>
  /**
   * Why 撤回 stopped early (2026-11-06 r2): no Outlook token. Every move it
   * never got to carries this same message in `errors`.
   */
  message?: string
}

/**
 * Reverse the most-recent batch's moves — and, since 2026-11-05, its deletes
 * (moves into Deleted Items) — by moving each message back to inbox.
 * No-op if snapshot is missing/expired. Clears the snapshot on completion
 * (including partial failures — a follow-up retry would address pre-moved
 * messages a second time, which Outlook rejects with 404).
 */
export async function executeUndo(api: OutlookApi): Promise<UndoResult> {
  // Synchronous guards before the first await (2026-11-05, invariant 6): a
  // double trigger (button + `u` key) used to replay every move of the
  // snapshot twice, and an undo racing a batch start could move mail the
  // new batch was about to process.
  if (undoInFlight) throw new Error('撤回進行中')
  if (executeClaimed) throw new Error('批次執行中,無法撤回')
  undoInFlight = true
  try {
    return await runUndo(api)
  } finally {
    undoInFlight = false
  }
}

async function runUndo(api: OutlookApi): Promise<UndoResult> {
  const snap = await getUndoSnapshot()
  if (!snap) return { attempted: 0, restored: 0, failed: 0, errors: [] }

  const errors: UndoResult['errors'] = []
  let restored = 0
  const restoredMoves: UndoSnapshot['moves'] = []
  let stopped: string | undefined
  await holdKeepAlive()
  try {
    for (let i = 0; i < snap.moves.length; i++) {
      const m = snap.moves[i]!
      try {
        await api.moveMessage(m.newMessageId, 'inbox')
        restored++
        restoredMoves.push(m)
      } catch (e) {
        // Circuit breaker (2026-11-06 r2), as in execute: with no Outlook
        // token no move can go out, and each one used to wait out its own
        // ~10 s token timeout. This and every remaining move get one message.
        if (e instanceof OutlookTokenUnavailableError) {
          stopped = undoTokenOutageMessage(e)
          for (const rest of snap.moves.slice(i)) errors.push({ subject: rest.subject, message: stopped })
          break
        }
        const message =
          e instanceof OutlookError
            ? `[${e.status}] ${e.message}`
            : e instanceof Error
            ? e.message
            : String(e)
        errors.push({ subject: m.subject, message })
      }
    }

    // ---- Learning reversal (H2) ---------------------------------------
    // The batch flushed its learning BEFORE the undo window even opened
    // (rule hits, thread-memory filings, freshly-minted ai_confirmed
    // rules). Reverting only the emails left all of that in place, so the
    // same misfiling was deterministically re-proposed next batch — and
    // got STICKIER each cycle. Reverse it for the moves that actually
    // restored; skip entirely when nothing came back (the mailbox still
    // matches what was learned). Each step is independent + non-fatal.
    // Restored DELETES (2026-11-05) taught nothing, so they don't count:
    // un-deleting alone must not delete the batch's minted rules.
    const restoredFilings = restoredMoves.filter((m) => !m.deleted)
    // (a) Undo counts as an override against the driving rules — feeds the
    //     empirical-accuracy demotion (high error rate → auto disable), same
    //     as editing the row before execute would have. That includes an
    //     auto-trash rule's restored delete (2026-11-06 r3: it carries its
    //     ruleId now), so a wrong auto-trash rule can be retired; the rest of
    //     the reversal below stays for filings only (a delete taught nothing).
    const overrideTally = new Map<string, number>()
    for (const m of restoredMoves) {
      if (m.ruleId) overrideTally.set(m.ruleId, (overrideTally.get(m.ruleId) ?? 0) + 1)
    }
    if (overrideTally.size > 0) {
      await bumpRuleOverrides(overrideTally).catch((e) =>
        console.warn('[mail-organizer] undo override tally failed (non-fatal)', e),
      )
    }
    if (restoredFilings.length > 0) {
      // (b) Delete the rules THIS batch's learning minted — AI-confirmed and
      //     (2026-11-06) override. deleteRule writes a tombstone, so the
      //     same wrong mapping can't be auto-re-learned next batch either.
      for (const id of snap.mintedRuleIds ?? []) {
        await deleteRule(id).catch((e) =>
          console.warn('[mail-organizer] undo minted-rule delete failed (non-fatal)', e),
        )
      }
      // (b2) Put every rule the batch revived or switched off back the way
      //      it was (2026-11-06) — the revived dormant rules back to sleep
      //      (left live, a revived rule re-proposed the undone misfiling
      //      next batch), the ones superseded / demoted / one-strike
      //      disabled back on. Not deleted: they predate the batch.
      if ((snap.revivedRules?.length ?? 0) > 0 || (snap.disabledRules?.length ?? 0) > 0) {
        await restoreRuleStates(snap.revivedRules ?? [], snap.disabledRules ?? []).catch((e) =>
          console.warn('[mail-organizer] undo rule-state reversal failed (non-fatal)', e),
        )
      }
      // (c) Contest the thread-memory entries this batch reinforced —
      //     conflictCount>0 stops the pre-filter from trusting them until
      //     they re-earn a stable streak.
      await markThreadMemoryConflicts({
        convIds: [...new Set(restoredMoves.map((m) => m.conversationId).filter((v): v is string => !!v))],
        normalizedSubjects: [
          ...new Set(restoredMoves.map((m) => m.normalizedSubject).filter((v): v is string => !!v)),
        ],
      }).catch((e) =>
        console.warn('[mail-organizer] undo thread-memory contest failed (non-fatal)', e),
      )
    }
  } finally {
    // The hold is released whatever these do (2026-11-06 r2): a rejected
    // storage call used to skip releaseKeepAlive, and the 20 s heartbeat then
    // kept the worker alive for the rest of the browser session.
    try {
      await clearUndoSnapshot()
      try {
        await chrome.alarms.clear(UNDO_ALARM_NAME)
      } catch {
        // ignore
      }
    } finally {
      await releaseKeepAlive()
    }
  }

  return {
    attempted: snap.moves.length,
    restored,
    failed: errors.length,
    errors,
    ...(stopped ? { message: stopped } : {}),
  }
}

/**
 * 撤回 stopped by a token outage (2026-11-06 r2). The snapshot is cleared
 * whatever happens (a replay would address already-restored mail again), so
 * there is no retry — the message says so and where the mail still is.
 */
function undoTokenOutageMessage(e: OutlookTokenUnavailableError): string {
  const reason = e.reason.trim().replace(/[。.！!\s]+$/u, '')
  return `無法取得 Outlook 授權,撤回已停止,這封信沒有移回收件匣、仍在歸檔後的位置${reason ? `(${reason})` : ''}。撤回無法重試,請到 Outlook 手動移回收件匣。`
}

export async function dismissUndo(): Promise<void> {
  await clearUndoSnapshot()
  try {
    await chrome.alarms.clear(UNDO_ALARM_NAME)
  } catch {
    // ignore
  }
}
