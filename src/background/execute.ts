// Execute flow.
//
// Drives a PlanItem[] through Outlook REST + persists progress to
// chrome.storage.session so the popup can poll and resume across reopens.
//
// Per-item handling:
//   move        → POST /messages/{id}/move (uses targetFolderId, or resolves from path)
//   delete      → DELETE /messages/{id}
//   new_folder  → POST createChildFolder (or top) + POST /messages/{id}/move
//   skip        → no-op
//
// Rule feedback:
//   - rule-sourced success → bumpRuleHit(ruleId)
//   - ai-sourced successful move → auto-create domain rule (confidence 0.7,
//     source: 'ai_confirmed') if no equivalent rule already covers it.
//     Skips senders on settings.internalDomains (firm-internal mail —
//     colleagues send across many cases; domain-routing is misleading).

import { findFolderByPath, joinFolderPath, OutlookApi, OutlookError } from '@/shared/outlook-api'
import {
  addRulesFilteringTombstones,
  bumpRuleHits,
  bumpRuleOverrides,
  decodeCompound,
  encodeCompound,
  extractCaseCodes,
  extractCourtCaseNumbers,
  extractSubjectSignal,
  isGenericProviderDomain,
  listRules,
  mutateRules,
  newRule,
  normalizeSignal,
} from '@/shared/rules'
import {
  addFolderToCache,
  addToRecentlyProcessed,
  addToSkipHistory,
  bumpMetrics,
  clearUndoSnapshot,
  getRuleTombstones,
  getSettings,
  getUndoSnapshot,
  recordConversationFilings,
  recordFolderActivityFromBatch,
  recordRuleEvents,
  recordSubjectFilings,
  setUndoSnapshot,
} from '@/shared/storage'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from '@/shared/normalize'
import {
  type ExecuteItemResult,
  type ExecuteItemStatus,
  type ExecuteState,
  type MailFolderNode,
  type PlanItem,
  type Rule,
  type RuleEvent,
  type RuleSnapshot,
  type UndoSnapshot,
  UNDO_WINDOW_MS,
} from '@/shared/types'
import { holdKeepAlive, releaseKeepAlive } from './keep-alive'

function isAbortError(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError'
}

// In-memory controller for the currently-running execute. Survives within a
// single service-worker lifetime; SW restart kills both the execute Promise
// AND this controller together, so no consistency issue.
let currentController: AbortController | null = null

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
  await chrome.storage.session.remove(STATE_KEY)
}

/**
 * Detect an `inProgress = true` state that has no live executor and recover.
 * Called on SW module load to handle MV3 idle-kill mid-run. Items in
 * `processing` become `error` (ambiguous — the operation may have partially
 * applied server-side); items in `queued` become `cancelled` (never started,
 * shouldn't be counted as a failure). Without this, `isExecuteRunning()`
 * would block the next batch forever.
 */
export async function recoverStaleExecuteState(): Promise<{ recovered: boolean; stuck: number }> {
  const state = await getExecuteState()
  if (!state?.inProgress) return { recovered: false, stuck: 0 }

  let cancelledCount = 0
  let errorCount = 0
  const results = state.results.map((r) => {
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
  const stuck = errorCount + cancelledCount

  await saveExecuteState({
    ...state,
    inProgress: false,
    cancelRequested: false,
    finishedAt: state.finishedAt ?? Date.now(),
    results,
    summary: {
      ...state.summary,
      cancelled: state.summary.cancelled + cancelledCount,
      errors: state.summary.errors + errorCount,
    },
  })
  console.warn(
    `[mail-organizer] recovered stuck executeState: ${errorCount} error, ${cancelledCount} cancelled`,
  )
  return { recovered: true, stuck }
}

export async function requestCancel(): Promise<void> {
  // Two-step write was racy: between read and write, a new batch could start
  // and our naive replace would clobber the new state. Now we re-read inside
  // the write and only flip cancelRequested when the same batch is still
  // running. The in-memory controller abort still fires unconditionally —
  // worst case it's a no-op against a stale controller, which is safe.
  const state = await getExecuteState()
  if (state && state.inProgress) {
    const cur = await getExecuteState()
    if (cur && cur.inProgress && cur.startedAt === state.startedAt) {
      await saveExecuteState({ ...cur, cancelRequested: true })
    }
    currentController?.abort()
  }
}

export async function isExecuteRunning(): Promise<boolean> {
  const state = await getExecuteState()
  return !!state?.inProgress
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
  if (await isExecuteRunning()) {
    throw new Error('已有批次執行中')
  }

  const state = initialState(plan)
  await saveExecuteState(state)

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

  const controller = new AbortController()
  currentController = controller
  await holdKeepAlive()

  // Tally rule hits as they happen, flush once at end. Atomic write of the
  // accumulated bumps avoids 50+ sequential read-modify-write cycles.
  const ruleHitTally = new Map<string, number>()
  // Ids that 404'd ErrorItemNotFound during this batch — provably dead,
  // recorded into the recentlyProcessed ledger alongside the successful
  // moves (see the post-loop write).
  const alreadyGoneIds: string[] = []
  // Rule overrides: a rule fired at classify time but the user edited
  // the resulting PlanItem before execute (changed action / target).
  // Detected via item.source !== 'rule' but item.originalRuleId still
  // present (PlanRow.tsx preserves it on edit via {...item} spread).
  // Bumped on the underlying rule's overrideCount so we can compute
  // empirical accuracy.
  const ruleOverrideTally = new Map<string, number>()
  // Thread memory writes: every successful move/folder_created teaches
  // us "this conversation / subject was filed here". Pre-filter step in
  // classifyPreflight uses this to short-circuit AI on follow-up
  // emails. Two tallies because subject map tracks conflictCount —
  // recordSubjectFilings handles the conflict logic.
  const convFilings: Array<{ convId: string; folderId: string; folderPath: string }> = []
  const subjectFilings: Array<{
    normalizedSubject: string
    folderId: string
    folderPath: string
  }> = []

  try {
    for (const i of indexOrder) {
      // Cancellation check between items (also catches abort mid-call below)
      const live = await getExecuteState()
      if (live?.cancelRequested) {
        state.results[i] = { ...state.results[i]!, status: 'cancelled' }
        state.summary.cancelled++
        break
      }

      const item = plan[i]!
      state.current = i
      state.results[i] = { ...state.results[i]!, status: 'processing' }
      await saveExecuteState(state)

      try {
        const outcome = await executeItem(item, tree, api, controller.signal)
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
        }
        // 404 ErrorItemNotFound = positive confirmation this id is dead.
        // Collect for the recentlyProcessed ledger below — without this,
        // a ghost entry that Outlook keeps listing would loop forever:
        // 404 → soft-skip → never recorded → reappears next batch.
        if (outcome.alreadyGone) alreadyGoneIds.push(item.emailId)

        if (status === 'moved') state.summary.moved++
        else if (status === 'deleted') state.summary.deleted++
        else if (status === 'folder_created') {
          state.summary.foldersCreated++
          state.summary.moved++ // also counts as a move
        } else if (status === 'skipped') state.summary.skipped++

        if (
          item.source === 'rule' &&
          item.ruleId &&
          (status === 'moved' || status === 'folder_created')
        ) {
          // Accumulate in-memory and flush once at batch end (see below).
          // Previously each hit ran an isolated mutateRules() — for a 50-hit
          // batch on a 1000-rule library that was 50 sequential storage
          // write cycles, adding 5-10s to the batch.
          ruleHitTally.set(item.ruleId, (ruleHitTally.get(item.ruleId) ?? 0) + 1)
        } else if (
          item.source !== 'rule' &&
          item.originalRuleId &&
          (status === 'moved' || status === 'folder_created' || status === 'deleted')
        ) {
          // The rule fired at classify time and the user then changed
          // the action / target — count as an override against the
          // originating rule. Status filter excludes skip (user opted
          // out, not necessarily "rule was wrong") and execute errors.
          ruleOverrideTally.set(
            item.originalRuleId,
            (ruleOverrideTally.get(item.originalRuleId) ?? 0) + 1,
          )
        }

        // Thread memory writeback. Record convId + normalized subject
        // for every successful move into a real folder. We do this even
        // for thread-sourced rows so the entry's timesFiled keeps
        // climbing (used by future tooling / debugging to spot strong
        // patterns). destinationFolderId comes from the move result
        // because Outlook assigns a NEW id after move — using
        // item.targetFolderId would record the source-side id and
        // pre-filter wouldn't match next batch.
        if (
          (status === 'moved' || status === 'folder_created') &&
          destinationFolderId &&
          destinationFolderPath
        ) {
          if (item.conversationId) {
            convFilings.push({
              convId: item.conversationId,
              folderId: destinationFolderId,
              folderPath: destinationFolderPath,
            })
          }
          const normalized = normalizeSubject(item.emailSubject ?? '')
          if (normalized.length >= MIN_NORMALIZED_SUBJECT_LEN) {
            subjectFilings.push({
              normalizedSubject: normalized,
              folderId: destinationFolderId,
              folderPath: destinationFolderPath,
            })
          }
        }
      } catch (e) {
        if (isAbortError(e)) {
          state.results[i] = { ...state.results[i]!, status: 'cancelled' }
          state.summary.cancelled++
          await saveExecuteState(state)
          break
        }
        // F5 (2026-06-03): a non-idempotent op (move/delete) that failed
        // with `uncertain` — the request reached Outlook and MAY have
        // committed server-side before the response was lost (network
        // blip / 5xx). outlook-api deliberately does NOT retry these
        // (retrying would 404 on the now-stale id). Marking them as a
        // hard `error` is wrong: it shows a scary red row AND makes the
        // item a 重試 candidate, reproducing the exact "已經移動成功卻叫我
        //再歸檔" loop this whole fix targets. Treat as a soft skip with
        // an honest "needs confirmation" note — the email is either at
        // its destination (move succeeded) or still in inbox
        // (re-classifiable next run); both are safe, neither warrants a
        // failure tally or a retry.
        if (e instanceof OutlookError && e.uncertain) {
          state.results[i] = {
            ...state.results[i]!,
            status: 'skipped',
            message:
              '網路中斷,這封可能已成功移動、也可能仍在收件匣。請到 Outlook 確認;若仍在收件匣,下次歸類會再處理。',
          }
          state.summary.skipped++
          await saveExecuteState(state)
          continue
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

      await saveExecuteState(state)
    }

    // Mark any never-started items (after break) as cancelled
    for (let i = 0; i < state.results.length; i++) {
      if (state.results[i]!.status === 'queued') {
        state.results[i] = { ...state.results[i]!, status: 'cancelled' }
        state.summary.cancelled++
      }
    }

    // F9 (2026-06-03): flush OVERRIDES before HITS. bumpRuleHits's
    // confidence-promotion gate reads `overrideCount / matchCount` to
    // decide whether a rule has earned a confidence bump. If hits flush
    // first, that gate sees the pre-batch overrideCount and can promote
    // a rule the user overrode in THIS very batch (its true post-batch
    // error rate would exceed the gate). Applying overrides first means
    // the promotion gate sees the up-to-date error rate. (Both write
    // different fields on the same rule, so ordering is otherwise free.)
    if (ruleOverrideTally.size > 0) {
      try {
        await bumpRuleOverrides(ruleOverrideTally)
      } catch (e) {
        console.warn('[mail-organizer] bumpRuleOverrides flush failed (non-fatal)', e)
      }
    }
    // Flush accumulated rule-hit bumps in a single atomic update.
    if (ruleHitTally.size > 0) {
      try {
        await bumpRuleHits(ruleHitTally)
      } catch (e) {
        console.warn('[mail-organizer] bumpRuleHits flush failed (non-fatal)', e)
      }
    }
    // Record the emails we actually moved/deleted into the recently-
    // processed ledger so the NEXT batch's inbox fetch filters them out.
    // Outlook's store is eventually consistent: for a short window after a
    // move/delete the inbox listing can still return the just-handled email
    // with its now-dead id, and the next batch would try to re-move it and
    // hit 404 ErrorItemNotFound — which is exactly the "already-moved
    // emails reappear in the continue list" bug. We record only DEFINITE
    // outcomes (moved / deleted / folder_created); `skipped` (incl. the
    // uncertain-network soft-skip) is left out so a genuinely-unmoved
    // email can still be retried next batch.
    const processedIds = [
      ...state.results
        .filter(
          (r) =>
            r.status === 'moved' ||
            r.status === 'deleted' ||
            r.status === 'folder_created',
        )
        .map((r) => r.emailId),
      // Plus ids Outlook confirmed dead via 404 this batch (already
      // moved/deleted earlier) — re-recording refreshes their TTL so a
      // long-lagging listing keeps them excluded round after round.
      ...alreadyGoneIds,
    ].filter(Boolean)
    if (processedIds.length > 0) {
      try {
        await addToRecentlyProcessed(processedIds)
      } catch (e) {
        console.warn('[mail-organizer] addToRecentlyProcessed failed (non-fatal)', e)
      }
    }
    // Flush thread memory — conversation + normalized subject.
    if (convFilings.length > 0) {
      try {
        await recordConversationFilings(convFilings)
      } catch (e) {
        console.warn('[mail-organizer] recordConversationFilings failed (non-fatal)', e)
      }
    }
    if (subjectFilings.length > 0) {
      try {
        await recordSubjectFilings(subjectFilings)
      } catch (e) {
        console.warn('[mail-organizer] recordSubjectFilings failed (non-fatal)', e)
      }
    }

    // Generate ai_confirmed + ai_overridden rules from successful runs —
    // both non-fatal. ai_confirmed captures "AI was right, lock it in";
    // ai_overridden captures "user changed AI's verdict, learn the new
    // mapping AND demote any existing rule that pointed to the wrong
    // target". They're mutually exclusive per item via wasUserOverride.
    let aiConfirmedAdded = 0
    let aiOverriddenAdded = 0
    try {
      aiConfirmedAdded = await generateAiConfirmedRules(plan, state.results)
    } catch (e) {
      console.warn('[mail-organizer] generateAiConfirmedRules failed (non-fatal)', e)
    }
    try {
      const r = await generateAiOverrideRules(plan, state.results)
      aiOverriddenAdded = r.added
      if (r.disabled > 0) {
        console.info(
          `[mail-organizer] ai_overridden disabled ${r.disabled} stale rule(s) pointing at the AI's old wrong target`,
        )
      }
    } catch (e) {
      console.warn('[mail-organizer] generateAiOverrideRules failed (non-fatal)', e)
    }
    state.rulesAdded = aiConfirmedAdded + aiOverriddenAdded

    // Conflict prevention moved into chooseLearningSignal (2026-05-27
    // redesign). When the user routes a same-domain email to a different
    // folder than an existing plain-domain rule's target, the learning
    // path upgrades to `compound (domain + 整段主旨)` and demotes the
    // stale plain-domain rule synchronously. No post-batch sweep needed;
    // conflicts can no longer accumulate via the normal learning flow.

    // Remember which emails the user chose to keep in inbox — next classify
    // pass auto-excludes them so they don't show up again to be re-decided.
    const skippedIds = state.results
      .filter((r) => r.status === 'skipped')
      .map((r) => r.emailId)
      .filter((id) => id)
    if (skippedIds.length > 0) {
      await addToSkipHistory(skippedIds).catch((e) =>
        console.warn('[mail-organizer] addToSkipHistory failed (non-fatal)', e),
      )
    }

    // Tally per-folder mail counts from this batch and feed the IdleScreen's
    // 「近日活動」quick-jump panel. Joins each result's emailId back to its
    // plan item to pull the resolved target path/id (results carry status
    // but not the target path).
    try {
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
      for (let i = 0; i < state.results.length; i++) {
        const r = state.results[i]!
        if (r.status !== 'moved' && r.status !== 'folder_created') continue
        const planItem = plan[i]
        if (!planItem) continue
        const folderId = sanitizeRuleTargetFolderId(planItem.targetFolderId)
        const folderPath = finalTargetPath(planItem)
        if (!folderId || !folderPath) continue
        const prev = folderTally.get(folderId)
        // Capture the moved email's subject/from so the "近日活動" panel
        // can show "what's the latest in this folder" without making the
        // user click 重新整理 to fire a separate Graph API scan. Last
        // result for the folder wins (plan iteration order is by index;
        // for batches sorted newest-first this is the latest email, for
        // oldest-first batches it'll be the most-recent one we moved
        // either way — both are "latest user saw" in practice).
        const latestMessage = {
          subject: (planItem.emailSubject ?? '').slice(0, 120),
          from: planItem.emailFrom ?? '',
          receivedAt: batchAt,
        }
        if (prev) {
          prev.count++
          prev.latestMessage = latestMessage
        } else {
          folderTally.set(folderId, {
            folderId,
            folderPath,
            count: 1,
            latestMessage,
          })
        }
      }
      if (folderTally.size > 0) {
        await recordFolderActivityFromBatch([...folderTally.values()], batchAt)
      }
    } catch (e) {
      console.warn('[mail-organizer] folder activity record failed (non-fatal)', e)
    }

    state.inProgress = false
    state.finishedAt = Date.now()
    await saveExecuteState(state)

    await bumpMetrics({
      moved: state.summary.moved,
      deleted: state.summary.deleted,
      foldersCreated: state.summary.foldersCreated,
      errors: state.summary.errors,
    })

    // Capture undo snapshot — only when there's actually something to undo.
    // Deletes and folder creates aren't reversed here (delete is recoverable
    // via Outlook's own Deleted Items; folder cleanup would amplify undo's
    // blast radius beyond what users expect from "撤回").
    try {
      await captureUndoSnapshot(state)
    } catch (e) {
      console.warn('[mail-organizer] captureUndoSnapshot failed (non-fatal)', e)
    }
  } catch (e) {
    state.inProgress = false
    state.finishedAt = Date.now()
    await saveExecuteState(state)
    throw e
  } finally {
    currentController = null
    await releaseKeepAlive()
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

async function executeItem(
  item: PlanItem,
  tree: MailFolderNode[],
  api: OutlookApi,
  signal: AbortSignal,
): Promise<ItemOutcome> {
  // Each network sub-call checks signal first. new_folder does
  // create + move (sometimes also list-siblings on 409) — without these
  // checks an in-flight cancel would let createFolder + moveMessage both
  // complete before the next item-boundary cancel test caught up.
  signal.throwIfAborted()

  if (item.action === 'skip') return { status: 'skipped' }

  if (item.action === 'delete') {
    try {
      await api.deleteMessage(item.emailId, signal)
      return { status: 'deleted' }
    } catch (e) {
      if (isAlreadyMovedError(e)) {
        // Delete target gone = deletion's already done by something
        // else (or by an earlier lost-response delete from us). Soft
        // success — don't surface as red error.
        return { status: 'skipped', message: ALREADY_MOVED_MESSAGE, alreadyGone: true }
      }
      throw e
    }
  }

  if (item.action === 'move') {
    const targetId = resolveTargetFolderId(item, tree)
    if (!targetId) throw new Error('目標資料夾未指定或無法在資料夾樹中找到')
    try {
      const res = await api.moveMessage(item.emailId, targetId, signal)
      return {
        status: 'moved',
        newMessageId: res.Id,
        destinationFolderId: targetId,
        destinationFolderPath: item.targetFolderPath,
      }
    } catch (e) {
      if (isAlreadyMovedError(e)) {
        return { status: 'skipped', message: ALREADY_MOVED_MESSAGE, alreadyGone: true }
      }
      throw e
    }
  }

  if (item.action === 'new_folder') {
    const name = item.suggestedFolderName?.trim()
    if (!name) throw new Error('未指定新資料夾名稱')
    let parentId: string | undefined
    if (item.suggestedParentPath) {
      const parent = findFolderByPath(tree, item.suggestedParentPath)
      if (parent) parentId = parent.id
      else throw new Error(`找不到父資料夾「${item.suggestedParentPath}」`)
    }

    // In-batch dedup: prior new_folder items have already created what's
    // needed. Reorder above guarantees those run first.
    const fullPath = joinFolderPath(item.suggestedParentPath, name)
    const existing = findFolderByPath(tree, fullPath)
    if (existing) {
      signal.throwIfAborted()
      try {
        const res = await api.moveMessage(item.emailId, existing.id, signal)
        return {
          status: 'moved',
          newMessageId: res.Id,
          destinationFolderId: existing.id,
          destinationFolderPath: fullPath,
        }
      } catch (e) {
        if (isAlreadyMovedError(e)) {
          return { status: 'skipped', message: ALREADY_MOVED_MESSAGE, alreadyGone: true }
        }
        throw e
      }
    }

    let created: { Id: string; DisplayName: string; ParentFolderId?: string }
    let recoveredFromExisting = false
    try {
      signal.throwIfAborted()
      created = parentId
        ? await api.createChildFolder(parentId, name, signal)
        : await api.createTopFolder(name, signal)
    } catch (e) {
      if (isAbortError(e)) throw e
      if (e instanceof OutlookError && e.status === 409) {
        console.warn(`[mail-organizer] createFolder 409 for ${fullPath}; recovering via listChildFolders`)
        signal.throwIfAborted()
        const siblings = parentId
          ? await api.listChildFolders(parentId)
          : await api.listTopFolders()
        const match = siblings.find((f) => f.DisplayName === name)
        if (!match) {
          throw new Error(`Outlook 回 409 但 listChildFolders 找不到同名「${name}」資料夾`)
        }
        created = match
        recoveredFromExisting = true
      } else {
        throw e
      }
    }

    // Splice into in-memory tree + persist (idempotent)
    spliceFolderIntoTree(tree, created, parentId)
    addFolderToCache(created, parentId).catch((e) =>
      console.warn('[mail-organizer] addFolderToCache failed (non-fatal)', e),
    )

    signal.throwIfAborted()
    try {
      const res = await api.moveMessage(item.emailId, created.Id, signal)
      return {
        status: recoveredFromExisting ? 'moved' : 'folder_created',
        newMessageId: res.Id,
        destinationFolderId: created.Id,
        destinationFolderPath: fullPath,
      }
    } catch (e) {
      if (isAlreadyMovedError(e)) {
        // Folder was already created (or recovered from 409) — that
        // bookkeeping stands. We just couldn't deposit this particular
        // message because its source Id is no longer valid. Tag the
        // outcome as skipped so the row doesn't appear as a hard error.
        return {
          status: 'skipped',
          message: ALREADY_MOVED_MESSAGE,
          alreadyGone: true,
          // Keep the folder we created visible to downstream code (so
          // sibling items in the same batch can use it as a target via
          // the spliced-into-tree path lookup).
          destinationFolderId: created.Id,
          destinationFolderPath: fullPath,
        }
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
 *   4. domain + subject feature → compound  (NEW 2026-05-22)
 *      e.g. `telco.example + 甲公司` — generic noun-phrase routing.
 *      Feature = a Chinese 4-12 char or Latin 6-12 char proper
 *      noun-ish token from the subject, excluding stopwords and
 *      Taiwan-legal patterns (those have their own slots above).
 *   5. subject feature alone → subject_keyword  (NEW 2026-05-22)
 *      Lower confidence than court-case because the feature heuristic
 *      is noisier; `bumpRuleHits` errorRate gate demotes bad features
 *      over time.
 *   6. usable domain (non-generic-provider, non-internal) → domain
 *   7. generic-provider sender → sender (exact email address)
 *      gmail.com / yahoo.com etc. as a domain is useless; a specific
 *      address like andy@gmail.com IS a precise routing signal.
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
 * DEMOTE the stale rule, but SKIPS creating a new rule. Rare path —
 * fires for empty / single-char subjects only.
 *
 * `featureKind` on subject_keyword / compound variants distinguishes:
 *   - 'court_case': signal is a Taiwan court case number — structural,
 *     safe at high confidence.
 *   - 'full_subject': signal is the entire normalized subject — used
 *     by the conflict-upgrade path and the internal-domain fallback.
 */
type LearningSignal =
  | { type: 'compound'; signal: string; descr: string; featureKind: 'court_case' | 'full_subject'; demoteOnly?: false }
  | { type: 'case_code'; signal: string; descr: string; demoteOnly?: false }
  | { type: 'subject_keyword'; signal: string; descr: string; featureKind: 'court_case' | 'full_subject'; demoteOnly?: false }
  | { type: 'domain'; signal: string; descr: string; demoteOnly?: boolean }
  | { type: 'sender'; signal: string; descr: string; demoteOnly?: boolean }

/**
 * Existing rules passed alongside the email being learned-from. Used by
 * the domain (P4) and sender (P5) branches to detect whether the email
 * would create a conflict with a stored same-signal-different-target
 * rule — that's the moment when "先廣後窄" upgrades to compound.
 *
 * When the user just routed an email and the existing broad rule was
 * about to fire elsewhere, that's the signal "the broad rule isn't
 * enough anymore". We use THIS email's subject to build a precise
 * compound rule, demoting the broad rule via the override pipeline.
 */
function chooseLearningSignal(
  item: PlanItem,
  internalDomainSet: ReadonlySet<string>,
  existingRules: Rule[],
): LearningSignal | null {
  const subject = item.emailSubject ?? ''
  const courtCases = extractCourtCaseNumbers(subject)
  const cc = courtCases[0]
  const codes = extractCaseCodes(subject)
  const codeAlpha = codes[0]
  // Normalized full subject (reply/forward prefixes stripped, lower-
  // cased, whitespace collapsed). Returns '' for empty / single-char
  // subjects below MIN_SUBJECT_SIGNAL_LEN.
  const subjectSignal = extractSubjectSignal(subject)

  const fromAddr = item.emailFrom ?? ''
  const at = fromAddr.lastIndexOf('@')
  const domain =
    at >= 0 ? fromAddr.slice(at + 1).toLowerCase().trim() : ''
  // "Usable" excludes internal domains (configurable via
  // settings.internalDomains) — colleagues forwarding multi-case mail
  // should not produce per-folder domain rules. Empty internalDomainSet
  // means solo mode, every external domain is usable.
  const usableDomain = domain && !internalDomainSet.has(domain) ? domain : ''

  // ---- Priority 1: compound with court case (structural ID fast-path)
  // Court case numbers are unique per case → no bloat risk. Build
  // compound + domain immediately for strongest routing. Confidence 0.9.
  if (usableDomain && cc) {
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
  // Court case numbers are globally unique; safe to route on subject
  // alone even without a domain anchor. Confidence 0.85.
  if (cc) {
    return {
      type: 'subject_keyword',
      signal: cc,
      descr: `主旨含 ${cc}`,
      featureKind: 'court_case',
    }
  }
  // ---- Priority 3: Latin case code (structural ID fast-path)
  // Firm-internal case codes (25A0067A etc.) — also unique per case.
  // Confidence 0.9.
  if (codeAlpha) {
    return { type: 'case_code', signal: codeAlpha, descr: `案件代號 ${codeAlpha}` }
  }
  // ---- Priority 4: usable domain (non-generic-provider)
  //
  // "先廣後窄" — start with a plain-domain rule. It covers all mail
  // from this client in one shot. Only when a CONFLICT exists (same
  // domain already routes to a DIFFERENT folder) do we upgrade to
  // compound (domain + 整段主旨), encoding what made the user pick THIS
  // folder over the existing one.
  if (usableDomain && !isGenericProviderDomain(usableDomain)) {
    // Compare via normalizeSignal so a stored rule signal of `@company-b.example`
    // (legitimate when entered through certain manual flows / import
    // sources) matches `company-b.example` from extractDomain. Direct
    // `.toLowerCase()` compare was missing the `@`-prefix case and the
    // trim-whitespace case, leaving conflicting domain rules undetected.
    const conflictExists = existingRules.some(
      (r) =>
        r.type === 'domain' &&
        normalizeSignal('domain', r.signal) === usableDomain &&
        r.enabled &&
        !r.orphaned &&
        r.targetFolderPath !== item.targetFolderPath,
    )
    if (conflictExists) {
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
      // normalizeSignal mirrors what storage / index lookup uses so a
      // sender stored as `User@Gmail.com` matches `user@gmail.com` here.
      const conflictExists = existingRules.some(
        (r) =>
          r.type === 'sender' &&
          normalizeSignal('sender', r.signal) === normalized &&
          r.enabled &&
          !r.orphaned &&
          r.targetFolderPath !== item.targetFolderPath,
      )
      if (conflictExists) {
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
  // colleagues, anonymous From, etc.). Domain rule isn't appropriate
  // here — internal colleagues mail about many cases.
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

function ruleNormSignal(r: Pick<Rule, 'type' | 'signal'>): string {
  return normalizeSignal(r.type, r.signal)
}

async function generateAiConfirmedRules(
  plan: PlanItem[],
  results: ExecuteItemResult[],
): Promise<number> {
  // Snapshot of currently-enabled rules — only used for dedup checks.
  // The atomic addRulesFilteringTombstones call below ensures we don't
  // clobber concurrent writers and respects user tombstones.
  const [existing, settings] = await Promise.all([listRules(), getSettings()])
  const internalDomainSet = new Set(
    settings.internalDomains.map((d) => d.toLowerCase()),
  )
  const newRules: Rule[] = []

  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]!
    const result = results[i]!
    if (result.status !== 'moved' && result.status !== 'folder_created') continue
    // Accept both:
    //   - source === 'ai': AI was confident, user accepted.
    //   - source === 'unresolved': either (a) AI was below threshold and
    //     the user filled in / confirmed the suggestion, or (b) AI failed
    //     entirely and the user picked a folder from scratch. Both are
    //     human-validated intent worth learning from. wasUserOverride
    //     below still separates "user changed AI's guess" → ai_overridden
    //     path from "user agreed with AI / had no AI to override" → here.
    if (item.source !== 'ai' && item.source !== 'unresolved') continue
    if (item.action !== 'move' && item.action !== 'new_folder') continue
    // User overrode the AI's verdict — handled by generateAiOverrideRules.
    // Skip here so we don't mis-attribute the user's choice to AI confidence.
    if (wasUserOverride(item)) continue

    const sig = chooseLearningSignal(item, internalDomainSet, existing)
    if (!sig) continue

    const targetPath =
      item.action === 'move'
        ? item.targetFolderPath
        : item.suggestedParentPath && item.suggestedFolderName
        ? joinFolderPath(item.suggestedParentPath, item.suggestedFolderName.trim())
        : undefined
    if (!targetPath) continue

    // Dedup against ALL rules (including disabled). Previously we only checked
    // enabled rules, which let AI-confirmed creation keep regenerating
    // identical disabled rules — they'd pile up forever after a user disables.
    const sigNorm = normalizeSignal(sig.type, sig.signal)
    const alreadyCovered =
      existing.some(
        (r) =>
          r.type === sig.type &&
          ruleNormSignal(r) === sigNorm &&
          r.targetFolderPath === targetPath,
      ) ||
      newRules.some(
        (r) =>
          r.type === sig.type &&
          ruleNormSignal(r) === sigNorm &&
          r.targetFolderPath === targetPath,
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
    // can't legitimately route to multiple folders).
    const isConflictProneSignal =
      sig.type === 'domain' ||
      sig.type === 'sender' ||
      (sig.type === 'subject_keyword' && sig.featureKind === 'full_subject') ||
      (sig.type === 'compound' && sig.featureKind === 'full_subject')
    if (isConflictProneSignal) {
      const hasConflict = existing.some(
        (r) =>
          r.type === sig.type &&
          r.enabled &&
          ruleNormSignal(r) === sigNorm &&
          r.targetFolderPath !== targetPath,
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
    const confidence = userResolved
      ? Math.min(0.95, baseConfidence + 0.05)
      : baseConfidence

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
  }

  // Filter against tombstones + write atomically. Closes the race where
  // a concurrent deleteRule could land between filter and write, letting
  // us auto-resurrect a rule the user just deleted.
  const { added } = await addRulesFilteringTombstones(newRules)
  return added.length
}

/**
 * Mirror of generateAiConfirmedRules for the "user disagreed" case.
 *
 * When the AI suggested move/new_folder → A but the user changed it to B and
 * the move to B executed successfully, the user has just taught us something
 * concrete about emails from that sender domain. We:
 *   1. Add a new ai_overridden rule (confidence 0.8 — higher than 0.7 used
 *      by ai_confirmed because user-validated > AI-validated).
 *   2. Disable any *non-user_manual* existing rule that points the same
 *      (type, signal) at a different target. This stops the AI from getting
 *      tugged back toward the wrong answer on the next batch.
 *
 * User_manual rules are sacred — never auto-disable them; the user knew what
 * they wanted when they hand-built that rule.
 */
export async function generateAiOverrideRules(
  plan: PlanItem[],
  results: ExecuteItemResult[],
): Promise<{ added: number; disabled: number }> {
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
  }
  const overrides: Override[] = []
  // Existing rules needed for chooseLearningSignal's conflict-prevention
  // check (priority 6 / 7) — if learning a plain-domain rule would
  // collide with an existing different-target rule, force compound.
  const [existing, settings] = await Promise.all([listRules(), getSettings()])
  const internalDomainSet = new Set(
    settings.internalDomains.map((d) => d.toLowerCase()),
  )

  for (let i = 0; i < plan.length; i++) {
    const item = plan[i]!
    const result = results[i]!
    if (result.status !== 'moved' && result.status !== 'folder_created') continue
    if (!wasUserOverride(item)) continue
    // We can only codify move-style overrides as rules. delete / skip are
    // valid user decisions but our Rule type has no action field — a rule
    // always means "move to target".
    if (item.action !== 'move' && item.action !== 'new_folder') continue
    const sig = chooseLearningSignal(item, internalDomainSet, existing)
    if (!sig) continue
    const target = finalTargetPath(item)
    if (!target) continue
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
    })
  }
  if (overrides.length === 0) return { added: 0, disabled: 0 }

  // Track the mutations so we can record audit events AFTER the mutex
  // releases. Doing it inside mutateRules would mean nested storage writes
  // before the lock yielded — annoying but not broken; we keep it outside
  // for cleaner ordering.
  const auditEvents: RuleEvent[] = []
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
    // and our write, letting us auto-resurrect the deleted rule.
    const tombKeys = new Set(
      (await getRuleTombstones()).map(
        (t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`,
      ),
    )
    const allowedOverrides = overrides.filter(
      (o) => !tombKeys.has(`${o.type}::${o.signalNorm}::${o.targetPath}`),
    )
    if (allowedOverrides.length === 0) {
      return { next: rules, result: { added: 0, disabled: 0 } }
    }
    let added = 0
    let disabled = 0
    const next = [...rules]
    const now = Date.now()

    for (const o of allowedOverrides) {
      const sameTypeSignalMatches = (r: Rule): boolean => {
        if (r.type !== o.type) return false
        return ruleNormSignal(r) === o.signalNorm
      }
      // "先廣後窄" upgrade demotion: when learning a compound rule built
      // from a conflict-upgrade (e.g. compound: vendor.com + 請款通知),
      // ALSO retire any enabled plain-domain (or plain-sender) rule for
      // the same anchor that points to a different folder. Reason: the
      // user's mental model is "the domain rule wasn't enough anymore",
      // so the broad rule should step aside. Future mail of this domain
      // whose subject doesn't match either compound rule falls back to
      // AI / conversation memory, which is the correct behavior.
      const isStalePlainRuleForCompoundUpgrade = (r: Rule): boolean => {
        if (o.type !== 'compound') return false
        const parsed = decodeCompound(o.signal)
        if (!parsed) return false
        for (const cond of parsed.conditions) {
          if (
            cond.type === 'domain' &&
            r.type === 'domain' &&
            ruleNormSignal(r) === normalizeSignal('domain', cond.value)
          ) {
            return true
          }
          if (
            cond.type === 'sender' &&
            r.type === 'sender' &&
            ruleNormSignal(r) === normalizeSignal('sender', cond.value)
          ) {
            return true
          }
        }
        return false
      }
      // If a rule for (type, signal, target) already exists, the user's
      // preferred mapping is already represented — still demote competitors.
      const alreadyCovered = next.some(
        (r) => sameTypeSignalMatches(r) && r.targetFolderPath === o.targetPath,
      )
      // Demote sibling rules with same (type, signal) but different target —
      // they're the "AI's wrong answer" we want to retire. user_manual is
      // sacred — the user knew what they were doing when they handbuilt it.
      for (let i = 0; i < next.length; i++) {
        const r = next[i]!
        if (!sameTypeSignalMatches(r) && !isStalePlainRuleForCompoundUpgrade(r)) continue
        if (r.targetFolderPath === o.targetPath) continue
        if (r.source === 'user_manual') continue
        if (!r.enabled) continue
        next[i] = { ...r, enabled: false }
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
      if (alreadyCovered) continue
      // Gate: don't auto-create plain-domain rules from a single user
      // override. Same domain often serves multiple cases — one override
      // isn't enough evidence to reroute every email from this domain.
      // Sender rules (specific From address) DO pass through: one
      // human-validated routing for an exact address is strong signal.
      if (o.type === 'domain') continue
      // demoteOnly: chooseLearningSignal detected a conflict but couldn't
      // extract a subject feature even with the lenient extractor. We
      // already demoted the stale rule above; creating a new rule for
      // the same plain signal would just re-create the same conflict
      // pattern next batch. Applies to sender variant (domain already
      // skipped above).
      if (o.demoteOnly) continue
      const created = newRule({
        type: o.type,
        signal: o.signal,
        targetFolderId: o.targetId,
        targetFolderPath: o.targetPath,
        // User-validated → higher confidence than ai_confirmed. Domain
        // rule creation is blocked above so we only get here for
        // compound / case_code / subject_keyword / sender — all of which
        // either uniquely identify a case or pinpoint a specific human.
        confidence: 0.95,
        source: 'ai_overridden',
      })
      next.push(created)
      auditEvents.push({
        kind: 'create',
        ruleId: created.id,
        at: now,
        actor: 'system',
        after: snapshotRule(created),
      })
      added++
    }

    return { next, result: { added, disabled } }
  })
  if (auditEvents.length > 0) {
    await recordRuleEvents(auditEvents).catch((e) =>
      console.warn('[mail-organizer] recordRuleEvents (ai-override) failed', e),
    )
  }
  return outcome
}

// ---- Undo ------------------------------------------------------------------

const UNDO_ALARM_NAME = 'undo-expire'

async function captureUndoSnapshot(state: ExecuteState): Promise<void> {
  const moves = state.results
    .filter(
      (r) =>
        (r.status === 'moved' || r.status === 'folder_created') &&
        r.newMessageId &&
        r.destinationFolderId,
    )
    .map((r) => ({
      newMessageId: r.newMessageId!,
      subject: r.subject,
      destinationFolderId: r.destinationFolderId!,
      destinationFolderPath: r.destinationFolderPath,
    }))

  if (moves.length === 0) return

  const now = Date.now()
  const snap: UndoSnapshot = {
    batchId: String(now),
    createdAt: now,
    expiresAt: now + UNDO_WINDOW_MS,
    moves,
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
}

/**
 * Reverse the most-recent batch's moves by moving each message back to inbox.
 * No-op if snapshot is missing/expired. Clears the snapshot on completion
 * (including partial failures — a follow-up retry would address pre-moved
 * messages a second time, which Outlook rejects with 404).
 */
export async function executeUndo(api: OutlookApi): Promise<UndoResult> {
  const snap = await getUndoSnapshot()
  if (!snap) return { attempted: 0, restored: 0, failed: 0, errors: [] }

  const errors: UndoResult['errors'] = []
  let restored = 0
  await holdKeepAlive()
  try {
    for (const m of snap.moves) {
      try {
        await api.moveMessage(m.newMessageId, 'inbox')
        restored++
      } catch (e) {
        const message =
          e instanceof OutlookError
            ? `[${e.status}] ${e.message}`
            : e instanceof Error
            ? e.message
            : String(e)
        errors.push({ subject: m.subject, message })
      }
    }
  } finally {
    await clearUndoSnapshot()
    try {
      await chrome.alarms.clear(UNDO_ALARM_NAME)
    } catch {
      // ignore
    }
    await releaseKeepAlive()
  }

  return {
    attempted: snap.moves.length,
    restored,
    failed: errors.length,
    errors,
  }
}

export async function dismissUndo(): Promise<void> {
  await clearUndoSnapshot()
  try {
    await chrome.alarms.clear(UNDO_ALARM_NAME)
  } catch {
    // ignore
  }
}
