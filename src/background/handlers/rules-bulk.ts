// Bulk rule-library handlers (2026-11-05): bulkRuleAction, clearDisabledRules,
// resolveConflict, splitConflictToCompound.
//
// Why: the last three used to loop deleteRule / toggleRule / upsertRule — N
// full read-modify-write rewrites of the whole rule array (each one firing
// storage.onChanged, i.e. a sync-push reschedule, and re-serialising the
// library) plus N separate read-append-writes of the history log. Clearing
// 300 dormant rules was 600 full storage rewrites. Each handler here does ONE
// mutateRules and ONE recordRuleEvents.
//
// The per-rule semantics deliberately mirror the single-rule helpers in
// shared/rules.ts so the bulk path can't drift from them:
//   delete → tombstone (type, normalizeSignal, target) written INSIDE the rules
//            lock, exactly like deleteRule — an auto-generator can't slip in
//            between and resurrect the triple; history 'delete' + snapshot.
//   enable → clears autoDisabledAt / autoDisabledReason (toggleRule).
//   disable→ leaves any old auto-disable provenance alone (toggleRule).
//   no-op toggles (already in the target state) write nothing and log nothing.
//   add    → confidence-capped, matching tombstones cleared inside the lock
//            (upsertRule's create path), history 'create'.
// Selection runs INSIDE the lock against the fresh library, so "every disabled
// rule" / "the losers of this conflict" can't go stale between read and write.

import type { PopupResponse } from '@/shared/messages'
import {
  applyConfidenceCap,
  mutateRules,
  newRule,
  normalizeSignal,
  snapshotOf,
  tombstoneDeletedAt,
} from '@/shared/rules'
import { addRuleTombstones, clearMatchingTombstones, recordRuleEvents } from '@/shared/storage'
import type { Rule, RuleEvent } from '@/shared/types'

type Actor = 'user' | 'system'

export type RuleBatch = {
  /** Rule ids to remove (tombstoned). Wins over `toggle` for the same id. */
  delete?: ReadonlySet<string>
  /** Rule id → desired `enabled`. Ids already in that state are no-ops. */
  toggle?: ReadonlyMap<string, boolean>
  /** Brand-new rules to append. */
  add?: readonly Rule[]
}

export type RuleBatchOutcome = {
  deleted: Rule[]
  /** The rules AFTER their toggle. */
  toggled: Rule[]
  /** As written (post confidence cap). */
  added: Rule[]
}

/**
 * Apply a mixed batch of deletes / toggles / adds in ONE locked rewrite and
 * ONE history append. `plan` runs inside the rules lock on the current
 * library and must be synchronous and side-effect free.
 */
export async function applyRuleBatch<T>(
  plan: (rules: readonly Rule[]) => { batch: RuleBatch; result: T },
  opts: { actor: Actor; label: string },
): Promise<{ outcome: RuleBatchOutcome; result: T }> {
  const at = Date.now()
  let outcome: RuleBatchOutcome = { deleted: [], toggled: [], added: [] }
  const result = await mutateRules<T>(async (rules) => {
    const { batch, result } = plan(rules)
    const del = batch.delete ?? new Set<string>()
    const toggle = batch.toggle ?? new Map<string, boolean>()
    const added = (batch.add ?? []).map(applyConfidenceCap)
    const deleted: Rule[] = []
    const toggled: Rule[] = []
    const next: Rule[] = []
    for (const r of rules) {
      if (del.has(r.id)) {
        deleted.push(r)
        continue
      }
      const want = toggle.get(r.id)
      if (want !== undefined && r.enabled !== want) {
        const updated: Rule = {
          ...r,
          enabled: want,
          ...(want ? { autoDisabledAt: undefined, autoDisabledReason: undefined } : {}),
        }
        toggled.push(updated)
        next.push(updated)
        continue
      }
      next.push(r)
    }
    if (deleted.length > 0) {
      await addRuleTombstones(
        deleted.map((r) => ({
          type: r.type,
          signalNorm: normalizeSignal(r.type, r.signal),
          targetFolderPath: r.targetFolderPath,
          deletedAt: tombstoneDeletedAt(r, at), // clock-skew safe (rules.ts)
        })),
      ).catch((e) => console.warn(`[mail-organizer] addRuleTombstones (${opts.label})`, e))
    }
    if (added.length > 0) {
      // After the tombstone write above: if a batch both deletes and re-adds
      // the same triple, the live rule must not be shadowed by its own ghost.
      await clearMatchingTombstones(
        added.map((r) => ({
          type: r.type,
          signalNorm: normalizeSignal(r.type, r.signal),
          targetFolderPath: r.targetFolderPath,
        })),
      ).catch((e) => console.warn(`[mail-organizer] clearMatchingTombstones (${opts.label})`, e))
      next.push(...added)
    }
    outcome = { deleted, toggled, added }
    const changed = deleted.length + toggled.length + added.length > 0
    // Same array ref on a no-op → mutateRules skips the write (and the
    // onChanged → sync-push it would trigger).
    return { next: changed ? next : rules, result }
  })

  const events: RuleEvent[] = [
    ...outcome.toggled.map(
      (r): RuleEvent => ({
        kind: 'toggle',
        ruleId: r.id,
        at,
        actor: opts.actor,
        enabled: r.enabled,
        signal: r.signal,
        type: r.type,
        targetFolderPath: r.targetFolderPath,
      }),
    ),
    ...outcome.deleted.map(
      (r): RuleEvent => ({ kind: 'delete', ruleId: r.id, at, actor: opts.actor, before: snapshotOf(r) }),
    ),
    ...outcome.added.map(
      (r): RuleEvent => ({ kind: 'create', ruleId: r.id, at, actor: opts.actor, after: snapshotOf(r) }),
    ),
  ]
  if (events.length > 0) {
    await recordRuleEvents(events).catch((e) =>
      console.warn(`[mail-organizer] recordRuleEvents (${opts.label}) failed (non-fatal)`, e),
    )
  }
  return { outcome, result }
}

// ---- bulkRuleAction ---------------------------------------------------------

/**
 * `{ op: 'enable' | 'disable' | 'delete'; ids: string[] }` →
 * `{ changed, missing }`. `changed` counts rules actually deleted / flipped;
 * ids already in the requested state are neither. `missing` = ids not in the
 * library at write time (deleted elsewhere, e.g. by a sync pull).
 */
export async function handleBulkRuleAction(msg: {
  op?: unknown
  ids?: unknown
}): Promise<PopupResponse> {
  const op = msg.op
  if (op !== 'enable' && op !== 'disable' && op !== 'delete') {
    return { ok: false, code: 'INVALID_INPUT', message: 'op 必須是 enable / disable / delete' }
  }
  const rawIds = msg.ids
  if (!Array.isArray(rawIds) || !rawIds.every((id): id is string => typeof id === 'string')) {
    return { ok: false, code: 'INVALID_INPUT', message: 'ids 必須是字串陣列' }
  }
  const ids = [...new Set(rawIds)]
  const { outcome, result } = await applyRuleBatch(
    (rules) => {
      const present = new Set(rules.map((r) => r.id))
      const hit = ids.filter((id) => present.has(id))
      const missing = ids.length - hit.length
      const batch: RuleBatch =
        op === 'delete'
          ? { delete: new Set(hit) }
          : { toggle: new Map(hit.map((id) => [id, op === 'enable'] as const)) }
      return { batch, result: { missing } }
    },
    { actor: 'user', label: `bulk ${op}` },
  )
  return {
    ok: true,
    data: { changed: outcome.deleted.length + outcome.toggled.length, missing: result.missing },
  }
}

// ---- clearDisabledRules -----------------------------------------------------

/** Delete every rule with enabled=false (tombstoned, so AI / scan won't re-learn them). */
export async function handleClearDisabledRules(): Promise<PopupResponse> {
  const { outcome } = await applyRuleBatch(
    (rules) => ({
      batch: { delete: new Set(rules.filter((r) => !r.enabled).map((r) => r.id)) },
      result: undefined,
    }),
    { actor: 'user', label: 'clearDisabledRules' },
  )
  return { ok: true, data: { deleted: outcome.deleted.length } }
}

// ---- resolveConflict --------------------------------------------------------

type ConflictStrategy = 'keep_highest' | 'disable_all' | 'keep_one'

/**
 * strategy:
 *   'keep_highest' — keep the most-confident rule, DELETE the rest
 *                    (writes tombstones — losers won't auto-resurrect)
 *   'keep_one'     — keep the rule the USER picked (msg.keepId), DELETE the
 *                    rest (same tombstone semantics). Manual counterpart of
 *                    keep_highest for when the right answer isn't the
 *                    highest-confidence one.
 *   'disable_all'  — disable every rule in this conflict group (reversible;
 *                    for "park it for now" intent)
 *
 * Losers are DELETED, not disabled: disabling left "zombie" rules greyed out
 * in the list, indistinguishable from manual disables, and the tombstone keeps
 * the next AI batch from regenerating the same (type, signal, target).
 */
export async function handleResolveConflict(msg: {
  ruleIds?: unknown
  strategy?: unknown
  keepId?: unknown
}): Promise<PopupResponse> {
  const ruleIds = Array.isArray(msg.ruleIds) ? (msg.ruleIds as string[]) : []
  const strategy = msg.strategy as ConflictStrategy
  if (ruleIds.length === 0 || !['keep_highest', 'disable_all', 'keep_one'].includes(strategy)) {
    return { ok: false, code: 'INVALID_INPUT', message: '參數錯誤' }
  }
  const requested = typeof msg.keepId === 'string' ? msg.keepId : ''
  const wanted = new Set(ruleIds)

  const { outcome, result } = await applyRuleBatch<
    { error: PopupResponse } | { keepId: string | null }
  >(
    (rules) => {
      const group = rules.filter((r) => wanted.has(r.id))
      if (group.length === 0) {
        return {
          batch: {},
          result: { error: { ok: false, code: 'NOT_FOUND', message: '找不到指定規則' } },
        }
      }
      if (strategy === 'disable_all') {
        return {
          batch: { toggle: new Map(group.map((r) => [r.id, false] as const)) },
          result: { keepId: null },
        }
      }
      let keepId: string
      if (strategy === 'keep_highest') {
        keepId = [...group].sort((a, b) => b.confidence - a.confidence)[0]!.id
      } else {
        if (!group.some((r) => r.id === requested)) {
          return {
            batch: {},
            result: { error: { ok: false, code: 'INVALID_INPUT', message: 'keepId 不在衝突群組內' } },
          }
        }
        keepId = requested
      }
      return {
        batch: { delete: new Set(group.filter((r) => r.id !== keepId).map((r) => r.id)) },
        result: { keepId },
      }
    },
    { actor: 'user', label: 'resolveConflict' },
  )
  if ('error' in result) return result.error
  return {
    ok: true,
    data: { disabled: outcome.toggled.length, deleted: outcome.deleted.length, kept: result.keepId },
  }
}

// ---- splitConflictToCompound ------------------------------------------------

type BaseCondition = { type: 'domain' | 'subject_keyword' | 'sender'; value: string }

/**
 * Each branch can supply 0..N keywords:
 *   - empty (0 kw)  → that rule stays enabled as a "fallback" (matches when
 *                     the simple condition holds but no keyword)
 *   - one or more   → create one compound per keyword (OR semantics for
 *                     multiple keywords targeting the same folder), disable
 *                     the original
 * Validation: at most ONE branch may be empty (otherwise the conflict
 * remains); all non-empty keywords must be unique across all branches.
 */
export async function handleSplitConflictToCompound(msg: { items?: unknown }): Promise<PopupResponse> {
  const items = Array.isArray(msg.items)
    ? (msg.items as Array<{ ruleId: string; keywords: string[] }>)
    : []
  if (items.length < 2) {
    return { ok: false, code: 'INVALID_INPUT', message: '至少要 2 條規則參與拆解' }
  }

  const sanitized = items.map((i) => ({
    ruleId: i.ruleId,
    keywords: (i.keywords ?? []).map((k) => k.trim()).filter((k) => k.length > 0),
  }))

  const fallbackCount = sanitized.filter((s) => s.keywords.length === 0).length
  if (fallbackCount > 1) {
    return { ok: false, code: 'TOO_MANY_FALLBACKS', message: '最多一個分支可留空作為 fallback' }
  }

  const allKeywords = sanitized.flatMap((s) => s.keywords.map((k) => k.toLowerCase()))
  if (new Set(allKeywords).size !== allKeywords.length) {
    return { ok: false, code: 'DUPLICATE_KEYWORDS', message: '同一關鍵字不可在多個分支使用' }
  }

  const { outcome, result } = await applyRuleBatch(
    (rules) => {
      const byId = new Map(rules.map((r) => [r.id, r]))
      const add: Rule[] = []
      const toggle = new Map<string, boolean>()
      let keptAsFallback = 0
      for (const branch of sanitized) {
        const original = byId.get(branch.ruleId)
        if (!original) continue
        if (branch.keywords.length === 0) {
          // Fallback branch — leave the original rule enabled, no compounds.
          keptAsFallback++
          continue
        }
        // Base conditions from the original rule.
        let baseConditions: BaseCondition[] = []
        if (original.type === 'compound') {
          try {
            baseConditions =
              (JSON.parse(original.signal) as { conditions?: BaseCondition[] }).conditions ?? []
          } catch {
            baseConditions = []
          }
        } else if (
          original.type === 'domain' ||
          original.type === 'subject_keyword' ||
          original.type === 'sender'
        ) {
          baseConditions = [{ type: original.type, value: original.signal }]
        } else {
          // case_code can't be folded into compound — skip.
          continue
        }
        for (const kw of branch.keywords) {
          const conditions = [...baseConditions, { type: 'subject_keyword' as const, value: kw }]
          add.push(
            newRule({
              type: 'compound',
              signal: JSON.stringify({ conditions }),
              targetFolderId: original.targetFolderId,
              targetFolderPath: original.targetFolderPath,
              confidence: original.confidence,
              source: 'user_manual',
            }),
          )
        }
        if (original.enabled) toggle.set(original.id, false)
      }
      return { batch: { add, toggle }, result: { keptAsFallback } }
    },
    { actor: 'user', label: 'splitConflictToCompound' },
  )
  return {
    ok: true,
    data: {
      created: outcome.added.length,
      disabled: outcome.toggled.length,
      keptAsFallback: result.keptAsFallback,
    },
  }
}
