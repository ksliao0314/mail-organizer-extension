// Rule engine.
//
// Storage:
//   chrome.storage.local['rules'] = Rule[]  (managed via getRules/setRules in storage.ts)
//
// Matching order — matchEmailWithIndex; the first tier that matches wins
// (2026-11-05). TYPE_PRIORITY below is only the within-library sort key —
// the real match path interleaves tiers by source and specificity:
//   1. case_code                — Latin case code in the subject (any source)
//   2. compound                 — AND of (domain | subject_keyword | sender);
//                                 longest subject condition first
//   3. court-case subject rule  — subject_keyword whose signal IS a court
//                                 case number (any source)
//   4. user_manual domain → subject_keyword → sender — hand-written rules
//                                 outrank every auto-learned broad rule
//   5. auto domain              — any party's address, except the mailbox
//                                 owner's own domains (buildRuleIndex opts)
//   6. body case pass           — the single case number in the fetched body
//                                 vs structural rules (AI path only)
//   7. auto subject_keyword     — substring, longest signal first
//   8. auto sender              — exact From address
//
// Within a tier: effective confidence (empirical accuracy over ALL fires —
// accepted hits + user overrides — once there are enough), then recency.

import { CASE_CLOSED_PURGE_DAYS } from './constants'
import {
  addRuleTombstones,
  clearMatchingTombstones,
  getActiveDays,
  getRuleTombstones,
  getRules,
  recordRuleEvents,
  setRules,
  utcDay,
} from './storage'
import type {
  Email,
  MailFolderNode,
  PlanItem,
  Rule,
  RuleEvent,
  RuleSnapshot,
  RuleSource,
  RuleTombstone,
  RuleType,
} from './types'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from './normalize'

// ---- Concurrency guard -----------------------------------------------------
//
// All rule mutations are read-modify-write on chrome.storage.local['rules'].
// Without serialization, concurrent writers (e.g. execute's bumpRuleHit during
// a long run + user upsertRule from options page) can clobber each other.
//
// The mutex is a Promise chain: each writer awaits the previous one before
// reading/writing, ensuring atomicity at the JS-engine level.

let writeChain: Promise<void> = Promise.resolve()

async function withRulesLock<T>(fn: () => Promise<T>): Promise<T> {
  const release = writeChain
  let resolve: () => void = () => {}
  writeChain = new Promise<void>((r) => {
    resolve = r
  })
  try {
    await release
    return await fn()
  } finally {
    resolve()
  }
}

export type MutateRulesOptions = {
  /**
   * Write `next` exactly as returned — do NOT stamp `editedAt`. ONLY for
   * writers whose rules carry an authoritative editedAt from elsewhere:
   * the sync pull merge (another machine's / the cloud's version) and backup
   * restore (the snapshot's version). Every other writer must leave this off
   * so its semantic edits get stamped (2026-11-05).
   */
  preserveEditedAt?: boolean
}

/**
 * Transactional rule mutation: read → transform → write, all under lock.
 * Use this for any operation that needs to read current state to compute next.
 *
 * editedAt (2026-11-05): this is the ONE place `Rule.editedAt` is stamped, so
 * no call site can forget it. Cross-machine sync merges per rule by editedAt;
 * without a reliable stamp a stale copy from another machine overwrote newer
 * edits (a re-target reverted → mail filed into the old case folder).
 */
export async function mutateRules<T>(
  fn: (rules: Rule[]) => { next: Rule[]; result: T } | Promise<{ next: Rule[]; result: T }>,
  opts: MutateRulesOptions = {},
): Promise<T> {
  return withRulesLock(async () => {
    const current = await getRules()
    const { next, result } = await fn(current)
    // Identity check (audit, perf): every transform follows copy-on-write —
    // it returns the SAME array ref only when truly nothing changed (all 18
    // call sites verified 2026-06). Skipping the write on that path spares a
    // full rule-library serialization per no-op sweep — runSweep fires on
    // EVERY SW wake (module-load listener), so idle wakes used to pay 2 full
    // writes each — and also suppresses the redundant sync push the rules
    // onChanged listener would schedule.
    if (next !== current) {
      await setRules(opts.preserveEditedAt ? next : stampEditedAt(current, next, Date.now()))
    }
    return result
  })
}

// Fields whose change is a real EDIT of the rule's behaviour (2026-11-05).
// Hit counters (matchCount / overrideCount / lastUsedAt) are deliberately
// absent: they change on every batch and sync merges them separately (max).
const SEMANTIC_RULE_FIELDS = [
  'type',
  'signal',
  'targetFolderId',
  'targetFolderPath',
  'confidence',
  'enabled',
  'source',
  'autoDisabledReason',
  'autoDisabledAt',
  'orphaned',
] as const

function sameSemanticValue(
  field: (typeof SEMANTIC_RULE_FIELDS)[number],
  a: unknown,
  b: unknown,
): boolean {
  // undefined / false are the same "not orphaned" state (mirrors diffSnapshots).
  if (field === 'orphaned') return !!a === !!b
  // An explicitly-cleared optional (`autoDisabledAt: undefined`) equals absent.
  return (a ?? undefined) === (b ?? undefined)
}

/**
 * Did `next` change the rule's behaviour relative to `prev`?
 *
 * One exception (2026-11-05): bumpRuleHits' auto-promotion (+confidence on a
 * matchCount threshold crossing) is counter-derived, not an edit. If it
 * stamped editedAt, a machine that merely matched mail while offline would
 * out-date a re-target / disable made elsewhere and revert it on its next
 * push. Its signature — confidence went UP, matchCount went UP, nothing else
 * semantic moved — never occurs on a user edit (the Options form never
 * bumps matchCount). The sync merge keeps the promotion via max-on-tie.
 */
function isSemanticChange(prev: Rule, next: Rule): boolean {
  let changed = false
  let confidenceOnly = true
  for (const f of SEMANTIC_RULE_FIELDS) {
    if (!sameSemanticValue(f, prev[f], next[f])) {
      changed = true
      if (f !== 'confidence') confidenceOnly = false
    }
  }
  if (!changed) return false
  if (
    confidenceOnly &&
    next.confidence > prev.confidence &&
    (next.matchCount ?? 0) > (prev.matchCount ?? 0)
  ) {
    return false
  }
  return true
}

function withEditedAt(r: Rule, editedAt: number | undefined): Rule {
  if (editedAt === undefined) {
    if (!('editedAt' in r)) return r
    const { editedAt: _drop, ...rest } = r
    void _drop
    return rest
  }
  return r.editedAt === editedAt ? r : { ...r, editedAt }
}

/**
 * Stamp editedAt on every rule in `next` whose semantic fields differ from its
 * `prev` version (same id), and on new rules. Everything else carries its
 * stored editedAt forward — a caller passing back a stale rule object (e.g. a
 * UI round-trip) can neither forge nor erase it. Monotonic per rule:
 * max(now, prev + 1), so an edit always supersedes the version it was based
 * on even when that version came from a machine whose clock runs ahead.
 * Exported for tests; production code reaches it only through mutateRules.
 */
export function stampEditedAt(prev: Rule[], next: Rule[], now: number): Rule[] {
  const prevById = new Map<string, Rule>()
  for (const r of prev) prevById.set(r.id, r)
  let out: Rule[] | null = null
  for (let i = 0; i < next.length; i++) {
    const r = next[i]!
    const p = prevById.get(r.id)
    if (p === r) continue // copy-on-write: untouched rule keeps its ref
    let stamped: Rule
    if (!p) stamped = withEditedAt(r, now)
    else if (isSemanticChange(p, r)) stamped = withEditedAt(r, Math.max(now, (p.editedAt ?? 0) + 1))
    else stamped = withEditedAt(r, p.editedAt)
    if (stamped !== r) {
      out ??= next.slice()
      out[i] = stamped
    }
  }
  return out ?? next
}

// ---- CRUD ------------------------------------------------------------------

export async function listRules(): Promise<Rule[]> {
  return getRules()
}

// ---- Audit-log helpers ----------------------------------------------------
//
// Every CRUD entry point below accepts an optional `actor` so callers can
// distinguish user-initiated changes (options page, plan-screen "停用此規則",
// rule import) from system-initiated ones (initial-scan, ai_confirmed,
// ai_overridden, reconcile auto-fix). Defaults to 'system' because the bulk
// of programmatic calls inside the SW are auto-derived; user paths are
// expected to pass `{ actor: 'user' }` explicitly.

type Actor = 'user' | 'system'

export function snapshotOf(r: Rule): RuleSnapshot {
  return {
    type: r.type,
    signal: r.signal,
    targetFolderPath: r.targetFolderPath,
    confidence: r.confidence,
    source: r.source,
    enabled: r.enabled,
    orphaned: r.orphaned,
  }
}

export function diffSnapshots(before: RuleSnapshot, after: RuleSnapshot): string[] {
  const changed: string[] = []
  if (before.type !== after.type) changed.push('type')
  if (before.signal !== after.signal) changed.push('signal')
  if (before.targetFolderPath !== after.targetFolderPath) changed.push('targetFolderPath')
  if (before.confidence !== after.confidence) changed.push('confidence')
  if (before.source !== after.source) changed.push('source')
  if (before.enabled !== after.enabled) changed.push('enabled')
  // Treat undefined / false as the same "not orphaned" state so we don't
  // emit a noisy event the first time the field appears.
  if (!!before.orphaned !== !!after.orphaned) changed.push('orphaned')
  return changed
}

async function logRuleEvents(events: RuleEvent[]): Promise<void> {
  await recordRuleEvents(events).catch((e) =>
    console.warn('[mail-organizer] recordRuleEvents failed (non-fatal)', e),
  )
}

/**
 * upsertRule(mode 'edit') target no longer exists — deleted elsewhere (another
 * tab, the stale sweep, a sync pull) after the editing page loaded. `code`
 * lets the SW map it without importing this class (2026-11-05).
 */
export class RuleNotFoundError extends Error {
  readonly code = 'NOT_FOUND' as const
  constructor(readonly ruleId: string) {
    super(`rule ${ruleId} not found`)
    this.name = 'RuleNotFoundError'
  }
}

export type UpsertMode = 'create' | 'edit'

/**
 * Overlay the fields the rule form owns onto the STORED rule (2026-11-05).
 * The Options page sends back its whole — possibly hours-old — copy;
 * writing that verbatim reset matchCount / overrideCount / lastUsedAt that
 * batches had bumped since the page loaded. Engine-owned fields (stats,
 * createdAt, source, sync stamps) always come from storage.
 */
function mergeEditableFields(stored: Rule, incoming: Rule): Rule {
  const next: Rule = {
    ...stored,
    type: incoming.type,
    signal: incoming.signal,
    targetFolderId: incoming.targetFolderId,
    targetFolderPath: incoming.targetFolderPath,
    confidence: incoming.confidence,
    enabled: incoming.enabled,
  }
  if (!stored.enabled && incoming.enabled) {
    // Re-enable clears the auto-disable provenance (same as toggleRule),
    // except that a closed-case retirement leaves its reason behind as the
    // user's exemption marker — see applyClosedCaseRetirement.
    next.autoDisabledAt = undefined
    next.autoDisabledReason = closedCaseExemptionOnEnable(stored)
  } else if (stored.enabled && !incoming.enabled && isClosedCaseExempt(stored)) {
    next.autoDisabledReason = undefined
  }
  if (
    stored.orphaned &&
    (stored.targetFolderId !== incoming.targetFolderId ||
      stored.targetFolderPath !== incoming.targetFolderPath)
  ) {
    // The form only offers live folders, so a re-target resolves the orphan.
    next.orphaned = false
  }
  return next
}

/**
 * Clear tombstones for a triple the user is (re)creating — including ones
 * stored in a legacy, pre-canonical form (2026-11-05: normalizeSignal now
 * folds court-case forms / whitespace / '@', so a raw-key clear would miss a
 * tombstone written by an older build and leave the ghost behind).
 */
async function clearTombstonesForRule(r: Rule): Promise<void> {
  const signalNorm = normalizeSignal(r.type, r.signal)
  const triples = [{ type: r.type, signalNorm, targetFolderPath: r.targetFolderPath }]
  const want = `${r.type}::${signalNorm}::${r.targetFolderPath}`
  for (const t of await getRuleTombstones()) {
    if (t.signalNorm !== signalNorm && tombstoneKey(t) === want) {
      triples.push({ type: t.type, signalNorm: t.signalNorm, targetFolderPath: t.targetFolderPath })
    }
  }
  await clearMatchingTombstones(triples)
}

/**
 * Create or edit a rule.
 *   - id exists → merge only the form-owned fields onto the stored rule
 *     (stats / createdAt / lastUsedAt preserved), any mode.
 *   - id missing + mode 'edit' → RuleNotFoundError: no insert, no tombstone
 *     clearing. Re-inserting a stale copy resurrected rules deleted after the
 *     page loaded.
 *   - id missing + mode 'create' or no mode (legacy internal callers that
 *     build fresh rules) → insert.
 */
export async function upsertRule(
  rule: Rule,
  opts?: { actor?: Actor; mode?: UpsertMode },
): Promise<Rule> {
  const actor = opts?.actor ?? 'system'
  let prevSnapshot: RuleSnapshot | undefined
  const saved = await mutateRules(async (rules) => {
    const idx = rules.findIndex((r) => r.id === rule.id)
    if (idx < 0 && opts?.mode === 'edit') throw new RuleNotFoundError(rule.id)
    const stored = idx >= 0 ? rules[idx]! : undefined
    // Enforce the per-type confidence cap on every write (see
    // applyConfidenceCap). For user_manual rules this is a no-op.
    const merged = applyConfidenceCap(stored ? mergeEditableFields(stored, rule) : rule)
    // Tombstone clear lives INSIDE the rules lock — otherwise an auto
    // generator could read tombstones (seeing this rule's ghost), filter
    // us out, and write rules between our clear and our write.
    await clearTombstonesForRule(merged).catch((e) =>
      console.warn('[mail-organizer] clearMatchingTombstones (upsert)', e),
    )
    if (stored) prevSnapshot = snapshotOf(stored)
    const next = [...rules]
    if (idx >= 0) next[idx] = merged
    else next.push(merged)
    return { next, result: merged }
  })
  const after = snapshotOf(saved)
  if (prevSnapshot) {
    const changedFields = diffSnapshots(prevSnapshot, after)
    if (changedFields.length > 0) {
      await logRuleEvents([
        {
          kind: 'edit',
          ruleId: saved.id,
          at: Date.now(),
          actor,
          before: prevSnapshot,
          after,
          changedFields,
        },
      ])
    }
  } else {
    await logRuleEvents([
      { kind: 'create', ruleId: saved.id, at: Date.now(), actor, after },
    ])
  }
  return saved
}

export async function deleteRule(id: string, opts?: { actor?: Actor }): Promise<boolean> {
  // Capture for audit log outside the lock — snapshot is just a value.
  let deletedSnapshot: RuleSnapshot | undefined
  const removed = await mutateRules(async (rules) => {
    const target = rules.find((r) => r.id === id)
    if (!target) return { next: rules, result: false }
    deletedSnapshot = snapshotOf(target)
    const next = rules.filter((r) => r.id !== id)
    // Tombstone the triple INSIDE the rules lock so auto-generators
    // (which also read tombstones under withRulesLock) can't race in
    // between rule-deletion and tombstone-write and auto-resurrect this.
    await addRuleTombstones([
      {
        type: target.type,
        signalNorm: normalizeSignal(target.type, target.signal),
        targetFolderPath: target.targetFolderPath,
        deletedAt: tombstoneDeletedAt(target),
      },
    ]).catch((e) =>
      console.warn('[mail-organizer] addRuleTombstones (delete)', e),
    )
    return { next, result: true }
  })
  if (removed && deletedSnapshot) {
    await logRuleEvents([
      {
        kind: 'delete',
        ruleId: id,
        at: Date.now(),
        actor: opts?.actor ?? 'system',
        before: deletedSnapshot,
      },
    ])
  }
  return removed
}

export async function toggleRule(
  id: string,
  enabled: boolean,
  opts?: { actor?: Actor },
): Promise<Rule | undefined> {
  // Return both the rule and a "did the toggle actually change anything"
  // flag so we can suppress event recording on no-ops (toggling enabled→true
  // when it was already true). Without this gate the history fills with
  // noise from idempotent toggle calls.
  //
  // Side-effect on re-enable: clear `autoDisabledAt` + `autoDisabledReason`.
  // Without this, a rule the user manually re-enabled would still wear the
  // "已自動休眠" badge — visual lie. Once user takes manual action, the
  // auto-disabled provenance is no longer current state.
  const outcome = await mutateRules<{ rule: Rule | undefined; changed: boolean }>((rules) => {
    const idx = rules.findIndex((r) => r.id === id)
    if (idx < 0) return { next: rules, result: { rule: undefined, changed: false } }
    const before = rules[idx]!
    if (before.enabled === enabled) {
      return { next: rules, result: { rule: before, changed: false } }
    }
    const updated: Rule = {
      ...before,
      enabled,
      // On enable: clear auto-disable provenance. On disable (user
      // chose to disable): leave any old auto-disable timestamp alone —
      // doesn't make sense to "auto-disable" something the user
      // explicitly disabled.
      //
      // Closed-case exemption (2026-11-05): re-enabling a 'case_closed'
      // rule keeps that reason on the now-enabled rule as the marker
      // applyClosedCaseRetirement honours ("the user wants this rule live
      // even though its folder is archived"); disabling an exempt rule
      // drops the marker so a later un-archive can't re-enable a rule the
      // user turned off.
      ...(enabled
        ? { autoDisabledAt: undefined, autoDisabledReason: closedCaseExemptionOnEnable(before) }
        : isClosedCaseExempt(before)
          ? { autoDisabledReason: undefined }
          : {}),
    }
    const next = [...rules]
    next[idx] = updated
    return { next, result: { rule: updated, changed: true } }
  })
  if (outcome.changed && outcome.rule) {
    await logRuleEvents([
      {
        kind: 'toggle',
        ruleId: id,
        at: Date.now(),
        actor: opts?.actor ?? 'system',
        enabled,
        signal: outcome.rule.signal,
        type: outcome.rule.type,
        targetFolderPath: outcome.rule.targetFolderPath,
      },
    ])
  }
  return outcome.rule
}

export async function bumpRuleHit(id: string): Promise<void> {
  await mutateRules((rules) => {
    const idx = rules.findIndex((r) => r.id === id)
    if (idx < 0) return { next: rules, result: undefined }
    const existing = rules[idx]!
    const next = [...rules]
    next[idx] = {
      ...existing,
      matchCount: existing.matchCount + 1,
      lastUsedAt: new Date().toISOString(),
    }
    return { next, result: undefined }
  })
}

/**
 * Batch variant of bumpRuleHit — used by execute.ts at the end of a batch so
 * a 50-email batch doesn't issue 50 sequential mutateRules calls (which would
 * each serialize the entire rules array through chrome.storage.local).
 *
 * Accepts a Map<ruleId, hitCount> and bumps every matching rule in one
 * critical section. lastUsedAt gets the now-timestamp; we don't try to
 * record per-hit timestamps because storage cost outweighs the value.
 */
// Confidence auto-promotion thresholds. Every 20 confirmed hits (rule fired
// AND user didn't override the action / target before execute) the rule's
// confidence gets +CONFIDENCE_STEP, capped at CONFIDENCE_CAP. user_manual
// rules are excluded — the lawyer set their confidence deliberately and
// auto-bumping would override that intent.
const CONFIDENCE_STEP_HITS = 20
const CONFIDENCE_STEP = 0.05
// Don't promote rules whose empirical error rate (overrideCount / matchCount)
// exceeds this. Without the gate, a 20/20-overridden rule would keep climbing
// — `effectiveConfidence` correctly demotes it at match time, but the
// displayed `confidence` in the rules UI would be a lie.
// 0.2 = 20% — generous enough to absorb noisy edges but tight enough to stop
// auto-promotion of rules the user clearly disagrees with.
const PROMOTION_MAX_ERROR_RATE = 0.2

/**
 * Type-aware confidence ceiling (2026-05-27 redesign).
 *
 * Plain-domain / plain-sender rules are intentional "broad catch-all"
 * signals — they should stay below compound rules so that when a
 * compound rule (domain + 整段主旨) is present, it wins on specificity
 * even if the plain-domain rule has accumulated many hits.
 *
 * Without per-type ceilings, a plain-domain with 200 hits would auto-
 * promote to 1.0 — pushing it above compound's typical 0.85 base and
 * inverting the precedence the "先廣後窄" design depends on.
 *
 * The cap is on AUTO-PROMOTION only. user_manual rules bypass bumpRuleHits
 * entirely (their confidence reflects the lawyer's deliberate intent), so
 * a user_manual plain-domain rule can carry 0.95 if the user said so.
 */
function confidenceCapForType(type: RuleType): number {
  switch (type) {
    case 'domain':
      return 0.7 // broad fallback — must stay below compound's 0.85
    case 'sender':
      return 0.75 // between domain and compound
    case 'subject_keyword':
      return 0.9 // structurally specific (full_subject or court_case)
    case 'compound':
    case 'case_code':
      return 0.95 // structurally unique identifiers
  }
}

/**
 * Apply the type-aware confidence cap on a rule about to be written.
 *
 * Without write-time enforcement, an imported rule library could carry
 * conf 0.95 plain-domain rules (e.g., from a pre-2026-05-27 backup),
 * which match-time tie-breakers would treat as "as authoritative as a
 * compound rule". Cross-type sort still goes through TYPE_PRIORITY first
 * so routing correctness is unaffected — but the in-bucket ordering
 * + UI confidence display were misleading.
 *
 * user_manual is exempt: the user deliberately picked the confidence.
 */
export function applyConfidenceCap(rule: Rule): Rule {
  if (rule.source === 'user_manual') return rule
  const cap = confidenceCapForType(rule.type)
  if (rule.confidence <= cap) return rule
  return { ...rule, confidence: cap }
}

function countThresholdCrossings(prev: number, next: number): number {
  // Number of integer multiples of CONFIDENCE_STEP_HITS in (prev, next].
  return Math.floor(next / CONFIDENCE_STEP_HITS) - Math.floor(prev / CONFIDENCE_STEP_HITS)
}

export async function bumpRuleHits(hits: Map<string, number>): Promise<{
  promoted: Array<{ ruleId: string; from: number; to: number }>
}> {
  if (hits.size === 0) return { promoted: [] }
  const promotionRecords: Array<{
    ruleId: string
    before: RuleSnapshot
    after: RuleSnapshot
  }> = []
  await mutateRules((rules) => {
    const now = new Date().toISOString()
    const next = rules.map((r) => {
      const inc = hits.get(r.id)
      if (!inc) return r
      const prevCount = r.matchCount
      const nextCount = prevCount + inc
      let nextConfidence = r.confidence
      if (r.source !== 'user_manual') {
        const steps = countThresholdCrossings(prevCount, nextCount)
        // Don't bump confidence on rules the user keeps overriding. The
        // empirical error rate gate is checked at promotion time, not at
        // creation — a rule that started clean and went bad still gets
        // its confidence frozen here. Rate over ALL fires (2026-11-05):
        // overrides never bump matchCount, so overrides/matchCount was a
        // rejected-to-accepted ratio, not an error rate.
        const errorRate = ruleErrorRate({ matchCount: nextCount, overrideCount: r.overrideCount })
        if (steps > 0 && errorRate <= PROMOTION_MAX_ERROR_RATE) {
          const cap = confidenceCapForType(r.type)
          const candidate = Math.min(cap, r.confidence + steps * CONFIDENCE_STEP)
          if (candidate > r.confidence) {
            const before = snapshotOf(r)
            const updated = {
              ...r,
              matchCount: nextCount,
              lastUsedAt: now,
              confidence: candidate,
            }
            promotionRecords.push({
              ruleId: r.id,
              before,
              after: snapshotOf(updated),
            })
            nextConfidence = candidate
          }
        }
      }
      return {
        ...r,
        matchCount: nextCount,
        lastUsedAt: now,
        confidence: nextConfidence,
      }
    })
    return { next, result: undefined }
  })
  if (promotionRecords.length > 0) {
    const at = Date.now()
    await logRuleEvents(
      promotionRecords.map((p) => ({
        kind: 'edit' as const,
        ruleId: p.ruleId,
        at,
        actor: 'system' as const,
        before: p.before,
        after: p.after,
        changedFields: ['confidence'],
      })),
    )
  }
  return {
    promoted: promotionRecords.map((p) => ({
      ruleId: p.ruleId,
      from: p.before.confidence,
      to: p.after.confidence,
    })),
  }
}

/**
 * Bump `overrideCount` for rules whose suggestion the user changed
 * before executing. Mirrors `bumpRuleHits` — accumulated map, single
 * atomic write. No threshold-promotion logic (overrides ≠ promotion);
 * the count just feeds `effectiveConfidence` / accuracy displays.
 */
export async function bumpRuleOverrides(overrides: Map<string, number>): Promise<void> {
  if (overrides.size === 0) return
  await mutateRules((rules) => {
    // Override is "rule fired but user disagreed" — still counts as the
    // rule firing, so refresh lastUsedAt alongside overrideCount.
    // Without this, the stale-sweep lastUsedAt branch would treat a
    // frequently-overridden rule as if it had never been used, then
    // disable it before the high-error-rate branch had a chance to
    // observe it. The high-error-rate path is the correct retirement
    // mechanism for these rules.
    const now = new Date().toISOString()
    const next = rules.map((r) => {
      const inc = overrides.get(r.id)
      if (!inc) return r
      return {
        ...r,
        overrideCount: (r.overrideCount ?? 0) + inc,
        lastUsedAt: now,
      }
    })
    return { next, result: undefined }
  })
}

// ---- Stale rule auto-disable (added 2026-05-22) ---------------------------
//
// Subject-feature learning can create many auto-generated rules over
// time. If a rule has never matched in N days, it's almost certainly
// noise; auto-disable it (flip `enabled = false`) so it stops cluttering
// the matching path, but keep the row so the user can re-enable if it
// turns out to be relevant later.
//
// We DO NOT delete here — that would lose audit history and require
// tombstone bookkeeping. enabled=false is reversible; the user can
// turn it back on in the rules UI.
//
// Skipped:
//   - user_manual rules: sacred, never auto-disabled.
//   - already-disabled rules: nothing to do.
//   - orphaned rules: separate concern (folder gone).

// Both cutoffs are in ACTIVE days (UTC days on which a batch was executed),
// not calendar days (2026-11-05). Wall-clock aging hard-deleted every learned
// rule once the user simply stopped classifying for ~100 days, and re-deleted
// restored backups on the next wake because they keep their old lastUsedAt.
// Counting only days the matcher actually ran keeps the original intent —
// "100 days of use without a hit = dead" — while freezing the clock whenever
// the tool sits idle.
const STALE_AGE_DAYS_DEFAULT = 100
// Class-differentiated cutoff for never-matched full_subject rules
// (2026-07-16). 實測: 規則庫 73% 是 full_subject 類、其中 71% 從未命中 —
// 大多數主旨一輩子只出現一次，這類死規則等 100 天太慢（同步預算已被
// 撐爆）。45 天足以涵蓋「主旨會再出現」的正常週期; 結構化案號 / domain /
// sender 維持 100 天不變。只影響 matchCount === 0 的分支 — 命中過的
// full_subject 規則證明過價值，仍走 100 天 last-used 標準。
const FULL_SUBJECT_NEVER_MATCHED_DAYS_DEFAULT = 45

/**
 * Active days evidenced by the rule library itself: every lastUsedAt is the
 * day a batch credited (or overrode) that rule, and AI-learned rules are
 * minted at execute time. Rules sync, so this carries the OTHER machines'
 * activity into this machine's sweep (keeping the no-tombstone hard delete
 * roughly convergent) and backfills history for libraries that predate the
 * recorded ledger. It can only UNDER-count real activity, which errs on the
 * side of keeping rules.
 */
export function deriveActiveDaysFromRules(rules: readonly Rule[]): string[] {
  const out = new Set<string>()
  for (const r of rules) {
    const used = utcDay(r.lastUsedAt)
    if (used) out.add(used)
    if (r.source === 'ai_confirmed' || r.source === 'ai_overridden') {
      const born = utcDay(r.createdAt)
      if (born) out.add(born)
    }
  }
  return [...out]
}

/**
 * Number of active days strictly AFTER the UTC day of `iso`, counting at most
 * `enough` (early exit). `daysDesc` must be unique and sorted descending.
 * Returns null when `iso` is unparseable — the caller must treat that as
 * "not stale", mirroring the old Number.isFinite guard.
 */
export function countActiveDaysAfter(
  daysDesc: readonly string[],
  iso: string | undefined,
  enough: number,
): number | null {
  const day = utcDay(iso)
  if (!day) return null
  let n = 0
  for (const d of daysDesc) {
    if (d <= day) break
    n++
    if (n >= enough) break
  }
  return n
}

// High-error-rate auto-disable thresholds. A rule that fired this many
// times, with at least this share of the fires overridden by the user, is
// clearly mis-routing; let it rest. User can re-enable in the rules UI if
// the recent overrides were exceptional. Both are measured over ALL fires
// (matchCount + overrideCount — see ruleFires), 2026-11-05.
const HIGH_ERROR_RATE_MIN_SAMPLES = 20
const HIGH_ERROR_RATE_THRESHOLD = 0.5

/**
 * Auto-disable reasons the learners may undo when they re-derive the same
 * (type, normalized signal, target) — fresh evidence the routing is right
 * (2026-11-05). Before, every learner treated ANY disabled copy as "already
 * covered", so the revive branch in addRulesFilteringTombstones was
 * unreachable: a one-strike ('override') or superseded rule stayed dead
 * forever AND blocked its own triple.
 *   - never revived: no reason (user-disabled — user intent), 'case_closed'
 *     (folder LOCATION decides, see applyClosedCaseRetirement),
 *     'legacy_token' (the sweep would re-disable it immediately).
 *   - 'stale' only exists on pre-hard-delete libraries; the same "the signal
 *     came back" logic applies.
 */
const REVIVABLE_DISABLE_REASONS: ReadonlySet<string> = new Set([
  'override',
  'superseded',
  'high-error-rate',
  'stale',
])

/** Disabled by the engine for a reason learning may reverse (see above). */
export function isRevivableAutoDisabled(
  r: Pick<Rule, 'enabled' | 'autoDisabledAt' | 'autoDisabledReason'>,
): boolean {
  if (r.enabled) return false
  if (r.autoDisabledReason) return REVIVABLE_DISABLE_REASONS.has(r.autoDisabledReason)
  // Reason-less auto-disable from older builds: the pre-2026-11-05 revive
  // branch keyed on autoDisabledAt alone — keep honouring it.
  return !!r.autoDisabledAt
}

/**
 * A rule switched back on by learning. Clears the auto-disable provenance
 * (as toggleRule does). A 'high-error-rate' rule also restarts its accuracy
 * window (2026-11-05): the stored rate is what disabled it and cannot change
 * while the rule is off, so keeping it would let the very next sweep disable
 * it again — a revive/disable flap that also restarts its dormant clock. The
 * confirming filing counts as its first accepted fire.
 */
export function revivedRule(r: Rule, nowIso: string): Rule {
  const freshWindow = r.autoDisabledReason === 'high-error-rate'
  return {
    ...r,
    enabled: true,
    autoDisabledAt: undefined,
    autoDisabledReason: undefined,
    ...(freshWindow ? { matchCount: 1, overrideCount: 0, lastUsedAt: nowIso } : {}),
  }
}

/**
 * Auto-disabled rules the sweep hard-deletes after DORMANT_PURGE_ACTIVE_DAYS
 * active days in that state (2026-11-05). They used to be kept forever —
 * blocking their triple and eating sync budget. The predicate reads only
 * synced fields (autoDisabledReason + autoDisabledAt) plus the active-day
 * clock, so — like the stale delete — every machine reaches the same verdict
 * and no tombstone is needed. user_manual is never purged.
 */
const DORMANT_PURGE_REASONS: ReadonlySet<string> = new Set([
  'override',
  'superseded',
  'high-error-rate',
])
const DORMANT_PURGE_ACTIVE_DAYS = 60

/**
 * True when a rule's subject condition is an ENTIRE normalized subject
 * (featureKind 'full_subject') rather than a structural court-case number.
 * featureKind isn't persisted on Rule, so classify from the stored signal:
 * subject_keyword whose signal is NOT a case number, or compound whose
 * subject condition is NOT a case number. Mirrors execute.ts's
 * isStructuralCourtCaseRule (the complement, for subject-bearing types).
 */
export function isFullSubjectRule(r: Pick<Rule, 'type' | 'signal'>): boolean {
  if (r.type === 'subject_keyword') return courtCaseSignal(r.signal) == null
  if (r.type === 'compound') {
    const parsed = decodeCompound(r.signal)
    const subj = parsed?.conditions.find((c) => c.type === 'subject_keyword')
    return subj != null && courtCaseSignal(subj.value) == null
  }
  return false
}

/**
 * Hard-delete rules that share the same `(type, normalizeSignal, targetFolderPath)`
 * triplet. Storage hygiene migration — these arise from historical paths
 * that pushed rules without checking the triplet (pre-2026-05-22
 * `generateAiConfirmedRules`, pre-removal `auto-conflict-resolver`, the
 * sync engine's merge step before its dedup tightening). Each triplet
 * keeps exactly one survivor:
 *
 *   1. Prefer enabled over disabled.
 *   2. Prefer source rank: user_manual > ai_overridden > ai_confirmed > auto_scan.
 *      (User intent always wins; AI override stronger than AI confirm;
 *      auto_scan is weakest since it's bulk-generated.)
 *   3. Tiebreak: higher matchCount (more "proven").
 *   4. Final tiebreak: older createdAt (most established).
 *
 * The survivor then absorbs the group's usage history (2026-11-05): max
 * lastUsedAt, max matchCount / overrideCount, min createdAt. Picking it
 * without looking at lastUsedAt used to drop the copy used yesterday and
 * keep a stale twin that the very same sweep then hard-deleted. Max (not
 * sum) because duplicates are mostly copies of ONE rule (sync / import) —
 * summing would double-count shared history.
 *
 * Like stale deletion: no tombstone, no audit. These are storage
 * duplicates, not policy decisions — they shouldn't show up in the
 * audit trail or block future learning.
 */
export type DedupeReport = {
  /** Number of duplicate rules removed (survivors stay; never counted here). */
  removed: number
  /** Number of distinct triplets that had any duplicates at all (≥ 1 removed). */
  groupsAffected: number
}

const SOURCE_RANK: Record<RuleSource, number> = {
  user_manual: 0,
  ai_overridden: 1,
  ai_confirmed: 2,
  auto_scan: 3,
}

function compareDedupeSurvivor(a: Rule, b: Rule): number {
  // F8 (2026-06-03): user_manual is sacred — it must NEVER be
  // dropped in favour of an auto/AI rule, even when the manual
  // rule is DISABLED and the auto one is enabled. Previously
  // `enabled` was the first sort key, so a disabled user_manual
  // rule lost to an enabled auto_scan/ai_confirmed rule sharing
  // the same triple and got hard-deleted (no tombstone, no
  // audit), silently destroying deliberate user intent. Putting
  // user_manual ahead of the enabled-preference restores the
  // invariant. For NON-user_manual rules, enabled-first still
  // holds (keep the rule that's actually routing).
  const aUM = a.source === 'user_manual'
  const bUM = b.source === 'user_manual'
  if (aUM !== bUM) return aUM ? -1 : 1
  if (a.enabled !== b.enabled) return a.enabled ? -1 : 1
  const rA = SOURCE_RANK[a.source] ?? 9
  const rB = SOURCE_RANK[b.source] ?? 9
  if (rA !== rB) return rA - rB
  if (a.matchCount !== b.matchCount) return b.matchCount - a.matchCount
  return a.createdAt.localeCompare(b.createdAt)
}

function isoMs(iso: string | undefined): number {
  const ms = iso ? Date.parse(iso) : NaN
  return Number.isFinite(ms) ? ms : NaN
}

/** Fold a duplicate group's usage history into its survivor. */
function mergeDuplicateStats(survivor: Rule, group: readonly Rule[]): Rule {
  let lastUsedAt = survivor.lastUsedAt
  let createdAt = survivor.createdAt
  let matchCount = survivor.matchCount
  let overrideCount = survivor.overrideCount
  for (const r of group) {
    const used = isoMs(r.lastUsedAt)
    if (Number.isFinite(used) && !(isoMs(lastUsedAt) >= used)) lastUsedAt = r.lastUsedAt
    const born = isoMs(r.createdAt)
    if (Number.isFinite(born) && !(isoMs(createdAt) <= born)) createdAt = r.createdAt
    if (r.matchCount > matchCount) matchCount = r.matchCount
    if ((r.overrideCount ?? 0) > (overrideCount ?? 0)) overrideCount = r.overrideCount
  }
  return {
    ...survivor,
    matchCount,
    createdAt,
    ...(lastUsedAt !== undefined ? { lastUsedAt } : {}),
    ...(overrideCount !== undefined ? { overrideCount } : {}),
  }
}

function dedupeTransform(rules: Rule[]): { next: Rule[]; report: DedupeReport } {
  const buckets = new Map<string, Rule[]>()
  for (const r of rules) {
    const key = ruleTripleKey(r)
    const arr = buckets.get(key)
    if (arr) arr.push(r)
    else buckets.set(key, [r])
  }
  let removed = 0
  let groupsAffected = 0
  // Keyed on object identity, not id: duplicate ids exist in the wild
  // (pre-2026-11-05 import merge), and an id-keyed keep-set kept or dropped
  // both copies together.
  const replacement = new Map<Rule, Rule>()
  const dropped = new Set<Rule>()
  for (const arr of buckets.values()) {
    if (arr.length === 1) continue
    groupsAffected++
    removed += arr.length - 1
    const sorted = [...arr].sort(compareDedupeSurvivor)
    replacement.set(sorted[0]!, mergeDuplicateStats(sorted[0]!, arr))
    for (const r of sorted.slice(1)) dropped.add(r)
  }
  if (removed === 0) return { next: rules, report: { removed: 0, groupsAffected: 0 } }
  const next: Rule[] = []
  for (const r of rules) {
    if (dropped.has(r)) continue
    next.push(replacement.get(r) ?? r)
  }
  return { next, report: { removed, groupsAffected } }
}

export async function dedupeRulesByKey(): Promise<DedupeReport> {
  return mutateRules((rules) => {
    const { next, report } = dedupeTransform(rules)
    return { next, result: report }
  })
}

export type StaleAutoDisableReport = {
  /** Rules whose `enabled` was flipped to false in this sweep (legacy_token
   *  and high-error-rate buckets). Stay in storage; surface in the rules
   *  UI as "已自動休眠" so user can re-enable if needed. */
  disabled: Rule[]
  /** Counts by reason. `stale` counts rules that were hard-deleted (not
   *  in `disabled` — kept here for telemetry). `dormantPurged` counts
   *  auto-disabled rules hard-deleted after DORMANT_PURGE_ACTIVE_DAYS. */
  byReason: { stale: number; highErrorRate: number; legacyToken: number; dormantPurged: number }
}

/**
 * Boundary timestamp — anything before this came from the deprecated
 * extractSubjectFeature / extractSubjectFeatureLenient / token-pass
 * design (4+ Chinese / 6+ Latin proper-noun fragments). After this
 * date, subject_keyword rules use the new extractSubjectSignal output
 * (entire normalized subject) and stay valid.
 */
const REDESIGN_CUTOFF_ISO = '2026-05-27T00:00:00.000Z'

/**
 * Detect a legacy subject_keyword rule from any pre-2026-05-27 path.
 *
 * Three deprecated paths produced over-broad subject_keyword rules:
 *   1. initial-scan token cross-folder pass (source = 'auto_scan',
 *      signal = 3–8 char pure-Chinese fragment)
 *   2. chooseLearningSignal old P4/P5 with `extractSubjectFeature`
 *      (source = 'ai_confirmed', signal = 4+ Chinese / 6+ Latin token)
 *   3. chooseLearningSignal old conflict-upgrade fallback with
 *      `extractSubjectFeatureLenient` (source = 'ai_confirmed' /
 *      'ai_overridden', signal = 2+ Chinese / 3+ Latin token)
 *
 * Detection strategy:
 *   - Exclude structural IDs (court_case / case_code signals) — those
 *     are valid regardless of vintage.
 *   - Exclude user_manual — sacred, user knew what they wanted.
 *   - For source = auto_scan: ALWAYS legacy (only path was the old
 *     token pass; new initial-scan never produces subject_keyword).
 *   - For source = ai_confirmed / ai_overridden: legacy only if
 *     createdAt is before the redesign cutoff. New full_subject rules
 *     stay valid.
 */
function isLegacyTokenSubjectRule(r: Rule): boolean {
  if (r.type !== 'subject_keyword') return false
  if (extractCourtCaseNumbers(r.signal).length > 0) return false
  if (extractCaseCodes(r.signal).length > 0) return false
  if (r.source === 'user_manual') return false
  if (r.source === 'auto_scan') return true
  // ai_confirmed / ai_overridden — only pre-cutoff are legacy.
  return r.createdAt < REDESIGN_CUTOFF_ISO
}

type StaleSweepOptions = {
  now: number
  staleDays: number
  fullSubjectDays: number
  /** Active days, unique, sorted DESC (see buildActiveDaysDesc). */
  daysDesc: readonly string[]
}

function buildActiveDaysDesc(
  recordedDays: readonly string[],
  rules: readonly Rule[],
): string[] {
  return [...new Set([...recordedDays, ...deriveActiveDaysFromRules(rules)])].sort().reverse()
}

function staleTransform(
  rules: Rule[],
  o: StaleSweepOptions,
): { next: Rule[]; report: StaleAutoDisableReport; events: RuleEvent[] } {
  const disabled: Rule[] = []
  const byReason = { stale: 0, highErrorRate: 0, legacyToken: 0, dormantPurged: 0 }
  const events: RuleEvent[] = []
  const next: Rule[] = []
  for (const r of rules) {
    // Untouched paths — keep as-is.
    if (r.source === 'user_manual' || r.orphaned) {
      next.push(r)
      continue
    }
    if (!r.enabled) {
      // Dormant purge (2026-11-05) — see DORMANT_PURGE_REASONS. Hard
      // delete, no tombstone / audit, same as the stale branch.
      if (r.autoDisabledReason && DORMANT_PURGE_REASONS.has(r.autoDisabledReason)) {
        const dormant = countActiveDaysAfter(o.daysDesc, r.autoDisabledAt, DORMANT_PURGE_ACTIVE_DAYS)
        if (dormant !== null && dormant >= DORMANT_PURGE_ACTIVE_DAYS) {
          byReason.dormantPurged++
          continue
        }
      }
      next.push(r)
      continue
    }

    // Fires = accepted hits + overrides (2026-11-05). Overrides never bump
    // matchCount, so the old matchCount-only gates treated a 100%-wrong rule
    // as "never matched" (and, its lastUsedAt refreshed by every override,
    // never stale either) — it was never demoted, disabled or swept.
    const fires = ruleFires(r)
    let reason: 'stale' | 'high-error-rate' | 'legacy_token' | null = null
    if (isLegacyTokenSubjectRule(r)) {
      reason = 'legacy_token'
    } else if (fires === 0) {
      // full_subject 類的死規則走 45 個使用日短週期 (見
      // FULL_SUBJECT_NEVER_MATCHED_DAYS_DEFAULT 的說明)。謂詞依賴同步欄位
      // (type / signal / matchCount / overrideCount / createdAt) + 使用日，
      // 使用日大部分由同步規則本身推導 → 各機器 sweep 判定大致相同、維持無
      // tombstone 硬刪的收斂性。
      const cutoff = isFullSubjectRule(r) ? o.fullSubjectDays : o.staleDays
      const idle = countActiveDaysAfter(o.daysDesc, r.createdAt, cutoff)
      if (idle !== null && idle >= cutoff) {
        reason = 'stale'
      }
    } else if (
      fires >= HIGH_ERROR_RATE_MIN_SAMPLES &&
      ruleErrorRate(r) >= HIGH_ERROR_RATE_THRESHOLD
    ) {
      reason = 'high-error-rate'
    }
    // Last-used stale (added 2026-05-27): rule fired in the past but
    // has not been used for `staleDays` ACTIVE days. Catches "topic ran
    // for a few weeks then moved on" patterns.
    if (!reason && fires > 0) {
      const idle = countActiveDaysAfter(o.daysDesc, r.lastUsedAt ?? r.createdAt, o.staleDays)
      if (idle !== null && idle >= o.staleDays) {
        reason = 'stale'
      }
    }
    if (!reason) {
      next.push(r)
      continue
    }

    // Stale → hard delete. Don't add to `next`, don't write tombstone
    // (so future learning can re-discover the same signal/target if
    // the inbox starts seeing it again), don't push audit event.
    // Only telemetry count survives, and even that lives in memory
    // and is dropped after the sweep returns.
    if (reason === 'stale') {
      byReason.stale++
      continue
    }

    // legacy_token / high-error-rate → soft disable + audit trail.
    const updated: Rule = {
      ...r,
      enabled: false,
      autoDisabledAt: new Date(o.now).toISOString(),
      autoDisabledReason: reason,
    }
    next.push(updated)
    disabled.push(updated)
    if (reason === 'high-error-rate') byReason.highErrorRate++
    else byReason.legacyToken++
    events.push({
      kind: 'toggle',
      ruleId: r.id,
      at: o.now,
      actor: 'system',
      enabled: false,
      type: r.type,
      signal: r.signal,
      targetFolderPath: r.targetFolderPath,
    })
  }
  // Nothing disabled AND nothing deleted → `next` holds exactly the same
  // refs in the same order. Return the ORIGINAL array so mutateRules'
  // identity check skips the full-library rewrite.
  const changed = disabled.length > 0 || byReason.stale > 0 || byReason.dormantPurged > 0
  return { next: changed ? next : rules, report: { disabled, byReason }, events }
}

/**
 * Auto-retire rules — different fates per bucket:
 *
 * 1. **Stale** (HARD DELETE, no trace):
 *    - never fired + `staleDays` ACTIVE days since createdAt, OR
 *    - fired but `staleDays` ACTIVE days since lastUsedAt
 *    (active day = a UTC day on which a batch was executed; see
 *    deriveActiveDaysFromRules / recordActiveDay — idle periods don't count)
 *    The signal isn't reaching the inbox anymore. Per user request:
 *      "100 天沒用的規則就完全刪除、不留紀錄、不影響將來學習"
 *    Removed from rules array entirely. NO tombstone (AI can re-learn
 *    the same signal/target later if mail flow changes). NO audit
 *    event (Rule events log stays focused on real user / AI actions).
 *    NO error log (this is the design, not an error).
 *
 * 2. **High error rate** (SOFT DISABLE):
 *    - fires (matchCount + overrideCount) ≥ HIGH_ERROR_RATE_MIN_SAMPLES
 *    - overrideCount / fires ≥ HIGH_ERROR_RATE_THRESHOLD
 *    The rule is mis-routing — user keeps fixing its verdicts. Surface
 *    in "自動休眠" so user can decide whether the overrides were
 *    exceptional or the rule is fundamentally wrong.
 *
 * 3. **Legacy token** (SOFT DISABLE):
 *    - Pre-2026-05-27 subject_keyword rule from the deprecated
 *      token/fragment design. Surface in "自動休眠" so user can
 *      reactivate any that are still relevant.
 *
 * 4. **Dormant purge** (HARD DELETE, 2026-11-05): a rule auto-disabled as
 *    'superseded' / 'override' / 'high-error-rate' for
 *    DORMANT_PURGE_ACTIVE_DAYS active days.
 *
 * Skipped (never touched):
 *   - `source === 'user_manual'` (sacred)
 *   - `orphaned === true` (handled by reconcile UI)
 *   - disabled rules other than the dormant-purge reasons above
 */
export async function autoDisableStaleRules(
  options: {
    now?: number
    /** Cutoff in ACTIVE days (default 100). */
    staleDays?: number
    /** Never-matched full_subject cutoff in ACTIVE days (default 45). */
    fullSubjectNeverMatchedDays?: number
    /** Recorded active days (UTC 'YYYY-MM-DD'); defaults to the stored ledger.
     *  Always unioned with deriveActiveDaysFromRules(library). */
    activeDays?: readonly string[]
  } = {},
): Promise<StaleAutoDisableReport> {
  const now = options.now ?? Date.now()
  const recordedDays = options.activeDays ?? (await getActiveDays())
  const { report, events } = await mutateRules((rules) => {
    // Computed ONCE from the pre-sweep library so deletions in this pass
    // can't shift the clock for rules evaluated later in the same pass.
    const out = staleTransform(rules, {
      now,
      staleDays: options.staleDays ?? STALE_AGE_DAYS_DEFAULT,
      fullSubjectDays:
        options.fullSubjectNeverMatchedDays ?? FULL_SUBJECT_NEVER_MATCHED_DAYS_DEFAULT,
      daysDesc: buildActiveDaysDesc(recordedDays, rules),
    })
    return { next: out.next, result: out }
  })
  if (events.length > 0) {
    await logRuleEvents(events)
  }
  return report
}

/** Exact segment-prefix match: path IS the prefix, or lives under it. */
function pathUnderAnyPrefix(path: string, prefixes: readonly string[]): boolean {
  return prefixes.some((p) => p.length > 0 && (path === p || path.startsWith(p + '/')))
}

/**
 * Closed-case exemption marker (2026-11-05): an ENABLED rule that still
 * carries autoDisabledReason 'case_closed'. Written when the user re-enables
 * a retired rule (toggleRule / upsertRule) — "keep this one live even though
 * its folder is archived". Phase 1 then leaves it alone instead of retiring
 * it again on the next preflight, and drops the marker once the folder is
 * back outside every archive prefix. Built from existing synced fields only
 * (enabled + autoDisabledReason), so it propagates without a new Rule field.
 */
function isClosedCaseExempt(r: Pick<Rule, 'enabled' | 'autoDisabledReason'>): boolean {
  return r.enabled && r.autoDisabledReason === 'case_closed'
}

/** autoDisabledReason a disabled rule keeps when it is switched back on. */
function closedCaseExemptionOnEnable(
  before: Pick<Rule, 'autoDisabledReason'>,
): Rule['autoDisabledReason'] {
  return before.autoDisabledReason === 'case_closed' ? 'case_closed' : undefined
}

function closedCaseRetirementTransform(
  rules: Rule[],
  closedCasePrefixes: readonly string[],
  nowIso: string,
): { next: Rule[]; retired: Rule[]; revived: Rule[] } {
  const retired: Rule[] = []
  const revived: Rule[] = []
  let changed = false
  const next = rules.map((r) => {
    if (r.orphaned) return r
    const underArchive =
      closedCasePrefixes.length > 0 && pathUnderAnyPrefix(r.targetFolderPath, closedCasePrefixes)
    if (r.enabled && underArchive) {
      // The user re-enabled this rule while its folder was archived — an
      // explicit, newer statement than the archive location. Respect it.
      if (isClosedCaseExempt(r)) return r
      const updated: Rule = {
        ...r,
        enabled: false,
        autoDisabledAt: nowIso,
        autoDisabledReason: 'case_closed',
      }
      retired.push(updated)
      changed = true
      return updated
    }
    // Revive fires even when the prefix list is emptied — un-designating
    // an archive tree releases every rule it had retired.
    if (!r.enabled && r.autoDisabledReason === 'case_closed' && !underArchive) {
      const updated: Rule = {
        ...r,
        enabled: true,
        autoDisabledAt: undefined,
        autoDisabledReason: undefined,
      }
      revived.push(updated)
      changed = true
      return updated
    }
    // Folder left the archive: the exemption has nothing left to exempt.
    // Dropping it means a later re-archive retires the rule normally.
    if (isClosedCaseExempt(r) && !underArchive) {
      changed = true
      return { ...r, autoDisabledReason: undefined }
    }
    return r
  })
  return { next: changed ? next : rules, retired, revived }
}

function closedCaseToggleEvents(rules: readonly Rule[], at: number): RuleEvent[] {
  return rules.map((r) => ({
    kind: 'toggle',
    ruleId: r.id,
    at,
    actor: 'system',
    enabled: r.enabled,
    type: r.type,
    signal: r.signal,
    targetFolderPath: r.targetFolderPath,
  }))
}

/**
 * Two-phase closed-case retirement, Phase 1 (2026-07-16) — SYMMETRIC:
 *   - retire: an ENABLED rule whose targetFolderPath sits under any
 *     closed-case archive prefix → enabled=false + reason 'case_closed'.
 *     Applies to ALL sources including user_manual — the user moving the
 *     case folder into the archive tree is a NEWER statement of intent
 *     than the rule, and a rule pointing into the archive silently files
 *     new mail where nobody looks. (user_manual is still exempt from the
 *     Phase-2 hard delete below.) Exception (2026-11-05): a rule the user
 *     re-enabled AFTER it was retired (exemption marker, see
 *     isClosedCaseExempt) is an even newer statement and stays live.
 *   - revive: a rule retired with 'case_closed' whose path is NO LONGER
 *     under any archive prefix (folder moved back out — reconcile has
 *     just refreshed stored paths by folderId) → re-enable. This is the
 *     case-reopen path; it does NOT rely on new mail arriving — the
 *     folder's LOCATION is the single source of truth, so learning never
 *     revives 'case_closed' (REVIVABLE_DISABLE_REASONS).
 *
 * Call sites: classify preflight (right after reconcileRulesAgainstTree,
 * fresh paths, before matching — closes the "rule still routes into the
 * archive" window) and runSweep (best-effort catch-up with stored paths;
 * Phase 1 also propagates cross-machine via rule sync, so a machine that
 * never runs preflight still converges).
 */
export async function applyClosedCaseRetirement(
  closedCasePrefixes: readonly string[],
  options: { now?: number } = {},
): Promise<{ retired: number; revived: number; rules: Rule[] }> {
  const nowMs = options.now ?? Date.now()
  // Returns the post-mutation array so the classify-preflight caller can
  // refresh its in-memory snapshot — it captured `rules` from the reconcile
  // mutateRules BEFORE this ran, and building the match index from that
  // stale capture would let a just-retired rule still route this round.
  const out = await mutateRules((rules) => {
    const t = closedCaseRetirementTransform(rules, closedCasePrefixes, new Date(nowMs).toISOString())
    return { next: t.next, result: t }
  })
  const auditEvents = closedCaseToggleEvents([...out.retired, ...out.revived], nowMs)
  if (auditEvents.length > 0) {
    await logRuleEvents(auditEvents)
  }
  return { retired: out.retired.length, revived: out.revived.length, rules: out.next }
}

function purgeClosedCaseTransform(
  rules: Rule[],
  now: number,
  purgeMs: number,
): { next: Rule[]; purged: number } {
  let purged = 0
  const next = rules.filter((r) => {
    if (r.enabled || r.autoDisabledReason !== 'case_closed') return true
    if (r.source === 'user_manual') return true
    const disabledMs = Date.parse(r.autoDisabledAt ?? '')
    if (!Number.isFinite(disabledMs) || now - disabledMs < purgeMs) return true
    purged++
    return false
  })
  return { next: purged > 0 ? next : rules, purged }
}

/**
 * Two-phase closed-case retirement, Phase 2 (2026-07-16): hard-delete
 * rules that have sat in 'case_closed' for CASE_CLOSED_PURGE_DAYS+.
 * No tombstone — mirrors the stale-delete semantics, so a genuinely
 * reopened case re-learns via first-sight once the twin is gone.
 *
 * SYNC-SAFETY INVARIANT (do not weaken): a no-tombstone hard delete is
 * only convergent because this predicate is a pure function of SYNCED
 * fields (autoDisabledReason + autoDisabledAt) and runs in runSweep on
 * EVERY machine — each machine deletes the same rules itself, so a
 * union-mode pull can resurrect a copy for at most one sweep interval.
 * Event-triggered deletion (preflight etc.) must never hard-delete
 * directly; it goes through the Phase-1 soft-disable that propagates.
 *
 * user_manual rules are exempt — destroying a user-created artifact
 * crosses a line; they stay soft-disabled (visible in the rules UI)
 * until the user deletes them.
 */
export async function purgeExpiredClosedCaseRules(
  options: { now?: number; purgeDays?: number } = {},
): Promise<{ purged: number }> {
  const now = options.now ?? Date.now()
  const purgeMs = (options.purgeDays ?? CASE_CLOSED_PURGE_DAYS) * 86_400_000
  const purged = await mutateRules((rules) => {
    const t = purgeClosedCaseTransform(rules, now, purgeMs)
    return { next: t.next, result: t.purged }
  })
  return { purged }
}

export type RuleSweepReport = {
  dedupe: DedupeReport
  retirement: { retired: number; revived: number }
  closedCasePurged: number
  stale: StaleAutoDisableReport
}

/**
 * The whole daily sweep — dedupe → closed-case retire/revive → closed-case
 * purge → stale / error-rate / dormant retirement — as ONE mutateRules
 * transaction (2026-11-05): one library read, at most one write, so at most
 * one sync push per day instead of up to four. Each step is the same
 * transform its standalone export runs, applied in the same order the
 * sweep always used; a step that throws is skipped (logged) so one bad
 * step can't block the others' cleanup. `closedCasePrefixes` omitted → the
 * two closed-case steps are skipped (see below).
 */
export async function runRuleSweepPass(
  options: {
    closedCasePrefixes?: readonly string[]
    now?: number
    staleDays?: number
    fullSubjectNeverMatchedDays?: number
    activeDays?: readonly string[]
    purgeDays?: number
  } = {},
): Promise<RuleSweepReport> {
  const now = options.now ?? Date.now()
  const nowIso = new Date(now).toISOString()
  const recordedDays = options.activeDays ?? (await getActiveDays())
  const purgeMs = (options.purgeDays ?? CASE_CLOSED_PURGE_DAYS) * 86_400_000
  const step = <T>(name: string, fallback: T, fn: () => T): T => {
    try {
      return fn()
    } catch (e) {
      console.warn(`[mail-organizer] sweep step ${name} failed (non-fatal)`, e)
      return fallback
    }
  }
  const { report, events } = await mutateRules((original) => {
    let rules = original
    const dedupe = step('dedupe', { next: rules, report: { removed: 0, groupsAffected: 0 } }, () =>
      dedupeTransform(rules),
    )
    rules = dedupe.next
    // Closed-case steps only run with a KNOWN prefix list: [] means "the user
    // un-designated every archive" and revives every retired rule, so an
    // unknown list (settings unreadable) must skip them, not default to [].
    const prefixes = options.closedCasePrefixes
    const noRetirement = { next: rules, retired: [] as Rule[], revived: [] as Rule[] }
    const retirement = prefixes
      ? step('closed-case-retire', noRetirement, () =>
          closedCaseRetirementTransform(rules, prefixes, nowIso),
        )
      : noRetirement
    rules = retirement.next
    const noPurge = { next: rules, purged: 0 }
    const purge = prefixes
      ? step('closed-case-purge', noPurge, () => purgeClosedCaseTransform(rules, now, purgeMs))
      : noPurge
    rules = purge.next
    const emptyStale: StaleAutoDisableReport = {
      disabled: [],
      byReason: { stale: 0, highErrorRate: 0, legacyToken: 0, dormantPurged: 0 },
    }
    const stale = step(
      'stale',
      { next: rules, report: emptyStale, events: [] as RuleEvent[] },
      () =>
        staleTransform(rules, {
          now,
          staleDays: options.staleDays ?? STALE_AGE_DAYS_DEFAULT,
          fullSubjectDays:
            options.fullSubjectNeverMatchedDays ?? FULL_SUBJECT_NEVER_MATCHED_DAYS_DEFAULT,
          // Clock from the PRE-pass library — same "computed once" rule as
          // the standalone sweep, so this pass's own deletions can't shift it.
          daysDesc: buildActiveDaysDesc(recordedDays, original),
        }),
    )
    rules = stale.next
    return {
      next: rules,
      result: {
        report: {
          dedupe: dedupe.report,
          retirement: { retired: retirement.retired.length, revived: retirement.revived.length },
          closedCasePurged: purge.purged,
          stale: stale.report,
        },
        events: [
          ...closedCaseToggleEvents([...retirement.retired, ...retirement.revived], now),
          ...stale.events,
        ],
      },
    }
  })
  if (events.length > 0) {
    await logRuleEvents(events)
  }
  return report
}

/**
 * Minimum fires before we trust the empirical accuracy number over the
 * configured confidence. Below this, the configured confidence (which
 * encodes "intent" — user_manual gets 0.95, auto_scan gets 0.65, etc.)
 * is the better signal because the empirical sample is too small.
 */
const ACCURACY_MIN_SAMPLES = 10

/**
 * Times the rule's verdict reached a plan that was executed: accepted hits
 * (matchCount) + user overrides (overrideCount). The denominator for every
 * accuracy / error-rate number (2026-11-05) — overrides never bump
 * matchCount, so matchCount alone is "accepted fires", not "fires".
 */
export function ruleFires(r: Pick<Rule, 'matchCount' | 'overrideCount'>): number {
  return (r.matchCount ?? 0) + (r.overrideCount ?? 0)
}

/** Share of fires the user overrode; 0 for a rule that never fired. */
export function ruleErrorRate(r: Pick<Rule, 'matchCount' | 'overrideCount'>): number {
  const fires = ruleFires(r)
  return fires > 0 ? (r.overrideCount ?? 0) / fires : 0
}

/**
 * Configured confidence blended with empirical accuracy. When a rule has
 * fired enough times (>= ACCURACY_MIN_SAMPLES fires), trust the empirical
 * track record (accepted / fires); otherwise fall back to the configured
 * value. Used by the matcher's in-tier tiebreak so a rule that's been
 * proven unreliable (high overrideCount) demotes itself naturally — before
 * 2026-11-05 a rule that was ONLY ever overridden kept matchCount 0 and
 * never reached the sample floor, so it was never demoted.
 */
export function effectiveConfidence(r: Pick<Rule, 'confidence' | 'matchCount' | 'overrideCount'>): number {
  const fires = ruleFires(r)
  if (fires >= ACCURACY_MIN_SAMPLES) {
    return Math.max(0, Math.min(1, (r.matchCount ?? 0) / fires))
  }
  return r.confidence
}

/**
 * Append a batch of new rules atomically. Used by user-initiated paths
 * (manual rule entry, import) where tombstones should NOT block re-creation —
 * the user is explicitly saying "I want this rule back". Auto-generators
 * should use `addRulesFilteringTombstones` instead.
 */
export async function addRules(newRules: Rule[], opts?: { actor?: Actor }): Promise<void> {
  if (newRules.length === 0) return
  const cappedRules = newRules.map(applyConfidenceCap)
  await mutateRules((rules) => ({ next: [...rules, ...cappedRules], result: undefined }))
  const actor = opts?.actor ?? 'system'
  const now = Date.now()
  await logRuleEvents(
    cappedRules.map((r) => ({
      kind: 'create' as const,
      ruleId: r.id,
      at: now,
      actor,
      after: snapshotOf(r),
    })),
  )
}

/**
 * Auto-generator variant of addRules: reads tombstones AND writes rules
 * inside the same critical section. Closes the race where a `deleteRule`
 * landing between (filter check) and (rule write) lets us re-add a rule
 * the user just deleted. Returns the survivors so callers can metric the
 * dropped count.
 */
export async function addRulesFilteringTombstones(
  candidates: Rule[],
  opts?: { actor?: Actor },
): Promise<{ added: Rule[]; dropped: number; reEnabled: Rule[] }> {
  if (candidates.length === 0) return { added: [], dropped: 0, reEnabled: [] }
  // Pre-cap before any dedup / filter logic — keeps the rest of the
  // function dealing with already-capped rules.
  candidates = candidates.map(applyConfidenceCap)
  let survivors: Rule[] = []
  let reEnabledRules: Rule[] = []
  await mutateRules(async (rules) => {
    const tombstones = await getRuleTombstones()
    const afterTombstone = filterByTombstones(candidates, tombstones)
    // Dedup against the live rules array INSIDE the mutex — closes the
    // race where two concurrent generateAiConfirmedRules calls both
    // snapshot rules, both pass their own dedup check, and both write
    // identical rules into the array. With the dedup check INSIDE the
    // lock, the second writer sees the first's rule and drops it.
    //
    // Auto-re-enable: when a candidate re-derives the triple of a rule the
    // ENGINE disabled for a revivable reason (isRevivableAutoDisabled —
    // 'override' / 'superseded' / 'high-error-rate' …), the fresh evidence
    // says the routing is right again: flip it back on instead of adding a
    // twin. User-disabled rules (no reason) and 'case_closed' /
    // 'legacy_token' stay disabled. Never revive into a conflict: if another
    // ENABLED rule now routes the same (type, signal) elsewhere, the twin
    // stays dormant — e.g. a case number the user re-targeted A→B must not
    // get its A rule back from old mail still sitting in folder A
    // (2026-11-05).
    //
    // Key shape mirrors filterByTombstones: (type, signalNorm, target).
    const byKey = new Map<string, { enabled: boolean; revivable?: number }>()
    const enabledTargets = new Map<string, Set<string>>()
    rules.forEach((r, idx) => {
      const sig = `${r.type}::${normalizeSignal(r.type, r.signal)}`
      const key = `${sig}::${r.targetFolderPath}`
      const slot = byKey.get(key) ?? { enabled: false }
      if (r.enabled) {
        slot.enabled = true
        const targets = enabledTargets.get(sig) ?? new Set<string>()
        targets.add(r.targetFolderPath)
        enabledTargets.set(sig, targets)
      } else if (slot.revivable === undefined && isRevivableAutoDisabled(r)) {
        slot.revivable = idx
      }
      byKey.set(key, slot)
    })
    const seenInBatch = new Set<string>()
    const reEnabledIndexes = new Map<number, Rule>()
    const nowIso = new Date().toISOString()
    survivors = afterTombstone.filter((r) => {
      const sig = `${r.type}::${normalizeSignal(r.type, r.signal)}`
      const key = `${sig}::${r.targetFolderPath}`
      const existing = byKey.get(key)
      if (existing) {
        if (!existing.enabled && existing.revivable !== undefined) {
          const conflicting = [...(enabledTargets.get(sig) ?? [])].some(
            (t) => t !== r.targetFolderPath,
          )
          if (!conflicting && !reEnabledIndexes.has(existing.revivable)) {
            // revivedRule also clears autoDisabledReason, honouring the
            // types.ts contract ("Cleared along with autoDisabledAt on
            // re-enable") like toggleRule — otherwise the rule would end up
            // enabled with a stale reason that round-trips through
            // export/sync and misreports the rule.
            reEnabledIndexes.set(existing.revivable, revivedRule(rules[existing.revivable]!, nowIso))
          }
        }
        return false
      }
      if (seenInBatch.has(key)) return false
      seenInBatch.add(key)
      return true
    })
    // Materialise reEnabledRules + write the modified rules array.
    let next = rules
    if (reEnabledIndexes.size > 0) {
      next = rules.map((r, i) => reEnabledIndexes.get(i) ?? r)
      reEnabledRules = Array.from(reEnabledIndexes.values())
    }
    if (survivors.length === 0 && reEnabledIndexes.size === 0) {
      return { next: rules, result: undefined }
    }
    return {
      next: survivors.length > 0 ? [...next, ...survivors] : next,
      result: undefined,
    }
  })
  if (survivors.length > 0) {
    const actor = opts?.actor ?? 'system'
    const now = Date.now()
    await logRuleEvents(
      survivors.map((r) => ({
        kind: 'create' as const,
        ruleId: r.id,
        at: now,
        actor,
        after: snapshotOf(r),
      })),
    )
  }
  if (reEnabledRules.length > 0) {
    const now = Date.now()
    await logRuleEvents(
      reEnabledRules.map((r) => ({
        kind: 'toggle' as const,
        ruleId: r.id,
        at: now,
        actor: 'system' as const,
        enabled: true,
        type: r.type,
        signal: r.signal,
        targetFolderPath: r.targetFolderPath,
      })),
    )
  }
  return {
    added: survivors,
    dropped: candidates.length - survivors.length,
    reEnabled: reEnabledRules,
  }
}

/** Canonical (type, normalized signal, target) key — the identity every
 *  dedupe / tombstone / revive comparison uses. */
export function ruleTripleKey(r: Pick<Rule, 'type' | 'signal' | 'targetFolderPath'>): string {
  return `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
}

/**
 * Canonical key of a tombstone. The stored signalNorm is re-normalized
 * (normalizeSignal is idempotent on its own output) so tombstones written
 * by older builds — before court-case forms / whitespace / '@' were folded,
 * or with compound conditions in their raw form — still match the rule they
 * were written for (2026-11-05). Comparing the raw stored string would let
 * those deletions silently stop blocking re-learning.
 */
/**
 * deletedAt for a tombstone of rule `r` (2026-11-05). Sync applies a tombstone
 * only to rule versions with max(createdAt, editedAt) <= deletedAt, so a
 * machine whose clock runs BEHIND the one that last edited the rule would
 * mint a tombstone that never suppresses that version — the delete silently
 * fails to propagate. Stamping strictly after the rule's own timestamps makes
 * the delete win regardless of clock skew.
 */
export function tombstoneDeletedAt(
  r: Pick<Rule, 'createdAt' | 'editedAt'>,
  now: number = Date.now(),
): number {
  const created = Date.parse(r.createdAt)
  return Math.max(
    now,
    typeof r.editedAt === 'number' && Number.isFinite(r.editedAt) ? r.editedAt + 1 : 0,
    Number.isFinite(created) ? created + 1 : 0,
  )
}

export function tombstoneKey(
  t: Pick<RuleTombstone, 'type' | 'signalNorm' | 'targetFolderPath'>,
): string {
  return `${t.type}::${normalizeSignal(t.type, t.signalNorm)}::${t.targetFolderPath}`
}

/**
 * Given a set of candidate rules + the current tombstones, return the
 * subset whose (type, signalNorm, targetFolderPath) is NOT tombstoned.
 * Auto-generators (initial-scan, ai_confirmed, ai_overridden) should run
 * candidates through this before calling addRules to honor prior user
 * deletions.
 */
export function filterByTombstones(candidates: Rule[], tombstones: RuleTombstone[]): Rule[] {
  if (tombstones.length === 0) return candidates
  const blocked = new Set(tombstones.map(tombstoneKey))
  return candidates.filter((r) => !blocked.has(ruleTripleKey(r)))
}

/** Convenience: load tombstones + filter in one call. */
export async function filterByCurrentTombstones(candidates: Rule[]): Promise<Rule[]> {
  const tombstones = await getRuleTombstones()
  return filterByTombstones(candidates, tombstones)
}

export function newRule(input: {
  type: RuleType
  signal: string
  targetFolderId: string
  targetFolderPath: string
  confidence: number
  source: RuleSource
  enabled?: boolean
}): Rule {
  return {
    id: crypto.randomUUID(),
    type: input.type,
    signal: input.signal,
    targetFolderId: input.targetFolderId,
    targetFolderPath: input.targetFolderPath,
    confidence: input.confidence,
    matchCount: 0,
    enabled: input.enabled ?? true,
    createdAt: new Date().toISOString(),
    source: input.source,
  }
}

// ---- Reconcile against folder tree -----------------------------------------

type FolderIndex = {
  byId: Map<string, MailFolderNode>
  byPath: Map<string, MailFolderNode>
}

function indexFolderTree(tree: MailFolderNode[]): FolderIndex {
  const byId = new Map<string, MailFolderNode>()
  const byPath = new Map<string, MailFolderNode>()
  const walk = (n: MailFolderNode) => {
    byId.set(n.id, n)
    byPath.set(n.path, n)
    for (const c of n.children) walk(c)
  }
  for (const root of tree) walk(root)
  return { byId, byPath }
}

export type ReconcileSummary = {
  pathsUpdated: number
  newlyOrphaned: number
  unorphaned: number
}

/**
 * Sync each rule's stored `targetFolderPath` and `orphaned` flag against the
 * live folder tree. Detects three drifts:
 *
 *   1. Folder renamed (Id resolves but path differs) → update path in place.
 *   2. Folder deleted (Id and path both miss) → mark orphaned.
 *   3. Previously orphaned rule now resolves again (rare: user undeletes the
 *      folder, or restored from backup) → clear orphaned.
 *
 * Falls back to path lookup when Id is empty/sentinel — older rules created
 * during AI-confirmed promotion sometimes have empty Id but a valid path.
 *
 * Pure function: callers persist the result themselves so this stays trivial
 * to unit test.
 */
export function reconcileRulesAgainstTree(
  rules: Rule[],
  tree: MailFolderNode[],
): { rules: Rule[]; summary: ReconcileSummary } {
  const idx = indexFolderTree(tree)
  let pathsUpdated = 0
  let newlyOrphaned = 0
  let unorphaned = 0

  const next = rules.map((r) => {
    // Resolve by Id first, then by stale path as fallback.
    const byId = r.targetFolderId ? idx.byId.get(r.targetFolderId) : undefined
    const byPath = idx.byPath.get(r.targetFolderPath)
    const resolved = byId ?? byPath

    if (!resolved) {
      if (r.orphaned) return r // already marked, no change
      newlyOrphaned++
      return { ...r, orphaned: true }
    }

    // Resolved — clear orphaned flag if it was set.
    let updated: Rule = r
    if (r.orphaned) {
      unorphaned++
      updated = { ...updated, orphaned: false }
    }
    // Sync path / id if folder was renamed or rule was created with empty Id.
    if (updated.targetFolderPath !== resolved.path) {
      pathsUpdated++
      updated = { ...updated, targetFolderPath: resolved.path }
    }
    if (!updated.targetFolderId || updated.targetFolderId !== resolved.id) {
      updated = { ...updated, targetFolderId: resolved.id }
    }
    return updated
  })

  return { rules: next, summary: { pathsUpdated, newlyOrphaned, unorphaned } }
}

// ---- Matching --------------------------------------------------------------

// More-specific rules win. Compound has multiple AND-conditions, so it's
// strictly more specific than any single-condition type. Without this,
// "domain company-b.example → A" would match before "domain company-b.example AND subject k → B"
// — the conflict-split feature wouldn't actually do anything useful.
//
// This map is the single source of truth for TYPE ordering. It controls:
//   1. `compareRulePriority` — sorts rules by type before tiebreaking on
//      effective confidence + recency (and the Options rule list mirrors it).
//   2. `matchEmailWithIndex` — the type order INSIDE each source class.
//      Since 2026-11-05 the match path is tiered by source as well
//      (user_manual broad rules before auto ones) and lifts structural
//      court-case subject rules above domain — see the file header for the
//      full tier list.
//
// If you change a value here, double-check both call sites still behave
// as intended. There was a 2026-05-22 incident where matchEmailWithIndex
// was rewritten to iterate sender BEFORE domain/subject_keyword, with
// TYPE_PRIORITY left at the old values — the divergence silently broke
// working subject_keyword rules. Lesson: within a source class this map
// and matchEmailWithIndex must stay in lockstep.
export const TYPE_PRIORITY: Record<RuleType, number> = {
  case_code: 1,
  compound: 2,
  domain: 3,
  subject_keyword: 4,
  sender: 5,
}

/**
 * Precedence helper used by classifyPreflight: when both a rule AND
 * thread memory hit on the same email, decide which wins.
 *
 * - Specific by construction (case_code / compound) → rule wins, since
 *   the signal uniquely identifies one case/topic.
 * - Hand-authored by the user (`user_manual`) → rule wins, since that
 *   represents the user's most explicit, most-recent intent — if they
 *   wrote it AFTER a thread was learned, the rule should override.
 * - Anything else (auto_scan / ai_confirmed / ai_overridden of domain
 *   / subject_keyword / sender) → thread wins, since thread memory
 *   captures actual past human filings, which is sharper than a broad
 *   auto-derived rule.
 *
 * Consistent with the matcher (2026-11-05): matchEmailWithIndex checks every
 * user_manual rule before any auto domain rule, so the outcome handed to
 * this helper is the hand-written rule whenever one matches — previously
 * the matcher returned the auto domain rule instead, this helper said
 * "thread wins", and the user's rule lost twice.
 *
 * Tests in tests/rules.test.ts pin this contract.
 */
export function ruleBeatsThread(
  rule: Pick<Rule, 'type' | 'source'> & { signal?: string },
): boolean {
  if (rule.type === 'case_code' || rule.type === 'compound') return true
  if (rule.source === 'user_manual') return true
  // A subject_keyword rule whose signal IS a court case number is a
  // structural unique identifier — treat it like case_code so it wins over
  // thread memory (優化 2026-07). All auto-learned court-case rules are the
  // subject_keyword type (chooseLearningSignal P2).
  if (rule.type === 'subject_keyword' && rule.signal && courtCaseSignal(rule.signal)) {
    return true
  }
  return false
}

export type MatchOutcome = {
  rule: Rule
  reason: string
}

export function matchEmail(email: Email, rules: Rule[]): MatchOutcome | null {
  // Convenience wrapper for callers (tests, ad-hoc) that don't want to
  // pre-build an index. Production batch paths (service-worker classify
  // loop, initial-scan) should use buildRuleIndex + matchEmailWithIndex
  // for O(1) domain/sender lookup and a single sort across the batch.
  return matchEmailWithIndex(email, buildRuleIndex(rules))
}

// ---- Compiled index (2026-11-05) -------------------------------------------
//
// Everything a rule needs at match time is derived ONCE per index build:
// compound signals are JSON-decoded once (≈73% of a real library is compound,
// and every one used to cost a JSON.parse per email in preflight AND harvest),
// subject needles are lowercased / whitespace-collapsed / court-case-detected
// once, and compounds are bucketed by their sender / domain condition so an
// email only tests the compounds that could possibly match it. Per-email
// features (lowercased addresses, subject forms, court cases) are likewise
// computed once per email instead of once per rule.

/** A subject needle pre-processed at index build. */
type SubjectNeedle = {
  /** Raw signal / condition value — echoed verbatim in the reason. */
  raw: string
  /** Canonical compact case number when the needle IS one, else null. */
  court: string | null
  /** Lowercased, whitespace-collapsed substring needle. */
  needle: string
}

type CompiledCondition =
  | { type: 'domain'; domain: string }
  | { type: 'sender'; raw: string; addr: string }
  | { type: 'subject_keyword'; subj: SubjectNeedle }

type CompiledCompound = {
  rule: Rule
  conds: CompiledCondition[]
  /** Position in the specificity-sorted compound tier; candidates gathered
   *  from the anchor buckets are tried in this order. */
  order: number
}

type CompiledSubjectRule = { rule: Rule; subj: SubjectNeedle }

function compileSubjectNeedle(raw: string): SubjectNeedle {
  return {
    raw,
    court: courtCaseSignal(raw),
    needle: collapseSubjectWhitespace(raw.toLowerCase()),
  }
}

function compileConditions(conds: readonly CompoundCondition[]): CompiledCondition[] {
  return conds.map((c): CompiledCondition => {
    if (c.type === 'domain') return { type: 'domain', domain: normalizeDomain(c.value) }
    if (c.type === 'sender') return { type: 'sender', raw: c.value, addr: c.value.toLowerCase().trim() }
    return { type: 'subject_keyword', subj: compileSubjectNeedle(c.value) }
  })
}

/** Per-email features, computed once per match call. */
type EmailFeatures = {
  subject: string
  subjectUpper: string
  /** Lowercased, whitespace-collapsed subject (the substring haystack). */
  hay: string
  /** Lowercased From address (untrimmed, like the sender-map lookup always was). */
  from: string | undefined
  /** Lowercased From + To + Cc + Bcc, From first. */
  addrs: string[]
  /** Domain of addrs[i] (text after the last '@'); '' when there is none. */
  domains: string[]
  courtCases: () => string[]
}

function emailFeatures(email: Email): EmailFeatures {
  const subject = email.Subject ?? ''
  const addrs = allEmailAddresses(email)
  let courtCases: string[] | undefined
  return {
    subject,
    subjectUpper: subject.toUpperCase(),
    hay: collapseSubjectWhitespace(subject.toLowerCase()),
    from: email.From?.EmailAddress?.Address?.toLowerCase(),
    addrs,
    domains: addrs.map((a) => {
      const at = a.lastIndexOf('@')
      return at < 0 ? '' : a.slice(at + 1)
    }),
    courtCases: () => (courtCases ??= extractCourtCaseNumbers(subject)),
  }
}

/**
 * Pre-built rule index for batch matching — one bucket per tier of the
 * match order in the file header, each pre-sorted so the first hit inside a
 * tier is the best candidate. The Rule[] / Map fields mirror the buckets for
 * callers that inspect them; matching runs off the compiled fields.
 */
export type RuleIndex = {
  caseCodeRules: Rule[]
  /** Specificity order: longest subject condition first (2026-11-05). */
  compoundRules: Rule[]
  /**
   * subject_keyword rules whose signal IS a court case number (優化
   * 2026-07). Kept in their own bucket and matched AFTER compound but
   * BEFORE domain — a structural unique identifier must outrank a broad
   * plain-domain rule (the same-client-multiple-cases hole). This bucket
   * order INTENTIONALLY diverges from TYPE_PRIORITY (subject_keyword=4);
   * compareRulePriority still sorts these as priority 4 for within-library
   * ordering, only the match path elevates them. Any source.
   */
  courtCaseSubjectRules: Rule[]
  /** AUTO domain rules by normalized domain; own-domain rules are left out. */
  domainMap: Map<string, Rule[]>
  /** AUTO non-court subject_keyword rules, longest signal first. */
  subjectKeywordRules: Rule[]
  /** AUTO sender rules by lowercased address. */
  senderMap: Map<string, Rule[]>
  /** user_manual tier (2026-11-05) — checked before every auto broad rule. */
  manualDomainMap: Map<string, Rule[]>
  manualSubjectKeywordRules: Rule[]
  manualSenderMap: Map<string, Rule[]>
  /** The mailbox owner's own domains (lowercased) — see buildRuleIndex. */
  ownDomains: ReadonlySet<string>
  // ---- compiled (internal) ----
  compiledCaseCodes: Array<{ rule: Rule; upper: string }>
  compiledCompounds: {
    bySender: Map<string, CompiledCompound[]>
    byDomain: Map<string, CompiledCompound[]>
    unanchored: CompiledCompound[]
    size: number
  }
  compiledCourtSubjects: CompiledSubjectRule[]
  compiledManualSubjects: CompiledSubjectRule[]
  compiledAutoSubjects: CompiledSubjectRule[]
}

// Subject_keyword has no good index — it's linear substring scan. To
// avoid "通知" (shorter, broad) winning over "112訴204" (longer,
// specific) just because the broad one happens to sort first by
// confidence, the subject buckets sort longest signal first, then
// effective confidence, then recency.
function subjectKeywordSort(a: Rule, b: Rule): number {
  const lenDelta = b.signal.length - a.signal.length
  if (lenDelta !== 0) return lenDelta
  const confDelta = effectiveConfidence(b) - effectiveConfidence(a)
  if (confDelta !== 0) return confDelta
  const aUsed = a.lastUsedAt ?? a.createdAt
  const bUsed = b.lastUsedAt ?? b.createdAt
  return bUsed.localeCompare(aUsed)
}

function longestSubjectCondition(conds: readonly CompiledCondition[]): number {
  let n = 0
  for (const c of conds) {
    if (c.type === 'subject_keyword' && c.subj.needle.length > n) n = c.subj.needle.length
  }
  return n
}

function pushTo<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const bucket = map.get(key)
  if (bucket) bucket.push(value)
  else map.set(key, [value])
}

/**
 * Build the match index.
 *
 * `ownDomains` (2026-11-05): the mailbox owner's own domain(s). The owner is
 * on To/Cc of essentially every inbound email, and a domain rule matches ANY
 * party — so an auto-learned `domain: <own domain>` (minted from the first AI
 * filing of a colleague's mail while internalDomains was still empty, the
 * default) claimed the whole inbox, and its empirical confidence climbed above
 * every client's rule. So:
 *   - AUTO domain rules never fire on an own-domain address (the own-domain
 *     rules themselves are left out of the index; other auto domain rules
 *     can't be matched by such an address anyway).
 *   - user_manual domain rules for an own domain stay honoured, but only on
 *     the From side: the owner's address on To/Cc would make a hand-written
 *     「所內網域 → 所內」 rule swallow all external mail too, which is
 *     clearly not what it means ("mail FROM colleagues").
 *   - compound / sender / subject rules are unaffected.
 */
export function buildRuleIndex(
  rules: Rule[],
  opts?: { ownDomains?: readonly string[] },
): RuleIndex {
  const ownDomains = new Set(
    (opts?.ownDomains ?? []).map((d) => normalizeDomain(d)).filter((d) => d.length > 0),
  )
  // Corrupt rows (unknown type / non-string signal — a bad import, a buggy
  // sync peer) are skipped rather than allowed to throw for every email.
  const enabled = rules
    .filter(
      (r) =>
        r.enabled &&
        !r.orphaned &&
        typeof r.signal === 'string' &&
        Object.prototype.hasOwnProperty.call(TYPE_PRIORITY, r.type),
    )
    .sort(compareRulePriority)
  const index: RuleIndex = {
    caseCodeRules: [],
    compoundRules: [],
    courtCaseSubjectRules: [],
    domainMap: new Map(),
    subjectKeywordRules: [],
    senderMap: new Map(),
    manualDomainMap: new Map(),
    manualSubjectKeywordRules: [],
    manualSenderMap: new Map(),
    ownDomains,
    compiledCaseCodes: [],
    compiledCompounds: { bySender: new Map(), byDomain: new Map(), unanchored: [], size: 0 },
    compiledCourtSubjects: [],
    compiledManualSubjects: [],
    compiledAutoSubjects: [],
  }
  const compounds: Array<{ rule: Rule; conds: CompiledCondition[]; subjectLen: number }> = []
  for (const r of enabled) {
    const manual = r.source === 'user_manual'
    switch (r.type) {
      case 'case_code':
        index.caseCodeRules.push(r)
        break
      case 'compound': {
        // Malformed compound → never matches (decodeCompound validates).
        const decoded = decodeCompound(r.signal)
        if (!decoded || decoded.conditions.length === 0) break
        const conds = compileConditions(decoded.conditions)
        compounds.push({ rule: r, conds, subjectLen: longestSubjectCondition(conds) })
        break
      }
      case 'domain': {
        const d = normalizeDomain(r.signal)
        if (!d) break
        if (manual) pushTo(index.manualDomainMap, d, r)
        else if (!ownDomains.has(d)) pushTo(index.domainMap, d, r)
        break
      }
      case 'subject_keyword':
        // Pure court-case-number subject rules go in their own high-priority
        // bucket (優化 2026-07); everything else is a plain subject rule.
        if (courtCaseSignal(r.signal)) index.courtCaseSubjectRules.push(r)
        else if (manual) index.manualSubjectKeywordRules.push(r)
        else index.subjectKeywordRules.push(r)
        break
      case 'sender': {
        const s = r.signal.toLowerCase().trim()
        if (!s) break
        pushTo(manual ? index.manualSenderMap : index.senderMap, s, r)
        break
      }
    }
  }
  index.courtCaseSubjectRules.sort(subjectKeywordSort)
  index.manualSubjectKeywordRules.sort(subjectKeywordSort)
  index.subjectKeywordRules.sort(subjectKeywordSort)
  // Compound specificity (2026-11-05): the LONGEST subject condition first,
  // then more conditions, then effective confidence / recency. Sorting on
  // confidence / recency alone let 「報告」→甲 claim 「RE: 乙案進度報告」
  // just because it fired more recently than the 乙案 rule.
  compounds.sort(
    (a, b) =>
      b.subjectLen - a.subjectLen ||
      b.conds.length - a.conds.length ||
      compareRulePriority(a.rule, b.rule),
  )
  compounds.forEach((c, order) => {
    const entry: CompiledCompound = { rule: c.rule, conds: c.conds, order }
    index.compoundRules.push(c.rule)
    // Bucket by one REQUIRED condition (sender is the most selective): an
    // email can only satisfy the compound if it carries that anchor.
    const sender = c.conds.find(
      (x): x is Extract<CompiledCondition, { type: 'sender' }> => x.type === 'sender',
    )
    const domain = c.conds.find(
      (x): x is Extract<CompiledCondition, { type: 'domain' }> => x.type === 'domain',
    )
    if (sender) pushTo(index.compiledCompounds.bySender, sender.addr, entry)
    else if (domain) pushTo(index.compiledCompounds.byDomain, domain.domain, entry)
    else index.compiledCompounds.unanchored.push(entry)
  })
  index.compiledCompounds.size = compounds.length
  index.compiledCaseCodes = index.caseCodeRules.map((rule) => ({
    rule,
    upper: rule.signal.toUpperCase(),
  }))
  const compileSubjects = (list: Rule[]): CompiledSubjectRule[] =>
    list.map((rule) => ({ rule, subj: compileSubjectNeedle(rule.signal) }))
  index.compiledCourtSubjects = compileSubjects(index.courtCaseSubjectRules)
  index.compiledManualSubjects = compileSubjects(index.manualSubjectKeywordRules)
  index.compiledAutoSubjects = compileSubjects(index.subjectKeywordRules)
  return index
}

function caseCodeReason(upper: string, raw: string, f: EmailFeatures): string | null {
  if (!upper || !f.subject) return null
  return f.subjectUpper.includes(upper) ? `主旨含案件代號「${raw}」` : null
}

function subjectReason(s: SubjectNeedle, f: EmailFeatures): string | null {
  if (!s.raw || !f.subject) return null
  // Canonical case-number equivalence (優化 2026-07): the learning side
  // stores the COMPACT form (「112訴304」) but official court notices use the
  // FULL form (「112年度訴字第304號」). Plain substring never matched across
  // the two forms — so a case-number rule couldn't even match the notice
  // that taught it, and (never firing) got hard-deleted by the stale sweep.
  // When the signal IS a pure case number, compare against the subject's
  // extracted case-number set; otherwise fall through to the literal.
  if (s.court && f.courtCases().includes(s.court)) return `主旨含案號「${s.court}」`
  // Whitespace-collapsed + digit-boundary-aware substring. Collapse fixes
  // learn-side normalized / match-side raw whitespace mismatches; the
  // boundary stops 「112訴20」 from matching 「112訴204」.
  return subjectIncludesBoundary(f.hay, s.needle) ? `主旨含「${s.raw}」` : null
}

/** First party (From first) on `domain`, as the reason; fromOnly restricts to From. */
function domainReason(domain: string, f: EmailFeatures, fromOnly = false): string | null {
  if (!domain) return null
  for (let i = 0; i < f.addrs.length; i++) {
    if (f.domains[i] !== domain) continue
    const isFrom = f.addrs[i] === f.from
    if (fromOnly && !isFrom) continue
    return `${isFrom ? '寄件人' : '收件人'}網域 @${domain}`
  }
  return null
}

function senderReason(addr: string, raw: string, f: EmailFeatures): string | null {
  if (!addr) return null
  return f.from === addr ? `寄件人 ${raw}` : null
}

function compoundReason(c: CompiledCompound, f: EmailFeatures): string | null {
  const reasons: string[] = []
  for (const cond of c.conds) {
    const r =
      cond.type === 'domain'
        ? domainReason(cond.domain, f)
        : cond.type === 'subject_keyword'
          ? subjectReason(cond.subj, f)
          : senderReason(cond.addr, cond.raw, f)
    if (!r) return null
    reasons.push(r)
  }
  return reasons.join(' + ')
}

function matchCompoundTier(index: RuleIndex, f: EmailFeatures): MatchOutcome | null {
  const c = index.compiledCompounds
  if (c.size === 0) return null
  const candidates: CompiledCompound[] = [...c.unanchored]
  if (f.from) {
    const bySender = c.bySender.get(f.from)
    if (bySender) candidates.push(...bySender)
  }
  const seen = new Set<string>()
  for (const d of f.domains) {
    if (!d || seen.has(d)) continue
    seen.add(d)
    const byDomain = c.byDomain.get(d)
    if (byDomain) candidates.push(...byDomain)
  }
  candidates.sort((a, b) => a.order - b.order)
  for (const cc of candidates) {
    const reason = compoundReason(cc, f)
    if (reason) return { rule: cc.rule, reason }
  }
  return null
}

function matchSubjectTier(list: readonly CompiledSubjectRule[], f: EmailFeatures): MatchOutcome | null {
  for (const s of list) {
    const reason = subjectReason(s.subj, f)
    if (reason) return { rule: s.rule, reason }
  }
  return null
}

/**
 * Domain tier: for every address (From first) look up its domain's bucket,
 * keep the candidate with the highest effective confidence. Own-domain
 * addresses never reach AUTO rules; for user_manual they count on the From
 * side only (see buildRuleIndex).
 */
function matchDomainTier(
  map: Map<string, Rule[]>,
  f: EmailFeatures,
  ownDomains: ReadonlySet<string>,
  manual: boolean,
): MatchOutcome | null {
  if (map.size === 0) return null
  let best: MatchOutcome | null = null
  let bestConf = -1
  for (let i = 0; i < f.addrs.length; i++) {
    const d = f.domains[i]!
    if (!d) continue
    const own = ownDomains.has(d)
    if (own && (!manual || f.addrs[i] !== f.from)) continue
    const top = map.get(d)?.[0] // bucket sorted desc — first dominates
    if (!top) continue
    const conf = effectiveConfidence(top)
    if (conf > bestConf) {
      const reason = domainReason(d, f, own)
      if (!reason) continue
      bestConf = conf
      best = { rule: top, reason }
    }
  }
  return best
}

function matchSenderTier(map: Map<string, Rule[]>, f: EmailFeatures): MatchOutcome | null {
  if (!f.from) return null
  for (const r of map.get(f.from) ?? []) {
    const reason = senderReason(r.signal.toLowerCase().trim(), r.signal, f)
    if (reason) return { rule: r, reason }
  }
  return null
}

/**
 * Match a single email against a pre-built RuleIndex — tiers in the order
 * documented in the file header; the first tier with a hit wins.
 *
 * The user_manual tier (2026-11-05): previously the walk went strictly by
 * type and ignored source, so a hand-written sender or subject rule never
 * fired whenever ANY auto domain rule matched ANY party of the email (To/Cc
 * included) — and thread memory then beat it too, because ruleBeatsThread
 * was handed the auto domain rule. (The old advice to "give the sender rule
 * a high explicit confidence" could not work: confidence only orders rules
 * WITHIN a type.) Structural identifiers (case_code / compound / court-case
 * subject) still come first, whatever their source. Inside the user_manual
 * tier the type order stays TYPE_PRIORITY's (domain → subject → sender), the
 * contract the 2026-05-22 incident pinned.
 */
export function matchEmailWithIndex(email: Email, index: RuleIndex): MatchOutcome | null {
  const f = emailFeatures(email)
  // 1. case_code
  for (const c of index.compiledCaseCodes) {
    const reason = caseCodeReason(c.upper, c.rule.signal, f)
    if (reason) return { rule: c.rule, reason }
  }
  // 2. compound
  const compound = matchCompoundTier(index, f)
  if (compound) return compound
  // 3. court-case subject rules (優化 2026-07) — a structural case-number
  // identifier outranks a broad plain-domain rule (same-client-multiple-
  // cases: 主旨帶 B 案案號的信不再被 @client.com→A 案規則搶先誤歸).
  const court = matchSubjectTier(index.compiledCourtSubjects, f)
  if (court) return court
  // 4. user_manual: domain → subject_keyword → sender
  const manual =
    matchDomainTier(index.manualDomainMap, f, index.ownDomains, true) ??
    matchSubjectTier(index.compiledManualSubjects, f) ??
    matchSenderTier(index.manualSenderMap, f)
  if (manual) return manual
  // 5. auto domain — O(K) Map lookups, K = address count (≤ 6 normally)
  const domain = matchDomainTier(index.domainMap, f, index.ownDomains, false)
  if (domain) return domain
  // 6. Internal-body case pass (batch-3 — after domain, before plain
  // subject_keyword). See matchBodyCaseWithIndex. Requires the fully-
  // fetched bodyText, so it is inert at preflight (BodyPreview only) and
  // fires in the AI path after getMessageBody (service-worker classifyAi).
  const body = matchBodyCaseWithIndex(email, index)
  if (body) return body
  // 7. auto subject_keyword — linear, longest signal first
  const subject = matchSubjectTier(index.compiledAutoSubjects, f)
  if (subject) return subject
  // 8. auto sender — O(1) lookup by From address
  return matchSenderTier(index.senderMap, f)
}

/**
 * Internal-body case pass (batch-3). Route a previously-unmatched email on a
 * case number found in its BODY — but ONLY:
 *   - when the subject carries NO case number (caseSignalsForMatch subject-
 *     first gate), and
 *   - the body has exactly ONE distinct case identifier (!bodyAmbiguous —
 *     bodies quote the opponent's OTHER case / prior cases), and
 *   - it matches a STRUCTURAL case rule (case_code / court-case), never a
 *     plain subject / domain / compound rule.
 * caseSignalsForMatch reads ONLY the fully-fetched bodyText, so this is inert
 * until a caller supplies it (the AI path, post getMessageBody) — never on the
 * 250-char preflight preview where ambiguity can't be judged. Exported so the
 * classifyAi loop can route body-case reps by rule (skipping the AI) after it
 * fetches the full body, keeping matchEmailWithIndex the single source of the
 * pass logic.
 */
export function matchBodyCaseWithIndex(email: Email, index: RuleIndex): MatchOutcome | null {
  if (!email.bodyText) return null
  const bodyCase = caseSignalsForMatch(email)
  if (bodyCase.source !== 'body' || bodyCase.bodyAmbiguous) return null
  for (const c of index.compiledCaseCodes) {
    if (bodyCase.caseCodes.some((code) => code.toUpperCase() === c.upper)) {
      return { rule: c.rule, reason: `內文含案件代號「${c.rule.signal}」` }
    }
  }
  for (const s of index.compiledCourtSubjects) {
    if (s.subj.court && bodyCase.courtCases.includes(s.subj.court)) {
      return { rule: s.rule, reason: `內文含案號「${s.subj.court}」` }
    }
  }
  return null
}

/**
 * Sort comparator for rule matching: type priority first, then
 * effective confidence (empirical accuracy when the rule has fired enough,
 * else the configured confidence), then recency.
 *
 * Using effective confidence here gives the engine a self-cleaning
 * property — a rule with bad track record drops in priority without
 * the user manually disabling it. Combined with the rule-health UI
 * surfacing "low accuracy" buckets, this nudges users to clean noisy
 * rules instead of leaving them at the top.
 */
export function compareRulePriority(a: Rule, b: Rule): number {
  const p = TYPE_PRIORITY[a.type] - TYPE_PRIORITY[b.type]
  if (p !== 0) return p
  const aConf = effectiveConfidence(a)
  const bConf = effectiveConfidence(b)
  if (aConf !== bConf) return bConf - aConf
  const aUsed = a.lastUsedAt ?? a.createdAt
  const bUsed = b.lastUsedAt ?? b.createdAt
  return bUsed.localeCompare(aUsed)
}

/**
 * Canonical domain: trimmed, leading '@' dropped, lowercased (2026-11-05:
 * trims first — `' @x.com'` used to keep its '@' and never match).
 */
function normalizeDomain(input: string): string {
  const s = input.trim()
  return (s.startsWith('@') ? s.slice(1) : s).trim().toLowerCase()
}

// ---- Generic-provider domain blocklist ------------------------------------
//
// Email providers used by countless unrelated senders. A domain rule keyed
// on these (e.g. "gmail.com → 某資料夾") is almost never the right routing
// signal — same domain serves every household / business / random sender on
// the internet. The auto-generators (initial-scan / ai_confirmed /
// ai_overridden) must skip building plain-domain rules for these.
//
// Compound rules (domain + 主旨關鍵字) are still allowed: those carry the
// extra subject signal that gives the routing meaning.
//
// Manual rules (user_manual) are never blocked — if the lawyer wants one,
// they can have it.
const GENERIC_PROVIDER_DOMAINS: ReadonlySet<string> = new Set([
  // Google
  'gmail.com', 'googlemail.com',
  // Yahoo
  'yahoo.com', 'yahoo.com.tw', 'yahoo.co.jp', 'ymail.com', 'rocketmail.com',
  // Microsoft
  'hotmail.com', 'hotmail.com.tw', 'outlook.com', 'live.com', 'msn.com',
  // Apple
  'icloud.com', 'me.com', 'mac.com',
  // Greater-China providers
  'qq.com', '163.com', '126.com', 'sina.com', 'sina.com.cn',
  'foxmail.com', 'aliyun.com',
  // Privacy-focused providers
  'protonmail.com', 'proton.me', 'pm.me',
  // Other commonly-seen consumer mail
  'aol.com', 'gmx.com', 'gmx.net', 'mail.com', 'zoho.com',
])

export function isGenericProviderDomain(domain: string): boolean {
  if (!domain) return false
  return GENERIC_PROVIDER_DOMAINS.has(normalizeDomain(domain))
}

function allEmailAddresses(email: Email): string[] {
  const out: string[] = []
  const from = email.From?.EmailAddress?.Address
  if (from) out.push(from.toLowerCase())
  for (const list of [email.ToRecipients, email.CcRecipients, email.BccRecipients]) {
    if (!list) continue
    for (const r of list) {
      const a = r.EmailAddress?.Address
      if (a) out.push(a.toLowerCase())
    }
  }
  return out
}

// ---- Compound --------------------------------------------------------------

export type CompoundCondition =
  | { type: 'domain'; value: string }
  | { type: 'subject_keyword'; value: string }
  | { type: 'sender'; value: string }

export type CompoundSignal = { conditions: CompoundCondition[] }

export function encodeCompound(conditions: CompoundCondition[]): string {
  return JSON.stringify({ conditions })
}

const COMPOUND_CONDITION_TYPES: ReadonlySet<string> = new Set([
  'domain',
  'subject_keyword',
  'sender',
])

/**
 * Validate parsed compound conditions (2026-11-05): every entry must be an
 * object with a known `type` and a non-blank string `value`, else the whole
 * compound is invalid (null). An imported `{"type":"domain","value":123}`
 * used to throw inside matching for EVERY email — and the sweep's catch then
 * silently skipped all cleanup.
 */
function validCompoundConditions(raw: unknown): CompoundCondition[] | null {
  if (!Array.isArray(raw)) return null
  const out: CompoundCondition[] = []
  for (const c of raw) {
    if (!c || typeof c !== 'object') return null
    const { type, value } = c as { type?: unknown; value?: unknown }
    if (typeof type !== 'string' || !COMPOUND_CONDITION_TYPES.has(type)) return null
    if (typeof value !== 'string' || value.trim().length === 0) return null
    out.push({ type, value } as CompoundCondition)
  }
  return out
}

/** Decode a compound rule signal; null when it is not a valid compound. */
export function decodeCompound(signal: string): CompoundSignal | null {
  if (typeof signal !== 'string') return null
  try {
    const parsed = JSON.parse(signal) as { conditions?: unknown } | null
    if (!parsed || typeof parsed !== 'object') return null
    const conditions = validCompoundConditions(parsed.conditions)
    return conditions ? { conditions } : null
  } catch {
    return null
  }
}

/**
 * Conditions of a compound for KEY purposes: the rule-signal form
 * `{conditions:[…]}` or the bare-array form normalizeSignal itself emits —
 * which is what older builds stored as a tombstone's signalNorm.
 */
function compoundConditionsForKey(signal: string): CompoundCondition[] | null {
  try {
    const parsed = JSON.parse(signal) as unknown
    if (Array.isArray(parsed)) return validCompoundConditions(parsed)
    if (parsed && typeof parsed === 'object') {
      return validCompoundConditions((parsed as { conditions?: unknown }).conditions)
    }
    return null
  } catch {
    return null
  }
}

// ---- Conflict detection ----------------------------------------------------

export type RuleConflict = {
  type: RuleType
  signal: string
  rules: Rule[] // 2+ rules with the same key, different target folders
}

/** Canonical subject needle: a pure case number in compact form, else
 *  lowercased + whitespace-collapsed (the matcher treats both as equal). */
function normalizeSubjectValue(signal: string): string {
  return courtCaseSignal(signal) ?? collapseSubjectWhitespace(signal.toLowerCase()).trim()
}

/**
 * Canonical identity of a signal — the key for tombstones, dedupe, conflict
 * detection and every learner's "already covered" check. Two signals the
 * MATCHER treats as equal must normalize equal (2026-11-05): full vs compact
 * court-case forms and full-width digits (「112年度訴字第304號」≡「112訴304」),
 * subject whitespace runs, and a padded / '@'-prefixed domain — before this
 * they were distinct keys, so tombstones, dedupe, conflict detection and
 * alreadyCovered all missed equivalents. Idempotent on its own output (see
 * tombstoneKey).
 */
export function normalizeSignal(type: RuleType, signal: string): string {
  if (typeof signal !== 'string') return ''
  switch (type) {
    case 'domain':
      return normalizeDomain(signal)
    case 'sender':
      return signal.toLowerCase().trim()
    case 'case_code':
      return signal.toUpperCase().trim()
    case 'subject_keyword':
      return normalizeSubjectValue(signal)
    case 'compound': {
      // Two compound rules with the same conditions in different orders
      // should be considered identical for conflict detection. Canonicalize
      // each value by its condition type, sort by (type, value), then
      // re-stringify so the key is canonical.
      const conds = compoundConditionsForKey(signal)
      if (!conds) return signal
      const canon = conds.map((c) => ({
        type: c.type,
        value:
          c.type === 'domain'
            ? normalizeDomain(c.value)
            : c.type === 'sender'
              ? c.value.toLowerCase().trim()
              : normalizeSubjectValue(c.value),
      }))
      canon.sort((a, b) =>
        a.type !== b.type ? a.type.localeCompare(b.type) : a.value.localeCompare(b.value),
      )
      return JSON.stringify(canon)
    }
  }
  return signal
}

export function findConflicts(rules: Rule[]): RuleConflict[] {
  const enabled = rules.filter((r) => r.enabled)
  const groups = new Map<string, Rule[]>()
  for (const r of enabled) {
    const key = `${r.type}::${normalizeSignal(r.type, r.signal)}`
    const arr = groups.get(key) ?? []
    arr.push(r)
    groups.set(key, arr)
  }
  const conflicts: RuleConflict[] = []
  for (const [key, arr] of groups) {
    if (arr.length < 2) continue
    const folders = new Set(arr.map((r) => r.targetFolderId))
    if (folders.size > 1) {
      conflicts.push({
        type: arr[0]!.type,
        signal: key.split('::').slice(1).join('::'),
        rules: arr,
      })
    }
  }
  return conflicts
}

// ---- Case-code extraction (for #8 initial scan) ---------------------------

// Looking at the user's case codes (e.g. "25A0067A", "26E0103A"):
//   2 digits + 1 letter + 4 digits + 1 letter
// Case-insensitive so subjects with "25a0067a" still match; we uppercase the
// capture group when emitting.
const CASE_CODE_RE = /\b(\d{2}[A-Z]\d{4}[A-Z])\b/gi

export function extractCaseCodes(text: string): string[] {
  if (!text) return []
  const codes = new Set<string>()
  for (const m of text.matchAll(CASE_CODE_RE)) {
    codes.add(m[1]!.toUpperCase())
  }
  return [...codes]
}

// ---- Taiwan court case number extraction ----------------------------------
//
// Taiwan court cases use 「年度 + 字 + 號」 naming. Compact form drops the
// 年度/字/號 connectors, e.g.:
//   112訴204         (年度 112, 訴字, 第 204 號)
//   114訴更一14      (年度 114, 訴更(一)字, 第 14 號)
//   114民著訴74      (年度 114, 民著訴字, 第 74 號)
//
// The user's actual compound rules are built around these — they're more
// important than the pure-Latin case codes the previous regex handled.
//
// Recognition strategy:
//   - Compact form: year (1XX) + 1-5 Chinese chars (case type) + 1-5 digits
//   - Full form:    year (1XX) + 年(度)? + Chinese chars + 字? + 第 + digits + 號?
// Lookbehind/ahead negation prevents matching parts of a longer number
// (avoids "2025年5月" → false match starting at "025").
//
// Date stopword filter: if the "case type" capture contains 年/月/日/時/分/秒,
// we reject the match — it's almost certainly a date expression that
// happened to fit the shape (e.g. "112年5月" looks like compact form).

// Case-type char class allows full-width / half-width parentheses so
// 分案字別 like 「訴更(一)字」/「訴更（一）」 are captured; normalizeCaseType
// strips the parens back out so the canonical compact form is 「訴更一」.
const COURT_CASE_COMPACT_RE = /(?<!\d)(1\d{2})([一-龥（）()]{1,5})(\d{1,5})(?!\d)/g
const COURT_CASE_FULL_RE = /(?<!\d)(1\d{2})年(?:度)?([一-龥（）()]{1,5}?)字?第(\d{1,5})號?/g
const DATE_STOPWORD_RE = /[年月日時分秒週]/

// Anchored variants — used to test whether a whole SIGNAL string *is* a
// court case number (vs merely contains one). A full-subject signal that
// embeds a case number (e.g. 「112訴204 通知會議」) must NOT be treated as a
// case-number signal — it keeps plain substring semantics.
const COURT_CASE_COMPACT_ANCHOR = /^(1\d{2})([一-龥（）()]{1,5})(\d{1,5})$/
const COURT_CASE_FULL_ANCHOR = /^(1\d{2})年(?:度)?([一-龥（）()]{1,5}?)字?第(\d{1,5})號?$/

/** Fold full-width digits (０-９, U+FF10–FF19) to ASCII so Taiwanese court
 *  e-notices that use full-width numerals normalize to the same canonical
 *  form as half-width. */
function foldFullWidthDigits(text: string): string {
  return text.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xff10 + 0x30))
}

/**
 * Normalize a captured case-type segment by dropping the 字 / 第 connectors
 * and any parentheses that some forms include but the compact form omits.
 * e.g. "上字第" → "上", "訴字" → "訴", "訴更(一)字" → "訴更一". If stripping
 * leaves an empty string the case has no real type name and we reject it.
 */
function normalizeCaseType(type: string): string {
  return type.replace(/[字第（）()]/g, '')
}

export function extractCourtCaseNumbers(text: string): string[] {
  if (!text) return []
  const folded = foldFullWidthDigits(text)
  const cases = new Set<string>()
  // Full form first (anchors on 年/字/第 keywords are unambiguous) so we
  // don't mis-extract from substrings that the compact pattern could also
  // match.
  for (const m of folded.matchAll(COURT_CASE_FULL_RE)) {
    const [, year, type, num] = m
    if (!year || !type || !num) continue
    if (DATE_STOPWORD_RE.test(type)) continue
    const cleanType = normalizeCaseType(type)
    if (!cleanType) continue
    cases.add(`${year}${cleanType}${num}`)
  }
  for (const m of folded.matchAll(COURT_CASE_COMPACT_RE)) {
    const [, year, type, num] = m
    if (!year || !type || !num) continue
    if (DATE_STOPWORD_RE.test(type)) continue
    const cleanType = normalizeCaseType(type)
    if (!cleanType) continue
    cases.add(`${year}${cleanType}${num}`)
  }
  return [...cases]
}

/**
 * If `signal` IS (not merely contains) a single court case number — in
 * compact 「112訴304」 or full 「112年度訴字第304號」 form — return its
 * canonical compact form; else null. This gates the canonical-equivalence
 * matching below so only pure case-number rules get cross-form matching;
 * full-subject signals that happen to embed a case number keep substring
 * semantics. Cheap: two anchored regexes, no allocation on the common miss.
 */
export function courtCaseSignal(signal: string): string | null {
  const s = foldFullWidthDigits(signal.trim())
  for (const re of [COURT_CASE_COMPACT_ANCHOR, COURT_CASE_FULL_ANCHOR]) {
    const m = re.exec(s)
    if (!m) continue
    const [, year, type, num] = m
    if (!year || !type || !num) continue
    if (DATE_STOPWORD_RE.test(type)) continue
    const cleanType = normalizeCaseType(type)
    if (cleanType) return `${year}${cleanType}${num}`
  }
  return null
}

// ---- Body case-number gate (batch-3) ---------------------------------------
//
// Email bodies are far noisier than subjects — they quote prior cases
// ("參照 111上YY"), mention the opponent's OTHER case ("相對人另案 112訴XX"),
// cite attachments. So the SINGLE gate below, shared by BOTH the matching side
// (rules) and the learning side (execute), enforces the conservative policy:
//   1. Subject-first: if the subject carries ANY structured case identifier,
//      use ONLY those and ignore the body entirely.
//   2. Body only as fallback, and only structured IDs (court-case / case_code)
//      — full-subject needles never touch the body (catastrophic false match).
//   3. `bodyAmbiguous` flags >1 distinct body identifier so the learning side
//      can refuse to mint a rule (matching may still use them; learning won't).
// Making this ONE function is the core alignment point: it stops the learning
// side from minting a rule the matching side would never fire (which the stale
// sweep would then hard-delete after 100 days).

export type EligibleCaseSignals = {
  courtCases: string[]
  caseCodes: string[]
  source: 'subject' | 'body' | 'none'
  /** Body carried >1 distinct case identifier — too ambiguous to learn from. */
  bodyAmbiguous: boolean
}

export function gateCaseSignals(
  subjectCourt: string[],
  subjectCodes: string[],
  bodyCourt: string[],
  bodyCodes: string[],
): EligibleCaseSignals {
  if (subjectCourt.length > 0 || subjectCodes.length > 0) {
    return {
      courtCases: subjectCourt,
      caseCodes: subjectCodes,
      source: 'subject',
      bodyAmbiguous: false,
    }
  }
  const distinct = new Set([...bodyCourt, ...bodyCodes])
  if (distinct.size === 0) {
    return { courtCases: [], caseCodes: [], source: 'none', bodyAmbiguous: false }
  }
  return {
    courtCases: bodyCourt,
    caseCodes: bodyCodes,
    source: 'body',
    bodyAmbiguous: distinct.size > 1,
  }
}

// Body case extraction reads the FULL fetched `bodyText` on both sides — no
// separate window slice (batch-3 review fix). Two invariants ride on this:
//   (1) Symmetry: matching (here) and learning (actionToPlanItem) both extract
//       from the same `email.bodyText` string, so a learned case rule can
//       always fire and ambiguity is judged identically — bound by shared
//       input, not by two constants that could drift apart.
//   (2) Ambiguity safety: the veto (bodyAmbiguous = >1 distinct case) must see
//       EVERY case number in the body. Slicing to a narrower window would hide
//       a second, deeper case number, letting an email that merely CITES
//       another party's case get routed there without AI review. The single
//       upper bound is getMessageBody's own fetch cap (the string we hold).
// Never reads the 250-char BodyPreview — at preflight (no bodyText) the body
// pass simply doesn't fire; body-case mail flows to the AI (full body) there.
export function caseSignalsForMatch(
  email: Pick<Email, 'Subject' | 'bodyText'>,
): EligibleCaseSignals {
  const body = email.bodyText ?? ''
  return gateCaseSignals(
    extractCourtCaseNumbers(email.Subject ?? ''),
    extractCaseCodes(email.Subject ?? ''),
    extractCourtCaseNumbers(body),
    extractCaseCodes(body),
  )
}

/** Learning-side view. Body identifiers are pre-computed at classify time into
 *  `bodyCaseNumbers`/`bodyCaseCodes` from the SAME full body window the matching
 *  side reads — NO 250-char bodyPreview fallback, which would let learning mint
 *  a rule the matching side can never fire (→ stale-swept). Absent fields ⇒ no
 *  body case (the email was never body-fetched). */
export function caseSignalsForLearning(
  item: Pick<PlanItem, 'emailSubject' | 'bodyCaseNumbers' | 'bodyCaseCodes'>,
): EligibleCaseSignals {
  return gateCaseSignals(
    extractCourtCaseNumbers(item.emailSubject ?? ''),
    extractCaseCodes(item.emailSubject ?? ''),
    item.bodyCaseNumbers ?? [],
    item.bodyCaseCodes ?? [],
  )
}

// Whitespace-collapse (folds U+3000 ideographic space, NBSP, runs of ASCII
// space — all common in Taiwanese official mail) so a signal learned from a
// subject with odd spacing still matches. `\s` in JS already covers U+3000
// and U+00A0.
function collapseSubjectWhitespace(s: string): string {
  return s.replace(/\s+/g, ' ')
}

const ANY_DIGIT_RE = /[0-9０-９]/

/**
 * Boundary-aware substring test for subject matching. Plain `includes`
 * lets 「112訴20」 (case 20) match a subject containing 「112訴204」 (case
 * 204) — a silent misfile at the highest-trust rule tier. When the needle
 * ends (or starts) with a digit, require the adjacent subject character to
 * NOT be a digit (either width). Non-numeric needles behave exactly like
 * `includes`.
 */
function subjectIncludesBoundary(haystack: string, needle: string): boolean {
  if (!needle) return false
  const startDigit = ANY_DIGIT_RE.test(needle[0]!)
  const endDigit = ANY_DIGIT_RE.test(needle[needle.length - 1]!)
  if (!startDigit && !endDigit) return haystack.includes(needle)
  let from = 0
  for (;;) {
    const idx = haystack.indexOf(needle, from)
    if (idx < 0) return false
    const before = idx > 0 ? haystack[idx - 1]! : ''
    const after = idx + needle.length < haystack.length ? haystack[idx + needle.length]! : ''
    const okBefore = !startDigit || !ANY_DIGIT_RE.test(before)
    const okAfter = !endDigit || !ANY_DIGIT_RE.test(after)
    if (okBefore && okAfter) return true
    from = idx + 1
  }
}

// ---- Single-email subject signal (整段主旨, 2026-05-27 redesign) ----------
//
// `extractSubjectSignal` is for single-email learning via
// `chooseLearningSignal`. Design axiom: same-domain + same-normalized-
// subject → must be the same case (user couldn't distinguish them
// either). So when a domain conflict surfaces, the discriminator is
// always *the subject itself* — no need to extract tokens / proper
// nouns.
//
// Previous design (2026-05-22) tried to pull a "proper noun" feature
// (4+ Chinese chars / 6+ Latin chars) out of the subject. This was
// noisy: passing-mention tokens like "請款通知" or "Phoenix" would
// create overfit rules. The new design uses the entire normalized
// subject as the signal — accuracy via exactness, with daily stale
// sweep capping rule library growth (100 days of no use → auto-
// disable).
//
// Confidence ladder reflects specificity:
//   - compound (domain + full subject):  0.85
//   - subject_keyword (full subject):    0.80
//   - subject_keyword (court case):      0.85  (existing — structural ID)
//   - compound (court case):             0.90  (existing — structural ID)

/**
 * Minimum length (after `normalizeSubject`) for a subject to qualify
 * as a LEARNED routing signal. Below this the subject is too generic
 * ("通知", "請款通知", "季報") to safely match emails.
 *
 * Raised from 2 to MIN_NORMALIZED_SUBJECT_LEN (8) on 2026-11-05. "The
 * domain anchor disambiguates" did not hold: a 2–7 char subject condition
 * minted on first sight, substring-matched unrelated mail of the same
 * domain / sender, and — being a compound — beat thread memory as well.
 * 8 is also the floor thread memory uses, so a subject too short to be
 * remembered is too short to become a permanent rule.
 */
const MIN_SUBJECT_SIGNAL_LEN = MIN_NORMALIZED_SUBJECT_LEN

/**
 * Convert a raw email subject into a routing signal — the entire
 * normalized subject string, or '' when too short / empty.
 *
 * Returns the same form `normalizeSubject` produces (reply/forward
 * prefixes stripped, lowercased, whitespace collapsed) so:
 *   - "Re: 關於甲公司股權移轉案" → "關於甲公司股權移轉案"
 *   - "FW: [External] Project Phoenix" → "project phoenix"
 *   - "Re: Re: 通知" → "" (too short to learn)
 *
 * Match-time semantics: `matchSubjectKeyword` does case-insensitive
 * substring match against the raw email subject. Since the signal is
 * already lowercase + prefix-stripped, it will match the raw subject
 * as a substring regardless of whatever Re:/Fwd: noise sits in front.
 */
export function extractSubjectSignal(rawSubject: string): string {
  if (!rawSubject) return ''
  const normalized = normalizeSubject(rawSubject)
  if (normalized.length < MIN_SUBJECT_SIGNAL_LEN) return ''
  return normalized
}

export function extractDomain(address: string | undefined | null): string | null {
  if (!address) return null
  const at = address.lastIndexOf('@')
  if (at < 0) return null
  return address.slice(at + 1).toLowerCase()
}
