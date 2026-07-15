// Rule engine.
//
// Storage:
//   chrome.storage.local['rules'] = Rule[]  (managed via getRules/setRules in storage.ts)
//
// Matching priority (lower number = checked first):
//   1. case_code        — specific identifier in subject (highest signal)
//   2. domain           — email-address suffix on any party
//   3. compound         — JSON-encoded AND of (domain | subject_keyword | sender)
//   4. subject_keyword  — substring in subject
//   5. sender           — exact From address
//
// Within the same type, higher `confidence` wins, then later `lastUsedAt`.

import {
  addRuleTombstones,
  clearMatchingTombstones,
  getRuleTombstones,
  getRules,
  recordRuleEvents,
  setRules,
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
import { normalizeSubject } from './normalize'

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

/**
 * Transactional rule mutation: read → transform → write, all under lock.
 * Use this for any operation that needs to read current state to compute next.
 */
export async function mutateRules<T>(
  fn: (rules: Rule[]) => { next: Rule[]; result: T } | Promise<{ next: Rule[]; result: T }>,
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
    if (next !== current) await setRules(next)
    return result
  })
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

export async function upsertRule(rule: Rule, opts?: { actor?: Actor }): Promise<Rule> {
  const actor = opts?.actor ?? 'system'
  // Enforce the per-type confidence cap on every write (see
  // applyConfidenceCap). For user_manual rules this is a no-op.
  const capped = applyConfidenceCap(rule)
  let prevSnapshot: RuleSnapshot | undefined
  const result = await mutateRules(async (rules) => {
    // Tombstone clear lives INSIDE the rules lock — otherwise an auto
    // generator could read tombstones (seeing this rule's ghost), filter
    // us out, and write rules between our clear and our write.
    await clearMatchingTombstones([
      {
        type: capped.type,
        signalNorm: normalizeSignal(capped.type, capped.signal),
        targetFolderPath: capped.targetFolderPath,
      },
    ]).catch((e) => console.warn('[mail-organizer] clearMatchingTombstones (upsert)', e))
    const idx = rules.findIndex((r) => r.id === capped.id)
    if (idx >= 0) prevSnapshot = snapshotOf(rules[idx]!)
    const next = [...rules]
    if (idx >= 0) next[idx] = capped
    else next.push(capped)
    return { next, result: capped }
  })
  const after = snapshotOf(capped)
  if (prevSnapshot) {
    const changedFields = diffSnapshots(prevSnapshot, after)
    if (changedFields.length > 0) {
      await logRuleEvents([
        {
          kind: 'edit',
          ruleId: capped.id,
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
      { kind: 'create', ruleId: capped.id, at: Date.now(), actor, after },
    ])
  }
  return result
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
        deletedAt: Date.now(),
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
      ...(enabled
        ? { autoDisabledAt: undefined, autoDisabledReason: undefined }
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
        // its confidence frozen here.
        const errorRate = nextCount > 0 ? (r.overrideCount ?? 0) / nextCount : 0
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

const STALE_AGE_DAYS_DEFAULT = 100
// High-error-rate auto-disable thresholds. A rule with this many
// override events out of total matches is clearly mis-routing; let it
// rest. User can re-enable in the rules UI if the recent overrides
// were exceptional.
const HIGH_ERROR_RATE_MIN_SAMPLES = 20
const HIGH_ERROR_RATE_THRESHOLD = 0.5

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

export async function dedupeRulesByKey(): Promise<DedupeReport> {
  let removed = 0
  let groupsAffected = 0
  await mutateRules((rules) => {
    const buckets = new Map<string, Rule[]>()
    for (const r of rules) {
      const key = `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
      const arr = buckets.get(key) ?? []
      arr.push(r)
      buckets.set(key, arr)
    }
    const keepIds = new Set<string>()
    for (const arr of buckets.values()) {
      if (arr.length === 1) {
        keepIds.add(arr[0]!.id)
        continue
      }
      groupsAffected++
      removed += arr.length - 1
      const sorted = [...arr].sort((a, b) => {
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
        const rA = SOURCE_RANK[a.source]
        const rB = SOURCE_RANK[b.source]
        if (rA !== rB) return rA - rB
        if (a.matchCount !== b.matchCount) return b.matchCount - a.matchCount
        return a.createdAt.localeCompare(b.createdAt)
      })
      keepIds.add(sorted[0]!.id)
    }
    if (removed === 0) {
      return { next: rules, result: undefined }
    }
    return {
      next: rules.filter((r) => keepIds.has(r.id)),
      result: undefined,
    }
  })
  return { removed, groupsAffected }
}

export type StaleAutoDisableReport = {
  /** Rules whose `enabled` was flipped to false in this sweep (legacy_token
   *  and high-error-rate buckets). Stay in storage; surface in the rules
   *  UI as "已自動休眠" so user can re-enable if needed. */
  disabled: Rule[]
  /** Counts by reason. `stale` counts rules that were hard-deleted (not
   *  in `disabled` — kept here for telemetry). */
  byReason: { stale: number; highErrorRate: number; legacyToken: number }
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

/**
 * Auto-retire rules in three scenarios — two different fates per bucket:
 *
 * 1. **Stale** (HARD DELETE, no trace):
 *    - never matched + older than `staleDays` days, OR
 *    - matchCount > 0 but lastUsedAt older than `staleDays` days
 *    The signal isn't reaching the inbox anymore. Per user request:
 *      "100 天沒用的規則就完全刪除、不留紀錄、不影響將來學習"
 *    Removed from rules array entirely. NO tombstone (AI can re-learn
 *    the same signal/target later if mail flow changes). NO audit
 *    event (Rule events log stays focused on real user / AI actions).
 *    NO error log (this is the design, not an error).
 *
 * 2. **High error rate** (SOFT DISABLE):
 *    - matchCount ≥ HIGH_ERROR_RATE_MIN_SAMPLES
 *    - overrideCount/matchCount ≥ HIGH_ERROR_RATE_THRESHOLD
 *    The rule is mis-routing — user keeps fixing its verdicts. Surface
 *    in "自動休眠" so user can decide whether the overrides were
 *    exceptional or the rule is fundamentally wrong.
 *
 * 3. **Legacy token** (SOFT DISABLE):
 *    - Pre-2026-05-27 subject_keyword rule from the deprecated
 *      token/fragment design. Surface in "自動休眠" so user can
 *      reactivate any that are still relevant.
 *
 * Skipped (never touched):
 *   - already-disabled rules
 *   - `source === 'user_manual'` (sacred)
 *   - `orphaned === true` (handled by reconcile UI)
 */
export async function autoDisableStaleRules(
  options: { now?: number; staleDays?: number } = {},
): Promise<StaleAutoDisableReport> {
  const now = options.now ?? Date.now()
  const staleDays = options.staleDays ?? STALE_AGE_DAYS_DEFAULT
  const cutoffMs = staleDays * 86_400_000
  const disabled: Rule[] = []
  const byReason = { stale: 0, highErrorRate: 0, legacyToken: 0 }
  const auditEvents: RuleEvent[] = []

  await mutateRules((rules) => {
    const next: Rule[] = []
    for (const r of rules) {
      // Untouched paths — keep as-is.
      if (!r.enabled || r.source === 'user_manual' || r.orphaned) {
        next.push(r)
        continue
      }

      // Determine which bucket (if any) this rule falls into.
      let reason: 'stale' | 'high-error-rate' | 'legacy_token' | null = null
      if (isLegacyTokenSubjectRule(r)) {
        reason = 'legacy_token'
      } else if (r.matchCount === 0) {
        const createdMs = Date.parse(r.createdAt)
        if (Number.isFinite(createdMs) && now - createdMs >= cutoffMs) {
          reason = 'stale'
        }
      } else if (r.matchCount >= HIGH_ERROR_RATE_MIN_SAMPLES) {
        const errRate = (r.overrideCount ?? 0) / r.matchCount
        if (errRate >= HIGH_ERROR_RATE_THRESHOLD) {
          reason = 'high-error-rate'
        }
      }
      // Last-used stale (added 2026-05-27): rule fired in the past but
      // has not been used for `staleDays` days. Catches "topic ran for
      // a few weeks then moved on" patterns.
      if (!reason) {
        const lastUsedStr = r.lastUsedAt ?? r.createdAt
        const lastUsedMs = Date.parse(lastUsedStr)
        if (Number.isFinite(lastUsedMs) && now - lastUsedMs >= cutoffMs && r.matchCount > 0) {
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
        autoDisabledAt: new Date(now).toISOString(),
        autoDisabledReason: reason,
      }
      next.push(updated)
      disabled.push(updated)
      if (reason === 'high-error-rate') byReason.highErrorRate++
      else byReason.legacyToken++
      auditEvents.push({
        kind: 'toggle',
        ruleId: r.id,
        at: now,
        actor: 'system',
        enabled: false,
        type: r.type,
        signal: r.signal,
        targetFolderPath: r.targetFolderPath,
      })
    }
    // Nothing disabled AND nothing stale-deleted → `next` holds exactly the
    // same refs in the same order. Return the ORIGINAL array so mutateRules'
    // identity check skips the full-library rewrite (this runs on every SW
    // wake via the module-load sweep).
    if (disabled.length === 0 && byReason.stale === 0) {
      return { next: rules, result: undefined }
    }
    return { next, result: undefined }
  })

  if (auditEvents.length > 0) {
    await logRuleEvents(auditEvents)
  }
  return { disabled, byReason }
}

/**
 * Minimum hits before we trust the empirical accuracy number over the
 * configured confidence. Below this, the configured confidence (which
 * encodes "intent" — user_manual gets 0.95, auto_scan gets 0.65, etc.)
 * is the better signal because the empirical sample is too small.
 */
const ACCURACY_MIN_SAMPLES = 10

/**
 * Configured confidence blended with empirical accuracy. When a rule has
 * fired enough times (>= ACCURACY_MIN_SAMPLES), trust the empirical
 * track record; otherwise fall back to the configured value. Used by
 * matchEmail's tiebreak so a rule that's been proven unreliable
 * (high overrideCount) demotes itself naturally.
 */
export function effectiveConfidence(r: Pick<Rule, 'confidence' | 'matchCount' | 'overrideCount'>): number {
  if (r.matchCount >= ACCURACY_MIN_SAMPLES) {
    const overrides = r.overrideCount ?? 0
    const empirical = (r.matchCount - overrides) / r.matchCount
    return Math.max(0, Math.min(1, empirical))
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
    // Auto-re-enable (2026-05-22): when a candidate matches an
    // auto-disabled existing rule (same type+signalNorm+target,
    // `autoDisabledAt` set), the user is implicitly confirming the
    // dormant rule was correct. Flip it back to `enabled=true` instead
    // of just skipping. User-disabled rules (no `autoDisabledAt`) stay
    // disabled — that's user intent.
    //
    // Key shape mirrors filterByTombstones: (type, signalNorm, target).
    const existingByKey = new Map<string, { idx: number; rule: Rule }>()
    rules.forEach((r, idx) => {
      const key = `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
      existingByKey.set(key, { idx, rule: r })
    })
    const seenInBatch = new Set<string>()
    const reEnabledIndexes = new Map<number, Rule>()
    survivors = afterTombstone.filter((r) => {
      const key = `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
      const existing = existingByKey.get(key)
      if (existing) {
        const e = existing.rule
        // Re-enable auto-disabled rule when the same signal+target
        // gets re-learned. User intent is implicit: "this routing
        // is still correct, you just hadn't seen mail for a while".
        if (!e.enabled && e.autoDisabledAt) {
          reEnabledIndexes.set(existing.idx, {
            ...e,
            enabled: true,
            autoDisabledAt: undefined,
            // Audit: also clear autoDisabledReason to honour the types.ts
            // contract ("Cleared along with autoDisabledAt on re-enable")
            // and match toggleRule — otherwise the rule ends up enabled
            // with a stale 'high-error-rate'/'legacy_token' reason that
            // round-trips through export/sync and misreports the rule.
            autoDisabledReason: undefined,
          })
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

/**
 * Given a set of candidate rules + the current tombstones, return the
 * subset whose (type, signalNorm, targetFolderPath) is NOT tombstoned.
 * Auto-generators (initial-scan, ai_confirmed, ai_overridden) should run
 * candidates through this before calling addRules to honor prior user
 * deletions.
 */
export function filterByTombstones(candidates: Rule[], tombstones: RuleTombstone[]): Rule[] {
  if (tombstones.length === 0) return candidates
  const blocked = new Set(
    tombstones.map((t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`),
  )
  return candidates.filter((r) => {
    const key = `${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`
    return !blocked.has(key)
  })
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
// This map is the SINGLE SOURCE OF TRUTH for type ordering. It controls:
//   1. `compareRulePriority` — sorts rules by type before tiebreaking on
//      effective confidence + recency.
//   2. `matchEmailWithIndex` — iterates buckets in this exact order.
//
// If you change a value here, double-check both call sites still behave
// as intended. There was a 2026-05-22 incident where matchEmailWithIndex
// was rewritten to iterate sender BEFORE domain/subject_keyword, with
// TYPE_PRIORITY left at the old values — the divergence silently broke
// working subject_keyword rules. Lesson: this map and matchEmailWithIndex
// must stay in lockstep.
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

/**
 * Pre-built rule index for batch matching. Bucketed by type so the
 * matcher can short-circuit on type priority and use O(1) Map lookups
 * for domain / sender rules (the two types where exact-string match is
 * the dominant cost in a 500+ rule library).
 *
 * Each bucket is pre-sorted by `compareRulePriority` so iteration order
 * inside a type returns the best candidate first.
 */
export type RuleIndex = {
  caseCodeRules: Rule[]
  compoundRules: Rule[]
  /**
   * subject_keyword rules whose signal IS a court case number (優化
   * 2026-07). Kept in their own bucket and matched AFTER compound but
   * BEFORE domain — a structural unique identifier must outrank a broad
   * plain-domain rule (the same-client-multiple-cases hole). This bucket
   * order INTENTIONALLY diverges from TYPE_PRIORITY (subject_keyword=4);
   * compareRulePriority still sorts these as priority 4 for within-library
   * ordering, only the match path elevates them.
   */
  courtCaseSubjectRules: Rule[]
  /** Map from normalized domain → rules targeting that domain. Sorted by priority. */
  domainMap: Map<string, Rule[]>
  /** Subject-keyword rules sorted by effective confidence desc (no easy index). */
  subjectKeywordRules: Rule[]
  /** Map from lowercased sender address → rules. Sorted by priority. */
  senderMap: Map<string, Rule[]>
}

export function buildRuleIndex(rules: Rule[]): RuleIndex {
  const enabled = rules.filter((r) => r.enabled && !r.orphaned).sort(compareRulePriority)
  const index: RuleIndex = {
    caseCodeRules: [],
    compoundRules: [],
    courtCaseSubjectRules: [],
    domainMap: new Map(),
    subjectKeywordRules: [],
    senderMap: new Map(),
  }
  // Subject_keyword has no good index — it's linear substring scan. To
  // avoid "通知" (shorter, broad) winning over "112訴204" (longer,
  // specific) just because the broad one happens to sort first by
  // confidence, we override the bucket's sort: longest signal first,
  // then effective confidence, then recency. This is local to the
  // bucket; doesn't affect other rule types' priority.
  const subjectKeywordSort = (a: Rule, b: Rule): number => {
    const lenDelta = b.signal.length - a.signal.length
    if (lenDelta !== 0) return lenDelta
    const confDelta = effectiveConfidence(b) - effectiveConfidence(a)
    if (confDelta !== 0) return confDelta
    const aUsed = a.lastUsedAt ?? a.createdAt
    const bUsed = b.lastUsedAt ?? b.createdAt
    return bUsed.localeCompare(aUsed)
  }
  for (const r of enabled) {
    switch (r.type) {
      case 'case_code':
        index.caseCodeRules.push(r)
        break
      case 'compound':
        index.compoundRules.push(r)
        break
      case 'domain': {
        const d = normalizeDomain(r.signal)
        if (!d) break
        let bucket = index.domainMap.get(d)
        if (!bucket) {
          bucket = []
          index.domainMap.set(d, bucket)
        }
        bucket.push(r)
        break
      }
      case 'subject_keyword':
        // Pure court-case-number subject rules go in their own high-priority
        // bucket (優化 2026-07); everything else stays a plain subject rule.
        if (courtCaseSignal(r.signal)) index.courtCaseSubjectRules.push(r)
        else index.subjectKeywordRules.push(r)
        break
      case 'sender': {
        const s = r.signal.toLowerCase().trim()
        if (!s) break
        let bucket = index.senderMap.get(s)
        if (!bucket) {
          bucket = []
          index.senderMap.set(s, bucket)
        }
        bucket.push(r)
        break
      }
    }
  }
  // Re-sort subject_keyword bucket by longest-signal-first. Other
  // buckets already have the right sort from `compareRulePriority`.
  index.subjectKeywordRules.sort(subjectKeywordSort)
  // Court-case bucket: distinct case numbers rarely collide, but keep the
  // same longest-first / confidence sort for determinism.
  index.courtCaseSubjectRules.sort(subjectKeywordSort)
  return index
}

/**
 * Match a single email against a pre-built RuleIndex. Mirrors
 * matchEmail's semantics (highest-priority type that matches wins, with
 * effective-confidence tiebreak inside the bucket) but uses Map lookups
 * to skip rules whose signal can't possibly match the email's
 * address/domain set.
 *
 * Domain match: iterate every address in the email; collect candidate
 * rules from `domainMap` keyed by each address's normalized domain;
 * pick the one with the highest effective confidence among matches.
 */
export function matchEmailWithIndex(email: Email, index: RuleIndex): MatchOutcome | null {
  // Order matches TYPE_PRIORITY: case_code(1) → compound(2) → domain(3)
  // → subject_keyword(4) → sender(5). Each bucket is pre-sorted by
  // `compareRulePriority` (or `subjectKeywordSort` for the subject
  // bucket — longest signal first).
  //
  // Historical note (2026-05-22): an earlier fix briefly promoted
  // sender above domain in match order, with the intent "sender is more
  // specific than domain so it should win". The promotion went further
  // than intended — it also placed sender above subject_keyword, which
  // broke working `工時審閱` style rules whenever the user had any
  // sender rule for the same From. Reverted to TYPE_PRIORITY order to
  // restore consistency between sort metadata and match path. Use a
  // `user_manual` sender rule with high explicit confidence if you
  // really need it to outrank a domain.

  // case_code (priority 1)
  for (const r of index.caseCodeRules) {
    const reason = matchCaseCode(r.signal, email)
    if (reason) return { rule: r, reason }
  }
  // compound (priority 2)
  for (const r of index.compoundRules) {
    const reason = matchCompound(r.signal, email)
    if (reason) return { rule: r, reason }
  }
  // court-case subject rules (優化 2026-07) — checked here, between compound
  // and domain, so a structural case-number identifier outranks a broad
  // plain-domain rule (same-client-multiple-cases: 主旨帶 B 案案號的信不再
  //被 @client.com→A 案規則搶先誤歸).
  for (const r of index.courtCaseSubjectRules) {
    const reason = matchSubjectKeyword(r.signal, email)
    if (reason) return { rule: r, reason }
  }
  // domain (priority 3) — O(K) where K = email address count (≤ 6 normally)
  let bestDomain: MatchOutcome | null = null
  let bestDomainConf = -1
  for (const addr of allEmailAddresses(email)) {
    const atIdx = addr.lastIndexOf('@')
    if (atIdx < 0) continue
    const domain = addr.slice(atIdx + 1)
    const candidates = index.domainMap.get(domain)
    if (!candidates) continue
    for (const r of candidates) {
      const reason = matchDomain(r.signal, email)
      if (reason) {
        const conf = effectiveConfidence(r)
        if (conf > bestDomainConf) {
          bestDomainConf = conf
          bestDomain = { rule: r, reason }
        }
        break // candidates sorted desc — first match dominates within this domain
      }
    }
  }
  if (bestDomain) return bestDomain
  // Internal-body case pass (batch-3, priority ~3.5 — after domain, before
  // plain subject_keyword). Fires ONLY when the subject carries no case number
  // AND the body has exactly ONE distinct case identifier (caseSignalsForMatch
  // gate). Bodies are noisy — they quote the opponent's OTHER case, cite prior
  // cases — so we route on a body case number only when it's unambiguous, and
  // only against STRUCTURAL case rules (case_code / court-case), never plain
  // subject / domain rules. Placed after domain so a stable domain rule wins
  // over a noisy body signal; anything ambiguous falls through to AI (safe).
  // NOTE: at preflight time email.bodyText is not yet fetched, so this reads
  // the 250-char BodyPreview — a first-250 hit routes here; a deeper one goes
  // to AI, which can then learn a rule for next time (B3-C6).
  const bodyCase = caseSignalsForMatch(email)
  if (bodyCase.source === 'body' && !bodyCase.bodyAmbiguous) {
    for (const r of index.caseCodeRules) {
      if (bodyCase.caseCodes.some((c) => c.toUpperCase() === r.signal.toUpperCase())) {
        return { rule: r, reason: `內文含案件代號「${r.signal}」` }
      }
    }
    for (const r of index.courtCaseSubjectRules) {
      const cc = courtCaseSignal(r.signal)
      if (cc && bodyCase.courtCases.includes(cc)) {
        return { rule: r, reason: `內文含案號「${cc}」` }
      }
    }
  }
  // subject_keyword (priority 4) — linear, no index possible without
  // bringing in a substring data structure (trie / aho-corasick).
  // Bucket is sorted longest-signal-first by buildRuleIndex so
  // "112訴204" wins over a shorter "通知" if both match.
  for (const r of index.subjectKeywordRules) {
    const reason = matchSubjectKeyword(r.signal, email)
    if (reason) return { rule: r, reason }
  }
  // sender (priority 5) — O(1) lookup by From address
  const fromAddr = email.From?.EmailAddress?.Address?.toLowerCase()
  if (fromAddr) {
    const candidates = index.senderMap.get(fromAddr)
    if (candidates) {
      for (const r of candidates) {
        const reason = matchSender(r.signal, email)
        if (reason) return { rule: r, reason }
      }
    }
  }
  return null
}

/**
 * Sort comparator for rule matching: type priority first, then
 * effective confidence (empirical accuracy when matchCount is high
 * enough, else the configured confidence), then recency.
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


function matchCaseCode(code: string, email: Email): string | null {
  if (!code || !email.Subject) return null
  if (email.Subject.toUpperCase().includes(code.toUpperCase())) {
    return `主旨含案件代號「${code}」`
  }
  return null
}

function normalizeDomain(input: string): string {
  return (input.startsWith('@') ? input.slice(1) : input).toLowerCase()
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

function matchDomain(signal: string, email: Email): string | null {
  const domain = normalizeDomain(signal)
  if (!domain) return null
  const addrs = allEmailAddresses(email)
  const hit = addrs.find((a) => a.endsWith('@' + domain))
  if (!hit) return null
  const fromAddr = email.From?.EmailAddress?.Address?.toLowerCase()
  const side = hit === fromAddr ? '寄件人' : '收件人'
  return `${side}網域 @${domain}`
}

function matchSender(address: string, email: Email): string | null {
  if (!address) return null
  const target = address.toLowerCase().trim()
  const from = email.From?.EmailAddress?.Address?.toLowerCase()
  if (from === target) return `寄件人 ${address}`
  return null
}

function matchSubjectKeyword(keyword: string, email: Email): string | null {
  if (!keyword || !email.Subject) return null
  // Canonical case-number equivalence (優化 2026-07): the learning side
  // stores the COMPACT form (「112訴304」) but official court notices use the
  // FULL form (「112年度訴字第304號」). Plain substring never matched across
  // the two forms — so a case-number rule couldn't even match the notice
  // that taught it, and (never firing) got hard-deleted by the 100-day
  // stale sweep. When the signal IS a pure case number, compare against the
  // subject's extracted case-number set instead.
  const cc = courtCaseSignal(keyword)
  if (cc) {
    if (extractCourtCaseNumbers(email.Subject).includes(cc)) {
      return `主旨含案號「${cc}」`
    }
    // fall through — the exact literal might still appear verbatim
  }
  // Whitespace-collapsed + digit-boundary-aware substring. Collapse fixes
  // learn-side normalized / match-side raw whitespace mismatches; the
  // boundary stops 「112訴20」 from matching 「112訴204」.
  const hay = collapseSubjectWhitespace(email.Subject.toLowerCase())
  const needle = collapseSubjectWhitespace(keyword.toLowerCase())
  if (subjectIncludesBoundary(hay, needle)) {
    return `主旨含「${keyword}」`
  }
  return null
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

export function decodeCompound(signal: string): CompoundSignal | null {
  try {
    const parsed = JSON.parse(signal) as CompoundSignal
    if (!parsed || !Array.isArray(parsed.conditions)) return null
    return parsed
  } catch {
    return null
  }
}

function matchCompound(signal: string, email: Email): string | null {
  const parsed = decodeCompound(signal)
  if (!parsed || parsed.conditions.length === 0) return null
  const reasons: string[] = []
  for (const cond of parsed.conditions) {
    let r: string | null = null
    if (cond.type === 'domain') r = matchDomain(cond.value, email)
    else if (cond.type === 'subject_keyword') r = matchSubjectKeyword(cond.value, email)
    else if (cond.type === 'sender') r = matchSender(cond.value, email)
    if (!r) return null
    reasons.push(r)
  }
  return reasons.join(' + ')
}

// ---- Conflict detection ----------------------------------------------------

export type RuleConflict = {
  type: RuleType
  signal: string
  rules: Rule[] // 2+ rules with the same key, different target folders
}

export function normalizeSignal(type: RuleType, signal: string): string {
  switch (type) {
    case 'domain':
      return normalizeDomain(signal)
    case 'sender':
      return signal.toLowerCase().trim()
    case 'case_code':
      return signal.toUpperCase().trim()
    case 'subject_keyword':
      return signal.toLowerCase().trim()
    case 'compound': {
      // Two compound rules with the same conditions in different orders
      // should be considered identical for conflict detection. Sort by
      // (type, value) then re-stringify so the key is canonical.
      const parsed = decodeCompound(signal)
      if (!parsed) return signal
      const sorted = [...parsed.conditions].sort((a, b) => {
        if (a.type !== b.type) return a.type.localeCompare(b.type)
        return a.value.toLowerCase().localeCompare(b.value.toLowerCase())
      })
      return JSON.stringify(sorted.map((c) => ({ type: c.type, value: c.value.toLowerCase().trim() })))
    }
  }
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

// Body case extraction operates on the SAME window on both sides so a learned
// rule can actually fire and ambiguity is judged consistently (batch-3 review
// fix). Matching reads ONLY the fully-fetched `bodyText` — NEVER the 250-char
// BodyPreview. Rationale: on a truncated preview the ambiguity check is blind
// to a second case number sitting past char 250, so an email that merely CITES
// another party's case in its first lines (真正主案在深處) could be silently
// routed to that cited case's folder without AI review. Requiring the full
// body means: at preflight (no bodyText) the body pass simply does not fire —
// body-case mail flows to the AI, which reads the full body and routes it
// correctly; the learned rule then fires on future SUBJECT occurrences.
const BODY_CASE_WINDOW = 800

export function caseSignalsForMatch(
  email: Pick<Email, 'Subject' | 'bodyText'>,
): EligibleCaseSignals {
  const body = (email.bodyText ?? '').slice(0, BODY_CASE_WINDOW)
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
 * as a routing signal. Below this the subject is too generic ("通知"
 * alone, single-char placeholders) to safely match emails.
 *
 * 2 chars is intentionally permissive — when paired with a domain in
 * a compound rule the domain itself disambiguates. Pure subject_keyword
 * rules (no domain) would benefit from a higher threshold, but the
 * gate above (P5) for those only fires when no domain is available, so
 * we accept the same threshold for simplicity.
 */
const MIN_SUBJECT_SIGNAL_LEN = 2

/**
 * Convert a raw email subject into a routing signal — the entire
 * normalized subject string, or '' when too short / empty.
 *
 * Returns the same form `normalizeSubject` produces (reply/forward
 * prefixes stripped, lowercased, whitespace collapsed) so:
 *   - "Re: 關於甲公司股權移轉案" → "關於甲公司股權移轉案"
 *   - "FW: [External] Project Phoenix" → "project phoenix"
 *   - "Re: Re: 通知" → "通知"
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
