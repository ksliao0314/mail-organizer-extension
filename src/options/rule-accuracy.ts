// Rule accuracy as the Options page shows it (2026-11-06).
//
// ONE formula everywhere: accepted ÷ all fires, where fires = matchCount
// (accepted hits) + overrideCount (the user corrected the rule's verdict) —
// i.e. shared/rules.ts ruleFires / ruleErrorRate, the same numbers the
// high-error-rate retirement uses. Before this only the rule row, the drawer
// and the summary line used it; the folder group header, its default
// collapse, the conflicts view and the cleanup list / warning still used the
// pre-2026-11-05 (match − override) ÷ match. So 2 hits + 4 overrides read 33 %
// in the row but −100 % in its folder header, and the cleanup warnings
// (≥ 5 / ≥ 10 MATCHES) never flagged a rule the user mostly overrides.

import { ruleErrorRate, ruleFires } from '@/shared/rules'
import type { Rule } from '@/shared/types'

type Counts = Pick<Rule, 'matchCount' | 'overrideCount'>

/** Accepted ÷ fires as a whole percent; null for a rule that never fired. */
export function ruleAccuracyPct(r: Counts): number | null {
  return ruleFires(r) > 0 ? Math.round((1 - ruleErrorRate(r)) * 100) : null
}

/** The same over a set of rules (folder header, library summary). */
export function aggregateAccuracy(rules: readonly Counts[]): {
  fires: number
  /** 0…1, null when none of the rules ever fired. */
  accuracy: number | null
} {
  let fires = 0
  let overrides = 0
  for (const r of rules) {
    fires += ruleFires(r)
    overrides += r.overrideCount ?? 0
  }
  return { fires, accuracy: fires > 0 ? (fires - overrides) / fires : null }
}

/** Fired at least `minFires` times and was overridden in more than `maxErrorRate` of them. */
export function isLowAccuracy(r: Counts, minFires: number, maxErrorRate = 0.3): boolean {
  return ruleFires(r) >= minFires && ruleErrorRate(r) > maxErrorRate
}
