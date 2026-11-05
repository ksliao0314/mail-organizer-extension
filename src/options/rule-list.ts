// Rule-list helpers for the Options page (2026-11-05).
//
// Kept out of App.tsx so they're unit-testable and so the component module
// keeps exporting only components (Vite fast-refresh boundary).

import type { Rule } from '@/shared/types'

function sameValue(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true
  if (a && b && typeof a === 'object' && typeof b === 'object') {
    return JSON.stringify(a) === JSON.stringify(b)
  }
  return false
}

/** Field-by-field equality of two rule records (Rule fields are flat). */
export function sameRule(a: Rule, b: Rule): boolean {
  if (a === b) return true
  const ka = Object.keys(a) as Array<keyof Rule>
  const kb = Object.keys(b) as Array<keyof Rule>
  if (ka.length !== kb.length) return false
  for (const k of ka) {
    if (!sameValue(a[k], b[k])) return false
  }
  return true
}

/**
 * Structural sharing for a freshly fetched rule list: every rule that is
 * unchanged keeps its PREVIOUS object identity, and if nothing changed at
 * all the previous array itself is returned (→ React bails out of the
 * state update entirely).
 *
 * Why: getRules round-trips the whole library through structured clone, so
 * every refresh used to mint new objects for all ~2000 rules, defeating the
 * memoized RuleRow and re-rendering the full table after every toggle —
 * and the Options page now refreshes on every external 'rules' change too.
 */
export function reconcileRuleList(prev: readonly Rule[], next: readonly Rule[]): Rule[] {
  const byId = new Map<string, Rule>()
  for (const r of prev) byId.set(r.id, r)
  let identical = prev.length === next.length
  const out = next.map((r, i) => {
    const old = byId.get(r.id)
    const kept = old && sameRule(old, r) ? old : r
    if (kept !== prev[i]) identical = false
    return kept
  })
  return identical ? (prev as Rule[]) : out
}

/** `bulkRuleAction` may report counts or id lists — normalise to a count. */
export function countOf(v: unknown): number {
  if (Array.isArray(v)) return v.length
  return typeof v === 'number' && Number.isFinite(v) ? v : 0
}

/** The SW's "no such handler" answers (older SW builds). */
export function isUnknownHandler(code: string | undefined): boolean {
  return code === 'UNKNOWN_TYPE' || code === 'UNKNOWN_REQUEST'
}
