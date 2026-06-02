// Rule health diagnostics.
//
// Categorizes a rule set into health buckets so the user can decide what to
// prune, refine, or fix. Pure function — UI computes once on render. No
// dependency on storage / chrome / network, so it's straightforward to unit
// test against any synthetic Rule[].
//
// We deliberately don't include "suspected missing rules" — that needs an
// execute-history feed which the codebase doesn't yet maintain.

import { findConflicts, isGenericProviderDomain } from './rules'
import type { Rule } from './types'

export type RuleHealthOptions = {
  /** `Date.now()` injected for testability. */
  now?: number
  /** A rule never used + older than this many days is "sleeping". */
  neverUsedAgeDays?: number
  /** A rule used at least once but not for this many days is also "sleeping". */
  staleUsedAgeDays?: number
  /** matchCount ≥ this counts a rule as "hot" for the hot-but-vague check. */
  hotMatchThreshold?: number
  /**
   * Patterns matching "vague" / unspecific target folder names. A rule with
   * matchCount ≥ hotMatchThreshold AND target path matching ANY of these is
   * flagged as "hot but vague" — suggests the user is bucketing a lot of
   * traffic into a generic folder and should split it.
   */
  vagueTargetPatterns?: readonly RegExp[]
}

const DEFAULT_VAGUE_PATTERNS: readonly RegExp[] = [
  /未分類/,
  /待釐清/,
  /待處理/,
  /其他/,
  /雜項/,
  /暫存/,
  /\bTBD\b/i,
  /\bTODO\b/i,
  /\bmisc\b/i,
  /^Inbox$/i,
]

const DAY_MS = 86_400_000

export type RuleHealthReport = {
  /** Enabled rules that haven't been useful for a long time (or ever). */
  sleeping: Rule[]
  /** Hot-traffic rules whose target folder looks like a catch-all bucket. */
  hotVague: Rule[]
  /** Rules flagged orphaned by reconcile — the target folder is gone. */
  orphaned: Rule[]
  /**
   * Plain-domain rules whose signal is a generic email provider (gmail.com,
   * yahoo.com, hotmail.com, etc.). One provider serves countless unrelated
   * senders so a folder-mapping by domain alone is almost never correct.
   * Suggest the user upgrade to compound (domain + 主旨關鍵字) or delete.
   */
  overBroad: Rule[]
  /** Rule IDs participating in at least one (type, signal) conflict. */
  conflictRuleIds: Set<string>
  /** Tallies, for fast badge rendering without recomputing. */
  counts: {
    sleeping: number
    hotVague: number
    orphaned: number
    overBroad: number
    conflicts: number
  }
}

export function computeRuleHealth(
  rules: Rule[],
  options: RuleHealthOptions = {},
): RuleHealthReport {
  const now = options.now ?? Date.now()
  const neverUsedAgeDays = options.neverUsedAgeDays ?? 30
  const staleUsedAgeDays = options.staleUsedAgeDays ?? 90
  const hotMatchThreshold = options.hotMatchThreshold ?? 10
  const vaguePatterns = options.vagueTargetPatterns ?? DEFAULT_VAGUE_PATTERNS

  const sleeping: Rule[] = []
  const hotVague: Rule[] = []
  const orphaned: Rule[] = []
  const overBroad: Rule[] = []

  for (const r of rules) {
    if (r.orphaned) {
      orphaned.push(r)
      // Orphaned rules are surfaced separately — don't double-count as sleeping.
      continue
    }
    if (!r.enabled) continue // ignore disabled rules for sleeping/hot signals

    // Over-broad detection — plain-domain rule on a generic provider.
    // Independent of activity (matchCount/lastUsedAt) because the problem
    // is structural: gmail.com → folder X mis-routes everyone who happens
    // to use Gmail to write to the user, regardless of how often it fires.
    if (r.type === 'domain' && isGenericProviderDomain(r.signal)) {
      overBroad.push(r)
      // Don't continue — a rule can also be sleeping or hot-vague AND
      // over-broad; surfacing in multiple buckets is fine because each
      // bucket frames a different "what to do about it" action.
    }

    // Sleeping detection
    const createdMs = Date.parse(r.createdAt)
    const lastUsedMs = r.lastUsedAt ? Date.parse(r.lastUsedAt) : null
    if (r.matchCount === 0 && Number.isFinite(createdMs)) {
      const ageDays = (now - createdMs) / DAY_MS
      if (ageDays >= neverUsedAgeDays) {
        sleeping.push(r)
        continue
      }
    } else if (lastUsedMs && Number.isFinite(lastUsedMs)) {
      const idleDays = (now - lastUsedMs) / DAY_MS
      if (idleDays >= staleUsedAgeDays) {
        sleeping.push(r)
        continue
      }
    }

    // Hot-but-vague detection
    if (r.matchCount >= hotMatchThreshold) {
      const target = r.targetFolderPath
      if (vaguePatterns.some((p) => p.test(target))) {
        hotVague.push(r)
      }
    }
  }

  // Sort each bucket by most informative first
  sleeping.sort((a, b) => {
    // Oldest createdAt first
    return Date.parse(a.createdAt) - Date.parse(b.createdAt)
  })
  hotVague.sort((a, b) => b.matchCount - a.matchCount)
  orphaned.sort((a, b) => a.signal.localeCompare(b.signal, 'zh-Hant'))
  overBroad.sort((a, b) => b.matchCount - a.matchCount)

  const conflicts = findConflicts(rules)
  const conflictRuleIds = new Set<string>()
  for (const c of conflicts) for (const r of c.rules) conflictRuleIds.add(r.id)

  return {
    sleeping,
    hotVague,
    orphaned,
    overBroad,
    conflictRuleIds,
    counts: {
      sleeping: sleeping.length,
      hotVague: hotVague.length,
      orphaned: orphaned.length,
      overBroad: overBroad.length,
      conflicts: conflicts.length,
    },
  }
}
