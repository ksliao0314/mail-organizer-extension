// Rule import / export.
//
// Serialize a rule set into a versioned JSON payload that can be saved as a
// file, then imported back on another machine (or after a clean reinstall).
//
// Format invariants:
//   - The top-level `type` discriminator MUST be 'mail-organizer-rules'.
//     Refuse anything else — we don't want to accidentally treat a stray JSON
//     file as a rule import.
//   - schemaVersion lets us evolve the format later. Currently v1.
//   - Each rule is strictly validated (type + signal must be non-empty,
//     enums must match) before we let it touch storage.
//
// Strategies on import:
//   - 'replace' — clear all existing rules, write imported set verbatim
//   - 'merge'   — keep existing rules; only add imported rules whose id is
//                 not already present AND whose (type, normalizeSignal)
//                 isn't held by a stored rule of equal or higher priority
//                 (planImportMerge). Existing matchCount / lastUsedAt stay
//                 intact; a lower-priority stored rule an imported one takes
//                 over is switched off as 'superseded' — after the folder
//                 reconcile, only for an imported rule whose folder resolved
//                 (applyImportTakeovers).

import { decodeCompound, normalizeSignal } from './rules'
import type { Rule, RuleSource, RuleTombstone, RuleType } from './types'

export const RULE_PAYLOAD_TYPE = 'mail-organizer-rules' as const
// v1: rules only
// v2: rules + tombstones (so deletions transfer when syncing across machines).
//     Validator preserves auto-disable state (autoDisabledAt /
//     autoDisabledReason / overrideCount / orphaned) — v1 stripped these.
export const RULE_SCHEMA_VERSION = 2

const VALID_RULE_TYPES: readonly RuleType[] = [
  'case_code',
  'compound',
  'domain',
  'subject_keyword',
  'sender',
]

const VALID_RULE_SOURCES: readonly RuleSource[] = [
  'auto_scan',
  'ai_confirmed',
  'ai_overridden',
  'user_manual',
]

/**
 * Hard cap on imported rule count. A normal user has tens to low-hundreds of
 * rules; 10k is generous and guards against a malformed/malicious file
 * exhausting memory before we can validate it.
 */
export const RULE_IMPORT_MAX = 10_000

export type RuleExportPayload = {
  type: typeof RULE_PAYLOAD_TYPE
  schemaVersion: number
  exportedAt: string
  appVersion: string
  ruleCount: number
  rules: Rule[]
  /**
   * v2+: tombstones for rules the user explicitly deleted. Included
   * so that "syncing rules across machines" actually transfers
   * deletion intent — without it, a deleted rule on machine A would
   * get auto-recreated on machine B as soon as it sees the matching
   * email pattern.
   *
   * Optional in v2 to allow callers that don't need them (e.g. partial
   * exports) to omit. Validator treats absent as empty array.
   */
  tombstones?: RuleTombstone[]
  tombstoneCount?: number
}

export function serializeRules(
  rules: Rule[],
  appVersion: string,
  tombstones: RuleTombstone[] = [],
): string {
  const payload: RuleExportPayload = {
    type: RULE_PAYLOAD_TYPE,
    schemaVersion: RULE_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    appVersion,
    ruleCount: rules.length,
    rules,
    tombstones,
    tombstoneCount: tombstones.length,
  }
  return JSON.stringify(payload, null, 2)
}

export type ParseRulesResult =
  | {
      ok: true
      rules: Rule[]
      /** v2+: tombstones from the payload. v1 payloads → empty array. */
      tombstones: RuleTombstone[]
      meta: { exportedAt: string; appVersion: string; schemaVersion: number }
    }
  | { ok: false; error: string }

export function parseRulesPayload(json: string): ParseRulesResult {
  let parsed: unknown
  try {
    parsed = JSON.parse(json)
  } catch (e) {
    return { ok: false, error: `JSON 解析失敗:${e instanceof Error ? e.message : String(e)}` }
  }
  if (!parsed || typeof parsed !== 'object') {
    return { ok: false, error: '檔案內容不是 JSON 物件' }
  }
  const obj = parsed as Record<string, unknown>

  if (obj.type !== RULE_PAYLOAD_TYPE) {
    return {
      ok: false,
      error: `不是規則匯出檔(缺少 type: "${RULE_PAYLOAD_TYPE}" 標記)`,
    }
  }
  if (typeof obj.schemaVersion !== 'number' || obj.schemaVersion > RULE_SCHEMA_VERSION) {
    return {
      ok: false,
      error: `不支援的 schemaVersion(此版本支援到 ${RULE_SCHEMA_VERSION},檔案是 ${String(obj.schemaVersion)})`,
    }
  }
  if (!Array.isArray(obj.rules)) {
    return { ok: false, error: 'rules 欄位不是陣列' }
  }
  if (obj.rules.length > RULE_IMPORT_MAX) {
    return {
      ok: false,
      error: `規則數量超過上限(${RULE_IMPORT_MAX} 條),拒絕匯入`,
    }
  }

  const validRules: Rule[] = []
  for (let i = 0; i < obj.rules.length; i++) {
    const r = obj.rules[i] as Record<string, unknown> | null | undefined
    const check = validateRule(r, i)
    if (!check.ok) return { ok: false, error: check.error }
    validRules.push(check.rule)
  }

  // v2: optional tombstones. v1 payloads simply have no tombstones field
  // → empty list, no behaviour change. Validate each entry minimally so
  // a malformed entry doesn't crash the importer.
  const validTombstones: RuleTombstone[] = []
  if (Array.isArray(obj.tombstones)) {
    for (let i = 0; i < obj.tombstones.length; i++) {
      const t = obj.tombstones[i] as Record<string, unknown> | null | undefined
      if (!t || typeof t !== 'object') continue
      if (typeof t.type !== 'string' || !VALID_RULE_TYPES.includes(t.type as RuleType)) continue
      if (typeof t.signalNorm !== 'string' || t.signalNorm.length === 0) continue
      if (typeof t.targetFolderPath !== 'string' || t.targetFolderPath.length === 0) continue
      if (typeof t.deletedAt !== 'number' || !Number.isFinite(t.deletedAt)) continue
      validTombstones.push({
        type: t.type as RuleType,
        // Re-canonicalize (2026-11-05): a backup from an older build stores
        // signalNorm in its pre-canonical form (full-form court case,
        // un-trimmed '@domain', raw compound values) — store the form
        // today's rules normalize to, so the imported deletion still blocks.
        signalNorm: normalizeSignal(t.type as RuleType, t.signalNorm) || t.signalNorm,
        targetFolderPath: t.targetFolderPath,
        deletedAt: t.deletedAt,
      })
    }
  }

  return {
    ok: true,
    rules: validRules,
    tombstones: validTombstones,
    meta: {
      exportedAt: typeof obj.exportedAt === 'string' ? obj.exportedAt : '',
      appVersion: typeof obj.appVersion === 'string' ? obj.appVersion : '',
      schemaVersion: obj.schemaVersion,
    },
  }
}

type ValidateResult = { ok: true; rule: Rule } | { ok: false; error: string }

function validateRule(r: Record<string, unknown> | null | undefined, idx: number): ValidateResult {
  if (!r || typeof r !== 'object') {
    return { ok: false, error: `第 ${idx + 1} 筆規則不是物件` }
  }
  const where = (field: string) => `第 ${idx + 1} 筆規則 ${field}`

  if (typeof r.id !== 'string' || r.id.length === 0) {
    return { ok: false, error: `${where('id')} 必須是非空字串` }
  }
  if (typeof r.type !== 'string' || !VALID_RULE_TYPES.includes(r.type as RuleType)) {
    return {
      ok: false,
      error: `${where('type')} 必須是 ${VALID_RULE_TYPES.join(' / ')} 之一(實際:${String(r.type)})`,
    }
  }
  if (typeof r.signal !== 'string' || r.signal.length === 0) {
    return { ok: false, error: `${where('signal')} 必須是非空字串` }
  }
  // Compound structure (2026-11-05): `{"conditions":[{"type":"domain",
  // "value":123}]}` used to import fine and then throw inside matching for
  // EVERY email (and inside the sweep, whose catch silently skipped all
  // cleanup). Reject anything decodeCompound — the matcher's own parser —
  // wouldn't accept, plus the condition-less form that can never match.
  if (r.type === 'compound') {
    const decoded = decodeCompound(r.signal)
    if (!decoded || decoded.conditions.length === 0) {
      return {
        ok: false,
        error: `${where('signal')} 不是有效的組合條件(需為 {"conditions":[{"type","value"}…]},type 為 domain / subject_keyword / sender、value 為非空字串)`,
      }
    }
  }
  if (typeof r.targetFolderId !== 'string') {
    return { ok: false, error: `${where('targetFolderId')} 必須是字串(可為空字串)` }
  }
  // Reject sentinels and obviously-bogus IDs at import time so they don't leak
  // into storage. Real Outlook IDs are long base64-like strings (100+ chars);
  // anything shorter that isn't empty is either a stale `pending:<...>` /
  // PLACEHOLDER sentinel or a hand-crafted import we shouldn't trust.
  if (
    r.targetFolderId !== '' &&
    (r.targetFolderId.length < 20 ||
      r.targetFolderId.startsWith('pending:') ||
      r.targetFolderId.startsWith('PLACEHOLDER'))
  ) {
    return {
      ok: false,
      error: `${where('targetFolderId')} 看起來不是有效的 Outlook 資料夾 ID(可為空字串或長度 ≥ 20 的真實 ID)`,
    }
  }
  if (typeof r.targetFolderPath !== 'string' || r.targetFolderPath.length === 0) {
    return { ok: false, error: `${where('targetFolderPath')} 必須是非空字串` }
  }
  if (typeof r.confidence !== 'number' || !Number.isFinite(r.confidence) || r.confidence < 0 || r.confidence > 1) {
    return { ok: false, error: `${where('confidence')} 必須是 0 到 1 的數字` }
  }
  if (typeof r.matchCount !== 'number' || !Number.isInteger(r.matchCount) || r.matchCount < 0) {
    return { ok: false, error: `${where('matchCount')} 必須是非負整數` }
  }
  if (typeof r.enabled !== 'boolean') {
    return { ok: false, error: `${where('enabled')} 必須是 boolean` }
  }
  if (typeof r.createdAt !== 'string' || r.createdAt.length === 0) {
    return { ok: false, error: `${where('createdAt')} 必須是非空字串` }
  }
  if (r.lastUsedAt !== undefined && typeof r.lastUsedAt !== 'string') {
    return { ok: false, error: `${where('lastUsedAt')} 若存在須為字串` }
  }
  if (typeof r.source !== 'string' || !VALID_RULE_SOURCES.includes(r.source as RuleSource)) {
    return {
      ok: false,
      error: `${where('source')} 必須是 ${VALID_RULE_SOURCES.join(' / ')} 之一(實際:${String(r.source)})`,
    }
  }

  const rule: Rule = {
    id: r.id,
    type: r.type as RuleType,
    signal: r.signal,
    targetFolderId: r.targetFolderId,
    targetFolderPath: r.targetFolderPath,
    confidence: r.confidence,
    matchCount: r.matchCount,
    enabled: r.enabled,
    createdAt: r.createdAt,
    source: r.source as RuleSource,
  }
  if (typeof r.lastUsedAt === 'string') rule.lastUsedAt = r.lastUsedAt
  // v2 payload preserves rule-lifecycle metadata that v1 silently
  // dropped. Without this, sync would lose:
  //   - overrideCount (the empirical accuracy track record)
  //   - autoDisabledAt / autoDisabledReason (why a rule is dormant)
  //   - orphaned flag (whether the target folder still exists on this
  //     machine — caller's reconcile re-evaluates this anyway, but
  //     keeping the bit for transit is harmless)
  if (typeof r.overrideCount === 'number' && Number.isFinite(r.overrideCount) && r.overrideCount >= 0) {
    rule.overrideCount = r.overrideCount
  }
  if (typeof r.autoDisabledAt === 'string') {
    rule.autoDisabledAt = r.autoDisabledAt
  }
  if (
    r.autoDisabledReason === 'stale' ||
    r.autoDisabledReason === 'high-error-rate' ||
    r.autoDisabledReason === 'legacy_token' ||
    // case_closed MUST survive import/transit: the two-phase retirement
    // sweep's hard-delete predicate keys on this value — stripping it here
    // would leave the other machine's copy soft-disabled forever (never
    // purged) and break cross-machine convergence.
    r.autoDisabledReason === 'case_closed' ||
    // Same for the 2026-11-05 reasons: the dormant purge and the learners'
    // revive check key on them.
    r.autoDisabledReason === 'superseded' ||
    r.autoDisabledReason === 'override'
  ) {
    rule.autoDisabledReason = r.autoDisabledReason
  }
  if (typeof r.orphaned === 'boolean') {
    rule.orphaned = r.orphaned
  }
  return { ok: true, rule }
}

// ---- Strategy application --------------------------------------------------

export type ImportStrategy = 'merge' | 'replace'

export type ImportPreview = {
  totalToImport: number
  /** Imported rules that 'merge' would skip: same id as a stored rule, a
   *  stored rule of equal or higher priority for the same (type, normalized
   *  signal) — see planImportMerge — or the same (type, normalized signal,
   *  target) as one earlier in the file. */
  duplicateCount: number
  /** Of duplicateCount: the stored rule that blocks it points at a DIFFERENT
   *  folder — skipped so the import can't create a conflict or bring back a
   *  mapping the user replaced (2026-11-06). */
  conflictCount: number
  /** Imported rules that would be added under 'merge'. */
  newCount: number
  /** Of newCount: imported rules that take a signal over from a stored rule
   *  of LOWER priority, or a switched-off one (2026-11-06, R5). */
  replaceCount: number
  /** Stored rules 'merge' may switch off ('superseded') because an imported
   *  rule of higher priority routes their signal to another folder — at
   *  most: only an imported rule whose folder resolves at import time takes
   *  over (applyImportTakeovers; importRules reports the actual number). */
  supersedeCount: number
  /** Current rule count before import. */
  existingCount: number
  /** Tombstones in the payload (v2+); 0 for v1 imports. */
  tombstoneCount: number
}

function signalKey(r: Pick<Rule, 'type' | 'signal'>): string {
  return `${r.type}::${normalizeSignal(r.type, r.signal)}`
}

/** Import priority of a rule's source: user_manual > ai_overridden > ai_confirmed > auto_scan. */
const SOURCE_PRIORITY: Record<RuleSource, number> = {
  user_manual: 3,
  ai_overridden: 2,
  ai_confirmed: 1,
  auto_scan: 0,
}

const priorityOf = (r: Pick<Rule, 'source'>): number => SOURCE_PRIORITY[r.source] ?? -1

type ImportMergePlan = {
  toAdd: Rule[]
  conflicting: number
  replacing: number
  /** Stored rule (identity) → the imported rules that would take its signal
   *  over; applied by applyImportTakeovers once the folders are reconciled. */
  takeovers: Map<Rule, Rule[]>
}

/**
 * The imported rules 'merge' would add (2026-11-05). Keyed on id FIRST:
 * re-importing an older backup of a rule edited since (new signal / target)
 * used to pass the old raw (type, signal) check and land as a second row
 * with the SAME id. Then on the canonical (type, normalized signal) — so
 * `@x.com` vs `x.com` or full vs compact court-case forms count as the same
 * signal — WITHOUT the target (2026-11-06): a stored rule for the signal
 * wins whatever folder it points at. Keying on the target added the backup's
 * old mapping next to the one the user replaced it with — a conflict that
 * also brought the replaced mapping back.
 *
 * …but only a stored rule of EQUAL OR HIGHER priority wins (R5, 2026-11-06):
 * an imported user_manual rule used to lose to any stored rule for its
 * signal, even an auto_scan one or one switched off. A stored rule blocks an
 * imported one when its source ranks at least as high AND it is live
 * (enabled, not orphaned) — or when the imported rule is not user_manual:
 * a dormant stored copy (user-disabled, superseded, …) still keeps an
 * imported AUTO rule from bringing the mapping back. When an enabled
 * imported rule does take a signal over, the live lower-priority stored
 * rules pointing elsewhere are switched off as 'superseded' (sync-safe —
 * the edit propagates — and revivable / purged like any superseded rule),
 * so the import never leaves two live rules routing one signal apart;
 * same-target ones are left for dedupe. That last step waits for the folder
 * reconcile (applyImportTakeovers, round 3). user_manual stored rules are
 * never touched. Within the file itself only exact (signal, target)
 * duplicates are dropped; the file's own structure is its business.
 */
function planImportMerge(existing: readonly Rule[], imported: readonly Rule[]): ImportMergePlan {
  const ids = new Set(existing.map((r) => r.id))
  const local = new Map<string, Rule[]>()
  for (const r of existing) {
    const key = signalKey(r)
    const bucket = local.get(key)
    if (bucket) bucket.push(r)
    else local.set(key, [r])
  }
  const fileTriples = new Set<string>()
  const toAdd: Rule[] = []
  const takeovers = new Map<Rule, Rule[]>()
  let conflicting = 0
  let replacing = 0
  for (const r of imported) {
    if (ids.has(r.id)) continue
    const key = signalKey(r)
    const same = local.get(key) ?? []
    const rank = priorityOf(r)
    const blockers = same.filter(
      (l) =>
        priorityOf(l) >= rank &&
        ((l.enabled && !l.orphaned) || r.source !== 'user_manual'),
    )
    if (blockers.length > 0) {
      if (!blockers.some((l) => l.targetFolderPath === r.targetFolderPath)) conflicting++
      continue
    }
    const triple = `${key}::${r.targetFolderPath}`
    if (fileTriples.has(triple)) continue
    ids.add(r.id)
    fileTriples.add(triple)
    toAdd.push(r)
    if (same.length > 0) replacing++
    if (!r.enabled) continue
    for (const l of same) {
      if (
        l.enabled &&
        l.source !== 'user_manual' &&
        priorityOf(l) < rank &&
        l.targetFolderPath !== r.targetFolderPath
      ) {
        const takers = takeovers.get(l)
        if (takers) takers.push(r)
        else takeovers.set(l, [r])
      }
    }
  }
  return { toAdd, conflicting, replacing, takeovers }
}

export function previewImport(
  existing: Rule[],
  imported: Rule[],
  tombstones: RuleTombstone[] = [],
): ImportPreview {
  const { toAdd, conflicting, replacing, takeovers } = planImportMerge(existing, imported)
  const newCount = toAdd.length
  return {
    totalToImport: imported.length,
    duplicateCount: imported.length - newCount,
    conflictCount: conflicting,
    newCount,
    replaceCount: replacing,
    supersedeCount: takeovers.size,
    existingCount: existing.length,
    tombstoneCount: tombstones.length,
  }
}

/**
 * The library an import produces: 'replace' — the file; 'merge' — the stored
 * rules (matchCount / lastUsedAt / user edits intact on every collision) plus
 * the imported rules planImportMerge adds. The takeovers of 'merge' are a
 * separate step, applyImportTakeovers, run once the folders are reconciled.
 */
export function applyImport(existing: Rule[], imported: Rule[], strategy: ImportStrategy): Rule[] {
  // replace: the file IS the library — but never write two rows with one id.
  if (strategy === 'replace') {
    const seen = new Set<string>()
    return imported.filter((r) => {
      if (seen.has(r.id)) return false
      seen.add(r.id)
      return true
    })
  }
  return [...existing, ...planImportMerge(existing, imported).toAdd]
}

/**
 * 'merge' takeovers (R5; round 3): switch off as 'superseded' each stored
 * rule an imported rule of higher priority takes over (planImportMerge) —
 * judged on `reconciled`, the imported library after the folder reconcile:
 * only when the stored rule is still live there AND at least one of its
 * takers is live there too (enabled, its folder resolved — not orphaned) and
 * still points elsewhere. Switching a working stored rule off for an
 * imported rule whose folder no longer exists left nothing routing the
 * signal. Run it only with a reconciled library: with no folder tree to
 * reconcile against, an imported rule's folder can't be confirmed.
 */
export function applyImportTakeovers(
  existing: readonly Rule[],
  imported: readonly Rule[],
  reconciled: readonly Rule[],
  nowIso: string = new Date().toISOString(),
): { next: Rule[]; superseded: Rule[] } {
  const { takeovers } = planImportMerge(existing, imported)
  const byId = new Map<string, Rule>()
  for (const r of reconciled) if (!byId.has(r.id)) byId.set(r.id, r)
  const live = (r: Rule | undefined): r is Rule => !!r && r.enabled && !r.orphaned
  const off = new Set<string>()
  for (const [stored, takers] of takeovers) {
    const now = byId.get(stored.id)
    if (!live(now) || now.source === 'user_manual') continue
    const taken = takers.some((t) => {
      const taker = byId.get(t.id)
      return live(taker) && taker.targetFolderPath !== now.targetFolderPath
    })
    if (taken) off.add(stored.id)
  }
  if (off.size === 0) return { next: [...reconciled], superseded: [] }
  const superseded: Rule[] = []
  const next = reconciled.map((r) => {
    if (!off.has(r.id) || !r.enabled) return r
    const updated: Rule = { ...r, enabled: false, autoDisabledAt: nowIso, autoDisabledReason: 'superseded' }
    superseded.push(updated)
    return updated
  })
  return { next, superseded }
}
