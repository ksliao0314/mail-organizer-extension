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
//   - 'merge'   — keep existing rules; only add imported rules whose
//                 (type, signal) does not already exist. Existing
//                 matchCount / lastUsedAt stay intact.

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
        signalNorm: t.signalNorm,
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
    r.autoDisabledReason === 'legacy_token'
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
  /** Imported rules whose (type, signal) already exists in current store. */
  duplicateCount: number
  /** Imported rules that would be added under 'merge'. */
  newCount: number
  /** Current rule count before import. */
  existingCount: number
  /** Tombstones in the payload (v2+); 0 for v1 imports. */
  tombstoneCount: number
}

function ruleKey(r: Pick<Rule, 'type' | 'signal'>): string {
  return `${r.type}::${r.signal}`
}

export function previewImport(
  existing: Rule[],
  imported: Rule[],
  tombstones: RuleTombstone[] = [],
): ImportPreview {
  const existingKeys = new Set(existing.map(ruleKey))
  let duplicateCount = 0
  for (const r of imported) {
    if (existingKeys.has(ruleKey(r))) duplicateCount++
  }
  return {
    totalToImport: imported.length,
    duplicateCount,
    newCount: imported.length - duplicateCount,
    existingCount: existing.length,
    tombstoneCount: tombstones.length,
  }
}

export function applyImport(
  existing: Rule[],
  imported: Rule[],
  strategy: ImportStrategy,
): Rule[] {
  if (strategy === 'replace') return [...imported]
  // merge: keep existing rules for any (type, signal) collision so we don't
  // wipe out their matchCount / lastUsedAt / user edits.
  const existingKeys = new Set(existing.map(ruleKey))
  const toAdd = imported.filter((r) => !existingKeys.has(ruleKey(r)))
  return [...existing, ...toAdd]
}
