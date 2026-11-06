// Compact wire codec for chrome.storage.sync (schemaVersion 3 + 4).
//
// Why: the dominant cost in a synced rule library is the folder reference.
// A Graph folder Id is ~125 bytes and every rule carries one; with 174
// rules pointing at ~40 distinct folders, `targetFolderId` + `targetFolderPath`
// is ~44% of the whole payload, almost all of it duplicated. Deduplicating
// folders into a dictionary + storing rules as compact rows (short keys,
// omitted defaults) roughly halves the bytes — the difference between fitting
// the 100 KB sync quota and blowing past it.
//
// v4 (2026-11-05) on top of v3:
//   - `ea` editedAt per row (sync merges per rule by it — see types.ts).
//   - compound signals as [kindCode, value] pairs instead of the escaped
//     `{"conditions":[{"type":…,"value":…}]}` JSON string (~82 B/row; compounds
//     are ~73% of a real library). Same for compound tombstone keys.
//   - canonical ISO timestamps as epoch-ms numbers (~13 B each).
//   - disabled-rule STUBS ([id, editedAt, reason, disabledAt]) so a disable
//     propagates even when the budget can't fit the rule's full row.
//   - compact tombstone rows [type, signalNorm, path, deletedAt].
//   Every compaction is applied only when it round-trips BYTE-EXACTLY, so a
//   decoded rule compares equal to the encoded one (no spurious edits).
//
// This module is PURE (Rule[] ↔ wire objects). No storage I/O — the sync
// engine chunks the output and writes it. All decode paths are defensive and
// SKIP what they cannot represent faithfully: a row with an unknown type /
// source / compound-kind code, or whose folder reference has no dictionary
// entry, is dropped rather than coerced (coercion used to turn unknown types
// into subject_keyword rules and dangling folders into EMPTY targets that
// replaced good local copies and re-spread on the next push). An unknown
// reason code keeps the rule and drops only the reason.

import type { Rule, RuleSource, RuleTombstone, RuleType } from './types'

// --- Enum codes -------------------------------------------------------------
// APPEND-ONLY. A code is a wire value shared across machines; never renumber
// an existing entry. New RuleType / RuleSource → append with the next integer.
// (Cross-version safety is also covered by downgrade protection in the engine,
// but keeping these append-only means even a hand-rolled reader stays correct.)

const TYPE_TO_CODE: Record<RuleType, number> = {
  case_code: 0,
  domain: 1,
  compound: 2,
  subject_keyword: 3,
  sender: 4,
}
const CODE_TO_TYPE: readonly RuleType[] = [
  'case_code',
  'domain',
  'compound',
  'subject_keyword',
  'sender',
]

const SOURCE_TO_CODE: Record<RuleSource, number> = {
  auto_scan: 0,
  ai_confirmed: 1,
  ai_overridden: 2,
  user_manual: 3,
}
const CODE_TO_SOURCE: readonly RuleSource[] = [
  'auto_scan',
  'ai_confirmed',
  'ai_overridden',
  'user_manual',
]

export type DisabledReason = NonNullable<Rule['autoDisabledReason']>
// RELEASE DISCIPLINE: a new reason code must ship in the SAME release as
// its first writer. A build that writes code N while another machine's
// build lacks entry N would silently strip the reason on decode and
// re-push the rule untagged — for case_closed that would permanently
// break the Phase-2 purge convergence.
const REASON_TO_CODE: Record<DisabledReason, number> = {
  stale: 0,
  'high-error-rate': 1,
  legacy_token: 2,
  case_closed: 3,
  override: 4, // 2026-11-05: one-strike disable after a user re-target
  superseded: 5, // 2026-11-05: broad rule silenced by a narrower one
}
const CODE_TO_REASON: readonly DisabledReason[] = [
  'stale',
  'high-error-rate',
  'legacy_token',
  'case_closed',
  'override',
  'superseded',
]

type CompoundKind = 'domain' | 'subject_keyword' | 'sender'
const KIND_TO_CODE: Record<CompoundKind, number> = {
  domain: 0,
  subject_keyword: 1,
  sender: 2,
}
const CODE_TO_KIND: readonly CompoundKind[] = ['domain', 'subject_keyword', 'sender']

function codeLookup<T>(table: readonly T[], code: unknown): T | undefined {
  return typeof code === 'number' && Number.isInteger(code) && code >= 0 && code < table.length
    ? table[code]
    : undefined
}

// --- Wire shapes ------------------------------------------------------------

/** One distinct target folder, addressed by its own stable key `x`. */
export type FolderDictEntry = {
  /**
   * Stable reference this folder is addressed by (row.f) — a hash of
   * (id,path), NOT an array position. Positional indices would renumber every
   * lexicographically-later folder whenever a new folder appears, rewriting
   * most rows' `f` and forcing chrome.storage.sync to re-set nearly all row
   * chunks (burning the per-hour write quota). An intrinsic hash keeps every
   * existing row byte-identical when an unrelated folder is added.
   */
  x: string
  /** targetFolderId (the expensive part — ~125 bytes of Graph id). */
  id: string
  /** targetFolderPath. */
  p: string
}

/** ISO timestamp, or (v4) its epoch-ms when that round-trips exactly. */
export type CompactTime = string | number

/** v4 compound conditions: [kindCode, value] per condition, in order. */
export type CompactCompound = Array<[number, string]>

/**
 * A rule as stored on the wire. Short keys; default-valued fields omitted
 * (enabled=true, matchCount=0, overrideCount=0, and all optionals). `f`
 * indexes into the folder dictionary.
 */
export type CompactRuleRow = {
  i: string // id
  t: number // type code
  s: string | CompactCompound // signal (v4: compact pairs for exact-round-trip compounds)
  f: string // folder reference (FolderDictEntry.x)
  c: number // confidence
  r: number // source code (r = "role/source"; s is taken by signal)
  ca: CompactTime // createdAt
  m?: number // matchCount (omit if 0)
  o?: number // overrideCount (omit if 0 / undefined)
  e?: 0 // enabled: omit if true, store 0 if false
  lu?: CompactTime // lastUsedAt (omit if absent)
  da?: CompactTime // autoDisabledAt (omit if absent)
  dr?: number // autoDisabledReason code (omit if absent)
  ea?: number // v4: editedAt epoch ms (omit if absent)
  // v4 additive (2026-11-06): statsResetAt epoch ms (omit if absent). A
  // v1.0.8 reader ignores unknown row keys — it just merges counters by max.
  sr?: number
}

/**
 * v4 disabled-state stub: `[id, editedAt|0, reasonCode|-1, autoDisabledAt]`
 * (trailing entries omitted when absent). Sent for a disabled rule whose full
 * row did not fit the byte budget, so the disable still propagates — before
 * v4 the budget dropped disabled rules first and the other machine's enabled
 * copy then re-enabled the rule everywhere. A stub never creates a rule; it
 * only updates one the receiver already has (by id).
 */
export type DisabledStub = [string, number, number?, CompactTime?]

/** One entry of the v4 row chunks: a full row or a disabled stub. */
export type WireRuleEntry = CompactRuleRow | DisabledStub

/**
 * v4 tombstone row: [typeCode, signalNorm (compact for compounds), path,
 * deletedAt, spellingHash?]. The 5th entry (2026-11-06, v1.0.9) is
 * RuleTombstone.spellingHash: `1` when it is the fingerprint of signalNorm
 * itself (2 bytes — it still marks the row as written by a build that records
 * it), else the fingerprint. A v1.0.8 decoder reads [0..3] only.
 */
export type CompactTombstone =
  | [number, string | CompactCompound, string, number]
  | [number, string | CompactCompound, string, number, 1 | string]

export type EncodedRules = {
  folders: FolderDictEntry[]
  rows: WireRuleEntry[]
}

/** A decoded disabled stub (see DisabledStub). */
export type DecodedStub = {
  id: string
  editedAt?: number
  autoDisabledReason?: DisabledReason
  autoDisabledAt?: string
}

export type DecodedRulePayload = {
  rules: Rule[]
  stubs: DecodedStub[]
  /** Rows dropped because they could not be decoded faithfully. */
  skipped: number
}

// --- helpers ----------------------------------------------------------------

/**
 * FNV-1a 64-bit hash → base36. Intrinsic to the folder key, so a folder's
 * reference is stable across pushes regardless of what other folders exist.
 */
function hashFolderKey(key: string): string {
  let h = 0xcbf29ce484222325n
  const mask = 0xffffffffffffffffn
  for (let i = 0; i < key.length; i++) {
    h ^= BigInt(key.charCodeAt(i))
    h = (h * 0x100000001b3n) & mask
  }
  return h.toString(36)
}

/**
 * Fast 53-bit content hash (cyrb53) → base36. Used for the push's
 * content fingerprint (meta.contentHash): identical payloads must hash
 * identically within one build so a no-op push can be skipped.
 */
export function contentHash(s: string): string {
  let h1 = 0xdeadbeef
  let h2 = 0x41c6ce57
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i)
    h1 = Math.imul(h1 ^ ch, 2654435761)
    h2 = Math.imul(h2 ^ ch, 1597334677)
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507)
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909)
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507)
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909)
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36)
}

function folderKey(id: string, p: string): string {
  // NUL separator can't appear in an id or path, so (id,p) → key is injective.
  return `${id}\u0000${p}`
}

function isoOf(ms: number): string | undefined {
  if (!Number.isFinite(ms) || Math.abs(ms) > 8.64e15) return undefined
  return new Date(ms).toISOString()
}

/** ISO → epoch-ms only when `new Date(ms).toISOString()` reproduces it exactly. */
export function compactTime(iso: string): CompactTime {
  const ms = Date.parse(iso)
  return Number.isFinite(ms) && isoOf(ms) === iso ? ms : iso
}

/** Inverse of compactTime; undefined for anything that isn't a time. */
export function expandTime(v: unknown): string | undefined {
  if (typeof v === 'string') return v
  if (typeof v === 'number') return isoOf(v)
  return undefined
}

function isPlainCondition(c: unknown): c is { type: CompoundKind; value: string } {
  if (!c || typeof c !== 'object' || Array.isArray(c)) return false
  const o = c as Record<string, unknown>
  return (
    Object.keys(o).length === 2 &&
    typeof o.type === 'string' &&
    Object.prototype.hasOwnProperty.call(KIND_TO_CODE, o.type) &&
    typeof o.value === 'string'
  )
}

function expandConditions(c: unknown): Array<{ type: CompoundKind; value: string }> | null {
  if (!Array.isArray(c)) return null
  const out: Array<{ type: CompoundKind; value: string }> = []
  for (const e of c) {
    if (!Array.isArray(e) || e.length !== 2) return null
    const kind = codeLookup(CODE_TO_KIND, e[0])
    if (!kind || typeof e[1] !== 'string') return null
    out.push({ type: kind, value: e[1] })
  }
  return out
}

/** Rule.signal of a compound rule (`{"conditions":[…]}`) → compact pairs, or null. */
export function compactCompoundSignal(signal: string): CompactCompound | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(signal)
  } catch {
    return null
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null
  const o = parsed as Record<string, unknown>
  if (Object.keys(o).length !== 1 || !Array.isArray(o.conditions)) return null
  if (!o.conditions.every(isPlainCondition)) return null
  const out: CompactCompound = (o.conditions as Array<{ type: CompoundKind; value: string }>).map(
    (c) => [KIND_TO_CODE[c.type], c.value],
  )
  // Only when the expansion reproduces the stored string byte-for-byte.
  return expandCompoundSignal(out) === signal ? out : null
}

export function expandCompoundSignal(c: unknown): string | null {
  const conds = expandConditions(c)
  return conds ? JSON.stringify({ conditions: conds }) : null
}

/** normalizeSignal('compound', …) output (a bare `[{type,value}]` array) → compact. */
export function compactCompoundNorm(norm: string): CompactCompound | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(norm)
  } catch {
    return null
  }
  if (!Array.isArray(parsed) || !parsed.every(isPlainCondition)) return null
  const out: CompactCompound = (parsed as Array<{ type: CompoundKind; value: string }>).map(
    (c) => [KIND_TO_CODE[c.type], c.value],
  )
  return expandCompoundNorm(out) === norm ? out : null
}

export function expandCompoundNorm(c: unknown): string | null {
  const conds = expandConditions(c)
  return conds ? JSON.stringify(conds) : null
}

// --- rule encode ------------------------------------------------------------

/**
 * Folder dictionary for `rules`. Deterministic: folders are sorted by
 * (id,path) and assigned hash references in that order, so the same library
 * always produces the same dictionary (stable chunk membership → no
 * write-quota churn).
 */
export function buildFolderDict(rules: Rule[]): {
  folders: FolderDictEntry[]
  refOf: (r: Pick<Rule, 'targetFolderId' | 'targetFolderPath'>) => string
} {
  const seen = new Map<string, { id: string; p: string }>()
  for (const r of rules) {
    const k = folderKey(r.targetFolderId, r.targetFolderPath)
    if (!seen.has(k)) seen.set(k, { id: r.targetFolderId, p: r.targetFolderPath })
  }
  const sortedKeys = [...seen.keys()].sort((a, b) => a.localeCompare(b))
  const refByKey = new Map<string, string>()
  const usedX = new Map<string, string>() // x → folderKey, for collision detection
  const folders: FolderDictEntry[] = []
  for (const k of sortedKeys) {
    let x = hashFolderKey(k)
    // Collision guard: two DISTINCT folders sharing a reference would misroute
    // a rule, and a legal-filing tool must never silently misfile. At 64 bits
    // this is astronomically rare; when it does happen, probe deterministically
    // (sorted-key order) so the assignment stays reproducible across pushes.
    let probe = 0
    while (usedX.has(x) && usedX.get(x) !== k) x = `${hashFolderKey(k)}.${++probe}`
    usedX.set(x, k)
    refByKey.set(k, x)
    const f = seen.get(k)!
    folders.push({ x, id: f.id, p: f.p })
  }
  return {
    folders,
    refOf: (r) => refByKey.get(folderKey(r.targetFolderId, r.targetFolderPath)) ?? '',
  }
}

/** Full v4 row for `r`, or null when its type/source has no wire code. */
export function encodeRuleRow(r: Rule, folderRef: string): CompactRuleRow | null {
  const t = TYPE_TO_CODE[r.type]
  const src = SOURCE_TO_CODE[r.source]
  if (t === undefined || src === undefined) return null
  const row: CompactRuleRow = {
    i: r.id,
    t,
    s: (r.type === 'compound' && compactCompoundSignal(r.signal)) || r.signal,
    f: folderRef,
    c: r.confidence,
    r: src,
    ca: compactTime(r.createdAt),
  }
  if (r.matchCount) row.m = r.matchCount
  if (r.overrideCount) row.o = r.overrideCount
  if (r.enabled === false) row.e = 0
  if (r.lastUsedAt) row.lu = compactTime(r.lastUsedAt)
  if (r.autoDisabledAt) row.da = compactTime(r.autoDisabledAt)
  if (r.autoDisabledReason) {
    const code = REASON_TO_CODE[r.autoDisabledReason]
    if (code !== undefined) row.dr = code
  }
  if (typeof r.editedAt === 'number' && r.editedAt > 0) row.ea = r.editedAt
  if (typeof r.statsResetAt === 'number' && Number.isFinite(r.statsResetAt) && r.statsResetAt > 0) {
    row.sr = r.statsResetAt
  }
  return row
}

/** Disabled-state stub for a disabled rule (see DisabledStub). */
export function encodeDisabledStub(r: Rule): DisabledStub {
  const out: Array<string | number> = [r.id, typeof r.editedAt === 'number' ? r.editedAt : 0]
  const code = r.autoDisabledReason ? REASON_TO_CODE[r.autoDisabledReason] : undefined
  if (code !== undefined || r.autoDisabledAt) {
    out.push(code ?? -1)
    if (r.autoDisabledAt) out.push(compactTime(r.autoDisabledAt))
  }
  return out as DisabledStub
}

/**
 * Rule[] → { folders, rows } with a full row per rule (rows follow input
 * order; rules with no wire code are omitted). The sync engine's budget
 * planner (sync-chunks.ts) builds the real push payload — this is the
 * straightforward whole-library encoding.
 */
export function encodeRulesForSync(rules: Rule[]): EncodedRules {
  const { folders, refOf } = buildFolderDict(rules)
  const rows: WireRuleEntry[] = []
  for (const r of rules) {
    const row = encodeRuleRow(r, refOf(r))
    if (row) rows.push(row)
  }
  return { folders, rows }
}

// --- rule decode ------------------------------------------------------------

function decodeFullRow(
  row: Record<string, unknown>,
  folderMap: Map<string, FolderDictEntry>,
): Rule | null {
  if (typeof row.i !== 'string' || row.i === '') return null
  const type = codeLookup(CODE_TO_TYPE, row.t)
  const source = codeLookup(CODE_TO_SOURCE, row.r)
  if (!type || !source) return null // unknown code → skip, never coerce
  let signal: string
  if (typeof row.s === 'string') signal = row.s
  else if (type === 'compound') {
    const s = expandCompoundSignal(row.s)
    if (s === null) return null
    signal = s
  } else return null
  // Item 4 (2026-11-05): an unresolved folder reference means the dictionary
  // chunk is missing or from another push. Decoding it to an EMPTY target
  // replaced the good local copy and re-spread on the next push — skip.
  const folder = typeof row.f === 'string' ? folderMap.get(row.f) : undefined
  if (!folder) return null
  const rule: Rule = {
    id: row.i,
    type,
    signal,
    targetFolderId: folder.id,
    targetFolderPath: folder.p,
    confidence: typeof row.c === 'number' && Number.isFinite(row.c) ? row.c : 0.7,
    matchCount: typeof row.m === 'number' && row.m > 0 ? row.m : 0,
    enabled: row.e !== 0,
    createdAt: expandTime(row.ca) ?? '',
    source,
  }
  if (typeof row.o === 'number' && row.o > 0) rule.overrideCount = row.o
  const lu = expandTime(row.lu)
  if (lu !== undefined) rule.lastUsedAt = lu
  const da = expandTime(row.da)
  if (da !== undefined) rule.autoDisabledAt = da
  const reason = codeLookup(CODE_TO_REASON, row.dr)
  if (reason) rule.autoDisabledReason = reason // unknown code → drop the reason only
  if (typeof row.ea === 'number' && Number.isFinite(row.ea) && row.ea > 0) rule.editedAt = row.ea
  if (typeof row.sr === 'number' && Number.isFinite(row.sr) && row.sr > 0) rule.statsResetAt = row.sr
  return rule
}

function decodeStub(row: unknown[]): DecodedStub | null {
  if (typeof row[0] !== 'string' || row[0] === '') return null
  const stub: DecodedStub = { id: row[0] }
  if (typeof row[1] === 'number' && Number.isFinite(row[1]) && row[1] > 0) stub.editedAt = row[1]
  const reason = codeLookup(CODE_TO_REASON, row[2])
  if (reason) stub.autoDisabledReason = reason
  const da = expandTime(row[3])
  if (da !== undefined) stub.autoDisabledAt = da
  return stub
}

/**
 * { folders, rows } → rules + disabled stubs. Reads v3 rows (string signals,
 * ISO times) and v4 rows alike. Fully defensive — this parses data that may
 * have been written by a different (newer/older/buggy) build:
 *   - rows without a string id, with an unknown type / source / compound-kind
 *     code, or with an unresolved folder reference are SKIPPED (counted).
 *   - an unknown autoDisabledReason code keeps the rule, minus the reason.
 */
export function decodeRulePayload(enc: { folders?: unknown; rows?: unknown }): DecodedRulePayload {
  const rows = Array.isArray(enc?.rows) ? (enc.rows as unknown[]) : []
  const folderMap = new Map<string, FolderDictEntry>()
  if (Array.isArray(enc?.folders)) {
    for (const f of enc.folders as unknown[]) {
      const e = f as FolderDictEntry | null
      if (e && typeof e.x === 'string' && typeof e.id === 'string' && typeof e.p === 'string') {
        folderMap.set(e.x, e)
      }
    }
  }
  const out: DecodedRulePayload = { rules: [], stubs: [], skipped: 0 }
  for (const row of rows) {
    if (Array.isArray(row)) {
      const stub = decodeStub(row)
      if (stub) out.stubs.push(stub)
      else out.skipped++
      continue
    }
    const rule =
      row && typeof row === 'object' ? decodeFullRow(row as Record<string, unknown>, folderMap) : null
    if (rule) out.rules.push(rule)
    else out.skipped++
  }
  return out
}

/** Back-compat wrapper: full rules only (stubs / skip counts dropped). */
export function decodeRulesFromSync(enc: { folders?: unknown; rows?: unknown }): Rule[] {
  return decodeRulePayload(enc).rules
}

// --- tombstones -------------------------------------------------------------

/**
 * RuleTombstone.spellingHash of a spelling in rules.ts legacySignalNorm form
 * (2026-11-06, S1): its contentHash (53-bit, base36, ~11 chars).
 */
export function spellingFingerprint(spelling: string): string {
  return contentHash(spelling)
}

export function encodeTombstone(t: RuleTombstone): CompactTombstone | null {
  const code = TYPE_TO_CODE[t.type]
  if (code === undefined || typeof t.signalNorm !== 'string') return null
  const s = t.type === 'compound' ? (compactCompoundNorm(t.signalNorm) ?? t.signalNorm) : t.signalNorm
  if (typeof t.spellingHash !== 'string') return [code, s, t.targetFolderPath, t.deletedAt]
  const own = t.spellingHash === spellingFingerprint(t.signalNorm)
  return [code, s, t.targetFolderPath, t.deletedAt, own ? 1 : t.spellingHash]
}

export function decodeTombstone(row: unknown): RuleTombstone | null {
  if (!Array.isArray(row) || row.length < 4) return null
  const type = codeLookup(CODE_TO_TYPE, row[0])
  if (!type) return null
  let signalNorm: string
  if (typeof row[1] === 'string') signalNorm = row[1]
  else if (type === 'compound') {
    const s = expandCompoundNorm(row[1])
    if (s === null) return null
    signalNorm = s
  } else return null
  if (typeof row[2] !== 'string') return null
  if (typeof row[3] !== 'number' || !Number.isFinite(row[3])) return null
  const t: RuleTombstone = { type, signalNorm, targetFolderPath: row[2], deletedAt: row[3] }
  // Optional spelling fingerprint (2026-11-06). Anything else there just
  // means "unknown" — never a reason to drop the deletion itself.
  if (row[4] === 1) t.spellingHash = spellingFingerprint(signalNorm)
  else if (typeof row[4] === 'string' && row[4] !== '') t.spellingHash = row[4]
  return t
}

// --- legacy (v1–v3) full-object readers ---------------------------------------

const RULE_TYPES: ReadonlySet<string> = new Set(CODE_TO_TYPE)
const RULE_SOURCES: ReadonlySet<string> = new Set(CODE_TO_SOURCE)
const REASONS: ReadonlySet<string> = new Set(CODE_TO_REASON)

/** v1–v3 tombstone object, validated. */
export function sanitizeLegacyTombstone(raw: unknown): RuleTombstone | null {
  if (!raw || typeof raw !== 'object') return null
  const t = raw as Record<string, unknown>
  if (typeof t.type !== 'string' || !RULE_TYPES.has(t.type)) return null
  if (typeof t.signalNorm !== 'string' || typeof t.targetFolderPath !== 'string') return null
  if (typeof t.deletedAt !== 'number' || !Number.isFinite(t.deletedAt)) return null
  return {
    type: t.type as RuleType,
    signalNorm: t.signalNorm,
    targetFolderPath: t.targetFolderPath,
    deletedAt: t.deletedAt,
    // Backups taken by this build carry it (2026-11-06); v1–v3 clouds never do.
    ...(typeof t.spellingHash === 'string' && t.spellingHash !== '' ? { spellingHash: t.spellingHash } : {}),
  }
}

/** v1/v2 full Rule object (syncRules_N chunks), validated field by field. */
export function sanitizeLegacyRule(raw: unknown): Rule | null {
  if (!raw || typeof raw !== 'object') return null
  const r = raw as Record<string, unknown>
  if (typeof r.id !== 'string' || r.id === '') return null
  if (typeof r.type !== 'string' || !RULE_TYPES.has(r.type)) return null
  if (typeof r.source !== 'string' || !RULE_SOURCES.has(r.source)) return null
  if (typeof r.signal !== 'string') return null
  if (typeof r.targetFolderId !== 'string' || typeof r.targetFolderPath !== 'string') return null
  const rule: Rule = {
    id: r.id,
    type: r.type as RuleType,
    signal: r.signal,
    targetFolderId: r.targetFolderId,
    targetFolderPath: r.targetFolderPath,
    confidence: typeof r.confidence === 'number' && Number.isFinite(r.confidence) ? r.confidence : 0.7,
    matchCount: typeof r.matchCount === 'number' && r.matchCount > 0 ? r.matchCount : 0,
    enabled: r.enabled !== false,
    createdAt: typeof r.createdAt === 'string' ? r.createdAt : '',
    source: r.source as RuleSource,
  }
  if (typeof r.overrideCount === 'number' && r.overrideCount > 0) rule.overrideCount = r.overrideCount
  if (typeof r.lastUsedAt === 'string') rule.lastUsedAt = r.lastUsedAt
  if (typeof r.autoDisabledAt === 'string') rule.autoDisabledAt = r.autoDisabledAt
  if (typeof r.autoDisabledReason === 'string' && REASONS.has(r.autoDisabledReason)) {
    rule.autoDisabledReason = r.autoDisabledReason as DisabledReason
  }
  return rule
}

// --- active-day ledger (2026-11-06) -------------------------------------------
//
// The stale sweep's clock (storage.ts activeDays) travels in the v4 settings
// item as an optional `d` field — a v1.0.8 reader only looks at `p` / `v`.
// Compact form: UTC day numbers ascending, the first absolute and the rest as
// deltas (`[20398,1,1,3,1]`) — ~2 bytes a day, so the full 400-day ledger
// costs about 1 KB.

const DAY_MS = 86_400_000
const DAY_RE = /^\d{4}-\d{2}-\d{2}$/
/** Day numbers outside this range are junk (1970 … ~2243). */
const MAX_DAY_NUMBER = 100_000

function dayNumber(day: string): number | undefined {
  if (typeof day !== 'string' || !DAY_RE.test(day)) return undefined
  const ms = Date.parse(`${day}T00:00:00.000Z`)
  if (!Number.isFinite(ms)) return undefined
  const n = Math.round(ms / DAY_MS)
  return n >= 0 && n <= MAX_DAY_NUMBER ? n : undefined
}

/** 'YYYY-MM-DD' days → compact wire form, newest `cap` kept. */
export function encodeActiveDays(days: readonly string[], cap: number): number[] {
  const nums = [...new Set(days.map(dayNumber).filter((n): n is number => n !== undefined))]
    .sort((a, b) => a - b)
    .slice(-Math.max(0, cap))
  const out: number[] = []
  for (let i = 0; i < nums.length; i++) out.push(i === 0 ? nums[i]! : nums[i]! - nums[i - 1]!)
  return out
}

/**
 * Inverse of encodeActiveDays — defensive: anything malformed yields [] (the
 * ledger only feeds the stale-sweep clock; a partial decode is never worth
 * the risk), and days after `latestMs` (a peer whose clock runs far ahead)
 * are dropped.
 */
export function decodeActiveDays(raw: unknown, latestMs: number = Date.now() + 2 * DAY_MS): string[] {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > 5000) return []
  const out: string[] = []
  let n = 0
  for (let i = 0; i < raw.length; i++) {
    const v = raw[i]
    if (typeof v !== 'number' || !Number.isInteger(v) || (i === 0 ? v < 0 : v < 1)) return []
    n = i === 0 ? v : n + v
    if (n > MAX_DAY_NUMBER) return []
    if (n * DAY_MS > latestMs) break
    out.push(new Date(n * DAY_MS).toISOString().slice(0, 10))
  }
  return out
}
