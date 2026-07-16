// Compact wire codec for chrome.storage.sync (schemaVersion 3).
//
// Why: the dominant cost in a synced rule library is the folder reference.
// A Graph folder Id is ~125 bytes and every rule carries one; with 174
// rules pointing at ~40 distinct folders, `targetFolderId` + `targetFolderPath`
// is ~44% of the whole payload, almost all of it duplicated. Deduplicating
// folders into a dictionary + storing rules as compact rows (short keys,
// omitted defaults) roughly halves the bytes — the difference between fitting
// the 100 KB sync quota and blowing past it.
//
// This module is PURE (Rule[] ↔ wire objects). No storage I/O — the sync
// engine chunks the output and writes it. All decode paths are defensive:
// malformed rows are skipped, unknown enum codes fall back to safe defaults,
// and a missing folder resolves to empty strings rather than corrupting
// neighbouring rules (folders are addressed by an explicit index, never by
// array position, so a dropped folder chunk can't shift the mapping).

import type { Rule, RuleSource, RuleType } from './types'

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

type DisabledReason = NonNullable<Rule['autoDisabledReason']>
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
}
const CODE_TO_REASON: readonly DisabledReason[] = [
  'stale',
  'high-error-rate',
  'legacy_token',
  'case_closed',
]

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
 * A rule as stored on the wire. Short keys; default-valued fields omitted
 * (enabled=true, matchCount=0, overrideCount=0, and all optionals). `f`
 * indexes into the folder dictionary.
 */
export type CompactRuleRow = {
  i: string // id
  t: number // type code
  s: string // signal
  f: string // folder reference (FolderDictEntry.x)
  c: number // confidence
  r: number // source code (r = "role/source"; s is taken by signal)
  ca: string // createdAt (ISO)
  m?: number // matchCount (omit if 0)
  o?: number // overrideCount (omit if 0 / undefined)
  e?: 0 // enabled: omit if true, store 0 if false
  lu?: string // lastUsedAt (omit if absent)
  da?: string // autoDisabledAt (omit if absent)
  dr?: number // autoDisabledReason code (omit if absent)
}

export type EncodedRules = {
  folders: FolderDictEntry[]
  rows: CompactRuleRow[]
}

function folderKey(id: string, p: string): string {
  // NUL separator can't appear in an id or path, so (id,p) → key is injective.
  return `${id}\u0000${p}`
}

/**
 * Rule[] → { folders, rows }. Deterministic: folders are sorted by
 * (id,path) and assigned indices in that order, so the same library always
 * produces the same dictionary (stable chunk membership → no write-quota
 * churn). Rows preserve the input order of `rules`.
 */
export function encodeRulesForSync(rules: Rule[]): EncodedRules {
  const seen = new Map<string, { id: string; p: string }>()
  for (const r of rules) {
    const k = folderKey(r.targetFolderId, r.targetFolderPath)
    if (!seen.has(k)) seen.set(k, { id: r.targetFolderId, p: r.targetFolderPath })
  }
  const sortedKeys = [...seen.keys()].sort((a, b) => a.localeCompare(b))
  const indexOf = new Map<string, string>()
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
    indexOf.set(k, x)
    const f = seen.get(k)!
    folders.push({ x, id: f.id, p: f.p })
  }

  const rows: CompactRuleRow[] = rules.map((r) => {
    const row: CompactRuleRow = {
      i: r.id,
      t: TYPE_TO_CODE[r.type] ?? TYPE_TO_CODE.subject_keyword,
      s: r.signal,
      f: indexOf.get(folderKey(r.targetFolderId, r.targetFolderPath)) ?? '',
      c: r.confidence,
      r: SOURCE_TO_CODE[r.source] ?? SOURCE_TO_CODE.ai_confirmed,
      ca: r.createdAt,
    }
    if (r.matchCount) row.m = r.matchCount
    if (r.overrideCount) row.o = r.overrideCount
    if (r.enabled === false) row.e = 0
    if (r.lastUsedAt) row.lu = r.lastUsedAt
    if (r.autoDisabledAt) row.da = r.autoDisabledAt
    if (r.autoDisabledReason) row.dr = REASON_TO_CODE[r.autoDisabledReason]
    return row
  })

  return { folders, rows }
}

/**
 * { folders, rows } → Rule[]. Fully defensive — this parses data that may
 * have been written by a different (newer/older/buggy) build, or arrive with
 * chunks missing:
 *   - rows without a string id are skipped (corrupt).
 *   - unknown type / source codes fall back to safe defaults.
 *   - a folder index with no dictionary entry resolves to empty strings
 *     (reconcileRulesAgainstTree will flag it orphaned on the next scan)
 *     rather than throwing or mis-pointing at another folder.
 */
export function decodeRulesFromSync(enc: EncodedRules): Rule[] {
  const rows = Array.isArray(enc?.rows) ? enc.rows : []
  const folderMap = new Map<string, FolderDictEntry>()
  if (Array.isArray(enc?.folders)) {
    for (const f of enc.folders) {
      if (f && typeof f.x === 'string' && typeof f.id === 'string' && typeof f.p === 'string') {
        folderMap.set(f.x, f)
      }
    }
  }

  const out: Rule[] = []
  for (const row of rows) {
    if (!row || typeof row.i !== 'string') continue
    const folder = folderMap.get(row.f)
    const rule: Rule = {
      id: row.i,
      type: CODE_TO_TYPE[row.t] ?? 'subject_keyword',
      signal: typeof row.s === 'string' ? row.s : '',
      targetFolderId: folder?.id ?? '',
      targetFolderPath: folder?.p ?? '',
      confidence: typeof row.c === 'number' ? row.c : 0.7,
      matchCount: typeof row.m === 'number' ? row.m : 0,
      enabled: row.e !== 0,
      createdAt: typeof row.ca === 'string' ? row.ca : '',
      source: CODE_TO_SOURCE[row.r] ?? 'ai_confirmed',
    }
    if (typeof row.o === 'number' && row.o > 0) rule.overrideCount = row.o
    if (typeof row.lu === 'string') rule.lastUsedAt = row.lu
    if (typeof row.da === 'string') rule.autoDisabledAt = row.da
    if (typeof row.dr === 'number' && CODE_TO_REASON[row.dr]) {
      rule.autoDisabledReason = CODE_TO_REASON[row.dr]
    }
    out.push(rule)
  }
  return out
}
