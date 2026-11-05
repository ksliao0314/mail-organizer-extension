// Chunking helpers + byte-budget planner for chrome.storage.sync writes.
//
// Why chunking: chrome.storage.sync caps each item at 8 KB and the
// entire bucket at 100 KB. Naive `set({ rules: bigArray })` blows up on
// any non-trivial library. We split rules + tombstones into chunks of
// ~6 KB each (well under the per-item ceiling, leaving headroom for
// future schema growth).
//
// All operations are pure JSON serialisation — no storage I/O. The
// sync engine module (`background/sync-engine.ts`) wires these into
// chrome.storage.sync calls.

import {
  CHUNK_BYTE_TARGET,
  FOLDER_ACTIVITY_SYNC_CAP,
  MAX_FOLDER_DICT_CHUNKS,
  MAX_RULE_CHUNKS,
  MAX_TOMBSTONE_CHUNKS,
  TOMBSTONE_SYNC_CAP,
} from './constants'
import {
  buildFolderDict,
  encodeDisabledStub,
  encodeRuleRow,
  encodeTombstone,
  type CompactRuleRow,
  type CompactTombstone,
  type DisabledStub,
  type FolderDictEntry,
  type WireRuleEntry,
} from './sync-codec'
import type { FolderActivity, Rule, RuleSource, RuleTombstone } from './types'

/** Sync envelope. Stored under one `chrome.storage.sync` key per chunk. */
export type SyncChunk<T> = {
  /** 0-indexed within its category. Used by callers to know how many to fetch. */
  index: number
  items: T[]
  /**
   * v4 (2026-11-05): pushId of the push that wrote this chunk. A pull only
   * accepts a snapshot whose every chunk carries the meta's pushId — chunks of
   * one push can reach another machine before / after its meta (torn read).
   */
  p?: string
}

const _utf8Encoder = new TextEncoder()

/**
 * UTF-8 byte length of a string — the unit chrome.storage.sync actually
 * enforces its 8 KB-per-item / 100 KB-total quota in.
 *
 * The old code used `JSON.stringify(x).length`, i.e. the UTF-16 code-unit
 * count. For ASCII that equals the byte count, but a CJK character is 1
 * code unit and 3 bytes. A rule library full of Chinese folder paths and
 * subject keywords therefore under-counted by up to 3×, so a chunk the
 * chunker believed was 6 KB could be 14 KB in storage → the
 * `kQuotaBytesPerItem quota exceeded` crash. Always measure in real bytes.
 */
export function utf8ByteLength(s: string): number {
  return _utf8Encoder.encode(s).length
}

/** UTF-8 byte size of a value once JSON-serialised (how sync bills it). */
export function jsonByteSize(value: unknown): number {
  return utf8ByteLength(JSON.stringify(value))
}

/** Bytes one chrome.storage.sync item costs: UTF-8 key + JSON-serialised value. */
export function writeBytes(key: string, value: unknown): number {
  return utf8ByteLength(key) + jsonByteSize(value)
}

// CHUNK_BYTE_TARGET re-exported for backward compat with existing
// imports; new code should import from '@/shared/constants'.
export { CHUNK_BYTE_TARGET } from './constants'

/**
 * Source filter for rules eligible for sync. Excludes:
 *   - auto_scan: cheaper to re-derive per machine via initial scan; would
 *     also balloon the synced set well past the quota.
 *   - orphaned: target folder gone, not useful on another machine either.
 *
 * Auto-disabled rules ARE included — the disable state IS the sync intent
 * (so the other machine doesn't re-create the same rule).
 */
const SYNCABLE_SOURCES: ReadonlySet<RuleSource> = new Set([
  'user_manual',
  'ai_confirmed',
  'ai_overridden',
])

export function shouldSyncRule(rule: Rule): boolean {
  if (rule.orphaned) return false
  if (!SYNCABLE_SOURCES.has(rule.source)) return false
  return true
}

const SOURCE_AUTHORITY: Record<RuleSource, number> = {
  user_manual: 3, // hand-made — highest intent
  ai_overridden: 2, // user corrected the AI — proven attention
  ai_confirmed: 1, // auto-learned
  auto_scan: 0, // never synced anyway
}

/** Locale-independent string order (localeCompare varies with the machine's ICU locale). */
function cmpStr(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

/**
 * Best-first ordering for the sync budget guard. When the library doesn't
 * fit the 100 KB quota, the lowest-value rules are the ones to drop from
 * the cloud (they stay in local storage — this only affects propagation).
 *
 * Priority, in order: enabled beats auto-disabled (a sleeping rule won't
 * fire on the other machine either — and since v4 its disable still
 * propagates through a stub); then source authority (hand-made > corrected >
 * auto); then track record (matchCount); then recency of use; then id. The
 * final id tie-break (2026-11-05) makes the order TOTAL: two machines holding
 * the same library must select the same budget-limited subset, otherwise each
 * one's push differs from the other's and neither is ever a no-op.
 */
export function compareRuleSyncPriority(a: Rule, b: Rule): number {
  if (a.enabled !== b.enabled) return a.enabled ? -1 : 1
  const auth = SOURCE_AUTHORITY[b.source] - SOURCE_AUTHORITY[a.source]
  if (auth !== 0) return auth
  const mc = (b.matchCount || 0) - (a.matchCount || 0)
  if (mc !== 0) return mc
  const aRecency = a.lastUsedAt ?? a.createdAt ?? ''
  const bRecency = b.lastUsedAt ?? b.createdAt ?? ''
  const rec = cmpStr(bRecency, aRecency)
  if (rec !== 0) return rec
  return cmpStr(a.id, b.id)
}

/**
 * Split rules into chunks targeting ~6KB each. Rules are sorted by id
 * before chunking so chunk membership is deterministic across runs —
 * we don't want unrelated rules to keep landing in different chunks on
 * every push, which would burn the per-hour write quota.
 *
 * If a single rule's serialised size somehow exceeds CHUNK_BYTE_TARGET
 * (very long signal / target path), it gets its own chunk. The hard
 * 8 KB per-item limit is still enforced by chrome.storage.sync; we
 * surface oversized rules as the engine catches the QUOTA error.
 *
 * (v1/v2 full-object layout — kept for reading legacy clouds in tests.)
 */
export function chunkRules(rules: Rule[]): SyncChunk<Rule>[] {
  const eligible = rules.filter(shouldSyncRule)
  return chunkArray(eligible, (r) => r.id)
}

// TOMBSTONE_SYNC_CAP re-exported from constants.ts for back-compat.
export { TOMBSTONE_SYNC_CAP } from './constants'

// FOLDER_ACTIVITY_SYNC_CAP re-exported from constants.ts for back-compat.
export { FOLDER_ACTIVITY_SYNC_CAP } from './constants'

export function chunkFolderActivity(
  activity: FolderActivity[],
  stamp?: string,
): SyncChunk<FolderActivity>[] {
  // Sort by lastActiveAt DESC to grab the most-recent entries (folderId
  // breaks ties so the selected set is deterministic).
  const recent = [...activity]
    .sort((a, b) => cmpStr(b.lastActiveAt, a.lastActiveAt) || cmpStr(a.folderId, b.folderId))
    .slice(0, FOLDER_ACTIVITY_SYNC_CAP)
  // Re-sort by folderId for stable chunk membership — without this, the
  // same set of entries could land in different chunks each push and burn
  // the per-hour write quota.
  return chunkArray(recent, (e) => e.folderId, stamp)
}

/** Id of a v4 row-chunk entry (full row or disabled stub). */
export function wireEntryId(e: WireRuleEntry): string {
  return Array.isArray(e) ? e[0] : (e as CompactRuleRow).i
}

/**
 * Chunk compact rule rows / stubs. Keyed by rule id for stable chunk
 * membership. NO source filtering / ordering happens here — the caller has
 * already selected the entries.
 */
export function chunkCompactRows(
  rows: WireRuleEntry[],
  stamp?: string,
): SyncChunk<WireRuleEntry>[] {
  return chunkArray(rows, wireEntryId, stamp)
}

/** Chunk the folder dictionary. Keyed by folder reference. */
export function chunkFolderDict(
  folders: FolderDictEntry[],
  stamp?: string,
): SyncChunk<FolderDictEntry>[] {
  return chunkArray(folders, (f) => String(f.x).padStart(6, '0'), stamp)
}

function compactTombstoneKey(t: CompactTombstone): string {
  return `${t[0]}::${typeof t[1] === 'string' ? t[1] : JSON.stringify(t[1])}::${t[2]}`
}

/** Chunk v4 compact tombstone rows (keyed by triple for stable membership). */
export function chunkCompactTombstones(
  rows: CompactTombstone[],
  stamp?: string,
): SyncChunk<CompactTombstone>[] {
  return chunkArray(rows, compactTombstoneKey, stamp)
}

/**
 * v3 full-object tombstone chunks (kept for writing legacy fixtures and for
 * readers of pre-v4 clouds). The v4 push selects tombstones with
 * planSyncPayload instead: newest first, inside a bounded byte share.
 */
export function chunkTombstones(tombstones: RuleTombstone[]): SyncChunk<RuleTombstone>[] {
  const recent = [...tombstones]
    .sort((a, b) => b.deletedAt - a.deletedAt)
    .slice(0, TOMBSTONE_SYNC_CAP)
  return chunkArray(recent, (t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`)
}

function chunkArray<T>(items: T[], keyFn: (item: T) => string, stamp?: string): SyncChunk<T>[] {
  const sorted = [...items].sort((a, b) => cmpStr(keyFn(a), keyFn(b)))
  const envelope = (index: number) =>
    jsonByteSize(stamp === undefined ? { index, items: [] } : { index, items: [], p: stamp })
  const make = (index: number, chunkItems: T[]): SyncChunk<T> =>
    stamp === undefined ? { index, items: chunkItems } : { index, items: chunkItems, p: stamp }
  const chunks: SyncChunk<T>[] = []
  let current: T[] = []
  // Envelope overhead, measured in real UTF-8 bytes (the unit sync bills in).
  let currentBytes = envelope(0)

  for (const item of sorted) {
    // +1 for the comma in the eventual array serialisation. Slight
    // over-estimate but safe.
    const itemSize = jsonByteSize(item) + 1
    if (current.length > 0 && currentBytes + itemSize > CHUNK_BYTE_TARGET) {
      chunks.push(make(chunks.length, current))
      current = []
      currentBytes = envelope(chunks.length)
    }
    current.push(item)
    currentBytes += itemSize
  }
  if (current.length > 0) {
    chunks.push(make(chunks.length, current))
  }
  return chunks
}

/**
 * Reverse of chunkRules — given an ordered list of chunks (any
 * order; we sort by `index`), concatenate items back into a single
 * array. Robust against missing / extra chunks: skips invalid entries
 * silently.
 */
export function unchunkArray<T>(chunks: SyncChunk<T>[]): T[] {
  // Filter null/invalid BEFORE sorting — sort comparator can't handle
  // missing index property without crashing.
  const valid = chunks.filter(
    (c): c is SyncChunk<T> => !!c && typeof c.index === 'number' && Array.isArray(c.items),
  )
  valid.sort((a, b) => a.index - b.index)
  const out: T[] = []
  for (const c of valid) out.push(...c.items)
  return out
}

// ---- v4 budget planner ------------------------------------------------------

export type SyncPlanInput = {
  /** The whole local library; non-syncable rules are filtered here. */
  rules: Rule[]
  tombstones: RuleTombstone[]
  /** Bytes available for row + folder + tombstone items (keys + values). */
  budgetBytes: number
  /** Tombstones' first claim (they may exceed it only into leftover budget). */
  tombstoneShareBytes: number
  /** Stamped as `p` on every chunk. */
  pushId: string
  keyPrefix: { rows: string; folders: string; tombs: string }
  /** For the "recent tombstone left out" count. */
  now: number
}

export type SyncPayloadPlan = {
  rowChunks: SyncChunk<WireRuleEntry>[]
  folderChunks: SyncChunk<FolderDictEntry>[]
  tombChunks: SyncChunk<CompactTombstone>[]
  /** Flattened chunk contents (deterministic order) — the content fingerprint input. */
  rows: WireRuleEntry[]
  folders: FolderDictEntry[]
  tombstones: CompactTombstone[]
  /** Exact bytes of every row / folder / tombstone write (keys + values). */
  bytes: number
  fullRowCount: number
  stubCount: number
  /** Syncable rules with NO cloud representation (no full row, no stub). */
  truncatedRuleCount: number
  /** Tombstones from the last 30 days that did not fit (older ones are left out by design). */
  truncatedTombstoneCount: number
  /** Every tombstone not in the cloud, for logging. */
  omittedTombstoneCount: number
}

// Leeway for JSON commas / chunk-boundary slack the item-cost model ignores.
const PLAN_SAFETY_BYTES = 256
const RECENT_TOMBSTONE_MS = 30 * 24 * 60 * 60 * 1000

/**
 * Select what fits the sync byte budget — ONE encoding pass plus a handful of
 * exact verifications (2026-11-05). The old loop re-encoded and re-serialised
 * the whole library once per dropped rule (~170× per push at 325 rules).
 *
 * Order of claims on the budget:
 *   1. Tombstones, NEWEST deletedAt first, up to `tombstoneShareBytes`. They
 *      used to be truncated by key order and charged before rules, so 300 old
 *      CJK tombstones crowded out the newest deletions AND ~69 rules.
 *   2. A stub for EVERY disabled syncable rule — never dropped (unless stubs
 *      alone overflow the whole budget, oldest edits dropped first). This is
 *      what makes disables propagate under budget pressure.
 *   3. Full rows in compareRuleSyncPriority order, largest prefix that fits
 *      (a disabled rule's full row replaces its stub, so it only costs the
 *      difference; a folder entry is charged on first use).
 *   4. Leftover budget → more tombstones, still newest first.
 */
export function planSyncPayload(input: SyncPlanInput): SyncPayloadPlan {
  const budget = Math.max(0, input.budgetBytes)
  const { pushId, keyPrefix } = input
  const envelopeBytes = (prefix: string) =>
    utf8ByteLength(prefix) + 3 + jsonByteSize({ index: 99, items: [], p: pushId })

  // --- candidates (one encode each) ---
  const ranked = input.rules.filter(shouldSyncRule).sort(compareRuleSyncPriority)
  const { folders: allFolders, refOf } = buildFolderDict(ranked)
  const folderByRef = new Map(allFolders.map((f) => [f.x, f]))
  type Cand = {
    rule: Rule
    row: CompactRuleRow
    ref: string
    rowCost: number
    folderCost: number // charged when this candidate is the folder's first user
    stub?: DisabledStub
    stubCost: number
  }
  const cands: Cand[] = []
  let unencodable = 0
  const firstUse = new Set<string>()
  for (const r of ranked) {
    const ref = refOf(r)
    const row = encodeRuleRow(r, ref)
    if (!row) {
      unencodable++
      continue
    }
    const stub = r.enabled ? undefined : encodeDisabledStub(r)
    let folderCost = 0
    if (!firstUse.has(ref)) {
      firstUse.add(ref)
      folderCost = jsonByteSize(folderByRef.get(ref)) + 1
    }
    cands.push({
      rule: r,
      row,
      ref,
      rowCost: jsonByteSize(row) + 1,
      folderCost,
      stub,
      stubCost: stub ? jsonByteSize(stub) + 1 : 0,
    })
  }

  // --- tombstones, newest first (key breaks ties → deterministic) ---
  const tombEnc: Array<{ e: CompactTombstone; cost: number; deletedAt: number }> = []
  for (const t of [...input.tombstones].sort(
    (a, b) =>
      b.deletedAt - a.deletedAt ||
      cmpStr(
        `${a.type}::${a.signalNorm}::${a.targetFolderPath}`,
        `${b.type}::${b.signalNorm}::${b.targetFolderPath}`,
      ),
  )) {
    if (tombEnc.length >= TOMBSTONE_SYNC_CAP) break
    const e = encodeTombstone(t)
    if (e) tombEnc.push({ e, cost: jsonByteSize(e) + 1, deletedAt: t.deletedAt })
  }

  // Rough per-category envelope estimate; exact bytes are verified below.
  const estimate = (itemBytes: number, prefix: string) => {
    if (itemBytes <= 0) return 0
    const env = envelopeBytes(prefix)
    return itemBytes + env * Math.ceil(itemBytes / Math.max(1024, CHUNK_BYTE_TARGET - env - 512))
  }
  const tombPrefixWithin = (limit: number, from = 0, start = 0) => {
    let n = from
    let acc = start
    while (n < tombEnc.length && estimate(acc + tombEnc[n]!.cost, keyPrefix.tombs) <= limit) {
      acc += tombEnc[n]!.cost
      n++
    }
    return n
  }

  // 1. Tombstones inside their share.
  let tombN = tombPrefixWithin(Math.min(input.tombstoneShareBytes, budget))
  const tombBytesOf = (n: number) => tombEnc.slice(0, n).reduce((s, t) => s + t.cost, 0)

  // 2. Stubs for every disabled rule (newest edit first if they overflow).
  const stubSet = new Set<string>()
  {
    const stubCands = cands
      .filter((c) => c.stub)
      .sort((a, b) => (b.rule.editedAt ?? 0) - (a.rule.editedAt ?? 0) || cmpStr(a.rule.id, b.rule.id))
    const room = budget - estimate(tombBytesOf(tombN), keyPrefix.tombs)
    let acc = 0
    for (const c of stubCands) {
      if (estimate(acc + c.stubCost, keyPrefix.rows) > room) break
      acc += c.stubCost
      stubSet.add(c.rule.id)
    }
  }

  // Marginal cost of promoting candidate i to a full row, in rank order.
  const marginal = cands.map(
    (c) => c.rowCost + c.folderCost - (stubSet.has(c.rule.id) ? c.stubCost : 0),
  )

  const materialize = (k: number, nTombs: number) => {
    const rows: WireRuleEntry[] = []
    const usedRefs = new Set<string>()
    let fullRowCount = 0
    let stubCount = 0
    for (let i = 0; i < cands.length; i++) {
      const c = cands[i]!
      if (i < k) {
        rows.push(c.row)
        usedRefs.add(c.ref)
        fullRowCount++
      } else if (c.stub && stubSet.has(c.rule.id)) {
        rows.push(c.stub)
        stubCount++
      }
    }
    const folders = allFolders.filter((f) => usedRefs.has(f.x))
    const tombs = tombEnc.slice(0, nTombs).map((t) => t.e)
    const rowChunks = chunkCompactRows(rows, pushId)
    const folderChunks = chunkFolderDict(folders, pushId)
    const tombChunks = chunkCompactTombstones(tombs, pushId)
    let bytes = 0
    for (const c of rowChunks) bytes += writeBytes(`${keyPrefix.rows}${c.index}`, c)
    for (const c of folderChunks) bytes += writeBytes(`${keyPrefix.folders}${c.index}`, c)
    for (const c of tombChunks) bytes += writeBytes(`${keyPrefix.tombs}${c.index}`, c)
    const capsOk =
      rowChunks.length <= MAX_RULE_CHUNKS &&
      folderChunks.length <= MAX_FOLDER_DICT_CHUNKS &&
      tombChunks.length <= MAX_TOMBSTONE_CHUNKS
    return { k, nTombs, rowChunks, folderChunks, tombChunks, bytes, capsOk, fullRowCount, stubCount }
  }
  type Mat = ReturnType<typeof materialize>
  const fits = (m: Mat) => m.capsOk && m.bytes <= budget

  // Pathological floor: stubs + tombstones alone overflow → shed tombstones,
  // then the oldest stubs, until an empty-row payload fits.
  let floor = materialize(0, tombN)
  while (!fits(floor) && (tombN > 0 || stubSet.size > 0)) {
    if (tombN > 0) tombN = Math.floor(tombN / 2)
    else {
      const oldest = cands
        .filter((c) => stubSet.has(c.rule.id))
        .sort((a, b) => (a.rule.editedAt ?? 0) - (b.rule.editedAt ?? 0))
      for (const c of oldest.slice(0, Math.max(1, Math.ceil(oldest.length / 4)))) {
        stubSet.delete(c.rule.id)
        marginal[cands.indexOf(c)] = c.rowCost + c.folderCost
      }
    }
    floor = materialize(0, tombN)
  }

  // 3. Largest full-row prefix that fits: estimate from prefix sums, then
  // correct with exact materialisations (each O(n); typically 2–3 of them).
  let best: Mat | null = fits(floor) ? floor : null
  if (best) {
    let k = 0
    {
      const base = floor.bytes
      let acc = 0
      while (k < cands.length && base + acc + marginal[k]! + PLAN_SAFETY_BYTES <= budget) {
        acc += marginal[k]!
        k++
      }
    }
    let hiFail = cands.length + 1
    for (let iter = 0; iter < 12 && k > best.k; iter++) {
      const m = materialize(k, tombN)
      if (fits(m)) {
        best = m
        const slack = budget - m.bytes - PLAN_SAFETY_BYTES
        let k2 = k
        let acc = 0
        while (k2 < cands.length && k2 + 1 < hiFail && acc + marginal[k2]! <= slack) {
          acc += marginal[k2]!
          k2++
        }
        if (k2 === k) break
        k = k2
      } else {
        hiFail = Math.min(hiFail, k)
        const over = m.capsOk ? m.bytes - budget + PLAN_SAFETY_BYTES : CHUNK_BYTE_TARGET
        let k2 = k
        let acc = 0
        while (k2 > best.k && acc < over) {
          k2--
          acc += Math.max(1, marginal[k2]!)
        }
        k = k2
      }
    }
  }
  if (!best) best = materialize(0, 0) // cannot happen once stubs/tombs are shed; defensive

  // 4. Leftover budget → more tombstones (newest first).
  if (tombN < tombEnc.length) {
    let n = tombPrefixWithin(budget - (best.bytes - estimate(tombBytesOf(tombN), keyPrefix.tombs)), tombN, tombBytesOf(tombN))
    for (let iter = 0; iter < 6 && n > tombN; iter++) {
      const m = materialize(best.k, n)
      if (fits(m)) {
        best = m
        tombN = n
        break
      }
      n = tombN + Math.floor((n - tombN) / 2)
    }
  }

  const tombs = best.tombChunks.flatMap((c) => c.items)
  const recentCutoff = input.now - RECENT_TOMBSTONE_MS
  const truncatedTombstoneCount = input.tombstones.filter((t) => t.deletedAt >= recentCutoff).length -
    tombEnc.slice(0, best.nTombs).filter((t) => t.deletedAt >= recentCutoff).length
  return {
    rowChunks: best.rowChunks,
    folderChunks: best.folderChunks,
    tombChunks: best.tombChunks,
    rows: best.rowChunks.flatMap((c) => c.items),
    folders: best.folderChunks.flatMap((c) => c.items),
    tombstones: tombs,
    bytes: best.bytes,
    fullRowCount: best.fullRowCount,
    stubCount: best.stubCount,
    truncatedRuleCount: unencodable + (cands.length - best.fullRowCount - best.stubCount),
    truncatedTombstoneCount: Math.max(0, truncatedTombstoneCount),
    omittedTombstoneCount: input.tombstones.length - best.nTombs,
  }
}
