// Chunking helpers for chrome.storage.sync writes.
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

import { CHUNK_BYTE_TARGET, FOLDER_ACTIVITY_SYNC_CAP, TOMBSTONE_SYNC_CAP } from './constants'
import type { FolderActivity, Rule, RuleSource, RuleTombstone } from './types'

/** Sync envelope. Stored under one `chrome.storage.sync` key per chunk. */
export type SyncChunk<T> = {
  /** 0-indexed within its category. Used by callers to know how many to fetch. */
  index: number
  items: T[]
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
): SyncChunk<FolderActivity>[] {
  // Sort by lastActiveAt DESC to grab the most-recent entries.
  const recent = [...activity]
    .sort((a, b) => b.lastActiveAt.localeCompare(a.lastActiveAt))
    .slice(0, FOLDER_ACTIVITY_SYNC_CAP)
  // Re-sort by folderId for stable chunk membership — without this, the
  // same set of entries could land in different chunks each push and burn
  // the per-hour write quota.
  return chunkArray(recent, (e) => e.folderId)
}

export function chunkTombstones(tombstones: RuleTombstone[]): SyncChunk<RuleTombstone>[] {
  const recent = [...tombstones]
    .sort((a, b) => b.deletedAt - a.deletedAt)
    .slice(0, TOMBSTONE_SYNC_CAP)
  // Re-sort by deterministic key (type + signalNorm + targetFolderPath) so
  // chunk membership is stable across pushes.
  recent.sort((a, b) => {
    const ka = `${a.type}::${a.signalNorm}::${a.targetFolderPath}`
    const kb = `${b.type}::${b.signalNorm}::${b.targetFolderPath}`
    return ka.localeCompare(kb)
  })
  return chunkArray(recent, (t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`)
}

function chunkArray<T>(items: T[], keyFn: (item: T) => string): SyncChunk<T>[] {
  const sorted = [...items].sort((a, b) => keyFn(a).localeCompare(keyFn(b)))
  const chunks: SyncChunk<T>[] = []
  let current: T[] = []
  let currentBytes = JSON.stringify({ index: 0, items: [] }).length // envelope overhead

  for (const item of sorted) {
    // +1 for the comma in the eventual array serialisation. Slight
    // over-estimate but safe.
    const itemSize = JSON.stringify(item).length + 1
    if (current.length > 0 && currentBytes + itemSize > CHUNK_BYTE_TARGET) {
      chunks.push({ index: chunks.length, items: current })
      current = []
      currentBytes = JSON.stringify({ index: chunks.length, items: [] }).length
    }
    current.push(item)
    currentBytes += itemSize
  }
  if (current.length > 0) {
    chunks.push({ index: chunks.length, items: current })
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
