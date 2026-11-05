// OWA manual-filing harvester — provenance filter (batch: harvester MVP).
//
// The lawyer sometimes files mail by dragging it into a case folder directly in
// Outlook, bypassing this extension. We want to LEARN that routing — but ONLY
// into thread-memory (local, non-authoritative, conflict-gated), never a hard
// rule, because there is no way to prove from Graph that a human (not our own
// move, not an Outlook server rule) put the mail there. `filterHarvestCandidates`
// is the safety valve: a pure, testable exclusion pipeline that keeps only mail
// we're confident is a NEW, human-originated filing signal.

import { matchEmailWithIndex, type RuleIndex } from '@/shared/rules'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from '@/shared/normalize'
import type { ModifiedMessage } from '@/shared/outlook-api'
import type { MovedByUsEntry } from '@/shared/storage'
import type { Email, ThreadMemoryEntry } from '@/shared/types'

// ---- Scan window / watermark (2026-11-05) ----------------------------------
//
// The harvester used to probe every leaf folder with `$top=10` (no paging):
// with no watermark there was no $filter at all, with one our own moves still
// ate the 10 slots, and the watermark was the client clock at the END of a
// ~200-request run — filings made meanwhile (or inside client/server clock
// skew) were lost for good. Now: ONE mailbox-wide `LastModifiedDateTime gt W`
// query, paged, filtered client-side to the target folders; W is taken at
// scan START minus a margin and only advances after a COMPLETE paging run.

/** First-run (and long-idle) lookback bound. */
export const HARVEST_LOOKBACK_MS = 14 * 24 * 60 * 60 * 1000
/** Overlap re-scanned every run — absorbs client↔server clock skew. Re-seen
 *  filings are harmless: an already-recorded route is filtered out by
 *  filterHarvestCandidates step (3). */
export const HARVEST_WATERMARK_MARGIN_MS = 10 * 60 * 1000
export const HARVEST_PAGE_SIZE = 100
/** 3000 modified messages per run — covers ~2 weeks for a busy mailbox. */
export const HARVEST_MAX_PAGES = 30

/** OData DateTimeOffset literal, second precision, UTC. */
export function toODataDateTime(ms: number): string {
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/**
 * Lower bound for this run's `LastModifiedDateTime gt …`: the stored
 * watermark clamped into [now − LOOKBACK, now − MARGIN]. No / garbage
 * watermark (first run) → the bounded lookback. A watermark in the future
 * (client clock was ahead, then corrected) is pulled back so it can't blind
 * the harvester until real time catches up.
 */
export function harvestWindowStart(lastHarvestAt: string | null, nowMs: number): string {
  const floor = nowMs - HARVEST_LOOKBACK_MS
  const ceil = nowMs - HARVEST_WATERMARK_MARGIN_MS
  const parsed = lastHarvestAt ? Date.parse(lastHarvestAt) : Number.NaN
  const w = Number.isFinite(parsed) ? Math.min(Math.max(parsed, floor), ceil) : floor
  return toODataDateTime(w)
}

/** Watermark to persist after a COMPLETE run that started at `scanStartMs`. */
export function nextHarvestWatermark(scanStartMs: number): string {
  return new Date(scanStartMs - HARVEST_WATERMARK_MARGIN_MS).toISOString()
}

/**
 * Keep the mailbox-wide rows that currently sit in a harvest target folder
 * (`targets`: folderId → folderPath), shaped for filterHarvestCandidates.
 */
export function selectHarvestScanItems(
  messages: ModifiedMessage[],
  targets: ReadonlyMap<string, string>,
): HarvestScanItem[] {
  const out: HarvestScanItem[] = []
  for (const m of messages) {
    const folderId = m.ParentFolderId
    if (!m.Id || !folderId) continue
    const folderPath = targets.get(folderId)
    if (folderPath === undefined) continue
    out.push({
      id: m.Id,
      subject: m.Subject ?? '',
      from: m.From?.EmailAddress?.Address,
      conversationId: m.ConversationId,
      folderId,
      folderPath,
    })
  }
  return out
}

/** One message the harvester scanned from a leaf folder. */
export type HarvestScanItem = {
  id: string
  subject: string
  from?: string
  conversationId?: string
  folderId: string
  folderPath: string
}

/** A suspected manual filing that survived every exclusion — ready to feed
 *  thread-memory (conversation + normalized subject). */
export type HarvestCandidate = {
  folderId: string
  folderPath: string
  conversationId?: string
  /** Present only when the subject is long enough to be a usable key. */
  normalizedSubject?: string
}

/**
 * Keep only scanned mail that is a NEW, human-originated filing signal. Drops,
 * in order:
 *   1. Our own moves — matched by the moved-by-us ledger (destination-side new
 *      id, and a convId+folder near-match as a second layer).
 *   2. Mail already covered by an ENABLED rule — either a rule moved it, or a
 *      rule already routes it, so re-learning is pointless / conflict-prone.
 *      (Graph can't tell "rule moved it" from "human moved it"; we cut the
 *      whole class to stay conservative — 寧缺勿濫.)
 *   3. A conversation / subject already pointing to the SAME folder in
 *      thread-memory — known route, not a new signal.
 * `ruleIndex` MUST be built from enabled rules only (buildRuleIndex already
 * excludes disabled rules).
 */
export function filterHarvestCandidates(
  scanned: HarvestScanItem[],
  opts: {
    movedByUs: MovedByUsEntry[]
    ruleIndex: RuleIndex
    convMemory: Record<string, ThreadMemoryEntry>
    subjectMemory: Record<string, ThreadMemoryEntry>
  },
): HarvestCandidate[] {
  const movedIds = new Set(opts.movedByUs.map((e) => e.newMessageId))
  const movedConvFolder = new Set(
    opts.movedByUs
      .filter((e) => e.conversationId)
      .map((e) => `${e.conversationId}::${e.destinationFolderId}`),
  )

  const out: HarvestCandidate[] = []
  for (const it of scanned) {
    if (!it.id || !it.folderId) continue
    // (1) our own move
    if (movedIds.has(it.id)) continue
    if (it.conversationId && movedConvFolder.has(`${it.conversationId}::${it.folderId}`)) continue

    // (2) already covered by an enabled rule
    const email: Email = {
      Id: it.id,
      Subject: it.subject ?? '',
      BodyPreview: '',
      From: { EmailAddress: { Address: it.from ?? '' } },
      ToRecipients: [],
      ReceivedDateTime: '',
      ParentFolderId: it.folderId,
    }
    if (matchEmailWithIndex(email, opts.ruleIndex)) continue

    // (3) already a known thread-memory route to the SAME folder
    const norm = normalizeSubject(it.subject ?? '')
    const hasNorm = norm.length >= MIN_NORMALIZED_SUBJECT_LEN
    const convKnown =
      !!it.conversationId && opts.convMemory[it.conversationId]?.folderId === it.folderId
    const subjKnown = hasNorm && opts.subjectMemory[norm]?.folderId === it.folderId
    if (convKnown || subjKnown) continue

    out.push({
      folderId: it.folderId,
      folderPath: it.folderPath,
      conversationId: it.conversationId,
      normalizedSubject: hasNorm ? norm : undefined,
    })
  }
  return out
}
