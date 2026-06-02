// Subject normalization used by:
//   - popup auto-propagate (within-batch siblings sharing a subject)
//   - thread memory pre-filter (cross-batch matching: previous batch
//     filed this subject to folder X → propose same target this batch)
//
// Strips reply / forward / system prefixes iteratively because
// real-world subjects have stacked prefixes like
// "Re: Fw: [External] 自動回覆: 主旨". Then collapses whitespace and
// lowercases so two clients sending "  Re:foo " vs "RE: FOO" still
// hash to the same key.
//
// Three categories of strippable prefix:
//   1. Reply/forward markers — Latin (RE / FW / FWD / REPLY / R) and
//      Chinese (回覆 / 答覆 / 轉寄 / 轉發 / 轉). Re[2]: style numbered
//      replies handled by the optional [\d+] suffix on the marker.
//   2. System auto-responses — Outlook's "自動回覆:" / "Out of Office:" /
//      "Automatic reply:" / "Auto-reply:".
//   3. Tag-style brackets — "[External]" / "[外部]" / "[已讀回條]" /
//      "[已讀回執]" — Exchange / O365 admins add these system-wide.

// One marker = one strip pass; iteration handles stacked prefixes.
const STRIP_PATTERNS: RegExp[] = [
  // Reply/forward markers with optional [n] count and colon/separator.
  /^\s*(?:re|fw|fwd|reply|r)(?:\s*\[\d+\])?[:：\s]+/i,
  /^\s*(?:回覆|答覆|轉寄|轉發|轉)[:：\s]+/,
  // System auto-responses (label part — actual subject usually follows).
  /^\s*(?:自動回覆|自動回复|out of office|automatic reply|auto[- ]?reply)[:：\s]+/i,
  // Tag-style bracketed prefixes (admin / system).
  /^\s*\[(?:external|外部|內部|internal|已讀回條|已讀回執|讀取回條|已讀|未讀|已讀回覆|附件警告|spam|junk|caution)\]\s*/i,
]

export function normalizeSubject(s: string): string {
  let prev = ''
  let cur = s
  while (cur !== prev) {
    prev = cur
    for (const re of STRIP_PATTERNS) {
      cur = cur.replace(re, '')
    }
  }
  return cur.trim().replace(/\s+/g, ' ').toLowerCase()
}

/**
 * Subjects shorter than this (after normalization) are too generic to
 * use as a cross-batch routing signal — "報告" / "請示" / "通知" /
 * "Reply" would otherwise pollute the subject memory with high-
 * confidence false positives.
 */
export const MIN_NORMALIZED_SUBJECT_LEN = 8
