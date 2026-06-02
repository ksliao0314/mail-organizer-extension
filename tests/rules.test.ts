import { describe, expect, it } from 'vitest'
import type { Email, Rule } from '@/shared/types'
import {
  addRules,
  applyConfidenceCap,
  bumpRuleHit,
  bumpRuleHits,
  bumpRuleOverrides,
  buildRuleIndex,
  dedupeRulesByKey,
  encodeCompound,
  extractCaseCodes,
  extractDomain,
  extractSubjectSignal,
  findConflicts,
  listRules,
  matchEmail,
  matchEmailWithIndex,
  newRule,
  ruleBeatsThread,
  toggleRule,
  upsertRule,
} from '@/shared/rules'

// ---- helpers ---------------------------------------------------------------

function makeEmail(partial: Partial<Email> & { Id: string }): Email {
  return {
    Subject: '',
    BodyPreview: '',
    From: { EmailAddress: { Address: 'unknown@example.com' } },
    ToRecipients: [],
    ReceivedDateTime: '2026-01-01T00:00:00Z',
    ParentFolderId: 'inbox',
    ...partial,
  } as Email
}

function rule(over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
    targetFolderPath: 'X/Y',
    confidence: 0.8,
    source: 'user_manual',
    ...over,
  })
}

// ---- matching --------------------------------------------------------------

describe('matchEmail', () => {
  it('matches case_code in subject case-insensitively', () => {
    const r = rule({ type: 'case_code', signal: '25A0067A' })
    const e = makeEmail({ Id: 'e1', Subject: 'RE: 25a0067a 案件討論' })
    const out = matchEmail(e, [r])
    expect(out?.rule.id).toBe(r.id)
    expect(out?.reason).toContain('25A0067A')
  })

  it('matches domain on sender', () => {
    const r = rule({ type: 'domain', signal: 'foodpanda.tw' })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'someone@foodpanda.tw' } },
    })
    const out = matchEmail(e, [r])
    expect(out?.rule.id).toBe(r.id)
    expect(out?.reason).toContain('寄件人網域')
  })

  it('matches domain on recipient CC', () => {
    const r = rule({ type: 'domain', signal: 'kgi.com' })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'me@example.com' } },
      ToRecipients: [{ EmailAddress: { Address: 'someone@kgi.com' } }],
    })
    const out = matchEmail(e, [r])
    expect(out?.rule.id).toBe(r.id)
    expect(out?.reason).toContain('收件人網域')
  })

  it('strips @ prefix from domain signal', () => {
    const r = rule({ type: 'domain', signal: '@foodpanda.tw' })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'x@foodpanda.tw' } },
    })
    expect(matchEmail(e, [r])).not.toBeNull()
  })

  it('matches sender exactly (case-insensitive)', () => {
    const r = rule({ type: 'sender', signal: 'alice.wang@example.com' })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'Alice.Wang@example.com' } },
    })
    expect(matchEmail(e, [r])?.rule.id).toBe(r.id)
  })

  it('matches subject_keyword case-insensitively', () => {
    const r = rule({ type: 'subject_keyword', signal: '工時審閱' })
    const e = makeEmail({ Id: 'e1', Subject: '請審閱：3月工時審閱表' })
    expect(matchEmail(e, [r])).not.toBeNull()
  })

  it('matches compound (AND of 2 conditions)', () => {
    const r = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'sender', value: 'alice.wang@example.com' },
        { type: 'subject_keyword', value: '工時審閱' },
      ]),
    })
    const fits = makeEmail({
      Id: 'e1',
      Subject: '3月工時審閱',
      From: { EmailAddress: { Address: 'alice.wang@example.com' } },
    })
    const onlyOne = makeEmail({
      Id: 'e2',
      Subject: '3月工時審閱',
      From: { EmailAddress: { Address: 'other@example.com' } },
    })
    expect(matchEmail(fits, [r])).not.toBeNull()
    expect(matchEmail(onlyOne, [r])).toBeNull()
  })

  it('respects priority: case_code (1) > domain (3)', () => {
    const codeRule = rule({
      type: 'case_code',
      signal: '25A0067A',
      targetFolderPath: 'A',
    })
    const domainRule = rule({
      type: 'domain',
      signal: 'foodpanda.tw',
      targetFolderPath: 'B',
    })
    const e = makeEmail({
      Id: 'e1',
      Subject: '25A0067A 報告',
      From: { EmailAddress: { Address: 'x@foodpanda.tw' } },
    })
    const out = matchEmail(e, [domainRule, codeRule])
    expect(out?.rule.type).toBe('case_code')
  })

  it('respects priority: compound (2) > domain (3) — split conflict semantics', () => {
    // After split-to-compound, the compound rule should win against the
    // original plain-domain rule even when both apply. Otherwise the split
    // feature is pointless.
    const domainRule = rule({
      type: 'domain',
      signal: 'dazn.com',
      targetFolderPath: '11sports',
    })
    const compoundRule = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'dazn.com' },
        { type: 'subject_keyword', value: '三二零九' },
      ]),
      targetFolderPath: 'DAZN',
    })
    const eWithKeyword = makeEmail({
      Id: 'e1',
      Subject: 'RE: 三二零九 案件討論',
      From: { EmailAddress: { Address: 'x@dazn.com' } },
    })
    const eWithoutKeyword = makeEmail({
      Id: 'e2',
      Subject: '一般通知',
      From: { EmailAddress: { Address: 'x@dazn.com' } },
    })

    // Compound wins when both conditions match
    expect(matchEmail(eWithKeyword, [domainRule, compoundRule])?.rule.type).toBe('compound')
    // Falls through to domain when keyword absent
    expect(matchEmail(eWithoutKeyword, [domainRule, compoundRule])?.rule.type).toBe('domain')
  })

  it('within same type, higher confidence wins', () => {
    const a = rule({ confidence: 0.7, targetFolderPath: 'low' })
    const b = rule({ confidence: 0.95, targetFolderPath: 'high' })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'x@example.com' } },
    })
    expect(matchEmail(e, [a, b])?.rule.id).toBe(b.id)
  })

  it('skips disabled rules', () => {
    const r = rule({ enabled: false })
    const e = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'x@example.com' } },
    })
    expect(matchEmail(e, [r])).toBeNull()
  })

  it('returns null when nothing matches', () => {
    const r = rule({ type: 'domain', signal: 'something-else.com' })
    const e = makeEmail({ Id: 'e1' })
    expect(matchEmail(e, [r])).toBeNull()
  })
})

// ---- conflict detection ----------------------------------------------------

describe('findConflicts', () => {
  it('detects same (type, signal) with different targets', () => {
    const a = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'A', targetFolderPath: 'A' })
    const b = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'B', targetFolderPath: 'B' })
    const conflicts = findConflicts([a, b])
    expect(conflicts).toHaveLength(1)
    expect(conflicts[0]!.rules).toHaveLength(2)
  })

  it('does not flag duplicates with same target', () => {
    const a = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'A' })
    const b = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'A' })
    expect(findConflicts([a, b])).toHaveLength(0)
  })

  it('compound conflict detection is order-insensitive', () => {
    const sigA = encodeCompound([
      { type: 'sender', value: 'x@y.com' },
      { type: 'subject_keyword', value: 'foo' },
    ])
    const sigB = encodeCompound([
      { type: 'subject_keyword', value: 'foo' },
      { type: 'sender', value: 'x@y.com' },
    ])
    const a = rule({ type: 'compound', signal: sigA, targetFolderId: 'A' })
    const b = rule({ type: 'compound', signal: sigB, targetFolderId: 'B' })
    expect(findConflicts([a, b])).toHaveLength(1)
  })

  it('ignores disabled rules', () => {
    const a = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'A' })
    const b = rule({ type: 'domain', signal: 'kgi.com', targetFolderId: 'B', enabled: false })
    expect(findConflicts([a, b])).toHaveLength(0)
  })
})

// ---- extraction helpers ----------------------------------------------------

describe('extractCaseCodes', () => {
  it('finds canonical pattern in text', () => {
    expect(extractCaseCodes('案件 25A0067A 進度')).toEqual(['25A0067A'])
    expect(extractCaseCodes('案件 25a0067a 進度')).toEqual(['25A0067A'])
  })

  it('dedupes multiple occurrences', () => {
    expect(extractCaseCodes('25A0067A vs 25A0067A 同案')).toEqual(['25A0067A'])
  })

  it('returns empty when no codes', () => {
    expect(extractCaseCodes('no codes here')).toEqual([])
    expect(extractCaseCodes('')).toEqual([])
  })

  it('rejects malformed patterns', () => {
    expect(extractCaseCodes('25A067A')).toEqual([])
    expect(extractCaseCodes('1234567')).toEqual([])
  })
})

describe('extractDomain', () => {
  it('lowercases domain', () => {
    expect(extractDomain('USER@Example.COM')).toBe('example.com')
  })

  it('handles null', () => {
    expect(extractDomain(null)).toBeNull()
    expect(extractDomain(undefined)).toBeNull()
    expect(extractDomain('')).toBeNull()
  })

  it('returns null for input without @', () => {
    expect(extractDomain('no-at-sign')).toBeNull()
  })
})

// ---- CRUD + atomicity ------------------------------------------------------

describe('rule CRUD', () => {
  it('upserts new and existing rules', async () => {
    const r1 = await upsertRule(rule({ signal: 'a.com' }))
    expect((await listRules())).toHaveLength(1)
    const r1b = await upsertRule({ ...r1, signal: 'a-updated.com' })
    expect((await listRules())).toHaveLength(1)
    expect(r1b.signal).toBe('a-updated.com')
  })

  it('addRules atomically appends multiple', async () => {
    const r1 = rule({ signal: 'a.com' })
    const r2 = rule({ signal: 'b.com' })
    await addRules([r1, r2])
    expect((await listRules())).toHaveLength(2)
  })

  it('toggleRule flips enabled', async () => {
    const r1 = await upsertRule(rule())
    await toggleRule(r1.id, false)
    const all = await listRules()
    expect(all[0]!.enabled).toBe(false)
  })

  it('bumpRuleHit increments matchCount + sets lastUsedAt', async () => {
    const r1 = await upsertRule(rule())
    await bumpRuleHit(r1.id)
    await bumpRuleHit(r1.id)
    const all = await listRules()
    expect(all[0]!.matchCount).toBe(2)
    expect(all[0]!.lastUsedAt).toBeDefined()
  })

  it('mutex serializes concurrent bumpRuleHit calls (no clobber)', async () => {
    // This is the P0-6 race fix: 10 concurrent bumps should result in exactly
    // 10, not some lossy partial number due to read-modify-write race.
    const r1 = await upsertRule(rule())
    await Promise.all(Array.from({ length: 10 }, () => bumpRuleHit(r1.id)))
    const all = await listRules()
    expect(all[0]!.matchCount).toBe(10)
  })
})

// ---- bumpRuleHits errorRate gate (added 2026-05-22) -----------------------
//
// Pre-2026-05-22, every 20 hits → confidence +0.05 unconditionally. A rule
// with high overrideCount (user keeps disagreeing) would still climb to 1.0
// raw confidence, misleading the UI even though effectiveConfidence at
// match-time already demoted it. Now the promotion only fires when
// errorRate (overrideCount / matchCount) ≤ 0.2.
describe('bumpRuleHits errorRate gate', () => {
  it('promotes confidence on healthy rules (low error rate)', async () => {
    const r1 = await upsertRule(rule({ confidence: 0.55, source: 'auto_scan' }))
    // 20 hits, 0 overrides → errorRate = 0% → should promote.
    await bumpRuleHits(new Map([[r1.id, 20]]))
    const all = await listRules()
    expect(all[0]!.confidence).toBeGreaterThan(0.55)
    expect(all[0]!.matchCount).toBe(20)
  })

  it('blocks promotion when error rate exceeds 20%', async () => {
    const r1 = await upsertRule(rule({ confidence: 0.55, source: 'auto_scan' }))
    // Seed 5 overrides first
    await bumpRuleOverrides(new Map([[r1.id, 5]]))
    // Then 20 hits — errorRate = 5/20 = 25% → should NOT promote.
    await bumpRuleHits(new Map([[r1.id, 20]]))
    const all = await listRules()
    expect(all[0]!.confidence).toBe(0.55)
    expect(all[0]!.matchCount).toBe(20)
    expect(all[0]!.overrideCount).toBe(5)
  })

  it('user_manual rules never get auto-promoted regardless of error rate', async () => {
    const r1 = await upsertRule(rule({ confidence: 0.8, source: 'user_manual' }))
    await bumpRuleHits(new Map([[r1.id, 40]]))
    const all = await listRules()
    expect(all[0]!.confidence).toBe(0.8)
  })
})

// ---- subject_keyword longest-match (added 2026-05-22) ---------------------
//
// Pre-2026-05-22, when two subject_keyword rules both matched, whichever
// sorted first by confidence won. That meant a broad "通知" could shadow a
// specific "112訴204" if its confidence happened to be higher. The fix:
// subjectKeywordRules bucket is re-sorted longest-signal-first inside
// buildRuleIndex, so the more specific signal always tries first.
describe('subject_keyword longest-match', () => {
  it('longer signal wins over shorter when both match', () => {
    const rBroad = rule({
      type: 'subject_keyword',
      signal: '通知',
      confidence: 0.9, // higher confidence — would have won pre-fix
      targetFolderPath: 'broad/folder',
    })
    const rSpecific = rule({
      type: 'subject_keyword',
      signal: '112訴204',
      confidence: 0.5, // lower confidence
      targetFolderPath: 'specific/folder',
    })
    const idx = buildRuleIndex([rBroad, rSpecific])
    const email = makeEmail({
      Id: 'e1',
      Subject: '112訴204的開庭通知',
    })
    const out = matchEmailWithIndex(email, idx)
    // Should pick the specific one because its signal is longer.
    expect(out?.rule.targetFolderPath).toBe('specific/folder')
  })

  it('shorter signal still wins when longer doesn\'t match', () => {
    const rShort = rule({
      type: 'subject_keyword',
      signal: '通知',
      targetFolderPath: 'short/folder',
    })
    const rLong = rule({
      type: 'subject_keyword',
      signal: 'completely-different-long-keyword',
      targetFolderPath: 'long/folder',
    })
    const idx = buildRuleIndex([rShort, rLong])
    const email = makeEmail({ Id: 'e1', Subject: '會議通知' })
    const out = matchEmailWithIndex(email, idx)
    expect(out?.rule.targetFolderPath).toBe('short/folder')
  })
})

// ---- TYPE_PRIORITY match order contract -----------------------------------
//
// History note: a brief 2026-05-22 fix moved sender ABOVE domain in
// matchEmailWithIndex's iteration order. The intent was "sender is more
// specific than domain", but it inadvertently also placed sender above
// subject_keyword — which silently broke working subject_keyword rules
// whenever a sender rule existed for the same From. Same-day revert
// restored TYPE_PRIORITY order. These tests pin the contract: domain
// outranks sender, subject_keyword outranks sender, neither outranks the
// other directly (different positions in TYPE_PRIORITY).
describe('TYPE_PRIORITY match order', () => {
  it('domain beats sender when both match (TYPE_PRIORITY 3 > 5)', () => {
    const rDomain = rule({
      type: 'domain',
      signal: 'gmail.com',
      targetFolderPath: 'misc/gmail',
    })
    const rSender = rule({
      type: 'sender',
      signal: 'andy@gmail.com',
      targetFolderPath: 'important/andy',
    })
    const idx = buildRuleIndex([rDomain, rSender])
    const email = makeEmail({
      Id: 'e1',
      From: { EmailAddress: { Address: 'andy@gmail.com' } },
    })
    const out = matchEmailWithIndex(email, idx)
    expect(out?.rule.targetFolderPath).toBe('misc/gmail')
  })

  it('subject_keyword beats sender (TYPE_PRIORITY 4 > 5)', () => {
    const rSender = rule({
      type: 'sender',
      signal: 'andy@gmail.com',
      targetFolderPath: 'misc/andy',
    })
    const rSubject = rule({
      type: 'subject_keyword',
      signal: '工時審閱',
      targetFolderPath: 'work/timesheet',
    })
    const idx = buildRuleIndex([rSender, rSubject])
    const email = makeEmail({
      Id: 'e1',
      Subject: '工時審閱詢問',
      From: { EmailAddress: { Address: 'andy@gmail.com' } },
    })
    const out = matchEmailWithIndex(email, idx)
    // subject_keyword (priority 4) wins over sender (priority 5)
    expect(out?.rule.targetFolderPath).toBe('work/timesheet')
  })

  it('sender wins when nothing higher matches', () => {
    const rSender = rule({
      type: 'sender',
      signal: 'andy@gmail.com',
      targetFolderPath: 'important/andy',
    })
    const idx = buildRuleIndex([rSender])
    const email = makeEmail({
      Id: 'e1',
      Subject: 'no keyword match here',
      From: { EmailAddress: { Address: 'andy@gmail.com' } },
    })
    const out = matchEmailWithIndex(email, idx)
    expect(out?.rule.targetFolderPath).toBe('important/andy')
  })
})

// ---- ruleBeatsThread precedence helper (added 2026-05-22) ------------------
//
// classifyPreflight uses this to decide whether a rule outcome should
// preempt a thread-memory hit. Pre-2026-05-22, thread memory always won
// — which meant a user_manual rule added AFTER the thread was learned
// could be invisibly shadowed. Fix: case_code / compound / user_manual
// outrank thread; everything else lets thread win.
describe('ruleBeatsThread precedence', () => {
  it('case_code rule always beats thread (even when source=auto_scan)', () => {
    expect(ruleBeatsThread({ type: 'case_code', source: 'auto_scan' })).toBe(true)
    expect(ruleBeatsThread({ type: 'case_code', source: 'ai_confirmed' })).toBe(true)
  })

  it('compound rule always beats thread', () => {
    expect(ruleBeatsThread({ type: 'compound', source: 'auto_scan' })).toBe(true)
    expect(ruleBeatsThread({ type: 'compound', source: 'ai_overridden' })).toBe(true)
  })

  it('user_manual rule beats thread regardless of type', () => {
    expect(ruleBeatsThread({ type: 'domain', source: 'user_manual' })).toBe(true)
    expect(ruleBeatsThread({ type: 'sender', source: 'user_manual' })).toBe(true)
    expect(ruleBeatsThread({ type: 'subject_keyword', source: 'user_manual' })).toBe(true)
  })

  it('broad auto rules let thread win', () => {
    expect(ruleBeatsThread({ type: 'domain', source: 'auto_scan' })).toBe(false)
    expect(ruleBeatsThread({ type: 'domain', source: 'ai_confirmed' })).toBe(false)
    expect(ruleBeatsThread({ type: 'sender', source: 'auto_scan' })).toBe(false)
    expect(ruleBeatsThread({ type: 'subject_keyword', source: 'ai_overridden' })).toBe(false)
  })
})

// ---- extractSubjectSignal (2026-05-27 redesign) ---------------------------
//
// Returns the entire normalized subject (reply/forward prefixes stripped,
// lowercased, whitespace collapsed) as a routing signal — or '' when the
// normalized result is below MIN_SUBJECT_SIGNAL_LEN (2 chars).
//
// Design: same-domain + same-normalized-subject → same case (user
// couldn't distinguish them either). So routing accuracy via exactness,
// no proper-noun heuristics, no stopword lists.
describe('extractSubjectSignal', () => {
  it('returns the entire normalized subject for typical Chinese cases', () => {
    expect(extractSubjectSignal('關於凱基證券股權移轉案')).toBe('關於凱基證券股權移轉案')
  })

  it('strips reply / forward / system prefixes (normalizeSubject behavior)', () => {
    expect(extractSubjectSignal('Re: [External] 凱基證券 詢問')).toBe('凱基證券 詢問')
    expect(extractSubjectSignal('Re: Re: Fw: 通知')).toBe('通知')
  })

  it('lowercases Latin segments and collapses whitespace', () => {
    expect(extractSubjectSignal('Project   Phoenix Q3 update')).toBe('project phoenix q3 update')
  })

  it('does NOT strip Taiwan-legal patterns (court case / case_code)', () => {
    // Unlike the old extractSubjectFeature, the court_case / case_code
    // path is selected EARLIER in chooseLearningSignal — by the time
    // extractSubjectSignal runs we know those don't apply. So full
    // subject is preserved verbatim (after prefix stripping).
    expect(extractSubjectSignal('112訴204 通知會議')).toBe('112訴204 通知會議')
  })

  it('returns empty for empty / single-char subjects (below MIN_SUBJECT_SIGNAL_LEN)', () => {
    expect(extractSubjectSignal('')).toBe('')
    expect(extractSubjectSignal('X')).toBe('')
    expect(extractSubjectSignal('  ')).toBe('')
  })

  it('returns empty when only system-prefix and tag, nothing else', () => {
    // After stripping "Re:" + "[External]" → only whitespace left.
    expect(extractSubjectSignal('Re: [External]')).toBe('')
  })

  it('accepts 2-char minimum (paired with domain in compound rule)', () => {
    // The compound rule (domain + subject) provides disambiguation, so
    // even short subjects like "通知" qualify when paired with a domain.
    expect(extractSubjectSignal('通知')).toBe('通知')
  })

  it('is deterministic — same input always yields same signal', () => {
    const subj = '關於凱基證券股權移轉案會議紀要'
    expect(extractSubjectSignal(subj)).toBe(extractSubjectSignal(subj))
  })

  it('signal can serve as substring match against raw email subject', () => {
    // matchSubjectKeyword does case-insensitive `includes`. Signal is
    // already prefix-stripped + lowercased, so a raw subject with reply
    // prefix matches via substring.
    const signal = extractSubjectSignal('Re: 關於凱基證券詢問')
    const rawIncoming = 'Re: Re: 關於凱基證券詢問'
    expect(rawIncoming.toLowerCase().includes(signal)).toBe(true)
  })
})

// ---- applyConfidenceCap (2026-05-27 redesign) -----------------------------
// Type-aware ceiling for auto-generated rules. Sync pull, import, restore,
// and execute-time learning all funnel through this so plain-domain rules
// can't be promoted (or arrive from an older client) above the cap and
// upset the per-type priority.
describe('applyConfidenceCap', () => {
  function makeRule(over: Partial<Rule>): Rule {
    return newRule({
      type: 'domain',
      signal: 'example.com',
      targetFolderId: 'F1',
      targetFolderPath: '03/X',
      confidence: 0.9,
      source: 'ai_confirmed',
      enabled: true,
      ...over,
    })
  }

  it('caps domain at 0.7', () => {
    const r = makeRule({ type: 'domain', confidence: 0.95 })
    expect(applyConfidenceCap(r).confidence).toBe(0.7)
  })

  it('caps sender at 0.75', () => {
    const r = makeRule({ type: 'sender', signal: 'a@gmail.com', confidence: 0.95 })
    expect(applyConfidenceCap(r).confidence).toBe(0.75)
  })

  it('caps subject_keyword at 0.9', () => {
    const r = makeRule({ type: 'subject_keyword', signal: '112訴204', confidence: 0.95 })
    expect(applyConfidenceCap(r).confidence).toBe(0.9)
  })

  it('caps compound at 0.95', () => {
    const r = makeRule({ type: 'compound', signal: encodeCompound([{ type: 'domain', value: 'x.com' }, { type: 'subject_keyword', value: 'y' }]), confidence: 1.0 })
    expect(applyConfidenceCap(r).confidence).toBe(0.95)
  })

  it('caps case_code at 0.95', () => {
    const r = makeRule({ type: 'case_code', signal: '25A0067A', confidence: 1.0 })
    expect(applyConfidenceCap(r).confidence).toBe(0.95)
  })

  it('exempts user_manual regardless of value', () => {
    const r = makeRule({ type: 'domain', source: 'user_manual', confidence: 1.0 })
    expect(applyConfidenceCap(r).confidence).toBe(1.0)
  })

  it('returns unchanged rule when already under cap', () => {
    const r = makeRule({ type: 'domain', confidence: 0.55 })
    const out = applyConfidenceCap(r)
    expect(out).toBe(r) // same reference — no new allocation
  })
})

describe('dedupeRulesByKey — user_manual sacredness (F8)', () => {
  function dup(over: Partial<Rule>): Rule {
    return newRule({
      type: 'domain',
      signal: 'dup.com',
      targetFolderId: 'F1',
      targetFolderPath: '03/Dup',
      confidence: 0.6,
      source: 'ai_confirmed',
      enabled: true,
      ...over,
    })
  }

  it('keeps a DISABLED user_manual rule over an ENABLED auto rule sharing the triple', async () => {
    // The exact bug: enabled-first sort used to drop the disabled
    // user_manual rule in favour of the enabled auto_scan one, hard-
    // deleting deliberate user intent with no tombstone.
    const manual = dup({ source: 'user_manual', enabled: false })
    const auto = dup({ source: 'auto_scan', enabled: true })
    await addRules([manual, auto])

    const report = await dedupeRulesByKey()
    expect(report.removed).toBe(1)

    const remaining = await listRules()
    expect(remaining).toHaveLength(1)
    // The survivor MUST be the user_manual rule, even though disabled.
    expect(remaining[0]!.id).toBe(manual.id)
    expect(remaining[0]!.source).toBe('user_manual')
  })

  it('among non-user_manual duplicates, still prefers the enabled one', async () => {
    const disabledAi = dup({ source: 'ai_confirmed', enabled: false })
    const enabledAuto = dup({ source: 'auto_scan', enabled: true })
    await addRules([disabledAi, enabledAuto])

    await dedupeRulesByKey()
    const remaining = await listRules()
    expect(remaining).toHaveLength(1)
    // No user_manual in the bucket → enabled-first holds.
    expect(remaining[0]!.id).toBe(enabledAuto.id)
  })

  it('among multiple user_manual duplicates, prefers the enabled one', async () => {
    const disabledManual = dup({ source: 'user_manual', enabled: false })
    const enabledManual = dup({ source: 'user_manual', enabled: true })
    await addRules([disabledManual, enabledManual])

    await dedupeRulesByKey()
    const remaining = await listRules()
    expect(remaining).toHaveLength(1)
    expect(remaining[0]!.id).toBe(enabledManual.id)
  })
})
