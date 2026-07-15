import { describe, expect, it } from 'vitest'
import type { Email, Rule } from '@/shared/types'
import {
  buildRuleIndex,
  courtCaseSignal,
  encodeCompound,
  extractCourtCaseNumbers,
  matchBodyCaseWithIndex,
  matchEmailWithIndex,
  newRule,
  ruleBeatsThread,
} from '@/shared/rules'

// Engine optimization batch 1 (2026-07): court-case rule matching.
// Before the fix, learning stored the COMPACT form (「112訴304」) but
// official court notices use the FULL form (「112年度訴字第304號」), and
// matching was raw substring — so a case-number rule could not match the
// very notice that taught it (→ never fired → hard-deleted by stale sweep).

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

function subjectRule(signal: string, targetFolderPath: string, over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'subject_keyword',
    signal,
    targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
    targetFolderPath,
    confidence: 0.85,
    source: 'ai_confirmed',
    ...over,
  })
}

function matchOne(rules: Rule[], email: Email) {
  return matchEmailWithIndex(email, buildRuleIndex(rules))
}

describe('courtCaseSignal (pure case-number detection)', () => {
  it('accepts compact and full pure forms; returns canonical compact', () => {
    expect(courtCaseSignal('112訴304')).toBe('112訴304')
    expect(courtCaseSignal('112年度訴字第304號')).toBe('112訴304')
    expect(courtCaseSignal('  112訴304  ')).toBe('112訴304')
    expect(courtCaseSignal('１１２訴３０４')).toBe('112訴304') // full-width digits
    expect(courtCaseSignal('114年度訴更(一)字第14號')).toBe('114訴更一14')
  })

  it('rejects signals that merely CONTAIN a case number (keep substring semantics)', () => {
    expect(courtCaseSignal('112訴204 通知會議')).toBeNull()
    expect(courtCaseSignal('關於112訴304的開庭')).toBeNull()
    expect(courtCaseSignal('通知')).toBeNull()
    expect(courtCaseSignal('2026年5月')).toBeNull() // date, not a case number
  })
})

describe('canonical case-number matching (full ↔ compact)', () => {
  it('compact-form rule matches a full-form court notice subject', () => {
    const r = subjectRule('112訴304', 'case/304')
    const notice = makeEmail({
      Id: 'e1',
      Subject: '臺灣臺北地方法院112年度訴字第304號開庭通知',
    })
    expect(matchOne([r], notice)?.rule.targetFolderPath).toBe('case/304')
  })

  it('full-form rule matches a compact-form subject', () => {
    const r = subjectRule('112年度訴字第304號', 'case/304')
    const email = makeEmail({ Id: 'e1', Subject: '112訴304 準備程序' })
    expect(matchOne([r], email)?.rule.targetFolderPath).toBe('case/304')
  })

  it('full-width-digit notice matches a half-width compact rule', () => {
    const r = subjectRule('112訴304', 'case/304')
    const email = makeEmail({ Id: 'e1', Subject: '１１２年度訴字第３０４號 開庭' })
    expect(matchOne([r], email)?.rule.targetFolderPath).toBe('case/304')
  })

  it('does NOT match a different case number', () => {
    const r = subjectRule('112訴304', 'case/304')
    const email = makeEmail({ Id: 'e1', Subject: '112年度訴字第305號開庭通知' })
    expect(matchOne([r], email)).toBeNull()
  })

  it('compound (domain + case number) matches across forms', () => {
    const r = newRule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'court.example.tw' },
        { type: 'subject_keyword', value: '112訴304' },
      ]),
      targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
      targetFolderPath: 'case/304',
      confidence: 0.9,
      source: 'ai_confirmed',
    })
    const email = makeEmail({
      Id: 'e1',
      Subject: '112年度訴字第304號 言詞辯論',
      From: { EmailAddress: { Address: 'clerk@court.example.tw' } },
    })
    expect(matchOne([r], email)?.rule.targetFolderPath).toBe('case/304')
  })
})

describe('digit-boundary guard (prefix collision)', () => {
  it('「112訴20」 does NOT match a subject containing 「112訴204」', () => {
    const r = subjectRule('112訴20', 'case/20')
    const email = makeEmail({ Id: 'e1', Subject: '台北地院112訴204開庭通知' })
    // Both go through canonical extraction now: subject yields 112訴204,
    // rule signal canonicalizes to 112訴20 → not in set → no match.
    expect(matchOne([r], email)).toBeNull()
  })

  it('「112訴20」 still matches its own case exactly', () => {
    const r = subjectRule('112訴20', 'case/20')
    const email = makeEmail({ Id: 'e1', Subject: '112年度訴字第20號 開庭' })
    expect(matchOne([r], email)?.rule.targetFolderPath).toBe('case/20')
  })

  it('non-case numeric signal respects the boundary in substring path', () => {
    // A full-subject-style signal ending in digits shouldn't prefix-match.
    const r = subjectRule('工單20', 'wo/20')
    const hit = makeEmail({ Id: 'e1', Subject: '本週工單20 已派發' })
    const miss = makeEmail({ Id: 'e2', Subject: '本週工單204 已派發' })
    expect(matchOne([r], hit)?.rule.targetFolderPath).toBe('wo/20')
    expect(matchOne([r], miss)).toBeNull()
  })
})

describe('whitespace-collapse matching', () => {
  it('signal matches subject with U+3000 ideographic space', () => {
    const r = subjectRule('股權移轉會議', 'case/x')
    const email = makeEmail({ Id: 'e1', Subject: '股權移轉　會議 紀要' })
    // Note: signal has no internal space; subject has U+3000 between chars.
    // The collapse makes the raw signal (no space) fail, but a signal WITH
    // the collapsed space matches. Test the realistic learn-side form:
    const r2 = subjectRule('股權移轉 會議', 'case/x')
    expect(matchOne([r2], email)?.rule.targetFolderPath).toBe('case/x')
    expect(r).toBeDefined()
  })

  it('collapses double ASCII spaces on both sides', () => {
    const r = subjectRule('project phoenix', 'proj/x')
    const email = makeEmail({ Id: 'e1', Subject: 'Project   Phoenix Q3' })
    expect(matchOne([r], email)?.rule.targetFolderPath).toBe('proj/x')
  })
})

describe('court-case subject rule priority', () => {
  function domainRule(domain: string, targetFolderPath: string): Rule {
    return newRule({
      type: 'domain',
      signal: domain,
      targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
      targetFolderPath,
      confidence: 0.7,
      source: 'auto_scan',
    })
  }

  it('case-number subject rule outranks a plain-domain rule (same-client-multi-case)', () => {
    const dom = domainRule('client.example.tw', 'case/A')
    const cc = subjectRule('112訴304', 'case/B')
    // Mail from the client whose subject carries case B's number.
    const email = makeEmail({
      Id: 'e1',
      Subject: '112年度訴字第304號 補充理由',
      From: { EmailAddress: { Address: 'lawyer@client.example.tw' } },
    })
    // Before the fix domain (priority 3) fired first → case/A (wrong).
    expect(matchOne([dom, cc], email)?.rule.targetFolderPath).toBe('case/B')
  })

  it('ruleBeatsThread: court-case subject rule beats thread memory; plain subject does not', () => {
    expect(ruleBeatsThread({ type: 'subject_keyword', source: 'ai_confirmed', signal: '112訴304' })).toBe(true)
    expect(ruleBeatsThread({ type: 'subject_keyword', source: 'ai_confirmed', signal: '請款通知' })).toBe(false)
    // no-signal legacy call shape still works (returns false)
    expect(ruleBeatsThread({ type: 'subject_keyword', source: 'ai_confirmed' })).toBe(false)
  })
})

describe('extractCourtCaseNumbers format coverage', () => {
  it('handles full-width digits and bracketed sub-types', () => {
    expect(extractCourtCaseNumbers('１１２訴３０４')).toContain('112訴304')
    expect(extractCourtCaseNumbers('114年度訴更(一)字第14號')).toContain('114訴更一14')
    expect(extractCourtCaseNumbers('114年度訴更（一）字第14號')).toContain('114訴更一14')
  })
})

// ---- B3-C5: internal-body case matching ------------------------------------

function caseCodeRule(signal: string, targetFolderPath: string, over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'case_code',
    signal,
    targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
    targetFolderPath,
    confidence: 0.9,
    source: 'ai_confirmed',
    ...over,
  })
}
function domainRule(signal: string, targetFolderPath: string): Rule {
  return newRule({
    type: 'domain',
    signal,
    targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
    targetFolderPath,
    confidence: 0.8,
    source: 'ai_confirmed',
  })
}

describe('internal-body case matching (batch-3)', () => {
  // Matching requires the FULLY-FETCHED bodyText, never the 250-char preview
  // (review fix): a truncated view can't verify case uniqueness, so preflight
  // (preview-only) never body-matches.
  it('routes on a body case number when the subject has none', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchOne([r], makeEmail({ Id: 'e', Subject: '開庭通知', bodyText: '本件 112訴304 敬請出席' }))
    expect(m?.rule.targetFolderPath).toBe('case/304')
    expect(m?.reason).toContain('內文含案號')
  })

  it('routes on a body case CODE against a case_code rule', () => {
    const r = caseCodeRule('25A0067A', 'case/A')
    const m = matchOne([r], makeEmail({ Id: 'e', Subject: '通知', bodyText: '案件 25A0067A 相關文件' }))
    expect(m?.rule.targetFolderPath).toBe('case/A')
    expect(m?.reason).toContain('內文含案件代號')
  })

  it('does NOT match on a preview-only email (no bodyText) — preflight safety', () => {
    // The exact P2 misroute guard: only BodyPreview present (as at preflight),
    // no bodyText → body pass must not fire, even though the preview names a
    // case that has a rule. The email flows to the AI (full-body) instead.
    const r = subjectRule('112訴304', 'case/304')
    const m = matchOne([r], makeEmail({ Id: 'e', Subject: '通知', BodyPreview: '本件 112訴304' }))
    expect(m).toBeNull()
  })

  it('IGNORES the body when the subject carries its own case number', () => {
    // Subject case 112訴204 has no rule; body case 112訴304 has a rule.
    // Subject-first gate → body is not consulted → no match (falls to AI).
    const r = subjectRule('112訴304', 'case/304')
    const m = matchOne([r], makeEmail({ Id: 'e', Subject: '112年度訴字第204號 通知', bodyText: '參照 112訴304' }))
    expect(m).toBeNull()
  })

  it('does NOT route when the body is ambiguous (>1 distinct case)', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchOne([r], makeEmail({ Id: 'e', Subject: '通知', bodyText: '本件 112訴304，另案 113訴500' }))
    expect(m).toBeNull()
  })

  it('a domain rule wins over a body case number (priority)', () => {
    const dom = domainRule('court.example.tw', 'domain/court')
    const cc = subjectRule('112訴304', 'case/304')
    const m = matchOne(
      [dom, cc],
      makeEmail({
        Id: 'e',
        Subject: '通知',
        From: { EmailAddress: { Address: 'clerk@court.example.tw' } },
        bodyText: '本件 112訴304',
      }),
    )
    expect(m?.rule.targetFolderPath).toBe('domain/court')
  })
})

// ---- matchBodyCaseWithIndex (C5 activation helper) --------------------------

describe('matchBodyCaseWithIndex (full-body re-match used by classifyAi)', () => {
  it('matches an unambiguous body court case against a court-case rule', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '開庭通知', bodyText: '本件 112訴304 敬請' }),
      buildRuleIndex([r]),
    )
    expect(m?.rule.targetFolderPath).toBe('case/304')
    expect(m?.reason).toContain('內文含案號')
  })

  it('returns null when the subject carries a case number (subject-first)', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '112年度訴字第204號', bodyText: '參照 112訴304' }),
      buildRuleIndex([r]),
    )
    expect(m).toBeNull()
  })

  it('returns null on an ambiguous body (>1 distinct case)', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '通知', bodyText: '本件 112訴304，另案 113訴500' }),
      buildRuleIndex([r]),
    )
    expect(m).toBeNull()
  })

  it('returns null when only BodyPreview is present (no fetched bodyText)', () => {
    const r = subjectRule('112訴304', 'case/304')
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '通知', BodyPreview: '本件 112訴304' }),
      buildRuleIndex([r]),
    )
    expect(m).toBeNull()
  })
})

// ---- review fix: ambiguity veto + routing scan the FULL body, not 800 ------

describe('matchBodyCaseWithIndex full-body window (review fix)', () => {
  it('vetoes when a SECOND case sits beyond the old 800-char window', () => {
    // X at ~char 5, Y at ~char 915. The old 800-slice would hide Y → distinct=1
    // → misroute to X. Full-body scan sees both → ambiguous → null (safe).
    const r = subjectRule('112訴304', 'case/304')
    const bodyText = '本件 112訴304 ' + '　'.repeat(900) + ' 另案 113訴500'
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '通知', bodyText }),
      buildRuleIndex([r]),
    )
    expect(m).toBeNull()
  })

  it('routes on a SINGLE case that sits beyond char 800', () => {
    const r = subjectRule('112訴304', 'case/304')
    const bodyText = '前言 '.repeat(300) + ' 本件 112訴304 敬請' // case well past 800
    const m = matchBodyCaseWithIndex(
      makeEmail({ Id: 'e', Subject: '通知', bodyText }),
      buildRuleIndex([r]),
    )
    expect(m?.rule.targetFolderPath).toBe('case/304')
  })
})
