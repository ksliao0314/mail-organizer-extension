import { describe, expect, it } from 'vitest'
import { caseSignalsForMatch, courtCaseSignal, extractCourtCaseNumbers, normalizeSignal } from '@/shared/rules'
import type { Email } from '@/shared/types'

describe('extractCourtCaseNumbers — compact form', () => {
  it('extracts 112訴204', () => {
    expect(extractCourtCaseNumbers('Re: 112訴204 開庭通知')).toEqual(['112訴204'])
  })

  it('extracts 114訴更一14 (with sub-class 更一)', () => {
    expect(extractCourtCaseNumbers('114訴更一14 補件')).toEqual(['114訴更一14'])
  })

  it('extracts 114民著訴74 (民著訴 type)', () => {
    expect(extractCourtCaseNumbers('Fwd: 114民著訴74 庭期變更')).toEqual([
      '114民著訴74',
    ])
  })

  it('extracts multiple case numbers in one subject', () => {
    // 113上字第50號 normalizes to 113上50 (strip 字 and 第).
    const out = extractCourtCaseNumbers('併案處理 112訴204 與 113上字第50號')
    expect(out).toContain('112訴204')
    expect(out).toContain('113上50')
  })
})

describe('extractCourtCaseNumbers — full form', () => {
  it('extracts 112年度訴字第204號 and normalizes to compact', () => {
    expect(extractCourtCaseNumbers('民國112年度訴字第204號')).toEqual(['112訴204'])
  })

  it('handles missing 度', () => {
    expect(extractCourtCaseNumbers('112年訴字第204號')).toEqual(['112訴204'])
  })

  // 2026-11-06: this used to assert that a full form WITHOUT 字 parses too
  // (「112年度訴第204」). 字 is now required before 第 — every real court
  // reference has it, and without it 「113年度財報第1季」 was a case. 號 stays
  // optional.
  it('handles a missing 號; 字 is required', () => {
    expect(extractCourtCaseNumbers('112年度訴字第204')).toEqual(['112訴204'])
    expect(extractCourtCaseNumbers('112年度訴第204')).toEqual([])
  })

  it('dedupes when same case appears in both forms', () => {
    const out = extractCourtCaseNumbers('案號 112訴204、即 112年度訴字第204號')
    expect(out).toEqual(['112訴204'])
  })
})

describe('extractCourtCaseNumbers — false positive rejection', () => {
  it('rejects "112年5月" (date)', () => {
    expect(extractCourtCaseNumbers('開庭日: 112年5月')).toEqual([])
  })

  it('rejects "112年5月1日"', () => {
    expect(extractCourtCaseNumbers('庭期 112年5月1日 上午')).toEqual([])
  })

  it('rejects 4-digit western year', () => {
    expect(extractCourtCaseNumbers('2025年訴字第10號')).toEqual([])
  })

  it('does not match a longer digit sequence (no 1XX boundary)', () => {
    // 1234 should not be split into "123" + ...
    expect(extractCourtCaseNumbers('編號 1234訴999')).toEqual([])
  })

  it('returns empty for empty / nullish input', () => {
    expect(extractCourtCaseNumbers('')).toEqual([])
  })

  it('does not extract from plain English subjects', () => {
    expect(extractCourtCaseNumbers('Hello world, please confirm')).toEqual([])
  })
})

// 2026-11-06: with no real 字別 after 年度, the full form backtracked out of
// its optional 度 and took 度 itself as the case type — 「113年度第1次董事會
// 議程」 became court case 「113度1」, and the first AI filing of it learned a
// case rule that swallowed every 「113年度第1…」 mail of that client.
describe('年度 + 第N次 / 期 / 號 is not a court case (度 is never a case type)', () => {
  const nonCases = [
    '113年度第1次董事會議程',
    '113年度第1期財報',
    '113年度第1號公告',
    '113年度第1次股東會',
    'RE: 113年度第２次臨時股東會 議事錄',
    '113度1',
  ]

  it.each(nonCases)('extracts nothing from 「%s」 (was 「113度…」)', (text) => {
    expect(extractCourtCaseNumbers(text)).toEqual([])
  })

  it('courtCaseSignal / normalizeSignal no longer canonicalize them to a case number', () => {
    expect(courtCaseSignal('113年度第1號')).toBeNull() // was 113度1
    expect(courtCaseSignal('113度1')).toBeNull()
    expect(normalizeSignal('subject_keyword', '113年度第1號')).toBe('113年度第1號')
  })

  it('the body case gate finds no case in them either', () => {
    const body = caseSignalsForMatch({
      Id: 'e',
      Subject: '會議通知',
      BodyPreview: '',
      bodyText: '茲訂於下週召開 113年度第1次董事會，議程如附件。',
    } as Email)
    expect(body.courtCases).toEqual([])
  })

  it.each([
    ['112年度訴字第304號', '112訴304'],
    ['112訴304', '112訴304'],
    ['113年度上字第5號', '113上5'],
    ['113年上字第5號', '113上5'],
    ['114年度訴更一字第14號', '114訴更一14'],
    ['１１２年度重訴字第１２號', '112重訴12'],
  ])('real case numbers still parse: 「%s」 → %s', (text, compact) => {
    expect(extractCourtCaseNumbers(`主旨 ${text} 開庭通知`)).toEqual([compact])
    expect(courtCaseSignal(text)).toBe(compact)
  })

  it('a real case next to a 年度 meeting number is still found', () => {
    expect(extractCourtCaseNumbers('113年度第1次會議 討論 113年度訴字第5號')).toEqual(['113訴5'])
  })
})

// 2026-11-06 (follow-up): the 度 fix closed one backtrack path; the class was
// the optional 字 — 「1XX年度<any words>第N」 parsed as a case (「113年度財報
// 第1季」 → 113財報1, 「113年度上午第1次」 → 113上午1). The full form now needs
// 字 right before 第 (whitespace allowed), in text and anchored alike.
describe('the full form needs 字 right before 第', () => {
  const nonCases = [
    '113年度財報第1季',
    '113年度上午第1次',
    '113年度第1次董事會',
    '113年度預算第2次追加',
    '113年度員工旅遊第3梯次',
    '１１３年度財報第１季',
    '113年度財報第1號', // the whole string — the anchored form canonicalized it to 113財報1
  ]

  it.each(nonCases)('「%s」 is no case — extraction, courtCaseSignal, normalizeSignal', (text) => {
    expect(extractCourtCaseNumbers(text)).toEqual([]) // was 113財報1 / 113上午1 / …
    expect(courtCaseSignal(text)).toBeNull()
    expect(normalizeSignal('subject_keyword', text)).not.toMatch(/^1\d{2}[^\d]+\d+$/)
  })

  it('nor in a body', () => {
    const body = caseSignalsForMatch({
      Id: 'e',
      Subject: '會議通知',
      BodyPreview: '',
      bodyText: '訂於 113年度上午第1次 全體會議討論。',
    } as Email)
    expect(body.courtCases).toEqual([])
  })

  it.each([
    ['112年度訴字第304號', '112訴304'],
    ['113年上字第5號', '113上5'],
    ['114年度訴更一字第14號', '114訴更一14'],
    ['114年度訴更（一）字第14號', '114訴更一14'],
    ['最高法院112年度台上字第123號', '112台上123'],
    ['１１２年度訴字第３０４號', '112訴304'],
    ['１１３年上字第５號', '113上5'],
    ['112年度訴字 第304號', '112訴304'],
    ['112年度訴字第304 開庭通知', '112訴304'],
    ['108年度家財訴字第7號', '108家財訴7'],
    ['103年度司執季字第38824號', '103司執季38824'],
  ])('a real reference still parses: 「%s」 → %s', (text, compact) => {
    expect(extractCourtCaseNumbers(`主旨 ${text}`)).toEqual([compact])
  })

  it('anchored too: courtCaseSignal / normalizeSignal keep canonicalizing real full forms', () => {
    expect(courtCaseSignal('最高法院112年度台上字第123號')).toBeNull() // a prefix: not the whole signal
    expect(courtCaseSignal('112年度台上字第123號')).toBe('112台上123')
    expect(courtCaseSignal('112年度訴字 第304號')).toBe('112訴304')
    expect(normalizeSignal('subject_keyword', '１１２年度訴字第３０４號')).toBe('112訴304')
  })

  it('compact forms are unchanged — including ones whose type shares a character with plain words', () => {
    expect(extractCourtCaseNumbers('112訴204 開庭')).toEqual(['112訴204'])
    expect(extractCourtCaseNumbers('108家財訴7 調解期日')).toEqual(['108家財訴7'])
    expect(courtCaseSignal('113財報1')).toBe('113財報1') // no character blacklist (see rules.ts)
  })

  it('the earlier 度 fixes still hold', () => {
    for (const text of ['113年度第1次董事會議程', '113年度第1期財報', '113年度第1號公告', '113年度字第5號', '113度1']) {
      expect(extractCourtCaseNumbers(text)).toEqual([])
      expect(courtCaseSignal(text)).toBeNull()
    }
  })
})

// extractCandidateSubjectTokens removed in 2026-05-27 redesign — tokenized
// subject_keyword rules conflicted with the 整段主旨 design. Tests deleted
// alongside the function.
