import { describe, expect, it } from 'vitest'
import { extractCourtCaseNumbers } from '@/shared/rules'

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

  it('handles missing 字 and 號', () => {
    expect(extractCourtCaseNumbers('112年度訴第204')).toEqual(['112訴204'])
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

// extractCandidateSubjectTokens removed in 2026-05-27 redesign — tokenized
// subject_keyword rules conflicted with the 整段主旨 design. Tests deleted
// alongside the function.
