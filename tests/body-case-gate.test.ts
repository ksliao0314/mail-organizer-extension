import { describe, expect, it } from 'vitest'
import {
  caseSignalsForLearning,
  caseSignalsForMatch,
  gateCaseSignals,
} from '@/shared/rules'
import type { Email, PlanItem } from '@/shared/types'

// The shared body-case gate is the single alignment point between the matching
// side (rules) and the learning side (execute). These lock the conservative
// policy: subject-first, body-fallback, structural-only, ambiguity-flagged.

describe('gateCaseSignals', () => {
  it('subject-first: a subject case number makes the body ignored entirely', () => {
    const r = gateCaseSignals(['112訴204'], [], ['113訴500'], ['25A0067A'])
    expect(r.source).toBe('subject')
    expect(r.courtCases).toEqual(['112訴204'])
    expect(r.caseCodes).toEqual([]) // subject codes, NOT the body's 25A0067A
    expect(r.bodyAmbiguous).toBe(false)
  })

  it('falls back to a single body identifier when the subject has none', () => {
    const r = gateCaseSignals([], [], ['112訴204'], [])
    expect(r.source).toBe('body')
    expect(r.courtCases).toEqual(['112訴204'])
    expect(r.bodyAmbiguous).toBe(false)
  })

  it('flags ambiguity when the body carries >1 distinct identifier', () => {
    const r = gateCaseSignals([], [], ['112訴204', '113訴500'], [])
    expect(r.source).toBe('body')
    expect(r.bodyAmbiguous).toBe(true)
  })

  it('a court number + a case code in the body are 2 distinct → ambiguous', () => {
    const r = gateCaseSignals([], [], ['112訴204'], ['25A0067A'])
    expect(r.bodyAmbiguous).toBe(true)
  })

  it('source=none when neither subject nor body has a case identifier', () => {
    const r = gateCaseSignals([], [], [], [])
    expect(r.source).toBe('none')
    expect(r.bodyAmbiguous).toBe(false)
  })
})

function email(over: Partial<Email>): Email {
  return { Id: 'e', Subject: '', BodyPreview: '', From: { EmailAddress: { Address: 'a@b.com' } }, ToRecipients: [], ReceivedDateTime: '', ParentFolderId: 'i', ...over } as Email
}

describe('caseSignalsForMatch', () => {
  it('prefers bodyText (800) over BodyPreview (250) when present', () => {
    const withText = caseSignalsForMatch(
      email({ Subject: '開庭通知', bodyText: '詳如 112訴204 所示', BodyPreview: '(截斷前段)' }),
    )
    expect(withText.source).toBe('body')
    expect(withText.courtCases).toEqual(['112訴204'])
  })

  it('does NOT read BodyPreview (matching requires the full bodyText)', () => {
    // Preview-only (as at preflight) → body ignored → no body-case routing.
    const r = caseSignalsForMatch(email({ Subject: '開庭通知', BodyPreview: '本件 112訴204' }))
    expect(r.source).toBe('none')
  })

  it('subject case number wins even if the body cites another case', () => {
    const r = caseSignalsForMatch(
      email({ Subject: '112年度訴字第204號 通知', bodyText: '參照另案 113訴500' }),
    )
    expect(r.source).toBe('subject')
    expect(r.courtCases).toEqual(['112訴204'])
  })
})

function planItem(over: Partial<PlanItem>): PlanItem {
  return { emailId: 'e', emailSubject: '', emailFrom: 'a@b.com', action: 'skip', confidence: 0, reason: '', source: 'ai', ...over }
}

describe('caseSignalsForLearning', () => {
  it('uses the pre-computed body identifiers from classify time', () => {
    const r = caseSignalsForLearning(
      planItem({ emailSubject: '開庭', bodyCaseNumbers: ['112訴204'], bodyCaseCodes: [] }),
    )
    expect(r.source).toBe('body')
    expect(r.courtCases).toEqual(['112訴204'])
  })

  it('does NOT re-extract from bodyPreview — absent pre-computed fields ⇒ no body case', () => {
    // Symmetry with matching: learning only trusts identifiers pre-computed
    // from the full body (bodyCaseNumbers), never the 250-char preview.
    const r = caseSignalsForLearning(planItem({ emailSubject: '開庭', bodyPreview: '本件 112訴204 敬請' }))
    expect(r.source).toBe('none')
  })

  it('refuses ambiguity: 2 distinct pre-computed body numbers → bodyAmbiguous', () => {
    const r = caseSignalsForLearning(
      planItem({ emailSubject: '開庭', bodyCaseNumbers: ['112訴204', '113訴500'], bodyCaseCodes: [] }),
    )
    expect(r.bodyAmbiguous).toBe(true)
  })
})

describe('closed-loop symmetry (match ⇔ learn on the same email)', () => {
  it('match and learn agree on the eligible signals for the same content', () => {
    const subject = '開庭通知'
    const body = '本件 112訴204 敬請出席'
    const m = caseSignalsForMatch(email({ Subject: subject, bodyText: body }))
    // Learn side receives the body identifiers pre-computed from the SAME body.
    const l = caseSignalsForLearning(
      planItem({ emailSubject: subject, bodyCaseNumbers: ['112訴204'], bodyCaseCodes: [] }),
    )
    expect(l.source).toBe(m.source)
    expect(l.courtCases).toEqual(m.courtCases)
    expect(l.caseCodes).toEqual(m.caseCodes)
    expect(l.bodyAmbiguous).toBe(m.bodyAmbiguous)
  })

  it('both ignore the body identically when the subject carries a case', () => {
    const subject = '112年度訴字第204號'
    const body = '另案 113訴500'
    const m = caseSignalsForMatch(email({ Subject: subject, bodyText: body }))
    const l = caseSignalsForLearning(
      planItem({ emailSubject: subject, bodyCaseNumbers: ['113訴500'], bodyCaseCodes: [] }),
    )
    expect(m.source).toBe('subject')
    expect(l.source).toBe('subject')
    expect(l.courtCases).toEqual(m.courtCases)
  })
})
