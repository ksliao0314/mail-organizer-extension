// B3-C6: learning a rule from a case number found in the email BODY.
// The conservative gates (subject-first, distinct===1, tightened source gate,
// lower confidence, NO compound) all live here + in the shared caseSignals
// gate. The closed-loop guarantee: what C5's body pass can MATCH, C6 LEARNS —
// so a body-learned rule can actually fire (never stale-swept for not firing).

import { beforeEach, describe, expect, it } from 'vitest'
import { generateAiConfirmedRules, generateAiOverrideRules } from '@/background/execute'
import { listRules } from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { PlanItem } from '@/shared/types'

beforeEach(async () => {
  await setRules([])
})

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: '開庭通知', // no case number in the subject
    emailFrom: 'clerk@court.example.tw',
    action: 'move',
    targetFolderPath: '03/案件/甲',
    confidence: 0.9,
    reason: '',
    source: 'ai', // AI-accepted (not overridden)
    ...over,
  }
}
const moved = (emailId: string) => ({
  emailId,
  subject: 'X',
  action: 'move' as const,
  status: 'moved' as const,
})

describe('generateAiConfirmedRules — body case learning', () => {
  it('learns a court case from the body as a PURE subject_keyword rule (never compound)', async () => {
    // Even though there is a usable domain, a body case must NOT become a
    // compound (matchCompound reads only the subject → would never fire).
    const item = planItem({ bodyCaseNumbers: ['112訴204'] })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.type).toBe('subject_keyword')
    expect(rules[0]!.signal).toBe('112訴204')
  })

  it('docks confidence by 0.1 for a body-derived case (0.85 → 0.75)', async () => {
    const item = planItem({ bodyCaseNumbers: ['112訴204'] })
    await generateAiConfirmedRules([item], [moved(item.emailId)])
    const rules = await listRules()
    expect(rules[0]!.confidence).toBeCloseTo(0.75, 6)
  })

  it('learns a body case CODE as a case_code rule at 0.8 (0.9 − 0.1)', async () => {
    const item = planItem({ emailSubject: '補件', bodyCaseCodes: ['25A0067A'] })
    await generateAiConfirmedRules([item], [moved(item.emailId)])
    const rules = await listRules()
    expect(rules[0]!.type).toBe('case_code')
    expect(rules[0]!.confidence).toBeCloseTo(0.8, 6)
  })

  it('refuses to learn when the body is ambiguous (>1 distinct case)', async () => {
    const item = planItem({ bodyCaseNumbers: ['112訴204', '113訴500'] })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    // No case signal survives the gate → falls through to domain, which is a
    // court.example.tw (usable) domain → a domain rule MIGHT form. Assert the
    // key property: no court-case rule for either ambiguous number.
    const rules = await listRules()
    expect(rules.some((x) => x.signal === '112訴204' || x.signal === '113訴500')).toBe(false)
    expect(r.added).toBeLessThanOrEqual(1)
  })

  it('does NOT let a thread/rule-routed item learn from a BODY case (source gate)', async () => {
    // A subject case can mint from any route (structural). A body case is only
    // trusted from ai / unresolved — not thread/rule.
    const item = planItem({ source: 'thread', bodyCaseNumbers: ['112訴204'] })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
  })

  it('regression: a SUBJECT case (+ usable domain) still learns a compound at 0.9, body ignored', async () => {
    const item = planItem({ emailSubject: '112年度訴字第204號 開庭', bodyCaseNumbers: ['113訴999'] })
    await generateAiConfirmedRules([item], [moved(item.emailId)])
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    // subject-first + usable domain → compound on the SUBJECT case; body ignored.
    expect(rules[0]!.type).toBe('compound')
    expect(rules[0]!.signal).toContain('112訴204')
    expect(rules[0]!.signal).not.toContain('113訴999')
    // compound court_case, no body penalty → unchanged 0.9
    expect(rules[0]!.confidence).toBeCloseTo(0.9, 6)
  })
})

describe('generateAiOverrideRules — body case learning', () => {
  it('mints a body-derived override rule at 0.85 (vs 0.95 for subject)', async () => {
    const item = planItem({
      source: 'ai',
      emailSubject: '通知',
      bodyCaseCodes: ['25A0067A'],
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('case_code')
    expect(rules[0]!.confidence).toBeCloseTo(0.85, 6)
  })
})
