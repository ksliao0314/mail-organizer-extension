import { describe, expect, it } from 'vitest'
import { generateAiOverrideRules } from '@/background/execute'
import { listRules } from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { PlanItem, Rule } from '@/shared/types'

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: '請款通知',
    emailFrom: 'alice@example.com',
    action: 'move',
    targetFolderPath: '03/Right',
    confidence: 0.7,
    reason: '',
    source: 'ai',
    aiOriginalAction: 'move',
    aiOriginalTargetFolderPath: '03/Wrong',
    ...over,
  }
}

function rule(over: Partial<Rule>): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/X',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'ai_confirmed',
    ...over,
  }
}

const okResult = (emailId: string) => ({
  emailId,
  subject: 'X',
  action: 'move' as const,
  status: 'moved' as const,
})

describe('AI generators prefer case_code over domain when subject has one', () => {
  it('generates a case_code rule when subject contains a case code', async () => {
    const item = planItem({
      emailSubject: 'Re: 25A0067A 補件通知',
      emailFrom: 'court@gov.example',
      action: 'move',
      targetFolderPath: '03/民事/X 案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/民事/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.type).toBe('case_code')
    expect(rules[0]!.signal).toBe('25A0067A')
    expect(rules[0]!.confidence).toBe(0.95)
    expect(rules[0]!.source).toBe('ai_overridden')
  })

  it('does NOT create a plain-domain rule when user overrides a single email — even with no case code', async () => {
    // Design choice: same domain often serves multiple cases at this firm.
    // A single user override isn't enough evidence to reroute every email
    // from this domain. Plain-domain rules need stronger evidence
    // (initial-scan cross-folder uniqueness, or multiple ai_confirmed).
    //
    // Post-2026-05-22: subject feature extraction may pull out a
    // generic noun phrase like "一般通知" and create a compound
    // (domain + feature) rule. The original concern was specifically
    // about plain-DOMAIN rules; compound rules with a feature still
    // require the feature to match, so they're MUCH narrower than a
    // domain-only rule. errorRate gate cleans up bad features.
    const item = planItem({
      emailSubject: 'Re: 一般通知',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    const rules = await listRules()
    // Original contract: no plain-domain rule from one override.
    expect(rules.filter((rl) => rl.type === 'domain')).toHaveLength(0)
    // New behavior: compound (domain + feature) may exist. The feature
    // heuristic could even fail to extract anything, in which case
    // r.added stays 0 — both outcomes are acceptable here.
    expect(r.added).toBeLessThanOrEqual(1)
  })

  it('disables only same-type same-signal sibling rules', async () => {
    // Existing case_code rule pointing at wrong target — should be disabled
    await setRules([
      rule({
        type: 'case_code',
        signal: '25A0067A',
        targetFolderPath: '03/民事/未分類',
        source: 'ai_confirmed',
      }),
      // Domain rule for same client — should NOT be disabled (different type)
      rule({
        type: 'domain',
        signal: 'court.example',
        targetFolderPath: '03/民事/未分類',
        source: 'auto_scan',
      }),
    ])
    const item = planItem({
      emailSubject: '25A0067A 開庭通知',
      emailFrom: 'court@court.example',
      action: 'move',
      targetFolderPath: '03/民事/X 案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/民事/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    expect(r.disabled).toBe(1)
    const rules = await listRules()
    const caseRule = rules.find((x) => x.type === 'case_code' && x.signal === '25A0067A' && x.targetFolderPath === '03/民事/未分類')
    expect(caseRule?.enabled).toBe(false) // disabled
    const domainRule = rules.find((x) => x.type === 'domain' && x.signal === 'court.example')
    expect(domainRule?.enabled).toBe(true) // untouched — different type
  })

  it('uses uppercase canonical case code regardless of subject casing', async () => {
    const item = planItem({
      emailSubject: 'Re: 25a0067a 補件',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/X 案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.signal).toBe('25A0067A')
  })
})
