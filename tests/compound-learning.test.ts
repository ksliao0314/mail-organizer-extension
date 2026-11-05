import { beforeEach, describe, expect, it } from 'vitest'
import {
  generateAiOverrideRules,
  type ExecuteItemResult,
} from '@/background/execute'
import { decodeCompound, listRules } from '@/shared/rules'
import { setRules, setSettings } from '@/shared/storage'
import type { Rule } from '@/shared/types'
import type { PlanItem } from '@/shared/types'

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: '',
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

const okResult = (emailId: string): ExecuteItemResult => ({
  emailId,
  subject: 'X',
  action: 'move',
  status: 'moved',
})

describe('AI generators — domain + Taiwan court case → compound rule', () => {
  beforeEach(async () => {
    // Pre-G1 the internal domain was the hardcoded constant 'example.com'.
    // After G1 generification, it lives in settings.internalDomains — empty by
    // default. Tests below were written against the old default, so we seed
    // it here to preserve their assertions.
    await setSettings({ internalDomains: ['example.com'] })
  })

  it('produces a compound rule (domain + courtCase) when both signals present', async () => {
    const item = planItem({
      emailSubject: 'Re: 112訴204 開庭通知',
      emailFrom: 'litigation@telco.example',
      action: 'move',
      targetFolderPath: '03/民事/乙併購案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/民事/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.type).toBe('compound')
    const parsed = decodeCompound(rules[0]!.signal)
    expect(parsed?.conditions).toEqual([
      { type: 'domain', value: 'telco.example' },
      { type: 'subject_keyword', value: '112訴204' },
    ])
    expect(rules[0]!.confidence).toBe(0.95)
    expect(rules[0]!.source).toBe('ai_overridden')
  })

  it('handles 114訴更一14 (sub-classified case)', async () => {
    const item = planItem({
      emailSubject: '114訴更一14 補件',
      emailFrom: 'court@agency.example',
      action: 'move',
      targetFolderPath: '03/行政/丙行政訴訟案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/行政/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('compound')
    const parsed = decodeCompound(rules[0]!.signal)
    expect(parsed?.conditions[1]?.value).toBe('114訴更一14')
  })

  it('canonicalizes full form 112年度訴字第204號 to compact in compound rule', async () => {
    const item = planItem({
      emailSubject: '請見 112年度訴字第204號 附件',
      emailFrom: 'a@telco.example',
      action: 'move',
      targetFolderPath: '03/X',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const parsed = decodeCompound((await listRules())[0]!.signal)
    expect(parsed?.conditions[1]?.value).toBe('112訴204')
  })

  it('falls back to subject_keyword when court case present but no usable domain', async () => {
    const item = planItem({
      emailSubject: '114民著訴74 期日通知',
      emailFrom: '',
      action: 'move',
      targetFolderPath: '03/民事/CompanyB',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/民事/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('subject_keyword')
    expect(rules[0]!.signal).toBe('114民著訴74')
  })

  it('skips internal domain even when court case is present', async () => {
    const item = planItem({
      emailSubject: '112訴204 內部討論',
      emailFrom: 'colleague@example.com',
      action: 'move',
      targetFolderPath: '03/民事/X 案',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/民事/Y 案',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    // internal domain blocked + court case alone → subject_keyword
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('subject_keyword')
    expect(rules[0]!.signal).toBe('112訴204')
  })

  it('still does NOT create plain-domain rule on user override (no-conflict case → no rule)', async () => {
    // "先廣後窄" v2 (2026-05-27): override + no existing same-domain
    // rule = nothing is learned. A single override is too weak to
    // build a broad domain rule (one override doesn't justify routing
    // every email from this domain), and there's no conflict signal
    // suggesting the domain rule is "insufficient". Next time the
    // user CONFIRMS an AI suggestion for this domain, the plain-
    // domain rule will be born.
    const item = planItem({
      emailSubject: '請款通知',
      emailFrom: 'billing@vendor.com',
      action: 'move',
      targetFolderPath: '05/發票',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '02/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    const rules = await listRules()
    // The original contract: no plain-domain rule from a single override.
    expect(rules.filter((rl) => rl.type === 'domain')).toHaveLength(0)
    // The new contract: nothing learned at all when there's no conflict.
    expect(r.added).toBe(0)
    expect(rules).toHaveLength(0)
  })

  it('upgrades to compound (domain + 整段主旨) when override conflicts with existing plain-domain rule', async () => {
    // Conflict scenario: existing `domain: vendor.com → 02/未分類`
    // rule, user just overrode AI's vendor.com routing to 05/發票.
    // That's the "網域規則已經不夠了，需要升級" signal — chooseLearningSignal
    // detects the conflict and builds compound (domain + 整段主旨)
    // for the new target; the demote loop in generateAiOverrideRules
    // disables the stale plain-domain rule.
    const seed: Rule = {
      id: 'seed-vendor',
      type: 'domain',
      signal: 'vendor.com',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '02/未分類',
      confidence: 0.55,
      matchCount: 5,
      enabled: true,
      createdAt: new Date().toISOString(),
      source: 'ai_confirmed',
    }
    await setRules([seed])

    // ≥ 8 chars: shorter subjects are no longer learnable (2026-11-05, see
    // the demote-only companion below).
    const item = planItem({
      emailSubject: '五月份請款通知單據',
      emailFrom: 'billing@vendor.com',
      action: 'move',
      targetFolderPath: '05/發票',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '02/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(1)
    expect(r.disabled).toBe(1) // stale plain-domain demoted

    const rules = await listRules()
    const compound = rules.find((rl) => rl.type === 'compound')
    expect(compound).toBeDefined()
    expect(compound!.targetFolderPath).toBe('05/發票')
    const parsed = decodeCompound(compound!.signal)
    expect(parsed?.conditions).toEqual([
      { type: 'domain', value: 'vendor.com' },
      { type: 'subject_keyword', value: '五月份請款通知單據' },
    ])
    // Old plain-domain rule still exists but is disabled — as an AUTO
    // disable ('superseded', 2026-11-05) so it can revive / be purged
    // later instead of reading as user-disabled forever.
    const oldDomainRule = rules.find((rl) => rl.id === seed.id)
    expect(oldDomainRule?.enabled).toBe(false)
    expect(oldDomainRule?.autoDisabledReason).toBe('superseded')
    expect(oldDomainRule?.autoDisabledAt).toBeDefined()
  })

  it('short subject (< 8 chars) on a conflict: demote the broad rule, learn nothing (2026-11-05)', async () => {
    await setRules([
      {
        id: 'seed-vendor-2',
        type: 'domain',
        signal: 'vendor.com',
        targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
        targetFolderPath: '02/未分類',
        confidence: 0.55,
        matchCount: 5,
        enabled: true,
        createdAt: '2026-01-01T00:00:00.000Z',
        source: 'ai_confirmed',
      },
    ])
    const item = planItem({
      emailSubject: '請款通知', // 4 chars — would have minted vendor.com + 請款通知
      emailFrom: 'billing@vendor.com',
      action: 'move',
      targetFolderPath: '05/發票',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '02/未分類',
    })
    const r = await generateAiOverrideRules([item], [okResult(item.emailId)])
    expect(r.added).toBe(0)
    expect(r.disabled).toBe(1)
    const [seed] = await listRules()
    expect(seed!.enabled).toBe(false)
    // Nothing narrower replaced it — the user's correction silenced it.
    expect(seed!.autoDisabledReason).toBe('override')
  })
})
