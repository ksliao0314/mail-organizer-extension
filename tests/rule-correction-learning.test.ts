// Correcting a RULE-routed row is learned (2026-11-05, item 3).
//
// PlanRow flips a rule-routed row's source 'rule'→'ai' when the user edits
// it, keeping originalRuleId but never setting aiOriginalAction (the rule,
// not the AI, produced the verdict). So wasUserOverride() was false: the
// birth gate treated the correction as a passive AI confirmation and held it
// back, and generateAiOverrideRules skipped it — the broad rule was never
// silenced. Repro: correcting client.example→A to B → added 0, disabled 0.
// Fixed on the learning side: originalRuleId + userTouched + a final target
// different from the originating rule's = explicit correction → override path.

import { describe, expect, it } from 'vitest'
import {
  generateAiConfirmedRules,
  generateAiOverrideRules,
  ruleCorrectionOf,
  type ExecuteItemResult,
} from '@/background/execute'
import { decodeCompound, encodeCompound, listRules, newRule } from '@/shared/rules'
import { setRules, setSettings } from '@/shared/storage'
import type { PlanItem, Rule } from '@/shared/types'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const SUBJECT = '第三季財報審閱意見書'

function domainRule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: FID,
      targetFolderPath: '03/A',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    id: 'broad-rule',
    createdAt: '2026-01-01T00:00:00.000Z',
    matchCount: 12,
    ...over,
  }
}

/** A row the domain rule routed to 03/A, re-targeted by the user (PlanRow.setTarget shape). */
function correctedRow(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: `RE: ${SUBJECT}`,
    emailFrom: 'cfo@client.example',
    action: 'move',
    targetFolderId: FID,
    targetFolderPath: '03/B',
    confidence: 0.7,
    reason: '寄件人網域 @client.example',
    source: 'ai', // flipped from 'rule' by the edit
    ruleId: undefined,
    originalRuleId: 'broad-rule',
    userTouched: true,
    ...over,
  }
}

const moved = (emailId: string): ExecuteItemResult => ({
  emailId,
  subject: 'X',
  action: 'move',
  status: 'moved',
})

describe('ruleCorrectionOf', () => {
  const r = domainRule()
  const byId = new Map([[r.id, r]])
  it('re-targeted rule row → the originating rule', () => {
    expect(ruleCorrectionOf(correctedRow(), byId)?.id).toBe('broad-rule')
  })
  it('not a correction: untouched (auto-propagated) sibling, same target, rule gone, still rule-sourced', () => {
    expect(ruleCorrectionOf(correctedRow({ userTouched: undefined }), byId)).toBeNull()
    expect(ruleCorrectionOf(correctedRow({ targetFolderPath: '03/A' }), byId)).toBeNull()
    expect(ruleCorrectionOf(correctedRow({ originalRuleId: 'deleted' }), byId)).toBeNull()
    expect(ruleCorrectionOf(correctedRow({ source: 'rule' }), byId)).toBeNull()
    expect(ruleCorrectionOf(correctedRow({ action: 'delete', targetFolderPath: undefined }), byId)).toBeNull()
  })
})

describe('the correction is learned through the override path', () => {
  it('repro: client.example→A corrected to B → compound learned, broad rule superseded', async () => {
    await setRules([domainRule()])
    const item = correctedRow()
    const conf = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(conf.added).toBe(0) // not mis-attributed as a passive AI confirmation
    const ov = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(ov.added).toBe(1) // was 0
    expect(ov.disabled).toBe(1) // was 0
    const rules = await listRules()
    const compound = rules.find((r) => r.type === 'compound')!
    expect(compound.targetFolderPath).toBe('03/B')
    expect(compound.source).toBe('ai_overridden')
    expect(decodeCompound(compound.signal)?.conditions).toEqual([
      { type: 'domain', value: 'client.example' },
      { type: 'subject_keyword', value: SUBJECT },
    ])
    const broad = rules.find((r) => r.id === 'broad-rule')!
    expect(broad.enabled).toBe(false)
    expect(broad.autoDisabledReason).toBe('superseded')
  })

  it('bypasses the birth gate: learned on first sight, no subject memory needed', async () => {
    // Internal mail (P6): the originating rule is a sender+subject compound.
    await setSettings({ internalDomains: ['firm.example'] })
    const wrong: Rule = {
      ...domainRule(),
      id: 'internal-rule',
      type: 'compound',
      signal: encodeCompound([
        { type: 'sender', value: 'colleague@firm.example' },
        { type: 'subject_keyword', value: SUBJECT },
      ]),
    }
    await setRules([wrong])
    const item = correctedRow({
      emailFrom: 'colleague@firm.example',
      originalRuleId: 'internal-rule',
    })
    const ov = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(ov.added).toBe(1)
    expect(ov.disabled).toBe(1) // same signal, wrong target → superseded
    const live = (await listRules()).filter((r) => r.enabled)
    expect(live.map((r) => r.targetFolderPath)).toEqual(['03/B'])
  })

  it('short subject: silence the broad rule, learn nothing new', async () => {
    await setRules([domainRule()])
    const item = correctedRow({ emailSubject: '請款通知' })
    const ov = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(ov).toMatchObject({ added: 0, disabled: 1 })
    const [broad] = await listRules()
    expect(broad!.enabled).toBe(false)
    expect(broad!.autoDisabledReason).toBe('override')
  })

  it('a user_manual originating rule is never silenced (but the correction is still learned)', async () => {
    await setRules([domainRule({ source: 'user_manual' })])
    const item = correctedRow()
    const ov = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(ov).toMatchObject({ added: 1, disabled: 0 })
    expect((await listRules()).find((r) => r.id === 'broad-rule')!.enabled).toBe(true)
  })

  it('an auto-propagated sibling (no userTouched) is not treated as a correction', async () => {
    await setRules([domainRule()])
    const item = correctedRow({ userTouched: undefined })
    const ov = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(ov).toMatchObject({ added: 0, disabled: 0 })
  })
})
