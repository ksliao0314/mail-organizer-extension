import { describe, expect, it } from 'vitest'
import {
  aiOriginalTargetPath,
  finalTargetPath,
  generateAiOverrideRules,
  wasUserOverride,
  type ExecuteItemResult,
} from '@/background/execute'
import { listRules } from '@/shared/rules'
import { getRuleEvents, setRules } from '@/shared/storage'
import type { PlanItem, Rule } from '@/shared/types'

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: 'X',
    emailFrom: 'alice@example.com',
    action: 'move',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    reason: '',
    source: 'ai',
    ...over,
  }
}

function rule(over: Partial<Rule>): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Wrong',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'ai_confirmed',
    ...over,
  }
}

function result(emailId: string, status: ExecuteItemResult['status']): ExecuteItemResult {
  return {
    emailId,
    subject: 'X',
    action: 'move',
    status,
  }
}

describe('wasUserOverride', () => {
  it('false when there was no AI suggestion (unresolved items)', () => {
    const item = planItem({ source: 'unresolved', aiOriginalAction: undefined })
    expect(wasUserOverride(item)).toBe(false)
  })

  it('false when final action+target match AI original', () => {
    const item = planItem({
      action: 'move',
      targetFolderPath: '03/Y',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    expect(wasUserOverride(item)).toBe(false)
  })

  it('true when action differs from AI', () => {
    const item = planItem({
      action: 'delete',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    expect(wasUserOverride(item)).toBe(true)
  })

  it('true when target differs from AI', () => {
    const item = planItem({
      action: 'move',
      targetFolderPath: '03/Z',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    expect(wasUserOverride(item)).toBe(true)
  })

  it('handles new_folder → move conversion', () => {
    const item = planItem({
      action: 'move',
      targetFolderPath: '03/Existing',
      aiOriginalAction: 'new_folder',
      aiOriginalSuggestedFolderName: 'NewFolder',
      aiOriginalSuggestedParentPath: '03',
    })
    expect(wasUserOverride(item)).toBe(true)
  })
})

describe('finalTargetPath / aiOriginalTargetPath', () => {
  it('returns move target', () => {
    const item = planItem({ action: 'move', targetFolderPath: 'a/b' })
    expect(finalTargetPath(item)).toBe('a/b')
  })

  it('returns new_folder joined path', () => {
    const item = planItem({
      action: 'new_folder',
      suggestedParentPath: 'a',
      suggestedFolderName: 'b',
    })
    expect(finalTargetPath(item)).toBe('a/b')
  })

  it('encodes slashes in new_folder name', () => {
    const item = planItem({
      action: 'new_folder',
      suggestedParentPath: 'a',
      suggestedFolderName: 'b/c',
    })
    expect(finalTargetPath(item)).toBe('a/b／c')
  })

  it('returns undefined for delete / skip', () => {
    expect(finalTargetPath(planItem({ action: 'delete' }))).toBeUndefined()
    expect(finalTargetPath(planItem({ action: 'skip' }))).toBeUndefined()
  })

  it('reads AI original target from snapshot fields', () => {
    const item = planItem({
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: 'x/y',
    })
    expect(aiOriginalTargetPath(item)).toBe('x/y')
  })
})

describe('generateAiOverrideRules', () => {
  it('does nothing when no items were overridden', async () => {
    const item = planItem({
      action: 'move',
      targetFolderPath: '03/Y',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(0)
    expect(r.disabled).toBe(0)
  })

  it('creates an ai_overridden rule when user picks a different target', async () => {
    // Subject carries a latin-style case code so chooseLearningSignal returns
    // case_code (not domain). Plain-domain rule creation is intentionally
    // suppressed in override learning — same domain often serves multiple
    // cases at this firm — so we exercise the case_code branch here.
    const item = planItem({
      emailSubject: 'Re: 案件 25A0067A 通知',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.type).toBe('case_code')
    expect(rules[0]!.signal).toBe('25A0067A')
    expect(rules[0]!.targetFolderPath).toBe('03/Right')
    expect(rules[0]!.source).toBe('ai_overridden')
    expect(rules[0]!.confidence).toBe(0.95)
  })

  it('disables a competing same-signal rule that points at the AI’s wrong target', async () => {
    await setRules([
      rule({ type: 'case_code', source: 'ai_confirmed', signal: '25A0067A', targetFolderPath: '03/Wrong' }),
    ])
    const item = planItem({
      emailSubject: 'Re: 案件 25A0067A 通知',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(1)
    expect(r.disabled).toBe(1)
    const rules = await listRules()
    expect(rules).toHaveLength(2)
    const wrong = rules.find((x) => x.targetFolderPath === '03/Wrong')!
    expect(wrong.enabled).toBe(false)
    const right = rules.find((x) => x.targetFolderPath === '03/Right')!
    expect(right.enabled).toBe(true)
    expect(right.source).toBe('ai_overridden')
  })

  it('preserves user_manual rules even when they point to a different target', async () => {
    await setRules([
      rule({ type: 'case_code', source: 'user_manual', signal: '25A0067A', targetFolderPath: '03/Manual' }),
    ])
    const item = planItem({
      emailSubject: 'Re: 案件 25A0067A 通知',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Manual',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(1)
    expect(r.disabled).toBe(0) // user_manual was NOT disabled
    const rules = await listRules()
    const manual = rules.find((x) => x.source === 'user_manual')!
    expect(manual.enabled).toBe(true)
  })

  it('skips items where final action is delete or skip (cannot codify as rule)', async () => {
    const deleteItem = planItem({
      emailFrom: 'alice@example.com',
      action: 'delete',
      targetFolderPath: undefined,
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    const skipItem = planItem({
      emailFrom: 'alice@example.com',
      action: 'skip',
      targetFolderPath: undefined,
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Y',
    })
    const r = await generateAiOverrideRules(
      [deleteItem, skipItem],
      [result(deleteItem.emailId, 'deleted'), result(skipItem.emailId, 'skipped')],
    )
    expect(r.added).toBe(0)
  })

  it('skips when the email move was not successful', async () => {
    const item = planItem({
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'error')])
    expect(r.added).toBe(0)
  })

  it('skips internal domain', async () => {
    const item = planItem({
      emailFrom: 'colleague@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(0)
  })

  it('skips creating a duplicate when the same (signal, target) rule already exists', async () => {
    await setRules([
      rule({ source: 'ai_confirmed', signal: 'example.com', targetFolderPath: '03/Right' }),
    ])
    const item = planItem({
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(0)
    // No same-signal-different-target rule existed to disable, either.
    expect(r.disabled).toBe(0)
  })

  it('handles unresolved items by not generating an override (no AI verdict to override)', async () => {
    const item = planItem({
      source: 'ai', // user set action=move which auto-promotes source to ai
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: undefined,
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.added).toBe(0)
  })
})

describe('generateAiOverrideRules — audit trail', () => {
  it('records a create event with actor=system for the new ai_overridden rule', async () => {
    const item = planItem({
      emailSubject: 'Re: 案件 25A0067A 通知',
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    const events = await getRuleEvents()
    const creates = events.filter((e) => e.kind === 'create')
    expect(creates).toHaveLength(1)
    expect(creates[0]).toMatchObject({ kind: 'create', actor: 'system' })
  })

  it('records toggle events when stale same-signal rules get auto-disabled', async () => {
    await setRules([
      rule({ source: 'ai_confirmed', signal: 'example.com', targetFolderPath: '03/Wrong' }),
    ])
    const item = planItem({
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    const events = await getRuleEvents()
    const toggles = events.filter((e) => e.kind === 'toggle')
    expect(toggles).toHaveLength(1)
    expect(toggles[0]).toMatchObject({
      kind: 'toggle',
      actor: 'system',
      enabled: false,
    })
  })

  it('sender + demoteOnly: demotes stale sender rule but skips creating new one when no subject feature', async () => {
    // Generic-provider domain (gmail.com) routes through P7 (sender) in
    // chooseLearningSignal. With a single-char subject ('X'), neither
    // strict nor lenient feature extractor can produce a discriminator,
    // so the signal becomes { type: 'sender', demoteOnly: true }.
    // Expected: existing sender→Wrong rule is demoted, but NO new sender→
    // Right rule is created (which would just re-create the same
    // conflict pattern next batch).
    await setRules([
      rule({
        type: 'sender',
        source: 'ai_confirmed',
        signal: 'alice@gmail.com',
        targetFolderPath: '03/Wrong',
      }),
    ])
    const item = planItem({
      emailSubject: 'X', // single char → no feature even with lenient extractor
      emailFrom: 'alice@gmail.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    const r = await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    expect(r.disabled).toBe(1) // stale rule demoted
    expect(r.added).toBe(0) // no new rule (demoteOnly skip)
    const after = await listRules()
    const enabledSenders = after.filter((rl) => rl.type === 'sender' && rl.enabled)
    expect(enabledSenders).toHaveLength(0)
  })

  it('no audit events when nothing changes (override blocked by tombstone)', async () => {
    // Existing rule that matches exactly what override would produce
    await setRules([
      rule({ source: 'ai_confirmed', signal: 'example.com', targetFolderPath: '03/Right' }),
    ])
    const item = planItem({
      emailFrom: 'alice@example.com',
      action: 'move',
      targetFolderPath: '03/Right',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    })
    // The override is allowed but already-covered → no new rule, no demote
    // (because no competing different-target rule exists). Events should be
    // empty (excluding the setRules-driven seed event which goes through a
    // path that doesn't audit).
    const before = (await getRuleEvents()).length
    await generateAiOverrideRules([item], [result(item.emailId, 'moved')])
    const after = (await getRuleEvents()).length
    expect(after).toBe(before)
  })
})
