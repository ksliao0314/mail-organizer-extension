import { describe, expect, it } from 'vitest'
import {
  addRules,
  deleteRule,
  newRule,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import {
  clearRuleHistory,
  getRuleEvents,
  recordRuleEvents,
} from '@/shared/storage'
import type { Rule, RuleEvent } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    source: 'user_manual',
    enabled: true,
    ...over,
  })
}

describe('rule history — basic recording', () => {
  it('records a create event when a new rule is upserted', async () => {
    const r = rule()
    await upsertRule(r, { actor: 'user' })
    const events = await getRuleEvents()
    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({
      kind: 'create',
      ruleId: r.id,
      actor: 'user',
    })
  })

  it('records an edit event with changedFields when an existing rule is updated', async () => {
    const r = rule({ targetFolderPath: '03/A', confidence: 0.7 })
    await upsertRule(r, { actor: 'user' })
    const updated = { ...r, targetFolderPath: '03/B', confidence: 0.9 }
    await upsertRule(updated, { actor: 'user' })
    const events = await getRuleEvents()
    expect(events).toHaveLength(2)
    expect(events[1]).toMatchObject({ kind: 'edit', ruleId: r.id, actor: 'user' })
    if (events[1]?.kind !== 'edit') throw new Error('expected edit event')
    expect(events[1].changedFields).toContain('targetFolderPath')
    expect(events[1].changedFields).toContain('confidence')
    expect(events[1].before.targetFolderPath).toBe('03/A')
    expect(events[1].after.targetFolderPath).toBe('03/B')
  })

  it('records a delete event with the pre-delete snapshot', async () => {
    const r = rule()
    await upsertRule(r, { actor: 'user' })
    await clearRuleHistory()
    await deleteRule(r.id, { actor: 'user' })
    const events = await getRuleEvents()
    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({ kind: 'delete', ruleId: r.id })
    if (events[0]?.kind !== 'delete') throw new Error('expected delete event')
    expect(events[0].before.signal).toBe('example.com')
  })

  it('records a toggle event only when state actually changes', async () => {
    const r = rule({ enabled: true })
    await upsertRule(r, { actor: 'user' })
    await clearRuleHistory()
    // No-op toggle (already enabled) → no event
    await toggleRule(r.id, true, { actor: 'user' })
    expect(await getRuleEvents()).toHaveLength(0)
    // Real toggle off → event
    await toggleRule(r.id, false, { actor: 'user' })
    const events = await getRuleEvents()
    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({
      kind: 'toggle',
      ruleId: r.id,
      enabled: false,
      actor: 'user',
    })
  })

  it('records bulk create events for addRules with caller-supplied actor', async () => {
    const rs = [rule({ signal: 'a.com' }), rule({ signal: 'b.com' })]
    await addRules(rs, { actor: 'system' })
    const events = await getRuleEvents()
    expect(events).toHaveLength(2)
    for (const e of events) {
      expect(e.kind).toBe('create')
      expect(e.actor).toBe('system')
    }
  })
})

describe('rule history — actor semantics', () => {
  it('defaults to actor=system when not specified', async () => {
    await upsertRule(rule())
    const events = await getRuleEvents()
    expect(events[0]?.actor).toBe('system')
  })

  it('allows mixing user and system events in a session', async () => {
    await upsertRule(rule({ signal: 'a.com' }), { actor: 'user' })
    await addRules([rule({ signal: 'b.com' })], { actor: 'system' })
    const events = await getRuleEvents()
    const actors = events.map((e) => e.actor).sort()
    expect(actors).toEqual(['system', 'user'])
  })
})

describe('rule history — storage cap + persistence', () => {
  it('keeps only the last 500 events (FIFO trim)', async () => {
    // Build 502 events via recordRuleEvents directly
    const events: RuleEvent[] = []
    for (let i = 0; i < 502; i++) {
      events.push({
        kind: 'toggle',
        ruleId: `r${i}`,
        at: Date.now() + i,
        actor: 'system',
        enabled: true,
        signal: `s${i}`,
        type: 'domain',
        targetFolderPath: '/x',
      })
    }
    await recordRuleEvents(events)
    const stored = await getRuleEvents()
    expect(stored).toHaveLength(500)
    // Oldest two should be trimmed
    expect((stored[0] as { ruleId: string }).ruleId).toBe('r2')
  })

  it('getRuleEvents(limit) returns only the most recent N', async () => {
    const events: RuleEvent[] = []
    for (let i = 0; i < 10; i++) {
      events.push({
        kind: 'toggle',
        ruleId: `r${i}`,
        at: i,
        actor: 'user',
        enabled: false,
        signal: `s${i}`,
        type: 'domain',
        targetFolderPath: '/x',
      })
    }
    await recordRuleEvents(events)
    const last3 = await getRuleEvents(3)
    expect(last3).toHaveLength(3)
    expect(last3.map((e) => (e as { ruleId: string }).ruleId)).toEqual([
      'r7',
      'r8',
      'r9',
    ])
  })

  it('clearRuleHistory wipes everything', async () => {
    await upsertRule(rule(), { actor: 'user' })
    expect(await getRuleEvents()).toHaveLength(1)
    await clearRuleHistory()
    expect(await getRuleEvents()).toHaveLength(0)
  })
})

describe('rule history — system-initiated paths leave audit trail', () => {
  it('addRules without actor records system events (auto-derived rules)', async () => {
    await addRules([rule({ source: 'ai_confirmed' }), rule({ source: 'auto_scan' })])
    const events = await getRuleEvents()
    expect(events).toHaveLength(2)
    expect(events.every((e) => e.actor === 'system')).toBe(true)
  })
})
