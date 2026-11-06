// Auto-disabled rule lifecycle (2026-11-05).
//
// Item 1 — revive: every learner used to treat ANY disabled copy of a triple
// as "already covered", so addRulesFilteringTombstones' revive branch was
// unreachable: a rule the engine one-strike-disabled ('override') or
// superseded stayed dead forever while blocking its own triple. AUTO-disabled
// rules ('override' | 'superseded' | 'high-error-rate') now revive when
// learning re-derives the same (type, normalized signal, target); user-disabled
// ones (no reason) keep blocking.
//
// Item 13 — bookkeeping + purge: silencing a broad rule records
// autoDisabledReason 'superseded' (+autoDisabledAt); the sweep hard-deletes
// AUTO-disabled rules after 60 ACTIVE days (no tombstone).

import { describe, expect, it } from 'vitest'
import {
  generateAiConfirmedRules,
  generateAiOverrideRules,
  type ExecuteItemResult,
} from '@/background/execute'
import {
  addRulesFilteringTombstones,
  autoDisableStaleRules,
  isRevivableAutoDisabled,
  listRules,
  newRule,
} from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { PlanItem, Rule } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const DAY = 86_400_000

// createdAt pinned: newRule stamps the real clock (time-bomb guard).
function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'vendor.example',
      targetFolderId: FID,
      targetFolderPath: '03/A',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  }
}

const autoOff = (reason: Rule['autoDisabledReason'], over: Partial<Rule> = {}): Rule =>
  rule({ enabled: false, autoDisabledReason: reason, autoDisabledAt: '2026-10-01T00:00:00.000Z', ...over })

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: '一般往來信件內容',
    emailFrom: 'billing@vendor.example',
    action: 'move',
    targetFolderPath: '03/A',
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}

const moved = (emailId: string): ExecuteItemResult => ({
  emailId,
  subject: 'X',
  action: 'move',
  status: 'moved',
})

describe('isRevivableAutoDisabled', () => {
  it('only engine reasons are revivable; user-disabled and location/legacy parks are not', () => {
    expect(isRevivableAutoDisabled(autoOff('override'))).toBe(true)
    expect(isRevivableAutoDisabled(autoOff('superseded'))).toBe(true)
    expect(isRevivableAutoDisabled(autoOff('high-error-rate'))).toBe(true)
    expect(isRevivableAutoDisabled(autoOff('case_closed'))).toBe(false)
    expect(isRevivableAutoDisabled(autoOff('legacy_token'))).toBe(false)
    expect(isRevivableAutoDisabled(rule({ enabled: false }))).toBe(false) // user-disabled
    expect(isRevivableAutoDisabled(rule())).toBe(false) // enabled
  })
})

describe('learning revives AUTO-disabled rules (item 1)', () => {
  it('AI re-confirming a one-strike (override) case rule switches it back on', async () => {
    const dormant = autoOff('override', { type: 'case_code', signal: '25A0067A' })
    await setRules([dormant])
    const item = planItem({ emailSubject: 'Re: 25A0067A 補件通知' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0) // no twin row
    expect(r.revived).toBe(1)
    expect(r.addedIds).toEqual([]) // undo must never delete a pre-existing rule
    const [after] = await listRules()
    expect(after!.id).toBe(dormant.id)
    expect(after!.enabled).toBe(true)
    expect(after!.autoDisabledReason).toBeUndefined()
    expect(after!.autoDisabledAt).toBeUndefined()
  })

  // 2026-11-06 (B1) — this used to assert the opposite. A superseded broad
  // rule is its anchor's conflict history (先廣後窄): reviving it on the next
  // A confirmation routed the following B-case mail to A without the AI. The
  // anchor stays narrow instead — tests/superseded-anchor.test.ts.
  it('a superseded plain-domain rule is NOT revived when the AI re-confirms its target', async () => {
    const dormant = autoOff('superseded')
    await setRules([dormant])
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.revived).toBe(0)
    expect(r.added).toBe(0) // the narrow form waits for its second observation
    expect((await listRules()).map((x) => x.enabled)).toEqual([false])
  })

  it('a plain-domain rule parked for another reason (high-error-rate) still revives', async () => {
    await setRules([autoOff('high-error-rate', { matchCount: 3, overrideCount: 20 })])
    const item = planItem()
    expect((await generateAiConfirmedRules([item], [moved(item.emailId)])).revived).toBe(1)
    expect((await listRules()).map((x) => x.enabled)).toEqual([true])
  })

  it('a revived high-error-rate rule gets a fresh accuracy window (no revive/disable flap)', async () => {
    const dormant = autoOff('high-error-rate', { matchCount: 10, overrideCount: 14 })
    await setRules([dormant])
    const item = planItem()
    expect((await generateAiConfirmedRules([item], [moved(item.emailId)])).revived).toBe(1)
    const [revived] = await listRules()
    expect(revived).toMatchObject({ enabled: true, matchCount: 1, overrideCount: 0 })
    // The next sweep must not switch it straight back off.
    const report = await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    expect(report.byReason.highErrorRate).toBe(0)
    expect((await listRules())[0]!.enabled).toBe(true)
  })

  it('user-disabled rules (no reason) keep blocking their triple', async () => {
    await setRules([rule({ enabled: false })])
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
    expect(r.revived).toBe(0)
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.enabled).toBe(false)
  })

  it('case_closed and legacy_token parks are not revived by learning', async () => {
    await setRules([
      autoOff('case_closed', { type: 'case_code', signal: '25A0067A' }),
      autoOff('legacy_token', { type: 'subject_keyword', signal: '甲公司' }),
    ])
    const a = planItem({ emailSubject: '25A0067A 補件' })
    const r = await generateAiConfirmedRules([a], [moved(a.emailId)])
    expect(r.revived).toBe(0)
    expect(r.added).toBe(0)
    expect((await listRules()).every((x) => !x.enabled)).toBe(true)
  })

  it('never revives into a conflict with a live same-signal rule (structural ids included)', async () => {
    // User re-targeted case X from A to B: X→A one-strike-disabled, X→B live.
    // Old mail of case X still sitting in folder A must not bring X→A back.
    await setRules([
      autoOff('override', { type: 'case_code', signal: '25A0067A', targetFolderPath: '03/A' }),
      rule({ type: 'case_code', signal: '25A0067A', targetFolderPath: '03/B', source: 'ai_overridden' }),
    ])
    const item = planItem({ emailSubject: '25A0067A 舊信' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.revived).toBe(0)
    expect(r.added).toBe(0)
    const a = (await listRules()).find((x) => x.targetFolderPath === '03/A')!
    expect(a.enabled).toBe(false)
  })

  it('addRulesFilteringTombstones applies the same conflict guard to any caller', async () => {
    const dormant = autoOff('high-error-rate')
    await setRules([dormant, rule({ targetFolderPath: '03/B' })])
    const res = await addRulesFilteringTombstones([rule()])
    expect(res.reEnabled).toHaveLength(0)
    expect(res.added).toHaveLength(0)
  })

  it('the override path revives its own auto-disabled twin and supersedes the wrong one', async () => {
    // X→B was one-strike-disabled earlier; the user now corrects X mail from
    // A back to B. Expect: X→A superseded, X→B revived — no duplicate X→B.
    const wrong = rule({ type: 'case_code', signal: '25A0067A', targetFolderPath: '03/A' })
    const dormant = autoOff('override', {
      type: 'case_code',
      signal: '25A0067A',
      targetFolderPath: '03/B',
      source: 'ai_overridden',
    })
    await setRules([wrong, dormant])
    const item = planItem({
      emailSubject: '25A0067A 開庭通知',
      targetFolderPath: '03/B',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/A',
    })
    const r = await generateAiOverrideRules([item], [moved(item.emailId)])
    expect(r).toMatchObject({ added: 0, disabled: 1, revived: 1 })
    const rules = await listRules()
    expect(rules).toHaveLength(2)
    expect(rules.find((x) => x.id === dormant.id)!.enabled).toBe(true)
    const w = rules.find((x) => x.id === wrong.id)!
    expect(w.enabled).toBe(false)
    expect(w.autoDisabledReason).toBe('superseded')
  })
})

describe('superseded bookkeeping + dormant purge (item 13)', () => {
  const NOW = Date.parse('2026-11-20T03:00:00.000Z')
  const at = (daysAgo: number) => new Date(NOW - daysAgo * DAY).toISOString()

  it('purges superseded / override / high-error-rate rules after 60 ACTIVE days; keeps 59', async () => {
    const old1 = autoOff('superseded', { signal: 'a.example', autoDisabledAt: at(60) })
    const old2 = autoOff('override', { signal: 'b.example', autoDisabledAt: at(65) })
    const old3 = autoOff('high-error-rate', { signal: 'c.example', autoDisabledAt: at(90) })
    const young = autoOff('superseded', { signal: 'd.example', autoDisabledAt: at(59) })
    await setRules([old1, old2, old3, young])
    const report = await autoDisableStaleRules({ now: NOW, activeDays: dailyActiveDays(NOW, 120) })
    expect(report.byReason.dormantPurged).toBe(3)
    expect((await listRules()).map((r) => r.signal)).toEqual(['d.example'])
  })

  it('never purges user_manual, user-disabled, case_closed or legacy_token rules', async () => {
    const keep = [
      autoOff('superseded', { signal: 'um.example', source: 'user_manual', autoDisabledAt: at(200) }),
      rule({ signal: 'user-off.example', enabled: false }),
      autoOff('legacy_token', { type: 'subject_keyword', signal: '甲公司', autoDisabledAt: at(200) }),
    ]
    await setRules(keep)
    const report = await autoDisableStaleRules({ now: NOW, activeDays: dailyActiveDays(NOW, 300) })
    expect(report.byReason.dormantPurged).toBe(0)
    expect(await listRules()).toHaveLength(3)
  })

  it('the dormant clock is frozen while the tool sits idle (calendar days do not count)', async () => {
    await setRules([autoOff('superseded', { autoDisabledAt: at(200) })])
    const report = await autoDisableStaleRules({ now: NOW, activeDays: dailyActiveDays(NOW, 10) })
    expect(report.byReason.dormantPurged).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })
})
