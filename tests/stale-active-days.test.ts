// Stale-sweep ACTIVE-day clock (2026-11-05).
//
// Regression for the audit P0: autoDisableStaleRules used wall-clock age, and
// the sweep runs on every SW wake — so a user who simply stopped classifying
// for ~100 days came back to a library with every learned rule hard-deleted
// (no tombstone / audit / notice), and a restored backup was deleted again on
// the next wake because it keeps its old lastUsedAt. Age is now counted in
// days on which a batch was actually executed.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  autoDisableStaleRules,
  countActiveDaysAfter,
  deriveActiveDaysFromRules,
  encodeCompound,
  listRules,
  newRule,
} from '@/shared/rules'
import { getActiveDays, getLastSweepAt, recordActiveDay, setRules, utcDay } from '@/shared/storage'
import type { Rule } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'

const DAY = 86_400_000
const LAST_BATCH = Date.parse('2026-07-16T03:00:00.000Z')

// newRule() only takes identity fields; apply matchCount / dates / source etc.
// on top so the overrides actually land. createdAt is pinned to a fixed past
// date: newRule stamps the REAL clock, and an AI-learned rule's birth day is
// itself an active day — on a later real date it would land "after" every
// fixture date and skew the counts (a time bomb caught by a future-clock run).
function learned(over: Partial<Rule> = {}): Rule {
  const base = newRule({
    type: 'domain',
    signal: 'client.example',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03進行中案件/某案',
    confidence: 0.7,
    source: 'ai_confirmed',
  })
  return { ...base, createdAt: '2026-01-01T00:00:00.000Z', ...over }
}

const fullSubject = (over: Partial<Rule> = {}) =>
  learned({
    type: 'compound',
    signal: encodeCompound([
      { type: 'sender', value: 'colleague@firm.example' },
      { type: 'subject_keyword', value: '請審閱三井租賃合約修訂版' },
    ]),
    ...over,
  })

/** A library that was in daily use until LAST_BATCH. */
function libraryUsedUntilJuly(): Rule[] {
  const used = learned({ signal: 'used.example', matchCount: 9 })
  used.createdAt = new Date(LAST_BATCH - 60 * DAY).toISOString()
  used.lastUsedAt = new Date(LAST_BATCH - 5 * DAY).toISOString()
  const caseRule = learned({ type: 'case_code', signal: '25A0067A', matchCount: 3 })
  caseRule.createdAt = new Date(LAST_BATCH - 90 * DAY).toISOString()
  caseRule.lastUsedAt = new Date(LAST_BATCH - 20 * DAY).toISOString()
  const neverMatched = fullSubject()
  neverMatched.createdAt = new Date(LAST_BATCH - 10 * DAY).toISOString()
  return [used, caseRule, neverMatched]
}

describe('idle extension does not age the library', () => {
  it('a year with no executed batch deletes nothing (was: every learned rule)', async () => {
    await setRules(libraryUsedUntilJuly())
    const report = await autoDisableStaleRules({ now: LAST_BATCH + 365 * DAY })
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(3)
  })

  it('a restored backup with old lastUsedAt survives the next sweep', async () => {
    // Same library re-imported in November after months of disuse: nothing
    // has been recorded since July, so the clock still reads July.
    await setRules(libraryUsedUntilJuly())
    await autoDisableStaleRules({ now: Date.parse('2026-11-05T00:00:00.000Z') })
    await autoDisableStaleRules({ now: Date.parse('2026-11-06T00:00:00.000Z') })
    expect(await listRules()).toHaveLength(3)
  })

  it('after resuming, rules age only by the days actually used', async () => {
    await setRules(libraryUsedUntilJuly())
    const resumed = Date.parse('2026-11-20T03:00:00.000Z')
    // 10 days of classification in November.
    const report = await autoDisableStaleRules({
      now: resumed,
      activeDays: dailyActiveDays(resumed, 10),
    })
    // never-matched full_subject: derived July days after its birth + 10 → < 45.
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(3)
  })
})

describe('active use still prunes dead rules', () => {
  it('never-matched full_subject is deleted after 45 days of use', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const r = fullSubject()
    r.createdAt = new Date(now - 50 * DAY).toISOString()
    await setRules([r])
    const report = await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 60) })
    expect(report.byReason.stale).toBe(1)
    expect(await listRules()).toHaveLength(0)
  })

  it('a matched rule unused for 100 days of use is deleted; 99 is kept', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const dead = learned({ signal: 'dead.example', matchCount: 4 })
    dead.lastUsedAt = new Date(now - 100 * DAY).toISOString()
    const alive = learned({ signal: 'alive.example', matchCount: 4 })
    alive.lastUsedAt = new Date(now - 99 * DAY).toISOString()
    await setRules([dead, alive])
    const report = await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 150) })
    expect(report.byReason.stale).toBe(1)
    expect((await listRules()).map((r) => r.signal)).toEqual(['alive.example'])
  })

  it('days evidenced by OTHER rules count (synced activity from another machine)', async () => {
    // No recorded ledger on this machine; the library itself shows 120
    // distinct days of hits.
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const rules: Rule[] = []
    for (let i = 0; i < 120; i++) {
      const r = learned({ signal: `d${i}.example`, matchCount: 1 })
      r.lastUsedAt = new Date(now - i * DAY).toISOString()
      rules.push(r)
    }
    await setRules(rules)
    const report = await autoDisableStaleRules({ now, activeDays: [] })
    // Rules last used on days 100..119 ago have ≥ 100 evidenced days after them.
    expect(report.byReason.stale).toBe(20)
    expect(await listRules()).toHaveLength(100)
  })

  it('unparseable createdAt is never treated as stale', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const r = fullSubject()
    r.createdAt = ''
    await setRules([r])
    const report = await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 200) })
    expect(report.byReason.stale).toBe(0)
  })

  it('user_manual is never stale-deleted', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const r = learned({ source: 'user_manual', matchCount: 2 })
    r.lastUsedAt = new Date(now - 300 * DAY).toISOString()
    await setRules([r])
    await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 350) })
    expect(await listRules()).toHaveLength(1)
  })
})

describe('active-day helpers', () => {
  it('countActiveDaysAfter counts strictly-after days and stops at `enough`', () => {
    const desc = ['2026-11-05', '2026-11-04', '2026-11-01', '2026-10-30']
    expect(countActiveDaysAfter(desc, '2026-11-01T23:00:00Z', 10)).toBe(2)
    expect(countActiveDaysAfter(desc, '2026-10-01T00:00:00Z', 3)).toBe(3)
    expect(countActiveDaysAfter(desc, 'garbage', 10)).toBeNull()
    expect(countActiveDaysAfter(desc, undefined, 10)).toBeNull()
  })

  it('deriveActiveDaysFromRules uses lastUsedAt of all rules + birth of AI-learned ones', () => {
    const a = learned({ source: 'auto_scan' })
    a.createdAt = '2026-05-01T00:00:00Z' // initial-scan birth: not an execute day
    const b = learned({ source: 'ai_overridden' })
    b.createdAt = '2026-06-02T08:00:00Z'
    b.lastUsedAt = '2026-06-10T08:00:00Z'
    expect(deriveActiveDaysFromRules([a, b]).sort()).toEqual(['2026-06-02', '2026-06-10'])
  })

  it('recordActiveDay is idempotent per UTC day and capped', async () => {
    const t = Date.parse('2026-11-05T01:00:00Z')
    await recordActiveDay(t)
    await recordActiveDay(t + 3 * 3_600_000)
    expect(await getActiveDays()).toEqual(['2026-11-05'])
    for (let i = 1; i <= 405; i++) await recordActiveDay(t + i * DAY)
    const days = await getActiveDays()
    expect(days).toHaveLength(400)
    expect(days[days.length - 1]).toBe(utcDay(t + 405 * DAY))
  })
})

// ---- sweep cadence ---------------------------------------------------------

type AlarmsMock = {
  get: ReturnType<typeof vi.fn>
  create: ReturnType<typeof vi.fn>
  onAlarm: { addListener: ReturnType<typeof vi.fn> }
}

describe('stale sweep cadence', () => {
  let alarms: AlarmsMock
  beforeEach(() => {
    alarms = {
      get: vi.fn(async () => undefined),
      create: vi.fn(async () => {}),
      onAlarm: { addListener: vi.fn() },
    }
    ;(globalThis as unknown as { chrome: { alarms?: AlarmsMock } }).chrome.alarms = alarms
  })
  afterEach(() => {
    delete (globalThis as unknown as { chrome: { alarms?: AlarmsMock } }).chrome.alarms
    vi.resetModules()
  })

  it('creates the daily alarm only when missing (no 5-minute re-arm per wake)', async () => {
    const { installStaleSweepListener } = await import('@/background/stale-sweep')
    alarms.get.mockResolvedValueOnce({ name: 'stale-rule-sweep', scheduledTime: Date.now() + 1 })
    installStaleSweepListener()
    await vi.waitFor(() => expect(alarms.get).toHaveBeenCalled())
    await new Promise((r) => setTimeout(r, 0))
    expect(alarms.create).not.toHaveBeenCalled()

    installStaleSweepListener() // alarm now missing
    await vi.waitFor(() => expect(alarms.create).toHaveBeenCalledTimes(1))
    expect(alarms.create.mock.calls[0]![1]).toMatchObject({ periodInMinutes: 1440 })
  })

  it('a non-forced sweep right after a sweep is skipped; force bypasses the gate', async () => {
    const { runSweep } = await import('@/background/stale-sweep')
    const first = await runSweep()
    expect(first.skipped).toBeUndefined()
    expect(await getLastSweepAt()).not.toBeNull()
    expect((await runSweep()).skipped).toBe(true)
    expect((await runSweep({ force: true })).skipped).toBeUndefined()
  })

  it('overlapping calls share one sweep', async () => {
    const { runSweep } = await import('@/background/stale-sweep')
    const a = runSweep()
    const b = runSweep()
    expect(b).toBe(a)
    await a
  })
})
