// Rule-library hygiene (2026-11-05).
//
// Item 9  — dedupeRulesByKey picked a survivor without looking at lastUsedAt,
//           dropping the copy used yesterday; the stale survivor was then
//           swept. The survivor now absorbs the group's history.
// Item 15 — normalizeSignal treated forms the matcher considers EQUAL as
//           distinct keys (full vs compact court cases, full-width digits,
//           whitespace runs, padded '@domain'), so tombstones / dedupe /
//           conflict detection / alreadyCovered all missed equivalents.
// Item 16 — runSweep does dedupe + closed-case retire/purge + stale in ONE
//           rules transaction: ≤ 1 write (so ≤ 1 sync push) per sweep.
// Item 10 — a malformed compound no longer aborts the sweep.
// B9 (2026-11-06) — canonical dedupe must not let a DISABLED user_manual
//           spelling absorb the ENABLED differently-spelled twin that was
//           routing the case.

import { describe, expect, it } from 'vitest'
import { generateAiOverrideRules } from '@/background/execute'
import { runSweep } from '@/background/stale-sweep'
import {
  addRulesFilteringTombstones,
  autoDisableStaleRules,
  dedupeRulesByKey,
  encodeCompound,
  filterByTombstones,
  findConflicts,
  listRules,
  matchEmail,
  newRule,
  normalizeSignal,
  runRuleSweepPass,
  tombstoneKey,
  upsertRule,
} from '@/shared/rules'
import {
  addRuleTombstones,
  getRuleTombstones,
  setRuleCounterBaseline,
  setRules,
  setSettings,
} from '@/shared/storage'
import type { Email, PlanItem, Rule } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const DAY = 86_400_000

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: FID,
      targetFolderPath: '03/A',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  }
}

describe('dedupe merges usage history into the survivor (item 9)', () => {
  it('keeps the newest lastUsedAt, max counts and the oldest createdAt', async () => {
    const survivor = rule({
      matchCount: 10, // wins the tiebreak
      lastUsedAt: '2026-03-01T00:00:00.000Z',
      createdAt: '2026-02-01T00:00:00.000Z',
    })
    const usedYesterday = rule({
      matchCount: 3,
      overrideCount: 4,
      lastUsedAt: '2026-11-04T00:00:00.000Z',
      createdAt: '2026-01-05T00:00:00.000Z',
    })
    await setRules([survivor, usedYesterday])
    const report = await dedupeRulesByKey()
    expect(report).toEqual({ removed: 1, groupsAffected: 1 })
    const [kept] = await listRules()
    expect(kept!.id).toBe(survivor.id)
    expect(kept!.lastUsedAt).toBe('2026-11-04T00:00:00.000Z')
    expect(kept!.matchCount).toBe(10)
    expect(kept!.overrideCount).toBe(4)
    expect(kept!.createdAt).toBe('2026-01-05T00:00:00.000Z')
  })

  it('counters come from the newest accuracy window (statsResetAt) — a revive is not undone by a twin', async () => {
    const revived = rule({ matchCount: 1, overrideCount: 0, statsResetAt: Date.parse('2026-11-05T00:00:00.000Z') })
    const oldTwin = rule({ matchCount: 10, overrideCount: 15, lastUsedAt: '2026-10-01T00:00:00.000Z' })
    await setRules([oldTwin, revived])
    await dedupeRulesByKey()
    const [kept] = await listRules()
    expect(kept!.matchCount).toBe(1)
    expect(kept!.overrideCount).toBe(0)
    expect(kept!.statsResetAt).toBe(Date.parse('2026-11-05T00:00:00.000Z'))
    expect(kept!.lastUsedAt).toBe('2026-10-01T00:00:00.000Z') // usage history still merges
  })

  it('repro: the merged survivor is no longer stale-swept right after dedupe', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const staleTwin = rule({ matchCount: 9, lastUsedAt: new Date(now - 150 * DAY).toISOString() })
    const freshTwin = rule({ matchCount: 2, lastUsedAt: new Date(now - DAY).toISOString() })
    await setRules([staleTwin, freshTwin])
    await dedupeRulesByKey()
    const report = await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 200) })
    expect(report.byReason.stale).toBe(0) // was 1: the survivor kept the 150-day-old lastUsedAt
    expect(await listRules()).toHaveLength(1)
  })

  // B9 (2026-11-06): canonical keys group spellings older builds kept apart.
  it('a DISABLED user_manual full-form rule never absorbs the ENABLED compact twin that routes the case', async () => {
    const userOff = rule({
      type: 'subject_keyword',
      signal: '112年度訴字第304號',
      source: 'user_manual',
      enabled: false,
    })
    const routing = rule({ type: 'subject_keyword', signal: '112訴304', matchCount: 7 })
    const routingCopy = rule({ type: 'subject_keyword', signal: '112訴304', matchCount: 2, overrideCount: 3 })
    await setRules([userOff, routing, routingCopy])
    expect(await dedupeRulesByKey()).toEqual({ removed: 1, groupsAffected: 1 })
    const after = await listRules()
    expect(after.map((r) => r.id).sort()).toEqual([userOff.id, routing.id].sort()) // F8: the user rule stays
    const twin = after.find((r) => r.id === routing.id)!
    expect(twin).toMatchObject({ enabled: true, matchCount: 7, overrideCount: 3 }) // absorbs its own spelling group
    // …so the case keeps routing (it silently stopped before).
    const out = matchEmail(
      {
        Id: 'm1',
        Subject: '臺北地院 112年度訴字第304號 開庭通知',
        BodyPreview: '',
        From: { EmailAddress: { Address: 'clerk@court.example' } },
        ToRecipients: [],
        ReceivedDateTime: '2026-11-01T00:00:00Z',
        ParentFolderId: 'inbox',
      } as Email,
      after,
    )
    expect(out?.rule.id).toBe(routing.id)
    // Stable: the next sweep's dedupe changes nothing.
    expect(await dedupeRulesByKey()).toEqual({ removed: 0, groupsAffected: 0 })
  })

  it('F8 unchanged for an identically-spelled twin: the disabled user_manual rule wins', async () => {
    const userOff = rule({ type: 'subject_keyword', signal: '112訴304', source: 'user_manual', enabled: false })
    const auto = rule({ type: 'subject_keyword', signal: '112訴304', matchCount: 7 })
    await setRules([userOff, auto])
    await dedupeRulesByKey()
    expect((await listRules()).map((r) => r.id)).toEqual([userOff.id])
  })

  it('duplicate ids across triples: the losing copy is still removed', async () => {
    // An id-keyed keep-set kept `a` alive because `b` (other triple) shares
    // its id — leaving the 03/A duplicate in place forever.
    const a = rule({ id: 'same-id', targetFolderPath: '03/A' })
    const b = rule({ id: 'same-id', targetFolderPath: '03/B' })
    const dupOfA = rule({ targetFolderPath: '03/A', matchCount: 5 }) // wins its triple
    await setRules([a, b, dupOfA])
    await dedupeRulesByKey()
    const after = await listRules()
    expect(after.map((r) => r.targetFolderPath).sort()).toEqual(['03/A', '03/B'])
    expect(after.find((r) => r.targetFolderPath === '03/A')!.id).toBe(dupOfA.id)
  })
})

describe('normalizeSignal canonicalizes what the matcher treats as equal (item 15)', () => {
  it('court-case forms, full-width digits and whitespace runs fold together', () => {
    expect(normalizeSignal('subject_keyword', '112年度訴字第304號')).toBe('112訴304')
    expect(normalizeSignal('subject_keyword', '１１２訴３０４')).toBe('112訴304')
    expect(normalizeSignal('subject_keyword', '  Project   Phoenix ')).toBe('project phoenix')
    expect(normalizeSignal('domain', ' @Client.Example ')).toBe('client.example')
    const full = encodeCompound([
      { type: 'subject_keyword', value: '112年度訴字第304號' },
      { type: 'domain', value: '@Client.example' },
    ])
    const compact = encodeCompound([
      { type: 'domain', value: 'client.example' },
      { type: 'subject_keyword', value: '112訴304' },
    ])
    expect(normalizeSignal('compound', full)).toBe(normalizeSignal('compound', compact))
  })

  it('is idempotent — a stored signalNorm re-normalizes to itself', () => {
    for (const [type, s] of [
      ['subject_keyword', '112年度訴字第304號'],
      ['domain', ' @x.example'],
      ['compound', encodeCompound([{ type: 'domain', value: 'x.example' }, { type: 'subject_keyword', value: 'A  B' }])],
    ] as const) {
      const once = normalizeSignal(type, s)
      expect(normalizeSignal(type, once)).toBe(once)
    }
  })

  it('conflict detection sees full- and compact-form case rules as the same signal', () => {
    const a = rule({ type: 'subject_keyword', signal: '112年度訴字第304號', targetFolderPath: '03/A' })
    const b = rule({
      type: 'subject_keyword',
      signal: '112訴304',
      targetFolderPath: '03/B',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE4',
    })
    expect(findConflicts([a, b])).toHaveLength(1)
  })

  it('a tombstone stored in a legacy form still blocks the canonical re-learn', async () => {
    // Pre-2026-11-05 builds stored lowercase+trim only.
    const legacy = { type: 'subject_keyword' as const, signalNorm: '112年度訴字第304號', targetFolderPath: '03/A', deletedAt: 1 }
    expect(tombstoneKey(legacy)).toBe('subject_keyword::112訴304::03/A')
    const candidate = rule({ type: 'subject_keyword', signal: '112訴304' })
    expect(filterByTombstones([candidate], [legacy])).toHaveLength(0)
    await addRuleTombstones([legacy])
    const res = await addRulesFilteringTombstones([candidate])
    expect(res.added).toHaveLength(0)
  })

  it('the override learner honours a legacy-form tombstone too', async () => {
    // Internal sender → no usable domain → P2 learns the pure case number.
    await setSettings({ internalDomains: ['x.example'] })
    await addRuleTombstones([
      { type: 'subject_keyword', signalNorm: '112年度訴字第304號', targetFolderPath: '03/Right', deletedAt: 1 },
    ])
    const item: PlanItem = {
      emailId: 'e1',
      emailSubject: '112年度訴字第304號 補充理由狀',
      emailFrom: 'a@x.example',
      action: 'move',
      targetFolderPath: '03/Right',
      confidence: 0.7,
      reason: '',
      source: 'ai',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '03/Wrong',
    }
    const r = await generateAiOverrideRules([item], [{ emailId: 'e1', subject: 'X', action: 'move', status: 'moved' }])
    expect(r.added).toBe(0)
  })

  it('re-creating the rule by hand clears its legacy-form tombstone', async () => {
    await addRuleTombstones([
      { type: 'subject_keyword', signalNorm: '112年度訴字第304號', targetFolderPath: '03/A', deletedAt: 1 },
    ])
    await upsertRule(rule({ type: 'subject_keyword', signal: '112訴304', source: 'user_manual' }), { mode: 'create' })
    expect(await getRuleTombstones()).toHaveLength(0)
  })
})

describe('runSweep is one rules transaction (item 16)', () => {
  it('dedupe + closed-case retire/purge + stale in ≤ 1 rules write, same outcomes', async () => {
    const now = Date.now()
    // closedCasePrefixes must be a subset of excludeFolderPrefixes — storage
    // enforces that invariant on every read/write since 2026-11-05.
    await setSettings({ excludeFolderPrefixes: ['05已完成案件'], closedCasePrefixes: ['05已完成案件'] })
    // This machine's counter baseline predates every rule below (2026-11-06),
    // so their whole records count — without it the first sweep would start
    // one and judge them only on later fires (tests/counter-baseline.test.ts).
    await setRuleCounterBaseline({ capturedAt: '2025-12-01T00:00:00.000Z', counts: {} })
    const dupA = rule({ signal: 'dup.example', matchCount: 1, lastUsedAt: new Date(now - DAY).toISOString() })
    const dupB = rule({ signal: 'dup.example', matchCount: 1, lastUsedAt: new Date(now - DAY).toISOString() })
    const toRetire = rule({ signal: 'archived.example', targetFolderPath: '05已完成案件/某案' })
    const toPurge = rule({
      signal: 'closed-long-ago.example',
      targetFolderPath: '05已完成案件/舊案',
      enabled: false,
      autoDisabledReason: 'case_closed',
      autoDisabledAt: new Date(now - 40 * DAY).toISOString(),
    })
    const highError = rule({ signal: 'wrong.example', matchCount: 5, overrideCount: 20, lastUsedAt: new Date(now - DAY).toISOString() })
    await setRules([dupA, dupB, toRetire, toPurge, highError])

    const setSpy = chrome.storage.local.set as unknown as { mock: { calls: Array<[Record<string, unknown>]> } }
    const before = setSpy.mock.calls.length
    const result = await runSweep({ force: true })
    const ruleWrites = setSpy.mock.calls.slice(before).filter(([items]) => 'rules' in items)
    expect(ruleWrites).toHaveLength(1)
    expect(result.disabledCount).toBe(1)

    const after = await listRules()
    expect(after.filter((r) => r.signal === 'dup.example')).toHaveLength(1)
    expect(after.find((r) => r.signal === 'archived.example')?.autoDisabledReason).toBe('case_closed')
    expect(after.find((r) => r.signal === 'closed-long-ago.example')).toBeUndefined()
    expect(after.find((r) => r.signal === 'wrong.example')?.autoDisabledReason).toBe('high-error-rate')
  })

  it('an UNKNOWN archive list skips the closed-case steps instead of reviving everything', async () => {
    const retired = rule({
      targetFolderPath: '05已完成案件/某案',
      enabled: false,
      autoDisabledReason: 'case_closed',
      autoDisabledAt: new Date().toISOString(),
    })
    await setRules([retired])
    await runRuleSweepPass({}) // e.g. settings unreadable → prefixes unknown
    expect((await listRules())[0]!.enabled).toBe(false)
    await runRuleSweepPass({ closedCasePrefixes: [] }) // explicit "no archives" → revive
    expect((await listRules())[0]!.enabled).toBe(true)
  })

  it('a no-op sweep writes nothing', async () => {
    await setRules([rule({ matchCount: 1, lastUsedAt: new Date().toISOString() })])
    const setSpy = chrome.storage.local.set as unknown as { mock: { calls: Array<[Record<string, unknown>]> } }
    const before = setSpy.mock.calls.length
    await runSweep({ force: true })
    expect(setSpy.mock.calls.slice(before).filter(([items]) => 'rules' in items)).toHaveLength(0)
  })

  it('a malformed compound no longer aborts the sweep (item 10)', async () => {
    const now = Date.parse('2026-11-20T03:00:00.000Z')
    const broken = rule({
      type: 'compound',
      signal: '{"conditions":[{"type":"subject_keyword","value":123}]}',
      createdAt: new Date(now - 5 * DAY).toISOString(),
    })
    const stale = rule({ signal: 'dead.example', matchCount: 3, lastUsedAt: new Date(now - 150 * DAY).toISOString() })
    await setRules([broken, stale])
    const report = await autoDisableStaleRules({ now, activeDays: dailyActiveDays(now, 200) })
    expect(report.byReason.stale).toBe(1)
    expect((await listRules()).map((r) => r.id)).toEqual([broken.id])
  })
})

// 2026-11-06: the parser read 「113年度第1次董事會議程」 as court case 「113度1」
// and the learners minted case rules on it. The fixed parser no longer
// matches them, but left enabled such a compound still looked like a free-
// text "narrowing" of the client (it kept the client's broad rule from being
// learned or revived). The sweep removes them, like other dead rules.
describe('rules learned on a 年度 pseudo case number are purged by the sweep', () => {
  const pseudoCompound = (over: Partial<Rule> = {}) =>
    rule({
      id: 'pseudo-compound',
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: '113度1' },
      ]),
      targetFolderPath: '03/董事會',
      matchCount: 4,
      lastUsedAt: new Date().toISOString(),
      ...over,
    })

  it('repro: the AUTO ones go (any state); real case rules and hand-written ones stay', async () => {
    const realCase = rule({
      id: 'real',
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: '112訴304' },
      ]),
      lastUsedAt: new Date().toISOString(),
    })
    const pseudoSubject = rule({ id: 'pseudo-subject', type: 'subject_keyword', signal: '１１３度２', source: 'ai_overridden' })
    const dormantPseudo = pseudoCompound({
      id: 'pseudo-dormant',
      targetFolderPath: '03/股東會',
      enabled: false,
      autoDisabledReason: 'override',
      autoDisabledAt: '2026-11-01T00:00:00.000Z',
    })
    const mine = rule({ id: 'mine', type: 'subject_keyword', signal: '113度1', source: 'user_manual' })
    await setRules([pseudoCompound(), realCase, pseudoSubject, dormantPseudo, mine])
    const report = await runRuleSweepPass({})
    expect(report.pseudoCasePurged).toBe(3)
    expect((await listRules()).map((r) => r.id).sort()).toEqual(['mine', 'real'])
    // A matching 「113年度第1期財報」 no longer reaches the board folder through it.
    expect(matchEmail({ Id: 'm', Subject: '113年度第1期財報', BodyPreview: '', From: { EmailAddress: { Address: 'ir@client.example' } } } as Email, await listRules())).toBeNull()
  })

  it('the daily sweep counts them with the dead rules it deleted', async () => {
    await setRules([pseudoCompound()])
    expect((await runSweep({ force: true })).staleDeletedCount).toBe(1)
    expect(await listRules()).toEqual([])
  })
})
