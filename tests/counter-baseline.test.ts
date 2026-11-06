// Pre-v1.0.8 override counts must not trip high-error-rate (2026-11-06, B7).
//
// de733bd counted a rule-routed row changed to 刪除 as an override. v1.0.8's
// predicate (fires ≥ 20, overrides / fires ≥ 0.5) then disabled broad rules
// on the FIRST sweep that the old one kept (m8/o12, m15/o15), and the dormant
// purge hard-deleted them 60 active days later. Each machine now keeps a
// LOCAL, unsynced baseline: rules that predate it are judged only on fires
// counted after it — and need 20 of them first.

import { describe, expect, it } from 'vitest'
import { runSweep } from '@/background/stale-sweep'
import { clearCloudState, pullNow, resetSyncEngineForTests } from '@/background/sync-engine'
import { serializeRules } from '@/shared/rule-io'
import {
  autoDisableStaleRules,
  bumpRuleHits,
  bumpRuleOverrides,
  dedupeRulesByKey,
  listRules,
  mutateRules,
  newRule,
  revivedRule,
  runRuleSweepPass,
} from '@/shared/rules'
import { getRuleCounterBaseline, setRuleCounterBaseline, setRules, setSettings } from '@/shared/storage'
import type { Rule } from '@/shared/types'
import { simulateOtherMachinePush } from './helpers/sync-sim'
import { loadServiceWorker } from './helpers/sw-harness'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'

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
    createdAt: '2026-06-01T00:00:00.000Z',
    lastUsedAt: new Date().toISOString(),
    ...over,
  }
}

const enabled = async (id: string) => (await listRules()).find((r) => r.id === id)?.enabled

describe('upgrade: legacy counts are baselined on the first sweep', () => {
  it('repro: m8/o12 and m15/o15 survive the first sweep (v1.0.8 disabled both)', async () => {
    const a = rule({ signal: 'a.example', matchCount: 8, overrideCount: 12 })
    const b = rule({ signal: 'b.example', matchCount: 15, overrideCount: 15 })
    await setRules([a, b])
    const result = await runSweep({ force: true })
    expect(result.disabledCount).toBe(0) // was 2
    expect(await enabled(a.id)).toBe(true)
    expect(await enabled(b.id)).toBe(true)
    expect((await getRuleCounterBaseline())?.counts).toEqual({ [a.id]: [8, 12], [b.id]: [15, 15] })
  })

  it('20 fresh fires, half of them overridden, still disable it', async () => {
    const a = rule({ matchCount: 8, overrideCount: 12 })
    await setRules([a])
    await runSweep({ force: true }) // baseline [8, 12]
    await bumpRuleHits(new Map([[a.id, 10]]))
    await bumpRuleOverrides(new Map([[a.id, 9]]))
    await runSweep({ force: true }) // 19 fresh fires — not enough evidence yet
    expect(await enabled(a.id)).toBe(true)
    await bumpRuleOverrides(new Map([[a.id, 1]]))
    await runSweep({ force: true }) // 20 fresh fires, 10 overridden
    expect((await listRules())[0]).toMatchObject({ enabled: false, autoDisabledReason: 'high-error-rate' })
  })

  it('fresh evidence that is mostly right keeps it, whatever the legacy total says', async () => {
    const a = rule({ matchCount: 8, overrideCount: 12 })
    await setRules([a])
    await runSweep({ force: true })
    await bumpRuleHits(new Map([[a.id, 12]]))
    await bumpRuleOverrides(new Map([[a.id, 8]])) // total m20/o20 = 50%, fresh 8/20 = 40%
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(true)
  })
})

describe('which rules the baseline covers', () => {
  it('a rule created after the baseline started is judged on its whole record', async () => {
    await setRuleCounterBaseline({ capturedAt: '2026-11-01T00:00:00.000Z', counts: {} })
    const fresh = rule({ createdAt: '2026-11-02T00:00:00.000Z', matchCount: 5, overrideCount: 20 })
    await setRules([fresh])
    await runSweep({ force: true })
    expect(await enabled(fresh.id)).toBe(false)
  })

  it('an OLD rule first seen later (e.g. arriving by sync) is baselined at first sight', async () => {
    await setRuleCounterBaseline({ capturedAt: '2026-11-01T00:00:00.000Z', counts: {} })
    const legacy = rule({ createdAt: '2026-05-01T00:00:00.000Z', matchCount: 5, overrideCount: 20 })
    await setRules([legacy])
    await runSweep({ force: true })
    expect(await enabled(legacy.id)).toBe(true)
    expect((await getRuleCounterBaseline())?.counts[legacy.id]).toEqual([5, 20])
  })

  it('counters reset by a revive count in full from then on; deleted rules leave the baseline', async () => {
    const a = rule({ matchCount: 8, overrideCount: 12 })
    const gone = rule({ signal: 'gone.example', matchCount: 1 })
    await setRules([a, gone])
    await runSweep({ force: true })
    // Revive's fresh window: (1, 0) — below the [8, 12] entry.
    await setRules([{ ...revivedRule({ ...a, enabled: false, autoDisabledReason: 'high-error-rate' }, new Date().toISOString()), matchCount: 1, overrideCount: 0 }])
    await runSweep({ force: true })
    expect((await getRuleCounterBaseline())?.counts).toEqual({ [a.id]: [0, 0] })
    await bumpRuleOverrides(new Map([[a.id, 19]])) // 20 fires since the reset, 19 overridden
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(false)
  })

  it('user_manual rules are never baselined (never swept either)', async () => {
    await setRules([rule({ source: 'user_manual', matchCount: 2, overrideCount: 30 })])
    await runSweep({ force: true })
    expect((await getRuleCounterBaseline())?.counts).toEqual({})
  })

  it('a revive that arrives with a lower epoch is not a reset: counters below the entry re-base', async () => {
    const a = rule({ matchCount: 8, overrideCount: 12, statsResetAt: 5_000 })
    await setRules([a])
    await runSweep({ force: true })
    expect((await getRuleCounterBaseline())?.epochs).toEqual({ [a.id]: 5_000 })
    // An older copy (earlier epoch, its own counts) replaces it wholesale.
    await setRules([{ ...a, statsResetAt: 1_000, matchCount: 2, overrideCount: 30 }])
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(true)
    expect((await getRuleCounterBaseline())?.counts[a.id]).toEqual([2, 30])
  })

  it('direct calls without a stored baseline keep judging whole records (no implicit start)', async () => {
    const a = rule({ matchCount: 8, overrideCount: 12 })
    await setRules([a])
    expect((await autoDisableStaleRules({ staleDays: 365 })).byReason.highErrorRate).toBe(1)
    await setRules([rule({ id: 'b', matchCount: 8, overrideCount: 12 })])
    expect((await runRuleSweepPass({})).stale.byReason.highErrorRate).toBe(1)
    expect(await getRuleCounterBaseline()).toBeNull()
  })
})

// R2 (2026-11-06): three ways the legacy record still got judged whole.
describe('the baseline cannot be bypassed (R2)', () => {
  it('(a) counters that went DOWN without a new epoch (replace-import / restore) re-base instead of resetting', async () => {
    const a = rule({ matchCount: 30, overrideCount: 14 })
    await setRules([a])
    await runSweep({ force: true }) // baseline [30, 14]
    // An older export carrying the pre-v1.0.8 record replaces the rule.
    await setRules([{ ...a, matchCount: 8, overrideCount: 12 }])
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(true) // was: "reset" → 8/12 judged whole → high-error-rate
    expect((await getRuleCounterBaseline())?.counts[a.id]).toEqual([8, 12])
    // Not an immunity: 20 fires counted here, half overridden, still disable it.
    await bumpRuleHits(new Map([[a.id, 10]]))
    await bumpRuleOverrides(new Map([[a.id, 10]]))
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(false)
  })

  it('(a) the same through the real import (replace): raised counts are covered too', async () => {
    const sw = await loadServiceWorker()
    await runSweep() // the SW's own wake-up sweep (module load) — let it finish on the empty library
    const a = rule({ matchCount: 30, overrideCount: 2 })
    await setRules([a])
    await runSweep({ force: true }) // baseline [30, 2]
    const payload = serializeRules([{ ...a, matchCount: 30, overrideCount: 40 }], '1.0.7')
    expect((await sw.send({ type: 'importRules', payload, strategy: 'replace' })).ok).toBe(true)
    await runSweep({ force: true })
    expect(await enabled(a.id)).toBe(true) // was: 38 "fresh" overrides → high-error-rate
  })

  it('(b) a twin with inflated overrides merged by dedupe does not read as fresh evidence', async () => {
    const survivor = rule({ id: 's', matchCount: 30, overrideCount: 2 })
    await setRules([survivor])
    await runSweep({ force: true }) // baseline s: [30, 2]
    // A legacy same-triple twin shows up (import / sync from an old library).
    const twin = rule({ id: 't', createdAt: '2026-05-01T00:00:00.000Z', matchCount: 8, overrideCount: 25 })
    await setRules([survivor, twin])
    const report = await runSweep({ force: true })
    expect(report.disabledCount).toBe(0) // was 1: [30, 25] vs base [30, 2] → 23 "fresh" overrides
    const [only] = await listRules()
    expect(only).toMatchObject({ id: 's', enabled: true, matchCount: 30, overrideCount: 25 })
    expect((await getRuleCounterBaseline())?.counts.s).toEqual([30, 25])
  })

  it('(b) dedupeRulesByKey keeps the baseline in step the same way', async () => {
    const survivor = rule({ id: 's', matchCount: 30, overrideCount: 2 })
    await setRules([survivor])
    await runSweep({ force: true })
    await setRules([survivor, rule({ id: 't', createdAt: '2026-05-01T00:00:00.000Z', matchCount: 8, overrideCount: 25 })])
    expect((await dedupeRulesByKey()).removed).toBe(1)
    expect((await getRuleCounterBaseline())?.counts.s).toEqual([30, 25])
  })

  it('(c) overrides another machine recorded arrive by sync pull already covered; local ones still count', async () => {
    resetSyncEngineForTests()
    await clearCloudState()
    await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
    const a = rule({ id: 'shared', matchCount: 10, overrideCount: 2 })
    await setRules([a])
    await runSweep({ force: true }) // baseline [10, 2]
    // The peer's copy carries its legacy overrides; sync max-merges counters.
    await simulateOtherMachinePush({ rules: [{ ...a, overrideCount: 30 }] })
    await pullNow('remote', { mode: 'union' })
    expect((await listRules())[0]).toMatchObject({ matchCount: 10, overrideCount: 30 })
    expect((await getRuleCounterBaseline())?.counts.shared).toEqual([10, 30])
    await runSweep({ force: true })
    expect(await enabled('shared')).toBe(true) // was: 28 "fresh" overrides → high-error-rate
    // What THIS machine counts afterwards is judged as before.
    await bumpRuleHits(new Map([['shared', 10]]))
    await bumpRuleOverrides(new Map([['shared', 10]]))
    await runSweep({ force: true })
    expect(await enabled('shared')).toBe(false)
  })

  it('(c) a rule new to this machine arrives covered whole, whatever its age', async () => {
    await setRuleCounterBaseline({ capturedAt: '2026-11-01T00:00:00.000Z', counts: {} })
    const fresh = rule({ id: 'peer-made', createdAt: '2026-11-03T00:00:00.000Z', matchCount: 5, overrideCount: 20 })
    await mutateRules((rules) => ({ next: [...rules, fresh], result: undefined }), { preserveEditedAt: true })
    expect((await getRuleCounterBaseline())?.counts['peer-made']).toEqual([5, 20])
    await runSweep({ force: true })
    expect(await enabled('peer-made')).toBe(true)
  })

  it('a pulled revive (newer epoch) is covered; a revive on THIS machine still opens a fresh window', async () => {
    const a = rule({ id: 'r', matchCount: 8, overrideCount: 12 })
    await setRules([a])
    await runSweep({ force: true })
    // Another machine revived it and counted 5 / 15 since.
    const pulled = { ...a, statsResetAt: Date.now(), matchCount: 5, overrideCount: 15 }
    await mutateRules((rules) => ({ next: rules.map(() => pulled), result: undefined }), { preserveEditedAt: true })
    await runSweep({ force: true })
    expect(await enabled('r')).toBe(true)
    expect((await getRuleCounterBaseline())?.counts.r).toEqual([5, 15])
  })
})
