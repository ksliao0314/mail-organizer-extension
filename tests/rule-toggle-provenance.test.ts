// One enable / disable transition for every user path (2026-11-06).
//
// B2 — bulk enable / disable (multi-select, folder-group "disable all",
//      resolveConflict disable_all, split-to-compound originals) mirrored the
//      OLD toggleRule and ignored the closed-case exemption:
//        (a) bulk-enabling a rule retired by 結案歸檔 dropped the 'case_closed'
//            marker → the next preflight retired it again;
//        (b) bulk-disabling an exempt rule kept the marker → when the folder
//            left the archive, Phase 1 re-enabled a rule the user turned off.
// B3 — a USER disable kept old auto provenance (toggleRule, the rule form and
//      the bulk path): a stale reason on an enabled rule (the old drawer could
//      save enabled + reason) made the user's disable look like an engine one
//      — revived by learning, purged by the dormant sweep.
// B11 — split-to-compound cleared tombstones by raw key; a tombstone an older
//      build wrote in a pre-canonical form for the new compound survived.

import { describe, expect, it } from 'vitest'
import { generateAiConfirmedRules, type ExecuteItemResult } from '@/background/execute'
import {
  handleBulkRuleAction,
  handleResolveConflict,
  handleSplitConflictToCompound,
} from '@/background/handlers/rules-bulk'
import {
  applyClosedCaseRetirement,
  autoDisableStaleRules,
  isRevivableAutoDisabled,
  listRules,
  newRule,
  toggleRule,
  upsertRule,
  withToggle,
} from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, setRules } from '@/shared/storage'
import type { PlanItem, Rule } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'

const ARCHIVE = '05已完成案件'
const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const NOW = Date.parse('2026-11-20T03:00:00.000Z')

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

const retired = (over: Partial<Rule> = {}) =>
  rule({
    targetFolderPath: `${ARCHIVE}/某案`,
    enabled: false,
    autoDisabledReason: 'case_closed',
    autoDisabledAt: '2026-11-01T00:00:00.000Z',
    ...over,
  })

/** Enabled, but still wearing an auto reason (the old drawer could save that). */
const staleProvenance = (over: Partial<Rule> = {}) =>
  rule({ autoDisabledReason: 'high-error-rate', autoDisabledAt: '2026-10-01T00:00:00.000Z', ...over })

const bulk = (op: 'enable' | 'disable', ids: string[]) => handleBulkRuleAction({ op, ids })

describe('withToggle — the shared transition', () => {
  it('enable keeps only the case_closed exemption marker; disable strips every reason', () => {
    expect(withToggle(retired(), true)).toMatchObject({ enabled: true, autoDisabledReason: 'case_closed' })
    expect(withToggle(retired(), true).autoDisabledAt).toBeUndefined()
    const off = withToggle(staleProvenance(), false)
    expect(off.enabled).toBe(false)
    expect(off.autoDisabledReason).toBeUndefined()
    expect(off.autoDisabledAt).toBeUndefined()
    const r = rule()
    expect(withToggle(r, true)).toBe(r) // no-op keeps the ref
  })
})

describe('bulk paths honour the closed-case exemption (B2)', () => {
  it('(a) bulk-enabling a retired rule keeps it live through the next retirement pass', async () => {
    const r = retired()
    await setRules([r])
    expect(await bulk('enable', [r.id])).toMatchObject({ ok: true, data: { changed: 1 } })
    expect((await listRules())[0]).toMatchObject({ enabled: true, autoDisabledReason: 'case_closed' })
    const second = await applyClosedCaseRetirement([ARCHIVE])
    expect(second.retired).toBe(0) // was 1
    expect((await listRules())[0]!.enabled).toBe(true)
  })

  it('(b) bulk-disabling an exempt rule drops the marker — un-archiving does not re-enable it', async () => {
    const r = retired()
    await setRules([r])
    await toggleRule(r.id, true, { actor: 'user' }) // exempt
    await bulk('disable', [r.id])
    const off = (await listRules())[0]!
    expect(off.autoDisabledReason).toBeUndefined() // was 'case_closed'
    await setRules([{ ...off, targetFolderPath: '03進行中案件/某案' }])
    expect((await applyClosedCaseRetirement([ARCHIVE])).revived).toBe(0) // was 1
    expect((await listRules())[0]!.enabled).toBe(false)
  })

  it('resolveConflict disable_all and split-to-compound originals go through the same transition', async () => {
    const a = retired({ signal: 'dup.example' })
    const b = rule({ signal: 'dup.example', targetFolderPath: '03/B' })
    await setRules([a, b])
    await toggleRule(a.id, true, { actor: 'user' }) // exempt marker on a
    await handleResolveConflict({ ruleIds: [a.id, b.id], strategy: 'disable_all' })
    expect((await listRules()).find((x) => x.id === a.id)!.autoDisabledReason).toBeUndefined()

    const c = staleProvenance({ signal: 'split.example' })
    const d = rule({ signal: 'split.example', targetFolderPath: '03/B' })
    await setRules([c, d])
    await handleSplitConflictToCompound({
      items: [
        { ruleId: c.id, keywords: ['甲案合約審閱'] },
        { ruleId: d.id, keywords: [] },
      ],
    })
    const original = (await listRules()).find((x) => x.id === c.id)!
    expect(original.enabled).toBe(false)
    expect(original.autoDisabledReason).toBeUndefined() // was 'high-error-rate' (looked engine-disabled)
  })
})

describe('a user disable carries no auto provenance (B3)', () => {
  const item = (): PlanItem => ({
    emailId: 'e1',
    emailSubject: '一般往來信件內容',
    emailFrom: 'billing@client.example',
    action: 'move',
    targetFolderPath: '03/A',
    confidence: 0.9,
    reason: '',
    source: 'ai',
  })
  const moved: ExecuteItemResult = { emailId: 'e1', subject: 'X', action: 'move', status: 'moved' }

  for (const [name, disable] of [
    ['toggleRule', (r: Rule) => toggleRule(r.id, false, { actor: 'user' })],
    ['the rule form', (r: Rule) => upsertRule({ ...r, enabled: false }, { actor: 'user', mode: 'edit' })],
    ['bulk disable', (r: Rule) => bulk('disable', [r.id])],
  ] as const) {
    it(`${name}: not revived by learning, not purged by the dormant sweep`, async () => {
      const r = staleProvenance()
      await setRules([r])
      await disable(r)
      const off = (await listRules())[0]!
      expect(off.enabled).toBe(false)
      expect(off.autoDisabledReason).toBeUndefined()
      expect(off.autoDisabledAt).toBeUndefined()
      expect(isRevivableAutoDisabled(off)).toBe(false)

      expect((await generateAiConfirmedRules([item()], [moved])).revived).toBe(0)
      const report = await autoDisableStaleRules({ now: NOW, activeDays: dailyActiveDays(NOW, 200) })
      expect(report.byReason.dormantPurged).toBe(0)
      expect(await listRules()).toHaveLength(1)
      expect((await listRules())[0]!.enabled).toBe(false)
    })
  }
})

describe('bulk add clears legacy-form tombstones (B11)', () => {
  it('split-to-compound removes a pre-canonical tombstone of the new compound triple', async () => {
    const a = rule({ signal: 'dup.example', source: 'user_manual' })
    const b = rule({ signal: 'dup.example', targetFolderPath: '03/B', source: 'user_manual' })
    await setRules([a, b])
    // Written by an older build: full-form case number, not folded to 112訴304.
    await addRuleTombstones([
      {
        type: 'compound',
        signalNorm: '[{"type":"domain","value":"dup.example"},{"type":"subject_keyword","value":"112年度訴字第304號"}]',
        targetFolderPath: '03/A',
        deletedAt: 1,
      },
    ])
    const res = await handleSplitConflictToCompound({
      items: [
        { ruleId: a.id, keywords: ['112年度訴字第304號'] },
        { ruleId: b.id, keywords: [] },
      ],
    })
    expect(res).toMatchObject({ ok: true, data: { created: 1 } })
    expect(await getRuleTombstones()).toHaveLength(0) // the legacy ghost survived the raw-key clear
  })
})
