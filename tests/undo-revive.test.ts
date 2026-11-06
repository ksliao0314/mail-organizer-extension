// 撤回 also reverses the revives its batch's learning made (2026-11-06, B6).
//
// Revived ids are deliberately NOT in mintedRuleIds (undo deletes those, and
// must never delete a rule older than the batch) — so undoing a misfiled
// batch left the rule it revived live, and it re-proposed the misfile next
// batch. The undo snapshot now records each revive with its pre-revive
// provenance; undo restores it unless the rule changed since.

import { beforeEach, describe, expect, it, vi } from 'vitest'
import { executeUndo, startExecute } from '@/background/execute'
import { OutlookError, type OutlookApi } from '@/shared/outlook-api'
import { listRules, newRule, toggleRule, upsertRule } from '@/shared/rules'
import { getUndoSnapshot, setRules } from '@/shared/storage'
import type { MailFolderNode, PlanItem, Rule } from '@/shared/types'

const CASES = 'FOLDER-CASES-0000000000000'
const A = 'FOLDER-A-00000000000000000'
const DISABLED_AT = '2026-10-01T00:00:00.000Z'

const tree = (): MailFolderNode[] => [
  {
    id: CASES,
    displayName: '案件',
    path: '案件',
    children: [{ id: A, displayName: '甲', path: '案件/甲', parentFolderId: CASES, children: [] }],
  },
]

function fakeApi() {
  const gone = new Set<string>()
  return {
    moveMessage: vi.fn(async (id: string, dest: string) => {
      if (gone.has(id)) throw new OutlookError(404, 'ErrorItemNotFound')
      gone.add(id)
      return { Id: `new-${id}`, ParentFolderId: dest }
    }),
    listChildFolders: vi.fn(async () => []),
    listTopFolders: vi.fn(async () => []),
  }
}
const asApi = (f: object) => f as unknown as OutlookApi

const dormant = (): Rule => ({
  ...newRule({
    type: 'domain',
    signal: 'client.example',
    targetFolderId: A,
    targetFolderPath: '案件/甲',
    confidence: 0.7,
    source: 'ai_confirmed',
  }),
  id: 'dormant',
  createdAt: '2026-01-01T00:00:00.000Z',
  enabled: false,
  autoDisabledReason: 'high-error-rate',
  autoDisabledAt: DISABLED_AT,
  matchCount: 3,
  overrideCount: 20,
})

/** The AI files a client.example mail into 甲 — its learning revives `dormant`. */
const aiRow = (): PlanItem => ({
  emailId: 'e1',
  emailSubject: '往來信件 e1',
  emailFrom: 'clerk@client.example',
  action: 'move',
  targetFolderId: A,
  targetFolderPath: '案件/甲',
  confidence: 0.9,
  reason: '',
  source: 'ai',
})

beforeEach(() => {
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.alarms = {
    create: vi.fn(async () => {}),
    clear: vi.fn(async () => true),
    get: vi.fn(async () => undefined),
    onAlarm: { addListener: vi.fn() },
  }
})

const stored = async () => (await listRules()).find((r) => r.id === 'dormant')!

describe('undo puts revived rules back to sleep (B6)', () => {
  it('repro: the batch revives the rule; 撤回 restores its old provenance', async () => {
    await setRules([dormant()])
    const api = fakeApi()
    await startExecute([aiRow()], tree(), asApi(api))
    expect((await stored()).enabled).toBe(true) // revived by learning
    expect((await getUndoSnapshot())?.revivedRules).toEqual([
      expect.objectContaining({ ruleId: 'dormant', autoDisabledReason: 'high-error-rate', autoDisabledAt: DISABLED_AT }),
    ])

    expect(await executeUndo(asApi(api))).toMatchObject({ restored: 1 })
    const after = await stored()
    expect(after.enabled).toBe(false) // was: still live → the misfile re-proposed next batch
    expect(after.autoDisabledReason).toBe('high-error-rate')
    expect(after.autoDisabledAt).toBe(DISABLED_AT)
  })

  it('a rule the user edited after the batch is left alone', async () => {
    await setRules([dormant()])
    const api = fakeApi()
    await startExecute([aiRow()], tree(), asApi(api))
    const live = await stored()
    await upsertRule({ ...live, confidence: 0.6 }, { actor: 'user', mode: 'edit' })
    await executeUndo(asApi(api))
    expect(await stored()).toMatchObject({ enabled: true, confidence: 0.6 })
  })

  it('a rule the user switched off after the batch stays a USER disable', async () => {
    await setRules([dormant()])
    const api = fakeApi()
    await startExecute([aiRow()], tree(), asApi(api))
    await toggleRule('dormant', false, { actor: 'user' })
    await executeUndo(asApi(api))
    const after = await stored()
    expect(after.enabled).toBe(false)
    expect(after.autoDisabledReason).toBeUndefined()
  })

  it('nothing restored (every move failed) → revives are kept, like minted rules', async () => {
    await setRules([dormant()])
    const api = fakeApi()
    await startExecute([aiRow()], tree(), asApi(api))
    api.moveMessage.mockRejectedValue(new OutlookError(403, 'denied'))
    expect(await executeUndo(asApi(api))).toMatchObject({ restored: 0, failed: 1 })
    expect((await stored()).enabled).toBe(true)
  })

  // R7 (2026-11-06): only the AI-confirmed learner's revives were recorded;
  // one made by the override learner stayed live after 撤回.
  it('a revive made by the OVERRIDE learner is put back to sleep too', async () => {
    const sender: Rule = { ...dormant(), type: 'sender', signal: 'boss@gmail.com', source: 'ai_overridden' }
    await setRules([sender])
    // The AI said 乙; the user changed the row to 甲 — the dormant 甲 mapping.
    const corrected: PlanItem = {
      ...aiRow(),
      emailFrom: 'boss@gmail.com',
      userTouched: true,
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '案件/乙',
    }
    const api = fakeApi()
    await startExecute([corrected], tree(), asApi(api))
    expect((await stored()).enabled).toBe(true) // revived by the override learner
    expect((await getUndoSnapshot())?.revivedRules).toEqual([
      expect.objectContaining({ ruleId: 'dormant', autoDisabledReason: 'high-error-rate' }),
    ]) // was undefined
    await executeUndo(asApi(api))
    expect(await stored()).toMatchObject({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: DISABLED_AT })
  })
})

// H2 consistency (2026-11-06): 撤回 means "these filings were wrong", so it
// reverses everything the batch's learning did to the library — the rules
// the OVERRIDE learner minted (only the AI-confirmed ones were deleted), and
// the rules the batch switched off: 'superseded' / 'override' demotions and
// the one-strike disable of a case rule whose row was re-filed elsewhere.
// Counters stay; a rule changed since the batch is left alone.
describe('撤回 reverses every rule change the batch made (H2)', () => {
  const B = 'FOLDER-B-00000000000000000'
  const tree2 = (): MailFolderNode[] => [
    {
      id: CASES,
      displayName: '案件',
      path: '案件',
      children: [
        { id: A, displayName: '甲', path: '案件/甲', parentFolderId: CASES, children: [] },
        { id: B, displayName: '乙', path: '案件/乙', parentFolderId: CASES, children: [] },
      ],
    },
  ]
  const ruleOf = (id: string, over: Partial<Rule>): Rule => ({
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: B,
      targetFolderPath: '案件/乙',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    id,
    createdAt: '2026-01-01T00:00:00.000Z',
    matchCount: 5,
    ...over,
  })
  const byId = async (id: string) => (await listRules()).find((r) => r.id === id)
  /** The AI said 乙; the user changed the row to 甲. */
  const overriddenTo甲 = (over: Partial<PlanItem> = {}): PlanItem => ({
    ...aiRow(),
    userTouched: true,
    aiOriginalAction: 'move',
    aiOriginalTargetFolderPath: '案件/乙',
    ...over,
  })

  it('repro: a rule the OVERRIDE learner minted is deleted', async () => {
    await setRules([])
    const api = fakeApi()
    await startExecute([overriddenTo甲({ emailFrom: 'boss@gmail.com' })], tree2(), asApi(api))
    const [minted] = await listRules()
    expect(minted).toMatchObject({ type: 'sender', source: 'ai_overridden', targetFolderPath: '案件/甲' })
    expect((await getUndoSnapshot())?.mintedRuleIds).toEqual([minted!.id]) // was undefined
    await executeUndo(asApi(api))
    expect(await listRules()).toEqual([]) // was: kept
  })

  it("repro: the broad rule the override learner superseded is switched back on, its counters untouched", async () => {
    await setRules([ruleOf('broad', {})])
    const api = fakeApi()
    await startExecute([overriddenTo甲({ emailSubject: '甲案合約審閱意見回覆' })], tree2(), asApi(api))
    expect(await byId('broad')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded' })
    expect((await listRules()).filter((r) => r.type === 'compound')).toHaveLength(1)

    await executeUndo(asApi(api))
    const broad = (await byId('broad'))!
    expect(broad.enabled).toBe(true) // was: left superseded
    expect(broad.autoDisabledReason).toBeUndefined()
    expect(broad.autoDisabledAt).toBeUndefined()
    expect(broad.matchCount).toBe(5)
    expect(await listRules()).toHaveLength(1) // the minted compound is gone
  })

  it('repro: the one-strike disable of a re-filed case rule is reversed', async () => {
    const caseRule = ruleOf('case', {
      type: 'case_code',
      signal: '25A0068B',
      targetFolderId: A,
      targetFolderPath: '案件/甲',
      confidence: 0.9,
    })
    await setRules([caseRule])
    const api = fakeApi()
    const refiled: PlanItem = {
      ...aiRow(),
      emailSubject: '25A0068B 補件通知',
      targetFolderId: B,
      targetFolderPath: '案件/乙',
      originalRuleId: 'case',
      userTouched: true,
    }
    await startExecute([refiled], tree2(), asApi(api))
    expect(await byId('case')).toMatchObject({ enabled: false, autoDisabledReason: 'override', overrideCount: 1 })

    await executeUndo(asApi(api))
    expect(await byId('case')).toMatchObject({ enabled: true, overrideCount: 1 }) // was: still off
    expect((await byId('case'))!.autoDisabledReason).toBeUndefined()
    // …and the case rule the correction taught (→ 乙) is gone with it.
    expect((await listRules()).map((r) => r.id)).toEqual(['case'])
  })

  it('a rule switched off, revived and superseded again within one batch is put back as it was BEFORE the batch', async () => {
    const caseRule = ruleOf('case', {
      type: 'case_code',
      signal: '25A0068B',
      targetFolderId: A,
      targetFolderPath: '案件/甲',
      confidence: 0.9,
    })
    await setRules([caseRule])
    const api = fakeApi()
    const plan: PlanItem[] = [
      // re-filed elsewhere → one-strike 'override', then the correction supersedes it…
      {
        ...aiRow(),
        emailId: 'e1',
        emailSubject: '25A0068B 補件通知',
        targetFolderId: B,
        targetFolderPath: '案件/乙',
        originalRuleId: 'case',
        userTouched: true,
      },
      // …while the AI's filing of another 25A0068B mail into 甲 revives it in between.
      { ...aiRow(), emailId: 'e2', emailSubject: '25A0068B 開庭通知' },
    ]
    await startExecute(plan, tree2(), asApi(api))
    expect(await byId('case')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded' })
    expect((await getUndoSnapshot())?.disabledRules).toEqual([expect.objectContaining({ ruleId: 'case' })])
    expect((await getUndoSnapshot())?.revivedRules).toBeUndefined()

    await executeUndo(asApi(api))
    const restored = (await byId('case'))!
    expect(restored.enabled).toBe(true)
    expect(restored.autoDisabledReason).toBeUndefined()
  })

  it('a dormant rule the AI-confirmed learner revived and the override learner superseded goes back to its old sleep', async () => {
    await setRules([
      ruleOf('case', {
        type: 'case_code',
        signal: '25A0068B',
        targetFolderId: A,
        targetFolderPath: '案件/甲',
        confidence: 0.9,
        enabled: false,
        autoDisabledReason: 'high-error-rate',
        autoDisabledAt: DISABLED_AT,
      }),
    ])
    const api = fakeApi()
    const plan: PlanItem[] = [
      { ...aiRow(), emailId: 'e1', emailSubject: '25A0068B 開庭通知' }, // AI → 甲: revives it
      {
        ...aiRow(),
        emailId: 'e2',
        emailSubject: '25A0068B 補件通知',
        targetFolderId: B,
        targetFolderPath: '案件/乙',
        userTouched: true,
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: '案件/甲',
      }, // the user moved this one to 乙: supersedes it again
    ]
    await startExecute(plan, tree2(), asApi(api))
    expect(await byId('case')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded' })
    await executeUndo(asApi(api))
    expect(await byId('case')).toMatchObject({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: DISABLED_AT })
    expect((await listRules()).map((r) => r.id)).toEqual(['case'])
  })

  it('a rule the user changed after the batch is left alone', async () => {
    await setRules([ruleOf('broad', {})])
    const api = fakeApi()
    await startExecute([overriddenTo甲({ emailSubject: '甲案合約審閱意見回覆' })], tree2(), asApi(api))
    const off = (await byId('broad'))!
    await upsertRule({ ...off, confidence: 0.5 }, { actor: 'user', mode: 'edit' })
    await executeUndo(asApi(api))
    expect(await byId('broad')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded', confidence: 0.5 })
  })
})
