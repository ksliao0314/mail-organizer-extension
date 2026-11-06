// Birth-gate crash window (2026-11-06, B10).
//
// If the worker died after recordSubjectFilings wrote but before the
// finalize step marker was saved, the replay snapshotted subject memory that
// ALREADY held this batch: the filing was counted twice and the birth gate
// saw "filed here before" — a full_subject rule minted on first sight. The
// pre-batch values of the batch's subject keys now go into the marker BEFORE
// the recording; a replay puts them back and records once.

import { beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import { executeUndo, recoverStaleExecuteState, startExecute } from '@/background/execute'
import { normalizeSubject } from '@/shared/normalize'
import type { OutlookApi } from '@/shared/outlook-api'
import { listRules, newRule } from '@/shared/rules'
import { clearUndoSnapshot, getSubjectMemory, getUndoSnapshot, setRules, setSettings } from '@/shared/storage'
import type { ExecuteState, MailFolderNode, PlanItem, Rule, ThreadMemoryEntry } from '@/shared/types'

const A = 'FOLDER-A-00000000000000000'
const PATH = '案件/甲'
const SUBJECT = '本週行政事項公告與提醒'
const KEY = normalizeSubject(SUBJECT)

/** Internal colleague mail → P6 sender + full-subject compound, birth-gated for source 'ai'. */
const row = (): PlanItem => ({
  emailId: 'e1',
  emailSubject: SUBJECT,
  emailFrom: 'colleague@firm.example',
  action: 'move',
  targetFolderId: A,
  targetFolderPath: PATH,
  confidence: 0.9,
  reason: '',
  source: 'ai',
  conversationId: 'c1',
})

const entry = (timesFiled: number): ThreadMemoryEntry => ({
  folderId: A,
  folderPath: PATH,
  lastFiledAt: new Date().toISOString(),
  timesFiled,
  conflictCount: 0,
  stableStreak: timesFiled,
})

beforeEach(async () => {
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.alarms = {
    create: vi.fn(async () => {}),
    clear: vi.fn(async () => true),
    get: vi.fn(async () => undefined),
    onAlarm: { addListener: vi.fn() },
  }
  await setSettings({ internalDomains: ['firm.example'] })
})

/**
 * A batch whose worker died inside the finalize steps: the move landed,
 * 'rules' / 'ledgers' / 'convMemory' were marked done, and the marker already
 * holds the pre-batch subject values (`prior`). `memory` is what subject
 * memory holds at death — with or without this batch's filing.
 */
async function seedDeathInSubjectStep(opts: {
  prior: ThreadMemoryEntry | null
  memory: ThreadMemoryEntry | null
}): Promise<void> {
  const plan = [row()]
  const startedAt = Date.now() - 5_000
  const state: ExecuteState = {
    inProgress: true,
    cancelRequested: false,
    startedAt,
    total: 1,
    current: 1,
    plan,
    results: [
      {
        emailId: 'e1',
        subject: SUBJECT,
        action: 'move',
        status: 'moved',
        newMessageId: 'new-e1',
        destinationFolderId: A,
        destinationFolderPath: PATH,
      },
    ],
    summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
    rulesAdded: 0,
  }
  await chrome.storage.session.set({
    executeState: state,
    executeFinalize: { startedAt, done: ['rules', 'ledgers', 'convMemory'], subjectPrior: { [KEY]: opts.prior } },
  })
  await chrome.storage.local.set({ subjectMemory: opts.memory ? { [KEY]: opts.memory } : {} })
}

describe('subjectMemory + birth gate replay is idempotent (B10)', () => {
  it('repro: death AFTER the recording — counted once, no rule on first sight', async () => {
    await seedDeathInSubjectStep({ prior: null, memory: entry(1) })
    expect(await recoverStaleExecuteState()).toMatchObject({ recovered: true })
    expect((await getSubjectMemory())[KEY]?.timesFiled).toBe(1) // was 2
    expect(await listRules()).toHaveLength(0) // was a full_subject compound minted on first sight
  })

  it('death BEFORE the recording — recorded once, still no rule', async () => {
    await seedDeathInSubjectStep({ prior: null, memory: null })
    await recoverStaleExecuteState()
    expect((await getSubjectMemory())[KEY]?.timesFiled).toBe(1)
    expect(await listRules()).toHaveLength(0)
  })

  it('a genuine earlier filing still opens the gate — and is not double-counted either', async () => {
    await seedDeathInSubjectStep({ prior: entry(1), memory: entry(2) })
    await recoverStaleExecuteState()
    expect((await getSubjectMemory())[KEY]?.timesFiled).toBe(2) // was 3
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.type).toBe('compound')
  })
})

describe('normal run', () => {
  it('persists the pre-batch values BEFORE recording, and drops them once learning ran', async () => {
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const markers: Array<{ done: string[]; subjectPrior?: unknown }> = []
    const original = sessionSet.getMockImplementation()!
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      const m = items.executeFinalize as { done: string[]; subjectPrior?: unknown } | undefined
      if (m) markers.push(JSON.parse(JSON.stringify(m)))
      return original(items)
    })
    try {
      const tree: MailFolderNode[] = [
        { id: 'CASES-00000000000000000000', displayName: '案件', path: '案件', children: [{ id: A, displayName: '甲', path: PATH, children: [] }] },
      ]
      const api = {
        moveMessage: vi.fn(async (id: string, dest: string) => ({ Id: `new-${id}`, ParentFolderId: dest })),
        listChildFolders: vi.fn(async () => []),
        listTopFolders: vi.fn(async () => []),
      }
      await startExecute([row()], tree, api as unknown as OutlookApi)
    } finally {
      sessionSet.mockImplementation(original)
    }
    const firstWithPrior = markers.find((m) => m.subjectPrior !== undefined)
    expect(firstWithPrior?.done).not.toContain('subjectMemory')
    expect(firstWithPrior?.subjectPrior).toEqual({ [KEY]: null })
    const afterLearning = markers.find((m) => m.done.includes('learning'))
    expect(afterLearning?.subjectPrior).toBeUndefined()
    expect((await getSubjectMemory())[KEY]?.timesFiled).toBe(1)
  })
})

// R8 (2026-11-06): the 'learning' step was not exactly-once. A death after
// the learner's rule write but before the step marker replayed the learner,
// which deduplicated against its own landed rules — mintedRuleIds / revives
// came back empty, so 撤回 could neither delete the learned rule nor put the
// revived one back to sleep, and the DoneScreen listed nothing learned.
describe('learning step replay is exactly-once (R8)', () => {
  const B = 'FOLDER-B-00000000000000000'
  const tree = (): MailFolderNode[] => [
    {
      id: 'CASES-00000000000000000000',
      displayName: '案件',
      path: '案件',
      children: [
        { id: A, displayName: '甲', path: PATH, children: [] },
        { id: B, displayName: '乙', path: '案件/乙', children: [] },
      ],
    },
  ]
  /** External-client mail the AI files into 甲 → domain rule (first sight). */
  const clientRow = (): PlanItem => ({ ...row(), emailFrom: 'clerk@client.example', emailSubject: '往來信件 e1' })
  const fakeApi = () => ({
    moveMessage: vi.fn(async (id: string, dest: string) => ({ Id: `new-${id}`, ParentFolderId: dest })),
    listChildFolders: vi.fn(async () => []),
    listTopFolders: vi.fn(async () => []),
  })
  type Intent = { added: Array<{ id: string }>; revived: unknown[]; disabled?: unknown[] }
  type Marker = {
    done: string[]
    learningIntent?: { confirmed?: Intent; override?: Intent }
    ruleDisables?: unknown[]
  }

  /** Run a real batch, recording every finalize-marker write (and what `rules` held at that moment). */
  async function runRecording(plan: PlanItem[], api: ReturnType<typeof fakeApi>) {
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    const writes: Array<{ marker: Marker; ruleIds: string[]; rules: Rule[] }> = []
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      const m = items.executeFinalize as Marker | undefined
      if (m) {
        const stored = ((await chrome.storage.local.get('rules')).rules ?? []) as Rule[]
        writes.push({
          marker: JSON.parse(JSON.stringify(m)),
          ruleIds: stored.map((r) => r.id),
          rules: JSON.parse(JSON.stringify(stored)),
        })
      }
      return original(items)
    })
    try {
      await startExecute(plan, tree(), api as unknown as OutlookApi)
    } finally {
      sessionSet.mockImplementation(original)
    }
    return writes
  }

  /** Put the finished batch back to "the worker died right after the learning write". */
  async function dieAfterLearningWrite(intent: Marker['learningIntent']): Promise<void> {
    const state = (await chrome.storage.session.get('executeState')).executeState as ExecuteState
    await chrome.storage.session.set({
      executeState: { ...state, inProgress: true, finishedAt: undefined, rulesAdded: 0, rulesAddedDetail: undefined },
      executeFinalize: {
        startedAt: state.startedAt,
        done: ['rules', 'ledgers', 'convMemory', 'subjectMemory'],
        learningIntent: intent,
      },
    })
    await clearUndoSnapshot()
  }

  it('normal run: the intent is persisted BEFORE the rule write and dropped once learning is recorded', async () => {
    const writes = await runRecording([clientRow()], fakeApi())
    const withIntent = writes.find((w) => w.marker.learningIntent?.confirmed)
    const plannedId = withIntent!.marker.learningIntent!.confirmed!.added[0]!.id
    expect(withIntent!.ruleIds).not.toContain(plannedId) // persisted before the write
    expect((await listRules()).map((r) => r.id)).toEqual([plannedId])
    expect(writes.find((w) => w.marker.done.includes('learning'))?.marker.learningIntent).toBeUndefined()
  })

  it('repro: death after the learned rule was written — the DoneScreen still lists it and 撤回 still removes it', async () => {
    const api = fakeApi()
    const writes = await runRecording([clientRow()], api)
    const intent = writes.find((w) => w.marker.learningIntent?.confirmed)!.marker.learningIntent
    const learnedId = intent!.confirmed!.added[0]!.id
    await dieAfterLearningWrite(intent)

    expect(await recoverStaleExecuteState()).toMatchObject({ recovered: true })
    const state = (await chrome.storage.session.get('executeState')).executeState as ExecuteState
    expect(state.rulesAdded).toBe(1) // was 0
    expect(state.rulesAddedDetail).toEqual([
      expect.objectContaining({ type: 'domain', signal: 'client.example', targetFolderPath: PATH }),
    ])
    expect((await getUndoSnapshot())?.mintedRuleIds).toEqual([learnedId]) // was undefined
    await executeUndo(api as unknown as OutlookApi)
    expect(await listRules()).toEqual([]) // was: the learned rule stayed
  })

  it('death after a REVIVE was written — 撤回 still puts the rule back to sleep', async () => {
    const DISABLED_AT = '2026-10-01T00:00:00.000Z'
    const dormant: Rule = {
      ...newRule({
        type: 'domain',
        signal: 'client.example',
        targetFolderId: A,
        targetFolderPath: PATH,
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
    }
    await setRules([dormant])
    const api = fakeApi()
    const writes = await runRecording([clientRow()], api)
    const intent = writes.find((w) => w.marker.learningIntent?.confirmed)!.marker.learningIntent
    expect(intent!.confirmed!.revived).toHaveLength(1)
    await dieAfterLearningWrite(intent)

    await recoverStaleExecuteState()
    expect((await getUndoSnapshot())?.revivedRules).toEqual([
      expect.objectContaining({ ruleId: 'dormant', autoDisabledReason: 'high-error-rate' }),
    ]) // was undefined
    await executeUndo(api as unknown as OutlookApi)
    expect((await listRules())[0]).toMatchObject({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: DISABLED_AT })
  })

  it('death BEFORE the write landed — the learner simply runs again', async () => {
    const api = fakeApi()
    const writes = await runRecording([clientRow()], api)
    const intent = writes.find((w) => w.marker.learningIntent?.confirmed)!.marker.learningIntent
    await setRules([]) // the write never happened
    await dieAfterLearningWrite(intent)

    await recoverStaleExecuteState()
    const rules = await listRules()
    expect(rules).toHaveLength(1)
    expect(rules[0]!.id).not.toBe(intent!.confirmed!.added[0]!.id)
    expect((await getUndoSnapshot())?.mintedRuleIds).toEqual([rules[0]!.id])
  })

  // H2 consistency (2026-11-06): what the batch switched off is reversed by
  // 撤回 too — and that survives a death between a write and its step marker.
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
    ...over,
  })

  it('death right after the OVERRIDE learner wrote — 撤回 still restores what it superseded and deletes what it minted', async () => {
    await setRules([ruleOf('broad', {})])
    const api = fakeApi()
    const corrected: PlanItem = {
      ...clientRow(),
      emailSubject: '甲案合約審閱意見回覆',
      userTouched: true,
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '案件/乙',
    }
    const writes = await runRecording([corrected], api)
    const intent = writes.find((w) => w.marker.learningIntent?.override)!.marker.learningIntent
    expect(intent!.override!.disabled).toHaveLength(1)
    await dieAfterLearningWrite(intent)

    await recoverStaleExecuteState()
    const minted = (await listRules()).find((r) => r.id !== 'broad')!
    expect((await getUndoSnapshot())?.mintedRuleIds).toEqual([minted.id])
    expect((await getUndoSnapshot())?.disabledRules).toEqual([expect.objectContaining({ ruleId: 'broad' })])
    await executeUndo(api as unknown as OutlookApi)
    expect(await listRules()).toEqual([expect.objectContaining({ id: 'broad', enabled: true })])
  })

  it('death after BOTH learners wrote — a revive the override learner then superseded still undoes to the pre-batch state', async () => {
    const DISABLED_AT = '2026-10-01T00:00:00.000Z'
    await setRules([
      ruleOf('case', {
        type: 'case_code',
        signal: '25A0068B',
        targetFolderId: A,
        targetFolderPath: PATH,
        confidence: 0.9,
        enabled: false,
        autoDisabledReason: 'high-error-rate',
        autoDisabledAt: DISABLED_AT,
      }),
    ])
    const api = fakeApi()
    const plan: PlanItem[] = [
      { ...clientRow(), emailId: 'e1', emailSubject: '25A0068B 開庭通知', conversationId: 'c1' },
      {
        ...clientRow(),
        emailId: 'e2',
        emailSubject: '25A0068B 補件通知',
        conversationId: 'c2',
        targetFolderId: B,
        targetFolderPath: '案件/乙',
        userTouched: true,
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: PATH,
      },
    ]
    const writes = await runRecording(plan, api)
    const intent = writes.find((w) => w.marker.learningIntent?.override)!.marker.learningIntent
    expect(intent!.confirmed!.revived).toHaveLength(1)
    await dieAfterLearningWrite(intent)

    await recoverStaleExecuteState()
    await executeUndo(api as unknown as OutlookApi)
    // Was: the revive no longer matched its post-revive fingerprint, so the
    // replay dropped it and undo switched the rule ON (the override's view).
    expect(await listRules()).toEqual([
      expect.objectContaining({ id: 'case', enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: DISABLED_AT }),
    ])
  })

  it('death right after the one-strike disable was written — 撤回 still switches the case rule back on', async () => {
    await setRules([
      ruleOf('case', { type: 'case_code', signal: '25A0068B', targetFolderId: A, targetFolderPath: PATH, confidence: 0.9 }),
    ])
    const api = fakeApi()
    const refiled: PlanItem = {
      ...clientRow(),
      emailSubject: '25A0068B 補件通知',
      targetFolderId: B,
      targetFolderPath: '案件/乙',
      originalRuleId: 'case',
      userTouched: true,
    }
    const writes = await runRecording([refiled], api)
    // The marker as written right BEFORE the disable, and the library right
    // after the 'rules' step — the worker died in between.
    const before = writes.find((w) => w.marker.ruleDisables)!
    expect(before.marker.done).toEqual([])
    const afterRules = writes.find((w) => w.marker.done.includes('rules'))!
    const state = (await chrome.storage.session.get('executeState')).executeState as ExecuteState
    await setRules(afterRules.rules)
    await chrome.storage.session.set({
      executeState: { ...state, inProgress: true, finishedAt: undefined },
      executeFinalize: { ...before.marker, startedAt: state.startedAt },
    })
    await clearUndoSnapshot()

    await recoverStaleExecuteState()
    expect((await getUndoSnapshot())?.disabledRules).toEqual([expect.objectContaining({ ruleId: 'case' })])
    await executeUndo(api as unknown as OutlookApi)
    expect((await listRules()).find((r) => r.id === 'case')).toMatchObject({ enabled: true, overrideCount: 1 })
  })
})
