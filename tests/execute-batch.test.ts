// Execute-loop regressions (2026-11-05 audit). Drives the REAL startExecute /
// startRetryFailed / executeUndo / recoverStaleExecuteState against a fake
// Outlook API (no network) and the chrome.storage mock.

import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import {
  executeUndo,
  getExecuteState,
  recoverStaleExecuteState,
  requestCancel,
  startExecute,
  startRetryFailed,
} from '@/background/execute'
import { OutlookError, type OutlookApi } from '@/shared/outlook-api'
import { listRules, newRule } from '@/shared/rules'
import {
  getConversationMemory,
  getMetrics,
  getMovedByUsLedger,
  getRecentlyProcessedIds,
  getUndoSnapshot,
  setFolderCache,
  setRules,
  setSettings,
} from '@/shared/storage'
import type { ExecuteState, MailFolderNode, PlanItem, Rule } from '@/shared/types'

const ROOT = 'ROOT-FOLDER-ID-00000000000'
const CASES = 'FOLDER-CASES-0000000000000'
const A = 'FOLDER-A-00000000000000000'
const B = 'FOLDER-B-00000000000000000'
const TRASH = 'WELLKNOWN-DELETED-00000000'
const OLDCASE = 'FOLDER-OLDCASE-00000000000'

function tree(extraCases: MailFolderNode[] = []): MailFolderNode[] {
  return [
    {
      id: CASES,
      displayName: '案件',
      path: '案件',
      parentFolderId: ROOT,
      children: [
        { id: A, displayName: '甲', path: '案件/甲', parentFolderId: CASES, children: [] },
        { id: B, displayName: '乙', path: '案件/乙', parentFolderId: CASES, children: [] },
        ...extraCases,
      ],
    },
  ]
}

function moveRow(emailId: string, targetFolderId: string, over: Partial<PlanItem> = {}): PlanItem {
  const path = targetFolderId === A ? '案件/甲' : targetFolderId === B ? '案件/乙' : undefined
  return {
    emailId,
    emailSubject: `主旨 ${emailId}`,
    emailFrom: 'clerk@client.example',
    action: 'move',
    targetFolderId,
    targetFolderPath: path,
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}

const notFound = (what = 'msg') =>
  new OutlookError(
    404,
    `Outlook REST POST /me/messages/${what}/move → 404 {"error":{"code":"ErrorItemNotFound"}}`,
  )

/** Fake OutlookApi: ids move once; moving a dead id 404s like Outlook. */
function fakeApi() {
  const gone = new Set<string>()
  const api = {
    moveMessage: vi.fn(async (id: string, dest: string) => {
      if (gone.has(id)) throw notFound(id)
      gone.add(id)
      return { Id: `new-${id}`, ParentFolderId: dest }
    }),
    moveToDeletedItems: vi.fn(async (id: string) => {
      if (gone.has(id)) throw notFound(id)
      gone.add(id)
      return { Id: `del-${id}`, ParentFolderId: TRASH }
    }),
    deleteMessage: vi.fn(async () => {}),
    createChildFolder: vi.fn(async (_parent: string, name: string) => ({
      Id: `CREATED-${name}-000000000000000`,
      DisplayName: name,
      ParentFolderId: _parent,
    })),
    createTopFolder: vi.fn(),
    listChildFolders: vi.fn(async () => [] as Array<{ Id: string; DisplayName: string }>),
    listTopFolders: vi.fn(async () => []),
  }
  return api
}
type FakeApi = ReturnType<typeof fakeApi> & Record<string, unknown>
const asApi = (f: object) => f as unknown as OutlookApi

beforeEach(() => {
  // keep-alive + undo use chrome.alarms; tests/setup.ts only mocks storage.
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.alarms = {
    create: vi.fn(async () => {}),
    clear: vi.fn(async () => true),
    get: vi.fn(async () => undefined),
    onAlarm: { addListener: vi.fn() },
  }
})

async function finalState(): Promise<ExecuteState> {
  const s = await getExecuteState()
  if (!s) throw new Error('no executeState')
  return s
}

async function waitIdle(): Promise<ExecuteState> {
  await vi.waitFor(async () => expect((await getExecuteState())?.inProgress).toBe(false))
  return finalState()
}

describe('start guards (item 10)', () => {
  it('two overlapping starts on a 3-item plan send 3 POST /move, not 6', async () => {
    const api = fakeApi()
    const plan = [moveRow('e1', A), moveRow('e2', A), moveRow('e3', B)]
    const first = startExecute(plan, tree(), asApi(api))
    const second = startExecute(plan, tree(), asApi(api))
    await expect(second).rejects.toThrow('已有批次執行中')
    await first
    expect(api.moveMessage).toHaveBeenCalledTimes(3)
    const s = await finalState()
    expect(s.summary).toMatchObject({ moved: 3, skipped: 0, errors: 0 })
  })

  it('overlapping retryFailed: one starts, the other gets ALREADY_RUNNING', async () => {
    const api = fakeApi()
    api.moveMessage.mockRejectedValueOnce(new OutlookError(403, 'denied'))
    await startExecute([moveRow('e1', A), moveRow('e2', B)], tree(), asApi(api))
    expect((await finalState()).summary.errors).toBe(1)

    const [r1, r2] = await Promise.all([
      startRetryFailed(asApi(api), async () => tree()),
      startRetryFailed(asApi(api), async () => tree()),
    ])
    expect(r1).toEqual({ ok: true, started: 1 })
    expect(r2).toMatchObject({ ok: false, code: 'ALREADY_RUNNING' })
    const s = await waitIdle()
    expect(s.total).toBe(1)
    expect(s.summary.moved).toBe(1)
  })

  it('a double 撤回 replays the snapshot once', async () => {
    const api = fakeApi()
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    const u1 = executeUndo(asApi(api))
    const u2 = executeUndo(asApi(api))
    await expect(u2).rejects.toThrow('撤回進行中')
    expect(await u1).toMatchObject({ attempted: 1, restored: 1 })
    expect(api.moveMessage.mock.calls.filter((c) => c[1] === 'inbox')).toHaveLength(1)
  })
})

describe('undo snapshot exists before the DoneScreen can see inProgress=false (item 9)', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())

  it('the first executeState write with inProgress=false already has the snapshot', async () => {
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    const seen: Array<Promise<Record<string, unknown>>> = []
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      const st = items.executeState as ExecuteState | undefined
      // Snapshot what a polling DoneScreen would read at this very moment.
      if (st && st.inProgress === false) seen.push(chrome.storage.local.get('undoSnapshot'))
      return original(items)
    })
    restore = () => sessionSet.mockImplementation(original)

    await startExecute([moveRow('e1', A)], tree(), asApi(fakeApi()))
    expect(seen.length).toBeGreaterThan(0)
    expect((await seen[0]!).undoSnapshot).toBeTruthy()
  })
})

describe('delete = move into Deleted Items, undoable (item 17)', () => {
  it('uses moveToDeletedItems (never DELETE) and 撤回 moves it back to the inbox', async () => {
    const api = fakeApi()
    await startExecute(
      [moveRow('e1', A), { ...moveRow('e2', A), action: 'delete', targetFolderId: undefined }],
      tree(),
      asApi(api),
    )
    expect(api.deleteMessage).not.toHaveBeenCalled()
    expect(api.moveToDeletedItems).toHaveBeenCalledWith('e2', expect.anything())
    const s = await finalState()
    expect(s.results[1]).toMatchObject({ status: 'deleted', newMessageId: 'del-e2', destinationFolderId: TRASH })
    const snap = await getUndoSnapshot()
    expect(snap?.moves.map((m) => [m.newMessageId, m.deleted ?? false])).toEqual([
      ['new-e1', false],
      ['del-e2', true],
    ])
    const undo = await executeUndo(asApi(api))
    expect(undo).toMatchObject({ attempted: 2, restored: 2, failed: 0 })
    expect(api.moveMessage).toHaveBeenCalledWith('del-e2', 'inbox')
  })
})

describe('rule overrides = re-target to a DIFFERENT folder only (item 8)', () => {
  function caseRule(id: string, code: string): Rule {
    return {
      ...newRule({
        type: 'case_code',
        signal: code,
        targetFolderId: A,
        targetFolderPath: '案件/甲',
        confidence: 0.9,
        source: 'ai_confirmed',
      }),
      id,
    }
  }

  it('deleting a rule-routed mail neither counts as an override nor kills the case rule', async () => {
    await setRules([caseRule('R-del', '25A0067A')])
    await startExecute(
      [
        {
          ...moveRow('e1', A),
          action: 'delete',
          source: 'ai',
          originalRuleId: 'R-del',
          userTouched: true,
        },
      ],
      tree(),
      asApi(fakeApi()),
    )
    const [r] = await listRules()
    expect(r!.enabled).toBe(true)
    expect(r!.overrideCount ?? 0).toBe(0)
  })

  it('a re-target elsewhere disables the structural rule with reason "override" (revivable)', async () => {
    await setRules([caseRule('R-move', '25A0068B')])
    await startExecute(
      [{ ...moveRow('e1', B), source: 'ai', originalRuleId: 'R-move', userTouched: true }],
      tree(),
      asApi(fakeApi()),
    )
    const [r] = await listRules()
    expect(r!.enabled).toBe(false)
    expect(r!.autoDisabledReason).toBe('override')
    expect(typeof r!.autoDisabledAt).toBe('string')
    expect(r!.overrideCount).toBe(1)
  })

  it('an edited row that still lands in the rule target is not an override', async () => {
    await setRules([caseRule('R-same', '25A0069C')])
    await startExecute(
      [{ ...moveRow('e1', A), source: 'ai', originalRuleId: 'R-same', userTouched: true }],
      tree(),
      asApi(fakeApi()),
    )
    const [r] = await listRules()
    expect(r!.enabled).toBe(true)
    expect(r!.overrideCount ?? 0).toBe(0)
  })
})

describe('new_folder 409 / lost-response recovery (item 11)', () => {
  const newFolderRow = (name: string): PlanItem => ({
    ...moveRow('e1', A),
    action: 'new_folder',
    targetFolderId: undefined,
    targetFolderPath: undefined,
    suggestedParentPath: '案件',
    suggestedFolderName: name,
  })
  const ACME = 'FOLDER-ACME-0000000000000'

  it('409 recovery matches the existing folder case-insensitively', async () => {
    const api = fakeApi()
    api.createChildFolder.mockRejectedValueOnce(new OutlookError(409, 'ErrorFolderExists'))
    api.listChildFolders.mockResolvedValueOnce([{ Id: ACME, DisplayName: 'ACME Corp' }])
    await startExecute([newFolderRow('Acme Corp')], tree(), asApi(api))
    const s = await finalState()
    expect(s.results[0]).toMatchObject({
      status: 'moved',
      destinationFolderId: ACME,
      destinationFolderPath: '案件/ACME Corp',
    })
    expect(api.moveMessage).toHaveBeenCalledWith('e1', ACME, expect.anything())
  })

  it('an existing folder in the tree is reused case-insensitively (no doomed POST)', async () => {
    const api = fakeApi()
    const existing: MailFolderNode = {
      id: ACME,
      displayName: 'ACME Corp',
      path: '案件/ACME Corp',
      parentFolderId: CASES,
      children: [],
    }
    await startExecute([newFolderRow('acme corp ')], tree([existing]), asApi(api))
    expect(api.createChildFolder).not.toHaveBeenCalled()
    expect((await finalState()).results[0]).toMatchObject({ status: 'moved', destinationFolderId: ACME })
  })

  it('a lost create response re-lists siblings and uses the folder it finds', async () => {
    const api = fakeApi()
    api.createChildFolder.mockRejectedValueOnce(new OutlookError(0, 'Network error on POST', undefined, true))
    api.listChildFolders.mockResolvedValueOnce([{ Id: ACME, DisplayName: 'Acme Corp' }])
    await startExecute([newFolderRow('Acme Corp')], tree(), asApi(api))
    expect((await finalState()).results[0]).toMatchObject({ status: 'folder_created', destinationFolderId: ACME })
  })

  it('a lost create with no folder found is an honest error — never "may have moved"', async () => {
    const api = fakeApi()
    api.createChildFolder.mockRejectedValueOnce(new OutlookError(0, 'Network error on POST', undefined, true))
    await startExecute([newFolderRow('Acme Corp')], tree(), asApi(api))
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('沒有移動')
    expect(row.message).not.toContain('可能已成功移動')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })
})

describe('404 on move verifies the destination (item 12)', () => {
  it('a purged destination is an error — the mail is NOT recorded as processed', async () => {
    const api = { ...fakeApi(), getFolder: vi.fn(async () => Promise.reject(notFound('folder'))) } as FakeApi
    api.moveMessage.mockRejectedValueOnce(notFound('e1'))
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('目標資料夾已不存在')
    expect((await getRecentlyProcessedIds()).has('e1')).toBe(false)
  })

  it('destination alive → the usual already-moved soft skip', async () => {
    const api = { ...fakeApi(), getFolder: vi.fn(async () => ({ Id: A, DisplayName: '甲' })) } as FakeApi
    api.moveMessage.mockRejectedValueOnce(notFound('e1'))
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    expect((await finalState()).results[0]).toMatchObject({ status: 'skipped', alreadyGone: true })
    expect((await getRecentlyProcessedIds()).has('e1')).toBe(true)
  })
})

describe('moves into Deleted Items & co. are refused (item 2)', () => {
  function oldUnfilteredTree(): MailFolderNode[] {
    return [
      ...tree(),
      {
        id: TRASH,
        displayName: '刪除的郵件',
        path: '刪除的郵件',
        parentFolderId: ROOT,
        children: [
          { id: OLDCASE, displayName: '甲', path: '刪除的郵件/甲', parentFolderId: TRASH, children: [] },
        ],
      },
    ]
  }

  it('a target inside the trash subtree of an old cached tree is refused (ids resolved via API)', async () => {
    const api = {
      ...fakeApi(),
      getWellKnownFolderIds: vi.fn(async () => new Map([['deleteditems', TRASH]])),
    } as FakeApi
    await startExecute([moveRow('e1', OLDCASE, { targetFolderPath: '刪除的郵件/甲' })], oldUnfilteredTree(), asApi(api))
    expect((await finalState()).results[0]!.status).toBe('error')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('uses the excluded ids persisted with the folder cache', async () => {
    await setFolderCache({ updatedAt: new Date().toISOString(), tree: tree(), excludedFolderIds: [TRASH] })
    const api = fakeApi()
    await startExecute([moveRow('e1', OLDCASE, { targetFolderPath: '刪除的郵件/甲' })], oldUnfilteredTree(), asApi(api))
    expect((await finalState()).results[0]!.status).toBe('error')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('a target missing from the tree whose live parent is the trash is refused', async () => {
    await setFolderCache({ updatedAt: new Date().toISOString(), tree: tree(), excludedFolderIds: [TRASH] })
    const api = {
      ...fakeApi(),
      getFolder: vi.fn(async (id: string) => ({ Id: id, DisplayName: '甲', ParentFolderId: TRASH })),
    } as FakeApi
    await startExecute([moveRow('e1', OLDCASE)], tree(), asApi(api))
    expect((await finalState()).results[0]!.status).toBe('error')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })
})

describe('cancel (items 14, 18)', () => {
  it('a POST aborted mid-flight is a soft "may have completed" row, the rest cancelled', async () => {
    const api = fakeApi()
    api.moveMessage.mockImplementationOnce(async () => {
      void requestCancel()
      throw new OutlookError(0, 'Cancelled while POST was in flight', undefined, true, true)
    })
    await startExecute([moveRow('e1', A), moveRow('e2', B), moveRow('e3', B)], tree(), asApi(api))
    const s = await finalState()
    expect(s.results[0]!.status).toBe('skipped')
    expect(s.results[0]!.message).toContain('已取消')
    expect(s.results.slice(1).map((r) => r.status)).toEqual(['cancelled', 'cancelled'])
    expect(s.cancelRequested).toBe(true)
  })

  it('requestCancel takes effect from memory at the next item boundary', async () => {
    const api = fakeApi()
    let calls = 0
    api.moveMessage.mockImplementation(async (id: string) => {
      calls++
      if (calls === 2) void requestCancel()
      return { Id: `new-${id}`, ParentFolderId: A }
    })
    const plan = Array.from({ length: 6 }, (_, i) => moveRow(`e${i}`, A))
    await startExecute(plan, tree(), asApi(api))
    const s = await finalState()
    expect(api.moveMessage).toHaveBeenCalledTimes(2)
    expect(s.summary).toMatchObject({ moved: 2, cancelled: 4 })
    expect(s.cancelRequested).toBe(true)
  })
})

describe('state persistence is throttled (item 18)', () => {
  it('a 30-item batch does not write the whole state twice per item nor read it per item', async () => {
    const set = chrome.storage.session.set as unknown as Mock
    const get = chrome.storage.session.get as unknown as Mock
    set.mockClear()
    get.mockClear()
    const plan = Array.from({ length: 30 }, (_, i) => moveRow(`e${i}`, i % 2 ? A : B))
    await startExecute(plan, tree(), asApi(fakeApi()))
    const stateWrites = set.mock.calls.filter((c) => 'executeState' in (c[0] as object)).length
    const stateReads = get.mock.calls.filter((c) => {
      const k = c[0]
      return k === 'executeState' || (Array.isArray(k) && k.includes('executeState'))
    }).length
    expect(stateWrites).toBeLessThan(10) // was 61 (1 + 2 per item)
    expect(stateReads).toBeLessThanOrEqual(1) // was 31 (1 + 1 per item)
    expect((await finalState()).summary.moved).toBe(30)
  })
})

describe('moved-by-us ledger only while the harvester is enabled (item 5)', () => {
  it('harvest off → nothing written', async () => {
    await startExecute([moveRow('e1', A, { conversationId: 'c1' })], tree(), asApi(fakeApi()))
    expect(await getMovedByUsLedger()).toEqual([])
  })

  it('harvest on → destination-side id recorded, slim', async () => {
    await setSettings({ harvestEnabled: true })
    await startExecute([moveRow('e1', A, { conversationId: 'c1' })], tree(), asApi(fakeApi()))
    expect(await getMovedByUsLedger()).toEqual([
      { newMessageId: 'new-e1', destinationFolderId: A, conversationId: 'c1', movedAt: expect.any(Number) },
    ])
  })
})

describe('SW-death recovery replays the post-loop bookkeeping (item 7)', () => {
  async function seedDeadBatch(opts: { marker?: string[] } = {}) {
    await setRules([
      {
        ...newRule({
          type: 'domain',
          signal: 'client.example',
          targetFolderId: A,
          targetFolderPath: '案件/甲',
          confidence: 0.7,
          source: 'ai_confirmed',
        }),
        id: 'R1',
      },
    ])
    await setSettings({ harvestEnabled: true })
    const plan: PlanItem[] = [
      moveRow('e1', A, { source: 'rule', ruleId: 'R1', conversationId: 'c1', emailSubject: '112訴204 開庭通知' }),
      moveRow('e2', B),
      moveRow('e3', B),
    ]
    const startedAt = Date.now() - 5_000
    const state: ExecuteState = {
      inProgress: true,
      cancelRequested: false,
      startedAt,
      total: 3,
      current: 1,
      plan,
      results: [
        {
          emailId: 'e1',
          subject: plan[0]!.emailSubject,
          action: 'move',
          status: 'moved',
          newMessageId: 'new-e1',
          destinationFolderId: A,
          destinationFolderPath: '案件/甲',
        },
        { emailId: 'e2', subject: 's', action: 'move', status: 'processing' },
        { emailId: 'e3', subject: 's', action: 'move', status: 'queued' },
      ],
      summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    await chrome.storage.session.set({ executeState: state })
    if (opts.marker) {
      await chrome.storage.session.set({ executeFinalize: { startedAt, done: opts.marker } })
    }
  }

  it('hits, thread memory, ledgers, metrics and undo all survive an idle-kill', async () => {
    await seedDeadBatch()
    const r = await recoverStaleExecuteState()
    expect(r).toEqual({ recovered: true, stuck: 2 })
    const [rule] = await listRules()
    expect(rule!.matchCount).toBe(1)
    expect((await getConversationMemory()).c1?.folderId).toBe(A)
    expect((await getRecentlyProcessedIds()).has('e1')).toBe(true)
    expect((await getMovedByUsLedger()).map((e) => e.newMessageId)).toEqual(['new-e1'])
    expect((await getUndoSnapshot())?.moves.map((m) => m.newMessageId)).toEqual(['new-e1'])
    expect((await getMetrics()).moved).toBe(1)
    const s = await finalState()
    expect(s.inProgress).toBe(false)
    expect(s.results.map((x) => x.status)).toEqual(['moved', 'error', 'cancelled'])
  })

  it('a death DURING bookkeeping replays only the unfinished steps', async () => {
    await seedDeadBatch({ marker: ['rules'] })
    await recoverStaleExecuteState()
    const [rule] = await listRules()
    expect(rule!.matchCount).toBe(0) // 'rules' already ran before the death
    expect((await getConversationMemory()).c1?.folderId).toBe(A) // the rest did not
  })

  it('concurrent recovery calls (module load + onStartup) replay once', async () => {
    await seedDeadBatch()
    await Promise.all([recoverStaleExecuteState(), recoverStaleExecuteState()])
    const [rule] = await listRules()
    expect(rule!.matchCount).toBe(1)
    expect((await getMetrics()).moved).toBe(1)
  })
})
