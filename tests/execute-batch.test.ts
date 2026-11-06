// Execute-loop regressions (2026-11-05 audit). Drives the REAL startExecute /
// startRetryFailed / executeUndo / recoverStaleExecuteState against a fake
// Outlook API (no network) and the chrome.storage mock.

import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import {
  ExecuteStartError,
  executeUndo,
  getExecuteState,
  launchExecute,
  recoverStaleExecuteState,
  requestCancel,
  startExecute,
  startRetryFailed,
} from '@/background/execute'
import { isKeepAliveHeartbeatActive } from '@/background/keep-alive'
import { EXECUTE_PARALLEL_ITEMS } from '@/shared/constants'
import {
  OutlookError,
  OutlookTokenUnavailableError,
  flattenFolderTree,
  type OutlookApi,
} from '@/shared/outlook-api'
import { listRules, newRule } from '@/shared/rules'
import {
  getBatchStats,
  getConversationMemory,
  getFolderActivity,
  getFolderCache,
  getMetrics,
  getMovedByUsLedger,
  getRecentlyProcessedIds,
  getSubjectMemory,
  getUndoSnapshot,
  setFolderCache,
  setRules,
  setSettings,
} from '@/shared/storage'
import type { ExecuteState, FolderCache, MailFolderNode, PlanItem, Rule } from '@/shared/types'

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
      getWellKnownFolderId: vi.fn(async (role: string) => (role === 'deleteditems' ? TRASH : undefined)),
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

// C2 (2026-11-06): v1.0.8 refused Junk Email and Sent Items (+ subfolders)
// too, and said 「可能已被刪除」 for every refusal. O1 (2026-11-06 r2): the
// tree leaves the trash / Drafts / Outbox out; their roots come with the
// folder cache (FolderCache.wellKnown).
describe('non-filing destinations (C2 / O1)', () => {
  const JUNK = 'WELLKNOWN-JUNK-00000000000'
  const JUNK_SUB = 'FOLDER-JUNK-SUB-0000000000'
  const SENT = 'WELLKNOWN-SENT-00000000000'
  const SENT_SUB = 'FOLDER-SENT-CLIENT-A-00000'
  const DRAFTS = 'WELLKNOWN-DRAFTS-000000000'
  const OUTBOX = 'WELLKNOWN-OUTBOX-000000000'
  /** The tree as fetched now: Junk / Sent in (roots tagged), the trash / Drafts / Outbox out. */
  function newTree(): MailFolderNode[] {
    return [
      ...tree(),
      {
        id: JUNK,
        displayName: '垃圾郵件',
        path: '垃圾郵件',
        parentFolderId: ROOT,
        wellKnown: 'junkemail',
        children: [{ id: JUNK_SUB, displayName: '可疑', path: '垃圾郵件/可疑', parentFolderId: JUNK, children: [] }],
      },
      {
        id: SENT,
        displayName: '寄件備份',
        path: '寄件備份',
        parentFolderId: ROOT,
        wellKnown: 'sentitems',
        children: [{ id: SENT_SUB, displayName: '客戶A', path: '寄件備份/客戶A', parentFolderId: SENT, children: [] }],
      },
    ]
  }
  const roots = (): FolderCache['wellKnown'] => ({
    complete: true,
    roots: [
      { id: TRASH, role: 'deleteditems', displayName: '刪除的郵件' },
      { id: DRAFTS, role: 'drafts' },
      { id: OUTBOX, role: 'outbox' },
      { id: JUNK, role: 'junkemail' },
      { id: SENT, role: 'sentitems' },
    ],
  })
  async function seedNewCache() {
    await setFolderCache({ updatedAt: new Date().toISOString(), tree: newTree(), wellKnown: roots() })
  }

  it('Junk / Sent Items and their subfolders are ordinary filing targets', async () => {
    await seedNewCache()
    const api = fakeApi()
    const plan = [JUNK, JUNK_SUB, SENT, SENT_SUB].map((id, i) => moveRow(`e${i}`, id))
    await startExecute(plan, newTree(), asApi(api))
    const s = await finalState()
    expect(s.results.map((r) => r.status)).toEqual(['moved', 'moved', 'moved', 'moved'])
    expect(api.moveMessage).toHaveBeenCalledTimes(4)
  })

  it('a folder deleted into the trash is refused with a message that says so (live parent walk)', async () => {
    await seedNewCache()
    const api = {
      ...fakeApi(),
      getFolder: vi.fn(async (id: string) => ({ Id: id, DisplayName: '甲', ParentFolderId: TRASH })),
    } as FakeApi
    await startExecute([moveRow('e1', OLDCASE, { targetFolderPath: '案件/甲' })], newTree(), asApi(api))
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('已在 Outlook 被刪除')
    expect(row.message).not.toContain('可能已被刪除')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('Drafts is refused as Drafts', async () => {
    await seedNewCache()
    const api = fakeApi()
    await startExecute([moveRow('e1', DRAFTS)], newTree(), asApi(api))
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('「草稿」或「寄件匣」')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('new_folder never creates a folder inside the trash (a tree that still lists it)', async () => {
    await seedNewCache()
    const api = fakeApi()
    const legacyTree: MailFolderNode[] = [
      ...newTree(),
      { id: TRASH, displayName: '刪除的郵件', path: '刪除的郵件', parentFolderId: ROOT, children: [] },
    ]
    await startExecute(
      [
        {
          ...moveRow('e1', A),
          action: 'new_folder',
          targetFolderId: undefined,
          targetFolderPath: undefined,
          suggestedParentPath: '刪除的郵件',
          suggestedFolderName: '新案',
        },
      ],
      legacyTree,
      asApi(api),
    )
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('不能在「已刪除的郵件」裡建立新資料夾')
    expect(api.createChildFolder).not.toHaveBeenCalled()
  })

  it('a v1.0.8 cache (all five roots "excluded") no longer blocks Sent Items once the names resolve', async () => {
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: tree(),
      excludedFolderIds: [TRASH, JUNK, SENT, DRAFTS, OUTBOX],
    })
    const ids: Record<string, string> = { deleteditems: TRASH, drafts: DRAFTS }
    const api = {
      ...fakeApi(),
      getWellKnownFolderId: vi.fn(async (role: string) => ids[role]),
      getFolder: vi.fn(async (id: string) => ({ Id: id, DisplayName: 'x', ParentFolderId: id === SENT_SUB ? SENT : ROOT })),
    } as FakeApi
    await startExecute([moveRow('e1', SENT_SUB), moveRow('e2', DRAFTS)], tree(), asApi(api))
    const s = await finalState()
    expect(s.results.map((r) => r.status)).toEqual(['moved', 'error'])
  })

  it('an incomplete cache looks up just the missing roots, each alone, once per batch', async () => {
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: newTree(),
      wellKnown: { complete: false, roots: [{ id: DRAFTS, role: 'drafts' }] },
    })
    const getWellKnownFolderId = vi.fn(async (role: string) => (role === 'deleteditems' ? TRASH : undefined))
    const api = { ...fakeApi(), getWellKnownFolderId } as FakeApi
    await startExecute([moveRow('e1', TRASH), moveRow('e2', DRAFTS), moveRow('e3', A)], newTree(), asApi(api))
    expect((await finalState()).results.map((r) => r.status)).toEqual(['deleted', 'error', 'moved'])
    // Drafts was known: only the trash and Outbox were asked for.
    expect(getWellKnownFolderId.mock.calls.map((c) => c[0]).sort()).toEqual(['deleteditems', 'outbox'])
  })

  // O1a (2026-11-06 r2): a move into 刪除的郵件 was learned like a filing — an
  // ai_confirmed domain rule → trash, conversation / subject memory → trash
  // (follow-ups auto-routed there at 0.95), folder activity and the AI's
  // recent-folders hint.
  describe('Deleted Items itself as a target runs as a delete (O1a)', () => {
    const trashRow = (id: string, over: Partial<PlanItem> = {}) =>
      moveRow(id, TRASH, { targetFolderPath: '刪除的郵件', conversationId: `conv-${id}`, ...over })

    it('an AI / user move: deleted, nothing learned or remembered, undo returns it to the inbox', async () => {
      await seedNewCache()
      const api = fakeApi()
      await startExecute(
        [trashRow('e1', { source: 'ai', aiOriginalAction: 'move', aiOriginalTargetFolderPath: '刪除的郵件' })],
        newTree(),
        asApi(api),
      )
      const s = await finalState()
      expect(s.results[0]).toMatchObject({ status: 'deleted', newMessageId: 'del-e1' })
      expect(s.summary).toMatchObject({ moved: 0, deleted: 1 })
      expect(api.moveToDeletedItems).toHaveBeenCalledWith('e1', expect.anything())
      expect(api.moveMessage).not.toHaveBeenCalled()
      expect(await listRules()).toEqual([]) // no ai_confirmed `domain client.example → 刪除的郵件`
      expect(await getConversationMemory()).toEqual({})
      expect(await getSubjectMemory()).toEqual({})
      expect(await getFolderActivity()).toEqual([])
      const snap = await getUndoSnapshot()
      expect(snap?.moves).toEqual([expect.objectContaining({ newMessageId: 'del-e1', deleted: true })])
      expect(await executeUndo(asApi(api))).toMatchObject({ restored: 1 })
      expect(api.moveMessage).toHaveBeenCalledWith('del-e1', 'inbox')
    })

    it('a rule-routed one still counts the rule hit', async () => {
      await seedNewCache()
      const trashRule: Rule = {
        ...newRule({
          type: 'sender',
          signal: 'spam@x.example',
          targetFolderId: TRASH,
          targetFolderPath: '刪除的郵件',
          confidence: 0.75,
          source: 'user_manual',
        }),
        id: 'R-trash',
      }
      await setRules([trashRule])
      await startExecute([trashRow('e1', { source: 'rule', ruleId: 'R-trash' })], newTree(), asApi(fakeApi()))
      expect((await finalState()).results[0]!.status).toBe('deleted')
      const [r] = await listRules()
      expect(r!.matchCount).toBe(1)
      expect(await getConversationMemory()).toEqual({})
    })
  })

  // O1f (2026-11-06 r2): the top-level 409 / lost-response recovery adopted a
  // hidden 「草稿」 without any check — and spliced it into the cached tree.
  describe('new_folder never adopts a non-filing folder (O1f)', () => {
    const topLevel = (name: string): PlanItem => ({
      ...moveRow('e1', A),
      action: 'new_folder',
      targetFolderId: undefined,
      targetFolderPath: undefined,
      suggestedParentPath: undefined,
      suggestedFolderName: name,
    })

    it('a 409 against Drafts itself is refused, never moved into or spliced', async () => {
      await seedNewCache()
      const api = fakeApi()
      api.createTopFolder.mockRejectedValueOnce(new OutlookError(409, 'ErrorFolderExists'))
      api.listTopFolders.mockResolvedValueOnce([{ Id: DRAFTS, DisplayName: '草稿' }] as never)
      await startExecute([topLevel('草稿')], newTree(), asApi(api))
      const row = (await finalState()).results[0]!
      expect(row.status).toBe('error')
      expect(row.message).toContain('「草稿」或「寄件匣」')
      expect(api.moveMessage).not.toHaveBeenCalled()
      expect(flattenFolderTree((await getFolderCache())!.tree).some((n) => n.id === DRAFTS)).toBe(false)
    })

    it('a lost create response that finds the trash itself is refused too', async () => {
      await seedNewCache()
      const api = fakeApi()
      api.createTopFolder.mockRejectedValueOnce(new OutlookError(0, 'Network error on POST', undefined, true))
      api.listTopFolders.mockResolvedValueOnce([{ Id: TRASH, DisplayName: '刪除的郵件' }] as never)
      await startExecute([topLevel('刪除的郵件')], newTree(), asApi(api))
      const row = (await finalState()).results[0]!
      expect(row.status).toBe('error')
      expect(row.message).toContain('和「已刪除的郵件」同名')
      expect(api.moveMessage).not.toHaveBeenCalled()
      expect(api.moveToDeletedItems).not.toHaveBeenCalled()
    })

    it('an in-batch twin that is a well-known root (a tree listing Outbox) is refused', async () => {
      await seedNewCache()
      const api = fakeApi()
      const listingOutbox: MailFolderNode[] = [
        ...newTree(),
        { id: OUTBOX, displayName: '寄件匣', path: '寄件匣', parentFolderId: ROOT, children: [] },
      ]
      await startExecute([topLevel('寄件匣')], listingOutbox, asApi(api))
      expect((await finalState()).results[0]!.status).toBe('error')
      expect(api.moveMessage).not.toHaveBeenCalled()
      expect(api.createTopFolder).not.toHaveBeenCalled()
    })
  })
})

// O4 (2026-11-06 r2): with no Outlook token (no / frozen OWA tab) every item
// waited out its own ~10 s token timeout — ≈ 28 minutes of identical error
// rows for 500 items.
describe('token outage stops the batch (O4)', () => {
  const outage = () => new OutlookTokenUnavailableError('Outlook 分頁沒有回應(可能被瀏覽器凍結或正忙)。', 'TAB_UNREACHABLE')

  it('the first token failure stops the batch; every unrun row says why and stays retryable', async () => {
    const api = fakeApi()
    api.moveMessage.mockImplementation(async () => {
      throw outage()
    })
    const plan = Array.from({ length: 12 }, (_, i) => moveRow(`e${i}`, i % 2 ? A : B))
    await startExecute(plan, tree(), asApi(api))
    const s = await finalState()
    // Only the moves already started when it hit — not one per item.
    expect(api.moveMessage.mock.calls.length).toBeLessThanOrEqual(EXECUTE_PARALLEL_ITEMS)
    expect(s.results.every((r) => r.status === 'error')).toBe(true)
    expect(new Set(s.results.map((r) => r.message)).size).toBe(1)
    expect(s.results[0]!.message).toContain('無法取得 Outlook 授權')
    expect(s.results[0]!.message).toContain('重試失敗項目')
    expect(s.summary).toMatchObject({ errors: 12, cancelled: 0 })

    // Once the tab is back, 「重試失敗項目」 takes every row.
    api.moveMessage.mockImplementation(async (id: string, dest: string) => ({ Id: `new-${id}`, ParentFolderId: dest }))
    expect(await startRetryFailed(asApi(api), async () => tree())).toEqual({ ok: true, started: 12 })
    expect((await waitIdle()).summary.moved).toBe(12)
  })

  it('a user cancel still wins for the rows that never ran', async () => {
    const api = fakeApi()
    api.moveMessage.mockImplementation(async () => {
      void requestCancel()
      throw outage()
    })
    await startExecute(Array.from({ length: 8 }, (_, i) => moveRow(`e${i}`, A)), tree(), asApi(api))
    const s = await finalState()
    expect(s.results.some((r) => r.status === 'cancelled')).toBe(true)
  })
})

// O2 (2026-11-06 r2): the start answered ok before the batch's first state
// write; when that write failed the popup polled a null state forever.
describe('launch waits for the initial state write (O2)', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())

  it('a failing first write rejects the launch with a clear error; nothing moved, claim released', async () => {
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      if ('executeState' in items) throw new Error('QUOTA_BYTES quota exceeded')
      return original(items)
    })
    restore = () => sessionSet.mockImplementation(original)
    const api = fakeApi()
    await expect(launchExecute([moveRow('e1', A)], tree(), asApi(api))).rejects.toBeInstanceOf(ExecuteStartError)
    expect(api.moveMessage).not.toHaveBeenCalled()
    expect(await getExecuteState()).toBeNull()

    restore()
    restore = null
    const { done } = await launchExecute([moveRow('e1', A)], tree(), asApi(api))
    // Resolved once the state is durable — before the batch is done.
    expect((await getExecuteState())?.total).toBe(1)
    await done
    expect((await finalState()).summary.moved).toBe(1)
  })

  it('a malformed plan rejects without leaving the executor claimed', async () => {
    await expect(launchExecute([null as unknown as PlanItem], tree(), asApi(fakeApi()))).rejects.toThrow()
    await startExecute([moveRow('e1', A)], tree(), asApi(fakeApi()))
    expect((await finalState()).summary.moved).toBe(1)
  })
})

// O3 (2026-11-06 r2): an awaited storage call ahead of releaseKeepAlive() in
// 撤回's finally — when it rejected, the hold leaked and the 20 s heartbeat
// kept the worker alive for the rest of the session.
describe('撤回 releases its keep-alive hold even when cleanup fails (O3)', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())

  it('clearUndoSnapshot rejecting no longer leaks the hold', async () => {
    const api = fakeApi()
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    const localRemove = chrome.storage.local.remove as unknown as Mock
    const original = localRemove.getMockImplementation()!
    localRemove.mockImplementation(async (keys: string | string[]) => {
      if ([keys].flat().includes('undoSnapshot')) throw new Error('storage unavailable')
      return original(keys)
    })
    restore = () => localRemove.mockImplementation(original)
    await expect(executeUndo(asApi(api))).rejects.toThrow('storage unavailable')
    expect(isKeepAliveHeartbeatActive()).toBe(false)
  })
})

describe('cancel (items 14, 18)', () => {
  const N = EXECUTE_PARALLEL_ITEMS
  /** Resolves once `n` moves are in flight at the same time. */
  function inFlightBarrier(n: number) {
    let started = 0
    let release: () => void = () => {}
    const all = new Promise<void>((r) => (release = r))
    return {
      arrive: async () => {
        if (++started === n) release()
        await all
      },
    }
  }

  it('POSTs aborted mid-flight are soft "may have completed" rows, the never-started rest cancelled', async () => {
    const api = fakeApi()
    const barrier = inFlightBarrier(N)
    api.moveMessage.mockImplementation(async (id: string, _dest: string, signal?: AbortSignal) => {
      await barrier.arrive()
      if (id === 'e1') void requestCancel()
      // Every POST in flight was on the wire when the cancel aborted it.
      if (signal?.aborted) {
        throw new OutlookError(0, 'Cancelled while POST was in flight', undefined, true, true)
      }
      return { Id: `new-${id}`, ParentFolderId: A }
    })
    const plan = Array.from({ length: N + 2 }, (_, i) => moveRow(`e${i + 1}`, i % 2 ? B : A))
    await startExecute(plan, tree(), asApi(api))
    const s = await finalState()
    expect(api.moveMessage).toHaveBeenCalledTimes(N) // nothing started after the cancel
    expect(s.results.map((r) => r.status)).toEqual([
      ...Array.from({ length: N }, () => 'skipped'),
      'cancelled',
      'cancelled',
    ])
    expect(s.results[0]!.message).toContain('已取消')
    expect(s.summary).toMatchObject({ moved: 0, skipped: N, cancelled: 2 })
    expect(s.cancelRequested).toBe(true)
  })

  it("requestCancel takes effect from memory at each worker's next item boundary", async () => {
    const api = fakeApi()
    const barrier = inFlightBarrier(N)
    let cancelled = false
    api.moveMessage.mockImplementation(async (id: string) => {
      await barrier.arrive()
      // The cancel lands while every worker's move is committing.
      if (!cancelled) {
        cancelled = true
        void requestCancel()
      }
      return { Id: `new-${id}`, ParentFolderId: A }
    })
    const plan = Array.from({ length: N + 3 }, (_, i) => moveRow(`e${i}`, A))
    await startExecute(plan, tree(), asApi(api))
    const s = await finalState()
    expect(api.moveMessage).toHaveBeenCalledTimes(N)
    expect(s.summary).toMatchObject({ moved: N, cancelled: 3 })
    expect(s.results.slice(N).every((r) => r.status === 'cancelled')).toBe(true)
    expect(s.cancelRequested).toBe(true)
  })
})

describe('parallel moves (2026-11-05, 放寬並加速)', () => {
  it('runs up to EXECUTE_PARALLEL_ITEMS moves at once, after every new_folder row, rows in plan order', async () => {
    const api = fakeApi()
    const events: string[] = []
    const create = api.createChildFolder.getMockImplementation()!
    api.createChildFolder.mockImplementation(async (parent: string, name: string) => {
      events.push('create')
      return create(parent, name)
    })
    let inFlight = 0
    let peak = 0
    api.moveMessage.mockImplementation(async (id: string, dest: string) => {
      events.push(`move:${id}`)
      peak = Math.max(peak, ++inFlight)
      await new Promise((r) => setTimeout(r, 5))
      inFlight--
      return { Id: `new-${id}`, ParentFolderId: dest }
    })
    const plan: PlanItem[] = Array.from({ length: 10 }, (_, i) => moveRow(`e${i}`, i % 2 ? B : A))
    plan[6] = {
      ...moveRow('e6', A),
      action: 'new_folder',
      targetFolderId: undefined,
      targetFolderPath: undefined,
      suggestedParentPath: '案件',
      suggestedFolderName: '丙',
    }
    await startExecute(plan, tree(), asApi(api))
    const s = await finalState()
    expect(events.slice(0, 2)).toEqual(['create', 'move:e6']) // the folder row first, alone
    expect(peak).toBe(EXECUTE_PARALLEL_ITEMS)
    expect(s.results.map((r) => r.emailId)).toEqual(plan.map((p) => p.emailId))
    expect(s.results.map((r) => r.status)).toEqual(plan.map((p) => (p.action === 'new_folder' ? 'folder_created' : 'moved')))
    expect(s.summary).toMatchObject({ moved: 10, foldersCreated: 1, errors: 0, cancelled: 0 })
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

// C4 (2026-11-06): the whole-state write is throttled, so after an SW death an
// item whose move was already on the wire could still read 'queued' →
// recovery said 「此筆未開始」 and skipped its bookkeeping. Now an in-flight
// record is persisted before every non-idempotent request.
describe('in-flight marker before non-idempotent requests (C4)', () => {
  it('each move goes out only after its index is persisted — while its row may still read queued', async () => {
    // The storage mock keeps object references; snapshot what was WRITTEN.
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    let written: ExecuteState | undefined
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      if (items.executeState) written = structuredClone(items.executeState) as ExecuteState
      return original(items)
    })
    const api = fakeApi()
    const plan = Array.from({ length: 5 }, (_, i) => moveRow(`e${i}`, i % 2 ? A : B))
    const seen: Array<{ marked: boolean; persisted: string | undefined }> = []
    api.moveMessage.mockImplementation(async (id: string, dest: string) => {
      const rec = (await chrome.storage.session.get('executeInFlight')).executeInFlight as
        | { indexes: number[] }
        | undefined
      const idx = plan.findIndex((p) => p.emailId === id)
      seen.push({ marked: !!rec?.indexes.includes(idx), persisted: written?.results[idx]?.status })
      return { Id: `new-${id}`, ParentFolderId: dest }
    })
    try {
      await startExecute(plan, tree(), asApi(api))
    } finally {
      sessionSet.mockImplementation(original)
    }
    expect(seen).toHaveLength(5)
    expect(seen.every((x) => x.marked)).toBe(true)
    expect(seen.some((x) => x.persisted === 'queued')).toBe(true) // the gap the marker covers
    expect((await chrome.storage.session.get('executeInFlight')).executeInFlight).toBeUndefined()
  })

  it('a folder create is marked too', async () => {
    const api = fakeApi()
    let markedAtCreate = false
    api.createChildFolder.mockImplementation(async (parent: string, name: string) => {
      const rec = (await chrome.storage.session.get('executeInFlight')).executeInFlight as
        | { indexes: number[] }
        | undefined
      markedAtCreate = !!rec?.indexes.includes(0)
      return { Id: `CREATED-${name}-000000000000000`, DisplayName: name, ParentFolderId: parent }
    })
    await startExecute(
      [
        {
          ...moveRow('e1', A),
          action: 'new_folder',
          targetFolderId: undefined,
          suggestedParentPath: '案件',
          suggestedFolderName: '丙',
        },
      ],
      tree(),
      asApi(api),
    )
    expect(markedAtCreate).toBe(true)
  })

  it('recovery: a marked row that never reached a final status is "may have completed", kept out of recentlyProcessed', async () => {
    const plan = ['e1', 'e2', 'e3', 'e4', 'e5'].map((id) => moveRow(id, A))
    const startedAt = Date.now() - 5_000
    const row = (i: number, status: ExecuteState['results'][number]['status']) => ({
      emailId: plan[i]!.emailId,
      subject: 's',
      action: 'move' as const,
      status,
    })
    const state: ExecuteState = {
      inProgress: true,
      cancelRequested: false,
      startedAt,
      total: 5,
      current: 2,
      plan,
      results: [
        { ...row(0, 'moved'), newMessageId: 'new-e1', destinationFolderId: A, destinationFolderPath: '案件/甲' },
        row(1, 'queued'), // marked: its move was on the wire
        row(2, 'processing'), // marked
        row(3, 'queued'), // never sent
        row(4, 'processing'), // started, nothing sent yet
      ],
      summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    await chrome.storage.session.set({ executeState: state, executeInFlight: { startedAt, indexes: [0, 1, 2] } })
    expect(await recoverStaleExecuteState()).toEqual({ recovered: true, stuck: 4 })
    const s = await finalState()
    expect(s.results.map((r) => r.status)).toEqual(['moved', 'skipped', 'skipped', 'cancelled', 'error'])
    expect(s.results[1]!.message).toContain('可能已完成,請到 Outlook 確認')
    expect(s.summary).toMatchObject({ moved: 1, skipped: 2, cancelled: 1, errors: 1 })
    const processed = await getRecentlyProcessedIds()
    expect(processed.has('e1')).toBe(true)
    expect(processed.has('e2') || processed.has('e3')).toBe(false)
    expect((await chrome.storage.session.get('executeInFlight')).executeInFlight).toBeUndefined()
  })
})

// C8 (2026-11-06): a death between a counting write and the finalize marker's
// `done` replayed the step — rule hits, metrics, timesFiled counted twice.
describe('finalize counting is exactly-once across a worker death (C8)', () => {
  const domainRule = (matchCount: number): Rule => ({
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: A,
      targetFolderPath: '案件/甲',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    id: 'R1',
    matchCount,
  })
  const memory = (timesFiled: number) => ({
    c1: {
      folderId: A,
      folderPath: '案件/甲',
      lastFiledAt: '2026-11-05T00:00:00.000Z',
      timesFiled,
      conflictCount: 0,
      stableStreak: timesFiled,
    },
  })

  async function seedDeadBatchAfterIntents(opts: { landed: boolean }) {
    await setRules([domainRule(opts.landed ? 6 : 5)])
    await chrome.storage.local.set({
      metrics: { moved: opts.landed ? 11 : 10, deleted: 0, foldersCreated: 0, errors: 0 },
      conversationMemory: memory(opts.landed ? 3 : 2),
    })
    const plan: PlanItem[] = [moveRow('e1', A, { source: 'rule', ruleId: 'R1', conversationId: 'c1' })]
    const startedAt = Date.now() - 5_000
    const state: ExecuteState = {
      inProgress: true,
      cancelRequested: false,
      startedAt,
      total: 1,
      current: 0,
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
      ],
      summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    await chrome.storage.session.set({
      executeState: state,
      // The worker died after recording its intents (and, if `landed`, after
      // the counting writes themselves) but before any step was marked done.
      executeFinalize: {
        startedAt,
        done: [],
        intents: {
          ruleHits: { pre: { R1: 5 }, inc: { R1: 1 } },
          convMemory: { pre: { c1: 2 }, inc: { c1: 1 } },
          metrics: { pre: { moved: 10 }, inc: { moved: 1 } },
        },
      },
    })
  }

  it('writes that already landed are not applied again', async () => {
    await seedDeadBatchAfterIntents({ landed: true })
    await recoverStaleExecuteState()
    expect((await listRules())[0]!.matchCount).toBe(6)
    expect((await getMetrics()).moved).toBe(11)
    expect((await getConversationMemory()).c1!.timesFiled).toBe(3)
  })

  it('writes that never landed are applied exactly once', async () => {
    await seedDeadBatchAfterIntents({ landed: false })
    await recoverStaleExecuteState()
    expect((await listRules())[0]!.matchCount).toBe(6)
    expect((await getMetrics()).moved).toBe(11)
    expect((await getConversationMemory()).c1!.timesFiled).toBe(3)
  })

  it('the normal path records each intent BEFORE its counting write', async () => {
    await setRules([domainRule(0)])
    const localSet = chrome.storage.local.set as unknown as Mock
    const original = localSet.getMockImplementation()!
    const intentsAtWrite: Record<string, string[]> = {}
    localSet.mockImplementation(async (items: Record<string, unknown>) => {
      for (const k of ['rules', 'metrics', 'conversationMemory']) {
        if (k in items && !(k in intentsAtWrite)) {
          const m = (await chrome.storage.session.get('executeFinalize')).executeFinalize as
            | { intents?: Record<string, unknown> }
            | undefined
          intentsAtWrite[k] = Object.keys(m?.intents ?? {})
        }
      }
      return original(items)
    })
    try {
      await startExecute(
        [moveRow('e1', A, { source: 'rule', ruleId: 'R1', conversationId: 'c1' })],
        tree(),
        asApi(fakeApi()),
      )
    } finally {
      localSet.mockImplementation(original)
    }
    expect(intentsAtWrite.rules).toContain('ruleHits')
    expect(intentsAtWrite.conversationMemory).toContain('convMemory')
    expect(intentsAtWrite.metrics).toContain('metrics')
    expect((await listRules())[0]!.matchCount).toBe(1)
    expect((await getMetrics()).moved).toBe(1)
  })

  it('the batch stats row is appended once', async () => {
    await seedDeadBatchAfterIntents({ landed: true })
    const startedAt = ((await chrome.storage.session.get('executeState')).executeState as ExecuteState).startedAt
    const statAt = new Date(startedAt + 1000).toISOString()
    await chrome.storage.local.set({
      batchStats: [{ at: statAt, planned: { rule: 1, thread: 0, ai: 0, unresolved: 0 }, moved: 1, deleted: 0, overrides: { rule: 0, ai: 0, thread: 0 }, gateResolved: 0 }],
    })
    const fin = (await chrome.storage.session.get('executeFinalize')).executeFinalize as Record<string, unknown>
    await chrome.storage.session.set({ executeFinalize: { ...fin, statAt } })
    await recoverStaleExecuteState()
    expect(await getBatchStats()).toHaveLength(1)
  })
})

// Gap 1 (2026-11-06 r2): a folder deleted in OWA AFTER the last tree fetch is
// still in the cached tree (up to 24 h) — execute trusted it, and mail filed
// there sat in the trash until the trash was emptied. An old cached tree's
// folders are now verified live (parent chain up to a known root, memoized
// per batch) before anything is filed into them.
describe('an old cached tree is verified live before filing into it (gap 1)', () => {
  const DRAFTS = 'WELLKNOWN-DRAFTS-000000000'
  const OUTBOX = 'WELLKNOWN-OUTBOX-000000000'
  async function seedCache(ageMs: number) {
    await setFolderCache({
      updatedAt: new Date(Date.now() - ageMs).toISOString(),
      tree: tree(),
      wellKnown: {
        complete: true,
        roots: [
          { id: TRASH, role: 'deleteditems' },
          { id: DRAFTS, role: 'drafts' },
          { id: OUTBOX, role: 'outbox' },
        ],
      },
    })
  }
  /** Live now: 甲 (A) was deleted in OWA — it sits in the trash; 乙 (B) is still under 案件. */
  function liveApi(parents: Record<string, string> = { [A]: TRASH, [B]: CASES, [CASES]: ROOT }) {
    const getFolder = vi.fn(async (id: string) => ({ Id: id, DisplayName: 'x', ParentFolderId: parents[id] }))
    return { ...fakeApi(), getFolder }
  }

  it('a folder trashed since the fetch is refused — one GET per distinct folder — and the cache invalidated', async () => {
    await seedCache(30 * 60_000)
    const api = liveApi()
    await startExecute([moveRow('e1', A), moveRow('e2', A), moveRow('e3', B), moveRow('e4', B)], tree(), asApi(api))
    const s = await finalState()
    expect(s.results.map((r) => r.status)).toEqual(['error', 'error', 'moved', 'moved'])
    expect(s.results[0]!.message).toContain('已在 Outlook 被刪除')
    expect(api.moveMessage).toHaveBeenCalledTimes(2)
    // 甲, 乙 and their shared parent 案件 — each fetched exactly once, even
    // with the parallel workers asking at the same time.
    expect(api.getFolder.mock.calls.map((c) => c[0]).sort()).toEqual([A, B, CASES].sort())
    // The next preflight refetches the tree (and reconcile orphans 甲's rules).
    expect((await getFolderCache())?.invalidatedAt).toBeTruthy()
  })

  it('a just-fetched tree is trusted: no GETs at all', async () => {
    await seedCache(60_000)
    const api = liveApi()
    await startExecute([moveRow('e1', A), moveRow('e2', B)], tree(), asApi(api))
    expect((await finalState()).results.map((r) => r.status)).toEqual(['moved', 'moved'])
    expect(api.getFolder).not.toHaveBeenCalled()
  })

  it('a failed lookup can’t tell — the move goes ahead as before, the cache stays as it is', async () => {
    await seedCache(30 * 60_000)
    const api = {
      ...fakeApi(),
      getFolder: vi.fn(async () => {
        throw new OutlookError(503, 'Service Unavailable')
      }),
    }
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    expect((await finalState()).results[0]!.status).toBe('moved')
    expect((await getFolderCache())?.invalidatedAt).toBeUndefined()
  })

  it('a target gone since the fetch (404 on the first hop) is DESTINATION_GONE, before any move', async () => {
    await seedCache(30 * 60_000)
    const api = {
      ...fakeApi(),
      getFolder: vi.fn(async () => {
        throw notFound('folder')
      }),
    }
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    const row = (await finalState()).results[0]!
    expect(row.status).toBe('error')
    expect(row.message).toContain('目標資料夾已不存在')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('a whole subtree trashed (its parent deleted) is caught too', async () => {
    await seedCache(30 * 60_000)
    const api = liveApi({ [A]: CASES, [CASES]: TRASH })
    await startExecute([moveRow('e1', A)], tree(), asApi(api))
    expect((await finalState()).results[0]!.status).toBe('error')
    expect(api.moveMessage).not.toHaveBeenCalled()
  })

  it('a folder this batch creates is trusted — no GET for it', async () => {
    await seedCache(30 * 60_000)
    const api = liveApi()
    await startExecute(
      [
        {
          ...moveRow('e1', B),
          action: 'new_folder',
          targetFolderId: undefined,
          targetFolderPath: undefined,
          suggestedParentPath: '案件',
          suggestedFolderName: '丙',
        },
        moveRow('e2', 'pending:e1', { targetFolderPath: '案件/丙' }),
      ],
      tree(),
      asApi(api),
    )
    expect((await finalState()).results.map((r) => r.status)).toEqual(['folder_created', 'moved'])
    expect(api.getFolder.mock.calls.map((c) => c[0])).toEqual([CASES]) // the parent, checked once
  })
})

// Gap 2 (2026-11-06 r2): 撤回 had no token circuit breaker — with no Outlook
// token every move it replays waited out its own ~10 s token timeout.
describe('撤回 stops on a token outage (gap 2)', () => {
  const outage = () => new OutlookTokenUnavailableError('請先打開並登入 outlook.office.com 分頁', 'NOT_ON_OWA')
  const undoCalls = (api: ReturnType<typeof fakeApi>) => api.moveMessage.mock.calls.filter((c) => c[1] === 'inbox')

  it('the first token failure stops it; the rest share one message saying 撤回 can’t be retried', async () => {
    const api = fakeApi()
    await startExecute([moveRow('e1', A), moveRow('e2', B), moveRow('e3', A)], tree(), asApi(api))
    api.moveMessage.mockImplementation(async (id: string, dest: string) => {
      if (dest === 'inbox') throw outage()
      return { Id: `new-${id}`, ParentFolderId: dest }
    })
    const r = await executeUndo(asApi(api))
    expect(undoCalls(api)).toHaveLength(1)
    expect(r).toMatchObject({ attempted: 3, restored: 0, failed: 3 })
    expect(new Set(r.errors.map((x) => x.message))).toEqual(new Set([r.message]))
    expect(r.message).toContain('撤回無法重試')
    expect(r.message).toContain('請先打開並登入 outlook.office.com 分頁')
  })

  it('stopping part-way keeps what was already restored', async () => {
    const api = fakeApi()
    await startExecute([moveRow('e1', A), moveRow('e2', B), moveRow('e3', A)], tree(), asApi(api))
    let n = 0
    api.moveMessage.mockImplementation(async (id: string, dest: string) => {
      if (dest === 'inbox' && ++n === 2) throw outage()
      return { Id: `back-${id}`, ParentFolderId: dest }
    })
    const r = await executeUndo(asApi(api))
    expect(undoCalls(api)).toHaveLength(2)
    expect(r).toMatchObject({ attempted: 3, restored: 1, failed: 2 })
  })
})

// ---- Round 3 (2026-11-06 r3) ------------------------------------------------

// Item 3: with Deleted Items unidentifiable both when the tree was fetched and
// at execute, a rule move to the trash ran as an ordinary filing — learned,
// remembered, 近日活動.
describe('a destination that may be the trash itself is filed but not learned from (r3 item 3)', () => {
  const DRAFTS = 'WELLKNOWN-DRAFTS-000000000'
  const OUTBOX = 'WELLKNOWN-OUTBOX-000000000'
  async function seedTrashlessCache() {
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: tree(),
      wellKnown: { complete: false, roots: [{ id: DRAFTS, role: 'drafts' }, { id: OUTBOX, role: 'outbox' }] },
    })
  }

  it('the trash lookup is retried alone; still unknown → moved, but no hit, memory, activity or rule', async () => {
    await seedTrashlessCache()
    const trashRule: Rule = {
      ...newRule({ type: 'sender', signal: 'spam@x.example', targetFolderId: TRASH, targetFolderPath: '刪除的郵件', confidence: 0.75, source: 'user_manual' }),
      id: 'R-trash',
    }
    await setRules([trashRule])
    const getWellKnownFolderId = vi.fn(async (role: string) => {
      if (role === 'deleteditems') throw new OutlookError(503, 'Service Unavailable')
      return undefined
    })
    const api = {
      ...fakeApi(),
      getWellKnownFolderId,
      // The trash is a top-level folder: its live parent is the mailbox root.
      getFolder: vi.fn(async (id: string) => ({ Id: id, DisplayName: 'x', ParentFolderId: ROOT })),
    }
    await startExecute(
      [
        moveRow('e1', TRASH, { targetFolderPath: '刪除的郵件', source: 'rule', ruleId: 'R-trash', conversationId: 'c1' }),
        moveRow('e2', TRASH, {
          targetFolderPath: '刪除的郵件',
          conversationId: 'c2',
          aiOriginalAction: 'move',
          aiOriginalTargetFolderPath: '刪除的郵件',
        }),
      ],
      tree(),
      asApi(api),
    )
    const s = await finalState()
    expect(s.results.map((r) => [r.status, r.unverifiedTarget])).toEqual([
      ['moved', true],
      ['moved', true],
    ])
    expect(getWellKnownFolderId).toHaveBeenCalledWith('deleteditems')
    expect((await listRules()).map((r) => [r.id, r.matchCount])).toEqual([['R-trash', 0]]) // no hit, nothing learned
    expect(await getConversationMemory()).toEqual({})
    expect(await getSubjectMemory()).toEqual({})
    expect(await getFolderActivity()).toEqual([])
    // Still undoable (the mail did move).
    expect((await getUndoSnapshot())?.moves).toHaveLength(2)
  })
})

// Item 4: live verification mapped a token outage to "can't tell" and
// memoized it — the POST then ran and tripped the breaker a full token wait
// late; an item already past runOne's entry check still sent its POST.
describe('a token outage during live verification trips the breaker at once (r3 item 4)', () => {
  async function seedOldCache() {
    await setFolderCache({
      updatedAt: new Date(Date.now() - 30 * 60_000).toISOString(),
      tree: tree(),
      wellKnown: { complete: true, roots: [{ id: TRASH, role: 'deleteditems' }] },
    })
  }
  const outage = () => new OutlookTokenUnavailableError('Outlook 分頁沒有回應', 'TAB_UNREACHABLE')

  it('no POST goes out — not for that folder, not for an item already past the guard', async () => {
    await seedOldCache()
    let releaseA: () => void = () => {}
    const gateA = new Promise<void>((r) => (releaseA = r))
    const getFolder = vi.fn(async (id: string) => {
      if (id === B) throw outage()
      if (id === A) await gateA // still verifying when B trips the breaker
      return { Id: id, DisplayName: 'x', ParentFolderId: id === CASES ? ROOT : CASES }
    })
    const api = { ...fakeApi(), getFolder }
    const run = startExecute([moveRow('e1', A), moveRow('e2', B), moveRow('e3', B)], tree(), asApi(api))
    await vi.waitFor(async () => expect((await getExecuteState())?.results[1]?.status).toBe('error'))
    releaseA()
    await run
    const s = await finalState()
    expect(s.results.map((r) => r.status)).toEqual(['error', 'error', 'error'])
    expect(new Set(s.results.map((r) => r.message)).size).toBe(1)
    expect(s.results[0]!.message).toContain('無法取得 Outlook 授權')
    expect(api.moveMessage).not.toHaveBeenCalled()
    expect(getFolder.mock.calls.filter((c) => c[0] === B)).toHaveLength(1) // memoized rejection
  })
})

// Item 7: a failed start had already dismissed the previous batch's 撤回 and
// (for 重試) wiped its failed rows, while the message said to try again.
describe('a start whose first write fails leaves the previous batch alone (r3 item 7)', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())
  function failStateWrites() {
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      if ('executeState' in items) throw new Error('QUOTA_BYTES quota exceeded')
      return original(items)
    })
    restore = () => sessionSet.mockImplementation(original)
  }

  it('the previous batch’s 撤回 survives a failed start', async () => {
    await startExecute([moveRow('e1', A)], tree(), asApi(fakeApi()))
    expect(await getUndoSnapshot()).toBeTruthy()
    failStateWrites()
    await expect(launchExecute([moveRow('e2', B)], tree(), asApi(fakeApi()))).rejects.toBeInstanceOf(ExecuteStartError)
    expect(await getUndoSnapshot()).toBeTruthy()
  })

  it('a failed 重試 start keeps the failed rows, so 重試 can be pressed again', async () => {
    const api = fakeApi()
    api.moveMessage.mockRejectedValueOnce(new OutlookError(403, 'denied'))
    await startExecute([moveRow('e1', A), moveRow('e2', B)], tree(), asApi(api))
    failStateWrites()
    const r = await startRetryFailed(asApi(api), async () => tree())
    expect(r).toMatchObject({ ok: false, code: 'EXECUTE_START_FAILED' })
    expect(r.ok ? '' : r.message).toContain('請再試一次')
    const prev = await finalState()
    expect(prev.results.map((x) => x.status)).toEqual(['error', 'moved'])
    restore!()
    restore = null
    expect(await startRetryFailed(asApi(api), async () => tree())).toEqual({ ok: true, started: 1 })
    expect((await waitIdle()).summary.moved).toBe(1)
  })
})

// Item 9: auto-trash rule hits are delete rows now (preflight); execute must
// count their hit, and 撤回 must count against the rule — and nothing else
// (H2: a delete taught nothing).
describe('a rule-routed delete (auto-trash rule) (r3 item 9)', () => {
  function trashRule(): Rule {
    return {
      ...newRule({ type: 'sender', signal: 'spam@x.example', targetFolderId: TRASH, targetFolderPath: '刪除的郵件', confidence: 0.75, source: 'user_manual' }),
      id: 'R-trash',
    }
  }
  const deleteRow = (id: string): PlanItem => ({
    ...moveRow(id, A),
    action: 'delete',
    targetFolderId: undefined,
    targetFolderPath: undefined,
    source: 'rule',
    ruleId: 'R-trash',
    originalRuleId: 'R-trash',
    conversationId: `conv-${id}`,
  })

  it('counts the rule hit, and its 撤回 entry carries the rule', async () => {
    await setRules([trashRule()])
    await startExecute([deleteRow('e1')], tree(), asApi(fakeApi()))
    expect((await finalState()).results[0]!.status).toBe('deleted')
    expect((await listRules())[0]!.matchCount).toBe(1)
    expect((await getUndoSnapshot())?.moves).toEqual([
      expect.objectContaining({ newMessageId: 'del-e1', deleted: true, ruleId: 'R-trash' }),
    ])
  })

  it('撤回 counts as an override against the auto-trash rule — and teaches nothing else', async () => {
    await setRules([trashRule()])
    const api = fakeApi()
    await startExecute([deleteRow('e1')], tree(), asApi(api))
    expect(await executeUndo(asApi(api))).toMatchObject({ restored: 1 })
    const [r] = await listRules()
    expect(r!.overrideCount).toBe(1)
    expect(r!.enabled).toBe(true)
    expect(await getConversationMemory()).toEqual({})
  })
})
