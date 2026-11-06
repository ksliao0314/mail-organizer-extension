// Soft thread hints name only folders the AI is offered (2026-11-06, round-3
// review item 1).
//
// Round 2 filtered the exemplars, the case map and the recent-folders hint
// down to the AI-visible folders, but not the soft thread hints (B2-C): they
// resolved the remembered folder against the FULL tree and fell back to the
// stored folderPath, so a hint could steer the AI to 垃圾郵件 / 寄件備份
// (hidden from it — valid only for rules, the picker and execute), to an
// excluded folder, or to a deleted folder's old path.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import { inboxMail, installFetchRouter, markBodiesFetched, seedClassifyEnv, waitFor, type FetchRouter } from './helpers/classify-env'

const tree = (): MailFolderNode[] => [
  {
    id: 'CASES-0000000000000000000',
    displayName: '案件',
    path: '案件',
    children: [
      { id: 'JIA-00000000000000000000', displayName: '甲', path: '案件/甲', children: [] },
      { id: 'YI-000000000000000000000', displayName: '乙', path: '案件/乙', children: [] },
    ],
  },
  {
    id: 'JUNK-0000000000000000000',
    displayName: '垃圾郵件',
    path: '垃圾郵件',
    wellKnown: 'junkemail',
    children: [],
  },
  {
    id: 'SENT-0000000000000000000',
    displayName: '寄件備份',
    path: '寄件備份',
    wellKnown: 'sentitems',
    children: [],
  },
]

/** An ambiguous (conflictCount > 0) memory entry: a hint, never an auto-route. */
const ambiguous = (folderId: string, folderPath: string) => ({
  folderId,
  folderPath,
  lastFiledAt: '2026-11-01T00:00:00.000Z',
  timesFiled: 3,
  conflictCount: 1,
})

let sw: SwHarness
let router: FetchRouter
beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await seedClassifyEnv()
  router = installFetchRouter()
})
afterEach(() => vi.unstubAllGlobals())

describe('soft thread hints', () => {
  it('only name folders the AI is offered: not Junk / Sent, not excluded, not a deleted folder', async () => {
    const settings = (await chrome.storage.local.get('settings')).settings as Record<string, unknown>
    await chrome.storage.local.set({
      settings: { ...settings, excludeFolderPrefixes: ['案件/乙'] },
      folderCache: { updatedAt: new Date().toISOString(), tree: tree(), wellKnown: { roots: [], complete: true } },
      conversationMemory: {
        'conv-junk': ambiguous('JUNK-0000000000000000000', '垃圾郵件'),
        'conv-sent': ambiguous('SENT-0000000000000000000', '寄件備份'),
        'conv-excluded': ambiguous('YI-000000000000000000000', '案件/乙'),
        // Deleted since: its id is gone from the tree, only the stored path is left.
        'conv-deleted': { ...ambiguous('GONE-0000000000000000000', '舊案/已刪除'), conflictCount: 0 },
        'conv-ok': ambiguous('JIA-00000000000000000000', '案件/甲'),
      },
    })
    const convs = ['conv-junk', 'conv-sent', 'conv-excluded', 'conv-deleted', 'conv-ok']
    router.inboxPages = [{ messages: convs.map((c, i) => inboxMail(i, { ConversationId: c })) }]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: convs.length })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    // None of them was auto-routed: all five go to the AI.
    expect((r.data as { unmatchedPreview: unknown[] }).unmatchedPreview).toHaveLength(5)

    const cache = (await chrome.storage.session.get('preflightCache')).preflightCache as {
      threadHints: Record<string, string>
    }
    expect(Object.keys(cache.threadHints)).toEqual(['m4'])
    expect(cache.threadHints.m4).toContain('「案件/甲」')

    // End to end: nothing the AI is not offered reaches the request.
    await markBodiesFetched()
    await sw.send({ type: 'classifyAi', runId: (r.data as { runId: string }).runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === 1)
    const prompt = JSON.stringify(router.claudeBodies[0])
    for (const hidden of ['垃圾郵件', '寄件備份', '舊案/已刪除', '線索:此對話近期在「案件/乙」']) {
      expect(prompt, hidden).not.toContain(hidden)
    }
    expect(prompt).toContain('線索:此對話近期在「案件/甲」與其他資料夾間擺盪')
  })

  it('a renamed folder is still hinted under its current path (resolved by id)', async () => {
    await chrome.storage.local.set({
      folderCache: { updatedAt: new Date().toISOString(), tree: tree(), wellKnown: { roots: [], complete: true } },
      subjectMemory: {
        // Stored under the folder's old name; the id still resolves.
        '每月工時報表彙整通知': { ...ambiguous('JIA-00000000000000000000', '案件/甲（舊名）'), harvested: true, conflictCount: 0 },
      },
    })
    router.inboxPages = [{ messages: [inboxMail(0, { Subject: '每月工時報表彙整通知' })] }]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 1 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    const cache = (await chrome.storage.session.get('preflightCache')).preflightCache as {
      threadHints: Record<string, string>
    }
    expect(cache.threadHints.m0).toBe('依你既有的 Outlook 歸檔，同主旨郵件近期歸於「案件/甲」（僅供參考）')
  })
})
