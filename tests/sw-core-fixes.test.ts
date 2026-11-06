// SW-core regressions (2026-11-06 review), driven through the REAL service
// worker dispatcher:
//   C3  — overlapping startExecute: the loser got ok:true, never ran, and its
//         ids were recorded as executed (never re-offered by the run).
//   C10 — resolveOwnDomains: identity cached 7 days across an account switch;
//         failure memo reset on every SW restart; 3 s budget tripped by queueing.
//   C11 — getFolderTree with no Outlook tab and no cache wrote a 401 to the
//         error log on every Options open.
//   C12 — navigateToFolder from the in-Outlook panel reloaded owaTabs[0], not
//         the tab hosting the panel.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { getErrorLog } from '@/shared/error-log'
import { OutlookApi } from '@/shared/outlook-api'
import { newRule } from '@/shared/rules'
import { getMailboxIdentity, setMailboxIdentity, setRules } from '@/shared/storage'
import type { PlanItem } from '@/shared/types'
import { FOLDER_ID, FOLDER_PATH, installFetchRouter, seedClassifyEnv, waitFor } from './helpers/classify-env'
import { EXTENSION_PAGE_SENDER, loadServiceWorker, type Sender, type SwHarness } from './helpers/sw-harness'

let sw: SwHarness
beforeAll(async () => {
  sw = await loadServiceWorker()
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

const moveItem = (id: string): PlanItem => ({
  emailId: id,
  emailSubject: `subject ${id}`,
  emailFrom: 'a@client.example',
  action: 'move',
  targetFolderId: FOLDER_ID,
  targetFolderPath: FOLDER_PATH,
  confidence: 0.9,
  reason: '',
  source: 'ai',
})

async function waitExecuteIdle() {
  await waitFor(async () => {
    const s = (await sw.send({ type: 'getExecuteState' })).data as { state: { inProgress: boolean } | null }
    return s.state == null || !s.state.inProgress
  })
}

describe('C3 — overlapping startExecute', () => {
  beforeEach(async () => {
    await seedClassifyEnv()
    installFetchRouter()
  })

  it('exactly one start is accepted, and only ITS ids are recorded as executed', async () => {
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    const [a, b] = await Promise.all([
      sw.send({ type: 'startExecute', plan: [moveItem('mA')] }),
      sw.send({ type: 'startExecute', plan: [moveItem('mB')] }),
    ])
    expect([a.ok, b.ok]).toEqual([true, false])
    expect(b.code).toBe('ALREADY_RUNNING')
    await waitExecuteIdle()
    const run = (await chrome.storage.session.get('classifyRun')).classifyRun as { executedIds: string[] }
    expect(run.executedIds).toEqual(['mA'])
  })
})

describe('C10 — resolveOwnDomains', () => {
  let meCalls: string[]
  beforeEach(async () => {
    await seedClassifyEnv()
    const router = installFetchRouter()
    router.inboxPages = [{ messages: [] }]
    const inner = globalThis.fetch as unknown as (url: string, init?: RequestInit) => Promise<Response>
    meCalls = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init?: RequestInit) => {
        if (/\/me(\?|$)/.test(url)) {
          meCalls.push(url)
          return new Response(JSON.stringify({ EmailAddress: 'me@firm-b.example' }), { status: 200 })
        }
        return inner(url, init)
      }),
    )
    await setRules([
      newRule({
        type: 'domain',
        signal: 'x.example',
        targetFolderId: FOLDER_ID,
        targetFolderPath: FOLDER_PATH,
        confidence: 0.7,
        source: 'ai_confirmed',
      }),
    ])
  })
  const tokenFor = (accountId?: string) =>
    chrome.storage.session.set({
      cachedToken: { secret: 'owa-token', expiresOn: Date.now() / 1000 + 3600, ...(accountId ? { accountId } : {}) },
    })

  it('an identity cached for ANOTHER account is refetched and re-keyed', async () => {
    await tokenFor('acct-B')
    await setMailboxIdentity({ address: 'me@firm-a.example', domain: 'firm-a.example', at: new Date().toISOString(), accountId: 'acct-A' })
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(meCalls).toHaveLength(1)
    expect(await getMailboxIdentity()).toMatchObject({ domain: 'firm-b.example', accountId: 'acct-B' })
  })

  it('the same account’s fresh identity is reused without a call', async () => {
    await tokenFor('acct-B')
    await setMailboxIdentity({ address: 'me@firm-b.example', domain: 'firm-b.example', at: new Date().toISOString(), accountId: 'acct-B' })
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(meCalls).toHaveLength(0)
  })

  it('the "just failed" memo survives an SW restart (session storage), per account', async () => {
    await tokenFor('acct-B')
    await chrome.storage.session.set({ mailboxIdentityFailure: { at: Date.now() - 60_000, accountId: 'acct-B' } })
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(meCalls).toHaveLength(0)
    // A different account isn't blocked by it.
    await tokenFor('acct-C')
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(meCalls).toHaveLength(1)
  })

  it('the lookup skips the rate limiter’s queue, so its short budget isn’t spent waiting', async () => {
    const getMyEmail = vi.spyOn(OutlookApi.prototype, 'getMyEmail')
    expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)
    expect(meCalls).toHaveLength(1)
    expect(getMyEmail).toHaveBeenCalledWith(expect.anything(), { unthrottled: true })
  })
})

describe('C11 — getFolderTree with no Outlook tab and no cache', () => {
  it('answers NOT_ON_OWA without writing the error log', async () => {
    sw.tabsQuery.mockResolvedValue([])
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r).toMatchObject({ ok: false, code: 'NOT_ON_OWA' })
    expect(await getErrorLog()).toEqual([])
  })

  it('…and a cache that must be refreshed (v1.0.8 format) is served without logging the missing token', async () => {
    sw.tabsQuery.mockResolvedValue([])
    const tree = [{ id: 'FOLDER-A-000000000000000000', displayName: '甲', path: '甲', children: [] }]
    await chrome.storage.local.set({
      folderCache: { updatedAt: new Date().toISOString(), tree, excludedFolderIds: ['WK-sent'] },
    })
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r).toMatchObject({ ok: true, data: { count: 1 } })
    expect(await getErrorLog()).toEqual([])
    // A REAL failure later in the same streak is still logged once.
    sw.tabsQuery.mockResolvedValue([
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
    ])
    sw.tabsSendMessage.mockResolvedValue({ ok: true, secret: 'tok', expiresOn: Date.now() / 1000 + 3600 })
    vi.stubGlobal('fetch', vi.fn(async () => new Response('forbidden', { status: 403 })))
    await sw.send({ type: 'getFolderTree' })
    await sw.send({ type: 'getFolderTree' })
    expect((await getErrorLog()).filter((e) => e.source === 'outlook:folder-tree')).toHaveLength(1)
  })

  it('a real Outlook failure is still logged', async () => {
    sw.tabsQuery.mockResolvedValue([
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
    ])
    sw.tabsSendMessage.mockResolvedValue({ ok: true, secret: 'tok', expiresOn: Date.now() / 1000 + 3600 })
    vi.stubGlobal('fetch', vi.fn(async () => new Response('forbidden', { status: 403 })))
    const r = await sw.send({ type: 'getFolderTree' })
    expect(r).toMatchObject({ ok: false, code: 'OUTLOOK_403' })
    expect((await getErrorLog()).some((e) => e.source === 'outlook:request')).toBe(true)
  })
})

describe('C12 — navigateToFolder', () => {
  const OTHER: chrome.tabs.Tab = { id: 7, active: false, url: 'https://outlook.office.com/mail/', windowId: 1 } as chrome.tabs.Tab
  const HOST: chrome.tabs.Tab = { id: 9, active: true, url: 'https://outlook.cloud.microsoft/mail/inbox', windowId: 2 } as chrome.tabs.Tab
  const update = () => (chrome.tabs as unknown as { update: ReturnType<typeof vi.fn> }).update

  beforeEach(() => {
    sw.tabsQuery.mockResolvedValue([OTHER, HOST])
    update().mockClear()
  })

  it('from the in-Outlook panel → the tab hosting the panel, on its own host', async () => {
    const panel: Sender = {
      ...EXTENSION_PAGE_SENDER,
      url: 'chrome-extension://test-extension/src/popup/index.html',
      tab: HOST,
      frameId: 3,
    }
    const r = await sw.sendFrom(panel, { type: 'navigateToFolder', folderId: 'AAMk/folder=' })
    expect(r.ok).toBe(true)
    expect(update()).toHaveBeenCalledWith(9, {
      active: true,
      url: `https://outlook.cloud.microsoft/mail/${encodeURIComponent('AAMk/folder=')}`,
    })
  })

  it('from the toolbar popup (no host tab) → an open Outlook tab, the active one first', async () => {
    sw.tabsQuery.mockResolvedValue([OTHER, { ...HOST, active: true }])
    await sw.send({ type: 'navigateToFolder', folderId: 'F1' })
    expect(update()).toHaveBeenCalledWith(9, expect.objectContaining({ url: 'https://outlook.cloud.microsoft/mail/F1' }))
  })

  it('a standalone extension tab (top frame) is never navigated itself', async () => {
    const optionsTab: Sender = {
      ...EXTENSION_PAGE_SENDER,
      tab: { id: 42, url: 'chrome-extension://test-extension/src/options/index.html' } as chrome.tabs.Tab,
      frameId: 0,
    }
    sw.tabsQuery.mockResolvedValue([OTHER])
    await sw.sendFrom(optionsTab, { type: 'navigateToFolder', folderId: 'F1' })
    expect(update()).toHaveBeenCalledWith(7, expect.objectContaining({ url: 'https://outlook.office.com/mail/F1' }))
  })
})
