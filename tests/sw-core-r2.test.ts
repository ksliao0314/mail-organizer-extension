// SW-core round-2 regressions (2026-11-06 r2), driven through the REAL
// service-worker dispatcher:
//   O2 — startExecute answered ok before the batch's first state write; when
//        that write failed the popup polled a null state forever.
//   O6 — mailbox identity: the "just failed" memo was a single slot (two
//        alternating accounts reset each other's); the answer was stamped
//        with the account peeked AFTER the lookup.
//   O7 — startInitialScan awaited the identity between its running check and
//        the scan's claim: a second overlapping start got ok:true and then
//        threw into the console.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import { getScanState } from '@/background/initial-scan'
import { isKeepAliveHeartbeatActive } from '@/background/keep-alive'
import { getErrorLog } from '@/shared/error-log'
import { getMailboxIdentity, setMailboxIdentity, setRules } from '@/shared/storage'
import { listRules, newRule } from '@/shared/rules'
import type { PlanItem } from '@/shared/types'
import {
  FOLDER_ID,
  FOLDER_PATH,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  seedClassifyEnv,
  waitFor,
} from './helpers/classify-env'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

let sw: SwHarness
beforeAll(async () => {
  sw = await loadServiceWorker()
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

const tokenFor = (accountId?: string) =>
  chrome.storage.session.set({
    cachedToken: { secret: 'owa-token', expiresOn: Date.now() / 1000 + 3600, ...(accountId ? { accountId } : {}) },
  })

describe('O2 — startExecute waits for the initial state write', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())

  it('a failing first write answers EXECUTE_START_FAILED instead of a hollow ok', async () => {
    await seedClassifyEnv()
    const router = installFetchRouter()
    const sessionSet = chrome.storage.session.set as unknown as Mock
    const original = sessionSet.getMockImplementation()!
    sessionSet.mockImplementation(async (items: Record<string, unknown>) => {
      if ('executeState' in items) throw new Error('QUOTA_BYTES quota exceeded')
      return original(items)
    })
    restore = () => sessionSet.mockImplementation(original)
    const plan: PlanItem[] = [
      {
        emailId: 'm1',
        emailSubject: 's',
        emailFrom: 'a@client.example',
        action: 'move',
        targetFolderId: FOLDER_ID,
        targetFolderPath: FOLDER_PATH,
        confidence: 0.9,
        reason: '',
        source: 'ai',
      },
    ]
    const r = await sw.send({ type: 'startExecute', plan })
    expect(r).toMatchObject({ ok: false, code: 'EXECUTE_START_FAILED' })
    expect(r.message).toContain('沒有開始')
    expect(router.outlookUrls.some((u) => u.endsWith('/move'))).toBe(false)
    // The executor was released: the next start runs.
    restore()
    restore = null
    expect((await sw.send({ type: 'startExecute', plan })).ok).toBe(true)
    await waitFor(async () => {
      const s = (await sw.send({ type: 'getExecuteState' })).data as { state: { inProgress: boolean } | null }
      return s.state?.inProgress === false
    })
  })
})

// O3 (2026-11-06 r2): the AI pass awaited its cleanup writes BEFORE
// releaseKeepAlive(); a rejected write leaked the hold, and the 20 s
// heartbeat then kept the worker alive for the whole browser session.
describe('O3 — the AI pass releases its keep-alive hold even when cleanup fails', () => {
  let restore: (() => void) | null = null
  afterEach(() => restore?.())

  it('a rejected classifyStage reset no longer leaks the hold', async () => {
    await seedClassifyEnv()
    const router = installFetchRouter()
    router.inboxPages = [{ messages: [inboxMail(1)] }]
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    await markBodiesFetched()
    const sessionRemove = chrome.storage.session.remove as unknown as Mock
    const original = sessionRemove.getMockImplementation()!
    sessionRemove.mockImplementation(async (keys: string | string[]) => {
      if ([keys].flat().includes('classifyStage')) throw new Error('storage unavailable')
      return original(keys)
    })
    restore = () => sessionRemove.mockImplementation(original)
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const r = await sw.send({ type: 'classifyAi', runId: (pf.data as { runId: string }).runId, excludeIds: [] })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    await waitFor(async () => {
      const p = (await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).data as { done?: boolean } | null
      return p?.done === true
    })
    // The cleanup failure is logged (not rethrown — the finished result stays).
    await waitFor(async () =>
      (await getErrorLog()).some(
        (e) => e.source === 'classify:loop' && (e.context as { phase?: string } | undefined)?.phase === 'cleanup',
      ),
    )
    expect(warn).toBeDefined()
    expect(isKeepAliveHeartbeatActive()).toBe(false)
  })
})

describe('O6 — mailbox identity', () => {
  let meCalls: string[]
  let onMe: () => Promise<Response>
  beforeEach(async () => {
    await seedClassifyEnv()
    const router = installFetchRouter()
    router.inboxPages = [{ messages: [] }]
    const inner = globalThis.fetch as unknown as (url: string, init?: RequestInit) => Promise<Response>
    meCalls = []
    onMe = async () => new Response(JSON.stringify({ error: { code: 'ErrorAccessDenied' } }), { status: 403 })
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init?: RequestInit) => {
        if (/\/me(\?|$)/.test(url)) {
          meCalls.push(url)
          return onMe()
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
  const preflight = async () => expect((await sw.send({ type: 'classifyPreflight', batchSize: 5 })).ok).toBe(true)

  it('the failure memo is per account: alternating accounts no longer reset each other', async () => {
    await tokenFor('acct-A')
    await preflight() // A fails → memo A
    await tokenFor('acct-B')
    await preflight() // B fails → memo B (used to overwrite A's)
    expect(meCalls).toHaveLength(2)
    await tokenFor('acct-A')
    await preflight()
    await tokenFor('acct-B')
    await preflight()
    expect(meCalls).toHaveLength(2) // both still remembered: no retry within the hour
  })

  it('an account change during the lookup: the answer is not stamped onto either account', async () => {
    await tokenFor('acct-A')
    onMe = async () => {
      await tokenFor('acct-B') // OWA switched accounts while /me was in flight
      return new Response(JSON.stringify({ EmailAddress: 'me@firm-a.example' }), { status: 200 })
    }
    await preflight()
    expect(meCalls).toHaveLength(1)
    // Round 1 stored firm-a stamped with acct-B — another mailbox's domain
    // excluded for a week.
    expect(await getMailboxIdentity()).toBeNull()
  })

  it('the account peeked BEFORE the lookup is the one stamped', async () => {
    await tokenFor('acct-A')
    onMe = async () => new Response(JSON.stringify({ EmailAddress: 'me@firm-a.example' }), { status: 200 })
    await preflight()
    expect(await getMailboxIdentity()).toMatchObject({ domain: 'firm-a.example', accountId: 'acct-A' })
  })
})

describe('O7 — overlapping startInitialScan', () => {
  const ROOT = 'ROOT-FOLDER-ID-00000000000'
  const FOLDER_A = 'FOLDER-A-00000000000000000'
  beforeEach(async () => {
    await chrome.storage.local.set({
      settings: { claudeApiKey: 'sk-ant-test', primaryRootPath: '案件', internalDomains: [] },
      folderCache: {
        updatedAt: new Date().toISOString(),
        tree: [
          {
            id: ROOT,
            displayName: '案件',
            path: '案件',
            children: [{ id: FOLDER_A, displayName: 'A', path: '案件/A', children: [] }],
          },
        ],
      },
    })
    await tokenFor()
    sw.tabsQuery.mockResolvedValue([
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
    ])
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        if (/\/me(\?|$)/.test(url)) return new Response(JSON.stringify({ EmailAddress: 'me@firm.example' }), { status: 200 })
        return new Response(JSON.stringify({ value: [] }), { status: 200 })
      }),
    )
  })

  it('exactly one start is accepted; the other gets ALREADY_RUNNING, nothing thrown into the console', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
    const [a, b] = await Promise.all([
      sw.send({ type: 'startInitialScan' }),
      sw.send({ type: 'startInitialScan' }),
    ])
    expect([a.ok, b.ok].sort()).toEqual([false, true])
    expect([a, b].find((r) => !r.ok)!.code).toBe('ALREADY_RUNNING')
    await waitFor(async () => (await getScanState())?.inProgress === false, 5000)
    expect(consoleError.mock.calls.filter((c) => String(c[0]).includes('startInitialScan failed'))).toEqual([])
  })

  it('the scan gets the resolved, account-checked identity (not the raw cache of another account)', async () => {
    await tokenFor('acct-B')
    // Cached for ANOTHER account (A = me@outlook.com) and /me fails now: the
    // scan must not treat A's address as this mailbox's own.
    await setMailboxIdentity({ address: 'me@outlook.com', domain: 'outlook.com', at: new Date().toISOString(), accountId: 'acct-A' })
    const from = (address: string, i: number) => ({
      Id: `m${i}`,
      Subject: `主旨 ${i}`,
      From: { EmailAddress: { Address: address } },
      ToRecipients: [],
    })
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        if (/\/me(\?|$)/.test(url)) return new Response('{"error":{"code":"ErrorAccessDenied"}}', { status: 403 })
        if (url.includes(`/MailFolders/${FOLDER_A}/messages`)) {
          return new Response(
            JSON.stringify({
              value: [from('me@outlook.com', 1), from('me@outlook.com', 2), from('pm@client.example', 3), from('pm@client.example', 4)],
            }),
            { status: 200 },
          )
        }
        return new Response(JSON.stringify({ value: [] }), { status: 200 })
      }),
    )
    const r = await sw.send({ type: 'startInitialScan' })
    expect(r.ok).toBe(true)
    await waitFor(async () => (await getScanState())?.inProgress === false, 5000)
    // For mailbox B, me@outlook.com is just another sender (round 1 read the
    // raw cache and skipped it as "the owner").
    expect((await listRules()).map((x) => `${x.type}:${x.signal}`).sort()).toEqual([
      'domain:client.example',
      'sender:me@outlook.com',
    ])
  })
})
