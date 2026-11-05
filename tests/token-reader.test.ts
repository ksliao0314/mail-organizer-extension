// OWA token acquisition (2026-11-05 audit, item 3).
//
//  - the MSAL reader required only `target.includes('Mail.ReadWrite')`, so a
//    graph.microsoft.com token with a later expiry could win → 401 loop;
//  - "nothing usable" gave no diagnostics (an encrypted cache looked like
//    "not signed in");
//  - the multi-account guard only counted VALID tokens of this origin;
//  - after an extension update the content script is orphaned ("Receiving end
//    does not exist") and the user saw "wait a few seconds" forever;
//  - a token with <60 s left forced a content-script round trip per request;
//  - token failures were invisible (now persisted as lastTokenError).

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { readOutlookTokenFromMsalCache } from '@/shared/msal-token-reader'

const nowSec = () => Math.floor(Date.now() / 1000)

function putToken(key: string, over: Record<string, unknown>) {
  localStorage.setItem(
    key,
    JSON.stringify({
      credentialType: 'AccessToken',
      homeAccountId: 'acct-1',
      secret: 'SECRET-' + key,
      target: 'https://outlook.office.com/Mail.ReadWrite https://outlook.office.com/Mail.Send',
      expiresOn: String(nowSec() + 3600),
      ...over,
    }),
  )
}

describe('readOutlookTokenFromMsalCache', () => {
  beforeEach(() => localStorage.clear())

  it('an Outlook-resource token beats a later-expiring Graph token that also lists Mail.ReadWrite', () => {
    putToken('msal.outlook', { expiresOn: String(nowSec() + 1800) })
    putToken('msal.graph', {
      target: 'https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/User.Read',
      expiresOn: String(nowSec() + 3500),
    })
    putToken('msal.bare-graph-scope', { target: 'Mail.ReadWrite User.Read', expiresOn: String(nowSec() + 3550) })
    const r = readOutlookTokenFromMsalCache()
    expect(r).toMatchObject({ ok: true, secret: 'SECRET-msal.outlook' })
  })

  it('accepts outlook.office365.com scopes too', () => {
    putToken('msal.o365', { target: 'https://outlook.office365.com/Mail.ReadWrite' })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: true, secret: 'SECRET-msal.o365' })
  })

  it('nothing usable → NO_TOKEN with diagnostic counts and no secret', () => {
    putToken('msal.graph', { target: 'https://graph.microsoft.com/Mail.ReadWrite' })
    putToken('msal.expired', { expiresOn: String(nowSec() - 7200) })
    localStorage.setItem('msal.encrypted-blob', 'v2.ZW5jcnlwdGVk')
    localStorage.setItem('unrelated', 'x')
    const r = readOutlookTokenFromMsalCache()
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.code).toBe('NO_TOKEN')
    expect(r.message).toContain('msal 鍵 3 個')
    expect(r.message).toContain('非 JSON 1')
    expect(r.message).toContain('AccessToken 2')
    expect(r.message).toContain('非 Outlook 資源 1')
    expect(r.message).toContain('已過期 1')
    expect(r.message).not.toContain('SECRET')
  })

  it('a cache with only non-JSON msal entries hints at an encrypted cache', () => {
    localStorage.setItem('msal.a', 'opaque-1')
    localStorage.setItem('msal.b', 'opaque-2')
    const r = readOutlookTokenFromMsalCache()
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.message).toContain('加密')
  })

  it('multi-account guard also counts an account whose Outlook token JUST lapsed', () => {
    putToken('msal.a', { homeAccountId: 'acct-A' })
    putToken('msal.b', { homeAccountId: 'acct-B', expiresOn: String(nowSec() - 300) })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: false, code: 'MULTIPLE_ACCOUNTS' })
  })

  it('…but not one that expired long ago (no day-long lockout after switching accounts)', () => {
    putToken('msal.a', { homeAccountId: 'acct-A' })
    putToken('msal.b', { homeAccountId: 'acct-B', expiresOn: String(nowSec() - 2 * 86400) })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: true, accountId: 'acct-A' })
  })

  it('is self-contained: its source runs standalone (what executeScript({ func }) ships)', () => {
    putToken('msal.outlook', {})
    const standalone = new Function(`return (${readOutlookTokenFromMsalCache.toString()})`)() as () => unknown
    expect(standalone()).toEqual(readOutlookTokenFromMsalCache())
  })
})

// ---- token.ts ---------------------------------------------------------------

type Tab = Partial<chrome.tabs.Tab> & { id: number }
let tabs: Tab[]
let sendMessage: ReturnType<typeof vi.fn>
let executeScript: ReturnType<typeof vi.fn>

function installChrome() {
  const c = (globalThis as unknown as { chrome: Record<string, unknown> }).chrome
  sendMessage = vi.fn()
  executeScript = vi.fn(async () => [])
  c.tabs = { query: vi.fn(async () => tabs), sendMessage }
  c.scripting = { executeScript }
  c.runtime = {
    getManifest: () => ({
      content_scripts: [{ js: ['assets/content.ts-loader-abc123.js'] }, { js: ['assets/owa-fab.ts-loader.js'] }],
    }),
  }
}

async function loadToken() {
  vi.resetModules()
  return import('@/background/token')
}

const okToken = (secret: string, ttlSec = 3600, accountId = 'acct-1') => ({
  ok: true,
  secret,
  expiresOn: nowSec() + ttlSec,
  accountId,
})
const orphaned = () =>
  new Error('Could not establish connection. Receiving end does not exist.')

describe('getOwaToken pipeline', () => {
  beforeEach(() => {
    tabs = [{ id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' }]
    installChrome()
  })
  afterEach(() => vi.restoreAllMocks())

  it('persists the last failure and clears it on success', async () => {
    const t = await loadToken()
    tabs = []
    const r = await t.getOwaToken()
    expect(r).toMatchObject({ ok: false, code: 'NOT_ON_OWA' })
    const rec = await t.getLastTokenError()
    expect(rec).toMatchObject({ code: 'NOT_ON_OWA', message: expect.any(String), at: expect.any(Number) })

    tabs = [{ id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' }]
    sendMessage.mockResolvedValue(okToken('fresh'))
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'fresh' })
    expect(await t.getLastTokenError()).toBeNull()
  })

  it('orphaned content script → re-injects the manifest token-reader files and retries', async () => {
    const t = await loadToken()
    sendMessage.mockRejectedValueOnce(orphaned()).mockResolvedValueOnce(okToken('after-reinject'))
    const r = await t.getOwaToken()
    expect(r).toMatchObject({ ok: true, secret: 'after-reinject' })
    expect(executeScript).toHaveBeenCalledWith({
      target: { tabId: 7 },
      files: ['assets/content.ts-loader-abc123.js'],
    })
  })

  it('re-injection that brings no listener back → reads the cache directly in the page', async () => {
    const t = await loadToken()
    sendMessage.mockRejectedValue(orphaned())
    executeScript.mockImplementation(async (opts: { files?: string[]; func?: unknown }) =>
      opts.func ? [{ result: okToken('direct') }] : [],
    )
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'direct' })
  })

  it('a discarded tab is reported as such, without injecting', async () => {
    const t = await loadToken()
    tabs = [{ id: 7, active: true, url: 'https://outlook.office.com/mail/', discarded: true }]
    sendMessage.mockRejectedValue(orphaned())
    expect(await t.getOwaToken()).toMatchObject({ ok: false, code: 'TAB_UNREACHABLE' })
    expect(executeScript).not.toHaveBeenCalled()
  })

  it('a cached token with 30 s left is used directly and refreshed in the background', async () => {
    const t = await loadToken()
    await chrome.storage.session.set({ cachedToken: { secret: 'cached', expiresOn: nowSec() + 30 } })
    sendMessage.mockResolvedValue(okToken('fresh'))
    // Old behaviour: <60 s left = "invalid" → awaited a round trip, returned 'fresh'.
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'cached' })
    await vi.waitFor(async () =>
      expect((await chrome.storage.session.get('cachedToken')).cachedToken).toMatchObject({ secret: 'fresh' }),
    )
    expect(sendMessage).toHaveBeenCalledTimes(1)
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'fresh' })
  })

  it('a failed background refresh keeps the still-usable token', async () => {
    const t = await loadToken()
    await chrome.storage.session.set({ cachedToken: { secret: 'cached', expiresOn: nowSec() + 30 } })
    sendMessage.mockResolvedValue({ ok: false, code: 'NO_TOKEN', message: 'x' })
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'cached' })
    await vi.waitFor(() => expect(sendMessage).toHaveBeenCalledTimes(1))
    await vi.waitFor(async () => expect(await t.getLastTokenError()).not.toBeNull())
    expect((await chrome.storage.session.get('cachedToken')).cachedToken).toMatchObject({ secret: 'cached' })
  })

  it('under 15 s left is NOT usable → foreground refresh', async () => {
    const t = await loadToken()
    await chrome.storage.session.set({ cachedToken: { secret: 'cached', expiresOn: nowSec() + 10 } })
    sendMessage.mockResolvedValue(okToken('fresh'))
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'fresh' })
  })

  it('OWA open on two origins with DIFFERENT accounts → MULTIPLE_ACCOUNTS', async () => {
    const t = await loadToken()
    tabs = [
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' },
      { id: 8, active: false, url: 'https://outlook.office365.com/mail/', status: 'complete' },
    ]
    sendMessage.mockImplementation(async (tabId: number) =>
      okToken(tabId === 7 ? 'a' : 'b', 3600, tabId === 7 ? 'acct-1' : 'acct-2'),
    )
    expect(await t.getOwaToken()).toMatchObject({ ok: false, code: 'MULTIPLE_ACCOUNTS' })
  })

  it('…same account on both origins is fine', async () => {
    const t = await loadToken()
    tabs = [
      { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' },
      { id: 8, active: false, url: 'https://outlook.office365.com/mail/', status: 'complete' },
    ]
    sendMessage.mockImplementation(async (tabId: number) => okToken(tabId === 7 ? 'a' : 'b'))
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'a' })
  })
})
