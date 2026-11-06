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

// C1 (2026-11-06): v1.0.8 accepted only the outlook.office.com / office365.com
// prefixes, so these working setups all became NO_TOKEN.
describe('readOutlookTokenFromMsalCache audience ranking (C1)', () => {
  beforeEach(() => localStorage.clear())

  it.each([
    ['outlook.cloud.microsoft (this user’s OWA)', 'https://outlook.cloud.microsoft/Mail.ReadWrite'],
    ['an explicit :443, any case', 'HTTPS://Outlook.Office.com:443/Mail.ReadWrite'],
    ['substrate.office.com', 'https://substrate.office.com/Mail.ReadWrite'],
    ['the Exchange Online app-id form', '00000002-0000-0ff1-ce00-000000000000/Mail.ReadWrite'],
  ])('accepts %s', (_label, target) => {
    putToken('msal.only', { target: `${target} openid profile` })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: true, secret: 'SECRET-msal.only' })
  })

  it('a known Outlook REST resource beats a later-expiring other (non-Graph) resource', () => {
    putToken('msal.cloud', { target: 'https://outlook.cloud.microsoft/Mail.ReadWrite', expiresOn: String(nowSec() + 600) })
    putToken('msal.substrate', {
      target: 'https://substrate.office.com/Mail.ReadWrite',
      expiresOn: String(nowSec() + 3500),
    })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: true, secret: 'SECRET-msal.cloud' })
  })

  it('…and the other resource is used when it is the only one left', () => {
    putToken('msal.cloud-expired', {
      target: 'https://outlook.cloud.microsoft/Mail.ReadWrite',
      expiresOn: String(nowSec() - 7200),
    })
    putToken('msal.substrate', { target: 'https://substrate.office.com/Mail.ReadWrite' })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: true, secret: 'SECRET-msal.substrate' })
  })

  it('never picks Graph — national clouds, the Graph app id or a bare scope — and lists the origins it saw', () => {
    putToken('msal.g1', { target: 'https://graph.microsoft.com/Mail.ReadWrite', expiresOn: String(nowSec() + 3599) })
    putToken('msal.g2', { target: 'https://graph.microsoft.us/Mail.ReadWrite' })
    putToken('msal.g3', { target: 'https://dod-graph.microsoft.us/Mail.ReadWrite' })
    putToken('msal.g4', { target: 'https://microsoftgraph.chinacloudapi.cn/Mail.ReadWrite' })
    putToken('msal.g5', { target: '00000003-0000-0000-c000-000000000000/Mail.ReadWrite' })
    putToken('msal.g6', { target: 'Mail.ReadWrite User.Read' })
    const r = readOutlookTokenFromMsalCache()
    expect(r).toMatchObject({ ok: false, code: 'NO_TOKEN' })
    if (r.ok) return
    expect(r.message).toContain('非 Outlook 資源 6')
    expect(r.message).toContain('https://graph.microsoft.com')
    expect(r.message).toContain('https://graph.microsoft.us')
    expect(r.message).toContain('(無資源前綴)')
    expect(r.message).not.toContain('SECRET')
  })

  it('the multi-account guard counts every usable rank', () => {
    putToken('msal.a', { homeAccountId: 'acct-A' })
    putToken('msal.b', { homeAccountId: 'acct-B', target: 'https://substrate.office.com/Mail.ReadWrite' })
    expect(readOutlookTokenFromMsalCache()).toMatchObject({ ok: false, code: 'MULTIPLE_ACCOUNTS' })
  })

  it('the ranked reader still runs standalone', () => {
    putToken('msal.cloud', { target: 'https://outlook.cloud.microsoft:443/Mail.ReadWrite' })
    const standalone = new Function(`return (${readOutlookTokenFromMsalCache.toString()})`)() as () => unknown
    expect(standalone()).toEqual(readOutlookTokenFromMsalCache())
    expect(standalone()).toMatchObject({ ok: true, secret: 'SECRET-msal.cloud' })
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

  it('keeps the token’s account id in the cache (C10) and peeks it without a refresh', async () => {
    const t = await loadToken()
    sendMessage.mockResolvedValue(okToken('fresh', 3600, 'acct-9'))
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'fresh', accountId: 'acct-9' })
    sendMessage.mockClear()
    expect(await t.peekCachedTokenAccountId()).toBe('acct-9')
    expect(await t.getOwaToken()).toMatchObject({ accountId: 'acct-9' })
    expect(sendMessage).not.toHaveBeenCalled()
  })
})

// C5 (2026-11-06): chrome.tabs.sendMessage has no timeout. A frozen tab left
// the shared in-flight refresh — and so every Outlook request — pending forever.
describe('token acquisition never hangs on a tab that does not answer (C5)', () => {
  beforeEach(() => {
    installChrome()
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })
  const never = () => new Promise<never>(() => {})

  it('a frozen tab on ANOTHER origin is "can’t tell": the token still arrives', async () => {
    const t = await loadToken()
    tabs = [
      { id: 7, active: true, url: 'https://outlook.cloud.microsoft/mail/', status: 'complete' },
      { id: 8, active: false, url: 'https://outlook.office.com/mail/', status: 'complete' },
    ]
    sendMessage.mockImplementation((tabId: number) => (tabId === 7 ? Promise.resolve(okToken('a')) : never()))
    let settled: unknown = 'pending'
    void t.getOwaToken().then((r) => (settled = r))
    await vi.advanceTimersByTimeAsync(1_600)
    expect(settled).toMatchObject({ ok: true, secret: 'a' })
    // The coalesced refresh was released too: the next caller isn't stuck.
    expect(await t.getOwaToken()).toMatchObject({ ok: true, secret: 'a' })
  })

  it('the chosen tab never answering → TAB_UNREACHABLE instead of hanging every caller', async () => {
    const t = await loadToken()
    tabs = [{ id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' }]
    sendMessage.mockImplementation(() => never())
    const results: unknown[] = []
    void t.getOwaToken().then((r) => results.push(r))
    void t.getOwaToken({ force: true }).then((r) => results.push(r)) // coalesces onto the same refresh
    await vi.advanceTimersByTimeAsync(10_500)
    expect(results).toHaveLength(2)
    expect(results[0]).toMatchObject({ ok: false, code: 'TAB_UNREACHABLE' })
    expect(executeScript).not.toHaveBeenCalled() // an unresponsive tab isn't re-injected into
  })
})
