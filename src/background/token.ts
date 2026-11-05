// Token cache + fetcher (service-worker side).
//
// Flow:
//   1. caller -> getOwaToken()
//   2. check chrome.storage.session for cached, unexpired token
//   3. if miss, locate an outlook.office.com tab
//   4. ask its content script via chrome.tabs.sendMessage
//      (orphaned listener → re-inject the script, then run the reader
//      directly via chrome.scripting.executeScript — 2026-11-05)
//   5. cache the result, return
//
// chrome.storage.session is in-memory, cleared on browser restart — fine for
// short-lived auth tokens. We deliberately do NOT store in chrome.storage.local
// to avoid persisting tokens to disk.

import type { ContentRequest, ContentResponse } from '@/shared/messages'
import { readOutlookTokenFromMsalCache, type MsalTokenReadResult } from '@/shared/msal-token-reader'

const SESSION_KEY = 'cachedToken'
const LAST_ERROR_KEY = 'lastTokenError'
// A cached token is USED while it has ≥ MIN_USABLE_SEC left, and refreshed
// in the BACKGROUND once it has < REFRESH_BUFFER_SEC left (2026-11-05). The
// old single 60 s buffer treated a 59-s-left token as missing: until OWA
// itself renewed (it renews late), EVERY request paid a content-script round
// trip that handed back the same token.
const MIN_USABLE_SEC = 15
const REFRESH_BUFFER_SEC = 60
// Background refresh attempts are throttled — while OWA hasn't renewed yet
// the content script keeps returning the same nearly-expired token.
const BACKGROUND_REFRESH_MIN_INTERVAL_MS = 15_000
// The crxjs loader imports the real module asynchronously, so its listener
// appears a moment after executeScript resolves.
const REINJECT_SETTLE_MS = 500

type CachedToken = { secret: string; expiresOn: number }

export type TokenResult =
  | { ok: true; secret: string; expiresOn: number }
  | {
      ok: false
      code: 'NO_TOKEN' | 'NOT_ON_OWA' | 'EXPIRED' | 'TAB_UNREACHABLE' | 'MULTIPLE_ACCOUNTS'
      message: string
    }

type TokenFailure = Extract<TokenResult, { ok: false }>

/** Last token-acquisition failure, for getStatus / the popup (2026-11-05). */
export type TokenErrorRecord = { code: string; message: string; at: number }

function nowSec(): number {
  return Date.now() / 1000
}

function isUsable(token: CachedToken | undefined | null): token is CachedToken {
  if (!token) return false
  return token.expiresOn - MIN_USABLE_SEC > nowSec()
}

function isNearExpiry(token: CachedToken): boolean {
  return token.expiresOn - REFRESH_BUFFER_SEC <= nowSec()
}

async function getCached(): Promise<CachedToken | null> {
  try {
    const result = await chrome.storage.session.get(SESSION_KEY)
    return (result[SESSION_KEY] as CachedToken | undefined) ?? null
  } catch {
    return null
  }
}

async function setCached(token: CachedToken | null): Promise<void> {
  if (token) {
    await chrome.storage.session.set({ [SESSION_KEY]: token })
  } else {
    await chrome.storage.session.remove(SESSION_KEY)
  }
}

/**
 * Token problems used to be invisible — the popup only saw "token invalid"
 * with no reason. The last failure is persisted (session storage; cleared on
 * the next success) so getStatus can show WHY (2026-11-05). Never contains
 * the token itself: messages are built from codes + counts only.
 */
async function recordTokenOutcome(result: TokenResult): Promise<void> {
  try {
    if (result.ok) {
      await chrome.storage.session.remove(LAST_ERROR_KEY)
    } else {
      const rec: TokenErrorRecord = { code: result.code, message: result.message, at: Date.now() }
      await chrome.storage.session.set({ [LAST_ERROR_KEY]: rec })
    }
  } catch {
    /* diagnostics only */
  }
}

export async function getLastTokenError(): Promise<{ code: string; message: string; at: number } | null> {
  try {
    const r = await chrome.storage.session.get(LAST_ERROR_KEY)
    const v = r[LAST_ERROR_KEY] as TokenErrorRecord | undefined
    if (v && typeof v.code === 'string' && typeof v.message === 'string' && typeof v.at === 'number') {
      return { code: v.code, message: v.message, at: v.at }
    }
  } catch {
    /* fall through */
  }
  return null
}

const OWA_URL_PATTERNS = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
]

async function listOwaTabs(): Promise<chrome.tabs.Tab[]> {
  return chrome.tabs.query({ url: OWA_URL_PATTERNS })
}

function pickOwaTab(tabs: chrome.tabs.Tab[]): chrome.tabs.Tab | null {
  // Prefer an active tab if multiple
  const active = tabs.find((t) => t.active)
  return active ?? tabs[0] ?? null
}

async function findOwaTab(): Promise<chrome.tabs.Tab | null> {
  return pickOwaTab(await listOwaTabs())
}

function isContentScriptNotReady(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err)
  // Chrome's canonical phrasing when the tab has no listener registered yet.
  return (
    msg.includes('Could not establish connection') ||
    msg.includes('Receiving end does not exist')
  )
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms))
}

function asTokenResult(resp: ContentResponse | MsalTokenReadResult | TokenFailure): TokenResult {
  // Drop extra fields (accountId) — only the documented shape leaves here.
  return resp.ok
    ? { ok: true, secret: resp.secret, expiresOn: resp.expiresOn }
    : { ok: false, code: resp.code, message: resp.message }
}

/**
 * The `js` files of the manifest's token-reader content script (2026-11-05).
 * It is the FIRST content_scripts entry in manifest.config.ts; crxjs rewrites
 * its path to a generated loader filename, so it must be read at runtime.
 */
function tokenReaderScriptFiles(): string[] {
  try {
    const entries = chrome.runtime.getManifest().content_scripts ?? []
    return (entries[0]?.js ?? []).filter((f): f is string => typeof f === 'string' && f.length > 0)
  } catch {
    return []
  }
}

async function reinjectTokenReader(tabId: number): Promise<boolean> {
  const files = tokenReaderScriptFiles()
  if (files.length === 0 || typeof chrome.scripting?.executeScript !== 'function') return false
  try {
    await chrome.scripting.executeScript({ target: { tabId }, files })
    return true
  } catch (e) {
    console.warn('[mail-organizer] content-script re-injection failed', e)
    return false
  }
}

/**
 * Run the reader directly in the page. Needed when re-injecting the files
 * doesn't restore a listener: the isolated world survives an extension
 * reload, and the crxjs loader's dynamic import() of an unchanged module URL
 * may resolve to the old, orphaned module without re-running it. (Not
 * verifiable in unit tests — implemented defensively.)
 */
async function readTokenViaExecuteScript(tabId: number): Promise<MsalTokenReadResult | null> {
  if (typeof chrome.scripting?.executeScript !== 'function') return null
  try {
    const [res] = await chrome.scripting.executeScript({
      target: { tabId },
      func: readOutlookTokenFromMsalCache,
    })
    const r = res?.result as MsalTokenReadResult | undefined
    if (r && typeof r === 'object' && typeof r.ok === 'boolean') return r
  } catch (e) {
    console.warn('[mail-organizer] direct token read via executeScript failed', e)
  }
  return null
}

async function askContentScript(
  tab: chrome.tabs.Tab,
): Promise<ContentResponse | MsalTokenReadResult | TokenFailure> {
  const tabId = tab.id!
  const req: ContentRequest = { type: 'fetch_token' }
  const noResponse: TokenFailure = {
    ok: false,
    code: 'NOT_ON_OWA',
    message: 'Content script 沒回應，OWA 分頁可能尚未載入完成',
  }
  const unreachable = (e: unknown): TokenFailure => ({
    ok: false,
    code: 'NOT_ON_OWA',
    message: `無法和 OWA 分頁通訊：${e instanceof Error ? e.message : String(e)}`,
  })
  const send = async (): Promise<ContentResponse | undefined> =>
    (await chrome.tabs.sendMessage(tabId, req)) as ContentResponse | undefined

  try {
    return (await send()) ?? noResponse
  } catch (e) {
    if (!isContentScriptNotReady(e)) return unreachable(e)
  }

  // No listener in the tab. A discarded tab has no live page at all.
  if (tab.discarded) {
    return {
      ok: false,
      code: 'TAB_UNREACHABLE',
      message: 'Outlook 分頁已被瀏覽器休眠(省記憶體)。請切換到該分頁讓它重新載入,再試一次。',
    }
  }

  // Still loading → the manifest injection just hasn't run yet: one retry
  // after a short delay (the original cold-start behaviour).
  if (tab.status === 'loading') {
    await delay(1500)
    try {
      return (await send()) ?? noResponse
    } catch (e) {
      if (!isContentScriptNotReady(e)) return unreachable(e)
    }
  }

  // Loaded but no listener = orphaned by an extension update/reload
  // (2026-11-05): before this, the user saw "wait a few seconds" forever
  // until they reloaded Outlook. Re-inject and retry once…
  if (await reinjectTokenReader(tabId)) {
    await delay(REINJECT_SETTLE_MS)
    try {
      const resp = await send()
      if (resp) return resp
    } catch (e) {
      if (!isContentScriptNotReady(e)) return unreachable(e)
    }
  }
  // …and as a last resort read the cache directly in the page.
  const direct = await readTokenViaExecuteScript(tabId)
  if (direct) return direct

  return {
    ok: false,
    code: 'TAB_UNREACHABLE',
    message: '無法和 Outlook 分頁通訊(擴充功能更新後通常需要重新載入 Outlook 分頁)。請重新整理 Outlook 分頁後再試。',
  }
}

/**
 * The content script only sees ITS origin's MSAL cache (2026-11-05). With
 * OWA open on two origins (e.g. outlook.office.com + outlook.office365.com)
 * signed into different accounts, the token we picked could belong to a
 * mailbox other than the one the user is looking at. Ask one tab per OTHER
 * origin and refuse on an account mismatch. Unanswering tabs are ignored
 * (can't tell → don't block).
 */
async function otherOriginAccountConflict(
  tabs: chrome.tabs.Tab[],
  chosen: chrome.tabs.Tab,
  chosenAccountId: string | undefined,
): Promise<boolean> {
  if (!chosenAccountId) return false
  const originOf = (t: chrome.tabs.Tab): string => {
    try {
      return t.url ? new URL(t.url).origin : ''
    } catch {
      return ''
    }
  }
  const chosenOrigin = originOf(chosen)
  const byOrigin = new Map<string, chrome.tabs.Tab>()
  for (const t of tabs) {
    const o = originOf(t)
    if (!o || o === chosenOrigin || t.id == null || t.discarded || byOrigin.has(o)) continue
    byOrigin.set(o, t)
  }
  for (const t of byOrigin.values()) {
    try {
      const resp = (await chrome.tabs.sendMessage(t.id!, { type: 'fetch_token' } satisfies ContentRequest)) as
        | MsalTokenReadResult
        | undefined
      if (resp?.ok && resp.accountId && resp.accountId !== chosenAccountId) return true
      if (resp && !resp.ok && resp.code === 'MULTIPLE_ACCOUNTS') return true
    } catch {
      /* orphaned / loading tab — can't tell */
    }
  }
  return false
}

// Coalesce concurrent token refreshes. Without this, 50 in-flight Outlook
// requests can each see a 401 + force-refresh, racing 50 askContentScript
// calls that reciprocally clobber the cache. With it, only the first call
// reaches the content script; the rest await the same promise.
let inFlightRefresh: Promise<TokenResult> | null = null
let lastBackgroundRefreshAt = 0

async function doFetchToken(opts: { keepUsableOnFailure: boolean }): Promise<TokenResult> {
  const result = await fetchTokenFromTabs()
  if (result.ok) {
    await setCached({ secret: result.secret, expiresOn: result.expiresOn })
  } else if (!(opts.keepUsableOnFailure && isUsable(await getCached()))) {
    // A failed BACKGROUND refresh must not throw away a token that still
    // works; a failed foreground/forced one (e.g. after a 401) must.
    await setCached(null)
  }
  await recordTokenOutcome(result)
  return result
}

async function fetchTokenFromTabs(): Promise<TokenResult> {
  const tabs = await listOwaTabs()
  const tab = pickOwaTab(tabs)
  if (!tab?.id) {
    return {
      ok: false,
      code: 'NOT_ON_OWA',
      message: '請先打開並登入 outlook.office.com 分頁',
    }
  }
  const resp = await askContentScript(tab)
  if (!resp.ok) return asTokenResult(resp)
  const accountId = 'accountId' in resp ? resp.accountId : undefined
  if (tabs.length > 1 && (await otherOriginAccountConflict(tabs, tab, accountId))) {
    return {
      ok: false,
      code: 'MULTIPLE_ACCOUNTS',
      message:
        '偵測到多個 Outlook 分頁登入了不同的 Microsoft 帳號。為避免整理到錯誤的信箱，請只保留要整理的那個帳號的 Outlook 分頁，再試一次。',
    }
  }
  return asTokenResult(resp)
}

function startRefresh(opts: { keepUsableOnFailure: boolean }): Promise<TokenResult> {
  const p: Promise<TokenResult> = doFetchToken(opts).finally(() => {
    if (inFlightRefresh === p) inFlightRefresh = null
  })
  inFlightRefresh = p
  return p
}

function maybeBackgroundRefresh(): void {
  if (inFlightRefresh) return
  if (Date.now() - lastBackgroundRefreshAt < BACKGROUND_REFRESH_MIN_INTERVAL_MS) return
  lastBackgroundRefreshAt = Date.now()
  void startRefresh({ keepUsableOnFailure: true }).catch(() => {
    /* background — the next foreground call surfaces failures */
  })
}

export async function getOwaToken(opts: { force?: boolean } = {}): Promise<TokenResult> {
  if (opts.force) {
    // Coalesce onto an in-flight refresh BEFORE any await. An in-flight
    // refresh is already fetching a brand-new token straight from the content
    // script (not the cache), so it is strictly fresher than whatever token
    // triggered a forced 401-retry. Letting force callers skip it was the old
    // 50-simultaneous-401s storm. `force` bypasses the *cache*, not an
    // in-flight fetch.
    if (inFlightRefresh) return inFlightRefresh
    return startRefresh({ keepUsableOnFailure: false })
  }

  // A usable cached token wins — even while a background refresh runs.
  const cached = await getCached()
  if (isUsable(cached)) {
    if (isNearExpiry(cached)) maybeBackgroundRefresh()
    return { ok: true, secret: cached.secret, expiresOn: cached.expiresOn }
  }
  // Re-check after the await — another caller may have started the
  // refresh while we were reading the cache.
  if (inFlightRefresh) return inFlightRefresh
  return startRefresh({ keepUsableOnFailure: false })
}

export async function clearTokenCache(): Promise<void> {
  await setCached(null)
}

/**
 * Read the cached token validity WITHOUT triggering a refresh on miss.
 * Used by getStatus so popup-open isn't gated by askContentScript's 1.5s
 * cold-start retry. If we don't have a valid cached token, the caller
 * can kick off a background refresh and rely on the next status check
 * (or the actual classify flow) to surface the fresh state.
 */
export async function peekCachedToken(): Promise<{ valid: boolean }> {
  const cached = await getCached()
  return { valid: isUsable(cached) }
}

export async function pingOwa(): Promise<{ ok: true } | { ok: false; code: 'NOT_ON_OWA'; message: string }> {
  const tab = await findOwaTab()
  if (!tab?.id) {
    return { ok: false, code: 'NOT_ON_OWA', message: '請先打開 outlook.office.com' }
  }
  return { ok: true }
}
