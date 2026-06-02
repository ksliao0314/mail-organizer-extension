// Token cache + fetcher (service-worker side).
//
// Flow:
//   1. caller -> getOwaToken()
//   2. check chrome.storage.session for cached, unexpired token
//   3. if miss, locate an outlook.office.com tab
//   4. ask its content script via chrome.tabs.sendMessage
//   5. cache the result, return
//
// chrome.storage.session is in-memory, cleared on browser restart — fine for
// short-lived auth tokens. We deliberately do NOT store in chrome.storage.local
// to avoid persisting tokens to disk.

import type { ContentRequest, ContentResponse } from '@/shared/messages'

const SESSION_KEY = 'cachedToken'
const REFRESH_BUFFER_SEC = 60 // refresh if expiring within this window

type CachedToken = { secret: string; expiresOn: number }

export type TokenResult =
  | { ok: true; secret: string; expiresOn: number }
  | { ok: false; code: 'NO_TOKEN' | 'NOT_ON_OWA' | 'EXPIRED' | 'TAB_UNREACHABLE'; message: string }

function nowSec(): number {
  return Date.now() / 1000
}

function isValid(token: CachedToken | undefined | null): token is CachedToken {
  if (!token) return false
  return token.expiresOn - REFRESH_BUFFER_SEC > nowSec()
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

const OWA_URL_PATTERNS = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
]

async function findOwaTab(): Promise<chrome.tabs.Tab | null> {
  const tabs = await chrome.tabs.query({ url: OWA_URL_PATTERNS })
  // Prefer an active tab if multiple
  const active = tabs.find((t) => t.active)
  return active ?? tabs[0] ?? null
}

function isContentScriptNotReady(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err)
  // Chrome's canonical phrasing when the tab has no listener registered yet.
  return (
    msg.includes('Could not establish connection') ||
    msg.includes('Receiving end does not exist')
  )
}

async function askContentScript(tabId: number): Promise<ContentResponse> {
  const req: ContentRequest = { type: 'fetch_token' }
  // One retry after a short delay handles the case where the user just opened
  // OWA and clicked the extension before the content script finished
  // injecting. Without this, the user has to reload OWA to recover.
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const resp = (await chrome.tabs.sendMessage(tabId, req)) as ContentResponse | undefined
      if (!resp) {
        return {
          ok: false,
          code: 'NOT_ON_OWA',
          message: 'Content script 沒回應，OWA 分頁可能尚未載入完成',
        }
      }
      return resp
    } catch (e) {
      if (attempt === 0 && isContentScriptNotReady(e)) {
        await new Promise((r) => setTimeout(r, 1500))
        continue
      }
      return {
        ok: false,
        code: 'NOT_ON_OWA',
        message: `無法和 OWA 分頁通訊：${e instanceof Error ? e.message : String(e)}`,
      }
    }
  }
  // Unreachable — loop either returns or breaks
  return { ok: false, code: 'NOT_ON_OWA', message: 'Content script 重試後仍無回應' }
}

// Coalesce concurrent token refreshes. Without this, 50 in-flight Outlook
// requests can each see a 401 + force-refresh, racing 50 askContentScript
// calls that reciprocally clobber the cache. With it, only the first call
// reaches the content script; the rest await the same promise.
let inFlightRefresh: Promise<TokenResult> | null = null

async function doFetchToken(): Promise<TokenResult> {
  const tab = await findOwaTab()
  if (!tab?.id) {
    return {
      ok: false,
      code: 'NOT_ON_OWA',
      message: '請先打開並登入 outlook.office.com 分頁',
    }
  }
  const resp = await askContentScript(tab.id)
  if (resp.ok) {
    await setCached({ secret: resp.secret, expiresOn: resp.expiresOn })
    return { ok: true, secret: resp.secret, expiresOn: resp.expiresOn }
  }
  await setCached(null)
  return resp as TokenResult
}

export async function getOwaToken(opts: { force?: boolean } = {}): Promise<TokenResult> {
  // Check in-flight refresh BEFORE any await so concurrent callers
  // piggyback immediately instead of each issuing their own storage read.
  // Without this, 50 simultaneous callers do 50 redundant storage.session
  // reads before one of them notices the in-flight refresh.
  if (!opts.force && inFlightRefresh) return inFlightRefresh

  if (!opts.force) {
    const cached = await getCached()
    if (isValid(cached)) {
      return { ok: true, secret: cached.secret, expiresOn: cached.expiresOn }
    }
    // Re-check after the await — another caller may have started the
    // refresh while we were reading the cache.
    if (inFlightRefresh) return inFlightRefresh
  }
  inFlightRefresh = doFetchToken().finally(() => {
    inFlightRefresh = null
  })
  return inFlightRefresh
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
  return { valid: isValid(cached) }
}

export async function pingOwa(): Promise<{ ok: true } | { ok: false; code: 'NOT_ON_OWA'; message: string }> {
  const tab = await findOwaTab()
  if (!tab?.id) {
    return { ok: false, code: 'NOT_ON_OWA', message: '請先打開 outlook.office.com' }
  }
  return { ok: true }
}
