// Content script — runs on outlook.office.com.
// Only job right now: hand the MSAL access token to the background on request.
// localStorage is shared between the page and the content-script isolated world,
// so we can read MSAL's cache directly.

import type { ContentRequest, ContentResponse } from '@/shared/messages'

function readMsalToken(): ContentResponse {
  const nowSec = Date.now() / 1000
  const candidates: Array<{ secret: string; expiresOn: number }> = []

  for (const key of Object.keys(localStorage)) {
    if (!key.includes('msal')) continue
    const raw = localStorage.getItem(key)
    if (!raw) continue

    let val: Record<string, unknown>
    try {
      val = JSON.parse(raw) as Record<string, unknown>
    } catch {
      continue
    }

    if (val.credentialType !== 'AccessToken') continue
    if (typeof val.target !== 'string' || !val.target.includes('Mail.ReadWrite')) continue
    if (typeof val.secret !== 'string' || !val.secret) continue
    if (typeof val.expiresOn !== 'string') continue

    const expiresOn = Number.parseInt(val.expiresOn, 10)
    if (!Number.isFinite(expiresOn)) continue
    if (expiresOn <= nowSec) continue

    candidates.push({ secret: val.secret, expiresOn })
  }

  if (candidates.length === 0) {
    return {
      ok: false,
      code: 'NO_TOKEN',
      message: 'localStorage 內找不到有效 MSAL token，可能尚未登入或 scope 不含 Mail.ReadWrite',
    }
  }

  // Pick the one expiring latest (most recently refreshed)
  candidates.sort((a, b) => b.expiresOn - a.expiresOn)
  const best = candidates[0]!
  return { ok: true, secret: best.secret, expiresOn: best.expiresOn }
}

chrome.runtime.onMessage.addListener(
  (msg: ContentRequest, _sender, sendResponse: (resp: ContentResponse) => void) => {
    if (msg?.type === 'fetch_token') {
      sendResponse(readMsalToken())
      return false // sync response
    }
    return false
  },
)

console.debug('[mail-organizer] content script ready on', location.host)
