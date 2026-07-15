// Content script — runs on outlook.office.com.
// Only job right now: hand the MSAL access token to the background on request.
// localStorage is shared between the page and the content-script isolated world,
// so we can read MSAL's cache directly.

import type { ContentRequest, ContentResponse } from '@/shared/messages'

function readMsalToken(): ContentResponse {
  const nowSec = Date.now() / 1000
  const candidates: Array<{ secret: string; expiresOn: number; homeAccountId: string }> = []

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

    candidates.push({
      secret: val.secret,
      expiresOn,
      homeAccountId: typeof val.homeAccountId === 'string' ? val.homeAccountId : '',
    })
  }

  if (candidates.length === 0) {
    return {
      ok: false,
      code: 'NO_TOKEN',
      message: 'localStorage 內找不到有效 MSAL token，可能尚未登入或 scope 不含 Mail.ReadWrite',
    }
  }

  // Multi-account guard (audit H4): when TWO accounts are signed into the
  // same OWA origin, localStorage holds a Mail.ReadWrite token for BOTH.
  // "Latest expiry" correlates with whichever account MSAL renewed most
  // recently — NOT with the mailbox on screen — and every destructive call
  // downstream goes to /me for whichever token wins. Picking wrong means
  // moving/deleting mail in the OTHER account's mailbox. We can't reliably
  // tell which mailbox the page is showing from the isolated world, so
  // refuse outright and tell the user to sign the extra account out.
  const distinctAccounts = new Set(
    candidates.map((c) => c.homeAccountId).filter((id) => id !== ''),
  )
  if (distinctAccounts.size > 1) {
    return {
      ok: false,
      code: 'MULTIPLE_ACCOUNTS',
      message:
        '偵測到此瀏覽器登入了多個 Microsoft 帳號。為避免整理到錯誤的信箱，請先登出其他帳號、只保留要整理的那一個，再重新載入 Outlook。',
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
