// Content script — runs on outlook.office.com.
// Only job right now: hand the MSAL access token to the background on request.
// localStorage is shared between the page and the content-script isolated world,
// so we can read MSAL's cache directly.
//
// The reader lives in src/shared/msal-token-reader.ts (2026-11-05) so the
// service worker can also run it via chrome.scripting.executeScript when this
// script's listener is orphaned after an extension update.

import type { ContentRequest, ContentResponse } from '@/shared/messages'
import { readOutlookTokenFromMsalCache } from '@/shared/msal-token-reader'

// No "already registered" global guard on purpose: after an extension reload
// the isolated world (and any global flag) survives while the OLD listener is
// dead, so a guard would block exactly the re-injection that heals it. A
// duplicate live listener is harmless — Chrome keeps the first response.
chrome.runtime.onMessage.addListener(
  (msg: ContentRequest, _sender, sendResponse: (resp: ContentResponse) => void) => {
    if (msg?.type === 'fetch_token') {
      sendResponse(readOutlookTokenFromMsalCache())
      return false // sync response
    }
    return false
  },
)

console.debug('[mail-organizer] content script ready on', location.host)
