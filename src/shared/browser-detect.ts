// Browser detection helper used by sync UI to render browser-aware
// hints. The underlying chrome.storage.sync API is identical across
// Chromium browsers — only the user-facing text changes ("via Microsoft
// account" on Edge, "via Google account" on Chrome).
//
// Detection is best-effort. User-agent strings can be spoofed, but
// nothing here is security-sensitive — wrong detection just shows
// slightly misleading help text.

export type BrowserKind = 'edge' | 'chrome' | 'firefox' | 'other'

/**
 * Identify the running browser via userAgent. Edge MUST be checked
 * before Chrome because Edge's UA contains "Chrome/" — order matters.
 */
export function detectBrowser(): BrowserKind {
  const ua = typeof navigator !== 'undefined' ? navigator.userAgent : ''
  if (/\bEdg\//.test(ua)) return 'edge'
  if (/\bFirefox\//.test(ua)) return 'firefox'
  if (/\bChrome\//.test(ua)) return 'chrome'
  return 'other'
}

/**
 * Short Chinese phrase describing where sync data goes through, used
 * inside sentences in the options UI.
 */
export function syncAccountDescription(browser: BrowserKind = detectBrowser()): string {
  switch (browser) {
    case 'edge':
      return '透過你登入 Edge 的 Microsoft 帳號'
    case 'chrome':
      return '透過你登入 Chrome 的 Google 帳號'
    case 'firefox':
      return '透過你登入 Firefox 的 Firefox Account'
    default:
      return '透過你的瀏覽器帳號'
  }
}

/**
 * Settings URL where the user enables sync + extension sync. Empty
 * string when unknown — caller should suppress the link instead of
 * rendering a dead one.
 */
export function syncSettingsUrl(browser: BrowserKind = detectBrowser()): string {
  switch (browser) {
    case 'edge':
      return 'edge://settings/profile/sync'
    case 'chrome':
      return 'chrome://settings/syncSetup'
    default:
      return ''
  }
}

/**
 * Display name for the browser — used in headlines / status messages.
 */
export function browserLabel(browser: BrowserKind = detectBrowser()): string {
  switch (browser) {
    case 'edge':
      return 'Edge'
    case 'chrome':
      return 'Chrome'
    case 'firefox':
      return 'Firefox'
    default:
      return '瀏覽器'
  }
}
