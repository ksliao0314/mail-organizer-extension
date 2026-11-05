// MSAL localStorage → Outlook REST access token (2026-11-05).
//
// Used in two places:
//   1. src/content/content.ts — the content script answers 'fetch_token'.
//   2. src/background/token.ts — last-resort chrome.scripting.executeScript
//      ({ func }) when the content script's listener is gone (orphaned by an
//      extension update) and re-injecting its files didn't bring it back.
//
// Because of (2) the function MUST be self-contained: executeScript ships
// only the function's source text into the page, so it may not reference
// imports, module-level constants or helpers — only its own body and
// built-ins. tests/token-reader.test.ts evaluates its toString() standalone
// to keep that true.

export type MsalTokenReadResult =
  | { ok: true; secret: string; expiresOn: number; accountId?: string }
  | { ok: false; code: 'NO_TOKEN' | 'MULTIPLE_ACCOUNTS'; message: string }

export function readOutlookTokenFromMsalCache(): MsalTokenReadResult {
  // Outlook REST v2.0 audiences. A token for another resource that merely
  // lists Mail.ReadWrite (graph.microsoft.com, or MSAL's bare-scope Graph
  // form) used to qualify — and with "latest expiry wins" it could beat the
  // Outlook token and 401-loop every request.
  const OUTLOOK_RESOURCES = ['https://outlook.office.com/', 'https://outlook.office365.com/']
  // An account whose Outlook token expired this recently was in active use;
  // OWA refreshes the on-screen account's token before expiry, so a short
  // window catches "two accounts in this origin" without blocking a user
  // for long after they switched accounts.
  const RECENTLY_ACTIVE_SEC = 15 * 60
  const nowSec = Date.now() / 1000

  const isOutlookMailScope = (target: string): boolean =>
    target.split(/\s+/).some((scope) => {
      const s = scope.toLowerCase()
      return OUTLOOK_RESOURCES.some((r) => s.startsWith(r)) && s.includes('mail.readwrite')
    })

  // Diagnostics for the failure message — counts only, never secrets/keys.
  let msalKeys = 0
  let jsonEntries = 0
  let nonJson = 0
  let accessTokens = 0
  let otherResource = 0
  let expired = 0

  const candidates: Array<{ secret: string; expiresOn: number; accountId: string }> = []
  const activeAccounts = new Set<string>()

  let store: Storage
  try {
    store = window.localStorage
  } catch {
    return { ok: false, code: 'NO_TOKEN', message: '無法讀取 Outlook 分頁的 localStorage' }
  }

  for (let i = 0; i < store.length; i++) {
    const key = store.key(i)
    if (!key || !key.includes('msal')) continue
    msalKeys++
    const raw = store.getItem(key)
    if (!raw) continue
    let val: unknown
    try {
      val = JSON.parse(raw)
    } catch {
      nonJson++
      continue
    }
    if (!val || typeof val !== 'object') {
      nonJson++
      continue
    }
    jsonEntries++
    const entry = val as Record<string, unknown>
    if (entry.credentialType !== 'AccessToken') continue
    accessTokens++
    const target = typeof entry.target === 'string' ? entry.target : ''
    if (!isOutlookMailScope(target)) {
      otherResource++
      continue
    }
    const expiresOn = Number.parseInt(String(entry.expiresOn ?? ''), 10)
    if (!Number.isFinite(expiresOn)) continue
    const accountId = typeof entry.homeAccountId === 'string' ? entry.homeAccountId : ''
    if (expiresOn <= nowSec) {
      expired++
      if (accountId && nowSec - expiresOn < RECENTLY_ACTIVE_SEC) activeAccounts.add(accountId)
      continue
    }
    if (typeof entry.secret !== 'string' || !entry.secret) continue
    if (accountId) activeAccounts.add(accountId)
    candidates.push({ secret: entry.secret, expiresOn, accountId })
  }

  // Multi-account guard (audit H4, widened 2026-11-05): two accounts in the
  // same OWA origin each hold an Outlook token; "latest expiry" follows
  // whichever MSAL renewed last, NOT the mailbox on screen, and every
  // destructive call goes to /me of the winning token. Counting only VALID
  // tokens missed the on-screen account whose token had just lapsed.
  if (activeAccounts.size > 1) {
    return {
      ok: false,
      code: 'MULTIPLE_ACCOUNTS',
      message:
        '偵測到此瀏覽器登入了多個 Microsoft 帳號。為避免整理到錯誤的信箱，請先登出其他帳號、只保留要整理的那一個，再重新載入 Outlook。',
    }
  }

  if (candidates.length === 0) {
    const encryptedHint =
      nonJson > 0 && accessTokens === 0
        ? ' MSAL 快取內容不是 JSON,可能是新版 Outlook 的加密快取,本工具目前讀不到。'
        : ''
    return {
      ok: false,
      code: 'NO_TOKEN',
      message:
        `localStorage 內找不到可用的 Outlook REST token,可能尚未登入或 token 已過期(請重新載入 Outlook)。` +
        `診斷:msal 鍵 ${msalKeys} 個、可解析 ${jsonEntries}、非 JSON ${nonJson}、AccessToken ${accessTokens}` +
        `(非 Outlook 資源 ${otherResource}、已過期 ${expired})。${encryptedHint}`,
    }
  }

  // Single account: pick the one expiring latest (most recently refreshed)
  candidates.sort((a, b) => b.expiresOn - a.expiresOn)
  const best = candidates[0]!
  return {
    ok: true,
    secret: best.secret,
    expiresOn: best.expiresOn,
    ...(best.accountId ? { accountId: best.accountId } : {}),
  }
}
