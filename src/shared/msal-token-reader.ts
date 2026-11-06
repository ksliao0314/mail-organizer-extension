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
  // Audience ranking (2026-11-06). v1.0.8 accepted ONLY the outlook.office.com
  // / outlook.office365.com prefixes, which turned working setups into "no
  // token": outlook.cloud.microsoft (this user's OWA), an explicit :443,
  // substrate.office.com and the Exchange app-id form all failed. Before that
  // (de733bd) any target listing Mail.ReadWrite qualified and "latest expiry
  // wins" let a Graph token beat the Outlook one → 401 loop. Now:
  //   rank 2 — a known Outlook REST resource (host match, any case, :443 ok);
  //   rank 1 — any other RESOURCE-QUALIFIED Mail.ReadWrite scope that isn't
  //            Microsoft Graph (incl. national clouds / Graph's app id);
  //   rank 0 — not usable: Graph, or a bare scope with no resource prefix
  //            (MSAL's Graph shorthand).
  // The best rank present wins; expiry and the multi-account guard decide
  // within it.
  const OUTLOOK_REST_HOSTS = ['outlook.office.com', 'outlook.office365.com', 'outlook.cloud.microsoft']
  const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000'
  // An account whose Outlook token expired this recently was in active use;
  // OWA refreshes the on-screen account's token before expiry, so a short
  // window catches "two accounts in this origin" without blocking a user
  // for long after they switched accounts.
  const RECENTLY_ACTIVE_SEC = 15 * 60
  const MAX_LISTED_RESOURCES = 8
  const nowSec = Date.now() / 1000

  /** `https://host[:port]/Scope` → { resource, name }; bare scope → resource ''. */
  const splitScope = (scope: string): { resource: string; name: string } => {
    const slash = scope.lastIndexOf('/')
    if (slash < 0) return { resource: '', name: scope }
    return { resource: scope.slice(0, slash), name: scope.slice(slash + 1) }
  }

  /** Display form of a scope's resource, for diagnostics (never a secret). */
  const resourceLabel = (resource: string): string => {
    if (!resource) return '(無資源前綴)'
    if (/^https?:\/\//i.test(resource)) {
      try {
        return new URL(resource).origin
      } catch {
        return resource.slice(0, 80)
      }
    }
    return resource.slice(0, 80)
  }

  const rankResource = (resource: string): number => {
    if (!resource) return 0 // bare scope = MSAL's Graph shorthand
    const lower = resource.toLowerCase()
    if (/^https?:\/\//.test(lower)) {
      let host = ''
      try {
        host = new URL(lower).hostname
      } catch {
        return 0
      }
      if (OUTLOOK_REST_HOSTS.includes(host)) return 2
      if (host.includes('graph.microsoft') || host === 'microsoftgraph.chinacloudapi.cn') return 0
      return 1
    }
    if (lower === GRAPH_APP_ID) return 0
    return 1
  }

  /** Best rank of any Mail.ReadWrite scope in `target`. */
  const rankTarget = (target: string): number => {
    let best = 0
    for (const scope of target.split(/\s+/)) {
      if (!scope) continue
      const { resource, name } = splitScope(scope)
      if (!name.toLowerCase().startsWith('mail.readwrite')) continue
      best = Math.max(best, rankResource(resource))
    }
    return best
  }

  // Diagnostics for the failure message — counts and resource origins only,
  // never secrets/keys.
  let msalKeys = 0
  let jsonEntries = 0
  let nonJson = 0
  let accessTokens = 0
  let otherResource = 0
  let expired = 0
  const resourcesSeen = new Set<string>()

  const candidates: Array<{ secret: string; expiresOn: number; accountId: string; rank: number }> = []
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
    for (const scope of target.split(/\s+/)) {
      if (scope && resourcesSeen.size < MAX_LISTED_RESOURCES) {
        resourcesSeen.add(resourceLabel(splitScope(scope).resource))
      }
    }
    const rank = rankTarget(target)
    if (rank === 0) {
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
    candidates.push({ secret: entry.secret, expiresOn, accountId, rank })
  }

  // Multi-account guard (audit H4, widened 2026-11-05): two accounts in the
  // same OWA origin each hold an Outlook token; "latest expiry" follows
  // whichever MSAL renewed last, NOT the mailbox on screen, and every
  // destructive call goes to /me of the winning token. Counting only VALID
  // tokens missed the on-screen account whose token had just lapsed. Every
  // usable rank counts (2026-11-06): one account's Outlook token next to
  // another account's substrate token is still two mailboxes.
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
    const resources = resourcesSeen.size > 0 ? `;token 資源:${[...resourcesSeen].join('、')}` : ''
    return {
      ok: false,
      code: 'NO_TOKEN',
      message:
        `localStorage 內找不到可用的 Outlook REST token,可能尚未登入或 token 已過期(請重新載入 Outlook)。` +
        `診斷:msal 鍵 ${msalKeys} 個、可解析 ${jsonEntries}、非 JSON ${nonJson}、AccessToken ${accessTokens}` +
        `(非 Outlook 資源 ${otherResource}、已過期 ${expired}${resources})。${encryptedHint}`,
    }
  }

  // Best rank first; within it, the one expiring latest (most recently refreshed).
  candidates.sort((a, b) => b.rank - a.rank || b.expiresOn - a.expiresOn)
  const best = candidates[0]!
  return {
    ok: true,
    secret: best.secret,
    expiresOn: best.expiresOn,
    ...(best.accountId ? { accountId: best.accountId } : {}),
  }
}
