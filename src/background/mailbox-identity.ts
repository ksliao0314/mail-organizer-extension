// The signed-in mailbox's identity (moved out of service-worker.ts, 2026-11-06 r2).
//
// Own mailbox domain (2026-11-05). With internalDomains empty (the default),
// learning used to mint `domain: <own domain>` from a colleague's mail, and
// since domain rules match ANY address incl. To/Cc, that one rule claimed
// nearly every inbound email. rules.ts keeps own-domain addresses from
// triggering AUTO domain rules, and learning / the initial scan never mint
// them — given the domain. Resolved from GET /me once and cached for a week;
// any failure falls back to the cached value, or to none (the pre-fix
// behaviour).
//
// 2026-11-06: the identity is keyed by the token's account (MSAL
// homeAccountId) — after an OWA account switch it was another mailbox's for
// up to a week; the "just failed, don't retry for an hour" memo lives in
// session storage (it reset on every SW restart, i.e. constantly); and the
// lookup skips the Outlook rate limiter's queue — its 3 s budget could expire
// while it merely waited behind a fan-out.
//
// 2026-11-06 r2:
//   - ONE account-checked accessor (getAccountCheckedIdentity) for every
//     reader. The learners and the initial scan read the raw cache: after an
//     account switch they excluded the OLD mailbox's domain while the matcher
//     already refused that identity.
//   - The failure memo is per account (a single slot let two alternating
//     accounts reset each other's).
//   - The 3 s budget covers the forced token refresh after a 401 too (the
//     request's token wait honours its signal — outlook-api.ts).
//   - The answer is stamped with the account peeked BEFORE the lookup, and
//     only kept when the account is still the same after it.

import type { OutlookApi } from '@/shared/outlook-api'
import { getMailboxIdentity, setMailboxIdentity, type MailboxIdentity } from '@/shared/storage'
import { peekCachedTokenAccountId } from './token'

export const MAILBOX_IDENTITY_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000
export const IDENTITY_LOOKUP_TIMEOUT_MS = 3_000
export const IDENTITY_RETRY_AFTER_FAILURE_MS = 60 * 60_000
const IDENTITY_FAILURE_KEY = 'mailboxIdentityFailure'

/** Failed lookups: account (MSAL homeAccountId; '' = unknown) → when (ms). */
type IdentityFailures = Record<string, number>

async function readIdentityFailures(): Promise<IdentityFailures> {
  try {
    const v = (await chrome.storage.session.get(IDENTITY_FAILURE_KEY))[IDENTITY_FAILURE_KEY] as unknown
    if (!v || typeof v !== 'object') return {}
    const byAccount = (v as { byAccount?: unknown }).byAccount
    if (byAccount && typeof byAccount === 'object') {
      const out: IdentityFailures = {}
      for (const [k, t] of Object.entries(byAccount as Record<string, unknown>)) {
        if (typeof t === 'number' && Number.isFinite(t)) out[k] = t
      }
      return out
    }
    // The earlier single slot: { at, accountId? }.
    const legacy = v as { at?: unknown; accountId?: unknown }
    if (typeof legacy.at === 'number') {
      return { [typeof legacy.accountId === 'string' ? legacy.accountId : '']: legacy.at }
    }
  } catch {
    /* unreadable → nothing recorded */
  }
  return {}
}

// Read-modify-write of the memo, serialised within this worker.
let failuresChain: Promise<void> = Promise.resolve()

function updateIdentityFailures(apply: (m: IdentityFailures) => void): Promise<void> {
  const run = failuresChain.then(async () => {
    const m = await readIdentityFailures()
    apply(m)
    const now = Date.now()
    for (const [k, t] of Object.entries(m)) if (!(now - t < IDENTITY_RETRY_AFTER_FAILURE_MS)) delete m[k]
    await chrome.storage.session.set({ [IDENTITY_FAILURE_KEY]: { byAccount: m } })
  })
  failuresChain = run.catch(() => {})
  return failuresChain
}

/** False only for an identity stamped with ANOTHER account than `accountId`. */
function belongsTo(identity: MailboxIdentity, accountId: string | undefined): boolean {
  return !(identity.accountId && accountId && identity.accountId !== accountId)
}

/**
 * The cached mailbox identity, unless it was stamped with another account
 * than the cached token's. One cached before account keying — or with no
 * token account known — can't be told apart and is used. Never throws.
 */
export async function getAccountCheckedIdentity(): Promise<MailboxIdentity | null> {
  const [cached, accountId] = await Promise.all([
    getMailboxIdentity().catch(() => null),
    peekCachedTokenAccountId().catch(() => undefined),
  ])
  return cached && belongsTo(cached, accountId) ? cached : null
}

/**
 * The signed-in mailbox: the account-checked cache while it is fresh, else a
 * bounded GET /me (never throws; on failure the usable cache, or null).
 * `fetch: false` = cache only: classify preflight refreshes the identity, so
 * the AI pass right after it and the background harvest never spend a call.
 */
export async function resolveMailboxIdentity(
  api: Pick<OutlookApi, 'getMyEmail'>,
  opts: { fetch?: boolean } = {},
): Promise<MailboxIdentity | null> {
  const [cached, accountId] = await Promise.all([
    getMailboxIdentity().catch(() => null),
    peekCachedTokenAccountId().catch(() => undefined),
  ])
  // Another account's identity is never used. One cached before account
  // keying is still trusted from cache, but refetched once to stamp it.
  const usable = cached !== null && belongsTo(cached, accountId) ? cached : null
  const cachedAt = usable ? Date.parse(usable.at) : NaN
  const fresh =
    usable !== null &&
    !(accountId && !usable.accountId) &&
    Number.isFinite(cachedAt) &&
    Date.now() - cachedAt < MAILBOX_IDENTITY_MAX_AGE_MS
  if (fresh || opts.fetch === false) return usable
  // A lookup that just failed (for this account) isn't retried for an hour,
  // so a persistently failing /me can't add its timeout to every preflight.
  const accountKey = accountId ?? ''
  const failedAt = (await readIdentityFailures())[accountKey]
  if (failedAt !== undefined && Date.now() - failedAt < IDENTITY_RETRY_AFTER_FAILURE_MS) return usable
  try {
    // Bounded: a failing /me would otherwise sit in request()'s 5xx backoff
    // (1 s + 3 s + 9 s) and hold up classify preflight — the identity is a
    // nice-to-have, never worth that.
    const address = (
      await api.getMyEmail(AbortSignal.timeout(IDENTITY_LOOKUP_TIMEOUT_MS), { unthrottled: true })
    )
      .trim()
      .toLowerCase()
    // Whose answer is it? Only the account the token had BEFORE the call,
    // and only if it still has it: a switch (or a 401 refresh landing on
    // another account) mid-call leaves it ambiguous — keep nothing.
    const after = await peekCachedTokenAccountId().catch(() => undefined)
    if (after !== accountId) {
      console.warn('[mail-organizer] mailbox identity lookup straddled an account change — not cached')
      return cached && belongsTo(cached, after) ? cached : null
    }
    const at = address.lastIndexOf('@')
    if (at > 0 && at < address.length - 1) {
      const identity: MailboxIdentity = {
        address,
        domain: address.slice(at + 1),
        at: new Date().toISOString(),
        ...(accountId ? { accountId } : {}),
      }
      await setMailboxIdentity(identity)
      await updateIdentityFailures((m) => {
        delete m[accountKey]
      })
      return identity
    }
  } catch (e) {
    console.warn('[mail-organizer] mailbox identity lookup failed (non-fatal)', e)
  }
  await updateIdentityFailures((m) => {
    m[accountKey] = Date.now()
  })
  return usable
}
