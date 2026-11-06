// Outlook REST v2.0 wrapper.
//
// Endpoint base: https://outlook.office.com/api/v2.0
// Auth:          Bearer <MSAL access token from OWA localStorage>
//
// Retry policy:
//   - 401 once: refresh token then retry
//   - 429 (any method) / 5xx / network (idempotent methods only):
//     exponential backoff (3^attempt sec), 3x max
//   - other 4xx: surface as OutlookError immediately
//
// Rate limiting via module-level TokenBucket (OUTLOOK_REQUEST_BURST burst /
// OUTLOOK_REQUESTS_PER_SEC per sec — 40 / 8 — FIFO reservations) + a
// module-level Semaphore capping requests in flight at 4.

import { OUTLOOK_REQUEST_BURST, OUTLOOK_REQUESTS_PER_SEC } from './constants'
import { Semaphore, TokenBucket, abortable, abortableSleep, linkAbortSignals } from './throttle'
import type {
  AiHiddenFolderRole,
  Email,
  MailFolder,
  MailFolderNode,
  WellKnownFolderRole,
  WellKnownFolderRoot,
  WellKnownFolders,
} from './types'

function isAbortError(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError'
}

const DEFAULT_BASE = 'https://outlook.office.com/api/v2.0'
// OUTLOOK_REQUESTS_PER_SEC steady, OUTLOOK_REQUEST_BURST burst (8/s, 40 since
// v1.0.9; 1/s, 10 before) — see constants.
const defaultBucket = new TokenBucket(OUTLOOK_REQUEST_BURST, OUTLOOK_REQUESTS_PER_SEC)

// Exchange throttles concurrent requests PER MAILBOX at ~4 (2026-11-05).
// Module-level so every OutlookApi instance shares one budget — the limit is
// the mailbox's, not the instance's.
export const MAX_CONCURRENT_REQUESTS = 4
const defaultInflight = new Semaphore(MAX_CONCURRENT_REQUESTS)

// Hosts allowed to receive our bearer token when following an absolute
// `@odata.nextLink` (2026-11-05). A paging URL comes from a response body;
// never send the token anywhere that isn't an Outlook REST endpoint.
const KNOWN_REST_ORIGINS = new Set([
  'https://outlook.office.com',
  'https://outlook.office365.com',
  'https://outlook.cloud.microsoft',
])

// Folder listings follow @odata.nextLink (2026-11-05). The cap only exists to
// stop a buggy/looping server; 50 pages × 200 children is far beyond any real
// mailbox, so hitting it is treated as an error rather than a silent
// truncation (a truncated listing is exactly what orphaned rules before).
const FOLDER_LIST_MAX_PAGES = 50

/**
 * Outlook REST distinguished folder names (GET /me/MailFolders/{name})
 * resolved with every folder tree (2026-11-06 r2 — see WellKnownFolders):
 *   - deleteditems / drafts / outbox: left OUT of the tree, root and subtree,
 *     and their subtrees are never listed (round 1 walked the whole trash on
 *     every fetch). Deleting folder 案件/甲 in OWA moves it to 刪除的郵件/甲
 *     with the SAME id — absent from the tree, its rules get orphaned, as in
 *     v1.0.8. Deleted Items ITSELF stays a valid rule target (auto-trash);
 *     execute runs such a move as a delete.
 *   - junkemail / sentitems: IN the tree (ordinary targets for rules, the
 *     picker and execute — v1.0.8 hid them and silently orphaned rules filing
 *     into them), roots tagged so the AI is never offered them.
 */
export const WELL_KNOWN_FOLDER_ROLES = [
  'deleteditems',
  'drafts',
  'outbox',
  'junkemail',
  'sentitems',
] as const satisfies readonly WellKnownFolderRole[]
const OMITTED_ROLES: ReadonlySet<WellKnownFolderRole> = new Set(['deleteditems', 'drafts', 'outbox'])
const AI_HIDDEN_ROLES: ReadonlySet<WellKnownFolderRole> = new Set(['junkemail', 'sentitems'])

function isOmitted(role: WellKnownFolderRole | undefined): boolean {
  return role !== undefined && OMITTED_ROLES.has(role)
}

// Distinguished-folder ids never change for a mailbox, but the signed-in
// account can (OWA account switch), so the memo is short-lived.
const WELL_KNOWN_TTL_MS = 10 * 60_000

/**
 * Per-request deadline, headers AND body (2026-11-06). fetch has none of its
 * own; with the 4-slot semaphore, four stalled requests blocked every Outlook
 * call for good. A timed-out GET retries like any transient failure; a timed-
 * out POST / DELETE / PATCH is `uncertain` (it may have committed) and is
 * never retried.
 */
export const OUTLOOK_REQUEST_TIMEOUT_MS = 60_000

// F6 (2026-06-03): upper bound on how long a single Retry-After can park
// the batch. Outlook normally returns small values, but a misbehaving
// proxy / throttle returning e.g. `Retry-After: 86400` would otherwise
// freeze the whole batch for 24h (abortable, but invisible — looks hung).
// 60s is generous for genuine throttling while keeping a runaway value
// from stalling the user indefinitely.
const MAX_RETRY_AFTER_MS = 60_000

/** Compute backoff wait, honouring Retry-After but clamped to a sane range. */
function computeRetryWaitMs(retryAfterHeader: string | null, attempt: number): number {
  // Audit (retry clamp): only treat a POSITIVE parsed Retry-After as valid.
  // A header like '-30' parses to a finite -30, which (without this guard)
  // bypassed the exponential fallback AND survived the upper clamp as a
  // negative wait → setTimeout fires next-tick → backoff collapses to an
  // immediate retry burst against a server that just asked us to slow down.
  const retryAfterSec = retryAfterHeader ? Number.parseInt(retryAfterHeader, 10) : NaN
  const base = Number.isFinite(retryAfterSec) && retryAfterSec > 0
    ? retryAfterSec * 1000
    : 1000 * Math.pow(3, attempt)
  // Clamp both ends: never negative, never longer than the max.
  return Math.min(Math.max(base, 0), MAX_RETRY_AFTER_MS)
}

// IMPORTANT (batch-3): do NOT add Body / UniqueBody here. This select feeds
// every list call (inbox scan, folder scans) — pulling full bodies would blow
// the response size 10-50× and invite 429s. Full body is fetched per-message,
// on demand, only for the AI-bucket subset via getMessageBody().
const MESSAGE_SELECT_DEFAULT = [
  'Id',
  'Subject',
  'BodyPreview',
  'ConversationId',
  'From',
  'ToRecipients',
  'CcRecipients',
  'ReceivedDateTime',
  'ParentFolderId',
  'IsRead',
  'HasAttachments',
  'Flag',
].join(',')

// Minimal HTML→text fallback used ONLY when the server ignores our
// Prefer:text header and still returns HTML. This is NOT a security
// sanitiser — the result is never rendered, only scanned for case numbers by
// regex — so a plain tag-strip + a few entity decodes is enough. (Self-writing
// a real sanitiser is unsafe in an MV3 worker with no DOMParser.)
function stripHtmlToText(html: string): string {
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\s+/g, ' ')
    .trim()
}

export class OutlookError extends Error {
  constructor(
    public readonly status: number,
    message: string,
    public readonly retryAfterSec?: number,
    /**
     * Set on errors where the underlying server-side state is genuinely
     * uncertain — typically a non-idempotent POST/DELETE/PATCH whose
     * request reached the server but whose response was lost (network
     * blip / TLS reset / 5xx). The operation MAY have succeeded; we
     * don't know. Callers (e.g. executeItem) can treat these
     * differently from "definitely failed" errors — e.g. don't surface
     * as a fatal retry candidate, since retry would either hit the same
     * blip or 404 (already moved).
     */
    public readonly uncertain: boolean = false,
    /**
     * Set when the caller's AbortSignal fired while a non-idempotent request
     * was already on the wire (or its 2xx body was being read) (2026-11-05).
     * Always paired with `uncertain` — the server may have committed the
     * move — so execute renders a soft "may have completed" row instead of
     * a plain 'cancelled' that lies when the move went through.
     */
    public readonly aborted: boolean = false,
  ) {
    super(message)
    this.name = 'OutlookError'
  }
}

/**
 * A request that never went out because no OWA token could be had — no
 * Outlook tab, signed out, an unreachable tab (2026-11-06). Still an
 * OutlookError(401) for every existing check; `tokenCode` is the token
 * layer's code (NOT_ON_OWA, NO_TOKEN, TAB_UNREACHABLE, …) when it gave one,
 * so a handler can answer "open Outlook first" without logging an error.
 */
export class OutlookTokenUnavailableError extends OutlookError {
  constructor(
    public readonly reason: string,
    public readonly tokenCode?: string,
  ) {
    super(401, `Token unavailable: ${reason}`)
    this.name = 'OutlookTokenUnavailableError'
  }
}

/** Response shape of a move (the moved message; Outlook assigns a new Id). */
export type MovedMessage = { Id: string; ParentFolderId?: string }

/** Mailbox-wide message row used by the manual-filing harvester. */
export type ModifiedMessage = Pick<Email, 'Id' | 'Subject' | 'From' | 'ConversationId' | 'ParentFolderId'> & {
  LastModifiedDateTime?: string
}

type Paged<T> = { value?: T[]; '@odata.nextLink'?: string }

type AttemptResult<T> =
  | { kind: 'ok'; value: T }
  | { kind: 'network'; error: unknown }
  | {
      kind: 'http'
      status: number
      statusText: string
      retryAfter: string | null
      text: string
    }

/**
 * RFC 9110 idempotent methods — safe to retry on network errors / 5xx
 * because re-executing them has no additional server-side effect.
 *
 * POST / DELETE / PATCH are explicitly NOT idempotent: a successful
 * server-side execution whose response is lost would, on retry, either
 * (a) double-process (worst: duplicate folder created), or (b) hit 404
 * because the resource is gone (move sees old message id stale). The
 * `/move` 404 we hit in production traced exactly to this: first POST
 * succeeded server-side, response lost, retry → ErrorItemNotFound. Fix:
 * only retry non-idempotent on explicit 429 (server promised "didn't
 * process") — for everything else, bail with an uncertain error.
 *
 * Note: POST for new folders is still safe because Outlook returns 409
 * on duplicate name, which createChildFolder / createTopFolder recover
 * from explicitly. POST for move/copy and DELETE for messages are the
 * dangerous ones.
 */
function isIdempotentMethod(method: string): boolean {
  const m = method.toUpperCase()
  return m === 'GET' || m === 'HEAD' || m === 'PUT' || m === 'OPTIONS'
}

/**
 * What to throw when the caller's signal aborts mid-request (2026-11-05).
 * For idempotent methods an abort is a plain cancel. For POST/DELETE/PATCH
 * the request was already on the wire, so the server may have committed it —
 * surfacing a bare AbortError made execute mark a row 'cancelled' even when
 * the move had gone through. (Aborts BEFORE the fetch starts — token bucket,
 * semaphore wait, backoff sleep — still surface as AbortError: nothing was
 * sent.)
 */
function abortedInFlight(e: unknown, method: string, idempotent: boolean): unknown {
  if (idempotent) return e
  return new OutlookError(
    0,
    `Cancelled while ${method} was in flight — the operation may have completed server-side`,
    undefined,
    true,
    true,
  )
}

/**
 * Exchange rejects `$orderby` properties that are not also the LEADING
 * properties of `$filter` (InefficientFilter, 2026-11-05). True when the
 * filter's first comparison is on ReceivedDateTime, i.e. when the default
 * `ReceivedDateTime desc` sort may be combined with it.
 */
export function filterLeadsWithReceivedDateTime(filter: string): boolean {
  return /^\s*\(*\s*ReceivedDateTime\b/i.test(filter)
}

export type TokenProvider = (opts?: { force?: boolean }) => Promise<string>
export type BaseUrlProvider = () => Promise<string> | string

export type OutlookApiOptions = {
  /** Injected for tests; production shares the module-level bucket. */
  bucket?: TokenBucket
  /** Injected for tests; production shares the module-level semaphore. */
  inflight?: Semaphore
  /** Injected for tests; production uses OUTLOOK_REQUEST_TIMEOUT_MS. */
  requestTimeoutMs?: number
}

export class OutlookApi {
  private readonly bucket: TokenBucket
  private readonly inflight: Semaphore
  private readonly requestTimeoutMs: number
  private wellKnownMemo: { at: number; roots: WellKnownFolderRoot[] } | null = null

  constructor(
    private readonly getToken: TokenProvider,
    // Optional base-URL provider — lets the service-worker resolve the
    // REST endpoint from the actual OWA tab's origin (office.com vs
    // cloud.microsoft etc.) so token audience matches the call host.
    // Defaults to DEFAULT_BASE (outlook.office.com) when omitted.
    private readonly getBaseUrl?: BaseUrlProvider,
    opts: OutlookApiOptions = {},
  ) {
    this.bucket = opts.bucket ?? defaultBucket
    this.inflight = opts.inflight ?? defaultInflight
    this.requestTimeoutMs = opts.requestTimeoutMs ?? OUTLOOK_REQUEST_TIMEOUT_MS
  }

  private async resolveBase(): Promise<string> {
    if (this.getBaseUrl) {
      try {
        return await this.getBaseUrl()
      } catch {
        /* fall back to default */
      }
    }
    return DEFAULT_BASE
  }

  /**
   * `path` is normally relative to the REST base; an absolute URL is an
   * `@odata.nextLink` (2026-11-05) and is only followed when it points at an
   * Outlook REST host — the bearer token must never leave for another origin.
   */
  private async resolveUrl(path: string): Promise<string> {
    const base = await this.resolveBase()
    if (!/^https?:\/\//i.test(path)) return `${base}${path}`
    let u: URL
    try {
      u = new URL(path)
    } catch {
      throw new OutlookError(0, `Refusing to follow a malformed nextLink: ${path.slice(0, 120)}`)
    }
    let baseOrigin = ''
    try {
      baseOrigin = new URL(base).origin
    } catch {
      /* ignore */
    }
    const allowedHost = u.origin === baseOrigin || KNOWN_REST_ORIGINS.has(u.origin)
    if (u.protocol !== 'https:' || !allowedHost || !u.pathname.toLowerCase().startsWith('/api/v2.0/')) {
      throw new OutlookError(0, `Refusing to follow a nextLink to an unexpected endpoint: ${u.origin}`)
    }
    return u.toString()
  }

  /**
   * One network round trip, holding a concurrency slot ONLY while the
   * request + response body are on the wire (2026-11-05) — never while we
   * sleep for backoff or wait for a token, so a parked retry can't starve the
   * other three slots.
   *
   * Bounded by requestTimeoutMs (2026-11-06): a timeout is reported like a
   * network error (GET → retried by request(); POST / DELETE / PATCH →
   * uncertain), while the CALLER's abort keeps its own semantics
   * (AbortError, or uncertain + aborted mid-POST). The slot is released on
   * every path.
   */
  private async attempt<T>(
    method: string,
    url: string,
    token: string,
    body: unknown,
    signal: AbortSignal | undefined,
    extraHeaders: Record<string, string> | undefined,
    idempotent: boolean,
  ): Promise<AttemptResult<T>> {
    const release = await this.inflight.acquire(signal)
    const deadline = new AbortController()
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      deadline.abort(new DOMException('Outlook request timed out', 'TimeoutError'))
    }, this.requestTimeoutMs)
    const linked = linkAbortSignals([signal, deadline.signal])
    const timeoutError = () =>
      new Error(`Outlook ${method} timed out after ${Math.round(this.requestTimeoutMs / 1000)}s`)
    try {
      let resp: Response
      try {
        resp = await fetch(url, {
          method,
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            Accept: 'application/json',
            ...extraHeaders,
          },
          body: body === undefined ? undefined : JSON.stringify(body),
          signal: linked.signal,
        })
      } catch (e) {
        if (signal?.aborted) throw abortedInFlight(e, method, idempotent)
        if (timedOut) return { kind: 'network', error: timeoutError() }
        if (isAbortError(e)) throw abortedInFlight(e, method, idempotent)
        return { kind: 'network', error: e }
      }
      if (resp.ok) {
        if (resp.status === 204) return { kind: 'ok', value: undefined as T }
        try {
          return { kind: 'ok', value: (await resp.json()) as T }
        } catch (e) {
          // 2xx already arrived: a non-idempotent op DID commit, we just lost
          // its body (e.g. the new message Id) — uncertain, not cancelled.
          if (signal?.aborted) throw abortedInFlight(e, method, idempotent)
          if (timedOut) {
            if (idempotent) return { kind: 'network', error: timeoutError() }
            throw new OutlookError(
              0,
              `${timeoutError().message} while reading the response — the operation completed server-side but its result was lost`,
              undefined,
              true,
            )
          }
          if (isAbortError(e)) throw abortedInFlight(e, method, idempotent)
          throw e
        }
      }
      const text = await resp.text().catch(() => '')
      return {
        kind: 'http',
        status: resp.status,
        statusText: resp.statusText,
        retryAfter: resp.headers.get('Retry-After'),
        text,
      }
    } finally {
      clearTimeout(timer)
      linked.dispose()
      release()
    }
  }

  private async request<T>(
    method: string,
    path: string,
    body?: unknown,
    signal?: AbortSignal,
    extraHeaders?: Record<string, string>,
    /**
     * `unthrottled` (2026-11-06): skip the rate limiter's FIFO queue — for a
     * single, rare, latency-bounded call (the mailbox identity lookup) that
     * would otherwise wait behind a whole fan-out. The concurrency cap still
     * applies; retries still back off.
     */
    opts: { unthrottled?: boolean } = {},
  ): Promise<T> {
    let forceToken = false
    let triedRefresh = false
    let backoffAttempt = 0
    const idempotent = isIdempotentMethod(method)

    while (true) {
      if (opts.unthrottled) signal?.throwIfAborted()
      else await this.bucket.take(1, signal)

      let token: string
      try {
        // The caller's signal bounds the token wait too (2026-11-06 r2): the
        // forced refresh after a 401 asks the OWA tab (up to ~10 s each), and
        // a short deadline like the identity lookup's 3 s used to not cover
        // it. Nothing has been sent yet, so this is a plain abort.
        token = await abortable(this.getToken({ force: forceToken }), signal)
      } catch (e) {
        if (signal?.aborted) throw signal.reason ?? new DOMException('Aborted', 'AbortError')
        const code = (e as { code?: unknown } | null)?.code
        throw new OutlookTokenUnavailableError(
          e instanceof Error ? e.message : String(e),
          typeof code === 'string' ? code : undefined,
        )
      }
      forceToken = false

      const url = await this.resolveUrl(path)
      const r = await this.attempt<T>(method, url, token, body, signal, extraHeaders, idempotent)
      if (r.kind === 'ok') return r.value

      if (r.kind === 'network') {
        const e = r.error
        if (idempotent && backoffAttempt < 3) {
          // Safe to retry — re-running a GET/HEAD/PUT has no
          // additional server effect.
          await abortableSleep(1000 * Math.pow(3, backoffAttempt), signal)
          backoffAttempt++
          continue
        }
        // Non-idempotent (POST/DELETE/PATCH): the request may have
        // landed and succeeded server-side before the network blip;
        // retrying would either double-process or hit 404 on a now-
        // stale resource id. Bail with uncertain=true so the caller can
        // surface it as "may have succeeded" instead of marking the
        // item as a hard failure.
        const msg = idempotent
          ? `Network error after retries: ${e instanceof Error ? e.message : String(e)}`
          : `Network error on ${method} (non-idempotent; not retried — operation may have completed server-side): ${e instanceof Error ? e.message : String(e)}`
        throw new OutlookError(0, msg, undefined, !idempotent)
      }

      if (r.status === 401 && !triedRefresh) {
        triedRefresh = true
        forceToken = true
        continue
      }

      // 429 — server explicitly said "didn't process, try again later".
      // Safe to retry for ALL methods.
      if (r.status === 429 && backoffAttempt < 3) {
        const waitMs = computeRetryWaitMs(r.retryAfter, backoffAttempt)
        await abortableSleep(waitMs, signal)
        backoffAttempt++
        continue
      }

      // 5xx — server crashed mid-request. For idempotent methods,
      // retry. For non-idempotent, the operation MAY have committed
      // server-side before the crash (Outlook is not transactional
      // about response delivery), so retrying could hit 404 on a
      // now-stale id. Bail with uncertain=true.
      if (r.status >= 500 && backoffAttempt < 3 && idempotent) {
        const waitMs = computeRetryWaitMs(r.retryAfter, backoffAttempt)
        await abortableSleep(waitMs, signal)
        backoffAttempt++
        continue
      }

      const uncertain = !idempotent && r.status >= 500
      const shownPath = /^https?:\/\//i.test(path) ? `${new URL(url).pathname}…` : path
      throw new OutlookError(
        r.status,
        `Outlook REST ${method} ${shownPath} → ${r.status} ${r.statusText} ${r.text.slice(0, 300)}`,
        undefined,
        uncertain,
      )
    }
  }

  /**
   * GET a collection and follow `@odata.nextLink` (2026-11-05). `complete`
   * is false when the page cap stopped us (or a looping nextLink was
   * detected) — the caller decides whether a partial listing is acceptable.
   */
  private async collectPages<T>(
    firstPath: string,
    opts: { maxPages: number; signal?: AbortSignal },
  ): Promise<{ items: T[]; complete: boolean; pages: number }> {
    const items: T[] = []
    const seen = new Set<string>()
    let next: string | undefined = firstPath
    let pages = 0
    while (next !== undefined && pages < opts.maxPages) {
      if (seen.has(next)) break // server echoed a nextLink we already fetched
      seen.add(next)
      const current: string = next
      const page: Paged<T> = await this.request<Paged<T>>('GET', current, undefined, opts.signal)
      pages++
      if (Array.isArray(page?.value)) items.push(...page.value)
      const link = page?.['@odata.nextLink']
      if (typeof link !== 'string' || link.length === 0) {
        next = undefined
      } else if (/^https?:\/\//i.test(link)) {
        next = link
      } else {
        // OData nextLinks are absolute; resolve a relative one defensively
        // against the page it came from (resolveUrl still vets the host).
        next = new URL(link, await this.resolveUrl(current)).toString()
      }
    }
    return { items, complete: next === undefined, pages }
  }

  private async listAllFolders(firstPath: string, what: string, signal?: AbortSignal): Promise<MailFolder[]> {
    const r = await this.collectPages<MailFolder>(firstPath, {
      maxPages: FOLDER_LIST_MAX_PAGES,
      signal,
    })
    if (!r.complete) {
      // Never hand back a silently truncated folder list — that is the
      // orphaned-rules / failed-409-recovery bug this paging exists to fix.
      throw new OutlookError(0, `資料夾清單分頁未完成(${what},已讀 ${r.pages} 頁)`)
    }
    return r.items
  }

  // ---- Folders --------------------------------------------------------------

  /** All top-level folders (follows nextLink — 2026-11-05). */
  async listTopFolders(signal?: AbortSignal): Promise<MailFolder[]> {
    return this.listAllFolders('/me/MailFolders?$top=100', 'top-level', signal)
  }

  /**
   * All child folders of `parentId` (follows nextLink — 2026-11-05). The old
   * single `$top=200` page dropped case folders 201+ under a busy parent:
   * their rules got orphaned, the AI never saw them, and new_folder's 409
   * recovery could never find the existing folder.
   */
  async listChildFolders(parentId: string, signal?: AbortSignal): Promise<MailFolder[]> {
    return this.listAllFolders(`/me/MailFolders/${parentId}/childFolders?$top=200`, parentId, signal)
  }

  /** One folder's identity + parent (used to verify move destinations). */
  async getFolder(folderId: string, signal?: AbortSignal): Promise<MailFolder> {
    return this.request<MailFolder>(
      'GET',
      `/me/MailFolders/${folderId}?$select=Id,DisplayName,ParentFolderId,ChildFolderCount`,
      undefined,
      signal,
    )
  }

  /**
   * Resolve the ids of WELL_KNOWN_FOLDER_ROLES (name → id). Defensive by
   * design: a lookup that fails is simply absent from the map — callers must
   * treat a missing name as "unknown", never as an error. Memoized briefly,
   * and only when every lookup produced a definitive answer.
   */
  async getWellKnownFolderIds(): Promise<Map<string, string>> {
    return new Map((await this.resolveWellKnownFolders()).roots.map((r) => [r.role, r.id]))
  }

  /**
   * Every well-known root that resolved, and whether every lookup gave a
   * definitive answer (2026-11-06 r2). Partial on transient failures — use
   * what resolved, never all-or-nothing (a single 429 used to leave the whole
   * trash subtree listed in a tree cached fresh for a day).
   */
  async resolveWellKnownFolders(): Promise<WellKnownFolders> {
    const { roots, failed } = await this.lookupWellKnownRoles()
    return { roots, complete: failed.size === 0 }
  }

  /**
   * One distinguished folder's id (2026-11-06 r3) — undefined when the
   * mailbox definitively has none; throws on a transient failure. Execute
   * retries just the roles its cached tree couldn't resolve.
   */
  async getWellKnownFolderId(role: WellKnownFolderRole, signal?: AbortSignal): Promise<string | undefined> {
    return (await this.lookupWellKnownRole(role, signal))?.id
  }

  private async lookupWellKnownRole(
    role: WellKnownFolderRole,
    signal?: AbortSignal,
  ): Promise<WellKnownFolderRoot | undefined> {
    try {
      const f = await this.request<{ Id?: string; DisplayName?: string }>(
        'GET',
        `/me/MailFolders/${role}?$select=Id,DisplayName`,
        undefined,
        signal,
      )
      if (typeof f?.Id !== 'string' || !f.Id) return undefined
      return {
        id: f.Id,
        role,
        ...(typeof f.DisplayName === 'string' && f.DisplayName ? { displayName: f.DisplayName } : {}),
      }
    } catch (e) {
      // A 4xx (e.g. this mailbox has no such folder / name unsupported) is a
      // stable answer; network / 5xx / token trouble is not.
      const transient =
        !(e instanceof OutlookError) ||
        e.status === 0 ||
        e.status === 401 ||
        e.status === 429 ||
        e.status >= 500
      if (transient) throw e
      return undefined
    }
  }

  /** Every role's lookup: the roots that resolved, and the TRANSIENT failures. */
  private async lookupWellKnownRoles(): Promise<{
    roots: WellKnownFolderRoot[]
    failed: Map<WellKnownFolderRole, unknown>
  }> {
    const memo = this.wellKnownMemo
    if (memo && Date.now() - memo.at < WELL_KNOWN_TTL_MS) {
      return { roots: memo.roots.map((r) => ({ ...r })), failed: new Map() }
    }
    const found = new Map<WellKnownFolderRole, WellKnownFolderRoot>()
    const failed = new Map<WellKnownFolderRole, unknown>()
    await Promise.all(
      WELL_KNOWN_FOLDER_ROLES.map(async (role) => {
        try {
          const root = await this.lookupWellKnownRole(role)
          if (root) found.set(role, root)
        } catch (e) {
          failed.set(role, e)
        }
      }),
    )
    // Stable order (role order), whatever order the lookups finished in.
    const roots = WELL_KNOWN_FOLDER_ROLES.flatMap((role) => {
      const r = found.get(role)
      return r ? [r] : []
    })
    if (failed.size === 0) this.wellKnownMemo = { at: Date.now(), roots: roots.map((r) => ({ ...r })) }
    return { roots, failed }
  }

  /**
   * Recursively walks the folder tree.
   * Sibling folders fetch in parallel; the shared TokenBucket spaces the
   * requests and the Semaphore caps how many are in flight.
   */
  async fetchFolderTree(): Promise<MailFolderNode[]> {
    return (await this.fetchFolderTreeWithMeta()).tree
  }

  /**
   * Like fetchFolderTree, plus the well-known roots it was built with
   * (2026-11-06 r2) — persisted with the folder cache: reconcile keeps a rule
   * on Deleted Items itself, execute recognises non-filing targets without
   * re-resolving. Deleted Items, Drafts and Outbox are left out with their
   * subtrees UNLISTED; Junk Email / Sent Items stay, roots tagged.
   *
   * A role whose lookup failed transiently keeps `previous`'s root
   * (2026-11-06 r3) — distinguished ids are stable for a mailbox — as long
   * as that folder is still at the top level (another account's id never
   * is). Only a role still unknown after that is listed like any folder,
   * and `wellKnown.complete` is false: the cache is refetched soon and
   * reconcile doesn't act on it. If EVERY lookup failed (no token, Outlook
   * unreachable), no tree is built at all — nothing in it could be
   * recognised, yet the same preflight would hand it to the AI — and the
   * failure is thrown (an OutlookTokenUnavailableError when that was the
   * cause), so the caller serves its cached tree instead.
   */
  async fetchFolderTreeWithMeta(
    opts: { previous?: WellKnownFolders } = {},
  ): Promise<{ tree: MailFolderNode[]; wellKnown: WellKnownFolders }> {
    const lookup = await this.lookupWellKnownRoles()
    if (lookup.failed.size === WELL_KNOWN_FOLDER_ROLES.length) {
      const errors = [...lookup.failed.values()]
      throw errors.find((e) => e instanceof OutlookTokenUnavailableError) ?? errors[0]
    }
    const top = await this.listTopFolders()
    const topIds = new Set(top.map((f) => f.Id))
    const roots = [...lookup.roots]
    let complete = true
    for (const role of lookup.failed.keys()) {
      const sticky = opts.previous?.roots.find((r) => r.role === role)
      if (sticky && topIds.has(sticky.id)) roots.push({ ...sticky })
      else complete = false
    }
    roots.sort((a, b) => WELL_KNOWN_FOLDER_ROLES.indexOf(a.role) - WELL_KNOWN_FOLDER_ROLES.indexOf(b.role))
    const roleById = new Map(roots.map((r) => [r.id, r.role] as const))
    const tree = await Promise.all(
      top.filter((f) => !isOmitted(roleById.get(f.Id))).map((f) => this.buildNode(f, '', roleById)),
    )
    return { tree, wellKnown: { roots, complete } }
  }

  private async buildNode(
    folder: MailFolder,
    parentPath: string,
    roleById: ReadonlyMap<string, WellKnownFolderRole>,
  ): Promise<MailFolderNode> {
    const path = joinFolderPath(parentPath || undefined, folder.DisplayName)
    const role = roleById.get(folder.Id)
    const node: MailFolderNode = {
      id: folder.Id,
      displayName: folder.DisplayName,
      parentFolderId: folder.ParentFolderId,
      path,
      children: [],
      ...(role && AI_HIDDEN_ROLES.has(role) ? { wellKnown: role as AiHiddenFolderRole } : {}),
    }
    if ((folder.ChildFolderCount ?? 0) > 0) {
      // Distinguished folders sit at the top level; filtered at every level
      // anyway (cheap, and never lists an omitted subtree).
      const children = (await this.listChildFolders(folder.Id)).filter((c) => !isOmitted(roleById.get(c.Id)))
      node.children = await Promise.all(children.map((c) => this.buildNode(c, path, roleById)))
    }
    return node
  }

  async createTopFolder(displayName: string, signal?: AbortSignal): Promise<MailFolder> {
    return this.request<MailFolder>('POST', '/me/MailFolders', { DisplayName: displayName }, signal)
  }

  async createChildFolder(parentId: string, displayName: string, signal?: AbortSignal): Promise<MailFolder> {
    return this.request<MailFolder>(
      'POST',
      `/me/MailFolders/${parentId}/childFolders`,
      { DisplayName: displayName },
      signal,
    )
  }

  // ---- Messages -------------------------------------------------------------

  async listInboxMessages(
    opts: { top?: number; skip?: number; select?: string; signal?: AbortSignal } = {},
  ): Promise<Email[]> {
    return this.listFolderMessages('inbox', opts)
  }

  async listFolderMessages(
    folderIdOrWellKnown: string,
    opts: {
      top?: number
      skip?: number
      select?: string
      /**
       * Optional OData $filter expression — server-side pre-filter so
       * we don't pull (and pay for) emails we don't need. Examples:
       *   - `endswith(from/emailAddress/address, '@company-b.example')`
       *   - `from/emailAddress/address eq 'support@company-b.example'`
       *
       * Kept generic — the auto-conflict-resolver that previously
       * consumed this was removed in 2026-05-27 (learning-time conflict
       * prevention replaced it); the parameter stays available for
       * future scoped-fetch needs.
       */
      filter?: string
      /**
       * Optional OData $orderby. Without a filter it defaults to
       * `ReceivedDateTime desc` (unchanged for the inbox / folder scans). With
       * a filter, the default is added ONLY when the filter leads with
       * ReceivedDateTime (2026-11-05): Exchange answers 400 InefficientFilter
       * when an orderby property isn't also the filter's leading property, so
       * the old always-on default made every `From/EmailAddress/Address`
       * server filter (autoUpgradeConflictRules) fail and fall back to the
       * client-side scan.
       */
      orderby?: string
      signal?: AbortSignal
    } = {},
  ): Promise<Email[]> {
    const top = opts.top ?? 50
    const skip = opts.skip ?? 0
    const select = opts.select ?? MESSAGE_SELECT_DEFAULT
    const parts = [`$top=${top}`, `$skip=${skip}`, `$select=${select}`]
    const orderby =
      opts.orderby ??
      (!opts.filter || filterLeadsWithReceivedDateTime(opts.filter) ? 'ReceivedDateTime desc' : undefined)
    if (orderby) parts.push(`$orderby=${encodeURIComponent(orderby)}`)
    if (opts.filter) {
      parts.push(`$filter=${encodeURIComponent(opts.filter)}`)
    }
    const r = await this.request<{ value: Email[] }>(
      'GET',
      `/me/MailFolders/${folderIdOrWellKnown}/messages?${parts.join('&')}`,
      undefined,
      opts.signal,
    )
    return r.value
  }

  /**
   * Fetch one message's body as plain text (batch-3). Uses UniqueBody so
   * quoted reply history is dropped — a case number quoted from a prior email
   * must not pollute this message's body case detection. `Prefer:text` makes
   * the server return plain text (no unsafe client-side HTML handling). The
   * stripHtmlToText fallback only fires if the server ignores Prefer. Truncated
   * defensively so a pathological body can't blow the response we hold in
   * memory; the prompt layer truncates again to its own budget.
   */
  async getMessageBody(messageId: string, signal?: AbortSignal): Promise<string> {
    const r = await this.request<{ UniqueBody?: { ContentType?: string; Content?: string } }>(
      'GET',
      `/me/messages/${messageId}?$select=UniqueBody`,
      undefined,
      signal,
      { Prefer: 'outlook.body-content-type="text"' },
    )
    const content = r.UniqueBody?.Content ?? ''
    const type = (r.UniqueBody?.ContentType ?? '').toLowerCase()
    const text = type === 'html' ? stripHtmlToText(content) : content
    return text.length > 4000 ? text.slice(0, 4000) : text
  }

  /**
   * Mailbox-wide "modified since" listing for the manual-filing harvester
   * (2026-11-05): ONE query over `/me/messages` instead of a `$top=10` probe
   * per leaf folder, paged via nextLink up to `maxPages`. The orderby
   * property is also the filter's leading property (Exchange rule).
   * `complete` is false when the page cap stopped the walk.
   *
   * OLDEST first (2026-11-06): with `desc`, a window holding more than
   * maxPages × pageSize modified messages (read-state changes count) was
   * never finished — every run re-read the newest pages and the older part
   * was never seen. Ascending, a capped run has covered a prefix of the
   * window, so the caller can resume from its last timestamp.
   *
   * KEYSET pages (2026-11-06 r2) instead of the server's `$skip` nextLink: a
   * row modified mid-scan jumps to the end of the order, which shifted every
   * later offset and skipped a row at the next page boundary. Each page after
   * the first asks for `LastModifiedDateTime ge <the previous page's last
   * timestamp>` (the server's own string, so precision always matches); rows
   * re-read at that boundary are de-duplicated by Id, the latest copy winning
   * (it carries the folder the mail is in NOW). Every page makes progress:
   * when a whole page shares one timestamp (a bulk change bigger than a
   * page), `$skip` steps through that timestamp's rows only.
   */
  async listMessagesModifiedSince(
    sinceIso: string,
    opts: { pageSize?: number; maxPages?: number; signal?: AbortSignal } = {},
  ): Promise<{ messages: ModifiedMessage[]; complete: boolean; pages: number }> {
    const pageSize = opts.pageSize ?? 100
    const maxPages = opts.maxPages ?? 20
    const byId = new Map<string, ModifiedMessage>()
    let bound: { op: 'gt' | 'ge'; at: string } = { op: 'gt', at: sinceIso }
    // Rows of the bound's own timestamp already read (only while a page
    // can't get past that timestamp).
    let skip = 0
    let pages = 0
    let complete = false
    while (pages < maxPages) {
      const parts = [
        `$filter=${encodeURIComponent(`LastModifiedDateTime ${bound.op} ${bound.at}`)}`,
        `$orderby=${encodeURIComponent('LastModifiedDateTime asc')}`,
        '$select=Id,Subject,From,ConversationId,LastModifiedDateTime,ParentFolderId',
        `$top=${pageSize}`,
        ...(skip > 0 ? [`$skip=${skip}`] : []),
      ]
      const page = await this.request<Paged<ModifiedMessage>>(
        'GET',
        `/me/messages?${parts.join('&')}`,
        undefined,
        opts.signal,
      )
      pages++
      const rows = Array.isArray(page?.value) ? page.value : []
      // Map.set keeps an existing key's position: first-seen order, latest copy.
      for (const m of rows) if (m?.Id) byId.set(m.Id, m)
      if (rows.length < pageSize) {
        complete = true
        break
      }
      const lastAt = rows[rows.length - 1]?.LastModifiedDateTime
      if (typeof lastAt !== 'string' || !lastAt) break // no keyset to continue from — incomplete
      if (bound.op === 'gt' || lastAt !== bound.at) {
        // Got past the bound: the next page starts at this page's last
        // timestamp (its rows there are read again and de-duplicated) —
        // unless the whole page WAS that timestamp, which must be skipped.
        const atLast = rows.filter((m) => m?.LastModifiedDateTime === lastAt).length
        bound = { op: 'ge', at: lastAt }
        skip = atLast === rows.length ? atLast : 0
      } else {
        // The whole page still sat on the bound's timestamp.
        skip += rows.length
      }
    }
    return { messages: [...byId.values()], complete, pages }
  }

  async moveMessage(messageId: string, destinationFolderId: string, signal?: AbortSignal): Promise<MovedMessage> {
    return this.request<MovedMessage>(
      'POST',
      `/me/messages/${messageId}/move`,
      { DestinationId: destinationFolderId },
      signal,
    )
  }

  /**
   * "Delete" = move into Deleted Items (2026-11-05). `DELETE /me/messages/{id}`
   * lands in Recoverable Items — invisible in OWA's 已刪除的郵件, which is
   * what the popup promises — and can't be reversed by 撤回. Outlook REST
   * v2.0 accepts the well-known name as DestinationId; if a host rejects it
   * (400 — the move was refused outright, nothing happened, so one retry
   * cannot double-move), retry once with the resolved folder id.
   */
  async moveToDeletedItems(messageId: string, signal?: AbortSignal): Promise<MovedMessage> {
    try {
      return await this.moveMessage(messageId, 'DeletedItems', signal)
    } catch (e) {
      if (e instanceof OutlookError && e.status === 400) {
        const id = (await this.getWellKnownFolderIds()).get('deleteditems')
        if (id) return this.moveMessage(messageId, id, signal)
      }
      throw e
    }
  }

  /**
   * Hard delete (lands in Recoverable Items). No longer used by execute —
   * see moveToDeletedItems — kept for API completeness.
   */
  async deleteMessage(messageId: string, signal?: AbortSignal): Promise<void> {
    await this.request<void>('DELETE', `/me/messages/${messageId}`, undefined, signal)
  }

  /**
   * Returns the signed-in user's primary email (e.g. for inferring firm
   * domain). `unthrottled` skips the rate limiter's queue (2026-11-06) — a
   * short timeout on this call used to expire while it merely queued.
   */
  async getMyEmail(signal?: AbortSignal, opts: { unthrottled?: boolean } = {}): Promise<string> {
    const r = await this.request<{ EmailAddress?: string }>(
      'GET',
      '/me',
      undefined,
      signal,
      undefined,
      opts,
    )
    return typeof r.EmailAddress === 'string' ? r.EmailAddress : ''
  }

  /**
   * One page of the inbox listing plus the server's continuation link
   * (2026-11-05, inbox starvation). Same $select / $orderby as
   * listInboxMessages; pass the returned `nextLink` back to get the next page
   * (it already encodes $top/$skip/$select, so `top`/`select` only shape the
   * first request).
   *
   * The link is re-based onto OUR resolved REST base: only its path + query
   * are kept, so the bearer token can never be sent to whatever host the
   * link happens to name.
   */
  async listInboxMessagesPage(opts: {
    top: number
    nextLink?: string | null
    select?: string
    signal?: AbortSignal
  }): Promise<{ messages: Email[]; nextLink: string | null }> {
    let path: string
    if (opts.nextLink) {
      let rel: string | null = null
      try {
        const u = new URL(opts.nextLink)
        const idx = u.pathname.toLowerCase().indexOf('/api/v2.0/')
        if (idx >= 0) rel = u.pathname.slice(idx + '/api/v2.0'.length) + u.search
      } catch {
        rel = null
      }
      if (!rel) {
        throw new OutlookError(0, `Unexpected @odata.nextLink: ${opts.nextLink.slice(0, 200)}`)
      }
      path = rel
    } else {
      const parts = [
        `$top=${opts.top}`,
        `$skip=0`,
        `$select=${opts.select ?? MESSAGE_SELECT_DEFAULT}`,
        `$orderby=${encodeURIComponent('ReceivedDateTime desc')}`,
      ]
      path = `/me/MailFolders/inbox/messages?${parts.join('&')}`
    }
    const r = await this.request<{ value?: Email[]; '@odata.nextLink'?: unknown }>(
      'GET',
      path,
      undefined,
      opts.signal,
    )
    const next = r['@odata.nextLink']
    return {
      messages: Array.isArray(r.value) ? r.value : [],
      nextLink: typeof next === 'string' && next.length > 0 ? next : null,
    }
  }
}

// ---- Helpers for path lookup -----------------------------------------------

/**
 * Path encoding strategy: Outlook allows DisplayName to contain '/', but our
 * MailFolderNode.path uses '/' as the segment separator. Without escaping, a
 * folder named "客戶/A" under "案件" computes path "案件/客戶/A" —
 * indistinguishable from a real 3-level structure. findFolderByPath would
 * pick whichever the tree walk hits first (non-deterministic).
 *
 * We replace '/' inside any segment (DisplayName) with U+FF0F FULLWIDTH
 * SOLIDUS before joining. The result is still a string, still uses '/' as
 * the separator, and U+FF0F looks nearly identical to '/' in CJK fonts so
 * UI display stays readable without an explicit decode step.
 *
 * Outlook API calls (createChildFolder etc.) continue to receive the RAW
 * displayName — only our path representation is encoded.
 *
 * The character U+FF0F was chosen because:
 *   - It's a single codepoint (cheap to encode/decode).
 *   - It's visually indistinguishable from '/' in most fonts.
 *   - It's extremely unlikely to appear in a real Outlook folder name.
 */
export const PATH_SEP = '/'
const SLASH_IN_NAME = '／'

export function encodeFolderName(name: string): string {
  return name.replace(/\//g, SLASH_IN_NAME)
}

export function joinFolderPath(parent: string | undefined, name: string): string {
  const encoded = encodeFolderName(name)
  return parent ? `${parent}${PATH_SEP}${encoded}` : encoded
}

export function flattenFolderTree(tree: MailFolderNode[]): MailFolderNode[] {
  const out: MailFolderNode[] = []
  const walk = (node: MailFolderNode) => {
    out.push(node)
    for (const child of node.children) walk(child)
  }
  for (const n of tree) walk(n)
  return out
}

export function findFolderByPath(tree: MailFolderNode[], path: string): MailFolderNode | undefined {
  return flattenFolderTree(tree).find((n) => n.path === path)
}

/**
 * The folders the AI may name (2026-11-06 r2): flattenFolderTree minus the
 * subtrees under a root tagged `wellKnown` (Junk Email / Sent Items). Those
 * stay valid targets for rules, the picker and execute, but an AI that could
 * see them proposed auto-junk moves that were then learned as rules.
 */
export function flattenAiVisibleFolders(tree: MailFolderNode[]): MailFolderNode[] {
  const out: MailFolderNode[] = []
  const walk = (node: MailFolderNode) => {
    if (node.wellKnown) return
    out.push(node)
    for (const child of node.children) walk(child)
  }
  for (const n of tree) walk(n)
  return out
}
