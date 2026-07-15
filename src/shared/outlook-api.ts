// Outlook REST v2.0 wrapper.
//
// Endpoint base: https://outlook.office.com/api/v2.0
// Auth:          Bearer <MSAL access token from OWA localStorage>
//
// Retry policy:
//   - 401 once: refresh token then retry
//   - 429 / 5xx / network: exponential backoff (3^attempt sec), 3x max
//   - other 4xx: surface as OutlookError immediately
//
// Rate limiting via module-level TokenBucket (10 burst / 1 per sec).

import { TokenBucket, abortableSleep } from './throttle'
import type { Email, MailFolder, MailFolderNode } from './types'

function isAbortError(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError'
}

const DEFAULT_BASE = 'https://outlook.office.com/api/v2.0'
const bucket = new TokenBucket(10, 1)

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
  ) {
    super(message)
    this.name = 'OutlookError'
  }
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

export type TokenProvider = (opts?: { force?: boolean }) => Promise<string>
export type BaseUrlProvider = () => Promise<string> | string

export class OutlookApi {
  constructor(
    private readonly getToken: TokenProvider,
    // Optional base-URL provider — lets the service-worker resolve the
    // REST endpoint from the actual OWA tab's origin (office.com vs
    // cloud.microsoft etc.) so token audience matches the call host.
    // Defaults to DEFAULT_BASE (outlook.office.com) when omitted.
    private readonly getBaseUrl?: BaseUrlProvider,
  ) {}

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

  private async request<T>(
    method: string,
    path: string,
    body?: unknown,
    signal?: AbortSignal,
    extraHeaders?: Record<string, string>,
  ): Promise<T> {
    let forceToken = false
    let triedRefresh = false
    let backoffAttempt = 0

    while (true) {
      await bucket.take(1, signal)

      let token: string
      try {
        token = await this.getToken({ force: forceToken })
      } catch (e) {
        throw new OutlookError(401, `Token unavailable: ${e instanceof Error ? e.message : String(e)}`)
      }
      forceToken = false

      const base = await this.resolveBase()
      const idempotent = isIdempotentMethod(method)
      let resp: Response
      try {
        resp = await fetch(`${base}${path}`, {
          method,
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            Accept: 'application/json',
            ...extraHeaders,
          },
          body: body === undefined ? undefined : JSON.stringify(body),
          signal,
        })
      } catch (e) {
        if (isAbortError(e)) throw e // never retry on abort
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

      if (resp.ok) {
        if (resp.status === 204) return undefined as T
        return (await resp.json()) as T
      }

      if (resp.status === 401 && !triedRefresh) {
        triedRefresh = true
        forceToken = true
        continue
      }

      // 429 — server explicitly said "didn't process, try again later".
      // Safe to retry for ALL methods.
      if (resp.status === 429 && backoffAttempt < 3) {
        const waitMs = computeRetryWaitMs(resp.headers.get('Retry-After'), backoffAttempt)
        await abortableSleep(waitMs, signal)
        backoffAttempt++
        continue
      }

      // 5xx — server crashed mid-request. For idempotent methods,
      // retry. For non-idempotent, the operation MAY have committed
      // server-side before the crash (Outlook is not transactional
      // about response delivery), so retrying could hit 404 on a
      // now-stale id. Bail with uncertain=true.
      if (resp.status >= 500 && backoffAttempt < 3 && idempotent) {
        const waitMs = computeRetryWaitMs(resp.headers.get('Retry-After'), backoffAttempt)
        await abortableSleep(waitMs, signal)
        backoffAttempt++
        continue
      }

      const errText = await resp.text().catch(() => '')
      const uncertain = !idempotent && resp.status >= 500
      throw new OutlookError(
        resp.status,
        `Outlook REST ${method} ${path} → ${resp.status} ${resp.statusText} ${errText.slice(0, 300)}`,
        undefined,
        uncertain,
      )
    }
  }

  // ---- Folders --------------------------------------------------------------

  async listTopFolders(): Promise<MailFolder[]> {
    const r = await this.request<{ value: MailFolder[] }>(
      'GET',
      '/me/MailFolders?$top=100',
    )
    return r.value
  }

  async listChildFolders(parentId: string): Promise<MailFolder[]> {
    const r = await this.request<{ value: MailFolder[] }>(
      'GET',
      `/me/MailFolders/${parentId}/childFolders?$top=200`,
    )
    return r.value
  }

  /**
   * Recursively walks the folder tree.
   * Sibling folders fetch in parallel; the global TokenBucket serializes
   * actual network calls so this won't violate the rate limit.
   */
  async fetchFolderTree(): Promise<MailFolderNode[]> {
    const top = await this.listTopFolders()
    return Promise.all(top.map((f) => this.buildNode(f, '')))
  }

  private async buildNode(folder: MailFolder, parentPath: string): Promise<MailFolderNode> {
    const path = joinFolderPath(parentPath || undefined, folder.DisplayName)
    const node: MailFolderNode = {
      id: folder.Id,
      displayName: folder.DisplayName,
      parentFolderId: folder.ParentFolderId,
      path,
      children: [],
    }
    if ((folder.ChildFolderCount ?? 0) > 0) {
      const children = await this.listChildFolders(folder.Id)
      node.children = await Promise.all(children.map((c) => this.buildNode(c, path)))
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
      signal?: AbortSignal
    } = {},
  ): Promise<Email[]> {
    const top = opts.top ?? 50
    const skip = opts.skip ?? 0
    const select = opts.select ?? MESSAGE_SELECT_DEFAULT
    const parts = [
      `$top=${top}`,
      `$skip=${skip}`,
      `$select=${select}`,
      `$orderby=ReceivedDateTime%20desc`,
    ]
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

  async moveMessage(messageId: string, destinationFolderId: string, signal?: AbortSignal): Promise<{ Id: string }> {
    return this.request<{ Id: string }>(
      'POST',
      `/me/messages/${messageId}/move`,
      { DestinationId: destinationFolderId },
      signal,
    )
  }

  async deleteMessage(messageId: string, signal?: AbortSignal): Promise<void> {
    await this.request<void>('DELETE', `/me/messages/${messageId}`, undefined, signal)
  }

  /** Returns the signed-in user's primary email (e.g. for inferring firm domain). */
  async getMyEmail(signal?: AbortSignal): Promise<string> {
    const r = await this.request<{ EmailAddress?: string }>(
      'GET',
      '/me',
      undefined,
      signal,
    )
    return typeof r.EmailAddress === 'string' ? r.EmailAddress : ''
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
