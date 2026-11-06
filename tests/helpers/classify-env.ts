// Shared fixture for SW-level classify-flow tests (2026-11-05).
//
// Seeds what the real service worker needs to run classifyPreflight →
// classifyAi → startExecute end to end — settings, a fresh folder cache, a
// cached OWA token — and routes fetch(): Outlook REST (inbox pages, bodies,
// moves) vs the Anthropic Messages API (a per-test handler). No real network.

import { vi } from 'vitest'
import type { Email, MailFolderNode } from '@/shared/types'
import { actionsJson, answerEvents, sseResponse, sseText } from './claude-sse'

export const FOLDER_ID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
export const FOLDER_PATH = '內部資料/工時審閱'

export function folderTree(): MailFolderNode[] {
  return [
    {
      id: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE0',
      displayName: '內部資料',
      path: '內部資料',
      children: [{ id: FOLDER_ID, displayName: '工時審閱', path: FOLDER_PATH, children: [] }],
    },
  ]
}

export async function seedClassifyEnv(opts: { model?: string; batchSize?: number } = {}): Promise<void> {
  await chrome.storage.local.set({
    settings: {
      claudeApiKey: 'sk-ant-test',
      claudeModel: opts.model ?? 'claude-sonnet-5-5',
      batchSize: opts.batchSize ?? 50,
      internalDomains: [],
    },
    // A current-format cache (2026-11-06 r3): one without `wellKnown` is
    // legacy and refetched on first use.
    folderCache: { updatedAt: new Date().toISOString(), tree: folderTree(), wellKnown: { roots: [], complete: true } },
  })
  await chrome.storage.session.set({
    cachedToken: { secret: 'owa-token', expiresOn: Date.now() / 1000 + 3600 },
  })
}

export function inboxMail(i: number, over: Partial<Email> = {}): Email {
  return {
    Id: `m${i}`,
    Subject: `subject ${i}`,
    BodyPreview: `preview ${i}`,
    From: { EmailAddress: { Address: `sender${i}@ai.example`, Name: `S${i}` } },
    ToRecipients: [],
    ReceivedDateTime: new Date(Date.UTC(2026, 10, 1, 0, 0, 0) - i * 60_000).toISOString(),
    Flag: { FlagStatus: 'NotFlagged' },
    ...over,
  } as Email
}

export type InboxPage = { messages: Email[]; nextLink?: string }

export type ClaudeHandler = (req: { body: Record<string, unknown>; signal?: AbortSignal | null }) => Promise<Response>

export type FetchRouter = {
  /** Queue of pages for inbox listings (first request, then nextLinks); consumed. */
  inboxPages: InboxPage[]
  claude: ClaudeHandler
  bodyStatus: number
  outlookUrls: string[]
  claudeBodies: Array<Record<string, unknown>>
  claudeSignals: Array<AbortSignal | null | undefined>
}

/** Number of emails a classify request asks about ("請分類以下 N 封郵件"). */
export function requestedEmailCount(body: Record<string, unknown>): number {
  const content = (body.messages as Array<{ content: Array<{ text?: string }> }>)[0]!.content
  const m = /請分類以下 (\d+) 封郵件/.exec(content[content.length - 1]!.text ?? '')
  return m ? Number(m[1]) : 0
}

/** A healthy answer that keeps every email in inbox (action skip, confident). */
export function answerAll(body: Record<string, unknown>, action = 'skip'): Response {
  const n = requestedEmailCount(body)
  return sseResponse(
    sseText(answerEvents(actionsJson(Array.from({ length: n }, (_, i) => ({ emailIndex: i, action }))))),
    { chunkBytes: 64 },
  )
}

/** A promise you resolve later, plus a fetch-like wait that honours abort. */
export function gate() {
  let open: () => void = () => {}
  const opened = new Promise<void>((r) => (open = r))
  const wait = (signal?: AbortSignal | null) =>
    new Promise<void>((resolve, reject) => {
      if (signal?.aborted) return reject(new DOMException('Aborted', 'AbortError'))
      signal?.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')))
      void opened.then(resolve)
    })
  return { open, wait }
}

export function installFetchRouter(): FetchRouter {
  const router: FetchRouter = {
    inboxPages: [],
    claude: async ({ body }) => answerAll(body),
    bodyStatus: 200,
    outlookUrls: [],
    claudeBodies: [],
    claudeSignals: [],
  }
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string, init: RequestInit = {}) => {
      if (url.startsWith('https://api.anthropic.com/')) {
        const body = JSON.parse(String(init.body)) as Record<string, unknown>
        router.claudeBodies.push(body)
        router.claudeSignals.push(init.signal)
        return router.claude({ body, signal: init.signal })
      }
      router.outlookUrls.push(url)
      const json = (v: unknown, status = 200) =>
        new Response(JSON.stringify(v), { status, headers: { 'content-type': 'application/json' } })
      if (url.includes('/me/MailFolders/inbox/messages')) {
        const page = router.inboxPages.shift() ?? { messages: [] }
        return json({
          value: page.messages,
          ...(page.nextLink ? { '@odata.nextLink': page.nextLink } : {}),
        })
      }
      if (url.includes('$select=UniqueBody')) {
        return router.bodyStatus === 200
          ? json({ UniqueBody: { ContentType: 'text', Content: '' } })
          : json({ error: { code: 'ErrorItemNotFound' } }, router.bodyStatus)
      }
      const move = /\/me\/messages\/([^/]+)\/move$/.exec(url)
      if (move && init.method === 'POST') return json({ Id: `moved-${move[1]}` })
      if (init.method === 'DELETE') return new Response(null, { status: 204 })
      return json({ value: [] })
    }),
  )
  return router
}

/**
 * Pretend every unmatched email's body was already fetched, so the loop makes
 * no Outlook body calls (keeps the module-level 10-token Outlook rate bucket
 * out of the tests' timing).
 */
export async function markBodiesFetched(): Promise<void> {
  const raw = await chrome.storage.session.get('preflightCache')
  const cache = raw.preflightCache as { unmatchedEmails: Email[] } | undefined
  if (!cache) return
  for (const m of cache.unmatchedEmails) m.bodyText = ''
  await chrome.storage.session.set({ preflightCache: cache })
}

export async function waitFor(cond: () => boolean | Promise<boolean>, timeoutMs = 3000): Promise<void> {
  const start = Date.now()
  while (!(await cond())) {
    if (Date.now() - start > timeoutMs) throw new Error('waitFor timed out')
    await new Promise((r) => setTimeout(r, 5))
  }
}
