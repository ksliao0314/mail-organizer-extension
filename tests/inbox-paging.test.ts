// Inbox starvation (2026-11-05, item 13).
//
// classifyPreflight and peekNextBatch fetched ONE page (top = batchSize +
// ghosts) and filtered kept / just-processed / flagged mail client-side with
// no paging: once the newest page was all kept or flagged, classify reported
// 0 emails although older mail existed. They now follow @odata.nextLink
// (OutlookApi.listInboxMessagesPage) until batchSize eligible or a cap.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { OutlookApi, OutlookError } from '@/shared/outlook-api'
import { addToSkipHistory } from '@/shared/storage'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import { inboxMail, installFetchRouter, seedClassifyEnv, type FetchRouter } from './helpers/classify-env'

const BASE = 'https://outlook.office.com/api/v2.0'

describe('OutlookApi.listInboxMessagesPage', () => {
  let calls: string[]
  beforeEach(() => {
    calls = []
  })
  afterEach(() => vi.unstubAllGlobals())

  function stub(bodies: unknown[]) {
    const queue = [...bodies]
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        calls.push(url)
        return new Response(JSON.stringify(queue.shift() ?? { value: [] }), { status: 200 })
      }),
    )
  }
  const api = () => new OutlookApi(async () => 'tok', () => BASE)

  it('first page: same $select / $orderby as listInboxMessages, and surfaces @odata.nextLink', async () => {
    stub([{ value: [inboxMail(0)], '@odata.nextLink': `${BASE}/me/MailFolders/inbox/messages?%24skip=10` }])
    const page = await api().listInboxMessagesPage({ top: 10 })
    expect(calls[0]).toContain('/me/MailFolders/inbox/messages?$top=10&$skip=0&$select=Id,Subject,BodyPreview')
    expect(calls[0]).toContain(`$orderby=${encodeURIComponent('ReceivedDateTime desc')}`)
    expect(page.messages.map((m) => m.Id)).toEqual(['m0'])
    expect(page.nextLink).toBe(`${BASE}/me/MailFolders/inbox/messages?%24skip=10`)
  })

  it('follows a nextLink re-based onto OUR base — the token never goes to the host the link names', async () => {
    stub([{ value: [] }])
    const page = await api().listInboxMessagesPage({
      top: 10,
      nextLink: "https://evil.example/api/v2.0/me/MailFolders('inbox')/messages?%24top=10&%24skip=10",
    })
    expect(calls[0]).toBe(`${BASE}/me/MailFolders('inbox')/messages?%24top=10&%24skip=10`)
    expect(page.nextLink).toBeNull()
  })

  it('rejects a nextLink it cannot re-base', async () => {
    stub([])
    await expect(
      api().listInboxMessagesPage({ top: 10, nextLink: 'not a url' }),
    ).rejects.toBeInstanceOf(OutlookError)
    expect(calls).toHaveLength(0)
  })
})

describe('preflight / peek page past kept and flagged mail', () => {
  let sw: SwHarness
  let router: FetchRouter
  beforeAll(async () => {
    sw = await loadServiceWorker()
  })
  beforeEach(async () => {
    await seedClassifyEnv()
    router = installFetchRouter()
  })
  afterEach(() => vi.unstubAllGlobals())

  const flagged = (i: number) => inboxMail(i, { Flag: { FlagStatus: 'Flagged' } })
  const link = (skip: number) => `${BASE}/me/MailFolders/inbox/messages?%24top=5&%24skip=${skip}`

  it('classifyPreflight finds the older mail below a page of kept + flagged mail', async () => {
    await addToSkipHistory(['m3', 'm4'])
    router.inboxPages = [
      { messages: [flagged(0), flagged(1), flagged(2), inboxMail(3), inboxMail(4)], nextLink: link(5) },
      { messages: [inboxMail(5), inboxMail(6), inboxMail(7)] },
    ]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 3 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    const d = r.data as {
      unmatchedPreview: Array<{ id: string }>
      flaggedCount: number
      preFilteredCount: number
    }
    expect(d.unmatchedPreview.map((m) => m.id)).toEqual(['m5', 'm6', 'm7'])
    expect(d.flaggedCount).toBe(3)
    expect(d.preFilteredCount).toBe(2)
    const inboxCalls = router.outlookUrls.filter((u) => u.includes('/inbox/messages'))
    expect(inboxCalls[1]).toBe(link(5))
  })

  it('peekNextBatch pages the same way, with the narrow Id,Flag select', async () => {
    router.inboxPages = [
      { messages: [flagged(0), flagged(1), flagged(2)], nextLink: link(3) },
      { messages: [inboxMail(3), inboxMail(4)], nextLink: link(5) },
    ]
    const r = await sw.send({ type: 'peekNextBatch', batchSize: 2 })
    expect(r.data).toEqual({ eligibleCount: 2, cappedAtBatchSize: true, totalFetched: 5 })
    expect(router.outlookUrls[0]).toContain('$select=Id,Flag')
  })

  it('stops after 5 pages when every page is flagged', async () => {
    router.inboxPages = Array.from({ length: 8 }, (_, p) => ({
      messages: [flagged(p * 2), flagged(p * 2 + 1)],
      nextLink: link((p + 1) * 2),
    }))
    const r = await sw.send({ type: 'peekNextBatch', batchSize: 2 })
    expect(r.data).toMatchObject({ eligibleCount: 0, totalFetched: 10 })
    expect(router.outlookUrls.filter((u) => u.includes('/inbox/messages'))).toHaveLength(5)
  })
})
