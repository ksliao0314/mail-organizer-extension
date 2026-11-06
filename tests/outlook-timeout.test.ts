// Per-request deadline (C9, 2026-11-06). fetch has no timeout of its own;
// with the 4-slot semaphore, four stalled requests used to block every
// Outlook call for good. Now each attempt is bounded — headers AND body — and
// a timeout means: GET → retried like a network blip; POST / DELETE / PATCH →
// `uncertain` (may have committed), never retried. A caller's abort keeps its
// own meaning, and the slot is released on every path.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { OUTLOOK_REQUEST_TIMEOUT_MS, OutlookApi, OutlookError } from '@/shared/outlook-api'
import { Semaphore, TokenBucket } from '@/shared/throttle'

let fetchSpy: ReturnType<typeof vi.fn>
beforeEach(() => {
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
})
afterEach(() => vi.restoreAllMocks())

const json = (v: unknown, status = 200) => new Response(JSON.stringify(v), { status })

/** Like real fetch: never answers, rejects with the signal's reason when aborted. */
function hang(_input: unknown, init?: RequestInit): Promise<Response> {
  return new Promise((_resolve, reject) => {
    const s = init?.signal
    s?.addEventListener('abort', () => reject(s.reason ?? new DOMException('Aborted', 'AbortError')))
  })
}

function makeApi(inflight = new Semaphore(4), requestTimeoutMs = 40) {
  return {
    inflight,
    api: new OutlookApi(async () => 'tok', undefined, {
      bucket: new TokenBucket(1000, 1000),
      inflight,
      requestTimeoutMs,
    }),
  }
}

describe('Outlook request deadline (C9)', () => {
  it('defaults to 60 s', () => {
    expect(OUTLOOK_REQUEST_TIMEOUT_MS).toBe(60_000)
  })

  it('a POST that never answers → uncertain (not "aborted"), sent once, slot released', async () => {
    const { api, inflight } = makeApi(new Semaphore(1))
    fetchSpy.mockImplementationOnce(hang)
    const err = await api.moveMessage('msg-1', 'folder-1').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(OutlookError)
    expect(err).toMatchObject({ status: 0, uncertain: true, aborted: false })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect(inflight.inFlight).toBe(0)
    // The single slot is usable again.
    fetchSpy.mockResolvedValueOnce(json({ value: [] }))
    await expect(api.listInboxMessages({ top: 1 })).resolves.toEqual([])
  })

  it('a GET that times out is retried like any transient failure', async () => {
    const { api, inflight } = makeApi()
    fetchSpy.mockImplementationOnce(hang).mockResolvedValueOnce(json({ value: [{ Id: 'm1' }] }))
    const out = await api.listInboxMessages({ top: 1 })
    expect(out.map((m) => m.Id)).toEqual(['m1'])
    expect(fetchSpy).toHaveBeenCalledTimes(2)
    expect(inflight.inFlight).toBe(0)
  })

  it('a 2xx whose body stalls: a POST is uncertain — it committed, only its result was lost', async () => {
    const { api, inflight } = makeApi()
    fetchSpy.mockImplementationOnce(async (_input: unknown, init?: RequestInit) => {
      const body = new ReadableStream({
        start(controller) {
          init?.signal?.addEventListener('abort', () => controller.error(new DOMException('Aborted', 'AbortError')))
        },
      })
      return new Response(body, { status: 200 })
    })
    const err = await api.moveMessage('msg-1', 'folder-1').catch((e: unknown) => e)
    expect(err).toMatchObject({ uncertain: true, aborted: false })
    expect((err as Error).message).toContain('result was lost')
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect(inflight.inFlight).toBe(0)
  })

  it('the caller’s abort is still the caller’s: AbortError for a GET, uncertain + aborted for a POST', async () => {
    const { api, inflight } = makeApi(new Semaphore(4), 10_000)
    fetchSpy.mockImplementation(hang)
    const ctl1 = new AbortController()
    const get = api.listInboxMessages({ top: 1, signal: ctl1.signal })
    get.catch(() => {})
    await vi.waitFor(() => expect(fetchSpy).toHaveBeenCalledTimes(1))
    ctl1.abort()
    await expect(get).rejects.toMatchObject({ name: 'AbortError' })

    const ctl2 = new AbortController()
    const post = api.moveMessage('m', 'f', ctl2.signal)
    post.catch(() => {})
    await vi.waitFor(() => expect(fetchSpy).toHaveBeenCalledTimes(2))
    ctl2.abort()
    await expect(post).rejects.toMatchObject({ uncertain: true, aborted: true })
    expect(inflight.inFlight).toBe(0)
  })

  it('getMyEmail({ unthrottled }) does not queue behind a drained rate limiter (C10)', async () => {
    // One token, refilled every 10 s: a throttled second request would wait ~10 s.
    const api = new OutlookApi(async () => 'tok', undefined, { bucket: new TokenBucket(1, 0.1) })
    fetchSpy.mockImplementation(async () => json({ value: [], EmailAddress: 'me@firm.example' }))
    await api.listInboxMessages({ top: 1 }) // drains the bucket
    const started = Date.now()
    await expect(api.getMyEmail(AbortSignal.timeout(1_000), { unthrottled: true })).resolves.toBe('me@firm.example')
    expect(Date.now() - started).toBeLessThan(500)
    // …while the same call through the queue would blow that budget.
    await expect(api.getMyEmail(AbortSignal.timeout(300))).rejects.toBeTruthy()
  })

  it('does not leave a listener on the caller’s long-lived signal per request', async () => {
    const { api } = makeApi()
    const ctl = new AbortController()
    const add = vi.spyOn(ctl.signal, 'addEventListener')
    const remove = vi.spyOn(ctl.signal, 'removeEventListener')
    fetchSpy.mockImplementation(async () => json({ value: [] }))
    for (let i = 0; i < 5; i++) await api.listInboxMessages({ top: 1, signal: ctl.signal })
    const added = add.mock.calls.filter((c) => c[0] === 'abort').length
    const removed = remove.mock.calls.filter((c) => c[0] === 'abort').length
    expect(added).toBeGreaterThan(0)
    expect(removed).toBe(added)
  })
})
