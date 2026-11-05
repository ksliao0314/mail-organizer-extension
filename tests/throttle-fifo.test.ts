// Rate limiting + concurrency (2026-11-05 audit, item 4).
//
// TokenBucket used to compute every waiter's sleep from the same pre-debit
// balance and clamp at 0: 30 concurrent take() → 10 immediately and the other
// 20 ALL together ~1 s later — 30 requests in ~1 s against a 1 req/s budget.
// It now reserves FIFO (balance may go negative) so waiters come out spaced.
// Separately, nothing capped requests in flight; OutlookApi.request() now
// holds a Semaphore slot (max 4 — Exchange's per-mailbox limit) per round trip.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Semaphore, TokenBucket } from '@/shared/throttle'
import { MAX_CONCURRENT_REQUESTS, OutlookApi } from '@/shared/outlook-api'

describe('TokenBucket FIFO reservation', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  it('30 concurrent take() on a 10-burst / 1-per-sec bucket are spaced ~1 s apart', async () => {
    const bucket = new TokenBucket(10, 1)
    let done = 0
    const all = Array.from({ length: 30 }, () => bucket.take().then(() => done++))

    await vi.advanceTimersByTimeAsync(0)
    expect(done).toBe(10) // the burst

    // Old code: all 30 finished here (~1 s). Now only ONE more per second.
    await vi.advanceTimersByTimeAsync(1000)
    expect(done).toBe(11)
    await vi.advanceTimersByTimeAsync(4000)
    expect(done).toBe(15)

    await vi.advanceTimersByTimeAsync(15_000)
    expect(done).toBe(30)
    await Promise.all(all)
  })

  it('an aborted waiter rejects and gives its reservation back', async () => {
    const bucket = new TokenBucket(1, 1)
    await bucket.take() // burst used
    const ctl = new AbortController()
    const aborted = bucket.take(1, ctl.signal)
    ctl.abort()
    await expect(aborted).rejects.toBeDefined()
    // The refunded slot means the next taker waits ~1 s, not ~2 s.
    let later = false
    void bucket.take().then(() => (later = true))
    await vi.advanceTimersByTimeAsync(1000)
    expect(later).toBe(true)
  })
})

describe('Semaphore', () => {
  it('never lets more than `max` holders in, and grants FIFO', async () => {
    const sem = new Semaphore(2)
    const order: number[] = []
    const releases: Array<() => void> = []
    const waiters = [1, 2, 3, 4].map((n) =>
      sem.acquire().then((release) => {
        order.push(n)
        releases.push(release)
      }),
    )
    await Promise.resolve()
    await Promise.resolve()
    expect(sem.inFlight).toBe(2)
    expect(order).toEqual([1, 2])
    releases[0]!()
    releases[0]!() // idempotent — must not free a second slot
    await Promise.resolve()
    expect(order).toEqual([1, 2, 3])
    expect(sem.inFlight).toBe(2)
    releases[1]!()
    releases[2]!()
    await Promise.all(waiters)
    expect(order).toEqual([1, 2, 3, 4])
  })

  it('a waiter aborted in the queue rejects without consuming a slot', async () => {
    const sem = new Semaphore(1)
    const release = await sem.acquire()
    const ctl = new AbortController()
    const queued = sem.acquire(ctl.signal)
    ctl.abort()
    await expect(queued).rejects.toBeDefined()
    release()
    expect(sem.inFlight).toBe(0)
    const again = await sem.acquire()
    expect(sem.inFlight).toBe(1)
    again()
  })
})

describe('OutlookApi request concurrency', () => {
  let fetchSpy: ReturnType<typeof vi.fn>
  beforeEach(() => {
    fetchSpy = vi.fn()
    ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
  })
  afterEach(() => vi.restoreAllMocks())

  it('a 12-wide Promise.all never has more than 4 requests in flight', async () => {
    let inFlight = 0
    let peak = 0
    fetchSpy.mockImplementation(async () => {
      inFlight++
      peak = Math.max(peak, inFlight)
      await new Promise((r) => setTimeout(r, 5))
      inFlight--
      return new Response(JSON.stringify({ UniqueBody: { ContentType: 'text', Content: 'x' } }), {
        status: 200,
      })
    })
    // Generous bucket so only the semaphore is under test.
    const api = new OutlookApi(async () => 'tok', undefined, { bucket: new TokenBucket(1000, 1000) })
    await Promise.all(Array.from({ length: 12 }, (_, i) => api.getMessageBody(`m${i}`)))
    expect(fetchSpy).toHaveBeenCalledTimes(12)
    expect(MAX_CONCURRENT_REQUESTS).toBe(4)
    expect(peak).toBe(4)
  })
})
