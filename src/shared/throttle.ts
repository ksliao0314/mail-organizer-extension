// Token-bucket rate limiter + concurrency semaphore.
// The Outlook budget (OUTLOOK_REQUESTS_PER_SEC steady, OUTLOOK_REQUEST_BURST
// burst — 8/s and 40 since v1.0.9) lives in constants.ts; this module only
// implements the mechanisms.

function abortReason(signal: AbortSignal): unknown {
  return signal.reason ?? new DOMException('Aborted', 'AbortError')
}

export class TokenBucket {
  private tokens: number
  private lastRefillMs: number

  constructor(
    private readonly capacity: number,
    private readonly refillPerSec: number,
  ) {
    this.tokens = capacity
    this.lastRefillMs = Date.now()
  }

  private refill(): void {
    const now = Date.now()
    const elapsedSec = (now - this.lastRefillMs) / 1000
    this.tokens = Math.min(this.capacity, this.tokens + elapsedSec * this.refillPerSec)
    this.lastRefillMs = now
  }

  async take(count = 1, signal?: AbortSignal): Promise<void> {
    // Early-out so a cancelled batch doesn't even claim a token (the next
    // outbound fetch would just error anyway, but reserving the token first
    // gave abort a 100-1000ms latency tail).
    if (signal?.aborted) throw abortReason(signal)
    this.refill()
    // FIFO reservation (2026-11-05): debit FIRST, letting the balance go
    // negative, and sleep for exactly the deficit this caller created. The
    // old code computed every waiter's sleep from the same pre-debit balance
    // and then clamped at 0 — 30 concurrent take() calls released 10
    // immediately and the remaining 20 all together ~1 s later, i.e. a burst
    // of 30 in ~1 s against a 1/s budget. With reservations the N-th waiter
    // beyond capacity waits N/refillPerSec, so waiters come out spaced.
    this.tokens -= count
    if (this.tokens >= 0) return
    const waitMs = Math.ceil((-this.tokens / this.refillPerSec) * 1000)
    try {
      await abortableSleep(waitMs, signal)
    } catch (e) {
      // Give the reservation back — an aborted waiter never sends a request,
      // so it must not push everyone queued behind it further out.
      this.refill()
      this.tokens = Math.min(this.capacity, this.tokens + count)
      throw e
    }
  }
}

/**
 * FIFO counting semaphore (2026-11-05). Bounds how many Outlook requests are
 * in flight at once: Exchange throttles per mailbox at ~4 concurrent
 * requests, and every fan-out in the codebase (fetchFolderTree's sibling
 * Promise.all, the SW's 25-wide getMessageBody Promise.all) used to put
 * EVERYTHING on the wire at once — the token bucket only spaces start times,
 * it doesn't cap concurrency.
 */
export class Semaphore {
  private active = 0
  private readonly waiters: Array<{ grant: () => void }> = []

  constructor(private readonly max: number) {}

  /** Number of holders right now (for tests / diagnostics). */
  get inFlight(): number {
    return this.active
  }

  /**
   * Resolves with an idempotent release function once a slot is free.
   * Rejects (without consuming a slot) when `signal` aborts while waiting.
   */
  acquire(signal?: AbortSignal): Promise<() => void> {
    if (signal?.aborted) return Promise.reject(abortReason(signal))
    return new Promise<() => void>((resolve, reject) => {
      let released = false
      const release = () => {
        if (released) return
        released = true
        this.active--
        this.pump()
      }
      const waiter = {
        grant: () => {
          signal?.removeEventListener('abort', onAbort)
          this.active++
          resolve(release)
        },
      }
      const onAbort = () => {
        const idx = this.waiters.indexOf(waiter)
        if (idx >= 0) this.waiters.splice(idx, 1)
        reject(abortReason(signal!))
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.waiters.push(waiter)
      this.pump()
    })
  }

  private pump(): void {
    while (this.active < this.max && this.waiters.length > 0) {
      this.waiters.shift()!.grant()
    }
  }
}

/**
 * One signal that aborts when any input does (AbortSignal.any, 2026-11-06) —
 * plus `dispose()`, which detaches it from the inputs again. Callers pass a
 * batch's long-lived signal into every request; without the detach each
 * request would leave a listener on it for the rest of the batch.
 */
export function linkAbortSignals(signals: Array<AbortSignal | undefined>): {
  signal: AbortSignal
  dispose: () => void
} {
  const ctl = new AbortController()
  const detach: Array<() => void> = []
  for (const s of signals) {
    if (!s) continue
    if (s.aborted) {
      ctl.abort(s.reason)
      break
    }
    const onAbort = () => ctl.abort(s.reason)
    s.addEventListener('abort', onAbort, { once: true })
    detach.push(() => s.removeEventListener('abort', onAbort))
  }
  return {
    signal: ctl.signal,
    dispose: () => {
      for (const d of detach.splice(0)) d()
    },
  }
}

/**
 * `p`, unless `signal` aborts first — then rejects with the abort reason
 * (2026-11-06 r2). `p` itself keeps running (e.g. a shared token refresh
 * other callers still wait on); only this caller stops waiting.
 */
export function abortable<T>(p: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return p
  if (signal.aborted) {
    p.catch(() => {}) // nobody waits on it any more
    return Promise.reject(abortReason(signal))
  }
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => {
      p.catch(() => {})
      reject(abortReason(signal))
    }
    signal.addEventListener('abort', onAbort, { once: true })
    p.then(
      (v) => {
        signal.removeEventListener('abort', onAbort)
        resolve(v)
      },
      (e) => {
        signal.removeEventListener('abort', onAbort)
        reject(e)
      },
    )
  })
}

export function abortableSleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(abortReason(signal))
      return
    }
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort)
      resolve()
    }, ms)
    const onAbort = () => {
      clearTimeout(timer)
      reject(abortReason(signal!))
    }
    signal?.addEventListener('abort', onAbort, { once: true })
  })
}
