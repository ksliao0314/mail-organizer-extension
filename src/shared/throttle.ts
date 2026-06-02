// Token-bucket rate limiter.
// Outlook REST tolerates ~60 req/min (per handoff). We model it as 1 token/sec
// steady-state with 10 token burst so short flurries don't gate on the bucket.

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
    if (signal?.aborted) {
      throw signal.reason ?? new DOMException('Aborted', 'AbortError')
    }
    this.refill()
    if (this.tokens >= count) {
      this.tokens -= count
      return
    }
    const deficit = count - this.tokens
    const waitMs = Math.ceil((deficit / this.refillPerSec) * 1000)
    await abortableSleep(waitMs, signal)
    // After the sleep, check again — abortableSleep already rejects on abort
    // but defensively re-check before debiting tokens.
    if (signal?.aborted) {
      throw signal.reason ?? new DOMException('Aborted', 'AbortError')
    }
    this.refill()
    this.tokens = Math.max(0, this.tokens - count)
  }
}

export function abortableSleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(signal.reason ?? new DOMException('Aborted', 'AbortError'))
      return
    }
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort)
      resolve()
    }, ms)
    const onAbort = () => {
      clearTimeout(timer)
      reject(signal?.reason ?? new DOMException('Aborted', 'AbortError'))
    }
    signal?.addEventListener('abort', onAbort, { once: true })
  })
}
