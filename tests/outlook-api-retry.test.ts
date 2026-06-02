// Regression tests for the 2026-05-27 non-idempotent retry fix:
//
// Original bug: outlook-api `request()` retried fetch() failures and 5xx
// for ALL HTTP methods. For POST /me/messages/{id}/move (and DELETE
// /messages/{id}), a successful server-side execution whose response
// was lost (network blip, TLS reset, transient 5xx) would be retried
// with the SAME (now-stale) message Id. The retry hit 404
// ErrorItemNotFound — surfacing as a hard error in the popup even
// though the message had already moved to its intended destination.
//
// Fix:
//   1. Only retry NON-idempotent methods on explicit 429 (rate limit —
//      server promised "didn't process"). Network errors + 5xx → bail
//      immediately with `uncertain: true` on the OutlookError so callers
//      can render "may have succeeded" rather than "failed".
//   2. Idempotent methods (GET / HEAD / PUT / OPTIONS) keep retrying on
//      network + 5xx, since re-executing them is safe.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { OutlookApi, OutlookError } from '@/shared/outlook-api'

let fetchSpy: ReturnType<typeof vi.fn>

beforeEach(() => {
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
})

afterEach(() => {
  vi.restoreAllMocks()
})

function makeApi(): OutlookApi {
  // Token provider always returns a fixed token; tests don't exercise
  // the 401 refresh path.
  return new OutlookApi(async () => 'fake-token')
}

// ---- POST: network error → bail with uncertain ---------------------------

describe('POST non-idempotent — no auto-retry on network error', () => {
  it('moveMessage rejects fetch once → throws immediately (no retry)', async () => {
    fetchSpy.mockRejectedValueOnce(new TypeError('Network down'))
    const api = makeApi()
    await expect(api.moveMessage('msg-id', 'folder-id')).rejects.toMatchObject({
      status: 0,
      uncertain: true,
    })
    // Critical: fetch called exactly ONCE. Before the fix, this was 4
    // (initial + 3 retries) — each retry posted /move again, and a
    // server-side success on the first try would 404 on retry.
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it('deleteMessage rejects fetch once → throws immediately (no retry)', async () => {
    fetchSpy.mockRejectedValueOnce(new TypeError('Connection reset'))
    const api = makeApi()
    await expect(api.deleteMessage('msg-id')).rejects.toMatchObject({
      status: 0,
      uncertain: true,
    })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it('OutlookError carries uncertain=true so caller can render "may have succeeded"', async () => {
    fetchSpy.mockRejectedValueOnce(new TypeError('Network blip'))
    const api = makeApi()
    try {
      await api.moveMessage('msg-id', 'folder-id')
      expect.unreachable('should have thrown')
    } catch (e) {
      expect(e).toBeInstanceOf(OutlookError)
      expect((e as OutlookError).uncertain).toBe(true)
      expect((e as OutlookError).message).toMatch(/non-idempotent/i)
    }
  })
})

// ---- POST: 5xx → bail with uncertain -------------------------------------

describe('POST non-idempotent — no auto-retry on 5xx', () => {
  it('moveMessage 500 → throws immediately with uncertain=true', async () => {
    fetchSpy.mockResolvedValueOnce(
      new Response('Internal Server Error', { status: 500 }),
    )
    const api = makeApi()
    await expect(api.moveMessage('msg-id', 'folder-id')).rejects.toMatchObject({
      status: 500,
      uncertain: true,
    })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it('moveMessage 503 → throws immediately (no retry)', async () => {
    fetchSpy.mockResolvedValueOnce(
      new Response('Service Unavailable', { status: 503 }),
    )
    const api = makeApi()
    await expect(api.moveMessage('msg-id', 'folder-id')).rejects.toMatchObject({
      status: 503,
      uncertain: true,
    })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })
})

// ---- POST: 429 → still retried (server explicit "didn't process") --------

describe('POST non-idempotent — still retries on 429', () => {
  it('moveMessage 429 then 200 → succeeds on second attempt', async () => {
    fetchSpy
      .mockResolvedValueOnce(
        new Response('rate limited', {
          status: 429,
          headers: { 'Retry-After': '0' },
        }),
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ Id: 'new-message-id' }), { status: 200 }),
      )
    const api = makeApi()
    const res = await api.moveMessage('msg-id', 'folder-id')
    expect(res).toEqual({ Id: 'new-message-id' })
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })
})

// ---- GET: idempotent — still retries on network + 5xx --------------------

describe('GET idempotent — retries on network error', () => {
  it('listInboxMessages: network fail then 200 → second attempt succeeds', async () => {
    fetchSpy
      .mockRejectedValueOnce(new TypeError('Network blip'))
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ value: [] }), { status: 200 }),
      )
    const api = makeApi()
    const r = await api.listInboxMessages({ top: 1 })
    expect(r).toEqual([])
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })

  it('listInboxMessages: 500 then 200 → second attempt succeeds', async () => {
    fetchSpy
      .mockResolvedValueOnce(
        new Response('boom', {
          status: 500,
          headers: { 'Retry-After': '0' },
        }),
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ value: [] }), { status: 200 }),
      )
    const api = makeApi()
    const r = await api.listInboxMessages({ top: 1 })
    expect(r).toEqual([])
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })
})

// ---- 404 surfaces normally (no retry, uncertain=false) -------------------

describe('404 on non-idempotent — surfaces normally', () => {
  it('moveMessage 404 → throws status=404, uncertain=false', async () => {
    fetchSpy.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          error: {
            code: 'ErrorItemNotFound',
            message: 'The specified object was not found in the store.',
          },
        }),
        { status: 404, statusText: 'Not Found' },
      ),
    )
    const api = makeApi()
    try {
      await api.moveMessage('msg-id', 'folder-id')
      expect.unreachable('should have thrown')
    } catch (e) {
      expect(e).toBeInstanceOf(OutlookError)
      expect((e as OutlookError).status).toBe(404)
      // 4xx is "definitely didn't succeed" from the client's POV —
      // caller's 404 handler treats this as already-moved soft-skip.
      // We don't set uncertain on 4xx; only 5xx + network gets it.
      expect((e as OutlookError).uncertain).toBe(false)
      expect((e as OutlookError).message).toMatch(/ErrorItemNotFound/)
    }
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })
})
