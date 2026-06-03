// Regression tests for the recentlyProcessed ledger (2026-06-03).
//
// Root-cause fix for "already-moved emails reappear in the continue
// list". Outlook's message store is eventually consistent: for a short
// window after a move/delete, the inbox listing can still return the
// just-handled email (with its now-dead id), so the next batch would
// try to re-move it and 404. We record moved/deleted ids into this
// ledger and filter the next batch's inbox fetch against it. The ledger
// auto-expires after RECENTLY_PROCESSED_TTL_MS so it only needs to
// outlast Outlook's propagation lag, not user sessions.

import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  addToRecentlyProcessed,
  getRecentlyProcessed,
  getRecentlyProcessedIds,
} from '@/shared/storage'
import { RECENTLY_PROCESSED_TTL_MS } from '@/shared/constants'

afterEach(() => {
  vi.useRealTimers()
})

describe('recentlyProcessed ledger', () => {
  it('records ids and returns them within the TTL window', async () => {
    await addToRecentlyProcessed(['a', 'b', 'c'])
    const ids = await getRecentlyProcessedIds()
    expect(ids.has('a')).toBe(true)
    expect(ids.has('b')).toBe(true)
    expect(ids.has('c')).toBe(true)
    expect(ids.size).toBe(3)
  })

  it('ignores empty / falsy ids', async () => {
    await addToRecentlyProcessed(['x', '', 'y'])
    const ids = await getRecentlyProcessedIds()
    expect(ids.has('x')).toBe(true)
    expect(ids.has('y')).toBe(true)
    expect(ids.has('')).toBe(false)
    expect(ids.size).toBe(2)
  })

  it('is idempotent on duplicate ids', async () => {
    await addToRecentlyProcessed(['dup'])
    await addToRecentlyProcessed(['dup', 'dup'])
    const rec = await getRecentlyProcessed()
    expect(Object.keys(rec)).toEqual(['dup'])
  })

  it('no-op on empty input', async () => {
    await addToRecentlyProcessed([])
    expect(await getRecentlyProcessedIds()).toEqual(new Set())
  })

  it('excludes entries older than the TTL', async () => {
    vi.useFakeTimers()
    vi.setSystemTime(1_000_000)
    await addToRecentlyProcessed(['old'])
    // Jump past the TTL.
    vi.setSystemTime(1_000_000 + RECENTLY_PROCESSED_TTL_MS + 1)
    const ids = await getRecentlyProcessedIds()
    expect(ids.has('old')).toBe(false)
  })

  it('keeps a fresh entry while dropping an expired one', async () => {
    vi.useFakeTimers()
    vi.setSystemTime(1_000_000)
    await addToRecentlyProcessed(['stale'])
    vi.setSystemTime(1_000_000 + RECENTLY_PROCESSED_TTL_MS + 1)
    await addToRecentlyProcessed(['fresh'])
    const ids = await getRecentlyProcessedIds()
    expect(ids.has('stale')).toBe(false)
    expect(ids.has('fresh')).toBe(true)
  })

  it('re-touching an id refreshes its TTL window', async () => {
    vi.useFakeTimers()
    vi.setSystemTime(1_000_000)
    await addToRecentlyProcessed(['keep'])
    // Just before expiry, re-record it → extends the window.
    vi.setSystemTime(1_000_000 + RECENTLY_PROCESSED_TTL_MS - 1)
    await addToRecentlyProcessed(['keep'])
    // Now advance to just past the ORIGINAL expiry — should still be present
    // because the second touch reset the clock.
    vi.setSystemTime(1_000_000 + RECENTLY_PROCESSED_TTL_MS + 1)
    const ids = await getRecentlyProcessedIds()
    expect(ids.has('keep')).toBe(true)
  })
})
