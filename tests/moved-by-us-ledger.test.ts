import { beforeEach, describe, expect, it } from 'vitest'
import { MOVED_BY_US_LEDGER_TTL_MS, THREAD_MEMORY_MAX_AGE_MS } from '@/shared/constants'
import { addToMovedByUsLedger, getMovedByUsLedger } from '@/shared/storage'

const DAY = 24 * 60 * 60 * 1000

beforeEach(async () => {
  await chrome.storage.local.clear()
})

const entry = (over: Partial<Parameters<typeof addToMovedByUsLedger>[0][number]> = {}) => ({
  newMessageId: 'nid-' + Math.random().toString(36).slice(2),
  destinationFolderId: 'fid1',
  conversationId: 'conv1',
  normalizedSubject: '112訴204 開庭',
  movedAt: Date.now(),
  ...over,
})

describe('movedByUsLedger', () => {
  it('writes and reads entries', async () => {
    await addToMovedByUsLedger([entry({ newMessageId: 'a' }), entry({ newMessageId: 'b' })])
    const ledger = await getMovedByUsLedger()
    expect(ledger.map((e) => e.newMessageId).sort()).toEqual(['a', 'b'])
  })

  it('dedups by newMessageId — newest write wins', async () => {
    await addToMovedByUsLedger([entry({ newMessageId: 'x', destinationFolderId: 'old' })])
    await addToMovedByUsLedger([entry({ newMessageId: 'x', destinationFolderId: 'new' })])
    const ledger = await getMovedByUsLedger()
    expect(ledger).toHaveLength(1)
    expect(ledger[0]!.destinationFolderId).toBe('new')
  })

  it('drops entries missing newMessageId or destinationFolderId', async () => {
    await addToMovedByUsLedger([
      entry({ newMessageId: '' }),
      entry({ newMessageId: 'ok', destinationFolderId: '' }),
      entry({ newMessageId: 'keep' }),
    ])
    const ledger = await getMovedByUsLedger()
    expect(ledger.map((e) => e.newMessageId)).toEqual(['keep'])
  })

  it('drops entries past the TTL on write (TTL = thread-memory max age, 2026-11-05)', async () => {
    await addToMovedByUsLedger([
      entry({ newMessageId: 'stale', movedAt: Date.now() - THREAD_MEMORY_MAX_AGE_MS - DAY }),
    ])
    expect(await getMovedByUsLedger()).toHaveLength(0)
  })

  it('keeps a 22-day-old own move (was dropped at the old 21-day TTL → re-harvested)', async () => {
    await addToMovedByUsLedger([entry({ newMessageId: 'old-own-move', movedAt: Date.now() - 22 * DAY })])
    expect((await getMovedByUsLedger()).map((e) => e.newMessageId)).toEqual(['old-own-move'])
  })

  it('ledger TTL never falls below the thread-memory retention it protects', () => {
    expect(MOVED_BY_US_LEDGER_TTL_MS).toBeGreaterThanOrEqual(THREAD_MEMORY_MAX_AGE_MS)
  })

  it('stores slim entries — normalizedSubject (never read by the harvester) is dropped', async () => {
    await addToMovedByUsLedger([
      { ...entry({ newMessageId: 'slim' }), normalizedSubject: '112訴204 開庭' } as never,
    ])
    const [stored] = await getMovedByUsLedger()
    expect(stored).toEqual({
      newMessageId: 'slim',
      destinationFolderId: 'fid1',
      conversationId: 'conv1',
      movedAt: expect.any(Number),
    })
    expect(stored).not.toHaveProperty('normalizedSubject')
  })

  it('is a no-op on empty input', async () => {
    await addToMovedByUsLedger([])
    expect(await getMovedByUsLedger()).toHaveLength(0)
  })

  it('getMovedByUsLedger returns [] when nothing stored', async () => {
    expect(await getMovedByUsLedger()).toEqual([])
  })
})
