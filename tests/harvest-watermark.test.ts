// harvestManualFilings lookback & watermark (2026-11-05 audit, item 6), driven
// through the REAL service-worker dispatcher with a fake OWA tab + fetch.
//
// Old behaviour: no watermark → no $filter; with one → still only $top=10 per
// leaf folder (our own moves ate the slots); the watermark was the client
// clock at the END of a ~200-request run, so filings made meanwhile (or inside
// clock skew) were lost. Now: one mailbox-wide modified-since query, paged,
// watermark = scan START − margin, advanced only after a complete run; first
// run bounded to a 14-day lookback; plus a synchronous in-flight guard.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  HARVEST_LOOKBACK_MS,
  HARVEST_WATERMARK_MARGIN_MS,
  harvestWindowStart,
  nextHarvestWatermark,
  selectHarvestScanItems,
} from '@/background/harvest'
import {
  getConversationMemory,
  getLastHarvestAt,
  setFolderCache,
  setLastHarvestAt,
  setSettings,
} from '@/shared/storage'
import type { MailFolderNode } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const CASES = 'FOLDER-CASES-0000000000000'
const JIA = 'FOLDER-JIA-000000000000000'
const YI = 'FOLDER-YI-0000000000000000'
const TREE: MailFolderNode[] = [
  {
    id: CASES,
    displayName: '案件',
    path: '案件',
    children: [
      { id: JIA, displayName: '甲', path: '案件/甲', parentFolderId: CASES, children: [] },
      { id: YI, displayName: '乙', path: '案件/乙', parentFolderId: CASES, children: [] },
    ],
  },
]

let sw: SwHarness
let fetchSpy: ReturnType<typeof vi.fn>

beforeAll(async () => {
  sw = await loadServiceWorker()
})

beforeEach(async () => {
  sw.tabsQuery.mockResolvedValue([
    { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
  ])
  sw.tabsSendMessage.mockResolvedValue({ ok: true, secret: 'tok', expiresOn: Date.now() / 1000 + 3600 })
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
  await setSettings({ harvestEnabled: true, primaryRootPath: '案件' })
  await setFolderCache({ updatedAt: new Date().toISOString(), tree: TREE })
})
afterEach(() => vi.restoreAllMocks())

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status })
}
const sinceOf = (url: string) =>
  decodeURIComponent(url).match(/LastModifiedDateTime gt (\S+?)(&|$)/)?.[1] ?? ''

describe('pure window helpers', () => {
  const now = Date.parse('2026-11-05T12:00:00Z')
  it('first run → bounded lookback; stale watermark clamped; future watermark pulled back', () => {
    expect(harvestWindowStart(null, now)).toBe('2026-10-22T12:00:00Z')
    expect(harvestWindowStart('garbage', now)).toBe('2026-10-22T12:00:00Z')
    expect(harvestWindowStart('2026-09-01T00:00:00Z', now)).toBe('2026-10-22T12:00:00Z')
    expect(harvestWindowStart('2026-11-05T09:30:00.000Z', now)).toBe('2026-11-05T09:30:00Z')
    expect(harvestWindowStart('2026-11-06T00:00:00Z', now)).toBe('2026-11-05T11:50:00Z')
  })
  it('watermark = scan start − margin', () => {
    expect(nextHarvestWatermark(now)).toBe('2026-11-05T11:50:00.000Z')
  })
  it('keeps only rows sitting in a target folder', () => {
    const out = selectHarvestScanItems(
      [
        { Id: 'a', Subject: 's', ParentFolderId: JIA, From: { EmailAddress: { Address: 'x@y.example' } } },
        { Id: 'b', Subject: 's', ParentFolderId: 'INBOX-ID' },
        { Id: '', Subject: 's', ParentFolderId: JIA },
      ] as never,
      new Map([[JIA, '案件/甲']]),
    )
    expect(out).toEqual([
      { id: 'a', subject: 's', from: 'x@y.example', conversationId: undefined, folderId: JIA, folderPath: '案件/甲' },
    ])
  })
})

describe('harvestManualFilings handler', () => {
  it('first run: one mailbox-wide query with a 14-day lookback; watermark = START − margin', async () => {
    const before = Date.now()
    fetchSpy.mockImplementation(async (input: unknown) => {
      const url = String(input)
      if (!url.includes('/me/messages?')) throw new Error(`unexpected ${url}`)
      return json({
        value: [
          { Id: 'm1', Subject: '新案件 往來函', ConversationId: 'c1', ParentFolderId: JIA },
          { Id: 'm2', Subject: 'inbox noise', ConversationId: 'c2', ParentFolderId: 'INBOX-ID' },
        ],
      })
    })
    const r = await sw.send({ type: 'harvestManualFilings' })
    const after = Date.now()
    expect(r).toMatchObject({ ok: true, data: { candidates: 1, complete: true, scannedFolders: 2 } })
    // Exactly ONE listing request — not one per leaf folder.
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    const since = Date.parse(sinceOf(String(fetchSpy.mock.calls[0]![0])))
    expect(since).toBeGreaterThanOrEqual(before - HARVEST_LOOKBACK_MS - 1000)
    expect(since).toBeLessThanOrEqual(after - HARVEST_LOOKBACK_MS)
    const w = Date.parse((await getLastHarvestAt())!)
    expect(w).toBeGreaterThanOrEqual(before - HARVEST_WATERMARK_MARGIN_MS)
    expect(w).toBeLessThanOrEqual(after - HARVEST_WATERMARK_MARGIN_MS)
    expect((await getConversationMemory()).c1).toMatchObject({ folderId: JIA, harvested: true })
    expect((await getConversationMemory()).c2).toBeUndefined()
  })

  it('uses the stored watermark as the window start', async () => {
    const w = new Date(Date.now() - 3 * 60 * 60 * 1000).toISOString()
    await setLastHarvestAt(w)
    fetchSpy.mockResolvedValue(json({ value: [] }))
    await sw.send({ type: 'harvestManualFilings' })
    expect(Date.parse(sinceOf(String(fetchSpy.mock.calls[0]![0])))).toBe(Math.floor(Date.parse(w) / 1000) * 1000)
  })

  it('a failure mid-paging leaves the watermark untouched', async () => {
    const w = new Date(Date.now() - 60 * 60 * 1000).toISOString()
    await setLastHarvestAt(w)
    fetchSpy
      .mockResolvedValueOnce(
        json({
          value: [{ Id: 'm1', Subject: 'x', ParentFolderId: JIA }],
          '@odata.nextLink': 'https://outlook.office.com/api/v2.0/me/messages?$skip=100',
        }),
      )
      .mockResolvedValueOnce(new Response('bad request', { status: 400 }))
    const r = await sw.send({ type: 'harvestManualFilings' })
    expect(r.ok).toBe(false)
    expect(await getLastHarvestAt()).toBe(w)
  })

  it('a second run while one is in flight is refused synchronously', async () => {
    let release: () => void = () => {}
    const gate = new Promise<void>((r) => (release = r))
    fetchSpy.mockImplementation(async () => {
      await gate
      return json({ value: [] })
    })
    const first = sw.send({ type: 'harvestManualFilings' })
    const second = await sw.send({ type: 'harvestManualFilings' })
    expect(second).toMatchObject({ ok: false, code: 'ALREADY_RUNNING' })
    release()
    expect((await first).ok).toBe(true)
    // …and the guard is released afterwards.
    fetchSpy.mockResolvedValue(json({ value: [] }))
    expect((await sw.send({ type: 'harvestManualFilings' })).ok).toBe(true)
  })
})
