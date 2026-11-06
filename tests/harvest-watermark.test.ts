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
  HARVEST_MAX_PAGES,
  HARVEST_PAGE_SIZE,
  HARVEST_WATERMARK_MARGIN_MS,
  harvestWindowStart,
  nextHarvestWatermark,
  resumeHarvestWatermark,
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
  it('capped run: resume one second before the last row; always past the window start (C7)', () => {
    const rows = (...ts: Array<string | undefined>) => ts.map((t, i) => ({ Id: `m${i}`, LastModifiedDateTime: t }))
    expect(
      resumeHarvestWatermark(rows('2026-11-03T10:00:00Z', '2026-11-03T10:05:07Z') as never, '2026-11-01T00:00:00Z'),
    ).toBe('2026-11-03T10:05:06.000Z')
    // A whole capped run inside the window start's own second: step past it.
    expect(
      resumeHarvestWatermark(rows('2026-11-03T10:05:07Z', '2026-11-03T10:05:07Z') as never, '2026-11-03T10:05:06Z'),
    ).toBe('2026-11-03T10:05:07.000Z')
    expect(resumeHarvestWatermark(rows(undefined, 'garbage') as never, '2026-11-01T00:00:00Z')).toBeNull()
  })
  // O5 (2026-11-06 r2): rounding DOWN landed back on the window start when the
  // timestamps carried fractions of a second — `gt 10:05:07` still matches
  // 10:05:07.5, so a capped run of one bulk change never advanced.
  it('a capped run inside one (fractional) second still moves the watermark past it (O5)', () => {
    const rows = (...ts: string[]) => ts.map((t, i) => ({ Id: `m${i}`, LastModifiedDateTime: t }))
    const start = '2026-11-03T10:05:07Z'
    const next = resumeHarvestWatermark(
      rows('2026-11-03T10:05:07.4000000Z', '2026-11-03T10:05:07.5000000Z') as never,
      start,
    )
    expect(next).toBe('2026-11-03T10:05:08.000Z')
    // …and the window the next run asks for is strictly later.
    expect(Date.parse(harvestWindowStart(next, Date.parse('2026-11-05T12:00:00Z')))).toBeGreaterThan(Date.parse(start))
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
    const t0 = Date.now() - 50 * 60 * 1000
    fetchSpy
      .mockResolvedValueOnce(
        // A FULL page — there is more to read…
        json({
          value: Array.from({ length: HARVEST_PAGE_SIZE }, (_, i) => ({
            Id: `m${i}`,
            Subject: 'x',
            ParentFolderId: JIA,
            LastModifiedDateTime: new Date(t0 + i * 1000).toISOString(),
          })),
        }),
      )
      // …and reading it fails.
      .mockResolvedValueOnce(new Response('bad request', { status: 400 }))
    const r = await sw.send({ type: 'harvestManualFilings' })
    expect(r.ok).toBe(false)
    expect(await getLastHarvestAt()).toBe(w)
  })

  // C7 (2026-11-06): more modified mail than the page cap (read-state changes
  // count) meant every run was incomplete, so the watermark never moved, each
  // run paid all 30 pages on the same newest mail, and the rest was never seen.
  // Keyset pages since O5 (2026-11-06 r2): the server below honours the
  // filter / $top / $skip of each request like Exchange does.
  it('a PAGE-CAPPED run advances the watermark to where it stopped; the next run resumes there', async () => {
    // Relative to now: the window start is clamped to the 14-day lookback,
    // so fixed dates break once the real clock moves past them.
    const now = Math.floor(Date.now() / 60_000) * 60_000
    const w0 = new Date(now - 5 * 86_400_000).toISOString()
    await setLastHarvestAt(w0)
    const base = now - 4 * 86_400_000
    const iso = (ms: number) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z')
    const rows = Array.from({ length: HARVEST_PAGE_SIZE * HARVEST_MAX_PAGES * 2 }, (_, i) => ({
      Id: `m${i}`,
      Subject: `s${i}`,
      ParentFolderId: JIA,
      LastModifiedDateTime: iso(base + i * 60_000),
    }))
    let lastServed = 0
    fetchSpy.mockImplementation(async (input: unknown) => {
      const url = decodeURIComponent(String(input))
      const f = /LastModifiedDateTime (gt|ge) (\S+?)(&|$)/.exec(url)!
      const bound = Date.parse(f[2]!)
      const top = Number(/\$top=(\d+)/.exec(url)![1])
      const skip = Number(/\$skip=(\d+)/.exec(url)?.[1] ?? 0)
      const value = rows
        .filter((r) => (f[1] === 'gt' ? Date.parse(r.LastModifiedDateTime) > bound : Date.parse(r.LastModifiedDateTime) >= bound))
        .slice(skip, skip + top)
      for (const r of value) lastServed = Math.max(lastServed, Date.parse(r.LastModifiedDateTime))
      return json({ value })
    })
    const r = await sw.send({ type: 'harvestManualFilings' })
    expect(r).toMatchObject({ ok: true, data: { complete: false } })
    expect(fetchSpy).toHaveBeenCalledTimes(HARVEST_MAX_PAGES)
    const decoded = decodeURIComponent(String(fetchSpy.mock.calls[0]![0]))
    expect(decoded).toContain('$orderby=LastModifiedDateTime asc')
    expect(await getLastHarvestAt()).toBe(new Date(lastServed - 1000).toISOString())

    fetchSpy.mockReset()
    fetchSpy.mockResolvedValue(json({ value: [] }))
    await sw.send({ type: 'harvestManualFilings' })
    expect(Date.parse(sinceOf(String(fetchSpy.mock.calls[0]![0])))).toBe(lastServed - 1000)
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
