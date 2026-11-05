// Start guards + keep-alive cadence (2026-11-05 audit, items 10 & 15).
//
// startInitialScan's guard flag was set AFTER `await isScanRunning()`, so two
// overlapping starts (double-click, FAB iframe + Options tab) both read false
// and both scanned (CLAUDE.md invariant 6). The keep-alive alarm used 0.4 min,
// below Chrome's 30 s minimum for packed extensions.

import { beforeEach, describe, expect, it, vi } from 'vitest'
import { startInitialScan } from '@/background/initial-scan'
import { holdKeepAlive, KEEPALIVE_PERIOD_MINUTES, releaseKeepAlive } from '@/background/keep-alive'
import type { OutlookApi } from '@/shared/outlook-api'
import type { MailFolderNode } from '@/shared/types'

let alarmsCreate: ReturnType<typeof vi.fn>
beforeEach(() => {
  alarmsCreate = vi.fn(async () => {})
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.alarms = {
    create: alarmsCreate,
    clear: vi.fn(async () => true),
    get: vi.fn(async () => undefined),
    onAlarm: { addListener: vi.fn() },
  }
})

const TREE: MailFolderNode[] = [
  {
    id: 'FOLDER-CASES-0000000000000',
    displayName: '案件',
    path: '案件',
    children: [
      { id: 'FOLDER-A-00000000000000000', displayName: '甲', path: '案件/甲', children: [] },
      { id: 'FOLDER-B-00000000000000000', displayName: '乙', path: '案件/乙', children: [] },
    ],
  },
]

describe('startInitialScan synchronous guard', () => {
  it('two overlapping starts → the second fails fast; folders are scanned once', async () => {
    const listFolderMessages = vi.fn(async () => [])
    const api = { listFolderMessages } as unknown as OutlookApi
    const opts = { rootPath: '案件', tree: TREE, api, excludePrefixes: [], internalDomains: [] }
    const first = startInitialScan(opts)
    const second = startInitialScan(opts)
    await expect(second).rejects.toThrow('已有掃描中的批次')
    await first
    expect(listFolderMessages).toHaveBeenCalledTimes(3) // 案件, 案件/甲, 案件/乙 — once each
    // Guard released afterwards: a later scan can run.
    await expect(startInitialScan(opts)).resolves.toBeUndefined()
  })
})

describe('keep-alive alarm', () => {
  it("uses Chrome's 30 s minimum period (0.4 min is rounded up by Chrome anyway)", async () => {
    expect(KEEPALIVE_PERIOD_MINUTES).toBeGreaterThanOrEqual(0.5)
    await holdKeepAlive()
    expect(alarmsCreate).toHaveBeenCalledWith(expect.any(String), { periodInMinutes: 0.5 })
    await releaseKeepAlive()
  })
})
