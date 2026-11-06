// Start guards + keep-alive cadence (2026-11-05 audit, items 10 & 15).
//
// startInitialScan's guard flag was set AFTER `await isScanRunning()`, so two
// overlapping starts (double-click, FAB iframe + Options tab) both read false
// and both scanned (CLAUDE.md invariant 6). The keep-alive alarm used 0.4 min,
// below Chrome's 30 s minimum for packed extensions.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { startInitialScan } from '@/background/initial-scan'
import {
  holdKeepAlive,
  isKeepAliveHeartbeatActive,
  KEEPALIVE_HEARTBEAT_MS,
  KEEPALIVE_PERIOD_MINUTES,
  releaseKeepAlive,
} from '@/background/keep-alive'
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

// C6 (2026-11-06): the 30 s alarm floor is the idle window itself, and long
// waits (Claude back-offs up to 60 s, rate-limiter queues) make no extension
// API calls — the worker could be idle-killed mid-batch.
describe('keep-alive heartbeat (C6)', () => {
  let getPlatformInfo: ReturnType<typeof vi.fn>
  beforeEach(() => {
    vi.useFakeTimers()
    getPlatformInfo = vi.fn(async () => ({ os: 'mac' }))
    const c = (globalThis as unknown as { chrome: Record<string, unknown> }).chrome
    c.runtime = { ...(c.runtime as object), getPlatformInfo }
  })
  afterEach(() => vi.useRealTimers())

  it('calls a cheap extension API every 20 s while held, and stops — no timer left — on release', async () => {
    await holdKeepAlive()
    await holdKeepAlive() // nested holder
    expect(isKeepAliveHeartbeatActive()).toBe(true)
    await vi.advanceTimersByTimeAsync(KEEPALIVE_HEARTBEAT_MS * 3)
    expect(getPlatformInfo).toHaveBeenCalledTimes(3)
    await releaseKeepAlive()
    await vi.advanceTimersByTimeAsync(KEEPALIVE_HEARTBEAT_MS)
    expect(getPlatformInfo).toHaveBeenCalledTimes(4) // still one holder left
    await releaseKeepAlive()
    expect(isKeepAliveHeartbeatActive()).toBe(false)
    expect(vi.getTimerCount()).toBe(0)
    await vi.advanceTimersByTimeAsync(KEEPALIVE_HEARTBEAT_MS * 5)
    expect(getPlatformInfo).toHaveBeenCalledTimes(4)
  })

  it('a holder released while the alarm create is pending leaves no alarm behind', async () => {
    let finishCreate: () => void = () => {}
    alarmsCreate.mockImplementationOnce(() => new Promise<void>((r) => (finishCreate = r)))
    const clear = (chrome.alarms as unknown as { clear: ReturnType<typeof vi.fn> }).clear
    const hold = holdKeepAlive()
    await releaseKeepAlive()
    clear.mockClear()
    finishCreate()
    await hold
    expect(clear).toHaveBeenCalledTimes(1)
    expect(isKeepAliveHeartbeatActive()).toBe(false)
  })
})
