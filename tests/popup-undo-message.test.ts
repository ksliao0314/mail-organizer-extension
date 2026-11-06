// 撤回 stopped by a token outage (2026-11-06 r2): every row carries the same
// shared reason, and the result also has it as `message`. Shown once at the
// top, not repeated after every subject.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true
if (!(globalThis as { ResizeObserver?: unknown }).ResizeObserver) {
  ;(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
    observe() {}
    disconnect() {}
  }
}

type Msg = { type: string; [k: string]: unknown }

const EXEC_STARTED = Date.UTC(2026, 10, 6, 9, 0, 0)
const doneExecute = {
  inProgress: false,
  cancelRequested: false,
  startedAt: EXEC_STARTED,
  finishedAt: EXEC_STARTED + 4000,
  total: 1,
  current: 1,
  results: [{ emailId: 'm0', subject: '規則命中', action: 'move', status: 'moved' }],
  summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
  rulesAdded: 0,
}

function status(over: Record<string, unknown> = {}) {
  return {
    owaConnected: true,
    tokenValid: true,
    apiKeyConfigured: true,
    apiKeyPreview: 'sk-ant-…',
    model: 'claude-sonnet-5-5',
    folderCacheAgeMs: 1000,
    folderCacheCount: 1,
    rulesCount: 0,
    excludePrefixes: [],
    batchSize: 50,
    prefetchNextBatch: false,
    internalDomains: [],
    primaryRootPath: '',
    ...over,
  }
}

let root: Root | null = null
afterEach(() => {
  act(() => root?.unmount())
  root = null
  document.body.innerHTML = ''
  delete (globalThis as unknown as { chrome: { runtime?: unknown } }).chrome.runtime
  vi.resetModules()
})

async function mountPopup(handler: (msg: Msg) => unknown | Promise<unknown>): Promise<Msg[]> {
  const sent: Msg[] = []
  const defaults = (msg: Msg): unknown => {
    switch (msg.type) {
      case 'getStatus':
        return status()
      case 'getFolderActivity':
        return { activity: [] }
      case 'getWeeklyDigest':
        return { shouldShow: false }
      case 'getOnboardingState':
        return { needed: false }
      case 'getUndoSnapshot':
        return { snapshot: null }
      case 'peekNextBatch':
        return { eligibleCount: 2, cappedAtBatchSize: false }
      case 'getPopupState':
        return null
      case 'getExecuteState':
        return { state: doneExecute }
      case 'getAiClassifyProgress':
        return null
      default:
        return undefined
    }
  }
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.runtime = {
    id: 'test-extension',
    getURL: (p: string) => `chrome-extension://test-extension/${p}`,
    openOptionsPage: vi.fn(),
    sendMessage: vi.fn(async (msg: Msg) => {
      sent.push(msg)
      const custom = await handler(msg) // a handler may hold an answer back (race tests)
      return { ok: true, data: custom === undefined ? defaults(msg) : custom }
    }),
  }
  const { default: App } = await import('@/popup/App')
  const host = document.createElement('div')
  document.body.appendChild(host)
  await act(async () => {
    root = createRoot(host)
    root.render(createElement(App))
  })
  return sent
}

const text = () => document.body.textContent ?? ''
const button = (label: string) =>
  [...document.querySelectorAll('button')].find((b) => (b.textContent ?? '').trim().startsWith(label))

async function clickHome() {
  await vi.waitFor(() => expect(button('回主畫面')).toBeDefined())
  await act(async () => button('回主畫面')!.click())
  await vi.waitFor(() => expect(text()).toContain('開始歸類'))
}


describe('undo stopped by a token outage', () => {
  it('shows the shared reason once, then the subjects', async () => {
    const reason = '無法取得 Outlook 授權,撤回已停止。撤回無法重試,請到 Outlook 手動移回收件匣。'
    await mountPopup((msg) => {
      if (msg.type === 'getUndoSnapshot') {
        return {
          snapshot: {
            batchStartedAt: EXEC_STARTED,
            expiresAt: Date.now() + 30_000,
            moves: [
              { emailId: 'a', subject: '甲案通知', newMessageId: 'na', fromFolderId: 'inbox', toFolderId: 'F1' },
              { emailId: 'b', subject: '乙案通知', newMessageId: 'nb', fromFolderId: 'inbox', toFolderId: 'F1' },
            ],
          },
        }
      }
      if (msg.type === 'executeUndo') {
        return {
          attempted: 2,
          restored: 0,
          failed: 2,
          message: reason,
          errors: [
            { subject: '甲案通知', message: reason },
            { subject: '乙案通知', message: reason },
          ],
        }
      }
      return undefined
    })
    await vi.waitFor(() => expect(button('撤回')).toBeDefined())
    await act(async () => button('撤回')!.click())
    await vi.waitFor(() => expect(text()).toContain('撤回部分失敗'))
    expect(text().split(reason).length - 1).toBe(1)
    expect(text()).toContain('甲案通知')
    expect(text()).toContain('乙案通知')
  })
})
