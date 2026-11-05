// Popup classify state machine (2026-11-05, items 1-3).
//
// Renders the real popup <App/> under jsdom against a stubbed service worker
// and checks what it SENDS:
//   - 繼續歸檔 consumes the run's remaining items by run (no re-classify of
//     mail whose AI finished during a partial execute) and a double click
//     runs one chain only;
//   - the DoneScreen prefetch leaves such a run alone;
//   - 取消 cancels the displayed run; progress of another run is ignored.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { PlanItem } from '@/shared/types'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true
if (!(globalThis as { ResizeObserver?: unknown }).ResizeObserver) {
  ;(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
    observe() {}
    disconnect() {}
  }
}
if (!Element.prototype.scrollIntoView) Element.prototype.scrollIntoView = () => {}

type Msg = { type: string; [k: string]: unknown }

const EXEC_STARTED = Date.UTC(2026, 10, 5, 9, 0, 0)

function item(id: string, subject: string): PlanItem {
  return {
    emailId: id,
    emailSubject: subject,
    emailFrom: `${id}@client.example`,
    action: 'skip',
    confidence: 0.9,
    reason: 'fixture',
    source: 'ai',
  }
}

function runProgress(over: Record<string, unknown> = {}) {
  const aiPlan = [item('m3', '第三封待歸檔'), item('m4', '第四封待歸檔')]
  return {
    runId: 'run-a',
    totalEmails: 2,
    completedEmails: 2,
    chunks: 1,
    completedChunks: 1,
    ruleCount: 0,
    aiCount: aiPlan.length,
    aiUnresolvedCount: 0,
    aiUnclassifiedCount: 0,
    aiCursor: aiPlan.length,
    // Classified BEFORE the execute started — the old startedAt heuristic
    // called this residue and re-classified it.
    startedAt: EXEC_STARTED - 60_000,
    done: true,
    rulePlan: [],
    aiPlan,
    folderTree: [],
    aiError: null,
    ...over,
  }
}

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

async function mountPopup(handler: (msg: Msg) => unknown): Promise<Msg[]> {
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
      case 'getFolderTree':
        return { tree: [], count: 0 }
      case 'getPopupState':
        return null
      case 'getExecuteState':
        return { state: null }
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
      const custom = handler(msg)
      const data = custom === undefined ? defaults(msg) : custom
      return { ok: true, data }
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

describe('DoneScreen → 繼續歸檔下一批', () => {
  it('consumes the run whose AI finished during the partial execute — no re-classify — and ignores a double click', async () => {
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getExecuteState') return { state: doneExecute }
      if (msg.type === 'getAiClassifyProgress') return runProgress()
      return undefined
    })
    await vi.waitFor(() => expect(button('繼續歸檔下一批')).toBeDefined())
    const btn = button('繼續歸檔下一批')!
    await act(async () => {
      btn.click()
      btn.click() // double-submit (item 3)
    })
    await vi.waitFor(() => expect(text()).toContain('第三封待歸檔'))
    expect(text()).toContain('第四封待歸檔')
    expect(sent.filter((m) => m.type === 'getAiClassifyProgress')).toHaveLength(1)
    expect(sent.some((m) => m.type === 'classifyPreflight')).toBe(false)
    expect(sent.some((m) => m.type === 'cancelClassify' || m.type === 'clearAiClassifyProgress')).toBe(false)
  })

  it('starts a fresh classify when the run has nothing left (everything executed)', async () => {
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getExecuteState') return { state: doneExecute }
      if (msg.type === 'getAiClassifyProgress') {
        return runProgress({ aiPlan: [], aiCount: 0, aiCursor: 2 })
      }
      if (msg.type === 'classifyPreflight') {
        return {
          runId: 'run-b',
          rulePlan: [],
          unmatchedPreview: [],
          folderTree: [],
          aiConfidenceThreshold: 0.5,
          preFilteredCount: 0,
          flaggedCount: 0,
          requestedBatchSize: 50,
        }
      }
      return undefined
    })
    await vi.waitFor(() => expect(button('繼續歸檔下一批')).toBeDefined())
    await act(async () => button('繼續歸檔下一批')!.click())
    await vi.waitFor(() => expect(sent.some((m) => m.type === 'classifyPreflight')).toBe(true))
    expect(sent.find((m) => m.type === 'cancelClassify')).toEqual({ type: 'cancelClassify', runId: 'run-a' })
  })
})

describe('DoneScreen → 繼續歸檔下一批 after a failed chunk', () => {
  it('re-classifies fresh when only 「AI 分類失敗」 placeholders are left', async () => {
    const failed = { ...item('m3', '分類失敗的信'), action: 'skip' as const, source: 'unresolved' as const, unclassified: true }
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getExecuteState') return { state: doneExecute }
      if (msg.type === 'getAiClassifyProgress') {
        return runProgress({ aiPlan: [failed], aiCount: 1, aiUnresolvedCount: 1, aiUnclassifiedCount: 1 })
      }
      if (msg.type === 'classifyPreflight') return { runId: 'run-b', rulePlan: [], unmatchedPreview: [], folderTree: [], aiConfidenceThreshold: 0.5, preFilteredCount: 0, flaggedCount: 0, requestedBatchSize: 50 }
      return undefined
    })
    await vi.waitFor(() => expect(button('繼續歸檔下一批')).toBeDefined())
    await act(async () => button('繼續歸檔下一批')!.click())
    await vi.waitFor(() => expect(sent.some((m) => m.type === 'classifyPreflight')).toBe(true))
  })
})

describe('DoneScreen prefetch', () => {
  it('leaves a finished-but-unconsumed run alone instead of re-classifying it', async () => {
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getStatus') return status({ prefetchNextBatch: true })
      if (msg.type === 'getExecuteState') return { state: doneExecute }
      if (msg.type === 'getAiClassifyProgress') return runProgress()
      return undefined
    })
    await vi.waitFor(() =>
      expect(sent.some((m) => m.type === 'getAiClassifyProgress' && m.view === 'meta')).toBe(true),
    )
    await new Promise((r) => setTimeout(r, 50))
    expect(sent.some((m) => m.type === 'classifyPreflight')).toBe(false)
    expect(sent.some((m) => m.type === 'cancelClassify' || m.type === 'clearAiClassifyProgress')).toBe(false)
  })
})

describe('plan with AI still running', () => {
  const running = () =>
    runProgress({
      runId: 'run-x',
      done: false,
      ruleCount: 1,
      rulePlan: [item('m0', '規則命中的信')],
      aiPlan: [],
      aiCount: 0,
      aiCursor: 0,
      completedEmails: 0,
      completedChunks: 0,
      startedAt: EXEC_STARTED,
    })

  it('取消 sends cancelClassify for the displayed run', async () => {
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getAiClassifyProgress') return running()
      return undefined
    })
    await vi.waitFor(() => expect(text()).toContain('規則命中的信'))
    await act(async () => button('取消')!.click())
    expect(sent).toContainEqual({ type: 'cancelClassify', runId: 'run-x' })
  })

  it('ignores progress of another run and stops waiting', async () => {
    await mountPopup((msg) => {
      if (msg.type === 'getAiClassifyProgress') {
        if (msg.view === 'poll') {
          // Another surface started run-y; its items must not leak in.
          return { ...running(), runId: 'run-y', aiPlan: [item('z9', '別的批次的信')], aiCount: 1, aiCursor: 1 }
        }
        return running()
      }
      return undefined
    })
    await vi.waitFor(() => expect(text()).toContain('規則命中的信'))
    await vi.waitFor(() => expect(text()).toContain('AI 分類已中止'), { timeout: 2000 })
    expect(text()).not.toContain('別的批次的信')
  })
})
