// DoneScreen 「回主畫面」(finishAndReset) only ends THIS surface's work
// (2026-11-06 audit E6).
//
// It used to send cancelClassify with NO runId — "cancel whatever run is
// current" — plus unconditional clearExecuteState / clearPopupState. With the
// toolbar popup and the in-OWA panel open at once, a stale DoneScreen in one
// of them, clicked after the other had consumed (繼續歸檔 adopts the SAME run
// id) or started a run, killed that live run: the other surface reopened with
// no banner and the rest of its mail silently missing. Now: cancel only the
// runs this surface owns, by id, and touch shared state only while the
// DoneScreen's batch is still the current execute state.
//
// Round 2 (2026-11-06 review U3 / U4): the DoneScreen's OWN prefetch run is
// cancelled even when the other surface has moved on — `unlessAdopted`, so a
// run 繼續歸檔 took over elsewhere survives (the SW side is covered in
// classify-run-adoption.test.ts) — and a prefetch preflight that lands while
// 回主畫面 is awaiting the execute state never goes on to classifyAi.

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

const preflight = (runId: string) => ({
  runId,
  rulePlan: [],
  unmatchedPreview: [{ id: 'n1', subject: '下一批', from: 'x@client.example', fromName: '', received: '' }],
  folderTree: [],
  aiConfidenceThreshold: 0.5,
  preFilteredCount: 0,
  flaggedCount: 0,
  requestedBatchSize: 50,
})

describe('DoneScreen 回主畫面', () => {
  it('never sends a run-less cancelClassify (that cancelled whatever run was current)', async () => {
    const sent = await mountPopup(() => undefined)
    await clickHome()
    expect(sent.filter((m) => m.type === 'cancelClassify' && m.runId === undefined)).toEqual([])
    // Its own batch is still current → the shared batch / plan state is cleared as before.
    expect(sent.some((m) => m.type === 'clearExecuteState')).toBe(true)
    expect(sent.some((m) => m.type === 'clearPopupState')).toBe(true)
  })

  it("cancels this DoneScreen's own prefetch run — by id, unless another surface adopted it", async () => {
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getStatus') return status({ prefetchNextBatch: true })
      if (msg.type === 'classifyPreflight') return preflight('run-p')
      if (msg.type === 'classifyAi') return { started: true, runId: 'run-p', chunks: 1, totalEmails: 1 }
      return undefined
    })
    await vi.waitFor(() => expect(sent.some((m) => m.type === 'classifyAi')).toBe(true))
    await clickHome()
    expect(sent.filter((m) => m.type === 'cancelClassify')).toEqual([
      { type: 'cancelClassify', runId: 'run-p', unlessAdopted: true },
    ])
  })

  it('a stale DoneScreen (the other surface already moved on) leaves the shared state alone — its prefetch run goes only if not adopted', async () => {
    // This surface prefetched run-p; then the toolbar popup clicked 繼續歸檔,
    // which cleared the execute state and ADOPTED run-p — same id. A plain
    // cancel of "our" run id here would kill the other surface's live plan;
    // `unlessAdopted` lets the SW spare it. Had the toolbar popup gone HOME
    // instead (U4), nothing adopted run-p and this cancel is what stops it.
    let otherSurfaceMovedOn = false
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getStatus') return status({ prefetchNextBatch: true })
      if (msg.type === 'classifyPreflight') return preflight('run-p')
      if (msg.type === 'classifyAi') return { started: true, runId: 'run-p', chunks: 1, totalEmails: 1 }
      if (msg.type === 'getExecuteState') return { state: otherSurfaceMovedOn ? null : doneExecute }
      return undefined
    })
    await vi.waitFor(() => expect(sent.some((m) => m.type === 'classifyAi')).toBe(true))
    otherSurfaceMovedOn = true
    const before = sent.length

    await clickHome()
    const after = sent.slice(before)
    expect(after.filter((m) => m.type === 'cancelClassify')).toEqual([
      { type: 'cancelClassify', runId: 'run-p', unlessAdopted: true },
    ])
    expect(after.some((m) => m.type === 'clearExecuteState')).toBe(false)
    expect(after.some((m) => m.type === 'clearPopupState')).toBe(false)
  })

  it('a prefetch preflight landing while 回主畫面 awaits the execute state never starts classifyAi — its run is dropped', async () => {
    let landPreflight: (() => void) | null = null
    let answerExecState: (() => void) | null = null
    let homeClicked = false
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getStatus') return status({ prefetchNextBatch: true })
      if (msg.type === 'classifyPreflight') {
        return new Promise((resolve) => {
          landPreflight = () => resolve(preflight('run-late'))
        })
      }
      if (msg.type === 'getExecuteState' && homeClicked) {
        return new Promise((resolve) => {
          answerExecState = () => resolve({ state: doneExecute })
        })
      }
      if (msg.type === 'classifyAi') return { started: true, runId: 'run-late', chunks: 1, totalEmails: 1 }
      return undefined
    })
    await vi.waitFor(() => expect(landPreflight).not.toBeNull())
    await vi.waitFor(() => expect(button('回主畫面')).toBeDefined())

    // 回主畫面 while the prefetch preflight is still out…
    homeClicked = true
    await act(async () => button('回主畫面')!.click())
    await vi.waitFor(() => expect(answerExecState).not.toBeNull())
    // …and it lands during 回主畫面's getExecuteState await.
    await act(async () => landPreflight!())
    await act(async () => answerExecState!())
    await vi.waitFor(() => expect(text()).toContain('開始歸類'))
    await new Promise((r) => setTimeout(r, 20))

    expect(sent.some((m) => m.type === 'classifyAi')).toBe(false)
    expect(sent).toContainEqual({ type: 'cancelClassify', runId: 'run-late', unlessAdopted: true })
  })

  it('繼續歸檔 tells the SW it adopted the run it consumes, before dropping the batch', async () => {
    const offerable = {
      runId: 'run-p',
      totalEmails: 1,
      completedEmails: 1,
      chunks: 1,
      completedChunks: 1,
      ruleCount: 0,
      aiCount: 1,
      aiUnresolvedCount: 0,
      aiUnclassifiedCount: 0,
      aiCursor: 1,
      startedAt: EXEC_STARTED + 1000,
      done: true,
      rulePlan: [],
      aiPlan: [{ emailId: 'n1', emailSubject: '下一批', emailFrom: 'x@client.example', action: 'skip', confidence: 0.9, reason: '', source: 'ai' }],
      folderTree: [],
      aiError: null,
    }
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getAiClassifyProgress') return offerable
      return undefined
    })
    await vi.waitFor(() => expect(button('繼續歸檔下一批')).toBeDefined())
    await act(async () => button('繼續歸檔下一批')!.click())
    await vi.waitFor(() => expect(text()).toContain('下一批'))
    const adopt = sent.findIndex((m) => m.type === 'adoptClassifyRun')
    expect(sent[adopt]).toEqual({ type: 'adoptClassifyRun', runId: 'run-p' })
    expect(adopt).toBeLessThan(sent.findIndex((m) => m.type === 'clearExecuteState'))
  })

  it('a newer batch from another surface also counts as moved on', async () => {
    let newer = false
    const sent = await mountPopup((msg) => {
      if (msg.type === 'getExecuteState') {
        return { state: newer ? { ...doneExecute, startedAt: EXEC_STARTED + 60_000 } : doneExecute }
      }
      return undefined
    })
    await vi.waitFor(() => expect(button('回主畫面')).toBeDefined())
    newer = true
    const before = sent.length
    await clickHome()
    const after = sent.slice(before)
    expect(after.some((m) => m.type === 'clearExecuteState')).toBe(false) // not ours to clear
    expect(after.some((m) => m.type === 'cancelClassify')).toBe(false)
  })
})
