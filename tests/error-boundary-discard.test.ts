// Root error boundary: a way out of a crash loop (2026-11-06 audit E5).
//
// The fallback only offered 重新載入 — but a reload restores the same saved
// plan, so a crash caused by that plan's data came straight back, forever.
// The popup now also offers 「捨棄暫存計畫並重新載入」: clearPopupState +
// cancelClassify, then reload.

import { act, createElement, type ComponentType } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ErrorBoundary } from '@/components/ui/error-boundary'
import { discardSavedPlan } from '@/popup/crash-recovery'
import { buttonContaining, stubRuntime, type Msg, type Reply } from './helpers/ui-harness'

let root: Root | null = null
let consoleSpy: ReturnType<typeof vi.spyOn> | null = null
afterEach(() => {
  act(() => root?.unmount())
  root = null
  document.body.innerHTML = ''
  delete (globalThis as unknown as { chrome: { runtime?: unknown } }).chrome.runtime
  consoleSpy?.mockRestore()
  consoleSpy = null
})

function quietConsole() {
  // React logs every caught render error; the boundary logs it again.
  consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {})
}

async function render(node: ReturnType<typeof createElement>) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  await act(async () => {
    root = createRoot(host)
    root.render(node)
  })
  return host
}

function Boom(): never {
  throw new Error('render exploded')
}

describe('popup crash recovery', () => {
  it('a plan whose saved data crashes the popup stops crashing after 捨棄暫存計畫並重新載入', async () => {
    quietConsole()
    // A saved plan the popup can't render (corrupt snapshot) — every plain
    // reload restores it and crashes again.
    let saved: unknown = {
      phase: 'plan',
      updatedAt: Date.now(),
      plan: { items: null, folderTree: [], usage: null, summary: { ruleHits: 0, aiHandled: 0 }, banner: null },
    }
    const sw = stubRuntime((m: Msg): Reply | undefined => {
      switch (m.type) {
        case 'getStatus':
          return {
            ok: true,
            data: {
              owaConnected: true,
              tokenValid: true,
              apiKeyConfigured: true,
              apiKeyPreview: 'sk-…',
              model: 'claude-sonnet-4-6',
              folderCacheAgeMs: null,
              folderCacheCount: 0,
              rulesCount: 0,
              excludePrefixes: [],
              batchSize: 50,
              internalDomains: [],
              primaryRootPath: '',
            },
          }
        case 'getOnboardingState':
          return { ok: true, data: { needed: false } }
        case 'getFolderActivity':
          return { ok: true, data: { activity: [] } }
        case 'getExecuteState':
          return { ok: true, data: { state: null } }
        case 'getAiClassifyProgress':
          return { ok: true, data: null }
        case 'getPopupState':
          return { ok: true, data: saved }
        case 'clearPopupState':
          saved = null
          return { ok: true }
        case 'cancelClassify':
          return { ok: true, data: { cancelled: true } }
        default:
          return undefined
      }
    })
    const { default: App } = await import('@/popup/App')
    let generation = 0
    const reload = vi.fn(() => {
      // What a page reload does: a fresh tree.
      act(() => root!.render(tree(++generation)))
    })
    const tree = (key: number) =>
      createElement(
        ErrorBoundary,
        { key, onDiscardState: discardSavedPlan, reload },
        createElement(App as ComponentType),
      )

    const host = await render(tree(generation))
    await vi.waitFor(() => expect(host.textContent).toContain('畫面發生錯誤'))

    // A plain reload restores the same plan → same crash.
    act(() => buttonContaining(host, '重新載入')!.click())
    await vi.waitFor(() => expect(reload).toHaveBeenCalledTimes(1))
    await vi.waitFor(() => expect(host.textContent).toContain('畫面發生錯誤'))

    const discard = await vi.waitFor(() => buttonContaining(host, '捨棄暫存計畫並重新載入')!)
    await act(async () => discard.click())
    await vi.waitFor(() => expect(reload).toHaveBeenCalledTimes(2))
    expect(sw.count('clearPopupState')).toBe(1)
    expect(sw.last('cancelClassify')).toEqual({ type: 'cancelClassify' })
    // Discard ran BEFORE the reload: the fresh tree renders the idle screen.
    await vi.waitFor(() => expect(host.textContent).toContain('開始歸類'))
    expect(host.textContent).not.toContain('畫面發生錯誤')
  })

  it('a FINISHED batch whose done screen crashes stops crashing after the discard (2026-11-06 U6)', async () => {
    quietConsole()
    // The popup reopens onto a finished batch's DoneScreen BEFORE it looks at
    // the saved plan — discarding only the plan left this one looping.
    let executeState: unknown = {
      inProgress: false,
      cancelRequested: false,
      startedAt: Date.now() - 5000,
      finishedAt: Date.now(),
      total: 1,
      current: 1,
      results: null, // data the done screen can't render
      summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    const sw = stubRuntime((m: Msg): Reply | undefined => {
      switch (m.type) {
        case 'getStatus':
          return {
            ok: true,
            data: {
              owaConnected: true,
              tokenValid: true,
              apiKeyConfigured: true,
              model: 'claude-sonnet-4-6',
              folderCacheCount: 0,
              excludePrefixes: [],
              batchSize: 50,
              internalDomains: [],
              primaryRootPath: '',
            },
          }
        case 'getOnboardingState':
          return { ok: true, data: { needed: false } }
        case 'getFolderActivity':
          return { ok: true, data: { activity: [] } }
        case 'getExecuteState':
          return { ok: true, data: { state: executeState } }
        case 'clearExecuteState':
          executeState = null
          return { ok: true, data: { cleared: true } }
        case 'getAiClassifyProgress':
        case 'getPopupState':
          return { ok: true, data: null }
        case 'clearPopupState':
        case 'dismissUndo':
          return { ok: true }
        case 'cancelClassify':
          return { ok: true, data: { cancelled: false } }
        default:
          return undefined
      }
    })
    const { default: App } = await import('@/popup/App')
    let generation = 0
    const reload = vi.fn(() => act(() => root!.render(tree(++generation))))
    const tree = (key: number) =>
      createElement(ErrorBoundary, { key, onDiscardState: discardSavedPlan, reload }, createElement(App as ComponentType))

    const host = await render(tree(generation))
    await vi.waitFor(() => expect(host.textContent).toContain('畫面發生錯誤'))
    await act(async () => buttonContaining(host, '捨棄暫存計畫並重新載入')!.click())
    await vi.waitFor(() => expect(reload).toHaveBeenCalledTimes(1))
    expect(sw.last('clearExecuteState')).toEqual({ type: 'clearExecuteState', unlessRunning: true })
    expect(sw.count('dismissUndo')).toBe(1)
    await vi.waitFor(() => expect(host.textContent).toContain('開始歸類'))
    expect(host.textContent).not.toContain('畫面發生錯誤')
  })

  it('never discards a RUNNING batch: the SW refuses and its undo is left alone', async () => {
    const sw = stubRuntime((m: Msg): Reply | undefined => {
      if (m.type === 'clearExecuteState') {
        return m.unlessRunning === true ? { ok: true, data: { cleared: false } } : { ok: true, data: { cleared: true } }
      }
      if (m.type === 'clearPopupState' || m.type === 'dismissUndo') return { ok: true }
      if (m.type === 'cancelClassify') return { ok: true, data: { cancelled: true } }
      return undefined
    })
    await discardSavedPlan()
    expect(sw.last('clearExecuteState')).toEqual({ type: 'clearExecuteState', unlessRunning: true })
    expect(sw.count('dismissUndo')).toBe(0)
    expect(sw.count('clearPopupState')).toBe(1)
  })

  it('a hung service worker does not trap the user — the reload still happens', async () => {
    quietConsole()
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    try {
      const reload = vi.fn()
      const never = vi.fn(() => new Promise<void>(() => {}))
      const host = await render(
        createElement(ErrorBoundary, { onDiscardState: never, reload }, createElement(Boom)),
      )
      const discard = buttonContaining(host, '捨棄暫存計畫並重新載入')!
      await act(async () => discard.click())
      expect(never).toHaveBeenCalledTimes(1)
      expect(reload).not.toHaveBeenCalled()
      expect(buttonContaining(host, '正在捨棄')!.disabled).toBe(true) // no double submit
      await act(async () => {
        await vi.advanceTimersByTimeAsync(3000)
      })
      expect(reload).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('the options page (no onDiscardState) keeps the single reload button', async () => {
    quietConsole()
    const host = await render(createElement(ErrorBoundary, null, createElement(Boom)))
    expect(host.textContent).toContain('畫面發生錯誤')
    expect(buttonContaining(host, '捨棄暫存計畫')).toBeUndefined()
    expect(host.querySelectorAll('button')).toHaveLength(1)
  })
})
