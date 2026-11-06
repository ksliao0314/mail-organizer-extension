// Framed popup / options UI waits for a verified bridge (2026-11-06 review
// U8 + U9).
//
// U9 — both pages are web-accessible to the Outlook origins (the FAB shows
// the popup in its panel and the options page in its settings modal — so
// options/index.html can't leave web_accessible_resources either), which lets
// an Outlook page script embed them as fully privileged extension pages and
// overlay them to steal clicks. Pairing only guarded the message channel.
// Framed, the UI now renders only after the SW-verified pairing completes;
// a page script can't complete it.
//
// U8 — a content script orphaned by an extension update can't pair (no
// runtime for a nonce; an older version's hellos carry none), and the panel's
// close / Esc / navigation silently did nothing. After the FAB's re-offer
// window the panel now says so.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { BridgeGate, PAIRING_NOTICE_AFTER_MS } from '@/components/ui/bridge-gate'
import { createParentBridge, type BridgeWindow } from '@/shared/parent-bridge'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

type Listener = (e: MessageEvent) => void
const STALE = '與 Outlook 分頁的連線已失效（擴充功能已更新），請重新整理 Outlook 分頁'

function frame(embedded = true) {
  const parent = { name: 'owa-top' }
  let listener: Listener | null = null
  const win: BridgeWindow = {
    parent: null,
    addEventListener: (_t: 'message', l: Listener) => {
      listener = l
    },
    removeEventListener: () => {},
  }
  ;(win as { parent: unknown }).parent = embedded ? parent : win
  const port = () => ({ postMessage: vi.fn(), close: vi.fn() })
  const hello = (nonce?: string, source: unknown = parent) =>
    listener?.({
      data: nonce === undefined ? { type: 'mail-organizer/parent-hello' } : { type: 'mail-organizer/parent-hello', nonce },
      source,
      ports: [port()],
    } as unknown as MessageEvent)
  return { win, hello }
}

let root: Root | null = null
afterEach(() => {
  act(() => root?.unmount())
  root = null
  document.body.innerHTML = ''
  vi.useRealTimers()
})

async function render(bridge: ReturnType<typeof createParentBridge>) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  await act(async () => {
    root = createRoot(host)
    root.render(createElement(BridgeGate, { bridge }, createElement('button', null, '執行 3 個')))
  })
  return host
}

const settle = () => act(async () => {
  await new Promise((r) => setTimeout(r, 0))
})

describe('BridgeGate', () => {
  it('top-level (toolbar popup / standalone tab): the UI renders at once', async () => {
    const { win } = frame(false)
    const bridge = createParentBridge(win, vi.fn(async () => false))
    bridge.init()
    const host = await render(bridge)
    expect(host.textContent).toContain('執行 3 個')
  })

  it('framed: nothing clickable until the SW verifies the pairing nonce', async () => {
    const { win, hello } = frame()
    const redeem = vi.fn(async (nonce: string) => nonce === 'n-good')
    const bridge = createParentBridge(win, redeem)
    bridge.init()
    const host = await render(bridge)
    expect(host.querySelector('button')).toBeNull()

    hello('n-forged') // e.g. a page script that embedded the panel itself
    await settle()
    expect(redeem).toHaveBeenCalledWith('n-forged')
    expect(host.querySelector('button')).toBeNull()

    hello('n-good') // the FAB content script
    await settle()
    expect(host.textContent).toContain('執行 3 個')
  })

  it('never paired (orphaned / pre-update content script): says to reload Outlook after the re-offer window', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const { win, hello } = frame()
    const bridge = createParentBridge(win, vi.fn(async () => true))
    bridge.init()
    const host = await render(bridge)
    hello() // an older content script's hello: no nonce → refused
    await act(async () => {
      await vi.advanceTimersByTimeAsync(1000)
    })
    expect(host.textContent).toContain('正在連線 Outlook 分頁')
    expect(host.textContent).not.toContain(STALE)

    await act(async () => {
      await vi.advanceTimersByTimeAsync(PAIRING_NOTICE_AFTER_MS)
    })
    expect(host.querySelector('[role="alert"]')?.textContent).toContain(STALE)
    expect(host.querySelector('button')).toBeNull()

    // A pairing that still completes later shows the UI after all.
    hello('n-late')
    await act(async () => {
      await vi.advanceTimersByTimeAsync(0)
    })
    expect(host.textContent).toContain('執行 3 個')
    expect(host.textContent).not.toContain(STALE)
  })
})

describe('parent bridge onConnected / isEmbedded', () => {
  it('fires once, after the verified hello only', async () => {
    const { win, hello } = frame()
    const bridge = createParentBridge(win, vi.fn(async (n: string) => n === 'ok'))
    bridge.init()
    const seen = vi.fn()
    const off = bridge.onConnected(seen)
    expect(bridge.isEmbedded()).toBe(true)
    hello('bad')
    await new Promise((r) => setTimeout(r, 0))
    expect(seen).not.toHaveBeenCalled()
    hello('ok')
    await new Promise((r) => setTimeout(r, 0))
    expect(seen).toHaveBeenCalledTimes(1)
    hello('ok') // already paired — refused, no second call
    await new Promise((r) => setTimeout(r, 0))
    expect(seen).toHaveBeenCalledTimes(1)
    off()
    expect(createParentBridge(frame(false).win, vi.fn()).isEmbedded()).toBe(false)
  })
})
