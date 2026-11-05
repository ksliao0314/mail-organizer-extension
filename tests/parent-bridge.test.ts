// Popup/options ↔ FAB content-script bridge handshake (2026-11-05 audit).
//
// The old bridge pinned its target origin from ANY `parent-hello` — no source
// check — so a page script that reached the iframe could re-pin it and
// black-hole close / open-options / navigate, and every message was posted to
// the OWA origin where all JS worlds on the page could read the folder paths.
// Now the content script hands over a private MessagePort; the frame accepts
// only the FIRST hello from window.parent that carries one.

import { describe, expect, it, vi } from 'vitest'
import { createParentBridge, type BridgeWindow } from '@/shared/parent-bridge'

type Listener = (e: MessageEvent) => void

function fakePort() {
  return { postMessage: vi.fn(), close: vi.fn() }
}

function fakeFrameWindow() {
  const parent = { name: 'owa-top' }
  let listener: Listener | null = null
  const win: BridgeWindow = {
    parent,
    addEventListener: vi.fn((_t: 'message', l: Listener) => {
      listener = l
    }),
    removeEventListener: vi.fn(),
  }
  const deliver = (data: unknown, source: unknown, ports: unknown[] = []) =>
    listener?.({ data, source, ports } as unknown as MessageEvent)
  return { win, parent, deliver }
}

const HELLO = { type: 'mail-organizer/parent-hello' }

describe('parent bridge handshake', () => {
  it('accepts the first hello from window.parent that carries a port, acks, and flushes the queue', () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const bridge = createParentBridge(win)
    bridge.init()
    // Posted before the handshake → queued, not lost.
    expect(bridge.post({ type: 'mail-organizer/prewarm-sidebar', paths: ['A/B'] })).toBe(false)

    const port = fakePort()
    deliver(HELLO, parent, [port])
    expect(bridge.isConnected()).toBe(true)
    expect(port.postMessage).toHaveBeenNthCalledWith(1, { type: 'mail-organizer/bridge-ack' })
    expect(port.postMessage).toHaveBeenNthCalledWith(2, {
      type: 'mail-organizer/prewarm-sidebar',
      paths: ['A/B'],
    })

    expect(bridge.post({ type: 'mail-organizer/close-panel' })).toBe(true)
    expect(port.postMessage).toHaveBeenLastCalledWith({ type: 'mail-organizer/close-panel' })
  })

  it('refuses a hello whose source is not window.parent (a page script reaching the frame)', () => {
    const { win, deliver } = fakeFrameWindow()
    const bridge = createParentBridge(win)
    bridge.init()
    const evil = fakePort()
    deliver(HELLO, { name: 'some-other-window' }, [evil])
    expect(bridge.isConnected()).toBe(false)
    expect(evil.close).toHaveBeenCalled()
    expect(evil.postMessage).not.toHaveBeenCalled()
  })

  it('refuses a portless hello (the old protocol) and any hello after pairing', () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const bridge = createParentBridge(win)
    bridge.init()
    deliver(HELLO, parent, [])
    expect(bridge.isConnected()).toBe(false)

    const first = fakePort()
    deliver(HELLO, parent, [first])
    const second = fakePort()
    deliver(HELLO, parent, [second]) // late re-offer / hijack attempt
    expect(second.close).toHaveBeenCalled()
    bridge.post({ type: 'mail-organizer/close-panel' })
    expect(first.postMessage).toHaveBeenLastCalledWith({ type: 'mail-organizer/close-panel' })
    expect(second.postMessage).not.toHaveBeenCalled()
  })

  it('queue:false reports "not delivered" without queueing (caller has a fallback)', () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const bridge = createParentBridge(win)
    expect(bridge.post({ type: 'mail-organizer/open-options' }, { queue: false })).toBe(false)
    const port = fakePort()
    deliver(HELLO, parent, [port])
    // Only the ack — the open-options must NOT fire late.
    expect(port.postMessage).toHaveBeenCalledTimes(1)
  })

  it('is a no-op at top level (toolbar popup / standalone options tab)', () => {
    const win = {
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
    } as unknown as BridgeWindow & { parent: unknown }
    ;(win as { parent: unknown }).parent = win
    const bridge = createParentBridge(win)
    bridge.init()
    expect(win.addEventListener).not.toHaveBeenCalled()
    expect(bridge.post({ type: 'mail-organizer/close-panel' })).toBe(false)
  })
})
