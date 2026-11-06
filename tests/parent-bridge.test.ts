// Popup/options ↔ FAB content-script bridge handshake (2026-11-05 audit).
//
// The old bridge pinned its target origin from ANY `parent-hello` — no source
// check — so a page script that reached the iframe could re-pin it and
// black-hole close / open-options / navigate, and every message was posted to
// the OWA origin where all JS worlds on the page could read the folder paths.
// Now the content script hands over a private MessagePort; the frame accepts
// only the FIRST hello from window.parent that carries one.
//
// 2026-11-06: `event.source === window.parent` can't tell the content script
// from an Outlook page script (same window), so every hello also carries a
// one-time nonce the SW issued to the content script of this tab; the port is
// accepted only once the SW confirms (and consumes) it. Nothing — not even
// the ack or the queued folder paths — goes out before that.

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

/** A stand-in for the SW: nonces it "issued" redeem once, everything else fails. */
function fakeSw(issued: string[] = ['n-good']) {
  const live = new Set(issued)
  const redeem = vi.fn(async (nonce: string) => live.delete(nonce))
  return { redeem }
}

const hello = (nonce?: string) =>
  nonce === undefined
    ? { type: 'mail-organizer/parent-hello' }
    : { type: 'mail-organizer/parent-hello', nonce }

const settle = () => new Promise((r) => setTimeout(r, 0))

describe('parent bridge handshake', () => {
  it('accepts a hello from window.parent only after the SW confirms its nonce, then acks and flushes', async () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const sw = fakeSw()
    const bridge = createParentBridge(win, sw.redeem)
    bridge.init()
    // Posted before the handshake → queued, not lost.
    expect(bridge.post({ type: 'mail-organizer/prewarm-sidebar', paths: ['A/B'] })).toBe(false)

    const port = fakePort()
    deliver(hello('n-good'), parent, [port])
    // Nothing goes out while the nonce is being checked.
    expect(port.postMessage).not.toHaveBeenCalled()
    expect(bridge.isConnected()).toBe(false)

    await settle()
    expect(sw.redeem).toHaveBeenCalledWith('n-good')
    expect(bridge.isConnected()).toBe(true)
    expect(port.postMessage).toHaveBeenNthCalledWith(1, { type: 'mail-organizer/bridge-ack' })
    expect(port.postMessage).toHaveBeenNthCalledWith(2, {
      type: 'mail-organizer/prewarm-sidebar',
      paths: ['A/B'],
    })

    expect(bridge.post({ type: 'mail-organizer/close-panel' })).toBe(true)
    expect(port.postMessage).toHaveBeenLastCalledWith({ type: 'mail-organizer/close-panel' })
  })

  it('refuses a hello from window.parent whose nonce the SW does not confirm (an Outlook page script)', async () => {
    // A page script shares window with the content script, so its hello has
    // the very same `source` — and it may arrive FIRST (the real hello waits
    // for the iframe load event). Its nonce is a guess; the SW says no.
    const { win, parent, deliver } = fakeFrameWindow()
    const sw = fakeSw(['n-real'])
    const bridge = createParentBridge(win, sw.redeem)
    bridge.init()
    bridge.post({ type: 'mail-organizer/prewarm-sidebar', paths: ['Secret/Folder'] })

    const evil = fakePort()
    deliver(hello('n-guess'), parent, [evil])
    await settle()
    expect(bridge.isConnected()).toBe(false)
    expect(evil.close).toHaveBeenCalled()
    expect(evil.postMessage).not.toHaveBeenCalled() // no ack, no folder paths

    // The real content script's hello still pairs afterwards.
    const real = fakePort()
    deliver(hello('n-real'), parent, [real])
    await settle()
    expect(bridge.isConnected()).toBe(true)
    expect(real.postMessage).toHaveBeenNthCalledWith(2, {
      type: 'mail-organizer/prewarm-sidebar',
      paths: ['Secret/Folder'],
    })
  })

  it('a page that embeds the panel itself never pairs: no nonce, or a failing SW, means no port', async () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const redeem = vi.fn(async () => {
      throw new Error('Extension context invalidated')
    })
    const bridge = createParentBridge(win, redeem)
    bridge.init()
    bridge.post({ type: 'mail-organizer/prewarm-sidebar', paths: ['A/B'] })

    const portless = fakePort()
    deliver(hello(), parent, [portless]) // the pre-2026-11-06 protocol: no nonce
    deliver(hello(''), parent, [portless])
    deliver(hello('x'.repeat(500)), parent, [portless]) // junk, not worth a SW call
    deliver(hello(42 as unknown as string), parent, [portless])
    await settle()
    expect(redeem).not.toHaveBeenCalled()
    expect(portless.close).toHaveBeenCalled()

    const p = fakePort()
    deliver(hello('n-1'), parent, [p])
    await settle()
    expect(redeem).toHaveBeenCalledTimes(1)
    expect(bridge.isConnected()).toBe(false)
    expect(p.close).toHaveBeenCalled()
    expect(p.postMessage).not.toHaveBeenCalled()
  })

  it('refuses a hello whose source is not window.parent, even with a valid nonce', async () => {
    const { win, deliver } = fakeFrameWindow()
    const sw = fakeSw()
    const bridge = createParentBridge(win, sw.redeem)
    bridge.init()
    const evil = fakePort()
    deliver(hello('n-good'), { name: 'some-other-window' }, [evil])
    await settle()
    expect(sw.redeem).not.toHaveBeenCalled()
    expect(bridge.isConnected()).toBe(false)
    expect(evil.close).toHaveBeenCalled()
    expect(evil.postMessage).not.toHaveBeenCalled()
  })

  it('refuses a portless hello and any hello after pairing; two verified hellos → first one wins', async () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const sw = fakeSw(['n-1', 'n-2', 'n-3'])
    const bridge = createParentBridge(win, sw.redeem)
    bridge.init()
    deliver(hello('n-1'), parent, [])
    await settle()
    expect(bridge.isConnected()).toBe(false)

    // Two offers in flight at once (the content script re-offers on a short
    // backoff while the first nonce is still being redeemed).
    const first = fakePort()
    const second = fakePort()
    deliver(hello('n-2'), parent, [first])
    deliver(hello('n-3'), parent, [second])
    await settle()
    expect(bridge.isConnected()).toBe(true)
    expect(second.close).toHaveBeenCalled()
    expect(second.postMessage).not.toHaveBeenCalled()

    const late = fakePort()
    deliver(hello('n-1'), parent, [late]) // late re-offer / hijack attempt
    await settle()
    expect(late.close).toHaveBeenCalled()
    bridge.post({ type: 'mail-organizer/close-panel' })
    expect(first.postMessage).toHaveBeenLastCalledWith({ type: 'mail-organizer/close-panel' })
    expect(late.postMessage).not.toHaveBeenCalled()
  })

  it('queue:false reports "not delivered" without queueing (caller has a fallback)', async () => {
    const { win, parent, deliver } = fakeFrameWindow()
    const sw = fakeSw()
    const bridge = createParentBridge(win, sw.redeem)
    expect(bridge.post({ type: 'mail-organizer/open-options' }, { queue: false })).toBe(false)
    const port = fakePort()
    deliver(hello('n-good'), parent, [port])
    await settle()
    // Only the ack — the open-options must NOT fire late.
    expect(port.postMessage).toHaveBeenCalledTimes(1)
  })

  it('is a no-op at top level (toolbar popup / standalone options tab)', () => {
    const win = {
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
    } as unknown as BridgeWindow & { parent: unknown }
    ;(win as { parent: unknown }).parent = win
    const redeem = vi.fn(async () => true)
    const bridge = createParentBridge(win, redeem)
    bridge.init()
    expect(win.addEventListener).not.toHaveBeenCalled()
    expect(bridge.post({ type: 'mail-organizer/close-panel' })).toBe(false)
    expect(redeem).not.toHaveBeenCalled()
  })
})
