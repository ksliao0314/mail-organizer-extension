// FAB content script ↔ iframe bridge (2026-11-05 audit).
//
//   - The shadow root is CLOSED (page scripts can't walk into it to reach the
//     iframes and inject a forged handshake).
//   - On iframe load the content script offers a private MessagePort in
//     `parent-hello`, posted to the frame with targetOrigin = the extension
//     origin; commands are then accepted ONLY over that port — a
//     window.postMessage from the page (even one claiming our origin) does
//     nothing.
//   - Extension URLs are resolved once; after an extension update (orphaned
//     content script) the FAB tells the user to reload instead of opening a
//     half-working panel.
//   - (2026-11-06) every hello carries a fresh one-time nonce from the SW
//     ('issuePanelPairingNonce'), which the frame redeems before trusting the
//     port — no nonce, no hello.

import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'

const EXT = 'chrome-extension://test-extension'
let shadow: ShadowRoot | null = null
let host: HTMLElement | null = null
let nonceSeq = 0
let nonceAvailable = true
const runtime: { id: string | undefined } & Record<string, unknown> = {
  id: 'test-extension',
  getURL: (p: string) => `${EXT}/${p}`,
  sendMessage: vi.fn(async (msg: { type?: string }) => {
    if (msg?.type === 'issuePanelPairingNonce') {
      return nonceAvailable
        ? { ok: true, data: { nonce: `nonce-${++nonceSeq}` } }
        : { ok: false, code: 'NO_TAB', message: 'no tab' }
    }
    return { ok: true, data: { show: true } }
  }),
  onMessage: { addListener: vi.fn() },
}
const realAttachShadow = Element.prototype.attachShadow

beforeAll(async () => {
  // Capture the closed root the same way a test harness would — the page
  // itself has no such hook.
  Element.prototype.attachShadow = function (init: ShadowRootInit) {
    const root = realAttachShadow.call(this, init)
    shadow = root
    return root
  }
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.runtime = runtime
  await import('@/content/owa-fab')
  await vi.waitFor(() => expect(document.getElementById('mail-organizer-fab-host')).not.toBeNull())
  host = document.getElementById('mail-organizer-fab-host')
})

afterAll(() => {
  // Tear the FAB down through its own visibility hook so its observers /
  // timers don't fire after the jsdom environment is gone.
  const onMessage = runtime.onMessage as { addListener: ReturnType<typeof vi.fn> }
  const listener = onMessage.addListener.mock.calls[0]?.[0] as
    | ((msg: { type: string; show: boolean }) => void)
    | undefined
  listener?.({ type: 'owaFabVisibility', show: false })
  Element.prototype.attachShadow = realAttachShadow
  delete (globalThis as unknown as { chrome: { runtime?: unknown } }).chrome.runtime
})

function parts() {
  const root = shadow!
  return {
    fab: root.querySelector<HTMLButtonElement>('.fab')!,
    backdrop: root.querySelector<HTMLDivElement>('.panel-backdrop')!,
    iframe: root.querySelector<HTMLIFrameElement>('.iframe')!,
  }
}

describe('OWA FAB bridge', () => {
  it('uses a closed shadow root', () => {
    expect(host).not.toBeNull()
    expect(host!.shadowRoot).toBeNull()
    expect(shadow).not.toBeNull()
  })

  it('pairs over a private MessagePort and accepts commands only through it', async () => {
    const { fab, backdrop, iframe } = parts()
    fab.click()
    expect(backdrop.classList.contains('open')).toBe(true)
    expect(iframe.src).toBe(`${EXT}/src/popup/index.html`)

    const postMessage = vi.fn()
    Object.defineProperty(iframe, 'contentWindow', { value: { postMessage }, configurable: true })
    iframe.dispatchEvent(new Event('load'))

    await vi.waitFor(() => expect(postMessage).toHaveBeenCalledTimes(1))
    const [msg, opts] = postMessage.mock.calls[0]!
    // The hello carries the nonce the SW just issued to this content script.
    expect(runtime.sendMessage).toHaveBeenCalledWith({ type: 'issuePanelPairingNonce' })
    expect(msg).toEqual({ type: 'mail-organizer/parent-hello', nonce: `nonce-${nonceSeq}` })
    expect(opts.targetOrigin).toBe(EXT)
    const port = opts.transfer[0] as MessagePort
    expect(port).toBeInstanceOf(MessagePort)

    // A page script spoofing a command via window.postMessage — even with
    // our origin — must be ignored now.
    window.dispatchEvent(
      new MessageEvent('message', { data: { type: 'mail-organizer/close-panel' }, origin: EXT }),
    )
    await new Promise((r) => setTimeout(r, 20))
    expect(backdrop.classList.contains('open')).toBe(true)

    // The real frame: ack, then a command over the port.
    port.postMessage({ type: 'mail-organizer/bridge-ack' })
    port.postMessage({ type: 'mail-organizer/close-panel' })
    await vi.waitFor(() => expect(backdrop.classList.contains('open')).toBe(false))
    port.close()
  })

  it('posts no hello without an SW nonce, and every re-offer carries a fresh one', async () => {
    const { fab, backdrop, iframe } = parts()
    fab.click()
    expect(backdrop.classList.contains('open')).toBe(true)
    const postMessage = vi.fn()
    Object.defineProperty(iframe, 'contentWindow', { value: { postMessage }, configurable: true })
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    try {
      nonceAvailable = false
      iframe.dispatchEvent(new Event('load'))
      await vi.advanceTimersByTimeAsync(100)
      expect(postMessage).not.toHaveBeenCalled() // SW said no → nothing offered

      nonceAvailable = true
      await vi.advanceTimersByTimeAsync(100) // first retry (150 ms)
      expect(postMessage).toHaveBeenCalledTimes(1)
      await vi.advanceTimersByTimeAsync(400) // no ack yet → second retry
      expect(postMessage).toHaveBeenCalledTimes(2)
      const nonces = postMessage.mock.calls.map(([m]) => (m as { nonce: string }).nonce)
      expect(nonces[0]).toMatch(/^nonce-/)
      expect(nonces[1]).toMatch(/^nonce-/)
      expect(nonces[0]).not.toBe(nonces[1])

      // The frame acks the second offer; retries stop and its port carries commands.
      const port = postMessage.mock.calls[1]![1].transfer[0] as MessagePort
      port.postMessage({ type: 'mail-organizer/bridge-ack' })
      port.postMessage({ type: 'mail-organizer/close-panel' })
      vi.useRealTimers()
      await vi.waitFor(() => expect(backdrop.classList.contains('open')).toBe(false))
      port.close()
    } finally {
      vi.useRealTimers()
      nonceAvailable = true
    }
  })

  it('after an extension update, a FAB click asks for an Outlook reload instead of opening', () => {
    const { fab, backdrop } = parts()
    expect(backdrop.classList.contains('open')).toBe(false)
    runtime.id = undefined // context invalidated
    try {
      fab.click()
      expect(backdrop.classList.contains('open')).toBe(false)
      expect(shadow!.querySelector('.nav-toast')?.textContent).toContain('重新整理')
    } finally {
      runtime.id = 'test-extension'
    }
  })
})
