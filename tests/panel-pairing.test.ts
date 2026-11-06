// In-OWA panel bridge pairing anchored in chrome.runtime (2026-11-06 audit E4).
//
// The popup / options pages are web-accessible to Outlook, and Outlook page
// scripts share `window` with the FAB content script — so `event.source ===
// window.parent` couldn't tell a page script's `parent-hello` from ours. A page
// script could embed the popup itself and pair (receiving 近日活動 folder
// paths), or race its hello into the real panel before the content script's
// load-time hello. Now the SW issues the content script a one-time nonce bound
// to its tab, and the frame accepts a port only after redeeming that nonce
// from the same tab. Drives the REAL SW dispatcher + the real bridge.

import { beforeAll, describe, expect, it, vi } from 'vitest'
import { createParentBridge, type BridgeWindow } from '@/shared/parent-bridge'
import {
  PAIRING_NONCE_TTL_MS,
  issuePanelPairingNonce,
  redeemPanelPairingNonce,
} from '@/background/handlers/panel-pairing'
import {
  CONTENT_SCRIPT_SENDER,
  EXTENSION_PAGE_SENDER,
  loadServiceWorker,
  type Sender,
  type SwHarness,
} from './helpers/sw-harness'

let sw: SwHarness
beforeAll(async () => {
  sw = await loadServiceWorker()
})

/** The popup page inside the FAB iframe of tab `tabId`. */
function panelIn(tabId: number): Sender {
  return {
    id: 'test-extension',
    url: 'chrome-extension://test-extension/src/popup/index.html',
    origin: 'chrome-extension://test-extension',
    tab: { id: tabId } as chrome.tabs.Tab,
    frameId: 3,
  }
}

function contentScriptIn(tabId: number): Sender {
  return { ...CONTENT_SCRIPT_SENDER, tab: { id: tabId } as chrome.tabs.Tab }
}

async function issue(sender: Sender = CONTENT_SCRIPT_SENDER): Promise<string> {
  const r = await sw.sendFrom(sender, { type: 'issuePanelPairingNonce' })
  expect(r.ok).toBe(true)
  return (r.data as { nonce: string }).nonce
}

async function redeem(sender: Sender, nonce: unknown): Promise<boolean> {
  const r = await sw.sendFrom(sender, { type: 'redeemPanelPairingNonce', nonce })
  expect(r.ok).toBe(true)
  return (r.data as { valid: boolean }).valid
}

describe('pairing nonce handlers', () => {
  it('the FAB content script may ask for a nonce (it is on the content-script allowlist)', async () => {
    const nonce = await issue()
    expect(nonce).toMatch(/^[0-9a-f]{32}$/)
    expect(await issue()).not.toBe(nonce) // fresh every time
  })

  it('the panel in the SAME tab redeems it exactly once', async () => {
    const nonce = await issue(contentScriptIn(7))
    expect(await redeem(panelIn(7), nonce)).toBe(true)
    expect(await redeem(panelIn(7), nonce)).toBe(false) // consumed
  })

  it('a frame in another tab gets valid:false — and does not burn the real frame’s nonce', async () => {
    const nonce = await issue(contentScriptIn(7))
    expect(await redeem(panelIn(8), nonce)).toBe(false)
    expect(await redeem(panelIn(7), nonce)).toBe(true)
  })

  it('guesses, junk and tab-less senders never redeem', async () => {
    await issue(contentScriptIn(7))
    expect(await redeem(panelIn(7), 'not-a-nonce')).toBe(false)
    expect(await redeem(panelIn(7), 42)).toBe(false)
    expect(await redeem(panelIn(7), undefined)).toBe(false)
    const nonce = await issue(contentScriptIn(7))
    expect(await redeem(EXTENSION_PAGE_SENDER, nonce)).toBe(false) // no tab (standalone page)
    expect(await redeem(panelIn(7), nonce)).toBe(true)
  })

  it('a content-script sender cannot redeem (only extension pages may)', async () => {
    const nonce = await issue(contentScriptIn(7))
    expect(
      await sw.sendFrom(contentScriptIn(7), { type: 'redeemPanelPairingNonce', nonce }),
    ).toMatchObject({ ok: false, code: 'FORBIDDEN' })
    expect(await redeem(panelIn(7), nonce)).toBe(true) // the refused attempt consumed nothing
  })

  it('a sender without a tab gets no nonce', async () => {
    expect(await sw.sendFrom(EXTENSION_PAGE_SENDER, { type: 'issuePanelPairingNonce' })).toMatchObject({
      ok: false,
      code: 'NO_TAB',
    })
  })

  it('nonces expire', () => {
    const t0 = 1_000_000
    const sender = { tab: { id: 31 } as chrome.tabs.Tab }
    const issued = issuePanelPairingNonce(sender, t0)
    const nonce = (issued as { data: { nonce: string } }).data.nonce
    expect(redeemPanelPairingNonce(nonce, sender, t0 + PAIRING_NONCE_TTL_MS + 1)).toEqual({
      ok: true,
      data: { valid: false },
    })
  })

  it('memory stays bounded per tab — the oldest nonces of a spamming tab are dropped first', () => {
    const sender = { tab: { id: 32 } as chrome.tabs.Tab }
    const t0 = 2_000_000
    const all: string[] = []
    for (let i = 0; i < 40; i++) {
      all.push((issuePanelPairingNonce(sender, t0) as { data: { nonce: string } }).data.nonce)
    }
    expect(redeemPanelPairingNonce(all[0], sender, t0)).toMatchObject({ data: { valid: false } })
    expect(redeemPanelPairingNonce(all[39], sender, t0)).toMatchObject({ data: { valid: true } })
  })
})

describe('bridge + SW end to end', () => {
  function frameIn(tabId: number) {
    const parent = { name: 'owa-top' }
    let listener: ((e: MessageEvent) => void) | null = null
    const win: BridgeWindow = {
      parent,
      addEventListener: (_t, l) => {
        listener = l
      },
      removeEventListener: () => {},
    }
    // The production redeemer, routed through the real dispatcher with this
    // frame's sender (the browser fills in sender.tab for the iframe page).
    const bridge = createParentBridge(win, async (nonce) => redeem(panelIn(tabId), nonce))
    bridge.init()
    const deliver = (data: unknown, ports: unknown[]) =>
      listener?.({ data, source: parent, ports } as unknown as MessageEvent)
    return { bridge, deliver }
  }
  const port = () => ({ postMessage: vi.fn(), close: vi.fn() })
  const hello = (nonce?: string) => ({ type: 'mail-organizer/parent-hello', nonce })

  it('pairs with the content script’s hello; a page script’s hello (same window) never pairs', async () => {
    const { bridge, deliver } = frameIn(7)
    bridge.post({ type: 'mail-organizer/prewarm-sidebar', paths: ['Cases/Acme'] })

    // The page script races in first — it can't call chrome.runtime, so its
    // nonce is made up (or replayed from nowhere).
    const evil = port()
    deliver(hello('0123456789abcdef0123456789abcdef'), [evil])
    deliver(hello(), [evil])
    await vi.waitFor(() => expect(evil.close).toHaveBeenCalledTimes(2))
    expect(bridge.isConnected()).toBe(false)
    expect(evil.postMessage).not.toHaveBeenCalled()

    // The FAB content script's real hello.
    const real = port()
    deliver(hello(await issue(contentScriptIn(7))), [real])
    await vi.waitFor(() => expect(bridge.isConnected()).toBe(true))
    expect(real.postMessage).toHaveBeenNthCalledWith(1, { type: 'mail-organizer/bridge-ack' })
    expect(real.postMessage).toHaveBeenNthCalledWith(2, {
      type: 'mail-organizer/prewarm-sidebar',
      paths: ['Cases/Acme'],
    })
  })

  it('a nonce issued to the content script of another tab does not pair', async () => {
    const { bridge, deliver } = frameIn(9)
    const p = port()
    deliver(hello(await issue(contentScriptIn(7))), [p])
    await vi.waitFor(() => expect(p.close).toHaveBeenCalled())
    expect(bridge.isConnected()).toBe(false)
  })
})
