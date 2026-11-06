// Cross-frame bridge for the popup / options pages when they run inside the
// OWA floating-button iframes (content/owa-fab.ts is the other end).
//
// Handshake over a private MessageChannel (2026-11-05). The previous
// protocol learned the parent's origin from ANY `parent-hello` — no source
// check — so a page script that could reach this frame (owa-fab's shadow
// root used to be OPEN; and our own `iframe-ready` ping, broadcast with '*',
// handed every listener on the OWA page this frame's WindowProxy as
// `event.source`) could re-pin the target origin and silently black-hole
// close-panel / open-options / navigate-folder. And because messages were
// posted to the OWA origin, every JS world on the page (OWA itself, other
// extensions' content scripts) received prewarm-sidebar / navigate-folder
// payloads — i.e. the user's folder names and paths.
//
// Now:
//   - the FAB content script creates a MessageChannel per offer and
//     transfers one port inside `parent-hello`, posted straight to
//     iframe.contentWindow with targetOrigin = the extension origin;
//   - every hello carries a one-time nonce the content script got from the
//     service worker (2026-11-06). `event.source === window.parent` can't
//     tell the content script from an Outlook PAGE script — they share that
//     window — and these pages are web-accessible to Outlook, so a page
//     script could embed this page itself, or race its own hello into the
//     real panel before the content script's (offered on iframe `load`,
//     while this listener runs from module evaluation). So a hello's port is
//     accepted only after the SW confirms its nonce was issued to a content
//     script in this frame's own tab (redeeming consumes it; see
//     background/handlers/panel-pairing.ts). Page scripts can't reach
//     chrome.runtime, so they can't mint one;
//   - we accept only the FIRST verified hello from `window.parent`; any
//     other hello is refused (its ports closed), so a late or foreign hello
//     can't steal the channel;
//   - everything afterwards travels over that port — nothing is broadcast
//     to the host page, and NOTHING is sent before pairing: messages posted
//     earlier wait in a bounded queue that only a verified port receives;
//   - we ack over the port so the content script can stop re-offering.
//
// Pairing protects the data channel only. A page script can still embed
// these (fully privileged) pages and overlay them — clickjacking — so when
// framed, the UI itself renders only after pairing (2026-11-06,
// components/ui/bridge-gate.tsx); onConnected() is its signal.
//
// The listener is installed at module-evaluation time from main.tsx
// (init is idempotent) so a hello arriving before React mounts isn't lost.

import { send } from './send-message'

const HELLO = 'mail-organizer/parent-hello'
const ACK = 'mail-organizer/bridge-ack'
const MAX_QUEUED = 20
/** Real nonces are 32 hex chars — anything far longer is junk, not worth a SW round-trip. */
const MAX_NONCE_LEN = 128

/** The slice of `Window` the bridge touches — injectable for unit tests. */
export type BridgeWindow = {
  readonly parent: unknown
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void
  removeEventListener(type: 'message', listener: (e: MessageEvent) => void): void
}

/**
 * Ask the service worker whether `nonce` was issued to the FAB content
 * script of this frame's tab (and consume it). Injectable for unit tests.
 */
export type RedeemNonce = (nonce: string) => Promise<boolean>

export type PostOptions = {
  /**
   * Queue the message until the handshake completes (default true). Pass
   * false when the caller has its own fallback for "no live bridge" (e.g.
   * open options in a tab instead) — otherwise the queued message would
   * also fire later and the action would happen twice.
   */
  queue?: boolean
}

export type ParentBridge = {
  /** Start listening for the handshake. Idempotent; returns a no-op cleanup. */
  init(): () => void
  /** Send to the FAB content script. Returns true only if delivered now. */
  post(message: object, opts?: PostOptions): boolean
  /** A verified FAB content script is on the other end (pairing completed). */
  isConnected(): boolean
  /** This page is framed (the in-OWA panel / settings modal) rather than top-level. */
  isEmbedded(): boolean
  /**
   * Called once when pairing completes (2026-11-06) — BridgeGate shows the
   * framed UI only from then on. Returns an unsubscribe.
   */
  onConnected(listener: () => void): () => void
}

function closePorts(ports: readonly MessagePort[]): void {
  for (const p of ports) {
    try {
      p.close()
    } catch {
      /* ignore */
    }
  }
}

export function createParentBridge(win: BridgeWindow, redeemNonce: RedeemNonce): ParentBridge {
  let port: MessagePort | null = null
  let listening = false
  const queue: object[] = []
  const connectedListeners = new Set<() => void>()
  const embedded = () => win.parent !== win

  function adopt(offered: MessagePort, verified: boolean) {
    // Unverified, or another hello won the race while this one was being
    // checked — release the port so it can't linger.
    if (!verified || port) {
      closePorts([offered])
      return
    }
    port = offered
    try {
      port.postMessage({ type: ACK })
      while (queue.length > 0) port.postMessage(queue.shift()!)
    } catch (err) {
      console.warn('[mail-organizer] parent bridge: port send failed', err)
    }
    for (const listener of [...connectedListeners]) {
      try {
        listener()
      } catch (err) {
        console.warn('[mail-organizer] parent bridge: listener failed', err)
      }
    }
  }

  function onMessage(e: MessageEvent) {
    const data = e.data as { type?: unknown; nonce?: unknown } | null
    if (!data || typeof data !== 'object' || data.type !== HELLO) return
    const ports = e.ports ? Array.from(e.ports) : []
    const [offered, ...extra] = ports
    closePorts(extra)
    const nonce =
      typeof data.nonce === 'string' && data.nonce.length <= MAX_NONCE_LEN ? data.nonce : ''
    if (port || e.source !== win.parent || !offered || !nonce) {
      // Already paired, not from our parent window, no channel or no nonce
      // offered (the pre-2026-11-06 protocol) — refuse.
      if (offered) closePorts([offered])
      return
    }
    void Promise.resolve()
      .then(() => redeemNonce(nonce))
      .then(
        (verified) => adopt(offered, verified === true),
        () => adopt(offered, false),
      )
  }

  function init(): () => void {
    if (!embedded()) return () => {}
    if (!listening) {
      win.addEventListener('message', onMessage)
      listening = true
    }
    // Deliberately persistent: the handshake may arrive at any point in the
    // page's life (a StrictMode effect re-run must not drop the listener).
    return () => {}
  }

  function post(message: object, opts?: PostOptions): boolean {
    // Top-level (toolbar popup / standalone options tab) — no parent.
    if (!embedded()) return false
    init()
    if (!port) {
      if (opts?.queue === false) return false
      if (queue.length >= MAX_QUEUED) queue.shift()
      queue.push(message)
      return false
    }
    try {
      port.postMessage(message)
      return true
    } catch (err) {
      console.warn('[mail-organizer] parent bridge: post failed', err)
      return false
    }
  }

  function onConnected(listener: () => void): () => void {
    connectedListeners.add(listener)
    return () => {
      connectedListeners.delete(listener)
    }
  }

  return { init, post, isConnected: () => port !== null, isEmbedded: embedded, onConnected }
}

/** The production redeemer: the SW checks the nonce against this frame's tab. */
async function redeemWithServiceWorker(nonce: string): Promise<boolean> {
  try {
    const r = await send<{ valid?: boolean }>({ type: 'redeemPanelPairingNonce', nonce })
    return r?.ok === true && r.data?.valid === true
  } catch {
    // SW unreachable / extension reloaded — can't verify, so don't pair.
    return false
  }
}

const defaultBridge: ParentBridge | null =
  typeof window !== 'undefined' ? createParentBridge(window, redeemWithServiceWorker) : null

export function initParentBridge(): () => void {
  return defaultBridge ? defaultBridge.init() : () => {}
}

/** This page's bridge (null outside a browser window) — for BridgeGate. */
export function getParentBridge(): ParentBridge | null {
  return defaultBridge
}

export function postToParent(message: object, opts?: PostOptions): boolean {
  return defaultBridge ? defaultBridge.post(message, opts) : false
}
