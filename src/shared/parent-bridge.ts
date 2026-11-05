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
//   - the FAB content script creates a MessageChannel per iframe load and
//     transfers one port inside `parent-hello`, posted straight to
//     iframe.contentWindow with targetOrigin = the extension origin;
//   - we accept only the FIRST hello whose `event.source === window.parent`
//     and that carries a port; any other hello is refused (its ports
//     closed), so a late or foreign hello can't steal the channel;
//   - everything afterwards travels over that port — nothing is broadcast
//     to the host page, and there is no `iframe-ready` ping any more;
//   - we ack over the port so the content script can stop re-offering.
//
// The listener is installed at module-evaluation time from main.tsx
// (init is idempotent) so a hello arriving before React mounts isn't lost;
// messages posted before the handshake completes are queued (bounded).

const HELLO = 'mail-organizer/parent-hello'
const ACK = 'mail-organizer/bridge-ack'
const MAX_QUEUED = 20

/** The slice of `Window` the bridge touches — injectable for unit tests. */
export type BridgeWindow = {
  readonly parent: unknown
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void
  removeEventListener(type: 'message', listener: (e: MessageEvent) => void): void
}

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
  isConnected(): boolean
}

export function createParentBridge(win: BridgeWindow): ParentBridge {
  let port: MessagePort | null = null
  let listening = false
  const queue: object[] = []
  const embedded = () => win.parent !== win

  function onMessage(e: MessageEvent) {
    const data = e.data as { type?: unknown } | null
    if (!data || typeof data !== 'object' || data.type !== HELLO) return
    const ports = e.ports ? Array.from(e.ports) : []
    const offered = ports[0]
    if (port || e.source !== win.parent || !offered) {
      // Already paired, not from our parent window, or no channel offered —
      // refuse, and release whatever was transferred so it can't linger.
      for (const p of ports) {
        try {
          p.close()
        } catch {
          /* ignore */
        }
      }
      return
    }
    port = offered
    try {
      port.postMessage({ type: ACK })
      while (queue.length > 0) port.postMessage(queue.shift()!)
    } catch (err) {
      console.warn('[mail-organizer] parent bridge: port send failed', err)
    }
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

  return { init, post, isConnected: () => port !== null }
}

const defaultBridge: ParentBridge | null =
  typeof window !== 'undefined' ? createParentBridge(window) : null

export function initParentBridge(): () => void {
  return defaultBridge ? defaultBridge.init() : () => {}
}

export function postToParent(message: object, opts?: PostOptions): boolean {
  return defaultBridge ? defaultBridge.post(message, opts) : false
}
