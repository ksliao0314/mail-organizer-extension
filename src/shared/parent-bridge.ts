// Cross-frame postMessage helper for popup / options iframes.
//
// Why: outbound window.parent.postMessage(..., '*') broadcasts the payload to
// whatever origin happens to host the iframe. Any extension content script
// attached to the OWA tab can listen and read it. Even if today's payload is
// just type+folderPath, the '*' is a foot-gun for future devs adding a more
// sensitive field. This module learns the parent origin via a one-time
// handshake (iframe pings parent → parent replies with hello → iframe
// stores e.origin) and uses it for all outbound sends.
//
// The initial iframe-ready ping is itself sent with '*', but it carries no
// data beyond the type identifier — safe to broadcast.

let parentOrigin: string | null = null

export function initParentBridge(): () => void {
  if (window.parent === window) {
    // Top-level (toolbar popup / standalone options tab) — no parent to talk to.
    return () => {}
  }
  function onMessage(e: MessageEvent) {
    const data = e.data as { type?: string } | undefined
    if (data?.type === 'mail-organizer/parent-hello') {
      parentOrigin = e.origin
    }
  }
  window.addEventListener('message', onMessage)
  try {
    window.parent.postMessage({ type: 'mail-organizer/iframe-ready' }, '*')
  } catch (e) {
    console.warn('[mail-organizer] iframe-ready ping failed', e)
  }
  return () => window.removeEventListener('message', onMessage)
}

export function postToParent(message: object): boolean {
  if (window.parent === window) return false
  if (!parentOrigin) {
    console.warn('[mail-organizer] postToParent before parent-hello — dropping', message)
    return false
  }
  window.parent.postMessage(message, parentOrigin)
  return true
}
