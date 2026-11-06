// In-OWA panel bridge pairing nonces (2026-11-06).
//
// The popup / options pages are web-accessible to the Outlook origins so the
// FAB content script (content/owa-fab.ts) can show them in an iframe. That
// also means an Outlook PAGE script can embed them itself — and since page
// scripts share `window` with the content script, a `parent-hello` they post
// passes the frame's `event.source === window.parent` check: they could pair
// with a frame they embedded (and receive the 近日活動 folder paths the popup
// sends on mount), or race the content script's own hello into the real
// panel. Pairing is therefore anchored in chrome.runtime, which page scripts
// cannot reach:
//
//   1. owa-fab asks for a one-time nonce ('issuePanelPairingNonce', a
//      content-script message) and puts it in its `parent-hello`;
//   2. the frame — an extension page, whose runtime messages carry the tab it
//      is embedded in — redeems it ('redeemPanelPairingNonce') and accepts
//      the hello's port only if the nonce was issued to a content script in
//      the SAME tab (shared/parent-bridge.ts). Redeeming consumes it.
//
// Nonces live in SW memory only: short-lived, single-use, never persisted or
// logged. An SW restart between issue and redeem just fails that one offer —
// the content script re-offers with a fresh nonce on its retry schedule.

import type { PopupResponse } from '@/shared/messages'

/** How long an issued nonce stays redeemable. Pairing normally takes < 1 s. */
export const PAIRING_NONCE_TTL_MS = 60_000
/** Each FAB iframe load issues ≤ 6 (first offer + 5 retries); two frames per tab. */
const MAX_NONCES_PER_TAB = 16
const MAX_NONCES_TOTAL = 512

type Entry = { tabId: number; expiresAt: number }
const nonces = new Map<string, Entry>()

type Sender = Pick<chrome.runtime.MessageSender, 'tab'> | undefined

function senderTabId(sender: Sender): number | null {
  const id = sender?.tab?.id
  return typeof id === 'number' && Number.isInteger(id) && id >= 0 ? id : null
}

function prune(now: number): void {
  for (const [nonce, e] of nonces) if (e.expiresAt <= now) nonces.delete(nonce)
}

function randomNonce(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(16))
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('')
}

/** Content script → SW: a fresh nonce bound to the sender's tab. */
export function issuePanelPairingNonce(sender: Sender, now = Date.now()): PopupResponse {
  const tabId = senderTabId(sender)
  if (tabId === null) {
    return { ok: false, code: 'NO_TAB', message: 'Pairing nonces are only issued to a tab' }
  }
  prune(now)
  // Bound memory — Map iteration is insertion order, so oldest go first.
  const ofTab = [...nonces].filter(([, e]) => e.tabId === tabId)
  for (const [nonce] of ofTab.slice(0, Math.max(0, ofTab.length - MAX_NONCES_PER_TAB + 1))) {
    nonces.delete(nonce)
  }
  while (nonces.size >= MAX_NONCES_TOTAL) {
    const oldest = nonces.keys().next().value
    if (oldest === undefined) break
    nonces.delete(oldest)
  }
  const nonce = randomNonce()
  nonces.set(nonce, { tabId, expiresAt: now + PAIRING_NONCE_TTL_MS })
  return { ok: true, data: { nonce } }
}

/**
 * Panel frame → SW: is this the nonce a content script in MY tab was issued?
 * A match is consumed; anything else answers `valid: false` and consumes
 * nothing (a wrong-tab probe must not burn the real frame's nonce).
 */
export function redeemPanelPairingNonce(
  nonce: unknown,
  sender: Sender,
  now = Date.now(),
): PopupResponse {
  prune(now)
  const tabId = senderTabId(sender)
  const entry = typeof nonce === 'string' ? nonces.get(nonce) : undefined
  const valid = !!entry && tabId !== null && entry.tabId === tabId
  if (valid) nonces.delete(nonce as string)
  return { ok: true, data: { valid } }
}
