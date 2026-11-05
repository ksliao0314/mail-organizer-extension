// Load the REAL service-worker module under the test chrome mock and return a
// `send()` that drives its runtime.onMessage dispatcher with exactly the
// message shapes the UI sends. Contract tests built on this catch UI↔SW drift
// (the export "undefined" file bug 75e436e, the import `json`/`payload` bug)
// that unit tests of either side alone can't see.
//
// tests/setup.ts only mocks chrome.storage; the SW also touches runtime /
// alarms / tabs / action / windows at module load, so those get inert stubs
// here. Import once per test file (vitest isolates modules per file).

import { beforeEach, vi } from 'vitest'

type Listener = (...args: unknown[]) => unknown
type Response = { ok: boolean; code?: string; message?: string; data?: unknown }
export type Sender = chrome.runtime.MessageSender

function event() {
  const listeners: Listener[] = []
  return {
    listeners,
    addListener: (l: Listener) => listeners.push(l),
    removeListener: () => {},
    hasListener: () => false,
  }
}

/**
 * The sender `send()` uses: an extension page (2026-11-05). The dispatcher
 * only serves extension pages (popup / options, incl. the in-OWA iframe
 * panel); content-script senders are limited to an allowlist — use
 * `sendFrom(CONTENT_SCRIPT_SENDER, msg)` to exercise that path.
 */
export const EXTENSION_PAGE_SENDER: Sender = {
  id: 'test-extension',
  url: 'chrome-extension://test-extension/src/options/index.html',
}

export const CONTENT_SCRIPT_SENDER: Sender = {
  id: 'test-extension',
  url: 'https://outlook.office.com/mail/inbox',
  origin: 'https://outlook.office.com',
  tab: { id: 7 } as chrome.tabs.Tab,
  frameId: 0,
}

export type SwHarness = {
  send: (msg: Record<string, unknown>) => Promise<Response>
  sendFrom: (sender: Sender, msg: unknown) => Promise<Response>
  /**
   * Deliver a chrome.storage.onChanged event to every listener the SW
   * registered. The setup.ts mock records listeners but never fires them;
   * SW caches keyed on onChanged (getStatus counts) need this after a test
   * writes rules / folderCache directly through storage helpers.
   */
  emitStorageChange: (
    changes: Record<string, chrome.storage.StorageChange>,
    area?: chrome.storage.AreaName,
  ) => Promise<void>
  tabsSendMessage: ReturnType<typeof vi.fn>
  tabsQuery: ReturnType<typeof vi.fn>
}

let storageListeners: Listener[] = []

async function emit(
  changes: Record<string, chrome.storage.StorageChange>,
  area: chrome.storage.AreaName = 'local',
): Promise<void> {
  for (const l of storageListeners) l(changes, area)
  // Let async listener work (e.g. the owa-fab broadcast) settle.
  await new Promise((r) => setTimeout(r, 0))
}

// setup.ts wipes the mock stores before every test WITHOUT firing
// storage.onChanged; real Chrome would fire it for every removed key. Mirror
// that for the keys SW in-memory caches watch, so a cached getStatus count
// from one test can't leak into the next (2026-11-05).
beforeEach(async () => {
  if (storageListeners.length === 0) return
  for (const l of storageListeners) l({ rules: {}, folderCache: {} }, 'local')
})

export async function loadServiceWorker(): Promise<SwHarness> {
  const c = (globalThis as unknown as { chrome: Record<string, unknown> }).chrome
  const onMessage = event()
  const tabsSendMessage = vi.fn(async () => undefined)
  const tabsQuery = vi.fn(async () => [] as chrome.tabs.Tab[])

  // Record storage.onChanged listeners ourselves (still forwarding to the
  // setup.ts vi.fn so `.mock.calls` keeps working for existing tests).
  const storage = c.storage as { onChanged: { addListener: (l: Listener) => void } }
  const forward = storage.onChanged.addListener
  storageListeners = []
  storage.onChanged.addListener = vi.fn((l: Listener) => {
    storageListeners.push(l)
    forward(l)
  })

  c.runtime = {
    id: 'test-extension',
    onMessage,
    onInstalled: event(),
    onStartup: event(),
    getURL: (p: string) => `chrome-extension://test-extension/${p}`,
    getManifest: () => ({ version: '0.0.0-test' }),
    sendMessage: vi.fn(async () => undefined),
    openOptionsPage: vi.fn(async () => {}),
  }
  c.alarms = {
    get: vi.fn(async () => undefined),
    create: vi.fn(async () => {}),
    clear: vi.fn(async () => true),
    onAlarm: event(),
  }
  c.tabs = {
    query: tabsQuery,
    sendMessage: tabsSendMessage,
    update: vi.fn(async () => ({})),
    create: vi.fn(async () => ({})),
    onCreated: event(),
    onUpdated: event(),
    onRemoved: event(),
    onActivated: event(),
  }
  c.action = { setPopup: vi.fn(async () => {}), onClicked: event() }
  c.windows = { update: vi.fn(async () => ({})) }
  c.downloads = { download: vi.fn(async () => 1) }

  await import('@/background/service-worker')
  const dispatch = onMessage.listeners[0]
  if (!dispatch) throw new Error('service worker registered no runtime.onMessage listener')
  const sendFrom = (sender: Sender, msg: unknown) =>
    new Promise<Response>((resolve) => {
      dispatch(msg, sender, resolve)
    })
  const send = (msg: Record<string, unknown>) => sendFrom(EXTENSION_PAGE_SENDER, msg)
  return { send, sendFrom, emitStorageChange: emit, tabsSendMessage, tabsQuery }
}
