// Load the REAL service-worker module under the test chrome mock and return a
// `send()` that drives its runtime.onMessage dispatcher with exactly the
// message shapes the UI sends. Contract tests built on this catch UI↔SW drift
// (the export "undefined" file bug 75e436e, the import `json`/`payload` bug)
// that unit tests of either side alone can't see.
//
// tests/setup.ts only mocks chrome.storage; the SW also touches runtime /
// alarms / tabs / action / windows at module load, so those get inert stubs
// here. Import once per test file (vitest isolates modules per file).

import { vi } from 'vitest'

type Listener = (...args: unknown[]) => unknown
type Response = { ok: boolean; code?: string; message?: string; data?: unknown }

function event() {
  const listeners: Listener[] = []
  return {
    listeners,
    addListener: (l: Listener) => listeners.push(l),
    removeListener: () => {},
    hasListener: () => false,
  }
}

export type SwHarness = {
  send: (msg: Record<string, unknown>) => Promise<Response>
  tabsSendMessage: ReturnType<typeof vi.fn>
  tabsQuery: ReturnType<typeof vi.fn>
}

export async function loadServiceWorker(): Promise<SwHarness> {
  const c = (globalThis as unknown as { chrome: Record<string, unknown> }).chrome
  const onMessage = event()
  const tabsSendMessage = vi.fn(async () => undefined)
  const tabsQuery = vi.fn(async () => [] as chrome.tabs.Tab[])
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
  const send = (msg: Record<string, unknown>) =>
    new Promise<Response>((resolve) => {
      dispatch(msg, { id: 'test-extension' }, resolve)
    })
  return { send, tabsSendMessage, tabsQuery }
}
