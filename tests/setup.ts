// Mock chrome.storage.local / session / sync for tests. The real chrome
// global is only available in extension contexts; tests run under
// jsdom + Node.

import { vi } from 'vitest'

const stores: Record<'local' | 'session' | 'sync', Record<string, unknown>> = {
  local: {},
  session: {},
  sync: {},
}

function makeArea(area: 'local' | 'session' | 'sync') {
  return {
    get: vi.fn(async (keys?: string | string[] | null) => {
      if (keys == null) return { ...stores[area] }
      if (typeof keys === 'string') {
        return keys in stores[area] ? { [keys]: stores[area][keys] } : {}
      }
      const out: Record<string, unknown> = {}
      for (const k of keys) {
        if (k in stores[area]) out[k] = stores[area][k]
      }
      return out
    }),
    set: vi.fn(async (items: Record<string, unknown>) => {
      Object.assign(stores[area], items)
    }),
    remove: vi.fn(async (keys: string | string[]) => {
      const arr = Array.isArray(keys) ? keys : [keys]
      for (const k of arr) delete stores[area][k]
    }),
    clear: vi.fn(async () => {
      stores[area] = {}
    }),
    QUOTA_BYTES: 102_400,
    QUOTA_BYTES_PER_ITEM: 8_192,
  }
}

;(globalThis as unknown as { chrome: unknown }).chrome = {
  storage: {
    local: makeArea('local'),
    session: makeArea('session'),
    sync: makeArea('sync'),
    onChanged: {
      addListener: vi.fn(),
      removeListener: vi.fn(),
    },
  },
}

// Reset between tests
import { beforeEach } from 'vitest'
beforeEach(() => {
  stores.local = {}
  stores.session = {}
  stores.sync = {}
})
