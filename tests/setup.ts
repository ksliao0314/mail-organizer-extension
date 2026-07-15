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
    // Mirror Chrome's real accounting: each item costs its key's UTF-8
    // bytes + its JSON-serialised value's UTF-8 bytes. Crucially NOT
    // String.length (which under-counts CJK 3×) — tests rely on this to
    // catch byte-vs-char regressions in the sync chunker / usage meter.
    getBytesInUse: vi.fn(async (keys?: string | string[] | null) => {
      const enc = new TextEncoder()
      const wanted =
        keys == null
          ? Object.keys(stores[area])
          : typeof keys === 'string'
            ? [keys]
            : keys
      let bytes = 0
      for (const k of wanted) {
        if (!(k in stores[area])) continue
        bytes += enc.encode(k).length + enc.encode(JSON.stringify(stores[area][k])).length
      }
      return bytes
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
