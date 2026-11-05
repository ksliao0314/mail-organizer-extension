// Options page loads the folder tree on mount (2026-11-05 audit).
//
// The only loadTree() call used to be inside the rule drawer's `editing`
// effect, while the drawer's 編輯 button is `disabled={!tree}` — a deadlock
// that left 新增規則 / 編輯 / the root-folder picker / the 近日活動 filter
// picker unusable. This renders the real Options <App/> under jsdom with a
// stub service worker and asserts the tree is requested without any click.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true
// jsdom has no layout APIs; the Options mount effect scrolls to #rules.
if (!Element.prototype.scrollIntoView) Element.prototype.scrollIntoView = () => {}

describe('Options page', () => {
  let root: Root | null = null
  afterEach(() => {
    act(() => root?.unmount())
    root = null
    document.body.innerHTML = ''
    delete (globalThis as unknown as { chrome: { runtime?: unknown } }).chrome.runtime
  })

  it('requests the folder tree on mount, before any user action', async () => {
    const requested: string[] = []
    ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.runtime = {
      id: 'test-extension',
      getURL: (p: string) => `chrome-extension://test-extension/${p}`,
      getManifest: () => ({ version: '0.0.0-test' }),
      openOptionsPage: vi.fn(),
      sendMessage: vi.fn(async (msg: { type: string }) => {
        requested.push(msg.type)
        if (msg.type === 'getFolderTree') {
          return { ok: true, data: { tree: [], count: 0 } }
        }
        return { ok: false, code: 'TEST_STUB', message: 'stubbed' }
      }),
    }
    const { default: App } = await import('@/options/App')
    const host = document.createElement('div')
    document.body.appendChild(host)
    await act(async () => {
      root = createRoot(host)
      root.render(createElement(App))
    })
    await vi.waitFor(() => expect(requested).toContain('getFolderTree'))
  })
})
