// Web storage + root error boundary (2026-11-05 audit).
//
// In the OWA-embedded iframe with third-party storage blocked, merely
// READING window.localStorage / window.sessionStorage throws SecurityError —
// including `typeof sessionStorage`, which evaluates the getter. The popup
// read localStorage unguarded in a useState initializer, and the options
// helpers ran `typeof` outside their try. And with no error boundary, one
// render-time throw left a blank panel.

import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { readWebStorage, writeWebStorage } from '@/popup/web-storage'
import { ErrorBoundary } from '@/components/ui/error-boundary'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const realLocal = Object.getOwnPropertyDescriptor(window, 'localStorage')
const realSession = Object.getOwnPropertyDescriptor(window, 'sessionStorage')

function blockStorage() {
  const thrower = {
    configurable: true,
    get() {
      throw new DOMException('The operation is insecure.', 'SecurityError')
    },
  }
  Object.defineProperty(window, 'localStorage', thrower)
  Object.defineProperty(window, 'sessionStorage', thrower)
}

afterEach(() => {
  if (realLocal) Object.defineProperty(window, 'localStorage', realLocal)
  if (realSession) Object.defineProperty(window, 'sessionStorage', realSession)
  document.body.innerHTML = ''
})

describe('guarded web storage', () => {
  it('reads null and swallows writes when the storage getter itself throws', () => {
    blockStorage()
    expect(() => window.localStorage).toThrow() // the failure mode being guarded
    expect(readWebStorage('local', 'k')).toBeNull()
    expect(readWebStorage('session', 'k')).toBeNull()
    expect(() => writeWebStorage('local', 'k', 'v')).not.toThrow()
    expect(() => writeWebStorage('session', 'k', 'v')).not.toThrow()
  })

  it('round-trips normally when storage is available', () => {
    writeWebStorage('session', 'mo-test', '1')
    expect(readWebStorage('session', 'mo-test')).toBe('1')
  })
})

describe('root ErrorBoundary', () => {
  it('replaces a crashed tree with a friendly fallback and a reload button', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {})
    function Boom(): never {
      throw new Error('render exploded')
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(createElement(ErrorBoundary, null, createElement(Boom)))
    })
    expect(host.textContent).toContain('畫面發生錯誤')
    expect(host.textContent).toContain('render exploded')
    expect(host.querySelector('button')?.textContent).toContain('重新載入')
    act(() => root.unmount())
    spy.mockRestore()
  })
})
