// Render the real popup / options <App/> under jsdom against a stub service
// worker (2026-11-05). Every chrome.runtime.sendMessage goes to `handler`;
// anything it doesn't answer gets { ok:false, code:'TEST_STUB' } — the same
// shape a real SW error has, so the UI's error paths run.

import { act, createElement, type ComponentType } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { vi } from 'vitest'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true
// jsdom has no layout APIs.
if (!Element.prototype.scrollIntoView) Element.prototype.scrollIntoView = () => {}
if (!(globalThis as { ResizeObserver?: unknown }).ResizeObserver) {
  ;(globalThis as { ResizeObserver?: unknown }).ResizeObserver = class {
    observe() {}
    unobserve() {}
    disconnect() {}
  }
}

export type Msg = { type: string } & Record<string, unknown>
export type Reply = { ok: boolean; code?: string; message?: string; data?: unknown }

export function stubRuntime(handler: (msg: Msg) => Reply | undefined | Promise<Reply | undefined>) {
  const sent: Msg[] = []
  const runtime = {
    id: 'test-extension',
    getURL: (p: string) => `chrome-extension://test-extension/${p}`,
    getManifest: () => ({ version: '0.0.0-test' }),
    openOptionsPage: vi.fn(),
    sendMessage: vi.fn(async (msg: Msg) => {
      sent.push(msg)
      const r = await handler(msg)
      return r ?? { ok: false, code: 'TEST_STUB', message: 'stubbed' }
    }),
  }
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.runtime = runtime
  return {
    sent,
    runtime,
    count: (type: string) => sent.filter((m) => m.type === type).length,
    last: (type: string) => [...sent].reverse().find((m) => m.type === type),
  }
}

export async function mountApp(App: ComponentType): Promise<{ host: HTMLElement; root: Root }> {
  const host = document.createElement('div')
  document.body.appendChild(host)
  let root!: Root
  await act(async () => {
    root = createRoot(host)
    root.render(createElement(App))
  })
  return { host, root }
}

export function cleanup(root: Root | null) {
  act(() => root?.unmount())
  document.body.innerHTML = ''
  delete (globalThis as unknown as { chrome: { runtime?: unknown } }).chrome.runtime
}

/** Latest chrome.storage.onChanged listener the page registered (setup.ts mock). */
export function latestStorageListener(): (changes: Record<string, unknown>, area: string) => void {
  const calls = (chrome.storage.onChanged.addListener as unknown as { mock: { calls: unknown[][] } })
    .mock.calls
  const fn = calls[calls.length - 1]?.[0]
  if (typeof fn !== 'function') throw new Error('no storage listener registered')
  return fn as (changes: Record<string, unknown>, area: string) => void
}

export function click(el: Element | null | undefined) {
  if (!el) throw new Error('click target missing')
  act(() => {
    ;(el as HTMLElement).click()
  })
}

export function keydown(target: EventTarget, key: string, init: KeyboardEventInit = {}) {
  const ev = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init })
  act(() => {
    target.dispatchEvent(ev)
  })
  return ev
}

/** Set a React-controlled input's value the way a user edit would. */
export function typeInto(el: HTMLInputElement | HTMLTextAreaElement, value: string) {
  const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype
  const setter = Object.getOwnPropertyDescriptor(proto, 'value')!.set!
  act(() => {
    setter.call(el, value)
    el.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

export function buttonByText(root: ParentNode, text: string): HTMLButtonElement | undefined {
  return Array.from(root.querySelectorAll('button')).find(
    (b) => (b.textContent ?? '').replace(/\s+/g, ' ').trim() === text,
  )
}

export function buttonContaining(root: ParentNode, text: string): HTMLButtonElement | undefined {
  return Array.from(root.querySelectorAll('button')).find((b) => (b.textContent ?? '').includes(text))
}
