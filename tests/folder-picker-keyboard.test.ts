// FolderPicker keyboard contract (2026-11-05 audit).
//
//   - ↑/↓ move an aria-activedescendant cursor, Enter picks it (there used
//     to be no keyboard selection at all).
//   - Esc on an OPEN dropdown closes only the dropdown and does not reach
//     document-level Esc handlers (it used to close the whole rule drawer,
//     discarding the edit). Esc on a CLOSED picker passes through.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { FolderPicker } from '@/popup/components/FolderPicker'
import type { MailFolderNode } from '@/shared/types'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const tree: MailFolderNode[] = [
  {
    id: 'A',
    displayName: 'Alpha',
    path: 'Alpha',
    children: [
      { id: 'A1', displayName: 'One', path: 'Alpha/One', children: [] },
      { id: 'A2', displayName: 'Two', path: 'Alpha/Two', children: [] },
    ],
  },
]

let root: Root | null = null
afterEach(() => {
  act(() => root?.unmount())
  root = null
  document.body.innerHTML = ''
})

async function mount(onSelect = vi.fn()) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  await act(async () => {
    root = createRoot(host)
    root.render(createElement(FolderPicker, { tree, onSelect }))
  })
  const input = host.querySelector('input')!
  return { host, input, onSelect }
}

function key(el: Element, k: string) {
  const ev = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })
  act(() => {
    el.dispatchEvent(ev)
  })
  return ev
}

describe('FolderPicker keyboard', () => {
  it('arrows move the active option and Enter selects it', async () => {
    const { input, onSelect } = await mount()
    act(() => input.focus())
    expect(input.getAttribute('aria-expanded')).toBe('true')
    const first = input.getAttribute('aria-activedescendant')
    expect(first).toBeTruthy()
    expect(document.getElementById(first!)?.textContent).toContain('Alpha')

    key(input, 'ArrowDown')
    const second = input.getAttribute('aria-activedescendant')!
    expect(second).not.toBe(first)
    expect(document.getElementById(second)?.textContent).toContain('Alpha/One')

    key(input, 'Enter')
    expect(onSelect).toHaveBeenCalledWith(expect.objectContaining({ id: 'A1', path: 'Alpha/One' }))
    expect(input.getAttribute('aria-expanded')).toBe('false')
  })

  it('Esc on an open dropdown closes only the dropdown; on a closed one it passes through', async () => {
    const { input } = await mount()
    const docEsc = vi.fn()
    document.addEventListener('keydown', docEsc)
    try {
      act(() => input.focus())
      expect(input.getAttribute('aria-expanded')).toBe('true')
      const ev = key(input, 'Escape')
      expect(input.getAttribute('aria-expanded')).toBe('false')
      expect(ev.defaultPrevented).toBe(true)
      expect(docEsc).not.toHaveBeenCalled()

      key(input, 'Escape') // closed now → next layer's business
      expect(docEsc).toHaveBeenCalledTimes(1)
    } finally {
      document.removeEventListener('keydown', docEsc)
    }
  })
})
