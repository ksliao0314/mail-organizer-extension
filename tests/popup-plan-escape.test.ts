// Esc / ? on the plan screen inside the OWA panel (2026-11-06 review U1 + U7).
//
// - The App-level Esc → close-panel listener registers at mount, so it ran
//   BEFORE the plan screen's own Esc (clear the cursor): one Esc cleared the
//   ring AND closed the whole panel. Now Esc clears the cursor first and only
//   an Esc with nothing to clear closes the panel.
// - A focused plan-row checkbox counted as a typing surface for the App-level
//   handlers: `?` did nothing and Esc never closed the panel — unlike the plan
//   screen's own shortcuts, which treat the checkbox as list chrome.

import type { Root } from 'react-dom/client'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import type { PlanItem } from '@/shared/types'
import { cleanup, click, keydown, mountApp, stubRuntime } from './helpers/ui-harness'

const posted = vi.hoisted(() => [] as Array<{ type?: string }>)
vi.mock('@/shared/parent-bridge', () => ({
  initParentBridge: () => () => {},
  postToParent: (m: { type?: string }) => {
    posted.push(m)
    return true
  },
}))

const realParent = Object.getOwnPropertyDescriptor(window, 'parent')
beforeAll(() => {
  // Embedded in the OWA floating-button iframe.
  Object.defineProperty(window, 'parent', { value: { name: 'owa-top' }, configurable: true })
})
afterAll(() => {
  if (realParent) Object.defineProperty(window, 'parent', realParent)
})

let root: Root | null = null
afterEach(() => {
  cleanup(root)
  root = null
  posted.length = 0
})

const closes = () => posted.filter((m) => m.type === 'mail-organizer/close-panel').length

const items: PlanItem[] = ['first mail', 'second mail', 'third mail'].map((s, i) => ({
  emailId: `e${i}`,
  emailSubject: s,
  emailFrom: 'a@acme.example',
  action: 'skip',
  confidence: 0.9,
  reason: '',
  source: 'ai',
}))

async function renderPlan() {
  stubRuntime((m) => {
    switch (m.type) {
      case 'getStatus':
        return {
          ok: true,
          data: {
            owaConnected: true,
            tokenValid: true,
            apiKeyConfigured: true,
            model: 'claude-sonnet-4-6',
            folderCacheCount: 0,
            excludePrefixes: [],
            batchSize: 50,
            internalDomains: [],
            primaryRootPath: '',
          },
        }
      case 'getOnboardingState':
        return { ok: true, data: { needed: false } }
      case 'getFolderActivity':
        return { ok: true, data: { activity: [] } }
      case 'getExecuteState':
        return { ok: true, data: { state: null } }
      case 'getAiClassifyProgress':
        return { ok: true, data: null }
      case 'getPopupState':
        return {
          ok: true,
          data: {
            phase: 'plan',
            updatedAt: Date.now(),
            plan: { items, folderTree: [], usage: null, summary: { ruleHits: 0, aiHandled: 3 }, banner: null },
          },
        }
      case 'savePopupState':
        return { ok: true }
      default:
        return undefined
    }
  })
  const { default: App } = await import('@/popup/App')
  const r = await mountApp(App)
  root = r.root
  await vi.waitFor(() => expect(r.host.querySelector('[data-plan-row-index="2"]')).not.toBeNull())
  return r.host
}

const row = (host: HTMLElement, i: number) => host.querySelector<HTMLElement>(`[data-plan-row-index="${i}"]`)!
const hasCursor = (r: HTMLElement) => r.className.includes('ring-2')

describe('Esc in the OWA panel, plan screen', () => {
  it('clears the keyboard cursor first; only the next Esc closes the panel', async () => {
    const host = await renderPlan()
    keydown(document.body, 'j')
    await vi.waitFor(() => expect(hasCursor(row(host, 0))).toBe(true))

    expect(keydown(document.body, 'Escape').defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(hasCursor(row(host, 0))).toBe(false))
    expect(closes()).toBe(0)

    keydown(document.body, 'Escape')
    expect(closes()).toBe(1)
  })

  it('a focused row checkbox is list chrome: Esc clears the cursor, then closes the panel', async () => {
    const host = await renderPlan()
    const cb = row(host, 1).querySelector<HTMLInputElement>('input[type="checkbox"]')!
    // A mouse click leaves the checkbox focused (and selects the row).
    cb.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    cb.focus()
    click(cb)
    keydown(cb, 'j')
    await vi.waitFor(() => expect(hasCursor(row(host, 0))).toBe(true))

    keydown(cb, 'Escape')
    await vi.waitFor(() => expect(hasCursor(row(host, 0))).toBe(false))
    expect(closes()).toBe(0)
    keydown(cb, 'Escape')
    expect(closes()).toBe(1)
  })
})

describe('? with a row checkbox focused', () => {
  it('opens the shortcut help like anywhere else on the plan screen', async () => {
    const host = await renderPlan()
    const cb = row(host, 0).querySelector<HTMLInputElement>('input[type="checkbox"]')!
    cb.focus()
    keydown(cb, '?')
    await vi.waitFor(() => expect(document.querySelector('[aria-label="鍵盤捷徑"]')).not.toBeNull())
  })
})
