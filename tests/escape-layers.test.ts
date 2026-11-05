// Escape is innermost-first (2026-11-05 audit). In the OWA floating-button
// iframes the pages post close-panel / close-options to the parent on Esc;
// that listener registers at mount, BEFORE any overlay's own handler, so an
// open layer must make it stand down:
//   - popup: Esc with the ? shortcuts dialog open closed the whole OWA panel
//   - options: Esc with the 篩選 popover or the enable-sync modal open closed
//     the whole options modal
//   - rule drawer: Esc meant for the FolderPicker dropdown closed the drawer
//     and discarded the edit

import type { Root } from 'react-dom/client'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, Rule } from '@/shared/types'
import {
  buttonByText,
  buttonContaining,
  cleanup,
  click,
  keydown,
  mountApp,
  stubRuntime,
  type Msg,
  type Reply,
} from './helpers/ui-harness'

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
  // Pretend we're embedded in the OWA FAB iframe.
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
  window.location.hash = ''
})

const closes = (type: string) => posted.filter((m) => m.type === type).length

const tree: MailFolderNode[] = [
  {
    id: 'ROOT',
    displayName: 'Cases',
    path: 'Cases',
    children: [{ id: 'F1', displayName: 'Acme', path: 'Cases/Acme', children: [] }],
  },
]
const theRule: Rule = {
  id: 'r1',
  type: 'domain',
  signal: 'acme.example',
  targetFolderId: 'F1',
  targetFolderPath: 'Cases/Acme',
  confidence: 0.9,
  matchCount: 2,
  enabled: true,
  createdAt: '2026-10-01T00:00:00.000Z',
  source: 'user_manual',
}

function optionsSw(extra?: (m: Msg) => Reply | undefined) {
  return stubRuntime((m) => {
    const custom = extra?.(m)
    if (custom) return custom
    if (m.type === 'getStatus') {
      return {
        ok: true,
        data: {
          owaConnected: true,
          tokenValid: true,
          apiKeyConfigured: true,
          excludePrefixes: [],
          closedCasePrefixes: [],
          internalDomains: [],
          internalSubjectCategories: [],
          recentActivityIncludePrefixes: [],
          recentActivityIncludeLeafNames: [],
          primaryRootPath: '',
          syncEnabled: false,
        },
      }
    }
    if (m.type === 'getRules') return { ok: true, data: { rules: [theRule], conflicts: [] } }
    if (m.type === 'getFolderTree') return { ok: true, data: { tree, count: 2 } }
    return undefined
  })
}

async function renderOptions(hash: string, extra?: (m: Msg) => Reply | undefined) {
  window.location.hash = hash
  const s = optionsSw(extra)
  const { default: App } = await import('@/options/App')
  const r = await mountApp(App)
  root = r.root
  return { host: r.host, s }
}

describe('options: Esc closes the innermost layer only', () => {
  it('篩選 popover', async () => {
    const { host } = await renderOptions('#rules-library/all')
    click(await vi.waitFor(() => buttonContaining(host, '篩選')!))
    expect(document.querySelector('[aria-label="篩選條件"]')).not.toBeNull()

    keydown(document.body, 'Escape')
    expect(document.querySelector('[aria-label="篩選條件"]')).toBeNull()
    expect(closes('mail-organizer/close-options')).toBe(0)

    keydown(document.body, 'Escape') // nothing open now → close the modal
    expect(closes('mail-organizer/close-options')).toBe(1)
  })

  it('enable-sync modal', async () => {
    const { host } = await renderOptions('#data', (m) =>
      m.type === 'getSyncStatus'
        ? {
            ok: true,
            data: { enabled: false, lastSyncAt: '', machineId: 'm1', bytesInUse: 0, bytesQuota: 102400 },
          }
        : undefined,
    )
    click(await vi.waitFor(() => buttonByText(host, '啟用同步')!))
    expect(document.querySelector('[aria-labelledby="enable-sync-title"]')).not.toBeNull()

    keydown(document.body, 'Escape')
    expect(document.querySelector('[aria-labelledby="enable-sync-title"]')).toBeNull()
    expect(closes('mail-organizer/close-options')).toBe(0)
  })

  it('FolderPicker dropdown inside the rule drawer', async () => {
    const { host } = await renderOptions('#rules-library/all')
    const cell = await vi.waitFor(() => {
      const td = Array.from(host.querySelectorAll('td')).find((x) => x.textContent?.includes('acme.example'))
      expect(td).toBeTruthy()
      return td!
    })
    click(cell)
    const drawer = await vi.waitFor(() => document.querySelector('[aria-label="規則詳細"]') as HTMLElement)
    click(await vi.waitFor(() => {
      const b = buttonContaining(drawer, '編輯')
      expect(b?.disabled).toBe(false)
      return b!
    }))
    const picker = await vi.waitFor(() => drawer.querySelector<HTMLInputElement>('input[role="combobox"]')!)
    picker.focus()
    await vi.waitFor(() => expect(picker.getAttribute('aria-expanded')).toBe('true'))

    keydown(picker, 'Escape') // closes the dropdown only
    expect(picker.getAttribute('aria-expanded')).toBe('false')
    expect(document.querySelector('[aria-label="規則詳細"]')).not.toBeNull()

    keydown(picker, 'Escape') // still typing in a field — the edit survives
    expect(document.querySelector('[aria-label="規則詳細"]')).not.toBeNull()
    expect(closes('mail-organizer/close-options')).toBe(0)
  })
})

describe('popup: Esc with the ? shortcuts dialog open', () => {
  it('closes the dialog, not the whole OWA panel', async () => {
    stubRuntime((m) => {
      if (m.type === 'getStatus') {
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
      }
      if (m.type === 'getOnboardingState') return { ok: true, data: { needed: false } }
      if (m.type === 'getFolderActivity') return { ok: true, data: { activity: [] } }
      return undefined
    })
    const { default: App } = await import('@/popup/App')
    const r = await mountApp(App)
    root = r.root

    keydown(document.body, '?')
    await vi.waitFor(() => expect(document.querySelector('[aria-label="鍵盤捷徑"]')).not.toBeNull())
    keydown(document.body, 'Escape')
    await vi.waitFor(() => expect(document.querySelector('[aria-label="鍵盤捷徑"]')).toBeNull())
    expect(closes('mail-organizer/close-panel')).toBe(0)

    keydown(document.body, 'Escape')
    expect(closes('mail-organizer/close-panel')).toBe(1)
  })
})
