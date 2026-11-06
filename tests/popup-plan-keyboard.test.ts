// Plan screen: Enter / Space vs the keyboard cursor (2026-11-06 audit E1).
//
// The j/k cursor is virtual (it never moves DOM focus), but Chrome leaves
// focus on whatever button the mouse last clicked — and nearly the whole plan
// row is a <button>. v1.0.8 made a focused button own Enter / Space (so Enter
// on 執行 executes), which in practice meant:
//   - Space never selected the cursor row, although the ? help says
//     「x / Space 勾選」;
//   - click row 1 → j → Enter toggled row 1, not the cursor row 2;
//   - after clicking 「待確認」 or 「依資料夾」, Space / Enter flipped that
//     filter / sort again — and 「待確認」 didn't reset the cursor, so the
//     ring sat on an index that now held another email;
//   - a clicked row checkbox (an INPUT) swallowed every shortcut.
// Inside the list the cursor row wins again; controls outside it (執行) keep
// their native activation.

import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, PlanItem } from '@/shared/types'
import { focusOwnsPlanActivation, preventsPointerFocus } from '@/popup/keyboard'
import {
  buttonByText,
  buttonContaining,
  cleanup,
  click,
  keydown,
  mountApp,
  stubRuntime,
  type Reply,
} from './helpers/ui-harness'

const tree: MailFolderNode[] = [
  {
    id: 'ROOT',
    displayName: 'Cases',
    path: 'Cases',
    children: [{ id: 'F1', displayName: 'Acme', path: 'Cases/Acme', children: [] }],
  },
]

function item(over: Partial<PlanItem>): PlanItem {
  return {
    emailId: 'e1',
    emailSubject: 'subject',
    emailFrom: 'a@acme.example',
    action: 'move',
    targetFolderId: 'F1',
    targetFolderPath: 'Cases/Acme',
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}

const ITEMS = [
  item({ emailId: 'e1', emailSubject: 'first mail' }),
  item({ emailId: 'e2', emailSubject: 'second mail' }),
  item({ emailId: 'e3', emailSubject: 'third mail', source: 'unresolved', confidence: 0.2 }),
]

function planSw(items: PlanItem[]) {
  return stubRuntime((m) => {
    switch (m.type) {
      case 'getStatus':
        return {
          ok: true,
          data: {
            owaConnected: true,
            tokenValid: true,
            apiKeyConfigured: true,
            apiKeyPreview: 'sk-…',
            model: 'claude-sonnet-4-6',
            folderCacheAgeMs: null,
            folderCacheCount: 0,
            rulesCount: 0,
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
            plan: {
              items,
              folderTree: tree,
              usage: null,
              summary: { ruleHits: 0, aiHandled: items.length },
              banner: null,
            },
          },
        } satisfies Reply
      case 'savePopupState':
      case 'clearPopupState':
        return { ok: true }
      default:
        return undefined
    }
  })
}

let root: Root | null = null
afterEach(() => {
  cleanup(root)
  root = null
})

async function renderPlan(items = ITEMS) {
  planSw(items)
  const { default: App } = await import('@/popup/App')
  const r = await mountApp(App)
  root = r.root
  await vi.waitFor(() => expect(r.host.querySelector('[data-plan-row-index="1"]')).not.toBeNull())
  return r.host
}

const row = (host: HTMLElement, i: number) =>
  host.querySelector<HTMLElement>(`[data-plan-row-index="${i}"]`)!
/** The big expand toggle that makes up most of a row. */
const toggleOf = (r: HTMLElement) => r.querySelector<HTMLButtonElement>('button')!
const checkboxOf = (r: HTMLElement) => r.querySelector<HTMLInputElement>('input[type="checkbox"]')!
const hasCursor = (r: HTMLElement) => r.className.includes('ring-2')
const isExpanded = (r: HTMLElement) => !!buttonContaining(r, '收起')

/** A real mouse click on a button: Chrome focuses it, then clicks. */
function mouseClick(el: HTMLElement) {
  el.focus()
  click(el)
}

function mousedown(el: HTMLElement) {
  const ev = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
  el.dispatchEvent(ev)
  return ev
}

describe('Enter / Space on the plan list follow the keyboard cursor', () => {
  it('Space with focus left on a row toggle selects the cursor row (as the ? help says)', async () => {
    const host = await renderPlan()
    const r2 = row(host, 1)
    mouseClick(toggleOf(r2)) // sets the cursor on row 2 and leaves focus on its toggle
    await vi.waitFor(() => expect(hasCursor(r2)).toBe(true))
    expect(document.activeElement).toBe(toggleOf(r2))

    const ev = keydown(toggleOf(r2), ' ')
    expect(ev.defaultPrevented).toBe(true) // the button's own click must not also fire
    await vi.waitFor(() => expect(checkboxOf(r2).checked).toBe(true))
  })

  it('click row 1 → j → Enter expands the cursor row 2, not the last-clicked row', async () => {
    const host = await renderPlan()
    const r1 = row(host, 0)
    const r2 = row(host, 1)
    mouseClick(toggleOf(r1))
    await vi.waitFor(() => expect(isExpanded(r1)).toBe(true))
    keydown(toggleOf(r1), 'j')
    await vi.waitFor(() => expect(hasCursor(r2)).toBe(true))

    const ev = keydown(toggleOf(r1), 'Enter')
    expect(ev.defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(isExpanded(r2)).toBe(true))
    expect(isExpanded(r1)).toBe(true) // row 1 untouched
  })

  it('a focused row checkbox no longer swallows the shortcuts; Space acts on the cursor row', async () => {
    const host = await renderPlan()
    const r1 = row(host, 0)
    const r2 = row(host, 1)
    const cb2 = checkboxOf(r2)
    cb2.focus()
    keydown(cb2, 'j')
    await vi.waitFor(() => expect(hasCursor(r1)).toBe(true))

    const ev = keydown(cb2, ' ')
    expect(ev.defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(checkboxOf(r1).checked).toBe(true))
    expect(cb2.checked).toBe(false)
  })

  // (A real Tab onto a row moves the cursor there — popup-plan-cursor.test.ts.)
  it('with no cursor, a focused list control keeps its own Enter / Space', async () => {
    const host = await renderPlan()
    const t2 = toggleOf(row(host, 1))
    t2.focus()
    expect(keydown(t2, ' ').defaultPrevented).toBe(false)
    expect(keydown(t2, 'Enter').defaultPrevented).toBe(false)
  })

  it('a Tab-focused control in the cursor row’s open detail panel keeps Enter / Space; mouse clicks there take no focus', async () => {
    const host = await renderPlan()
    const r1 = row(host, 0)
    const r2 = row(host, 1)
    mouseClick(toggleOf(r1))
    await vi.waitFor(() => expect(isExpanded(r1)).toBe(true))
    const del = buttonByText(r1, '刪除')!
    expect(del.closest('[data-plan-row-detail]')).not.toBeNull()
    expect(mousedown(del).defaultPrevented).toBe(true) // a click won't park focus on it

    del.focus() // Tab
    expect(keydown(del, 'Enter').defaultPrevented).toBe(false)
    expect(keydown(del, ' ').defaultPrevented).toBe(false)
    await new Promise((r) => setTimeout(r, 20))
    expect(isExpanded(r1)).toBe(true)
    expect(checkboxOf(r1).checked).toBe(false)

    // Cursor moves to row 2 — now focus and cursor disagree, the cursor wins.
    keydown(del, 'j')
    await vi.waitFor(() => expect(hasCursor(r2)).toBe(true))
    expect(keydown(del, 'Enter').defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(isExpanded(r2)).toBe(true))
  })

  it('「待確認」 and 「依資料夾」 take no focus on a mouse click, and 「待確認」 resets the cursor', async () => {
    const host = await renderPlan()
    const chip = buttonContaining(host, '待確認 1')!
    const sort = buttonContaining(host, '依資料夾')!
    expect(mousedown(chip).defaultPrevented).toBe(true)
    expect(mousedown(sort).defaultPrevented).toBe(true)

    click(row(host, 1)) // cursor on 'second mail'
    await vi.waitFor(() => expect(hasCursor(row(host, 1))).toBe(true))
    click(chip) // only 'third mail' (needs attention) remains
    await vi.waitFor(() => expect(host.querySelector('[data-plan-row-index="1"]')).toBeNull())
    // The ring must not silently jump onto a different email.
    expect(host.querySelectorAll('.ring-2')).toHaveLength(0)

    // Focus still on the list (not the chip): Space → cursor row after j.
    keydown(document.body, 'j')
    await vi.waitFor(() => expect(hasCursor(row(host, 0))).toBe(true))
    keydown(document.body, ' ')
    await vi.waitFor(() => expect(checkboxOf(row(host, 0)).checked).toBe(true))
  })
})

describe('focusOwnsPlanActivation / preventsPointerFocus', () => {
  function fixture() {
    document.body.innerHTML = `
      <button id="exec">執行</button>
      <ul data-plan-list="">
        <li>
          <div data-plan-row-index="0">
            <input id="cb0" type="checkbox" />
            <button id="t0">row 0</button>
            <div data-plan-row-detail=""><button id="d0">刪除</button><input id="in0" /></div>
          </div>
        </li>
        <li><button id="hdr">全選此組</button></li>
      </ul>`
    return (id: string) => document.getElementById(id)!
  }

  it('routes Enter / Space by where focus is and where the cursor is', () => {
    const $ = fixture()
    expect(focusOwnsPlanActivation($('exec'), 0)).toBe(true) // outside the list
    expect(focusOwnsPlanActivation($('t0'), 0)).toBe(false) // row toggle → cursor row
    expect(focusOwnsPlanActivation($('t0'), null)).toBe(true) // no cursor → native
    expect(focusOwnsPlanActivation($('cb0'), 0)).toBe(false)
    expect(focusOwnsPlanActivation($('hdr'), 0)).toBe(false)
    expect(focusOwnsPlanActivation($('d0'), 0)).toBe(true) // detail of the cursor row
    expect(focusOwnsPlanActivation($('d0'), 1)).toBe(false) // detail of another row
    expect(focusOwnsPlanActivation(document.body, 0)).toBe(false)
  })

  it('only button-likes skip pointer focus — inputs and options keep it', () => {
    const $ = fixture()
    expect(preventsPointerFocus($('d0'))).toBe(true)
    expect(preventsPointerFocus($('in0'))).toBe(false)
    expect(preventsPointerFocus($('cb0'))).toBe(false)
  })
})
