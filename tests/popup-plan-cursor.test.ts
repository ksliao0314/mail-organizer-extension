// Plan-screen keyboard cursor, round 2 (2026-11-06 review U1 / U2).
//
// U1 — keyboard-only use. The cursor never followed Tab: Tab to row 1 +
// Enter set it there (the activation's click bubbled to the row), and from
// then on Enter / Space went to row 1 whatever was focused — Tab to row 2 +
// Enter collapsed row 1, Space on row 2's checkbox selected row 1, and the
// in-list controls that aren't rows (全選此組, 展開 / 收合同對話) could not be
// activated by keyboard at all. Now Tab moves the cursor onto the row it
// lands on, and a Tabbed-to non-row control keeps its own Enter / Space.
// Focus a mouse click left behind still routes to the cursor row (E1).
//
// U2 — the cursor was a list INDEX. With 依資料夾 sorting an AI chunk merged
// mid-review sorts in above it, and a conversation cluster that splits shows
// its hidden siblings; the ring stayed on the index, the emails moved, and
// the next d / s / x / Enter hit a different email. It is an email now.

import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, PlanItem } from '@/shared/types'
import {
  buttonByText,
  buttonContaining,
  cleanup,
  click,
  keydown,
  mountApp,
  stubRuntime,
  typeInto,
} from './helpers/ui-harness'

const tree: MailFolderNode[] = [
  {
    id: 'ROOT',
    displayName: 'Cases',
    path: 'Cases',
    children: ['A', 'B', 'C'].map((n) => ({ id: `F${n}`, displayName: n, path: `Cases/${n}`, children: [] })),
  },
]

function item(over: Partial<PlanItem>): PlanItem {
  return {
    emailId: 'e1',
    emailSubject: 'subject',
    emailFrom: 'a@acme.example',
    action: 'move',
    targetFolderId: 'FA',
    targetFolderPath: 'Cases/A',
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}

const STATUS = {
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
}

type AiFeed = { runId: string; poll: () => PlanItem[] }

function planSw(items: PlanItem[], ai?: AiFeed) {
  return stubRuntime((m) => {
    switch (m.type) {
      case 'getStatus':
        return { ok: true, data: STATUS }
      case 'getOnboardingState':
        return { ok: true, data: { needed: false } }
      case 'getFolderActivity':
        return { ok: true, data: { activity: [] } }
      case 'getExecuteState':
        return { ok: true, data: { state: null } }
      case 'getAiClassifyProgress': {
        if (!ai) return { ok: true, data: null }
        const aiPlan = m.view === 'poll' ? ai.poll() : []
        return {
          ok: true,
          data: {
            runId: ai.runId,
            totalEmails: 5,
            completedEmails: aiPlan.length,
            chunks: 2,
            completedChunks: aiPlan.length > 0 ? 1 : 0,
            ruleCount: 0,
            aiCount: aiPlan.length,
            aiUnresolvedCount: 0,
            aiUnclassifiedCount: 0,
            aiCursor: aiPlan.length,
            startedAt: Date.now(),
            done: false,
            aiPlan,
            rulePlan: [],
            folderTree: tree,
          },
        }
      }
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
              aiPendingFlag: !!ai,
              runId: ai?.runId ?? null,
            },
          },
        }
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

async function renderPlan(items: PlanItem[], ai?: AiFeed) {
  planSw(items, ai)
  const { default: App } = await import('@/popup/App')
  const r = await mountApp(App)
  root = r.root
  await vi.waitFor(() => expect(r.host.querySelector('[data-plan-row-index="1"]')).not.toBeNull())
  return r.host
}

const rows = (host: HTMLElement) => Array.from(host.querySelectorAll<HTMLElement>('[data-plan-row-index]'))
const row = (host: HTMLElement, subject: string) => {
  const r = rows(host).find((el) => el.textContent?.includes(subject))
  if (!r) throw new Error(`no row for ${subject}`)
  return r
}
const toggleOf = (r: HTMLElement) => r.querySelector<HTMLButtonElement>('button')!
const checkboxOf = (r: HTMLElement) => r.querySelector<HTMLInputElement>('input[type="checkbox"]')!
const hasCursor = (r: HTMLElement) => r.className.includes('ring-2')
const isExpanded = (r: HTMLElement) => !!buttonContaining(r, '收起')
const cursorRows = (host: HTMLElement) => rows(host).filter(hasCursor)

/** Keyboard focus: the Tab keydown fires where focus is, then focus moves. */
function tabTo(el: HTMLElement) {
  keydown(document.activeElement ?? document.body, 'Tab')
  el.focus()
}

/** A real mouse click: pointer down (focus moves), then the click. */
function mouseClick(el: HTMLElement) {
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
  el.focus()
  click(el)
}

/**
 * A key press as the browser handles it: keydown, and unless a handler
 * cancelled it, the focused button / checkbox's native activation (a click).
 */
function press(el: HTMLElement, key: 'Enter' | ' ') {
  const ev = keydown(el, key)
  const activates =
    el.tagName === 'BUTTON' || (el instanceof HTMLInputElement && el.type === 'checkbox' && key === ' ')
  if (!ev.defaultPrevented && activates) click(el)
  return ev
}

const ITEMS = [
  item({ emailId: 'e1', emailSubject: 'first mail' }),
  item({ emailId: 'e2', emailSubject: 'second mail' }),
  item({ emailId: 'e3', emailSubject: 'third mail' }),
]

describe('U1: the cursor follows keyboard focus', () => {
  it('Tab to row 1 + Enter, then Tab on to row 2 + Enter: each Enter opens the row Tab landed on', async () => {
    const host = await renderPlan(ITEMS)
    const r1 = row(host, 'first mail')
    const r2 = row(host, 'second mail')

    tabTo(toggleOf(r1))
    await vi.waitFor(() => expect(hasCursor(r1)).toBe(true))
    press(toggleOf(r1), 'Enter')
    await vi.waitFor(() => expect(isExpanded(r1)).toBe(true))

    // On through row 1's open detail panel (its controls keep the cursor on
    // row 1) to row 2's checkbox and toggle.
    tabTo(buttonByText(r1, '刪除')!)
    expect(hasCursor(r1)).toBe(true)
    tabTo(checkboxOf(r2))
    tabTo(toggleOf(r2))
    await vi.waitFor(() => expect(hasCursor(r2)).toBe(true))
    expect(hasCursor(r1)).toBe(false)

    press(toggleOf(r2), 'Enter')
    await vi.waitFor(() => expect(isExpanded(r2)).toBe(true))
    expect(isExpanded(r1)).toBe(true) // row 1 was not collapsed under the user
  })

  it('Space on a Tab-focused row checkbox selects THAT row, even with the cursor elsewhere', async () => {
    const host = await renderPlan(ITEMS)
    const r1 = row(host, 'first mail')
    const r2 = row(host, 'second mail')
    click(r1) // mouse: cursor on row 1
    await vi.waitFor(() => expect(hasCursor(r1)).toBe(true))

    tabTo(checkboxOf(r2))
    await vi.waitFor(() => expect(hasCursor(r2)).toBe(true))
    press(checkboxOf(r2), ' ')
    await vi.waitFor(() => expect(checkboxOf(r2).checked).toBe(true))
    expect(checkboxOf(r1).checked).toBe(false)
  })

  it('a window regaining focus (no Tab) does not pull the cursor back to the focused row', async () => {
    const host = await renderPlan(ITEMS)
    const r1 = row(host, 'first mail')
    const r3 = row(host, 'third mail')
    tabTo(toggleOf(r1))
    await vi.waitFor(() => expect(hasCursor(r1)).toBe(true))
    keydown(toggleOf(r1), 'j')
    keydown(toggleOf(r1), 'j')
    await vi.waitFor(() => expect(hasCursor(r3)).toBe(true))

    // Switching away and back re-fires focusin on the same element.
    toggleOf(r1).blur()
    toggleOf(r1).focus()
    await new Promise((r) => setTimeout(r, 20))
    expect(hasCursor(r3)).toBe(true)
    // …and Enter still means the cursor row (the old focus no longer owns it).
    expect(keydown(toggleOf(r1), 'Enter').defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(isExpanded(r3)).toBe(true))
    expect(isExpanded(r1)).toBe(false)
  })

  it('Tabbed-to list controls that are not rows keep their own Enter / Space while a cursor exists', async () => {
    const host = await renderPlan([
      item({ emailId: 'c1', emailSubject: 'thread one', conversationId: 'conv-x' }),
      item({ emailId: 'c2', emailSubject: 'thread two', conversationId: 'conv-x' }),
      item({ emailId: 'e3', emailSubject: 'solo mail', targetFolderId: 'FB', targetFolderPath: 'Cases/B' }),
    ])
    click(buttonContaining(host, '依資料夾')!) // group headers with 全選此組
    const solo = row(host, 'solo mail')
    click(solo) // cursor on the solo row
    await vi.waitFor(() => expect(hasCursor(row(host, 'solo mail'))).toBe(true))

    const expand = buttonContaining(host, '還有 1 封同對話')!
    tabTo(expand)
    expect(press(expand, 'Enter').defaultPrevented).toBe(false)
    await vi.waitFor(() => expect(host.textContent).toContain('thread two'))
    expect(isExpanded(row(host, 'solo mail'))).toBe(false) // the cursor row was left alone

    const collapse = buttonContaining(host, '收合 2 封同對話')!
    tabTo(collapse)
    expect(press(collapse, 'Enter').defaultPrevented).toBe(false)
    await vi.waitFor(() => expect(host.textContent).not.toContain('thread two'))

    const selectGroup = buttonByText(host, '全選此組')! // group Cases/A: c1 + c2
    tabTo(selectGroup)
    expect(press(selectGroup, ' ').defaultPrevented).toBe(false)
    await vi.waitFor(() => expect(checkboxOf(row(host, 'thread one')).checked).toBe(true))
    expect(checkboxOf(row(host, 'solo mail')).checked).toBe(false)
  })

  it('focus a MOUSE click left on a list control still routes Space to the cursor row (E1)', async () => {
    const host = await renderPlan([
      item({ emailId: 'e1', emailSubject: 'first mail' }),
      item({ emailId: 'e2', emailSubject: 'second mail', targetFolderId: 'FB', targetFolderPath: 'Cases/B' }),
      item({ emailId: 'e3', emailSubject: 'third mail', targetFolderId: 'FC', targetFolderPath: 'Cases/C' }),
    ])
    click(buttonContaining(host, '依資料夾')!)
    click(row(host, 'first mail'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'first mail'))).toBe(true))

    const groups = Array.from(host.querySelectorAll('button')).filter((b) => b.textContent?.trim() === '全選此組')
    mouseClick(groups[1]!) // group Cases/B → selects 'second mail'; focus stays on the button
    await vi.waitFor(() => expect(checkboxOf(row(host, 'second mail')).checked).toBe(true))

    expect(press(groups[1]!, ' ').defaultPrevented).toBe(true)
    await vi.waitFor(() => expect(checkboxOf(row(host, 'first mail')).checked).toBe(true))
  })
})

describe('U2: the cursor stays on its email when rows move', () => {
  it('an AI chunk sorting in ABOVE the cursor (依資料夾) — d still hits the cursor email', async () => {
    let merged = false
    const host = await renderPlan(
      [
        item({ emailId: 'e1', emailSubject: 'bravo mail', targetFolderId: 'FB', targetFolderPath: 'Cases/B' }),
        item({ emailId: 'e2', emailSubject: 'charlie mail', targetFolderId: 'FC', targetFolderPath: 'Cases/C' }),
        item({ emailId: 'e4', emailSubject: 'delta mail', targetFolderId: 'FC', targetFolderPath: 'Cases/C' }),
      ],
      {
        runId: 'run-x',
        poll: () => (merged ? [item({ emailId: 'e0', emailSubject: 'alpha mail' })] : []),
      },
    )
    click(buttonContaining(host, '依資料夾')!)
    click(row(host, 'charlie mail'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'charlie mail'))).toBe(true))

    merged = true // the next poll (500 ms) brings 'alpha mail' → Cases/A, sorted first
    await vi.waitFor(() => expect(host.textContent).toContain('alpha mail'), { timeout: 3000 })
    expect(row(host, 'alpha mail').getAttribute('data-plan-row-index')).toBe('0')
    expect(cursorRows(host)).toEqual([row(host, 'charlie mail')])

    keydown(document.body, 'd')
    await vi.waitFor(() => expect(row(host, 'charlie mail').textContent).toContain('已刪除的郵件'))
    expect(row(host, 'bravo mail').textContent).not.toContain('已刪除的郵件')
  })

  it('a conversation cluster that splits under the cursor — d still hits the cursor email', async () => {
    const host = await renderPlan([
      item({ emailId: 'c1', emailSubject: 'thread one', conversationId: 'conv-x' }),
      item({ emailId: 'c2', emailSubject: 'thread two', conversationId: 'conv-x' }),
      item({ emailId: 'e3', emailSubject: 'solo mail' }),
    ])
    // c2 is folded under c1 (same conversation, same decision).
    expect(host.textContent).not.toContain('thread two')
    click(checkboxOf(row(host, 'thread one'))) // select c1 (no cursor move)
    click(row(host, 'solo mail'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'solo mail'))).toBe(true))

    // 全部保留 on c1 only: the cluster no longer agrees → c2 shows up above the cursor row.
    click(buttonByText(host, '全部保留')!)
    await vi.waitFor(() => expect(host.textContent).toContain('thread two'))
    expect(cursorRows(host)).toEqual([row(host, 'solo mail')])

    keydown(document.body, 'd')
    await vi.waitFor(() => expect(row(host, 'solo mail').textContent).toContain('已刪除的郵件'))
    expect(row(host, 'thread two').textContent).not.toContain('已刪除的郵件')
  })

  it('a sibling folded away by 收合 hands the cursor to its cluster row; a filtered-out email drops it', async () => {
    const host = await renderPlan([
      item({ emailId: 'c1', emailSubject: 'thread one', conversationId: 'conv-x' }),
      item({ emailId: 'c2', emailSubject: 'thread two', conversationId: 'conv-x' }),
      item({ emailId: 'e3', emailSubject: 'solo mail', source: 'unresolved', confidence: 0.2 }),
    ])
    click(buttonContaining(host, '還有 1 封同對話')!)
    click(row(host, 'thread two'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'thread two'))).toBe(true))
    click(buttonContaining(host, '收合 2 封同對話')!)
    await vi.waitFor(() => expect(host.textContent).not.toContain('thread two'))
    expect(cursorRows(host)).toEqual([row(host, 'thread one')])

    click(buttonContaining(host, '待確認 1')!) // only 'solo mail' stays
    await vi.waitFor(() => expect(host.textContent).not.toContain('thread one'))
    expect(cursorRows(host)).toEqual([])
    click(buttonContaining(host, '待確認 1')!) // back to all: the cursor does not come back
    await vi.waitFor(() => expect(host.textContent).toContain('thread one'))
    expect(cursorRows(host)).toEqual([])
  })
})

// Round 3 (2026-11-06): following the email is right when the list reorders
// AROUND the cursor (the U2 cases above). When the user's own action on the
// cursor row moves that row — d / s / a retarget into another 依資料夾
// group, out of the filter — the ring used to follow it there: off-screen
// into the delete group, where the next d did nothing (already 刪除) and j
// went on from there; under a filter it vanished and j restarted at the top.
// Now the cursor moves on to the next mail, as f58c9dd's index cursor did.
describe('round 3: the user’s own action on the cursor row moves the cursor ON', () => {
  const GROUPED = [
    item({ emailId: 'a0', emailSubject: 'alpha zero' }),
    item({ emailId: 'a1', emailSubject: 'alpha one' }),
    item({ emailId: 'a2', emailSubject: 'alpha two' }),
    item({ emailId: 'b0', emailSubject: 'bravo zero', targetFolderId: 'FB', targetFolderPath: 'Cases/B' }),
  ]
  const isDeleted = (r: HTMLElement) => (r.textContent ?? '').includes('已刪除的郵件')

  it('依資料夾: d sends the row to the delete group — the cursor stays put on the next mail; d / j go on from there', async () => {
    const host = await renderPlan(GROUPED)
    click(buttonContaining(host, '依資料夾')!)
    click(row(host, 'alpha one'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'alpha one'))).toBe(true))

    keydown(document.body, 'd')
    await vi.waitFor(() => expect(isDeleted(row(host, 'alpha one'))).toBe(true))
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'alpha two')]))
    expect(row(host, 'alpha two').getAttribute('data-plan-row-index')).toBe('1') // same place on screen

    keydown(document.body, 'd') // acts on the next mail — no longer a no-op on the deleted one
    await vi.waitFor(() => expect(isDeleted(row(host, 'alpha two'))).toBe(true))
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'bravo zero')]))
    keydown(document.body, 'k')
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'alpha zero')]))
    expect(isDeleted(row(host, 'alpha zero'))).toBe(false)
    expect(isDeleted(row(host, 'bravo zero'))).toBe(false)
  })

  it('a filter the edit takes the row out of — the cursor moves to the next mail, not back to the top', async () => {
    const host = await renderPlan([
      item({ emailId: 'm0', emailSubject: 'memo zero', emailFrom: 'x@client.example' }),
      item({ emailId: 'm1', emailSubject: 'memo one', emailFrom: 'x@client.example' }),
      item({ emailId: 'm2', emailSubject: 'memo two', emailFrom: 'x@client.example' }),
      item({ emailId: 'x0', emailSubject: 'other', emailFrom: 'x@client.example', targetFolderId: 'FB', targetFolderPath: 'Cases/B' }),
    ])
    keydown(document.body, '/')
    const input = await vi.waitFor(() => {
      const el = host.querySelector<HTMLInputElement>('input[aria-label="過濾條件"]')
      expect(el).not.toBeNull()
      return el!
    })
    typeInto(input, 'cases/a') // matches the target path only
    keydown(input, 'Enter')
    input.blur()
    await vi.waitFor(() => expect(host.textContent).not.toContain('other'))

    click(row(host, 'memo one'))
    await vi.waitFor(() => expect(hasCursor(row(host, 'memo one'))).toBe(true))
    keydown(document.body, 'd') // no target any more → drops out of the filter
    await vi.waitFor(() => expect(host.textContent).not.toContain('memo one'))
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'memo two')]))

    keydown(document.body, 'd')
    await vi.waitFor(() => expect(host.textContent).not.toContain('memo two'))
    expect(host.textContent).toContain('刪 2')
    expect(isDeleted(row(host, 'memo zero'))).toBe(false)
  })

  it('a click on 保留 in the row’s own panel, and 全部刪除 with the cursor row selected, move it on too', async () => {
    const host = await renderPlan(GROUPED)
    click(buttonContaining(host, '依資料夾')!)
    const r1 = row(host, 'alpha one')
    click(r1.querySelector('button')!) // expand (mouse) — the cursor lands here
    await vi.waitFor(() => expect(buttonByText(r1, '保留')).toBeDefined())
    click(buttonByText(r1, '保留')!)
    await vi.waitFor(() => expect(row(host, 'alpha one').textContent).toContain('保留在收件夾'))
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'alpha two')]))

    // Bulk: the cursor row and the one after it go together → the next UNCHANGED mail.
    click(checkboxOf(row(host, 'alpha two')))
    click(checkboxOf(row(host, 'bravo zero')))
    click(buttonByText(host, '全部刪除')!)
    await vi.waitFor(() => expect(isDeleted(row(host, 'alpha two'))).toBe(true))
    await vi.waitFor(() => expect(cursorRows(host)).toEqual([row(host, 'alpha one')]))
  })

  it('an edit that leaves the row where it is keeps the cursor on it (default order)', async () => {
    const host = await renderPlan(GROUPED)
    click(row(host, 'alpha one'))
    keydown(document.body, 'd')
    await vi.waitFor(() => expect(isDeleted(row(host, 'alpha one'))).toBe(true))
    expect(cursorRows(host)).toEqual([row(host, 'alpha one')])
  })
})
