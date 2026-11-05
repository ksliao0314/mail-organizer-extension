// Popup plan / done / idle screens (2026-11-05 audit). Renders the REAL
// popup <App/> against a stub SW (see helpers/ui-harness).
//
//   - Enter on a focused button is that button's activation: clicking a row
//     sets the keyboard cursor, and Enter on 執行 used to toggle the cursor
//     row instead (the plan shortcut handler only exempted INPUT/TEXTAREA)
//   - 「刪除」copy says what it does — a move to 已刪除的郵件 (it said
//     「永久刪除」 in the row while the footer promised it was recoverable)
//   - 重試 is in-flight guarded; ALREADY_RUNNING shows the running batch
//     instead of an error screen
//   - the token-wait poll really polls every second (it fired once) and the
//     banner shows the SW's tokenError reason

import { act } from 'react'
import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, PlanItem } from '@/shared/types'
import {
  buttonContaining,
  cleanup,
  click,
  keydown,
  mountApp,
  stubRuntime,
  type Msg,
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

function status(over: Record<string, unknown> = {}) {
  return {
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
    ...over,
  }
}

function popupSw(extra: (m: Msg) => Reply | undefined | Promise<Reply | undefined>) {
  return stubRuntime(async (m) => {
    const custom = await extra(m)
    if (custom) return custom
    switch (m.type) {
      case 'getStatus':
        return { ok: true, data: status() }
      case 'getOnboardingState':
        return { ok: true, data: { needed: false } }
      case 'getFolderActivity':
        return { ok: true, data: { activity: [] } }
      case 'getExecuteState':
        return { ok: true, data: { state: null } }
      case 'getAiClassifyProgress':
        return { ok: true, data: null }
      case 'getUndoSnapshot':
        return { ok: true, data: { snapshot: null } }
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

async function renderPopup() {
  const { default: App } = await import('@/popup/App')
  const r = await mountApp(App)
  root = r.root
  return r.host
}

function planState(items: PlanItem[]): Reply {
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
  }
}

describe('plan screen', () => {
  it('Enter on a focused button activates the button, not the keyboard-cursor row', async () => {
    popupSw((m) =>
      m.type === 'getPopupState'
        ? planState([item({ emailId: 'e1', emailSubject: 'first mail' }), item({ emailId: 'e2', emailSubject: 'second mail' })])
        : undefined,
    )
    const host = await renderPopup()
    const row = await vi.waitFor(() => {
      const el = Array.from(host.querySelectorAll<HTMLElement>('.group.relative')).find((x) =>
        x.textContent?.includes('second mail'),
      )
      expect(el).toBeTruthy()
      return el!
    })
    click(row) // sets the keyboard cursor on this row
    await vi.waitFor(() => expect(row.className).toContain('ring-2'))
    expect(buttonContaining(row, '收起')).toBeUndefined()

    const execute = buttonContaining(host, '執行')!
    execute.focus()
    keydown(execute, 'Enter')
    await new Promise((r) => setTimeout(r, 20))
    expect(buttonContaining(row, '收起')).toBeUndefined() // row did NOT toggle open

    // Navigation shortcuts still work from a focused button.
    keydown(execute, 'j')
    await vi.waitFor(() => expect(row.className).toContain('ring-2'))
  })

  it('describes 刪除 as a move to 已刪除的郵件 everywhere', async () => {
    popupSw((m) =>
      m.type === 'getPopupState'
        ? planState([item({ emailId: 'd1', emailSubject: 'spam', action: 'delete', targetFolderId: undefined, targetFolderPath: undefined })])
        : undefined,
    )
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('spam'))
    expect(host.textContent).not.toContain('永久刪除')
    expect(host.textContent).toContain('移到「已刪除的郵件」')
    expect(host.textContent).toContain('含 1 件移到「已刪除的郵件」')
  })
})

describe('done screen 重試', () => {
  it('ignores a double-click and shows the running batch on ALREADY_RUNNING', async () => {
    const doneState = {
      inProgress: false,
      cancelRequested: false,
      startedAt: 1,
      finishedAt: 2,
      total: 1,
      current: 1,
      results: [{ emailId: 'e1', subject: 'S', action: 'move', status: 'error', message: 'boom' }],
      summary: { moved: 0, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 1 },
      rulesAdded: 0,
    }
    let releaseRetry: (r: Reply) => void = () => {}
    let executing = false
    const s = popupSw((m) => {
      if (m.type === 'getExecuteState') {
        return {
          ok: true,
          data: { state: executing ? { ...doneState, inProgress: true, finishedAt: undefined } : doneState },
        }
      }
      if (m.type === 'retryFailed') return new Promise<Reply>((res) => (releaseRetry = res))
      return undefined
    })
    const host = await renderPopup()
    const retry = await vi.waitFor(() => buttonContaining(host, '重試')!)
    click(retry)
    click(retry)
    await vi.waitFor(() => expect(s.count('retryFailed')).toBe(1))
    expect(retry.disabled).toBe(true)

    executing = true
    await act(async () => {
      releaseRetry({ ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' })
    })
    await vi.waitFor(() => expect(host.textContent).toContain('執行中'))
    expect(host.textContent).not.toContain('ALREADY_RUNNING')
    expect(s.count('retryFailed')).toBe(1)
  })
})

describe('token wait', () => {
  it('polls getStatus every second for up to 10 s and shows the SW reason', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })
    try {
      const s = popupSw((m) =>
        m.type === 'getStatus'
          ? {
              ok: true,
              data: status({
                tokenValid: false,
                tokenError: { code: 'TAB_UNREACHABLE', message: 'no content script' },
              }),
            }
          : undefined,
      )
      const host = await renderPopup()
      await vi.waitFor(() => expect(s.count('getStatus')).toBe(1))
      await vi.waitFor(() =>
        expect(host.textContent).toContain('請重新整理 Outlook 分頁（擴充功能更新後需要）'),
      )

      await act(async () => {
        await vi.advanceTimersByTimeAsync(3100)
      })
      expect(s.count('getStatus')).toBeGreaterThanOrEqual(4) // initial + 3 polls

      await act(async () => {
        await vi.advanceTimersByTimeAsync(10_000)
      })
      const settled = s.count('getStatus')
      expect(settled).toBeLessThanOrEqual(11)
      await act(async () => {
        await vi.advanceTimersByTimeAsync(5_000)
      })
      expect(s.count('getStatus')).toBe(settled) // stopped at the ceiling
    } finally {
      vi.useRealTimers()
    }
  })
})
