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

// token.ts's message for a tab the browser discarded (Memory Saver) — the fix
// is to switch to the tab, not to reload it.
const SLEEPING_TAB_MESSAGE =
  'Outlook 分頁已被瀏覽器休眠(省記憶體)。請切換到該分頁讓它重新載入,再試一次。'

/** The header pill's label per token error (2026-11-06 review U11d). */
const PILL: Record<string, string> = {
  TAB_UNREACHABLE: 'Outlook 分頁沒有回應',
  NOT_ON_OWA: 'Outlook 分頁需要重新整理',
  MULTIPLE_ACCOUNTS: '多個 Outlook 帳號',
}

describe('token issue text (2026-11-06 audit E3)', () => {
  async function bannerFor(tokenError: { code: string; message: string }) {
    popupSw((m) =>
      m.type === 'getStatus' ? { ok: true, data: status({ tokenValid: false, tokenError }) } : undefined,
    )
    const host = await renderPopup()
    await vi.waitFor(() =>
      expect(host.textContent).toContain(PILL[tokenError.code] ?? 'Outlook 登入過期'),
    )
    return host
  }

  it('the header pill names the actual problem — 「登入過期」 only for a missing / expired token (U11d)', async () => {
    const host = await bannerFor({ code: 'TAB_UNREACHABLE', message: SLEEPING_TAB_MESSAGE })
    expect(host.textContent).toContain('Outlook 分頁沒有回應')
    expect(host.textContent).not.toContain('Outlook 登入過期')
    cleanup(root)
    root = null
    const host2 = await bannerFor({ code: 'NO_TOKEN', message: '' })
    expect(host2.textContent).toContain('Outlook 登入過期')
  })

  it("shows the SW's own TAB_UNREACHABLE reason (sleeping tab → switch to it) instead of a generic reload hint", async () => {
    const host = await bannerFor({ code: 'TAB_UNREACHABLE', message: SLEEPING_TAB_MESSAGE })
    await vi.waitFor(() => expect(host.textContent).toContain(SLEEPING_TAB_MESSAGE))
    expect(host.textContent).not.toContain('請重新整理 Outlook 分頁（擴充功能更新後需要）')
  })

  it('falls back to the reload hint when TAB_UNREACHABLE has no message, and for NOT_ON_OWA', async () => {
    const host = await bannerFor({ code: 'TAB_UNREACHABLE', message: '' })
    await vi.waitFor(() =>
      expect(host.textContent).toContain('請重新整理 Outlook 分頁（擴充功能更新後需要）'),
    )
    cleanup(root)
    root = null
    const host2 = await bannerFor({ code: 'NOT_ON_OWA', message: 'Content script 沒回應' })
    await vi.waitFor(() =>
      expect(host2.textContent).toContain('請重新整理 Outlook 分頁（擴充功能更新後需要）'),
    )
    expect(host2.textContent).not.toContain('Content script')
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
                tokenError: { code: 'TAB_UNREACHABLE', message: SLEEPING_TAB_MESSAGE },
              }),
            }
          : undefined,
      )
      const host = await renderPopup()
      await vi.waitFor(() => expect(s.count('getStatus')).toBe(1))
      await vi.waitFor(() => expect(host.textContent).toContain(SLEEPING_TAB_MESSAGE))

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

describe('executing progress (2026-11-06 review U11c)', () => {
  it('counts finished items — not the index of the item a worker started last', async () => {
    // new_folder items run first and 3 workers run in parallel: item 2 (the
    // new_folder) is done, a worker just picked up item 0 → current = 0.
    const running = {
      inProgress: true,
      cancelRequested: false,
      startedAt: Date.now() - 2000,
      total: 3,
      current: 0,
      results: [
        { emailId: 'a', subject: 'mail a', action: 'move', status: 'processing' },
        { emailId: 'b', subject: 'mail b', action: 'move', status: 'queued' },
        { emailId: 'c', subject: 'mail c', action: 'new_folder', status: 'folder_created' },
      ],
      summary: { moved: 1, deleted: 0, foldersCreated: 1, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    popupSw((m) => (m.type === 'getExecuteState' ? { ok: true, data: { state: running } } : undefined))
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('執行中'))
    expect(host.textContent).toContain('1 / 3')
    expect(host.textContent).toContain('33%')
    expect(host.textContent).not.toContain('0 / 3')
  })
})

describe('「AI N」 in the plan summary (2026-11-06 review U11e)', () => {
  const refused = item({
    emailId: 'r1',
    emailSubject: 'refused mail',
    action: 'skip',
    source: 'unresolved',
    confidence: 0,
    targetFolderId: undefined,
    targetFolderPath: undefined,
  })
  const decided = item({ emailId: 'd1', emailSubject: 'decided mail' })
  const run = {
    runId: 'run-a',
    totalEmails: 2,
    completedEmails: 2,
    chunks: 1,
    completedChunks: 1,
    ruleCount: 0,
    aiCount: 2,
    aiUnresolvedCount: 1,
    aiUnclassifiedCount: 0,
    aiCursor: 2,
    startedAt: Date.now() - 60_000,
    done: true,
    rulePlan: [],
    aiPlan: [decided, refused],
    folderTree: tree,
    aiError: null,
  }

  it('reopening onto a finished run does not count the refused row', async () => {
    popupSw((m) => (m.type === 'getAiClassifyProgress' ? { ok: true, data: run } : undefined))
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('refused mail'))
    expect(host.textContent).toContain('規則 0 · AI 1')
  })

  it('繼續歸檔 onto the run counts it the same way', async () => {
    const done = {
      inProgress: false,
      cancelRequested: false,
      startedAt: Date.now() - 120_000,
      finishedAt: Date.now() - 100_000,
      total: 1,
      current: 1,
      results: [{ emailId: 'm0', subject: 'm0', action: 'move', status: 'moved' }],
      summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
      rulesAdded: 0,
    }
    popupSw((m) => {
      if (m.type === 'getExecuteState') return { ok: true, data: { state: done } }
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: run }
      if (m.type === 'peekNextBatch') return { ok: true, data: { eligibleCount: 2, cappedAtBatchSize: false } }
      if (m.type === 'clearExecuteState' || m.type === 'dismissUndo' || m.type === 'adoptClassifyRun') {
        return { ok: true }
      }
      return undefined
    })
    const host = await renderPopup()
    const go = await vi.waitFor(() => buttonContaining(host, '繼續歸檔下一批')!)
    await act(async () => go.click())
    await vi.waitFor(() => expect(host.textContent).toContain('refused mail'))
    expect(host.textContent).toContain('規則 0 · AI 1')
  })
})
