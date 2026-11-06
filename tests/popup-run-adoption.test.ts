// Taking a classify run over as this surface's plan tells the SW first
// (2026-11-06 review round 3, item 1).
//
// A DoneScreen's 回主畫面 cancels its own prefetch run with `unlessAdopted`
// (round 2, U4) — so every path that turns a run into THIS surface's plan has
// to mark it adopted, and before the plan is shown:
//   - the reopen resume never did: panel A sits on its DoneScreen while its
//     prefetch run R classifies; the toolbar popup B goes home, reopens and
//     resumes R as its plan; A's 回主畫面 then cancelled R and B showed
//     「AI 分類已中止」;
//   - 繼續歸檔 sent the adopt fire-and-forget, so a cancel landing between its
//     progress read and the adopt still won.

import type { Root } from 'react-dom/client'
import { act } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { PlanItem } from '@/shared/types'
import { buttonContaining, cleanup, mountApp, stubRuntime, type Msg, type Reply } from './helpers/ui-harness'

function item(id: string, subject: string): PlanItem {
  return {
    emailId: id,
    emailSubject: subject,
    emailFrom: `${id}@client.example`,
    action: 'skip',
    confidence: 0.9,
    reason: '',
    source: 'ai',
  }
}

/** Run R: a DoneScreen prefetch, still classifying, with one item ready. */
function runR(over: Record<string, unknown> = {}) {
  return {
    runId: 'run-r',
    totalEmails: 3,
    completedEmails: 1,
    chunks: 1,
    completedChunks: 0,
    ruleCount: 1,
    aiCount: 0,
    aiUnresolvedCount: 0,
    aiUnclassifiedCount: 0,
    aiCursor: 0,
    startedAt: Date.now() - 5000,
    done: false,
    rulePlan: [item('r0', 'prefetched mail')],
    aiPlan: [],
    folderTree: [],
    aiError: null,
    ...over,
  }
}

const STATUS = {
  owaConnected: true,
  tokenValid: true,
  apiKeyConfigured: true,
  apiKeyPreview: 'sk-…',
  model: 'claude-sonnet-5-5',
  folderCacheAgeMs: 1000,
  folderCacheCount: 1,
  rulesCount: 0,
  excludePrefixes: [],
  batchSize: 50,
  prefetchNextBatch: false,
  internalDomains: [],
  primaryRootPath: '',
}

const doneExecute = {
  inProgress: false,
  cancelRequested: false,
  startedAt: Date.now() - 60_000,
  finishedAt: Date.now() - 55_000,
  total: 1,
  current: 1,
  results: [{ emailId: 'm0', subject: 'm0', action: 'move', status: 'moved' }],
  summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
  rulesAdded: 0,
}

function sw(extra: (m: Msg) => Reply | undefined | Promise<Reply | undefined>) {
  return stubRuntime(async (m) => {
    const custom = await extra(m)
    if (custom) return custom
    switch (m.type) {
      case 'getStatus':
        return { ok: true, data: STATUS }
      case 'getOnboardingState':
        return { ok: true, data: { needed: false } }
      case 'getFolderActivity':
        return { ok: true, data: { activity: [] } }
      case 'getWeeklyDigest':
        return { ok: true, data: { shouldShow: false } }
      case 'getExecuteState':
        return { ok: true, data: { state: null } }
      case 'getPopupState':
        return { ok: true, data: null }
      case 'getUndoSnapshot':
        return { ok: true, data: { snapshot: null } }
      case 'peekNextBatch':
        return { ok: true, data: { eligibleCount: 3, cappedAtBatchSize: false } }
      case 'savePopupState':
      case 'clearPopupState':
      case 'clearExecuteState':
      case 'dismissUndo':
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

describe('reopen resume adopts the run it shows', () => {
  it('a prefetch run resumed as the plan is adopted (so its DoneScreen’s 回主畫面 spares it)', async () => {
    const s = sw((m) => {
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: runR() }
      if (m.type === 'adoptClassifyRun') return { ok: true, data: { adopted: true } }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('prefetched mail'))
    expect(s.last('adoptClassifyRun')).toEqual({ type: 'adoptClassifyRun', runId: 'run-r' })
    expect(host.textContent).toContain('AI 分類中') // still polling the live run
  })

  it('also when a saved plan of the same run is restored', async () => {
    const s = sw((m) => {
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: runR() }
      if (m.type === 'adoptClassifyRun') return { ok: true, data: { adopted: true } }
      if (m.type === 'getPopupState') {
        return {
          ok: true,
          data: {
            phase: 'plan',
            updatedAt: Date.now(),
            plan: {
              items: [item('r0', 'prefetched mail')],
              folderTree: [],
              usage: null,
              summary: { ruleHits: 1, aiHandled: 0 },
              banner: null,
              aiPendingFlag: true,
              runId: 'run-r',
            },
          },
        }
      }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('prefetched mail'))
    expect(s.last('adoptClassifyRun')).toEqual({ type: 'adoptClassifyRun', runId: 'run-r' })
  })

  it('a run cancelled since it was read is not resumed', async () => {
    sw((m) => {
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: runR() }
      if (m.type === 'adoptClassifyRun') return { ok: true, data: { adopted: false } }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('開始歸類'))
    expect(host.textContent).not.toContain('prefetched mail')
  })
})

describe('繼續歸檔 awaits the adopt before switching screens', () => {
  it('nothing moves on (no batch cleared, no plan shown) until the SW has recorded the adoption', async () => {
    let answerAdopt: ((r: Reply) => void) | null = null
    const s = sw((m) => {
      if (m.type === 'getExecuteState') return { ok: true, data: { state: doneExecute } }
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: runR() }
      if (m.type === 'adoptClassifyRun') {
        return new Promise<Reply>((resolve) => {
          answerAdopt = resolve
        })
      }
      return undefined
    })
    const host = await renderPopup()
    const go = await vi.waitFor(() => {
      const b = buttonContaining(host, '繼續歸檔下一批')
      expect(b).toBeDefined()
      return b!
    })
    await act(async () => go.click())
    await vi.waitFor(() => expect(answerAdopt).not.toBeNull())
    await new Promise((r) => setTimeout(r, 20))
    expect(s.count('clearExecuteState')).toBe(0)
    expect(host.textContent).not.toContain('prefetched mail')

    await act(async () => answerAdopt!({ ok: true, data: { adopted: true } }))
    await vi.waitFor(() => expect(host.textContent).toContain('prefetched mail'))
    expect(s.count('clearExecuteState')).toBe(1)
    expect(s.count('classifyPreflight')).toBe(0)
  })

  it('a run cancelled before the adopt landed is not shown — the next batch is classified afresh', async () => {
    const s = sw((m) => {
      if (m.type === 'getExecuteState') return { ok: true, data: { state: doneExecute } }
      if (m.type === 'getAiClassifyProgress') return { ok: true, data: runR() }
      if (m.type === 'adoptClassifyRun') return { ok: true, data: { adopted: false } }
      if (m.type === 'classifyPreflight') return { ok: false, code: 'TEST_STOP', message: 'stop here' }
      return undefined
    })
    const host = await renderPopup()
    const go = await vi.waitFor(() => {
      const b = buttonContaining(host, '繼續歸檔下一批')
      expect(b).toBeDefined()
      return b!
    })
    await act(async () => go.click())
    await vi.waitFor(() => expect(s.count('classifyPreflight')).toBe(1))
    expect(host.textContent).not.toContain('prefetched mail')
  })
})
