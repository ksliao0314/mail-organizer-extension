// Refusals in the popup (2026-11-06, round-2 review K5 / K7).
//
//   - a refused row explained 「AI 信心過低,沒給出明確分類」 like a low-confidence
//     answer, and its 待確認 tooltip said 「AI 信心低於門檻…」;
//   - the non-fatal refusal banner was prefixed 「AI 分類失敗：」 and titled
//     like a failure;
//   - categories showed as raw API ids (shared label helper, see
//     refusal-isolation.test.ts);
//   - mail left out of the batch because Claude refused it before is counted
//     next to flagged mail when a batch starts.

import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { RuleTraceBlock } from '@/popup/components/PlanRow'
import type { MailFolderNode, PlanItem } from '@/shared/types'
import { buttonContaining, cleanup, click, mountApp, stubRuntime, type Msg, type Reply } from './helpers/ui-harness'

const tree: MailFolderNode[] = [
  { id: 'ROOT', displayName: 'Cases', path: 'Cases', children: [] },
]

function row(over: Partial<PlanItem>): PlanItem {
  return {
    emailId: 'e1',
    emailSubject: 'subject',
    emailFrom: 'a@acme.example',
    action: 'skip',
    confidence: 0,
    reason: '',
    source: 'unresolved',
    ...over,
  }
}

const refusedEmail = row({
  emailId: 'r1',
  emailSubject: '被拒的那封',
  unclassified: true,
  reason: 'Claude 拒絕分類「被拒的那封」（類別：資訊安全），請手動處理',
  refusal: { scope: 'email', category: 'cyber' },
})
const refusedBatch = row({
  emailId: 'r2',
  emailSubject: '整批被拒',
  unclassified: true,
  reason: 'Claude 拒絕了這封所在的整批分類請求（類別：未提供），無法確定是哪一封；下一批會再試',
  refusal: { scope: 'batch', category: null },
})
const failed = row({ emailId: 'f1', emailSubject: '呼叫失敗', unclassified: true, reason: 'AI 分類失敗：CLASSIFIER_HTTP_529' })
const gated = row({ emailId: 'g1', emailSubject: '信心不足', reason: '信心 0.30 < 0.5 (自動保留)' })

function status() {
  return {
    owaConnected: true,
    tokenValid: true,
    apiKeyConfigured: true,
    apiKeyPreview: 'sk-…',
    model: 'claude-sonnet-5-5',
    folderCacheAgeMs: null,
    folderCacheCount: 0,
    rulesCount: 0,
    excludePrefixes: [],
    batchSize: 50,
    internalDomains: [],
    primaryRootPath: '',
  }
}

function popupSw(extra: (m: Msg) => Reply | undefined) {
  return stubRuntime(async (m) => {
    const custom = extra(m)
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
      case 'getPopupState':
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

/** A finished run the popup resumes from (full view), with this banner and these AI rows. */
function finishedRun(aiPlan: PlanItem[], aiError: { code: string; message: string } | null): Reply {
  return {
    ok: true,
    data: {
      runId: 'run-a',
      totalEmails: aiPlan.length,
      completedEmails: aiPlan.length,
      chunks: 1,
      completedChunks: 1,
      ruleCount: 0,
      aiCount: aiPlan.length,
      aiUnresolvedCount: aiPlan.length,
      aiUnclassifiedCount: aiPlan.filter((i) => i.unclassified).length,
      aiCursor: aiPlan.length,
      startedAt: 1,
      done: true,
      rulePlan: [],
      aiPlan,
      folderTree: tree,
      aiError,
    },
  }
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

/** The 待確認 badge's tooltip on the plan row showing `subject`. */
function attentionTooltip(host: HTMLElement, subject: string): string | null {
  const rowEl = Array.from(host.querySelectorAll<HTMLElement>('.group.relative')).find((x) =>
    x.textContent?.includes(subject),
  )
  const badge = Array.from(rowEl?.querySelectorAll<HTMLElement>('[title]') ?? []).find(
    (el) => el.textContent?.trim() === '待確認',
  )
  return badge?.getAttribute('title') ?? null
}

describe('refusal banner', () => {
  it('has its own title and no 「AI 分類失敗：」 prefix — the rest of the batch was classified', async () => {
    const message = 'Claude 拒絕分類 1 封郵件：「被拒的那封」（類別：資訊安全），已標為未分類，請手動處理'
    popupSw((m) =>
      m.type === 'getAiClassifyProgress'
        ? finishedRun([refusedEmail, gated], { code: 'CLASSIFIER_REFUSAL', message })
        : undefined,
    )
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('被拒的那封'))
    const text = host.textContent ?? ''
    expect(text).toContain('Claude 拒絕分類部分郵件、需手動處理')
    expect(text).toContain(message)
    expect(text).not.toContain('AI 分類失敗：Claude 拒絕')
    expect(text).not.toContain('AI 分類發生問題')
  })

  it('a real failure keeps the failure wording', async () => {
    popupSw((m) =>
      m.type === 'getAiClassifyProgress'
        ? finishedRun([failed, gated], { code: 'CLASSIFIER_HTTP_529', message: 'Claude 服務暫時過載（529）' })
        : undefined,
    )
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('呼叫失敗'))
    expect(host.textContent).toContain('AI 分類失敗：Claude 服務暫時過載（529）')
  })
})

describe('unresolved rows say why they have no decision', () => {
  it('in the collapsed 待確認 tooltip', async () => {
    popupSw((m) =>
      m.type === 'getAiClassifyProgress'
        ? finishedRun([refusedEmail, refusedBatch, failed, gated], null)
        : undefined,
    )
    const host = await renderPopup()
    await vi.waitFor(() => expect(host.textContent).toContain('信心不足'))
    expect(attentionTooltip(host, '被拒的那封')).toBe('Claude 拒絕分類這封郵件，請你手動處理')
    expect(attentionTooltip(host, '整批被拒')).toBe('Claude 拒絕了這封所在的整批請求，請你決定')
    expect(attentionTooltip(host, '呼叫失敗')).toBe('AI 沒有完成這封的分類，請你決定')
    expect(attentionTooltip(host, '信心不足')).toBe('AI 信心低於門檻、未產生明確分類，請你決定')
  })

  it('in the expanded trace (refused rows never claim low confidence)', async () => {
    const render = async (item: PlanItem) => {
      const host = document.createElement('div')
      document.body.appendChild(host)
      const r = createRoot(host)
      await act(async () => r.render(createElement(RuleTraceBlock, { item, sourceDetailBadge: null })))
      const text = host.textContent ?? ''
      act(() => r.unmount())
      host.remove()
      return text
    }
    const email = await render(refusedEmail)
    expect(email).toContain('Claude 拒絕分類這封郵件（類別：資訊安全），請你手動決定歸檔位置')
    expect(email).not.toContain('信心過低')
    expect(email).not.toContain('cyber')
    const batch = await render(refusedBatch)
    expect(batch).toContain('Claude 拒絕了這封所在的整批請求、無法確定是哪一封；下一批會再試')
    expect(batch).not.toContain('信心過低')
    const fail = await render(failed)
    expect(fail).toContain('AI 沒有完成這封的分類，下一批會重新分類')
    expect(fail).not.toContain('信心過低')
    expect(await render(gated)).toContain('AI 信心過低,沒給出明確分類,請你逐封決定')
  })
})

// The DoneScreen of a finished execute, for 繼續歸檔下一批.
const doneExecute = {
  inProgress: false,
  cancelRequested: false,
  startedAt: Date.UTC(2026, 10, 5, 9, 0, 0),
  finishedAt: Date.UTC(2026, 10, 5, 9, 0, 4),
  total: 1,
  current: 1,
  results: [{ emailId: 'm0', subject: '規則命中', action: 'move', status: 'moved' }],
  summary: { moved: 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
  rulesAdded: 0,
}

describe('繼續歸檔下一批 adopting a run (e.g. a DoneScreen prefetch)', () => {
  // Round-3 review item 2: the adoption only looked at the run's aiError, so
  // what its preflight left out — refused mail included — was never shown.
  it('shows what the run’s preflight left out, like a batch started here', async () => {
    const run = finishedRun([gated], null)
    ;(run.data as Record<string, unknown>).excluded = { kept: 0, flagged: 2, refused: 1 }
    popupSw((m) => {
      if (m.type === 'getExecuteState') return { ok: true, data: { state: doneExecute } }
      if (m.type === 'getAiClassifyProgress') return run
      if (['adoptClassifyRun', 'clearExecuteState', 'dismissUndo'].includes(m.type)) return { ok: true }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(buttonContaining(host, '繼續歸檔下一批')).toBeTruthy())
    click(buttonContaining(host, '繼續歸檔下一批'))
    await vi.waitFor(() => expect(host.textContent).toContain('信心不足'))
    expect(host.textContent).toContain('部分郵件已預先排除')
    expect(host.textContent).toContain('2 件標記待處理 + 1 件 Claude 先前拒絕分類（需手動歸檔）')
  })

  // Round-3 review item 3: a run whose items could not be stored keeps its
  // counts for the summary — but has nothing to adopt.
  it('does not adopt a run stored without its items: a fresh batch starts instead', async () => {
    const run = finishedRun([], { code: 'STORAGE_FAILED', message: '瀏覽器暫存空間不足' })
    Object.assign(run.data as Record<string, unknown>, {
      ruleCount: 1,
      aiCount: 2,
      aiUnresolvedCount: 0,
      aiUnclassifiedCount: 0,
      planDropped: true,
    })
    const sw = popupSw((m) => {
      if (m.type === 'getExecuteState') return { ok: true, data: { state: doneExecute } }
      if (m.type === 'getAiClassifyProgress') return run
      if (['clearExecuteState', 'dismissUndo', 'cancelClassify'].includes(m.type)) return { ok: true }
      if (m.type === 'classifyPreflight') return { ok: false, code: 'TEST_STOP', message: 'stop here' }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(buttonContaining(host, '繼續歸檔下一批')).toBeTruthy())
    click(buttonContaining(host, '繼續歸檔下一批'))
    await vi.waitFor(() => expect(sw.count('classifyPreflight')).toBe(1))
    expect(sw.count('adoptClassifyRun')).toBe(0)
    expect(sw.last('cancelClassify')).toEqual({ type: 'cancelClassify', runId: 'run-a' })
  })
})

describe('starting a batch', () => {
  it('counts mail left out because Claude refused it before, next to flagged mail (K5)', async () => {
    popupSw((m) => {
      if (m.type === 'classifyPreflight') {
        return {
          ok: true,
          data: {
            runId: 'run-b',
            rulePlan: [],
            unmatchedPreview: [],
            folderTree: tree,
            aiConfidenceThreshold: 0.5,
            preFilteredCount: 0,
            flaggedCount: 2,
            refusedCount: 3,
            requestedBatchSize: 50,
          },
        }
      }
      return undefined
    })
    const host = await renderPopup()
    await vi.waitFor(() => expect(buttonContaining(host, '開始歸類')).toBeTruthy())
    click(buttonContaining(host, '開始歸類'))
    await vi.waitFor(() => expect(host.textContent).toContain('部分郵件已預先排除'))
    expect(host.textContent).toContain('2 件標記待處理 + 3 件 Claude 先前拒絕分類（需手動歸檔）')
  })
})
