// Executing a batch must not hide never-classified mail (2026-11-05 audit).
//
// When the Claude call fails mid-run (429 / 529 / network) or the response
// omits an email, the SW fills the plan with action:'skip' placeholders. The
// popup tells the user 「規則命中與已完成部分仍可執行」; executing then wrote
// every action==='skip' row into skipHistory, so those emails vanished from
// every batch for 60 days, shown as 「先前選擇保留」. Placeholders are now
// flagged `unclassified` and only recorded when the user explicitly chose 保留.
// AI / confidence-gate skip decisions keep the old behaviour.

import { beforeEach, describe, expect, it, vi } from 'vitest'
import { startExecute } from '@/background/execute'
import type { OutlookApi } from '@/shared/outlook-api'
import { getActiveDays, getSkipHistory } from '@/shared/storage'
import type { PlanItem } from '@/shared/types'

beforeEach(() => {
  // keep-alive + undo use chrome.alarms; tests/setup.ts only mocks storage.
  ;(globalThis as unknown as { chrome: Record<string, unknown> }).chrome.alarms = {
    create: vi.fn(async () => {}),
    clear: vi.fn(async () => true),
    get: vi.fn(async () => undefined),
    onAlarm: { addListener: vi.fn() },
  }
})

function skipRow(over: Partial<PlanItem>): PlanItem {
  return {
    emailId: 'id',
    emailSubject: '主旨',
    emailFrom: 'someone@client.example',
    action: 'skip',
    confidence: 0,
    reason: '',
    source: 'unresolved',
    ...over,
  }
}

describe('skipHistory after execute', () => {
  it('records real keep decisions but not never-classified placeholders', async () => {
    const plan: PlanItem[] = [
      skipRow({ emailId: 'api-failed', unclassified: true, reason: 'AI 分類失敗：HTTP_529' }),
      skipRow({ emailId: 'ai-omitted', unclassified: true, reason: 'AI 未回傳此 index 的判斷' }),
      skipRow({ emailId: 'failed-but-user-kept', unclassified: true, userTouched: true }),
      skipRow({
        emailId: 'low-confidence',
        reason: '信心 0.42 < 0.5 (自動保留)',
        aiOriginalAction: 'move',
      }),
      skipRow({ emailId: 'ai-said-skip' }),
    ]
    // Skip rows never touch Outlook, so an empty api object suffices.
    await startExecute(plan, [], {} as OutlookApi)

    const kept = Object.keys(await getSkipHistory()).sort()
    expect(kept).toEqual(['ai-said-skip', 'failed-but-user-kept', 'low-confidence'])
  })

  it('an executed batch records today as an active day (stale-sweep clock)', async () => {
    await startExecute([skipRow({ emailId: 'x' })], [], {} as OutlookApi)
    expect(await getActiveDays()).toEqual([new Date().toISOString().slice(0, 10)])
  })
})
