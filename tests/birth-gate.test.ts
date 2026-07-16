// 刀1 (2026-07-16): full_subject 兩次觀察出生閘門。
// 實測 73% 的規則庫是 compound/full_subject、其中 71% 從未命中 — 被動
// AI 確認的 full_subject 訊號改為「SubjectMemory 中已有一次先前的真實
// 歸檔到同一資料夾」才鑄。此檔定義該 contract 的全部邊界。
//
// 時序前提: execute 管線在呼叫 generateAiConfirmedRules 之前已跑
// recordSubjectFilings — 所以「首見」在 memory 中是 timesFiled=1（本次），
// 「第二次觀察」是 timesFiled=2（一次先前 + 本次）。測試直接以
// recordSubjectFilings 的呼叫次數模擬。

import { beforeEach, describe, expect, it } from 'vitest'
import { generateAiConfirmedRules } from '@/background/execute'
import { listRules } from '@/shared/rules'
import { recordSubjectFilings, setRules, setSettings } from '@/shared/storage'
import { normalizeSubject } from '@/shared/normalize'
import type { PlanItem } from '@/shared/types'

const SUBJECT = '請審閱三井租賃合約修訂版'
const FOLDER = '01所內資料/合約審閱'

beforeEach(async () => {
  await chrome.storage.local.clear()
  await setRules([])
  // 內部網域 → chooseLearningSignal 走 P6（sender + full_subject compound），
  // 這正是實測中最大的鑄造來源。
  await setSettings({ internalDomains: ['firm.example'] })
})

function planItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: SUBJECT,
    emailFrom: 'colleague@firm.example',
    action: 'move',
    targetFolderPath: FOLDER,
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}
const moved = (emailId: string) => ({
  emailId,
  subject: 'X',
  action: 'move' as const,
  status: 'moved' as const,
})

async function fileSubject(times: number, opts: { folder?: string; harvested?: boolean } = {}) {
  for (let i = 0; i < times; i++) {
    await recordSubjectFilings(
      [
        {
          normalizedSubject: normalizeSubject(SUBJECT),
          folderId: 'FID_' + (opts.folder ?? FOLDER),
          folderPath: opts.folder ?? FOLDER,
        },
      ],
      { harvested: opts.harvested === true },
    )
  }
}

describe('full_subject 出生閘門 (兩次觀察)', () => {
  it('首見 (timesFiled=1) 的被動 AI 確認 → 不鑄', async () => {
    await fileSubject(1) // 只有本次
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
    expect(await listRules()).toHaveLength(0)
  })

  it('完全沒有 subject memory（記錄失敗的防禦情境）→ 不鑄', async () => {
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
  })

  it('第二次觀察 (timesFiled=2、同資料夾、無衝突) → 鑄 compound', async () => {
    await fileSubject(2)
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('compound')
  })

  it('主旨曾在資料夾間搖擺 (conflictCount>0) → 不鑄', async () => {
    await fileSubject(1, { folder: '01所內資料/別的地方' })
    await fileSubject(2) // 換回目標資料夾 → conflictCount 1
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
  })

  it('harvested-only 的先前紀錄不算觀察 → 不鑄', async () => {
    // 兩次都是 harvest（folder-content 弱證據）→ entry.harvested=true
    await fileSubject(2, { harvested: true })
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
  })

  it('使用者主動解決 (source=unresolved) → 首見即鑄（明確人為意圖）', async () => {
    const item = planItem({ source: 'unresolved' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
  })

  it('結構化案號不受閘門影響 → 首見即鑄', async () => {
    // 主旨含案號 → P1 compound/court_case，非 full_subject。
    const item = planItem({ emailSubject: '113年度訴字第204號 開庭通知' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(['compound', 'subject_keyword', 'case_code']).toContain(rules[0]!.type)
  })

  it('短主旨（<8 字、SubjectMemory 永不記錄）回退為首見即鑄，不被無限封鎖', async () => {
    // 覆核 F1: SubjectMemory 寫入端以 MIN_NORMALIZED_SUBJECT_LEN=8 為門檻，
    // 2-7 字主旨永遠不會有 entry — 對它們套閘門是無限封鎖而非兩次觀察。
    const item = planItem({ emailSubject: '委任狀用印' }) // 5 字
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
  })

  it('外部網域（P4 → domain 規則，非 full_subject）不受閘門影響', async () => {
    const item = planItem({ emailFrom: 'contact@company-a.example' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('domain')
  })
})
