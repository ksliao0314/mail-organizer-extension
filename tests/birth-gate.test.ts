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
import { getSubjectMemory, recordSubjectFilings, setRules, setSettings } from '@/shared/storage'
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

  it('短主旨（<8 字）不再學成 full_subject 規則（2026-11-05）', async () => {
    // 原本（覆核 F1）2-7 字主旨因 SubjectMemory 永不記錄而回退為首見即鑄。
    // 2026-11-05 起 extractSubjectSignal 的門檻與 SubjectMemory 一致（8）：
    // 這類短主旨本身就太泛（子字串誤中同寄件人的無關信、又以 compound 身分
    // 蓋過 thread memory），不是閘門要等的訊號 — 交給 AI / thread memory。
    const item = planItem({ emailSubject: '委任狀用印' }) // 5 字
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(0)
    expect(await listRules()).toHaveLength(0)
  })

  it('外部網域（P4 → domain 規則，非 full_subject）不受閘門影響', async () => {
    const item = planItem({ emailFrom: 'contact@company-a.example' })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules[0]!.type).toBe('domain')
  })
})

// ---- 2026-11-05 (item 7): the gate reads memory BEFORE this batch's filing,
// and a thread-routed subject repeat counts as the second observation.
//
// In the real flow the second same-subject email is routed by SUBJECT memory
// (source 'thread'), and thread-routed non-structural rows never learned — so
// passive full_subject rules could never be minted. And because execute
// recorded the filing first, the `harvested` exclusion was dead: the real
// filing had already cleared the flag. startExecute now snapshots subject
// memory before recording and passes it as `priorSubjectMemory`.
describe('full_subject 出生閘門 — 先前快照 + thread 重複 (2026-11-05)', () => {
  /** Mirror execute: snapshot (a copy, as startExecute takes), then record this batch's filing. */
  async function executeBatchFiling() {
    const prior = { ...(await getSubjectMemory()) }
    await fileSubject(1)
    return prior
  }

  it('首見（快照中無紀錄）→ 不鑄', async () => {
    const prior = await executeBatchFiling()
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)], { priorSubjectMemory: prior })
    expect(r.added).toBe(0)
  })

  it('第二封同主旨信由 thread memory（subject）路由、使用者未改 → 鑄', async () => {
    await fileSubject(1) // first email, filed by an earlier batch
    const prior = await executeBatchFiling()
    const item = planItem({
      source: 'thread',
      threadMatch: { kind: 'subject', previousFolderPath: FOLDER },
      targetFolderId: 'FID_' + FOLDER,
    })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)], { priorSubjectMemory: prior })
    expect(r.added).toBe(1) // was 0: thread rows never reached the gate
    expect((await listRules())[0]!.type).toBe('compound')
  })

  it('同對話的回覆（convId 路由）不算第二次觀察 → 不鑄', async () => {
    await fileSubject(1)
    const prior = await executeBatchFiling()
    const item = planItem({
      source: 'thread',
      threadMatch: { kind: 'convId', previousFolderPath: FOLDER },
    })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)], { priorSubjectMemory: prior })
    expect(r.added).toBe(0)
  })

  it('使用者改過目標的 thread 列不算確認 → 不鑄', async () => {
    await fileSubject(1)
    const prior = await executeBatchFiling()
    const item = planItem({
      source: 'thread',
      threadMatch: { kind: 'subject', previousFolderPath: FOLDER },
      userTouched: true,
    })
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)], { priorSubjectMemory: prior })
    expect(r.added).toBe(0)
  })

  it('harvested-only 的先前紀錄在快照下確實不算（舊流程中該檢查已失效）', async () => {
    await fileSubject(2, { harvested: true })
    const prior = await executeBatchFiling() // the real filing clears the flag AFTER the snapshot
    expect((await getSubjectMemory())[normalizeSubject(SUBJECT)]!.harvested).toBeUndefined()
    const item = planItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)], { priorSubjectMemory: prior })
    expect(r.added).toBe(0)
  })

  it('同批兩個不同對話的同主旨信 → 視為跨對話重複，鑄一條', async () => {
    const prior = await getSubjectMemory()
    const a = planItem({ conversationId: 'conv-a' })
    const b = planItem({ conversationId: 'conv-b' })
    const r = await generateAiConfirmedRules([a, b], [moved(a.emailId), moved(b.emailId)], {
      priorSubjectMemory: prior,
    })
    expect(r.added).toBe(1)
  })

  it('同批同一對話的兩封 → 仍是單一 thread，不鑄', async () => {
    const prior = await getSubjectMemory()
    const a = planItem({ conversationId: 'conv-a' })
    const b = planItem({ conversationId: 'conv-a' })
    const r = await generateAiConfirmedRules([a, b], [moved(a.emailId), moved(b.emailId)], {
      priorSubjectMemory: prior,
    })
    expect(r.added).toBe(0)
  })
})
