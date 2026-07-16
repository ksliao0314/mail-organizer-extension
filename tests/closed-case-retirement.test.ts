// 刀3 (2026-07-16): 兩階段結案退場。
// Phase 1 (applyClosedCaseRetirement): 路徑進歸檔樹 → 軟停用
// 'case_closed'（同步傳播）；搬出 → 對稱復職。資料夾位置是唯一真相 —
// 不靠信件觸發（alreadyCovered 會擋 re-mint，mail-driven re-enable
// 永遠不會發生）。
// Phase 2 (purgeExpiredClosedCaseRules): 停用滿 30 天 → 硬刪不留
// tombstone（sweep 逐機收斂; user_manual 豁免）。

import { beforeEach, describe, expect, it } from 'vitest'
import {
  applyClosedCaseRetirement,
  listRules,
  newRule,
  purgeExpiredClosedCaseRules,
} from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { Rule } from '@/shared/types'

const ARCHIVE = '05已完成案件'
const NOW = Date.parse('2026-07-16T00:00:00.000Z')
const daysAgo = (d: number) => new Date(NOW - d * 86_400_000).toISOString()

beforeEach(async () => {
  await chrome.storage.local.clear()
})

function rule(over: Partial<Rule>): Rule {
  return newRule({
    type: 'compound',
    signal: '{"conditions":[{"type":"subject_keyword","value":"某案主旨"}]}',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03進行中案件/某案',
    confidence: 0.85,
    source: 'ai_confirmed',
    ...over,
  })
}

describe('Phase 1 — applyClosedCaseRetirement 退場', () => {
  it('路徑在歸檔樹下的 enabled 規則 → 停用 + case_closed + 時間戳', async () => {
    await setRules([rule({ targetFolderPath: `${ARCHIVE}/某案/往來書狀` })])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.retired).toBe(1)
    const [after] = await listRules()
    expect(after!.enabled).toBe(false)
    expect(after!.autoDisabledReason).toBe('case_closed')
    expect(after!.autoDisabledAt).toBe(new Date(NOW).toISOString())
  })

  it('路徑等於 prefix 本身也算在樹下', async () => {
    await setRules([rule({ targetFolderPath: ARCHIVE })])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.retired).toBe(1)
  })

  it('segment 精確比對: 「05已完成案件X」不被 prefix「05已完成案件」誤傷', async () => {
    await setRules([rule({ targetFolderPath: '05已完成案件X/還在辦' })])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.retired).toBe(0)
    expect((await listRules())[0]!.enabled).toBe(true)
  })

  it('user_manual 也退場（資料夾位置是較新的意思表示）', async () => {
    await setRules([
      rule({ source: 'user_manual', targetFolderPath: `${ARCHIVE}/某案` }),
    ])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.retired).toBe(1)
  })

  it('orphaned 規則不動（另一條管線的事）', async () => {
    const orphan = rule({ targetFolderPath: `${ARCHIVE}/某案` })
    orphan.orphaned = true
    await setRules([orphan])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.retired).toBe(0)
  })

  it('空 prefix 清單 → 不退任何規則', async () => {
    await setRules([rule({ targetFolderPath: `${ARCHIVE}/某案` })])
    const r = await applyClosedCaseRetirement([], { now: NOW })
    expect(r.retired).toBe(0)
  })

  it('回傳更新後的規則陣列（preflight 要用它重建 match index）', async () => {
    await setRules([rule({ targetFolderPath: `${ARCHIVE}/某案` })])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.rules).toHaveLength(1)
    expect(r.rules[0]!.enabled).toBe(false)
  })
})

describe('Phase 1 — 對稱復職', () => {
  it('case_closed 規則的路徑已不在歸檔樹（資料夾搬出）→ 復職、清 reason', async () => {
    const retired = rule({ targetFolderPath: '03進行中案件/重啟的案子' })
    retired.enabled = false
    retired.autoDisabledAt = daysAgo(10)
    retired.autoDisabledReason = 'case_closed'
    await setRules([retired])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.revived).toBe(1)
    const [after] = await listRules()
    expect(after!.enabled).toBe(true)
    expect(after!.autoDisabledReason).toBeUndefined()
    expect(after!.autoDisabledAt).toBeUndefined()
  })

  it('取消勾選歸檔 prefix（清空清單）→ 全部 case_closed 規則復職', async () => {
    const retired = rule({ targetFolderPath: `${ARCHIVE}/某案` })
    retired.enabled = false
    retired.autoDisabledReason = 'case_closed'
    retired.autoDisabledAt = daysAgo(5)
    await setRules([retired])
    const r = await applyClosedCaseRetirement([], { now: NOW })
    expect(r.revived).toBe(1)
  })

  it('其他原因停用的規則（stale / 使用者手動）不被復職', async () => {
    const userDisabled = rule({ targetFolderPath: '03進行中案件/某案' })
    userDisabled.enabled = false // 使用者手動停用，無 reason
    const staleDisabled = rule({ targetFolderPath: '03進行中案件/另案' })
    staleDisabled.enabled = false
    staleDisabled.autoDisabledReason = 'high-error-rate'
    staleDisabled.autoDisabledAt = daysAgo(5)
    await setRules([userDisabled, staleDisabled])
    const r = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(r.revived).toBe(0)
    const after = await listRules()
    expect(after.every((x) => !x.enabled)).toBe(true)
  })

  it('退場↔復職往返是冪等的（第二次呼叫零變動）', async () => {
    await setRules([rule({ targetFolderPath: `${ARCHIVE}/某案` })])
    await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    const second = await applyClosedCaseRetirement([ARCHIVE], { now: NOW })
    expect(second.retired).toBe(0)
    expect(second.revived).toBe(0)
  })
})

describe('Phase 2 — purgeExpiredClosedCaseRules 硬刪', () => {
  function retiredRule(daysDisabled: number, over: Partial<Rule> = {}): Rule {
    const r = rule({ targetFolderPath: `${ARCHIVE}/某案`, ...over })
    r.enabled = false
    r.autoDisabledReason = 'case_closed'
    r.autoDisabledAt = daysAgo(daysDisabled)
    return r
  }

  it('停用 31 天 → 刪除; 29 天 → 保留', async () => {
    await setRules([retiredRule(31), retiredRule(29)])
    const { purged } = await purgeExpiredClosedCaseRules({ now: NOW })
    expect(purged).toBe(1)
    expect(await listRules()).toHaveLength(1)
  })

  it('user_manual 永不硬刪（停用 100 天仍保留）', async () => {
    await setRules([retiredRule(100, { source: 'user_manual' })])
    const { purged } = await purgeExpiredClosedCaseRules({ now: NOW })
    expect(purged).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })

  it('其他 reason 的停用規則不受 purge 影響', async () => {
    const other = rule({})
    other.enabled = false
    other.autoDisabledReason = 'stale'
    other.autoDisabledAt = daysAgo(100)
    await setRules([other])
    const { purged } = await purgeExpiredClosedCaseRules({ now: NOW })
    expect(purged).toBe(0)
  })

  it('autoDisabledAt 缺失/壞值 → 防禦性保留', async () => {
    const broken = retiredRule(31)
    broken.autoDisabledAt = undefined
    await setRules([broken])
    const { purged } = await purgeExpiredClosedCaseRules({ now: NOW })
    expect(purged).toBe(0)
  })
})
