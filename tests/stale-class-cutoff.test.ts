// 刀2 (2026-07-16): full_subject 死規則的 45 天類別化回收。
// matchCount===0 的 full_subject 類（subject_keyword / compound 且主旨
// 條件非案號）在 45 天即走 stale 硬刪；結構化案號 / domain / sender
// 維持 100 天。謂詞只依賴同步欄位 → 每台機器 sweep 各自收斂。

import { beforeEach, describe, expect, it } from 'vitest'
import { autoDisableStaleRules, encodeCompound, isFullSubjectRule, newRule } from '@/shared/rules'
import { listRules } from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { Rule } from '@/shared/types'

const NOW = Date.parse('2026-07-16T00:00:00.000Z')
const daysAgo = (d: number) => new Date(NOW - d * 86_400_000).toISOString()

function rule(over: Partial<Rule>): Rule {
  return newRule({
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03進行中案件/某案',
    confidence: 0.8,
    source: 'ai_confirmed',
    ...over,
  })
}

const fullSubjectCompound = (over: Partial<Rule> = {}) =>
  rule({
    type: 'compound',
    signal: encodeCompound([
      { type: 'sender', value: 'colleague@firm.example' },
      { type: 'subject_keyword', value: '請審閱三井租賃合約修訂版' },
    ]),
    ...over,
  })

const courtCaseCompound = (over: Partial<Rule> = {}) =>
  rule({
    type: 'compound',
    signal: encodeCompound([
      { type: 'domain', value: 'court.example.tw' },
      { type: 'subject_keyword', value: '113年度訴字第204號' },
    ]),
    ...over,
  })

describe('isFullSubjectRule 分類', () => {
  it('compound 主旨條件非案號 → true; 案號 → false', () => {
    expect(isFullSubjectRule(fullSubjectCompound())).toBe(true)
    expect(isFullSubjectRule(courtCaseCompound())).toBe(false)
  })
  it('subject_keyword 全主旨 → true; 案號 → false; 其他 type → false', () => {
    expect(isFullSubjectRule(rule({ type: 'subject_keyword', signal: '合約審閱通知' }))).toBe(true)
    expect(isFullSubjectRule(rule({ type: 'subject_keyword', signal: '113年度訴字第204號' }))).toBe(false)
    expect(isFullSubjectRule(rule({ type: 'domain', signal: 'x.example' }))).toBe(false)
    expect(isFullSubjectRule(rule({ type: 'sender', signal: 'a@x.example' }))).toBe(false)
  })
})

describe('never-matched full_subject 45 天回收', () => {
  it('full_subject compound、未命中、50 天 → 硬刪', async () => {
    const r = fullSubjectCompound()
    r.createdAt = daysAgo(50)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(1)
    expect(await listRules()).toHaveLength(0)
  })

  it('full_subject compound、未命中、40 天 → 保留（未達 45）', async () => {
    const r = fullSubjectCompound()
    r.createdAt = daysAgo(40)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })

  it('案號 compound、未命中、50 天 → 保留（結構化走 100 天）', async () => {
    const r = courtCaseCompound()
    r.createdAt = daysAgo(50)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })

  it('domain、未命中、50 天 → 保留（非 full_subject 類）', async () => {
    const r = rule({})
    r.createdAt = daysAgo(50)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(0)
  })

  it('full_subject 但命中過 → 不走 45 天（last-used 100 天標準）', async () => {
    const r = fullSubjectCompound()
    r.createdAt = daysAgo(50)
    r.matchCount = 3
    r.lastUsedAt = daysAgo(50)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })

  it('user_manual full_subject → 神聖不可回收', async () => {
    const r = fullSubjectCompound({ source: 'user_manual' })
    r.createdAt = daysAgo(200)
    await setRules([r])
    const report = await autoDisableStaleRules({ now: NOW })
    expect(report.byReason.stale).toBe(0)
    expect(await listRules()).toHaveLength(1)
  })
})
