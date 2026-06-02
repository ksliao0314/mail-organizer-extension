import { describe, expect, it } from 'vitest'
import { computeRuleHealth } from '@/shared/rule-health'
import type { Rule } from '@/shared/types'

const NOW = new Date('2026-05-19T00:00:00.000Z').getTime()
const DAY = 86_400_000

function rule(over: Partial<Rule>): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: new Date(NOW - 7 * DAY).toISOString(),
    source: 'user_manual',
    ...over,
  }
}

describe('computeRuleHealth — sleeping', () => {
  it('marks never-used rules older than threshold as sleeping', () => {
    const r = rule({ matchCount: 0, createdAt: new Date(NOW - 60 * DAY).toISOString() })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.sleeping).toBe(1)
  })

  it('does NOT mark fresh never-used rules as sleeping', () => {
    const r = rule({ matchCount: 0, createdAt: new Date(NOW - 5 * DAY).toISOString() })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.sleeping).toBe(0)
  })

  it('marks once-used-but-idle rules as sleeping', () => {
    const r = rule({
      matchCount: 5,
      lastUsedAt: new Date(NOW - 120 * DAY).toISOString(),
    })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.sleeping).toBe(1)
  })

  it('does NOT mark recently-used rules as sleeping', () => {
    const r = rule({
      matchCount: 5,
      lastUsedAt: new Date(NOW - 10 * DAY).toISOString(),
    })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.sleeping).toBe(0)
  })

  it('ignores disabled rules', () => {
    const r = rule({
      enabled: false,
      matchCount: 0,
      createdAt: new Date(NOW - 60 * DAY).toISOString(),
    })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.sleeping).toBe(0)
  })

  it('sorts sleeping by oldest createdAt first', () => {
    const old = rule({
      id: 'old',
      matchCount: 0,
      createdAt: new Date(NOW - 90 * DAY).toISOString(),
    })
    const newer = rule({
      id: 'newer',
      matchCount: 0,
      createdAt: new Date(NOW - 35 * DAY).toISOString(),
    })
    const h = computeRuleHealth([newer, old], { now: NOW })
    expect(h.sleeping.map((r) => r.id)).toEqual(['old', 'newer'])
  })
})

describe('computeRuleHealth — hotVague', () => {
  it('flags high matchCount rules targeting 未分類', () => {
    const r = rule({ matchCount: 50, targetFolderPath: '02未分類' })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.hotVague).toBe(1)
  })

  it('flags 其他 / 待釐清 / 雜項 variants', () => {
    const rules = [
      rule({ id: '1', matchCount: 12, targetFolderPath: '02/其他' }),
      rule({ id: '2', matchCount: 30, targetFolderPath: '01待釐清' }),
      rule({ id: '3', matchCount: 15, targetFolderPath: 'X/雜項' }),
    ]
    const h = computeRuleHealth(rules, { now: NOW })
    expect(h.counts.hotVague).toBe(3)
  })

  it('does NOT flag specific target paths even at high matchCount', () => {
    const r = rule({ matchCount: 100, targetFolderPath: '案件/客戶A/合約' })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.hotVague).toBe(0)
  })

  it('does NOT flag vague targets below threshold', () => {
    const r = rule({ matchCount: 3, targetFolderPath: '02未分類' })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.hotVague).toBe(0)
  })

  it('sorts hotVague by highest matchCount first', () => {
    const a = rule({ id: 'a', matchCount: 12, targetFolderPath: '02未分類' })
    const b = rule({ id: 'b', matchCount: 100, targetFolderPath: '02未分類' })
    const h = computeRuleHealth([a, b], { now: NOW })
    expect(h.hotVague.map((r) => r.id)).toEqual(['b', 'a'])
  })
})

describe('computeRuleHealth — orphaned', () => {
  it('surfaces orphaned rules separately and does not also flag them sleeping', () => {
    const r = rule({
      orphaned: true,
      matchCount: 0,
      createdAt: new Date(NOW - 365 * DAY).toISOString(),
    })
    const h = computeRuleHealth([r], { now: NOW })
    expect(h.counts.orphaned).toBe(1)
    expect(h.counts.sleeping).toBe(0)
  })
})

describe('computeRuleHealth — conflicts', () => {
  it('collects rule IDs from findConflicts', () => {
    // Two enabled domain rules same signal different targetFolderId → conflict
    const a = rule({
      id: 'a',
      signal: 'company-a.example',
      targetFolderId: 'AAA' + 'a'.repeat(50),
      targetFolderPath: '03/A',
    })
    const b = rule({
      id: 'b',
      signal: 'company-a.example',
      targetFolderId: 'BBB' + 'b'.repeat(50),
      targetFolderPath: '03/B',
    })
    const h = computeRuleHealth([a, b], { now: NOW })
    expect(h.counts.conflicts).toBeGreaterThan(0)
    expect(h.conflictRuleIds.has('a')).toBe(true)
    expect(h.conflictRuleIds.has('b')).toBe(true)
  })

  it('no conflicts when signals differ', () => {
    const a = rule({ id: 'a', signal: 'a.com' })
    const b = rule({ id: 'b', signal: 'b.com' })
    const h = computeRuleHealth([a, b], { now: NOW })
    expect(h.counts.conflicts).toBe(0)
  })
})

describe('computeRuleHealth — options', () => {
  it('respects custom thresholds', () => {
    const r = rule({ matchCount: 0, createdAt: new Date(NOW - 14 * DAY).toISOString() })
    expect(computeRuleHealth([r], { now: NOW }).counts.sleeping).toBe(0)
    expect(
      computeRuleHealth([r], { now: NOW, neverUsedAgeDays: 7 }).counts.sleeping,
    ).toBe(1)
  })

  it('respects custom vague patterns', () => {
    const r = rule({ matchCount: 20, targetFolderPath: '01測試桶' })
    expect(computeRuleHealth([r], { now: NOW }).counts.hotVague).toBe(0)
    expect(
      computeRuleHealth([r], { now: NOW, vagueTargetPatterns: [/測試桶/] }).counts.hotVague,
    ).toBe(1)
  })
})
