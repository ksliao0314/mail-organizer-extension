// Options rule-list structural sharing (2026-11-05 audit).
//
// getRules round-trips the whole library through structured clone, so every
// refresh minted new objects for every rule — defeating the memoized RuleRow
// and re-rendering the full 2000-row table after each toggle. The refresh
// now keeps the previous object for every unchanged rule.

import { describe, expect, it } from 'vitest'
import { countOf, isUnknownHandler, reconcileRuleList, sameRule } from '@/options/rule-list'
import type { Rule } from '@/shared/types'

function rule(id: string, over: Partial<Rule> = {}): Rule {
  return {
    id,
    type: 'domain',
    signal: `${id}.example`,
    targetFolderId: 'F1',
    targetFolderPath: 'A/B',
    confidence: 0.7,
    matchCount: 3,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'ai_confirmed',
    ...over,
  }
}

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T

describe('reconcileRuleList', () => {
  it('returns the previous array itself when nothing changed', () => {
    const prev = [rule('a'), rule('b')]
    expect(reconcileRuleList(prev, clone(prev))).toBe(prev)
  })

  it('keeps identity for unchanged rules and swaps only the changed one', () => {
    const prev = [rule('a'), rule('b'), rule('c')]
    const next = clone(prev)
    next[1] = { ...next[1]!, enabled: false }
    const out = reconcileRuleList(prev, next)
    expect(out).not.toBe(prev)
    expect(out[0]).toBe(prev[0])
    expect(out[1]).toBe(next[1])
    expect(out[2]).toBe(prev[2])
  })

  it('follows the new order and drops / adds rules', () => {
    const prev = [rule('a'), rule('b')]
    const next = [clone(prev[1]!), rule('c')]
    const out = reconcileRuleList(prev, next)
    expect(out.map((r) => r.id)).toEqual(['b', 'c'])
    expect(out[0]).toBe(prev[1])
  })

  it('treats an added optional field as a change', () => {
    const a = rule('a')
    expect(sameRule(a, { ...a, lastUsedAt: '2026-02-01T00:00:00.000Z' })).toBe(false)
    expect(sameRule(a, clone(a))).toBe(true)
  })
})

describe('bulk response helpers', () => {
  it('normalises counts or id lists', () => {
    expect(countOf(3)).toBe(3)
    expect(countOf(['x', 'y'])).toBe(2)
    expect(countOf(undefined)).toBe(0)
  })
  it('recognises the SW "no such handler" codes', () => {
    expect(isUnknownHandler('UNKNOWN_REQUEST')).toBe(true)
    expect(isUnknownHandler('UNKNOWN_TYPE')).toBe(true)
    expect(isUnknownHandler('NOT_FOUND')).toBe(false)
  })
})
