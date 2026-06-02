import { describe, expect, it } from 'vitest'
import { buildExamplesBlock, selectExemplars } from '@/shared/classifier'
import type { Rule, RuleSource, RuleType } from '@/shared/types'

function rule(over: Partial<Rule>): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    matchCount: 20,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  }
}

describe('selectExemplars', () => {
  it('returns empty when no rules clear the matchCount threshold', () => {
    const rules = [rule({ matchCount: 1 }), rule({ matchCount: 2 })]
    expect(selectExemplars(rules)).toHaveLength(0)
  })

  it('drops disabled / orphaned rules', () => {
    const rules = [
      rule({ matchCount: 30, enabled: false }),
      rule({ matchCount: 30, orphaned: true }),
      rule({ matchCount: 30, signal: 'kept.com' }),
    ]
    const picked = selectExemplars(rules)
    expect(picked.map((r) => r.signal)).toEqual(['kept.com'])
  })

  it('prefers user_manual over ai_confirmed over auto_scan', () => {
    const a = rule({ id: 'a', source: 'auto_scan', matchCount: 100 })
    const b = rule({ id: 'b', source: 'ai_confirmed', matchCount: 100 })
    const c = rule({ id: 'c', source: 'user_manual', matchCount: 100 })
    const picked = selectExemplars([a, b, c])
    expect(picked[0]!.id).toBe('c') // user_manual first
  })

  it('hits each rule type once before doubling up', () => {
    // 3 user_manual domain rules + 1 user_manual case_code rule
    const rules: Rule[] = [
      rule({ id: 'd1', type: 'domain', signal: 'a.com', matchCount: 100 }),
      rule({ id: 'd2', type: 'domain', signal: 'b.com', matchCount: 99 }),
      rule({ id: 'd3', type: 'domain', signal: 'c.com', matchCount: 98 }),
      rule({ id: 'cc1', type: 'case_code', signal: '25A0067A', matchCount: 10 }),
    ]
    const picked = selectExemplars(rules)
    // Must include the lone case_code rule even though its matchCount is lower
    expect(picked.map((r) => r.id)).toContain('cc1')
  })

  it('caps at 8 examples', () => {
    const many: Rule[] = Array.from({ length: 20 }, (_, i) =>
      rule({ id: `r${i}`, signal: `domain${i}.com`, matchCount: 50 }),
    )
    expect(selectExemplars(many)).toHaveLength(8)
  })
})

describe('buildExamplesBlock', () => {
  it('returns empty string when no exemplars', () => {
    expect(buildExamplesBlock([])).toBe('')
  })

  it('formats each rule type humanly', () => {
    const rules: Rule[] = [
      rule({ type: 'domain', signal: 'kgi.com', targetFolderPath: '03/凱基證券' }),
      rule({ type: 'case_code', signal: '25A0067A', targetFolderPath: '03/X 案' }),
      rule({ type: 'subject_keyword', signal: '請款', targetFolderPath: '05/發票' }),
      rule({ type: 'sender', signal: 'admin@gov.tw', targetFolderPath: '03/行政' }),
    ]
    const block = buildExamplesBlock(rules)
    expect(block).toMatch(/寄件人網域 @kgi\.com → 03\/凱基證券/)
    expect(block).toMatch(/主旨含案件代號 25A0067A → 03\/X 案/)
    expect(block).toMatch(/主旨含「請款」 → 05\/發票/)
    expect(block).toMatch(/寄件人 admin@gov\.tw → 03\/行政/)
  })

  it('warns AI that examples are guidance, not absolute', () => {
    const r = rule({ matchCount: 30 })
    expect(buildExamplesBlock([r])).toMatch(/僅供參考/)
  })
})

// Sanity check: a realistic mix runs to completion
describe('selectExemplars integration', () => {
  it('handles a 50-rule mixed source set', () => {
    const types: RuleType[] = ['domain', 'case_code', 'subject_keyword', 'sender']
    const sources: RuleSource[] = ['auto_scan', 'ai_confirmed', 'ai_overridden', 'user_manual']
    const mix: Rule[] = []
    for (let i = 0; i < 50; i++) {
      mix.push(
        rule({
          id: `r${i}`,
          type: types[i % types.length]!,
          source: sources[i % sources.length]!,
          signal: `s${i}`,
          matchCount: 5 + (i % 30),
        }),
      )
    }
    const picked = selectExemplars(mix)
    expect(picked.length).toBeGreaterThan(0)
    expect(picked.length).toBeLessThanOrEqual(8)
  })
})
