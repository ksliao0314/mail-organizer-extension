import { describe, expect, it } from 'vitest'
import {
  chunkRules,
  chunkTombstones,
  CHUNK_BYTE_TARGET,
  shouldSyncRule,
  TOMBSTONE_SYNC_CAP,
  unchunkArray,
} from '@/shared/sync-chunks'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  }
}

describe('shouldSyncRule', () => {
  it('includes user_manual / ai_confirmed / ai_overridden', () => {
    const sources: RuleSource[] = ['user_manual', 'ai_confirmed', 'ai_overridden']
    for (const s of sources) {
      expect(shouldSyncRule(rule({ source: s }))).toBe(true)
    }
  })

  it('excludes auto_scan (re-derived per machine via initial scan)', () => {
    expect(shouldSyncRule(rule({ source: 'auto_scan' }))).toBe(false)
  })

  it('excludes orphaned rules (target folder gone, not useful on other machine)', () => {
    expect(shouldSyncRule(rule({ orphaned: true }))).toBe(false)
  })

  it('includes auto-disabled rules (the disable IS the sync intent)', () => {
    expect(
      shouldSyncRule(
        rule({
          enabled: false,
          autoDisabledAt: '2026-01-01T00:00:00.000Z',
          autoDisabledReason: 'stale',
        }),
      ),
    ).toBe(true)
  })
})

describe('chunkRules', () => {
  it('produces 0 chunks for empty input', () => {
    expect(chunkRules([])).toEqual([])
  })

  it('produces 1 chunk for a few rules', () => {
    const chunks = chunkRules([
      rule({ signal: 'a.com' }),
      rule({ signal: 'b.com' }),
      rule({ signal: 'c.com' }),
    ])
    expect(chunks).toHaveLength(1)
    expect(chunks[0]!.index).toBe(0)
    expect(chunks[0]!.items).toHaveLength(3)
  })

  it('respects ~6KB target per chunk', () => {
    // Each rule ~300 bytes. 100 rules ≈ 30 KB → expect ~5 chunks.
    const rules = Array.from({ length: 100 }, (_, i) =>
      rule({ signal: `r${i}.example.com` }),
    )
    const chunks = chunkRules(rules)
    expect(chunks.length).toBeGreaterThan(1)
    for (const c of chunks) {
      // Allow some over-estimate slack but not blow out.
      expect(JSON.stringify(c).length).toBeLessThan(CHUNK_BYTE_TARGET + 1024)
    }
  })

  it('skips auto_scan + orphaned rules', () => {
    const rules: Rule[] = [
      rule({ signal: 'keep.com', source: 'user_manual' }),
      rule({ signal: 'skip-auto.com', source: 'auto_scan' }),
      rule({ signal: 'skip-orph.com', source: 'user_manual', orphaned: true }),
    ]
    const chunks = chunkRules(rules)
    const flat = chunks.flatMap((c) => c.items)
    expect(flat.map((r) => r.signal)).toEqual(['keep.com'])
  })

  it('produces deterministic chunking (same input → same chunks)', () => {
    const rules = Array.from({ length: 30 }, (_, i) =>
      rule({ id: `id-${i.toString().padStart(3, '0')}`, signal: `r${i}.com` }),
    )
    const a = chunkRules(rules)
    const b = chunkRules([...rules].reverse()) // different input order
    // Same chunk count + same items per chunk in same order.
    expect(a.length).toBe(b.length)
    for (let i = 0; i < a.length; i++) {
      expect(a[i]!.items.map((r) => r.id)).toEqual(b[i]!.items.map((r) => r.id))
    }
  })

  it('round-trips via unchunkArray (deterministic order preserved)', () => {
    const rules = Array.from({ length: 50 }, (_, i) =>
      rule({ id: `id-${i.toString().padStart(3, '0')}`, signal: `r${i}.com` }),
    )
    const chunks = chunkRules(rules)
    const reconstructed = unchunkArray(chunks)
    expect(reconstructed).toHaveLength(50)
    // Ids should be in sorted order (deterministic chunking).
    expect(reconstructed.map((r) => r.id)).toEqual(
      [...rules].sort((a, b) => a.id.localeCompare(b.id)).map((r) => r.id),
    )
  })
})

describe('chunkTombstones', () => {
  function tomb(deletedAt: number, signalNorm: string): RuleTombstone {
    return {
      type: 'domain',
      signalNorm,
      targetFolderPath: 'X',
      deletedAt,
    }
  }

  it('caps to TOMBSTONE_SYNC_CAP, keeping the most recent', () => {
    const tombs = Array.from({ length: TOMBSTONE_SYNC_CAP + 100 }, (_, i) =>
      tomb(i * 1000, `t${i}.com`),
    )
    const chunks = chunkTombstones(tombs)
    const flat = unchunkArray(chunks)
    expect(flat).toHaveLength(TOMBSTONE_SYNC_CAP)
    // The newest 500 by deletedAt should be included.
    const expectedNewestIndex = TOMBSTONE_SYNC_CAP + 99 // index 0 = oldest
    const expectedOldestKept = expectedNewestIndex - TOMBSTONE_SYNC_CAP + 1
    const keptSignals = new Set(flat.map((t) => t.signalNorm))
    expect(keptSignals.has(`t${expectedNewestIndex}.com`)).toBe(true)
    expect(keptSignals.has(`t${expectedOldestKept}.com`)).toBe(true)
    expect(keptSignals.has(`t0.com`)).toBe(false)
  })

  it('produces 0 chunks for empty input', () => {
    expect(chunkTombstones([])).toEqual([])
  })
})

describe('unchunkArray', () => {
  it('handles chunks in arbitrary order via index', () => {
    const reconstructed = unchunkArray([
      { index: 2, items: ['c', 'd'] },
      { index: 0, items: ['a'] },
      { index: 1, items: ['b'] },
    ])
    expect(reconstructed).toEqual(['a', 'b', 'c', 'd'])
  })

  it('skips invalid chunk entries', () => {
    const reconstructed = unchunkArray([
      { index: 0, items: ['a'] },
      // @ts-expect-error testing invalid input
      null,
      { index: 1, items: ['b'] },
      // @ts-expect-error
      { index: 2 } as unknown,
    ])
    expect(reconstructed).toEqual(['a', 'b'])
  })
})
