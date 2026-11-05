import { describe, expect, it } from 'vitest'
import {
  chunkRules,
  chunkTombstones,
  CHUNK_BYTE_TARGET,
  compareRuleSyncPriority,
  jsonByteSize,
  planSyncPayload,
  shouldSyncRule,
  TOMBSTONE_SYNC_CAP,
  unchunkArray,
  utf8ByteLength,
  type SyncPlanInput,
} from '@/shared/sync-chunks'
import { SYNC_TOMBSTONE_BUDGET_SHARE, SYNC_TOTAL_BUDGET_BYTES } from '@/shared/constants'
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
    for (const s of sources) expect(shouldSyncRule(rule({ source: s }))).toBe(true)
  })

  it('excludes auto_scan (re-derived per machine via initial scan)', () => {
    expect(shouldSyncRule(rule({ source: 'auto_scan' }))).toBe(false)
  })

  it('excludes orphaned rules (target folder gone, not useful on other machine)', () => {
    expect(shouldSyncRule(rule({ orphaned: true }))).toBe(false)
  })

  it('includes auto-disabled rules (the disable IS the sync intent)', () => {
    expect(
      shouldSyncRule(rule({ enabled: false, autoDisabledAt: '2026-01-01T00:00:00.000Z', autoDisabledReason: 'stale' })),
    ).toBe(true)
  })
})

describe('compareRuleSyncPriority (budget-guard drop order)', () => {
  it('enabled rules rank above auto-disabled ones', () => {
    const on = rule({ enabled: true, source: 'ai_confirmed' })
    const off = rule({ enabled: false, source: 'user_manual', matchCount: 999 })
    expect(compareRuleSyncPriority(on, off)).toBeLessThan(0)
  })

  it('among enabled rules, source authority wins (user_manual > ai_confirmed)', () => {
    const manual = rule({ enabled: true, source: 'user_manual', matchCount: 0 })
    const ai = rule({ enabled: true, source: 'ai_confirmed', matchCount: 500 })
    expect(compareRuleSyncPriority(manual, ai)).toBeLessThan(0)
  })

  it('same source → higher matchCount wins', () => {
    const proven = rule({ source: 'ai_confirmed', matchCount: 50 })
    const fresh = rule({ source: 'ai_confirmed', matchCount: 1 })
    expect(compareRuleSyncPriority(proven, fresh)).toBeLessThan(0)
  })

  it('sorting a mixed set puts the drop-me rules last', () => {
    const rules = [
      rule({ id: 'disabled', enabled: false, source: 'user_manual', matchCount: 100 }),
      rule({ id: 'ai-cold', source: 'ai_confirmed', matchCount: 0 }),
      rule({ id: 'manual-hot', source: 'user_manual', matchCount: 20 }),
      rule({ id: 'ai-hot', source: 'ai_confirmed', matchCount: 80 }),
    ]
    const order = [...rules].sort(compareRuleSyncPriority).map((r) => r.id)
    expect(order[0]).toBe('manual-hot')
    expect(order[order.length - 1]).toBe('disabled')
  })

  it('is a TOTAL order: exact ties break by id (two machines pick the same subset)', () => {
    const a = rule({ id: 'a' })
    const b = rule({ id: 'b' })
    expect(compareRuleSyncPriority(a, b)).toBeLessThan(0)
    expect(compareRuleSyncPriority(b, a)).toBeGreaterThan(0)
    expect([b, a].sort(compareRuleSyncPriority).map((r) => r.id)).toEqual(['a', 'b'])
  })
})

describe('chunkRules', () => {
  it('produces 0 chunks for empty input', () => {
    expect(chunkRules([])).toEqual([])
  })

  it('produces 1 chunk for a few rules', () => {
    const chunks = chunkRules([rule({ signal: 'a.com' }), rule({ signal: 'b.com' }), rule({ signal: 'c.com' })])
    expect(chunks).toHaveLength(1)
    expect(chunks[0]!.index).toBe(0)
    expect(chunks[0]!.items).toHaveLength(3)
  })

  it('respects ~6KB target per chunk', () => {
    const chunks = chunkRules(Array.from({ length: 100 }, (_, i) => rule({ signal: `r${i}.example.com` })))
    expect(chunks.length).toBeGreaterThan(1)
    for (const c of chunks) expect(JSON.stringify(c).length).toBeLessThan(CHUNK_BYTE_TARGET + 1024)
  })

  it('keeps every CJK-heavy chunk under the 8KB per-item byte limit (kQuotaBytesPerItem regression)', () => {
    const PER_ITEM_LIMIT = 8_192
    const rules = Array.from({ length: 60 }, (_, i) =>
      rule({
        id: `id-${i.toString().padStart(3, '0')}`,
        type: 'subject_keyword',
        signal: `台灣高等法院民事判決一一四年度上字第${i}號損害賠償等事件通知`,
        targetFolderPath: `01進行中案件/２０２４年度/王小明與陳大文股權轉讓糾紛/往來文件與書狀/第${i}類`,
      }),
    )
    const chunks = chunkRules(rules)
    expect(chunks.length).toBeGreaterThan(1)
    for (const c of chunks) {
      expect(jsonByteSize(c) + utf8ByteLength(`syncRules_${c.index}`)).toBeLessThan(PER_ITEM_LIMIT)
    }
    const worst = chunks.reduce((m, c) => Math.max(m, jsonByteSize(c)), 0)
    const worstChars = chunks.reduce((m, c) => Math.max(m, JSON.stringify(c).length), 0)
    expect(worst).toBeGreaterThan(worstChars * 1.3)
  })

  it('skips auto_scan + orphaned rules', () => {
    const chunks = chunkRules([
      rule({ signal: 'keep.com', source: 'user_manual' }),
      rule({ signal: 'skip-auto.com', source: 'auto_scan' }),
      rule({ signal: 'skip-orph.com', source: 'user_manual', orphaned: true }),
    ])
    expect(chunks.flatMap((c) => c.items).map((r) => r.signal)).toEqual(['keep.com'])
  })

  it('produces deterministic chunking (same input → same chunks)', () => {
    const rules = Array.from({ length: 30 }, (_, i) => rule({ id: `id-${i.toString().padStart(3, '0')}`, signal: `r${i}.com` }))
    const a = chunkRules(rules)
    const b = chunkRules([...rules].reverse())
    expect(a.length).toBe(b.length)
    for (let i = 0; i < a.length; i++) expect(a[i]!.items.map((r) => r.id)).toEqual(b[i]!.items.map((r) => r.id))
  })

  it('round-trips via unchunkArray (deterministic order preserved)', () => {
    const rules = Array.from({ length: 50 }, (_, i) => rule({ id: `id-${i.toString().padStart(3, '0')}`, signal: `r${i}.com` }))
    const reconstructed = unchunkArray(chunkRules(rules))
    expect(reconstructed.map((r) => r.id)).toEqual([...rules].map((r) => r.id).sort())
  })
})

describe('chunkTombstones (legacy v3 layout)', () => {
  function tomb(deletedAt: number, signalNorm: string): RuleTombstone {
    return { type: 'domain', signalNorm, targetFolderPath: 'X', deletedAt }
  }

  it('caps to TOMBSTONE_SYNC_CAP, keeping the most recent', () => {
    const tombs = Array.from({ length: TOMBSTONE_SYNC_CAP + 100 }, (_, i) => tomb(i * 1000, `t${i}.com`))
    const flat = unchunkArray(chunkTombstones(tombs))
    expect(flat).toHaveLength(TOMBSTONE_SYNC_CAP)
    const kept = new Set(flat.map((t) => t.signalNorm))
    expect(kept.has(`t${TOMBSTONE_SYNC_CAP + 99}.com`)).toBe(true)
    expect(kept.has(`t0.com`)).toBe(false)
  })

  it('produces 0 chunks for empty input', () => {
    expect(chunkTombstones([])).toEqual([])
  })
})

describe('unchunkArray', () => {
  it('handles chunks in arbitrary order via index', () => {
    expect(
      unchunkArray([
        { index: 2, items: ['c', 'd'] },
        { index: 0, items: ['a'] },
        { index: 1, items: ['b'] },
      ]),
    ).toEqual(['a', 'b', 'c', 'd'])
  })

  it('skips invalid chunk entries', () => {
    expect(
      unchunkArray([
        { index: 0, items: ['a'] },
        // @ts-expect-error testing invalid input
        null,
        { index: 1, items: ['b'] },
        // @ts-expect-error
        { index: 2 } as unknown,
      ]),
    ).toEqual(['a', 'b'])
  })
})

// ---- v4 budget planner -------------------------------------------------------

const PREFIX = { rows: 'syncRows4_', folders: 'syncFolders4_', tombs: 'syncTombs4_' }
const SHARE = Math.floor(SYNC_TOMBSTONE_BUDGET_SHARE * SYNC_TOTAL_BUDGET_BYTES)

function plan(over: Partial<SyncPlanInput>) {
  return planSyncPayload({
    rules: [],
    tombstones: [],
    budgetBytes: 80 * 1024,
    tombstoneShareBytes: SHARE,
    pushId: 'pushid000001',
    keyPrefix: PREFIX,
    now: Date.now(),
    ...over,
  })
}

function caseRule(i: number, over: Partial<Rule> = {}): Rule {
  return rule({
    id: `case-${String(i).padStart(4, '0')}`,
    type: 'compound',
    signal: JSON.stringify({
      conditions: [
        { type: 'domain', value: `firm${i % 40}.com.tw` },
        { type: 'subject_keyword', value: `一一四年度上字第${i}號損害賠償等事件 開庭通知` },
      ],
    }),
    targetFolderId: 'AAMkADAwMDAwMDAwLWJiYmItY2NjYy1kZGRk' + 'z'.repeat(80) + (i % 60),
    targetFolderPath: `01進行中案件/客戶第${i % 60}號`,
    source: 'ai_confirmed',
    matchCount: 500 - i,
    ...over,
  })
}

function cjkTomb(i: number, deletedAt: number): RuleTombstone {
  return {
    type: 'subject_keyword',
    signalNorm: `一一三年度訴字第${i}號請求返還不當得利事件言詞辯論通知書`,
    targetFolderPath: `01進行中案件/已結案/客戶第${i}號`,
    deletedAt,
  }
}

describe('Item 5: tombstones newest first, inside a bounded byte share', () => {
  it('300 older CJK tombstones no longer crowd out the 150 newest deletions', () => {
    const now = Date.now()
    const old = Array.from({ length: 300 }, (_, i) => cjkTomb(i, now - 90 * 86_400_000 + i))
    const newest = Array.from({ length: 150 }, (_, i) => cjkTomb(1000 + i, now - i * 1000))
    const p = plan({ tombstones: [...old, ...newest], rules: Array.from({ length: 300 }, (_, i) => caseRule(i)) })
    const kept = new Set(p.tombstones.map((t) => t[1] as string))
    const newestKept = newest.filter((t) => kept.has(t.signalNorm)).length
    const oldKept = old.filter((t) => kept.has(t.signalNorm)).length
    // Before: chunked by KEY order and truncated by chunk count → the newest
    // deletions (sorting after the old ones) were the ones dropped.
    expect(newestKept).toBeGreaterThan(0)
    if (oldKept > 0) expect(newestKept).toBe(newest.length)
    // Every kept tombstone is at least as new as every dropped one.
    const keptTimes = [...old, ...newest].filter((t) => kept.has(t.signalNorm)).map((t) => t.deletedAt)
    const droppedTimes = [...old, ...newest].filter((t) => !kept.has(t.signalNorm)).map((t) => t.deletedAt)
    if (droppedTimes.length > 0) expect(Math.min(...keptTimes)).toBeGreaterThanOrEqual(Math.max(...droppedTimes))
  })

  it('a big deletion history costs a budget-pressured library (almost) no rules', () => {
    const rules = Array.from({ length: 400 }, (_, i) => caseRule(i))
    const without = plan({ rules })
    const tombs = Array.from({ length: 500 }, (_, i) => cjkTomb(i, Date.now() - i * 60_000))
    const withTombs = plan({ rules, tombstones: tombs })
    expect(without.truncatedRuleCount).toBeGreaterThan(0) // really under pressure
    // Tombstone bytes stay within their share (plus chunk envelopes)…
    const tombBytes = withTombs.tombChunks.reduce((s, c) => s + jsonByteSize(c) + 16, 0)
    expect(tombBytes).toBeLessThanOrEqual(SHARE + 512)
    // …so rules lose at most that share, not the ~69 rules of the old order.
    const rowBytes = jsonByteSize(without.rows) / Math.max(1, without.rows.length)
    expect(without.fullRowCount - withTombs.fullRowCount).toBeLessThanOrEqual(Math.ceil(SHARE / rowBytes) + 2)
  })

  it('tombstones may use budget the rules leave unused', () => {
    const tombs = Array.from({ length: 450 }, (_, i) => cjkTomb(i, Date.now() - i * 60_000))
    const p = plan({ rules: [caseRule(1)], tombstones: tombs })
    const tombBytes = p.tombChunks.reduce((s, c) => s + jsonByteSize(c), 0)
    expect(tombBytes).toBeGreaterThan(SHARE)
    expect(p.bytes).toBeLessThanOrEqual(80 * 1024)
  })
})

describe('Item 13b: one-pass planner', () => {
  it('fills the budget exactly once: never over it, and the next rule would not fit', () => {
    const rules = Array.from({ length: 600 }, (_, i) => caseRule(i))
    const budget = 70 * 1024
    const p = plan({ rules, budgetBytes: budget })
    expect(p.bytes).toBeLessThanOrEqual(budget)
    expect(p.truncatedRuleCount).toBe(600 - p.fullRowCount)
    // Near-maximal: what's left is smaller than ~3 rows + a chunk envelope.
    const avgRow = jsonByteSize(p.rows) / p.rows.length
    expect(budget - p.bytes).toBeLessThan(3 * avgRow + 600)
    // The kept rows are exactly the top of the priority order.
    const keptIds = new Set(p.rows.map((e) => (Array.isArray(e) ? e[0] : e.i)))
    const top = [...rules].sort(compareRuleSyncPriority).slice(0, p.fullRowCount)
    expect(top.every((r) => keptIds.has(r.id))).toBe(true)
  })

  it('every chunk carries the pushId and stays under the per-item limit', () => {
    const p = plan({
      rules: Array.from({ length: 300 }, (_, i) => caseRule(i)),
      tombstones: Array.from({ length: 50 }, (_, i) => cjkTomb(i, i)),
    })
    for (const c of [...p.rowChunks, ...p.folderChunks, ...p.tombChunks]) {
      expect(c.p).toBe('pushid000001')
      expect(jsonByteSize(c) + 20).toBeLessThan(8_192)
    }
  })

  it('a 2000-rule library plans quickly (the old loop re-encoded once per dropped rule)', () => {
    const rules = Array.from({ length: 2000 }, (_, i) => caseRule(i))
    const t0 = Date.now()
    const p = plan({ rules })
    expect(Date.now() - t0).toBeLessThan(4000)
    expect(p.fullRowCount).toBeGreaterThan(100)
  })

  it('disabled rules ride as stubs when their full rows do not fit', () => {
    const rules = [
      ...Array.from({ length: 500 }, (_, i) => caseRule(i)),
      ...Array.from({ length: 30 }, (_, i) =>
        caseRule(900 + i, { enabled: false, autoDisabledReason: 'stale', autoDisabledAt: '2026-10-01T00:00:00.000Z' }),
      ),
    ]
    const p = plan({ rules })
    expect(p.stubCount).toBe(30)
    expect(p.truncatedRuleCount).toBe(530 - p.fullRowCount - 30)
  })
})
