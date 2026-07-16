import { describe, expect, it } from 'vitest'
import {
  decodeRulesFromSync,
  encodeRulesForSync,
  type EncodedRules,
} from '@/shared/sync-codec'
import { jsonByteSize } from '@/shared/sync-chunks'
import type { Rule } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '01案件/子資料夾',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  }
}

describe('sync-codec round-trip', () => {
  it('encode → decode reproduces a minimal rule exactly', () => {
    const r = rule()
    const back = decodeRulesFromSync(encodeRulesForSync([r]))
    expect(back).toHaveLength(1)
    expect(back[0]).toEqual(r)
  })

  it('preserves every optional field across the round-trip', () => {
    const r = rule({
      type: 'compound',
      source: 'ai_overridden',
      signal: '台灣高等法院｜損害賠償',
      confidence: 0.83,
      matchCount: 42,
      overrideCount: 7,
      enabled: false,
      lastUsedAt: '2026-06-30T10:20:30.000Z',
      autoDisabledAt: '2026-07-01T00:00:00.000Z',
      autoDisabledReason: 'high-error-rate',
    })
    const back = decodeRulesFromSync(encodeRulesForSync([r]))
    expect(back[0]).toEqual(r)
  })

  it('round-trips every autoDisabledReason (incl. case_closed for two-phase retirement)', () => {
    const reasons: NonNullable<Rule['autoDisabledReason']>[] = [
      'stale',
      'high-error-rate',
      'legacy_token',
      'case_closed',
    ]
    const rules = reasons.map((reason, i) =>
      rule({
        id: `r-${i}`,
        enabled: false,
        autoDisabledAt: '2026-07-01T00:00:00.000Z',
        autoDisabledReason: reason,
      }),
    )
    const back = decodeRulesFromSync(encodeRulesForSync(rules))
    // case_closed 掉了 → 對機的 Phase 2 purge 謂詞永不觸發（不收斂）。
    expect(back.map((r) => r.autoDisabledReason)).toEqual(reasons)
  })

  it('round-trips all rule types and sources', () => {
    const types: Rule['type'][] = ['case_code', 'domain', 'compound', 'subject_keyword', 'sender']
    const sources: Rule['source'][] = ['auto_scan', 'ai_confirmed', 'ai_overridden', 'user_manual']
    const rules = types.flatMap((t, i) =>
      sources.map((src) => rule({ id: `${t}-${src}`, type: t, source: src, confidence: 0.5 + i * 0.1 })),
    )
    const back = decodeRulesFromSync(encodeRulesForSync(rules))
    // Order preserved; each field intact.
    expect(back).toEqual(rules)
  })

  it('a batch of realistic rules round-trips field-for-field', () => {
    const rules = Array.from({ length: 50 }, (_, i) =>
      rule({
        id: `id-${i}`,
        type: i % 2 ? 'subject_keyword' : 'case_code',
        signal: `一一四年度上字第${i}號`,
        targetFolderId: 'AAMk' + 'x'.repeat(100) + (i % 8), // 8 distinct folders
        targetFolderPath: `01進行中/2024/客戶${i % 8}`,
        matchCount: i,
        source: i % 3 ? 'ai_confirmed' : 'user_manual',
      }),
    )
    const back = decodeRulesFromSync(encodeRulesForSync(rules))
    expect(back).toEqual(rules)
  })
})

describe('sync-codec folder dedup', () => {
  it('stores each distinct folder exactly once', () => {
    const rules = Array.from({ length: 30 }, (_, i) =>
      rule({ id: `id-${i}`, targetFolderId: 'FID_A', targetFolderPath: 'A/path' }),
    )
    const enc = encodeRulesForSync(rules)
    expect(enc.folders).toHaveLength(1)
    expect(enc.rows).toHaveLength(30)
    expect(enc.rows.every((r) => r.f === enc.folders[0]!.x)).toBe(true)
  })

  it('is dramatically smaller than the full-rule serialisation for a folder-heavy library', () => {
    // 100 rules across 10 folders with realistic ~125-byte Graph ids.
    const rules = Array.from({ length: 100 }, (_, i) =>
      rule({
        id: `id-${i}`,
        targetFolderId: 'AAMkADAwMDAwMDAwLWJiYmItY2NjYy1kZGRk' + 'z'.repeat(90) + (i % 10),
        targetFolderPath: `01進行中案件/2024年度/客戶第${i % 10}號/往來書狀`,
        signal: `主旨關鍵字${i}`,
      }),
    )
    const full = jsonByteSize(rules)
    const enc = encodeRulesForSync(rules)
    const compact = jsonByteSize(enc.folders) + jsonByteSize(enc.rows)
    // Expect at least a 35% reduction — folder ids collapse 100→10.
    expect(compact).toBeLessThan(full * 0.65)
  })

  it('folder reference is STABLE when an unrelated folder is added (no write churn)', () => {
    // Regression: positional folder indices renumbered every later folder when
    // a new one appeared, rewriting most rows' f → chrome.storage.sync re-set
    // nearly all chunks and burned the write quota. Hash-based refs are intrinsic.
    const before = encodeRulesForSync([
      rule({ id: 'r1', targetFolderId: 'FID_MMM', targetFolderPath: 'M' }),
    ])
    const fBefore = before.rows[0]!.f
    // Add a rule whose folder sorts BEFORE FID_MMM — positional indexing would
    // have bumped FID_MMM from x=0 to x=1 and changed r1's f.
    const after = encodeRulesForSync([
      rule({ id: 'r1', targetFolderId: 'FID_MMM', targetFolderPath: 'M' }),
      rule({ id: 'r2', targetFolderId: 'FID_AAA', targetFolderPath: 'A' }),
    ])
    const r1After = after.rows.find((x) => x.i === 'r1')!
    expect(r1After.f).toBe(fBefore) // unchanged → its row chunk stays byte-identical
  })

  it('distinct folders get distinct references (no accidental collision)', () => {
    const enc = encodeRulesForSync(
      Array.from({ length: 50 }, (_, i) =>
        rule({ id: `r${i}`, targetFolderId: `FID_${i}`, targetFolderPath: `p${i}` }),
      ),
    )
    const refs = new Set(enc.folders.map((f) => f.x))
    expect(refs.size).toBe(50)
  })

  it('deterministic: same library → identical encoding (stable chunk membership)', () => {
    const rules = Array.from({ length: 20 }, (_, i) =>
      rule({ id: `id-${i}`, targetFolderId: `F${i % 5}`, targetFolderPath: `p${i % 5}` }),
    )
    const a = encodeRulesForSync(rules)
    const b = encodeRulesForSync([...rules].reverse())
    // Folder dictionary is order-independent; rows follow input order.
    expect(a.folders).toEqual(b.folders)
  })
})

describe('sync-codec defensive decode', () => {
  it('skips rows with no id', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'ok', t: 1, s: 'x', f: 'h0', c: 0.7, r: 3, ca: '2026-01-01T00:00:00.000Z' }, { f: 'h0' } as never],
    }
    const back = decodeRulesFromSync(enc)
    expect(back.map((r) => r.id)).toEqual(['ok'])
  })

  it('a missing folder reference resolves to empty strings, not a crash or mis-point', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'ha', id: 'FID_A', p: 'A' }],
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'nope', c: 0.7, r: 3, ca: '2026-01-01T00:00:00.000Z' }],
    }
    const back = decodeRulesFromSync(enc)
    expect(back[0]!.targetFolderId).toBe('')
    expect(back[0]!.targetFolderPath).toBe('')
  })

  it('unknown type / source codes fall back to safe defaults', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'r1', t: 99, s: 'x', f: 'h0', c: 0.7, r: 99, ca: '2026-01-01T00:00:00.000Z' }],
    }
    const back = decodeRulesFromSync(enc)
    expect(back[0]!.type).toBe('subject_keyword')
    expect(back[0]!.source).toBe('ai_confirmed')
  })

  it('tolerates a garbage payload without throwing', () => {
    expect(decodeRulesFromSync({ folders: null as never, rows: null as never })).toEqual([])
    expect(decodeRulesFromSync({} as never)).toEqual([])
  })

  it('folder reference survives even if folders arrive out of order (explicit x, not position)', () => {
    const enc: EncodedRules = {
      folders: [
        { x: 'hb', id: 'FID_B', p: 'B' },
        { x: 'ha', id: 'FID_A', p: 'A' },
      ],
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'hb', c: 0.7, r: 3, ca: '2026-01-01T00:00:00.000Z' }],
    }
    const back = decodeRulesFromSync(enc)
    expect(back[0]!.targetFolderId).toBe('FID_B')
  })
})
