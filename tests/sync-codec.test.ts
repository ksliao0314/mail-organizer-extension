import { describe, expect, it } from 'vitest'
import {
  compactCompoundNorm,
  compactCompoundSignal,
  compactTime,
  decodeRulePayload,
  decodeRulesFromSync,
  decodeTombstone,
  encodeDisabledStub,
  encodeRuleRow,
  encodeRulesForSync,
  encodeTombstone,
  expandTime,
  sanitizeLegacyRule,
  sanitizeLegacyTombstone,
  type CompactRuleRow,
  type EncodedRules,
} from '@/shared/sync-codec'
import { encodeCompound, normalizeSignal } from '@/shared/rules'
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

  it('preserves every optional field across the round-trip (incl. editedAt)', () => {
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
      editedAt: 1_780_000_000_123,
    })
    const back = decodeRulesFromSync(encodeRulesForSync([r]))
    expect(back[0]).toEqual(r)
  })

  it('round-trips every autoDisabledReason, incl. override + superseded (item 12)', () => {
    const reasons: NonNullable<Rule['autoDisabledReason']>[] = [
      'stale',
      'high-error-rate',
      'legacy_token',
      'case_closed',
      'override',
      'superseded',
    ]
    const rules = reasons.map((reason, i) =>
      rule({ id: `r-${i}`, enabled: false, autoDisabledAt: '2026-07-01T00:00:00.000Z', autoDisabledReason: reason }),
    )
    const back = decodeRulesFromSync(encodeRulesForSync(rules))
    expect(back.map((r) => r.autoDisabledReason)).toEqual(reasons)
    // Stubs carry the reason too.
    const stubs = decodeRulePayload({ folders: [], rows: rules.map(encodeDisabledStub) }).stubs
    expect(stubs.map((s) => s.autoDisabledReason)).toEqual(reasons)
  })

  it('round-trips all rule types and sources', () => {
    const types: Rule['type'][] = ['case_code', 'domain', 'compound', 'subject_keyword', 'sender']
    const sources: Rule['source'][] = ['auto_scan', 'ai_confirmed', 'ai_overridden', 'user_manual']
    const rules = types.flatMap((t, i) =>
      sources.map((src) => rule({ id: `${t}-${src}`, type: t, source: src, confidence: 0.5 + i * 0.1 })),
    )
    expect(decodeRulesFromSync(encodeRulesForSync(rules))).toEqual(rules)
  })

  it('a batch of realistic rules round-trips field-for-field', () => {
    const rules = Array.from({ length: 50 }, (_, i) =>
      rule({
        id: `id-${i}`,
        type: i % 2 ? 'subject_keyword' : 'case_code',
        signal: `一一四年度上字第${i}號`,
        targetFolderId: 'AAMk' + 'x'.repeat(100) + (i % 8),
        targetFolderPath: `01進行中/2024/客戶${i % 8}`,
        matchCount: i,
        source: i % 3 ? 'ai_confirmed' : 'user_manual',
      }),
    )
    expect(decodeRulesFromSync(encodeRulesForSync(rules))).toEqual(rules)
  })
})

describe('Item 13a: compact compound signals', () => {
  const compound = encodeCompound([
    { type: 'domain', value: 'kgi.com' },
    { type: 'subject_keyword', value: '某某股份有限公司 損害賠償事件 進度報告' },
  ])

  it('a canonical compound signal travels as [kind, value] pairs and round-trips exactly', () => {
    const r = rule({ type: 'compound', signal: compound })
    const row = encodeRuleRow(r, 'f')!
    expect(Array.isArray(row.s)).toBe(true)
    const back = decodeRulesFromSync({ folders: [{ x: 'f', id: r.targetFolderId, p: r.targetFolderPath }], rows: [row] })
    expect(back[0]!.signal).toBe(compound)
  })

  it('saves the escaped {"conditions":…} wrapper (~80 B per compound row)', () => {
    const r = rule({ type: 'compound', signal: compound })
    const v4 = jsonByteSize(encodeRuleRow(r, 'f'))
    const v3 = jsonByteSize({ ...encodeRuleRow(r, 'f'), s: compound, ca: r.createdAt })
    expect(v3 - v4).toBeGreaterThanOrEqual(70)
  })

  it('a non-canonical compound (key order / extra field) stays a string and still round-trips', () => {
    const odd = JSON.stringify({ conditions: [{ value: 'kgi.com', type: 'domain' }] })
    const extra = JSON.stringify({ conditions: [{ type: 'domain', value: 'kgi.com', note: 'x' }] })
    for (const signal of [odd, extra, 'not json']) {
      expect(compactCompoundSignal(signal)).toBeNull()
      const r = rule({ type: 'compound', signal })
      const row = encodeRuleRow(r, 'f')!
      expect(row.s).toBe(signal)
      const back = decodeRulesFromSync({ folders: [{ x: 'f', id: r.targetFolderId, p: r.targetFolderPath }], rows: [row] })
      expect(back[0]!.signal).toBe(signal)
    }
  })

  it('compound tombstone keys are compacted too and round-trip exactly', () => {
    const norm = normalizeSignal('compound', compound)
    expect(compactCompoundNorm(norm)).not.toBeNull()
    const t = { type: 'compound' as const, signalNorm: norm, targetFolderPath: '案件/甲', deletedAt: 123 }
    const enc = encodeTombstone(t)!
    expect(typeof enc[1]).not.toBe('string')
    expect(decodeTombstone(enc)).toEqual(t)
    expect(jsonByteSize(enc)).toBeLessThan(jsonByteSize(t) - 60)
  })

  it('canonical ISO timestamps travel as epoch ms; anything else stays verbatim', () => {
    expect(compactTime('2026-01-01T00:00:00.000Z')).toBe(Date.UTC(2026, 0, 1))
    expect(compactTime('2026-01-01T00:00:00Z')).toBe('2026-01-01T00:00:00Z')
    expect(expandTime(Date.UTC(2026, 0, 1))).toBe('2026-01-01T00:00:00.000Z')
    const r = rule({ createdAt: '2026-05-22T00:00:00Z', lastUsedAt: 'garbage' })
    expect(decodeRulesFromSync(encodeRulesForSync([r]))[0]).toEqual(r)
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
    expect(enc.rows.every((r) => (r as CompactRuleRow).f === enc.folders[0]!.x)).toBe(true)
  })

  it('is dramatically smaller than the full-rule serialisation for a folder-heavy library', () => {
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
    expect(jsonByteSize(enc.folders) + jsonByteSize(enc.rows)).toBeLessThan(full * 0.65)
  })

  it('folder reference is STABLE when an unrelated folder is added (no write churn)', () => {
    const before = encodeRulesForSync([rule({ id: 'r1', targetFolderId: 'FID_MMM', targetFolderPath: 'M' })])
    const fBefore = (before.rows[0] as CompactRuleRow).f
    const after = encodeRulesForSync([
      rule({ id: 'r1', targetFolderId: 'FID_MMM', targetFolderPath: 'M' }),
      rule({ id: 'r2', targetFolderId: 'FID_AAA', targetFolderPath: 'A' }),
    ])
    const r1After = after.rows.find((x) => (x as CompactRuleRow).i === 'r1') as CompactRuleRow
    expect(r1After.f).toBe(fBefore)
  })

  it('distinct folders get distinct references (no accidental collision)', () => {
    const enc = encodeRulesForSync(
      Array.from({ length: 50 }, (_, i) => rule({ id: `r${i}`, targetFolderId: `FID_${i}`, targetFolderPath: `p${i}` })),
    )
    expect(new Set(enc.folders.map((f) => f.x)).size).toBe(50)
  })

  it('deterministic: same library → identical encoding (stable chunk membership)', () => {
    const rules = Array.from({ length: 20 }, (_, i) =>
      rule({ id: `id-${i}`, targetFolderId: `F${i % 5}`, targetFolderPath: `p${i % 5}` }),
    )
    expect(encodeRulesForSync(rules).folders).toEqual(encodeRulesForSync([...rules].reverse()).folders)
  })
})

describe('sync-codec defensive decode (item 7d: skip, never coerce)', () => {
  const ca = '2026-01-01T00:00:00.000Z'

  it('skips rows with no id', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'ok', t: 1, s: 'x', f: 'h0', c: 0.7, r: 3, ca }, { f: 'h0' } as never],
    }
    expect(decodeRulesFromSync(enc).map((r) => r.id)).toEqual(['ok'])
  })

  it('a missing folder reference SKIPS the row (no empty target replacing a good copy)', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'ha', id: 'FID_A', p: 'A' }],
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'nope', c: 0.7, r: 3, ca }],
    }
    const out = decodeRulePayload(enc)
    expect(out.rules).toEqual([])
    expect(out.skipped).toBe(1)
  })

  it('an unknown type code SKIPS the row (was coerced to subject_keyword)', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'r1', t: 99, s: 'x', f: 'h0', c: 0.7, r: 3, ca }],
    }
    expect(decodeRulesFromSync(enc)).toEqual([])
  })

  it('an unknown source code SKIPS the row (coercion could demote a user_manual rule)', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'h0', c: 0.7, r: 99, ca }],
    }
    expect(decodeRulesFromSync(enc)).toEqual([])
  })

  it('an unknown reason code keeps the rule and drops only the reason', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'h0', c: 0.7, r: 3, ca, e: 0, dr: 99 }],
    }
    const [r] = decodeRulesFromSync(enc)
    expect(r!.enabled).toBe(false)
    expect(r!.autoDisabledReason).toBeUndefined()
  })

  it('an unknown compound-kind code SKIPS the row', () => {
    const enc: EncodedRules = {
      folders: [{ x: 'h0', id: 'F', p: 'p' }],
      rows: [{ i: 'r1', t: 2, s: [[7, 'x']], f: 'h0', c: 0.7, r: 3, ca }],
    }
    expect(decodeRulesFromSync(enc)).toEqual([])
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
      rows: [{ i: 'r1', t: 1, s: 'x', f: 'hb', c: 0.7, r: 3, ca }],
    }
    expect(decodeRulesFromSync(enc)[0]!.targetFolderId).toBe('FID_B')
  })

  it('tombstone rows with unknown types / malformed fields are skipped', () => {
    expect(decodeTombstone([99, 'x', 'p', 1])).toBeNull()
    expect(decodeTombstone([1, 'x', 'p'])).toBeNull()
    expect(decodeTombstone([1, 'x', 'p', 'soon'])).toBeNull()
    expect(decodeTombstone([1, 'x.com', 'p', 5])).toEqual({ type: 'domain', signalNorm: 'x.com', targetFolderPath: 'p', deletedAt: 5 })
  })

  it('legacy (v1/v2) full rule objects are validated field by field', () => {
    expect(sanitizeLegacyRule({ ...rule(), type: 'bogus' })).toBeNull()
    expect(sanitizeLegacyRule({ ...rule(), source: 'bogus' })).toBeNull()
    expect(sanitizeLegacyRule({ ...rule(), autoDisabledReason: 'bogus' })!.autoDisabledReason).toBeUndefined()
    const ok = rule({ lastUsedAt: '2026-02-02T00:00:00.000Z' })
    expect(sanitizeLegacyRule(ok)).toEqual(ok)
    expect(sanitizeLegacyTombstone({ type: 'domain', signalNorm: 'a', targetFolderPath: 'p', deletedAt: 'x' })).toBeNull()
  })
})

describe('disabled stubs', () => {
  it('[id, editedAt, reason, disabledAt] round-trips; a manual disable is just [id, editedAt]', () => {
    const auto = rule({ id: 'a', enabled: false, autoDisabledReason: 'case_closed', autoDisabledAt: '2026-07-01T00:00:00.000Z', editedAt: 99 })
    const manual = rule({ id: 'm', enabled: false, editedAt: 7 })
    expect(encodeDisabledStub(manual)).toEqual(['m', 7])
    const out = decodeRulePayload({ folders: [], rows: [encodeDisabledStub(auto), encodeDisabledStub(manual)] })
    expect(out.rules).toEqual([])
    expect(out.stubs).toEqual([
      { id: 'a', editedAt: 99, autoDisabledReason: 'case_closed', autoDisabledAt: '2026-07-01T00:00:00.000Z' },
      { id: 'm', editedAt: 7 },
    ])
  })
})
