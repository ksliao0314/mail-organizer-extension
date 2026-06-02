import { describe, expect, it } from 'vitest'
import {
  applyImport,
  parseRulesPayload,
  previewImport,
  RULE_PAYLOAD_TYPE,
  serializeRules,
} from '@/shared/rule-io'
import type { Rule } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    id: over.id ?? 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: 'X/Y',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  }
}

describe('serializeRules', () => {
  it('produces a valid versioned payload (v2 schema)', () => {
    const json = serializeRules([rule({ signal: 'foo.com' })], '0.0.1')
    const parsed = JSON.parse(json)
    expect(parsed.type).toBe(RULE_PAYLOAD_TYPE)
    expect(parsed.schemaVersion).toBe(2)
    expect(parsed.appVersion).toBe('0.0.1')
    expect(parsed.ruleCount).toBe(1)
    expect(parsed.rules).toHaveLength(1)
    expect(parsed.rules[0].signal).toBe('foo.com')
    expect(typeof parsed.exportedAt).toBe('string')
    // v2: tombstones field present (empty by default).
    expect(parsed.tombstones).toEqual([])
    expect(parsed.tombstoneCount).toBe(0)
  })

  it('round-trips through parseRulesPayload', () => {
    const original = [rule({ signal: 'a.com' }), rule({ signal: 'b.com', type: 'sender' })]
    const json = serializeRules(original, '0.0.1')
    const result = parseRulesPayload(json)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.rules).toHaveLength(2)
    expect(result.rules[0]!.signal).toBe('a.com')
    expect(result.rules[1]!.type).toBe('sender')
    expect(result.tombstones).toEqual([])
  })

  // ---- v2 tombstones (2026-05-22) ----------------------------------------
  describe('v2 tombstones round-trip', () => {
    it('serializes provided tombstones', () => {
      const tombs = [
        {
          type: 'domain' as const,
          signalNorm: 'spammer.com',
          targetFolderPath: 'A/B',
          deletedAt: 1716_000_000_000,
        },
      ]
      const json = serializeRules([rule()], '0.0.1', tombs)
      const parsed = JSON.parse(json)
      expect(parsed.tombstones).toHaveLength(1)
      expect(parsed.tombstones[0].signalNorm).toBe('spammer.com')
      expect(parsed.tombstoneCount).toBe(1)
    })

    it('parses tombstones back via parseRulesPayload', () => {
      const tombs = [
        {
          type: 'sender' as const,
          signalNorm: 'noisy@example.com',
          targetFolderPath: 'Bin',
          deletedAt: 1716_000_000_000,
        },
        {
          type: 'subject_keyword' as const,
          signalNorm: '電子報',
          targetFolderPath: 'Trash',
          deletedAt: 1716_001_000_000,
        },
      ]
      const json = serializeRules([rule()], '0.0.1', tombs)
      const result = parseRulesPayload(json)
      expect(result.ok).toBe(true)
      if (!result.ok) return
      expect(result.tombstones).toHaveLength(2)
      expect(result.tombstones[0]!.signalNorm).toBe('noisy@example.com')
    })

    it('drops malformed tombstone entries silently (validation per field)', () => {
      const raw = {
        type: RULE_PAYLOAD_TYPE,
        schemaVersion: 2,
        exportedAt: '2026-05-22T00:00:00Z',
        appVersion: '0.0.1',
        ruleCount: 0,
        rules: [],
        tombstones: [
          { type: 'domain', signalNorm: 'ok.com', targetFolderPath: 'X', deletedAt: 1 },
          { type: 'BADTYPE', signalNorm: 'x', targetFolderPath: 'X', deletedAt: 1 },
          { type: 'domain', signalNorm: '', targetFolderPath: 'X', deletedAt: 1 },
          { type: 'domain', signalNorm: 'y.com', targetFolderPath: '', deletedAt: 1 },
          { type: 'domain', signalNorm: 'z.com', targetFolderPath: 'X', deletedAt: 'notnum' },
        ],
      }
      const result = parseRulesPayload(JSON.stringify(raw))
      expect(result.ok).toBe(true)
      if (!result.ok) return
      // Only the first entry should survive validation.
      expect(result.tombstones).toHaveLength(1)
      expect(result.tombstones[0]!.signalNorm).toBe('ok.com')
    })

    it('treats payload with no tombstones field as empty tombstone list (back-compat)', () => {
      const raw = {
        type: RULE_PAYLOAD_TYPE,
        schemaVersion: 1, // legacy v1 payload
        exportedAt: '2026-01-01T00:00:00Z',
        appVersion: '0.0.0',
        ruleCount: 1,
        rules: [rule({ signal: 'legacy.com' })],
        // No tombstones field
      }
      const result = parseRulesPayload(JSON.stringify(raw))
      expect(result.ok).toBe(true)
      if (!result.ok) return
      expect(result.tombstones).toEqual([])
      expect(result.rules).toHaveLength(1)
    })
  })

  // ---- v2 preserves rule lifecycle metadata that v1 dropped --------------
  describe('v2 preserves auto-disable / override metadata', () => {
    it('round-trips autoDisabledAt + autoDisabledReason', () => {
      const r = rule({
        signal: 'sleeping.com',
        enabled: false,
        autoDisabledAt: '2026-04-01T00:00:00.000Z',
        autoDisabledReason: 'stale',
      })
      const json = serializeRules([r], '0.0.1')
      const result = parseRulesPayload(json)
      expect(result.ok).toBe(true)
      if (!result.ok) return
      const restored = result.rules[0]!
      expect(restored.autoDisabledAt).toBe('2026-04-01T00:00:00.000Z')
      expect(restored.autoDisabledReason).toBe('stale')
    })

    it('round-trips overrideCount + orphaned', () => {
      const r = rule({ signal: 'ovr.com', overrideCount: 5, orphaned: true, matchCount: 30 })
      const json = serializeRules([r], '0.0.1')
      const result = parseRulesPayload(json)
      expect(result.ok).toBe(true)
      if (!result.ok) return
      const restored = result.rules[0]!
      expect(restored.overrideCount).toBe(5)
      expect(restored.orphaned).toBe(true)
    })

    it('rejects invalid autoDisabledReason values silently (field dropped, rule preserved)', () => {
      const r = rule({ signal: 'x.com' })
      const json = serializeRules([r], '0.0.1')
      // Hand-inject a bogus value into the JSON and re-parse.
      const obj = JSON.parse(json)
      obj.rules[0].autoDisabledReason = 'made-up-value'
      const result = parseRulesPayload(JSON.stringify(obj))
      expect(result.ok).toBe(true)
      if (!result.ok) return
      expect(result.rules[0]!.autoDisabledReason).toBeUndefined()
    })
  })
})

describe('parseRulesPayload validation', () => {
  it('rejects non-JSON', () => {
    const r = parseRulesPayload('not json')
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toMatch(/JSON/)
  })

  it('rejects missing type marker', () => {
    const r = parseRulesPayload(JSON.stringify({ rules: [] }))
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toMatch(/type/)
  })

  it('rejects wrong type marker', () => {
    const r = parseRulesPayload(
      JSON.stringify({ type: 'mail-organizer-diagnostic', schemaVersion: 1, rules: [] }),
    )
    expect(r.ok).toBe(false)
  })

  it('rejects future schemaVersion', () => {
    const r = parseRulesPayload(
      JSON.stringify({ type: RULE_PAYLOAD_TYPE, schemaVersion: 99, rules: [] }),
    )
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toMatch(/schemaVersion/)
  })

  it('accepts empty rule list', () => {
    const r = parseRulesPayload(
      JSON.stringify({ type: RULE_PAYLOAD_TYPE, schemaVersion: 1, rules: [] }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.rules).toHaveLength(0)
  })

  it('rejects rule with invalid type enum', () => {
    const r = parseRulesPayload(
      serializeRules(
        // @ts-expect-error testing runtime guard
        [{ ...rule(), type: 'not-a-real-type' }],
        '0.0.1',
      ),
    )
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toMatch(/type/)
  })

  it('rejects rule with confidence out of range', () => {
    const r = parseRulesPayload(serializeRules([rule({ confidence: 1.5 })], '0.0.1'))
    expect(r.ok).toBe(false)
  })

  it('rejects rule with negative matchCount', () => {
    const r = parseRulesPayload(serializeRules([rule({ matchCount: -1 })], '0.0.1'))
    expect(r.ok).toBe(false)
  })

  it('rejects rule with empty signal', () => {
    const r = parseRulesPayload(serializeRules([rule({ signal: '' })], '0.0.1'))
    expect(r.ok).toBe(false)
  })

  it('rejects rule with empty targetFolderPath', () => {
    const r = parseRulesPayload(serializeRules([rule({ targetFolderPath: '' })], '0.0.1'))
    expect(r.ok).toBe(false)
  })

  it('allows empty targetFolderId (path-only rules are valid)', () => {
    const r = parseRulesPayload(serializeRules([rule({ targetFolderId: '' })], '0.0.1'))
    expect(r.ok).toBe(true)
  })

  it('reports which rule failed', () => {
    const good = rule({ signal: 'good.com' })
    const bad = rule({ signal: '' })
    const json = serializeRules([good, bad], '0.0.1')
    const r = parseRulesPayload(json)
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toMatch(/第 2 筆/)
  })
})

describe('previewImport', () => {
  it('counts duplicates by (type, signal)', () => {
    const existing = [
      rule({ type: 'domain', signal: 'a.com' }),
      rule({ type: 'domain', signal: 'b.com' }),
    ]
    const imported = [
      rule({ type: 'domain', signal: 'a.com' }), // duplicate
      rule({ type: 'sender', signal: 'a.com' }), // different type — not duplicate
      rule({ type: 'domain', signal: 'c.com' }), // new
    ]
    const p = previewImport(existing, imported)
    expect(p.existingCount).toBe(2)
    expect(p.totalToImport).toBe(3)
    expect(p.duplicateCount).toBe(1)
    expect(p.newCount).toBe(2)
  })
})

describe('applyImport', () => {
  it('replace strategy overwrites existing entirely', () => {
    const existing = [rule({ signal: 'a.com' }), rule({ signal: 'b.com' })]
    const imported = [rule({ signal: 'c.com' })]
    const result = applyImport(existing, imported, 'replace')
    expect(result).toHaveLength(1)
    expect(result[0]!.signal).toBe('c.com')
  })

  it('merge strategy keeps existing duplicates, adds new', () => {
    const existing = [
      rule({ id: 'existing-1', signal: 'a.com', matchCount: 99 }),
      rule({ id: 'existing-2', signal: 'b.com' }),
    ]
    const imported = [
      rule({ id: 'imp-1', signal: 'a.com', matchCount: 0 }), // duplicate by (type, signal)
      rule({ id: 'imp-2', signal: 'c.com' }), // new
    ]
    const result = applyImport(existing, imported, 'merge')
    expect(result).toHaveLength(3)
    // Existing rule for a.com should be preserved (matchCount 99, not 0)
    const aRule = result.find((r) => r.signal === 'a.com')
    expect(aRule?.id).toBe('existing-1')
    expect(aRule?.matchCount).toBe(99)
    // c.com was added
    expect(result.some((r) => r.signal === 'c.com')).toBe(true)
  })

  it('merge with no overlap concatenates', () => {
    const existing = [rule({ signal: 'a.com' })]
    const imported = [rule({ signal: 'b.com' })]
    const result = applyImport(existing, imported, 'merge')
    expect(result).toHaveLength(2)
  })

  it('replace with empty imported wipes everything', () => {
    const existing = [rule({ signal: 'a.com' }), rule({ signal: 'b.com' })]
    const result = applyImport(existing, [], 'replace')
    expect(result).toHaveLength(0)
  })
})
