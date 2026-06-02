import { describe, expect, it } from 'vitest'
import {
  addRules,
  addRulesFilteringTombstones,
  autoDisableStaleRules,
  dedupeRulesByKey,
  deleteRule,
  filterByCurrentTombstones,
  filterByTombstones,
  listRules,
  newRule,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import {
  addRuleTombstones,
  clearAllRuleTombstones,
  getRuleTombstones,
} from '@/shared/storage'
import type { Rule, RuleTombstone } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03/Y',
    confidence: 0.7,
    source: 'ai_confirmed',
    enabled: true,
    ...over,
  })
}

describe('tombstones — record on delete', () => {
  it('writes a tombstone when a rule is deleted', async () => {
    const r = rule()
    await addRules([r])
    await deleteRule(r.id)
    const tombs = await getRuleTombstones()
    expect(tombs.length).toBe(1)
    expect(tombs[0]).toMatchObject({
      type: 'domain',
      signalNorm: 'example.com',
      targetFolderPath: '03/Y',
    })
  })

  it('does not write a tombstone when delete misses (no such id)', async () => {
    await deleteRule('nonexistent-id')
    expect(await getRuleTombstones()).toHaveLength(0)
  })
})

describe('tombstones — block auto-regeneration', () => {
  it('filterByTombstones removes blocked candidates', () => {
    const candidates = [
      rule({ signal: 'kept.com', targetFolderPath: '03/A' }),
      rule({ signal: 'blocked.com', targetFolderPath: '03/B' }),
    ]
    const tombs: RuleTombstone[] = [
      { type: 'domain', signalNorm: 'blocked.com', targetFolderPath: '03/B', deletedAt: 1 },
    ]
    const result = filterByTombstones(candidates, tombs)
    expect(result.map((r) => r.signal)).toEqual(['kept.com'])
  })

  it('case_code tombstone uses uppercase comparison', () => {
    const candidates = [
      rule({ type: 'case_code', signal: '25A0067A', targetFolderPath: '03/X' }),
    ]
    const tombs: RuleTombstone[] = [
      { type: 'case_code', signalNorm: '25A0067A', targetFolderPath: '03/X', deletedAt: 1 },
    ]
    expect(filterByTombstones(candidates, tombs)).toHaveLength(0)
  })

  it('returns all candidates when no tombstones exist', () => {
    const candidates = [rule(), rule({ signal: 'b.com' })]
    expect(filterByTombstones(candidates, [])).toHaveLength(2)
  })

  it('filterByCurrentTombstones loads tombstones from storage', async () => {
    await addRuleTombstones([
      { type: 'domain', signalNorm: 'gone.com', targetFolderPath: '03/Y', deletedAt: 0 },
    ])
    const candidates = [
      rule({ signal: 'gone.com', targetFolderPath: '03/Y' }),
      rule({ signal: 'alive.com', targetFolderPath: '03/Y' }),
    ]
    const survived = await filterByCurrentTombstones(candidates)
    expect(survived.map((r) => r.signal)).toEqual(['alive.com'])
  })
})

describe('tombstones — cleared on user re-create', () => {
  it('upsertRule clears matching tombstone so the rule sticks', async () => {
    // Plant a tombstone manually
    await addRuleTombstones([
      { type: 'domain', signalNorm: 'example.com', targetFolderPath: '03/Y', deletedAt: 0 },
    ])
    expect(await getRuleTombstones()).toHaveLength(1)
    // Now user re-creates the same triple manually
    await upsertRule(rule({ source: 'user_manual', targetFolderPath: '03/Y' }))
    expect(await getRuleTombstones()).toHaveLength(0)
  })

  it('upsertRule does NOT clear an unrelated tombstone', async () => {
    await addRuleTombstones([
      { type: 'domain', signalNorm: 'other.com', targetFolderPath: '03/Z', deletedAt: 0 },
    ])
    await upsertRule(rule({ signal: 'example.com', targetFolderPath: '03/Y' }))
    expect(await getRuleTombstones()).toHaveLength(1)
  })
})

describe('tombstones — cap', () => {
  it('clearAllRuleTombstones wipes everything', async () => {
    await addRuleTombstones([
      { type: 'domain', signalNorm: 'a.com', targetFolderPath: 'X', deletedAt: 0 },
      { type: 'domain', signalNorm: 'b.com', targetFolderPath: 'Y', deletedAt: 1 },
    ])
    expect(await getRuleTombstones()).toHaveLength(2)
    await clearAllRuleTombstones()
    expect(await getRuleTombstones()).toHaveLength(0)
  })
})

// Sanity guard: filter logic for normalized signals — the rule's signal may
// be stored with @ prefix or mixed case, but tombstone.signalNorm is the
// canonical form (lowercase, no @).
describe('tombstones — signal normalization', () => {
  it('blocks a rule with @-prefixed signal when tombstone is unprefixed', () => {
    const candidates = [rule({ signal: '@Company-A.EXAMPLE', targetFolderPath: '03/Z' })]
    const tombs: RuleTombstone[] = [
      { type: 'domain', signalNorm: 'company-a.example', targetFolderPath: '03/Z', deletedAt: 0 },
    ]
    expect(filterByTombstones(candidates, tombs)).toHaveLength(0)
  })
})

describe('listRules after delete', () => {
  it('the deleted rule does not return from listRules', async () => {
    const r = rule()
    await addRules([r])
    expect((await listRules()).find((x) => x.id === r.id)).toBeTruthy()
    await deleteRule(r.id)
    expect((await listRules()).find((x) => x.id === r.id)).toBeFalsy()
  })
})

// ---- addRulesFilteringTombstones dedup-against-existing (added 2026-05-22) ----
//
// Pre-2026-05-22 the function only filtered tombstones — two concurrent
// generateAiConfirmedRules calls both snapshot rules, both pass their
// pre-mutex dedup, both write identical rules → duplicates in storage.
// Fix: dedup against the live rules array INSIDE the mutex closure.
describe('addRulesFilteringTombstones dedup', () => {
  it('skips a candidate already present in storage (same type/signal/target)', async () => {
    const existing = rule({ signal: 'a.com', targetFolderPath: 'X' })
    await addRules([existing])
    const dup = rule({ signal: 'a.com', targetFolderPath: 'X' })
    const { added, dropped } = await addRulesFilteringTombstones([dup])
    expect(added).toHaveLength(0)
    expect(dropped).toBe(1)
    // Should still be only one entry total.
    expect((await listRules()).filter((r) => r.signal === 'a.com')).toHaveLength(1)
  })

  it('keeps candidates with same signal but different target', async () => {
    const existing = rule({ signal: 'b.com', targetFolderPath: 'X' })
    await addRules([existing])
    const distinct = rule({ signal: 'b.com', targetFolderPath: 'Y' })
    const { added } = await addRulesFilteringTombstones([distinct])
    expect(added).toHaveLength(1)
    expect((await listRules()).filter((r) => r.signal === 'b.com')).toHaveLength(2)
  })

  it('dedups duplicates within the same batch', async () => {
    const a = rule({ signal: 'c.com', targetFolderPath: 'X' })
    const b = rule({ signal: 'c.com', targetFolderPath: 'X' }) // same triple
    const { added, dropped } = await addRulesFilteringTombstones([a, b])
    expect(added).toHaveLength(1)
    expect(dropped).toBe(1)
  })

  it('survives concurrent batches without producing duplicates (race fix)', async () => {
    // Simulates the original bug scenario: two parallel callers both want
    // to add the same rule. Pre-fix this yielded 2 entries.
    const a = rule({ signal: 'd.com', targetFolderPath: 'X' })
    const b = rule({ signal: 'd.com', targetFolderPath: 'X' })
    await Promise.all([
      addRulesFilteringTombstones([a]),
      addRulesFilteringTombstones([b]),
    ])
    const final = (await listRules()).filter((r) => r.signal === 'd.com')
    expect(final).toHaveLength(1)
  })

  it('still respects tombstones in the same call', async () => {
    // Mix: candidate1 tombstoned, candidate2 not. Only candidate2 lands.
    const r1 = rule({ signal: 'e1.com', targetFolderPath: 'X' })
    const r2 = rule({ signal: 'e2.com', targetFolderPath: 'Y' })
    await addRuleTombstones([
      {
        type: 'domain',
        signalNorm: 'e1.com',
        targetFolderPath: 'X',
        deletedAt: Date.now(),
      } as RuleTombstone,
    ])
    const { added } = await addRulesFilteringTombstones([r1, r2])
    expect(added.map((r) => r.signal)).toEqual(['e2.com'])
  })

  // ---- Auto re-enable for previously auto-disabled rules (2026-05-22) ----
  //
  // When the daily sweep auto-disables a rule and later the same signal+
  // target gets re-learned (AI confirms it again), the dedup path should
  // RE-ENABLE the dormant rule instead of skipping. User-disabled rules
  // (autoDisabledAt absent) stay disabled — that's user intent.
  it('re-enables auto-disabled rule when same signal+target re-learned', async () => {
    // Set up a soft-disabled rule directly (high-error-rate / legacy_token
    // both soft-disable; stale is hard-delete and never re-enables).
    const r = rule({ signal: 'foo.com', targetFolderPath: 'X', source: 'ai_confirmed' })
    r.enabled = false
    r.autoDisabledAt = new Date().toISOString()
    r.autoDisabledReason = 'high-error-rate'
    await addRules([r])
    let after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledAt).toBeDefined()

    // Now re-learn the same signal+target
    const dup = rule({ signal: 'foo.com', targetFolderPath: 'X', source: 'ai_confirmed' })
    const result = await addRulesFilteringTombstones([dup])
    expect(result.added).toHaveLength(0) // no NEW row
    expect(result.reEnabled).toHaveLength(1) // existing re-enabled
    after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
    expect(after?.autoDisabledAt).toBeUndefined()
  })

  it('does NOT re-enable user-disabled rules (no autoDisabledAt)', async () => {
    const original = await upsertRule(
      rule({ signal: 'bar.com', targetFolderPath: 'Y', source: 'ai_confirmed' }),
    )
    // User manually disables
    await toggleRule(original.id, false, { actor: 'user' })
    let after = (await listRules()).find((r) => r.id === original.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledAt).toBeUndefined()

    const dup = rule({ signal: 'bar.com', targetFolderPath: 'Y' })
    const result = await addRulesFilteringTombstones([dup])
    expect(result.added).toHaveLength(0)
    expect(result.reEnabled).toHaveLength(0)
    after = (await listRules()).find((r) => r.id === original.id)
    expect(after?.enabled).toBe(false) // stayed disabled
  })
})

// ---- autoDisableStaleRules — high error rate (2026-05-22) ------------------
describe('autoDisableStaleRules — high error rate branch', () => {
  function staleRule(over: Partial<Rule> = {}): Rule {
    return newRule({
      type: 'domain',
      signal: 'x.com',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.7,
      source: 'ai_confirmed',
      enabled: true,
      ...over,
    })
  }

  it('disables rule with matchCount ≥ 20 and overrideRate ≥ 0.5', async () => {
    // Build a rule with 20 hits and 12 overrides (60% error)
    const r = staleRule({ signal: 'badroute.com' })
    r.matchCount = 20
    r.overrideCount = 12
    await addRules([r])

    const { disabled, byReason } = await autoDisableStaleRules({
      now: Date.now(),
      staleDays: 365, // not stale by age
    })
    expect(disabled.find((d) => d.id === r.id)).toBeDefined()
    expect(byReason.highErrorRate).toBe(1)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledReason).toBe('high-error-rate')
  })

  it('does NOT disable rule with matchCount ≥ 20 and overrideRate < 0.5', async () => {
    const r = staleRule({ signal: 'goodroute.com' })
    r.matchCount = 30
    r.overrideCount = 5 // ~17%
    await addRules([r])
    const { byReason } = await autoDisableStaleRules({
      now: Date.now(),
      staleDays: 365,
    })
    expect(byReason.highErrorRate).toBe(0)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })

  it('does NOT disable user_manual rule even with high error rate', async () => {
    const r = staleRule({ signal: 'sacred.com', source: 'user_manual' })
    r.matchCount = 30
    r.overrideCount = 20
    await addRules([r])
    await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })

  it('requires minimum sample (20) before high-error-rate disable', async () => {
    // 10 hits, 8 overrides (80% error but only 10 samples).
    const r = staleRule({ signal: 'newrule.com' })
    r.matchCount = 10
    r.overrideCount = 8
    await addRules([r])
    await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })
})

// ---- autoDisableStaleRules — legacy_token branch (2026-05-27) ------------
// Pre-2026-05-27 design used `extractCandidateSubjectTokens` to build
// 3–8 char Chinese fragment subject_keyword rules (source = auto_scan).
// Those rules over-fire on unrelated mail containing the fragment.
// Proactive cleanup: any auto_scan + subject_keyword rule with a
// fragment-shape signal gets auto-disabled regardless of activity.
describe('autoDisableStaleRules — legacy_token branch', () => {
  it('disables auto_scan + subject_keyword rule with token-shape signal', async () => {
    const r = newRule({
      type: 'subject_keyword',
      signal: '甲公司', // 4-char pure Chinese → legacy token shape
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '03/X',
      confidence: 0.75,
      source: 'auto_scan',
      enabled: true,
    })
    r.matchCount = 50 // high usage doesn't save it — design is over-broad
    r.lastUsedAt = new Date().toISOString() // recently used either
    await addRules([r])

    const { disabled, byReason } = await autoDisableStaleRules({
      now: Date.now(),
      staleDays: 365,
    })
    expect(disabled.find((d) => d.id === r.id)).toBeDefined()
    expect(byReason.legacyToken).toBe(1)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledReason).toBe('legacy_token')
  })

  it('keeps subject_keyword rule with court-case signal (structural ID)', async () => {
    const r = newRule({
      type: 'subject_keyword',
      signal: '112訴204', // court case — not a legacy fragment
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '03/X',
      confidence: 0.85,
      source: 'auto_scan', // hypothetical — auto_scan never made these, but defend anyway
      enabled: true,
    })
    r.matchCount = 5
    await addRules([r])

    await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })

  it('keeps post-cutoff ai_confirmed subject_keyword (new full_subject design)', async () => {
    // Post-2026-05-27 ai_confirmed subject_keyword = new design's full
    // subject signal (e.g. P6 internal-domain fallback). Stays.
    const r = newRule({
      type: 'subject_keyword',
      signal: '請款通知會議',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '05/發票',
      confidence: 0.8,
      source: 'ai_confirmed',
      enabled: true,
    })
    // newRule() sets createdAt = now, which is post-cutoff.
    r.matchCount = 5
    r.lastUsedAt = new Date().toISOString()
    await addRules([r])

    await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })

  it('disables pre-cutoff ai_confirmed subject_keyword (old extractSubjectFeature output)', async () => {
    // Rule was created BEFORE the 2026-05-27 redesign — its signal came
    // from extractSubjectFeature (4+ Chinese token). Over-broad. Retire.
    const r = newRule({
      type: 'subject_keyword',
      signal: '甲公司',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '03/X',
      confidence: 0.7,
      source: 'ai_confirmed',
      enabled: true,
    })
    r.createdAt = '2026-05-20T00:00:00.000Z' // pre-cutoff
    r.matchCount = 30
    r.lastUsedAt = new Date().toISOString()
    await addRules([r])

    const { byReason } = await autoDisableStaleRules({
      now: Date.now(),
      staleDays: 365,
    })
    expect(byReason.legacyToken).toBe(1)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledReason).toBe('legacy_token')
  })

  it('disables pre-cutoff ai_overridden subject_keyword (old lenient extractor)', async () => {
    // extractSubjectFeatureLenient produced shorter tokens via override path.
    const r = newRule({
      type: 'subject_keyword',
      signal: '三菱',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '03/Y',
      confidence: 0.7,
      source: 'ai_overridden',
      enabled: true,
    })
    r.createdAt = '2026-05-22T12:00:00.000Z'
    r.matchCount = 10
    await addRules([r])

    const { byReason } = await autoDisableStaleRules({
      now: Date.now(),
      staleDays: 365,
    })
    expect(byReason.legacyToken).toBe(1)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
  })

  it('keeps pre-cutoff user_manual subject_keyword (sacred)', async () => {
    const r = newRule({
      type: 'subject_keyword',
      signal: '甲公司',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '03/X',
      confidence: 0.95,
      source: 'user_manual',
      enabled: true,
    })
    r.createdAt = '2026-05-20T00:00:00.000Z' // pre-cutoff — but user_manual!
    r.matchCount = 5
    r.lastUsedAt = new Date().toISOString()
    await addRules([r])

    await autoDisableStaleRules({ now: Date.now(), staleDays: 365 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true) // user_manual stays
  })
})

// ---- autoDisableStaleRules — lastUsedAt stale branch (2026-05-27) ---------
// Subject-type rules (subject_keyword, compound) get auto-disabled when
// they've not matched for > staleDays days. Catches "thread ran for a
// few weeks then moved on" — without this the subject-as-signal redesign
// would leak rules forever.
describe('autoDisableStaleRules — lastUsedAt stale branch', () => {
  it('disables rule with lastUsedAt older than staleDays (matchCount > 0)', async () => {
    // Rule fired 200 days ago, hasn't been used since.
    const now = Date.now()
    const longAgo = now - 200 * 86_400_000
    const r = newRule({
      type: 'subject_keyword',
      signal: '關於某案件的詢問',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.8,
      source: 'ai_confirmed',
      enabled: true,
    })
    r.matchCount = 5 // has fired before
    r.lastUsedAt = new Date(longAgo).toISOString()
    await addRules([r])

    const { disabled, byReason } = await autoDisableStaleRules({
      now,
      staleDays: 100, // 200 > 100 → stale
    })
    // Stale rules are HARD-DELETED — not in `disabled` (soft-disable list)
    // and not in storage. byReason.stale still counts (telemetry).
    expect(disabled.find((d) => d.id === r.id)).toBeUndefined()
    expect(byReason.stale).toBe(1)
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after).toBeUndefined() // hard-deleted, gone from storage
  })

  it('hard-deletes stale rule WITHOUT writing a tombstone', async () => {
    // The design lets future learning re-discover the same signal/target
    // if mail flow resumes — that requires the tombstone library to stay
    // clean of stale-deletion marks.
    const now = Date.now()
    const r = newRule({
      type: 'subject_keyword',
      signal: '已棄置的主旨關鍵字',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.8,
      source: 'ai_confirmed',
      enabled: true,
    })
    r.matchCount = 5
    r.lastUsedAt = new Date(now - 200 * 86_400_000).toISOString()
    await addRules([r])
    const tombsBefore = (await getRuleTombstones()).length

    await autoDisableStaleRules({ now, staleDays: 100 })

    const tombsAfter = (await getRuleTombstones()).length
    expect(tombsAfter).toBe(tombsBefore) // no new tombstone for stale delete
  })

  it('keeps rule that fired recently (lastUsedAt within staleDays)', async () => {
    const now = Date.now()
    const recently = now - 30 * 86_400_000 // 30 days ago
    const r = newRule({
      type: 'subject_keyword',
      signal: '近期還在用的主旨',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.8,
      source: 'ai_confirmed',
      enabled: true,
    })
    r.matchCount = 12
    r.lastUsedAt = new Date(recently).toISOString()
    await addRules([r])

    await autoDisableStaleRules({ now, staleDays: 100 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
  })

  it('does NOT auto-disable user_manual subject rules even when stale by lastUsedAt', async () => {
    const now = Date.now()
    const longAgo = now - 200 * 86_400_000
    const r = newRule({
      type: 'subject_keyword',
      signal: '使用者手動建立的主旨規則',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.95,
      source: 'user_manual',
      enabled: true,
    })
    r.matchCount = 5
    r.lastUsedAt = new Date(longAgo).toISOString()
    await addRules([r])

    await autoDisableStaleRules({ now, staleDays: 100 })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true) // sacred — never auto-disabled
  })
})

// ---- toggleRule clears autoDisabledAt on re-enable (2026-05-22) -----------
describe('toggleRule clears auto-disable provenance on re-enable', () => {
  function ruleWithAuto(over: Partial<Rule> = {}): Rule {
    return newRule({
      type: 'domain',
      signal: 'z.com',
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'Z',
      confidence: 0.7,
      source: 'ai_confirmed',
      enabled: false,
      ...over,
    })
  }

  it('user re-enable clears autoDisabledAt + autoDisabledReason', async () => {
    const r = ruleWithAuto()
    r.autoDisabledAt = new Date().toISOString()
    r.autoDisabledReason = 'stale'
    await addRules([r])
    await toggleRule(r.id, true, { actor: 'user' })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(true)
    expect(after?.autoDisabledAt).toBeUndefined()
    expect(after?.autoDisabledReason).toBeUndefined()
  })

  it('user disable leaves any existing auto-disable trace untouched (edge case)', async () => {
    // Rule already auto-disabled; user calls toggle(disabled) — no-op
    // change but still shouldn't gain auto-disable metadata.
    const r = ruleWithAuto()
    r.autoDisabledAt = '2026-01-01T00:00:00.000Z'
    r.autoDisabledReason = 'stale'
    await addRules([r])
    await toggleRule(r.id, false, { actor: 'user' })
    const after = (await listRules()).find((rr) => rr.id === r.id)
    expect(after?.enabled).toBe(false)
    expect(after?.autoDisabledAt).toBe('2026-01-01T00:00:00.000Z')
  })
})

// ---- dedupeRulesByKey — historical duplicate cleanup (2026-05-27) -------
describe('dedupeRulesByKey', () => {
  it('collapses identical (type, signal, target) triplets, preserving best survivor', async () => {
    // 3 ai_confirmed dups for company-b.example → 11 sports. They differ only in
    // id and matchCount. Higher matchCount should survive.
    const r1 = rule({ type: 'domain', signal: 'company-b.example', targetFolderPath: '03/11 sports' })
    r1.matchCount = 3
    const r2 = rule({ type: 'domain', signal: 'company-b.example', targetFolderPath: '03/11 sports' })
    r2.matchCount = 10
    const r3 = rule({ type: 'domain', signal: 'company-b.example', targetFolderPath: '03/11 sports' })
    r3.matchCount = 5
    await addRules([r1, r2, r3])

    const { removed, groupsAffected } = await dedupeRulesByKey()
    expect(removed).toBe(2)
    expect(groupsAffected).toBe(1)
    const after = await listRules()
    expect(after).toHaveLength(1)
    expect(after[0]?.id).toBe(r2.id) // highest matchCount kept
  })

  it('user_manual rule wins over ai_confirmed even with lower matchCount', async () => {
    const aiR = rule({ type: 'domain', signal: 'foo.com', targetFolderPath: 'X', source: 'ai_confirmed' })
    aiR.matchCount = 50
    const userR = rule({ type: 'domain', signal: 'foo.com', targetFolderPath: 'X', source: 'user_manual' })
    userR.matchCount = 0 // never matched but it's user_manual
    await addRules([aiR, userR])

    await dedupeRulesByKey()
    const after = await listRules()
    expect(after).toHaveLength(1)
    expect(after[0]?.source).toBe('user_manual')
  })

  it('enabled rule wins over disabled with same source', async () => {
    const disabledR = rule({ type: 'domain', signal: 'bar.com', targetFolderPath: 'Y', source: 'ai_confirmed' })
    disabledR.enabled = false
    disabledR.matchCount = 20
    const enabledR = rule({ type: 'domain', signal: 'bar.com', targetFolderPath: 'Y', source: 'ai_confirmed' })
    enabledR.matchCount = 5
    await addRules([disabledR, enabledR])

    await dedupeRulesByKey()
    const after = await listRules()
    expect(after).toHaveLength(1)
    expect(after[0]?.id).toBe(enabledR.id)
  })

  it('keeps unique-triplet rules untouched (different targets remain)', async () => {
    // company-b.example → A and company-b.example → B are a real conflict (different
    // targets), NOT a duplicate. Both stay.
    const a = rule({ type: 'domain', signal: 'company-b.example', targetFolderPath: 'A' })
    const b = rule({ type: 'domain', signal: 'company-b.example', targetFolderPath: 'B' })
    await addRules([a, b])

    const { removed } = await dedupeRulesByKey()
    expect(removed).toBe(0)
    expect(await listRules()).toHaveLength(2)
  })

  it('does NOT write tombstones for removed duplicates', async () => {
    // Hard delete = no tombstone, otherwise future re-learning of the
    // same signal/target would be blocked by the tombstone library.
    const r1 = rule({ type: 'domain', signal: 'baz.com', targetFolderPath: 'Z' })
    const r2 = rule({ type: 'domain', signal: 'baz.com', targetFolderPath: 'Z' })
    await addRules([r1, r2])
    const tombsBefore = (await getRuleTombstones()).length
    await dedupeRulesByKey()
    const tombsAfter = (await getRuleTombstones()).length
    expect(tombsAfter).toBe(tombsBefore)
  })
})
