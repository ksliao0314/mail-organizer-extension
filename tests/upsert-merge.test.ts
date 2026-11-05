// Editing a rule from a stale Options page (2026-11-05, item 6).
//
// upsertRule did `next[idx] = <the UI's whole object>`: the page's copy is
// as old as the page, so every hit / override / lastUsedAt bumped since it
// loaded was wiped by an edit (repro: 37 hits bumped after page load → 0
// after the edit), and a stale copy of a rule deleted meanwhile was simply
// re-inserted. Now: merge only the form-owned fields onto the stored rule;
// mode 'edit' on a missing id → NOT_FOUND.

import { beforeAll, describe, expect, it } from 'vitest'
import {
  bumpRuleHits,
  bumpRuleOverrides,
  deleteRule,
  listRules,
  newRule,
  RuleNotFoundError,
  upsertRule,
} from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, setRules } from '@/shared/storage'
import type { Rule } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const FID2 = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE4'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: FID,
      targetFolderPath: '03/A',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  }
}

describe('upsertRule merges editable fields onto the stored rule', () => {
  it('repro: hits bumped after the page loaded survive an edit from that page', async () => {
    const original = rule()
    await setRules([original])
    const pageCopy = { ...original } // what the Options page holds
    await bumpRuleHits(new Map([[original.id, 37]]))
    await bumpRuleOverrides(new Map([[original.id, 2]]))
    const stored = (await listRules())[0]!

    const saved = await upsertRule(
      { ...pageCopy, targetFolderId: FID2, targetFolderPath: '03/B', confidence: 0.6 },
      { actor: 'user', mode: 'edit' },
    )
    const [after] = await listRules()
    // mutateRules stamps editedAt on the stored copy (sync, 2026-11-05).
    expect(after!.editedAt).toBeGreaterThan(0)
    expect({ ...after, editedAt: undefined }).toEqual({ ...saved, editedAt: undefined })
    expect(after!.matchCount).toBe(37) // was 0
    expect(after!.overrideCount).toBe(2)
    expect(after!.lastUsedAt).toBe(stored.lastUsedAt)
    expect(after!.createdAt).toBe('2026-01-01T00:00:00.000Z')
    expect(after!.targetFolderPath).toBe('03/B')
    expect(after!.confidence).toBe(0.6)
  })

  it("mode 'edit' on a rule deleted meanwhile → NOT_FOUND, no resurrection, tombstone kept", async () => {
    const r = rule()
    await setRules([r])
    await deleteRule(r.id, { actor: 'user' })
    expect(await getRuleTombstones()).toHaveLength(1)
    await expect(upsertRule({ ...r, confidence: 0.6 }, { mode: 'edit' })).rejects.toBeInstanceOf(
      RuleNotFoundError,
    )
    await expect(upsertRule(r, { mode: 'edit' })).rejects.toMatchObject({ code: 'NOT_FOUND' })
    expect(await listRules()).toHaveLength(0)
    expect(await getRuleTombstones()).toHaveLength(1)
  })

  it("mode 'create' and the legacy no-mode call insert a new rule (and clear its tombstone)", async () => {
    await addRuleTombstones([
      { type: 'domain', signalNorm: 'client.example', targetFolderPath: '03/A', deletedAt: 1 },
    ])
    await upsertRule(rule(), { mode: 'create' })
    await upsertRule(rule({ signal: 'other.example' }))
    expect(await listRules()).toHaveLength(2)
    expect(await getRuleTombstones()).toHaveLength(0)
  })

  it('re-enabling through the form clears auto-disable provenance', async () => {
    const r = rule({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: '2026-10-01T00:00:00.000Z' })
    await setRules([r])
    const saved = await upsertRule({ ...r, enabled: true }, { mode: 'edit' })
    expect(saved.enabled).toBe(true)
    expect(saved.autoDisabledReason).toBeUndefined()
    expect(saved.autoDisabledAt).toBeUndefined()
  })

  it('engine-owned fields in the UI copy are ignored; the per-type cap still applies', async () => {
    const r = rule({ matchCount: 5 })
    await setRules([r])
    const saved = await upsertRule(
      { ...r, matchCount: 0, source: 'user_manual', confidence: 0.99 },
      { mode: 'edit' },
    )
    expect(saved.source).toBe('ai_confirmed')
    expect(saved.matchCount).toBe(5)
    expect(saved.confidence).toBe(0.7) // domain cap
  })

  it('a re-target clears a stale orphaned flag', async () => {
    const r = rule({ orphaned: true })
    await setRules([r])
    const saved = await upsertRule({ ...r, targetFolderId: FID2, targetFolderPath: '03/B' }, { mode: 'edit' })
    expect(saved.orphaned).toBe(false)
  })
})

describe("SW case 'upsertRule' contract", () => {
  let sw: SwHarness
  beforeAll(async () => {
    sw = await loadServiceWorker()
  })

  it('maps NOT_FOUND for an edit of a deleted rule', async () => {
    const r = rule()
    const res = await sw.send({ type: 'upsertRule', rule: r, mode: 'edit' })
    expect(res).toEqual({ ok: false, code: 'NOT_FOUND', message: '這條規則已在別處被刪除，請重新整理' })
    expect(await listRules()).toHaveLength(0)
  })

  it('returns the STORED rule (live stats), not the page copy', async () => {
    const r = rule()
    await setRules([r])
    await bumpRuleHits(new Map([[r.id, 4]]))
    const res = await sw.send({ type: 'upsertRule', rule: { ...r, confidence: 0.65 }, mode: 'edit' })
    expect(res.ok).toBe(true)
    const saved = (res.data as { rule: Rule }).rule
    expect(saved.matchCount).toBe(4)
    expect(saved.confidence).toBe(0.65)
  })
})
