// Error-rate retirement over ALL fires (2026-11-05, item 2).
//
// Overrides never increment matchCount (execute bumps overrideCount instead),
// so every accuracy gate that divided by matchCount — the effectiveConfidence
// demotion floor (≥10), the high-error-rate disable floor (≥20), the
// promotion gate — counted ACCEPTED hits only. A rule that was wrong every
// single time never reached either floor; and because overrides refresh
// lastUsedAt it never went stale either. Repro: 60 overrides → matchCount 0,
// confidence 0.85, still enabled and routing.

import { describe, expect, it } from 'vitest'
import {
  autoDisableStaleRules,
  bumpRuleOverrides,
  buildRuleIndex,
  effectiveConfidence,
  listRules,
  matchEmailWithIndex,
  newRule,
  ruleErrorRate,
  ruleFires,
} from '@/shared/rules'
import { computeRuleHealth } from '@/shared/rule-health'
import { setRules } from '@/shared/storage'
import type { Email, Rule } from '@/shared/types'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'compound',
      signal: JSON.stringify({
        conditions: [
          { type: 'domain', value: 'client.example' },
          { type: 'subject_keyword', value: '每月例行報告提交' },
        ],
      }),
      targetFolderId: FID,
      targetFolderPath: '03/A',
      confidence: 0.85,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-10-01T00:00:00.000Z',
    ...over,
  }
}

describe('fires = accepted hits + overrides', () => {
  it('ruleFires / ruleErrorRate', () => {
    expect(ruleFires({ matchCount: 3, overrideCount: 2 })).toBe(5)
    expect(ruleFires({ matchCount: 0 })).toBe(0)
    expect(ruleErrorRate({ matchCount: 3, overrideCount: 1 })).toBe(0.25)
    expect(ruleErrorRate({ matchCount: 0, overrideCount: 0 })).toBe(0)
  })

  it('effectiveConfidence is accepted / fires once there are ≥ 10 fires', () => {
    expect(effectiveConfidence({ confidence: 0.85, matchCount: 10, overrideCount: 10 })).toBe(0.5)
    // Old formula: (10 − 10) / 10 = 0 — overrides were counted against hits
    // twice (once as missing hits, once subtracted).
    expect(effectiveConfidence({ confidence: 0.85, matchCount: 0, overrideCount: 12 })).toBe(0)
    // Below the floor the configured confidence stands.
    expect(effectiveConfidence({ confidence: 0.85, matchCount: 2, overrideCount: 5 })).toBe(0.85)
  })
})

describe('the always-overridden rule (repro)', () => {
  it('60 overrides, 0 hits → demoted to 0 and auto-disabled as high-error-rate', async () => {
    const r = rule()
    await setRules([r])
    await bumpRuleOverrides(new Map([[r.id, 60]]))
    const [bumped] = await listRules()
    expect(bumped!.matchCount).toBe(0)
    expect(effectiveConfidence(bumped!)).toBe(0) // was: 0.85
    const report = await autoDisableStaleRules({ now: Date.parse('2026-11-05T00:00:00Z') })
    expect(report.byReason.highErrorRate).toBe(1)
    const [after] = await listRules()
    expect(after!.enabled).toBe(false)
    expect(after!.autoDisabledReason).toBe('high-error-rate')
  })

  it('a demoted rule loses its in-tier ordering to an untested sibling', () => {
    // Same compound tier, same specificity: the 0%-accurate rule sorts last.
    const bad = rule({ targetFolderPath: '03/Bad', matchCount: 0, overrideCount: 15, confidence: 0.95 })
    const fresh = rule({ targetFolderPath: '03/Fresh', confidence: 0.85 })
    const email = {
      Id: 'e1',
      Subject: 'RE: 每月例行報告提交',
      BodyPreview: '',
      From: { EmailAddress: { Address: 'pm@client.example' } },
      ToRecipients: [],
      ReceivedDateTime: '2026-11-01T00:00:00Z',
      ParentFolderId: 'inbox',
    } as Email
    const out = matchEmailWithIndex(email, buildRuleIndex([bad, fresh]))
    expect(out?.rule.targetFolderPath).toBe('03/Fresh')
  })

  it('a rule the user keeps correcting is not reported as "sleeping"', () => {
    const now = Date.parse('2026-11-05T00:00:00Z')
    const corrected = rule({
      createdAt: '2026-07-01T00:00:00.000Z',
      overrideCount: 5,
      lastUsedAt: '2026-11-04T00:00:00.000Z',
    })
    const report = computeRuleHealth([corrected], { now })
    expect(report.sleeping).toHaveLength(0)
  })
})
