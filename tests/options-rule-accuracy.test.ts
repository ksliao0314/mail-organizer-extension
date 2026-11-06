// One accuracy formula across the Options page (2026-11-06 audit E2).
//
// Accepted ÷ ALL fires (matchCount + overrideCount — overrides never bump
// matchCount) was used by the rule row, drawer and summary line only. The
// folder group header, its default collapse, the conflicts view and the
// cleanup list / warning still used (match − override) ÷ match: 2 hits + 4
// overrides showed 33 % in the row but −100 % in the header, and the cleanup
// warnings (≥ 5 / ≥ 10 MATCHES) never flagged a rule the user mostly overrides.

import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Rule } from '@/shared/types'
import { aggregateAccuracy, isLowAccuracy, ruleAccuracyPct } from '@/options/rule-accuracy'
import { cleanup, mountApp, stubRuntime } from './helpers/ui-harness'

function status() {
  return {
    owaConnected: true,
    tokenValid: true,
    apiKeyConfigured: true,
    apiKeyPreview: 'sk-…test',
    model: 'claude-sonnet-4-6',
    rulesCount: 0,
    excludePrefixes: [],
    closedCasePrefixes: [],
    batchSize: 50,
    aiConfidenceThreshold: 0.5,
    skipFlagged: true,
    showOwaFab: true,
    prefetchNextBatch: false,
    recentActivityIncludePrefixes: [],
    recentActivityIncludeLeafNames: [],
    internalDomains: [],
    primaryRootPath: 'Cases',
    internalSubjectCategories: [],
    aiIncludeFewShotExamples: true,
    harvestEnabled: false,
    syncEnabled: false,
    syncMachineId: 'm1',
    lastSyncAt: '',
  }
}

function rule(id: string, over: Partial<Rule> = {}): Rule {
  return {
    id,
    type: 'domain',
    signal: `${id}.example`,
    targetFolderId: 'F1',
    targetFolderPath: 'Cases/Acme',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-10-01T00:00:00.000Z',
    source: 'ai_confirmed',
    ...over,
  }
}

function sw(rules: Rule[], conflicts: unknown[] = []) {
  return stubRuntime((msg) => {
    switch (msg.type) {
      case 'getStatus':
        return { ok: true, data: status() }
      case 'getRules':
        return { ok: true, data: { rules: JSON.parse(JSON.stringify(rules)), conflicts } }
      case 'getFolderTree':
        return { ok: true, data: { tree: [], count: 0 } }
      default:
        return undefined
    }
  })
}

let root: Root | null = null
afterEach(() => {
  cleanup(root)
  root = null
  window.location.hash = ''
})

async function renderOptions(hash: string) {
  window.location.hash = hash
  const { default: App } = await import('@/options/App')
  const r = await mountApp(App)
  root = r.root
  return r.host
}

// 2 accepted hits + 4 overrides = 6 fires, 33 % accepted.
const MOSTLY_OVERRIDDEN = { matchCount: 2, overrideCount: 4 }

describe('rule-accuracy helpers', () => {
  it('accepted ÷ fires, never ÷ hits', () => {
    expect(ruleAccuracyPct(MOSTLY_OVERRIDDEN)).toBe(33)
    expect(ruleAccuracyPct({ matchCount: 0, overrideCount: 3 })).toBe(0)
    expect(ruleAccuracyPct({ matchCount: 0 })).toBeNull()
    expect(aggregateAccuracy([MOSTLY_OVERRIDDEN, { matchCount: 4 }])).toEqual({ fires: 10, accuracy: 0.6 })
    expect(aggregateAccuracy([{ matchCount: 0 }]).accuracy).toBeNull()
    expect(isLowAccuracy(MOSTLY_OVERRIDDEN, 5)).toBe(true) // 6 fires, 67 % overridden
    expect(isLowAccuracy(MOSTLY_OVERRIDDEN, 10)).toBe(false) // not enough fires
  })
})

describe('Options shows the same accuracy everywhere', () => {
  it('folder group header agrees with its row (33 %, not −100 %)', async () => {
    sw([rule('r1', MOSTLY_OVERRIDDEN)])
    const host = await renderOptions('#rules-library/all')
    const header = await vi.waitFor(() => {
      const el = host.querySelector<HTMLButtonElement>('button[aria-label$="Cases/Acme"]')
      expect(el).not.toBeNull()
      return el!
    })
    const row = await vi.waitFor(() => {
      const el = Array.from(host.querySelectorAll('tr')).find((tr) => tr.textContent?.includes('r1.example'))
      expect(el).toBeTruthy()
      return el!
    })
    expect(row.textContent).toContain('33%')
    expect(header.textContent).toContain('33%')
    expect(host.textContent).not.toContain('-100%')
  })

  it('conflicts view uses the same number', async () => {
    sw(
      [rule('r1', MOSTLY_OVERRIDDEN), rule('r2', { signal: 'r1.example', targetFolderPath: 'Cases/Other', matchCount: 4 })],
      [{ type: 'domain', signal: 'r1.example', ruleIds: ['r1', 'r2'], targets: ['Cases/Acme', 'Cases/Other'] }],
    )
    const host = await renderOptions('#rules-library/conflicts')
    await vi.waitFor(() => expect(host.textContent).toContain('→ Cases/Acme'))
    expect(host.textContent).toContain('命中 2 · 33%')
    expect(host.textContent).toContain('命中 4 · 100%')
    expect(host.textContent).not.toContain('-100%')
  })

  it('the cleanup list and the low-accuracy warning count fires, not matches', async () => {
    sw([
      rule('mostly-wrong', MOSTLY_OVERRIDDEN), // 6 fires, 67 % overridden
      rule('often-wrong', { matchCount: 3, overrideCount: 8 }), // 11 fires, 73 %
      rule('always-wrong', { matchCount: 0, overrideCount: 5 }), // fired 5×, all overridden
      rule('idle'), // never fired
    ])
    const host = await renderOptions('#rules')
    await vi.waitFor(() => expect(host.textContent).toContain('建議清理 TOP 5'))
    const text = host.textContent ?? ''
    expect(text).toContain('錯誤率 100%')
    expect(text).toContain('錯誤率 73%')
    expect(text).toContain('錯誤率 67%')
    // Only the rule that truly never fired is 「從未命中」 — not the one the
    // user overrode every single time.
    expect(text.match(/從未命中/g)).toHaveLength(1)
    // ≥ 10 fires and > 30 % overridden: only 'often-wrong'.
    expect(text).toContain('1 條規則錯誤率 > 30%')
    // Library line: 5 accepted of 22 fires.
    expect(text).toContain('22.7%')
  })
})
