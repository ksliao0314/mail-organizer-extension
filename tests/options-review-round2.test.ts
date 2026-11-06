// Options page, 2026-11-06 review round 2 (U10, U11a, U11b).
//
// U10  — 「回復到此」 is a rollback for EVERY machine now (synced rules the
//        snapshot lacks are deleted with tombstones that sync), but the
//        confirm still said only 「本機規則會被覆蓋」 and the result line never
//        said how many rules went.
// U11a — 「N 條規則錯誤率 > 30%」 → 查看 opened 規則健康度, which has no
//        accuracy bucket: the N rules were nowhere to be seen.
// U11b — opened with no Outlook tab, the page said 「等待資料夾樹載入…」
//        forever, and never asked again once Outlook was opened.

import { act } from 'react'
import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, Rule } from '@/shared/types'
import { buttonByText, buttonContaining, cleanup, click, mountApp, stubRuntime, type Msg, type Reply } from './helpers/ui-harness'

function status(over: Record<string, unknown> = {}) {
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
    syncEnabled: true,
    syncMachineId: 'm1',
    lastSyncAt: '',
    ...over,
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

const tree: MailFolderNode[] = [
  { id: 'ROOT', displayName: 'Cases', path: 'Cases', children: [{ id: 'F1', displayName: 'Acme', path: 'Cases/Acme', children: [] }] },
]

function sw(extra: (m: Msg) => Reply | undefined) {
  return stubRuntime((m) => {
    const custom = extra(m)
    if (custom) return custom
    switch (m.type) {
      case 'getStatus':
        return { ok: true, data: status() }
      case 'getRules':
        return { ok: true, data: { rules: [], conflicts: [] } }
      case 'getFolderTree':
        return { ok: true, data: { tree, count: 2 } }
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
  vi.restoreAllMocks()
  vi.useRealTimers()
})

async function renderOptions(hash: string) {
  window.location.hash = hash
  const { default: App } = await import('@/options/App')
  const r = await mountApp(App)
  root = r.root
  return r.host
}

describe('U10: 回復到此', () => {
  it('the confirm says the restore deletes on every machine; the result says how many', async () => {
    const confirm = vi.spyOn(window, 'confirm').mockReturnValue(true)
    const s = sw((m) => {
      if (m.type === 'getSyncStatus') {
        return { ok: true, data: { enabled: true, lastSyncAt: '', machineId: 'm1', bytesInUse: 0, bytesQuota: 102400 } }
      }
      if (m.type === 'listSyncBackups') {
        return {
          ok: true,
          data: {
            backups: [
              { snapshotAt: '2026-11-01T00:00:00.000Z', direction: 'manual', label: '手動', ruleCount: 3, tombstoneCount: 0 },
            ],
          },
        }
      }
      if (m.type === 'restoreSyncBackup') {
        return { ok: true, data: { restored: true, removedCount: 2, undeletedCount: 1 } }
      }
      return undefined
    })
    const host = await renderOptions('#data')
    const found = <T,>(el: T | undefined): T => {
      expect(el).toBeDefined()
      return el!
    }
    click(await vi.waitFor(() => found(buttonContaining(host, '同步歷史 / 回復'))))
    click(await vi.waitFor(() => found(buttonByText(host, '回復到此'))))

    const prompt = String(confirm.mock.calls[0]?.[0] ?? '')
    expect(prompt).toContain('所有機器')
    expect(prompt).toContain('刪除')
    expect(prompt).toContain('回復備份前')
    expect(s.last('restoreSyncBackup')).toEqual({ type: 'restoreSyncBackup', snapshotAt: '2026-11-01T00:00:00.000Z' })
    await vi.waitFor(() => expect(host.textContent).toContain('刪除 2 條備份之後新增的規則（所有機器）'))
    expect(host.textContent).toContain('恢復 1 條之後被刪除的規則')
  })
})

describe('U11a: 錯誤率 > 30% → 查看', () => {
  it('opens the rule library filtered to exactly those rules', async () => {
    const rules = [
      rule('often-wrong', { matchCount: 3, overrideCount: 8 }), // 11 fires, 73 % overridden
      rule('mostly-right', { matchCount: 20, overrideCount: 1 }),
      rule('few-fires', { matchCount: 2, overrideCount: 4 }), // < 10 fires: not in the warning
    ]
    sw((m) => (m.type === 'getRules' ? { ok: true, data: { rules, conflicts: [] } } : undefined))
    const host = await renderOptions('#rules')
    const warning = await vi.waitFor(() => {
      const li = Array.from(host.querySelectorAll('li')).find((x) => x.textContent?.includes('1 條規則錯誤率 > 30%'))
      expect(li).toBeTruthy()
      return li!
    })
    click(buttonContaining(warning, '查看')!)

    const chip = await vi.waitFor(() => {
      const b = buttonContaining(host, '錯誤率高')
      expect(b?.getAttribute('aria-pressed')).toBe('true')
      return b!
    })
    expect(chip.textContent).toContain('1')
    await vi.waitFor(() => expect(host.textContent).toContain('often-wrong.example'))
    expect(host.textContent).not.toContain('mostly-right.example')
    expect(host.textContent).not.toContain('few-fires.example')
  })
})

describe('U11b: folder tree with no Outlook tab', () => {
  it('says to open Outlook, and loads the tree once Outlook is there — no page reload', async () => {
    let outlookOpen = false
    const s = sw((m) => {
      if (m.type === 'getStatus') {
        return { ok: true, data: status({ owaConnected: outlookOpen, tokenValid: outlookOpen, syncEnabled: false }) }
      }
      if (m.type === 'getFolderTree' && !outlookOpen) {
        return { ok: false, code: 'NOT_ON_OWA', message: '請先打開並登入 outlook.office.com 分頁' }
      }
      return undefined
    })
    const host = await renderOptions('#engine')
    await vi.waitFor(() => expect(host.textContent).toContain('請打開並登入 Outlook 網頁版分頁'))
    const requests = s.count('getFolderTree')

    // The user opens Outlook and comes back to this tab.
    outlookOpen = true
    act(() => {
      window.dispatchEvent(new Event('focus'))
    })
    await vi.waitFor(() => expect(s.count('getFolderTree')).toBeGreaterThan(requests))
    await vi.waitFor(() => {
      expect(host.textContent).not.toContain('請打開並登入 Outlook 網頁版分頁')
      expect(host.textContent).not.toContain('等待資料夾樹載入') // the tree is in
    })
    expect(buttonContaining(host, '新增資料夾(含子資料夾)')!.disabled).toBe(false)
  })

  it('also retries on its own while the page stays visible', async () => {
    vi.useFakeTimers({ toFake: ['setInterval', 'clearInterval'] })
    let outlookOpen = false
    const s = sw((m) => {
      if (m.type === 'getStatus') {
        return { ok: true, data: status({ owaConnected: false, tokenValid: false, syncEnabled: false }) }
      }
      if (m.type === 'getFolderTree' && !outlookOpen) {
        return { ok: false, code: 'NOT_ON_OWA', message: 'no tab' }
      }
      return undefined
    })
    const host = await renderOptions('#engine')
    await vi.waitFor(() => expect(host.textContent).toContain('請打開並登入 Outlook 網頁版分頁'))
    const requests = s.count('getFolderTree')
    outlookOpen = true
    await act(async () => {
      await vi.advanceTimersByTimeAsync(10_000)
    })
    expect(s.count('getFolderTree')).toBeGreaterThan(requests)
    await vi.waitFor(() => expect(host.textContent).not.toContain('請打開並登入 Outlook 網頁版分頁'))
  })
})
