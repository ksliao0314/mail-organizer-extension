// Options page keeps in step with writes made elsewhere and never writes
// stale snapshots back (2026-11-05 audit). Renders the REAL Options <App/>
// against a stub SW (see helpers/ui-harness).
//
//   - external 'settings' / 'rules' writes (sync pull, popup, sweep) trigger
//     a re-fetch; other storage areas don't
//   - 歸類偏好 saves ONLY the field the user edited (it used to write all
//     three from its mount-time snapshot, reverting synced-in values)
//   - drawer save sends mode:'edit'; NOT_FOUND (deleted elsewhere) is
//     explained, refreshed, and closes the drawer
//   - bulk actions are ONE bulkRuleAction call + one refresh, with a
//     per-rule fallback (still one refresh) for SWs without the handler
//   - the confidence slider is clamped to the SW's per-type cap for
//     system-learned rules
//   - storage usage card renders when the SW supports it, hides otherwise

import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { MailFolderNode, Rule } from '@/shared/types'
import {
  buttonByText,
  buttonContaining,
  cleanup,
  click,
  latestStorageListener,
  mountApp,
  stubRuntime,
  typeInto,
  type Msg,
  type Reply,
} from './helpers/ui-harness'

const tree: MailFolderNode[] = [
  {
    id: 'ROOT',
    displayName: 'Cases',
    path: 'Cases',
    children: [{ id: 'F1', displayName: 'Acme', path: 'Cases/Acme', children: [] }],
  },
]

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
    internalDomains: ['corp.example'],
    primaryRootPath: 'Cases',
    internalSubjectCategories: ['工時'],
    aiIncludeFewShotExamples: true,
    harvestEnabled: false,
    syncEnabled: false,
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
    matchCount: 4,
    enabled: true,
    createdAt: '2026-10-01T00:00:00.000Z',
    source: 'ai_confirmed',
    ...over,
  }
}

function sw(opts: { rules?: Rule[]; extra?: (msg: Msg) => Reply | undefined } = {}) {
  let rules = opts.rules ?? []
  const h = stubRuntime((msg) => {
    const custom = opts.extra?.(msg)
    if (custom) return custom
    switch (msg.type) {
      case 'getStatus':
        return { ok: true, data: status({ rulesCount: rules.length }) }
      case 'getRules':
        return { ok: true, data: { rules: JSON.parse(JSON.stringify(rules)), conflicts: [] } }
      case 'getFolderTree':
        return { ok: true, data: { tree, count: 2 } }
      case 'setSettings':
        return { ok: true }
      default:
        return undefined
    }
  })
  return { ...h, setRules: (next: Rule[]) => (rules = next) }
}

let root: Root | null = null
afterEach(() => {
  cleanup(root)
  root = null
  window.location.hash = ''
})

async function renderOptions(hash = '') {
  window.location.hash = hash
  const { default: App } = await import('@/options/App')
  const r = await mountApp(App)
  root = r.root
  return r.host
}

describe('Options live refresh', () => {
  it('re-fetches status on an external settings write and rules on a rules write — local area only', async () => {
    const s = sw()
    await renderOptions()
    await vi.waitFor(() => expect(s.count('getStatus')).toBeGreaterThan(0))
    await vi.waitFor(() => expect(s.count('getRules')).toBeGreaterThan(0))
    const listener = latestStorageListener()
    const statusBefore = s.count('getStatus')
    const rulesBefore = s.count('getRules')

    listener({ settings: { newValue: {} } }, 'sync')
    listener({ cachedToken: { newValue: {} } }, 'session')
    await new Promise((r) => setTimeout(r, 350))
    expect(s.count('getStatus')).toBe(statusBefore)

    listener({ settings: { newValue: {} } }, 'local')
    await vi.waitFor(() => expect(s.count('getStatus')).toBe(statusBefore + 1))
    expect(s.count('getRules')).toBe(rulesBefore)

    listener({ rules: { newValue: [] } }, 'local')
    await vi.waitFor(() => expect(s.count('getRules')).toBe(rulesBefore + 1))
  })
})

describe('歸類偏好 card', () => {
  it('saves only the edited field (no stale write-back of the others)', async () => {
    const s = sw()
    const host = await renderOptions('#engine')
    const input = await vi.waitFor(() => {
      const el = host.querySelector<HTMLInputElement>('input[placeholder="輸入幾個你想用來分類內部信的關鍵字"]')
      expect(el?.value).toBe('工時')
      return el!
    })
    typeInto(input, '工時、薪資')
    click(buttonByText(host, '儲存'))
    await vi.waitFor(() => expect(s.count('setSettings')).toBe(1))
    expect(s.last('setSettings')!.patch).toEqual({ internalSubjectCategories: ['工時', '薪資'] })
  })
})

describe('read-modify-write from fresh settings', () => {
  it('近日活動 filter removes from FRESH settings and patches only that list', async () => {
    let prefixes = ['Cases/A']
    const s = sw({
      extra: (m) =>
        m.type === 'getStatus'
          ? {
              ok: true,
              data: status({
                recentActivityIncludePrefixes: prefixes,
                recentActivityIncludeLeafNames: ['待簽核'],
              }),
            }
          : undefined,
    })
    const host = await renderOptions('#engine')
    const remove = await vi.waitFor(() => {
      const b = host.querySelector<HTMLButtonElement>('button[aria-label="移除 Cases/A"]')
      expect(b).not.toBeNull()
      return b!
    })
    prefixes = ['Cases/A', 'Cases/B'] // synced in from another machine; page not refreshed yet
    click(remove)
    await vi.waitFor(() => expect(s.count('setSettings')).toBe(1))
    // Old code: { prefixes: [], leaves: ['待簽核'] } from the stale snapshot.
    expect(s.last('setSettings')!.patch).toEqual({ recentActivityIncludePrefixes: ['Cases/B'] })
  })

  it('結案歸檔 checkbox adds to the FRESH closedCasePrefixes list', async () => {
    let closed: string[] = []
    const s = sw({
      extra: (m) =>
        m.type === 'getStatus'
          ? {
              ok: true,
              data: status({ excludePrefixes: ['Archive', 'Old'], closedCasePrefixes: closed }),
            }
          : undefined,
    })
    const host = await renderOptions('#data')
    const box = await vi.waitFor(() => {
      const label = Array.from(host.querySelectorAll('label')).find(
        (l) => l.textContent?.trim() === 'Archive',
      )
      const input = label?.querySelector<HTMLInputElement>('input[type="checkbox"]')
      expect(input).toBeTruthy()
      return input!
    })
    closed = ['Old'] // another machine ticked 'Old' meanwhile
    click(box)
    await vi.waitFor(() => expect(s.count('setSettings')).toBe(1))
    expect(s.last('setSettings')!.patch).toEqual({ closedCasePrefixes: ['Old', 'Archive'] })
  })
})

async function openDrawerFor(host: HTMLElement, signal: string) {
  const cell = await vi.waitFor(() => {
    const el = Array.from(host.querySelectorAll('td')).find((td) => td.textContent?.includes(signal))
    expect(el).toBeTruthy()
    return el!
  })
  click(cell)
  return await vi.waitFor(() => {
    const d = document.querySelector('[role="dialog"][aria-label="規則詳細"]')
    expect(d).not.toBeNull()
    return d as HTMLElement
  })
}

describe('rule drawer save', () => {
  it("sends mode:'edit', and on NOT_FOUND explains, refreshes and closes the drawer", async () => {
    const s = sw({
      rules: [rule('r1')],
      extra: (m) =>
        m.type === 'upsertRule' ? { ok: false, code: 'NOT_FOUND', message: 'gone' } : undefined,
    })
    const host = await renderOptions('#rules-library/all')
    const drawer = await openDrawerFor(host, 'r1.example')
    const edit = await vi.waitFor(() => {
      const b = buttonContaining(drawer, '編輯')
      expect(b?.disabled).toBe(false) // tree loaded on mount
      return b!
    })
    click(edit)
    s.setRules([]) // deleted on another machine meanwhile
    const rulesBefore = s.count('getRules')
    click(await vi.waitFor(() => buttonByText(drawer, '儲存')!))

    await vi.waitFor(() => expect(s.count('upsertRule')).toBe(1))
    const sent = s.last('upsertRule')!
    expect(sent.mode).toBe('edit')
    expect((sent.rule as Rule).id).toBe('r1')
    await vi.waitFor(() => expect(document.body.textContent).toContain('修改沒有儲存'))
    await vi.waitFor(() => expect(s.count('getRules')).toBeGreaterThan(rulesBefore))
    await vi.waitFor(() =>
      expect(document.querySelector('[role="dialog"][aria-label="規則詳細"]')).toBeNull(),
    )
  })

  it('clamps the confidence slider to the per-type cap for a system-learned rule', async () => {
    sw({ rules: [rule('cap', { source: 'ai_confirmed', type: 'domain', confidence: 0.7 })] })
    const host = await renderOptions('#rules-library/all')
    const drawer = await openDrawerFor(host, 'cap.example')
    click(await vi.waitFor(() => buttonContaining(drawer, '編輯')!))
    const slider = await vi.waitFor(() => drawer.querySelector<HTMLInputElement>('input[type="range"]')!)
    expect(slider.max).toBe('0.7')
    expect(drawer.textContent).toContain('上限 0.70')
  })
})

describe('bulk actions', () => {
  async function selectAllAndDisable(host: HTMLElement) {
    const selectAll = await vi.waitFor(() => {
      const el = host.querySelector<HTMLInputElement>('input[aria-label="全選顯示中的規則"]')
      expect(el).not.toBeNull()
      return el!
    })
    click(selectAll)
    click(await vi.waitFor(() => buttonByText(host, '停用')!))
  }

  it('disables the whole selection in ONE bulkRuleAction call and refreshes once', async () => {
    const s = sw({
      rules: [rule('a'), rule('b'), rule('c')],
      extra: (m) =>
        m.type === 'bulkRuleAction' ? { ok: true, data: { changed: 3, missing: 0 } } : undefined,
    })
    const host = await renderOptions('#rules-library/all')
    await vi.waitFor(() => expect(s.count('getRules')).toBeGreaterThan(0))
    const rulesBefore = s.count('getRules')
    await selectAllAndDisable(host)
    await vi.waitFor(() => expect(s.count('bulkRuleAction')).toBe(1))
    const msg = s.last('bulkRuleAction')!
    expect(msg.op).toBe('disable')
    expect([...(msg.ids as string[])].sort()).toEqual(['a', 'b', 'c'])
    expect(s.count('toggleRule')).toBe(0)
    await vi.waitFor(() => expect(s.count('getRules')).toBe(rulesBefore + 1))
  })

  it('falls back to per-rule calls on an older SW — still a single refresh', async () => {
    const s = sw({
      rules: [rule('a'), rule('b'), rule('c')],
      extra: (m) =>
        m.type === 'bulkRuleAction'
          ? { ok: false, code: 'UNKNOWN_REQUEST', message: 'Unknown request type' }
          : m.type === 'toggleRule'
            ? { ok: true }
            : undefined,
    })
    const host = await renderOptions('#rules-library/all')
    await vi.waitFor(() => expect(s.count('getRules')).toBeGreaterThan(0))
    const rulesBefore = s.count('getRules')
    await selectAllAndDisable(host)
    await vi.waitFor(() => expect(s.count('toggleRule')).toBe(3))
    await vi.waitFor(() => expect(s.count('getRules')).toBe(rulesBefore + 1))
    await new Promise((r) => setTimeout(r, 50))
    expect(s.count('getRules')).toBe(rulesBefore + 1)
  })
})

describe('storage usage card', () => {
  it('shows usage + top keys when the SW supports getStorageUsage', async () => {
    sw({
      extra: (m) =>
        m.type === 'getStorageUsage'
          ? {
              ok: true,
              data: {
                bytesInUse: 2 * 1024 * 1024,
                quotaBytes: 10 * 1024 * 1024,
                topKeys: [{ key: 'rules', bytes: 1024 * 1024 }],
              },
            }
          : undefined,
    })
    const host = await renderOptions('#data')
    await vi.waitFor(() => expect(host.textContent).toContain('本機儲存空間'))
    expect(host.textContent).toContain('2.00 MB')
    expect(host.textContent).toContain('(20%)')
    expect(host.textContent).toContain('規則')
  })

  it('renders nothing when the handler is unavailable', async () => {
    const s = sw()
    const host = await renderOptions('#data')
    await vi.waitFor(() => expect(s.count('getStorageUsage')).toBe(1))
    await new Promise((r) => setTimeout(r, 20))
    expect(host.textContent).not.toContain('本機儲存空間')
  })
})
