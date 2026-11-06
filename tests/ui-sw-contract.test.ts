// UI ↔ service-worker contract tests (2026-11-05 audit).
//
// Drive the REAL SW dispatcher with the exact message shapes the UI sends.
// Covers two audit findings:
//   - Rule import always failed (「缺少 payload」): Options sent `json`, the SW
//     read `payload`. Import is also the recovery path for a lost library, so
//     export → import must round-trip.
//   - The owa-fab content script read chrome.storage settings (with the Claude
//     API key) and console.logged them on every Outlook page load. It now only
//     ever receives a visibility boolean from the SW.

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { afterEach, beforeAll, describe, expect, it, vi, type Mock } from 'vitest'
import { CONTENT_SCRIPT_MESSAGE_TYPES, SW_MESSAGE_TYPES } from '@/shared/messages'
import { addRules, listRules, newRule, normalizeSignal } from '@/shared/rules'
import {
  getRuleEvents,
  getRuleTombstones,
  getSettings,
  setFolderCache,
  setRules,
  setSettings,
} from '@/shared/storage'
import type { MailFolderNode, Rule } from '@/shared/types'
import {
  CONTENT_SCRIPT_SENDER,
  EXTENSION_PAGE_SENDER,
  loadServiceWorker,
  type Sender,
  type SwHarness,
} from './helpers/sw-harness'

let sw: SwHarness
beforeAll(async () => {
  sw = await loadServiceWorker()
})

function rule(signal: string) {
  return newRule({
    type: 'domain',
    signal,
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: '03進行中案件/某案',
    confidence: 0.7,
    source: 'user_manual',
  })
}

describe('rule import contract', () => {
  it('export → import round-trips with the exact message shapes the Options page sends', async () => {
    await addRules([rule('a.example'), rule('b.example')])
    const exported = await sw.send({ type: 'exportRules' })
    expect(exported.ok).toBe(true)
    const payload = (exported.data as { payload: string }).payload
    expect(typeof payload).toBe('string')

    await setRules([]) // library lost
    // Shape copied from options/App.tsx handleImportFile.
    const r = await sw.send({ type: 'importRules', payload, strategy: 'merge' })
    expect(r.ok).toBe(true)
    const d = r.data as Record<string, number>
    for (const k of ['added', 'replaced', 'total', 'skipped', 'orphanedCount', 'importedTombstones', 'cappedCount', 'cappedMaxDelta']) {
      expect(typeof d[k]).toBe('number')
    }
    expect((await listRules()).map((x) => x.signal).sort()).toEqual(['a.example', 'b.example'])
  })

  it('the Options page really sends `payload` (guards the UI side of the contract)', () => {
    const src = readFileSync(resolve(__dirname, '../src/options/App.tsx'), 'utf8')
    const call = src.match(/type:\s*'importRules'[^}]*\}/)?.[0] ?? ''
    expect(call).toContain('payload')
  })

  it('legacy `json` field is still accepted', async () => {
    await addRules([rule('c.example')])
    const payload = ((await sw.send({ type: 'exportRules' })).data as { payload: string }).payload
    await setRules([])
    const r = await sw.send({ type: 'importRules', json: payload, strategy: 'merge' })
    expect(r.ok).toBe(true)
    expect(await listRules()).toHaveLength(1)
  })

  // Folders on this machine: the imported rule's 某案 (when `withTarget`) and
  // the local rule's 別案.
  const OTHER = 'AAMkAD-other-folder-id-000000'
  const folders = (withTarget: boolean): MailFolderNode[] => [
    {
      id: 'AAMkAD-root-folder-id-00000000',
      displayName: '03進行中案件',
      path: '03進行中案件',
      children: [
        ...(withTarget
          ? [{ id: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3', displayName: '某案', path: '03進行中案件/某案', children: [] }]
          : []),
        { id: OTHER, displayName: '別案', path: '03進行中案件/別案', children: [] },
      ],
    },
  ]
  async function importOver(withTarget: boolean) {
    const mine = rule('d.example')
    await setRules([mine])
    const payload = ((await sw.send({ type: 'exportRules' })).data as { payload: string }).payload
    const scan = { ...rule('d.example'), id: 'scan', source: 'auto_scan' as const, targetFolderId: OTHER, targetFolderPath: '03進行中案件/別案' }
    await setRules([scan])
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: folders(withTarget),
      // current-format cache; one without wellKnown is refetched as legacy (round 3)
      wellKnown: { roots: [], complete: true },
    })
    const r = await sw.send({ type: 'importRules', payload, strategy: 'merge' })
    expect(r.ok).toBe(true)
    return { r, mine }
  }

  it('R5: an imported hand-written rule takes over from a local auto one — reported in the counts and the history', async () => {
    const { r, mine } = await importOver(true)
    expect(r.data).toMatchObject({ added: 1, skipped: 0, conflictSkipped: 0, replacedLocal: 1, supersededLocal: 1 })
    const rules = await listRules()
    expect(rules.find((x) => x.id === mine.id)?.enabled).toBe(true)
    expect(rules.find((x) => x.id === 'scan')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded' })
    const events = await getRuleEvents(10)
    expect(events.some((e) => e.kind === 'toggle' && e.ruleId === 'scan' && e.enabled === false)).toBe(true)
  })

  // Round 3: the takeover was decided before the reconcile, so a working
  // local rule was switched off for an imported rule whose folder does not
  // exist here — and nothing routed that signal any more.
  it('R5: no takeover by an imported rule whose folder does not exist here', async () => {
    const { r, mine } = await importOver(false)
    expect(r.data).toMatchObject({ added: 1, supersededLocal: 0, orphanedCount: 1 })
    const rules = await listRules()
    expect(rules.find((x) => x.id === 'scan')).toMatchObject({ enabled: true }) // was: superseded
    expect(rules.find((x) => x.id === 'scan')!.autoDisabledReason).toBeUndefined()
    expect(rules.find((x) => x.id === mine.id)).toMatchObject({ enabled: true, orphaned: true })
  })
})

describe('OWA floating button never sees the API key', () => {
  it('getOwaFabVisibility returns only the boolean', async () => {
    await setSettings({ claudeApiKey: 'sk-ant-api03-SECRET-TEST', showOwaFab: false })
    const r = await sw.send({ type: 'getOwaFabVisibility' })
    expect(r).toEqual({ ok: true, data: { show: false } })
    expect(JSON.stringify(r)).not.toContain('sk-ant')
  })

  it('defaults to shown', async () => {
    const r = await sw.send({ type: 'getOwaFabVisibility' })
    expect(r).toEqual({ ok: true, data: { show: true } })
  })

  it('the SW pushes a one-boolean message to OWA tabs only when the setting flips', async () => {
    const calls = (chrome.storage.onChanged.addListener as unknown as { mock: { calls: unknown[][] } }).mock
      .calls
    const listener = calls
      .map((c) => c[0] as (changes: unknown, area: string) => void)
      .find((fn) => fn.toString().includes('owaFabShown'))
    expect(listener).toBeDefined()

    sw.tabsQuery.mockResolvedValue([{ id: 7 } as chrome.tabs.Tab])
    sw.tabsSendMessage.mockClear()
    const key = 'sk-ant-api03-SECRET-TEST'
    listener!(
      {
        settings: {
          oldValue: { claudeApiKey: key, showOwaFab: true },
          newValue: { claudeApiKey: key, showOwaFab: true, lastSyncAt: 'x' },
        },
      },
      'local',
    )
    await new Promise((r) => setTimeout(r, 0))
    expect(sw.tabsSendMessage).not.toHaveBeenCalled() // no flip, no message

    listener!(
      {
        settings: {
          oldValue: { claudeApiKey: key, showOwaFab: true },
          newValue: { claudeApiKey: key, showOwaFab: false },
        },
      },
      'local',
    )
    await vi_waitFor(() => sw.tabsSendMessage.mock.calls.length > 0)
    expect(sw.tabsSendMessage).toHaveBeenCalledWith(7, { type: 'owaFabVisibility', show: false })
    expect(JSON.stringify(sw.tabsSendMessage.mock.calls)).not.toContain('sk-ant')
  })

  it('content scripts never touch chrome.storage (where the API key lives)', () => {
    const dir = resolve(__dirname, '../src/content')
    for (const f of readdirSync(dir).filter((n) => n.endsWith('.ts'))) {
      const code = readFileSync(resolve(dir, f), 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/.*$/gm, '')
      expect(code, f).not.toMatch(/chrome\.storage/)
    }
  })
})

async function vi_waitFor(cond: () => boolean, timeoutMs = 1000): Promise<void> {
  const start = Date.now()
  while (!cond()) {
    if (Date.now() - start > timeoutMs) throw new Error('waitFor timed out')
    await new Promise((r) => setTimeout(r, 5))
  }
}

// ---------------------------------------------------------------------------
// 2026-11-05 SW-core batch
// ---------------------------------------------------------------------------

function ruleAt(signal: string, folder: string, over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal,
      // ≥ 20 chars, or getRules' auto-heal blanks the id.
      targetFolderId: `AAMkAD-folder-id-${folder}-padding-padding`,
      targetFolderPath: `03進行中案件/${folder}`,
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    ...over,
  }
}

function leaf(id: string, path: string): MailFolderNode {
  return { id, displayName: path.split('/').pop()!, path, children: [] }
}

/** chrome.storage.local.set calls that wrote `key` since the last mockClear. */
function writesOf(key: string): number {
  return vi
    .mocked(chrome.storage.local.set)
    .mock.calls.filter((c) => key in (c[0] as Record<string, unknown>)).length
}

describe('dispatcher only serves extension pages (sender allowlist)', () => {
  it('a content-script sender may only ask for the FAB visibility', async () => {
    await addRules([rule('keep.example')])
    for (const type of ['wipeAllRules', 'exportDiagnostic', 'setSettings', 'getStatus', 'importRules']) {
      const r = await sw.sendFrom(CONTENT_SCRIPT_SENDER, { type, patch: { batchSize: 1 } })
      expect(r, type).toMatchObject({ ok: false, code: 'FORBIDDEN' })
    }
    expect(await listRules()).toHaveLength(1) // wipe refused
    expect((await getSettings()).batchSize).toBe(50) // setSettings refused
    expect(await sw.sendFrom(CONTENT_SCRIPT_SENDER, { type: 'getOwaFabVisibility' })).toEqual({
      ok: true,
      data: { show: true },
    })
  })

  it('refuses senders with no URL, another extension, or a look-alike URL', async () => {
    const senders: Sender[] = [
      { id: 'test-extension' },
      { id: 'other-extension', url: 'chrome-extension://other-extension/src/options/index.html' },
      { id: 'test-extension', url: 'chrome-extension://other-extension/src/options/index.html' },
      { id: 'test-extension', url: 'https://evil.example/chrome-extension://test-extension/' },
    ]
    for (const s of senders) {
      expect(await sw.sendFrom(s, { type: 'exportDiagnostic' }), JSON.stringify(s)).toMatchObject({
        ok: false,
        code: 'FORBIDDEN',
      })
    }
    // malformed payloads from a content script are refused, not parsed
    expect(await sw.sendFrom(CONTENT_SCRIPT_SENDER, null)).toMatchObject({ code: 'FORBIDDEN' })
  })

  it('every message a content script sends is on the content-script allowlist', () => {
    // Otherwise it would pass extension-page-sender tests and be FORBIDDEN in
    // production, where its sender.url is the Outlook page.
    const dir = resolve(__dirname, '../src/content')
    const sent = new Set<string>()
    for (const f of readdirSync(dir).filter((n) => n.endsWith('.ts'))) {
      const code = readFileSync(resolve(dir, f), 'utf8')
      for (const m of code.matchAll(/runtime\.sendMessage\(\s*\{\s*type:\s*'([A-Za-z]+)'/g)) sent.add(m[1]!)
      // A non-literal send would slip past the scan — keep them literal.
      expect(code, f).not.toMatch(/runtime\.sendMessage\(\s*[^{\s]/)
    }
    expect(sent.size).toBeGreaterThan(0)
    expect([...sent].filter((t) => !(CONTENT_SCRIPT_MESSAGE_TYPES as readonly string[]).includes(t))).toEqual([])
  })

  it('the in-OWA iframe panel (an extension page inside an Outlook tab) keeps full access', async () => {
    const panel: Sender = {
      id: 'test-extension',
      url: 'chrome-extension://test-extension/src/popup/index.html',
      origin: 'chrome-extension://test-extension',
      tab: { id: 7 } as chrome.tabs.Tab,
      frameId: 3,
    }
    expect((await sw.sendFrom(panel, { type: 'getRules' })).ok).toBe(true)
  })
})

describe('message contract typing', () => {
  it('unknown or missing types answer UNKNOWN_TYPE', async () => {
    // A type the stale messages.ts PopupRequest advertised but nothing served.
    expect(await sw.send({ type: 'classify' })).toMatchObject({ ok: false, code: 'UNKNOWN_TYPE' })
    expect(await sw.sendFrom(EXTENSION_PAGE_SENDER, { nope: 1 })).toMatchObject({ code: 'UNKNOWN_TYPE' })
    expect(await sw.sendFrom(EXTENSION_PAGE_SENDER, undefined)).toMatchObject({ code: 'UNKNOWN_TYPE' })
  })

  it('every message type the UI sources send is one the SW serves', () => {
    // popup/App.tsx and options/App.tsx keep their own UNTYPED send() copies,
    // so the compiler can't see their literals — scan them.
    const roots = ['../src/popup', '../src/options', '../src/content'].map((p) => resolve(__dirname, p))
    const files: string[] = []
    const walk = (dir: string) => {
      for (const name of readdirSync(dir)) {
        const full = join(dir, name)
        if (statSync(full).isDirectory()) walk(full)
        else if (/\.tsx?$/.test(name)) files.push(full)
      }
    }
    roots.forEach(walk)
    // send(…) / send<T>(…) / chrome.runtime.sendMessage(…) whose first
    // argument is an object literal starting with `type: '…'`.
    const call = /\b(?:send|sendMessage)\s*(?:<(?:[^()]|\((?:[^()]|\([^()]*\))*\))*?>)?\(\s*\{\s*type:\s*'([A-Za-z]+)'/g
    const sent = new Set<string>()
    for (const f of files) {
      for (const m of readFileSync(f, 'utf8').matchAll(call)) sent.add(m[1]!)
    }
    expect(sent.size).toBeGreaterThan(40) // the scan really found the call sites
    const served = new Set<string>(SW_MESSAGE_TYPES)
    expect([...sent].filter((t) => !served.has(t))).toEqual([])
  })
})

describe('toggleRule', () => {
  it('a missing id is NOT_FOUND (was ok:true with rule undefined)', async () => {
    expect(await sw.send({ type: 'toggleRule', ruleId: 'gone', enabled: false })).toMatchObject({
      ok: false,
      code: 'NOT_FOUND',
    })
    const r = ruleAt('a.example', 'A')
    await setRules([r])
    expect(await sw.send({ type: 'toggleRule', ruleId: r.id, enabled: false })).toMatchObject({
      ok: true,
      data: { rule: { id: r.id, enabled: false } },
    })
  })
})

describe('setSettings handler', () => {
  it('drops closed-case prefixes that are not excluded, even in a closed-only patch', async () => {
    await sw.send({ type: 'setSettings', patch: { excludeFolderPrefixes: ['05已完成案件'] } })
    const r = await sw.send({
      type: 'setSettings',
      patch: { closedCasePrefixes: ['05已完成案件', '99其他'] },
    })
    expect(r).toMatchObject({ ok: true, data: { settings: { closedCasePrefixes: ['05已完成案件'] } } })
  })
})

describe('refreshFolderActivity single-flight', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('concurrent refreshes (toolbar popup + in-OWA panel) share ONE scan', async () => {
    sw.tabsQuery.mockResolvedValue([])
    await setSettings({ primaryRootPath: '案件' })
    await setFolderCache({
      updatedAt: new Date().toISOString(),
      tree: [
        {
          ...leaf('AAMkAD-root-folder-id-padding', '案件'),
          children: [
            leaf('AAMkAD-leaf-a-folder-id-padding', '案件/甲'),
            leaf('AAMkAD-leaf-b-folder-id-padding', '案件/乙'),
          ],
        },
      ],
    })
    await chrome.storage.session.set({
      cachedToken: { secret: 'test-token', expiresOn: Date.now() / 1000 + 3600 },
    })
    let release!: () => void
    const gate = new Promise<void>((r) => {
      release = r
    })
    const fetchMock = vi.fn(async () => {
      await gate
      return {
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              Id: 'm1',
              Subject: 'latest',
              ReceivedDateTime: '2026-11-01T00:00:00Z',
              From: { EmailAddress: { Address: 'x@client.example' } },
            },
          ],
        }),
      }
    })
    vi.stubGlobal('fetch', fetchMock)

    const fromPopup = sw.send({ type: 'refreshFolderActivity' })
    const fromPanel = sw.send({ type: 'refreshFolderActivity' })
    await vi_waitFor(() => fetchMock.mock.calls.length >= 1)
    release()
    const [a, b] = await Promise.all([fromPopup, fromPanel])
    expect(a).toEqual(b)
    expect(a).toMatchObject({ ok: true, data: { scanned: 2, matched: 2, errored: 0 } })
    expect(fetchMock).toHaveBeenCalledTimes(2) // one scan × 2 folders, not two scans

    // The slot clears afterwards: a later refresh runs a fresh scan.
    expect((await sw.send({ type: 'refreshFolderActivity' })).ok).toBe(true)
    expect(fetchMock).toHaveBeenCalledTimes(4)
  })
})

describe('getStatus count cache', () => {
  const tree3 = () => [
    leaf('AAMkAD-x-folder-id-padding-pad', 'X'),
    leaf('AAMkAD-y-folder-id-padding-pad', 'Y'),
    leaf('AAMkAD-z-folder-id-padding-pad', 'Z'),
  ]

  it('serves counts without re-reading rules / folder tree until they change', async () => {
    await setRules([ruleAt('a.example', 'A'), ruleAt('b.example', 'B')])
    await setFolderCache({ updatedAt: new Date().toISOString(), tree: tree3() })
    await sw.emitStorageChange({ rules: {}, folderCache: {} })
    expect((await sw.send({ type: 'getStatus' })).data).toMatchObject({
      rulesCount: 2,
      folderCacheCount: 3,
    })

    const get = vi.mocked(chrome.storage.local.get)
    get.mockClear()
    expect((await sw.send({ type: 'getStatus' })).data).toMatchObject({
      rulesCount: 2,
      folderCacheCount: 3,
    })
    const keysRead = get.mock.calls.map((c) => c[0])
    expect(keysRead).not.toContain('rules')
    expect(keysRead).not.toContain('folderCache')

    await addRules([ruleAt('c.example', 'C')])
    await sw.emitStorageChange({ rules: {} })
    expect((await sw.send({ type: 'getStatus' })).data).toMatchObject({ rulesCount: 3 })

    await setFolderCache({ updatedAt: new Date().toISOString(), tree: tree3().slice(0, 1) })
    await sw.emitStorageChange({ folderCache: {} })
    expect((await sw.send({ type: 'getStatus' })).data).toMatchObject({ folderCacheCount: 1 })
  })

  it('a write that lands mid-recompute is not cached as a stale count', async () => {
    await setRules([ruleAt('a.example', 'A')])
    await sw.emitStorageChange({ rules: {} })
    const get = chrome.storage.local.get as unknown as Mock<
      (keys?: unknown) => Promise<Record<string, unknown>>
    >
    const realGet = get.getMockImplementation()!
    let release!: () => void
    const gate = new Promise<void>((r) => {
      release = r
    })
    let gated = false
    get.mockImplementation(async (keys) => {
      const v = await realGet(keys)
      if (keys === 'rules' && !gated) {
        gated = true
        await gate
      }
      return v
    })
    try {
      const pending = sw.send({ type: 'getStatus' }) // reads 1 rule, then stalls
      await vi_waitFor(() => gated)
      await addRules([ruleAt('b.example', 'B')]) // write + its onChanged land meanwhile
      await sw.emitStorageChange({ rules: {} })
      release()
      expect((await pending).data).toMatchObject({ rulesCount: 1 })
    } finally {
      get.mockImplementation(realGet)
    }
    expect((await sw.send({ type: 'getStatus' })).data).toMatchObject({ rulesCount: 2 })
  })
})

describe('bulk rule operations: one rewrite + one history append each', () => {
  it('bulkRuleAction disable / enable / delete', async () => {
    const a = ruleAt('a.example', 'A')
    const b = ruleAt('b.example', 'B')
    const c = ruleAt('c.example', 'C', {
      enabled: false,
      autoDisabledAt: '2026-10-01T00:00:00.000Z',
      autoDisabledReason: 'high-error-rate',
    })
    await setRules([a, b, c])
    const set = vi.mocked(chrome.storage.local.set)

    set.mockClear()
    const off = await sw.send({ type: 'bulkRuleAction', op: 'disable', ids: [a.id, b.id, c.id, 'ghost', a.id] })
    // c was already disabled (no-op, no event); 'ghost' isn't in the library.
    expect(off).toEqual({ ok: true, data: { changed: 2, missing: 1 } })
    expect(writesOf('rules')).toBe(1)
    expect(writesOf('ruleHistory')).toBe(1)
    expect((await listRules()).every((r) => !r.enabled)).toBe(true)
    expect((await getRuleEvents()).filter((e) => e.kind === 'toggle')).toHaveLength(2)

    set.mockClear()
    expect(await sw.send({ type: 'bulkRuleAction', op: 'enable', ids: [c.id] })).toEqual({
      ok: true,
      data: { changed: 1, missing: 0 },
    })
    const cNow = (await listRules()).find((r) => r.id === c.id)!
    expect(cNow).toMatchObject({ enabled: true })
    // Manual enable clears auto-disable provenance, exactly like toggleRule.
    expect(cNow.autoDisabledAt).toBeUndefined()
    expect(cNow.autoDisabledReason).toBeUndefined()

    set.mockClear()
    expect(await sw.send({ type: 'bulkRuleAction', op: 'delete', ids: [a.id, b.id] })).toEqual({
      ok: true,
      data: { changed: 2, missing: 0 },
    })
    expect(writesOf('rules')).toBe(1)
    expect(writesOf('ruleHistory')).toBe(1)
    expect((await listRules()).map((r) => r.id)).toEqual([c.id])
    // Tombstoned like deleteRule, so AI / scan won't re-learn them.
    const tombKeys = (await getRuleTombstones()).map((t) => `${t.type}::${t.signalNorm}::${t.targetFolderPath}`)
    for (const r of [a, b]) {
      expect(tombKeys).toContain(`${r.type}::${normalizeSignal(r.type, r.signal)}::${r.targetFolderPath}`)
    }
    expect((await getRuleEvents()).filter((e) => e.kind === 'delete').map((e) => e.ruleId).sort()).toEqual(
      [a.id, b.id].sort(),
    )

    set.mockClear()
    expect(await sw.send({ type: 'bulkRuleAction', op: 'delete', ids: [] })).toEqual({
      ok: true,
      data: { changed: 0, missing: 0 },
    })
    expect(writesOf('rules')).toBe(0) // no-op → no rewrite (and no sync push)

    expect(await sw.send({ type: 'bulkRuleAction', op: 'nuke', ids: [c.id] })).toMatchObject({
      code: 'INVALID_INPUT',
    })
    expect(await sw.send({ type: 'bulkRuleAction', op: 'delete', ids: [1] })).toMatchObject({
      code: 'INVALID_INPUT',
    })
  })

  it('clearDisabledRules deletes every disabled rule in one rewrite, tombstoned', async () => {
    const keep = ruleAt('keep.example', 'K')
    const d1 = ruleAt('d1.example', 'D1', { enabled: false })
    const d2 = ruleAt('d2.example', 'D2', { enabled: false, source: 'user_manual' })
    await setRules([keep, d1, d2])
    vi.mocked(chrome.storage.local.set).mockClear()
    expect(await sw.send({ type: 'clearDisabledRules' })).toEqual({ ok: true, data: { deleted: 2 } })
    expect(writesOf('rules')).toBe(1)
    expect(writesOf('ruleHistory')).toBe(1)
    expect((await listRules()).map((r) => r.id)).toEqual([keep.id])
    expect(await getRuleTombstones()).toHaveLength(2)
  })

  it('resolveConflict keep_highest / keep_one / disable_all', async () => {
    const lo = ruleAt('dup.example', 'A', { confidence: 0.5 })
    const hi = ruleAt('dup.example', 'B', { confidence: 0.7 })
    const mid = ruleAt('dup.example', 'C', { confidence: 0.6 })
    await setRules([lo, hi, mid])
    vi.mocked(chrome.storage.local.set).mockClear()
    expect(
      await sw.send({ type: 'resolveConflict', ruleIds: [lo.id, hi.id, mid.id], strategy: 'keep_highest' }),
    ).toEqual({ ok: true, data: { disabled: 0, deleted: 2, kept: hi.id } })
    expect(writesOf('rules')).toBe(1)
    expect(writesOf('ruleHistory')).toBe(1)
    expect((await listRules()).map((r) => r.id)).toEqual([hi.id])
    expect(await getRuleTombstones()).toHaveLength(2)

    await setRules([lo, hi, mid])
    expect(
      await sw.send({ type: 'resolveConflict', ruleIds: [lo.id, hi.id], strategy: 'keep_one', keepId: mid.id }),
    ).toMatchObject({ ok: false, code: 'INVALID_INPUT' })
    expect(
      await sw.send({ type: 'resolveConflict', ruleIds: [lo.id, hi.id], strategy: 'keep_one', keepId: lo.id }),
    ).toEqual({ ok: true, data: { disabled: 0, deleted: 1, kept: lo.id } })

    await setRules([lo, { ...hi, enabled: false }, mid])
    vi.mocked(chrome.storage.local.set).mockClear()
    expect(
      await sw.send({ type: 'resolveConflict', ruleIds: [lo.id, hi.id, mid.id], strategy: 'disable_all' }),
    ).toEqual({ ok: true, data: { disabled: 2, deleted: 0, kept: null } })
    expect(writesOf('rules')).toBe(1)
    expect((await listRules()).every((r) => !r.enabled)).toBe(true)

    expect(
      await sw.send({ type: 'resolveConflict', ruleIds: ['ghost'], strategy: 'disable_all' }),
    ).toMatchObject({ ok: false, code: 'NOT_FOUND' })
  })

  it('splitConflictToCompound creates compounds + disables originals in one rewrite', async () => {
    const a = ruleAt('dup.example', 'A')
    const b = ruleAt('dup.example', 'B')
    await setRules([a, b])
    vi.mocked(chrome.storage.local.set).mockClear()
    const r = await sw.send({
      type: 'splitConflictToCompound',
      items: [
        { ruleId: a.id, keywords: ['甲案', ' 乙案 '] },
        { ruleId: b.id, keywords: [] },
      ],
    })
    expect(r).toEqual({ ok: true, data: { created: 2, disabled: 1, keptAsFallback: 1 } })
    expect(writesOf('rules')).toBe(1)
    expect(writesOf('ruleHistory')).toBe(1)
    const after = await listRules()
    expect(after.find((x) => x.id === a.id)!.enabled).toBe(false)
    expect(after.find((x) => x.id === b.id)!.enabled).toBe(true)
    const compounds = after.filter((x) => x.type === 'compound')
    expect(compounds).toHaveLength(2)
    for (const cpd of compounds) {
      expect(cpd).toMatchObject({ source: 'user_manual', targetFolderPath: a.targetFolderPath, enabled: true })
    }
    expect(compounds.map((cpd) => JSON.parse(cpd.signal).conditions[1].value).sort()).toEqual(['乙案', '甲案'])
    const kinds = (await getRuleEvents()).map((e) => e.kind).sort()
    expect(kinds).toEqual(['create', 'create', 'toggle'])

    expect(
      await sw.send({
        type: 'splitConflictToCompound',
        items: [
          { ruleId: a.id, keywords: [] },
          { ruleId: b.id, keywords: [] },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'TOO_MANY_FALLBACKS' })
  })
})

describe('getStorageUsage handler', () => {
  it('returns runtime quota + biggest keys (names only); null quota under unlimitedStorage', async () => {
    await setSettings({ claudeApiKey: 'sk-ant-api03-SECRET-TEST' })
    await setRules([ruleAt('a.example', 'A'), ruleAt('b.example', 'B')])
    const r = await sw.send({ type: 'getStorageUsage' })
    expect(r.ok).toBe(true)
    const d = r.data as {
      bytesInUse: number
      quotaBytes: number | null
      topKeys: Array<{ key: string; bytes: number }>
    }
    expect(Object.keys(d).sort()).toEqual(['bytesInUse', 'quotaBytes', 'topKeys'])
    expect(d.quotaBytes).toBe(chrome.storage.local.QUOTA_BYTES)
    expect(d.bytesInUse).toBeGreaterThan(0)
    expect(d.topKeys.map((k) => k.key)).toEqual(expect.arrayContaining(['rules', 'settings']))
    const sizes = d.topKeys.map((k) => k.bytes)
    expect([...sizes].sort((x, y) => y - x)).toEqual(sizes)
    expect(JSON.stringify(r)).not.toContain('sk-ant')

    const runtime = chrome.runtime as unknown as { getManifest: () => unknown }
    const original = runtime.getManifest
    runtime.getManifest = () => ({ version: '0.0.0-test', permissions: ['storage', 'unlimitedStorage'] })
    try {
      expect((await sw.send({ type: 'getStorageUsage' })).data).toMatchObject({ quotaBytes: null })
    } finally {
      runtime.getManifest = original
    }
  })
})

describe('toolbar icon → Outlook, without browser-wide tab listeners', () => {
  it('registers no tabs.onCreated / onUpdated / onRemoved listeners and re-asserts the popup', () => {
    const tabs = chrome.tabs as unknown as Record<string, { listeners: unknown[] }>
    expect(tabs.onCreated!.listeners).toHaveLength(0)
    expect(tabs.onUpdated!.listeners).toHaveLength(0)
    expect(tabs.onRemoved!.listeners).toHaveLength(0)
    expect(chrome.action.setPopup).toHaveBeenCalledWith({ popup: 'src/popup/index.html' })
  })

  it('openOwa opens Outlook when no tab exists and focuses one when it does', async () => {
    sw.tabsQuery.mockResolvedValue([])
    expect(await sw.send({ type: 'openOwa' })).toEqual({ ok: true, data: { action: 'created' } })
    expect(chrome.tabs.create).toHaveBeenCalledWith({ url: 'https://outlook.office.com/mail/', active: true })

    sw.tabsQuery.mockResolvedValue([{ id: 3, windowId: 9 } as chrome.tabs.Tab])
    expect(await sw.send({ type: 'openOwa' })).toEqual({ ok: true, data: { action: 'focused' } })
    expect(chrome.tabs.update).toHaveBeenCalledWith(3, { active: true })
    expect(chrome.windows.update).toHaveBeenCalledWith(9, { focused: true })
    sw.tabsQuery.mockResolvedValue([])
  })

  it('the popup redirects only as the toolbar popup, never inside the OWA iframe', () => {
    const src = readFileSync(resolve(__dirname, '../src/popup/App.tsx'), 'utf8')
    const at = src.indexOf("type: 'openOwa'")
    expect(at).toBeGreaterThan(0)
    const effect = src.slice(Math.max(0, at - 900), at + 120)
    expect(effect).toContain('status.owaConnected')
    expect(effect).toContain('window.top !== window')
    expect(effect).toContain('window.close()')
  })
})
