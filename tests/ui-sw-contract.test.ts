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

import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { beforeAll, describe, expect, it } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import { setRules, setSettings } from '@/shared/storage'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

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
