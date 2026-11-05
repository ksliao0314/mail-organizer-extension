import { describe, expect, it } from 'vitest'
import { restoreBackup, SYNC_BACKUPS_KEY } from '@/background/sync-engine'
import {
  addToSkipHistory,
  clearSkipHistory,
  getRules,
  getSettings,
  getSkipHistoryCount,
  setRules,
  setSettings,
} from '@/shared/storage'
import { DEFAULT_CLAUDE_MODEL } from '@/shared/classifier'
import { DEFAULT_SETTINGS, type Rule, type Settings } from '@/shared/types'

function rule(over: Partial<Rule>): Rule {
  return {
    id: over.id ?? 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'a'.repeat(50),
    targetFolderPath: 'X/Y',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: new Date().toISOString(),
    source: 'user_manual',
    ...over,
  }
}

describe('getRules auto-heal', () => {
  it('strips pending: prefix on read and writes back', async () => {
    const dirty = rule({ targetFolderId: 'pending:emailX' + 'a'.repeat(50) })
    await setRules([dirty])
    const after = await getRules()
    expect(after[0]!.targetFolderId).toBe('')
    // Subsequent read should NOT have to clean again (no infinite write loop)
    const after2 = await getRules()
    expect(after2[0]!.targetFolderId).toBe('')
  })

  it('strips PLACEHOLDER ids', async () => {
    const dirty = rule({ targetFolderId: 'PLACEHOLDER_FOLDER_ID' })
    await setRules([dirty])
    const after = await getRules()
    expect(after[0]!.targetFolderId).toBe('')
  })

  it('strips too-short ids', async () => {
    const dirty = rule({ targetFolderId: 'short' })
    await setRules([dirty])
    const after = await getRules()
    expect(after[0]!.targetFolderId).toBe('')
  })

  it('leaves valid ids untouched', async () => {
    const valid = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
    await setRules([rule({ targetFolderId: valid })])
    const after = await getRules()
    expect(after[0]!.targetFolderId).toBe(valid)
  })
})

describe('getSettings sanitize', () => {
  it('drops legacy unknown keys (e.g. confidenceThreshold)', async () => {
    // Simulate old data with legacy field
    await chrome.storage.local.set({
      settings: {
        claudeApiKey: 'sk-ant-test',
        claudeModel: 'claude-sonnet-4-6',
        batchSize: 50,
        excludeFolderPrefixes: ['05已完成案件'],
        aiConfidenceThreshold: 0.5,
        skipFlagged: true,
        confidenceThreshold: 0.85, // legacy
        someJunkField: 'whatever',
      },
    })
    const cleaned = await getSettings()
    expect(cleaned).not.toHaveProperty('confidenceThreshold')
    expect(cleaned).not.toHaveProperty('someJunkField')
    expect(cleaned.claudeApiKey).toBe('sk-ant-test')
    expect(cleaned.aiConfidenceThreshold).toBe(0.5)
    expect(cleaned.skipFlagged).toBe(true)
    // Should have persisted the cleaned version
    const raw = (await chrome.storage.local.get('settings'))['settings']
    expect(raw).not.toHaveProperty('confidenceThreshold')
  })

  it('skipFlagged defaults to true when not set', async () => {
    await chrome.storage.local.set({ settings: { claudeApiKey: 'x' } })
    const cleaned = await getSettings()
    expect(cleaned.skipFlagged).toBe(true)
  })

  it('skipFlagged respects user override to false', async () => {
    await chrome.storage.local.set({ settings: { skipFlagged: false } })
    const cleaned = await getSettings()
    expect(cleaned.skipFlagged).toBe(false)
  })

  it('handles malformed types by falling back to defaults', async () => {
    await chrome.storage.local.set({
      settings: {
        claudeModel: 123, // wrong type
        batchSize: 'oops', // wrong type
      },
    })
    const cleaned = await getSettings()
    expect(cleaned.claudeModel).toBe(DEFAULT_SETTINGS.claudeModel)
    expect(cleaned.batchSize).toBe(50)
  })

  // 2026-11-05: claudeModel syncs to every machine; a retired / junk id used
  // to fail every AI email with a raw 404. Unknown ids map to the default,
  // known ids (including the legacy 4.x ones) are kept — no migration.
  it('maps an unknown / retired claudeModel to the default', async () => {
    await chrome.storage.local.set({ settings: { claudeModel: 'claude-3-opus-20240229' } })
    expect((await getSettings()).claudeModel).toBe(DEFAULT_SETTINGS.claudeModel)
  })

  it('keeps every known model id as stored (no migration of legacy picks)', async () => {
    for (const id of ['claude-sonnet-4-6', 'claude-haiku-4-5-20251001', 'claude-opus-4-7', 'claude-opus-5-5']) {
      await chrome.storage.local.set({ settings: { claudeModel: id } })
      expect((await getSettings()).claudeModel).toBe(id)
    }
  })

  it('new installs default to Sonnet 5.5', async () => {
    expect(DEFAULT_SETTINGS.claudeModel).toBe('claude-sonnet-5-5')
    expect(DEFAULT_SETTINGS.claudeModel).toBe(DEFAULT_CLAUDE_MODEL)
    expect((await getSettings()).claudeModel).toBe('claude-sonnet-5-5')
  })

  it('persists patch via setSettings cleanly', async () => {
    await setSettings({ batchSize: 100 })
    const after = await getSettings()
    expect(after.batchSize).toBe(100)
  })

  it('harvestEnabled defaults to false when unset', async () => {
    await chrome.storage.local.set({ settings: { claudeApiKey: 'x' } })
    const cleaned = await getSettings()
    expect(cleaned.harvestEnabled).toBe(false)
  })

  it('closedCasePrefixes round-trips through setSettings (3-layer whitelist)', async () => {
    // (2026-11-05) closedCasePrefixes must be a subset of excludeFolderPrefixes
    // on every write, so the prefix has to be excluded first.
    await setSettings({ excludeFolderPrefixes: ['05已完成案件'], closedCasePrefixes: ['05已完成案件'] })
    const after = await getSettings()
    expect(after.closedCasePrefixes).toEqual(['05已完成案件'])
    // Re-sanitize on next read keeps it.
    expect((await getSettings()).closedCasePrefixes).toEqual(['05已完成案件'])
  })

  it('closedCasePrefixes defaults to [] when unset', async () => {
    await chrome.storage.local.set({ settings: { claudeApiKey: 'x' } })
    expect((await getSettings()).closedCasePrefixes).toEqual([])
  })

  it('harvestEnabled round-trips through setSettings (was silently stripped)', async () => {
    // Regression: harvestEnabled was in the Settings type + DEFAULT_SETTINGS
    // but absent from ALLOWED_SETTINGS_KEYS and sanitizeSettings, so the
    // Options toggle wrote true and getSettings dropped it back to false.
    await setSettings({ harvestEnabled: true })
    const after = await getSettings()
    expect(after.harvestEnabled).toBe(true)
    // And it must survive a re-sanitize on the next read.
    const again = await getSettings()
    expect(again.harvestEnabled).toBe(true)
  })
})

// (2026-11-05) storage.setSettings used to merge `{ ...stored, ...patch }`
// verbatim. Only the SW 'setSettings' handler clamped values, but sync pull and
// restoreBackup call storage.setSettings directly — a cloud / backup payload
// with batchSize 0 (every batch empty) or aiConfidenceThreshold 7 (every AI
// item held back) was stored AND read back as-is.
describe('setSettings sanitizes the MERGED result (every writer)', () => {
  async function rawStored(): Promise<Record<string, unknown>> {
    return ((await chrome.storage.local.get('settings'))['settings'] ?? {}) as Record<string, unknown>
  }

  it('clamps batchSize to 1–200 and aiConfidenceThreshold to 0–1, stored and read back', async () => {
    await setSettings({ batchSize: 0, aiConfidenceThreshold: 7 })
    expect(await rawStored()).toMatchObject({ batchSize: 1, aiConfidenceThreshold: 1 })
    expect(await getSettings()).toMatchObject({ batchSize: 1, aiConfidenceThreshold: 1 })

    await setSettings({ batchSize: 999, aiConfidenceThreshold: -0.5 })
    expect(await getSettings()).toMatchObject({ batchSize: 200, aiConfidenceThreshold: 0 })

    await setSettings({ batchSize: 12.7 })
    expect((await getSettings()).batchSize).toBe(12)
  })

  it('legacy out-of-range values already in storage are clamped on read', async () => {
    await chrome.storage.local.set({ settings: { batchSize: 0, aiConfidenceThreshold: 7 } })
    expect(await getSettings()).toMatchObject({ batchSize: 1, aiConfidenceThreshold: 1 })
  })

  it('trims, dedupes and drops empty string-list entries', async () => {
    await setSettings({
      // An empty exclude prefix used to exclude EVERY folder (startsWith('')).
      excludeFolderPrefixes: [' 05已完成案件 ', '05已完成案件', '', '   ', '06封存/'],
      recentActivityIncludePrefixes: ['案件/', ' 案件/ ', ''],
      recentActivityIncludeLeafNames: ['文章審閱', '文章審閱 ', ''],
      internalSubjectCategories: [' 工時', '工時', ''],
      internalDomains: ['@Example.COM ', 'example.com', ' @ ', ''],
    })
    const s = await getSettings()
    expect(s.excludeFolderPrefixes).toEqual(['05已完成案件', '06封存/'])
    expect(s.recentActivityIncludePrefixes).toEqual(['案件/'])
    expect(s.recentActivityIncludeLeafNames).toEqual(['文章審閱'])
    expect(s.internalSubjectCategories).toEqual(['工時'])
    expect(s.internalDomains).toEqual(['example.com'])
    expect(await rawStored()).toMatchObject({ internalDomains: ['example.com'] })
  })

  it('keeps closedCasePrefixes ⊆ excludeFolderPrefixes on every write', async () => {
    // closed-only write naming a prefix that isn't excluded → dropped
    await setSettings({ closedCasePrefixes: ['05已完成案件'] })
    expect((await getSettings()).closedCasePrefixes).toEqual([])

    // normalized match against the (un-normalized) exclude entry
    await setSettings({
      excludeFolderPrefixes: ['05已完成案件/', '06封存'],
      closedCasePrefixes: [' /05已完成案件 ', '06封存', '06封存', '99其他'],
    })
    expect((await getSettings()).closedCasePrefixes).toEqual(['05已完成案件', '06封存'])

    // an unrelated write leaves it alone
    await setSettings({ lastSyncAt: '2026-11-05T00:00:00.000Z' })
    expect((await getSettings()).closedCasePrefixes).toEqual(['05已完成案件', '06封存'])

    // an exclude-only write prunes it, against the FRESH stored value
    await setSettings({ excludeFolderPrefixes: ['06封存'] })
    expect((await getSettings()).closedCasePrefixes).toEqual(['06封存'])
    expect(await rawStored()).toMatchObject({ closedCasePrefixes: ['06封存'] })
  })

  it('normalizes primaryRootPath', async () => {
    await setSettings({ primaryRootPath: ' /案件/ ' })
    expect((await getSettings()).primaryRootPath).toBe('案件')
  })

  it('drops unknown keys at write time instead of persisting them', async () => {
    await setSettings({ batchSize: 30, junk: 'x', confidenceThreshold: 0.9 } as Partial<Settings>)
    const raw = await rawStored()
    expect(raw).not.toHaveProperty('junk')
    expect(raw).not.toHaveProperty('confidenceThreshold')
    expect(raw.batchSize).toBe(30)
  })

  it('an explicit undefined in a patch does not reset the field (API key survives)', async () => {
    await setSettings({ claudeApiKey: 'sk-ant-api03-KEEP' })
    await setSettings({ claudeApiKey: undefined, batchSize: 40 })
    expect(await getSettings()).toMatchObject({ claudeApiKey: 'sk-ant-api03-KEEP', batchSize: 40 })
  })

  it('backup restore (a direct storage.setSettings caller) is clamped too', async () => {
    const snapshotAt = '2026-11-01T00:00:00.000Z'
    const { claudeApiKey: _k, ...base } = DEFAULT_SETTINGS
    void _k
    await chrome.storage.local.set({
      [SYNC_BACKUPS_KEY]: [
        {
          snapshotAt,
          direction: 'pre-pull',
          payload: {
            rules: [],
            tombstones: [],
            settings: {
              ...base,
              batchSize: 0,
              aiConfidenceThreshold: 7,
              excludeFolderPrefixes: [],
              closedCasePrefixes: ['05已完成案件'],
            },
          },
        },
      ],
    })
    expect(await restoreBackup(snapshotAt)).toMatchObject({ restored: true })
    expect(await getSettings()).toMatchObject({
      batchSize: 1,
      aiConfidenceThreshold: 1,
      closedCasePrefixes: [],
    })
    expect(await rawStored()).toMatchObject({ batchSize: 1, aiConfidenceThreshold: 1 })
  })

  it('setSettings returns exactly what a subsequent read returns (idempotent normalization)', async () => {
    const written = await setSettings({
      batchSize: -4,
      excludeFolderPrefixes: [' / x', '/ x'],
      closedCasePrefixes: ['/ x'],
      primaryRootPath: '/ / 案件 / ',
    })
    expect(await getSettings()).toEqual(written)
    expect(await setSettings({})).toEqual(written)
  })
})

describe('skipHistory', () => {
  it('dedupes adds and counts correctly', async () => {
    const added1 = await addToSkipHistory(['a', 'b', 'c'])
    expect(added1).toBe(3)
    const added2 = await addToSkipHistory(['b', 'c', 'd'])
    expect(added2).toBe(1) // only 'd' is new
    expect(await getSkipHistoryCount()).toBe(4)
  })

  it('clearSkipHistory returns count and zeroes', async () => {
    await addToSkipHistory(['x', 'y'])
    const cleared = await clearSkipHistory()
    expect(cleared).toBe(2)
    expect(await getSkipHistoryCount()).toBe(0)
  })

  it('ignores empty ids', async () => {
    const added = await addToSkipHistory(['', 'real-id', ''])
    expect(added).toBe(1)
    expect(await getSkipHistoryCount()).toBe(1)
  })
})
