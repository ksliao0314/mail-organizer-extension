import { describe, expect, it } from 'vitest'
import {
  addToSkipHistory,
  clearSkipHistory,
  getRules,
  getSettings,
  getSkipHistoryCount,
  setRules,
  setSettings,
} from '@/shared/storage'
import type { Rule } from '@/shared/types'

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
    expect(cleaned.claudeModel).toBe('claude-sonnet-4-6')
    expect(cleaned.batchSize).toBe(50)
  })

  it('persists patch via setSettings cleanly', async () => {
    await setSettings({ batchSize: 100 })
    const after = await getSettings()
    expect(after.batchSize).toBe(100)
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
