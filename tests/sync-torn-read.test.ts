// Item 4 (2026-11-05): torn reads. A push's meta and chunks are written in one
// set() on the pushing machine, but reach another machine item by item. The
// pull used to read meta and chunks separately and decode whatever it found:
// a folder ref with no dictionary entry became an EMPTY target that replaced
// the good local copy and re-spread on the next push. Now every v4 chunk
// carries the meta's pushId; a mismatch aborts the pull without writing,
// retries with backoff, and a late chunk arrival triggers a re-pull.

import { beforeEach, describe, expect, it, vi } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import { getRuleTombstones, getSettings, setSettings } from '@/shared/storage'
import {
  buildPushPayload,
  getSyncStatus,
  pullNow,
  quiesce,
  resetSyncEngineForTests,
} from '@/background/sync-engine'
import { SYNC_TORN_RETRY_MAX_ATTEMPTS } from '@/shared/constants'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { Rule, RuleSource } from '@/shared/types'
import type { FolderDictEntry } from '@/shared/sync-codec'
import { simulateOtherMachinePush, syncListener } from './helpers/sync-sim'

function rule(over: Partial<Rule> = {}): Rule {
  // newRule() ignores id / counters — apply overrides on the result.
  return {
    ...newRule({
      type: 'domain',
      signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: '案件/甲',
      confidence: 0.7,
      source: (over.source as RuleSource) ?? 'user_manual',
      enabled: over.enabled,
    }),
    ...over,
  }
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

/** Push #2's full writes, built but not stored. */
function nextPush(rules: Rule[], pushId: string) {
  return buildPushPayload({
    machineId: 'other-machine',
    rules,
    tombstones: [{ type: 'domain', signalNorm: 'tomb.com', targetFolderPath: 'X', deletedAt: Date.now() }],
    folderActivity: [],
    settings: { ...DEFAULT_SETTINGS, claudeModel: 'claude-haiku-5-5' },
    updatedAt: new Date().toISOString(),
    pushId,
    now: Date.now(),
  })
}

describe('Item 4: a torn snapshot aborts the pull without writing anything', () => {
  it('a chunk from another push (pushId mismatch) → nothing written, torn reported', async () => {
    await addRules([rule({ id: 'local', signal: 'local.com' })])
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00001')
    // push #2's meta + settings landed, its row chunk is still push #1's.
    await simulateOtherMachinePush({ rules: [rule({ id: 'old', signal: 'old.com' })] })
    const stale = (await chrome.storage.sync.get('syncRows4_0')).syncRows4_0
    await chrome.storage.sync.set({ ...second.writes, syncRows4_0: stale })
    const r = await pullNow('remote', { mode: 'union' })
    expect(r).toMatchObject({ pulled: false, torn: true })
    expect((await listRules()).map((x) => x.id)).toEqual(['local'])
    expect(await getRuleTombstones()).toEqual([])
    expect((await getSettings()).claudeModel).not.toBe('claude-haiku-5-5')
  })

  it('a missing chunk → same abort', async () => {
    await addRules([rule({ id: 'local', signal: 'local.com' })])
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00002')
    const { syncFolders4_0: _late, ...early } = second.writes
    void _late
    await chrome.storage.sync.set(early)
    const r = await pullNow('remote', { mode: 'replace' })
    expect(r).toMatchObject({ pulled: false, torn: true })
    expect((await listRules()).map((x) => x.id)).toEqual(['local'])
  })

  it('once the late chunks land, the next pull merges the complete push', async () => {
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00003')
    const { syncFolders4_0: late, ...early } = second.writes
    await chrome.storage.sync.set(early)
    expect((await pullNow('remote', { mode: 'union' })).pulled).toBe(false)
    await chrome.storage.sync.set({ syncFolders4_0: late })
    expect((await pullNow('remote', { mode: 'union' })).pulled).toBe(true)
    expect((await listRules()).map((x) => x.id)).toEqual(['a'])
    expect((await getSettings()).claudeModel).toBe('claude-haiku-5-5')
  })

  it('a late chunk arrival alone (no syncMeta in the change) triggers a re-pull', async () => {
    const listener = syncListener()
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00004')
    const { syncFolders4_0: late, ...early } = second.writes
    await chrome.storage.sync.set(early)
    expect((await pullNow('remote', { mode: 'union' })).torn).toBe(true)
    resetSyncEngineForTests() // drop the backoff timer: prove the chunk event alone re-pulls
    await chrome.storage.sync.set({ syncFolders4_0: late })
    vi.useFakeTimers()
    try {
      // Before: the listener returned unless the change contained syncMeta.
      listener({ syncFolders4_0: { newValue: late } }, 'sync')
      await vi.advanceTimersByTimeAsync(300)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    expect((await listRules()).map((x) => x.id)).toEqual(['a'])
  })

  it('a torn read retries on its own with backoff', async () => {
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00005')
    const { syncFolders4_0: late, ...early } = second.writes
    await chrome.storage.sync.set(early)
    vi.useFakeTimers()
    try {
      expect((await pullNow('remote', { mode: 'union' })).torn).toBe(true)
      await chrome.storage.sync.set({ syncFolders4_0: late })
      await vi.advanceTimersByTimeAsync(2_500)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    expect((await listRules()).map((x) => x.id)).toEqual(['a'])
  })

  it('a persistently torn cloud surfaces in lastError after repeated attempts', async () => {
    const second = nextPush([rule({ id: 'a', signal: 'a.com' })], 'pushtwo00006')
    const { syncFolders4_0: _late, ...early } = second.writes
    void _late
    await chrome.storage.sync.set(early)
    for (let i = 0; i < SYNC_TORN_RETRY_MAX_ATTEMPTS; i++) {
      expect((await pullNow('remote', { mode: 'union' })).torn).toBe(true)
    }
    expect((await getSyncStatus()).lastError?.reason).toContain('incomplete')
  })
})

describe('Item 4: a row whose folder ref did not resolve never replaces the local copy', () => {
  it('the row is skipped; the local target survives', async () => {
    await addRules([rule({ id: 'keep', signal: 'keep.com' })])
    const local = (await listRules())[0]!
    // A buggy writer (or a dict chunk that lost the entry) — same pushId, so
    // the generation check passes, but the folder ref dangles.
    const payload = nextPush([{ ...local, editedAt: local.editedAt! + 5000, targetFolderPath: '案件/乙' }], 'pushtwo00007')
    const folders = payload.writes['syncFolders4_0'] as { index: number; items: FolderDictEntry[]; p: string }
    await chrome.storage.sync.set({ ...payload.writes, syncFolders4_0: { ...folders, items: [] } })
    const r = await pullNow('remote', { mode: 'replace' })
    expect(r.pulled).toBe(true)
    expect(r.mergeMode).toBe('union') // an imperfect snapshot never deletes
    const after = (await listRules())[0]!
    // Before: decoded to targetFolderId '' / targetFolderPath '' and replaced it.
    expect(after.targetFolderId).toBe(local.targetFolderId)
    expect(after.targetFolderPath).toBe('案件/甲')
  })
})
