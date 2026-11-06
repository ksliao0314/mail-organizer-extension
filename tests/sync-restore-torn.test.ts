// 2026-11-06 review fixes:
//   A5  a restore (rollback) holds on every machine
//   A6  a cloud left torn under one pushId no longer locks machines out

import { beforeEach, describe, expect, it, vi } from 'vitest'
import { addRules, listRules, newRule, upsertRule } from '@/shared/rules'
import { getRuleTombstones, setSettings } from '@/shared/storage'
import {
  clearCloudState,
  listBackups,
  mergeCloudIntoLocal,
  pullNow,
  pushNow,
  quiesce,
  resetSyncEngineForTests,
  restoreBackup,
  SYNC_BACKUPS_KEY,
  type SyncBackup,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import { decodeRulePayload, decodeTombstone, type FolderDictEntry, type WireRuleEntry } from '@/shared/sync-codec'
import { SYNC_TORN_RECOVERY_MIN_MS, SYNC_TORN_RECOVERY_MIN_READS } from '@/shared/constants'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { readCloudMeta, simulateOtherMachinePush, syncListener } from './helpers/sync-sim'

const FOLDER_A = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const FOLDER_B = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdiYmJi'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: over.type ?? 'domain',
      signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
      targetFolderId: FOLDER_A,
      targetFolderPath: '案件/甲',
      confidence: 0.7,
      source: (over.source as RuleSource) ?? 'user_manual',
      enabled: over.enabled,
    }),
    ...over,
  }
}

async function byId(id: string): Promise<Rule | undefined> {
  return (await listRules()).find((r) => r.id === id)
}

async function cloud(): Promise<{ rules: Rule[]; tombstones: RuleTombstone[] }> {
  const all = await chrome.storage.sync.get(null)
  const pick = <T,>(prefix: string) =>
    unchunkArray(
      Object.keys(all)
        .filter((k) => k.startsWith(prefix))
        .map((k) => all[k] as { index: number; items: T[] }),
    )
  return {
    rules: decodeRulePayload({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') }).rules,
    tombstones: pick<unknown>('syncTombs4_').map(decodeTombstone).filter((t): t is RuleTombstone => t !== null),
  }
}

async function seedBackup(rules: Rule[]): Promise<string> {
  const { claudeApiKey: _k, ...settings } = DEFAULT_SETTINGS
  void _k
  const snapshotAt = new Date(Date.now() - 60_000).toISOString()
  const backup: SyncBackup = { snapshotAt, direction: 'manual', payload: { rules, tombstones: [], settings } }
  await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: [backup] })
  return snapshotAt
}

/** Every item the cloud meta references carries its pushId. */
async function cloudIntact(): Promise<boolean> {
  const all = await chrome.storage.sync.get(null)
  const meta = await readCloudMeta()
  return Object.keys(all)
    .filter((k) => k !== 'syncMeta')
    .every((k) => (all[k] as { p?: string }).p === meta!.pushId)
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- A5 -------------------------------------------------------------------------

describe('A5: a restore holds on every machine', () => {
  it('a rolled-back re-target is not reverted by the other machine’s newer copy', async () => {
    await addRules([rule({ id: 'R', signal: 'client.example' })])
    const at = await seedBackup([(await byId('R'))!])
    await upsertRule({ ...(await byId('R'))!, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' }, { actor: 'user' })
    const edited = (await byId('R'))!
    await pushNow('edit') // the edit reached the other machine
    await restoreBackup(at)
    expect((await byId('R'))!.targetFolderPath).toBe('案件/甲')
    // The other machine still holds the 乙 version and pushes it.
    await simulateOtherMachinePush({ rules: [edited] })
    await pullNow('remote', { mode: 'union' })
    // Before: the restored version kept its old editedAt and lost this merge
    // — the UI said 「已回復」 while the machines disagreed.
    expect((await byId('R'))!.targetFolderPath).toBe('案件/甲')
    await pushNow('after')
    expect((await cloud()).rules.find((r) => r.id === 'R')!.targetFolderPath).toBe('案件/甲')
  })

  it('a rule created after the snapshot is removed with a tombstone that propagates', async () => {
    await addRules([rule({ id: 'keep', signal: 'keep.example' })])
    const at = await seedBackup([(await byId('keep'))!])
    await addRules([
      rule({ id: 'later', signal: 'later.example', source: 'ai_confirmed' }),
      rule({ id: 'scan', signal: 'scan.example', source: 'auto_scan' }),
    ])
    const later = (await byId('later'))!
    await pushNow('later-reached-the-cloud')
    const r = await restoreBackup(at)
    expect(r).toMatchObject({ restored: true, removedCount: 1 })
    expect((await listRules()).map((x) => x.id)).toEqual(['keep'])
    const tombs = await getRuleTombstones()
    // Per-device auto_scan rules leave no tombstone (initial scan may re-learn them).
    expect(tombs.map((t) => t.signalNorm)).toEqual(['later.example'])
    expect(tombs[0]!.deletedAt).toBeGreaterThan(later.editedAt!)
    // The other machine still has `later` and pushes it: it does not come back here…
    await simulateOtherMachinePush({ rules: [later] })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('later')).toBeUndefined()
    // …and once we push, the other machine drops its copy too.
    await pushNow('after')
    const c = await cloud()
    const peer = mergeCloudIntoLocal({
      local: [later],
      localTombstones: [],
      cloudRules: c.rules,
      cloudStubs: [],
      cloudTombstones: c.tombstones,
      mode: 'union',
    })
    // Before: removed locally without a trace → the peer pushed it right back.
    expect(peer.next.map((x) => x.id)).toEqual(['keep'])
  })

  it('still caps confidence and keeps per-device settings', async () => {
    const hot = { ...rule({ id: 'hot', signal: 'hot.example', source: 'ai_confirmed' }), confidence: 0.99 }
    const at = await seedBackup([hot])
    await restoreBackup(at)
    expect((await byId('hot'))!.confidence).toBe(0.7)
  })
})

// ---- A6 -------------------------------------------------------------------------

describe('A6: a permanently torn cloud', () => {
  it('its meta owner rewrites it instead of skipping the push as "unchanged"', async () => {
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    await pushNow('first')
    const ours = (await readCloudMeta())!.pushId
    // Another machine's push left a chunk under OUR meta (its own meta lost).
    const row = (await chrome.storage.sync.get('syncRows4_0')).syncRows4_0 as object
    await chrome.storage.sync.set({ syncRows4_0: { ...row, p: 'otherpush001' } })
    const r = await pushNow('manual')
    // Before: same content hash → "unchanged" → the cloud stayed torn for everyone else.
    expect(r.pushed).toBe(true)
    expect(r.unchanged).toBeUndefined()
    expect((await readCloudMeta())!.pushId).not.toBe(ours)
    expect(await cloudIntact()).toBe(true)
  })

  it('the owner notices a foreign chunk under its meta and repairs on its own', async () => {
    const listener = syncListener()
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    await pushNow('first')
    const row = (await chrome.storage.sync.get('syncRows4_0')).syncRows4_0 as object
    const straggler = { ...row, p: 'otherpush002' }
    await chrome.storage.sync.set({ syncRows4_0: straggler })
    vi.useFakeTimers()
    try {
      // Before: our own meta → "echo" → ignored.
      listener({ syncRows4_0: { newValue: straggler } }, 'sync')
      await vi.advanceTimersByTimeAsync(1_000)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    expect(await cloudIntact()).toBe(true)
  })

  it('another machine rewrites it from local after repeated torn reads of the same push, after a backup', async () => {
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    const payload = await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.example' })] })
    // One row chunk belongs to another push — and stays that way (its owner is gone).
    await chrome.storage.sync.set({ syncRows4_0: { ...(payload.writes['syncRows4_0'] as object), p: 'ghostpush001' } })
    for (let i = 0; i < SYNC_TORN_RECOVERY_MIN_READS; i++) {
      // Within the first minute: it could still be a push in flight.
      expect((await pushNow('manual')).pushed).toBe(false)
    }
    const streak = (await chrome.storage.local.get('syncTornStreak')).syncTornStreak as { firstAt: number }
    await chrome.storage.local.set({
      syncTornStreak: { ...streak, firstAt: streak.firstAt - SYNC_TORN_RECOVERY_MIN_MS },
    })
    const r = await pushNow('manual')
    // Before: "pull before push failed" forever — not even 立即上傳 could push.
    expect(r.pushed).toBe(true)
    expect((await listBackups()).some((b) => b.trigger === 'torn-recovery')).toBe(true)
    expect((await readCloudMeta())!.sourceMachineId).toBe('this-machine')
    expect(await cloudIntact()).toBe(true)
    expect((await cloud()).rules.map((x) => x.id)).toEqual(['mine'])
    expect((await chrome.storage.local.get('syncTornStreak')).syncTornStreak).toBeUndefined()
  })

  it('never rewrites a torn cloud that carries a wipe this machine has not acted on', async () => {
    await simulateOtherMachinePush({ rules: [] })
    expect((await pullNow('join', { mode: 'union' })).pulled).toBe(true)
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    const payload = await simulateOtherMachinePush({ wipe: { id: 'wipe-x' }, rules: [rule({ id: 't', signal: 't.example' })] })
    await chrome.storage.sync.set({ syncRows4_0: { ...(payload.writes['syncRows4_0'] as object), p: 'ghostpush002' } })
    for (let i = 0; i < SYNC_TORN_RECOVERY_MIN_READS; i++) await pushNow('manual')
    const streak = (await chrome.storage.local.get('syncTornStreak')).syncTornStreak as { firstAt: number }
    await chrome.storage.local.set({ syncTornStreak: { ...streak, firstAt: 0 } })
    expect((await pushNow('manual')).pushed).toBe(false)
    expect((await readCloudMeta())!.sourceMachineId).toBe('other-machine')
  })

  it('a different torn pushId restarts the streak', async () => {
    const first = await simulateOtherMachinePush({ rules: [rule({ id: 'a', signal: 'a.example' })] })
    await chrome.storage.sync.set({ syncRows4_0: { ...(first.writes['syncRows4_0'] as object), p: 'ghost00000001' } })
    for (let i = 0; i < SYNC_TORN_RECOVERY_MIN_READS; i++) await pullNow('remote', { mode: 'union' })
    const second = await simulateOtherMachinePush({ rules: [rule({ id: 'b', signal: 'b.example' })] })
    await chrome.storage.sync.set({ syncRows4_0: { ...(second.writes['syncRows4_0'] as object), p: 'ghost00000002' } })
    await pullNow('remote', { mode: 'union' })
    expect((await chrome.storage.local.get('syncTornStreak')).syncTornStreak).toMatchObject({
      pushId: second.meta.pushId,
      count: 1,
    })
  })
})
