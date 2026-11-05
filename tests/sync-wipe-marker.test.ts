// Cross-machine wipe propagation (audit P0-1 + P0-3, 2026-05-27; wipe ids,
// 2026-11-05). Covers:
//
//   1. pushNow({ wipeMarker: true }) stamps SyncMeta.wipeMarker with our
//      machineId + a fresh random id.
//   2. A normal push PRESERVES an existing wipeMarker in cloud meta
//      (so a 3rd machine connecting later still trips the wipe).
//   3. A participating machine applies an UNSEEN marker exactly once:
//        - drops local syncable rules, keeps auto_scan + local tombstones
//        - pins a pre-wipe backup, persists a remoteWipeNotice
//   4. Echo guard: we never trip our own marker.
//   5. Item 6: markers are matched by id, never by comparing timestamps
//      written by different machines' clocks; first participation (incl.
//      re-enable after disable) records the marker without applying it.

import { beforeEach, describe, expect, it } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, setSettings } from '@/shared/storage'
import {
  clearCloudState,
  listBackups,
  pullNow,
  pushNow,
  readRemoteWipeNotice,
  REMOTE_WIPE_NOTICE_KEY,
  resetSyncEngineForTests,
  SYNC_LOCAL_STATE_KEY,
  type RemoteWipeNotice,
  type SyncMeta,
} from '@/background/sync-engine'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { readCloudMeta, simulateOtherMachinePush, simulateV3Push } from './helpers/sync-sim'

function rule(over: Partial<Rule> = {}): Rule {
  // newRule() ignores id / counters — apply overrides on the result.
  return {
    ...newRule({
      type: 'domain',
      signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
      targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
      targetFolderPath: 'X/Y',
      confidence: 0.7,
      source: (over.source as RuleSource) ?? 'user_manual',
      enabled: over.enabled,
    }),
    ...over,
  }
}

function tombstone(over: Partial<RuleTombstone> = {}): RuleTombstone {
  return {
    type: 'domain',
    signalNorm: over.signalNorm ?? 'a.com',
    targetFolderPath: over.targetFolderPath ?? 'X/Y',
    deletedAt: over.deletedAt ?? Date.now(),
    ...over,
  }
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

/** Make this machine a sync participant (it has merged the cloud once). */
async function participate(): Promise<void> {
  await simulateOtherMachinePush({ rules: [] })
  const r = await pullNow('join', { mode: 'union' })
  expect(r.pulled).toBe(true)
}

// ---- 1. Push with wipeMarker stamps SyncMeta -----------------------------

describe('pushNow({ wipeMarker: true }) — stamps SyncMeta.wipeMarker', () => {
  it('stamps marker with our machineId and a fresh id on a wipe push', async () => {
    const before = new Date().toISOString()
    await pushNow('post-wipe', { wipeMarker: true })
    const meta = (await readCloudMeta())!
    expect(meta.wipeMarker!.byMachineId).toBe('this-machine')
    expect(meta.wipeMarker!.at >= before).toBe(true)
    expect(meta.wipeMarker!.id).toMatch(/^[a-z0-9]{8,}$/)
    await pushNow('post-wipe-2', { wipeMarker: true })
    expect((await readCloudMeta())!.wipeMarker!.id).not.toBe(meta.wipeMarker!.id)
  })

  it('non-wipe push leaves wipeMarker undefined when cloud had none', async () => {
    await addRules([rule()])
    await pushNow('normal')
    expect((await readCloudMeta())!.wipeMarker).toBeUndefined()
  })
})

// ---- 2. Normal push preserves an existing wipeMarker ----------------------

describe('Normal push preserves cloud wipeMarker', () => {
  it('does not erase an earlier wipeMarker from another machine', async () => {
    await simulateOtherMachinePush({ wipe: { id: 'wipe-old', at: '2026-05-26T00:00:00.000Z' } })
    await addRules([rule({ signal: 'normal.com' })])
    await pushNow('normal') // pulls first (records the marker), then pushes
    const meta = (await readCloudMeta())!
    expect(meta.wipeMarker).toEqual({
      id: 'wipe-old',
      at: '2026-05-26T00:00:00.000Z',
      byMachineId: 'other-machine',
    })
    expect(meta.sourceMachineId).toBe('this-machine')
  })

  it('wipe push REPLACES an older marker with a fresh one', async () => {
    await simulateOtherMachinePush({ wipe: { id: 'wipe-old', at: '2026-05-26T00:00:00.000Z' } })
    await pushNow('post-wipe', { wipeMarker: true })
    const meta = (await readCloudMeta())!
    expect(meta.wipeMarker!.byMachineId).toBe('this-machine')
    expect(meta.wipeMarker!.id).not.toBe('wipe-old')
  })
})

// ---- 3. Pull applies remote wipe -----------------------------------------

describe('pullNow — applies an unseen remote wipeMarker', () => {
  it('drops syncable rules, PRESERVES local tombstones, pins a backup, persists notice', async () => {
    await participate()
    await addRules([
      rule({ signal: 'keep-auto.com', source: 'auto_scan' }), // not synced
      rule({ signal: 'syncable.com', source: 'ai_confirmed' }),
      rule({ signal: 'manual.com', source: 'user_manual' }),
    ])
    await addRuleTombstones([tombstone({ signalNorm: 'old.com' }), tombstone({ signalNorm: 'older.com' })])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-1' } })
    await pullNow('test', { mode: 'union' })

    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['keep-auto.com'])
    // F3: this machine's own deletion intents survive a REMOTE wipe.
    expect((await getRuleTombstones()).map((t) => t.signalNorm).sort()).toEqual(['old.com', 'older.com'])
    const notice = (await chrome.storage.local.get(REMOTE_WIPE_NOTICE_KEY))[REMOTE_WIPE_NOTICE_KEY] as
      | RemoteWipeNotice
      | undefined
    expect(notice!.byMachineId).toBe('other-machine')
    expect(notice!.droppedRuleCount).toBe(2)
    // Item 8: the pre-wipe snapshot the banner points to is pinned.
    const pinned = (await listBackups()).find((b) => b.pinned)
    expect(pinned?.trigger).toBe('remote-wipe')
    expect(pinned!.payload.rules.map((r) => r.signal).sort()).toEqual([
      'keep-auto.com',
      'manual.com',
      'syncable.com',
    ])
  })

  it('the pinned pre-wipe snapshot survives backup rotation', async () => {
    await participate()
    await addRules([rule({ signal: 'precious.com' })])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-rot' } })
    await pullNow('wipe', { mode: 'union' })
    for (let i = 0; i < 8; i++) await pullNow(`replace-${i}`, { mode: 'replace' })
    const backups = await listBackups()
    expect(backups.filter((b) => !b.pinned).length).toBe(5)
    expect(backups.find((b) => b.pinned)!.payload.rules.map((r) => r.signal)).toEqual(['precious.com'])
  })

  it('F1: a first-time participant is NOT wiped — and the recorded marker never trips later', async () => {
    await addRules([
      rule({ signal: 'local-work.com', source: 'user_manual' }),
      rule({ signal: 'learned.com', source: 'ai_confirmed' }),
    ])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-before-join' } })
    await pullNow('first-enable', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['learned.com', 'local-work.com'])
    expect(await readRemoteWipeNotice()).toBeUndefined()
    // Now a participant: the SAME marker (carried by later pushes) is seen.
    await simulateOtherMachinePush({ rules: [] })
    await pullNow('later', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal).sort()).toEqual(['learned.com', 'local-work.com'])
  })

  it('readRemoteWipeNotice returns the stored notice', async () => {
    await participate()
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-n' } })
    await pullNow('test', { mode: 'union' })
    expect((await readRemoteWipeNotice())!.byMachineId).toBe('other-machine')
  })
})

// ---- 4. Echo guard ---------------------------------------------------------

describe('Echo guard — own wipeMarker does not self-wipe', () => {
  it('skips remote-wipe when marker.byMachineId === ours', async () => {
    await participate()
    await addRules([rule({ signal: 'self.com', source: 'user_manual' })])
    await simulateOtherMachinePush({ machineId: 'this-machine', wipe: { id: 'own-wipe' } })
    await pullNow('test', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal)).toContain('self.com')
  })
})

// ---- 5. Item 6: ids, not clocks ---------------------------------------------

describe('Item 6: a wipe applies once per id, independent of machine clocks', () => {
  it('applied exactly once even when this machine runs behind the wiper (clock skew)', async () => {
    await participate()
    await addRules([rule({ signal: 'pre-wipe.com', source: 'ai_confirmed' })])
    // The wiper's clock runs an hour ahead of ours.
    await simulateOtherMachinePush({ wipe: { id: 'skewed', at: new Date(Date.now() + 3_600_000).toISOString() } })
    await pullNow('first', { mode: 'union' })
    expect(await listRules()).toEqual([])
    // The user re-creates a rule and this machine pushes: under the old rule
    // lastSyncAt became OUR clock (< marker.at) and the next pull re-applied
    // the wipe, deleting the re-created rule.
    await addRules([rule({ signal: 'recreated.com', source: 'user_manual' })])
    expect((await pushNow('local-change')).pushed).toBe(true)
    await simulateOtherMachinePush({ rules: [] }) // a later push, same marker carried
    await pullNow('second', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal)).toEqual(['recreated.com'])
  })

  it('a machine that pushes before pulling still applies the wipe (its push pulls first)', async () => {
    await participate()
    await addRules([rule({ signal: 'stale-on-b.com', source: 'ai_confirmed' })])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-b' } })
    // B edits something and its debounced push fires BEFORE any pull ran.
    await addRules([rule({ signal: 'new-on-b.com', source: 'user_manual' })])
    const r = await pushNow('local-change')
    expect(r.pushed).toBe(true)
    // Before: B's push stamped lastSyncAt past marker.at → never applied, and
    // re-uploaded every pre-wipe rule.
    expect(await listRules()).toEqual([])
    expect((await readCloudMeta())!.ruleCount).toBe(0)
  })

  it('disabling sync resets participation: re-enabling records an old marker without applying it', async () => {
    const { handleDisableSync, handleEnableSync } = await import('@/background/handlers/sync')
    await participate()
    await simulateOtherMachinePush({ wipe: { id: 'wipe-while-away' } })
    await handleDisableSync({ keepCloud: true })
    // While sync is off the user builds rules here.
    await addRules([rule({ signal: 'built-offline.com', source: 'user_manual' })])
    const r = await handleEnableSync()
    expect(r.ok).toBe(true)
    expect((await listRules()).map((x) => x.signal)).toEqual(['built-offline.com'])
  })

  it('upgrade from v1.0.7 (no local state): the current legacy marker is judged once by lastSyncAt', async () => {
    // Applied: issued after this machine's last v1.0.7 sync.
    await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
    await setSettings({ lastSyncAt: '2026-06-01T00:00:00.000Z' })
    await addRules([rule({ signal: 'gone.com', source: 'ai_confirmed' })])
    await simulateV3Push({ wipeMarker: { at: '2026-07-01T00:00:00.000Z', byMachineId: 'v3-machine' } })
    await pullNow('upgrade', { mode: 'union' })
    expect(await listRules()).toEqual([])
  })

  it('upgrade from v1.0.7: a legacy marker older than lastSyncAt (already applied) is not re-applied', async () => {
    await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
    await setSettings({ lastSyncAt: '2026-08-01T00:00:00.000Z' })
    await addRules([rule({ signal: 'kept.com', source: 'ai_confirmed' })])
    await simulateV3Push({ wipeMarker: { at: '2026-07-01T00:00:00.000Z', byMachineId: 'v3-machine' } })
    await pullNow('upgrade', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal)).toEqual(['kept.com'])
    // And it is now recorded: the v4 push that carries it forward never trips it.
    await pushNow('after')
    await pullNow('again', { mode: 'union' })
    expect((await listRules()).map((r) => r.signal)).toEqual(['kept.com'])
  })

  it('a legacy (v3) marker carried into v4 keeps one stable id', async () => {
    await simulateV3Push({ wipeMarker: { at: '2026-07-01T00:00:00.000Z', byMachineId: 'v3-machine' } })
    await addRules([rule({ signal: 'x.com' })])
    await pushNow('upgrade-cloud')
    const meta = (await readCloudMeta()) as SyncMeta
    expect(meta.schemaVersion).toBe(4)
    expect(meta.wipeMarker).toEqual({
      id: 'legacy:v3-machine:2026-07-01T00:00:00.000Z',
      at: '2026-07-01T00:00:00.000Z',
      byMachineId: 'v3-machine',
    })
  })
})
