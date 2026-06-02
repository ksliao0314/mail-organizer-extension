// Tests for the cross-machine wipe propagation feature (audit P0-1 +
// P0-3 fixes, 2026-05-27). Covers:
//
//   1. pushNow({ wipeMarker: true }) stamps SyncMeta.wipeMarker with
//      our machineId + current ISO.
//   2. A normal push PRESERVES an existing wipeMarker in cloud meta
//      (so a 3rd machine connecting later still trips the wipe).
//   3. pullNow detects a fresh wipeMarker from ANOTHER machine and:
//        - drops local syncable rules
//        - clears local tombstones
//        - persists a remoteWipeNotice entry in chrome.storage.local
//   4. Echo guard: we never trip our own marker.
//   5. Staleness guard: marker.at <= lastSyncAt → no re-trip.
//
// Setup mirrors sync-engine.test.ts conventions (fake-storage,
// per-test seeding of syncMachineId).

import { beforeEach, describe, expect, it } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import {
  addRuleTombstones,
  getRuleTombstones,
  getSettings,
  setSettings,
} from '@/shared/storage'
import {
  clearCloudState,
  pullNow,
  pushNow,
  readRemoteWipeNotice,
  REMOTE_WIPE_NOTICE_KEY,
  type RemoteWipeNotice,
  type SyncMeta,
} from '@/background/sync-engine'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return newRule({
    type: 'domain',
    signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
    targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
    targetFolderPath: 'X/Y',
    confidence: 0.7,
    source: (over.source as RuleSource) ?? 'user_manual',
    ...over,
  })
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
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

async function readCloudMeta(): Promise<SyncMeta | undefined> {
  const r = await chrome.storage.sync.get('syncMeta')
  return r.syncMeta as SyncMeta | undefined
}

// ---- 1. Push with wipeMarker stamps SyncMeta -----------------------------

describe('pushNow({ wipeMarker: true }) — stamps SyncMeta.wipeMarker', () => {
  it('stamps marker with our machineId on a wipe push', async () => {
    const before = new Date().toISOString()
    await pushNow('post-wipe', { wipeMarker: true })
    const meta = await readCloudMeta()
    expect(meta).toBeDefined()
    expect(meta!.wipeMarker).toBeDefined()
    expect(meta!.wipeMarker!.byMachineId).toBe('this-machine')
    expect(meta!.wipeMarker!.at >= before).toBe(true)
  })

  it('non-wipe push leaves wipeMarker undefined when cloud had none', async () => {
    await addRules([rule()])
    await pushNow('normal')
    const meta = await readCloudMeta()
    expect(meta).toBeDefined()
    expect(meta!.wipeMarker).toBeUndefined()
  })
})

// ---- 2. Normal push preserves an existing wipeMarker ----------------------

describe('Normal push preserves cloud wipeMarker', () => {
  it('does not erase an earlier wipeMarker from another machine', async () => {
    // Simulate: another machine pushed with a wipeMarker earlier.
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: '2026-05-26T00:00:00Z',
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        recentPushes: [{ machineId: 'other-machine', at: '2026-05-26T00:00:00Z' }],
        wipeMarker: {
          at: '2026-05-26T00:00:00Z',
          byMachineId: 'other-machine',
        },
      } satisfies SyncMeta,
    })
    // We pull first so our lastSyncAt advances past the marker
    // (otherwise our own pull would TRIP the marker and clear local).
    await pullNow('first-pull', { mode: 'union' })

    // Now do a normal push. It should keep the original marker around
    // so a 3rd machine arriving later still trips it.
    await addRules([rule({ signal: 'normal.com' })])
    await pushNow('normal')
    const meta = await readCloudMeta()
    expect(meta!.wipeMarker).toEqual({
      at: '2026-05-26T00:00:00Z',
      byMachineId: 'other-machine',
    })
    // The pusher (us) is now the source, but the wipeMarker still
    // credits the original issuing machine.
    expect(meta!.sourceMachineId).toBe('this-machine')
  })

  it('wipe push REPLACES an older marker with a fresh one', async () => {
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: '2026-05-26T00:00:00Z',
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: {
          at: '2026-05-26T00:00:00Z',
          byMachineId: 'other-machine',
        },
      } satisfies SyncMeta,
    })
    // A fresh wipe push replaces the marker — same field, different
    // issuer + newer timestamp.
    await pushNow('post-wipe', { wipeMarker: true })
    const meta = await readCloudMeta()
    expect(meta!.wipeMarker!.byMachineId).toBe('this-machine')
    expect(meta!.wipeMarker!.at > '2026-05-26T00:00:00Z').toBe(true)
  })
})

// ---- 3. Pull applies remote wipe -----------------------------------------

describe('pullNow — applies remote wipeMarker', () => {
  it('drops syncable rules but PRESERVES local tombstones + persists notice', async () => {
    // F1: this machine must look like it has ALREADY participated in
    // sync (lastSyncAt non-empty + older than the wipe) for the marker
    // to trip. A first-time participant (lastSyncAt === '') is exempt.
    await setSettings({ lastSyncAt: '2020-01-01T00:00:00.000Z' })
    // Seed our local with rules + tombstones (some syncable, some not)
    await addRules([
      rule({ signal: 'keep-auto.com', source: 'auto_scan' }), // not synced
      rule({ signal: 'syncable.com', source: 'ai_confirmed' }),
      rule({ signal: 'manual.com', source: 'user_manual' }),
    ])
    await addRuleTombstones([
      tombstone({ signalNorm: 'old.com' }),
      tombstone({ signalNorm: 'older.com' }),
    ])

    // Simulate another machine pushed with a wipeMarker NEWER than
    // our lastSyncAt.
    const wipeAt = new Date().toISOString()
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: wipeAt,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: wipeAt, byMachineId: 'other-machine' },
      } satisfies SyncMeta,
    })

    await pullNow('test', { mode: 'union' })

    // Syncable rules dropped; auto_scan rule kept
    const remaining = await listRules()
    const remainingSignals = remaining.map((r) => r.signal).sort()
    expect(remainingSignals).toEqual(['keep-auto.com'])

    // F3: local tombstones PRESERVED (this machine's own deletion
    // intents — the user didn't click "全部刪除" here, so they must
    // survive to keep blocking resurrection).
    const tomb = await getRuleTombstones()
    expect(tomb.map((t) => t.signalNorm).sort()).toEqual(['old.com', 'older.com'])

    // Notice persisted
    const notice = (await chrome.storage.local.get(REMOTE_WIPE_NOTICE_KEY))[
      REMOTE_WIPE_NOTICE_KEY
    ] as RemoteWipeNotice | undefined
    expect(notice).toBeDefined()
    expect(notice!.byMachineId).toBe('other-machine')
    expect(notice!.at).toBe(wipeAt)
    expect(notice!.droppedRuleCount).toBe(2) // syncable + manual
  })

  it('F1: first-time participant (lastSyncAt empty) is NOT wiped', async () => {
    // lastSyncAt defaults to '' (set in beforeEach via setSettings that
    // doesn't touch it). A machine joining sync for the first time must
    // keep its locally-built rules even if cloud carries an old wipe.
    await addRules([
      rule({ signal: 'local-work.com', source: 'user_manual' }),
      rule({ signal: 'learned.com', source: 'ai_confirmed' }),
    ])
    const wipeAt = new Date().toISOString()
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: wipeAt,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: wipeAt, byMachineId: 'other-machine' },
      } satisfies SyncMeta,
    })

    await pullNow('first-enable', { mode: 'union' })

    const remaining = await listRules()
    expect(remaining.map((r) => r.signal).sort()).toEqual([
      'learned.com',
      'local-work.com',
    ])
    // No wipe → no notice.
    expect(await readRemoteWipeNotice()).toBeUndefined()
  })

  it('F2: wipe apply advances lastSyncAt to marker time (no re-trip)', async () => {
    await setSettings({ lastSyncAt: '2020-01-01T00:00:00.000Z' })
    await addRules([rule({ signal: 'gone.com', source: 'ai_confirmed' })])
    const wipeAt = new Date().toISOString()
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: wipeAt,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: wipeAt, byMachineId: 'other-machine' },
      } satisfies SyncMeta,
    })

    await pullNow('test', { mode: 'union' })
    // lastSyncAt must now be >= marker.at so a subsequent pull of the
    // same marker does NOT re-trip the wipe.
    const settings = await getSettings()
    expect(settings.lastSyncAt >= wipeAt).toBe(true)
  })

  it('readRemoteWipeNotice returns the stored notice', async () => {
    await setSettings({ lastSyncAt: '2020-01-01T00:00:00.000Z' })
    // Setup: cause a pull with remote wipe so the notice is persisted.
    await addRules([rule({ signal: 'x.com', source: 'user_manual' })])
    const wipeAt = new Date().toISOString()
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: wipeAt,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: wipeAt, byMachineId: 'other-machine' },
      } satisfies SyncMeta,
    })
    await pullNow('test', { mode: 'union' })

    const notice = await readRemoteWipeNotice()
    expect(notice).toBeDefined()
    expect(notice!.byMachineId).toBe('other-machine')
  })
})

// ---- 4. Echo guard: own wipeMarker never trips ---------------------------

describe('Echo guard — own wipeMarker does not self-wipe', () => {
  it('skips remote-wipe when marker.byMachineId === ours', async () => {
    await addRules([rule({ signal: 'self.com', source: 'user_manual' })])
    const wipeAt = new Date().toISOString()
    // Marker issued by us — should NOT trigger local wipe even though
    // cloud has fresh-looking marker
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'this-machine',
        updatedAt: wipeAt,
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: wipeAt, byMachineId: 'this-machine' },
      } satisfies SyncMeta,
    })

    await pullNow('test', { mode: 'union' })

    const remaining = await listRules()
    // The user_manual rule survives because the marker is OURS
    expect(remaining.map((r) => r.signal)).toContain('self.com')
  })
})

// ---- 5. Staleness guard --------------------------------------------------

describe('Staleness guard — marker.at <= lastSyncAt does not re-trip', () => {
  it('skips when we already applied this marker (lastSyncAt advanced past marker.at)', async () => {
    await addRules([rule({ signal: 'persist.com', source: 'user_manual' })])
    // Our lastSyncAt is AFTER the marker — meaning we already processed
    // this wipe last time. Should NOT re-trip.
    await setSettings({ lastSyncAt: '2026-06-01T00:00:00Z' })
    await chrome.storage.sync.set({
      syncMeta: {
        schemaVersion: 2,
        sourceMachineId: 'other-machine',
        updatedAt: '2026-05-26T00:00:00Z',
        ruleCount: 0,
        tombstoneCount: 0,
        ruleChunkCount: 0,
        tombstoneChunkCount: 0,
        wipeMarker: { at: '2026-05-26T00:00:00Z', byMachineId: 'other-machine' },
      } satisfies SyncMeta,
    })

    await pullNow('test', { mode: 'union' })

    const remaining = await listRules()
    expect(remaining.map((r) => r.signal)).toContain('persist.com')
  })
})

// ---- Cleanup ----

beforeEach(async () => {
  await clearCloudState()
})
