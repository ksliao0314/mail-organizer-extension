// First-participation merge ("adoption") and the v1.0.7 → v4 transition
// (2026-11-06 review fixes A1, A4, A12).
//
// v1.0.7 (sync v3) never let a tombstone remove a rule the cloud still had,
// so real v3 libraries hold "live rule + newer tombstone of its triple"
// pairs. v1.0.8's first v4 merge deleted every such rule — on every machine,
// without a backup. Now the first merge keeps them (re-stamped so the verdict
// holds everywhere), takes a pinned 「升級同步格式前」 backup first, and a
// machine re-enabling sync is a first-time participant again.

import { beforeEach, describe, expect, it } from 'vitest'
import { addRules, listRules, newRule } from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, getSettings, setRules, setSettings } from '@/shared/storage'
import {
  backupLabel,
  clearCloudState,
  listBackups,
  mergeCloudIntoLocal,
  pullNow,
  pushNow,
  readRemoteWipeNotice,
  resetSyncEngineForTests,
  restoreBackup,
  SYNC_LOCAL_STATE_KEY,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import { decodeRulesFromSync, type FolderDictEntry, type WireRuleEntry } from '@/shared/sync-codec'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { simulateOtherMachinePush, simulateV3Push } from './helpers/sync-sim'

const FOLDER_A = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const FOLDER_B = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdiYmJi'
const LAST_V3_SYNC = '2026-10-01T00:00:00.000Z'

/** A rule as v1.0.7 stored it: no editedAt. */
function v3Rule(over: Partial<Rule> = {}): Rule {
  const r: Rule = {
    ...newRule({
      type: over.type ?? 'domain',
      signal: over.signal ?? `r${Math.random().toString(36).slice(2)}.com`,
      targetFolderId: FOLDER_A,
      targetFolderPath: '案件/甲',
      confidence: 0.7,
      source: (over.source as RuleSource) ?? 'ai_confirmed',
    }),
    createdAt: '2026-08-01T00:00:00.000Z',
    ...over,
  }
  delete r.editedAt
  return r
}

function tomb(signalNorm: string, deletedAt: string, over: Partial<RuleTombstone> = {}): RuleTombstone {
  return { type: 'domain', signalNorm, targetFolderPath: '案件/甲', deletedAt: Date.parse(deletedAt), ...over }
}

async function cloudRules(): Promise<Rule[]> {
  const cloud = await chrome.storage.sync.get(null)
  const pick = <T,>(prefix: string) =>
    unchunkArray(
      Object.keys(cloud)
        .filter((k) => k.startsWith(prefix))
        .map((k) => cloud[k] as { index: number; items: T[] }),
    )
  return decodeRulesFromSync({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') })
}

/** This machine ran v1.0.7 with sync on and last synced at LAST_V3_SYNC. */
async function asUpgradedV3Machine(): Promise<void> {
  await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine', lastSyncAt: LAST_V3_SYNC })
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

describe('A1: the first v4 merge keeps the live rules v3 kept', () => {
  it('a live rule next to a tombstone v3 had processed survives, re-stamped above it, after a pinned upgrade backup', async () => {
    await asUpgradedV3Machine()
    // One of two duplicates was deleted before the last v3 sync; v3 kept the
    // other because the cloud still had it.
    const live = v3Rule({ id: 'dup-2', signal: 'twin.example' })
    const t = tomb('twin.example', '2026-09-15T00:00:00.000Z')
    await setRules([live])
    await addRuleTombstones([t])
    await simulateV3Push({ rules: [live], tombstones: [t], updatedAt: LAST_V3_SYNC })

    expect((await pullNow('upgrade', { mode: 'union' })).pulled).toBe(true)

    // Before: the rule was gone and there was no backup ("0 rules, 0 backups").
    const kept = (await listRules()).find((r) => r.id === 'dup-2')
    expect(kept).toBeDefined()
    expect(kept!.editedAt).toBeGreaterThan(t.deletedAt)
    const upgrade = (await listBackups()).find((b) => b.trigger === 'upgrade')!
    expect(upgrade.pinned).toBe(true)
    expect(backupLabel(upgrade)).toBe('升級同步格式前')
    expect(upgrade.payload.rules.map((r) => r.id)).toEqual(['dup-2'])
    expect(upgrade.payload.tombstones).toEqual([t])

    // Convergent: once pushed, a machine that merges WITHOUT adoption (v1.0.8,
    // or this build after its own first merge) keeps it too.
    expect((await pushNow('after-upgrade')).pushed).toBe(true)
    const peer = mergeCloudIntoLocal({
      local: [live],
      localTombstones: [t],
      cloudRules: await cloudRules(),
      cloudStubs: [],
      cloudTombstones: [t],
      mode: 'union',
    })
    expect(peer.next.map((r) => r.id)).toEqual(['dup-2'])
  })

  it('a deletion made after the last v3 sync still propagates', async () => {
    await asUpgradedV3Machine()
    const live = v3Rule({ id: 'doomed', signal: 'deleted-later.example' })
    await setRules([live])
    await simulateV3Push({
      tombstones: [tomb('deleted-later.example', '2026-10-20T00:00:00.000Z')],
      updatedAt: '2026-10-20T00:00:01.000Z',
    })
    await pullNow('upgrade', { mode: 'union' })
    expect(await listRules()).toEqual([])
  })

  it("this machine's own tombstones never delete its live rules on the first merge (never-synced machine, first enable)", async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await setSettings({ syncEnabled: false })
    const live = v3Rule({ id: 'kept-locally', signal: 'mine.example', source: 'user_manual' })
    await setRules([live])
    // A duplicate of it was deleted here last week; the library kept routing with this one.
    await addRuleTombstones([tomb('mine.example', new Date(Date.now() - 7 * 86_400_000).toISOString())])
    await simulateOtherMachinePush({ rules: [v3Rule({ id: 'theirs', signal: 'theirs.example' })] })
    expect((await handleEnableSync()).ok).toBe(true)
    expect((await listRules()).map((r) => r.id).sort()).toEqual(['kept-locally', 'theirs'])
  })

  it('a canonical-only collision never deletes a user_manual rule (auto rules follow canonical keys)', async () => {
    await asUpgradedV3Machine()
    const manual = v3Rule({
      id: 'manual-full-form',
      type: 'subject_keyword',
      signal: '112年度訴字第304號',
      source: 'user_manual',
    })
    const auto = v3Rule({ id: 'auto-full-form', type: 'subject_keyword', signal: '112年度訴字第305號' })
    await setRules([manual, auto])
    // Deleted compact auto twins, written in their own (compact) spelling —
    // after the last v3 sync, so the time rule doesn't cover them.
    const compact304 = tomb('112訴304', '2026-10-15T00:00:00.000Z', { type: 'subject_keyword' })
    const compact305 = tomb('112訴305', '2026-10-15T00:00:00.000Z', { type: 'subject_keyword' })
    await simulateV3Push({ tombstones: [compact304, compact305], updatedAt: '2026-10-15T00:00:01.000Z' })
    await pullNow('upgrade', { mode: 'union' })
    const ids = (await listRules()).map((r) => r.id)
    expect(ids).toEqual(['manual-full-form'])
    // Re-stamped above the colliding tombstone: the same tombstone arriving
    // again (next pull, other machines) can't take it either.
    expect((await listRules())[0]!.editedAt).toBeGreaterThan(compact304.deletedAt)
  })

  it('an incoming rule v3 would have added is accepted too (re-stamped); a delete made here still wins', async () => {
    await asUpgradedV3Machine()
    // The other machine kept a live twin next to a deletion both machines
    // processed under v3; this machine never had that twin.
    const twin = v3Rule({ id: 'their-twin', signal: 'twin.example' })
    const processed = tomb('twin.example', '2026-09-15T00:00:00.000Z')
    // A user_manual rule there collides only canonically with an auto twin
    // deleted HERE after the last sync.
    const manual = v3Rule({ id: 'their-manual', type: 'subject_keyword', signal: '113年度上字第88號', source: 'user_manual' })
    const ownCompact = tomb('113上88', '2026-10-20T00:00:00.000Z', { type: 'subject_keyword' })
    // …and a rule deleted here after the last sync (same spelling) stays deleted.
    const gone = v3Rule({ id: 'deleted-here', signal: 'gone.example' })
    const ownDelete = tomb('gone.example', '2026-10-20T00:00:00.000Z')
    await addRuleTombstones([processed, ownCompact, ownDelete])
    await simulateV3Push({ rules: [twin, manual, gone], tombstones: [processed], updatedAt: LAST_V3_SYNC })
    await pullNow('upgrade', { mode: 'union' })
    const rules = await listRules()
    expect(rules.map((r) => r.id).sort()).toEqual(['their-manual', 'their-twin'])
    expect(rules.find((r) => r.id === 'their-twin')!.editedAt).toBeGreaterThan(processed.deletedAt)
    expect(rules.find((r) => r.id === 'their-manual')!.editedAt).toBeGreaterThan(ownCompact.deletedAt)
  })

  it('adoption runs once: later merges apply v4 semantics', async () => {
    await asUpgradedV3Machine()
    await setRules([v3Rule({ id: 'x', signal: 'x.example' })])
    await simulateV3Push({ rules: [], updatedAt: LAST_V3_SYNC })
    await pullNow('upgrade', { mode: 'union' })
    expect((await listRules()).map((r) => r.id)).toEqual(['x'])
    // A local tombstone written now (e.g. a twin deleted here) suppresses it
    // at the next (non-adoption) merge — steady-state v4 semantics.
    await addRuleTombstones([{ type: 'domain', signalNorm: 'x.example', targetFolderPath: '案件/甲', deletedAt: Date.now() }])
    await simulateOtherMachinePush({ rules: [] })
    await pullNow('later', { mode: 'union' })
    expect(await listRules()).toEqual([])
  })
})

describe('A4: an upgraded machine never loses its v3-era edits without a backup', () => {
  it('cloud wins the first merge (no editedAt, no settings base), and the upgrade backup restores rules + tombstones + settings everywhere', async () => {
    await asUpgradedV3Machine()
    await setSettings({ batchSize: 33 })
    // Edited here while stuck on v1.0.7 (never pushed): re-targeted to 乙.
    const mine = v3Rule({ id: 'edited-on-v3', signal: 'client.example', targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' })
    const ownTomb = tomb('removed.example', '2026-09-01T00:00:00.000Z')
    await setRules([mine])
    await addRuleTombstones([ownTomb])
    await simulateV3Push({
      rules: [{ ...mine, targetFolderId: FOLDER_A, targetFolderPath: '案件/甲' }],
      settings: { claudeModel: 'claude-opus-4-7', batchSize: 44 },
      updatedAt: LAST_V3_SYNC,
    })
    await pullNow('upgrade', { mode: 'union' })
    // Policy kept: a v3-era rule has no editedAt, so the tie goes to the cloud,
    // and with no settings base every cloud value wins…
    expect((await listRules())[0]!.targetFolderPath).toBe('案件/甲')
    expect((await getSettings()).batchSize).toBe(44)
    // …but nothing is lost: the pinned snapshot holds all three.
    const upgrade = (await listBackups()).find((b) => b.trigger === 'upgrade')!
    expect(upgrade.payload.rules[0]!.targetFolderPath).toBe('案件/乙')
    expect(upgrade.payload.tombstones).toEqual([ownTomb])
    expect(upgrade.payload.settings.batchSize).toBe(33)

    await pushNow('after-upgrade')
    await restoreBackup(upgrade.snapshotAt)
    const restored = (await listRules())[0]!
    expect(restored.targetFolderPath).toBe('案件/乙')
    expect((await getSettings()).batchSize).toBe(33)
    expect((await getRuleTombstones()).map((t) => t.signalNorm)).toContain('removed.example')
    // The rollback is a new edit: it wins the next merge on every machine.
    expect((await pushNow('after-restore')).pushed).toBe(true)
    expect((await cloudRules())[0]!.targetFolderPath).toBe('案件/乙')
  })
})

describe('A12: participation comes from v4 state, not a stale lastSyncAt', () => {
  it('a machine that switched sync off under v1.0.7 re-enables without applying a wipe issued while it was away', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
    await setSettings({ syncEnabled: false, lastSyncAt: '2026-06-01T00:00:00.000Z' })
    await addRules([newRule({
      type: 'domain',
      signal: 'built-while-off.example',
      targetFolderId: FOLDER_A,
      targetFolderPath: '案件/甲',
      confidence: 0.9,
      source: 'user_manual',
    })])
    await simulateOtherMachinePush({ wipe: { id: 'wipe-while-away', at: '2026-07-01T00:00:00.000Z' } })
    expect((await handleEnableSync()).ok).toBe(true)
    // Before: no v4 state + lastSyncAt → "participant" → the July wipe was applied.
    expect((await listRules()).map((r) => r.signal)).toEqual(['built-while-off.example'])
    expect(await readRemoteWipeNotice()).toBeUndefined()
  })

  it('…but its first merge is still an upgrade: pinned backup, v3-processed tombstones are not news', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
    await setSettings({ syncEnabled: false, lastSyncAt: LAST_V3_SYNC })
    const live = v3Rule({ id: 'kept-by-v3', signal: 'twin.example' })
    await setRules([live])
    await simulateOtherMachinePush({ tombstones: [tomb('twin.example', '2026-09-15T00:00:00.000Z')] })
    expect((await handleEnableSync()).ok).toBe(true)
    expect((await listRules()).map((r) => r.id)).toEqual(['kept-by-v3'])
    expect((await listBackups()).map((b) => b.trigger)).toContain('upgrade')
  })
})
