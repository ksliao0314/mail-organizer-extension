// Round-2 review fixes for cross-machine sync (2026-11-06):
//   S1  adoption no longer resurrects a user_manual rule deleted under v4
//   S2  a restore never rolls hit counters back (no false high-error disable)
//   S3  a genuine edit to a locally orphaned rule survives the next pull;
//       stubs and full rows treat an orphaned copy alike
//   S4  an edit that moves a rule onto a tombstoned triple outranks it
//   S5  deleting one of two canonical twins keeps the other everywhere
//   S6  a failed enable resets participation
//   S7  a machine enabling sync can repair a permanently torn cloud
//   S8  auto_scan rules age on this machine's own active days
//   S9  folder-activity metric changes are coalesced like hit counters
//   S10 a v1.0.8 peer's echo of our push is not pushed over again

import { beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import {
  addRules,
  autoDisableStaleRules,
  bumpRuleHits,
  deleteRule,
  legacySignalNorm,
  listRules,
  mutateRules,
  newRule,
  reconcileRulesAgainstTree,
  ruleTombstone,
  stampEditedAt,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import {
  addRuleTombstones,
  getActiveDays,
  getLocalActiveDays,
  getRuleTombstones,
  getSettings,
  mergeActiveDaysFromSync,
  recordActiveDay,
  recordFolderActivityFromBatch,
  setRuleCounterBaseline,
  setRules,
  setSettings,
} from '@/shared/storage'
import {
  buildPushPayload,
  clearCloudState,
  listBackups,
  mergeCloudIntoLocal,
  pullNow,
  pushNow,
  quiesce,
  readRemoteWipeNotice,
  resetSyncEngineForTests,
  restoreBackup,
  schedulePush,
  SYNC_BACKUPS_KEY,
  SYNC_LOCAL_STATE_KEY,
  type SyncBackup,
  type SyncMeta,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import {
  contentHash,
  decodeRulePayload,
  decodeTombstone,
  encodeRuleRow,
  encodeTombstone,
  spellingFingerprint,
  type FolderDictEntry,
  type WireRuleEntry,
} from '@/shared/sync-codec'
import { SYNC_TORN_RECOVERY_MIN_MS } from '@/shared/constants'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { MailFolderNode, Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'
import { machines, readCloudMeta, simulateOtherMachinePush, simulateV3Push } from './helpers/sync-sim'

const FOLDER_A = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const FOLDER_B = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdiYmJi'
const LAST_V3_SYNC = '2026-10-01T00:00:00.000Z'

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

/** A rule as v1.0.7 stored it: no editedAt, created in August. */
function v3Rule(over: Partial<Rule> = {}): Rule {
  const r = rule({ createdAt: '2026-08-01T00:00:00.000Z', ...over })
  delete r.editedAt
  return r
}

async function byId(id: string): Promise<Rule | undefined> {
  return (await listRules()).find((r) => r.id === id)
}

const ids = async () => (await listRules()).map((r) => r.id).sort()

async function cloud(): Promise<{ rules: Rule[]; tombstones: RuleTombstone[] }> {
  const all = await chrome.storage.sync.get(null)
  const pick = <T,>(prefix: string) =>
    unchunkArray(
      Object.keys(all)
        .filter((k) => k.startsWith(prefix))
        .map((k) => all[k] as { index: number; items: T[] }),
    )
  return {
    rules: decodeRulePayload({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') })
      .rules,
    tombstones: pick<unknown>('syncTombs4_').map(decodeTombstone).filter((t): t is RuleTombstone => t !== null),
  }
}

/** Every item the cloud meta references carries its pushId. */
async function cloudIntact(): Promise<boolean> {
  const all = await chrome.storage.sync.get(null)
  const meta = await readCloudMeta()
  return Object.keys(all)
    .filter((k) => k !== 'syncMeta')
    .every((k) => (all[k] as { p?: string }).p === meta!.pushId)
}

/** This machine ran v1.0.7 with sync on and last synced at LAST_V3_SYNC. */
async function asUpgradedV3Machine(): Promise<void> {
  await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine', lastSyncAt: LAST_V3_SYNC })
}

/** Machine A pushes `rules`; machine B (current afterwards) joins and merges them. */
async function twoMachinesSharing(rules: Rule[]): Promise<ReturnType<typeof machines>> {
  const m = machines()
  await m.use('A')
  await setSettings({ syncEnabled: true, syncMachineId: 'machine-a' })
  await setRules(rules)
  expect((await pushNow('a-initial')).pushed).toBe(true)
  await m.use('B')
  await setSettings({ syncEnabled: true, syncMachineId: 'machine-b' })
  expect((await pullNow('b-join', { mode: 'union' })).pulled).toBe(true)
  return m
}

const syncSet = () => chrome.storage.sync.set as unknown as Mock

const TREE_WITHOUT_A: MailFolderNode[] = [{ id: FOLDER_B, displayName: '乙', path: '案件/乙', children: [] }]
const TREE_WITH_A: MailFolderNode[] = [
  { id: FOLDER_A, displayName: '甲', path: '案件/甲', children: [] },
  { id: FOLDER_B, displayName: '乙', path: '案件/乙', children: [] },
]

async function reconcile(tree: MailFolderNode[]): Promise<void> {
  await mutateRules((rules) => ({ next: reconcileRulesAgainstTree(rules, tree).rules, result: undefined }))
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- S1 -------------------------------------------------------------------------

describe('S1: adoption never resurrects a user_manual rule deleted under v4', () => {
  const full = (n: number, over: Partial<Rule> = {}) =>
    rule({ id: `manual-${n}`, type: 'subject_keyword', signal: `112年度訴字第${n}號`, ...over })

  it('every new tombstone fingerprints the deleted rule’s own spelling (old readers ignore it)', () => {
    const t = ruleTombstone(full(304))
    expect(t).toMatchObject({ signalNorm: '112訴304', spellingHash: spellingFingerprint('112年度訴字第304號') })
    const row = encodeTombstone(t)!
    expect(decodeTombstone(row)).toEqual(t)
    // What a v1.0.8 decoder sees: the first four entries — the same deletion.
    const { spellingHash: _h, ...v108View } = t
    void _h
    expect(decodeTombstone(row.slice(0, 4))).toEqual(v108View)
    // Spelling = signalNorm → a 2-byte marker.
    const ascii = ruleTombstone(rule({ signal: 'client.example' }))
    expect(encodeTombstone(ascii)![4]).toBe(1)
    expect(decodeTombstone(encodeTombstone(ascii))!.spellingHash).toBe(spellingFingerprint('client.example'))
    // A compound's legacy spelling can be several times its canonical key —
    // only a short fingerprint travels.
    const compound = rule({
      type: 'compound',
      signal: JSON.stringify({
        conditions: [
          { type: 'domain', value: 'Firm.example' },
          { type: 'subject_keyword', value: '一一四年度上字第88號損害賠償等事件  開庭通知' },
        ],
      }),
    })
    const ct = ruleTombstone(compound)
    expect(ct.spellingHash).toBe(spellingFingerprint(legacySignalNorm('compound', compound.signal)))
    const crow = encodeTombstone(ct)!
    expect(typeof crow[4]).toBe('string')
    expect((crow[4] as string).length).toBeLessThan(16)
    expect(decodeTombstone(crow)).toEqual(ct)
    // Anything else in the 5th slot only loses the fingerprint, never the deletion.
    expect(decodeTombstone([...row.slice(0, 4), { junk: true }])).toEqual(v108View)
  })

  it('(a) a v1.0.7 machine upgrading after a v1.0.8+ machine deleted the rule does not bring it back', async () => {
    await asUpgradedV3Machine()
    await setRules([304, 305, 306].map((n) => { const r = full(n, { createdAt: '2026-08-01T00:00:00.000Z' }); delete r.editedAt; return r }))
    await simulateOtherMachinePush({
      tombstones: [
        // v1.0.8 deleted the user's own 304: canonical form, nothing else.
        { type: 'subject_keyword', signalNorm: '112訴304', targetFolderPath: '案件/甲', deletedAt: Date.parse('2026-11-05T09:00:00.000Z') },
        // v1.0.9 deleted 305 on a machine whose clock ran weeks behind: the
        // spelling still says whose deletion it is.
        { ...ruleTombstone(full(305)), deletedAt: Date.parse('2026-10-20T00:00:00.000Z') },
        // v1.0.7 deleted a compact AUTO twin 「112訴306」 — a mere collision with
        // the user's 306: still protected (A1).
        { type: 'subject_keyword', signalNorm: '112訴306', targetFolderPath: '案件/甲', deletedAt: Date.parse('2026-10-20T00:00:00.000Z') },
      ],
    })
    expect((await pullNow('upgrade', { mode: 'union' })).pulled).toBe(true)
    // Before: 304 and 305 were read as collisions too — kept, re-stamped above
    // their own deletions, and re-published to every machine.
    expect(await ids()).toEqual(['manual-306'])
  })

  it('(b) A switches sync off, B deletes the rule, A re-enables: it stays deleted (ASCII control too)', async () => {
    const { handleDisableSync, handleEnableSync } = await import('@/background/handlers/sync')
    const at = Date.now() - 60_000
    const m = await twoMachinesSharing([full(304, { editedAt: at }), rule({ id: 'ascii', signal: 'client.example', editedAt: at })])
    await m.use('A')
    await handleDisableSync({ keepCloud: true })
    await m.use('B')
    await deleteRule('manual-304', { actor: 'user' })
    await deleteRule('ascii', { actor: 'user' })
    expect((await pushNow('b-delete')).pushed).toBe(true)
    await m.use('A')
    expect((await handleEnableSync()).ok).toBe(true)
    expect(await listRules()).toEqual([])
    await m.use('B')
    await pullNow('b-after', { mode: 'union' })
    expect(await listRules()).toEqual([])
  })

  it('(c) A switches sync off, deletes the rule itself, re-enables: deleted everywhere', async () => {
    const { handleDisableSync, handleEnableSync } = await import('@/background/handlers/sync')
    const m = await twoMachinesSharing([full(304, { editedAt: Date.now() - 60_000 })])
    await m.use('A')
    await handleDisableSync({ keepCloud: true })
    await deleteRule('manual-304', { actor: 'user' })
    expect((await handleEnableSync()).ok).toBe(true)
    expect(await listRules()).toEqual([])
    expect((await cloud()).rules).toEqual([])
    await m.use('B')
    await pullNow('b-after', { mode: 'union' })
    expect(await listRules()).toEqual([])
  })

  it('(d) an upgraded machine that deletes the rule before its first sync keeps it deleted', async () => {
    await asUpgradedV3Machine()
    const manual = v3Rule({ id: 'manual-304', type: 'subject_keyword', signal: '112年度訴字第304號' })
    await setRules([manual])
    // The other machine (still v1.0.7) has it.
    await simulateV3Push({ rules: [manual], updatedAt: LAST_V3_SYNC })
    await deleteRule('manual-304', { actor: 'user' })
    expect((await pullNow('upgrade', { mode: 'union' })).pulled).toBe(true)
    expect(await listRules()).toEqual([])
    expect((await pushNow('after')).pushed).toBe(true)
    expect((await cloud()).rules).toEqual([])
  })
})

// ---- S2 -------------------------------------------------------------------------

async function seedBackup(rules: Rule[]): Promise<string> {
  const { claudeApiKey: _k, ...settings } = DEFAULT_SETTINGS
  void _k
  const snapshotAt = new Date(Date.now() - 60_000).toISOString()
  const backup: SyncBackup = { snapshotAt, direction: 'manual', payload: { rules, tombstones: [], settings } }
  await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: [backup] })
  return snapshotAt
}

describe('S2: a restore keeps the current hit counters', () => {
  it('restoring an older snapshot no longer re-judges the legacy record and disables the rule', async () => {
    // A broad rule carrying pre-v1.0.8 override counts; the sweep baselined it at first sight.
    const broad = rule({ id: 'broad', signal: 'client.example', source: 'ai_confirmed', matchCount: 22, overrideCount: 26, statsResetAt: undefined })
    await setRules([broad])
    await setRuleCounterBaseline({ capturedAt: '2026-11-06T00:00:00.000Z', counts: { broad: [22, 26] } })
    // Snapshot from before the last few batches, with another target.
    const at = await seedBackup([{ ...broad, matchCount: 20, overrideCount: 25, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' }])
    expect((await restoreBackup(at)).restored).toBe(true)
    expect(await byId('broad')).toMatchObject({ targetFolderPath: '案件/乙', matchCount: 22, overrideCount: 26 })
    await autoDisableStaleRules({ activeDays: dailyActiveDays(Date.now(), 3) })
    // Before: 20 / 25 sat below the baseline → "reset" → 45 fires, 56 % overridden → disabled.
    expect((await byId('broad'))!.enabled).toBe(true)
  })

  it('a newer counter epoch in the snapshot still wins (same rule as the sync merge)', async () => {
    const r = rule({ id: 'e', source: 'ai_confirmed', matchCount: 40, overrideCount: 30 })
    await setRules([r])
    const at = await seedBackup([{ ...r, confidence: 0.6, matchCount: 2, overrideCount: 0, statsResetAt: Date.now() - 5_000 }])
    await restoreBackup(at)
    expect(await byId('e')).toMatchObject({ confidence: 0.6, matchCount: 2 })
    expect((await byId('e'))!.overrideCount).toBeUndefined()
  })
})

// ---- S3 -------------------------------------------------------------------------

describe('S3: an orphaned rule keeps a genuine local edit', () => {
  it('orphanedAt is derived like editedAt: set on orphaning, kept, dropped on un-orphan', () => {
    const p = { ...rule({ id: 'x' }), editedAt: 100 }
    const orphaned = stampEditedAt([p], [{ ...p, orphaned: true }], 5000)[0]!
    expect(orphaned).toMatchObject({ orphaned: true, orphanedAt: 5000, editedAt: 100 })
    const kept = stampEditedAt([orphaned], [{ ...orphaned, matchCount: 3, orphanedAt: 1 }], 6000)[0]!
    expect(kept.orphanedAt).toBe(5000) // a caller can't forge it
    const back = stampEditedAt([kept], [{ ...kept, orphaned: false }], 7000)[0]!
    expect(back.orphanedAt).toBeUndefined()
    expect(back.editedAt).toBe(100)
    // Orphaned by v1.0.8 (its stamp in editedAt, no orphanedAt): counters don't
    // vouch for it, a real edit does.
    const old = { ...rule({ id: 'y' }), orphaned: true, editedAt: 900 }
    expect(stampEditedAt([old], [{ ...old, matchCount: 1 }], 8000)[0]!.orphanedAt).toBeUndefined()
    expect(stampEditedAt([old], [{ ...old, enabled: false }], 8000)[0]).toMatchObject({ orphanedAt: 900, editedAt: 8000 })
    // Never on the wire.
    expect(JSON.stringify(encodeRuleRow({ ...orphaned }, 'f'))).not.toContain('5000')
  })

  it('the user disables an orphaned rule: the next pull keeps the disable, and the un-orphan publishes it', async () => {
    await addRules([rule({ id: 'R', signal: 'client.example' })])
    await pushNow('initial')
    const peerCopy = (await byId('R'))!
    await reconcile(TREE_WITHOUT_A)
    await toggleRule('R', false, { actor: 'user' })
    const disabled = (await byId('R'))!
    expect(disabled).toMatchObject({ orphaned: true, enabled: false })
    // The other machine pushes its older copy.
    await simulateOtherMachinePush({ rules: [peerCopy] })
    await pullNow('remote', { mode: 'union' })
    // Before: re-enabled from the older cloud copy, under the user's newer stamp.
    expect(await byId('R')).toMatchObject({ enabled: false, orphaned: true, editedAt: disabled.editedAt })
    await reconcile(TREE_WITH_A)
    expect((await pushNow('after')).pushed).toBe(true)
    expect((await cloud()).rules.find((r) => r.id === 'R')!.enabled).toBe(false)
  })

  it('an orphaned copy takes a newer disable stub as it takes a newer full row', () => {
    const merge = (local: Rule, stubEditedAt: number) =>
      mergeCloudIntoLocal({
        local: [local],
        localTombstones: [],
        cloudRules: [],
        cloudStubs: [{ id: local.id, editedAt: stubEditedAt, autoDisabledReason: 'superseded', autoDisabledAt: '2026-11-06T00:00:00.000Z' }],
        cloudTombstones: [],
        mode: 'union',
      }).next[0]!
    const local: Rule = { ...rule({ id: 's' }), orphaned: true, orphanedAt: 5, editedAt: 10 }
    // Before: stubs were ignored for orphaned copies (full rows were applied).
    expect(merge(local, 20)).toMatchObject({ enabled: false, orphaned: true, orphanedAt: 5, autoDisabledReason: 'superseded', editedAt: 20 })
    expect(merge(local, 5)).toBe(local) // an older stub loses to the newer local edit
    // A possible v1.0.8 orphan stamp loses to the stub, as it loses to a full row (A2).
    const suspect: Rule = { ...rule({ id: 'v' }), orphaned: true, editedAt: 99 }
    expect(merge(suspect, 20)).toMatchObject({ enabled: false, orphaned: true })
  })
})

// ---- S4 -------------------------------------------------------------------------

describe('S4: moving a rule onto a tombstoned triple', () => {
  it('a re-target / re-signal over a fast-clock deletion is stamped above it and survives the round trip', async () => {
    await setRules([
      { ...rule({ id: 'resignal', signal: 'old.example' }), editedAt: 1000 },
      { ...rule({ id: 'retarget', signal: 'moved.example' }), editedAt: 1000 },
    ])
    const future = Date.now() + 10 * 60_000
    const tombs: RuleTombstone[] = [
      { type: 'domain', signalNorm: 'new.example', targetFolderPath: '案件/甲', deletedAt: future },
      { type: 'domain', signalNorm: 'moved.example', targetFolderPath: '案件/乙', deletedAt: future },
    ]
    await addRuleTombstones(tombs)
    await upsertRule({ ...(await byId('resignal'))!, signal: 'new.example' }, { actor: 'user' })
    await upsertRule({ ...(await byId('retarget'))!, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' }, { actor: 'user' })
    expect((await byId('resignal'))!.editedAt).toBeGreaterThan(future)
    expect((await byId('retarget'))!.editedAt).toBeGreaterThan(future)
    // A peer whose clock runs ahead still holds both tombstones.
    await simulateOtherMachinePush({ tombstones: tombs })
    await pullNow('remote', { mode: 'union' })
    // Before: stamped max(now, prev + 1) < deletedAt → both suppressed, gone everywhere.
    expect(await ids()).toEqual(['resignal', 'retarget'])
  })
})

// ---- S5 -------------------------------------------------------------------------

describe('S5: deleting one of two canonical twins keeps the other everywhere', () => {
  // The B9 pair dedupe keeps: a disabled user_manual full form + an enabled auto compact form.
  const T0 = Date.now() - 3_600_000
  const manualFull = (): Rule => ({
    ...rule({ id: 'manual-full', type: 'subject_keyword', signal: '112年度訴字第304號', enabled: false }),
    editedAt: T0,
  })
  const autoCompact = (): Rule => ({
    ...rule({ id: 'auto-compact', type: 'subject_keyword', signal: '112訴304', source: 'ai_confirmed' }),
    editedAt: T0,
  })

  it('deleting the auto twin keeps the user_manual twin — here, on the peer, with no ping-pong', async () => {
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    expect(await pushNow('b-echo')).toMatchObject({ unchanged: true })
    await m.use('A')
    syncSet().mockClear()
    await deleteRule('auto-compact', { actor: 'user' })
    // Round 3: no re-stamp — the tombstone records the deleted spelling and
    // never suppresses the twin (a re-stamp could revert a peer's edit).
    expect((await getRuleTombstones())[0]!.spellingHash).toBe(spellingFingerprint('112訴304'))
    expect((await byId('manual-full'))!.editedAt).toBe(T0)
    expect((await pushNow('a-delete')).pushed).toBe(true)
    expect(syncSet()).toHaveBeenCalledTimes(1)
    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    // Before: the canonical tombstone suppressed the user_manual twin as well (invariant 3).
    expect(await ids()).toEqual(['manual-full'])
    expect(await pushNow('b-follow-up')).toMatchObject({ unchanged: true })
    expect(syncSet()).toHaveBeenCalledTimes(1)
    // A's own later merges keep it too.
    await m.use('A')
    await pullNow('a-later', { mode: 'union' })
    expect(await ids()).toEqual(['manual-full'])
  })

  it('「清除已停用規則」 deleting the user_manual twin keeps the enabled auto twin everywhere', async () => {
    const { handleClearDisabledRules } = await import('@/background/handlers/rules-bulk')
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    await m.use('A')
    expect(await handleClearDisabledRules()).toMatchObject({ ok: true, data: { deleted: 1 } })
    expect(await ids()).toEqual(['auto-compact'])
    expect((await pushNow('a-clear')).pushed).toBe(true)
    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    expect(await ids()).toEqual(['auto-compact'])
    expect(await pushNow('b-follow-up')).toMatchObject({ unchanged: true })
    await m.use('A')
    await pullNow('a-later', { mode: 'union' })
    expect(await ids()).toEqual(['auto-compact'])
  })

  it('a delete with no surviving twin re-stamps nothing', async () => {
    await setRules([{ ...rule({ id: 'solo', signal: 'solo.example' }), editedAt: 5 }, { ...rule({ id: 'other', signal: 'other.example' }), editedAt: 6 }])
    await deleteRule('solo', { actor: 'user' })
    expect((await byId('other'))!.editedAt).toBe(6)
  })
})

// ---- S6 -------------------------------------------------------------------------

describe('S6: a failed enable resets participation', () => {
  it('a wipe issued before the next enable is recorded, not applied', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await setSettings({ syncEnabled: false })
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.example' })] })
    // The union pull succeeds; the push-back fails.
    syncSet().mockImplementationOnce(async () => {
      throw new Error('network down')
    })
    expect((await handleEnableSync()).ok).toBe(false)
    expect((await getSettings()).syncEnabled).toBe(false)
    const state = (await chrome.storage.local.get(SYNC_LOCAL_STATE_KEY))[SYNC_LOCAL_STATE_KEY] as { participating: boolean }
    expect(state.participating).toBe(false)
    // While sync is off here, the other machine wipes.
    await simulateOtherMachinePush({ wipe: { id: 'wipe-between' }, rules: [] })
    expect((await handleEnableSync()).ok).toBe(true)
    // Before: still "participating" after the failed enable → the wipe was applied.
    expect(await ids()).toEqual(['mine', 'theirs'])
    expect(await readRemoteWipeNotice()).toBeUndefined()
  })

  it('an upgraded v1.0.7 machine whose first enable fails still upgrades on the next one', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await chrome.storage.local.remove(SYNC_LOCAL_STATE_KEY)
    await setSettings({ syncEnabled: false, lastSyncAt: LAST_V3_SYNC })
    await setRules([v3Rule({ id: 'kept-by-v3', signal: 'twin.example' })])
    // A twin deleted before the last v3 sync — v3 kept this one.
    const processed: RuleTombstone = { type: 'domain', signalNorm: 'twin.example', targetFolderPath: '案件/甲', deletedAt: Date.parse('2026-09-15T00:00:00.000Z') }
    const payload = await simulateOtherMachinePush({ tombstones: [processed] })
    // First attempt: the cloud is mid-push (torn) → fails, nothing merged.
    await chrome.storage.sync.set({ syncTombs4_0: { ...(payload.writes['syncTombs4_0'] as object), p: 'ghostpush004' } })
    expect((await handleEnableSync()).ok).toBe(false)
    // The other machine finishes its push; the next enable is the upgrade merge.
    await simulateOtherMachinePush({ tombstones: [processed] })
    expect((await handleEnableSync()).ok).toBe(true)
    // Before: the rollback dropped the v1.0.7 lastSyncAt → the processed
    // tombstone counted as news (rule deleted) and no upgrade backup was taken.
    expect(await ids()).toEqual(['kept-by-v3'])
    expect((await listBackups()).map((b) => b.trigger)).toContain('upgrade')
  }, 20_000)
})

// ---- S7 -------------------------------------------------------------------------

describe('S7: enabling sync against a permanently torn cloud', () => {
  it('repairs it once the torn-cloud recovery is allowed — backup first, carried wipe recorded not applied', async () => {
    const { handleEnableSync } = await import('@/background/handlers/sync')
    await setSettings({ syncEnabled: false })
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    const payload = await simulateOtherMachinePush({
      wipe: { id: 'old-wipe', at: '2026-10-01T00:00:00.000Z' },
      rules: [rule({ id: 'theirs', signal: 'theirs.example' })],
    })
    // One row chunk belongs to another push, and stays that way (its owner is gone).
    await chrome.storage.sync.set({ syncRows4_0: { ...(payload.writes['syncRows4_0'] as object), p: 'ghostpush003' } })

    const first = await handleEnableSync()
    // Within the first minute it may still be a push in flight.
    expect(first.ok).toBe(false)
    expect((first as { message?: string }).message).toContain('上傳中')
    expect((await getSettings()).syncEnabled).toBe(false)
    // The failed enable kept the torn-read streak (S6's reset must not clear it).
    const streak = (await chrome.storage.local.get('syncTornStreak')).syncTornStreak as { firstAt: number; count: number }
    expect(streak.count).toBeGreaterThanOrEqual(3)
    await chrome.storage.local.set({ syncTornStreak: { ...streak, firstAt: streak.firstAt - SYNC_TORN_RECOVERY_MIN_MS } })

    // Before: every enable failed — no machine could ever join this cloud.
    const second = await handleEnableSync()
    expect(second.ok).toBe(true)
    expect((await listBackups()).some((b) => b.trigger === 'torn-recovery')).toBe(true)
    expect((await readCloudMeta())!.sourceMachineId).toBe('this-machine')
    expect(await cloudIntact()).toBe(true)
    expect((await cloud()).rules.map((r) => r.id)).toEqual(['mine'])

    // The wipe marker it carried forward is old news here (F1): the other
    // machine's next push carries it again, and nothing is wiped.
    await simulateOtherMachinePush({ rules: [] })
    await pullNow('remote', { mode: 'union' })
    expect(await ids()).toEqual(['mine'])
    expect(await readRemoteWipeNotice()).toBeUndefined()
  }, 20_000)

  it('if the cloud heals meanwhile, that push merges only after a first-enable backup', async () => {
    // Not yet participating (sync just switched on); another machine's cloud.
    await addRules([rule({ id: 'mine', signal: 'mine.example' })])
    await addRuleTombstones([{ type: 'domain', signalNorm: 'old.example', targetFolderPath: '案件/甲', deletedAt: 1000 }])
    await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.example' })] })
    expect((await pushNow('first-enable-torn-recovery')).pushed).toBe(true)
    const backup = (await listBackups()).find((b) => b.trigger === 'first-enable')
    expect(backup?.payload.rules.map((r) => r.id)).toEqual(['mine'])
    expect(backup?.payload.tombstones.map((t) => t.signalNorm)).toEqual(['old.example'])
    expect(await ids()).toEqual(['mine', 'theirs'])
    // Once participating, pushes take no backup.
    await upsertRule({ ...(await byId('mine'))!, confidence: 0.6 }, { actor: 'user' })
    await pushNow('later')
    expect((await listBackups()).filter((b) => b.trigger === 'first-enable')).toHaveLength(1)
  })
})

// ---- S8 -------------------------------------------------------------------------

describe('S8: per-device auto_scan rules age on this machine’s own clock', () => {
  it('local and fleet ledgers are kept apart; the push ledger is their union', async () => {
    await recordActiveDay(Date.parse('2026-06-01T12:00:00Z'))
    await mergeActiveDaysFromSync(['2026-06-01', '2026-07-01', '2026-08-01'])
    expect(await getLocalActiveDays()).toEqual(['2026-06-01'])
    expect(await getActiveDays()).toEqual(['2026-06-01', '2026-07-01', '2026-08-01'])
    // A batch here never reads the fleet days into the local ledger.
    await recordActiveDay(Date.parse('2026-09-01T12:00:00Z'))
    expect(await getLocalActiveDays()).toEqual(['2026-06-01', '2026-09-01'])
    // Days already known write nothing.
    const set = chrome.storage.local.set as unknown as Mock
    set.mockClear()
    await mergeActiveDaysFromSync(['2026-06-01', '2026-09-01', '2026-08-01'])
    expect(set).not.toHaveBeenCalled()
  })

  it('a machine idle since June keeps its auto_scan rules; the synced rule follows the fleet clock', async () => {
    const lastUsed = '2026-06-01T12:00:00.000Z'
    const born = '2026-05-01T00:00:00.000Z'
    const scan = { ...rule({ id: 'scan', signal: 'scan.example', source: 'auto_scan' }), createdAt: born, matchCount: 3, lastUsedAt: lastUsed }
    const learned = { ...rule({ id: 'learned', signal: 'learned.example', source: 'ai_confirmed' }), createdAt: born, editedAt: Date.parse(born), matchCount: 3, lastUsedAt: lastUsed }
    await setRules([scan, learned])
    for (const d of ['2026-05-30', '2026-05-31', '2026-06-01']) await recordActiveDay(Date.parse(`${d}T12:00:00Z`))
    // The busy machine classified every day since.
    await simulateOtherMachinePush({ rules: [learned], activeDays: dailyActiveDays(Date.now(), 150) })
    await pullNow('remote', { mode: 'union' })
    await autoDisableStaleRules()
    // Before: the merged ledger aged the auto_scan rule 150 days too — deleted on return.
    expect(await ids()).toEqual(['scan'])
  })
})

// ---- S9 -------------------------------------------------------------------------

async function localState(): Promise<{ counterPushDueAt?: number } | undefined> {
  return (await chrome.storage.local.get(SYNC_LOCAL_STATE_KEY))[SYNC_LOCAL_STATE_KEY] as
    | { counterPushDueAt?: number }
    | undefined
}

describe('S9: folder-activity metrics are coalesced like hit counters', () => {
  it('a batch into folders already in the synced list waits; a new folder pushes at once', async () => {
    await addRules([rule({ id: 'c', signal: 'c.example', source: 'ai_confirmed' })])
    await recordFolderActivityFromBatch([{ folderId: FOLDER_A, folderPath: '案件/甲', count: 1 }], new Date(Date.now() - 60_000).toISOString())
    await pushNow('initial')
    syncSet().mockClear()
    vi.useFakeTimers()
    try {
      await bumpRuleHits(new Map([['c', 2]]))
      await recordFolderActivityFromBatch(
        [{ folderId: FOLDER_A, folderPath: '案件/甲', count: 2, latestMessage: { subject: '開庭通知', from: 'court@example.gov', receivedAt: new Date().toISOString() } }],
        new Date().toISOString(),
      )
      schedulePush('local-mutation', 0)
      await vi.advanceTimersByTimeAsync(50)
      // Before: lastActiveAt / recentCount / lastBatchAt were semantic → ~88 KB per batch.
      expect(syncSet()).not.toHaveBeenCalled()
      expect((await localState())?.counterPushDueAt).toBeGreaterThan(Date.now())

      await recordFolderActivityFromBatch([{ folderId: FOLDER_B, folderPath: '案件/乙', count: 1 }], new Date().toISOString())
      schedulePush('local-mutation', 0)
      await vi.advanceTimersByTimeAsync(50)
      expect(syncSet()).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    const activity = unchunkArray(
      Object.entries(await chrome.storage.sync.get(null))
        .filter(([k]) => k.startsWith('syncActivity4_'))
        .map(([, v]) => v as { index: number; items: Array<{ folderId: string; recentCount: number }> }),
    )
    // The prompt push carried the pending metrics along.
    expect(activity.find((e) => e.folderId === FOLDER_A)!.recentCount).toBe(2)
    expect((await localState())?.counterPushDueAt).toBeUndefined()
  })
})

// ---- S10 ------------------------------------------------------------------------

/** A v1.0.8 machine re-pushes what it merged: rows without `sr`, 4-entry tombstones, no semanticHash / daysHash / ledger. */
async function simulateV108Echo(pushId: string): Promise<void> {
  const all = await chrome.storage.sync.get(null)
  const meta = all.syncMeta as SyncMeta
  const writes: Record<string, unknown> = {}
  const restamp = <T,>(k: string, map: (items: T[]) => T[]) => {
    const c = all[k] as { index: number; items: T[] }
    writes[k] = { index: c.index, items: map(c.items), p: pushId }
  }
  const keys = Object.keys(all).filter((k) => k !== 'syncMeta')
  for (const k of keys) {
    if (k === 'syncSettings4') writes[k] = { p: pushId, v: (all[k] as { v: unknown }).v }
    else if (k.startsWith('syncRows4_'))
      restamp<WireRuleEntry>(k, (items) =>
        items.map((e) => {
          if (Array.isArray(e)) return e
          const { sr: _sr, ...rest } = e
          void _sr
          return rest
        }),
      )
    else if (k.startsWith('syncTombs4_')) restamp<unknown[]>(k, (items) => items.map((t) => t.slice(0, 4)))
    else restamp<unknown>(k, (items) => items)
  }
  const flat = (prefix: string) =>
    unchunkArray(keys.filter((k) => k.startsWith(prefix)).map((k) => writes[k] as { index: number; items: unknown[] }))
  const hash = contentHash(
    JSON.stringify([
      flat('syncRows4_'),
      flat('syncFolders4_'),
      flat('syncTombs4_'),
      flat('syncActivity4_'),
      (writes['syncSettings4'] as { v: unknown }).v,
    ]),
  )
  const { semanticHash: _s, daysHash: _d, ...rest } = meta
  void _s
  void _d
  writes.syncMeta = { ...rest, sourceMachineId: 'v108-machine', pushId, contentHash: hash, updatedAt: new Date().toISOString() }
  await chrome.storage.sync.set(writes)
}

describe('S10: a v1.0.8 peer’s echo is not pushed over again', () => {
  it('differs only in what v1.0.8 drops → no-op; a real change still pushes', async () => {
    await setRules([
      { ...rule({ id: 'rev', signal: 'rev.example', source: 'ai_confirmed', matchCount: 1 }), statsResetAt: Date.now() - 1000, editedAt: 5 },
      { ...rule({ id: 'gone', type: 'subject_keyword', signal: '112年度訴字第7號' }), editedAt: 6 },
    ])
    await deleteRule('gone', { actor: 'user' })
    expect((await pushNow('mine')).pushed).toBe(true)
    await simulateV108Echo('v108push0001')
    await pullNow('remote', { mode: 'union' })
    syncSet().mockClear()
    // Before: the echo's contentHash differed (no `sr`, no spelling) → a full re-push, which the v1.0.8 peer echoed again.
    expect(await pushNow('follow-up')).toMatchObject({ unchanged: true })
    expect(syncSet()).not.toHaveBeenCalled()
    // Local keeps the v1.0.9 fields.
    expect((await byId('rev'))!.statsResetAt).toBeGreaterThan(0)
    expect((await getRuleTombstones())[0]!.spellingHash).toBe(spellingFingerprint('112年度訴字第7號'))
    await upsertRule({ ...(await byId('rev'))!, confidence: 0.6 }, { actor: 'user' })
    expect((await pushNow('edit')).pushed).toBe(true)
    expect(syncSet()).toHaveBeenCalledTimes(1)
  })

  it('the v1.0.8-equivalent fingerprint equals contentHash when nothing v1.0.9-only is pushed', () => {
    const input = {
      machineId: 'a',
      rules: [rule({ id: 'plain', signal: 'plain.example' })],
      tombstones: [],
      folderActivity: [],
      settings: DEFAULT_SETTINGS,
      updatedAt: new Date().toISOString(),
      pushId: 'p00000000001',
      now: Date.now(),
    }
    const p = buildPushPayload(input)
    expect(p.v108ContentHash).toBe(p.contentHash)
    const withSr = buildPushPayload({ ...input, rules: [{ ...input.rules[0]!, statsResetAt: 1_800_000_000_000 }] })
    expect(withSr.v108ContentHash).toBe(p.contentHash)
    expect(withSr.contentHash).not.toBe(p.contentHash)
  })
})
