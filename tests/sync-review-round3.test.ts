// Round-3 review fixes for cross-machine sync (2026-11-06):
//   T1  one tombstone slot per deleted spelling — deleting the second
//       canonical twin used to overwrite the first deletion's spelling, and
//       adoption then resurrected the user's deleted manual twin everywhere
//   T2  a tombstone that records a spelling deletes only that spelling
//       (replaces the S5 twin re-stamp, which could revert a peer's edit and
//       didn't help a twin this machine never pushed)

import { beforeEach, describe, expect, it, type Mock } from 'vitest'
import {
  deleteRule,
  listRules,
  mutateRules,
  newRule,
  reconcileRulesAgainstTree,
  ruleTripleKey,
  toggleRule,
  tombstoneKey,
} from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, setRules, setSettings } from '@/shared/storage'
import {
  buildPushPayload,
  clearCloudState,
  mergeCloudIntoLocal,
  pullNow,
  pushNow,
  resetSyncEngineForTests,
  restoreBackup,
  SYNC_BACKUPS_KEY,
  type SyncBackup,
  type SyncMeta,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import {
  decodeRulePayload,
  decodeTombstone,
  spellingFingerprint,
  type FolderDictEntry,
  type WireRuleEntry,
} from '@/shared/sync-codec'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { MailFolderNode, Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { machines, simulateOtherMachinePush } from './helpers/sync-sim'

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

// The pair dedupe keeps (B9): a disabled user_manual full form + an enabled auto compact form.
const T0 = Date.now() - 3_600_000
const manualFull = (): Rule => ({
  ...rule({ id: 'manual-full', type: 'subject_keyword', signal: '112年度訴字第304號', enabled: false }),
  editedAt: T0,
})
const autoCompact = (over: Partial<Rule> = {}): Rule => ({
  ...rule({ id: 'auto-compact', type: 'subject_keyword', signal: '112訴304', source: 'ai_confirmed' }),
  editedAt: T0,
  ...over,
})
const TRIPLE = { type: 'subject_keyword' as const, signalNorm: '112訴304', targetFolderPath: '案件/甲' }
const H_MANUAL = spellingFingerprint('112年度訴字第304號')
const H_AUTO = spellingFingerprint('112訴304')

async function byId(id: string): Promise<Rule | undefined> {
  return (await listRules()).find((r) => r.id === id)
}

const ids = async () => (await listRules()).map((r) => r.id).sort()

const byDeletedAt = (list: RuleTombstone[]) => [...list].sort((a, b) => a.deletedAt - b.deletedAt)

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

/**
 * A v1.0.8 machine merges the cloud and re-pushes it: tombstones applied to
 * every spelling of their triple, one tombstone per triple, no spellings, no
 * semanticHash / ledger.
 */
async function simulateV108Repush(): Promise<void> {
  const c = await cloud()
  const newest = new Map<string, number>()
  for (const t of c.tombstones) newest.set(tombstoneKey(t), Math.max(newest.get(tombstoneKey(t)) ?? 0, t.deletedAt))
  const kept = c.rules.filter((r) => {
    const d = newest.get(ruleTripleKey(r))
    return d === undefined || d < Math.max(Date.parse(r.createdAt) || 0, r.editedAt ?? 0)
  })
  const oneEach = new Map<string, RuleTombstone>()
  for (const t of c.tombstones) {
    const k = `${t.type}::${t.signalNorm}::${t.targetFolderPath}`
    const cur = oneEach.get(k)
    if (!cur || t.deletedAt > cur.deletedAt) {
      oneEach.set(k, { type: t.type, signalNorm: t.signalNorm, targetFolderPath: t.targetFolderPath, deletedAt: t.deletedAt })
    }
  }
  const payload = await simulateOtherMachinePush({ machineId: 'v108-machine', rules: kept, tombstones: [...oneEach.values()] })
  const { semanticHash: _s, daysHash: _d, ...meta } = payload.meta
  void _s
  void _d
  await chrome.storage.sync.set({ syncMeta: meta satisfies SyncMeta })
}

const syncSet = () => chrome.storage.sync.set as unknown as Mock

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- T1 -------------------------------------------------------------------------

describe('T1: one tombstone slot per deleted spelling', () => {
  it('addRuleTombstones keeps each spelling and the newer deletedAt, and folds a v1.0.8 echo', async () => {
    const manual: RuleTombstone = { ...TRIPLE, deletedAt: 2000, spellingHash: H_MANUAL }
    const auto: RuleTombstone = { ...TRIPLE, deletedAt: 3000, spellingHash: H_AUTO }
    await addRuleTombstones([manual])
    await addRuleTombstones([auto])
    // Before: keyed by the triple alone → [auto]; the manual twin's spelling was gone.
    expect(byDeletedAt(await getRuleTombstones())).toEqual([manual, auto])
    // An older copy arriving later no longer overwrites the newer one.
    await addRuleTombstones([{ ...manual, deletedAt: 1000 }])
    expect(byDeletedAt(await getRuleTombstones())).toEqual([manual, auto])
    // v1.0.8 re-pushes `auto` without its spelling: the same deletion, folded.
    await addRuleTombstones([{ ...TRIPLE, deletedAt: 3000 }])
    expect(byDeletedAt(await getRuleTombstones())).toEqual([manual, auto])
    // A genuine spelling-less deletion (v1.0.8 / an import) is kept as such.
    await addRuleTombstones([{ ...TRIPLE, deletedAt: 4000 }])
    expect(byDeletedAt(await getRuleTombstones())).toEqual([manual, auto, { ...TRIPLE, deletedAt: 4000 }])
  })

  it('B deletes the manual twin, then the auto twin; A re-enables sync: neither comes back', async () => {
    const { handleDisableSync, handleEnableSync } = await import('@/background/handlers/sync')
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    await m.use('A')
    await handleDisableSync({ keepCloud: true })
    await m.use('B')
    await deleteRule('manual-full', { actor: 'user' })
    await deleteRule('auto-compact', { actor: 'user' })
    expect((await getRuleTombstones()).map((t) => t.spellingHash).sort()).toEqual([H_AUTO, H_MANUAL].sort())
    expect((await pushNow('b-delete')).pushed).toBe(true)
    await m.use('A')
    expect((await handleEnableSync()).ok).toBe(true)
    // Before: the one stored tombstone carried 「112訴304」's spelling only, so
    // adoption read the manual twin's own deletion as a canonical collision —
    // kept, re-stamped above it and re-published.
    expect(await ids()).toEqual([])
    await m.use('B')
    syncSet().mockClear()
    await pullNow('b-after', { mode: 'union' })
    expect(await ids()).toEqual([])
    // Both spellings round-trip unchanged: B's follow-up is a no-op.
    expect(await pushNow('b-follow-up')).toMatchObject({ unchanged: true })
    expect(syncSet()).not.toHaveBeenCalled()
  })

  it('a pull keeps a cloud deletion of another spelling, newer or older than the local one', async () => {
    const now = Date.now()
    const manual: RuleTombstone = { ...TRIPLE, deletedAt: now - 5000, spellingHash: H_MANUAL }
    const auto: RuleTombstone = { ...TRIPLE, deletedAt: now - 1000, spellingHash: H_AUTO }
    const third: RuleTombstone = { ...TRIPLE, deletedAt: now - 9000, spellingHash: spellingFingerprint('112年度訴字第304号') }
    await addRuleTombstones([manual])
    await simulateOtherMachinePush({ tombstones: [auto, third] })
    await pullNow('remote', { mode: 'union' })
    // Before: `auto` (newer) overwrote `manual`; `third` (older) was dropped.
    expect(byDeletedAt(await getRuleTombstones())).toEqual([third, manual, auto])
  })

  it('restore keeps a snapshot’s deletion of another spelling', async () => {
    const now = Date.now()
    const auto: RuleTombstone = { ...TRIPLE, deletedAt: now - 1000, spellingHash: H_AUTO }
    const manual: RuleTombstone = { ...TRIPLE, deletedAt: now - 5000, spellingHash: H_MANUAL }
    await addRuleTombstones([auto])
    const { claudeApiKey: _k, ...settings } = DEFAULT_SETTINGS
    void _k
    const snapshotAt = new Date(now - 60_000).toISOString()
    const backup: SyncBackup = {
      snapshotAt,
      direction: 'pre-pull',
      trigger: 'first-enable',
      payload: { rules: [], tombstones: [manual], settings },
    }
    await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: [backup] })
    expect((await restoreBackup(snapshotAt)).restored).toBe(true)
    // Before: newest-per-triple → `manual` (older than `auto`) was skipped.
    expect(byDeletedAt(await getRuleTombstones())).toEqual([manual, auto])
  })

  it('both spellings push identically whatever order they are stored in (no ping-pong)', () => {
    // One bulk delete of both twins → one deletedAt for both.
    const a: RuleTombstone = { ...TRIPLE, deletedAt: 5000, spellingHash: H_AUTO }
    const m: RuleTombstone = { ...TRIPLE, deletedAt: 5000, spellingHash: H_MANUAL }
    const push = (tombstones: RuleTombstone[]) =>
      buildPushPayload({
        machineId: 'x',
        rules: [],
        tombstones,
        folderActivity: [],
        settings: DEFAULT_SETTINGS,
        updatedAt: new Date().toISOString(),
        pushId: 'p00000000001',
        now: Date.now(),
      }).contentHash
    expect(push([a, m])).toBe(push([m, a]))
  })
})

// ---- T2 -------------------------------------------------------------------------

const TREE_WITHOUT_A: MailFolderNode[] = [{ id: FOLDER_B, displayName: '乙', path: '案件/乙', children: [] }]

function caseRule(i: number, over: Partial<Rule> = {}): Rule {
  return rule({
    id: `case-${String(i).padStart(4, '0')}`,
    type: 'compound',
    signal: JSON.stringify({
      conditions: [
        { type: 'domain', value: `firm${i % 40}.com.tw` },
        { type: 'subject_keyword', value: `一一四年度上字第${i}號損害賠償等事件 開庭通知` },
      ],
    }),
    targetFolderId: 'AAMkADAwMDAwMDAwLWJiYmItY2NjYy1kZGRk' + 'z'.repeat(80) + (i % 60),
    targetFolderPath: `01進行中案件/客戶第${i % 60}號`,
    source: 'ai_confirmed',
    matchCount: 1000 - (i % 1000),
    ...over,
  })
}

function cjkTomb(i: number, deletedAt: number): RuleTombstone {
  return {
    type: 'subject_keyword',
    signalNorm: `一一三年度訴字第${i}號請求返還不當得利事件言詞辯論通知書`,
    targetFolderPath: `01進行中案件/已結案/客戶第${i}號`,
    deletedAt,
  }
}

describe('T2: a tombstone with a spelling deletes only that spelling', () => {
  it('(a) a peer’s unsynced edit to the surviving twin is not reverted by the deletion', async () => {
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    // B re-enables the manual twin — not pushed yet …
    await toggleRule('manual-full', true, { actor: 'user' })
    // … while A deletes the auto twin and pushes.
    await m.use('A')
    await deleteRule('auto-compact', { actor: 'user' })
    expect((await pushNow('a-delete')).pushed).toBe(true)
    await m.use('B')
    expect((await pushNow('b-edit')).pushed).toBe(true) // pulls A's deletion first
    // Before: A had re-stamped its untouched copy above the deletion — newer
    // than B's edit — so B's re-enable was reverted.
    expect(await byId('manual-full')).toMatchObject({ enabled: true })
    expect(await byId('auto-compact')).toBeUndefined()
    await m.use('A')
    await pullNow('a-remote', { mode: 'union' })
    expect(await ids()).toEqual(['manual-full'])
    expect((await byId('manual-full'))!.enabled).toBe(true)
  })

  it('(b) the twin survives on the peer when this machine never pushes it — orphaned here', async () => {
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    await m.use('A')
    await mutateRules((rules) => ({ next: reconcileRulesAgainstTree(rules, TREE_WITHOUT_A).rules, result: undefined }))
    await deleteRule('auto-compact', { actor: 'user' })
    expect((await pushNow('a-delete')).pushed).toBe(true)
    expect((await cloud()).rules).toEqual([]) // orphaned rules stay local
    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    // Before: B's copy predates the deletion and was suppressed by the
    // canonical tombstone; A's re-stamped copy never left A.
    expect(await ids()).toEqual(['manual-full'])
  })

  it('(b) …or squeezed out of this machine’s push by a bulk deletion', async () => {
    // Never matched and the oldest: the lowest sync priority there is.
    const m = await twoMachinesSharing([manualFull(), autoCompact({ createdAt: '2020-01-01T00:00:00.000Z', matchCount: 0 })])
    await m.use('A')
    const core = Array.from({ length: 120 }, (_, i) => caseRule(i, { source: 'user_manual' }))
    const neverMatched = Array.from({ length: 400 }, (_, i) => caseRule(3000 + i, { matchCount: 0 }))
    await setRules([...(await listRules()), ...core, ...neverMatched])
    await addRuleTombstones(Array.from({ length: 300 }, (_, i) => cjkTomb(i, Date.now() - (i + 10) * 1000)))
    const { handleClearDisabledRules } = await import('@/background/handlers/rules-bulk')
    expect(await handleClearDisabledRules()).toMatchObject({ ok: true, data: { deleted: 1 } })
    expect((await pushNow('a-clear')).pushed).toBe(true)
    const pushed = await cloud()
    expect(pushed.rules.some((r) => r.id === 'auto-compact')).toBe(false) // left out by the budget
    expect(pushed.tombstones.some((t) => t.spellingHash === H_MANUAL)).toBe(true)
    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    // Before: suppressed by the manual twin's canonical tombstone → gone on B.
    expect(await byId('auto-compact')).toBeDefined()
    expect(await byId('manual-full')).toBeUndefined()
  }, 30_000)

  it('a v1.0.8 machine’s re-push of the deletion (spelling stripped) is not a delete-every-spelling here', async () => {
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    await m.use('A')
    await deleteRule('auto-compact', { actor: 'user' })
    expect((await pushNow('a-delete')).pushed).toBe(true)
    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    expect(await ids()).toEqual(['manual-full'])
    // A v1.0.8 peer applies it to every spelling (drops ITS manual twin) and re-pushes.
    await simulateV108Repush()
    await pullNow('b-v108', { mode: 'union' })
    // Kept as a spelling-less deletion, it would delete the manual twin here too.
    expect(await ids()).toEqual(['manual-full'])
    expect((await getRuleTombstones()).map((t) => t.spellingHash)).toEqual([H_AUTO])
    await m.use('A')
    await pullNow('a-v108', { mode: 'union' })
    expect(await ids()).toEqual(['manual-full'])
    expect((await getRuleTombstones()).map((t) => t.spellingHash)).toEqual([H_AUTO])
  })

  it('a machine that first saw only the stripped copy gets the twin back for good once the spelled deletion arrives', async () => {
    const m = await twoMachinesSharing([manualFull(), autoCompact()])
    await m.use('A')
    await deleteRule('auto-compact', { actor: 'user' })
    expect((await pushNow('a-delete')).pushed).toBe(true)
    await simulateV108Repush()
    // B was offline for A's push: it only ever sees the stripped copy — a
    // delete-every-spelling as far as it can tell.
    await m.use('B')
    await pullNow('b-v108', { mode: 'union' })
    expect(await ids()).toEqual([])
    // A merges the re-push (keeps its twin) and pushes the spelled deletion + the twin.
    await m.use('A')
    expect((await pushNow('a-again')).pushed).toBe(true)
    await m.use('B')
    await pullNow('b-a', { mode: 'union' })
    expect(await ids()).toEqual(['manual-full'])
    expect((await getRuleTombstones()).map((t) => t.spellingHash)).toEqual([H_AUTO])
    // The v1.0.8 peer re-pushes the stripped copy again: B keeps the twin now.
    await simulateV108Repush()
    await pullNow('b-v108-again', { mode: 'union' })
    // Before: B's stored stripped copy "covered" the spelled one, which was
    // never stored — so this re-push deleted the twin again.
    expect(await ids()).toEqual(['manual-full'])
  })

  it('adoption weighs only the deletions that delete the rule: another spelling’s never costs it its protection', () => {
    const old = Date.parse('2026-08-01T00:00:00.000Z')
    const auto = { ...autoCompact({ createdAt: '2026-08-01T00:00:00.000Z' }) }
    delete auto.editedAt
    const res = mergeCloudIntoLocal({
      local: [auto],
      // This machine's own spelling-less tombstone of the triple (it always coexisted with the rule) …
      localTombstones: [{ ...TRIPLE, deletedAt: old + 1000 }],
      cloudRules: [],
      cloudStubs: [],
      // … and another machine's deletion of the OTHER spelling.
      cloudTombstones: [{ ...TRIPLE, deletedAt: old + 2000, spellingHash: H_MANUAL }],
      mode: 'union',
      adoption: {},
    })
    // Before: the other spelling's deletion counted against it → dropped.
    expect(res.next.map((r) => r.id)).toEqual(['auto-compact'])
    expect(res.next[0]!.editedAt).toBeGreaterThan(old + 2000)
  })

  it('control: a spelling-less deletion (v1.0.8, an import) still deletes every spelling', async () => {
    await setRules([manualFull(), autoCompact()])
    await simulateOtherMachinePush({ tombstones: [{ ...TRIPLE, deletedAt: Date.now() }] })
    await pullNow('remote', { mode: 'union' })
    expect(await ids()).toEqual([])
  })
})
