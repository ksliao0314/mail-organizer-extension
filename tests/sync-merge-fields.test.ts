// Per-field merge semantics fixed 2026-11-06:
//   A2  `orphaned` is a per-device observation, not an edit
//   A3  revived counters carry an epoch (statsResetAt) through sync
//   A9  a rule created / re-enabled over a skewed tombstone outranks it
//   A10 duplicate ids are no longer re-stamped on every write
//   A13 an unknown claudeModel from a newer peer is neither adopted nor bounced
//   A14 the stale-sweep active-day ledger is union-merged through sync

import { beforeEach, describe, expect, it } from 'vitest'
import {
  addRules,
  addRulesFilteringTombstones,
  autoDisableStaleRules,
  deleteRule,
  listRules,
  mutateRules,
  newRule,
  reconcileRulesAgainstTree,
  revivedRule,
  stampEditedAt,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import {
  addRuleTombstones,
  getActiveDays,
  getRuleTombstones,
  getSettings,
  recordActiveDay,
  setRules,
  setSettings,
} from '@/shared/storage'
import {
  buildPushPayload,
  clearCloudState,
  mergeCloudIntoLocal,
  pullNow,
  pushNow,
  resetSyncEngineForTests,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import {
  decodeActiveDays,
  decodeRulePayload,
  encodeActiveDays,
  encodeRuleRow,
  type FolderDictEntry,
  type WireRuleEntry,
} from '@/shared/sync-codec'
import { ACTIVE_DAYS_CAP } from '@/shared/constants'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { MailFolderNode, Rule, RuleSource } from '@/shared/types'
import { dailyActiveDays } from './helpers/active-days'
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

async function byId(id: string): Promise<Rule | undefined> {
  return (await listRules()).find((r) => r.id === id)
}

async function cloudRules(): Promise<Rule[]> {
  const cloud = await chrome.storage.sync.get(null)
  const pick = <T,>(prefix: string) =>
    unchunkArray(
      Object.keys(cloud)
        .filter((k) => k.startsWith(prefix))
        .map((k) => cloud[k] as { index: number; items: T[] }),
    )
  return decodeRulePayload({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') })
    .rules
}

/** Folder tree that lacks 案件/甲 (its folder was deleted / is unreachable here). */
const TREE_WITHOUT_A: MailFolderNode[] = [
  { id: FOLDER_B, displayName: '乙', path: '案件/乙', children: [] },
]
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

// ---- A2 -------------------------------------------------------------------------

describe('A2: orphaned is per-device', () => {
  it('orphaning / un-orphaning never stamps editedAt', async () => {
    await addRules([rule({ id: 'o1', signal: 'o1.example' })])
    const before = (await byId('o1'))!.editedAt
    await reconcile(TREE_WITHOUT_A)
    expect((await byId('o1'))!.orphaned).toBe(true)
    expect((await byId('o1'))!.editedAt).toBe(before)
    await reconcile(TREE_WITH_A)
    expect((await byId('o1'))!.orphaned).toBeFalsy()
    expect((await byId('o1'))!.editedAt).toBe(before)
  })

  it("scenario A: B orphans R after A re-targeted it — B adopts A's target and lets reconcile re-check it", async () => {
    await addRules([rule({ id: 'R', signal: 'client.example' })])
    const original = (await byId('R'))!
    // A re-targets R to 乙 …
    const onA: Rule = { ...original, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙', editedAt: original.editedAt! + 1000 }
    // … and B, before pulling that, finds 甲 gone and orphans R (stamped by v1.0.8).
    await setRules([{ ...original, orphaned: true, editedAt: original.editedAt! + 5000 }])
    await simulateOtherMachinePush({ rules: [onA] })
    await pullNow('remote', { mode: 'union' })
    const merged = (await byId('R'))!
    // Before: B's newer orphan stamp won → dead target kept forever (never pushed).
    expect(merged.targetFolderPath).toBe('案件/乙')
    expect(merged.orphaned).toBeFalsy()
  })

  it("scenario B: B's un-orphan no longer re-publishes its stale copy over A's re-target", async () => {
    await addRules([rule({ id: 'R', signal: 'client.example' })])
    const original = (await byId('R'))!
    await reconcile(TREE_WITHOUT_A) // B: orphaned (old target 甲)
    // A re-targeted R right after it was created; B un-orphans a moment later.
    const onA: Rule = { ...original, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙', editedAt: original.editedAt! + 1 }
    await new Promise((r) => setTimeout(r, 5))
    await reconcile(TREE_WITH_A) // B: the folder is back → un-orphan
    await simulateOtherMachinePush({ rules: [onA] })
    expect((await pushNow('local-change')).pushed).toBe(true) // pulls A's edit first
    expect((await byId('R'))!.targetFolderPath).toBe('案件/乙')
    expect((await cloudRules()).find((r) => r.id === 'R')!.targetFolderPath).toBe('案件/乙')
  })

  it('a local copy orphaned at the SAME target as the cloud copy stays orphaned', () => {
    const local: Rule = { ...rule({ id: 's' }), orphaned: true, editedAt: 10 }
    const cloud: Rule = { ...local, orphaned: undefined, enabled: false, editedAt: 20 }
    const [m] = mergeCloudIntoLocal({ local: [local], localTombstones: [], cloudRules: [cloud], cloudStubs: [], cloudTombstones: [], mode: 'union' }).next
    expect(m).toMatchObject({ orphaned: true, enabled: false })
  })
})

// ---- A3 -------------------------------------------------------------------------

describe('A3: a revived rule keeps its fresh counters across machines', () => {
  it('the peer’s pre-reset counters no longer come back and re-disable it', async () => {
    const tired: Rule = {
      ...rule({ id: 'rev', type: 'subject_keyword', signal: '某案件 開庭通知', source: 'ai_confirmed' }),
      enabled: false,
      autoDisabledReason: 'high-error-rate',
      autoDisabledAt: '2026-10-01T00:00:00.000Z',
      matchCount: 10,
      overrideCount: 30,
    }
    await addRules([tired])
    // Fresh evidence the routing is right → learning revives it (counters reset).
    const res = await addRulesFilteringTombstones([{ ...tired, id: 'candidate', enabled: true, matchCount: 0, overrideCount: 0 }])
    expect(res.reEnabled.map((r) => r.id)).toEqual(['rev'])
    const revived = (await byId('rev'))!
    expect(revived).toMatchObject({ enabled: true, matchCount: 1, overrideCount: 0 })
    expect(revived.statsResetAt).toBeGreaterThan(0)
    // The other machine merged the revive with max counters (v1.0.8 logic) and pushes.
    await simulateOtherMachinePush({ rules: [{ ...revived, statsResetAt: undefined, matchCount: 10, overrideCount: 30 }] })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('rev')).toMatchObject({ matchCount: 1, overrideCount: 0, statsResetAt: revived.statsResetAt })
    // Before: max → 10 / 30 again → the next sweep disabled it (and restarted its dormant clock).
    await autoDisableStaleRules({ activeDays: dailyActiveDays(Date.now(), 3) })
    expect((await byId('rev'))!.enabled).toBe(true)
  })

  it('equal epochs still take the max; the newer epoch wins either way round', () => {
    const base = { ...rule({ id: 'e', source: 'ai_confirmed' }), editedAt: 5 }
    const merge = (l: Partial<Rule>, c: Partial<Rule>) =>
      mergeCloudIntoLocal({
        local: [{ ...base, ...l }],
        localTombstones: [],
        cloudRules: [{ ...base, ...c }],
        cloudStubs: [],
        cloudTombstones: [],
        mode: 'union',
      }).next[0]!
    expect(merge({ statsResetAt: 7, matchCount: 2 }, { statsResetAt: 7, matchCount: 9 }).matchCount).toBe(9)
    expect(merge({ statsResetAt: 7, matchCount: 2 }, { statsResetAt: 3, matchCount: 9, overrideCount: 9 })).toMatchObject({
      matchCount: 2,
      statsResetAt: 7,
    })
    expect(merge({ matchCount: 40, overrideCount: 30 }, { statsResetAt: 7, matchCount: 1 })).toMatchObject({
      matchCount: 1,
      statsResetAt: 7,
    })
    expect(merge({ matchCount: 40, overrideCount: 30 }, { statsResetAt: 7, matchCount: 1 }).overrideCount).toBeUndefined()
  })

  it('the epoch rides as an optional row key a v1.0.8 decoder can skip', () => {
    const r = { ...rule({ id: 'w' }), statsResetAt: 1_800_000_000_000 }
    const row = encodeRuleRow(r, 'f1')!
    expect(row.sr).toBe(1_800_000_000_000)
    const folders = [{ x: 'f1', id: r.targetFolderId, p: r.targetFolderPath }]
    expect(decodeRulePayload({ folders, rows: [row] }).rules[0]!.statsResetAt).toBe(1_800_000_000_000)
    // v1.0.8 reads only the keys it knows: the same row minus `sr` is a valid rule.
    const { sr: _sr, ...asV108Sees } = row
    void _sr
    const old = decodeRulePayload({ folders, rows: [asV108Sees] })
    expect(old.skipped).toBe(0)
    expect(old.rules[0]!.statsResetAt).toBeUndefined()
  })

  it('revivedRule opens a newer epoch than any previous one', () => {
    const r = { ...rule(), enabled: false, autoDisabledReason: 'high-error-rate' as const, statsResetAt: 9e12 }
    expect(revivedRule(r, '2026-11-06T00:00:00.000Z').statsResetAt).toBe(9e12 + 1)
  })
})

// ---- A9 -------------------------------------------------------------------------

describe('A9: re-creating over a skewed tombstone', () => {
  it('a rule re-created within the skew of a fast-clock deletion survives the round trip', async () => {
    const fast = Date.now() + 10 * 60_000
    await setRules([{ ...rule({ id: 'old', signal: 'skew.example' }), editedAt: fast }]) // edited on a machine 10 min ahead
    await deleteRule('old', { actor: 'user' })
    const t = (await getRuleTombstones())[0]!
    expect(t.deletedAt).toBeGreaterThan(Date.now())
    await upsertRule(rule({ id: 'again', signal: 'skew.example' }), { actor: 'user', mode: 'create' })
    expect((await byId('again'))!.editedAt).toBeGreaterThan(t.deletedAt)
    // The other machine still carries the tombstone.
    await simulateOtherMachinePush({ tombstones: [t] })
    await pullNow('remote', { mode: 'union' })
    // Before: stamped `now` < deletedAt → suppressed after one round trip.
    expect(await byId('again')).toBeDefined()
  })

  it('a re-enable over such a tombstone is stamped above it too', async () => {
    await setRules([{ ...rule({ id: 'sleeper', signal: 'skew2.example' }), enabled: false, editedAt: 1000 }])
    const future = Date.now() + 10 * 60_000
    await addRuleTombstones([{ type: 'domain', signalNorm: 'skew2.example', targetFolderPath: '案件/甲', deletedAt: future }])
    await toggleRule('sleeper', true, { actor: 'user' })
    expect((await byId('sleeper'))!.editedAt).toBeGreaterThan(future)
  })

  it('an ordinary edit is not pushed above a tombstone (only creation / re-enable is)', async () => {
    await setRules([{ ...rule({ id: 'plain', signal: 'plain.example' }), editedAt: 1000 }])
    const future = Date.now() + 10 * 60_000
    await addRuleTombstones([{ type: 'domain', signalNorm: 'plain.example', targetFolderPath: '案件/甲', deletedAt: future }])
    await upsertRule({ ...(await byId('plain'))!, confidence: 0.6 }, { actor: 'user' })
    expect((await byId('plain'))!.editedAt).toBeLessThan(future)
  })
})

// ---- A10 ------------------------------------------------------------------------

describe('A10: duplicate ids', () => {
  it('an unrelated write leaves both copies of a duplicated id with their own editedAt', async () => {
    const a1 = { ...rule({ id: 'dup', signal: 'dup.example' }), editedAt: 100 }
    const a2 = { ...rule({ id: 'dup', signal: 'dup.example' }), targetFolderId: FOLDER_B, targetFolderPath: '案件/乙', editedAt: 200 }
    const other = { ...rule({ id: 'other', signal: 'other.example' }), editedAt: 300 }
    await setRules([a1, a2, other])
    for (let i = 0; i < 3; i++) await upsertRule({ ...other, confidence: 0.6 + i / 100 }, { actor: 'user' })
    // Before: copy #1 was compared with copy #2 (a "re-target") and re-stamped
    // `now` on every write → permanently the newest version in sync.
    expect((await listRules()).filter((r) => r.id === 'dup').map((r) => r.editedAt)).toEqual([100, 200])
  })

  it('each modified copy is compared with its own occurrence', () => {
    const a1 = { ...rule({ id: 'd' }), editedAt: 100 }
    const a2 = { ...rule({ id: 'd' }), targetFolderPath: '案件/乙', editedAt: 200 }
    // A counter-only change to both copies (bumpRuleHits maps by id).
    const next = stampEditedAt([a1, a2], [{ ...a1, matchCount: 1 }, { ...a2, matchCount: 1 }], 5000)
    expect(next.map((r) => r.editedAt)).toEqual([100, 200])
  })
})

// ---- A13 ------------------------------------------------------------------------

describe('A13: a claudeModel this build does not know', () => {
  it('is neither adopted nor bounced back; a later local choice is pushed', async () => {
    await setSettings({ claudeModel: 'claude-opus-5-5' })
    await simulateOtherMachinePush({ settings: { claudeModel: 'claude-future-9-9' } })
    await pullNow('remote', { mode: 'union' })
    // Before: sanitize turned it into the DEFAULT model locally…
    expect((await getSettings()).claudeModel).toBe('claude-opus-5-5')
    // …and the next push wrote that default over the newer peer's choice.
    const r = await pushNow('local-change')
    expect(r).toMatchObject({ pushed: true, unchanged: true })
    const cloudModel = async () =>
      ((await chrome.storage.sync.get('syncSettings4')).syncSettings4 as { v: { claudeModel: string } }).v.claudeModel
    expect(await cloudModel()).toBe('claude-future-9-9')
    // The user picks another model here: that one goes out.
    await setSettings({ claudeModel: 'claude-sonnet-4-6' })
    expect((await pushNow('model-change')).pushed).toBe(true)
    expect(await cloudModel()).toBe('claude-sonnet-4-6')
  })

  it('a pending local model edit is not overwritten by the foreign value', async () => {
    await simulateOtherMachinePush({ settings: { claudeModel: 'claude-haiku-5-5' } })
    await pullNow('join', { mode: 'union' })
    await setSettings({ claudeModel: 'claude-opus-4-7' }) // edited here, not pushed yet
    await simulateOtherMachinePush({ settings: { claudeModel: 'claude-future-9-9' } })
    await pullNow('remote', { mode: 'union' })
    await pullNow('remote-again', { mode: 'union' })
    expect((await getSettings()).claudeModel).toBe('claude-opus-4-7')
    await pushNow('push-edit')
    expect(
      ((await chrome.storage.sync.get('syncSettings4')).syncSettings4 as { v: { claudeModel: string } }).v.claudeModel,
    ).toBe('claude-opus-4-7')
  })
})

// ---- A14 ------------------------------------------------------------------------

describe('A14: the active-day ledger is union-merged through sync', () => {
  it('rides in the settings item (`d`), which keeps the v1.0.8 shape', async () => {
    await recordActiveDay(Date.parse('2026-11-01T10:00:00Z'))
    await recordActiveDay(Date.parse('2026-11-03T10:00:00Z'))
    await addRules([rule({ signal: 'x.example' })])
    await pushNow('test')
    const item = (await chrome.storage.sync.get('syncSettings4')).syncSettings4 as {
      p: string
      v: Record<string, unknown>
      d: number[]
    }
    expect(typeof item.p).toBe('string')
    expect(item.v.claudeModel).toBeDefined()
    expect(decodeActiveDays(item.d)).toEqual(['2026-11-01', '2026-11-03'])
  })

  it('a pull adds the other machine’s days; the full ledger stays small', async () => {
    const theirs = dailyActiveDays(Date.parse('2026-11-05T12:00:00Z'), ACTIVE_DAYS_CAP)
    await recordActiveDay(Date.parse('2026-11-06T09:00:00Z'))
    const payload = await simulateOtherMachinePush({ activeDays: theirs })
    expect(JSON.stringify((payload.writes['syncSettings4'] as { d: number[] }).d).length).toBeLessThan(1500)
    await pullNow('remote', { mode: 'union' })
    const days = await getActiveDays()
    expect(days.length).toBe(ACTIVE_DAYS_CAP) // cap kept, newest win
    expect(days).toContain('2026-11-06')
    expect(days).toContain('2026-11-05')
  })

  it('two machines reach the same stale-sweep verdict (no-tombstone deletes converge)', async () => {
    const m = machines()
    const born = '2026-06-01T00:00:00.000Z'
    const stale: Rule = { ...rule({ id: 'stale', signal: 'quiet.example', source: 'ai_confirmed' }), createdAt: born, editedAt: Date.parse(born) }
    // A has classified every day since; B only on a few days.
    await m.use('A')
    await setSettings({ syncEnabled: true, syncMachineId: 'machine-a' })
    await setRules([stale])
    for (const d of dailyActiveDays(Date.now(), 120)) await recordActiveDay(Date.parse(`${d}T12:00:00Z`))
    await pushNow('a')
    await m.use('B')
    await setSettings({ syncEnabled: true, syncMachineId: 'machine-b' })
    await setRules([stale])
    for (const d of dailyActiveDays(Date.now(), 5)) await recordActiveDay(Date.parse(`${d}T12:00:00Z`))
    await pullNow('remote', { mode: 'union' })
    // Before: B's ledger had 5 days → B kept the rule A swept, and re-pushed it.
    await autoDisableStaleRules()
    expect(await listRules()).toEqual([])
    await m.use('A')
    await autoDisableStaleRules()
    expect(await listRules()).toEqual([])
  })

  it('a malformed or far-future ledger is ignored', () => {
    expect(decodeActiveDays('nope')).toEqual([])
    expect(decodeActiveDays([20000, 0])).toEqual([])
    expect(decodeActiveDays([20000, -3])).toEqual([])
    const tomorrowPlus = Math.floor(Date.now() / 86_400_000) + 30
    expect(decodeActiveDays([tomorrowPlus])).toEqual([])
    const enc = encodeActiveDays(['2026-11-03', '2026-11-01', 'junk', '2026-11-01'], 10)
    expect(decodeActiveDays(enc)).toEqual(['2026-11-01', '2026-11-03'])
  })

  it('the ledger forces a push only for a day the cloud lacks (no ping-pong over a dropped future day)', async () => {
    const today = new Date().toISOString().slice(0, 10)
    const farFuture = new Date(Date.now() + 30 * 86_400_000).toISOString().slice(0, 10)
    await addRules([rule({ id: 'x', signal: 'x.example' })])
    await pushNow('initial')
    // A peer whose clock runs a month ahead pushes a "future" active day,
    // which this machine's decoder refuses.
    await simulateOtherMachinePush({ rules: [(await byId('x'))!], activeDays: ['2026-10-01', farFuture] })
    await pullNow('remote', { mode: 'union' })
    expect(await getActiveDays()).toEqual(['2026-10-01'])
    // Lacking that day is not news: the follow-up push is a no-op, so the
    // peer is not provoked into re-pushing it forever.
    expect(await pushNow('follow-up')).toMatchObject({ unchanged: true })
    // A day the cloud doesn't know IS news, even with nothing else changed.
    await recordActiveDay(Date.parse(`${today}T12:00:00Z`))
    const r = await pushNow('new-day')
    expect(r.pushed).toBe(true)
    expect(r.unchanged).toBeUndefined()
    const d = ((await chrome.storage.sync.get('syncSettings4')).syncSettings4 as { d: number[] }).d
    expect(decodeActiveDays(d)).toContain(today)
  })

  it('an identical ledger keeps the echo push a no-op', () => {
    const input = {
      machineId: 'a',
      rules: [],
      tombstones: [],
      folderActivity: [],
      settings: DEFAULT_SETTINGS,
      activeDays: ['2026-11-01'],
      updatedAt: new Date().toISOString(),
      now: Date.now(),
    }
    const one = buildPushPayload({ ...input, pushId: 'p00000000001' })
    const two = buildPushPayload({ ...input, pushId: 'p00000000002' })
    expect(one.contentHash).toBe(two.contentHash)
    expect(one.daysHash).toBe(two.daysHash)
    expect(buildPushPayload({ ...input, activeDays: [], pushId: 'p00000000003' }).daysHash).not.toBe(one.daysHash)
  })
})
