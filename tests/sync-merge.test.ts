// Cross-machine merge semantics (2026-11-05): editedAt stamping in
// mutateRules, per-id merge by editedAt, tombstone timing, disabled stubs,
// one sync queue + pull-before-push, deferred local changes, and restore
// through mutateRules. See sync-engine.ts header for the model.

import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  addRules,
  bumpRuleHits,
  bumpRuleOverrides,
  deleteRule,
  listRules,
  mutateRules,
  newRule,
  stampEditedAt,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import { addRuleTombstones, getRuleTombstones, getSettings, setSettings } from '@/shared/storage'
import {
  buildPushPayload,
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
import { decodeRulesFromSync, type FolderDictEntry, type WireRuleEntry } from '@/shared/sync-codec'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { Rule, RuleSource } from '@/shared/types'
import { readCloudMeta, simulateOtherMachinePush, syncListener } from './helpers/sync-sim'

const FOLDER_A = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const FOLDER_B = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdiYmJi'

function rule(over: Partial<Rule> = {}): Rule {
  // newRule() ignores id / counters — apply overrides on the result.
  return {
    ...newRule({
      type: 'domain',
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
  return decodeRulesFromSync({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') })
}

beforeEach(async () => {
  resetSyncEngineForTests()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- mutateRules stamps editedAt ------------------------------------------

describe('Item 2: mutateRules stamps editedAt on semantic edits only', () => {
  it('a new rule is stamped; a re-target stamps a newer editedAt', async () => {
    const r = rule({ id: 'r1' })
    await addRules([r])
    const created = (await byId('r1'))!
    expect(created.editedAt).toBeGreaterThan(0)
    await upsertRule({ ...created, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' }, { actor: 'user' })
    expect((await byId('r1'))!.editedAt).toBeGreaterThan(created.editedAt!)
  })

  it('hit counters never stamp it (bumpRuleHits without promotion, bumpRuleOverrides)', async () => {
    await addRules([rule({ id: 'c1', source: 'ai_confirmed' })])
    const before = (await byId('c1'))!.editedAt
    await bumpRuleHits(new Map([['c1', 3]]))
    await bumpRuleOverrides(new Map([['c1', 1]]))
    const after = (await byId('c1'))!
    expect(after.matchCount).toBe(3)
    expect(after.overrideCount).toBe(1)
    expect(after.editedAt).toBe(before)
  })

  it('auto-promotion (confidence up together with matchCount) is not an edit', async () => {
    await addRules([
      rule({ id: 'p1', type: 'subject_keyword', signal: '某案件進度', source: 'ai_confirmed', confidence: 0.7 }),
    ])
    const before = (await byId('p1'))!.editedAt
    await bumpRuleHits(new Map([['p1', 20]]))
    const after = (await byId('p1'))!
    expect(after.confidence).toBeCloseTo(0.75)
    // A promotion that stamped editedAt would let a machine that merely matched
    // mail offline out-date (and revert) a re-target made elsewhere.
    expect(after.editedAt).toBe(before)
  })

  it('disabling (toggle) is an edit', async () => {
    await addRules([rule({ id: 't1' })])
    const before = (await byId('t1'))!.editedAt!
    await toggleRule('t1', false, { actor: 'user' })
    expect((await byId('t1'))!.editedAt).toBeGreaterThan(before)
  })

  it('a stale rule object written back can neither forge nor erase editedAt', async () => {
    await addRules([rule({ id: 's1' })])
    const stored = (await byId('s1'))!
    const { editedAt: _e, ...withoutEditedAt } = stored
    void _e
    await upsertRule(withoutEditedAt as Rule) // e.g. a UI round-trip, nothing changed
    expect((await byId('s1'))!.editedAt).toBe(stored.editedAt)
    await upsertRule({ ...stored, editedAt: stored.editedAt! + 10 ** 9 })
    expect((await byId('s1'))!.editedAt).toBe(stored.editedAt)
  })

  it('stays monotonic when the stored version came from a clock running ahead', () => {
    const now = Date.now()
    const prev = { ...rule({ id: 'm1' }), editedAt: now + 600_000 }
    const [edited] = stampEditedAt([prev], [{ ...prev, targetFolderPath: '案件/乙' }], now)
    expect(edited!.editedAt).toBe(prev.editedAt + 1)
  })

  it('preserveEditedAt writes the incoming editedAt untouched', async () => {
    await mutateRules(() => ({ next: [{ ...rule({ id: 'x1' }), editedAt: 1234 }], result: undefined }), {
      preserveEditedAt: true,
    })
    expect((await byId('x1'))!.editedAt).toBe(1234)
  })
})

// ---- per-id merge -----------------------------------------------------------

describe('Item 2: pull merges per rule id by editedAt', () => {
  it('Bug repro: a stale cloud copy no longer reverts a newer local re-target', async () => {
    // Other machine still has the OLD target; this machine re-targeted later.
    const original = rule({ id: 'case-1', signal: 'kgi.com' })
    await addRules([original])
    const stale = { ...(await byId('case-1'))! }
    await upsertRule({ ...stale, targetFolderId: FOLDER_B, targetFolderPath: '案件/乙' }, { actor: 'user' })
    await simulateOtherMachinePush({ rules: [{ ...stale, matchCount: 7 }] })
    await pullNow('remote', { mode: 'union' })
    const merged = (await byId('case-1'))!
    // Before: the cloud copy replaced ours wholesale → mail filed into 案件/甲.
    expect(merged.targetFolderPath).toBe('案件/乙')
    expect(merged.targetFolderId).toBe(FOLDER_B)
    expect(merged.matchCount).toBe(7) // counters still merge (max)
  })

  it('a newer edit from the other machine wins', async () => {
    await addRules([rule({ id: 'n1' })])
    const local = (await byId('n1'))!
    await simulateOtherMachinePush({
      rules: [{ ...local, targetFolderPath: '案件/乙', editedAt: local.editedAt! + 5000 }],
    })
    await pullNow('remote', { mode: 'union' })
    expect((await byId('n1'))!.targetFolderPath).toBe('案件/乙')
  })

  it('pre-v1.0.8 rules (no editedAt on either side) tie → cloud wins, as before', () => {
    const local = { ...rule({ id: 'l1' }), editedAt: undefined }
    const cloud = { ...local, targetFolderPath: '案件/乙' }
    const out = mergeCloudIntoLocal({
      local: [local],
      localTombstones: [],
      cloudRules: [cloud],
      cloudStubs: [],
      cloudTombstones: [],
      mode: 'union',
    })
    expect(out.next[0]!.targetFolderPath).toBe('案件/乙')
  })

  it('counters take the max, lastUsedAt the latest, createdAt the earliest', () => {
    const t = (d: number) => new Date(Date.UTC(2026, 0, d)).toISOString()
    const local: Rule = {
      ...rule({ id: 'k1' }),
      matchCount: 9,
      overrideCount: 1,
      lastUsedAt: t(20),
      createdAt: t(5),
      editedAt: 100,
    }
    const cloud: Rule = { ...local, matchCount: 4, overrideCount: 3, lastUsedAt: t(10), createdAt: t(2), editedAt: 50 }
    const [m] = mergeCloudIntoLocal({
      local: [local],
      localTombstones: [],
      cloudRules: [cloud],
      cloudStubs: [],
      cloudTombstones: [],
      mode: 'union',
    }).next
    expect(m).toMatchObject({ matchCount: 9, overrideCount: 3, lastUsedAt: t(20), createdAt: t(2), editedAt: 100 })
  })

  it('on a tie an auto rule keeps the higher (promoted) confidence', () => {
    const local: Rule = { ...rule({ id: 'q1', source: 'ai_confirmed', type: 'subject_keyword', signal: 'x' }), confidence: 0.8, editedAt: 10 }
    const cloud: Rule = { ...local, confidence: 0.75 }
    const [m] = mergeCloudIntoLocal({
      local: [local],
      localTombstones: [],
      cloudRules: [cloud],
      cloudStubs: [],
      cloudTombstones: [],
      mode: 'union',
    }).next
    expect(m!.confidence).toBe(0.8)
  })

  it('an unchanged merge returns the SAME array (no rules write, no echo push)', () => {
    const local = [{ ...rule({ id: 'u1' }), editedAt: 5 }]
    const out = mergeCloudIntoLocal({
      local,
      localTombstones: [],
      cloudRules: [{ ...local[0]! }],
      cloudStubs: [],
      cloudTombstones: [],
      mode: 'union',
    })
    expect(out.next).toBe(local)
  })
})

// ---- tombstone timing -------------------------------------------------------

describe('Item 3: a tombstone only suppresses versions older than the deletion', () => {
  it('(a) an incoming cloud rule older than a LOCAL tombstone is not re-added nor re-published', async () => {
    await addRules([rule({ id: 'd1', signal: 'deleted-here.com' })])
    const old = (await byId('d1'))!
    await deleteRule('d1', { actor: 'user' })
    // The other machine hasn't seen the deletion yet and still pushes the rule.
    await simulateOtherMachinePush({ rules: [old] })
    await pullNow('remote', { mode: 'union' })
    // Before: "cloud rules not in local → always add" resurrected it.
    expect(await byId('d1')).toBeUndefined()
    await pushNow('after')
    expect((await cloudRules()).find((r) => r.id === 'd1')).toBeUndefined()
  })

  it('(b) a rule re-created after an old deletion survives that tombstone', async () => {
    const monthAgo = Date.now() - 30 * 86_400_000
    await addRules([rule({ id: 'again', signal: 'recreated.com' })]) // created now
    await simulateOtherMachinePush({
      tombstones: [{ type: 'domain', signalNorm: 'recreated.com', targetFolderPath: '案件/甲', deletedAt: monthAgo }],
    })
    await pullNow('remote', { mode: 'union' })
    // Before: the month-old tombstone deleted the re-created rule.
    expect(await byId('again')).toBeDefined()
  })

  it('an edit made after the deletion beats it (incoming)', async () => {
    const deletedAt = Date.now() - 60_000
    await addRuleTombstones([{ type: 'domain', signalNorm: 'edited-later.com', targetFolderPath: '案件/甲', deletedAt }])
    const r = { ...rule({ id: 'e1', signal: 'edited-later.com' }), createdAt: new Date(deletedAt - 86_400_000).toISOString(), editedAt: deletedAt + 1000 }
    await simulateOtherMachinePush({ rules: [r] })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('e1')).toBeDefined()
  })

  it('a cloud deletion newer than the local version still drops it (delete propagates)', async () => {
    await addRules([rule({ id: 'g1', signal: 'gone.com' })])
    const local = (await byId('g1'))!
    await simulateOtherMachinePush({
      tombstones: [{ type: 'domain', signalNorm: 'gone.com', targetFolderPath: '案件/甲', deletedAt: local.editedAt! + 1 }],
    })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('g1')).toBeUndefined()
  })

  it('an older cloud tombstone never overwrites a newer local one', async () => {
    const now = Date.now()
    await addRuleTombstones([{ type: 'domain', signalNorm: 'twice.com', targetFolderPath: '案件/甲', deletedAt: now }])
    await simulateOtherMachinePush({
      tombstones: [{ type: 'domain', signalNorm: 'twice.com', targetFolderPath: '案件/甲', deletedAt: now - 86_400_000 }],
    })
    await pullNow('remote', { mode: 'union' })
    const t = (await getRuleTombstones()).find((x) => x.signalNorm === 'twice.com')!
    expect(t.deletedAt).toBe(now)
  })
})

// ---- disabled stubs -----------------------------------------------------------

function bulky(i: number): Rule {
  return {
    ...rule({
      id: `bulk-${String(i).padStart(4, '0')}`,
      type: 'subject_keyword',
      signal: `一一四年度上字第${i}號損害賠償等事件通知書主旨與附件清單`,
      targetFolderId: 'AAMkADAwMDAwMDAwLWJiYmItY2NjYy1kZGRk' + 'z'.repeat(90) + i,
      targetFolderPath: `01進行中案件/2024年度/客戶第${i}號/往來書狀`,
      source: 'ai_confirmed',
    }),
    matchCount: 1000 - i,
  }
}

describe('Item 1: a disable propagates even when the budget drops the full row', () => {
  function otherMachineLibrary(disabled: Rule): Rule[] {
    return [...Array.from({ length: 400 }, (_, i) => bulky(i)), disabled]
  }

  it('the disabled rule is sent as a stub; only enabled rules count as truncated', () => {
    const x: Rule = {
      ...rule({ id: 'x-disabled', signal: 'sleepy.com', source: 'ai_confirmed' }),
      enabled: false,
      autoDisabledReason: 'stale',
      autoDisabledAt: '2026-10-01T00:00:00.000Z',
      editedAt: Date.now(),
    }
    const payload = buildPushPayload({
      machineId: 'a',
      rules: otherMachineLibrary(x),
      tombstones: [],
      folderActivity: [],
      settings: DEFAULT_SETTINGS,
      updatedAt: new Date().toISOString(),
      pushId: 'pid000000001',
      now: Date.now(),
    })
    const entry = payload.plan.rows.find((e) => (Array.isArray(e) ? e[0] : e.i) === 'x-disabled')
    expect(Array.isArray(entry)).toBe(true)
    expect(payload.plan.stubCount).toBe(1)
    expect(payload.plan.truncatedRuleCount).toBe(401 - payload.plan.fullRowCount - 1)
  })

  it("the other machine's enabled copy gets disabled by the stub", async () => {
    await addRules([rule({ id: 'x-disabled', signal: 'sleepy.com', source: 'ai_confirmed' })])
    const mine = (await byId('x-disabled'))!
    const disabled: Rule = {
      ...mine,
      enabled: false,
      autoDisabledReason: 'stale',
      autoDisabledAt: '2026-10-01T00:00:00.000Z',
      editedAt: mine.editedAt! + 1000,
    }
    await simulateOtherMachinePush({ rules: otherMachineLibrary(disabled) })
    await pullNow('remote', { mode: 'union' })
    // Before: dropped by the budget → stayed enabled here → re-enabled there.
    expect(await byId('x-disabled')).toMatchObject({
      enabled: false,
      autoDisabledReason: 'stale',
      autoDisabledAt: '2026-10-01T00:00:00.000Z',
      targetFolderPath: '案件/甲',
    })
  })

  it('a stub never creates a rule this machine does not have', async () => {
    const disabled: Rule = { ...rule({ id: 'never-here', source: 'ai_confirmed' }), enabled: false, editedAt: Date.now() }
    await simulateOtherMachinePush({ rules: otherMachineLibrary(disabled) })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('never-here')).toBeUndefined()
  })

  it('a newer local re-enable beats an older stub', async () => {
    await addRules([rule({ id: 'x-disabled', signal: 'sleepy.com', source: 'ai_confirmed' })])
    const mine = (await byId('x-disabled'))!
    const disabled: Rule = { ...mine, enabled: false, autoDisabledReason: 'stale', editedAt: mine.editedAt! - 1000 }
    await simulateOtherMachinePush({ rules: otherMachineLibrary(disabled) })
    await pullNow('remote', { mode: 'union' })
    expect((await byId('x-disabled'))!.enabled).toBe(true)
  })
})

// ---- one queue, pull before push, deferred local changes -------------------

describe('Item 2: one sync queue; a push merges the unmerged cloud first', () => {
  it("the first push pulls first, so the other machine's rule survives our push", async () => {
    await addRules([rule({ id: 'mine', signal: 'mine.com' })])
    await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.com' })] })
    const r = await pushNow('local-change')
    expect(r.pushed).toBe(true)
    // Before: no pull before push → the cloud lost 'theirs'.
    expect((await cloudRules()).map((x) => x.id).sort()).toEqual(['mine', 'theirs'])
    expect(await byId('theirs')).toBeDefined()
  })

  it('a push issued during a pull uploads the post-merge state', async () => {
    await addRules([rule({ id: 'mine', signal: 'mine.com' })])
    await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.com' })] })
    await Promise.all([pullNow('remote', { mode: 'union' }), pushNow('concurrent')])
    expect((await cloudRules()).map((x) => x.id).sort()).toEqual(['mine', 'theirs'])
  })

  it('settings survive the pull before push: the remote change is not pushed back stale, a pending local edit is kept', async () => {
    // Common base: both machines in sync.
    await simulateOtherMachinePush({ settings: { claudeModel: 'claude-sonnet-4-6', batchSize: 10 } })
    await pullNow('join', { mode: 'union' })
    // The other machine changes batchSize; this one (not yet aware) edits claudeModel.
    await simulateOtherMachinePush({ settings: { claudeModel: 'claude-sonnet-4-6', batchSize: 77 } })
    await setSettings({ claudeModel: 'claude-opus-5-5' })
    expect((await pushNow('local-settings-edit')).pushed).toBe(true)
    const cloud = ((await chrome.storage.sync.get('syncSettings4')).syncSettings4 as {
      v: { claudeModel: string; batchSize: number }
    }).v
    expect(cloud.batchSize).toBe(77) // not reverted to the pre-pull 10
    expect(cloud.claudeModel).toBe('claude-opus-5-5') // the pull didn't overwrite the pending edit
    expect(await getSettings()).toMatchObject({ batchSize: 77, claudeModel: 'claude-opus-5-5' })
  })

  it('after a refused pull (torn cloud) the push refuses to overwrite it', async () => {
    await addRules([rule({ id: 'mine', signal: 'mine.com' })])
    const payload = await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.com' })] })
    // One row chunk from a DIFFERENT push (arrived early / late).
    await chrome.storage.sync.set({ syncRows4_0: { ...(payload.writes['syncRows4_0'] as object), p: 'otherpush000' } })
    const r = await pushNow('local-change')
    expect(r.pushed).toBe(false)
    expect(r.reason).toContain('pull before push failed')
    expect((await readCloudMeta())!.sourceMachineId).toBe('other-machine')
  })

  it('local changes made during a pull are deferred and pushed, not dropped', async () => {
    const listener = syncListener()
    await simulateOtherMachinePush({ rules: [rule({ id: 'theirs', signal: 'theirs.com' })] })
    vi.useFakeTimers()
    try {
      await pullNow('remote', { mode: 'union' })
      // Still inside the pull's echo window: a batch finishes and adds a rule.
      await addRules([rule({ id: 'during', signal: 'during.com' })])
      listener({ rules: { newValue: [] } }, 'local')
      await vi.advanceTimersByTimeAsync(1_000 + 5_000 + 100)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    // Before: `if (pullInProgress) return` dropped it — never pushed.
    expect((await readCloudMeta())!.sourceMachineId).toBe('this-machine')
    expect((await cloudRules()).map((x) => x.id).sort()).toEqual(['during', 'theirs'])
  })

  it('a remote change arriving during a pull is queued and merged, not dropped', async () => {
    const listener = syncListener()
    await simulateOtherMachinePush({ rules: [rule({ id: 'first', signal: 'first.com' })] })
    const getMock = chrome.storage.sync.get as ReturnType<typeof vi.fn>
    const realGet = getMock.getMockImplementation() as (
      keys?: string | string[] | null,
    ) => Promise<Record<string, unknown>>
    let fired = false
    getMock.mockImplementation(async (keys?: string | string[] | null) => {
      const out = await realGet(keys)
      if (!fired && Array.isArray(keys) && keys.includes('syncMeta') && keys.includes('syncSettings4')) {
        fired = true
        // The other machine pushes again while our pull is mid-read.
        const second = await simulateOtherMachinePush({
          rules: [rule({ id: 'first', signal: 'first.com' }), rule({ id: 'second', signal: 'second.com' })],
        })
        listener({ syncMeta: { newValue: second.meta } }, 'sync')
        await new Promise((r) => setTimeout(r, 400)) // keep this pull in flight past the debounce
      }
      return out
    })
    try {
      await pullNow('manual', { mode: 'union' })
      await new Promise((r) => setTimeout(r, 300))
      await quiesce()
    } finally {
      getMock.mockImplementation(realGet)
    }
    // Before: the remote pull hit "pull already in flight" and was dropped.
    expect((await listRules()).map((r) => r.id).sort()).toEqual(['first', 'second'])
  })
})

// ---- restore -----------------------------------------------------------------

async function seedBackup(rules: Rule[]): Promise<string> {
  const { claudeApiKey: _k, ...settings } = DEFAULT_SETTINGS
  void _k
  const snapshotAt = new Date(Date.now() - 60_000).toISOString()
  const backup: SyncBackup = { snapshotAt, direction: 'manual', payload: { rules, tombstones: [], settings } }
  await chrome.storage.local.set({ [SYNC_BACKUPS_KEY]: [backup] })
  return snapshotAt
}

describe('Item 9: restoreBackup goes through mutateRules', () => {
  it('waits for the rules lock — a slow concurrent mutation can no longer overwrite the restore', async () => {
    await addRules([rule({ id: 'current', signal: 'current.com' })])
    const at = await seedBackup([{ ...rule({ id: 'restored', signal: 'restored.com' }), editedAt: 42 }])
    let release: () => void = () => {}
    const gate = new Promise<void>((r) => (release = r))
    const slow = mutateRules(async (rs) => {
      await gate
      return { next: [...rs, rule({ id: 'late-write', signal: 'late.com' })], result: undefined }
    })
    const restoring = restoreBackup(at)
    await new Promise((r) => setTimeout(r, 20))
    release()
    await Promise.all([slow, restoring])
    // Before: raw setRules landed first, then the slow writer's stale RMW
    // overwrote the restored library.
    expect((await listRules()).map((r) => r.id)).toEqual(['restored'])
  })

  it('a restored version identical to the current one keeps its version; a different one is a new edit (A5)', async () => {
    await addRules([rule({ id: 'same', signal: 'same.com' })])
    const same = (await byId('same'))!
    const before = Date.now()
    const at = await seedBackup([same, { ...rule({ id: 'gone', signal: 'gone-here.com' }), editedAt: 4242 }])
    await restoreBackup(at)
    expect((await byId('same'))!.editedAt).toBe(same.editedAt)
    // 2026-11-05 kept the snapshot's editedAt 4242: any newer copy anywhere
    // then reverted the rollback at the next pull.
    expect((await byId('gone'))!.editedAt).toBeGreaterThanOrEqual(before)
  })

  it('a rule deleted after the snapshot is un-deleted: tombstone cleared, re-stamped newer', async () => {
    await addRules([rule({ id: 'oops', signal: 'oops.com' })])
    const snap = (await byId('oops'))!
    const at = await seedBackup([snap])
    await deleteRule('oops', { actor: 'user' })
    const tomb = (await getRuleTombstones()).find((t) => t.signalNorm === 'oops.com')!
    await addRuleTombstones([{ type: 'domain', signalNorm: 'unrelated.com', targetFolderPath: 'X', deletedAt: 1 }])
    const r = await restoreBackup(at)
    expect(r).toMatchObject({ restored: true, undeletedCount: 1 })
    const back = (await byId('oops'))!
    expect(back.editedAt).toBeGreaterThan(tomb.deletedAt)
    const tombs = await getRuleTombstones()
    expect(tombs.map((t) => t.signalNorm)).toEqual(['unrelated.com'])
    // The same deletion arriving from the cloud later can't suppress it again.
    await simulateOtherMachinePush({ tombstones: [tomb] })
    await pullNow('remote', { mode: 'union' })
    expect(await byId('oops')).toBeDefined()
  })
})
