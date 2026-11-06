// 2026-11-06 review fixes:
//   A7  only recently edited disables get a guaranteed stub
//   A8  recent deletions outrank low-priority rule rows for the sync budget
//   A11 counter-only changes are pushed at most once per coalescing window
//       (plus the no-ping-pong property, checked on two simulated machines)

import { beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import { addRules, bumpRuleHits, listRules, newRule, upsertRule } from '@/shared/rules'
import { recordFolderActivityFromBatch, setSettings } from '@/shared/storage'
import {
  buildPushPayload,
  clearCloudState,
  pullNow,
  pushNow,
  quiesce,
  resetSyncEngineForTests,
  schedulePush,
  SYNC_LOCAL_STATE_KEY,
} from '@/background/sync-engine'
import { unchunkArray } from '@/shared/sync-chunks'
import { decodeRulePayload, type FolderDictEntry, type WireRuleEntry } from '@/shared/sync-codec'
import { SYNC_COUNTER_PUSH_COALESCE_MS } from '@/shared/constants'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { Rule, RuleSource, RuleTombstone } from '@/shared/types'
import { machines } from './helpers/sync-sim'

const FOLDER_A = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const DAY = 86_400_000

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

/** A compound case rule like a real library's (~73% of rows). */
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

function plan(rules: Rule[], tombstones: RuleTombstone[] = []) {
  return buildPushPayload({
    machineId: 'm',
    rules,
    tombstones,
    folderActivity: [],
    settings: DEFAULT_SETTINGS,
    updatedAt: new Date().toISOString(),
    pushId: 'pushid000001',
    now: Date.now(),
  }).plan
}

const enabledRows = (p: ReturnType<typeof plan>) =>
  p.rows.filter((e) => !Array.isArray(e) && (e as { e?: 0 }).e !== 0).length

beforeEach(async () => {
  resetSyncEngineForTests()
  await clearCloudState()
  await setSettings({ syncEnabled: true, syncMachineId: 'this-machine' })
})

// ---- A7 -------------------------------------------------------------------------

describe('A7: long-settled disables no longer crowd enabled rules out', () => {
  const enabled = Array.from({ length: 700 }, (_, i) => caseRule(i))
  const disabled = (editedAt: number) =>
    Array.from({ length: 600 }, (_, i) =>
      caseRule(2000 + i, { enabled: false, autoDisabledReason: 'legacy_token', editedAt, matchCount: 0 }),
    )

  it('700 enabled + 600 disabled long ago: as many enabled rows as with no disabled rules at all', () => {
    const baseline = plan(enabled)
    const p = plan([...enabled, ...disabled(Date.now() - 90 * DAY)])
    // Before: 600 stubs claimed the budget first (measured 189 enabled rows vs 239 on v3).
    expect(enabledRows(p)).toBe(enabledRows(baseline))
    expect(p.stubCount).toBe(0)
    // Left out by design — not reported as truncated.
    expect(p.truncatedRuleCount).toBe(baseline.truncatedRuleCount)
  })

  it('a fresh disable still gets its guaranteed stub', () => {
    const p = plan([...enabled, ...disabled(Date.now() - 90 * DAY), caseRule(9000, { enabled: false, editedAt: Date.now() })])
    expect(p.stubCount).toBe(1)
    expect(p.rows.some((e) => Array.isArray(e) && e[0] === 'case-9000')).toBe(true)
  })
})

// ---- A8 -------------------------------------------------------------------------

describe('A8: recent deletions outrank low-priority rows', () => {
  function cjkTomb(i: number, deletedAt: number): RuleTombstone {
    return {
      type: 'subject_keyword',
      signalNorm: `一一三年度訴字第${i}號請求返還不當得利事件言詞辯論通知書`,
      targetFolderPath: `01進行中案件/已結案/客戶第${i}號`,
      deletedAt,
    }
  }
  const core = Array.from({ length: 120 }, (_, i) => caseRule(i, { source: 'user_manual' }))
  const neverMatched = Array.from({ length: 400 }, (_, i) => caseRule(3000 + i, { matchCount: 0 }))
  const bulkDelete = Array.from({ length: 300 }, (_, i) => cjkTomb(i, Date.now() - i * 1000))

  it('a 300-rule bulk delete reaches the cloud without costing a single core rule', () => {
    const baseline = plan([...core, ...neverMatched])
    const p = plan([...core, ...neverMatched], bulkDelete)
    // Before: the 15 % share held ~94 of them; the peer kept routing with the rest.
    expect(p.tombstones.length).toBeGreaterThanOrEqual(200)
    const kept = new Set(p.rows.map((e) => (Array.isArray(e) ? e[0] : e.i)))
    expect(core.every((r) => kept.has(r.id))).toBe(true)
    expect(p.fullRowCount).toBeLessThan(baseline.fullRowCount) // never-matched rows made room
    expect(p.bytes).toBeLessThanOrEqual(88 * 1024)
  })

  it('old deletions do not get that priority', () => {
    const old = bulkDelete.map((t) => ({ ...t, deletedAt: t.deletedAt - 90 * DAY }))
    const baseline = plan([...core, ...neverMatched])
    const p = plan([...core, ...neverMatched], old)
    expect(p.tombstones.length).toBeLessThan(150)
    expect(baseline.fullRowCount - p.fullRowCount).toBeLessThan(80)
  })
})

// ---- A11 ------------------------------------------------------------------------

const syncSet = () => chrome.storage.sync.set as unknown as Mock
async function localState(): Promise<{ counterPushDueAt?: number } | undefined> {
  return (await chrome.storage.local.get(SYNC_LOCAL_STATE_KEY))[SYNC_LOCAL_STATE_KEY] as
    | { counterPushDueAt?: number }
    | undefined
}
async function cloudRule(id: string): Promise<Rule | undefined> {
  const cloud = await chrome.storage.sync.get(null)
  const pick = <T,>(prefix: string) =>
    unchunkArray(
      Object.keys(cloud)
        .filter((k) => k.startsWith(prefix))
        .map((k) => cloud[k] as { index: number; items: T[] }),
    )
  return decodeRulePayload({ folders: pick<FolderDictEntry>('syncFolders4_'), rows: pick<WireRuleEntry>('syncRows4_') })
    .rules.find((r) => r.id === id)
}

/** Fire the debounced push path (what a local storage change schedules). */
async function debouncedPush(advanceMs = 50): Promise<void> {
  schedulePush('local-mutation', 0)
  await vi.advanceTimersByTimeAsync(advanceMs)
}

describe('A11: counter-only changes are coalesced', () => {
  it('a hit-count bump is not pushed at once; one coalesced push carries it later', async () => {
    await addRules([rule({ id: 'c', signal: 'c.example', source: 'ai_confirmed' })])
    await pushNow('initial')
    syncSet().mockClear()
    vi.useFakeTimers()
    try {
      await bumpRuleHits(new Map([['c', 3]]))
      await debouncedPush()
      // Before: every batch's counter bump rewrote all ~88 KB of chunks.
      expect(syncSet()).not.toHaveBeenCalled()
      expect((await localState())?.counterPushDueAt).toBeGreaterThan(Date.now())
      await bumpRuleHits(new Map([['c', 2]]))
      await debouncedPush()
      expect(syncSet()).not.toHaveBeenCalled()
      await vi.advanceTimersByTimeAsync(SYNC_COUNTER_PUSH_COALESCE_MS + 100)
    } finally {
      vi.useRealTimers()
    }
    await quiesce()
    expect(syncSet()).toHaveBeenCalledTimes(1)
    expect((await cloudRule('c'))!.matchCount).toBe(5)
    expect((await localState())?.counterPushDueAt).toBeUndefined()
  })

  it('semantic edits, folder activity and 立即上傳 still push at once (carrying pending counters)', async () => {
    await addRules([rule({ id: 'c', signal: 'c.example', source: 'ai_confirmed' })])
    await pushNow('initial')
    vi.useFakeTimers()
    try {
      await bumpRuleHits(new Map([['c', 4]]))
      await debouncedPush()
      syncSet().mockClear()
      await upsertRule({ ...(await listRules())[0]!, confidence: 0.6 }, { actor: 'user' })
      await debouncedPush()
      expect(syncSet()).toHaveBeenCalledTimes(1)
      expect((await cloudRule('c'))!).toMatchObject({ matchCount: 4, confidence: 0.6 })
      expect((await localState())?.counterPushDueAt).toBeUndefined()

      await bumpRuleHits(new Map([['c', 1]]))
      await recordFolderActivityFromBatch([{ folderId: FOLDER_A, folderPath: '案件/甲', count: 1 }], new Date().toISOString())
      await debouncedPush()
      expect(syncSet()).toHaveBeenCalledTimes(2)

      await bumpRuleHits(new Map([['c', 1]]))
      await debouncedPush()
      expect(syncSet()).toHaveBeenCalledTimes(2) // deferred
    } finally {
      vi.useRealTimers()
    }
    expect((await pushNow('manual')).pushed).toBe(true) // 立即上傳 flushes
    expect(syncSet()).toHaveBeenCalledTimes(3)
    expect((await cloudRule('c'))!.matchCount).toBe(6)
  })

  it('switching sync off flushes counter changes still waiting', async () => {
    const { handleDisableSync } = await import('@/background/handlers/sync')
    await addRules([rule({ id: 'c', signal: 'c.example', source: 'ai_confirmed' })])
    await pushNow('initial')
    vi.useFakeTimers()
    try {
      await bumpRuleHits(new Map([['c', 7]]))
      await debouncedPush()
    } finally {
      vi.useRealTimers()
    }
    expect((await cloudRule('c'))!.matchCount).toBe(0)
    await handleDisableSync({ keepCloud: true })
    expect((await cloudRule('c'))!.matchCount).toBe(7)
  })

  it('no ping-pong: an edit costs one write and the peer’s follow-up push is a no-op', async () => {
    const m = machines()
    await m.use('A')
    await setSettings({ syncEnabled: true, syncMachineId: 'machine-a' })
    await addRules([rule({ id: 'R', signal: 'client.example', source: 'ai_confirmed' })])
    await pushNow('a-initial')
    await m.use('B')
    await setSettings({ syncEnabled: true, syncMachineId: 'machine-b' })
    await pullNow('b-join', { mode: 'union' })
    expect(await pushNow('b-echo')).toMatchObject({ unchanged: true })

    syncSet().mockClear()
    await m.use('A')
    await upsertRule({ ...(await listRules())[0]!, confidence: 0.65 }, { actor: 'user' })
    expect((await pushNow('a-edit')).pushed).toBe(true)
    expect(syncSet()).toHaveBeenCalledTimes(1)

    await m.use('B')
    await pullNow('b-remote', { mode: 'union' })
    expect((await listRules())[0]!.confidence).toBe(0.65)
    expect(await pushNow('b-follow-up')).toMatchObject({ unchanged: true })
    // B classified mail meanwhile: its counter-only difference waits too.
    vi.useFakeTimers()
    try {
      await bumpRuleHits(new Map([['R', 2]]))
      await debouncedPush()
    } finally {
      vi.useRealTimers()
    }
    expect(syncSet()).toHaveBeenCalledTimes(1)
  })
})
