// A classify run whose session-storage writes fail (2026-11-06, round-2
// review K6).
//
// A rejected progress write used to break mergeReady's promise chain: every
// later merge was skipped, `await mergeChain` threw out of a try…finally with
// no catch, and the detached loop's rejection went unhandled — all chunks were
// still paid for, progress stayed done:false / completedChunks:0, classifyAi
// answered alreadyStarted and the popup waited forever. 500-email batches make
// quota pressure plausible (≈4.5 MB of JSON across session keys).
//
// Now a failed write stops new chunks and the run ends in a terminal done:true
// state with STORAGE_FAILED — shrunk until it fits — and nothing escapes the
// detached loop unlogged. The folder tree is no longer re-stored with every
// progress write.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { getErrorLog } from '@/shared/error-log'
import { addRules, newRule } from '@/shared/rules'
import type { PlanItem } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import {
  FOLDER_ID,
  FOLDER_PATH,
  answerAll,
  folderTree,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  seedClassifyEnv,
  waitFor,
  type FetchRouter,
} from './helpers/classify-env'

let sw: SwHarness
let router: FetchRouter
const restores: Array<() => void> = []
beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await seedClassifyEnv()
  router = installFetchRouter()
})
afterEach(() => {
  while (restores.length > 0) restores.pop()!()
  vi.unstubAllGlobals()
})

async function preflight(n: number): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await markBodiesFetched()
  return (r.data as { runId: string }).runId
}

type Stored = {
  done: boolean
  completedChunks: number
  aiPlan: PlanItem[]
  rulePlan: PlanItem[]
  folderTree?: unknown
  aiError?: { code: string; message: string } | null
}
const stored = async () =>
  ((await chrome.storage.session.get('aiClassifyProgress')).aiClassifyProgress ?? null) as Stored | null

/**
 * Make chrome.storage.session.set reject the writes `reject` picks, like a
 * full quota (restored after the test). Swapped directly: the setup mock is
 * already a vi.fn, and vi.spyOn would hand back that same mock.
 */
function failSessionWrites(reject: (items: Record<string, unknown>) => boolean) {
  const session = chrome.storage.session as unknown as {
    set: (items: Record<string, unknown>) => Promise<void>
  }
  const original = session.set
  session.set = async (items: Record<string, unknown>) => {
    if (reject(items)) throw new Error('Session storage quota bytes exceeded. Values were not stored.')
    return original.call(session, items)
  }
  const restore = () => {
    session.set = original
  }
  restores.push(restore)
  return restore
}

const progressWith = (items: Record<string, unknown>, pick: (p: Stored) => boolean) => {
  const p = items.aiClassifyProgress as Stored | undefined
  return !!p && pick(p)
}

describe('a progress write that fails mid-run', () => {
  it('stops new chunks and ends in a terminal STORAGE_FAILED state the popup can see', async () => {
    const runId = await preflight(125) // 5 chunks
    // The quota holds the small initial snapshot but nothing with items in it.
    const restore = failSessionWrites((items) => progressWith(items, (p) => p.aiPlan.length > 0))
    const r = await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    expect(r.ok).toBe(true)

    await waitFor(async () => (await stored())?.done === true)
    const p = (await stored())!
    expect(p.aiError!.code).toBe('STORAGE_FAILED')
    expect(p.aiError!.message).toContain('每次處理筆數')
    // The bare terminal shape: counts + error, no items.
    expect(p.aiPlan).toEqual([])
    // Chunk 0 went first and its answer fanned out chunks 1-2; chunk 3 took
    // chunk 0's lane in the same tick its merge write was queued. That write
    // failed, so nothing was scheduled after it: chunk 4 was never sent.
    expect(router.claudeBodies).toHaveLength(4)
    const sent = router.claudeBodies.map((b) => JSON.stringify(b))
    expect(sent.some((b) => b.includes('subject 100'))).toBe(false)

    const view = (await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).data as {
      done: boolean
      aiError: { code: string }
    }
    expect(view.done).toBe(true)
    expect(view.aiError.code).toBe('STORAGE_FAILED')
    const log = (await getErrorLog()).filter((e) => e.source === 'classify:storage')
    expect(log).toHaveLength(1)
    expect(JSON.stringify(log)).not.toContain('sk-ant-test')
    // The loop is over: the run is not stuck "already started" forever with a
    // live loop behind it — a new batch classifies normally.
    restore()
    router.claudeBodies.length = 0
    const next = await preflight(2)
    await sw.send({ type: 'classifyAi', runId: next, excludeIds: [] })
    await waitFor(async () => (await stored())?.done === true && (await stored())!.aiPlan.length === 2)
  })

  // Round-3 review item 3: the bare shape dropped the rule plan too, so the
  // popup's summary read 規則 0 / AI 0 next to the rows it still showed.
  it('the bare terminal state keeps the counts; with no items stored it offers nothing to adopt', async () => {
    await addRules([
      newRule({
        type: 'domain',
        signal: 'rule.example',
        targetFolderId: FOLDER_ID,
        targetFolderPath: FOLDER_PATH,
        confidence: 0.7,
        source: 'user_manual',
      }),
    ])
    router.inboxPages = [
      {
        messages: [
          inboxMail(0, { From: { EmailAddress: { Address: 'a@rule.example', Name: 'R' } } }),
          inboxMail(1),
          inboxMail(2),
        ],
      },
    ]
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 3 })
    expect(pf.ok, JSON.stringify(pf)).toBe(true)
    await markBodiesFetched()
    router.claude = async ({ body }) => answerAll(body, 'delete')
    failSessionWrites((items) => progressWith(items, (p) => p.aiPlan.length > 0))
    await sw.send({ type: 'classifyAi', runId: (pf.data as { runId: string }).runId, excludeIds: [] })
    await waitFor(async () => (await stored())?.done === true)
    expect((await stored())!.rulePlan).toEqual([]) // the bare shape

    const meta = (await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).data as Record<string, unknown>
    expect(meta).toMatchObject({
      done: true,
      ruleCount: 1,
      aiCount: 2,
      aiUnresolvedCount: 0,
      aiUnclassifiedCount: 0,
      planDropped: true,
    })
    // Nothing to resume from storage: 繼續 / 完成 treat it as residue.
    await sw.send({ type: 'clearExecuteState' })
    expect(await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).toMatchObject({ ok: true, data: null })
  })

  it('keeps what fits: a slimmer terminal state when only the full one is too big', async () => {
    const runId = await preflight(3)
    // Reject only progress writes that still carry body previews.
    failSessionWrites((items) =>
      progressWith(items, (p) => [...p.aiPlan, ...p.rulePlan].some((i) => (i.bodyPreview ?? '') !== '')),
    )
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await stored())?.done === true)
    const p = (await stored())!
    expect(p.aiError!.code).toBe('STORAGE_FAILED')
    expect(p.aiPlan.map((i) => i.emailId)).toEqual(['m0', 'm1', 'm2'])
    expect(p.aiPlan.every((i) => i.bodyPreview === '')).toBe(true)
  })
})

describe('anything else that escapes the detached loop', () => {
  it('a failing cleanup after the final write is logged but keeps the finished result', async () => {
    const runId = await preflight(2)
    const session = chrome.storage.session as unknown as {
      remove: (keys: string | string[]) => Promise<void>
    }
    const original = session.remove
    session.remove = async (keys: string | string[]) => {
      if ([keys].flat().includes('preflightCache')) throw new Error('storage unavailable')
      return original.call(session, keys)
    }
    restores.push(() => {
      session.remove = original
    })
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await getErrorLog()).some((e) => e.source === 'classify:loop'))
    const p = (await stored())!
    expect(p.done).toBe(true)
    expect(p.aiError ?? null).toBeNull()
    expect(p.aiPlan.map((i) => i.emailId)).toEqual(['m0', 'm1'])
  })

  it('is logged and ends the run in a terminal state instead of an unhandled rejection', async () => {
    const runId = await preflight(2)
    // The stage display write is the loop's first session write after the
    // claim — outside any per-chunk handling.
    failSessionWrites((items) => 'classifyStage' in items)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await stored())?.done === true)
    const p = (await stored())!
    expect(p.aiError!.code).toBe('AI_FAILED')
    expect(router.claudeBodies).toHaveLength(0)
    const log = (await getErrorLog()).filter((e) => e.source === 'classify:loop')
    expect(log).toHaveLength(1)
    expect(JSON.stringify(log)).not.toContain('sk-ant-test')
  })
})

describe('a failed initial progress write', () => {
  it('ends the run before it spends anything', async () => {
    const runId = await preflight(2)
    failSessionWrites((items) => 'aiClassifyProgress' in items)
    const r = await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    expect(r).toMatchObject({ ok: false, code: 'STORAGE_FAILED' })
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(router.claudeBodies).toHaveLength(0)
    expect((await getErrorLog()).some((e) => e.source === 'classify:storage')).toBe(true)
  })
})

describe('the folder tree', () => {
  it('is no longer re-stored with every progress write; the full view reads the folder cache', async () => {
    const runId = await preflight(2)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await stored())?.done === true)
    expect(await stored()).not.toHaveProperty('folderTree')
    const full = (await sw.send({ type: 'getAiClassifyProgress', view: 'full' })).data as {
      folderTree: unknown
    }
    expect(full.folderTree).toEqual(folderTree())
    const poll = (await sw.send({ type: 'getAiClassifyProgress', view: 'poll', aiFrom: 0 })).data
    expect(poll).not.toHaveProperty('folderTree')
  })
})
