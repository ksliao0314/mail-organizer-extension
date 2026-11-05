// Classify runs have an owner (2026-11-05, items 1 + 3).
//
// Before: 取消 / 完成 only deleted storage keys while the SW loop kept calling
// Claude for every remaining chunk and rewrote progress; 取消 → 開始歸類 ran two
// loops on one key (double spend), the old loop's done:true ended the new
// poll early, and its finally deleted the new run's preflight cache. Now a
// preflight mints a runId, the loop checks ownership before every Claude call
// and progress write, and cancel / supersede abort the in-flight call.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import {
  answerAll,
  gate,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  seedClassifyEnv,
  waitFor,
  type FetchRouter,
} from './helpers/classify-env'

let sw: SwHarness
let router: FetchRouter
beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await seedClassifyEnv()
  router = installFetchRouter()
})
afterEach(() => vi.unstubAllGlobals())

async function preflight(n: number): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await markBodiesFetched()
  return (r.data as { runId: string }).runId
}

type Progress = { runId: string; done: boolean; aiPlan?: unknown[]; aiCount: number }
async function progress(): Promise<Progress | null> {
  return ((await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).data ?? null) as Progress | null
}

describe('cancelClassify', () => {
  it('stops the loop before its next Claude call, aborts the in-flight one, and leaves nothing to resume', async () => {
    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    const runId = await preflight(30) // 2 chunks (25 + 5)
    const started = await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    expect(started.data).toMatchObject({ started: true, runId, chunks: 2 })
    await waitFor(() => router.claudeBodies.length === 1)

    const c = await sw.send({ type: 'cancelClassify', runId })
    expect(c.data).toEqual({ cancelled: true })
    expect(router.claudeSignals[0]!.aborted).toBe(true)
    g.open()
    await new Promise((r) => setTimeout(r, 30))

    expect(router.claudeBodies).toHaveLength(1) // chunk 2 never called
    expect(await progress()).toBeNull()
    const session = await chrome.storage.session.get(['aiClassifyProgress', 'preflightCache', 'classifyRun'])
    expect(session).toEqual({})
  })

  it('a stale cancel (other runId) is a no-op for the current run', async () => {
    const runId = await preflight(2)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    const c = await sw.send({ type: 'cancelClassify', runId: 'run-old' })
    expect(c.data).toEqual({ cancelled: false })
    await waitFor(async () => (await progress())?.done === true)
    expect((await progress())!.runId).toBe(runId)
  })

  it('clearAiClassifyProgress also invalidates the run (its loop stops writing)', async () => {
    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    const runId = await preflight(30)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === 1)
    await sw.send({ type: 'clearAiClassifyProgress' })
    g.open()
    await new Promise((r) => setTimeout(r, 30))
    expect(router.claudeBodies).toHaveLength(1)
    expect(await progress()).toBeNull()
  })
})

describe('a newer run supersedes the old one', () => {
  it('old classifyAi is refused; the old loop neither writes progress nor deletes the new preflight cache', async () => {
    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    const runA = await preflight(3)
    await sw.send({ type: 'classifyAi', runId: runA, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === 1)

    // 取消 → 開始歸類 from the popup = a fresh preflight.
    const runB = await preflight(3)
    expect(runB).not.toBe(runA)
    expect(router.claudeSignals[0]!.aborted).toBe(true) // A's spend stops now
    g.open()
    await new Promise((r) => setTimeout(r, 30))

    const stale = await sw.send({ type: 'classifyAi', runId: runA, excludeIds: [] })
    expect(stale).toMatchObject({ ok: false, code: 'STALE_RUN' })
    // A's finally must not have removed B's preflight context.
    const cache = (await chrome.storage.session.get('preflightCache')).preflightCache as { runId: string }
    expect(cache.runId).toBe(runB)
    expect(await progress()).toBeNull() // A wrote nothing after losing the run

    router.claude = async ({ body }) => answerAll(body)
    const startedB = await sw.send({ type: 'classifyAi', runId: runB, excludeIds: [] })
    expect(startedB.ok).toBe(true)
    await waitFor(async () => (await progress())?.done === true)
    const p = (await progress())!
    expect(p.runId).toBe(runB)
    expect(p.aiCount).toBe(3)
    expect(router.claudeBodies).toHaveLength(2) // A once (aborted), B once
  })
})

describe('one classifyAi loop per run (item 3)', () => {
  it('a double-submitted classifyAi starts a single loop', async () => {
    const runId = await preflight(2)
    const [a, b] = await Promise.all([
      sw.send({ type: 'classifyAi', runId, excludeIds: [] }),
      sw.send({ type: 'classifyAi', runId, excludeIds: [] }),
    ])
    const datas = [a.data, b.data] as Array<{ started: boolean; alreadyStarted?: boolean }>
    expect(datas.filter((d) => d.started)).toHaveLength(1)
    expect(datas.filter((d) => d.alreadyStarted)).toHaveLength(1)
    await waitFor(async () => (await progress())?.done === true)
    // …and a late third one doesn't re-run the finished run either.
    const late = await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    expect(late.ok === false || (late.data as { started: boolean }).started === false).toBe(true)
    await new Promise((r) => setTimeout(r, 20))
    expect(router.claudeBodies).toHaveLength(1)
  })
})
