// SW side of the popup's cross-surface cleanup (2026-11-06 review U4 / U6).
//
// U4 — a DoneScreen's 回主畫面 now cancels its OWN prefetch run even when the
// other popup surface has already moved on (before, nothing stopped it: it
// kept calling Claude and the next open landed in its plan). But 繼續歸檔 on
// the other surface may have taken that very run over (same run id), so the
// cancel carries `unlessAdopted` and 繼續歸檔 records the adoption first.
//
// U6 — the error boundary's discard also drops a FINISHED batch's execute
// state (a crash on the done screen used to loop: mount reads the execute
// state first). `unlessRunning` makes the SW refuse while a batch runs.

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

async function startRun(n: number): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await markBodiesFetched()
  const runId = (r.data as { runId: string }).runId
  const ai = await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
  expect(ai.ok, JSON.stringify(ai)).toBe(true)
  return runId
}

const currentRunId = async () =>
  ((await chrome.storage.session.get('classifyRun')).classifyRun as { runId?: string } | undefined)?.runId ?? null

describe('cancelClassify unlessAdopted', () => {
  it('spares a run 繼續歸檔 adopted; a plain cancel (取消 on the adopting surface) still stops it', async () => {
    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    const runId = await startRun(30) // 2 chunks — still running
    await waitFor(() => router.claudeBodies.length === 1)

    expect((await sw.send({ type: 'adoptClassifyRun', runId })).data).toEqual({ adopted: true })
    const spared = await sw.send({ type: 'cancelClassify', runId, unlessAdopted: true })
    expect(spared.data).toEqual({ cancelled: false })
    expect(router.claudeSignals[0]!.aborted).toBe(false)
    expect(await currentRunId()).toBe(runId)

    const plain = await sw.send({ type: 'cancelClassify', runId })
    expect(plain.data).toEqual({ cancelled: true })
    expect(router.claudeSignals[0]!.aborted).toBe(true)
    g.open()
  })

  it('cancels a prefetch run nobody adopted (the other surface went home first)', async () => {
    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    const runId = await startRun(30)
    await waitFor(() => router.claudeBodies.length === 1)

    const r = await sw.send({ type: 'cancelClassify', runId, unlessAdopted: true })
    expect(r.data).toEqual({ cancelled: true })
    expect(router.claudeSignals[0]!.aborted).toBe(true)
    g.open()
    await new Promise((res) => setTimeout(res, 30))
    expect(router.claudeBodies).toHaveLength(1) // chunk 2 never called
    expect(await currentRunId()).toBeNull()
  })

  it('only the current run can be adopted, and the mark survives later rewrites of the run record', async () => {
    router.inboxPages = [{ messages: [inboxMail(0), inboxMail(1)] }]
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 2 })
    await markBodiesFetched()
    const runId = (pf.data as { runId: string }).runId
    expect((await sw.send({ type: 'adoptClassifyRun', runId: 'run-old' })).data).toEqual({ adopted: false })
    expect((await sw.send({ type: 'adoptClassifyRun' })).data).toEqual({ adopted: false })
    expect((await sw.send({ type: 'adoptClassifyRun', runId })).data).toEqual({ adopted: true })

    // classifyAi rewrites the run record (aiStartedAt) after the adoption.
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => {
      const p = (await sw.send({ type: 'getAiClassifyProgress', view: 'meta' })).data as { done?: boolean } | null
      return p?.done === true
    })
    expect((await sw.send({ type: 'cancelClassify', runId, unlessAdopted: true })).data).toEqual({
      cancelled: false,
    })
    expect(await currentRunId()).toBe(runId)
  })
})

describe('clearExecuteState unlessRunning', () => {
  const state = (inProgress: boolean) => ({
    inProgress,
    cancelRequested: false,
    startedAt: Date.now() - 5000,
    ...(inProgress ? {} : { finishedAt: Date.now() }),
    total: 1,
    current: 0,
    plan: [],
    results: [{ emailId: 'm1', subject: 's', action: 'move', status: inProgress ? 'processing' : 'moved' }],
    summary: { moved: inProgress ? 0 : 1, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
    rulesAdded: 0,
  })

  it('keeps a running batch’s state', async () => {
    await chrome.storage.session.set({ executeState: state(true) })
    const r = await sw.send({ type: 'clearExecuteState', unlessRunning: true })
    expect(r).toEqual({ ok: true, data: { cleared: false } })
    expect((await chrome.storage.session.get('executeState')).executeState).toBeDefined()
  })

  it('drops a finished batch’s state', async () => {
    await chrome.storage.session.set({ executeState: state(false) })
    const r = await sw.send({ type: 'clearExecuteState', unlessRunning: true })
    expect(r).toEqual({ ok: true, data: { cleared: true } })
    expect((await chrome.storage.session.get('executeState')).executeState).toBeUndefined()
  })
})
