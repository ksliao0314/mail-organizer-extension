// Parallel AI chunks (2026-11-05, 放寬並加速) end to end through the real SW:
//
//   - the first request goes alone; the rest wait until its answer starts
//     streaming (its prompt-cache write is readable by then; since
//     2026-11-06 a thinking block alone no longer counts — a refusal can still
//     end the response there), then fan out to AI_PARALLEL_CHUNKS requests in
//     flight, never more;
//   - results merge into the plan in chunk order, and completedChunks counts
//     only the contiguous merged prefix, whatever order chunks finish in;
//   - a 429 drops the rest of the run to one request at a time;
//   - a run-fatal error stops new requests; chunks that finished keep their
//     results (even ones after the failed chunk); the rest are unclassified;
//   - cancel aborts every in-flight request and sends nothing new.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { AI_PARALLEL_CHUNKS } from '@/shared/constants'
import { getErrorLog } from '@/shared/error-log'
import type { PlanItem } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import { actionsJson, answerEvents, jsonErrorResponse, sseText } from './helpers/claude-sse'
import {
  gate,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  requestedEmailCount,
  seedClassifyEnv,
  waitFor,
  type FetchRouter,
} from './helpers/classify-env'

const CHUNK = 25
const LANES = AI_PARALLEL_CHUNKS

let sw: SwHarness
let router: FetchRouter
beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await seedClassifyEnv()
  router = installFetchRouter()
})
afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

async function preflight(n: number): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await markBodiesFetched()
  return (r.data as { runId: string }).runId
}

type Progress = {
  done: boolean
  completedChunks: number
  aiPlan: PlanItem[]
  aiError?: { code: string; message: string } | null
}
const full = async () =>
  ((await sw.send({ type: 'getAiClassifyProgress', view: 'full' })).data ?? null) as Progress | null

/** Which chunk a request carries: its first email's inbox index / CHUNK. */
function chunkOf(body: Record<string, unknown>): number {
  const content = (body.messages as Array<{ content: Array<{ text?: string }> }>)[0]!.content
  const m = /subject (\d+)/.exec(content[content.length - 1]!.text ?? '')
  return Math.floor(Number(m![1]) / CHUNK)
}

const chunkOfEmail = (emailId: string) => Math.floor(Number(emailId.slice(1)) / CHUNK)

/**
 * A 200 stream that begins at once (message_start, the thinking block and the
 * start of the answer's text block: the moment a real answer has started
 * streaming) and sends the rest of an answer that deletes every email only
 * after `rest()` resolves.
 */
function heldAnswer(body: Record<string, unknown>, rest: () => Promise<void>): Response {
  const n = requestedEmailCount(body)
  const events = answerEvents(
    actionsJson(Array.from({ length: n }, (_, i) => ({ emailIndex: i, action: 'delete' }))),
  )
  const textStart =
    events.findIndex((e) => (e.data as { content_block?: { type?: string } }).content_block?.type === 'text') + 1
  const enc = new TextEncoder()
  const parts = [enc.encode(sseText(events.slice(0, textStart))), enc.encode(sseText(events.slice(textStart)))]
  let step = 0
  return new Response(
    new ReadableStream<Uint8Array>({
      async pull(controller) {
        if (step === 1) {
          try {
            await rest()
          } catch (e) {
            controller.error(e)
            return
          }
        }
        if (step >= parts.length) {
          controller.close()
          return
        }
        controller.enqueue(parts[step++]!)
      },
    }),
    { status: 200, headers: { 'content-type': 'text/event-stream; charset=utf-8' } },
  )
}

const settle = () => new Promise((r) => setTimeout(r, 30))
const ids = (plan: PlanItem[]) => plan.map((p) => p.emailId)
const inboxIds = (from: number, to: number) => Array.from({ length: to - from }, (_, i) => `m${from + i}`)
const range = (from: number, to: number) => Array.from({ length: to - from }, (_, i) => from + i)

describe('fan-out', () => {
  it('waits for the first stream, keeps AI_PARALLEL_CHUNKS in flight, merges in chunk order', async () => {
    expect(LANES).toBeGreaterThan(1)
    const chunks = LANES + 2
    const prefill = gate()
    const tails = new Map<number, ReturnType<typeof gate>>()
    let inFlight = 0
    let peak = 0
    router.claude = async ({ body, signal }) => {
      const c = chunkOf(body)
      const tail = gate()
      tails.set(c, tail)
      peak = Math.max(peak, ++inFlight)
      // The first response hasn't begun streaming yet (prefill).
      if (c === 0) await prefill.wait(signal)
      return heldAnswer(body, async () => {
        await tail.wait(signal)
        inFlight--
      })
    }
    const runId = await preflight(CHUNK * chunks)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })

    await waitFor(() => router.claudeBodies.length === 1)
    await settle()
    expect(router.claudeBodies).toHaveLength(1) // cold cache: the others wait

    prefill.open()
    await waitFor(() => router.claudeBodies.length === LANES)
    await settle()
    expect(router.claudeBodies.map(chunkOf)).toEqual(range(0, LANES)) // lane cap

    // The newest chunk in flight finishes first: the next chunk starts, but
    // the plan and the progress count still wait for chunk 0.
    tails.get(LANES - 1)!.open()
    await waitFor(() => router.claudeBodies.length === LANES + 1)
    await settle()
    expect((await full())!.completedChunks).toBe(0)
    expect((await full())!.aiPlan).toHaveLength(0)

    tails.get(0)!.open()
    await waitFor(async () => (await full())!.completedChunks === 1)
    expect(ids((await full())!.aiPlan)).toEqual(inboxIds(0, CHUNK))
    await waitFor(() => router.claudeBodies.length === chunks)

    for (const g of tails.values()) g.open()
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(ids(p.aiPlan)).toEqual(inboxIds(0, CHUNK * chunks))
    expect(p.aiPlan.every((x) => x.action === 'delete')).toBe(true)
    expect(p.aiError ?? null).toBeNull()
    expect(peak).toBe(LANES)
  })
})

describe('rate limits', () => {
  it('a 429 drops the rest of the run to one request at a time', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const chunks = LANES + 3
    // The first wave finishes only after chunk 1's 429 has been handled.
    const hold = gate()
    const order: number[] = []
    const busyAtStart: number[] = []
    let inFlight = 0
    let limited = false
    router.claude = async ({ body, signal }) => {
      const c = chunkOf(body)
      order.push(c)
      busyAtStart.push(inFlight)
      if (c === 1 && !limited) {
        limited = true
        return jsonErrorResponse(429, 'rate_limit_error', 'rate limited', { 'retry-after': '1' })
      }
      inFlight++
      return heldAnswer(body, async () => {
        if (c < LANES) await hold.wait(signal)
        inFlight--
      })
    }
    const runId = await preflight(CHUNK * chunks)
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    // Chunk 1's 429 is handled: its backoff is now sleeping.
    for (let i = 0; i < 500 && !warn.mock.calls.some((c) => String(c[0]).includes('HTTP_429')); i++) {
      await vi.advanceTimersByTimeAsync(0)
    }
    expect(warn.mock.calls.some((c) => String(c[0]).includes('HTTP_429'))).toBe(true)
    hold.open()
    for (let i = 0; i < 40 && !(await full())?.done; i++) await vi.advanceTimersByTimeAsync(5_000)
    vi.useRealTimers()

    const p = (await full())!
    expect(p.done).toBe(true)
    expect(p.aiPlan.some((x) => x.unclassified)).toBe(false)
    expect(ids(p.aiPlan)).toEqual(inboxIds(0, CHUNK * chunks))
    // The first wave, chunk 1's retry, then every later chunk strictly alone.
    expect(order).toEqual([...range(0, LANES), 1, ...range(LANES, chunks)])
    expect(busyAtStart.slice(LANES)).toEqual(range(LANES - 1, chunks).map(() => 0))
  })
})

describe('run-fatal error', () => {
  it('stops new requests, keeps every finished chunk (even after the failed one), stubs the rest', async () => {
    const chunks = LANES + 3
    const hold = gate()
    router.claude = async ({ body, signal }) => {
      const c = chunkOf(body)
      if (c === 1) return jsonErrorResponse(401, 'authentication_error', 'invalid x-api-key')
      return heldAnswer(body, () => hold.wait(signal))
    }
    const runId = await preflight(CHUNK * chunks)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === LANES)
    await settle() // chunk 1's 401 is handled before the others finish
    hold.open()
    await waitFor(async () => (await full())?.done === true)

    const p = (await full())!
    expect(router.claudeBodies).toHaveLength(LANES) // nothing sent after the failure
    expect(ids(p.aiPlan)).toEqual(inboxIds(0, CHUNK * chunks))
    const stubbedChunks = [...new Set(p.aiPlan.filter((x) => x.unclassified).map((x) => chunkOfEmail(x.emailId)))]
    expect(stubbedChunks).toEqual([1, ...range(LANES, chunks)])
    expect(p.aiPlan.filter((x) => x.action === 'delete')).toHaveLength(CHUNK * (LANES - 1))
    const unclassified = CHUNK * (chunks - LANES + 1)
    expect(p.aiError!.code).toBe('CLASSIFIER_HTTP_401')
    expect(p.aiError!.message).toContain(`剩餘 ${unclassified} 封未分類`)
    const log = (await getErrorLog()).filter((e) => e.source === 'classify:api')
    expect(log).toHaveLength(1)
    expect(log[0]!.context).toMatchObject({ code: 'HTTP_401', fatal: true, chunk: 2, unclassified })
  })
})

describe('cancel', () => {
  it('aborts every in-flight request and sends nothing new', async () => {
    const hold = gate()
    router.claude = async ({ body, signal }) => heldAnswer(body, () => hold.wait(signal))
    const runId = await preflight(CHUNK * (LANES + 2))
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === LANES)

    const c = await sw.send({ type: 'cancelClassify', runId })
    expect(c.data).toEqual({ cancelled: true })
    expect(router.claudeSignals.every((s) => s?.aborted === true)).toBe(true)
    hold.open()
    await settle()
    expect(router.claudeBodies).toHaveLength(LANES)
    expect(await full()).toBeNull()
  })
})
