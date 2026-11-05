// Classify-loop error policy end to end through the real SW (2026-11-05,
// items 4, 5, 9, 15).
//
//   - a 429 that outlasts the (now patient, retry-after-aware) retries stops
//     the run gracefully: every remaining email becomes an `unclassified`
//     placeholder and the failure reaches the error log (the dispatcher's
//     classify:api branch never saw loop errors);
//   - a retired model is one clear, fatal error — not a raw 404 per chunk;
//   - a refusal is chunk-local; the next chunk still gets classified;
//   - body-fetch failures stay fail-soft but are logged (once per chunk);
//   - setSettings refuses model ids we don't know.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { MODEL_UNAVAILABLE_MESSAGE } from '@/shared/classifier'
import { getErrorLog } from '@/shared/error-log'
import { getSettings } from '@/shared/storage'
import type { PlanItem } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import { answerEvents, jsonErrorResponse, sseResponse, sseText } from './helpers/claude-sse'
import {
  answerAll,
  inboxMail,
  installFetchRouter,
  markBodiesFetched,
  requestedEmailCount,
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
afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

async function preflight(n: number, opts: { bodies?: 'fetched' | 'live' } = {}): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  if (opts.bodies !== 'live') await markBodiesFetched()
  return (r.data as { runId: string }).runId
}

type Full = {
  done: boolean
  aiPlan: PlanItem[]
  aiError?: { code: string; message: string } | null
}
const full = async () =>
  ((await sw.send({ type: 'getAiClassifyProgress', view: 'full' })).data ?? null) as Full | null

describe('retries exhausted', () => {
  it('a 429 storm is waited out with patience, then the run stops gracefully and is logged', async () => {
    router.claude = async () =>
      jsonErrorResponse(429, 'rate_limit_error', 'Number of request tokens has exceeded your rate limit', {
        'retry-after': '1',
      })
    const runId = await preflight(30) // 2 chunks
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    for (let i = 0; i < 15 && !(await full())?.done; i++) await vi.advanceTimersByTimeAsync(20_000)
    vi.useRealTimers()

    const p = (await full())!
    expect(p.done).toBe(true)
    // 1 call + 4 retries (5 / 10 / 20 / 40 s ≥ retry-after) for chunk 1; the
    // run then stops — chunk 2 is never sent into a rate-limited account.
    expect(router.claudeBodies).toHaveLength(5)
    expect(p.aiPlan).toHaveLength(30)
    expect(p.aiPlan.every((x) => x.unclassified === true && x.action === 'skip')).toBe(true)
    expect(p.aiError!.code).toBe('CLASSIFIER_HTTP_429')
    expect(p.aiError!.message).toContain('流量限制（429）')
    expect(p.aiError!.message).toContain('剩餘 30 封未分類')

    const log = (await getErrorLog()).filter((e) => e.source === 'classify:api')
    expect(log).toHaveLength(1)
    expect(log[0]!.context).toMatchObject({ code: 'HTTP_429', status: 429, fatal: true, retriedMs: 75_000, unclassified: 30 })
  })
})

describe('model unavailable', () => {
  it('a retired model id fails once with an actionable message (no retries, no raw 404)', async () => {
    router.claude = async () => jsonErrorResponse(404, 'not_found_error', 'model: claude-3-opus-20240229')
    const runId = await preflight(3)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(router.claudeBodies).toHaveLength(1)
    expect(p.aiError).toEqual({ code: 'CLASSIFIER_MODEL_UNAVAILABLE', message: MODEL_UNAVAILABLE_MESSAGE })
    expect(p.aiPlan.every((x) => x.unclassified)).toBe(true)
    // Only failure placeholders left → residue: the purge drops the run so the
    // next batch re-classifies that mail fresh instead of re-showing failures.
    await sw.send({ type: 'clearExecuteState' })
    expect(await full()).toBeNull()
  })

  it('setSettings refuses unknown model ids and accepts known ones', async () => {
    await sw.send({ type: 'setSettings', patch: { claudeModel: 'claude-made-up-9' } })
    expect((await getSettings()).claudeModel).toBe('claude-sonnet-5-5')
    await sw.send({ type: 'setSettings', patch: { claudeModel: 'claude-opus-4-7' } })
    expect((await getSettings()).claudeModel).toBe('claude-opus-4-7')
  })
})

describe('refusal', () => {
  it('is chunk-local: that chunk becomes placeholders with the category, the next chunk is still classified', async () => {
    let call = 0
    router.claude = async ({ body }) => {
      call++
      if (call === 1) {
        return sseResponse(
          sseText(answerEvents('', { stopReason: 'refusal', stopDetails: { type: 'refusal', category: 'cyber' } })),
        )
      }
      expect(requestedEmailCount(body)).toBe(5)
      return answerAll(body, 'delete')
    }
    const runId = await preflight(30)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(router.claudeBodies).toHaveLength(2)
    const refused = p.aiPlan.filter((x) => x.unclassified)
    expect(refused).toHaveLength(25)
    expect(refused[0]!.reason).toBe('Claude 拒絕分類此批（類別：cyber）')
    expect(p.aiPlan.filter((x) => x.action === 'delete')).toHaveLength(5)
    expect(p.aiError).toEqual({
      code: 'CLASSIFIER_REFUSAL',
      message: 'Claude 拒絕分類其中 25 封（類別：cyber），已標為未分類',
    })
    const log = (await getErrorLog()).filter((e) => e.source === 'classify:api')
    expect(log[0]!.context).toMatchObject({ code: 'REFUSAL', category: 'cyber', fatal: false })
  })
})

describe('body fetch failures', () => {
  it('are logged once per chunk and the emails are still classified from their preview', async () => {
    router.bodyStatus = 404
    const runId = await preflight(2, { bodies: 'live' })
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(p.aiPlan).toHaveLength(2)
    expect(p.aiPlan.some((x) => x.unclassified)).toBe(false)
    const log = (await getErrorLog()).filter((e) => e.source === 'classify:body')
    expect(log).toHaveLength(1)
    expect(log[0]!.context).toMatchObject({ failed: 2, codes: ['HTTP_404'] })
  })
})
