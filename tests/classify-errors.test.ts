// Classify-loop error policy end to end through the real SW (2026-11-05,
// items 4, 5, 9, 15).
//
//   - a 429 that outlasts the (now patient, retry-after-aware) retries stops
//     the run gracefully: every remaining email becomes an `unclassified`
//     placeholder and the failure reaches the error log (the dispatcher's
//     classify:api branch never saw loop errors);
//   - a retired model is one clear, fatal error — not a raw 404 per chunk;
//   - a refusal is isolated to the refused email(s) by bisection (2026-11-06):
//     only they become placeholders, named in the banner, and are remembered
//     so later batches leave them out; before a run bisects or blames, it
//     needs proof the prompt every request shares is acceptable (an answered
//     call, or ONE probe with a harmless email) — a refusal every request
//     trips is never bisected, never pinned on an email, and remembered
//     across runs (round-2 review K1 / K2 / K4 / K5);
//   - body-fetch failures stay fail-soft but are logged (once per chunk);
//   - setSettings refuses model ids we don't know.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { MODEL_UNAVAILABLE_MESSAGE, type ClassifierUsage } from '@/shared/classifier'
import { getErrorLog } from '@/shared/error-log'
import { addRefusedEmails, getRefusalSuspicion, getRefusedEmails } from '@/shared/refusal-ledger'
import { addRules, newRule } from '@/shared/rules'
import { getSettings } from '@/shared/storage'
import type { PlanItem } from '@/shared/types'
import { batchRefusedReason } from '@/background/refusal-isolation'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
import { answerEvents, jsonErrorResponse, refusalEvents, sseResponse, sseText } from './helpers/claude-sse'
import {
  FOLDER_ID,
  FOLDER_PATH,
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
afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

type PreflightData = {
  runId: string
  unmatchedPreview: Array<{ id: string }>
  refusedCount: number
  flaggedCount: number
}
let lastPreflight: PreflightData

async function preflight(n: number, opts: { bodies?: 'fetched' | 'live' } = {}): Promise<string> {
  router.inboxPages = [{ messages: Array.from({ length: n }, (_, i) => inboxMail(i)) }]
  const r = await sw.send({ type: 'classifyPreflight', batchSize: n })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  if (opts.bodies !== 'live') await markBodiesFetched()
  lastPreflight = r.data as PreflightData
  return lastPreflight.runId
}

type Full = {
  done: boolean
  completedChunks?: number
  aiPlan: PlanItem[]
  aiError?: { code: string; message: string } | null
  usage?: ClassifierUsage
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

// ---- refusal isolation (2026-11-06, D1) ------------------------------------

/** Inbox indexes of the emails a classify request carries (inboxMail subjects). */
function subjectsOf(body: Record<string, unknown>): number[] {
  return [...emailBlock(body).matchAll(/主旨:subject (\d+)/g)].map((m) => Number(m[1]))
}

function emailBlock(body: Record<string, unknown>): string {
  const content = (body.messages as Array<{ content: Array<{ text?: string }> }>)[0]!.content
  return content[content.length - 1]!.text ?? ''
}

/** The run's probe: the shared prompt with one harmless synthetic email. */
const isProbe = (body: Record<string, unknown>) => emailBlock(body).includes('主旨:下週會議時間確認')

/** A decline before any output, as the API streams it (no content blocks). */
function refusalResponse(category: string | null, details: Record<string, unknown> = {}): Response {
  return sseResponse(sseText(refusalEvents(category, { details })))
}

/** Refuse every request carrying one of `culprits`; answer the rest (delete). */
function refuseEmails(culprits: number[], category: string | null = 'cyber', details: Record<string, unknown> = {}) {
  return async ({ body }: { body: Record<string, unknown> }) =>
    subjectsOf(body).some((i) => culprits.includes(i))
      ? refusalResponse(category, details)
      : answerAll(body, 'delete')
}

async function runBatch(n: number): Promise<Full> {
  const runId = await preflight(n)
  await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
  await waitFor(async () => (await full())?.done === true)
  return (await full())!
}

describe('refusal', () => {
  it('is isolated to the refused email: only it is unclassified (named, with its category); the rest is classified', async () => {
    router.claude = refuseEmails([7])
    const p = await runBatch(30)

    // Chunk 0: the refused call, ONE probe (the shared prompt with a harmless
    // email — answered, so the refusal is about some email of the chunk), then
    // 10 bisection calls (25 → 13 → 6 → 3 → 2 → 1); chunk 1: one call.
    expect(router.claudeBodies).toHaveLength(13)
    expect(router.claudeBodies.filter(isProbe)).toHaveLength(1)
    expect(isProbe(router.claudeBodies[1]!)).toBe(true)
    expect(p.aiPlan).toHaveLength(30)
    expect(p.aiPlan.some((x) => x.emailId.includes('probe'))).toBe(false) // the probe's answer is discarded
    const unclassified = p.aiPlan.filter((x) => x.unclassified)
    expect(unclassified.map((x) => x.emailId)).toEqual(['m7'])
    expect(unclassified[0]!.reason).toBe('Claude 拒絕分類「subject 7」（類別：資訊安全），請手動處理')
    expect(unclassified[0]!.refusal).toEqual({ scope: 'email', category: 'cyber' })
    expect(unclassified[0]!.source).toBe('unresolved')
    expect(p.aiPlan.filter((x) => x.action === 'delete')).toHaveLength(29)
    expect(p.aiError).toEqual({
      code: 'CLASSIFIER_REFUSAL',
      message: 'Claude 拒絕分類 1 封郵件：「subject 7」（類別：資訊安全），已標為未分類，請手動處理',
    })
    // Every attempt counted: the 7 answered calls (probe included; 120 in / 40
    // out each); the 6 cyber declines before any output are not billed.
    expect(p.usage).toMatchObject({ inputTokens: 840, outputTokens: 280 })
    expect(p.usage!.costUsd).toBeCloseTo((840 * 2 + 280 * 10) / 1e6, 12)

    const log = (await getErrorLog()).filter((e) => e.source === 'classify:api')
    expect(log).toHaveLength(1)
    expect(log[0]!.context).toMatchObject({
      code: 'REFUSAL',
      category: 'cyber',
      fatal: false,
      refused: 1,
      unclassified: 1,
      isolationCalls: 10,
    })
    expect(JSON.stringify(log)).not.toContain('subject 7') // no mail content in the log
    expect(await getRefusedEmails()).toEqual({
      m7: { at: expect.any(Number), category: 'cyber', model: 'claude-sonnet-5-5' },
    })
    expect(await getRefusalSuspicion()).toBeNull()
  })

  it('counts billed declines too (bio is billed even before any output)', async () => {
    router.claude = refuseEmails([3], 'bio')
    const p = await runBatch(5)
    // [0..4] refused → probe answered → [0,1,2] answered → [3,4] refused → [3] refused, [4] answered.
    expect(router.claudeBodies.map(subjectsOf)).toEqual([[0, 1, 2, 3, 4], [], [0, 1, 2], [3, 4], [3], [4]])
    expect(p.usage).toMatchObject({ inputTokens: 3 * 300 + 3 * 120, outputTokens: 3 * 40 })
  })

  // Round-2 review K5: a ledgered email still took a batch slot (re-shown as a
  // placeholder every batch) — an inbox topped by ≥ batchSize of them made
  // every batch show only them and 繼續 spin in place.
  it('a refused email is left out of later batches and counted for the user; another model retries it', async () => {
    router.claude = refuseEmails([3])
    await runBatch(5)
    expect(router.claudeBodies).toHaveLength(6) // 1 + probe + 4 bisection calls

    router.claudeBodies.length = 0
    const next = await runBatch(5)
    expect(lastPreflight.refusedCount).toBe(1)
    expect(lastPreflight.unmatchedPreview.map((m) => m.id)).toEqual(['m0', 'm1', 'm2', 'm4'])
    expect(router.claudeBodies.map(subjectsOf)).toEqual([[0, 1, 2, 4]])
    expect(next.aiPlan.map((x) => x.emailId)).toEqual(['m0', 'm1', 'm2', 'm4'])
    expect(next.aiError ?? null).toBeNull()

    // The 繼續 probe agrees: refused mail is no slot of the next batch.
    router.inboxPages = [{ messages: Array.from({ length: 5 }, (_, i) => inboxMail(i)) }]
    const peek = await sw.send({ type: 'peekNextBatch', batchSize: 5 })
    expect(peek.data).toMatchObject({ eligibleCount: 4 })

    // The ledger is per model: switching models gives the email a fresh try.
    await sw.send({ type: 'setSettings', patch: { claudeModel: 'claude-opus-5-5' } })
    router.claudeBodies.length = 0
    router.claude = async ({ body }) => answerAll(body, 'delete')
    const other = await runBatch(5)
    expect(lastPreflight.refusedCount).toBe(0)
    expect(router.claudeBodies.map(subjectsOf)).toEqual([[0, 1, 2, 3, 4]])
    expect(other.aiPlan.some((x) => x.unclassified)).toBe(false)
  })

  it('a page of only refused mail takes no slot: the scan pages past it', async () => {
    await addRefusedEmails(
      [0, 1, 2].map((i) => ({ emailId: `m${i}`, category: 'cyber', model: 'claude-sonnet-5-5' })),
    )
    router.inboxPages = [
      {
        messages: [0, 1, 2].map((i) => inboxMail(i)),
        nextLink: 'https://outlook.office.com/api/v2.0/me/MailFolders/inbox/messages?%24skip=3',
      },
      { messages: [3, 4, 5].map((i) => inboxMail(i)) },
    ]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 3 })
    const d = r.data as PreflightData
    expect(d.refusedCount).toBe(3)
    expect(d.unmatchedPreview.map((m) => m.id)).toEqual(['m3', 'm4', 'm5'])
  })

  // Round-3 review item 2: the exclusion ran before rule matching, so for up
  // to 7 days a refused email didn't get even its user_manual rule or its
  // thread routing — neither needs the AI.
  it('a refused email still gets its rule and its thread routing; only AI-bound refused mail is left out', async () => {
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
    await chrome.storage.local.set({
      conversationMemory: {
        'conv-filed': {
          folderId: FOLDER_ID,
          folderPath: FOLDER_PATH,
          lastFiledAt: '2026-11-01T00:00:00.000Z',
          timesFiled: 2,
          conflictCount: 0,
        },
      },
    })
    await addRefusedEmails(
      [0, 1, 2].map((i) => ({ emailId: `m${i}`, category: 'cyber', model: 'claude-sonnet-5-5' })),
    )
    router.inboxPages = [
      {
        messages: [
          inboxMail(0, { From: { EmailAddress: { Address: 'a@rule.example', Name: 'R' } } }), // a rule routes it
          inboxMail(1, { ConversationId: 'conv-filed' }), // thread memory routes it
          inboxMail(2), // only the AI could — left out
          inboxMail(3),
        ],
      },
    ]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 1 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    const d = r.data as PreflightData & { rulePlan: PlanItem[] }
    expect(d.rulePlan.map((p) => [p.emailId, p.source])).toEqual([
      ['m0', 'rule'],
      ['m1', 'thread'],
    ])
    // Refused mail takes no batch slot: batchSize 1 is still m3's.
    expect(d.unmatchedPreview.map((m) => m.id)).toEqual(['m3'])
    expect(d.refusedCount).toBe(1)
  })

  it('the exclusion counts ride along with the run, for a surface that adopts it later', async () => {
    await addRefusedEmails([{ emailId: 'm1', category: 'cyber', model: 'claude-sonnet-5-5' }])
    router.claude = async ({ body }) => answerAll(body, 'delete')
    router.inboxPages = [
      { messages: [inboxMail(0, { Flag: { FlagStatus: 'Flagged' } }), inboxMail(1), inboxMail(2)] },
    ]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 2 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    await markBodiesFetched()
    await sw.send({ type: 'classifyAi', runId: (r.data as PreflightData).runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    for (const view of ['meta', 'poll', 'full'] as const) {
      const v = (await sw.send({ type: 'getAiClassifyProgress', view })).data as { excluded?: unknown }
      expect(v.excluded, view).toEqual({ kept: 0, flagged: 1, refused: 1 })
    }
  })

  it('an email ledgered after the preflight is not sent: its row says it was refused before', async () => {
    router.claude = async ({ body }) => answerAll(body, 'delete')
    const runId = await preflight(3)
    await addRefusedEmails([{ emailId: 'm1', category: 'bio', model: 'claude-sonnet-5-5' }])
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(router.claudeBodies.map(subjectsOf)).toEqual([[0, 2]])
    const row = p.aiPlan.find((x) => x.emailId === 'm1')!
    expect(row.reason).toBe('Claude 先前拒絕分類「subject 1」（類別：生物安全），這次未再送出，請手動處理')
    expect(row.refusal).toEqual({ scope: 'email', category: 'bio' })
    expect(p.aiError!.code).toBe('CLASSIFIER_REFUSAL')
  })

  // Round-2 review K2: with one AI email (or only culprits) and no answered
  // call in the run, the refusal stayed generic and unremembered — the email
  // was re-sent and refused every batch.
  it('a lone refused email is named and remembered once the probe clears the shared prompt', async () => {
    router.claude = refuseEmails([0])
    const p = await runBatch(1)
    // The email alone (refused), then the probe (answered): the refusal is its own.
    expect(router.claudeBodies).toHaveLength(2)
    expect(isProbe(router.claudeBodies[1]!)).toBe(true)
    expect(p.aiPlan).toHaveLength(1)
    expect(p.aiPlan[0]!.reason).toBe('Claude 拒絕分類「subject 0」（類別：資訊安全），請手動處理')
    expect(p.aiPlan[0]!.refusal).toEqual({ scope: 'email', category: 'cyber' })
    expect(Object.keys(await getRefusedEmails())).toEqual(['m0'])

    // Not re-sent: the next batch leaves it out.
    router.claudeBodies.length = 0
    await preflight(1)
    expect(lastPreflight.refusedCount).toBe(1)
    expect(lastPreflight.unmatchedPreview).toEqual([])
    expect(router.claudeBodies).toHaveLength(0)
  })

  it('two culprits and nothing else: both are named and remembered', async () => {
    router.claude = refuseEmails([0, 1])
    const p = await runBatch(2)
    // [0,1] refused → probe answered → [0] refused, [1] refused.
    expect(router.claudeBodies.map(subjectsOf)).toEqual([[0, 1], [], [0], [1]])
    expect(p.aiPlan.map((x) => x.refusal?.scope)).toEqual(['email', 'email'])
    expect(Object.keys(await getRefusedEmails()).sort()).toEqual(['m0', 'm1'])
    expect(p.aiError!.message).toContain('Claude 拒絕分類 2 封郵件：「subject 0」、「subject 1」')
  })

  // Round-2 review K1: a refusal of the prompt every request shares was
  // bisected to the 12-call chunk budget before anything noticed (14 calls for
  // 30 emails, every batch; v1.0.8 spent 2).
  it('a refusal every request trips is not bisected: one probe, one call per chunk, remembered for later runs', async () => {
    router.claude = async () => refusalResponse('cyber')
    const p = await runBatch(30)
    // Chunk 0, the probe (refused too → the shared prompt is the problem), chunk 1.
    expect(router.claudeBodies).toHaveLength(3)
    expect(isProbe(router.claudeBodies[1]!)).toBe(true)
    expect(p.aiPlan).toHaveLength(30)
    expect(p.aiPlan.every((x) => x.unclassified && x.refusal?.scope === 'batch')).toBe(true)
    expect(p.aiPlan[0]!.reason).toBe(batchRefusedReason('cyber'))
    expect(p.aiError!.code).toBe('CLASSIFIER_REFUSAL')
    // K4: every unclassified email counted, not just the first chunk's 25.
    expect(p.aiError!.message).toContain('Claude 拒絕分類其中 30 封郵件（類別：資訊安全）')
    expect(p.aiError!.message).toContain('連只含一封測試郵件的請求也被拒絕')
    expect(await getRefusedEmails()).toEqual({})
    expect(await getRefusalSuspicion()).toMatchObject({ model: 'claude-sonnet-5-5', category: 'cyber' })

    // The next batch with the same prompt: no probe, no bisection — v1.0.8's
    // one call per chunk.
    router.claudeBodies.length = 0
    const again = await runBatch(30)
    expect(router.claudeBodies).toHaveLength(2)
    expect(router.claudeBodies.some(isProbe)).toBe(false)
    expect(again.aiError!.message).toContain('Claude 拒絕分類其中 30 封郵件')
  })

  // Round-2 review K1: a refusal after the thinking block opened the
  // cold-cache gate, so three lanes bisected at once (39 calls for 75 emails,
  // 41 for 125, every batch).
  it('a refusal after streamed thinking never fans out: one lane, one probe, one call per chunk', async () => {
    let inFlight = 0
    let peak = 0
    router.claude = async () => {
      peak = Math.max(peak, ++inFlight)
      const bytes = new TextEncoder().encode(sseText(refusalEvents('cyber', { when: 'after-thinking' })))
      let sent = false
      return new Response(
        new ReadableStream<Uint8Array>({
          async pull(controller) {
            if (sent) {
              inFlight--
              controller.close()
              return
            }
            sent = true
            await new Promise((r) => setTimeout(r, 2)) // let any fan-out happen
            controller.enqueue(bytes)
          },
        }),
        { status: 200, headers: { 'content-type': 'text/event-stream' } },
      )
    }
    const p = await runBatch(125)
    // 5 chunks + the probe; strictly one request at a time.
    expect(router.claudeBodies).toHaveLength(6)
    expect(peak).toBe(1)
    expect(p.aiPlan.filter((x) => x.unclassified)).toHaveLength(125)
    expect(p.aiError!.message).toContain('Claude 拒絕分類其中 125 封郵件')

    router.claudeBodies.length = 0
    peak = 0
    await runBatch(125)
    expect(router.claudeBodies).toHaveLength(5)
    expect(peak).toBe(1)
  })

  it('an answer after a shared verdict lifts it: a later refused chunk is isolated, the suspicion cleared', async () => {
    router.claude = async ({ body }) => {
      const s = subjectsOf(body)
      if (isProbe(body) || s.every((i) => i < 25)) return refusalResponse('cyber') // chunk 0 + probe
      return s.includes(52) ? refusalResponse('cyber') : answerAll(body, 'delete')
    }
    const p = await runBatch(55)
    // 1 (chunk 0) + 1 (probe) + 1 (chunk 1, answered) + 5 (chunk 2: 5 → 3 → 2 | 1, 2).
    expect(router.claudeBodies).toHaveLength(8)
    expect(p.aiPlan.filter((x) => x.unclassified)).toHaveLength(26)
    expect(p.aiPlan.find((x) => x.emailId === 'm52')!.reason).toBe(
      'Claude 拒絕分類「subject 52」（類別：資訊安全），請手動處理',
    )
    expect(Object.keys(await getRefusedEmails())).toEqual(['m52'])
    // The answered chunk proved the prompt: the remembered suspicion is gone.
    expect(await getRefusalSuspicion()).toBeNull()
    expect(p.aiError!.message).toBe(
      'Claude 拒絕分類 1 封郵件：「subject 52」（類別：資訊安全），已標為未分類，請手動處理；' +
        '另有 25 封郵件所在的整批請求被拒（類別：資訊安全），無法確定是哪一封，下一批會再試',
    )
  })

  it('a refusal whose fallback was skipped (recommended_model) is isolated but not remembered', async () => {
    router.claude = refuseEmails([1], 'cyber', { recommended_model: 'claude-sonnet-5' })
    const p = await runBatch(4)
    expect(p.aiPlan.filter((x) => x.unclassified).map((x) => x.emailId)).toEqual(['m1'])
    expect(await getRefusedEmails()).toEqual({})
  })

  it('cancelling while the probe is in flight stops everything: no further calls, no progress', async () => {
    const hold = gate()
    router.claude = async ({ body, signal }) => {
      if (router.claudeBodies.length === 1) return refusalResponse('cyber')
      await hold.wait(signal)
      return answerAll(body, 'delete')
    }
    const runId = await preflight(6)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === 2) // the probe in flight
    expect(isProbe(router.claudeBodies[1]!)).toBe(true)
    await sw.send({ type: 'cancelClassify', runId })
    hold.open()
    await new Promise((r) => setTimeout(r, 30))
    expect(router.claudeBodies).toHaveLength(2)
    expect(await full()).toBeNull()
    expect(await getRefusedEmails()).toEqual({})
    expect(await getRefusalSuspicion()).toBeNull()
  })

  // Round-3 review item 4: an answer that fails to parse is still an answer —
  // the shared prompt was accepted. The probe's set the gate to proven but
  // never re-ran the scheduler, and an ordinary chunk's didn't count at all,
  // so the three lanes only opened at the first successful bisection call.
  it('a probe answered with unparsable JSON still proves the prompt: the run fans out at once', async () => {
    const hold = gate()
    router.claude = async ({ body, signal }) => {
      if (isProbe(body)) return sseResponse(sseText(answerEvents('這不是 JSON')))
      const s = subjectsOf(body)
      if (s.includes(7)) return refusalResponse('cyber')
      // Chunk 0's bisection waits until chunks 1 and 2 are on their way.
      if (s.every((i) => i < 25)) await hold.wait(signal)
      return answerAll(body, 'delete')
    }
    const runId = await preflight(75)
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    const sent = (first: number) => router.claudeBodies.some((b) => subjectsOf(b)[0] === first)
    await waitFor(() => sent(25) && sent(50))
    hold.open()
    await waitFor(async () => (await full())?.done === true)
    const p = (await full())!
    expect(p.aiPlan.filter((x) => x.unclassified).map((x) => x.emailId)).toEqual(['m7'])
    expect(Object.keys(await getRefusedEmails())).toEqual(['m7'])
  })

  it('an ordinary answer that fails to parse proves the prompt too: it clears a remembered suspicion', async () => {
    router.claude = async () => refusalResponse('cyber')
    await runBatch(3) // the probe is refused too → suspicion remembered
    expect(await getRefusalSuspicion()).not.toBeNull()

    router.claude = async () => sseResponse(sseText(answerEvents('這不是 JSON')))
    const p = await runBatch(3)
    expect(p.aiError!.code).toBe('CLASSIFIER_PARSE_ERROR')
    expect(await getRefusalSuspicion()).toBeNull()
  })

  it('a probe that hits a run-fatal error ends the run with that error', async () => {
    router.claude = async ({ body }) =>
      isProbe(body)
        ? jsonErrorResponse(401, 'authentication_error', 'invalid x-api-key')
        : refusalResponse('cyber')
    const p = await runBatch(30)
    expect(router.claudeBodies).toHaveLength(2) // chunk 0 + the probe; chunk 1 never sent
    expect(p.aiError!.code).toBe('CLASSIFIER_HTTP_401')
    expect(p.aiPlan).toHaveLength(30)
    expect(p.aiPlan.every((x) => x.unclassified)).toBe(true)
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
