// Partial execution vs classify progress (2026-11-05, items 2 + 14).
//
// Since 「先執行 N 個」 the user may execute while AI still runs. The old
// residue guard compared progress.startedAt with the execute's start: an AI
// run that finished DURING the execute looked like residue, so 繼續歸檔 / the
// DoneScreen prefetch threw it away and re-classified the same mail (paid
// twice), and the clearExecuteState purge deleted a still-running run's
// progress. Now the run records the ids submitted to execute and never offers
// them again; only a finished run with nothing left is purged.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { addRules, newRule } from '@/shared/rules'
import type { PlanItem } from '@/shared/types'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'
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
afterEach(() => vi.unstubAllGlobals())

type FullView = {
  runId: string
  done: boolean
  ruleCount: number
  aiCount: number
  aiCursor: number
  rulePlan: PlanItem[]
  aiPlan: PlanItem[]
  folderTree?: unknown
}
const view = async (v: 'full' | 'poll' | 'meta', aiFrom = 0) =>
  ((await sw.send({ type: 'getAiClassifyProgress', view: v, aiFrom })).data ?? null) as FullView | null

async function executeAndWait(plan: PlanItem[]) {
  const r = await sw.send({ type: 'startExecute', plan })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await waitFor(async () => {
    const s = (await sw.send({ type: 'getExecuteState' })).data as { state: { inProgress: boolean } | null }
    return s.state != null && !s.state.inProgress
  })
}

describe('「先執行 N 個」 while AI still classifies', () => {
  it('keeps the rest of the run, never re-offers executed mail, purges only empty residue', async () => {
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
          ...[0, 1, 2].map((i) =>
            inboxMail(i, { From: { EmailAddress: { Address: `a${i}@rule.example`, Name: 'R' } } }),
          ),
          ...[3, 4, 5].map((i) => inboxMail(i)),
        ],
      },
    ]
    const pf = await sw.send({ type: 'classifyPreflight', batchSize: 6 })
    const { runId, rulePlan } = pf.data as { runId: string; rulePlan: PlanItem[] }
    expect(rulePlan.map((p) => p.emailId)).toEqual(['m0', 'm1', 'm2'])
    await markBodiesFetched()

    const g = gate()
    router.claude = async ({ body, signal }) => {
      await g.wait(signal)
      return answerAll(body)
    }
    await sw.send({ type: 'classifyAi', runId, excludeIds: [] })
    await waitFor(() => router.claudeBodies.length === 1)

    // The popup's 「先執行 3 個」: only the rule hits exist yet.
    await executeAndWait(rulePlan)
    // Read-time exclusion: the loop hasn't written since, yet the executed
    // rule hits are already gone from what the popup can see.
    expect((await view('meta'))!.ruleCount).toBe(0)

    // Item 14: the executing-screen poll no longer ships the whole plan.
    const st = (await sw.send({ type: 'getExecuteState' })).data as { state: Record<string, unknown> }
    expect(st.state).not.toHaveProperty('plan')
    expect(st.state).toHaveProperty('results')

    // 繼續歸檔 / 完成 drop the execute state while AI still runs: the live
    // run must survive (the old purge deleted its done:false progress).
    await sw.send({ type: 'clearExecuteState' })
    expect((await view('meta'))!.done).toBe(false)

    g.open() // AI finishes AFTER the execute started
    await waitFor(async () => (await view('meta'))?.done === true)

    const full = (await view('full'))!
    expect(full.runId).toBe(runId)
    expect(full.rulePlan).toEqual([])
    expect(full.aiPlan.map((p) => p.emailId)).toEqual(['m3', 'm4', 'm5'])
    expect(full.folderTree).toBeDefined()
    // Write-time exclusion too: the stored progress never re-gained them.
    const stored = (await chrome.storage.session.get('aiClassifyProgress')).aiClassifyProgress as {
      rulePlan: PlanItem[]
    }
    expect(stored.rulePlan).toEqual([])

    // Item 14: the poll view carries neither the tree nor the rule plan, and
    // only AI items past the cursor.
    const poll = (await view('poll', 0))!
    expect(poll).not.toHaveProperty('folderTree')
    expect(poll).not.toHaveProperty('rulePlan')
    expect(poll.aiPlan).toHaveLength(3)
    expect((await view('poll', full.aiCursor))!.aiPlan).toEqual([])

    // A finished run that still has unexecuted items is NOT residue: it stays
    // for 繼續歸檔 instead of being re-classified (the pay-twice bug).
    await sw.send({ type: 'clearExecuteState' })
    expect((await view('meta'))!.aiCount).toBe(3)
    expect(router.claudeBodies).toHaveLength(1)

    // Execute the rest → now it IS residue and the purge drops the run.
    await executeAndWait(full.aiPlan)
    expect((await view('meta'))!.aiCount).toBe(0)
    await sw.send({ type: 'clearExecuteState' })
    expect(await view('full')).toBeNull()
    expect(await chrome.storage.session.get(['aiClassifyProgress', 'classifyRun'])).toEqual({})
  })
})
