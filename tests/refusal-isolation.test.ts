// Refusal isolation + refused-email ledger (2026-11-06, D1).
//
//   - a refused subset is bisected until the refused item(s) stand alone;
//     every other subset is answered normally;
//   - extra calls stop at the per-chunk and the shared per-run budget — what
//     is left unsplit is reported, never silently dropped;
//   - parse failures stay chunk-local, a fatal error or a cancel stops the
//     isolation at once;
//   - the ledger keeps refused ids for REFUSED_EMAILS_TTL_MS, newest first,
//     capped.

import { describe, expect, it } from 'vitest'
import { ClassifierError } from '@/shared/classifier'
import { REFUSED_EMAILS_CAP, REFUSED_EMAILS_TTL_MS } from '@/shared/constants'
import { refusalCategoriesText, refusalCategoryLabel } from '@/shared/refusal-category'
import {
  addRefusedEmails,
  clearRefusalSuspicion,
  getRefusalSuspicion,
  getRefusedEmailIds,
  getRefusedEmails,
  setRefusalSuspicion,
  type RefusalSuspicion,
} from '@/shared/refusal-ledger'
import {
  batchRefusedReason,
  classifyIsolatingRefusals,
  createRefusalGate,
  describeRefusals,
  previouslyRefusedReason,
  refusedEmailReason,
  sharedPromptProbeEmail,
} from '@/background/refusal-isolation'

const refusal = (category: string | null = 'cyber', promptKey?: string) =>
  new ClassifierError('REFUSAL', `Claude 拒絕分類此批（類別：${category ?? '未提供'}）`, {
    category,
    ...(promptKey ? { promptKey } : {}),
  })

/** A classify() double: refuses any subset containing a culprit; records calls. */
function classifier(culprits: number[], opts: { failOn?: (s: number[]) => unknown } = {}) {
  const calls: number[][] = []
  const classify = async (subset: number[]) => {
    calls.push(subset)
    const failure = opts.failOn?.(subset)
    if (failure) throw failure
    if (subset.some((x) => culprits.includes(x))) throw refusal()
    return { answered: subset }
  }
  return { calls, classify }
}

const range = (n: number) => Array.from({ length: n }, (_, i) => i)
const budget = (remaining = 100) => ({ remaining })
const open = { isCancelled: () => false }

describe('classifyIsolatingRefusals', () => {
  it('bisects down to the refused item; every other subset is answered', async () => {
    const { calls, classify } = classifier([7])
    const r = await classifyIsolatingRefusals(range(25), classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      ...open,
    })
    expect(r.refused.map((x) => x.item)).toEqual([7])
    expect(r.refused[0]!.error.category).toBe('cyber')
    expect(r.answered.flatMap((a) => a.items).sort((a, b) => a - b)).toEqual(range(25).filter((x) => x !== 7))
    // 1 + 2 per halving level (25 → 13 → 6 → 3 → 2 → 1).
    expect(r.calls).toBe(11)
    expect(calls).toHaveLength(11)
    expect(r.unisolated).toEqual([])
    expect(r.fatal).toBeUndefined()
    expect(r.aborted).toBe(false)
  })

  it('finds several culprits, and a lone refused item needs no extra call', async () => {
    const two = classifier([2, 9])
    const r = await classifyIsolatingRefusals(range(10), two.classify, {
      maxExtraCalls: 20,
      runBudget: budget(),
      ...open,
    })
    expect(r.refused.map((x) => x.item)).toEqual([2, 9])
    expect(r.answered.flatMap((a) => a.items)).toHaveLength(8)

    const lone = classifier([0])
    const one = await classifyIsolatingRefusals([0], lone.classify, {
      maxExtraCalls: 0,
      runBudget: budget(0),
      ...open,
    })
    expect(one.refused.map((x) => x.item)).toEqual([0])
    expect(one.calls).toBe(1)
  })

  it('leaves refused subsets unsplit once the chunk budget is spent', async () => {
    const { calls, classify } = classifier(range(8)) // everything refused
    const r = await classifyIsolatingRefusals(range(8), classify, {
      maxExtraCalls: 3,
      runBudget: budget(),
      ...open,
    })
    expect(calls).toHaveLength(4) // the first call + 3 extra
    expect(r.answered).toEqual([])
    // Everything not refused on its own is reported, nothing is lost.
    const reported = [...r.refused.map((x) => [x.item]), ...r.unisolated.map((u) => u.items)].flat()
    expect(reported.sort((a, b) => a - b)).toEqual(range(8))
    expect(r.unisolated.every((u) => u.error.code === 'REFUSAL')).toBe(true)
  })

  it('shares the run budget across chunks', async () => {
    const run = budget(2)
    const first = classifier([1])
    await classifyIsolatingRefusals(range(4), first.classify, { maxExtraCalls: 12, runBudget: run, ...open })
    expect(run.remaining).toBe(0)
    const second = classifier([1])
    const r = await classifyIsolatingRefusals(range(4), second.classify, {
      maxExtraCalls: 12,
      runBudget: run,
      ...open,
    })
    expect(second.calls).toHaveLength(1)
    expect(r.unisolated.flatMap((u) => u.items)).toEqual(range(4))
  })

  it('asks before every extra call (a suspended isolation leaves subsets unsplit)', async () => {
    let allowed = true
    const { calls, classify } = classifier([0])
    const r = await classifyIsolatingRefusals(range(8), classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      extraCallsAllowed: () => {
        const ok = allowed
        allowed = false // allow exactly one extra call
        return ok
      },
      ...open,
    })
    expect(calls.map((s) => s.join())).toEqual(['0,1,2,3,4,5,6,7', '0,1,2,3'])
    expect(r.unisolated.map((u) => u.items)).toEqual([[0, 1], [2, 3], [4, 5, 6, 7]])
  })

  it('keeps a parse failure chunk-local and goes on with the other half', async () => {
    const { classify } = classifier([0], {
      failOn: (s) => (s.length === 2 && s[0] === 2 ? new ClassifierError('PARSE_ERROR', 'bad json') : null),
    })
    const r = await classifyIsolatingRefusals(range(4), classify, { maxExtraCalls: 12, runBudget: budget(), ...open })
    expect(r.parseFailed.map((p) => p.items)).toEqual([[2, 3]])
    expect(r.refused.map((x) => x.item)).toEqual([0])
    expect(r.answered.map((a) => a.items)).toEqual([[1]])
  })

  it('stops at a run-fatal error and at a cancel', async () => {
    const fatal = new ClassifierError('HTTP_401', 'invalid x-api-key')
    const f = classifier([0], { failOn: (s) => (s[0] === 2 ? fatal : null) })
    const r = await classifyIsolatingRefusals(range(4), f.classify, { maxExtraCalls: 12, runBudget: budget(), ...open })
    expect(r.fatal).toBe(fatal)
    expect(f.calls.map((s) => s.join())).toEqual(['0,1,2,3', '0,1', '0', '1', '2,3'])

    let cancelled = false
    const c = classifier([0], {
      failOn: (s) => {
        if (s.length === 2) cancelled = true
        return null
      },
    })
    const rc = await classifyIsolatingRefusals(range(4), c.classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      isCancelled: () => cancelled,
    })
    expect(rc.aborted).toBe(true)
    expect(c.calls).toHaveLength(2) // nothing after the cancel

    const aborted = classifier([], { failOn: () => new ClassifierError('ABORTED', '分類已取消') })
    const ra = await classifyIsolatingRefusals(range(4), aborted.classify, { maxExtraCalls: 12, runBudget: budget(), ...open })
    expect(ra.aborted).toBe(true)
  })

  // 2026-11-06 (K1 / K2): bisecting and blaming wait for the run's verdict.
  it('asks once, at the first refusal, whether the refusal can be pinned on these items', async () => {
    const asked: ClassifierError[] = []
    const { calls, classify } = classifier(range(8)) // the shared prompt: everything refused
    const r = await classifyIsolatingRefusals(range(8), classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      attributable: async (e) => {
        asked.push(e)
        return false
      },
      ...open,
    })
    expect(calls).toHaveLength(1) // no bisection at all
    expect(asked).toHaveLength(1)
    expect(r.refused).toEqual([])
    expect(r.unisolated.map((u) => u.items)).toEqual([range(8)])

    // A single item refused on its own is only reported as `refused` when
    // attributable — otherwise it stays unisolated (not blamed).
    const lone = classifier([0])
    const notBlamed = await classifyIsolatingRefusals([0], lone.classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      attributable: () => false,
      ...open,
    })
    expect(notBlamed.refused).toEqual([])
    expect(notBlamed.unisolated.map((u) => u.items)).toEqual([[0]])
    const blamed = await classifyIsolatingRefusals([0], classifier([0]).classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      attributable: () => true,
      ...open,
    })
    expect(blamed.refused.map((x) => x.item)).toEqual([0])
  })

  it('a cancel while the verdict was pending stops the isolation', async () => {
    let cancelled = false
    const { calls, classify } = classifier([3])
    const r = await classifyIsolatingRefusals(range(4), classify, {
      maxExtraCalls: 12,
      runBudget: budget(),
      attributable: async () => {
        cancelled = true
        return true
      },
      isCancelled: () => cancelled,
    })
    expect(r.aborted).toBe(true)
    expect(calls).toHaveLength(1)
  })
})

describe('RefusalGate', () => {
  /** A gate plus `verdict(refusal)` that passes this test's probe (counted). */
  function gateWith(
    over: Partial<Parameters<typeof createRefusalGate>[0]> & { probe?: () => Promise<void> } = {},
  ) {
    const { probe: probeImpl, ...deps } = over
    const remembered: RefusalSuspicion[] = []
    let forgotten = 0
    let probes = 0
    const gate = createRefusalGate({
      model: 'claude-sonnet-5-5',
      remembered: null,
      remember: async (s) => void remembered.push(s),
      forget: async () => void forgotten++,
      ttlMs: 1000,
      now: () => 5_000,
      ...deps,
    })
    const probe = async () => {
      probes++
      await probeImpl?.()
    }
    return {
      gate,
      verdict: (r: ClassifierError) => gate.verdict(r, probe),
      remembered,
      forgotten: () => forgotten,
      probes: () => probes,
    }
  }

  it('one probe for the whole run, however many lanes ask at once; an answered probe means isolate', async () => {
    let release!: () => void
    const held = new Promise<void>((r) => (release = r))
    const { gate, verdict, probes } = gateWith({ probe: () => held })
    gate.noteRefusal()
    expect(gate.blocksFanOut()).toBe(true)
    // Each lane brings its own probe; only the first one ever runs.
    let otherLaneProbes = 0
    const otherLane = async () => void otherLaneProbes++
    const verdicts = Promise.all([verdict(refusal()), gate.verdict(refusal(), otherLane), verdict(refusal())])
    release()
    expect(await verdicts).toEqual(['isolate', 'isolate', 'isolate'])
    expect(probes()).toBe(1)
    expect(otherLaneProbes).toBe(0)
    expect(gate.blocksFanOut()).toBe(false)
    expect(gate.sharedConcluded()).toBe(false)
  })

  it('a refused probe means shared: remembered with the prompt key; no second probe', async () => {
    const { gate, verdict, remembered, probes } = gateWith({
      probe: async () => {
        throw refusal('reasoning_extraction', 'key-A')
      },
    })
    gate.noteRefusal()
    expect(await verdict(refusal('cyber', 'key-A'))).toBe('shared')
    expect(await verdict(refusal('cyber', 'key-A'))).toBe('shared')
    expect(probes()).toBe(1)
    expect(gate.sharedConcluded()).toBe(true)
    expect(gate.blocksFanOut()).toBe(true)
    expect(remembered).toEqual([
      { model: 'claude-sonnet-5-5', promptKey: 'key-A', category: 'reasoning_extraction', at: 5_000 },
    ])
  })

  it('a remembered suspicion for the same model + prompt (within its TTL) skips the probe', async () => {
    const s = { model: 'claude-sonnet-5-5', promptKey: 'key-A', category: 'cyber', at: 4_500 }
    const same = gateWith({ remembered: s })
    expect(await same.verdict(refusal('cyber', 'key-A'))).toBe('shared')
    expect(same.probes()).toBe(0)
    // Another prompt, another model or an expired suspicion: probe again.
    const other = gateWith({ remembered: s })
    expect(await other.verdict(refusal('cyber', 'key-B'))).toBe('isolate')
    expect(other.probes()).toBe(1)
    const otherModel = gateWith({ remembered: { ...s, model: 'claude-opus-5-5' } })
    expect(await otherModel.verdict(refusal('cyber', 'key-A'))).toBe('isolate')
    const expired = gateWith({ remembered: { ...s, at: 3_000 } })
    expect(await expired.verdict(refusal('cyber', 'key-A'))).toBe('isolate')
  })

  it('an answer clears a remembered suspicion for its prompt and lifts the shared verdict', async () => {
    const s = { model: 'claude-sonnet-5-5', promptKey: 'key-A', category: 'cyber', at: 4_500 }
    const { gate, verdict, forgotten } = gateWith({ remembered: s })
    gate.noteRefusal()
    expect(await verdict(refusal('cyber', 'key-A'))).toBe('shared')
    gate.noteSuccess('key-B') // another prompt's answer proves nothing about key-A's suspicion
    expect(forgotten()).toBe(0)
    gate.noteSuccess('key-A')
    expect(forgotten()).toBe(1)
    expect(await verdict(refusal('cyber', 'key-A'))).toBe('isolate')
    expect(gate.sharedConcluded()).toBe(false)
    expect(gate.blocksFanOut()).toBe(false)
  })

  it('a probe that fails otherwise is unavailable (no isolation) and hands its fatal error out once', async () => {
    const fatal = new ClassifierError('HTTP_401', 'invalid x-api-key')
    const { gate, verdict } = gateWith({
      probe: async () => {
        throw fatal
      },
    })
    expect(await verdict(refusal())).toBe('unavailable')
    expect(gate.takeProbeFailure()).toEqual({ error: fatal })
    expect(gate.takeProbeFailure()).toBeNull()
    const aborted = gateWith({
      probe: async () => {
        throw new ClassifierError('ABORTED', '分類已取消')
      },
    })
    expect(await aborted.verdict(refusal())).toBe('unavailable')
    expect(aborted.gate.takeProbeFailure()).toBeNull()
  })

  // Round-3 review item 4: the scheduler's one-lane hold has to lift the
  // moment the run is proven — the probe set the state but told no one.
  it('onProven fires once, on the first proof — a parse-failed probe answer included', async () => {
    let proven = 0
    const { gate, verdict } = gateWith({
      onProven: () => void proven++,
      probe: async () => {
        throw new ClassifierError('PARSE_ERROR', 'bad json')
      },
    })
    gate.noteRefusal()
    expect(await verdict(refusal())).toBe('isolate')
    expect(proven).toBe(1)
    gate.noteSuccess('key-A')
    expect(proven).toBe(1) // already proven: no second signal
    const answered = gateWith({ onProven: () => void proven++ })
    answered.gate.noteSuccess(undefined)
    expect(proven).toBe(2)
    const probed = gateWith({ onProven: () => void proven++ })
    expect(await probed.verdict(refusal())).toBe('isolate')
    expect(proven).toBe(3)
  })

  it('a parse failure of the probe still proves the prompt was accepted', async () => {
    const { verdict } = gateWith({
      probe: async () => {
        throw new ClassifierError('PARSE_ERROR', 'bad json')
      },
    })
    expect(await verdict(refusal())).toBe('isolate')
  })

  it('the probe email is harmless synthetic mail, never a real Outlook id', () => {
    const probe = sharedPromptProbeEmail()
    expect(probe.Id).toBe('mail-organizer:shared-prompt-probe')
    expect(probe.Subject).toBe('下週會議時間確認')
    expect(probe.bodyText).toBeTruthy()
  })
})

describe('refusal texts', () => {
  it('name the subject (shortened) and a neutral category label — never the raw API id', () => {
    expect(refusedEmailReason('Exploit kit 報價', 'cyber')).toBe('Claude 拒絕分類「Exploit kit 報價」（類別：資訊安全），請手動處理')
    expect(previouslyRefusedReason('', null)).toBe('Claude 先前拒絕分類「(空主旨)」（類別：未提供），這次未再送出，請手動處理')
    expect(refusedEmailReason('x'.repeat(60), 'bio')).toContain(`「${'x'.repeat(40)}…」`)
    expect(batchRefusedReason('frontier_llm')).toBe(
      'Claude 拒絕了這封所在的整批分類請求（類別：AI 模型開發），無法確定是哪一封；下一批會再試',
    )
    expect(
      describeRefusals({
        named: [
          { subject: 'a', category: 'cyber' },
          { subject: 'b', category: 'bio' },
          { subject: 'c', category: 'cyber' },
          { subject: 'd', category: null },
        ],
        unattributed: 0,
        unattributedCategories: [],
        shared: false,
      }),
    ).toBe('Claude 拒絕分類 4 封郵件：「a」、「b」、「c」 等（類別：資訊安全、生物安全、未提供），已標為未分類，請手動處理')
  })

  it('category labels: one helper for every known id; unknown → 其他, none → 未提供', () => {
    expect(
      ['cyber', 'bio', 'frontier_llm', 'reasoning_extraction', 'general_harms'].map(refusalCategoryLabel),
    ).toEqual(['資訊安全', '生物安全', 'AI 模型開發', '推理過程擷取', '一般使用政策'])
    expect(refusalCategoryLabel('some_new_category')).toBe('其他')
    expect(refusalCategoryLabel(null)).toBe('未提供')
    expect(refusalCategoryLabel('')).toBe('未提供')
    expect(refusalCategoriesText(['cyber', null, 'cyber', 'bio'])).toBe('資訊安全、未提供、生物安全')
    expect(refusalCategoriesText([])).toBe('未提供')
  })

  it('the run banner counts every unattributed email and category, with or without named ones (K4)', () => {
    expect(
      describeRefusals({ named: [], unattributed: 125, unattributedCategories: ['cyber', 'bio'], shared: false }),
    ).toBe(
      'Claude 拒絕了 125 封郵件所在的整批請求（類別：資訊安全、生物安全），無法確定是哪一封；已標為未分類，下一批會再試',
    )
    expect(
      describeRefusals({
        named: [{ subject: 'a', category: 'cyber' }],
        unattributed: 24,
        unattributedCategories: [null],
        shared: false,
      }),
    ).toBe(
      'Claude 拒絕分類 1 封郵件：「a」（類別：資訊安全），已標為未分類，請手動處理；' +
        '另有 24 封郵件所在的整批請求被拒（類別：未提供），無法確定是哪一封，下一批會再試',
    )
    const shared = describeRefusals({ named: [], unattributed: 30, unattributedCategories: ['cyber'], shared: true })!
    expect(shared).toContain('Claude 拒絕分類其中 30 封郵件（類別：資訊安全）')
    expect(shared).toContain('連只含一封測試郵件的請求也被拒絕')
    expect(shared).toContain('關閉「AI 是否參考既有規則範例」')
    expect(describeRefusals({ named: [], unattributed: 0, unattributedCategories: [], shared: false })).toBeNull()
  })
})

describe('refused-email ledger', () => {
  it('keeps entries for the TTL, newest wins, capped', async () => {
    const t0 = Date.UTC(2026, 10, 6)
    await addRefusedEmails([{ emailId: 'm1', category: 'cyber', model: 'claude-sonnet-5-5' }], t0)
    await addRefusedEmails([{ emailId: 'm1', category: 'bio', model: 'claude-haiku-5-5' }, { emailId: '', category: null, model: 'x' }], t0 + 1000)
    expect(await getRefusedEmails(t0 + 2000)).toEqual({
      m1: { at: t0 + 1000, category: 'bio', model: 'claude-haiku-5-5' },
    })
    expect(await getRefusedEmails(t0 + 1000 + REFUSED_EMAILS_TTL_MS - 1)).toHaveProperty('m1')
    expect(await getRefusedEmails(t0 + 1000 + REFUSED_EMAILS_TTL_MS + 1)).toEqual({})

    const many = Array.from({ length: REFUSED_EMAILS_CAP + 5 }, (_, i) => ({
      emailId: `x${i}`,
      category: null,
      model: 'claude-sonnet-5-5',
    }))
    await addRefusedEmails(many, t0 + 5000)
    expect(Object.keys(await getRefusedEmails(t0 + 6000))).toHaveLength(REFUSED_EMAILS_CAP)
  })

  it('ignores a malformed stored value', async () => {
    await chrome.storage.local.set({ refusedEmails: { m1: { at: 'soon' }, m2: 'x', m3: { at: Date.now(), category: 'cyber', model: 'claude-sonnet-5-5' } } })
    expect(Object.keys(await getRefusedEmails())).toEqual(['m3'])
    await chrome.storage.local.set({ refusedEmails: ['m1'] })
    expect(await getRefusedEmails()).toEqual({})
  })

  it('lists the ids ledgered for one model (what the inbox scan leaves out)', async () => {
    await addRefusedEmails([
      { emailId: 'a', category: 'cyber', model: 'claude-sonnet-5-5' },
      { emailId: 'b', category: null, model: 'claude-opus-5-5' },
    ])
    expect([...(await getRefusedEmailIds('claude-sonnet-5-5'))]).toEqual(['a'])
    expect([...(await getRefusedEmailIds('claude-haiku-5-5'))]).toEqual([])
  })
})

describe('shared-prompt refusal suspicion storage', () => {
  it('round-trips, rejects malformed values, clears', async () => {
    expect(await getRefusalSuspicion()).toBeNull()
    const s = { model: 'claude-sonnet-5-5', promptKey: 'abc', category: null, at: 1 }
    await setRefusalSuspicion(s)
    expect(await getRefusalSuspicion()).toEqual(s)
    await chrome.storage.local.set({ refusalSuspicion: { model: 'x', promptKey: '', category: null, at: 1 } })
    expect(await getRefusalSuspicion()).toBeNull()
    await setRefusalSuspicion(s)
    await clearRefusalSuspicion()
    expect(await chrome.storage.local.get('refusalSuspicion')).toEqual({})
  })
})
