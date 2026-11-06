// 先廣後窄 silencing must outlast one batch (2026-11-06, B1).
//
// 1. client.example → 甲案 (broad domain rule).
// 2. The user corrects one 乙案 mail: compound(domain + subject) → 乙案 is
//    learned and the broad rule is auto-disabled as 'superseded'.
// 3. Next batch the AI files a 甲案 mail → the broad rule was REVIVED (the
//    revive guard only looked at same-(type, signal) twins; the compound has
//    another signal).
// 4. The next 乙案 mail with a new subject was routed to 甲案 by the rule and
//    never reached the AI.
// Same-batch variant: an A filing revived broad→A while a B filing minted
// domain→B — two live conflicting domain rules.

import { describe, expect, it } from 'vitest'
import {
  generateAiConfirmedRules,
  generateAiOverrideRules,
  type ExecuteItemResult,
} from '@/background/execute'
import { normalizeSubject } from '@/shared/normalize'
import {
  addRulesFilteringTombstones,
  buildRuleIndex,
  decodeCompound,
  encodeCompound,
  listRules,
  matchEmailWithIndex,
  newRule,
} from '@/shared/rules'
import { setRules } from '@/shared/storage'
import type { Email, PlanItem, Rule, ThreadMemoryEntry } from '@/shared/types'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const A = '03/甲案'
const B = '03/乙案'
const SUBJ_B1 = '乙案股權移轉進度報告'
const SUBJ_A = '甲案合約審閱意見回覆'

function broad(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: FID,
      targetFolderPath: A,
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    id: 'broad',
    createdAt: '2026-01-01T00:00:00.000Z',
    matchCount: 12,
    ...over,
  }
}

const superseded = (over: Partial<Rule> = {}) =>
  broad({ enabled: false, autoDisabledReason: 'superseded', autoDisabledAt: '2026-11-01T00:00:00.000Z', ...over })

function compound(subject: string, target: string, over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: subject },
      ]),
      targetFolderId: FID,
      targetFolderPath: target,
      confidence: 0.85,
      source: 'ai_overridden',
    }),
    createdAt: '2026-11-01T00:00:00.000Z',
    ...over,
  }
}

let seq = 0
function aiFiling(subject: string, target: string, over: Partial<PlanItem> = {}): PlanItem {
  seq++
  return {
    emailId: `e${seq}`,
    emailSubject: subject,
    emailFrom: 'pm@client.example',
    action: 'move',
    targetFolderId: FID,
    targetFolderPath: target,
    confidence: 0.9,
    reason: '',
    source: 'ai',
    conversationId: `c${seq}`,
    ...over,
  }
}

/** A row the broad rule routed to 甲案, re-targeted by the user (PlanRow shape). */
const correction = (subject: string, target: string) =>
  aiFiling(subject, target, { originalRuleId: 'broad', userTouched: true, confidence: 0.7 })

const moved = (item: PlanItem): ExecuteItemResult => ({
  emailId: item.emailId,
  subject: item.emailSubject ?? '',
  action: 'move',
  status: 'moved',
})

/** Subject memory before the batch: `subject` already filed once into `folderPath`. */
const seenBefore = (...entries: Array<[string, string]>): Record<string, ThreadMemoryEntry> =>
  Object.fromEntries(
    entries.map(([subject, folderPath]) => [
      normalizeSubject(subject),
      { folderId: FID, folderPath, lastFiledAt: '2026-10-01T00:00:00.000Z', timesFiled: 1, conflictCount: 0, stableStreak: 1 },
    ]),
  )

const inbound = (subject: string): Email =>
  ({
    Id: 'm1',
    Subject: subject,
    BodyPreview: '',
    From: { EmailAddress: { Address: 'pm@client.example' } },
    ToRecipients: [],
    ReceivedDateTime: '2026-11-06T00:00:00Z',
    ParentFolderId: 'inbox',
  }) as Email

const liveDomainRules = async () => (await listRules()).filter((r) => r.enabled && r.type === 'domain')

describe('a superseded anchor stays narrow (B1)', () => {
  it('repro: after the correction, an AI 甲案 filing neither revives nor re-mints the broad rule', async () => {
    await setRules([broad()])
    const fix = correction(SUBJ_B1, B)
    expect(await generateAiOverrideRules([fix], [moved(fix)])).toMatchObject({ added: 1, disabled: 1 })

    const a = aiFiling(SUBJ_A, A)
    const r = await generateAiConfirmedRules([a], [moved(a)])
    expect(r.revived).toBe(0) // was 1
    const rules = await listRules()
    expect(rules.find((x) => x.id === 'broad')).toMatchObject({ enabled: false, autoDisabledReason: 'superseded' })
    expect(await liveDomainRules()).toHaveLength(0)
    // The next 乙案 mail with a NEW subject goes to the AI instead of 甲案.
    expect(matchEmailWithIndex(inbound('乙案補充文件寄送通知'), buildRuleIndex(rules))).toBeNull()
  })

  it('on its second observation the 甲案 subject is learned NARROW (domain + subject), not broad', async () => {
    await setRules([superseded(), compound(SUBJ_B1, B)])
    const a = aiFiling(SUBJ_A, A)
    const r = await generateAiConfirmedRules([a], [moved(a)], { priorSubjectMemory: seenBefore([SUBJ_A, A]) })
    expect(r).toMatchObject({ added: 1, revived: 0 })
    const learned = (await listRules()).find((x) => x.enabled && x.targetFolderPath === A)!
    expect(learned.type).toBe('compound')
    expect(decodeCompound(learned.signal)?.conditions).toEqual([
      { type: 'domain', value: 'client.example' },
      { type: 'subject_keyword', value: normalizeSubject(SUBJ_A) },
    ])
  })

  it('a 乙案 filing on a superseded anchor does not mint a fresh broad rule pointing elsewhere', async () => {
    await setRules([superseded()])
    const b = aiFiling('乙案開會時間確認通知', B)
    const r = await generateAiConfirmedRules([b], [moved(b)])
    expect(r.added).toBe(0) // was 1: domain client.example → 乙案
    expect(await liveDomainRules()).toHaveLength(0)
  })

  it('a correction with a subject too short to narrow learns nothing on the confirmed path either', async () => {
    await setRules([superseded()])
    const a = aiFiling('RE: 回覆', A)
    const r = await generateAiConfirmedRules([a], [moved(a)])
    expect(r).toMatchObject({ added: 0, revived: 0 })
    expect(await liveDomainRules()).toHaveLength(0)
  })

  it('while an AUTO narrow rule on the anchor routes elsewhere, the anchor stays narrow after the broad row is purged', async () => {
    await setRules([compound(SUBJ_B1, B)]) // the superseded broad rule is gone (dormant purge)
    const a = aiFiling(SUBJ_A, A)
    const r = await generateAiConfirmedRules([a], [moved(a)], { priorSubjectMemory: seenBefore([SUBJ_A, A]) })
    expect(r.added).toBe(1)
    expect(await liveDomainRules()).toHaveLength(0) // was: domain client.example → 甲案
  })

  it('a hand-written compound on the domain does not block the broad rule (broad + own exception is legitimate)', async () => {
    await setRules([compound('請款通知單', '03/帳務', { source: 'user_manual' })])
    const a = aiFiling(SUBJ_A, A)
    expect((await generateAiConfirmedRules([a], [moved(a)])).added).toBe(1)
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([A])
  })
})

describe('in-batch candidates and revives are part of the conflict guard (B1)', () => {
  it('one batch filing the same domain into two folders learns no broad rule for either', async () => {
    await setRules([])
    const a = aiFiling(SUBJ_A, A)
    const b = aiFiling('乙案開會時間確認通知', B)
    const r = await generateAiConfirmedRules([a, b], [moved(a), moved(b)])
    expect(r.added).toBe(0) // was 2 conflicting domain rules
    // With both subjects seen before, each side is learned narrow.
    await setRules([])
    const r2 = await generateAiConfirmedRules([a, b], [moved(a), moved(b)], {
      priorSubjectMemory: seenBefore([SUBJ_A, A], ['乙案開會時間確認通知', B]),
    })
    expect(r2.added).toBe(2)
    expect((await listRules()).every((x) => x.type === 'compound')).toBe(true)
  })

  it('a case-numbered filing elsewhere in the batch is not an anchor conflict (structural tiers outrank domain)', async () => {
    await setRules([])
    const general = aiFiling(SUBJ_A, A)
    const court = aiFiling('112訴204 言詞辯論期日通知', '03/112訴204')
    const r = await generateAiConfirmedRules([general, court], [moved(general), moved(court)])
    expect(r.added).toBe(2) // the case rule + the broad rule for the general mail
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([A])
  })

  it('same-batch variant: an A filing does not revive broad→A while a B filing is in the batch', async () => {
    await setRules([broad({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: '2026-11-01T00:00:00.000Z' })])
    const a = aiFiling(SUBJ_A, A)
    const b = aiFiling('乙案開會時間確認通知', B)
    const r = await generateAiConfirmedRules([a, b], [moved(a), moved(b)])
    expect(r).toMatchObject({ added: 0, revived: 0 }) // was: broad→A revived + domain→B minted
    expect(await liveDomainRules()).toHaveLength(0)
  })

  it('the central guard refuses a plain-anchor revive while a compound on the anchor routes elsewhere (any caller, e.g. the scan)', async () => {
    await setRules([superseded(), compound(SUBJ_B1, B)])
    const res = await addRulesFilteringTombstones([broad({ id: 'scan-candidate', source: 'auto_scan', enabled: true })])
    expect(res.reEnabled).toHaveLength(0)
    expect(res.added).toHaveLength(0)
    expect(await liveDomainRules()).toHaveLength(0)
  })

  it('the central guard refuses a revive that conflicts with another candidate of the same call', async () => {
    await setRules([broad({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: '2026-11-01T00:00:00.000Z' })])
    const res = await addRulesFilteringTombstones([
      broad({ id: 'cand-a', enabled: true }),
      broad({ id: 'cand-b', enabled: true, targetFolderPath: B }),
    ])
    expect(res.reEnabled).toHaveLength(0)
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([B])
  })

  it('a revive with no conflict anywhere still goes through', async () => {
    await setRules([superseded({ autoDisabledReason: 'high-error-rate' })])
    const res = await addRulesFilteringTombstones([broad({ id: 'cand', enabled: true })])
    expect(res.reEnabled.map((x) => x.id)).toEqual(['broad'])
    expect(res.revivedFrom[0]).toMatchObject({ enabled: false, autoDisabledReason: 'high-error-rate' })
  })
})

/** A row `r` routed (un-edited) — the shape preflight builds. */
const routedBy = (r: Rule, subject: string, over: Partial<PlanItem> = {}): PlanItem =>
  aiFiling(subject, r.targetFolderPath, {
    source: 'rule',
    ruleId: r.id,
    originalRuleId: r.id,
    matchedRule: { id: r.id, type: r.type, signal: r.signal, source: r.source, matchCount: r.matchCount, enabled: true },
    ...over,
  })

// R3 (2026-11-06): batchAnchorTargets counted a filing that a rule ranking
// above every auto domain rule had routed — the user's own exception, any
// compound — as the anchor going "elsewhere". With 500-mail batches nearly
// every batch holds one such filing per busy client, so the broad rule was
// never learned. That rule keeps routing its mail next to a broad one.
describe('filings routed by rules above the domain tier are not anchor conflicts (R3)', () => {
  it('repro: a hand-written compound exception routes one client mail, the AI confirms another → broad rule learned', async () => {
    const mine = compound('請款通知單', '03/帳務', { id: 'mine', source: 'user_manual' })
    await setRules([mine])
    const exception = routedBy(mine, '三月份請款通知單')
    const a = aiFiling(SUBJ_A, A)
    const r = await generateAiConfirmedRules([exception, a], [moved(exception), moved(a)])
    expect(r.added).toBe(1) // was 0: the anchor looked conflicted
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([A])
    // The exception still wins for its own mail.
    const rules = await listRules()
    expect(matchEmailWithIndex(inbound('四月份請款通知單'), buildRuleIndex(rules))?.rule.id).toBe('mine')
  })

  it('the same for a hand-written sender rule', async () => {
    const boss = { ...broad({ id: 'boss', type: 'sender', signal: 'pm@client.example', source: 'user_manual' }), targetFolderPath: '03/老闆' }
    await setRules([boss])
    const routed = routedBy(boss, '週會議程')
    const a = aiFiling(SUBJ_A, A)
    expect((await generateAiConfirmedRules([routed, a], [moved(routed), moved(a)])).added).toBe(1)
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([A])
  })

  it('a filing routed by an AUTO rule below the domain tier still counts', async () => {
    const subj = { ...broad({ id: 'subj', type: 'subject_keyword', signal: '週會議程確認通知', source: 'ai_confirmed' }), targetFolderPath: B }
    await setRules([subj])
    const routed = routedBy(subj, '週會議程確認通知')
    const a = aiFiling(SUBJ_A, A)
    expect((await generateAiConfirmedRules([routed, a], [moved(routed), moved(a)])).added).toBe(0)
    expect(await liveDomainRules()).toHaveLength(0)
  })
})

// R6 (2026-11-06): a rule-routed row learns structural signals on first
// sight — including a row the USER's own rule routed. A hand-written
// exception inside a case taught the whole case to follow it.
describe('rows routed by a hand-written rule teach no structural rule (R6)', () => {
  const billing = () =>
    compound('112訴204 帳單', '03/帳務', { id: 'mine', source: 'user_manual' })

  it('repro: the 帳單 mail of case 112訴204 routed by the user exception learns nothing', async () => {
    const mine = billing()
    await setRules([mine])
    const row = routedBy(mine, '112訴204 帳單 三月份')
    const r = await generateAiConfirmedRules([row], [moved(row)])
    expect(r.added).toBe(0) // was 1: compound (client.example + 112訴204) → 03/帳務
    expect(await listRules()).toEqual([mine])
    // So the case's other mail does not follow the exception into 帳務.
    expect(matchEmailWithIndex(inbound('112訴204 開庭通知'), buildRuleIndex(await listRules()))).toBeNull()
  })

  it('a row an AUTO rule routed still mints the case rule on first sight', async () => {
    const auto = broad({ targetFolderPath: '03/112訴204' })
    await setRules([auto])
    const row = routedBy(auto, '112訴204 開庭通知')
    const r = await generateAiConfirmedRules([row], [moved(row)])
    expect(r.added).toBe(1)
    const learned = (await listRules()).find((x) => x.type === 'compound')!
    expect(learned.targetFolderPath).toBe('03/112訴204')
  })

  // Narrowed (round 3): R6 skipped mail ANY hand-written rule routed, so a
  // plain hand-written domain rule's mail stopped teaching its case number —
  // the rule that also covers the case's mail from other senders.
  it('repro: a row a hand-written PLAIN domain rule routed still teaches its case code', async () => {
    const mine = broad({ id: 'mine', source: 'user_manual', targetFolderPath: '03/甲' })
    await setRules([mine])
    const row = routedBy(mine, '25A0067A 開庭通知')
    const r = await generateAiConfirmedRules([row], [moved(row)])
    expect(r.added).toBe(1) // was 0
    expect((await listRules()).find((x) => x.type === 'case_code')).toMatchObject({
      signal: '25A0067A',
      targetFolderPath: '03/甲',
    })
  })

  it('…and its court case number (domain + case compound)', async () => {
    const mine = broad({ id: 'mine', source: 'user_manual', targetFolderPath: '03/甲' })
    await setRules([mine])
    const row = routedBy(mine, '112訴204 開庭通知')
    expect((await generateAiConfirmedRules([row], [moved(row)])).added).toBe(1) // was 0
    const learned = (await listRules()).find((x) => x.type === 'compound')!
    expect(decodeCompound(learned.signal)?.conditions).toEqual([
      { type: 'domain', value: 'client.example' },
      { type: 'subject_keyword', value: '112訴204' },
    ])
  })

  it('a hand-written subject or case_code rule still teaches nothing', async () => {
    const subject = { ...broad({ id: 'mine-subject', source: 'user_manual', targetFolderPath: '03/帳務' }), type: 'subject_keyword' as const, signal: '帳單' }
    await setRules([subject])
    const row = routedBy(subject, '112訴204 帳單')
    expect((await generateAiConfirmedRules([row], [moved(row)])).added).toBe(0)
    const code = { ...broad({ id: 'mine-code', source: 'user_manual', targetFolderPath: '03/乙' }), type: 'case_code' as const, signal: '25A0067A' }
    await setRules([code])
    const row2 = routedBy(code, '25A0067A 112訴204 併案通知')
    expect((await generateAiConfirmedRules([row2], [moved(row2)])).added).toBe(0)
  })
})

// R4 (2026-11-06): the revive guard counted EVERY compound on the anchor —
// a case-number compound or the user's own exception included — so a
// high-error-rate / stale domain rule of any client with one could never
// come back. Only AUTO full-subject compounds (the 先廣後窄 trace) hold the
// anchor, as in the learners' narrowedByAnchor.
describe('structural and hand-written compounds do not hold the anchor against a revive (R4)', () => {
  const caseCompound = () =>
    compound('112訴204', '03/112訴204', { id: 'case', source: 'ai_confirmed', confidence: 0.9 })
  const ownException = () => compound('請款通知單', '03/帳務', { id: 'mine', source: 'user_manual' })
  const tired = () => broad({ enabled: false, autoDisabledReason: 'high-error-rate', autoDisabledAt: '2026-11-01T00:00:00.000Z' })

  it('repro: the central guard revives the broad rule next to a case compound and a hand-written exception', async () => {
    await setRules([tired(), caseCompound(), ownException()])
    const res = await addRulesFilteringTombstones([broad({ id: 'cand', enabled: true })])
    expect(res.reEnabled.map((x) => x.id)).toEqual(['broad']) // was []
  })

  it('…and so does the AI-confirmed learner', async () => {
    await setRules([tired(), caseCompound(), ownException()])
    const a = aiFiling(SUBJ_A, A)
    const r = await generateAiConfirmedRules([a], [moved(a)])
    expect(r.revived).toBe(1) // was 0
    expect((await liveDomainRules()).map((x) => x.targetFolderPath)).toEqual([A])
  })

  it('an in-call structural compound candidate does not block it either; an AUTO full-subject one still does', async () => {
    await setRules([tired()])
    const res = await addRulesFilteringTombstones([broad({ id: 'cand', enabled: true }), caseCompound()])
    expect(res.reEnabled.map((x) => x.id)).toEqual(['broad'])
    await setRules([tired()])
    const res2 = await addRulesFilteringTombstones([broad({ id: 'cand', enabled: true }), compound(SUBJ_B1, B)])
    expect(res2.reEnabled).toHaveLength(0)
  })
})

// Round 3: an ENABLED but ORPHANED rule (its folder is gone) routes nothing,
// yet the same-signal conflict checks counted it — so a rule an import had
// switched off for it could never come back, and nothing routed the signal.
describe('an orphaned rule is no conflict', () => {
  const orphanedMine = (over: Partial<Rule> = {}) =>
    broad({ id: 'mine', source: 'user_manual', targetFolderPath: '03/已刪除', enabled: true, orphaned: true, ...over })

  it('repro: the central guard revives the rule an orphaned import left switched off', async () => {
    const auto = broad({ id: 'auto', type: 'case_code', signal: '25A0067A', enabled: false, autoDisabledReason: 'superseded', autoDisabledAt: '2026-11-01T00:00:00.000Z' })
    await setRules([auto, orphanedMine({ type: 'case_code', signal: '25A0067A' })])
    const a = aiFiling('25A0067A 開庭通知', A)
    const r = await generateAiConfirmedRules([a], [moved(a)])
    expect(r.revived).toBe(1) // was 0: the orphaned twin counted as a live conflict
    expect((await listRules()).find((x) => x.id === 'auto')).toMatchObject({ enabled: true })
  })

  it('the AI-confirmed learner learns the broad rule next to an orphaned hand-written one', async () => {
    await setRules([orphanedMine()])
    const a = aiFiling(SUBJ_A, A)
    expect((await generateAiConfirmedRules([a], [moved(a)])).added).toBe(1) // was 0
    expect((await liveDomainRules()).filter((x) => !x.orphaned).map((x) => x.targetFolderPath)).toEqual([A])
  })
})
