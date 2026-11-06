// Match-path tiers (2026-11-05).
//
// Item 4 — hand-written sender / subject rules lost to auto domain rules:
//   the walk went strictly by type, so ANY auto domain rule matching ANY party
//   (To/Cc included) beat a user_manual sender or subject_keyword rule — and
//   thread memory then beat it too, because ruleBeatsThread saw the domain
//   rule. user_manual rules now form their own tier ahead of every auto broad
//   rule (structural tiers stay on top).
// Item 8 — compound specificity: the compound tier sorted by confidence /
//   recency only, so 「報告」→甲 claimed 「RE: 乙案進度報告」 when it fired
//   more recently; now longest subject condition first.
// Item 10 — one malformed compound must not break matching for every email.
// Item 16 — compounds decoded once and bucketed by anchor; behaviour pinned.

import { describe, expect, it } from 'vitest'
import {
  buildRuleIndex,
  encodeCompound,
  matchBodyCaseWithIndex,
  matchEmailWithIndex,
  newRule,
  ruleBeatsThread,
} from '@/shared/rules'
import type { Email, Rule } from '@/shared/types'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'client.example',
      targetFolderId: FID,
      targetFolderPath: 'auto/domain',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  }
}

function email(over: Partial<Email> = {}): Email {
  return {
    Id: 'e1',
    Subject: '',
    BodyPreview: '',
    From: { EmailAddress: { Address: 'someone@client.example' } },
    ToRecipients: [{ EmailAddress: { Address: 'me@firm.example' } }],
    ReceivedDateTime: '2026-11-01T00:00:00Z',
    ParentFolderId: 'inbox',
    ...over,
  } as Email
}

const match = (rules: Rule[], e: Email, ownDomains?: string[]) =>
  matchEmailWithIndex(e, buildRuleIndex(rules, { ownDomains }))

describe('user_manual tier beats auto broad rules (item 4)', () => {
  it('a hand-written sender rule beats an auto domain rule matched via a Cc party', () => {
    const auto = rule({ signal: 'cc-party.example' })
    const manual = rule({
      type: 'sender',
      signal: 'boss@gmail.com',
      source: 'user_manual',
      targetFolderPath: 'manual/boss',
    })
    const e = email({
      From: { EmailAddress: { Address: 'boss@gmail.com' } },
      CcRecipients: [{ EmailAddress: { Address: 'x@cc-party.example' } }],
    })
    const out = match([auto, manual], e)
    expect(out?.rule.targetFolderPath).toBe('manual/boss') // was auto/domain
    // …and the preflight precedence now sees the hand-written rule:
    expect(ruleBeatsThread(out!.rule)).toBe(true)
  })

  it('a hand-written subject rule beats the auto domain rule of the sender', () => {
    const auto = rule()
    const manual = rule({
      type: 'subject_keyword',
      signal: '工時審閱',
      source: 'user_manual',
      targetFolderPath: 'manual/timesheet',
    })
    const out = match([auto, manual], email({ Subject: 'RE: 十月工時審閱' }))
    expect(out?.rule.targetFolderPath).toBe('manual/timesheet')
  })

  it('a hand-written domain rule beats an auto domain rule with higher empirical confidence', () => {
    const auto = rule({ signal: 'cc-party.example', matchCount: 50, confidence: 0.7 })
    const manual = rule({ signal: 'client.example', source: 'user_manual', targetFolderPath: 'manual/client', confidence: 0.6 })
    const e = email({ CcRecipients: [{ EmailAddress: { Address: 'x@cc-party.example' } }] })
    expect(match([auto, manual], e)?.rule.targetFolderPath).toBe('manual/client')
  })

  it('structural tiers still come first: compound and court-case rules outrank user_manual subject rules', () => {
    const manual = rule({
      type: 'subject_keyword',
      signal: '開庭通知',
      source: 'user_manual',
      targetFolderPath: 'manual/hearings',
    })
    const court = rule({ type: 'subject_keyword', signal: '112訴204', targetFolderPath: 'case/204' })
    const compound = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: '113訴9' },
      ]),
      targetFolderPath: 'case/9',
    })
    expect(match([manual, court], email({ Subject: '112訴204 開庭通知' }))?.rule.targetFolderPath).toBe('case/204')
    expect(match([manual, compound], email({ Subject: '113訴9 開庭通知' }))?.rule.targetFolderPath).toBe('case/9')
  })

  it('without a hand-written match the auto order is unchanged (domain before subject before sender)', () => {
    const dom = rule()
    const subj = rule({ type: 'subject_keyword', signal: '請款通知單', targetFolderPath: 'auto/subject' })
    const snd = rule({ type: 'sender', signal: 'someone@client.example', targetFolderPath: 'auto/sender' })
    expect(match([snd, subj, dom], email({ Subject: '請款通知單' }))?.rule.targetFolderPath).toBe('auto/domain')
    expect(match([snd, subj], email({ Subject: '請款通知單' }))?.rule.targetFolderPath).toBe('auto/subject')
  })
})

describe('compound specificity (item 8)', () => {
  const compound = (subject: string, target: string, over: Partial<Rule> = {}) =>
    rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: subject },
      ]),
      targetFolderPath: target,
      ...over,
    })

  it('the longest subject condition wins even when the shorter rule fired more recently', () => {
    const broad = compound('報告', 'case/甲', { lastUsedAt: '2026-11-04T00:00:00.000Z', confidence: 0.95 })
    const specific = compound('乙案進度報告', 'case/乙', { lastUsedAt: '2026-06-01T00:00:00.000Z' })
    const out = match([broad, specific], email({ Subject: 'RE: 乙案進度報告' }))
    expect(out?.rule.targetFolderPath).toBe('case/乙') // was case/甲
    // The broad one still catches what only it matches.
    expect(match([broad, specific], email({ Subject: '甲案報告' }))?.rule.targetFolderPath).toBe('case/甲')
  })

  // B4 (2026-11-06): "longest subject first" ignored source and structure.
  it('a hand-written compound beats a longer auto one (vendor + 請款 vs vendor + 三月份請款單…)', () => {
    const manual = compound('請款', 'manual/billing', { source: 'user_manual' })
    const auto = compound('三月份請款單與發票明細', 'auto/march')
    const out = match([auto, manual], email({ Subject: 'RE: 三月份請款單與發票明細' }))
    expect(out?.rule.targetFolderPath).toBe('manual/billing') // was auto/march
  })

  it('a court-case compound beats a full-subject compound that merely contains the case number', () => {
    const court = compound('112訴204', 'case/204')
    const fullSubject = compound('112訴204 股權移轉案進度報告', 'case/progress', { source: 'ai_overridden' })
    const out = match([fullSubject, court], email({ Subject: '112訴204 股權移轉案進度報告' }))
    expect(out?.rule.targetFolderPath).toBe('case/204') // was case/progress
    // …and so does a Latin case-code compound.
    const code = compound('25A0067A', 'case/25A0067A')
    const full2 = compound('25A0067A 補件通知與期日', 'case/notice')
    expect(match([full2, code], email({ Subject: '25A0067A 補件通知與期日' }))?.rule.targetFolderPath).toBe(
      'case/25A0067A',
    )
  })

  it('a hand-written compound with no subject condition beats an auto one that has one', () => {
    const manual = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'sender', value: 'someone@client.example' },
        { type: 'domain', value: 'client.example' },
      ]),
      source: 'user_manual',
      targetFolderPath: 'manual/someone',
    })
    const auto = compound('季度營運報告', 'auto/quarterly')
    expect(match([auto, manual], email({ Subject: '季度營運報告' }))?.rule.targetFolderPath).toBe('manual/someone')
  })

  it('structural still outranks hand-written inside the compound tier', () => {
    const manual = compound('請款', 'manual/billing', { source: 'user_manual' })
    const court = compound('112訴204', 'case/204')
    expect(match([manual, court], email({ Subject: '112訴204 律師費請款' }))?.rule.targetFolderPath).toBe('case/204')
  })

  // R1 (2026-11-06): B4 ranked every auto case-number compound above a
  // hand-written compound whose subject merely CONTAINS that case number —
  // the user's exception inside the case. v1.0.8 routed 「112訴204 帳單
  // 三月份」 to 帳務; B4 sent it to the case folder, and correcting that
  // one-strike-disabled the case rule and learned (domain + 112訴204) → 帳務.
  it('a hand-written compound naming a case number beats the auto case compound (the exception inside the case)', () => {
    const manual = compound('112訴204 帳單', 'manual/帳務', { source: 'user_manual' })
    const auto = compound('112訴204', 'case/甲')
    const rules = [auto, manual]
    expect(match(rules, email({ Subject: '112訴204 帳單 三月份' }))?.rule.targetFolderPath).toBe('manual/帳務') // was case/甲
    // The rest of the case still goes to the case folder.
    expect(match(rules, email({ Subject: '112訴204 開庭通知' }))?.rule.targetFolderPath).toBe('case/甲')
  })

  it('the same for a Latin case code inside the hand-written subject; an AUTO one that contains it still loses', () => {
    const manual = compound('25A0067A 帳單', 'manual/帳務', { source: 'user_manual' })
    const auto = compound('25A0067A', 'case/25A0067A')
    expect(match([auto, manual], email({ Subject: '25A0067A 帳單' }))?.rule.targetFolderPath).toBe('manual/帳務')
    const autoFull = compound('25A0067A 帳單', 'auto/帳務', { source: 'ai_overridden' })
    expect(match([autoFull, auto], email({ Subject: '25A0067A 帳單' }))?.rule.targetFolderPath).toBe('case/25A0067A')
  })
})

// Tier 0 (2026-11-06): R1 across tiers. A hand-written rule whose subject
// names a case id still lost to the auto case_code (tier 1) and to the auto
// case-number subject rules (tier 3) / compounds (tier 2); correcting the
// misfile one-strike-disabled the case rule and taught the whole case to go
// to the exception's folder.
describe('tier 0: hand-written rules naming a case id beat every auto case rule', () => {
  const caseCode = (code: string, target: string, over: Partial<Rule> = {}) =>
    rule({ type: 'case_code', signal: code, targetFolderPath: target, confidence: 0.9, ...over })
  const compound = (subject: string, target: string, over: Partial<Rule> = {}) =>
    rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: 'client.example' },
        { type: 'subject_keyword', value: subject },
      ]),
      targetFolderPath: target,
      ...over,
    })

  it('repro: a hand-written compound naming a Latin case code beats the auto case_code rule', () => {
    const rules = [
      caseCode('25A0067A', 'case/25A0067A'),
      compound('25A0067A 帳單', 'manual/帳務', { source: 'user_manual' }),
    ]
    expect(match(rules, email({ Subject: '25A0067A 帳單 三月份' }))?.rule.targetFolderPath).toBe('manual/帳務') // was case/25A0067A
    expect(match(rules, email({ Subject: '25A0067A 補件通知' }))?.rule.targetFolderPath).toBe('case/25A0067A')
  })

  it('repro: a hand-written subject rule naming a court case number beats the auto case-number rules', () => {
    const manual = rule({ type: 'subject_keyword', signal: '112訴204 帳單', source: 'user_manual', targetFolderPath: 'manual/帳務' })
    const autoCompound = compound('112訴204', 'case/甲', { confidence: 0.9 })
    const autoSubject = rule({ type: 'subject_keyword', signal: '112訴204', targetFolderPath: 'case/甲2', confidence: 0.85 })
    expect(match([autoCompound, manual], email({ Subject: '112訴204 帳單 三月份' }))?.rule.targetFolderPath).toBe('manual/帳務') // was case/甲
    expect(match([autoSubject, manual], email({ Subject: '112訴204 帳單 三月份' }))?.rule.targetFolderPath).toBe('manual/帳務') // was case/甲2
    expect(match([autoCompound, autoSubject, manual], email({ Subject: '112訴204 開庭通知' }))?.rule.targetFolderPath).toBe('case/甲')
  })

  it('a hand-written case_code rule keeps its place (tier 1): ahead of auto compounds, behind the tier-0 exception', () => {
    const manualCode = caseCode('25A0067A', 'manual/code', { source: 'user_manual' })
    const autoCompound = compound('25A0067A', 'auto/compound')
    expect(match([autoCompound, manualCode], email({ Subject: '25A0067A 進度' }))?.rule.targetFolderPath).toBe('manual/code')
    const exception = rule({ type: 'subject_keyword', signal: '25A0067A 帳單', source: 'user_manual', targetFolderPath: 'manual/帳務' })
    expect(match([manualCode, exception], email({ Subject: '25A0067A 帳單' }))?.rule.targetFolderPath).toBe('manual/帳務')
  })

  it('a hand-written pure case-number subject rule still serves the body case pass', () => {
    const manual = rule({ type: 'subject_keyword', signal: '112訴204', source: 'user_manual', targetFolderPath: 'manual/case' })
    const index = buildRuleIndex([manual])
    expect(matchEmailWithIndex(email({ Subject: '112訴204 通知' }), index)?.rule.targetFolderPath).toBe('manual/case')
    const bodyOnly = email({ Subject: '開庭通知', bodyText: '本院 112年度訴字第204號 定於下週開庭' })
    expect(matchBodyCaseWithIndex(bodyOnly, index)?.rule.targetFolderPath).toBe('manual/case')
  })
})

describe('compiled compound index (item 16)', () => {
  it('sender-, domain- and subject-only compounds all match through their buckets', () => {
    const bySender = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'sender', value: 'Boss@Client.example' },
        { type: 'subject_keyword', value: '董事會議程確認' },
      ]),
      targetFolderPath: 'c/sender',
    })
    const byDomain = rule({
      type: 'compound',
      signal: encodeCompound([
        { type: 'domain', value: '@Partner.example' },
        { type: 'subject_keyword', value: '合作備忘錄草案' },
      ]),
      targetFolderPath: 'c/domain',
    })
    const subjectOnly = rule({
      type: 'compound',
      signal: encodeCompound([{ type: 'subject_keyword', value: '年度稽核排程' }]),
      targetFolderPath: 'c/subject',
    })
    const rules = [bySender, byDomain, subjectOnly]
    const a = match(rules, email({ From: { EmailAddress: { Address: 'boss@client.example' } }, Subject: '董事會議程確認' }))
    expect(a?.rule.targetFolderPath).toBe('c/sender')
    expect(a?.reason).toBe('寄件人 Boss@Client.example + 主旨含「董事會議程確認」')
    const b = match(
      rules,
      email({ ToRecipients: [{ EmailAddress: { Address: 'x@partner.example' } }], Subject: '合作備忘錄草案 v2' }),
    )
    expect(b?.rule.targetFolderPath).toBe('c/domain')
    expect(b?.reason).toBe('收件人網域 @partner.example + 主旨含「合作備忘錄草案」')
    expect(match(rules, email({ Subject: '年度稽核排程' }))?.rule.targetFolderPath).toBe('c/subject')
    // Anchor present but the AND fails → no match.
    expect(match(rules, email({ From: { EmailAddress: { Address: 'boss@client.example' } }, Subject: '無關' }))).toBeNull()
  })
})

describe('malformed rules never throw per email (item 10)', () => {
  it('a compound with a non-string value is skipped; the rest of the library still matches', () => {
    const broken = rule({
      type: 'compound',
      signal: '{"conditions":[{"type":"domain","value":123}]}',
      targetFolderPath: 'broken',
    })
    const notJson = rule({ type: 'compound', signal: '{not json', targetFolderPath: 'broken2' })
    const ok = rule()
    const index = buildRuleIndex([broken, notJson, ok])
    expect(() => matchEmailWithIndex(email(), index)).not.toThrow()
    expect(matchEmailWithIndex(email(), index)?.rule.targetFolderPath).toBe('auto/domain')
  })

  it('rows with an unknown type or non-string signal are ignored', () => {
    const corrupt = [
      { ...rule(), type: 'bogus' as Rule['type'] },
      { ...rule(), signal: 42 as unknown as string },
    ]
    expect(() => buildRuleIndex(corrupt)).not.toThrow()
    expect(matchEmailWithIndex(email(), buildRuleIndex(corrupt))).toBeNull()
  })
})
