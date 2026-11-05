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
