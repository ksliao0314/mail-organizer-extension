// Own-mailbox domain hijack (2026-11-05, item 5).
//
// With settings.internalDomains = [] (the default), the first AI filing of a
// colleague's mail learned `domain: <own domain>`. Domain rules match ANY
// address incl. To/Cc, and the mailbox owner is on To/Cc of every inbound
// email — so the rule claimed nearly all mail, and its empirical confidence
// rose above the clients' rules. Repro: a court notice To me@firm.example →
// 所內/行政.

import { describe, expect, it } from 'vitest'
import {
  generateAiConfirmedRules,
  generateAiOverrideRules,
  type ExecuteItemResult,
} from '@/background/execute'
import { startInitialScan } from '@/background/initial-scan'
import { buildRuleIndex, listRules, matchEmailWithIndex, newRule } from '@/shared/rules'
import { getMailboxIdentity, setMailboxIdentity, setRules } from '@/shared/storage'
import type { OutlookApi } from '@/shared/outlook-api'
import type { Email, MailFolderNode, PlanItem, Rule } from '@/shared/types'

const FID = 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3'
const ME = 'me@firm.example'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    ...newRule({
      type: 'domain',
      signal: 'firm.example',
      targetFolderId: FID,
      targetFolderPath: '所內/行政',
      confidence: 0.7,
      source: 'ai_confirmed',
    }),
    createdAt: '2026-01-01T00:00:00.000Z',
    ...over,
  }
}

function email(from: string, over: Partial<Email> = {}): Email {
  return {
    Id: 'e-' + Math.random().toString(36).slice(2),
    Subject: '開庭通知',
    BodyPreview: '',
    From: { EmailAddress: { Address: from } },
    ToRecipients: [{ EmailAddress: { Address: ME } }],
    ReceivedDateTime: '2026-11-01T00:00:00Z',
    ParentFolderId: 'inbox',
    ...over,
  } as Email
}

describe('mailbox identity storage', () => {
  it('round-trips, normalizes, and derives the domain from the address', async () => {
    expect(await getMailboxIdentity()).toBeNull()
    await setMailboxIdentity({ address: ' Me@Firm.Example ', domain: '', at: '2026-11-05T00:00:00Z' })
    expect(await getMailboxIdentity()).toEqual({
      address: 'me@firm.example',
      domain: 'firm.example',
      at: '2026-11-05T00:00:00Z',
    })
  })

  it('ignores an address without a usable domain', async () => {
    await setMailboxIdentity({ address: 'not-an-address', domain: '', at: '' })
    expect(await getMailboxIdentity()).toBeNull()
  })
})

describe('buildRuleIndex ownDomains', () => {
  it('repro: an auto own-domain rule no longer claims a court notice addressed to me', () => {
    const own = rule()
    const notice = email('clerk@court.example.tw')
    // Without the guard the rule fires through the To address.
    expect(matchEmailWithIndex(notice, buildRuleIndex([own]))?.rule.id).toBe(own.id)
    expect(matchEmailWithIndex(notice, buildRuleIndex([own], { ownDomains: ['firm.example'] }))).toBeNull()
  })

  it('client domain rules are unaffected by the owner being on To', () => {
    const client = rule({ signal: 'client.example', targetFolderPath: '03/Client' })
    const own = rule()
    const out = matchEmailWithIndex(
      email('pm@client.example'),
      buildRuleIndex([own, client], { ownDomains: ['firm.example'] }),
    )
    expect(out?.rule.targetFolderPath).toBe('03/Client')
  })

  it('a user_manual own-domain rule is honoured on the From side only', () => {
    const manual = rule({ source: 'user_manual', targetFolderPath: '所內/同事' })
    const index = buildRuleIndex([manual], { ownDomains: ['FIRM.example'] })
    const colleague = matchEmailWithIndex(email('colleague@firm.example'), index)
    expect(colleague?.rule.id).toBe(manual.id)
    expect(colleague?.reason).toBe('寄件人網域 @firm.example')
    expect(matchEmailWithIndex(email('pm@client.example'), index)).toBeNull()
  })
})

const moved = (emailId: string): ExecuteItemResult => ({
  emailId,
  subject: 'X',
  action: 'move',
  status: 'moved',
})

function colleagueItem(over: Partial<PlanItem> = {}): PlanItem {
  return {
    emailId: 'e-' + Math.random().toString(36).slice(2),
    emailSubject: '本週行政事項公告與提醒',
    emailFrom: 'colleague@firm.example',
    action: 'move',
    targetFolderPath: '所內/行政',
    confidence: 0.9,
    reason: '',
    source: 'unresolved', // no birth gate — isolates the domain decision
    ...over,
  }
}

describe('learners never mint an own-domain domain rule', () => {
  it('AI-confirmed path treats the own domain like an internal one (sender + subject instead)', async () => {
    await setMailboxIdentity({ address: ME, domain: 'firm.example', at: '' })
    const item = colleagueItem()
    const r = await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect(r.added).toBe(1)
    const rules = await listRules()
    expect(rules.filter((x) => x.type === 'domain')).toHaveLength(0)
    expect(rules[0]!.type).toBe('compound')
  })

  it('without the identity the same filing still learns the hijacking domain rule (baseline)', async () => {
    const item = colleagueItem()
    await generateAiConfirmedRules([item], [moved(item.emailId)])
    expect((await listRules()).map((x) => `${x.type}:${x.signal}`)).toEqual(['domain:firm.example'])
  })

  it('override path: no own-domain conflict upgrade either', async () => {
    await setMailboxIdentity({ address: ME, domain: 'firm.example', at: '' })
    await setRules([rule({ targetFolderPath: '所內/其他' })])
    const item = colleagueItem({
      source: 'ai',
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: '所內/其他',
    })
    await generateAiOverrideRules([item], [moved(item.emailId)])
    const created = (await listRules()).filter((x) => x.source === 'ai_overridden')
    expect(created).toHaveLength(1)
    expect(created[0]!.signal).toContain('colleague@firm.example') // sender-anchored, not domain
  })

  it('initial scan skips the own domain', async () => {
    await setMailboxIdentity({ address: ME, domain: 'firm.example', at: '' })
    const tree: MailFolderNode[] = [
      {
        id: 'root-id-xxxxxxxxxxxxxxxxxxxx',
        displayName: '案件',
        path: '案件',
        children: [
          { id: 'folder-a-xxxxxxxxxxxxxxxxxxx', displayName: 'A', path: '案件/A', children: [] },
        ],
      },
    ]
    const mails: Partial<Email>[] = [
      { From: { EmailAddress: { Address: 'colleague@firm.example' } }, Subject: '一' },
      { From: { EmailAddress: { Address: 'colleague@firm.example' } }, Subject: '二' },
      { From: { EmailAddress: { Address: 'pm@client.example' } }, Subject: '三' },
      { From: { EmailAddress: { Address: 'pm@client.example' } }, Subject: '四' },
    ]
    const api = {
      listFolderMessages: async (id: string) =>
        id === 'folder-a-xxxxxxxxxxxxxxxxxxx'
          ? mails.map((m, i) => ({ Id: `m${i}`, ToRecipients: [], ...m }))
          : [],
    } as unknown as OutlookApi
    await startInitialScan({ rootPath: '案件', tree, api, excludePrefixes: [], internalDomains: [] })
    const domains = (await listRules()).filter((r) => r.type === 'domain').map((r) => r.signal)
    expect(domains).toEqual(['client.example'])
  })
})
