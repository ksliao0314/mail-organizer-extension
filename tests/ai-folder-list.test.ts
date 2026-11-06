// The AI's folder list (O1c, 2026-11-06 r2). Round 1 put Junk Email / Sent
// Items (and the trash root) back into the folder tree — and with it into the
// AI's folder block: the AI proposed `move → 垃圾郵件` / `→ 刪除的郵件`,
// those moves were learned as auto-junk / auto-trash rules, and new_folder
// under 刪除的郵件 always failed in execute. The trash / Drafts / Outbox are
// no longer in the tree at all; Junk / Sent stay valid targets for rules, the
// picker and execute, but the AI is only offered ordinary folders, and
// whatever it names outside that list becomes an unresolved row.

import { describe, expect, it } from 'vitest'
import { actionToPlanItem, buildActiveFoldersBlock, buildFolderBlock } from '@/shared/classifier'
import type { Email, MailFolderNode } from '@/shared/types'

const tree = (): MailFolderNode[] => [
  {
    id: 'CASES-0000000000000000000',
    displayName: '案件',
    path: '案件',
    children: [{ id: 'JIA-00000000000000000000', displayName: '甲', path: '案件/甲', children: [] }],
  },
  {
    id: 'JUNK-0000000000000000000',
    displayName: '垃圾郵件',
    path: '垃圾郵件',
    wellKnown: 'junkemail',
    children: [{ id: 'JUNK-SUB-000000000000000', displayName: '可疑', path: '垃圾郵件/可疑', children: [] }],
  },
  {
    id: 'SENT-0000000000000000000',
    displayName: '寄件備份',
    path: '寄件備份',
    wellKnown: 'sentitems',
    children: [{ id: 'SENT-A-00000000000000000', displayName: '客戶A', path: '寄件備份/客戶A', children: [] }],
  },
]

const email = { Id: 'e1', Subject: '廣告', BodyPreview: '', From: { EmailAddress: { Address: 'promo@shop.example' } } } as Email

describe('folder block offers only ordinary folders', () => {
  it('leaves the Junk / Sent subtrees out', () => {
    const block = buildFolderBlock(tree(), [])
    expect(block).toContain('- 案件/甲')
    expect(block).not.toContain('垃圾郵件')
    expect(block).not.toContain('寄件備份')
  })

  it('the recent-folders hint leaves them out too (the ledger still records filings there)', () => {
    const block = buildActiveFoldersBlock(['垃圾郵件', '案件/甲', '寄件備份/客戶A', '垃圾郵件/舊的'], [], tree())
    expect(block).toContain('- 案件/甲')
    expect(block).not.toContain('垃圾郵件')
    expect(block).not.toContain('寄件備份')
  })
})

describe('AI picks outside the offered list become unresolved', () => {
  it('a move into Junk is a 保留 row — never a move that would execute and be learned', () => {
    const p = actionToPlanItem(email, { emailIndex: 0, action: 'move', targetFolderPath: '垃圾郵件', confidence: 0.95, reason: '廣告' }, tree())
    expect(p).toMatchObject({ action: 'skip', source: 'unresolved', aiOriginalAction: 'move', aiOriginalTargetFolderPath: '垃圾郵件' })
    expect(p.targetFolderId).toBeUndefined()
    expect(p.reason).toContain('不在可用資料夾清單')
  })

  it('…and into a Sent Items subfolder', () => {
    const p = actionToPlanItem(email, { emailIndex: 0, action: 'move', targetFolderPath: '寄件備份/客戶A', confidence: 0.9, reason: '' }, tree())
    expect(p).toMatchObject({ action: 'skip', source: 'unresolved' })
  })

  it('an ordinary folder still resolves', () => {
    const p = actionToPlanItem(email, { emailIndex: 0, action: 'move', targetFolderPath: '案件/甲', confidence: 0.9, reason: '' }, tree())
    expect(p).toMatchObject({ action: 'move', source: 'ai', targetFolderId: 'JIA-00000000000000000000' })
  })

  it('new_folder under Junk / Sent is unresolved', () => {
    const p = actionToPlanItem(
      email,
      { emailIndex: 0, action: 'new_folder', suggestedParentPath: '寄件備份', suggestedFolderName: '客戶B', confidence: 0.9, reason: '' },
      tree(),
    )
    expect(p).toMatchObject({ action: 'skip', source: 'unresolved', aiOriginalAction: 'new_folder' })
    expect(p.reason).toContain('不在可用資料夾清單')
  })

  it('a "new" top-level folder that IS the Junk folder is unresolved (execute would just reuse it)', () => {
    const p = actionToPlanItem(
      email,
      { emailIndex: 0, action: 'new_folder', suggestedFolderName: '垃圾郵件', confidence: 0.9, reason: '' },
      tree(),
    )
    expect(p).toMatchObject({ action: 'skip', source: 'unresolved' })
  })

  it('a genuinely new folder under an ordinary parent is untouched', () => {
    const p = actionToPlanItem(
      email,
      { emailIndex: 0, action: 'new_folder', suggestedParentPath: '案件', suggestedFolderName: '乙', confidence: 0.9, reason: '' },
      tree(),
    )
    expect(p).toMatchObject({ action: 'new_folder', source: 'ai', suggestedParentPath: '案件', suggestedFolderName: '乙' })
  })
})

// Integration follow-ups (2026-11-06 r2).
describe('nothing outside the offered list reaches the prompt or executes', () => {
  it('an AI pick under an EXCLUDED prefix is a 保留 row — an untouched row no longer files into it', () => {
    const p = actionToPlanItem(
      email,
      { emailIndex: 0, action: 'move', targetFolderPath: '案件/甲', confidence: 0.9, reason: '' },
      tree(),
      ['案件'],
    )
    expect(p).toMatchObject({ action: 'skip', source: 'unresolved', aiOriginalAction: 'move', aiOriginalTargetFolderPath: '案件/甲' })
    expect(p.targetFolderId).toBeUndefined()
    expect(p.targetFolderPath).toBeUndefined()
    expect(p.reason).toContain('排除清單')
  })

  it('the recent-folders hint drops paths no longer in the tree (刪除的郵件 recorded by older builds)', () => {
    const block = buildActiveFoldersBlock(['刪除的郵件', '案件/甲', '案件/已刪除的資料夾'], [], tree())
    expect(block).toContain('- 案件/甲')
    expect(block).not.toContain('刪除的郵件')
    expect(block).not.toContain('已刪除的資料夾')
  })
})

// r3 (2026-11-06): the tree no longer lists Deleted Items / Drafts / Outbox,
// so a top-level "new" 「刪除的郵件」 looked valid in the plan — and failed at
// execute (409 against the trash itself, refused). Their names come in from
// the folder cache's well-known roots.
describe('a top-level new folder named like an omitted well-known folder (r3)', () => {
  const newTop = (name: string) =>
    actionToPlanItem(
      email,
      { emailIndex: 0, action: 'new_folder', suggestedFolderName: name, confidence: 0.9, reason: '' },
      tree(),
      [],
      ['刪除的郵件', '草稿', '寄件匣'],
    )

  it('is unresolved (case-insensitively, by display name)', () => {
    expect(newTop('刪除的郵件')).toMatchObject({ action: 'skip', source: 'unresolved', aiOriginalAction: 'new_folder' })
    expect(newTop(' 草稿 ')).toMatchObject({ action: 'skip', source: 'unresolved' })
  })

  it('…only at the top level; an ordinary new name is untouched', () => {
    const nested = actionToPlanItem(
      email,
      { emailIndex: 0, action: 'new_folder', suggestedParentPath: '案件', suggestedFolderName: '草稿', confidence: 0.9, reason: '' },
      tree(),
      [],
      ['草稿'],
    )
    expect(nested).toMatchObject({ action: 'new_folder', source: 'ai' })
    expect(newTop('客戶B')).toMatchObject({ action: 'new_folder', source: 'ai' })
  })
})
