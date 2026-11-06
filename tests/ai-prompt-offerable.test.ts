// Exemplars and the case map only name folders the AI is offered (2026-11-06
// r2): a rule into 垃圾郵件 / 寄件備份 or a folder no longer in the tree used
// to appear as "寄件人網域 @x → 垃圾郵件" and steer the model toward a target
// it can't use (the pick becomes a 保留 row).

import { afterEach, describe, expect, it, vi } from 'vitest'
import { classifyBatch } from '@/shared/classifier'
import { DEFAULT_SETTINGS, type Email, type MailFolderNode, type Rule, type Settings } from '@/shared/types'

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
    children: [],
  },
]

const rule = (over: Partial<Rule>): Rule => ({
  id: 'r-' + Math.random().toString(36).slice(2),
  type: 'domain',
  signal: 'client.example',
  targetFolderId: 'JIA-00000000000000000000',
  targetFolderPath: '案件/甲',
  confidence: 0.9,
  matchCount: 30,
  enabled: true,
  createdAt: '2026-01-01T00:00:00.000Z',
  source: 'user_manual',
  ...over,
})

async function promptBody(rules: Rule[]): Promise<string> {
  let body = ''
  vi.stubGlobal(
    'fetch',
    vi.fn(async (_url: string, init: RequestInit) => {
      body = typeof init.body === 'string' ? init.body : ''
      return new Response(
        JSON.stringify({
          content: [{ type: 'text', text: '{"actions":[{"emailIndex":0,"action":"skip","confidence":0.5,"reason":"x"}]}' }],
          usage: { input_tokens: 1, output_tokens: 1, cache_creation_input_tokens: 0, cache_read_input_tokens: 0 },
          stop_reason: 'end_turn',
        }),
        { status: 200 },
      )
    }),
  )
  const settings: Settings = { ...DEFAULT_SETTINGS, claudeApiKey: 'sk-ant-test', aiIncludeFewShotExamples: true }
  await classifyBatch(
    {
      emails: [{ Id: 'e1', Subject: '112年度訴字第204號', From: { EmailAddress: { Address: 'a@b.example' } } } as Email],
      folderTree: tree(),
      excludePrefixes: [],
      rules,
    },
    settings,
  )
  return body
}

describe('exemplars and the case map', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('keep rules into offered folders', async () => {
    const body = await promptBody([rule({}), rule({ type: 'subject_keyword', signal: '112訴204' })])
    expect(body).toContain('client.example')
    expect(body).toContain('112訴204')
  })

  it('drop rules into Junk / Sent or folders no longer in the tree', async () => {
    const body = await promptBody([
      rule({ signal: 'spam-sender.example', targetFolderId: 'JUNK-0000000000000000000', targetFolderPath: '垃圾郵件' }),
      rule({ signal: 'gone-client.example', targetFolderId: 'GONE-0000000000000000000', targetFolderPath: '案件/已刪除' }),
      rule({ type: 'subject_keyword', signal: '112訴204', targetFolderPath: '垃圾郵件', targetFolderId: 'JUNK-0000000000000000000' }),
    ])
    expect(body).not.toContain('spam-sender.example')
    expect(body).not.toContain('gone-client.example')
    expect(body).not.toContain('112訴204 →')
  })
})
