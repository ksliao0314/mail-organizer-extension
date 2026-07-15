import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  actionToPlanItem,
  buildActiveFoldersBlock,
  buildCaseMapBlock,
  buildEmailBlock,
  buildFolderBlock,
  classifyBatch,
  parseAiActions,
  salvageTruncatedArray,
} from '@/shared/classifier'
import { DEFAULT_SETTINGS } from '@/shared/types'
import type { Email, MailFolderNode, Rule, Settings } from '@/shared/types'

function tree(): MailFolderNode[] {
  return [
    {
      id: 'f1',
      displayName: '內部資料',
      path: '內部資料',
      children: [
        { id: 'f1a', displayName: '工時審閱', path: '內部資料/工時審閱', children: [] },
      ],
    },
    {
      id: 'f2',
      displayName: '05已完成案件',
      path: '05已完成案件',
      children: [
        { id: 'f2a', displayName: 'old', path: '05已完成案件/old', children: [] },
      ],
    },
  ]
}

// ---- parseAiActions --------------------------------------------------------

describe('parseAiActions', () => {
  it('parses clean JSON array', () => {
    const txt = '[{"emailIndex":0,"action":"move","targetFolderPath":"A","confidence":0.9,"reason":"x"}]'
    expect(parseAiActions(txt)).toHaveLength(1)
  })

  it('handles markdown-fence wrapped output', () => {
    const txt = '```json\n[{"emailIndex":0,"action":"delete","confidence":1,"reason":"x"}]\n```'
    const r = parseAiActions(txt)
    expect(r[0]!.action).toBe('delete')
  })

  it('extracts array from prose-wrapped output', () => {
    const txt = '好的，這是結果：[{"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x"}] 完成'
    expect(parseAiActions(txt)).toHaveLength(1)
  })

  it('handles brackets inside strings (balanced parser)', () => {
    const txt = '[{"emailIndex":0,"action":"move","targetFolderPath":"A","confidence":0.9,"reason":"foo[bar]baz"}]'
    expect(parseAiActions(txt)[0]!.reason).toBe('foo[bar]baz')
  })

  it('throws ClassifierError on completely unparseable input', () => {
    expect(() => parseAiActions('completely not JSON at all')).toThrow()
  })

  // ---- Loose JSON recovery (added 2026-05-22) ----------------------------
  //
  // Pre-2026-05-22, strict / balanced parse fails would throw. Now there's
  // a final pass that strips trailing commas, // line comments, /* */ block
  // comments, and smart quotes — covers the most common LLM mishaps even
  // when the prompt forbids them.
  describe('loose JSON recovery', () => {
    it('repairs trailing comma before closing bracket', () => {
      const txt =
        '[{"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x"},]'
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })

    it('repairs trailing comma before closing brace inside object', () => {
      const txt =
        '[{"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x",}]'
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })

    it('strips // line comments', () => {
      const txt = `[
        // this is a comment
        {"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x"}
      ]`
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })

    it('strips /* */ block comments', () => {
      const txt =
        '[/* note */ {"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x"}]'
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })

    it('converts smart quotes to straight quotes', () => {
      // Smart double quotes around keys/values. Note: the JSON itself is
      // valid otherwise (just wrong quote char).
      const txt =
        '[{“emailIndex”:0,“action”:“skip”,“confidence”:0.3,“reason”:“x”}]'
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })

    it('combines repairs (trailing comma + comment)', () => {
      const txt = `[
        {"emailIndex":0,"action":"skip","confidence":0.3,"reason":"x"}, // note
      ]`
      const r = parseAiActions(txt)
      expect(r).toHaveLength(1)
    })
  })
})

// ---- salvageTruncatedArray -------------------------------------------------

describe('salvageTruncatedArray', () => {
  it('recovers prefix when last item is truncated mid-string', () => {
    const txt = '[ {"emailIndex":0,"action":"move","targetFolderPath":"A","confidence":0.9,"reason":"ok"}, {"emailIndex":1,"action":"move","targetFolderPath":"B/中文還沒打完'
    const salvaged = salvageTruncatedArray(txt)
    expect(salvaged).not.toBeNull()
    const parsed = parseAiActions(salvaged!)
    expect(parsed).toHaveLength(1)
    expect(parsed[0]!.emailIndex).toBe(0)
  })

  it('returns null when truncated before first complete object', () => {
    const txt = '[ {"emailIndex":0,"action":"mov'
    expect(salvageTruncatedArray(txt)).toBeNull()
  })

  it('handles escaped quotes correctly', () => {
    const txt = '[ {"emailIndex":0,"action":"move","targetFolderPath":"A","confidence":0.9,"reason":"with \\"quote\\""}, {"emailIndex":1,"action":"de'
    const salvaged = salvageTruncatedArray(txt)
    expect(salvaged).not.toBeNull()
    const parsed = parseAiActions(salvaged!)
    expect(parsed).toHaveLength(1)
  })

  it('handles brackets inside strings without confusing depth tracker', () => {
    const txt = '[ {"emailIndex":0,"action":"move","targetFolderPath":"A","confidence":0.9,"reason":"contains [bracket] inside"}, {"emailIndex":1,"action":"sk'
    const salvaged = salvageTruncatedArray(txt)
    expect(salvaged).not.toBeNull()
    expect(parseAiActions(salvaged!)).toHaveLength(1)
  })
})

// ---- buildFolderBlock ------------------------------------------------------

describe('buildFolderBlock', () => {
  it('lists real paths', () => {
    const block = buildFolderBlock(tree(), [])
    expect(block).toContain('內部資料')
    expect(block).toContain('內部資料/工時審閱')
  })

  it('excludes paths under prefix', () => {
    const block = buildFolderBlock(tree(), ['05已完成案件'])
    expect(block).not.toContain('05已完成案件/old')
    expect(block).not.toContain('- 05已完成案件\n')
    expect(block).toContain('排除清單')
  })
})

// ---- buildEmailBlock -------------------------------------------------------

function email(over: Partial<Email> & { Id: string }): Email {
  return {
    Subject: '',
    BodyPreview: '',
    From: { EmailAddress: { Address: 'x@y.com', Name: 'X' } },
    ToRecipients: [],
    ReceivedDateTime: '2026-01-01T00:00:00Z',
    ParentFolderId: 'inbox',
    ...over,
  } as Email
}

describe('buildEmailBlock', () => {
  it('numbers emails [0], [1], ...', () => {
    const block = buildEmailBlock([
      email({ Id: 'e1', Subject: 'hello' }),
      email({ Id: 'e2', Subject: 'world' }),
    ])
    expect(block).toContain('[0]')
    expect(block).toContain('[1]')
    expect(block).toContain('hello')
    expect(block).toContain('world')
  })

  it('truncates long previews to 200 chars', () => {
    const longText = 'A'.repeat(500)
    const block = buildEmailBlock([email({ Id: 'e', BodyPreview: longText })])
    expect(block).not.toContain('A'.repeat(201))
  })

  // B2-B: case-number detection line
  it('surfaces a detected court case number from the subject', () => {
    const block = buildEmailBlock([
      email({ Id: 'e', Subject: '關於 112年度訴字第204號 開庭通知' }),
    ])
    expect(block).toMatch(/識別碼:.*112訴204/)
  })

  it('detects a case number from the body preview when the subject has none', () => {
    const block = buildEmailBlock([
      email({ Id: 'e', Subject: '開庭通知', BodyPreview: '案號 112訴204 敬請出席' }),
    ])
    expect(block).toMatch(/識別碼:.*112訴204/)
  })

  it('omits the 識別碼 line when no case number is present', () => {
    const block = buildEmailBlock([email({ Id: 'e', Subject: '午餐揪團' })])
    expect(block).not.toContain('識別碼:')
  })

  // B2-C: soft thread hint line, keyed by email Id
  it('renders a thread hint for the matching email only', () => {
    const block = buildEmailBlock(
      [email({ Id: 'e1', Subject: 'a' }), email({ Id: 'e2', Subject: 'b' })],
      { threadHints: { e2: '此對話近期曾歸於「客戶A/訴訟」（僅供參考）' } },
    )
    const [rec0, rec1] = block.split('\n\n')
    expect(rec0).not.toContain('線索:')
    expect(rec1).toContain('線索:此對話近期曾歸於「客戶A/訴訟」')
  })
})

// ---- buildCaseMapBlock (B2-B) ----------------------------------------------

function caseRule(over: Partial<Rule>): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'case_code',
    signal: '25A0067A',
    targetFolderId: 'fid',
    targetFolderPath: '03/X 案',
    confidence: 0.9,
    matchCount: 10,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  } as Rule
}

describe('buildCaseMapBlock', () => {
  it('is empty when no case rules exist', () => {
    expect(buildCaseMapBlock([caseRule({ type: 'domain', signal: 'x.com' })])).toBe('')
  })

  it('maps case_code and court-case subject rules to their folders', () => {
    const block = buildCaseMapBlock([
      caseRule({ type: 'case_code', signal: '25a0067a', targetFolderPath: '03/X 案' }),
      caseRule({ type: 'subject_keyword', signal: '112訴204', targetFolderPath: '03/Y 案' }),
    ])
    expect(block).toMatch(/25A0067A → 03\/X 案/) // case_code upper-cased
    expect(block).toMatch(/112訴204 → 03\/Y 案/)
  })

  it('excludes a subject_keyword rule whose signal is not a bare case number', () => {
    const block = buildCaseMapBlock([
      caseRule({ type: 'subject_keyword', signal: '請款通知', targetFolderPath: '05/發票' }),
    ])
    expect(block).toBe('')
  })

  it('skips disabled / orphaned rules', () => {
    const block = buildCaseMapBlock([
      caseRule({ signal: '25A0001A', enabled: false }),
      caseRule({ signal: '25A0002A', orphaned: true }),
    ])
    expect(block).toBe('')
  })
})

// ---- buildActiveFoldersBlock (B2-D) ----------------------------------------

describe('buildActiveFoldersBlock', () => {
  it('is empty for undefined / empty input', () => {
    expect(buildActiveFoldersBlock(undefined, [])).toBe('')
    expect(buildActiveFoldersBlock([], [])).toBe('')
  })

  it('lists folders and drops excluded / duplicate paths', () => {
    const block = buildActiveFoldersBlock(
      ['03/進行中', '05已完成案件/old', '03/進行中', '03/另一案'],
      ['05已完成案件'],
    )
    expect(block).toContain('本工具近期歸檔的資料夾')
    expect(block).toContain('- 03/進行中')
    expect(block).toContain('- 03/另一案')
    expect(block).not.toContain('05已完成案件')
    // de-duplicated
    expect(block.match(/03\/進行中/g)).toHaveLength(1)
  })

  it('caps the list at 15 folders', () => {
    const many = Array.from({ length: 30 }, (_, i) => `F/${i}`)
    const block = buildActiveFoldersBlock(many, [])
    expect(block.split('\n').filter((l) => l.startsWith('- '))).toHaveLength(15)
  })
})

// ---- actionToPlanItem ------------------------------------------------------

describe('actionToPlanItem', () => {
  it('resolves valid path to id (source=ai)', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      { emailIndex: 0, action: 'move', targetFolderPath: '內部資料/工時審閱', confidence: 0.9, reason: 'r' },
      tree(),
    )
    expect(item.action).toBe('move')
    expect(item.targetFolderId).toBe('f1a')
    expect(item.source).toBe('ai')
  })

  it('marks unresolved when path missing in tree', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      { emailIndex: 0, action: 'move', targetFolderPath: 'imaginary/path', confidence: 0.9, reason: 'r' },
      tree(),
    )
    expect(item.source).toBe('unresolved')
    expect(item.reason).toContain('AI 指定的路徑不存在')
  })

  it('passes through delete', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      { emailIndex: 0, action: 'delete', confidence: 0.95, reason: '電子報' },
      tree(),
    )
    expect(item.action).toBe('delete')
    expect(item.source).toBe('ai')
  })

  it('passes through new_folder when name provided', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      {
        emailIndex: 0,
        action: 'new_folder',
        suggestedFolderName: 'AA',
        suggestedParentPath: '內部資料',
        confidence: 0.9,
        reason: 'r',
      },
      tree(),
    )
    expect(item.action).toBe('new_folder')
    expect(item.suggestedFolderName).toBe('AA')
  })

  it('clamps invalid confidence to 0-1', () => {
    const tooHigh = actionToPlanItem(
      email({ Id: 'e' }),
      { emailIndex: 0, action: 'delete', confidence: 5, reason: 'r' },
      tree(),
    )
    expect(tooHigh.confidence).toBe(1)

    const negative = actionToPlanItem(
      email({ Id: 'e' }),
      { emailIndex: 0, action: 'delete', confidence: -3, reason: 'r' },
      tree(),
    )
    expect(negative.confidence).toBe(0)
  })

  // ---- new_folder parent validation (added 2026-05-22) -----------------
  //
  // Without this, AI hallucinating a parent path (or typo) caused a 404
  // at folder-create time mid-execute. Now caught at plan-build time and
  // demoted to unresolved so the user can pick the right parent in the
  // plan UI.
  it('marks new_folder unresolved when suggestedParentPath missing', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      {
        emailIndex: 0,
        action: 'new_folder',
        suggestedFolderName: 'New',
        suggestedParentPath: 'imaginary/parent',
        confidence: 0.9,
        reason: 'r',
      },
      tree(),
    )
    expect(item.action).toBe('skip')
    expect(item.source).toBe('unresolved')
    expect(item.reason).toContain('父資料夾')
    // aiOriginal* preserved so user sees AI's intent
    expect(item.aiOriginalAction).toBe('new_folder')
    expect(item.aiOriginalSuggestedFolderName).toBe('New')
  })

  it('keeps new_folder when suggestedParentPath valid', () => {
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      {
        emailIndex: 0,
        action: 'new_folder',
        suggestedFolderName: 'New',
        suggestedParentPath: '內部資料',
        confidence: 0.9,
        reason: 'r',
      },
      tree(),
    )
    expect(item.action).toBe('new_folder')
    expect(item.source).toBe('ai')
  })

  it('keeps new_folder when suggestedParentPath omitted (top-level)', () => {
    // Empty parentPath → create at root, no validation needed.
    const item = actionToPlanItem(
      email({ Id: 'e' }),
      {
        emailIndex: 0,
        action: 'new_folder',
        suggestedFolderName: 'New',
        confidence: 0.9,
        reason: 'r',
      },
      tree(),
    )
    expect(item.action).toBe('new_folder')
    expect(item.source).toBe('ai')
  })
})

// ---- buildEmailBlock long-subject truncation (added 2026-05-22) -----------

describe('buildEmailBlock subject truncation', () => {
  it('truncates subject to 200 chars + ellipsis', () => {
    const longSubject = 'a'.repeat(300)
    const block = buildEmailBlock([
      {
        Id: 'e',
        Subject: longSubject,
        From: { EmailAddress: { Address: 'x@y.com' } },
      } as Email,
    ])
    expect(block).toContain('a'.repeat(200) + '…')
    expect(block).not.toContain('a'.repeat(201))
  })

  it('leaves short subjects intact', () => {
    const block = buildEmailBlock([
      {
        Id: 'e',
        Subject: 'short subject',
        From: { EmailAddress: { Address: 'x@y.com' } },
      } as Email,
    ])
    expect(block).toContain('short subject')
    expect(block).not.toContain('…')
  })
})

// ---- classifyBatch exemplars privacy gate (added 2026-05-22) ---------------
//
// `settings.aiIncludeFewShotExamples` controls whether user-validated
// rules' target paths get sent to Claude as few-shot examples. The
// underlying `selectExemplars` / `buildExamplesBlock` are exercised in
// tests/exemplars.test.ts; here we verify the GATE actually omits the
// examples block from the outgoing request body when the flag is off.

const exemplarRule = (overrides: Partial<Rule> = {}): Rule => ({
  id: 'r-' + Math.random().toString(36).slice(2),
  type: 'domain',
  signal: 'sensitive-client.com',
  targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
  targetFolderPath: '03/案件/敏感客戶',
  confidence: 0.9,
  matchCount: 30,
  enabled: true,
  createdAt: '2026-01-01T00:00:00.000Z',
  source: 'user_manual',
  ...overrides,
})

function mockClaudeResponse(): Response {
  const body = {
    content: [
      { type: 'text', text: '[{"emailIndex":0,"action":"skip","confidence":0.5,"reason":"x"}]' },
    ],
    usage: {
      input_tokens: 100,
      output_tokens: 20,
      cache_creation_input_tokens: 0,
      cache_read_input_tokens: 0,
    },
    stop_reason: 'end_turn',
  }
  return new Response(JSON.stringify(body), { status: 200 })
}

describe('classifyBatch exemplars opt-out', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  async function captureRequestBody(
    settings: Settings,
    rules: Rule[],
  ): Promise<string> {
    let capturedBody = ''
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url: string, init: RequestInit) => {
        capturedBody = typeof init.body === 'string' ? init.body : ''
        return mockClaudeResponse()
      }),
    )
    await classifyBatch(
      {
        emails: [
          {
            Id: 'e1',
            Subject: 'hi',
            From: { EmailAddress: { Address: 'x@y.com' } },
          } as Email,
        ],
        folderTree: tree(),
        excludePrefixes: [],
        rules,
      },
      settings,
    )
    return capturedBody
  }

  it('OMITS examples block when aiIncludeFewShotExamples=false', async () => {
    const settings: Settings = {
      ...DEFAULT_SETTINGS,
      claudeApiKey: 'sk-ant-test',
      aiIncludeFewShotExamples: false,
    }
    const body = await captureRequestBody(settings, [exemplarRule()])
    // The describeRule path used by buildExamplesBlock wouldn't be in
    // the body. The sensitive path should NEVER appear.
    expect(body).not.toContain('sensitive-client.com')
    expect(body).not.toContain('敏感客戶')
  })

  it('INCLUDES examples block when aiIncludeFewShotExamples=true (default)', async () => {
    const settings: Settings = {
      ...DEFAULT_SETTINGS,
      claudeApiKey: 'sk-ant-test',
      aiIncludeFewShotExamples: true,
    }
    const body = await captureRequestBody(settings, [exemplarRule()])
    // With the flag on AND a high-confidence high-matchCount user_manual
    // rule (qualifies for selectExemplars), the target path SHOULD be
    // in the request body.
    expect(body).toContain('敏感客戶')
  })

  it('OMITS examples block when rules array is empty even with flag on', async () => {
    const settings: Settings = {
      ...DEFAULT_SETTINGS,
      claudeApiKey: 'sk-ant-test',
      aiIncludeFewShotExamples: true,
    }
    const body = await captureRequestBody(settings, [])
    expect(body).not.toContain('敏感客戶')
  })
})
