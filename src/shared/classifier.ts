// Claude API classifier.
//
// Two-stage classification pipeline expects this module to:
//   1. Accept a batch of emails + folder tree + settings
//   2. Build a cached prompt prefix (system + folder block) and a fresh user
//      block (the email batch)
//   3. Call Anthropic /v1/messages, parse the JSON array response
//   4. Map each AiAction back to a PlanItem (resolving folder paths to IDs)
//
// Prompt caching:
//   - System block carries cache_control: ephemeral
//   - Folder tree user-block carries cache_control: ephemeral
//   - Email user-block is fresh
// Subsequent batches with the same model + folder tree hit cache, saving ~4K
// input tokens per call.

import { flattenFolderTree } from './outlook-api'
import { decodeCompound } from './rules'
import type { Email, MailFolderNode, PlanItem, Rule, Settings } from './types'

const API_URL = 'https://api.anthropic.com/v1/messages'
const ANTHROPIC_VERSION = '2023-06-01'
// Extended prompt cache TTL (1 hour). Default ephemeral is 5 min — too
// short for a lawyer who classifies morning inbox then comes back hours
// later for the afternoon. With this beta + ttl:'1h' on cache_control
// markers, the same system prompt + folder tree + examples cache hits
// across longer gaps without re-creation cost.
const ANTHROPIC_BETA = 'extended-cache-ttl-2025-04-11'

/**
 * Build the system prompt from settings. All domain-specific phrasing
 * (firm / lawyer / case file) is parameterised so a non-lawyer user
 * gets a neutral "email-organising assistant" prompt instead of one
 * that talks about cases they don't have.
 *
 * What changes with settings:
 *   - `primaryRootPath`  → example path in the JSON sample; default
 *     placeholder shows generic "<根資料夾>/<分類>/<子分類>" when unset.
 *   - `internalDomains`  → adds a rule line that tells the model to
 *     categorise mail from these domains by subject (since "all mail
 *     from @firm.com" rarely maps to one folder). Omitted when empty.
 *   - `internalSubjectCategories` → appended as concrete examples to
 *     the internal-mail rule. Omitted when no internal domains anyway.
 *
 * The Taiwan court / case-code regexes still run downstream (rules.ts
 * extractors), but they fire 0 hits when subjects don't match — so
 * keeping the prompt neutral doesn't break the lawyer use case.
 *
 * Cache impact: every settings change invalidates the prompt cache
 * (text bytes change). In practice these settings rarely change after
 * onboarding, so prompt caching still pays off for daily workflow.
 */
function buildSystemPrompt(settings: Settings): string {
  const exampleRoot = settings.primaryRootPath || '<根資料夾>'
  const hasInternal = settings.internalDomains.length > 0
  const domainList = settings.internalDomains.map((d) => `@${d}`).join(' / ')
  const catList =
    settings.internalSubjectCategories.length > 0
      ? `,類別範例:${settings.internalSubjectCategories.join('、')}`
      : ''
  const internalLine = hasInternal
    ? `內部 ${domainList} 寄收:依主旨判斷該歸到哪一個資料夾${catList}`
    : ''
  const numberedRules = [
    'targetFolderPath / suggestedParentPath 必須出現在「可用資料夾」清單中,不要編造路徑',
    '排除清單下的子資料夾不可作為 targetFolderPath',
    'confidence 0.0–1.0:1.0 = 精確識別碼匹配;0.7+ = 寄件人/收件人網域強相關;< 0.5 應考慮改 "skip"',
    ...(internalLine ? [internalLine] : []),
    '外部寄件人 / 收件人:用網域對應到主題相關的資料夾',
    '主旨含明確識別碼(編號、案號、訂單號等)的、優先用該識別碼定位',
    'reason 30 字內、講明判斷的 key signal',
  ]
    .map((line, i) => `${i + 1}. ${line}`)
    .join('\n')

  return `你是郵件歸類助手。使用者會給你他的資料夾結構與一批郵件,請逐封決定歸類動作。

可用動作(嚴格四選一):
- "move"        移到一個既有資料夾,targetFolderPath 必填
- "delete"      刪除(明確無關郵件:電子報、廣告、活動通知)
- "new_folder"  需要新建資料夾,suggestedFolderName + suggestedParentPath 必填
- "skip"        判斷不出來、低信心,使用者會手動處理

判斷原則:
${numberedRules}

嚴格 JSON array、不要任何前後文、不要 markdown fence:
[
  { "emailIndex": 0, "action": "move", "targetFolderPath": "${exampleRoot}/<分類>/<子分類>", "confidence": 0.9, "reason": "寄件網域 @<外部網域>" },
  { "emailIndex": 1, "action": "delete", "confidence": 0.95, "reason": "電子報通訊" }
]
`
}

export class ClassifierError extends Error {
  constructor(
    public readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'ClassifierError'
  }
}

export type AiAction = {
  emailIndex: number
  action: 'move' | 'delete' | 'new_folder' | 'skip'
  targetFolderPath?: string
  suggestedFolderName?: string
  suggestedParentPath?: string
  confidence: number
  reason: string
}

export type ClassifierUsage = {
  inputTokens: number
  outputTokens: number
  cacheCreationTokens: number
  cacheReadTokens: number
}

export type ClassifierResult = {
  plan: PlanItem[]
  rawResponse: string
  usage: ClassifierUsage
  /**
   * Present when the model hit max_tokens. Indicates how many emails were
   * actually classified before truncation; the remainder are stub-skipped via
   * the backstop and should be flagged to the user.
   */
  truncated?: { parsedCount: number; totalRequested: number }
}

export type ClassifierInput = {
  emails: Email[]
  folderTree: MailFolderNode[]
  excludePrefixes: string[]
  /**
   * Optional pool of known rules to mine for few-shot examples. Passing the
   * full set is fine — `selectExemplars` picks the most informative subset.
   * Omit to skip the examples block entirely (saves some prompt tokens but
   * loses in-context learning).
   */
  rules?: Rule[]
}

// ---- Prompt building -------------------------------------------------------

function pathExcluded(path: string, excludePrefixes: string[]): boolean {
  for (const p of excludePrefixes) {
    if (path === p) return true
    if (path.startsWith(p + '/')) return true
  }
  return false
}

// ---- Few-shot examples -----------------------------------------------------

const EXEMPLAR_MAX = 8
const EXEMPLAR_MIN_MATCH_COUNT = 5

/**
 * Pick a small set of high-quality rules to use as in-context examples.
 * Goals:
 *   - Prefer user-validated sources (user_manual > ai_overridden) over
 *     auto-scan or AI-confirmed (which can be noisy).
 *   - Skip disabled / orphaned rules — they teach the wrong thing.
 *   - Hit at least one example per rule type when available so the AI sees
 *     domain / case_code / subject_keyword / sender / compound diversity.
 *   - Cap at EXEMPLAR_MAX so prompt-cache stays warm and we don't pay for
 *     a wall of examples.
 */
export function selectExemplars(rules: Rule[]): Rule[] {
  const candidates = rules.filter(
    (r) => r.enabled && !r.orphaned && r.matchCount >= EXEMPLAR_MIN_MATCH_COUNT,
  )
  // Source priority: user_manual = 3 (gold standard, hand-built),
  // ai_overridden = 2 (user corrected AI), ai_confirmed = 1, auto_scan = 0.
  const sourceRank: Record<Rule['source'], number> = {
    user_manual: 3,
    ai_overridden: 2,
    ai_confirmed: 1,
    auto_scan: 0,
  }
  const sorted = [...candidates].sort((a, b) => {
    const s = sourceRank[b.source] - sourceRank[a.source]
    if (s !== 0) return s
    if (b.matchCount !== a.matchCount) return b.matchCount - a.matchCount
    return a.signal.localeCompare(b.signal)
  })

  // Greedy fill, prefer type diversity. Walk the sorted list twice: first
  // pass picks at most one per type, second pass fills remaining slots from
  // whatever's left.
  const picked: Rule[] = []
  const seenTypes = new Set<string>()
  for (const r of sorted) {
    if (picked.length >= EXEMPLAR_MAX) break
    if (seenTypes.has(r.type)) continue
    picked.push(r)
    seenTypes.add(r.type)
  }
  for (const r of sorted) {
    if (picked.length >= EXEMPLAR_MAX) break
    if (picked.includes(r)) continue
    picked.push(r)
  }
  return picked
}

function describeRule(r: Rule): string {
  switch (r.type) {
    case 'domain':
      return `寄件人網域 @${r.signal.replace(/^@/, '')}`
    case 'sender':
      return `寄件人 ${r.signal}`
    case 'case_code':
      return `主旨含案件代號 ${r.signal}`
    case 'subject_keyword':
      return `主旨含「${r.signal}」`
    case 'compound': {
      const parsed = decodeCompound(r.signal)
      if (!parsed) return r.signal
      const parts = parsed.conditions.map((c) => {
        if (c.type === 'domain') return `網域 @${c.value.replace(/^@/, '')}`
        if (c.type === 'sender') return `寄件人 ${c.value}`
        return `主旨含「${c.value}」`
      })
      return parts.join(' 且 ')
    }
  }
}

export function buildExamplesBlock(exemplars: Rule[]): string {
  if (exemplars.length === 0) return ''
  const lines = exemplars.map((r) => `- ${describeRule(r)} → ${r.targetFolderPath}`)
  return [
    '歷史分類範例（律師確認過的規則，僅供參考、新郵件仍需獨立判斷）：',
    lines.join('\n'),
  ].join('\n')
}

export function buildFolderBlock(tree: MailFolderNode[], excludePrefixes: string[]): string {
  const flat = flattenFolderTree(tree)
  const lines = flat
    .filter((n) => !pathExcluded(n.path, excludePrefixes))
    .map((n) => '- ' + n.path)
  const excludeLines = excludePrefixes.map((p) => '- ' + p + '（及其子資料夾全部排除）')
  return [
    '可用資料夾（target 必須出自此清單）：',
    lines.join('\n'),
    '',
    '排除清單（不可作為 target）：',
    excludeLines.join('\n'),
  ].join('\n')
}

function recipientAddrs(list?: Email['ToRecipients']): string {
  if (!list) return ''
  return list
    .map((r) => r.EmailAddress?.Address)
    .filter(Boolean)
    .slice(0, 5)
    .join(', ')
}

// Subjects can technically be arbitrary length but rarely carry useful
// classification signal past 200 chars. Truncating defensively keeps the
// prompt cheap and protects against accidental DoS via a single huge
// subject blowing up the request budget.
const MAX_SUBJECT_LEN = 200

export function buildEmailBlock(emails: Email[]): string {
  return emails
    .map((m, i) => {
      const fromName = m.From?.EmailAddress?.Name ?? ''
      const fromAddr = m.From?.EmailAddress?.Address ?? '(unknown)'
      const to = recipientAddrs(m.ToRecipients)
      const cc = recipientAddrs(m.CcRecipients)
      const date = (m.ReceivedDateTime ?? '').slice(0, 10)
      const subjectRaw = m.Subject ?? ''
      const subject =
        subjectRaw.length > MAX_SUBJECT_LEN
          ? subjectRaw.slice(0, MAX_SUBJECT_LEN) + '…'
          : subjectRaw
      const preview = (m.BodyPreview ?? '').slice(0, 200).replace(/\s+/g, ' ').trim()
      return `[${i}] 主旨:${subject || '(空主旨)'}
    寄件:${fromName ? fromName + ' ' : ''}<${fromAddr}>
    收件:${to}${cc ? ' / CC: ' + cc : ''}
    日期:${date}
    預覽:${preview || '(無預覽)'}`
    })
    .join('\n\n')
}

// ---- Response parsing ------------------------------------------------------

/**
 * Find the first balanced [...] in `text` starting at or after `startIdx`.
 * Tracks string context to ignore brackets inside strings. Returns null if no
 * balanced array found. Used as fallback when AI wraps the JSON in prose.
 */
function findBalancedArray(text: string, startIdx = 0): string | null {
  const open = text.indexOf('[', startIdx)
  if (open < 0) return null
  let depth = 0
  let inString = false
  let escape = false
  for (let i = open; i < text.length; i++) {
    const ch = text[i]
    if (escape) {
      escape = false
      continue
    }
    if (inString) {
      if (ch === '\\') escape = true
      else if (ch === '"') inString = false
      continue
    }
    if (ch === '"') {
      inString = true
      continue
    }
    if (ch === '[') depth++
    else if (ch === ']') {
      depth--
      if (depth === 0) return text.slice(open, i + 1)
    }
  }
  return null
}

/**
 * Common "loose JSON" repairs that Claude (and other LLMs) occasionally
 * emit despite the strict instructions:
 *   - Trailing commas before `]` or `}` — most frequent issue.
 *   - `//` line comments inside the array (helpful but invalid).
 *   - Smart quotes (" / " instead of ").
 *
 * Strip-and-retry runs ONLY after the strict parse and balanced-array
 * parse both fail, so the happy path stays cheap. Returns the repaired
 * text or null when no repair would change the input.
 */
function loosenJson(text: string): string | null {
  let out = text
  // Strip // line comments (avoid stripping :// in URLs by requiring the
  // comment to start at a line beginning or after whitespace).
  out = out.replace(/(^|\s)\/\/[^\n]*/g, '$1')
  // Strip /* */ block comments.
  out = out.replace(/\/\*[\s\S]*?\*\//g, '')
  // Drop trailing commas before } or ].
  out = out.replace(/,(\s*[}\]])/g, '$1')
  // Smart quotes → straight (only when they look JSON-like, surrounding
  // keys/values; conservative replace inside strings could break content).
  out = out.replace(/[“”]/g, '"')
  return out !== text ? out : null
}

export function parseAiActions(content: string): AiAction[] {
  const trimmed = content.trim()
  try {
    const direct = JSON.parse(trimmed)
    if (Array.isArray(direct)) return direct as AiAction[]
  } catch {
    /* fall through */
  }

  // Search for a balanced array (handles AI prose-wrapped output, markdown
  // fences, etc.). Try multiple starting positions in case of nested examples.
  let startFrom = 0
  while (startFrom < trimmed.length) {
    const candidate = findBalancedArray(trimmed, startFrom)
    if (!candidate) break
    try {
      const parsed = JSON.parse(candidate)
      if (Array.isArray(parsed)) return parsed as AiAction[]
    } catch {
      /* continue searching past this candidate */
    }
    startFrom = trimmed.indexOf(candidate, startFrom) + candidate.length
  }

  // Last resort: apply common "loose JSON" repairs (trailing commas /
  // line comments / smart quotes), then retry both the direct parse
  // and the balanced-array search. Many real-world LLM mishaps fall
  // into these categories.
  const loosened = loosenJson(trimmed)
  if (loosened) {
    try {
      const direct = JSON.parse(loosened)
      if (Array.isArray(direct)) return direct as AiAction[]
    } catch {
      /* keep going */
    }
    const candidate = findBalancedArray(loosened, 0)
    if (candidate) {
      try {
        const parsed = JSON.parse(candidate)
        if (Array.isArray(parsed)) return parsed as AiAction[]
      } catch {
        /* fall through */
      }
    }
  }

  throw new ClassifierError(
    'PARSE_ERROR',
    `Claude 回應無法解析為 JSON array。前 200 字:${trimmed.slice(0, 200)}`,
  )
}

/**
 * Recover what we can from a max_tokens-truncated response.
 * Walks `[ {item}, {item}, {trun…` to the last complete `}` at array depth,
 * then closes the array. Returns null if even the first item is incomplete.
 */
export function salvageTruncatedArray(content: string): string | null {
  const arrayStart = content.indexOf('[')
  if (arrayStart < 0) return null

  let lastCompleteObjectEnd = -1
  let objectDepth = 0
  let inString = false
  let escape = false

  // Track only object-brace depth ({}); array-brackets stay implicit (we know
  // we're inside the top-level array).
  for (let i = arrayStart + 1; i < content.length; i++) {
    const ch = content[i]
    if (escape) {
      escape = false
      continue
    }
    if (inString) {
      if (ch === '\\') escape = true
      else if (ch === '"') inString = false
      continue
    }
    if (ch === '"') {
      inString = true
      continue
    }
    if (ch === '{') objectDepth++
    else if (ch === '}') {
      objectDepth--
      if (objectDepth === 0) lastCompleteObjectEnd = i
    }
  }

  if (lastCompleteObjectEnd < 0) return null
  // Slice [first…lastComplete}], then close with ]
  return content.slice(arrayStart, lastCompleteObjectEnd + 1) + ']'
}

// ---- Action → PlanItem -----------------------------------------------------

function clampConfidence(n: unknown): number {
  if (typeof n !== 'number' || !Number.isFinite(n)) return 0
  return Math.max(0, Math.min(1, n))
}

export function actionToPlanItem(
  email: Email,
  action: AiAction,
  folderTree: MailFolderNode[],
): PlanItem {
  const base = {
    emailId: email.Id,
    emailSubject: email.Subject ?? '',
    emailFrom: email.From?.EmailAddress?.Address ?? '',
    bodyPreview: email.BodyPreview ?? '',
    conversationId: email.ConversationId,
    confidence: clampConfidence(action.confidence),
    reason: (action.reason ?? '').trim().slice(0, 200),
  }

  if (action.action === 'move' && action.targetFolderPath) {
    const node = flattenFolderTree(folderTree).find((n) => n.path === action.targetFolderPath)
    return {
      ...base,
      action: 'move',
      targetFolderId: node?.id,
      targetFolderPath: action.targetFolderPath,
      source: node ? 'ai' : 'unresolved',
      // Snapshot the AI's verdict so execute can detect user overrides later.
      aiOriginalAction: 'move',
      aiOriginalTargetFolderPath: action.targetFolderPath,
      ...(node ? {} : { reason: base.reason + '（AI 指定的路徑不存在）' }),
    }
  }

  if (action.action === 'delete') {
    return {
      ...base,
      action: 'delete',
      source: 'ai',
      aiOriginalAction: 'delete',
    }
  }

  if (action.action === 'new_folder' && action.suggestedFolderName) {
    // Validate suggestedParentPath against the live tree — AI sometimes
    // hallucinates a parent that doesn't exist (e.g. typo, or invents
    // a "logical" parent that the user actually doesn't have). Without
    // this check, execute hits a 404 at folder-create time. Marking
    // unresolved here lets the user pick the right parent in the plan UI.
    const parentPath = action.suggestedParentPath
    const parentValid =
      !parentPath ||
      flattenFolderTree(folderTree).some((n) => n.path === parentPath)
    if (!parentValid) {
      return {
        ...base,
        action: 'skip',
        source: 'unresolved',
        reason: base.reason + `(AI 建議的父資料夾「${parentPath}」不存在)`,
        // Preserve AI's intent so popup can show what it WANTED to do.
        aiOriginalAction: 'new_folder',
        aiOriginalSuggestedFolderName: action.suggestedFolderName,
        aiOriginalSuggestedParentPath: parentPath,
      }
    }
    return {
      ...base,
      action: 'new_folder',
      suggestedFolderName: action.suggestedFolderName,
      suggestedParentPath: action.suggestedParentPath,
      source: 'ai',
      aiOriginalAction: 'new_folder',
      aiOriginalSuggestedFolderName: action.suggestedFolderName,
      aiOriginalSuggestedParentPath: action.suggestedParentPath,
    }
  }

  return { ...base, action: 'skip', source: 'unresolved' }
}

// ---- Main entry ------------------------------------------------------------

export async function classifyBatch(
  input: ClassifierInput,
  settings: Settings,
): Promise<ClassifierResult> {
  if (!settings.claudeApiKey) {
    throw new ClassifierError('NO_API_KEY', 'Claude API key 未設定')
  }
  if (input.emails.length === 0) {
    return {
      plan: [],
      rawResponse: '',
      usage: { inputTokens: 0, outputTokens: 0, cacheCreationTokens: 0, cacheReadTokens: 0 },
    }
  }

  const folderBlock = buildFolderBlock(input.folderTree, input.excludePrefixes)
  const emailBlock = buildEmailBlock(input.emails)
  // Few-shot mining is privacy-sensitive: user rule target paths (e.g.
  // "客戶X 離婚案") would round-trip through the LLM. Honour the
  // settings.aiIncludeFewShotExamples flag — defaults to ON because
  // examples improve accuracy meaningfully, but lawyers / others with
  // sensitive folder names can disable in Options.
  const exemplars =
    input.rules && settings.aiIncludeFewShotExamples
      ? selectExemplars(input.rules)
      : []
  const examplesBlock = buildExamplesBlock(exemplars)

  // Output budget scales with batch. Chinese tokens are ~1.5× English in
  // Anthropic's tokenizer; folder paths + reasons are mostly Chinese, so we
  // need a generous per-email allowance. Sonnet 4.6 supports max_tokens up to
  // 64K so we can afford to be wide.
  // 50 emails → ~22K | 100 emails → ~42K | hard cap 64K
  const maxTokens = Math.min(64000, 2048 + input.emails.length * 400)

  const body = {
    model: settings.claudeModel,
    max_tokens: maxTokens,
    system: [
      {
        type: 'text',
        text: buildSystemPrompt(settings),
        cache_control: { type: 'ephemeral', ttl: '1h' },
      },
    ],
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: folderBlock,
            cache_control: { type: 'ephemeral', ttl: '1h' },
          },
          // Examples block — only included when we have validated exemplars
          // to share. cache_control too so reruns with the same rule set
          // share the cache hit even though the email batch changes.
          ...(examplesBlock
            ? [
                {
                  type: 'text' as const,
                  text: examplesBlock,
                  cache_control: { type: 'ephemeral' as const, ttl: '1h' as const },
                },
              ]
            : []),
          {
            type: 'text',
            text: `請分類以下 ${input.emails.length} 封郵件：\n\n${emailBlock}`,
          },
        ],
      },
    ],
  }

  let resp: Response
  try {
    resp = await fetch(API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': settings.claudeApiKey,
        'anthropic-version': ANTHROPIC_VERSION,
        'anthropic-beta': ANTHROPIC_BETA,
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify(body),
    })
  } catch (e) {
    throw new ClassifierError(
      'NETWORK',
      `Claude API 網路錯誤：${e instanceof Error ? e.message : String(e)}`,
    )
  }

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '')
    throw new ClassifierError(
      `HTTP_${resp.status}`,
      `Claude API ${resp.status} ${resp.statusText}: ${errText.slice(0, 300)}`,
    )
  }

  const data = (await resp.json()) as {
    content: Array<{ type: string; text?: string }>
    usage: {
      input_tokens: number
      output_tokens: number
      cache_creation_input_tokens?: number
      cache_read_input_tokens?: number
    }
    stop_reason: string
  }

  const text = data.content
    .filter((c) => c.type === 'text' && typeof c.text === 'string')
    .map((c) => c.text!)
    .join('')

  // If the model hit max_tokens, the array is truncated mid-item and
  // parseAiActions would throw. Salvage the complete prefix instead — the
  // backstop downstream marks the unparsed tail as skip.
  let truncated: ClassifierResult['truncated']
  let rawActions: AiAction[]
  if (data.stop_reason === 'max_tokens') {
    const salvaged = salvageTruncatedArray(text)
    if (!salvaged) {
      throw new ClassifierError(
        'PARSE_ERROR',
        `Claude 回應 max_tokens 截斷且無法救回任何項目。前 200 字：${text.slice(0, 200)}`,
      )
    }
    rawActions = parseAiActions(salvaged)
    truncated = { parsedCount: rawActions.length, totalRequested: input.emails.length }
    console.warn(
      `[mail-organizer] Claude hit max_tokens; salvaged ${rawActions.length}/${input.emails.length} items`,
    )
  } else {
    rawActions = parseAiActions(text)
  }

  // Dedup: AI sometimes returns duplicate emailIndex entries (model glitch).
  // Keep the one with highest confidence so we don't move/delete the same
  // email twice.
  const dedupedByIndex = new Map<number, AiAction>()
  for (const action of rawActions) {
    if (typeof action?.emailIndex !== 'number') continue
    const existing = dedupedByIndex.get(action.emailIndex)
    const aConf = clampConfidence(action.confidence)
    const eConf = existing ? clampConfidence(existing.confidence) : -1
    if (!existing || aConf > eConf) dedupedByIndex.set(action.emailIndex, action)
  }
  const actions = [...dedupedByIndex.values()]

  const plan: PlanItem[] = actions.map((action) => {
    const email = input.emails[action.emailIndex]
    if (!email) {
      return {
        emailId: '',
        emailSubject: '',
        emailFrom: '',
        bodyPreview: '',
        action: 'skip',
        confidence: 0,
        reason: `AI emailIndex ${action.emailIndex} 超出範圍`,
        source: 'unresolved',
      }
    }
    return actionToPlanItem(email, action, input.folderTree)
  })

  // Backstop: if AI returned fewer items than emails, mark missing ones as skip
  const covered = new Set(actions.map((a) => a.emailIndex))
  input.emails.forEach((m, i) => {
    if (!covered.has(i)) {
      plan.push({
        emailId: m.Id,
        emailSubject: m.Subject ?? '',
        emailFrom: m.From?.EmailAddress?.Address ?? '',
        bodyPreview: m.BodyPreview ?? '',
        conversationId: m.ConversationId,
        action: 'skip',
        confidence: 0,
        reason: 'AI 未回傳此 index 的判斷',
        source: 'unresolved',
      })
    }
  })

  return {
    plan,
    rawResponse: text,
    usage: {
      inputTokens: data.usage.input_tokens,
      outputTokens: data.usage.output_tokens,
      cacheCreationTokens: data.usage.cache_creation_input_tokens ?? 0,
      cacheReadTokens: data.usage.cache_read_input_tokens ?? 0,
    },
    ...(truncated ? { truncated } : {}),
  }
}
