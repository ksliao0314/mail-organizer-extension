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
import { courtCaseSignal, decodeCompound, extractCaseCodes, extractCourtCaseNumbers } from './rules'
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
    'reason 50 字內、先講判斷的 key signal 再下結論',
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

嚴格 JSON array、不要任何前後文、不要 markdown fence。
每筆先寫 reason(判斷依據)再寫 action —— 讓結論被自己的推理制約、提高準確度:
[
  { "emailIndex": 0, "reason": "寄件網域 @<外部網域>、主旨提及該客戶案件", "action": "move", "targetFolderPath": "${exampleRoot}/<分類>/<子分類>", "confidence": 0.9 },
  { "emailIndex": 1, "reason": "電子報通訊、與案件無關", "action": "delete", "confidence": 0.95 }
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
  /**
   * Per-email soft hints keyed by `email.Id` (B2-C). The preflight attaches a
   * short note for emails whose conversation / subject was recently filed
   * somewhere but was too ambiguous (conflictCount>0) to auto-route — a weak
   * nudge the AI can weigh, explicitly framed as uncertain.
   */
  threadHints?: Record<string, string>
  /**
   * Folder paths this tool has recently filed mail into, most-recent first
   * (B2-D). Rendered as a reference-only block so the AI biases toward
   * still-active case folders instead of misfiling into a closed case's
   * folder. Excluded / capped further inside; omit or [] to skip the block.
   */
  recentFolders?: string[]
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

const EXEMPLAR_MAX = 12
// Source-aware eligibility (優化 2026-07): user-validated rules need no
// statistical proof — a rule the user hand-built or corrected AI to is
// authoritative at matchCount 0. Only auto-derived rules must earn their
// spot with hits.
const EXEMPLAR_MIN_MATCH: Record<Rule['source'], number> = {
  user_manual: 0,
  ai_overridden: 0,
  ai_confirmed: 3,
  auto_scan: 5,
}
// Reserve a few slots for RECENT user corrections regardless of matchCount —
// "AI erred on this class last week and the user fixed it" is the single most
// instructive example, but ai_overridden rules are born at matchCount 0 and
// would otherwise never qualify.
const EXEMPLAR_RECENT_RESERVED = 3
const EXEMPLAR_RECENT_AGE_MS = 30 * 24 * 60 * 60 * 1000
// Cap examples pointing at the SAME folder so a high-traffic folder can't
// monopolize the block and skew the AI toward 1-2 destinations.
const EXEMPLAR_PER_FOLDER_CAP = 2

/**
 * Pick a small set of high-quality rules to use as in-context examples.
 * Goals:
 *   - Prefer user-validated sources (user_manual > ai_overridden) over
 *     auto-scan or AI-confirmed (which can be noisy).
 *   - Skip disabled / orphaned rules — they teach the wrong thing.
 *   - Reserve slots for recent user corrections (in-context learning's
 *     highest-value signal).
 *   - Diversity: one per type first, then cap per target folder.
 *   - Cap at EXEMPLAR_MAX so prompt-cache stays warm.
 * `nowMs` is injectable for deterministic tests.
 */
export function selectExemplars(rules: Rule[], nowMs: number = Date.now()): Rule[] {
  const sourceRank: Record<Rule['source'], number> = {
    user_manual: 3,
    ai_overridden: 2,
    ai_confirmed: 1,
    auto_scan: 0,
  }
  // Secondary sort key: log-bucket the matchCount so tiny per-batch changes
  // (a rule ticking 41→42) don't reshuffle the exemplar set and needlessly
  // invalidate the prompt cache.
  const matchBucket = (r: Rule): number => Math.floor(Math.log2(r.matchCount + 1))
  const bySalience = (a: Rule, b: Rule): number => {
    const s = sourceRank[b.source] - sourceRank[a.source]
    if (s !== 0) return s
    const m = matchBucket(b) - matchBucket(a)
    if (m !== 0) return m
    return a.signal.localeCompare(b.signal)
  }

  const usable = rules.filter((r) => r.enabled && !r.orphaned)
  const picked: Rule[] = []
  const folderCount = new Map<string, number>()
  const canTake = (r: Rule, enforceFolderCap: boolean): boolean => {
    if (picked.includes(r)) return false
    if (enforceFolderCap && (folderCount.get(r.targetFolderPath) ?? 0) >= EXEMPLAR_PER_FOLDER_CAP) {
      return false
    }
    return true
  }
  const take = (r: Rule) => {
    picked.push(r)
    folderCount.set(r.targetFolderPath, (folderCount.get(r.targetFolderPath) ?? 0) + 1)
  }

  // Pass 0: reserve slots for recent user corrections (matchCount-exempt).
  const recent = usable
    .filter(
      (r) =>
        (r.source === 'ai_overridden' || r.source === 'user_manual') &&
        nowMs - Date.parse(r.createdAt) <= EXEMPLAR_RECENT_AGE_MS,
    )
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt))
  for (const r of recent) {
    if (picked.length >= EXEMPLAR_RECENT_RESERVED) break
    if (canTake(r, true)) take(r)
  }

  // Candidates for the main fill: meet the source-aware match threshold.
  const candidates = usable
    .filter((r) => r.matchCount >= EXEMPLAR_MIN_MATCH[r.source])
    .sort(bySalience)

  // Pass 1: type diversity (one per type), honouring the per-folder cap.
  const seenTypes = new Set<string>()
  for (const r of candidates) {
    if (picked.length >= EXEMPLAR_MAX) break
    if (seenTypes.has(r.type)) continue
    if (!canTake(r, true)) continue
    take(r)
    seenTypes.add(r.type)
  }
  // Pass 2: fill remaining slots, still capping per folder.
  for (const r of candidates) {
    if (picked.length >= EXEMPLAR_MAX) break
    if (!canTake(r, true)) continue
    take(r)
  }
  // Pass 3: if still short (small library where the folder cap starved us),
  // relax the folder cap so we don't ship a near-empty examples block.
  for (const r of candidates) {
    if (picked.length >= EXEMPLAR_MAX) break
    if (!canTake(r, false)) continue
    take(r)
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

// Case-number → folder map (B2-B). Unlike `selectExemplars` (a diverse SAMPLE
// capped per folder / per type), this lists EVERY known case identifier so a
// subject bearing case 「112訴500」 can be routed even when the sampled
// exemplars didn't happen to include a sibling case in that folder. Capped so
// a heavy user with hundreds of case rules can't blow the prompt budget.
const CASE_MAP_MAX = 25

export function buildCaseMapBlock(rules: Rule[]): string {
  const seen = new Set<string>()
  const entries: { code: string; path: string; matchCount: number }[] = []
  for (const r of rules) {
    if (!r.enabled || r.orphaned) continue
    // case_code rules store the canonical code as their signal; other types
    // (subject_keyword learned from a court-case subject) qualify only when
    // the whole signal IS a court case number.
    const code = r.type === 'case_code' ? r.signal.trim().toUpperCase() : courtCaseSignal(r.signal)
    if (!code) continue
    const key = code + '→' + r.targetFolderPath
    if (seen.has(key)) continue
    seen.add(key)
    entries.push({ code, path: r.targetFolderPath, matchCount: r.matchCount })
  }
  if (entries.length === 0) return ''
  entries.sort((a, b) => b.matchCount - a.matchCount || a.code.localeCompare(b.code))
  const lines = entries.slice(0, CASE_MAP_MAX).map((e) => `- ${e.code} → ${e.path}`)
  return [
    '已知案號 / 代號對照（同一案號的新郵件通常歸同一資料夾；相近號碼可當推斷線索，但仍需依內容判斷）：',
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

// Recently-active folders (B2-D). folderActivity tells us which case folders
// this tool has been filing into lately; a lawyer's closed cases keep their
// folders forever, so without this the AI happily routes a new email into a
// case that wrapped up two years ago. Reference-only framing — this is a bias
// signal, not a constraint. NOT cached: activity churns every batch, and
// pinning it into the cached prefix would thrash the cache. Capped + exclude-
// filtered (defence in depth; the caller pre-filters too).
const ACTIVE_FOLDERS_MAX = 15

export function buildActiveFoldersBlock(
  recentFolders: string[] | undefined,
  excludePrefixes: string[],
): string {
  if (!recentFolders || recentFolders.length === 0) return ''
  const seen = new Set<string>()
  const picked: string[] = []
  for (const p of recentFolders) {
    if (!p || pathExcluded(p, excludePrefixes) || seen.has(p)) continue
    seen.add(p)
    picked.push(p)
    if (picked.length >= ACTIVE_FOLDERS_MAX) break
  }
  if (picked.length === 0) return ''
  return [
    '本工具近期歸檔的資料夾（僅供參考，代表這些案子仍在進行中；不代表這批郵件一定屬於它們）：',
    picked.map((p) => '- ' + p).join('\n'),
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

/**
 * Extract structured case identifiers (Taiwan court case numbers + Latin case
 * codes) from an email's subject and body preview (B2-B). Subject matches
 * first (higher precision), preview second (fallback coverage). Deduped,
 * capped at 4 so a body listing many cases can't blow up the line. These are
 * high-precision regexes, so surfacing them explicitly stops the AI from
 * missing a case number buried in a long subject.
 */
function detectCaseSignals(subject: string, preview: string): string[] {
  const out: string[] = []
  const seen = new Set<string>()
  const add = (arr: string[]) => {
    for (const s of arr) {
      if (seen.has(s)) continue
      seen.add(s)
      out.push(s)
    }
  }
  add(extractCourtCaseNumbers(subject))
  add(extractCaseCodes(subject))
  add(extractCourtCaseNumbers(preview))
  add(extractCaseCodes(preview))
  return out.slice(0, 4)
}

export function buildEmailBlock(
  emails: Email[],
  opts?: { threadHints?: Record<string, string> },
): string {
  return emails
    .map((m, i) => {
      // Whitespace-collapse Subject and the From display name, same as
      // BodyPreview below (audit P3): the record format is line-oriented
      // ("[i] 主旨:… / 寄件:…"), and RFC 2047 encoded-word headers can
      // decode to contain literal CR/LF — a crafted inbound email could
      // otherwise terminate its own record and append forged lines
      // (including a fake "[k] 主旨:…" shadowing another real email).
      const fromName = (m.From?.EmailAddress?.Name ?? '').replace(/\s+/g, ' ').trim()
      const fromAddr = (m.From?.EmailAddress?.Address ?? '(unknown)').replace(/\s+/g, ' ').trim()
      const to = recipientAddrs(m.ToRecipients)
      const cc = recipientAddrs(m.CcRecipients)
      const date = (m.ReceivedDateTime ?? '').slice(0, 10)
      const subjectRaw = (m.Subject ?? '').replace(/\s+/g, ' ').trim()
      const subject =
        subjectRaw.length > MAX_SUBJECT_LEN
          ? subjectRaw.slice(0, MAX_SUBJECT_LEN) + '…'
          : subjectRaw
      const preview = (m.BodyPreview ?? '').slice(0, 200).replace(/\s+/g, ' ').trim()
      // B2-B: surface detected case identifiers as their own line so the AI
      // treats them as a strong routing signal even when the subject is noisy.
      const cases = detectCaseSignals(subjectRaw, preview)
      const caseLine = cases.length ? `\n    識別碼:${cases.join(', ')}` : ''
      // B2-C: soft thread hint (uncertain — the preflight already auto-routed
      // the confident ones; anything reaching the AI is ambiguous).
      const hint = opts?.threadHints?.[m.Id]
      const hintLine = hint ? `\n    線索:${hint}` : ''
      return `[${i}] 主旨:${subject || '(空主旨)'}
    寄件:${fromName ? fromName + ' ' : ''}<${fromAddr}>
    收件:${to}${cc ? ' / CC: ' + cc : ''}
    日期:${date}
    預覽:${preview || '(無預覽)'}${caseLine}${hintLine}`
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
  // Enforced here, not just in the prompt (audit P2): the folder block
  // hides excluded folders from the AI, but the model can still NAME an
  // excluded path (from the email's content, or hallucination) — and the
  // full tree resolves it to a real targetFolderId that would execute.
  // Default [] keeps existing callers/tests source-compatible.
  excludePrefixes: string[] = [],
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
    // Excluded path → unresolved (no targetFolderId), same shape as the
    // path-not-found case: the row shows in the plan for the user to
    // re-target manually instead of silently moving into a folder they
    // explicitly excluded.
    if (pathExcluded(action.targetFolderPath, excludePrefixes)) {
      return {
        ...base,
        action: 'move',
        targetFolderPath: action.targetFolderPath,
        source: 'unresolved',
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: action.targetFolderPath,
        reason: base.reason + '（AI 指定的路徑在排除清單）',
      }
    }
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
    // A parent inside the exclude list is as unusable as a nonexistent one
    // — creating a child there would put mail into an excluded subtree.
    const parentValid =
      !parentPath ||
      (!pathExcluded(parentPath, excludePrefixes) &&
        flattenFolderTree(folderTree).some((n) => n.path === parentPath))
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
  const emailBlock = buildEmailBlock(input.emails, { threadHints: input.threadHints })
  // Few-shot mining is privacy-sensitive: user rule target paths (e.g.
  // "客戶X 離婚案") would round-trip through the LLM. Honour the
  // settings.aiIncludeFewShotExamples flag — defaults to ON because
  // examples improve accuracy meaningfully, but lawyers / others with
  // sensitive folder names can disable in Options.
  // Also drop exemplars targeting EXCLUDED folders (audit P2): the folder
  // block hides excluded paths from the AI, but an exemplar line
  // "寄件人網域 @x → 排除的資料夾" actively steered the model into naming
  // exactly those paths.
  const fewShot = Boolean(input.rules && settings.aiIncludeFewShotExamples)
  const exemplars = fewShot
    ? selectExemplars(input.rules!).filter(
        (r) => !pathExcluded(r.targetFolderPath, input.excludePrefixes),
      )
    : []
  const examplesBlock = buildExamplesBlock(exemplars)
  // Case-number → folder map (B2-B). Same privacy gate as exemplars (it
  // exposes target paths) and the same exclude filter. Built from the full
  // rule set so every known case identifier is routable, not just the
  // sampled exemplars.
  const caseMapBlock = fewShot
    ? buildCaseMapBlock(
        input.rules!.filter((r) => !pathExcluded(r.targetFolderPath, input.excludePrefixes)),
      )
    : ''
  // Recently-active folders (B2-D). NOT gated on aiIncludeFewShotExamples: it
  // exposes only folder PATHS the tool has been filing into (already visible
  // in the folder block), no rule signals. Rendered fresh (uncached) below.
  const activeFoldersBlock = buildActiveFoldersBlock(input.recentFolders, input.excludePrefixes)

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
          // Case-number → folder map (B2-B). Cached alongside the examples;
          // it's derived from the (slow-changing) rule set, not the batch.
          ...(caseMapBlock
            ? [
                {
                  type: 'text' as const,
                  text: caseMapBlock,
                  cache_control: { type: 'ephemeral' as const, ttl: '1h' as const },
                },
              ]
            : []),
          // Recently-active folders (B2-D) — FRESH (uncached): folderActivity
          // churns every batch, so caching it would thrash the prefix cache
          // and cost more than it saves. Placed just before the email batch.
          ...(activeFoldersBlock
            ? [{ type: 'text' as const, text: activeFoldersBlock }]
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

  // Drop actions whose emailIndex maps to no real email (out-of-range or
  // NaN) instead of fabricating a ghost PlanItem with emailId '' (audit
  // P3): ghosts inflated plan.length past the chunk's email count (progress
  // showed >100%), and two ghosts collide on the same empty-string React
  // key downstream. The backstop below already guarantees every REAL email
  // gets a plan entry, so dropping loses nothing.
  const plan: PlanItem[] = actions.flatMap((action) => {
    const email = input.emails[action.emailIndex]
    if (!email) {
      console.warn(`[mail-organizer] AI 回傳無效 emailIndex ${action.emailIndex}，已丟棄`)
      return []
    }
    return [actionToPlanItem(email, action, input.folderTree, input.excludePrefixes)]
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
