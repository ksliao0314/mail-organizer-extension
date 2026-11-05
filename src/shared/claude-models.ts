// Known Claude models this extension offers (2026-11-05).
//
// Single source of truth for the Options model picker, settings validation
// (storage.sanitizeSettings + the SW setSettings handler), the per-generation
// request parameters in classifier.ts, and the cost estimate.
//
// Deliberately dependency-free: storage.ts imports it, and importing
// classifier.ts from storage.ts would form a cycle (classifier → rules →
// storage). classifier.ts re-exports everything here, so callers can treat
// classifier.ts as the public home of the table.

export type ClaudeModelFamily = 'opus' | 'sonnet' | 'haiku'

/** USD per million tokens. */
export type ClaudeModelPrice = {
  input: number
  output: number
  /** Cache read. Mostly 0.1× input, but NOT on every model (Opus 5.5 is $0.20 = 0.05×). */
  cacheRead: number
  /** 1-hour cache write (this client marks every breakpoint ttl:'1h') = 2× input. */
  cacheWrite1h: number
}

export type ClaudeModelInfo = {
  id: string
  family: ClaudeModelFamily
  /**
   * 5 = the Claude *.5 line: thinking is always on (effort is the only
   * control), sampling params / budget_tokens / prefill are rejected, safety
   * classifiers can answer stop_reason 'refusal'. 4 = legacy 4.x request
   * surface (no thinking unless asked, no effort on Haiku 4.5).
   */
  generation: 4 | 5
  label: string
  /** Older generation kept selectable for users who pinned it. */
  legacy: boolean
  price: ClaudeModelPrice
  /**
   * Server-side refusal fallbacks (`fallbacks: "default"` under beta
   * server-side-fallback-2026-07-01). Claude API only; Haiku 5.5 has none
   * (a fallbacks list is a 400 there, "default" is a no-op).
   */
  serverFallback: boolean
  /**
   * `output_config.format` (JSON schema) is documented as supported. Models
   * missing from the documented list still work through the prompt + tolerant
   * parser, and a 400 on the format is caught and retried without it anyway.
   */
  structuredOutputs: boolean
}

export const CLAUDE_MODELS: readonly ClaudeModelInfo[] = [
  {
    id: 'claude-sonnet-5-5',
    family: 'sonnet',
    generation: 5,
    label: 'Sonnet 5.5（推薦）',
    legacy: false,
    price: { input: 2, output: 10, cacheRead: 0.2, cacheWrite1h: 4 },
    serverFallback: true,
    structuredOutputs: true,
  },
  {
    id: 'claude-haiku-5-5',
    family: 'haiku',
    generation: 5,
    label: 'Haiku 5.5（快、便宜）',
    legacy: false,
    // ≤100K-token prompt rate card (classify chunks stay well under it).
    price: { input: 0.1, output: 0.5, cacheRead: 0.01, cacheWrite1h: 0.2 },
    serverFallback: false,
    structuredOutputs: true,
  },
  {
    id: 'claude-opus-5-5',
    family: 'opus',
    generation: 5,
    label: 'Opus 5.5（最強）',
    legacy: false,
    price: { input: 4, output: 20, cacheRead: 0.2, cacheWrite1h: 8 },
    serverFallback: true,
    structuredOutputs: true,
  },
  {
    id: 'claude-sonnet-4-6',
    family: 'sonnet',
    generation: 4,
    label: 'Sonnet 4.6（舊版）',
    legacy: true,
    price: { input: 3, output: 15, cacheRead: 0.3, cacheWrite1h: 6 },
    serverFallback: false,
    structuredOutputs: false,
  },
  {
    id: 'claude-haiku-4-5-20251001',
    family: 'haiku',
    generation: 4,
    label: 'Haiku 4.5（舊版）',
    legacy: true,
    price: { input: 1, output: 5, cacheRead: 0.1, cacheWrite1h: 2 },
    serverFallback: false,
    structuredOutputs: true,
  },
  {
    id: 'claude-opus-4-7',
    family: 'opus',
    generation: 4,
    label: 'Opus 4.7（舊版）',
    legacy: true,
    price: { input: 5, output: 25, cacheRead: 0.5, cacheWrite1h: 10 },
    serverFallback: false,
    structuredOutputs: false,
  },
]

/** Default for NEW installs only — stored known ids are never migrated. */
export const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-5-5'

export function findClaudeModel(id: string | null | undefined): ClaudeModelInfo | undefined {
  if (typeof id !== 'string') return undefined
  return CLAUDE_MODELS.find((m) => m.id === id)
}

export function isKnownClaudeModel(id: unknown): id is string {
  return typeof id === 'string' && CLAUDE_MODELS.some((m) => m.id === id)
}

/** Current-generation model of a family — the price fallback for ids we don't list. */
export function familyFallbackModel(id: string | null | undefined): ClaudeModelInfo {
  const m = (id ?? '').toLowerCase()
  const family: ClaudeModelFamily = m.includes('opus')
    ? 'opus'
    : m.includes('haiku')
      ? 'haiku'
      : // Sonnet is the safe middle estimate for anything unrecognised.
        'sonnet'
  return CLAUDE_MODELS.find((x) => x.family === family && !x.legacy)!
}
