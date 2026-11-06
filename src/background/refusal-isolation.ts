// Refusal isolation (2026-11-06).
//
// A refusal used to stub every rep of the chunk (up to 25 emails + their
// thread siblings) with one shared reason. The stubs are `unclassified`, so
// they were never recorded in skipHistory and came back — the next batch and
// every 繼續 re-sent the same chunk and got refused again, and the banner never
// said which email caused it. With server-side fallback retrying only some
// refusal categories (and Haiku 5.5 having none), one email could take a whole
// chunk down on every batch.
//
// Now a refused subset is bisected — sequentially, inside the chunk's own lane
// — until the refused email(s) stand alone; every other subset is classified
// normally. The extra calls are bounded per chunk and per run (constants.ts).
//
// Bisecting only helps when the refusal is about some email. When it is about
// the prompt every request shares (system prompt, folder names, rule examples,
// case table), every subset is refused and the bisection is pure waste — and
// with parallel lanes it ran in up to three lanes at once, every batch
// (2026-11-06 review, K1). So before a run isolates (or blames) anything, the
// RefusalGate below needs proof the shared prompt is acceptable: a call of the
// run that was answered, or — when the first refusal comes before any answer —
// ONE probe for the whole run, carrying the shared prompt and a harmless
// synthetic email. A refused probe means "shared": nothing is isolated or
// blamed, and the suspicion is remembered across runs until a call with that
// prompt is answered (refusal-ledger.ts).

import { ClassifierError } from '@/shared/classifier'
import { refusalCategoriesText, refusalCategoryLabel } from '@/shared/refusal-category'
import type { RefusalSuspicion } from '@/shared/refusal-ledger'
import type { Email } from '@/shared/types'

export type RefusalIsolation<T, R> = {
  /** Subsets Claude answered, in call order. */
  answered: Array<{ items: T[]; result: R }>
  /** Items Claude refused on their own (single-item calls), attributable to them. */
  refused: Array<{ item: T; error: ClassifierError }>
  /**
   * Refused subsets left unsplit: the call budget ran out, or the refusal
   * could not be attributed to single items (see `attributable`).
   */
  unisolated: Array<{ items: T[]; error: ClassifierError }>
  /** Subsets whose answer could not be parsed (chunk-local, never split). */
  parseFailed: Array<{ items: T[]; error: ClassifierError }>
  /**
   * Run-fatal error from any call (retries exhausted, auth, model gone…).
   * Isolation stops there; items in none of the lists were never answered.
   */
  fatal?: unknown
  /** Cancelled / superseded mid-way: the caller drops everything. */
  aborted: boolean
  /** Claude calls made, the first one included. */
  calls: number
}

/**
 * Classify `items` with one call; on a REFUSAL split the subset in halves and
 * recurse (first half, then second), until a single refused item remains.
 * `classify` must throw ClassifierError codes as classifyBatch does (it is
 * normally classifyBatch wrapped in withClaudeRetry).
 */
export async function classifyIsolatingRefusals<T, R>(
  items: readonly T[],
  classify: (subset: T[]) => Promise<R>,
  opts: {
    /** Calls this chunk may add after its first refusal. */
    maxExtraCalls: number
    /** The run's shared allowance of extra calls; decremented per call. */
    runBudget: { remaining: number }
    /** Asked before each extra call; false leaves the subset unsplit. */
    extraCallsAllowed?: () => boolean
    /**
     * Asked once, at the first refusal — before any extra call, and before an
     * item refused on its own is reported in `refused`. False: the refusal is
     * not about these items (it is the prompt every request shares), so every
     * refused subset is left unsplit in `unisolated` instead. Normally
     * RefusalGate.verdict(); it may wait.
     */
    attributable?: (refusal: ClassifierError) => boolean | Promise<boolean>
    /** Checked before every call — cancelled / superseded runs stop spending. */
    isCancelled: () => boolean
  },
): Promise<RefusalIsolation<T, R>> {
  const out: RefusalIsolation<T, R> = {
    answered: [],
    refused: [],
    unisolated: [],
    parseFailed: [],
    aborted: false,
    calls: 0,
  }
  let extraCalls = 0
  let attribution: boolean | null = null
  const mayAttribute = async (refusal: ClassifierError): Promise<boolean> => {
    if (attribution === null) attribution = opts.attributable ? await opts.attributable(refusal) : true
    return attribution
  }

  const visit = async (subset: T[], parentRefusal: ClassifierError | null): Promise<void> => {
    if (subset.length === 0 || out.aborted || out.fatal !== undefined) return
    if (opts.isCancelled()) {
      out.aborted = true
      return
    }
    if (parentRefusal) {
      if (
        extraCalls >= opts.maxExtraCalls ||
        opts.runBudget.remaining <= 0 ||
        opts.extraCallsAllowed?.() === false
      ) {
        out.unisolated.push({ items: subset, error: parentRefusal })
        return
      }
      extraCalls++
      opts.runBudget.remaining--
    }
    out.calls++
    let result: R
    try {
      result = await classify(subset)
    } catch (e) {
      const ce = e instanceof ClassifierError ? e : null
      if (opts.isCancelled() || ce?.code === 'ABORTED') {
        out.aborted = true
        return
      }
      if (ce?.code === 'REFUSAL') {
        const attributed = await mayAttribute(ce)
        // The verdict may have waited for a probe: re-check before going on.
        if (opts.isCancelled()) {
          out.aborted = true
          return
        }
        if (!attributed) {
          out.unisolated.push({ items: subset, error: ce })
          return
        }
        if (subset.length === 1) {
          out.refused.push({ item: subset[0]!, error: ce })
          return
        }
        const mid = Math.ceil(subset.length / 2)
        await visit(subset.slice(0, mid), ce)
        await visit(subset.slice(mid), ce)
        return
      }
      if (ce?.code === 'PARSE_ERROR') {
        out.parseFailed.push({ items: subset, error: ce })
        return
      }
      out.fatal = e
      return
    }
    out.answered.push({ items: subset, result })
  }

  await visit([...items], null)
  return out
}

// ---- Run-level verdict: is the shared prompt acceptable? -----------------------

/**
 * `isolate`: some call of the run was answered — the shared prompt is fine, a
 * refusal is about the emails it carried. `shared`: a probe with the shared
 * prompt and a harmless email was refused too (or a remembered suspicion for
 * this very prompt still holds) — refusals are not about any email.
 * `unavailable`: the probe failed for another reason (cancel, rate limit…).
 */
export type RefusalVerdict = 'isolate' | 'shared' | 'unavailable'

export type RefusalGate = {
  /**
   * A call of this run was answered — parsable or not: either way the shared
   * prompt was accepted (also clears a remembered suspicion for its prompt).
   */
  noteSuccess(promptKey: string | undefined): void
  /** A call of this run ended in a refusal. */
  noteRefusal(): void
  /**
   * The run's verdict for `refusal`. At most ONE probe per run: concurrent
   * callers wait for the same one, so only one lane ever spends calls finding
   * out (and none bisects before the run has had a success). `probe` is the
   * caller's call with the run's shared prompt and sharedPromptProbeEmail():
   * resolves when answered, throws ClassifierError codes as classifyBatch
   * does; its answer is discarded. Only the first caller's probe ever runs.
   */
  verdict(refusal: ClassifierError, probe: () => Promise<void>): Promise<RefusalVerdict>
  /**
   * The scheduler must stay at one lane: a refusal was seen and no call has
   * been answered yet — a refusing response must not keep the cold-cache gate
   * open as if it had been answered.
   */
  blocksFanOut(): boolean
  /** A run-fatal error the probe hit — handed out once, to the lane that reports it. */
  takeProbeFailure(): { error: unknown } | null
  /**
   * The run concluded the shared prompt itself is refused, and no answer has
   * contradicted it since (banner hint).
   */
  sharedConcluded(): boolean
}

export function createRefusalGate(deps: {
  model: string
  /** The remembered suspicion, read once per run (null: none / unreadable). */
  remembered: RefusalSuspicion | null
  remember: (s: RefusalSuspicion) => Promise<void>
  forget: () => Promise<void>
  ttlMs: number
  now?: () => number
  /**
   * Called once, when the run first has proof the shared prompt is answered —
   * by a noted success or by the probe, a parse-failed answer included
   * (round 3): the scheduler's one-lane hold (blocksFanOut) lifts right then,
   * instead of at the next unrelated scheduling event.
   */
  onProven?: () => void
}): RefusalGate {
  const now = deps.now ?? Date.now
  let state: 'unknown' | 'proven' | 'shared' | 'unavailable' = 'unknown'
  let refusalSeen = false
  let remembered = deps.remembered
  let probing: Promise<void> | null = null
  let probeFailure: { error: unknown } | null = null

  const becomeProven = () => {
    if (state === 'proven') return
    state = 'proven'
    deps.onProven?.()
  }

  const noteSuccess = (promptKey: string | undefined) => {
    becomeProven()
    if (remembered && promptKey && remembered.promptKey === promptKey) {
      remembered = null
      void deps.forget().catch(() => {})
    }
  }

  const runProbe = async (refusal: ClassifierError, probe: () => Promise<void>): Promise<void> => {
    try {
      await probe()
      becomeProven()
    } catch (e) {
      const ce = e instanceof ClassifierError ? e : null
      if (ce?.code === 'REFUSAL') {
        // A success elsewhere in the meantime outweighs the probe.
        if (state !== 'unknown') return
        state = 'shared'
        const promptKey = ce.promptKey ?? refusal.promptKey
        if (promptKey) {
          remembered = { model: deps.model, promptKey, category: ce.category ?? null, at: now() }
          await deps.remember(remembered).catch(() => {})
        }
      } else if (ce?.code === 'PARSE_ERROR') {
        // Answered, just not parsable: the prompt itself was accepted.
        noteSuccess(ce.promptKey)
      } else {
        if (state === 'unknown') state = 'unavailable'
        if (ce?.code !== 'ABORTED') probeFailure = { error: e }
      }
    }
  }

  return {
    noteSuccess,
    noteRefusal() {
      refusalSeen = true
    },
    async verdict(refusal, probe) {
      for (;;) {
        if (state === 'proven') return 'isolate'
        if (state === 'shared') return 'shared'
        if (state === 'unavailable') return 'unavailable'
        const r = remembered
        if (
          r &&
          r.model === deps.model &&
          !!refusal.promptKey &&
          r.promptKey === refusal.promptKey &&
          now() - r.at < deps.ttlMs
        ) {
          state = 'shared'
          return 'shared'
        }
        probing ??= runProbe(refusal, probe)
        await probing
        if (state === 'unknown') state = 'unavailable' // never loop on a probe that decided nothing
      }
    },
    blocksFanOut() {
      return refusalSeen && state !== 'proven'
    },
    takeProbeFailure() {
      const f = probeFailure
      probeFailure = null
      return f
    },
    sharedConcluded() {
      return state === 'shared'
    },
  }
}

/** Id of the synthetic probe email — never a real Outlook id. */
export const SHARED_PROMPT_PROBE_ID = 'mail-organizer:shared-prompt-probe'

/**
 * The harmless email a probe carries: an everyday meeting confirmation, so a
 * refusal of the probe can only come from the prompt every request shares.
 * Never shown, never executed — its answer is discarded.
 */
export function sharedPromptProbeEmail(): Email {
  const text = '您好，想確認下週二下午兩點的會議時間是否方便，謝謝。'
  return {
    Id: SHARED_PROMPT_PROBE_ID,
    Subject: '下週會議時間確認',
    BodyPreview: text,
    bodyText: text,
    From: { EmailAddress: { Address: 'colleague@example.com', Name: '同事' } },
    ToRecipients: [],
    ReceivedDateTime: new Date().toISOString(),
    ParentFolderId: '',
  }
}

// ---- User-facing text (Traditional Chinese, neutral wording) ----------------

const SUBJECT_SHOWN_CHARS = 40

function shownSubject(subject: string | undefined): string {
  const s = (subject ?? '').replace(/\s+/g, ' ').trim()
  if (!s) return '(空主旨)'
  return s.length > SUBJECT_SHOWN_CHARS ? s.slice(0, SUBJECT_SHOWN_CHARS) + '…' : s
}

/** Plan-row reason for an email Claude just refused on its own. */
export function refusedEmailReason(subject: string | undefined, category: string | null): string {
  return `Claude 拒絕分類「${shownSubject(subject)}」（類別：${refusalCategoryLabel(category)}），請手動處理`
}

/** Plan-row reason for an email skipped because Claude refused it before. */
export function previouslyRefusedReason(subject: string | undefined, category: string | null): string {
  return `Claude 先前拒絕分類「${shownSubject(subject)}」（類別：${refusalCategoryLabel(category)}），這次未再送出，請手動處理`
}

/** Plan-row reason for an email whose whole request was refused, not pinned on it. */
export function batchRefusedReason(category: string | null): string {
  return `Claude 拒絕了這封所在的整批分類請求（類別：${refusalCategoryLabel(category)}），無法確定是哪一封；下一批會再試`
}

/**
 * The run's refusal banner, accumulated over every chunk (2026-11-06, K4):
 * the emails refused on their own (named, first three subjects) and the ones
 * whose whole request was refused without being pinned on them. `shared`: the
 * run concluded the shared prompt itself is refused — say where to look.
 * Null when nothing was refused.
 */
export function describeRefusals(r: {
  named: ReadonlyArray<{ subject: string | undefined; category: string | null }>
  unattributed: number
  unattributedCategories: ReadonlyArray<string | null>
  shared: boolean
}): string | null {
  const parts: string[] = []
  if (r.named.length > 0) {
    const shown = r.named
      .slice(0, 3)
      .map((x) => `「${shownSubject(x.subject)}」`)
      .join('、')
    parts.push(
      `Claude 拒絕分類 ${r.named.length} 封郵件：${shown}${r.named.length > 3 ? ' 等' : ''}` +
        `（類別：${refusalCategoriesText(r.named.map((x) => x.category))}），已標為未分類，請手動處理`,
    )
  }
  if (r.unattributed > 0) {
    const categories = refusalCategoriesText(r.unattributedCategories)
    const named = parts.length > 0
    parts.push(
      r.shared
        ? `${named ? '另有' : 'Claude 拒絕分類其中'} ${r.unattributed} 封郵件${named ? '被拒' : ''}` +
            `（類別：${categories}）：連只含一封測試郵件的請求也被拒絕，原因可能在資料夾名稱或規則範例，` +
            '可到設定 → 分類引擎關閉「AI 是否參考既有規則範例」後再試；已標為未分類，下一批會再試'
        : named
          ? `另有 ${r.unattributed} 封郵件所在的整批請求被拒（類別：${categories}），無法確定是哪一封，下一批會再試`
          : `Claude 拒絕了 ${r.unattributed} 封郵件所在的整批請求（類別：${categories}），無法確定是哪一封；` +
            '已標為未分類，下一批會再試',
    )
  }
  return parts.length > 0 ? parts.join('；') : null
}
