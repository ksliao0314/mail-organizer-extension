import { Fragment, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { LayoutContext, useLayout } from './layout-context'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import {
  AlertTriangle,
  CheckCircle2,
  Circle,
  FolderTree,
  HelpCircle,
  Loader2,
  Mail,
  Pause,
  Play,
  RefreshCw,
  RotateCcw,
  Search,
  Settings as SettingsIcon,
  Sparkles,
  X,
  XCircle,
} from 'lucide-react'
import { PlanRow } from './components/PlanRow'
import { FolderPicker } from './components/FolderPicker'
import { cn } from '@/lib/utils'
import { flattenFolderTree, joinFolderPath } from '@/shared/outlook-api'
import { estimateUsageCostUsd, formatUsdApprox } from '@/shared/classifier'
import { OnboardingWizard } from './components/OnboardingWizard'
import { initParentBridge, postToParent } from '@/shared/parent-bridge'
import { normalizeSubject } from '@/shared/normalize'
import type {
  ExecuteItemStatus,
  ExecuteState as SharedExecuteState,
  MailFolderNode,
  PlanAction,
  PlanItem,
  UndoSnapshot,
} from '@/shared/types'

// ---- Background protocol shims --------------------------------------------

type Ok<T> = { ok: true; data?: T }
type Err = { ok: false; code: string; message: string }
async function send<T>(req: unknown): Promise<Ok<T> | Err> {
  return chrome.runtime.sendMessage(req)
}

type StatusData = {
  owaConnected: boolean
  tokenValid: boolean
  apiKeyConfigured: boolean
  apiKeyPreview: string | null
  model: string
  folderCacheAgeMs: number | null
  folderCacheCount: number
  rulesCount: number
  excludePrefixes: string[]
  batchSize: number
  /** Pipeline mode — pre-classify next batch while user reviews DoneScreen. */
  prefetchNextBatch?: boolean
  /** Firm internal domains. Drives onboarding banner when empty. */
  internalDomains: string[]
  /** Primary case-tracking root. Drives onboarding banner when empty. */
  primaryRootPath: string
}

type ClassifyUsage = { inputTokens: number; outputTokens: number; cacheCreationTokens: number; cacheReadTokens: number }

type PreviewEmail = { id: string; subject: string; from: string; fromName: string; received: string }

type FolderActivityRow = {
  folderId: string
  folderPath: string
  lastActiveAt: string
  recentCount: number
  lastBatchAt?: string
  /**
   * Captured during refresh scan — top-1 message in the folder. Used by
   * FolderActivityListItem's title tooltip so the user can preview what's
   * latest in the case without clicking through.
   */
  latestMessage?: {
    subject: string
    from: string
    receivedAt: string
  }
  /**
   * SW annotation: true when folderId is in current folderCache; false
   * when synced in from another machine that has a folderId we don't
   * recognise (account mismatch edge case); undefined when no cache
   * exists yet (treat as "unknown, render normally" rather than greyed).
   */
  localExists?: boolean
}

// Recent Activity display filter moved to the SW (handler
// `getFolderActivity`). It reads `recentActivityIncludePrefixes` +
// `recentActivityIncludeLeafNames` from Settings so the user can edit
// the rules in options without a code change. Popup only renders.

type WeeklyDigestSummary = {
  daysSpan: number
  moved: number
  deleted: number
  foldersCreated: number
  errors: number
  rulesAdded: number
  rulesCountNow: number
  sleepingCount: number
  orphanedCount: number
  conflictsCount: number
}

type PreflightData = {
  rulePlan: PlanItem[]
  unmatchedPreview: PreviewEmail[]
  folderTree: MailFolderNode[]
  aiConfidenceThreshold: number
  /** Emails pre-filtered out because user already chose 'skip' previously. */
  preFilteredCount: number
  /** Emails skipped because they're flagged as 待處理 in Outlook. */
  flaggedCount: number
  /** Original batch size requested, before pre-filter. */
  requestedBatchSize: number
}

type AiProgress = {
  totalEmails: number
  completedEmails: number
  chunks: number
  completedChunks: number
  rulePlan: PlanItem[]
  folderTree: MailFolderNode[]
  aiPlan: PlanItem[]
  skippedByUser?: number
  usage?: ClassifyUsage
  aiError?: { code: string; message: string } | null
  startedAt: number
  done: boolean
}

// ---- State machine --------------------------------------------------------

// Popup uses a subset of the SW's ExecuteState — `plan` is intentionally
// stripped (popup never reads it; not worth holding a copy in popup memory
// for large batches). Derived via Omit so adding fields in shared/types.ts
// propagates here automatically.
type ExecuteState = Omit<SharedExecuteState, 'plan'>

type Phase =
  | { kind: 'idle' }
  | { kind: 'classifying'; startedAt: number }
  | {
      kind: 'plan'
      items: PlanItem[]
      folderTree: MailFolderNode[]
      usage: ClassifyUsage | null
      summary: { ruleHits: number; aiHandled: number }
      banner?: { code: string; message: string }
      /** When defined, AI is still classifying — execute is disabled. */
      aiPending?: {
        startedAt: number
        totalEmails: number
        completedEmails: number
        chunks: number
        completedChunks: number
      }
    }
  | { kind: 'executing'; state: ExecuteState }
  | { kind: 'done'; state: ExecuteState }
  | { kind: 'error'; code: string; message: string }

// ---- Small UI helpers -----------------------------------------------------

function StatusPill({
  ok,
  label,
  tooltip,
  pending,
}: {
  ok: boolean
  label: string
  tooltip?: string
  pending?: boolean
}) {
  return (
    <Badge variant={ok ? 'success' : pending ? 'muted' : 'warning'} className="gap-1" title={tooltip}>
      {pending ? (
        <Loader2 className="size-2.5 animate-spin" />
      ) : ok ? (
        <CheckCircle2 className="size-2.5" />
      ) : (
        <Circle className="size-2.5" />
      )}
      {label}
    </Badge>
  )
}

// Compact health indicator next to the settings icon. Green = everything
// the user needs to start classifying is in place (OWA tab + token + Claude
// API key); amber = at least one issue. The tooltip names the specific
// issue(s) so the user knows what to fix before reaching for the gear icon.
function HealthDot({
  status,
  loading,
}: {
  status: StatusData | null
  loading: boolean
}) {
  const issues: string[] = []
  if (status) {
    if (!status.owaConnected) issues.push('Outlook 未連線')
    else if (!status.tokenValid) issues.push('Outlook 登入過期')
    if (!status.apiKeyConfigured) issues.push('Claude API key 未設定')
  }
  const healthy = !!status && !loading && issues.length === 0
  const tooltip = healthy
    ? '連線、token、API key 都正常'
    : loading
      ? '檢查連線中…'
      : issues.length > 0
        ? issues.join('、')
        : '連線狀態尚未確認'
  return (
    <span
      role="status"
      aria-label={tooltip}
      title={tooltip}
      className={cn(
        'inline-block size-2.5 rounded-full',
        healthy ? 'bg-emerald-500' : 'bg-amber-500',
      )}
    />
  )
}

function planSummary(items: PlanItem[]) {
  const counts: Record<PlanAction, number> = { move: 0, delete: 0, new_folder: 0, skip: 0 }
  for (const i of items) counts[i.action]++
  return counts
}

// normalizeSubject moved to '@/shared/normalize' for reuse by SW thread
// memory pre-filter.

/**
 * Build a picker-only tree that includes both real folders AND virtual nodes
 * for every distinct `new_folder` destination still pending in this plan.
 *
 * Virtual nodes carry id `pending:<emailId>` (a sentinel — execute.ts treats
 * any non-valid id as "resolve by path at execute time"). The execute layer
 * runs new_folder items first, so by the time the dependent move executes the
 * real folder exists in the in-memory tree (spliced) and path lookup succeeds.
 */
// Banner code → 律師可讀標題（UI/UX 檢討 2026-07）：舊版把 TRUNCATED /
// PRE_FILTERED 這種機器碼直接渲染成粗體標題。原始 code 移入 title 屬性，
// 回報問題時 hover 仍查得到。
const BANNER_TITLE: Record<string, string> = {
  TRUNCATED: 'AI 回應被截斷、部分郵件未分類',
  PRE_FILTERED: '部分郵件已預先排除',
  CONFIDENCE_GATED: '部分郵件 AI 信心不足、已自動保留',
}
function bannerTitle(code: string): string {
  if (BANNER_TITLE[code]) return BANNER_TITLE[code]
  if (code.startsWith('CLASSIFIER_') || code === 'AI_FAILED') return 'AI 分類發生問題'
  return '分類時發生問題'
}

function buildAugmentedTree(tree: MailFolderNode[], items: PlanItem[]): MailFolderNode[] {
  const cloned: MailFolderNode[] = JSON.parse(JSON.stringify(tree))

  function findNode(nodes: MailFolderNode[], path: string): MailFolderNode | undefined {
    for (const n of nodes) {
      if (n.path === path) return n
      const found = findNode(n.children, path)
      if (found) return found
    }
    return undefined
  }

  const addedPaths = new Set<string>()
  for (const item of items) {
    if (item.action !== 'new_folder') continue
    const name = item.suggestedFolderName?.trim()
    const parentPath = item.suggestedParentPath?.trim()
    if (!name || !parentPath) continue

    const fullPath = joinFolderPath(parentPath, name)
    if (addedPaths.has(fullPath)) continue
    addedPaths.add(fullPath)

    const parent = findNode(cloned, parentPath)
    if (!parent) continue
    if (parent.children.some((c) => c.displayName === name)) continue

    parent.children.push({
      id: `pending:${item.emailId}`,
      displayName: name,
      path: fullPath,
      children: [],
    })
  }
  return cloned
}

// ---- App ------------------------------------------------------------------

export default function App() {
  const [phase, setPhase] = useState<Phase>({ kind: 'idle' })
  const [status, setStatus] = useState<StatusData | null>(null)
  // `?` opens a small overlay listing keyboard shortcuts. Single source of
  // truth — same overlay is reachable from any screen.
  const [showShortcuts, setShowShortcuts] = useState(false)
  const [statusLoading, setStatusLoading] = useState(true)
  const [batchSize, setBatchSize] = useState(50)
  const [folderActivity, setFolderActivity] = useState<FolderActivityRow[]>([])
  const [weeklyDigest, setWeeklyDigest] = useState<WeeklyDigestSummary | null>(null)
  // First-run wizard state. `null` = not yet fetched; `false` = checked
  // and not needed; `true` = show wizard. Loaded once on App mount.
  // The wizard only requires an API key — no folder tree needed because
  // the pipeline scans everything by default. Setting a primary root
  // path is an optional later refinement in Options.
  const [wizardNeeded, setWizardNeeded] = useState<boolean | null>(null)
  const rootRef = useRef<HTMLElement>(null)
  const [isWide, setIsWide] = useState(false)
  // Wall-clock budget for the "waiting for OWA token" polling loop
  // below. Survives status re-renders so each new refreshStatus result
  // doesn't reset the 10s ceiling.
  const tokenPollStartRef = useRef<number | null>(null)

  // Observe the main element's actual rendered width so the layout adapts
  // when the popup is loaded inside the FAB iframe (reading-pane mode).
  // Hysteresis: enter wide at 780, exit at 740 — prevents thrashing if the
  // host page width grazes the boundary (e.g. user drag-resizes the
  // browser, OWA reading pane animations).
  useEffect(() => {
    const el = rootRef.current
    if (!el) return
    const observer = new ResizeObserver(() => {
      const w = el.clientWidth
      setIsWide((prev) => (prev ? w >= 740 : w >= 780))
    })
    observer.observe(el)
    return () => observer.disconnect()
  }, [])

  // ESC closes the FAB panel when the popup is loaded inside the iframe.
  // Same cross-frame trick as options: keyboard events can't bubble out
  // of an iframe, so we postMessage the parent. Skip INPUT/TEXTAREA so
  // ESC inside a FolderPicker or rule editor still acts on the field
  // instead of nuking the whole panel.
  // No-op when running as the standalone Chrome toolbar popup
  // (window.parent === window) — Chrome handles popup close itself.
  useEffect(() => {
    if (window.parent === window) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      const t = e.target as HTMLElement | null
      const tag = t?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return
      postToParent({ type: 'mail-organizer/close-panel' })
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [])

  // Handshake with the FAB content script so postToParent learns the
  // parent's origin. No-op in toolbar mode (window.parent === window).
  useEffect(() => initParentBridge(), [])

  // Global `?` key opens the shortcut cheatsheet. Skip when typing in
  // form fields so a literal "?" in a text input doesn't pop the overlay.
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (e.metaKey || e.ctrlKey || e.altKey) return
      const t = e.target as HTMLElement | null
      const tag = t?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return
      if (e.key === '?' || (e.key === '/' && e.shiftKey)) {
        e.preventDefault()
        setShowShortcuts((v) => !v)
      } else if (e.key === 'Escape' && showShortcuts) {
        e.preventDefault()
        setShowShortcuts(false)
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [showShortcuts])

  const refreshStatus = useCallback(async () => {
    setStatusLoading(true)
    const r = await send<StatusData>({ type: 'getStatus' })
    setStatusLoading(false)
    if (r.ok && r.data) {
      setStatus(r.data)
      if (r.data.batchSize) setBatchSize(r.data.batchSize)
    }
  }, [])

  const refreshFolderActivity = useCallback(async () => {
    // SW now applies the include-prefixes / leaf-names filter (configurable
    // in options) and returns only the rows the user wants to see.
    const r = await send<{ activity: FolderActivityRow[] }>({ type: 'getFolderActivity' })
    if (r.ok && r.data) {
      setFolderActivity(r.data.activity)
    }
  }, [])

  const refreshWeeklyDigest = useCallback(async () => {
    const r = await send<
      | { shouldShow: false }
      | ({ shouldShow: true } & WeeklyDigestSummary)
    >({ type: 'getWeeklyDigest' })
    if (r.ok && r.data && r.data.shouldShow) {
      const { shouldShow: _ignored, ...summary } = r.data
      setWeeklyDigest(summary)
    } else {
      setWeeklyDigest(null)
    }
  }, [])

  // Pull the digest once at startup. Re-pulling later would just re-show
  // the card the user dismissed, so once-per-popup-open is enough.
  useEffect(() => {
    void refreshWeeklyDigest()
  }, [refreshWeeklyDigest])

  // First-run wizard probe. Asks the SW whether onboarding is needed —
  // SW knows the full picture (API key, cloud state, local backups) so
  // the wizard can pick the right branch (truly new vs "sync hasn't
  // pulled yet"). Runs once per popup open.
  useEffect(() => {
    void (async () => {
      const r = await send<{ needed: boolean }>({ type: 'getOnboardingState' })
      setWizardNeeded(r.ok && r.data ? r.data.needed : false)
    })()
  }, [])

  // Option B — Pipeline prewarm of the OWA sidebar. As soon as we know
  // which folders will appear in 「近日活動」, ship the list to the FAB
  // content script so it can pre-expand each ancestor chain in OWA.
  // When the user later clicks a row, navigation completes in ~100ms
  // instead of the 5-15s OWA needs to lazy-load 50-70+ child folders.
  //
  // We trigger only once per popup session (the ref). Re-runs would be
  // free thanks to expandOwaTreeitem's idempotent guard, but skipping
  // them avoids spamming postMessage. Also no-op in toolbar mode where
  // there's no parent OWA tab to drive.
  const prewarmSentRef = useRef(false)
  useEffect(() => {
    if (prewarmSentRef.current) return
    if (folderActivity.length === 0) return
    if (window.parent === window) return
    prewarmSentRef.current = true
    const paths = folderActivity
      .map((row) => row.folderPath)
      .filter((p): p is string => typeof p === 'string' && p.length > 0)
    if (paths.length > 0) {
      postToParent({ type: 'mail-organizer/prewarm-sidebar', paths })
    }
  }, [folderActivity])

  async function dismissWeeklyDigest() {
    setWeeklyDigest(null)
    await send({ type: 'dismissWeeklyDigest' })
  }

  // Auto-refresh the recent-activity list whenever the background updates
  // chrome.storage.local['folderActivity'] — fires after every execute batch
  // via recordFolderActivityFromBatch, so the user sees the just-archived
  // folders appear without having to close+reopen the popup.
  useEffect(() => {
    const onChanged = (
      changes: { [key: string]: chrome.storage.StorageChange },
      area: chrome.storage.AreaName,
    ) => {
      if (area !== 'local') return
      // Re-fetch when folderActivity changes (execute wrote new data, or
      // a sync pull merged in another machine's activity) OR when settings
      // change (filter rules may have changed, the same activity now
      // resolves to a different visible list).
      if (changes.folderActivity || changes.settings) {
        void refreshFolderActivity()
      }
    }
    chrome.storage.onChanged.addListener(onChanged)
    return () => chrome.storage.onChanged.removeListener(onChanged)
  }, [refreshFolderActivity])

  // Cross-machine sync use case: when popup opens on machine B, we may
  // have just pulled folderActivity from machine A — those entries'
  // latestMessage snapshots reflect A's view at A's push time, which
  // could be hours stale by now. Auto-trigger a Graph API refresh if it
  // has been > 30 min since the last successful refresh on THIS machine,
  // so the lawyer's "quick browse" actually shows fresh subjects.
  //
  // Gated:
  //   - only when status loaded (we need OWA connected + folderCache)
  //   - only when in idle phase (don't pull Graph during active classify)
  //   - run once per popup session (ref guard)
  //   - skip if last refresh was within 30 min (covers the case where
  //     user is bouncing in and out of the popup all morning)
  const autoRefreshDoneRef = useRef(false)
  useEffect(() => {
    if (autoRefreshDoneRef.current) return
    if (!status?.owaConnected) return
    if (status.folderCacheCount === 0) return
    if (phase.kind !== 'idle') return
    autoRefreshDoneRef.current = true
    void (async () => {
      try {
        const f = await send<{ lastRefreshAt: string | null }>({
          type: 'getFolderActivityFreshness',
        })
        const last = f.ok ? f.data?.lastRefreshAt : null
        const STALE_MS = 30 * 60 * 1000
        if (last && Date.now() - new Date(last).getTime() < STALE_MS) return
        // Skip if user hasn't configured a root path yet — refresh would
        // hit NO_ROOT_CONFIGURED and surface as an error.
        if (!status.primaryRootPath) return
        const r = await send({ type: 'refreshFolderActivity' })
        if (r.ok) await refreshFolderActivity()
      } catch (e) {
        // Best-effort — silent on failure. The user can still hit the
        // refresh button manually.
        console.debug('[mail-organizer] auto-refresh folderActivity skipped', e)
      }
    })()
  }, [status, phase.kind, refreshFolderActivity])

  // Pipeline mode (Feat 1) — when the user lands on DoneScreen and has
  // opted in to prefetch, silently kick off classify for the next batch.
  // The SW writes results to aiClassifyProgress as usual; the popup mount
  // logic stays on DoneScreen because executeState.finishedAt outranks
  // aiProgress. When the user clicks 繼續歸檔下一批, onContinue checks
  // for a finished prefetch first and skips the slow fetch+classify chain.
  // Each DoneScreen visit triggers at most once via `prefetchStartedRef`.
  const prefetchStartedRef = useRef<string | null>(null)
  useEffect(() => {
    if (phase.kind !== 'done') {
      prefetchStartedRef.current = null
      return
    }
    if (!status?.prefetchNextBatch || !status.owaConnected || !status.tokenValid || !status.apiKeyConfigured) {
      return
    }
    // Don't re-trigger for the same DoneScreen instance.
    const phaseKey = String(phase.state.startedAt)
    if (prefetchStartedRef.current === phaseKey) return
    prefetchStartedRef.current = phaseKey
    // Capture the just-finished execute's start time while phase is narrowed
    // to 'done' — the freshness watermark for the staleness check below.
    const execStartedAt = phase.state.startedAt

    void (async () => {
      // Cancellation (audit P2): the ref is reset to null the moment phase
      // leaves 'done' (see the effect's early-return above), so re-checking
      // it after EVERY await makes leaving DoneScreen an effective cancel.
      // Without this, a quick 繼續歸檔下一批 raced the still-in-flight prefetch
      // chain into a SECOND classifyPreflight+classifyAi (double Anthropic
      // spend, interleaved progress writes); 回主畫面 similarly let the chain
      // resurrect AI progress that finishAndReset had just cleared.
      const cancelled = () => prefetchStartedRef.current !== phaseKey
      // Audit: the just-executed batch leaves ITS OWN (older) AI classify
      // progress in session storage — nothing clears it between classify-done
      // and DoneScreen mount. The old guard bailed on ANY progress, so it
      // mistook that stale residue for a live prefetch and never started one
      // → pipeline mode was a silent no-op on every AI batch. Only bail when
      // the progress is strictly NEWER than the execute we just finished (a
      // genuine in-flight/done prefetch); otherwise clear the residue and
      // proceed. Mirrors continueToNextBatch's startedAt staleness guard.
      const existing = await send<AiProgress | null>({ type: 'getAiClassifyProgress' })
      if (cancelled()) return
      if (
        existing.ok &&
        existing.data &&
        (existing.data.done === false ||
          (typeof existing.data.startedAt === 'number' &&
            existing.data.startedAt > execStartedAt))
      ) {
        // 活的 classify（進行中，可能是部分執行留下的）或真正的
        // prefetch — 都不要動它。殘渣必然 done:true 且早於 execute。
        return
      }
      if (existing.ok && existing.data) {
        // Stale residue from the batch we just executed — clear it so the
        // resume path can't promote it and so our prefetch starts clean.
        await send({ type: 'clearAiClassifyProgress' })
        if (cancelled()) return
      }
      // Headless preflight + classifyAi. Same SW handlers the foreground
      // path uses; just don't touch popup phase so DoneScreen stays put.
      const pf = await send<PreflightData>({
        type: 'classifyPreflight',
        batchSize,
        forceFresh: false,
      })
      if (cancelled()) return
      if (!pf.ok || !pf.data) return
      if (pf.data.unmatchedPreview.length === 0) return // nothing for AI to do
      await send({ type: 'classifyAi', excludeIds: [] })
    })().catch((e) => console.warn('[mail-organizer] prefetch failed (non-fatal)', e))
  }, [phase, status, batchSize])

  useEffect(() => {
    void (async () => {
      await refreshStatus()
      void refreshFolderActivity()
      // Resume order matters:
      //   1. execute in progress / finished → executing / done view
      //   2. AI classify in progress / freshly done → plan view with progress
      //   3. otherwise idle
      const execResp = await send<{ state: ExecuteState | null }>({ type: 'getExecuteState' })
      if (execResp.ok && execResp.data?.state) {
        const s = execResp.data.state
        if (s.inProgress) {
          setPhase({ kind: 'executing', state: s })
          return
        }
        if (s.finishedAt) {
          setPhase({ kind: 'done', state: s })
          return
        }
      }
      // Fetch popup state + AI progress together so we can merge: user edits
      // (in popupState) are the source of truth, with new AI items not yet
      // saved coming from aiClassifyProgress.
      const [stateResp, aiResp] = await Promise.all([
        send<{
          phase: 'preview' | 'plan'
          updatedAt: number
          preview?: {
            rulePlan: PlanItem[]
            unmatched: PreviewEmail[]
            folderTree: MailFolderNode[]
            aiConfidenceThreshold: number
            excludedIds: string[]
            preFilteredCount?: number
            flaggedCount?: number
            requestedBatchSize?: number
          }
          plan?: {
            items: PlanItem[]
            folderTree?: MailFolderNode[]
            usage: ClassifyUsage | null
            summary: { ruleHits: number; aiHandled: number }
            banner: { code: string; message: string } | null
            aiPendingFlag?: boolean
          }
        } | null>({ type: 'getPopupState' }),
        send<AiProgress | null>({ type: 'getAiClassifyProgress' }),
      ])

      const saved = stateResp.ok ? stateResp.data : null
      const aiProg = aiResp.ok ? aiResp.data : null

      // Preview is no longer a phase. Legacy snapshots from prior versions
      // can still arrive — drop them silently so we fall through to idle.
      if (saved?.phase === 'preview') {
        void send({ type: 'clearPopupState' })
      }

      // Plan phase: combine popupState items (with edits) + aiClassifyProgress
      // (aiPending status + any new AI items not yet seen).
      const hasPopupPlan = saved?.phase === 'plan' && saved.plan
      const hasAi = aiProg != null

      if (!hasPopupPlan && !hasAi) {
        // No saved state at all → idle (let mount fall through)
        return
      }

      let items: PlanItem[]
      let folderTree: MailFolderNode[]
      let usage: ClassifyUsage | null
      let summary: { ruleHits: number; aiHandled: number }
      let banner: { code: string; message: string } | undefined

      if (hasPopupPlan) {
        items = saved!.plan!.items
        // Older snapshots persisted folderTree inline; new ones don't.
        // Prefer the inline tree when present (cheap), fall back to a fresh
        // fetch from the SW's folderCache otherwise.
        if (saved!.plan!.folderTree && saved!.plan!.folderTree.length > 0) {
          folderTree = saved!.plan!.folderTree
        } else {
          const treeResp = await send<{ tree: MailFolderNode[] }>({ type: 'getFolderTree' })
          folderTree = treeResp.ok && treeResp.data ? treeResp.data.tree : []
        }
        usage = saved!.plan!.usage
        summary = saved!.plan!.summary
        banner = saved!.plan!.banner ?? undefined
        // popupState may be missing AI items that arrived after its last
        // debounced save — merge them in whether or not classify has since
        // finished. (Audit P2: this was gated on !aiProg.done, so a classify
        // that COMPLETED while the popup was closed had its post-close items
        // silently dropped on reopen — despite the aiPending card promising
        // "popup 可關、回來會自動接著看".)
        if (aiProg) {
          const seen = new Set(items.map((i) => i.emailId))
          const fresh = aiProg.aiPlan.filter((ai) => !seen.has(ai.emailId))
          if (fresh.length > 0) items = [...items, ...fresh]
          if (aiProg.done) {
            // Adopt the final run's usage / summary / error banner — the
            // stale snapshot predates completion. Same formulas as the
            // in-popup poll's done-branch below.
            usage = aiProg.usage ?? usage
            summary = {
              ruleHits: aiProg.rulePlan.length,
              aiHandled:
                aiProg.aiPlan.length -
                (aiProg.skippedByUser ?? 0) -
                aiProg.aiPlan.filter((x) => x.source === 'unresolved').length,
            }
            if (aiProg.aiError) {
              banner = {
                code: aiProg.aiError.code,
                message:
                  aiProg.aiError.code === 'TRUNCATED' || aiProg.aiError.code === 'CONFIDENCE_GATED'
                    ? aiProg.aiError.message
                    : `AI 分類失敗：${aiProg.aiError.message}。規則命中與已完成部分仍可執行。`,
              }
            }
          }
        }
      } else if (aiProg) {
        items = [...aiProg.rulePlan, ...aiProg.aiPlan]
        folderTree = aiProg.folderTree
        usage = aiProg.usage ?? null
        summary = {
          ruleHits: aiProg.rulePlan.length,
          aiHandled: aiProg.aiPlan.length - (aiProg.skippedByUser ?? 0),
        }
        banner = aiProg.aiError
          ? { code: aiProg.aiError.code, message: aiProg.aiError.message }
          : undefined
      } else {
        return
      }

      const aiPending =
        aiProg && !aiProg.done
          ? {
              startedAt: aiProg.startedAt,
              totalEmails: aiProg.totalEmails,
              completedEmails: aiProg.completedEmails,
              chunks: aiProg.chunks,
              completedChunks: aiProg.completedChunks,
            }
          : undefined

      setPhase({
        kind: 'plan',
        items,
        folderTree,
        usage,
        summary,
        banner,
        aiPending,
      })
    })()
  }, [refreshStatus])

  // Auto-poll status while we're stuck on "OWA connected but no token
  // yet". Without this, the SW's silent background getOwaToken refresh
  // (kicked off by getStatus when peekCachedToken misses) succeeds in
  // 1-2s but the popup never finds out — the user has to close + reopen
  // the popup to see a fresh status. getStatus deliberately doesn't
  // block on the refresh because askContentScript can take ~1.5s on
  // OWA cold-start, but we still need a feedback loop.
  //
  // Cadence: poll once per second, cap total wait at 10s. If still
  // stuck after that, user can click the refresh button (or close +
  // reopen as before). This is intentionally short — most successful
  // refreshes land in the first 2-3 ticks; longer waits typically
  // mean MSAL itself is doing a silent token rotation and we won't
  // help by hammering the SW.
  useEffect(() => {
    const waiting = !!status?.owaConnected && !status.tokenValid
    if (!waiting) {
      tokenPollStartRef.current = null
      return
    }
    if (tokenPollStartRef.current === null) {
      tokenPollStartRef.current = Date.now()
    }
    if (Date.now() - tokenPollStartRef.current > 10_000) return
    const id = window.setTimeout(() => {
      void refreshStatus()
    }, 1000)
    return () => clearTimeout(id)
  }, [status?.owaConnected, status?.tokenValid, refreshStatus])

  // Refresh status when the popup regains visibility / focus. Mainly
  // helps the OWA FAB iframe case where the popup persists across tab
  // switches — user pops out to Outlook, does something, comes back,
  // expects to see fresh state. The toolbar popup closes on focus
  // loss so these listeners are a no-op there.
  useEffect(() => {
    const onVisible = () => {
      if (document.visibilityState === 'visible') {
        void refreshStatus()
      }
    }
    document.addEventListener('visibilitychange', onVisible)
    window.addEventListener('focus', onVisible)
    return () => {
      document.removeEventListener('visibilitychange', onVisible)
      window.removeEventListener('focus', onVisible)
    }
  }, [refreshStatus])

  // Poll execute state every 500ms while running.
  useEffect(() => {
    if (phase.kind !== 'executing') return
    const id = setInterval(async () => {
      const r = await send<{ state: ExecuteState | null }>({ type: 'getExecuteState' })
      if (!r.ok || !r.data?.state) return
      const s = r.data.state
      if (s.inProgress) setPhase({ kind: 'executing', state: s })
      else setPhase({ kind: 'done', state: s })
    }, 500)
    return () => clearInterval(id)
  }, [phase.kind])

  // Persist preview / plan phase to chrome.storage.session so an accidental
  // popup close doesn't wipe in-progress edits. Debounced 400ms.
  // We DO save during aiPending now — user edits while AI is still chunk-
  // classifying should also survive a popup close. Resume logic on mount
  // prefers popupState items (with edits) and overlays new AI items from
  // progress that aren't yet in popupState.
  useEffect(() => {
    if (phase.kind !== 'plan') return

    const id = setTimeout(() => {
      const snapshot = {
        phase: 'plan' as const,
        updatedAt: Date.now(),
        plan: {
          items: phase.items,
          // folderTree intentionally NOT included — it can balloon to
          // hundreds of KB and trip chrome.storage.session quota when paired
          // with a large plan. Resume fetches a fresh tree from the SW.
          usage: phase.usage,
          summary: phase.summary,
          banner: phase.banner ?? null,
          // aiPending hint only — actual progress numbers come from
          // aiClassifyProgress on resume.
          aiPendingFlag: !!phase.aiPending,
        },
      }
      void send({ type: 'savePopupState', state: snapshot })
    }, 400)
    return () => clearTimeout(id)
  }, [phase])

  // Poll AI classify progress when we're in plan phase with aiPending set.
  useEffect(() => {
    if (phase.kind !== 'plan' || !phase.aiPending) return
    const id = setInterval(async () => {
      const r = await send<AiProgress | null>({ type: 'getAiClassifyProgress' })
      if (!r.ok || !r.data) return
      const prog = r.data
      setPhase((p) => {
        if (p.kind !== 'plan' || !p.aiPending) return p

        // MERGE rather than replace. Previously we did
        //   items = [...prog.rulePlan, ...prog.aiPlan]
        // which clobbered any in-popup user edits every 500ms — making
        // rule-matched items snap back to their original action / target.
        // Keep `p.items` as the source of truth for items we already know
        // about; only append AI items the SW has produced since last tick.
        const existingIds = new Set(p.items.map((i) => i.emailId))
        const freshAi = prog.aiPlan.filter((ai) => !existingIds.has(ai.emailId))
        const items = freshAi.length > 0 ? [...p.items, ...freshAi] : p.items

        if (prog.done) {
          return {
            ...p,
            items,
            usage: prog.usage ?? null,
            summary: {
              ruleHits: prog.rulePlan.length,
              aiHandled:
                prog.aiPlan.length -
                (prog.skippedByUser ?? 0) -
                prog.aiPlan.filter((x) => x.source === 'unresolved').length,
            },
            aiPending: undefined,
            banner: prog.aiError
              ? {
                  code: prog.aiError.code,
                  message:
                    prog.aiError.code === 'TRUNCATED' || prog.aiError.code === 'CONFIDENCE_GATED'
                      ? prog.aiError.message
                      : `AI 分類失敗：${prog.aiError.message}。規則命中與已完成部分仍可執行。`,
                }
              : p.banner,
          }
        }
        // Nothing changed this tick (the common case — chunks complete every
        // 10-30s but we poll at 500ms): return the SAME reference so React
        // skips the re-render entirely. Without this, every tick minted a new
        // phase object → full 100-row PlanRow re-render twice a second AND
        // re-armed the 400ms-debounced savePopupState (debounce < poll
        // interval → never coalesces → ~2 session writes/sec for the whole
        // classify).
        if (
          freshAi.length === 0 &&
          p.aiPending.completedEmails === prog.completedEmails &&
          p.aiPending.completedChunks === prog.completedChunks
        ) {
          return p
        }
        return {
          ...p,
          items,
          aiPending: {
            ...p.aiPending,
            completedEmails: prog.completedEmails,
            completedChunks: prog.completedChunks,
          },
        }
      })
    }, 500)
    return () => clearInterval(id)
  }, [phase.kind, phase.kind === 'plan' && phase.aiPending != null])

  // Pipeline mode (Feat 1) — when the DoneScreen prefetch effect already
  // kicked off classify in the background, consume that progress instead
  // of starting a fresh classify. Saves the 30-60s the user would
  // otherwise wait. Falls back to the normal flow if no prefetch exists.
  async function continueToNextBatch() {
    // Capture the just-finished execute's start time BEFORE clearing its
    // state — this is the freshness watermark for the staleness guard
    // below. We're called from the DoneScreen so phase is 'done'.
    const executeStartedAt =
      phase.kind === 'done' ? phase.state.startedAt : null

    // Don't clear AI progress yet — we might need to consume a prefetched
    // result. Clear other per-batch state safely.
    void send({ type: 'clearPopupState' })
    void send({ type: 'clearExecuteState' })
    void send({ type: 'dismissUndo' })

    const aiResp = await send<AiProgress | null>({ type: 'getAiClassifyProgress' })
    const prefetched = aiResp.ok ? aiResp.data : null

    // STALENESS GUARD (2026-06-03 — root cause of「已歸檔的信混進下一批」).
    // Session storage still holds the AI progress of the batch we JUST
    // executed: nothing clears it between classify-done and here
    // (confirmExecute only clears popupState; preflight only clears it
    // when the NEXT classify starts). Without this check, prefetch-
    // disabled users (the default) get that residue promoted as the
    // "next batch" — the exact same plan whose ids are all dead after
    // the move, so every row 404s「訊息已不在原位置」. The earlier
    // recentlyProcessed-ledger fix never fires on this path because no
    // inbox fetch happens at all.
    //
    // Discriminator: the 執行 button is disabled while aiPending, so the
    // executed cycle's classify ALWAYS finished (last progress write
    // included) before execute began → its startedAt < execute's
    // startedAt. A genuine DoneScreen prefetch only starts after the
    // execute finished → its startedAt > execute's startedAt. Strictly
    // newer ⇔ genuine prefetch; anything else gets discarded and we
    // re-classify from a fresh inbox fetch (where the ledger filter
    // handles Outlook's listing lag).
    // 刀3 (2026-07)：done === false 的 progress 一律視為「活的」。殘渣
    // 必然是 done:true（見上面的推導）— 但部分執行（AI 分類中先執行規則
    // 命中）打破了「執行前 classify 必已完成」的前提：執行後 classify
    // 仍在跑、其 startedAt 早於 execute，卻是貨真價實的進行中工作，
    // 不能被當殘渣清掉。
    const isFreshPrefetch =
      prefetched != null &&
      (prefetched.done === false ||
        (executeStartedAt != null &&
          typeof prefetched.startedAt === 'number' &&
          prefetched.startedAt > executeStartedAt))
    if (!prefetched || !isFreshPrefetch) {
      // No prefetch available (or stale residue) — clean start.
      await send({ type: 'clearAiClassifyProgress' })
      void startClassify(false)
      return
    }
    // Promote the prefetched result into a plan phase. If AI is still
    // running we land in plan with aiPending so progress keeps polling.
    //
    // 刀3：部分執行後，progress 裡仍留著「剛剛已執行那批」的 plan 項 —
    // 把已終局處理（移動/刪除/建夾/保留）的 email 濾掉，只讓「還沒處理
    // 的」進入新 plan。error/cancelled 保留 — 那些信還在收件匣，該給
    // 使用者重試的機會。
    const executedIds = new Set(
      phase.kind === 'done'
        ? phase.state.results
            .filter(
              (r) =>
                r.status === 'moved' ||
                r.status === 'deleted' ||
                r.status === 'folder_created' ||
                r.status === 'skipped',
            )
            .map((r) => r.emailId)
        : [],
    )
    const items: PlanItem[] = [...prefetched.rulePlan, ...prefetched.aiPlan].filter(
      (i) => !executedIds.has(i.emailId),
    )
    if (items.length === 0 && prefetched.done) {
      // Edge case: prefetch completed but inbox was empty — fall through.
      await send({ type: 'clearAiClassifyProgress' })
      void startClassify(false)
      return
    }
    setPhase({
      kind: 'plan',
      items,
      folderTree: prefetched.folderTree,
      usage: prefetched.usage ?? null,
      summary: {
        ruleHits: prefetched.rulePlan.length,
        aiHandled: prefetched.aiPlan.length - (prefetched.skippedByUser ?? 0),
      },
      banner: prefetched.aiError
        ? { code: prefetched.aiError.code, message: prefetched.aiError.message }
        : undefined,
      aiPending: prefetched.done
        ? undefined
        : {
            startedAt: prefetched.startedAt ?? Date.now(),
            totalEmails: prefetched.totalEmails,
            completedEmails: prefetched.aiPlan.length,
            chunks: prefetched.chunks,
            completedChunks: prefetched.completedChunks,
          },
    })
    void refreshStatus()
  }

  async function startClassify(forceFresh = false) {
    // Fresh classify replaces any stale popup state from a prior session.
    void send({ type: 'clearPopupState' })
    setPhase({ kind: 'classifying', startedAt: Date.now() })
    const r = await send<PreflightData>({ type: 'classifyPreflight', batchSize, forceFresh })
    if (!r.ok) {
      setPhase({ kind: 'error', code: r.code, message: r.message })
      return
    }
    if (!r.data) {
      setPhase({ kind: 'error', code: 'EMPTY', message: '背景沒回傳資料' })
      return
    }
    const data = r.data

    // Build a banner describing what got pre-filtered, if anything.
    const parts: string[] = []
    if (data.preFilteredCount > 0) parts.push(`${data.preFilteredCount} 件先前選擇保留`)
    if (data.flaggedCount > 0) parts.push(`${data.flaggedCount} 件標記待處理`)
    const banner =
      parts.length > 0
        ? {
            code: 'PRE_FILTERED' as const,
            message: `已從收件夾排除 ${parts.join(' + ')}。可到設定頁調整略過機制。`,
          }
        : undefined

    // No unmatched → no AI needed, go straight to plan.
    if (data.unmatchedPreview.length === 0) {
      setPhase({
        kind: 'plan',
        items: data.rulePlan,
        folderTree: data.folderTree,
        usage: null,
        summary: { ruleHits: data.rulePlan.length, aiHandled: 0 },
        banner,
      })
      void refreshStatus()
      return
    }

    // Auto-send everything unmatched to AI — the lawyer does not want to
    // pre-screen 50 emails before AI runs. Plan opens immediately showing
    // rule hits so the user can review while AI works in background.
    const aiResp = await send<{ started: boolean; chunks: number; totalEmails: number }>({
      type: 'classifyAi',
      excludeIds: [],
    })
    if (!aiResp.ok) {
      setPhase({ kind: 'error', code: aiResp.code, message: aiResp.message })
      return
    }
    setPhase({
      kind: 'plan',
      items: data.rulePlan,
      folderTree: data.folderTree,
      usage: null,
      summary: { ruleHits: data.rulePlan.length, aiHandled: 0 },
      banner,
      aiPending: {
        startedAt: Date.now(),
        totalEmails: aiResp.data?.totalEmails ?? data.unmatchedPreview.length,
        completedEmails: 0,
        chunks: aiResp.data?.chunks ?? 0,
        completedChunks: 0,
      },
    })
  }

  // Synchronous re-entry guard for 執行 (audit P2): a double-click used to
  // fire startExecute twice — the second either hit the SW's ALREADY_RUNNING
  // guard and stranded the popup on ErrorScreen, or (response slow) started
  // duplicate work. The ref flips before the first await, so click #2 is a
  // no-op regardless of message timing.
  const executeSubmittingRef = useRef(false)
  async function confirmExecute() {
    if (phase.kind !== 'plan') return
    if (executeSubmittingRef.current) return
    executeSubmittingRef.current = true
    // Bg holds the tree in folderCache already; no need to ship ~200 KB back.
    let r: Awaited<ReturnType<typeof send<{ started: boolean; total: number }>>>
    try {
      r = await send<{ started: boolean; total: number }>({
        type: 'startExecute',
        plan: phase.items,
      })
    } finally {
      executeSubmittingRef.current = false
    }
    if (!r.ok && r.code !== 'ALREADY_RUNNING') {
      setPhase({ kind: 'error', code: r.code, message: r.message })
      return
    }
    // ALREADY_RUNNING falls through: the batch we just tried to start IS
    // already running (our own double-send or another popup surface) — show
    // the executing view and let the getExecuteState poll adopt real state
    // instead of stranding the user on an error screen.
    // Execute now owns the workflow — popup state becomes stale, drop it.
    void send({ type: 'clearPopupState' })
    // Seed an initial executing state immediately so UI doesn't flicker.
    setPhase({
      kind: 'executing',
      state: {
        inProgress: true,
        cancelRequested: false,
        startedAt: Date.now(),
        total: phase.items.length,
        current: 0,
        results: phase.items.map((p) => ({
          emailId: p.emailId,
          subject: p.emailSubject,
          action: p.action,
          status: 'queued',
        })),
        summary: { moved: 0, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
        rulesAdded: 0,
      },
    })
  }

  async function cancelExecute() {
    await send({ type: 'cancelExecute' })
  }

  async function retryFailed() {
    const r = await send<{ started: number }>({ type: 'retryFailed' })
    if (!r.ok) {
      setPhase({ kind: 'error', code: r.code, message: r.message })
      return
    }
    // Seed a fresh executing state — the actual content will be replaced by polling.
    setPhase({
      kind: 'executing',
      state: {
        inProgress: true,
        cancelRequested: false,
        startedAt: Date.now(),
        total: r.data?.started ?? 0,
        current: 0,
        results: [],
        summary: { moved: 0, deleted: 0, foldersCreated: 0, skipped: 0, cancelled: 0, errors: 0 },
        rulesAdded: 0,
      },
    })
  }

  async function finishAndReset() {
    await Promise.all([
      send({ type: 'clearExecuteState' }),
      send({ type: 'clearPopupState' }),
      send({ type: 'clearAiClassifyProgress' }),
    ])
    setPhase({ kind: 'idle' })
    void refreshStatus()
  }

  function backToIdle() {
    void send({ type: 'clearPopupState' })
    setPhase({ kind: 'idle' })
  }

  function updateItem(next: PlanItem) {
    setPhase((p) => {
      if (p.kind !== 'plan') return p
      return { ...p, items: p.items.map((i) => (i.emailId === next.emailId ? next : i)) }
    })
  }

  function bulkApply(ids: Set<string>, transform: (item: PlanItem) => PlanItem) {
    setPhase((p) => {
      if (p.kind !== 'plan') return p
      return {
        ...p,
        items: p.items.map((i) => (ids.has(i.emailId) ? transform(i) : i)),
      }
    })
  }

  function openOptions() {
    // When the popup is loaded inside the OWA FAB iframe, asking Chrome to
    // open the options page would spawn a new tab — jarring because the user
    // is in OWA mid-flow. Instead, post a message to the parent so the FAB
    // content script can overlay a big options iframe on top of OWA.
    // Detected via `window.parent !== window` which is true only when we're
    // embedded; the regular toolbar popup is a top-level browser popup.
    if (window.parent !== window) {
      postToParent({ type: 'mail-organizer/open-options' })
    } else {
      chrome.runtime.openOptionsPage()
    }
  }

  const canStart =
    status?.owaConnected && status?.tokenValid && status?.apiKeyConfigured && phase.kind === 'idle'

  return (
    <LayoutContext.Provider value={{ isWide }}>
    <main
      ref={rootRef}
      className="w-full min-w-[540px] max-w-[1600px] min-h-[520px] max-h-[100vh] overflow-y-auto bg-background text-foreground font-sans"
    >
      {/* Header */}
      <header className="sticky top-0 z-10 border-b border-border bg-background/95 backdrop-blur px-4 py-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Mail className="size-4 text-foreground" />
            <h1 className="text-sm font-semibold tracking-tight">Mail Organizer</h1>
          </div>
          <div className="flex items-center gap-2">
            <PhaseContext phase={phase} model={status?.model} />
            <HealthDot status={status} loading={statusLoading} />
            <Button
              variant="ghost"
              size="icon"
              onClick={() => setShowShortcuts(true)}
              title="鍵盤捷徑 (?)"
            >
              <HelpCircle className="size-4" />
            </Button>
            <Button variant="ghost" size="icon" onClick={openOptions} title="設定">
              <SettingsIcon className="size-4" />
            </Button>
          </div>
        </div>

        {(() => {
          // Show status chips only when there's something the user needs
          // to know about: loading-in-progress, or an actual problem.
          // Healthy steady-state hides the row entirely — the IdleScreen
          // also surfaces issues as an amber banner, so chips here are
          // redundant noise for daily use.
          const owaIssue = status && !status.owaConnected
          const tokenIssue = status && status.owaConnected && !status.tokenValid
          const keyIssue = status && !status.apiKeyConfigured
          const hasIssue = !!(owaIssue || tokenIssue || keyIssue)
          if (!statusLoading && !hasIssue) return null
          return (
            <div className="mt-2 flex flex-wrap gap-1.5">
              {statusLoading && (
                <StatusPill ok={false} label="檢查連線中" pending />
              )}
              {owaIssue && (
                <StatusPill ok={false} label="未連線 Outlook" tooltip="請打開 Outlook 網頁分頁並登入" />
              )}
              {tokenIssue && (
                <StatusPill ok={false} label="Outlook 登入過期" tooltip="OWA 可能剛開、稍等幾秒再試" />
              )}
              {keyIssue && (
                <StatusPill ok={false} label="缺 Claude API key" tooltip="請到設定頁設定 API key" />
              )}
            </div>
          )
        })()}
      </header>

      {/* Body */}
      <div className="p-4">
        {phase.kind === 'idle' &&
          (wizardNeeded === true ? (
            // First-run wizard takes over the idle area. Distinguishes
            // truly-new users from "another machine pushed, sync hasn't
            // pulled yet" so we don't accidentally have them set up
            // from scratch and collide with the eventual pull.
            <OnboardingWizard
              onComplete={async () => {
                setWizardNeeded(false)
                await refreshStatus() // re-derive owaConnected / canStart
              }}
            />
          ) : (
            <IdleScreen
              status={status}
              batchSize={batchSize}
              canStart={!!canStart}
              onStart={() => void startClassify(false)}
              onRefreshTree={() => void startClassify(true)}
              onOpenOptions={openOptions}
              folderActivity={folderActivity}
              onRefreshActivity={refreshFolderActivity}
              weeklyDigest={weeklyDigest}
              onDismissDigest={() => void dismissWeeklyDigest()}
            />
          ))}

        {phase.kind === 'classifying' && <ClassifyingScreen startedAt={phase.startedAt} batchSize={batchSize} />}

        {phase.kind === 'plan' && (
          <PlanScreen
            items={phase.items}
            tree={phase.folderTree}
            excludePrefixes={status?.excludePrefixes ?? ['05已完成案件']}
            usage={phase.usage}
            model={status?.model}
            summary={phase.summary}
            banner={phase.banner}
            aiPending={phase.aiPending}
            onChange={updateItem}
            onBulkApply={bulkApply}
            onBack={backToIdle}
            onConfirmExecute={confirmExecute}
            onDismissBanner={() =>
              setPhase((p) => (p.kind === 'plan' ? { ...p, banner: undefined } : p))
            }
            onToggleRule={async (ruleId, enabled) => {
              // Throw on Err so PlanRow's catch can roll back the optimistic
              // localDisabled flag. send() resolves both Ok and Err — without
              // this gate, an Err response would silently leave the UI
              // displaying "已停用" while the rule is still firing.
              const r = await send({ type: 'toggleRule', ruleId, enabled })
              if (!r.ok) throw new Error(r.message || 'toggle failed')
            }}
          />
        )}

        {phase.kind === 'executing' && (
          <ExecutingScreen state={phase.state} onCancel={cancelExecute} />
        )}

        {phase.kind === 'done' && (
          <DoneScreen
            state={phase.state}
            batchSize={batchSize}
            onReset={finishAndReset}
            onRetry={retryFailed}
            onContinue={() => void continueToNextBatch()}
          />
        )}

        {phase.kind === 'error' && (
          <ErrorScreen
            code={phase.code}
            message={phase.message}
            onBack={backToIdle}
          />
        )}
      </div>
      {showShortcuts && <ShortcutsOverlay onClose={() => setShowShortcuts(false)} />}
    </main>
    </LayoutContext.Provider>
  )
}

// Compact phase-context badge shown in the header during non-idle phases.
// Tells the user "what is this popup doing right now" without making them
// scan the body.
//
// UI/UX 檢討 (2026-07)：plan 階段不再重複顯示封數（摘要卡的「共 N 封」
// 是這個數字唯一的家）、也不再掛 model 徽章 — 模型資訊屬於設定頁，
// 每日批次審閱時是純遙測噪音。model 徽章只留在 classifying 階段
// （那一刻「哪個模型在計費」才是相關資訊）。
function PhaseContext({ phase, model }: { phase: Phase; model?: string }) {
  const modelShort =
    model && phase.kind === 'classifying'
      ? model.replace(/^claude-/, '').replace(/-\d{8}$/, '')
      : null
  let label: string | null = null
  if (phase.kind === 'classifying') label = 'AI 分類中…'
  else if (phase.kind === 'executing') label = `處理中 ${phase.state.current} / ${phase.state.total}`
  else if (phase.kind === 'done')
    label =
      phase.state.summary.errors > 0
        ? `完成(${phase.state.summary.errors} 失敗)`
        : '已完成'

  if (!label && !modelShort) return null
  return (
    <div className="hidden sm:flex items-center gap-1.5 text-[10px] text-muted-foreground">
      {label && <span className="tabular-nums">{label}</span>}
      {modelShort && phase.kind !== 'idle' && (
        <span
          className="font-mono px-1.5 py-0.5 rounded bg-muted text-[9px]"
          title={`分類模型:${model}`}
        >
          {modelShort}
        </span>
      )}
    </div>
  )
}

function ShortcutsOverlay({ onClose }: { onClose: () => void }) {
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-label="鍵盤捷徑"
    >
      <div
        className="rounded-md border border-border bg-card text-card-foreground shadow-xl w-[420px] max-w-[90vw] p-5 space-y-3"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between">
          <h2 className="text-sm font-semibold">鍵盤捷徑</h2>
          <button
            type="button"
            onClick={onClose}
            className="text-muted-foreground hover:text-foreground"
            aria-label="關閉"
          >
            <XCircle className="size-4" />
          </button>
        </div>
        <dl className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-1.5 text-xs">
          {[
            { keys: ['j', 'k'], desc: '上下移動焦點' },
            { keys: ['x', 'Space'], desc: '勾選 / 取消勾選此項' },
            { keys: ['Enter'], desc: '展開 / 收起此項' },
            { keys: ['d'], desc: '把這封改為「刪除」' },
            { keys: ['s'], desc: '把這封改為「保留」' },
            { keys: ['/'], desc: '過濾清單 — 主旨 / 寄件人 / 路徑' },
            { keys: ['u'], desc: '完成畫面 — 撤回剛剛的批次' },
            { keys: ['?'], desc: '顯示這個說明' },
            { keys: ['Esc'], desc: '關閉面板 / 收起選單' },
          ].map((row) => (
            <Fragment key={row.desc}>
              <dt className="flex items-center gap-1 font-mono">
                {row.keys.map((k, i) => (
                  <Fragment key={k}>
                    {i > 0 && <span className="text-muted-foreground">/</span>}
                    <kbd className="px-1.5 py-0.5 rounded bg-muted border border-border text-[10px]">{k}</kbd>
                  </Fragment>
                ))}
              </dt>
              <dd className="self-center">{row.desc}</dd>
            </Fragment>
          ))}
        </dl>
        <p className="text-[10px] text-muted-foreground pt-2 border-t border-border">
          AI 分類進行中時鍵盤捷徑暫停 — 等清單載入完成後再用。
        </p>
      </div>
    </div>
  )
}

// ---- Idle ----------------------------------------------------------------

function IdleScreen({
  status,
  batchSize,
  canStart,
  onStart,
  onRefreshTree,
  onOpenOptions,
  folderActivity,
  onRefreshActivity,
  weeklyDigest,
  onDismissDigest,
}: {
  status: StatusData | null
  batchSize: number
  canStart: boolean
  onStart: () => void
  onRefreshTree: () => void
  onOpenOptions: () => void
  folderActivity: FolderActivityRow[]
  onRefreshActivity: () => Promise<void> | void
  weeklyDigest: WeeklyDigestSummary | null
  onDismissDigest: () => void
}) {
  const { isWide } = useLayout()
  const issues: { label: string; action?: { label: string; onClick: () => void } }[] = []
  if (status) {
    if (!status.owaConnected)
      issues.push({ label: '請先打開並登入 Outlook 網頁版分頁、再重開此 popup' })
    if (status.owaConnected && !status.tokenValid)
      issues.push({ label: '無法讀取 Outlook 登入資訊、稍等幾秒再試' })
    if (!status.apiKeyConfigured)
      issues.push({
        label: '尚未設定 Claude API key',
        action: { label: '前往設定', onClick: onOpenOptions },
      })
    // Firm-style settings (internal domains / primary root / category
    // hints) are ALL optional — a sole-practitioner or anyone with a
    // personal Gmail-only workflow shouldn't be nagged. Initial scan
    // surfaces its own "needs root path" error when invoked without
    // configuration; classify works fine with all three empty.
  }

  // First-time user (no recent activity yet) sees the start button as
  // the hero — there's nothing else useful to show. Once the lawyer has
  // run at least one batch, Recent Activity becomes the visual anchor
  // (they often open the popup to jump to a case, not to start a new
  // batch); start gets a compact strip below it.
  const hasActivity = folderActivity.length > 0

  const startStrip = (
    <div className="flex items-center gap-3">
      <Button size="lg" onClick={onStart} disabled={!canStart} className="flex-1">
        <Play /> 開始歸類
      </Button>
      <div className="flex flex-col gap-0.5 text-[11px] text-muted-foreground">
        <span>本次取最近 {batchSize} 封</span>
        <div className="flex items-center gap-2">
          <button type="button" onClick={onOpenOptions} className="hover:underline">
            設定
          </button>
          <span aria-hidden="true">·</span>
          <button
            type="button"
            onClick={onRefreshTree}
            className="hover:underline disabled:opacity-50"
            disabled={!canStart}
          >
            <RefreshCw className="inline size-3 mr-0.5" />
            重新偵測資料夾
          </button>
        </div>
      </div>
    </div>
  )

  return (
    <div className="space-y-4">
      {issues.length > 0 && (
        <ul className="space-y-2 rounded-md border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900">
          {issues.map((issue, i) => (
            <li key={i} className="flex items-start gap-2">
              <AlertTriangle className="size-3.5 mt-0.5 shrink-0" />
              <span className="flex-1">{issue.label}</span>
              {issue.action && (
                <Button size="sm" variant="outline" className="h-6 px-2 text-[11px]" onClick={issue.action.onClick}>
                  {issue.action.label}
                </Button>
              )}
            </li>
          ))}
        </ul>
      )}

      {weeklyDigest && (
        <WeeklyDigestCard summary={weeklyDigest} onDismiss={onDismissDigest} onOpenOptions={onOpenOptions} />
      )}

      {/* Onboarding checklist — surfaces only on first run (no activity
          yet) and only after the user is past the blocking issues. Once
          they've completed even one classify batch, hasActivity flips
          true and this disappears. Each step has its own done-check so
          partial completers see what's left. */}
      {!hasActivity && status && status.apiKeyConfigured && status.tokenValid && (
        <OnboardingChecklist status={status} onOpenOptions={onOpenOptions} />
      )}

      {hasActivity ? (
        <>
          <RecentActivityPanel
            rows={folderActivity}
            isWide={isWide}
            onRefresh={onRefreshActivity}
            hero
          />
          {startStrip}
        </>
      ) : (
        <>
          {startStrip}
          <RecentActivityPanel
            rows={folderActivity}
            isWide={isWide}
            onRefresh={onRefreshActivity}
          />
        </>
      )}
    </div>
  )
}

// ---- Onboarding checklist ------------------------------------------------

/**
 * Two-step "first run" guide. Only renders when the user is past the
 * hard blockers (API key set, OWA token valid) AND has no folder
 * activity yet (= first batch hasn't been completed). Disappears
 * permanently once the user runs their first classify, since
 * `folderActivity` will be non-empty.
 *
 * Why no "set API key" step: the parent already gates rendering on
 * `apiKeyConfigured=true`, so the step would always be ✓ — misleading.
 * The blocking issues banner above handles the "API key not set" case
 * with its own prompt.
 */
function OnboardingChecklist({
  status,
  onOpenOptions,
}: {
  status: StatusData
  onOpenOptions: () => void
}) {
  const steps: Array<{ done: boolean; label: string; hint?: string }> = [
    {
      done: status.primaryRootPath.length > 0,
      label: '選擇主要根資料夾(建議、非必要)',
      hint: '設定頁 → 歸類偏好 → 主要根資料夾',
    },
    {
      done: false,
      label: '按下方「開始歸類」跑第一批',
      hint: '系統會依資料夾結構 + AI 判斷分類',
    },
  ]
  const allDone = steps.every((s) => s.done)
  if (allDone) return null
  return (
    <div className="rounded-md border border-border bg-card p-3 space-y-2">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-medium">第一次使用?照這幾步走</h3>
        <button
          type="button"
          onClick={onOpenOptions}
          className="text-[10px] text-muted-foreground hover:text-foreground hover:underline"
        >
          打開設定
        </button>
      </div>
      <ol className="space-y-1.5 text-xs">
        {steps.map((step, i) => (
          <li key={i} className="flex items-start gap-2">
            <span
              className={cn(
                'shrink-0 mt-0.5 inline-flex items-center justify-center size-4 rounded-full text-[10px] font-mono',
                step.done
                  ? 'bg-emerald-100 text-emerald-800'
                  : 'bg-muted text-muted-foreground',
              )}
              aria-label={step.done ? '已完成' : '待完成'}
            >
              {step.done ? '✓' : i + 1}
            </span>
            <div className="flex-1">
              <div className={cn(step.done && 'text-muted-foreground line-through')}>
                {step.label}
              </div>
              {step.hint && !step.done && (
                <div className="text-[10px] text-muted-foreground mt-0.5">
                  {step.hint}
                </div>
              )}
            </div>
          </li>
        ))}
      </ol>
    </div>
  )
}

// ---- Weekly digest card --------------------------------------------------

function WeeklyDigestCard({
  summary,
  onDismiss,
  onOpenOptions,
}: {
  summary: WeeklyDigestSummary
  onDismiss: () => void
  onOpenOptions: () => void
}) {
  const totalHandled = summary.moved + summary.deleted + summary.foldersCreated
  const healthHint =
    summary.sleepingCount + summary.orphanedCount + summary.conflictsCount > 0
  return (
    <div className="rounded-md border border-foreground/30 bg-card p-3 space-y-2 text-xs">
      <div className="flex items-baseline justify-between">
        <h3 className="text-sm font-medium">
          近期摘要
          <span className="ml-2 text-[10px] text-muted-foreground tabular-nums font-normal">
            最近 {summary.daysSpan} 天
          </span>
        </h3>
        <button
          type="button"
          onClick={onDismiss}
          className="text-muted-foreground hover:text-foreground"
          aria-label="知道了"
          title="知道了 — 7 天後再顯示下一份"
        >
          <XCircle className="size-3.5" />
        </button>
      </div>
      {totalHandled === 0 && summary.rulesAdded === 0 ? (
        <p className="text-muted-foreground italic">這段時間沒有歸檔活動。</p>
      ) : (
        <ul className="space-y-0.5 text-muted-foreground">
          {totalHandled > 0 && (
            <li>
              處理 <span className="font-mono tabular-nums text-foreground">{totalHandled}</span> 封 —
              移動 {summary.moved} · 刪除 {summary.deleted} · 新建資料夾 {summary.foldersCreated}
            </li>
          )}
          {summary.rulesAdded > 0 && (
            <li>
              規則庫淨增 <span className="font-mono tabular-nums text-foreground">{summary.rulesAdded}</span> 條
              · 目前共 {summary.rulesCountNow} 條
            </li>
          )}
          {summary.errors > 0 && (
            <li className="text-amber-700">執行失敗 {summary.errors} 件</li>
          )}
        </ul>
      )}
      {healthHint && (
        <div className="flex items-center justify-between rounded border border-amber-200 bg-amber-50 px-2 py-1.5 text-[11px] text-amber-900">
          <span>
            規則健康度待處理:
            {summary.sleepingCount > 0 && <> 休眠 {summary.sleepingCount}</>}
            {summary.conflictsCount > 0 && <> · 衝突 {summary.conflictsCount}</>}
            {summary.orphanedCount > 0 && <> · 目標遺失 {summary.orphanedCount}</>}
          </span>
          <button
            type="button"
            onClick={onOpenOptions}
            className="text-amber-900 hover:underline shrink-0"
          >
            前往整理
          </button>
        </div>
      )}
    </div>
  )
}

// ---- Recent activity quick-jump panel ------------------------------------
//
// Reads from chrome.storage.local.folderActivity, populated by execute.ts
// after each successful classify batch (Phase 1) and on-demand via a Graph
// API scan (Phase 3 manual refresh). Click a row → SW handler opens the
// folder in OWA (focuses existing OWA tab or opens a new one).

function RecentActivityPanel({
  rows,
  isWide,
  onRefresh,
  hero = false,
}: {
  rows: FolderActivityRow[]
  isWide: boolean
  onRefresh: () => Promise<void> | void
  /** When true, render as a prominent card (lawyer has prior activity);
   *  otherwise render as a compact bottom-of-IdleScreen hint. */
  hero?: boolean
}) {
  const [refreshing, setRefreshing] = useState(false)
  const [refreshError, setRefreshError] = useState<string | null>(null)
  const [showAll, setShowAll] = useState(false)

  const PREVIEW_LIMIT = isWide ? 14 : 6
  const visibleRows = showAll ? rows : rows.slice(0, PREVIEW_LIMIT)
  const hasMore = !showAll && rows.length > PREVIEW_LIMIT

  const [refreshOk, setRefreshOk] = useState<string | null>(null)
  // Auto-clear the success message so it doesn't pile up across refreshes.
  useEffect(() => {
    if (!refreshOk) return
    const t = window.setTimeout(() => setRefreshOk(null), 3000)
    return () => window.clearTimeout(t)
  }, [refreshOk])

  async function handleRefresh() {
    setRefreshing(true)
    setRefreshError(null)
    setRefreshOk(null)
    try {
      const r = await send<{ scanned: number; matched: number; errored: number; capped: boolean }>({
        type: 'refreshFolderActivity',
      })
      if (!r.ok) {
        setRefreshError(r.message || r.code || '掃描失敗')
      } else {
        await onRefresh()
        const d = r.data
        const parts: string[] = []
        if (d) {
          parts.push(`已掃 ${d.scanned} 個資料夾`)
          if (d.matched > 0) parts.push(`更新 ${d.matched}`)
          if (d.errored > 0) parts.push(`${d.errored} 個失敗`)
          if (d.capped) parts.push('已達 200 個上限')
        } else {
          parts.push('完成')
        }
        setRefreshOk(parts.join(' · '))
      }
    } finally {
      setRefreshing(false)
    }
  }

  async function handleClick(row: FolderActivityRow) {
    if (window.parent !== window) {
      // FAB iframe mode: ask the content script in OWA to do an in-page
      // navigation (click the folder's treeitem → OWA's SPA navigates
      // without reloading). This is the no-flash path. The folder name is
      // the leaf segment of folderPath — OWA's treeitem uses data-folder-name
      // which matches displayName.
      const folderName = row.folderPath.split('/').filter(Boolean).pop() || ''
      postToParent({
        type: 'mail-organizer/navigate-folder',
        folderName,
        folderPath: row.folderPath,
        folderId: row.folderId,
      })
      // Close the panel so the user sees the navigation result.
      postToParent({ type: 'mail-organizer/close-panel' })
    } else {
      // Toolbar popup mode: no embedded OWA window to message; ask SW to
      // switch / open the OWA tab via chrome.tabs (full reload, but
      // unavoidable in this mode — Chrome popup closes on focus loss anyway).
      await send({ type: 'navigateToFolder', folderId: row.folderId })
    }
  }

  if (rows.length === 0) {
    return (
      <div className="space-y-2 border-t border-border pt-3">
        <div className="flex items-baseline justify-between">
          <h3 className="text-xs font-medium text-muted-foreground">近日活動</h3>
          <button
            type="button"
            onClick={() => void handleRefresh()}
            disabled={refreshing}
            className="text-[10px] text-muted-foreground hover:text-foreground hover:underline disabled:opacity-50"
            title="呼叫 Outlook API 掃描所有案件資料夾的最新信件 — 用來補上手動拖入(不是透過歸類)的信件紀錄。可能要 10-30 秒"
          >
            {refreshing ? '掃描中…(可能要 10-30 秒)' : '掃描 Outlook'}
          </button>
        </div>
        <p className="text-[11px] text-muted-foreground italic">
          執行歸類後、最近收信的案件會出現在這裡、點一下直接跳到該資料夾。
        </p>
        {refreshError && (
          <p className="text-[10px] text-red-700">{refreshError}</p>
        )}
      </div>
    )
  }

  return (
    <div
      className={cn(
        hero
          ? 'space-y-3 rounded-md border border-border bg-card p-3'
          : 'space-y-2 border-t border-border pt-3',
      )}
    >
      <div className="flex items-baseline justify-between">
        <h3
          className={cn(
            'font-medium',
            hero ? 'text-sm text-foreground' : 'text-xs text-muted-foreground',
          )}
        >
          近日活動 <span className="font-mono tabular-nums text-muted-foreground">({rows.length})</span>
        </h3>
        <button
          type="button"
          onClick={() => void handleRefresh()}
          disabled={refreshing}
          className="text-[10px] text-muted-foreground hover:text-foreground hover:underline disabled:opacity-50"
          title="呼叫 Outlook API、掃描所有案件資料夾的最新信件時間 — 用來補上手動拖入(不是透過歸類)的信件紀錄。可能要 10-30 秒"
        >
          {refreshing ? '掃描中…(10-30 秒)' : '重新整理'}
        </button>
      </div>

      <ul className={cn('grid gap-1', isWide ? 'grid-cols-2 gap-x-3' : 'grid-cols-1')}>
        {visibleRows.map((row) => (
          <li key={row.folderId}>
            <FolderActivityListItem row={row} onClick={() => void handleClick(row)} />
          </li>
        ))}
      </ul>

      {hasMore && (
        <button
          type="button"
          onClick={() => setShowAll(true)}
          className="w-full text-[11px] text-muted-foreground hover:text-foreground hover:underline py-1"
        >
          查看全部 ({rows.length - PREVIEW_LIMIT} 更多) →
        </button>
      )}
      {showAll && rows.length > PREVIEW_LIMIT && (
        <button
          type="button"
          onClick={() => setShowAll(false)}
          className="w-full text-[11px] text-muted-foreground hover:text-foreground hover:underline py-1"
        >
          收起
        </button>
      )}

      {refreshError && <p className="text-[10px] text-red-700">{refreshError}</p>}
      {refreshOk && <p className="text-[10px] text-emerald-700">{refreshOk}</p>}
    </div>
  )
}

function FolderActivityListItem({
  row,
  onClick,
}: {
  row: FolderActivityRow
  onClick: () => void
}) {
  // Split path so the last segment (the case folder) is bold and the
  // ancestor path renders as muted breadcrumb.
  const segments = row.folderPath.split('/').filter(Boolean)
  const leaf = segments[segments.length - 1] || row.folderPath
  const ancestors = segments.slice(0, -1).join(' / ')
  const relative = formatRelative(row.lastActiveAt)
  // Build a multi-line tooltip showing path + latest message (when known
  // from a refresh scan). Native browser title respects \n so the user
  // gets readable preview on hover without us reaching for a positioned
  // popover.
  // Synced-in row whose folderId doesn't exist on THIS machine. Could
  // happen if the other machine uses a different Outlook account, or
  // if folderCache here is stale. Render greyed-out, disable click —
  // better than clicking to a 404 in OWA. Undefined (no cache at all)
  // = first-run state, render normally.
  const isMissingLocally = row.localExists === false
  const tooltipParts: string[] = [row.folderPath]
  tooltipParts.push(
    `最後活動：${new Date(row.lastActiveAt).toLocaleString('zh-TW')}`,
  )
  if (isMissingLocally) {
    tooltipParts.push('')
    tooltipParts.push('⚠ 此資料夾在本機 Outlook 找不到')
    tooltipParts.push('(可能是從另一台機器同步、但本機帳戶不同)')
  }
  if (row.latestMessage) {
    tooltipParts.push('')
    tooltipParts.push(`最新一封：${row.latestMessage.subject || '(無主旨)'}`)
    if (row.latestMessage.from) {
      tooltipParts.push(`寄件人：${row.latestMessage.from}`)
    }
    tooltipParts.push(
      `收件：${new Date(row.latestMessage.receivedAt).toLocaleString('zh-TW')}`,
    )
  }
  return (
    <button
      type="button"
      onClick={isMissingLocally ? undefined : onClick}
      disabled={isMissingLocally}
      className={cn(
        'w-full text-left rounded-md border border-transparent px-2 py-1.5 text-xs transition-colors',
        isMissingLocally
          ? 'opacity-50 cursor-not-allowed'
          : 'hover:border-border hover:bg-accent/30',
      )}
      title={tooltipParts.join('\n')}
    >
      <div className="flex items-baseline gap-2">
        <span className="font-medium truncate flex-1">{leaf}</span>
        {isMissingLocally && (
          <span className="text-[9px] text-amber-700 shrink-0">
            ⚠ 本機不存在
          </span>
        )}
        <span className="font-mono text-[10px] text-muted-foreground shrink-0 tabular-nums">
          {relative}
        </span>
      </div>
      {ancestors && (
        <div className="text-[10px] text-muted-foreground truncate mt-0.5">
          {ancestors}
        </div>
      )}
      {row.latestMessage?.subject && (
        <div
          className="text-[10px] text-muted-foreground italic truncate mt-0.5"
          title={row.latestMessage.subject}
        >
          {row.latestMessage.subject}
        </div>
      )}
    </button>
  )
}

function formatRelative(iso: string): string {
  const t = Date.parse(iso)
  if (!Number.isFinite(t)) return ''
  const diffMs = Date.now() - t
  // Clamp negative (future) timestamps to 0 — happens if client clock is
  // ahead of server, or a malformed ISO crept in. Otherwise we'd render
  // nonsense like「-3 分前」.
  const min = Math.max(0, Math.floor(diffMs / 60_000))
  if (min < 1) return '剛剛'
  if (min < 60) return `${min} 分前`
  const hr = Math.floor(min / 60)
  if (hr < 24) return `${hr} 小時前`
  const day = Math.floor(hr / 24)
  if (day < 7) return `${day} 天前`
  // Beyond a week, show absolute date (no year if same year)
  const d = new Date(t)
  const y = d.getFullYear()
  const m = d.getMonth() + 1
  const dd = d.getDate()
  if (y === new Date().getFullYear()) return `${m}/${dd}`
  return `${y}/${m}/${dd}`
}

// ---- Classifying ---------------------------------------------------------

type ClassifyStageData =
  | { stage: 'idle' }
  | { stage: 'fetching_tree' }
  | { stage: 'fetching_inbox' }
  | { stage: 'matching_rules'; total: number }
  | { stage: 'calling_ai'; toClassify: number }

const STAGE_LABEL: Record<ClassifyStageData['stage'], string> = {
  idle: '處理中…',
  fetching_tree: '抓取資料夾樹…',
  fetching_inbox: '抓取收件夾…',
  matching_rules: '套用規則中…',
  calling_ai: 'Claude 分類中…',
}

function ClassifyingScreen({ startedAt, batchSize }: { startedAt: number; batchSize: number }) {
  const [elapsed, setElapsed] = useState(0)
  const [stageData, setStageData] = useState<ClassifyStageData>({ stage: 'idle' })

  useEffect(() => {
    const elapsedId = setInterval(() => setElapsed(Date.now() - startedAt), 250)
    const stageId = setInterval(async () => {
      const r = await send<ClassifyStageData>({ type: 'getClassifyStage' })
      if (r.ok && r.data) setStageData(r.data)
    }, 400)
    return () => {
      clearInterval(elapsedId)
      clearInterval(stageId)
    }
  }, [startedAt])

  const seconds = Math.floor(elapsed / 1000)
  const label = STAGE_LABEL[stageData.stage]
  const detail =
    stageData.stage === 'calling_ai'
      ? ` · ${stageData.toClassify} 封送 AI`
      : stageData.stage === 'matching_rules'
      ? ` · 共 ${stageData.total} 封`
      : ''

  return (
    <div className="flex flex-col items-center justify-center py-12 gap-3 text-center">
      <Loader2 className="size-8 animate-spin text-foreground" />
      <div className="space-y-1">
        <div className="text-sm font-medium">{label}</div>
        <div className="text-xs text-muted-foreground tabular-nums">
          已 {seconds}s · {batchSize} 封{detail} · 首批可能 30-60s
        </div>
      </div>
    </div>
  )
}

// ---- Plan ----------------------------------------------------------------

// Group key for the "依資料夾分組" sort (覆核 P1-2). Same-destination rows
// share a key so they cluster together; the reviewer can then eyeball "all N
// mails going to 客戶A/訴訟" as a block instead of hunting them scattered
// across an 80-row list. Non-move actions get their own synthetic buckets.
function targetGroupKey(item: PlanItem): string {
  if (item.action === 'move') return item.targetFolderPath ?? '（未指定目標）'
  if (item.action === 'new_folder') {
    return item.suggestedParentPath && item.suggestedFolderName
      ? `${joinFolderPath(item.suggestedParentPath, item.suggestedFolderName)}（新建）`
      : '（新建・未完成）'
  }
  if (item.action === 'delete') return '（刪除）'
  return '（保留）'
}

function PlanScreen({
  items,
  tree,
  excludePrefixes,
  usage,
  model,
  summary,
  banner,
  aiPending,
  onChange,
  onBulkApply,
  onBack,
  onConfirmExecute,
  onDismissBanner,
  onToggleRule,
}: {
  items: PlanItem[]
  tree: MailFolderNode[]
  excludePrefixes: string[]
  usage: ClassifyUsage | null
  model?: string
  summary: { ruleHits: number; aiHandled: number }
  banner?: { code: string; message: string }
  aiPending?: {
    startedAt: number
    totalEmails: number
    completedEmails: number
    chunks: number
    completedChunks: number
  }
  onChange: (item: PlanItem) => void
  onBulkApply: (ids: Set<string>, transform: (item: PlanItem) => PlanItem) => void
  onBack: () => void
  onConfirmExecute: () => void
  onDismissBanner: () => void
  onToggleRule: (ruleId: string, enabled: boolean) => Promise<void>
}) {
  const counts = useMemo(() => planSummary(items), [items])
  // Picker shows real folders + the pending destinations in this same plan, so
  // user can re-target later items to a folder that doesn't exist in Outlook
  // yet (execute will create + reorder).
  //
  // Perf (audit): buildAugmentedTree deep-clones the whole folder tree
  // (JSON round-trip, ~200KB) — but it only READS the new_folder destination
  // triples. Keying it directly on `items` recomputed the clone on EVERY
  // edit (action toggles, d/s shortcuts, move-target picks). Derive a cheap
  // proxy key from just those triples so unrelated edits reuse the cached
  // tree; only a change to an actual pending destination re-clones.
  const pendingFoldersKey = useMemo(
    () =>
      items
        .filter((i) => i.action === 'new_folder')
        .map(
          (i) =>
            `${i.emailId}\u0000${i.suggestedFolderName ?? ''}\u0000${i.suggestedParentPath ?? ''}`,
        )
        .join('\u0001'),
    [items],
  )
  const augmentedTree = useMemo(
    () => buildAugmentedTree(tree, items),
    // `pendingFoldersKey` is a complete proxy for the slice of `items` that
    // buildAugmentedTree reads — when it's unchanged, the cached result is
    // identical by construction.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [tree, pendingFoldersKey],
  )

  // Path index — flatten once per augmented-tree change, share across all
  // PlanRow instances. Avoids 50× redundant flattens on each onChange.
  const validPaths = useMemo<ReadonlySet<string>>(() => {
    const s = new Set<string>()
    for (const n of flattenFolderTree(augmentedTree)) s.add(n.path)
    return s
  }, [augmentedTree])

  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())

  // Filter state — `/` key narrows the visible list by substring against
  // subject / from / target path / reason. Each comma-separated token
  // becomes a chip; multiple chips AND together so each one further
  // narrows. Selection survives filter changes (the user may select then
  // filter to verify), but focusedIndex resets because the index is into
  // filteredItems and would otherwise point at hidden rows.
  const [filterChips, setFilterChips] = useState<string[]>([])
  const [filterDraft, setFilterDraft] = useState('')
  const [filterVisible, setFilterVisible] = useState(false)
  const filterInputRef = useRef<HTMLInputElement | null>(null)

  // 「待確認」過濾（UI/UX 檢討 2026-07）：80 封批次裡律師真正要親自看的
  // 是 unresolved / 低信心那 5-20 封，其他都是她已信任的規則命中。摘要卡
  // 的「待確認 N」chip 一鍵切到只看這批 — 這是整個畫面最重要的減噪工具。
  const [attentionOnly, setAttentionOnly] = useState(false)
  const needsAttention = useCallback(
    (item: PlanItem) =>
      item.source === 'unresolved' || (item.source !== 'rule' && item.confidence < 0.5),
    [],
  )
  const attentionCount = useMemo(
    () => items.filter(needsAttention).length,
    [items, needsAttention],
  )

  const filteredItems = useMemo(() => {
    let base = items
    if (attentionOnly) base = base.filter(needsAttention)
    if (filterChips.length === 0) return base
    const needles = filterChips.map((c) => c.toLowerCase())
    return base.filter((item) => {
      const haystack = [
        item.emailSubject ?? '',
        item.emailFrom ?? '',
        item.targetFolderPath ?? '',
        item.suggestedParentPath ?? '',
        item.suggestedFolderName ?? '',
        item.reason ?? '',
      ]
        .join(' ')
        .toLowerCase()
      return needles.every((n) => haystack.includes(n))
    })
  }, [items, filterChips, attentionOnly, needsAttention])

  function commitFilterDraft() {
    const t = filterDraft.trim()
    if (!t) return
    setFilterChips((prev) => (prev.includes(t) ? prev : [...prev, t]))
    setFilterDraft('')
  }

  function clearFilter() {
    setFilterChips([])
    setFilterDraft('')
    setFilterVisible(false)
  }

  // Auto-propagate toast — shows once a same-subject sync fires. The hint
  // banner that used to live inside each PlanRow is gone (visual noise), so
  // without this the propagation is silent and feels "broken" when siblings
  // happen to be off-screen.
  const [propagateToast, setPropagateToast] = useState<{ count: number; key: number } | null>(null)
  useEffect(() => {
    if (!propagateToast) return
    const t = window.setTimeout(() => setPropagateToast(null), 3000)
    return () => window.clearTimeout(t)
  }, [propagateToast])

  // Conversation-cluster collapse: same-ConversationId items that share the
  // SAME action+target are rendered as one "rep" row plus a "+N 封同對話"
  // hint to expand. Reduces a thread of 8 reply emails (all going to same
  // case folder) from 8 ~80px rows to 1 + hint.
  //
  // Only collapse when:
  //   - members > 1
  //   - all share same action + targetFolderPath (otherwise mixed
  //     decisions, can't safely hide siblings)
  //
  // Per-group expand state survives action changes within the plan
  // (Set keyed by conversationId). Default: collapsed.
  const [expandedConversations, setExpandedConversations] = useState<Set<string>>(
    new Set(),
  )

  // 依資料夾分組排序 (覆核 P1-2). When on, reorder so same-target rows sit
  // together. Secondary key = conversationId (falls back to emailId) so a
  // conversation cluster's members stay contiguous within their folder bucket
  // — the collapse-link placement logic below relies on that contiguity.
  // Array.sort is stable (ES2019+), so ties preserve the original order.
  const [sortMode, setSortMode] = useState<'default' | 'byTarget'>('default')
  const orderedItems = useMemo(() => {
    if (sortMode !== 'byTarget') return filteredItems
    // Strict code-point compare (NOT localeCompare): the group-header dedup
    // and the conversation-cluster contiguity checks downstream test group
    // keys with strict `!==`. localeCompare returns 0 for canonically-equal
    // but code-point-distinct strings (e.g. NFC vs NFD "café"), which would
    // make this an inconsistent (invalid) ordering vs the strict dedup —
    // interleaving two such groups, duplicating headers and splitting
    // clusters. Code-point compare keeps sort + dedup on the same equality.
    const cmp = (x: string, y: string): number => (x < y ? -1 : x > y ? 1 : 0)
    return [...filteredItems].sort((a, b) => {
      const kc = cmp(targetGroupKey(a), targetGroupKey(b))
      if (kc !== 0) return kc
      return cmp(a.conversationId ?? a.emailId, b.conversationId ?? b.emailId)
    })
  }, [filteredItems, sortMode])

  // Per-target-group tallies (from filteredItems so hidden siblings count too)
  // — the group header shows the true "N mails → this folder" and 全選此組
  // selects ALL of them, not just the rows currently rendered.
  const targetGroups = useMemo(() => {
    const ids = new Map<string, string[]>()
    for (const item of filteredItems) {
      const k = targetGroupKey(item)
      const arr = ids.get(k) ?? []
      arr.push(item.emailId)
      ids.set(k, arr)
    }
    return ids
  }, [filteredItems])

  const conversationGroups = useMemo(() => {
    const byConv = new Map<string, PlanItem[]>()
    for (const item of orderedItems) {
      const cid = item.conversationId
      if (!cid) continue
      const arr = byConv.get(cid) ?? []
      arr.push(item)
      byConv.set(cid, arr)
    }
    // Drop solo / non-unified groups — render those normally.
    const collapsible = new Map<
      string,
      { rep: PlanItem; siblings: PlanItem[] }
    >()
    for (const [cid, members] of byConv) {
      if (members.length < 2) continue
      const first = members[0]!
      const unified = members.every(
        (m) =>
          m.action === first.action &&
          m.targetFolderPath === first.targetFolderPath &&
          m.suggestedFolderName === first.suggestedFolderName &&
          m.suggestedParentPath === first.suggestedParentPath,
      )
      if (!unified) continue
      collapsible.set(cid, { rep: first, siblings: members.slice(1) })
    }
    return collapsible
  }, [orderedItems])

  // The items actually rendered: skip siblings of COLLAPSED groups.
  const displayItems = useMemo(() => {
    return orderedItems.filter((item) => {
      const cid = item.conversationId
      if (!cid) return true
      const group = conversationGroups.get(cid)
      if (!group) return true // not a unified group, render normally
      // This item is part of a unified group.
      if (item.emailId === group.rep.emailId) return true // always render rep
      // Sibling — render only if group is expanded.
      return expandedConversations.has(cid)
    })
  }, [orderedItems, conversationGroups, expandedConversations])

  function toggleConversation(cid: string) {
    setExpandedConversations((prev) => {
      const next = new Set(prev)
      if (next.has(cid)) next.delete(cid)
      else next.add(cid)
      return next
    })
  }

  // Auto-prune selection if items disappear (shouldn't happen mid-plan but defensive)
  useEffect(() => {
    const valid = new Set(items.map((i) => i.emailId))
    let changed = false
    const filtered = new Set<string>()
    for (const id of selectedIds) {
      if (valid.has(id)) filtered.add(id)
      else changed = true
    }
    if (changed) setSelectedIds(filtered)
  }, [items, selectedIds])

  function toggleSelect(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }

  // ---- Keyboard navigation ------------------------------------------------
  // Focused index drives j/k navigation, x toggles select, Enter expands the
  // row (via tokens that PlanRow watches), d/s set action without leaving the
  // keyboard. Move requires a target so keyboard does not bind a key for it.

  const [focusedIndex, setFocusedIndex] = useState<number | null>(null)
  const [scrollIntoViewToken, setScrollIntoViewToken] = useState(0)
  const [toggleExpandToken, setToggleExpandToken] = useState(0)
  // 持久化（UI/UX 檢討 2026-07）：舊版是 component-local state，每次開
  // popup 提示都重生 — 每天用的人被迫反覆關同一條提示。localStorage 在
  // extension popup 內跨 session 存活；? overlay 仍可隨時查快捷鍵。
  const [shortcutsHintDismissed, setShortcutsHintDismissed] = useState(
    () => localStorage.getItem('mo.shortcutsHintDismissed') === '1',
  )
  const dismissShortcutsHint = useCallback(() => {
    setShortcutsHintDismissed(true)
    try {
      localStorage.setItem('mo.shortcutsHintDismissed', '1')
    } catch {
      // quota / private mode — session-only dismissal is still fine
    }
  }, [])

  // Auto-clear focus when items disappear (defensive — same scenario as
  // selection pruning above). F15: keep focus in-range for displayItems
  // (the rendered rows), since j/k navigation now indexes that. Collapsing
  // a conversation cluster shrinks displayItems, so a focusedIndex at the
  // old tail must clamp down to stay on a visible row.
  useEffect(() => {
    if (focusedIndex !== null && focusedIndex >= displayItems.length) {
      setFocusedIndex(displayItems.length > 0 ? displayItems.length - 1 : null)
    }
  }, [displayItems.length, focusedIndex])

  // Reset focus when the filter narrows differently, OR the sort mode flips
  // (both reorder displayItems, so a held focusedIndex would point at a
  // different row than the user expects). Without this, j/k could land on an
  // item the user can no longer see while their visible list scrolls past
  // unexpected entries.
  useEffect(() => {
    setFocusedIndex(null)
  }, [filterChips, sortMode])

  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      // Don't hijack typing in inputs/textareas/contenteditable elements
      const t = e.target as HTMLElement | null
      const tag = t?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return
      // Don't compete with browser/system shortcuts
      if (e.metaKey || e.ctrlKey || e.altKey) return

      // `/` opens the filter input. Available even during AI classify so
      // the user can pre-narrow the rule-hit portion of the plan while
      // AI is still working.
      if (e.key === '/' && !e.shiftKey) {
        e.preventDefault()
        setFilterVisible(true)
        // Wait for the input to mount on next frame before focusing.
        window.setTimeout(() => filterInputRef.current?.focus(), 0)
        return
      }

      // 刀3 (UI/UX 檢討 2026-07)：AI 分類中不再封鎖鍵盤 — 律師每批開頭
      // 有 30-60 秒盯著已可審的規則命中列卻按不了 j/k/d/s，一天 2-3 批
      // 就是每天數分鐘死時間。安全性：focusedIndex 對 displayItems.length
      // 有 clamp，AI 結果只會 append、不會位移既有 index。

      // F15 (2026-06-03): navigate the VISIBLE list (displayItems), not
      // filteredItems. Rows are rendered from displayItems, which omits
      // siblings of collapsed conversation clusters. Indexing into
      // filteredItems let the j/k cursor land on a hidden sibling — no
      // focus ring showed, and d/s/Enter then mutated an email the user
      // couldn't see. Keying everything off displayItems means the cursor
      // only ever lands on a rendered row.
      const len = displayItems.length
      if (len === 0) return

      const key = e.key
      const lower = key.toLowerCase()

      if (lower === 'j' || key === 'ArrowDown') {
        e.preventDefault()
        setFocusedIndex((i) => (i === null ? 0 : Math.min(i + 1, len - 1)))
        setScrollIntoViewToken((v) => v + 1)
        setShortcutsHintDismissed(true)
      } else if (lower === 'k' || key === 'ArrowUp') {
        e.preventDefault()
        setFocusedIndex((i) => (i === null ? 0 : Math.max(i - 1, 0)))
        setScrollIntoViewToken((v) => v + 1)
        setShortcutsHintDismissed(true)
      } else if (key === ' ' || lower === 'x') {
        if (focusedIndex !== null) {
          e.preventDefault()
          const item = displayItems[focusedIndex]
          if (item) toggleSelect(item.emailId)
          setShortcutsHintDismissed(true)
        }
      } else if (key === 'Enter') {
        if (focusedIndex !== null) {
          e.preventDefault()
          setToggleExpandToken((v) => v + 1)
          setShortcutsHintDismissed(true)
        }
      } else if (lower === 'd') {
        if (focusedIndex !== null) {
          e.preventDefault()
          const item = displayItems[focusedIndex]
          if (item && item.action !== 'delete') {
            handleItemChange({
              ...item,
              action: 'delete',
              targetFolderId: undefined,
              targetFolderPath: undefined,
              suggestedFolderName: undefined,
              suggestedParentPath: undefined,
              userTouched: true,
              source: item.source === 'rule' ? 'ai' : item.source,
              ruleId: item.source === 'rule' ? undefined : item.ruleId,
            })
          }
          setShortcutsHintDismissed(true)
        }
      } else if (lower === 's') {
        if (focusedIndex !== null) {
          e.preventDefault()
          const item = displayItems[focusedIndex]
          if (item && item.action !== 'skip') {
            handleItemChange({
              ...item,
              action: 'skip',
              targetFolderId: undefined,
              targetFolderPath: undefined,
              suggestedFolderName: undefined,
              suggestedParentPath: undefined,
              userTouched: true,
              source: item.source === 'rule' ? 'ai' : item.source,
              ruleId: item.source === 'rule' ? undefined : item.ruleId,
            })
          }
          setShortcutsHintDismissed(true)
        }
      } else if (key === 'Escape') {
        setFocusedIndex(null)
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [displayItems, focusedIndex, aiPending, onChange])

  // Subject groups for "💡 N items share same subject" hints
  const subjectGroups = useMemo(() => {
    const m = new Map<string, string[]>()
    for (const item of items) {
      const sub = item.emailSubject?.trim()
      if (!sub) continue
      const norm = normalizeSubject(sub)
      if (!norm) continue
      const arr = m.get(norm) ?? []
      arr.push(item.emailId)
      m.set(norm, arr)
    }
    return m
  }, [items])

  function sameSubjectOtherIds(source: PlanItem): string[] {
    const sub = source.emailSubject?.trim()
    if (!sub) return []
    const norm = normalizeSubject(sub)
    if (!norm) return []
    const ids = subjectGroups.get(norm) ?? []
    return ids.filter((id) => id !== source.emailId)
  }

  function buildPropagateTransform(source: PlanItem): (item: PlanItem) => PlanItem {
    return (item): PlanItem => {
      const newSource: PlanItem['source'] = item.source === 'rule' ? 'ai' : item.source
      const newRuleId = item.source === 'rule' ? undefined : item.ruleId

      // Special case: source wants to create a new folder. We can't copy the
      // new_folder action verbatim (would try to create N duplicates); instead
      // turn targets into "move to pending" — execute's reorder+dedup handles
      // it.
      if (source.action === 'new_folder') {
        const name = source.suggestedFolderName?.trim()
        const parent = source.suggestedParentPath?.trim()
        const path = name && parent ? joinFolderPath(parent, name) : undefined
        return {
          ...item,
          action: 'move',
          targetFolderId: path ? `pending:${source.emailId}` : undefined,
          targetFolderPath: path,
          suggestedFolderName: undefined,
          suggestedParentPath: undefined,
          source: newSource,
          ruleId: newRuleId,
        }
      }

      return {
        ...item,
        action: source.action,
        targetFolderId: source.targetFolderId,
        targetFolderPath: source.targetFolderPath,
        suggestedFolderName: source.suggestedFolderName,
        suggestedParentPath: source.suggestedParentPath,
        source: newSource,
        ruleId: newRuleId,
      }
    }
  }

  function propagateFrom(source: PlanItem) {
    // "Format painter": copy source's action+target to all OTHER checked rows.
    const otherIds = new Set([...selectedIds].filter((id) => id !== source.emailId))
    if (otherIds.size === 0) return
    // Mark painted rows userTouched (audit P2): the painter is an EXPLICIT
    // user decision, so same-subject auto-propagation must not quietly
    // overwrite it later. The auto-propagation path in handleItemChange
    // deliberately does NOT set the flag (documented last-write-wins for
    // untouched siblings) — hence the wrapper here instead of baking it
    // into the shared transform.
    const base = buildPropagateTransform(source)
    onBulkApply(otherIds, (item) => ({ ...base(item), userTouched: true }))
    setSelectedIds(new Set())
  }

  // Wrap onChange to auto-propagate to same-subject siblings whenever the
  // user lands on a "complete" state (move with target / delete / skip /
  // fully-specified new_folder). Last-write-wins: subsequent edits override
  // earlier ones across the whole same-subject group.
  function handleItemChange(item: PlanItem) {
    onChange(item)
    const complete =
      item.action === 'delete' ||
      item.action === 'skip' ||
      (item.action === 'move' && !!item.targetFolderPath) ||
      (item.action === 'new_folder' &&
        !!item.suggestedFolderName?.trim() &&
        !!item.suggestedParentPath?.trim())
    if (!complete) return
    const siblingIds = sameSubjectOtherIds(item)
    if (siblingIds.length === 0) return
    // Skip siblings the user has already explicitly edited — they made a
    // deliberate decision we shouldn't quietly overwrite. Bulk apply still
    // hits unedited siblings (which is the common case the auto-propagate
    // is for: AI dropped the whole same-subject group into the same target,
    // user adjusts one, the rest should follow).
    const itemsById = new Map(items.map((i) => [i.emailId, i]))
    const targetIds = new Set(
      siblingIds.filter((id) => !itemsById.get(id)?.userTouched),
    )
    if (targetIds.size === 0) return
    onBulkApply(targetIds, buildPropagateTransform(item))
    setPropagateToast({ count: targetIds.size, key: Date.now() })
  }

  // Bulk 「移到…」（UI/UX 檢討 2026-07）— 勾選後直接選資料夾套用，
  // 取代舊的四步隱藏流程。userTouched 同批次刪除/保留的理由。
  const [bulkMoveOpen, setBulkMoveOpen] = useState(false)
  useEffect(() => {
    // 勾選清空時收起 picker，避免殘留一個沒有對象的選擇器。
    if (selectedIds.size === 0) setBulkMoveOpen(false)
  }, [selectedIds])
  function applyMoveTo(node: MailFolderNode) {
    onBulkApply(selectedIds, (item) => ({
      ...item,
      action: 'move' as const,
      targetFolderId: node.id,
      targetFolderPath: node.path,
      suggestedFolderName: undefined,
      suggestedParentPath: undefined,
      source: item.source === 'rule' ? 'ai' : item.source,
      ruleId: item.source === 'rule' ? undefined : item.ruleId,
      userTouched: true,
    }))
    setBulkMoveOpen(false)
    setSelectedIds(new Set())
  }

  // Bulk 全部刪除/全部保留 set userTouched (audit P2): these are explicit
  // user decisions — without the flag, a later same-subject auto-propagation
  // could silently flip a bulk-保留 row back to 刪除.
  function applyDelete() {
    onBulkApply(selectedIds, (item) => ({
      ...item,
      action: 'delete' as const,
      targetFolderId: undefined,
      targetFolderPath: undefined,
      suggestedFolderName: undefined,
      suggestedParentPath: undefined,
      source: item.source === 'rule' ? 'ai' : item.source,
      ruleId: item.source === 'rule' ? undefined : item.ruleId,
      userTouched: true,
    }))
    setSelectedIds(new Set())
  }

  function applySkip() {
    onBulkApply(selectedIds, (item) => ({
      ...item,
      action: 'skip' as const,
      targetFolderId: undefined,
      targetFolderPath: undefined,
      suggestedFolderName: undefined,
      suggestedParentPath: undefined,
      source: item.source === 'rule' ? 'ai' : item.source,
      ruleId: item.source === 'rule' ? undefined : item.ruleId,
      userTouched: true,
    }))
    setSelectedIds(new Set())
  }

  const movableValid = items.filter(
    (i) =>
      (i.action === 'move' && i.targetFolderPath) ||
      i.action === 'delete' ||
      (i.action === 'new_folder' && i.suggestedFolderName && i.suggestedParentPath),
  ).length

  // 「繼續歸檔」after the last batch can land here when there's nothing left
  // to process. Render a clear empty state instead of a 0-row plan UI.
  if (items.length === 0 && !aiPending) {
    return (
      <div className="space-y-3">
        {banner && (
          <div className="flex items-start gap-2 rounded-md border border-amber-200 bg-amber-50 p-2.5 text-xs">
            <AlertTriangle className="size-3.5 mt-0.5 text-amber-700 shrink-0" />
            <div className="flex-1" title={banner.code}>
              <div className="font-medium text-amber-900">{bannerTitle(banner.code)}</div>
              <div className="text-amber-800">{banner.message}</div>
            </div>
          </div>
        )}
        <div className="rounded-md border border-emerald-200 bg-emerald-50 p-6 text-center space-y-2">
          <CheckCircle2 className="size-8 text-emerald-700 mx-auto" />
          <div className="text-sm font-medium text-emerald-900">信箱已清空</div>
          <div className="text-xs text-emerald-800">沒有更多郵件需要歸類</div>
        </div>
        <Button onClick={onBack} className="w-full">回主畫面</Button>
      </div>
    )
  }

  return (
    <div className="space-y-3">
      {/* 前導區合併（UI/UX 檢討 2026-07）：舊版清單上方可同時疊 AI 進度卡
          ＋amber banner＋toast＋摘要卡（含 token telemetry）四塊，600px
          彈窗第一封信被推到 fold 之下。現在：AI 進度併入摘要卡一行、
          banner 中文標題、toast 改浮層不佔版面、token/chunk 遙測移除
          （成本統計收進 title tooltip）。 */}
      <div className="rounded-md border border-border bg-card p-3 space-y-2">
        <div className="flex items-baseline justify-between text-xs">
          <div className="flex flex-wrap items-baseline gap-2">
            <span className="font-medium">共 {items.length} 封</span>
            {(filterChips.length > 0 || attentionOnly) && (
              <span className="text-muted-foreground">
                · 顯示 <span className="tabular-nums">{filteredItems.length}</span>
              </span>
            )}
            <Badge variant="success" className="gap-1">移 {counts.move}</Badge>
            <Badge variant="danger" className="gap-1">刪 {counts.delete}</Badge>
            <Badge variant="warning" className="gap-1">新 {counts.new_folder}</Badge>
            {counts.skip > 0 && <Badge variant="muted">留 {counts.skip}</Badge>}
            {attentionCount > 0 && (
              <button
                type="button"
                onClick={() => setAttentionOnly((v) => !v)}
                className={cn(
                  'rounded-full border px-2 py-0.5 text-[10px] font-medium transition-colors',
                  attentionOnly
                    ? 'border-amber-500 bg-amber-100 text-amber-900'
                    : 'border-amber-300 bg-amber-50 text-amber-800 hover:bg-amber-100',
                )}
                title={
                  attentionOnly
                    ? '目前只顯示待確認 — 點擊恢復全部'
                    : '只看 AI 沒把握、需要你決定的郵件'
                }
              >
                待確認 {attentionCount}
              </button>
            )}
            {items.length > 2 && (
              <button
                type="button"
                onClick={() => setSortMode((m) => (m === 'byTarget' ? 'default' : 'byTarget'))}
                className={cn(
                  'inline-flex items-center gap-1 rounded-full border px-2 py-0.5 text-[10px] font-medium transition-colors',
                  sortMode === 'byTarget'
                    ? 'border-foreground bg-foreground text-background'
                    : 'border-border text-muted-foreground hover:bg-accent',
                )}
                title={
                  sortMode === 'byTarget'
                    ? '目前依目標資料夾分組 — 點擊恢復原順序'
                    : '把要進同一個資料夾的信聚在一起，方便逐夾核對'
                }
              >
                <FolderTree className="size-3" />
                依資料夾
              </button>
            )}
          </div>
          <span
            className="font-mono text-[10px] text-muted-foreground tabular-nums shrink-0"
            title={
              usage
                ? `本批 token：輸入 ${usage.inputTokens} · 輸出 ${usage.outputTokens}` +
                  (usage.cacheReadTokens > 0 ? ` · 快取命中 ${usage.cacheReadTokens}` : '') +
                  `\n費用為估算（依模型 list price 概算，非帳單）`
                : undefined
            }
          >
            規則 {summary.ruleHits} · AI {summary.aiHandled}
            {usage && (
              <span className="text-muted-foreground/70">
                {' · 約 '}
                {formatUsdApprox(estimateUsageCostUsd(usage, model ?? ''))}
              </span>
            )}
          </span>
          <button
            type="button"
            onClick={() => setFilterVisible(true)}
            className="text-muted-foreground hover:text-foreground shrink-0 ml-2"
            aria-label="過濾清單"
            title="過濾清單（快捷鍵 /）"
          >
            <Search className="size-3.5" />
          </button>
        </div>
        {aiPending && (
          <div
            className="space-y-1"
            title="popup 可關 — 背景會繼續、回來會自動接著看"
          >
            <div className="flex items-center gap-2 text-[11px]">
              <Loader2 className="size-3 animate-spin shrink-0" />
              <span className="text-muted-foreground">
                AI 分類中 {aiPending.completedEmails}/{aiPending.totalEmails} — 已完成的可先審閱
              </span>
            </div>
            <div className="h-1 w-full overflow-hidden rounded-full bg-muted">
              <div
                className="h-full bg-foreground transition-all duration-300"
                style={{
                  width: `${
                    aiPending.totalEmails > 0
                      ? Math.round((aiPending.completedEmails / aiPending.totalEmails) * 100)
                      : 0
                  }%`,
                }}
              />
            </div>
          </div>
        )}
      </div>
      {banner && (
        <div className="flex items-start gap-2 rounded-md border border-amber-200 bg-amber-50 p-2.5 text-xs">
          <AlertTriangle className="size-3.5 mt-0.5 text-amber-700 shrink-0" />
          <div className="flex-1" title={banner.code}>
            <div className="font-medium text-amber-900">{bannerTitle(banner.code)}</div>
            <div className="text-amber-800">{banner.message}</div>
          </div>
          <button type="button" onClick={onDismissBanner} className="text-amber-700 hover:text-amber-900">
            <XCircle className="size-3.5" />
          </button>
        </div>
      )}
      {propagateToast && (
        <div
          key={propagateToast.key}
          className="fixed bottom-20 inset-x-4 z-40 flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-xs text-emerald-900 shadow-md animate-in fade-in slide-in-from-bottom-1 duration-200"
        >
          <Sparkles className="size-3.5 text-emerald-700 shrink-0" />
          <span className="flex-1">
            已同步套用到 <span className="font-mono tabular-nums font-medium">{propagateToast.count}</span> 件同主旨郵件
          </span>
          <button
            type="button"
            onClick={() => setPropagateToast(null)}
            className="text-emerald-700 hover:text-emerald-900"
            aria-label="關閉"
          >
            <XCircle className="size-3.5" />
          </button>
        </div>
      )}

      {(filterVisible || filterChips.length > 0) && (
        <div
          className={cn(
            'flex items-center gap-1.5 flex-wrap rounded-md border bg-background px-2.5 py-1.5 text-xs',
            filterChips.length > 0 ? 'border-foreground/40' : 'border-border',
          )}
        >
          <Search className="size-3.5 text-muted-foreground shrink-0" />
          {filterChips.map((chip) => (
            <span
              key={chip}
              className="inline-flex items-center gap-1 rounded bg-muted px-1.5 py-0.5 text-[11px]"
            >
              <span className="font-mono">{chip}</span>
              <button
                type="button"
                onClick={() =>
                  setFilterChips((prev) => prev.filter((c) => c !== chip))
                }
                className="text-muted-foreground hover:text-foreground"
                aria-label={`移除過濾條件 ${chip}`}
              >
                <X className="size-3" />
              </button>
            </span>
          ))}
          <input
            ref={filterInputRef}
            value={filterDraft}
            onChange={(e) => setFilterDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ',' || e.key === '，') {
                e.preventDefault()
                commitFilterDraft()
              } else if (
                e.key === 'Backspace' &&
                filterDraft === '' &&
                filterChips.length > 0
              ) {
                e.preventDefault()
                setFilterChips((prev) => prev.slice(0, -1))
              } else if (e.key === 'Escape') {
                e.preventDefault()
                if (filterDraft) {
                  setFilterDraft('')
                } else {
                  clearFilter()
                }
              }
            }}
            onBlur={commitFilterDraft}
            placeholder={
              filterChips.length > 0
                ? '繼續加條件…'
                : '過濾主旨 / 寄件人 / 路徑(Enter 加條件、Esc 關閉)'
            }
            className="flex-1 min-w-[120px] bg-transparent outline-none text-xs"
            aria-label="過濾條件"
          />
          <span className="text-[10px] text-muted-foreground tabular-nums shrink-0">
            {filteredItems.length}/{items.length}
          </span>
          <button
            type="button"
            onClick={clearFilter}
            className="text-muted-foreground hover:text-foreground shrink-0"
            aria-label="關閉過濾"
            title="清除並關閉(Esc)"
          >
            <X className="size-3.5" />
          </button>
        </div>
      )}

      {selectedIds.size > 0 && (
        <div className="rounded-md border-2 border-foreground bg-card p-2.5 text-xs space-y-2">
          <div className="flex items-center justify-between gap-2 flex-wrap">
            <span className="font-medium">已選 {selectedIds.size} 項</span>
            <div className="flex items-center gap-1.5 flex-wrap">
              {/* 「移到…」直接上批次列（UI/UX 檢討 2026-07）：批次移動是
                  勾選後最常見的期待，舊版要走「展開任一張卡→卡內選目標→
                  按全部套用」四步隱藏流程、還得靠一段常駐教學文字補救。
                  展開卡內的「全部套用到勾選」保留作進階格式刷。 */}
              <Button
                size="sm"
                variant={bulkMoveOpen ? 'default' : 'outline'}
                onClick={() => setBulkMoveOpen((v) => !v)}
              >
                移到…
              </Button>
              {/* 全部刪除降為 outline-destructive：實心紅留給真正的警示。
                  舊版它與底部轉紅的執行鈕形成兩顆意義不同的紅色主鈕。 */}
              <Button
                size="sm"
                variant="outline"
                className="border-red-300 text-red-700 hover:bg-red-50"
                onClick={applyDelete}
              >
                全部刪除
              </Button>
              <Button size="sm" variant="outline" onClick={applySkip}>
                全部保留
              </Button>
              <button
                type="button"
                onClick={() => setSelectedIds(new Set(filteredItems.map((i) => i.emailId)))}
                className="text-[10px] text-muted-foreground hover:underline px-1"
              >
                {filterChips.length > 0
                  ? `全選顯示中 (${filteredItems.length})`
                  : `全選 (${items.length})`}
              </button>
              <button
                type="button"
                onClick={() => setSelectedIds(new Set())}
                className="text-[10px] text-muted-foreground hover:underline px-1"
              >
                清除
              </button>
            </div>
          </div>
          {bulkMoveOpen && (
            <div className="space-y-1">
              <label className="text-[10px] text-muted-foreground">
                目標資料夾（套用到勾選的 {selectedIds.size} 項）
              </label>
              <FolderPicker
                tree={augmentedTree}
                excludePrefixes={excludePrefixes}
                onSelect={applyMoveTo}
              />
            </div>
          )}
        </div>
      )}

      {filterChips.length > 0 && filteredItems.length === 0 ? (
        <div className="rounded-md border border-dashed border-border bg-muted/30 p-6 text-center space-y-2 text-xs">
          <Search className="size-6 text-muted-foreground mx-auto" />
          <div className="text-sm font-medium">沒有符合過濾條件的郵件</div>
          <div className="text-muted-foreground">
            條件:{filterChips.join(' · ')}
          </div>
          <Button size="sm" variant="outline" onClick={clearFilter}>
            清除過濾
          </Button>
        </div>
      ) : (
        // Single list in both narrow and wide layouts. The earlier
        // wide-mode side trace panel duplicated the row's expanded
        // RuleTraceBlock + "停用此規則" button — pure visual noise once
        // inline expand existed. PlanRow's own layout uses useLayout()
        // to widen its content in wide mode, so we don't need a
        // different render here.
        <ul className="space-y-1.5">
          {displayItems.map((it, idx) => {
            const cid = it.conversationId
            const group = cid ? conversationGroups.get(cid) : undefined
            // This item is the rep of a unified conversation group?
            const isRepOfGroup = group && group.rep.emailId === it.emailId
            const isCollapsed =
              isRepOfGroup && !expandedConversations.has(cid!)
            // 依資料夾分組 (P1-2)：sortMode='byTarget' 時，group key 與前一
            // 列不同就先插一條群組分隔列（路徑 + 該資料夾總筆數 + 全選此
            // 組）。群組計數取自 targetGroups（涵蓋被折疊的同對話 siblings），
            // 全選也把隱藏 siblings 一起勾。
            const gkey = sortMode === 'byTarget' ? targetGroupKey(it) : null
            const showGroupHeader =
              gkey !== null &&
              (idx === 0 || targetGroupKey(displayItems[idx - 1]!) !== gkey)
            const groupSize = gkey !== null ? (targetGroups.get(gkey)?.length ?? 0) : 0
            // F15 (2026-06-03): focus ring keys off the displayItems
            // index (the rendered position), matching keyboard nav which
            // now also indexes displayItems. Previously this used the
            // filteredItems index, so a cursor on a hidden sibling
            // matched no rendered row and the ring vanished.
            return (
              <Fragment key={it.emailId}>
                {showGroupHeader && (
                  <li className="pt-2 first:pt-0">
                    <div className="flex items-center gap-2 rounded bg-muted/50 px-2 py-1">
                      <FolderTree className="size-3 text-muted-foreground shrink-0" />
                      <span
                        className="font-mono text-[11px] font-medium truncate"
                        title={gkey!}
                      >
                        {gkey}
                      </span>
                      <span className="text-[10px] text-muted-foreground tabular-nums shrink-0">
                        {groupSize} 封
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          setSelectedIds((prev) => {
                            const next = new Set(prev)
                            for (const id of targetGroups.get(gkey!) ?? []) next.add(id)
                            return next
                          })
                        }
                        className="ml-auto shrink-0 text-[10px] text-muted-foreground hover:text-foreground hover:underline"
                        title="勾選所有要進這個資料夾的信"
                      >
                        全選此組
                      </button>
                    </div>
                  </li>
                )}
                <li className="contents">
                <PlanRow
                  item={it}
                  tree={augmentedTree}
                  validPaths={validPaths}
                  excludePrefixes={excludePrefixes}
                  selected={selectedIds.has(it.emailId)}
                  otherSelectedCount={
                    selectedIds.size - (selectedIds.has(it.emailId) ? 1 : 0)
                  }
                  focused={focusedIndex === idx}
                  scrollIntoViewToken={
                    focusedIndex === idx ? scrollIntoViewToken : undefined
                  }
                  toggleExpandToken={
                    focusedIndex === idx ? toggleExpandToken : undefined
                  }
                  onChange={handleItemChange}
                  onToggleSelect={() => toggleSelect(it.emailId)}
                  onPropagate={() => propagateFrom(it)}
                  onToggleRule={onToggleRule}
                  onActivate={() => setFocusedIndex(idx)}
                />
                {/* Conversation cluster hint: shown directly under the
                    rep when this group has siblings. Collapsed = "+ N
                    封同對話 [展開]"; expanded = "↑ 收合 N 封同對話"
                    rendered after the last sibling (handled below by
                    checking idx vs siblings boundary). */}
                {isRepOfGroup && isCollapsed && (
                  <button
                    type="button"
                    onClick={() => toggleConversation(cid!)}
                    className="ml-7 mt-0 text-[11px] text-muted-foreground hover:text-foreground underline underline-offset-2 inline-flex items-center gap-1 py-1"
                    title="同對話的其他信件已折疊（決策相同：移到相同資料夾）"
                  >
                    + 還有 {group!.siblings.length} 封同對話 — 展開檢視
                  </button>
                )}
                {/* After the LAST sibling of an expanded group, show
                    collapse button. Sibling rows themselves carry no
                    extra UI. */}
                {!isRepOfGroup && cid && expandedConversations.has(cid) && (() => {
                  const group2 = conversationGroups.get(cid)
                  if (!group2) return null
                  const nextItem = displayItems[idx + 1]
                  const isLastSibling =
                    !nextItem || nextItem.conversationId !== cid
                  if (!isLastSibling) return null
                  return (
                    <button
                      type="button"
                      onClick={() => toggleConversation(cid)}
                      className="ml-7 text-[11px] text-muted-foreground hover:text-foreground underline underline-offset-2 inline-flex items-center gap-1 py-1"
                    >
                      ↑ 收合 {group2.siblings.length + 1} 封同對話
                    </button>
                  )
                })()}
                </li>
              </Fragment>
            )
          })}
        </ul>
      )}

      {/* 底欄減層（UI/UX 檢討 2026-07）：舊版最多疊三層（刪除警示條＋
          鍵盤提示＋按鈕列）。刪除警示併入按鈕列小字＋執行鈕 label；執行鈕
          固定 default（不再因含刪除而轉紅 — 舊版和批次列的紅色刪除鈕形成
          兩顆意義不同的紅鈕，主動作反而失去辨識度）。 */}
      <div className="sticky bottom-0 -mx-4 px-4 py-3 border-t border-border bg-background/95 backdrop-blur space-y-2">
        {!shortcutsHintDismissed && selectedIds.size === 0 && !aiPending && (
          <div className="flex items-center justify-between gap-2 text-[10px] text-muted-foreground">
            <span>
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">j</kbd>/
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">k</kbd> 上下 ·
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">x</kbd> 勾選 ·
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">Enter</kbd> 展開 ·
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">d</kbd>/
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">s</kbd> 刪/留 ·
              <kbd className="font-mono px-1 mx-0.5 bg-muted rounded">/</kbd> 過濾
            </span>
            <button
              type="button"
              className="text-muted-foreground hover:text-foreground"
              onClick={dismissShortcutsHint}
              aria-label="關閉鍵盤提示（不再顯示）"
              title="關閉後不再顯示；按 ? 可隨時查快捷鍵"
            >
              ×
            </button>
          </div>
        )}
        <div className="flex items-center justify-between gap-2">
          <Button variant="ghost" onClick={onBack}>取消</Button>
          {counts.delete > 0 && (
            <span
              className="text-[10px] text-amber-800 text-right flex-1"
              title="刪除的郵件可從 Outlook「已刪除的郵件」資料夾回復"
            >
              含 {counts.delete} 件刪除・可回復
            </span>
          )}
          {/* 刀3：AI 分類中不再鎖死執行 — 已完成的（規則命中＋已分類的
              AI 項）可先執行，剩餘的分類繼續在背景跑，完成後從「繼續歸檔」
              接手（consume 路徑會濾掉已執行的信）。 */}
          {aiPending ? (
            <Button
              variant="outline"
              onClick={onConfirmExecute}
              disabled={movableValid === 0}
              title="AI 還在分類剩餘郵件 — 先執行目前清單，分類完成後可從「繼續歸檔」接續處理其餘郵件"
            >
              <Play /> 先執行 {movableValid} 個
            </Button>
          ) : (
            <Button onClick={onConfirmExecute} disabled={movableValid === 0}>
              <Play /> 執行 {movableValid} 個
            </Button>
          )}
        </div>
      </div>
    </div>
  )
}

// ---- Executing ------------------------------------------------------------

const ITEM_STATUS_META: Record<ExecuteItemStatus, { icon: string; tone: string }> = {
  queued: { icon: '○', tone: 'text-muted-foreground' },
  processing: { icon: '◌', tone: 'text-foreground' },
  moved: { icon: '✓', tone: 'text-emerald-700' },
  deleted: { icon: '✗', tone: 'text-red-700' },
  folder_created: { icon: '✓', tone: 'text-amber-700' },
  skipped: { icon: '—', tone: 'text-muted-foreground' },
  cancelled: { icon: '⊘', tone: 'text-muted-foreground' },
  error: { icon: '!', tone: 'text-red-700' },
}

const ACTION_LABEL_SHORT: Record<PlanAction, string> = {
  move: '移',
  delete: '刪',
  new_folder: '新',
  skip: '留',
}

function ExecutingScreen({
  state,
  onCancel,
}: {
  state: ExecuteState
  onCancel: () => void
}) {
  const pct = state.total > 0 ? Math.round((state.current / state.total) * 100) : 0
  // Arm the cancel button only after a short beat (audit P2): PlanScreen's
  // 執行 and this screen's 取消 render in the same sticky-bar spot, so the
  // second click of a double-click on 執行 landed HERE and cancelled the
  // batch the first click just started. 600ms comfortably outlasts a
  // double-click without delaying a real cancel decision.
  const [cancelArmed, setCancelArmed] = useState(false)
  useEffect(() => {
    const t = window.setTimeout(() => setCancelArmed(true), 600)
    return () => window.clearTimeout(t)
  }, [])
  return (
    <div className="space-y-3">
      <div className="rounded-md border border-border bg-card p-3 space-y-2">
        <div className="flex items-baseline justify-between text-xs">
          <div className="font-medium">
            {state.cancelRequested ? '取消中…' : '執行中'}
            <span className="ml-2 text-muted-foreground tabular-nums">
              {state.current} / {state.total}
            </span>
          </div>
          <span className="font-mono text-muted-foreground tabular-nums">{pct}%</span>
        </div>
        <div className="h-1.5 w-full overflow-hidden rounded-full bg-muted">
          <div
            className="h-full bg-foreground transition-all duration-300"
            style={{ width: `${pct}%` }}
          />
        </div>
        <div className="flex gap-1.5 text-[10px]">
          <Badge variant="success">移 {state.summary.moved}</Badge>
          <Badge variant="danger">刪 {state.summary.deleted}</Badge>
          <Badge variant="warning">新 {state.summary.foldersCreated}</Badge>
          {state.summary.errors > 0 && <Badge variant="danger">錯 {state.summary.errors}</Badge>}
        </div>
      </div>

      <ul className="space-y-1 max-h-72 overflow-y-auto">
        {state.results.map((r) => {
          const meta = ITEM_STATUS_META[r.status]
          return (
            <li key={r.emailId} className="flex items-start gap-2 text-xs px-2 py-1.5 border-b border-border/40 last:border-0">
              <span className={cn('font-mono w-4 shrink-0', meta.tone)}>{meta.icon}</span>
              <span className="text-[10px] text-muted-foreground font-mono shrink-0 mt-0.5">[{ACTION_LABEL_SHORT[r.action]}]</span>
              <span className="flex-1 truncate">{r.subject || '（無主旨）'}</span>
              {r.message && (
                <span className="text-[10px] text-red-700 truncate max-w-[160px]" title={r.message}>
                  {r.message.slice(0, 40)}
                </span>
              )}
            </li>
          )
        })}
      </ul>

      <div className="sticky bottom-0 -mx-4 px-4 py-3 border-t border-border bg-background/95 backdrop-blur flex items-center justify-between">
        <span className="text-xs text-muted-foreground">popup 可關 — 背景會繼續</span>
        <Button variant="outline" onClick={onCancel} disabled={state.cancelRequested || !cancelArmed}>
          <Pause /> {state.cancelRequested ? '取消中…' : '取消'}
        </Button>
      </div>
    </div>
  )
}

// ---- Done -----------------------------------------------------------------

type UndoResultPayload = {
  attempted: number
  restored: number
  failed: number
  errors: Array<{ subject: string; message: string }>
}

// Compact rule-type labels for the DoneScreen learned-rules list (覆核 P1-3).
const LEARNED_TYPE_LABEL: Record<string, string> = {
  domain: '網域',
  sender: '寄件人',
  case_code: '案號',
  subject_keyword: '主旨',
  compound: '組合',
}

function DoneScreen({
  state,
  batchSize,
  onReset,
  onRetry,
  onContinue,
}: {
  state: ExecuteState
  batchSize: number
  onReset: () => void
  onRetry: () => void
  onContinue: () => void
}) {
  const elapsedSec = state.finishedAt ? Math.round((state.finishedAt - state.startedAt) / 1000) : 0
  const errored = state.results.filter((r) => r.status === 'error')
  const canceled = state.cancelRequested

  // Peek the next batch on mount so 「繼續歸檔下一批」isn't a misleading
  // default action when inbox is empty after this batch (the common case
  // for a 50-email-per-batch workflow). Lightweight: one Graph API call
  // with select=Id,Flag, applies the same skip-history + skipFlagged
  // filters preflight uses. peekResult === null = still loading; show
  // button optimistically. Failure also leaves peekResult null → falls
  // through to legacy behaviour, never blocks the user.
  const [peekResult, setPeekResult] = useState<{
    eligibleCount: number
    cappedAtBatchSize: boolean
  } | null>(null)
  useEffect(() => {
    let cancelled = false
    void send<{ eligibleCount: number; cappedAtBatchSize: boolean }>({
      type: 'peekNextBatch',
      batchSize,
    }).then((r) => {
      if (cancelled) return
      if (r.ok && r.data) setPeekResult(r.data)
    })
    return () => {
      cancelled = true
    }
  }, [batchSize])

  // Undo window state. Loads once on mount; ticks every second to drive the
  // countdown; transitions to running/done/failed once the user acts.
  const [undoSnap, setUndoSnap] = useState<UndoSnapshot | null>(null)
  const [undoStage, setUndoStage] = useState<'idle' | 'running' | 'done' | 'failed'>('idle')
  const [undoResult, setUndoResult] = useState<UndoResultPayload | null>(null)
  const [undoError, setUndoError] = useState<string | null>(null)
  const [now, setNow] = useState(() => Date.now())

  useEffect(() => {
    let cancelled = false
    void send<{ snapshot: UndoSnapshot | null }>({ type: 'getUndoSnapshot' }).then((r) => {
      if (cancelled) return
      if (r.ok) setUndoSnap(r.data?.snapshot ?? null)
    })
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (!undoSnap || undoStage !== 'idle') return
    const id = setInterval(() => setNow(Date.now()), 1000)
    return () => clearInterval(id)
  }, [undoSnap, undoStage])

  const remainingMs = undoSnap ? Math.max(0, undoSnap.expiresAt - now) : 0
  const showUndoBanner = !!undoSnap && remainingMs > 0 && undoStage === 'idle'

  async function performUndo() {
    if (!undoSnap) return
    setUndoStage('running')
    setUndoError(null)
    const r = await send<UndoResultPayload>({ type: 'executeUndo' })
    if (r.ok) {
      const data = r.data!
      setUndoResult(data)
      setUndoStage(data.failed > 0 ? 'failed' : 'done')
    } else {
      setUndoError(r.message)
      setUndoStage('failed')
    }
  }

  async function dismissUndoLocally() {
    await send({ type: 'dismissUndo' })
    setUndoSnap(null)
  }

  // Keyboard shortcut: `u` fires the undo when the banner is visible. Same
  // input-guard rules as PlanScreen — don't hijack typing or modifier combos.
  useEffect(() => {
    if (!showUndoBanner) return
    function onKey(e: KeyboardEvent) {
      const t = e.target as HTMLElement | null
      const tag = t?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return
      if (e.metaKey || e.ctrlKey || e.altKey) return
      if (e.key.toLowerCase() === 'u') {
        e.preventDefault()
        void performUndo()
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [showUndoBanner])

  return (
    <div className="space-y-4">
      <div className={cn(
        'rounded-md border p-4 space-y-2',
        errored.length === 0 && !canceled ? 'border-emerald-200 bg-emerald-50' : 'border-amber-200 bg-amber-50',
      )}>
        <div className="flex items-center gap-2">
          {errored.length === 0 && !canceled ? (
            <CheckCircle2 className="size-5 text-emerald-700" />
          ) : (
            <AlertTriangle className="size-5 text-amber-700" />
          )}
          <h2 className="text-sm font-semibold">
            {canceled ? '已取消' : errored.length === 0 ? '全部完成' : `完成（${errored.length} 個錯誤）`}
          </h2>
        </div>
        <div className="text-xs space-y-0.5 pl-7">
          <div>耗時 <span className="font-mono">{elapsedSec}s</span></div>
          <div>
            移動 {state.summary.moved} · 刪除 {state.summary.deleted} · 新建 {state.summary.foldersCreated} · 保留 {state.summary.skipped}
          </div>
          {state.rulesAdded > 0 && (
            <div className="text-emerald-800">
              <div className="flex items-center gap-1">
                <Sparkles className="size-3 shrink-0" />
                <span>學會 {state.rulesAdded} 條新規則 — 下次同類郵件自動歸檔</span>
              </div>
              {state.rulesAddedDetail && state.rulesAddedDetail.length > 0 && (
                <details className="mt-1 ml-4">
                  <summary className="cursor-pointer select-none text-[11px] text-emerald-700 hover:underline">
                    看是哪幾條規則 →
                  </summary>
                  <ul className="mt-1 space-y-0.5">
                    {state.rulesAddedDetail.map((d, i) => (
                      <li
                        key={i}
                        className="text-[10px] text-emerald-900/85 truncate"
                        title={`${d.signal} → ${d.targetFolderPath}`}
                      >
                        <span className="text-emerald-700">
                          {d.source === 'ai_overridden' ? '你改的' : 'AI'}·
                          {LEARNED_TYPE_LABEL[d.type] ?? d.type}
                        </span>{' '}
                        <span className="font-mono">{d.signal}</span> →{' '}
                        <span className="font-mono">{d.targetFolderPath}</span>
                      </li>
                    ))}
                  </ul>
                  <button
                    type="button"
                    onClick={() => {
                      void chrome.tabs.create({
                        url: `${chrome.runtime.getURL('src/options/index.html')}#rules-library/all`,
                      })
                    }}
                    className="mt-1 text-[10px] text-emerald-700 hover:underline"
                  >
                    在規則庫檢視 / 編輯 →
                  </button>
                </details>
              )}
            </div>
          )}
        </div>
      </div>

      {showUndoBanner && (
        <div className="flex items-center gap-2 text-[11px] text-muted-foreground">
          <button
            type="button"
            onClick={() => void performUndo()}
            className="underline underline-offset-2 hover:text-foreground transition-colors"
            title="鍵盤捷徑:u"
          >
            撤回 {undoSnap!.moves.length} 件搬移
          </button>
          <span className="font-mono tabular-nums">({Math.ceil(remainingMs / 1000)}s)</span>
          <button
            type="button"
            onClick={() => void dismissUndoLocally()}
            className="ml-auto text-muted-foreground/60 hover:text-foreground"
            aria-label="不撤回"
          >
            ×
          </button>
        </div>
      )}

      {undoStage === 'running' && (
        <div className="rounded-md border border-border bg-muted/40 p-2.5 text-xs flex items-center gap-2">
          <Loader2 className="size-4 animate-spin" />
          撤回中…
        </div>
      )}

      {undoStage === 'done' && undoResult && (
        <div className="rounded-md border border-emerald-200 bg-emerald-50 p-2.5 text-xs flex items-center gap-1.5">
          <CheckCircle2 className="size-4 text-emerald-700" />
          已撤回 {undoResult.restored} 封郵件回到收件夾
        </div>
      )}

      {undoStage === 'failed' && (
        <div className="rounded-md border border-red-200 bg-red-50 p-2.5 text-xs space-y-1">
          {undoResult ? (
            <>
              <div className="font-medium text-red-900">
                撤回部分失敗:成功 {undoResult.restored} 件 / 失敗 {undoResult.failed} 件
              </div>
              {undoResult.errors.slice(0, 3).map((e, i) => (
                <div key={i} className="text-[10px] text-red-800 truncate">
                  • {e.subject || '(無主旨)'}: {e.message}
                </div>
              ))}
              {undoResult.errors.length > 3 && (
                <div className="text-[10px] text-red-700">…及其他 {undoResult.errors.length - 3} 件</div>
              )}
            </>
          ) : (
            <div className="text-red-800">{undoError ?? '撤回失敗、原因未知'}</div>
          )}
        </div>
      )}

      {errored.length > 0 && (
        <div className="space-y-1.5">
          <div className="text-xs font-medium text-red-900">錯誤明細</div>
          <ul className="space-y-1">
            {errored.map((r) => (
              <li key={r.emailId} className="rounded-md border border-red-200 bg-red-50/50 p-2 text-xs space-y-0.5">
                <div className="font-medium truncate">{r.subject}</div>
                <div className="text-red-800 text-[10px]">{r.message}</div>
              </li>
            ))}
          </ul>
        </div>
      )}

      {errored.length > 0 && (
        <Button onClick={onRetry} className="w-full" variant="default">
          <RotateCcw /> 重試 {errored.length} 個錯誤
        </Button>
      )}
      {/*
       * Three render modes:
       *   1. peek === null (still loading)  → show continue + back (legacy)
       *   2. eligibleCount === 0            → empty-state hint, 回主畫面 only
       *   3. eligibleCount > 0              → continue with count, back secondary
       */}
      {peekResult?.eligibleCount === 0 ? (
        <>
          <div className="rounded-md border border-emerald-200 bg-emerald-50 p-3 text-center text-xs text-emerald-900">
            <CheckCircle2 className="size-4 text-emerald-700 mx-auto mb-1" />
            信箱已清空,沒有更多郵件需要歸類
          </div>
          <Button onClick={onReset} className="w-full">
            回主畫面
          </Button>
        </>
      ) : (
        <div className="flex gap-2">
          <Button
            onClick={onContinue}
            className="flex-1"
            variant={errored.length > 0 ? 'outline' : 'default'}
          >
            <Play />
            {peekResult && peekResult.eligibleCount > 0
              ? `繼續歸檔下一批 (還有 ${peekResult.eligibleCount}${
                  peekResult.cappedAtBatchSize ? '+' : ''
                } 封)`
              : '繼續歸檔下一批'}
          </Button>
          <Button onClick={onReset} variant="ghost">
            回主畫面
          </Button>
        </div>
      )}
    </div>
  )
}

// ---- Error ---------------------------------------------------------------

function ErrorScreen({ code, message, onBack }: { code: string; message: string; onBack: () => void }) {
  return (
    <div className="space-y-3">
      <div className="rounded-md border border-amber-300 bg-amber-50 p-3 text-xs space-y-1.5">
        <div className="flex items-center gap-2 font-medium text-amber-900">
          <XCircle className="size-4" />
          {code}
        </div>
        <div className="text-amber-800 break-words">{message}</div>
      </div>
      <Button variant="outline" onClick={onBack}>返回</Button>
    </div>
  )
}
