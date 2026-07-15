import { Fragment, memo, type ChangeEvent, type KeyboardEvent as ReactKeyboardEvent, type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { cn } from '@/lib/utils'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { FolderPicker } from '@/popup/components/FolderPicker'
import { AlertTriangle, BarChart3, CheckCircle2, ChevronDown, Download, KeyRound, ListChecks, Loader2, Mail, Pencil, Play, Plus, Power, PowerOff, Radar, RotateCcw, Search, Settings as SettingsIcon, ShieldAlert, Trash2, Upload, X } from 'lucide-react'
import {
  decodeCompound,
  encodeCompound,
  extractCaseCodes,
  extractCourtCaseNumbers,
  type CompoundCondition,
} from '@/shared/rules'
import { computeRuleHealth } from '@/shared/rule-health'
import { initParentBridge, postToParent } from '@/shared/parent-bridge'
import {
  browserLabel,
  detectBrowser,
  syncAccountDescription,
  syncSettingsUrl,
} from '@/shared/browser-detect'
import type { MailFolderNode, Rule, RuleEvent, RuleSource, RuleType, Metrics } from '@/shared/types'

type Ok<T> = { ok: true; data?: T }
type Err = { ok: false; code: string; message: string }

/**
 * Sub-views inside the full-screen Rule Library (2026-05-27 redesign).
 * Mapped to URL hashes `#rules-library/<view>` so users can bookmark
 * specific views (e.g. linking the conflicts inbox).
 */
type RuleLibrarySubView =
  | 'all' // every rule, with filters / search / bulk actions
  | 'conflicts' // current rule conflicts (most are auto-resolved now)
  | 'dormant' // auto-disabled rules (legacy_token, high-error-rate)
  | 'health' // top hits / low accuracy / stale dashboards
  | 'scan' // initial scan tool
  | 'history' // rule edit audit log

const RULE_LIBRARY_SUBVIEWS: RuleLibrarySubView[] = [
  'all',
  'conflicts',
  'dormant',
  'health',
  'scan',
  'history',
]
async function send<T>(req: unknown): Promise<Ok<T> | Err> {
  return chrome.runtime.sendMessage(req)
}

type StatusData = {
  owaConnected: boolean
  tokenValid: boolean
  apiKeyConfigured: boolean
  apiKeyPreview: string | null
  model: string
  rulesCount: number
  excludePrefixes: string[]
  batchSize: number
  aiConfidenceThreshold: number
  skipFlagged: boolean
  showOwaFab: boolean
  prefetchNextBatch: boolean
  recentActivityIncludePrefixes: string[]
  recentActivityIncludeLeafNames: string[]
  internalDomains: string[]
  primaryRootPath: string
  internalSubjectCategories: string[]
  aiIncludeFewShotExamples: boolean
  syncEnabled: boolean
  syncMachineId: string
  lastSyncAt: string
}

type ScanItemStatus = 'queued' | 'processing' | 'done' | 'no_domains' | 'empty' | 'error'

type ScanState = {
  inProgress: boolean
  cancelRequested: boolean
  startedAt: number
  finishedAt?: number
  total: number
  current: number
  rootPath: string
  results: Array<{
    folderPath: string
    status: ScanItemStatus
    emailsScanned?: number
    domainsFound?: Array<{ domain: string; count: number; added: boolean }>
    sendersFound?: Array<{ address: string; count: number; added: boolean }>
    rulesAdded?: number
    message?: string
  }>
  summary: {
    foldersScanned: number
    rulesAdded: number
    foldersWithNoExternalDomains: number
    errors: number
  }
}

const MODEL_OPTIONS = [
  { id: 'claude-sonnet-4-6', label: 'Sonnet 4.6（精準、推薦）' },
  { id: 'claude-haiku-4-5-20251001', label: 'Haiku 4.5（快、便宜）' },
  { id: 'claude-opus-4-7', label: 'Opus 4.7（最強、最貴）' },
]

/**
 * Collapsible wrapper for advanced / diagnostic sections — uses native
 * <details> so state is managed by the browser (survives re-renders, no
 * useState overhead). The summary is a clickable strip styled to match
 * other Card titles; the body holds the original Card / Section.
 *
 * Sections wrapped:
 *   - 初始掃描 (one-time setup)
 *   - 規則效益 / 規則健康度 (diagnostic)
 *   - 歷史統計 / 規則編輯紀錄 / 診斷匯出 (read-only data views)
 *
 * Frequently-edited settings (API key, model, batch size, threshold, OWA FAB,
 * skip mechanisms, rules CRUD) stay always-visible.
 */
/**
 * Format an ISO timestamp as a short relative phrase: "剛剛 / 5 分前 /
 * 2 小時前 / 3 天前 / 2026-05-26". Inline rather than importing from
 * popup to avoid coupling between the popup and options bundles.
 */
function formatRelativeTime(iso: string | undefined): string {
  if (!iso) return ''
  const t = Date.parse(iso)
  if (!Number.isFinite(t)) return ''
  const diffMs = Date.now() - t
  const min = Math.max(0, Math.floor(diffMs / 60_000))
  if (min < 1) return '剛剛'
  if (min < 60) return `${min} 分前`
  const hr = Math.floor(min / 60)
  if (hr < 24) return `${hr} 小時前`
  const day = Math.floor(hr / 24)
  if (day < 7) return `${day} 天前`
  const d = new Date(t)
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

/**
 * Turn a raw sync error reason string into a user-friendly hint + CTA.
 *
 * The raw `reason` field comes from doPush / doPull / chrome.storage.sync
 * errors. A lawyer doesn't need to know what schemaVersion is — they
 * need to know what to DO. This mapping covers the common failure modes:
 *
 *   - schema downgrade refused → upgrade extension on this device
 *   - quota exceeded          → clean up sleeping / orphaned rules
 *   - quota write/min hit     → just wait, retry
 *   - schemaVersion mismatch  → upgrade extension
 *   - network / fetch fail    → retry now
 *   - generic 'in flight'     → friendly "already running" message
 *   - fallthrough             → echo the raw reason
 */
function describeSyncError(
  reason: string,
  source: 'push' | 'pull' | 'pull-remote',
): { hint: string; cta?: { label: string; action: 'upgrade' | 'cleanup' | 'retry' } } {
  const r = reason.toLowerCase()
  if (r.includes('schemaversion')) {
    return {
      hint:
        source === 'push'
          ? '雲端規則的版本比這台機器新、無法覆蓋。請先把此瀏覽器的擴充功能升級到最新版本、然後再試一次。'
          : '雲端規則的版本格式跟這台機器不符。請把這台機器的擴充功能升級到最新版本。',
      cta: { label: '到擴充功能管理頁檢查版本', action: 'upgrade' },
    }
  }
  if (r.includes('quota_bytes') || (r.includes('quota') && r.includes('exceed'))) {
    return {
      hint:
        '同步資料量超過 100KB 上限。建議到下方規則庫清掉長期沒命中的「休眠規則」、再重試上傳。',
      cta: { label: '到規則庫清理', action: 'cleanup' },
    }
  }
  if (r.includes('max_write_operations') || r.includes('write_operations_per')) {
    return {
      hint: '同步寫入太頻繁、被瀏覽器暫時限速。稍等一分鐘後會自動恢復。',
      cta: { label: '立即重試', action: 'retry' },
    }
  }
  if (r.includes('network') || r.includes('failed to fetch') || r.includes('offline')) {
    return {
      hint: '網路連線問題、稍後會自動重試。',
      cta: { label: '立即重試', action: 'retry' },
    }
  }
  if (r.includes('in flight') || r.includes('already')) {
    return { hint: '另一次同步正在執行、這次跳過、稍後會再嘗試。' }
  }
  if (r.includes('sync disabled')) {
    return { hint: '同步已停用、所以這次操作被跳過。' }
  }
  if (r.includes('no cloud state')) {
    return { hint: '雲端還沒有資料。第一次啟用同步時、按「立即上傳」推送本機規則。' }
  }
  // Fallback: show raw reason, no CTA.
  return { hint: reason }
}

function CollapsibleSection({
  title,
  icon,
  children,
  defaultOpen = false,
}: {
  title: string
  icon?: ReactNode
  children: ReactNode
  defaultOpen?: boolean
}) {
  return (
    <details className="group" open={defaultOpen}>
      <summary className="cursor-pointer list-none rounded-md border border-border bg-muted/40 hover:bg-muted px-3 py-2 flex items-center gap-2 text-sm font-medium select-none">
        {icon}
        <span className="flex-1">{title}</span>
        <span className="text-[10px] text-muted-foreground group-open:hidden">展開</span>
        <ChevronDown className="size-3.5 text-muted-foreground transition-transform group-open:rotate-180" />
      </summary>
      <div className="mt-2">
        {children}
      </div>
    </details>
  )
}

export default function App() {
  // When the options page is loaded inside the FAB iframe (window.parent !== window),
  // pressing ESC should close the parent modal. Iframe keyboard events don't
  // bubble across the frame boundary, so we have to postMessage out. The FAB
  // content script listens for `mail-organizer/close-options` and clears the
  // overlay. Standalone (chrome.runtime.openOptionsPage) tabs ignore ESC.
  useEffect(() => {
    if (window.parent === window) return // standalone tab
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      // Don't intercept ESC while the user is mid-edit in an input/textarea
      // — that breaks "ESC to clear value" expectations on form fields.
      const t = e.target as HTMLElement | null
      const tag = t?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return
      // If a layered overlay (rule-detail drawer, confirm dialog, actions
      // menu) is open, Esc belongs to IT — bail and let its own handler
      // close just that layer. Their stopImmediatePropagation can't save us
      // here: this App-level listener registers at mount, BEFORE any
      // overlay's listener, and document keydown listeners fire in
      // registration order — so we'd close the whole options iframe first
      // and the drawer's handler would fire into a dead UI.
      if (document.querySelector('[role="dialog"], [role="alertdialog"], [role="menu"]')) return
      postToParent({ type: 'mail-organizer/close-options' })
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [])

  // Handshake with the FAB content script so postToParent learns the
  // parent's origin. No-op in standalone tab mode.
  useEffect(() => initParentBridge(), [])

  const [status, setStatus] = useState<StatusData | null>(null)
  const [keyInput, setKeyInput] = useState('')
  const [keyError, setKeyError] = useState<string | null>(null)
  const [keySaved, setKeySaved] = useState(false)
  // Whether the full API-key controls (preview + input + clear) are
  // visible. When the key is already configured, default to collapsed
  // so the daily-use options page doesn't dump a "manage your secret"
  // prompt at the top. User clicks 「管理」 to expand if they need to
  // change keys; auto-collapses after a successful save.
  const [keyCardExpanded, setKeyCardExpanded] = useState(false)
  const [savingModel, setSavingModel] = useState(false)
  const [batchSavedAt, setBatchSavedAt] = useState<number | null>(null)
  const [scanState, setScanState] = useState<ScanState | null>(null)
  // Initialized empty; an effect below seeds it from settings.primaryRootPath
  // once status loads. New users who haven't completed onboarding see the
  // input empty + a hint pointing them at the firm settings card.
  const [scanRootPath, setScanRootPath] = useState('')
  const [scanError, setScanError] = useState<string | null>(null)
  const [rules, setRules] = useState<Rule[]>([])
  const [conflicts, setConflicts] = useState<Array<{ type: string; signal: string; ruleIds: string[]; targets: string[] }>>([])
  const [tree, setTree] = useState<MailFolderNode[] | null>(null)
  const [metrics, setMetrics] = useState<Metrics | null>(null)
  const [ruleSearch, setRuleSearch] = useState('')
  const [skipCount, setSkipCount] = useState<number>(0)
  const [skipClearedAt, setSkipClearedAt] = useState<number | null>(null)
  // Default to 'rules' on open — that's the section the lawyer touches
  // most (規則庫 CRUD); 連線 is set-and-forget. The scroll-on-mount effect
  // below smooth-scrolls to #rules so the user lands on the right area
  // without an extra click on the sidebar TOC.
  const [activeSection, setActiveSection] = useState<string>('rules')
  // Full-screen Rule Library view — opens when the user clicks
  // 「查看完整規則庫 →」 in the 規則庫 tab's dashboard. Replaces the entire
  // Options tab layout (per the 2026-05-27 redesign — the rule library
  // is the most demanding workflow and warrants its own screen).
  // `ruleLibraryView` controls which sub-view (sidebar item) is active.
  const [showRuleLibrary, setShowRuleLibrary] = useState(false)
  const [ruleLibraryView, setRuleLibraryView] = useState<RuleLibrarySubView>('all')
  // Floating toast for rule-mutation feedback. Three kinds:
  //   ok   — green (e.g. "✓ 已刪除 (gmail.com)")
  //   warn — amber (e.g. "規則已不存在 — 可能被其他流程移除過")
  //   err  — red   (e.g. "刪除失敗:storage quota exceeded")
  // Auto-dismiss after 3s; manually closable via the X.
  const [ruleToast, setRuleToast] = useState<{ kind: 'ok' | 'warn' | 'err'; msg: string; key: number } | null>(null)
  useEffect(() => {
    if (!ruleToast) return
    const t = window.setTimeout(() => setRuleToast(null), 3000)
    return () => window.clearTimeout(t)
  }, [ruleToast])

  // Cross-section edit request — Rule Health rows can ask the Rules section
  // to open the editor on a specific rule. Set the id here, RulesSection's
  // useEffect picks it up, expands the editor, and scrolls the row into view.
  const [requestedEditRuleId, setRequestedEditRuleId] = useState<string | null>(null)
  const requestEditRule = useCallback((ruleId: string) => {
    setRuleSearch('') // clear search so the row is visible if filtered out
    setRequestedEditRuleId(ruleId)
    // Route into the full-screen Rule Library 「全部」 view — only that
    // view mounts RuleAllView, which consumes requestedEditRuleId.
    // Without this, clicking "編輯" from 健康度 / 自動休眠 / 衝突 sub-views
    // queued the id but the editor never opened until the user later
    // navigated back to the 全部 view, where the drawer then popped
    // open for a possibly-stale rule id.
    window.location.hash = 'rules-library/all'
  }, [])

  // 刀3 (2026-07)：popup 判斷依據面板的「編輯這條規則 →」深連結。popup
  // 以 chrome.tabs.create 開 options 並帶 ?edit=<ruleId>；這裡消費一次
  // 後從網址移除（replaceState），避免重新整理時再度彈出編輯抽屜。
  useEffect(() => {
    const params = new URLSearchParams(window.location.search)
    const editId = params.get('edit')
    if (!editId) return
    requestEditRule(editId)
    params.delete('edit')
    const qs = params.toString()
    window.history.replaceState(
      null,
      '',
      `${window.location.pathname}${qs ? `?${qs}` : ''}${window.location.hash}`,
    )
  }, [requestEditRule])

  const refresh = useCallback(async () => {
    const r = await send<StatusData>({ type: 'getStatus' })
    if (r.ok && r.data) {
      setStatus(r.data)
      // Seed the scan root input from settings ON FIRST LOAD only — the
      // user may type a one-off override after that; we don't want to
      // clobber their edits on every status refresh.
      setScanRootPath((prev) => (prev.length > 0 ? prev : r.data!.primaryRootPath))
    }
  }, [])

  const refreshScan = useCallback(async () => {
    const r = await send<{ state: ScanState | null }>({ type: 'getScanState' })
    if (r.ok && r.data) setScanState(r.data.state)
  }, [])

  const refreshRules = useCallback(async () => {
    const r = await send<{ rules: Rule[]; conflicts: Array<{ type: string; signal: string; ruleIds: string[]; targets: string[] }> }>(
      { type: 'getRules' },
    )
    if (r.ok && r.data) {
      setRules(r.data.rules)
      setConflicts(r.data.conflicts)
    }
  }, [])

  const refreshMetrics = useCallback(async () => {
    const r = await send<Metrics>({ type: 'getMetrics' })
    if (r.ok && r.data) setMetrics(r.data)
  }, [])

  const refreshSkipCount = useCallback(async () => {
    const r = await send<{ count: number }>({ type: 'getSkipHistoryCount' })
    if (r.ok && r.data) setSkipCount(r.data.count)
  }, [])

  const loadTree = useCallback(async () => {
    if (tree) return tree
    const r = await send<{ tree: MailFolderNode[] }>({ type: 'getFolderTree' })
    if (r.ok && r.data) {
      setTree(r.data.tree)
      return r.data.tree
    }
    return null
  }, [tree])

  useEffect(() => {
    void refresh()
    void refreshScan()
    void refreshRules()
    void refreshMetrics()
    void refreshSkipCount()
  }, [refresh, refreshScan, refreshRules, refreshMetrics, refreshSkipCount])

  // Scroll the page to the rules section on first mount. Run once (no
  // deps) and after a microtask so the section's DOM exists. We use the
  // browser-native anchor scroll — same path the sidebar TOC uses — so
  // the user's scroll position lands consistent with their next manual
  // navigation. Skip if URL already has a hash (user deep-linked).
  useEffect(() => {
    if (window.location.hash) return
    requestAnimationFrame(() => {
      const el = document.getElementById('rules')
      el?.scrollIntoView({ block: 'start', behavior: 'instant' as ScrollBehavior })
    })
  }, [])

  async function clearSkipHistory() {
    const r = await send<{ cleared: number }>({ type: 'clearSkipHistory' })
    if (r.ok) {
      void refreshSkipCount()
      setSkipClearedAt(Date.now())
      setTimeout(() => setSkipClearedAt(null), 2000)
    }
  }

  // toggleSkipFlagged removed 2026-05-27 — skipping flagged mail is now
  // hardcoded always-on. The storage field still exists for backward
  // compat (old machines that sync over their setting won't error) but
  // is never written to from the UI and never read by the classifier.

  async function toggleShowOwaFab(value: boolean) {
    await send({ type: 'setSettings', patch: { showOwaFab: value } })
    await refresh()
  }

  async function togglePrefetchNextBatch(value: boolean) {
    await send({ type: 'setSettings', patch: { prefetchNextBatch: value } })
    await refresh()
  }

  async function toggleAiIncludeFewShot(value: boolean) {
    await send({ type: 'setSettings', patch: { aiIncludeFewShotExamples: value } })
    await refresh()
  }

  // Poll while scan in progress
  useEffect(() => {
    if (!scanState?.inProgress) return
    const id = setInterval(() => void refreshScan(), 1000)
    return () => clearInterval(id)
  }, [scanState?.inProgress, refreshScan])

  // When scan finishes, also refresh status (rule count changed)
  useEffect(() => {
    if (scanState && !scanState.inProgress && scanState.finishedAt) {
      void refresh()
    }
  }, [scanState?.inProgress, scanState?.finishedAt, refresh])

  // Scroll-spy: highlight the TOC entry whose section is currently dominating
  // the upper-middle of the viewport. rootMargin biases toward the entry the
  // user is actively reading rather than just-scrolled-past headers.
  useEffect(() => {
    // Tab-mode now: nav is a true tab switcher (one section visible at a
    // time), not a scroll-spy. Listen for hash changes so opening the
    // page with #rules / #engine deep-links to that tab.
    //
    // Backward compatibility for the old 5-tab layout (#connection,
    // #skip) — these were merged into #engine and #data respectively
    // in the 2026-05-27 reorg. Existing bookmarks still land somewhere
    // sensible instead of falling through to the default.
    const LEGACY_HASH_REDIRECT: Record<string, string> = {
      connection: 'engine',
      skip: 'data',
    }
    const onHash = () => {
      const rawH = window.location.hash.replace('#', '')
      // Full-screen Rule Library mode — `#rules-library` or
      // `#rules-library/<sub>`. Replaces the tab layout entirely.
      if (rawH === 'rules-library' || rawH.startsWith('rules-library/')) {
        setShowRuleLibrary(true)
        const sub = rawH === 'rules-library' ? 'all' : rawH.slice('rules-library/'.length)
        if ((RULE_LIBRARY_SUBVIEWS as string[]).includes(sub)) {
          setRuleLibraryView(sub as RuleLibrarySubView)
        }
        return
      }
      setShowRuleLibrary(false)
      const h = LEGACY_HASH_REDIRECT[rawH] ?? rawH
      if (['engine', 'rules', 'data'].includes(h)) {
        setActiveSection(h)
      }
    }
    onHash() // initial
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  async function startScan() {
    setScanError(null)
    const r = await send({ type: 'startInitialScan', rootPath: scanRootPath })
    if (!r.ok) {
      setScanError(r.message)
      return
    }
    await refreshScan()
  }

  async function cancelScan() {
    await send({ type: 'cancelScan' })
  }

  async function clearScan() {
    await send({ type: 'clearScanState' })
    setScanState(null)
  }

  // Returns true on success. Audit: the boolean lets bulk actions detect a
  // backend failure (send resolves to {ok:false} without throwing), instead
  // of counting every {ok:false} as a success and silently clearing it from
  // the selection.
  async function toggleRuleEnabled(rule: Rule, enabled: boolean): Promise<boolean> {
    const r = await send({ type: 'toggleRule', ruleId: rule.id, enabled })
    void refreshRules()
    if (!r.ok) {
      setRuleToast({ kind: 'err', msg: `操作失敗:${r.message || r.code}`, key: Date.now() })
    }
    return r.ok
  }

  // Returns true when the rule is gone (deleted now, or already absent), false
  // on a backend error (so a bulk delete keeps it selected for retry).
  async function deleteRule(ruleId: string): Promise<boolean> {
    // Look up signal BEFORE the delete so the toast can name what we removed
    // — the rule is gone from `rules` after refreshRules() runs.
    const before = rules.find((r) => r.id === ruleId)
    const label = before
      ? before.type === 'compound'
        ? '複合規則'
        : before.signal
      : ''
    const r = await send<{ deleted: boolean }>({ type: 'deleteRule', ruleId })
    await refreshRules()
    if (!r.ok) {
      setRuleToast({ kind: 'err', msg: `刪除失敗:${r.message || r.code}`, key: Date.now() })
      return false
    }
    if (!r.data?.deleted) {
      setRuleToast({
        kind: 'warn',
        msg: '規則已不存在 — 可能被其他流程移除過',
        key: Date.now(),
      })
      // The rule is gone, which is the goal — treat as success for bulk.
      return true
    }
    setRuleToast({ kind: 'ok', msg: `已刪除 ${label}`, key: Date.now() })
    return true
  }

  async function upgradeDomainRuleToCompound(rule: Rule, keyword: string) {
    const kw = keyword.trim()
    if (!kw) {
      setRuleToast({ kind: 'warn', msg: '請輸入主旨關鍵字', key: Date.now() })
      return
    }
    // Build the compound signal — domain + subject_keyword AND condition.
    const compoundSignal = encodeCompound([
      { type: 'domain', value: rule.signal },
      { type: 'subject_keyword', value: kw },
    ])
    const created = await send<{ rule: Rule }>({
      type: 'createRule',
      input: {
        type: 'compound',
        signal: compoundSignal,
        targetFolderId: rule.targetFolderId,
        targetFolderPath: rule.targetFolderPath,
        confidence: 0.9,
        source: 'user_manual',
      },
    })
    if (!created.ok) {
      setRuleToast({ kind: 'err', msg: `升級失敗:${created.message}`, key: Date.now() })
      return
    }
    // Delete the original plain-domain rule. Writes a tombstone so AI won't
    // auto-resurrect the same (domain → target) mapping next batch.
    const del = await send<{ deleted: boolean }>({ type: 'deleteRule', ruleId: rule.id })
    await refreshRules()
    if (!del.ok) {
      setRuleToast({
        kind: 'warn',
        msg: '已建立複合規則,但原 domain 規則刪除失敗、請手動刪除',
        key: Date.now(),
      })
      return
    }
    setRuleToast({
      kind: 'ok',
      msg: `已升級:@${rule.signal} + 主旨「${kw}」`,
      key: Date.now(),
    })
  }

  async function clearDisabledRules() {
    const r = await send<{ deleted: number }>({ type: 'clearDisabledRules' })
    await refreshRules()
    if (!r.ok) {
      setRuleToast({ kind: 'err', msg: `清除失敗:${r.message || r.code}`, key: Date.now() })
      return r
    }
    const n = r.data?.deleted ?? 0
    setRuleToast({
      kind: n > 0 ? 'ok' : 'warn',
      msg: n > 0 ? `已清除 ${n} 條已停用規則` : '沒有已停用規則需要清除',
      key: Date.now(),
    })
    return r
  }

  async function createRule(input: {
    type: RuleType
    signal: string
    targetFolderId: string
    targetFolderPath: string
    confidence: number
  }) {
    const r = await send<{ rule: Rule }>({
      type: 'createRule',
      input: { ...input, source: 'user_manual' as RuleSource },
    })
    if (r.ok) void refreshRules()
    return r
  }

  async function upsertRule(rule: Rule) {
    const r = await send<{ rule: Rule }>({ type: 'upsertRule', rule })
    if (r.ok) void refreshRules()
    return r
  }

  async function resolveConflict(
    ruleIds: string[],
    strategy: 'keep_highest' | 'disable_all' | 'keep_one',
    keepId?: string,
  ) {
    await send({ type: 'resolveConflict', ruleIds, strategy, ...(keepId ? { keepId } : {}) })
    void refreshRules()
  }

  async function splitConflictToCompound(items: Array<{ ruleId: string; keywords: string[] }>) {
    const r = await send<{ created: number; disabled: number; keptAsFallback: number }>({
      type: 'splitConflictToCompound',
      items,
    })
    if (r.ok) void refreshRules()
    return r
  }

  async function autoUpgradeConflictRules(ruleIds: string[]) {
    const r = await send<{ created: number; disabled: number }>({
      type: 'autoUpgradeConflictRules',
      ruleIds,
    })
    if (r.ok) void refreshRules()
    return r
  }

  async function wipeAllRulesClean() {
    const r = await send<{
      cleared: number
      cloudCleared?: boolean
      cloudClearStatus?: 'skipped' | 'cleared' | 'failed'
      postWipePushed?: boolean
      syncEnabled?: boolean
    }>({ type: 'wipeAllRules' })
    if (r.ok) {
      void refreshRules()
      // If the wipeMarker push didn't go through, the user's intent
      // ("wipe everywhere") didn't propagate. Surface this so they
      // know to retry or check sync health.
      if (r.data?.syncEnabled && r.data.postWipePushed === false) {
        alert(
          '本機已清除,但雲端標記未成功寫入。其他機器不會自動同步清除,請手動再執行一次「立即上傳」,或在每台機器上各自清除。',
        )
      }
    } else {
      alert(`刪除失敗:${r.message ?? r.code}`)
    }
    return r
  }

  async function suggestSplitKeywords(folderIds: string[]) {
    return send<{ suggestions: string[][]; recentSubjects: string[][] }>({
      type: 'suggestSplitKeywords',
      folderIds,
    })
  }

  async function saveKey() {
    setKeyError(null)
    setKeySaved(false)
    const r = await send({ type: 'setApiKey', key: keyInput })
    if (!r.ok) {
      setKeyError(r.message)
      return
    }
    setKeyInput('')
    setKeySaved(true)
    // Auto-collapse the key card after a successful save — the user no
    // longer needs to see the input prompt every time they open options.
    setKeyCardExpanded(false)
    void refresh()
    setTimeout(() => setKeySaved(false), 2000)
  }

  async function clearKey() {
    const r = await send({ type: 'setApiKey', key: '' })
    if (r.ok) {
      setKeyInput('')
      // After clearing, expand so user is prompted to set a new one.
      setKeyCardExpanded(true)
      void refresh()
    }
  }

  async function changeModel(modelId: string) {
    setSavingModel(true)
    await send({ type: 'setSettings', patch: { claudeModel: modelId } })
    await refresh()
    setSavingModel(false)
  }

  async function changeBatchSize(n: number) {
    await send({ type: 'setSettings', patch: { batchSize: n } })
    await refresh()
    setBatchSavedAt(Date.now())
    setTimeout(() => setBatchSavedAt(null), 1500)
  }

  const [thresholdDraft, setThresholdDraft] = useState<number | null>(null)
  const thresholdValue = thresholdDraft ?? status?.aiConfidenceThreshold ?? 0.5
  async function commitThreshold(value: number) {
    await send({ type: 'setSettings', patch: { aiConfidenceThreshold: value } })
    await refresh()
    setThresholdDraft(null)
  }

  const [exportError, setExportError] = useState<string | null>(null)
  const [exportSavedAt, setExportSavedAt] = useState<number | null>(null)

  async function exportDiagnostic() {
    setExportError(null)
    setExportSavedAt(null)
    const r = await send<unknown>({ type: 'exportDiagnostic' })
    if (!r.ok) {
      setExportError(`匯出失敗：${r.message}`)
      return
    }
    if (!r.data) {
      setExportError('匯出失敗：背景沒回傳資料')
      return
    }
    let url: string | null = null
    try {
      const json = JSON.stringify(r.data, null, 2)
      const blob = new Blob([json], { type: 'application/json' })
      url = URL.createObjectURL(blob)
      const ts = new Date().toISOString().slice(0, 10)
      const filename = `mail-organizer-diagnostic-${ts}.json`

      // Prefer chrome.downloads — anchor.click() is blocked in some MV3
      // contexts and silently does nothing.
      if (chrome.downloads?.download) {
        await chrome.downloads.download({ url, filename, saveAs: false })
      } else {
        const a = document.createElement('a')
        a.href = url
        a.download = filename
        document.body.appendChild(a)
        a.click()
        a.remove()
      }
      setExportSavedAt(Date.now())
    } catch (e) {
      setExportError(`匯出失敗：${e instanceof Error ? e.message : String(e)}`)
    } finally {
      // Give the download a moment to start before revoking
      if (url) setTimeout(() => URL.revokeObjectURL(url!), 10_000)
      setTimeout(() => setExportSavedAt(null), 3000)
    }
  }

  return (
    <main className="min-h-screen bg-background text-foreground font-sans">
      {ruleToast && (
        <div
          key={ruleToast.key}
          className={`fixed top-4 right-4 z-50 flex items-center gap-2 rounded-md border px-3 py-2 text-xs shadow-md animate-in fade-in slide-in-from-top-2 duration-200 ${
            ruleToast.kind === 'ok'
              ? 'border-emerald-300 bg-emerald-50 text-emerald-900'
              : ruleToast.kind === 'warn'
                ? 'border-amber-300 bg-amber-50 text-amber-900'
                : 'border-red-300 bg-red-50 text-red-900'
          }`}
        >
          <span className="flex-1">{ruleToast.msg}</span>
          <button
            type="button"
            onClick={() => setRuleToast(null)}
            className="opacity-70 hover:opacity-100"
            aria-label="關閉"
          >
            <X className="size-3.5" />
          </button>
        </div>
      )}
      {showRuleLibrary ? (
        <RuleLibraryView
          view={ruleLibraryView}
          rules={rules}
          conflicts={conflicts}
          tree={tree}
          loadTree={loadTree}
          search={ruleSearch}
          setSearch={setRuleSearch}
          metrics={metrics}
          status={status}
          scanState={scanState}
          scanRootPath={scanRootPath}
          setScanRootPath={setScanRootPath}
          scanError={scanError}
          startScan={startScan}
          cancelScan={cancelScan}
          clearScan={clearScan}
          requestedEditRuleId={requestedEditRuleId}
          onEditConsumed={() => setRequestedEditRuleId(null)}
          onToggle={toggleRuleEnabled}
          onDelete={deleteRule}
          onCreate={createRule}
          onUpsert={upsertRule}
          onResolveConflict={resolveConflict}
          onSplitConflict={splitConflictToCompound}
          onAutoUpgradeConflict={autoUpgradeConflictRules}
          onSuggestKeywords={suggestSplitKeywords}
          onAfterImport={refreshRules}
          onClearDisabled={clearDisabledRules}
          onWipeAll={wipeAllRulesClean}
          onUpgradeToCompound={upgradeDomainRuleToCompound}
          onEditRule={requestEditRule}
          onBack={() => {
            window.location.hash = 'rules'
          }}
          onChangeView={(v) => {
            window.location.hash = `rules-library/${v}`
          }}
        />
      ) : (
      <div className="mx-auto max-w-5xl px-6 py-10">
        <header className="space-y-1 mb-8">
          <div className="flex items-center gap-2">
            <Mail className="size-5" />
            <h1 className="text-xl font-semibold tracking-tight">Mail Organizer — 設定</h1>
          </div>
          <p className="text-sm text-muted-foreground">
            API key、模型、批次大小、排除資料夾、初始掃描、規則管理、歷史統計。
          </p>
        </header>

        <div className="md:grid md:grid-cols-[10rem_minmax(0,1fr)] md:gap-8">
          <aside className="hidden md:block">
            <nav className="sticky top-6 flex flex-col gap-0.5 text-[13px]">
              {([
                ['engine', '分類引擎'],
                ['rules', '規則庫'],
                ['data', '資料與同步'],
              ] as const).map(([id, label]) => (
                <a
                  key={id}
                  href={`#${id}`}
                  onClick={(e) => {
                    // Tab-mode: clicking the nav switches the visible
                    // section instead of scroll-anchoring within one
                    // long page. preventDefault to keep the URL change
                    // managed by hashchange listener (single code path).
                    e.preventDefault()
                    window.location.hash = id
                    setActiveSection(id)
                    // Scroll to top so the user lands at the top of
                    // the newly-visible section.
                    window.scrollTo({ top: 0, behavior: 'instant' })
                  }}
                  className={cn(
                    'block px-2.5 py-1.5 rounded-md transition-colors',
                    activeSection === id
                      ? 'bg-accent text-foreground font-medium'
                      : 'text-muted-foreground hover:bg-accent/50 hover:text-foreground',
                  )}
                >
                  {label}
                </a>
              ))}
            </nav>
          </aside>

          <div className="min-w-0 space-y-6">
        {activeSection === 'engine' && (<>
        <div id="engine" className="scroll-mt-6">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">連線</h2>
        </div>
        {/* Claude API key — collapsed by default once configured */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <KeyRound className="size-4" />
              Claude API key
              {status?.apiKeyConfigured && <CheckCircle2 className="size-4 text-emerald-600" />}
              {status?.apiKeyConfigured && !keyCardExpanded && (
                <span className="ml-auto flex items-center gap-2 text-[10px] font-normal text-muted-foreground">
                  <code className="font-mono">{status.apiKeyPreview}</code>
                  <button
                    type="button"
                    onClick={() => setKeyCardExpanded(true)}
                    className="text-foreground/70 hover:text-foreground hover:underline"
                  >
                    管理
                  </button>
                </span>
              )}
            </CardTitle>
            {(keyCardExpanded || !status?.apiKeyConfigured) && (
              <CardDescription>
                存在 chrome.storage.local，sandbox 在 extension 內,外部網頁 / 其他 extension 不可讀。不會 echo 到 console / log / UI。
              </CardDescription>
            )}
          </CardHeader>
          {(keyCardExpanded || !status?.apiKeyConfigured) && (
            <CardContent className="space-y-3">
              {status?.apiKeyConfigured ? (
                <div className="flex items-center justify-between rounded-md border border-border bg-muted/40 p-3 text-xs">
                  <span>已設定 · <code className="font-mono">{status.apiKeyPreview}</code></span>
                  <div className="flex items-center gap-1">
                    <Button variant="ghost" size="sm" onClick={clearKey} className="text-destructive hover:text-destructive">
                      <Trash2 className="size-3.5" />
                      清除
                    </Button>
                    <Button variant="ghost" size="sm" onClick={() => setKeyCardExpanded(false)}>
                      收合
                    </Button>
                  </div>
                </div>
              ) : (
                <div className="rounded-md border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900 flex items-center gap-2">
                  <ShieldAlert className="size-4 shrink-0" />
                  尚未設定 — 設定後才能呼叫 Claude 進行 AI 分類
                </div>
              )}

              <div className="flex gap-2">
                <Input
                  type="password"
                  autoComplete="off"
                  placeholder="sk-ant-…"
                  value={keyInput}
                  onChange={(e) => setKeyInput(e.target.value)}
                  className="font-mono text-xs flex-1"
                />
                <Button onClick={saveKey} disabled={!keyInput.trim()}>
                  {status?.apiKeyConfigured ? '更換' : '儲存'}
                </Button>
              </div>
              {keyError && <p className="text-xs text-destructive">{keyError}</p>}
              {keySaved && <p className="text-xs text-emerald-700">已儲存</p>}
            </CardContent>
          )}
        </Card>

        {/* Classification preferences (G1-G5 generification, refined
            2026-05-22). Single card, two zones:
              - 主要根資料夾 (recommended, drives initial scan + prompt example)
              - 內部信件規則 (OPTIONAL, collapsed by default — only useful
                if the user has a workplace email with shared domain)
            Replaces the previous「事務所設定」title which assumed lawyer
            context. Empty everywhere = sole practitioner / personal
            account; the AI still works, just without same-domain
            grouping. */}
        {/* 歸類偏好 — moved from the old 連線 tab (2026-05-27 reorg).
            Anchors the AI's understanding of the user's folder layout;
            structurally an engine setting, not a connection setting. */}
        <div className="pt-4">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">歸類偏好</h2>
        </div>
        <ClassificationPrefsCard status={status} tree={tree} onSaved={refresh} />

        <div className="pt-4">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">引擎參數</h2>
        </div>
        {/* Model */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">分類模型</CardTitle>
            <CardDescription>
              影響分類品質與成本。Sonnet 4.6 是平衡選擇；Haiku 4.5 便宜 5×；Opus 4.7 是 ceiling。
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-2">
            {MODEL_OPTIONS.map((opt) => (
              <label
                key={opt.id}
                className="flex items-center gap-3 rounded-md border border-border bg-background hover:bg-accent px-3 py-2 cursor-pointer transition-colors"
              >
                <input
                  type="radio"
                  name="model"
                  value={opt.id}
                  checked={status?.model === opt.id}
                  onChange={() => changeModel(opt.id)}
                  disabled={savingModel}
                  className="accent-foreground"
                />
                <span className="text-sm">{opt.label}</span>
                <code className="ml-auto text-[10px] font-mono text-muted-foreground">{opt.id}</code>
              </label>
            ))}
          </CardContent>
        </Card>

        {/* Batch size */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">每次處理筆數</CardTitle>
            <CardDescription>從收件夾取最近的 N 封做歸類。可在 popup 臨時覆寫。</CardDescription>
          </CardHeader>
          <CardContent className="space-y-2">
            <div className="flex gap-2">
              {[25, 50, 100].map((n) => (
                <button
                  key={n}
                  type="button"
                  onClick={() => changeBatchSize(n)}
                  className={`flex-1 rounded-md border px-3 py-2 text-sm font-medium transition-colors ${
                    status?.batchSize === n
                      ? 'border-foreground bg-foreground text-background'
                      : 'border-border bg-background hover:bg-accent'
                  }`}
                >
                  {n}
                </button>
              ))}
            </div>
            {batchSavedAt && <p className="text-xs text-emerald-700">已儲存</p>}
          </CardContent>
        </Card>

        {/* AI confidence threshold */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">AI 信心門檻</CardTitle>
            <CardDescription>
              AI 回傳信心低於此值的建議會自動標為「保留」，不執行 move/delete。
              <span className="font-medium">較高 = 更保守</span>（多手動處理）；
              <span className="font-medium">較低 = 更自動</span>（風險高）。
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-2">
            <div className="flex items-center gap-3">
              <input
                type="range"
                min={0}
                max={1}
                step={0.05}
                value={thresholdValue}
                onChange={(e) => setThresholdDraft(Number(e.target.value))}
                onMouseUp={(e) => void commitThreshold(Number((e.target as HTMLInputElement).value))}
                onTouchEnd={(e) => void commitThreshold(Number((e.target as HTMLInputElement).value))}
                // Keyboard adjustments (Arrow/Home/End/Page keys) fire
                // change but never mouseup/touchend — without these two
                // fallbacks the display updated while the setting silently
                // never saved.
                onKeyUp={(e) => {
                  const k = e.key
                  if (
                    k === 'ArrowLeft' || k === 'ArrowRight' || k === 'ArrowUp' || k === 'ArrowDown' ||
                    k === 'Home' || k === 'End' || k === 'PageUp' || k === 'PageDown'
                  ) {
                    void commitThreshold(Number((e.target as HTMLInputElement).value))
                  }
                }}
                onBlur={(e) => {
                  if (thresholdDraft !== null) {
                    void commitThreshold(Number(e.target.value))
                  }
                }}
                className="flex-1 accent-foreground"
              />
              <span className="font-mono tabular-nums w-12 text-right text-sm">
                {thresholdValue.toFixed(2)}
              </span>
            </div>
            <div className="flex justify-between text-[10px] text-muted-foreground font-mono">
              <span>0.00 全自動</span>
              <span>0.50 平衡（建議）</span>
              <span>1.00 全手動</span>
            </div>
          </CardContent>
        </Card>

        <div className="pt-4">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">進階</h2>
        </div>
        {/* OWA floating launcher */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">在 Outlook 頁面顯示浮動圖示</CardTitle>
            <CardDescription>
              在 Outlook 網頁右下角顯示一個快捷圖示、點下去直接在 OWA 內展開歸類面板、不用切到瀏覽器工具列。瀏覽器工具列的 icon 同時保留可用。
            </CardDescription>
          </CardHeader>
          <CardContent>
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={status?.showOwaFab ?? true}
                onChange={(e) => void toggleShowOwaFab(e.target.checked)}
                className="size-4 accent-foreground"
              />
              <span className="text-sm">啟用 OWA 浮動圖示</span>
              <span className="ml-auto text-[10px] text-muted-foreground">
                {status?.showOwaFab === false ? '已停用' : '已啟用'}
              </span>
            </label>
            <p className="mt-1.5 text-[10px] text-muted-foreground">
              切換立即生效、已開啟的 Outlook 分頁自動更新、不需要重新整理。
            </p>
          </CardContent>
        </Card>

        {/* Pipeline mode — pre-classify the next batch in the background */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">Pipeline 模式(背景預跑下一批)</CardTitle>
            <CardDescription>
              執行完一批後、popup 停在「已完成」畫面時,SW 會偷偷把下一批的 classify 跑完。
              下次按「繼續歸檔下一批」幾乎瞬間進入計畫畫面、省 30-60 秒等待。
              代價:**如果你關掉 popup 不繼續、預跑的 Claude token 就浪費了**。
              建議只在你通常會連續做 2 批以上時開啟。
            </CardDescription>
          </CardHeader>
          <CardContent>
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={status?.prefetchNextBatch ?? false}
                onChange={(e) => void togglePrefetchNextBatch(e.target.checked)}
                className="size-4 accent-foreground"
              />
              <span className="text-sm">啟用 Pipeline 模式</span>
              <span className="ml-auto text-[10px] text-muted-foreground">
                {status?.prefetchNextBatch ? '已啟用' : '已停用'}
              </span>
            </label>
            <p className="mt-1.5 text-[10px] text-muted-foreground">
              預設關閉。打開後在「已完成」畫面停留期間,下一批會在背景跑(可在 token 用量看到)。
            </p>
          </CardContent>
        </Card>

        {/* AI few-shot examples (privacy toggle) */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">AI 是否參考既有規則範例</CardTitle>
            <CardDescription>
              啟用時,你已有的規則(包含 target 資料夾路徑)會被當成 few-shot 範例送到 Claude。
              關閉後 AI 只看資料夾結構與郵件本身、不會看到你的規則,但分類準確度可能略降。
              若你的資料夾名稱含敏感資訊(客戶名、案號等),建議關閉。
            </CardDescription>
          </CardHeader>
          <CardContent>
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={status?.aiIncludeFewShotExamples ?? true}
                onChange={(e) => void toggleAiIncludeFewShot(e.target.checked)}
                className="size-4 accent-foreground"
              />
              <span className="text-sm">啟用 few-shot 範例(預設開)</span>
              <span className="ml-auto text-[10px] text-muted-foreground">
                {status?.aiIncludeFewShotExamples === false ? '已停用' : '已啟用'}
              </span>
            </label>
            <p className="mt-1.5 text-[10px] text-muted-foreground">
              切換立即生效、下次分類就會反映。規則「直接命中」邏輯不受此設定影響、只影響送 Claude 的 prompt。
            </p>
          </CardContent>
        </Card>

        <div className="pt-4">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">篩選</h2>
        </div>
        {/* Recent Activity filter — configurable include rules */}
        <RecentActivityFilterCard status={status} tree={tree} onSaved={refresh} />
        </>)}

        {activeSection === 'rules' && (<>
        <div id="rules" className="scroll-mt-6 pt-2">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">規則庫</h2>
        </div>
        {/* Dashboard summary — the rules tab is now a "health overview"
            with KPIs and warnings; full management opens in its own
            full-screen view via the button at the bottom. (2026-05-27
            redesign — see RuleLibraryView for the workspace.) */}
        <RuleLibrarySummaryCard
          rules={rules}
          conflicts={conflicts}
          onOpenLibrary={(sub) => {
            const target = sub ? `rules-library/${sub}` : 'rules-library'
            window.location.hash = target
          }}
        />
        </>)}


        {activeSection === 'data' && (<>
        <div id="data" className="scroll-mt-6 pt-2">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">同步</h2>
        </div>
        {/* Cross-machine sync (Edge / Chrome account-level sync) — moved
            from the engine tab (2026-05-27 reorg). It's a data-replication
            concern, not a classification setting. */}
        {/* Sync ops (manual pull, first-enable union merge, rollback,
            applied remote wipe) rewrite the LOCAL rule library — refresh
            rules + metrics too, not just status, or the rule-library UI
            keeps showing pre-pull data and a drawer edit saved afterwards
            can overwrite freshly-pulled changes with stale copies. */}
        <CrossMachineSyncCard
          status={status}
          onChanged={async () => {
            await refresh()
            await refreshRules()
            await refreshMetrics()
          }}
        />

        <div className="pt-4">
          <h2 className="text-xs font-semibold tracking-wide text-muted-foreground uppercase border-b border-border pb-1.5">統計與紀錄</h2>
        </div>
        {/* Stats — read-only data view; collapsed by default */}
        {metrics && (
          <CollapsibleSection title="歷史統計" icon={<BarChart3 className="size-4" />}>
            <StatsSection metrics={metrics} />
          </CollapsibleSection>
        )}

        {/* Rule edit history — audit log; collapsed by default */}
        <CollapsibleSection title="規則編輯紀錄" icon={<ListChecks className="size-4" />}>
          <RuleHistorySection />
        </CollapsibleSection>

        {/* Centralised error log — surfaces silent SW failures. Moved
            from the engine tab (2026-05-27 reorg) — it's a diagnostic
            artefact, not a classification setting. */}
        <ErrorLogCard />

        {/* Skip history — moved from the removed 略過機制 tab. The
            flagged-mail toggle is now hardcoded always-on (skipping
            flagged mail is the default behavior). Only the kept-in-
            inbox history remains here as a clean-up utility. */}
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">保留歷史</CardTitle>
            <CardDescription>
              歷次執行時、你選「保留」的郵件 ID 會記錄、下次掃描自動排除。需要重新評估時點清空。
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-2 text-xs">
            <div className="flex items-baseline justify-between">
              <span className="text-muted-foreground">已記錄</span>
              <span className="tabular-nums font-mono">{skipCount} 件</span>
            </div>
            {skipCount > 0 ? (
              <div className="flex items-center gap-2">
                <Button variant="outline" size="sm" onClick={clearSkipHistory}>
                  <Trash2 className="size-3.5" />
                  清空保留歷史
                </Button>
                <span className="text-[10px] text-muted-foreground">
                  清空後、下次掃描會把這些信再列入評估
                </span>
              </div>
            ) : (
              <p className="text-[11px] text-muted-foreground italic">
                尚未累積、第一次執行後開始記錄
              </p>
            )}
            {skipClearedAt && <p className="text-[10px] text-emerald-700">已清空</p>}
          </CardContent>
        </Card>

        {/* Other settings + diagnostic export — collapsed by default */}
        <CollapsibleSection title="其他與診斷" icon={<SettingsIcon className="size-4" />}>
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <SettingsIcon className="size-4" />
              其他
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-3 text-xs">
            <div className="flex items-baseline justify-between">
              <span className="text-muted-foreground">排除資料夾前綴</span>
              <div className="flex gap-1 flex-wrap justify-end">
                {(status?.excludePrefixes ?? []).map((p) => (
                  <Badge key={p} variant="outline" className="font-mono">{p}</Badge>
                ))}
              </div>
            </div>
            <div className="flex items-baseline justify-between">
              <span className="text-muted-foreground flex items-center gap-1">
                <ListChecks className="size-3.5" />
                目前規則數
              </span>
              <span className="tabular-nums">{status?.rulesCount ?? 0}</span>
            </div>
            <div className="pt-1 border-t border-border space-y-1.5">
              <Button variant="outline" size="sm" onClick={exportDiagnostic}>
                <Download className="size-3.5" /> 匯出診斷檔（JSON）
              </Button>
              <p className="text-[10px] text-muted-foreground">
                包含設定 / 規則 / 統計 / 最近一次執行摘要。API key 不會匯出。
              </p>
              {exportSavedAt && <p className="text-[10px] text-emerald-700">已下載</p>}
              {exportError && <p className="text-[10px] text-red-700">{exportError}</p>}
            </div>
            <SettingsExportImportRow status={status} onImported={refresh} />
          </CardContent>
        </Card>
        </CollapsibleSection>
        </>)}

            <div className="flex items-center justify-center gap-3 pt-4 text-[10px] text-muted-foreground">
              <span>v{chrome.runtime.getManifest().version}</span>
              <span aria-hidden>·</span>
              {/*
                Dev-quality-of-life: reload the extension without
                having to go to chrome://extensions and click the
                reload icon. chrome.runtime.reload() restarts the
                whole extension, which re-reads source from disk —
                useful after `npm run build` to pick up new code.
                The Options page itself reloads as part of this.
              */}
              <button
                type="button"
                onClick={() => {
                  if (
                    window.confirm(
                      '重新載入擴充？\n\n會中斷目前的擴充作業（執行中的歸類批次會失效）。',
                    )
                  ) {
                    chrome.runtime.reload()
                  }
                }}
                className="inline-flex items-center gap-1 hover:text-foreground hover:underline underline-offset-2 transition-colors"
                title="重新讀取 dist/ 並重啟擴充（chrome.runtime.reload）— 開發時用、不必去 chrome://extensions"
              >
                <RotateCcw className="size-3" />
                重新載入擴充
              </button>
            </div>
          </div>
        </div>
      </div>
      )}
    </main>
  )
}

// ---- Rules section -------------------------------------------------------

const TYPE_LABEL: Record<RuleType, string> = {
  case_code: '案件代號',
  domain: '網域',
  compound: '複合',
  subject_keyword: '主旨關鍵字',
  sender: '寄件人',
}

/**
 * Refined type label for a rule. For subject_keyword rules the signal
 * shape determines what it actually represents:
 *   - Taiwan court case number (112訴204) → 「案號」
 *   - Latin case code (25A0067A) → 「案件代號」
 *   - Otherwise → 「主旨」(full-subject signal, post-2026-05-27 design)
 *
 * Helps the user distinguish at a glance:
 *   "112訴204"   [案號]   → 03/民事/甲公司案件
 *   "25A0067A"   [案件代號]→ 03/民事/甲公司案件
 *   "請款通知"   [主旨]   → 05/發票
 *
 * Defaults to TYPE_LABEL for non-subject_keyword types.
 */
function refinedTypeLabel(type: RuleType, signal: string): string {
  if (type !== 'subject_keyword') return TYPE_LABEL[type]
  if (extractCourtCaseNumbers(signal).length > 0) return '案號'
  if (extractCaseCodes(signal).length > 0) return '案件代號'
  return '主旨'
}

const SOURCE_LABEL: Record<RuleSource, string> = {
  auto_scan: '掃描',
  ai_confirmed: 'AI 確認',
  ai_overridden: 'AI 覆蓋學習',
  user_manual: '手動',
}

/**
 * Compact breakdown strip shown in the Rules library card header.
 * Quick scan of:
 *   - rule type counts (case_code / compound / domain / subject_keyword / sender)
 *   - source counts (user_manual / ai_confirmed / ai_overridden / auto_scan)
 *   - auto-disabled count (rules sleep-swept by the daily background sweep)
 *
 * Goal: let user spot rule-library bloat (e.g. "我有 800 條 subject_keyword
 * 規則?") at a glance, before it requires opening the 規則健康度 card.
 * Auto-disabled count is highlighted because they're typically subject-
 * feature rules that didn't pan out.
 */
// ============================================================
// Rule Library — Dashboard summary card (in Options 規則庫 tab)
// ============================================================

/**
 * Compact health-overview card shown on the 規則庫 tab. Replaces the
 * old approach of cramming all rule UI into the narrow tab column.
 * Click 「查看完整規則庫」 to open the full-screen workspace below.
 *
 * Sub-jump support: callers can pass a sub-view to deep-link into
 * a specific view (e.g. "conflicts" link when N > 0).
 */
function RuleLibrarySummaryCard({
  rules,
  conflicts,
  onOpenLibrary,
}: {
  rules: Rule[]
  conflicts: Array<{ type: string; signal: string; ruleIds: string[]; targets: string[] }>
  onOpenLibrary: (sub?: RuleLibrarySubView) => void
}) {
  const stats = useMemo(() => {
    const enabled = rules.filter((r) => r.enabled).length
    const dormant = rules.filter((r) => !r.enabled && r.autoDisabledAt).length
    const totalHits = rules.reduce((sum, r) => sum + r.matchCount, 0)
    const totalOverrides = rules.reduce((sum, r) => sum + (r.overrideCount ?? 0), 0)
    const accuracy = totalHits > 0 ? (totalHits - totalOverrides) / totalHits : 1
    return { total: rules.length, enabled, dormant, totalHits, accuracy }
  }, [rules])

  // Per-type breakdown for the chart-like row, sorted by count desc.
  const typeBreakdown = useMemo(() => {
    const counts: Record<string, number> = {}
    for (const r of rules) {
      if (!r.enabled) continue
      counts[r.type] = (counts[r.type] ?? 0) + 1
    }
    return Object.entries(counts).sort((a, b) => b[1] - a[1])
  }, [rules])

  // Top hits in the last "all time" (we don't have time-windowed hits
  // here — surface top by total matchCount as a reasonable proxy).
  const topHits = useMemo(
    () =>
      [...rules]
        .filter((r) => r.matchCount > 0)
        .sort((a, b) => b.matchCount - a.matchCount)
        .slice(0, 5),
    [rules],
  )

  // Warning: rules with low empirical accuracy (effectiveConfidence
  // would demote them, but they're still firing).
  const lowAccuracyCount = useMemo(
    () =>
      rules.filter((r) => {
        if (!r.enabled) return false
        if (r.matchCount < 10) return false
        const overrides = r.overrideCount ?? 0
        return overrides / r.matchCount > 0.3
      }).length,
    [rules],
  )

  // Warning: rules that have been silent for 70+ days but won't be
  // swept until 100 days. Users may want to clean them up early.
  const nearStaleCount = useMemo(() => {
    const now = Date.now()
    const SEVENTY_D_MS = 70 * 86_400_000
    return rules.filter((r) => {
      if (!r.enabled) return false
      if (r.source === 'user_manual') return false
      const lastUsed = r.lastUsedAt ? new Date(r.lastUsedAt).getTime() : new Date(r.createdAt).getTime()
      return now - lastUsed >= SEVENTY_D_MS
    }).length
  }, [rules])

  const accuracyPct = (stats.accuracy * 100).toFixed(1)
  const accuracyLabel = stats.totalHits === 0 ? '—' : `${accuracyPct}%`

  // Color-grade the accuracy label so problem libraries flag themselves
  const accuracyClass =
    stats.totalHits === 0
      ? 'text-muted-foreground'
      : stats.accuracy >= 0.9
        ? 'text-emerald-700'
        : stats.accuracy >= 0.75
          ? 'text-amber-700'
          : 'text-red-700'

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">規則庫概況</CardTitle>
        <CardDescription>
          看一眼整體健康度。詳細管理、編輯、衝突解決請點下方按鈕進入完整規則庫。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-5">
        {/* ---- KPI row ----------------------------------------- */}
        <div className="grid grid-cols-3 gap-4">
          <div className="space-y-0.5">
            <div className="text-2xl font-semibold tabular-nums">{stats.total}</div>
            <div className="text-[10px] text-muted-foreground uppercase tracking-wide">
              總規則數
            </div>
            <div className="text-[10px] text-muted-foreground tabular-nums">
              {stats.enabled} 啟用 · {stats.dormant} 休眠
            </div>
          </div>
          <div className="space-y-0.5">
            <div className="text-2xl font-semibold tabular-nums">{stats.totalHits.toLocaleString()}</div>
            <div className="text-[10px] text-muted-foreground uppercase tracking-wide">
              累積命中
            </div>
            <div className="text-[10px] text-muted-foreground tabular-nums">
              覆蓋率自動跑
            </div>
          </div>
          <div className="space-y-0.5">
            <div className={cn('text-2xl font-semibold tabular-nums', accuracyClass)}>
              {accuracyLabel}
            </div>
            <div className="text-[10px] text-muted-foreground uppercase tracking-wide">
              準確率
            </div>
            <div className="text-[10px] text-muted-foreground tabular-nums">
              命中扣除 override
            </div>
          </div>
        </div>

        {/* ---- Type breakdown (mini bar chart) ----------------- */}
        {typeBreakdown.length > 0 && (
          <div className="space-y-2 rounded-md border border-border bg-muted/30 p-3">
            <div className="text-[10px] text-muted-foreground uppercase tracking-wide font-medium">
              啟用規則類型分佈
            </div>
            <div className="space-y-1.5">
              {typeBreakdown.map(([type, count]) => {
                const pct = stats.enabled > 0 ? (count / stats.enabled) * 100 : 0
                return (
                  <div key={type} className="flex items-center gap-3 text-[11px]">
                    <Badge variant="outline" className="font-mono text-[9px] shrink-0 w-16 justify-center">
                      {TYPE_LABEL[type as RuleType] ?? type}
                    </Badge>
                    <div className="flex-1 h-1.5 rounded-full bg-muted overflow-hidden">
                      <div
                        className="h-full bg-foreground/70 transition-all"
                        style={{ width: `${pct}%` }}
                      />
                    </div>
                    <span className="tabular-nums font-mono text-muted-foreground w-8 text-right">
                      {count}
                    </span>
                  </div>
                )
              })}
            </div>
          </div>
        )}

        {/* ---- Top 5 hits ------------------------------------- */}
        {topHits.length > 0 && (
          <div className="space-y-2">
            <div className="text-[10px] text-muted-foreground uppercase tracking-wide font-medium">
              命中 TOP 5
            </div>
            <ul className="space-y-1">
              {topHits.map((r) => (
                <li key={r.id} className="flex items-center gap-2 text-xs">
                  <Badge variant="success" className="tabular-nums w-10 justify-center text-[10px]">
                    {r.matchCount}
                  </Badge>
                  <Badge variant="outline" className="font-mono text-[9px] shrink-0">
                    {refinedTypeLabel(r.type, r.signal)}
                  </Badge>
                  <span className="font-mono truncate min-w-0 flex-1">
                    {r.type === 'compound' ? formatCompoundSignal(r.signal) : r.signal}
                  </span>
                  <span className="text-muted-foreground truncate hidden md:inline">
                    → {r.targetFolderPath}
                  </span>
                </li>
              ))}
            </ul>
          </div>
        )}

        {/* ---- Warnings (only render the ones that fire) ------ */}
        {(conflicts.length > 0 || lowAccuracyCount > 0 || nearStaleCount > 0) && (
          <div className="space-y-1.5 rounded-md border border-amber-200 bg-amber-50 p-3 text-xs">
            <div className="text-[10px] text-amber-900 uppercase tracking-wide font-medium">
              注意事項
            </div>
            <ul className="space-y-1">
              {conflicts.length > 0 && (
                <li className="flex items-center justify-between gap-2 text-amber-900">
                  <span className="flex items-center gap-1.5">
                    <AlertTriangle className="size-3" />
                    {conflicts.length} 個規則衝突
                  </span>
                  <button
                    type="button"
                    onClick={() => onOpenLibrary('conflicts')}
                    className="text-amber-900 underline hover:no-underline text-[11px]"
                  >
                    查看 →
                  </button>
                </li>
              )}
              {lowAccuracyCount > 0 && (
                <li className="flex items-center justify-between gap-2 text-amber-900">
                  <span>
                    {lowAccuracyCount} 條規則錯誤率 &gt; 30%
                  </span>
                  <button
                    type="button"
                    onClick={() => onOpenLibrary('health')}
                    className="text-amber-900 underline hover:no-underline text-[11px]"
                  >
                    查看 →
                  </button>
                </li>
              )}
              {nearStaleCount > 0 && (
                <li className="flex items-center justify-between gap-2 text-amber-900">
                  <span>
                    {nearStaleCount} 條規則 70+ 天未命中(100 天會自動清除)
                  </span>
                  <button
                    type="button"
                    onClick={() => onOpenLibrary('health')}
                    className="text-amber-900 underline hover:no-underline text-[11px]"
                  >
                    查看 →
                  </button>
                </li>
              )}
            </ul>
          </div>
        )}

        {/* ---- Primary CTA ------------------------------------- */}
        <div className="flex items-center justify-between pt-2 border-t border-border">
          <span className="text-[11px] text-muted-foreground">
            完整管理、搜尋、批次操作、衝突處理等都在完整規則庫
          </span>
          <Button onClick={() => onOpenLibrary()}>
            查看完整規則庫 →
          </Button>
        </div>
      </CardContent>
    </Card>
  )
}

// ============================================================
// Rule Library — Full-screen workspace
// ============================================================

type RuleLibraryViewProps = {
  view: RuleLibrarySubView
  rules: Rule[]
  conflicts: Array<{ type: string; signal: string; ruleIds: string[]; targets: string[] }>
  tree: MailFolderNode[] | null
  loadTree: () => Promise<MailFolderNode[] | null>
  search: string
  setSearch: (s: string) => void
  metrics: Metrics | null
  status: StatusData | null
  scanState: ScanState | null
  scanRootPath: string
  setScanRootPath: (s: string) => void
  scanError: string | null
  startScan: () => Promise<void>
  cancelScan: () => Promise<void>
  clearScan: () => Promise<void>
  requestedEditRuleId: string | null
  onEditConsumed: () => void
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
  onDelete: (ruleId: string) => Promise<boolean>
  onCreate: (input: { type: RuleType; signal: string; targetFolderId: string; targetFolderPath: string; confidence: number }) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onUpsert: (rule: Rule) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onResolveConflict: (
    ruleIds: string[],
    strategy: 'keep_highest' | 'disable_all' | 'keep_one',
    keepId?: string,
  ) => Promise<void>
  onSplitConflict: (items: Array<{ ruleId: string; keywords: string[] }>) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onAutoUpgradeConflict: (ruleIds: string[]) => Promise<{ ok: true; data?: { created: number; disabled: number } } | { ok: false; code: string; message: string }>
  onSuggestKeywords: (folderIds: string[]) => Promise<{ ok: true; data?: { suggestions: string[][]; recentSubjects: string[][] } } | { ok: false; code: string; message: string }>
  onAfterImport: () => void | Promise<void>
  onClearDisabled: () => Promise<{ ok: true; data?: { deleted: number } } | { ok: false; code: string; message: string }>
  onWipeAll: () => Promise<{ ok: true; data?: { cleared: number; cloudCleared?: boolean; syncEnabled?: boolean } } | { ok: false; code: string; message: string }>
  onUpgradeToCompound: (rule: Rule, keyword: string) => Promise<void>
  onEditRule: (ruleId: string) => void
  onBack: () => void
  onChangeView: (v: RuleLibrarySubView) => void
}

function RuleLibraryView(props: RuleLibraryViewProps) {
  const { view, rules, conflicts, onBack, onChangeView } = props

  // Sidebar item counts — surface conflicts / dormant as live badges.
  const dormantCount = useMemo(
    () => rules.filter((r) => !r.enabled && r.autoDisabledAt).length,
    [rules],
  )

  type NavItem = {
    id: RuleLibrarySubView
    label: string
    count?: number
    badge?: 'danger' | 'muted'
    section?: 'main' | 'admin'
  }
  const navItems: NavItem[] = [
    { id: 'all', label: '全部規則', count: rules.length, section: 'main' },
    { id: 'conflicts', label: '衝突', count: conflicts.length, badge: conflicts.length > 0 ? 'danger' : 'muted', section: 'main' },
    { id: 'dormant', label: '自動休眠', count: dormantCount, section: 'main' },
    { id: 'health', label: '規則健康度', section: 'main' },
    { id: 'scan', label: '初次掃描', section: 'admin' },
    { id: 'history', label: '編輯紀錄', section: 'admin' },
  ]

  return (
    <div className="min-h-screen flex">
      {/* Sidebar */}
      <aside className="w-56 shrink-0 border-r border-border bg-muted/30 flex flex-col">
        <div className="p-4 border-b border-border">
          <button
            type="button"
            onClick={onBack}
            className="text-xs text-muted-foreground hover:text-foreground transition-colors flex items-center gap-1"
          >
            ← 返回設定
          </button>
          <h1 className="text-base font-semibold mt-2">規則庫</h1>
          <div className="text-[11px] text-muted-foreground tabular-nums mt-0.5">
            {rules.length} 條 · {conflicts.length} 衝突
          </div>
        </div>
        <nav className="p-2 flex-1 overflow-y-auto">
          {navItems.map((item, idx) => {
            const isAdminStart =
              item.section === 'admin' &&
              navItems[idx - 1]?.section === 'main'
            return (
              <Fragment key={item.id}>
                {isAdminStart && (
                  <div className="my-2 border-t border-border" />
                )}
                <button
                  type="button"
                  onClick={() => onChangeView(item.id)}
                  className={cn(
                    'w-full px-2.5 py-1.5 rounded-md text-sm flex items-center justify-between transition-colors',
                    view === item.id
                      ? 'bg-accent text-foreground font-medium'
                      : 'text-muted-foreground hover:bg-accent/50 hover:text-foreground',
                  )}
                >
                  <span className="flex items-center gap-2">
                    {item.label}
                    {item.id === 'conflicts' && conflicts.length > 0 && (
                      <span className="size-1.5 rounded-full bg-red-500" aria-hidden="true" />
                    )}
                  </span>
                  {typeof item.count === 'number' && item.count > 0 && (
                    <span
                      className={cn(
                        'text-[10px] tabular-nums font-mono',
                        item.badge === 'danger' ? 'text-red-700' : 'text-muted-foreground',
                      )}
                    >
                      {item.count}
                    </span>
                  )}
                </button>
              </Fragment>
            )
          })}
        </nav>
      </aside>

      {/* Main content */}
      <div className="flex-1 min-w-0 overflow-y-auto">
        <div className="max-w-6xl mx-auto p-6 space-y-4">
          <RuleLibraryViewBody {...props} />
        </div>
      </div>
    </div>
  )
}

/**
 * Renders the right-hand panel of the rule library based on `view`.
 * For Phase 1 (2026-05-27 redesign) we embed existing components for
 * each view — the long-term plan rebuilds the "all" view as a wider
 * table with side detail drawer. Other views are mostly relocated.
 */
function RuleLibraryViewBody(props: RuleLibraryViewProps) {
  const {
    view,
    rules,
    conflicts,
    tree,
    loadTree,
    status,
    scanState,
    scanRootPath,
    setScanRootPath,
    scanError,
    startScan,
    cancelScan,
    clearScan,
    requestedEditRuleId,
    onEditConsumed,
    onToggle,
    onDelete,
    onCreate,
    onUpsert,
    onResolveConflict,
    onSplitConflict,
    onAutoUpgradeConflict,
    onSuggestKeywords,
    onAfterImport,
    onClearDisabled,
    onWipeAll,
    onEditRule,
  } = props

  if (view === 'all') {
    return (
      <RuleAllView
        rules={rules}
        tree={tree}
        loadTree={loadTree}
        onToggle={onToggle}
        onDelete={onDelete}
        onCreate={onCreate}
        onUpsert={onUpsert}
        onAfterImport={onAfterImport}
        onClearDisabled={onClearDisabled}
        onWipeAll={onWipeAll}
        requestedEditId={requestedEditRuleId}
        onEditConsumed={onEditConsumed}
      />
    )
  }

  if (view === 'conflicts') {
    return (
      <ConflictsView
        rules={rules}
        conflicts={conflicts}
        tree={tree}
        onResolveConflict={onResolveConflict}
        onSplitConflict={onSplitConflict}
        onAutoUpgradeConflict={onAutoUpgradeConflict}
        onSuggestKeywords={onSuggestKeywords}
        onEditRule={onEditRule}
      />
    )
  }

  if (view === 'dormant') {
    return <DormantRulesView rules={rules} onToggle={onToggle} onDelete={onDelete} />
  }

  if (view === 'health') {
    return (
      <RuleHealthSection
        rules={rules}
        onToggle={onToggle}
        onDelete={onDelete}
        onEdit={onEditRule}
      />
    )
  }

  if (view === 'scan') {
    return (
      <Card id="initial-scan">
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-sm">
            <Radar className="size-4" />
            初次掃描
          </CardTitle>
          <CardDescription>
            掃指定根資料夾下的每個子資料夾、各取最近 20 封信，外部網域出現 ≥ 2 次就自動生 <code className="font-mono">domain</code> 規則。一次性 setup 動作、跑完歸類就有大量規則命中。
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="space-y-1">
            <label className="text-xs text-muted-foreground">根資料夾路徑</label>
            <Input
              value={scanRootPath}
              onChange={(e) => setScanRootPath(e.target.value)}
              placeholder={status?.primaryRootPath || '尚未設定主要根資料夾、請輸入路徑'}
              className="font-mono text-xs"
              disabled={scanState?.inProgress}
            />
            {status && !status.primaryRootPath && (
              <div className="text-[10px] text-amber-700">
                ⚠ 設定 → 分類引擎 → 歸類偏好 尚未設定主要根資料夾、本次掃描需手動輸入。
              </div>
            )}
          </div>

          {scanError && (
            <div className="rounded-md border border-red-200 bg-red-50 p-2 text-xs text-red-900">
              {scanError}
            </div>
          )}

          {(!scanState || (!scanState.inProgress && !scanState.finishedAt)) && (
            <Button onClick={startScan} disabled={!status?.tokenValid}>
              <Play /> 開始掃描
            </Button>
          )}

          {scanState && scanState.inProgress && (
            <div className="space-y-2">
              <div className="flex items-baseline justify-between text-xs">
                <span className="font-medium flex items-center gap-1.5">
                  <Loader2 className="size-3.5 animate-spin" />
                  {scanState.cancelRequested ? '取消中…' : `掃描中(${scanState.current}/${scanState.total})`}
                </span>
                <span className="font-mono text-muted-foreground tabular-nums">
                  +{scanState.summary.rulesAdded} 規則 · {scanState.summary.errors} 錯
                </span>
              </div>
              <div className="h-1.5 w-full rounded-full bg-muted overflow-hidden">
                <div
                  className="h-full bg-foreground transition-all duration-300"
                  style={{ width: `${(scanState.current / scanState.total) * 100}%` }}
                />
              </div>
              <div className="text-xs text-muted-foreground font-mono truncate">
                {scanState.results[scanState.current]?.folderPath ?? ''}
              </div>
              <Button variant="outline" size="sm" onClick={cancelScan} disabled={scanState.cancelRequested}>
                <X /> 取消
              </Button>
            </div>
          )}

          {scanState && !scanState.inProgress && scanState.finishedAt && (
            <ScanResults state={scanState} onReset={clearScan} />
          )}
        </CardContent>
      </Card>
    )
  }

  if (view === 'history') {
    return <RuleHistorySection />
  }

  return null
}

// ============================================================
// Rule Library — Dormant rules view (new)
// ============================================================

/**
 * Lists rules that were auto-disabled (legacy_token / high-error-rate).
 * Stale-100d rules are hard-deleted by design, so they never show up
 * here. Each row exposes [復活] + [永久刪除] actions.
 */
function DormantRulesView({
  rules,
  onToggle,
  onDelete,
}: {
  rules: Rule[]
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
  onDelete: (ruleId: string) => Promise<boolean>
}) {
  const dormant = useMemo(
    () =>
      rules
        .filter((r) => !r.enabled && r.autoDisabledAt)
        .sort((a, b) => {
          const aT = a.autoDisabledAt ?? ''
          const bT = b.autoDisabledAt ?? ''
          return bT.localeCompare(aT)
        }),
    [rules],
  )

  if (dormant.length === 0) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="text-sm">自動休眠</CardTitle>
          <CardDescription>
            被系統自動停用的規則(舊版拆詞型或長期高錯誤率)。可一鍵復活或永久刪除。
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="text-center py-8 text-sm text-muted-foreground">
            <CheckCircle2 className="size-8 mx-auto mb-2 text-emerald-600" />
            <p>沒有休眠的規則</p>
          </div>
        </CardContent>
      </Card>
    )
  }

  const reasonLabel = (reason?: string): string => {
    if (reason === 'legacy_token') return '舊版拆詞型(可能誤命中)'
    if (reason === 'high-error-rate') return '高錯誤率'
    if (reason === 'stale') return '長期未命中'
    return '未知'
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">自動休眠({dormant.length})</CardTitle>
        <CardDescription>
          按休眠時間倒序。100 天未命中規則已硬刪不在此列。
        </CardDescription>
      </CardHeader>
      <CardContent>
        <ul className="divide-y divide-border">
          {dormant.map((r) => (
            <li key={r.id} className="py-2 flex items-start gap-3 text-xs">
              <div className="flex-1 min-w-0 space-y-0.5">
                <div className="flex items-center gap-2">
                  <Badge variant="outline" className="font-mono text-[9px] shrink-0">
                    {refinedTypeLabel(r.type, r.signal)}
                  </Badge>
                  <span className="font-mono truncate min-w-0">
                    {r.type === 'compound' ? formatCompoundSignal(r.signal) : r.signal}
                  </span>
                </div>
                <div className="text-muted-foreground truncate">→ {r.targetFolderPath}</div>
                <div className="text-[10px] text-muted-foreground">
                  原因:{reasonLabel(r.autoDisabledReason)}
                  {r.autoDisabledAt && ` · ${formatDate(r.autoDisabledAt)}`}
                  {r.matchCount > 0 && ` · 過去命中 ${r.matchCount} 次`}
                </div>
              </div>
              <div className="flex items-center gap-1 shrink-0">
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => void onToggle(r, true)}
                  className="text-[10px] h-7"
                >
                  復活
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  onClick={() => void onDelete(r.id)}
                  className="text-[10px] h-7 text-destructive hover:text-destructive"
                >
                  <Trash2 className="size-3" />
                </Button>
              </div>
            </li>
          ))}
        </ul>
      </CardContent>
    </Card>
  )
}

function formatDate(iso: string): string {
  try {
    const d = new Date(iso)
    return d.toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' })
  } catch {
    return iso
  }
}

// formatRelativeTime is defined earlier in the file; reuse it.

// ============================================================
// Rule Library — 「全部」view (Phase 2 redesign)
// ============================================================
// TYPE_LABEL / SOURCE_LABEL / formatCompoundSignal live further down
// in the file (originally defined for RulesSection). We reuse those.

type RuleAllViewProps = {
  rules: Rule[]
  tree: MailFolderNode[] | null
  loadTree: () => Promise<MailFolderNode[] | null>
  // Return true on success so bulk actions can detect a failed backend write
  // (audit: a resolved {ok:false} must keep the item selected, not be counted
  // as a silent success).
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
  onDelete: (ruleId: string) => Promise<boolean>
  onCreate: (input: { type: RuleType; signal: string; targetFolderId: string; targetFolderPath: string; confidence: number }) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onUpsert: (rule: Rule) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onAfterImport: () => void | Promise<void>
  onClearDisabled: () => Promise<{ ok: true; data?: { deleted: number } } | { ok: false; code: string; message: string }>
  onWipeAll: () => Promise<{ ok: true; data?: { cleared: number; cloudCleared?: boolean; syncEnabled?: boolean } } | { ok: false; code: string; message: string }>
  requestedEditId: string | null
  onEditConsumed: () => void
}

/**
 * Wide-table rule library view. Replaces the cramped `RulesSection`
 * for the 「全部」 sub-view. Features:
 *   - Toolbar: search + 3 filter dropdowns + sort + actions
 *   - Wide table with checkbox column for bulk select
 *   - Right-side detail drawer (slide-in) when clicking a row
 *   - Bulk action bar floats when selection > 0
 *   - Add-rule form opens inline above the table
 *   - Import/Export reuses existing message handlers
 */
function RuleAllView({
  rules,
  tree,
  loadTree,
  onToggle,
  onDelete,
  onCreate,
  onUpsert,
  onAfterImport,
  onClearDisabled,
  onWipeAll,
  requestedEditId,
  onEditConsumed,
}: RuleAllViewProps) {
  const [search, setSearch] = useState('')
  const [typeFilter, setTypeFilter] = useState<RuleType | 'all'>('all')
  const [sourceFilter, setSourceFilter] = useState<RuleSource | 'all'>('all')
  const [statusFilter, setStatusFilter] = useState<'all' | 'enabled' | 'disabled'>('enabled')
  const [sortKey, setSortKey] = useState<'matchCount' | 'created' | 'lastUsed' | 'confidence'>('matchCount')
  // View mode toggle — grouped by target folder (default, matches the
  // lawyer's "see all rules for this case at once" mental model) or
  // flat sorted list. sessionStorage so a quick reload preserves it,
  // not localStorage (a deliberate fresh-start reload resets it).
  const [viewMode, setViewMode] = useState<'grouped' | 'flat'>(() => {
    if (typeof sessionStorage !== 'undefined') {
      const stored = sessionStorage.getItem('mo-rule-view-mode')
      if (stored === 'flat' || stored === 'grouped') return stored
    }
    return 'grouped'
  })
  useEffect(() => {
    if (typeof sessionStorage !== 'undefined') {
      sessionStorage.setItem('mo-rule-view-mode', viewMode)
    }
  }, [viewMode])
  // Per-folder collapse state for grouped view. Same sessionStorage policy.
  const [collapsed, setCollapsed] = useState<Set<string>>(() => {
    if (typeof sessionStorage !== 'undefined') {
      const stored = sessionStorage.getItem('mo-rule-collapsed-folders')
      if (stored) {
        try {
          const arr = JSON.parse(stored) as unknown
          if (Array.isArray(arr) && arr.every((x) => typeof x === 'string')) {
            return new Set(arr)
          }
        } catch {
          /* ignore malformed */
        }
      }
    }
    return new Set()
  })
  useEffect(() => {
    if (typeof sessionStorage !== 'undefined') {
      sessionStorage.setItem('mo-rule-collapsed-folders', JSON.stringify([...collapsed]))
    }
  }, [collapsed])
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [detailId, setDetailId] = useState<string | null>(null)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [adding, setAdding] = useState(false)
  const [bulkBusy, setBulkBusy] = useState(false)

  // External edit request from dashboard / health view — auto-open
  // detail drawer in edit mode.
  useEffect(() => {
    if (!requestedEditId) return
    setDetailId(requestedEditId)
    setEditingId(requestedEditId)
    onEditConsumed()
  }, [requestedEditId, onEditConsumed])

  // Keyboard shortcuts (Phase 4):
  //   - Esc: close detail drawer (also clears any in-progress edit)
  //   - /  : focus search input (table-style shortcut familiar from
  //          GitHub, Linear, etc.) — only when no input is focused
  const searchInputRef = useRef<HTMLInputElement | null>(null)
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      // IME guard: cancelling a Chinese/Japanese composition fires a
      // keydown with key='Escape' and isComposing=true (legacy keyCode
      // 229). That Esc means "drop the composition", NOT "close the
      // drawer" — swallowing it here closed the drawer and discarded the
      // user's whole in-progress edit.
      if (e.isComposing || e.keyCode === 229) return
      if (e.key === 'Escape' && detailId) {
        // If a nested ConfirmDialog (role="alertdialog") is open on top of
        // the drawer, let *it* consume Esc (cancel the dialog) — don't tear
        // down the whole drawer underneath it. This effect's listener is
        // registered before the dialog's (the drawer mounts first), so
        // without this guard our stopImmediatePropagation would fire first
        // and close the drawer, unmounting the dialog mid-confirm.
        if (document.querySelector('[role="alertdialog"]')) return
        // Stop the event before any outer (iframe-FAB-mode) listener
        // closes the entire Options window — Esc on a drawer should
        // dismiss the drawer, not the whole UI. See the document-level
        // Esc handler around line 239 that posts close-options to the
        // FAB iframe parent.
        e.stopImmediatePropagation()
        e.preventDefault()
        setDetailId(null)
        setEditingId(null)
        return
      }
      if (e.key === '/' && !detailId) {
        const target = e.target as HTMLElement | null
        const tag = target?.tagName
        const isInput =
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          target?.isContentEditable === true
        if (!isInput) {
          e.preventDefault()
          searchInputRef.current?.focus()
          searchInputRef.current?.select()
        }
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [detailId])

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase()
    let list = rules.filter((r) => {
      if (typeFilter !== 'all' && r.type !== typeFilter) return false
      if (sourceFilter !== 'all' && r.source !== sourceFilter) return false
      if (statusFilter === 'enabled' && !r.enabled) return false
      if (statusFilter === 'disabled' && r.enabled) return false
      if (q) {
        const signalLower = (r.type === 'compound' ? formatCompoundSignal(r.signal) : r.signal).toLowerCase()
        if (!signalLower.includes(q) && !r.targetFolderPath.toLowerCase().includes(q)) {
          return false
        }
      }
      return true
    })
    // Sort
    list = [...list].sort((a, b) => {
      switch (sortKey) {
        case 'matchCount':
          return b.matchCount - a.matchCount
        case 'created':
          return b.createdAt.localeCompare(a.createdAt)
        case 'lastUsed': {
          const aUsed = a.lastUsedAt ?? a.createdAt
          const bUsed = b.lastUsedAt ?? b.createdAt
          return bUsed.localeCompare(aUsed)
        }
        case 'confidence':
          return b.confidence - a.confidence
      }
    })
    return list
  }, [rules, search, typeFilter, sourceFilter, statusFilter, sortKey])

  const allFilteredSelected =
    filtered.length > 0 && filtered.every((r) => selected.has(r.id))

  function toggleSelectAll() {
    if (allFilteredSelected) {
      setSelected(new Set())
    } else {
      setSelected(new Set(filtered.map((r) => r.id)))
    }
  }
  // Stable identities (useCallback) so the memoized RuleRow can skip
  // re-rendering unchanged rows on search keystrokes / filter changes.
  const toggleOne = useCallback((id: string) => {
    setSelected((s) => {
      const next = new Set(s)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }, [])
  const handleToggleRule = useCallback(
    (r: Rule) => {
      void onToggle(r, !r.enabled)
    },
    [onToggle],
  )

  // Bulk action helper — survives partial failures (network blip on
  // rule 3 of 10 shouldn't lose the work on rules 1–2 or stop rules
  // 4–10). Each item runs inside its own try/catch; remaining items
  // are removed from selection (the operation was attempted), failed
  // items stay selected so the user can retry.
  // The op resolves to `false` when the backend reported failure (send →
  // {ok:false} without throwing) — audit fix: that must count as a failure
  // and keep the item selected, not be silently swallowed as success. `void`
  // / `undefined` returns still count as success (no signal = no failure).
  async function runBulk(
    ids: string[],
    op: (id: string) => Promise<boolean | void>,
  ): Promise<{ ok: number; failed: string[] }> {
    let ok = 0
    const failed: string[] = []
    for (const id of ids) {
      try {
        const res = await op(id)
        if (res === false) failed.push(id)
        else ok++
      } catch (e) {
        console.warn('[mail-organizer] bulk action item failed', id, e)
        failed.push(id)
      }
    }
    return { ok, failed }
  }

  async function bulkDisable() {
    setBulkBusy(true)
    try {
      const ids = Array.from(selected)
      const { failed } = await runBulk(ids, async (id) => {
        const r = rules.find((x) => x.id === id)
        // No-op (already disabled / not found) is success; otherwise the
        // backend result decides.
        if (!r || !r.enabled) return true
        return await onToggle(r, false)
      })
      // Clear the survivors from selection; keep failed for retry.
      setSelected(new Set(failed))
    } finally {
      setBulkBusy(false)
    }
  }
  async function bulkEnable() {
    setBulkBusy(true)
    try {
      const ids = Array.from(selected)
      const { failed } = await runBulk(ids, async (id) => {
        const r = rules.find((x) => x.id === id)
        if (!r || r.enabled) return true
        return await onToggle(r, true)
      })
      setSelected(new Set(failed))
    } finally {
      setBulkBusy(false)
    }
  }
  async function bulkDelete() {
    if (!confirm(`確定刪除選取的 ${selected.size} 條規則?(寫入墓碑、AI 不會再生成)`)) return
    setBulkBusy(true)
    try {
      const ids = Array.from(selected)
      // If the currently-open detail drawer references a rule we're
      // about to delete, close it first so the user doesn't see a
      // ghost drawer pointing at a deleted id (the drawer's
      // `detailRule` lookup returns undefined → unmounts cleanly,
      // but editing state would linger).
      if (detailId && ids.includes(detailId)) {
        setDetailId(null)
        setEditingId(null)
      }
      const { failed } = await runBulk(ids, (id) => onDelete(id))
      setSelected(new Set(failed))
    } finally {
      setBulkBusy(false)
    }
  }

  const detailRule = useMemo(
    () => (detailId ? rules.find((r) => r.id === detailId) ?? null : null),
    [detailId, rules],
  )

  return (
    <div className="space-y-4">
      {/* ---- Toolbar -------------------------------------------- */}
      <div className="flex flex-wrap items-center gap-2">
        <div className="relative flex-1 min-w-[200px] max-w-md">
          <Search className="absolute left-2 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground pointer-events-none" />
          <Input
            ref={searchInputRef}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="搜尋 signal 或目標資料夾…（/ 快速聚焦)"
            className="pl-7 text-xs h-8"
          />
        </div>
        <select
          value={typeFilter}
          onChange={(e) => setTypeFilter(e.target.value as RuleType | 'all')}
          className="text-xs h-8 rounded-md border border-border bg-background px-2"
        >
          <option value="all">類型:全部</option>
          {(['case_code', 'compound', 'domain', 'subject_keyword', 'sender'] as RuleType[]).map((t) => (
            <option key={t} value={t}>類型:{TYPE_LABEL[t]}</option>
          ))}
        </select>
        <select
          value={sourceFilter}
          onChange={(e) => setSourceFilter(e.target.value as RuleSource | 'all')}
          className="text-xs h-8 rounded-md border border-border bg-background px-2"
        >
          <option value="all">來源:全部</option>
          {(['user_manual', 'ai_overridden', 'ai_confirmed', 'auto_scan'] as RuleSource[]).map((s) => (
            <option key={s} value={s}>來源:{SOURCE_LABEL[s]}</option>
          ))}
        </select>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value as 'all' | 'enabled' | 'disabled')}
          className="text-xs h-8 rounded-md border border-border bg-background px-2"
        >
          <option value="enabled">狀態:啟用</option>
          <option value="disabled">狀態:停用</option>
          <option value="all">狀態:全部</option>
        </select>
        <select
          value={sortKey}
          onChange={(e) => setSortKey(e.target.value as typeof sortKey)}
          className="text-xs h-8 rounded-md border border-border bg-background px-2"
        >
          <option value="matchCount">排序:命中數</option>
          <option value="lastUsed">排序:最近使用</option>
          <option value="created">排序:建立時間</option>
          <option value="confidence">排序:Confidence</option>
        </select>
        <select
          value={viewMode}
          onChange={(e) => setViewMode(e.target.value as 'grouped' | 'flat')}
          className="text-xs h-8 rounded-md border border-border bg-background px-2"
          aria-label="檢視模式"
        >
          <option value="grouped">檢視:按資料夾分組</option>
          <option value="flat">檢視:平面清單</option>
        </select>
        <div className="ml-auto flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={() => setAdding(true)}
            disabled={!tree || tree.length === 0}
            title={!tree ? '需先載入資料夾樹' : ''}
          >
            <Plus className="size-3.5" />
            新增規則
          </Button>
          <RuleLibraryActionsMenu
            onAfterImport={onAfterImport}
            onClearDisabled={onClearDisabled}
            onWipeAll={onWipeAll}
            rulesCount={rules.length}
            disabledCount={rules.filter((r) => !r.enabled).length}
          />
        </div>
      </div>

      {/* Result summary line */}
      <div className="text-[11px] text-muted-foreground flex items-center justify-between">
        <span>
          顯示 {filtered.length} / {rules.length} 條
          {selected.size > 0 && ` · 已選 ${selected.size}`}
        </span>
        {(search || typeFilter !== 'all' || sourceFilter !== 'all' || statusFilter !== 'enabled') && (
          <button
            type="button"
            onClick={() => {
              setSearch('')
              setTypeFilter('all')
              setSourceFilter('all')
              setStatusFilter('enabled')
            }}
            className="text-muted-foreground hover:text-foreground underline"
          >
            重設篩選
          </button>
        )}
      </div>

      {/* ---- Inline add form ------------------------------------ */}
      {adding && tree && (
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">新增規則</CardTitle>
          </CardHeader>
          <CardContent>
            <RuleForm
              tree={tree}
              rules={rules}
              onCancel={() => setAdding(false)}
              onSubmit={async (values) => {
                const r = await onCreate(values)
                if (r.ok) setAdding(false)
                return r
              }}
            />
          </CardContent>
        </Card>
      )}

      {/* ---- Table ---------------------------------------------- */}
      {filtered.length === 0 ? (
        <div className="text-center py-16 text-sm text-muted-foreground border border-dashed border-border rounded-md">
          {rules.length === 0 ? (
            <>
              <p>規則庫是空的</p>
              <p className="text-[11px] mt-1">點右上「新增規則」或執行「初次掃描」開始</p>
            </>
          ) : (
            <>
              <p>沒有符合篩選條件的規則</p>
              <p className="text-[11px] mt-1">試試重設篩選</p>
            </>
          )}
        </div>
      ) : viewMode === 'grouped' ? (
        <RuleGroupedTable
          rules={filtered}
          allRules={rules}
          selected={selected}
          detailId={detailId}
          collapsed={collapsed}
          setCollapsed={setCollapsed}
          onToggleSelectAll={toggleSelectAll}
          allFilteredSelected={allFilteredSelected}
          onSelect={toggleOne}
          onOpenDetail={setDetailId}
          onToggle={onToggle}
        />
      ) : (
        <div className="rounded-md border border-border overflow-hidden bg-background">
          <table className="w-full text-xs">
            <thead className="bg-muted/50">
              <tr className="text-left text-[10px] uppercase tracking-wide text-muted-foreground">
                <th className="px-2 py-2 w-8">
                  <input
                    type="checkbox"
                    checked={allFilteredSelected}
                    onChange={toggleSelectAll}
                    className="accent-foreground"
                    aria-label="全選"
                  />
                </th>
                <th className="px-2 py-2 w-20">類型</th>
                <th className="px-2 py-2">Signal</th>
                <th className="px-2 py-2">→ 目標資料夾</th>
                <th className="px-2 py-2 w-16 text-right">命中</th>
                <th className="px-2 py-2 w-20 text-right">Conf</th>
                <th className="px-2 py-2 w-20">來源</th>
                <th className="px-2 py-2 w-12"></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => (
                <RuleRow
                  key={r.id}
                  rule={r}
                  selected={selected.has(r.id)}
                  highlighted={detailId === r.id}
                  showTarget
                  onSelect={toggleOne}
                  onOpenDetail={setDetailId}
                  onToggle={handleToggleRule}
                />
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ---- Bulk action bar (floating) ------------------------- */}
      {selected.size > 0 && (
        <div className="sticky bottom-4 z-30 mx-auto max-w-2xl">
          <div className="rounded-md border border-foreground bg-foreground text-background px-4 py-2 shadow-lg flex items-center gap-3">
            <span className="text-xs font-medium tabular-nums">
              已選 {selected.size} 條
            </span>
            <div className="flex items-center gap-1 ml-auto">
              <Button
                size="sm"
                variant="ghost"
                className="text-background hover:bg-background/10 text-[11px] h-7"
                onClick={() => void bulkEnable()}
                disabled={bulkBusy}
              >
                <Power className="size-3" />
                啟用
              </Button>
              <Button
                size="sm"
                variant="ghost"
                className="text-background hover:bg-background/10 text-[11px] h-7"
                onClick={() => void bulkDisable()}
                disabled={bulkBusy}
              >
                <PowerOff className="size-3" />
                停用
              </Button>
              <Button
                size="sm"
                variant="ghost"
                className="text-red-300 hover:text-red-200 hover:bg-background/10 text-[11px] h-7"
                onClick={() => void bulkDelete()}
                disabled={bulkBusy}
              >
                <Trash2 className="size-3" />
                刪除
              </Button>
              <button
                type="button"
                onClick={() => setSelected(new Set())}
                className="text-background/60 hover:text-background ml-2"
                aria-label="取消選取"
              >
                <X className="size-3.5" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---- Detail drawer -------------------------------------- */}
      {detailRule && (
        <RuleDetailDrawer
          rule={detailRule}
          tree={tree}
          rules={rules}
          editing={editingId === detailRule.id}
          onStartEdit={() => setEditingId(detailRule.id)}
          onCancelEdit={() => setEditingId(null)}
          onClose={() => {
            setDetailId(null)
            setEditingId(null)
          }}
          onToggle={onToggle}
          onDelete={async (id) => {
            await onDelete(id)
            setDetailId(null)
            setEditingId(null)
          }}
          onUpsert={async (rule) => {
            const r = await onUpsert(rule)
            if (r.ok) setEditingId(null)
            return r
          }}
          loadTree={loadTree}
        />
      )}
    </div>
  )
}

// ============================================================
// Rule Library — Action menu (import / export / clear / wipe)
// ============================================================

function RuleLibraryActionsMenu({
  onAfterImport,
  onClearDisabled,
  onWipeAll,
  rulesCount,
  disabledCount,
}: {
  onAfterImport: () => void | Promise<void>
  onClearDisabled: () => Promise<{ ok: true; data?: { deleted: number } } | { ok: false; code: string; message: string }>
  onWipeAll: () => Promise<{ ok: true; data?: { cleared: number; cloudCleared?: boolean; syncEnabled?: boolean } } | { ok: false; code: string; message: string }>
  rulesCount: number
  disabledCount: number
}) {
  const [open, setOpen] = useState(false)
  const [confirmWipe, setConfirmWipe] = useState(false)
  const [confirmClear, setConfirmClear] = useState(false)
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const menuRef = useRef<HTMLDivElement | null>(null)
  const triggerRef = useRef<HTMLButtonElement | null>(null)
  // A11y: list of menuitem elements, snapshot when menu opens, so arrow
  // keys can navigate without re-querying the DOM on every key press.
  // Updated by the panel's ref callback below.
  const menuItemsRef = useRef<HTMLElement[]>([])

  // Close on outside click
  useEffect(() => {
    if (!open) return
    function onClick(e: MouseEvent) {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', onClick)
    return () => document.removeEventListener('mousedown', onClick)
  }, [open])

  // A11y: Esc closes + restores focus to the trigger. Arrow keys
  // navigate between menu items while open.
  useEffect(() => {
    if (!open) return
    function onKey(e: KeyboardEvent) {
      if (e.key === 'Escape') {
        e.stopPropagation()
        e.preventDefault()
        setOpen(false)
        // Restore focus to the trigger button so the user's Tab
        // position is preserved.
        triggerRef.current?.focus()
        return
      }
      // F17 (2026-06-03): navigate only ENABLED items. Disabled menu
      // buttons (e.g. 匯出 when 0 rules, 清除已停用 when none disabled)
      // are non-focusable; including them made arrow nav appear to
      // "stall" when it landed on one. Filter to focusable items per
      // keypress so the WAI-ARIA "skip disabled" contract actually holds.
      const focusable = () =>
        menuItemsRef.current.filter(
          (el) => el && !(el as HTMLButtonElement).disabled,
        )
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        const items = focusable()
        if (items.length === 0) return
        e.preventDefault()
        const active = document.activeElement as HTMLElement | null
        const idx = active ? items.indexOf(active) : -1
        const delta = e.key === 'ArrowDown' ? 1 : -1
        // Wrap. -1 (no current focus) + ArrowDown = first; + ArrowUp = last.
        const nextIdx = idx === -1
          ? (delta > 0 ? 0 : items.length - 1)
          : (idx + delta + items.length) % items.length
        items[nextIdx]?.focus()
        return
      }
      if (e.key === 'Home') {
        e.preventDefault()
        focusable()[0]?.focus()
        return
      }
      if (e.key === 'End') {
        e.preventDefault()
        const items = focusable()
        items[items.length - 1]?.focus()
        return
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [open])

  // When the menu opens, focus the first item so keyboard users land
  // inside. Defer one tick so the items have rendered.
  useEffect(() => {
    if (!open) {
      // Clear the items array on close so arrow keys don't navigate to
      // stale (detached) item nodes the next time the menu reopens.
      menuItemsRef.current = []
      return
    }
    const t = setTimeout(() => {
      menuItemsRef.current[0]?.focus()
    }, 0)
    return () => clearTimeout(t)
  }, [open])

  async function handleExport() {
    setOpen(false)
    const r = await send<{ json: string; filename: string }>({ type: 'exportRules' })
    if (!r.ok || !r.data) {
      // Same feedback idiom as handleImportFile below.
      alert(`匯出失敗:${(!r.ok && (r.message ?? r.code)) || '背景沒回傳資料'}`)
      return
    }
    let url: string | null = null
    try {
      const blob = new Blob([r.data.json], { type: 'application/json' })
      url = URL.createObjectURL(blob)
      // Prefer chrome.downloads — a detached anchor.click() is blocked in
      // some MV3 contexts (notably the FAB iframe) and silently does
      // nothing; and revoking the blob URL synchronously could cut off the
      // download before it started. Mirrors exportDiagnostic /
      // exportSettings, which were hardened for exactly this.
      if (chrome.downloads?.download) {
        await chrome.downloads.download({ url, filename: r.data.filename, saveAs: false })
      } else {
        const a = document.createElement('a')
        a.href = url
        a.download = r.data.filename
        document.body.appendChild(a)
        a.click()
        a.remove()
      }
    } catch (e) {
      alert(`匯出失敗:${e instanceof Error ? e.message : String(e)}`)
    } finally {
      if (url) {
        const u = url
        setTimeout(() => URL.revokeObjectURL(u), 10_000)
      }
    }
  }

  async function handleImportFile(e: ChangeEvent<HTMLInputElement>) {
    setOpen(false)
    const file = e.target.files?.[0]
    if (!file) return
    const text = await file.text()
    e.target.value = ''
    // Use replace strategy — the simplest "merge" for now. Power users
    // can use the legacy flow if needed.
    const r = await send<{
      added: number
      replaced: number
      total: number
      skipped: number
      orphanedCount: number
      importedTombstones: number
      cappedCount: number
      cappedMaxDelta: number
    }>({ type: 'importRules', json: text, strategy: 'merge' })
    if (r.ok) {
      await onAfterImport()
      // Surface confidence-cap impact so the user isn't surprised
      // their imported rules show lower trust than the backup file
      // recorded. Threshold of ≥1 cap-affected rule — we don't pop
      // up for trivial 0-delta cases.
      const d = r.data
      if (d && d.cappedCount > 0) {
        alert(
          `匯入完成,但 ${d.cappedCount} 條規則的信心度因目前系統上限被調低` +
            ` (最大降幅 ${d.cappedMaxDelta.toFixed(2)})。\n\n` +
            `這是預期行為:新版針對不同規則類型設了上限 ` +
            `(網域 0.7 / 寄件人 0.75 / 主旨 0.9 / 複合與案號 0.95),` +
            `避免廣域規則蓋過更精準的同 folder 規則。` +
            `user_manual 規則不受限。`,
        )
      }
    } else {
      alert(`匯入失敗:${r.message ?? r.code}`)
    }
  }

  // Callback ref that collects menu-item buttons in DOM order. Excludes
  // disabled items so arrow keys skip over them (matches WAI-ARIA menu
  // pattern — disabled items are skipped in navigation).
  function setItemRef(idx: number, el: HTMLButtonElement | null) {
    if (!el) return
    menuItemsRef.current[idx] = el
  }

  return (
    <div className="relative" ref={menuRef}>
      <Button
        ref={triggerRef}
        variant="outline"
        size="sm"
        onClick={() => setOpen((o) => !o)}
        aria-haspopup="menu"
        aria-expanded={open}
      >
        <SettingsIcon className="size-3.5" />
        操作
        <ChevronDown className="size-3" />
      </Button>
      {open && (
        <div
          role="menu"
          aria-label="規則庫操作"
          className="absolute right-0 top-full mt-1 z-20 min-w-[180px] rounded-md border border-border bg-background shadow-md py-1 text-xs"
        >
          <button
            type="button"
            role="menuitem"
            ref={(el) => setItemRef(0, el)}
            onClick={() => {
              setOpen(false)
              fileInputRef.current?.click()
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-accent focus:bg-accent focus:outline-none flex items-center gap-2"
          >
            <Upload className="size-3.5" />
            匯入規則
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="application/json,.json"
            hidden
            onChange={(e) => void handleImportFile(e)}
          />
          <button
            type="button"
            role="menuitem"
            ref={(el) => setItemRef(1, el)}
            onClick={handleExport}
            disabled={rulesCount === 0}
            className="w-full text-left px-3 py-1.5 hover:bg-accent focus:bg-accent focus:outline-none disabled:opacity-50 flex items-center gap-2"
          >
            <Download className="size-3.5" />
            匯出規則
          </button>
          <div role="separator" className="border-t border-border my-1" />
          <button
            type="button"
            role="menuitem"
            ref={(el) => setItemRef(2, el)}
            onClick={() => {
              setOpen(false)
              setConfirmClear(true)
            }}
            disabled={disabledCount === 0}
            className="w-full text-left px-3 py-1.5 hover:bg-accent focus:bg-accent focus:outline-none disabled:opacity-50 flex items-center gap-2"
          >
            <Trash2 className="size-3.5" />
            清除已停用 ({disabledCount})
          </button>
          <button
            type="button"
            role="menuitem"
            ref={(el) => setItemRef(3, el)}
            onClick={() => {
              setOpen(false)
              setConfirmWipe(true)
            }}
            disabled={rulesCount === 0}
            className="w-full text-left px-3 py-1.5 hover:bg-red-50 focus:bg-red-50 focus:outline-none text-red-700 disabled:opacity-50 flex items-center gap-2"
          >
            <Trash2 className="size-3.5" />
            全部刪除…
          </button>
        </div>
      )}
      {confirmClear && (
        <ConfirmDialog
          title={`刪除 ${disabledCount} 條已停用規則?`}
          description="會寫入墓碑、AI 不會再自動生成。"
          confirmLabel="刪除"
          danger
          onConfirm={async () => {
            await onClearDisabled()
            setConfirmClear(false)
          }}
          onCancel={() => setConfirmClear(false)}
        />
      )}
      {confirmWipe && (
        <ConfirmDialog
          title={`刪除全部 ${rulesCount} 條規則?`}
          description="同時清空 AI 記憶 + 墓碑庫 + 審計日誌。不可復原。已啟用同步時會自動觸發其他機器在下次同步時一併清除。"
          confirmLabel="全部刪除"
          danger
          onConfirm={async () => {
            await onWipeAll()
            setConfirmWipe(false)
          }}
          onCancel={() => setConfirmWipe(false)}
        />
      )}
    </div>
  )
}

function ConfirmDialog({
  title,
  description,
  confirmLabel,
  danger,
  onConfirm,
  onCancel,
}: {
  title: string
  description?: string
  confirmLabel: string
  danger?: boolean
  onConfirm: () => Promise<void> | void
  onCancel: () => void
}) {
  const [busy, setBusy] = useState(false)
  // Esc cancels the dialog instead of bubbling out to the iframe-FAB
  // listener that would close the entire Options window. Cancel button
  // also gets autofocus so Enter on cancel doesn't accidentally confirm
  // a destructive action.
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (e.key === 'Escape' && !busy) {
        e.stopImmediatePropagation()
        e.preventDefault()
        onCancel()
      }
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [busy, onCancel])
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="confirm-dialog-title"
    >
      {/* Backdrop click cancels — non-destructive interpretation of
          "I clicked away by mistake". */}
      <button
        type="button"
        aria-label="取消並關閉"
        onClick={() => !busy && onCancel()}
        className="absolute inset-0 cursor-default"
      />
      <div className="relative bg-background rounded-md border border-border shadow-lg max-w-md w-full p-5 space-y-3">
        <h3 id="confirm-dialog-title" className="text-sm font-medium">{title}</h3>
        {description && (
          <p className="text-xs text-muted-foreground">{description}</p>
        )}
        <div className="flex justify-end gap-2 pt-2">
          <Button
            variant="outline"
            size="sm"
            onClick={onCancel}
            disabled={busy}
            autoFocus
          >
            取消
          </Button>
          <Button
            size="sm"
            className={danger ? 'bg-red-600 hover:bg-red-700 text-white border-red-700' : ''}
            disabled={busy}
            onClick={async () => {
              setBusy(true)
              try {
                await onConfirm()
              } finally {
                setBusy(false)
              }
            }}
          >
            {confirmLabel}
          </Button>
        </div>
      </div>
    </div>
  )
}

// ============================================================
// Rule Library — Detail drawer (slide from right)
// ============================================================

function RuleDetailDrawer({
  rule,
  tree,
  rules,
  editing,
  onStartEdit,
  onCancelEdit,
  onClose,
  onToggle,
  onDelete,
  onUpsert,
  loadTree,
}: {
  rule: Rule
  tree: MailFolderNode[] | null
  rules: Rule[]
  editing: boolean
  onStartEdit: () => void
  onCancelEdit: () => void
  onClose: () => void
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
  onDelete: (ruleId: string) => Promise<void>
  onUpsert: (rule: Rule) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  loadTree: () => Promise<MailFolderNode[] | null>
}) {
  const accuracyPct =
    rule.matchCount > 0
      ? Math.round(((rule.matchCount - (rule.overrideCount ?? 0)) / rule.matchCount) * 100)
      : null
  const [confirmDelete, setConfirmDelete] = useState(false)

  // Load tree on open if not loaded (so edit form can show folder picker).
  useEffect(() => {
    if (!tree && editing) {
      void loadTree()
    }
  }, [editing, tree, loadTree])

  // ---- A11y: focus trap + focus restoration -------------------------------
  //
  // - On open: remember the previously-focused element so we can restore
  //   focus on close (typical pattern when a "row click" launches a modal —
  //   user expects focus to come back to that row).
  // - While open: pressing Tab cycles within the drawer; Shift+Tab cycles
  //   backwards. Tab from the last item wraps to the first; Shift+Tab from
  //   the first wraps to the last. Prevents focus from leaking to the
  //   underlying rules table (which is still in the DOM behind the backdrop).
  // - On open: auto-focus the close button so a keyboard user lands inside
  //   the drawer instead of having to Tab in from wherever they were.
  const drawerRef = useRef<HTMLDivElement | null>(null)
  const closeButtonRef = useRef<HTMLButtonElement | null>(null)
  useEffect(() => {
    const previouslyFocused = document.activeElement as HTMLElement | null
    // Defer the focus to next tick so the drawer's elements are in the DOM
    // and the animate-in class has applied (auto-focus during the
    // mount-render race can land focus before the close button paints).
    const t = setTimeout(() => {
      closeButtonRef.current?.focus()
    }, 0)
    return () => {
      clearTimeout(t)
      // Restore focus to whatever opened the drawer (typically a RuleRow).
      if (previouslyFocused && document.contains(previouslyFocused)) {
        try {
          previouslyFocused.focus()
        } catch {
          /* ignore — element may be detached / unfocusable */
        }
      }
    }
  }, [])
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (e.key !== 'Tab') return
      const root = drawerRef.current
      if (!root) return
      // F16 (2026-06-03): when a nested ConfirmDialog (role="alertdialog")
      // is open and focus is inside it, let IT manage focus — don't yank
      // back into the drawer. The dialog renders as a sibling of drawerRef
      // (outside it), so without this bail the drawer's "focus escaped,
      // snap it back" branch fires on every Tab inside the dialog, making
      // the destructive-confirm button unreachable by keyboard.
      const activeEl = document.activeElement as HTMLElement | null
      if (activeEl?.closest('[role="alertdialog"]')) return
      // Snapshot focusable elements inside the drawer right before each
      // Tab so dynamic content (edit form opening, delete confirm dialog,
      // etc.) is included in the cycle. Cheap — the drawer has at most
      // 15-20 focusable elements.
      const focusables = Array.from(
        root.querySelectorAll<HTMLElement>(
          'a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])',
        ),
      ).filter((el) => !el.hidden && el.offsetParent !== null)
      if (focusables.length === 0) return
      const first = focusables[0]!
      const last = focusables[focusables.length - 1]!
      const active = document.activeElement as HTMLElement | null
      if (e.shiftKey) {
        // Backward — wrap to last when at first OR focus is outside the
        // drawer (defensive: if focus somehow escaped, snap it back).
        if (active === first || !root.contains(active)) {
          e.preventDefault()
          last.focus()
        }
      } else {
        // Forward — wrap to first when at last OR outside.
        if (active === last || !root.contains(active)) {
          e.preventDefault()
          first.focus()
        }
      }
    }
    // Use capture phase so we beat any element-level Tab handlers inside
    // form widgets (e.g. a custom FolderPicker that might consume Tab).
    document.addEventListener('keydown', onKey, true)
    return () => document.removeEventListener('keydown', onKey, true)
  }, [])

  return (
    <div
      className="fixed inset-0 z-40 flex justify-end"
      role="dialog"
      aria-modal="true"
      aria-label="規則詳細"
    >
      {/* Backdrop */}
      <button
        type="button"
        aria-label="關閉"
        onClick={onClose}
        // tabIndex=-1 keeps the backdrop out of the Tab cycle — a sighted
        // user can click it, but keyboard users use the dedicated close
        // button + Esc handler instead. Without this, every Tab cycle
        // hits "Close (backdrop)" as a hidden first stop.
        tabIndex={-1}
        className="absolute inset-0 bg-black/30"
      />
      {/* Drawer */}
      <div
        ref={drawerRef}
        className="relative w-full max-w-md bg-background border-l border-border shadow-xl overflow-y-auto animate-in slide-in-from-right duration-200"
      >
        <div className="sticky top-0 bg-background border-b border-border px-4 py-3 flex items-center justify-between">
          <h2 className="text-sm font-medium">規則詳細</h2>
          <button
            type="button"
            ref={closeButtonRef}
            onClick={onClose}
            className="text-muted-foreground hover:text-foreground"
            aria-label="關閉"
          >
            <X className="size-4" />
          </button>
        </div>
        <div className="p-4 space-y-4 text-xs">
          {editing && tree ? (
            <RuleForm
              tree={tree}
              initial={rule}
              rules={rules.filter((r) => r.id !== rule.id)}
              onCancel={onCancelEdit}
              onSubmit={async (values) => {
                const updated: Rule = { ...rule, ...values }
                return onUpsert(updated)
              }}
            />
          ) : (
            <>
              {/* Identity */}
              <section className="space-y-2">
                <Field label="類型">
                  <Badge variant="outline" className="font-mono text-[10px]">
                    {refinedTypeLabel(rule.type, rule.signal)}
                  </Badge>
                </Field>
                <Field label="Signal">
                  <code className="font-mono break-all text-[11px]">
                    {rule.type === 'compound' ? formatCompoundSignal(rule.signal) : rule.signal}
                  </code>
                </Field>
                <Field label="目標資料夾">
                  <code className="font-mono break-all text-[11px]">{rule.targetFolderPath}</code>
                </Field>
              </section>

              <div className="border-t border-border" />

              {/* Stats */}
              <section className="grid grid-cols-2 gap-3">
                <Field label="Confidence" tight>
                  <span className="font-mono tabular-nums">{rule.confidence.toFixed(2)}</span>
                </Field>
                <Field label="命中次數" tight>
                  <span className="font-mono tabular-nums">{rule.matchCount}</span>
                </Field>
                <Field label="Override 次數" tight>
                  <span className="font-mono tabular-nums">{rule.overrideCount ?? 0}</span>
                </Field>
                <Field label="準確率" tight>
                  {accuracyPct !== null ? (
                    <span
                      className={cn(
                        'font-mono tabular-nums',
                        accuracyPct >= 90
                          ? 'text-emerald-700'
                          : accuracyPct >= 70
                            ? 'text-amber-700'
                            : 'text-red-700',
                      )}
                    >
                      {accuracyPct}%
                    </span>
                  ) : (
                    <span className="text-muted-foreground">—</span>
                  )}
                </Field>
              </section>

              <div className="border-t border-border" />

              {/* Provenance */}
              <section className="space-y-2">
                <Field label="來源">
                  <Badge variant="outline" className="text-[10px]">
                    {SOURCE_LABEL[rule.source]}
                  </Badge>
                </Field>
                <Field label="狀態">
                  {rule.enabled ? (
                    <Badge variant="success" className="text-[10px]">啟用</Badge>
                  ) : (
                    <Badge variant="outline" className="text-[10px] text-muted-foreground">
                      停用
                      {rule.autoDisabledReason && ` · ${rule.autoDisabledReason}`}
                    </Badge>
                  )}
                </Field>
                <Field label="建立時間">
                  <span>{formatDate(rule.createdAt)}</span>
                  <span className="text-muted-foreground ml-1">({formatRelativeTime(rule.createdAt)})</span>
                </Field>
                {rule.lastUsedAt && (
                  <Field label="最近使用">
                    <span>{formatDate(rule.lastUsedAt)}</span>
                    <span className="text-muted-foreground ml-1">({formatRelativeTime(rule.lastUsedAt)})</span>
                  </Field>
                )}
                {rule.orphaned && (
                  <div className="rounded-md border border-amber-200 bg-amber-50 p-2 text-amber-900">
                    ⚠ Orphaned — 目標資料夾不存在或 ID 已失效。請編輯重新指定。
                  </div>
                )}
              </section>

              <div className="border-t border-border" />

              {/* Actions */}
              <div className="flex flex-wrap gap-2">
                <Button size="sm" variant="outline" onClick={onStartEdit} disabled={!tree}>
                  <Pencil className="size-3.5" />
                  編輯
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => void onToggle(rule, !rule.enabled)}
                >
                  {rule.enabled ? <PowerOff className="size-3.5" /> : <Power className="size-3.5" />}
                  {rule.enabled ? '停用' : '啟用'}
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="text-destructive hover:text-destructive ml-auto"
                  onClick={() => setConfirmDelete(true)}
                >
                  <Trash2 className="size-3.5" />
                  刪除
                </Button>
              </div>
            </>
          )}
        </div>
      </div>
      {confirmDelete && (
        <ConfirmDialog
          title="刪除這條規則?"
          description="會寫入墓碑、AI 不會再自動生成同樣的規則。"
          confirmLabel="刪除"
          danger
          onConfirm={async () => {
            await onDelete(rule.id)
            setConfirmDelete(false)
          }}
          onCancel={() => setConfirmDelete(false)}
        />
      )}
    </div>
  )
}

function Field({
  label,
  children,
  tight,
}: {
  label: string
  children: ReactNode
  tight?: boolean
}) {
  return (
    <div className={cn('flex', tight ? 'flex-col gap-0.5' : 'items-baseline gap-2')}>
      <span
        className={cn(
          'text-[10px] uppercase tracking-wide text-muted-foreground',
          tight ? '' : 'min-w-[80px]',
        )}
      >
        {label}
      </span>
      <span className="text-xs">{children}</span>
    </div>
  )
}

// ============================================================
// Rule Library — Shared row + grouped table (Phase 5 redesign)
// ============================================================

/**
 * Per-type ordering used inside each folder group. Matches the
 * structural-ID-first preference: case_code > compound > domain >
 * subject_keyword > sender. Mirrors TYPE_PRIORITY in shared/rules.ts.
 */
const RULE_TYPE_PRIORITY: Record<RuleType, number> = {
  case_code: 1,
  compound: 2,
  domain: 3,
  subject_keyword: 4,
  sender: 5,
}

/**
 * One table row. Reused by both flat view (`showTarget`=true to show
 * the target-folder column) and grouped view (`showTarget`=false —
 * the target is in the group header).
 */
// Perf (audit): memoized with id/rule-based callbacks. The 全部規則 table
// renders up to 200-800 rows; with inline per-row arrows every search
// keystroke / filter change / checkbox toggle re-rendered EVERY row (each
// running refinedTypeLabel's regex extraction twice). With stable callbacks
// from the parents, memo lets unchanged rows skip re-render entirely.
const RuleRow = memo(function RuleRow({
  rule,
  selected,
  highlighted,
  showTarget,
  onSelect,
  onOpenDetail,
  onToggle,
}: {
  rule: Rule
  selected: boolean
  highlighted: boolean
  showTarget: boolean
  onSelect: (id: string) => void
  onOpenDetail: (id: string) => void
  onToggle: (rule: Rule) => void
}) {
  const typeLabel = refinedTypeLabel(rule.type, rule.signal)
  const openDetail = () => onOpenDetail(rule.id)
  const accuracyPct =
    rule.matchCount > 0
      ? Math.round(((rule.matchCount - (rule.overrideCount ?? 0)) / rule.matchCount) * 100)
      : null
  const accuracyClass =
    accuracyPct === null
      ? 'text-muted-foreground'
      : accuracyPct >= 90
        ? 'text-emerald-700'
        : accuracyPct >= 70
          ? 'text-amber-700'
          : 'text-red-700'
  // A11y: keyboard activation. The row is focusable (tabIndex=0); Enter
  // and Space open the detail drawer. We DON'T put role="button" on the
  // <tr> because that breaks the grid/row screen reader semantics. Instead
  // we lean on the focusable + onKeyDown contract, plus an explicit
  // aria-label so screen readers announce what the row is for.
  //
  // Why a row click is the activation gesture: every body cell already
  // had onClick={openDetail} via mouse. The keyboard equivalent is to
  // land focus on the row and press Enter. The toggle button + checkbox
  // intercept their own clicks via stopPropagation so they remain
  // independent targets.
  function handleRowKey(e: ReactKeyboardEvent<HTMLTableRowElement>) {
    if (e.target !== e.currentTarget) return // event came from a child
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault()
      openDetail()
    }
  }
  return (
    <tr
      tabIndex={0}
      onKeyDown={handleRowKey}
      aria-label={`${typeLabel} ${rule.signal} → ${rule.targetFolderPath}`}
      className={cn(
        'border-t border-border hover:bg-accent/30 transition-colors',
        'focus:outline-none focus:ring-2 focus:ring-foreground/20 focus:ring-inset',
        !rule.enabled && 'opacity-50',
        highlighted && 'bg-accent/50',
      )}
    >
      <td className="px-2 py-1.5">
        <input
          type="checkbox"
          checked={selected}
          onChange={() => onSelect(rule.id)}
          className="accent-foreground"
          aria-label={`選取 ${rule.signal}`}
          onClick={(e) => e.stopPropagation()}
        />
      </td>
      <td className="px-2 py-1.5 cursor-pointer" onClick={openDetail}>
        <Badge variant="outline" className="font-mono text-[9px]">
          {typeLabel}
        </Badge>
      </td>
      <td className="px-2 py-1.5 cursor-pointer max-w-0" onClick={openDetail}>
        <div className="font-mono truncate" title={rule.signal}>
          {rule.type === 'compound' ? formatCompoundSignal(rule.signal) : rule.signal}
        </div>
      </td>
      {showTarget && (
        <td
          className="px-2 py-1.5 cursor-pointer max-w-0 text-muted-foreground"
          onClick={openDetail}
        >
          <div className="truncate" title={rule.targetFolderPath}>
            {rule.targetFolderPath}
          </div>
        </td>
      )}
      <td
        className="px-2 py-1.5 text-right font-mono tabular-nums cursor-pointer"
        onClick={openDetail}
      >
        <span>{rule.matchCount}</span>
        {accuracyPct !== null && (
          <span className={cn('ml-1 text-[10px]', accuracyClass)}>
            {accuracyPct}%
          </span>
        )}
      </td>
      <td
        className="px-2 py-1.5 text-right font-mono tabular-nums cursor-pointer"
        onClick={openDetail}
      >
        {rule.confidence.toFixed(2)}
      </td>
      <td className="px-2 py-1.5 cursor-pointer" onClick={openDetail}>
        <span className="text-[10px] text-muted-foreground">
          {SOURCE_LABEL[rule.source]}
        </span>
      </td>
      <td className="px-2 py-1.5">
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation()
            onToggle(rule)
          }}
          title={rule.enabled ? '停用' : '啟用'}
          aria-label={`${rule.enabled ? '停用' : '啟用'}規則 ${rule.signal}`}
          className="text-muted-foreground hover:text-foreground"
        >
          {rule.enabled ? <Power className="size-3.5" /> : <PowerOff className="size-3.5" />}
        </button>
      </td>
    </tr>
  )
})

/**
 * Grouped table — rules clustered by target folder. Each folder
 * shows up as a header with aggregate stats (count, hits, accuracy)
 * + a hover-reveal "stop all rules in this folder" action.
 *
 * Order:
 *   1. Folder groups sorted by total enabled matchCount (DESC),
 *      tiebreak by folder path (ASC) for deterministic layout.
 *   2. Inside each group, rules ordered by RULE_TYPE_PRIORITY,
 *      tiebreak by matchCount (DESC).
 *
 * Stat aggregates use ENABLED rules only — matches what the user
 * sees by default ("狀態:啟用" filter). If the user has disabled
 * rules visible (via filter), disabled rows still render but don't
 * contribute to the per-folder accuracy / hit number.
 */
function RuleGroupedTable({
  rules,
  allRules,
  selected,
  detailId,
  collapsed,
  setCollapsed,
  onToggleSelectAll,
  allFilteredSelected,
  onSelect,
  onOpenDetail,
  onToggle,
}: {
  rules: Rule[]
  allRules: Rule[]
  selected: Set<string>
  detailId: string | null
  collapsed: Set<string>
  setCollapsed: (v: Set<string>) => void
  onToggleSelectAll: () => void
  allFilteredSelected: boolean
  onSelect: (id: string) => void
  onOpenDetail: (id: string) => void
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
}) {
  void allRules // currently unused; reserved for future "this folder also has N hidden rules" hint
  // Stable wrapper so the memoized RuleRow's onToggle prop keeps identity
  // across re-renders (see RuleRow's memo note).
  const handleToggleRule = useCallback(
    (r: Rule) => {
      void onToggle(r, !r.enabled)
    },
    [onToggle],
  )
  const groups = useMemo(() => {
    const map = new Map<string, Rule[]>()
    for (const r of rules) {
      const arr = map.get(r.targetFolderPath) ?? []
      arr.push(r)
      map.set(r.targetFolderPath, arr)
    }
    type Group = {
      path: string
      rules: Rule[]
      totalHits: number
      totalOverrides: number
      enabledCount: number
    }
    const out: Group[] = []
    for (const [path, rs] of map.entries()) {
      const sorted = [...rs].sort((a, b) => {
        const p = RULE_TYPE_PRIORITY[a.type] - RULE_TYPE_PRIORITY[b.type]
        if (p !== 0) return p
        return b.matchCount - a.matchCount
      })
      const enabled = sorted.filter((r) => r.enabled)
      out.push({
        path,
        rules: sorted,
        totalHits: enabled.reduce((s, r) => s + r.matchCount, 0),
        totalOverrides: enabled.reduce((s, r) => s + (r.overrideCount ?? 0), 0),
        enabledCount: enabled.length,
      })
    }
    // Active folders first (more hits = more relevant), then alphabetical.
    out.sort((a, b) => {
      if (a.totalHits !== b.totalHits) return b.totalHits - a.totalHits
      return a.path.localeCompare(b.path)
    })
    return out
  }, [rules, allRules])

  function toggleCollapse(path: string) {
    const next = new Set(collapsed)
    if (next.has(path)) next.delete(path)
    else next.add(path)
    setCollapsed(next)
  }

  return (
    <div className="space-y-3">
      {/* Top-level select-all row, mirrors the flat-table thead */}
      <div className="flex items-center gap-2 text-[11px] text-muted-foreground px-1">
        <input
          type="checkbox"
          checked={allFilteredSelected}
          onChange={onToggleSelectAll}
          className="accent-foreground"
          aria-label="全選顯示中的規則"
        />
        <span>選取顯示中的全部規則 ({rules.length} 條)</span>
      </div>

      {groups.map((g) => {
        const isCollapsed = collapsed.has(g.path)
        const accPct =
          g.totalHits > 0
            ? Math.round(((g.totalHits - g.totalOverrides) / g.totalHits) * 100)
            : null
        const accClass =
          accPct === null
            ? 'text-muted-foreground'
            : accPct >= 90
              ? 'text-emerald-700'
              : accPct >= 70
                ? 'text-amber-700'
                : 'text-red-700'
        return (
          <div
            key={g.path}
            className="rounded-md border border-border overflow-hidden bg-background"
          >
            {/* Folder header — collapse toggle is the main button; bulk
                action is a sibling button (not nested) to keep the HTML
                semantically valid. Both live in a flex row that handles
                hover styling via the `group` class on the wrapper. */}
            <div className="bg-muted/40 hover:bg-muted/60 transition-colors flex items-center gap-2 px-3 py-2 group">
              <button
                type="button"
                onClick={() => toggleCollapse(g.path)}
                className="flex items-center gap-2 text-left flex-1 min-w-0"
                aria-expanded={!isCollapsed}
                aria-label={`${isCollapsed ? '展開' : '摺疊'} ${g.path}`}
              >
                <ChevronDown
                  className={cn(
                    'size-3.5 text-muted-foreground transition-transform shrink-0',
                    isCollapsed && '-rotate-90',
                  )}
                  aria-hidden="true"
                />
                <span className="font-mono text-xs truncate flex-1" title={g.path}>
                  {g.path}
                </span>
                <span className="text-[10px] tabular-nums text-muted-foreground shrink-0 hidden sm:inline">
                  {g.rules.length} 條 · {g.totalHits.toLocaleString()} 命中
                  {accPct !== null && (
                    <>
                      {' · '}
                      <span className={accClass}>{accPct}%</span>
                    </>
                  )}
                </span>
              </button>
              {/* Bulk-disable: hover-reveal on pointer; always visible
                  via focus so keyboard users still reach it. */}
              {g.enabledCount > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    if (!confirm(`停用「${g.path}」底下 ${g.enabledCount} 條啟用中的規則?`)) return
                    void (async () => {
                      for (const r of g.rules) {
                        if (r.enabled) {
                          try {
                            await onToggle(r, false)
                          } catch (err) {
                            console.warn(
                              '[mail-organizer] folder bulk-disable item failed',
                              r.id,
                              err,
                            )
                          }
                        }
                      }
                    })()
                  }}
                  className="opacity-0 group-hover:opacity-100 focus:opacity-100 transition-opacity text-[10px] px-2 py-0.5 rounded border border-border bg-background hover:bg-accent shrink-0"
                  title="停用此資料夾的所有啟用規則"
                >
                  全部停用
                </button>
              )}
            </div>
            {/* Body — table without the target column */}
            {!isCollapsed && (
              <table className="w-full text-xs">
                <thead className="bg-muted/20 border-t border-border">
                  <tr className="text-left text-[10px] uppercase tracking-wide text-muted-foreground">
                    <th className="px-2 py-1.5 w-8"></th>
                    <th className="px-2 py-1.5 w-20">類型</th>
                    <th className="px-2 py-1.5">Signal</th>
                    <th className="px-2 py-1.5 w-16 text-right">命中</th>
                    <th className="px-2 py-1.5 w-20 text-right">Conf</th>
                    <th className="px-2 py-1.5 w-20">來源</th>
                    <th className="px-2 py-1.5 w-12"></th>
                  </tr>
                </thead>
                <tbody>
                  {g.rules.map((r) => (
                    <RuleRow
                      key={r.id}
                      rule={r}
                      selected={selected.has(r.id)}
                      highlighted={detailId === r.id}
                      showTarget={false}
                      onSelect={onSelect}
                      onOpenDetail={onOpenDetail}
                      onToggle={handleToggleRule}
                    />
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )
      })}
    </div>
  )
}

// ============================================================
// Rule Library — Conflicts view (Phase 3 standalone)
// ============================================================

type ConflictsViewProps = {
  rules: Rule[]
  conflicts: Array<{ type: string; signal: string; ruleIds: string[]; targets: string[] }>
  tree: MailFolderNode[] | null
  onResolveConflict: (
    ruleIds: string[],
    strategy: 'keep_highest' | 'disable_all' | 'keep_one',
    keepId?: string,
  ) => Promise<void>
  /** Jump to 全部規則 view with this rule's detail drawer + editor open. */
  onEditRule: (ruleId: string) => void
  onSplitConflict: (items: Array<{ ruleId: string; keywords: string[] }>) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
  onAutoUpgradeConflict: (ruleIds: string[]) => Promise<{ ok: true; data?: { created: number; disabled: number } } | { ok: false; code: string; message: string }>
  onSuggestKeywords: (folderIds: string[]) => Promise<{ ok: true; data?: { suggestions: string[][]; recentSubjects: string[][] } } | { ok: false; code: string; message: string }>
}

/**
 * Standalone conflicts view. Each conflict group is shown as a card
 * with the auto-upgrade button as primary action; destructive
 * fallbacks (keep highest, disable all) are tucked behind 「進階」.
 */
function ConflictsView({
  rules,
  conflicts,
  onResolveConflict,
  onAutoUpgradeConflict,
  onEditRule,
}: ConflictsViewProps) {
  type UpgradeState =
    | { status: 'idle' }
    | { status: 'pending' }
    | { status: 'done'; created: number; disabled: number }
    | { status: 'error'; message: string }
  const [upgradeStateByKey, setUpgradeStateByKey] = useState<
    Record<string, UpgradeState>
  >({})

  const conflictKey = (c: { ruleIds: string[] }): string =>
    [...c.ruleIds].sort().join('|')

  if (conflicts.length === 0) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="text-sm">衝突</CardTitle>
          <CardDescription>
            同類型同條件、但指向不同資料夾的規則。新設計下會在學習階段自動防衝突,衝突清單應該很少出現;若出現通常是歷史規則或跨機器同步帶入。
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="text-center py-12 text-sm text-muted-foreground">
            <CheckCircle2 className="size-10 mx-auto mb-3 text-emerald-600" />
            <p className="font-medium text-foreground">目前沒有衝突</p>
            <p className="text-[11px] mt-1">規則庫運作正常,學習階段已自動防衝突</p>
          </div>
        </CardContent>
      </Card>
    )
  }

  return (
    <div className="space-y-4">
      <div className="rounded-md border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900">
        <p className="font-medium flex items-center gap-1.5">
          <AlertTriangle className="size-3.5" />
          發現 {conflicts.length} 個規則衝突
        </p>
        <p className="text-[11px] mt-1">
          建議用「自動升級為 compound 規則」處理 — 系統會抓取每條規則的目標資料夾近期信件主旨,自動建立 compound (網域/寄件人 + 主旨) 規則並停用舊規則。
        </p>
      </div>

      <ul className="space-y-3">
        {conflicts.map((c) => {
          const cKey = conflictKey(c)
          const upgrade = upgradeStateByKey[cKey] ?? { status: 'idle' as const }
          const supportsAutoUpgrade = c.type === 'domain' || c.type === 'sender'
          // Resolve the group's full Rule objects so each side of the
          // conflict shows its evidence (confidence / hits / accuracy /
          // source) and offers per-rule actions — the card used to show
          // bare target paths with no way to edit or to know which rule
          // 「保留最高」 would keep.
          const group = c.ruleIds
            .map((id) => rules.find((r) => r.id === id))
            .filter((r): r is Rule => !!r)
          // EXACTLY the SW resolver's keep_highest selection (sort by
          // confidence desc; stable sort breaks ties by group order) — so
          // the「會保留」badge always matches what the button really does.
          const winner =
            group.length > 0
              ? [...group].sort((a, b) => b.confidence - a.confidence)[0]
              : undefined

          return (
            <Card key={cKey}>
              <CardContent className="p-4 space-y-3">
                <div>
                  <div className="flex items-center gap-2 text-xs">
                    <Badge variant="outline" className="font-mono text-[10px]">
                      {TYPE_LABEL[c.type as RuleType] ?? c.type}
                    </Badge>
                    <code className="font-mono text-xs">{c.signal}</code>
                  </div>
                  <ul className="mt-2 space-y-1">
                    {group.map((r) => {
                      const acc =
                        r.matchCount > 0
                          ? Math.round(
                              ((r.matchCount - (r.overrideCount ?? 0)) / r.matchCount) * 100,
                            )
                          : null
                      return (
                        <li
                          key={r.id}
                          className="flex items-center gap-2 flex-wrap rounded border border-border bg-muted/20 px-2 py-1.5 text-[11px]"
                        >
                          <span
                            className="font-mono flex-1 min-w-[160px] truncate"
                            title={r.targetFolderPath}
                          >
                            → {r.targetFolderPath}
                          </span>
                          {winner?.id === r.id && group.length > 1 && (
                            <Badge
                              variant="outline"
                              className="text-[9px] shrink-0 border-emerald-400 text-emerald-700"
                              title="信心值最高 —「保留最高、刪其餘」會保留這條"
                            >
                              保留最高會留這條
                            </Badge>
                          )}
                          {!r.enabled && (
                            <Badge variant="outline" className="text-[9px] shrink-0">
                              已停用
                            </Badge>
                          )}
                          <span className="font-mono text-muted-foreground shrink-0 tabular-nums">
                            conf {r.confidence.toFixed(2)}
                          </span>
                          <span className="font-mono text-muted-foreground shrink-0 tabular-nums">
                            命中 {r.matchCount}
                            {acc !== null ? ` · ${acc}%` : ''}
                          </span>
                          <span className="text-[10px] text-muted-foreground shrink-0">
                            {SOURCE_LABEL[r.source]}
                          </span>
                          <button
                            type="button"
                            onClick={() => onEditRule(r.id)}
                            className="underline underline-offset-2 text-muted-foreground hover:text-foreground shrink-0"
                            title="到「全部規則」開啟這條規則的編輯抽屜"
                          >
                            編輯
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              const others = group.length - 1
                              if (
                                !confirm(
                                  `保留「${r.targetFolderPath}」這條、刪除其餘 ${others} 條?\n(刪除會寫入墓碑、AI 不會再自動學回同樣的規則)`,
                                )
                              ) {
                                return
                              }
                              void onResolveConflict(c.ruleIds, 'keep_one', r.id)
                            }}
                            className="underline underline-offset-2 text-red-700 hover:text-red-800 shrink-0"
                            title="手動指定保留這條，刪除群組內其他規則"
                          >
                            保留這條
                          </button>
                        </li>
                      )
                    })}
                    {/* Stale conflict data (rule ids no longer resolvable
                        locally) — fall back to the bare target list. */}
                    {group.length === 0 &&
                      c.targets.map((t, i) => (
                        <li key={i} className="font-mono text-[11px] text-muted-foreground pl-3">
                          → {t}
                        </li>
                      ))}
                  </ul>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  {supportsAutoUpgrade && (
                    <Button
                      size="sm"
                      disabled={upgrade.status === 'pending'}
                      onClick={() => {
                        setUpgradeStateByKey((s) => ({
                          ...s,
                          [cKey]: { status: 'pending' },
                        }))
                        void onAutoUpgradeConflict(c.ruleIds).then((r) => {
                          if (r.ok && r.data) {
                            setUpgradeStateByKey((s) => ({
                              ...s,
                              [cKey]: {
                                status: 'done',
                                created: r.data!.created,
                                disabled: r.data!.disabled,
                              },
                            }))
                          } else if (!r.ok) {
                            setUpgradeStateByKey((s) => ({
                              ...s,
                              [cKey]: { status: 'error', message: r.message },
                            }))
                          }
                        })
                      }}
                    >
                      {upgrade.status === 'pending'
                        ? '升級中…'
                        : '自動升級為 compound 規則'}
                    </Button>
                  )}
                  {upgrade.status === 'done' && (
                    <span className="text-[11px] text-emerald-700">
                      ✓ 新增 {upgrade.created} 條 · 停用 {upgrade.disabled} 條
                    </span>
                  )}
                  {upgrade.status === 'error' && (
                    <span className="text-[11px] text-red-700">
                      ✗ {upgrade.message}
                    </span>
                  )}
                  <details className="text-[11px] ml-auto">
                    <summary className="cursor-pointer text-muted-foreground hover:text-foreground">
                      進階
                    </summary>
                    <div className="mt-2 flex gap-1 flex-wrap">
                      <Button
                        size="sm"
                        variant="outline"
                        className="text-[10px] h-7"
                        onClick={() => void onResolveConflict(c.ruleIds, 'disable_all')}
                      >
                        全部停用
                      </Button>
                      <Button
                        size="sm"
                        variant="outline"
                        className="text-[10px] h-7 border-red-300 text-red-700 hover:bg-red-50"
                        disabled={!winner}
                        title={
                          winner
                            ? `會保留信心最高的「${winner.targetFolderPath}」(conf ${winner.confidence.toFixed(2)})`
                            : '無法解析衝突群組的規則'
                        }
                        onClick={() => {
                          if (!winner) return
                          if (
                            !confirm(
                              `保留信心最高的「${winner.targetFolderPath}」(conf ${winner.confidence.toFixed(2)})、刪除其餘 ${group.length - 1} 條?\n(刪除會寫入墓碑、AI 不會再自動學回同樣的規則)`,
                            )
                          ) {
                            return
                          }
                          void onResolveConflict(c.ruleIds, 'keep_highest')
                        }}
                      >
                        保留最高、刪其餘
                      </Button>
                    </div>
                  </details>
                </div>
              </CardContent>
            </Card>
          )
        })}
      </ul>
    </div>
  )
}

const TYPE_PLACEHOLDER: Record<RuleType, string> = {
  case_code: '25A0067A',
  domain: 'vendor.example（不需要 @）',
  compound: '',
  subject_keyword: '工時審閱',
  sender: '例:contact@example.com',
}

const COMPOUND_TYPE_LABEL: Record<CompoundCondition['type'], string> = {
  sender: '寄件人',
  domain: '網域',
  subject_keyword: '主旨關鍵字',
}

function formatCompoundSignal(signal: string): string {
  const parsed = decodeCompound(signal)
  if (!parsed) return '(無效 JSON)'
  return parsed.conditions.map((c) => `${COMPOUND_TYPE_LABEL[c.type]}=${c.value}`).join(' AND ')
}

type RuleFormValues = {
  type: RuleType
  signal: string
  targetFolderId: string
  targetFolderPath: string
  confidence: number
  enabled: boolean
}

type RuleFormProps = {
  tree: MailFolderNode[]
  initial?: Rule
  /** Existing rules for proactive conflict detection. Optional — if not
   *  provided, the form skips conflict checking. */
  rules?: Rule[]
  onCancel: () => void
  onSubmit: (values: RuleFormValues) => Promise<{ ok: true; data?: unknown } | { ok: false; code: string; message: string }>
}

function RuleForm({ tree, initial, rules, onCancel, onSubmit }: RuleFormProps) {
  const [type, setType] = useState<RuleType>(initial?.type ?? 'domain')
  const [signal, setSignal] = useState(initial?.signal ?? '')
  const [target, setTarget] = useState<{ id: string; path: string } | null>(
    initial ? { id: initial.targetFolderId, path: initial.targetFolderPath } : null,
  )
  const [confidence, setConfidence] = useState(initial?.confidence ?? 0.85)
  const [enabled, setEnabled] = useState(initial?.enabled ?? true)
  const [error, setError] = useState<string | null>(null)

  // ---- State-sync on rule switch ------------------------------------------
  //
  // useState only reads `initial` on first render. Without this effect, two
  // sequential edits in the same drawer session — say, click rule A's
  // "編輯" → cancel → click rule B's "編輯" — would show rule A's old
  // values when editing B, because the form mounted once for A and is
  // being reused for B.
  //
  // We key the reset on `initial?.id` so:
  //   - Switching to a different rule resets every field (intended)
  //   - The SAME rule getting a sync-driven update mid-edit does NOT
  //     reset (so the user's in-flight typing isn't clobbered by the
  //     pulled values). The next time they cancel and re-edit, they'll
  //     see the fresh values.
  //   - Switching from edit-existing to add-new (initial becomes
  //     undefined) clears everything.
  const lastInitialIdRef = useRef<string | undefined>(initial?.id)
  useEffect(() => {
    const nextId = initial?.id
    if (nextId === lastInitialIdRef.current) return
    lastInitialIdRef.current = nextId
    setType(initial?.type ?? 'domain')
    setSignal(initial?.signal ?? '')
    setTarget(
      initial
        ? { id: initial.targetFolderId, path: initial.targetFolderPath }
        : null,
    )
    setConfidence(initial?.confidence ?? 0.85)
    setEnabled(initial?.enabled ?? true)
    setError(null)
  }, [initial])
  // Conflict warning derived from the current draft against existing
  // rules. Computed live as the user types so they see "this same
  // signal already points at folder X" before saving. Excludes the
  // rule being edited so editing a rule doesn't flag itself as a
  // conflict with itself.
  const conflicts = useMemo<Rule[]>(() => {
    if (!rules || rules.length === 0) return []
    if (!signal.trim()) return []
    const sigNorm =
      type === 'compound' ? signal : signal.toLowerCase().trim().replace(/^@/, '')
    return rules.filter(
      (r) =>
        r.id !== initial?.id &&
        r.type === type &&
        (r.type === 'compound'
          ? r.signal === sigNorm
          : r.signal.toLowerCase().trim().replace(/^@/, '') === sigNorm) &&
        r.targetFolderPath !== (target?.path ?? '__none__'),
    )
  }, [rules, type, signal, target, initial?.id])

  function changeType(next: RuleType) {
    if (next === type) return

    // Smart conversion so editing a rule's type doesn't always wipe the value:
    //   simple → compound : wrap current signal as the first condition so user
    //     can immediately add subject_keyword on top
    //   compound → simple : pull first condition's value if its type matches,
    //     otherwise clear and let user re-enter
    //   simple → simple (e.g. domain → sender) : keep signal as-is
    //   anything involving case_code : reset (case_code can't be compound)
    if (next === 'compound' && type !== 'compound') {
      const compatible = type === 'domain' || type === 'sender' || type === 'subject_keyword'
      const conditions =
        compatible && signal.trim()
          ? [{ type, value: signal.trim() } as CompoundCondition]
          : []
      setSignal(conditions.length > 0 ? encodeCompound(conditions) : '')
    } else if (type === 'compound' && next !== 'compound') {
      const parsed = decodeCompound(signal)
      const first = parsed?.conditions[0]
      if (
        first &&
        (next === 'domain' || next === 'sender' || next === 'subject_keyword') &&
        first.type === next
      ) {
        setSignal(first.value)
      } else {
        setSignal('')
      }
    } else if (next === 'case_code' || type === 'case_code') {
      // case_code can't share signal format with the other simple types
      setSignal('')
    }
    // simple ↔ simple (domain / sender / subject_keyword): keep signal

    setType(next)
  }

  async function submit() {
    setError(null)
    if (!target) {
      setError('請選擇目標資料夾')
      return
    }
    if (type === 'compound') {
      const parsed = decodeCompound(signal)
      if (!parsed || parsed.conditions.length === 0) {
        setError('複合規則至少要有 1 個條件')
        return
      }
      if (parsed.conditions.some((c) => !c.value.trim())) {
        setError('每個條件的值都不能空')
        return
      }
    } else if (!signal.trim()) {
      setError('signal 不可空')
      return
    }
    const r = await onSubmit({
      type,
      signal: type === 'compound' ? signal : signal.trim(),
      targetFolderId: target.id,
      targetFolderPath: target.path,
      confidence,
      enabled,
    })
    if (!r.ok) setError(r.message)
  }

  return (
    <div className="rounded-md border border-border bg-muted/30 p-3 space-y-3 text-xs">
      <div className="flex items-center gap-2 flex-wrap">
        <label className="text-[10px] text-muted-foreground">類型</label>
        <select
          value={type}
          onChange={(e) => changeType(e.target.value as RuleType)}
          className="rounded-md border border-input bg-background px-2 py-1 text-xs"
        >
          <option value="case_code">案件代號</option>
          <option value="domain">網域</option>
          <option value="subject_keyword">主旨關鍵字</option>
          <option value="sender">寄件人</option>
          <option value="compound">複合（多條件）</option>
        </select>
        {initial && type !== 'compound' && (
          <span className="text-[10px] text-muted-foreground">
            升級為「複合」可加主旨條件
          </span>
        )}
        {initial && type === 'compound' && (
          <span className="text-[10px] text-muted-foreground">已升級為多條件、可繼續加</span>
        )}
      </div>

      {type === 'compound' ? (
        <CompoundEditor value={signal} onChange={setSignal} />
      ) : (
        <div className="space-y-1">
          <label className="text-[10px] text-muted-foreground">signal</label>
          <Input
            value={signal}
            onChange={(e) => setSignal(e.target.value)}
            placeholder={TYPE_PLACEHOLDER[type]}
            className="text-xs font-mono"
          />
        </div>
      )}

      <div className="space-y-1">
        <label className="text-[10px] text-muted-foreground">目標資料夾</label>
        <FolderPicker
          tree={tree}
          value={target?.path}
          onSelect={(n) => setTarget({ id: n.id, path: n.path })}
          excludePrefixes={['05已完成案件']}
        />
      </div>

      <div className="flex items-center gap-2">
        <label className="text-[10px] text-muted-foreground">信心</label>
        <input
          type="range"
          min={0.5}
          max={1}
          step={0.05}
          value={confidence}
          onChange={(e) => setConfidence(Number(e.target.value))}
          className="flex-1 accent-foreground"
        />
        <span className="font-mono tabular-nums w-10 text-right">{confidence.toFixed(2)}</span>
      </div>

      {initial && (
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={enabled}
            onChange={(e) => setEnabled(e.target.checked)}
            className="accent-foreground"
          />
          <span>啟用</span>
        </label>
      )}

      {error && <p className="text-red-700">{error}</p>}

      {/* Proactive conflict warning — non-blocking; rules.ts at match time
          uses TYPE_PRIORITY + effectiveConfidence to pick a winner, but the
          user usually means "this rule wins" so we flag the surprise. */}
      {conflicts.length > 0 && (
        <div className="rounded-md border border-amber-300 bg-amber-50 p-2 text-[10px] text-amber-900 space-y-1">
          <div className="font-medium flex items-center gap-1">
            <AlertTriangle className="size-3" />
            {conflicts.length === 1
              ? '已有相同 signal 但不同目標的規則'
              : `已有 ${conflicts.length} 條相同 signal 但不同目標的規則`}
          </div>
          <ul className="space-y-0.5 pl-4">
            {conflicts.slice(0, 3).map((r) => (
              <li key={r.id} className="font-mono truncate">
                → {r.targetFolderPath}{' '}
                <span className="text-amber-700/80">
                  ({r.source}
                  {r.enabled ? '' : '・已停用'})
                </span>
              </li>
            ))}
            {conflicts.length > 3 && (
              <li className="text-amber-700/70">…還有 {conflicts.length - 3} 條</li>
            )}
          </ul>
          <div>
            建議:儲存後到「規則衝突」區處理、或改用 compound 規則(domain + 主旨關鍵字)分流。
          </div>
        </div>
      )}

      <div className="flex justify-end gap-2 pt-1">
        <Button size="sm" variant="ghost" onClick={onCancel}>取消</Button>
        <Button size="sm" onClick={submit}>{initial ? '儲存' : '新增'}</Button>
      </div>
    </div>
  )
}

function CompoundEditor({
  value,
  onChange,
}: {
  value: string
  onChange: (encoded: string) => void
}) {
  const [conditions, setConditions] = useState<CompoundCondition[]>(() => {
    const parsed = decodeCompound(value)
    return parsed?.conditions ?? []
  })

  function update(next: CompoundCondition[]) {
    setConditions(next)
    onChange(encodeCompound(next))
  }

  function setCondition(i: number, patch: Partial<CompoundCondition>) {
    const next = conditions.map((c, idx) =>
      idx === i ? ({ ...c, ...patch } as CompoundCondition) : c,
    )
    update(next)
  }

  function addCondition() {
    update([...conditions, { type: 'sender', value: '' }])
  }

  function removeCondition(i: number) {
    update(conditions.filter((_, idx) => idx !== i))
  }

  const PLACEHOLDER: Record<CompoundCondition['type'], string> = {
    sender: '例:contact@example.com',
    domain: 'vendor.example',
    subject_keyword: '工時審閱',
  }

  return (
    <div className="space-y-2">
      <div className="flex items-baseline justify-between">
        <label className="text-[10px] text-muted-foreground">條件清單</label>
        <span className="text-[10px] text-muted-foreground">所有條件都需滿足（AND）</span>
      </div>

      {conditions.length === 0 && (
        <p className="text-[11px] italic text-muted-foreground">尚無條件 — 至少要加 1 個</p>
      )}

      {conditions.map((c, i) => (
        <div key={i} className="flex gap-1 items-center">
          {i > 0 && <span className="text-[9px] text-muted-foreground font-mono px-1">AND</span>}
          {i === 0 && <span className="text-[9px] text-muted-foreground font-mono px-1 invisible">AND</span>}
          <select
            value={c.type}
            onChange={(e) => setCondition(i, { type: e.target.value as CompoundCondition['type'] })}
            className="rounded-md border border-input bg-background px-1.5 py-1 text-xs"
          >
            <option value="sender">寄件人</option>
            <option value="domain">網域</option>
            <option value="subject_keyword">主旨</option>
          </select>
          <Input
            value={c.value}
            onChange={(e) => setCondition(i, { value: e.target.value })}
            placeholder={PLACEHOLDER[c.type]}
            className="text-xs font-mono flex-1 h-8"
          />
          <Button
            variant="ghost"
            size="icon"
            className="size-8 shrink-0 text-muted-foreground hover:text-red-700"
            onClick={() => removeCondition(i)}
            title="移除條件"
          >
            <X className="size-3.5" />
          </Button>
        </div>
      ))}

      <Button size="sm" variant="outline" onClick={addCondition}>
        <Plus className="size-3.5" /> 新增條件
      </Button>
    </div>
  )
}

// ---- Rule effectiveness — removed 2026-05-27 ----------------------------
// The cards moved to RuleLibrarySummaryCard + RuleHealthSection inside
// the new full-screen RuleLibraryView. Local STALE_MS / STALE_DAYS
// constants were unused after the removal and were deleted.

/**
 * Recent Activity panel filter editor. Two list-of-strings fields:
 *   - Include prefixes: paths the row's folderPath must start with
 *   - Include leaf names: folder names that match by leaf regardless
 *     of where in the tree they sit
 *
 * Local "draft" state so the user can edit several lines and hit Save
 * without each keystroke pinging the SW. Empty lines + duplicates get
 * trimmed before save.
 */
/**
 * Classification preferences card (renamed from "事務所設定" 2026-05-22).
 *
 * Two zones:
 *  1. 主要根資料夾 — top-level setting, recommended for most users.
 *     Drives initial scan default + classifier prompt's example path.
 *     Picked via FolderPicker so the user always sees their actual
 *     tree, never a placeholder leaking another user's folder names.
 *  2. 內部信件規則 — collapsed-by-default optional section. Only
 *     useful for users with a workplace email (firm / company /
 *     school) where same-domain mail should route differently from
 *     external mail. Personal Gmail / Outlook users leave it empty.
 *
 * Empty everywhere is a fully-valid state: classify still works, just
 * with zero same-domain grouping and ad-hoc folder path entry at
 * scan time. Placeholders carry NO real folder names — they describe
 * the FORMAT only, so a non-lawyer doesn't see "案件" or
 * "工時/薪資/利衝" and wonder why their inbox doesn't look like that.
 */
function ClassificationPrefsCard({
  status,
  tree,
  onSaved,
}: {
  status: StatusData | null
  tree: MailFolderNode[] | null
  onSaved: () => Promise<void> | void
}) {
  const initialDomains = useMemo(
    () => (status?.internalDomains ?? []).join('\n'),
    [status?.internalDomains],
  )
  const initialRoot = useMemo(
    () => status?.primaryRootPath ?? '',
    [status?.primaryRootPath],
  )
  const initialCategories = useMemo(
    () => (status?.internalSubjectCategories ?? []).join('、'),
    [status?.internalSubjectCategories],
  )
  const [domainsDraft, setDomainsDraft] = useState(initialDomains)
  const [rootDraft, setRootDraft] = useState(initialRoot)
  const [catDraft, setCatDraft] = useState(initialCategories)
  const [savedAt, setSavedAt] = useState<number | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [detecting, setDetecting] = useState(false)
  const [detectError, setDetectError] = useState<string | null>(null)
  // Success message that surfaces the actual detected email — important
  // for the multi-account case where the user might be logged into a
  // personal account in the active OWA tab without realising. Without
  // the message they'd just see a new domain appear and not know which
  // account it came from.
  const [detectInfo, setDetectInfo] = useState<string | null>(null)
  // Internal-rules zone open state. Default open if the user already
  // has internal domains configured (returning user), default closed
  // for fresh installs so the new user isn't drawn to a feature they
  // probably don't need.
  const [internalOpen, setInternalOpen] = useState(false)

  // Re-seed drafts when status arrives / changes.
  useEffect(() => {
    setDomainsDraft(initialDomains)
    setRootDraft(initialRoot)
    setCatDraft(initialCategories)
    // Open the internal-rules zone whenever there's existing content
    // to edit. Won't toggle back closed during the same session if the
    // user opens it manually with no content.
    if (initialDomains.trim() || initialCategories.trim()) {
      setInternalOpen(true)
    }
  }, [initialDomains, initialRoot, initialCategories])

  const parseDomains = (s: string): string[] => {
    const lines = s
      .split(/[\r\n,,]+/)
      .map((l) => l.trim().toLowerCase().replace(/^@/, ''))
      .filter((l) => l.length > 0)
    return Array.from(new Set(lines))
  }
  const parseCategories = (s: string): string[] => {
    const items = s
      .split(/[\r\n,,、]+/)
      .map((l) => l.trim())
      .filter((l) => l.length > 0)
    return Array.from(new Set(items))
  }

  const dirty =
    parseDomains(domainsDraft).join('\n') !==
      parseDomains(initialDomains).join('\n') ||
    rootDraft.trim() !== initialRoot ||
    parseCategories(catDraft).join('、') !==
      parseCategories(initialCategories).join('、')

  async function save() {
    setSaveError(null)
    const r = await send({
      type: 'setSettings',
      patch: {
        internalDomains: parseDomains(domainsDraft),
        primaryRootPath: rootDraft.trim(),
        internalSubjectCategories: parseCategories(catDraft),
      },
    })
    if (r.ok) {
      setSavedAt(Date.now())
      await onSaved()
      window.setTimeout(() => setSavedAt(null), 2500)
    } else {
      setSaveError(r.message || r.code || '儲存失敗')
    }
  }

  async function autoDetect() {
    setDetectError(null)
    setDetectInfo(null)
    setDetecting(true)
    try {
      const r = await send<{ domain: string; email: string }>({
        type: 'detectUserDomain',
      })
      if (r.ok && r.data?.domain) {
        const existing = parseDomains(domainsDraft)
        if (existing.includes(r.data.domain)) {
          setDetectError(
            `已包含 @${r.data.domain}(偵測自 ${r.data.email})`,
          )
          return
        }
        setDomainsDraft([...existing, r.data.domain].join('\n'))
        // Surface the actual email so the user can spot wrong-account
        // case (multi-tab OWA where the active tab isn't the firm one).
        setDetectInfo(
          `已加入 @${r.data.domain}(偵測自 ${r.data.email})。若這不是預期帳號、請手動編輯。`,
        )
      } else if (!r.ok) {
        setDetectError(r.message || r.code || '偵測失敗')
      }
    } catch (e) {
      setDetectError(e instanceof Error ? e.message : String(e))
    } finally {
      setDetecting(false)
    }
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">歸類偏好</CardTitle>
        <CardDescription>
          告訴系統你常用的資料夾結構,讓 AI 給出更貼近你習慣的歸類建議。全部留空也能跑,只是 AI 沒有額外脈絡可參考。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        {/* ---- Primary root folder (top priority) ----------------- */}
        <div className="space-y-1">
          <label className="text-xs text-muted-foreground">
            主要根資料夾(建議設定)
          </label>
          {tree && tree.length > 0 ? (
            <FolderPicker
              tree={tree}
              value={rootDraft}
              placeholder="搜尋並選擇你的主要工作資料夾…"
              onSelect={(n) => setRootDraft(n.path)}
            />
          ) : (
            <Input
              value={rootDraft}
              onChange={(e) => setRootDraft(e.target.value)}
              placeholder="先到「資料夾快取」按重新偵測、就會出現可選清單"
              className="font-mono text-xs"
              disabled
            />
          )}
          <p className="text-[10px] text-muted-foreground">
            初始掃描預設從這個資料夾開始、AI prompt 的範例路徑也會以這個為字首。
            可隨時清空或更換。
          </p>
        </div>

        {/* ---- Optional: internal-mail rules ---------------------- */}
        <div className="rounded-md border border-border">
          <button
            type="button"
            onClick={() => setInternalOpen((v) => !v)}
            className="w-full flex items-center justify-between px-3 py-2 text-left hover:bg-accent/40 transition-colors"
            aria-expanded={internalOpen}
          >
            <span className="text-xs">
              <span className="font-medium">內部信件規則</span>
              <span className="ml-1.5 text-[10px] text-muted-foreground">(選用)</span>
            </span>
            <ChevronDown
              className={cn(
                'size-3.5 text-muted-foreground transition-transform',
                internalOpen && 'rotate-180',
              )}
            />
          </button>
          {internalOpen && (
            <div className="px-3 pb-3 pt-1 space-y-3 border-t border-border">
              <p className="text-[10px] text-muted-foreground leading-relaxed">
                若你有工作信箱(公司、學校等共用網域的組織)、可在這裡告訴 AI
                哪些網域是「內部」、避免內部寄件人地址被當成獨立分類學成規則。
                個人帳號通常不需要。
              </p>

              <div className="space-y-1">
                <div className="flex items-center justify-between">
                  <label className="text-xs text-muted-foreground">
                    內部網域(一行一個,可加多個)
                  </label>
                  <button
                    type="button"
                    onClick={() => void autoDetect()}
                    disabled={detecting || !status?.tokenValid}
                    className="text-[10px] underline underline-offset-2 hover:text-foreground disabled:opacity-40 disabled:no-underline"
                    title={
                      status?.tokenValid
                        ? '從你登入的 Outlook 帳號偵測網域'
                        : '需要先在 OWA 分頁登入'
                    }
                  >
                    {detecting ? '偵測中…' : '自動偵測'}
                  </button>
                </div>
                <textarea
                  value={domainsDraft}
                  onChange={(e) => setDomainsDraft(e.target.value)}
                  rows={2}
                  className="w-full rounded-md border border-border bg-background px-2 py-1.5 text-xs font-mono"
                  placeholder="一行一個網域(例如直接按「自動偵測」帶入)"
                />
                {detectError && (
                  <p className="text-[10px] text-red-700">{detectError}</p>
                )}
                {detectInfo && (
                  <p className="text-[10px] text-emerald-700">{detectInfo}</p>
                )}
                <p className="text-[10px] text-muted-foreground">
                  這些網域的寄件人會被視為「內部」、不會被學成獨立網域規則。
                </p>
              </div>

              <div className="space-y-1">
                <label className="text-xs text-muted-foreground">
                  內部信件分類提示(可選,逗號或頓號分隔)
                </label>
                <Input
                  value={catDraft}
                  onChange={(e) => setCatDraft(e.target.value)}
                  placeholder="輸入幾個你想用來分類內部信的關鍵字"
                  className="font-mono text-xs"
                />
                <p className="text-[10px] text-muted-foreground">
                  寫進 AI prompt 當提示,讓內部信能依主旨對應到你的資料夾。留空 AI 會自行判斷。
                </p>
              </div>
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          <Button onClick={save} disabled={!dirty}>
            儲存
          </Button>
          {savedAt && <span className="text-[10px] text-emerald-700">已儲存</span>}
          {saveError && <span className="text-[10px] text-red-700">{saveError}</span>}
        </div>
      </CardContent>
    </Card>
  )
}

/**
 * Settings JSON export / import (G5, 2026-05-22). Lets the user back up
 * their configuration (firm settings + display preferences + auto rules
 * thresholds) and re-apply on a fresh install or share a template with
 * a colleague at the same firm.
 *
 * What's EXCLUDED on export:
 *   - claudeApiKey (sensitive, never round-trips through file I/O)
 *
 * On import: only known Settings keys are picked up; unknown keys
 * silently dropped by the setSettings handler's whitelist. claudeApiKey
 * in the file is ignored even if present.
 */
function SettingsExportImportRow({
  status,
  onImported,
}: {
  status: StatusData | null
  onImported: () => Promise<void> | void
}) {
  const [error, setError] = useState<string | null>(null)
  const [ok, setOk] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  async function exportSettings() {
    setError(null)
    setOk(null)
    if (!status) {
      setError('狀態未載入,稍候再試')
      return
    }
    // Build a clean settings-only payload from status. status already
    // excludes claudeApiKey by virtue of the getStatus handler not
    // exposing it (only apiKeyConfigured + apiKeyPreview).
    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      settings: {
        claudeModel: status.model,
        batchSize: status.batchSize,
        excludeFolderPrefixes: status.excludePrefixes,
        aiConfidenceThreshold: status.aiConfidenceThreshold,
        skipFlagged: status.skipFlagged,
        showOwaFab: status.showOwaFab,
        prefetchNextBatch: status.prefetchNextBatch,
        recentActivityIncludePrefixes: status.recentActivityIncludePrefixes,
        recentActivityIncludeLeafNames: status.recentActivityIncludeLeafNames,
        internalDomains: status.internalDomains,
        primaryRootPath: status.primaryRootPath,
        internalSubjectCategories: status.internalSubjectCategories,
        aiIncludeFewShotExamples: status.aiIncludeFewShotExamples,
      },
    }
    let url: string | null = null
    try {
      const blob = new Blob([JSON.stringify(payload, null, 2)], {
        type: 'application/json',
      })
      url = URL.createObjectURL(blob)
      const ts = new Date().toISOString().slice(0, 10)
      const filename = `mail-organizer-settings-${ts}.json`
      if (chrome.downloads?.download) {
        await chrome.downloads.download({ url, filename, saveAs: false })
      } else {
        const a = document.createElement('a')
        a.href = url
        a.download = filename
        document.body.appendChild(a)
        a.click()
        a.remove()
      }
      setOk('設定已匯出')
      window.setTimeout(() => setOk(null), 2500)
    } catch (e) {
      setError(`匯出失敗:${e instanceof Error ? e.message : String(e)}`)
    } finally {
      if (url) window.setTimeout(() => URL.revokeObjectURL(url!), 5000)
    }
  }

  // Bump when settings-export schema changes incompatibly. Newer
  // versions can be imported but the user is warned that some fields
  // may not apply; older versions stay compatible because the
  // setSettings handler whitelist-filters unknown keys.
  const SUPPORTED_IMPORT_VERSION = 1

  async function onFilePicked(e: ChangeEvent<HTMLInputElement>) {
    setError(null)
    setOk(null)
    const file = e.target.files?.[0]
    e.target.value = '' // allow re-picking the same file
    if (!file) return
    try {
      const text = await file.text()
      const parsed = JSON.parse(text) as {
        version?: number
        settings?: Record<string, unknown>
      }
      const patch = parsed.settings
      if (!patch || typeof patch !== 'object') {
        setError('JSON 結構異常:缺少 settings 物件')
        return
      }
      const fileVersion = typeof parsed.version === 'number' ? parsed.version : 0
      if (fileVersion > SUPPORTED_IMPORT_VERSION) {
        const proceed = window.confirm(
          `設定檔版本 ${fileVersion} 比目前支援的版本 ${SUPPORTED_IMPORT_VERSION} 新、` +
            `部分欄位可能無法套用。仍要繼續嗎?`,
        )
        if (!proceed) return
      } else if (fileVersion === 0) {
        // No version field — could be a hand-edited JSON. Warn but
        // proceed; the setSettings whitelist will reject obvious junk.
        const proceed = window.confirm(
          '設定檔沒有 version 欄位、可能是手動編輯或舊版格式。仍要繼續嗎?',
        )
        if (!proceed) return
      }
      // claudeApiKey is whitelist-blocked by the setSettings handler — even
      // if someone slipped it into the file, it won't be written here.
      const r = await send({ type: 'setSettings', patch })
      if (r.ok) {
        await onImported()
        setOk('設定已匯入')
        window.setTimeout(() => setOk(null), 2500)
      } else {
        setError(r.message || r.code || '匯入失敗')
      }
    } catch (e) {
      setError(`解析失敗:${e instanceof Error ? e.message : String(e)}`)
    }
  }

  return (
    <div className="pt-1 border-t border-border space-y-1.5">
      <div className="flex flex-wrap items-center gap-2">
        <Button variant="outline" size="sm" onClick={() => void exportSettings()}>
          <Download className="size-3.5" /> 匯出設定(JSON)
        </Button>
        <Button
          variant="outline"
          size="sm"
          onClick={() => fileInputRef.current?.click()}
        >
          <Upload className="size-3.5" /> 匯入設定
        </Button>
        <input
          ref={fileInputRef}
          type="file"
          accept="application/json,.json"
          className="hidden"
          onChange={(e) => void onFilePicked(e)}
        />
      </div>
      <p className="text-[10px] text-muted-foreground">
        匯出/匯入所有設定(不含 API key)。可用來在新裝置上還原配置、
        或在團隊內分享通用設定範本。
      </p>
      {ok && <p className="text-[10px] text-emerald-700">{ok}</p>}
      {error && <p className="text-[10px] text-red-700">{error}</p>}
    </div>
  )
}

// ---- Cross-machine sync card (2026-05-22) ---------------------------------
//
// Uses chrome.storage.sync, which is browser-account-bound:
//   - Edge → Microsoft account
//   - Chrome → Google account
//   - Cross-browser doesn't auto-sync (use clipboard P3 instead)
//
// Sync scope (intentional):
//   ✓ Rules where source ∈ {user_manual, ai_confirmed, ai_overridden}
//   ✓ Recent tombstones (capped to fit 100 KB quota)
//   ✓ Settings minus claudeApiKey
//   ✗ auto_scan rules (re-derived per machine)
//   ✗ folderCache / folderActivity / skipHistory / memories (per-machine)
//   ✗ claudeApiKey (security)
function CrossMachineSyncCard({
  status,
  onChanged,
}: {
  status: StatusData | null
  onChanged: () => Promise<void> | void
}) {
  type SyncStatus = {
    enabled: boolean
    lastSyncAt: string
    machineId: string
    cloud?: {
      sourceMachineId: string
      updatedAt: string
      ruleCount: number
      tombstoneCount: number
      isUs: boolean
      recentPushes: Array<{ machineId: string; at: string }>
    }
    bytesInUse: number
    bytesQuota: number
    lastError?: {
      at: string
      source: 'push' | 'pull' | 'pull-remote'
      reason: string
    }
  }
  type BackupSummary = {
    snapshotAt: string
    direction: 'pre-push' | 'pre-pull' | 'manual'
    ruleCount: number
    tombstoneCount: number
  }

  const [syncStatus, setSyncStatus] = useState<SyncStatus | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [ok, setOk] = useState<string | null>(null)
  const [historyOpen, setHistoryOpen] = useState(false)
  const [backups, setBackups] = useState<BackupSummary[]>([])
  const [showEnableModal, setShowEnableModal] = useState(false)
  // Remote-wipe notice: set when ANOTHER machine ran "全部刪除規則"
  // and the wipeMarker rode in on our next pull. This card surfaces
  // the event so the user understands why their local syncable rules
  // disappeared — they can roll back from backup if it wasn't
  // intended.
  type RemoteWipeNotice = {
    byMachineId: string
    at: string
    appliedAt: string
    droppedRuleCount: number
  }
  const [remoteWipeNotice, setRemoteWipeNotice] = useState<RemoteWipeNotice | null>(null)

  const refreshSync = useCallback(async () => {
    const [r, rw] = await Promise.all([
      send<SyncStatus>({ type: 'getSyncStatus' }),
      send<{ notice: RemoteWipeNotice | null }>({ type: 'getRemoteWipeNotice' }),
    ])
    if (r.ok && r.data) setSyncStatus(r.data)
    if (rw.ok && rw.data) setRemoteWipeNotice(rw.data.notice)
  }, [])

  const refreshBackups = useCallback(async () => {
    const r = await send<{ backups: BackupSummary[] }>({ type: 'listSyncBackups' })
    if (r.ok && r.data) setBackups(r.data.backups)
  }, [])

  useEffect(() => {
    void refreshSync()
  }, [refreshSync, status])

  useEffect(() => {
    if (historyOpen) void refreshBackups()
  }, [historyOpen, refreshBackups])

  const browser = detectBrowser()
  const accountDesc = syncAccountDescription(browser)
  const settingsUrl = syncSettingsUrl(browser)

  async function enable() {
    setBusy(true)
    setError(null)
    setOk(null)
    const r = await send<{
      action: 'push' | 'pull' | 'union-merge'
      pushed?: boolean
      pulled?: boolean
      reason?: string
      ruleCount?: number
    }>({
      type: 'enableSync',
    })
    setBusy(false)
    if (!r.ok) {
      setError(r.message || r.code)
      return
    }
    let msg = '已上傳本機規則到雲端'
    if (r.data?.action === 'union-merge') {
      msg = `已合併本機與雲端規則 (雲端 ${r.data.ruleCount ?? 0} 條 + 本機獨有 → 合併後雙向同步)`
    } else if (r.data?.action === 'pull') {
      msg = '已套用雲端規則'
    }
    setOk(msg)
    await refreshSync()
    await onChanged()
    setShowEnableModal(false)
    setTimeout(() => setOk(null), 4000)
  }

  async function disableKeepCloud() {
    setBusy(true)
    setError(null)
    const r = await send({ type: 'disableSync', keepCloud: true })
    setBusy(false)
    if (!r.ok) setError(r.message || r.code)
    else setOk('已停用同步(雲端資料保留)')
    await refreshSync()
    await onChanged()
    setTimeout(() => setOk(null), 3000)
  }

  async function disableAndWipe() {
    if (!window.confirm('確定停用同步並清除雲端資料?\n本機資料不受影響、可隨時重新啟用。')) return
    setBusy(true)
    setError(null)
    const r = await send({ type: 'disableSync', keepCloud: false })
    setBusy(false)
    if (!r.ok) setError(r.message || r.code)
    else setOk('已停用同步並清除雲端資料')
    await refreshSync()
    await onChanged()
    setTimeout(() => setOk(null), 3000)
  }

  async function manualPush() {
    setBusy(true)
    setError(null)
    const r = await send<{
      pushed: boolean
      reason?: string
      truncatedRuleCount?: number
      truncatedTombstoneCount?: number
    }>({ type: 'pushSyncNow' })
    setBusy(false)
    if (!r.ok) {
      setError(r.message || r.code)
    } else if (r.data?.pushed) {
      const truncRules = r.data.truncatedRuleCount ?? 0
      const truncTombs = r.data.truncatedTombstoneCount ?? 0
      if (truncRules > 0 || truncTombs > 0) {
        // Truncation is a sync degradation worth flagging loudly, not
        // a transient OK message — keep on screen until user reads.
        setError(
          `已上傳、但有 ${truncRules} 條規則` +
            (truncTombs > 0 ? ` + ${truncTombs} 條墓碑` : '') +
            ' 因為超過 chunk 上限沒同步上去。請清理「自動休眠」規則。',
        )
      } else {
        setOk('已上傳到雲端')
        setTimeout(() => setOk(null), 3000)
      }
    } else {
      setError(`推送失敗:${r.data?.reason ?? '未知原因'}`)
    }
    await refreshSync()
  }

  async function manualPull() {
    setBusy(true)
    setError(null)
    const r = await send<{ pulled: boolean; reason?: string; ruleCount?: number; tombstoneCount?: number }>({
      type: 'pullSyncNow',
    })
    setBusy(false)
    if (!r.ok) setError(r.message || r.code)
    else if (r.data?.pulled) setOk(`已拉下雲端規則(${r.data.ruleCount ?? 0} 條)`)
    else setError(`拉取失敗:${r.data?.reason ?? '未知原因'}`)
    await refreshSync()
    await onChanged()
    setTimeout(() => setOk(null), 3000)
  }

  async function rollback(snapshotAt: string) {
    if (!window.confirm('回復到這個備份?目前的本機規則會被覆蓋(新增一個 manual 備份、可再回復)')) return
    setBusy(true)
    setError(null)
    const r = await send<{ restored: boolean }>({
      type: 'restoreSyncBackup',
      snapshotAt,
    })
    setBusy(false)
    if (!r.ok) setError(r.message || r.code)
    else if (r.data?.restored) setOk('已回復')
    else setError('找不到此備份')
    await refreshSync()
    await refreshBackups()
    await onChanged()
    setTimeout(() => setOk(null), 3000)
  }

  const enabled = syncStatus?.enabled ?? false
  const quotaPct =
    syncStatus && syncStatus.bytesQuota > 0
      ? Math.round((syncStatus.bytesInUse / syncStatus.bytesQuota) * 100)
      : 0

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">跨機器同步</CardTitle>
        <CardDescription className="space-y-1">
          <span className="block">
            把規則、墓碑、設定(扣 API key){accountDesc}同步到你另一台裝同款瀏覽器的機器。
          </span>
          <span className="block text-[10px] text-muted-foreground">
            ✓ 同步:user_manual / ai_confirmed / ai_overridden 規則、墓碑、設定 ·
            ✗ 不同步:auto_scan 規則(各機器各自跑初始掃描)、API key、本機快取
          </span>
          <span className="block text-[10px] text-muted-foreground">
            ⚠ Edge ↔ Chrome 之間不會自動同步、跨瀏覽器請用規則庫「操作 → 匯出規則 / 匯入規則」的 JSON 檔
          </span>
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-3 text-xs">
        {!enabled ? (
          <>
            <div className="rounded-md border border-border bg-muted/30 p-3 space-y-1">
              <div className="text-[11px]">目前未啟用同步。</div>
              <div className="text-[10px] text-muted-foreground">
                啟用前請確認兩台機器都登入相同{accountDesc.replace('透過你登入 ', '').replace(' 的', '')}、
                且開啟瀏覽器同步 + 擴充功能同步開關。
                {settingsUrl && (
                  <>
                    {' · '}
                    <button
                      type="button"
                      className="underline underline-offset-2 hover:text-foreground"
                      onClick={() => {
                        // chrome:// / edge:// URLs can't be opened via window.open() —
                        // copy to clipboard instead.
                        void navigator.clipboard.writeText(settingsUrl)
                        setOk(`已複製 ${settingsUrl} — 貼到網址列開啟`)
                        setTimeout(() => setOk(null), 3000)
                      }}
                    >
                      複製設定頁網址
                    </button>
                  </>
                )}
              </div>
            </div>
            <Button
              onClick={() => setShowEnableModal(true)}
              disabled={busy}
            >
              啟用同步
            </Button>
          </>
        ) : (
          <>
            {remoteWipeNotice && (
              <div className="rounded-md border border-amber-400 bg-amber-50 p-3 space-y-1.5">
                <div className="flex items-start gap-2 text-[11px]">
                  <span className="font-medium text-amber-900">
                    ⓘ 另一台機器執行了「全部刪除」
                  </span>
                  <button
                    type="button"
                    className="ml-auto text-[10px] text-amber-900/70 hover:text-amber-950 underline underline-offset-2"
                    onClick={async () => {
                      await send({ type: 'dismissRemoteWipeNotice' })
                      setRemoteWipeNotice(null)
                    }}
                  >
                    我知道了
                  </button>
                </div>
                <div className="text-[11px] text-amber-950">
                  本機已自動同步清除 {remoteWipeNotice.droppedRuleCount} 條
                  跨機器規則 + 全部墓碑(發生於{' '}
                  {new Date(remoteWipeNotice.at).toLocaleString('zh-TW')})。
                </div>
                <div className="text-[10px] text-amber-900/80">
                  若不是本人在另一機器清除,請開「同步歷史」回滾至清除前的快照。
                </div>
              </div>
            )}
            {syncStatus?.lastError &&
              (() => {
                // #7: classify the error reason and produce an action
                // hint + optional CTA. The raw reason ('unsupported
                // cloud schemaVersion=2 ...') is technical noise to a
                // lawyer — actionable text + a button gets them out of
                // the broken state faster.
                const action = describeSyncError(
                  syncStatus.lastError.reason,
                  syncStatus.lastError.source,
                )
                return (
                  <div className="rounded-md border border-red-300 bg-red-50 p-3 space-y-1.5">
                    <div className="flex items-start gap-2 text-[11px]">
                      <span className="font-medium text-red-700">
                        ⚠ 上次同步失敗
                      </span>
                      <span className="text-red-700/80 text-[10px]">
                        {syncStatus.lastError.source === 'push'
                          ? '(上傳)'
                          : syncStatus.lastError.source === 'pull'
                            ? '(下載)'
                            : '(來自另一台機器的變更未能拉下)'}
                      </span>
                      <button
                        type="button"
                        className="ml-auto text-[10px] text-red-700/70 hover:text-red-900 underline underline-offset-2"
                        onClick={async () => {
                          await send({ type: 'dismissSyncError' })
                          await refreshSync()
                        }}
                      >
                        清除提示
                      </button>
                    </div>
                    <div className="text-[11px] text-red-900 font-medium">
                      {action.hint}
                    </div>
                    {action.cta && (
                      <div>
                        <button
                          type="button"
                          onClick={() => {
                            if (action.cta!.action === 'retry') {
                              void (async () => {
                                await send({
                                  type:
                                    syncStatus.lastError!.source === 'push'
                                      ? 'pushSyncNow'
                                      : 'pullSyncNow',
                                })
                                await refreshSync()
                              })()
                            } else if (action.cta!.action === 'upgrade') {
                              // chrome:// can't be opened from extension
                              // pages; copy to clipboard.
                              void navigator.clipboard.writeText(
                                'chrome://extensions',
                              )
                              setOk('已複製 chrome://extensions — 貼到網址列檢查擴充版本')
                              setTimeout(() => setOk(null), 4000)
                            } else if (action.cta!.action === 'cleanup') {
                              // Scroll to rules section so user can clean
                              // up sleeping / stale rules.
                              const el =
                                document.getElementById('rules')
                              el?.scrollIntoView({ behavior: 'smooth' })
                            }
                          }}
                          className="text-[10px] bg-red-700 hover:bg-red-800 text-white rounded px-2 py-1"
                        >
                          {action.cta.label}
                        </button>
                      </div>
                    )}
                    <details className="text-[10px] text-red-900/60">
                      <summary className="cursor-pointer">技術細節</summary>
                      <div className="break-all mt-1">
                        {syncStatus.lastError.reason}
                      </div>
                      <div className="mt-0.5">
                        發生時間：
                        {new Date(syncStatus.lastError.at).toLocaleString('zh-TW')}
                      </div>
                    </details>
                  </div>
                )
              })()}
            <div className="rounded-md border border-emerald-200 bg-emerald-50 p-3 space-y-1.5">
              <div className="flex items-center gap-2 text-[11px]">
                <CheckCircle2 className="size-3.5 text-emerald-700" />
                <span className="font-medium">同步已啟用</span>
                <span className="text-muted-foreground">·</span>
                <span className="text-muted-foreground">本機 ID:{syncStatus?.machineId.slice(0, 8)}…</span>
              </div>
              <div className="text-[10px] text-muted-foreground space-y-0.5">
                <div>上次同步:{syncStatus?.lastSyncAt ? new Date(syncStatus.lastSyncAt).toLocaleString('zh-TW') : '從未'}</div>
                {syncStatus?.cloud && (
                  <div>
                    雲端狀態:{syncStatus.cloud.ruleCount} 條規則 · {syncStatus.cloud.tombstoneCount} 條墓碑 ·
                    {syncStatus.cloud.isUs ? '本機' : '另一台機器'}寫入於{' '}
                    {new Date(syncStatus.cloud.updatedAt).toLocaleString('zh-TW')}
                  </div>
                )}
                <div>
                  Quota 使用:{Math.round((syncStatus?.bytesInUse ?? 0) / 1024)} KB /{' '}
                  {Math.round((syncStatus?.bytesQuota ?? 102400) / 1024)} KB
                  {quotaPct > 0 && ` (${quotaPct}%)`}
                  {quotaPct >= 80 && <span className="ml-1 text-amber-700">⚠ 接近上限</span>}
                </div>
              </div>
              {/* #6: multi-machine visibility — show each machine's last
                  push, newest first. The current machine is highlighted
                  ("本機") so the user can spot themselves in the list. */}
              {syncStatus?.cloud && syncStatus.cloud.recentPushes.length > 0 && (
                <div className="pt-1 border-t border-emerald-200/60">
                  <div className="text-[10px] text-muted-foreground mb-0.5">
                    最近活動 ({syncStatus.cloud.recentPushes.length} 台機器)
                  </div>
                  <ul className="space-y-0.5">
                    {syncStatus.cloud.recentPushes.map((p) => {
                      const isMe = p.machineId === syncStatus.machineId
                      return (
                        <li
                          key={p.machineId}
                          className="flex items-center gap-2 text-[10px] font-mono"
                        >
                          <span
                            className={
                              isMe
                                ? 'inline-flex items-center rounded-md bg-seal/20 text-seal-dim px-1 py-px font-medium'
                                : 'text-muted-foreground'
                            }
                          >
                            {isMe ? '本機' : p.machineId.slice(0, 8) + '…'}
                          </span>
                          <span className="text-muted-foreground tabular-nums ml-auto">
                            {formatRelativeTime(p.at)}
                          </span>
                        </li>
                      )
                    })}
                  </ul>
                </div>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <Button size="sm" variant="outline" onClick={() => void manualPush()} disabled={busy}>
                立即上傳
              </Button>
              <Button size="sm" variant="outline" onClick={() => void manualPull()} disabled={busy}>
                立即下載
              </Button>
              <button
                type="button"
                onClick={() => setHistoryOpen((v) => !v)}
                className="text-[10px] underline underline-offset-2 hover:text-foreground"
              >
                {historyOpen ? '收起' : '展開'}同步歷史 / 回復
              </button>
              <Button
                size="sm"
                variant="ghost"
                className="ml-auto text-muted-foreground"
                onClick={() => void disableKeepCloud()}
                disabled={busy}
              >
                停用同步(保留雲端)
              </Button>
              <Button
                size="sm"
                variant="ghost"
                className="text-red-700"
                onClick={() => void disableAndWipe()}
                disabled={busy}
              >
                停用並清除雲端
              </Button>
            </div>

            {historyOpen && (
              <div className="rounded-md border border-border bg-muted/30 p-2 space-y-1">
                <div className="text-[10px] font-medium">本機同步快照(最近 5 筆)</div>
                {backups.length === 0 ? (
                  <div className="text-[10px] text-muted-foreground py-1">尚無快照</div>
                ) : (
                  <ul className="space-y-1">
                    {backups.map((b) => {
                      const dirLabel =
                        b.direction === 'pre-push'
                          ? '上傳前'
                          : b.direction === 'pre-pull'
                            ? '下載前'
                            : '手動'
                      return (
                        <li
                          key={b.snapshotAt}
                          className="flex items-center gap-2 text-[10px] font-mono"
                        >
                          <span className="text-muted-foreground">{dirLabel}</span>
                          <span>{new Date(b.snapshotAt).toLocaleString('zh-TW')}</span>
                          <span className="text-muted-foreground tabular-nums">
                            規則 {b.ruleCount} / 墓碑 {b.tombstoneCount}
                          </span>
                          <button
                            type="button"
                            className="ml-auto text-[10px] underline underline-offset-2 hover:text-foreground disabled:opacity-50"
                            disabled={busy}
                            onClick={() => void rollback(b.snapshotAt)}
                          >
                            回復到此
                          </button>
                        </li>
                      )
                    })}
                  </ul>
                )}
              </div>
            )}
          </>
        )}

        {ok && <p className="text-[10px] text-emerald-700">{ok}</p>}
        {error && <p className="text-[10px] text-red-700">{error}</p>}
      </CardContent>

      {showEnableModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 p-4">
          <div className="bg-card rounded-md border border-border p-4 max-w-md space-y-3 text-xs">
            <h3 className="text-sm font-medium">啟用跨機器同步</h3>
            <div className="space-y-1 text-[11px]">
              <div>啟用前請確認:</div>
              <ul className="list-disc list-inside space-y-0.5 text-muted-foreground pl-1">
                <li>兩台機器都登入相同的{browserLabel(browser)}帳號</li>
                <li>{browserLabel(browser)} 同步開啟、且「擴充功能」同步勾選</li>
                <li>第一次同步:若雲端已有資料(另一台先啟用過)、會拉下來;若沒有、會以「這台」為主上傳</li>
                <li>每次推 / 拉前都會自動本機備份、可從「同步歷史」回復</li>
              </ul>
            </div>
            {/* Surface the failure INSIDE the modal — the card-level error
                banner renders underneath this fixed z-50 overlay, so a
                failed enable used to look like the button silently did
                nothing. */}
            {error && (
              <p className="text-[11px] text-red-700" role="alert">
                啟用失敗:{error}
              </p>
            )}
            <div className="flex justify-end gap-2 pt-1">
              <Button size="sm" variant="ghost" onClick={() => setShowEnableModal(false)} disabled={busy}>
                取消
              </Button>
              <Button size="sm" onClick={() => void enable()} disabled={busy}>
                {busy ? '啟用中…' : '我確認、啟用'}
              </Button>
            </div>
          </div>
        </div>
      )}
    </Card>
  )
}

/**
 * Centralised error log viewer. Reads from chrome.storage.local['errorLog']
 * (populated by `logError()` calls scattered across the SW) and renders
 * a scrollable list with source-prefix grouping + clear button.
 *
 * Collapsed by default — most days the log is empty and the user doesn't
 * need to look. The header chip surfaces the count so they know when to
 * expand.
 */
function ErrorLogCard() {
  type ErrorEntry = {
    at: string
    source: string
    message: string
    context?: Record<string, unknown>
  }
  const [entries, setEntries] = useState<ErrorEntry[]>([])
  const [open, setOpen] = useState(false)
  const [loading, setLoading] = useState(false)

  const refresh = useCallback(async () => {
    setLoading(true)
    const r = await send<{ entries: ErrorEntry[] }>({
      type: 'getErrorLog',
      limit: 100,
    })
    if (r.ok && r.data) setEntries(r.data.entries)
    setLoading(false)
  }, [])

  useEffect(() => {
    void refresh()
  }, [refresh])

  async function handleClear() {
    if (!window.confirm('清空錯誤紀錄?')) return
    await send({ type: 'clearErrorLog' })
    await refresh()
  }

  const count = entries.length
  // Group by source-prefix (before colon) so the user sees the failure
  // domains at a glance: "sync: 3, classify: 1, outlook: 2".
  const summary = useMemo(() => {
    const by = new Map<string, number>()
    for (const e of entries) {
      const prefix = e.source.split(':')[0] ?? e.source
      by.set(prefix, (by.get(prefix) ?? 0) + 1)
    }
    return [...by.entries()].sort((a, b) => b[1] - a[1])
  }, [entries])

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm flex items-center gap-2">
          <AlertTriangle className="size-4" />
          錯誤紀錄
          {count > 0 && (
            <span className="ml-1 inline-flex items-center rounded-full bg-amber-100 px-1.5 py-0.5 text-[10px] font-medium text-amber-800">
              {count}
            </span>
          )}
        </CardTitle>
        <CardDescription className="text-[11px]">
          背景服務裡的失敗訊息會收在這裡。每天看一眼、不正常的話可發現。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-2">
        {count === 0 ? (
          <p className="text-[11px] text-muted-foreground italic">
            {loading ? '載入中…' : '沒有錯誤紀錄 ✓'}
          </p>
        ) : (
          <>
            <div className="flex items-center gap-2 text-[11px]">
              <span className="text-muted-foreground">分類：</span>
              {summary.map(([prefix, n]) => (
                <span
                  key={prefix}
                  className="inline-flex items-center rounded-md border border-border bg-muted/30 px-1.5 py-0.5 font-mono text-[10px]"
                >
                  {prefix} · {n}
                </span>
              ))}
              <button
                type="button"
                onClick={() => setOpen((v) => !v)}
                className="ml-auto text-[10px] underline underline-offset-2 hover:text-foreground"
              >
                {open ? '收起' : '展開'}
              </button>
              <button
                type="button"
                onClick={() => void handleClear()}
                className="text-[10px] text-red-700 underline underline-offset-2 hover:text-red-900"
              >
                清空
              </button>
            </div>
            {open && (
              <div className="rounded-md border border-border bg-muted/20 max-h-72 overflow-y-auto">
                <ul className="divide-y divide-border">
                  {[...entries].reverse().map((e, i) => (
                    <li key={`${e.at}-${i}`} className="px-2 py-1.5">
                      <div className="flex items-baseline gap-2 text-[10px] font-mono">
                        <span className="text-muted-foreground tabular-nums">
                          {new Date(e.at).toLocaleString('zh-TW')}
                        </span>
                        <span className="font-medium text-amber-700">
                          {e.source}
                        </span>
                      </div>
                      <div className="text-[11px] mt-0.5 break-words">
                        {e.message}
                      </div>
                      {e.context && Object.keys(e.context).length > 0 && (
                        <div className="text-[10px] mt-0.5 font-mono text-muted-foreground break-all">
                          {JSON.stringify(e.context)}
                        </div>
                      )}
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </>
        )}
      </CardContent>
    </Card>
  )
}

/**
 * 近日活動顯示範圍 — picker-driven redesign (2026-05-27).
 *
 * Previous UX was two textareas where users had to type folder paths
 * with the right semantics (`案件資料夾/` with trailing slash for
 * descendants, leaf name without slash) — error-prone and required
 * domain knowledge of the matcher's internals.
 *
 * New UX:
 *   - Chip-list shows currently-configured entries (folder vs leaf
 *     name tagged separately).
 *   - Two `+ 新增` buttons open a FolderPicker overlay; user selects
 *     a folder from their actual tree. We store path-without-slash
 *     for folder mode (matcher handles "folder itself + descendants"
 *     via the new semantics in folder-activity-filter.ts).
 *   - Save is immediate on add/remove — no separate "儲存" step.
 *   - Free-text fallback exists behind 「進階」 disclosure for power
 *     users / weird match cases (legacy `案件資料夾/` form still
 *     accepted by the matcher).
 */
function RecentActivityFilterCard({
  status,
  tree,
  onSaved,
}: {
  status: StatusData | null
  tree: MailFolderNode[] | null
  onSaved: () => Promise<void> | void
}) {
  const prefixes = status?.recentActivityIncludePrefixes ?? []
  const leaves = status?.recentActivityIncludeLeafNames ?? []
  const totalCount = prefixes.length + leaves.length

  const [adding, setAdding] = useState<null | 'prefix' | 'leaf'>(null)
  const [savedAt, setSavedAt] = useState<number | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [advancedOpen, setAdvancedOpen] = useState(false)
  const [saving, setSaving] = useState(false)

  async function commit(nextPrefixes: string[], nextLeaves: string[]) {
    // In-flight guard (audit: stale-read race). `prefixes`/`leaves` come from
    // the `status` prop, which only refreshes after send()+onSaved() finish —
    // two async round-trips. A second add/remove fired before that refresh
    // would rebuild its next-array from the STALE snapshot and overwrite the
    // first write, resurrecting a just-removed pill (or dropping a just-added
    // one). Block re-entry here, and disable the pill controls + picker while
    // saving, so each mutation sees fresh state.
    if (saving) return
    setSaving(true)
    setSaveError(null)
    try {
      const r = await send({
        type: 'setSettings',
        patch: {
          recentActivityIncludePrefixes: Array.from(new Set(nextPrefixes)),
          recentActivityIncludeLeafNames: Array.from(new Set(nextLeaves)),
        },
      })
      if (r.ok) {
        setSavedAt(Date.now())
        await onSaved()
        window.setTimeout(() => setSavedAt(null), 2000)
      } else {
        setSaveError(r.message || r.code || '儲存失敗')
      }
    } finally {
      setSaving(false)
    }
  }

  async function addPrefix(node: MailFolderNode) {
    if (prefixes.includes(node.path)) {
      setAdding(null)
      return
    }
    await commit([...prefixes, node.path], leaves)
    setAdding(null)
  }
  async function addLeaf(node: MailFolderNode) {
    if (leaves.includes(node.displayName)) {
      setAdding(null)
      return
    }
    await commit(prefixes, [...leaves, node.displayName])
    setAdding(null)
  }
  async function removePrefix(value: string) {
    await commit(prefixes.filter((p) => p !== value), leaves)
  }
  async function removeLeaf(value: string) {
    await commit(prefixes, leaves.filter((l) => l !== value))
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">近日活動 顯示範圍</CardTitle>
        <CardDescription>
          設定 popup「近日活動」面板要顯示哪些資料夾的活動。
          <span className="font-medium">全部留空 = 顯示所有活動</span>(預設)。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-3 text-xs">
        {totalCount === 0 ? (
          <div className="rounded-md border border-dashed border-border p-4 text-center text-muted-foreground">
            <p>目前顯示所有資料夾的活動</p>
            <p className="text-[11px] mt-1">下方加入篩選條件以限縮顯示範圍</p>
          </div>
        ) : (
          <ul className="space-y-1">
            {prefixes.map((p) => (
              <li
                key={`p-${p}`}
                className="flex items-center gap-2 p-1.5 rounded border border-border bg-muted/30"
              >
                <Badge
                  variant="outline"
                  className="text-[9px] font-mono shrink-0"
                  title="包含此資料夾及其所有子資料夾"
                >
                  資料夾
                </Badge>
                <code className="font-mono text-xs flex-1 truncate" title={p}>
                  {p}
                </code>
                <button
                  type="button"
                  onClick={() => void removePrefix(p)}
                  disabled={saving}
                  className="text-muted-foreground hover:text-destructive shrink-0 disabled:opacity-40 disabled:pointer-events-none"
                  aria-label={`移除 ${p}`}
                >
                  <X className="size-3.5" />
                </button>
              </li>
            ))}
            {leaves.map((l) => (
              <li
                key={`l-${l}`}
                className="flex items-center gap-2 p-1.5 rounded border border-border bg-muted/30"
              >
                <Badge
                  variant="outline"
                  className="text-[9px] font-mono shrink-0"
                  title="不論在哪一層,只要最末層資料夾名稱符合就顯示"
                >
                  葉節點
                </Badge>
                <code className="font-mono text-xs flex-1 truncate" title={l}>
                  {l}
                </code>
                <button
                  type="button"
                  onClick={() => void removeLeaf(l)}
                  disabled={saving}
                  className="text-muted-foreground hover:text-destructive shrink-0 disabled:opacity-40 disabled:pointer-events-none"
                  aria-label={`移除 ${l}`}
                >
                  <X className="size-3.5" />
                </button>
              </li>
            ))}
          </ul>
        )}

        {adding === 'prefix' ? (
          <div className="space-y-2 p-3 rounded-md border border-foreground/40 bg-muted/20">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-medium">
                選擇資料夾(會包含此資料夾及其所有子資料夾)
              </span>
              <button
                type="button"
                onClick={() => setAdding(null)}
                className="text-muted-foreground hover:text-foreground"
                aria-label="取消"
              >
                <X className="size-3.5" />
              </button>
            </div>
            {tree && tree.length > 0 ? (
              <FolderPicker
                tree={tree}
                onSelect={(node) => {
                  if (!saving) void addPrefix(node)
                }}
                placeholder="搜尋或點選資料夾…"
              />
            ) : (
              <p className="text-[11px] text-muted-foreground">
                資料夾樹尚未載入 — 先到設定→分類引擎→資料夾快取按重新偵測
              </p>
            )}
          </div>
        ) : adding === 'leaf' ? (
          <div className="space-y-2 p-3 rounded-md border border-foreground/40 bg-muted/20">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-medium">
                選擇任一資料夾,系統會以該資料夾的「名稱」當匹配關鍵字
              </span>
              <button
                type="button"
                onClick={() => setAdding(null)}
                className="text-muted-foreground hover:text-foreground"
                aria-label="取消"
              >
                <X className="size-3.5" />
              </button>
            </div>
            <p className="text-[10px] text-muted-foreground">
              例:選了<code className="font-mono">內部資料/待簽核</code>會儲存<code className="font-mono">待簽核</code>,
              之後任何路徑只要最末層叫「待簽核」都會被列入(跨類別找同名資料夾用)。
            </p>
            {tree && tree.length > 0 ? (
              <FolderPicker
                tree={tree}
                onSelect={(node) => {
                  if (!saving) void addLeaf(node)
                }}
                placeholder="搜尋或點選資料夾…"
              />
            ) : (
              <p className="text-[11px] text-muted-foreground">
                資料夾樹尚未載入 — 先到設定→分類引擎→資料夾快取按重新偵測
              </p>
            )}
          </div>
        ) : (
          <div className="flex flex-wrap items-center gap-2">
            <Button
              size="sm"
              variant="outline"
              onClick={() => setAdding('prefix')}
              disabled={!tree || tree.length === 0 || saving}
            >
              <Plus className="size-3.5" />
              新增資料夾(含子資料夾)
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => setAdding('leaf')}
              disabled={!tree || tree.length === 0 || saving}
            >
              <Plus className="size-3.5" />
              新增葉節點名稱
            </Button>
            {!tree && (
              <span className="text-[10px] text-muted-foreground">
                等待資料夾樹載入…
              </span>
            )}
          </div>
        )}

        {savedAt && (
          <p className="text-[10px] text-emerald-700">已儲存</p>
        )}
        {saveError && (
          <p className="text-[10px] text-red-700">{saveError}</p>
        )}

        {/* Power-user free-text fallback. Hidden by default since the
            picker covers all normal cases. Useful when the user wants
            to anticipate a folder that doesn't exist yet, or to enter
            the legacy `案件資料夾/` form explicitly for descendant-only
            matching. */}
        <details
          open={advancedOpen}
          onToggle={(e) => setAdvancedOpen((e.target as HTMLDetailsElement).open)}
          className="border-t border-border pt-2"
        >
          <summary className="cursor-pointer text-[11px] text-muted-foreground hover:text-foreground select-none">
            進階:手動編輯(每行一條)
          </summary>
          <AdvancedManualEditor
            prefixes={prefixes}
            leaves={leaves}
            onCommit={commit}
            parentSaving={saving}
          />
        </details>
      </CardContent>
    </Card>
  )
}

/**
 * Advanced fallback editor — preserves the old "two textareas" UX as
 * an escape hatch for power users. Kept inside RecentActivityFilterCard
 * so it can't accidentally be reused elsewhere with the wrong intent.
 */
function AdvancedManualEditor({
  prefixes,
  leaves,
  onCommit,
  parentSaving,
}: {
  prefixes: string[]
  leaves: string[]
  onCommit: (nextPrefixes: string[], nextLeaves: string[]) => Promise<void>
  /** Parent card's in-flight commit flag. commit() drops re-entrant calls
   *  silently (the stale-read-race guard), so this button must also be
   *  disabled while a pill commit is in flight — otherwise a click here
   *  would no-op with a green-looking spinner flash and no error. */
  parentSaving: boolean
}) {
  const [prefixDraft, setPrefixDraft] = useState(prefixes.join('\n'))
  const [leafDraft, setLeafDraft] = useState(leaves.join('\n'))
  const [saving, setSaving] = useState(false)

  // Re-seed drafts when parent state changes from the picker UI above.
  useEffect(() => {
    setPrefixDraft(prefixes.join('\n'))
  }, [prefixes])
  useEffect(() => {
    setLeafDraft(leaves.join('\n'))
  }, [leaves])

  const parseList = (s: string): string[] =>
    Array.from(
      new Set(s.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 0)),
    )

  const dirty =
    parseList(prefixDraft).join('\n') !== prefixes.join('\n') ||
    parseList(leafDraft).join('\n') !== leaves.join('\n')

  async function save() {
    setSaving(true)
    try {
      await onCommit(parseList(prefixDraft), parseList(leafDraft))
    } finally {
      setSaving(false)
    }
  }

  return (
    <div className="mt-2 space-y-2">
      <div className="grid gap-2 md:grid-cols-2">
        <div className="space-y-1">
          <label className="text-[10px] text-muted-foreground">
            路徑(每行一條)
          </label>
          <textarea
            value={prefixDraft}
            onChange={(e) => setPrefixDraft(e.target.value)}
            rows={3}
            className="w-full rounded-md border border-border bg-background px-2 py-1.5 text-[11px] font-mono"
            placeholder={'案件資料夾\n研究專案/'}
          />
          <p className="text-[10px] text-muted-foreground">
            無斜線=資料夾+子資料夾。 結尾斜線=只比對子資料夾(舊版相容)。
          </p>
        </div>
        <div className="space-y-1">
          <label className="text-[10px] text-muted-foreground">
            葉節點名稱(每行一條)
          </label>
          <textarea
            value={leafDraft}
            onChange={(e) => setLeafDraft(e.target.value)}
            rows={3}
            className="w-full rounded-md border border-border bg-background px-2 py-1.5 text-[11px] font-mono"
            placeholder={'待簽核\n工時審閱'}
          />
        </div>
      </div>
      <Button onClick={() => void save()} disabled={!dirty || saving || parentSaving} size="sm">
        {saving ? '儲存中…' : '儲存手動編輯'}
      </Button>
    </div>
  )
}

// ---- Rule health section -------------------------------------------------

/**
 * "Run stale-sweep now" row. Background already runs daily, but the
 * button lets the user trigger immediately + see how many got
 * auto-disabled. Useful right after adopting subject-feature learning
 * to clean up old rules that never matched.
 */
function StaleSweepActionRow() {
  const [running, setRunning] = useState(false)
  const [result, setResult] = useState<string | null>(null)

  async function run() {
    setRunning(true)
    setResult(null)
    const r = await send<{ disabledCount: number; staleDeletedCount?: number }>({
      type: 'runStaleSweep',
    })
    setRunning(false)
    if (r.ok && r.data) {
      const disabled = r.data.disabledCount
      const deleted = r.data.staleDeletedCount ?? 0
      // Report BOTH outcomes — stale rules are hard-deleted (not disabled),
      // so a sweep that only deleted used to read as "nothing to do".
      const parts: string[] = []
      if (deleted > 0) parts.push(`已刪除 ${deleted} 條休眠規則`)
      if (disabled > 0) parts.push(`已自動停用 ${disabled} 條規則`)
      setResult(parts.length > 0 ? parts.join('、') : '沒有符合條件的規則需要清理')
      window.setTimeout(() => setResult(null), 4000)
    } else if (!r.ok) {
      setResult(`失敗:${r.message ?? r.code}`)
    }
  }

  return (
    <div className="flex items-center gap-2 text-[10px] text-muted-foreground border-b border-border pb-2 mb-1">
      <button
        type="button"
        onClick={() => void run()}
        disabled={running}
        className="underline underline-offset-2 hover:text-foreground disabled:opacity-50"
        title="現在執行一次休眠規則掃除(平時每天背景自動跑一次)"
      >
        {running ? '掃除中…' : '立即執行休眠掃除'}
      </button>
      {result && <span className="text-emerald-700">{result}</span>}
    </div>
  )
}

function RuleHealthSection({
  rules,
  onToggle,
  onDelete,
  onEdit,
}: {
  rules: Rule[]
  onToggle: (rule: Rule, enabled: boolean) => Promise<boolean>
  onDelete: (ruleId: string) => Promise<boolean>
  /** Jump to Rules section and open the editor on this rule. */
  onEdit: (ruleId: string) => void
}) {
  const health = useMemo(() => computeRuleHealth(rules), [rules])
  const [openBucket, setOpenBucket] = useState<
    null | 'sleeping' | 'hotVague' | 'orphaned' | 'overBroad' | 'conflicts'
  >(null)

  const conflictRules = useMemo(
    () => rules.filter((r) => health.conflictRuleIds.has(r.id)),
    [rules, health.conflictRuleIds],
  )

  const totalSignals =
    health.counts.sleeping +
    health.counts.hotVague +
    health.counts.orphaned +
    health.counts.overBroad +
    health.counts.conflicts

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-sm">
          <AlertTriangle className="size-4" />
          規則健康度
          {totalSignals === 0 ? (
            <Badge variant="success" className="gap-1">
              <CheckCircle2 className="size-3" />
              全綠
            </Badge>
          ) : (
            <Badge variant="warning">{totalSignals} 個訊號</Badge>
          )}
        </CardTitle>
        <CardDescription>
          長期未命中、命中多但路徑模糊、彼此衝突、目標資料夾不存在等情況。點分類展開檢視與處理。
          每天背景自動把 100 天未命中的「休眠」規則刪除(不可復原、不留紀錄;同訊號日後可重新學回)、不影響手動規則。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-2 text-xs">
        <StaleSweepActionRow />
        <HealthBucketButton
          label="休眠"
          count={health.counts.sleeping}
          active={openBucket === 'sleeping'}
          onClick={() => setOpenBucket((b) => (b === 'sleeping' ? null : 'sleeping'))}
          description="從未命中 30 天 / 或上次命中 ≥ 90 天前。每日自動掃除會把 ≥ 100 天未命中的規則刪除(不可復原、不留紀錄;同訊號日後可重新學回)"
        />
        <HealthBucketButton
          label="模糊熱門"
          count={health.counts.hotVague}
          active={openBucket === 'hotVague'}
          onClick={() => setOpenBucket((b) => (b === 'hotVague' ? null : 'hotVague'))}
          description="高命中數但目標含「未分類 / 其他 / 雜項」等通用資料夾 — 建議細分"
        />
        <HealthBucketButton
          label="衝突"
          count={health.counts.conflicts}
          active={openBucket === 'conflicts'}
          onClick={() => setOpenBucket((b) => (b === 'conflicts' ? null : 'conflicts'))}
          description="同類型同條件對應不同資料夾 — 衝突時改由 AI 逐封判斷"
        />
        <HealthBucketButton
          label="目標遺失"
          count={health.counts.orphaned}
          active={openBucket === 'orphaned'}
          onClick={() => setOpenBucket((b) => (b === 'orphaned' ? null : 'orphaned'))}
          description="規則指定的目標資料夾在 Outlook 中已不存在 — 命中時會被跳過"
        />
        <HealthBucketButton
          label="過於廣泛"
          count={health.counts.overBroad}
          active={openBucket === 'overBroad'}
          onClick={() => setOpenBucket((b) => (b === 'overBroad' ? null : 'overBroad'))}
          description="純 domain 規則綁在 gmail.com / yahoo.com 等通用信箱 — 建議升級為複合或刪除"
        />

        {openBucket === 'sleeping' && (
          <HealthRuleList
            rules={health.sleeping}
            emptyText="沒有休眠規則"
            actions={(r) => (
              <>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => onEdit(r.id)}
                  title="在規則庫展開編輯"
                >
                  <Pencil className="size-3" /> 編輯
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => void onToggle(r, false)}
                >
                  停用
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 px-2 text-[10px] text-red-700 hover:text-red-700"
                  onClick={() => void onDelete(r.id)}
                >
                  刪除
                </Button>
              </>
            )}
          />
        )}
        {openBucket === 'hotVague' && (
          <HealthRuleList
            rules={health.hotVague}
            emptyText="沒有模糊熱門規則"
            actions={(r) => (
              <>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => onEdit(r.id)}
                  title="在規則庫展開編輯 — 可改 signal 或 target"
                >
                  <Pencil className="size-3" /> 編輯
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 px-2 text-[10px] text-red-700 hover:text-red-700"
                  onClick={() => void onDelete(r.id)}
                >
                  刪除
                </Button>
              </>
            )}
          />
        )}
        {openBucket === 'conflicts' && (
          <HealthRuleList
            rules={conflictRules}
            emptyText="沒有衝突"
            actions={(r) => (
              <>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => onEdit(r.id)}
                  title="在規則庫展開編輯"
                >
                  <Pencil className="size-3" /> 編輯
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 px-2 text-[10px] text-red-700 hover:text-red-700"
                  onClick={() => void onDelete(r.id)}
                >
                  刪除
                </Button>
              </>
            )}
          />
        )}
        {openBucket === 'orphaned' && (
          <HealthRuleList
            rules={health.orphaned}
            emptyText="沒有目標遺失"
            actions={(r) => (
              <>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => onEdit(r.id)}
                  title="在規則庫展開編輯 — 重新指定 target"
                >
                  <Pencil className="size-3" /> 重新指定 target
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 px-2 text-[10px] text-red-700 hover:text-red-700"
                  onClick={() => void onDelete(r.id)}
                >
                  刪除
                </Button>
              </>
            )}
          />
        )}
        {openBucket === 'overBroad' && (
          <HealthRuleList
            rules={health.overBroad}
            emptyText="沒有過於廣泛的 domain 規則"
            actions={(r) => (
              <>
                <Button
                  size="sm"
                  variant="outline"
                  className="h-6 px-2 text-[10px]"
                  onClick={() => onEdit(r.id)}
                  title="在規則庫展開編輯 — 可改成複合規則(domain + 主旨關鍵字)"
                >
                  <Pencil className="size-3" /> 編輯
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 px-2 text-[10px] text-red-700 hover:text-red-700"
                  onClick={() => void onDelete(r.id)}
                >
                  刪除
                </Button>
              </>
            )}
          />
        )}
      </CardContent>
    </Card>
  )
}

function HealthBucketButton({
  label,
  count,
  active,
  onClick,
  description,
}: {
  label: string
  count: number
  active: boolean
  onClick: () => void
  description: string
}) {
  const isClean = count === 0
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={isClean}
      className={cn(
        'w-full text-left rounded-md border px-2.5 py-1.5 transition-colors',
        isClean
          ? 'border-border bg-muted/30 text-muted-foreground cursor-default'
          : active
            ? 'border-foreground bg-accent'
            : 'border-border bg-card hover:bg-accent/50 cursor-pointer',
      )}
    >
      <div className="flex items-baseline justify-between gap-2">
        <span className="font-medium">{label}</span>
        <span
          className={cn(
            'font-mono tabular-nums',
            isClean ? 'text-muted-foreground' : count > 0 ? 'text-amber-800' : '',
          )}
        >
          {count}
        </span>
      </div>
      <div className="text-[10px] text-muted-foreground mt-0.5">{description}</div>
    </button>
  )
}

function HealthRuleList({
  rules,
  emptyText,
  actions,
}: {
  rules: Rule[]
  emptyText: string
  actions: (r: Rule) => ReactNode
}) {
  if (rules.length === 0) {
    return <p className="text-[11px] text-muted-foreground italic px-2">{emptyText}</p>
  }
  return (
    <ul className="space-y-1 max-h-72 overflow-y-auto rounded-md border border-border bg-muted/20 p-2">
      {rules.map((r) => (
        <li key={r.id} className="flex items-center gap-2 text-[11px]">
          <Badge variant="outline" className="font-mono text-[9px]">
            {refinedTypeLabel(r.type, r.signal)}
          </Badge>
          <span className="font-mono truncate min-w-0 flex-1" title={r.signal}>
            {r.type === 'compound' ? formatCompoundSignal(r.signal) : r.signal}
          </span>
          <span className="font-mono text-[10px] text-muted-foreground truncate">
            → {r.targetFolderPath}
          </span>
          <span className="font-mono text-[10px] tabular-nums text-muted-foreground shrink-0">
            {r.matchCount} 中
          </span>
          <span className="flex gap-1 shrink-0">{actions(r)}</span>
        </li>
      ))}
    </ul>
  )
}

// ---- Rule edit history section -------------------------------------------
//
// Audit log surface for everything that touched the rule library: user edits
// (toggle / delete / target change / split-into-compound) and system writes
// (initial-scan, ai_confirmed, ai_overridden, reconcile, import). Lets the
// user trace "when did this rule appear / why was X disabled" without
// needing to remember.
//
// Filter chips: 全部 / 我改的 / 系統 / 動作類型 — most-recent first, capped
// at 100 entries (storage already trims at 500).

const RULE_TYPE_LABEL_TRACE: Record<RuleType, string> = {
  domain: 'domain',
  sender: 'sender',
  case_code: 'case_code',
  subject_keyword: 'subject',
  compound: 'compound',
}

const KIND_LABEL: Record<RuleEvent['kind'], string> = {
  create: '新增',
  edit: '編輯',
  toggle: '切換',
  delete: '刪除',
}

const FIELD_LABEL_TRACE: Record<string, string> = {
  type: '類型',
  signal: 'signal',
  targetFolderPath: '目標',
  confidence: '信心',
  source: '來源',
  enabled: '啟用',
  orphaned: '目標遺失',
}

function formatHistoryWhen(at: number): string {
  const diff = Date.now() - at
  if (diff < 0) return '剛剛'
  const sec = Math.floor(diff / 1000)
  if (sec < 60) return `${sec} 秒前`
  const min = Math.floor(sec / 60)
  if (min < 60) return `${min} 分前`
  const hr = Math.floor(min / 60)
  if (hr < 24) return `${hr} 小時前`
  const day = Math.floor(hr / 24)
  if (day < 7) return `${day} 天前`
  const d = new Date(at)
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
}

function shortSignal(sig: string, type: RuleType): string {
  if (type !== 'compound') return sig.length > 40 ? sig.slice(0, 40) + '…' : sig
  try {
    const parsed = JSON.parse(sig) as { conditions: Array<{ type: string; value: string }> }
    return parsed.conditions
      .map((c) => {
        const v = c.value.length > 20 ? c.value.slice(0, 20) + '…' : c.value
        if (c.type === 'domain') return `@${v.replace(/^@/, '')}`
        if (c.type === 'sender') return v
        return `「${v}」`
      })
      .join(' + ')
  } catch {
    return sig.slice(0, 40)
  }
}

function RuleHistorySection() {
  const [events, setEvents] = useState<RuleEvent[] | null>(null)
  const [filterActor, setFilterActor] = useState<'all' | 'user' | 'system'>('all')
  const [filterKind, setFilterKind] = useState<'all' | RuleEvent['kind']>('all')
  const [confirmClear, setConfirmClear] = useState(false)

  const refresh = useCallback(async () => {
    const r = await send<{ events: RuleEvent[] }>({ type: 'getRuleHistory', limit: 100 })
    if (r.ok && r.data) setEvents(r.data.events)
  }, [])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const filtered = useMemo(() => {
    if (!events) return []
    let out = events
    if (filterActor !== 'all') out = out.filter((e) => e.actor === filterActor)
    if (filterKind !== 'all') out = out.filter((e) => e.kind === filterKind)
    // Most-recent first for display
    return [...out].reverse()
  }, [events, filterActor, filterKind])

  const totalUser = useMemo(() => events?.filter((e) => e.actor === 'user').length ?? 0, [events])
  const totalSystem = useMemo(() => events?.filter((e) => e.actor === 'system').length ?? 0, [events])

  async function clearHistory() {
    await send({ type: 'clearRuleHistory' })
    setConfirmClear(false)
    await refresh()
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-sm">
          <ListChecks className="size-4" />
          規則編輯紀錄
          {events && (
            <Badge variant="muted">{events.length}</Badge>
          )}
        </CardTitle>
        <CardDescription>
          所有規則庫變動的審計紀錄(最多 500 筆,顯示最近 100 筆)。「我改的」是你在這個頁面或 plan 階段點停用 / 拆解等動作;「系統」是 AI 確認 / 覆蓋學習 / 初始掃描 / reconcile 自動加的。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-3 text-xs">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-muted-foreground">作者:</span>
          {(['all', 'user', 'system'] as const).map((a) => (
            <button
              key={a}
              type="button"
              onClick={() => setFilterActor(a)}
              className={cn(
                'rounded px-2 py-0.5 text-[11px] border transition-colors',
                filterActor === a
                  ? 'border-foreground bg-foreground text-background'
                  : 'border-border bg-card hover:bg-accent',
              )}
            >
              {a === 'all' ? `全部 ${events?.length ?? 0}` : a === 'user' ? `我改的 ${totalUser}` : `系統 ${totalSystem}`}
            </button>
          ))}
          <span className="text-muted-foreground ml-2">動作:</span>
          {(['all', 'create', 'edit', 'toggle', 'delete'] as const).map((k) => (
            <button
              key={k}
              type="button"
              onClick={() => setFilterKind(k)}
              className={cn(
                'rounded px-2 py-0.5 text-[11px] border transition-colors',
                filterKind === k
                  ? 'border-foreground bg-foreground text-background'
                  : 'border-border bg-card hover:bg-accent',
              )}
            >
              {k === 'all' ? '全部' : KIND_LABEL[k]}
            </button>
          ))}
          <div className="ml-auto">
            {confirmClear ? (
              <span className="flex gap-1">
                <button
                  type="button"
                  className="text-[11px] text-red-700 hover:underline"
                  onClick={() => void clearHistory()}
                >
                  確定清空
                </button>
                <button
                  type="button"
                  className="text-[11px] text-muted-foreground hover:underline"
                  onClick={() => setConfirmClear(false)}
                >
                  取消
                </button>
              </span>
            ) : (
              <button
                type="button"
                className="text-[11px] text-muted-foreground hover:text-foreground"
                onClick={() => setConfirmClear(true)}
              >
                清空紀錄
              </button>
            )}
          </div>
        </div>

        {!events ? (
          <p className="text-muted-foreground italic">載入中…</p>
        ) : filtered.length === 0 ? (
          <p className="text-muted-foreground italic">
            {events.length === 0 ? '尚無紀錄。下次規則變動會在這裡顯示。' : '此過濾條件下無紀錄。'}
          </p>
        ) : (
          <ul className="space-y-1 max-h-[28rem] overflow-y-auto">
            {filtered.map((e, idx) => (
              <li
                key={`${e.at}-${idx}`}
                className={cn(
                  'rounded-md border p-2 space-y-0.5',
                  e.actor === 'user' ? 'border-foreground/30 bg-card' : 'border-border bg-muted/20',
                )}
              >
                <div className="flex items-baseline gap-2 flex-wrap">
                  <span className="text-muted-foreground tabular-nums">{formatHistoryWhen(e.at)}</span>
                  <Badge variant={e.actor === 'user' ? 'default' : 'muted'} className="text-[9px]">
                    {e.actor === 'user' ? '我' : '系統'}
                  </Badge>
                  <span className="font-medium">{KIND_LABEL[e.kind]}</span>
                  {e.kind === 'create' && (
                    <>
                      <Badge variant="outline" className="text-[9px] font-mono">
                        {RULE_TYPE_LABEL_TRACE[e.after.type]}
                      </Badge>
                      <span className="font-mono truncate">{shortSignal(e.after.signal, e.after.type)}</span>
                      <span className="text-muted-foreground">→</span>
                      <span className="font-mono text-muted-foreground truncate">{e.after.targetFolderPath}</span>
                    </>
                  )}
                  {e.kind === 'delete' && (
                    <>
                      <Badge variant="outline" className="text-[9px] font-mono">
                        {RULE_TYPE_LABEL_TRACE[e.before.type]}
                      </Badge>
                      <span className="font-mono truncate line-through">{shortSignal(e.before.signal, e.before.type)}</span>
                      <span className="text-muted-foreground">→</span>
                      <span className="font-mono text-muted-foreground truncate line-through">{e.before.targetFolderPath}</span>
                    </>
                  )}
                  {e.kind === 'toggle' && (
                    <>
                      <Badge variant="outline" className="text-[9px] font-mono">
                        {RULE_TYPE_LABEL_TRACE[e.type]}
                      </Badge>
                      <span className="font-mono truncate">{shortSignal(e.signal, e.type)}</span>
                      <span className="text-muted-foreground">→</span>
                      <span className="font-mono text-muted-foreground truncate">{e.targetFolderPath}</span>
                      <Badge variant={e.enabled ? 'success' : 'warning'} className="text-[9px]">
                        {e.enabled ? '已啟用' : '已停用'}
                      </Badge>
                    </>
                  )}
                  {e.kind === 'edit' && (
                    <>
                      <Badge variant="outline" className="text-[9px] font-mono">
                        {RULE_TYPE_LABEL_TRACE[e.after.type]}
                      </Badge>
                      <span className="font-mono truncate">{shortSignal(e.after.signal, e.after.type)}</span>
                    </>
                  )}
                </div>
                {e.kind === 'edit' && (
                  <div className="text-[10px] text-muted-foreground pl-1 space-y-0.5">
                    {e.changedFields.map((field) => {
                      const beforeVal = (e.before as unknown as Record<string, unknown>)[field]
                      const afterVal = (e.after as unknown as Record<string, unknown>)[field]
                      const label = FIELD_LABEL_TRACE[field] ?? field
                      return (
                        <div key={field} className="font-mono">
                          {label}: <span className="line-through">{String(beforeVal)}</span> → <span className="text-foreground">{String(afterVal)}</span>
                        </div>
                      )
                    })}
                  </div>
                )}
              </li>
            ))}
          </ul>
        )}
      </CardContent>
    </Card>
  )
}

// ---- Stats section -------------------------------------------------------

function StatsSection({ metrics }: { metrics: Metrics }) {
  const total = metrics.moved + metrics.deleted + metrics.foldersCreated + metrics.errors
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-sm">
          <BarChart3 className="size-4" />
          歷史統計
        </CardTitle>
        <CardDescription>累計動作數，自安裝以來。</CardDescription>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-4 gap-3 text-center">
          <StatBox label="移動" value={metrics.moved} tone="text-emerald-700" />
          <StatBox label="刪除" value={metrics.deleted} tone="text-red-700" />
          <StatBox label="新建" value={metrics.foldersCreated} tone="text-amber-700" />
          <StatBox label="錯誤" value={metrics.errors} tone="text-red-700" />
        </div>
        {total === 0 && (
          <p className="mt-3 text-[11px] text-muted-foreground text-center italic">
            尚未執行過歸類批次
          </p>
        )}
      </CardContent>
    </Card>
  )
}

function StatBox({ label, value, tone }: { label: string; value: number; tone: string }) {
  return (
    <div className="rounded-md border border-border bg-muted/30 p-3">
      <div className={`text-2xl font-semibold font-mono tabular-nums ${tone}`}>{value}</div>
      <div className="text-[11px] text-muted-foreground mt-0.5">{label}</div>
    </div>
  )
}

function ScanResults({ state, onReset }: { state: ScanState; onReset: () => void }) {
  const elapsedSec = state.finishedAt ? Math.round((state.finishedAt - state.startedAt) / 1000) : 0
  const errorRows = state.results.filter((r) => r.status === 'error')
  const successRows = state.results.filter((r) => r.status === 'done' && (r.rulesAdded ?? 0) > 0)

  return (
    <div className="space-y-3">
      <div className={`rounded-md border p-3 space-y-1.5 ${errorRows.length === 0 ? 'border-emerald-200 bg-emerald-50' : 'border-amber-200 bg-amber-50'}`}>
        <div className="flex items-center gap-2 text-sm font-medium">
          {errorRows.length === 0 ? <CheckCircle2 className="size-4 text-emerald-700" /> : <AlertTriangle className="size-4 text-amber-700" />}
          完成（{elapsedSec}s）
        </div>
        <div className="text-xs space-y-0.5 pl-6">
          <div>掃描 <span className="font-mono">{state.summary.foldersScanned}</span> 個資料夾、生成 <span className="font-mono">{state.summary.rulesAdded}</span> 條規則</div>
          <div className="text-muted-foreground">
            無外部網域 {state.summary.foldersWithNoExternalDomains} · 錯 {state.summary.errors}
          </div>
        </div>
      </div>

      {successRows.length > 0 && (
        <details className="text-xs">
          <summary className="cursor-pointer text-muted-foreground">展開 {successRows.length} 條規則來源</summary>
          <ul className="mt-2 space-y-1 max-h-72 overflow-y-auto">
            {successRows.map((r) => (
              <li key={r.folderPath} className="rounded-md border border-border bg-muted/30 p-2 space-y-0.5">
                <div className="font-mono truncate">{r.folderPath}</div>
                <div className="flex flex-wrap gap-1">
                  {r.domainsFound?.filter((d) => d.added).map((d) => (
                    <Badge key={`d:${d.domain}`} variant="success" className="font-mono">
                      @{d.domain} ×{d.count}
                    </Badge>
                  ))}
                  {r.sendersFound?.filter((s) => s.added).map((s) => (
                    <Badge key={`s:${s.address}`} variant="success" className="font-mono">
                      {s.address} ×{s.count}
                    </Badge>
                  ))}
                </div>
              </li>
            ))}
          </ul>
        </details>
      )}

      {errorRows.length > 0 && (
        <div className="space-y-1">
          <div className="text-xs font-medium text-red-900">錯誤 ({errorRows.length})</div>
          <ul className="space-y-1">
            {errorRows.slice(0, 10).map((r) => (
              <li key={r.folderPath} className="rounded-md border border-red-200 bg-red-50/50 p-2 text-xs space-y-0.5">
                <div className="font-mono truncate">{r.folderPath}</div>
                <div className="text-red-800 text-[10px]">{r.message}</div>
              </li>
            ))}
          </ul>
        </div>
      )}

      <Button variant="outline" size="sm" onClick={onReset}>
        <RotateCcw /> 清除掃描結果（重置這張卡）
      </Button>
    </div>
  )
}
