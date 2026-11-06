// MV3 service worker — central dispatcher.
//
// The full list of request types handle() serves is SW_MESSAGE_TYPES in
// shared/messages.ts (the switch below is compile-time exhaustive over it).
// Only extension pages (popup / options, including the in-OWA iframe panel)
// may call them; content scripts are limited to CONTENT_SCRIPT_MESSAGE_TYPES
// (same file).

import { OutlookApi, OutlookError, flattenFolderTree } from '@/shared/outlook-api'
import {
  ClassifierError,
  classifyBatch,
  isKnownClaudeModel,
  withClaudeRetry,
  type ClassifierUsage,
} from '@/shared/classifier'
import {
  clearExecuteState,
  dismissUndo,
  executeUndo,
  getExecuteState,
  installUndoExpireListener,
  isExecuteBusy,
  isExecuteRunning,
  recoverStaleExecuteState,
  requestCancel,
  startExecute,
  startRetryFailed,
} from './execute'
import {
  clearScanState,
  getScanState,
  isScanRunning,
  recoverStaleScanState,
  requestScanCancel,
  startInitialScan,
} from './initial-scan'
import { installActionRouter, openOrFocusOwa } from './action-router'
import { logError } from '@/shared/error-log'
import {
  handleBulkRuleAction,
  handleClearDisabledRules,
  handleResolveConflict,
  handleSplitConflictToCompound,
} from './handlers/rules-bulk'
import {
  handleClearErrorLog,
  handleDisableSync,
  handleDismissSyncError,
  handleEnableSync,
  handleGetErrorLog,
  handleGetSyncStatus,
  handleListSyncBackups,
  handlePullSyncNow,
  handlePushSyncNow,
  handleRestoreSyncBackup,
} from './handlers/sync'
import { holdKeepAlive, installKeepAliveListener, releaseKeepAlive } from './keep-alive'
import { installStaleSweepListener, runSweep } from './stale-sweep'
// installSyncListener is the SW-top-level wiring; getSyncStatus / listBackups
// are needed by the inline getOnboardingState handler. Other sync handlers
// live in ./handlers/sync.
import {
  clearCloudState,
  dismissRemoteWipeNotice,
  getSyncStatus,
  installSyncListener,
  listBackups,
  pushNow,
  quiesce,
  readRemoteWipeNotice,
} from './sync-engine'
import { filterFolderActivity } from '@/shared/folder-activity-filter'
import {
  CONTENT_SCRIPT_MESSAGE_TYPES,
  type MessageType,
  type PopupRequest,
  type PopupResponse,
} from '@/shared/messages'
import {
  AI_PARALLEL_CHUNKS,
  BATCH_SIZE_MAX,
  BATCH_SIZE_MIN,
  INBOX_PAGE_SIZE_MAX,
  INBOX_PAGE_SIZE_MIN,
  INBOX_SCAN_MAX_MESSAGES,
  INBOX_SCAN_MIN_MESSAGES,
} from '@/shared/constants'
import {
  addRulesFilteringTombstones,
  applyClosedCaseRetirement,
  applyConfidenceCap,
  deleteRule,
  diffSnapshots,
  encodeCompound,
  extractSubjectSignal,
  findConflicts,
  listRules,
  buildRuleIndex,
  matchBodyCaseWithIndex,
  matchEmailWithIndex,
  mutateRules,
  newRule,
  reconcileRulesAgainstTree,
  type ReconcileSummary,
  ruleBeatsThread,
  snapshotOf,
  toggleRule,
  upsertRule,
} from '@/shared/rules'
import {
  HARVEST_MAX_PAGES,
  HARVEST_PAGE_SIZE,
  filterHarvestCandidates,
  harvestWindowStart,
  nextHarvestWatermark,
  selectHarvestScanItems,
  type HarvestScanItem,
} from '@/background/harvest'
import {
  applyImport,
  type ImportStrategy,
  parseRulesPayload,
  previewImport,
  serializeRules,
} from '@/shared/rule-io'
import {
  addRuleTombstones,
  clearAllAiMemory,
  clearAllRuleTombstones,
  clearRuleHistory,
  clearSkipHistory,
  getConversationMemory,
  getFolderActivity,
  getFolderActivityRefreshAt,
  getFolderCache,
  getLastHarvestAt,
  getMetrics,
  getMailboxIdentity,
  getMovedByUsLedger,
  getRecentlyProcessedIds,
  getRuleEvents,
  getRuleTombstones,
  getSettings,
  getSkipHistory,
  getSkipHistoryCount,
  getStorageTopKeys,
  getStorageUsage,
  getSubjectMemory,
  getUndoSnapshot,
  getWeeklyDigestState,
  mergeFolderActivityScan,
  recordConversationFilings,
  recordRuleEvents,
  recordSubjectFilings,
  setFolderActivityRefreshAt,
  setLastHarvestAt,
  setFolderCache,
  setMailboxIdentity,
  setSettings,
  setWeeklyDigestState,
} from '@/shared/storage'
import { computeRuleHealth } from '@/shared/rule-health'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from '@/shared/normalize'
import type {
  Email,
  MailFolderNode,
  PlanItem,
  Rule,
  RuleEvent,
  Settings,
} from '@/shared/types'
import { getLastTokenError, getOwaToken, peekCachedToken, pingOwa } from './token'

const FOLDER_CACHE_TTL_MS = 24 * 60 * 60 * 1000
const CLASSIFY_STAGE_KEY = 'classifyStage'
const PREFLIGHT_CACHE_KEY = 'preflightCache'
const AI_PROGRESS_KEY = 'aiClassifyProgress'
const POPUP_STATE_KEY = 'popupState'
const AI_CHUNK_SIZE = 25

export type ClassifyStage =
  | { stage: 'idle' }
  | { stage: 'fetching_tree' }
  | { stage: 'fetching_inbox' }
  | { stage: 'matching_rules'; total: number }
  | { stage: 'calling_ai'; toClassify: number }

async function setClassifyStage(stage: ClassifyStage): Promise<void> {
  if (stage.stage === 'idle') {
    await chrome.storage.session.remove(CLASSIFY_STAGE_KEY)
  } else {
    await chrome.storage.session.set({ [CLASSIFY_STAGE_KEY]: stage })
  }
}

/**
 * Group inbox messages by Outlook ConversationId. Returns one representative
 * per thread (most recent ReceivedDateTime) plus a map of siblings keyed by
 * the rep's emailId so the caller can fan out the AI decision after the
 * representative is classified. Emails without a ConversationId are
 * returned as their own rep with no siblings.
 */
function groupEmailsByConversation(emails: Email[]): {
  reps: Email[]
  siblingsByRepId: Map<string, Email[]>
} {
  const byConv = new Map<string, Email[]>()
  const solos: Email[] = []
  for (const m of emails) {
    if (!m.ConversationId) {
      solos.push(m)
      continue
    }
    const arr = byConv.get(m.ConversationId) ?? []
    arr.push(m)
    byConv.set(m.ConversationId, arr)
  }
  const reps: Email[] = []
  const siblingsByRepId = new Map<string, Email[]>()
  for (const members of byConv.values()) {
    // Most-recent message becomes representative — it carries the latest
    // direction of the thread (e.g. final court ruling decided after 5
    // back-and-forth emails).
    members.sort((a, b) => b.ReceivedDateTime.localeCompare(a.ReceivedDateTime))
    const rep = members[0]!
    reps.push(rep)
    if (members.length > 1) siblingsByRepId.set(rep.Id, members.slice(1))
  }
  // Preserve the original input order for stable plan rendering — solos
  // were already in order, threaded reps come after.
  reps.push(...solos)
  return { reps, siblingsByRepId }
}

/**
 * Build a PlanItem for a thread sibling by cloning the rep's classified
 * action onto the sibling's identity. Keeps targetFolderId / target path
 * / suggestedFolderName etc. from the rep; rewrites identity-bound fields
 * (emailId, Subject, From, BodyPreview, ConversationId) to the sibling's
 * own. Marked `threadInherited: true` so the UI can label it.
 */
function inheritPlanItemForSibling(rep: PlanItem, sibling: Email): PlanItem {
  return {
    ...rep,
    emailId: sibling.Id,
    emailSubject: sibling.Subject ?? '',
    emailFrom: sibling.From?.EmailAddress?.Address ?? '',
    bodyPreview: sibling.BodyPreview ?? '',
    conversationId: sibling.ConversationId,
    threadInherited: true,
    reason: rep.reason
      ? `${rep.reason}（同對話 thread 繼承）`
      : '同對話 thread 繼承',
  }
}

// Folder-tree refresh failure streak (2026-11-05). getOrFetchFolderTree used
// to fall back to the stale cache with only a console.warn — forever, with
// nothing user-visible. It still serves the cache (a stale tree beats a dead
// popup), but now logs to the central error log ONCE per failure streak and
// records the streak in session storage so callers / getStatus can surface
// 「資料夾清單未更新」. Session storage (not a module variable) so an SW
// idle-restart mid-streak doesn't re-log every wake.
const FOLDER_TREE_STALE_KEY = 'folderTreeStale'

export type FolderTreeStaleness = {
  /** When the current failure streak started (unix ms). */
  since: number
  lastError: string
  /** updatedAt of the cache being served instead. */
  cacheUpdatedAt: string
}

async function getFolderTreeStaleness(): Promise<FolderTreeStaleness | null> {
  try {
    const r = await chrome.storage.session.get(FOLDER_TREE_STALE_KEY)
    return (r[FOLDER_TREE_STALE_KEY] as FolderTreeStaleness | undefined) ?? null
  } catch {
    return null
  }
}

async function getOrFetchFolderTreeWithMeta(
  forceFresh: boolean,
): Promise<{ tree: MailFolderNode[]; stale: FolderTreeStaleness | null }> {
  const cache = await getFolderCache()
  const cacheFresh = cache && Date.now() - new Date(cache.updatedAt).getTime() < FOLDER_CACHE_TTL_MS

  if (!forceFresh && cacheFresh) {
    return { tree: cache!.tree, stale: await getFolderTreeStaleness() }
  }

  try {
    const { tree, excludedFolderIds } = await api.fetchFolderTreeWithMeta()
    await setFolderCache({ updatedAt: new Date().toISOString(), tree, excludedFolderIds })
    await chrome.storage.session.remove(FOLDER_TREE_STALE_KEY).catch(() => {})
    return { tree, stale: null }
  } catch (e) {
    // Transient failures (network, 401, etc) shouldn't block the user when a
    // stale tree is still usable — keep serving it, but make it visible.
    if (cache) {
      const msg = e instanceof Error ? e.message : String(e)
      const prev = await getFolderTreeStaleness()
      const stale: FolderTreeStaleness = {
        since: prev?.since ?? Date.now(),
        lastError: msg.slice(0, 300),
        cacheUpdatedAt: cache.updatedAt,
      }
      if (!prev) {
        await logError('outlook:folder-tree', msg, {
          ageMs: Date.now() - new Date(cache.updatedAt).getTime(),
          forceFresh,
        }).catch(() => {})
      }
      await chrome.storage.session.set({ [FOLDER_TREE_STALE_KEY]: stale }).catch(() => {})
      return { tree: cache.tree, stale }
    }
    throw e
  }
}

async function getOrFetchFolderTree(forceFresh: boolean): Promise<MailFolderNode[]> {
  return (await getOrFetchFolderTreeWithMeta(forceFresh)).tree
}

// Synchronous in-flight guard for harvestManualFilings (2026-11-05,
// invariant 6): set before the handler's first await so a double-click (or
// two Options tabs) can't run two overlapping mailbox scans.
let harvestInFlight = false

// Cache the resolved REST base URL across requests so we don't issue
// chrome.tabs.query on every API call. Expires after 60s; long enough to
// stay warm during a batch, short enough to pick up tab changes.
const OWA_REST_HOST_PATTERNS = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
]
const DEFAULT_REST_BASE = 'https://outlook.office.com/api/v2.0'
let cachedBaseUrl: { value: string; expiresAt: number } | null = null

// Concurrency guard for the wipeAllRules handler so two near-simultaneous
// invocations (e.g., user double-clicks "全部刪除") can't interleave their
// quiesce/clearCloud sequences and tear down sync state twice.
let wipeInFlight = false

async function resolveOutlookBase(): Promise<string> {
  if (cachedBaseUrl && Date.now() < cachedBaseUrl.expiresAt) {
    return cachedBaseUrl.value
  }
  let base = DEFAULT_REST_BASE
  try {
    const tabs = await chrome.tabs.query({ url: OWA_REST_HOST_PATTERNS })
    const tab = tabs.find((t) => t.active) ?? tabs[0]
    if (tab?.url) {
      const origin = new URL(tab.url).origin
      base = `${origin}/api/v2.0`
    }
  } catch {
    /* keep default */
  }
  cachedBaseUrl = { value: base, expiresAt: Date.now() + 60_000 }
  return base
}

const api = new OutlookApi(
  async (opts) => {
    const r = await getOwaToken({ force: opts?.force })
    if (!r.ok) throw new Error(r.message)
    return r.secret
  },
  resolveOutlookBase,
)

// Own mailbox domain (2026-11-05). With internalDomains empty (the default),
// learning used to mint `domain: <own domain>` from a colleague's mail, and
// since domain rules match ANY address incl. To/Cc, that one rule claimed
// nearly every inbound email. rules.ts now keeps own-domain addresses from
// triggering AUTO domain rules (and learning never mints them) — given the
// domain. Resolved from GET /me once and cached for a week; any failure falls
// back to the cached value, or to none (the pre-fix behaviour).
const MAILBOX_IDENTITY_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000
let identityLookupFailedAt = 0

// `fetch: false` = cache only: classify preflight refreshes the identity, so
// the AI pass right after it and the background harvest never spend a call.
async function resolveOwnDomains(opts: { fetch?: boolean } = {}): Promise<string[]> {
  const cached = await getMailboxIdentity().catch(() => null)
  const cachedAt = cached ? Date.parse(cached.at) : NaN
  const fresh =
    cached !== null && Number.isFinite(cachedAt) && Date.now() - cachedAt < MAILBOX_IDENTITY_MAX_AGE_MS
  if (fresh || opts.fetch === false) return cached ? [cached.domain] : []
  // A lookup that just failed isn't retried for an hour (SW lifetime), so a
  // persistently failing /me can't add its timeout to every preflight.
  if (Date.now() - identityLookupFailedAt < 60 * 60_000) return cached ? [cached.domain] : []
  try {
    // Bounded: a failing /me would otherwise sit in request()'s 5xx backoff
    // (1 s + 3 s + 9 s) and hold up classify preflight — the identity is a
    // nice-to-have, never worth that.
    const address = (await api.getMyEmail(AbortSignal.timeout(3000))).trim().toLowerCase()
    const at = address.lastIndexOf('@')
    if (at > 0 && at < address.length - 1) {
      const identity = { address, domain: address.slice(at + 1), at: new Date().toISOString() }
      await setMailboxIdentity(identity)
      return [identity.domain]
    }
  } catch (e) {
    console.warn('[mail-organizer] mailbox identity lookup failed (non-fatal)', e)
  }
  identityLookupFailedAt = Date.now()
  return cached ? [cached.domain] : []
}

chrome.runtime.onInstalled.addListener((details) => {
  console.debug('[mail-organizer] installed', details.reason)
  void recoverStaleStates(`installed (${details.reason})`)
})

chrome.runtime.onStartup.addListener(() => {
  console.debug('[mail-organizer] startup')
  void recoverStaleStates('browser-startup')
})

/**
 * Module-load recovery. MV3 service workers can be idle-killed mid-operation
 * (the executor Promise dies but the persisted `inProgress` flag stays true),
 * which would block all future operations. On every SW (re)load we sweep for
 * stale states and mark them failed so the user isn't stuck. SW wake from
 * idle doesn't fire onStartup, so we MUST do it at module load too.
 */
void recoverStaleStates('module-load')

// Storage usage check on SW startup. Without `unlimitedStorage`,
// chrome.storage.local writes fail once the quota (read at runtime — 10 MB on
// current Chrome; the old hard-coded 5 MB double-counted) is exceeded, so warn
// at >80% before rule writes start dropping. With `unlimitedStorage` there is
// no quota and `approaching` is never true. The Options page shows the same
// numbers live via 'getStorageUsage' (2026-11-05).
void getStorageUsage()
  .then((usage) => {
    if (usage.approaching && usage.percentUsed !== null) {
      console.warn(
        '[mail-organizer] chrome.storage.local usage at',
        `${(usage.percentUsed * 100).toFixed(1)}%`,
        `(${usage.bytesInUse} / ${usage.quotaBytes} bytes)`,
        '— consider exporting rules + clearing skip-history / tombstones.',
      )
    }
  })
  .catch(() => {
    /* silent — startup check is best-effort */
  })

// Keep-alive alarm listener (lazy — only fires when something is holding)
installKeepAliveListener()

// Auto-expire the undo snapshot when its alarm fires (30s window).
installUndoExpireListener()

// Daily sweep: auto-disable rules with 0 hits older than 90 days.
// Idempotent — re-registering the alarm is safe.
installStaleSweepListener()

// Cross-machine sync listener. Detects remote pushes (other machine
// signed into the same browser sync account) and pulls them in.
installSyncListener()

// Dynamic toolbar-icon routing. When no OWA tab is open, clicking the
// icon opens one (solving the chicken-and-egg "extension needs Outlook,
// but I haven't opened Outlook yet" problem). When an OWA tab IS open,
// the popup behaves as normal.
installActionRouter()

// OWA floating-button visibility (2026-11-05). The owa-fab content script
// used to read chrome.storage.local['settings'] — which holds the Claude
// API key — and console.log the whole object on every Outlook page load;
// its storage.onChanged listener also received the full old+new settings
// (key included) on every settings write. The content script now never
// touches extension storage: it asks via 'getOwaFabVisibility' and the SW
// pushes a one-boolean 'owaFabVisibility' message when the setting flips
// (any write path: Options toggle, sync pull, backup restore).
chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local' || !changes.settings) return
  const before = owaFabShown(changes.settings.oldValue)
  const after = owaFabShown(changes.settings.newValue)
  if (before !== after) void broadcastOwaFabVisibility(after)
})

function owaFabShown(settings: unknown): boolean {
  // Default ON — first-run users get the FAB without opting in.
  return (settings as { showOwaFab?: boolean } | undefined)?.showOwaFab !== false
}

async function broadcastOwaFabVisibility(show: boolean): Promise<void> {
  try {
    const tabs = await chrome.tabs.query({ url: OWA_REST_HOST_PATTERNS })
    await Promise.all(
      tabs.map((t) =>
        t.id == null
          ? undefined
          : chrome.tabs.sendMessage(t.id, { type: 'owaFabVisibility', show }).catch(() => {
              /* tab without our content script (loading / orphaned) */
            }),
      ),
    )
  } catch (e) {
    console.warn('[mail-organizer] owa-fab visibility broadcast failed (non-fatal)', e)
  }
}

// getStatus count cache (2026-11-05). getStatus runs on every popup open /
// focus and once a second while the popup waits for a token, and it only
// needs two COUNTS — yet it read and deserialised the whole rule library and
// the whole folder tree each time. Cache the counts; drop them whenever
// storage.onChanged reports a 'rules' or 'folderCache' write (from any
// context). `statusCountsGen` closes the race where a write lands while a
// recompute is mid-read: that recompute's result is returned but not cached.
type StatusCounts = {
  rulesCount: number
  folderCacheCount: number
  folderCacheUpdatedAt: string | null
}
let statusCounts: StatusCounts | null = null
let statusCountsGen = 0

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local') return
  if (changes.rules || changes.folderCache) {
    statusCountsGen++
    statusCounts = null
  }
})

async function getStatusCounts(): Promise<StatusCounts> {
  if (statusCounts) return statusCounts
  const gen = statusCountsGen
  const [cache, rules] = await Promise.all([getFolderCache(), listRules()])
  const counts: StatusCounts = {
    rulesCount: rules.length,
    folderCacheCount: cache ? flattenFolderTree(cache.tree).length : 0,
    folderCacheUpdatedAt: cache?.updatedAt ?? null,
  }
  if (gen === statusCountsGen) statusCounts = counts
  return counts
}

// Single-flight guard for refreshFolderActivity (2026-11-05). The toolbar
// popup and the in-OWA panel both auto-refresh on open, and each refresh is
// a sequential scan of up to 200 folders (minutes, at the shared 1 req/s
// Outlook token bucket); two overlapping runs doubled the API load and the
// wait for an identical result. Concurrent callers share the in-flight
// promise. Checked + set synchronously in the handler, before any await
// (TOCTOU). Cleared in `finally`; a pathologically hung scan is also dropped
// whenever Chrome recycles the SW (module state is per-SW-lifetime).
let refreshFolderActivityInFlight: Promise<PopupResponse> | null = null

/**
 * Phase 3 — manual refresh via Graph API. Scans every case folder under
 * the user's root path, queries each for its most-recent message, and
 * merges results into the activity store. Catches mail the user dragged
 * in manually (bypassing the extension's classify flow) that otherwise
 * wouldn't appear in the quick-jump panel. Only ever called through the
 * refreshFolderActivity single-flight guard in handle().
 */
async function scanFolderActivity(): Promise<PopupResponse> {
  try {
    const cache = await getFolderCache()
    if (!cache) {
      return { ok: false, code: 'NO_CACHE', message: '尚未抓過資料夾、請先跑一次歸類' }
    }
    const tok = await getOwaToken()
    if (!tok.ok) {
      return { ok: false, code: 'NO_TOKEN', message: '無法取得 OWA token' }
    }
    const settings = await getSettings()
    // Scan folders under the user-configured primary root path (e.g.
    // '案件') to avoid wasting API calls on archive trees.
    // Empty primaryRootPath means onboarding's incomplete — surface
    // that to the user rather than guessing.
    const rootPath = settings.primaryRootPath
    if (!rootPath) {
      return {
        ok: false,
        code: 'NO_ROOT_CONFIGURED',
        message: '尚未設定主要根資料夾(設定頁 → 歸類偏好 → 主要根資料夾)',
      }
    }
    const root = flattenFolderTree(cache.tree).find((f) => f.path === rootPath)
    if (!root) {
      return { ok: false, code: 'NO_ROOT', message: `找不到「${rootPath}」根目錄` }
    }
    // Walk the root subtree, collecting LEAF folders (no children) — these
    // are case-level folders (e.g. 法顧/vendor). Skipping interior
    // category nodes (e.g. 法顧) which don't have direct case mail.
    const targets: MailFolderNode[] = []
    function collectLeaves(nodes: MailFolderNode[]): void {
      for (const n of nodes) {
        if (n.children.length === 0) targets.push(n)
        else collectLeaves(n.children)
      }
    }
    collectLeaves(root.children)
    // Apply exclude-prefixes too — same blocklist the classify flow uses.
    const excludes = settings.excludeFolderPrefixes
    const scanTargets = targets.filter(
      (n) => !excludes.some((p) => n.path.startsWith(p)),
    )
    // Cap to avoid runaway scans on huge mailboxes — 200 covers a year's
    // worth of cases comfortably.
    const SCAN_CAP = 200
    const limited = scanTargets.slice(0, SCAN_CAP)
    const scans: Array<{
      folderId: string
      folderPath: string
      latestMessageAt: string
      latestMessage?: { subject: string; from: string; receivedAt: string }
    }> = []
    let scanned = 0
    let errored = 0
    // Sequential with small interim delays. Graph API allows bursts but
    // sustained N rapid requests can hit throttle; sequential at 50/s is
    // safe and the user's clicking refresh manually so wait time is
    // acceptable.
    for (const node of limited) {
      try {
        const msgs = await api.listFolderMessages(node.id, {
          top: 1,
          // Subject + From added so the popup row's hover tooltip can
          // surface "what's the latest in this folder" without an
          // extra round-trip per row.
          select: 'Id,Subject,From,ReceivedDateTime',
        })
        if (msgs.length > 0) {
          const top = msgs[0]!
          const latest = top.ReceivedDateTime
          if (latest) {
            const subject = (top.Subject ?? '').slice(0, 120)
            const from = top.From?.EmailAddress?.Address ?? ''
            scans.push({
              folderId: node.id,
              folderPath: node.path,
              latestMessageAt: latest,
              latestMessage: { subject, from, receivedAt: latest },
            })
          }
        }
        scanned++
      } catch (e) {
        errored++
        console.warn(`[mail-organizer] refresh scan failed for ${node.path}`, e)
      }
    }
    await mergeFolderActivityScan(scans)
    // Stamp "we just refreshed" only when at least one folder's
    // Graph API call succeeded. Bug #L: previously we stamped even
    // when ALL scans threw (transient OWA token expiry, network
    // blip), which would block the popup's 30-min auto-refresh
    // policy from retrying — user would see stale data for the
    // next 30 minutes despite repeated popup opens.
    if (scanned > 0) {
      await setFolderActivityRefreshAt(new Date().toISOString())
    }
    return {
      ok: true,
      data: {
        scanned,
        errored,
        matched: scans.length,
        capped: scanTargets.length > SCAN_CAP,
      },
    }
  } catch (e) {
    return {
      ok: false,
      code: 'SCAN_ERROR',
      message: e instanceof Error ? e.message : String(e),
    }
  }
}

async function recoverStaleStates(reason: string): Promise<void> {
  try {
    await recoverStaleExecuteState()
    await recoverStaleScanState()
    // AI classify progress: if marked in-progress with no live worker → mark done with error
    const aiRaw = await chrome.storage.session.get(AI_PROGRESS_KEY)
    const aiProg = aiRaw[AI_PROGRESS_KEY] as { done?: boolean; aiError?: unknown } | undefined
    if (aiProg && aiProg.done === false) {
      await chrome.storage.session.set({
        [AI_PROGRESS_KEY]: {
          ...aiProg,
          done: true,
          aiError: aiProg.aiError ?? {
            code: 'SW_INTERRUPTED',
            message: 'Service worker 中斷、AI 分類未完成。可重新開始歸類。',
          },
        },
      })
      console.warn(`[mail-organizer] recovered stuck aiClassifyProgress (${reason})`)
    }
    // Clear stale stage display
    await chrome.storage.session.remove(CLASSIFY_STAGE_KEY)
  } catch (e) {
    console.error('[mail-organizer] recoverStaleStates failed', e)
  }
}

/**
 * Sender allowlist (2026-11-05). The dispatcher used to ignore the sender
 * entirely, so any script running as our content script — e.g. a compromised
 * OWA renderer, which hosts the content scripts — could send wipeAllRules,
 * exportDiagnostic (settings + full rule library), setSettings,
 * importRules… Non-extension-page senders may now only use
 * CONTENT_SCRIPT_MESSAGE_TYPES (shared/messages.ts; today just the FAB
 * visibility boolean). The in-OWA panel is NOT a content script: it is our
 * popup page in an iframe (its sender.url is the chrome-extension:// page URL,
 * filled in by the browser process), so it keeps full access.
 */
const CONTENT_SCRIPT_ALLOWED: ReadonlySet<string> = new Set<MessageType>(
  CONTENT_SCRIPT_MESSAGE_TYPES,
)

function isExtensionPageSender(sender: chrome.runtime.MessageSender): boolean {
  // getURL('') is this extension's own origin + '/', i.e.
  // `chrome-extension://${chrome.runtime.id}/` on Chrome / Edge.
  const ownPrefix = chrome.runtime.getURL('')
  return (
    sender.id === chrome.runtime.id &&
    typeof sender.url === 'string' &&
    sender.url.startsWith(ownPrefix)
  )
}

chrome.runtime.onMessage.addListener(
  (raw: unknown, sender, sendResponse: (resp: PopupResponse) => void) => {
    const type = (raw as { type?: unknown } | null | undefined)?.type
    const allowed =
      isExtensionPageSender(sender) ||
      (typeof type === 'string' && CONTENT_SCRIPT_ALLOWED.has(type))
    if (!allowed) {
      // Console only — writing refusals to the persisted error log would hand
      // an attacker a way to flush real entries out of the 200-entry log.
      console.warn('[mail-organizer] refused message from non-extension-page sender', {
        type: typeof type === 'string' ? type.slice(0, 64) : typeof type,
        origin: sender.origin,
      })
      sendResponse({
        ok: false,
        code: 'FORBIDDEN',
        message: 'This request is only accepted from Mail Organizer extension pages',
      })
      return false
    }
    if (typeof type !== 'string') {
      sendResponse({ ok: false, code: 'UNKNOWN_TYPE', message: 'Request has no string `type`' })
      return false
    }
    // Unknown strings are fine here — handle()'s default answers UNKNOWN_TYPE.
    const msg = raw as PopupRequest
    void (async () => {
      try {
        sendResponse(await handle(msg))
      } catch (e) {
        // Surface unhandled handler errors in the centralised error log
        // (#2) so a silent failure shows up in the Options "錯誤紀錄"
        // card. Original console.warn-only behaviour was invisible to
        // anyone without the SW devtools console open.
        const errMsg = e instanceof Error ? e.message : String(e)
        if (e instanceof OutlookError) {
          await logError('outlook:request', errMsg, {
            handler: msg.type,
            status: e.status,
          }).catch(() => {})
          sendResponse({
            ok: false,
            code: `OUTLOOK_${e.status}`,
            message: e.message,
          })
        } else if (e instanceof ClassifierError) {
          await logError('classify:api', errMsg, {
            handler: msg.type,
            code: e.code,
          }).catch(() => {})
          sendResponse({
            ok: false,
            code: `CLASSIFIER_${e.code}`,
            message: e.message,
          })
        } else {
          await logError('handler:unhandled', errMsg, { handler: msg.type }).catch(
            () => {},
          )
          sendResponse({
            ok: false,
            code: 'UNHANDLED',
            message: errMsg,
          })
        }
      }
    })()
    return true
  },
)

// ---- Classify run ownership (2026-11-05) ----------------------------------
//
// Every classify cycle is a RUN with an id minted by classifyPreflight. The
// SW tracks the one current run (memory + session storage, so a restarted
// worker still knows it). A newer preflight supersedes the current run;
// cancelClassify / clearAiClassifyProgress invalidate it. The classifyAi loop
// checks ownership before each Claude call and each progress write, and its
// in-flight call is aborted, so a cancelled or superseded loop stops spending
// and can't overwrite / end / delete a newer run's state. Before this, 取消 /
// 完成 only deleted storage keys while the loop kept calling Claude for every
// remaining chunk; 取消 then 開始歸類 ran two loops on the same key (double
// spend), the old loop's done:true ended the new poll early, and its finally
// deleted the new run's preflight cache.
//
// The run also records which emails were submitted to execute while it was
// current (startExecute). Its progress never offers them again — that, not the
// old startedAt-vs-execute heuristic, is what keeps executed mail from being
// resurrected after a partial execute (「先執行 N 個」 while AI still runs).

const CLASSIFY_RUN_KEY = 'classifyRun'

type ClassifyRunRecord = {
  runId: string
  createdAt: number
  /** Set when classifyAi starts this run's loop; a second start would double-spend. */
  aiStartedAt?: number
  /** Email ids submitted to execute while this run was current. */
  executedIds: string[]
}

/** Memory mirror of the session record; `undefined` = not loaded in this SW instance yet. */
let classifyRun: ClassifyRunRecord | null | undefined
/** Abort controllers of classify loops running in THIS SW instance, by runId. */
const classifyLoops = new Map<string, AbortController>()

async function loadClassifyRun(): Promise<ClassifyRunRecord | null> {
  if (classifyRun !== undefined) return classifyRun
  const stored = (await chrome.storage.session.get(CLASSIFY_RUN_KEY))[CLASSIFY_RUN_KEY] as
    | ClassifyRunRecord
    | undefined
  // Re-check after the await: another handler may have set the run meanwhile.
  if (classifyRun === undefined) classifyRun = stored ?? null
  return classifyRun
}

/**
 * Memory is updated synchronously, before the storage write is even issued —
 * ownership checks read memory with no await in between (the TOCTOU iron law
 * from the sync engine), so a cancel takes effect at once.
 */
function setClassifyRun(next: ClassifyRunRecord | null): Promise<void> {
  classifyRun = next
  return next
    ? chrome.storage.session.set({ [CLASSIFY_RUN_KEY]: next })
    : chrome.storage.session.remove(CLASSIFY_RUN_KEY)
}

function ownsClassifyRun(runId: string): boolean {
  return classifyRun != null && classifyRun.runId === runId
}

function abortClassifyLoopsExcept(keepRunId: string | null): void {
  for (const [id, ctrl] of classifyLoops) {
    if (id !== keepRunId) ctrl.abort()
  }
}

function newClassifyRunId(): string {
  const rand =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2)
  return `run-${Date.now().toString(36)}-${rand}`
}

/**
 * Invalidate the current run (or only `runId`, when given — a stale cancel
 * must not kill a newer run started from another popup surface), abort its
 * loop and clear its progress + preflight cache.
 */
async function cancelClassifyRun(runId?: string): Promise<boolean> {
  const current = await loadClassifyRun()
  if (runId && current && current.runId !== runId) return false
  const dropRun = setClassifyRun(null)
  abortClassifyLoopsExcept(null)
  await dropRun
  await chrome.storage.session.remove([AI_PROGRESS_KEY, PREFLIGHT_CACHE_KEY])
  // A preflight that claimed a new run meanwhile owns the stage display.
  if (classifyRun == null) await setClassifyStage({ stage: 'idle' })
  return true
}

/** startExecute hook: these emails are leaving the inbox — never re-offer them. */
async function recordClassifyRunExecuted(ids: string[]): Promise<void> {
  try {
    const current = await loadClassifyRun()
    if (!current || ids.length === 0 || !ownsClassifyRun(current.runId)) return
    const merged = new Set(classifyRun!.executedIds)
    for (const id of ids) merged.add(id)
    await setClassifyRun({ ...classifyRun!, executedIds: [...merged] })
  } catch (e) {
    console.warn('[mail-organizer] recordClassifyRunExecuted failed (non-fatal)', e)
  }
}

/** What classifyPreflight hands classifyAi through session storage. */
type PreflightCache = {
  runId: string
  unmatchedEmails: Email[]
  tree: MailFolderNode[]
  rulePlan: PlanItem[]
  threadHints?: Record<string, string>
  createdAt: number
}

/**
 * AI progress as stored. `aiPlan` holds the loop's output minus executed ids;
 * `aiSeq[k]` is the append position of `aiPlan[k]` in the loop's output and
 * `aiTotal` its length — the poll cursor, which stays valid while executed
 * items are filtered out of the stored list.
 */
type StoredAiProgress = {
  runId: string
  totalEmails: number
  completedEmails: number
  chunks: number
  completedChunks: number
  rulePlan: PlanItem[]
  folderTree: MailFolderNode[]
  aiPlan: PlanItem[]
  aiSeq: number[]
  aiTotal: number
  skippedByUser?: number
  usage?: ClassifierUsage
  aiError?: { code: string; message: string } | null
  startedAt: number
  done: boolean
}

type AiProgressView = 'full' | 'poll' | 'meta'

/**
 * What getAiClassifyProgress returns (item 14). The popup used to receive the
 * whole folder tree (~190 KB at 800 folders) plus both full plans every
 * 500 ms. Now: 'full' (resume / consume — once) carries the plans + tree;
 * 'poll' carries only the AI items appended since the popup's cursor; 'meta'
 * carries counts only. Executed ids are filtered again here because an
 * execute that starts after the loop's last write has no write to ride on.
 */
function viewAiProgress(
  p: StoredAiProgress,
  executedIds: readonly string[],
  view: AiProgressView,
  aiFrom: number,
) {
  const executed = new Set(executedIds)
  const rulePlan = p.rulePlan.filter((i) => !executed.has(i.emailId))
  const aiPlan: PlanItem[] = []
  const aiSeq: number[] = []
  p.aiPlan.forEach((item, k) => {
    if (executed.has(item.emailId)) return
    aiPlan.push(item)
    aiSeq.push(p.aiSeq?.[k] ?? k)
  })
  const meta = {
    runId: p.runId,
    totalEmails: p.totalEmails,
    completedEmails: p.completedEmails,
    chunks: p.chunks,
    completedChunks: p.completedChunks,
    ruleCount: rulePlan.length,
    aiCount: aiPlan.length,
    aiUnresolvedCount: aiPlan.filter((x) => x.source === 'unresolved').length,
    // Never-classified placeholders (failed / refused chunks). A finished run
    // offering nothing else is residue: re-classifying that mail fresh beats
    // re-showing stale failure rows.
    aiUnclassifiedCount: aiPlan.filter((x) => x.unclassified === true).length,
    aiCursor: p.aiTotal ?? p.aiPlan.length,
    skippedByUser: p.skippedByUser,
    usage: p.usage,
    aiError: p.aiError,
    startedAt: p.startedAt,
    done: p.done,
  }
  if (view === 'meta') return meta
  if (view === 'poll') return { ...meta, aiPlan: aiPlan.filter((_, k) => aiSeq[k]! >= aiFrom) }
  return { ...meta, rulePlan, aiPlan, folderTree: p.folderTree }
}

// ---- Inbox paging (2026-11-05, inbox starvation) ---------------------------
//
// Preflight and peekNextBatch used to fetch ONE page (top = batchSize +
// ghosts) and filter skipHistory / recentlyProcessed / flagged client-side.
// Once the newest page was all kept or flagged mail, classify reported 0
// emails although older mail sat right below. Follow @odata.nextLink until we
// have batchSize eligible messages or hit a cap.
// Sized from the batch (2026-11-05, 放寬並加速): pages of
// [INBOX_PAGE_SIZE_MIN, INBOX_PAGE_SIZE_MAX] messages, total scanned = 3× the
// batch within [INBOX_SCAN_MIN_MESSAGES, INBOX_SCAN_MAX_MESSAGES], so a
// 500-email batch isn't cut short at a fixed 500-message / 5-page budget once
// kept or flagged mail sits in between, and a 25-email batch still needs at
// most ~6 calls to look 500 messages deep.
export function inboxScanBudget(
  batchSize: number,
  ghosts: number,
): { pageSize: number; maxMessages: number; maxPages: number } {
  const pageSize = Math.min(
    INBOX_PAGE_SIZE_MAX,
    Math.max(INBOX_PAGE_SIZE_MIN, batchSize + Math.min(ghosts, 100)),
  )
  const maxMessages = Math.min(
    INBOX_SCAN_MAX_MESSAGES,
    Math.max(INBOX_SCAN_MIN_MESSAGES, batchSize * 3),
  )
  return { pageSize, maxMessages, maxPages: Math.ceil(maxMessages / pageSize) + 1 }
}

/** A batchSize from a message is clamped like the stored setting. */
function clampBatchSize(n: unknown, fallback: number): number {
  return typeof n === 'number' && Number.isFinite(n)
    ? Math.max(BATCH_SIZE_MIN, Math.min(BATCH_SIZE_MAX, Math.floor(n)))
    : fallback
}

async function scanInboxForBatch(opts: {
  batchSize: number
  select?: string
  skipIds: ReadonlySet<string>
  recentlyProcessed: ReadonlySet<string>
}): Promise<{
  eligible: Email[]
  preFilteredCount: number
  flaggedCount: number
  scanned: number
  moreMayExist: boolean
}> {
  // Over-fetch the first page by the ghost count (just-moved mail Outlook
  // still lists) so a fully-stale top-N can't starve the batch on its own.
  const budget = inboxScanBudget(opts.batchSize, opts.recentlyProcessed.size)
  const top = budget.pageSize
  const eligible: Email[] = []
  let preFilteredCount = 0
  let flaggedCount = 0
  let scanned = 0
  let pages = 0
  let nextLink: string | null = null
  // $skip paging shifts when mail arrives between two page reads: the next
  // page then re-lists the previous page's tail. Without this the same email
  // could land in one batch twice (duplicate plan rows, a 404 on execute,
  // double learning input).
  const seen = new Set<string>()
  do {
    let page: { messages: Email[]; nextLink: string | null }
    try {
      page = await api.listInboxMessagesPage({ top, nextLink, select: opts.select })
    } catch (e) {
      // First page failing = nothing to classify, surface it as before. A
      // later page failing just ends the scan with what we already have.
      if (pages === 0) throw e
      await logError('classify:inbox-page', e instanceof Error ? e.message : String(e), {
        page: pages + 1,
        status: e instanceof OutlookError ? e.status : undefined,
      })
      nextLink = null
      break
    }
    pages++
    scanned += page.messages.length
    for (const m of page.messages) {
      if (seen.has(m.Id)) continue
      seen.add(m.Id)
      if (opts.skipIds.has(m.Id)) {
        preFilteredCount++
        continue
      }
      // Ghosts (just moved/deleted by us, still listed) are invisible
      // bookkeeping — not counted anywhere the user sees.
      if (opts.recentlyProcessed.has(m.Id)) continue
      // Skip flagged mail (hardcoded always-on since 2026-05-27): the user is
      // tracking it. Read live from this fetch, never persisted, so removing
      // the flag in Outlook brings the email back next run.
      if (m.Flag?.FlagStatus === 'Flagged') {
        flaggedCount++
        continue
      }
      eligible.push(m)
    }
    nextLink = page.nextLink
  } while (
    eligible.length < opts.batchSize &&
    nextLink &&
    pages < budget.maxPages &&
    scanned < budget.maxMessages
  )
  return {
    eligible,
    preFilteredCount,
    flaggedCount,
    scanned,
    moreMayExist: eligible.length > opts.batchSize || nextLink != null,
  }
}

/** Banner text for a classify failure; the popup wraps it in 「AI 分類失敗：…」. */
function describeClassifyFailure(e: unknown, leftUnclassified: number): string {
  const ce = e instanceof ClassifierError ? e : null
  const tail = `剩餘 ${leftUnclassified} 封未分類、下一批會重新分類`
  switch (ce?.code) {
    case 'MODEL_UNAVAILABLE':
    case 'REFUSAL':
    case 'PARSE_ERROR':
    case 'NO_API_KEY':
      return ce.message
    case 'HTTP_401':
    case 'HTTP_403':
      return `Claude API key 無效或沒有權限（${ce.status}），請到設定頁更新 API key；${tail}`
    case 'HTTP_429':
      return `Claude API 流量限制（429），等待重試後仍未恢復，${tail}`
    case 'HTTP_529':
      return `Claude 服務暫時過載（529），等待重試後仍未恢復，${tail}`
    case 'NETWORK':
      return `連不上 Claude API（網路中斷），${tail}`
    default:
      if (ce && (/^HTTP_5\d\d$/.test(ce.code) || ce.code === 'STREAM_ERROR')) {
        return `Claude 服務暫時異常（${ce.code.replace('HTTP_', '')}），等待重試後仍未恢復，${tail}`
      }
      return ce ? ce.message : e instanceof Error ? e.message : String(e)
  }
}

async function handle(msg: PopupRequest): Promise<PopupResponse> {
  // No `await` may run before the switch: single-flight guards inside cases
  // (refreshFolderActivity) rely on check+set happening synchronously.
  switch (msg.type) {
    // ---- Content-script queries ---------------------------------------------

    case 'getOwaFabVisibility': {
      // Only the boolean leaves the SW — never the settings object (it holds
      // claudeApiKey). See the storage.onChanged broadcast near the top.
      const settings = await getSettings()
      return { ok: true, data: { show: owaFabShown(settings) } }
    }

    // ---- Production handlers (used by popup #6) ----------------------------

    case 'getStatus': {
      // Counts come from the onChanged-invalidated cache (see getStatusCounts)
      // instead of re-reading every rule + the whole folder tree per call.
      const [settings, counts] = await Promise.all([getSettings(), getStatusCounts()])
      const owa = await pingOwa()
      // Don't block popup-open on a fresh token fetch — askContentScript's
      // cold-start retry can wait ~1.5s when OWA's content script hasn't
      // injected yet. Peek the cache instead. If stale, kick off a
      // background refresh so the cache is warm by the time the user
      // actually clicks 開始歸類 (which awaits getOwaToken anyway).
      let tokenValid = false
      if (owa.ok) {
        const peek = await peekCachedToken()
        tokenValid = peek.valid
        if (!peek.valid) {
          void getOwaToken().catch(() => {
            /* background refresh — its failure is recorded by token.ts (lastTokenError) */
          })
        }
      }
      // The REAL reason the token is unavailable (2026-11-05): before, the
      // popup could only say 「稍等幾秒再試」 — the multi-account guard's
      // message never reached the user, nor "reload Outlook after an extension
      // update". The background refresh above records its failure; the popup
      // re-polls every second, so a fresh reason shows up within a poll or two.
      // Ignore stale entries so an hour-old failure can't mislead.
      let tokenError: { code: string; message: string } | null = null
      if (owa.ok && !tokenValid) {
        const last = await getLastTokenError().catch(() => null)
        if (last && Date.now() - last.at < 5 * 60_000) {
          tokenError = { code: last.code, message: last.message }
        }
      }
      return {
        ok: true,
        data: {
          owaConnected: owa.ok,
          tokenValid,
          tokenError,
          apiKeyConfigured: settings.claudeApiKey.length > 0,
          apiKeyPreview: settings.claudeApiKey.length > 0 ? maskApiKey(settings.claudeApiKey) : null,
          model: settings.claudeModel,
          folderCacheAgeMs:
            counts.folderCacheUpdatedAt !== null
              ? Date.now() - new Date(counts.folderCacheUpdatedAt).getTime()
              : null,
          folderCacheCount: counts.folderCacheCount,
          rulesCount: counts.rulesCount,
          excludePrefixes: settings.excludeFolderPrefixes,
          closedCasePrefixes: settings.closedCasePrefixes ?? [],
          batchSize: settings.batchSize,
          aiConfidenceThreshold: settings.aiConfidenceThreshold,
          skipFlagged: settings.skipFlagged,
          showOwaFab: settings.showOwaFab,
          prefetchNextBatch: settings.prefetchNextBatch,
          recentActivityIncludePrefixes: settings.recentActivityIncludePrefixes,
          recentActivityIncludeLeafNames: settings.recentActivityIncludeLeafNames,
          internalDomains: settings.internalDomains,
          primaryRootPath: settings.primaryRootPath,
          internalSubjectCategories: settings.internalSubjectCategories,
          aiIncludeFewShotExamples: settings.aiIncludeFewShotExamples,
          harvestEnabled: settings.harvestEnabled === true,
          syncEnabled: settings.syncEnabled,
          syncMachineId: settings.syncMachineId,
          lastSyncAt: settings.lastSyncAt,
        },
      }
    }

    case 'openOwa': {
      // Sent by the TOOLBAR popup when getStatus says no Outlook tab is open
      // (2026-11-05) — it then closes itself. Replaces the old setPopup('')
      // swapping driven by browser-wide tab listeners (see action-router.ts).
      const action = await openOrFocusOwa()
      return { ok: true, data: { action } }
    }

    case 'getFolderActivity': {
      // Used by IdleScreen 「近日活動」panel. SW-side filter using the
      // user's allowlist (prefixes / leaf-names) so popup doesn't keep
      // its own filter state in sync with settings.
      //
      // Filter contract is in `shared/folder-activity-filter.ts` so it
      // can be unit-tested. Critical case there: empty allowlists →
      // return all rows (was previously "return nothing", which broke
      // fresh installs after generification removed the hardcoded
      // prefix defaults).
      //
      // Each row is annotated with `localExists` — true when its
      // folderId is present in the current folderCache. Cross-machine
      // sync (folderActivity v2) may bring in entries whose folderId
      // refers to a folder that doesn't exist on THIS machine (e.g.
      // user used to be on a different Microsoft account here). Popup
      // greys those out and disables the click — better than a 404 on
      // navigation.
      const [activity, settings, cache] = await Promise.all([
        getFolderActivity(),
        getSettings(),
        getFolderCache(),
      ])
      const filtered = filterFolderActivity(
        activity,
        settings.recentActivityIncludePrefixes,
        new Set(settings.recentActivityIncludeLeafNames),
      )
      const localIds = cache
        ? new Set(flattenFolderTree(cache.tree).map((n) => n.id))
        : null
      // localExists is undefined when we have NO cache at all (first run
      // before classify) — popup interprets that as "unknown, render
      // normally". Only the explicit `false` triggers the degraded UI.
      const annotated = filtered.map((row) => ({
        ...row,
        localExists: localIds ? localIds.has(row.folderId) : undefined,
      }))
      return { ok: true, data: { activity: annotated } }
    }

    case 'getWeeklyDigest': {
      // First-show seeding: store a baseline so next time we can compute
      // a meaningful delta. Don't show a digest on the very first popup
      // open (nothing to summarize yet).
      const state = await getWeeklyDigestState()
      const currentMetrics = await getMetrics()
      const currentRules = await listRules()
      if (!state) {
        await setWeeklyDigestState({
          lastShownAt: new Date().toISOString(),
          snapshot: { metrics: currentMetrics, rulesCount: currentRules.length },
        })
        return { ok: true, data: { shouldShow: false } }
      }
      const lastMs = new Date(state.lastShownAt).getTime()
      const daysSince = (Date.now() - lastMs) / (1000 * 60 * 60 * 24)
      if (daysSince < 7) {
        return { ok: true, data: { shouldShow: false } }
      }
      const health = computeRuleHealth(currentRules)
      return {
        ok: true,
        data: {
          shouldShow: true,
          daysSpan: Math.round(daysSince),
          moved: Math.max(0, currentMetrics.moved - state.snapshot.metrics.moved),
          deleted: Math.max(0, currentMetrics.deleted - state.snapshot.metrics.deleted),
          foldersCreated: Math.max(
            0,
            currentMetrics.foldersCreated - state.snapshot.metrics.foldersCreated,
          ),
          errors: Math.max(0, currentMetrics.errors - state.snapshot.metrics.errors),
          rulesAdded: Math.max(0, currentRules.length - state.snapshot.rulesCount),
          rulesCountNow: currentRules.length,
          sleepingCount: health.counts.sleeping,
          orphanedCount: health.counts.orphaned,
          conflictsCount: health.counts.conflicts,
        },
      }
    }

    case 'dismissWeeklyDigest': {
      // Snapshot current state so the next digest reports the delta from
      // here. Resets the 7-day timer.
      const currentMetrics = await getMetrics()
      const currentRules = await listRules()
      await setWeeklyDigestState({
        lastShownAt: new Date().toISOString(),
        snapshot: { metrics: currentMetrics, rulesCount: currentRules.length },
      })
      return { ok: true }
    }

    case 'navigateToFolder': {
      // Click on a recent-activity row opens the folder in OWA. Empirically
      // verified URL pattern (no trailing slash):
      //
      //   https://{owa-host}/mail/{url-encoded-folder-id}
      //
      // The folder ID is the same base64-with-/= identifier we already get
      // from Graph API (`/me/mailFolders` → folder.Id) — encodeURIComponent
      // takes care of escaping `/` to %2F and `=` to %3D so the path is
      // legal. CRUCIAL: use the SAME host as the existing OWA tab — Microsoft
      // has three (office.com / office365.com / cloud.microsoft) and cross-
      // host navigation makes OWA bail to the home view, not the folder.
      const folderId = typeof msg.folderId === 'string' ? msg.folderId : ''
      if (!folderId) {
        return { ok: false, code: 'NO_FOLDER_ID', message: '未提供資料夾 ID' }
      }
      const owaTabs = await chrome.tabs.query({
        url: [
          'https://outlook.office.com/*',
          'https://outlook.office365.com/*',
          'https://outlook.cloud.microsoft/*',
        ],
      })
      // Default host for the "no OWA tab open" fallback. office.com is the
      // most-installed of the three.
      let host = 'https://outlook.office.com'
      if (owaTabs.length > 0 && owaTabs[0]!.url) {
        try {
          host = new URL(owaTabs[0]!.url).origin
        } catch {
          // malformed url, keep the fallback host
        }
      }
      const targetUrl = `${host}/mail/${encodeURIComponent(folderId)}`
      if (owaTabs.length > 0) {
        const tab = owaTabs[0]!
        await chrome.tabs.update(tab.id!, { active: true, url: targetUrl })
        if (tab.windowId !== undefined) {
          await chrome.windows.update(tab.windowId, { focused: true })
        }
      } else {
        await chrome.tabs.create({ url: targetUrl, active: true })
      }
      return { ok: true }
    }

    case 'refreshFolderActivity': {
      // Single-flight (2026-11-05): popup + in-OWA panel both auto-refresh on
      // open; a second caller shares the in-flight ~200-request scan instead
      // of starting another. Check + set happen synchronously — no await
      // between them (see refreshFolderActivityInFlight).
      if (refreshFolderActivityInFlight) return refreshFolderActivityInFlight
      const run = scanFolderActivity()
      refreshFolderActivityInFlight = run
      try {
        return await run
      } finally {
        if (refreshFolderActivityInFlight === run) refreshFolderActivityInFlight = null
      }
    }

    case 'harvestManualFilings': {
      // Manual-filing harvester (opt-in). Finds mail the lawyer filed by hand
      // in Outlook (a move rewrites LastModifiedDateTime, so recently-MOVED —
      // not recently-received — mail surfaces), filters out our own moves /
      // rule-covered / already-known routes, and feeds what's left into
      // thread-memory as HARVESTED (soft hint, never auto-route, never a
      // rule). Decoupled from refreshFolderActivity (quick-jump).
      //
      // 2026-11-05: ONE mailbox-wide `LastModifiedDateTime gt W` query
      // (paged) replaces the per-leaf `$top=10` probes; see harvest.ts for
      // the window / watermark rules.
      if (harvestInFlight) {
        return { ok: false, code: 'ALREADY_RUNNING', message: '手動歸檔掃描已在進行中' }
      }
      harvestInFlight = true
      try {
        // Watermark candidate = scan START (minus margin, see harvest.ts):
        // anything filed while this run pages through the mailbox is newer
        // than it and gets picked up next time.
        const scanStartedAt = Date.now()
        const settings = await getSettings()
        if (!settings.harvestEnabled) {
          return { ok: true, data: { scannedFolders: 0, candidates: 0, disabled: true } }
        }
        const cache = await getFolderCache()
        if (!cache) return { ok: false, code: 'NO_CACHE', message: '尚未抓過資料夾、請先跑一次歸類' }
        const tok = await getOwaToken()
        if (!tok.ok) return { ok: false, code: 'NO_TOKEN', message: '無法取得 OWA token' }
        const rootPath = settings.primaryRootPath
        if (!rootPath) {
          return { ok: false, code: 'NO_ROOT_CONFIGURED', message: '尚未設定主要根資料夾' }
        }
        const root = flattenFolderTree(cache.tree).find((f) => f.path === rootPath)
        if (!root) return { ok: false, code: 'NO_ROOT', message: `找不到「${rootPath}」根目錄` }
        const targets: MailFolderNode[] = []
        const collectLeaves = (nodes: MailFolderNode[]): void => {
          for (const n of nodes) {
            if (n.children.length === 0) targets.push(n)
            else collectLeaves(n.children)
          }
        }
        collectLeaves(root.children)
        const excludes = settings.excludeFolderPrefixes
        // No folder cap any more: the target set is a client-side id filter
        // on one query, so it costs no requests.
        const targetFolders = new Map(
          targets
            .filter((n) => !excludes.some((p) => n.path.startsWith(p)))
            .map((n) => [n.id, n.path] as const),
        )
        const since = harvestWindowStart(await getLastHarvestAt(), scanStartedAt)
        const listing = await api.listMessagesModifiedSince(since, {
          pageSize: HARVEST_PAGE_SIZE,
          maxPages: HARVEST_MAX_PAGES,
        })
        const scanned: HarvestScanItem[] = selectHarvestScanItems(listing.messages, targetFolders)
        const scannedFolders = targetFolders.size
        const errored = listing.complete ? 0 : 1
        const [rules, movedByUs, convMemory, subjectMemory] = await Promise.all([
          listRules(),
          getMovedByUsLedger(),
          getConversationMemory(),
          getSubjectMemory(),
        ])
        const candidates = filterHarvestCandidates(scanned, {
          movedByUs,
          ruleIndex: buildRuleIndex(rules, { ownDomains: await resolveOwnDomains({ fetch: false }) }),
          convMemory,
          subjectMemory,
        })
        const convFilings = candidates
          .filter((c) => c.conversationId)
          .map((c) => ({ convId: c.conversationId!, folderId: c.folderId, folderPath: c.folderPath }))
        const subjectFilings = candidates
          .filter((c) => c.normalizedSubject)
          .map((c) => ({
            normalizedSubject: c.normalizedSubject!,
            folderId: c.folderId,
            folderPath: c.folderPath,
          }))
        if (convFilings.length > 0) await recordConversationFilings(convFilings, { harvested: true })
        if (subjectFilings.length > 0) await recordSubjectFilings(subjectFilings, { harvested: true })
        // Advance the watermark ONLY after a COMPLETE paging run, and only to
        // the scan START minus the margin (never the end-of-run clock): a
        // page-capped or failed run leaves it untouched so the unseen part of
        // the window is retried, and filings made during the run stay newer
        // than the watermark.
        if (listing.complete && targetFolders.size > 0) {
          await setLastHarvestAt(nextHarvestWatermark(scanStartedAt))
        }
        return {
          ok: true,
          data: {
            scannedFolders,
            errored,
            candidates: candidates.length,
            scannedMessages: listing.messages.length,
            complete: listing.complete,
          },
        }
      } catch (e) {
        return {
          ok: false,
          code: 'HARVEST_ERROR',
          message: e instanceof Error ? e.message : String(e),
        }
      } finally {
        harvestInFlight = false
      }
    }

    case 'getOnboardingState': {
      // First-run wizard decision logic. The popup needs to distinguish
      // THREE cases on a fresh machine:
      //
      //   1. Genuinely new user (no API key, no rules, no sync, no
      //      cloud data) — show the full setup wizard.
      //
      //   2. "Already used on another machine" — cloud has data from
      //      a different machineId. We DON'T want to push them into
      //      "set up from scratch", which would create new rules that
      //      conflict with the eventual sync pull. Offer to enable
      //      sync + pull first.
      //
      //   3. Reinstalled with backups intact — chrome.storage.local
      //      still has syncBackups from previous session. Offer
      //      restoreSyncBackup as an option.
      //
      // Each signal is independent so the UI can compose them.
      const [settings, rules, syncStatus, backups] = await Promise.all([
        getSettings(),
        listRules(),
        getSyncStatus(),
        listBackups(),
      ])
      // "Needs onboarding" = no API key AND user hasn't dismissed.
      //
      // Only the API key is genuinely required for the first classify
      // to work — primaryRootPath / internalDomains / etc. are
      // optional config that the user can tune in Options later, the
      // pipeline already scans the whole folder tree by default.
      //
      // Rule count is NOT a wizard trigger either: a user who has the
      // API key but no rules yet is in a normal "haven't classified
      // anything yet" state, not a wizard-needed state.
      const hasApiKey = settings.claudeApiKey.length > 0
      const hasRules = rules.length > 0
      const hasPrimaryRoot = settings.primaryRootPath.length > 0
      const dismissed = settings.onboardingDismissed
      const needed = !dismissed && !hasApiKey
      // Cloud has another machine's data?
      const cloudHasOtherMachineData =
        !!syncStatus.cloud && !syncStatus.cloud.isUs
      return {
        ok: true,
        data: {
          needed,
          // Why the wizard might / might-not show, exposed for UI clarity.
          reasons: {
            hasApiKey,
            hasRules,
            hasPrimaryRoot,
            dismissed,
          },
          cloud: cloudHasOtherMachineData
            ? {
                ruleCount: syncStatus.cloud!.ruleCount,
                tombstoneCount: syncStatus.cloud!.tombstoneCount,
                updatedAt: syncStatus.cloud!.updatedAt,
                sourceMachineId: syncStatus.cloud!.sourceMachineId,
              }
            : undefined,
          syncEnabled: syncStatus.enabled,
          hasLocalBackups: backups.length > 0,
          backupCount: backups.length,
        },
      }
    }

    case 'getFolderActivityFreshness': {
      // Used by popup IdleScreen to decide whether to auto-trigger
      // refreshFolderActivity on open. Returns the timestamp of the
      // last successful refresh (Graph API scan) — null if never run.
      // Popup compares against `Date.now()` and triggers if > 30 min.
      const at = await getFolderActivityRefreshAt()
      return { ok: true, data: { lastRefreshAt: at } }
    }

    case 'detectUserDomain': {
      // Used by the FirmSettingsCard's 「自動偵測」 button. Single API
      // call to /me, returns the user's primary email + its domain. The
      // popup / options page then uses the domain as the suggested
      // internalDomain entry. No persistence here — caller decides
      // whether to apply.
      try {
        const email = await api.getMyEmail()
        if (!email) {
          return { ok: false, code: 'NO_EMAIL', message: 'Outlook 未回傳 email' }
        }
        const at = email.lastIndexOf('@')
        if (at < 0 || at === email.length - 1) {
          return {
            ok: false,
            code: 'INVALID_EMAIL',
            message: `Outlook 回傳的 email 格式異常:${email}`,
          }
        }
        return {
          ok: true,
          data: { email, domain: email.slice(at + 1).toLowerCase().trim() },
        }
      } catch (e) {
        return {
          ok: false,
          code: 'DETECT_FAILED',
          message: e instanceof Error ? e.message : String(e),
        }
      }
    }

    case 'peekNextBatch': {
      // Lightweight "would the next batch have anything?" probe used by
      // the DoneScreen to decide whether to show 「繼續歸檔下一批」.
      //
      // Reuses preflight's pre-filters (skip-history + skipFlagged) but
      // skips folder tree, rule reconcile, rule matching, and AI — all
      // we need is a count. One Graph API call with the narrowest
      // possible select (Id + Flag) keeps this under ~200ms.
      //
      // Returns:
      //   - eligibleCount:    number of emails that would enter the next batch
      //   - cappedAtBatchSize: true when the fetch hit the batchSize ceiling
      //                       (could be more emails beyond this page)
      //
      // Non-fatal — failure returns { ok:false } and the popup falls back to
      // the old always-show-button behaviour.
      try {
        const settings = await getSettings()
        const batchSize = clampBatchSize(msg.batchSize, settings.batchSize)
        // Same scan as classifyPreflight (ledgers, ghost over-fetch, paging
        // past kept / flagged mail) so the peek number predicts what classify
        // will actually act on — a peek of "0 封可處理" while older mail sat
        // below a page of flagged mail was the starvation bug's UI face.
        const [skipHistory, recentlyProcessed] = await Promise.all([
          getSkipHistory(),
          getRecentlyProcessedIds(),
        ])
        const scan = await scanInboxForBatch({
          batchSize,
          // Id for skip-history match, Flag for the flagged filter — that's all.
          select: 'Id,Flag',
          skipIds: new Set(Object.keys(skipHistory)),
          recentlyProcessed,
        })
        return {
          ok: true,
          data: {
            // One classify run handles at most batchSize — cap the peek
            // number to what the next batch will really take on.
            eligibleCount: Math.min(scan.eligible.length, batchSize),
            cappedAtBatchSize: scan.moreMayExist,
            totalFetched: scan.scanned,
          },
        }
      } catch (e) {
        return {
          ok: false,
          code: 'PEEK_FAILED',
          message: e instanceof Error ? e.message : String(e),
        }
      }
    }

    case 'classifyPreflight': {
      // Phase 1 — fetch tree + inbox, run rules, hand back unmatched preview.
      // Popup shows preview screen for user to opt-out specific emails before
      // any tokens are spent. Full Email objects stashed in session storage so
      // Phase 2 doesn't need to re-fetch.
      //
      // A preflight starts a NEW run (2026-11-05): claim it synchronously,
      // before any await, so the previous run's loop sees it lost ownership
      // at its very next check and its in-flight Claude call is aborted now.
      const runId = newClassifyRunId()
      const claimRun = setClassifyRun({ runId, createdAt: Date.now(), executedIds: [] })
      abortClassifyLoopsExcept(runId)
      await claimRun

      const settings = await getSettings()
      if (!settings.claudeApiKey) {
        return { ok: false, code: 'NO_API_KEY', message: '請先到設定頁設定 Claude API key' }
      }
      const batchSize = clampBatchSize(msg.batchSize, settings.batchSize)
      const forceFresh = msg.forceFresh === true

      // Wipe any leftover state from a prior run so popup mount-time resume
      // doesn't pick up stale data after this preflight starts a fresh batch.
      await Promise.all([
        chrome.storage.session.remove(AI_PROGRESS_KEY),
        chrome.storage.session.remove(POPUP_STATE_KEY),
        chrome.storage.session.remove(PREFLIGHT_CACHE_KEY),
      ])

      try {
        await setClassifyStage({ stage: 'fetching_tree' })
        const tree = await getOrFetchFolderTree(forceFresh)

        await setClassifyStage({ stage: 'fetching_inbox' })
        // Exclusion ledgers + paged scan: see scanInboxForBatch (ghost
        // over-fetch on the first page, @odata.nextLink past pages of kept /
        // flagged mail, capped).
        const [skipHistory, recentlyProcessed] = await Promise.all([
          getSkipHistory(),
          getRecentlyProcessedIds(),
        ])
        const skipIds = new Set(Object.keys(skipHistory))
        const scan = await scanInboxForBatch({ batchSize, skipIds, recentlyProcessed })

        // Reconcile rule paths against the live tree before matching: if a
        // folder was renamed/deleted in Outlook, update or mark orphan so
        // matching won't queue inevitable failures and the UI reflects truth.
        //
        // The whole reconcile (read → diff → write → audit) lives inside
        // mutateRules so it can't race against bumpRuleHits, addRules, or any
        // other concurrent rule mutator. Reconcile USED to call raw setRules
        // outside the lock — the only mutator that did — which left a narrow
        // window for clobber between, e.g., a finishing execute batch's
        // bumpRuleHits and a popup-triggered reconcile firing simultaneously.
        const reconcileEvents: RuleEvent[] = []
        let reconcileSummary: ReconcileSummary = {
          pathsUpdated: 0,
          newlyOrphaned: 0,
          unorphaned: 0,
        }
        let rules = await mutateRules<Rule[]>((current) => {
          const recon = reconcileRulesAgainstTree(current, tree)
          reconcileSummary = recon.summary
          if (
            recon.summary.pathsUpdated > 0 ||
            recon.summary.newlyOrphaned > 0 ||
            recon.summary.unorphaned > 0
          ) {
            const beforeMap = new Map(current.map((r) => [r.id, r]))
            const now = Date.now()
            for (const after of recon.rules) {
              const before = beforeMap.get(after.id)
              if (!before) continue
              const beforeSnap = snapshotOf(before)
              const afterSnap = snapshotOf(after)
              const changedFields = diffSnapshots(beforeSnap, afterSnap)
              if (changedFields.length === 0) continue
              reconcileEvents.push({
                kind: 'edit',
                ruleId: after.id,
                at: now,
                actor: 'system',
                before: beforeSnap,
                after: afterSnap,
                changedFields,
              })
            }
          }
          return { next: recon.rules, result: recon.rules }
        })
        if (reconcileEvents.length > 0) {
          await recordRuleEvents(reconcileEvents).catch((e) =>
            console.warn('[mail-organizer] reconcile audit failed (non-fatal)', e),
          )
        }
        if (
          reconcileSummary.pathsUpdated > 0 ||
          reconcileSummary.newlyOrphaned > 0 ||
          reconcileSummary.unorphaned > 0
        ) {
          console.info(
            `[mail-organizer] reconcile: ${reconcileSummary.pathsUpdated} paths updated, ` +
              `${reconcileSummary.newlyOrphaned} newly orphaned, ${reconcileSummary.unorphaned} unorphaned`,
          )
        }

        // 結案退場 Phase 1 (2026-07-16) — reconcile 剛把規則路徑按
        // folderId 刷新過，此刻的 targetFolderPath 是即時真相。搬進
        // 「結案歸檔」樹的規則在 matching 之前軟停用（同時修掉「規則
        // 默默把新信歸進已歸檔資料夾」的舊 bug），搬出者對稱復職。
        // 硬刪除只發生在 runSweep 的 Phase 2（30 天後、逐機收斂）。
        try {
          const closedPrefixes = settings.closedCasePrefixes ?? []
          const retirement = await applyClosedCaseRetirement(closedPrefixes)
          // Refresh the in-memory snapshot — `rules` above was captured from
          // the reconcile mutateRules; without this, buildRuleIndex below
          // would still see just-retired rules as enabled for this round.
          rules = retirement.rules
          if (retirement.retired > 0 || retirement.revived > 0) {
            console.info(
              `[mail-organizer] closed-case retirement: ${retirement.retired} retired, ` +
                `${retirement.revived} revived (archive prefixes: ${closedPrefixes.length})`,
            )
          }
        } catch (e) {
          console.warn('[mail-organizer] closed-case retirement failed (non-fatal)', e)
        }

        // Pre-filters (applied inside scanInboxForBatch): drop emails the user
        // already decided to keep in inbox (skip history), ones we just
        // moved/deleted (recentlyProcessed — Outlook's eventually-consistent
        // listing can still return them with now-dead ids; acting on them
        // would 404) and Outlook-flagged ("待處理") ones. The banner counts
        // only user-visible exclusions (kept + flagged); ghost drops are
        // invisible bookkeeping. (The OTHER entrance of "已歸檔的信混進下一批"
        // — resurrecting the executed batch's classify progress — is closed
        // by the run's executed-id exclusion, see recordClassifyRunExecuted.)
        const { preFilteredCount, flaggedCount } = scan
        // Final slice keeps the batch contract (≤ batchSize) after the
        // over-fetch / paging above.
        const emails = scan.eligible.slice(0, batchSize)

        await setClassifyStage({ stage: 'matching_rules', total: emails.length })
        const rulePlan: PlanItem[] = []
        const unmatched: Email[] = []
        // Thread memory pre-filter: routes follow-up emails (internal
        // replies, fwd, etc.) to the same folder we filed earlier
        // emails of the same conversation / subject — closes the gap
        // where AI sees only internal-to-internal mail with vague
        // subjects and gives up at 0.5 confidence.
        //   1. ConversationId match (exact, confidence 0.95)
        //   2. Normalized-subject match (looser, gated on
        //      conflictCount===0 and min length 8, confidence 0.85)
        const [convMemory, subjectMemory] = await Promise.all([
          getConversationMemory(),
          getSubjectMemory(),
        ])
        const folderIndex = new Map(
          flattenFolderTree(tree).map((n) => [n.id, n]),
        )
        const matchThreadMemory = (
          m: Email,
        ): { folderId: string; folderPath: string; kind: 'convId' | 'subject'; previousFolderPath: string } | null => {
          if (m.ConversationId) {
            const e = convMemory[m.ConversationId]
            // Same gating as subjectMemory: only trust the memory when
            // conflictCount===0. A convId that's been filed to two
            // different folders recently is ambiguous; let AI / rules
            // decide rather than auto-routing to the latest target.
            // conflictCount earns its way back to 0 via DECAY_AFTER_STABLE
            // consecutive same-folder filings (see recordConversationFilings).
            // `!e.harvested`: a harvested-only entry (inferred from Outlook
            // drag-drop, never plan-reviewed) is NOT trusted for auto-route —
            // it only surfaces as a soft hint below until a real filing
            // confirms it.
            if (e && (e.conflictCount ?? 0) === 0 && !e.harvested) {
              const node = folderIndex.get(e.folderId)
              if (node) {
                return {
                  folderId: node.id,
                  folderPath: node.path,
                  kind: 'convId',
                  previousFolderPath: e.folderPath,
                }
              }
            }
          }
          const norm = normalizeSubject(m.Subject ?? '')
          if (norm.length >= MIN_NORMALIZED_SUBJECT_LEN) {
            const e = subjectMemory[norm]
            if (e && (e.conflictCount ?? 0) === 0 && !e.harvested) {
              const node = folderIndex.get(e.folderId)
              if (node) {
                return {
                  folderId: node.id,
                  folderPath: node.path,
                  kind: 'subject',
                  previousFolderPath: e.folderPath,
                }
              }
            }
          }
          return null
        }
        // Soft thread hint (B2-C): for emails that DON'T get an auto-route
        // above (they reach the AI), still surface any recent-but-ambiguous
        // filing history as a weak, explicitly-uncertain nudge. An unmatched
        // email with a memory entry means the entry was too ambiguous
        // (conflictCount>0) or its folder no longer resolves — exactly the
        // case where the AI would otherwise 0.5-guess.
        const softThreadHint = (m: Email): string | null => {
          if (m.ConversationId) {
            const e = convMemory[m.ConversationId]
            if (e) {
              const path = folderIndex.get(e.folderId)?.path ?? e.folderPath
              if (path) {
                if (e.harvested && (e.conflictCount ?? 0) === 0) {
                  return `依你既有的 Outlook 歸檔，此對話近期歸於「${path}」（僅供參考）`
                }
                return (e.conflictCount ?? 0) > 0
                  ? `此對話近期在「${path}」與其他資料夾間擺盪、歸屬不穩定（僅供參考）`
                  : `此對話近期曾歸於「${path}」（僅供參考）`
              }
            }
          }
          const norm = normalizeSubject(m.Subject ?? '')
          if (norm.length >= MIN_NORMALIZED_SUBJECT_LEN) {
            const e = subjectMemory[norm]
            if (e) {
              const path = folderIndex.get(e.folderId)?.path ?? e.folderPath
              if (path) {
                return e.harvested
                  ? `依你既有的 Outlook 歸檔，同主旨郵件近期歸於「${path}」（僅供參考）`
                  : `同主旨郵件近期曾歸於「${path}」（僅供參考）`
              }
            }
          }
          return null
        }
        // emailId → hint, persisted with the preflight so classifyAi can feed
        // it to the model without re-reading thread memory.
        const threadHints: Record<string, string> = {}
        // Build the rule index ONCE per batch. Before this, matchEmail
        // re-sorted the full rules array for every email — O(N log N)
        // per email × N emails = O(N² log N) of pure sort work that the
        // indexed path replaces with a single sort plus O(1) Map lookups
        // for domain / sender (the two largest buckets in practice).
        // Own-domain addresses never trigger AUTO domain rules (2026-11-05).
        const ruleIndex = buildRuleIndex(rules, { ownDomains: await resolveOwnDomains() })
        // Rule precedence over thread memory is captured by the exported
        // `ruleBeatsThread` helper in shared/rules.ts so the contract has
        // a single source of truth and gets exercised by unit tests.
        for (const m of emails) {
          const threadHit = matchThreadMemory(m)
          const outcome = matchEmailWithIndex(m, ruleIndex)
          // Precedence (final order, after the 2026-05 fix):
          //   1. Rule that beats thread (case_code / compound / user_manual)
          //   2. Thread memory hit
          //   3. Broad rule (domain / subject_keyword / sender from auto paths)
          //   4. AI fallback (unmatched bucket)
          //
          // Why: thread memory used to short-circuit the entire rule
          // check, which shadowed user-built rules added AFTER a thread
          // was learned. With this ordering, user_manual / case_code /
          // compound rules always win, while broad auto-rules step
          // aside for thread (preserving the original use case: pull
          // internal replies into the same folder as the external
          // first-touch email).
          if (outcome && ruleBeatsThread(outcome.rule)) {
            rulePlan.push({
              emailId: m.Id,
              emailSubject: m.Subject ?? '',
              emailFrom: m.From?.EmailAddress?.Address ?? '',
              bodyPreview: m.BodyPreview ?? '',
              conversationId: m.ConversationId,
              action: 'move',
              targetFolderId: outcome.rule.targetFolderId,
              targetFolderPath: outcome.rule.targetFolderPath,
              confidence: outcome.rule.confidence,
              reason: outcome.reason,
              source: 'rule',
              ruleId: outcome.rule.id,
              originalRuleId: outcome.rule.id,
              matchedRule: {
                id: outcome.rule.id,
                type: outcome.rule.type,
                signal: outcome.rule.signal,
                source: outcome.rule.source,
                matchCount: outcome.rule.matchCount,
                lastUsedAt: outcome.rule.lastUsedAt,
                enabled: outcome.rule.enabled,
              },
            })
            continue
          }
          if (threadHit) {
            rulePlan.push({
              emailId: m.Id,
              emailSubject: m.Subject ?? '',
              emailFrom: m.From?.EmailAddress?.Address ?? '',
              bodyPreview: m.BodyPreview ?? '',
              conversationId: m.ConversationId,
              action: 'move',
              targetFolderId: threadHit.folderId,
              targetFolderPath: threadHit.folderPath,
              confidence: threadHit.kind === 'convId' ? 0.95 : 0.85,
              reason:
                threadHit.kind === 'convId'
                  ? `此對話 thread 上次已歸於「${threadHit.previousFolderPath}」`
                  : `同主旨的歷史信件上次歸於「${threadHit.previousFolderPath}」`,
              source: 'thread',
              threadMatch: {
                kind: threadHit.kind,
                previousFolderPath: threadHit.previousFolderPath,
              },
              // If a broad rule also matched and agrees on the destination,
              // remember it so execute can keep its hit-count / lastUsedAt
              // alive (優化 2026-07: thread-shadowed rules were starved →
              // stale-swept). Compare by id first, path as fallback.
              ...(outcome &&
              ((outcome.rule.targetFolderId &&
                outcome.rule.targetFolderId === threadHit.folderId) ||
                outcome.rule.targetFolderPath === threadHit.folderPath)
                ? { agreeingRuleId: outcome.rule.id }
                : {}),
            })
            continue
          }
          if (outcome) {
            rulePlan.push({
              emailId: m.Id,
              emailSubject: m.Subject ?? '',
              emailFrom: m.From?.EmailAddress?.Address ?? '',
              bodyPreview: m.BodyPreview ?? '',
              conversationId: m.ConversationId,
              action: 'move',
              targetFolderId: outcome.rule.targetFolderId,
              targetFolderPath: outcome.rule.targetFolderPath,
              confidence: outcome.rule.confidence,
              reason: outcome.reason,
              source: 'rule',
              ruleId: outcome.rule.id,
              originalRuleId: outcome.rule.id,
              matchedRule: {
                id: outcome.rule.id,
                type: outcome.rule.type,
                signal: outcome.rule.signal,
                source: outcome.rule.source,
                matchCount: outcome.rule.matchCount,
                lastUsedAt: outcome.rule.lastUsedAt,
                enabled: outcome.rule.enabled,
              },
            })
          } else {
            unmatched.push(m)
            const hint = softThreadHint(m)
            if (hint) threadHints[m.Id] = hint
          }
        }

        // Superseded or cancelled while we were fetching / matching: this
        // run is dead — don't write its cache over a newer run's.
        if (!ownsClassifyRun(runId)) {
          return {
            ok: false,
            code: 'STALE_RUN',
            message: '這次分類已被取消或被新的分類取代，請重新開始歸類',
          }
        }
        // Persist preflight context so classifyAi can run independently
        const cache: PreflightCache = {
          runId,
          unmatchedEmails: unmatched,
          tree,
          rulePlan,
          threadHints,
          createdAt: Date.now(),
        }
        await chrome.storage.session.set({ [PREFLIGHT_CACHE_KEY]: cache })

        return {
          ok: true,
          data: {
            runId,
            rulePlan,
            unmatchedPreview: unmatched.map((m) => ({
              id: m.Id,
              subject: m.Subject ?? '',
              from: m.From?.EmailAddress?.Address ?? '',
              fromName: m.From?.EmailAddress?.Name ?? '',
              received: m.ReceivedDateTime,
            })),
            folderTree: tree,
            aiConfidenceThreshold: settings.aiConfidenceThreshold,
            preFilteredCount,
            flaggedCount,
            requestedBatchSize: batchSize,
          },
        }
      } finally {
        // A newer run owns the stage display now; leave it alone.
        if (ownsClassifyRun(runId)) await setClassifyStage({ stage: 'idle' })
      }
    }

    case 'classifyAi': {
      // Phase 2 — chunked, progressive AI classification.
      // Fire-and-forget: SW writes progress to session storage each chunk,
      // popup polls for incremental updates.
      const settings = await getSettings()
      const excludeIds = new Set((msg.excludeIds as string[] | undefined) ?? [])
      const run = await loadClassifyRun()
      // Callers that predate run ids adopt the current run.
      const runId = typeof msg.runId === 'string' ? msg.runId : run?.runId
      const staleRun = {
        ok: false as const,
        code: 'STALE_RUN',
        message: '這次分類已被取消或被新的分類取代，請重新開始歸類',
      }
      if (!run || !runId || run.runId !== runId) return staleRun
      const pf = (await chrome.storage.session.get(PREFLIGHT_CACHE_KEY))[
        PREFLIGHT_CACHE_KEY
      ] as PreflightCache | undefined
      // One loop per run (items 1 + 3): a double-submitted 繼續歸檔 / a
      // prefetch racing the foreground must never classify the same run twice.
      const alreadyStarted = async () => {
        const prog = (await chrome.storage.session.get(AI_PROGRESS_KEY))[AI_PROGRESS_KEY] as
          | StoredAiProgress
          | undefined
        return {
          ok: true as const,
          data: {
            started: false,
            alreadyStarted: true,
            runId,
            chunks: prog?.runId === runId ? prog.chunks : 0,
            totalEmails: prog?.runId === runId ? prog.totalEmails : 0,
          },
        }
      }
      if (!ownsClassifyRun(runId)) return staleRun
      if (classifyRun!.aiStartedAt != null || classifyLoops.has(runId)) return alreadyStarted()
      if (!pf || pf.runId !== runId) {
        return { ok: false, code: 'NO_PREFLIGHT', message: '請先重新分類（preflight 過期）' }
      }

      const toClassify: Email[] = []
      const toSkip: Email[] = []
      for (const m of pf.unmatchedEmails) {
        if (excludeIds.has(m.Id)) toSkip.push(m)
        else toClassify.push(m)
      }

      // Smart batching: group by Outlook ConversationId so Claude only
      // classifies one representative per thread. Siblings inherit the
      // decision after the rep returns. Saves ~30% of AI calls on a
      // typical inbox where back-and-forth threads dominate. Emails
      // without a ConversationId become their own one-member group.
      const { reps, siblingsByRepId } = groupEmailsByConversation(toClassify)

      const chunks: Email[][] = []
      for (let i = 0; i < reps.length; i += AI_CHUNK_SIZE) {
        chunks.push(reps.slice(i, i + AI_CHUNK_SIZE))
      }

      // Immutable cycle birth time — informational now (the popup shows it);
      // which cycle a progress object belongs to is decided by runId.
      const classifyStartedAt = Date.now()

      // Claim the run synchronously — no await between these checks and the
      // claim — so a concurrent duplicate classifyAi is refused above.
      if (!ownsClassifyRun(runId)) return staleRun
      if (classifyRun!.aiStartedAt != null || classifyLoops.has(runId)) return alreadyStarted()
      const controller = new AbortController()
      classifyLoops.set(runId, controller)
      // Any older loop still unwinding (superseded) gets aborted for good.
      abortClassifyLoopsExcept(runId)
      const claimWrite = setClassifyRun({ ...classifyRun!, aiStartedAt: Date.now() })

      // The loop's output, append-only. Persisted minus executed ids, with
      // each kept item's append position as the poll cursor (see
      // StoredAiProgress).
      let aiPlan: PlanItem[] = []
      /**
       * Write progress iff this run is still current. Ownership check and the
       * storage call happen synchronously together, so a cancel that ran
       * before can't be undone by this write and a cancel after it removes
       * it. Returns null (and writes nothing) once the run is gone.
       */
      const writeProgress = (
        fields: Pick<StoredAiProgress, 'completedChunks' | 'done'> &
          Partial<Pick<StoredAiProgress, 'skippedByUser' | 'usage' | 'aiError'>>,
        planOut: PlanItem[] = aiPlan,
      ): Promise<void> | null => {
        if (!ownsClassifyRun(runId)) return null
        const executed = new Set(classifyRun!.executedIds)
        const keptAi: PlanItem[] = []
        const keptSeq: number[] = []
        planOut.forEach((item, k) => {
          if (executed.has(item.emailId)) return
          keptAi.push(item)
          keptSeq.push(k)
        })
        const progress: StoredAiProgress = {
          runId,
          totalEmails: toClassify.length,
          // Classification work done, executed items included — keeps the
          // progress bar monotonic across a partial execute.
          completedEmails: aiPlan.length,
          chunks: chunks.length,
          rulePlan: pf.rulePlan.filter((i) => !executed.has(i.emailId)),
          folderTree: pf.tree,
          aiPlan: keptAi,
          aiSeq: keptSeq,
          aiTotal: planOut.length,
          startedAt: classifyStartedAt,
          ...fields,
        }
        return chrome.storage.session.set({ [AI_PROGRESS_KEY]: progress })
      }

      // Initial progress snapshot (rulePlan + folderTree embedded so popup can
      // resume from this state alone if it reopens mid-run)
      const initialWrite = writeProgress({ completedChunks: 0, done: chunks.length === 0 })
      await claimWrite
      await initialWrite

      // Run the AI loop in the background; popup polls.
      void (async () => {
        const totalUsage: ClassifierUsage = {
          inputTokens: 0,
          outputTokens: 0,
          cacheCreationTokens: 0,
          cacheReadTokens: 0,
        }
        const addUsage = (u: ClassifierUsage) => {
          totalUsage.inputTokens += u.inputTokens
          totalUsage.outputTokens += u.outputTokens
          totalUsage.cacheCreationTokens += u.cacheCreationTokens
          totalUsage.cacheReadTokens += u.cacheReadTokens
        }
        let aiError: { code: string; message: string } | null = null
        const threshold = settings.aiConfidenceThreshold
        let gatedCount = 0
        let truncatedCount = 0
        const ownsRun = () => ownsClassifyRun(runId)

        await holdKeepAlive()
        try {
          if (ownsRun()) {
            await setClassifyStage({ stage: 'calling_ai', toClassify: toClassify.length })
          }

          // Pull current rules once per classify run so few-shot examples
          // reflect the user's latest state. Read fresh — preflight cache
          // might be from minutes ago and rules can change (manual edits,
          // ai_overridden generated by a previous batch).
          const currentRules = await listRules()
          // Rule index for the batch-3 full-body re-match. These reps missed
          // every rule at preflight (subject/domain/thread), but preflight had
          // no body — now that we fetch the full body per chunk, an unambiguous
          // body case number can route a rep by rule and skip the AI call.
          // Own-domain addresses never trigger AUTO domain rules (2026-11-05).
          const aiRuleIndex = buildRuleIndex(currentRules, {
            ownDomains: await resolveOwnDomains({ fetch: false }),
          })
          // Recently-active folders (B2-D): most-recent-first paths of folders
          // this tool has been filing into. Biases the AI away from a case's
          // long-closed folder. Read once per run; churns slowly relative to a
          // single classify. Sorted desc by lastActiveAt; classifier caps +
          // exclude-filters. Fail-soft: activity is a nicety, never block AI.
          let recentFolders: string[] = []
          try {
            const activity = await getFolderActivity()
            recentFolders = [...activity]
              .sort((a, b) => (b.lastActiveAt ?? '').localeCompare(a.lastActiveAt ?? ''))
              .map((a) => a.folderPath)
              .filter(Boolean)
          } catch (e) {
            console.warn('[mail-organizer] getFolderActivity for AI hint failed', e)
          }

          // Chunks run up to AI_PARALLEL_CHUNKS at a time (2026-11-05, 放寬並加速);
          // a 500-email batch was 20 strictly sequential Claude calls. Each chunk
          // resolves to an outcome (never rejects); outcomes merge into aiPlan
          // strictly in chunk order, so the plan order and the progress bar
          // (completedChunks = contiguous merged chunks) behave as before. Any
          // 429 / 529 drops the run to one call at a time. A run-fatal error
          // stops scheduling; chunks already in flight finish and keep their
          // results; whatever never completed becomes `unclassified` below.
          type ChunkOutcome =
            | { kind: 'done'; items: PlanItem[] }
            | { kind: 'fatal'; items: PlanItem[]; error: unknown; retriedMs: number }
            | { kind: 'aborted' }
          let throttled = false
          // Cold-cache gate: one request goes first; the rest wait until its
          // response begins streaming, the point where its prompt-cache write
          // (system + folder tree + rules, identical for every chunk of this
          // run) becomes readable. Without it the first wave would each pay a
          // 1h-TTL cache write for the same prefix and read none of it.
          let warm = false
          let firstChunkLocalError: { code: string; message: string } | null = null

          const processChunk = async (ci: number): Promise<ChunkOutcome> => {
            const chunk = chunks[ci]!
            // Full-body fetch (batch-3): ONLY the conversation reps that
            // actually reach Claude get their UniqueBody pulled, so body case
            // numbers surface in the prompt (B3-C4) and can sink into learning
            // (B3-C6). Rule/thread-routed mail and non-rep siblings are never
            // fetched — this single subset gate is what bounds the extra Graph
            // round-trips + token cost. Fail-soft: on error the email keeps
            // only its 250-char BodyPreview. The undefined-guard makes SW-
            // restart recovery re-runs idempotent.
            const bodyFailures: string[] = []
            await Promise.all(
              chunk.map(async (m) => {
                if (m.bodyText !== undefined) return
                try {
                  m.bodyText = await api.getMessageBody(m.Id, controller.signal)
                } catch (e) {
                  if (controller.signal.aborted) return
                  bodyFailures.push(
                    e instanceof OutlookError
                      ? `HTTP_${e.status}`
                      : e instanceof Error
                        ? e.name
                        : 'error',
                  )
                  console.warn('[mail-organizer] getMessageBody failed (non-fatal)', m.Id, e)
                }
              }),
            )
            if (!ownsRun()) return { kind: 'aborted' }
            if (bodyFailures.length > 0) {
              // Still fail-soft (the preview is classified instead), but no
              // longer console-only (item 15): one entry per chunk, not per
              // email, so an Outlook outage can't flood the 200-entry log.
              await logError(
                'classify:body',
                `正文抓取失敗 ${bodyFailures.length}/${chunk.length} 封，改用 250 字預覽分類`,
                {
                  chunk: ci + 1,
                  chunks: chunks.length,
                  failed: bodyFailures.length,
                  codes: [...new Set(bodyFailures)].slice(0, 5),
                },
              )
            }
            // Full-body re-match (batch-3 C5 activation): now that the body is
            // fetched, an UNAMBIGUOUS body case number can route a rep by rule —
            // skipping the AI. Only reps that still don't match go to Claude.
            // Body-routed reps (+ their siblings) join the plan as source:'rule'
            // and get hit-counted by execute like any rule match, so they never
            // stale-sweep (closes the P3 window-churn). Ambiguity + subject-
            // first + structural-only gating lives in matchBodyCaseWithIndex.
            const aiReps: Email[] = []
            const bodyRouted: PlanItem[] = []
            for (const m of chunk) {
              const bodyOutcome = matchBodyCaseWithIndex(m, aiRuleIndex)
              if (!bodyOutcome) {
                aiReps.push(m)
                continue
              }
              const repItem: PlanItem = {
                emailId: m.Id,
                emailSubject: m.Subject ?? '',
                emailFrom: m.From?.EmailAddress?.Address ?? '',
                bodyPreview: m.BodyPreview ?? '',
                conversationId: m.ConversationId,
                action: 'move',
                targetFolderId: bodyOutcome.rule.targetFolderId,
                targetFolderPath: bodyOutcome.rule.targetFolderPath,
                confidence: bodyOutcome.rule.confidence,
                reason: bodyOutcome.reason,
                source: 'rule',
                ruleId: bodyOutcome.rule.id,
                originalRuleId: bodyOutcome.rule.id,
                matchedRule: {
                  id: bodyOutcome.rule.id,
                  type: bodyOutcome.rule.type,
                  signal: bodyOutcome.rule.signal,
                  source: bodyOutcome.rule.source,
                  matchCount: bodyOutcome.rule.matchCount,
                  lastUsedAt: bodyOutcome.rule.lastUsedAt,
                  enabled: bodyOutcome.rule.enabled,
                },
              }
              bodyRouted.push(repItem)
              const sibs = siblingsByRepId.get(m.Id)
              if (sibs) {
                for (const sib of sibs) bodyRouted.push(inheritPlanItemForSibling(repItem, sib))
              }
            }
            // Ownership re-check right before spending tokens (item 1).
            if (!ownsRun()) return { kind: 'aborted' }
            let retriedMs = 0
            try {
              // Only the still-unmatched reps hit Claude; classifyBatch returns
              // an empty result for aiReps=[] (no API call, no cost).
              const result = await withClaudeRetry(
                () =>
                  classifyBatch(
                    {
                      emails: aiReps,
                      folderTree: pf.tree,
                      excludePrefixes: settings.excludeFolderPrefixes,
                      rules: currentRules,
                      threadHints: pf.threadHints,
                      recentFolders,
                    },
                    settings,
                    { signal: controller.signal, onResponseStart: () => markWarm() },
                  ),
                {
                  signal: controller.signal,
                  onRetry: ({ attempt, waitMs, error }) => {
                    retriedMs += waitMs
                    // Rate-limited / overloaded: stop fanning out for the rest
                    // of the run (the in-flight calls finish on their own).
                    if (error.code === 'HTTP_429' || error.code === 'HTTP_529') throttled = true
                    console.warn(
                      `[mail-organizer] Claude 暫時性錯誤 ${error.code}，` +
                        `${Math.round(waitMs / 1000)}s 後重試 (#${attempt})`,
                    )
                  },
                },
              )
              if (result.truncated) {
                // Audit: floor at 0. parsedCount is the PRE-dedup raw action
                // count, so a max_tokens response whose salvaged prefix
                // repeats an emailIndex (a documented model glitch) can make
                // this term negative — and a negative term would cancel a
                // sibling chunk's real positive, silently suppressing the
                // "部分 chunk 觸頂截斷" warning even though emails WERE cut.
                truncatedCount += Math.max(0, aiReps.length - result.truncated.parsedCount)
              }
              const chunkPlan = result.plan.map((p) => {
                if (p.action !== 'skip' && p.source === 'ai' && p.confidence < threshold) {
                  gatedCount++
                  // Preserve aiOriginal* so that downstream learning
                  // (generateAiConfirmedRules / generateAiOverrideRules)
                  // can still tell whether the user's eventual decision
                  // matched or overrode the AI's low-confidence guess.
                  // Without this, the user resolving the item would
                  // produce no rule at all — wasted intent signal.
                  return {
                    ...p,
                    action: 'skip' as const,
                    targetFolderId: undefined,
                    targetFolderPath: undefined,
                    suggestedFolderName: undefined,
                    suggestedParentPath: undefined,
                    source: 'unresolved' as const,
                    reason: `信心 ${p.confidence.toFixed(2)} < ${threshold} (自動保留)`,
                  }
                }
                return p
              })
              // Fan out each rep's decision to its thread siblings. After
              // this step the plan length matches the email count the user
              // expects (not the rep count).
              const fannedOut: PlanItem[] = []
              for (const repItem of chunkPlan) {
                fannedOut.push(repItem)
                const sibs = siblingsByRepId.get(repItem.emailId)
                if (sibs) {
                  for (const sib of sibs) fannedOut.push(inheritPlanItemForSibling(repItem, sib))
                }
              }
              addUsage(result.usage)
              return { kind: 'done', items: [...bodyRouted, ...fannedOut] }
            } catch (e) {
              // Cancelled / superseded (item 1): the run is gone — no stubs,
              // no banner, no progress write; just stop spending.
              if (
                !ownsRun() ||
                controller.signal.aborted ||
                (e instanceof ClassifierError && e.code === 'ABORTED')
              ) {
                return { kind: 'aborted' }
              }
              const ce = e instanceof ClassifierError ? e : null
              // A parse failure or a mid-stream refusal still consumed tokens.
              if (ce?.usage) addUsage(ce.usage)
              // Chunk-local vs run-fatal. PARSE_ERROR (a glitch in THIS
              // chunk's response) and REFUSAL (item 9: a safety classifier
              // declined THIS chunk's content) stub-skip just this chunk and
              // keep classifying the rest. Everything else — retries exhausted
              // (the service is still down / rate-limited), auth, invalid
              // request, model unavailable — stops the run gracefully: every
              // email that never got classified becomes an `unclassified`
              // placeholder (never recorded in skipHistory, so it comes back
              // next batch).
              const chunkLocal = ce?.code === 'PARSE_ERROR' || ce?.code === 'REFUSAL'
              if (!chunkLocal) return { kind: 'fatal', items: bodyRouted, error: e, retriedMs }
              const stubReason = ce?.code === 'REFUSAL' ? ce.message : `AI 分類失敗：CLASSIFIER_${ce!.code}`
              const stubs = unclassifiedStubs(aiReps, stubReason, new Set(bodyRouted.map((p) => p.emailId)))
              firstChunkLocalError ??= {
                code: `CLASSIFIER_${ce!.code}`,
                message:
                  ce?.code === 'REFUSAL'
                    ? `Claude 拒絕分類其中 ${stubs.length} 封（類別：${ce.category ?? '未提供'}），已標為未分類`
                    : describeClassifyFailure(e, stubs.length),
              }
              // The dispatcher's classify:api branch never sees these (the
              // loop runs detached from any message), so log here (item 4).
              await logError('classify:api', ce?.message ?? String(e), {
                code: ce!.code,
                ...(ce?.status !== undefined ? { status: ce.status } : {}),
                ...(ce?.category !== undefined ? { category: ce.category } : {}),
                model: settings.claudeModel,
                chunk: ci + 1,
                chunks: chunks.length,
                unclassified: stubs.length,
                retriedMs,
                fatal: false,
              })
              return { kind: 'done', items: [...bodyRouted, ...stubs] }
            }
          }

          /** `unclassified` placeholders for reps (+ their thread siblings) not in `already`. */
          const unclassifiedStubs = (
            reps: Email[],
            reason: string,
            already: Set<string>,
          ): PlanItem[] => {
            const stub = (m: Email): PlanItem => ({
              emailId: m.Id,
              emailSubject: m.Subject ?? '',
              emailFrom: m.From?.EmailAddress?.Address ?? '',
              bodyPreview: m.BodyPreview ?? '',
              conversationId: m.ConversationId,
              action: 'skip',
              confidence: 0,
              reason,
              source: 'unresolved',
              unclassified: true,
            })
            const out: PlanItem[] = []
            for (const m of reps) {
              if (already.has(m.Id)) continue
              out.push(stub(m))
              already.add(m.Id)
              // Siblings also need a stub or they vanish from the plan.
              for (const sib of siblingsByRepId.get(m.Id) ?? []) {
                if (already.has(sib.Id)) continue
                out.push(stub(sib))
                already.add(sib.Id)
              }
            }
            return out
          }

          const outcomes: (ChunkOutcome | undefined)[] = new Array(chunks.length)
          let merged = 0 // chunks [0, merged) are in aiPlan
          let fatalAt = -1
          let nextChunk = 0
          let inFlight = 0
          let mergeChain: Promise<void> = Promise.resolve()
          // Merge every contiguous finished chunk, in order, then persist once.
          const mergeReady = () => {
            mergeChain = mergeChain.then(async () => {
              const before = merged
              while (merged < chunks.length && outcomes[merged]?.kind === 'done') {
                aiPlan = [
                  ...aiPlan,
                  ...(outcomes[merged] as { kind: 'done'; items: PlanItem[] }).items,
                ]
                merged++
              }
              if (merged > before) {
                const w = writeProgress({ completedChunks: merged, done: false })
                if (w) await w
              }
            })
          }
          let resolveAll!: () => void
          const allDone = new Promise<void>((r) => {
            resolveAll = r
          })
          const lanes = () => (throttled || !warm ? 1 : Math.max(1, AI_PARALLEL_CHUNKS))
          const pump = () => {
            while (
              ownsRun() &&
              !controller.signal.aborted &&
              fatalAt < 0 &&
              nextChunk < chunks.length &&
              inFlight < lanes()
            ) {
              const ci = nextChunk++
              inFlight++
              void processChunk(ci)
                .catch(
                  // processChunk handles every expected failure; anything else
                  // (a bug) must end the run gracefully, not hang it.
                  (e): ChunkOutcome => ({ kind: 'fatal', items: [], error: e, retriedMs: 0 }),
                )
                .then((outcome) => {
                  outcomes[ci] = outcome
                  if (outcome.kind === 'fatal' && (fatalAt < 0 || ci < fatalAt)) fatalAt = ci
                  inFlight--
                  mergeReady()
                  pump()
                })
            }
            if (inFlight === 0) resolveAll()
          }
          function markWarm() {
            if (warm) return
            warm = true
            pump()
          }
          pump()
          await allDone
          await mergeChain

          if (!ownsRun()) return

          // Everything from the first chunk that didn't merge: completed chunks
          // keep their items (in order); a fatal chunk keeps its body-routed
          // items; anything never classified becomes an `unclassified` stub.
          if (merged < chunks.length) {
            const fatal =
              fatalAt >= 0
                ? (outcomes[fatalAt] as { kind: 'fatal'; error: unknown; retriedMs: number })
                : null
            const ce = fatal && fatal.error instanceof ClassifierError ? fatal.error : null
            const stubReason = `AI 分類失敗：${ce ? `CLASSIFIER_${ce.code}` : 'AI_FAILED'}`
            const already = new Set(aiPlan.map((p) => p.emailId))
            let stubbed = 0
            for (let ci = merged; ci < chunks.length; ci++) {
              const o = outcomes[ci]
              if (o && o.kind !== 'aborted') {
                for (const item of o.items) already.add(item.emailId)
                aiPlan = [...aiPlan, ...o.items]
              }
              if (!o || o.kind !== 'done') {
                const stubs = unclassifiedStubs(chunks[ci]!, stubReason, already)
                stubbed += stubs.length
                aiPlan = [...aiPlan, ...stubs]
              }
            }
            if (fatal) {
              aiError = {
                code: ce ? `CLASSIFIER_${ce.code}` : 'AI_FAILED',
                message: describeClassifyFailure(fatal.error, stubbed),
              }
              // The dispatcher's classify:api branch never sees these (the
              // loop runs detached from any message), so log here (item 4).
              await logError('classify:api', ce?.message ?? String(fatal.error), {
                code: ce?.code ?? 'AI_FAILED',
                ...(ce?.status !== undefined ? { status: ce.status } : {}),
                ...(ce?.category !== undefined ? { category: ce.category } : {}),
                model: settings.claudeModel,
                chunk: fatalAt + 1,
                chunks: chunks.length,
                unclassified: stubbed,
                retriedMs: fatal.retriedMs,
                fatal: true,
              })
            }
          }
          aiError ??= firstChunkLocalError

          if (!ownsRun()) return

          // Append skip stubs for emails the user opted out of
          const skipPlan: PlanItem[] = toSkip.map((m) => ({
            emailId: m.Id,
            emailSubject: m.Subject ?? '',
            emailFrom: m.From?.EmailAddress?.Address ?? '',
            bodyPreview: m.BodyPreview ?? '',
            conversationId: m.ConversationId,
            action: 'skip',
            confidence: 0,
            reason: '使用者於 preview 階段排除',
            source: 'unresolved',
          }))

          // Build banner code if multiple warnings stacked
          let banner: { code: string; message: string } | null = aiError
          if (!banner && truncatedCount > 0) {
            banner = {
              code: 'TRUNCATED',
              message: `部分 chunk 觸頂截斷、約 ${truncatedCount} 件可能不完整。建議下次縮小 batch。`,
            }
          }
          if (!banner && gatedCount > 0) {
            banner = {
              code: 'CONFIDENCE_GATED',
              message: `${gatedCount} 件信心低於 ${threshold}，已自動標為「保留」。可到設定頁調整門檻。`,
            }
          }

          const finalWrite = writeProgress(
            {
              completedChunks: chunks.length,
              skippedByUser: skipPlan.length,
              usage: totalUsage,
              aiError: banner,
              done: true,
            },
            [...aiPlan, ...skipPlan],
          )
          if (finalWrite) await finalWrite
        } finally {
          classifyLoops.delete(runId)
          if (ownsRun()) {
            // Issued together right after the ownership check: a newer run
            // can't have written its preflight cache yet (preflight claims
            // the run synchronously first), so this can only drop our own.
            await Promise.all([
              setClassifyStage({ stage: 'idle' }),
              chrome.storage.session.remove(PREFLIGHT_CACHE_KEY),
            ])
          }
          await releaseKeepAlive()
        }
      })()

      return {
        ok: true,
        data: { started: true, runId, chunks: chunks.length, totalEmails: toClassify.length },
      }
    }

    case 'getAiClassifyProgress': {
      const view: AiProgressView =
        msg.view === 'poll' || msg.view === 'meta' ? msg.view : 'full'
      const aiFrom =
        typeof msg.aiFrom === 'number' && Number.isFinite(msg.aiFrom) ? Math.max(0, msg.aiFrom) : 0
      const [run, raw] = await Promise.all([
        loadClassifyRun(),
        chrome.storage.session.get(AI_PROGRESS_KEY),
      ])
      const prog = raw[AI_PROGRESS_KEY] as StoredAiProgress | undefined
      // Only the current run's progress exists for the popup.
      if (!prog || !run || prog.runId !== run.runId) return { ok: true, data: null }
      return { ok: true, data: viewAiProgress(prog, run.executedIds, view, aiFrom) }
    }

    case 'clearAiClassifyProgress': {
      // Clearing progress also invalidates its run (item 1) — deleting the
      // key alone left the loop running and re-creating it.
      await cancelClassifyRun(typeof msg.runId === 'string' ? msg.runId : undefined)
      return { ok: true }
    }

    case 'cancelClassify': {
      // 取消 / 完成 / a superseding surface. With a runId only that run is
      // cancelled (stale cancels are no-ops); without one, the current run.
      const cancelled = await cancelClassifyRun(
        typeof msg.runId === 'string' ? msg.runId : undefined,
      )
      return { ok: true, data: { cancelled } }
    }

    case 'savePopupState': {
      // Snapshot of popup's preview/plan phase so an accidental close doesn't
      // wipe the user's in-progress edits. Cleared on idle/execute/retry.
      await chrome.storage.session.set({ [POPUP_STATE_KEY]: msg.state })
      return { ok: true }
    }

    case 'getPopupState': {
      const r = await chrome.storage.session.get(POPUP_STATE_KEY)
      return { ok: true, data: r[POPUP_STATE_KEY] ?? null }
    }

    case 'clearPopupState': {
      await chrome.storage.session.remove(POPUP_STATE_KEY)
      return { ok: true }
    }

    case 'getSkipHistoryCount': {
      const count = await getSkipHistoryCount()
      return { ok: true, data: { count } }
    }

    case 'clearSkipHistory': {
      const cleared = await clearSkipHistory()
      return { ok: true, data: { cleared } }
    }

    case 'setApiKey': {
      const key = typeof msg.key === 'string' ? msg.key.trim() : ''
      if (!key) {
        await setSettings({ claudeApiKey: '' })
        return { ok: true, data: { saved: true, cleared: true } }
      }
      if (!key.startsWith('sk-ant-')) {
        return { ok: false, code: 'INVALID_KEY', message: 'API key 應以 sk-ant- 開頭' }
      }
      await setSettings({ claudeApiKey: key })
      return { ok: true, data: { saved: true, preview: maskApiKey(key) } }
    }

    case 'setSettings': {
      const raw = (msg.patch ?? {}) as Record<string, unknown>
      // Whitelist + type-check each field. Unknown fields and bad types are
      // silently dropped — caller never gets to write garbage into storage.
      // claudeApiKey is intentionally excluded; use the setApiKey handler.
      const clean: Partial<Settings> = {}
      // Known model ids only (2026-11-05): the value syncs to every machine,
      // and an unknown / retired id fails every AI email with a 404.
      if (isKnownClaudeModel(raw.claudeModel)) {
        clean.claudeModel = raw.claudeModel
      }
      if (typeof raw.batchSize === 'number' && Number.isFinite(raw.batchSize)) {
        clean.batchSize = Math.max(BATCH_SIZE_MIN, Math.min(BATCH_SIZE_MAX, Math.floor(raw.batchSize)))
      }
      if (
        Array.isArray(raw.excludeFolderPrefixes) &&
        raw.excludeFolderPrefixes.every((s) => typeof s === 'string')
      ) {
        clean.excludeFolderPrefixes = raw.excludeFolderPrefixes
      }
      if (
        Array.isArray(raw.closedCasePrefixes) &&
        raw.closedCasePrefixes.every((s) => typeof s === 'string')
      ) {
        // Trim + drop empties — the retirement predicate does exact
        // segment-prefix matching, a stray space would silently never match.
        clean.closedCasePrefixes = raw.closedCasePrefixes
          .map((s) => s.trim().replace(/^\/+|\/+$/g, ''))
          .filter((s) => s.length > 0)
      }
      if (
        typeof raw.aiConfidenceThreshold === 'number' &&
        Number.isFinite(raw.aiConfidenceThreshold)
      ) {
        clean.aiConfidenceThreshold = Math.max(0, Math.min(1, raw.aiConfidenceThreshold))
      }
      if (typeof raw.skipFlagged === 'boolean') {
        clean.skipFlagged = raw.skipFlagged
      }
      if (typeof raw.showOwaFab === 'boolean') {
        clean.showOwaFab = raw.showOwaFab
      }
      if (typeof raw.prefetchNextBatch === 'boolean') {
        clean.prefetchNextBatch = raw.prefetchNextBatch
      }
      if (
        Array.isArray(raw.recentActivityIncludePrefixes) &&
        raw.recentActivityIncludePrefixes.every((s) => typeof s === 'string')
      ) {
        clean.recentActivityIncludePrefixes = raw.recentActivityIncludePrefixes
          .map((s) => s.trim())
          .filter((s) => s.length > 0)
      }
      if (
        Array.isArray(raw.recentActivityIncludeLeafNames) &&
        raw.recentActivityIncludeLeafNames.every((s) => typeof s === 'string')
      ) {
        clean.recentActivityIncludeLeafNames = raw.recentActivityIncludeLeafNames
          .map((s) => s.trim())
          .filter((s) => s.length > 0)
      }
      if (
        Array.isArray(raw.internalDomains) &&
        raw.internalDomains.every((s) => typeof s === 'string')
      ) {
        // Lowercase + de-duplicate + drop empties so storage stays canonical.
        const norm = (raw.internalDomains as string[])
          .map((s) => s.toLowerCase().trim().replace(/^@/, ''))
          .filter(Boolean)
        clean.internalDomains = Array.from(new Set(norm))
      }
      if (typeof raw.primaryRootPath === 'string') {
        // Trim leading/trailing slashes — initial-scan / classifier compare
        // against MailFolderNode.path which has no leading slash and no
        // trailing slash.
        clean.primaryRootPath = raw.primaryRootPath.trim().replace(/^\/+|\/+$/g, '')
      }
      if (
        Array.isArray(raw.internalSubjectCategories) &&
        raw.internalSubjectCategories.every((s) => typeof s === 'string')
      ) {
        clean.internalSubjectCategories = (raw.internalSubjectCategories as string[])
          .map((s) => s.trim())
          .filter(Boolean)
      }
      if (typeof raw.aiIncludeFewShotExamples === 'boolean') {
        clean.aiIncludeFewShotExamples = raw.aiIncludeFewShotExamples
      }
      if (typeof raw.onboardingDismissed === 'boolean') {
        clean.onboardingDismissed = raw.onboardingDismissed
      }
      if (typeof raw.harvestEnabled === 'boolean') {
        clean.harvestEnabled = raw.harvestEnabled
      }
      // 結案歸檔 prefix 必須維持「排除清單的子集」— (2026-11-05) now enforced
      // by storage.setSettings on the MERGED result under the settings lock,
      // for every writer (this handler, sync pull, backup restore). The old
      // in-handler prune read the current closedCasePrefixes OUTSIDE that
      // lock (a concurrent sync pull could be clobbered with a stale base) and
      // only ran when the patch carried excludeFolderPrefixes.
      const next = await setSettings(clean)
      return {
        ok: true,
        data: { settings: { ...next, claudeApiKey: next.claudeApiKey ? '••••' : '' } },
      }
    }

    case 'startExecute': {
      const plan = msg.plan as PlanItem[] | undefined
      const tree = msg.tree as MailFolderNode[] | undefined
      if (!plan || !Array.isArray(plan) || plan.length === 0) {
        return { ok: false, code: 'EMPTY_PLAN', message: 'plan 為空' }
      }
      const resolvedTree = tree ?? (await getFolderCache())?.tree
      if (!resolvedTree) {
        return { ok: false, code: 'NO_TREE', message: '資料夾樹不存在，請回去重新分類' }
      }
      // isExecuteBusy() is re-checked synchronously right before the call
      // (2026-11-05): startExecute claims the executor before ITS first
      // await, so check + claim are atomic and an overlapping start gets
      // ALREADY_RUNNING instead of a misleading ok.
      if ((await isExecuteRunning()) || isExecuteBusy()) {
        return { ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' }
      }
      await recordClassifyRunExecuted(plan.map((p) => p.emailId)) // 2026-11-05: never re-offered by this run's progress
      // Fire-and-forget; popup polls getExecuteState.
      void startExecute(plan, resolvedTree, api).catch((err) =>
        console.error('[mail-organizer] startExecute failed', err),
      )
      return { ok: true, data: { started: true, total: plan.length } }
    }

    case 'getExecuteState': {
      const state = await getExecuteState()
      // Slim poll payload (2026-11-05): the popup polls this every 500 ms
      // while executing and never reads `plan` (its ExecuteState type omits
      // it) — shipping the full plan each tick was pure messaging overhead.
      if (!state) return { ok: true, data: { state: null } }
      const { plan: _plan, ...slim } = state
      return { ok: true, data: { state: slim } }
    }

    case 'cancelExecute': {
      await requestCancel()
      return { ok: true }
    }

    case 'clearExecuteState': {
      // Second entrance of the stale-progress bug (2026-06-03): after
      // 完成 clears the execute state, a popup reopen's resume path
      // finds no executeState + no popupState but DOES find the
      // executed batch's leftover AI progress — and promotes it as a
      // live plan, re-offering already-moved emails.
      //
      // 2026-11-05: the old purge (progress.startedAt <= execute.startedAt)
      // broke with partial execution — it deleted a still-running run's
      // done:false progress, and a run that finished DURING the execute
      // looked like residue, so its remaining mail was re-classified (paid
      // twice). The run's executed-id exclusion now keeps moved mail out of
      // every read, so only a FINISHED run with nothing left to offer is
      // residue. A live run (done:false) is never purged.
      const [run, raw] = await Promise.all([
        loadClassifyRun(),
        chrome.storage.session.get(AI_PROGRESS_KEY),
      ])
      const prog = raw[AI_PROGRESS_KEY] as StoredAiProgress | undefined
      if (prog) {
        if (!run || prog.runId !== run.runId) {
          await chrome.storage.session.remove(AI_PROGRESS_KEY)
        } else if (prog.done) {
          const left = viewAiProgress(prog, run.executedIds, 'meta', 0)
          if (left.ruleCount + left.aiCount - left.aiUnclassifiedCount === 0) {
            await cancelClassifyRun(run.runId)
          }
        }
      }
      await clearExecuteState()
      return { ok: true }
    }

    case 'getUndoSnapshot': {
      const snap = await getUndoSnapshot()
      return { ok: true, data: { snapshot: snap } }
    }

    case 'executeUndo': {
      try {
        const result = await executeUndo(api)
        return { ok: true, data: result }
      } catch (e) {
        const message =
          e instanceof OutlookError
            ? `[${e.status}] ${e.message}`
            : e instanceof Error
            ? e.message
            : String(e)
        return { ok: false, code: 'UNDO_FAILED', message }
      }
    }

    case 'dismissUndo': {
      await dismissUndo()
      return { ok: true }
    }

    case 'retryFailed': {
      // Whole flow lives in execute.ts (2026-11-05) so the executor is
      // claimed synchronously before the first await (invariant 6).
      const r = await startRetryFailed(api, async () => (await getFolderCache())?.tree)
      return r.ok
        ? { ok: true, data: { started: r.started } }
        : { ok: false, code: r.code, message: r.message }
    }

    case 'startInitialScan': {
      const settings = await getSettings()
      // Caller-specified rootPath wins (lets the options UI scan an
      // arbitrary subtree without changing the firm setting), but the
      // primary root from settings is the canonical default.
      const rootPath =
        typeof msg.rootPath === 'string' && msg.rootPath.length > 0
          ? msg.rootPath
          : settings.primaryRootPath
      if (!rootPath) {
        return {
          ok: false,
          code: 'NO_ROOT_CONFIGURED',
          message: '尚未設定主要根資料夾(設定頁 → 歸類偏好 → 主要根資料夾)',
        }
      }

      if (await isScanRunning()) {
        return { ok: false, code: 'ALREADY_RUNNING', message: '已有掃描中的批次' }
      }

      // Fail fast if OWA isn't reachable or token unavailable — otherwise the
      // scan loop would log "NOT_ON_OWA" error against every folder.
      const owa = await pingOwa()
      if (!owa.ok) {
        return { ok: false, code: owa.code, message: owa.message }
      }
      const tok = await getOwaToken()
      if (!tok.ok) {
        return { ok: false, code: tok.code, message: tok.message }
      }

      const tree = (await getFolderCache())?.tree ?? (await getOrFetchFolderTree(false))

      // Fail fast with a SYNCHRONOUS error the UI can render. The scan
      // itself is fire-and-forgotten below, so its own throw for this
      // condition (root folder absent from the tree) was invisible — the
      // UI got ok:true, refreshed to a null scan state, and silently
      // showed the start button again with the reason buried in the SW
      // console.
      const hasTargets = flattenFolderTree(tree).some((n) => {
        const inRoot = n.path === rootPath || n.path.startsWith(rootPath + '/')
        if (!inRoot) return false
        return !settings.excludeFolderPrefixes.some(
          (p) => n.path === p || n.path.startsWith(p + '/'),
        )
      })
      if (!hasTargets) {
        return {
          ok: false,
          code: 'NO_TARGETS',
          message: `找不到根資料夾「${rootPath}」或其子資料夾(可能已改名/被排除清單涵蓋)`,
        }
      }

      void startInitialScan({
        rootPath,
        tree,
        api,
        excludePrefixes: settings.excludeFolderPrefixes,
        internalDomains: settings.internalDomains,
      }).catch((err) => console.error('[mail-organizer] startInitialScan failed', err))
      return { ok: true, data: { started: true, rootPath } }
    }

    case 'getScanState': {
      const state = await getScanState()
      return { ok: true, data: { state } }
    }

    case 'cancelScan': {
      await requestScanCancel()
      return { ok: true }
    }

    case 'clearScanState': {
      await clearScanState()
      return { ok: true }
    }

    case 'getRules': {
      const rules = await listRules()
      const conflicts = findConflicts(rules)
      return {
        ok: true,
        data: {
          rules,
          conflicts: conflicts.map((c) => ({
            type: c.type,
            signal: c.signal,
            ruleIds: c.rules.map((r) => r.id),
            targets: c.rules.map((r) => r.targetFolderPath),
          })),
        },
      }
    }

    case 'createRule': {
      const input = msg.input as Parameters<typeof newRule>[0] | undefined
      if (!input || !input.type || !input.signal || !input.targetFolderId || !input.targetFolderPath) {
        return { ok: false, code: 'INVALID_RULE', message: '欠缺必填欄位' }
      }
      const rule = newRule({
        type: input.type,
        signal: input.signal,
        targetFolderId: input.targetFolderId,
        targetFolderPath: input.targetFolderPath,
        confidence: typeof input.confidence === 'number' ? input.confidence : 0.85,
        source: input.source ?? 'user_manual',
        enabled: input.enabled ?? true,
      })
      await upsertRule(rule, { actor: 'user' })
      return { ok: true, data: { rule } }
    }

    case 'upsertRule': {
      const rule = msg.rule as Rule | undefined
      if (!rule?.id) return { ok: false, code: 'INVALID_RULE', message: '規則資料無效' }
      // mode (2026-11-05): 'edit' must not re-insert a rule deleted after
      // the Options page loaded; absent = legacy merge-or-insert.
      const mode = msg.mode === 'create' || msg.mode === 'edit' ? msg.mode : undefined
      try {
        // Return the STORED rule: upsertRule merges the editable fields onto
        // it, so the page gets live stats back instead of its stale copy.
        const saved = await upsertRule(rule, { actor: 'user', mode })
        return { ok: true, data: { rule: saved } }
      } catch (e) {
        if ((e as { code?: unknown } | null)?.code === 'NOT_FOUND') {
          return { ok: false, code: 'NOT_FOUND', message: '這條規則已在別處被刪除，請重新整理' }
        }
        throw e
      }
    }

    case 'deleteRule': {
      const id = typeof msg.ruleId === 'string' ? msg.ruleId : null
      if (!id) return { ok: false, code: 'NO_ID', message: '未提供 ruleId' }
      const deleted = await deleteRule(id, { actor: 'user' })
      return { ok: true, data: { deleted } }
    }

    case 'toggleRule': {
      const id = typeof msg.ruleId === 'string' ? msg.ruleId : null
      const enabled = msg.enabled === true
      if (!id) return { ok: false, code: 'NO_ID', message: '未提供 ruleId' }
      const updated = await toggleRule(id, enabled, { actor: 'user' })
      if (!updated) {
        // (2026-11-05) Was ok:true with rule:undefined — the popup's
        // 「停用此規則」 then showed "已停用" for a rule that no longer exists
        // (deleted in Options / by a sync pull) and the Options bulk path
        // counted it as a success.
        return { ok: false, code: 'NOT_FOUND', message: '找不到指定規則(可能已被刪除或同步移除)' }
      }
      return { ok: true, data: { rule: updated } }
    }

    case 'bulkRuleAction': {
      // { op: 'enable' | 'disable' | 'delete'; ids: string[] } →
      // { changed, missing }. One locked rewrite + one history append for the
      // whole selection (see handlers/rules-bulk.ts) instead of one round
      // trip — and one full rule-library rewrite — per rule.
      return handleBulkRuleAction(msg as { op?: unknown; ids?: unknown })
    }

    case 'clearAllRules': {
      // INTENTIONAL DEV / DIAGNOSTIC HANDLER — no UI caller. Reachable
      // only via `chrome.runtime.sendMessage({type:'clearAllRules'})`
      // from the extension's background console. User-facing equivalent
      // is `importRules` with `strategy:'replace'` and an empty payload,
      // which goes through the merge-preview UX in Options.
      //
      // Record one delete event per rule so the history reflects the wipe.
      // Bulk path — capture snapshots before we obliterate the store.
      // F12: clear INSIDE mutateRules so a concurrent rule mutation
      // (a finishing batch's bumpRuleHits, an ai_confirmed generation)
      // can't interleave its read-modify-write around our raw write and
      // resurrect rules after the wipe.
      const existing = await mutateRules<Rule[]>((current) => {
        return { next: [], result: current }
      })
      if (existing.length > 0) {
        const now = Date.now()
        await recordRuleEvents(
          existing.map((r) => ({
            kind: 'delete' as const,
            ruleId: r.id,
            at: now,
            actor: 'user' as const,
            before: {
              type: r.type,
              signal: r.signal,
              targetFolderPath: r.targetFolderPath,
              confidence: r.confidence,
              source: r.source,
              enabled: r.enabled,
            },
          })),
        ).catch((e) =>
          console.warn('[mail-organizer] recordRuleEvents (clearAllRules) failed', e),
        )
      }
      return { ok: true, data: { cleared: existing.length } }
    }

    case 'wipeAllRules': {
      // User-initiated complete wipe for "start clean" workflow. Per the
      // 2026-05-27 "keep library clean" design:
      //   1. Rules array → emptied
      //   2. Tombstone library → cleared (otherwise a follow-up import
      //      could be partially blocked, and AI would be unable to
      //      re-learn signals that the user used to have)
      //   3. AI memory (conversation + subject) → cleared (otherwise
      //      the next batch would still route based on pre-wipe thread
      //      / subject decisions, contaminating "fresh start")
      //   4. Rule events log → cleared (no audit trail of pre-wipe state)
      //
      // Sync handling (2026-05-27 refactor — fixes audit P0 race +
      // tombstone-resurrection):
      //   - quiesce drains in-flight push + cancels pending debounce
      //   - clearCloudState wipes cloud-side keys
      //   - If clearCloudState FAILS, we ABORT BEFORE touching local
      //     state. Otherwise: cloud full + local empty + tombstones
      //     cleared → next push uses our empty tombstones, losing the
      //     user's deletion intent across machines (audit P0-3:
      //     "tombstone resurrection on clearCloud-fail"). Better to
      //     refuse the wipe and surface the underlying network /
      //     storage error than to half-wipe in a way the user can't
      //     recover from cleanly.
      //   - After local wipe succeeds, pushNow('post-wipe', {
      //     wipeMarker: true }) stamps SyncMeta.wipeMarker so OTHER
      //     machines detect the wipe on their next pull and apply it
      //     locally (audit P0-1: "cross-machine wipe propagation").
      //     Without this, machine B would keep its syncable rules
      //     and push them back to cloud, undoing the wipe.
      //
      // Concurrency: a single in-flight guard prevents the user
      // double-clicking the "全部刪除" button from interleaving two
      // wipe sequences. The second call returns WIPE_IN_FLIGHT until
      // the first finishes.
      //
      // skipHistory + folderActivity intentionally preserved — they're
      // user-explicit decisions / UI display, not routing learning input.
      if (wipeInFlight) {
        return {
          ok: false,
          code: 'WIPE_IN_FLIGHT',
          message: '已有清除作業進行中,請稍候再試',
        }
      }
      wipeInFlight = true
      try {
        const existing = await listRules()
        const settings = await getSettings()
        if (settings.syncEnabled) {
          // Step 1: drain pending push + clear cloud FIRST. If this
          // fails the local state is still intact (audit P0-3) and
          // the user can retry or disable sync first.
          try {
            await quiesce()
            await clearCloudState()
          } catch (e) {
            const errMsg = e instanceof Error ? e.message : String(e)
            console.warn('[mail-organizer] clearCloudState (wipe) failed', e)
            return {
              ok: false,
              code: 'CLOUD_CLEAR_FAILED',
              message:
                '雲端清除失敗,本機規則未動。請檢查網路後重試,或先在「跨機器同步」停用同步再清除。錯誤:' +
                errMsg.slice(0, 200),
            }
          }
        }
        // Step 2: local wipe. Order doesn't matter — these are
        // independent local-storage writes. Each wrapped because we
        // don't want a single failure to skip the others (e.g.
        // tombstone clear failing shouldn't leave AI memory intact —
        // a fresh start should be a fresh start).
        // F12: empty INSIDE mutateRules so a concurrent rule mutation
        // (a finishing batch's bumpRuleHits, an ai_confirmed generation,
        // an Options upsert) can't read-modify-write around a raw
        // setRules([]) and resurrect rules after the wipe. wipeInFlight
        // only blocks a second WIPE, not other rule writers.
        await mutateRules(() => ({ next: [], result: undefined }))
        try {
          await clearAllRuleTombstones()
        } catch (e) {
          console.warn('[mail-organizer] clearAllRuleTombstones (wipe) failed', e)
        }
        try {
          await clearAllAiMemory()
        } catch (e) {
          console.warn('[mail-organizer] clearAllAiMemory (wipe) failed', e)
        }
        try {
          await clearRuleHistory()
        } catch (e) {
          console.warn('[mail-organizer] clearRuleHistory (wipe) failed', e)
        }
        // Step 3: stamp wipeMarker into cloud so OTHER machines pick
        // up the wipe on their next pull. Bypass the 5s debounce —
        // the longer we wait, the more likely another machine
        // pushes their (pre-wipe) state and we'd have to overwrite
        // it. pushNow runs synchronously.
        let postWipePushOk = true
        if (settings.syncEnabled) {
          try {
            const r = await pushNow('post-wipe', { wipeMarker: true })
            postWipePushOk = r.pushed === true
            if (!postWipePushOk) {
              console.warn(
                '[mail-organizer] post-wipe push did not complete',
                r.reason,
              )
            }
          } catch (e) {
            postWipePushOk = false
            console.warn('[mail-organizer] post-wipe push threw', e)
          }
        }
        return {
          ok: true,
          data: {
            cleared: existing.length,
            cloudCleared: settings.syncEnabled,
            cloudClearStatus: settings.syncEnabled ? 'cleared' : 'skipped',
            postWipePushed: settings.syncEnabled ? postWipePushOk : undefined,
            syncEnabled: settings.syncEnabled,
          },
        }
      } finally {
        wipeInFlight = false
      }
    }

    case 'exportRules': {
      // v2 (2026-05-22) bundles tombstones into the payload so syncing
      // across machines transfers deletion intent — without this, a
      // deleted rule on A would get auto-recreated on B as soon as it
      // sees a matching email pattern.
      const [rules, tombstones] = await Promise.all([listRules(), getRuleTombstones()])
      const payload = serializeRules(
        rules,
        chrome.runtime.getManifest().version,
        tombstones,
      )
      return {
        ok: true,
        data: {
          payload,
          count: rules.length,
          tombstoneCount: tombstones.length,
        },
      }
    }

    case 'previewRuleImport': {
      const payload = typeof msg.payload === 'string' ? msg.payload : ''
      if (!payload) {
        return { ok: false, code: 'BAD_INPUT', message: '缺少 payload' }
      }
      const parsed = parseRulesPayload(payload)
      if (!parsed.ok) {
        return { ok: false, code: 'PARSE_ERROR', message: parsed.error }
      }
      const existing = await listRules()
      const preview = previewImport(existing, parsed.rules, parsed.tombstones)
      return { ok: true, data: { preview, meta: parsed.meta } }
    }

    case 'importRules': {
      // `payload` is the contract. `json` is accepted as an alias: the
      // Options page sent `json` since the initial commit, so every import
      // failed with 「缺少 payload」 (2026-11-05 audit) — same drift class as
      // the export "undefined" file bug (75e436e). Covered by a contract test
      // that drives this dispatcher with the exact UI message shape.
      const payload =
        typeof msg.payload === 'string'
          ? msg.payload
          : typeof msg.json === 'string'
            ? msg.json
            : ''
      const strategy = msg.strategy as ImportStrategy | undefined
      if (!payload) {
        return { ok: false, code: 'BAD_INPUT', message: '缺少 payload' }
      }
      if (strategy !== 'merge' && strategy !== 'replace') {
        return { ok: false, code: 'BAD_INPUT', message: 'strategy 必須是 merge 或 replace' }
      }
      const parsed = parseRulesPayload(payload)
      if (!parsed.ok) {
        return { ok: false, code: 'PARSE_ERROR', message: parsed.error }
      }
      // All under the rules mutex so concurrent bumpRuleHit / upsert
      // can't race with the import. Reconcile happens in the same critical
      // section against whichever folder tree we have cached locally — this
      // catches the cross-machine case where imported targetFolderIds point
      // to folders that don't exist on the current Outlook account.
      //
      // We capture the diff inside mutateRules so the audit log reflects
      // exactly what changed. Replace strategy emits delete events for all
      // previously-existing rules; both strategies emit create events for
      // newly-added rules. Edit events aren't possible here because the
      // import key is (type, signal) — a "matching" rule keeps its old id
      // under merge (it's skipped, not edited).
      const cachedTree = (await getFolderCache())?.tree
      let auditEvents: RuleEvent[] = []
      // Track confidence-lowering by the cap so the user sees how
      // much trust their backup leaked. Backups from an older build
      // (pre-2026-05-13 per-type caps) can carry domain rules at
      // conf=1.0, sender at 0.95, etc. Silently capping makes the
      // user think "import worked fine" — surface the delta so they
      // understand "47 rules got their conf trimmed; lowest cap was
      // domain → 0.7".
      let cappedCount = 0
      let cappedMaxDelta = 0
      const result = await mutateRules((existing) => {
        // Pre-cap imported rules — backups / shared JSON may carry
        // confidence above the per-type ceiling we now enforce
        // elsewhere. Skipping this would let imports silently
        // bypass the cap.
        const cappedImported = parsed.rules.map((r) => {
          const capped = applyConfidenceCap(r)
          // Round to avoid float-precision noise (e.g. 0.69999... vs 0.7
          // from JSON parsing). Threshold 0.0005 ignores trivial cap
          // hits where the rule was effectively at the ceiling already.
          const delta = (r.confidence ?? 0) - (capped.confidence ?? 0)
          if (delta > 0.0005) {
            cappedCount++
            if (delta > cappedMaxDelta) cappedMaxDelta = delta
          }
          return capped
        })
        const merged = applyImport(existing, cappedImported, strategy)
        const added = strategy === 'merge' ? merged.length - existing.length : cappedImported.length
        const replaced = strategy === 'replace' ? existing.length : 0
        let orphanedCount = 0
        let next = merged
        if (cachedTree && cachedTree.length > 0) {
          const recon = reconcileRulesAgainstTree(merged, cachedTree)
          next = recon.rules
          orphanedCount = next.filter((r) => r.orphaned).length
        }
        const now = Date.now()
        const existingIds = new Set(existing.map((r) => r.id))
        const nextIds = new Set(next.map((r) => r.id))
        const events: RuleEvent[] = []
        for (const r of existing) {
          if (!nextIds.has(r.id)) {
            events.push({
              kind: 'delete',
              ruleId: r.id,
              at: now,
              actor: 'user',
              before: {
                type: r.type,
                signal: r.signal,
                targetFolderPath: r.targetFolderPath,
                confidence: r.confidence,
                source: r.source,
                enabled: r.enabled,
              },
            })
          }
        }
        for (const r of next) {
          if (!existingIds.has(r.id)) {
            events.push({
              kind: 'create',
              ruleId: r.id,
              at: now,
              actor: 'user',
              after: {
                type: r.type,
                signal: r.signal,
                targetFolderPath: r.targetFolderPath,
                confidence: r.confidence,
                source: r.source,
                enabled: r.enabled,
              },
            })
          }
        }
        auditEvents = events
        return {
          next,
          result: {
            added,
            replaced,
            total: next.length,
            skipped: parsed.rules.length - added,
            orphanedCount,
          },
        }
      })
      if (auditEvents.length > 0) {
        await recordRuleEvents(auditEvents).catch((e) =>
          console.warn('[mail-organizer] recordRuleEvents (importRules) failed', e),
        )
      }
      // v2 payload: also merge tombstones so deletions transfer across
      // machines. addRuleTombstones is dedup-by-triple internally, so
      // already-present tombstones are no-ops.
      //
      // Why this matters: without it, the user's deletion of a noisy
      // rule on machine A would get auto-recreated on machine B as
      // soon as it sees a matching email pattern (via
      // ai_confirmed / auto_scan / etc.) — defeating the deletion.
      let importedTombstones = 0
      if (parsed.tombstones.length > 0) {
        try {
          await addRuleTombstones(parsed.tombstones)
          importedTombstones = parsed.tombstones.length
        } catch (e) {
          console.warn('[mail-organizer] addRuleTombstones (importRules) failed', e)
        }
      }
      return {
        ok: true,
        data: {
          ...result,
          importedTombstones,
          // Surface how many imported rules had their confidence
          // trimmed by the per-type cap (and the worst-case delta in
          // 0..1 units). Lets the options UI explain "your backup had
          // 47 rules above the current ceiling; we lowered them to
          // match the new defaults" instead of silently changing
          // user-trusted numbers.
          cappedCount,
          cappedMaxDelta: Math.round(cappedMaxDelta * 100) / 100,
        },
      }
    }

    case 'getRuleHistory': {
      const limit = typeof msg.limit === 'number' && msg.limit > 0 ? msg.limit : 100
      const events = await getRuleEvents(limit)
      return { ok: true, data: { events } }
    }

    case 'clearRuleHistory': {
      await clearRuleHistory()
      return { ok: true }
    }

    case 'resolveConflict':
      // keep_highest / keep_one delete the losers (tombstoned), disable_all
      // disables the group — one locked rewrite + one history append
      // (2026-11-05; previously one deleteRule / toggleRule per rule).
      return handleResolveConflict(msg as { ruleIds?: unknown; strategy?: unknown; keepId?: unknown })

    case 'clearDisabledRules':
      // Delete every enabled=false rule, tombstoned exactly like deleteRule so
      // AI / initial-scan won't recreate them — in ONE rewrite (2026-11-05).
      return handleClearDisabledRules()

    case 'runStaleSweep': {
      // Manual trigger for the daily auto-disable sweep. Same operation
      // as the chrome.alarms-driven path. Returns both the soft-disabled
      // count (legacy_token / high-error-rate) and the stale hard-delete
      // count so the options button can report what actually happened.
      const { disabledCount, staleDeletedCount } = await runSweep({ force: true })
      return { ok: true, data: { disabledCount, staleDeletedCount } }
    }

    // ---- Cross-machine sync (chrome.storage.sync) ----------------------
    // Implementations live in ./handlers/sync.ts. Each `handleX` is
    // self-contained: routing here, logic there.

    case 'getSyncStatus':
      return handleGetSyncStatus()
    case 'enableSync':
      return handleEnableSync()
    case 'disableSync':
      return handleDisableSync(msg as { keepCloud?: boolean })
    case 'pushSyncNow':
      return handlePushSyncNow()
    case 'pullSyncNow':
      return handlePullSyncNow()
    case 'listSyncBackups':
      return handleListSyncBackups()
    case 'restoreSyncBackup':
      return handleRestoreSyncBackup(msg as { snapshotAt?: unknown })
    case 'dismissSyncError':
      return handleDismissSyncError()
    case 'getRemoteWipeNotice': {
      const notice = await readRemoteWipeNotice()
      return { ok: true, data: { notice: notice ?? null } }
    }
    case 'dismissRemoteWipeNotice': {
      await dismissRemoteWipeNotice()
      return { ok: true, data: {} }
    }

    case 'getErrorLog':
      return handleGetErrorLog(msg as { limit?: unknown })
    case 'clearErrorLog':
      return handleClearErrorLog()

    case 'suggestSplitKeywords': {
      // Observe recent subjects in each target folder of a conflict and
      // surface "distinctive" tokens: shows up frequently in folder A but
      // rarely in others. User picks from suggestions instead of guessing.
      const folderIds = Array.isArray(msg.folderIds) ? (msg.folderIds as string[]) : []
      if (folderIds.length < 2) {
        return { ok: false, code: 'INVALID_INPUT', message: '需要至少 2 個資料夾' }
      }

      const TOP_N = 30
      const KEYWORD_COUNT = 5

      // Fetch subjects from each folder in parallel
      const subjectsPerFolder = await Promise.all(
        folderIds.map(async (fid) => {
          try {
            const emails = await api.listFolderMessages(fid, {
              top: TOP_N,
              select: 'Id,Subject,ReceivedDateTime',
            })
            return emails.map((m) => m.Subject ?? '').filter((s) => s.length > 0)
          } catch (e) {
            console.warn(`[mail-organizer] suggestSplitKeywords: failed to fetch ${fid}`, e)
            return []
          }
        }),
      )

      // Tokenize a subject into 2-6 char Chinese / ASCII word tokens
      const tokenize = (s: string): string[] => {
        const cleaned = s
          .replace(/^(re|fw|fwd|轉寄|回覆)[:：\s]+/gi, '')
          .replace(/[【】「」『』（）()\[\]\.，,。、!?:;'"／\/\-\s]+/g, ' ')
          .trim()
        if (!cleaned) return []
        const tokens = new Set<string>()
        // Chinese: contiguous runs of CJK chars, take 2-4 char windows
        const cjkRuns = cleaned.match(/[一-鿿]+/g) ?? []
        for (const run of cjkRuns) {
          for (let len = 2; len <= 4 && len <= run.length; len++) {
            for (let i = 0; i + len <= run.length; i++) {
              tokens.add(run.slice(i, i + len))
            }
          }
        }
        // ASCII: words of length 3+
        const asciiWords = cleaned.match(/[A-Za-z][A-Za-z0-9]{2,}/g) ?? []
        for (const w of asciiWords) tokens.add(w.toLowerCase())
        // Numbers like 26B0013A
        const codes = cleaned.match(/\d{2}[A-Z]\d{4}[A-Z]/gi) ?? []
        for (const c of codes) tokens.add(c.toUpperCase())
        return [...tokens]
      }

      const counts: Map<string, number>[] = subjectsPerFolder.map((subjects) => {
        const m = new Map<string, number>()
        for (const s of subjects) {
          for (const t of tokenize(s)) m.set(t, (m.get(t) ?? 0) + 1)
        }
        return m
      })

      // For each folder, pick tokens that appear at least 2 times AND don't
      // appear in any other folder's counts (or appear much less). Rank by
      // raw frequency.
      const suggestionsPerFolder = counts.map((selfCounts, idx) => {
        const otherCounts = counts.filter((_, i) => i !== idx)
        const candidates: Array<{ token: string; count: number }> = []
        for (const [token, count] of selfCounts) {
          if (count < 2) continue
          const otherMax = Math.max(0, ...otherCounts.map((c) => c.get(token) ?? 0))
          if (otherMax >= count) continue // too common across folders
          candidates.push({ token, count })
        }
        candidates.sort((a, b) => b.count - a.count)
        return candidates.slice(0, KEYWORD_COUNT).map((c) => c.token)
      })

      // Also include recent subjects (truncated) so user can eyeball
      const recentSubjects = subjectsPerFolder.map((subs) =>
        subs.slice(0, 8).map((s) => (s.length > 50 ? s.slice(0, 50) + '…' : s)),
      )

      return {
        ok: true,
        data: {
          suggestions: suggestionsPerFolder,
          recentSubjects,
        },
      }
    }

    case 'splitConflictToCompound':
      // Per-branch keywords → user_manual compounds; originals disabled. One
      // locked rewrite + one history append (2026-11-05; previously one
      // toggleRule per original plus one upsertRule per new compound).
      return handleSplitConflictToCompound(msg as { items?: unknown })

    case 'autoUpgradeConflictRules': {
      // Auto-upgrade a domain / sender conflict to compound rules
      // (domain + 整段主旨), without requiring the user to type
      // keywords. For each conflicting rule:
      //   1. Server-side $filter the rule's target folder so we only
      //      get messages that actually matched this rule's signal.
      //   2. For each unique normalized subject, build a compound rule
      //      `(domain|sender + 主旨) → that rule's target`.
      //   3. Disable the original plain rule.
      //
      // Stays out of the daily flow — only fires when the user opens
      // Options 衝突 tab and clicks 「升級為 compound 規則」, so the cost
      // (one OData fetch per conflicting rule) is bounded and opt-in.
      const ruleIds = Array.isArray(msg.ruleIds) ? (msg.ruleIds as string[]) : []
      if (ruleIds.length < 2) {
        return { ok: false, code: 'INVALID_INPUT', message: '至少要 2 條規則參與升級' }
      }
      const all = await listRules()
      const targets = ruleIds
        .map((rid) => all.find((r) => r.id === rid))
        .filter((r): r is Rule => !!r && r.enabled && !r.orphaned)
      if (targets.length < 2) {
        return { ok: false, code: 'INVALID_INPUT', message: '規則已不存在或已停用' }
      }
      // Only domain / sender conflicts supported. case_code / compound /
      // subject_keyword conflicts represent genuine ambiguity the user
      // should resolve manually (fall through to the existing buttons).
      if (!targets.every((r) => r.type === 'domain' || r.type === 'sender')) {
        return { ok: false, code: 'UNSUPPORTED_TYPE', message: '僅支援網域 / 寄件人衝突自動升級' }
      }
      const anchorType = targets[0]!.type
      if (!targets.every((r) => r.type === anchorType)) {
        return { ok: false, code: 'MIXED_TYPES', message: '所有規則必須是同類型' }
      }

      const TOP_PER_FOLDER = 30
      const TOP_CLIENT_FALLBACK = 100
      const newRules: Rule[] = []
      const disabled: string[] = []
      const fetchErrors: Array<{ signal: string; folder: string; error: string }> = []
      const emptyFolders: Array<{ signal: string; folder: string }> = []

      // Pre-fetch fresh folder tree once so we can recover from stale
      // targetFolderIds (cross-machine sync writes IDs from the other
      // machine's mailbox — those 404 on this machine). Falling back to
      // path lookup gets us the correct local ID.
      let folderById: Map<string, MailFolderNode> | null = null
      let folderByPath: Map<string, MailFolderNode> | null = null
      try {
        const tree = await getOrFetchFolderTree(false)
        const flat = flattenFolderTree(tree)
        folderById = new Map(flat.map((f) => [f.id, f]))
        folderByPath = new Map(flat.map((f) => [f.path, f]))
      } catch (e) {
        console.warn('[mail-organizer] autoUpgradeConflictRules: folder tree fetch failed', e)
      }

      for (const rule of targets) {
        // Narrow rule.type to the compound-condition-eligible subset.
        // Already enforced above ("UNSUPPORTED_TYPE"), but TypeScript
        // doesn't propagate that filter — narrow locally.
        const anchorCondType: 'domain' | 'sender' =
          rule.type === 'domain' ? 'domain' : 'sender'

        // Resolve targetFolderId — prefer stored, fall back to path
        // lookup against current tree (handles sync-from-other-machine
        // and folder-renamed-or-recreated scenarios).
        let effectiveFolderId = rule.targetFolderId
        if (folderById && folderByPath) {
          if (!effectiveFolderId || !folderById.has(effectiveFolderId)) {
            const byPathHit = folderByPath.get(rule.targetFolderPath)
            if (byPathHit) {
              effectiveFolderId = byPathHit.id
            }
          }
        }
        if (!effectiveFolderId) {
          fetchErrors.push({
            signal: rule.signal,
            folder: rule.targetFolderPath,
            error: '找不到對應的資料夾(可能已刪除或重新命名)',
          })
          continue
        }
        // Re-derive the CURRENT path from the resolved folder. If the
        // folder was renamed out-of-band (Outlook UI), the rule's
        // stored `targetFolderPath` is stale even though the id is
        // still valid. Without this re-derivation, any compound rule
        // we mint inherits the stale path and surfaces wrong in the
        // UI. Reconcile-on-preflight eventually fixes it, but new
        // rules created here shouldn't bake in known-stale data.
        const liveFolder = folderById?.get(effectiveFolderId)
        const freshFolderPath = liveFolder?.path ?? rule.targetFolderPath

        // OData $filter constrains the fetch to mail matching this
        // rule's signal — uses PascalCase per Outlook REST v2.0 spec.
        // If the filter is rejected by the server OR returns 0 mail,
        // fall back to fetching un-filtered top-N + client-side filter
        // (some mailbox configurations / older endpoints have quirks
        // with endswith on From). Client-side fallback is bounded at
        // TOP_CLIENT_FALLBACK to avoid pulling huge folders.
        const escapedSignal = rule.signal.replace(/'/g, "''")
        const signalLowered = rule.signal.toLowerCase()
        const matchesSignal = (addr: string): boolean => {
          const a = addr.toLowerCase()
          return anchorCondType === 'domain'
            ? a.endsWith(`@${signalLowered}`)
            : a === signalLowered
        }
        const filter =
          anchorCondType === 'domain'
            ? `endswith(From/EmailAddress/Address, '@${escapedSignal}')`
            : `From/EmailAddress/Address eq '${escapedSignal}'`
        let emails: Awaited<ReturnType<typeof api.listFolderMessages>>
        let filterPath: 'server' | 'client-fallback' = 'server'
        try {
          emails = await api.listFolderMessages(effectiveFolderId, {
            top: TOP_PER_FOLDER,
            select: 'Id,Subject,From,ReceivedDateTime',
            filter,
          })
        } catch (e) {
          // Server filter rejected (4xx) — try without filter
          console.warn(
            `[mail-organizer] autoUpgradeConflictRules: server filter rejected for ${rule.targetFolderPath}, falling back to client-side filter`,
            e,
          )
          try {
            const unfiltered = await api.listFolderMessages(effectiveFolderId, {
              top: TOP_CLIENT_FALLBACK,
              select: 'Id,Subject,From,ReceivedDateTime',
            })
            emails = unfiltered.filter((m) =>
              matchesSignal(m.From?.EmailAddress?.Address ?? ''),
            )
            filterPath = 'client-fallback'
          } catch (e2) {
            const errMsg = e2 instanceof Error ? e2.message : String(e2)
            fetchErrors.push({
              signal: rule.signal,
              folder: rule.targetFolderPath,
              error: errMsg.slice(0, 200),
            })
            continue
          }
        }

        // Server filter returned 0 — possible silent failure on certain
        // mailbox setups. Try the client-side fallback once before
        // declaring the folder empty.
        if (emails.length === 0 && filterPath === 'server') {
          try {
            const unfiltered = await api.listFolderMessages(effectiveFolderId, {
              top: TOP_CLIENT_FALLBACK,
              select: 'Id,Subject,From,ReceivedDateTime',
            })
            const clientFiltered = unfiltered.filter((m) =>
              matchesSignal(m.From?.EmailAddress?.Address ?? ''),
            )
            if (clientFiltered.length > 0) {
              emails = clientFiltered
              filterPath = 'client-fallback'
            }
          } catch (e) {
            // Non-fatal — treat as truly empty below.
            console.warn(
              `[mail-organizer] autoUpgradeConflictRules: client-fallback fetch failed for ${rule.targetFolderPath}`,
              e,
            )
          }
        }

        // Dedup subjects via the normalized form `extractSubjectSignal`
        // produces (Re:/Fwd:/whitespace stripped). Caps per rule so a
        // single folder can't explode the rule library.
        const MAX_NEW_RULES_PER_CONFLICTING_RULE = 10
        const seenSignals = new Set<string>()
        for (const m of emails) {
          if (seenSignals.size >= MAX_NEW_RULES_PER_CONFLICTING_RULE) break
          const subjectSignal = extractSubjectSignal(m.Subject ?? '')
          if (!subjectSignal) continue
          if (seenSignals.has(subjectSignal)) continue
          seenSignals.add(subjectSignal)
          const compoundSignal = encodeCompound([
            { type: anchorCondType, value: rule.signal },
            { type: 'subject_keyword', value: subjectSignal },
          ])
          newRules.push(
            newRule({
              type: 'compound',
              signal: compoundSignal,
              // Use the resolved folder id (path-resolved when stored id
              // was stale) so the new rule matches the local mailbox.
              // `freshFolderPath` reflects the folder's current display
              // path — avoids baking a stale rename into the new rule.
              targetFolderId: effectiveFolderId,
              targetFolderPath: freshFolderPath,
              confidence: 0.85, // matches compound + full_subject in execute
              source: 'ai_confirmed',
            }),
          )
        }
        if (seenSignals.size === 0) {
          // Folder had no mail matching this signal (or all subjects were
          // empty/single-char). Don't demote — the rule might still be
          // valid; user can manually disable if needed.
          emptyFolders.push({ signal: rule.signal, folder: rule.targetFolderPath })
          continue
        }
        // Fetch succeeded AND at least one compound built → safe to
        // demote the broad rule. User's "升級" click expressed intent
        // that the broad rule is no longer enough.
        disabled.push(rule.id)
      }

      // Apply: add new compounds (via tombstone-aware path to respect
      // user deletions), then disable original rules.
      let createdCount = 0
      if (newRules.length > 0) {
        const { added } = await addRulesFilteringTombstones(newRules)
        createdCount = added.length
      }
      for (const id of disabled) {
        try {
          await toggleRule(id, false, { actor: 'system' })
        } catch (e) {
          console.warn('[mail-organizer] autoUpgradeConflictRules: toggle failed', id, e)
        }
      }

      // If we couldn't make any progress, surface a useful error
      // instead of silently returning `{created:0, disabled:0}` — the
      // user clicked something, they deserve to know why nothing changed.
      if (createdCount === 0 && disabled.length === 0) {
        const reason = fetchErrors.length > 0
          ? `抓取失敗 ${fetchErrors.length} 條: ${fetchErrors.map(e => `${e.signal}→${e.folder} (${e.error})`).join(' | ')}`
          : emptyFolders.length > 0
          ? `${emptyFolders.length} 個資料夾未抓到符合 signal 的信件 (可能資料夾為空 / 信件已被搬走 / signal 不再對應)`
          : '未知原因 — 請查看 service worker 控制台'
        return { ok: false, code: 'NO_PROGRESS', message: reason }
      }

      return {
        ok: true,
        data: {
          created: createdCount,
          disabled: disabled.length,
          fetchErrors: fetchErrors.length,
          emptyFolders: emptyFolders.length,
        },
      }
    }

    case 'getMetrics': {
      const metrics = await getMetrics()
      return { ok: true, data: metrics }
    }

    case 'getFolderTree': {
      const tree = await getOrFetchFolderTree(msg.forceFresh === true)
      return { ok: true, data: { tree, count: flattenFolderTree(tree).length } }
    }

    case 'getClassifyStage': {
      const r = await chrome.storage.session.get(CLASSIFY_STAGE_KEY)
      const stage = (r[CLASSIFY_STAGE_KEY] as ClassifyStage | undefined) ?? { stage: 'idle' }
      return { ok: true, data: stage }
    }

    case 'getStorageUsage': {
      // (2026-11-05) Usage used to be visible only in the SW console and the
      // diagnostic export. quotaBytes is the RUNTIME quota (null with
      // unlimitedStorage); topKeys = largest keys by Chrome's own accounting,
      // names only — never values (settings holds the API key).
      const [usage, topKeys] = await Promise.all([getStorageUsage(), getStorageTopKeys()])
      return {
        ok: true,
        data: { bytesInUse: usage.bytesInUse, quotaBytes: usage.quotaBytes, topKeys },
      }
    }

    case 'exportDiagnostic': {
      const [settings, cache, rules, metrics, skipCount, storageUsage, storageTopKeys] =
        await Promise.all([
          getSettings(),
          getFolderCache(),
          listRules(),
          getMetrics(),
          getSkipHistoryCount(),
          getStorageUsage(),
          getStorageTopKeys(),
        ])
      const lastExecute = await getExecuteState()
      // API key is sensitive — never include in export
      const { claudeApiKey: _omit, ...safeSettings } = settings
      return {
        ok: true,
        data: {
          exportedAt: new Date().toISOString(),
          version: chrome.runtime.getManifest().version,
          settings: { ...safeSettings, claudeApiKey: '[redacted]' },
          rules,
          metrics,
          skipHistoryCount: skipCount,
          storageUsage: {
            bytesInUse: storageUsage.bytesInUse,
            // null = unlimitedStorage (no enforced quota).
            quotaBytes: storageUsage.quotaBytes,
            percentUsed:
              storageUsage.percentUsed === null
                ? null
                : Math.round(storageUsage.percentUsed * 100) / 100,
            approaching: storageUsage.approaching,
            // Key names + sizes only (never values) — tells a diagnosis WHAT
            // is filling storage.
            topKeys: storageTopKeys,
          },
          folderCache: cache
            ? {
                updatedAt: cache.updatedAt,
                totalNodes: flattenFolderTree(cache.tree).length,
                topLevelCount: cache.tree.length,
                // omit full tree (large + may contain client names)
              }
            : null,
          lastExecute: lastExecute
            ? {
                startedAt: lastExecute.startedAt,
                finishedAt: lastExecute.finishedAt,
                summary: lastExecute.summary,
                rulesAdded: lastExecute.rulesAdded,
                errors: lastExecute.results
                  .filter((r) => r.status === 'error')
                  .map((r) => ({ subject: r.subject, action: r.action, message: r.message }))
                  .slice(0, 20),
              }
            : null,
        },
      }
    }

    default: {
      // Compile-time exhaustiveness (2026-11-05): a MessageType without a
      // `case` fails tsc here. At runtime this is reached by strings outside
      // the union (stale UI build, typo in an untyped sender).
      const unhandled: never = msg.type
      return {
        ok: false,
        code: 'UNKNOWN_TYPE',
        message: `Unknown request type: ${String(unhandled)}`,
      }
    }
  }
}

function maskApiKey(key: string): string {
  if (key.length <= 14) return '••••'
  return `${key.slice(0, 10)}••••${key.slice(-4)}`
}
