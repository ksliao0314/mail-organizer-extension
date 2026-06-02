// MV3 service worker — central dispatcher.
//
// Production handlers used by popup and options:
//   getStatus / classify / setApiKey / setSettings
//   startExecute / getExecuteState / cancelExecute / clearExecuteState / retryFailed
//   startInitialScan / getScanState / cancelScan / clearScanState
//   getRules / createRule / upsertRule / deleteRule / toggleRule / clearAllRules
//   getMetrics / getFolderTree / getClassifyStage

import { OutlookApi, OutlookError, flattenFolderTree } from '@/shared/outlook-api'
import { ClassifierError, classifyBatch, type ClassifierUsage } from '@/shared/classifier'
import {
  clearExecuteState,
  dismissUndo,
  executeUndo,
  getExecuteState,
  installUndoExpireListener,
  isExecuteRunning,
  recoverStaleExecuteState,
  requestCancel,
  startExecute,
} from './execute'
import {
  clearScanState,
  getScanState,
  isScanRunning,
  recoverStaleScanState,
  requestScanCancel,
  startInitialScan,
} from './initial-scan'
import { installActionRouter } from './action-router'
import { logError } from '@/shared/error-log'
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
import type { PopupResponse } from '@/shared/messages'
import {
  addRulesFilteringTombstones,
  applyConfidenceCap,
  deleteRule,
  diffSnapshots,
  encodeCompound,
  extractSubjectSignal,
  findConflicts,
  listRules,
  buildRuleIndex,
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
  getMetrics,
  getRuleEvents,
  getRuleTombstones,
  getSettings,
  getSkipHistory,
  getSkipHistoryCount,
  getStorageUsage,
  getSubjectMemory,
  getUndoSnapshot,
  getWeeklyDigestState,
  mergeFolderActivityScan,
  recordRuleEvents,
  setFolderActivityRefreshAt,
  setFolderCache,
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
import { getOwaToken, peekCachedToken, pingOwa } from './token'

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

async function getOrFetchFolderTree(forceFresh: boolean): Promise<MailFolderNode[]> {
  const cache = await getFolderCache()
  const cacheFresh = cache && Date.now() - new Date(cache.updatedAt).getTime() < FOLDER_CACHE_TTL_MS

  if (!forceFresh && cacheFresh) {
    return cache!.tree
  }

  try {
    const tree = await api.fetchFolderTree()
    await setFolderCache({ updatedAt: new Date().toISOString(), tree })
    return tree
  } catch (e) {
    // Transient failures (network, 401, etc) shouldn't block the user when a
    // stale tree is still usable. Fall back, log so the staleness is traceable.
    if (cache) {
      console.warn(
        '[mail-organizer] folder tree fetch failed, falling back to stale cache',
        { ageMs: Date.now() - new Date(cache.updatedAt).getTime(), error: e },
      )
      return cache.tree
    }
    throw e
  }
}

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

// Storage usage check on SW startup. chrome.storage.local has a 5 MB cap
// and writes silently fail once exceeded. Log a clear warning at >80%
// so the user (or anyone inspecting the SW console) can act before
// rule writes start dropping.
void getStorageUsage()
  .then((usage) => {
    if (usage.approaching) {
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

type AnyRequest = { type: string; [k: string]: unknown }

chrome.runtime.onMessage.addListener(
  (msg: AnyRequest, _sender, sendResponse: (resp: PopupResponse) => void) => {
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

async function handle(msg: AnyRequest): Promise<PopupResponse> {
  switch (msg.type) {
    // ---- Production handlers (used by popup #6) ----------------------------

    case 'getStatus': {
      const [settings, cache, rules] = await Promise.all([
        getSettings(),
        getFolderCache(),
        listRules(),
      ])
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
            /* background refresh — silent on failure, classify will surface it */
          })
        }
      }
      return {
        ok: true,
        data: {
          owaConnected: owa.ok,
          tokenValid,
          apiKeyConfigured: settings.claudeApiKey.length > 0,
          apiKeyPreview: settings.claudeApiKey.length > 0 ? maskApiKey(settings.claudeApiKey) : null,
          model: settings.claudeModel,
          folderCacheAgeMs: cache ? Date.now() - new Date(cache.updatedAt).getTime() : null,
          folderCacheCount: cache ? flattenFolderTree(cache.tree).length : 0,
          rulesCount: rules.length,
          excludePrefixes: settings.excludeFolderPrefixes,
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
          syncEnabled: settings.syncEnabled,
          syncMachineId: settings.syncMachineId,
          lastSyncAt: settings.lastSyncAt,
        },
      }
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
      // Phase 3 — manual refresh via Graph API. Scans every case folder under
      // the user's root path, queries each for its most-recent message, and
      // merges results into the activity store. Catches mail the user dragged
      // in manually (bypassing the extension's classify flow) that otherwise
      // wouldn't appear in the quick-jump panel.
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
        // are case-level folders (e.g. 法顧/foodpanda). Skipping interior
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
        const batchSize =
          typeof msg.batchSize === 'number' ? msg.batchSize : settings.batchSize
        const [rawEmails, skipHistory] = await Promise.all([
          api.listInboxMessages({
            top: batchSize,
            // Id for skip-history match, Flag for skipFlagged filter — that's all.
            select: 'Id,Flag',
          }),
          getSkipHistory(),
        ])
        const skipIds = new Set(Object.keys(skipHistory))
        const afterSkipHistory =
          skipIds.size > 0 ? rawEmails.filter((m) => !skipIds.has(m.Id)) : rawEmails
        // Skip flagged mail is now hardcoded always-on (2026-05-27 reorg).
        // Lawyers use Outlook's follow-up flag to mark "I'm tracking this"
        // — auto-routing those was always a footgun. The UI toggle was
        // removed; the storage field is preserved for backward compat but
        // ignored here.
        const eligible = afterSkipHistory.filter(
          (m) => m.Flag?.FlagStatus !== 'Flagged',
        )
        return {
          ok: true,
          data: {
            eligibleCount: eligible.length,
            cappedAtBatchSize: rawEmails.length >= batchSize,
            totalFetched: rawEmails.length,
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
      const settings = await getSettings()
      if (!settings.claudeApiKey) {
        return { ok: false, code: 'NO_API_KEY', message: '請先到設定頁設定 Claude API key' }
      }
      const batchSize = typeof msg.batchSize === 'number' ? msg.batchSize : settings.batchSize
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
        const [rawEmails, skipHistory] = await Promise.all([
          api.listInboxMessages({ top: batchSize }),
          getSkipHistory(),
        ])

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
        const rules = await mutateRules<Rule[]>((current) => {
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

        // Pre-filter pass 1: drop emails the user already decided to keep in
        // inbox (skip history).
        const skipIds = new Set(Object.keys(skipHistory))
        const afterSkipHistory =
          skipIds.size > 0 ? rawEmails.filter((m) => !skipIds.has(m.Id)) : rawEmails
        const preFilteredCount = rawEmails.length - afterSkipHistory.length

        // Pre-filter pass 2: skip Outlook-flagged ("待處理") messages when
        // Skip flagged mail is hardcoded always-on (2026-05-27 reorg).
        // User is actively tracking these — don't auto-route them.
        //
        // IMPORTANT: this reads each email's live Flag.FlagStatus from the
        // Outlook fetch above — we do NOT persist a flagged-id list. When
        // the user removes the flag in Outlook, the very next classify run
        // sees FlagStatus !== 'Flagged' and includes the email automatically.
        const emails = afterSkipHistory.filter(
          (m) => m.Flag?.FlagStatus !== 'Flagged',
        )
        const flaggedCount = afterSkipHistory.length - emails.length

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
            if (e && (e.conflictCount ?? 0) === 0) {
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
            if (e && (e.conflictCount ?? 0) === 0) {
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
        // Build the rule index ONCE per batch. Before this, matchEmail
        // re-sorted the full rules array for every email — O(N log N)
        // per email × N emails = O(N² log N) of pure sort work that the
        // indexed path replaces with a single sort plus O(1) Map lookups
        // for domain / sender (the two largest buckets in practice).
        const ruleIndex = buildRuleIndex(rules)
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
          }
        }

        // Persist preflight context so classifyAi can run independently
        await chrome.storage.session.set({
          [PREFLIGHT_CACHE_KEY]: {
            unmatchedEmails: unmatched,
            tree,
            rulePlan,
            createdAt: Date.now(),
          },
        })

        return {
          ok: true,
          data: {
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
        await setClassifyStage({ stage: 'idle' })
      }
    }

    case 'classifyAi': {
      // Phase 2 — chunked, progressive AI classification.
      // Fire-and-forget: SW writes progress to session storage each chunk,
      // popup polls for incremental updates.
      const settings = await getSettings()
      const excludeIds = new Set((msg.excludeIds as string[] | undefined) ?? [])
      const pf = (await chrome.storage.session.get(PREFLIGHT_CACHE_KEY))[
        PREFLIGHT_CACHE_KEY
      ] as
        | {
            unmatchedEmails: Email[]
            tree: MailFolderNode[]
            rulePlan: PlanItem[]
            createdAt: number
          }
        | undefined
      if (!pf) {
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

      // Initial progress snapshot (rulePlan + folderTree embedded so popup can
      // resume from this state alone if it reopens mid-run)
      await chrome.storage.session.set({
        [AI_PROGRESS_KEY]: {
          totalEmails: toClassify.length,
          completedEmails: 0,
          chunks: chunks.length,
          completedChunks: 0,
          rulePlan: pf.rulePlan,
          folderTree: pf.tree,
          aiPlan: [],
          startedAt: Date.now(),
          done: chunks.length === 0,
        },
      })

      // Run the AI loop in the background; popup polls.
      void (async () => {
        let aiPlan: PlanItem[] = []
        let totalUsage: ClassifierUsage = {
          inputTokens: 0,
          outputTokens: 0,
          cacheCreationTokens: 0,
          cacheReadTokens: 0,
        }
        let aiError: { code: string; message: string } | null = null
        const threshold = settings.aiConfidenceThreshold
        let gatedCount = 0
        let truncatedCount = 0

        await holdKeepAlive()
        try {
          await setClassifyStage({ stage: 'calling_ai', toClassify: toClassify.length })

          // Pull current rules once per classify run so few-shot examples
          // reflect the user's latest state. Read fresh — preflight cache
          // might be from minutes ago and rules can change (manual edits,
          // ai_overridden generated by a previous batch).
          const currentRules = await listRules()

          for (let ci = 0; ci < chunks.length; ci++) {
            try {
              const result = await classifyBatch(
                {
                  emails: chunks[ci]!,
                  folderTree: pf.tree,
                  excludePrefixes: settings.excludeFolderPrefixes,
                  rules: currentRules,
                },
                settings,
              )
              if (result.truncated) {
                truncatedCount += chunks[ci]!.length - result.truncated.parsedCount
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
              // this step, `aiPlan` length matches the email count the user
              // expects (not the rep count).
              const fannedOut: PlanItem[] = []
              for (const repItem of chunkPlan) {
                fannedOut.push(repItem)
                const sibs = siblingsByRepId.get(repItem.emailId)
                if (sibs) {
                  for (const sib of sibs) {
                    fannedOut.push(inheritPlanItemForSibling(repItem, sib))
                  }
                }
              }
              aiPlan = [...aiPlan, ...fannedOut]
              totalUsage.inputTokens += result.usage.inputTokens
              totalUsage.outputTokens += result.usage.outputTokens
              totalUsage.cacheCreationTokens += result.usage.cacheCreationTokens
              totalUsage.cacheReadTokens += result.usage.cacheReadTokens
            } catch (e) {
              aiError =
                e instanceof ClassifierError
                  ? { code: `CLASSIFIER_${e.code}`, message: e.message }
                  : { code: 'AI_FAILED', message: e instanceof Error ? e.message : String(e) }
              // Skip remaining chunks; downstream backstop turns them to skip.
              // `chunks.slice(ci).flat()` *includes the chunk that just failed*
              // — dedup against aiPlan so successful entries from earlier in the
              // same chunk (or popup-pushed duplicates) don't get a phantom
              // "skip" stub piled on top, which would skew completedEmails
              // and downstream summary counts.
              const remaining = chunks.slice(ci).flat()
              const already = new Set(aiPlan.map((p) => p.emailId))
              const errReason = `AI 分類失敗：${aiError.code}`
              const skipStub = (m: Email): PlanItem => ({
                emailId: m.Id,
                emailSubject: m.Subject ?? '',
                emailFrom: m.From?.EmailAddress?.Address ?? '',
                bodyPreview: m.BodyPreview ?? '',
                conversationId: m.ConversationId,
                action: 'skip',
                confidence: 0,
                reason: errReason,
                source: 'unresolved',
              })
              for (const m of remaining) {
                if (already.has(m.Id)) continue
                aiPlan.push(skipStub(m))
                already.add(m.Id)
                // Siblings also need a skip stub or they vanish from the
                // plan entirely.
                const sibs = siblingsByRepId.get(m.Id)
                if (sibs) {
                  for (const sib of sibs) {
                    if (already.has(sib.Id)) continue
                    aiPlan.push(skipStub(sib))
                    already.add(sib.Id)
                  }
                }
              }
              break
            }

            // After each chunk, push partial plan to progress
            await chrome.storage.session.set({
              [AI_PROGRESS_KEY]: {
                totalEmails: toClassify.length,
                completedEmails: aiPlan.length,
                chunks: chunks.length,
                completedChunks: ci + 1,
                rulePlan: pf.rulePlan,
                folderTree: pf.tree,
                aiPlan,
                startedAt: Date.now(),
                done: false,
              },
            })
          }

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

          await chrome.storage.session.set({
            [AI_PROGRESS_KEY]: {
              totalEmails: toClassify.length,
              completedEmails: aiPlan.length,
              chunks: chunks.length,
              completedChunks: chunks.length,
              rulePlan: pf.rulePlan,
              folderTree: pf.tree,
              aiPlan: [...aiPlan, ...skipPlan],
              skippedByUser: skipPlan.length,
              usage: totalUsage,
              aiError: banner,
              startedAt: Date.now(),
              done: true,
            },
          })
        } finally {
          await setClassifyStage({ stage: 'idle' })
          // Preflight context no longer needed
          await chrome.storage.session.remove(PREFLIGHT_CACHE_KEY)
          await releaseKeepAlive()
        }
      })()

      return { ok: true, data: { started: true, chunks: chunks.length, totalEmails: toClassify.length } }
    }

    case 'getAiClassifyProgress': {
      const r = await chrome.storage.session.get(AI_PROGRESS_KEY)
      return { ok: true, data: r[AI_PROGRESS_KEY] ?? null }
    }

    case 'clearAiClassifyProgress': {
      await chrome.storage.session.remove(AI_PROGRESS_KEY)
      await chrome.storage.session.remove(PREFLIGHT_CACHE_KEY)
      return { ok: true }
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
      if (typeof raw.claudeModel === 'string' && raw.claudeModel.length > 0) {
        clean.claudeModel = raw.claudeModel
      }
      if (typeof raw.batchSize === 'number' && Number.isFinite(raw.batchSize)) {
        clean.batchSize = Math.max(1, Math.min(200, Math.floor(raw.batchSize)))
      }
      if (
        Array.isArray(raw.excludeFolderPrefixes) &&
        raw.excludeFolderPrefixes.every((s) => typeof s === 'string')
      ) {
        clean.excludeFolderPrefixes = raw.excludeFolderPrefixes
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
      if (await isExecuteRunning()) {
        return { ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' }
      }
      // Fire-and-forget; popup polls getExecuteState.
      void startExecute(plan, resolvedTree, api).catch((err) =>
        console.error('[mail-organizer] startExecute failed', err),
      )
      return { ok: true, data: { started: true, total: plan.length } }
    }

    case 'getExecuteState': {
      const state = await getExecuteState()
      return { ok: true, data: { state } }
    }

    case 'cancelExecute': {
      await requestCancel()
      return { ok: true }
    }

    case 'clearExecuteState': {
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
      if (await isExecuteRunning()) {
        return { ok: false, code: 'ALREADY_RUNNING', message: '已有批次執行中' }
      }
      const prev = await getExecuteState()
      if (!prev || prev.inProgress || !prev.plan) {
        return { ok: false, code: 'NO_PREV', message: '沒有可重試的批次' }
      }
      const failedItems: PlanItem[] = []
      for (let i = 0; i < prev.results.length; i++) {
        if (prev.results[i]?.status === 'error') {
          failedItems.push(prev.plan[i]!)
        }
      }
      if (failedItems.length === 0) {
        return { ok: false, code: 'NO_FAILED', message: '沒有失敗項目' }
      }
      const tree = (await getFolderCache())?.tree
      if (!tree) {
        return { ok: false, code: 'NO_TREE', message: '資料夾樹不存在，請先重新分類' }
      }
      // Discard the pending undo snapshot — its emails belong to the previous
      // batch. Otherwise the retry's DoneScreen could still expose the stale
      // "撤回" link pointing at the wrong batch's moves.
      await dismissUndo()
      await clearExecuteState()
      void startExecute(failedItems, tree, api).catch((err) =>
        console.error('[mail-organizer] retryFailed failed', err),
      )
      return { ok: true, data: { started: failedItems.length } }
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
      await upsertRule(rule, { actor: 'user' })
      return { ok: true, data: { rule } }
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
      return { ok: true, data: { rule: updated } }
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
      const payload = typeof msg.payload === 'string' ? msg.payload : ''
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

    case 'resolveConflict': {
      // strategy:
      //   'keep_highest' — keep the most-confident rule, DELETE the rest
      //                    (writes tombstones — losers won't auto-resurrect)
      //   'disable_all'  — disable every rule in this conflict group
      //                    (reversible; for "park it for now" intent)
      const ruleIds = Array.isArray(msg.ruleIds) ? (msg.ruleIds as string[]) : []
      const strategy = msg.strategy as 'keep_highest' | 'disable_all'
      if (ruleIds.length === 0 || !['keep_highest', 'disable_all'].includes(strategy)) {
        return { ok: false, code: 'INVALID_INPUT', message: '參數錯誤' }
      }
      const all = await listRules()
      const group = all.filter((r) => ruleIds.includes(r.id))
      if (group.length === 0) {
        return { ok: false, code: 'NOT_FOUND', message: '找不到指定規則' }
      }

      let keepId: string | null = null
      if (strategy === 'keep_highest') {
        const sorted = [...group].sort((a, b) => b.confidence - a.confidence)
        keepId = sorted[0]!.id
      }

      let disabled = 0
      let deleted = 0
      for (const r of group) {
        if (strategy === 'disable_all') {
          if (!r.enabled) continue
          await toggleRule(r.id, false, { actor: 'user' })
          disabled++
        } else {
          // keep_highest: loser rules get deleted (NOT disabled). Previously
          // this path only flipped enabled=false — leading to "zombie" rules
          // staying in the list as greyed-out entries, indistinguishable
          // from manual disables. Delete writes a tombstone so the same
          // (type, signal, target) won't auto-regenerate next AI batch.
          if (r.id === keepId) continue
          const ok = await deleteRule(r.id, { actor: 'user' })
          if (ok) deleted++
        }
      }
      return { ok: true, data: { disabled, deleted, kept: keepId } }
    }

    case 'clearDisabledRules': {
      // Bulk-delete every rule with enabled=false. Goes through deleteRule so
      // each one writes a tombstone — AI / initial-scan won't recreate them
      // with the same (type, signal, target) on the next pass.
      const all = await listRules()
      const disabledIds = all.filter((r) => !r.enabled).map((r) => r.id)
      let deleted = 0
      for (const id of disabledIds) {
        const ok = await deleteRule(id, { actor: 'user' })
        if (ok) deleted++
      }
      return { ok: true, data: { deleted } }
    }

    case 'runStaleSweep': {
      // Manual trigger for the daily auto-disable sweep. Same operation
      // as the chrome.alarms-driven path. Returns the count of rules
      // auto-disabled this run.
      const { disabledCount } = await runSweep()
      return { ok: true, data: { disabledCount } }
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

    case 'splitConflictToCompound': {
      // Each branch can supply 0..N keywords:
      //   - empty (0 kw)   → that rule stays enabled as a "fallback" (matches
      //                       when the simple condition holds but no keyword)
      //   - one or more   → create one compound per keyword (OR semantics for
      //                       multiple keywords targeting same folder), disable
      //                       the original
      // Validation: at most ONE branch may be empty (otherwise conflict remains);
      // all non-empty keywords must be unique across all branches.
      const items = Array.isArray(msg.items)
        ? (msg.items as Array<{ ruleId: string; keywords: string[] }>)
        : []
      if (items.length < 2) {
        return { ok: false, code: 'INVALID_INPUT', message: '至少要 2 條規則參與拆解' }
      }

      const sanitized = items.map((i) => ({
        ruleId: i.ruleId,
        keywords: (i.keywords ?? [])
          .map((k) => k.trim())
          .filter((k) => k.length > 0),
      }))

      const fallbackCount = sanitized.filter((s) => s.keywords.length === 0).length
      if (fallbackCount > 1) {
        return {
          ok: false,
          code: 'TOO_MANY_FALLBACKS',
          message: '最多一個分支可留空作為 fallback',
        }
      }

      const allKeywords = sanitized.flatMap((s) =>
        s.keywords.map((k) => k.toLowerCase()),
      )
      const distinctKeywords = new Set(allKeywords)
      if (distinctKeywords.size !== allKeywords.length) {
        return {
          ok: false,
          code: 'DUPLICATE_KEYWORDS',
          message: '同一關鍵字不可在多個分支使用',
        }
      }

      const all = await listRules()
      const newRules: Rule[] = []
      let disabled = 0
      let keptAsFallback = 0

      for (const branch of sanitized) {
        const original = all.find((r) => r.id === branch.ruleId)
        if (!original) continue

        if (branch.keywords.length === 0) {
          // Fallback branch — leave the original rule enabled, no compounds
          keptAsFallback++
          continue
        }

        // Extract base conditions from the original rule
        let baseConditions: Array<{
          type: 'domain' | 'subject_keyword' | 'sender'
          value: string
        }> = []
        if (original.type === 'compound') {
          const parsed = (() => {
            try {
              return JSON.parse(original.signal) as { conditions: typeof baseConditions }
            } catch {
              return null
            }
          })()
          baseConditions = parsed?.conditions ?? []
        } else if (
          original.type === 'domain' ||
          original.type === 'subject_keyword' ||
          original.type === 'sender'
        ) {
          baseConditions = [{ type: original.type, value: original.signal }]
        } else {
          // case_code can't be folded into compound — skip
          continue
        }

        for (const kw of branch.keywords) {
          const conditions = [...baseConditions, { type: 'subject_keyword' as const, value: kw }]
          newRules.push(
            newRule({
              type: 'compound',
              signal: JSON.stringify({ conditions }),
              targetFolderId: original.targetFolderId,
              targetFolderPath: original.targetFolderPath,
              confidence: original.confidence,
              source: 'user_manual',
            }),
          )
        }

        if (original.enabled) {
          await toggleRule(original.id, false, { actor: 'user' })
          disabled++
        }
      }

      for (const r of newRules) await upsertRule(r, { actor: 'user' })
      return {
        ok: true,
        data: { created: newRules.length, disabled, keptAsFallback },
      }
    }

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

    case 'exportDiagnostic': {
      const [settings, cache, rules, metrics, skipCount, storageUsage] = await Promise.all([
        getSettings(),
        getFolderCache(),
        listRules(),
        getMetrics(),
        getSkipHistoryCount(),
        getStorageUsage(),
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
            quotaBytes: storageUsage.quotaBytes,
            percentUsed: Math.round(storageUsage.percentUsed * 100) / 100,
            approaching: storageUsage.approaching,
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

    default:
      return {
        ok: false,
        code: 'UNKNOWN_REQUEST',
        message: `Unknown request type: ${msg.type}`,
      }
  }
}

function maskApiKey(key: string): string {
  if (key.length <= 14) return '••••'
  return `${key.slice(0, 10)}••••${key.slice(-4)}`
}
