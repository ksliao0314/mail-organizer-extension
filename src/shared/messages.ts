import type { PlanItem, Metrics } from './types'

// popup / options → background ----------------------------------------------

/**
 * Every request `type` the service-worker dispatcher (`handle()` in
 * background/service-worker.ts) serves (2026-11-05).
 *
 * Why a single list: UI↔SW drift (the export "undefined" file bug 75e436e,
 * the import `json`/`payload` bug) used to be invisible until runtime because
 * the SW switched on an untyped `string` and every sender passed `unknown`.
 * Now `handle()` takes `{ type: MessageType }` and its switch is exhaustive
 * (a `never` check in `default`), so tsc fails when
 *   - a `case` label is not listed here (typo / undeclared handler), or
 *   - a type listed here has no `case`.
 * `send()` in shared/send-message.ts takes the same request type, so typed
 * senders can't name a type the SW doesn't serve. tests/ui-sw-contract.test.ts
 * additionally scans the UI sources (whose local `send()` copies are untyped).
 *
 * To add a message: append it to the matching group below AND add the `case`.
 */
export const SW_MESSAGE_TYPES = [
  // ---- content scripts (the ONLY types a non-extension-page sender may use;
  //      allowlist enforced in the dispatcher) ----
  'getOwaFabVisibility',
  'issuePanelPairingNonce',

  // ---- in-OWA panel bridge pairing (extension pages only; 2026-11-06) ----
  'redeemPanelPairingNonce',

  // ---- status / navigation ----
  'getStatus',
  'openOwa',
  'getOnboardingState',
  'navigateToFolder',
  'detectUserDomain',
  'getStorageUsage',
  'exportDiagnostic',
  'getMetrics',

  // ---- recent activity / weekly digest / harvest ----
  'getFolderActivity',
  'getFolderActivityFreshness',
  'refreshFolderActivity',
  'harvestManualFilings',
  'getWeeklyDigest',
  'dismissWeeklyDigest',

  // ---- classify pipeline ----
  'peekNextBatch',
  'classifyPreflight',
  'classifyAi',
  'getAiClassifyProgress',
  'clearAiClassifyProgress',
  'cancelClassify',
  'adoptClassifyRun',
  'getClassifyStage',
  'getFolderTree',
  'savePopupState',
  'getPopupState',
  'clearPopupState',
  'getSkipHistoryCount',
  'clearSkipHistory',

  // ---- settings ----
  'setApiKey',
  'setSettings',

  // ---- execute / undo ----
  'startExecute',
  'getExecuteState',
  'cancelExecute',
  'clearExecuteState',
  'getUndoSnapshot',
  'executeUndo',
  'dismissUndo',
  'retryFailed',

  // ---- initial scan ----
  'startInitialScan',
  'getScanState',
  'cancelScan',
  'clearScanState',

  // ---- rule library ----
  'getRules',
  'createRule',
  'upsertRule',
  'deleteRule',
  'toggleRule',
  'bulkRuleAction',
  'clearAllRules',
  'wipeAllRules',
  'exportRules',
  'previewRuleImport',
  'importRules',
  'getRuleHistory',
  'clearRuleHistory',
  'resolveConflict',
  'clearDisabledRules',
  'runStaleSweep',
  'suggestSplitKeywords',
  'splitConflictToCompound',
  'autoUpgradeConflictRules',

  // ---- cross-machine sync ----
  'getSyncStatus',
  'enableSync',
  'disableSync',
  'pushSyncNow',
  'pullSyncNow',
  'listSyncBackups',
  'restoreSyncBackup',
  'dismissSyncError',
  'getRemoteWipeNotice',
  'dismissRemoteWipeNotice',

  // ---- error log ----
  'getErrorLog',
  'clearErrorLog',
] as const

export type MessageType = (typeof SW_MESSAGE_TYPES)[number]

/**
 * The ONLY request types a sender that is not one of our extension pages —
 * i.e. a content script in an Outlook tab — may use (2026-11-05). Everything
 * else from such a sender is refused with FORBIDDEN by the dispatcher.
 * A NEW `chrome.runtime.sendMessage` in src/content/* must be added here, or
 * it will be refused in production; tests/ui-sw-contract.test.ts scans the
 * content scripts to catch exactly that.
 */
export const CONTENT_SCRIPT_MESSAGE_TYPES: readonly MessageType[] = [
  'getOwaFabVisibility',
  // One-time nonce the FAB puts in its iframe handshake (2026-11-06) — see
  // background/handlers/panel-pairing.ts.
  'issuePanelPairingNonce',
]

/** A request to the service worker. Handler-specific fields ride alongside `type`. */
export type PopupRequest = { type: MessageType; [k: string]: unknown }

export type PopupResponse =
  | { ok: true; data?: unknown }
  | { ok: false; code: string; message: string }

// background → popup (long-running progress, pushed via runtime.connect) ----

export type ProgressStage = 'fetching' | 'classifying' | 'executing' | 'scanning'

export type BackgroundEvent =
  | { type: 'progress'; stage: ProgressStage; current: number; total: number; note?: string }
  | { type: 'plan_ready'; items: PlanItem[] }
  | { type: 'item_done'; emailId: string; status: 'moved' | 'deleted' | 'folder_created' | 'error'; message?: string }
  | { type: 'done'; metrics: Metrics }
  | { type: 'error'; code: string; message: string }

// background ↔ content -------------------------------------------------------

export type ContentRequest = { type: 'fetch_token' }

export type ContentResponse =
  | { ok: true; secret: string; expiresOn: number }
  | {
      ok: false
      code: 'NO_TOKEN' | 'EXPIRED' | 'NOT_ON_OWA' | 'MULTIPLE_ACCOUNTS'
      message: string
    }

// channel names --------------------------------------------------------------

export const PORT_PROGRESS = 'mail-organizer:progress'
