import type { PlanItem, Rule, Metrics } from './types'

// popup → background ---------------------------------------------------------

export type PopupRequest =
  | { type: 'classify'; batchSize?: number }
  | { type: 'execute'; plan: PlanItem[] }
  | { type: 'cancel' }
  | { type: 'getRules' }
  | { type: 'upsertRule'; rule: Rule }
  | { type: 'deleteRule'; ruleId: string }
  | { type: 'getMetrics' }
  | { type: 'refreshFolderCache' }
  | { type: 'pingOwa' }
  | { type: 'initialScan' }

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
  | { ok: false; code: 'NO_TOKEN' | 'EXPIRED' | 'NOT_ON_OWA'; message: string }

// channel names --------------------------------------------------------------

export const PORT_PROGRESS = 'mail-organizer:progress'
