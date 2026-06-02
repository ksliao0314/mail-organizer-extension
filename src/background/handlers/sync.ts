// SW request handlers for cross-machine sync + centralised error log.
//
// Extracted from service-worker.ts to keep the dispatcher manageable.
// Each export below is a small async function taking the parsed message
// fragment it cares about, returning a PopupResponse. The dispatcher in
// service-worker.ts pattern-matches on msg.type and forwards.
//
// Why extract: sync has been the most actively modified subsystem (17+
// bug fixes since launch) and lives as a 150-line block of switch cases
// in the middle of a 2,000-line file. Isolating means future sync
// changes don't churn the dispatcher; sync tests can stub at this
// boundary; the dispatcher reads as routing, not implementation.

import {
  clearCloudState,
  dismissSyncError,
  getSyncStatus,
  listBackups,
  pullNow,
  pushNow,
  quiesce,
  restoreBackup,
} from '../sync-engine'
import { clearErrorLog, getErrorLog } from '@/shared/error-log'
import { setSettings } from '@/shared/storage'
import type { PopupResponse } from '@/shared/messages'

export async function handleGetSyncStatus(): Promise<PopupResponse> {
  const status = await getSyncStatus()
  return { ok: true, data: status }
}

/**
 * First-time enable. Flow:
 *   1. Mark sync enabled (settings.syncEnabled = true)
 *   2. If cloud already has data and it's NOT ours (different machineId):
 *      a. UNION-pull (preserve our local rules cloud didn't have)
 *      b. then push merged result so cloud now has the union
 *   3. Else just push.
 *
 * Bug #M: surface real failure. If pull or push-back fails, roll back
 * syncEnabled and return ok:false — otherwise the UI would show "啟用
 * 成功" green while sync is actually broken.
 */
export async function handleEnableSync(): Promise<PopupResponse> {
  await setSettings({ syncEnabled: true })
  const status = await getSyncStatus()
  let action: 'push' | 'pull' | 'union-merge' = 'push'
  let pull: Awaited<ReturnType<typeof pullNow>> | undefined
  let push: Awaited<ReturnType<typeof pushNow>> | undefined
  if (status.cloud && !status.cloud.isUs) {
    action = 'union-merge'
    pull = await pullNow('first-enable', { mode: 'union' })
    // Don't attempt push-back if the pull itself failed — we'd be
    // pushing pre-merge local state which could lose data.
    if (pull.pulled) {
      push = await pushNow('first-enable-push-back')
    }
  } else {
    push = await pushNow('first-enable')
  }
  const ok =
    action === 'union-merge'
      ? pull?.pulled === true && push?.pushed === true
      : push?.pushed === true
  if (!ok) {
    // Roll back the syncEnabled flip — sync didn't actually start, so
    // leaving syncEnabled=true would mean every subsequent local
    // mutation schedules a doomed push.
    await setSettings({ syncEnabled: false })
    const reason =
      pull && !pull.pulled
        ? `下載失敗:${pull.reason ?? '未知'}`
        : push && !push.pushed
          ? `上傳失敗:${push.reason ?? '未知'}`
          : '同步失敗'
    return { ok: false, code: 'ENABLE_SYNC_FAILED', message: reason }
  }
  return {
    ok: true,
    data: {
      action,
      pulled: pull?.pulled,
      pushed: push?.pushed,
      ruleCount: pull?.ruleCount,
      tombstoneCount: pull?.tombstoneCount,
      mergeMode: pull?.mergeMode,
    },
  }
}

/**
 * Disable on this machine. keepCloud=true (default) keeps cloud state
 * intact so other machines can keep syncing. keepCloud=false wipes
 * cloud — but only after quiesce()ing any in-flight push (Bug #N).
 */
export async function handleDisableSync(msg: {
  keepCloud?: boolean
}): Promise<PopupResponse> {
  const keepCloud = msg.keepCloud !== false
  await setSettings({ syncEnabled: false })
  if (!keepCloud) {
    await quiesce()
    await clearCloudState()
  }
  return { ok: true, data: { keepCloud } }
}

export async function handlePushSyncNow(): Promise<PopupResponse> {
  const r = await pushNow('manual')
  return { ok: true, data: r }
}

export async function handlePullSyncNow(): Promise<PopupResponse> {
  const r = await pullNow('manual')
  return { ok: true, data: r }
}

export async function handleListSyncBackups(): Promise<PopupResponse> {
  const backups = await listBackups()
  // Don't ship the full payloads — let the UI summarise and only fetch
  // detail when the user actually clicks rollback.
  const summary = backups.map((b) => ({
    snapshotAt: b.snapshotAt,
    direction: b.direction,
    ruleCount: b.payload.rules.length,
    tombstoneCount: b.payload.tombstones.length,
  }))
  return { ok: true, data: { backups: summary } }
}

export async function handleRestoreSyncBackup(msg: {
  snapshotAt?: unknown
}): Promise<PopupResponse> {
  const snapshotAt = typeof msg.snapshotAt === 'string' ? msg.snapshotAt : ''
  if (!snapshotAt) {
    return { ok: false, code: 'BAD_INPUT', message: '缺少 snapshotAt' }
  }
  const r = await restoreBackup(snapshotAt)
  return { ok: true, data: r }
}

/** User-acknowledged dismissal of the persistent sync-error banner. */
export async function handleDismissSyncError(): Promise<PopupResponse> {
  await dismissSyncError()
  return { ok: true }
}

// ---- error log -----------------------------------------------------------

export async function handleGetErrorLog(msg: {
  limit?: unknown
}): Promise<PopupResponse> {
  const limit =
    typeof msg.limit === 'number' && msg.limit > 0 ? msg.limit : 100
  const entries = await getErrorLog(limit)
  return { ok: true, data: { entries } }
}

export async function handleClearErrorLog(): Promise<PopupResponse> {
  await clearErrorLog()
  return { ok: true }
}
