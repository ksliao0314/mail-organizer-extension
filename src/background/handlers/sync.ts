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
  resetSyncParticipation,
  restoreBackup,
} from '../sync-engine'
import { clearErrorLog, getErrorLog } from '@/shared/error-log'
import { SYNC_SCHEMA_VERSION } from '@/shared/constants'
import { setSettings } from '@/shared/storage'
import type { PopupResponse } from '@/shared/messages'

export async function handleGetSyncStatus(): Promise<PopupResponse> {
  const status = await getSyncStatus()
  return { ok: true, data: status }
}

/**
 * User-facing reason for a failed enable (2026-11-05). The raw engine reason
 * for a schema refusal read "unsupported cloud schemaVersion=2 (we support
 * 3)" and the UI turned it into "upgrade THIS machine" — wrong when the cloud
 * is the OLDER side. Older clouds are now read and rewritten; only a NEWER
 * cloud is refused, and that really does need this machine upgraded.
 */
function describeEnableFailure(kind: 'pull' | 'push', reason: string | undefined): string {
  const r = reason ?? '未知'
  const newer = /schemaVersion=(\d+)/.exec(r)
  if (newer && Number(newer[1]) > SYNC_SCHEMA_VERSION) {
    return (
      `雲端同步資料由較新版本的擴充功能寫入(格式 v${newer[1]})、這台機器(格式 v${SYNC_SCHEMA_VERSION})無法讀取。` +
      '請先把這台機器的擴充功能更新到最新版本、再啟用同步。'
    )
  }
  if (/incomplete/i.test(r)) {
    return '雲端資料正在由另一台機器上傳中、尚未完整。請稍候幾秒再按一次「啟用同步」。'
  }
  return `${kind === 'pull' ? '下載失敗' : '上傳失敗'}:${r}`
}

/**
 * First-time enable. Flow:
 *   1. Mark sync enabled (settings.syncEnabled = true)
 *   2. If cloud already has data and it's NOT ours (different machineId):
 *      a. UNION-pull (preserve our local rules cloud didn't have) — any
 *         readable schema (v1–v4); a local backup is taken first, since
 *         this is the biggest merge a library ever goes through
 *      b. then push merged result so cloud now has the union (written as
 *         the current schema, upgrading an older cloud)
 *   3. Else just push (which itself pulls first if the cloud changed).
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
    pull = await pullNow('first-enable', { mode: 'union', backup: true })
    if (!pull.pulled && pull.torn) {
      // Another machine is mid-push; its chunks usually land within a second.
      await new Promise((r) => setTimeout(r, 1500))
      pull = await pullNow('first-enable-retry', { mode: 'union', backup: true })
    }
    if (!pull.pulled && pull.reason === 'no cloud state') {
      // Cloud was cleared between the status read and the pull.
      action = 'push'
      pull = undefined
      push = await pushNow('first-enable')
    } else if (pull.pulled) {
      // Don't attempt push-back if the pull itself failed — we'd be
      // pushing pre-merge local state which could lose data.
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
        ? describeEnableFailure('pull', pull.reason)
        : push && !push.pushed
          ? describeEnableFailure('push', push.reason)
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
  // Re-enabling later must count as FIRST participation (2026-11-05): a wipe
  // marker still in the cloud then is recorded as seen, never applied to the
  // library this machine builds while sync is off.
  await resetSyncParticipation()
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
  // UNION, never replace (2026-11-05). Replace drops every syncable local
  // rule that isn't in the cloud — but doPush's byte budget deliberately
  // leaves the lowest-priority rules OUT of the cloud ("they stay in local
  // storage"), so 立即下載 (and the pull-error banner's retry, which sends
  // the same message) permanently deleted exactly the rules whose only copy
  // was on this machine, disabled user_manual rules included. Upstream
  // deletions still propagate in union mode via tombstones.
  const r = await pullNow('manual', { mode: 'union' })
  return { ok: true, data: r }
}

export async function handleListSyncBackups(): Promise<PopupResponse> {
  const backups = await listBackups()
  // Don't ship the full payloads — let the UI summarise and only fetch
  // detail when the user actually clicks rollback.
  // Tombstones are no longer snapshotted (2026-11-05) — count defensively so
  // old and new backups both summarise.
  const summary = backups.map((b) => ({
    snapshotAt: b.snapshotAt,
    direction: b.direction,
    trigger: b.trigger,
    pinned: b.pinned === true,
    ruleCount: Array.isArray(b.payload?.rules) ? b.payload.rules.length : 0,
    tombstoneCount: Array.isArray(b.payload?.tombstones) ? b.payload.tombstones.length : 0,
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
