// Daily stale-rule sweep.
//
// Once a day: dedupe, closed-case retirement/purge, then
// autoDisableStaleRules (rules.ts) — soft-disables high-error-rate /
// legacy_token rules and HARD-deletes stale ones. Staleness is counted in
// ACTIVE days (UTC days a batch was executed), never calendar days, so an
// idle extension doesn't age its library (2026-11-05). user_manual,
// disabled and orphaned rules are never touched by the stale branch.
//
// SYNC-SAFETY INVARIANT for the HARD-DELETE paths in this sweep (stale
// deletion, case_closed purge — no tombstone written): a no-tombstone
// delete is only convergent across machines because (a) the predicate is
// a function of SYNCED rule fields plus the active-day clock — and that
// clock is mostly derived from the synced rules themselves
// (deriveActiveDaysFromRules), with each machine's own recorded days on
// top — and (b) this sweep runs on EVERY machine (daily alarm + startup
// catch-up). Each machine deletes (nearly) the same rules itself, so a
// union-mode sync pull can resurrect a copy for at most a sweep interval or
// two; a machine whose clock lags only ever keeps a rule LONGER. NEVER hard-delete from an event-triggered,
// single-machine code path (preflight, message handler) without a
// tombstone — that resurrects forever via the other machine's cloud copy.
// Event paths must go through a soft-disable that propagates via sync
// (see applyClosedCaseRetirement Phase 1), leaving the delete to us.
//
// Background-only: needs chrome.alarms which only the service worker can
// register. Wires up in service-worker.ts via installStaleSweepListener().

import { SWEEP_MIN_INTERVAL_MS } from '@/shared/constants'
import {
  applyClosedCaseRetirement,
  autoDisableStaleRules,
  dedupeRulesByKey,
  purgeExpiredClosedCaseRules,
} from '@/shared/rules'
import { getLastSweepAt, getSettings, setLastSweepAt } from '@/shared/storage'

const STALE_SWEEP_ALARM = 'stale-rule-sweep'
const SWEEP_INTERVAL_MINUTES = 24 * 60 // daily

export type SweepResult = {
  disabledCount: number
  staleDeletedCount: number
  /** True when a non-forced call was skipped because a sweep ran < SWEEP_MIN_INTERVAL_MS ago. */
  skipped?: boolean
}

/**
 * Register the daily alarm + handler. Safe to call on every SW start
 * (module load).
 *
 * The alarm is created only when it doesn't already exist (2026-11-05).
 * Previously every SW start re-created it with delayInMinutes:5, and
 * create() with the same name REPLACES the pending alarm — since the SW
 * wakes on most tab events, the "daily" sweep actually fired ~every 5 min
 * (twice per wake: module load + onAlarm). The startup sweep is now gated
 * by lastSweepAt inside runSweep.
 */
export function installStaleSweepListener(): void {
  chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name !== STALE_SWEEP_ALARM) return
    void runSweep()
  })
  void ensureSweepAlarm()
  // Catch-up sweep on SW startup (e.g. browser was closed when the alarm was
  // due). Gated: a no-op unless the last sweep is ≥ SWEEP_MIN_INTERVAL_MS old.
  void runSweep()
}

async function ensureSweepAlarm(): Promise<void> {
  try {
    if (await chrome.alarms.get(STALE_SWEEP_ALARM)) return
    await chrome.alarms.create(STALE_SWEEP_ALARM, {
      // First sweep ~5 min after (re)registration so a freshly-installed
      // extension doesn't sweep immediately (give the user time to settle).
      delayInMinutes: 5,
      periodInMinutes: SWEEP_INTERVAL_MINUTES,
    })
  } catch (e) {
    console.warn('[mail-organizer] stale sweep alarm setup failed (non-fatal)', e)
  }
}

// Coalesces overlapping calls (module-load sweep + the alarm's own wake fire
// together). Assigned synchronously before any await — TOCTOU invariant.
let sweepInFlight: Promise<SweepResult> | null = null

/**
 * Entry for the alarm, SW startup, and the options "run sweep now" button
 * (`force: true`, which bypasses the interval gate).
 */
export function runSweep(opts: { force?: boolean } = {}): Promise<SweepResult> {
  if (sweepInFlight) return sweepInFlight
  sweepInFlight = runSweepGated(opts.force === true).finally(() => {
    sweepInFlight = null
  })
  return sweepInFlight
}

async function runSweepGated(force: boolean): Promise<SweepResult> {
  if (!force) {
    try {
      const last = await getLastSweepAt()
      if (last !== null && Date.now() - last < SWEEP_MIN_INTERVAL_MS) {
        return { disabledCount: 0, staleDeletedCount: 0, skipped: true }
      }
    } catch {
      /* unreadable stamp → just sweep */
    }
  }
  const result = await sweepOnce()
  await setLastSweepAt(Date.now()).catch(() => {})
  return result
}

async function sweepOnce(): Promise<SweepResult> {
  try {
    // Dedupe first — collapsing same-triplet duplicates before the stale
    // check lets stats (matchCount etc.) on the kept rule reflect the
    // full history of all dups merged into it. Hard delete, no audit.
    await dedupeRulesByKey()
    // 結案退場 (2026-07-16): Phase 1 best-effort catch-up on stored paths
    // (primary hook is classify preflight, right after reconcile refreshes
    // paths — this covers machines that never open the popup), then Phase 2
    // purge of rules retired ≥ CASE_CLOSED_PURGE_DAYS ago. Order matters:
    // retire/revive BEFORE purging so a folder moved back out of the
    // archive revives instead of aging toward deletion.
    try {
      const settings = await getSettings()
      const retirement = await applyClosedCaseRetirement(settings.closedCasePrefixes ?? [])
      const { purged } = await purgeExpiredClosedCaseRules()
      if (retirement.retired > 0 || retirement.revived > 0 || purged > 0) {
        console.info(
          `[mail-organizer] closed-case sweep: ${retirement.retired} retired, ` +
            `${retirement.revived} revived, ${purged} purged`,
        )
      }
    } catch (e) {
      console.warn('[mail-organizer] closed-case sweep failed (non-fatal)', e)
    }
    const { disabled, byReason } = await autoDisableStaleRules()
    // Only log soft-disable events (legacy_token / high-error-rate).
    // Stale deletions are intentionally silent per the design: they
    // leave no audit trail, no tombstone, no console log — the rule
    // library just gets cleaner. The COUNT is still returned so the
    // manual "run sweep now" button can tell the user what happened
    // (audit: reporting only disabledCount made a stale-deleting sweep
    // look like a no-op).
    if (disabled.length > 0) {
      console.info(
        `[mail-organizer] stale sweep: auto-disabled ${disabled.length} rule(s)`,
        disabled.map((r) => `${r.type}::${r.signal}`),
      )
    }
    return { disabledCount: disabled.length, staleDeletedCount: byReason.stale }
  } catch (e) {
    console.warn('[mail-organizer] stale sweep failed (non-fatal)', e)
    return { disabledCount: 0, staleDeletedCount: 0 }
  }
}
