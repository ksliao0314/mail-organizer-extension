// Daily stale-rule sweep.
//
// Once a day, walk the rules array and auto-disable rules that:
//   - have never matched (matchCount === 0)
//   - were created more than 90 days ago
//   - are not user_manual (sacred)
//   - are not already disabled / orphaned
//
// Reversible — `enabled` flips to false, `autoDisabledAt` timestamp set.
// Rule stays in storage; user can re-enable in options.
//
// SYNC-SAFETY INVARIANT for the HARD-DELETE paths in this sweep (stale
// deletion, case_closed purge — no tombstone written): a no-tombstone
// delete is only convergent across machines because (a) the predicate is
// a pure function of SYNCED rule fields, and (b) this sweep runs on EVERY
// machine (each SW startup + daily alarm). Each machine deletes the same
// rules itself, so a union-mode sync pull can resurrect a copy for at
// most one sweep interval. NEVER hard-delete from an event-triggered,
// single-machine code path (preflight, message handler) without a
// tombstone — that resurrects forever via the other machine's cloud copy.
// Event paths must go through a soft-disable that propagates via sync
// (see applyClosedCaseRetirement Phase 1), leaving the delete to us.
//
// Background-only: needs chrome.alarms which only the service worker can
// register. Wires up in service-worker.ts via installStaleSweepListener().

import {
  applyClosedCaseRetirement,
  autoDisableStaleRules,
  dedupeRulesByKey,
  purgeExpiredClosedCaseRules,
} from '@/shared/rules'
import { getSettings } from '@/shared/storage'

const STALE_SWEEP_ALARM = 'stale-rule-sweep'
const SWEEP_INTERVAL_MINUTES = 24 * 60 // daily

/**
 * Register the daily alarm + handler. Idempotent — safe to call from
 * `runtime.onInstalled` AND `runtime.onStartup` AND on module load
 * (covers SW restart scenarios).
 *
 * `chrome.alarms.create` with the same name overwrites the previous
 * schedule, so this won't pile up duplicates.
 */
export function installStaleSweepListener(): void {
  chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name !== STALE_SWEEP_ALARM) return
    void runSweep()
  })
  void chrome.alarms.create(STALE_SWEEP_ALARM, {
    // Run first sweep ~5 min after register so a freshly-installed
    // extension doesn't sweep immediately (give user time to settle).
    delayInMinutes: 5,
    periodInMinutes: SWEEP_INTERVAL_MINUTES,
  })
  // Also run an immediate, fire-and-forget sweep on every SW startup —
  // makes legacy_token / lastUsedAt cleanup visible on extension reload
  // instead of waiting for the 5-minute alarm. The sweep is idempotent
  // (already-disabled rules are skipped), so running on every wake is
  // cheap.
  void runSweep()
}

/**
 * Public entry for manual invocation (e.g. "run sweep now" button in
 * options page). Same operation as the alarm-triggered path.
 */
export async function runSweep(): Promise<{ disabledCount: number; staleDeletedCount: number }> {
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
