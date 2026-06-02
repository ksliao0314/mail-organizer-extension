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
// Background-only: needs chrome.alarms which only the service worker can
// register. Wires up in service-worker.ts via installStaleSweepListener().

import { autoDisableStaleRules, dedupeRulesByKey } from '@/shared/rules'

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
export async function runSweep(): Promise<{ disabledCount: number }> {
  try {
    // Dedupe first — collapsing same-triplet duplicates before the stale
    // check lets stats (matchCount etc.) on the kept rule reflect the
    // full history of all dups merged into it. Hard delete, no audit.
    await dedupeRulesByKey()
    const { disabled } = await autoDisableStaleRules()
    // Only log soft-disable events (legacy_token / high-error-rate).
    // Stale deletions are intentionally silent per the design: they
    // leave no audit trail, no tombstone, no console log — the rule
    // library just gets cleaner.
    if (disabled.length > 0) {
      console.info(
        `[mail-organizer] stale sweep: auto-disabled ${disabled.length} rule(s)`,
        disabled.map((r) => `${r.type}::${r.signal}`),
      )
    }
    return { disabledCount: disabled.length }
  } catch (e) {
    console.warn('[mail-organizer] stale sweep failed (non-fatal)', e)
    return { disabledCount: 0 }
  }
}
