// MV3 service workers get idle-killed after ~30s of no events. Our execute
// loop has sleeps (token bucket) and Anthropic fetches that can take many
// seconds — during those gaps Chrome may decide to unload us.
//
// chrome.alarms fires events that count as "activity", keeping the SW alive.
// We schedule a tight-cadence alarm only while long-running work is
// happening, and clear it when idle.

const KEEPALIVE_ALARM_NAME = 'mail-organizer:keepalive'
// 0.5 min = Chrome's minimum alarm period for packed extensions (2026-11-05).
// The old 0.4 was below it — Chrome silently rounds it up to 30 s (with a
// warning), so the "24 s, under the idle window" this used to promise never
// held. 30 s is still enough as a backstop: the execute loop's own
// chrome.storage / fetch activity also resets the worker's idle timer.
export const KEEPALIVE_PERIOD_MINUTES = 0.5

// Reference count so concurrent long-running operations don't fight each
// other over the alarm lifecycle.
let activeHolders = 0

export async function holdKeepAlive(): Promise<void> {
  activeHolders++
  if (activeHolders === 1) {
    try {
      await chrome.alarms.create(KEEPALIVE_ALARM_NAME, { periodInMinutes: KEEPALIVE_PERIOD_MINUTES })
    } catch (e) {
      console.warn('[mail-organizer] keep-alive create failed', e)
    }
  }
}

export async function releaseKeepAlive(): Promise<void> {
  activeHolders = Math.max(0, activeHolders - 1)
  if (activeHolders === 0) {
    try {
      await chrome.alarms.clear(KEEPALIVE_ALARM_NAME)
    } catch (e) {
      console.warn('[mail-organizer] keep-alive clear failed', e)
    }
  }
}

/**
 * Wrap a long-running operation so the SW stays awake throughout. Safe to
 * nest — internal refcount tracks active holders.
 */
export async function withKeepAlive<T>(fn: () => Promise<T>): Promise<T> {
  await holdKeepAlive()
  try {
    return await fn()
  } finally {
    await releaseKeepAlive()
  }
}

/**
 * Clear a keepalive alarm orphaned by a previous SW generation.
 *
 * `activeHolders` lives in memory and resets to 0 every time the SW is
 * idle-killed, but `chrome.alarms` persists across generations. If the SW
 * was killed mid-execute (before `releaseKeepAlive` ran), the alarm survives
 * and keeps firing every ~30s forever — pointlessly waking the SW and
 * draining battery — because no in-memory holder remains to release it, and
 * recovery marks stale work *failed* rather than resuming it under a fresh
 * hold. Call this once at startup, before any work can re-acquire a hold.
 * Guarded on `activeHolders === 0` so it never clears a live alarm.
 */
export async function clearOrphanKeepAlive(): Promise<void> {
  if (activeHolders > 0) return
  try {
    await chrome.alarms.clear(KEEPALIVE_ALARM_NAME)
  } catch (e) {
    console.warn('[mail-organizer] keep-alive orphan clear failed', e)
  }
}

/**
 * Register a no-op alarm listener so the alarm event actually wakes /
 * pings the SW. Called once at module load. Also clears any alarm orphaned
 * by a previous SW generation (see clearOrphanKeepAlive).
 */
export function installKeepAliveListener(): void {
  chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === KEEPALIVE_ALARM_NAME) {
      // No-op — just receiving the event keeps SW alive.
      // (Logging here would flood the console.)
    }
  })
  // Best-effort: drop an orphaned alarm from a prior generation. Safe because
  // module-load recovery does not resume work under a keepalive hold, so
  // activeHolders is still 0 here; if real work starts later it re-creates
  // the alarm through holdKeepAlive.
  void clearOrphanKeepAlive()
}
