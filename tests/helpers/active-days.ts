// Shared test helper for the stale-sweep ACTIVE-day clock (2026-11-05).
//
// autoDisableStaleRules counts rule age in days on which a batch was actually
// executed, not calendar days. Tests that model "N days have passed" while the
// user kept classifying every day pass `activeDays: dailyActiveDays(now, n)`.

const DAY_MS = 86_400_000

/** UTC days ('YYYY-MM-DD') for "a batch ran every day" over the `count` days ending at `nowMs`. */
export function dailyActiveDays(nowMs: number, count: number): string[] {
  const out: string[] = []
  for (let i = 0; i < count; i++) {
    out.push(new Date(nowMs - i * DAY_MS).toISOString().slice(0, 10))
  }
  return out
}
