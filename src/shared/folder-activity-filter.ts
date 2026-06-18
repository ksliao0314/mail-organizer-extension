// Filter predicate for "近日活動" panel display, extracted from the SW
// handler so it can be unit-tested. Inlined predicate had a regression
// on 2026-05-22: empty allowlists silently hid all activity for fresh
// installs (where the generification removed the hardcoded prefix
// defaults). Now: empty = no filter = show everything.

import type { FolderActivity } from './types'
import { encodeFolderName } from './outlook-api'

/**
 * Filter folder activity rows for the IdleScreen panel.
 *
 * Semantics:
 *   - Empty allowlists (both `prefixes` and `leafNames` empty) → return
 *     `rows` unchanged. Treats the absence of configuration as
 *     "no preference", which is the natural default for new users
 *     who haven't visited Options yet.
 *   - Either non-empty → return rows where path starts with any
 *     prefix OR leaf segment matches any leafName.
 *
 * Leaf-name matching: `row.folderPath` is built with `joinFolderPath`, which
 * runs each segment through `encodeFolderName` (a literal '/' inside a folder
 * name becomes '／'). The Options UI stores leaf names as the raw
 * `node.displayName`, so a folder whose name contains '/' (e.g. "2024/Q1")
 * would never match a path segment unless we encode the allowlist the same
 * way. We encode both sides here so the comparison is apples-to-apples.
 *
 * Pure function; trivially test-able.
 */
export function filterFolderActivity(
  rows: FolderActivity[],
  prefixes: readonly string[],
  leafNames: ReadonlySet<string>,
): FolderActivity[] {
  if (prefixes.length === 0 && leafNames.size === 0) {
    return rows.slice() // copy to avoid caller-side mutation surprises
  }
  const encodedLeafNames =
    leafNames.size > 0 ? new Set([...leafNames].map(encodeFolderName)) : leafNames
  return rows.filter((row) => {
    for (const p of prefixes) {
      if (!p) continue
      if (p.endsWith('/')) {
        // Legacy form (typed manually before the picker UI). Trailing
        // slash explicitly means "descendants only, not the folder
        // itself". Preserved for back-compat with users who already
        // entered `案件資料夾/` style strings.
        if (row.folderPath.startsWith(p)) return true
      } else {
        // New form (picker output): folder itself + descendants. The
        // explicit `/` separator on the descendant check stops
        // `案件` from matching `案件2`.
        if (row.folderPath === p || row.folderPath.startsWith(`${p}/`)) {
          return true
        }
      }
    }
    if (encodedLeafNames.size > 0) {
      const segments = row.folderPath.split('/').filter(Boolean)
      const leaf = segments[segments.length - 1]
      if (leaf && encodedLeafNames.has(leaf)) return true
    }
    return false
  })
}
