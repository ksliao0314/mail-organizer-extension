// Centralised error log persisted in chrome.storage.local.
//
// Why: SW errors fan out into ~30 console.warn sites — auth failures,
// quota errors, parse errors, sync errors, retry exhaustions. console
// is fine while you have the SW devtools open, but the lawyer doesn't.
// When something quietly stops working, there's no UI surface to look
// at and no way to share the symptom for diagnosis.
//
// Design:
//   - One module-level helper `logError(source, message, context?)`.
//   - Capped at ERROR_LOG_CAP entries, oldest evicted FIFO.
//   - Read via `getErrorLog(limit?)`; clear via `clearErrorLog()`.
//   - The Options page hosts a card showing recent entries + filter.
//
// Source codes are domain-prefixed strings ('sync:push', 'classify:api',
// 'execute:move', etc.) so the UI can filter / colour-code.

import { ERROR_LOG_CAP } from './constants'

export type ErrorLogEntry = {
  at: string // ISO timestamp
  source: string // domain-prefixed: 'sync:push', 'classify:api', etc.
  message: string
  // Lightweight metadata. Never put PII or API keys here — this is
  // user-readable + diagnostic-export-shippable.
  context?: Record<string, unknown>
}

const KEY = 'errorLog'

// Bug #U: logError does read-modify-write on chrome.storage.local. Two
// concurrent calls (e.g. sync push and remote-pull both failing at the
// same time — exactly when logging matters most) would each read the
// same N-entry list, append their own entry, write — last writer wins,
// losing the earlier append. Serialize through this chain.
let writeChain: Promise<void> = Promise.resolve()

async function withErrorLogLock<T>(fn: () => Promise<T>): Promise<T> {
  const release = writeChain
  let resolve: () => void = () => {}
  writeChain = new Promise<void>((r) => {
    resolve = r
  })
  try {
    await release
    return await fn()
  } finally {
    resolve()
  }
}

/**
 * Append one error entry. Safe to call from any context (popup, SW,
 * options, content script) — it just writes to chrome.storage.local
 * which is shared. Failures here are swallowed so the logger itself
 * can't cause a cascade.
 *
 * Bug #U: serialized via withErrorLogLock — concurrent logError calls
 * can't clobber each other's append.
 */
export async function logError(
  source: string,
  message: string,
  context?: Record<string, unknown>,
): Promise<void> {
  try {
    await withErrorLogLock(async () => {
      const entry: ErrorLogEntry = {
        at: new Date().toISOString(),
        source,
        message,
        context,
      }
      const stored = await chrome.storage.local.get(KEY)
      const existing = (stored[KEY] as ErrorLogEntry[] | undefined) ?? []
      const next = [...existing, entry]
      // FIFO eviction once over cap. Append-newest, drop-front means a
      // burst of errors doesn't lose the oldest first-cause entry until
      // we genuinely fill ERROR_LOG_CAP slots.
      const trimmed =
        next.length > ERROR_LOG_CAP
          ? next.slice(next.length - ERROR_LOG_CAP)
          : next
      await chrome.storage.local.set({ [KEY]: trimmed })
    })
  } catch (e) {
    // Logger MUST NOT throw. Worst case we lose this entry.
    // eslint-disable-next-line no-console
    console.debug('[mail-organizer] error-log write failed', e)
  }
}

export async function getErrorLog(limit?: number): Promise<ErrorLogEntry[]> {
  const stored = await chrome.storage.local.get(KEY)
  const all = (stored[KEY] as ErrorLogEntry[] | undefined) ?? []
  if (limit && limit < all.length) {
    return all.slice(all.length - limit)
  }
  return all
}

export async function clearErrorLog(): Promise<void> {
  await chrome.storage.local.remove(KEY)
}
