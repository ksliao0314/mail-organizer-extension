// Refused-email ledger (2026-11-06).
//
// emailId → when / why Claude refused to classify it, and with which model.
// Without it a refused email was re-sent with its whole chunk every batch and
// every 繼續 — and refused again: its placeholder row is `unclassified`, so
// execute never records it in skipHistory and it always comes back. Ledgered
// emails (same model, within the TTL) now take no batch slot; preflight still
// lets rules and thread memory route them, and leaves only the rest out of the
// AI bucket — counted for the user, like flagged mail.
//
// chrome.storage.local, own key and own write lock; never synced, backed up or
// exported (another machine finds its refusals itself). Entries expire after
// REFUSED_EMAILS_TTL_MS so a classifier change gets a fresh chance.
//
// Also here: the shared-prompt refusal suspicion (`refusalSuspicion`, same
// storage rules) — see RefusalSuspicion.

import { REFUSED_EMAILS_CAP, REFUSED_EMAILS_TTL_MS } from './constants'

const KEY_REFUSED_EMAILS = 'refusedEmails'
const KEY_REFUSAL_SUSPICION = 'refusalSuspicion'

export type RefusedEmailEntry = {
  /** unix ms of the refusal. */
  at: number
  /** stop_details.category; null when the API gave none. */
  category: string | null
  /** Model that refused (the selected one; fallbacks had already declined too). */
  model: string
}

export type RefusedEmails = Record<string, RefusedEmailEntry>

function isEntry(v: unknown): v is RefusedEmailEntry {
  if (!v || typeof v !== 'object') return false
  const e = v as Record<string, unknown>
  return (
    typeof e.at === 'number' &&
    Number.isFinite(e.at) &&
    (typeof e.category === 'string' || e.category === null) &&
    typeof e.model === 'string'
  )
}

function prune(rec: Record<string, unknown>, now: number): RefusedEmails {
  const cutoff = now - REFUSED_EMAILS_TTL_MS
  const fresh: Array<[string, RefusedEmailEntry]> = []
  for (const [id, e] of Object.entries(rec)) {
    if (id && isEntry(e) && e.at >= cutoff) fresh.push([id, e])
  }
  if (fresh.length > REFUSED_EMAILS_CAP) {
    fresh.sort((a, b) => b[1].at - a[1].at) // newest first
    fresh.length = REFUSED_EMAILS_CAP
  }
  return Object.fromEntries(fresh)
}

async function readRaw(): Promise<Record<string, unknown>> {
  const stored = await chrome.storage.local.get(KEY_REFUSED_EMAILS)
  const raw = stored[KEY_REFUSED_EMAILS]
  return raw && typeof raw === 'object' && !Array.isArray(raw) ? (raw as Record<string, unknown>) : {}
}

/** Current (TTL-pruned) ledger. */
export async function getRefusedEmails(now: number = Date.now()): Promise<RefusedEmails> {
  return prune(await readRaw(), now)
}

let writeChain: Promise<void> = Promise.resolve()

/** Record refused email ids (newest wins on duplicates); prunes TTL + cap. */
export async function addRefusedEmails(
  entries: Array<{ emailId: string; category: string | null; model: string }>,
  now: number = Date.now(),
): Promise<void> {
  const usable = entries.filter((e) => e.emailId)
  if (usable.length === 0) return
  const run = writeChain.then(async () => {
    const rec = await readRaw()
    for (const e of usable) rec[e.emailId] = { at: now, category: e.category, model: e.model }
    await chrome.storage.local.set({ [KEY_REFUSED_EMAILS]: prune(rec, now) })
  })
  // Keep the chain alive past a failed write; the caller still sees the error.
  writeChain = run.catch(() => {})
  return run
}

/** Ids ledgered for `model` (TTL-pruned) — what the inbox scan leaves out. */
export async function getRefusedEmailIds(model: string, now: number = Date.now()): Promise<Set<string>> {
  const ids = new Set<string>()
  for (const [id, entry] of Object.entries(await getRefusedEmails(now))) {
    if (entry.model === model) ids.add(id)
  }
  return ids
}

// ---- Shared-prompt refusal suspicion -------------------------------------------

/**
 * "Requests with this model and this prompt prefix are refused whatever email
 * they carry" (2026-11-06). Written when a one-email probe sent with the run's
 * shared prompt (system prompt, folder list, rule examples, case table) was
 * refused too; while it holds (REFUSAL_SUSPICION_TTL_MS), later runs neither
 * probe nor bisect a refusal of the same prefix — they cost what v1.0.8 did,
 * one call per chunk. Cleared as soon as a call with that prefix is answered.
 */
export type RefusalSuspicion = {
  model: string
  /** sharedPromptKey() of the refused requests (classifier.ts). */
  promptKey: string
  /** stop_details.category of the probe's refusal (display / diagnostics). */
  category: string | null
  /** unix ms. */
  at: number
}

function isSuspicion(v: unknown): v is RefusalSuspicion {
  if (!v || typeof v !== 'object') return false
  const s = v as Record<string, unknown>
  return (
    typeof s.model === 'string' &&
    typeof s.promptKey === 'string' &&
    s.promptKey.length > 0 &&
    (typeof s.category === 'string' || s.category === null) &&
    typeof s.at === 'number' &&
    Number.isFinite(s.at)
  )
}

/** The remembered suspicion, or null (none / malformed). TTL is the caller's check. */
export async function getRefusalSuspicion(): Promise<RefusalSuspicion | null> {
  const raw = (await chrome.storage.local.get(KEY_REFUSAL_SUSPICION))[KEY_REFUSAL_SUSPICION]
  return isSuspicion(raw) ? raw : null
}

export async function setRefusalSuspicion(s: RefusalSuspicion): Promise<void> {
  await chrome.storage.local.set({ [KEY_REFUSAL_SUSPICION]: s })
}

export async function clearRefusalSuspicion(): Promise<void> {
  await chrome.storage.local.remove(KEY_REFUSAL_SUSPICION)
}
