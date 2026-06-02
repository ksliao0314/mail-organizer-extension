// Initial scan.
//
// One-shot setup: walk a root path (default 案件), fetch the latest
// emails in each subfolder, count external domains, and create `auto_scan`
// domain rules for any domain that hits the folder ≥ MIN_COUNT times.
//
// Why count ≥ 2? Single-occurrence domains might be a stray CC; two or more
// shows the folder is a "home" for that domain. Confidence scales with count.

import { OutlookApi, OutlookError, flattenFolderTree } from '@/shared/outlook-api'
import {
  addRulesFilteringTombstones,
  encodeCompound,
  extractCaseCodes,
  extractCourtCaseNumbers,
  extractDomain,
  isGenericProviderDomain,
  listRules,
  newRule,
} from '@/shared/rules'
import { recordConversationFilings, recordSubjectFilings } from '@/shared/storage'
import { MIN_NORMALIZED_SUBJECT_LEN, normalizeSubject } from '@/shared/normalize'
import type { MailFolderNode, Rule } from '@/shared/types'
import { holdKeepAlive, releaseKeepAlive } from './keep-alive'

const SCAN_KEY = 'scanState'
const TOP_PER_FOLDER = 20
const MIN_COUNT = 2
// Domain rules carry an intentionally lower ceiling than the unique-signal
// rule types — even folder-unique domains can hide "same client across
// multiple cases" ramping up over time. Subject-specific rules (case_code /
// compound / subject_keyword) and AI per-email should still win.
const DOMAIN_MAX_CONFIDENCE = 0.65
const BASE_CONFIDENCE = 0.55
const FLUSH_EVERY_N_FOLDERS = 10

// Case codes (e.g. 25A0067A) are unique enough that two occurrences in one
// folder are already strong evidence — no need for the higher domain
// threshold. Confidence is high because case codes are essentially the
// strongest signal a Taiwanese law-firm email can carry.
const MIN_COUNT_CASE_CODE = 2
const CASE_CODE_CONFIDENCE = 0.95

// Taiwan court case numbers (e.g. 112訴204) are even stronger when paired
// with the sending domain: the COMPOUND of (domain + caseNumber) routes the
// specific case-client pair while letting other emails from the same domain
// flow elsewhere. ≥ 2 co-occurrences in one folder is sufficient signal.
const MIN_COUNT_COURT_CASE_COMPOUND = 2
const COURT_CASE_COMPOUND_CONFIDENCE = 0.9

// Subject-token cross-folder pass (3–8 char Chinese chunks) was removed
// in the 2026-05-27 redesign. The token heuristic conflicted with the
// "整段主旨 = signal" principle — token-extracted subject_keyword rules
// could fire on unrelated mail that happened to share the same noun
// fragment. case_code + court_case structural IDs are still extracted
// via dedicated passes above.

// Generic-provider sender extraction (gmail.com / yahoo.com / 等).
// Plain-domain rules can't represent these — gmail.com hosts countless
// unrelated senders — but a SPECIFIC address (andy@gmail.com) is a real
// routing signal. We extract these per-folder with the same cross-folder
// uniqueness gate used for plain domains: ≥ 2 in this folder + ≤ 1 in any
// other folder. Confidence matches the execute-time sender-fallback
// (0.7) so behaviour is consistent across the two paths.
const MIN_COUNT_SENDER = 2
const SENDER_CONFIDENCE = 0.7

function isAbortError(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError'
}

// In-memory abort controller for the running scan. Used to interrupt
// in-flight listFolderMessages immediately on user cancel.
let currentController: AbortController | null = null

export type ScanItemStatus = 'queued' | 'processing' | 'done' | 'no_domains' | 'empty' | 'error'

export type ScanItemResult = {
  folderPath: string
  status: ScanItemStatus
  emailsScanned?: number
  domainsFound?: Array<{ domain: string; count: number; added: boolean }>
  /** Case codes found in subject lines and whether each became a rule. */
  caseCodesFound?: Array<{ code: string; count: number; added: boolean }>
  /** (domain, court case number) compounds detected in this folder. */
  compoundsFound?: Array<{ domain: string; courtCase: string; count: number; added: boolean }>
  /** Generic-provider senders (e.g. andy@gmail.com) — `added` flipped by the cross-folder pass. */
  sendersFound?: Array<{ address: string; count: number; added: boolean }>
  rulesAdded?: number
  message?: string
}

export type ScanState = {
  inProgress: boolean
  cancelRequested: boolean
  startedAt: number
  finishedAt?: number
  total: number
  current: number
  rootPath: string
  results: ScanItemResult[]
  summary: {
    foldersScanned: number
    rulesAdded: number
    foldersWithNoExternalDomains: number
    errors: number
  }
}

export async function getScanState(): Promise<ScanState | null> {
  const r = await chrome.storage.session.get(SCAN_KEY)
  return (r[SCAN_KEY] as ScanState | undefined) ?? null
}

async function saveScanState(state: ScanState): Promise<void> {
  await chrome.storage.session.set({ [SCAN_KEY]: state })
}

export async function clearScanState(): Promise<void> {
  await chrome.storage.session.remove(SCAN_KEY)
}

/**
 * Same pattern as `recoverStaleExecuteState`: an `inProgress` scan with no
 * live worker means the SW was killed mid-run. Mark it failed so the next
 * scan can start.
 */
export async function recoverStaleScanState(): Promise<{ recovered: boolean }> {
  const state = await getScanState()
  if (!state?.inProgress) return { recovered: false }
  await saveScanState({
    ...state,
    inProgress: false,
    cancelRequested: false,
    finishedAt: state.finishedAt ?? Date.now(),
    results: state.results.map((r) =>
      r.status === 'processing' || r.status === 'queued'
        ? { ...r, status: 'error' as const, message: 'Service worker 中斷、掃描未完成' }
        : r,
    ),
  })
  console.warn('[mail-organizer] recovered stuck scanState')
  return { recovered: true }
}

export async function requestScanCancel(): Promise<void> {
  const state = await getScanState()
  if (state?.inProgress) {
    await saveScanState({ ...state, cancelRequested: true })
    currentController?.abort()
  }
}

export async function isScanRunning(): Promise<boolean> {
  return !!(await getScanState())?.inProgress
}

export async function startInitialScan(opts: {
  rootPath: string
  tree: MailFolderNode[]
  api: OutlookApi
  excludePrefixes: string[]
  /** Firm internal domains — sender addresses on these are skipped when
   *  extracting per-folder sender candidates (colleagues cross-work
   *  many cases). Empty array = solo mode, all senders eligible. */
  internalDomains: string[]
}): Promise<void> {
  const { rootPath, tree, api, excludePrefixes, internalDomains } = opts
  // Lowercase + Set for O(1) lookup in the per-email loop.
  const internalDomainSet = new Set(internalDomains.map((d) => d.toLowerCase()))
  const isInternalDomain = (d: string | null): boolean => {
    if (!d) return false
    return internalDomainSet.has(d.toLowerCase())
  }
  if (await isScanRunning()) throw new Error('已有掃描中的批次')

  const flat = flattenFolderTree(tree)
  const targets = flat.filter((n) => {
    const inRoot = n.path === rootPath || n.path.startsWith(rootPath + '/')
    if (!inRoot) return false
    const excluded = excludePrefixes.some((p) => n.path === p || n.path.startsWith(p + '/'))
    return !excluded
  })

  if (targets.length === 0) {
    throw new Error(`找不到根資料夾「${rootPath}」或其子資料夾`)
  }

  const state: ScanState = {
    inProgress: true,
    cancelRequested: false,
    startedAt: Date.now(),
    total: targets.length,
    current: 0,
    rootPath,
    results: targets.map((n) => ({ folderPath: n.path, status: 'queued' as const })),
    summary: { foldersScanned: 0, rulesAdded: 0, foldersWithNoExternalDomains: 0, errors: 0 },
  }
  await saveScanState(state)

  // Snapshot existing rules for dedup. Build a fast key set for O(1) lookups.
  // Tracks domain / case_code / compound / subject_keyword so a follow-up
  // scan doesn't produce duplicates for any of these auto-derived kinds.
  const existing = await listRules()
  const knownKeys = new Set<string>()
  for (const r of existing) {
    if (r.type === 'domain') {
      const dom = r.signal.toLowerCase().replace(/^@/, '')
      knownKeys.add(`domain::${dom}::${r.targetFolderId}`)
    } else if (r.type === 'case_code') {
      knownKeys.add(`case_code::${r.signal.toUpperCase().trim()}::${r.targetFolderId}`)
    } else if (r.type === 'compound') {
      // Use the raw (canonical) signal string — encodeCompound produces a
      // stable JSON serialization so it's safe to compare directly.
      knownKeys.add(`compound::${r.signal}::${r.targetFolderId}`)
    } else if (r.type === 'subject_keyword') {
      knownKeys.add(`subject_keyword::${r.signal.toLowerCase().trim()}::${r.targetFolderId}`)
    } else if (r.type === 'sender') {
      knownKeys.add(`sender::${r.signal.toLowerCase().trim()}::${r.targetFolderId}`)
    }
  }

  // Same pattern for plain-domain rules: collect per-folder, then decide in
  // a cross-folder pass. Two gates apply (see post-loop section):
  //   1. Generic-provider domains (gmail.com / yahoo.com / 等) — never
  //      auto-build a plain-domain rule. Same provider serves countless
  //      unrelated senders.
  //   2. Cross-folder spread — domains showing up in ≥ 2 folders mean "same
  //      client across multiple cases" (dominant pattern at this firm). A
  //      single plain-domain rule can't represent that; we leave routing
  //      to AI per-email.
  //
  // Compound rules (domain + court case) are NOT subject to these gates —
  // they're built per-folder in the main loop because the court-case half
  // already provides per-case specificity.
  type FolderDomainData = {
    folderId: string
    folderPath: string
    domainCounts: Map<string, number>
  }
  const folderDomainData: FolderDomainData[] = []

  // Parallel collection for generic-provider senders (specific addresses
  // like andy@gmail.com). Plain-domain rules can't represent gmail.com
  // sensibly, but individual addresses absolutely can — and the
  // cross-folder uniqueness gate still applies (an address showing in
  // multiple folders means the person works across cases, can't be
  // routed by one rule).
  //
  // Skip senders on `internalDomains`: colleagues send about many
  // cases, so even a folder-unique colleague address is misleading.
  type FolderSenderData = {
    folderId: string
    folderPath: string
    senderCounts: Map<string, number> // key: lowercased full email address
  }
  const folderSenderData: FolderSenderData[] = []

  const buffer: Rule[] = [] // staged rules awaiting flush
  let lastFlushAt = -1

  // Thread-memory seeds collected during the scan. Flushed once at the
  // end of the scan (or on cancel/throw) — no point updating after each
  // folder because the scan typically completes in seconds. Same dedup
  // semantics as runtime filings: latest folder wins for ConvId, subject
  // accumulates conflictCount when filed to different folders.
  const convFilings: Array<{ convId: string; folderId: string; folderPath: string }> = []
  const subjectFilings: Array<{
    normalizedSubject: string
    folderId: string
    folderPath: string
  }> = []

  const flush = async () => {
    if (buffer.length === 0) return
    const toFlush = buffer.splice(0)
    try {
      // Respect prior user deletions — don't auto-resurrect what they killed.
      // Filter + write atomically under rules lock to avoid resurrecting
      // rules whose tombstones land between our check and our write.
      await addRulesFilteringTombstones(toFlush)
    } catch (e) {
      console.error('[mail-organizer] scan flush failed; putting back into buffer', e)
      buffer.unshift(...toFlush)
      throw e
    }
  }

  const controller = new AbortController()
  currentController = controller
  await holdKeepAlive()

  try {
    for (let i = 0; i < targets.length; i++) {
      const live = await getScanState()
      if (live?.cancelRequested) break

      const node = targets[i]!
      state.current = i
      state.results[i] = { folderPath: node.path, status: 'processing' }
      await saveScanState(state)

      try {
        const emails = await api.listFolderMessages(node.id, {
          top: TOP_PER_FOLDER,
          // Subject pulled for case-code / token extraction. ConversationId
          // pulled so we can seed thread memory in the same fetch — no
          // extra round-trips, the data's already coming back.
          select: 'Id,Subject,From,ReceivedDateTime,ConversationId',
          signal: controller.signal,
        })

        if (emails.length === 0) {
          state.results[i] = { folderPath: node.path, status: 'empty', emailsScanned: 0 }
          state.summary.foldersScanned++
          await saveScanState(state)
          continue
        }

        const domainCounts = new Map<string, number>()
        const caseCodeCounts = new Map<string, number>()
        // (domain, courtCase) co-occurrences for compound rules. Same domain
        // + court case appearing together ≥ MIN_COUNT_COURT_CASE_COMPOUND
        // times in this folder turns into a compound rule.
        const compoundCounts = new Map<string, { domain: string; courtCase: string; count: number }>()
        // Generic-provider sender counts (key: lowercased full address).
        // Only addresses on gmail.com / yahoo.com / 等 get tracked here —
        // non-generic addresses are already routable via the plain-domain
        // rule, so a parallel sender rule would just shadow it.
        const senderCounts = new Map<string, number>()
        for (const m of emails) {
          const addr = m.From?.EmailAddress?.Address
          const d = extractDomain(addr ?? null)
          if (d && !isInternalDomain(d)) {
            domainCounts.set(d, (domainCounts.get(d) ?? 0) + 1)
            if (addr && isGenericProviderDomain(d)) {
              const normAddr = addr.toLowerCase().trim()
              if (normAddr) {
                senderCounts.set(normAddr, (senderCounts.get(normAddr) ?? 0) + 1)
              }
            }
          }
          // Case codes in subject — strongest single-condition signal for
          // a Taiwanese law-firm inbox. extractCaseCodes returns canonical
          // uppercase form ([A-Z] segments) and de-dupes within one subject.
          for (const code of extractCaseCodes(m.Subject ?? '')) {
            caseCodeCounts.set(code, (caseCodeCounts.get(code) ?? 0) + 1)
          }
          // Taiwan court case numbers in subject, co-counted with domain so
          // we can build (domain + courtCase) compound rules.
          const courtCases = extractCourtCaseNumbers(m.Subject ?? '')
          if (d && !isInternalDomain(d) && courtCases.length > 0) {
            for (const cc of courtCases) {
              const key = `${d}::${cc}`
              const prev = compoundCounts.get(key)
              compoundCounts.set(key, {
                domain: d,
                courtCase: cc,
                count: (prev?.count ?? 0) + 1,
              })
            }
          }
          // Seed thread memory. Cheap — data's already in hand. Means the
          // first real classify batch after scan already has thread coverage
          // instead of waiting for AI to slowly fill it in.
          if (m.ConversationId) {
            convFilings.push({
              convId: m.ConversationId,
              folderId: node.id,
              folderPath: node.path,
            })
          }
          const norm = normalizeSubject(m.Subject ?? '')
          if (norm.length >= MIN_NORMALIZED_SUBJECT_LEN) {
            subjectFilings.push({
              normalizedSubject: norm,
              folderId: node.id,
              folderPath: node.path,
            })
          }
        }

        folderDomainData.push({
          folderId: node.id,
          folderPath: node.path,
          domainCounts,
        })
        folderSenderData.push({
          folderId: node.id,
          folderPath: node.path,
          senderCounts,
        })

        const sortedDomains = [...domainCounts.entries()].sort((a, b) => b[1] - a[1])
        const sortedCaseCodes = [...caseCodeCounts.entries()].sort((a, b) => b[1] - a[1])
        const sortedCompounds = [...compoundCounts.values()].sort((a, b) => b.count - a.count)
        const sortedSenders = [...senderCounts.entries()].sort((a, b) => b[1] - a[1])
        // Single-folder UI: show every detected domain with count, but mark
        // them all as not-yet-added. The cross-folder pass after the main
        // loop decides what becomes a real plain-domain rule and flips the
        // `added` flag on the matching entry.
        const reportedDomains: Array<{ domain: string; count: number; added: boolean }> =
          sortedDomains
            .filter(([, count]) => count >= MIN_COUNT)
            .map(([domain, count]) => ({ domain, count, added: false }))
        const reportedCaseCodes: Array<{ code: string; count: number; added: boolean }> = []
        const reportedCompounds: Array<{ domain: string; courtCase: string; count: number; added: boolean }> = []
        // Same pattern: show every generic-provider sender ≥ MIN_COUNT_SENDER
        // with `added=false`; the cross-folder pass below flips the flag.
        const reportedSenders: Array<{ address: string; count: number; added: boolean }> =
          sortedSenders
            .filter(([, count]) => count >= MIN_COUNT_SENDER)
            .map(([address, count]) => ({ address, count, added: false }))
        let added = 0

        for (const [code, count] of sortedCaseCodes) {
          if (count < MIN_COUNT_CASE_CODE) {
            reportedCaseCodes.push({ code, count, added: false })
            continue
          }
          const key = `case_code::${code}::${node.id}`
          if (knownKeys.has(key)) {
            reportedCaseCodes.push({ code, count, added: false })
            continue
          }
          buffer.push(
            newRule({
              type: 'case_code',
              signal: code,
              targetFolderId: node.id,
              targetFolderPath: node.path,
              confidence: CASE_CODE_CONFIDENCE,
              source: 'auto_scan',
            }),
          )
          knownKeys.add(key)
          reportedCaseCodes.push({ code, count, added: true })
          added++
        }

        for (const { domain, courtCase, count } of sortedCompounds) {
          if (count < MIN_COUNT_COURT_CASE_COMPOUND) {
            reportedCompounds.push({ domain, courtCase, count, added: false })
            continue
          }
          const signal = encodeCompound([
            { type: 'domain', value: domain },
            { type: 'subject_keyword', value: courtCase },
          ])
          const key = `compound::${signal}::${node.id}`
          if (knownKeys.has(key)) {
            reportedCompounds.push({ domain, courtCase, count, added: false })
            continue
          }
          buffer.push(
            newRule({
              type: 'compound',
              signal,
              targetFolderId: node.id,
              targetFolderPath: node.path,
              confidence: COURT_CASE_COMPOUND_CONFIDENCE,
              source: 'auto_scan',
            }),
          )
          knownKeys.add(key)
          reportedCompounds.push({ domain, courtCase, count, added: true })
          added++
        }

        // Status is provisional — the cross-folder pass at the end may still
        // promote some `reportedDomains` entries to `added=true` and bump
        // rulesAdded / status accordingly. We use 'done' optimistically so a
        // folder with case_code or compound hits doesn't display "no_domains".
        state.results[i] = {
          folderPath: node.path,
          status: 'done',
          emailsScanned: emails.length,
          domainsFound: reportedDomains,
          caseCodesFound: reportedCaseCodes.length > 0 ? reportedCaseCodes : undefined,
          compoundsFound: reportedCompounds.length > 0 ? reportedCompounds : undefined,
          sendersFound: reportedSenders.length > 0 ? reportedSenders : undefined,
          rulesAdded: added,
        }
        state.summary.foldersScanned++
        state.summary.rulesAdded += added
      } catch (e) {
        if (isAbortError(e)) break
        state.results[i] = {
          folderPath: node.path,
          status: 'error',
          message:
            e instanceof OutlookError
              ? `[${e.status}] ${e.message}`
              : e instanceof Error
              ? e.message
              : String(e),
        }
        state.summary.errors++
      }

      await saveScanState(state)

      // Periodic flush so a long scan doesn't lose all rules on failure
      if (i - lastFlushAt >= FLUSH_EVERY_N_FOLDERS && buffer.length > 0) {
        await flush().catch((e) => console.error('[mail-organizer] periodic flush failed', e))
        lastFlushAt = i
      }
    }

    // ---- Cross-folder uniqueness helper ---------------------------------
    //
    // Used by both Pass A (plain-domain) and Pass A2 (generic-provider
    // sender). The shared predicate is:
    //
    //   "Signal appears ≥ MIN_COUNT in exactly ONE folder, AND ≤ MIN_COUNT-1
    //    (= 1 with MIN_COUNT=2) in every OTHER folder."
    //
    // This is the conservative "same client across multiple cases" filter
    // — if a signal anchors two distinct folders, we can't routes by it
    // and AI per-email decides instead.
    //
    // Pass A previously expressed this as "build folderMap only for
    // entries ≥ MIN_COUNT, then require size === 1"; Pass A2 expressed
    // it as "scan for unique ≥ MIN_COUNT entry, then verify others ≤ 1".
    // Identical predicate, two inconsistent code paths — consolidated
    // here for future maintenance.
    //
    // `folderData` is the per-folder count table; each entry has its own
    // signal namespace (domain or sender address). Returns the winning
    // (folderId, count) or null if the predicate fails.
    type FolderCountSnapshot = {
      folderId: string
      folderPath: string
      counts: Map<string, number>
    }
    const findUniqueFolderCandidate = (
      signal: string,
      folderData: FolderCountSnapshot[],
      minCount: number,
    ): { folderId: string; folderPath: string; count: number } | null => {
      let candidate: { folderId: string; folderPath: string; count: number } | null = null
      const others: Array<{ folderId: string; count: number }> = []
      for (const fd of folderData) {
        const c = fd.counts.get(signal)
        if (!c) continue
        if (c >= minCount) {
          if (candidate !== null) return null // ambiguous: 2+ folders meet candidate threshold
          candidate = { folderId: fd.folderId, folderPath: fd.folderPath, count: c }
        } else {
          others.push({ folderId: fd.folderId, count: c })
        }
      }
      if (!candidate) return null
      // No "other" folder should exceed minCount-1 (= 1 with minCount=2).
      // Note: every entry here has count < minCount by construction, so
      // this loop is effectively a no-op when minCount === 2 (1 is the
      // only sub-threshold value). Kept explicit for clarity and to
      // self-document the predicate in case minCount changes.
      const maxInOthers = minCount - 1
      for (const o of others) {
        if (o.count > maxInOthers) return null
      }
      return candidate
    }

    // ---- Cross-folder pass A: plain-domain rules with both gates --------
    //
    // Gates applied IN ADDITION to the uniqueness helper above:
    //
    //   Gate 1 (generic provider): never auto-build for gmail.com /
    //   yahoo.com / 等 — one provider serves countless unrelated senders.
    //   Pass A2 below handles these as specific-address sender rules.
    //
    // Domains passing the gate → build the plain-domain rule and flip
    // the matching reportedDomains entry's `added` flag so the scan UI
    // shows the truth instead of "detected but silently dropped".
    try {
      // Aggregate the union of all domains seen across folders, then
      // probe via the shared helper.
      const allDomains = new Set<string>()
      for (const fd of folderDomainData) {
        for (const d of fd.domainCounts.keys()) allDomains.add(d)
      }
      const folderDomainSnap: FolderCountSnapshot[] = folderDomainData.map((fd) => ({
        folderId: fd.folderId,
        folderPath: fd.folderPath,
        counts: fd.domainCounts,
      }))
      for (const domain of allDomains) {
        if (isGenericProviderDomain(domain)) continue
        const cand = findUniqueFolderCandidate(domain, folderDomainSnap, MIN_COUNT)
        if (!cand) continue
        const key = `domain::${domain}::${cand.folderId}`
        if (knownKeys.has(key)) continue
        // Confidence scales with count, capped at DOMAIN_MAX_CONFIDENCE
        // (0.65) — even folder-unique domains can hide same-client-
        // multiple-cases ramping up over time; user / AI per-email
        // should still be able to win on a per-email basis.
        const confidence = Math.min(
          DOMAIN_MAX_CONFIDENCE,
          BASE_CONFIDENCE + cand.count * 0.05,
        )
        buffer.push(
          newRule({
            type: 'domain',
            signal: domain,
            targetFolderId: cand.folderId,
            targetFolderPath: cand.folderPath,
            confidence,
            source: 'auto_scan',
          }),
        )
        knownKeys.add(key)
        const folderResult = state.results.find((r) => r.folderPath === cand.folderPath)
        if (folderResult?.domainsFound) {
          const entry = folderResult.domainsFound.find((d) => d.domain === domain)
          if (entry) entry.added = true
        }
        state.summary.rulesAdded++
      }
    } catch (e) {
      console.warn('[mail-organizer] cross-folder domain pass failed (non-fatal)', e)
    }

    // ---- Cross-folder pass A2: generic-provider sender rules -----------
    //
    // Parallel to Pass A but for SPECIFIC addresses (e.g. andy@gmail.com).
    // Already filtered at collection time to generic providers only — a
    // non-generic sender will be routable via its plain-domain rule.
    // Uniqueness predicate is identical to Pass A's; the only difference
    // is the signal type (full email vs domain).
    try {
      const allSenders = new Set<string>()
      for (const fd of folderSenderData) {
        for (const s of fd.senderCounts.keys()) allSenders.add(s)
      }
      const folderSenderSnap: FolderCountSnapshot[] = folderSenderData.map((fd) => ({
        folderId: fd.folderId,
        folderPath: fd.folderPath,
        counts: fd.senderCounts,
      }))
      for (const address of allSenders) {
        const cand = findUniqueFolderCandidate(address, folderSenderSnap, MIN_COUNT_SENDER)
        if (!cand) continue
        const key = `sender::${address}::${cand.folderId}`
        if (knownKeys.has(key)) continue
        buffer.push(
          newRule({
            type: 'sender',
            signal: address,
            targetFolderId: cand.folderId,
            targetFolderPath: cand.folderPath,
            confidence: SENDER_CONFIDENCE,
            source: 'auto_scan',
          }),
        )
        knownKeys.add(key)
        const folderResult = state.results.find((r) => r.folderPath === cand.folderPath)
        if (folderResult?.sendersFound) {
          const entry = folderResult.sendersFound.find((s) => s.address === address)
          if (entry) entry.added = true
        }
        state.summary.rulesAdded++
      }
    } catch (e) {
      console.warn('[mail-organizer] cross-folder sender pass failed (non-fatal)', e)
    }

    // Reflect the cross-folder domain / sender passes back onto per-folder
    // rulesAdded so the summary metric "foldersWithNoExternalDomains" is
    // accurate. A folder that contributed only a plain-domain or sender
    // rule (and no case_code / compound) would otherwise still show as
    // "no rules added" since the per-folder loop ran before the
    // cross-folder gates decided to build them.
    for (const fr of state.results) {
      const domainAdded = fr.domainsFound?.filter((d) => d.added).length ?? 0
      const senderAdded = fr.sendersFound?.filter((s) => s.added).length ?? 0
      const total = domainAdded + senderAdded
      if (total > 0) fr.rulesAdded = (fr.rulesAdded ?? 0) + total
    }
    state.summary.foldersWithNoExternalDomains = state.results.filter(
      (r) => r.status === 'done' && (r.rulesAdded ?? 0) === 0,
    ).length

    // Cross-folder pass B (unique subject tokens) was removed in the
    // 2026-05-27 redesign — see the constant block at top of file.

    state.inProgress = false
    state.finishedAt = Date.now()
    await saveScanState(state)
  } catch (e) {
    state.inProgress = false
    state.finishedAt = Date.now()
    await saveScanState(state)
    throw e
  } finally {
    // Final flush — runs even on cancel / throw, so accumulated rules persist.
    if (buffer.length > 0) {
      try {
        await flush()
      } catch (e) {
        console.error('[mail-organizer] final scan flush failed; rules may be lost', e)
      }
    }
    // Seed thread memory in parallel with rule flush. Non-fatal: a
    // failure here just means the first real batch loses the head-start;
    // rules still went out fine.
    if (convFilings.length > 0 || subjectFilings.length > 0) {
      try {
        await Promise.all([
          convFilings.length > 0 ? recordConversationFilings(convFilings) : Promise.resolve(),
          subjectFilings.length > 0 ? recordSubjectFilings(subjectFilings) : Promise.resolve(),
        ])
      } catch (e) {
        console.warn('[mail-organizer] thread memory seed failed (non-fatal)', e)
      }
    }
    currentController = null
    await releaseKeepAlive()
  }
}
