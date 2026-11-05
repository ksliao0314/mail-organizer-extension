import { beforeEach, describe, expect, it } from 'vitest'
import {
  getConversationMemory,
  getSubjectMemory,
  recordConversationFilings,
  recordSubjectFilings,
} from '@/shared/storage'

beforeEach(async () => {
  await chrome.storage.local.clear()
})

const filing = (folderId = 'fidA') => [{ convId: 'c1', folderId, folderPath: 'p/' + folderId }]

describe('harvested flag on thread-memory (C3)', () => {
  it('marks an entry harvested when the only filing is harvested', async () => {
    await recordConversationFilings(filing(), { harvested: true })
    const mem = await getConversationMemory()
    expect(mem.c1!.harvested).toBe(true)
  })

  it('a real (plan-reviewed) filing does NOT set the flag', async () => {
    await recordConversationFilings(filing())
    const mem = await getConversationMemory()
    expect(mem.c1!.harvested).toBeUndefined()
  })

  it('a real filing CONFIRMS a harvested entry — flag is cleared', async () => {
    await recordConversationFilings(filing(), { harvested: true })
    await recordConversationFilings(filing()) // real filing, same conv+folder
    const mem = await getConversationMemory()
    expect(mem.c1!.harvested).toBeUndefined()
    expect(mem.c1!.timesFiled).toBe(2)
  })

  it('a harvest filing on an already-confirmed entry does NOT re-flag it', async () => {
    await recordConversationFilings(filing()) // real first
    await recordConversationFilings(filing(), { harvested: true }) // harvest after
    const mem = await getConversationMemory()
    expect(mem.c1!.harvested).toBeUndefined()
  })

  it('a harvest filing NEVER mutates a confirmed entry (provenance firewall, 2026-11-05)', async () => {
    // Repro: the lawyer's own (extension-executed) filing confirmed c1 → Y.
    // 22+ days later the moved-by-us ledger had forgotten a move of ours into
    // X, so the harvester re-saw it as a "manual filing" — and flipped the
    // CONFIRMED entry to {folderId: X, conflictCount: 1, previousFolderId: Y}.
    await recordConversationFilings(filing('fidY')) // real, plan-reviewed
    const before = (await getConversationMemory()).c1
    await recordConversationFilings(filing('fidX'), { harvested: true })
    const after = (await getConversationMemory()).c1
    expect(after).toEqual(before)
    expect(after!.folderId).toBe('fidY')
    expect(after!.conflictCount ?? 0).toBe(0)
    expect(after!.previousFolderId).toBeUndefined()
  })

  it('subject memory gets the same firewall', async () => {
    const subj = [{ normalizedSubject: '112訴204 開庭通知', folderId: 'fidY', folderPath: 'p/fidY' }]
    await recordSubjectFilings(subj)
    await recordSubjectFilings([{ ...subj[0]!, folderId: 'fidX', folderPath: 'p/fidX' }], {
      harvested: true,
    })
    const e = (await getSubjectMemory())['112訴204 開庭通知']!
    expect(e.folderId).toBe('fidY')
    expect(e.timesFiled).toBe(1)
    expect(e.conflictCount ?? 0).toBe(0)
  })

  it('a second folder-content filing (e.g. re-scan) does NOT clear an existing harvested flag', async () => {
    // The P2 fix: harvest marks C harvested; a later initial-scan re-seed —
    // now also {harvested:true} — must keep it soft, not silently escalate it
    // to an auto-route. Only a real execute filing clears it.
    await recordConversationFilings(filing(), { harvested: true })
    await recordConversationFilings(filing(), { harvested: true })
    const mem = await getConversationMemory()
    expect(mem.c1!.harvested).toBe(true)
  })
})
