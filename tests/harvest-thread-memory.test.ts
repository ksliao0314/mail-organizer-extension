import { beforeEach, describe, expect, it } from 'vitest'
import { getConversationMemory, recordConversationFilings } from '@/shared/storage'

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
