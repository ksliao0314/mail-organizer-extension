// The initial scan resolves the mailbox identity itself (2026-11-06, B12).
//
// initial-scan read only the CACHED identity, which only a classify
// preflight fills — so on a fresh install, or for an upgrader who scans
// before classifying, Pass A counted colleagues' mail and could mint
// `domain: <own domain>` (a rule the owner's address on To/Cc lets match
// nearly every inbound email). The startInitialScan handler now awaits
// resolveOwnDomains (bounded GET /me, fail-soft) and passes it in.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { getScanState } from '@/background/initial-scan'
import { listRules } from '@/shared/rules'
import { getMailboxIdentity } from '@/shared/storage'
import { waitFor } from './helpers/classify-env'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const ROOT = 'ROOT-FOLDER-ID-00000000000'
const FOLDER_A = 'FOLDER-A-00000000000000000'

let sw: SwHarness
beforeAll(async () => {
  sw = await loadServiceWorker()
})
beforeEach(async () => {
  await chrome.storage.local.set({
    settings: { claudeApiKey: 'sk-ant-test', primaryRootPath: '案件', internalDomains: [] },
    folderCache: {
      updatedAt: new Date().toISOString(),
      tree: [
        {
          id: ROOT,
          displayName: '案件',
          path: '案件',
          children: [{ id: FOLDER_A, displayName: 'A', path: '案件/A', children: [] }],
        },
      ],
    },
  })
  await chrome.storage.session.set({
    cachedToken: { secret: 'owa-token', expiresOn: Date.now() / 1000 + 3600 },
  })
  sw.tabsQuery.mockResolvedValue([
    { id: 7, active: true, url: 'https://outlook.office.com/mail/', status: 'complete' } as chrome.tabs.Tab,
  ])
})
afterEach(() => vi.unstubAllGlobals())

/** Outlook REST stub: GET /me → `me` (null = a definitive 403, no retry backoff);
 *  folder A holds colleague + client mail. */
function stubOutlook(me: string | null) {
  const meCalls: string[] = []
  const from = (address: string, i: number) => ({
    Id: `m${i}`,
    Subject: `主旨 ${i}`,
    From: { EmailAddress: { Address: address } },
    ToRecipients: [],
  })
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string) => {
      const json = (v: unknown, status = 200) =>
        new Response(JSON.stringify(v), { status, headers: { 'content-type': 'application/json' } })
      if (/\/me(\?|$)/.test(url)) {
        meCalls.push(url)
        return me === null ? json({ error: { code: 'ErrorAccessDenied' } }, 403) : json({ EmailAddress: me })
      }
      if (url.includes(`/MailFolders/${FOLDER_A}/messages`)) {
        return json({
          value: [
            from('colleague@firm.example', 1),
            from('colleague@firm.example', 2),
            from('pm@client.example', 3),
            from('pm@client.example', 4),
          ],
        })
      }
      return json({ value: [] })
    }),
  )
  return meCalls
}

async function scan(): Promise<void> {
  const r = await sw.send({ type: 'startInitialScan' })
  expect(r.ok, JSON.stringify(r)).toBe(true)
  await waitFor(async () => (await getScanState())?.inProgress === false, 5000)
}

describe('startInitialScan resolves the mailbox identity (B12)', () => {
  it('fresh install, empty identity cache: GET /me runs and the own domain is not minted', async () => {
    const meCalls = stubOutlook('Me@Firm.example')
    expect(await getMailboxIdentity()).toBeNull()
    await scan()
    expect(meCalls).toHaveLength(1)
    expect(await getMailboxIdentity()).toMatchObject({ domain: 'firm.example' })
    const domains = (await listRules()).filter((r) => r.type === 'domain').map((r) => r.signal)
    expect(domains).toEqual(['client.example']) // was ['firm.example', 'client.example']
  })

  it('a failing /me does not block the scan (fail-soft)', async () => {
    stubOutlook(null)
    await scan()
    const domains = (await listRules()).filter((r) => r.type === 'domain').map((r) => r.signal).sort()
    expect(domains).toEqual(['client.example', 'firm.example']) // pre-fix behaviour, but the scan ran
  })
})
