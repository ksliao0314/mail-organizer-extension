// Own-mailbox-domain wiring in the real classify preflight (2026-11-05).
//
// rules.ts keeps own-domain addresses from triggering AUTO domain rules when
// buildRuleIndex gets `ownDomains`; this checks the SW actually resolves the
// mailbox identity (GET /me, cached) and passes it. Before the fix, an
// auto-learned `domain: firm.example` rule claimed every email addressed to
// me@firm.example — i.e. nearly all inbound mail.

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { newRule } from '@/shared/rules'
import { getMailboxIdentity, setMailboxIdentity, setRules } from '@/shared/storage'
import { FOLDER_ID, FOLDER_PATH, inboxMail, installFetchRouter, seedClassifyEnv } from './helpers/classify-env'
import { loadServiceWorker, type SwHarness } from './helpers/sw-harness'

const ownDomainRule = () => ({
  ...newRule({
    type: 'domain',
    signal: 'firm.example',
    targetFolderId: FOLDER_ID,
    targetFolderPath: FOLDER_PATH,
    confidence: 0.7,
    source: 'ai_confirmed',
  }),
  createdAt: '2026-10-01T00:00:00.000Z',
})

const toMe = (i: number) =>
  inboxMail(i, {
    From: { EmailAddress: { Address: `clerk${i}@court.example`, Name: 'Court' } },
    ToRecipients: [{ EmailAddress: { Address: 'me@firm.example', Name: 'Me' } }],
  })

describe('classify preflight passes the mailbox domain to the rule index', () => {
  let sw: SwHarness
  beforeAll(async () => {
    sw = await loadServiceWorker()
  })
  beforeEach(async () => {
    await seedClassifyEnv()
  })
  afterEach(() => vi.unstubAllGlobals())

  function routeMe(address: string | null) {
    const router = installFetchRouter()
    const inner = globalThis.fetch as unknown as (url: string, init?: RequestInit) => Promise<Response>
    const meCalls: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init?: RequestInit) => {
        if (/\/me(\?|$)/.test(url)) {
          meCalls.push(url)
          return address === null
            ? new Response('{}', { status: 500 })
            : new Response(JSON.stringify({ EmailAddress: address }), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              })
        }
        return inner(url, init)
      }),
    )
    return { router, meCalls }
  }

  it('looks up GET /me once, stores the identity, and an own-domain AUTO rule no longer routes mail', async () => {
    await setRules([ownDomainRule()])
    const { router, meCalls } = routeMe('Me@Firm.example')
    router.inboxPages = [{ messages: [toMe(0), toMe(1)] }]

    const r = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    const d = r.data as { unmatchedPreview: Array<{ id: string }> }
    // Both go to the AI bucket instead of being swallowed by the own-domain rule.
    expect(d.unmatchedPreview.map((m) => m.id).sort()).toEqual(['m0', 'm1'])
    expect(meCalls).toHaveLength(1)
    expect(await getMailboxIdentity()).toMatchObject({ address: 'me@firm.example', domain: 'firm.example' })
  })

  it('a fresh cached identity is reused without another /me call', async () => {
    await setRules([ownDomainRule()])
    await setMailboxIdentity({ address: 'me@firm.example', domain: 'firm.example', at: new Date().toISOString() })
    const { router, meCalls } = routeMe('me@firm.example')
    router.inboxPages = [{ messages: [toMe(0)] }]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    expect(meCalls).toHaveLength(0)
    expect((r.data as { unmatchedPreview: Array<{ id: string }> }).unmatchedPreview.map((m) => m.id)).toEqual(['m0'])
  })

  it('if /me fails and nothing is cached, classification still works (pre-fix behaviour)', async () => {
    await setRules([ownDomainRule()])
    const { router } = routeMe(null)
    router.inboxPages = [{ messages: [toMe(0)] }]
    const r = await sw.send({ type: 'classifyPreflight', batchSize: 5 })
    expect(r.ok, JSON.stringify(r)).toBe(true)
    expect(await getMailboxIdentity()).toBeNull()
  })
})
