// Mailbox identity, round 2 (2026-11-06 r2):
//   O6a — the learners and the initial scan read the RAW cached identity: after
//         an OWA account switch they kept excluding the previous mailbox's
//         domain while the matcher already refused that identity. One
//         account-checked accessor now serves every reader.
//   O6c — the identity lookup's 3 s budget didn't cover the forced token
//         refresh after a 401 (up to ~10 s per OWA tab ask): the request's
//         token wait now honours the caller's signal.

import { afterEach, describe, expect, it, vi } from 'vitest'
import { generateAiConfirmedRules } from '@/background/execute'
import { getAccountCheckedIdentity } from '@/background/mailbox-identity'
import { OutlookApi } from '@/shared/outlook-api'
import { listRules } from '@/shared/rules'
import { setMailboxIdentity } from '@/shared/storage'
import { TokenBucket } from '@/shared/throttle'
import type { ExecuteItemResult, PlanItem } from '@/shared/types'

afterEach(() => vi.unstubAllGlobals())

const tokenFor = (accountId?: string) =>
  chrome.storage.session.set({
    cachedToken: { secret: 'owa-token', expiresOn: Date.now() / 1000 + 3600, ...(accountId ? { accountId } : {}) },
  })
const IDENTITY_A = { address: 'me@firm-a.example', domain: 'firm-a.example', at: new Date().toISOString(), accountId: 'acct-A' }

describe('O6a — one account-checked identity for every reader', () => {
  it('never returns another account’s identity; an unknown account can’t be told apart', async () => {
    await setMailboxIdentity(IDENTITY_A)
    await tokenFor('acct-B')
    expect(await getAccountCheckedIdentity()).toBeNull()
    await tokenFor('acct-A')
    expect(await getAccountCheckedIdentity()).toMatchObject({ domain: 'firm-a.example' })
    await chrome.storage.session.remove('cachedToken')
    expect(await getAccountCheckedIdentity()).toMatchObject({ domain: 'firm-a.example' })
  })

  const item = (): PlanItem => ({
    emailId: 'e1',
    emailSubject: '季度合作會議紀錄與後續',
    emailFrom: 'pm@firm-a.example',
    action: 'move',
    targetFolderPath: '案件/甲',
    confidence: 0.9,
    reason: '',
    source: 'unresolved', // no birth gate — isolates the domain decision
  })
  const moved: ExecuteItemResult = { emailId: 'e1', subject: 'X', action: 'move', status: 'moved' }

  it('learners: after a switch to another mailbox, the OLD mailbox’s domain is an ordinary one', async () => {
    await setMailboxIdentity(IDENTITY_A)
    await tokenFor('acct-B')
    await generateAiConfirmedRules([item()], [moved])
    expect((await listRules()).map((r) => `${r.type}:${r.signal}`)).toEqual(['domain:firm-a.example'])
  })

  it('learners: for the same account the own domain is still never learned', async () => {
    await setMailboxIdentity(IDENTITY_A)
    await tokenFor('acct-A')
    await generateAiConfirmedRules([item()], [moved])
    expect((await listRules()).filter((r) => r.type === 'domain')).toEqual([])
  })
})

describe('O6c — a request deadline covers the token refresh after a 401', () => {
  it('a 401 whose forced token refresh hangs gives up when the caller’s signal does', async () => {
    const getToken = vi.fn(async (opts?: { force?: boolean }) =>
      opts?.force ? new Promise<string>(() => {}) : 'tok',
    )
    vi.stubGlobal('fetch', vi.fn(async () => new Response('expired', { status: 401 })))
    const api = new OutlookApi(getToken, undefined, { bucket: new TokenBucket(1000, 1000) })
    const started = Date.now()
    await expect(api.getMyEmail(AbortSignal.timeout(50), { unthrottled: true })).rejects.toMatchObject({
      name: 'TimeoutError',
    })
    expect(Date.now() - started).toBeLessThan(2000)
    expect(getToken).toHaveBeenLastCalledWith({ force: true })
  })

  it('a token failure without an abort is still OutlookTokenUnavailableError', async () => {
    const getToken = vi.fn(async () => {
      throw Object.assign(new Error('請先打開並登入 outlook.office.com 分頁'), { code: 'NOT_ON_OWA' })
    })
    vi.stubGlobal('fetch', vi.fn())
    const api = new OutlookApi(getToken, undefined, { bucket: new TokenBucket(1000, 1000) })
    await expect(api.getMyEmail(AbortSignal.timeout(5_000))).rejects.toMatchObject({
      name: 'OutlookTokenUnavailableError',
      tokenCode: 'NOT_ON_OWA',
    })
  })
})
