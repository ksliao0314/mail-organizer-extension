// batch-3: on-demand full-body fetch (getMessageBody) + the regression lock
// that keeps Body/UniqueBody OUT of the shared list $select.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { OutlookApi } from '@/shared/outlook-api'

let fetchSpy: ReturnType<typeof vi.fn>

beforeEach(() => {
  fetchSpy = vi.fn()
  ;(globalThis as unknown as { fetch: typeof fetch }).fetch = fetchSpy as unknown as typeof fetch
})
afterEach(() => vi.restoreAllMocks())

function makeApi(): OutlookApi {
  return new OutlookApi(async () => 'fake-token')
}
function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), { status: 200 })
}

describe('getMessageBody', () => {
  it('requests UniqueBody with Prefer:text and returns the plain text', async () => {
    fetchSpy.mockResolvedValueOnce(
      jsonResponse({ UniqueBody: { ContentType: 'text', Content: '本件 112訴204 敬請出席' } }),
    )
    const api = makeApi()
    const text = await api.getMessageBody('msg-1')
    expect(text).toBe('本件 112訴204 敬請出席')

    const [url, init] = fetchSpy.mock.calls[0]!
    expect(url).toContain('/me/messages/msg-1')
    expect(url).toContain('$select=UniqueBody')
    expect((init as RequestInit).headers).toMatchObject({
      Prefer: 'outlook.body-content-type="text"',
    })
  })

  it('strips tags when the server ignores Prefer and returns HTML', async () => {
    fetchSpy.mockResolvedValueOnce(
      jsonResponse({
        UniqueBody: { ContentType: 'html', Content: '<div><p>本件 112訴204</p><br></div>' },
      }),
    )
    const api = makeApi()
    const text = await api.getMessageBody('msg-2')
    expect(text).toBe('本件 112訴204')
  })

  it('returns empty string when the body is missing', async () => {
    fetchSpy.mockResolvedValueOnce(jsonResponse({}))
    const api = makeApi()
    expect(await api.getMessageBody('msg-3')).toBe('')
  })
})

describe('MESSAGE_SELECT_DEFAULT regression lock', () => {
  it('list calls NEVER $select Body or UniqueBody (would blow up list responses)', async () => {
    fetchSpy.mockResolvedValueOnce(jsonResponse({ value: [] }))
    const api = makeApi()
    await api.listInboxMessages()

    const url = String(fetchSpy.mock.calls[0]![0])
    expect(url).toContain('$select=')
    expect(url).not.toContain('UniqueBody')
    // "Body" as a standalone select token — BodyPreview (which IS allowed)
    // must not trip this.
    expect(/(?<![A-Za-z])Body(?![A-Za-z])/.test(url)).toBe(false)
    expect(url).toContain('BodyPreview') // sanity: the cheap preview stays
  })
})
