// Regression tests for the 2026-05-27 404-as-skipped behaviour:
//
// The companion fix to non-idempotent-retry suppression: when a move /
// delete returns 404 ErrorItemNotFound, the message has effectively
// already moved (either by us in a now-lost first response, by another
// rule, or by the user in OWA). Re-running the same operation will hit
// the same 404. Mark as `skipped` so the UI doesn't show a red error
// row inviting a futile retry.
//
// `isAlreadyMovedError` is exported specifically so we can unit-test
// the discrimination here — vs invoking the full executeItem (which
// would require a real OutlookApi + chrome.storage harness).

import { describe, expect, it } from 'vitest'
import { isAlreadyMovedError } from '@/background/execute'
import { OutlookError } from '@/shared/outlook-api'

describe('isAlreadyMovedError', () => {
  it('returns true for 404 with ErrorItemNotFound in the message', () => {
    const err = new OutlookError(
      404,
      "Outlook REST POST /me/messages/abc/move → 404 Not Found {\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"The specified object was not found in the store.\"}}",
    )
    expect(isAlreadyMovedError(err)).toBe(true)
  })

  it('returns true for 404 with case-insensitive ErrorItemNotFound', () => {
    // Belt-and-braces — the marker check uses /i flag.
    const err = new OutlookError(
      404,
      'errorItemNotFound: stale id',
    )
    expect(isAlreadyMovedError(err)).toBe(true)
  })

  it('returns false for 404 WITHOUT the marker (e.g. folder gone)', () => {
    // A 404 without ErrorItemNotFound is some other issue — folder
    // path typo, deleted folder, mailbox access problem. We don't
    // want to swallow those as "already moved"; they should surface
    // normally so the user can diagnose.
    const err = new OutlookError(
      404,
      'Outlook REST GET /me/MailFolders/foo → 404 Not Found ErrorFolderNotFound',
    )
    expect(isAlreadyMovedError(err)).toBe(false)
  })

  it('returns false for non-404 OutlookErrors even if message mentions ErrorItemNotFound', () => {
    // Status code is the load-bearing check — only 404 means "thing is
    // gone". A 500 with "ErrorItemNotFound" in the body is a server
    // bug, not a stable already-moved signal.
    const err = new OutlookError(
      500,
      'Server bug: ErrorItemNotFound in 500 response',
    )
    expect(isAlreadyMovedError(err)).toBe(false)
  })

  it('returns false for plain Errors (not OutlookError)', () => {
    expect(isAlreadyMovedError(new Error('ErrorItemNotFound'))).toBe(false)
  })

  it('returns false for non-Error values (defensive)', () => {
    expect(isAlreadyMovedError(undefined)).toBe(false)
    expect(isAlreadyMovedError(null)).toBe(false)
    expect(isAlreadyMovedError('ErrorItemNotFound')).toBe(false)
    expect(isAlreadyMovedError({ status: 404, message: 'ErrorItemNotFound' })).toBe(false)
  })

  it('returns true for the exact production error text the user reported', () => {
    // Reproduces the specific failure observed in the field 2026-05-27:
    //   "[404] Outlook REST POST /me/messages/AQMk...PAAHJX2_CgAAAA==/move
    //    → 404 {"error":{"code":"ErrorItemNotFound","message":"The
    //    specified object was not found in the store., The process
    //    failed to get the correct properties."}}"
    const err = new OutlookError(
      404,
      '[404] Outlook REST POST /me/messages/AQMkADcwMWM5/move → 404 {"error":{"code":"ErrorItemNotFound","message":"The specified object was not found in the store., The process failed to get the correct properties."}}',
    )
    expect(isAlreadyMovedError(err)).toBe(true)
  })
})
