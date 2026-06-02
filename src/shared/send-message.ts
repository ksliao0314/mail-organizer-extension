// Thin chrome.runtime.sendMessage wrapper with discriminated union return.
//
// Extracted to a shared module so popup components (notably the
// onboarding wizard) can use it without circular-importing from
// popup/App.tsx, where the original definition lives. Identical
// signature so call sites in App.tsx need no change.

type Ok<T> = { ok: true; data?: T }
type Err = { ok: false; code: string; message: string }

export async function send<T>(req: unknown): Promise<Ok<T> | Err> {
  return chrome.runtime.sendMessage(req)
}
