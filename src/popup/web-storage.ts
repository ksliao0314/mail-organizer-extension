// Crash-proof Web Storage access for the popup + options pages (2026-11-05).
//
// Why a helper: when these pages run inside the OWA floating-button iframe
// and the browser blocks third-party storage, merely READING the
// `window.localStorage` / `window.sessionStorage` getter throws a
// SecurityError — so even `typeof sessionStorage === 'undefined'` (which
// evaluates the getter) throws, and an unguarded read inside a useState
// initializer takes the whole React tree down with it. Everything stored
// here is a per-viewer convenience (a dismissed hint, a remembered filter),
// so "storage unavailable" must degrade to "not remembered", never crash.

export type WebStorageArea = 'local' | 'session'

function resolveArea(area: WebStorageArea): Storage | null {
  try {
    if (typeof window === 'undefined') return null
    const s = area === 'local' ? window.localStorage : window.sessionStorage
    return s ?? null
  } catch {
    return null
  }
}

export function readWebStorage(area: WebStorageArea, key: string): string | null {
  try {
    return resolveArea(area)?.getItem(key) ?? null
  } catch {
    return null
  }
}

export function writeWebStorage(area: WebStorageArea, key: string, value: string): void {
  try {
    resolveArea(area)?.setItem(key, value)
  } catch {
    // quota exceeded / storage disabled — the value is simply not remembered
  }
}
