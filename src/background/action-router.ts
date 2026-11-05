// Toolbar-icon → Outlook routing.
//
// Goal (unchanged): clicking the icon with no Outlook web tab open should
// land the lawyer in Outlook — the extension can't do anything until OWA is
// open, so the icon doubles as a one-click shortcut to the workspace.
//
// Mechanism (2026-11-05 rewrite):
//   - The toolbar popup is ALWAYS wired. When the TOOLBAR popup opens and
//     getStatus reports owaConnected=false, it sends `openOwa` (→
//     openOrFocusOwa below) and closes itself (popup/App.tsx). The in-OWA
//     iframe panel never does this — it is by definition already inside
//     Outlook.
//   - Previously this module swapped chrome.action.setPopup('') / popup path
//     from tabs.onCreated / onUpdated / onRemoved listeners so a click with no
//     OWA tab fired action.onClicked instead. Registering those listeners made
//     Chrome wake this service worker on practically every navigation in every
//     tab, browser-wide, just to re-run one tabs.query. They are gone.
//
// Notes:
//   - setPopup state persists across SW restarts AND browser restarts (Bug #P).
//     A user upgrading from a build that last left setPopup('') would get a
//     dead icon (no popup, and onClicked only opens Outlook), so we re-assert
//     the popup path once per SW start. Idempotent, no storage IO, wakes nothing.
//   - action.onClicked stays registered as a safety net: it only fires while
//     the popup is unset, which should now never happen — but if it does, the
//     click still does the useful thing.

const OWA_TAB_PATTERNS = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
]

const POPUP_PATH = 'src/popup/index.html'

/**
 * Where to send the user when no OWA tab is open.
 *
 * `/mail/` is the universal entry — for work accounts it lands on the
 * inbox directly; for personal Microsoft accounts Microsoft will detect
 * and redirect to outlook.live.com (which falls OUTSIDE our
 * host_permissions, so the extension's content scripts won't work there
 * — but the user will see Outlook is up, can switch to their work
 * profile if applicable, etc).
 *
 * Future enhancement: persist the LAST KNOWN OWA origin (cloud.microsoft
 * vs office.com vs office365.com) when we see one, and prefer that on
 * next open. For now the office.com entry is the safest universal default.
 */
const DEFAULT_OUTLOOK_URL = 'https://outlook.office.com/mail/'

/**
 * Focus an existing Outlook web tab, or open one.
 *
 * Re-queries tabs RIGHT BEFORE acting rather than trusting the caller's view
 * (the popup's getStatus may be a moment old, and the user may have opened
 * OWA in another window meanwhile) — never open a duplicate Outlook tab when
 * one exists.
 */
export async function openOrFocusOwa(): Promise<'focused' | 'created'> {
  const existing = await chrome.tabs.query({ url: OWA_TAB_PATTERNS })
  if (existing.length > 0) {
    // `lastAccessed` would pick the freshest tab, but not all Chromium
    // versions populate it — first match is good enough.
    const tab = existing[0]!
    if (tab.id !== undefined) {
      await chrome.tabs.update(tab.id, { active: true })
    }
    if (tab.windowId !== undefined) {
      await chrome.windows.update(tab.windowId, { focused: true })
    }
    return 'focused'
  }
  // chrome.tabs.create rather than chrome.windows.create — same window,
  // less disruption to the user's existing layout.
  await chrome.tabs.create({ url: DEFAULT_OUTLOOK_URL, active: true })
  return 'created'
}

/** Wire up at SW module load (every wake-up). Registers NO tab listeners. */
export function installActionRouter(): void {
  chrome.action.onClicked.addListener(() => {
    void openOrFocusOwa().catch((e) =>
      console.warn('[mail-organizer] failed to open / focus Outlook tab', e),
    )
  })

  void Promise.resolve()
    .then(() => chrome.action.setPopup({ popup: POPUP_PATH }))
    .catch((e) => console.warn('[mail-organizer] setPopup failed', e))
}
