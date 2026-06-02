// Dynamic toolbar-icon behaviour:
//
//   - When at least one Outlook web tab is open anywhere → keep the
//     default_popup wired. Clicking the icon shows the popup UI.
//   - When NO Outlook tab is open → unset the popup. Clicking the icon
//     fires action.onClicked, which opens a new Outlook tab.
//
// This solves the chicken-and-egg "extension can't connect to Outlook
// because I haven't opened Outlook yet" problem — the icon becomes a
// one-click shortcut to the workspace the lawyer actually needs.
//
// Implementation notes:
//   - chrome.action.setPopup({popup:''}) makes onClicked fire on next
//     click. setPopup({popup:'src/popup/index.html'}) restores the popup.
//   - We listen on tabs.onCreated / onUpdated / onRemoved and re-evaluate
//     after each. Cheap (single chrome.tabs.query call per event).
//   - Top-level listener registration so MV3 SW idle-kill + wake doesn't
//     leave the icon in a stale state — handlers fire whenever Chrome
//     decides, and we re-derive the correct popup config each time.

const OWA_TAB_PATTERNS = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
]

const POPUP_PATH = 'src/popup/index.html'

/**
 * Where to send the user when they click the icon with no OWA tab open.
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
 * Returns true when at least one tab is open at any of the OWA hosts.
 * Doesn't care about active / focused — any window, any tab counts.
 */
async function hasOpenOwaTab(): Promise<boolean> {
  try {
    const tabs = await chrome.tabs.query({ url: OWA_TAB_PATTERNS })
    return tabs.length > 0
  } catch (e) {
    // chrome.tabs.query can fail in early SW startup (host_permissions
    // not yet materialised) — assume "no OWA" so the icon at least
    // offers a way to open one.
    console.debug('[mail-organizer] hasOpenOwaTab query failed', e)
    return false
  }
}

/**
 * Set the icon's popup config based on current OWA tab presence.
 *
 *   - OWA open  → popup wired ('src/popup/index.html')
 *   - OWA gone  → popup blank ('') → next click fires onClicked
 *
 * Idempotent — calling repeatedly with the same state is a no-op as far
 * as user-visible behaviour goes. chrome.action.setPopup writes are
 * cheap (no storage IO).
 */
async function syncPopupConfig(): Promise<void> {
  const owaOpen = await hasOpenOwaTab()
  try {
    await chrome.action.setPopup({
      popup: owaOpen ? POPUP_PATH : '',
    })
  } catch (e) {
    console.warn('[mail-organizer] setPopup failed', e)
  }
}

// Debounce so tabs.onUpdated bursts (loading → complete → favicon → title
// changes per page nav) collapse into one sync. 150 ms feels instant to
// the user but absorbs typical event clusters from a single page load.
let syncDebounceTimer: ReturnType<typeof setTimeout> | null = null
function scheduleSyncPopupConfig(): void {
  if (syncDebounceTimer) clearTimeout(syncDebounceTimer)
  syncDebounceTimer = setTimeout(() => {
    syncDebounceTimer = null
    void syncPopupConfig()
  }, 150)
}

/**
 * onClicked handler — only fires when popup is unset (i.e. no OWA tab,
 * according to the last syncPopupConfig sweep).
 *
 * Bug #P: setPopup state persists across SW restarts. If SW was killed
 * with setPopup='' and the user manually opened OWA in another window
 * before syncPopupConfig had a chance to re-run, the persisted '' wins
 * and onClicked fires when the user clicks. Without a live check here
 * we'd open a duplicate Outlook tab — annoying.
 *
 * So we re-query OWA tabs RIGHT BEFORE acting. If one exists, focus it
 * (defense against the stale-popup-state race). Otherwise, open new.
 */
async function handleIconClick(): Promise<void> {
  try {
    // Live check — never trust the popup config alone.
    const existing = await chrome.tabs.query({ url: OWA_TAB_PATTERNS })
    if (existing.length > 0) {
      // Pick the most-recently-active one (heuristic: highest `id` is
      // usually the freshest tab, but `lastAccessed` would be ideal —
      // not all Chrome versions populate it, so fall back to first).
      const tab = existing[0]!
      if (tab.id !== undefined) {
        await chrome.tabs.update(tab.id, { active: true })
      }
      if (tab.windowId !== undefined) {
        await chrome.windows.update(tab.windowId, { focused: true })
      }
      // Resync popup config so next click goes the popup path. The
      // tabs.onUpdated event from .update should fire this anyway, but
      // do it eagerly to close the race.
      void syncPopupConfig()
      return
    }
    // No OWA tab anywhere — genuinely need a new one.
    // chrome.tabs.create rather than chrome.windows.create — same
    // window, less disruption to the user's existing layout.
    await chrome.tabs.create({ url: DEFAULT_OUTLOOK_URL, active: true })
  } catch (e) {
    console.warn('[mail-organizer] failed to open / focus Outlook tab', e)
  }
}

/**
 * Wire up the listeners. Call at SW module load (every wake-up).
 *
 * We listen to:
 *   - tabs.onCreated: a new tab might be OWA (we should wire popup back)
 *   - tabs.onUpdated: a tab navigated (could become OWA, or leave OWA)
 *   - tabs.onRemoved: a tab closed (could have been the only OWA tab)
 *   - windows.onFocusChanged: closing a window can take OWA tabs with
 *     it; onRemoved fires per tab but defensively re-check on window
 *     focus loss too.
 *
 * Each listener calls syncPopupConfig — cheap, idempotent.
 */
export function installActionRouter(): void {
  chrome.action.onClicked.addListener(() => {
    void handleIconClick()
  })

  chrome.tabs.onCreated.addListener(scheduleSyncPopupConfig)

  // Listen to ALL onUpdated events (not just changeInfo.url), debounced.
  //   Why: when an OWA tab navigates AWAY to a non-permitted URL (e.g.
  //   user goes outlook.office.com → google.com), changeInfo.url is
  //   undefined (we lack permission to see the new URL). A URL-only
  //   guard would miss this transition, leaving setPopup in "popup
  //   wired" state even though no OWA tab exists. The debounce absorbs
  //   the storm of non-URL onUpdated events (favicon, title, loading).
  chrome.tabs.onUpdated.addListener(scheduleSyncPopupConfig)

  chrome.tabs.onRemoved.addListener(scheduleSyncPopupConfig)

  // Initial sync on SW startup — covers the case where Chrome boots
  // with OWA already open in a restored session, or the SW was idle-
  // killed and woke without any tab events firing.
  void syncPopupConfig()
}
