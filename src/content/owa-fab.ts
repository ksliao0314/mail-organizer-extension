// Floating action button injected on OWA pages.
//
// Parallel entry point to the Chrome toolbar icon: the lawyer works entirely
// inside the Outlook tab, and reaching the toolbar means moving the mouse to
// the top-right of the whole browser. Putting our launcher in OWA's lower-
// right (next to Microsoft's Copilot button) keeps everything one click away.
//
// Architecture:
//   - One Shadow DOM host (#mail-organizer-fab-host) appended to <body>.
//   - Shadow tree contains the FAB <button> + an <iframe> panel pointing at
//     our existing popup index.html. The popup runs in its native chrome-
//     extension:// origin so chrome.* APIs, MSAL token fetch, and all the
//     existing classify / execute flows work without any change.
//   - Click outside / Escape collapses the panel. Click the FAB again toggles.
//   - chrome.storage.onChanged hot-reloads the FAB when the user flips the
//     "show OWA floating icon" setting in options — no page reload needed.

const HOST_ID = 'mail-organizer-fab-host'
const LOG_PREFIX = '[mail-organizer-fab]'

// Visual: white pill matching the Copilot launcher next to us. Icon is
// lucide-react's "folder-down" — represents "organize mail into folders",
// matches what the tool actually does and reads less like another AI
// launcher (a wand-sparkles icon next to Copilot was being mistaken for
// a Copilot extension).
const ICON_SVG = `
<svg width="22" height="22" viewBox="0 0 24 24" fill="none"
     stroke="#111110" stroke-width="2"
     stroke-linecap="round" stroke-linejoin="round"
     aria-hidden="true">
  <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>
  <path d="M12 10v6"/>
  <path d="m15 13-3 3-3-3"/>
</svg>`

// Try to locate the Copilot button so we can anchor our FAB directly under
// it. OWA renders this button with an aria-label in the active locale —
// english "Copilot", traditional chinese 「Copilot」 (the brand name stays
// in latin chars), so a case-insensitive substring on the latin word works
// across locales. Fall back through a couple of other known attributes
// before giving up.
function findCopilotAnchor(): HTMLElement | null {
  const selectors = [
    '[aria-label*="Copilot" i]',
    '[data-icon-name="Copilot" i]',
    '[data-app-name*="Copilot" i]',
    '[title*="Copilot" i]',
  ]
  for (const sel of selectors) {
    const all = document.querySelectorAll<HTMLElement>(sel)
    for (const el of Array.from(all)) {
      // Skip invisible matches (some are hidden helper nodes)
      if (el.offsetParent === null) continue
      const r = el.getBoundingClientRect()
      if (r.width < 16 || r.height < 16) continue
      return el
    }
  }
  return null
}


function isOwaTopFrame(): boolean {
  // Skip if we're somehow loaded inside an iframe (OWA has nested frames for
  // reading-pane previews, etc.). The FAB only makes sense at the top level.
  return window.top === window.self
}

// Find a treeitem element by its OWA folder display name.
//
// OWA normalizes ASCII letters in `data-folder-name` to lowercase
// (verified empirically: "CompanyB" → "company-b", "公司甲v.公司乙ABC..." →
// "公司甲v.公司乙abc..."), while CJK characters are preserved verbatim. Our
// stored folder names come from Graph API which keeps the original case,
// so we have to lowercase both sides for the comparison.
//
// Fallback selectors handle the case where Microsoft renames the
// data-folder-name attribute in a future OWA build — we still want the
// FAB navigation to work rather than silently fail. We log the fallback
// hit so a regression is visible in console.
function findOwaTreeitem(folderName: string): HTMLElement | null {
  const target = folderName.toLowerCase()

  const items = document.querySelectorAll<HTMLElement>(
    '[role="treeitem"][data-folder-name]',
  )
  for (const el of Array.from(items)) {
    const attr = el.getAttribute('data-folder-name')
    if (attr && attr.toLowerCase() === target) return el
  }

  // Fallback: aria-label match. OWA's treeitems usually have an aria-label
  // like "案件名稱" or "案件名稱, 已讀, 5 個郵件" — accept exact or prefix.
  const labelled = document.querySelectorAll<HTMLElement>(
    `[role="treeitem"][aria-label]`,
  )
  for (const el of Array.from(labelled)) {
    const label = el.getAttribute('aria-label')?.toLowerCase() ?? ''
    if (label === target || label.startsWith(`${target},`)) {
      console.warn(
        LOG_PREFIX,
        'findOwaTreeitem hit via aria-label fallback — data-folder-name selector may have changed:',
        folderName,
      )
      return el
    }
  }

  return null
}

// Wait until predicate returns true or timeout. Used after click-to-expand
// to wait for child treeitems to render into the DOM.
function waitForDom(check: () => boolean, timeoutMs: number): Promise<boolean> {
  return new Promise((resolve) => {
    if (check()) return resolve(true)
    const start = Date.now()
    const id = window.setInterval(() => {
      if (check()) {
        window.clearInterval(id)
        resolve(true)
      } else if (Date.now() - start > timeoutMs) {
        window.clearInterval(id)
        resolve(false)
      }
    }, 50)
  })
}

// Surface a brief inline toast inside the FAB host's shadow DOM. Used when
// navigation can't find a folder — replaces the old behaviour of falling
// through to location.href which would reload the tab and land the user on
// inbox (because OWA frequently doesn't honor a deep folderId on cold init).
// Auto-dismisses after 6 seconds; manually closable via × button. Re-calling
// before the timer fires replaces any prior toast (no stacking).
function showNavToast(title: string, detail?: string): void {
  const host = document.getElementById(HOST_ID)
  if (!host || !host.shadowRoot) {
    console.warn(LOG_PREFIX, 'no shadow root, falling back to console:', title, detail)
    return
  }
  // Replace any existing toast to avoid stacking.
  host.shadowRoot.querySelectorAll('.nav-toast').forEach((el) => el.remove())

  const toast = document.createElement('div')
  toast.className = 'nav-toast'
  toast.setAttribute('role', 'alert')

  const body = document.createElement('div')
  body.className = 'nav-toast-body'
  const titleEl = document.createElement('div')
  titleEl.className = 'nav-toast-title'
  titleEl.textContent = title
  body.appendChild(titleEl)
  if (detail) {
    const detailEl = document.createElement('div')
    detailEl.className = 'nav-toast-detail'
    detailEl.textContent = detail
    body.appendChild(detailEl)
  }
  toast.appendChild(body)

  const dismiss = document.createElement('button')
  dismiss.className = 'nav-toast-dismiss'
  dismiss.type = 'button'
  dismiss.setAttribute('aria-label', '關閉提示')
  dismiss.textContent = '×'
  let timer: number | null = null
  dismiss.onclick = () => {
    if (timer !== null) window.clearTimeout(timer)
    toast.remove()
  }
  toast.appendChild(dismiss)

  host.shadowRoot.appendChild(toast)
  timer = window.setTimeout(() => {
    toast.remove()
    timer = null
  }, 6000)
}

// Click a treeitem with OWA-appropriate event sequence. OWA's React handlers
// may listen to pointerdown / mousedown / click — fire the full chain.
function clickOwaTreeitem(el: HTMLElement): void {
  el.scrollIntoView({ block: 'nearest', behavior: 'instant' as ScrollBehavior })
  try {
    el.focus()
  } catch {
    /* ignore */
  }
  const opts = {
    bubbles: true,
    cancelable: true,
    view: window,
    button: 0,
    buttons: 1,
  }
  try {
    el.dispatchEvent(new PointerEvent('pointerdown', opts))
    el.dispatchEvent(new MouseEvent('mousedown', opts))
    el.dispatchEvent(new PointerEvent('pointerup', opts))
    el.dispatchEvent(new MouseEvent('mouseup', opts))
    el.dispatchEvent(new MouseEvent('click', opts))
  } catch {
    el.click()
  }
}

// Navigate OWA to a target folder by driving its sidebar DOM. We
// deliberately do NOT fall back to anchor-click / pushState / location.href
// — empirically all three end up reloading the tab AND landing on inbox
// (OWA's cold-init doesn't honor a deep folderId in the URL). Reloading +
// landing on inbox is a worse outcome than failing visibly, so when the
// sidebar walk can't find the folder we show an inline toast and let the
// user expand the sidebar manually.
async function navigateOwaToFolder(
  folderName: string,
  folderPath: string,
  folderId: string,
): Promise<void> {
  console.log(LOG_PREFIX, 'navigate request:', { folderName, folderPath, folderId })

  const segments = folderPath.split('/').filter(Boolean)
  if (segments.length === 0 && !folderName) {
    console.warn(LOG_PREFIX, 'no path/name to navigate, aborting')
    return
  }

  const startPath = location.pathname
  const targetPath = folderId
    ? `/mail/${encodeURIComponent(folderId)}`
    : ''

  console.log(LOG_PREFIX, 'attempting DOM-walk navigation')
  const domClickOk = await tryDomClickNavigation(folderName, segments)
  if (domClickOk) {
    // Best-effort URL detection just for logging — doesn't gate behavior.
    void waitForDom(() => {
      if (location.pathname === startPath) return false
      if (targetPath && !location.pathname.includes(encodeURIComponent(folderId)))
        return false
      return true
    }, 3000).then((ok) => {
      if (ok) console.log(LOG_PREFIX, '✓ URL settled to', location.pathname)
      else console.warn(LOG_PREFIX, 'URL did not update within 3s; OWA may have rejected')
    })
    return
  }

  // DOM walk failed — show a clear inline toast. We refuse to fall through
  // to location.href because OWA's cold-init reliably drops the folderId
  // and lands on inbox (the very symptom this code is preventing).
  console.warn(LOG_PREFIX, 'navigation failed for', folderPath)
  showNavToast(
    `找不到「${folderName}」資料夾`,
    `請手動展開 Outlook 左側 sidebar 到「${folderPath}」、再點本面板的「近日活動」重試。`,
  )
}

// Expand a treeitem WITHOUT navigating to it. Critical that we don't
// navigate — the original implementation called clickOwaTreeitem(ancestor)
// which BOTH expanded AND navigated; mid-traversal URL hops + the long
// wait OWA needs to lazy-load children combined to make ancestors flip
// the page state under our feet.
//
// Empirically verified against outlook.cloud.microsoft (this user's
// build, Nov 2026):
//
//   - ArrowRight keystroke on the focused treeitem expands it cleanly
//     (aria-expanded → "true") with no URL change.
//   - After expansion, OWA fetches the child folder list via Graph API
//     and renders the first batch into DOM. For folders with 50-70+
//     children this can take 5-15 seconds while the renderer churns.
//   - The treeitem contains exactly one `<button class="fui-Button …">`
//     (the chevron) with NO aria-label and NO class containing "chevron"
//     or "expand". Selector-by-label was a dead end — we use the
//     "first button child" heuristic for the fallback.
//   - Plain `clickOwaTreeitem(el)` triggers a "navigate AND expand"
//     combo, so we never use it on ancestors any more. That fallback
//     was actively harmful — it changed `location.pathname` mid-
//     traversal, which downstream code interpreted as "navigation
//     succeeded" and stopped trying.
async function expandOwaTreeitem(
  el: HTMLElement,
  childName: string,
): Promise<boolean> {
  if (el.getAttribute('aria-expanded') === 'true' && findOwaTreeitem(childName)) {
    return true
  }
  el.scrollIntoView({ block: 'nearest', behavior: 'instant' as ScrollBehavior })
  try {
    el.focus()
  } catch {
    /* ignore */
  }

  // Primary: ARIA Right-Arrow keystroke. Verified to work on the user's
  // outlook.cloud.microsoft build.
  const arrowOpts: KeyboardEventInit = {
    bubbles: true,
    cancelable: true,
    key: 'ArrowRight',
    code: 'ArrowRight',
    keyCode: 39,
  }
  try {
    el.dispatchEvent(new KeyboardEvent('keydown', arrowOpts))
    el.dispatchEvent(new KeyboardEvent('keyup', arrowOpts))
  } catch {
    /* ignore */
  }
  // Wait generously — OWA lazy-loads child folder list on first expand
  // and a parent with 50+ children can take many seconds to populate.
  // 15s is far longer than the renderer normally needs but well short of
  // a user perceiving the action as "stuck".
  if (await waitForDom(() => findOwaTreeitem(childName) !== null, 15_000)) {
    return true
  }

  // Fallback: click the chevron button. OWA renders exactly one button
  // child inside each treeitem (the chevron icon) — `fui-Button` is
  // Microsoft's Fluent UI class. No aria-label so we identify by
  // "first button descendant".
  const chevron = el.querySelector<HTMLButtonElement>(
    'button.fui-Button, button[class*="fui-Button"], button',
  )
  if (chevron && chevron !== el) {
    try {
      chevron.click()
    } catch {
      /* ignore */
    }
    if (await waitForDom(() => findOwaTreeitem(childName) !== null, 10_000)) {
      return true
    }
  }

  // Out of safe options. We deliberately do NOT fall through to a
  // navigating click on the treeitem itself — that would change the URL
  // mid-traversal and corrupt the user's view (the very symptom we
  // are fixing). Caller's failure path shows a toast.
  return false
}

// Module-level throttle so concurrent prewarm requests don't pile up
// (popup might re-trigger on Recent Activity storage change). We don't
// need queueing — overlapping calls just merge into one in-flight pass.
let prewarmInFlight: Promise<void> | null = null

// Pipeline-style optimisation: when the popup loads Recent Activity, it
// ships the visible folderPaths here so we can silently pre-expand each
// ancestor chain. Subsequent navigate-folder clicks then short-circuit
// inside expandOwaTreeitem (aria-expanded already true + child found)
// and complete in <100ms instead of the 5-15s cold-load.
//
// Sequential, not parallel: OWA's sidebar reacts poorly to overlapping
// expand events and the per-folder lazy fetch already saturates its
// renderer. We chain ancestor expansions, accepting that the slowest
// path dominates total time.
//
// Idempotent by design — `expandOwaTreeitem` checks aria-expanded and
// child presence before doing any work, so a second prewarm pass is
// effectively free.
async function prewarmSidebarPaths(paths: string[]): Promise<void> {
  if (prewarmInFlight) return prewarmInFlight
  prewarmInFlight = (async () => {
    console.log(LOG_PREFIX, 'prewarm: starting for', paths.length, 'paths')
    let expanded = 0
    let skipped = 0
    for (const path of paths) {
      const segments = path.split('/').filter(Boolean)
      if (segments.length <= 1) continue // top-level — no ancestor to expand
      // Walk ancestors. Stop early if any expansion fails — deeper
      // levels can't be reached anyway.
      for (let i = 0; i < segments.length - 1; i++) {
        const ancestorName = segments[i]!
        const childName = segments[i + 1]!
        const el = findOwaTreeitem(ancestorName)
        if (!el) break
        if (el.getAttribute('aria-expanded') === 'true' && findOwaTreeitem(childName)) {
          skipped++
          continue
        }
        const ok = await expandOwaTreeitem(el, childName)
        if (!ok) break
        expanded++
      }
    }
    console.log(LOG_PREFIX, 'prewarm: done — expanded', expanded, 'skipped', skipped)
  })().finally(() => {
    prewarmInFlight = null
  })
  return prewarmInFlight
}

// DOM-walk navigation: descend the sidebar tree, expanding each ancestor
// via `expandOwaTreeitem` so we don't accidentally navigate to the
// ancestor itself, then click the leaf to select+navigate.
//
// Generous timeouts because OWA virtualizes deeper levels of the tree —
// children only render after the parent expands AND the section scrolls
// into view. Tight timeouts were the original cause of "fall through to
// reload" hits the user reported.
//
// Returns true if we successfully clicked the leaf treeitem. Caller does
// NOT fall back to URL-based strategies on false return — those reliably
// land on inbox after OWA's cold init, which is worse than failing visibly.
async function tryDomClickNavigation(
  folderName: string,
  segments: string[],
): Promise<boolean> {
  // Expand each ancestor if collapsed
  for (let i = 0; i < segments.length - 1; i++) {
    const ancestorName = segments[i]!
    let ancestorEl = findOwaTreeitem(ancestorName)
    if (!ancestorEl) {
      // Top-level might be off-screen in a long sidebar — extend the wait
      // to give virtualized rendering a chance.
      const appeared = await waitForDom(
        () => findOwaTreeitem(ancestorName) !== null,
        2000,
      )
      if (!appeared) {
        console.warn(LOG_PREFIX, 'ancestor not in DOM after 2s:', ancestorName)
        return false
      }
      ancestorEl = findOwaTreeitem(ancestorName)
    }
    if (!ancestorEl) return false
    const nextName = segments[i + 1]!
    if (ancestorEl.getAttribute('aria-expanded') === 'true' && findOwaTreeitem(nextName)) {
      // Already expanded and child rendered — skip expansion step.
      continue
    }
    console.log(LOG_PREFIX, 'expanding ancestor:', ancestorName)
    const opened = await expandOwaTreeitem(ancestorEl, nextName)
    if (!opened) {
      console.warn(LOG_PREFIX, 'expand did not reveal child:', nextName)
      return false
    }
  }

  // Click leaf
  const leafName = folderName || segments[segments.length - 1] || ''
  if (!leafName) return false
  let target = findOwaTreeitem(leafName)
  if (!target) {
    // Leaf may not be in the first render batch of the parent's
    // lazy-loaded children. OWA appends children incrementally as it
    // fetches them, so we give it the same long window as the ancestor
    // expansion above.
    const appeared = await waitForDom(
      () => findOwaTreeitem(leafName) !== null,
      10_000,
    )
    if (!appeared) {
      console.warn(LOG_PREFIX, 'target treeitem not found after 10s:', leafName)
      // Diagnostic dump: what data-folder-name values ARE present?
      const allNames = Array.from(
        document.querySelectorAll<HTMLElement>('[role="treeitem"][data-folder-name]'),
      )
        .map((el) => el.getAttribute('data-folder-name') || '')
        .filter((n) => n.length > 0)
      const prefix = leafName.slice(0, Math.max(1, Math.floor(leafName.length / 3)))
      const closeMatches = allNames.filter((n) => n.includes(prefix.toLowerCase()))
      console.warn(
        LOG_PREFIX,
        'looked for exact:',
        JSON.stringify(leafName),
        '— close matches:',
        closeMatches,
      )
      return false
    }
    target = findOwaTreeitem(leafName)
  }
  if (!target) return false

  console.log(LOG_PREFIX, 'clicking target treeitem:', leafName)
  target.scrollIntoView({ block: 'nearest', behavior: 'instant' as ScrollBehavior })
  clickOwaTreeitem(target)
  return true
}

async function shouldShow(): Promise<boolean> {
  try {
    const r = await chrome.storage.local.get('settings')
    const s = r.settings as { showOwaFab?: boolean } | undefined
    // Default ON — first-run users get the FAB without having to opt in.
    const show = s?.showOwaFab !== false
    console.log(LOG_PREFIX, 'shouldShow', { settings: s, show })
    return show
  } catch (e) {
    console.warn(LOG_PREFIX, 'shouldShow read failed, defaulting ON', e)
    return true
  }
}

function buildShadow(host: HTMLElement): {
  fab: HTMLButtonElement
  panelBackdrop: HTMLDivElement
  panel: HTMLDivElement
  iframe: HTMLIFrameElement
  optionsBackdrop: HTMLDivElement
  optionsBack: HTMLButtonElement
  optionsClose: HTMLButtonElement
  optionsIframe: HTMLIFrameElement
} {
  const shadow = host.attachShadow({ mode: 'open' })
  // Styles scoped to the shadow root — OWA's global CSS can't reach in,
  // and our styles can't leak out.
  const style = document.createElement('style')
  style.textContent = `
    :host { all: initial; }
    .fab {
      position: fixed;
      /* top / right set dynamically by positionFab() — anchored to OWA's
         Copilot button when found, fallback top-right estimate otherwise. */
      top: 220px;
      right: 24px;
      width: 44px;
      height: 44px;
      border-radius: 50%;
      border: 1px solid rgba(0, 0, 0, 0.06);
      background: white;
      cursor: pointer;
      /* Match the Copilot launcher's soft pillow shadow — light + ambient. */
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08), 0 1px 3px rgba(0, 0, 0, 0.06);
      display: flex;
      align-items: center;
      justify-content: center;
      transition: transform 0.15s ease, box-shadow 0.15s ease;
      pointer-events: auto;
      font-family: system-ui, -apple-system, "Segoe UI", "Noto Sans TC", sans-serif;
    }
    .fab:hover {
      transform: translateY(-1px);
      box-shadow: 0 6px 18px rgba(0, 0, 0, 0.12), 0 2px 4px rgba(0, 0, 0, 0.08);
    }
    .fab:active { transform: translateY(0); }
    .fab:focus-visible { outline: 2px solid #5B7EE5; outline-offset: 2px; }

    /* Panel modal: dim backdrop + centered card the same size as the
       settings modal below, so both feel like the same UI affordance.
       Earlier iterations tried to fill the OWA reading-pane region, but
       a fixed-size centered modal gives a more predictable + consistent
       experience between "open plan panel" and "open settings". */
    .panel-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(15, 15, 15, 0.45);
      display: none;
      align-items: center;
      justify-content: center;
      z-index: 2147483646;
      pointer-events: auto;
      animation: fadeIn 140ms ease-out;
    }
    .panel-backdrop.open { display: flex; }

    .panel {
      width: min(1100px, calc(100vw - 48px));
      height: min(800px, calc(100vh - 48px));
      background: white;
      border-radius: 14px;
      box-shadow: 0 20px 60px rgba(0, 0, 0, 0.25);
      overflow: hidden;
      pointer-events: auto;
      animation: modalIn 180ms ease-out;
    }

    .iframe {
      width: 100%;
      height: 100%;
      border: none;
      display: block;
    }

    /* Options modal — opens when the popup posts the 'open-options' message.
       Full-viewport dim backdrop + centered iframe loading our options page.
       Sized generously because options has 6 sections + sidebar nav. */
    .options-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(15, 15, 15, 0.45);
      display: none;
      align-items: center;
      justify-content: center;
      z-index: 2147483647;
      pointer-events: auto;
      animation: fadeIn 140ms ease-out;
    }
    .options-backdrop.open { display: flex; }

    .options-modal {
      width: min(1100px, calc(100vw - 48px));
      height: min(800px, calc(100vh - 48px));
      background: white;
      border-radius: 14px;
      box-shadow: 0 20px 60px rgba(0, 0, 0, 0.25);
      overflow: hidden;
      display: flex;
      flex-direction: column;
      animation: modalIn 180ms ease-out;
    }

    .options-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 14px;
      border-bottom: 1px solid rgba(0, 0, 0, 0.08);
      background: white;
    }

    /* Left "back to main" affordance — same behavior as the × on the right
       (both close the settings modal and reveal the plan panel underneath),
       but the icon + text wording makes it obvious that you're navigating
       back, not closing the whole extension. */
    .options-back {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      border-radius: 6px;
      border: none;
      background: transparent;
      cursor: pointer;
      color: #444;
      font-size: 12px;
      font-family: system-ui, -apple-system, "Segoe UI", "Noto Sans TC", sans-serif;
    }
    .options-back:hover { background: rgba(0, 0, 0, 0.06); color: #111; }
    .options-back svg { width: 14px; height: 14px; flex-shrink: 0; }

    .options-close {
      width: 28px;
      height: 28px;
      border-radius: 6px;
      border: none;
      background: transparent;
      cursor: pointer;
      color: #666;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
      line-height: 1;
    }
    .options-close:hover { background: rgba(0, 0, 0, 0.06); color: #111; }

    .options-iframe {
      flex: 1;
      width: 100%;
      border: none;
      display: block;
    }

    /* Nav-failure toast: bottom-center, slides up. Used when DOM-click
       navigation can't find the folder in OWA's sidebar and we refuse to
       fall through to a destructive page reload. */
    .nav-toast {
      position: fixed;
      bottom: 24px;
      left: 50%;
      transform: translateX(-50%);
      max-width: 480px;
      padding: 10px 14px;
      background: #1f2937;
      color: white;
      border-radius: 8px;
      font-size: 13px;
      line-height: 1.45;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
      pointer-events: auto;
      z-index: 2147483647;
      animation: toastIn 200ms ease-out;
      font-family: system-ui, -apple-system, "Segoe UI", "Noto Sans TC", sans-serif;
      display: flex;
      align-items: flex-start;
      gap: 12px;
    }
    .nav-toast .nav-toast-body { flex: 1; min-width: 0; }
    .nav-toast .nav-toast-title { font-weight: 600; margin-bottom: 2px; }
    .nav-toast .nav-toast-detail { opacity: 0.85; font-size: 12px; }
    .nav-toast .nav-toast-dismiss {
      background: transparent;
      border: none;
      color: white;
      cursor: pointer;
      opacity: 0.7;
      font-size: 18px;
      line-height: 1;
      padding: 0 4px;
    }
    .nav-toast .nav-toast-dismiss:hover { opacity: 1; }

    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
    @keyframes modalIn {
      from { opacity: 0; transform: scale(0.97); }
      to { opacity: 1; transform: scale(1); }
    }
    @keyframes toastIn {
      from { opacity: 0; transform: translate(-50%, 8px); }
      to { opacity: 1; transform: translate(-50%, 0); }
    }
  `

  const fab = document.createElement('button')
  fab.className = 'fab'
  fab.type = 'button'
  fab.setAttribute('aria-label', 'Mail Organizer')
  fab.title = 'Mail Organizer — 開啟歸類面板'
  fab.innerHTML = ICON_SVG

  // Backdrop wraps the panel so it can be dim-centered like the settings
  // modal. Click on the dim area (target === backdrop) closes the panel.
  const panelBackdrop = document.createElement('div')
  panelBackdrop.className = 'panel-backdrop'
  const panel = document.createElement('div')
  panel.className = 'panel'

  const iframe = document.createElement('iframe')
  iframe.className = 'iframe'
  iframe.title = 'Mail Organizer'
  // Don't set src yet — defer until the first time the user opens the panel
  // (setOpen(true) in createFab). Two reasons:
  //   1. Skip popup-load cost on every OWA page that may never open the panel
  //   2. Until the iframe has a chrome-extension:// src, other extensions
  //      (notably Claude in Chrome) can still inspect the OWA tab — Chrome's
  //      cross-extension iframe isolation otherwise blocks them.
  panel.appendChild(iframe)
  panelBackdrop.appendChild(panel)

  // Settings overlay — hidden by default, shown when the popup iframe posts
  // 'mail-organizer/open-options'. Larger than the FAB panel because
  // options has many sections.
  const optionsBackdrop = document.createElement('div')
  optionsBackdrop.className = 'options-backdrop'
  const optionsModal = document.createElement('div')
  optionsModal.className = 'options-modal'
  const optionsHeader = document.createElement('div')
  optionsHeader.className = 'options-header'
  // Left: ← 返回主畫面 (close modal, return to the plan panel underneath).
  const optionsBack = document.createElement('button')
  optionsBack.className = 'options-back'
  optionsBack.type = 'button'
  optionsBack.setAttribute('aria-label', '返回主畫面')
  optionsBack.title = '返回主畫面 (Esc)'
  optionsBack.innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
         stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="m12 19-7-7 7-7"/>
      <path d="M19 12H5"/>
    </svg>
    <span>返回主畫面</span>
  `
  // Right: × (same close action — kept for "just close" muscle memory).
  const optionsClose = document.createElement('button')
  optionsClose.className = 'options-close'
  optionsClose.type = 'button'
  optionsClose.setAttribute('aria-label', '關閉設定')
  optionsClose.title = '關閉 (Esc)'
  optionsClose.textContent = '×'
  optionsHeader.appendChild(optionsBack)
  optionsHeader.appendChild(optionsClose)
  const optionsIframe = document.createElement('iframe')
  optionsIframe.className = 'options-iframe'
  // Lazy-load the URL: only set src when first opened so we don't pay the
  // page-load cost on every OWA page just in case the user opens settings.
  optionsIframe.title = 'Mail Organizer — 設定'
  optionsModal.appendChild(optionsHeader)
  optionsModal.appendChild(optionsIframe)
  optionsBackdrop.appendChild(optionsModal)

  shadow.appendChild(style)
  shadow.appendChild(fab)
  shadow.appendChild(panelBackdrop)
  shadow.appendChild(optionsBackdrop)

  return {
    fab,
    panelBackdrop,
    panel,
    iframe,
    optionsBackdrop,
    optionsBack,
    optionsClose,
    optionsIframe,
  }
}

let teardown: (() => void) | null = null

function createFab() {
  if (!isOwaTopFrame()) {
    console.log(LOG_PREFIX, 'skip (not top frame)')
    return
  }
  if (document.getElementById(HOST_ID)) {
    console.log(LOG_PREFIX, 'skip (host already exists)')
    return
  }
  // Wait for body — content_scripts run at document_idle so this is normally
  // ready, but defend against an edge case during SPA early load.
  if (!document.body) {
    console.log(LOG_PREFIX, 'body not ready, defer to DOMContentLoaded')
    window.addEventListener('DOMContentLoaded', () => createFab(), { once: true })
    return
  }

  console.log(LOG_PREFIX, 'creating FAB host')
  const host = document.createElement('div')
  host.id = HOST_ID
  // The host element itself is just a positioning anchor with no painted
  // surface. pointer-events:none lets clicks pass through the empty area to
  // OWA; the FAB / panel set pointer-events:auto on themselves.
  host.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:2147483646;'
  document.body.appendChild(host)

  const {
    fab,
    panelBackdrop,
    iframe,
    optionsBackdrop,
    optionsBack,
    optionsClose,
    optionsIframe,
  } = buildShadow(host)

  let open = false
  let optionsOpen = false
  // Grace timers: after closing, keep the iframe React tree mounted for
  // a short window so quick close→reopen doesn't pay the load cost (and
  // doesn't lose the user's in-flight UI state). After the grace window,
  // we swap to about:blank to release the React app — otherwise its
  // setIntervals keep running for the lifetime of the OWA tab.
  const IFRAME_UNLOAD_GRACE_MS = 30_000
  let panelUnloadTimer: number | null = null
  let optionsUnloadTimer: number | null = null
  const POPUP_URL = chrome.runtime.getURL('src/popup/index.html')
  const OPTIONS_URL = chrome.runtime.getURL('src/options/index.html')

  function setOpen(next: boolean) {
    open = next
    panelBackdrop.classList.toggle('open', open)
    fab.setAttribute('aria-expanded', String(open))
    if (open) {
      if (panelUnloadTimer !== null) {
        window.clearTimeout(panelUnloadTimer)
        panelUnloadTimer = null
      }
      if (!iframe.src || iframe.src === 'about:blank') {
        // Lazy-set the popup iframe src on first open — keeps the OWA page
        // free of cross-extension iframes until actually needed (see
        // buildShadow for rationale).
        iframe.src = POPUP_URL
      }
    } else if (iframe.src && iframe.src !== 'about:blank') {
      panelUnloadTimer = window.setTimeout(() => {
        iframe.src = 'about:blank'
        panelUnloadTimer = null
      }, IFRAME_UNLOAD_GRACE_MS)
    }
  }
  function setOptionsOpen(next: boolean) {
    optionsOpen = next
    optionsBackdrop.classList.toggle('open', optionsOpen)
    if (optionsOpen) {
      if (optionsUnloadTimer !== null) {
        window.clearTimeout(optionsUnloadTimer)
        optionsUnloadTimer = null
      }
      if (!optionsIframe.src || optionsIframe.src === 'about:blank') {
        // Lazy-set src on first open so we don't pay the load cost upfront.
        optionsIframe.src = OPTIONS_URL
      }
    } else if (optionsIframe.src && optionsIframe.src !== 'about:blank') {
      optionsUnloadTimer = window.setTimeout(() => {
        optionsIframe.src = 'about:blank'
        optionsUnloadTimer = null
      }, IFRAME_UNLOAD_GRACE_MS)
    }
  }

  function onFabClick() {
    setOpen(!open)
  }

  function onPanelBackdropClick(e: MouseEvent) {
    // Close only when clicking the dim backdrop itself, not the panel box.
    // (Same pattern as the settings modal — keeps clicks inside the modal
    // from accidentally dismissing it.)
    if (e.target === panelBackdrop) setOpen(false)
  }

  function onDocKey(e: KeyboardEvent) {
    if (e.key !== 'Escape') return
    // Settings modal wins precedence over the plan panel — close whichever
    // is currently open from the outermost layer first.
    if (optionsOpen) {
      setOptionsOpen(false)
      return
    }
    if (open) {
      setOpen(false)
      fab.focus()
    }
  }

  function onPostMessage(e: MessageEvent) {
    // The popup / options iframe sends these. Strict origin + type check so
    // a random OWA postMessage can't trigger us.
    //   iframe-ready:    iframe says "I'm up" — reply with parent-hello so
    //                    it learns our origin and can postMessage back with
    //                    a specific targetOrigin (no more '*' broadcasts).
    //   open-options:    popup's "設定" link → show the modal
    //   close-options:   options page's ESC handler → hide the modal
    //   close-panel:     popup's ESC handler → hide the FAB panel
    //   navigate-folder: popup's "近日活動" row click → ask OWA to navigate
    //                    to that folder in-place (no reload)
    const expectedOrigin = new URL(chrome.runtime.getURL('')).origin
    // Log EVERY incoming message so we can prove the content script is
    // actually receiving them — diagnostic for cases where the user clicks
    // a row but sees nothing happen.
    const inData = (e.data as { type?: string } | undefined) || undefined
    if (inData?.type?.startsWith('mail-organizer/')) {
      console.log(LOG_PREFIX, 'incoming postMessage', {
        origin: e.origin,
        expected: expectedOrigin,
        match: e.origin === expectedOrigin,
        type: inData.type,
      })
    }
    if (e.origin !== expectedOrigin) return
    const data = e.data as
      | {
          type?: string
          folderName?: string
          folderPath?: string
          folderId?: string
          paths?: string[]
        }
      | undefined
    if (data?.type === 'mail-organizer/iframe-ready') {
      // Reply with hello so the iframe learns our origin. Target the
      // extension origin so the reply isn't broadcast.
      const src = e.source as Window | null
      if (src) {
        src.postMessage(
          { type: 'mail-organizer/parent-hello' },
          { targetOrigin: expectedOrigin },
        )
      }
    } else if (data?.type === 'mail-organizer/open-options') {
      setOptionsOpen(true)
    } else if (data?.type === 'mail-organizer/close-options') {
      setOptionsOpen(false)
    } else if (data?.type === 'mail-organizer/close-panel') {
      // If options modal is currently open, close it first (ESC closes the
      // top-most layer). Otherwise close the panel.
      if (optionsOpen) setOptionsOpen(false)
      else setOpen(false)
    } else if (data?.type === 'mail-organizer/navigate-folder') {
      void navigateOwaToFolder(
        data.folderName || '',
        data.folderPath || '',
        data.folderId || '',
      )
    } else if (data?.type === 'mail-organizer/prewarm-sidebar' && Array.isArray(data.paths)) {
      // Pipeline-style optimisation (Option B): popup ships the list of
      // folderPaths visible in 「近日活動」when it mounts; we silently
      // pre-expand the ancestors so the user's later click navigates
      // almost instantly instead of waiting for OWA to lazy-load 70+
      // children. expandOwaTreeitem's idempotent guard (aria-expanded
      // === 'true' && child found) makes re-trigger free.
      void prewarmSidebarPaths(
        data.paths.filter((p): p is string => typeof p === 'string'),
      )
    }
  }

  // Also push parent-hello when an iframe finishes loading — covers the
  // case where our iframe-ready listener isn't ready when the iframe boots.
  function sendHelloToIframe(targetIframe: HTMLIFrameElement) {
    const extOrigin = new URL(chrome.runtime.getURL('')).origin
    targetIframe.contentWindow?.postMessage(
      { type: 'mail-organizer/parent-hello' },
      extOrigin,
    )
  }
  iframe.addEventListener('load', () => sendHelloToIframe(iframe))
  optionsIframe.addEventListener('load', () => sendHelloToIframe(optionsIframe))

  function onOptionsCloseClick() {
    setOptionsOpen(false)
  }

  function onOptionsBackClick() {
    // "← 返回主畫面" — same effect as close: hides the modal, the plan
    // panel underneath becomes visible again. We give the user both
    // affordances (left back arrow, right ×) because settings is reached
    // via the plan panel, so "back" matches the user's mental model
    // better than "close".
    setOptionsOpen(false)
  }

  function onOptionsBackdropClick(e: MouseEvent) {
    // Close only when clicking the dim backdrop itself, not the modal box.
    //
    // User preference (2026-05-26): clicking the backdrop should close
    // EVERYTHING — settings modal + underlying plan panel — not just
    // return to the plan panel. "Outside-click" mental model expects
    // "dismiss the overlay I'm looking at", and from the user's view
    // the FAB panel + options are both overlays on OWA. The explicit
    // ← 返回主畫面 button stays as the "I just want to go back, not
    // close everything" affordance.
    if (e.target === optionsBackdrop) {
      setOptionsOpen(false)
      setOpen(false)
    }
  }

  fab.addEventListener('click', onFabClick)
  panelBackdrop.addEventListener('click', onPanelBackdropClick)
  document.addEventListener('keydown', onDocKey, true)
  window.addEventListener('message', onPostMessage)
  optionsBack.addEventListener('click', onOptionsBackClick)
  optionsClose.addEventListener('click', onOptionsCloseClick)
  optionsBackdrop.addEventListener('click', onOptionsBackdropClick)

  // --- Anchor positioning: keep the FAB tucked under Copilot ------------
  //
  // OWA's top toolbar shifts position when the window resizes and when OWA
  // toggles its layout (compact / reading-pane / mobile breakpoint). A
  // hard-coded top:220px would drift out of place. So we look up the Copilot
  // button on every position-relevant event and place the FAB ~8px below it
  // (centered horizontally on the same column). Falls back to the CSS
  // defaults (top:220 right:24) if Copilot isn't on the page.
  //
  // The panel itself is a CSS-centered modal — its size is set by `calc()`
  // against viewport units, so viewport resize doesn't need any JS work.
  const FAB_SIZE = 44
  const GAP_BELOW_COPILOT = 8

  function positionFab() {
    const anchor = findCopilotAnchor()
    let fabTop: number
    let fabLeft: number
    if (anchor) {
      const r = anchor.getBoundingClientRect()
      fabTop = r.bottom + GAP_BELOW_COPILOT
      fabLeft = r.left + (r.width - FAB_SIZE) / 2
    } else {
      // Fallback when Copilot isn't on the page: pin to viewport top-right.
      fabTop = 220
      fabLeft = window.innerWidth - 24 - FAB_SIZE
    }

    fab.style.top = `${fabTop}px`
    fab.style.left = `${fabLeft}px`
    fab.style.right = 'auto'
    fab.style.bottom = 'auto'
  }

  // Reposition triggers:
  //   1. window.resize — viewport size changed
  //   2. ResizeObserver on the anchor — Copilot button moved or resized
  //      (covers OWA toolbar layout shifts that don't fire window.resize)
  //   3. MutationObserver on document.body subtree — used only while
  //      anchor is missing (OWA SPA still loading); detaches itself the
  //      moment we find the anchor.
  // No 2-second polling backstop: those wakeups were burning CPU for 8h
  // a day for nothing once anchor was stable.
  let anchorRO: ResizeObserver | null = null
  let bodyMO: MutationObserver | null = null
  let lastAnchor: HTMLElement | null = null

  const stopBodyMO = () => {
    if (bodyMO) {
      bodyMO.disconnect()
      bodyMO = null
    }
  }
  const startBodyMO = () => {
    if (bodyMO) return
    bodyMO = new MutationObserver(() => {
      const anchor = findCopilotAnchor()
      if (anchor) {
        stopBodyMO()
        positionFab()
      }
    })
    bodyMO.observe(document.body, { childList: true, subtree: true })
  }

  const positionFabAndObserve = () => {
    positionFab()
    const anchor = findCopilotAnchor()
    if (anchor && anchor !== lastAnchor) {
      lastAnchor = anchor
      if (anchorRO) anchorRO.disconnect()
      anchorRO = new ResizeObserver(() => positionFab())
      anchorRO.observe(anchor)
      stopBodyMO()
    } else if (!anchor) {
      lastAnchor = null
      startBodyMO()
    }
  }

  requestAnimationFrame(positionFabAndObserve)
  const onResize = () => positionFab()
  window.addEventListener('resize', onResize)

  teardown = () => {
    fab.removeEventListener('click', onFabClick)
    panelBackdrop.removeEventListener('click', onPanelBackdropClick)
    document.removeEventListener('keydown', onDocKey, true)
    window.removeEventListener('resize', onResize)
    window.removeEventListener('message', onPostMessage)
    optionsBack.removeEventListener('click', onOptionsBackClick)
    optionsClose.removeEventListener('click', onOptionsCloseClick)
    optionsBackdrop.removeEventListener('click', onOptionsBackdropClick)
    if (anchorRO) anchorRO.disconnect()
    stopBodyMO()
    if (panelUnloadTimer !== null) window.clearTimeout(panelUnloadTimer)
    if (optionsUnloadTimer !== null) window.clearTimeout(optionsUnloadTimer)
    host.remove()
    teardown = null
  }
}

function removeFab() {
  if (teardown) teardown()
  document.getElementById(HOST_ID)?.remove()
}

async function init() {
  console.log(LOG_PREFIX, 'init on', location.host, 'readyState=', document.readyState)
  if (await shouldShow()) createFab()
}

// Hot-reload on settings change so the toggle in options takes effect
// without requiring a page reload.
chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local' || !changes.settings) return
  const next = changes.settings.newValue as { showOwaFab?: boolean } | undefined
  const show = next?.showOwaFab !== false
  if (show && !document.getElementById(HOST_ID)) createFab()
  if (!show) removeFab()
})

console.log(LOG_PREFIX, 'script loaded')
void init()
