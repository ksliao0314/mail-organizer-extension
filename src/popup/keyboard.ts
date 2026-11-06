// Keyboard-target classification for the document-level shortcut handlers
// in the popup + options pages (2026-11-05).
//
// Document `keydown` listeners see every key press, including ones meant for
// whatever element has focus. The old guards only exempted INPUT / TEXTAREA,
// so e.g. Enter on a focused 執行 button toggled the keyboard-cursor row
// instead of executing, and Esc on a dropdown / overlay tore down the whole
// panel. These helpers let each handler decide "is this key someone else's?".

function asElement(t: EventTarget | null): Element | null {
  return t && typeof (t as Element).closest === 'function' ? (t as Element) : null
}

/** Typing surfaces: every key (letters, Esc, arrows) belongs to them. */
export function isTextEntryTarget(t: EventTarget | null): boolean {
  const el = asElement(t)
  if (!el) return false
  if ((el as HTMLElement).isContentEditable) return true
  const tag = el.tagName
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT'
}

/**
 * Widgets that consume arrow / letter keys natively (or via their own
 * handlers): selects, contenteditable, ARIA composite widgets.
 */
export function ownsNavigationKeys(t: EventTarget | null): boolean {
  const el = asElement(t)
  if (!el) return false
  if (isTextEntryTarget(el)) return true
  return !!el.closest(
    '[contenteditable=""], [contenteditable="true"], [role="menuitem"], [role="option"], [role="listbox"], [role="combobox"], [role="slider"], [role="spinbutton"], [role="textbox"]',
  )
}

/**
 * Elements whose native Enter / Space activation must win over a
 * document-level shortcut (buttons, links, disclosure summaries, ARIA
 * button-likes). Navigation keys (j/k, arrows) still reach the shortcut.
 */
export function ownsActivationKeys(t: EventTarget | null): boolean {
  const el = asElement(t)
  if (!el) return false
  if (ownsNavigationKeys(el)) return true
  return !!el.closest(
    'button, a[href], summary, [role="button"], [role="link"], [role="checkbox"], [role="switch"], [role="tab"]',
  )
}

// ---- Popup plan list (2026-11-06) ------------------------------------------
//
// The plan screen's j/k cursor is VIRTUAL (`focusedIndex`) — it never moves
// DOM focus — while Chrome leaves DOM focus on whatever button the mouse last
// clicked, and nearly the whole plan row is a <button> (the expand toggle).
// So the focused element and the row with the cursor ring routinely
// disagree. v1.0.8's "a focused button owns Enter / Space" (added so Enter on
// 執行 executes) therefore sent those keys to the last-clicked row, filter
// chip or sort toggle: Space never selected, and click row 1 → j → Enter
// toggled row 1. Inside the list the cursor row wins again; controls outside
// it keep their native activation.

/** On the plan list container (`<ul>`). */
export const PLAN_LIST_ATTR = 'data-plan-list'
/** On each PlanRow root: its index in the rendered list (= focusedIndex space). */
export const PLAN_ROW_INDEX_ATTR = 'data-plan-row-index'
/** On a PlanRow's expanded detail panel (action buttons, 判斷依據, 收起…). */
export const PLAN_ROW_DETAIL_ATTR = 'data-plan-row-detail'

/**
 * A row's checkbox is list chrome, not a typing surface: clicking one leaves
 * it focused, and j/k/x/d/s must keep working from there (it used to swallow
 * every shortcut as an INPUT).
 */
export function isPlanListCheckbox(t: EventTarget | null): boolean {
  const el = asElement(t)
  return (
    !!el &&
    el.tagName === 'INPUT' &&
    (el as HTMLInputElement).type === 'checkbox' &&
    !!el.closest(`[${PLAN_LIST_ATTR}]`)
  )
}

/**
 * Should Enter / Space be left to the focused element's own activation on
 * the plan screen? True for:
 *   - a control OUTSIDE the plan list (執行, banner ×, bulk bar…) — v1.0.8;
 *   - any list control while there is no keyboard cursor;
 *   - a control in the CURSOR row's open detail panel — DOM focus and the
 *     cursor agree on the email. (Mouse clicks there don't move focus — see
 *     PlanRow — so this is a control the user Tabbed to.)
 *   - a list control that is NOT part of a row (全選此組, 「+ 還有 N 封同對話」,
 *     「↑ 收合」) when the user Tabbed to it (`keyboardFocused`, 2026-11-06):
 *     with a cursor present, these used to be unreachable from the keyboard.
 * False → the key acts on the cursor row (Space = 勾選, Enter = 展開/收起),
 * even with DOM focus on another row's toggle / checkbox or a list header
 * that the mouse left focused. Tabbing onto a row moves the cursor there
 * (PlanScreen), so for row controls "cursor row" = the row the user Tabbed to.
 */
export function focusOwnsPlanActivation(
  t: EventTarget | null,
  cursorIndex: number | null,
  keyboardFocused = false,
): boolean {
  const el = asElement(t)
  if (!el || !ownsActivationKeys(el)) return false
  if (!el.closest(`[${PLAN_LIST_ATTR}]`)) return true
  if (cursorIndex === null) return true
  const row = el.closest(`[${PLAN_ROW_INDEX_ATTR}]`)
  if (!row) return keyboardFocused
  return (
    !!el.closest(`[${PLAN_ROW_DETAIL_ATTR}]`) &&
    row.getAttribute(PLAN_ROW_INDEX_ATTR) === String(cursorIndex)
  )
}

/**
 * Where did the current DOM focus come from (2026-11-06)? The plan list needs
 * to tell a Tab from a mouse click: Tab-focus moves the keyboard cursor onto
 * that row and lets a non-row list control keep its own Enter / Space, while
 * focus a mouse click merely left behind must not steer anything (the E1
 * case). Only a Tab / Shift+Tab counts — a window regaining focus re-fires
 * focusin on the old element, which must not yank the cursor back.
 *
 * `release()` = the user moved on with the keyboard cursor (j / k): from then
 * on the Tabbed-to element no longer owns Enter / Space.
 */
export type FocusOriginTracker = {
  /** `t` got focus from Tab and nothing (click, cursor move) happened since. */
  isKeyboardFocused(t: EventTarget | null): boolean
  release(): void
  dispose(): void
}

export function createFocusOriginTracker(doc: Document = document): FocusOriginTracker {
  let lastWasTab = false
  let keyboardFocused: Element | null = null
  const onPointer = () => {
    lastWasTab = false
    keyboardFocused = null
  }
  const onKey = (e: KeyboardEvent) => {
    lastWasTab = e.key === 'Tab'
  }
  const onFocusIn = (e: FocusEvent) => {
    keyboardFocused = lastWasTab ? asElement(e.target) : null
  }
  // Capture phase on the document: runs before any handler can stop the
  // event, and (for focusin) before React's own onFocus on the list.
  doc.addEventListener('pointerdown', onPointer, true)
  doc.addEventListener('mousedown', onPointer, true)
  doc.addEventListener('keydown', onKey, true)
  doc.addEventListener('focusin', onFocusIn, true)
  return {
    isKeyboardFocused: (t) => t !== null && t === keyboardFocused,
    release: () => {
      keyboardFocused = null
    },
    dispose: () => {
      doc.removeEventListener('pointerdown', onPointer, true)
      doc.removeEventListener('mousedown', onPointer, true)
      doc.removeEventListener('keydown', onKey, true)
      doc.removeEventListener('focusin', onFocusIn, true)
    },
  }
}

/**
 * Buttons that shouldn't take DOM focus from a MOUSE click (2026-11-06):
 * plan-list toggles that reshape the list under the cursor (待確認, 依資料夾)
 * and the buttons in a row's detail panel. They stay Tab-focusable — this
 * only stops a mouse click from leaving focus parked on them, where the next
 * Space / Enter would re-activate them instead of reaching the cursor row.
 * Inputs / comboboxes / options are left alone (they need focus to work).
 */
export function preventsPointerFocus(t: EventTarget | null): boolean {
  const el = asElement(t)
  if (!el || ownsNavigationKeys(el)) return false
  return !!el.closest('button, summary, [role="button"]')
}

/**
 * An overlay layer (dialog, menu, open listbox) is showing — Esc belongs to
 * the innermost layer, never to a "close the whole surface" handler.
 */
export function overlayLayerOpen(root: ParentNode = document): boolean {
  return !!root.querySelector(
    '[role="dialog"], [role="alertdialog"], [role="menu"], [role="listbox"]',
  )
}

/** IME composition cancel arrives as keydown Esc with isComposing / 229. */
export function isImeComposing(e: KeyboardEvent): boolean {
  return e.isComposing || e.keyCode === 229
}
