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
