import { createContext, useContext } from 'react'

// The popup renders in two very different containers:
//   - Chrome toolbar popup: ~540px wide, height capped, vertical layout
//   - OWA FAB iframe: fills the reading-pane region (900-1400px), full height
//
// Children (PlanRow / PlanScreen) read `isWide` to switch layouts. Threshold
// at 760px is the standard tablet-landscape breakpoint and matches the width
// where the multi-column row layout starts feeling comfortable.
//
// Lives in its own module so both App (the provider) and downstream
// components (the consumers) can import without circular dependencies.

export const LayoutContext = createContext<{ isWide: boolean }>({ isWide: false })

export function useLayout(): { isWide: boolean } {
  return useContext(LayoutContext)
}
