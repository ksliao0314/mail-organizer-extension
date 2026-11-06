import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import { ErrorBoundary } from '@/components/ui/error-boundary'
import { BridgeGate } from '@/components/ui/bridge-gate'
import { initParentBridge } from '@/shared/parent-bridge'
import { discardSavedPlan } from './crash-recovery'
// Variable fonts — self-hosted via @fontsource-variable (no Google Fonts
// CDN call, avoids MV3 CSP issues + saves a network round-trip per
// popup open). Geist = body text (Linear's typographic language;
// distinct from generic system / Inter). JetBrains Mono = numbers /
// signal / folder paths (tabular figures, programmer-grade legibility).
// Latin subsets only (2026-11-05) — see src/fonts.css.
import '@/fonts.css'
import '@/index.css'

// Listen for the FAB content script's handshake BEFORE React mounts
// (2026-11-05): the hello is posted on the iframe's load event, and React's
// first passive effects can run after that — a listener installed from a
// useEffect could miss it. Idempotent; a no-op in the toolbar popup.
initParentBridge()

// onDiscardState (2026-11-06): a reload restores the saved plan, so a crash
// caused by that plan's data would repeat — give the user a way to drop it.
// BridgeGate (2026-11-06): framed in Outlook, the UI renders only once the
// FAB content script has paired — see components/ui/bridge-gate.tsx.
createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary onDiscardState={discardSavedPlan}>
      <BridgeGate>
        <App />
      </BridgeGate>
    </ErrorBoundary>
  </StrictMode>,
)
