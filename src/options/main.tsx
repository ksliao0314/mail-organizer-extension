import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import { ErrorBoundary } from '@/components/ui/error-boundary'
import { BridgeGate } from '@/components/ui/bridge-gate'
import { initParentBridge } from '@/shared/parent-bridge'
// Same font setup as popup (see popup/main.tsx for why).
// Latin subsets only (2026-11-05) — see src/fonts.css.
import '@/fonts.css'
import '@/index.css'

// Pair with the FAB content script before React mounts — see popup/main.tsx
// (2026-11-05). No-op when the options page is a standalone tab.
initParentBridge()

// Framed in Outlook's settings modal, the options UI renders only once the
// FAB content script has paired (2026-11-06) — options/index.html has to stay
// web-accessible for that modal, so a page script could embed it too. See
// components/ui/bridge-gate.tsx. No-op as a standalone tab.
createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <BridgeGate>
        <App />
      </BridgeGate>
    </ErrorBoundary>
  </StrictMode>,
)
