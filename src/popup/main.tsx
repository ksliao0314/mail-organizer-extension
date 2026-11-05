import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import { ErrorBoundary } from '@/components/ui/error-boundary'
import { initParentBridge } from '@/shared/parent-bridge'
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

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
