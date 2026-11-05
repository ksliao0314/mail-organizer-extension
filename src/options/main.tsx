import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import { ErrorBoundary } from '@/components/ui/error-boundary'
import { initParentBridge } from '@/shared/parent-bridge'
// Same font setup as popup (see popup/main.tsx for why).
// Latin subsets only (2026-11-05) — see src/fonts.css.
import '@/fonts.css'
import '@/index.css'

// Pair with the FAB content script before React mounts — see popup/main.tsx
// (2026-11-05). No-op when the options page is a standalone tab.
initParentBridge()

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
