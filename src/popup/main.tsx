import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
// Variable fonts — self-hosted via @fontsource-variable (no Google Fonts
// CDN call, avoids MV3 CSP issues + saves a network round-trip per
// popup open). Geist = body text (Linear's typographic language;
// distinct from generic system / Inter). JetBrains Mono = numbers /
// signal / folder paths (tabular figures, programmer-grade legibility).
import '@fontsource-variable/geist'
import '@fontsource-variable/jetbrains-mono'
import '@/index.css'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
