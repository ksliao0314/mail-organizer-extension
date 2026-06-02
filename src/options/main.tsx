import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
// Same font setup as popup (see popup/main.tsx for why).
import '@fontsource-variable/geist'
import '@fontsource-variable/jetbrains-mono'
import '@/index.css'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
