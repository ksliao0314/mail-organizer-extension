import { defineConfig } from 'vite'
import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { crx } from '@crxjs/vite-plugin'
import manifest from './manifest.config'

export default defineConfig({
  plugins: [react(), tailwindcss(), crx({ manifest })],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
    },
  },
  build: {
    target: 'esnext',
    // Sourcemaps were 2.6 MB of the 3.5 MB dist and the built manifest exposed the
    // content scripts' maps to OWA pages. Opt in for debugging: SOURCEMAP=1 npm run build
    sourcemap: process.env.SOURCEMAP === '1',
  },
})
