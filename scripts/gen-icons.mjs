// Render public/icons/icon.svg into 16/48/128 PNGs for the extension manifest.
//
// Run via `npm run icons` or implicitly by `npm run build`. Sharp is pure JS
// wrapper over libvips with platform prebuilds; install is non-interactive.

import sharp from 'sharp'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const root = resolve(__dirname, '..')
const svgPath = resolve(root, 'public/icons/icon.svg')
const outDir = resolve(root, 'public/icons')

const svg = readFileSync(svgPath)

for (const size of [16, 32, 48, 128]) {
  const file = resolve(outDir, `icon-${size}.png`)
  await sharp(svg, { density: 320 })
    .resize(size, size, { fit: 'contain', kernel: 'lanczos3' })
    .png({ compressionLevel: 9 })
    .toFile(file)
  console.log(`✓ ${file}`)
}
