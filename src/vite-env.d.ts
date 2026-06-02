/// <reference types="vite/client" />

// @fontsource-variable/* packages ship CSS-only — no TS declarations.
// Vite handles them fine at runtime, but tsc errors without these stubs.
declare module '@fontsource-variable/geist'
declare module '@fontsource-variable/jetbrains-mono'
