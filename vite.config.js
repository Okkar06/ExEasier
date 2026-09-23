import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { viteSingleFile } from 'vite-plugin-singlefile'

// `npm run build` inlines all JS/CSS into one HTML file that opens straight
// from disk (file://), so the app can be used fully offline.
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  publicDir: false,
})
