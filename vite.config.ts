import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    // ExcelJS (~930 kB) is lazy-loaded on Generate, so its chunk size doesn't affect page load
    chunkSizeWarningLimit: 1000,
  },
})
