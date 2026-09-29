import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { env } from 'node:process'

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/api': { target: env.AYA_API_PROXY_TARGET || 'http://127.0.0.1:8001', changeOrigin: true },
    },
  },
})
