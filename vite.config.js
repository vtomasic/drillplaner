import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    host: true, // im LAN erreichbar (z.B. iPad im gleichen WLAN)
    port: Number(process.env.PORT) || 5173,
    strictPort: false,
  },
})
