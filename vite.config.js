import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [basicSsl()],
  build: {
    rollupOptions: { input: "taskpane.html" }
  },
  server: {
    port: 3000,
    strictPort: true,
    proxy: {
      "/api": { target: "http://localhost:3001", changeOrigin: true }
    }
  }
});
