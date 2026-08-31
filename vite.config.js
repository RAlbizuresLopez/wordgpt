import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async () => ({
  build: {
    rollupOptions: { input: "taskpane.html" }
  },
  server: {
    port: 3000,
    strictPort: true,
    https: await getHttpsServerOptions(),
    proxy: {
      "/api": { target: "http://localhost:3001", changeOrigin: true }
    }
  }
}));
