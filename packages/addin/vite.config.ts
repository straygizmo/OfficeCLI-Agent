import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// Reuse the certificate created by `office-addin-dev-certs install` so the
// dev server is trusted by Office (same cert the host process uses).
function devCerts() {
  const dir = path.join(os.homedir(), ".office-addin-dev-certs");
  const cert = path.join(dir, "localhost.crt");
  const key = path.join(dir, "localhost.key");
  if (fs.existsSync(cert) && fs.existsSync(key)) {
    return { cert: fs.readFileSync(cert), key: fs.readFileSync(key) };
  }
  return undefined;
}

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    strictPort: true,
    https: devCerts(),
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: true,
    rollupOptions: {
      input: { taskpane: path.resolve(__dirname, "taskpane.html") },
    },
  },
});
