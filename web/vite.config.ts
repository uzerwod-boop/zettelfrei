import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";
const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8"));
// base "./": läuft auf GitHub Pages, lokal und in der Desktop-App (Tauri)
export default defineConfig({ plugins: [react()], base: "./", build: { outDir: "dist", sourcemap: false }, define: { __APP_VERSION__: JSON.stringify(pkg.version) } });
