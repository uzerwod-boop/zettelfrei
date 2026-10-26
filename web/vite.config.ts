import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
// base "./": läuft unter /app/ auf GitHub Pages, lokal und später in der Desktop-App (Tauri)
export default defineConfig({ plugins: [react()], base: "./", build: { outDir: "dist", sourcemap: false } });
