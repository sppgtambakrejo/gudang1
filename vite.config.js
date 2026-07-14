import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Base path is injected automatically for GitHub Pages by the deploy workflow
// (set via VITE_BASE_PATH env var to "/<repo-name>/"). Locally it defaults to "/".
export default defineConfig({
  plugins: [react()],
  base: process.env.VITE_BASE_PATH || "/",
});
