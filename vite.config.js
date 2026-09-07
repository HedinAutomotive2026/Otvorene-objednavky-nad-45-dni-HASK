import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base so the build works both locally and when hosted
// under a GitHub Pages project path (https://<user>.github.io/<repo>/).
export default defineConfig({
  plugins: [react()],
  base: "./",
});
