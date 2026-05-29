import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  // Use relative asset URLs so the app still loads when hosted from a subfolder.
  base: "./",
  plugins: [react()],
});
