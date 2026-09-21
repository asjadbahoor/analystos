import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // In `npm run dev`, forward AI calls to the Flask backend on :5000
      "/api": "http://localhost:5000",
    },
  },
  build: {
    outDir: "dist",
  },
});
