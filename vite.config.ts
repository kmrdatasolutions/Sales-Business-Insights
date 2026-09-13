import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],

  base: "/Sales-Business-Insights/",

  server: {
    port: 5173,
    host: "localhost"
  }
});