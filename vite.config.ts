import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],

  base: "/sales-business-insights/",

  server: {
    port: 5173,
    host: "localhost"
  }
});