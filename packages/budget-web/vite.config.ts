import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { VitePWA } from "vite-plugin-pwa";

// GitHub Pages serves the site under /<repo-name>/. The deploy workflow passes
// PAGES_BASE; the default matches the repository name (web-plan.md section 4.5). The dev server
// (used by the local demo and the e2e tests) serves from the root.
export default defineConfig(({ command }) => ({
  base:
    process.env.PAGES_BASE ??
    (command === "serve" ? "/" : "/gas-budget-scripts/"),
  plugins: [
    react(),
    tailwindcss(),
    VitePWA({
      registerType: "prompt",
      includeAssets: ["icon.svg"],
      manifest: {
        name: "Fast Budget",
        short_name: "Budget",
        description:
          "Fast categorizing and splitting for the budget spreadsheet",
        start_url: ".",
        scope: ".",
        display: "standalone",
        background_color: "#f8fafc",
        theme_color: "#0f766e",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          {
            src: "icon-512.png",
            sizes: "512x512",
            type: "image/png",
            purpose: "maskable",
          },
        ],
      },
      workbox: {
        // Precache the app shell only. Never cache Apps Script responses.
        globPatterns: ["**/*.{js,css,html,svg,png,webmanifest}"],
        globIgnores: ["licenses.html"],
        navigateFallbackDenylist: [/licenses\.html$/],
        runtimeCaching: [],
      },
    }),
  ],
}));
