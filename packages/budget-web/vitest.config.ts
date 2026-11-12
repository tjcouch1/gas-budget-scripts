import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

process.env.TZ = "America/Chicago";

export default defineConfig({
  plugins: [react()],
  test: {
    environment: "jsdom",
    include: ["src/**/*.test.{ts,tsx}"],
    setupFiles: ["./src/test/setup.ts"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.{ts,tsx}"],
      exclude: ["src/test/**", "src/main.tsx", "src/**/*.test.*"],
      reporter: ["text", "html"],
    },
  },
});
