import { defineConfig } from "vitest/config";

// Date handling in the fingerprint uses local time (GAS runs in the spreadsheet time
// zone); pin it so tests are deterministic.
process.env.TZ = "America/Chicago";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      reporter: ["text", "html"],
    },
  },
});
