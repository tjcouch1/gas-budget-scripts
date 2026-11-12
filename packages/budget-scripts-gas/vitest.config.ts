import { defineConfig } from "vitest/config";

// Sheet-name/date formatting in src/ uses local time; pin it so tests are deterministic.
process.env.TZ = "America/Chicago";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    setupFiles: ["./test/setup.ts"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      reporter: ["text", "html"],
    },
  },
});
