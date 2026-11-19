import { defineConfig, devices } from "@playwright/test";

// Ports differ from the interactive demo (8787 / 5173) so both can run at once.
const BACKEND_PORT = 8788;
const WEB_PORT = 5174;
const isCI = Boolean(process.env.CI);

export default defineConfig({
  testDir: "./e2e",
  // Tests share one in-memory backend and reset it before each test
  fullyParallel: false,
  workers: 1,
  retries: isCI ? 1 : 0,
  reporter: isCI
    ? [["github"], ["html", { open: "never" }]]
    : [["list"], ["html", { open: "never" }]],
  use: {
    ...devices["Desktop Chrome"],
    viewport: { width: 390, height: 844 },
    baseURL: `http://127.0.0.1:${WEB_PORT}`,
    trace: "retain-on-failure",
  },
  projects: [{ name: "chromium", use: { browserName: "chromium" } }],
  webServer: [
    {
      // Real doPost/doGet on in-memory GAS mocks with fake data
      command: "pnpm --dir ../budget-scripts-gas run dev:server",
      env: { DEMO_PORT: String(BACKEND_PORT) },
      url: `http://127.0.0.1:${BACKEND_PORT}/__debug/health`,
      reuseExistingServer: !isCI,
      timeout: 60_000,
    },
    {
      command: `pnpm exec vite --host 127.0.0.1 --port ${WEB_PORT} --strictPort`,
      url: `http://127.0.0.1:${WEB_PORT}/`,
      reuseExistingServer: !isCI,
      timeout: 60_000,
    },
  ],
});
