/** Entry point of the demo backend (bundled and run by `scripts/dev-server.ts`) */
import { DemoBackend, DEMO_API_KEY } from "./backend.ts";
import { createDemoServer, EXEC_PATH } from "./server.ts";

const port = Number(process.env.DEMO_PORT ?? 8787);
const backend = new DemoBackend();
const server = createDemoServer(backend);

server.listen(port, "127.0.0.1", () => {
  console.log("");
  console.log("Demo backend (real doPost/doGet on in-memory GAS mocks)");
  console.log(`  Web app URL: http://127.0.0.1:${port}${EXEC_PATH}`);
  console.log(`  API key:     ${DEMO_API_KEY}`);
  console.log("  Web app (Vite): http://127.0.0.1:5173/");
  console.log("  Paste both into the web app's Settings dialog.");
  console.log("  State is in memory; restart (or POST /__debug/reset) resets.");
  console.log("");
});

function shutdown() {
  server.close(() => process.exit(0));
  setTimeout(() => process.exit(0), 500).unref();
}
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
