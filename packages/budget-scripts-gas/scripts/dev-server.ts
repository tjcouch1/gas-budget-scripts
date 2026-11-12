/**
 * @file scripts/dev-server.ts
 *
 * Bundles `dev/main.ts` (the local demo backend) with esbuild and runs it in this process. Node
 * cannot run the TypeScript sources directly because they use extensionless imports.
 *
 * Output goes to `.demo/` (gitignored), never to `dist/` (which clasp pushes).
 */
import * as path from "node:path";
import { pathToFileURL } from "node:url";
import * as esbuild from "esbuild";

// Matches appsscript.json: sheet names and dates in src/ use local time
process.env.TZ = "America/Chicago";

const outfile = path.resolve(".demo/server.mjs");

await esbuild.build({
  entryPoints: [path.resolve("dev/main.ts")],
  outfile,
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node24",
  logLevel: "warning",
});

await import(pathToFileURL(outfile).href);
