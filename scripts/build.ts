/**
 * @file scripts/build.ts
 *
 * Build pipeline for gas-budget-scripts.
 *
 * Why this script exists:
 * Google Apps Script (GAS) executes all files in a single flat global context on Google's servers
 * and does not support ES module `import` / `export` syntax between files.
 *
 * This script uses esbuild to bundle our modular TypeScript codebase into a single `dist/index.js`
 * wrapped in an IIFE. Crucially:
 * 1. Output is kept UNMINIFIED so the code remains readable in the online editor at script.google.com.
 * 2. Exported functions from `src/index.ts` are dynamically discovered and declared as top-level
 *    `function <name>(...)` statements in the bundle footer so GAS's static parser can register triggers
 *    (e.g., `onOpen`, `onEditInstalled`), custom spreadsheet formulas (e.g., `=sheetNameOffset()`),
 *    and entries in the web editor's "Run" dropdown.
 * 3. The `appsscript.json` manifest is copied to `dist/` so clasp can push directly from `dist/`.
 * 4. Supports `--watch` for rebuild-on-save and `--clasp` for live-syncing to Google Apps Script.
 */

import { spawn, spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import * as esbuild from "esbuild";

/** Directory where clasp-ready build artifacts are emitted */
const distDir = path.resolve("dist");

/** Main entry file containing exported GAS functions */
const entryFile = path.resolve("src/index.ts");

/** Google Apps Script project manifest required in the root of the push directory */
const manifestFile = path.resolve("appsscript.json");

/**
 * Ensures the output distribution directory exists before writing artifacts.
 */
function ensureDistDir(): void {
  if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir, { recursive: true });
  }
}

/**
 * Copies the `appsscript.json` manifest to `dist/`.
 *
 * Clasp configured with `"rootDir": "./dist"` requires the manifest to be co-located
 * with the bundled JavaScript files when pushing to Google.
 */
function copyManifest(): void {
  ensureDistDir();
  fs.copyFileSync(manifestFile, path.join(distDir, "appsscript.json"));
}

/**
 * Dynamically discovers all exported function names from an entry file by scanning its source text.
 *
 * This avoids manually maintaining an array of exposed functions that could fall out of sync
 * whenever new functions or custom spreadsheet formulas are added to `src/index.ts`.
 *
 * @param filePath Absolute path to the source file to inspect (typically `src/index.ts`).
 * @returns An array of function identifiers that are exported by the file.
 */
function getExportedFunctionNames(filePath: string): string[] {
  const content = fs.readFileSync(filePath, "utf-8");
  const functions: string[] = [];

  // Match direct function exports: `export function myFunction(...)` / `export async function myFunction(...)`
  // Regex breakdown:
  //   ^               - Start of line (thanks to 'm' flag)
  //   export\s+       - The keyword 'export' followed by one or more spaces
  //   (?:async\s+)?   - Optional 'async' keyword followed by one or more spaces
  //   function\s+     - The keyword 'function' followed by one or more spaces
  //   ([a-zA-Z0-9_$]+)- Capture group 1: valid JavaScript identifier characters
  //   \s*\(           - Optional whitespace followed by opening parenthesis of argument list
  //   /gm             - Global matching across all lines
  const fnRegex =
    /^export\s+(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(/gm;
  let match: RegExpExecArray | null;
  while ((match = fnRegex.exec(content)) !== null) {
    functions.push(match[1]);
  }

  // Match grouped exports: `export { fn1, fn2 as alias }`
  // Regex breakdown:
  //   ^export\s*\{    - Line starting with 'export {'
  //   ([^}]+)         - Capture group 1: everything inside the braces
  //   \}              - Closing brace
  const namedExportRegex = /^export\s*\{([^}]+)\}/gm;
  while ((match = namedExportRegex.exec(content)) !== null) {
    const names = match[1]
      .split(",")
      .map((entry) => {
        // If an item is aliased (`fn as myAlias`), we want the exposed alias (`myAlias`)
        const parts = entry.trim().split(/\s+as\s+/);
        return parts[parts.length - 1].trim();
      })
      .filter(Boolean);
    functions.push(...names);
  }

  return functions;
}

/**
 * Generates JavaScript code to append to the end of the esbuild bundle.
 *
 * Why this is necessary for Google Apps Script:
 * When esbuild creates an IIFE bundle (`var _entry = (() => { ... })()`), the exported
 * functions exist as properties on `_entry`. However, Google Apps Script's static parser
 * parses the top-level AST of script files to discover:
 * 1. Simple triggers like `onOpen(e)` and installed triggers like `onEditInstalled(e)`.
 * 2. Custom spreadsheet formula functions (e.g., `=sheetNameOffset(-1, 0)`).
 * 3. Functions available in the script editor's "Run" / "Debug" dropdown.
 *
 * If a function is only attached to `globalThis` or nested inside `_entry`, GAS will not
 * detect it as a callable top-level function. This footer bridges that gap by generating explicit
 * top-level `function <name>(...args)` definitions that forward calls to `_entry`. These declarations
 * are also global properties, so name-based lookups (e.g., `Ui.addItem` menu entries) resolve them too.
 *
 * @param exposedFunctions List of function names to expose globally.
 * @returns JavaScript code to be appended to the bundle.
 */
function createFooter(exposedFunctions: string[]): string {
  return `
// Explicit top-level function declarations for Google Apps Script AST parser:
// allows triggers (onOpen), custom formulas (=sheetNameOffset), and the editor dropdown to find them.
${exposedFunctions
  .map(
    (fn) => `function ${fn}(...args) {\n  return _entry.${fn}(...args);\n}`
  )
  .join("\n\n")}
`;
}

/**
 * Main build and watch orchestrator.
 *
 * Flags:
 *   --watch : Keeps esbuild alive and rebuilds `dist/index.js` on every source change.
 *   --clasp : (Used with --watch by `pnpm start`) Spawns `clasp push --watch` to upload
 *             changes to Google Apps Script as soon as esbuild writes to `dist/`.
 */
async function run(): Promise<void> {
  const isWatch = process.argv.includes("--watch");
  const isClasp = process.argv.includes("--clasp");

  ensureDistDir();
  copyManifest();

  const outfile = path.join(distDir, "index.js");

  const buildOptions: esbuild.BuildOptions = {
    entryPoints: [entryFile],
    bundle: true,
    outfile,
    format: "iife",
    globalName: "_entry",
    target: "es2022",
    // Keep unminified so code is human-readable and inspectable on script.google.com
    minify: false,
    // Output is written in onEnd so the GAS footer can be regenerated from the current
    // exports of src/index.ts on every (re)build, including in --watch mode.
    write: false,
    plugins: [
      {
        name: "gas-footer-writer",
        setup(build) {
          build.onEnd((result) => {
            if (result.errors.length > 0) return;

            const exposedFunctions = getExportedFunctionNames(entryFile);
            console.log(
              `Discovered ${exposedFunctions.length} exported functions from src/index.ts:\n` +
                exposedFunctions.map((f) => `  - ${f}`).join("\n")
            );

            const bundle = result.outputFiles?.find(
              (file) => path.resolve(file.path) === outfile
            );
            if (!bundle) throw new Error(`esbuild did not produce ${outfile}`);

            const contents = bundle.text + createFooter(exposedFunctions);
            // Skip identical writes so clasp --watch does not push unchanged bundles
            const existing = fs.existsSync(outfile)
              ? fs.readFileSync(outfile, "utf-8")
              : undefined;
            if (existing !== contents) fs.writeFileSync(outfile, contents);

            // Ensure manifest stays copied if dist was cleaned or touched
            copyManifest();
            console.log(
              `[${new Date().toLocaleTimeString()}] Successfully built dist/index.js and dist/appsscript.json.`
            );
          });
        },
      },
    ],
  };

  if (isWatch) {
    // Persistent context allows fast incremental rebuilds as files in src/ change
    const ctx = await esbuild.context(buildOptions);
    await ctx.watch();
    console.log("esbuild is watching for file changes in src/...");

    if (isClasp) {
      console.log("Starting clasp push --watch on dist/...");
      // Spawn clasp watch to push changes from dist/ to script.google.com
      const claspProcess = spawn(
        "pnpm",
        ["exec", "clasp", "push", "--watch"],
        {
          stdio: "inherit",
          shell: true,
        }
      );

      // Gracefully terminate child clasp process when the user stops the script (Ctrl+C)
      const cleanup = async () => {
        await ctx.dispose();
        if (process.platform === "win32" && claspProcess.pid !== undefined) {
          // With `shell: true`, kill() only stops the cmd.exe wrapper on Windows; kill the whole tree
          spawnSync("taskkill", ["/pid", String(claspProcess.pid), "/T", "/F"], {
            stdio: "ignore",
          });
        } else {
          claspProcess.kill();
        }
        process.exit(0);
      };

      process.on("SIGINT", cleanup);
      process.on("SIGTERM", cleanup);
    }
  } else {
    // Single one-off production build
    await esbuild.build(buildOptions);
    console.log("Build completed successfully.");
  }
}

run().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
