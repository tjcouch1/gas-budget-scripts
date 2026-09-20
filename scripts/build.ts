import { spawn } from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import * as esbuild from "esbuild";

const distDir = path.resolve("dist");
const entryFile = path.resolve("src/index.ts");
const manifestFile = path.resolve("appsscript.json");

function ensureDistDir() {
  if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir, { recursive: true });
  }
}

function copyManifest() {
  ensureDistDir();
  fs.copyFileSync(manifestFile, path.join(distDir, "appsscript.json"));
}

/**
 * Dynamically discovers all exported function names from src/index.ts.
 */
function getExportedFunctionNames(filePath: string): string[] {
  const content = fs.readFileSync(filePath, "utf-8");
  const functions: string[] = [];

  // Match: export function name(
  const fnRegex = /^export\s+function\s+([a-zA-Z0-9_$]+)\s*\(/gm;
  let match: RegExpExecArray | null;
  while ((match = fnRegex.exec(content)) !== null) {
    functions.push(match[1]);
  }

  // Match: export { name1, name2 }
  const namedExportRegex = /^export\s*\{([^}]+)\}/gm;
  while ((match = namedExportRegex.exec(content)) !== null) {
    const names = match[1]
      .split(",")
      .map((n) => n.trim().split(/\s+as\s+/).pop()!.trim())
      .filter(Boolean);
    functions.push(...names);
  }

  return functions;
}

function createFooter(exposedFunctions: string[]): string {
  return `
// Expose functions to globalThis
Object.assign(globalThis, _entry);

// Explicit top-level function declarations for Google Apps Script triggers, custom formulas, and the editor
${exposedFunctions
  .map(
    (fn) => `function ${fn}(...args) {\n  return _entry.${fn}(...args);\n}`
  )
  .join("\n\n")}
`;
}

async function run() {
  const isWatch = process.argv.includes("--watch");
  const isClasp = process.argv.includes("--clasp");

  ensureDistDir();
  copyManifest();

  const exposedFunctions = getExportedFunctionNames(entryFile);
  console.log(
    `Discovered ${exposedFunctions.length} exported functions from src/index.ts:\n` +
      exposedFunctions.map((f) => `  - ${f}`).join("\n")
  );

  const buildOptions: esbuild.BuildOptions = {
    entryPoints: [entryFile],
    bundle: true,
    outfile: path.join(distDir, "index.js"),
    format: "iife",
    globalName: "_entry",
    target: "es2022",
    minify: false,
    footer: {
      js: createFooter(exposedFunctions),
    },
    plugins: [
      {
        name: "gas-rebuild-notifier",
        setup(build) {
          build.onEnd((result) => {
            if (result.errors.length === 0) {
              copyManifest();
              console.log(
                `[${new Date().toLocaleTimeString()}] Successfully built dist/index.js and dist/appsscript.json.`
              );
            }
          });
        },
      },
    ],
  };

  if (isWatch) {
    const ctx = await esbuild.context(buildOptions);
    await ctx.watch();
    console.log("esbuild is watching for file changes in src/...");

    if (isClasp) {
      console.log("Starting clasp push --watch on dist/...");
      const claspProcess = spawn("pnpm", ["exec", "clasp", "push", "--watch"], {
        stdio: "inherit",
        shell: true,
      });

      const cleanup = () => {
        ctx.dispose();
        claspProcess.kill();
        process.exit(0);
      };

      process.on("SIGINT", cleanup);
      process.on("SIGTERM", cleanup);
    }
  } else {
    await esbuild.build(buildOptions);
    console.log("Build completed successfully.");
  }
}

run().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
