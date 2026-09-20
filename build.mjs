import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as path from "node:path";

const distDir = path.resolve("dist");
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

// Copy appsscript.json to dist/
fs.copyFileSync("appsscript.json", path.join(distDir, "appsscript.json"));

const exposedFunctions = [
  "onOpen",
  "onEditInstalled",
  "getAndRecordSomeReceiptsNoMark",
  "getAndRecordReceiptsNoMark",
  "getAndRecordSomeReceiptsAndMark",
  "getAndRecordReceiptsAndMark",
  "addTransactionSheet",
  "addTransactionSheets",
  "sheetNameOffset",
  "getMonthOffset",
  "logTheadById",
  "logReceipts",
  "logVariables",
  "testLogTransactionSheetInfos",
  "testRangeCachePerformance",
  "testSplit",
];

const footerJs = `
// Expose functions to globalThis
Object.assign(globalThis, _entry);

// Explicit top-level function declarations for Google Apps Script triggers, custom formulas, and the editor
${exposedFunctions
  .map(
    (fn) =>
      `function ${fn}(...args) {\n  return _entry.${fn}(...args);\n}`
  )
  .join("\n\n")}
`;

await esbuild.build({
  entryPoints: ["src/index.ts"],
  bundle: true,
  outfile: "dist/index.js",
  format: "iife",
  globalName: "_entry",
  target: "es2022",
  minify: false,
  footer: {
    js: footerJs,
  },
});

console.log("Build completed: dist/index.js and dist/appsscript.json generated.");
