# AGENTS.md — Coding, Security & Architecture Guidelines

All AI coding agents and contributors MUST adhere strictly to the rules below.

## 0. Current Repository Layout (read first)

**Today** this repository is a **single-package** `pnpm` project: a container-bound Google Apps Script (GAS) project written in TypeScript.

```text
src/index.ts                  # GAS entry point: triggers, menu actions, custom sheet formulas (all exported)
src/util/budgeting.util.ts    # Receipt import, transaction sheet creation, splitTransaction
src/util/variables.util.ts    # Variables!A1:B10 + per-sheet SheetVariables readers (memoized)
src/util/util.ts, spreadsheet-util.util.ts
src/models/                   # receipt-info, thread-info, thread-list (Gmail parsing)
scripts/build.ts              # esbuild bundler -> dist/index.js (+ top-level GAS function footer)
test/*.test.ts                # Vitest characterization tests
test/gas-mocks/               # In-memory SpreadsheetApp / GmailApp / Logger mocks
.github/workflows/ci.yml      # CI: format:check, lint, typecheck, test:coverage, build
appsscript.json               # GAS manifest (copied to dist/ by the build)
web-plan.md                   # Plan for the PWA frontend + GAS JSON API (not yet implemented)
```

**Planned (not yet present):** `web-plan.md` Step 1 converts the repo into a `pnpm` workspace with `packages/budget-scripts-gas` (the code above), `packages/budget-shared` (`@budget/shared`) and `packages/budget-web` (React/Vite PWA). Rules below that name `packages/...` apply once those packages exist; until then, apply them to the equivalent root paths. When the migration lands, update section 0 of this file.

---

## 1. CRITICAL SECURITY RULES (ZERO TOLERANCE)

1. **NEVER Hardcode Secrets, URLs, or Personal Identifiers:**
   - **NEVER** commit, hardcode, or place real Google Apps Script Web App Deployment URLs (`https://script.google.com/macros/s/<id>/exec`), Script IDs, API keys, tokens, Spreadsheet IDs, or personal email addresses in any source file, test file, configuration file, `.env` file committed to git, or documentation. Use placeholders such as `SCRIPT_PROJECT_ID`, `<deployment-id>`, and `@example.com` addresses in tests (see `test/gas-mocks/index.ts`). Public vendor sender addresses already in `src/` (Chase/PayPal/Venmo) are not personal identifiers.
   - In the GAS code, all secrets (such as `API_KEY`) MUST be read at runtime from `PropertiesService.getScriptProperties()`.
   - In the web app, the GAS Deployment URL and `API_KEY` MUST be provided by the user at runtime via the UI Settings dialog.
   - `.clasp*.json` files are gitignored (they contain the Script ID). Never commit them.
2. **NEVER Store Plaintext Credentials in `localStorage`:**
   - In the web app, connection credentials (`gasWebAppUrl` and `apiKey`) MUST be encrypted using the Web Crypto API (`crypto.subtle`) with a non-extractable (`extractable: false`) `AES-GCM` 256-bit `CryptoKey` stored in `IndexedDB`.
   - Only encrypted ciphertext (`iv` + `ciphertext`) may be persisted. Decrypt only in-memory, inside the request execution scope.
   - This protects against copying stored data off the device; it does NOT protect against script running in the page (XSS). Never add third-party scripts loaded at runtime, `eval`, `new Function`, or `dangerouslySetInnerHTML` with sheet data.
3. **Never authenticate via URL query strings.** The API key travels only in the JSON body of `POST` requests (GAS `doPost` cannot read request headers).
4. **Never write unvalidated user text that GAS will interpret as a formula.** `Range.setValues()` treats any string beginning with `=` as a formula. Cost formulas must be produced only by the shared formula builder from an expression validated by the shared math parser. Free-text fields (e.g. `Name`) beginning with `=`, `+`, `-`, or `@` must be rejected or prefixed with `'` server-side.

---

## 2. AGENT SAFETY: THINGS YOU MUST NOT RUN

The scripts below change the user's **live** Google Apps Script project and spreadsheet. Do not run them unless the user explicitly asks you to in this session:

- `pnpm run push`, `pnpm start` (runs `clasp push --watch`), `pnpm run login`, `pnpm run connect`, `pnpm exec clasp ...` (`push`, `deploy`, `undeploy`, `run`, ...).
- Anything that calls a deployed `/exec` URL.

Never edit `dist/` by hand (it is generated and gitignored). Do not commit unless asked.

Branches: `main` is development and is used with a **test** budget sheet. `deploy-budget-sheet` is production and is used with the real sheet. Do not merge to or push `deploy-budget-sheet` unless asked.

---

## 3. DATA INTEGRITY & SPREADSHEET COMPATIBILITY RULES

1. **Preserve existing spreadsheet behavior.** The tests in `test/` are characterization tests that pin current cell values, formulas, labels, and error messages. A change that makes a test fail is a behavior change. Only update the test when the behavior change is intended and stated in the task or plan, and call out each updated expectation in your report. Never weaken an assertion just to make it pass.
2. **Prevent Silent Data Overwrites (web app):**
   - Never overwrite a sheet row from the web app without verifying, **on the server inside `LockService.getScriptLock()`**, that the row still matches the client's `baseFingerprint` (the row's state at the client's last pull). The fingerprint algorithm lives in the shared package and must be identical on client and server.
   - If the row changed, reject the write and return the current row. The client surfaces a 3-way conflict (`Original Value`, `Discarded Local Change`, `Current Remote Value`) and requires an explicit choice (`Accept Remote` or `Overwrite with Mine`).
3. **Spreadsheet formula and split conventions.** These match `splitTransaction` in `src/util/budgeting.util.ts`, which is pinned by `test/budgeting.test.ts`:
   - **Parent cost cell:** a plain number cost `C` becomes `=C-R[n]C[0]`, where `n` is the new child's row offset from the parent. An existing formula cost keeps its text, and `-R[n]C[0]` is appended. A web-initiated N-way split of an unsplit row writes `=C-R[1]C[0]-R[2]C[0]...-R[N]C[0]`.
   - Note: real Sheets returns A1 notation from `getFormulas()` (e.g. `=45-C12`). Use `getFormulasR1C1()` when you need relative references. The test mocks store formula text verbatim and do not convert notation.
   - **Child cost cell (Pre-Tax ON):** `=${TaxMultiplier}*(${expression})`. The native split writes `=${TaxMultiplier}*(0)`.
   - **Child cost cell (Pre-Tax OFF):** `=${expression}` with no outer parentheses (e.g. `=1+3.23*2`), or a raw number when the input is a single numeric literal.
   - **Split names:** the parent name gets `" | "` appended only if it does not already contain `"|"`. A child name is `"<parent name before ' | '> | <item note>"`. `doTransactionRowInfosSeemEqual` groups rows by equal date AND (equal prefix before `"|"` OR equal first third of the name). Keep this so native split/grouping keeps working.
   - **Native split columns:** a new child row copies `Paid with` (type) and leaves `Envelope` (category) blank. Shifting rows down moves only Date/Name/Cost (cols 0–2) and Envelope/Paid with (cols 4–5), not col 3 or col 6. Do not change this without a stated decision (see `web-plan.md` open questions).
4. **Gmail processing (current behavior):** `recordReceipts` labels each successful thread `Receipts/Scripted` + `Receipts` and archives it **one thread at a time**. Errored threads are written to the sheet (red name cell) but NOT labeled, so they are re-scanned and re-written on later runs. Non-receipt-only threads are dropped without labels. The planned batch labeling with `Receipts/Ignored` / `Receipts/Error` (see `web-plan.md` §4.2) is a deliberate behavior change. Implement it only as specified there, and update the pinned tests accordingly.

---

## 4. BUILD, TOOLING & GAS CONVENTIONS

1. **Tooling & Package Manager:**
   - Always use `pnpm` (Volta-pinned: Node `24.16.0`, `pnpm 10.34.5`). Never use `npm` or `yarn`.
   - TypeScript: `typescript` is aliased to `@typescript/typescript6` (typescript-eslint needs the TS 6 API), while `@typescript/native` (TypeScript 7) provides `tsc` for `typecheck`. Keep both aliases. New packages must use the root-installed versions; do not add a separate `typescript` version.
   - ESLint (flat config `eslint.config.js`, typescript-eslint + `eslint-config-prettier`) and Prettier (`.prettierrc.json`) are enforced in pre-commit and CI.
   - Once the shared package exists, keep shared DTOs, the fingerprint function, the inline math parser, and the formula builder in `@budget/shared`. Never duplicate them in frontend or backend.
2. **GAS build pipeline (`scripts/build.ts`):**
   - The entry point is `src/index.ts`. Any function that GAS must call (triggers, menu actions, custom formulas, `doGet`/`doPost`, editor-run helpers) MUST be exported from `src/index.ts` as `export function name(...)` or `export { name } from "..."`. The build discovers exports by regex and emits top-level `function name(...args)` wrappers.
   - Menu actions are dispatched by name via the `menuActions` map in `src/index.ts`. Add new menu items to both `menu` and `menuActions`.
   - Keep `minify: false` so deployed code stays readable in the Apps Script editor.
   - In code reachable from `doGet`/`doPost` or time-driven triggers, use `SpreadsheetApp.getActiveSpreadsheet()`, never `getActiveSheet()`/`getActiveRange()`/`setActiveRange()`, and never call `SpreadsheetApp.getUi()` without a try/catch. `getUi()` throws outside an interactive editor session. `variables.util.ts` already uses `getActiveSpreadsheet()`.
   - `src/` keeps module-level caches (`Variables`, Gmail labels). Each GAS execution starts fresh, so do not rely on caches persisting across requests.
3. **GAS & Gmail performance:**
   - In **new or modified** code, never call `getValues()`/`getFormulas()`/`setValues()`/`setNote()`/`setBackground()` inside a row-by-row loop. Read and write contiguous ranges in one batch (`getValues`, `getFormulas`, `getNotes`, `getBackgrounds`, `setValues`, `setNotes`, `setBackgrounds`). Existing code (`splitTransaction`'s row scan, `recordReceiptsOnSheet`'s per-row note/background writes) still does per-row I/O. Do not refactor it opportunistically. Change it only with tests that prove identical output.
   - Gmail batch APIs (`GmailLabel.addToThreads`, `GmailApp.moveThreadsToArchive`) accept at most 100 threads per call. Chunk larger arrays.
4. **Web app → GAS requests:** send `POST` with `Content-Type: text/plain;charset=utf-8`, body `JSON.stringify(payload)`, and `redirect: "follow"`. This avoids a CORS preflight, which `script.google.com` does not answer. `doPost` must always return `ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON)` and report errors in the JSON body. GAS cannot set HTTP status codes.

---

## 5. TESTING & VERIFICATION

1. Tests use Vitest with in-memory GAS mocks (`test/gas-mocks/`). `test/setup.ts` calls `vi.resetModules()` and reinstalls fresh globals before every test, so import `src/` modules dynamically inside each test (`const load = () => import("../src/...")`). `vitest.config.ts` pins `TZ=America/Chicago`, which matches `appsscript.json`.
2. Mocks implement only what `src/` calls today. They do not evaluate formulas or adjust relative references on `copyTo`. When new code uses another GAS API (e.g. `PropertiesService`, `LockService`, `ContentService`, `Utilities`, `GmailLabel.addToThreads`, `GmailApp.moveThreadsToArchive`, `Range.getFormulasR1C1`), extend the mocks minimally and add tests.
3. Every behavior change or new feature needs tests. Backend API handlers are tested by calling the exported `doPost`/`doGet` with mock event objects.
4. **Before completing any task, run and pass all of:** `pnpm run format:check`, `pnpm run lint`, `pnpm run typecheck`, `pnpm test`, and `pnpm run build`. This is the same set as Husky pre-commit plus CI. After the workspace migration, these root scripts must cover every package.
