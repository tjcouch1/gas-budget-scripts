# Fast Budget PWA & Google Sheets API — Master Architecture & Execution Plan

> **Purpose of this Document:** A self-contained blueprint for an AI coding agent (or developer) in a fresh session to implement a **prototype** GitHub Pages PWA frontend for the budgeting spreadsheet. It covers the manual setup checklist, priorities, the spreadsheet schema, the existing `gas-budget-scripts` architecture (verified against `src/`), security rules, the API contract, conflict resolution, the Gmail performance fix, testing/CI, and ordered implementation steps.
>
> Read `AGENTS.md` first. Its rules override anything here.
>
> **Status markers used below:** **[VERIFIED]** = checked against `src/`/`test/` at the time of writing. **[USER-PROVIDED]** = describes the spreadsheet itself, which is not visible to code. Treat it as the user's description, and the agent must code defensively against it. **[ASSUMPTION]** = a decision made without verifying the real sheet; code defensively and re-check it during manual testing (§7). **[OPEN Qn]** = a genuinely unresolved question (§7).

---

## 0. Manual User Setup Checklist (Things You Configure Outside the Code)

The agent cannot do these. Phases are referenced by the step numbers in §5.

### A. Google Sheet Setup (before testing Step 2's API)

1. **Add 4 new rows to the per-sheet `SheetVariables` block.**
   - The `SheetVariables` value in `Variables!A1:B10` is an A1 range (e.g. `H1:I3`). That same range holds `[key, value]` pairs on **every** non-`Variables` tab.
   - Expand the `SheetVariables` value by 4 rows (e.g. `H1:I3` → `H1:I7`). Then, on the **Template tab** (`TemplateName`) and on **every pay-period tab you want to view in the app** (at least the current one), add:
     - **`EnvelopesStart`**: first envelope-name data cell of the Envelope table (e.g. `J4`).
     - **`EnvelopesMax`**: max number of envelope rows (e.g. `30`).
     - **`PaidWithStart`**: first data cell of the Paid With table.
     - **`PaidWithMax`**: max number of rows in the Paid With table (e.g. `15`).
   - Before expanding, check that the 4 new rows of that range are **empty on all existing tabs**. Any non-empty cells there would be read as extra keys. The code treats the 4 new keys as **optional** (§3.3), so older tabs without them keep working for receipt import and native splitting. The app only shows envelopes for tabs that have them (it hides envelope balances on tabs without them; see §3.3). The owner adds the keys to the Template and to the latest few pay-period tabs. Older tabs simply lack envelope data in the app.
2. Tabs created by `addTransactionSheet()` copy the Template, so they pick up the new keys automatically.

### A2. First manual test target

Point the deployed web app (and the PWA's Settings dialog) at your **test** budget sheet first: `main` branch ↔ test sheet, `deploy-budget-sheet` ↔ real sheet. Nothing in the implementation depends on which sheet is used.

### B. Google Apps Script Setup (after Step 2 code exists)

1. **clasp config after the monorepo move (Step 1).** Move your local, gitignored `.clasp.json` (and any `.clasp - prod.json` / `.clasp - test.json` variants) into `packages/budget-scripts-gas/`, keeping `"rootDir": "./dist"`. `.claspignore` moves there in the same commit. Run `pnpm run login` once if needed.
2. **Push the code yourself** (agents must not, per `AGENTS.md` §2): `pnpm run push` from the repo root. **Test against the test budget sheet first** (the one connected on `main`).
3. **Re-authorize** if the editor prompts (Step 2 adds `PropertiesService`, `LockService`, `ContentService`, `Utilities`, and the Gmail batch label/archive calls; the Gmail and Sheets scopes are the same as before, so a prompt may not appear). Deploying as a Web App (item 5) asks you to authorize as the deploying user.
4. **Generate & store `API_KEY`.** In the Apps Script editor (`pnpm run open`), choose `setupGenerateApiKey` in the **Run** dropdown and run it. It stores a random 64-hex-character key in **Project Settings → Script Properties** as `API_KEY` and prints it once to the Execution log. Copy it into the PWA Settings dialog. It refuses to overwrite an existing key: to rotate it, delete the `API_KEY` property in **Project Settings → Script Properties**, then run `setupGenerateApiKey` again; old clients then get `unauthorized`. You can also create the `API_KEY` property manually. Until `API_KEY` exists, every API call returns `unauthorized`.
5. **Deploy as a Web App.** **Deploy → New deployment → type: Web app**:
   - **Execute as:** `Me`. All requests run as the sheet owner, so a partner using the PWA acts with the owner's Sheet/Gmail access.
   - **Who has access:** `Anyone`. Do not pick "Anyone with Google account": browser `fetch()` from GitHub Pages cannot carry Google auth cookies, and requests fail on a login redirect. Security comes from the `API_KEY`.
   - Copy the `https://script.google.com/macros/s/<deployment-id>/exec` URL into the PWA Settings dialog. **Never commit it.**
6. **After every later code change:** `pnpm run push` updates only the editor's HEAD. The `/exec` URL keeps serving the version it was deployed with until you go to **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. The URL stays the same. (CLI alternative: `pnpm exec clasp deploy -i <deployment-id>`.)
7. **Triggers:** your existing installed triggers keep working unchanged: `onEditInstalled` on edit, and the time-driven trigger that runs "import 25 and mark" (`getAndRecordReceiptsAndMark`) every 12 hours **[USER-PROVIDED]**. Scheduled runs must keep working with no changes (§4.4 locking). Do not recreate or edit the trigger. Since Step 2 the receipt imports and the split checkbox wait up to 30 s for the script lock; if a web-app write holds it longer, the scheduled run fails and Apps Script emails you (the next run 12 h later retries).
8. **Verify the bundle loads in Apps Script V8** (Step 2). Agents cannot run GAS, so after the first `pnpm run push`, open the editor, run `onOpen`/a menu action once, and confirm there is no syntax error. **Decision (Step 2):** the es2022 bundle contained class fields and `#private` fields (`ThreadList`, `ReceiptInfo`, `ApiError`). The repo had no evidence that an es2022 esbuild bundle had ever run in Apps Script (the deployed branch predates the esbuild pipeline), so `scripts/build.ts` now uses `target: "es2019"`. esbuild then lowers class fields, `#private`, `?.`, `??` and `??=` to plain ES2019 helpers with the same behavior. The bundle uses no DOM or Node APIs (checked: no `window`, `document`, `crypto`, `TextEncoder`, `require`, `process`, `Buffer`). If everything loads, you may try `es2022` later. It is not required.
9. **Check the [ASSUMPTION] items in §7** (envelope table layout, Paid With table, Tax Deductible column) against the real sheet during first manual testing.

### C. Gmail Labels (automatic)

`getGmailLabel(name)` in `budgeting.util.ts` creates labels on first use **[VERIFIED]**. Today the code uses `Receipts/Scripted` and `Receipts`. After Step 2.2 it also uses `Receipts/Ignored` and `Receipts/Error` (§4.2). Once Step 2.2 is deployed, optionally run `getAndRecordReceiptsAndMark` once from the editor. That labels the backlog of non-receipt threads so later runs skip them.

### D. GitHub Repository Setup (for hosting the PWA)

1. **Pages is not enabled yet on `tjcouch1/gas-budget-scripts`.** Go to **Settings → Pages → Build and deployment → Source: "GitHub Actions"** (not "Deploy from a branch"; there is no `gh-pages` branch). `.github/workflows/deploy-web.yml` (Step 4) then publishes `packages/budget-web/dist` from `main` (and on manual `workflow_dispatch`).
2. The published site is public. That is acceptable because the bundle contains no secrets. Default URL `https://tjcouch1.github.io/gas-budget-scripts/`, no custom domain.
3. Open-source license page: the web package generates `licenses.html` at install/build time (Step 3 item 8). Nothing to configure manually.
4. After the first successful deploy, open the Pages URL on a phone and use "Add to Home Screen" (installed PWAs keep IndexedDB storage longer on iOS), then enter the web app URL and API key in Settings.

---

## 1. Project Overview & Core Objective

An automated Google Sheets budget backed by a TypeScript, container-bound GAS project (`gas-budget-scripts`). It imports receipt data from Gmail (Chase, PayPal, Venmo, and receipts forwarded by a partner) into bi-weekly pay-period tabs. An in-sheet checkbox menu and split checkboxes, driven by the installed `onEditInstalled` trigger, let scripts run on mobile **[VERIFIED]**.

**The Problem:** The Sheets UI is slow to open on mobile and desktop, slow to recalculate, and tedious for finding, categorizing, and splitting expenses. Full receipt import (`getAndRecordReceiptsAndMark`) reportedly takes 150–360 s, close to GAS's 6-minute limit. Part of the cause is in the code (§4.2).

**The Solution:** A fast, mobile-first PWA on GitHub Pages (`packages/budget-web`) that talks to a JSON Web App endpoint (`doGet`/`doPost`) added to the existing GAS code (`packages/budget-scripts-gas`). Types, the fingerprint function, and math/formula utilities are shared through `packages/budget-shared`.

- **The Google Sheet remains the single source of truth.** Everything the PWA writes must be identical to what the existing scripts or a human editing the sheet would write. Existing behavior (triggers, menu, custom formulas `sheetNameOffset`/`getMonthOffset`, native split, receipt import output) must stay identical, except for the explicitly listed Gmail-labeling change in §4.2. The existing tests in `test/` are the safety net for this.
- The PWA must work on **Android, iOS, Windows, and macOS** for **at least 2 users** (owner + partner). All GAS executions run as the owner (`Execute as: Me`).

---

## 2. Strictly Ordered Priorities & Requirements

### Priority #1: Ease & Speed of Finding, Categorizing, and Splitting Expenses

- **Instant Discovery:** Show newly imported/uncategorized expenses (blank `Envelope`) on open, with fast search/filter across name, envelope, paid-with, and date.
- **Fast Categorization and Editing:** UX priority order is (1) **categorize (`Envelope`) and split must be the fastest, smoothest actions**: tapping a transaction shows envelope chips right away (one tap to assign, optimistic), with a quick-split entry point on the same sheet; (2) **`Cost` edits** next; (3) `Name`, `Paid with`, and `Tax Deductible` (`Yes` or blank) live in a secondary "Edit details" sheet. The prototype supports editing **all** of Name, Cost, Envelope, Paid with, Tax Deductible, and Split, each with fingerprint conflict checks (§4.4).
- **First-Class Splitting UX (Inline Math & Tax Multiplier):**
  - Split one transaction into 2+ rows across envelopes in one action.
  - Split amounts accept **inline arithmetic** (`1+3.23*2`, `4.23*2+12.09+3.44*6`) for line-item splitting.
  - Each split row has a **"Pre-Tax (apply tax multiplier)"** toggle that uses `TaxMultiplier` from the API (never hardcoded):
    - **ON:** child `Cost` = `=${TaxMultiplier}*(4.23*2+12.09+3.44*6)`.
    - **OFF:** `=1+3.23*2` (no outer parentheses) for an expression, or the raw number for a single numeric literal.
    - Parent `Cost` and `Name` follow the native conventions exactly (§3.4, `AGENTS.md` §3.3).
- **Manual Receipt Sync:** a refresh button that imports up to 25 new Gmail receipts (`syncReceipts`). Target latency is a few seconds once the backlog is labeled. It is not guaranteed, because GAS cold starts alone take 1–3 s.

### Priority #2: Envelopes & Instant Balance Visibility

- Installed PWA opens to cached envelope balances immediately from IndexedDB (stale-while-revalidate), then refreshes in the background.
- Default **4-metric view**: `Envelope Size | Spent | Remain | Remain Cumul`, with `Remain` the most prominent.
- Toggle to show the full 11-column table (YTD/Projected).

### Priority #3: Pay Periods & Big Picture

- 14-day pay-period tabs named `M/D/YY - M/D/YY`, with ` (Gap)` appended for gap periods. Created by `addTransactionSheet()`, e.g. `3/3/24 - 3/16/24`. Month/day are **not** zero-padded **[VERIFIED]**.
- YTD/Projected envelope metrics, per-transaction `Tax Deductible` totals, and `Paid with` methods (e.g. `Credit`, `<ForwardName> Venmo`; the import code produces `Credit` and `<prefix> PayPal/Venmo` types **[VERIFIED]**). Prototype scope is a read-only display of what the sheet already computes. Post-prototype items are in §6.

---

## 3. Existing Codebase Architecture & Spreadsheet Schema Reference

### 3.1 Existing Build & Tooling **[VERIFIED]**

- **Node/pnpm:** Volta-pinned `node 24.16.0`, `pnpm 10.34.5`. Single package today (no `pnpm-workspace.yaml`).
- **GAS bundler (`scripts/build.ts`):** esbuild bundles `src/index.ts` into an unminified IIFE (`var _entry = (() => {...})()`) at `dist/index.js`, resolved relative to the **current working directory**. It regex-scans `src/index.ts` for `export function <name>(` and `export { ... }`, then appends `function <name>(...args) { return _entry.<name>(...args); }` wrappers so GAS registers triggers (`onOpen`, `onEditInstalled`), menu functions, custom formulas (`sheetNameOffset`, `getMonthOffset`), and, once added, `doGet`/`doPost`. It also copies `appsscript.json` to `dist/`. `--watch --clasp` (`pnpm start`) live-pushes.
- **clasp:** `.clasp.json` (gitignored) uses `"rootDir": "./dist"`. `.claspignore` whitelists only `**/*.js` and `**/appsscript.json`.
- **Manifest:** `appsscript.json` sets timeZone `America/Chicago`, runtime V8, and has no `webapp` section yet.
- **Quality gates:** Husky pre-commit runs `format:check`, `lint`, `typecheck`, `test`. `.github/workflows/ci.yml` runs on pushes to `main` and on PRs: `pnpm install --frozen-lockfile`, `format:check`, `lint`, `typecheck`, `test:coverage`, `build`.
- **TypeScript:** `typescript` → `@typescript/typescript6` (for typescript-eslint). `@typescript/native` (TS 7) provides `tsc` (`typecheck` = `tsc --noEmit && tsc --noEmit -p tsconfig.test.json`).
- **Tests:** Vitest `^5`, `test/**/*.test.ts`, `TZ=America/Chicago`, setup resets modules and installs mocks before each test. Mocks: `MockSpreadsheetApp`/`MockSpreadsheet`/`MockSheet`/`MockRange` (values, formulas, notes, backgrounds, `copyTo`, checkboxes, `getUi` toggle via `uiAvailable`), `MockGmailApp`/`MockThread`/`MockLabel` (per-thread `addLabel`/`moveToArchive` only), and `Logger`. No `PropertiesService`, `LockService`, `ContentService`, `Utilities`, `getFormulasR1C1`, `GmailLabel.addToThreads`, or `GmailApp.moveThreadsToArchive` yet. (Step 2 added all of these; see Step 2 item 6.)

### 3.2 Global Variables (`Variables` sheet, `Variables!A1:B10`) **[VERIFIED]**

`src/util/variables.util.ts` reads `[key, value]` rows from `Variables!A1:B10` via `SpreadsheetApp.getActiveSpreadsheet().getRange(...)`. That already works in web-app executions. The earlier "fix `getActiveSheet()`" note is obsolete. All 7 keys are required (it throws otherwise). 3 rows remain free.

| Key                   | Type      | Description                                                                 |
| :-------------------- | :-------- | :-------------------------------------------------------------------------- |
| `Menu`                | A1 string | Top-left cell of the in-sheet checkbox menu                                 |
| `SheetVariables`      | A1 string | Range (same on every tab) holding per-sheet `[key, value]` pairs            |
| `PayPeriodDays`       | number    | Pay period length in days (e.g. `14`)                                       |
| `TemplateName`        | string    | Tab duplicated by `addTransactionSheet()`                                   |
| `TaxMultiplier`       | number    | Pre-tax split multiplier (e.g. `1.0825`)                                    |
| `ForwardEmailAddress` | string    | Address of the person forwarding receipts (added to the Gmail search query) |
| `ForwardName`         | string    | Prefix for that person's `Paid with` values (e.g. `<ForwardName> Venmo`)    |

### 3.3 Per-Sheet Variables (`SheetVariables` range on each tab)

| Key                    | Status                            | Description                                                                  |
| :--------------------- | :-------------------------------- | :--------------------------------------------------------------------------- |
| `TransactionsStart`    | Existing, required **[VERIFIED]** | Top-left cell (Date column, first transaction row) of the transactions table |
| `TransactionsMax`      | Existing, required **[VERIFIED]** | Max transaction rows on the tab                                              |
| `SplitCheckboxesStart` | Existing, required **[VERIFIED]** | First cell of the 1-column split-checkbox range (`TransactionsMax` rows)     |
| **`EnvelopesStart`**   | **NEW, optional**                 | First envelope-name data cell of the Envelope table                          |
| **`EnvelopesMax`**     | **NEW, optional**                 | Max envelope rows                                                            |
| **`PaidWithStart`**    | **NEW, optional**                 | First data cell of the Paid With table (left-most column = option names)     |
| **`PaidWithMax`**      | **NEW, optional**                 | Max Paid With rows                                                           |

**Do NOT add the new keys to `SHEET_VARIABLE_NAMES`.** `getSheetVariables()` throws if any listed key is missing, which would break receipt import into older tabs and the native split checkbox on them. Add them to the `SheetVariables` type as optional (`EnvelopesStart?: string`, etc.). The API validates them only when it reads envelopes or paid-with options. If they are missing, the API returns `envelopes: null` / `paidWithOptions: null` plus a `warnings` entry, not an error. **Decision:** these keys are optional in code; the owner adds them to the Template and the latest few sheets. Features that depend on them degrade gracefully on tabs lacking them (hide envelope balances; with no option list, the `Envelope`/`Paid with` chips are hidden and those fields are not editable on that tab).

### 3.4 Transaction Table Columns (from `TransactionsStart`)

Existing code touches columns 0–2 (`getTransactionsRange`, 3 cols) and 4–5 (`getTransactionMetadataRange` = `range.offset(0, 4, rows, 2)`) **[VERIFIED]**. The API reads **7 columns** (`offset(0, 0, TransactionsMax, 7)`):

- **0 `Date`**: `Date` object (imported receipts write `Date`s). Blank cells read as `""`.
- **1 `Name`**: merchant/description. Split convention: rows in a split group contain `|` (`"<Merchant> | <Item note>"`). `doTransactionRowInfosSeemEqual` treats rows as the same group when dates are equal AND (the text before the first `"|"`, right-trimmed, is equal OR the first ⌊min(len)/3⌋ characters are equal) **[VERIFIED]**. A name-cell note holds receipt notes/errors. Background `#FF0000` = error, `#E8A9CA` = note **[VERIFIED]**.
- **2 `Cost`**: number or formula. Empty cost on import gets background `#E8A9CA` **[VERIFIED]**.
- **3 `Current Balance`**: read-only sheet formula **[USER-PROVIDED]**. Never write it. Native split does not move it.
- **4 `Envelope`** (`category` in code) **[VERIFIED position]**.
- **5 `Paid with`** (`type` in code) **[VERIFIED position]**.
- **6 `Tax Deductible`**: `Yes` or blank **[USER-PROVIDED]**. Existing code never reads, writes, or **shifts** this column. Native split not moving cols 3 and 6 is tracked in a GitHub issue. The prototype must **not** change native split behavior, and the web split matches native for parity (new child's col 6 left blank).

**Native `splitTransaction(sheet, index)` behavior that the web split must reproduce [VERIFIED, pinned in `test/budgeting.test.ts`]:**

1. Finds the group: the clicked row plus following rows that "seem equal" to it. Rejects an empty row, a row whose previous row seems equal (middle of a group), or a group with no room after it.
2. Parent row: name gets `" | "` appended unless it already contains `"|"`. Cost: number `C` → `=C-R[k]C[0]`. Existing formula `=F` → `=F-R[k]C[0]`, where `k` = current group size, i.e. the new child's offset.
3. New child row, appended directly after the group: same date, name `"<parent name before ' | '> | "`, cost `=${TaxMultiplier}*(0)`, `Envelope` blank, `Paid with` copied from parent.
4. If the row after the group is non-empty, everything from there to the first empty row after the last used row shifts down one row via `copyTo` (values, formulas, notes, formats). This covers cols 0–2 and 4–5 only. The vacated row is cleared, including notes and backgrounds.
5. Real Sheets returns A1 from `getFormulas()`: after one split the parent reads back as e.g. `=5.5-C6`, so a second split produces `=5.5-C6-R[2]C[0]`. The mocks store text verbatim.

### 3.5 Envelope Table Columns (from `EnvelopesStart`) **[USER-PROVIDED]** **[ASSUMPTION]**

11 contiguous columns (`offset(0, 0, EnvelopesMax, 11)`), read-only (all formulas/sheet-managed). Read **display values** (`getDisplayValues()`) for presentation and raw values for numbers:
0 `Envelope` · 1 `Envelope Available Cumulative` · 2 `Envelope Size`* · 3 `Spent`* · 4 `Remain`* (primary) · 5 `Remain Cumul`* · 6 `Envelope Size YTD` · 7 `Spent YTD` · 8 `Spent Average YTD` · 9 `Envelope Size Year Projected` · 10 `Spent Year Projected` (*core 4).
The owner is unsure whether the table has header, total, or blank separator rows inside `EnvelopesMax`, so read defensively: skip rows whose col 0 is blank, skip rows that look like headers or totals (col 0 equals `Envelope`, `Total`, or `Totals`, case-insensitive, trimmed), and never throw on non-numeric cells (show the display value, treat the number as `null`). Envelope names in col 0 are also the category choices for the `Envelope` column. Unverified assumption: re-check against the real sheet during manual testing.

### 3.6 Paid With Table (from `PaidWithStart`) **[USER-PROVIDED]** **[ASSUMPTION]**

Read column 0 of `PaidWithMax` rows as the `Paid with` option list. Skip blanks, header-like rows (`Paid with`, `Total`), and the summary row labelled `Tax deductible` (compare case-insensitively, trimmed). The owner is unsure the `Tax deductible` row exists or how it is labelled, so this is tolerant by design and an unverified assumption to check during manual testing.

---

## 4. Key Architectural Decisions

### 4.1 Monorepo Structure (`pnpm` workspaces)

```text
gas-budget-scripts/
├── AGENTS.md                        # Coding & security guardrails (update §0 after migration)
├── web-plan.md                      # This plan
├── package.json                     # Workspace root: shared devDeps (TS6/TS7 aliases, eslint, prettier, vitest, husky) + delegating scripts
├── pnpm-workspace.yaml              # packages: ["packages/*"]
├── eslint.config.js, .prettierrc.json, .prettierignore   # stay at root, cover all packages
├── .github/workflows/ci.yml         # updated for the workspace
├── .github/workflows/deploy-web.yml # NEW: GitHub Pages deploy (Step 4)
├── .github/actions/get-volta-node-version/action.yml # NEW (Step 4): copied from the owner's game-boy-camera repo
└── packages/
    ├── budget-shared/               # @budget/shared — source-only TS package ("exports": "./src/index.ts"), no build step
    │   └── src/ index.ts, types.ts (DTOs), fingerprint.ts, math.ts (parser + formula builder), names.ts (split-name helpers, free-text guard)
    │   └── test/*.test.ts
    ├── budget-scripts-gas/          # @budget/scripts-gas — today's root GAS project, moved with git mv
    │   ├── appsscript.json, .claspignore, scripts/build.ts, tsconfig.json, tsconfig.test.json, vitest.config.ts
    │   ├── src/ index.ts (+ doGet/doPost/setupGenerateApiKey exports), api/ (router, auth, handlers, serialize), models/, util/
    │   └── test/ (existing tests + gas-mocks, plus api.test.ts)
    └── budget-web/                  # @budget/web — Vite + React + TS + Tailwind + vite-plugin-pwa
        ├── vite.config.ts, index.html, public/ (icons; licenses.html is generated + gitignored)
        ├── scripts/generate-licenses.ts, license-checker.config.json, additional-licenses/ (Step 3 item 8)
        └── src/ crypto/ (credential vault), api/ (client), sync/ (cache, queue, conflicts), components/, test/
```

esbuild bundles `@budget/shared` into `dist/index.js` (it resolves workspace TS source). Vite does the same. Keep shared code free of DOM and GAS globals (pure TS, `lib: ES2022`).

### 4.2 Fixing the Slow Gmail Receipt Import

**Why it is slow [VERIFIED]:**

1. `ThreadList`'s constructor drops threads with no receipts and no errors (`hasInformation() === false`), so non-receipt emails (statements, alerts, marketing) from the four senders are never labeled. The search query excludes only `label:receipts`, `label:receipts-cru-reimburse`, `label:receipts-tax-deductible`, and `label:receipts-scripted`, plus `in:inbox`. These threads are therefore re-fetched and re-parsed on every run.
2. Threads with errors are **written to the sheet** (red name cell) but skipped when marking, so every run re-parses them **and writes duplicate rows**.
3. `markThreadProcessed()` calls `addLabel` ×2 + `moveToArchive` per thread (3 RPCs each).

**Change (implement exactly; this is the one intended behavior change. It is the **intended fix** for the GitHub issue about failed emails being re-recorded on every run, and that issue owns the final label/archive decision, so follow it if it differs):**

1. `ThreadList` keeps the threads it currently discards in a new `ignoredThreads: GmailThread[]`. `threadInfos` keeps its current contents so `getAllReceiptInfos()` output is unchanged.
2. In `recordReceipts(threadList, shouldMarkProcessed)`, when `shouldMarkProcessed` is true, derive `valid` (no errors), `errored` (errors > 0), and `ignored`. Then call the batch APIs in chunks of ≤100:
   - `Receipts/Scripted` → `[...valid, ...errored, ...ignored]`
   - `Receipts` → `valid`, then `GmailApp.moveThreadsToArchive(valid)`
   - `Receipts/Error` → `errored` (left in inbox, so the user sees it)
   - `Receipts/Ignored` → `ignored` (left in inbox)
   - When `shouldMarkProcessed` is false, keep the existing `Would mark thread ... processed` log per valid thread.
3. Keep the error count, alert, and throw behavior (`numErrors` → `getUi().alert`, else throw) unchanged for menu/trigger callers. Extract the core into `recordReceiptsCore(...)` that **returns** `{ receiptInfosAdded, numErrors, errors }` without touching `getUi()`. `recordReceipts` wraps it (alert/throw as today), and the API calls the core directly.
4. Update the pinned tests in `test/budgeting.test.ts` (“labels … and archives threads when marking”, “does not mark errored threads”, “does nothing for a non-receipt-only thread”) to the new expectations. Add `addToThreads` and `moveThreadsToArchive` to the Gmail mocks. List each changed expectation in the commit/PR description.

### 4.3 Credential Security in the Browser (non-extractable `CryptoKey` vault)

- No URL or key in git, `.env`, or the bundle. Nothing plaintext in `localStorage`.
- `packages/budget-web/src/crypto/credential-vault.ts`:
  1. On first save, `crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"])`, then store the `CryptoKey` object in IndexedDB (structured-clonable).
  2. Encrypt `JSON.stringify({ gasWebAppUrl, apiKey })` with a fresh random 12-byte IV and store only `{ iv, ciphertext }` in IndexedDB.
  3. Decrypt inside the API client's request function only. Do not hold plaintext in React state or module globals. The Settings dialog shows the key masked and never pre-fills it.
  4. Put IndexedDB access behind a small `KeyValueStore` interface so tests can inject an in-memory store.
- **What this protects (be accurate):** the raw key bytes cannot be exported by JavaScript, so a copied IndexedDB/localStorage dump, a backup, or a browser-sync leak yields only ciphertext. It does **not** add OS-keychain encryption, and it does **not** stop malicious script running on the page from calling `decrypt`. That is why XSS hygiene in `AGENTS.md` §1 matters.
- Validate `gasWebAppUrl` on save: must match `^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec$`.
- iOS Safari may evict storage of sites not added to the Home Screen after about 7 days unused. The app then shows the Settings dialog again, and Home Screen install avoids it. Call `navigator.storage.persist()` where supported.

### 4.4 API Contract (GAS Web App)

**Transport [required by GAS/CORS]:** every data call is a `POST` to the `/exec` URL with `Content-Type: text/plain;charset=utf-8`, body `JSON.stringify(request)`, and `redirect: "follow"`. GAS answers with a 302 to `script.googleusercontent.com`, which returns the JSON with `Access-Control-Allow-Origin: *`. `doPost` reads `JSON.parse(e.postData.contents)`. GAS cannot see request headers or set status codes. `doGet` returns only `{ ok: true, data: { service: "budget-api", apiVersion: 1 } }` (no auth, no data) for a connectivity check.

**Request envelope** (types in `@budget/shared/types.ts`):

```ts
type ApiRequest =
  | { apiVersion: 1; apiKey: string; action: "getState"; sheetName?: string }
  | {
      apiVersion: 1;
      apiKey: string;
      action: "updateTransaction";
      sheetName: string;
      rowIndex: number;
      baseFingerprint: string;
      changes: Partial<{
        name: string;
        costExpression: string; // validated by the shared math parser
        costPreTax: boolean; // default false; true applies TaxMultiplier like a split child
        envelope: string;
        paidWith: string;
        taxDeductible: "Yes" | "";
      }>;
      force?: boolean;
    }
  | {
      apiVersion: 1;
      apiKey: string;
      action: "splitTransaction";
      sheetName: string;
      rowIndex: number;
      baseFingerprint: string;
      splits: SplitItemInput[];
      force?: boolean;
    }
  | { apiVersion: 1; apiKey: string; action: "syncReceipts" };
type SplitItemInput = {
  itemNote: string;
  expression: string;
  preTax: boolean;
  envelope: string;
  paidWith?: string;
};
type ApiResponse<T> =
  | { ok: true; data: T; warnings?: string[] }
  | {
      ok: false;
      error: {
        code:
          | "unauthorized"
          | "bad_request"
          | "not_found"
          | "conflict"
          | "busy"
          | "server_error";
        message: string;
      };
      conflict?: { current: TransactionRowDto };
    };
```

- `rowIndex` = 0-based index from `TransactionsStart`, as used by `splitTransaction`.
- Auth: compare `apiKey` with `PropertiesService.getScriptProperties().getProperty("API_KEY")` using a constant-time compare. If the property is unset, always return `unauthorized`.
- Validate every field server-side: the sheet name must be one of `getTransactionSheetInfos()`, `rowIndex` must be within `[0, TransactionsMax)`, `envelope`/`paidWith` must be in the sheet's option lists (or blank) when those lists are available, and `expression`/`costExpression` must parse with the shared parser. Reject free-text (`name`, `itemNote`, `envelope`, `paidWith`) starting with `= + - @` (AGENTS.md §1.4). The prototype edits Name, Cost, Envelope, Paid with, Tax deductible, and Split. It does **not** edit Date, and never writes col 3 (`Current Balance`). Cost edits write `formatSplitCostFormula`-style output (number or `=expr`, tax-wrapped only when `costPreTax`). Cost edits on a split **parent** (a cost formula that subtracts its children's cells) are rejected with `bad_request` ("edit the child rows"). `name` edits on a row in a split group (name contains `|`) are rejected with `bad_request` in the prototype, because they could break group matching.
- Mutations run under `LockService.getScriptLock().tryLock(20000)`. On failure return `busy`. The client retries with backoff.
- Wrap the handler in try/catch: unexpected errors → `server_error` with the message. Never include stack traces containing sheet data beyond the message.
- **`getState` response (`PayPeriodStateDto`):**
  - `sheets: { name, startDate, endDate, isGap }[]` from `getTransactionSheetInfos()`, sorted newest first. Dates are `yyyy-MM-dd` via `Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), …)`.
  - `sheetName`: the requested one, else the sheet whose `[startDate, endDate]` contains now, else the newest.
  - `taxMultiplier`, `forwardName`.
  - `transactions: TransactionRowDto[]`: one batch each of `getValues`, `getFormulas`, `getDisplayValues`, `getNotes`, `getBackgrounds` on the 7-column range. Trailing empty rows are trimmed, interior empty rows kept as `null`. Each DTO has `{ rowIndex, date: "yyyy-MM-dd" | "", name, cost: number | null, costFormula: string | null, costDisplay, envelope, paidWith, taxDeductible, note, isError, isNote, splitGroupId, fingerprint }`. `splitGroupId` uses the same "seem equal" logic, exported from `budgeting.util.ts` (not re-implemented).
  - `envelopes: EnvelopeRowDto[] | null`, `paidWithOptions: string[] | null`, `envelopeOptions: string[] | null`.
- **Fingerprint (`@budget/shared/fingerprint.ts`, same function on both sides):** a 32-bit FNV-1a hex digest (pure TS, synchronous, so it runs in GAS without `Utilities`) of `JSON.stringify([date, name, costFormula ?? cost, envelope, paidWith, taxDeductible])`. Field values are normalized exactly as in the DTO: dates as `yyyy-MM-dd` in the spreadsheet time zone, strings trimmed only at the right, empty cells as `""`. The server recomputes it from a fresh batch read of that one row inside the lock.
- **`updateTransaction`:** inside the lock, re-read row `rowIndex`. If its fingerprint ≠ `baseFingerprint` and `!force`, return `conflict` with `current`. Otherwise write only the changed cells: Name/Cost with one `setValues` on `offset(rowIndex, 1, 1, 2)` and cols 4–6 with one `setValues` on `offset(rowIndex, 4, 1, 3)`, writing unchanged values back as read (an unchanged Cost formula must be written back from `getFormulas()`, not from its value). Never touch col 0 or col 3. Return the updated DTO.
- **`splitTransaction`:** inside the lock, check the fingerprint the same way. Then perform the split by **generalizing the existing `Budgeting.splitTransaction`** so it accepts N child specs, not by writing a second implementation. The checkbox path must keep calling it with the current defaults and stay byte-for-byte identical (existing tests must pass unmodified). For each child: name = `"<parent prefix> | <itemNote>"`, cost = `formatSplitCostFormula(expression, preTax, taxMultiplier)`, envelope = given, paid-with = given or the parent's. Parent formula: append `-R[k]C[0]` for each new child offset `k` to the existing number/formula (rule 2 of §3.4). That yields `=C-R[1]C[0]-…-R[N]C[0]` for an unsplit parent and stays correct for an already-split group. Shift following rows down by N in one `copyTo` per column block. Use the same columns as native (see the GitHub issue on cols 3/6; do not change native split in the prototype). Return the refreshed `getState` for that sheet.
- **`syncReceipts`:** under the lock, call `getChaseReceipts(0, 25)` + `recordReceiptsCore(threadList, true)` (§4.2). Return `{ receiptInfosAdded, numErrors, errors }` plus the refreshed state. It never calls `getUi()`. **The API never auto-creates pay-period sheets** (conservative default). If a receipt fits no sheet, the existing error ("Could not find transaction sheet…") is returned as `bad_request`, and the UI shows "No sheet for this pay period" with an explicit button that calls the existing add-transaction-sheet action. Scheduled (time-driven) runs keep their existing behavior.
- **Locking for existing entry points:** wrap the bodies of `getAndRecordReceipts*` exports and the split branch of `onEditInstalled` in the same script lock (`waitLock(30000)`) so the owner's 12-hourly time-driven import and a PWA sync cannot record the same threads twice. This changes no outputs, and scheduled runs must keep working unchanged (if the lock cannot be obtained within the wait, the scheduled run fails loudly rather than silently skipping; Apps Script emails failures). Add a test that the lock is acquired and released.
- **Manifest:** add `"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }` to `appsscript.json`, so deployments made with `clasp deploy` match §0.B.5.

**Implementation notes (Step 2).** These refine the contract above. All type changes are additive and optional:

- **Errors and order of checks:** a missing or non-JSON body returns `bad_request`. Next, the API key is checked (`unauthorized`, before anything else in the request is read). Then `apiVersion` and `action` are checked (`bad_request`). Field validation happens before the lock, so invalid requests never take it. Inside the lock, the fingerprint is checked (`conflict`), and only then the row-dependent rules: an empty row, a split-parent cost edit, and a name edit on a `|` row are `bad_request`. Empty rows are not in `getState`, so only a `force` request can reach the empty-row check.
- **Option lists:** blank `envelope`/`paidWith` is always allowed. A non-blank value must be in the sheet's list, and is rejected with `bad_request` when the sheet has no list (§3.3: not editable on tabs without the optional keys). A split item's `paidWith` that is omitted, `null`, or `""` copies the parent's value, like the native split. `splits` holds 1-50 items, and `preTax` is required on each.
- **`costPreTax`** without `costExpression` is `bad_request`. Unknown keys in `changes` (e.g. `date`) are `bad_request`.
- **Split parent detection** uses `getFormulasR1C1()` on the cost cell: a relative reference to a row below in the same column (`R[k]C` / `R[k]C[0]`, k ≥ 1).
- **Native-error mapping:** the native split's sheet-content errors (no content, middle of a group, no room, not enough empty rows) and the import's "Could not find transaction sheet" / "not enough empty transaction rows" errors become `bad_request`. Anything else is `server_error` with the message only.
- **`splitGroupId`** is the 0-based `rowIndex` of the group's first row (as a string), only for groups of 2 or more rows. Single rows get `null`.
- **Dates:** `sheets[].startDate/endDate` are formatted at the fixed GMT-06:00 offset they are parsed at by `getTransactionSheetInfos`. With the spreadsheet zone, end dates would show the next day during daylight saving time. Transaction dates use `Utilities.formatDate` in the spreadsheet time zone. Sheets whose names match the pattern but do not parse as dates are left out.
- **Shared types (additive):** `SyncReceiptsResultDto.receiptInfosAdded` is the total row count, and a new optional `receiptInfosAddedBySheet` holds the per-sheet map. `ServiceInfoDto` types the `doGet` data.
- **Locking:** `getState` is read-only and takes no lock. Mutations and `syncReceipts` use `tryLock(20000)` and return `busy` on failure. The 4 `getAndRecordReceipts*` exports and the split checkbox use `waitLock(30000)`. Every path releases the lock in `finally`.

### 4.5 Sync Engine, Offline & Conflict Resolution (frontend)

- **Cache:** the last `getState` per sheet goes in IndexedDB (`idb` library). It is budget data, not credentials, and is stored **unencrypted** (accepted for the prototype; encryption at rest is on the roadmap, §6), with a "Clear local data" button in Settings. Render from cache first, then refresh with `getState` on open, on focus, and after each successful mutation.
- **Status pill:** `Synced` · `Refreshing…` · `Syncing N changes…` · `Offline (N queued)` · `Conflict`.
- **Pending queue (IndexedDB, FIFO):** `{ id, sheetName, rowIndex, baseFingerprint, baseRow, action, payload, createdAt }`. The UI applies the change optimistically. Flush when online, one mutation at a time, and stop at the first `conflict` or `busy`. A split also takes a `rowIndex`, so later queued mutations on the same sheet are re-based after a split succeeds: re-pull, then drop and re-queue mutations whose rows moved (match by fingerprint within the refreshed state). If no match is found, they become conflicts.
- **Conflicts:** the server's fingerprint check is authoritative. Pulling on the client first is only an optimization. On `conflict`, revert the optimistic change, store `{ baseRow, proposed, current }`, and show the modal with **Original**, **Your discarded change**, **Current sheet value**:
  - **Accept Remote** → drop the mutation and keep the remote value.
  - **Overwrite with Mine** → resend with `baseFingerprint = current.fingerprint` (not `force`), so a third concurrent change is still caught.
- **Service worker (`vite-plugin-pwa`, `generateSW`):** precache the app shell only. Never cache `script.google.com` / `script.googleusercontent.com` responses in the SW (`navigateFallbackDenylist` + no runtime caching for those hosts). Use `registerType: "prompt"` to show an "Update available" toast.
- **GitHub Pages base path:** `vite.config.ts` `base` is read from `process.env.PAGES_BASE` with default `"/gas-budget-scripts/"` (the game-boy repo uses `base: "./"`; either works if the PWA `scope`/`start_url` match). The manifest uses `start_url: "."` and `scope: "."`. Use a hash router (`HashRouter`) so deep links work without a 404 fallback.

---

## 5. Step-by-Step Prototype Implementation Plan

Each step ends with all quality gates green (`AGENTS.md` §5.4) and a separate commit (only if the user asked for commits).

### Step 1: Workspace Migration & `@budget/shared`

1. `git mv src scripts test appsscript.json .claspignore tsconfig.json tsconfig.test.json vitest.config.ts packages/budget-scripts-gas/`. Create `packages/budget-scripts-gas/package.json` with the GAS scripts (`build`, `start`, `push`, `pull`, `open`, `login`, `logout`, `connect`, `typecheck`, `test`, `test:coverage`) and its runtime-adjacent devDeps (`@google/clasp`, `@types/google-apps-script`, `esbuild`). Note in §0.B.1 (already written) that `.clasp.json` moves too.
2. Root `package.json`: keep shared devDeps (`typescript` TS6 alias, `@typescript/native`, eslint stack, prettier, vitest, `@vitest/coverage-v8`, husky, `@types/node`) and `volta`. Scripts:
   - `build` = `pnpm -r run build`, `typecheck` = `pnpm -r run typecheck`, `test` = `pnpm -r run test`, `test:coverage` = `pnpm -r run test:coverage`.
   - `push`/`start`/`open`/`login`/`pull` = `pnpm --filter @budget/scripts-gas run <same>`.
   - `lint`/`format`/`format:check` stay root-level over the whole repo.
3. Add `pnpm-workspace.yaml`. Keep `onlyBuiltDependencies`/esbuild handling working (`pnpm rebuild esbuild` note in README). Run `pnpm install` to regenerate the lockfile.
4. Update `eslint.config.js` globs (`packages/*/scripts/**`, `packages/*/*.config.{js,ts}`; `ignores` for `packages/*/dist/**`, `**/coverage/**`), `.prettierignore`, README paths, and `AGENTS.md` §0.
5. CI (`ci.yml`) keeps the same steps. They now delegate to the workspace through the root scripts. Verify `node-version-file: package.json` still finds `volta.node` at the root.
6. Create `packages/budget-shared` (source-only; `package.json` with `"name": "@budget/shared"`, `"type": "module"`, `"exports": { ".": "./src/index.ts" }`, `typecheck`/`test` scripts):
   - DTOs and request/response types from §4.4 (including `changes` with `name`, `costExpression`, `costPreTax`, `envelope`, `paidWith`, `taxDeductible`).
   - `isUnsafeFreeText(s): boolean` (starts with `=`, `+`, `-`, or `@` after trimming left) for the AGENTS.md §1.4 guard, in `names.ts`.
   - `evaluateMathExpression(expr): number | null`: a recursive-descent parser for numbers (incl. leading `.5`), unary minus, `+ - * /`, parentheses, and whitespace. No `eval`/`Function`. Return `null` on any invalid input, division by zero, or a non-finite result.
   - `formatSplitCostFormula(expr, applyTax, taxMultiplier): string | number`: trims, validates with the parser (throws on invalid), and strips a leading `=`. If `applyTax`, return `=${taxMultiplier}*(${expr})`. Else return `Number(expr)` if it is a single numeric literal (incl. negative), else `=${expr}`, never adding outer parentheses.
   - `computeRowFingerprint(row)` (§4.4) and `buildSplitChildName(parentName, itemNote)` / `buildSplitParentName(parentName)` that mirror §3.4 rule 2–3.
   - Vitest tests: `=1.0825*(0)`-style output, `1+3.23*2` → `=1+3.23*2`, `5` → `5`, `-5` → `-5`, `(1+2)` with tax OFF → `=(1+2)` (user-typed parentheses kept), invalid/injection strings (`=IMPORTXML(...)`, `1+`, `2/0`, `SUM(1)`) rejected, and fingerprint stability across key order and `Date` vs string inputs.
7. **`.gitignore`:** add `packages/*/dist/`, `packages/*/coverage/`, `.pnpm-store/`, and the generated license artifacts of the future web package (`packages/budget-web/public/licenses.html`, `packages/budget-web/licenses.json`, `packages/budget-web/licenses-raw.json`) now, so they are never committed (mirrors the owner's game-boy-camera repo). The license tooling itself (`generate-licenses.ts` etc.) is part of the web package and lands in **Step 3 item 8**, not Step 1. Add `pnpm.onlyBuiltDependencies` for `esbuild` in the root `package.json` if pnpm requires it.
8. Then `pnpm run build` must still produce a `dist/index.js` with the same discovered export list as before the move (compare console output).

### Step 2: GAS changes (`packages/budget-scripts-gas`)

> **Status: implemented** (branch `frontend-plan`). Items 1-7 are done. Remaining for the owner: §0.B.2-9 (push, `setupGenerateApiKey`, deploy the web app, check that V8 loads the bundle). Implementation notes and the small, additive contract details are in §4.4 under "Implementation notes (Step 2)".

1. **`variables.util.ts`:** add the 4 optional keys to the `SheetVariables` type only (§3.3). The `getActiveSpreadsheet()` change is already done.
2. **Gmail optimization (§4.2)** in `thread-list.model.ts` and `budgeting.util.ts`, plus mock and test updates.
3. **API:** `src/api/` (`router.ts`, `auth.ts`, `handlers.ts`, `serialize.ts`) implementing §4.4. Export `doGet`, `doPost`, and `setupGenerateApiKey` from `src/index.ts`. `setupGenerateApiKey` stores `Utilities.getUuid() + Utilities.getUuid()` (dashes removed) as `API_KEY` and `Logger.log`s it once. It should refuse to overwrite an existing key unless the code is edited to pass `true`.
4. Generalize `splitTransaction` (§4.4) and export the "seem equal" helper for `splitGroupId`.
5. Add the locking for existing entry points and the `webapp` manifest block (§4.4).
   - **Risk to check during manual testing (agents cannot run GAS):** confirm the esbuild output loads in Apps Script V8 (see §0.B.8). If class fields / `#private` syntax or other newer syntax causes problems, set `target: "es2019"` in `scripts/build.ts`.
6. **Mocks to add** in `test/gas-mocks/`: `PropertiesService` (in-memory `getScriptProperties().getProperty/setProperty`), `LockService` (`getScriptLock()` with `tryLock/waitLock/releaseLock` and a flag to simulate contention), `ContentService` (`createTextOutput(s).setMimeType(m)` returning an object with `getContent()`), `Utilities` (`getUuid`, `formatDate` using `Intl` with the given time zone), `MockSpreadsheet.getSpreadsheetTimeZone()` → `"America/Chicago"`, `MockRange.getFormulasR1C1()` (return the stored text), `MockLabel.addToThreads(threads)`, and `MockGmailApp.moveThreadsToArchive(threads)`. The batch methods should throw on >100 threads, to test chunking. Register them in `installGasGlobals`.
7. **Tests (`test/api.test.ts`)** call the exported `doPost({ postData: { contents: JSON.stringify(req) } })` and `JSON.parse(output.getContent())`:
   - Auth: missing, wrong, or unset key → `unauthorized`.
   - Validation errors, including formula injection in `itemNote`/`envelope`.
   - `getState` shape, date formatting, trailing-row trimming, and envelopes `null` + warning when the optional keys are missing.
   - `updateTransaction`: envelope/paid-with/tax-deductible edits write only cols 4–6; name/cost edits write only cols 1–2; col 0 and col 3 are never touched; cost expression becomes a number or `=expr` (tax-wrapped only with `costPreTax`); unsafe name/envelope text, invalid expressions, split-parent cost edits, and split-group name edits are rejected; fingerprint mismatch → `conflict` with `current`; Overwrite with the new fingerprint succeeds.
   - `splitTransaction`: N=1 output equals the checkbox split for the same input (parity test); N=3 with pre-tax mixes; split of an already-split group; following rows shifted by N; no-room error.
   - `syncReceipts`: records and labels, returns errors without throwing when `uiAvailable = false`.
   - `busy` when the lock is contended.
   - All existing tests still pass, except the Gmail-labeling expectations listed in §4.2 item 4.

### Step 3: Frontend PWA (`packages/budget-web`)

**Status: implemented (prototype).** Deviations and frontend assumptions about the API (the GAS backend must match):

- Package name is `@budget/web` (`pnpm --filter @budget/web ...`). Tabs use a tiny hash-based switch (`#/expenses`, `#/envelopes`) instead of `react-router`'s `HashRouter` (two views only). shadcn/ui was not used; Tailwind v4 classes only.
- `updateTransaction` is assumed to return the updated `TransactionRowDto` as `data` (a full `PayPeriodStateDto` is also accepted). `splitTransaction` is assumed to return the refreshed `PayPeriodStateDto` for the sheet. `conflict` responses carry `conflict.current`.
- `syncReceipts` returns `SyncReceiptsResultDto` (`state` included). A `bad_request` whose message contains "transaction sheet" is shown as "No sheet for this pay period"; the API has no add-sheet action, so the UI tells the owner to add the sheet in the spreadsheet (no button).
- Queued edits to the same row are merged; a later edit or split on a row with an earlier queued change rides on that change's `baseFingerprint` and is re-based to the server's returned fingerprint. After a split, later queued edits are re-based by fingerprint (else they become conflicts). "Overwrite with Mine" resends with `baseFingerprint = current.fingerprint`, never `force`.
- `EnvelopeRowDto.display`/`values` index 3 is treated as `Remain` (the plan's cols 1-10 after the name).

1. Scaffold Vite + React 19 + TypeScript + Tailwind CSS v4 + `vite-plugin-pwa`. Optional: shadcn/ui components copied into the repo (no runtime CDN). Pin exact versions in `package.json`. Add `dev`, `build` (`tsc --noEmit && vite build`), `preview`, `typecheck`, and `test` scripts. App icons go in `public/` (generated, generic budget icon). Run `vite build` with `PAGES_BASE` (§4.5).
2. `crypto/credential-vault.ts` + **Connection Settings dialog**: URL + key inputs, a "Test connection" button that calls `getState`, and "Forget credentials".
3. `api/client.ts`: a single `callApi(request)` that decrypts credentials, sends the POST (§4.4), parses the envelope, maps errors to typed exceptions, and times out after 60 s. It handles a non-JSON response (e.g. the HTML login page that appears when the deployment isn't `Anyone`) with a clear message.
4. `sync/` (`useBudgetSync`): cache, queue, status pill, and conflict modal (§4.5).
5. **Envelopes view:** pay-period selector (defaults to the current period); 4-metric compact list with `Remain` emphasized (negative in red); "Show all columns" toggle for the 11-column table (horizontal scroll on mobile). If `envelopes` is `null`, show a hint pointing to §0.A.
6. **Expenses view & drawer:** filter chips (`Uncategorized` = blank envelope, `All`, `Tax Deductible`, `Errors/Notes`), search over name/envelope/paid-with/date, and split groups shown nested.
   - **Row tap sheet (fast path):** tapping a transaction opens a bottom sheet with envelope chips on top (one tap assigns the envelope, optimistic), a prominent **Split** button, and an inline **Cost** field (accepts inline arithmetic). A secondary "Edit details" section holds Name, Paid with, and Tax Deductible (Yes or blank). Keep taps to a minimum for categorize and split.
   - **Split mode:** rows of {item note, expression input with live evaluated value, Pre-Tax toggle labelled with the actual `taxMultiplier`, envelope}. Shows a live **Remaining to allocate** = parent current value − Σ child values, using the evaluated (tax-applied) amounts. The parent's current value comes from its `cost`/`costDisplay`. Submit is disabled while any expression is invalid.
   - The UI never writes Date or col 3. Hide Cost editing for split parents and Name editing for split-group rows (the server rejects them anyway, §4.4). If no pay-period sheet exists for a date, show "No sheet for this pay period" with an explicit add-sheet button (never auto-create).
7. **Frontend tests** (Vitest, `environment: "jsdom"`, `@testing-library/react`): vault round-trip with an in-memory store and Node's WebCrypto (`globalThis.crypto.subtle`), checking that no plaintext reaches the store; API client request format (method, `text/plain` header, `redirect: "follow"`, body has `apiKey`) with mocked `fetch`; queue/conflict reducer (success, conflict → modal state, Accept Remote, Overwrite resends with the new fingerprint, offline queueing); split drawer remaining-to-allocate math and invalid-expression handling.
8. **Open-source license publishing (copy/adapt from the owner's `game-boy-camera-picture-extractor`, `packages/gbcam-extract-web`):**
   - Files in `packages/budget-web/`: `scripts/generate-licenses.ts`, `license-checker.config.json` (the same empty-field template: `name`, `version`, `description`, `licenses`, `copyright`, `licenseFile: "none"`, `licenseText: "none"`, `licenseModified: "no"`, `path: null`), `additional-licenses/additional-licenses.json` (+ `.schema.json`, `.types.ts`; entries for non-npm attributions, initially `[]`).
   - Dependencies (devDependencies): `license-checker`, `@types/license-checker`.
   - `package.json` script: `"postinstall": "node scripts/generate-licenses.ts"` (Node 24 runs TypeScript natively; same as the game-boy repo). It runs on every `pnpm install` and in CI before the build. `build` must also ensure `licenses.html` exists (run `pnpm run generate:licenses` first).
   - How it works: `license-checker` scans `<repo root>/node_modules/.pnpm` (the recursive install location in a pnpm workspace; the script computes it as `../../node_modules/.pnpm` from the package dir), using `license-checker.config.json` as `customPath`. For each package it extracts a copyright line from the license text, merges `additional-licenses.json` entries first (keys prefixed `00_`), then writes `licenses-raw.json`, `licenses.json`, and a self-contained dark/light HTML page to `public/licenses.html` (Vite copies it to `dist/`). The page's back link and icon must use relative paths that work under the Pages base path (`./`), not `/`, and its title is "Open Source Licenses - Fast Budget".
   - The three generated files are gitignored (Step 1 item 7). The app links to `./licenses.html` from Settings/About. Keep the root `LICENSE` (MIT) and copy it to `packages/budget-web/LICENSE` as the game-boy repo does.
   - Verify it still works with this repo's pnpm (`onlyBuiltDependencies`, lockfile frozen in CI) and that `postinstall` does not fail on a fresh clone before `public/` exists (`mkdir -p`, as the script does).

### Step 4: GitHub Pages Deploy & CI

**Status: implemented.** `.github/workflows/deploy-web.yml` (push to `main` with the path filter, plus `workflow_dispatch`; `PAGES_BASE=/<repo-name>/`), `.github/actions/get-volta-node-version/`. `ci.yml` is unchanged: the root scripts already cover `@budget/web` (`format:check`, `lint`, `typecheck`, `test:coverage`, `build`). The owner still has to enable Pages with source "GitHub Actions" (§0.D).

1. `.github/workflows/deploy-web.yml`: `on: push` to `main` with `paths: ["packages/budget-web/**", "packages/budget-shared/**", "pnpm-lock.yaml", ".github/workflows/deploy-web.yml"]` plus `workflow_dispatch`. Deploys from **`main`**; there is no `gh-pages` branch. Modeled on the owner's `game-boy-camera-picture-extractor/.github/workflows/deploy.yml` (which deploys from a `production` branch; here `main`), using the Pages-from-Actions source.
   - `permissions: { contents: read, pages: write, id-token: write }`, `concurrency: { group: pages, cancel-in-progress: false }`.
   - Steps: checkout → `pnpm/action-setup@v4` (10.34.5) → `setup-node@v4` (`node-version-file: package.json`, cache pnpm; or the copied `.github/actions/get-volta-node-version` composite action that reads `volta.node`, as the game-boy repo does) → `pnpm install --frozen-lockfile` → `pnpm --filter @budget/web run test` → `PAGES_BASE=/${{ github.event.repository.name }}/ pnpm --filter @budget/web run build` → `actions/configure-pages@v5` → `actions/upload-pages-artifact@v3` (`path: packages/budget-web/dist`) → a separate `deploy` job (`needs: build`, `environment: { name: github-pages, url: ${{ steps.deployment.outputs.page_url }} }`) running `actions/deploy-pages@v4`. The game-boy repo's release/version-bump jobs are **not** copied. `pnpm install` runs the web `postinstall` license generation, so `licenses.html` is in the artifact. Owner must enable Pages with source "GitHub Actions" first (§0.D).
2. `ci.yml` needs no new steps: the root scripts cover all packages after Step 1. Make sure `pnpm run build` builds the web app too, so CI catches build breaks.

### Step 5: Hand-off checklist (agent reports to user)

- What the user must do now: §0.A, §0.A2 (point the deployed app at the **test** sheet first), §0.B.2–9, §0.D (enable Pages with source "GitHub Actions").
- List of any changed test expectations (should be only §4.2 item 4).
- Which [ASSUMPTION] items were exercised, and any remaining open questions.

---

## 6. Post-Prototype Roadmap

1. Accounts & Tax tab (Paid With summary stats incl. the `Tax deductible` row). "Add pay-period sheets" (`addTransactionSheets`) button. Manual transaction entry (cash/check).
2. Editing Name/Cost/Date of existing rows (with formula-injection guard). Visual line-item split builder producing `=${TaxMultiplier}*(a+b)` or `=a+b`.
3. Per-user API keys / revocation, or Google OAuth (Google sign-in) instead of the shared key. With OAuth the `API_KEY` is replaced by verified Google identity plus an allowlist of permitted emails stored in Script Properties (no shared secret). Optional WebAuthn unlock of the vault.
4. Encrypt cached budget data at rest (the IndexedDB cache is unencrypted in the prototype).
5. Batch the remaining per-row I/O in `splitTransaction`/`recordReceiptsOnSheet` (with parity tests).

---

## 7. Open Questions & Unverified Assumptions

Resolved by the owner and folded into the plan above: optional per-sheet keys (§3.3), 12-hourly time-driven import (§0.B.7, §4.4), Pages from `main` via GitHub Actions (§0.D, Step 4), test sheet first (§0.A2), no auto-created sheets (§4.4), unencrypted cache for the prototype (§4.5, §6), shared API key for the prototype (§6), editable fields (§2, §4.4).

**Unverified assumptions** (owner does not know; the code is tolerant, re-check during manual testing against the real sheet):

| #   | Assumption                                                                                                                                     |
| --- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| Q1  | The envelope table has the 11 columns in §3.5, possibly with header/total/blank rows inside `EnvelopesMax`. The reader skips such rows.        |
| Q2  | The Paid With table's first column holds option names, possibly with a `Tax deductible` summary row. The reader skips it and header-like rows. |

**Still open, tracked as GitHub issues** (not blockers; the prototype keeps current behavior):

- Native split not moving cols 3 and 6 (see the GitHub issue). The prototype must not change native split.
- Failed/ignored Gmail threads being re-recorded and the label/archive policy (see the GitHub issue). §4.2 is the intended fix.
