# Fast Budget PWA & Google Sheets API — Master Architecture & Execution Plan

> **Purpose of this Document:** A self-contained blueprint for an AI coding agent (or developer) in a fresh session to implement a **prototype** GitHub Pages PWA frontend for the budgeting spreadsheet. It covers the manual setup checklist, priorities, the spreadsheet schema, the existing `gas-budget-scripts` architecture (verified against `src/`), security rules, the API contract, conflict resolution, the Gmail performance fix, testing/CI, and ordered implementation steps.
>
> Read `AGENTS.md` first. Its rules override anything here.
>
> **Status markers used below:** **[VERIFIED]** = checked against `src/`/`test/` at the time of writing. **[USER-PROVIDED]** = describes the spreadsheet itself, which is not visible to code. Treat it as the user's description, and the agent must code defensively against it. **[OPEN Qn]** = an unresolved decision only the user can make. See §7. Where an open question blocks a step, implement the stated **default** and keep it easy to change.

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
   - Before expanding, check that the 4 new rows of that range are **empty on all existing tabs**. Any non-empty cells there would be read as extra keys. The code treats the 4 new keys as **optional** (§3.3), so older tabs without them keep working for receipt import and native splitting. The app only shows envelopes for tabs that have them.
2. Tabs created by `addTransactionSheet()` copy the Template, so they pick up the new keys automatically.

### B. Google Apps Script Setup (after Step 2 code exists)

1. **clasp config after the monorepo move (Step 1).** Move your local, gitignored `.clasp.json` (and any `.clasp - prod.json` / `.clasp - test.json` variants) into `packages/budget-scripts-gas/`, keeping `"rootDir": "./dist"`. `.claspignore` moves there in the same commit. Run `pnpm run login` once if needed.
2. **Push the code yourself** (agents must not, per `AGENTS.md` §2): `pnpm run push` from the repo root. **Test against the test budget sheet first** (the one connected on `main`).
3. **Re-authorize** when the editor prompts (new services: `PropertiesService`, `LockService`, `ContentService`).
4. **Generate & store `API_KEY`.** In the Apps Script editor (`pnpm run open`), choose `setupGenerateApiKey` in the **Run** dropdown and run it. It stores a random key in **Project Settings → Script Properties** as `API_KEY` and prints it once to the Execution log. Copy it into the PWA Settings dialog. To rotate the key, run it again; old clients then get `unauthorized`. You can also create the `API_KEY` property manually.
5. **Deploy as a Web App.** **Deploy → New deployment → type: Web app**:
   - **Execute as:** `Me`. All requests run as the sheet owner, so a partner using the PWA acts with the owner's Sheet/Gmail access.
   - **Who has access:** `Anyone`. Do not pick "Anyone with Google account": browser `fetch()` from GitHub Pages cannot carry Google auth cookies, and requests fail on a login redirect. Security comes from the `API_KEY`.
   - Copy the `https://script.google.com/macros/s/<deployment-id>/exec` URL into the PWA Settings dialog. **Never commit it.**
6. **After every later code change:** `pnpm run push` updates only the editor's HEAD. The `/exec` URL keeps serving the version it was deployed with until you go to **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. The URL stays the same. (CLI alternative: `pnpm exec clasp deploy -i <deployment-id>`.)
7. **Triggers:** your existing installed triggers (`onEditInstalled` on edit, and any time-driven trigger for receipt import, [OPEN Q6]) keep working unchanged.

### C. Gmail Labels (automatic)

`getGmailLabel(name)` in `budgeting.util.ts` creates labels on first use **[VERIFIED]**. Today the code uses `Receipts/Scripted` and `Receipts`. After Step 2.2 it also uses `Receipts/Ignored` and `Receipts/Error` (§4.2). Once Step 2.2 is deployed, optionally run `getAndRecordReceiptsAndMark` once from the editor. That labels the backlog of non-receipt threads so later runs skip them.

### D. GitHub Repository Setup (for hosting the PWA)

1. **Settings → Pages → Build and deployment → Source: GitHub Actions.** `.github/workflows/deploy-web.yml` (Step 4) publishes `packages/budget-web/dist`.
2. GitHub Pages on a **private** repo requires a paid plan, and the published site is public either way. That is acceptable because the bundle contains no secrets ([OPEN Q7]).

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
- **One-Tap Categorization:** Set or change `Envelope`, `Paid with`, and `Tax Deductible` (`Yes`/`No`/blank) with optimistic UI.
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
- **Tests:** Vitest `^5`, `test/**/*.test.ts`, `TZ=America/Chicago`, setup resets modules and installs mocks before each test. Mocks: `MockSpreadsheetApp`/`MockSpreadsheet`/`MockSheet`/`MockRange` (values, formulas, notes, backgrounds, `copyTo`, checkboxes, `getUi` toggle via `uiAvailable`), `MockGmailApp`/`MockThread`/`MockLabel` (per-thread `addLabel`/`moveToArchive` only), and `Logger`. No `PropertiesService`, `LockService`, `ContentService`, `Utilities`, `getFormulasR1C1`, `GmailLabel.addToThreads`, or `GmailApp.moveThreadsToArchive` yet.

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

**Do NOT add the new keys to `SHEET_VARIABLE_NAMES`.** `getSheetVariables()` throws if any listed key is missing, which would break receipt import into older tabs and the native split checkbox on them. Add them to the `SheetVariables` type as optional (`EnvelopesStart?: string`, etc.). The API validates them only when it reads envelopes or paid-with options. If they are missing, the API returns `envelopes: null` / `paidWithOptions: null` plus a `warnings` entry, not an error ([OPEN Q5]).

### 3.4 Transaction Table Columns (from `TransactionsStart`)

Existing code touches columns 0–2 (`getTransactionsRange`, 3 cols) and 4–5 (`getTransactionMetadataRange` = `range.offset(0, 4, rows, 2)`) **[VERIFIED]**. The API reads **7 columns** (`offset(0, 0, TransactionsMax, 7)`):

- **0 `Date`**: `Date` object (imported receipts write `Date`s). Blank cells read as `""`.
- **1 `Name`**: merchant/description. Split convention: rows in a split group contain `|` (`"<Merchant> | <Item note>"`). `doTransactionRowInfosSeemEqual` treats rows as the same group when dates are equal AND (the text before the first `"|"`, right-trimmed, is equal OR the first ⌊min(len)/3⌋ characters are equal) **[VERIFIED]**. A name-cell note holds receipt notes/errors. Background `#FF0000` = error, `#E8A9CA` = note **[VERIFIED]**.
- **2 `Cost`**: number or formula. Empty cost on import gets background `#E8A9CA` **[VERIFIED]**.
- **3 `Current Balance`**: read-only sheet formula **[USER-PROVIDED]**. Never write it. Native split does not move it.
- **4 `Envelope`** (`category` in code) **[VERIFIED position]**.
- **5 `Paid with`** (`type` in code) **[VERIFIED position]**.
- **6 `Tax Deductible`**: `Yes`/`No`/blank **[USER-PROVIDED]**. Existing code never reads, writes, or **shifts** this column ([OPEN Q3]).

**Native `splitTransaction(sheet, index)` behavior that the web split must reproduce [VERIFIED, pinned in `test/budgeting.test.ts`]:**

1. Finds the group: the clicked row plus following rows that "seem equal" to it. Rejects an empty row, a row whose previous row seems equal (middle of a group), or a group with no room after it.
2. Parent row: name gets `" | "` appended unless it already contains `"|"`. Cost: number `C` → `=C-R[k]C[0]`. Existing formula `=F` → `=F-R[k]C[0]`, where `k` = current group size, i.e. the new child's offset.
3. New child row, appended directly after the group: same date, name `"<parent name before ' | '> | "`, cost `=${TaxMultiplier}*(0)`, `Envelope` blank, `Paid with` copied from parent.
4. If the row after the group is non-empty, everything from there to the first empty row after the last used row shifts down one row via `copyTo` (values, formulas, notes, formats). This covers cols 0–2 and 4–5 only. The vacated row is cleared, including notes and backgrounds.
5. Real Sheets returns A1 from `getFormulas()`: after one split the parent reads back as e.g. `=5.5-C6`, so a second split produces `=5.5-C6-R[2]C[0]`. The mocks store text verbatim.

### 3.5 Envelope Table Columns (from `EnvelopesStart`) **[USER-PROVIDED]** ([OPEN Q1])

11 contiguous columns (`offset(0, 0, EnvelopesMax, 11)`), read-only (all formulas/sheet-managed). Read **display values** (`getDisplayValues()`) for presentation and raw values for numbers:
0 `Envelope` · 1 `Envelope Available Cumulative` · 2 `Envelope Size`* · 3 `Spent`* · 4 `Remain`* (primary) · 5 `Remain Cumul`* · 6 `Envelope Size YTD` · 7 `Spent YTD` · 8 `Spent Average YTD` · 9 `Envelope Size Year Projected` · 10 `Spent Year Projected` (*core 4).
Skip rows whose col 0 is blank. Envelope names in col 0 are also the category choices for the `Envelope` column.

### 3.6 Paid With Table (from `PaidWithStart`) **[USER-PROVIDED]** ([OPEN Q2])

Read column 0 of `PaidWithMax` rows as the `Paid with` option list. Skip blanks and the summary row labelled `Tax deductible` (compare case-insensitively, trimmed).

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
├── .github/workflows/deploy-web.yml # NEW: GitHub Pages deploy
└── packages/
    ├── budget-shared/               # @budget/shared — source-only TS package ("exports": "./src/index.ts"), no build step
    │   └── src/ index.ts, types.ts (DTOs), fingerprint.ts, math.ts (parser + formula builder), names.ts (split-name helpers)
    │   └── test/*.test.ts
    ├── budget-scripts-gas/          # @budget/scripts-gas — today's root GAS project, moved with git mv
    │   ├── appsscript.json, .claspignore, scripts/build.ts, tsconfig.json, tsconfig.test.json, vitest.config.ts
    │   ├── src/ index.ts (+ doGet/doPost/setupGenerateApiKey exports), api/ (router, auth, handlers, serialize), models/, util/
    │   └── test/ (existing tests + gas-mocks, plus api.test.ts)
    └── budget-web/                  # @budget/web — Vite + React + TS + Tailwind + vite-plugin-pwa
        ├── vite.config.ts, index.html, public/ (icons)
        └── src/ crypto/ (credential vault), api/ (client), sync/ (cache, queue, conflicts), components/, test/
```

esbuild bundles `@budget/shared` into `dist/index.js` (it resolves workspace TS source). Vite does the same. Keep shared code free of DOM and GAS globals (pure TS, `lib: ES2022`).

### 4.2 Fixing the Slow Gmail Receipt Import

**Why it is slow [VERIFIED]:**

1. `ThreadList`'s constructor drops threads with no receipts and no errors (`hasInformation() === false`), so non-receipt emails (statements, alerts, marketing) from the four senders are never labeled. The search query excludes only `label:receipts`, `label:receipts-cru-reimburse`, `label:receipts-tax-deductible`, and `label:receipts-scripted`, plus `in:inbox`. These threads are therefore re-fetched and re-parsed on every run.
2. Threads with errors are **written to the sheet** (red name cell) but skipped when marking, so every run re-parses them **and writes duplicate rows**.
3. `markThreadProcessed()` calls `addLabel` ×2 + `moveToArchive` per thread (3 RPCs each).

**Change (implement exactly; this is the one intended behavior change, pending [OPEN Q4]):**

1. `ThreadList` keeps the threads it currently discards in a new `ignoredThreads: GmailThread[]`. `threadInfos` keeps its current contents so `getAllReceiptInfos()` output is unchanged.
2. In `recordReceipts(threadList, shouldMarkProcessed)`, when `shouldMarkProcessed` is true, derive `valid` (no errors), `errored` (errors > 0), and `ignored`. Then call the batch APIs in chunks of ≤100:
   - `Receipts/Scripted` → `[...valid, ...errored, ...ignored]`
   - `Receipts` → `valid`, then `GmailApp.moveThreadsToArchive(valid)`
   - `Receipts/Error` → `errored` (left in inbox, so the user sees it)
   - `Receipts/Ignored` → `ignored` (left in inbox; default for [OPEN Q4])
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
        envelope: string;
        paidWith: string;
        taxDeductible: "Yes" | "No" | "";
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
- Validate every field server-side: the sheet name must be one of `getTransactionSheetInfos()`, `rowIndex` must be within `[0, TransactionsMax)`, `envelope`/`paidWith` must be in the sheet's option lists (or blank) when those lists are available, and `expression` must parse with the shared parser. Reject free-text starting with `= + - @` (§AGENTS 1.4). Prototype scope does **not** edit Date, Name, or Cost of existing rows ([OPEN Q12]).
- Mutations run under `LockService.getScriptLock().tryLock(20000)`. On failure return `busy`. The client retries with backoff.
- Wrap the handler in try/catch: unexpected errors → `server_error` with the message. Never include stack traces containing sheet data beyond the message.
- **`getState` response (`PayPeriodStateDto`):**
  - `sheets: { name, startDate, endDate, isGap }[]` from `getTransactionSheetInfos()`, sorted newest first. Dates are `yyyy-MM-dd` via `Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), …)`.
  - `sheetName`: the requested one, else the sheet whose `[startDate, endDate]` contains now, else the newest.
  - `taxMultiplier`, `forwardName`.
  - `transactions: TransactionRowDto[]`: one batch each of `getValues`, `getFormulas`, `getDisplayValues`, `getNotes`, `getBackgrounds` on the 7-column range. Trailing empty rows are trimmed, interior empty rows kept as `null`. Each DTO has `{ rowIndex, date: "yyyy-MM-dd" | "", name, cost: number | null, costFormula: string | null, costDisplay, envelope, paidWith, taxDeductible, note, isError, isNote, splitGroupId, fingerprint }`. `splitGroupId` uses the same "seem equal" logic, exported from `budgeting.util.ts` (not re-implemented).
  - `envelopes: EnvelopeRowDto[] | null`, `paidWithOptions: string[] | null`, `envelopeOptions: string[] | null`.
- **Fingerprint (`@budget/shared/fingerprint.ts`, same function on both sides):** a 32-bit FNV-1a hex digest (pure TS, synchronous, so it runs in GAS without `Utilities`) of `JSON.stringify([date, name, costFormula ?? cost, envelope, paidWith, taxDeductible])`. Field values are normalized exactly as in the DTO: dates as `yyyy-MM-dd` in the spreadsheet time zone, strings trimmed only at the right, empty cells as `""`. The server recomputes it from a fresh batch read of that one row inside the lock.
- **`updateTransaction`:** inside the lock, re-read row `rowIndex`. If its fingerprint ≠ `baseFingerprint` and `!force`, return `conflict` with `current`. Otherwise write only the changed cells of cols 4–6 with one `setValues` on `offset(rowIndex, 4, 1, 3)`, writing back unchanged values as read. Return the updated DTO.
- **`splitTransaction`:** inside the lock, check the fingerprint the same way. Then perform the split by **generalizing the existing `Budgeting.splitTransaction`** so it accepts N child specs, not by writing a second implementation. The checkbox path must keep calling it with the current defaults and stay byte-for-byte identical (existing tests must pass unmodified). For each child: name = `"<parent prefix> | <itemNote>"`, cost = `formatSplitCostFormula(expression, preTax, taxMultiplier)`, envelope = given, paid-with = given or the parent's. Parent formula: append `-R[k]C[0]` for each new child offset `k` to the existing number/formula (rule 2 of §3.4). That yields `=C-R[1]C[0]-…-R[N]C[0]` for an unsplit parent and stays correct for an already-split group. Shift following rows down by N in one `copyTo` per column block. Use the same columns as native, pending [OPEN Q3]. Return the refreshed `getState` for that sheet.
- **`syncReceipts`:** under the lock, call `getChaseReceipts(0, 25)` + `recordReceiptsCore(threadList, true)` (§4.2). Return `{ receiptInfosAdded, numErrors, errors }` plus the refreshed state. It never calls `getUi()`. If a receipt fits no sheet, the existing error ("Could not find transaction sheet…") is returned as `bad_request` ([OPEN Q9]).
- **Locking for existing entry points:** wrap the bodies of `getAndRecordReceipts*` exports and the split branch of `onEditInstalled` in the same script lock (`waitLock(30000)`) so a time-driven import and a PWA sync cannot record the same threads twice. This changes no outputs. Add a test that the lock is acquired and released.
- **Manifest:** add `"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }` to `appsscript.json`, so deployments made with `clasp deploy` match §0.B.5.

### 4.5 Sync Engine, Offline & Conflict Resolution (frontend)

- **Cache:** the last `getState` per sheet goes in IndexedDB (`idb` library). It is budget data, not credentials, and is stored unencrypted ([OPEN Q10]). Render from cache first, then refresh with `getState` on open, on focus, and after each successful mutation.
- **Status pill:** `Synced` · `Refreshing…` · `Syncing N changes…` · `Offline (N queued)` · `Conflict`.
- **Pending queue (IndexedDB, FIFO):** `{ id, sheetName, rowIndex, baseFingerprint, baseRow, action, payload, createdAt }`. The UI applies the change optimistically. Flush when online, one mutation at a time, and stop at the first `conflict` or `busy`. A split also takes a `rowIndex`, so later queued mutations on the same sheet are re-based after a split succeeds: re-pull, then drop and re-queue mutations whose rows moved (match by fingerprint within the refreshed state). If no match is found, they become conflicts.
- **Conflicts:** the server's fingerprint check is authoritative. Pulling on the client first is only an optimization. On `conflict`, revert the optimistic change, store `{ baseRow, proposed, current }`, and show the modal with **Original**, **Your discarded change**, **Current sheet value**:
  - **Accept Remote** → drop the mutation and keep the remote value.
  - **Overwrite with Mine** → resend with `baseFingerprint = current.fingerprint` (not `force`), so a third concurrent change is still caught.
- **Service worker (`vite-plugin-pwa`, `generateSW`):** precache the app shell only. Never cache `script.google.com` / `script.googleusercontent.com` responses in the SW (`navigateFallbackDenylist` + no runtime caching for those hosts). Use `registerType: "prompt"` to show an "Update available" toast.
- **GitHub Pages base path:** `vite.config.ts` `base: "/<repo-name>/"`, read from `process.env.PAGES_BASE` with default `"/gas-budget-scripts/"`. The manifest uses `start_url: "."` and `scope: "."`. Use a hash router (`HashRouter`) so deep links work without a 404 fallback.

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
   - DTOs and request/response types from §4.4.
   - `evaluateMathExpression(expr): number | null`: a recursive-descent parser for numbers (incl. leading `.5`), unary minus, `+ - * /`, parentheses, and whitespace. No `eval`/`Function`. Return `null` on any invalid input, division by zero, or a non-finite result.
   - `formatSplitCostFormula(expr, applyTax, taxMultiplier): string | number`: trims, validates with the parser (throws on invalid), and strips a leading `=`. If `applyTax`, return `=${taxMultiplier}*(${expr})`. Else return `Number(expr)` if it is a single numeric literal (incl. negative), else `=${expr}`, never adding outer parentheses.
   - `computeRowFingerprint(row)` (§4.4) and `buildSplitChildName(parentName, itemNote)` / `buildSplitParentName(parentName)` that mirror §3.4 rule 2–3.
   - Vitest tests: `=1.0825*(0)`-style output, `1+3.23*2` → `=1+3.23*2`, `5` → `5`, `-5` → `-5`, `(1+2)` with tax OFF → `=(1+2)` (user-typed parentheses kept), invalid/injection strings (`=IMPORTXML(...)`, `1+`, `2/0`, `SUM(1)`) rejected, and fingerprint stability across key order and `Date` vs string inputs.
7. Then `pnpm run build` must still produce a `dist/index.js` with the same discovered export list as before the move (compare console output).

### Step 2: GAS changes (`packages/budget-scripts-gas`)

1. **`variables.util.ts`:** add the 4 optional keys to the `SheetVariables` type only (§3.3). The `getActiveSpreadsheet()` change is already done.
2. **Gmail optimization (§4.2)** in `thread-list.model.ts` and `budgeting.util.ts`, plus mock and test updates.
3. **API:** `src/api/` (`router.ts`, `auth.ts`, `handlers.ts`, `serialize.ts`) implementing §4.4. Export `doGet`, `doPost`, and `setupGenerateApiKey` from `src/index.ts`. `setupGenerateApiKey` stores `Utilities.getUuid() + Utilities.getUuid()` (dashes removed) as `API_KEY` and `Logger.log`s it once. It should refuse to overwrite an existing key unless the code is edited to pass `true`.
4. Generalize `splitTransaction` (§4.4) and export the "seem equal" helper for `splitGroupId`.
5. Add the locking for existing entry points and the `webapp` manifest block (§4.4).
6. **Mocks to add** in `test/gas-mocks/`: `PropertiesService` (in-memory `getScriptProperties().getProperty/setProperty`), `LockService` (`getScriptLock()` with `tryLock/waitLock/releaseLock` and a flag to simulate contention), `ContentService` (`createTextOutput(s).setMimeType(m)` returning an object with `getContent()`), `Utilities` (`getUuid`, `formatDate` using `Intl` with the given time zone), `MockSpreadsheet.getSpreadsheetTimeZone()` → `"America/Chicago"`, `MockRange.getFormulasR1C1()` (return the stored text), `MockLabel.addToThreads(threads)`, and `MockGmailApp.moveThreadsToArchive(threads)`. The batch methods should throw on >100 threads, to test chunking. Register them in `installGasGlobals`.
7. **Tests (`test/api.test.ts`)** call the exported `doPost({ postData: { contents: JSON.stringify(req) } })` and `JSON.parse(output.getContent())`:
   - Auth: missing, wrong, or unset key → `unauthorized`.
   - Validation errors, including formula injection in `itemNote`/`envelope`.
   - `getState` shape, date formatting, trailing-row trimming, and envelopes `null` + warning when the optional keys are missing.
   - `updateTransaction`: success writes only cols 4–6 and leaves cols 0–3 untouched; fingerprint mismatch → `conflict` with `current`; Overwrite with the new fingerprint succeeds.
   - `splitTransaction`: N=1 output equals the checkbox split for the same input (parity test); N=3 with pre-tax mixes; split of an already-split group; following rows shifted by N; no-room error.
   - `syncReceipts`: records and labels, returns errors without throwing when `uiAvailable = false`.
   - `busy` when the lock is contended.
   - All existing tests still pass, except the Gmail-labeling expectations listed in §4.2 item 4.

### Step 3: Frontend PWA (`packages/budget-web`)

1. Scaffold Vite + React 19 + TypeScript + Tailwind CSS v4 + `vite-plugin-pwa`. Optional: shadcn/ui components copied into the repo (no runtime CDN). Pin exact versions in `package.json`. Add `dev`, `build` (`tsc --noEmit && vite build`), `preview`, `typecheck`, and `test` scripts. App icons go in `public/` (generated, generic budget icon). Run `vite build` with `PAGES_BASE` (§4.5).
2. `crypto/credential-vault.ts` + **Connection Settings dialog**: URL + key inputs, a "Test connection" button that calls `getState`, and "Forget credentials".
3. `api/client.ts`: a single `callApi(request)` that decrypts credentials, sends the POST (§4.4), parses the envelope, maps errors to typed exceptions, and times out after 60 s. It handles a non-JSON response (e.g. the HTML login page that appears when the deployment isn't `Anyone`) with a clear message.
4. `sync/` (`useBudgetSync`): cache, queue, status pill, and conflict modal (§4.5).
5. **Envelopes view:** pay-period selector (defaults to the current period); 4-metric compact list with `Remain` emphasized (negative in red); "Show all columns" toggle for the 11-column table (horizontal scroll on mobile). If `envelopes` is `null`, show a hint pointing to §0.A.
6. **Expenses view & drawer:** filter chips (`Uncategorized` = blank envelope, `All`, `Tax Deductible`, `Errors/Notes`), search over name/envelope/paid-with/date, and split groups shown nested.
   - Drawer: Envelope / Paid with / Tax Deductible selectors (option lists from the API).
   - **Split mode:** rows of {item note, expression input with live evaluated value, Pre-Tax toggle labelled with the actual `taxMultiplier`, envelope}. Shows a live **Remaining to allocate** = parent current value − Σ child values, using the evaluated (tax-applied) amounts. The parent's current value comes from its `cost`/`costDisplay`. Submit is disabled while any expression is invalid.
   - The UI never writes Name, Date, or Cost of existing rows in the prototype.
7. **Frontend tests** (Vitest, `environment: "jsdom"`, `@testing-library/react`): vault round-trip with an in-memory store and Node's WebCrypto (`globalThis.crypto.subtle`), checking that no plaintext reaches the store; API client request format (method, `text/plain` header, `redirect: "follow"`, body has `apiKey`) with mocked `fetch`; queue/conflict reducer (success, conflict → modal state, Accept Remote, Overwrite resends with the new fingerprint, offline queueing); split drawer remaining-to-allocate math and invalid-expression handling.

### Step 4: GitHub Pages Deploy & CI

1. `.github/workflows/deploy-web.yml`: `on: push` to `main` with `paths: ["packages/budget-web/**", "packages/budget-shared/**", "pnpm-lock.yaml", ".github/workflows/deploy-web.yml"]` plus `workflow_dispatch` (branch per [OPEN Q7]).
   - `permissions: { contents: read, pages: write, id-token: write }`, `concurrency: { group: pages, cancel-in-progress: false }`.
   - Steps: checkout → `pnpm/action-setup` (10.34.5) → `setup-node` (`node-version-file: package.json`, cache pnpm) → `pnpm install --frozen-lockfile` → `pnpm --filter @budget/web run test` → `PAGES_BASE=/${{ github.event.repository.name }}/ pnpm --filter @budget/web run build` → `actions/configure-pages@v5` → `actions/upload-pages-artifact@v3` (`path: packages/budget-web/dist`) → `actions/deploy-pages@v4` in an `environment: github-pages` job.
2. `ci.yml` needs no new steps: the root scripts cover all packages after Step 1. Make sure `pnpm run build` builds the web app too, so CI catches build breaks.

### Step 5: Hand-off checklist (agent reports to user)

- What the user must do now: §0.A, §0.B.2–6, §0.D.
- List of any changed test expectations (should be only §4.2 item 4).
- Remaining open questions.

---

## 6. Post-Prototype Roadmap

1. Accounts & Tax tab (Paid With summary stats incl. the `Tax deductible` row). "Add pay-period sheets" (`addTransactionSheets`) button. Manual transaction entry (cash/check).
2. Editing Name/Cost/Date of existing rows (with formula-injection guard). Visual line-item split builder producing `=${TaxMultiplier}*(a+b)` or `=a+b`.
3. Per-user API keys / revocation, or Google OAuth instead of a shared key. Optional WebAuthn unlock of the vault.
4. Batch the remaining per-row I/O in `splitTransaction`/`recordReceiptsOnSheet` (with parity tests).

---

## 7. Open Questions (only the user can answer; defaults in use until answered)

| #   | Question                                                                                                                                                                                                    | Default used by the plan                                                                                              |
| --- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Q1  | Envelope table: are the 11 columns and their order in §3.5 exact? Are there header/total/blank separator rows inside `EnvelopesMax`?                                                                        | As listed. Skip rows with blank col 0. Read-only.                                                                     |
| Q2  | Paid With table: is the `Tax deductible` stats row inside the `PaidWithMax` range, and what is its exact label? Are the `Envelope`/`Paid with` cells data-validation dropdowns that reference these tables? | Skip a row labelled `Tax deductible` (case-insensitive). Validate writes against the lists.                           |
| Q3  | Is column offset 6 `Tax Deductible` on every transaction tab? The native split does **not** shift col 6 (or col 3) when moving rows down. Is that intended, or a bug to fix (in native split too)?          | Web split matches native (cols 0–2, 4–5 shifted) for parity. The new child's col 6 is left blank. Flag in the report. |
| Q4  | Gmail: OK to label errored threads `Receipts/Scripted` + `Receipts/Error` (stops automatic retries **and** stops duplicate red rows)? Should ignored/errored threads be archived or left in the inbox?      | Label both. Archive neither (only valid receipts are archived, as today).                                             |
| Q5  | Should the 4 new sheet variables be required on every tab (you backfill all tabs) or optional?                                                                                                              | Optional (§3.3).                                                                                                      |
| Q6  | Is there a time-driven trigger running receipt import? Which function and how often?                                                                                                                        | Assume yes (`getAndRecordReceiptsAndMark`). Hence the shared lock.                                                    |
| Q7  | Is GitHub Pages available on this repo (public, or paid plan)? Which branch should deploy the PWA (`main` or `deploy-budget-sheet`)? Custom domain?                                                         | Deploy from `main`. Default `github.io/<repo>/` URL.                                                                  |
| Q8  | Which sheet should the prototype be tested against first?                                                                                                                                                   | The test budget sheet bound to the `main` script project.                                                             |
| Q9  | If a new receipt's date has no pay-period tab, should `syncReceipts` call `addTransactionSheets()` first?                                                                                                   | No. Return the error. The user adds sheets via the menu.                                                              |
| Q10 | Is caching budget data (not credentials) unencrypted in the browser's IndexedDB acceptable on both users' devices?                                                                                          | Yes, with a "Clear local data" button in Settings.                                                                    |
| Q11 | Is one shared `API_KEY` for both users acceptable for the prototype?                                                                                                                                        | Yes.                                                                                                                  |
| Q12 | Should the prototype edit Name/Cost of existing (non-split) rows, or only categorize + split?                                                                                                               | Only categorize + split.                                                                                              |
