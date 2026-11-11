# Fast Budget PWA & Google Sheets API — Master Architecture & Execution Plan

> **Purpose of this Document:** This file is a self-contained blueprint designed for an AI coding assistant or developer in a fresh session. It contains the manual setup checklist, business priorities, pain points, spreadsheet schemas, existing codebase architecture (`gas-budget-scripts`), security rules, conflict-resolution specs, Gmail performance fixes, and step-by-step prototype instructions.

---

## 0. Manual User Setup Checklist (Things You Need to Configure Outside the Code)

Before or during testing of the prototype, complete these manual setup steps in your Google Sheet, Google Apps Script console, Gmail, and GitHub repository:

### A. Google Sheet Setup (Required Before Testing the API)
1. **Add 4 New Rows to the Per-Sheet `SheetVariables` Block:**
   - Open your Google Sheet and go to your **Template tab** (referenced by `TemplateName` in the `Variables` sheet) **and** your **current active `MM/DD/YY - MM/DD/YY` pay-period tab** (plus any recent tabs you want to view in the app).
   - Locate the per-sheet variables table (at the range defined by `SheetVariables` in `Variables!A1:B10`) where `TransactionsStart`, `TransactionsMax`, and `SplitCheckboxesStart` currently live.
   - Expand `SheetVariables` in the `Variables` sheet if needed so the range covers 4 additional rows, and add these 4 key-value pairs to the per-sheet variables block:
     - **`EnvelopesStart`**: Top-left data cell of the Envelope table (first envelope name cell, e.g., `J4`).
     - **`EnvelopesMax`**: Max number of envelope rows in the Envelope table (e.g., `30`).
     - **`PaidWithStart`**: Top-left data cell of the Paid With table (first payment option cell).
     - **`PaidWithMax`**: Max number of rows in the Paid With table (e.g., `15`).

### B. Google Apps Script (GAS) Setup (Required After Pushing Phase 1 Code)
1. **Update `.clasp.json` Path (Local Machine):**
   - When the GAS package moves to `packages/budget-scripts-gas`, either keep `.clasp.json` inside `packages/budget-scripts-gas/` with `"rootDir": "./dist"` or at repo root with `"rootDir": "./packages/budget-scripts-gas/dist"` (the build scripts will handle this automatically).
2. **Generate & Store `API_KEY` in GAS Script Properties:**
   - Run the one-time helper function `setupGenerateApiKey()` from the GAS Editor (`script.google.com`) **Run** dropdown (or add `API_KEY` manually under **Project Settings -> Script Properties**).
   - Copy the generated secret key from the Execution Log to paste into the PWA Settings modal once.
3. **Deploy as a Web App:**
   - In the GAS Editor at `script.google.com`, click **Deploy -> New deployment -> Select type: Web app**.
   - Set **Execute as:** `Me` (ensures both you and your partner interact with your Gmail and Google Sheet using your credentials).
   - Set **Who has access:** `Anyone` (required so browser `fetch()` requests work without CORS redirect failures; requests are authenticated via your `API_KEY`).
   - Copy the resulting `https://script.google.com/macros/s/.../exec` Web App URL to paste into the PWA Settings modal.

### C. Gmail Labels Setup (Automatic + Optional Clean-Up)
1. **Labels Created Automatically by the Script:**
   - `getGmailLabel(name)` in `budgeting.util.ts` automatically creates Gmail labels if they don't exist (`GmailApp.createLabel(name)`). The script will automatically create and manage:
     - `Receipts/Scripted` (applied to **every** email thread touched by the script—valid receipts, ignored non-receipts, and errored threads).
     - `Receipts` (applied to valid receipt threads).
     - `Receipts/Ignored` (applied alongside `Receipts/Scripted` to non-receipt emails from Chase/PayPal/Venmo/Partner so they are never re-scanned).
     - `Receipts/Error` (applied alongside `Receipts/Scripted` to threads that threw parsing errors so they don't block future runs).
2. **Optional One-Time Run:**
   - Once the new Gmail batch-labeling code is pushed, trigger `getAndRecordReceiptsAndMark` once from the GAS editor to clear out and label all historical non-receipt threads currently sitting in your inbox. Every run after that will take **<3 seconds**.

### D. GitHub Repository Setup (For Hosting the PWA)
1. In your GitHub repo settings, go to **Settings -> Pages** and set **Build and deployment -> Source** to **GitHub Actions** so `.github/workflows/deploy-web.yml` automatically publishes `packages/budget-web/dist`.

---

## 1. Project Overview & Core Objective
We have an automated Google Sheets budget backed by a TypeScript Google Apps Script (GAS) project (`gas-budget-scripts`). It pulls receipt data from Gmail (Chase, PayPal, Venmo, and forwarded partner receipts) into bi-weekly pay-period sheets and uses checkbox `onEdit` triggers so scripts can be run on mobile.

**The Problem:** The Google Sheets UI is slow to open and navigate on mobile and desktop, takes too long to recalculate dependencies, makes finding/categorizing/splitting expenses tedious, and the scheduled Gmail receipt sync (`getAndRecordReceiptsAndMark`) currently takes 150s–360s (nearing GAS's 6-minute hard timeout).

**The Solution:** Build a fast, mobile-first Progressive Web App (PWA) hosted on GitHub Pages (`packages/budget-web`) that talks to a headless JSON Web App endpoint (`doGet` / `doPost`) added to the existing GAS codebase (`packages/budget-scripts-gas`), sharing DRY types and math utilities via a shared package (`packages/budget-shared`).
- The Google Sheet remains the single source of truth and retains 100% backward compatibility (all formulas, split syntax, and cell notes work identically whether edited in the PWA or directly in the Google Sheet).
- The PWA works across **Android, iOS, Windows, and macOS** for **at least 2 users** (the primary owner + partner), while all GAS executions run under the primary owner's Google account (`Execute as: Me`).

---

## 2. Strictly Ordered Priorities & Requirements

### Priority #1 (Top Priority): Ease & Speed of Finding, Categorizing, and Splitting Expenses
- **Instant Discovery:** View newly imported/uncategorized expenses immediately upon opening the app, with fast search and filtering across merchant name, envelope, paid-with method, and date.
- **One-Tap Categorization:** Assign or change an expense's `Envelope`, `Paid with` method, and `Tax Deductible` (`Yes`/`No`) flag with zero lag.
- **First-Class Splitting UX (with Inline Math & Tax Multiplier):**
  - Easily split a transaction into 2+ rows across multiple envelopes.
  - Support typing **inline arithmetic expressions** in split amount boxes (e.g., `1+3.23*2` or `4.23*2+12.09+3.44*6`) so receipts can be split rapidly by line item.
  - Provide a per-split-row **"Pre-Tax (Apply Tax Multiplier)"** toggle:
    - When **Pre-Tax is ON**: Writes `=${TaxMultiplier}*(4.23*2+12.09+3.44*6)` to the child row's `Cost` cell in the sheet.
    - When **Pre-Tax is OFF**: Writes `=1+3.23*2` (**no** outer parentheses) if an expression is used, or the raw number if a plain number is entered.
    - Updates the parent row's `Cost` cell using the exact existing R1C1 relative formula structure: `=${originalCost}-R[1]C[0]-R[2]C[0]...` and appends `" | "` to the parent and child `Name` cells so the sheet's native `splitTransaction` logic remains 100% compatible.
- **Fast Manual Receipt Sync:** A refresh button in the PWA that pulls new Gmail receipts in 1–3 seconds.

### Priority #2 (Second Priority): Budget Categories ("Envelopes") & Instant Balance Visibility
- One-click launch from mobile or desktop home screen (installed PWA) showing up-to-date money in each envelope in **<50ms** from local cache (`stale-while-revalidate`).
- Display the **4 most frequently used envelope columns** clearly without friction:
  1. **`Remain`** (Most important number at a glance — remaining balance for the current 2-week period)
  2. **`Remain Cumul`** (Cumulative remaining balance across periods)
  3. **`Envelope Size`** (Budgeted amount for the period)
  4. **`Spent`** (Amount spent in the period)
- Provide an expand/view toggle on the Envelope list/table to switch between the default **4-metric view** (`Envelope Size | Spent | Remain | Remain Cumul`) and the **full 11-column YTD/Projected table**.

### Priority #3 (Third Priority): Bi-Weekly Pay Periods & Big-Picture Financial Tracking
- **Two-Week Pay Periods:** Track money across 14-day pay-period tabs (`MM/DD/YY - MM/DD/YY` and `(Gap)` sheets).
- **Monthly & Annual Visibility:** Account for monthly expenditures and expose YTD/Year-Projected envelope metrics (`Envelope Size YTD`, `Spent YTD`, `Spent Average YTD`, `Envelope Size Year Projected`, `Spent Year Projected`).
- **Tax-Deductible Tracking:** Track the per-transaction `Tax Deductible` (`Yes`/`No`) column and surface period/annual tax-deductible totals.
- **Multi-Account Running Totals:** Track running balances (`Current Balance at Transaction`) and `Paid with` methods (`Credit`, `E-Check`, `Cash`, `TJ Paypal`, `TJ Venmo`, etc.).

---

## 3. Existing Codebase Architecture & Spreadsheet Schema Reference

### 3.1 Existing Build & Tooling (`gas-budget-scripts`)
- **Node & Package Manager:** Pinned via Volta (`node 24.16.0`, `pnpm 10.34.5`).
- **GAS Bundler (`scripts/build.ts`):** Uses `esbuild` to bundle `src/index.ts` into an unminified IIFE (`var _entry = (() => { ... })()`) at `dist/index.js`. It scans `src/index.ts` via regex for `export function <name>(` and `export { ... }`, appending top-level `function <name>(...args) { return _entry.<name>(...args); }` declarations at the bottom of `dist/index.js` so Google Apps Script's AST parser registers triggers (`onOpen`, `onEditInstalled`), custom formulas (`sheetNameOffset`, `getMonthOffset`), and Web App entry points (`doGet`, `doPost`).
- **Clasp Config:** `.clasp.json` points `"rootDir": "./dist"` and `scripts/build.ts` copies `appsscript.json` into `dist/`.
- **Testing & Hooks:** `vitest` (`^5.0.1`) and `tsc --noEmit` run via Husky pre-commit hook.

### 3.2 Global Variables (`Variables` Sheet — `Variables!A1:B10`)
Read by `src/util/variables.util.ts`:
| Variable Key | Type | Description |
| :--- | :--- | :--- |
| `Menu` | `string` (A1) | Top-left cell of the in-sheet mobile checkbox menu |
| `SheetVariables` | `string` (A1) | Per-sheet range where `SheetVariables` key-value pairs live on each pay-period tab |
| `PayPeriodDays` | `number` | Length of one pay period in days (e.g., `14`) |
| `TemplateName` | `string` | Name of the template tab duplicated by `addTransactionSheet()` |
| `TaxMultiplier` | `number` | Multiplier for computing tax on pre-tax split items (e.g., `1.09`) |
| `ForwardEmailAddress` | `string` | Partner's email address for forwarded receipts |
| `ForwardName` | `string` | Partner's display name prefix for `Paid with` (e.g., `<ForwardName> Venmo`) |

> **CRITICAL BUG TO FIX IN `variables.util.ts`:**
> Line 55 currently uses `SpreadsheetApp.getActiveSheet().getRange(VARIABLES_RANGE)`. In a headless Web App execution (`doGet`/`doPost`), `getActiveSheet()` may not behave reliably or may default to the first tab. Change this to `SpreadsheetApp.getActiveSpreadsheet().getRange(VARIABLES_RANGE)`.

### 3.3 Per-Sheet Variables (`SheetVariables` Range on Each Pay-Period Tab)
Existing keys in `src/util/variables.util.ts` + **4 new keys to add**:
| SheetVariable Key | Status | Description |
| :--- | :--- | :--- |
| `TransactionsStart` | Existing | Top-left cell (`Date` column, row 1) of the transactions table |
| `TransactionsMax` | Existing | Max number of transaction rows supported on the sheet |
| `SplitCheckboxesStart` | Existing | Top-left cell of the 1-column Split Checkboxes range |
| **`EnvelopesStart`** | **NEW** | Top-left data cell (`Envelope` name column, first envelope row) of the Envelope table |
| **`EnvelopesMax`** | **NEW** | Max number of envelope rows in the Envelope table |
| **`PaidWithStart`** | **NEW** | Top-left data cell of the Paid With table (left-most column holds dropdown options) |
| **`PaidWithMax`** | **NEW** | Max number of rows in the Paid With table |

### 3.4 Exact Transaction Table Column Offsets (Starting at `TransactionsStart`)
Each transaction row spans **7 contiguous columns** (`offset(0, 0, TransactionsMax, 7)`):
- **Col Offset `0` (`Date`):** Transaction Date (`Date` object or string)
- **Col Offset `1` (`Name`):** Merchant / Description string.
  - *Split convention:* Parent and child rows in a split group contain `" | "` in the name (`"<Merchant> | <Item description>"`). `doTransactionRowInfosSeemEqual` checks if dates match and if the prefix before `"|"` (or first 1/3 of the string) matches.
  - *Notes & Error styling:* Cell notes store receipt notes/errors. Background `#FF0000` indicates a receipt error; `#E8A9CA` indicates a receipt note.
- **Col Offset `2` (`Cost`):** Number or formula string.
  - *Empty cost:* Background `#E8A9CA`.
  - *Split parent formula:* `=${originalCost}-R[1]C[0]-R[2]C[0]...`
  - *Split child formula:* `=${TaxMultiplier}*(<expr>)` (if pre-tax) or `=<expr>` (if post-tax expression, no outer parens) or number.
- **Col Offset `3` (`Current Balance`):** Read-only sheet formula calculating the running balance at each transaction.
- **Col Offset `4` (`Envelope`):** Budget envelope category string (called `category` in existing GAS code; e.g., `Groceries`, `Hospitality`, `Gas`, `Work reimburse`).
- **Col Offset `5` (`Paid with`):** Payment medium string (called `type` in existing GAS code; populated from the left-most column of the Paid With table, e.g., `Credit`, `E-Check`, `Cash`, `TJ Paypal`, `TJ Venmo`).
- **Col Offset `6` (`Tax Deductible`):** `"Yes"` or `"No"` (or blank).

*(Note: In existing `budgeting.util.ts`, `getTransactionMetadataRange(range)` is `range.offset(0, 4, range.getNumRows(), 2)`, which only covered offsets `4` and `5`. We will expand or add a helper to cover offset `6` for `Tax Deductible` as well.)*

### 3.5 Exact Envelope Table Column Offsets (Starting at `EnvelopesStart`)
Each envelope row spans **11 contiguous columns** (`offset(0, 0, EnvelopesMax, 11)`):
- **Col `0`:** `Envelope` (Category name, e.g., `Groceries`, `Hospitality`, `Gas`)
- **Col `1`:** `Envelope Available Cumulative`
- **Col `2`:** `Envelope Size` *(Core 4 metric)*
- **Col `3`:** `Spent` *(Core 4 metric)*
- **Col `4`:** `Remain` *(Core 4 metric — #1 primary at-a-glance number)*
- **Col `5`:** `Remain Cumul` *(Core 4 metric)*
- **Col `6`:** `Envelope Size YTD`
- **Col `7`:** `Spent YTD`
- **Col `8`:** `Spent Average YTD`
- **Col `9`:** `Envelope Size Year Projected`
- **Col `10`:** `Spent Year Projected`

---

## 4. Key Architectural Decisions & Optimizations

### 4.1 Monorepo Structure (`pnpm` Workspaces)
Convert the repository into a `pnpm` workspace monorepo with 3 packages under `/packages`:
```text
gas-budget-scripts/
├── AGENTS.md                        # Strict coding & security guardrails for AI agents
├── PLAN.md                          # This master execution plan
├── package.json                     # Workspace root scripts (build, test, typecheck)
├── pnpm-workspace.yaml              # packages: ["packages/*"]
└── packages/
    ├── budget-shared/               # @budget/shared: Shared TS interfaces, DTOs, inline math parser/formatter
    │   ├── package.json
    │   ├── tsconfig.json
    │   └── src/
    │       ├── index.ts
    │       ├── types.ts             # EnvelopeDto, TransactionDto, PayPeriodSheetDto, ApiPayloads, ConflictDto
    │       └── math.ts              # Safe arithmetic parser + GAS formula builder (=1.09*(...) vs =1+3.23*2)
    ├── budget-scripts-gas/          # @budget/scripts-gas: Existing GAS project + new Web App API
    │   ├── package.json
    │   ├── tsconfig.json
    │   ├── appsscript.json
    │   ├── scripts/
    │   │   └── build.ts             # Existing esbuild bundler + top-level AST footer generator
    │   └── src/
    │       ├── index.ts             # Exports existing triggers/menus + NEW doGet(e) & doPost(e)
    │       ├── api.ts               # JSON API handlers, batch sheet readers, optimistic lock checks
    │       ├── models/              # Existing receipt-info, thread-info, thread-list models
    │       └── util/                # Existing budgeting, variables, spreadsheet-util, util
    └── budget-web/                  # @budget/web: Vite + React + Tailwind CSS + shadcn/ui PWA
        ├── package.json
        ├── vite.config.ts           # Configured with vite-plugin-pwa & GitHub Pages base path
        └── src/
            ├── crypto/              # IndexedDB non-extractable AES-GCM CryptoKey credential vault
            ├── sync/                # Stale-while-revalidate cache, pull-before-push queue, 3-way conflict engine
            └── components/          # Envelopes view, Expenses view, Split/Categorize drawer, Conflict modal
```

### 4.2 Fixing the 150s–360s Gmail Receipt Sync Bottleneck (<3s Execution)
- **Why `getAndRecordReceiptsAndMark` Currently Takes 150s–360s:**
  1. In `src/models/thread-list.model.ts`, when an email from Chase, PayPal, Venmo, or `ForwardEmailAddress` is **not** a receipt (e.g., statement alerts, security codes, marketing) and has no errors (`threadInfo.hasInformation() === false`), it is omitted from `this.threadInfos` and **never gets labeled `Receipts/Scripted`**. Similarly, threads with `errors.length > 0` are skipped in `recordReceipts()`.
  2. Consequently, every run of `getAndRecordReceiptsAndMark()` (`start=null, max=null`) re-downloads and re-parses every non-receipt email and errored email from the inbox since the beginning of time.
  3. `markThreadProcessed()` calls `thread.addLabel()` and `thread.moveToArchive()` inside a per-thread loop (3 sequential RPC calls per thread).
- **The Solution (To Implement in `thread-list.model.ts` & `budgeting.util.ts`):**
  1. Track **three thread arrays** on `ThreadList`:
     - `validReceiptThreads`: Threads with `receiptInfos.length > 0 && errors.length === 0`.
     - `ignoredThreads`: Threads with `receiptInfos.length === 0 && errors.length === 0` (non-receipt emails).
     - `erroredThreads`: Threads with `errors.length > 0`.
  2. When `shouldMarkProcessed === true` in `recordReceipts()`, use **batch `GmailLabel.addToThreads(threads)`** (chunked in batches of 100 if `>100` on initial clean-up):
     - Apply `Receipts/Scripted` to **all touched threads** (`[...validReceiptThreads, ...ignoredThreads, ...erroredThreads]`). Because `getTJGmailSearchQuery()` already includes `NOT label:receipts-scripted`, none of these threads will ever be re-fetched again!
     - Apply `Receipts` to `validReceiptThreads` and call `GmailApp.moveThreadsToArchive(validReceiptThreads)`.
     - Apply `Receipts/Ignored` to `ignoredThreads`.
     - Apply `Receipts/Error` to `erroredThreads`.

### 4.3 Credential Security in the Browser (Option B: Non-Extractable `CryptoKey` Vault)
- **Zero Hardcoded Secrets:** Neither the GAS Web App Deployment URL nor the `API_KEY` is ever committed to git, `.env`, or source files.
- **No Plaintext Storage:** Credentials are never stored in plaintext `localStorage`.
- **Implementation (`packages/budget-web/src/crypto/credential-vault.ts`):**
  1. Generate a 256-bit `AES-GCM` `CryptoKey` via `window.crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"])` with `extractable: false`.
  2. Persist the opaque `CryptoKey` handle inside `IndexedDB` (encrypted at rest by the browser via OS keychain: Windows DPAPI, macOS/iOS Keychain, Android Keystore).
  3. Encrypt `{ gasWebAppUrl, apiKey }` using `crypto.subtle.encrypt` with a random 12-byte IV and store only `{ iv, ciphertext }` in `IndexedDB`.
  4. Decrypt only in volatile local function scope immediately prior to calling `fetch()`, requiring zero PIN entry on startup.

### 4.4 Background Sync Indicator & Pull-Before-Push Conflict Resolution
- **Persistent Sync Status Indicator:** Header badge showing `Synced`, `Pulling...`, `Syncing N changes...`, `Offline (N queued)`, or `Conflict Detected`.
- **No Silent Overwrites (Pull-Before-Push Queue):**
  1. Every cached transaction row carries a deterministic `rowFingerprint` (hash of `date | name | costFormulaOrValue | envelope | paidWith | taxDeductible`) representing the state when it was last pulled (`baseState`).
  2. When a user categorizes or splits an expense in the UI, it applies optimistically to the UI state and enters `pendingMutations` with `{ rowIndex, sheetName, baseState, proposedState }`.
  3. When flushing `pendingMutations`:
     - The client/server fetches the latest state from the sheet first (inside `LockService.getScriptLock()`).
     - If the target row's current state on the sheet **no longer matches `baseState`**, that local mutation is **not pushed** and is moved to a `conflicts` list.
     - The UI updates to reflect the newly pulled remote data and displays a **Conflict Warning Modal/Banner** showing all 3 states clearly:
       - **Original (`What it was before`):** e.g., `Envelope: [Uncategorized], Cost: $45.00`
       - **Discarded Local Change (`What you changed it to`):** e.g., `Envelope: Groceries`
       - **Current Sheet Value (`What it is now`):** e.g., `Envelope: Hospitality`
     - The user must explicitly click either **"Accept Remote Changes"** or **"Overwrite with My Changes"**.

---

## 5. Step-by-Step Prototype Implementation Plan

### Step 1: Monorepo Migration & `@budget/shared` Package
1. Create `pnpm-workspace.yaml` and move the existing GAS files (`src/`, `scripts/`, `appsscript.json`) into `packages/budget-scripts-gas/` while keeping git hooks and `pnpm run build` / `pnpm run push` / `pnpm test` working seamlessly from root.
2. Create `packages/budget-shared` containing:
   - DTO interfaces (`EnvelopeRowDto`, `TransactionRowDto`, `PayPeriodSummaryDto`, `GetSheetDataResponse`, `PostActionRequest`, `ConflictDetailDto`).
   - `evaluateMathExpression(expr: string): number | null` — safe parser supporting `+`, `-`, `*`, `/`, `(`, `)`, and decimals.
   - `formatSplitCostFormula(expr: string, applyTax: boolean, taxMultiplier: number): string | number` — returns `=${taxMultiplier}*(${cleanExpr})` when `applyTax` is true, or `=${cleanExpr}` (no outer parens) when `applyTax` is false and `cleanExpr` contains operators, or a `number` if it's a single numeric literal.
   - Unit tests in `vitest` verifying formula formatting and math evaluation.

### Step 2: Sheet Variable Additions, Gmail Optimization & GAS JSON API (`packages/budget-scripts-gas`)
1. **Update `src/util/variables.util.ts`:**
   - Fix `getVariablesRange()` to use `SpreadsheetApp.getActiveSpreadsheet().getRange(VARIABLES_RANGE)`.
   - Add `EnvelopesStart`, `EnvelopesMax`, `PaidWithStart`, `PaidWithMax` to `SHEET_VARIABLE_NAMES` and `SheetVariables`.
2. **Optimize Gmail Labeling in `src/models/thread-list.model.ts` & `src/util/budgeting.util.ts`:**
   - Track `ignoredThreads` and `erroredThreads` alongside valid receipt threads.
   - Batch-apply `Receipts/Scripted` to all touched threads, `Receipts` + archive to valid receipt threads, `Receipts/Ignored` to non-receipt threads, and `Receipts/Error` to errored threads.
3. **Create `src/api.ts` and export `doGet`, `doPost`, and `setupGenerateApiKey` from `src/index.ts`:**
   - **Authentication:** Validate `API_KEY` from `PropertiesService.getScriptProperties()`.
   - **Fast Batch Reader (`getPayPeriodState(sheetName?)`):**
     - Reads the active/latest pay-period sheet (or requested `sheetName`) + list of available pay-period tab names (`getTransactionSheetInfos()`).
     - Reads `TransactionsStart` (`TransactionsMax` x 7 columns: values, formulas, notes, backgrounds) in one batch.
     - Reads `EnvelopesStart` (`EnvelopesMax` x 11 columns: values) in one batch.
     - Reads `PaidWithStart` (`PaidWithMax` x 1 column: values, filtering out `"Tax deductible"` stats row from dropdown choices) in one batch.
   - **Mutation Handlers in `doPost(e)` (wrapped in `LockService.getScriptLock()`):**
     - `action: "updateTransaction"`: Validates `baseFingerprint` against current row. If matched, updates `Envelope`, `Paid with`, `Tax Deductible`, `Name`, or `Cost` and returns updated sheet state. If mismatched, returns `{ status: "conflict", currentRow, ... }`.
     - `action: "splitTransaction"`: Validates `baseFingerprint` of parent row. Shifts subsequent rows down in a single batch `copyTo` (supporting splitting into $N$ rows at once), writes the parent formula `=${parentCost}-R[1]C[0]...` and `" | "` name suffix, writes all child rows with their pre-tax/post-tax formulas and categories, and returns the updated sheet state.
     - `action: "syncReceipts"`: Calls `Budgeting.getAndRecordReceipts(0, 25, true)` and returns newly recorded counts + refreshed sheet state.

### Step 3: Frontend PWA (`packages/budget-web`)
1. Scaffold Vite + React + TypeScript + Tailwind CSS + `shadcn/ui` + `vite-plugin-pwa`.
2. Implement `credential-vault.ts` (IndexedDB `extractable: false` `AES-GCM` `CryptoKey`) and a **Connection Settings Dialog** where the user pastes their GAS Web App URL and `API_KEY` once.
3. Implement the **Sync Engine & Conflict Manager** (`useBudgetSync` hook):
   - Renders immediately (<50ms) from local IndexedDB/localStorage snapshot.
   - Displays the top **Sync Status Pill** (`Synced`, `Pulling...`, `Syncing...`, `Conflict`).
   - Implements the Pull-Before-Push queue and the **Conflict Resolution Modal** (showing Before, Your Discarded Change, and Current Remote Value with "Accept Remote" vs. "Overwrite with Mine" buttons).
4. Build the **Envelopes View**:
   - Pay-period selector dropdown at top (defaults to current pay-period sheet).
   - Default **4-Column Compact View** showing `Envelope`, `Envelope Size`, `Spent`, `Remain` (highlighted primary badge), and `Remain Cumul` for every envelope.
   - **"Expand All Columns" Toggle** that reveals the full 11-column table including YTD and Year Projected stats.
5. Build the **Expenses / Inbox View & Split Drawer**:
   - Filter chips (`Uncategorized`, `All`, `Tax Deductible`, `Errors/Notes`) + instant search input.
   - Tapping any transaction opens the **Categorize & Split Drawer**:
     - Quick-select chips/dropdowns for `Envelope`, `Paid with`, and `Tax Deductible` (`Yes`/`No`).
     - **"Split Expense" Mode:** Dynamic list of split items. Each split item has an item label (appended after `" | "`), an **Inline Math Expression input** (`4.23*2+12.09`), a **Pre-Tax (`1.09x`) toggle**, and an `Envelope` selector, plus a live **"Remaining to Allocate"** preview calculated from the parent receipt total.
6. Add GitHub Actions workflow (`.github/workflows/deploy-web.yml`) to build `packages/budget-web` and deploy to GitHub Pages.

---

## 6. Post-Prototype Production Roadmap

1. **Multi-Sheet Big Picture & Paid-With Stats:**
   - Surface the `Paid with` summary table stats (including the pay-period `Tax Deductible` stats row) in an Accounts & Tax tab.
   - Add "Add New Pay Period Sheet (`addTransactionSheets`)" button directly in the PWA.
   - Add manual transaction entry (for cash/check expenses without email receipts).
2. **Enhanced Line-Item Split Builder & OAuth:**
   - Visual receipt line-item builder that constructs `=1.09*(item1+item2)` or `=item1+item2` from individual item chips, plus optional Google OAuth / WebAuthn unlock.