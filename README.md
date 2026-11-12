# gas-budget-scripts

Google Apps Scripts for budgeting.

---

## Summary

This repository contains a Google Apps Script TypeScript project for managing and recording transaction receipts into Google Sheets with a [container-bound Apps Script project](https://developers.google.com/apps-script/guides/bound).

This is a `pnpm` workspace. The GAS project lives in `packages/budget-scripts-gas` (`@budget/scripts-gas`) and shared pure TypeScript utilities (DTOs, math parser, cost-formula builder, row fingerprint) live in `packages/budget-shared` (`@budget/shared`). The PWA frontend lives in `packages/budget-web` (`@budget/web`, React + Vite) and is deployed to GitHub Pages by `.github/workflows/deploy-web.yml`; see [`web-plan.md`](web-plan.md).

GAS code is written using standard modern TypeScript (`import` / `export`) in `packages/budget-scripts-gas/src/`, bundled using `esbuild` into human-readable, unminified JavaScript in `dist/`, and synchronized to Google Apps Script via `@google/clasp`.

---

## Prerequisites & Tooling

We use **[Volta](https://volta.sh)** to pin and manage compatible Node and `pnpm` versions.

1. **Install Volta** (if not already installed): follow instructions at [volta.sh](https://volta.sh).
2. **Install pnpm with Volta**:
   ```bash
   volta install pnpm@10
   ```
   Volta will automatically detect the pinned versions in `package.json` (`node 24.16.0` and `pnpm 10.34.5`).

---

## Setup

1. **Install dependencies**:

   ```bash
   pnpm install
   ```

   > **Note on esbuild**: In `pnpm` v10+, dependency build scripts are restricted by default. If `pnpm` prompts you about `esbuild`, run:
   >
   > ```bash
   > pnpm rebuild esbuild
   > ```

2. **Log into Google via clasp**:

   ```bash
   pnpm run login
   ```

3. **Configure clasp connection**:
   Clone or connect your Apps Script project:
   ```bash
   pnpm run connect   # runs `clasp clone SCRIPT_PROJECT_ID --rootDir ./dist` inside packages/budget-scripts-gas
   ```
   > **IMPORTANT**: Your `.clasp.json` lives in `packages/budget-scripts-gas/` (it is gitignored; if you had one in the repo root before the workspace move, move it there along with any variants). It (as well as any environment variants like `.clasp - prod.json` and `.clasp - test.json`) **must** set:
   >
   > ```json
   > "rootDir": "./dist"
   > ```
   >
   > Clasp pushes the bundled files from `dist/` rather than the raw TypeScript files in `src/`.

---

## Development Workflow

### Building

```bash
pnpm run build
```

This runs every package's `build` script. For the GAS package that is `packages/budget-scripts-gas/scripts/build.ts` (output in `packages/budget-scripts-gas/dist/`), which:

- Dynamically discovers all exported functions in `src/index.ts` (triggers, menu handlers, and custom spreadsheet formulas like `=sheetNameOffset()`).
- Bundles everything into an unminified, human-readable `dist/index.js` with top-level function declarations for the Apps Script runtime.
- Copies `appsscript.json` into `dist/`.

### Watching & Live Sync

```bash
pnpm start
```

Starts `esbuild` in watch mode on `src/` and spawns `clasp push --watch` on `dist/`. Any change to source files is immediately rebuilt and uploaded to Google Apps Script.

### Manual Push

```bash
pnpm run push
```

Builds the latest bundles and runs `clasp push`.

### Web app (PWA)

```bash
pnpm --filter @budget/web run dev       # dev server (served under /gas-budget-scripts/)
pnpm --filter @budget/web run build     # tsc + vite build -> packages/budget-web/dist (set PAGES_BASE to change the base path)
pnpm --filter @budget/web run preview   # serve the production build
```

The web app asks for the Apps Script web app URL and API key in its Settings dialog and stores them encrypted on the device; nothing secret is in the bundle. `pnpm install` (and `build`) generate `packages/budget-web/public/licenses.html` (open source licenses, gitignored). To publish: enable **Settings -> Pages -> Source: GitHub Actions** once, also allow the production branch under **Settings -> Environments -> github-pages -> Deployment branches** (the default policy only allows `main`). The site deploys from `deploy-budget-sheet`, the production branch used with the real sheet: merge `main` into `deploy-budget-sheet` (a push that touches the web or shared packages), or run "Deploy web app to GitHub Pages" manually on that branch. Pushes to `main` do not deploy.

### Local demo (fake backend)

```bash
pnpm run demo
```

Starts the **demo backend** and the Vite dev server together. The backend is the real `doPost`/`doGet` from `packages/budget-scripts-gas/src` running on the in-memory GAS mocks, seeded with obviously fake data: a Variables sheet, a Template, three pay-period tabs (the oldest has no optional `EnvelopesStart`/`PaidWithStart` keys), ~27 transactions including uncategorized ones, an existing split group, envelope balances, Paid With options, and fake Gmail receipts (Chase/Venmo/PayPal) that "Import receipts" adds. Nothing touches Google. Then:

1. Open <http://127.0.0.1:5173/>.
2. In Settings paste the web app URL `http://127.0.0.1:8787/macros/s/demo/exec` and the API key `dev-key` (both are printed when the backend starts). Local `http://localhost`/`127.0.0.1` URLs are accepted **only in dev builds** (`import.meta.env.DEV`); production builds accept only `script.google.com` deployment URLs.

State lives in memory; restart the demo to reset it. Limits: the mock evaluates arithmetic and relative (`R[1]C[0]`) cell formulas on read, so costs and Current Balance stay live, but the envelope table (Spent/Remain/...) holds plain numbers and does **not** update when transactions change. Run the backend alone with `pnpm --filter @budget/scripts-gas run dev:server` (`DEMO_PORT` changes the port, default 8787). The server also exposes test-only `/__debug/*` endpoints (see AGENTS.md section 5).

### End-to-end tests

```bash
pnpm exec playwright install chromium   # once (from packages/budget-web, or add --with-deps on Linux)
pnpm run test:e2e
```

Playwright (Chromium) starts the demo backend (port 8788) and the Vite dev server (port 5174), resets the backend before every test, and drives the real UI: settings, loading, categorizing, cost math, splitting (add tax on/off), all pay periods, details, conflicts, offline queue, receipt import, and tabs without optional keys. Assertions on the sheet read the mock sheet through `/__debug/sheet`. The report is written to `packages/budget-web/playwright-report/`. `pnpm test` stays unit-only and fast.

### Open Apps Script Project

```bash
pnpm run open
```

Opens the connected project in your browser at `script.google.com`.

### Type Checking, Linting, Formatting & Testing

```bash
pnpm run typecheck      # tsc (TypeScript 7) in every package (src, then src + test + configs)
pnpm run lint           # ESLint (typescript-eslint); `pnpm run lint:fix` to auto-fix
pnpm run format         # Prettier --write; `pnpm run format:check` to verify only
pnpm test               # Vitest unit tests in every package (packages/*/test/**/*.test.ts)
pnpm run test:watch     # Vitest in watch mode
pnpm run test:coverage  # Vitest with v8 coverage (HTML report in packages/*/coverage/)
```

Tests run against in-memory mocks of the Apps Script globals (`SpreadsheetApp`, `GmailApp`, `Logger`) in `packages/budget-scripts-gas/test/gas-mocks/`. They are characterization tests: they pin the current spreadsheet behavior (cell values, formulas, labels) so refactors can be verified. Mocks do not evaluate formulas (unless the demo server opts in) or adjust relative references on `copyTo`. Sample emails with their expected parse results (`test/fixtures/receipts.ts`) and sample sheet data (`test/fixtures/sheets.ts`) are shared by the unit tests and the demo backend: change an email format in one place.

> **TypeScript versions**: `typescript` is aliased to `@typescript/typescript6` because typescript-eslint needs the TS 6 API, while `@typescript/native` (TypeScript 7) provides the `tsc` used by `typecheck`. Revisit once typescript-eslint supports TS 7.

> **Git Hook**: Husky runs `format:check`, `lint`, `typecheck` and `test` as a pre-commit hook.

> **CI**: `.github/workflows/ci.yml` runs format check, lint, typecheck, tests with coverage and build on pushes to `main` and on pull requests, plus a separate `e2e` job (Playwright, Chromium) that uploads the Playwright report as an artifact when it fails.

---

## Branches

- `main` - Active development. Use with a test budget sheet.
- `deploy-budget-sheet` - Production code. Use with the real budget sheet.
