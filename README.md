# gas-budget-scripts

Google Apps Scripts for budgeting.

---

## Summary

This repository contains a Google Apps Script TypeScript project for managing and recording transaction receipts into Google Sheets with a [container-bound Apps Script project](https://developers.google.com/apps-script/guides/bound).

This is a `pnpm` workspace. The GAS project lives in `packages/budget-scripts-gas` (`@budget/scripts-gas`) and shared pure TypeScript utilities (DTOs, math parser, cost-formula builder, row fingerprint) live in `packages/budget-shared` (`@budget/shared`). A PWA frontend (`packages/budget-web`) is planned in [`web-plan.md`](web-plan.md).

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

Tests run against in-memory mocks of the Apps Script globals (`SpreadsheetApp`, `GmailApp`, `Logger`) in `packages/budget-scripts-gas/test/gas-mocks/`. They are characterization tests: they pin the current spreadsheet behavior (cell values, formulas, labels) so refactors can be verified. Mocks do not evaluate formulas or adjust relative references on `copyTo`.

> **TypeScript versions**: `typescript` is aliased to `@typescript/typescript6` because typescript-eslint needs the TS 6 API, while `@typescript/native` (TypeScript 7) provides the `tsc` used by `typecheck`. Revisit once typescript-eslint supports TS 7.

> **Git Hook**: Husky runs `format:check`, `lint`, `typecheck` and `test` as a pre-commit hook.

> **CI**: `.github/workflows/ci.yml` runs format check, lint, typecheck, tests with coverage and build on pushes to `main` and on pull requests.

---

## Branches

- `main` - Active development. Use with a test budget sheet.
- `deploy-budget-sheet` - Production code. Use with the real budget sheet.
