# gas-budget-scripts

Google Apps Scripts for budgeting.

---

## Summary

This repository contains a Google Apps Script TypeScript project for managing and recording transaction receipts into Google Sheets with a [container-bound Apps Script project](https://developers.google.com/apps-script/guides/bound).

Code is written using standard modern TypeScript (`import` / `export`) in `src/`, bundled using `esbuild` into human-readable, unminified JavaScript in `dist/`, and synchronized to Google Apps Script via `@google/clasp`.

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
   pnpm exec clasp clone SCRIPT_PROJECT_ID --rootDir ./dist
   ```
   > **IMPORTANT**: Your `.clasp.json` (as well as any environment variants like `.clasp - prod.json` and `.clasp - test.json`) **must** set:
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

This runs `scripts/build.ts`, which:

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

### Type Checking & Testing

```bash
pnpm run typecheck    # Runs tsc --noEmit
pnpm test             # Runs vitest unit tests
pnpm run test:watch   # Runs vitest in watch mode
```

> **Git Hook**: Husky runs `pnpm run typecheck` and `pnpm run test` automatically as a pre-commit hook before every commit.

---

## Branches

- `main` - Active development. Use with a test budget sheet.
- `deploy-budget-sheet` - Production code. Use with the real budget sheet.
