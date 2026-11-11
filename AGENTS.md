# AGENTS.md — Coding, Security & Architecture Guidelines

This repository is a `pnpm` monorepo containing a Google Apps Script backend (`packages/budget-scripts-gas`), a shared TypeScript library (`packages/budget-shared`), and a React/Vite Progressive Web App frontend (`packages/budget-web`).

All AI coding agents and contributors MUST adhere strictly to the rules below.

---

## 1. CRITICAL SECURITY RULES (ZERO TOLERANCE)

1. **NEVER Hardcode Secrets, URLs, or Personal Identifiers:**
   - **NEVER** commit, hardcode, or place Google Apps Script Web App Deployment URLs (`https://script.google.com/macros/s/...`), API keys, tokens, Spreadsheet IDs, or personal email addresses in any source file, test file, configuration file, `.env` file committed to git, or documentation.
   - In `packages/budget-scripts-gas`, all secrets (such as `API_KEY`) MUST be read dynamically at runtime from `PropertiesService.getScriptProperties()`.
   - In `packages/budget-web`, the GAS Deployment URL and `API_KEY` MUST be provided by the user at runtime via the UI Settings dialog.
2. **NEVER Store Plaintext Credentials in `localStorage`:**
   - In `packages/budget-web`, sensitive connection credentials (`gasWebAppUrl` and `apiKey`) MUST be encrypted using the Web Crypto API (`window.crypto.subtle`) with a non-extractable (`extractable: false`) `AES-GCM` 256-bit `CryptoKey` stored in `IndexedDB`.
   - Only encrypted ciphertext (`iv` + `ciphertext`) may be persisted to disk. Decryption must happen strictly in-memory inside the request execution scope.

---

## 2. DATA INTEGRITY & CONFLICT RESOLUTION RULES

1. **Prevent Silent Data Overwrites:**
   - Never blindly overwrite a Google Sheet row from the web app without verifying that the row on the sheet still matches the `baseFingerprint` (the state of the row when the user last pulled).
   - Always pull/verify remote changes before applying queued local mutations.
   - If a row was modified externally since the client's last pull, reject the automatic push, surface a clear 3-way conflict warning (`Original Value`, `Discarded Local Change`, `Current Remote Value`), and require explicit user confirmation (`Accept Remote` vs. `Overwrite with Mine`).
2. **Preserve 100% Spreadsheet Formula Compatibility:**
   - When splitting transactions from the web app, always write formulas that match the spreadsheet's native conventions:
     - **Parent cost cell:** `=${originalCost}-R[1]C[0]-R[2]C[0]...`
     - **Child cost cell (Pre-Tax toggle ON):** `=${TaxMultiplier}*(${expression})`
     - **Child cost cell (Pre-Tax toggle OFF):** `=${expression}` (do NOT wrap in unnecessary outer parentheses, e.g., `=1+3.23*2`), or a raw number if no math operators are present.
     - **Split Name convention:** Ensure split group parent and child names retain the `" | "` separator (`"<Merchant> | <Item Note>"`) so `doTransactionRowInfosSeemEqual` continues to recognize split groups.

---

## 3. MONOREPO, GMAIL & BUILD CONVENTIONS

1. **Tooling & Package Manager:**
   - Always use `pnpm` (managed via Volta: Node `24.16.0`, `pnpm 10.34.5`). Do not use `npm` or `yarn`.
   - Keep shared TypeScript types, DTOs, and pure business utilities (such as the inline math expression parser and formula formatter) in `packages/budget-shared` (`@budget/shared`) to maintain DRY code across frontend and backend.
2. **Google Apps Script Build Pipeline (`packages/budget-scripts-gas`):**
   - Entry point is `src/index.ts`. Any function that must be callable by Google Apps Script triggers, menus, custom sheet formulas, or HTTP Web App requests (`doGet`, `doPost`) MUST be exported from `src/index.ts` so `scripts/build.ts` generates a top-level function declaration in `dist/index.js`.
   - Keep `minify: false` in `scripts/build.ts` so deployed code remains human-readable in the online Apps Script editor.
   - In `src/util/variables.util.ts` (and any code invoked by `doGet`/`doPost`), use `SpreadsheetApp.getActiveSpreadsheet()` rather than `SpreadsheetApp.getActiveSheet()` when resolving global ranges like `Variables!A1:B10`.
3. **Google Apps Script & Gmail Performance Best Practices:**
   - Never call `range.getValues()`, `range.getFormulas()`, or `range.setValues()` inside a row-by-row loop. Always read and write contiguous ranges in a single batch call.
   - When marking Gmail threads processed in `budgeting.util.ts`, always use batch `GmailLabel.addToThreads(threads)` and `GmailApp.moveThreadsToArchive(threads)` rather than per-thread loops. Ensure `Receipts/Scripted` is applied to all touched threads (`valid`, `Receipts/Ignored`, and `Receipts/Error`) so non-receipt and errored threads are never re-scanned on future runs.
   - When sending `POST` requests from `packages/budget-web` to the GAS `doPost` endpoint, use `Content-Type: text/plain;charset=utf-8` with `JSON.stringify(payload)` and `redirect: "follow"` to avoid CORS preflight (`OPTIONS`) failures on `script.google.com`.
4. **Verification Before Committing:**
   - Always run `pnpm run typecheck` and `pnpm test` across all packages before completing a task.