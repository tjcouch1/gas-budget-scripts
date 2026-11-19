# Receipt rules review (branch `receipt-rules` vs `frontend-plan`)

Scope: `git diff frontend-plan` plus the untracked files, read against AGENTS.md (binding), `receipt-rules-plan.md` (spec; section 11 binding) and the old parser (`git show frontend-plan:packages/budget-scripts-gas/src/models/thread-list.model.ts`). Every finding below was verified by reading the code; where noted, by running a Node snippet. No code was changed.

Paths: `gas/` = `packages/budget-scripts-gas/`, `shared/` = `packages/budget-shared/`, `web/` = `packages/budget-web/`.

Summary: 0 blockers, 1 major, 6 minor, 6 nits. Parity with the old parser holds for every branch I traced (details in "Verified OK"). The security rules hold (no secrets, no formula writes, no `dangerouslySetInnerHTML`/`eval`, query builder injection-proof, read-only fetch never labels/archives, candidate regexes never executed server-side). The main gaps are error handling for a malformed `ForwardEmailAddress` variable, a limit mismatch that can make the fetcher unusable, stale docs, and a few UI state/a11y details.

---

## Findings (most severe first)

### 1. [major] A malformed `ForwardEmailAddress` stops every import with a misleading message and the wrong API code

- `shared/src/receipt-query.ts:53-54` throws `Error("Invalid sender address: <x>")` for the forward address too.
- `gas/src/util/receipt-rules.util.ts:60-65` (`getReceiptSearchQuery`) does not wrap that error.
- `gas/src/api/receipt-rules-handlers.ts:52-60` (`mapInvalidConfig`) only recognizes the `ReceiptRules variable is invalid` prefix; `gas/src/api/handlers.ts:563-575` (`handleSyncReceipts`) therefore lets it fall through to `server_error`. `handleFetchReceiptEmails` (`receipt-rules-handlers.ts:147-162`) maps the same condition to `invalid_config` with a different text (`ForwardEmailAddress variable is invalid: ...`), and `handleGetReceiptRules` (`:213-219`) swallows it (`query: null`, `issues: []`).

Problem: before this branch the Gmail query was built by string interpolation, so a value such as `Sam <sam@example.com>` or `sam@example.com ` (trailing space, as Sheets happily stores) only broke forwarded-email detection; Chase/PayPal/Venmo imports still ran. Now `STRICT_EMAIL_REGEX` rejects it and `getChaseReceipts` throws before `GmailApp.search`. Failing loudly is the plan's decision (section 9), but:

- the menu/trigger alert says `Invalid sender address: sam@example.com ` (the owner looks for a sender, not for `ForwardEmailAddress`);
- the web app shows a generic `server_error` notice with no hint, because `syncReceipts` does not map it to `invalid_config`;
- the Customize variables dialog loads with `query: null` and no explanation; only the fetcher's query preview shows "Invalid sender address".

Scenario: owner's sheet has `ForwardEmailAddress = "Sam <sam@example.com>"` (or a trailing space). After deploying, the time-driven import fails every run with `Invalid sender address: Sam <sam@example.com>`; "Sync receipts" in the PWA shows a red generic error; nothing points at the Variables sheet.

Fix (red -> green):

1. `gas/test/receipt-rules-util.test.ts`: add a test seeding `ForwardEmailAddress` as `"Sam <sam@example.com>"` and assert `getReceiptSearchQuery()` throws `/^ForwardEmailAddress variable is invalid: .*Fix it on the Variables sheet/`.
2. `gas/test/api-receipt-rules.test.ts`: with the same seed, `syncReceipts` and `fetchReceiptEmails` (no `senders`) both return `error.code === "invalid_config"` with that message; `getReceiptRules` returns `query: null` (unchanged) — optionally add a `queryError: string | null` to `ReceiptRulesStateDto` (`shared/src/types.ts`) and show it in `CustomizeVariablesDialog` above the tabs (only if the same agent owns both files; see groups).
3. `gas/src/util/receipt-rules.util.ts`: in `getReceiptSearchQuery`, catch the builder error and rethrow `Error("ForwardEmailAddress variable is invalid: <original message>. Fix it on the Variables sheet (a plain address like name@example.com).")`. Export the prefix constant.
4. `gas/src/api/receipt-rules-handlers.ts`: `mapInvalidConfig` treats both prefixes as `invalid_config`; remove the ad-hoc try/catch in `handleFetchReceiptEmails` (it becomes redundant).
5. Keep `SyncEngine`'s hint `"variables"` for `invalid_config`; the dialog already shows `ForwardEmailAddress` read-only with "Edit these two in the sheet's Variables tab", so the hint is still useful.

### 2. [minor] Fetch sender cap (20) is lower than the configurable sender count (25): the fetcher is unusable for 21-25 senders

- `shared/src/receipt-rules.ts:544` allows 1-25 senders.
- `gas/src/api/receipt-rules-handlers.ts:41,87-102` rejects `senders` arrays longer than `FETCH_MAX_SENDERS = 20`.
- `web/src/components/variables/EmailFetcher.tsx:15,54-60` always sends every candidate sender (`senders: allSenders`) and shows "At most 20 sender addresses in total" when `senders.length + extras.length > 20`, which disables Fetch.

Scenario: the owner adds a 21st sender in the form. The rules validate and save, but "Fetch emails" is disabled with "At most 20 sender addresses in total" and there is no way to preview anything (the fetcher has no option to send a subset).

Fix: define one constant in `shared/src/receipt-query.ts`, e.g. `export const MAX_QUERY_SENDERS = 30;` (25 configurable + 5 extras), use it in `validateSenders` and in `EmailFetcher` (replace the local `MAX_SENDERS`), and update the `badRequest` text to use the constant. Red tests: `api-receipt-rules.test.ts` "rejects %j" case with 31 senders (and accepts 25 defaults-plus-extras), `EmailFetcher.test.tsx` renders 25 senders plus one extra and expects Fetch enabled.

### 3. [minor] Stale documentation after the range change

- `AGENTS.md:23` still says `src/util/variables.util.ts  #   Variables!A1:B10 + per-sheet SheetVariables readers`.
- `web-plan.md:18`, `web-plan.md:116`, `web-plan.md:118` still say `Variables!A1:B10` and "3 rows remain free" (the banner added at line 61 contradicts section 3.2 below it).
- `AGENTS.md` section 0 does not list `gas/src/util/a1.util.ts` (new shared A1 helpers imported by both setup utilities).
- `receipt-rules-plan.md:24` and section 9 call out only the sender-address case-insensitivity, but `applyReceiptRules` (`shared/src/receipt-rules.ts:845-852`) also compares the `To` address with `ForwardEmailAddress` case-insensitively (old code: exact `===`). This is a second intentional micro-change (pinned by the test "uses ForwardName when sent To the forward address (case-insensitive)"); it should be stated in the plan/README so nobody "fixes" it later.

Fix: edit the four doc locations (text only). No tests.

### 4. [minor] Save error text stays on screen after the user fixes the rules

- `web/src/components/variables/CustomizeVariablesDialog.tsx:150-154` (`edit`) only resets `save` when it is `"saved"`; `"error"` and `"conflict"` persist (`:277-302`).

Scenario: Save fails with `bad_request` ("ReceiptRules is invalid: senders[0].address: ...", listed under the buttons). The user fixes the address in the form; the red server message and its issue list stay visible next to the now-enabled Save button, contradicting the live validation ("0 problems") until the next save completes.

Fix: in `edit`, reset any non-idle, non-saving state: `if (save.kind !== "idle" && save.kind !== "saving") setSave({ kind: "idle" })`. Keep `"conflict"` until reload? No: a conflict is about `expectedRaw`, which the next save will re-check, so clearing it on edit is fine, but keep the Reload button reachable by also offering it when `save.kind === "idle" && loaded.issues.length > 0`? Not needed; simplest is to clear on edit. Red test in `CustomizeVariablesDialog.test.tsx`: after the "lists server validation issues from a rejected save" flow, `fireEvent.change` on "Default type prefix" and assert `screen.queryByText(/server says no/)` is null.

### 5. [minor] `getCompiledReceiptRules` memo ignores `ForwardEmailAddress`

- `gas/src/util/receipt-rules.util.ts:18,36-52` caches by `raw` only, but validation (`compileReceiptRules(..., { forwardEmailAddress })`) depends on `ForwardEmailAddress`.

Scenario (demo server/long-lived process, or any future code path that changes the variable within an execution): `bustVariablesCache()` is called, `ForwardEmailAddress` now equals a configured sender, but because `raw` is unchanged the stale compiled object is returned and the "sender must not be the ForwardEmailAddress" rule is not enforced. On real GAS each execution starts fresh, so today this only affects the demo backend and tests, but the comment claims "bustVariablesCache makes getVariables re-read, which recompiles" which is only true when the JSON text changed.

Fix: key the cache on both values: `cache = { raw, forwardEmailAddress, compiled }` and compare both. Red test in `receipt-rules-util.test.ts`: compile, change `B6` (ForwardEmailAddress) to `no.reply.alerts@chase.com`, `bustVariablesCache()`, expect `getCompiledReceiptRules()` to throw the invalid-config message for `senders[0].address`.

### 6. [minor] Preview and import can diverge on long bodies without an explicit warning in the result

- `gas/src/api/receipt-rules-handlers.ts:39-40,175-178` truncates `plainBody` at 20,000 and `htmlBody` at 60,000 characters; the import (`ThreadList`) runs the same regexes on the full body.
- `web/src/components/variables/EmailResults.tsx:151-157` only shows "(truncated by the server)" inside the collapsed `<details>` summary; the outcome badge and captures say nothing.

Scenario: a body rule whose match sits after the 20,000th character shows "Not a receipt" in the preview while the real import records a receipt (or vice versa). The owner "fixes" a rule that was already working.

Fix: in `EmailResults.tsx`, when `email.plainBodyTruncated || email.htmlBodyTruncated`, render a visible amber line next to the badge: "Body truncated by the server; the import sees the full text, so this result may differ." Red test in `EmailResults.test.tsx` ("notes truncated bodies" can be tightened to assert the sentence is outside the `<details>`). No server change.

### 7. [minor] Test gaps worth closing

- Note text keeps the original case of the sender address (`gas/src/models/thread-list.model.ts:119` uses `extractEmailAddress(message.getFrom())` on purpose), but `gas/test/receipts.test.ts:163` only pins a lower-case address, so a refactor to `outcome.senderAddress` would pass. Add a non-receipt fixture with `from: "Chase <No.Reply.Alerts@Chase.com>"` and assert the note contains `From: No.Reply.Alerts@Chase.com`.
- No API test for `fetchReceiptEmails`/`syncReceipts` with an invalid `ForwardEmailAddress` (covered by finding 1's red tests).
- No API test that `saveReceiptRules` rejects an oversize config (validation is unit-tested in `shared/test/receipt-rules-validate.test.ts` "oversize JSON", but the handler path that turns it into `bad_request` with `issues[0].path === "(root)"` is not). Add one `it` in `api-receipt-rules.test.ts` using a 41,000-character `category`-padded config.
- No API test that `saveReceiptRules` returns a readable error when `Variables!A1:B30` has no free row (`ensureGlobalVariableRow` throws; the router maps it to `server_error` with the message). Add one `it` reusing the 23 filler rows from `global-variables-setup.test.ts`.
- e2e: the plan's C6 "paste invalid JSON in the JSON tab" flow is unit-tested only (`RulesEditor.test.tsx`); acceptable, but note it in the plan.

### 8. [nit] Tabs are not keyboard-navigable as a tablist

- `web/src/components/variables/CustomizeVariablesDialog.tsx:209-232` and `RulesEditor.tsx:154-173` render `role="tablist"`/`role="tab"` buttons without `tabIndex` management or arrow-key handling, and on `lg` screens the tablist is `hidden` while the `tabpanel`s keep `aria-labelledby` pointing at hidden tabs.

Fix: add `onKeyDown` handling ArrowLeft/ArrowRight/Home/End that moves `section` and focus, set `tabIndex={section === s.id ? 0 : -1}`; on `lg` drop `role="tabpanel"`/`aria-labelledby` (plain `<section aria-label>`), or keep the tablist visible. Red test: `userEvent.keyboard("{ArrowRight}")` on the selected tab selects the next one.

### 9. [nit] Unclosed `${` in a template is silently literal

- `shared/src/receipt-rules.ts:130,287` only scans complete `${...}` placeholders; `"${subject.name"` passes validation and expands to the literal text `${subject.name`.

Fix: in `checkTemplate`, after the placeholder scan, `if (template.replace(PLACEHOLDER_SCAN, "").includes("${")) issues.add(path, "has an unclosed ${ placeholder")`. Red test in `receipt-rules-validate.test.ts`.

### 10. [nit] Fetcher error text for an empty candidate address

- `web/src/components/variables/EmailFetcher.tsx:91-105` shows `Invalid sender address: ` (nothing after the colon) when a freshly added sender still has an empty address, because `buildReceiptSearchQuery` throws with the empty string.

Fix: before building the preview, if `allSenders.some((s) => s.trim() === "")` set `preview.error = "A sender in the rules has no address yet."`. Red test in `EmailFetcher.test.tsx` ("disables Fetch when a candidate sender is invalid" currently asserts only `/Invalid sender/`).

### 11. [nit] Forward-subject wrapper changes top-level alternation semantics (improvement, document it)

- `shared/src/receipt-rules.ts:658-661` builds `^(?:Fwd: )(?:<pattern without ^>)`; the old `getForwardSubjectRegExp` built `^Fwd: <pattern without ^>`. For a pattern with top-level alternation such as `^a|b`, old matched `b` without the prefix (verified in Node: old `true`, new `false`). None of the default patterns have top-level alternation, so parity for the defaults holds; this is a correctness improvement for custom rules. Add one sentence to the `ReceiptRule.subject` TSDoc and a test in `receipt-rules-apply.test.ts` pinning that `^a|b` requires the prefix for `b`.

### 12. [nit] `setupAddOptionalGlobalVariables` and `saveReceiptRules` fail when a stray value sits in column A of row 30

- `gas/src/util/global-variables-setup.util.ts:314-332` (`findFreeRows`) only considers rows below the last non-empty column-A cell. A heading or comment in `A30` makes both the setup helper and the first web save fail with "No empty row left" even though rows 8-29 are free.

The error message tells the owner what to do, so this is acceptable; mention it in `RECEIPT_RULES_NOTE`/README ("keep column A below the variables empty") or relax the search to "first row below the last _known variable name_ where both cells are empty". No test change unless the behavior changes.

### 13. [nit] `DEFAULT_RECEIPT_RULES` is a mutable shared object used as React state and as the compiled config

- `shared/src/receipt-rules-defaults.ts:11` exports a plain object; `CustomizeVariablesDialog` seeds `candidate` with it, `parseReceiptRulesText` returns it as `config`, and `compileReceiptRules` stores the reference. All current edits copy before writing (`withOptional`, spreads), so no bug today.

Fix (hardening): deep-freeze it in development (`Object.freeze` recursively) or return `structuredClone(DEFAULT_RECEIPT_RULES)` from `parseReceiptRulesText`. Test: mutate a nested field of the parse result and assert `DEFAULT_RECEIPT_RULES` is unchanged.

---

## Verified OK (no finding)

Behavioral parity (engine + `DEFAULT_RECEIPT_RULES` vs the old parser):

- Forwarded detection (`from === ForwardEmailAddress`, now case-insensitive and guarded for empty), `fromPatterns` order and the `tentativeFrom` fallback, the three body markers with skip 5/6/6 and "whole body when no marker", `forwardable: false` => blank receipt (type undefined), explicit `ForwardName` prefix with fall-through to the `To` check and then `defaultTypePrefix` when `ForwardName` is empty.
- Type prefixes: `fixed` for Chase, `recipient` for PayPal/Venmo; type is set even when no rule matches.
- Cost: `text ? parseFloat(text) : undefined`, negate after parse (`NaN` and `-0` preserved), body-sourced cost for `paypal-received`.
- Names: `${body.name}` / `${body.name} - ${body.details}` / `${subject.name}` with `nameFallback`, `(Gas)` suffix and category only when a receipt; empty captured name stays falsy.
- Error messages (3) identical; `No function...` now carries the lower-cased address as planned; the note's `From:` keeps original case.
- Gmail query byte-identical to the pinned string; the forward address is appended last; `in:inbox` + exclusions fixed; subjects quoted and character-filtered; `newer_than` integer-checked; all addresses strict-regex checked. Injection-proof.
- Forward-subject regex wrapping cannot throw for any pattern that validated alone (checked `^+a`, `^*a`, `^{1}a`, `^?a`, `^|^`, `(?<=x)^a`, `^` in Node).
- Edge fixtures were pinned on the old parser first (`receipts.test.ts` still runs them through `ThreadList`), and the engine reproduces all of them.

Security and GAS conventions:

- `setupAddOptionalGlobalVariables` exported from `src/index.ts`; `getActiveSpreadsheet()` everywhere; no `getUi()` in new code; no `getValues/setValues` inside row loops in new code (one `setValues` per added optional row as the plan allows; notes written with one `getNotes/setNotes`).
- `saveReceiptRules`: lock (`tryLock:20000`/`release`), re-read inside the lock, `expectedRaw` check, full validation incl. size, `JSON.stringify(config, null, 2)` (starts with `{` or is `""`), `bustVariablesCache()` after the write, never executes rule regexes (compile only).
- `fetchReceiptEmails`: no lock, no label creation, no `addLabel`/archive (asserted by `expectGmailUntouched`), `maxThreads` default 10, clamp at 50, message cap 100, body caps, total cap; `senders` given => `ReceiptRules` not read.
- Web: all email text is rendered as React children (`<pre>` in `EmailResults.tsx`); no `dangerouslySetInnerHTML`, `eval` or `new Function` in `web/src` or `shared/src`; e2e "email content is shown as text" checks `window.__pwned` stays undefined. The preview runs `applyReceiptRules` from `@budget/shared` (same module as the server). Credentials/API untouched (`services.api` is the existing authenticated client).
- Demo backend `bustVariablesCache()` per request mirrors a fresh GAS execution; the in-execution behavior (save then get in one module instance) is covered by `api-receipt-rules.test.ts` and `receipt-rules-util.test.ts`, so nothing is hidden.
- Docs: README owner steps (push, run `setupAddOptionalGlobalVariables`, redeploy a new version) are correct; AGENTS.md section 3.4 sentence and debug endpoint list are accurate.

---

## Design notes on the known follow-up work (not findings)

1. Unknown senders => "not a receipt". Today an unknown `From` inside a matched thread (for example the owner replying "thanks" in a receipt thread) raises an error, which keeps the thread in the inbox with `Receipts/Error`. Turning it into a note changes labeling: a thread with one receipt and one reply would be archived with a note instead of flagged. Decide that explicitly and pin it with a fixture in `receipts.test.ts`; also give `RuleOutcome` a `reason` (`"unknown-sender" | "not-forwardable" | "no-forward-header" | ...`) so `EmailResults` can explain "Not a receipt" instead of "no rule matched". Three tests pin the current message (`shared/test/receipt-rules-apply.test.ts`, `gas/test/receipt-rules-engine.test.ts`, `gas/test/receipts.test.ts`) and must change together.
2. Non-forward messages from `ForwardEmailAddress` => "not a receipt". Keep the error when the subject starts with a forward prefix but no `fromPattern` matches (that is the case that catches a new Gmail forward format), and return "not a receipt" only when the subject has no prefix. Otherwise a silently changed forward format turns into silently ignored receipts.
3. `extractEmailAddress` preferring `<...>` and stripping quotes: the regex is shared by sender detection, the forwarded `From:` fallback, the `To` check, the note text and the demo backend's search filter, so one change covers all; add fixtures for `"service@paypal.com" <service@paypal.com>` and a bare quoted address. It also changes finding 7's note-text expectation if the From contains quotes.
4. New PayPal "X sent you $Y USD" rule: changing `DEFAULT_RECEIPT_RULES` breaks the "deep-equals the plan's JSON" test and section 2.4 on purpose; more importantly, owners who ran the setup helper have the old defaults JSON in the cell and will not get the new rule. The README should say "after upgrading, open Customize variables and Reset to built-in defaults + Save (or Clear variable)". Order the rule after `paypal-sent`/`paypal-received`; the subject anchor `^(?<name>.+) sent you \$(?<cost>.+) USD$` cannot collide with `^You sent a ...` but `.+` names should be pinned with fixtures containing `$` and non-breaking spaces.
5. Narrowing the query so the forward address only matches forward-prefixed subjects: Gmail cannot scope `subject:` to one address inside a single `from:(a OR b OR fwd)`; it needs `(from:a OR from:b OR (from:fwd subject:Fwd))` (changing the pinned query text and `useRealisticGmailSearch`'s parser), and Gmail's subject search is token-based, so `subject:Fwd` matches anywhere in the subject, not only as a prefix. Low value for the added complexity; the engine-level change in note 2 gives the same user-visible result.
6. Access scope (accepted by owner decisions 2 and 3, recorded here for awareness): `fetchReceiptEmails` with `senders` + `includeProcessed` lets any API-key holder read message bodies from arbitrary senders across all mail (read-only, 100 messages per call, unlimited calls). Before this branch the API key gave sheet access only. If that ever matters, restrict `senders` to addresses in the saved rules plus the candidate's, or require `newerThanDays` for `includeProcessed`.

---

## Fix groups (parallel, disjoint files)

**Group A [Sonnet] GAS + shared query/engine + fetcher.** Findings 1, 2, 5, 7 (tests), 9, 10, 11, 13.
Files: `gas/src/util/receipt-rules.util.ts`, `gas/src/api/receipt-rules-handlers.ts`, `gas/src/api/handlers.ts`, `shared/src/receipt-query.ts`, `shared/src/receipt-rules.ts`, `shared/src/receipt-rules-defaults.ts`, `web/src/components/variables/EmailFetcher.tsx`, tests: `gas/test/receipt-rules-util.test.ts`, `gas/test/api-receipt-rules.test.ts`, `gas/test/receipts.test.ts`, `gas/test/fixtures/receipts.ts`, `shared/test/receipt-rules-validate.test.ts`, `shared/test/receipt-rules-apply.test.ts`, `web/src/components/variables/EmailFetcher.test.tsx`. Do not touch `shared/src/types.ts` or the dialog files (Group B). Finding 1's optional `queryError` DTO field is Group B's if wanted; Group A only needs the thrown message and `mapInvalidConfig`.

**Group B [Sonnet] Web dialog UX/a11y.** Findings 4, 6, 8 (and optionally the `queryError` field from finding 1: add `queryError?: string | null` to `ReceiptRulesStateDto` in `shared/src/types.ts` and render it; the server side of that field would then need a one-line follow-up in Group A's handler after both land, so prefer to skip it unless coordinated).
Files: `web/src/components/variables/CustomizeVariablesDialog.tsx`, `RulesEditor.tsx`, `EmailResults.tsx`, their `.test.tsx` files, `shared/src/types.ts` (only if the optional field is added).

**Group C [Sonnet, or Haiku-class] Docs.** Findings 3, 12 (note text), 7 (plan note about the JSON e2e), 11 (TSDoc sentence can be left to Group A since it edits `receipt-rules.ts`).
Files: `AGENTS.md`, `web-plan.md`, `README.md`, `receipt-rules-plan.md` (sections 1/9: add the `To`-address case-insensitivity and the alternation note; section 8 C6: JSON-tab flow is unit-only), `shared/src/receipt-rules-defaults.ts` only for the `RECEIPT_RULES_NOTE` sentence about keeping column A empty below the variables — coordinate with Group A, which also edits that file for finding 13; simplest is to give the note sentence to Group A and keep Group C to Markdown only.

After all groups: run `pnpm run format:check`, `lint`, `typecheck`, `test`, `build`, `test:e2e` (AGENTS.md section 5.6).
