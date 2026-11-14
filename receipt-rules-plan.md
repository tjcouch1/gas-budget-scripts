# Receipt Rules Plan: data-driven receipt parsing, configurable Gmail senders, "Customize variables" UI

Status: IMPLEMENTED on branch `receipt-rules` (all Section 8 tasks A1-D1 done; full verification run, uncommitted). Branch `receipt-rules`. Implementers follow AGENTS.md (binding) and this plan, one task at a time, strict red -> green TDD. Nothing here runs clasp, pushes, or calls a deployed `/exec`.

Paths are relative to the repo root unless stated. `gas/` means `packages/budget-scripts-gas/`, `shared/` means `packages/budget-shared/`, `web/` means `packages/budget-web/`.

---

## 1. Summary of decisions

| Topic                | Decision                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Rule engine location | `shared/src/receipt-rules.ts` (+ `receipt-rules-defaults.ts`, `receipt-query.ts`). Pure TS, no globals, runs in GAS and the browser.                                                                                                                                                                                                                                                                                                                                         |
| Storage              | ONE new optional GLOBAL variable, key `ReceiptRules`, value = pretty-printed JSON (`JSON.stringify(config, null, 2)`), on the Variables sheet. The Gmail sender list is derived from `senders[].address` inside that JSON (single source of truth; no second variable). `ForwardEmailAddress` stays a separate existing variable and is always added to the query.                                                                                                           |
| Variables range      | `Variables!A1:B10` becomes `Variables!A1:B30` (`VARIABLES_RANGE`, exported from `variables.util.ts` and reused by both setup helpers). Two tests that pin the range text change (Section 7).                                                                                                                                                                                                                                                                                 |
| Defaults             | Missing key, empty or whitespace-only cell => `DEFAULT_RECEIPT_RULES` (encodes today's behavior exactly). The setup helper writes the key AND the defaults JSON plus a long cell note. The web UI shows a "Reset to built-in defaults" action which clears the cell.                                                                                                                                                                                                         |
| Invalid value        | `getCompiledReceiptRules()` throws `Error("ReceiptRules variable is invalid: <path>: <message>. Fix it in the web app (Settings > Customize variables) or clear the cell to use the built-in defaults.")` BEFORE any Gmail search or sheet write. Menu/trigger imports fail loudly (GAS emails trigger failures). The API maps it to the new error code `invalid_config`; the web app shows the message with an "Open Customize variables" action. Never silently mis-parse. |
| Gmail query          | Built only by `buildReceiptSearchQuery()` from validated addresses. Always adds `in:inbox` and the four fixed label exclusions (unless the read-only fetcher asks for `includeProcessed`). No raw query text anywhere.                                                                                                                                                                                                                                                       |
| API                  | Three new actions: `fetchReceiptEmails` (read-only, no labels), `getReceiptRules` (read), `saveReceiptRules` (write, validated, under the script lock, optional `expectedRaw` conflict check). Same API key auth.                                                                                                                                                                                                                                                            |
| Regex safety         | Server never executes candidate regexes (the client runs the preview). Save compiles every regex and enforces size limits (Section 3.4). ReDoS of saved rules is accepted as the owner's responsibility and documented in the TSDoc/note.                                                                                                                                                                                                                                    |
| Web editor           | Both a structured form and a raw JSON textarea (tab switch), no new dependency. Live preview runs the shared engine in the browser against emails fetched by the API. Email bodies are rendered as text in `<pre>` (React escapes; never `dangerouslySetInnerHTML`).                                                                                                                                                                                                         |
| Characterization     | New edge-case fixtures pinned against the OLD parser first (Task B1), then the engine is written to satisfy all fixtures (A2), then `ThreadList` switches to the engine (B3) while a parity test holds old == new, then the old regex code is deleted (B4). Existing tests are not weakened.                                                                                                                                                                                 |

Sender address comparison becomes case-insensitive (today: exact). This cannot change any existing test and only turns a would-be "No function to handle receipt message" error into a parse. The `To`-address check for `ForwardEmailAddress` (`recipientPrefix()`, step 4 of 2.3) is also compared case-insensitively (today: exact). These are the two intentional micro-changes; both are called out again in Section 9, and the `To` case is pinned by "uses ForwardName when sent To the forward address (case-insensitive)" in `shared/test/receipt-rules-apply.test.ts`.

---

## 2. Rule engine (shared)

### 2.1 Types (`shared/src/receipt-rules.ts`)

```ts
export const RECEIPT_RULES_VERSION = 1 as const;

/** A JavaScript regular expression given as JSON-safe source text (no surrounding slashes). */
export interface RegexSpec {
  /** RegExp source, e.g. "^You paid (?<name>.+) \\$(?<cost>.+)$". Named groups are what templates reference. */
  pattern: string;
  /** Any of "i", "m", "s", "u" (each at most once). Never "g", "y" or "d". Default "". */
  flags?: string;
}

export type BodySource = "plainBody" | "htmlBody";

export interface ReceiptRule {
  /** Unique within its sender. Shown in the web app's preview ("matched rule: chase-purchase"). 1-40 chars, [A-Za-z0-9_-]. */
  id: string;
  /**
   * Required anchor: run against the subject. The first rule (in order) whose subject matches wins.
   * For a forwarded email the engine uses "^Fwd: " + pattern-without-its-leading-"^" (today's getForwardSubjectRegExp).
   */
  subject: RegexSpec;
  /**
   * Optional second match, run once after the subject matched. Its groups are available to templates as
   * ${body.<group>}. For a forwarded email it runs on the original (unwrapped) body, see ForwardingConfig.
   * When it does not match, every ${body.*} reference is "missing" (see `name`).
   */
  body?: RegexSpec & { source: BodySource };
  /**
   * Template for the receipt name. Placeholders: ${subject.<group>} and ${body.<group>}. Any other text is literal.
   * If ANY referenced group is missing (regex did not match, group did not participate, or body not available)
   * the template yields nothing and `nameFallback` is tried the same way. If both fail, name is undefined.
   * Omit for rules that have no name.
   */
  name?: string;
  nameFallback?: string;
  /**
   * Where the cost comes from. The group's text goes through parseFloat (exactly like today, so "1,234.56" is 1).
   * An empty or missing group gives no cost. `negate` multiplies by -1 (refunds, money received). Omit for no cost.
   */
  cost?: { source: "subject" | "body"; group: string; negate?: boolean };
  /** Envelope/category written to the sheet when this rule yields a receipt (cost or name present). */
  category?: string;
}

export interface SenderRules {
  /** Exact sender address (compared case-insensitively). Strict syntax, see STRICT_EMAIL_REGEX. */
  address: string;
  /** Base of the "Paid with" type: "Credit", "Paypal", "Venmo". 1-40 chars, no "=" "+" "-" "@" start. */
  type: string;
  /**
   * "fixed": type is `type` verbatim.
   * "recipient": type is "<prefix> <type>" where prefix is ForwardName if the email was forwarded by the forwarder
   * or was sent To ForwardEmailAddress (and ForwardName is non-empty), else `defaultTypePrefix`.
   */
  typePrefix: "fixed" | "recipient";
  /**
   * true: emails From ForwardEmailAddress whose body says they were forwarded from this sender are parsed with these
   * rules (subject must start with "Fwd: "). false: such forwards are not receipts (they only get a note).
   */
  forwardable: boolean;
  /** Evaluated in order; first subject match wins. May be empty (sender is known, nothing is ever a receipt). */
  rules: ReceiptRule[];
}

export interface ForwardingConfig {
  /**
   * Regexes with a named group `email`, run in order on the forwarded email's whole plain body to find the original
   * sender ("From: PayPal <service@paypal.com>"). The group may hold a formal address; the address is extracted.
   * None matching => error "Could not get 'from' email address from forwarded email with subject '<subject>'".
   */
  fromPatterns: RegexSpec[];
  /**
   * Markers that start the forwarded content. The first marker found in the plain body wins: the original body is
   * everything from the marker on, split on "\n", with the first `skipLines` lines dropped. None found => whole body.
   */
  bodyMarkers: { marker: string; skipLines: number }[];
}

export interface ReceiptRulesConfig {
  version: typeof RECEIPT_RULES_VERSION;
  /** Prefix for "recipient" types when the email was not for the forwarder. Today "TJ". 1-20 chars. */
  defaultTypePrefix: string;
  forwarding: ForwardingConfig;
  /** No two senders may share an address. None may equal ForwardEmailAddress (checked at save and compile time). */
  senders: SenderRules[];
}
```

The TSDoc above is the spec; copy it into the source and keep `RECEIPT_RULES_NOTE` (Section 4.3) consistent with it.

### 2.2 Functions

```ts
export interface RuleIssue {
  path: string;
  message: string;
} // path like "senders[1].rules[0].subject.pattern"

export type ValidationResult =
  { ok: true; config: ReceiptRulesConfig } | { ok: false; issues: RuleIssue[] };

/** Structural + semantic validation of parsed JSON (never throws). Collects up to 50 issues. */
export function validateReceiptRulesConfig(
  value: unknown,
  options?: { forwardEmailAddress?: string },
): ValidationResult;

/** JSON.parse + validate. Empty/whitespace text => { ok: true, config: DEFAULT_RECEIPT_RULES, isDefault: true }. */
export function parseReceiptRulesText(
  text: string | null | undefined,
  options?,
): ValidationResult & { isDefault?: boolean };

export interface CompiledReceiptRules {
  config: ReceiptRulesConfig;
  /** true when any rule uses body.source === "htmlBody" (GAS then calls message.getBody()) */
  usesHtmlBody: boolean;
  /** lower-cased sender addresses in config order */
  senderAddresses: string[];
  // internal: compiled RegExp per rule (subject, forwardSubject, body), per fromPattern
}
/** Throws Error("ReceiptRules ... invalid: <path>: <message> ...") on the first issue. */
export function compileReceiptRules(
  config: ReceiptRulesConfig,
  options?: { forwardEmailAddress?: string },
): CompiledReceiptRules;

export interface RuleEmail {
  from: string; // message.getFrom(): "Chase <no.reply.alerts@chase.com>" or bare address
  to: string; // message.getTo()
  subject: string;
  plainBody: string;
  htmlBody?: string; // only needed when compiled.usesHtmlBody
}
export interface RuleContext {
  forwardEmailAddress: string;
  forwardName: string;
}

export interface RuleOutcome {
  /** Lower-cased address the rules were chosen by (original sender for forwards). null only for an unknown forwarded sender. */
  senderAddress: string | null;
  forwarded: boolean;
  /** null: sender known but no rule's subject matched (or forwarded sender has no forwardable rules). */
  ruleId: string | null;
  receipt: {
    cost: number | undefined;
    name: string | undefined;
    category: string | undefined;
    type: string | undefined;
  };
  /** `!!cost || !!name` with JS truthiness (so 0 and NaN are NOT receipts), exactly ThreadList's test today. */
  isReceipt: boolean;
  captures: {
    subject: Record<string, string | undefined> | null;
    body: Record<string, string | undefined> | null;
  };
  warnings: string[]; // e.g. "Rule x needs the HTML body but it was not provided"
}
/**
 * Throws plain Error with these exact messages (pinned by tests):
 *  - "Could not get email address from <from>"
 *  - "No function to handle receipt message from <address>"
 *  - "Could not get 'from' email address from forwarded email with subject '<subject>'"
 */
export function applyReceiptRules(
  rules: CompiledReceiptRules,
  email: RuleEmail,
  context: RuleContext,
): RuleOutcome;

/** Formal-address extraction: /(?<email>[^<>\s ]+@[^<>\s ]+)/ ; throws "Could not get email address from <text>" */
export function extractEmailAddress(formal: string): string;

/** Expands ${subject.x}/${body.x}; returns undefined if any referenced group is undefined. Exported for tests. */
export function expandTemplate(
  template: string,
  groups: {
    subject: Record<string, string | undefined> | null;
    body: Record<string, string | undefined> | null;
  },
): string | undefined;
```

### 2.3 `applyReceiptRules` algorithm (must reproduce today's branches)

1. `fromAddress = extractEmailAddress(email.from).toLowerCase()` (throws as today).
2. If `fromAddress === context.forwardEmailAddress.toLowerCase()` and that is non-empty: forwarded path.
   a. For each `forwarding.fromPatterns` in order, `exec` on `email.plainBody`; first match wins. None => throw the forwarded-from error.
   b. `original = groups.email`; if `extractEmailAddress` succeeds on it use that, else use `original` as is (today's tentativeFrom fallback). Lower-case.
   c. Find sender with that address and `forwardable === true`. None => return `{ senderAddress: null, forwarded: true, ruleId: null, receipt: all undefined, isReceipt: false, captures: {null,null}, warnings: [] }` (today: blank ReceiptInfo => note).
   d. `body = unwrapForwardedBody(email.plainBody, forwarding.bodyMarkers)` (first found marker wins; `substring(indexOf(marker)).split("\n").slice(skipLines).join("\n")`); `subjectRegex = forwardSubject`; prefix: `context.forwardName || recipientPrefix()`.
3. Else: find sender by `fromAddress`. None => throw "No function to handle receipt message from <fromAddress>" (use the lower-cased extracted address). `body = email.plainBody`; `subjectRegex = subject`; prefix: `recipientPrefix()`.
4. `recipientPrefix()`: try `extractEmailAddress(email.to)`; on throw => undefined (today logs and ignores; the Logger message is not pinned by any test: implementers grep `Threw while trying` in `gas/test` to confirm, and drop it). If it equals `forwardEmailAddress` (case-insensitive) and `forwardName` non-empty => `forwardName`, else undefined. Final prefix = that `|| config.defaultTypePrefix`.
5. `type = sender.typePrefix === "fixed" ? sender.type : `${prefix} ${sender.type}``. Type is set even when no rule matches (today's behavior; not observable but keep).
6. For each rule in order: `m = subjectRegex.exec(email.subject)`; first match wins. No match => `ruleId: null`, receipt `{undefined, undefined, undefined, type}`.
7. Body: if `rule.body`, text = `source === "plainBody" ? body : email.htmlBody` (for forwarded + htmlBody: use `email.htmlBody` unchanged; add warning "forwarded HTML bodies are not unwrapped"). If `htmlBody` needed but undefined => body captures null + warning. `exec` once.
8. `cost`: group text `t = captures[source]?.[group]`; `cost = t ? parseFloat(t) : undefined`; if `negate` and cost !== undefined => `cost * -1`.
9. `name = expandTemplate(rule.name) ?? expandTemplate(rule.nameFallback)` (each only if defined).
10. `isReceipt = !!cost || !!name`; `category = isReceipt ? rule.category : undefined`.

Regexes are compiled once per `compileReceiptRules` and are NOT global, so `exec` is stateless.

### 2.4 Worked example: `DEFAULT_RECEIPT_RULES` (`shared/src/receipt-rules-defaults.ts`)

This is the exact JSON the setup helper writes (as a TS object literal `satisfies ReceiptRulesConfig`). Every current regex and branch is mapped; the comment after each entry names the old constant.

```json
{
  "version": 1,
  "defaultTypePrefix": "TJ",
  "forwarding": {
    "fromPatterns": [
      {
        "pattern": "[Ff]orwarded message.*\\s*\\r?\\n\\r?\\n?\\*?From:\\*?[ \\u00a0](?<email>.+)\\r?\\n"
      },
      {
        "pattern": "転送されたメッセージ.*\\s*\\r?\\n\\r?\\n?\\*?差出人:\\*?[ \\u00a0](?<email>.+)\\r?\\n"
      }
    ],
    "bodyMarkers": [
      { "marker": "---------- Forwarded message ---------", "skipLines": 5 },
      { "marker": "Begin forwarded message:", "skipLines": 6 },
      { "marker": "転送されたメッセージ:", "skipLines": 6 }
    ]
  },
  "senders": [
    {
      "address": "no.reply.alerts@chase.com",
      "type": "Credit",
      "typePrefix": "fixed",
      "forwardable": false,
      "rules": [
        {
          "id": "chase-purchase",
          "subject": {
            "pattern": "^You made a \\$(?<cost>.+) transaction with (?<name>.+)$"
          },
          "name": "${subject.name}",
          "cost": { "source": "subject", "group": "cost" }
        },
        {
          "id": "chase-refund",
          "subject": {
            "pattern": "^You have a \\$(?<cost>.+) credit pending on your credit card$"
          },
          "body": {
            "source": "plainBody",
            "pattern": "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n"
          },
          "name": "${body.name}",
          "cost": { "source": "subject", "group": "cost", "negate": true }
        },
        {
          "id": "chase-gas",
          "subject": { "pattern": "^You used your card at a gas station$" },
          "body": {
            "source": "plainBody",
            "pattern": "\\r?\\nMerchant\\s+(?<name>.*\\S)\\s*\\r?\\n"
          },
          "name": "${body.name} (Gas)",
          "category": "Gas"
        }
      ]
    },
    {
      "address": "service@paypal.com",
      "type": "Paypal",
      "typePrefix": "recipient",
      "forwardable": true,
      "rules": [
        {
          "id": "paypal-sent",
          "subject": {
            "pattern": "^You sent a \\$(?<cost>.+)[\\s\\u00a0]USD payment$"
          },
          "body": {
            "source": "plainBody",
            "pattern": "You sent \\$.+ to (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n"
          },
          "name": "${body.name} - ${body.details}",
          "cost": { "source": "subject", "group": "cost" }
        },
        {
          "id": "paypal-received",
          "subject": { "pattern": "^Money is waiting for you$" },
          "body": {
            "source": "plainBody",
            "pattern": "Accept your \\$(?<cost>.+)[\\s\\u00a0]USD from (?<name>.+)\\r?\\n\\r?\\n.*\\r?\\n.*\\r?\\n\\r?\\n(?<details>.+)\\r?\\n"
          },
          "name": "${body.name} - ${body.details}",
          "cost": { "source": "body", "group": "cost", "negate": true }
        }
      ]
    },
    {
      "address": "venmo@venmo.com",
      "type": "Venmo",
      "typePrefix": "recipient",
      "forwardable": true,
      "rules": [
        {
          "id": "venmo-paid",
          "subject": { "pattern": "^You paid (?<name>.+) \\$(?<cost>.+)$" },
          "body": {
            "source": "plainBody",
            "pattern": "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n"
          },
          "name": "${subject.name} - ${body.details}",
          "nameFallback": "${subject.name}",
          "cost": { "source": "subject", "group": "cost" }
        },
        {
          "id": "venmo-charge",
          "subject": {
            "pattern": "^You completed (?<name>.+)'s \\$(?<cost>.+) charge request$"
          },
          "body": {
            "source": "plainBody",
            "pattern": "charged\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n"
          },
          "name": "${subject.name} - ${body.details}",
          "nameFallback": "${subject.name}",
          "cost": { "source": "subject", "group": "cost" }
        },
        {
          "id": "venmo-received",
          "subject": { "pattern": "^(?<name>.+) paid you \\$(?<cost>.+)$" },
          "body": {
            "source": "plainBody",
            "pattern": "paid\\s*.*\\s*\\s*.*\\s*\\r?\\n(?<details>.+)\\r?\\n"
          },
          "name": "${subject.name} - ${body.details}",
          "nameFallback": "${subject.name}",
          "cost": { "source": "subject", "group": "cost", "negate": true }
        }
      ]
    }
  ]
}
```

Mapping of every old symbol/branch:

| Old code                                                                                                                                    | New place                                                                                                         |
| ------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `formalEmailAddressRegExp`, `getEmailAddress`                                                                                               | `extractEmailAddress` (engine, fixed, not configurable)                                                           |
| `getForwardSubjectRegExp`                                                                                                                   | engine: `"^Fwd: " + pattern.replace(/^\^/, "")` compiled per rule                                                 |
| `getForwardEmailAddress` (2 regexes, tentativeFrom fallback)                                                                                | `forwarding.fromPatterns` + step 2b                                                                               |
| `getForwardPlainBody` (3 markers, skip 5/6/6 lines)                                                                                         | `forwarding.bodyMarkers` + step 2d                                                                                |
| `getMessagePartInfo` (`cost ? parseFloat : undefined`)                                                                                      | step 8                                                                                                            |
| `chaseSubjectReceiptRegExp` / `RefundRegExp` / `GasRegExp`, `chaseBodyMerchantRegExp`                                                       | `chase-purchase`, `chase-refund`, `chase-gas` (body regex duplicated in two rules; that is fine)                  |
| Chase: refund keeps cost even if body fails; gas has no cost                                                                                | `chase-refund` cost from subject; `chase-gas` has no `cost`                                                       |
| Chase gas `category = "Gas"` only inside body-match branch                                                                                  | `category` applied only when `isReceipt` (gas without body match has no name and no cost => not a receipt anyway) |
| `type = "Credit"`                                                                                                                           | `typePrefix: "fixed"`                                                                                             |
| `paypalReceiptEmailAddress`, `venmoReceiptEmailAddress`                                                                                     | `senders[].address`                                                                                               |
| PayPal/Venmo `finalTypePrefix` logic (explicit ForwardName when forwarded, To-address check, "TJ" default, `!finalTypePrefix` falsy checks) | step 4/5 and `defaultTypePrefix`                                                                                  |
| PayPal sent: name only if body matches, cost from subject regardless                                                                        | `paypal-sent`: template fails => undefined name; cost from subject                                                |
| PayPal received: cost and name both only if body matches, cost negated                                                                      | `paypal-received`: cost source body + negate                                                                      |
| Venmo: name = subject.name, `- details` appended only if body matches                                                                       | `name` + `nameFallback`                                                                                           |
| Venmo received: `cost * -1`                                                                                                                 | `negate: true`                                                                                                    |
| Proxy `get` for ForwardEmailAddress; forwarded Chase => blank receipt                                                                       | `forwardable: false` on Chase => step 2c blank receipt                                                            |
| `getReceiptInfoMap[from]` undefined => throw "No function to handle receipt message from"                                                   | step 3                                                                                                            |
| Logger.log on To-address parse failure                                                                                                      | dropped (not pinned; confirm by grep)                                                                             |
| `getUnicodeAnalysis` (unused debug helper)                                                                                                  | delete                                                                                                            |

Nothing in today's behavior is inexpressible. Two things are deliberately NOT configurable: the formal-address extraction regex and the `Fwd: ` subject prefix (constant `FORWARD_SUBJECT_PREFIX = "Fwd: "` in the engine; promote to config later if needed).

### 2.5 Validation rules (`validateReceiptRulesConfig`)

Collect issues (do not stop at first) with precise `path`s:

- Root must be an object; `version === 1`.
- `defaultTypePrefix`: string, 1-20 chars, `isUnsafeFreeText` false (shared `names.ts`), no `|`.
- `forwarding.fromPatterns`: array 1-10; each `RegexSpec` valid (below) and its pattern contains `(?<email>`.
- `forwarding.bodyMarkers`: array 0-10; `marker` 1-100 chars non-empty; `skipLines` integer 0-20.
- `senders`: array 1-25. `address` matches `STRICT_EMAIL_REGEX`, max 254 chars; duplicates (case-insensitive) => issue on the later one; equals `options.forwardEmailAddress` (case-insensitive) => issue. `type` 1-40 chars, no `|`, `isUnsafeFreeText` false. `typePrefix` in set. `forwardable` boolean. `rules` array 0-25.
- Rule: `id` matches `/^[A-Za-z0-9_-]{1,40}$/`, unique within sender. `subject`, `body` valid RegexSpec; `body.source` in set. `name`/`nameFallback`: string 1-200, every `${a.b}` placeholder must have `a` in {subject, body} (`body` only if `rule.body` exists) and group `b` declared in that regex (scan pattern with `/\(\?<([A-Za-z_][A-Za-z0-9_]*)>/g`); the expanded text must not be able to start with `=`, `+`, `-`, `@` from literal text (check the literal prefix before the first placeholder with `isUnsafeFreeText`; group text is validated at write time by existing sheet code paths, which today already write names from regexes unchanged, so do not add new rejection there). `cost.source` in set (body only if `rule.body`), `cost.group` declared in that regex; `negate` boolean. `category` 1-60 chars, `isUnsafeFreeText` false.
- RegexSpec: `pattern` string 1-500 chars; `flags` matches `/^(?!.*(.).*\1)[imsu]*$/`; `new RegExp(pattern, flags)` must not throw (issue message = the RegExp error message).
- Unknown keys anywhere => issue "unknown key" (keeps typos loud).
- Serialized size: `JSON.stringify(config, null, 2).length <= MAX_RECEIPT_RULES_CHARS` (= 40_000; a cell holds 50,000).

`compileReceiptRules` runs validation (with `forwardEmailAddress`) and throws on the first issue with message `ReceiptRules variable is invalid: ${path}: ${message}. Fix it in the web app (Settings > Customize variables) or clear the cell to use the built-in defaults.`

### 2.6 Gmail query builder (`shared/src/receipt-query.ts`)

```ts
export const STRICT_EMAIL_REGEX =
  /^[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/;
export function isStrictEmailAddress(value: unknown): value is string; // also length <= 254
/** Allowed subject filter text: 1-100 chars, none of " ( ) { } : \ newline; trimmed non-empty */
export function isSafeSubjectFilter(value: unknown): value is string;

/** Fixed, never configurable (AGENTS.md section 3.4 labels) */
export const RECEIPT_LABEL_EXCLUSIONS =
  "NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";

export interface ReceiptQueryOptions {
  /** default false: adds "in:inbox" and the label exclusions. true (fetcher only): search all mail, no exclusions */
  includeProcessed?: boolean;
  /** each validated with isSafeSubjectFilter; emitted as subject:("a" OR "b") */
  subjects?: string[];
  /** 1-3650; emitted as newer_than:<n>d */
  newerThanDays?: number;
}
/**
 * `from:(a OR b OR <forwardEmailAddress>) in:inbox <exclusions> ` -- note the trailing space, pinned by
 * gas/test/budgeting.test.ts. Throws Error("Invalid sender address: <x>") for any address failing the strict regex.
 * forwardEmailAddress is appended last when non-empty (today it is always appended). No de-duplication.
 */
export function buildReceiptSearchQuery(
  senderAddresses: string[],
  forwardEmailAddress: string,
  options?: ReceiptQueryOptions,
): string;
```

With the defaults and `fwd@example.com` the output must be byte-identical to the string on `gas/test/budgeting.test.ts:156`. Extra clauses (`subject:(...)`, `newer_than:`) go after `from:(...)` and before `in:inbox`; when `includeProcessed` the `in:inbox` and exclusions are omitted but the trailing space stays.

---

## 3. GAS package changes

### 3.1 `gas/src/util/variables.util.ts`

- `export const VARIABLES_SHEET_NAME = "Variables"; export const VARIABLES_RANGE = "Variables!A1:B30";` (was A1:B10, private). `sheet-variables-setup.util.ts` imports these instead of its own copies.
- `Variables` type gains `/** Optional (see GLOBAL_VARIABLE_NOTES.ReceiptRules) */ ReceiptRules?: string;` NOT added to `VARIABLE_NAMES`.
- `export const OPTIONAL_VARIABLE_NAMES = ["ReceiptRules"] as const;`
- `export const GLOBAL_VARIABLE_NOTES = { ReceiptRules: RECEIPT_RULES_NOTE } as const;` where `RECEIPT_RULES_NOTE` is a string constant in `gas/src/util/receipt-rules.util.ts` (Section 4.3) to keep variables.util free of engine imports (import the constant only; no cycle since receipt-rules.util imports `getVariables` from variables.util... to avoid the cycle, define `RECEIPT_RULES_NOTE` in `shared/src/receipt-rules-defaults.ts` and import it in both).

### 3.2 `gas/src/util/receipt-rules.util.ts` (new)

```ts
let cache: { raw: string; compiled: CompiledReceiptRules } | undefined;
/** Compiles the ReceiptRules variable (memoized per execution by raw text; `bustVariablesCache` makes getVariables re-read). Throws the invalid-config Error. */
export function getCompiledReceiptRules(): CompiledReceiptRules;
/** Raw cell text ("" when the key is missing or empty) */
export function getReceiptRulesRaw(): string;
export function getReceiptSenderAddresses(): string[]; // compiled.senderAddresses
export function getReceiptSearchQuery(): string; // buildReceiptSearchQuery(senders, Variables.ForwardEmailAddress)
export function toRuleEmail(
  message: GoogleAppsScript.Gmail.GmailMessage,
  includeHtml: boolean,
): RuleEmail; // getFrom/getTo/getSubject/getPlainBody (+getBody when includeHtml)
export function getRuleContext(): RuleContext; // { forwardEmailAddress, forwardName } from Variables
```

Raw non-string cell (a number) => treated as invalid with path `ReceiptRules`, message "must be text (JSON)".

### 3.3 `gas/src/models/thread-list.model.ts` (after switch-over)

- Constructor: `constructor(threads, rules: CompiledReceiptRules = getCompiledReceiptRules(), context: RuleContext = getRuleContext())`. Both defaults evaluate BEFORE the loop, so an invalid variable throws out of the constructor (fail loudly), not into per-thread errors.
- Per message: `const email = toRuleEmail(message, rules.usesHtmlBody); const outcome = applyReceiptRules(rules, email, context); const receipt = new ReceiptInfo(message, outcome.receipt.cost, outcome.receipt.name, outcome.receipt.category, outcome.receipt.type);` then the EXISTING `if (!receipt.cost && !receipt.name)` note logic, unchanged. Errors thrown by the engine land in the same `catch` => same "Error while processing message ..." text. The note text `Message is not a receipt:\nFrom: ${from}...` needs `from`: use `extractEmailAddress(message.getFrom())` result (same text as before: the pre-lower-cased address; keep the original case by using `extractEmailAddress` directly, not `outcome.senderAddress`).
- Everything from `// #region regexp helper functions` through `getReceiptInfoMap` is deleted in Task B4. `compareReceiptInfosByDateAscending`, `ThreadList` body, `ignoredThreads`, `getAllReceiptInfos` are untouched.

### 3.4 `gas/src/util/budgeting.util.ts`

- `getTJGmailSearchQuery()` body becomes `return getReceiptSearchQuery();` (keep the function name to minimize diff; rename later if desired). `getChaseReceipts` is otherwise unchanged. Because `getReceiptSearchQuery` compiles the rules, an invalid variable throws before `GmailApp.search`.

### 3.5 API (`gas/src/api/*`) and shared DTOs (`shared/src/types.ts`)

Add to `ApiErrorCode`: `"invalid_config"`. Add to `ApiFailure`: `issues?: { path: string; message: string }[]`. `ApiError` constructor gets an optional 4th arg `issues?` and `toFailure` copies it.

```ts
export interface FetchReceiptEmailsRequest extends BaseRequest {
  action: "fetchReceiptEmails";
  /** Sender addresses for the query. Omit to use the sheet's ReceiptRules (or the built-in defaults). 1-20, strict syntax. */
  senders?: string[];
  /** Extra subject filters, 0-5, see isSafeSubjectFilter. */
  subjects?: string[];
  /** Default false (inbox, unprocessed only). true searches all mail including labeled/archived receipts. */
  includeProcessed?: boolean;
  /** Default false. true adds each message's HTML body (as text). */
  includeHtml?: boolean;
  /** 1-50, default 20 threads. */
  maxThreads?: number;
  /** 1-3650 days. */
  newerThanDays?: number;
}
export interface ReceiptEmailDto {
  threadId: string;
  messageId: string;
  from: string;
  to: string;
  subject: string;
  /** ISO 8601 */ date: string;
  plainBody: string;
  plainBodyTruncated: boolean;
  htmlBody?: string;
  htmlBodyTruncated?: boolean;
  /** Thread labels, e.g. ["Receipts", "Receipts/Scripted"] */ labels: string[];
}
export interface FetchReceiptEmailsResultDto {
  query: string;
  threadCount: number;
  emails: ReceiptEmailDto[];
  /** true when the message or size caps cut the list */ truncated: boolean;
}

export interface GetReceiptRulesRequest extends BaseRequest {
  action: "getReceiptRules";
}
export interface ReceiptRulesStateDto {
  /** Cell text, "" when missing/empty */ raw: string;
  /** Whether the ReceiptRules key exists on the Variables sheet */ hasVariable: boolean;
  /** Parsed config when valid (defaults when raw is empty) */ config: ReceiptRulesConfig | null;
  isDefault: boolean;
  issues: { path: string; message: string }[]; // non-empty when invalid
  forwardEmailAddress: string;
  forwardName: string;
  /** Built from `config` (or null when invalid) */ query: string | null;
}

export interface SaveReceiptRulesRequest extends BaseRequest {
  action: "saveReceiptRules";
  /** The config object, or null to clear the cell (back to built-in defaults). */
  rules: ReceiptRulesConfig | null;
  /** Optional optimistic check: the `raw` from the last getReceiptRules. Mismatch => error code "conflict". */
  expectedRaw?: string;
}
// response data: ReceiptRulesStateDto (after the write)
```

Add the three to the `ApiRequest` union (the web `ApiRequestBody` type follows automatically).

Handlers (`gas/src/api/handlers.ts` or a new `gas/src/api/receipt-rules-handlers.ts` imported by the router; prefer the new file to keep handlers.ts under control):

- `handleFetchReceiptEmails(req)`: validate fields (`badRequest` messages: "senders must be an array of 1-20 email addresses", "senders[2] is not a valid email address", "subjects[0] contains characters that are not allowed", "maxThreads must be an integer from 1 to 50", ...). Senders default `getReceiptSenderAddresses()` (invalid variable => `ApiError("invalid_config", message)`). `query = buildReceiptSearchQuery(senders, Variables.ForwardEmailAddress, {...})`. `threads = GmailApp.search(query, 0, maxThreads)`. Walk messages; stop at `FETCH_MAX_MESSAGES = 100` or when cumulative body chars exceed `FETCH_MAX_TOTAL_CHARS = 3_000_000` (`truncated = true`). `plainBody` cut at `FETCH_PLAIN_BODY_MAX = 20_000`, `htmlBody` at `FETCH_HTML_BODY_MAX = 60_000`, flags set. `labels = thread.getLabels().map(l => l.getName())`. No lock, no label writes, no archive. Date `message.getDate().toISOString()`.
- `handleGetReceiptRules()`: `raw = getReceiptRulesRaw()`; `parseReceiptRulesText(raw, { forwardEmailAddress })`; build DTO (no lock).
- `handleSaveReceiptRules(req)`: `rules` must be object or null; `expectedRaw` optional string. Under `tryWithScriptLock(fn, busy)` (same `busy` message as syncReceipts): re-read raw; if `expectedRaw !== undefined && expectedRaw !== raw` => `ApiError("conflict", "The rules on the sheet changed since you loaded them. Reload and try again.")` (no `conflict.current`). If `rules === null` => write `""`. Else `validateReceiptRulesConfig(rules, { forwardEmailAddress })`; issues => `ApiError("bad_request", "ReceiptRules is invalid: <first issue path>: <msg>" , undefined, issues)`. Text = `JSON.stringify(config, null, 2)`; length check. Write via `ensureGlobalVariableRow("ReceiptRules")` (Section 4.2) which returns the value `Range`; `setValues([[text]])` (text starts with `{`, never `=`). `bustVariablesCache()`. Return `handleGetReceiptRules().data`.

Router: three new `case`s. `handleSyncReceipts`: wrap so an error whose message starts with `ReceiptRules variable is invalid` becomes `ApiError("invalid_config", message)` (add a check before `SYNC_USER_ERROR_REGEX`).

### 3.6 Mocks (`gas/test/gas-mocks/gmail.ts`), minimal

- `MessageInit.html?: string`; `MockMessage.getBody()` returns `init.html ?? ""`.
- `MockThread.getLabels()` returns `this.labels`.
- `MockRange.getNotes/setNotes` already exist; nothing else needed for the Variables sheet.

### 3.7 Demo backend (`gas/dev/*`, test-only)

- `useRealisticGmailSearch`: parse the query: `from:(a OR b)` => keep threads whose first message's extracted from-address is in the set (case-insensitive); `subject:("x" OR "y")` => subject contains any (case-insensitive); if the query contains `in:inbox` => exclude archived and `Receipts/Scripted`-labeled threads (today's behavior), else include everything. Add `pnpm`-free unit coverage in `gas/test/demo-backend.test.ts`.
- `DemoBackend.dumpVariables()` => `[key, value][]` of `Variables!A1:B30`; server route `GET __debug/variables`. `reset()` already re-seeds (the seed does not include `ReceiptRules`, so each e2e starts on defaults). `addThreads` accepts `html`.
- `createDemoThreads()` unchanged; `demoInbox` already includes a non-receipt.

### 3.8 `gas/src/index.ts`

`export { setupAddOptionalGlobalVariables } from "./util/global-variables-setup.util";` (build discovers it). No menu change.

---

## 4. Setup helper and cell notes

### 4.1 `gas/src/util/global-variables-setup.util.ts` (new)

```ts
/** Finds the key's value Range in Variables!A1:B30, adding the key in the first fully-empty row below the last used row when missing. Throws if the range is full. Returns { valueRange, added: boolean }. */
export function ensureGlobalVariableRow(
  key: (typeof OPTIONAL_VARIABLE_NAMES)[number],
): { valueRange: Range; added: boolean };
/** Adds every missing optional global variable with its default value and a note on the key cell (GLOBAL_VARIABLE_NOTES). Existing keys/values/notes are kept. Batch I/O: one getValues, one setValues per added row, one getNotes/setNotes on the key column. Logs what it did. */
export function addOptionalGlobalVariables(): void;
/** Run once from the Apps Script editor */
export function setupAddOptionalGlobalVariables(): void;
```

Default value for `ReceiptRules` = `JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2)`. "Fully-empty row" = both key and value cells empty (same `isEmptyCell` as the sheet helper; move `isEmptyCell`, `toA1`, `columnToLetters` to a small `gas/src/util/a1.util.ts` and import from both helpers, or keep duplicates if that churns tests; prefer the shared helper). If a non-empty cell sits in column A or B below the last key within A1:B30 that is not a key/value pair, do not overwrite it: choose the first row where BOTH are empty. Call `bustVariablesCache()` afterwards.

### 4.2 Range growth safety

Rows 1-7 hold today's required keys in the owner's sheet; rows 8-10 may hold anything. Expanding the read range to row 30 only adds rows whose keys are ignored unless they match a known name, so unrelated content in A11:B30 is harmless for reads unless column A literally equals a variable name (then it is picked up, as rows 8-10 are today). The `SheetVariables` value cell (e.g. `H1:I8`) lives in column B of the Variables sheet and is untouched by this feature.

### 4.3 `RECEIPT_RULES_NOTE` (in `shared/src/receipt-rules-defaults.ts`)

Plain text, <= 2,500 chars, no Markdown. Content: what the variable is, "empty = built-in defaults", edit in the web app (Settings > Customize variables) which validates and tests before saving, the JSON shape in a compact outline (version, defaultTypePrefix, forwarding.fromPatterns/bodyMarkers, senders[].address/type/typePrefix/forwardable/rules[].id/subject/body/name/nameFallback/cost/category), placeholder syntax, first-match-wins, how the Gmail query is derived from `senders[].address` + ForwardEmailAddress with `in:inbox` and the label exclusions always added, that regexes run in Apps Script so slow patterns slow the import, and that an invalid value stops imports with an error until fixed or cleared. The `Variables.ReceiptRules` TSDoc says "see GLOBAL_VARIABLE_NOTES.ReceiptRules / RECEIPT_RULES_NOTE".

---

## 5. Web app

### 5.1 Plumbing

- `web/src/services.ts`: add `api: CallApi` to `Services` (the `callApi` already created). Tests that build a fake `Services` get the extra field.
- `web/src/sync/engine.ts` `syncReceipts`: `hint: error.code === "invalid_config" ? "variables" : ...`. `web/src/App.tsx`: when `notice.hint === "variables"` render an "Open Customize variables" button that opens the dialog (state `variablesOpen`).
- `web/src/components/SettingsDialog.tsx`: new button "Customize variables" (disabled until credentials are saved, `title` explains) calling a new prop `onOpenVariables: () => void`. App opens `<CustomizeVariablesDialog services onClose />` (it closes Settings first on small screens; both can stack on desktop but keep it simple: close Settings, open Variables; closing Variables returns to the app).

### 5.2 Components (`web/src/components/variables/`)

- `CustomizeVariablesDialog.tsx` (container; `Dialog wide title="Customize variables"`):
  - On mount: `services.api<ReceiptRulesStateDto>({ action: "getReceiptRules" })`. States: `loading | loaded | loadError`. Holds `candidate: ReceiptRulesConfig` (starts from `config` or, when invalid, from defaults with the server issues shown and the raw JSON prefilled in the JSON tab), `raw` (for `expectedRaw`), `dirty`, `emails: ReceiptEmailDto[]`, `fetchState`, `saveState`.
  - `useMemo(() => validateReceiptRulesConfig(candidate, { forwardEmailAddress }))` => `issues` shown live; `compiled` only when ok.
  - `useMemo(() => previewReceiptRules(compiled, emails, context))` (Section 5.3).
  - Layout: on `< lg` a `role="tablist"` with tabs "Rules", "Emails", "Results"; on `lg:` two columns (Rules | Emails + Results stacked), using Tailwind only. Footer bar (sticky bottom): "Reset to built-in defaults" (secondary; sets candidate to defaults, dirty), "Save to sheet" (primary, disabled when `issues.length > 0` or busy or not dirty), status/alert text with `role="status"`/`role="alert"`. Unsaved guard via existing `useUnsavedGuard`/`UnsavedPrompt` from `ui.tsx` when closing dirty.
  - Save: `saveReceiptRules` with `{ rules: candidate, expectedRaw: raw }`; on `conflict` show message + "Reload" button; on `bad_request` with `issues` list them (path + message); on success replace `raw/config`, `dirty=false`, status "Saved to the Variables sheet".
- `RulesEditor.tsx` (props: `config`, `onChange(config)`, `issues`, `forwardEmailAddress`, `forwardName`): header card showing read-only `ForwardEmailAddress`/`ForwardName` ("edit these in the sheet") and an editable `defaultTypePrefix`. Sub-tabs "Form" / "JSON".
  - `SenderCard.tsx`: address, type, typePrefix (radio: "Fixed" / "Prefix by recipient"), forwardable checkbox, rules list, "Add rule", "Remove sender" (ConfirmDialog). "Add sender" at the bottom.
  - `RuleCard.tsx`: id, subject pattern (+ flags), optional body (checkbox "Also match the body" => source select + pattern + flags), name, nameFallback, cost (checkbox "Has a cost" => source select, group text, negate checkbox), category. Each field `<label>`ed; issue text for that field (match by path prefix) under it with `aria-describedby`. Monospace `font-mono` for patterns/templates.
  - `ForwardingEditor.tsx`: collapsible `<details>` "Forwarded-email detection (advanced)": fromPatterns list, bodyMarkers list.
  - JSON tab: `<textarea aria-label="Rules JSON" className="font-mono">` + "Apply JSON" button; parse errors (SyntaxError message) and validation issues listed under it. "Copy" button optional.
- `EmailFetcher.tsx` (props: `senders` (from candidate), `forwardEmailAddress`, `onFetched(emails, query)`, `api`): shows the senders as chips (read-only) and inputs: "Extra sender addresses" (comma-separated, validated with `isStrictEmailAddress`), "Subject contains" (up to 5, `isSafeSubjectFilter`), checkboxes "Include already-processed emails", "Include HTML bodies", number "Max threads" (1-50, default 20), optional "Newer than (days)". Query preview `<code>` built client-side with `buildReceiptSearchQuery` (same function the server uses; the server's `query` is displayed after fetch). "Fetch emails" button; `role="status"` "Fetched N emails from M threads" / `role="alert"` errors. Clear statement: "Fetching never labels or archives anything."
- `EmailResults.tsx` (props: `items: PreviewItem[]`): list of `<article>` cards: badge (Receipt / Not a receipt / Error) with color + text, From/To/Subject/Date, matched rule id and sender, captures as a `<dl>` (subject.* and body.*), resulting receipt (name, cost formatted, type, category), warnings, error message. `<details>` "Plain body" and "HTML body (source)" each containing `<pre className="whitespace-pre-wrap break-words text-xs">` with the text as a React child (never `dangerouslySetInnerHTML`). Truncation flags shown. Empty state copy: "Fetch emails to test the rules."
- Mobile: single column, 16px gutters, inputs `min-h-11`, cards stack; results list uses `max-h` scroll inside the dialog (dialog already `max-h-[92vh] overflow-y-auto`).

### 5.3 `web/src/lib/receipt-rules-preview.ts` (pure)

```ts
export interface PreviewItem {
  email: ReceiptEmailDto;
  outcome: RuleOutcome | null; // null when `error`
  error: string | null; // Error message thrown by applyReceiptRules
}
export function previewReceiptRules(
  rules: CompiledReceiptRules | null,
  emails: ReceiptEmailDto[],
  context: RuleContext,
): PreviewItem[];
export function toRuleEmail(dto: ReceiptEmailDto): RuleEmail;
export function formatOutcome(item: PreviewItem): {
  label: "Receipt" | "Not a receipt" | "Error";
  tone: "ok" | "muted" | "error";
};
```

### 5.4 Client-side e2e expectations (Section 8, Task C6)

---

## 6. Caching, locking, limits (reference card)

- Per-execution memo: `receipt-rules.util.ts` cache keyed by raw text; `bustVariablesCache()` after save makes the next `getVariables()` re-read, which changes `raw`, which recompiles.
- Locks: `saveReceiptRules` and `syncReceipts` share the script lock (`tryWithScriptLock`, `busy`). `fetchReceiptEmails` and `getReceiptRules` take no lock.
- Limits: `MAX_RECEIPT_RULES_CHARS 40_000`; senders <= 25; rules/sender <= 25; pattern <= 500; template <= 200; fromPatterns <= 10; bodyMarkers <= 10; fetch: threads 1-50 (default 20), messages <= 100, plain body <= 20_000 chars, html <= 60_000 chars, total <= 3_000_000 chars; subjects <= 5; extra senders total <= 20.

---

## 7. Characterization strategy (keeping behavior identical)

1. Task B1 adds `receiptEdgeFixtures` to `gas/test/fixtures/receipts.ts` and runs them through the EXISTING parser in `gas/test/receipts.test.ts` (same `describe.each` loop plus a 4th group `["Edge cases", /^edge/]`). These pass against the old code with no production change, pinning the fallback branches the main fixtures do not cover:
   - `edgeChaseRefundNoMerchant`: refund subject, body "no merchant line" => `{ cost: -45, name: undefined, type: "Credit" }`.
   - `edgeChaseGasNoMerchant`: gas subject, body "nothing" => `null`.
   - `edgePaypalSentNoBody`: sent subject, body "nothing" => `{ cost: 25.5, name: undefined, type: "TJ Paypal" }`.
   - `edgePaypalReceivedNoBody`: "Money is waiting for you", body "nothing" => `null`.
   - `edgeVenmoPaidNoBody`: "You paid Jane Doe $15.00", body "" => `{ cost: 15, name: "Jane Doe" }`.
   - `edgeVenmoToForwarder`: Venmo paid with `to: "Sam <fwd@example.com>"` => `{ type: "Sam Venmo" }`.
   - `edgeVenmoToUnparsable`: `to: "undisclosed-recipients"` => `{ type: "TJ Venmo" }`.
   - `edgeForwardedChase`: from FWD, subject "Fwd: You made a $1.00 transaction with X", body with Forwarded header `From: Chase <no.reply.alerts@chase.com>` => `null` (Chase not forwardable).
   - `edgeForwardedTentativeFrom`: forwarded header `From: service@paypal.com` (bare address) + PayPal body => `{ cost: 25.5, type: "Sam Paypal" }`.
   - `edgeCostNotNumber`: Chase "You made a $abc transaction with Shop" => `{ name: "Shop" }` and `cost` is `NaN` (assert with `Number.isNaN`; add a `costIsNaN?: true` key to `ExpectedReceipt` handled by the test loop).
   - `edgeZeroCostNoName`: Chase refund "$0.00" with no merchant => `null` (cost -0 is falsy).
     The error-message tests already in `receipts.test.ts` ("No function to handle...", "Could not get email address from nobody", forwarded-from error, unknown forwarded sender => `[]`) stay as they are.
2. Task A2 implements the engine; `gas/test/receipt-rules-engine.test.ts` runs EVERY fixture (`receiptFixtures` + `receiptEdgeFixtures`) through `applyReceiptRules(compileReceiptRules(DEFAULT_RECEIPT_RULES), toRuleEmail(fixture.email), { forwardEmailAddress: "fwd@example.com", forwardName: "Sam" })` and asserts the same `expected` (null => `isReceipt === false`). It also asserts the three pinned error messages and that the defaults' query equals the pinned query string.
3. Task B3 switches `ThreadList`; all of `receipts.test.ts`, `budgeting.test.ts`, `api.test.ts`, `index.test.ts`, `demo-backend.test.ts` must pass unchanged. A parity test in `receipt-rules-engine.test.ts` compares `ThreadList` output to direct engine output for every fixture (cost/name/type/category/isReceipt); it stays after the old code is gone, because `ThreadList` remains the adapter.
4. Task B4 deletes the old regex code. Coverage of `thread-list.model.ts` must not drop below the current figure for the remaining lines (check `pnpm --filter @budget/scripts-gas run test:coverage`).

Tests whose expectations change on purpose (call out in reports): `gas/test/variables.test.ts` "ignores variables beyond row 10" (=> beyond row 30; move the 11th row to row 31) and the error message `in the range Variables!A1:B10!` (=> `A1:B30`); any test pinning `Variables!A1:B10` in `sheet-variables-setup.test.ts` (grep).

---

## 8. Ordered task list

Legend: [S] Sonnet, [O] Opus. "Red" = write the failing tests first; "Green" = make them pass with the smallest change. Every task ends with the package's `pnpm --filter <pkg> run test` and `typecheck` green and `pnpm run format:check`/`lint` on touched files. The final task runs the full AGENTS.md section 5.6 set.

### Group 1 (parallel; disjoint files)

**A1 [O] Shared types, defaults, validation.**
Files: `shared/src/receipt-rules.ts` (types, `validateReceiptRulesConfig`, `parseReceiptRulesText`, `compileReceiptRules`, `expandTemplate`, `extractEmailAddress`), `shared/src/receipt-rules-defaults.ts` (`DEFAULT_RECEIPT_RULES`, `RECEIPT_RULES_NOTE`, `MAX_RECEIPT_RULES_CHARS`), `shared/src/index.ts` exports, `shared/test/receipt-rules-validate.test.ts`.
Red tests: defaults validate ok and `isDefault` for "", " ", null; each rule in Section 2.5 produces an issue with the exact `path` (one test per rule category: bad version, bad flags "gi", invalid regex "(", missing `(?<email>`, template referencing `${body.x}` without body, cost group not declared, duplicate address, address equal to forwardEmailAddress, unknown key, oversize JSON, unsafe `type` "=x"); `compileReceiptRules` throws the exact invalid-config message; `expandTemplate` returns undefined on a missing group and expands literal `$` untouched; `extractEmailAddress` message. Acceptance: all green; `DEFAULT_RECEIPT_RULES` deep-equals the JSON in Section 2.4.

**A2 [O] Engine `applyReceiptRules` + fixture parity.** (depends on A1 types; same agent as A1 continues)
Files: `shared/src/receipt-rules.ts` (apply), `shared/test/receipt-rules-apply.test.ts` (inline small emails for: forwarded path with each of the 3 markers and both fromPatterns, forwardable=false, unknown sender throws, prefix logic incl. empty forwardName fallthrough, htmlBody source + missing html warning, nameFallback, negate, parseFloat semantics "1,234.56" => 1, category only when isReceipt), `gas/test/receipt-rules-engine.test.ts` (fixture-driven; needs B1's edge fixtures: if B1 is not merged yet, write the test to iterate `receiptFixtures` only and extend in B3).
Acceptance: every fixture expected matches; captures populated; `isReceipt` semantics.

**A3 [S] Query builder.**
Files: `shared/src/receipt-query.ts`, `shared/src/index.ts`, `shared/test/receipt-query.test.ts`.
Red tests: default senders + "fwd@example.com" => the exact pinned string (copy from `gas/test/budgeting.test.ts:156`); invalid address throws "Invalid sender address: x"; subjects emitted quoted; `includeProcessed` omits inbox/exclusions; `newerThanDays`; `isSafeSubjectFilter` rejects `"`, `(`, `:`; `isStrictEmailAddress` rejects spaces/parens/"OR".

**B1 [S] Edge fixtures pinned on the OLD parser.**
Files: `gas/test/fixtures/receipts.ts` (`receiptEdgeFixtures`, `ExpectedReceipt.costIsNaN?`), `gas/test/receipts.test.ts` (4th group; `costIsNaN` handling).
Acceptance: all new fixtures pass with NO `src/` change. If one does not, the plan's expectation is wrong: report the actual old behavior and pin THAT (the old code is the oracle).

**B2 [S] Variables range + optional `ReceiptRules` + `receipt-rules.util.ts`.** (needs A1 published types; can start on the variables part first)
Files: `gas/src/util/variables.util.ts`, `gas/src/util/sheet-variables-setup.util.ts` (import constants), `gas/src/util/receipt-rules.util.ts`, `gas/test/variables.test.ts` (update the two pinned expectations, add: `ReceiptRules` read when present, absent when missing), `gas/test/receipt-rules-util.test.ts` (defaults when missing/empty; compiled from valid JSON; invalid JSON throws the message; number cell throws; memo: second call returns same object; after `bustVariablesCache()` + cell change, recompiles; `getReceiptSearchQuery()` equals the pinned string; `toRuleEmail` calls `getBody` only when asked, needs `MockMessage.getBody` from B7 or add it here, minimal).
Acceptance: existing tests green except the two stated updates.

### Group 2 (after A1-A3, B1, B2)

**B3 [O] Switch `ThreadList` to the engine (parity).**
Files: `gas/src/models/thread-list.model.ts` (constructor params, per-message adapter; old code left in place but unused for this task), `gas/test/receipt-rules-engine.test.ts` (parity test ThreadList vs engine over all fixtures; invalid variable => `new ThreadList([])` throws before iterating).
Acceptance: full gas suite green with no expectation changes; parity test green.

**B4 [S] Delete the old parser code.**
Files: `gas/src/models/thread-list.model.ts` only. Remove `getUnicodeAnalysis` through `getReceiptInfoMap`, unused imports. Acceptance: suite green, lint clean, `dist` builds (`pnpm --filter @budget/scripts-gas run build`).

**B5 [S] Query via builder in `budgeting.util.ts`.**
Files: `gas/src/util/budgeting.util.ts` (`getTJGmailSearchQuery` => `getReceiptSearchQuery()`), `gas/test/budgeting.test.ts` (add: with a custom `ReceiptRules` containing one extra sender the search query lists it; invalid `ReceiptRules` makes `getChaseReceipts` throw before `GmailApp.search` is called, `searches` stays empty).
Acceptance: pinned query test unchanged and green.

**B6 [S] Setup helper for global variables.**
Files: `gas/src/util/global-variables-setup.util.ts`, `gas/src/util/a1.util.ts` (moved helpers) + `sheet-variables-setup.util.ts` imports, `gas/src/index.ts` export, `gas/test/global-variables-setup.test.ts` (adds key+defaults JSON+note into first fully-empty row; idempotent second run logs "Nothing to do" and leaves values/notes; skips a row with stray content in B; throws when A1:B30 has no empty row and changes nothing; `index.test.ts`-style check that the export exists in `src/index.ts`), `gas/test/sheet-variables-setup.test.ts` unchanged.
Acceptance: note text equals `GLOBAL_VARIABLE_NOTES.ReceiptRules`; value parses back to `DEFAULT_RECEIPT_RULES`.

### Group 3 (parallel after Group 2 starts; B7/B8 need B2, C-tasks need A1-A3 only)

**B7 [O] Mocks + `fetchReceiptEmails`.**
Files: `gas/test/gas-mocks/gmail.ts` (`html`, `getBody`, `getLabels`), `shared/src/types.ts` (DTOs, `invalid_config`, `issues`), `gas/src/api/api-error.ts` (`issues`), `gas/src/api/router.ts` (case + `issues` in `toFailure`), `gas/src/api/receipt-rules-handlers.ts` (`handleFetchReceiptEmails`), `gas/test/api-receipt-rules.test.ts` (via `doPost`: unauthorized; default senders build the pinned query and `GmailApp.searches[0]` is `{query, start: 0, max: 20}`; `senders` override; bad sender => bad_request message; `subjects` with `(` => bad_request; `includeHtml` returns `htmlBody`; truncation flags with a 25,000-char body; message cap 100 with `truncated: true`; labels listed; NO labels created/added and no archive after the call; invalid `ReceiptRules` => `invalid_config`).
Acceptance: all green; `handlers.ts` is not touched in this task.

**B8 [O] `getReceiptRules` / `saveReceiptRules` + `syncReceipts` invalid_config.**
Files: `gas/src/api/receipt-rules-handlers.ts`, `gas/src/api/router.ts`, `gas/src/api/handlers.ts` (syncReceipts mapping only), `gas/test/api-receipt-rules.test.ts` (get: missing key => `hasVariable false`, `isDefault true`, `config` = defaults, `query` pinned; invalid cell => `issues` non-empty, `config null`, `query null`; save valid => cell text is pretty JSON, key added when missing (reuse `ensureGlobalVariableRow` from B6), cache busted so a following `getReceiptRules` in the same module instance reflects it, lock `tryLock:20000`/`release` calls recorded; save invalid => `bad_request` with `issues[0].path`; save `null` => cell "" and `isDefault true`; `expectedRaw` mismatch => `conflict`; lock contended => `busy`; `syncReceipts` with invalid cell => `invalid_config`).

**B9 [S] Demo backend: realistic search, `__debug/variables`, html threads.**
Files: `gas/dev/backend.ts`, `gas/dev/server.ts`, `gas/test/demo-backend.test.ts` (query filters by from/subject/in:inbox; `dumpVariables`), `web/e2e/helpers.ts` (`backend.variables()`).

**C1 [S] Web plumbing.**
Files: `web/src/services.ts` (`api`), `web/src/sync/engine.ts` (hint "variables"), `web/src/App.tsx` (state + button + dialog mount placeholder), `web/src/components/SettingsDialog.tsx` (`onOpenVariables` button), tests: `web/src/sync/engine.test.ts` (invalid_config => hint "variables"), `web/src/components/SettingsDialog.test.tsx` (button present, disabled without saved credentials, calls prop), `web/src/App.test.tsx` fake services gains `api`.

**C2 [S] Preview lib.**
Files: `web/src/lib/receipt-rules-preview.ts`, `web/src/lib/receipt-rules-preview.test.ts` (uses `demoInbox` from the gas fixtures via a relative import like the gas tests do; a thrown engine error becomes `error`; null rules => items with `outcome null` and error "Rules are invalid"; `formatOutcome` labels).

**C3 [O] `RulesEditor` (form + JSON).**
Files: `web/src/components/variables/RulesEditor.tsx`, `SenderCard.tsx`, `RuleCard.tsx`, `ForwardingEditor.tsx`, `web/src/components/variables/RulesEditor.test.tsx` (renders defaults: 3 sender cards, rule ids visible; editing the Chase purchase name template calls `onChange` with the new template; toggling "Also match the body" adds/removes `body`; issue for `senders[0].rules[0].subject.pattern` appears under that field and is referenced by `aria-describedby`; JSON tab shows pretty JSON; "Apply JSON" with invalid JSON shows the SyntaxError text and does not call `onChange`; valid JSON calls `onChange`; "Add sender"/"Remove sender" (confirm) work; all inputs have labels (`getByLabelText`)).

**C4 [O] `EmailFetcher` + `EmailResults`.**
Files: `web/src/components/variables/EmailFetcher.tsx`, `EmailResults.tsx`, `web/src/components/variables/EmailFetcher.test.tsx` (query preview equals `buildReceiptSearchQuery` output; invalid extra sender shows inline error and disables Fetch; Fetch calls `api` with `{ action: "fetchReceiptEmails", senders: [...candidate + extras], includeHtml, includeProcessed, maxThreads }`; status text), `EmailResults.test.tsx` (badges; captures `<dl>`; a plain body containing `<script>alert(1)</script>` is rendered as text, `document.querySelector("script")` is null; truncation note; empty state).

**C5 [O] `CustomizeVariablesDialog` container.**
Files: `web/src/components/variables/CustomizeVariablesDialog.tsx`, `web/src/App.tsx` (mount), `web/src/components/variables/CustomizeVariablesDialog.test.tsx` (loads via fake `api`; invalid server state shows issues and prefilled JSON; live preview updates when a template changes; Save disabled while issues exist; Save sends `{ action: "saveReceiptRules", rules, expectedRaw }` and shows "Saved"; conflict shows Reload; Reset to defaults then Save sends defaults, and a second control "Clear variable (use built-in defaults)" sends `rules: null`; unsaved-close prompt; tabs have `role="tab"`/`aria-selected` and work on a 390px viewport (jsdom: just assert the tablist exists and switching changes the visible panel)).

### Group 4 (after everything)

**C6 [S] E2E.**
Files: `web/e2e/variables.spec.ts`. Flow: `connect` -> Settings -> "Customize variables" -> dialog visible -> Emails tab -> "Fetch emails" -> 5 cards (demo inbox) with "Demo Cafe" marked Receipt and "Your statement is ready" marked Not a receipt -> Rules tab -> change `chase-purchase` name template to `${subject.name}!` -> Results shows "Demo Cafe!" -> Save -> `backend.variables()` has `ReceiptRules` containing `"${subject.name}!"` -> close -> "Sync receipts" -> row named "Demo Cafe!" exists in `backend.dump(sheetName)`. Second test: paste invalid JSON in the JSON tab -> SyntaxError shown, candidate unchanged; then paste valid JSON with a bad regex `(` -> issue listed with its path, Save disabled. As built, the e2e suite does not run the JSON-tab paste step; that step is unit-tested only (`web/src/components/variables/RulesEditor.test.tsx`). (Server-side rejection of an invalid cell is unit-covered in B8; no debug endpoint writes the Variables sheet, so it is not e2e-tested.)

**D1 [S] Docs + full verification.**
Files: `AGENTS.md` section 0 (new files: `shared/src/receipt-rules*.ts`, `receipt-query.ts`, `gas/src/util/receipt-rules.util.ts`, `global-variables-setup.util.ts`, `gas/src/api/receipt-rules-handlers.ts`, `web/src/components/variables/`, `__debug/variables`), section 3.4 one sentence ("Receipt parsing is driven by the `ReceiptRules` global variable, defaults in `@budget/shared`"). Run `pnpm run format:check`, `lint`, `typecheck`, `test`, `build`, `test:e2e`; confirm `web/dist` contains no dev URL/key.

Parallelism: Group 1 = up to 3 agents (A1+A2 | A3+B1 | B2). Group 2 is sequential (B3 -> B4 -> B5 -> B6) by one agent while another starts C1-C2. Group 3 = 3 agents (B7+B8 | B9+C1+C2 | C3+C4 then C5). C6 and D1 last.

---

## 9. Risks

- `parseFloat` semantics and `NaN`/`-0` edge behaviors are preserved on purpose; do not "fix" them in this feature.
- Case-insensitive sender compare is a tiny intentional behavior change (never makes a parse fail that works today).
- Case-insensitive `To`-address compare against `ForwardEmailAddress` (`recipientPrefix()`) is a second tiny intentional change. It only differs from today when the `To` address matches the forward address in case alone. Do not change it back to exact compare.
- Forwarded subject regex: it is built as `<prefix>(?:<pattern without ^>)`, so a top-level alternation in a custom subject pattern (e.g. `^a|b`) still requires the prefix for every branch. The old `^Fwd: <pattern>` let the `b` branch match without the prefix. The default patterns have no top-level alternation, so defaults are unchanged.
- Expanding the Variables read range to row 30 picks up any stray `A11:A30` cell whose text equals a variable name; today rows 8-10 behave the same way. Document in the note.
- `ForwardEmailAddress` must satisfy `STRICT_EMAIL_REGEX` or the import now fails loudly (today it would produce a broken Gmail query silently). Message makes the cause clear.
- Regex cost on GAS: the owner can save a slow pattern. Mitigation: size caps, TSDoc warning, preview in the browser first. No server-side execution of unsaved candidates.
- Response size of `fetchReceiptEmails` with `includeHtml`: capped at ~3 MB; GAS and the client handle that.
- `syncReceipts` and `saveReceiptRules` share the lock; a save during an import answers `busy`, which the UI already understands.
- Shared package tests must not import GAS globals; fixture-driven engine tests therefore live in the GAS package.

## 10. Open questions for the owner (with recommended defaults)

1. Setup writes the defaults JSON into the cell (your example) vs. writing an empty value (built-in defaults keep flowing after code updates). Recommended: write the defaults JSON as planned, since the UI shows "differs from built-in defaults" and offers "Clear variable".
2. Who may save rules: there is one shared API key, so anyone with it (partner included) can save. Recommended: accept this; no roles in this feature.
3. `fetchReceiptEmails` with `includeProcessed` searches ALL mail from those senders (read-only). Recommended: keep it, default off, clearly labeled.
4. Sender comparison case-insensitive (recommended) vs. exact as today.
5. Should the `Fwd: ` subject prefix and the formal-address regex be configurable now? Recommended: no (constants), add later if a real need appears.

## 11. Owner decisions (binding; override anything above that conflicts)

1. The setup helper writes the pretty-printed `DEFAULT_RECEIPT_RULES` JSON into the `ReceiptRules` value cell when the key is missing (as recommended).
2. Anyone with the API key may save rules (no roles), as recommended.
3. `includeProcessed` stays: optional, default off, clearly labeled, read-only. ALSO: the fetcher UI has a "Max emails" number input (default **10**), sent as `maxThreads`. The API default `maxThreads` is 10; the server still clamps it to its cap.
4. Sender address comparison is case-insensitive (intentional small change; pin it with a test and call it out).
5. Forwarded-subject prefix becomes configurable: add `forwarding.subjectPrefixes: string[]` (default `["Fwd: "]`; 1-5 entries, each 1-20 chars, literal text, regex-escaped when building the forward subject regex). A forwarded email's subject matches when it starts with any prefix followed by the rule's subject pattern without its leading `^`. The default must reproduce today's behavior exactly. The formal-address extraction regex (`Name <email>` → `email`) stays in code (it is a no-op on plain addresses).
