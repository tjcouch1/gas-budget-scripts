/**
 * Data-driven receipt email parsing (see receipt-rules-plan.md section 2). Pure TS: runs in GAS and
 * the browser. Regexes are owner-supplied; slow (ReDoS-prone) patterns are the owner's
 * responsibility and slow down every import.
 */
import {
  DEFAULT_RECEIPT_RULES,
  MAX_RECEIPT_RULES_CHARS,
} from "./receipt-rules-defaults.ts";
import { isUnsafeFreeText } from "./names.ts";
import {
  DEFAULT_FORWARD_SUBJECT_PREFIXES,
  forwardPrefixToken,
  isStrictEmailAddress,
} from "./receipt-query.ts";

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
   * For a forwarded email the subject must start with one of `forwarding.subjectPrefixes` followed by
   * this pattern without its leading "^". The pattern is wrapped in a group, so a top-level
   * alternation such as `^a|b` needs the prefix for every alternative.
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
   * rules (subject must start with a forwarding subject prefix). false: such forwards are not receipts (they only get a note).
   */
  forwardable: boolean;
  /** Evaluated in order; first subject match wins. May be empty (sender is known, nothing is ever a receipt). */
  rules: ReceiptRule[];
}

export interface ForwardingConfig {
  /**
   * Literal subject prefixes of forwarded emails (regex-escaped when used). 1-5 entries of 1-20 chars,
   * each letters/digits plus an optional trailing ":" and whitespace (that word is the Gmail query's
   * `subject:` token, see forwardPrefixToken). Optional; default ["Fwd: "]. A message from
   * ForwardEmailAddress whose subject starts with none of them is "not a receipt" (not-a-forward).
   */
  subjectPrefixes?: string[];
  /**
   * Regexes with a named group `email`, run in order on the forwarded email's whole plain body to find the original
   * sender ("From: PayPal <service@paypal.com>"). The group may hold a formal address; the address is extracted.
   * None matching (the subject has a forward prefix) => error "Could not get 'from' email address from forwarded email with subject '<subject>'".
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

/** One validation problem; path like "senders[1].rules[0].subject.pattern" ("(root)" for the whole value) */
export interface RuleIssue {
  path: string;
  message: string;
}

export type ValidationResult =
  { ok: true; config: ReceiptRulesConfig } | { ok: false; issues: RuleIssue[] };

export interface ValidationOptions {
  /** When given, no sender may use this address (compared case-insensitively) */
  forwardEmailAddress?: string;
}

// #region validation

const ROOT_PATH = "(root)";
const MAX_ISSUES = 50;
const FLAGS_REGEX = /^(?!.*(.).*\1)[imsu]*$/;
const RULE_ID_REGEX = /^[A-Za-z0-9_-]{1,40}$/;
const GROUP_NAME_SCAN = /\(\?<([A-Za-z_][A-Za-z0-9_]*)>/g;
/** Anything that looks like a placeholder; each must be a valid TEMPLATE_PLACEHOLDER */
const PLACEHOLDER_SCAN = /\$\{([^}]*)\}/g;
const PLACEHOLDER_BODY = /^(subject|body)\.([A-Za-z_][A-Za-z0-9_]*)$/;
const TEMPLATE_PLACEHOLDER = /\$\{(subject|body)\.([A-Za-z_][A-Za-z0-9_]*)\}/g;
const UNSAFE_START_MESSAGE = "must not start with =, +, - or @";

type Obj = Record<string, unknown>;

class IssueList {
  readonly list: RuleIssue[] = [];
  add(path: string, message: string) {
    if (this.list.length < MAX_ISSUES) this.list.push({ path, message });
  }
}

function isObj(value: unknown): value is Obj {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function child(path: string, key: string | number) {
  if (typeof key === "number") return `${path}[${key}]`;
  return path === ROOT_PATH ? key : `${path}.${key}`;
}

function checkKeys(
  value: Obj,
  path: string,
  allowed: readonly string[],
  issues: IssueList,
) {
  for (const key of Object.keys(value))
    if (!allowed.includes(key)) issues.add(child(path, key), "unknown key");
}

/** Checks a required object; returns it or undefined (issue added) */
function requireObj(
  value: unknown,
  path: string,
  issues: IssueList,
): Obj | undefined {
  if (isObj(value)) return value;
  issues.add(path, value === undefined ? "is required" : "must be an object");
  return undefined;
}

function checkArray(
  value: unknown,
  path: string,
  min: number,
  max: number,
  what: string,
  issues: IssueList,
): unknown[] | undefined {
  if (Array.isArray(value) && value.length >= min && value.length <= max)
    return value;
  issues.add(path, `must be a list of ${min}-${max} ${what}`);
  return undefined;
}

interface TextRules {
  max: number;
  /** reject text starting with a formula character */
  safe?: boolean;
  noPipe?: boolean;
  /** reject whitespace-only text */
  nonBlank?: boolean;
}

function checkText(
  value: unknown,
  path: string,
  rules: TextRules,
  issues: IssueList,
): boolean {
  if (
    typeof value !== "string" ||
    value.length < 1 ||
    value.length > rules.max ||
    (rules.nonBlank && value.trim() === "")
  ) {
    issues.add(path, `must be text of 1-${rules.max} characters`);
    return false;
  }
  if (rules.noPipe && value.includes("|")) {
    issues.add(path, 'must not contain "|"');
    return false;
  }
  if (rules.safe && isUnsafeFreeText(value)) {
    issues.add(path, UNSAFE_START_MESSAGE);
    return false;
  }
  return true;
}

/** Named group names declared in a pattern */
function declaredGroups(pattern: string): Set<string> {
  return new Set([...pattern.matchAll(GROUP_NAME_SCAN)].map((m) => m[1]));
}

/**
 * Validates a RegexSpec (plus `extraKeys`, checked by the caller). Returns the declared group names
 * when the regex is valid, else undefined.
 */
function checkRegexSpec(
  value: unknown,
  path: string,
  issues: IssueList,
  extraKeys: readonly string[] = [],
): Set<string> | undefined {
  const spec = requireObj(value, path, issues);
  if (!spec) return undefined;
  checkKeys(spec, path, ["pattern", "flags", ...extraKeys], issues);
  const patternOk = checkText(
    spec.pattern,
    child(path, "pattern"),
    { max: 500 },
    issues,
  );
  let flagsOk = true;
  if (
    spec.flags !== undefined &&
    (typeof spec.flags !== "string" || !FLAGS_REGEX.test(spec.flags))
  ) {
    flagsOk = false;
    issues.add(
      child(path, "flags"),
      "may only contain i, m, s and u, each at most once",
    );
  }
  if (!patternOk || !flagsOk) return undefined;
  const pattern = spec.pattern as string;
  try {
    new RegExp(pattern, (spec.flags as string | undefined) ?? "");
  } catch (e) {
    issues.add(
      child(path, "pattern"),
      e instanceof Error ? e.message : String(e),
    );
    return undefined;
  }
  return declaredGroups(pattern);
}

/** Group names available to templates/cost; undefined entry = regex invalid (skip reference checks) */
interface RuleGroups {
  subject: Set<string> | undefined;
  /** null: the rule has no body */
  body: Set<string> | undefined | null;
}

function checkTemplate(
  value: unknown,
  path: string,
  groups: RuleGroups,
  issues: IssueList,
) {
  if (!checkText(value, path, { max: 200 }, issues)) return;
  const template = value as string;
  if (template.replace(PLACEHOLDER_SCAN, "").includes("${"))
    issues.add(path, "has an unclosed ${ placeholder");
  for (const [placeholder, inner] of template.matchAll(PLACEHOLDER_SCAN)) {
    const m = PLACEHOLDER_BODY.exec(inner);
    if (!m) {
      issues.add(
        path,
        `${placeholder} is not a placeholder; use \${subject.<group>} or \${body.<group>}`,
      );
      return;
    }
    const source = m[1] as "subject" | "body";
    const known = groups[source];
    if (known === null) {
      issues.add(path, `${placeholder} is used but the rule has no body`);
      return;
    }
    if (known && !known.has(m[2])) {
      issues.add(
        path,
        `${placeholder}: the ${source} pattern has no group named "${m[2]}"`,
      );
      return;
    }
  }
  const literalStart = template.split("${")[0];
  if (isUnsafeFreeText(literalStart)) issues.add(path, UNSAFE_START_MESSAGE);
}

function checkRule(
  value: unknown,
  path: string,
  seenIds: Set<string>,
  issues: IssueList,
) {
  const rule = requireObj(value, path, issues);
  if (!rule) return;
  checkKeys(
    rule,
    path,
    ["id", "subject", "body", "name", "nameFallback", "cost", "category"],
    issues,
  );
  if (typeof rule.id !== "string" || !RULE_ID_REGEX.test(rule.id))
    issues.add(
      child(path, "id"),
      "must be 1-40 letters, digits, underscores or dashes",
    );
  else if (seenIds.has(rule.id))
    issues.add(child(path, "id"), "duplicate rule id within this sender");
  else seenIds.add(rule.id);

  const groups: RuleGroups = {
    subject: checkRegexSpec(rule.subject, child(path, "subject"), issues),
    body: null,
  };
  if (rule.body !== undefined) {
    const bodyPath = child(path, "body");
    groups.body = checkRegexSpec(rule.body, bodyPath, issues, ["source"]);
    if (
      isObj(rule.body) &&
      rule.body.source !== "plainBody" &&
      rule.body.source !== "htmlBody"
    )
      issues.add(
        child(bodyPath, "source"),
        'must be "plainBody" or "htmlBody"',
      );
  }

  for (const key of ["name", "nameFallback"] as const)
    if (rule[key] !== undefined)
      checkTemplate(rule[key], child(path, key), groups, issues);

  if (rule.cost !== undefined) {
    const costPath = child(path, "cost");
    const cost = requireObj(rule.cost, costPath, issues);
    if (cost) {
      checkKeys(cost, costPath, ["source", "group", "negate"], issues);
      let known: Set<string> | undefined | null = undefined;
      if (cost.source !== "subject" && cost.source !== "body")
        issues.add(child(costPath, "source"), 'must be "subject" or "body"');
      else if (cost.source === "body" && groups.body === null)
        issues.add(
          child(costPath, "source"),
          "is body but the rule has no body",
        );
      else known = groups[cost.source];
      if (typeof cost.group !== "string" || cost.group === "")
        issues.add(child(costPath, "group"), "must be a group name");
      else if (known && !known.has(cost.group))
        issues.add(
          child(costPath, "group"),
          `the ${String(cost.source)} pattern has no group named "${cost.group}"`,
        );
      if (cost.negate !== undefined && typeof cost.negate !== "boolean")
        issues.add(child(costPath, "negate"), "must be true or false");
    }
  }

  if (rule.category !== undefined)
    checkText(
      rule.category,
      child(path, "category"),
      { max: 60, safe: true },
      issues,
    );
}

function checkForwarding(value: unknown, path: string, issues: IssueList) {
  const forwarding = requireObj(value, path, issues);
  if (!forwarding) return;
  checkKeys(
    forwarding,
    path,
    ["subjectPrefixes", "fromPatterns", "bodyMarkers"],
    issues,
  );

  if (forwarding.subjectPrefixes !== undefined) {
    const prefixesPath = child(path, "subjectPrefixes");
    const prefixes = checkArray(
      forwarding.subjectPrefixes,
      prefixesPath,
      1,
      5,
      "prefixes",
      issues,
    );
    prefixes?.forEach((prefix, i) => {
      if (
        checkText(prefix, child(prefixesPath, i), { max: 20 }, issues) &&
        forwardPrefixToken(prefix as string) === undefined
      )
        issues.add(
          child(prefixesPath, i),
          'must be letters or digits, optionally followed by ":" and spaces (e.g. "Fwd: ")',
        );
    });
  }

  const fromPath = child(path, "fromPatterns");
  checkArray(
    forwarding.fromPatterns,
    fromPath,
    1,
    10,
    "patterns",
    issues,
  )?.forEach((spec, i) => {
    const groups = checkRegexSpec(spec, child(fromPath, i), issues);
    if (groups && !groups.has("email"))
      issues.add(
        child(child(fromPath, i), "pattern"),
        "must contain a named group (?<email>...)",
      );
  });

  const markersPath = child(path, "bodyMarkers");
  checkArray(
    forwarding.bodyMarkers,
    markersPath,
    0,
    10,
    "markers",
    issues,
  )?.forEach((value, i) => {
    const markerPath = child(markersPath, i);
    const marker = requireObj(value, markerPath, issues);
    if (!marker) return;
    checkKeys(marker, markerPath, ["marker", "skipLines"], issues);
    checkText(
      marker.marker,
      child(markerPath, "marker"),
      { max: 100, nonBlank: true },
      issues,
    );
    const skip = marker.skipLines;
    if (
      typeof skip !== "number" ||
      !Number.isInteger(skip) ||
      skip < 0 ||
      skip > 20
    )
      issues.add(child(markerPath, "skipLines"), "must be a whole number 0-20");
  });
}

function checkSender(
  value: unknown,
  path: string,
  seenAddresses: Set<string>,
  forwardAddress: string,
  issues: IssueList,
) {
  const sender = requireObj(value, path, issues);
  if (!sender) return;
  checkKeys(
    sender,
    path,
    ["address", "type", "typePrefix", "forwardable", "rules"],
    issues,
  );
  const addressPath = child(path, "address");
  if (!isStrictEmailAddress(sender.address))
    issues.add(
      addressPath,
      "must be a plain email address like name@example.com",
    );
  else {
    const lower = sender.address.toLowerCase();
    if (seenAddresses.has(lower))
      issues.add(
        addressPath,
        "duplicate sender address (addresses are compared case-insensitively)",
      );
    else if (forwardAddress && lower === forwardAddress)
      issues.add(
        addressPath,
        "must not be the ForwardEmailAddress (forwarded emails are handled by the forwarding settings)",
      );
    seenAddresses.add(lower);
  }
  checkText(
    sender.type,
    child(path, "type"),
    { max: 40, safe: true, noPipe: true },
    issues,
  );
  if (sender.typePrefix !== "fixed" && sender.typePrefix !== "recipient")
    issues.add(child(path, "typePrefix"), 'must be "fixed" or "recipient"');
  if (typeof sender.forwardable !== "boolean")
    issues.add(child(path, "forwardable"), "must be true or false");
  const rulesPath = child(path, "rules");
  const seenIds = new Set<string>();
  checkArray(sender.rules, rulesPath, 0, 25, "rules", issues)?.forEach(
    (rule, i) => checkRule(rule, child(rulesPath, i), seenIds, issues),
  );
}

/** Structural + semantic validation of parsed JSON (never throws). Collects up to 50 issues. */
export function validateReceiptRulesConfig(
  value: unknown,
  options: ValidationOptions = {},
): ValidationResult {
  const issues = new IssueList();
  const root = isObj(value) ? value : undefined;
  if (!root) issues.add(ROOT_PATH, "must be a JSON object");
  else {
    checkKeys(
      root,
      ROOT_PATH,
      ["version", "defaultTypePrefix", "forwarding", "senders"],
      issues,
    );
    if (root.version !== RECEIPT_RULES_VERSION)
      issues.add("version", `must be ${RECEIPT_RULES_VERSION}`);
    checkText(
      root.defaultTypePrefix,
      "defaultTypePrefix",
      { max: 20, safe: true, noPipe: true },
      issues,
    );
    checkForwarding(root.forwarding, "forwarding", issues);
    const seenAddresses = new Set<string>();
    const forwardAddress = (options.forwardEmailAddress ?? "").toLowerCase();
    checkArray(root.senders, "senders", 1, 25, "senders", issues)?.forEach(
      (sender, i) =>
        checkSender(
          sender,
          child("senders", i),
          seenAddresses,
          forwardAddress,
          issues,
        ),
    );
    const size = JSON.stringify(root, null, 2).length;
    if (size > MAX_RECEIPT_RULES_CHARS)
      issues.add(
        ROOT_PATH,
        `is too long (${size} characters; the limit is ${MAX_RECEIPT_RULES_CHARS})`,
      );
  }
  return issues.list.length === 0
    ? { ok: true, config: value as ReceiptRulesConfig }
    : { ok: false, issues: issues.list };
}

/** JSON.parse + validate. Empty/whitespace text => { ok: true, config: DEFAULT_RECEIPT_RULES, isDefault: true }. */
export function parseReceiptRulesText(
  text: string | null | undefined,
  options: ValidationOptions = {},
): ValidationResult & { isDefault?: boolean } {
  if (text === null || text === undefined || text.trim() === "")
    return {
      ok: true,
      config: JSON.parse(JSON.stringify(DEFAULT_RECEIPT_RULES)),
      isDefault: true,
    };
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch (e) {
    return {
      ok: false,
      issues: [
        {
          path: ROOT_PATH,
          message: `Invalid JSON: ${e instanceof Error ? e.message : String(e)}`,
        },
      ],
    };
  }
  return validateReceiptRulesConfig(value, options);
}

// #endregion

// #region compile

export interface CompiledReceiptRules {
  config: ReceiptRulesConfig;
  /** true when any rule uses body.source === "htmlBody" (GAS then calls message.getBody()) */
  usesHtmlBody: boolean;
  /** lower-cased sender addresses in config order */
  senderAddresses: string[];
}

interface CompiledRule {
  rule: ReceiptRule;
  subject: RegExp;
  forwardSubject: RegExp;
  body: RegExp | undefined;
}

interface CompiledSender {
  sender: SenderRules;
  rules: CompiledRule[];
}

interface CompiledInternals {
  /** by lower-cased address */
  senders: Map<string, CompiledSender>;
  fromPatterns: RegExp[];
}

/** Compiled regexes, kept out of the public object so it stays plain data */
const internals = new WeakMap<CompiledReceiptRules, CompiledInternals>();

/** Escapes text for literal use in a RegExp (valid with and without the "u" flag) */
function escapeRegExp(text: string): string {
  return text.replace(/[\\^$.*+?()[\]{}|/]/g, "\\$&");
}

function toRegExp(spec: RegexSpec): RegExp {
  return new RegExp(spec.pattern, spec.flags ?? "");
}

/** Throws Error("ReceiptRules variable is invalid: <path>: <message>. ...") on the first issue. */
export function compileReceiptRules(
  config: ReceiptRulesConfig,
  options: ValidationOptions = {},
): CompiledReceiptRules {
  const result = validateReceiptRulesConfig(config, options);
  if (!result.ok) {
    const [{ path, message }] = result.issues;
    throw new Error(
      `ReceiptRules variable is invalid: ${path}: ${message}. Fix it in the web app (Settings > Customize variables) or clear the cell to use the built-in defaults.`,
    );
  }

  const prefixes =
    config.forwarding.subjectPrefixes ?? DEFAULT_FORWARD_SUBJECT_PREFIXES;
  const prefixSource = `^(?:${prefixes.map(escapeRegExp).join("|")})`;
  const senders = new Map<string, CompiledSender>();
  let usesHtmlBody = false;
  for (const sender of config.senders) {
    senders.set(sender.address.toLowerCase(), {
      sender,
      rules: sender.rules.map((rule) => {
        if (rule.body?.source === "htmlBody") usesHtmlBody = true;
        return {
          rule,
          subject: toRegExp(rule.subject),
          forwardSubject: new RegExp(
            `${prefixSource}(?:${rule.subject.pattern.replace(/^\^/, "")})`,
            rule.subject.flags ?? "",
          ),
          body: rule.body ? toRegExp(rule.body) : undefined,
        };
      }),
    });
  }

  const compiled: CompiledReceiptRules = {
    config,
    usesHtmlBody,
    senderAddresses: [...senders.keys()],
  };
  internals.set(compiled, {
    senders,
    fromPatterns: config.forwarding.fromPatterns.map(toRegExp),
  });
  return compiled;
}

// #endregion

// #region helpers

/** An address inside angle brackets: `Name <a@b.com>` */
const ANGLE_EMAIL_REGEX = /<\s*(?<email>[^<>\s\u00a0"']+@[^<>\s\u00a0"']+)\s*>/;
/** The first bare address token, without surrounding quotes */
const FORMAL_EMAIL_REGEX = /(?<email>[^<>\s\u00a0"']+@[^<>\s\u00a0"']+)/;

/**
 * Formal-address extraction. Prefers the address inside `<...>`, else the first `x@y` token, never
 * including quotes: `"service@paypal.com" <service@paypal.com>`, `"a@b.com"`, `Name <a@b.com>` and
 * `a@b.com` all give the plain address. Throws "Could not get email address from <text>"
 */
export function extractEmailAddress(formal: string): string {
  const email = (
    ANGLE_EMAIL_REGEX.exec(formal) ?? FORMAL_EMAIL_REGEX.exec(formal)
  )?.groups?.email;
  if (email === undefined)
    throw new Error(`Could not get email address from ${formal}`);
  return email;
}

export type TemplateGroups = {
  subject: Record<string, string | undefined> | null;
  body: Record<string, string | undefined> | null;
};

/** Expands ${subject.x}/${body.x}; returns undefined if any referenced group is undefined. Other text is literal. */
export function expandTemplate(
  template: string,
  groups: TemplateGroups,
): string | undefined {
  let missing = false;
  const text = template.replace(
    TEMPLATE_PLACEHOLDER,
    (_match, source: "subject" | "body", group: string) => {
      const value = groups[source]?.[group];
      if (value === undefined) {
        missing = true;
        return "";
      }
      return value;
    },
  );
  return missing ? undefined : text;
}

// #endregion

// #region apply

export interface RuleEmail {
  /** message.getFrom(): "Chase <no.reply.alerts@chase.com>" or a bare address */
  from: string;
  /** message.getTo() */
  to: string;
  subject: string;
  plainBody: string;
  /** only needed when compiled.usesHtmlBody */
  htmlBody?: string;
}

export interface RuleContext {
  forwardEmailAddress: string;
  forwardName: string;
}

export type RuleCaptures = Record<string, string | undefined>;

/**
 * Why an email is not a receipt:
 * - "unknown-sender": no sender has that address (the original sender, for a forward)
 * - "not-forwardable": a forward of a known sender whose `forwardable` is false
 * - "not-a-forward": from ForwardEmailAddress, but the subject does not start with a forward prefix
 * - "no-rule-matched": known sender, but no rule's subject matched
 * - "no-receipt-data": a rule matched but gave no cost and no name
 */
export type RuleOutcomeReason =
  | "unknown-sender"
  | "not-forwardable"
  | "not-a-forward"
  | "no-rule-matched"
  | "no-receipt-data";

export interface RuleOutcome {
  /** Lower-cased address the rules were chosen by (original sender for forwards). null when no sender's rules were used (unknown sender, not forwardable, not a forward). */
  senderAddress: string | null;
  forwarded: boolean;
  /** null: no rule's subject matched, or no sender's rules were used. */
  ruleId: string | null;
  /** null for a receipt; otherwise why it is not one */
  reason: RuleOutcomeReason | null;
  receipt: {
    cost: number | undefined;
    name: string | undefined;
    category: string | undefined;
    type: string | undefined;
  };
  /** `!!cost || !!name` with JS truthiness (so 0 and NaN are NOT receipts), exactly ThreadList's test today. */
  isReceipt: boolean;
  captures: { subject: RuleCaptures | null; body: RuleCaptures | null };
  /** e.g. "Rule x needs the HTML body but it was not provided" */
  warnings: string[];
}

/** The original body of a forwarded email: from the first found marker on, minus its first `skipLines` lines */
function unwrapForwardedBody(
  plainBody: string,
  markers: ForwardingConfig["bodyMarkers"],
): string {
  for (const { marker, skipLines } of markers) {
    const index = plainBody.indexOf(marker);
    if (index >= 0)
      return plainBody.substring(index).split("\n").slice(skipLines).join("\n");
  }
  return plainBody;
}

/** A non-receipt outcome for which no sender's rules were used */
function noRules(forwarded: boolean, reason: RuleOutcomeReason): RuleOutcome {
  return {
    senderAddress: null,
    forwarded,
    ruleId: null,
    reason,
    receipt: {
      cost: undefined,
      name: undefined,
      category: undefined,
      type: undefined,
    },
    isReceipt: false,
    captures: { subject: null, body: null },
    warnings: [],
  };
}

function tryExtractEmailAddress(text: string): string | undefined {
  try {
    return extractEmailAddress(text);
  } catch {
    return undefined;
  }
}

/**
 * Applies the rules to one email (receipt-rules-plan.md section 2.3). Unknown senders and messages
 * from ForwardEmailAddress without a forward subject prefix are not receipts (see RuleOutcomeReason).
 * Throws plain Error with these exact messages (pinned by tests):
 *  - "Could not get email address from <from>"
 *  - "Could not get 'from' email address from forwarded email with subject '<subject>'" (the subject
 *    starts with a forward prefix but no fromPattern matched: likely a new forward format)
 */
export function applyReceiptRules(
  rules: CompiledReceiptRules,
  email: RuleEmail,
  context: RuleContext,
): RuleOutcome {
  const compiled = internals.get(rules);
  if (!compiled)
    throw new Error(
      "applyReceiptRules needs the object from compileReceiptRules",
    );
  const { config } = rules;
  const fromAddress = extractEmailAddress(email.from).toLowerCase();
  const forwardAddress = context.forwardEmailAddress.toLowerCase();
  const forwarded = forwardAddress !== "" && fromAddress === forwardAddress;

  let sender: CompiledSender | undefined;
  let body = email.plainBody;
  let explicitPrefix: string | undefined;
  if (forwarded) {
    const prefixes =
      config.forwarding.subjectPrefixes ?? DEFAULT_FORWARD_SUBJECT_PREFIXES;
    if (!prefixes.some((prefix) => email.subject.startsWith(prefix)))
      return noRules(false, "not-a-forward");
    let original: string | undefined;
    for (const regex of compiled.fromPatterns) {
      original = regex.exec(email.plainBody)?.groups?.email;
      if (original !== undefined) break;
    }
    if (original === undefined)
      throw new Error(
        `Could not get 'from' email address from forwarded email with subject '${email.subject}'`,
      );
    const originalAddress = (
      tryExtractEmailAddress(original) ?? original
    ).toLowerCase();
    sender = compiled.senders.get(originalAddress);
    if (!sender) return noRules(true, "unknown-sender");
    if (!sender.sender.forwardable) return noRules(true, "not-forwardable");
    body = unwrapForwardedBody(email.plainBody, config.forwarding.bodyMarkers);
    explicitPrefix = context.forwardName || undefined;
  } else {
    sender = compiled.senders.get(fromAddress);
    if (!sender) return noRules(false, "unknown-sender");
  }

  const recipientPrefix = () => {
    const to = tryExtractEmailAddress(email.to);
    return to !== undefined &&
      to.toLowerCase() === forwardAddress &&
      context.forwardName
      ? context.forwardName
      : undefined;
  };
  const type =
    sender.sender.typePrefix === "fixed"
      ? sender.sender.type
      : `${explicitPrefix || recipientPrefix() || config.defaultTypePrefix} ${sender.sender.type}`;

  const senderAddress = sender.sender.address.toLowerCase();
  const warnings: string[] = [];
  for (const {
    rule,
    subject,
    forwardSubject,
    body: bodyRegex,
  } of sender.rules) {
    const subjectMatch = (forwarded ? forwardSubject : subject).exec(
      email.subject,
    );
    if (!subjectMatch) continue;

    const captures: RuleOutcome["captures"] = {
      subject: { ...subjectMatch.groups },
      body: null,
    };
    if (rule.body && bodyRegex) {
      let text: string | undefined = body;
      if (rule.body.source === "htmlBody") {
        text = email.htmlBody;
        if (text === undefined)
          warnings.push(
            `Rule ${rule.id} needs the HTML body but it was not provided`,
          );
        else if (forwarded)
          warnings.push("forwarded HTML bodies are not unwrapped");
      }
      const bodyMatch = text === undefined ? null : bodyRegex.exec(text);
      if (bodyMatch) captures.body = { ...bodyMatch.groups };
    }

    let cost: number | undefined;
    if (rule.cost) {
      const text = captures[rule.cost.source]?.[rule.cost.group];
      cost = text ? parseFloat(text) : undefined;
      if (rule.cost.negate && cost !== undefined) cost = cost * -1;
    }
    const expand = (template: string | undefined) =>
      template === undefined ? undefined : expandTemplate(template, captures);
    const name = expand(rule.name) ?? expand(rule.nameFallback);
    const isReceipt = !!cost || !!name;
    return {
      senderAddress,
      forwarded,
      ruleId: rule.id,
      reason: isReceipt ? null : "no-receipt-data",
      receipt: {
        cost,
        name,
        category: isReceipt ? rule.category : undefined,
        type,
      },
      isReceipt,
      captures,
      warnings,
    };
  }

  return {
    senderAddress,
    forwarded,
    ruleId: null,
    reason: "no-rule-matched",
    receipt: { cost: undefined, name: undefined, category: undefined, type },
    isReceipt: false,
    captures: { subject: null, body: null },
    warnings,
  };
}

// #endregion
