import {
  buildReceiptSearchQuery,
  type ReceiptQueryOptions,
  compileReceiptRules,
  parseReceiptRulesText,
  type CompiledReceiptRules,
  type RuleContext,
  type RuleEmail,
} from "@budget/shared";
import { getVariables } from "./variables.util";

/** Same text as `compileReceiptRules` throws, for issues found before compiling (JSON, cell type) */
function invalidConfigError(path: string, message: string) {
  return new Error(
    `ReceiptRules variable is invalid: ${path}: ${message}. Fix it in the web app (Settings > Customize variables) or clear the cell to use the built-in defaults.`,
  );
}

/** Message prefix when the ForwardEmailAddress variable cannot be used in the Gmail query */
export const FORWARD_ADDRESS_INVALID_PREFIX =
  "ForwardEmailAddress variable is invalid";

let cache:
  | {
      raw: unknown;
      forwardEmailAddress: string;
      compiled: CompiledReceiptRules;
    }
  | undefined;

/** Raw `ReceiptRules` cell value as read from the Variables sheet (undefined when the key is missing) */
function readRawCell(): unknown {
  return getVariables().ReceiptRules;
}

/** Raw cell text ("" when the key is missing or empty) */
export function getReceiptRulesRaw(): string {
  const raw = readRawCell();
  return raw === undefined || raw === null ? "" : String(raw);
}

/**
 * Compiles the ReceiptRules variable (empty or missing = built-in defaults). Memoized per execution
 * by raw cell value and ForwardEmailAddress; `bustVariablesCache` makes getVariables re-read, which
 * recompiles when either changed.
 * Throws "ReceiptRules variable is invalid: ..." for an invalid value
 */
export function getCompiledReceiptRules(): CompiledReceiptRules {
  const raw = readRawCell();
  const forwardEmailAddress = getVariables().ForwardEmailAddress;
  if (
    cache &&
    cache.raw === raw &&
    cache.forwardEmailAddress === forwardEmailAddress
  )
    return cache.compiled;

  if (raw !== undefined && raw !== null && typeof raw !== "string")
    throw invalidConfigError("ReceiptRules", "must be text (JSON)");

  const options = { forwardEmailAddress };
  const parsed = parseReceiptRulesText(raw as string | undefined, options);
  if (!parsed.ok) {
    const [{ path, message }] = parsed.issues;
    throw invalidConfigError(path, message);
  }
  const compiled = compileReceiptRules(parsed.config, options);
  cache = { raw, forwardEmailAddress, compiled };
  return compiled;
}

/** Lower-cased sender addresses of the compiled rules */
export function getReceiptSenderAddresses(): string[] {
  return getCompiledReceiptRules().senderAddresses;
}

/** Gmail search query for unprocessed receipt threads */
export function getReceiptSearchQuery(): string {
  return buildReceiptQuery(getReceiptSenderAddresses());
}

/**
 * Query for the given (already validated) sender addresses plus the ForwardEmailAddress variable.
 * A malformed ForwardEmailAddress throws "ForwardEmailAddress variable is invalid: ..."
 */
export function buildReceiptQuery(
  senders: string[],
  options: ReceiptQueryOptions = {},
): string {
  try {
    return buildReceiptSearchQuery(
      senders,
      getVariables().ForwardEmailAddress,
      options,
    );
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    throw new Error(
      `${FORWARD_ADDRESS_INVALID_PREFIX}: ${message}. Fix it on the Variables sheet (a plain address like name@example.com).`,
      { cause: e },
    );
  }
}

/** What the rule engine needs from a Gmail message. Calls `getBody()` only when `includeHtml` */
export function toRuleEmail(
  message: GoogleAppsScript.Gmail.GmailMessage,
  includeHtml: boolean,
): RuleEmail {
  const email: RuleEmail = {
    from: message.getFrom(),
    to: message.getTo(),
    subject: message.getSubject(),
    plainBody: message.getPlainBody(),
  };
  if (includeHtml) email.htmlBody = message.getBody();
  return email;
}

/** Forward address and name from the Variables sheet */
export function getRuleContext(): RuleContext {
  const { ForwardEmailAddress, ForwardName } = getVariables();
  return { forwardEmailAddress: ForwardEmailAddress, forwardName: ForwardName };
}
