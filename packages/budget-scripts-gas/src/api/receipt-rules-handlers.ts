/**
 * Receipt rules API (receipt-rules-plan.md section 3.5): read-only Gmail fetch for testing rules in
 * the browser, and reading/saving the ReceiptRules global variable. The server never runs candidate
 * (unsaved) rules against emails.
 */
import {
  isSafeSubjectFilter,
  isStrictEmailAddress,
  MAX_QUERY_SENDERS,
  parseReceiptRulesText,
  validateReceiptRulesConfig,
  type FetchReceiptEmailsResultDto,
  type ReceiptEmailDto,
  type ReceiptRulesStateDto,
} from "@budget/shared";
import { ensureGlobalVariableRow } from "../util/global-variables-setup.util";
import { tryWithScriptLock } from "../util/lock.util";
import {
  buildReceiptQuery,
  FORWARD_ADDRESS_INVALID_PREFIX,
  getReceiptRulesRaw,
  getReceiptSearchQuery,
  getReceiptSenderAddresses,
} from "../util/receipt-rules.util";
import { bustVariablesCache, getVariables } from "../util/variables.util";
import {
  ApiError,
  getErrorMessage,
  RECEIPT_IMPORT_BUSY_MESSAGE,
} from "./api-error";
import type { HandlerResult } from "./handlers";

type Json = Record<string, unknown>;

/** Default threads per fetch (owner decision 11.3) */
export const FETCH_DEFAULT_THREADS = 10;
/** Most threads per fetch; larger requests are clamped */
export const FETCH_MAX_THREADS = 50;
export const FETCH_MAX_MESSAGES = 100;
export const FETCH_MAX_TOTAL_CHARS = 3_000_000;
export const FETCH_PLAIN_BODY_MAX = 20_000;
export const FETCH_HTML_BODY_MAX = 60_000;
const FETCH_MAX_SUBJECTS = 5;

/** Message prefix of every invalid-ReceiptRules error thrown by receipt-rules.util */
export const INVALID_RECEIPT_RULES_PREFIX = "ReceiptRules variable is invalid";

function badRequest(message: string): never {
  throw new ApiError("bad_request", message);
}

/** Runs `fn`, turning an invalid ReceiptRules variable into `invalid_config` */
export function mapInvalidConfig<T>(fn: () => T): T {
  try {
    return fn();
  } catch (e) {
    const message = getErrorMessage(e);
    if (
      message.startsWith(INVALID_RECEIPT_RULES_PREFIX) ||
      message.startsWith(FORWARD_ADDRESS_INVALID_PREFIX)
    )
      throw new ApiError("invalid_config", getErrorMessage(e));
    throw e;
  }
}

function optionalBoolean(value: unknown, field: string): boolean {
  if (value === undefined) return false;
  if (typeof value !== "boolean") badRequest(`${field} must be a boolean`);
  return value;
}

/** Integer in [min, max]; with `clamp`, values above `max` become `max` instead of an error */
function optionalInteger(
  value: unknown,
  field: string,
  min: number,
  max: number,
  clamp = false,
): number | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < min ||
    (!clamp && value > max)
  )
    badRequest(`${field} must be an integer from ${min} to ${max}`);
  return Math.min(value, max);
}

function validateSenders(value: unknown): string[] | undefined {
  if (value === undefined) return undefined;
  if (
    !Array.isArray(value) ||
    value.length < 1 ||
    value.length > MAX_QUERY_SENDERS
  )
    badRequest(
      `senders must be an array of 1-${MAX_QUERY_SENDERS} email addresses`,
    );
  value.forEach((sender, i) => {
    if (!isStrictEmailAddress(sender))
      badRequest(`senders[${i}] is not a valid email address`);
  });
  return value as string[];
}

function validateSubjects(value: unknown): string[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value) || value.length > FETCH_MAX_SUBJECTS)
    badRequest(
      `subjects must be an array of 0-${FETCH_MAX_SUBJECTS} subject filters`,
    );
  value.forEach((subject, i) => {
    if (!isSafeSubjectFilter(subject))
      badRequest(`subjects[${i}] contains characters that are not allowed`);
  });
  return value as string[];
}

function cut(text: string, max: number) {
  return text.length > max
    ? { text: text.slice(0, max), truncated: true }
    : { text, truncated: false };
}

/**
 * Read-only: searches Gmail and returns the messages of the matching threads. No lock, never labels,
 * archives or marks anything
 */
export function handleFetchReceiptEmails(
  req: Json,
): HandlerResult<FetchReceiptEmailsResultDto> {
  const explicitSenders = validateSenders(req.senders);
  const subjects = validateSubjects(req.subjects);
  const includeProcessed = optionalBoolean(
    req.includeProcessed,
    "includeProcessed",
  );
  const includeHtml = optionalBoolean(req.includeHtml, "includeHtml");
  const maxThreads =
    optionalInteger(req.maxThreads, "maxThreads", 1, FETCH_MAX_THREADS, true) ??
    FETCH_DEFAULT_THREADS;
  const newerThanDays = optionalInteger(
    req.newerThanDays,
    "newerThanDays",
    1,
    3650,
  );

  const query = mapInvalidConfig(() => {
    const senders = explicitSenders ?? getReceiptSenderAddresses();
    // Request fields are validated above, so only the ForwardEmailAddress variable can fail here
    return buildReceiptQuery(senders, {
      includeProcessed,
      subjects,
      newerThanDays,
    });
  });

  const threads = GmailApp.search(query, 0, maxThreads);
  const emails: ReceiptEmailDto[] = [];
  let totalChars = 0;
  let truncated = false;
  outer: for (const thread of threads) {
    const labels = thread.getLabels().map((label) => label.getName());
    for (const message of thread.getMessages()) {
      if (emails.length >= FETCH_MAX_MESSAGES) {
        truncated = true;
        break outer;
      }
      const plain = cut(message.getPlainBody(), FETCH_PLAIN_BODY_MAX);
      const html = includeHtml
        ? cut(message.getBody(), FETCH_HTML_BODY_MAX)
        : undefined;
      const size = plain.text.length + (html?.text.length ?? 0);
      if (totalChars + size > FETCH_MAX_TOTAL_CHARS) {
        truncated = true;
        break outer;
      }
      totalChars += size;
      const email: ReceiptEmailDto = {
        threadId: thread.getId(),
        messageId: message.getId(),
        from: message.getFrom(),
        to: message.getTo(),
        subject: message.getSubject(),
        date: message.getDate().toISOString(),
        plainBody: plain.text,
        plainBodyTruncated: plain.truncated,
        labels,
      };
      if (html) {
        email.htmlBody = html.text;
        email.htmlBodyTruncated = html.truncated;
      }
      emails.push(email);
    }
  }

  return { data: { query, threadCount: threads.length, emails, truncated } };
}

/** The ReceiptRules variable as stored, parsed and validated. No lock, no writes */
export function handleGetReceiptRules(): HandlerResult<ReceiptRulesStateDto> {
  const variables = getVariables();
  const raw = getReceiptRulesRaw();
  const forwardEmailAddress = variables.ForwardEmailAddress;
  const parsed = parseReceiptRulesText(raw, { forwardEmailAddress });
  let query: string | null = null;
  if (parsed.ok)
    try {
      query = getReceiptSearchQuery();
    } catch {
      // e.g. an invalid ForwardEmailAddress: no query to show
    }
  return {
    data: {
      raw,
      hasVariable: Object.prototype.hasOwnProperty.call(
        variables,
        "ReceiptRules",
      ),
      config: parsed.ok ? parsed.config : null,
      isDefault: parsed.ok && parsed.isDefault === true,
      issues: parsed.ok ? [] : parsed.issues,
      forwardEmailAddress,
      forwardName: variables.ForwardName,
      query,
    },
  };
}

/**
 * Validates and writes the ReceiptRules variable (pretty JSON; null clears the cell) under the
 * script lock, adding the key to the Variables sheet when missing. Returns the state after the write
 */
export function handleSaveReceiptRules(
  req: Json,
): HandlerResult<ReceiptRulesStateDto> {
  const { rules, expectedRaw } = req;
  if (rules !== null && (typeof rules !== "object" || Array.isArray(rules)))
    badRequest("rules must be an object or null");
  if (expectedRaw !== undefined && typeof expectedRaw !== "string")
    badRequest("expectedRaw must be a string");

  return tryWithScriptLock(
    () => {
      // Re-read inside the lock
      bustVariablesCache();
      if (expectedRaw !== undefined && expectedRaw !== getReceiptRulesRaw())
        throw new ApiError(
          "conflict",
          "The rules on the sheet changed since you loaded them. Reload and try again.",
        );

      let text = "";
      if (rules !== null) {
        const result = validateReceiptRulesConfig(rules, {
          forwardEmailAddress: getVariables().ForwardEmailAddress,
        });
        if (!result.ok) {
          const [{ path, message }] = result.issues;
          throw new ApiError(
            "bad_request",
            `ReceiptRules is invalid: ${path}: ${message}`,
            undefined,
            result.issues,
          );
        }
        // Starts with "{", so Sheets never treats it as a formula. Size is checked by validation
        text = JSON.stringify(result.config, null, 2);
      }

      ensureGlobalVariableRow("ReceiptRules").valueRange.setValues([[text]]);
      bustVariablesCache();
      return handleGetReceiptRules();
    },
    () => {
      throw new ApiError("busy", RECEIPT_IMPORT_BUSY_MESSAGE);
    },
  );
}
