/** Gmail search query builder for receipt fetching. Pure; no GAS or DOM globals. */

export const STRICT_EMAIL_REGEX =
  /^[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/;

export function isStrictEmailAddress(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= 254 &&
    STRICT_EMAIL_REGEX.test(value)
  );
}

const UNSAFE_SUBJECT_CHARS = /["(){}:\\\r\n]/;

/** Allowed subject filter text: 1-100 chars, none of " ( ) { } : \ newline; trimmed non-empty */
export function isSafeSubjectFilter(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= 100 &&
    value.trim().length > 0 &&
    !UNSAFE_SUBJECT_CHARS.test(value)
  );
}

/** Most sender addresses in one query: 25 configurable senders plus a few extras */
export const MAX_QUERY_SENDERS = 30;

/** Fixed, never configurable (labels from AGENTS.md section 3.4) */
export const RECEIPT_LABEL_EXCLUSIONS =
  "NOT label:receipts NOT label:receipts-cru-reimburse NOT label:receipts-tax-deductible NOT label:receipts-scripted";

/** Used when `forwarding.subjectPrefixes` is omitted (today's behavior) */
export const DEFAULT_FORWARD_SUBJECT_PREFIXES: readonly string[] = ["Fwd: "];

/**
 * The Gmail search token for a forward subject prefix: the prefix without its trailing ":" and
 * whitespace ("Fwd: " => "Fwd", "FW:" => "FW"). undefined unless what is left is 1+ letters/digits
 * (any script), which keeps the token safe to put in a query unquoted.
 */
export function forwardPrefixToken(prefix: string): string | undefined {
  const token = prefix.replace(/[\s:]+$/u, "");
  return /^[\p{L}\p{N}]+$/u.test(token) ? token : undefined;
}

export interface ReceiptQueryOptions {
  /** default false: adds "in:inbox" and the label exclusions. true: search all mail, no exclusions */
  includeProcessed?: boolean;
  /** each validated with isSafeSubjectFilter; emitted as subject:("a" OR "b") */
  subjects?: string[];
  /** integer 1-3650; emitted as newer_than:<n>d */
  newerThanDays?: number;
  /**
   * Forward subject prefixes (default DEFAULT_FORWARD_SUBJECT_PREFIXES). Messages from the forward
   * address only match when the subject contains one of their tokens (see forwardPrefixToken)
   */
  forwardSubjectPrefixes?: readonly string[];
}

/**
 * `(from:(a OR b) OR (from:<forwardEmailAddress> subject:Fwd)) in:inbox <exclusions> ` (the trailing
 * space is pinned by tests). Several forward prefixes give `subject:(Fwd OR FW)`. Without a forward
 * address it is `from:(a OR b) ...`; without senders it is `(from:<fwd> subject:Fwd) ...`.
 *
 * Gmail search syntax: "OR" (upper case) joins terms, and parentheses group terms, including
 * nested groups and `op:(a OR b)` values, so `(X OR (Y Z))` means X or (Y and Z). Gmail cannot scope
 * `subject:` to one address inside a single `from:(...)`, hence the separate group. Its subject search
 * is word-based: `subject:Fwd` matches "Fwd" anywhere in the subject, not only as a prefix; the rule
 * engine checks the real prefix.
 *
 * Throws Error("Invalid sender address: <x>") for any address failing the strict regex and
 * Error("Invalid forward subject prefix: <p>") for a prefix without a safe token. No de-duplication.
 */
export function buildReceiptSearchQuery(
  senderAddresses: string[],
  forwardEmailAddress: string,
  options: ReceiptQueryOptions = {},
): string {
  for (const address of forwardEmailAddress
    ? [...senderAddresses, forwardEmailAddress]
    : senderAddresses)
    if (!isStrictEmailAddress(address))
      throw new Error(`Invalid sender address: ${address}`);

  const fromSenders = `from:(${senderAddresses.join(" OR ")})`;
  let from = fromSenders;
  if (forwardEmailAddress) {
    const tokens = (
      options.forwardSubjectPrefixes ?? DEFAULT_FORWARD_SUBJECT_PREFIXES
    ).map((prefix) => {
      const token = forwardPrefixToken(prefix);
      if (token === undefined)
        throw new Error(`Invalid forward subject prefix: ${prefix}`);
      return token;
    });
    const subject =
      tokens.length === 1 ? tokens[0] : `(${tokens.join(" OR ")})`;
    const fromForward = `(from:${forwardEmailAddress} subject:${subject})`;
    from =
      senderAddresses.length > 0
        ? `(${fromSenders} OR ${fromForward})`
        : fromForward;
  }
  const parts = [from];

  if (options.subjects && options.subjects.length > 0) {
    for (const subject of options.subjects)
      if (!isSafeSubjectFilter(subject))
        throw new Error(`Invalid subject filter: ${subject}`);
    parts.push(
      `subject:(${options.subjects.map((s) => `"${s}"`).join(" OR ")})`,
    );
  }

  if (options.newerThanDays !== undefined) {
    const days = options.newerThanDays;
    if (!Number.isInteger(days) || days < 1 || days > 3650)
      throw new Error(`Invalid newerThanDays: ${days}`);
    parts.push(`newer_than:${days}d`);
  }

  if (!options.includeProcessed)
    parts.push("in:inbox", RECEIPT_LABEL_EXCLUSIONS);

  return parts.join(" ") + " ";
}
