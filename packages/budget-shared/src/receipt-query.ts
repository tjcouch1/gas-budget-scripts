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

export interface ReceiptQueryOptions {
  /** default false: adds "in:inbox" and the label exclusions. true: search all mail, no exclusions */
  includeProcessed?: boolean;
  /** each validated with isSafeSubjectFilter; emitted as subject:("a" OR "b") */
  subjects?: string[];
  /** integer 1-3650; emitted as newer_than:<n>d */
  newerThanDays?: number;
}

/**
 * `from:(a OR b OR <forwardEmailAddress>) in:inbox <exclusions> ` (trailing space is pinned by tests).
 * Throws Error("Invalid sender address: <x>") for any address failing the strict regex.
 * forwardEmailAddress is appended last when non-empty. No de-duplication.
 */
export function buildReceiptSearchQuery(
  senderAddresses: string[],
  forwardEmailAddress: string,
  options: ReceiptQueryOptions = {},
): string {
  const addresses = forwardEmailAddress
    ? [...senderAddresses, forwardEmailAddress]
    : senderAddresses;
  for (const address of addresses)
    if (!isStrictEmailAddress(address))
      throw new Error(`Invalid sender address: ${address}`);

  const parts = [`from:(${addresses.join(" OR ")})`];

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
