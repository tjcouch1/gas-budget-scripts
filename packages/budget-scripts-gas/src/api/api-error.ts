import type { ApiErrorCode, TransactionRowDto } from "@budget/shared";

/** An error the API reports to the client with a specific error code (web-plan.md section 4.4) */
export class ApiError extends Error {
  readonly code: ApiErrorCode;
  /**
   * Present for `conflict` errors: the row as it currently is on the sheet, plus for a split group
   * edit the group's current rows
   */
  readonly conflict:
    { current: TransactionRowDto; rows?: TransactionRowDto[] } | undefined;
  /** Validation problems behind a `bad_request` (e.g. invalid ReceiptRules) */
  readonly issues: { path: string; message: string }[] | undefined;

  constructor(
    code: ApiErrorCode,
    message: string,
    conflict?: { current: TransactionRowDto; rows?: TransactionRowDto[] },
    issues?: { path: string; message: string }[],
  ) {
    super(message);
    this.name = "ApiError";
    this.code = code;
    this.conflict = conflict;
    this.issues = issues;
  }
}

/** `busy` message of calls that share the lock with receipt imports (syncReceipts, saveReceiptRules) */
export const RECEIPT_IMPORT_BUSY_MESSAGE =
  "Another change or receipt import is in progress. Please try again.";

/** Message of any thrown value without a stack trace */
export function getErrorMessage(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
