import type { ApiErrorCode, TransactionRowDto } from "@budget/shared";

/** An error the API reports to the client with a specific error code (web-plan.md section 4.4) */
export class ApiError extends Error {
  readonly code: ApiErrorCode;
  /** Present for `conflict` errors: the row as it currently is on the sheet */
  readonly conflict: { current: TransactionRowDto } | undefined;

  constructor(
    code: ApiErrorCode,
    message: string,
    conflict?: { current: TransactionRowDto },
  ) {
    super(message);
    this.name = "ApiError";
    this.code = code;
    this.conflict = conflict;
  }
}

/** Message of any thrown value without a stack trace */
export function getErrorMessage(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
