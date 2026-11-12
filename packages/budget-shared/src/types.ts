/**
 * Shared DTOs and API request/response types for the budget GAS Web App and the web client.
 * Pure types only: no DOM or GAS globals. See web-plan.md section 4.4.
 */

export const API_VERSION = 1;

export type TaxDeductibleValue = "Yes" | "";

/** Fields of an existing transaction row that the API may change. */
export interface TransactionChanges {
  name?: string;
  /** Validated by the shared math parser. */
  costExpression?: string;
  /** When true, the cost is wrapped with the sheet TaxMultiplier like a split child. Default false. */
  costPreTax?: boolean;
  envelope?: string;
  paidWith?: string;
  taxDeductible?: TaxDeductibleValue;
}

export interface SplitItemInput {
  itemNote: string;
  expression: string;
  preTax: boolean;
  envelope: string;
  paidWith?: string;
}

interface BaseRequest {
  apiVersion: typeof API_VERSION;
  apiKey: string;
}

export interface GetStateRequest extends BaseRequest {
  action: "getState";
  sheetName?: string;
}

/** Reads every pay-period tab's transactions (the "All pay periods" view). Read only. */
export interface GetAllTransactionsRequest extends BaseRequest {
  action: "getAllTransactions";
}

export interface UpdateTransactionRequest extends BaseRequest {
  action: "updateTransaction";
  sheetName: string;
  rowIndex: number;
  baseFingerprint: string;
  changes: TransactionChanges;
  force?: boolean;
}

export interface SplitTransactionRequest extends BaseRequest {
  action: "splitTransaction";
  sheetName: string;
  rowIndex: number;
  baseFingerprint: string;
  splits: SplitItemInput[];
  force?: boolean;
}

export interface SyncReceiptsRequest extends BaseRequest {
  action: "syncReceipts";
}

export type ApiRequest =
  | GetStateRequest
  | GetAllTransactionsRequest
  | UpdateTransactionRequest
  | SplitTransactionRequest
  | SyncReceiptsRequest;

export type ApiErrorCode =
  | "unauthorized"
  | "bad_request"
  | "not_found"
  | "conflict"
  | "busy"
  | "server_error";

export interface ApiSuccess<T> {
  ok: true;
  data: T;
  warnings?: string[];
}

export interface ApiFailure {
  ok: false;
  error: { code: ApiErrorCode; message: string };
  conflict?: { current: TransactionRowDto };
}

export type ApiResponse<T> = ApiSuccess<T> | ApiFailure;

export interface TransactionRowDto {
  /** 0-based index from TransactionsStart. */
  rowIndex: number;
  /** yyyy-MM-dd in the spreadsheet time zone, or "". */
  date: string;
  name: string;
  cost: number | null;
  costFormula: string | null;
  costDisplay: string;
  envelope: string;
  paidWith: string;
  taxDeductible: string;
  note: string;
  isError: boolean;
  isNote: boolean;
  splitGroupId: string | null;
  fingerprint: string;
}

export interface SheetInfoDto {
  name: string;
  /** yyyy-MM-dd */
  startDate: string;
  /** yyyy-MM-dd */
  endDate: string;
  isGap: boolean;
}

/** One pay-period tab in a `getAllTransactions` response. Edits use this `name` as `sheetName`. */
export interface SheetTransactionsDto extends SheetInfoDto {
  /** Null entries are interior empty rows (same as `PayPeriodStateDto.transactions`). */
  transactions: (TransactionRowDto | null)[];
  /** The tab's own option lists; null when it has no sheet variables for them. */
  envelopeOptions: string[] | null;
  paidWithOptions: string[] | null;
}

/** `getAllTransactions` response data: every transaction tab, newest first. */
export interface AllTransactionsDto {
  taxMultiplier: number;
  sheets: SheetTransactionsDto[];
}

export interface EnvelopeRowDto {
  name: string;
  /** Display strings for cols 1-10 of the envelope table (web-plan.md section 3.5). */
  display: string[];
  /** Numeric values where the cell is numeric, else null. Same order as `display`. */
  values: (number | null)[];
}

/** `doGet` response data (unauthenticated connectivity check). */
export interface ServiceInfoDto {
  service: "budget-api";
  apiVersion: typeof API_VERSION;
}

export interface PayPeriodStateDto {
  sheets: SheetInfoDto[];
  sheetName: string;
  taxMultiplier: number;
  forwardName: string;
  /** Null entries are interior empty rows. */
  transactions: (TransactionRowDto | null)[];
  envelopes: EnvelopeRowDto[] | null;
  paidWithOptions: string[] | null;
  envelopeOptions: string[] | null;
}

export interface SyncReceiptsResultDto {
  /** Total number of receipt rows added across all sheets. */
  receiptInfosAdded: number;
  /** Receipt rows added per transaction sheet name (only sheets that got rows). */
  receiptInfosAddedBySheet?: Record<string, number>;
  numErrors: number;
  errors: string[];
  state: PayPeriodStateDto;
}

/** The fields the fingerprint covers (web-plan.md section 4.4). */
export interface FingerprintRow {
  /** A Date (local time) or an already-formatted yyyy-MM-dd string; blank/undefined means "". */
  date: Date | string | null | undefined;
  name: string | null | undefined;
  cost: number | string | null | undefined;
  /** When present (the cell holds a formula) it is used instead of `cost`. */
  costFormula?: string | null;
  envelope: string | null | undefined;
  paidWith: string | null | undefined;
  taxDeductible: string | null | undefined;
}
