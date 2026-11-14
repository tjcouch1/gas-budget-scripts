/**
 * Shared DTOs and API request/response types for the budget GAS Web App and the web client.
 * Pure types only: no DOM or GAS globals. See web-plan.md section 4.4.
 */

import type { ReceiptRulesConfig } from "./receipt-rules.ts";

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
  /** The Name cell's note (plain text, not a formula). "" removes it. */
  note?: string;
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
  /** Required when `changes.note` is set: the note at the client's last pull (notes are not fingerprinted). */
  baseNote?: string;
  force?: boolean;
}

export interface SplitTransactionRequest extends BaseRequest {
  action: "splitTransaction";
  sheetName: string;
  rowIndex: number;
  baseFingerprint: string;
  splits: SplitItemInput[];
  /** New text after " | " in the original row's name (written with the split). Omit to keep the native `<name> | `. */
  parentItemNote?: string;
  /** New envelope for the original row (written with the split). Omit to leave it unchanged. */
  parentEnvelope?: string;
  force?: boolean;
}

/**
 * Deletes one row, or a whole split group together (rows below move up). Response data: the
 * refreshed `PayPeriodStateDto`, like `splitTransaction`.
 */
export interface DeleteTransactionRequest extends BaseRequest {
  action: "deleteTransaction";
  sheetName: string;
  /** First row of the split group, or the single row. */
  rowIndex: number;
  /** Every row to delete, contiguous and in order starting at `rowIndex`, with its fingerprint at the last pull. */
  rows: { rowIndex: number; baseFingerprint: string }[];
  force?: boolean;
}

/** The fields of a new transaction row (`addTransaction`). */
export interface NewTransactionFields {
  /** yyyy-MM-dd, read in the spreadsheet time zone. */
  date: string;
  /** Required; must not start with =, +, -, or @. */
  name: string;
  /** Validated by the shared math parser (no tax multiplier). */
  costExpression: string;
  /** "" or one of the sheet's envelope options. Default "". */
  envelope?: string;
  /** "" or one of the sheet's Paid with options. Default "". */
  paidWith?: string;
  /** Default "". */
  taxDeductible?: TaxDeductibleValue;
  /** The Name cell's note. Default "" (none). */
  note?: string;
}

/**
 * Appends one row after the last non-empty transaction row (never into an interior gap).
 * Idempotent per `requestId`: a repeat (e.g. an offline retry) returns the state without
 * inserting again. Response data: `AddTransactionResultDto`.
 */
export interface AddTransactionRequest extends BaseRequest {
  action: "addTransaction";
  sheetName: string;
  /** Client-generated id, 8-64 chars of A-Z a-z 0-9 _ -. Remembered for 6 hours. */
  requestId: string;
  fields: NewTransactionFields;
}

/** `addTransaction` response data: the refreshed state plus the row the new transaction went into. */
export interface AddTransactionResultDto extends PayPeriodStateDto {
  addedRowIndex: number;
}

export interface SyncReceiptsRequest extends BaseRequest {
  action: "syncReceipts";
  /** Default true. False imports only: no Gmail labels or archiving ("testing mode"). */
  mark?: boolean;
}

export type ApiRequest =
  | GetStateRequest
  | GetAllTransactionsRequest
  | UpdateTransactionRequest
  | SplitTransactionRequest
  | DeleteTransactionRequest
  | AddTransactionRequest
  | SyncReceiptsRequest
  | FetchReceiptEmailsRequest
  | GetReceiptRulesRequest
  | SaveReceiptRulesRequest;

export type ApiErrorCode =
  | "unauthorized"
  | "bad_request"
  | "not_found"
  | "conflict"
  | "busy"
  | "invalid_config"
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
  issues?: { path: string; message: string }[];
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
  /** 1-50, default 10 threads (owner decision 11.3). */
  maxThreads?: number;
  /** 1-3650 days. */
  newerThanDays?: number;
  /**
   * Forward subject prefixes for the forward address's query clause (the candidate rules'
   * forwarding.subjectPrefixes), 1-5, see forwardPrefixToken. Omit to use the saved rules' (or the
   * built-in default when `senders` is given).
   */
  forwardSubjectPrefixes?: string[];
}

export interface ReceiptEmailDto {
  threadId: string;
  messageId: string;
  from: string;
  to: string;
  subject: string;
  /** ISO 8601 */
  date: string;
  plainBody: string;
  plainBodyTruncated: boolean;
  htmlBody?: string;
  htmlBodyTruncated?: boolean;
  /** Thread labels, e.g. ["Receipts", "Receipts/Scripted"] */
  labels: string[];
}

export interface FetchReceiptEmailsResultDto {
  query: string;
  threadCount: number;
  emails: ReceiptEmailDto[];
  /** true when the message or size caps cut the list */
  truncated: boolean;
}

export interface GetReceiptRulesRequest extends BaseRequest {
  action: "getReceiptRules";
}

export interface ReceiptRulesStateDto {
  /** Cell text, "" when missing/empty */
  raw: string;
  /** Whether the ReceiptRules key exists on the Variables sheet */
  hasVariable: boolean;
  /** Parsed config when valid (defaults when raw is empty) */
  config: ReceiptRulesConfig | null;
  isDefault: boolean;
  /** Non-empty when invalid */
  issues: { path: string; message: string }[];
  forwardEmailAddress: string;
  forwardName: string;
  /** Built from `config` (or null when invalid) */
  query: string | null;
}

export interface SaveReceiptRulesRequest extends BaseRequest {
  action: "saveReceiptRules";
  /** The config object, or null to clear the cell (back to built-in defaults). */
  rules: ReceiptRulesConfig | null;
  /** Optional optimistic check: the `raw` from the last getReceiptRules. Mismatch => error code "conflict". */
  expectedRaw?: string;
}
// saveReceiptRules response data: ReceiptRulesStateDto (after the write)
