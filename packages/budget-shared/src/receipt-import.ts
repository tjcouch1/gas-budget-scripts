/**
 * Pure part of the receipt import: how fetched emails become transaction rows, notes and errors. The
 * GAS import (ThreadList) and the web app's email preview both call this, so what the preview shows is
 * what the import writes. Nothing here reads or writes Gmail or the sheet.
 */
import type { RuleOutcome } from "./receipt-rules.ts";

/** A JS Date or the GAS Date type (which lacks Symbol.toPrimitive) */
export interface DateLike {
  getTime(): number;
}

export const NEXT_ERROR_SEPARATOR =
  "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n";
export const NEXT_NOTE_SEPARATOR =
  "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n";
const NOTE_BODY_CHARS = 280;

export interface ImportMessageInput {
  messageId: string;
  subject: string;
  /** Address as written in the From header (original case), for the note text */
  fromAddress: string;
  date: DateLike;
  /** The date as the caller prints it (GAS: String(message.getDate())), for note/error text */
  dateText: string;
  plainBody: string | null | undefined;
  /** null when `error` is set */
  outcome: RuleOutcome | null;
  /** Text of the failure (GAS: `${e}`) when the message could not be processed */
  error: string | null;
}

export interface ImportRow {
  /** null for the empty row an errored thread without receipts gets */
  messageId: string | null;
  date: DateLike;
  cost: number | undefined;
  name: string | undefined;
  category: string | undefined;
  type: string | undefined;
  /** Empty string means no error (shown on the name cell, red) */
  errorMessage: string;
  /** Empty string means no note (shown on the name cell, pink) */
  note: string;
}

export interface ThreadImportPlan {
  rows: ImportRow[];
  notes: string[];
  errors: string[];
  /** Receipts or errors: the thread is recorded. Otherwise it is only labeled ignored. */
  hasInformation: boolean;
}

export function messageErrorText(args: {
  messageId: string;
  subject: string;
  dateText: string;
  threadId: string;
  error: string;
}): string {
  return `Error while processing message ${args.messageId} with subject ${args.subject} from date ${args.dateText} on thread ${args.threadId}. Skipping marking as processed. ${args.error}`;
}

export function threadErrorText(threadId: string, error: string): string {
  return `Error while processing thread with ID ${threadId}. Skipping marking as processed. ${error}`;
}

export function notAReceiptNote(args: {
  fromAddress: string;
  subject: string;
  dateText: string;
  threadId: string;
  plainBody: string | null | undefined;
}): string {
  return `Message is not a receipt:\nFrom: ${args.fromAddress}\nSubject: ${args.subject}\nDate: ${args.dateText}\nThread ID: ${args.threadId}\n280 Chars of Plain Body:\n${args.plainBody?.substring(0, NOTE_BODY_CHARS)}`;
}

export const joinThreadErrors = (errors: string[]) =>
  errors.join(NEXT_ERROR_SEPARATOR);
export const joinThreadNotes = (notes: string[]) =>
  notes.join(NEXT_NOTE_SEPARATOR);

/** Same decision the import makes per message: a receipt has a (truthy) cost or name */
export function isReceiptReceipt(receipt: RuleOutcome["receipt"]): boolean {
  return !!receipt.cost || !!receipt.name;
}

/** The rows, notes and errors one Gmail thread produces */
export function planThreadImport(input: {
  threadId: string;
  /** Date of the thread's last message: the date of the empty row for a thread with only errors. Called only then. */
  lastMessageDate: () => DateLike;
  messages: ImportMessageInput[];
}): ThreadImportPlan {
  const rows: ImportRow[] = [];
  const notes: string[] = [];
  const errors: string[] = [];
  for (const m of input.messages) {
    if (m.error !== null || !m.outcome) {
      errors.push(
        messageErrorText({
          messageId: m.messageId,
          subject: m.subject,
          dateText: m.dateText,
          threadId: input.threadId,
          error: m.error ?? "",
        }),
      );
      continue;
    }
    const { cost, name, category, type } = m.outcome.receipt;
    if (!cost && !name)
      notes.push(
        notAReceiptNote({
          fromAddress: m.fromAddress,
          subject: m.subject,
          dateText: m.dateText,
          threadId: input.threadId,
          plainBody: m.plainBody,
        }),
      );
    else
      rows.push({
        messageId: m.messageId,
        date: m.date,
        cost,
        name,
        category,
        type,
        errorMessage: "",
        note: "",
      });
  }

  const hasInformation = rows.length > 0 || errors.length > 0;
  if (rows.length === 0 && errors.length > 0)
    rows.push({
      messageId: null,
      date: input.lastMessageDate(),
      cost: undefined,
      name: undefined,
      category: undefined,
      type: undefined,
      errorMessage: "",
      note: "",
    });
  if (hasInformation) {
    if (errors.length > 0) rows[0].errorMessage = joinThreadErrors(errors);
    if (notes.length > 0) rows[0].note = joinThreadNotes(notes);
  }
  return { rows, notes, errors, hasInformation };
}

/** Stable sort by date ascending (the import records receipts oldest first) */
export function sortRowsByDate<T extends { date: DateLike }>(rows: T[]): T[] {
  return [...rows].sort((a, b) => a.date.getTime() - b.date.getTime());
}

export interface TransactionSheetRange {
  startDate: DateLike;
  endDate: DateLike;
}

/** The sheet whose [startDate, endDate) holds `date`, as the import picks it */
export function findTransactionSheet<T extends TransactionSheetRange>(
  date: DateLike,
  sheets: T[],
): T | undefined {
  return sheets.find(
    (s) =>
      date.getTime() >= s.startDate.getTime() &&
      date.getTime() < s.endDate.getTime(),
  );
}

export interface SheetAssignment<R, S> {
  /** Rows the import records, each with its sheet, in input order */
  recorded: { row: R; sheet: S }[];
  /**
   * Rows the import leaves out, in input order: every row of a thread that has a row no sheet fits.
   * `sheet` is undefined for the rows that fit no sheet themselves
   */
  skipped: { row: R; sheet: S | undefined }[];
}

/**
 * Picks each row's transaction sheet, as the import does. A thread with a row that fits no sheet is
 * left out whole (none of its rows are recorded and it is not marked), so a thread is never half
 * recorded and a later import picks it up once the pay-period sheet exists.
 *
 * @param threadOf the thread a row came from, compared with `===` (GAS: the GmailThread object;
 *   preview: the thread ID)
 */
export function assignRowsToSheets<
  R extends { date: DateLike },
  S extends TransactionSheetRange,
  K,
>(rows: R[], sheets: S[], threadOf: (row: R) => K): SheetAssignment<R, S> {
  const withSheet = rows.map((row) => ({
    row,
    sheet: findTransactionSheet(row.date, sheets),
  }));
  const threadsWithoutSheet = new Set<K>();
  for (const { row, sheet } of withSheet)
    if (!sheet) threadsWithoutSheet.add(threadOf(row));
  const recorded: { row: R; sheet: S }[] = [];
  const skipped: { row: R; sheet: S | undefined }[] = [];
  for (const entry of withSheet) {
    if (threadsWithoutSheet.has(threadOf(entry.row))) skipped.push(entry);
    else recorded.push({ row: entry.row, sheet: entry.sheet! });
  }
  return { recorded, skipped };
}
