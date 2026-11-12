/**
 * Reads sheet data into the shared DTOs (web-plan.md sections 3.4-3.6 and 4.4). Every read is a
 * batch read of a contiguous range (AGENTS.md section 4.3).
 */
import {
  computeRowFingerprint,
  type AllTransactionsDto,
  type EnvelopeRowDto,
  type PayPeriodStateDto,
  type SheetInfoDto,
  type TransactionRowDto,
} from "@budget/shared";
import {
  Budgeting,
  doTransactionRowInfosSeemEqual,
} from "../util/budgeting.util";
import { Variables } from "../util/variables.util";
import { ApiError } from "./api-error";

/** Number of transaction table columns the API reads (Date .. Tax Deductible) */
export const TRANSACTION_COLUMNS = 7;
/** Transaction column indexes from TransactionsStart */
export const COL = {
  date: 0,
  name: 1,
  cost: 2,
  currentBalance: 3,
  envelope: 4,
  paidWith: 5,
  taxDeductible: 6,
} as const;

/** Number of envelope table columns (web-plan.md section 3.5) */
const ENVELOPE_COLUMNS = 11;
/** Upper bound for the optional *Max sheet variables, to guard against typos */
const MAX_OPTIONAL_TABLE_ROWS = 1000;

const ERROR_BACKGROUND = "#ff0000";
const NOTE_BACKGROUND = "#e8a9ca";

const ENVELOPE_SKIP_LABELS = new Set(["envelope", "total", "totals"]);
const PAID_WITH_SKIP_LABELS = new Set([
  "paid with",
  "total",
  "totals",
  "tax deductible",
]);

type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type TransactionSheetInfo = ReturnType<
  typeof Budgeting.getTransactionSheetInfos
>[number];

/** Formats a value as yyyy-MM-dd in the spreadsheet time zone, or "" if it is not a valid Date */
function formatCellDate(value: unknown, timeZone: string): string {
  if (value instanceof Date)
    return Number.isNaN(value.getTime())
      ? ""
      : Utilities.formatDate(value, timeZone, "yyyy-MM-dd");
  return String(value ?? "").replace(/\s+$/, "");
}

/**
 * Formats a pay-period boundary from `getTransactionSheetInfos` as yyyy-MM-dd. Those dates are
 * parsed at GMT-06:00, so they are formatted at that fixed offset (formatting the end date in
 * America/Chicago during daylight saving time would give the next day).
 */
function formatSheetBoundaryDate(date: Date): string {
  const shifted = new Date(date.getTime() - 6 * 60 * 60 * 1000);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${shifted.getUTCFullYear()}-${pad(shifted.getUTCMonth() + 1)}-${pad(shifted.getUTCDate())}`;
}

function toText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value);
}

/** Transaction sheets with valid dates, newest first */
export function getSortedTransactionSheetInfos(): TransactionSheetInfo[] {
  return Budgeting.getTransactionSheetInfos()
    .filter(
      (info) =>
        !Number.isNaN(info.startDate.getTime()) &&
        !Number.isNaN(info.endDate.getTime()),
    )
    .sort((a, b) => b.startDate.getTime() - a.startDate.getTime());
}

/** Gets a transaction sheet by name. Throws `not_found` if it is not a transaction sheet */
export function getTransactionSheet(sheetName: string): Sheet {
  const info = getSortedTransactionSheetInfos().find(
    (i) => i.sheet.getName() === sheetName,
  );
  if (!info)
    throw new ApiError(
      "not_found",
      `No transaction sheet named "${sheetName}"`,
    );
  return info.sheet;
}

/** Raw batch-read data of the whole transaction table plus the DTO of every row */
export type TransactionTable = {
  /** The 7-column range of all TransactionsMax rows */
  range: GoogleAppsScript.Spreadsheet.Range;
  values: unknown[][];
  formulas: string[][];
  /** DTO for every row index, including empty rows */
  rows: TransactionRowDto[];
  /** Whether each row has content in any column but Current Balance */
  hasContent: boolean[];
};

/** Reads the transaction table of the sheet with one batch read per kind of data */
export function readTransactionTable(sheet: Sheet): TransactionTable {
  const sheetVariables = Variables.getSheetVariables(sheet);
  const range = sheet
    .getRange(sheetVariables.TransactionsStart)
    .offset(0, 0, Number(sheetVariables.TransactionsMax), TRANSACTION_COLUMNS);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const displayValues = range.getDisplayValues();
  const notes = range.getNotes();
  const backgrounds = range.getBackgrounds();
  const timeZone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  const hasContent = values.map((rowValues, r) =>
    rowValues.some(
      (value, c) =>
        c !== COL.currentBalance &&
        ((value !== "" && value !== null && value !== undefined) ||
          !!formulas[r][c]),
    ),
  );

  // Split groups exactly like the native split: a group is a row followed by rows that "seem equal" to it
  const groupHeads: (number | null)[] = [];
  const groupSizes = new Map<number, number>();
  let head: number | null = null;
  values.forEach((rowValues, r) => {
    if (!hasContent[r]) {
      head = null;
      groupHeads.push(null);
      return;
    }
    const rowInfo = {
      date: rowValues[COL.date] as GoogleAppsScript.Base.Date | undefined,
      name: (formulas[r][COL.name] || rowValues[COL.name]) as
        string | undefined,
    };
    if (head !== null) {
      const headInfo = {
        date: values[head][COL.date] as GoogleAppsScript.Base.Date | undefined,
        name: (formulas[head][COL.name] || values[head][COL.name]) as
          string | undefined,
      };
      if (!doTransactionRowInfosSeemEqual(rowInfo, headInfo)) head = r;
    } else head = r;
    groupHeads.push(head);
    groupSizes.set(head, (groupSizes.get(head) ?? 0) + 1);
  });

  const rows = values.map((rowValues, r): TransactionRowDto => {
    const costValue = rowValues[COL.cost];
    const background = toText(backgrounds[r][COL.name]).toLowerCase();
    const groupHead = groupHeads[r];
    const dto: Omit<TransactionRowDto, "fingerprint"> = {
      rowIndex: r,
      date: formatCellDate(rowValues[COL.date], timeZone),
      name: toText(rowValues[COL.name]),
      cost:
        typeof costValue === "number" && Number.isFinite(costValue)
          ? costValue
          : null,
      costFormula: formulas[r][COL.cost] || null,
      costDisplay: toText(displayValues[r][COL.cost]),
      envelope: toText(rowValues[COL.envelope]),
      paidWith: toText(rowValues[COL.paidWith]),
      taxDeductible: toText(rowValues[COL.taxDeductible]),
      note: toText(notes[r][COL.name]),
      isError: background === ERROR_BACKGROUND,
      isNote: background === NOTE_BACKGROUND,
      splitGroupId:
        groupHead !== null && (groupSizes.get(groupHead) ?? 0) > 1
          ? String(groupHead)
          : null,
    };
    return { ...dto, fingerprint: computeRowFingerprint(dto) };
  });

  return { range, values, formulas, rows, hasContent };
}

/** Transaction DTOs with trailing empty rows trimmed and interior empty rows as null */
export function toTransactionList(
  table: TransactionTable,
): (TransactionRowDto | null)[] {
  const lastIndex = table.hasContent.lastIndexOf(true);
  return table.rows
    .slice(0, lastIndex + 1)
    .map((row, r) => (table.hasContent[r] ? row : null));
}

/**
 * Gets the range of an optional table from per-sheet variables, or a warning if the variables are
 * missing or invalid
 */
function getOptionalTableRange(
  sheet: Sheet,
  startKey: "EnvelopesStart" | "PaidWithStart",
  maxKey: "EnvelopesMax" | "PaidWithMax",
  numColumns: number,
): { range: GoogleAppsScript.Spreadsheet.Range } | { warning: string } {
  const sheetVariables = Variables.getSheetVariables(sheet);
  const start = sheetVariables[startKey];
  const max = Number(sheetVariables[maxKey]);
  if (!start || !sheetVariables[maxKey])
    return {
      warning: `Sheet "${sheet.getName()}" has no ${startKey}/${maxKey} sheet variables.`,
    };
  if (!Number.isInteger(max) || max < 1 || max > MAX_OPTIONAL_TABLE_ROWS)
    return {
      warning: `Sheet "${sheet.getName()}" has an invalid ${maxKey} sheet variable.`,
    };
  try {
    return {
      range: sheet.getRange(String(start)).offset(0, 0, max, numColumns),
    };
  } catch {
    return {
      warning: `Sheet "${sheet.getName()}" has an invalid ${startKey} sheet variable.`,
    };
  }
}

function isSkippedLabel(name: string, skipLabels: Set<string>) {
  return name === "" || skipLabels.has(name.toLowerCase());
}

/** Option lists and envelope balances of a sheet. `null` where the optional sheet variables are missing */
export type SheetOptions = {
  envelopes: EnvelopeRowDto[] | null;
  envelopeOptions: string[] | null;
  paidWithOptions: string[] | null;
  warnings: string[];
};

/** Reads the envelope table and Paid With options of the sheet (web-plan.md sections 3.5 and 3.6) */
export function readSheetOptions(sheet: Sheet): SheetOptions {
  const warnings: string[] = [];

  let envelopes: EnvelopeRowDto[] | null = null;
  let envelopeOptions: string[] | null = null;
  const envelopeTable = getOptionalTableRange(
    sheet,
    "EnvelopesStart",
    "EnvelopesMax",
    ENVELOPE_COLUMNS,
  );
  if ("range" in envelopeTable) {
    const values = envelopeTable.range.getValues();
    const displayValues = envelopeTable.range.getDisplayValues();
    envelopes = [];
    values.forEach((rowValues, r) => {
      const name = toText(rowValues[0]).trim();
      if (isSkippedLabel(name, ENVELOPE_SKIP_LABELS)) return;
      envelopes!.push({
        name,
        display: displayValues[r].slice(1).map(toText),
        values: rowValues
          .slice(1)
          .map((value) =>
            typeof value === "number" && Number.isFinite(value) ? value : null,
          ),
      });
    });
    envelopeOptions = [...new Set(envelopes.map((envelope) => envelope.name))];
  } else
    warnings.push(
      `${envelopeTable.warning} Envelope balances and Envelope editing are unavailable for it.`,
    );

  let paidWithOptions: string[] | null = null;
  const paidWithTable = getOptionalTableRange(
    sheet,
    "PaidWithStart",
    "PaidWithMax",
    1,
  );
  if ("range" in paidWithTable) {
    paidWithOptions = [
      ...new Set(
        paidWithTable.range
          .getValues()
          .map(([value]) => toText(value).trim())
          .filter((name) => !isSkippedLabel(name, PAID_WITH_SKIP_LABELS)),
      ),
    ];
  } else
    warnings.push(
      `${paidWithTable.warning} Paid with editing is unavailable for it.`,
    );

  return { envelopes, envelopeOptions, paidWithOptions, warnings };
}

/** Builds the `getState` response for the sheet (or the current/newest one when not given) */
export function buildPayPeriodState(sheetName?: string): {
  state: PayPeriodStateDto;
  warnings: string[];
} {
  const infos = getSortedTransactionSheetInfos();
  if (infos.length === 0)
    throw new ApiError("not_found", "No transaction sheets found");

  let info: TransactionSheetInfo | undefined;
  if (sheetName !== undefined) {
    info = infos.find((i) => i.sheet.getName() === sheetName);
    if (!info)
      throw new ApiError(
        "not_found",
        `No transaction sheet named "${sheetName}"`,
      );
  } else {
    const now = new Date();
    info =
      infos.find((i) => i.startDate <= now && now <= i.endDate) ?? infos[0];
  }

  const sheets: SheetInfoDto[] = infos.map((i) => ({
    name: i.sheet.getName(),
    startDate: formatSheetBoundaryDate(i.startDate),
    endDate: formatSheetBoundaryDate(i.endDate),
    isGap: i.sheet.getName().includes("(Gap)"),
  }));

  const variables = Variables.getVariables();
  const options = readSheetOptions(info.sheet);

  return {
    state: {
      sheets,
      sheetName: info.sheet.getName(),
      taxMultiplier: Number(variables.TaxMultiplier),
      forwardName: toText(variables.ForwardName),
      transactions: toTransactionList(readTransactionTable(info.sheet)),
      envelopes: options.envelopes,
      paidWithOptions: options.paidWithOptions,
      envelopeOptions: options.envelopeOptions,
    },
    warnings: options.warnings,
  };
}

/**
 * Builds the `getAllTransactions` response: every transaction tab (newest first) with its
 * transactions and option lists. One batch read per kind of data per tab (no per-row I/O)
 */
export function buildAllTransactions(): AllTransactionsDto {
  const infos = getSortedTransactionSheetInfos();
  if (infos.length === 0)
    throw new ApiError("not_found", "No transaction sheets found");
  return {
    taxMultiplier: Number(Variables.getVariables().TaxMultiplier),
    sheets: infos.map((info) => {
      const options = readSheetOptions(info.sheet);
      return {
        name: info.sheet.getName(),
        startDate: formatSheetBoundaryDate(info.startDate),
        endDate: formatSheetBoundaryDate(info.endDate),
        isGap: info.sheet.getName().includes("(Gap)"),
        transactions: toTransactionList(readTransactionTable(info.sheet)),
        envelopeOptions: options.envelopeOptions,
        paidWithOptions: options.paidWithOptions,
      };
    }),
  };
}
