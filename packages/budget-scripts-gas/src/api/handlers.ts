/**
 * Handlers for the Web App JSON API actions (web-plan.md section 4.4). The router authenticates
 * the request first; handlers validate every field server-side.
 */
import {
  formatSplitCostFormula,
  getSplitGroupKey,
  type AddTransactionResultDto,
  isUnsafeFreeText,
  type AllTransactionsDto,
  type PayPeriodStateDto,
  type SyncReceiptsResultDto,
  type TaxDeductibleValue,
  type TransactionRowDto,
} from "@budget/shared";
import {
  Budgeting,
  type SplitChildSpec,
  type SplitParentSpec,
} from "../util/budgeting.util";
import { tryWithScriptLock } from "../util/lock.util";
import { Variables } from "../util/variables.util";
import {
  ApiError,
  getErrorMessage,
  RECEIPT_IMPORT_BUSY_MESSAGE,
} from "./api-error";
import { mapInvalidConfig } from "./receipt-rules-handlers";
import {
  buildAllTransactions,
  buildPayPeriodState,
  COL,
  getTransactionSheet,
  readSheetOptions,
  readTransactionTable,
  type SheetOptions,
  type TransactionTable,
} from "./serialize";

/** Result of a handler: the response data plus optional warnings */
export type HandlerResult<T> = { data: T; warnings?: string[] };

/** Most split rows one request may create */
const MAX_SPLITS = 50;
/** Number of receipt threads `syncReceipts` imports per call (same as "Import 25 receipts and mark") */
const SYNC_RECEIPTS_MAX_THREADS = 25;

/** Errors thrown by the native split that are caused by the sheet's contents, not by a server fault */
const SPLIT_USER_ERROR_REGEX =
  /has no content|seems to be in a transaction group already|No room on sheet|not enough empty transaction rows/;
/** Errors thrown by receipt recording that the user must fix on the sheet */
const SYNC_USER_ERROR_REGEX =
  /Could not find transaction sheet|not enough empty transaction rows/;

/**
 * A cost formula with a relative reference to a row below in the same column (e.g. `-R[1]C[0]`),
 * i.e. a split parent that subtracts its children. Matched against R1C1 formulas
 */
const SPLIT_PARENT_FORMULA_REGEX = /R\[[1-9]\d*\]C(?:\[0\])?(?![\d[])/;

// #region validation helpers

type Json = Record<string, unknown>;

function badRequest(message: string): never {
  throw new ApiError("bad_request", message);
}

function isObject(value: unknown): value is Json {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requireString(value: unknown, field: string): string {
  if (typeof value !== "string") badRequest(`${field} must be a string`);
  return value;
}

function optionalBoolean(value: unknown, field: string): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "boolean") badRequest(`${field} must be a boolean`);
  return value;
}

/** Free text written to the sheet must not be interpreted as a formula (AGENTS.md section 1.4) */
function requireSafeText(value: unknown, field: string): string {
  const text = requireString(value, field);
  if (isUnsafeFreeText(text))
    badRequest(`${field} must not start with =, +, -, or @`);
  return text;
}

function requireRowIndex(
  value: unknown,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
) {
  const transactionsMax = Number(
    Variables.getSheetVariables(sheet).TransactionsMax,
  );
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < 0 ||
    value >= transactionsMax
  )
    badRequest(`rowIndex must be an integer in [0, ${transactionsMax})`);
  return value;
}

/** Builds a cost cell value with the shared formula builder. Rejects invalid expressions */
function buildCost(expression: string, preTax: boolean, field: string) {
  try {
    return formatSplitCostFormula(
      expression,
      preTax,
      Number(Variables.getVariables().TaxMultiplier),
    );
  } catch {
    return badRequest(`${field} is not a valid expression`);
  }
}

/**
 * Validates an Envelope or Paid with value against the sheet's option list. Blank is always allowed.
 * A non-blank value needs the list (web-plan.md section 3.3: not editable on tabs without it)
 */
function requireOption(
  value: unknown,
  getOptions: () => string[] | null,
  field: string,
  sheetVariableName: string,
): string {
  const text = requireSafeText(value, field);
  if (text === "") return text;
  // Read the option lists only when a non-blank value needs checking
  const options = getOptions();
  if (!options)
    badRequest(
      `${field} cannot be set on this sheet because it has no ${sheetVariableName} sheet variables`,
    );
  if (!options.includes(text))
    badRequest(`${field} "${text}" is not one of the sheet's options`);
  return text;
}

// #endregion

/** Reads the sheet's option lists on first use only */
function sheetOptionsGetter(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): () => SheetOptions {
  let options: SheetOptions | undefined;
  return () => (options ??= readSheetOptions(sheet));
}

/** Gets the sheet and row index of a mutation request and validates them */
function getTarget(request: Json) {
  const sheetName = requireString(request.sheetName, "sheetName");
  const sheet = getTransactionSheet(sheetName);
  const rowIndex = requireRowIndex(request.rowIndex, sheet);
  const baseFingerprint = requireString(
    request.baseFingerprint,
    "baseFingerprint",
  );
  if (!baseFingerprint) badRequest("baseFingerprint must not be empty");
  const force = optionalBoolean(request.force, "force") ?? false;
  return { sheetName, sheet, rowIndex, baseFingerprint, force };
}

/**
 * Runs a mutation under the script lock (`busy` if not available) after re-reading the row and
 * checking its fingerprint (`conflict` with the current row if it changed and not `force`)
 */
function mutateUnderLock<T>(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
  baseFingerprint: string,
  force: boolean,
  mutate: (table: TransactionTable, current: TransactionRowDto) => T,
  baseNote?: string,
): T {
  return tryWithScriptLock(
    () => {
      const table = readTransactionTable(sheet);
      const current = table.rows[rowIndex];
      if (
        !force &&
        (current.fingerprint !== baseFingerprint ||
          (baseNote !== undefined && current.note !== baseNote))
      )
        throw new ApiError(
          "conflict",
          "The transaction changed on the sheet since it was loaded",
          { current },
        );
      if (!table.hasContent[rowIndex]) badRequest(`Row ${rowIndex} is empty`);
      return mutate(table, current);
    },
    () => {
      throw new ApiError(
        "busy",
        "Another change is in progress. Please try again.",
      );
    },
  );
}

/** `getAllTransactions`: the transactions of every pay-period sheet (read only, no lock) */
export function handleGetAllTransactions(): HandlerResult<AllTransactionsDto> {
  return { data: buildAllTransactions() };
}

/** `getState`: the state of one pay-period sheet (read only, no lock) */
export function handleGetState(
  request: Json,
): HandlerResult<PayPeriodStateDto> {
  const sheetName =
    request.sheetName === undefined || request.sheetName === null
      ? undefined
      : requireString(request.sheetName, "sheetName");
  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

type ValidatedChanges = {
  name?: string;
  cost?: string | number;
  envelope?: string;
  paidWith?: string;
  taxDeductible?: TaxDeductibleValue;
  note?: string;
};

const MAX_NOTE_LENGTH = 5000;

const CHANGE_KEYS = new Set([
  "name",
  "costExpression",
  "costPreTax",
  "envelope",
  "paidWith",
  "taxDeductible",
  "note",
]);

function validateChanges(
  changes: unknown,
  getOptions: () => SheetOptions,
): ValidatedChanges {
  if (!isObject(changes)) badRequest("changes must be an object");
  const unknownKeys = Object.keys(changes).filter(
    (key) => !CHANGE_KEYS.has(key),
  );
  if (unknownKeys.length > 0)
    badRequest(`Unknown changes: ${unknownKeys.join(", ")}`);

  const validated: ValidatedChanges = {};
  if (changes.name !== undefined)
    validated.name = requireSafeText(changes.name, "changes.name");

  const costPreTax = optionalBoolean(changes.costPreTax, "changes.costPreTax");
  if (changes.costExpression !== undefined)
    validated.cost = buildCost(
      requireString(changes.costExpression, "changes.costExpression"),
      costPreTax ?? false,
      "changes.costExpression",
    );
  else if (costPreTax !== undefined)
    badRequest("changes.costPreTax requires changes.costExpression");

  if (changes.envelope !== undefined)
    validated.envelope = requireOption(
      changes.envelope,
      () => getOptions().envelopeOptions,
      "changes.envelope",
      "EnvelopesStart/EnvelopesMax",
    );
  if (changes.paidWith !== undefined)
    validated.paidWith = requireOption(
      changes.paidWith,
      () => getOptions().paidWithOptions,
      "changes.paidWith",
      "PaidWithStart/PaidWithMax",
    );
  if (changes.taxDeductible !== undefined) {
    if (changes.taxDeductible !== "Yes" && changes.taxDeductible !== "")
      badRequest('changes.taxDeductible must be "Yes" or ""');
    validated.taxDeductible = changes.taxDeductible;
  }

  if (changes.note !== undefined) {
    const note = requireString(changes.note, "changes.note");
    if (note.length > MAX_NOTE_LENGTH)
      badRequest(`changes.note must be at most ${MAX_NOTE_LENGTH} characters`);
    validated.note = note;
  }

  if (Object.keys(validated).length === 0) badRequest("No changes given");
  return validated;
}

/**
 * `updateTransaction`: edits Name, Cost, Envelope, Paid with, and/or Tax deductible of one row.
 * Never writes Date (col 0) or Current Balance (col 3)
 */
export function handleUpdateTransaction(
  request: Json,
): HandlerResult<TransactionRowDto> {
  const { sheet, rowIndex, baseFingerprint, force } = getTarget(request);
  const changes = validateChanges(request.changes, sheetOptionsGetter(sheet));
  const baseNote =
    changes.note === undefined
      ? undefined
      : requireString(request.baseNote, "baseNote");

  const data = mutateUnderLock(
    sheet,
    rowIndex,
    baseFingerprint,
    force,
    (table, current) => {
      const rowValues = table.values[rowIndex];
      const rowFormulas = table.formulas[rowIndex];
      /** Unchanged cells are written back as read: formula if there is one, else value */
      const asRead = (col: number) => rowFormulas[col] || rowValues[col];

      // A row in a split group may only change the item note after its "|": the text before
      // it must stay the same, or the row would fall out of its group.
      if (
        changes.name !== undefined &&
        current.name.includes("|") &&
        !(
          changes.name.includes("|") &&
          getSplitGroupKey(changes.name) === getSplitGroupKey(current.name)
        )
      )
        badRequest(
          "Cannot rename a row in a split group (its name contains |) because that could break the group",
        );

      if (changes.cost !== undefined) {
        const costFormulaR1C1 = table.range
          .offset(rowIndex, COL.cost, 1, 1)
          .getFormulasR1C1()[0][0];
        if (SPLIT_PARENT_FORMULA_REGEX.test(costFormulaR1C1 ?? ""))
          badRequest(
            "Cannot edit the cost of a split parent row; edit the child rows",
          );
      }

      if (changes.name !== undefined || changes.cost !== undefined)
        table.range
          .offset(rowIndex, COL.name, 1, 2)
          .setValues([
            [
              changes.name ?? asRead(COL.name),
              changes.cost ?? asRead(COL.cost),
            ],
          ]);

      if (
        changes.envelope !== undefined ||
        changes.paidWith !== undefined ||
        changes.taxDeductible !== undefined
      )
        table.range
          .offset(rowIndex, COL.envelope, 1, 3)
          .setValues([
            [
              changes.envelope ?? asRead(COL.envelope),
              changes.paidWith ?? asRead(COL.paidWith),
              changes.taxDeductible ?? asRead(COL.taxDeductible),
            ],
          ]);

      if (changes.note !== undefined)
        table.range.offset(rowIndex, COL.name, 1, 1).setNote(changes.note);

      return readTransactionTable(sheet).rows[rowIndex];
    },
    baseNote,
  );

  return { data };
}

/** Validates the split items and converts them into child specs for `Budgeting.splitTransaction` */
function validateSplits(
  splits: unknown,
  getOptions: () => SheetOptions,
): SplitChildSpec[] {
  if (!Array.isArray(splits) || splits.length < 1 || splits.length > MAX_SPLITS)
    badRequest(`splits must be an array of 1 to ${MAX_SPLITS} items`);
  return splits.map((split, i): SplitChildSpec => {
    const field = `splits[${i}]`;
    if (!isObject(split)) badRequest(`${field} must be an object`);
    const preTax = optionalBoolean(split.preTax, `${field}.preTax`);
    if (preTax === undefined) badRequest(`${field}.preTax must be a boolean`);
    return {
      itemNote: requireSafeText(split.itemNote, `${field}.itemNote`),
      cost: buildCost(
        requireString(split.expression, `${field}.expression`),
        preTax,
        `${field}.expression`,
      ),
      category: requireOption(
        split.envelope,
        () => getOptions().envelopeOptions,
        `${field}.envelope`,
        "EnvelopesStart/EnvelopesMax",
      ),
      type:
        split.paidWith === undefined || split.paidWith === null
          ? undefined
          : requireOption(
              split.paidWith,
              () => getOptions().paidWithOptions,
              `${field}.paidWith`,
              "PaidWithStart/PaidWithMax",
            ) || undefined,
    };
  });
}

/**
 * `splitTransaction`: splits a row (or an already split group) into N new child rows using the
 * native `Budgeting.splitTransaction`. Returns the refreshed state of the sheet
 */
export function handleSplitTransaction(
  request: Json,
): HandlerResult<PayPeriodStateDto> {
  const { sheetName, sheet, rowIndex, baseFingerprint, force } =
    getTarget(request);
  const getOptions = sheetOptionsGetter(sheet);
  const childSpecs = validateSplits(request.splits, getOptions);
  const parentSpec: SplitParentSpec = {};
  if (request.parentItemNote !== undefined)
    parentSpec.itemNote = requireSafeText(
      request.parentItemNote,
      "parentItemNote",
    );
  if (request.parentEnvelope !== undefined)
    parentSpec.category = requireOption(
      request.parentEnvelope,
      () => getOptions().envelopeOptions,
      "parentEnvelope",
      "EnvelopesStart/EnvelopesMax",
    );

  mutateUnderLock(sheet, rowIndex, baseFingerprint, force, () => {
    try {
      Budgeting.splitTransaction(sheet, rowIndex, childSpecs, parentSpec);
    } catch (e) {
      if (SPLIT_USER_ERROR_REGEX.test(getErrorMessage(e)))
        badRequest(getErrorMessage(e));
      throw e;
    }
  });

  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

/** Most rows one `deleteTransaction` request may delete */
const MAX_DELETE_ROWS = 100;

/** Validates `deleteTransaction`'s `rows`: 1..MAX contiguous rows starting at `rowIndex` */
function validateDeleteRows(rows: unknown, rowIndex: number) {
  if (!Array.isArray(rows) || rows.length < 1 || rows.length > MAX_DELETE_ROWS)
    badRequest(`rows must be an array of 1 to ${MAX_DELETE_ROWS} items`);
  return rows.map((row, i) => {
    const field = `rows[${i}]`;
    if (!isObject(row)) badRequest(`${field} must be an object`);
    if (row.rowIndex !== rowIndex + i)
      badRequest(
        "rows must list contiguous row indexes in order, starting at rowIndex",
      );
    const baseFingerprint = requireString(
      row.baseFingerprint,
      `${field}.baseFingerprint`,
    );
    if (!baseFingerprint)
      badRequest(`${field}.baseFingerprint must not be empty`);
    return { rowIndex: rowIndex + i, baseFingerprint };
  });
}

/**
 * `deleteTransaction`: deletes one row, or every row of a split group together. Under the script
 * lock it re-reads the table, checks every row's fingerprint (`conflict` with the first changed
 * row unless `force`), and checks that the rows are non-empty and form exactly one whole split
 * group (or one row not in a group). Then the rows are cleared in place (a gap; no row moves). Returns the refreshed sheet state
 */
export function handleDeleteTransaction(
  request: Json,
): HandlerResult<PayPeriodStateDto> {
  const sheetName = requireString(request.sheetName, "sheetName");
  const sheet = getTransactionSheet(sheetName);
  const rowIndex = requireRowIndex(request.rowIndex, sheet);
  const rows = validateDeleteRows(request.rows, rowIndex);
  const force = optionalBoolean(request.force, "force") ?? false;
  const transactionsMax = Number(
    Variables.getSheetVariables(sheet).TransactionsMax,
  );
  if (rowIndex + rows.length > transactionsMax)
    badRequest(`rows must be within [0, ${transactionsMax})`);

  tryWithScriptLock(
    () => {
      const table = readTransactionTable(sheet);
      if (!force)
        rows.forEach(({ rowIndex: r, baseFingerprint }) => {
          if (table.rows[r].fingerprint !== baseFingerprint)
            throw new ApiError(
              "conflict",
              "The transaction changed on the sheet since it was loaded",
              { current: table.rows[r] },
            );
        });
      rows.forEach(({ rowIndex: r }) => {
        if (!table.hasContent[r]) badRequest(`Row ${r} is empty`);
      });
      const groupId = table.rows[rowIndex].splitGroupId;
      const groupRows =
        groupId === null
          ? [rowIndex]
          : table.rows
              .filter((row) => row.splitGroupId === groupId)
              .map((row) => row.rowIndex);
      if (
        groupRows.length !== rows.length ||
        groupRows.some((r, i) => r !== rows[i].rowIndex)
      )
        badRequest(
          groupId === null
            ? "rows must list only the one row, because it is not in a split group"
            : `rows must list every row of the split group (rows ${groupRows[0]}-${groupRows[groupRows.length - 1]}) and nothing else`,
        );
      Budgeting.deleteTransactionRows(sheet, rowIndex, rows.length);
    },
    () => {
      throw new ApiError(
        "busy",
        "Another change is in progress. Please try again.",
      );
    },
  );

  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

const ADD_FIELD_KEYS = new Set([
  "date",
  "name",
  "costExpression",
  "envelope",
  "paidWith",
  "taxDeductible",
  "note",
]);
const REQUEST_ID_REGEX = /^[A-Za-z0-9_-]{8,64}$/;
const DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/;
/** How long `addTransaction` remembers a requestId (the CacheService maximum: 6 hours) */
const ADD_REQUEST_ID_TTL_SECONDS = 6 * 60 * 60;

/** Parses yyyy-MM-dd as midnight in the spreadsheet time zone. Rejects impossible dates */
function requireDate(value: unknown, field: string): Date {
  const text = requireString(value, field);
  if (!DATE_REGEX.test(text)) badRequest(`${field} must be yyyy-MM-dd`);
  const timeZone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  let date: Date;
  try {
    date = Utilities.parseDate(text, timeZone, "yyyy-MM-dd");
  } catch {
    return badRequest(`${field} is not a valid date`);
  }
  // The parser is lenient (2024-02-30 becomes March 1): require a round trip
  if (Utilities.formatDate(date, timeZone, "yyyy-MM-dd") !== text)
    badRequest(`${field} is not a valid date`);
  return date;
}

/** Validates `addTransaction`'s fields into the cells of the new row */
function validateNewFields(fields: unknown, getOptions: () => SheetOptions) {
  if (!isObject(fields)) badRequest("fields must be an object");
  const unknownKeys = Object.keys(fields).filter(
    (key) => !ADD_FIELD_KEYS.has(key),
  );
  if (unknownKeys.length > 0)
    badRequest(`Unknown fields: ${unknownKeys.join(", ")}`);

  const date = requireDate(fields.date, "fields.date");
  const name = requireSafeText(fields.name, "fields.name");
  if (name.trim() === "") badRequest("fields.name must not be empty");
  const costExpression = requireString(
    fields.costExpression,
    "fields.costExpression",
  );
  if (costExpression.trim() === "")
    badRequest("fields.costExpression must not be empty");
  const cost = buildCost(costExpression, false, "fields.costExpression");
  const envelope =
    fields.envelope === undefined
      ? ""
      : requireOption(
          fields.envelope,
          () => getOptions().envelopeOptions,
          "fields.envelope",
          "EnvelopesStart/EnvelopesMax",
        );
  const paidWith =
    fields.paidWith === undefined
      ? ""
      : requireOption(
          fields.paidWith,
          () => getOptions().paidWithOptions,
          "fields.paidWith",
          "PaidWithStart/PaidWithMax",
        );
  const taxDeductible = fields.taxDeductible ?? "";
  if (taxDeductible !== "Yes" && taxDeductible !== "")
    badRequest('fields.taxDeductible must be "Yes" or ""');
  const note =
    fields.note === undefined ? "" : requireString(fields.note, "fields.note");
  if (note.length > MAX_NOTE_LENGTH)
    badRequest(`fields.note must be at most ${MAX_NOTE_LENGTH} characters`);
  return { date, name, cost, envelope, paidWith, taxDeductible, note };
}

/**
 * `addTransaction`: writes one new row into the first empty row after the last non-empty
 * transaction row (like receipt import; interior gaps are never filled), under the script lock.
 * A requestId seen in the last 6 hours (CacheService) returns the state without inserting again,
 * so a retried request never adds a duplicate. Never writes col 3 (Current Balance) or the split
 * checkboxes. Returns the refreshed state of the sheet and the new row's index
 */
export function handleAddTransaction(
  request: Json,
): HandlerResult<AddTransactionResultDto> {
  const sheetName = requireString(request.sheetName, "sheetName");
  const sheet = getTransactionSheet(sheetName);
  const requestId = requireString(request.requestId, "requestId");
  if (!REQUEST_ID_REGEX.test(requestId))
    badRequest("requestId must be 8-64 characters of A-Z, a-z, 0-9, _ or -");
  const row = validateNewFields(request.fields, sheetOptionsGetter(sheet));
  const cacheKey = `addTransaction:${requestId}`;

  const addedRowIndex = tryWithScriptLock(
    () => {
      const cache = CacheService.getScriptCache();
      const remembered = cache.get(cacheKey);
      if (remembered !== null) return Number(remembered);

      const table = readTransactionTable(sheet);
      const rowIndex = table.hasContent.lastIndexOf(true) + 1;
      if (rowIndex >= table.rows.length)
        badRequest(
          `There is no empty transaction row left at the end of sheet ${sheetName}`,
        );
      table.range
        .offset(rowIndex, COL.date, 1, 3)
        .setValues([[row.date, row.name, row.cost]]);
      table.range
        .offset(rowIndex, COL.envelope, 1, 3)
        .setValues([[row.envelope, row.paidWith, row.taxDeductible]]);
      if (row.note)
        table.range.offset(rowIndex, COL.name, 1, 1).setNote(row.note);
      cache.put(cacheKey, String(rowIndex), ADD_REQUEST_ID_TTL_SECONDS);
      return rowIndex;
    },
    () => {
      throw new ApiError(
        "busy",
        "Another change is in progress. Please try again.",
      );
    },
  );

  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: { ...state, addedRowIndex }, warnings };
}

/**
 * `syncReceipts`: imports and marks up to 25 receipt threads (like "Import 25 receipts and mark")
 * under the script lock. Errors in emails are returned, never shown with `getUi()`. Never creates
 * pay-period sheets: a receipt that fits no sheet is a `bad_request`
 */
export function handleSyncReceipts(
  req: Json,
): HandlerResult<SyncReceiptsResultDto> {
  if (req.mark !== undefined && typeof req.mark !== "boolean")
    badRequest("mark must be a boolean");
  const mark = req.mark ?? true;
  return tryWithScriptLock(
    () => {
      let result: ReturnType<typeof Budgeting.recordReceiptsCore>;
      try {
        // An invalid ReceiptRules variable becomes `invalid_config`
        result = mapInvalidConfig(() =>
          Budgeting.recordReceiptsCore(
            Budgeting.getChaseReceipts(0, SYNC_RECEIPTS_MAX_THREADS),
            mark,
          ),
        );
      } catch (e) {
        if (SYNC_USER_ERROR_REGEX.test(getErrorMessage(e)))
          badRequest(getErrorMessage(e));
        throw e;
      }
      const { state, warnings } = buildPayPeriodState();
      return {
        data: {
          receiptInfosAdded: Object.values(result.receiptInfosAdded).reduce(
            (sum, count) => sum + count,
            0,
          ),
          receiptInfosAddedBySheet: result.receiptInfosAdded,
          numErrors: result.numErrors,
          errors: result.errors,
          state,
        },
        warnings,
      };
    },
    () => {
      throw new ApiError("busy", RECEIPT_IMPORT_BUSY_MESSAGE);
    },
  );
}
