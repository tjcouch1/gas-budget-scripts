/**
 * Handlers for the Web App JSON API actions (web-plan.md section 4.4). The router authenticates
 * the request first; handlers validate every field server-side.
 */
import {
  formatSplitCostFormula,
  isUnsafeFreeText,
  type PayPeriodStateDto,
  type SyncReceiptsResultDto,
  type TaxDeductibleValue,
  type TransactionRowDto,
} from "@budget/shared";
import { Budgeting, type SplitChildSpec } from "../util/budgeting.util";
import { tryWithScriptLock } from "../util/lock.util";
import { Variables } from "../util/variables.util";
import { ApiError, getErrorMessage } from "./api-error";
import {
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
): T {
  return tryWithScriptLock(
    () => {
      const table = readTransactionTable(sheet);
      const current = table.rows[rowIndex];
      if (current.fingerprint !== baseFingerprint && !force)
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
};

const CHANGE_KEYS = new Set([
  "name",
  "costExpression",
  "costPreTax",
  "envelope",
  "paidWith",
  "taxDeductible",
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

      if (changes.name !== undefined && current.name.includes("|"))
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

      return readTransactionTable(sheet).rows[rowIndex];
    },
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
  const childSpecs = validateSplits(request.splits, sheetOptionsGetter(sheet));

  mutateUnderLock(sheet, rowIndex, baseFingerprint, force, () => {
    try {
      Budgeting.splitTransaction(sheet, rowIndex, childSpecs);
    } catch (e) {
      if (SPLIT_USER_ERROR_REGEX.test(getErrorMessage(e)))
        badRequest(getErrorMessage(e));
      throw e;
    }
  });

  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

/**
 * `syncReceipts`: imports and marks up to 25 receipt threads (like "Import 25 receipts and mark")
 * under the script lock. Errors in emails are returned, never shown with `getUi()`. Never creates
 * pay-period sheets: a receipt that fits no sheet is a `bad_request`
 */
export function handleSyncReceipts(): HandlerResult<SyncReceiptsResultDto> {
  return tryWithScriptLock(
    () => {
      let result: ReturnType<typeof Budgeting.recordReceiptsCore>;
      try {
        const threadList = Budgeting.getChaseReceipts(
          0,
          SYNC_RECEIPTS_MAX_THREADS,
        );
        result = Budgeting.recordReceiptsCore(threadList, true);
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
      throw new ApiError(
        "busy",
        "Another change or receipt import is in progress. Please try again.",
      );
    },
  );
}
