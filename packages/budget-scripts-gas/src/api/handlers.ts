/**
 * Handlers for the Web App JSON API actions (web-plan.md section 4.4). The router authenticates
 * the request first; handlers validate every field server-side.
 */
import {
  buildSplitParentFormula,
  buildSplitParentName,
  formatSplitCostFormula,
  getSplitGroupKey,
  parseAmountInput,
  type AddTransactionResultDto,
  isUnsafeFreeText,
  MAX_SHEETS_PER_READ,
  type AllTransactionsDto,
  type PayPeriodStateDto,
  type SplitGroupResultDto,
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
const SYNC_USER_ERROR_REGEX = /not enough empty transaction rows/;

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

/**
 * `getAllTransactions`: the transactions of every pay-period sheet, or only the named ones when
 * `sheetNames` is given (read only, no lock)
 */
export function handleGetAllTransactions(
  request: Json = {},
): HandlerResult<AllTransactionsDto> {
  let sheetNames: string[] | undefined;
  if (request.sheetNames !== undefined) {
    if (
      !Array.isArray(request.sheetNames) ||
      request.sheetNames.length === 0 ||
      request.sheetNames.length > MAX_SHEETS_PER_READ ||
      request.sheetNames.some((name) => typeof name !== "string")
    )
      badRequest(
        `sheetNames must be an array of 1 to ${MAX_SHEETS_PER_READ} strings`,
      );
    sheetNames = request.sheetNames as string[];
  }
  return { data: buildAllTransactions(sheetNames) };
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
  /** A plain amount or math replacing a split parent's total (`buildSplitParentFormula`) */
  splitTotal?: string;
  envelope?: string;
  paidWith?: string;
  taxDeductible?: TaxDeductibleValue;
  note?: string;
};

const MAX_NOTE_LENGTH = 5000;

const CHANGE_KEYS = new Set([
  "name",
  "splitTotal",
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
  /** Field path in error messages */
  field = "changes",
  /** `splitTotal` is not a per-row change of `updateSplitGroup` */
  allowSplitTotal = true,
): ValidatedChanges {
  if (!isObject(changes)) badRequest(`${field} must be an object`);
  const unknownKeys = Object.keys(changes).filter(
    (key) =>
      !CHANGE_KEYS.has(key) || (key === "splitTotal" && !allowSplitTotal),
  );
  if (unknownKeys.length > 0)
    badRequest(`Unknown ${field}: ${unknownKeys.join(", ")}`);

  const validated: ValidatedChanges = {};
  if (changes.name !== undefined)
    validated.name = requireSafeText(changes.name, `${field}.name`);

  const costPreTax = optionalBoolean(changes.costPreTax, `${field}.costPreTax`);
  if (changes.costExpression !== undefined)
    validated.cost = buildCost(
      requireString(changes.costExpression, `${field}.costExpression`),
      costPreTax ?? false,
      `${field}.costExpression`,
    );
  else if (costPreTax !== undefined)
    badRequest(`${field}.costPreTax requires ${field}.costExpression`);

  if (changes.splitTotal !== undefined) {
    const total = requireString(changes.splitTotal, `${field}.splitTotal`);
    if (!parseAmountInput(total).ok)
      badRequest(`${field}.splitTotal must be a number or valid math`);
    if (validated.cost !== undefined)
      badRequest(
        `${field}.splitTotal cannot be combined with ${field}.costExpression`,
      );
    validated.splitTotal = total;
  }

  if (changes.envelope !== undefined)
    validated.envelope = requireOption(
      changes.envelope,
      () => getOptions().envelopeOptions,
      `${field}.envelope`,
      "EnvelopesStart/EnvelopesMax",
    );
  if (changes.paidWith !== undefined)
    validated.paidWith = requireOption(
      changes.paidWith,
      () => getOptions().paidWithOptions,
      `${field}.paidWith`,
      "PaidWithStart/PaidWithMax",
    );
  if (changes.taxDeductible !== undefined) {
    if (changes.taxDeductible !== "Yes" && changes.taxDeductible !== "")
      badRequest(`${field}.taxDeductible must be "Yes" or ""`);
    validated.taxDeductible = changes.taxDeductible;
  }

  if (changes.note !== undefined) {
    const note = requireString(changes.note, `${field}.note`);
    if (note.length > MAX_NOTE_LENGTH)
      badRequest(`${field}.note must be at most ${MAX_NOTE_LENGTH} characters`);
    validated.note = note;
  }

  if (Object.keys(validated).length === 0) badRequest("No changes given");
  return validated;
}

/**
 * `updateTransaction`: edits Name, Cost, Envelope, Paid with, and/or Tax deductible of one row,
 * or a split parent's total. Writes only that row (renaming a split group is `updateSplitGroup`).
 * Never writes Date (col 0) or Current Balance (col 3).
 *
 * Returns the row, or the whole pay-period state when the row is in a split group and its name,
 * cost, or total changed: the group's other displayed values changed with it (a child's cost
 * moves the parent's remainder). The web client saves split groups with `updateSplitGroup`,
 * which returns only the group's rows
 */
export function handleUpdateTransaction(
  request: Json,
): HandlerResult<TransactionRowDto | PayPeriodStateDto> {
  const { sheetName, sheet, rowIndex, baseFingerprint, force } =
    getTarget(request);
  const changes = validateChanges(request.changes, sheetOptionsGetter(sheet));
  const baseNote =
    changes.note === undefined
      ? undefined
      : requireString(request.baseNote, "baseNote");

  const { row, grouped } = mutateUnderLock(
    sheet,
    rowIndex,
    baseFingerprint,
    force,
    (table, current) => {
      const rowValues = table.values[rowIndex];
      const rowFormulas = table.formulas[rowIndex];
      /** Unchanged cells are written back as read: formula if there is one, else value */
      const asRead = (col: number) => rowFormulas[col] || rowValues[col];

      const groupId = current.splitGroupId;
      /** The first row of its group, or a row in no group */
      const isGroupHead = groupId === null || groupId === String(rowIndex);
      /** Rows of the group below a split parent (contiguous) */
      const childIndexes =
        groupId === String(rowIndex)
          ? table.rows
              .filter(
                (r) => r.rowIndex > rowIndex && r.splitGroupId === groupId,
              )
              .map((r) => r.rowIndex)
          : [];

      const name = changes.name;
      const renamesGroup =
        name !== undefined &&
        getSplitGroupKey(name) !== getSplitGroupKey(current.name);
      // A child row of a split group may only change the item note after its "|": the text
      // before it must stay the same, or the row would fall out of its group. A split parent
      // renames the whole group instead (its children get the new prefix below).
      if (
        name !== undefined &&
        current.name.includes("|") &&
        !isGroupHead &&
        !(name.includes("|") && !renamesGroup)
      )
        badRequest(
          "Cannot rename a row in a split group (its name contains |) because that could break the group",
        );
      // Renaming a split group rewrites its children's names, whose fingerprints this request
      // does not carry: that must go through `updateSplitGroup`, which checks every row.
      if (renamesGroup && childIndexes.length > 0)
        badRequest(
          "Cannot rename a split group with updateTransaction; use updateSplitGroup so every row is checked",
        );

      let cost = changes.cost;
      if (changes.splitTotal !== undefined) {
        const costFormulaR1C1 = table.range
          .offset(rowIndex, COL.cost, 1, 1)
          .getFormulasR1C1()[0][0];
        if (!SPLIT_PARENT_FORMULA_REGEX.test(costFormulaR1C1 ?? ""))
          badRequest(
            "changes.splitTotal is only allowed on a split parent row",
          );
        try {
          cost = buildSplitParentFormula(
            costFormulaR1C1 ?? "",
            changes.splitTotal,
          );
        } catch {
          badRequest(
            "changes.splitTotal cannot be applied: the cost formula does not start with a total",
          );
        }
      }

      if (changes.cost !== undefined) {
        const costFormulaR1C1 = table.range
          .offset(rowIndex, COL.cost, 1, 1)
          .getFormulasR1C1()[0][0];
        if (SPLIT_PARENT_FORMULA_REGEX.test(costFormulaR1C1 ?? ""))
          badRequest(
            "Cannot edit the cost of a split parent row; edit the child rows",
          );
      }

      if (name !== undefined || cost !== undefined)
        table.range
          .offset(rowIndex, COL.name, 1, 2)
          .setValues([[name ?? asRead(COL.name), cost ?? asRead(COL.cost)]]);

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

      const updated = readTransactionTable(sheet).rows[rowIndex];
      return {
        row: updated,
        grouped:
          (groupId !== null || updated.splitGroupId !== null) &&
          (name !== undefined || cost !== undefined),
      };
    },
    baseNote,
  );

  if (!grouped) return { data: row };
  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

/** Per-item fields written after the native split: Tax Deductible (col 6) and the Name note */
interface SplitChildExtras {
  taxDeductible?: TaxDeductibleValue;
  note?: string;
}

/** Validates a split item's optional `taxDeductible` and `note` like `updateTransaction` */
function validateSplitExtras(
  split: Record<string, unknown>,
  field: string,
): SplitChildExtras {
  const extras: SplitChildExtras = {};
  if (split.taxDeductible !== undefined) {
    if (split.taxDeductible !== "Yes" && split.taxDeductible !== "")
      badRequest(`${field}.taxDeductible must be "Yes" or ""`);
    extras.taxDeductible = split.taxDeductible;
  }
  if (split.note !== undefined) {
    const note = requireString(split.note, `${field}.note`);
    if (note.length > MAX_NOTE_LENGTH)
      badRequest(`${field}.note must be at most ${MAX_NOTE_LENGTH} characters`);
    extras.note = note;
  }
  return extras;
}

/**
 * Validates the split items and converts them into child specs for `Budgeting.splitTransaction`,
 * plus each item's extras (written after the native split)
 */
function validateSplits(
  splits: unknown,
  getOptions: () => SheetOptions,
  /** Field name in error messages */
  name = "splits",
): { specs: SplitChildSpec[]; extras: SplitChildExtras[] } {
  if (!Array.isArray(splits) || splits.length < 1 || splits.length > MAX_SPLITS)
    badRequest(`${name} must be an array of 1 to ${MAX_SPLITS} items`);
  const extras: SplitChildExtras[] = [];
  const specs = splits.map((split, i): SplitChildSpec => {
    const field = `${name}[${i}]`;
    if (!isObject(split)) badRequest(`${field} must be an object`);
    const preTax = optionalBoolean(split.preTax, `${field}.preTax`);
    if (preTax === undefined) badRequest(`${field}.preTax must be a boolean`);
    extras.push(validateSplitExtras(split, field));
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
  return { specs, extras };
}

/**
 * Writes the new children's Tax Deductible (col 6) and Name notes after the native split: one
 * batch call each over the contiguous new-child rows, and only when some item supplies them
 */
function writeSplitChildExtras(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  firstChildIndex: number,
  extras: SplitChildExtras[],
) {
  const topLeft = sheet.getRange(
    Variables.getSheetVariables(sheet).TransactionsStart,
  );
  if (extras.some((e) => e.taxDeductible !== undefined))
    topLeft
      .offset(firstChildIndex, COL.taxDeductible, extras.length, 1)
      .setValues(extras.map((e) => [e.taxDeductible ?? ""]));
  if (extras.some((e) => e.note !== undefined))
    topLeft
      .offset(firstChildIndex, COL.name, extras.length, 1)
      .setNotes(extras.map((e) => [e.note ?? ""]));
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
  const { specs: childSpecs, extras } = validateSplits(
    request.splits,
    getOptions,
  );
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
      const groupRange = Budgeting.splitTransaction(
        sheet,
        rowIndex,
        childSpecs,
        parentSpec,
      );
      // The new children are the last rows of the returned group range
      writeSplitChildExtras(
        sheet,
        rowIndex + groupRange.getNumRows() - childSpecs.length,
        extras,
      );
    } catch (e) {
      if (SPLIT_USER_ERROR_REGEX.test(getErrorMessage(e)))
        badRequest(getErrorMessage(e));
      throw e;
    }
  });

  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: state, warnings };
}

/** Most rows one `updateSplitGroup` request may list */
const MAX_GROUP_ROWS = 100;

/** A validated `updateSplitGroup` row */
interface GroupRowEdit {
  rowIndex: number;
  baseFingerprint: string;
  baseNote?: string;
  changes: ValidatedChanges;
}

/** Validates `updateSplitGroup`'s `rows`: 2..MAX contiguous rows in order, each with its fingerprint */
function validateGroupRows(
  rows: unknown,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  getOptions: () => SheetOptions,
): GroupRowEdit[] {
  if (!Array.isArray(rows) || rows.length < 2 || rows.length > MAX_GROUP_ROWS)
    badRequest(`rows must be an array of 2 to ${MAX_GROUP_ROWS} items`);
  const first = isObject(rows[0])
    ? requireRowIndex(rows[0].rowIndex, sheet)
    : badRequest("rows[0] must be an object");
  const transactionsMax = Number(
    Variables.getSheetVariables(sheet).TransactionsMax,
  );
  if (first + rows.length > transactionsMax)
    badRequest(`rows must be within [0, ${transactionsMax})`);
  return rows.map((row, i): GroupRowEdit => {
    const field = `rows[${i}]`;
    if (!isObject(row)) badRequest(`${field} must be an object`);
    if (row.rowIndex !== first + i)
      badRequest("rows must list contiguous row indexes in order");
    const baseFingerprint = requireString(
      row.baseFingerprint,
      `${field}.baseFingerprint`,
    );
    if (!baseFingerprint)
      badRequest(`${field}.baseFingerprint must not be empty`);
    const baseNote =
      row.baseNote === undefined
        ? undefined
        : requireString(row.baseNote, `${field}.baseNote`);
    const changes =
      row.changes === undefined
        ? {}
        : validateChanges(row.changes, getOptions, `${field}.changes`, false);
    if (i === 0 && changes.cost !== undefined)
      badRequest(
        "Cannot edit the cost of a split parent row; edit the child rows",
      );
    return { rowIndex: first + i, baseFingerprint, baseNote, changes };
  });
}

/** Validates `updateSplitGroup`'s `title`: the new text before " | " of every row */
function validateGroupTitle(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  const title = requireSafeText(value, "title").trim();
  if (title === "") badRequest("title must not be empty");
  if (title.includes("|")) badRequest("title must not contain |");
  return title;
}

/**
 * `updateSplitGroup`: saves every edit of one split group atomically (web-plan.md section 4.4).
 * Under the script lock it re-reads the table and checks that the listed rows are still exactly
 * the group (first row = parent) and that every row's fingerprint and `baseNote` match; on any
 * mismatch nothing is written and `conflict` returns the group's current rows. Then each column
 * that changed is written with one batch call over the group (names, costs, Envelope/Paid
 * with/Tax Deductible, notes), and `newItems` are appended with the native split. Returns the
 * group's rows, plus the sheet state only when `newItems` moved rows below the group
 */
export function handleUpdateSplitGroup(
  request: Json,
): HandlerResult<SplitGroupResultDto> {
  const sheetName = requireString(request.sheetName, "sheetName");
  const sheet = getTransactionSheet(sheetName);
  const getOptions = sheetOptionsGetter(sheet);
  const rows = validateGroupRows(request.rows, sheet, getOptions);
  const title = validateGroupTitle(request.title);
  let splitTotal: string | undefined;
  if (request.splitTotal !== undefined) {
    splitTotal = requireString(request.splitTotal, "splitTotal");
    if (!parseAmountInput(splitTotal).ok)
      badRequest("splitTotal must be a number or valid math");
  }
  const newItems =
    request.newItems === undefined
      ? undefined
      : validateSplits(request.newItems, getOptions, "newItems");
  if (
    title === undefined &&
    splitTotal === undefined &&
    newItems === undefined &&
    rows.every((row) => Object.keys(row.changes).length === 0)
  )
    badRequest("No changes given");

  const head = rows[0].rowIndex;
  const count = rows.length;

  const resultRows = tryWithScriptLock(
    () => {
      const table = readTransactionTable(sheet);
      const groupId = table.rows[head].splitGroupId;
      const groupIndexes =
        groupId === String(head)
          ? table.rows
              .filter((row) => row.splitGroupId === groupId)
              .map((row) => row.rowIndex)
          : [];
      const changed = rows.find(({ rowIndex, baseFingerprint, baseNote }) => {
        const current = table.rows[rowIndex];
        return (
          !table.hasContent[rowIndex] ||
          current.fingerprint !== baseFingerprint ||
          (baseNote !== undefined && current.note !== baseNote)
        );
      });
      const sameGroup =
        groupIndexes.length === count &&
        groupIndexes.every((r, i) => r === rows[i].rowIndex);
      if (changed || !sameGroup) {
        const indexes = [
          ...new Set([...rows.map((row) => row.rowIndex), ...groupIndexes]),
        ]
          .filter((r) => table.hasContent[r])
          .sort((a, b) => a - b);
        throw new ApiError(
          "conflict",
          "The split changed on the sheet since it was loaded",
          {
            current: table.rows[changed?.rowIndex ?? head],
            rows: indexes.map((r) => table.rows[r]),
          },
        );
      }

      /** Unchanged cells are written back as read: formula if there is one, else value */
      const asRead = (r: number, col: number) =>
        table.formulas[r][col] || table.values[r][col];
      /** Writes one column block over the group when any row changes it */
      const writeColumns = (
        col: number,
        numColumns: number,
        cells: (row: GroupRowEdit, i: number) => unknown[],
        changes: boolean,
      ) => {
        if (changes)
          table.range
            .offset(head, col, count, numColumns)
            .setValues(rows.map(cells));
      };

      // A row's own name may change only after its "|" (the group's title renames them all)
      rows.forEach(({ rowIndex, changes }, i) => {
        const current = table.rows[rowIndex].name;
        if (
          changes.name !== undefined &&
          !(
            current.includes("|") &&
            changes.name.includes("|") &&
            getSplitGroupKey(changes.name) === getSplitGroupKey(current)
          )
        )
          badRequest(
            `rows[${i}].changes.name must keep the text before | (rename the group with title)`,
          );
      });

      let parentCost: string | undefined;
      if (splitTotal !== undefined) {
        const costFormulaR1C1 = table.range
          .offset(head, COL.cost, 1, 1)
          .getFormulasR1C1()[0][0];
        if (!SPLIT_PARENT_FORMULA_REGEX.test(costFormulaR1C1 ?? ""))
          badRequest(
            "splitTotal is only allowed when the first row is a split parent",
          );
        try {
          parentCost = buildSplitParentFormula(
            costFormulaR1C1 ?? "",
            splitTotal,
          );
        } catch {
          badRequest(
            "splitTotal cannot be applied: the cost formula does not start with a total",
          );
        }
      }

      writeColumns(
        COL.name,
        1,
        ({ rowIndex, changes }, i) => {
          const name = changes.name ?? asRead(rowIndex, COL.name);
          if (title === undefined) return [name];
          const text = String(name);
          const bar = text.indexOf("|");
          // Each row keeps its text after "|"; a parent without "|" gets " | " (AGENTS.md 3.3)
          if (bar < 0) return [i === 0 ? buildSplitParentName(title) : text];
          return [`${title} |${text.slice(bar + 1)}`];
        },
        title !== undefined ||
          rows.some((row) => row.changes.name !== undefined),
      );
      writeColumns(
        COL.cost,
        1,
        ({ rowIndex, changes }, i) => [
          (i === 0 ? parentCost : changes.cost) ?? asRead(rowIndex, COL.cost),
        ],
        parentCost !== undefined ||
          rows.some((row) => row.changes.cost !== undefined),
      );
      writeColumns(
        COL.envelope,
        3,
        ({ rowIndex, changes }) => [
          changes.envelope ?? asRead(rowIndex, COL.envelope),
          changes.paidWith ?? asRead(rowIndex, COL.paidWith),
          changes.taxDeductible ?? asRead(rowIndex, COL.taxDeductible),
        ],
        rows.some(
          ({ changes }) =>
            changes.envelope !== undefined ||
            changes.paidWith !== undefined ||
            changes.taxDeductible !== undefined,
        ),
      );
      if (rows.some((row) => row.changes.note !== undefined))
        table.range
          .offset(head, COL.name, count, 1)
          .setNotes(
            rows.map(({ rowIndex, changes }) => [
              changes.note ?? table.rows[rowIndex].note,
            ]),
          );

      if (newItems)
        try {
          const groupRange = Budgeting.splitTransaction(
            sheet,
            head,
            newItems.specs,
          );
          writeSplitChildExtras(
            sheet,
            head + groupRange.getNumRows() - newItems.specs.length,
            newItems.extras,
          );
        } catch (e) {
          if (SPLIT_USER_ERROR_REGEX.test(getErrorMessage(e)))
            badRequest(getErrorMessage(e));
          throw e;
        }

      const after = readTransactionTable(sheet);
      const afterGroupId = after.rows[head].splitGroupId;
      return afterGroupId === null
        ? [after.rows[head]]
        : after.rows.filter((row) => row.splitGroupId === afterGroupId);
    },
    () => {
      throw new ApiError(
        "busy",
        "Another change is in progress. Please try again.",
      );
    },
  );

  if (!newItems) return { data: { rows: resultRows } };
  const { state, warnings } = buildPayPeriodState(sheetName);
  return { data: { rows: resultRows, state }, warnings };
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
  const timeZone = Budgeting.getSpreadsheetTimeZone();
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
 * pay-period sheets: threads with a receipt that fits no sheet are skipped (left unmarked) and
 * counted in `receiptsWithoutSheet`; the rest are imported
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
          receiptsWithoutSheet: result.skippedReceipts.length,
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
