import { DEFAULT_RECEIPT_RULES } from "@budget/shared";
import { isEmptyCell, toA1 } from "./a1.util";
import {
  bustVariablesCache,
  GLOBAL_VARIABLE_NOTES,
  OPTIONAL_VARIABLE_NAMES,
  VARIABLES_RANGE,
  VARIABLES_SHEET_NAME,
} from "./variables.util";

type OptionalVariableName = (typeof OPTIONAL_VARIABLE_NAMES)[number];
type Range = GoogleAppsScript.Spreadsheet.Range;

/** Value written by the setup helper for each optional global key */
const OPTIONAL_VARIABLE_DEFAULTS: Record<OptionalVariableName, () => string> = {
  ReceiptRules: () => JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2),
};

function readVariablesRange() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getRange(VARIABLES_RANGE);
  return { range, values: range.getValues() };
}

/**
 * Row indexes (0-based in VARIABLES_RANGE) for `keys`: the first rows below the last key whose key
 * AND value cells are both empty. Throws, before anything is written, when there are too few
 */
function findFreeRows(values: unknown[][], keys: readonly string[]) {
  let lastKeyRow = -1;
  values.forEach(([key], r) => {
    if (!isEmptyCell(key)) lastKeyRow = r;
  });
  const rows: number[] = [];
  for (let r = lastKeyRow + 1; r < values.length; r++)
    if (
      rows.length < keys.length &&
      isEmptyCell(values[r][0]) &&
      isEmptyCell(values[r][1])
    )
      rows.push(r);
  if (rows.length < keys.length)
    throw new Error(
      `No empty row left in ${VARIABLES_RANGE} for ${keys.join(", ")}. Make room below the last variable (both columns empty) and try again. Nothing was changed.`,
    );
  return rows;
}

/** A1 address on the Variables sheet of the cell at (row, column) offsets inside VARIABLES_RANGE */
function variablesCellA1(
  range: Range,
  rowOffset: number,
  columnOffset: number,
) {
  return `${VARIABLES_SHEET_NAME}!${toA1(range.getRow() + rowOffset, range.getColumn() + columnOffset)}`;
}

/**
 * Finds the key's value Range in VARIABLES_RANGE. When the key is missing, writes the key (with its
 * GLOBAL_VARIABLE_NOTES note on the key cell) into the first row below the last key whose key and
 * value cells are both empty, and leaves the value empty. Throws if there is no such row
 */
export function ensureGlobalVariableRow(key: OptionalVariableName): {
  valueRange: Range;
  added: boolean;
} {
  const { range, values } = readVariablesRange();
  const existing = values.findIndex(([k]) => k === key);
  if (existing >= 0)
    return { valueRange: range.offset(existing, 1, 1, 1), added: false };

  const [row] = findFreeRows(values, [key]);
  const keyCell = range.offset(row, 0, 1, 1);
  keyCell.setValues([[key]]);
  keyCell.setNote(GLOBAL_VARIABLE_NOTES[key]);
  bustVariablesCache();
  return { valueRange: range.offset(row, 1, 1, 1), added: true };
}

/**
 * Adds every missing optional global variable (`ReceiptRules`) to the Variables sheet with its
 * default value (the pretty-printed built-in receipt rules) and a note on the key cell
 * (GLOBAL_VARIABLE_NOTES). Existing keys, values and notes are kept. Throws and changes nothing when
 * VARIABLES_RANGE has too few empty rows below the last key
 */
export function addOptionalGlobalVariables() {
  const { range, values } = readVariablesRange();
  const present = new Set(values.map(([key]) => String(key)));
  const missing = OPTIONAL_VARIABLE_NAMES.filter((key) => !present.has(key));
  if (missing.length === 0) {
    Logger.log(
      `${VARIABLES_SHEET_NAME} already has all optional global variables (${OPTIONAL_VARIABLE_NAMES.join(", ")}). Nothing to do.`,
    );
    return;
  }

  const rows = findFreeRows(values, missing);
  // Rows may not be contiguous (stray content is skipped): one write per added row
  rows.forEach((r, i) =>
    range
      .offset(r, 0, 1, 2)
      .setValues([[missing[i], OPTIONAL_VARIABLE_DEFAULTS[missing[i]]()]]),
  );

  // Notes on the key cells added by this run only; existing notes in the key column are kept
  const keyColumn = range.offset(0, 0, values.length, 1);
  const notes = keyColumn.getNotes();
  rows.forEach((r, i) => (notes[r][0] = GLOBAL_VARIABLE_NOTES[missing[i]]));
  keyColumn.setNotes(notes);

  bustVariablesCache();

  Logger.log(
    [
      ...rows.map(
        (r, i) =>
          `Added ${missing[i]} to ${variablesCellA1(range, r, 0)} with the built-in default value in ${variablesCellA1(range, r, 1)} (see the note on the key cell).`,
      ),
      "Edit the rules in the web app (Settings > Customize variables), or clear the value cell to use the built-in defaults.",
    ].join("\n"),
  );
}

/**
 * Run once from the Apps Script editor: adds the optional global variables used by the web app
 * (`ReceiptRules`) to the Variables sheet
 */
export function setupAddOptionalGlobalVariables() {
  addOptionalGlobalVariables();
}
