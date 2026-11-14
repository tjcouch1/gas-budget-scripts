import {
  bustVariablesCache,
  getVariables,
  OPTIONAL_SHEET_VARIABLE_NOTES,
} from "./variables.util";

const VARIABLES_SHEET_NAME = "Variables";
const VARIABLES_RANGE = `${VARIABLES_SHEET_NAME}!A1:B10`;

/** Optional per-sheet keys (web-plan.md §3.3) with a hint for what to put in each */
const OPTIONAL_SHEET_VARIABLES = [
  [
    "EnvelopesStart",
    "first envelope-name data cell of the Envelope table, e.g. J4",
  ],
  ["EnvelopesMax", "max number of envelope rows, e.g. 30"],
  [
    "PaidWithStart",
    "first data cell of the Paid With table (option names), e.g. P4",
  ],
  ["PaidWithMax", "max number of rows in the Paid With table, e.g. 15"],
] as const;

function columnToLetters(column: number) {
  let letters = "";
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

const toA1 = (row: number, column: number) =>
  `${columnToLetters(column)}${row}`;

const isEmptyCell = (value: unknown) =>
  value === "" || value === null || value === undefined;

/**
 * Adds the missing optional per-sheet keys (`EnvelopesStart`, `EnvelopesMax`, `PaidWithStart`,
 * `PaidWithMax`) with empty values to the Template tab's `SheetVariables` block. If the block has
 * too few empty rows, it expands the `SheetVariables` range in the Variables sheet, but only after
 * checking that the new rows are empty on every non-Variables tab (otherwise it throws and changes
 * nothing). Logs which cells to fill in
 */
export function addOptionalSheetVariables() {
  const variables = getVariables();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateName = variables.TemplateName;
  const template = spreadsheet.getSheetByName(templateName);
  if (!template)
    throw new Error(`Template sheet "${templateName}" was not found.`);

  const block = template.getRange(variables.SheetVariables);
  if (block.getNumColumns() < 2)
    throw new Error(
      `SheetVariables range ${variables.SheetVariables} must have at least 2 columns (key, value).`,
    );
  const firstRow = block.getRow();
  const firstColumn = block.getColumn();
  const numRows = block.getNumRows();
  const pairRange = template.getRange(firstRow, firstColumn, numRows, 2);
  const values = pairRange.getValues();
  const formulas = pairRange.getFormulas();

  const present = new Set(values.map(([key]) => String(key)));
  const missing = OPTIONAL_SHEET_VARIABLES.filter(([key]) => !present.has(key));
  if (missing.length === 0) {
    Logger.log(
      `${templateName} already has all optional sheet variables (${OPTIONAL_SHEET_VARIABLES.map(([key]) => key).join(", ")}). Nothing to do.`,
    );
    return;
  }

  const emptyRows = values
    .map((row, r) => (isEmptyCell(row[0]) && isEmptyCell(row[1]) ? r : -1))
    .filter((r) => r >= 0);
  const extraRows = Math.max(0, missing.length - emptyRows.length);

  let sheetVariablesA1 = variables.SheetVariables;
  if (extraRows > 0) {
    const blocked: string[] = [];
    spreadsheet.getSheets().forEach((sheet) => {
      if (sheet.getName() === VARIABLES_SHEET_NAME) return;
      sheet
        .getRange(firstRow + numRows, firstColumn, extraRows, 2)
        .getValues()
        .forEach((row, r) =>
          row.forEach((value, c) => {
            if (!isEmptyCell(value))
              blocked.push(
                `${sheet.getName()}!${toA1(firstRow + numRows + r, firstColumn + c)}`,
              );
          }),
        );
    });
    if (blocked.length > 0)
      throw new Error(
        `Cannot expand SheetVariables (${variables.SheetVariables}) by ${extraRows} row(s): these cells must be empty first: ${blocked.join(", ")}. Nothing was changed.`,
      );

    const variablesRange = spreadsheet.getRange(VARIABLES_RANGE);
    const keyRow = variablesRange
      .getValues()
      .findIndex(([key]) => key === "SheetVariables");
    if (keyRow < 0)
      throw new Error(`SheetVariables was not found in ${VARIABLES_RANGE}.`);
    sheetVariablesA1 = `${toA1(firstRow, firstColumn)}:${toA1(
      firstRow + numRows + extraRows - 1,
      firstColumn + block.getNumColumns() - 1,
    )}`;
    variablesRange.offset(keyRow, 1, 1, 1).setValues([[sheetVariablesA1]]);
  }

  // Keep existing cells (formulas included) and fill empty rows, then new rows, in one batch write
  const output: unknown[][] = values.map((row, r) =>
    row.map((value, c) => formulas[r][c] || value),
  );
  for (let i = 0; i < extraRows; i++) output.push(["", ""]);
  const targetRows = [
    ...emptyRows,
    ...Array.from({ length: extraRows }, (_, i) => numRows + i),
  ].slice(0, missing.length);
  targetRows.forEach((r, i) => (output[r] = [missing[i][0], ""]));
  template
    .getRange(firstRow, firstColumn, numRows + extraRows, 2)
    .setValues(output);

  // Notes on the key cells added by this run only; existing notes in the key column are kept
  const keyRange = template.getRange(
    firstRow,
    firstColumn,
    numRows + extraRows,
    1,
  );
  const notes = keyRange.getNotes();
  targetRows.forEach(
    (r, i) => (notes[r][0] = OPTIONAL_SHEET_VARIABLE_NOTES[missing[i][0]]),
  );
  keyRange.setNotes(notes);

  bustVariablesCache();

  Logger.log(
    [
      `Added ${missing.map(([key]) => key).join(", ")} to ${templateName}${
        extraRows > 0
          ? ` and expanded SheetVariables in ${VARIABLES_SHEET_NAME} to ${sheetVariablesA1}`
          : ""
      }. Fill in these cells on ${templateName}:`,
      ...targetRows.map(
        (r, i) =>
          `  ${templateName}!${toA1(firstRow + r, firstColumn + 1)}: ${missing[i][0]} = ${missing[i][1]}`,
      ),
      `New tabs made by addTransactionSheet copy ${templateName}, so they get these keys. Existing tabs do not: add the same keys there manually (or copy the rows) if you want envelopes and Paid With options on them in the web app. Tabs without them keep working; the app just hides those features.`,
    ].join("\n"),
  );
}

/**
 * Run once from the Apps Script editor: adds the optional per-sheet variables used by the web app
 * to the Template tab
 */
export function setupAddOptionalSheetVariables() {
  addOptionalSheetVariables();
}
