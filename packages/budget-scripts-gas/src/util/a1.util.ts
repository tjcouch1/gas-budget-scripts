/** 1-based column number to letters (1 => A, 27 => AA) */
export function columnToLetters(column: number) {
  let letters = "";
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

/** 1-based row and column to A1 notation */
export const toA1 = (row: number, column: number) =>
  `${columnToLetters(column)}${row}`;

/** Whether a cell value from getValues() is empty */
export const isEmptyCell = (value: unknown) =>
  value === "" || value === null || value === undefined;
