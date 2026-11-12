/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Minimal in-memory mock of the parts of SpreadsheetApp used by src/.
 *
 * Only what the code calls is implemented. Handles that get passed to src/ (sheets, ranges) are typed
 * `any` at the API boundary because the mocks only implement a small subset of the GAS interfaces. Notable simplifications:
 * - Cells are stored sparsely; the grid is unbounded.
 * - A string written via setValues/setValue that starts with "=" is stored as a formula and its
 *   value is the formula text itself (there is no formula evaluation).
 * - copyTo copies value/formula/note/background verbatim (no relative reference adjustment).
 */

export interface Cell {
  value: unknown;
  formula: string;
  note: string;
  background: string | null;
}

const emptyCell = (): Cell => ({
  value: "",
  formula: "",
  note: "",
  background: null,
});

/** Convert an A1 column label (e.g. "AB") to a 1-based number */
export function colToNumber(label: string): number {
  return [...label].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

/** Parse A1 notation like "A1" or "B2:D5" (without sheet prefix) */
export function parseA1(a1: string) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1);
  if (!m) throw new Error(`Bad A1 notation: ${a1}`);
  const col = colToNumber(m[1]);
  const row = Number(m[2]);
  const lastCol = m[3] ? colToNumber(m[3]) : col;
  const lastRow = m[4] ? Number(m[4]) : row;
  return { row, col, numRows: lastRow - row + 1, numCols: lastCol - col + 1 };
}

export class MockRange {
  constructor(
    private readonly sheet: MockSheet,
    private readonly row: number,
    private readonly col: number,
    private readonly numRows: number,
    private readonly numCols: number,
  ) {}

  getSheet(): any {
    return this.sheet;
  }
  getRow() {
    return this.row;
  }
  getColumn() {
    return this.col;
  }
  getLastRow() {
    return this.row + this.numRows - 1;
  }
  getLastColumn() {
    return this.col + this.numCols - 1;
  }
  getNumRows() {
    return this.numRows;
  }
  getNumColumns() {
    return this.numCols;
  }

  /** Relative offset like the real API; size defaults to this range's size */
  offset(
    rowOffset: number,
    colOffset: number,
    numRows?: number,
    numCols?: number,
  ): any {
    return new MockRange(
      this.sheet,
      this.row + rowOffset,
      this.col + colOffset,
      numRows ?? this.numRows,
      numCols ?? this.numCols,
    );
  }

  /** 1-based cell relative to this range's top-left */
  getCell(r: number, c: number): any {
    return new MockRange(this.sheet, this.row + r - 1, this.col + c - 1, 1, 1);
  }

  private map<T>(fn: (cell: Cell, r: number, c: number) => T): T[][] {
    return Array.from({ length: this.numRows }, (_, r) =>
      Array.from({ length: this.numCols }, (_, c) =>
        fn(this.sheet.cell(this.row + r, this.col + c), r, c),
      ),
    );
  }

  getValues() {
    return this.map((cell) => cell.value);
  }
  getDisplayValues() {
    return this.map((cell) => String(cell.value ?? ""));
  }
  getFormulas() {
    return this.map((cell) => cell.formula);
  }
  getValue() {
    return this.getValues()[0][0];
  }
  getNotes() {
    return this.map((cell) => cell.note);
  }
  getBackgrounds() {
    return this.map((cell) => cell.background);
  }

  setValues(values: unknown[][]) {
    if (
      values.length !== this.numRows ||
      values.some((r) => r.length !== this.numCols)
    )
      throw new Error(
        `Dimension mismatch: range is ${this.numRows}x${this.numCols} but data is ${values.length}x${values[0]?.length}`,
      );
    this.map((cell, r, c) => {
      const v = values[r][c];
      if (typeof v === "string" && v.startsWith("=")) {
        cell.formula = v;
        cell.value = v;
      } else {
        cell.formula = "";
        cell.value = v ?? "";
      }
    });
    return this;
  }
  setValue(value: unknown) {
    return this.setValues(
      Array.from({ length: this.numRows }, () =>
        Array(this.numCols).fill(value),
      ),
    );
  }
  setNote(note: string) {
    this.map((cell) => (cell.note = note));
    return this;
  }
  setNotes(notes: (string | null)[][]) {
    this.map((cell, r, c) => (cell.note = notes[r][c] ?? ""));
    return this;
  }
  setBackground(color: string | null) {
    this.map((cell) => (cell.background = color));
    return this;
  }
  setBackgrounds(colors: (string | null)[][]) {
    this.map((cell, r, c) => (cell.background = colors[r][c]));
    return this;
  }
  setTextStyle(style: unknown) {
    this.sheet.textStyles.push({ range: this, style });
    return this;
  }
  insertCheckboxes() {
    this.map((cell) => {
      // Checkboxes start unchecked (FALSE)
      if (cell.value === "") cell.value = false;
    });
    this.sheet.checkboxRanges.push(this);
    return this;
  }
  isChecked() {
    return this.getValue() === true;
  }
  clearContent() {
    this.map((cell) => {
      cell.value = "";
      cell.formula = "";
    });
    return this;
  }
  copyTo(dest: any) {
    // Snapshot first so overlapping source/destination behaves like a real copy
    const snapshot = this.map((cell) => ({ ...cell }));
    dest.map((cell, r, c) => {
      if (r < this.numRows && c < this.numCols)
        Object.assign(cell, snapshot[r][c]);
    });
    return this;
  }
}

export class MockSheet {
  private cells = new Map<string, Cell>();
  tabColor: string | null = null;
  textStyles: { range: MockRange; style: unknown }[] = [];
  checkboxRanges: MockRange[] = [];

  constructor(
    private name: string,
    readonly spreadsheet: MockSpreadsheet,
  ) {}

  /** Get (creating if needed) the cell at a 1-based position */
  cell(row: number, col: number): Cell {
    const key = `${row},${col}`;
    let c = this.cells.get(key);
    if (!c) {
      c = emptyCell();
      this.cells.set(key, c);
    }
    return c;
  }

  getName() {
    return this.name;
  }
  /** Test helper: rename the sheet */
  rename(name: string) {
    this.name = name;
  }
  /** 1-based position among the spreadsheet's sheets */
  getIndex() {
    return this.spreadsheet.sheets.indexOf(this) + 1;
  }
  setTabColor(color: string) {
    this.tabColor = color;
    return this;
  }

  getRange(a: string | number, col?: number, numRows = 1, numCols = 1): any {
    if (typeof a === "string") {
      const p = parseA1(a);
      return new MockRange(this, p.row, p.col, p.numRows, p.numCols);
    }
    return new MockRange(this, a, col!, numRows, numCols);
  }

  /** Test helper: write a 2D block of raw values starting at an A1 cell */
  seed(a1: string, rows: unknown[][]) {
    const p = parseA1(a1);
    this.getRange(p.row, p.col, rows.length, rows[0].length).setValues(rows);
    return this;
  }

  /** Copy all cells from another sheet (used for template duplication) */
  copyFrom(other: MockSheet) {
    other.cells.forEach((cell, key) => this.cells.set(key, { ...cell }));
  }
}

export class MockSpreadsheet {
  sheets: MockSheet[] = [];
  activeSheet: MockSheet | undefined;

  addSheet(name: string): any {
    const sheet = new MockSheet(name, this);
    this.sheets.push(sheet);
    this.activeSheet ??= sheet;
    return sheet;
  }
  getSheets() {
    return [...this.sheets];
  }
  getSheetByName(name: string): any {
    return this.sheets.find((s) => s.getName() === name) ?? null;
  }
  getActiveSheet(): any {
    return this.activeSheet!;
  }
  /** Supports "SheetName!A1:B2" like the real API */
  getRange(a1: string): any {
    const [sheetName, range] = a1.split("!");
    const sheet = this.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Range not found: ${a1}`);
    return sheet.getRange(range);
  }
  insertSheet(
    name: string,
    index: number,
    options?: { template?: MockSheet },
  ): any {
    const sheet = new MockSheet(name, this);
    if (options?.template) sheet.copyFrom(options.template);
    this.sheets.splice(index, 0, sheet);
    return sheet;
  }
}

export class MockMenu {
  items: ({ name: string; functionName: string } | "separator")[] = [];
  added = false;
  constructor(readonly name: string) {}
  addItem(name: string, functionName: string) {
    this.items.push({ name, functionName });
    return this;
  }
  addSeparator() {
    this.items.push("separator");
    return this;
  }
  addToUi() {
    this.added = true;
  }
}

export class MockUi {
  menus: MockMenu[] = [];
  alerts: string[] = [];
  createMenu(name: string) {
    const menu = new MockMenu(name);
    this.menus.push(menu);
    return menu;
  }
  alert(message: string) {
    this.alerts.push(message);
  }
}

export class MockSpreadsheetApp {
  ui = new MockUi();
  /** Set false to make getUi() throw (simulates mobile / trigger context) */
  uiAvailable = true;
  activeRange: any;
  constructor(readonly spreadsheet: MockSpreadsheet) {}

  getActiveSpreadsheet() {
    return this.spreadsheet;
  }
  getActiveSheet(): any {
    return this.spreadsheet.getActiveSheet();
  }
  getUi() {
    if (!this.uiAvailable)
      throw new Error("Cannot call SpreadsheetApp.getUi() from this context.");
    return this.ui;
  }
  newTextStyle() {
    const style = { bold: false };
    const builder = {
      setBold(b: boolean) {
        style.bold = b;
        return builder;
      },
      build: () => ({ ...style }),
    };
    return builder;
  }
  setActiveRange(range: any) {
    this.activeRange = range;
    return range;
  }
}
