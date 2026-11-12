/**
 * The demo backend: the real `doPost`/`doGet` from `src/` running on the in-memory GAS mocks that
 * the unit tests use, seeded with fake data. State lives in memory only.
 *
 * Formulas are evaluated by the mock on read (arithmetic and R1C1 relative references only).
 * Envelope balances and other sheet-computed cells are plain numbers in the seed, so they do not
 * update live when transactions change.
 */
import { doGet, doPost } from "../src/index.ts";
import {
  installGasGlobals,
  makeThread,
  type AnySheet,
  type GasEnv,
} from "../test/gas-mocks/index.ts";
import {
  DEMO_API_KEY,
  seedDemo,
  TRANSACTIONS_MAX,
  TRANSACTIONS_START,
} from "./seed.ts";

export { DEMO_API_KEY };

/** Columns of the transaction table the debug endpoints read and write */
const COLUMNS = [
  "date",
  "name",
  "cost",
  "currentBalance",
  "envelope",
  "paidWith",
  "taxDeductible",
] as const;

type Json = Record<string, unknown>;

export class DemoBackend {
  readonly env: GasEnv;

  constructor() {
    this.env = installGasGlobals();
    this.env.spreadsheet.evaluateFormulas = true;
    this.useRealisticGmailSearch();
    seedDemo(this.env);
  }

  /**
   * The real Gmail query excludes processed threads (`in:inbox NOT label:receipts-scripted`).
   * The base mock returns every thread, so filter the same way here.
   */
  private useRealisticGmailSearch() {
    const gmail = this.env.GmailApp;
    gmail.search = (query: string, start?: number, max?: number) => {
      gmail.searches.push({ query, start, max });
      const open = gmail.threads.filter(
        (thread) =>
          !thread.archived &&
          !thread.labels.some(
            (label: { name: string }) => label.name === "Receipts/Scripted",
          ),
      );
      return start === undefined
        ? open
        : open.slice(start, start + (max ?? Infinity));
    };
  }

  /** Replaces all state with the seed (used by `--reset`-style restarts and the e2e tests) */
  reset() {
    seedDemo(this.env);
    this.env.LockService.scriptLock.contended = false;
  }

  /** Runs a request body through the real `doPost`. Returns the JSON text GAS would return */
  post(body: string): string {
    return doPost({ postData: { contents: body } } as never).getContent();
  }

  get(): string {
    return doGet().getContent();
  }

  // #region debug helpers (test-only inspection and mutation, never part of the GAS API)

  private getSheet(name: string): AnySheet {
    const sheet = this.env.spreadsheet.getSheetByName(name);
    if (!sheet) throw new Error(`No sheet named "${name}"`);
    return sheet;
  }

  private transactionRange(sheet: AnySheet) {
    return sheet
      .getRange(TRANSACTIONS_START)
      .offset(0, 0, TRANSACTIONS_MAX, COLUMNS.length);
  }

  listSheets(): string[] {
    return this.env.spreadsheet.getSheets().map((sheet) => sheet.getName());
  }

  /** Raw contents of a sheet's transaction table (values as the mock evaluates them) */
  dumpSheet(name: string) {
    const sheet = this.getSheet(name);
    const range = this.transactionRange(sheet);
    const values = range.getValues() as unknown[][];
    const formulas = range.getFormulas() as string[][];
    const notes = range.getNotes() as string[][];
    const backgrounds = range.getBackgrounds() as (string | null)[][];
    const rows = values.map((rowValues, rowIndex) => ({
      rowIndex,
      values: rowValues.map((value) =>
        value instanceof Date ? value.toISOString() : value,
      ),
      formulas: formulas[rowIndex],
      notes: notes[rowIndex],
      backgrounds: backgrounds[rowIndex],
    }));
    const isUsed = (row: (typeof rows)[number]) =>
      row.values.some((value, c) => c !== 3 && value !== "" && value !== null);
    const lastUsed = rows.reduce(
      (last, row, i) => (isUsed(row) ? i : last),
      -1,
    );
    return { name, columns: COLUMNS, rows: rows.slice(0, lastUsed + 1) };
  }

  /**
   * Edits cells of one transaction row directly, like a person typing in the sheet. `cost` may be
   * a number or an `=` formula
   */
  mutateRow(body: Json) {
    const sheet = this.getSheet(String(body.sheetName));
    const rowIndex = Number(body.rowIndex);
    const changes = (body.changes ?? {}) as Json;
    const range = this.transactionRange(sheet);
    for (const [key, value] of Object.entries(changes)) {
      const col = COLUMNS.indexOf(key as (typeof COLUMNS)[number]);
      if (col < 0) throw new Error(`Unknown column ${key}`);
      range.offset(rowIndex, col, 1, 1).setValues([[value]]);
    }
  }

  /** Adds fake Gmail threads that the next `syncReceipts` will import */
  addThreads(threads: Json[]) {
    for (const t of threads)
      this.env.GmailApp.threads.push(
        makeThread({
          from: String(t.from),
          subject: String(t.subject),
          body: String(t.body ?? ""),
          to: t.to === undefined ? undefined : String(t.to),
          date: t.date ? new Date(String(t.date)) : new Date(),
        }),
      );
  }

  /** Summary of the Gmail threads, for asserting labels and archiving */
  dumpThreads() {
    return this.env.GmailApp.threads.map((thread) => ({
      id: thread.getId(),
      subject: thread.getMessages()[0].getSubject(),
      labels: thread.labels.map((label: { name: string }) => label.name),
      archived: thread.archived,
    }));
  }

  /** Simulates another execution holding the script lock, so mutations answer `busy` */
  setLockContended(contended: boolean) {
    this.env.LockService.scriptLock.contended = contended;
  }

  // #endregion
}
