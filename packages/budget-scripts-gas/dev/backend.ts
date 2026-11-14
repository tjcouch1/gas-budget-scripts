/**
 * The demo backend: the real `doPost`/`doGet` from `src/` running on the in-memory GAS mocks that
 * the unit tests use, seeded with fake data. State lives in memory only.
 *
 * Formulas are evaluated by the mock on read (arithmetic and R1C1 relative references only).
 * Envelope balances and other sheet-computed cells are plain numbers in the seed, so they do not
 * update live when transactions change.
 */
import { extractEmailAddress } from "@budget/shared";
import { doGet, doPost } from "../src/index.ts";
import {
  bustVariablesCache,
  VARIABLES_RANGE,
} from "../src/util/variables.util.ts";
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
    bustVariablesCache();
  }

  /**
   * Filters threads like Gmail would for the queries `buildReceiptSearchQuery` emits (the base mock
   * returns every thread): `from:(a OR b)` keeps threads whose first message's address is in the set
   * (case-insensitive), `subject:("x" OR "y")` keeps subjects containing any term
   * (case-insensitive), and `in:inbox` drops archived and `Receipts/Scripted` threads.
   * `newer_than:` and the label exclusions other than Receipts/Scripted are ignored
   */
  private useRealisticGmailSearch() {
    const gmail = this.env.GmailApp;
    gmail.search = (query: string, start?: number, max?: number) => {
      gmail.searches.push({ query, start, max });
      const fromMatch = /from:\(([^)]*)\)/.exec(query);
      const senders = fromMatch
        ? new Set(fromMatch[1].split(" OR ").map((a) => a.toLowerCase()))
        : undefined;
      const subjectMatch = /subject:\(([^)]*)\)/.exec(query);
      const subjects = subjectMatch
        ? [...subjectMatch[1].matchAll(/"([^"]*)"/g)].map(([, s]) =>
            s.toLowerCase(),
          )
        : undefined;
      const inboxOnly = query.includes("in:inbox");
      const matching = gmail.threads.filter((thread) => {
        const first = thread.getMessages()[0];
        if (
          senders &&
          !senders.has(extractEmailAddress(first.getFrom()).toLowerCase())
        )
          return false;
        const subject = String(first.getSubject()).toLowerCase();
        if (subjects && !subjects.some((s) => subject.includes(s)))
          return false;
        return (
          !inboxOnly ||
          (!thread.archived &&
            !thread.labels.some(
              (label: { name: string }) => label.name === "Receipts/Scripted",
            ))
        );
      });
      return start === undefined
        ? matching
        : matching.slice(start, start + (max ?? Infinity));
    };
  }

  /** Replaces all state with the seed (used by `--reset`-style restarts and the e2e tests) */
  reset() {
    seedDemo(this.env);
    this.env.LockService.scriptLock.contended = false;
  }

  /** Runs a request body through the real `doPost`. Returns the JSON text GAS would return */
  post(body: string): string {
    bustVariablesCache(); // every real GAS execution starts with empty module caches
    return doPost({ postData: { contents: body } } as never).getContent();
  }

  get(): string {
    bustVariablesCache();
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
          html: t.html === undefined ? undefined : String(t.html),
          to: t.to === undefined ? undefined : String(t.to),
          date: t.date ? new Date(String(t.date)) : new Date(),
        }),
      );
  }

  /** `[key, value]` rows of the Variables sheet's variables range (rows with an empty key are skipped) */
  dumpVariables(): [string, unknown][] {
    return (
      this.env.spreadsheet.getRange(VARIABLES_RANGE).getValues() as unknown[][]
    )
      .filter(([key]) => key !== "" && key !== null)
      .map(([key, value]) => [
        String(key),
        value instanceof Date ? value.toISOString() : value,
      ]);
  }

  /** Writes a Variables sheet value directly (bypasses validation; adds the key in the first empty row if missing) */
  setVariable(key: string, value: string) {
    const range = this.env.spreadsheet.getRange(VARIABLES_RANGE);
    const keys = range.getValues() as unknown[][];
    let row = keys.findIndex(([k]) => k === key);
    if (row < 0) row = keys.findIndex(([k]) => k === "" || k === null);
    if (row < 0) throw new Error(`No room for variable ${key}`);
    range.offset(row, 0, 1, 2).setValues([[key, value]]);
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
