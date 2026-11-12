/* eslint-disable @typescript-eslint/no-explicit-any, no-empty-pattern */
import { expect, test as base, type Page } from "@playwright/test";

export const BACKEND = "http://127.0.0.1:8788";
export const GAS_URL = `${BACKEND}/macros/s/demo/exec`;
export const API_KEY = "dev-key";

type DumpRow = {
  rowIndex: number;
  values: unknown[];
  formulas: string[];
  notes: string[];
  backgrounds: (string | null)[];
};

/** Column indexes of the transaction table */
export const COL = {
  date: 0,
  name: 1,
  cost: 2,
  balance: 3,
  envelope: 4,
  paidWith: 5,
  taxDeductible: 6,
} as const;

async function debug<T>(path: string, body?: unknown): Promise<T> {
  const response = await fetch(`${BACKEND}/__debug/${path}`, {
    method: body === undefined ? "GET" : "POST",
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return (await response.json()) as T;
}

/** Calls the real API directly (used to find sheet names and to read state) */
export async function callApi<T = any>(request: object): Promise<T> {
  const response = await fetch(GAS_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ apiVersion: 1, apiKey: API_KEY, ...request }),
  });
  return (await response.json()) as T;
}

export const backend = {
  reset: () => debug("reset", {}),
  sheets: () => debug<string[]>("sheets"),
  /** Name of the current (newest) pay-period sheet */
  async currentSheet(): Promise<string> {
    return (await callApi({ action: "getState" })).data.sheetName;
  },
  async dump(sheetName: string): Promise<DumpRow[]> {
    const dump = await debug<{ rows: DumpRow[] }>(
      `sheet?name=${encodeURIComponent(sheetName)}`,
    );
    return dump.rows;
  },
  /** Edits a row directly, like a person in the sheet */
  mutate: (sheetName: string, rowIndex: number, changes: object) =>
    debug("mutate", { sheetName, rowIndex, changes }),
  threads: () =>
    debug<{ subject: string; labels: string[]; archived: boolean }[]>(
      "threads",
    ),
};

export const test = base.extend<{ sheetName: string }>({
  sheetName: async ({}, use) => {
    await backend.reset();
    await use(await backend.currentSheet());
  },
});

export { expect };

/** Enters the demo URL and key in the Settings dialog, saves, and waits for the list */
export async function connect(page: Page) {
  await page.goto("/");
  const dialog = page.getByRole("dialog", { name: "Connection settings" });
  await dialog.getByLabel("Web app URL").fill(GAS_URL);
  await dialog.getByLabel("API key").fill(API_KEY);
  await dialog.getByRole("button", { name: "Save" }).click();
  await expect(dialog.getByRole("status")).toContainText("Saved");
  await dialog.getByRole("button", { name: "Close", exact: true }).click();
  await expect(page.getByRole("status").first()).toBeVisible();
  await expect(page.getByLabel("Pay period")).toBeVisible();
  await expect(syncedPill(page)).toBeVisible();
  // No stale "enter your credentials" banner from the first load
  await expect(page.getByRole("alert")).toHaveCount(0);
}

export function syncedPill(page: Page) {
  return page.locator('[data-status="synced"]');
}

/** Opens a transaction by its name in the expense list (searches so any filter works) */
export async function openTransaction(page: Page, name: string) {
  await page.getByLabel("Search").fill(name);
  await page
    .getByRole("button", { name: new RegExp(escapeRegExp(name)) })
    .first()
    .click();
  await expect(page.getByRole("dialog")).toBeVisible();
}

function escapeRegExp(text: string) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Index of the first row with this exact name in a sheet dump */
export function findRow(rows: DumpRow[], name: string) {
  const row = rows.find((r) => r.values[COL.name] === name);
  if (!row) throw new Error(`No row named ${name}`);
  return row;
}
