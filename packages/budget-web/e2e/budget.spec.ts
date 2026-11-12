import type { Page } from "@playwright/test";
import { CURRENT_PERIOD } from "../../budget-scripts-gas/test/fixtures/sheets.ts";
import {
  API_KEY,
  backend,
  callApi,
  COL,
  connect,
  expect,
  findRow,
  GAS_URL,
  openTransaction,
  syncedPill,
  test,
} from "./helpers.ts";

const UNCATEGORIZED_COUNT = CURRENT_PERIOD.filter((t) => !t.envelope).length;

test.describe("connection and loading", () => {
  test("enters settings, tests the connection, saves, and loads the state", async ({
    page,
  }) => {
    await page.goto("/");
    const dialog = page.getByRole("dialog", { name: "Connection settings" });
    await expect(dialog).toBeVisible();

    // The production URL rule still rejects non-Google URLs
    await dialog.getByLabel("Web app URL").fill("https://example.com/exec");
    await dialog.getByLabel("API key").fill(API_KEY);
    await dialog.getByRole("button", { name: "Save" }).click();
    await expect(dialog.getByRole("alert")).toContainText("URL must look like");

    await dialog.getByLabel("Web app URL").fill(GAS_URL);
    await dialog.getByRole("button", { name: "Test connection" }).click();
    await expect(dialog.getByRole("status")).toContainText(
      "Connected. 3 pay-period sheets found.",
    );
    await dialog.getByRole("button", { name: "Save" }).click();
    await expect(dialog.getByRole("status")).toContainText("Saved");
    await dialog.getByRole("button", { name: "Close" }).click();

    await expect(page.getByLabel("Pay period").locator("option")).toHaveCount(
      3,
    );
    await expect(
      page.getByRole("tab", { name: /^Uncategorized/ }),
    ).toContainText(String(UNCATEGORIZED_COUNT));
    await expect(page.getByText("Sample Cafe").first()).toBeVisible();
    await expect(syncedPill(page)).toBeVisible();
  });

  test("a wrong API key shows an error", async ({ page }) => {
    await page.goto("/");
    const dialog = page.getByRole("dialog", { name: "Connection settings" });
    await dialog.getByLabel("Web app URL").fill(GAS_URL);
    await dialog.getByLabel("API key").fill("wrong-key");
    await dialog.getByRole("button", { name: "Test connection" }).click();
    await expect(dialog.getByRole("alert")).toContainText(/API key/i);
  });

  test("credentials persist across reloads and are not in localStorage", async ({
    page,
  }) => {
    await connect(page);
    await page.reload();
    await expect(page.getByLabel("Pay period")).toBeVisible();
    await expect(page.getByRole("dialog")).toHaveCount(0);
    const stored = await page.evaluate(() =>
      JSON.stringify(Object.entries(localStorage)),
    );
    expect(stored).not.toContain(API_KEY);
  });
});

test.describe("editing", () => {
  test("categorizes with an envelope chip", async ({ page, sheetName }) => {
    await connect(page);
    await openTransaction(page, "Test Hardware");
    await page
      .getByRole("dialog")
      .getByRole("button", { name: /^Household/ })
      .click();
    await expect(page.getByRole("dialog")).toHaveCount(0);

    await expect
      .poll(async () => {
        const row = findRow(await backend.dump(sheetName), "Test Hardware");
        return row.values[COL.envelope];
      })
      .toBe("Household");
    await expect(syncedPill(page)).toBeVisible();
  });

  test("edits the cost with inline math", async ({ page, sheetName }) => {
    await connect(page);
    await openTransaction(page, "Imaginary Books");
    const dialog = page.getByRole("dialog");
    await dialog.getByLabel("Cost").fill("=1+3.23*2");
    await dialog.getByRole("button", { name: "Save cost" }).click();

    await expect
      .poll(
        async () =>
          findRow(await backend.dump(sheetName), "Imaginary Books").formulas[
            COL.cost
          ],
      )
      .toBe("=1+3.23*2");
    const row = findRow(await backend.dump(sheetName), "Imaginary Books");
    expect(row.values[COL.cost]).toBeCloseTo(7.46);

    // A plain number is stored as a number, not a formula
    await openTransaction(page, "Imaginary Books");
    await page.getByRole("dialog").getByLabel("Cost").fill("20.5");
    await page.getByRole("button", { name: "Save cost" }).click();
    await expect
      .poll(
        async () =>
          findRow(await backend.dump(sheetName), "Imaginary Books").values[
            COL.cost
          ],
      )
      .toBe(20.5);
    expect(
      findRow(await backend.dump(sheetName), "Imaginary Books").formulas[
        COL.cost
      ],
    ).toBe("");
  });

  test("pre-tax cost applies the tax multiplier", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Imaginary Books");
    const dialog = page.getByRole("dialog");
    await dialog.getByLabel("Cost").fill("10");
    await dialog.getByLabel(/Add tax/).check();
    await dialog.getByRole("button", { name: "Save cost" }).click();
    await expect
      .poll(
        async () =>
          findRow(await backend.dump(sheetName), "Imaginary Books").formulas[
            COL.cost
          ],
      )
      .toBe("=1.0825*(10)");
  });

  test("edits name, paid with, and tax deductible", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Faux Bakery");
    const dialog = page.getByRole("dialog");
    await dialog.getByText("Edit details").click();
    await dialog.getByLabel("Name", { exact: true }).fill("=SUM(1)");
    await expect(dialog.getByText("Names cannot start with")).toBeVisible();
    await expect(
      dialog.getByRole("button", { name: "Save details" }),
    ).toBeDisabled();

    await dialog.getByLabel("Name", { exact: true }).fill("Faux Bakery Two");
    await dialog.getByLabel("Paid with").selectOption("Debit");
    await dialog.getByLabel("Tax deductible").check();
    await dialog.getByRole("button", { name: "Save details" }).click();

    await expect
      .poll(async () => {
        const row = (await backend.dump(sheetName))[15];
        return [
          row.values[COL.name],
          row.values[COL.paidWith],
          row.values[COL.taxDeductible],
        ];
      })
      .toEqual(["Faux Bakery Two", "Debit", "Yes"]);
  });
});

test.describe("splitting", () => {
  test("splits with pre-tax off and writes the native formulas", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Jane Example - Lunch money");
    const dialog = page.getByRole("dialog");
    await dialog.getByRole("button", { name: "Split", exact: true }).click();

    await dialog.getByLabel("Item 1 note").fill("Pizza");
    await dialog.getByLabel("Item 1 amount").fill("1+3.23*2");
    await dialog.getByLabel("Item 1 envelope").selectOption("Dining Out");
    await dialog.getByRole("button", { name: "+ Add item" }).click();
    await dialog.getByLabel("Item 2 note").fill("Soda");
    await dialog.getByLabel("Item 2 amount").fill("2");
    await dialog.getByLabel("Item 2 envelope").selectOption("Groceries");
    await expect(dialog.getByTestId("remaining")).toContainText("$5.54");
    await dialog.getByRole("button", { name: "Split", exact: true }).click();

    await expect
      .poll(async () => (await backend.dump(sheetName))[12]?.values[COL.name])
      .toBe("Jane Example - Lunch money | Soda");
    const rows = await backend.dump(sheetName);
    const parent = findRow(rows, "Jane Example - Lunch money | ");
    expect(parent.rowIndex).toBe(10);
    expect(parent.formulas[COL.cost]).toBe("=15-R[1]C[0]-R[2]C[0]");
    const pizza = rows[11];
    expect(pizza.values[COL.name]).toBe("Jane Example - Lunch money | Pizza");
    expect(pizza.formulas[COL.cost]).toBe("=1+3.23*2");
    expect(pizza.values[COL.envelope]).toBe("Dining Out");
    expect(pizza.values[COL.paidWith]).toBe("TJ Venmo");
    const soda = rows[12];
    // A single numeric literal is stored as a plain number
    expect(soda.values[COL.cost]).toBe(2);
    expect(soda.formulas[COL.cost]).toBe("");
    expect(soda.values[COL.envelope]).toBe("Groceries");
    // The rows below moved down by two
    expect(rows[13].values[COL.name]).toBe("Fictional Grocer");
    await expect(syncedPill(page)).toBeVisible();
  });

  test("splits with pre-tax on and multiplies by the tax multiplier", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Dummy Department Store");
    const dialog = page.getByRole("dialog");
    await dialog.getByRole("button", { name: "Split", exact: true }).click();
    await dialog.getByLabel("Item 1 note").fill("Napkins");
    await dialog.getByLabel("Item 1 amount").fill("4.23*2+12.09");
    await dialog
      .getByLabel(/Add tax/)
      .first()
      .check();
    await dialog.getByLabel("Item 1 envelope").selectOption("Household");
    await dialog.getByRole("button", { name: "+ Add item" }).click();
    await dialog.getByLabel("Item 2 note").fill("Gift");
    await dialog.getByLabel("Item 2 amount").fill("10");
    await dialog.getByLabel("Item 2 envelope").selectOption("Gifts");
    await dialog.getByRole("button", { name: "Split", exact: true }).click();

    await expect
      .poll(async () =>
        (await backend.dump(sheetName)).some(
          (r) => r.values[COL.name] === "Dummy Department Store | Napkins",
        ),
      )
      .toBe(true);
    const rows = await backend.dump(sheetName);
    const napkins = findRow(rows, "Dummy Department Store | Napkins");
    expect(napkins.formulas[COL.cost]).toBe("=1.0825*(4.23*2+12.09)");
    expect(napkins.values[COL.cost]).toBeCloseTo(1.0825 * 20.55);
    const gift = findRow(rows, "Dummy Department Store | Gift");
    expect(gift.values[COL.cost]).toBe(10);
    expect(gift.rowIndex).toBe(napkins.rowIndex + 1);
  });

  test("names of rows in a split group cannot be edited", async ({ page }) => {
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | Paper towels");
    const dialog = page.getByRole("dialog");
    await dialog.getByText("Edit details").click();
    await expect(
      dialog.getByText("Names of split rows cannot be edited here."),
    ).toBeVisible();
  });
});

test.describe("conflicts", () => {
  async function editRemotelyThenChooseEnvelope(page: Page, sheetName: string) {
    await connect(page);
    await openTransaction(page, "Test Hardware");
    // Someone edits the row in the sheet after the app loaded it
    await backend.mutate(sheetName, 6, {
      name: "Test Hardware Remote",
      cost: 30,
    });
    await page
      .getByRole("dialog")
      .getByRole("button", { name: /^Household/ })
      .click();
    const conflict = page.getByRole("dialog", {
      name: "This row changed in the sheet",
    });
    await expect(conflict).toBeVisible();
    await expect(
      conflict.getByRole("region", { name: "Original Value" }),
    ).toContainText("Test Hardware");
    await expect(
      conflict.getByRole("region", { name: "Discarded Local Change" }),
    ).toContainText("Household");
    await expect(
      conflict.getByRole("region", { name: "Current Remote Value" }),
    ).toContainText("Test Hardware Remote");
    return conflict;
  }

  test("Accept Remote keeps the sheet's value", async ({ page, sheetName }) => {
    const conflict = await editRemotelyThenChooseEnvelope(page, sheetName);
    await conflict.getByRole("button", { name: "Accept Remote" }).click();
    await expect(conflict).toHaveCount(0);
    await expect(page.getByText("Test Hardware Remote").first()).toBeVisible();
    const row = (await backend.dump(sheetName))[6];
    expect(row.values[COL.name]).toBe("Test Hardware Remote");
    expect(row.values[COL.envelope]).toBe("");
  });

  test("Overwrite with Mine writes the local change over the remote row", async ({
    page,
    sheetName,
  }) => {
    const conflict = await editRemotelyThenChooseEnvelope(page, sheetName);
    await conflict.getByRole("button", { name: "Overwrite with Mine" }).click();
    await expect(conflict).toHaveCount(0);
    await expect
      .poll(async () => (await backend.dump(sheetName))[6].values[COL.envelope])
      .toBe("Household");
    // Only the field the user changed was written
    expect((await backend.dump(sheetName))[6].values[COL.name]).toBe(
      "Test Hardware Remote",
    );
    await expect(syncedPill(page)).toBeVisible();
  });
});

test.describe("offline", () => {
  test("queues a change offline and syncs it on reconnect", async ({
    page,
    context,
    sheetName,
  }) => {
    await connect(page);
    await context.setOffline(true);
    await page.evaluate(() => window.dispatchEvent(new Event("offline")));
    await expect(page.locator('[data-status="offline"]')).toBeVisible();

    await openTransaction(page, "Imaginary Books");
    await page
      .getByRole("dialog")
      .getByRole("button", { name: /^Entertainment/ })
      .click();
    await expect(page.locator('[data-status="offline"]')).toContainText(
      "1 queued",
    );
    expect(
      findRow(await backend.dump(sheetName), "Imaginary Books").values[
        COL.envelope
      ],
    ).toBe("");

    await context.setOffline(false);
    await expect
      .poll(
        async () =>
          findRow(await backend.dump(sheetName), "Imaginary Books").values[
            COL.envelope
          ],
      )
      .toBe("Entertainment");
    await expect(syncedPill(page)).toBeVisible();
  });
});

test.describe("receipts", () => {
  test("Import receipts adds the fake Gmail receipts as rows", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    const before = (await backend.dump(sheetName)).length;
    await page.getByRole("button", { name: "Sync receipts" }).click();
    await expect(page.getByText("Demo Cafe").first()).toBeVisible();

    const rows = await backend.dump(sheetName);
    expect(rows.length).toBe(before + 4);
    const names = rows.map((r) => String(r.values[COL.name]));
    expect(names).toContain("Demo Cafe");
    expect(names).toContain("Jane Doe - Pizza night");
    expect(names.some((n) => n.startsWith("Fake Fuel 42"))).toBe(true);
    const threads = await backend.threads();
    expect(threads.filter((t) => t.archived)).toHaveLength(4);
    expect(
      threads.find((t) => t.subject === "Your statement is ready")?.labels,
    ).toContain("Receipts/Ignored");

    // A second import finds nothing new
    await page.getByRole("button", { name: "Sync receipts" }).click();
    await expect(syncedPill(page)).toBeVisible();
    expect((await backend.dump(sheetName)).length).toBe(before + 4);
  });
});

test.describe("sheets without the optional variables", () => {
  test("an older tab degrades gracefully", async ({ page }) => {
    await connect(page);
    const select = page.getByLabel("Pay period");
    const oldest = await select.locator("option").last().getAttribute("value");
    await select.selectOption(oldest!);

    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Test Hardware");
    const dialog = page.getByRole("dialog");
    await expect(
      dialog.getByText("This sheet has no envelope list"),
    ).toBeVisible();
    await expect(
      dialog.getByRole("button", { name: /^Household/ }),
    ).toHaveCount(0);
    await dialog.getByText("Edit details").click();
    await expect(dialog.getByLabel("Paid with")).toHaveCount(0);
    await dialog.getByRole("button", { name: "Close" }).click();

    await page.getByRole("button", { name: "Envelopes" }).click();
    await expect(
      page.getByText("No envelope data for this pay period."),
    ).toBeVisible();
  });
});

test.describe("envelopes", () => {
  test("shows the balances per envelope", async ({ page }) => {
    await connect(page);
    await page.getByRole("button", { name: "Envelopes" }).click();
    await expect(page.getByTestId("remain-Groceries")).toBeVisible();
    await page.getByLabel("Show all columns").check();
    await expect(page.getByRole("table")).toContainText("Spent YTD");
  });
});

test.describe("all pay periods", () => {
  test("shows uncategorized rows of every tab and edits a row of another tab", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    // Not loaded until chosen
    await expect(page.getByText("Faux Bakery")).toHaveCount(1); // current tab only
    const all = (await callApi({ action: "getAllTransactions" })).data;
    const names: string[] = all.sheets.map((s: { name: string }) => s.name);
    expect(names).toHaveLength(3);
    const previous = names[1];

    await page
      .getByLabel("Transactions shown")
      .selectOption({ label: "All pay periods" });
    // One heading per tab, newest first
    for (const name of names)
      await expect(page.getByRole("heading", { name })).toBeVisible();
    // The previous tab's uncategorized "Faux Bakery" is now listed too
    await expect(page.getByText("Faux Bakery")).toHaveCount(2);
    // Searching spans every tab
    await page.getByLabel("Search").fill("Test Hardware");
    await expect(page.getByRole("heading", { name: names[2] })).toBeVisible();

    await page.getByLabel("Search").fill("Faux Bakery");
    await page
      .getByRole("heading", { name: previous })
      .locator("xpath=following-sibling::div")
      .getByRole("button", { name: /Faux Bakery/ })
      .click();
    const dialog = page.getByRole("dialog");
    await dialog.getByRole("button", { name: /Groceries/ }).click();

    await expect
      .poll(
        async () =>
          findRow(await backend.dump(previous), "Faux Bakery").values[
            COL.envelope
          ],
      )
      .toBe("Groceries");
    // The current tab's own Faux Bakery is untouched
    expect(
      findRow(await backend.dump(sheetName), "Faux Bakery").values[
        COL.envelope
      ],
    ).toBe("");
    await expect(syncedPill(page)).toBeVisible();
  });

  test("splits a row of another tab", async ({ page, sheetName }) => {
    expect(sheetName).toBeTruthy(); // the fixture resets the backend
    await connect(page);
    const all = (await callApi({ action: "getAllTransactions" })).data;
    const previous: string = all.sheets[1].name;
    await page
      .getByLabel("Transactions shown")
      .selectOption({ label: "All pay periods" });
    await page.getByLabel("Search").fill("Faux Bakery");
    const dialog = page.getByRole("dialog");
    await page
      .getByRole("heading", { name: previous })
      .locator("xpath=following-sibling::div")
      .getByRole("button", { name: /Faux Bakery/ })
      .click();
    await dialog.getByRole("button", { name: "Split", exact: true }).click();
    await dialog.getByLabel("Item 1 note").fill("Rolls");
    await dialog.getByLabel("Item 1 amount").fill("2.5");
    await expect(dialog.getByTestId("remaining")).toContainText("$3.75");
    await dialog.getByRole("button", { name: "Split", exact: true }).click();
    await expect
      .poll(async () =>
        (await backend.dump(previous)).some(
          (r) => r.values[COL.name] === "Faux Bakery | Rolls",
        ),
      )
      .toBe(true);
    const rows = await backend.dump(previous);
    const parent = findRow(rows, "Faux Bakery | ");
    expect(parent.formulas[COL.cost]).toBe("=6.25-R[1]C[0]");
  });
});
