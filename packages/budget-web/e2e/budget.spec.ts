import type { Page } from "@playwright/test";
import { receiptThreadFixtures } from "../../budget-scripts-gas/test/fixtures/receipts.ts";
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
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy(); // the fixture resets the backend (other spec files edit it)
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
    await dialog.getByRole("button", { name: "Close", exact: true }).click();

    // Three pay-period tabs plus the "All" option
    await expect(page.getByLabel("Pay period").locator("option")).toHaveCount(
      4,
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
    // Unsplit rows have no tax toggle
    await expect(dialog.getByLabel(/Add tax/)).toHaveCount(0);
    await dialog.getByLabel("Cost").fill("=1+3.23*2");
    await dialog.getByRole("button", { name: "Save & close" }).click();

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
    await page.getByRole("button", { name: "Save & close" }).click();
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

  test("adds a note to a row from Edit details", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Faux Bakery");
    const dialog = page.getByRole("dialog");
    await expect(dialog.getByLabel("Note", { exact: true })).toHaveCount(0);
    await dialog.getByText("Edit details").click();
    await dialog.getByRole("button", { name: "Add note" }).click();
    await expect(dialog.getByLabel("Note", { exact: true })).toBeFocused();
    await page.keyboard.type("Paid back by Pat");
    await dialog.getByRole("button", { name: "Save & close" }).click();
    await expect
      .poll(
        async () =>
          findRow(await backend.dump(sheetName), "Faux Bakery").notes[COL.name],
      )
      .toBe("Paid back by Pat");

    // The note shows when the row is opened again
    await openTransaction(page, "Faux Bakery");
    await expect(
      page.getByRole("dialog").getByLabel("Note", { exact: true }),
    ).toHaveValue("Paid back by Pat");
  });

  test("edits name, paid with, and tax deductible", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Faux Bakery");
    const dialog = page.getByRole("dialog");
    // The name is the editable title; Paid with and Tax deductible are under Edit details
    await dialog.getByLabel("Name", { exact: true }).fill("=SUM(1)");
    await expect(dialog.getByText("Names cannot start with")).toBeVisible();
    await expect(
      dialog.getByRole("button", { name: "Save & close" }),
    ).toBeDisabled();

    await dialog.getByLabel("Name", { exact: true }).fill("Faux Bakery Two");
    await dialog.getByText("Edit details").click();
    await dialog
      .getByRole("group", { name: "Paid with" })
      .getByRole("button", { name: "Debit", exact: true })
      .click();
    await dialog.getByLabel("Tax deductible").check();
    await expect(dialog.getByTestId("dirty-indicator")).toBeVisible();
    await dialog.getByRole("button", { name: "Save & close" }).click();

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
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const item1 = dialog.getByRole("group", { name: "Split item 1" });
    await item1.getByLabel("Item note").fill("Pizza");
    await item1.getByLabel("Amount").fill("1+3.23*2");
    // Add tax starts checked for new items
    await expect(item1.getByLabel(/Add tax/)).toBeChecked();
    await item1.getByLabel(/Add tax/).uncheck();
    await item1.getByRole("button", { name: /^Dining Out/ }).click();
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const item2 = dialog.getByRole("group", { name: "Split item 2" });
    await item2.getByLabel("Item note").fill("Soda");
    await item2.getByLabel("Amount").fill("2");
    await item2.getByLabel(/Add tax/).uncheck();
    await item2.getByRole("button", { name: /^Groceries/ }).click();
    await expect(dialog.getByTestId("remaining")).toHaveText("$5.54");
    await dialog.getByRole("button", { name: "Save & close" }).click();

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
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const item1 = dialog.getByRole("group", { name: "Split item 1" });
    await item1.getByLabel("Item note").fill("Napkins");
    await item1.getByLabel("Amount").fill("4.23*2+12.09");
    // The original card has no tax toggle; the new item's is checked
    await expect(dialog.getByLabel(/Add tax/)).toHaveCount(1);
    await expect(item1.getByLabel(/Add tax/)).toBeChecked();
    await item1.getByRole("button", { name: /^Household/ }).click();
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const item2 = dialog.getByRole("group", { name: "Split item 2" });
    await item2.getByLabel("Item note").fill("Gift");
    await item2.getByLabel("Amount").fill("10");
    await item2.getByLabel(/Add tax/).uncheck();
    await item2.getByRole("button", { name: /^Gifts/ }).click();
    await dialog.getByRole("button", { name: "Save & close" }).click();

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

  test("renames a split row's item note in its card, keeping the merchant prefix", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | Paper towels");
    const dialog = page.getByRole("dialog");
    // The merchant prefix is the static title; only the notes are inputs
    await expect(dialog.getByTestId("name-prefix")).toHaveText("Fake Mart |");
    const towels = dialog.getByRole("group", { name: "Split item 1" });
    await expect(towels.getByLabel("Item note")).toHaveValue("Paper towels");
    await towels.getByLabel("Item note").fill("Kitchen towels");
    await dialog.getByRole("button", { name: "Save & close" }).click();

    await expect
      .poll(async () => (await backend.dump(sheetName))[1].values[COL.name])
      .toBe("Fake Mart | Kitchen towels");
    const rows = await backend.dump(sheetName);
    // Cost, envelope, and the rest of the group are untouched
    expect(rows[1].formulas[COL.cost]).toBe("=1.0825*(9.5)");
    expect(rows[1].values[COL.envelope]).toBe("Household");
    expect(rows[0].values[COL.name]).toBe("Fake Mart | ");
    expect(rows[2].values[COL.name]).toBe("Fake Mart | Birthday card");
  });

  test("renames the split parent note and an unsplit row from the title", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | ");
    await page
      .getByRole("dialog")
      .getByRole("group", { name: "Original" })
      .getByLabel("Item note")
      .fill("Cart");
    await page.getByRole("button", { name: "Save & close" }).click();
    await expect
      .poll(async () => (await backend.dump(sheetName))[0].values[COL.name])
      .toBe("Fake Mart | Cart");
    // Still one group: the cost formula is untouched
    expect((await backend.dump(sheetName))[0].formulas[COL.cost]).toBe(
      "=87.5-R[1]C[0]-R[2]C[0]",
    );

    await openTransaction(page, "Sample Cafe");
    await page
      .getByRole("dialog")
      .getByLabel("Name", { exact: true })
      .fill("Sample Bistro");
    await page.getByRole("button", { name: "Save & close" }).click();
    await expect
      .poll(async () => (await backend.dump(sheetName))[3].values[COL.name])
      .toBe("Sample Bistro");
  });

  test("edits an existing split group's child in the drawer", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | Birthday card");
    const dialog = page.getByRole("dialog");
    // The whole group opens: the original's amount is the remainder of the $87.50 total
    const original = dialog.getByRole("group", { name: "Original" });
    await expect(original.getByTestId("remaining")).toHaveText("$73.23");
    await expect(original.getByLabel(/Add tax/)).toHaveCount(0);
    const towels = dialog.getByRole("group", { name: "Split item 1" });
    await expect(towels.getByLabel("Amount")).toHaveValue("9.5");
    await expect(towels.getByLabel(/Add tax/)).toBeChecked();

    const card = dialog.getByRole("group", { name: "Split item 2" });
    await expect(card.getByLabel("Item note")).toHaveValue("Birthday card");
    await expect(card.getByLabel(/Add tax/)).not.toBeChecked();
    await card.getByLabel("Amount").fill("10");
    await card.getByLabel(/Add tax/).check();
    await expect(original.getByTestId("remaining")).toHaveText("$66.39");
    await card.getByText("Edit details").click();
    await card
      .getByRole("group", { name: "Paid with" })
      .getByRole("button", { name: "Debit", exact: true })
      .click();
    await card.getByLabel("Tax deductible").check();
    await dialog.getByRole("button", { name: "Save & close" }).click();

    await expect
      .poll(async () => (await backend.dump(sheetName))[2].formulas[COL.cost])
      .toBe("=1.0825*(10)");
    const rows = await backend.dump(sheetName);
    expect(rows[2].values[COL.paidWith]).toBe("Debit");
    expect(rows[2].values[COL.taxDeductible]).toBe("Yes");
    // The parent and the other child are untouched
    expect(rows[0].formulas[COL.cost]).toBe("=87.5-R[1]C[0]-R[2]C[0]");
    expect(rows[1].formulas[COL.cost]).toBe("=1.0825*(9.5)");
    await expect(syncedPill(page)).toBeVisible();
  });

  test("splits with the original's note and envelope edited", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Jane Example - Lunch money");
    const dialog = page.getByRole("dialog");
    await dialog.getByRole("button", { name: "Add split item" }).click();

    const original = dialog.getByRole("group", { name: "Original" });
    await expect(original.getByLabel("Cost")).toHaveCount(0);
    await original.getByLabel("Item note").fill("Mine");
    await original.getByRole("button", { name: /^Dining Out/ }).click();
    // Staged, not the one-tap save
    await expect(dialog).toBeVisible();
    const item1 = dialog.getByRole("group", { name: "Split item 1" });
    await item1.getByLabel("Item note").fill("Theirs");
    await item1.getByLabel("Amount").fill("5");
    await item1.getByLabel(/Add tax/).uncheck();
    await item1.getByRole("button", { name: /^Groceries/ }).click();
    await expect(dialog.getByTestId("remaining")).toHaveText("$10.00");
    await dialog.getByRole("button", { name: "Save & close" }).click();

    await expect
      .poll(async () => (await backend.dump(sheetName))[11]?.values[COL.name])
      .toBe("Jane Example - Lunch money | Theirs");
    const rows = await backend.dump(sheetName);
    const parent = rows[10];
    expect(parent.values[COL.name]).toBe("Jane Example - Lunch money | Mine");
    expect(parent.formulas[COL.cost]).toBe("=15-R[1]C[0]");
    expect(parent.values[COL.envelope]).toBe("Dining Out");
    expect(rows[11].values[COL.envelope]).toBe("Groceries");
    expect(rows[11].values[COL.cost]).toBe(5);
    await expect(syncedPill(page)).toBeVisible();
  });

  test("dismissing a split with changes asks Save / Discard / Go back", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await openTransaction(page, "Jane Example - Lunch money");
    const dialog = page.getByRole("dialog", { name: "Transaction" });
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const amount = dialog
      .getByRole("group", { name: "Split item 1" })
      .getByLabel("Amount");
    await amount.fill("5");
    await dialog.getByRole("button", { name: "Close", exact: true }).click();
    const prompt = page.getByRole("alertdialog", { name: "Unsaved changes" });
    await prompt.getByRole("button", { name: "Go back" }).click();
    await expect(amount).toHaveValue("5");
    await dialog.getByRole("button", { name: "Close", exact: true }).click();
    await prompt.getByRole("button", { name: "Discard" }).click();
    await expect(page.getByRole("dialog")).toHaveCount(0);
    expect((await backend.dump(sheetName))[10].values[COL.name]).toBe(
      "Jane Example - Lunch money",
    );
  });
});

test.describe("unsaved changes and save status", () => {
  test("dismissing with changes asks Save / Discard / Go back; no prompt without changes", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    // Nothing changed: closing is immediate
    await openTransaction(page, "Faux Bakery");
    await page
      .getByRole("dialog")
      .getByRole("button", { name: "Close", exact: true })
      .click();
    await expect(page.getByRole("dialog")).toHaveCount(0);

    await openTransaction(page, "Faux Bakery");
    const dialog = page.getByRole("dialog", { name: "Transaction" });
    const prompt = page.getByRole("alertdialog", { name: "Unsaved changes" });
    const name = dialog.getByLabel("Name", { exact: true });
    await name.fill("Faux Bakery Edited");
    await expect(dialog.getByTestId("dirty-indicator")).toBeVisible();

    // Go back keeps editing (x button)
    await dialog.getByRole("button", { name: "Close", exact: true }).click();
    await prompt.getByRole("button", { name: "Go back" }).click();
    await expect(prompt).toHaveCount(0);
    await expect(name).toHaveValue("Faux Bakery Edited");

    // Escape opens the prompt too; Discard drops the edit
    await page.keyboard.press("Escape");
    await expect(prompt).toBeVisible();
    await prompt.getByRole("button", { name: "Discard" }).click();
    await expect(page.getByRole("dialog")).toHaveCount(0);
    expect(findRow(await backend.dump(sheetName), "Faux Bakery")).toBeTruthy();

    // Tapping the backdrop with changes prompts too
    await openTransaction(page, "Faux Bakery");
    await name.fill("Nope");
    await page.mouse.click(195, 20);
    await expect(prompt).toBeVisible();
    await prompt.getByRole("button", { name: "Go back" }).click();

    // Swipe down (drag the handle) opens the prompt; Save writes the change
    await name.fill("Faux Bakery Edited");
    const handle = dialog.getByTestId("drag-handle");
    const box = (await handle.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, box.y + 200, { steps: 5 });
    await page.mouse.up();
    await expect(prompt).toBeVisible();
    await prompt.getByRole("button", { name: "Save" }).click();
    await expect(page.getByRole("dialog")).toHaveCount(0);
    await expect
      .poll(async () =>
        (await backend.dump(sheetName)).some(
          (r) => r.values[COL.name] === "Faux Bakery Edited",
        ),
      )
      .toBe(true);
    // The discarded edits never reached the sheet
    expect(
      (await backend.dump(sheetName)).some(
        (r) => r.values[COL.name] === "Nope",
      ),
    ).toBe(false);
  });

  test("each row shows Saving…, then Saved ✓ once the sheet has the change", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    // Slow the write down so the pending state is observable
    await page.route("**/macros/s/demo/exec", async (route) => {
      if (route.request().postData()?.includes('"updateTransaction"'))
        await new Promise((r) => setTimeout(r, 800));
      await route.continue();
    });
    await openTransaction(page, "Test Hardware");
    await page
      .getByRole("dialog")
      .getByRole("button", { name: /^Household/ })
      .click();
    const status = page
      .getByRole("button", { name: /Test Hardware/ })
      .getByTestId("row-status");
    await expect(status).toHaveText("Saving…");
    await expect(status).toHaveText("Saved ✓");
    expect(
      findRow(await backend.dump(sheetName), "Test Hardware").values[
        COL.envelope
      ],
    ).toBe("Household");
    // It clears itself a few seconds later
    await expect(status).toHaveCount(0, { timeout: 8000 });
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
    await expect(
      page
        .getByRole("button", { name: /Imaginary Books/ })
        .getByTestId("row-status"),
    ).toHaveText("Queued – offline");
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

  test("a forwarded non-receipt from an unknown sender, with a reply, is ignored without an error", async ({
    page,
    sheetName,
  }) => {
    await backend.addThreads([
      { messages: receiptThreadFixtures.forwardedFromUnknownSender },
    ]);
    await connect(page);
    const before = (await backend.dump(sheetName)).length;
    await page.getByRole("button", { name: "Sync receipts" }).click();
    await expect(page.getByText("Demo Cafe").first()).toBeVisible();

    expect((await backend.dump(sheetName)).length).toBe(before + 4);
    const thread = (await backend.threads()).find(
      (t) => t.subject === "Fwd: Dinner Friday?",
    );
    expect(thread?.labels).toEqual(["Receipts/Scripted", "Receipts/Ignored"]);
    expect(thread?.archived).toBe(false);
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
    await expect(dialog.getByRole("group", { name: "Paid with" })).toHaveCount(
      0,
    );
    await dialog.getByRole("button", { name: "Close", exact: true }).click();

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
    // The old separate scope selector is gone; "All" lives in the pay-period dropdown
    await expect(page.getByLabel("Transactions shown")).toHaveCount(0);
    // Not loaded until chosen
    await expect(page.getByText("Faux Bakery")).toHaveCount(1); // current tab only
    const all = (await callApi({ action: "getAllTransactions" })).data;
    const names: string[] = all.sheets.map((s: { name: string }) => s.name);
    expect(names).toHaveLength(3);
    const previous = names[1];

    await page.getByLabel("Pay period").selectOption({ label: "All" });
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
    await page.getByLabel("Pay period").selectOption({ label: "All" });
    await page.getByLabel("Search").fill("Faux Bakery");
    const dialog = page.getByRole("dialog");
    await page
      .getByRole("heading", { name: previous })
      .locator("xpath=following-sibling::div")
      .getByRole("button", { name: /Faux Bakery/ })
      .click();
    await dialog.getByRole("button", { name: "Add split item" }).click();
    const item = dialog.getByRole("group", { name: "Split item 1" });
    await item.getByLabel("Item note").fill("Rolls");
    await item.getByLabel("Amount").fill("2.5");
    await item.getByLabel(/Add tax/).uncheck();
    await expect(dialog.getByTestId("remaining")).toHaveText("$3.75");
    await dialog.getByRole("button", { name: "Save & close" }).click();
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

test.describe("deleting", () => {
  test("deletes a single row after confirming; the row is blank and rows below (with a split group) do not move", async ({
    page,
    sheetName,
  }) => {
    // A split group below the row to delete, whose relative formula must stay untouched
    const before = await backend.dump(sheetName);
    const hardware = findRow(before, "Test Hardware");
    const split = await callApi({
      action: "splitTransaction",
      sheetName,
      rowIndex: hardware.rowIndex,
      baseFingerprint: (await callApi({ action: "getState", sheetName })).data
        .transactions[hardware.rowIndex].fingerprint,
      splits: [
        { itemNote: "Nails", expression: "5", preTax: false, envelope: "" },
      ],
    });
    expect(split.ok).toBe(true);
    const withSplit = await backend.dump(sheetName);
    const cafe = findRow(withSplit, "Sample Cafe");
    const cafeBefore = withSplit[cafe.rowIndex];
    const usedBefore = withSplit.length;

    await connect(page);
    await openTransaction(page, "Sample Cafe");
    const drawer = page.getByRole("dialog");
    await drawer.getByRole("button", { name: "Delete transaction" }).click();
    const confirm = page.getByRole("alertdialog", {
      name: "Delete transaction?",
    });
    await expect(confirm).toContainText("Sample Cafe · $12.35");
    await expect(confirm.getByRole("button", { name: "Cancel" })).toBeFocused();
    // Escape and Cancel only close the confirmation; nothing is deleted
    await page.keyboard.press("Escape");
    await expect(confirm).toHaveCount(0);
    await expect(drawer).toBeVisible();
    await drawer.getByRole("button", { name: "Delete transaction" }).click();
    await confirm.getByRole("button", { name: "Cancel" }).click();
    await expect(confirm).toHaveCount(0);
    expect(findRow(await backend.dump(sheetName), "Sample Cafe").rowIndex).toBe(
      cafe.rowIndex,
    );

    await drawer.getByRole("button", { name: "Delete transaction" }).click();
    await confirm.getByRole("button", { name: "Delete" }).click();
    await expect(drawer).toHaveCount(0);
    await expect
      .poll(
        async () =>
          (await backend.dump(sheetName))[cafe.rowIndex].values[COL.name],
      )
      .toBe("");
    const rows = await backend.dump(sheetName);
    expect(rows).toHaveLength(usedBefore);
    // The deleted row is blank in cols A-C and E-G; its Current Balance formula stays
    const blank = rows[cafe.rowIndex];
    expect(blank.values.filter((_, c) => c !== 3).every((v) => v === "")).toBe(
      true,
    );
    expect(blank.formulas[3]).toBe(cafeBefore.formulas[3]);
    // Every other row is unchanged (formulas, notes, backgrounds), none moved
    const stable = (r: (typeof rows)[number]) => [
      r.formulas,
      r.notes,
      r.backgrounds,
    ];
    rows.forEach((r, i) => {
      if (i !== cafe.rowIndex) expect(stable(r)).toEqual(stable(withSplit[i]));
    });
    // Only this Sample Cafe was deleted (a later one, on another day, stays)
    expect(
      rows.filter((r) => r.values[COL.name] === "Sample Cafe"),
    ).toHaveLength(1);
    // The split group did not move; its relative formula still subtracts its child
    const parent = findRow(rows, "Test Hardware | ");
    expect(parent.rowIndex).toBe(hardware.rowIndex);
    expect(parent.formulas[COL.cost]).toBe("=27.83-R[1]C[0]");
    expect(parent.values[COL.cost]).toBeCloseTo(22.83);
    expect(rows[parent.rowIndex + 1].values[COL.name]).toBe(
      "Test Hardware | Nails",
    );
    await expect(syncedPill(page)).toBeVisible();
    // It was the only uncategorized Sample Cafe
    await page.getByLabel("Search").fill("Sample Cafe");
    await expect(page.getByRole("button", { name: /Sample Cafe/ })).toHaveCount(
      0,
    );
  });

  test("deletes a whole split group from one of its rows", async ({
    page,
    sheetName,
  }) => {
    const before = await backend.dump(sheetName);
    const usedBefore = before.length;
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | Birthday card");
    const drawer = page.getByRole("dialog");
    await drawer.getByRole("button", { name: "Delete split" }).click();
    const confirm = page.getByRole("alertdialog", { name: "Delete split?" });
    await expect(confirm).toContainText("Fake Mart · $87.50");
    await expect(confirm).toContainText("all 3 rows of this split");
    await confirm.getByRole("button", { name: "Delete" }).click();
    await expect(drawer).toHaveCount(0);

    const group = before.filter((r) =>
      String(r.values[COL.name]).startsWith("Fake Mart"),
    );
    expect(group).toHaveLength(3);
    await expect
      .poll(
        async () =>
          (await backend.dump(sheetName))[group[0].rowIndex].values[COL.name],
      )
      .toBe("");
    const rows = await backend.dump(sheetName);
    // Rows are cleared in place: the table keeps its length and no other row moves
    expect(rows).toHaveLength(usedBefore);
    expect(
      rows.some((r) => String(r.values[COL.name]).startsWith("Fake Mart")),
    ).toBe(false);
    group.forEach((g) => {
      const r = rows[g.rowIndex];
      expect(r.values.filter((_, c) => c !== 3).every((v) => v === "")).toBe(
        true,
      );
      expect(r.formulas[3]).toBe(g.formulas[3]);
    });
    const stable = (r: (typeof rows)[number]) => [
      r.formulas,
      r.notes,
      r.backgrounds,
    ];
    rows.forEach((r, i) => {
      if (!group.some((g) => g.rowIndex === i))
        expect(stable(r)).toEqual(stable(before[i]));
    });
    await expect(syncedPill(page)).toBeVisible();
    await page.getByLabel("Search").fill("Fake Mart");
    await expect(page.getByRole("button", { name: /Fake Mart/ })).toHaveCount(
      0,
    );
  });
});
