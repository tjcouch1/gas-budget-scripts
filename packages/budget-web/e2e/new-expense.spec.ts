import {
  backend,
  callApi,
  COL,
  connect,
  expect,
  findRow,
  syncedPill,
  test,
} from "./helpers.ts";

test.describe("new expense", () => {
  test("adds an expense from the button at the end of the list", async ({
    page,
    sheetName,
  }) => {
    const before = await backend.dump(sheetName);
    const lastFilled = before.reduce(
      (last, r) =>
        r.values.some((v, c) => c !== COL.balance && v !== "" && v !== null)
          ? r.rowIndex
          : last,
      -1,
    );
    const { data: state } = await callApi({ action: "getState", sheetName });
    const period = state.sheets.find(
      (s: { name: string }) => s.name === sheetName,
    );

    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    const button = page.getByRole("button", { name: "New expense" });
    // At the bottom of the list: the user scrolls down to reach it
    await button.scrollIntoViewIfNeeded();
    await button.click();

    const dialog = page.getByRole("dialog", { name: "New expense" });
    await dialog.getByLabel("Date").fill(period.startDate);
    await dialog.getByLabel("Name").fill("E2E Farmers Market");
    await dialog.getByLabel("Cost").fill("4.25*2");
    await expect(dialog.getByText("$8.50")).toBeVisible();
    await dialog
      .getByRole("group", { name: "Envelope" })
      .getByRole("button", { name: "Groceries", exact: true })
      .click();
    await dialog
      .getByRole("group", { name: "Paid with" })
      .getByRole("button", { name: "Cash", exact: true })
      .click();
    await dialog.getByLabel("Note (optional)").fill("from the market");
    await dialog.getByRole("button", { name: "Save" }).click();
    await expect(dialog).toHaveCount(0);

    await expect(syncedPill(page)).toBeVisible();
    await expect
      .poll(async () =>
        (await backend.dump(sheetName)).some(
          (r) => r.values[COL.name] === "E2E Farmers Market",
        ),
      )
      .toBe(true);
    const row = findRow(await backend.dump(sheetName), "E2E Farmers Market");
    // Written right after the last non-empty row
    expect(row.rowIndex).toBe(lastFilled + 1);
    expect(row.formulas[COL.cost]).toBe("=4.25*2");
    expect(row.values[COL.envelope]).toBe("Groceries");
    expect(row.values[COL.paidWith]).toBe("Cash");
    expect(row.notes[COL.name]).toBe("from the market");

    // It is listed, confirmed by the sheet
    await page.getByLabel("Search").fill("E2E Farmers");
    const listed = page.getByRole("button", { name: /E2E Farmers Market/ });
    await expect(listed).toBeVisible();
    await expect(listed).toContainText("Groceries");
    await expect(listed).toContainText("Cash");
  });
});
