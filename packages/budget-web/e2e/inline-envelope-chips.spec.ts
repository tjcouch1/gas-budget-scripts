import {
  backend,
  connect,
  expect,
  findRow,
  COL,
  syncedPill,
  test,
} from "./helpers.ts";

test.describe("inline envelope chips setting", () => {
  test("uncategorized rows show chips under the row; choosing saves and the chips stay", async ({
    page,
    sheetName,
  }) => {
    await page.addInitScript(() =>
      localStorage.setItem("budget.inlineEnvelopeChips", "1"),
    );
    await connect(page);
    // No dropdown in this mode
    await expect(page.getByLabel("Envelope for Test Hardware")).toHaveCount(0);
    const row = page.locator("[data-row-key]", { hasText: "Test Hardware" });
    const group = row.getByRole("group", { name: "Envelope" });
    await expect(group).toBeVisible();
    // No title or expand button
    await expect(group.getByRole("button", { expanded: false })).toHaveCount(0);

    await group.getByRole("button", { name: /^Household/ }).click();
    // Chosen without opening the transaction
    await expect(page.getByRole("dialog")).toHaveCount(0);
    await expect
      .poll(async () => {
        const saved = findRow(await backend.dump(sheetName), "Test Hardware");
        return saved.values[COL.envelope];
      })
      .toBe("Household");
    await expect(syncedPill(page)).toBeVisible();
    // The chooser stays on the row, showing the selection
    await expect(
      row.getByRole("group", { name: "Envelope" }).getByRole("button", {
        name: /^Household/,
      }),
    ).toHaveAttribute("aria-pressed", "true");
  });

  test("the Settings toggle switches the dropdown to chips", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await connect(page);
    await expect(page.getByLabel("Envelope for Test Hardware")).toBeVisible();
    await page.getByRole("button", { name: "Settings" }).click();
    await page
      .getByLabel(/Choose envelopes with chips instead of the dropdown/)
      .check();
    await page.getByRole("button", { name: "Close", exact: true }).click();
    await expect(page.getByLabel("Envelope for Test Hardware")).toHaveCount(0);
  });
});
