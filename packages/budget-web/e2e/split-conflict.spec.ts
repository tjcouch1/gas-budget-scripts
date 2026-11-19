import type { Page } from "@playwright/test";
import {
  backend,
  COL,
  connect,
  expect,
  openTransaction,
  syncedPill,
  test,
} from "./helpers.ts";

/**
 * A split group is saved in one `updateSplitGroup` request: when any row of the group changed in
 * the sheet, nothing is written and the whole group is resolved in one 3-way dialog.
 */
test.describe("split group conflicts", () => {
  /** Edits the Fake Mart split in the drawer while someone edits one of its rows in the sheet. */
  async function editSplitWhileSheetChanges(page: Page, sheetName: string) {
    await connect(page);
    await page.getByRole("tab", { name: /^All/ }).click();
    await openTransaction(page, "Fake Mart | Birthday card");
    // Someone moves the paper towels to another envelope after the app loaded the split
    await backend.mutate(sheetName, 1, { envelope: "Groceries" });
    const drawer = page.getByRole("dialog");
    await drawer.getByLabel("Name", { exact: true }).fill("Fake Market");
    await drawer
      .getByRole("group", { name: "Split item 2" })
      .getByLabel("Amount")
      .fill("5");
    await drawer.getByRole("button", { name: "Save & close" }).click();

    const conflict = page.getByRole("dialog", {
      name: "This split changed in the sheet",
    });
    await expect(conflict).toBeVisible();
    // Nothing of the save was written
    const rows = await backend.dump(sheetName);
    expect(rows.slice(0, 3).map((r) => r.values[COL.name])).toEqual([
      "Fake Mart | ",
      "Fake Mart | Paper towels",
      "Fake Mart | Birthday card",
    ]);
    expect(rows[2].values[COL.cost]).toBe(3.99);

    // One 3-way table per row of the split
    const towels = conflict.getByRole("region", { name: "Split item 1" });
    await expect(
      towels.getByRole("row", { name: /Current remote/ }),
    ).toContainText("Groceries");
    await expect(towels.getByRole("row", { name: /Original/ })).toContainText(
      "Household",
    );
    const card = conflict.getByRole("region", { name: "Split item 2" });
    await expect(
      card.getByRole("row", { name: /Discarded local change/ }),
    ).toContainText("Fake Market | Birthday card");
    await expect(
      card.getByRole("row", { name: /Current remote/ }),
    ).toContainText("Fake Mart | Birthday card");
    return conflict;
  }

  test("Accept Remote keeps the whole split as the sheet has it", async ({
    page,
    sheetName,
  }) => {
    const conflict = await editSplitWhileSheetChanges(page, sheetName);
    await conflict.getByRole("button", { name: "Accept Remote" }).click();
    await expect(conflict).toHaveCount(0);
    await expect(syncedPill(page)).toBeVisible();
    const rows = await backend.dump(sheetName);
    expect(rows.slice(0, 3).map((r) => r.values[COL.name])).toEqual([
      "Fake Mart | ",
      "Fake Mart | Paper towels",
      "Fake Mart | Birthday card",
    ]);
    expect(rows[1].values[COL.envelope]).toBe("Groceries");
    expect(rows[2].values[COL.cost]).toBe(3.99);
  });

  test("Overwrite with Mine saves the whole split again, keeping the sheet's other edit", async ({
    page,
    sheetName,
  }) => {
    const conflict = await editSplitWhileSheetChanges(page, sheetName);
    await conflict.getByRole("button", { name: "Overwrite with Mine" }).click();
    await expect(conflict).toHaveCount(0);
    await expect
      .poll(async () => (await backend.dump(sheetName))[0].values[COL.name])
      .toBe("Fake Market | ");
    const rows = await backend.dump(sheetName);
    expect(rows.slice(1, 3).map((r) => r.values[COL.name])).toEqual([
      "Fake Market | Paper towels",
      "Fake Market | Birthday card",
    ]);
    expect(rows[2].values[COL.cost]).toBe(5);
    // The envelope edited in the sheet was not part of the save, so it stays
    expect(rows[1].values[COL.envelope]).toBe("Groceries");
    await expect(syncedPill(page)).toBeVisible();
  });
});
