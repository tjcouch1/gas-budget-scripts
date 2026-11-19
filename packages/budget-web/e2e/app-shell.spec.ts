import {
  API_KEY,
  backend,
  connect,
  expect,
  findRow,
  GAS_URL,
  syncedPill,
  test,
} from "./helpers.ts";

test.describe("app shell", () => {
  test("first-run settings closes itself after a working save", async ({
    page,
  }) => {
    // Fake timers so the 2-second auto-close can be skipped instead of waited for
    await page.clock.install();
    await page.goto("/");
    const dialog = page.getByRole("dialog", { name: "Connection settings" });
    await dialog.getByLabel("Web app URL").fill(GAS_URL);
    await dialog.getByLabel("API key").fill(API_KEY);
    await dialog.getByRole("button", { name: "Save" }).click();
    await expect(dialog.getByRole("status")).toContainText("Connected.");
    await page.clock.fastForward(2000);
    await expect(dialog).toBeHidden();
    await expect(page.getByLabel("Pay period")).toBeVisible();
  });

  test("Settings opened manually does not close by itself", async ({
    page,
  }) => {
    await page.clock.install();
    await connect(page);
    await page.getByRole("button", { name: "Settings" }).click();
    const dialog = page.getByRole("dialog", { name: "Connection settings" });
    await dialog.getByLabel("Web app URL").fill(GAS_URL);
    await dialog.getByLabel("API key").fill(API_KEY);
    await dialog.getByRole("button", { name: "Save" }).click();
    await expect(dialog.getByRole("status")).toContainText("Saved");
    await expect(syncedPill(page)).toBeVisible();
    // Well past the 2-second auto-close a first-run save would use
    await page.clock.fastForward(5000);
    await expect(dialog).toBeVisible();
  });

  test("Sync pulls changes made in the sheet", async ({ page, sheetName }) => {
    await connect(page);
    const rows = await backend.dump(sheetName);
    const row = findRow(rows, "Sample Cafe");
    await expect(page.getByText("Sample Cafe").first()).toBeVisible();
    await backend.mutate(sheetName, row.rowIndex, {
      name: "Renamed In Sheet",
    });
    await page.getByRole("button", { name: "Sync now" }).click();
    await expect(page.getByText("Renamed In Sheet").first()).toBeVisible();
    await expect(syncedPill(page)).toBeVisible();
  });
});

test.describe("testing mode", () => {
  test("imports receipts without marking them in Gmail", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    const before = (await backend.dump(sheetName)).length;
    await page.getByRole("button", { name: "Settings" }).click();
    const dialog = page.getByRole("dialog", { name: "Connection settings" });
    await dialog
      .getByRole("checkbox", { name: /Testing mode: import receipts/ })
      .check();
    await dialog.getByRole("button", { name: "Close" }).click();
    const button = page.getByRole("button", { name: "Sync receipts" });
    await expect(button).toHaveText("Import receipts (no mark)");
    await button.click();
    await expect(page.getByText("Demo Cafe").first()).toBeVisible();

    expect((await backend.dump(sheetName)).length).toBe(before + 4);
    const threads = await backend.threads();
    expect(threads.filter((t) => t.archived)).toHaveLength(0);
    expect(threads.every((t) => t.labels.length === 0)).toBe(true);
  });
});
