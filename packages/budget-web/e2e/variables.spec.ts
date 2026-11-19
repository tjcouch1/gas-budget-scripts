import type { Page } from "@playwright/test";
import { backend, connect, COL, expect, test } from "./helpers.ts";

/** Opens Settings, then the Customize variables dialog, and fetches the demo inbox */
async function openVariables(page: Page) {
  await page.getByRole("button", { name: "Settings" }).click();
  await page.getByRole("button", { name: "Customize variables" }).click();
  const dialog = page.getByRole("dialog", { name: "Customize variables" });
  await expect(dialog).toBeVisible();
  return dialog;
}

type Dialog = ReturnType<Page["getByRole"]>;

/** Switches tabs (the dialog shows one panel at a time at the test viewport width) */
async function showTab(dialog: Dialog, name: RegExp | string) {
  await dialog.getByRole("tab", { name }).click();
}

/** Fetches on the Emails tab, then shows the Results tab */
async function fetchEmails(dialog: Dialog) {
  await showTab(dialog, "Emails");
  await dialog.getByRole("button", { name: "Fetch emails" }).click();
  await expect(dialog.getByTestId("server-query")).toBeVisible();
  await showTab(dialog, /^Results/);
}

test.describe("Customize variables", () => {
  test("fetches the inbox and previews which emails are receipts", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await connect(page);
    const dialog = await openVariables(page);
    await fetchEmails(dialog);

    const cafe = dialog.getByRole("article", { name: /Demo Cafe/ });
    await expect(cafe.getByTestId("outcome-badge")).toContainText("Receipt");
    await expect(cafe.getByTestId("receipt-name")).toHaveText("Demo Cafe");
    await expect(cafe.getByTestId("receipt-cost")).toContainText("18.42");
    await expect(
      dialog
        .getByRole("article", { name: "Your statement is ready" })
        .getByTestId("outcome-badge"),
    ).toContainText("Not a receipt");
    await showTab(dialog, "Emails");
    await expect(dialog.getByTestId("server-query")).toContainText("from:(");
  });

  test("editing a rule changes the live preview", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await connect(page);
    const dialog = await openVariables(page);
    await fetchEmails(dialog);
    await expect(
      dialog.getByRole("article", { name: /Demo Cafe/ }),
    ).toBeVisible();

    await showTab(dialog, "Rules");
    await dialog
      .getByRole("group", { name: "Rule chase-purchase" })
      .getByLabel("Name", { exact: true })
      .fill("${subject.name}!");
    await showTab(dialog, /^Results/);
    const cafe = dialog.getByRole("article", { name: /Demo Cafe/ });
    await expect(cafe.getByTestId("receipt-name")).toHaveText("Demo Cafe!");
  });

  test("an invalid regex shows a problem and disables Save", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await connect(page);
    const dialog = await openVariables(page);
    await dialog
      .getByRole("group", { name: "Rule chase-purchase" })
      .getByRole("textbox", { name: "Subject pattern" })
      .fill("(");
    await expect(
      dialog.getByRole("group", { name: "Rule problems" }),
    ).toBeVisible();
    await expect(
      dialog.getByRole("button", { name: "Save to sheet" }),
    ).toBeDisabled();
  });

  test("saving writes the variable and the next import uses the saved rules", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    const dialog = await openVariables(page);
    await dialog
      .getByRole("group", { name: "Rule chase-purchase" })
      .getByLabel("Name", { exact: true })
      .fill("${subject.name}!");
    await dialog.getByRole("button", { name: "Save to sheet" }).click();
    await expect(dialog.getByRole("status")).toContainText(
      "Saved to the Variables sheet",
    );

    const variables = await backend.variables();
    const saved = variables.find(([key]) => key === "ReceiptRules");
    expect(String(saved?.[1])).toContain("${subject.name}!");

    await dialog.getByRole("button", { name: "Close", exact: true }).click();
    await page.getByRole("button", { name: "Sync receipts" }).click();
    await expect(page.getByText("Demo Cafe!").first()).toBeVisible();
    const names = (await backend.dump(sheetName)).map((r) =>
      String(r.values[COL.name]),
    );
    expect(names).toContain("Demo Cafe!");
  });

  test("Max emails limits the fetch and Include processed finds archived mail", async ({
    page,
    sheetName,
  }) => {
    await connect(page);
    await page.getByRole("button", { name: "Sync receipts" }).click();
    await expect(page.getByText("Demo Cafe").first()).toBeVisible();
    expect((await backend.dump(sheetName)).length).toBeGreaterThan(0);

    const dialog = await openVariables(page);
    await fetchEmails(dialog);
    // Processed emails are excluded by default
    await expect(dialog.getByRole("article")).toHaveCount(0);

    await showTab(dialog, "Emails");
    await dialog
      .getByRole("checkbox", { name: /Include already-processed emails/ })
      .check();
    await fetchEmails(dialog);
    await expect(dialog.getByRole("article")).toHaveCount(5);

    await showTab(dialog, "Emails");
    await dialog.getByLabel("Max emails").fill("2");
    await fetchEmails(dialog);
    await expect(dialog.getByRole("article")).toHaveCount(2);
  });

  test("an invalid ReceiptRules cell shows the notice with an action", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await connect(page);
    await backend.setVariable("ReceiptRules", "{ not json");
    await page.getByRole("button", { name: "Sync receipts" }).click();
    const notice = page.getByRole("alert");
    await expect(notice).toContainText("ReceiptRules");
    await notice
      .getByRole("button", { name: "Open Customize variables" })
      .click();
    await expect(
      page.getByRole("dialog", { name: "Customize variables" }),
    ).toBeVisible();
  });

  test("email content is shown as text, never as HTML", async ({
    page,
    sheetName,
  }) => {
    expect(sheetName).toBeTruthy();
    await backend.addThreads([
      {
        from: "Chase <no.reply.alerts@chase.com>",
        subject: "You made a $5.00 transaction with <b>Bold</b> Shop",
        body: "<script>window.__pwned = 1</script>",
        html: '<p><img src=x onerror="window.__pwned = 1">hi</p>',
      },
    ]);
    await connect(page);
    const dialog = await openVariables(page);
    await showTab(dialog, "Emails");
    await dialog.getByLabel("Include HTML bodies").check();
    await fetchEmails(dialog);

    const article = dialog.getByRole("article", { name: /Bold/ });
    await expect(article.getByTestId("receipt-name")).toHaveText(
      "<b>Bold</b> Shop",
    );
    await expect(
      dialog.locator("article b, article script, article img"),
    ).toHaveCount(0);
    expect(
      await page.evaluate(
        () => (window as unknown as { __pwned?: number }).__pwned,
      ),
    ).toBeUndefined();
  });
});
