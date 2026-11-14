import { describe, expect, it } from "vitest";
import { getEnv } from "./gas-mocks";

const load = () => import("../dev/backend.ts");

type Json = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

async function create() {
  const { DemoBackend, DEMO_API_KEY } = await load();
  const backend = new DemoBackend();
  const call = (request: Json): Json =>
    JSON.parse(
      backend.post(
        JSON.stringify({ apiVersion: 1, apiKey: DEMO_API_KEY, ...request }),
      ),
    );
  return { backend, call };
}

describe("mock formula evaluation (opt-in)", () => {
  it("keeps formula text as the value by default", () => {
    const sheet = getEnv().spreadsheet.addSheet("S");
    sheet.seed("A1", [[2, "=R[0]C[-1]*3"]]);
    expect(sheet.getRange("B1").getValue()).toBe("=R[0]C[-1]*3");
  });

  it("evaluates arithmetic and R1C1 references when enabled", () => {
    const env = getEnv();
    env.spreadsheet.evaluateFormulas = true;
    const sheet = env.spreadsheet.addSheet("S");
    sheet.seed("A1", [[2, "=RC[-1]*3", "=R[1]C[-2]-R[2]C[-2]+1.5*(2)"]]);
    sheet.seed("A2", [[10]]);
    sheet.seed("A3", [[4]]);
    expect(sheet.getRange("B1").getValue()).toBe(6);
    expect(sheet.getRange("C1").getValue()).toBe(9);
    expect(sheet.getRange("B1").getDisplayValues()).toEqual([["6.00"]]);
    // The formula text is still what getFormulas returns
    expect(sheet.getRange("B1").getFormulas()).toEqual([["=RC[-1]*3"]]);
  });

  it("leaves unevaluable formulas as text", () => {
    const env = getEnv();
    env.spreadsheet.evaluateFormulas = true;
    const sheet = env.spreadsheet.addSheet("S");
    sheet.seed("A1", [["text", "=R[0]C[-1]*2"]]);
    expect(sheet.getRange("B1").getValue()).toBe("=R[0]C[-1]*2");
  });
});

describe("demo backend", () => {
  it("serves the seeded pay periods; the oldest lacks the optional keys", async () => {
    const { call } = await create();
    const state = call({ action: "getState" });
    expect(state.ok).toBe(true);
    expect(state.data.sheets).toHaveLength(3);
    expect(state.data.taxMultiplier).toBe(1.0825);
    expect(state.data.envelopes.length).toBeGreaterThan(5);
    expect(state.data.paidWithOptions).toContain("Credit");
    expect(
      state.data.transactions.filter((t: Json | null) => t && t.envelope === "")
        .length,
    ).toBeGreaterThan(5);

    const oldest = call({
      action: "getState",
      sheetName: state.data.sheets[2].name,
    });
    expect(oldest.ok).toBe(true);
    expect(oldest.data.envelopes).toBeNull();
    expect(oldest.data.paidWithOptions).toBeNull();
    expect(oldest.warnings.length).toBeGreaterThan(0);
  });

  it("evaluates the existing split group's formulas", async () => {
    const { call } = await create();
    const [parent, child] = call({ action: "getState" }).data.transactions;
    expect(parent.costFormula).toBe("=87.5-R[1]C[0]-R[2]C[0]");
    expect(child.cost).toBeCloseTo(1.0825 * 9.5);
    expect(parent.cost).toBeCloseTo(87.5 - 1.0825 * 9.5 - 3.99);
    expect(parent.splitGroupId).not.toBeNull();
  });

  it("rejects a wrong API key", async () => {
    const { backend } = await create();
    const response = JSON.parse(
      backend.post(
        JSON.stringify({ apiVersion: 1, apiKey: "nope", action: "getState" }),
      ),
    );
    expect(response.ok).toBe(false);
    expect(response.error.code).toBe("unauthorized");
  });

  it("imports the fake receipts once and labels them", async () => {
    const { call, backend } = await create();
    const before = call({ action: "getState" }).data.transactions.length;
    const first = call({ action: "syncReceipts" });
    expect(first.ok).toBe(true);
    expect(first.data.receiptInfosAdded).toBe(4);
    expect(first.data.numErrors).toBe(0);
    expect(first.data.state.transactions.length).toBe(before + 4);

    const second = call({ action: "syncReceipts" });
    expect(second.data.receiptInfosAdded).toBe(0);
    const threads = backend.dumpThreads();
    expect(
      threads.filter((t) => t.archived && t.labels.includes("Receipts")),
    ).toHaveLength(4);
    expect(
      threads.find((t) => t.subject === "Your statement is ready")?.labels,
    ).toContain("Receipts/Ignored");
  });

  it("reset restores the seed", async () => {
    const { call, backend } = await create();
    call({ action: "syncReceipts" });
    backend.reset();
    expect(call({ action: "syncReceipts" }).data.receiptInfosAdded).toBe(4);
  });

  it("debug helpers mutate and dump rows", async () => {
    const { call, backend } = await create();
    const sheetName = call({ action: "getState" }).data.sheetName;
    backend.mutateRow({
      sheetName,
      rowIndex: 3,
      changes: { name: "Edited elsewhere", cost: "=2+3" },
    });
    const dump = backend.dumpSheet(sheetName);
    expect(dump.rows[3].values[1]).toBe("Edited elsewhere");
    expect(dump.rows[3].formulas[2]).toBe("=2+3");
    expect(dump.rows[3].values[2]).toBe(5);
  });

  it("Gmail search filters by the query's from, subject and in:inbox", async () => {
    const { call } = await create();
    const fetch = (fields: Json = {}) =>
      call({ action: "fetchReceiptEmails", ...fields });
    const subjects = (res: Json) => res.data.emails.map((e: Json) => e.subject);

    expect(fetch().data.emails).toHaveLength(5);
    // from: compares extracted addresses case-insensitively
    const venmo = fetch({ senders: ["VENMO@venmo.com"] });
    expect(venmo.data.emails).toHaveLength(1);
    expect(venmo.data.emails[0].from).toMatch(/venmo@venmo\.com/);
    expect(fetch({ senders: ["nobody@example.com"] }).data.emails).toEqual([]);
    // subject: any of the quoted terms, case-insensitive
    expect(
      subjects(fetch({ subjects: ["demo cafe", "Your statement"] })),
    ).toEqual([
      "You made a $18.42 transaction with Demo Cafe",
      "Your statement is ready",
    ]);

    // in:inbox drops archived and Receipts/Scripted threads; includeProcessed searches all mail
    call({ action: "syncReceipts" });
    expect(fetch().data.emails).toEqual([]);
    const all = fetch({ includeProcessed: true });
    expect(all.data.emails).toHaveLength(5);
    expect(all.data.emails[0].labels).toContain("Receipts/Scripted");
  });

  it("addThreads accepts html for fetchReceiptEmails with includeHtml", async () => {
    const { call, backend } = await create();
    backend.addThreads([
      {
        from: "Shop <receipts@shop.example.com>",
        subject: "Your order",
        body: "plain",
        html: "<b>Total $5</b>",
      },
    ]);
    const res = call({
      action: "fetchReceiptEmails",
      senders: ["receipts@shop.example.com"],
      includeHtml: true,
    });
    expect(res.data.emails).toHaveLength(1);
    expect(res.data.emails[0].htmlBody).toBe("<b>Total $5</b>");
  });

  it("dumpVariables lists the Variables sheet's key/value rows, including saved ReceiptRules", async () => {
    const { call, backend } = await create();
    const before = backend.dumpVariables();
    expect(before).toContainEqual(["ForwardName", "Sam"]);
    expect(before.map(([key]) => key)).not.toContain("ReceiptRules");

    const { DEFAULT_RECEIPT_RULES } = await import("@budget/shared");
    expect(
      call({ action: "saveReceiptRules", rules: DEFAULT_RECEIPT_RULES }).ok,
    ).toBe(true);
    expect(backend.dumpVariables()).toContainEqual([
      "ReceiptRules",
      JSON.stringify(DEFAULT_RECEIPT_RULES, null, 2),
    ]);

    backend.reset();
    expect(backend.dumpVariables().map(([key]) => key)).not.toContain(
      "ReceiptRules",
    );
  });
});
