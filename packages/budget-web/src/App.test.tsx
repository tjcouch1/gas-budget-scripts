import { describe, expect, it, vi } from "vitest";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import type { PayPeriodStateDto, TransactionRowDto } from "@budget/shared";
import { App } from "./App.tsx";
import { MemoryKeyValueStore } from "./crypto/kv-store.ts";
import { createServices } from "./services.ts";
import { makeRow, makeState } from "./test/fixtures.ts";

const URL = "https://script.google.com/macros/s/TEST/exec";

interface Sent {
  action: string;
  [k: string]: unknown;
}

async function setup(respond: (req: Sent) => unknown) {
  const sent: Sent[] = [];
  const fetchImpl = vi.fn((_url: string, init: RequestInit) => {
    const req = JSON.parse(init.body as string) as Sent;
    sent.push(req);
    const body = respond(req);
    return Promise.resolve({
      text: () => Promise.resolve(JSON.stringify(body)),
    } as Response);
  });
  const services = createServices({
    vaultStore: new MemoryKeyValueStore(),
    cacheStore: new MemoryKeyValueStore(),
    fetchImpl: fetchImpl as unknown as typeof fetch,
  });
  await services.vault.save({ gasWebAppUrl: URL, apiKey: "KEY" });
  render(<App services={services} />);
  return { sent, services, user: userEvent.setup() };
}

const rows = (): TransactionRowDto[] => [
  makeRow(0, { name: "Aldi" }),
  makeRow(1, { name: "Shell", envelope: "Gas" }),
];

describe("App", () => {
  it("opens to uncategorized expenses and categorizes in two taps", async () => {
    const state = makeState(rows());
    const { sent, user } = await setup((req) => {
      if (req.action === "getState") return { ok: true, data: state };
      return {
        ok: true,
        data: { ...rows()[0], envelope: "Groceries", fingerprint: "new" },
      };
    });
    await user.click(await screen.findByRole("button", { name: /Aldi/ }));
    expect(screen.queryByRole("button", { name: /^Shell/ })).toBeNull(); // filtered out
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    await waitFor(() =>
      expect(sent.find((s) => s.action === "updateTransaction")).toMatchObject({
        sheetName: state.sheetName,
        rowIndex: 0,
        changes: { envelope: "Groceries" },
        apiKey: "KEY",
      }),
    );
    await waitFor(() =>
      expect(screen.getByRole("status", { name: "" })).toBeInTheDocument(),
    );
  });

  it("shows the conflict dialog and lets the user accept the remote value", async () => {
    const state = makeState(rows());
    const remote = makeRow(0, {
      name: "Aldi",
      envelope: "Gas",
      fingerprint: "remote",
    });
    const { user } = await setup((req) => {
      if (req.action === "getState") return { ok: true, data: state };
      return {
        ok: false,
        error: { code: "conflict", message: "changed" },
        conflict: { current: remote },
      };
    });
    await user.click(await screen.findByRole("button", { name: /Aldi/ }));
    await user.click(screen.getByRole("button", { name: /Groceries/ }));
    expect(
      await screen.findByRole("dialog", {
        name: "This row changed in the sheet",
      }),
    ).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveAttribute(
      "data-status",
      "conflict",
    );
    await user.click(screen.getByRole("button", { name: "Accept Remote" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeNull());
  });

  it("shows the Settings dialog when no credentials are saved", async () => {
    const services = createServices({
      vaultStore: new MemoryKeyValueStore(),
      cacheStore: new MemoryKeyValueStore(),
      fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    render(<App services={services} />);
    expect(
      await screen.findByRole("dialog", { name: "Connection settings" }),
    ).toBeInTheDocument();
    // The key input is never pre-filled.
    expect(screen.getByLabelText("API key")).toHaveValue("");
  });

  it("switches to the envelopes view", async () => {
    const state: PayPeriodStateDto = makeState(rows());
    const { user } = await setup(() => ({ ok: true, data: state }));
    await screen.findByRole("button", { name: /Aldi/ });
    await user.click(screen.getByRole("button", { name: "Envelopes" }));
    expect(screen.getByTestId("remain-Groceries")).toHaveTextContent("$280.00");
  });
});

describe("App: all pay periods", () => {
  const OLD = "2/17/26 - 3/2/26";
  const oldRow = makeRow(0, { name: "Old Hardware", fingerprint: "old-fp" });

  it("loads every tab only when chosen, searches across tabs, and edits with the row's own sheet", async () => {
    window.location.hash = "#/expenses";
    const state = makeState(rows());
    const all = {
      taxMultiplier: 1.0825,
      sheets: [
        {
          name: state.sheetName,
          startDate: "2026-03-03",
          endDate: "2026-03-16",
          isGap: false,
          transactions: rows(),
          envelopeOptions: ["Groceries", "Gas"],
          paidWithOptions: null,
        },
        {
          name: OLD,
          startDate: "2026-02-17",
          endDate: "2026-03-02",
          isGap: false,
          transactions: [oldRow],
          envelopeOptions: ["Groceries", "Gas"],
          paidWithOptions: null,
        },
      ],
    };
    const { sent, user } = await setup((req) => {
      if (req.action === "getState") return { ok: true, data: state };
      if (req.action === "getAllTransactions") return { ok: true, data: all };
      return {
        ok: true,
        data: { ...oldRow, envelope: "Gas", fingerprint: "n" },
      };
    });
    await screen.findByRole("button", { name: /Aldi/ });
    expect(sent.some((s) => s.action === "getAllTransactions")).toBe(false);
    expect(screen.queryByRole("tab", { name: /Tax deductible/ })).toBeNull();

    await user.selectOptions(
      screen.getByLabelText("Transactions shown"),
      "All pay periods",
    );
    await user.click(
      await screen.findByRole("button", { name: /Old Hardware/ }),
    );
    // Heading per tab
    expect(screen.getByRole("heading", { name: OLD })).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: /Gas/ }));
    await waitFor(() =>
      expect(sent.find((s) => s.action === "updateTransaction")).toMatchObject({
        sheetName: OLD,
        rowIndex: 0,
        baseFingerprint: "old-fp",
        changes: { envelope: "Gas" },
      }),
    );
  });
});
