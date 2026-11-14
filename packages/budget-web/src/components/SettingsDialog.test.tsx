import { afterEach, describe, expect, it, vi } from "vitest";
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { MemoryKeyValueStore } from "../crypto/kv-store.ts";
import { createServices } from "../services.ts";
import { makeRow, makeState } from "../test/fixtures.ts";
import { SettingsDialog } from "./SettingsDialog.tsx";

const URL = "https://script.google.com/macros/s/TEST/exec";

async function setup(fetchImpl?: typeof fetch, saved = false) {
  const vaultStore = new MemoryKeyValueStore();
  const services = createServices({
    vaultStore,
    cacheStore: new MemoryKeyValueStore(),
    fetchImpl:
      fetchImpl ??
      (vi.fn(() =>
        Promise.resolve({
          text: () =>
            Promise.resolve(JSON.stringify({ ok: true, data: makeState([]) })),
        }),
      ) as unknown as typeof fetch),
  });
  if (saved) await services.vault.save({ gasWebAppUrl: URL, apiKey: "k" });
  const onChanged = vi.fn();
  render(
    <SettingsDialog
      services={services}
      onClose={() => {}}
      onChanged={onChanged}
    />,
  );
  return { services, vaultStore, onChanged, user: userEvent.setup() };
}

describe("SettingsDialog", () => {
  it("saves encrypted credentials and clears the inputs", async () => {
    const { user, vaultStore, onChanged } = await setup();
    await user.type(screen.getByLabelText("Web app URL"), URL);
    await user.type(screen.getByLabelText("API key"), "sekrit-key");
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(await screen.findByText(/Saved \(encrypted/)).toBeInTheDocument();
    expect(onChanged).toHaveBeenCalled();
    expect(screen.getByLabelText("API key")).toHaveValue("");
    expect(vaultStore.data.has("credential-blob")).toBe(true);
    expect(JSON.stringify([...vaultStore.data.keys()])).not.toContain("sekrit");
  });

  it("rejects a malformed URL", async () => {
    const { user, vaultStore } = await setup();
    await user.type(
      screen.getByLabelText("Web app URL"),
      "https://example.com/exec",
    );
    await user.type(screen.getByLabelText("API key"), "k");
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(await screen.findByRole("alert")).toHaveTextContent(
      /URL must look like/,
    );
    expect(vaultStore.data.size).toBe(0);
  });

  it("tests the connection without saving", async () => {
    const { user, vaultStore } = await setup();
    await user.type(screen.getByLabelText("Web app URL"), URL);
    await user.type(screen.getByLabelText("API key"), "k");
    await user.click(screen.getByRole("button", { name: "Test connection" }));
    expect(
      await screen.findByText(/Connected\. 2 pay-period sheets/),
    ).toBeInTheDocument();
    expect(vaultStore.data.size).toBe(0);
  });

  it("forgets credentials", async () => {
    const { user, vaultStore } = await setup(undefined, true);
    const forget = await screen.findByRole("button", {
      name: "Forget credentials",
    });
    await vi.waitFor(() => expect(forget).toBeEnabled());
    await user.click(forget);
    expect(
      await screen.findByText(/removed from this device/),
    ).toBeInTheDocument();
    expect(vaultStore.data.size).toBe(0);
  });

  it("links to the licenses page with a relative URL", async () => {
    await setup();
    expect(screen.getByRole("link", { name: /licenses/i })).toHaveAttribute(
      "href",
      "./licenses.html",
    );
  });
});

describe("SettingsDialog switching sheets", () => {
  const OTHER = "https://script.google.com/macros/s/OTHER/exec";

  async function setupSwitch() {
    const state = makeState([makeRow(0, { name: "Aldi" })]);
    const fetchImpl = vi.fn((_u: string, init: RequestInit) => {
      const req = JSON.parse(init.body as string);
      if (req.action === "getState")
        return Promise.resolve({
          text: () =>
            Promise.resolve(JSON.stringify({ ok: true, data: state })),
        });
      return Promise.reject(new TypeError("offline"));
    }) as unknown as typeof fetch;
    const services = createServices({
      vaultStore: new MemoryKeyValueStore(),
      cacheStore: new MemoryKeyValueStore(),
      fetchImpl,
    });
    await services.vault.save({ gasWebAppUrl: URL, apiKey: "k" });
    await services.engine.refresh();
    const clear = vi.spyOn(services, "clearLocalData");
    render(
      <SettingsDialog
        services={services}
        onClose={() => {}}
        onChanged={vi.fn()}
      />,
    );
    return { services, clear, user: userEvent.setup() };
  }

  it("clears local data when the URL changes", async () => {
    const { services, clear, user } = await setupSwitch();
    await user.type(screen.getByLabelText("Web app URL"), OTHER);
    await user.type(screen.getByLabelText("API key"), "k2");
    await user.click(screen.getByRole("button", { name: "Save" }));
    await screen.findByText(/Saved \(encrypted/);
    expect(clear).toHaveBeenCalledTimes(1);
    expect(services.engine.getSnapshot().states).toEqual({});
    expect((await services.vault.load())?.gasWebAppUrl).toBe(OTHER);
  });

  it("keeps local data for the same URL (surrounding spaces ignored)", async () => {
    const { services, clear, user } = await setupSwitch();
    await user.type(screen.getByLabelText("Web app URL"), `  ${URL} `);
    await user.type(screen.getByLabelText("API key"), "new-key");
    await user.click(screen.getByRole("button", { name: "Save" }));
    await screen.findByText(/Saved \(encrypted/);
    expect(clear).not.toHaveBeenCalled();
    expect(Object.keys(services.engine.getSnapshot().states)).toHaveLength(1);
  });

  it("requires a second click to discard unsynced changes", async () => {
    const { services, clear, user } = await setupSwitch();
    const sheet = Object.keys(services.engine.getSnapshot().states)[0]!;
    const row = services.engine.getSnapshot().states[sheet].transactions[0]!;
    services.engine.enqueueUpdate(sheet, row, { note: "x" });
    await user.type(screen.getByLabelText("Web app URL"), OTHER);
    await user.type(screen.getByLabelText("API key"), "k2");
    await user.click(screen.getByRole("button", { name: "Save" }));
    expect(
      await screen.findByText(
        "1 unsynced change from the old sheet will be discarded.",
      ),
    ).toBeInTheDocument();
    expect(clear).not.toHaveBeenCalled();
    expect((await services.vault.load())?.gasWebAppUrl).toBe(URL);
    await user.click(
      screen.getByRole("button", { name: "Discard and switch" }),
    );
    await screen.findByText(/Saved \(encrypted/);
    expect(clear).toHaveBeenCalledTimes(1);
    expect(services.engine.getSnapshot().queue).toEqual([]);
    expect((await services.vault.load())?.gasWebAppUrl).toBe(OTHER);
  });

  it("forgetting credentials clears local data", async () => {
    const { services, clear, user } = await setupSwitch();
    const forget = await screen.findByRole("button", {
      name: "Forget credentials",
    });
    await vi.waitFor(() => expect(forget).toBeEnabled());
    await user.click(forget);
    await screen.findByText(/removed from this device/);
    expect(clear).toHaveBeenCalled();
    expect(services.engine.getSnapshot().states).toEqual({});
  });
});

describe("SettingsDialog testing mode", () => {
  it("toggles the per-device preference", async () => {
    localStorage.clear();
    const { user } = await setup();
    const box = screen.getByRole("checkbox", {
      name: /Testing mode: import receipts without marking them in Gmail/,
    });
    expect(box).not.toBeChecked();
    await user.click(box);
    expect(localStorage.getItem("budget.testingMode")).toBe("1");
    await user.click(box);
    expect(localStorage.getItem("budget.testingMode")).toBeNull();
  });
});

describe("SettingsDialog auto close", () => {
  async function save(onChanged: () => Promise<boolean>, auto: boolean) {
    vi.useFakeTimers({ shouldAdvanceTime: true });
    const services = createServices({
      vaultStore: new MemoryKeyValueStore(),
      cacheStore: new MemoryKeyValueStore(),
      fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    const onClose = vi.fn();
    render(
      <SettingsDialog
        services={services}
        onClose={onClose}
        onChanged={onChanged}
        autoCloseOnConnect={auto}
      />,
    );
    const user = userEvent.setup({ advanceTimers: vi.advanceTimersByTime });
    await user.type(screen.getByLabelText("Web app URL"), URL);
    await user.type(screen.getByLabelText("API key"), "k");
    await user.click(screen.getByRole("button", { name: "Save" }));
    return onClose;
  }

  afterEach(() => vi.useRealTimers());

  it("shows Connected. and closes after ~2 seconds when enabled", async () => {
    const onClose = await save(() => Promise.resolve(true), true);
    expect(await screen.findByText("Saved. Connected.")).toBeInTheDocument();
    expect(onClose).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(2100);
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("does not close when the connection fails or auto-close is off", async () => {
    const failed = await save(() => Promise.resolve(false), true);
    await screen.findByText(/Saved \(encrypted/);
    await vi.advanceTimersByTimeAsync(3000);
    expect(failed).not.toHaveBeenCalled();
    vi.useRealTimers();
  });

  it("does not close when opened manually", async () => {
    const manual = await save(() => Promise.resolve(true), false);
    await screen.findByText(/Saved \(encrypted/);
    await vi.advanceTimersByTimeAsync(3000);
    expect(manual).not.toHaveBeenCalled();
  });
});
