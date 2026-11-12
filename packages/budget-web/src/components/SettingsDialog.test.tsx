import { describe, expect, it, vi } from "vitest";
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { MemoryKeyValueStore } from "../crypto/kv-store.ts";
import { createServices } from "../services.ts";
import { makeState } from "../test/fixtures.ts";
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
