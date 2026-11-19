import { describe, expect, it, vi } from "vitest";
import { MemoryKeyValueStore } from "./crypto/kv-store.ts";
import { createServices } from "./services.ts";

describe("createServices", () => {
  it("exposes the authenticated api client", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve({
        text: () =>
          Promise.resolve(JSON.stringify({ ok: true, data: { raw: "" } })),
      }),
    ) as unknown as typeof fetch;
    const services = createServices({
      vaultStore: new MemoryKeyValueStore(),
      cacheStore: new MemoryKeyValueStore(),
      fetchImpl,
    });
    await services.vault.save({
      gasWebAppUrl: "https://script.google.com/macros/s/TEST/exec",
      apiKey: "k",
    });
    const result = await services.api({ action: "getReceiptRules" });
    expect(result.data).toEqual({ raw: "" });
  });
});
