import { afterEach, describe, expect, it, vi } from "vitest";
import { randomId } from "./random-id.ts";

describe("randomId", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("uses randomUUID when available", () => {
    expect(randomId()).toMatch(/^[0-9a-f-]{36}$/);
  });

  it("falls back to getRandomValues in an insecure context (no randomUUID)", () => {
    vi.stubGlobal("crypto", {
      getRandomValues: globalThis.crypto.getRandomValues.bind(
        globalThis.crypto,
      ),
    });
    const a = randomId();
    expect(a).toMatch(/^[0-9a-f]{32}$/);
    expect(randomId()).not.toBe(a);
  });
});
