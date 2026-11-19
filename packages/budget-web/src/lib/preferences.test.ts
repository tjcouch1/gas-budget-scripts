import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  ENVELOPE_EXPANDED_KEY,
  getEnvelopeExpanded,
  setEnvelopeExpanded,
} from "./preferences.ts";

describe("envelope expanded preference", () => {
  beforeEach(() => localStorage.clear());

  it("defaults to collapsed and round-trips", () => {
    expect(getEnvelopeExpanded()).toBe(false);
    setEnvelopeExpanded(true);
    expect(localStorage.getItem(ENVELOPE_EXPANDED_KEY)).toBe("1");
    expect(getEnvelopeExpanded()).toBe(true);
    setEnvelopeExpanded(false);
    expect(getEnvelopeExpanded()).toBe(false);
  });

  it("survives unavailable storage", () => {
    const spy = vi
      .spyOn(Storage.prototype, "getItem")
      .mockImplementation(() => {
        throw new Error("blocked");
      });
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw new Error("blocked");
    });
    expect(getEnvelopeExpanded()).toBe(false);
    expect(() => setEnvelopeExpanded(true)).not.toThrow();
    spy.mockRestore();
    vi.restoreAllMocks();
  });
});

describe("recent envelopes", () => {
  beforeEach(() => localStorage.clear());

  it("keeps the most recent choice first without duplicates", async () => {
    const { getRecentEnvelopes, recordRecentEnvelope } =
      await import("./preferences.ts");
    expect(getRecentEnvelopes()).toEqual([]);
    recordRecentEnvelope("Gas");
    recordRecentEnvelope("Groceries");
    recordRecentEnvelope("Gas");
    recordRecentEnvelope("");
    expect(getRecentEnvelopes()).toEqual(["Gas", "Groceries"]);
  });

  it("ignores corrupt storage", async () => {
    const { getRecentEnvelopes } = await import("./preferences.ts");
    localStorage.setItem("budget.recentEnvelopes", "{oops");
    expect(getRecentEnvelopes()).toEqual([]);
  });
});

describe("inline envelope chips preference", () => {
  beforeEach(() => localStorage.clear());

  it("defaults to the dropdown and round-trips", async () => {
    const { getInlineEnvelopeChips, setInlineEnvelopeChips } =
      await import("./preferences.ts");
    expect(getInlineEnvelopeChips()).toBe(false);
    setInlineEnvelopeChips(true);
    expect(localStorage.getItem("budget.inlineEnvelopeChips")).toBe("1");
    expect(getInlineEnvelopeChips()).toBe(true);
    setInlineEnvelopeChips(false);
    expect(getInlineEnvelopeChips()).toBe(false);
  });
});
