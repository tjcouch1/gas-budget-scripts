import { afterEach, describe, expect, it, vi } from "vitest";
import { getTestingMode, setTestingMode } from "./testing-mode.ts";

afterEach(() => {
  vi.restoreAllMocks();
  localStorage.clear();
});

describe("testing mode preference", () => {
  it("defaults off and persists", () => {
    expect(getTestingMode()).toBe(false);
    setTestingMode(true);
    expect(localStorage.getItem("budget.testingMode")).toBe("1");
    expect(getTestingMode()).toBe(true);
    setTestingMode(false);
    expect(getTestingMode()).toBe(false);
  });

  it("does not throw when storage is unavailable", () => {
    vi.spyOn(Storage.prototype, "getItem").mockImplementation(() => {
      throw new Error("blocked");
    });
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw new Error("blocked");
    });
    expect(getTestingMode()).toBe(false);
    expect(() => setTestingMode(true)).not.toThrow();
  });
});
