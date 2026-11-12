import { describe, expect, it } from "vitest";
import { isGroupName, joinNameParts, splitNameParts } from "./name-parts.ts";

describe("name parts", () => {
  it("splits a split-row name into prefix and item note", () => {
    expect(splitNameParts("Costco | milk | 2%")).toEqual({
      prefix: "Costco",
      note: "milk | 2%",
    });
    expect(splitNameParts("Costco | ")).toEqual({ prefix: "Costco", note: "" });
    expect(splitNameParts("A|b")).toEqual({ prefix: "A|b", note: "" });
  });
  it("joins them back the way the sheet does", () => {
    expect(joinNameParts("Costco", "milk")).toBe("Costco | milk");
    expect(joinNameParts("Costco", "")).toBe("Costco | ");
    const { prefix, note } = splitNameParts("Costco | milk");
    expect(joinNameParts(prefix, note)).toBe("Costco | milk");
  });
  it("detects group names", () => {
    expect(isGroupName("A | b")).toBe(true);
    expect(isGroupName("A")).toBe(false);
  });
});
