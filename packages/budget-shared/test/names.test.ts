import { describe, expect, it } from "vitest";
import {
  buildSplitChildName,
  buildSplitParentName,
  getSplitNamePrefix,
  isUnsafeFreeText,
} from "../src/index.ts";

describe("buildSplitParentName", () => {
  it("appends ' | ' to an unsplit name", () => {
    expect(buildSplitParentName("Store")).toBe("Store | ");
  });

  it("leaves a name that already contains '|' unchanged", () => {
    expect(buildSplitParentName("Store | ")).toBe("Store | ");
    expect(buildSplitParentName("Store | milk")).toBe("Store | milk");
    expect(buildSplitParentName("Store|x")).toBe("Store|x");
  });
});

describe("buildSplitChildName", () => {
  it("uses the merchant before ' | ' plus the item note", () => {
    expect(buildSplitChildName("Store", "milk")).toBe("Store | milk");
    expect(buildSplitChildName("Store | ", "eggs")).toBe("Store | eggs");
    expect(buildSplitChildName("Store | milk", "eggs")).toBe("Store | eggs");
  });

  it("allows an empty item note like the native split", () => {
    expect(buildSplitChildName("Store | ", "")).toBe("Store | ");
  });

  it("extracts the prefix", () => {
    expect(getSplitNamePrefix("A | B | C")).toBe("A");
    expect(getSplitNamePrefix("A")).toBe("A");
  });
});

describe("isUnsafeFreeText", () => {
  it.each(["=SUM(A1)", "+1", "-1", "@cmd", "  =1", "\t-x"])(
    "flags %j",
    (text) => {
      expect(isUnsafeFreeText(text)).toBe(true);
    },
  );

  it.each(["", "Store", "a=b", "1+1", "Target - milk", "5"])(
    "allows %j",
    (text) => {
      expect(isUnsafeFreeText(text)).toBe(false);
    },
  );
});
