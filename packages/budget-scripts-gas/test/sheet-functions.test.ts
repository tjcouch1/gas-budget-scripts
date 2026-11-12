import { describe, expect, it } from "vitest";
import { getEnv } from "./gas-mocks";

const load = () => import("../src/index");

describe("getMonthOffset", () => {
  it("offsets months forward, wrapping around the year", async () => {
    const { getMonthOffset } = await load();
    expect(getMonthOffset("January", 0, 1)).toBe("January");
    expect(getMonthOffset("January", 1, 1)).toBe("February");
    expect(getMonthOffset("November", 2, 1)).toBe("January");
    expect(getMonthOffset("December", 13, 1)).toBe("January");
  });

  it("returns a message for unknown months", async () => {
    const { getMonthOffset } = await load();
    expect(getMonthOffset("Smarch", 1, 1)).toBe("No month named Smarch");
    // Case sensitive
    expect(getMonthOffset("january", 1, 1)).toBe("No month named january");
  });

  it("returns undefined for negative offsets past January", async () => {
    const { getMonthOffset } = await load();
    expect(getMonthOffset("March", -1, 1)).toBe("February");
    // NOTE: possible bug: JS `%` keeps the sign, so MONTHS[-1] is undefined instead of wrapping to December
    expect(getMonthOffset("January", -1, 1)).toBeUndefined();
  });
});

describe("sheetNameOffset", () => {
  it("returns the name of the sheet offset from the active sheet", async () => {
    const ss = getEnv().spreadsheet;
    ss.addSheet("A");
    const b = ss.addSheet("B");
    ss.addSheet("C");
    ss.activeSheet = b;
    const { sheetNameOffset } = await load();
    expect(sheetNameOffset(0, 1)).toBe("B");
    expect(sheetNameOffset(1, 1)).toBe("C");
    expect(sheetNameOffset(-1, 1)).toBe("A");
  });

  it("returns a message when there is no sheet at the offset", async () => {
    const ss = getEnv().spreadsheet;
    ss.addSheet("A");
    const { sheetNameOffset } = await load();
    expect(sheetNameOffset(1, 1)).toBe("No sheet at offset 1");
    expect(sheetNameOffset(-1, 1)).toBe("No sheet at offset -1");
  });
});
