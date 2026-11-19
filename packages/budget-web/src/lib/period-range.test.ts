import { describe, expect, it } from "vitest";
import type { SheetInfoDto } from "@budget/shared";
import {
  parseRangeOption,
  rangeOptionValue,
  rangeStart,
  sheetsInRange,
} from "./period-range.ts";

const sheet = (startDate: string, endDate: string): SheetInfoDto => ({
  name: `${startDate} - ${endDate}`,
  startDate,
  endDate,
  isGap: false,
});

// Newest first, like the server sends them
const SHEETS = [
  sheet("2026-11-06", "2026-11-19"),
  sheet("2026-10-23", "2026-11-05"),
  sheet("2026-08-14", "2026-08-27"),
  sheet("2026-05-01", "2026-05-14"),
  sheet("2026-01-02", "2026-01-15"),
  sheet("2025-12-19", "2026-01-01"),
  sheet("2025-10-10", "2025-10-23"),
];
const TODAY = new Date(2026, 10, 18); // 2026-11-18

describe("period ranges", () => {
  it("computes the window start", () => {
    expect(rangeStart("ytd", TODAY)).toBe("2026-01-01");
    expect(rangeStart("3m", TODAY)).toBe("2026-08-18");
    expect(rangeStart("6m", TODAY)).toBe("2026-05-18");
    expect(rangeStart("1y", TODAY)).toBe("2025-11-18");
    // Day clamped to the shorter month
    expect(rangeStart("3m", new Date(2026, 4, 31))).toBe("2026-02-28");
  });

  it("keeps the periods that overlap the window ending today", () => {
    const names = (r: Parameters<typeof sheetsInRange>[1]) =>
      sheetsInRange(SHEETS, r, TODAY).map((s) => s.startDate);
    expect(names("all")).toHaveLength(SHEETS.length);
    expect(names("ytd")).toEqual([
      "2026-11-06",
      "2026-10-23",
      "2026-08-14",
      "2026-05-01",
      "2026-01-02",
      "2025-12-19", // ends 2026-01-01, inside the year
    ]);
    // Aug 14-27 overlaps a window starting Aug 18
    expect(names("3m")).toEqual(["2026-11-06", "2026-10-23", "2026-08-14"]);
    // May 1-14 ends before the window starts (May 18)
    expect(names("6m")).toEqual(["2026-11-06", "2026-10-23", "2026-08-14"]);
    expect(names("1y")).toHaveLength(6);
  });

  it("excludes periods that start after today", () => {
    const future = [sheet("2026-11-20", "2026-12-03"), ...SHEETS];
    expect(sheetsInRange(future, "3m", TODAY)).toHaveLength(3);
  });

  it("round-trips dropdown values", () => {
    expect(parseRangeOption(rangeOptionValue("6m"))).toBe("6m");
    expect(parseRangeOption("3/3/26 - 3/16/26")).toBeNull();
    expect(parseRangeOption("\u0000range:bogus")).toBeNull();
  });
});
