import type { SheetInfoDto } from "@budget/shared";

/** A multi-period choice in the pay-period dropdown. "all" is every pay period. */
export type PeriodRange = "all" | "ytd" | "3m" | "6m" | "1y";

export const PERIOD_RANGES: {
  id: Exclude<PeriodRange, "all">;
  label: string;
}[] = [
  { id: "ytd", label: "Year to date" },
  { id: "3m", label: "Last 3 months" },
  { id: "6m", label: "Last 6 months" },
  { id: "1y", label: "Last year" },
];

/** Prefix of the dropdown value of a range option (a leading NUL never collides with a sheet name). */
const RANGE_PREFIX = "\u0000range:";

export const rangeOptionValue = (id: Exclude<PeriodRange, "all">): string =>
  RANGE_PREFIX + id;

/** The range a dropdown value stands for, or null when it is not a (non-"all") range option. */
export function parseRangeOption(value: string): PeriodRange | null {
  if (!value.startsWith(RANGE_PREFIX)) return null;
  const id = value.slice(RANGE_PREFIX.length);
  return PERIOD_RANGES.some((r) => r.id === id) ? (id as PeriodRange) : null;
}

const pad = (n: number) => String(n).padStart(2, "0");
const formatDate = (d: Date) =>
  `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/** First day (yyyy-MM-dd) of the window that ends on `today`. */
export function rangeStart(
  range: Exclude<PeriodRange, "all">,
  today: Date,
): string {
  if (range === "ytd") return `${today.getFullYear()}-01-01`;
  const months = range === "3m" ? 3 : range === "6m" ? 6 : 12;
  // Clamp the day so e.g. May 31 minus 3 months is Feb 28/29, not a March date
  const start = new Date(today.getFullYear(), today.getMonth() - months, 1);
  const lastDay = new Date(
    start.getFullYear(),
    start.getMonth() + 1,
    0,
  ).getDate();
  start.setDate(Math.min(today.getDate(), lastDay));
  return formatDate(start);
}

/**
 * The pay periods to show for a range: those whose date range overlaps the window ending today
 * (an approximation, since pay periods are not aligned with months). "all" keeps every period.
 * Order is kept.
 */
export function sheetsInRange(
  sheets: SheetInfoDto[],
  range: PeriodRange,
  today: Date = new Date(),
): SheetInfoDto[] {
  if (range === "all") return sheets;
  const start = rangeStart(range, today);
  const end = formatDate(today);
  // yyyy-MM-dd strings compare correctly as text
  return sheets.filter((s) => s.endDate >= start && s.startDate <= end);
}
