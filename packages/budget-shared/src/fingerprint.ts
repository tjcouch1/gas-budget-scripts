import type { FingerprintRow } from "./types.ts";

function pad(n: number, width = 2): string {
  return String(n).padStart(width, "0");
}

/** Formats a Date as yyyy-MM-dd using local time (GAS runs in the spreadsheet time zone). */
export function formatLocalDate(date: Date): string {
  return `${pad(date.getFullYear(), 4)}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function normalizeText(value: string | null | undefined): string {
  return (value ?? "").replace(/\s+$/, "");
}

function normalizeDate(value: Date | string | null | undefined): string {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "" : formatLocalDate(value);
  }
  return normalizeText(value);
}

function normalizeCost(row: FingerprintRow): string | number {
  if (row.costFormula) return normalizeText(row.costFormula);
  const cost = row.cost;
  if (cost === null || cost === undefined) return "";
  if (typeof cost === "number") return cost;
  return normalizeText(cost);
}

/** 32-bit FNV-1a hash of a string, as 8 lowercase hex digits. */
export function fnv1a32(input: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

/**
 * Fingerprint of a transaction row's editable state (web-plan.md section 4.4). Pure
 * and synchronous so it runs identically in the web client and in GAS.
 * Digest of JSON.stringify([date, name, costFormula ?? cost, envelope, paidWith, taxDeductible]).
 */
export function computeRowFingerprint(row: FingerprintRow): string {
  return fnv1a32(
    JSON.stringify([
      normalizeDate(row.date),
      normalizeText(row.name),
      normalizeCost(row),
      normalizeText(row.envelope),
      normalizeText(row.paidWith),
      normalizeText(row.taxDeductible),
    ]),
  );
}
