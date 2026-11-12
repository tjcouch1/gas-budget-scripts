import {
  computeRowFingerprint,
  type PayPeriodStateDto,
  type TransactionRowDto,
} from "@budget/shared";

export function makeRow(
  rowIndex: number,
  over: Partial<TransactionRowDto> = {},
): TransactionRowDto {
  const base = {
    date: "2026-03-05",
    name: `Store ${rowIndex}`,
    cost: 10 as number | null,
    costFormula: null as string | null,
    envelope: "",
    paidWith: "Credit",
    taxDeductible: "",
    ...over,
  };
  return {
    rowIndex,
    date: base.date,
    name: base.name,
    cost: base.cost,
    costFormula: base.costFormula,
    costDisplay: over.costDisplay ?? `$${(base.cost ?? 0).toFixed(2)}`,
    envelope: base.envelope,
    paidWith: base.paidWith,
    taxDeductible: base.taxDeductible,
    note: "",
    isError: false,
    isNote: false,
    splitGroupId: null,
    fingerprint: computeRowFingerprint(base),
    ...over,
  };
}

export function makeState(
  transactions: (TransactionRowDto | null)[],
  over: Partial<PayPeriodStateDto> = {},
): PayPeriodStateDto {
  return {
    sheets: [
      {
        name: "3/3/26 - 3/16/26",
        startDate: "2026-03-03",
        endDate: "2026-03-16",
        isGap: false,
      },
      {
        name: "2/17/26 - 3/2/26",
        startDate: "2026-02-17",
        endDate: "2026-03-02",
        isGap: false,
      },
    ],
    sheetName: "3/3/26 - 3/16/26",
    taxMultiplier: 1.0825,
    forwardName: "Pat",
    transactions,
    envelopes: [
      {
        name: "Groceries",
        display: [
          "$0",
          "$400.00",
          "$120.00",
          "$280.00",
          "$280.00",
          "",
          "",
          "",
          "",
          "",
        ],
        values: [0, 400, 120, 280, 280, null, null, null, null, null],
      },
      {
        name: "Gas",
        display: [
          "$0",
          "$100.00",
          "$130.00",
          "-$30.00",
          "-$30.00",
          "",
          "",
          "",
          "",
          "",
        ],
        values: [0, 100, 130, -30, -30, null, null, null, null, null],
      },
    ],
    paidWithOptions: ["Credit", "Pat Venmo"],
    envelopeOptions: ["Groceries", "Gas"],
    ...over,
  };
}

export const flush = () => new Promise<void>((r) => setTimeout(r, 0));
