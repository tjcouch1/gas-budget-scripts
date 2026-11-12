import { useState } from "react";
import type { SplitItemInput } from "@budget/shared";
import { formatMoney } from "../lib/format.ts";
import {
  summarizeSplit,
  toSplitInputs,
  type SplitDraftItem,
} from "../lib/split-math.ts";
import { Button, inputClass } from "./ui.tsx";

let keyCounter = 0;
const newItem = (): SplitDraftItem => ({
  key: `item-${keyCounter++}`,
  itemNote: "",
  expression: "",
  preTax: false,
  envelope: "",
});

/**
 * Mirrors the spreadsheet: the original row stays on top and keeps whatever the
 * new items do not take (a locked, live-calculated entry). Each item below it is
 * a new child row.
 */
export function SplitEditor({
  parentName,
  parentValue,
  taxMultiplier,
  envelopeOptions,
  onSubmit,
  onCancel,
}: {
  parentName: string;
  /** The parent's current value; the parent keeps whatever is not allocated. */
  parentValue: number | null;
  taxMultiplier: number;
  envelopeOptions: string[] | null;
  onSubmit: (splits: SplitItemInput[]) => void;
  onCancel: () => void;
}) {
  const [items, setItems] = useState<SplitDraftItem[]>(() => [newItem()]);
  const summary = summarizeSplit(parentValue, items, taxMultiplier);
  const overAllocated = summary.remaining < -0.005;

  const update = (key: string, patch: Partial<SplitDraftItem>) =>
    setItems((list) =>
      list.map((i) => (i.key === key ? { ...i, ...patch } : i)),
    );

  return (
    <form
      className="space-y-3"
      onSubmit={(e) => {
        e.preventDefault();
        if (summary.canSubmit) onSubmit(toSplitInputs(items));
      }}
    >
      <p className="text-sm text-slate-600 dark:text-slate-400">
        Each item below is taken out of the original. Amounts accept math like{" "}
        <code>4.23*2+12.09</code>.
      </p>
      <div
        data-testid="original-entry"
        className={`rounded-xl border border-dashed p-3 ${
          overAllocated
            ? "border-red-400 bg-red-50 dark:bg-red-950/30"
            : "border-slate-300 bg-slate-100 dark:border-slate-600 dark:bg-slate-800/60"
        }`}
      >
        <div className="flex items-baseline justify-between gap-3">
          <div className="min-w-0">
            <p className="text-xs font-medium uppercase tracking-wide text-slate-500">
              Remaining on original
            </p>
            <p className="truncate text-sm text-slate-700 dark:text-slate-300">
              {parentName || "(no name)"} · was {formatMoney(parentValue)}
            </p>
          </div>
          <p
            data-testid="remaining"
            className={`shrink-0 text-xl font-semibold tabular-nums ${
              overAllocated ? "text-red-600" : ""
            }`}
          >
            {formatMoney(summary.remaining)}
          </p>
        </div>
        {overAllocated && (
          <p role="alert" className="mt-1 text-xs text-red-600">
            The items add up to more than the original amount. Reduce an item to
            continue.
          </p>
        )}
      </div>
      {items.map((item, i) => {
        const ev = summary.items[i];
        const invalid = item.expression.trim() !== "" && ev.value === null;
        return (
          <fieldset
            key={item.key}
            className="relative space-y-2 rounded-xl border border-slate-200 p-3 dark:border-slate-700"
          >
            <legend className="px-1 text-xs font-medium text-slate-500">
              Item {i + 1}
            </legend>
            {items.length > 1 && (
              <button
                type="button"
                aria-label="Remove item"
                title="Remove item"
                onClick={() =>
                  setItems((l) => l.filter((x) => x.key !== item.key))
                }
                className="absolute right-1 top-1 flex size-9 items-center justify-center rounded-lg text-xl leading-none text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
              >
                ×
              </button>
            )}
            <input
              className={`${inputClass} ${items.length > 1 ? "pr-10" : ""}`}
              aria-label={`Item ${i + 1} note`}
              placeholder="Item note (optional)"
              value={item.itemNote}
              onChange={(e) => update(item.key, { itemNote: e.target.value })}
            />
            {ev.noteUnsafe && (
              <p className="text-xs text-red-600">
                Notes cannot start with = + - or @.
              </p>
            )}
            <div className="flex items-center gap-2">
              <input
                className={`${inputClass} ${invalid ? "border-red-500" : ""}`}
                aria-label={`Item ${i + 1} amount`}
                inputMode="decimal"
                placeholder="Amount, e.g. 4.23*2+3"
                value={item.expression}
                onChange={(e) =>
                  update(item.key, { expression: e.target.value })
                }
              />
              <span
                className="w-24 shrink-0 text-right text-sm tabular-nums"
                aria-label={`Item ${i + 1} value`}
              >
                {ev.value === null ? "" : formatMoney(ev.value)}
              </span>
            </div>
            <label className="flex min-h-11 items-center gap-2 text-sm">
              <input
                type="checkbox"
                checked={item.preTax}
                onChange={(e) => update(item.key, { preTax: e.target.checked })}
              />
              <span>
                Add tax <span className="tabular-nums">×{taxMultiplier}</span>
                <span className="block text-xs text-slate-500">
                  Multiplies this amount by the tax rate
                </span>
              </span>
            </label>
            {envelopeOptions && (
              <select
                className={inputClass}
                aria-label={`Item ${i + 1} envelope`}
                value={item.envelope}
                onChange={(e) => update(item.key, { envelope: e.target.value })}
              >
                <option value="">Envelope…</option>
                {envelopeOptions.map((o) => (
                  <option key={o} value={o}>
                    {o}
                  </option>
                ))}
              </select>
            )}
          </fieldset>
        );
      })}
      <Button onClick={() => setItems((l) => [...l, newItem()])}>
        + Add item
      </Button>
      <div className="flex gap-2">
        <Button variant="primary" type="submit" disabled={!summary.canSubmit}>
          Split
        </Button>
        <Button onClick={onCancel}>Back</Button>
      </div>
    </form>
  );
}
