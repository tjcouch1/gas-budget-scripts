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
  const [items, setItems] = useState<SplitDraftItem[]>(() => [
    newItem(),
    newItem(),
  ]);
  const summary = summarizeSplit(parentValue, items, taxMultiplier);

  const update = (key: string, patch: Partial<SplitDraftItem>) =>
    setItems((list) =>
      list.map((i) => (i.key === key ? { ...i, ...patch } : i)),
    );

  function fillRest(index: number) {
    const own = summary.items[index].value ?? 0;
    const rest = summary.remaining + own;
    const item = items[index];
    const raw = item.preTax ? rest / taxMultiplier : rest;
    update(item.key, { expression: String(Math.round(raw * 100) / 100) });
  }

  return (
    <form
      className="space-y-3"
      onSubmit={(e) => {
        e.preventDefault();
        if (summary.canSubmit) onSubmit(toSplitInputs(items));
      }}
    >
      <p className="text-sm text-slate-600 dark:text-slate-400">
        Splitting <strong>{parentName}</strong> ({formatMoney(parentValue)}).
        Amounts accept math like <code>4.23*2+12.09</code>.
      </p>
      {items.map((item, i) => {
        const ev = summary.items[i];
        const invalid = item.expression.trim() !== "" && ev.value === null;
        return (
          <fieldset
            key={item.key}
            className="space-y-2 rounded-xl border border-slate-200 p-3 dark:border-slate-700"
          >
            <legend className="px-1 text-xs font-medium text-slate-500">
              Item {i + 1}
            </legend>
            <input
              className={inputClass}
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
            <div className="flex flex-wrap items-center gap-3">
              <label className="flex min-h-11 items-center gap-2 text-sm">
                <input
                  type="checkbox"
                  checked={item.preTax}
                  onChange={(e) =>
                    update(item.key, { preTax: e.target.checked })
                  }
                />
                Pre-Tax (× {taxMultiplier})
              </label>
              <Button variant="ghost" onClick={() => fillRest(i)}>
                Fill rest
              </Button>
              {items.length > 1 && (
                <Button
                  variant="ghost"
                  onClick={() =>
                    setItems((l) => l.filter((x) => x.key !== item.key))
                  }
                >
                  Remove
                </Button>
              )}
            </div>
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
      <p
        data-testid="remaining"
        className={`text-base font-semibold ${summary.remaining < -0.005 ? "text-red-600" : ""}`}
      >
        Remaining to allocate: {formatMoney(summary.remaining)}
      </p>
      <div className="flex gap-2">
        <Button variant="primary" type="submit" disabled={!summary.canSubmit}>
          Split
        </Button>
        <Button onClick={onCancel}>Back</Button>
      </div>
    </form>
  );
}
