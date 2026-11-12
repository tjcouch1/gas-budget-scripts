import { useState } from "react";
import { isUnsafeFreeText, type SplitItemInput } from "@budget/shared";
import { formatMoney } from "../lib/format.ts";
import { splitNameParts } from "../lib/name-parts.ts";
import {
  summarizeSplit,
  toSplitInputs,
  type SplitDraftItem,
} from "../lib/split-math.ts";
import {
  Button,
  Dialog,
  inputClass,
  UnsavedPrompt,
  useUnsavedGuard,
} from "./ui.tsx";

let keyCounter = 0;
const newItem = (): SplitDraftItem => ({
  key: `item-${keyCounter++}`,
  itemNote: "",
  expression: "",
  preTax: false,
  envelope: "",
});

const cardClass =
  "relative space-y-2 rounded-xl border p-3 border-slate-200 dark:border-slate-700";

function EnvelopeSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}) {
  return (
    <select
      className={inputClass}
      aria-label={label}
      value={value}
      onChange={(e) => onChange(e.target.value)}
    >
      <option value="">Envelope…</option>
      {[...options, ...(value && !options.includes(value) ? [value] : [])].map(
        (o) => (
          <option key={o} value={o}>
            {o}
          </option>
        ),
      )}
    </select>
  );
}

/** The item note after " | ", with the merchant prefix shown as static text before it. */
function NoteField({
  prefix,
  label,
  value,
  onChange,
}: {
  prefix: string;
  label: string;
  value: string;
  onChange: (value: string) => void;
}) {
  return (
    <div className="flex items-center gap-2">
      <span
        data-testid="name-prefix"
        className="max-w-[40%] shrink-0 truncate text-sm text-slate-500"
      >
        {prefix} |
      </span>
      <input
        className={inputClass}
        aria-label={label}
        placeholder="Item note (optional)"
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  );
}

/**
 * A column of identical cards, like the spreadsheet. The first card is the
 * original row: its cost is calculated (original minus the other cards) and
 * read-only, but its item note and envelope are editable. Each card after it
 * is a new child row.
 */
export function SplitEditor({
  parentName,
  parentValue,
  parentEnvelope,
  taxMultiplier,
  envelopeOptions,
  onSubmit,
  onClose,
}: {
  parentName: string;
  /** The parent's current value; the parent keeps whatever is not allocated. */
  parentValue: number | null;
  parentEnvelope: string;
  taxMultiplier: number;
  envelopeOptions: string[] | null;
  /** `parent` holds only the original row's fields that were changed. */
  onSubmit: (
    splits: SplitItemInput[],
    parent: { itemNote?: string; envelope?: string },
  ) => void;
  onClose: () => void;
}) {
  const { prefix, note: initialParentNote } = splitNameParts(parentName);
  const [items, setItems] = useState<SplitDraftItem[]>(() => [newItem()]);
  const [parentNote, setParentNote] = useState(initialParentNote);
  const [parentEnv, setParentEnv] = useState(parentEnvelope);
  const summary = summarizeSplit(parentValue, items, taxMultiplier);
  const overAllocated = summary.remaining < -0.005;
  const parentNoteUnsafe = isUnsafeFreeText(parentNote);
  const canSubmit = summary.canSubmit && !parentNoteUnsafe;

  const dirty =
    parentNote !== initialParentNote ||
    parentEnv !== parentEnvelope ||
    items.length > 1 ||
    items.some(
      (i) => i.itemNote !== "" || i.expression !== "" || i.preTax || i.envelope,
    );

  const update = (key: string, patch: Partial<SplitDraftItem>) =>
    setItems((list) =>
      list.map((i) => (i.key === key ? { ...i, ...patch } : i)),
    );

  function save() {
    if (!canSubmit) return;
    onSubmit(toSplitInputs(items), {
      ...(parentNote !== initialParentNote
        ? { itemNote: parentNote.trim() }
        : {}),
      ...(parentEnv !== parentEnvelope ? { envelope: parentEnv } : {}),
    });
  }

  const guard = useUnsavedGuard({ dirty, onDiscard: onClose });

  return (
    <>
      <Dialog title="Split transaction" onClose={guard.requestClose}>
        <form
          className="space-y-3"
          onSubmit={(e) => {
            e.preventDefault();
            save();
          }}
        >
          <p className="text-sm text-slate-600 dark:text-slate-400">
            Each item below is taken out of the original. Amounts accept math
            like <code>4.23*2+12.09</code>.
          </p>

          <div
            data-testid="original-entry"
            className={`${cardClass} ${
              overAllocated
                ? "!border-red-400 bg-red-50 dark:bg-red-950/30"
                : "bg-slate-50 dark:bg-slate-800/50"
            }`}
          >
            <p className="text-xs font-medium uppercase tracking-wide text-slate-500">
              Remaining on original
            </p>
            <NoteField
              prefix={prefix}
              label="Original note"
              value={parentNote}
              onChange={setParentNote}
            />
            {parentNoteUnsafe && (
              <p className="text-xs text-red-600">
                Notes cannot start with = + - or @.
              </p>
            )}
            <input
              data-testid="remaining"
              className={`${inputClass} font-semibold tabular-nums disabled:opacity-80 ${
                overAllocated ? "text-red-600" : ""
              }`}
              aria-label="Original amount"
              value={formatMoney(summary.remaining)}
              readOnly
              disabled
            />
            <label className="flex min-h-11 items-center gap-2 text-sm opacity-60">
              <input type="checkbox" disabled checked={false} readOnly />
              <span>
                Add tax <span className="tabular-nums">×{taxMultiplier}</span>
                <span className="block text-xs text-slate-500">
                  Not used for the original: its amount is calculated
                </span>
              </span>
            </label>
            {envelopeOptions && (
              <EnvelopeSelect
                label="Original envelope"
                value={parentEnv}
                options={envelopeOptions}
                onChange={setParentEnv}
              />
            )}
            <p className="text-xs text-slate-500">
              Was {formatMoney(parentValue)}
            </p>
            {overAllocated && (
              <p role="alert" className="text-xs text-red-600">
                The items add up to more than the original amount. Reduce an
                item to continue.
              </p>
            )}
          </div>

          {items.map((item, i) => {
            const ev = summary.items[i];
            const invalid = item.expression.trim() !== "" && ev.value === null;
            return (
              <div
                key={item.key}
                data-testid="split-item"
                className={cardClass}
              >
                <div className="flex items-center justify-between">
                  <p className="text-xs font-medium uppercase tracking-wide text-slate-500">
                    Item {i + 1}
                  </p>
                  {items.length > 1 && (
                    <button
                      type="button"
                      aria-label={`Remove item ${i + 1}`}
                      title="Remove item"
                      onClick={() =>
                        setItems((l) => l.filter((x) => x.key !== item.key))
                      }
                      className="-mr-1 -mt-1 flex size-9 items-center justify-center rounded-lg text-xl leading-none text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                    >
                      ×
                    </button>
                  )}
                </div>
                <NoteField
                  prefix={prefix}
                  label={`Item ${i + 1} note`}
                  value={item.itemNote}
                  onChange={(v) => update(item.key, { itemNote: v })}
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
                    onChange={(e) =>
                      update(item.key, { preTax: e.target.checked })
                    }
                  />
                  <span>
                    Add tax{" "}
                    <span className="tabular-nums">×{taxMultiplier}</span>
                    <span className="block text-xs text-slate-500">
                      Multiplies this amount by the tax rate
                    </span>
                  </span>
                </label>
                {envelopeOptions && (
                  <EnvelopeSelect
                    label={`Item ${i + 1} envelope`}
                    value={item.envelope}
                    options={envelopeOptions}
                    onChange={(v) => update(item.key, { envelope: v })}
                  />
                )}
              </div>
            );
          })}
          <Button onClick={() => setItems((l) => [...l, newItem()])}>
            + Add item
          </Button>
          <div className="sticky bottom-0 -mx-4 flex items-center gap-3 border-t border-slate-200 bg-white px-4 pt-3 dark:border-slate-700 dark:bg-slate-900">
            <Button
              variant="primary"
              type="submit"
              className="flex-1"
              disabled={!canSubmit}
            >
              Save split
            </Button>
            {dirty && (
              <span
                data-testid="dirty-indicator"
                role="status"
                className="flex items-center gap-1 text-sm font-medium text-amber-700 dark:text-amber-400"
              >
                <span aria-hidden="true">●</span> Unsaved changes
              </span>
            )}
          </div>
        </form>
      </Dialog>
      {guard.prompting && (
        <UnsavedPrompt
          canSave={canSubmit}
          onSave={save}
          onDiscard={onClose}
          onBack={() => guard.setPrompting(false)}
        />
      )}
    </>
  );
}
