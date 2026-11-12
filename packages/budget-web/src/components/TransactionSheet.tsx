import { useState } from "react";
import {
  evaluateMathExpression,
  isUnsafeFreeText,
  type EnvelopeRowDto,
  type SplitItemInput,
  type TransactionChanges,
} from "@budget/shared";
import { formatMoney } from "../lib/format.ts";
import { envelopeRemain, type ViewRow } from "../lib/rows.ts";
import { SplitEditor } from "./SplitEditor.tsx";
import { Button, Dialog, inputClass } from "./ui.tsx";

export interface TransactionSheetProps {
  row: ViewRow;
  /** True when the row is the first of an existing split group. */
  isSplitParent: boolean;
  /** True when the row is a later member of a split group. */
  isSplitChild: boolean;
  taxMultiplier: number;
  /** Envelope choices, most used first; null when the sheet has no envelope list. */
  envelopeOptions: string[] | null;
  envelopes: EnvelopeRowDto[] | null;
  paidWithOptions: string[] | null;
  onChange: (changes: TransactionChanges) => void;
  onSplit: (splits: SplitItemInput[]) => void;
  onClose: () => void;
}

/**
 * Bottom sheet for one transaction. Fast path first: envelope chips (one tap
 * assigns and closes), Split, then Cost. Name, Paid with and Tax deductible live
 * under "Edit details".
 */
export function TransactionSheet(props: TransactionSheetProps) {
  const { row, onChange, onClose } = props;
  const [mode, setMode] = useState<"main" | "split">("main");
  const [cost, setCost] = useState(
    row.costFormula ?? (row.cost === null ? "" : String(row.cost)),
  );
  const [preTax, setPreTax] = useState(false);
  const [name, setName] = useState(row.name);
  const [paidWith, setPaidWith] = useState(row.paidWith);
  const [taxDeductible, setTaxDeductible] = useState(
    row.taxDeductible.trim().toLowerCase() === "yes",
  );

  const inGroup = row.splitGroupId !== null || row.name.includes("|");
  const canEditCost = !props.isSplitParent && !row.splitPending;
  const canEditName = !inGroup && !row.splitPending;
  const canSplit = !props.isSplitChild && !row.splitPending;

  const costExpr = cost.trim().replace(/^=/, "");
  const costValue = evaluateMathExpression(costExpr);
  const costInvalid = cost.trim() !== "" && costValue === null;
  const nameUnsafe = canEditName && isUnsafeFreeText(name);

  if (mode === "split") {
    return (
      <Dialog title="Split transaction" onClose={onClose}>
        <SplitEditor
          parentName={row.name}
          parentValue={row.cost}
          taxMultiplier={props.taxMultiplier}
          envelopeOptions={props.envelopeOptions}
          onSubmit={(splits) => {
            props.onSplit(splits);
            onClose();
          }}
          onCancel={() => setMode("main")}
        />
      </Dialog>
    );
  }

  const costChanged =
    canEditCost &&
    cost.trim() !== "" &&
    (preTax || cost.trim() !== (row.costFormula ?? String(row.cost ?? "")));

  function saveCost() {
    if (!costChanged || costInvalid) return;
    onChange({
      costExpression: costExpr,
      ...(preTax ? { costPreTax: true } : {}),
    });
    onClose();
  }

  function saveDetails() {
    const changes: TransactionChanges = {};
    if (canEditName && name !== row.name) changes.name = name;
    if (props.paidWithOptions && paidWith !== row.paidWith)
      changes.paidWith = paidWith;
    const wasYes = row.taxDeductible.trim().toLowerCase() === "yes";
    if (taxDeductible !== wasYes)
      changes.taxDeductible = taxDeductible ? "Yes" : "";
    if (Object.keys(changes).length > 0) onChange(changes);
    onClose();
  }

  return (
    <Dialog title={row.name || "Transaction"} onClose={onClose}>
      <p className="mb-3 text-sm text-slate-600 dark:text-slate-400">
        {row.date} · {row.costDisplay || formatMoney(row.cost)}
        {row.paidWith ? ` · ${row.paidWith}` : ""}
      </p>
      {row.note && (
        <p className="mb-3 rounded-lg bg-pink-100 p-2 text-sm text-pink-900 dark:bg-pink-900/30 dark:text-pink-100">
          {row.note}
        </p>
      )}

      {props.envelopeOptions ? (
        <section aria-label="Envelope" className="mb-4">
          <h3 className="mb-2 text-sm font-medium">Envelope</h3>
          <div className="flex flex-wrap gap-2">
            {props.envelopeOptions.map((name) => {
              const remain = envelopeRemain(props.envelopes, name);
              const selected = row.envelope === name;
              return (
                <button
                  key={name}
                  type="button"
                  aria-pressed={selected}
                  disabled={row.splitPending}
                  onClick={() => {
                    if (!selected) onChange({ envelope: name });
                    onClose();
                  }}
                  className={`min-h-11 rounded-full px-4 text-sm ring-1 transition-colors disabled:opacity-50 ${
                    selected
                      ? "bg-teal-700 text-white ring-teal-700"
                      : "bg-white ring-slate-300 hover:bg-teal-50 dark:bg-slate-800 dark:ring-slate-600 dark:hover:bg-slate-700"
                  }`}
                >
                  {name}
                  {remain && remain.display && (
                    <span
                      className={`ml-2 text-xs ${
                        remain.value !== null && remain.value < 0
                          ? selected
                            ? "text-red-200"
                            : "text-red-600"
                          : "opacity-70"
                      }`}
                    >
                      {remain.display}
                    </span>
                  )}
                </button>
              );
            })}
            {row.envelope && (
              <button
                type="button"
                onClick={() => {
                  onChange({ envelope: "" });
                  onClose();
                }}
                className="min-h-11 rounded-full px-4 text-sm text-slate-500 ring-1 ring-dashed ring-slate-300"
              >
                Clear
              </button>
            )}
          </div>
        </section>
      ) : (
        <p className="mb-4 text-sm text-slate-500">
          This sheet has no envelope list, so envelopes cannot be edited here.
        </p>
      )}

      <div className="mb-4 flex flex-wrap items-end gap-2">
        {canSplit && (
          <Button variant="primary" onClick={() => setMode("split")}>
            Split
          </Button>
        )}
        {row.splitPending && (
          <span className="text-sm text-slate-500">Split queued…</span>
        )}
      </div>

      {canEditCost ? (
        <form
          className="mb-4 space-y-2"
          onSubmit={(e) => {
            e.preventDefault();
            saveCost();
          }}
        >
          <label className="block text-sm font-medium" htmlFor="cost-input">
            Cost
          </label>
          <div className="flex items-center gap-2">
            <input
              id="cost-input"
              className={`${inputClass} ${costInvalid ? "border-red-500" : ""}`}
              inputMode="decimal"
              value={cost}
              onChange={(e) => setCost(e.target.value)}
              aria-invalid={costInvalid}
            />
            <span className="w-24 shrink-0 text-right text-sm tabular-nums">
              {costValue === null
                ? ""
                : formatMoney(
                    preTax ? costValue * props.taxMultiplier : costValue,
                  )}
            </span>
            <Button type="submit" disabled={!costChanged || costInvalid}>
              Save cost
            </Button>
          </div>
          <label className="flex min-h-11 items-center gap-2 text-sm">
            <input
              type="checkbox"
              checked={preTax}
              onChange={(e) => setPreTax(e.target.checked)}
            />
            Pre-Tax (× {props.taxMultiplier})
          </label>
        </form>
      ) : (
        props.isSplitParent && (
          <p className="mb-4 text-sm text-slate-500">
            This row is a split parent. Edit the child rows to change amounts.
          </p>
        )
      )}

      <details className="rounded-xl border border-slate-200 p-3 dark:border-slate-700">
        <summary className="min-h-11 cursor-pointer py-2 text-sm font-medium">
          Edit details
        </summary>
        <div className="space-y-3 pt-2">
          <label className="block text-sm font-medium">
            Name
            <input
              className={`${inputClass} mt-1`}
              value={name}
              disabled={!canEditName}
              onChange={(e) => setName(e.target.value)}
            />
          </label>
          {!canEditName && (
            <p className="text-xs text-slate-500">
              Names of split rows cannot be edited here.
            </p>
          )}
          {nameUnsafe && (
            <p className="text-xs text-red-600">
              Names cannot start with = + - or @.
            </p>
          )}
          {props.paidWithOptions && (
            <label className="block text-sm font-medium">
              Paid with
              <select
                className={`${inputClass} mt-1`}
                value={paidWith}
                onChange={(e) => setPaidWith(e.target.value)}
              >
                <option value=""></option>
                {[
                  ...props.paidWithOptions,
                  ...(paidWith && !props.paidWithOptions.includes(paidWith)
                    ? [paidWith]
                    : []),
                ].map((o) => (
                  <option key={o} value={o}>
                    {o}
                  </option>
                ))}
              </select>
            </label>
          )}
          <label className="flex min-h-11 items-center gap-2 text-sm">
            <input
              type="checkbox"
              checked={taxDeductible}
              onChange={(e) => setTaxDeductible(e.target.checked)}
            />
            Tax deductible
          </label>
          <Button variant="primary" onClick={saveDetails} disabled={nameUnsafe}>
            Save details
          </Button>
        </div>
      </details>
    </Dialog>
  );
}
