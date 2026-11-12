import { useState } from "react";
import {
  evaluateMathExpression,
  isUnsafeFreeText,
  type EnvelopeRowDto,
  type SplitItemInput,
  type TransactionChanges,
} from "@budget/shared";
import { formatMoney } from "../lib/format.ts";
import {
  isGroupName,
  joinNameParts,
  splitNameParts,
} from "../lib/name-parts.ts";
import { envelopeRemain, type ViewRow } from "../lib/rows.ts";
import { SplitEditor } from "./SplitEditor.tsx";
import {
  Button,
  Dialog,
  inputClass,
  UnsavedPrompt,
  useUnsavedGuard,
} from "./ui.tsx";

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
  onSplit: (
    splits: SplitItemInput[],
    parent: { itemNote?: string; envelope?: string },
  ) => void;
  onClose: () => void;
}

/**
 * Bottom sheet for one transaction. The title is the editable name. Edits are
 * staged locally and saved together with "Save & close"; tapping an envelope
 * chip is the fast path: it saves the staged edits plus the envelope and closes.
 * Closing with unsaved edits asks Save / Discard / Go back.
 */
export function TransactionSheet(props: TransactionSheetProps) {
  const { row, onChange, onClose } = props;
  const [mode, setMode] = useState<"main" | "split">("main");
  const initialCost =
    row.costFormula ?? (row.cost === null ? "" : String(row.cost));
  const [cost, setCost] = useState(initialCost);
  const [preTax, setPreTax] = useState(false);
  const inGroup = isGroupName(row.name);
  const parts = splitNameParts(row.name);
  // In a split group only the item note after " | " is edited; the merchant prefix groups the rows.
  const initialName = inGroup ? parts.note : row.name;
  const [name, setName] = useState(initialName);
  const [paidWith, setPaidWith] = useState(row.paidWith);
  const wasYes = row.taxDeductible.trim().toLowerCase() === "yes";
  const [taxDeductible, setTaxDeductible] = useState(wasYes);

  const canEditCost = !props.isSplitParent && !row.splitPending;
  const canEditName = !row.splitPending;
  const canSplit = !props.isSplitChild && !row.splitPending;

  const costExpr = cost.trim().replace(/^=/, "");
  const costValue = evaluateMathExpression(costExpr);
  const costInvalid = canEditCost && cost.trim() !== "" && costValue === null;
  const nameUnsafe = canEditName && isUnsafeFreeText(name);

  const nameChanged = canEditName && name !== initialName;
  const costChanged =
    canEditCost &&
    cost.trim() !== "" &&
    (preTax || cost.trim() !== initialCost.trim());
  const paidWithChanged =
    props.paidWithOptions !== null && paidWith !== row.paidWith;
  const taxChanged = taxDeductible !== wasYes;
  const dirty = nameChanged || costChanged || paidWithChanged || taxChanged;
  const invalid = costInvalid || nameUnsafe;

  /** Every staged edit as one update. */
  function stagedChanges(): TransactionChanges {
    const changes: TransactionChanges = {};
    if (nameChanged)
      changes.name = inGroup ? joinNameParts(parts.prefix, name) : name;
    if (costChanged) {
      changes.costExpression = costExpr;
      if (preTax) changes.costPreTax = true;
    }
    if (paidWithChanged) changes.paidWith = paidWith;
    if (taxChanged) changes.taxDeductible = taxDeductible ? "Yes" : "";
    return changes;
  }

  function save(extra: TransactionChanges = {}) {
    if (invalid) return;
    const changes = { ...stagedChanges(), ...extra };
    if (Object.keys(changes).length > 0) onChange(changes);
    onClose();
  }

  const guard = useUnsavedGuard({ dirty, onDiscard: onClose });

  if (mode === "split") {
    return (
      <SplitEditor
        parentName={row.name}
        parentValue={row.cost}
        parentEnvelope={row.envelope}
        taxMultiplier={props.taxMultiplier}
        envelopeOptions={props.envelopeOptions}
        onSubmit={(splits, parent) => {
          props.onSplit(splits, parent);
          onClose();
        }}
        onClose={onClose}
      />
    );
  }

  const heading = (
    <div className="min-w-0 flex-1">
      {canEditName ? (
        <div className="flex min-w-0 items-baseline gap-1 border-b border-dashed border-slate-300 focus-within:border-teal-600 dark:border-slate-600">
          {inGroup && (
            <span
              data-testid="name-prefix"
              className="max-w-[45%] shrink-0 truncate text-lg font-semibold text-slate-500"
            >
              {parts.prefix} |
            </span>
          )}
          <input
            aria-label={inGroup ? "Item note" : "Name"}
            aria-invalid={nameUnsafe}
            placeholder={inGroup ? "item note" : "Name"}
            className="min-w-0 flex-1 bg-transparent py-1 text-lg font-semibold outline-none placeholder:font-normal placeholder:text-slate-400"
            value={name}
            onChange={(e) => setName(e.target.value)}
          />
          <span aria-hidden="true" className="text-slate-400">
            ✎
          </span>
        </div>
      ) : (
        <h2 className="text-lg font-semibold">{row.name || "Transaction"}</h2>
      )}
      {nameUnsafe && (
        <p role="alert" className="mt-1 text-xs text-red-600">
          Names cannot start with = + - or @.
        </p>
      )}
    </div>
  );

  return (
    <>
      <Dialog
        title="Transaction"
        heading={heading}
        onClose={guard.requestClose}
      >
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
              {props.envelopeOptions.map((option) => {
                const remain = envelopeRemain(props.envelopes, option);
                const selected = row.envelope === option;
                return (
                  <button
                    key={option}
                    type="button"
                    aria-pressed={selected}
                    disabled={row.splitPending || invalid}
                    onClick={() => save(selected ? {} : { envelope: option })}
                    className={`min-h-11 rounded-full px-4 text-sm ring-1 transition-colors disabled:opacity-50 ${
                      selected
                        ? "bg-teal-700 text-white ring-teal-700"
                        : "bg-white ring-slate-300 hover:bg-teal-50 dark:bg-slate-800 dark:ring-slate-600 dark:hover:bg-slate-700"
                    }`}
                  >
                    {option}
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
                  disabled={invalid}
                  onClick={() => save({ envelope: "" })}
                  className="min-h-11 rounded-full px-4 text-sm text-slate-500 ring-1 ring-dashed ring-slate-300 disabled:opacity-50"
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

        <div className="mb-4 flex flex-wrap items-center gap-2">
          {canSplit && (
            <Button
              variant="primary"
              onClick={() => setMode("split")}
              disabled={dirty}
            >
              Split
            </Button>
          )}
          {dirty && canSplit && (
            <span className="text-xs text-slate-500">
              Save your changes before splitting.
            </span>
          )}
          {row.splitPending && (
            <span className="text-sm text-slate-500">Split queued…</span>
          )}
        </div>

        {canEditCost ? (
          <div className="mb-4 space-y-2">
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
            </div>
            <label className="flex min-h-11 items-center gap-2 text-sm">
              <input
                type="checkbox"
                checked={preTax}
                onChange={(e) => setPreTax(e.target.checked)}
              />
              Add tax ×{props.taxMultiplier}
            </label>
          </div>
        ) : (
          props.isSplitParent && (
            <p className="mb-4 text-sm text-slate-500">
              This row is a split parent. Edit the child rows to change amounts.
            </p>
          )
        )}

        <details className="mb-4 rounded-xl border border-slate-200 p-3 dark:border-slate-700">
          <summary className="min-h-11 cursor-pointer py-2 text-sm font-medium">
            Edit details
          </summary>
          <div className="space-y-3 pt-2">
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
          </div>
        </details>

        <div className="sticky bottom-0 -mx-4 flex items-center gap-3 border-t border-slate-200 bg-white px-4 pt-3 dark:border-slate-700 dark:bg-slate-900">
          <Button
            variant="primary"
            className="flex-1"
            onClick={() => save()}
            disabled={invalid}
          >
            Save &amp; close
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
      </Dialog>
      {guard.prompting && (
        <UnsavedPrompt
          canSave={!invalid}
          onSave={() => save()}
          onDiscard={onClose}
          onBack={() => guard.setPrompting(false)}
        />
      )}
    </>
  );
}
