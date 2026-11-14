import { useRef, useState } from "react";
import {
  evaluateMathExpression,
  isUnsafeFreeText,
  type NewTransactionFields,
} from "@budget/shared";
import { formatMoney, formatShortDate } from "../lib/format.ts";
import { ChipChooser } from "./ChipChooser.tsx";
import {
  Button,
  Dialog,
  UnsavedPrompt,
  inputClass,
  useUnsavedGuard,
} from "./ui.tsx";

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

/** The browser's local date as yyyy-MM-dd. */
export function todayLocalIso(now = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

interface Draft {
  date: string;
  name: string;
  cost: string;
  envelope: string;
  paidWith: string;
  taxDeductible: boolean;
  note: string;
}

/**
 * Drawer for a new transaction ("New expense") on the selected pay period. The sheet writes it
 * after the last row. A date outside the pay period is allowed after an explicit choice.
 */
export function NewExpenseSheet({
  period,
  envelopeOptions,
  paidWithOptions,
  today = todayLocalIso(),
  onSave,
  onClose,
}: {
  /** The selected pay period's first and last day (yyyy-MM-dd); null when unknown. */
  period: { startDate: string; endDate: string } | null;
  envelopeOptions: string[] | null;
  paidWithOptions: string[] | null;
  /** The default date (yyyy-MM-dd). */
  today?: string;
  onSave: (fields: NewTransactionFields) => void;
  onClose: () => void;
}) {
  const [initial] = useState<Draft>(() => ({
    date: today,
    name: "",
    cost: "",
    envelope: "",
    paidWith: "",
    taxDeductible: false,
    note: "",
  }));
  const [draft, setDraftState] = useState<Draft>(initial);
  const [askDate, setAskDate] = useState(false);
  const dateInput = useRef<HTMLInputElement>(null);
  const setDraft = (patch: Partial<Draft>) =>
    setDraftState((d) => ({ ...d, ...patch }));

  const expression = draft.cost.trim().replace(/^=/, "").trim();
  const value = expression ? evaluateMathExpression(expression) : null;
  const costInvalid = expression !== "" && value === null;
  const nameUnsafe = isUnsafeFreeText(draft.name);
  const dateValid = ISO_DATE.test(draft.date);
  const outsidePeriod =
    dateValid &&
    period !== null &&
    (draft.date < period.startDate || draft.date > period.endDate);
  const canSave =
    dateValid && draft.name.trim() !== "" && !nameUnsafe && value !== null;
  const dirty = (Object.keys(initial) as (keyof Draft)[]).some(
    (k) => draft[k] !== initial[k],
  );
  const guard = useUnsavedGuard({ dirty, onDiscard: onClose });

  const save = (confirmedDate = false) => {
    if (!canSave) return;
    if (outsidePeriod && !confirmedDate) {
      guard.setPrompting(false);
      setAskDate(true);
      return;
    }
    onSave({
      date: draft.date,
      name: draft.name,
      costExpression: expression,
      envelope: draft.envelope,
      paidWith: draft.paidWith,
      taxDeductible: draft.taxDeductible ? "Yes" : "",
      note: draft.note,
    });
    onClose();
  };

  return (
    <Dialog title="New expense" onClose={guard.requestClose}>
      <form
        className="space-y-3"
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <label className="block text-sm font-medium">
          Date
          <input
            ref={dateInput}
            type="date"
            className={`${inputClass} mt-1`}
            aria-invalid={!dateValid}
            value={draft.date}
            onChange={(e) => {
              setAskDate(false);
              setDraft({ date: e.target.value });
            }}
          />
        </label>
        {outsidePeriod && period && (
          <div
            role="alert"
            className="space-y-2 rounded-lg bg-amber-50 p-3 text-sm text-amber-900 dark:bg-amber-950/40 dark:text-amber-100"
          >
            <p>
              {formatShortDate(draft.date)} is outside this pay period (
              {formatShortDate(period.startDate)} –{" "}
              {formatShortDate(period.endDate)}). It will still be added to this
              pay period&apos;s sheet.
            </p>
            {askDate && (
              <div className="flex flex-wrap gap-2">
                <Button variant="primary" size="sm" onClick={() => save(true)}>
                  Save anyway
                </Button>
                <Button
                  size="sm"
                  onClick={() => {
                    setAskDate(false);
                    dateInput.current?.focus();
                  }}
                >
                  Change date
                </Button>
              </div>
            )}
          </div>
        )}
        <div>
          <input
            className={`${inputClass} ${nameUnsafe ? "border-red-500" : ""}`}
            aria-label="Name"
            aria-invalid={nameUnsafe}
            placeholder="Name"
            value={draft.name}
            onChange={(e) => setDraft({ name: e.target.value })}
          />
          {nameUnsafe && (
            <p className="mt-1 text-sm text-red-600">
              Names must not start with = + - or @.
            </p>
          )}
        </div>
        <div className="flex items-center gap-2">
          <input
            className={`${inputClass} ${costInvalid ? "border-red-500" : ""}`}
            aria-label="Cost"
            aria-invalid={costInvalid}
            inputMode="decimal"
            placeholder="Amount, e.g. 4.23*2+3"
            value={draft.cost}
            onChange={(e) => setDraft({ cost: e.target.value })}
          />
          <span className="w-24 shrink-0 text-right text-sm tabular-nums">
            {formatMoney(value)}
          </span>
        </div>
        {envelopeOptions && (
          <ChipChooser
            label="Envelope"
            options={envelopeOptions}
            value={draft.envelope}
            onSelect={(envelope) => setDraft({ envelope })}
          />
        )}
        {paidWithOptions && (
          <ChipChooser
            label="Paid with"
            options={paidWithOptions}
            value={draft.paidWith}
            onSelect={(paidWith) => setDraft({ paidWith })}
          />
        )}
        <label className="flex min-h-11 items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={draft.taxDeductible}
            onChange={(e) => setDraft({ taxDeductible: e.target.checked })}
          />
          Tax deductible
        </label>
        <textarea
          aria-label="Note (optional)"
          rows={2}
          placeholder="Note (optional)"
          className="w-full rounded-lg border border-pink-200 bg-pink-50 px-3 py-2 text-sm text-pink-950 outline-none focus:border-pink-400 focus:ring-2 focus:ring-pink-400/30 dark:border-pink-900 dark:bg-pink-950/30 dark:text-pink-100"
          value={draft.note}
          onChange={(e) => setDraft({ note: e.target.value })}
        />
        <Button
          type="submit"
          variant="primary"
          className="w-full"
          disabled={!canSave}
        >
          Save
        </Button>
      </form>
      {guard.prompting && (
        <UnsavedPrompt
          canSave={canSave}
          onSave={() => save()}
          onDiscard={onClose}
          onBack={() => guard.setPrompting(false)}
        />
      )}
    </Dialog>
  );
}
