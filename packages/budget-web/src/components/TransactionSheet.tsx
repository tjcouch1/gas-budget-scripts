import { useLayoutEffect, useRef, useState, type ReactNode } from "react";
import {
  getSplitNamePrefix,
  isUnsafeFreeText,
  splitParentFormulaParts,
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
import {
  checkPlainAmount,
  evaluateRowDraft,
  initialRowDraft,
  type RowDraft,
  type RowDraftOptions,
} from "../lib/row-draft.ts";
import { type ViewRow } from "../lib/rows.ts";
import {
  newSplitItem,
  summarizeSplit,
  toSplitInputs,
  type SplitDraftItem,
} from "../lib/split-math.ts";
import { ChipChooser } from "./ChipChooser.tsx";
import { EnvelopeChipChooser } from "./EnvelopeChipChooser.tsx";
import {
  getDetailsOpen,
  getPaidWithExpanded,
  setDetailsOpen,
  setPaidWithExpanded,
} from "../lib/preferences.ts";
import {
  Button,
  ConfirmDialog,
  Dialog,
  inputClass,
  UnsavedPrompt,
  useUnsavedGuard,
} from "./ui.tsx";

export interface TransactionSheetProps {
  /**
   * The tapped row's split group in sheet order, or just the row when it is not
   * split. The first row is the original (the split parent). No synthetic rows.
   */
  rows: ViewRow[];
  /** rowIndex of the tapped row: its card is scrolled to the top of the drawer. */
  focusRowIndex?: number;
  taxMultiplier: number;
  /** Envelope choices, most used first; null when the sheet has no envelope list. */
  envelopeOptions: string[] | null;
  /** Envelope choices in the sheet's layout order (shown when the chooser is expanded). */
  envelopeLayoutOptions?: string[] | null;
  envelopes: EnvelopeRowDto[] | null;
  paidWithOptions: string[] | null;
  onUpdate: (row: ViewRow, changes: TransactionChanges) => void;
  /** Splits `row` (the original); `parent.itemNote` names it when it was not split yet. */
  onSplit: (
    row: ViewRow,
    splits: SplitItemInput[],
    parent: { itemNote?: string },
  ) => void;
  /** Deletes the transaction, or every row of its split group. Omitted: no Delete button. */
  onDelete?: (rows: ViewRow[]) => void;
  onClose: () => void;
}

const cardClass =
  "min-w-0 space-y-2 rounded-xl border border-slate-200 p-3 dark:border-slate-700";
const unsafeMessage = "Names cannot start with = + - or @.";

/** The card title, sitting in the top border like a fieldset legend. */
function CardLegend({ children }: { children: ReactNode }) {
  return (
    <legend className="px-1 text-xs font-medium uppercase tracking-wide text-slate-500">
      {children}
    </legend>
  );
}

/** The item note after " | ", with the merchant prefix shown as static text before it. */
function ItemNoteField({
  prefix,
  value,
  onChange,
  unsafe,
}: {
  prefix: string;
  value: string;
  onChange: (value: string) => void;
  unsafe: boolean;
}) {
  return (
    <div>
      <div className="flex items-center gap-2">
        <span className="max-w-[40%] shrink-0 truncate text-sm text-slate-500">
          {prefix} |
        </span>
        <input
          className={inputClass}
          aria-label="Item note"
          aria-invalid={unsafe}
          placeholder="Item note (optional)"
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
      </div>
      {unsafe && (
        <p role="alert" className="mt-1 text-xs text-red-600">
          {unsafeMessage}
        </p>
      )}
    </div>
  );
}

/** Cost math input with its evaluated value, and optionally the "Add tax" toggle. */
function AmountField({
  label,
  value,
  onChange,
  display,
  invalid,
  error,
  plain = false,
  preTax,
  onPreTax,
  taxMultiplier,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  display: number | null;
  invalid: boolean;
  /** Inline message under the field while it is invalid. */
  error?: string | null;
  /** A plain amount (a number only) rather than a math expression. */
  plain?: boolean;
  preTax?: boolean;
  onPreTax?: (preTax: boolean) => void;
  taxMultiplier: number;
}) {
  return (
    <div>
      <div className="flex items-center gap-2">
        <input
          className={`${inputClass} ${invalid ? "border-red-500" : ""}`}
          aria-label={label}
          aria-invalid={invalid}
          inputMode="decimal"
          placeholder={plain ? "Amount, e.g. 12.34" : "Amount, e.g. 4.23*2+3"}
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
        <span className="w-24 shrink-0 text-right text-sm tabular-nums">
          {formatMoney(display)}
        </span>
      </div>
      {invalid && error && (
        <p role="alert" className="mt-1 text-xs text-red-600">
          {error}
        </p>
      )}
      {onPreTax && (
        <label className="flex min-h-11 items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={preTax}
            onChange={(e) => onPreTax(e.target.checked)}
          />
          Add tax <span className="tabular-nums">×{taxMultiplier}</span>
        </label>
      )}
    </div>
  );
}

/** The fields under "Edit details" (Paid with, Tax deductible, Add note) and the note box. */
interface DetailsValues {
  paidWith: string;
  taxDeductible: boolean;
  note: string;
}

/**
 * The Name cell's note box and the collapsed "Edit details". `follow`: start in, and write
 * toggles to, the saved details/Paid with settings.
 */
function RowDetails({
  values,
  hasNote,
  onChange,
  paidWithOptions,
  follow,
}: {
  values: DetailsValues;
  /** True when the row had a note when the drawer opened. */
  hasNote: boolean;
  onChange: (patch: Partial<DetailsValues>) => void;
  paidWithOptions: string[] | null;
  follow: boolean;
}) {
  const [noteShown, setNoteShown] = useState(hasNote);
  const [detailsOpen, setDetailsOpenState] = useState(
    () => follow && getDetailsOpen(),
  );
  return (
    <>
      {noteShown && (
        <textarea
          aria-label="Note"
          // Mounted by "Add note": start typing straight away.
          autoFocus={!hasNote}
          rows={2}
          placeholder="Note"
          className="w-full rounded-lg border border-pink-200 bg-pink-50 px-3 py-2 text-sm text-pink-950 outline-none focus:border-pink-400 focus:ring-2 focus:ring-pink-400/30 dark:border-pink-900 dark:bg-pink-950/30 dark:text-pink-100"
          value={values.note}
          onChange={(e) => onChange({ note: e.target.value })}
        />
      )}
      <details
        open={detailsOpen}
        onToggle={(e) => {
          const open = e.currentTarget.open;
          if (open === detailsOpen) return;
          setDetailsOpenState(open);
          if (follow) setDetailsOpen(open);
        }}
        className="rounded-lg open:border open:border-slate-200 open:px-3 open:pb-2 dark:open:border-slate-700"
      >
        <summary className="cursor-pointer py-1.5 text-sm font-medium text-slate-600 dark:text-slate-300">
          Edit details
        </summary>
        {paidWithOptions && (
          <ChipChooser
            label="Paid with"
            options={paidWithOptions}
            value={values.paidWith}
            onSelect={(paidWith) => onChange({ paidWith })}
            defaultExpanded={follow && getPaidWithExpanded()}
            onExpandedChange={follow ? setPaidWithExpanded : undefined}
          />
        )}
        <div className="flex flex-wrap items-center gap-x-4">
          <label className="flex min-h-11 items-center gap-2 text-sm">
            <input
              type="checkbox"
              checked={values.taxDeductible}
              onChange={(e) => onChange({ taxDeductible: e.target.checked })}
            />
            Tax deductible
          </label>
          {!noteShown && (
            <Button variant="ghost" onClick={() => setNoteShown(true)}>
              Add note
            </Button>
          )}
        </div>
      </details>
    </>
  );
}

/**
 * One existing row: its name field, cost (or the split remainder), envelope,
 * paid with, the Name cell's note, and the collapsed "Edit details".
 */
function RowCard({
  label,
  rowIndex,
  nameField,
  costField,
  draft,
  hasNote,
  setDraft,
  envelopeChooser,
  paidWithOptions,
  bordered,
  follow,
}: {
  label: string;
  /** Split mode: a bordered card with its label as a legend on the border. */
  bordered: boolean;
  /** Start in, and write toggles to, the saved details/Paid with settings. */
  follow: boolean;
  rowIndex: number;
  nameField: ReactNode;
  costField: ReactNode;
  draft: RowDraft;
  /** True when the row had a note when the drawer opened. */
  hasNote: boolean;
  setDraft: (patch: Partial<RowDraft>) => void;
  envelopeChooser: ReactNode;
  paidWithOptions: string[] | null;
}) {
  return (
    <fieldset
      aria-label={label}
      data-row-index={rowIndex}
      className={bordered ? cardClass : "min-w-0 space-y-2"}
    >
      {bordered && <CardLegend>{label}</CardLegend>}
      {nameField}
      {costField}
      {envelopeChooser}
      <RowDetails
        values={draft}
        hasNote={hasNote}
        onChange={setDraft}
        paidWithOptions={paidWithOptions}
        follow={follow}
      />
    </fieldset>
  );
}

/**
 * Drawer for one transaction, or for a whole split group as a column of
 * cards like the sheet: the original row (the remainder) on top, one card per
 * split item below. Every edit, including an envelope tap, is staged and saved
 * together with "Save & close".
 */
export function TransactionSheet(props: TransactionSheetProps) {
  const { rows, taxMultiplier, onClose } = props;
  const parent = rows[0];
  const parentHasBar = isGroupName(parent.name);
  const prefixParts = splitNameParts(parent.name);
  const existingGroup = rows.length > 1;
  const splitPending = rows.some((r) => r.splitPending);

  const optionsFor = (i: number): RowDraftOptions => {
    const row = rows[i];
    const nameMode = row.splitPending
      ? "none"
      : isGroupName(row.name)
        ? "itemNote"
        : "name";
    const costMode =
      i > 0 ? "child" : existingGroup || splitPending ? "none" : "plain";
    return { nameMode, costMode, taxMultiplier };
  };
  const [initials] = useState(() =>
    rows.map((r, i) => initialRowDraft(r, optionsFor(i))),
  );
  const [drafts, setDrafts] = useState<RowDraft[]>(initials);
  const [items, setItems] = useState<SplitDraftItem[]>([]);
  /** Item note for an unsplit original that is being split. */
  const [parentNewNote, setParentNewNote] = useState("");
  const [showInfo, setShowInfo] = useState(false);
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  /** The split's name before " | " (the title of a split-named row), as edited. */
  const [titleDraft, setTitleDraft] = useState(prefixParts.prefix);
  /** The split total being edited (pencil), or null while it is not being edited. */
  const [totalDraft, setTotalDraft] = useState<string | null>(null);
  /**
   * Key of the new split item that is the focused card (the last one added), or null while the
   * focused card is an existing row. Only the focused card follows the saved settings.
   */
  const [focusedItemKey, setFocusedItemKey] = useState<string | null>(null);
  const body = useRef<HTMLDivElement>(null);

  const splitMode = existingGroup || splitPending || items.length > 0;
  const evals = rows.map((r, i) =>
    evaluateRowDraft(r, initials[i], drafts[i], optionsFor(i)),
  );
  const rowsTotal = rows.reduce((s, r) => s + (r.cost ?? 0), 0);
  /** The leading total of the parent's `=<total>-R[1]C[0]-...` formula, when it has one. */
  const totalParts =
    existingGroup && !splitPending
      ? splitParentFormulaParts(parent.costFormula)
      : null;
  const totalEditable = totalParts !== null && totalParts.total !== null;
  const totalCheck = totalDraft === null ? null : checkPlainAmount(totalDraft);
  const totalError = totalCheck?.error ?? null;
  /** The new total to save, when it was edited to a different valid amount. */
  const newTotal =
    totalCheck?.amount != null &&
    totalParts?.total != null &&
    Math.abs(Number(totalCheck.amount) - totalParts.total) > 1e-9
      ? totalCheck.amount
      : null;
  const total =
    existingGroup || splitPending
      ? totalCheck?.amount != null
        ? Number(totalCheck.amount)
        : rowsTotal
      : (evals[0].value ?? 0);
  const childrenSum = evals.slice(1).reduce((s, e) => s + (e.value ?? 0), 0);
  const summary = summarizeSplit(total - childrenSum, items, taxMultiplier);
  const overAllocated = splitMode && summary.remaining < -0.005;
  const itemsInvalid = summary.items.some(
    (e) => e.value === null || e.noteUnsafe,
  );
  const askParentNewNote = !existingGroup && !parentHasBar && items.length > 0;
  const parentNewNoteUnsafe =
    askParentNewNote && isUnsafeFreeText(parentNewNote);
  /** A split-named row's title (the text before " | ") is edited on its own. */
  const splitTitleEditable = optionsFor(0).nameMode === "itemNote";
  const title = titleDraft.trim();
  const titleChanged = splitTitleEditable && title !== prefixParts.prefix;
  const titleError = !splitTitleEditable
    ? null
    : title === ""
      ? "Enter a name"
      : isUnsafeFreeText(title)
        ? unsafeMessage
        : title.includes("|")
          ? "The name of a split cannot contain |"
          : null;
  /** The parent's changes: its card's edits, the split's new title, and the new total. */
  const parentChanges: TransactionChanges = {
    ...evals[0].changes,
    // The new prefix with the parent's item note; " | " is kept even when the note is blank.
    ...(titleChanged ? { name: joinNameParts(title, drafts[0].name) } : {}),
    ...(newTotal !== null ? { splitTotal: newTotal } : {}),
  };
  const invalid =
    evals.some((e) => e.costInvalid || e.nameUnsafe) ||
    itemsInvalid ||
    overAllocated ||
    parentNewNoteUnsafe ||
    titleError !== null ||
    totalError !== null;
  const dirty =
    Object.keys(parentChanges).length > 0 ||
    evals.slice(1).some((e) => Object.keys(e.changes).length > 0) ||
    items.length > 0;

  const setDraft = (i: number, patch: Partial<RowDraft>) =>
    setDrafts((list) => list.map((d, j) => (j === i ? { ...d, ...patch } : d)));
  const updateItem = (key: string, patch: Partial<SplitDraftItem>) =>
    setItems((list) =>
      list.map((i) => (i.key === key ? { ...i, ...patch } : i)),
    );

  function save() {
    if (invalid) return;
    // The children first: their names still carry the prefix the sheet has. Then the original,
    // whose rename moves every child to the new prefix, and last a split, so it sees the
    // original's new name and envelope.
    rows.forEach((row, i) => {
      if (i > 0 && Object.keys(evals[i].changes).length > 0)
        props.onUpdate(row, evals[i].changes);
    });
    if (Object.keys(parentChanges).length > 0)
      props.onUpdate(parent, parentChanges);
    if (items.length > 0)
      props.onSplit(
        parent,
        toSplitInputs(items),
        askParentNewNote && parentNewNote.trim()
          ? { itemNote: parentNewNote.trim() }
          : {},
      );
    onClose();
  }

  const guard = useUnsavedGuard({ dirty, onDiscard: onClose });

  // A queued split has no real rows yet for its new items, so the group cannot be deleted until it is saved.
  const canDelete =
    props.onDelete !== undefined &&
    !splitPending &&
    rows.every((r) => !r.synthetic && r.rowIndex >= 0);
  const deletesGroup = rows.length > 1;
  const deleteAmount = deletesGroup
    ? formatMoney(rows.reduce((s, r) => s + (r.cost ?? 0), 0))
    : parent.cost !== null
      ? formatMoney(parent.cost)
      : parent.costDisplay;
  function confirmDelete() {
    setConfirmingDelete(false);
    props.onDelete?.(rows);
    onClose();
  }

  /** Scrolls the drawer so the card matching `selector` sits near its top. */
  const scrollToCard = (selector: string) => {
    const card = body.current?.querySelector<HTMLElement>(selector);
    const dialog = card?.closest<HTMLElement>('[role="dialog"]');
    if (!card || !dialog || typeof card.getBoundingClientRect !== "function")
      return;
    const delta =
      card.getBoundingClientRect().top - dialog.getBoundingClientRect().top;
    // jsdom has no layout: every rect is zero.
    if (delta > 0) dialog.scrollTop += delta - 72;
  };

  // Opening a child row of a group scrolls its card near the top.
  useLayoutEffect(() => {
    if (
      props.focusRowIndex === undefined ||
      props.focusRowIndex === parent.rowIndex
    )
      return;
    scrollToCard(`[data-row-index="${props.focusRowIndex}"]`);
    // Only on mount.
  }, []);

  // "Add split item" scrolls the new (focused) card near the top and focuses its amount.
  useLayoutEffect(() => {
    if (focusedItemKey === null) return;
    const selector = `[data-item-key="${focusedItemKey}"]`;
    scrollToCard(selector);
    body.current
      ?.querySelector<HTMLInputElement>(
        `${selector} input[aria-label="Amount"]`,
      )
      ?.focus({ preventScroll: true });
  }, [focusedItemKey]);

  const parentName = drafts[0].name;
  const prefix = parentHasBar ? titleDraft : parentName;
  const titleEditable = optionsFor(0).nameMode === "name";
  const titleInput = (
    value: string,
    onChange: (value: string) => void,
    invalidTitle: boolean,
  ) => (
    <div className="flex min-w-0 items-baseline gap-1 border-b border-dashed border-slate-300 focus-within:border-teal-600 dark:border-slate-600">
      <input
        aria-label="Name"
        aria-invalid={invalidTitle}
        placeholder="Name"
        className="min-w-0 flex-1 bg-transparent py-1 text-lg font-semibold outline-none placeholder:font-normal placeholder:text-slate-400"
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
      <span aria-hidden="true" className="text-slate-400">
        ✎
      </span>
    </div>
  );
  const headingError = titleEditable
    ? evals[0].nameUnsafe
      ? unsafeMessage
      : null
    : titleError;
  const heading = (
    <div className="min-w-0 flex-1">
      {titleEditable ? (
        titleInput(
          parentName,
          (name) => setDraft(0, { name }),
          evals[0].nameUnsafe,
        )
      ) : splitTitleEditable ? (
        // The split's name without its " | " (the sheet keeps it; see AGENTS.md section 3.3).
        titleInput(titleDraft, setTitleDraft, titleError !== null)
      ) : parentHasBar ? (
        <h2
          data-testid="name-prefix"
          className="truncate text-lg font-semibold"
        >
          {prefixParts.prefix}
        </h2>
      ) : (
        <h2 className="text-lg font-semibold">
          {parent.name || "Transaction"}
        </h2>
      )}
      {headingError && (
        <p role="alert" className="mt-1 text-xs text-red-600">
          {headingError}
        </p>
      )}
    </div>
  );

  const envelopeChooserFor = (
    value: string,
    onSelect: (value: string) => void,
    disabled = false,
    follow = false,
  ) =>
    props.envelopeOptions && (
      <EnvelopeChipChooser
        value={value}
        options={props.envelopeOptions}
        layoutOptions={props.envelopeLayoutOptions}
        envelopes={props.envelopes}
        onSelect={onSelect}
        disabled={disabled}
        follow={follow}
      />
    );

  /** Name field of an existing row: the item note, or the whole name for an ungrouped name. */
  const nameFieldFor = (i: number) => {
    const mode = optionsFor(i).nameMode;
    const row = rows[i];
    const rowPrefix = splitNameParts(row.name).prefix;
    if (mode === "itemNote")
      return (
        <ItemNoteField
          // Rows of the split show the title as it is being edited.
          prefix={
            splitTitleEditable && rowPrefix === prefixParts.prefix
              ? titleDraft
              : rowPrefix
          }
          value={drafts[i].name}
          onChange={(name) => setDraft(i, { name })}
          unsafe={evals[i].nameUnsafe}
        />
      );
    if (mode === "name" && i > 0)
      return (
        <div>
          <input
            className={inputClass}
            aria-label="Name"
            aria-invalid={evals[i].nameUnsafe}
            value={drafts[i].name}
            onChange={(e) => setDraft(i, { name: e.target.value })}
          />
          {evals[i].nameUnsafe && (
            <p role="alert" className="mt-1 text-xs text-red-600">
              {unsafeMessage}
            </p>
          )}
        </div>
      );
    if (i === 0 && askParentNewNote)
      return (
        <ItemNoteField
          prefix={prefix}
          value={parentNewNote}
          onChange={setParentNewNote}
          unsafe={parentNewNoteUnsafe}
        />
      );
    return null;
  };

  /**
   * Only the focused card follows (and writes) the saved settings: the unsplit row or the clicked
   * card of a group, until "Add split item" moves the focus to the new item.
   */
  const followsPreference = (i: number) =>
    focusedItemKey === null &&
    (rows.length === 1 || rows[i].rowIndex === props.focusRowIndex);

  const remainder = (
    <div>
      <div className="flex items-center gap-2">
        <span className="text-sm text-slate-500">Amount</span>
        <output
          data-testid="remaining"
          aria-label="Remainder amount"
          className={`text-lg font-semibold tabular-nums ${overAllocated ? "text-red-600" : ""}`}
        >
          {formatMoney(summary.remaining)}
        </output>
        <button
          type="button"
          aria-label="About this amount"
          aria-expanded={showInfo}
          onClick={() => setShowInfo((s) => !s)}
          className="flex min-h-11 min-w-11 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
        >
          ⓘ
        </button>
      </div>
      {showInfo && (
        <p className="text-xs text-slate-500">
          Calculated as the remainder from the split items below.
        </p>
      )}
      {overAllocated && (
        <p role="alert" className="text-xs text-red-600">
          The split items add up to more than the total. Reduce an item to
          continue.
        </p>
      )}
    </div>
  );

  const childCount = rows.length - 1;

  return (
    <>
      <Dialog
        title="Transaction"
        heading={heading}
        onClose={guard.requestClose}
      >
        <div ref={body}>
          {totalDraft !== null ? (
            <div className="mb-3">
              <div className="flex items-center gap-2 text-sm text-slate-600 dark:text-slate-400">
                <span>{parent.date} · Total</span>
                <input
                  className={`${inputClass} max-w-40 ${totalError ? "border-red-500" : ""}`}
                  aria-label="Total"
                  aria-invalid={totalError !== null}
                  inputMode="decimal"
                  autoFocus
                  placeholder="Total, e.g. 45.67"
                  value={totalDraft}
                  onChange={(e) => setTotalDraft(e.target.value)}
                />
                {totalCheck?.amount != null && (
                  <span className="tabular-nums">{formatMoney(total)}</span>
                )}
              </div>
              {totalError && (
                <p role="alert" className="mt-1 text-xs text-red-600">
                  {totalError}
                </p>
              )}
            </div>
          ) : (
            <p className="mb-3 flex items-center gap-1 text-sm text-slate-600 dark:text-slate-400">
              <span>
                {parent.date} ·{" "}
                {splitMode
                  ? `Total ${formatMoney(total)}`
                  : parent.costDisplay || formatMoney(parent.cost)}
              </span>
              {totalEditable && (
                <button
                  type="button"
                  aria-label="Edit total"
                  title="Edit the total before splitting"
                  onClick={() => setTotalDraft(String(totalParts?.total ?? ""))}
                  className="flex min-h-11 min-w-11 items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
                >
                  ✎
                </button>
              )}
            </p>
          )}
          {!props.envelopeOptions && (
            <p className="mb-3 text-sm text-slate-500">
              This sheet has no envelope list, so envelopes cannot be edited
              here.
            </p>
          )}

          <div className="space-y-3">
            {rows.map((row, i) => (
              <RowCard
                key={row.key}
                label={
                  i === 0
                    ? splitMode
                      ? "Remainder"
                      : "Transaction"
                    : `Split item ${i}`
                }
                rowIndex={row.rowIndex}
                bordered={splitMode}
                follow={followsPreference(i)}
                nameField={nameFieldFor(i)}
                costField={
                  i === 0 ? (
                    splitMode ? (
                      remainder
                    ) : (
                      <AmountField
                        label="Cost"
                        value={drafts[0].cost}
                        onChange={(cost) => setDraft(0, { cost })}
                        display={evals[0].value}
                        invalid={evals[0].costInvalid}
                        error={evals[0].costError}
                        plain
                        taxMultiplier={taxMultiplier}
                      />
                    )
                  ) : (
                    <AmountField
                      label="Amount"
                      value={drafts[i].cost}
                      onChange={(cost) => setDraft(i, { cost })}
                      display={evals[i].value}
                      invalid={evals[i].costInvalid}
                      error={evals[i].costError}
                      preTax={drafts[i].preTax}
                      onPreTax={(preTax) => setDraft(i, { preTax })}
                      taxMultiplier={taxMultiplier}
                    />
                  )
                }
                draft={drafts[i]}
                hasNote={row.note !== ""}
                setDraft={(patch) => setDraft(i, patch)}
                envelopeChooser={envelopeChooserFor(
                  drafts[i].envelope,
                  (envelope) => setDraft(i, { envelope }),
                  false,
                  followsPreference(i),
                )}
                paidWithOptions={props.paidWithOptions}
              />
            ))}

            {items.map((item, k) => {
              const n = childCount + k + 1;
              const ev = summary.items[k];
              return (
                <fieldset
                  key={item.key}
                  aria-label={`Split item ${n}`}
                  data-testid="split-item"
                  data-item-key={item.key}
                  className={`${cardClass} relative border-dashed pt-4`}
                >
                  <CardLegend>Split item {n} · new</CardLegend>
                  <button
                    type="button"
                    aria-label={`Remove split item ${n}`}
                    title="Remove item"
                    onClick={() =>
                      setItems((l) => l.filter((x) => x.key !== item.key))
                    }
                    className="absolute -top-4 right-2 flex h-8 w-8 items-center justify-center rounded-full bg-white text-xl leading-none text-slate-500 hover:bg-slate-100 dark:bg-slate-900 dark:hover:bg-slate-800"
                  >
                    ×
                  </button>
                  <ItemNoteField
                    prefix={prefix}
                    value={item.itemNote}
                    onChange={(itemNote) => updateItem(item.key, { itemNote })}
                    unsafe={ev.noteUnsafe}
                  />
                  <AmountField
                    label="Amount"
                    value={item.expression}
                    onChange={(expression) =>
                      updateItem(item.key, { expression })
                    }
                    display={ev.value}
                    invalid={item.expression.trim() !== "" && ev.value === null}
                    preTax={item.preTax}
                    onPreTax={(preTax) => updateItem(item.key, { preTax })}
                    taxMultiplier={taxMultiplier}
                  />
                  {envelopeChooserFor(
                    item.envelope,
                    (envelope) => updateItem(item.key, { envelope }),
                    false,
                    item.key === focusedItemKey,
                  )}
                  <RowDetails
                    values={item}
                    hasNote={false}
                    onChange={(patch) => updateItem(item.key, patch)}
                    paidWithOptions={props.paidWithOptions}
                    follow={item.key === focusedItemKey}
                  />
                </fieldset>
              );
            })}
          </div>

          <div className="my-4 flex items-center justify-between gap-2">
            {splitPending ? (
              <span className="text-sm text-slate-500">Split queued…</span>
            ) : (
              <Button
                onClick={() => {
                  const item = newSplitItem({
                    paidWith: drafts[0].paidWith,
                    taxDeductible: drafts[0].taxDeductible,
                  });
                  setItems((l) => [...l, item]);
                  setFocusedItemKey(item.key);
                }}
              >
                <span aria-hidden="true">+ </span>Add split item
              </Button>
            )}
            {props.onDelete && (
              <Button
                variant="danger"
                onClick={() => setConfirmingDelete(true)}
                disabled={!canDelete}
              >
                {deletesGroup ? "Delete split" : "Delete transaction"}
              </Button>
            )}
          </div>
          {props.onDelete && splitPending && (
            <p className="-mt-2 mb-4 text-xs text-slate-500">
              Can be deleted once the queued split is saved.
            </p>
          )}

          <div className="sticky bottom-0 -mx-4 flex items-center gap-3 border-t border-slate-200 bg-white px-4 pt-3 dark:border-slate-700 dark:bg-slate-900">
            <Button
              variant="primary"
              className="flex-1"
              onClick={() => save()}
              disabled={invalid || !dirty}
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
        </div>
      </Dialog>
      {confirmingDelete && (
        <ConfirmDialog
          title={deletesGroup ? "Delete split?" : "Delete transaction?"}
          confirmLabel="Delete"
          onConfirm={confirmDelete}
          onCancel={() => setConfirmingDelete(false)}
        >
          <p className="font-medium text-slate-900 dark:text-slate-100">
            {deletesGroup ? getSplitNamePrefix(parent.name) : parent.name} ·{" "}
            {deleteAmount}
          </p>
          <p className="mt-1">
            {deletesGroup
              ? `This deletes all ${rows.length} rows of this split from the sheet.`
              : "This deletes the row from the sheet."}{" "}
            Unsaved edits in this drawer are discarded.
          </p>
        </ConfirmDialog>
      )}
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
