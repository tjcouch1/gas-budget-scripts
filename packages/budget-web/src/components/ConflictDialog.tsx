import type { TransactionRowDto } from "@budget/shared";
import { groupResend } from "../sync/engine.ts";
import { proposedGroupRows, proposedRow } from "../sync/optimistic.ts";
import type { Conflict, GroupMutation } from "../sync/types.ts";
import { Button, Dialog } from "./ui.tsx";

const FIELDS: { label: string; get: (r: TransactionRowDto) => string }[] = [
  { label: "Name", get: (r) => r.name },
  {
    label: "Cost",
    // The plain number, never a formatted "$1,234.50", so all three columns compare alike.
    get: (r) =>
      r.costFormula ?? (r.cost !== null ? String(r.cost) : r.costDisplay),
  },
  { label: "Envelope", get: (r) => r.envelope },
  { label: "Paid with", get: (r) => r.paidWith },
  { label: "Tax deductible", get: (r) => r.taxDeductible },
  { label: "Note", get: (r) => r.note },
];

function Column({
  title,
  row,
  other,
  replacement,
}: {
  title: string;
  row: TransactionRowDto | null;
  other: TransactionRowDto | null;
  /** Shown instead of the fields (e.g. for a deleted row). */
  replacement?: string;
}) {
  return (
    <section
      aria-label={title}
      className="rounded-xl bg-slate-50 p-3 ring-1 ring-slate-200 dark:bg-slate-800 dark:ring-slate-700"
    >
      <h3 className="mb-2 text-sm font-semibold">{title}</h3>
      {replacement ? (
        <p className="text-sm font-semibold text-amber-700 dark:text-amber-300">
          {replacement}
        </p>
      ) : row ? (
        <dl className="space-y-1 text-sm">
          {FIELDS.map((f) => {
            const value = f.get(row);
            const differs = other !== null && f.get(other) !== value;
            return (
              <div key={f.label}>
                <dt className="text-xs text-slate-500">{f.label}</dt>
                <dd
                  className={
                    differs
                      ? "font-semibold text-amber-700 dark:text-amber-300"
                      : ""
                  }
                >
                  {value || "(blank)"}
                </dd>
              </div>
            );
          })}
        </dl>
      ) : (
        <p className="text-sm text-slate-500">This row no longer exists.</p>
      )}
    </section>
  );
}

/** The group table's fields: like `FIELDS`, but costs as values (a remainder's formula hides its value). */
const GROUP_FIELDS: { label: string; get: (r: TransactionRowDto) => string }[] =
  FIELDS.map((f) =>
    f.label === "Cost"
      ? {
          label: "Cost",
          get: (r) =>
            r.cost !== null
              ? String(Number(r.cost.toPrecision(12)))
              : (r.costFormula ?? r.costDisplay),
        }
      : f,
  );

const VERSIONS = [
  { key: "original", label: "Original" },
  { key: "mine", label: "Discarded local change" },
  { key: "remote", label: "Current remote" },
] as const;

/**
 * 3-way conflict of a split group save: one table per row of the group (Original / Discarded
 * local change / Current remote), values that differ from the original highlighted, and one
 * choice for the whole group.
 */
function GroupConflictDialog({
  conflict,
  mutation,
  taxMultiplier,
  onResolve,
}: {
  conflict: Conflict;
  mutation: GroupMutation;
  taxMultiplier: number;
  onResolve: (choice: "accept" | "overwrite") => void;
}) {
  const currentRows =
    conflict.currentRows ?? (conflict.current ? [conflict.current] : []);
  const proposed = proposedGroupRows(mutation, taxMultiplier);
  const indexes = [
    ...new Set([
      ...mutation.payload.rows.map((r) => r.rowIndex),
      ...currentRows.map((r) => r.rowIndex),
    ]),
  ].sort((a, b) => a - b);
  const canOverwrite = groupResend(mutation, currentRows) !== null;
  const items = mutation.payload.newItems ?? [];
  return (
    <Dialog title="This split changed in the sheet" wide>
      <p className="mb-3 text-sm text-slate-600 dark:text-slate-400">
        Someone edited this split after you last synced. Your changes to it were
        not saved. Choose which version of the whole split to keep.
      </p>
      {items.length > 0 && (
        <p className="mb-3 text-sm">
          Your discarded change also added {items.length} new item
          {items.length === 1 ? "" : "s"}:{" "}
          {items
            .map((s) => `${s.itemNote || "(no note)"} ${s.expression}`)
            .join("; ")}
        </p>
      )}
      <div className="space-y-3">
        {indexes.map((rowIndex, position) => {
          const i = mutation.payload.rows.findIndex(
            (r) => r.rowIndex === rowIndex,
          );
          const original = i >= 0 ? mutation.payload.rows[i].baseRow : null;
          const versions = {
            original,
            mine: i >= 0 ? proposed[i] : null,
            remote: currentRows.find((r) => r.rowIndex === rowIndex) ?? null,
          };
          const label = position === 0 ? "Remainder" : `Split item ${position}`;
          return (
            <section
              key={rowIndex}
              aria-label={label}
              className="rounded-xl bg-slate-50 p-3 ring-1 ring-slate-200 dark:bg-slate-800 dark:ring-slate-700"
            >
              <h3 className="mb-2 text-sm font-semibold">{label}</h3>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[36rem] text-left text-sm">
                  <thead>
                    <tr className="text-xs text-slate-500">
                      <th scope="col" className="pr-2 font-normal">
                        <span className="sr-only">Version</span>
                      </th>
                      {GROUP_FIELDS.map((f) => (
                        <th
                          key={f.label}
                          scope="col"
                          className="pr-2 font-normal"
                        >
                          {f.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {VERSIONS.map((v) => {
                      const row = versions[v.key];
                      return (
                        <tr key={v.key} className="align-top">
                          <th
                            scope="row"
                            className="pr-2 text-xs font-medium text-slate-500"
                          >
                            {v.label}
                          </th>
                          {row ? (
                            GROUP_FIELDS.map((f) => {
                              const value = f.get(row);
                              const differs =
                                v.key !== "original" &&
                                original !== null &&
                                f.get(original) !== value;
                              return (
                                <td
                                  key={f.label}
                                  data-changed={differs ? "true" : undefined}
                                  className={`break-words pr-2 ${differs ? "font-semibold text-amber-700 dark:text-amber-300" : ""}`}
                                >
                                  {value || "(blank)"}
                                </td>
                              );
                            })
                          ) : (
                            <td
                              colSpan={GROUP_FIELDS.length}
                              className="text-slate-500"
                            >
                              {v.key === "remote"
                                ? "This row is no longer in the sheet."
                                : "Not in the split when you edited it."}
                            </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>
          );
        })}
      </div>
      <div className="mt-4 flex flex-wrap gap-2">
        <Button variant="primary" onClick={() => onResolve("accept")}>
          Accept Remote
        </Button>
        <Button
          variant="danger"
          disabled={!canOverwrite}
          onClick={() => onResolve("overwrite")}
        >
          Overwrite with Mine
        </Button>
      </div>
    </Dialog>
  );
}

/** 3-way conflict: Original / Discarded Local Change / Current Remote. */
export function ConflictDialog({
  conflict,
  taxMultiplier,
  onResolve,
}: {
  conflict: Conflict;
  taxMultiplier: number;
  onResolve: (choice: "accept" | "overwrite") => void;
}) {
  const { mutation, current } = conflict;
  if (mutation.action === "updateSplitGroup")
    return (
      <GroupConflictDialog
        conflict={conflict}
        mutation={mutation}
        taxMultiplier={taxMultiplier}
        onResolve={onResolve}
      />
    );
  const isDelete = mutation.action === "deleteTransaction";
  // For a delete, compare the row that changed (it may be any row of the split group).
  const original = isDelete
    ? (mutation.payload.rows.find((r) => r.rowIndex === current?.rowIndex)
        ?.baseRow ?? mutation.baseRow)
    : mutation.baseRow;
  const mine = proposedRow(mutation, original, taxMultiplier);
  const deleteCount = isDelete ? mutation.payload.rows.length : 0;
  return (
    <Dialog title="This row changed in the sheet" wide>
      <p className="mb-3 text-sm text-slate-600 dark:text-slate-400">
        Someone edited this transaction after you last synced. Choose which
        value to keep.
      </p>
      {mutation.action === "splitTransaction" && (
        <p className="mb-3 text-sm">
          Your discarded change was a split into{" "}
          {mutation.payload.splits.length} item
          {mutation.payload.splits.length === 1 ? "" : "s"}:{" "}
          {mutation.payload.splits
            .map((s) => `${s.itemNote || "(no note)"} ${s.expression}`)
            .join("; ")}
        </p>
      )}
      {isDelete && (
        <p className="mb-3 text-sm">
          Your discarded change was a delete of{" "}
          {deleteCount === 1
            ? "this transaction"
            : `all ${deleteCount} rows of this split`}
          . Accept Remote keeps the row(s); Overwrite with Mine deletes them
          anyway.
        </p>
      )}
      <div className="grid gap-3 sm:grid-cols-3">
        <Column title="Original Value" row={original} other={null} />
        <Column
          title="Discarded Local Change"
          row={mine}
          other={original}
          replacement={
            isDelete
              ? deleteCount === 1
                ? "Delete"
                : `Delete (all ${deleteCount} rows)`
              : undefined
          }
        />
        <Column title="Current Remote Value" row={current} other={original} />
      </div>
      <div className="mt-4 flex flex-wrap gap-2">
        <Button variant="primary" onClick={() => onResolve("accept")}>
          Accept Remote
        </Button>
        <Button
          variant="danger"
          disabled={current === null}
          onClick={() => onResolve("overwrite")}
        >
          Overwrite with Mine
        </Button>
      </div>
    </Dialog>
  );
}
