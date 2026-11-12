import type { TransactionRowDto } from "@budget/shared";
import { proposedRow } from "../sync/optimistic.ts";
import type { Conflict } from "../sync/types.ts";
import { Button, Dialog } from "./ui.tsx";

const FIELDS: { label: string; get: (r: TransactionRowDto) => string }[] = [
  { label: "Name", get: (r) => r.name },
  { label: "Cost", get: (r) => r.costFormula ?? r.costDisplay },
  { label: "Envelope", get: (r) => r.envelope },
  { label: "Paid with", get: (r) => r.paidWith },
  { label: "Tax deductible", get: (r) => r.taxDeductible },
];

function Column({
  title,
  row,
  other,
}: {
  title: string;
  row: TransactionRowDto | null;
  other: TransactionRowDto | null;
}) {
  return (
    <section
      aria-label={title}
      className="rounded-xl bg-slate-50 p-3 ring-1 ring-slate-200 dark:bg-slate-800 dark:ring-slate-700"
    >
      <h3 className="mb-2 text-sm font-semibold">{title}</h3>
      {row ? (
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
  const original = mutation.baseRow;
  const mine = proposedRow(mutation, original, taxMultiplier);
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
      <div className="grid gap-3 sm:grid-cols-3">
        <Column title="Original Value" row={original} other={null} />
        <Column title="Discarded Local Change" row={mine} other={original} />
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
