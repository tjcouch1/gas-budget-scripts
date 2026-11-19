import { useState } from "react";
import type { PayPeriodStateDto } from "@budget/shared";
import {
  CORE_COLUMN_INDEXES,
  ENVELOPE_COLUMNS,
  REMAIN_INDEX,
} from "../lib/rows.ts";

export function EnvelopesView({
  state,
  warnings,
}: {
  state: PayPeriodStateDto;
  warnings: string[];
}) {
  const [showAll, setShowAll] = useState(false);
  const envelopes = state.envelopes;

  if (!envelopes) {
    return (
      <div className="rounded-xl bg-amber-50 p-4 text-sm text-amber-900 dark:bg-amber-900/20 dark:text-amber-100">
        <p className="font-medium">No envelope data for this pay period.</p>
        <p className="mt-1">
          This sheet is missing the optional EnvelopesStart / EnvelopesMax
          variables (see the setup checklist). Older tabs simply have no
          envelope balances.
        </p>
        {warnings.map((w) => (
          <p key={w} className="mt-1 opacity-80">
            {w}
          </p>
        ))}
      </div>
    );
  }

  return (
    <div>
      <label className="mb-3 flex min-h-11 items-center gap-2 text-sm">
        <input
          type="checkbox"
          checked={showAll}
          onChange={(e) => setShowAll(e.target.checked)}
        />
        Show all columns
      </label>
      {showAll ? (
        <div className="overflow-x-auto rounded-xl ring-1 ring-slate-200 dark:ring-slate-700">
          <table className="w-full min-w-max text-sm">
            <thead className="bg-slate-100 text-left dark:bg-slate-800">
              <tr>
                <th className="px-3 py-2">Envelope</th>
                {ENVELOPE_COLUMNS.map((c) => (
                  <th key={c} className="px-3 py-2 text-right">
                    {c}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {envelopes.map((e) => (
                <tr
                  key={e.name}
                  className="border-t border-slate-100 dark:border-slate-800"
                >
                  <td className="px-3 py-2 font-medium">{e.name}</td>
                  {ENVELOPE_COLUMNS.map((c, i) => (
                    <td key={c} className="px-3 py-2 text-right tabular-nums">
                      {e.display[i] ?? ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <ul className="space-y-2">
          {envelopes.map((e) => {
            const remain = e.values[REMAIN_INDEX];
            return (
              <li
                key={e.name}
                className="rounded-xl bg-white p-3 shadow-sm ring-1 ring-slate-200 dark:bg-slate-900 dark:ring-slate-700"
              >
                <div className="flex items-baseline justify-between gap-3">
                  <span className="font-medium">{e.name}</span>
                  <span
                    data-testid={`remain-${e.name}`}
                    className={`text-xl font-semibold tabular-nums ${
                      remain !== null && remain < 0
                        ? "text-red-600"
                        : "text-teal-700 dark:text-teal-400"
                    }`}
                  >
                    {e.display[REMAIN_INDEX]}
                  </span>
                </div>
                <dl className="mt-1 grid grid-cols-3 gap-2 text-xs text-slate-500 dark:text-slate-400">
                  {CORE_COLUMN_INDEXES.filter((i) => i !== REMAIN_INDEX).map(
                    (i) => (
                      <div key={i}>
                        <dt>{ENVELOPE_COLUMNS[i]}</dt>
                        <dd className="tabular-nums text-slate-700 dark:text-slate-300">
                          {e.display[i]}
                        </dd>
                      </div>
                    ),
                  )}
                </dl>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}
