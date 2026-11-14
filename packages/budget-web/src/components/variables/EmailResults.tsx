import type { RuleCaptures } from "@budget/shared";
import { formatMoney } from "../../lib/format.ts";
import {
  formatOutcome,
  type PreviewItem,
} from "../../lib/receipt-rules-preview.ts";

const TONES = {
  ok: "bg-teal-100 text-teal-900 dark:bg-teal-900/60 dark:text-teal-100",
  muted: "bg-slate-200 text-slate-800 dark:bg-slate-700 dark:text-slate-100",
  error: "bg-red-100 text-red-900 dark:bg-red-900/60 dark:text-red-100",
};

const preClass =
  "mt-1 max-h-80 overflow-auto rounded-lg bg-slate-100 p-2 text-xs whitespace-pre-wrap break-words dark:bg-slate-800";

function captureEntries(
  prefix: string,
  groups: RuleCaptures | null,
): [string, string | undefined][] {
  return groups
    ? Object.entries(groups).map(([key, value]) => [`${prefix}.${key}`, value])
    : [];
}

function formatDate(iso: string): string {
  const date = new Date(iso);
  return Number.isNaN(date.getTime()) ? iso : date.toLocaleString();
}

/** Preview of the rules over fetched emails. All email text is rendered as React text (escaped). */
export function EmailResults({ items }: { items: PreviewItem[] }) {
  if (items.length === 0)
    return (
      <p className="text-sm text-slate-500 dark:text-slate-400">
        Fetch emails to test the rules.
      </p>
    );
  return (
    <ul aria-label="Email results" className="space-y-3">
      {items.map((item, i) => (
        <li key={`${item.email.messageId}-${i}`}>
          <ResultCard item={item} />
        </li>
      ))}
    </ul>
  );
}

function ResultCard({ item }: { item: PreviewItem }) {
  const { email, outcome, error } = item;
  const { label, tone, detail } = formatOutcome(item);
  const captures = outcome
    ? [
        ...captureEntries("subject", outcome.captures.subject),
        ...captureEntries("body", outcome.captures.body),
      ]
    : [];
  const receipt = outcome?.receipt;
  return (
    <article
      aria-label={email.subject}
      className="space-y-2 rounded-xl bg-white p-3 text-sm ring-1 ring-slate-200 dark:bg-slate-800/60 dark:ring-slate-700"
    >
      <div className="flex flex-wrap items-start justify-between gap-2">
        <h4 className="font-medium break-words">{email.subject}</h4>
        <span
          data-testid="outcome-badge"
          className={`shrink-0 rounded-full px-2 py-0.5 text-xs font-semibold ${TONES[tone]}`}
        >
          {detail ? `${label}: ${detail}` : label}
        </span>
      </div>
      <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-0.5 text-xs text-slate-600 dark:text-slate-400">
        <dt>From</dt>
        <dd className="break-all">{email.from}</dd>
        <dt>To</dt>
        <dd className="break-all">{email.to}</dd>
        <dt>Date</dt>
        <dd>{formatDate(email.date)}</dd>
        {email.labels.length > 0 && (
          <>
            <dt>Labels</dt>
            <dd>{email.labels.join(", ")}</dd>
          </>
        )}
      </dl>

      {error && (
        <p className="text-red-700 dark:text-red-300 break-words">{error}</p>
      )}

      {outcome && (
        <p className="text-xs">
          Rule:{" "}
          {outcome.ruleId ? <code>{outcome.ruleId}</code> : <span>none</span>}
          {outcome.senderAddress && (
            <>
              {" "}
              (sender <code className="break-all">{outcome.senderAddress}</code>
              {outcome.forwarded ? ", forwarded" : ""})
            </>
          )}
        </p>
      )}

      {receipt && outcome.isReceipt && (
        <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-0.5">
          <dt className="text-slate-500 dark:text-slate-400">Name</dt>
          <dd data-testid="receipt-name">{receipt.name ?? ""}</dd>
          <dt className="text-slate-500 dark:text-slate-400">Cost</dt>
          <dd data-testid="receipt-cost">
            {receipt.cost === undefined ? "" : formatMoney(receipt.cost)}
          </dd>
          <dt className="text-slate-500 dark:text-slate-400">Paid with</dt>
          <dd>{receipt.type ?? ""}</dd>
          <dt className="text-slate-500 dark:text-slate-400">Envelope</dt>
          <dd>{receipt.category ?? ""}</dd>
        </dl>
      )}

      {captures.length > 0 && (
        <dl
          role="list"
          aria-label="Captures"
          className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-0.5 rounded-lg bg-slate-50 p-2 text-xs dark:bg-slate-900"
        >
          {captures.map(([key, value]) => (
            <div key={key} role="listitem" className="contents">
              <dt className="font-mono">{key}</dt>
              <dd className="font-mono break-all">
                {value === undefined ? "(none)" : value}
              </dd>
            </div>
          ))}
        </dl>
      )}

      {outcome && outcome.warnings.length > 0 && (
        <ul className="list-disc pl-5 text-xs text-amber-700 dark:text-amber-300">
          {outcome.warnings.map((warning, i) => (
            <li key={i}>{warning}</li>
          ))}
        </ul>
      )}

      {(email.plainBodyTruncated || email.htmlBodyTruncated) && (
        <p className="text-xs text-amber-700 dark:text-amber-300">
          Body truncated by the server; the import sees the full text, so this
          result may differ.
        </p>
      )}

      <details>
        <summary className="flex min-h-11 cursor-pointer items-center">
          Plain body
          {email.plainBodyTruncated && " (truncated by the server)"}
        </summary>
        <pre className={preClass}>{email.plainBody}</pre>
      </details>
      {email.htmlBody !== undefined && (
        <details>
          <summary className="flex min-h-11 cursor-pointer items-center">
            HTML body (source)
            {email.htmlBodyTruncated && " (truncated by the server)"}
          </summary>
          <pre className={preClass}>{email.htmlBody}</pre>
        </details>
      )}
    </article>
  );
}
