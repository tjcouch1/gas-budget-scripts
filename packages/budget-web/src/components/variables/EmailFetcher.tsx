import { useState } from "react";
import {
  buildReceiptSearchQuery,
  isSafeSubjectFilter,
  isStrictEmailAddress,
  MAX_QUERY_SENDERS,
  type FetchReceiptEmailsRequest,
  type FetchReceiptEmailsResultDto,
  type ReceiptEmailDto,
  type TransactionSheetRangeDto,
  type RuleIssue,
} from "@budget/shared";
import type { CallApi } from "../../api/client.ts";
import { Button } from "../ui.tsx";
import { CheckboxField, TextField } from "./fields.tsx";

const MAX_SUBJECTS = 5;

const splitList = (text: string) =>
  text
    .split(",")
    .map((part) => part.trim())
    .filter((part) => part.length > 0);

const issue = (message: string): RuleIssue[] => [{ path: "", message }];

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

/** Read-only Gmail search for testing the rules. Never labels or archives. */
export function EmailFetcher({
  senders,
  forwardEmailAddress,
  forwardSubjectPrefixes,
  onFetched,
  api,
}: {
  /** Sender addresses from the candidate rules. */
  senders: string[];
  forwardEmailAddress: string;
  /** The candidate rules' forwarding.subjectPrefixes (undefined = built-in default). */
  forwardSubjectPrefixes?: string[];
  onFetched: (
    emails: ReceiptEmailDto[],
    query: string,
    transactionSheets?: TransactionSheetRangeDto[],
  ) => void;
  api: CallApi;
}) {
  const [extraText, setExtraText] = useState("");
  const [subjectText, setSubjectText] = useState("");
  const [includeProcessed, setIncludeProcessed] = useState(false);
  const [includeHtml, setIncludeHtml] = useState(false);
  const [maxText, setMaxText] = useState("10");
  const [daysText, setDaysText] = useState("");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [serverQuery, setServerQuery] = useState<string | null>(null);

  const extras = splitList(extraText);
  const subjects = splitList(subjectText);
  const extraIssues = (() => {
    const bad = extras.find((address) => !isStrictEmailAddress(address));
    if (bad) return issue(`${bad} is not a valid email address`);
    if (senders.length + extras.length > MAX_QUERY_SENDERS)
      return issue(`At most ${MAX_QUERY_SENDERS} sender addresses in total`);
    return [];
  })();
  const subjectIssues = (() => {
    const bad = subjects.find((subject) => !isSafeSubjectFilter(subject));
    if (bad)
      return issue(
        `${bad} contains characters that are not allowed (" ( ) { } : \\)`,
      );
    if (subjects.length > MAX_SUBJECTS)
      return issue(`At most ${MAX_SUBJECTS} subject filters`);
    return [];
  })();
  const maxThreads = Number(maxText);
  const maxIssues =
    Number.isInteger(maxThreads) && maxThreads >= 1 && maxThreads <= 50
      ? []
      : issue("Enter a whole number from 1 to 50");
  const newerThanDays = daysText.trim() === "" ? undefined : Number(daysText);
  const daysIssues =
    newerThanDays === undefined ||
    (Number.isInteger(newerThanDays) &&
      newerThanDays >= 1 &&
      newerThanDays <= 3650)
      ? []
      : issue("Enter a whole number from 1 to 3650, or leave it empty");

  const allSenders = [...senders, ...extras];
  const inputsOk =
    extraIssues.length === 0 &&
    subjectIssues.length === 0 &&
    maxIssues.length === 0 &&
    daysIssues.length === 0;
  const preview = (() => {
    if (!inputsOk) return { query: null, error: null };
    if (allSenders.some((address) => address.trim() === ""))
      return {
        query: null,
        error: "A sender in the rules has no address yet.",
      };
    try {
      return {
        query: buildReceiptSearchQuery(allSenders, forwardEmailAddress, {
          includeProcessed,
          subjects,
          newerThanDays,
          forwardSubjectPrefixes,
        }),
        error: null,
      };
    } catch (e) {
      return { query: null, error: e instanceof Error ? e.message : String(e) };
    }
  })();
  const canFetch = !busy && preview.query !== null && allSenders.length > 0;

  const fetchEmails = async () => {
    setBusy(true);
    setError(null);
    setStatus(null);
    const request: Omit<FetchReceiptEmailsRequest, "apiVersion" | "apiKey"> = {
      action: "fetchReceiptEmails",
      senders: allSenders,
      includeHtml,
      includeProcessed,
      maxThreads,
      ...(subjects.length > 0 ? { subjects } : {}),
      ...(newerThanDays !== undefined ? { newerThanDays } : {}),
      ...(forwardSubjectPrefixes ? { forwardSubjectPrefixes } : {}),
    };
    try {
      const { data } = await api<FetchReceiptEmailsResultDto>(request);
      setServerQuery(data.query);
      setStatus(
        `Fetched ${plural(data.emails.length, "email")} from ${plural(data.threadCount, "thread")}` +
          (data.truncated ? " (list cut short by the size limits)" : ""),
      );
      if (data.transactionSheets)
        onFetched(data.emails, data.query, data.transactionSheets);
      else onFetched(data.emails, data.query);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  return (
    <section aria-label="Fetch emails" className="space-y-3">
      <p className="text-sm text-slate-600 dark:text-slate-400">
        Fetching never labels or archives anything. The emails are only used to
        preview the rules here.
      </p>
      <div>
        <p className="text-sm font-medium">Senders from the rules</p>
        <ul
          aria-label="Senders from the rules"
          className="mt-1 flex flex-wrap gap-1"
        >
          {senders.map((address, i) => (
            <li
              key={i}
              className="rounded-full bg-slate-100 px-2 py-0.5 font-mono text-xs break-all dark:bg-slate-800"
            >
              {address || "(empty)"}
            </li>
          ))}
        </ul>
      </div>
      <TextField
        label="Extra sender addresses"
        value={extraText}
        mono
        placeholder="a@example.com, b@example.com"
        hint="Comma-separated. Only used for this fetch."
        issues={extraIssues}
        onChange={setExtraText}
      />
      <TextField
        label="Subject contains"
        value={subjectText}
        placeholder="transaction, credit"
        hint="Comma-separated, up to 5."
        issues={subjectIssues}
        onChange={setSubjectText}
      />
      <div className="grid gap-3 sm:grid-cols-2">
        <TextField
          label="Max emails"
          type="number"
          value={maxText}
          hint="Threads to search, 1-50."
          issues={maxIssues}
          onChange={setMaxText}
        />
        <TextField
          label="Newer than (days)"
          type="number"
          value={daysText}
          issues={daysIssues}
          onChange={setDaysText}
        />
      </div>
      <CheckboxField
        label="Include already-processed emails (read-only search of all mail, including labeled and archived receipts)"
        checked={includeProcessed}
        onChange={setIncludeProcessed}
      />
      <CheckboxField
        label="Include HTML bodies"
        checked={includeHtml}
        onChange={setIncludeHtml}
      />
      <div>
        <p className="text-sm font-medium">Gmail query</p>
        <code
          data-testid="query-preview"
          className="mt-1 block rounded-lg bg-slate-100 p-2 text-xs break-all dark:bg-slate-800"
        >
          {preview.query ?? "(fix the inputs above)"}
        </code>
        {preview.error && (
          <p
            role="alert"
            className="mt-1 text-sm text-red-600 dark:text-red-400"
          >
            {preview.error}
          </p>
        )}
      </div>
      <Button
        variant="primary"
        disabled={!canFetch}
        onClick={() => void fetchEmails()}
      >
        {busy ? "Fetching…" : "Fetch emails"}
      </Button>
      {status && (
        <p role="status" className="text-sm text-slate-700 dark:text-slate-300">
          {status}
        </p>
      )}
      {error && (
        <p role="alert" className="text-sm text-red-600 dark:text-red-400">
          {error}
        </p>
      )}
      {serverQuery !== null && (
        <p className="text-xs text-slate-500 dark:text-slate-400">
          Server query:{" "}
          <code data-testid="server-query" className="break-all">
            {serverQuery}
          </code>
        </p>
      )}
    </section>
  );
}
