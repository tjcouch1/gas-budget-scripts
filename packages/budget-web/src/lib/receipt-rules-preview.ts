import {
  applyReceiptRules,
  findTransactionSheet,
  planThreadImport,
  sortRowsByDate,
  type ImportMessageInput,
  type TransactionSheetRangeDto,
  type CompiledReceiptRules,
  type ReceiptEmailDto,
  type RuleContext,
  type RuleEmail,
  type RuleOutcome,
  type RuleOutcomeReason,
} from "@budget/shared";

/** One transaction row the import would write */
export interface ImportRowPreview {
  date: Date;
  name: string | undefined;
  cost: number | undefined;
  category: string | undefined;
  type: string | undefined;
  /** Non-empty: the name cell gets this note and a red background */
  errorMessage: string;
  /** Non-empty (no error): the name cell gets this note and a pink background */
  note: string;
  /** Name of the transaction sheet it goes on, null when none fits (the import would fail), undefined when sheets are unknown */
  sheet: string | null | undefined;
}

export interface ImportPreview {
  rows: ImportRowPreview[];
  /** Why this email adds no row, or null when it does */
  skip: string | null;
  /** The thread already has the Receipts/Scripted label (a normal import would not even fetch it) */
  alreadyProcessed: boolean;
}

export interface PreviewItem {
  email: ReceiptEmailDto;
  /** What the import would write for this email, computed by the shared import planner */
  import: ImportPreview;
  /** null when `error` is set */
  outcome: RuleOutcome | null;
  /** Message thrown by applyReceiptRules, or "Rules are invalid" without compiled rules. */
  error: string | null;
}

export function toRuleEmail(dto: ReceiptEmailDto): RuleEmail {
  return {
    from: dto.from,
    to: dto.to,
    subject: dto.subject,
    plainBody: dto.plainBody,
    ...(dto.htmlBody !== undefined ? { htmlBody: dto.htmlBody } : {}),
  };
}

type RuleResult = Omit<PreviewItem, "import">;

const EMPTY_IMPORT: ImportPreview = {
  rows: [],
  skip: null,
  alreadyProcessed: false,
};

const SCRIPTED_LABEL = "Receipts/Scripted";

/**
 * Runs the shared engine over fetched emails, then the shared import planner (the same code the
 * sheet's import uses) to show the rows each email would add. Pure: nothing is written anywhere.
 */
export function previewReceiptRules(
  rules: CompiledReceiptRules | null,
  emails: ReceiptEmailDto[],
  context: RuleContext,
  sheets?: TransactionSheetRangeDto[],
): PreviewItem[] {
  const results: RuleResult[] = emails.map((email) => {
    if (!rules) return { email, outcome: null, error: "Rules are invalid" };
    try {
      return {
        email,
        outcome: applyReceiptRules(rules, toRuleEmail(email), context),
        error: null,
      };
    } catch (error) {
      return {
        email,
        outcome: null,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  });
  if (!rules)
    return results.map((r) => ({
      ...r,
      import: {
        ...EMPTY_IMPORT,
        skip: "Rules are invalid",
        alreadyProcessed: r.email.labels.includes(SCRIPTED_LABEL),
      },
    }));

  const sheetRanges = sheets?.map((s) => ({
    name: s.name,
    startDate: new Date(s.startDate),
    endDate: new Date(s.endDate),
  }));
  const sheetFor = (date: Date) =>
    sheetRanges
      ? (findTransactionSheet(date, sheetRanges)?.name ?? null)
      : undefined;

  // Plan per thread, as the import does
  const imports = new Map<ReceiptEmailDto, ImportPreview>();
  const threads = new Map<string, RuleResult[]>();
  for (const r of results)
    threads.set(r.email.threadId, [
      ...(threads.get(r.email.threadId) ?? []),
      r,
    ]);
  for (const [threadId, items] of threads) {
    const dates = items.map((r) => new Date(r.email.date));
    const lastMessageDate = new Date(
      Math.max(...dates.map((d) => d.getTime())),
    );
    const messages = items.map((r, i): ImportMessageInput => ({
      messageId: r.email.messageId,
      subject: r.email.subject,
      fromAddress: r.email.from,
      date: dates[i]!,
      dateText: String(dates[i]),
      plainBody: r.email.plainBody,
      outcome: r.outcome,
      error: r.outcome ? null : `Error: ${r.error ?? ""}`,
    }));
    const plan = planThreadImport({
      threadId,
      lastMessageDate: () => lastMessageDate,
      messages,
    });
    const rows = sortRowsByDate(plan.rows);
    items.forEach((r, i) => {
      // The empty row of an errored thread belongs to the thread's first email
      const mine = plan.rows.filter(
        (row) =>
          row.messageId === r.email.messageId ||
          (row.messageId === null && i === 0),
      );
      const own = rows.filter((row) => mine.includes(row));
      let skip: string | null = null;
      if (own.length === 0 && r.outcome) {
        const reason = r.outcome.reason
          ? REASON_TEXT[r.outcome.reason]
          : REASON_TEXT["no-receipt-data"];
        skip = `Not a receipt (${reason}): ${
          plan.hasInformation
            ? "no row for this message, only a note on the thread's first row"
            : "the thread is only labeled Receipts/Ignored"
        }`;
      }
      imports.set(r.email, {
        skip,
        alreadyProcessed: r.email.labels.includes(SCRIPTED_LABEL),
        rows: own.map((row) => ({
          date: new Date(row.date.getTime()),
          name: row.name,
          cost: row.cost,
          category: row.category,
          type: row.type,
          errorMessage: row.errorMessage,
          note: row.note,
          sheet: sheetFor(new Date(row.date.getTime())),
        })),
      });
    });
  }
  return results.map((r) => ({ ...r, import: imports.get(r.email)! }));
}

/** Why an email is not a receipt, in words */
export const REASON_TEXT: Record<RuleOutcomeReason, string> = {
  "unknown-sender": "unknown sender",
  "not-forwardable": "forwards of this sender are not used",
  "not-a-forward": "from the forward address but not a forward",
  "no-rule-matched": "no rule matched",
  "no-receipt-data": "no cost or name found",
};

export function formatOutcome(item: PreviewItem): {
  label: "Receipt" | "Not a receipt" | "Error";
  tone: "ok" | "muted" | "error";
  /** Only for "Not a receipt": why (REASON_TEXT) */
  detail?: string;
} {
  if (item.error !== null || !item.outcome)
    return { label: "Error", tone: "error" };
  if (item.outcome.isReceipt) return { label: "Receipt", tone: "ok" };
  const { reason } = item.outcome;
  return {
    label: "Not a receipt",
    tone: "muted",
    ...(reason ? { detail: REASON_TEXT[reason] } : {}),
  };
}
