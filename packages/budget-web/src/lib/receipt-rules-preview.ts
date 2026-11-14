import {
  applyReceiptRules,
  type CompiledReceiptRules,
  type ReceiptEmailDto,
  type RuleContext,
  type RuleEmail,
  type RuleOutcome,
  type RuleOutcomeReason,
} from "@budget/shared";

export interface PreviewItem {
  email: ReceiptEmailDto;
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

/** Runs the shared engine over fetched emails, the same code the sheet's import uses. */
export function previewReceiptRules(
  rules: CompiledReceiptRules | null,
  emails: ReceiptEmailDto[],
  context: RuleContext,
): PreviewItem[] {
  return emails.map((email) => {
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
