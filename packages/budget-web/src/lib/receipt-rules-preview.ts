import {
  applyReceiptRules,
  type CompiledReceiptRules,
  type ReceiptEmailDto,
  type RuleContext,
  type RuleEmail,
  type RuleOutcome,
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

export function formatOutcome(item: PreviewItem): {
  label: "Receipt" | "Not a receipt" | "Error";
  tone: "ok" | "muted" | "error";
} {
  if (item.error !== null || !item.outcome)
    return { label: "Error", tone: "error" };
  return item.outcome.isReceipt
    ? { label: "Receipt", tone: "ok" }
    : { label: "Not a receipt", tone: "muted" };
}
