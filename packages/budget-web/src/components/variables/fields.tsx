import { useId, type ReactNode } from "react";
import type { RuleIssue } from "@budget/shared";
import { inputClass } from "../ui.tsx";
import type { FieldHelpKey } from "./help.ts";
import { HelpTip } from "./HelpTip.tsx";

/** Issues whose path is `path` itself or something inside it ("a.b" matches "a.b", "a.b.c", "a.b[0]"). */
export function issuesAt(issues: RuleIssue[], path: string): RuleIssue[] {
  return issues.filter(
    (issue) =>
      issue.path === path ||
      issue.path.startsWith(`${path}.`) ||
      issue.path.startsWith(`${path}[`),
  );
}

/** Issues for exactly this path. */
export function issuesFor(issues: RuleIssue[], path: string): RuleIssue[] {
  return issues.filter((issue) => issue.path === path);
}

export const monoInputClass = `${inputClass} font-mono text-sm`;

export function IssueText({ id, issues }: { id: string; issues: RuleIssue[] }) {
  if (issues.length === 0) return null;
  return (
    <p id={id} className="mt-1 text-sm text-red-600 dark:text-red-400">
      {issues.map((issue) => issue.message).join("; ")}
    </p>
  );
}

/** A labelled text input with its validation issues linked by aria-describedby. */
export function TextField({
  label,
  value,
  onChange,
  issues,
  mono = false,
  placeholder,
  hint,
  type = "text",
  className = "",
  help,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  issues: RuleIssue[];
  mono?: boolean;
  placeholder?: string;
  hint?: ReactNode;
  type?: "text" | "number" | "email";
  className?: string;
  help?: FieldHelpKey;
}) {
  const id = useId();
  const issueId = `${id}-issues`;
  return (
    <div className={className}>
      <div className="flex flex-wrap items-center">
        <label htmlFor={id} className="text-sm font-medium">
          {label}
        </label>
        {help && <HelpTip help={help} label={label} />}
      </div>
      <input
        id={id}
        type={type}
        value={value}
        placeholder={placeholder}
        spellCheck={false}
        autoComplete="off"
        aria-invalid={issues.length > 0 ? true : undefined}
        aria-describedby={issues.length > 0 ? issueId : undefined}
        onChange={(e) => onChange(e.target.value)}
        className={`mt-1 ${mono ? monoInputClass : inputClass}`}
      />
      {hint && (
        <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
          {hint}
        </p>
      )}
      <IssueText id={issueId} issues={issues} />
    </div>
  );
}

export function CheckboxField({
  label,
  checked,
  onChange,
  help,
}: {
  label: string;
  checked: boolean;
  onChange: (checked: boolean) => void;
  help?: FieldHelpKey;
}) {
  const box = (
    <label className="flex min-h-11 items-center gap-2 text-sm">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="h-5 w-5"
      />
      {label}
    </label>
  );
  if (!help) return box;
  return (
    <div className="flex flex-wrap items-center">
      {box}
      <HelpTip help={help} label={label} />
    </div>
  );
}

export function SelectField<T extends string>({
  label,
  value,
  options,
  onChange,
  issues,
  help,
}: {
  help?: FieldHelpKey;
  label: string;
  value: T;
  options: { value: T; label: string }[];
  onChange: (value: T) => void;
  issues: RuleIssue[];
}) {
  const id = useId();
  const issueId = `${id}-issues`;
  return (
    <div>
      <div className="flex flex-wrap items-center">
        <label htmlFor={id} className="text-sm font-medium">
          {label}
        </label>
        {help && <HelpTip help={help} label={label} />}
      </div>
      <select
        id={id}
        value={value}
        aria-invalid={issues.length > 0 ? true : undefined}
        aria-describedby={issues.length > 0 ? issueId : undefined}
        onChange={(e) => onChange(e.target.value as T)}
        className={`mt-1 ${inputClass}`}
      >
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
      <IssueText id={issueId} issues={issues} />
    </div>
  );
}

/** Returns a copy of `obj` with `key` set, or removed when `value` is "" or undefined. */
export function withOptional<T extends object, K extends keyof T>(
  obj: T,
  key: K,
  value: T[K] | "" | undefined,
): T {
  const next = { ...obj };
  if (value === "" || value === undefined) delete next[key];
  else next[key] = value;
  return next;
}
