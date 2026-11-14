import {
  DEFAULT_FORWARD_SUBJECT_PREFIXES,
  type ForwardingConfig,
  type RuleIssue,
} from "@budget/shared";
import { Button } from "../ui.tsx";
import { issuesAt, issuesFor, TextField, withOptional } from "./fields.tsx";

const replaceAt = <T,>(list: T[], index: number, value: T) =>
  list.map((item, i) => (i === index ? value : item));
const removeAt = <T,>(list: T[], index: number) =>
  list.filter((_, i) => i !== index);

/** "Forwarded-email detection (advanced)": subject prefixes, from patterns, body markers. */
export function ForwardingEditor({
  forwarding,
  issues,
  onChange,
}: {
  forwarding: ForwardingConfig;
  issues: RuleIssue[];
  onChange: (forwarding: ForwardingConfig) => void;
}) {
  const base = "forwarding";
  const at = (sub: string) => issuesFor(issues, `${base}.${sub}`);
  const prefixes = forwarding.subjectPrefixes ?? [
    ...DEFAULT_FORWARD_SUBJECT_PREFIXES,
  ];
  const setPrefixes = (next: string[]) =>
    onChange({ ...forwarding, subjectPrefixes: next });
  const hasIssues = issuesAt(issues, base).length > 0;
  return (
    <details
      open={hasIssues || undefined}
      className="rounded-xl bg-white p-3 ring-1 ring-slate-200 dark:bg-slate-800/60 dark:ring-slate-700"
    >
      <summary className="flex min-h-11 cursor-pointer items-center text-sm font-medium">
        Forwarded-email detection (advanced)
      </summary>
      <div className="mt-2 space-y-4">
        {issuesFor(issues, base).length > 0 && (
          <p className="text-sm text-red-600 dark:text-red-400">
            {issuesFor(issues, base)
              .map((issue) => issue.message)
              .join("; ")}
          </p>
        )}
        <section className="space-y-2">
          <h4 className="text-sm font-semibold">Forwarded subject prefixes</h4>
          {prefixes.map((prefix, i) => (
            <div key={i} className="flex items-end gap-2">
              <TextField
                className="flex-1"
                label={`Forward subject prefix ${i + 1}`}
                value={prefix}
                mono
                issues={at(`subjectPrefixes[${i}]`)}
                onChange={(value) => setPrefixes(replaceAt(prefixes, i, value))}
              />
              <Button
                variant="ghost"
                aria-label={`Remove forward subject prefix ${i + 1}`}
                onClick={() => setPrefixes(removeAt(prefixes, i))}
              >
                Remove
              </Button>
            </div>
          ))}
          <Button onClick={() => setPrefixes([...prefixes, ""])}>
            Add prefix
          </Button>
        </section>

        <section className="space-y-2">
          <h4 className="text-sm font-semibold">
            Original sender patterns (need a named group <code>email</code>)
          </h4>
          {forwarding.fromPatterns.map((spec, i) => (
            <div
              key={i}
              className="grid gap-2 sm:grid-cols-[1fr_6rem_auto] sm:items-end"
            >
              <TextField
                label={`From pattern ${i + 1}`}
                value={spec.pattern}
                mono
                issues={[
                  ...at(`fromPatterns[${i}]`),
                  ...at(`fromPatterns[${i}].pattern`),
                ]}
                onChange={(pattern) =>
                  onChange({
                    ...forwarding,
                    fromPatterns: replaceAt(forwarding.fromPatterns, i, {
                      ...spec,
                      pattern,
                    }),
                  })
                }
              />
              <TextField
                label={`From pattern ${i + 1} flags`}
                value={spec.flags ?? ""}
                mono
                issues={at(`fromPatterns[${i}].flags`)}
                onChange={(flags) =>
                  onChange({
                    ...forwarding,
                    fromPatterns: replaceAt(
                      forwarding.fromPatterns,
                      i,
                      withOptional(spec, "flags", flags),
                    ),
                  })
                }
              />
              <Button
                variant="ghost"
                aria-label={`Remove from pattern ${i + 1}`}
                onClick={() =>
                  onChange({
                    ...forwarding,
                    fromPatterns: removeAt(forwarding.fromPatterns, i),
                  })
                }
              >
                Remove
              </Button>
            </div>
          ))}
          <Button
            onClick={() =>
              onChange({
                ...forwarding,
                fromPatterns: [
                  ...forwarding.fromPatterns,
                  { pattern: "From: (?<email>.+)" },
                ],
              })
            }
          >
            Add from pattern
          </Button>
        </section>

        <section className="space-y-2">
          <h4 className="text-sm font-semibold">Forwarded body markers</h4>
          {forwarding.bodyMarkers.map((entry, i) => (
            <div
              key={i}
              className="grid gap-2 sm:grid-cols-[1fr_8rem_auto] sm:items-end"
            >
              <TextField
                label={`Body marker ${i + 1}`}
                value={entry.marker}
                issues={[
                  ...at(`bodyMarkers[${i}]`),
                  ...at(`bodyMarkers[${i}].marker`),
                ]}
                onChange={(marker) =>
                  onChange({
                    ...forwarding,
                    bodyMarkers: replaceAt(forwarding.bodyMarkers, i, {
                      ...entry,
                      marker,
                    }),
                  })
                }
              />
              <TextField
                label={`Body marker ${i + 1} lines to skip`}
                type="number"
                value={String(entry.skipLines)}
                issues={at(`bodyMarkers[${i}].skipLines`)}
                onChange={(value) =>
                  onChange({
                    ...forwarding,
                    bodyMarkers: replaceAt(forwarding.bodyMarkers, i, {
                      ...entry,
                      skipLines: value === "" ? 0 : Number(value),
                    }),
                  })
                }
              />
              <Button
                variant="ghost"
                aria-label={`Remove body marker ${i + 1}`}
                onClick={() =>
                  onChange({
                    ...forwarding,
                    bodyMarkers: removeAt(forwarding.bodyMarkers, i),
                  })
                }
              >
                Remove
              </Button>
            </div>
          ))}
          <Button
            onClick={() =>
              onChange({
                ...forwarding,
                bodyMarkers: [
                  ...forwarding.bodyMarkers,
                  { marker: "", skipLines: 0 },
                ],
              })
            }
          >
            Add body marker
          </Button>
        </section>
      </div>
    </details>
  );
}
