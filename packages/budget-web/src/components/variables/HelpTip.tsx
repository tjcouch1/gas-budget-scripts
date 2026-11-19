import { useId, useState } from "react";
import { FIELD_HELP, type FieldHelpKey } from "./help.ts";

/**
 * An info button for a field: hover shows the native `title`; tapping toggles the same text inline
 * (works on touch and for screen readers via aria-describedby).
 */
export function HelpTip({
  help,
  label,
}: {
  help: FieldHelpKey;
  /** The field's name, for the button's accessible name. */
  label: string;
}) {
  const [open, setOpen] = useState(false);
  const id = useId();
  const text = FIELD_HELP[help];
  return (
    <>
      <button
        type="button"
        title={text}
        aria-label={`Help for ${label}`}
        aria-expanded={open}
        aria-describedby={open ? id : undefined}
        onClick={() => setOpen((o) => !o)}
        className="ml-1 inline-flex h-5 w-5 shrink-0 items-center justify-center rounded-full border border-slate-400 text-xs leading-none text-slate-600 hover:bg-slate-100 dark:border-slate-500 dark:text-slate-300 dark:hover:bg-slate-700"
      >
        <span aria-hidden="true">i</span>
      </button>
      {open && (
        <p
          id={id}
          role="note"
          className="mt-1 basis-full text-xs text-slate-600 dark:text-slate-300"
        >
          {text}
        </p>
      )}
    </>
  );
}
