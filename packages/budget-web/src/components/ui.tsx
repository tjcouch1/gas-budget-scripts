import { useEffect, type ReactNode } from "react";

export function Button({
  variant = "secondary",
  className = "",
  ...props
}: React.ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: "primary" | "secondary" | "danger" | "ghost";
}) {
  const styles = {
    primary: "bg-teal-700 text-white hover:bg-teal-800 disabled:bg-teal-700/40",
    secondary:
      "bg-white text-slate-800 ring-1 ring-slate-300 hover:bg-slate-50 disabled:opacity-50 dark:bg-slate-800 dark:text-slate-100 dark:ring-slate-600 dark:hover:bg-slate-700",
    danger: "bg-red-600 text-white hover:bg-red-700 disabled:opacity-50",
    ghost:
      "text-slate-700 hover:bg-slate-200/60 dark:text-slate-200 dark:hover:bg-slate-700/60",
  }[variant];
  return (
    <button
      type="button"
      className={`min-h-11 rounded-lg px-4 py-2 text-sm font-medium transition-colors disabled:cursor-not-allowed ${styles} ${className}`}
      {...props}
    />
  );
}

export const inputClass =
  "min-h-11 w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-base text-slate-900 outline-none focus:border-teal-600 focus:ring-2 focus:ring-teal-600/30 dark:border-slate-600 dark:bg-slate-900 dark:text-slate-100";

/** Modal dialog; on small screens it is a bottom sheet. */
export function Dialog({
  title,
  onClose,
  children,
  wide = false,
}: {
  title: string;
  onClose?: () => void;
  children: ReactNode;
  wide?: boolean;
}) {
  useEffect(() => {
    if (!onClose) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  return (
    <div
      className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/50 sm:items-center"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose?.();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className={`pb-safe max-h-[92vh] w-full overflow-y-auto rounded-t-2xl bg-white p-4 shadow-xl sm:rounded-2xl dark:bg-slate-900 ${wide ? "sm:max-w-3xl" : "sm:max-w-lg"}`}
      >
        <div className="mb-3 flex items-start justify-between gap-3">
          <h2 className="text-lg font-semibold">{title}</h2>
          {onClose && (
            <button
              type="button"
              aria-label="Close"
              onClick={onClose}
              className="min-h-11 min-w-11 rounded-lg text-xl text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
            >
              ×
            </button>
          )}
        </div>
        {children}
      </div>
    </div>
  );
}
