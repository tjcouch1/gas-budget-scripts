import { useEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";

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

/** Modal dialog; on small screens it is a bottom sheet that can be dragged down to dismiss. */
export function Dialog({
  title,
  heading,
  onClose,
  children,
  wide = false,
}: {
  /** Accessible name of the dialog; also the visible heading unless `heading` is given. */
  title: string;
  /** Replaces the visible heading (e.g. an editable title). */
  heading?: ReactNode;
  onClose?: () => void;
  children: ReactNode;
  wide?: boolean;
}) {
  const [dragY, setDragY] = useState(0);
  const dragStart = useRef<number | null>(null);
  useEffect(() => {
    if (!onClose) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  const dismissable = onClose !== undefined;
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
        style={dragY > 0 ? { transform: `translateY(${dragY}px)` } : undefined}
        className={`pb-safe max-h-[92vh] w-full overflow-y-auto rounded-t-2xl bg-white p-4 pt-2 shadow-xl sm:rounded-2xl sm:pt-4 dark:bg-slate-900 ${wide ? "sm:max-w-3xl" : "sm:max-w-lg"}`}
      >
        {dismissable && (
          <div
            data-testid="drag-handle"
            aria-hidden="true"
            className="mx-auto mb-2 flex h-6 w-full cursor-grab touch-none items-center justify-center sm:hidden"
            onPointerDown={(e) => {
              dragStart.current = e.clientY;
              e.currentTarget.setPointerCapture?.(e.pointerId);
            }}
            onPointerMove={(e) => {
              if (dragStart.current !== null)
                setDragY(Math.max(0, e.clientY - dragStart.current));
            }}
            onPointerUp={(e) => {
              const dy =
                dragStart.current === null ? 0 : e.clientY - dragStart.current;
              dragStart.current = null;
              setDragY(0);
              if (dy > DRAG_DISMISS_PX) onClose?.();
            }}
            onPointerCancel={() => {
              dragStart.current = null;
              setDragY(0);
            }}
          >
            <span className="h-1.5 w-10 rounded-full bg-slate-300 dark:bg-slate-600" />
          </div>
        )}
        <div className="mb-3 flex items-start justify-between gap-3">
          {heading ?? <h2 className="text-lg font-semibold">{title}</h2>}
          {onClose && (
            <button
              type="button"
              aria-label="Close"
              onClick={onClose}
              className="min-h-11 min-w-11 shrink-0 rounded-lg text-xl text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800"
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

/** Dragging the sheet down further than this dismisses it. */
const DRAG_DISMISS_PX = 90;

/**
 * Save / Discard / Go back guard for closing a dialog that has unsaved changes.
 * `requestClose` closes at once when nothing changed, otherwise opens the prompt
 * (and toggles it off again, so Escape while it is open means "Go back").
 */
export function useUnsavedGuard({
  dirty,
  onDiscard,
}: {
  dirty: boolean;
  onDiscard: () => void;
}) {
  const [prompting, setPrompting] = useState(false);
  const requestClose = () => {
    if (!dirty) onDiscard();
    else setPrompting((p) => !p);
  };
  return { prompting, setPrompting, requestClose };
}

export function UnsavedPrompt({
  canSave,
  onSave,
  onDiscard,
  onBack,
}: {
  canSave: boolean;
  onSave: () => void;
  onDiscard: () => void;
  onBack: () => void;
}) {
  return createPortal(
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/60 p-4 sm:items-center">
      <div
        role="alertdialog"
        aria-modal="true"
        aria-label="Unsaved changes"
        className="w-full max-w-sm space-y-3 rounded-2xl bg-white p-4 shadow-xl dark:bg-slate-900"
      >
        <h2 className="text-lg font-semibold">Unsaved changes</h2>
        <p className="text-sm text-slate-600 dark:text-slate-400">
          Save your changes before closing, or discard them?
        </p>
        <div className="flex flex-col gap-2">
          <Button variant="primary" onClick={onSave} disabled={!canSave}>
            Save
          </Button>
          <Button variant="danger" onClick={onDiscard}>
            Discard
          </Button>
          <Button autoFocus onClick={onBack}>
            Go back
          </Button>
        </div>
      </div>
    </div>,
    document.body,
  );
}
