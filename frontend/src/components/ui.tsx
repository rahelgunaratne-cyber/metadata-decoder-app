import type {
  ButtonHTMLAttributes,
  InputHTMLAttributes,
  ReactNode,
} from "react";

export function cn(...parts: (string | false | null | undefined)[]): string {
  return parts.filter(Boolean).join(" ");
}

type ButtonVariant = "primary" | "secondary" | "ghost" | "danger";

export function Button({
  variant = "primary",
  className,
  children,
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { variant?: ButtonVariant }) {
  const base =
    "inline-flex items-center justify-center gap-2 rounded-lg px-4 py-2 text-sm font-semibold transition-colors disabled:cursor-not-allowed disabled:opacity-50 focus:outline-none focus-visible:ring-2 focus-visible:ring-navy/40";
  const variants: Record<ButtonVariant, string> = {
    primary: "bg-navy text-white hover:bg-navy-700",
    secondary: "bg-white text-navy border border-slate-200 hover:bg-slate-50",
    ghost: "text-slate-600 hover:bg-slate-100",
    danger: "bg-white text-red-600 border border-red-200 hover:bg-red-50",
  };
  return (
    <button className={cn(base, variants[variant], className)} {...props}>
      {children}
    </button>
  );
}

export function Card({ className, children }: { className?: string; children: ReactNode }) {
  return (
    <div
      className={cn(
        "rounded-xl border border-slate-200 bg-white shadow-[0_1px_3px_rgba(0,0,0,0.05)]",
        className
      )}
    >
      {children}
    </div>
  );
}

export function Input({
  className,
  changed = false,
  ...props
}: InputHTMLAttributes<HTMLInputElement> & { changed?: boolean }) {
  return (
    <input
      className={cn(
        "w-full rounded-md border bg-white px-2.5 py-1.5 text-sm text-ink",
        "focus:border-navy focus:outline-none focus:ring-2 focus:ring-navy/15",
        "disabled:bg-slate-50 disabled:text-slate-400",
        changed ? "border-navy ring-2 ring-navy/20" : "border-slate-300",
        className
      )}
      {...props}
    />
  );
}

export function Spinner({ className }: { className?: string }) {
  return (
    <span
      className={cn(
        "inline-block h-4 w-4 animate-spin rounded-full border-2 border-slate-300 border-t-navy",
        className
      )}
      aria-label="loading"
    />
  );
}

export function Badge({
  children,
  tone = "neutral",
}: {
  children: ReactNode;
  tone?: "neutral" | "artist" | "isrc" | "missing" | "format" | "clean";
}) {
  const tones: Record<string, string> = {
    neutral: "bg-slate-100 text-slate-600",
    artist: "bg-artist text-artist-ink",
    isrc: "bg-isrc text-isrc-ink",
    missing: "bg-missing text-missing-ink",
    format: "bg-format text-format-ink",
    clean: "bg-clean text-clean-ink",
  };
  return (
    <span
      className={cn(
        "inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-medium",
        tones[tone]
      )}
    >
      {children}
    </span>
  );
}

export function EmptyState({ title, hint }: { title: string; hint?: string }) {
  return (
    <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 bg-white/60 px-6 py-12 text-center">
      <p className="text-sm font-semibold text-slate-600">{title}</p>
      {hint && <p className="mt-1 text-xs text-slate-400">{hint}</p>}
    </div>
  );
}

export function ConfirmDialog({
  open,
  title,
  message,
  confirmLabel = "Apply",
  pending = false,
  onConfirm,
  onCancel,
}: {
  open: boolean;
  title: string;
  message: ReactNode;
  confirmLabel?: string;
  pending?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/30 px-4">
      <Card className="w-full max-w-md p-6">
        <h3 className="text-base font-semibold text-ink">{title}</h3>
        <div className="mt-2 text-sm text-muted">{message}</div>
        <div className="mt-6 flex justify-end gap-2">
          <Button variant="secondary" onClick={onCancel} disabled={pending}>
            Cancel
          </Button>
          <Button onClick={onConfirm} disabled={pending}>
            {pending ? <Spinner /> : null}
            {confirmLabel}
          </Button>
        </div>
      </Card>
    </div>
  );
}

export function Toast({ message, tone = "ok" }: { message: string; tone?: "ok" | "error" }) {
  return (
    <div
      className={cn(
        "fixed bottom-6 right-6 z-50 rounded-lg px-4 py-3 text-sm font-medium shadow-lg",
        tone === "ok" ? "bg-clean text-clean-ink" : "bg-red-50 text-red-700 border border-red-200"
      )}
    >
      {message}
    </div>
  );
}
