import type { ReactNode } from "react";
import { cn } from "./ui";

export function DataTable({
  head,
  children,
  scroll = false,
}: {
  head: ReactNode;
  children: ReactNode;
  /** Cap the height and scroll the body, keeping the header pinned. Use for
   * dense tabs (artist clusters, ISRC conflicts) with 100+ rows. */
  scroll?: boolean;
}) {
  return (
    <div
      className={cn(
        "scroll-area overflow-x-auto rounded-xl border border-slate-200 bg-white",
        scroll && "max-h-[62vh] overflow-y-auto"
      )}
    >
      <table className="w-full border-collapse text-sm">
        <thead>
          <tr className="text-left text-xs font-semibold uppercase tracking-wide text-white">
            {head}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}

export function Th({ children, className = "" }: { children?: ReactNode; className?: string }) {
  return (
    <th className={`sticky top-0 z-10 whitespace-nowrap bg-navy px-3 py-2.5 ${className}`}>
      {children}
    </th>
  );
}

export function Td({ children, className = "" }: { children?: ReactNode; className?: string }) {
  return (
    <td className={`border-b border-slate-100 px-3 py-2 align-middle ${className}`}>{children}</td>
  );
}

/** A column header that marks an editable ("yellow") input column, matching the
 * desktop tool's convention. */
export function EditHead({ children }: { children: ReactNode }) {
  return (
    <th className="sticky top-0 z-10 whitespace-nowrap bg-isrc px-3 py-2.5 text-left text-xs font-semibold uppercase tracking-wide text-isrc-ink">
      {children}
    </th>
  );
}
