import { useMemo, useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { api } from "../../lib/api";
import type { Results, Row, ScanDetail } from "../../lib/types";
import { Button, ConfirmDialog, Input } from "../../components/ui";
import { DataTable, EditHead, Td, Th } from "../../components/Table";

function rowKey(r: Row): string {
  return `${r["Excel Row"]}|${r.Column}`;
}

export function MissingTab({
  scanId,
  results,
  onApplied,
  notify,
}: {
  scanId: string;
  results: Results;
  onApplied: (d: ScanDetail) => void;
  notify: (msg: string, tone?: "ok" | "error") => void;
}) {
  const cells = results.missingCells;
  const initial = useMemo(
    () => Object.fromEntries(cells.map((c) => [rowKey(c), String(c["Fill Value"] ?? "")])),
    [cells]
  );
  const [values, setValues] = useState<Record<string, string>>(() => ({ ...initial }));
  const [confirmOpen, setConfirmOpen] = useState(false);

  const fillCount = useMemo(
    () => cells.filter((c) => (values[rowKey(c)] ?? "").trim()).length,
    [cells, values]
  );

  const apply = useMutation({
    mutationFn: () => {
      const payload = cells.map((c) => ({
        excel_row: Number(c["Excel Row"]),
        column: String(c.Column ?? ""),
        title: String(c["Track Title"] ?? ""),
        artist: String(c["Track Display Artist"] ?? ""),
        suggested: String(c["Suggested Fill"] ?? ""),
        fill_value: values[rowKey(c)] ?? "",
        source: String(c["Suggestion source"] ?? ""),
      }));
      return api.applyMissing(scanId, payload);
    },
    onSuccess: (d) => {
      onApplied(d);
      const a = (d as ScanDetail & { applied?: { fills: number } }).applied;
      notify(`Filled ${a?.fills ?? 0} cell(s).`);
    },
    onError: (e: Error) => {
      setConfirmOpen(false);
      notify(e.message, "error");
    },
  });

  if (cells.length === 0) {
    return <p className="py-8 text-center text-sm text-muted">No missing required fields.</p>;
  }

  return (
    <div className="space-y-4">
      <p className="text-sm text-muted">
        One row per blank required cell. <span className="font-medium">Fill value</span> is
        pre-filled where the scanner can infer it confidently — confirm or override, clear to skip.
      </p>
      <DataTable
        scroll
        head={
          <>
            <Th>Row</Th>
            <Th>Column</Th>
            <Th>Track</Th>
            <Th>Reason</Th>
            <EditHead>Fill value</EditHead>
            <Th>Source</Th>
          </>
        }
      >
        {cells.map((c: Row) => {
          const k = rowKey(c);
          return (
            <tr key={k} className="hover:bg-slate-50">
              <Td className="text-xs text-muted">{String(c["Excel Row"] ?? "")}</Td>
              <Td className="text-xs font-medium">{String(c.Column ?? "")}</Td>
              <Td className="max-w-[14rem] truncate text-xs">{String(c["Track Title"] ?? "")}</Td>
              <Td className="text-xs text-muted">{String(c["Reason for missing"] ?? "")}</Td>
              <Td className="w-56">
                <Input
                  value={values[k] ?? ""}
                  changed={(values[k] ?? "") !== (initial[k] ?? "")}
                  placeholder="—"
                  onChange={(e) => setValues((p) => ({ ...p, [k]: e.target.value }))}
                />
              </Td>
              <Td className="max-w-[16rem] truncate text-[11px] text-slate-400">
                {String(c["Suggestion source"] ?? "")}
              </Td>
            </tr>
          );
        })}
      </DataTable>
      <div className="flex justify-end">
        <Button onClick={() => setConfirmOpen(true)} disabled={apply.isPending}>
          Apply fills
        </Button>
      </div>

      <ConfirmDialog
        open={confirmOpen}
        title="Fill missing fields?"
        pending={apply.isPending}
        confirmLabel="Apply fills"
        onCancel={() => setConfirmOpen(false)}
        onConfirm={() => apply.mutate()}
        message={
          <>
            <span className="font-medium text-ink">{fillCount}</span> cell(s) will be filled with
            the values shown. Empty rows are skipped.
          </>
        }
      />
    </div>
  );
}
