import { useMemo, useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { api } from "../../lib/api";
import type { Results, Row, ScanDetail } from "../../lib/types";
import { Button, ConfirmDialog, Input } from "../../components/ui";
import { DataTable, EditHead, Td, Th } from "../../components/Table";

interface RowEdit {
  confirm_ok: boolean;
  corrected_isrc: string;
}

function rowKey(r: Row): string {
  return `${r.ISRC}|${r["Excel Row"]}`;
}

export function IsrcTab({
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
  const rows = results.isrcConflicts;
  const [edits, setEdits] = useState<Record<string, RowEdit>>(() =>
    Object.fromEntries(
      rows.map((r) => [
        rowKey(r),
        {
          confirm_ok: Boolean(r["Confirm OK?"]),
          corrected_isrc: String(r["Corrected ISRC"] ?? ""),
        },
      ])
    )
  );
  const [confirmOpen, setConfirmOpen] = useState(false);

  const summary = useMemo(() => {
    let corrections = 0;
    let confirms = 0;
    for (const r of rows) {
      const e = edits[rowKey(r)];
      if (e?.corrected_isrc?.trim()) corrections += 1;
      else if (e?.confirm_ok) confirms += 1;
    }
    return { corrections, confirms };
  }, [rows, edits]);

  const apply = useMutation({
    mutationFn: () => {
      const payload = rows.map((r) => {
        const e = edits[rowKey(r)];
        return {
          conflict_id: String(r.Conflict ?? ""),
          isrc: String(r.ISRC ?? ""),
          excel_row: Number(r["Excel Row"]),
          title: String(r["Track Title"] ?? ""),
          artist: String(r["Track Display Artist"] ?? ""),
          confirm_ok: e?.confirm_ok ?? false,
          corrected_isrc: e?.corrected_isrc ?? "",
        };
      });
      return api.applyIsrc(scanId, payload);
    },
    onSuccess: (d) => {
      onApplied(d);
      const a = (d as ScanDetail & {
        applied?: { corrections: number; confirmedOk: number; warnings: string[] };
      }).applied;
      notify(`Applied ${a?.corrections ?? 0} ISRC fix(es), ${a?.confirmedOk ?? 0} confirmed OK.`);
    },
    onError: (e: Error) => {
      setConfirmOpen(false);
      notify(e.message, "error");
    },
  });

  const set = (k: string, patch: Partial<RowEdit>) =>
    setEdits((p) => ({ ...p, [k]: { ...p[k], ...patch } }));

  if (rows.length === 0) {
    return <p className="py-8 text-center text-sm text-muted">No duplicate-ISRC conflicts.</p>;
  }

  return (
    <div className="space-y-4">
      <p className="text-sm text-muted">
        These rows share an ISRC with a different artist. Tick{" "}
        <span className="font-medium">OK</span> if a duplicate is intentional (remembered across
        scans), or type a corrected ISRC to fix it.
      </p>
      <DataTable
        scroll
        head={
          <>
            <Th>Conflict</Th>
            <Th>ISRC</Th>
            <Th>Row</Th>
            <Th>Track</Th>
            <Th>Artist</Th>
            <EditHead>OK?</EditHead>
            <EditHead>Corrected ISRC</EditHead>
          </>
        }
      >
        {rows.map((r: Row) => {
          const k = rowKey(r);
          const e = edits[k];
          return (
            <tr key={k} className="hover:bg-slate-50">
              <Td className="font-mono text-xs text-muted">{String(r.Conflict ?? "")}</Td>
              <Td className="font-mono text-xs">{String(r.ISRC ?? "")}</Td>
              <Td className="text-xs text-muted">{String(r["Excel Row"] ?? "")}</Td>
              <Td className="max-w-[14rem] truncate text-xs">{String(r["Track Title"] ?? "")}</Td>
              <Td className="max-w-[12rem] truncate text-xs">
                {String(r["Track Display Artist"] ?? "")}
              </Td>
              <Td className="text-center">
                <input
                  type="checkbox"
                  checked={e?.confirm_ok ?? false}
                  onChange={(ev) => set(k, { confirm_ok: ev.target.checked })}
                  className="h-4 w-4 accent-[color:var(--color-navy)]"
                />
              </Td>
              <Td className="w-44">
                <Input
                  value={e?.corrected_isrc ?? ""}
                  changed={Boolean(e?.corrected_isrc?.trim())}
                  placeholder="—"
                  onChange={(ev) => set(k, { corrected_isrc: ev.target.value })}
                />
              </Td>
            </tr>
          );
        })}
      </DataTable>
      <div className="flex justify-end">
        <Button onClick={() => setConfirmOpen(true)} disabled={apply.isPending}>
          Apply ISRC fixes
        </Button>
      </div>

      <ConfirmDialog
        open={confirmOpen}
        title="Apply ISRC fixes?"
        pending={apply.isPending}
        confirmLabel="Apply ISRC fixes"
        onCancel={() => setConfirmOpen(false)}
        onConfirm={() => apply.mutate()}
        message={
          <>
            <span className="font-medium text-ink">{summary.corrections}</span> ISRC(s) will be
            corrected and <span className="font-medium text-ink">{summary.confirms}</span> confirmed
            as intentional duplicates (remembered across future scans).
          </>
        }
      />
    </div>
  );
}
