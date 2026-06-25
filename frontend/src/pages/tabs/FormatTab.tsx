import { useMemo, useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { api } from "../../lib/api";
import type { Results, Row, ScanDetail } from "../../lib/types";
import { Badge, Button, ConfirmDialog, Input } from "../../components/ui";
import { DataTable, EditHead, Td, Th } from "../../components/Table";

const SPLIT_RE = /^Artist \d+ Master Split$/;

function splitColumns(rows: Row[]): string[] {
  const set = new Set<string>();
  rows.forEach((r) => Object.keys(r).forEach((k) => SPLIT_RE.test(k) && set.add(k)));
  return [...set].sort((a, b) => {
    const na = Number(a.match(/\d+/)?.[0] ?? 0);
    const nb = Number(b.match(/\d+/)?.[0] ?? 0);
    return na - nb;
  });
}

function num(v: unknown): number {
  const n = parseFloat(String(v).replace("%", ""));
  return Number.isFinite(n) ? n : 0;
}

export function FormatTab({
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
  const corrections = results.formatCorrections;
  const splits = results.splitsReview;
  const cols = results.formatColumns;
  const splitCols = useMemo(() => splitColumns(splits), [splits]);

  const [cellVals, setCellVals] = useState<Record<string, string>>(() =>
    Object.fromEntries(
      corrections.map((c) => [`${c["Excel Row"]}|${c.Column}`, String(c["Corrected Value"] ?? "")])
    )
  );
  const [splitVals, setSplitVals] = useState<Record<string, Record<string, string>>>(() =>
    Object.fromEntries(
      splits.map((s) => [
        String(s["Excel Row"]),
        Object.fromEntries(splitCols.map((c) => [c, s[c] == null ? "" : String(s[c])])),
      ])
    )
  );
  const [confirmOpen, setConfirmOpen] = useState(false);

  const correctedCount = useMemo(
    () => corrections.filter((c) => (cellVals[`${c["Excel Row"]}|${c.Column}`] ?? "").trim()).length,
    [corrections, cellVals]
  );
  const initialSplit = useMemo(
    () =>
      Object.fromEntries(
        splits.map((s) => [
          String(s["Excel Row"]),
          Object.fromEntries(splitCols.map((c) => [c, s[c] == null ? "" : String(s[c])])),
        ])
      ),
    [splits, splitCols]
  );

  const apply = useMutation({
    mutationFn: () => {
      const cell_corrections = corrections.map((c) => ({
        type: String(c.Type ?? ""),
        excel_row: Number(c["Excel Row"]),
        column: String(c.Column ?? ""),
        found: String(c["Found Value"] ?? ""),
        corrected: cellVals[`${c["Excel Row"]}|${c.Column}`] ?? "",
      }));
      const split_rows = splits.map((s) => ({
        excel_row: Number(s["Excel Row"]),
        splits: splitVals[String(s["Excel Row"])] ?? {},
      }));
      return api.applyFormat(scanId, cell_corrections, split_rows);
    },
    onSuccess: (d) => {
      onApplied(d);
      const a = (d as ScanDetail & {
        applied?: { cells: number; splitWrites: number; columnsStripped: string[] };
      }).applied;
      notify(
        `Applied ${a?.cells ?? 0} value fix(es), ${a?.splitWrites ?? 0} split write(s), ${
          a?.columnsStripped?.length ?? 0
        } column(s) de-percented.`
      );
    },
    onError: (e: Error) => {
      setConfirmOpen(false);
      notify(e.message, "error");
    },
  });

  const nothing = corrections.length === 0 && splits.length === 0 && cols.length === 0;
  if (nothing) {
    return <p className="py-8 text-center text-sm text-muted">No format issues.</p>;
  }

  return (
    <div className="space-y-8">
      {cols.length > 0 && (
        <div className="rounded-lg border border-format/60 bg-format/30 p-3 text-sm text-format-ink">
          <span className="font-semibold">Auto-fixed on Apply:</span>{" "}
          {cols.map((c) => String(c.Column)).join(", ")} — percent formatting will be stripped so
          values upload correctly.
        </div>
      )}

      {corrections.length > 0 && (
        <div className="space-y-3">
          <h3 className="text-sm font-semibold text-ink">ISRC / UPC format</h3>
          <DataTable
            head={
              <>
                <Th>Type</Th>
                <Th>Row</Th>
                <Th>Column</Th>
                <Th>Found</Th>
                <EditHead>Corrected value</EditHead>
              </>
            }
          >
            {corrections.map((c: Row) => {
              const k = `${c["Excel Row"]}|${c.Column}`;
              return (
                <tr key={k} className="hover:bg-slate-50">
                  <Td className="text-xs text-muted">{String(c.Type ?? "")}</Td>
                  <Td className="text-xs text-muted">{String(c["Excel Row"] ?? "")}</Td>
                  <Td className="text-xs font-medium">{String(c.Column ?? "")}</Td>
                  <Td className="font-mono text-xs">{String(c["Found Value"] ?? "")}</Td>
                  <Td className="w-48">
                    <Input
                      value={cellVals[k] ?? ""}
                      changed={Boolean((cellVals[k] ?? "").trim())}
                      placeholder="—"
                      onChange={(e) => setCellVals((p) => ({ ...p, [k]: e.target.value }))}
                    />
                  </Td>
                </tr>
              );
            })}
          </DataTable>
        </div>
      )}

      {splits.length > 0 && (
        <div className="space-y-3">
          <h3 className="text-sm font-semibold text-ink">Master splits (must sum to 100)</h3>
          <DataTable
            head={
              <>
                <Th>Row</Th>
                <Th>Track</Th>
                {splitCols.map((c) => (
                  <EditHead key={c}>{c.replace(" Master Split", "")}</EditHead>
                ))}
                <Th>Sum</Th>
              </>
            }
          >
            {splits.map((s: Row) => {
              const rk = String(s["Excel Row"]);
              const rowVals = splitVals[rk] ?? {};
              const sum = splitCols.reduce((acc, c) => acc + num(rowVals[c]), 0);
              const ok = Math.abs(sum - 100) <= 0.5;
              return (
                <tr key={rk} className="hover:bg-slate-50">
                  <Td className="text-xs text-muted">{rk}</Td>
                  <Td className="max-w-[12rem] truncate text-xs">
                    {String(s["Track Title"] ?? "")}
                  </Td>
                  {splitCols.map((c) => (
                    <Td key={c} className="w-20">
                      <Input
                        value={rowVals[c] ?? ""}
                        changed={(rowVals[c] ?? "") !== (initialSplit[rk]?.[c] ?? "")}
                        onChange={(e) =>
                          setSplitVals((p) => ({
                            ...p,
                            [rk]: { ...p[rk], [c]: e.target.value },
                          }))
                        }
                      />
                    </Td>
                  ))}
                  <Td>
                    <Badge tone={ok ? "clean" : "artist"}>{sum.toFixed(0)}</Badge>
                  </Td>
                </tr>
              );
            })}
          </DataTable>
        </div>
      )}

      <div className="flex justify-end">
        <Button onClick={() => setConfirmOpen(true)} disabled={apply.isPending}>
          Apply format fixes
        </Button>
      </div>

      <ConfirmDialog
        open={confirmOpen}
        title="Apply format fixes?"
        pending={apply.isPending}
        confirmLabel="Apply format fixes"
        onCancel={() => setConfirmOpen(false)}
        onConfirm={() => apply.mutate()}
        message={
          <>
            <span className="font-medium text-ink">{correctedCount}</span> value fix(es),{" "}
            <span className="font-medium text-ink">{splits.length}</span> master-split row(s), and{" "}
            <span className="font-medium text-ink">{cols.length}</span> column(s) de-percented.
          </>
        }
      />
    </div>
  );
}
