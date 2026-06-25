import { useMemo, useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { api } from "../../lib/api";
import type { Results, Row, ScanDetail } from "../../lib/types";
import { Button, ConfirmDialog, Input } from "../../components/ui";
import { DataTable, EditHead, Td, Th } from "../../components/Table";

function parseVariants(s: string): string[] {
  return String(s || "")
    .split(";")
    .map((p) => p.replace(/\s*\(\d+\)\s*$/, "").trim())
    .filter(Boolean);
}

export function ArtistTab({
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
  const clusters = results.artistClusters;
  const initial = useMemo(
    () => Object.fromEntries(clusters.map((c) => [String(c.Cluster), String(c.Correction ?? "")])),
    [clusters]
  );
  const [corrections, setCorrections] = useState<Record<string, string>>(() => ({ ...initial }));
  const [confirmOpen, setConfirmOpen] = useState(false);

  const summary = useMemo(() => {
    let apply = 0;
    let leave = 0;
    let skip = 0;
    for (const c of clusters) {
      const v = (corrections[String(c.Cluster)] ?? "").trim();
      if (!v) skip += 1;
      else if (v.toUpperCase() === "LEAVE") leave += 1;
      else apply += 1;
    }
    return { apply, leave, skip };
  }, [clusters, corrections]);

  const apply = useMutation({
    mutationFn: () => {
      const payload = clusters.map((c) => ({
        cluster_id: String(c.Cluster),
        correction: corrections[String(c.Cluster)] ?? "",
        variants: parseVariants(String(c["Variants (count)"] ?? "")),
      }));
      return api.applyArtist(scanId, payload);
    },
    onSuccess: (d) => {
      onApplied(d);
      const a = (d as ScanDetail & { applied?: { replacements: number; leaveAdded: number } })
        .applied;
      notify(`Applied ${a?.replacements ?? 0} name fix(es), ${a?.leaveAdded ?? 0} left as-is.`);
    },
    onError: (e: Error) => {
      setConfirmOpen(false);
      notify(e.message, "error");
    },
  });

  const set = (id: string, v: string) => setCorrections((p) => ({ ...p, [id]: v }));

  if (clusters.length === 0) {
    return <p className="py-8 text-center text-sm text-muted">No artist-name clusters flagged.</p>;
  }

  return (
    <div className="space-y-4">
      <p className="text-sm text-muted">
        Each cluster groups names the scanner thinks are the same artist. The{" "}
        <span className="font-medium">Correction</span> is pre-filled with the most likely
        spelling — keep it, edit it, clear it to skip, or click <span className="font-medium">Leave</span>{" "}
        for genuinely-different artists (remembered across scans).
      </p>
      <DataTable
        scroll
        head={
          <>
            <Th>Cluster</Th>
            <Th>Variants found</Th>
            <EditHead>Correction</EditHead>
            <Th></Th>
          </>
        }
      >
        {clusters.map((c: Row) => {
          const id = String(c.Cluster);
          const value = corrections[id] ?? "";
          return (
            <tr key={id} className="hover:bg-slate-50">
              <Td className="font-mono text-xs text-muted">{id}</Td>
              <Td className="max-w-md text-xs">{String(c["Variants (count)"] ?? "")}</Td>
              <Td className="w-56">
                <Input
                  value={value}
                  changed={value !== (initial[id] ?? "")}
                  onChange={(e) => set(id, e.target.value)}
                />
              </Td>
              <Td>
                <Button
                  variant="ghost"
                  className="px-2 py-1 text-xs"
                  onClick={() => set(id, "LEAVE")}
                >
                  Leave
                </Button>
              </Td>
            </tr>
          );
        })}
      </DataTable>
      <div className="flex justify-end">
        <Button onClick={() => setConfirmOpen(true)} disabled={apply.isPending}>
          Apply name fixes
        </Button>
      </div>

      <ConfirmDialog
        open={confirmOpen}
        title="Apply artist-name fixes?"
        pending={apply.isPending}
        confirmLabel="Apply name fixes"
        onCancel={() => setConfirmOpen(false)}
        onConfirm={() => apply.mutate()}
        message={
          <>
            <span className="font-medium text-ink">{summary.apply}</span> cluster(s) will have
            their variants replaced, <span className="font-medium text-ink">{summary.leave}</span>{" "}
            marked Leave, and <span className="font-medium text-ink">{summary.skip}</span> skipped.
            This rewrites cells across all artist columns and re-scans the sheet.
          </>
        }
      />
    </div>
  );
}
