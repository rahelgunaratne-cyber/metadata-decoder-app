import { useMemo, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { api } from "../lib/api";
import type { Scan } from "../lib/types";
import { CountPills } from "../components/CountPills";
import { Button, Card, EmptyState, Input, Spinner, Toast } from "../components/ui";

function timeAgo(iso: string): string {
  if (!iso) return "";
  const then = new Date(iso).getTime();
  const mins = Math.round((Date.now() - then) / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.round(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  return new Date(iso).toLocaleDateString();
}

function UploadZone() {
  const navigate = useNavigate();
  const qc = useQueryClient();
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const upload = useMutation({
    mutationFn: (file: File) => api.createScan(file),
    onSuccess: (data) => {
      qc.invalidateQueries({ queryKey: ["scans"] });
      navigate(`/scans/${data.scan.id}`);
    },
    onError: (e: Error) => setError(e.message),
  });

  function handleFiles(files: FileList | null) {
    setError(null);
    const file = files?.[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setError("Please choose an .xlsx spreadsheet.");
      return;
    }
    upload.mutate(file);
  }

  return (
    <Card
      className={`relative overflow-hidden border-2 border-dashed p-10 text-center transition-colors ${
        dragging ? "border-navy bg-slate-50" : "border-slate-200"
      }`}
    >
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          handleFiles(e.dataTransfer.files);
        }}
      >
        <div className="mx-auto mb-4 grid h-14 w-14 place-items-center rounded-2xl bg-slate-100 text-2xl">
          ⬆
        </div>
        <h2 className="text-lg font-semibold text-ink">Scan a metadata sheet</h2>
        <p className="mx-auto mt-1 max-w-md text-sm text-muted">
          Drag an <span className="font-medium">.xlsx</span> file here, or choose one. The
          decoder checks artist names, ISRCs, missing fields, and formats — then lets you fix
          everything right here.
        </p>
        <div className="mt-5 flex justify-center">
          {upload.isPending ? (
            <span className="inline-flex items-center gap-2 text-sm font-medium text-navy">
              <Spinner /> Scanning…
            </span>
          ) : (
            <Button onClick={() => inputRef.current?.click()}>Choose file</Button>
          )}
        </div>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx"
          className="hidden"
          onChange={(e) => handleFiles(e.target.files)}
        />
      </div>
      {error && <Toast message={error} tone="error" />}
    </Card>
  );
}

function ScanRow({ scan, onDelete }: { scan: Scan; onDelete: (id: string) => void }) {
  const navigate = useNavigate();
  return (
    <tr
      className="cursor-pointer border-b border-slate-100 last:border-0 hover:bg-slate-50"
      onClick={() => navigate(`/scans/${scan.id}`)}
    >
      <td className="px-4 py-3">
        <div className="font-medium text-ink">{scan.filename}</div>
        <div className="text-xs text-muted">
          {scan.tracks_sheet}
          {scan.is_rescan ? " · re-scanned" : ""} · {timeAgo(scan.updated_at)}
        </div>
        {scan.detected_format && scan.detected_format !== "unknown" && (
          <span className="mt-0.5 inline-flex items-center rounded-full bg-navy/10 px-2 py-0.5 text-[10px] font-medium text-navy">
            {scan.detected_format}
          </span>
        )}
        {scan.detected_format === "unknown" && (
          <span className="mt-0.5 inline-flex items-center rounded-full bg-amber-100 px-2 py-0.5 text-[10px] font-medium text-amber-700">
            Unrecognized format
          </span>
        )}
      </td>
      <td className="px-4 py-3">
        <CountPills counts={scan.counts} />
      </td>
      <td className="px-2 py-3 text-right" onClick={(e) => e.stopPropagation()}>
        <Button
          variant="ghost"
          className="px-2 text-slate-400 hover:text-red-600"
          title="Remove from dashboard"
          onClick={() => onDelete(scan.id)}
        >
          ✕
        </Button>
      </td>
    </tr>
  );
}

export function HomePage() {
  const qc = useQueryClient();
  const [filter, setFilter] = useState("");
  const { data, isLoading } = useQuery({ queryKey: ["scans"], queryFn: api.listScans });

  const del = useMutation({
    mutationFn: (id: string) => api.deleteScan(id),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["scans"] }),
  });

  const scans = data?.scans ?? [];
  const filtered = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return scans;
    return scans.filter(
      (s) =>
        s.filename.toLowerCase().includes(q) ||
        (s.tracks_sheet || "").toLowerCase().includes(q)
    );
  }, [scans, filter]);

  return (
    <div className="space-y-8">
      <UploadZone />

      <div>
        <div className="mb-3 flex items-center justify-between gap-4">
          <h2 className="text-base font-semibold text-ink">
            Scans <span className="text-muted">({scans.length})</span>
          </h2>
          {scans.length > 0 && (
            <Input
              placeholder="Search by sheet name or tab…"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              className="max-w-xs"
            />
          )}
        </div>

        {isLoading ? (
          <div className="grid place-items-center py-12">
            <Spinner className="h-6 w-6" />
          </div>
        ) : filtered.length === 0 ? (
          <EmptyState
            title={scans.length === 0 ? "No scans yet" : "No matches"}
            hint={scans.length === 0 ? "Upload a sheet above to get started." : undefined}
          />
        ) : (
          <Card className="overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-navy text-left text-xs font-semibold uppercase tracking-wide text-white">
                  <th className="px-4 py-2.5">Sheet</th>
                  <th className="px-4 py-2.5">Issues</th>
                  <th className="px-2 py-2.5"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((s) => (
                  <ScanRow key={s.id} scan={s} onDelete={del.mutate} />
                ))}
              </tbody>
            </table>
          </Card>
        )}
      </div>
    </div>
  );
}
