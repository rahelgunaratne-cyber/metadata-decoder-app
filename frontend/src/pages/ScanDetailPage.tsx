import { useEffect, useState } from "react";
import { Link, useParams } from "react-router-dom";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { api, downloadFile } from "../lib/api";
import type { Counts, ScanDetail } from "../lib/types";
import { CountPills } from "../components/CountPills";
import { Button, Card, Spinner, Toast } from "../components/ui";
import { ArtistTab } from "./tabs/ArtistTab";
import { IsrcTab } from "./tabs/IsrcTab";
import { MissingTab } from "./tabs/MissingTab";
import { FormatTab } from "./tabs/FormatTab";

type TabKey = "artist" | "isrc" | "missing" | "format";

const TAB_META: { key: TabKey; label: string; count: (c: Counts) => number }[] = [
  { key: "artist", label: "Artist names", count: (c) => c.artistTypos },
  { key: "isrc", label: "Duplicate ISRCs", count: (c) => c.isrcConflicts },
  { key: "missing", label: "Missing fields", count: (c) => c.missingFields },
  { key: "format", label: "Formats", count: (c) => c.formatIssues },
];

export function ScanDetailPage() {
  const { id = "" } = useParams();
  const qc = useQueryClient();
  const [tab, setTab] = useState<TabKey>("artist");
  const [toast, setToast] = useState<{ msg: string; tone: "ok" | "error" } | null>(null);

  const { data, isLoading, isError, error } = useQuery({
    queryKey: ["scan", id],
    queryFn: () => api.getScan(id),
  });

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 3500);
    return () => clearTimeout(t);
  }, [toast]);

  function notify(msg: string, tone: "ok" | "error" = "ok") {
    setToast({ msg, tone });
  }

  function onApplied(detail: ScanDetail) {
    qc.setQueryData(["scan", id], detail);
    qc.invalidateQueries({ queryKey: ["scans"] });
  }

  if (isLoading) {
    return (
      <div className="grid place-items-center py-20">
        <Spinner className="h-6 w-6" />
      </div>
    );
  }
  if (isError || !data) {
    return (
      <Card className="p-8 text-center">
        <p className="text-sm text-red-600">{(error as Error)?.message ?? "Scan not found."}</p>
        <Link to="/" className="mt-3 inline-block text-sm font-medium text-navy">
          ← Back to scans
        </Link>
      </Card>
    );
  }

  const { scan, results } = data;
  const tabProps = { scanId: id, results, onApplied, notify };

  return (
    <div className="space-y-6">
      <div>
        <Link to="/" className="text-sm font-medium text-muted hover:text-navy">
          ← All scans
        </Link>
      </div>

      <div className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <h1 className="text-xl font-bold text-ink">{scan.filename}</h1>
          <p className="mt-0.5 text-sm text-muted">
            Tab: <span className="font-medium">{scan.tracks_sheet}</span>
            {scan.is_rescan ? " · re-scanned" : ""}
          </p>
          <div className="mt-2">
            <CountPills counts={scan.counts} />
          </div>
        </div>
        <div className="flex gap-2">
          <Button variant="secondary" onClick={() => downloadFile(id, "issues")}>
            Download report
          </Button>
          <Button variant="secondary" onClick={() => downloadFile(id, "annotated")}>
            Download annotated
          </Button>
        </div>
      </div>

      {scan.counts.total === 0 ? (
        <Card className="p-10 text-center">
          <div className="text-3xl">✓</div>
          <p className="mt-2 text-base font-semibold text-clean-ink">
            All clean — this sheet is ready to ingest.
          </p>
        </Card>
      ) : (
        <>
          <div className="flex flex-wrap gap-1 border-b border-slate-200">
            {TAB_META.map((t) => {
              const n = t.count(scan.counts);
              const active = tab === t.key;
              return (
                <button
                  key={t.key}
                  onClick={() => setTab(t.key)}
                  className={`-mb-px flex items-center gap-2 border-b-2 px-4 py-2.5 text-sm font-medium transition-colors ${
                    active
                      ? "border-navy text-navy"
                      : "border-transparent text-muted hover:text-ink"
                  }`}
                >
                  {t.label}
                  <span
                    className={`rounded-full px-1.5 py-0.5 text-xs ${
                      n > 0 ? "bg-slate-200 text-slate-700" : "bg-clean text-clean-ink"
                    }`}
                  >
                    {n}
                  </span>
                </button>
              );
            })}
          </div>

          {/* Keying on updated_at remounts the active tab after each apply so its
              local edit state is rebuilt from the fresh scan results. */}
          <div key={`${tab}-${scan.updated_at}`}>
            {tab === "artist" && <ArtistTab {...tabProps} />}
            {tab === "isrc" && <IsrcTab {...tabProps} />}
            {tab === "missing" && <MissingTab {...tabProps} />}
            {tab === "format" && <FormatTab {...tabProps} />}
          </div>
        </>
      )}

      {toast && <Toast message={toast.msg} tone={toast.tone} />}
    </div>
  );
}
