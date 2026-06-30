import type { Results, Row } from "../../lib/types";
import { DataTable, Td, Th } from "../../components/Table";
import { Badge } from "../../components/ui";

export function SplitsTab({ results }: { results: Results }) {
  const splitErrors = results.splitErrors ?? [];
  const idMismatches = results.idMismatches ?? [];

  if (splitErrors.length === 0 && idMismatches.length === 0) {
    return <p className="py-8 text-center text-sm text-muted">No splits issues.</p>;
  }

  return (
    <div className="space-y-8">
      {splitErrors.length > 0 && (
        <div className="space-y-3">
          <h3 className="text-sm font-semibold text-ink">
            Split totals — Net rows must sum to 100%
          </h3>
          <DataTable
            head={
              <>
                <Th>Release ID</Th>
                <Th>ISRC</Th>
                <Th>Track Title</Th>
                <Th>Rows</Th>
                <Th>Net Total</Th>
                <Th>Difference</Th>
              </>
            }
          >
            {splitErrors.map((r: Row, i: number) => {
              const diff = Number(r["Difference from 100"] ?? 0);
              return (
                <tr key={i} className="hover:bg-slate-50">
                  <Td className="font-mono text-xs text-muted">{String(r["Release ID"] ?? "")}</Td>
                  <Td className="font-mono text-xs">{String(r["ISRC"] ?? "")}</Td>
                  <Td className="max-w-[16rem] truncate text-xs">{String(r["Track Title"] ?? "")}</Td>
                  <Td className="text-xs text-muted">{String(r["Rows"] ?? "")}</Td>
                  <Td>
                    <Badge tone={Math.abs(diff) <= 0.5 ? "clean" : "artist"}>
                      {String(r["Split Total"] ?? "")}%
                    </Badge>
                  </Td>
                  <Td className="text-xs text-muted">
                    {diff > 0 ? `+${diff}` : String(diff)}
                  </Td>
                </tr>
              );
            })}
          </DataTable>
        </div>
      )}

      {idMismatches.length > 0 && (
        <div className="space-y-3">
          <h3 className="text-sm font-semibold text-ink">ID mismatches</h3>
          <p className="text-xs text-muted">
            The same Client or Account ID maps to more than one name, or the same name
            maps to more than one ID. These should each be a 1-to-1 relationship.
          </p>
          <DataTable
            head={
              <>
                <Th>Type</Th>
                <Th>ID</Th>
                <Th>Names found</Th>
                <Th>Rows</Th>
              </>
            }
          >
            {idMismatches.map((r: Row, i: number) => (
              <tr key={i} className="hover:bg-slate-50">
                <Td className="text-xs text-muted">{String(r["Type"] ?? "")}</Td>
                <Td className="font-mono text-xs">{String(r["ID"] ?? "")}</Td>
                <Td className="max-w-[24rem] text-xs">{String(r["Names found"] ?? "")}</Td>
                <Td className="text-xs text-muted">{String(r["Rows"] ?? "")}</Td>
              </tr>
            ))}
          </DataTable>
        </div>
      )}
    </div>
  );
}
