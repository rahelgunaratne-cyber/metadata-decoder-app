import type { Counts } from "../lib/types";
import { Badge } from "./ui";

export function CountPills({ counts }: { counts: Counts }) {
  if (counts.total === 0) {
    return <Badge tone="clean">All clean</Badge>;
  }
  return (
    <div className="flex flex-wrap gap-1.5">
      {counts.artistTypos > 0 && <Badge tone="artist">{counts.artistTypos} artist</Badge>}
      {counts.isrcConflicts > 0 && <Badge tone="isrc">{counts.isrcConflicts} ISRC</Badge>}
      {counts.missingFields > 0 && <Badge tone="missing">{counts.missingFields} missing</Badge>}
      {counts.formatIssues > 0 && <Badge tone="format">{counts.formatIssues} format</Badge>}
    </div>
  );
}
