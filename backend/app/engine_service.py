"""
The bridge between the web API and the vendored desktop engine.

Responsibilities:
  - Run a scan on an uploaded workbook and persist the results + generated
    workbooks (annotated copy, issues report) to storage.
  - Apply each of the four correction types from JSON the UI sends (instead of
    from edited Excel files), re-annotate, re-scan, and update everything.
  - Read/write the org-wide LEAVE decisions from the database, materializing
    them into a temp directory the engine reads from.

The engine's apply_to_workbook() functions already accept structured Python
inputs, so we reuse them directly and never round-trip through the issues
spreadsheet the desktop tool relied on.
"""
from __future__ import annotations

import json
import os
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

# Importing the engine package puts the vendored module directory on sys.path
# so the absolute `from scan_metadata import ...` statements inside the apply
# scripts resolve correctly.
import engine  # noqa: F401
import apply_corrections as ac
import apply_format_corrections as afc
import apply_isrc_corrections as aic
import apply_missing_corrections as amc
import scan_metadata as sm

from .db import LEAVE_ARTIST, LEAVE_ISRC, get_db
from .storage import get_storage


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _key(scan_id: str, name: str) -> str:
    return f"{scan_id}/{name}"


class EngineService:
    def __init__(self) -> None:
        self.storage = get_storage()
        self.db = get_db()

    # ---- LEAVE record helpers ---------------------------------------------

    def _materialize_leave_dir(self, leave_dir: Path) -> None:
        """Write the org-wide LEAVE decisions into JSON files the engine reads."""
        artist_records = self.db.list_leave(LEAVE_ARTIST)
        isrc_records = self.db.list_leave(LEAVE_ISRC)
        # save_leave_records / save_isrc_leave_records compute the normalized
        # forms and write the exact on-disk schema the loaders expect.
        sm.save_leave_records(leave_dir, artist_records)
        sm.save_isrc_leave_records(leave_dir, isrc_records)

    def _add_artist_leave(self, new_records: list[dict]) -> int:
        existing = self.db.list_leave(LEAVE_ARTIST)
        seen = {
            frozenset(sm.normalize(v) for v in r.get("variants", []) if v)
            for r in existing
        }
        to_add = []
        for rec in new_records:
            sig = frozenset(sm.normalize(v) for v in rec.get("variants", []) if v)
            if not sig or sig in seen:
                continue
            seen.add(sig)
            to_add.append(rec)
        return self.db.add_leave(LEAVE_ARTIST, to_add) if to_add else 0

    def _add_isrc_leave(self, new_records: list[dict]) -> int:
        existing = self.db.list_leave(LEAVE_ISRC)
        seen = {
            frozenset(tuple(p) for p in r.get("signature", []))
            for r in existing
        }
        to_add = []
        for rec in new_records:
            sig = frozenset(tuple(p) for p in rec.get("signature", []))
            if not sig or sig in seen:
                continue
            seen.add(sig)
            # Make JSON-serializable (signature as list of [isrc, artist]).
            to_add.append({
                "isrc": rec.get("isrc", ""),
                "signature": [list(p) for p in rec.get("signature", [])],
                "rows": list(rec.get("rows", [])),
                "note": rec.get("note", ""),
                "added_at": rec.get("added_at", _now()),
            })
        return self.db.add_leave(LEAVE_ISRC, to_add) if to_add else 0

    # ---- Results assembly --------------------------------------------------

    @staticmethod
    def _results_payload(summaries: tuple) -> dict:
        (
            issues,
            cluster_summary,
            isrc_summary,
            missing_summary,
            missing_per_cell,
            format_column_summary,
            format_row_summary,
            format_corrections,
            splits_review,
            stats,
            split_errors,
            id_mismatches,
        ) = summaries
        return {
            "stats": stats,
            "issues": issues,
            "artistClusters": cluster_summary,
            "isrcConflicts": isrc_summary,
            "missingSummary": missing_summary,
            "missingCells": missing_per_cell,
            "formatColumns": format_column_summary,
            "formatRows": format_row_summary,
            "formatCorrections": format_corrections,
            "splitsReview": splits_review,
            "splitErrors": split_errors,
            "idMismatches": id_mismatches,
            "detectedFormat": stats.get("detected_format", "label-engine"),
        }

    @staticmethod
    def _counts(stats: dict) -> dict:
        return {
            "artistTypos": stats.get("artist_typo_cells", 0),
            "isrcConflicts": stats.get("isrc_conflict_cells", 0),
            "isrcConflictGroups": stats.get("isrc_conflicts", 0),
            "missingFields": stats.get("missing_field_issues", 0),
            "formatIssues": stats.get("format_issues", 0),
            "splitErrors": stats.get("splits_errors", 0),
            "idMismatches": stats.get("id_mismatches", 0),
            "splitsIssues": stats.get("splits_issues", 0),
            "total": stats.get("total_issues", 0),
        }

    def _persist_scan_outputs(
        self, scan_id: str, work_path: Path, summaries: tuple, tmp: Path
    ) -> dict:
        """Generate the issues + annotated workbooks from `work_path`, upload all
        three artifacts, and return the results payload."""
        (issues, cluster_summary, isrc_summary, missing_summary, missing_per_cell,
         format_column_summary, format_row_summary, format_corrections,
         splits_review, stats, split_errors, id_mismatches) = summaries

        issues_path = tmp / "issues.xlsx"
        annotated_path = tmp / "annotated.xlsx"

        sm.write_report(
            issues_path, issues, cluster_summary, isrc_summary, missing_summary,
            missing_per_cell, format_column_summary, format_row_summary,
            format_corrections, splits_review, work_path,
            split_errors=split_errors,
            id_mismatches=id_mismatches,
            detected_format=stats.get("detected_format", ""),
        )
        sm.write_annotated_copy(work_path, annotated_path, issues, stats["tracks_sheet"])

        results = self._results_payload(summaries)
        self.storage.write(_key(scan_id, "annotated.xlsx"), annotated_path.read_bytes())
        self.storage.write(_key(scan_id, "issues.xlsx"), issues_path.read_bytes())
        self.storage.write(
            _key(scan_id, "results.json"),
            json.dumps(results, ensure_ascii=False).encode("utf-8"),
        )
        return results

    # ---- Public operations -------------------------------------------------

    def create_scan(self, *, data: bytes, filename: str, user_email: str) -> dict:
        scan_id = uuid.uuid4().hex
        self.storage.write(_key(scan_id, "original.xlsx"), data)

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            leave_dir = tmp / "leave"
            leave_dir.mkdir()
            self._materialize_leave_dir(leave_dir)

            in_path = tmp / "input.xlsx"
            in_path.write_bytes(data)

            summaries = sm.analyze(in_path, issues_output_path=None, project_dir=leave_dir)
            stats = summaries[9]
            results = self._persist_scan_outputs(scan_id, in_path, summaries, tmp)

        scan = {
            "id": scan_id,
            "filename": filename,
            "created_at": _now(),
            "updated_at": _now(),
            "uploaded_by": user_email,
            "tracks_sheet": stats.get("tracks_sheet", ""),
            "sheets_scanned": stats.get("sheets_scanned", []),
            "detected_format": stats.get("detected_format", "label-engine"),
            "other_sheets": stats.get("other_sheets_with_track_isrc", []),
            "is_rescan": False,
            "status": "done",
            "counts": self._counts(stats),
            "keys": {
                "original": _key(scan_id, "original.xlsx"),
                "annotated": _key(scan_id, "annotated.xlsx"),
                "issues": _key(scan_id, "issues.xlsx"),
                "results": _key(scan_id, "results.json"),
            },
        }
        self.db.create_scan(scan)
        return {"scan": scan, "results": results}

    def get_results(self, scan_id: str) -> dict | None:
        scan = self.db.get_scan(scan_id)
        if not scan:
            return None
        raw = self.storage.read(scan["keys"]["results"])
        return json.loads(raw)

    def get_scan(self, scan_id: str) -> dict | None:
        return self.db.get_scan(scan_id)

    def list_scans(self) -> list[dict]:
        return self.db.list_scans()

    def file_bytes(self, scan_id: str, which: str) -> tuple[bytes, str] | None:
        scan = self.db.get_scan(scan_id)
        if not scan or which not in {"annotated", "issues", "original"}:
            return None
        data = self.storage.read(scan["keys"][which])
        base = Path(scan["filename"]).stem
        suffix = {"annotated": "_annotated", "issues": "_issues", "original": ""}[which]
        return data, f"{base}{suffix}.xlsx"

    def delete_scan(self, scan_id: str) -> bool:
        scan = self.db.get_scan(scan_id)
        if not scan:
            return False
        self.storage.delete_prefix(f"{scan_id}/")
        return self.db.delete_scan(scan_id)

    # ---- Applying corrections ---------------------------------------------

    def _rescan_and_persist(self, scan_id: str, work_bytes: bytes) -> dict:
        """Given the freshly-corrected working workbook bytes, re-scan and
        persist the new artifacts + counts. Returns {scan, results, ...}."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            leave_dir = tmp / "leave"
            leave_dir.mkdir()
            self._materialize_leave_dir(leave_dir)

            work_path = tmp / "work.xlsx"
            work_path.write_bytes(work_bytes)

            summaries = sm.analyze(work_path, issues_output_path=None, project_dir=leave_dir)
            stats = summaries[9]
            results = self._persist_scan_outputs(scan_id, work_path, summaries, tmp)

        fields = {
            "updated_at": _now(),
            "is_rescan": True,
            "counts": self._counts(stats),
            "tracks_sheet": stats.get("tracks_sheet", ""),
            "detected_format": stats.get("detected_format", "label-engine"),
            "sheets_scanned": stats.get("sheets_scanned", []),
        }
        self.db.update_scan(scan_id, fields)
        scan = self.db.get_scan(scan_id)
        return {"scan": scan, "results": results}

    def _download_working(self, scan_id: str, tmp: Path) -> Path:
        scan = self.db.get_scan(scan_id)
        work_path = tmp / "work.xlsx"
        work_path.write_bytes(self.storage.read(scan["keys"]["annotated"]))
        return work_path

    def apply_artist(self, scan_id: str, clusters: list[dict]) -> dict:
        replacements: dict[str, str] = {}
        leave_records: list[dict] = []
        for c in clusters:
            corr = (c.get("correction") or "").strip()
            variants = c.get("variants") or []
            if not corr:
                continue
            if corr.upper() == sm.LEAVE_MARKER:
                leave_records.append({
                    "variants": variants,
                    "cluster_id": c.get("cluster_id", ""),
                    "note": "Marked LEAVE in the webapp.",
                    "added_at": _now(),
                })
                continue
            for v in variants:
                key = sm.normalize(v)
                if key:
                    replacements[key] = corr

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            work_path = self._download_working(scan_id, tmp)
            if replacements:
                ac.apply_to_workbook(work_path, replacements)
            work_bytes = work_path.read_bytes()

        leave_added = self._add_artist_leave(leave_records) if leave_records else 0
        out = self._rescan_and_persist(scan_id, work_bytes)
        out["applied"] = {"replacements": len(replacements), "leaveAdded": leave_added}
        return out

    def apply_isrc(self, scan_id: str, rows: list[dict]) -> dict:
        review = []
        for r in rows:
            ok = bool(r.get("confirm_ok"))
            review.append({
                "conflict_id": r.get("conflict_id", ""),
                "isrc": r.get("isrc", ""),
                "excel_row": int(r["excel_row"]),
                "title": r.get("title", ""),
                "artist": r.get("artist", ""),
                "confirm_ok_raw": "OK" if ok else "",
                "confirm_ok": ok,
                "corrected_isrc": (r.get("corrected_isrc") or "").strip(),
            })
        replacements, leave_records, warnings = aic.plan_actions(review)

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            work_path = self._download_working(scan_id, tmp)
            if replacements:
                aic.apply_to_workbook(work_path, replacements)
            work_bytes = work_path.read_bytes()

        leave_added = self._add_isrc_leave(leave_records) if leave_records else 0
        out = self._rescan_and_persist(scan_id, work_bytes)
        out["applied"] = {
            "corrections": len(replacements),
            "confirmedOk": leave_added,
            "warnings": warnings,
        }
        return out

    def apply_missing(self, scan_id: str, fills_in: list[dict]) -> dict:
        fills = []
        for f in fills_in:
            value = (f.get("fill_value") or "").strip()
            if not value:
                continue
            fills.append({
                "excel_row": int(f["excel_row"]),
                "column": f.get("column", ""),
                "title": f.get("title", ""),
                "artist": f.get("artist", ""),
                "suggested": f.get("suggested", ""),
                "fill_value": value,
                "source": f.get("source", ""),
            })

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            work_path = self._download_working(scan_id, tmp)
            applied = 0
            if fills:
                audit, _skipped = amc.apply_to_workbook(work_path, fills)
                applied = len(audit)
            work_bytes = work_path.read_bytes()

        out = self._rescan_and_persist(scan_id, work_bytes)
        out["applied"] = {"fills": applied}
        return out

    def apply_format(
        self,
        scan_id: str,
        cell_corrections_in: list[dict],
        split_rows_in: list[dict],
    ) -> dict:
        cell_corrections = [
            {
                "type": c.get("type", ""),
                "excel_row": int(c["excel_row"]),
                "column": c.get("column", ""),
                "found": c.get("found", ""),
                "corrected": (c.get("corrected") or "").strip(),
            }
            for c in cell_corrections_in
        ]
        split_rows = [
            {"excel_row": int(s["excel_row"]), "splits": s.get("splits", {})}
            for s in split_rows_in
            if s.get("splits")
        ]

        # Percent-formatted columns are auto-stripped (no UI input) — derive the
        # list from the current scan's column-wide format flags.
        results = self.get_results(scan_id) or {}
        pct_columns = [
            row.get("Column", "")
            for row in results.get("formatColumns", [])
            if "% format" in str(row.get("Issue", "")).lower() and row.get("Column")
        ]

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            work_path = self._download_working(scan_id, tmp)
            cell_audit, split_audit, stripped = afc.apply_to_workbook(
                work_path, cell_corrections, split_rows, pct_columns
            )
            work_bytes = work_path.read_bytes()

        out = self._rescan_and_persist(scan_id, work_bytes)
        out["applied"] = {
            "cells": len(cell_audit),
            "splitWrites": len(split_audit),
            "columnsStripped": stripped,
        }
        return out


_service: EngineService | None = None


def get_service() -> EngineService:
    global _service
    if _service is None:
        _service = EngineService()
    return _service
