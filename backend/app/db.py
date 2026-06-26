"""
Metadata + decision persistence — Firestore only.

Collections:
  - scans:        one document per uploaded sheet (metadata + issue counts +
                  storage keys). The big results payload lives in GCS as
                  results.json; only a pointer is kept here.
  - leave_artist: artist clusters marked "LEAVE" (intentionally similar names).
  - leave_isrc:   ISRC conflicts confirmed OK (intentional duplicates).
"""
from __future__ import annotations

from .config import get_settings

SCANS = "scans"
LEAVE_ARTIST = "leave_artist"
LEAVE_ISRC = "leave_isrc"


class Database:
    def __init__(self, project: str, database: str) -> None:
        from google.cloud import firestore

        self._fs = firestore.Client(project=project, database=database)

    def create_scan(self, scan: dict) -> None:
        self._fs.collection(SCANS).document(scan["id"]).set(scan)

    def get_scan(self, scan_id: str) -> dict | None:
        doc = self._fs.collection(SCANS).document(scan_id).get()
        return doc.to_dict() if doc.exists else None

    def list_scans(self) -> list[dict]:
        docs = (
            self._fs.collection(SCANS)
            .order_by("created_at", direction="DESCENDING")
            .stream()
        )
        return [d.to_dict() for d in docs]

    def update_scan(self, scan_id: str, fields: dict) -> None:
        self._fs.collection(SCANS).document(scan_id).update(fields)

    def delete_scan(self, scan_id: str) -> bool:
        ref = self._fs.collection(SCANS).document(scan_id)
        if not ref.get().exists:
            return False
        ref.delete()
        return True

    def list_leave(self, kind: str) -> list[dict]:
        return [d.to_dict() for d in self._fs.collection(kind).stream()]

    def add_leave(self, kind: str, records: list[dict]) -> int:
        col = self._fs.collection(kind)
        for rec in records:
            col.add(rec)
        return len(records)


_db: Database | None = None


def get_db() -> Database:
    global _db
    if _db is None:
        s = get_settings()
        _db = Database(s.gcp_project, s.firestore_database)
    return _db
