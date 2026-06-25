"""
Metadata + decision persistence.

Stores:
  - scans:        one document per uploaded sheet (metadata + issue counts +
                  storage keys). The big results payload lives in storage as
                  results.json; only a pointer is kept here.
  - leave_artist: artist clusters marked "LEAVE" (intentionally similar names).
  - leave_isrc:   ISRC conflicts confirmed OK (intentional duplicates).

Two backends, chosen by whether GOOGLE_CLOUD_PROJECT is set:
  - FirestoreDB (production)
  - LocalDB     (a single JSON file on disk, for local dev / tests)
"""
from __future__ import annotations

import json
import os
import threading
from abc import ABC, abstractmethod
from typing import Any

from .config import get_settings

SCANS = "scans"
LEAVE_ARTIST = "leave_artist"
LEAVE_ISRC = "leave_isrc"


class Database(ABC):
    @abstractmethod
    def create_scan(self, scan: dict) -> None: ...

    @abstractmethod
    def get_scan(self, scan_id: str) -> dict | None: ...

    @abstractmethod
    def list_scans(self) -> list[dict]: ...

    @abstractmethod
    def update_scan(self, scan_id: str, fields: dict) -> None: ...

    @abstractmethod
    def delete_scan(self, scan_id: str) -> bool: ...

    @abstractmethod
    def list_leave(self, kind: str) -> list[dict]: ...

    @abstractmethod
    def add_leave(self, kind: str, records: list[dict]) -> int: ...


class LocalDB(Database):
    """A simple JSON-file-backed store, adequate for single-process dev."""

    def __init__(self, path: str) -> None:
        self.path = path
        self._lock = threading.Lock()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if not os.path.exists(path):
            self._write({SCANS: {}, LEAVE_ARTIST: [], LEAVE_ISRC: []})

    def _read(self) -> dict:
        with open(self.path, "r", encoding="utf-8") as fh:
            return json.load(fh)

    def _write(self, data: dict) -> None:
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def create_scan(self, scan: dict) -> None:
        with self._lock:
            data = self._read()
            data[SCANS][scan["id"]] = scan
            self._write(data)

    def get_scan(self, scan_id: str) -> dict | None:
        with self._lock:
            return self._read()[SCANS].get(scan_id)

    def list_scans(self) -> list[dict]:
        with self._lock:
            scans = list(self._read()[SCANS].values())
        scans.sort(key=lambda s: s.get("created_at", ""), reverse=True)
        return scans

    def update_scan(self, scan_id: str, fields: dict) -> None:
        with self._lock:
            data = self._read()
            if scan_id in data[SCANS]:
                data[SCANS][scan_id].update(fields)
                self._write(data)

    def delete_scan(self, scan_id: str) -> bool:
        with self._lock:
            data = self._read()
            if scan_id in data[SCANS]:
                del data[SCANS][scan_id]
                self._write(data)
                return True
            return False

    def list_leave(self, kind: str) -> list[dict]:
        with self._lock:
            return list(self._read().get(kind, []))

    def add_leave(self, kind: str, records: list[dict]) -> int:
        with self._lock:
            data = self._read()
            existing = data.setdefault(kind, [])
            existing.extend(records)
            self._write(data)
            return len(records)


class FirestoreDB(Database):
    def __init__(self, project: str, database: str) -> None:
        from google.cloud import firestore  # imported lazily

        self._fs = firestore.Client(project=project, database=database)

    def create_scan(self, scan: dict) -> None:
        self._fs.collection(SCANS).document(scan["id"]).set(scan)

    def get_scan(self, scan_id: str) -> dict | None:
        doc = self._fs.collection(SCANS).document(scan_id).get()
        return doc.to_dict() if doc.exists else None

    def list_scans(self) -> list[dict]:
        from google.cloud.firestore_v1.base_query import FieldFilter  # noqa: F401

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
        settings = get_settings()
        if settings.use_firestore:
            _db = FirestoreDB(settings.gcp_project, settings.firestore_database)
        else:
            _db = LocalDB(os.path.join(settings.data_dir, "db.json"))
    return _db
