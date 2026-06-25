"""
File storage abstraction.

Stores the uploaded sheet plus the generated working/annotated/issues
workbooks and the results JSON for each scan. Two backends:
  - GcsStorage: Google Cloud Storage (production).
  - LocalStorage: a directory on disk (local dev / tests).

Keys look like "<scan_id>/working.xlsx". The backend chooses the
implementation based on whether GCS_BUCKET is configured.
"""
from __future__ import annotations

import os
import shutil
from abc import ABC, abstractmethod

from .config import get_settings


class Storage(ABC):
    @abstractmethod
    def write(self, key: str, data: bytes) -> None: ...

    @abstractmethod
    def read(self, key: str) -> bytes: ...

    @abstractmethod
    def exists(self, key: str) -> bool: ...

    @abstractmethod
    def delete_prefix(self, prefix: str) -> None: ...


class LocalStorage(Storage):
    def __init__(self, root: str) -> None:
        self.root = root
        os.makedirs(self.root, exist_ok=True)

    def _path(self, key: str) -> str:
        # Keys are forward-slash separated; map onto the local filesystem.
        return os.path.join(self.root, *key.split("/"))

    def write(self, key: str, data: bytes) -> None:
        path = self._path(key)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(data)

    def read(self, key: str) -> bytes:
        with open(self._path(key), "rb") as fh:
            return fh.read()

    def exists(self, key: str) -> bool:
        return os.path.exists(self._path(key))

    def delete_prefix(self, prefix: str) -> None:
        path = self._path(prefix)
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)


class GcsStorage(Storage):
    def __init__(self, bucket_name: str) -> None:
        from google.cloud import storage as gcs  # imported lazily

        self._client = gcs.Client()
        self._bucket = self._client.bucket(bucket_name)

    def write(self, key: str, data: bytes) -> None:
        blob = self._bucket.blob(key)
        blob.upload_from_string(
            data,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if key.endswith(".xlsx")
            else "application/octet-stream",
        )

    def read(self, key: str) -> bytes:
        return self._bucket.blob(key).download_as_bytes()

    def exists(self, key: str) -> bool:
        return self._bucket.blob(key).exists()

    def delete_prefix(self, prefix: str) -> None:
        for blob in self._client.list_blobs(self._bucket, prefix=prefix):
            blob.delete()


_storage: Storage | None = None


def get_storage() -> Storage:
    global _storage
    if _storage is None:
        settings = get_settings()
        if settings.use_gcs:
            _storage = GcsStorage(settings.gcs_bucket)
        else:
            _storage = LocalStorage(os.path.join(settings.data_dir, "files"))
    return _storage
