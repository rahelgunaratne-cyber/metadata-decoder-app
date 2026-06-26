"""
File storage — Google Cloud Storage only.

Keys look like "<scan_id>/working.xlsx".
"""
from __future__ import annotations

from .config import get_settings


class Storage:
    def __init__(self, bucket_name: str) -> None:
        from google.cloud import storage as gcs

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
        _storage = Storage(get_settings().gcs_bucket)
    return _storage
