"""
Runtime configuration, read from environment variables.

The app runs in two modes:
  - Local dev (no GCP): set ENV=local. Files are stored on the local
    filesystem under DATA_DIR, scan metadata in a local JSON DB, and auth is
    disabled by default so you can click around without a Google sign-in.
  - Production (Cloud Run): GCS_BUCKET and GOOGLE_CLOUD_PROJECT are set, so
    files go to Cloud Storage and metadata to Firestore. Auth is enabled and
    enforces the allowed email domain.
"""
from __future__ import annotations

import os
from functools import lru_cache


def _bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


class Settings:
    def __init__(self) -> None:
        self.env: str = os.environ.get("ENV", "local").strip().lower()

        # GCP wiring. When GCS_BUCKET is empty we fall back to local storage,
        # which makes local development possible with zero GCP setup.
        self.gcp_project: str = os.environ.get("GOOGLE_CLOUD_PROJECT", "").strip()
        self.gcs_bucket: str = os.environ.get("GCS_BUCKET", "").strip()
        self.firestore_database: str = os.environ.get("FIRESTORE_DATABASE", "(default)").strip()

        # Local-mode storage location.
        self.data_dir: str = os.environ.get(
            "DATA_DIR", os.path.join(os.getcwd(), ".local_data")
        )

        # Auth. Default: enabled in prod, disabled locally.
        self.auth_enabled: bool = _bool("AUTH_ENABLED", default=(self.env != "local"))
        self.oauth_client_id: str = os.environ.get("OAUTH_CLIENT_ID", "").strip()
        self.allowed_email_domain: str = os.environ.get(
            "ALLOWED_EMAIL_DOMAIN", "createmusicgroup.com"
        ).strip().lower()
        # Optional comma-separated allowlist of individual emails that bypass
        # the domain check (handy for external contractors / demos).
        self.allowed_emails: set[str] = {
            e.strip().lower()
            for e in os.environ.get("ALLOWED_EMAILS", "").split(",")
            if e.strip()
        }

        # Where the built frontend lives inside the container.
        self.static_dir: str = os.environ.get(
            "STATIC_DIR", os.path.join(os.path.dirname(os.path.dirname(__file__)), "static")
        )

        # Max upload size (bytes). Default 25 MB.
        self.max_upload_bytes: int = int(os.environ.get("MAX_UPLOAD_BYTES", str(25 * 1024 * 1024)))

    @property
    def use_gcs(self) -> bool:
        return bool(self.gcs_bucket)

    @property
    def use_firestore(self) -> bool:
        return bool(self.gcp_project)


@lru_cache
def get_settings() -> Settings:
    return Settings()
