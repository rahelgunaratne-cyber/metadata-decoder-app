"""FastAPI entrypoint.

Serves the JSON API under /api/* and, in production, the built React single-
page app for everything else (so the whole product is one Cloud Run service
on one URL).
"""
from __future__ import annotations

import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.requests import Request

from .api import scans
from .config import get_settings

settings = get_settings()
app = FastAPI(title="Metadata Decoder", version="1.0.0")

# During local dev the Vite dev server runs on a different origin; allow it.
if settings.env == "local":
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
        allow_methods=["*"],
        allow_headers=["*"],
    )

app.include_router(scans.router)


@app.get("/api/config")
async def public_config() -> dict:
    """Non-secret values the frontend needs at boot (the OAuth client ID and
    whether auth is on)."""
    return {
        "authEnabled": settings.auth_enabled,
        "oauthClientId": settings.oauth_client_id,
        "allowedDomain": settings.allowed_email_domain,
    }


@app.get("/api/healthz")
async def healthz() -> dict:
    return {"status": "ok"}


# ---- Static SPA (production) ----------------------------------------------
# The Dockerfile builds the frontend into backend/static. When present, serve
# it: hashed assets straight from /assets, and index.html for any non-API path
# so client-side routing (e.g. /scans/<id>) works on refresh.
_static_dir = settings.static_dir
if os.path.isdir(_static_dir):
    app.mount("/assets", StaticFiles(directory=os.path.join(_static_dir, "assets")), name="assets")

    @app.get("/{full_path:path}")
    async def spa(full_path: str, request: Request):
        if full_path.startswith("api/"):
            return JSONResponse({"detail": "Not found"}, status_code=404)
        candidate = os.path.join(_static_dir, full_path)
        if full_path and os.path.isfile(candidate):
            return FileResponse(candidate)
        return FileResponse(os.path.join(_static_dir, "index.html"))
