"""API routes for scans and corrections. Every route requires a signed-in,
domain-allowed user (enforced by the require_user dependency)."""
from __future__ import annotations

import io

from fastapi import APIRouter, Depends, HTTPException, UploadFile, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from ..auth import User, require_user
from ..config import Settings, get_settings
from ..engine_service import get_service

router = APIRouter(prefix="/api", tags=["scans"])

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---- Request bodies --------------------------------------------------------


class ArtistCluster(BaseModel):
    cluster_id: str = ""
    correction: str = ""
    variants: list[str] = []


class ArtistCorrections(BaseModel):
    clusters: list[ArtistCluster]


class IsrcRow(BaseModel):
    conflict_id: str = ""
    isrc: str = ""
    excel_row: int
    title: str = ""
    artist: str = ""
    confirm_ok: bool = False
    corrected_isrc: str = ""


class IsrcCorrections(BaseModel):
    rows: list[IsrcRow]


class MissingFill(BaseModel):
    excel_row: int
    column: str
    title: str = ""
    artist: str = ""
    suggested: str = ""
    fill_value: str = ""
    source: str = ""


class MissingCorrections(BaseModel):
    fills: list[MissingFill]


class FormatCell(BaseModel):
    type: str = ""
    excel_row: int
    column: str
    found: str = ""
    corrected: str = ""


class SplitRow(BaseModel):
    excel_row: int
    splits: dict[str, object] = {}


class FormatCorrections(BaseModel):
    cell_corrections: list[FormatCell] = []
    split_rows: list[SplitRow] = []


# ---- Routes ----------------------------------------------------------------


@router.get("/me")
async def me(user: User = Depends(require_user)) -> dict:
    return {"email": user.get("email"), "name": user.get("name"), "picture": user.get("picture")}


@router.post("/scans", status_code=status.HTTP_201_CREATED)
async def create_scan(
    file: UploadFile,
    user: User = Depends(require_user),
    settings: Settings = Depends(get_settings),
) -> dict:
    name = file.filename or "sheet.xlsx"
    if not name.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    if len(data) > settings.max_upload_bytes:
        raise HTTPException(status_code=413, detail="File is too large.")
    try:
        return get_service().create_scan(data=data, filename=name, user_email=user.email)
    except Exception as exc:  # surface engine failures cleanly
        raise HTTPException(status_code=422, detail=f"Could not scan this sheet: {exc}")


@router.get("/scans")
async def list_scans(user: User = Depends(require_user)) -> dict:
    return {"scans": get_service().list_scans()}


@router.get("/scans/{scan_id}")
async def get_scan(scan_id: str, user: User = Depends(require_user)) -> dict:
    svc = get_service()
    scan = svc.get_scan(scan_id)
    if not scan:
        raise HTTPException(status_code=404, detail="Scan not found.")
    return {"scan": scan, "results": svc.get_results(scan_id)}


@router.delete("/scans/{scan_id}")
async def delete_scan(scan_id: str, user: User = Depends(require_user)) -> dict:
    if not get_service().delete_scan(scan_id):
        raise HTTPException(status_code=404, detail="Scan not found.")
    return {"deleted": scan_id}


@router.get("/scans/{scan_id}/files/{which}")
async def download_file(scan_id: str, which: str, user: User = Depends(require_user)):
    result = get_service().file_bytes(scan_id, which)
    if result is None:
        raise HTTPException(status_code=404, detail="File not found.")
    data, download_name = result
    return StreamingResponse(
        io.BytesIO(data),
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


def _require_scan(scan_id: str):
    svc = get_service()
    if not svc.get_scan(scan_id):
        raise HTTPException(status_code=404, detail="Scan not found.")
    return svc


@router.post("/scans/{scan_id}/corrections/artist")
async def apply_artist(scan_id: str, body: ArtistCorrections, user: User = Depends(require_user)) -> dict:
    svc = _require_scan(scan_id)
    return svc.apply_artist(scan_id, [c.model_dump() for c in body.clusters])


@router.post("/scans/{scan_id}/corrections/isrc")
async def apply_isrc(scan_id: str, body: IsrcCorrections, user: User = Depends(require_user)) -> dict:
    svc = _require_scan(scan_id)
    return svc.apply_isrc(scan_id, [r.model_dump() for r in body.rows])


@router.post("/scans/{scan_id}/corrections/missing")
async def apply_missing(scan_id: str, body: MissingCorrections, user: User = Depends(require_user)) -> dict:
    svc = _require_scan(scan_id)
    return svc.apply_missing(scan_id, [f.model_dump() for f in body.fills])


@router.post("/scans/{scan_id}/corrections/format")
async def apply_format(scan_id: str, body: FormatCorrections, user: User = Depends(require_user)) -> dict:
    svc = _require_scan(scan_id)
    return svc.apply_format(
        scan_id,
        [c.model_dump() for c in body.cell_corrections],
        [s.model_dump() for s in body.split_rows],
    )
