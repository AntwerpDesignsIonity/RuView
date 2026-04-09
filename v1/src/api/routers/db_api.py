"""
Database-backed REST+WebSocket API endpoints for the sensing database.

Adds to the existing FastAPI app:
  GET  /api/v1/db/readings       — query readings (filters: last_minutes, motion, presence)
  GET  /api/v1/db/telemetry      — query ESP32 telemetry
  GET  /api/v1/db/sessions       — list sessions
  GET  /api/v1/db/sessions/{id}  — session stats
  GET  /api/v1/db/devices        — list devices with online time
  GET  /api/v1/db/export/csv     — export readings as CSV download
  GET  /api/v1/db/export/tensor  — export training windows as .pt
  POST /api/v1/db/layouts        — create/update room layout
  GET  /api/v1/db/layouts        — list layouts

All responses include security headers (CSP, HSTS, X-Content-Type, X-Frame).
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Query, Response
from fastapi.responses import FileResponse, JSONResponse

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/v1/db", tags=["database"])

# Lazy singleton — initialized when the sensing server starts
_db = None


def get_db():
    """Get or create the global SensingDB singleton."""
    global _db
    if _db is None:
        from v1.src.database.sensing_db import SensingDB
        _db = SensingDB()
    return _db


def _secure_response(data, status_code: int = 200) -> JSONResponse:
    """Wrap response with security headers."""
    resp = JSONResponse(content=data, status_code=status_code)
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["X-XSS-Protection"] = "1; mode=block"
    resp.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    resp.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'"
    resp.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return resp


# ---------------------------------------------------------------
# Readings
# ---------------------------------------------------------------
@router.get("/readings")
async def get_readings(
    last_minutes: Optional[float] = Query(None, ge=0.1, le=10080, description="Last N minutes"),
    session_id: Optional[str] = Query(None, max_length=36),
    device_id: Optional[str] = Query(None, max_length=16),
    motion_level: Optional[str] = Query(None, pattern="^(absent|idle|active|rapid)$"),
    presence_only: bool = Query(False),
    limit: int = Query(1000, ge=1, le=50000),
):
    db = get_db()
    rows = db.query_readings(
        session_id=session_id,
        device_id=device_id,
        last_minutes=last_minutes,
        limit=limit,
        motion_level=motion_level,
        presence_only=presence_only,
    )
    return _secure_response({"count": len(rows), "readings": rows})


# ---------------------------------------------------------------
# Telemetry
# ---------------------------------------------------------------
@router.get("/telemetry")
async def get_telemetry(
    last_minutes: Optional[float] = Query(None, ge=0.1, le=10080),
    device_id: Optional[str] = Query(None, max_length=16),
    limit: int = Query(500, ge=1, le=10000),
):
    db = get_db()
    rows = db.query_telemetry(device_id=device_id, last_minutes=last_minutes, limit=limit)
    return _secure_response({"count": len(rows), "telemetry": rows})


# ---------------------------------------------------------------
# Sessions
# ---------------------------------------------------------------
@router.get("/sessions")
async def list_sessions():
    db = get_db()
    rows = db._conn.execute(
        "SELECT * FROM sessions ORDER BY started_utc DESC LIMIT 100"
    ).fetchall()
    return _secure_response({"sessions": [dict(r) for r in rows]})


@router.get("/sessions/{session_id}")
async def get_session_stats(session_id: str):
    if len(session_id) > 36:
        return _secure_response({"error": "invalid session_id"}, 400)
    db = get_db()
    stats = db.get_session_stats(session_id)
    return _secure_response(stats)


# ---------------------------------------------------------------
# Devices
# ---------------------------------------------------------------
@router.get("/devices")
async def list_devices():
    db = get_db()
    rows = db._conn.execute("SELECT * FROM devices ORDER BY last_seen_utc DESC").fetchall()
    devices = []
    for r in rows:
        d = dict(r)
        d["online_time"] = db.get_device_online_time(d["device_id"])
        devices.append(d)
    return _secure_response({"devices": devices})


# ---------------------------------------------------------------
# Exports
# ---------------------------------------------------------------
@router.get("/export/csv")
async def export_csv(
    last_minutes: Optional[float] = Query(60, ge=1, le=10080),
    session_id: Optional[str] = Query(None, max_length=36),
):
    db = get_db()
    path = db.export_csv(last_minutes=last_minutes, session_id=session_id)
    return FileResponse(
        path,
        media_type="text/csv",
        filename=Path(path).name,
        headers={"X-Content-Type-Options": "nosniff"},
    )


@router.get("/export/tensor")
async def export_tensor(
    window_size: int = Query(20, ge=5, le=200),
    step: int = Query(10, ge=1, le=100),
    session_id: Optional[str] = Query(None, max_length=36),
):
    db = get_db()
    out_dir = db.export_training_windows(
        window_size=window_size, step=step, session_id=session_id,
    )
    return _secure_response({
        "status": "exported",
        "output_dir": out_dir,
        "message": "Training windows saved as .pt files",
    })


# ---------------------------------------------------------------
# Layouts
# ---------------------------------------------------------------
@router.get("/layouts")
async def list_layouts():
    db = get_db()
    rows = db._conn.execute("SELECT * FROM layouts ORDER BY updated_utc DESC").fetchall()
    return _secure_response({"layouts": [dict(r) for r in rows]})


@router.post("/layouts")
async def create_layout(
    name: str = Query(..., min_length=1, max_length=100),
    width_m: float = Query(10.0, ge=1, le=1000),
    height_m: float = Query(10.0, ge=1, le=1000),
    floor: int = Query(0, ge=-10, le=200),
    description: Optional[str] = Query(None, max_length=500),
):
    db = get_db()
    lid = db.upsert_layout(
        name=name, width_m=width_m, height_m=height_m,
        floor=floor, description=description,
    )
    return _secure_response({"layout_id": lid, "status": "created"}, 201)
