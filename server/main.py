"""
FastAPI web server for Input Studio.
Provides REST API endpoints matching the desktop app functionality.
"""
from __future__ import annotations

import io
import os
import sys
import uuid
import json
import time
import asyncio
import zipfile
import threading
from collections import deque
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, UploadFile, HTTPException, Depends, Query, Body, Request
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

# Add parent directory to path to import app.py
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import the existing Api class
from app import Api, ROOT, _ensure_dirs, LOCAL, PROJECTS_DIR

app = FastAPI(title="Input Studio Web API")

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify actual origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Session storage (in production, use Redis or database)
_sessions: dict[str, Api] = {}

# ---- Cost guardrails for public web deployment ----
_RATE_LIMIT_WINDOW_SEC = int(os.environ.get("INPUTSTUDIO_RATE_WINDOW_SEC", "60"))
# Default relaxed for single-tenant VPS operation
_RATE_LIMIT_REQUESTS = int(os.environ.get("INPUTSTUDIO_RATE_REQUESTS", "600"))
_MAX_UPLOAD_BYTES = int(os.environ.get("INPUTSTUDIO_MAX_UPLOAD_MB", "200")) * 1024 * 1024
_HEAVY_CONCURRENCY = max(1, int(os.environ.get("INPUTSTUDIO_MAX_HEAVY_CONCURRENCY", "6")))
_HEAVY_WAIT_TIMEOUT_SEC = float(os.environ.get("INPUTSTUDIO_HEAVY_WAIT_TIMEOUT_SEC", "15.0"))

_rate_lock = threading.Lock()
_rate_buckets: dict[str, deque[float]] = {}
_heavy_semaphore = asyncio.Semaphore(_HEAVY_CONCURRENCY)

_UPLOAD_PATH_HINTS = (
    "/api/projects/create",
    "/api/projects/upload",
    "/api/upload-project-zip",
    "/append-pdf",
)
_HEAVY_PATH_HINTS = (
    "/api/projects/create",
    "/api/projects/upload",
    "/api/upload-project-zip",
    "/save",
    "/export",
    "/append-pdf",
    "/copy-page",
    "/delete-page",
    "/reorder-pages",
)


def _client_ip(request: Request) -> str:
    xff = str(request.headers.get("x-forwarded-for") or "").strip()
    if xff:
        return xff.split(",")[0].strip() or "unknown"
    if request.client and request.client.host:
        return str(request.client.host)
    return "unknown"


def _is_upload_path(path: str) -> bool:
    return any(hint in path for hint in _UPLOAD_PATH_HINTS)


def _is_heavy_path(path: str) -> bool:
    return any(hint in path for hint in _HEAVY_PATH_HINTS)


def _is_rate_limited(ip: str) -> bool:
    now = time.monotonic()
    with _rate_lock:
        bucket = _rate_buckets.get(ip)
        if bucket is None:
            bucket = deque()
            _rate_buckets[ip] = bucket
        while bucket and (now - bucket[0]) > _RATE_LIMIT_WINDOW_SEC:
            bucket.popleft()
        if len(bucket) >= _RATE_LIMIT_REQUESTS:
            return True
        bucket.append(now)
        return False


async def _read_upload_limited(upload_file: UploadFile, label: str) -> bytes:
    total = 0
    chunks: list[bytes] = []
    chunk_size = 1024 * 1024
    while True:
        chunk = await upload_file.read(chunk_size)
        if not chunk:
            break
        total += len(chunk)
        if total > _MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"{label}が大きすぎます。{_MAX_UPLOAD_BYTES // (1024 * 1024)}MB以下にしてください。",
            )
        chunks.append(chunk)
    return b"".join(chunks)


@app.middleware("http")
async def protect_api_resources(request: Request, call_next):
    path = str(request.url.path or "")
    if not path.startswith("/api/"):
        return await call_next(request)

    if _is_rate_limited(_client_ip(request)):
        return JSONResponse(
            status_code=429,
            content={"code": "RATE_LIMITED", "detail": "リクエストが多すぎます。少し待ってから再試行してください。"},
        )

    if request.method in ("POST", "PUT", "PATCH") and _is_upload_path(path):
        content_length = request.headers.get("content-length")
        if content_length:
            try:
                if int(content_length) > _MAX_UPLOAD_BYTES:
                    return JSONResponse(
                        status_code=413,
                        content={
                            "code": "UPLOAD_TOO_LARGE",
                            "detail": f"アップロードサイズ上限は{_MAX_UPLOAD_BYTES // (1024 * 1024)}MBです。",
                        },
                    )
            except Exception:
                pass

    acquired = False
    if _is_heavy_path(path):
        try:
            await asyncio.wait_for(_heavy_semaphore.acquire(), timeout=_HEAVY_WAIT_TIMEOUT_SEC)
            acquired = True
        except asyncio.TimeoutError:
            return JSONResponse(
                status_code=503,
                content={"code": "SERVER_BUSY", "detail": "現在アクセスが集中しています。少し待って再試行してください。"},
            )

    try:
        return await call_next(request)
    finally:
        if acquired:
            _heavy_semaphore.release()


def get_or_create_session(session_id: str | None = None) -> tuple[str, Api]:
    """Get or create a session."""
    if session_id and session_id in _sessions:
        return session_id, _sessions[session_id]
    
    # Create new session
    new_session_id = str(uuid.uuid4())
    api = Api()
    _sessions[new_session_id] = api
    return new_session_id, api


@app.get("/")
async def root():
    """Root endpoint - serve index.html."""
    ui_dir = ROOT / "ui"
    index_path = ui_dir / "index.html"
    if index_path.exists():
        return FileResponse(str(index_path), media_type="text/html")
    return {"message": "Input Studio Web API - UI not found"}


@app.post("/api/session")
async def create_session():
    """Create a new session."""
    session_id, api = get_or_create_session()
    return {"session_id": session_id}


@app.post("/api/projects/create")
async def create_project(
    pdf_file: UploadFile = File(...),
    session_id: str = Query(...),
):
    """Create a new project from uploaded PDF."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Read PDF data (size-limited)
    pdf_data = await _read_upload_limited(pdf_file, "PDF")
    
    # Save PDF temporarily and create project
    try:
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_data)
            tmp_path = tmp.name
        
        result = api.create_project_from_pdf_simple(tmp_path)
        
        # Clean up temp file
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        
        if not result.get("ok"):
            raise HTTPException(status_code=400, detail=result.get("errors", ["Failed to create project"]))
        
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/projects/upload")
async def upload_project(
    project_file: UploadFile = File(...),
    pdf_file: UploadFile | None = File(None),
    session_id: str = Query(...),
):
    """Upload a project JSON file and create project on server. Optionally include PDF."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    try:
        # Read project JSON (size-limited)
        project_data = await _read_upload_limited(project_file, "プロジェクトJSON")
        project_json = json.loads(project_data.decode('utf-8'))
        
        # Extract project name
        project_name = project_json.get("project", "project")
        
        # Create project directory
        import time
        import uuid
        from app import _safe_name, _write_json, _now_iso, PROJECTS_DIR
        
        stem = _safe_name(project_name)
        pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}-{stem}"
        proj_dir = PROJECTS_DIR / pid
        proj_dir.mkdir(parents=True, exist_ok=True)
        
        # Save project.json
        proj_json = proj_dir / "project.json"
        project_json["updated_at"] = _now_iso()
        project_json["pdf"] = "template.pdf"
        _write_json(proj_json, project_json)
        
        if pdf_file and pdf_file.filename:
            pdf_data = await _read_upload_limited(pdf_file, "PDF")
            pdf_path = proj_dir / "template.pdf"
            with open(pdf_path, "wb") as f:
                f.write(pdf_data)
        
        return {
            "ok": True,
            "path": str(proj_json),
            "project_id": pid,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/upload-project-zip")
async def upload_project_zip(
    zip_file: UploadFile = File(...),
    session_id: str = Query(...),
):
    """Upload a project ZIP (project.json + template.pdf bundled)."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    if not zip_file.filename or not zip_file.filename.lower().endswith(".zip"):
        raise HTTPException(status_code=400, detail="ZIPファイルを選択してください")
    try:
        from app import _safe_name, _write_json, _now_iso
        zip_data = await _read_upload_limited(zip_file, "ZIP")
        if not zip_data:
            raise HTTPException(status_code=400, detail="ZIPファイルが空です")
        if len(zip_data) < 4:
            raise HTTPException(status_code=400, detail="ファイルが短すぎます。ZIPファイルを選択してください。")
        if zip_data[:2] != b"PK":
            if zip_data.lstrip()[:1] == b"<":
                raise HTTPException(status_code=400, detail="ZIPファイルではありません。別のファイルを選択してください。")
            raise HTTPException(status_code=400, detail="無効なZIPファイルです。正しいZIP形式の案件ファイルを選択してください。")
        with zipfile.ZipFile(io.BytesIO(zip_data), "r") as zf:
            names = zf.namelist()
            project_json_name = next((n for n in names if n.replace("\\", "/").lower().endswith("project.json")), None)
            if not project_json_name:
                raise HTTPException(status_code=400, detail="ZIPにproject.jsonが含まれていません")
            project_json = json.loads(zf.read(project_json_name).decode("utf-8"))
        project_name = project_json.get("project", "project")
        stem = _safe_name(project_name)
        pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}-{stem}"
        proj_dir = PROJECTS_DIR / pid
        proj_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(zip_data), "r") as zf:
            for name in names:
                if name.endswith("/") or ".." in name:
                    continue
                target = proj_dir / name.replace("\\", "/")
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(zf.read(name))
        proj_json = proj_dir / "project.json"
        project_json["updated_at"] = _now_iso()
        project_json["pdf"] = project_json.get("pdf", "template.pdf")
        _write_json(proj_json, project_json)
        return {"ok": True, "path": str(proj_json), "project_id": pid}
    except HTTPException:
        raise
    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="ZIPファイルが壊れているか、形式が正しくありません。「プロジェクトを保存」で作成したZIPを使用してください。")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/projects/{project_id}/load")
async def load_project(project_id: str, session_id: str = Query(...)):
    """Load a project."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Find project directory
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    result = api.load_project(str(proj_json))
    return result


@app.get("/api/projects/{project_id}/preview/{page_index}")
async def get_preview(
    project_id: str,
    page_index: int,
    session_id: str = Query(...),
):
    """Get preview PNG for a page."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    result = api.get_preview_png_base64_page(page_index)
    if not result.get("ok"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to generate preview"))
    
    return result


@app.post("/api/projects/{project_id}/save")
async def save_project(
    project_id: str,
    session_id: str = Query(...),
    make_filled_pdf: bool = Query(False),
):
    """Save current project."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    result = api.save_current_project(make_filled_pdf=make_filled_pdf)
    return result


@app.post("/api/projects/{project_id}/values")
async def set_value(
    project_id: str,
    tag: str = Query(...),
    value: str = Query(...),
    session_id: str = Query(...),
):
    """Set a tag value."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    result = api.set_value(tag, value)
    return result


@app.post("/api/projects/{project_id}/placements")
async def add_text_field(
    project_id: str,
    tag: str = Query(...),
    page: int = Query(...),
    x: float = Query(...),
    y: float = Query(...),
    font_size: int = Query(...),
    session_id: str = Query(...),
):
    """Add a text field placement."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    result = api.add_text_field(tag, page, x, y, font_size)
    return result


@app.get("/api/projects/{project_id}/export")
async def export_pdf(
    project_id: str,
    session_id: str = Query(...),
):
    """Export filled PDF."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    # Save and export
    result = api.save_current_project(make_filled_pdf=True)
    if not result.get("ok"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to export PDF"))
    
    pdf_path = result.get("filled_pdf") or result.get("pdf")
    if not pdf_path or not Path(pdf_path).exists():
        raise HTTPException(status_code=500, detail="PDF file not found")
    
    return FileResponse(
        pdf_path,
        media_type="application/pdf",
        filename=f"{project_id}_filled.pdf",
    )


@app.get("/api/projects/{project_id}/export-json")
async def export_project_json(
    project_id: str,
    session_id: str = Query(...),
):
    """Export project.json for save-as / project save."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    api.load_project(str(proj_json))
    result = api.save_current_project(make_filled_pdf=False)
    if not result.get("ok"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to save project"))
    with open(proj_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    return JSONResponse(content=data)


@app.get("/api/projects/{project_id}/export-zip")
async def export_project_zip(
    project_id: str,
    session_id: str = Query(...),
):
    """Export project as ZIP (project.json + template.pdf + sources)."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    api.load_project(str(proj_json))
    result = api.save_current_project(make_filled_pdf=False)
    if not result.get("ok"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to save project"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(str(proj_json), "project.json")
        proj_data = json.loads(proj_json.read_text(encoding="utf-8"))
        pdf_name = proj_data.get("pdf", "template.pdf")
        pdf_path = proj_dir / pdf_name
        if pdf_path.exists():
            zf.write(str(pdf_path), pdf_name)
        for sub in ("sources",):
            src_dir = proj_dir / sub
            if src_dir.exists() and src_dir.is_dir():
                for fp in src_dir.rglob("*"):
                    if fp.is_file():
                        zf.write(str(fp), f"{sub}/{fp.relative_to(src_dir).as_posix()}")
    buf.seek(0)
    zip_bytes = buf.getvalue()
    if len(zip_bytes) < 100:
        raise HTTPException(status_code=500, detail="ZIPの作成に失敗しました（空のZIP）")
    stem = proj_data.get("project", "project") or "project"
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in stem)[:64]
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{safe}.zip"'},
    )


def _ensure_project_has_pdf(proj_dir: Path, project_json: dict) -> None:
    """Raise HTTPException if project has no template PDF."""
    pdf_name = project_json.get("pdf", "template.pdf")
    pdf_path = proj_dir / pdf_name
    if not pdf_path.exists():
        raise HTTPException(
            status_code=400,
            detail="この案件にはPDFが含まれていません。PDFから新規で作成した案件でお試しください。",
        )


@app.post("/api/projects/{project_id}/append-pdf")
async def append_pdf_to_project(
    project_id: str,
    pdf_file: UploadFile = File(...),
    session_id: str = Query(...),
):
    """Append PDF pages to current project."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    with open(proj_json, "r", encoding="utf-8") as f:
        proj_data = json.load(f)
    _ensure_project_has_pdf(proj_dir, proj_data)
    api.load_project(str(proj_json))
    try:
        import tempfile
        pdf_data = await _read_upload_limited(pdf_file, "PDF")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_data)
            tmp_path = tmp.name
        result = api.append_pdf_to_project(tmp_path)
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        if not result.get("ok"):
            raise HTTPException(status_code=400, detail=result.get("error", "Failed to append PDF"))
        return result
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/projects/{project_id}/copy-page")
async def copy_page(
    project_id: str,
    page_index: int = Query(...),
    session_id: str = Query(...),
):
    """Copy current page with elements."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    with open(proj_json, "r", encoding="utf-8") as f:
        proj_data = json.load(f)
    _ensure_project_has_pdf(proj_dir, proj_data)
    api.load_project(str(proj_json))
    result = api.copy_page_with_elements(page_index)
    if not result.get("ok"):
        raise HTTPException(status_code=400, detail=result.get("error", "Failed to copy page"))
    return result


@app.post("/api/projects/{project_id}/delete-page")
async def delete_page(
    project_id: str,
    page_index: int = Query(...),
    session_id: str = Query(...),
):
    """Delete current page."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    with open(proj_json, "r", encoding="utf-8") as f:
        proj_data = json.load(f)
    _ensure_project_has_pdf(proj_dir, proj_data)
    api.load_project(str(proj_json))
    result = api.delete_page_from_project(page_index)
    if not result.get("ok"):
        raise HTTPException(status_code=400, detail=result.get("error", "Failed to delete page"))
    return result


@app.post("/api/projects/{project_id}/reorder-pages")
async def reorder_pages(
    project_id: str,
    body: dict = Body(...),
    session_id: str = Query(...),
):
    """Reorder pages."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    api = _sessions[session_id]
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    with open(proj_json, "r", encoding="utf-8") as f:
        proj_data = json.load(f)
    _ensure_project_has_pdf(proj_dir, proj_data)
    api.load_project(str(proj_json))
    order = body.get("order") if isinstance(body, dict) else None
    if not isinstance(order, list):
        raise HTTPException(status_code=400, detail={"code": "INVALID_ORDER", "detail": "order は配列で指定してください"})
    result = api.reorder_pages(order)
    if not result.get("ok"):
        raise HTTPException(status_code=400, detail=result.get("error", "Failed to reorder pages"))
    return result


# Additional endpoints
class ReorderPagesBody(BaseModel):
    order: list[int]


class PlacementUpdate(BaseModel):
    tag: str | None = None
    page: int | None = None
    x: float | None = None
    y: float | None = None
    font_size: int | None = None
    color: str | None = None
    line_height: float | None = None
    letter_spacing: float | None = None
    writing_mode: str | None = None


@app.post("/api/projects/{project_id}/placements/{fid}")
async def update_placement(
    project_id: str,
    fid: str,
    patch: PlacementUpdate,
    session_id: str = Query(...),
):
    """Update a placement."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    patch_dict = patch.dict(exclude_none=True)
    result = api.update_placement(fid, patch_dict)
    return result


@app.post("/api/projects/{project_id}/placements/{fid}/position")
async def set_element_pos(
    project_id: str,
    fid: str,
    x: float = Query(...),
    y: float = Query(...),
    session_id: str = Query(...),
):
    """Set element position."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    result = api.set_element_pos(fid, x, y)
    return result


@app.get("/api/projects/{project_id}/elements/{fid}")
async def get_element_info(
    project_id: str,
    fid: str,
    session_id: str = Query(...),
):
    """Get element info."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    result = api.get_element_info(fid)
    return result


@app.delete("/api/projects/{project_id}/elements")
async def delete_elements(
    project_id: str,
    fids: list[str] = Query(...),
    session_id: str = Query(...),
):
    """Delete elements."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    result = api.delete_elements(fids)
    return result


class ProjectPayload(BaseModel):
    tags: list[str] | None = None
    values: dict[str, str] | None = None
    placements: dict[str, dict] | None = None


@app.post("/api/projects/{project_id}/payload")
async def set_project_payload(
    project_id: str,
    payload: ProjectPayload,
    session_id: str = Query(...),
):
    """Set project payload."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    
    # Ensure project is loaded
    proj_dir = PROJECTS_DIR / project_id
    proj_json = proj_dir / "project.json"
    if not proj_json.exists():
        raise HTTPException(status_code=404, detail="Project not found")
    
    api.load_project(str(proj_json))
    
    payload_dict = payload.dict(exclude_none=True)
    result = api.set_project_payload(payload_dict)
    return result


@app.get("/api/workers")
async def get_workers(session_id: str = Query(...)):
    """Get workers."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    result = api.get_workers()
    return result


@app.get("/api/admin/settings")
async def get_admin_settings(session_id: str = Query(...)):
    """Get admin settings."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    result = api.get_admin_settings()
    return result


class SettingsPatch(BaseModel):
    ui_mode: str | None = None
    default_font_size: int | None = None
    view_zoom: float | None = None


@app.post("/api/admin/settings")
async def update_admin_settings(
    patch: SettingsPatch,
    session_id: str = Query(...),
):
    """Update admin settings."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    patch_dict = patch.dict(exclude_none=True)
    result = api.update_admin_settings(patch_dict)
    return result


@app.post("/api/ui/mode")
async def set_ui_mode(
    mode: str = Query(...),
    session_id: str = Query(...),
):
    """Set UI mode."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    result = api.set_ui_mode(mode)
    return result


@app.post("/api/work/start")
async def start_work(
    worker_id: str = Query(...),
    session_id: str = Query(...),
):
    """Start work."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    result = api.start_work(worker_id)
    return result


@app.post("/api/work/private/toggle")
async def toggle_private(session_id: str = Query(...)):
    """Toggle private mode."""
    if session_id not in _sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    api = _sessions[session_id]
    result = api.toggle_private()
    return result


# Mount static files (UI and web_api.js)
ui_dir = ROOT / "ui"
server_dir = Path(__file__).parent

# Serve web_api.js from server directory
@app.get("/web_api.js")
async def web_api_js():
    web_api_path = server_dir / "web_api.js"
    if web_api_path.exists():
        response = FileResponse(web_api_path, media_type="application/javascript")
        # Prevent caching during development
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response
    raise HTTPException(status_code=404, detail="web_api.js not found")


@app.get("/ad-config.js")
async def ad_config_js():
    """Serve client-side ad configuration from environment variables."""
    enabled = str(os.environ.get("INPUTSTUDIO_ADS_ENABLED", "0")).strip().lower() in ("1", "true", "yes", "on")
    provider = str(os.environ.get("INPUTSTUDIO_AD_PROVIDER", "adsense")).strip().lower() or "adsense"
    config = {
        "enabled": enabled,
        "provider": provider,
        "adsense": {
            "client": str(os.environ.get("INPUTSTUDIO_ADSENSE_CLIENT", "")).strip(),
            "slots": {
                "gate": str(os.environ.get("INPUTSTUDIO_AD_SLOT_GATE", "")).strip(),
                "panel": str(os.environ.get("INPUTSTUDIO_AD_SLOT_PANEL", "")).strip(),
                "panelBottom": str(os.environ.get("INPUTSTUDIO_AD_SLOT_PANEL_BOTTOM", "")).strip(),
                "unlock": str(os.environ.get("INPUTSTUDIO_AD_SLOT_UNLOCK", "")).strip(),
            },
        },
        "unlock": {
            "minSeconds": max(1, int(os.environ.get("INPUTSTUDIO_UNLOCK_AD_SECONDS", "3"))),
        },
    }
    body = "window.__INPUTSTUDIO_AD_CONFIG__ = " + json.dumps(config, ensure_ascii=False) + ";"
    response = Response(content=body, media_type="application/javascript")
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

# Mount static files for UI assets (CSS, JS, images, etc.)
# Note: API routes must be defined BEFORE static file mounting
if ui_dir.exists():
    # Mount UI static files - this will serve files from ui directory
    # Files like app.js, styles.css will be served automatically
    app.mount("/static", StaticFiles(directory=str(ui_dir)), name="static")
    
    # Serve static files from root (for app.js, styles.css, etc.)
    @app.get("/{file_path:path}")
    async def serve_static(file_path: str):
        """Serve static files from ui directory."""
        # Skip API routes
        if file_path.startswith("api/"):
            raise HTTPException(status_code=404, detail="Not found")
        
        # Handle root and index.html
        if file_path == "" or file_path == "index.html":
            index_path = ui_dir / "index.html"
            if index_path.exists():
                return FileResponse(str(index_path), media_type="text/html")
            raise HTTPException(status_code=404, detail="index.html not found")
        
        # Serve other static files
        file_path_obj = Path(file_path)
        # Security: prevent directory traversal
        if ".." in str(file_path_obj) or str(file_path_obj).startswith("/"):
            raise HTTPException(status_code=403, detail="Forbidden")
        
        full_path = ui_dir / file_path_obj
        if full_path.exists() and full_path.is_file() and str(full_path).startswith(str(ui_dir)):
            # Determine media type
            media_types = {
                ".css": "text/css",
                ".js": "application/javascript",
                ".html": "text/html",
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".svg": "image/svg+xml",
                ".json": "application/json",
            }
            ext = full_path.suffix.lower()
            media_type = media_types.get(ext, "application/octet-stream")
            
            return FileResponse(str(full_path), media_type=media_type)
        
        raise HTTPException(status_code=404, detail="File not found")


if __name__ == "__main__":
    # Ensure directories exist
    _ensure_dirs()
    
    port = int(os.environ.get("PORT", 8001))  # Changed default to 8001 to avoid conflicts
    uvicorn.run(app, host="0.0.0.0", port=port)
