"""
FastAPI web server for Input Studio.
Provides REST API endpoints matching the desktop app functionality.
"""
from __future__ import annotations

import os
import sys
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, UploadFile, HTTPException, Depends, Query
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
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
    """Root endpoint."""
    return {"message": "Input Studio Web API"}


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
    
    # Read PDF data
    pdf_data = await pdf_file.read()
    
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


# Additional endpoints
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

if ui_dir.exists():
    app.mount("/ui", StaticFiles(directory=str(ui_dir), html=True), name="ui")
    
    # Serve web_api.js from server directory
    @app.get("/web_api.js")
    async def web_api_js():
        web_api_path = server_dir / "web_api.js"
        if web_api_path.exists():
            return FileResponse(web_api_path, media_type="application/javascript")
        raise HTTPException(status_code=404, detail="web_api.js not found")
    
    app.mount("/", StaticFiles(directory=str(ui_dir), html=True), name="root")


if __name__ == "__main__":
    # Ensure directories exist
    _ensure_dirs()
    
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
