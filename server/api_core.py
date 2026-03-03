"""
Core API logic shared between desktop and web versions.
This module contains the business logic without UI-specific dependencies.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
import sys
import time
import uuid
import zipfile
import threading
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4

try:
    import fitz  # PyMuPDF (in-process PDF renderer)
except Exception:
    fitz = None


def _norm_writing_mode(v: Any) -> str:
    s = str(v or "").strip().lower()
    if s in ("vertical", "v", "tate", "縦", "縦書き"):
        return "vertical"
    return "horizontal"


def _draw_vertical_text(
    c: canvas.Canvas,
    *,
    x_pt: float,
    y_top_pt: float,
    text: str,
    font_name: str,
    fs_pt: float,
    line_h: float,
    letter_s_pt: float,
) -> None:
    """Simple tategaki approximation."""
    c.setFont(font_name, fs_pt)

    leading = float(fs_pt) * float(line_h or 1.2)
    step_y = leading + float(letter_s_pt or 0.0)
    step_x = float(fs_pt) * 1.10 + float(letter_s_pt or 0.0)

    cols = str(text or "").splitlines() or [""]
    baseline_shift_em = PDF_BASELINE_SHIFT_EM
    try:
        ascent = float(pdfmetrics.getAscent(font_name) or 0) / 1000.0 * fs_pt
    except Exception:
        ascent = fs_pt * 0.8
    y0 = float(y_top_pt) - float(ascent) - (baseline_shift_em * float(fs_pt))

    for col_idx, col in enumerate(cols):
        cx = float(x_pt) - (col_idx * step_x)
        cy = float(y0)
        for ch in col:
            if ch == "\r":
                continue
            if ch == " ":
                cy -= step_y
                continue
            if ord(ch) < 128:
                c.saveState()
                c.translate(cx, cy)
                c.rotate(90)
                c.setFont(font_name, fs_pt)
                c.drawString(0, -fs_pt * 0.3, ch)
                c.restoreState()
            else:
                c.setFont(font_name, fs_pt)
                c.drawString(cx, cy, ch)
            cy -= step_y


# Constants
RENDER_DPI = 150
PDF_BASELINE_SHIFT_EM = 0.26
PDF_LETTER_SPACING_FACTOR = 0.72
DEFAULT_LETTER_SPACING = 1.2


def _default_local_data_dir() -> Path:
    """Store user data directory."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if base:
        return Path(base) / "InputStudio" / "_local_data"
    return Path.home() / "AppData" / "Local" / "InputStudio" / "_local_data"


def _ensure_dirs(base_dir: Path) -> None:
    """Ensure required directories exist."""
    base_dir.mkdir(parents=True, exist_ok=True)
    (base_dir / "projects").mkdir(parents=True, exist_ok=True)


def _read_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _safe_name(s: str) -> str:
    keep = []
    for ch in s.strip():
        if ch.isalnum() or ch in ("-", "_", " ", "."):
            keep.append(ch)
        else:
            keep.append("_")
    out = "".join(keep).strip().replace(" ", "_")
    return out or "project"


@dataclass
class LoadedProject:
    path: Path
    data: dict[str, Any]


class ApiCore:
    """Core API logic without UI-specific dependencies."""
    
    def __init__(self, base_dir: Path, resource_root: Path) -> None:
        self._base_dir = base_dir
        self._resource_root = resource_root
        _ensure_dirs(base_dir)
        
        self._project: LoadedProject | None = None
        self._last_project_path: str | None = None
        self._ui_mode: str = "worker"
        self._working_worker_id: str | None = None
        self._private: bool = False
        self._page_count: int = 1
        self._render_lock = threading.Lock()
        self._page_cache: OrderedDict[int, str] = OrderedDict()
        self._cache_max_pages = 12
        self._fitz_doc = None
        self._fitz_pdf_path: str | None = None
        
        # Paths
        self._projects_dir = base_dir / "projects"
        self._workers_path = base_dir / "workers.json"
        self._admin_settings_path = base_dir / "admin_settings.json"
        
        # Font path
        self._bundled_font_path = (resource_root / "assets" / "fonts" / "NotoSansJP-Regular.ttf")
        if not self._bundled_font_path.exists():
            # Fallback to parent directory
            self._bundled_font_path = (Path(__file__).parent.parent / "assets" / "fonts" / "NotoSansJP-Regular.ttf")

    def create_project_from_pdf(self, pdf_data: bytes, pdf_filename: str) -> dict[str, Any]:
        """Create a new project from uploaded PDF."""
        try:
            stem = _safe_name(Path(pdf_filename).stem)
            pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}-{stem}"
            proj_dir = self._projects_dir / pid
            proj_dir.mkdir(parents=True, exist_ok=True)
            pdf_dst = proj_dir / "template.pdf"
            pdf_dst.write_bytes(pdf_data)

            data: dict[str, Any] = {
                "project": stem,
                "pdf": str(pdf_dst.name),
                "dpi": RENDER_DPI,
                "ui_mode": "worker",
                "tags": [],
                "values": {},
                "placements": {},
                "created_at": _now_iso(),
                "updated_at": _now_iso(),
            }
            proj_json = proj_dir / "project.json"
            _write_json(proj_json, data)
            self._last_project_path = str(proj_json)
            self._load_project_internal(str(proj_json))
            return {"ok": True, "path": str(proj_json), "project_id": pid}
        except Exception as e:
            return {"ok": False, "errors": [str(e)]}

    def _load_project_internal(self, path: str) -> bool:
        """Internal project loading."""
        try:
            p = Path(path)
            if not p.exists():
                return False
            data = _read_json(p, None)
            if not isinstance(data, dict):
                return False

            # Schema normalization
            changed = False
            placements0 = data.get("placements")
            if isinstance(placements0, dict):
                newp: dict[str, Any] = {}
                for k, v in placements0.items():
                    if isinstance(v, dict) and "tag" in v:
                        newp[str(k)] = v
                        continue
                    if isinstance(v, dict):
                        fid = f"f_{uuid.uuid4().hex[:8]}"
                        nv = dict(v)
                        nv["tag"] = str(k)
                        newp[fid] = nv
                        changed = True
                if changed:
                    data["placements"] = newp
            else:
                data["placements"] = {}

            tags0 = data.get("tags")
            tags_list: list[str] = [str(t) for t in tags0] if isinstance(tags0, list) else []
            tagset = {t for t in tags_list if t.strip()}
            for _, pl in (data.get("placements") or {}).items():
                if isinstance(pl, dict):
                    t = str(pl.get("tag") or "").strip()
                    if t and t not in tagset:
                        tags_list.append(t)
                        tagset.add(t)
                        changed = True
            data["tags"] = tags_list

            placements1 = data.get("placements") if isinstance(data.get("placements"), dict) else {}
            for _, pl in placements1.items():
                if not isinstance(pl, dict):
                    continue
                wm = _norm_writing_mode(pl.get("writing_mode"))
                if pl.get("writing_mode") != wm:
                    pl["writing_mode"] = wm
                    changed = True

            if changed:
                data["updated_at"] = _now_iso()
                try:
                    _write_json(p, data)
                except Exception:
                    pass

            self._project = LoadedProject(path=p, data=data)
            self._last_project_path = str(p)
            self._ui_mode = str(data.get("ui_mode") or "worker")
            pdf_path = self._pdf_path()

            try:
                if self._fitz_doc is not None:
                    try:
                        self._fitz_doc.close()
                    except Exception:
                        pass
                    self._fitz_doc = None
                if fitz is not None:
                    self._fitz_doc = fitz.open(str(pdf_path))
                    self._fitz_pdf_path = str(pdf_path)
                    self._page_count = max(1, int(self._fitz_doc.page_count))
                else:
                    self._fitz_pdf_path = None
                    self._page_count = max(1, int(len(PdfReader(str(pdf_path)).pages)))
            except Exception:
                self._fitz_doc = None
                self._fitz_pdf_path = None
                self._page_count = 1

            self._page_cache.clear()
            if changed:
                _write_json(self._project.path, self._project.data)
            return True
        except Exception:
            return False

    def load_project(self, path: str) -> dict[str, Any]:
        """Load a project."""
        if not self._load_project_internal(path):
            return {"ok": False, "error": "not_found"}
        
        data = self._project.data
        return {
            "ok": True,
            "project": data.get("project") or Path(path).parent.name,
            "tags": list(data.get("tags") or []),
            "values": dict(data.get("values") or {}),
            "placements": dict(data.get("placements") or {}),
            "ui_mode": self._ui_mode,
            "path": str(self._project.path),
            "page_count": self._page_count,
        }

    def _ensure_project_loaded(self) -> bool:
        """Ensure project is loaded."""
        if self._project:
            return True
        if not self._last_project_path:
            return False
        return self._load_project_internal(self._last_project_path)

    def _pdf_path(self) -> Path:
        if not self._project:
            raise RuntimeError("no project")
        pdf_name = str(self._project.data.get("pdf") or "template.pdf")
        return (self._project.path.parent / pdf_name).resolve()

    def _cache_dir(self) -> Path:
        if not self._project:
            raise RuntimeError("no project")
        key = str(self._project.path.parent).encode("utf-8", errors="ignore")
        hid = hashlib.sha1(key).hexdigest()[:16]
        d = self._base_dir / "_cache_pages" / hid
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _cache_png_path(self, page_index: int) -> Path:
        return self._cache_dir() / f"page_{int(page_index):04d}.png"

    def _page_image_size(self, page_index: int) -> tuple[int, int]:
        try:
            if self._fitz_doc is not None:
                pi = int(page_index)
                if pi < 0:
                    pi = 0
                if pi >= int(self._fitz_doc.page_count):
                    pi = int(self._fitz_doc.page_count) - 1
                page = self._fitz_doc.load_page(pi)
                r = page.rect
                w_px = int(round(float(r.width) / 72.0 * RENDER_DPI))
                h_px = int(round(float(r.height) / 72.0 * RENDER_DPI))
                return max(1, w_px), max(1, h_px)
            reader = PdfReader(str(self._pdf_path()))
            page = reader.pages[int(page_index)]
            w_pt = float(page.mediabox.width)
            h_pt = float(page.mediabox.height)
            w_px = int(round(w_pt / 72.0 * RENDER_DPI))
            h_px = int(round(h_pt / 72.0 * RENDER_DPI))
            return max(1, w_px), max(1, h_px)
        except Exception:
            return 600, 800

    def get_preview_png_base64_page(self, page_index: int) -> dict[str, Any]:
        """Get preview PNG for a page."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        
        # Implementation continues in next part...
        # This is a large method, will be completed in the next file
        return {"ok": False, "error": "not_implemented"}

    # Additional methods will be added...
    # For now, creating the FastAPI server structure
