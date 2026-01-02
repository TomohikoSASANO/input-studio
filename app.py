from __future__ import annotations

import base64
import hashlib
import json
import shutil
import time
import uuid
import zipfile
import threading
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import webview
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

try:
    import fitz  # PyMuPDF (in-process PDF renderer)
except Exception:  # pragma: no cover
    fitz = None


ROOT = Path(__file__).resolve().parent
UI_DIR = ROOT / "ui"
LOCAL = ROOT / "_local_data"
PROJECTS_DIR = LOCAL / "projects"
WORKERS_PATH = LOCAL / "workers.json"
ADMIN_SETTINGS_PATH = LOCAL / "admin_settings.json"

# Render DPI for preview images and coordinate system in this app.
RENDER_DPI = 150


def _ensure_dirs() -> None:
    LOCAL.mkdir(parents=True, exist_ok=True)
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)


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


def _pick_file(title: str, filetypes: list[tuple[str, str]], initialdir: str | None = None) -> str | None:
    # tkinter is stdlib; make sure the dialog is visible on pywebview contexts.
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    try:
        root.lift()
        root.focus_force()
    except Exception:
        pass
    try:
        path = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
            initialdir=initialdir or str(Path.home()),
            parent=root,
        )
        return path or None
    finally:
        try:
            root.destroy()
        except Exception:
            pass


@dataclass
class LoadedProject:
    path: Path
    data: dict[str, Any]


class Api:
    def __init__(self) -> None:
        _ensure_dirs()
        self._project: LoadedProject | None = None
        self._ui_mode: str = "worker"
        self._last_dir: str | None = None
        self._working_worker_id: str | None = None
        self._private: bool = False
        self._page_count: int = 1
        self._render_lock = threading.Lock()
        self._page_cache: "OrderedDict[int, str]" = OrderedDict()
        self._cache_max_pages = 12
        self._fitz_doc = None
        self._fitz_pdf_path: str | None = None

    # --- dialogs ---
    def pick_project(self) -> dict[str, Any]:
        p = _pick_file("案件（プロジェクト）を開く", [("Project JSON", "*.json"), ("All", "*.*")], self._last_dir)
        if not p:
            return {"ok": False}
        self._last_dir = str(Path(p).resolve().parent)
        return {"ok": True, "path": str(Path(p).resolve())}

    def pick_pdf(self) -> dict[str, Any]:
        p = _pick_file("PDFを選択", [("PDF", "*.pdf"), ("All", "*.*")], self._last_dir)
        if not p:
            return {"ok": False}
        self._last_dir = str(Path(p).resolve().parent)
        return {"ok": True, "path": str(Path(p).resolve())}

    # --- projects ---
    def create_project_from_pdf_simple(self, pdf_path: str) -> dict[str, Any]:
        """
        Non-form PDF assumed.
        Create an empty project (no tags/placements), just bind the PDF for preview/export.
        """
        try:
            src = Path(pdf_path).resolve()
            if not src.exists():
                return {"ok": False, "errors": ["PDFが見つかりませんでした"]}

            stem = _safe_name(src.stem)
            pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}-{stem}"
            proj_dir = PROJECTS_DIR / pid
            proj_dir.mkdir(parents=True, exist_ok=True)
            pdf_dst = proj_dir / "template.pdf"
            shutil.copy2(src, pdf_dst)

            data: dict[str, Any] = {
                "project": stem,
                "pdf": str(pdf_dst.name),
                "dpi": RENDER_DPI,
                "ui_mode": "worker",
                "tags": [],
                "values": {},
                "placements": {},  # tag -> {page,x,y,font_size}
                "created_at": _now_iso(),
                "updated_at": _now_iso(),
            }
            proj_json = proj_dir / "project.json"
            _write_json(proj_json, data)
            return {"ok": True, "path": str(proj_json)}
        except Exception as e:
            return {"ok": False, "errors": [str(e)]}

    def load_project(self, path: str) -> dict[str, Any]:
        try:
            p = Path(path).resolve()
            if not p.exists():
                return {"ok": False, "error": "not_found"}
            data = _read_json(p, None)
            if not isinstance(data, dict):
                return {"ok": False, "error": "invalid_json"}
            self._project = LoadedProject(path=p, data=data)
            self._ui_mode = str(data.get("ui_mode") or "worker")
            pdf_path = str(self._pdf_path())

            # Open PDF once. If PyMuPDF is available, rendering stays in-process
            # (no poppler subprocess => no black window on page changes).
            try:
                if self._fitz_doc is not None:
                    try:
                        self._fitz_doc.close()
                    except Exception:
                        pass
                    self._fitz_doc = None
                if fitz is not None:
                    self._fitz_doc = fitz.open(pdf_path)
                    self._fitz_pdf_path = pdf_path
                    self._page_count = max(1, int(self._fitz_doc.page_count))
                else:
                    self._fitz_pdf_path = None
                    self._page_count = max(1, int(len(PdfReader(pdf_path).pages)))
            except Exception:
                self._fitz_doc = None
                self._fitz_pdf_path = None
                self._page_count = 1

            self._page_cache.clear()
            return {
                "ok": True,
                "project": data.get("project") or p.parent.name,
                "tags": list(data.get("tags") or []),
                "values": dict(data.get("values") or {}),
                "drop_dir": str((p.parent / "exports").resolve()),
                "ui_mode": self._ui_mode,
                "path": str(p),
                "page_count": self._page_count,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _cache_dir(self) -> Path:
        if not self._project:
            raise RuntimeError("no project")
        # NOTE: Keep cache path ASCII-only for WebView file:// reliability.
        # Project folders may include non-ASCII (e.g. Japanese) names.
        key = str(self._project.path.parent).encode("utf-8", errors="ignore")
        hid = hashlib.sha1(key).hexdigest()[:16]
        d = LOCAL / "_cache_pages" / hid
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _cache_png_path(self, page_index: int) -> Path:
        return self._cache_dir() / f"page_{int(page_index):04d}.png"

    def _file_url(self, path: Path, bust: bool = True) -> str:
        # Use file:// URL so we don't send huge base64 over the JS bridge.
        p = path.resolve()
        url = p.as_uri()
        if bust:
            url = f"{url}?t={int(time.time() * 1000)}"
        return url

    def _cache_get(self, page_index: int) -> str | None:
        try:
            if page_index in self._page_cache:
                val = self._page_cache.pop(page_index)
                self._page_cache[page_index] = val
                return val
        except Exception:
            return None
        return None

    def _cache_put(self, page_index: int, data_url: str) -> None:
        try:
            if page_index in self._page_cache:
                self._page_cache.pop(page_index, None)
            self._page_cache[page_index] = data_url
            while len(self._page_cache) > self._cache_max_pages:
                self._page_cache.popitem(last=False)
        except Exception:
            pass

    def _render_page_png_url(self, idx: int) -> tuple[str, int, int]:
        # disk cache first (instant + no huge bridge payload)
        cache_png = self._cache_png_path(idx)
        if cache_png.exists():
            w, h = self._page_image_size(idx)
            return self._file_url(cache_png, bust=True), w, h

        img = None
        # Preferred: in-process rendering (no external process / no black window)
        try:
            if self._fitz_doc is not None and fitz is not None:
                pi = int(idx)
                if pi < 0:
                    pi = 0
                if pi >= int(self._fitz_doc.page_count):
                    pi = int(self._fitz_doc.page_count) - 1
                page = self._fitz_doc.load_page(pi)
                scale = RENDER_DPI / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=True)
                b0 = pix.tobytes("png")

                import io
                from PIL import Image

                img = Image.open(io.BytesIO(b0)).convert("RGBA")
        except Exception:
            img = None

        # Fallback: pdf2image (may spawn poppler subprocess)
        if img is None:
            pdf = self._pdf_path()
            images = convert_from_path(
                str(pdf),
                dpi=RENDER_DPI,
                first_page=idx + 1,
                last_page=idx + 1,
            )
            if not images:
                raise RuntimeError("render_failed")
            img = images[0].convert("RGBA")

        # overlay
        try:
            from PIL import ImageDraw, ImageFont

            draw = ImageDraw.Draw(img)
            try:
                font_cache: dict[int, Any] = {}

                def font(sz: int):
                    if sz in font_cache:
                        return font_cache[sz]
                    try:
                        f = ImageFont.truetype("arial.ttf", sz)
                    except Exception:
                        f = ImageFont.load_default()
                    font_cache[sz] = f
                    return f

            except Exception:
                font = lambda sz: None  # type: ignore

            placements = dict(self._project.data.get("placements") or {})
            values = dict(self._project.data.get("values") or {})
            for k, p in placements.items():
                if not isinstance(p, dict):
                    continue
                if int(p.get("page") or 0) != idx:
                    continue
                text = str(values.get(k) or "").replace("<br>", "\n")
                if not text.strip():
                    continue
                x = float(p.get("x") or 0)
                y = float(p.get("y") or 0)
                fs = int(p.get("font_size") or 14)
                draw.text((x, y), text, fill=(10, 10, 10, 255), font=font(fs))
        except Exception:
            pass

        # Save to disk cache and return file URL.
        try:
            img.save(cache_png, format="PNG")
        except Exception:
            # fallback: still try to return as base64 if saving fails
            import io
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            b = buf.getvalue()
            w, h = img.size
            return f"data:image/png;base64,{base64.b64encode(b).decode('ascii')}", w, h
        w, h = img.size
        return self._file_url(cache_png, bust=True), w, h

    def get_preview_png_base64_page(self, page_index: int) -> dict[str, Any]:
        """
        Render an explicit page index with current overlays.
        """
        if not self._project:
            return {"ok": False, "error": "no_project"}
        try:
            idx = int(page_index or 0)
            if idx < 0:
                idx = 0
            if idx >= self._page_count:
                idx = self._page_count - 1

            hit = self._cache_get(idx)
            if hit:
                w, h = self._page_image_size(idx)
                return {"ok": True, "png": hit, "page_display_width": w, "page_display_height": h, "page_index": idx}

            with self._render_lock:
                hit2 = self._cache_get(idx)
                if hit2:
                    w, h = self._page_image_size(idx)
                    return {"ok": True, "png": hit2, "page_display_width": w, "page_display_height": h, "page_index": idx}
                png, w, h = self._render_page_png_url(idx)
                self._cache_put(idx, png)

            # prefetch neighbor pages in background (for fast rapid paging)
            def _prefetch(n: int) -> None:
                try:
                    if n < 0 or n >= self._page_count:
                        return
                    if self._cache_get(n):
                        return
                    with self._render_lock:
                        if self._cache_get(n):
                            return
                        png2, _, _ = self._render_page_png_url(n)
                        self._cache_put(n, png2)
                except Exception:
                    return

            for n in (idx + 1, idx - 1):
                threading.Thread(target=_prefetch, args=(n,), daemon=True).start()

            return {
                "ok": True,
                "png": png,
                "page_display_width": w,
                "page_display_height": h,
                "page_index": idx,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_current_project(self) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        try:
            self._project.data["ui_mode"] = self._ui_mode
            self._project.data["updated_at"] = _now_iso()
            _write_json(self._project.path, self._project.data)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # --- mode / workers ---
    def set_ui_mode(self, mode: str) -> dict[str, Any]:
        m = str(mode or "")
        if m not in ("admin", "worker"):
            return {"ok": False, "error": "invalid_mode"}
        self._ui_mode = m
        if self._project:
            self._project.data["ui_mode"] = m
            _write_json(self._project.path, self._project.data)
        return {"ok": True}

    def get_admin_settings(self) -> dict[str, Any]:
        s = _read_json(ADMIN_SETTINGS_PATH, {"ui_mode": "worker"})
        if not isinstance(s, dict):
            s = {"ui_mode": "worker"}
        return {"ok": True, "settings": s}

    def get_workers(self) -> dict[str, Any]:
        rows = _read_json(WORKERS_PATH, [])
        if not isinstance(rows, list):
            rows = []
        # minimal normalization
        workers = []
        for r in rows:
            if not isinstance(r, dict):
                continue
            if not r.get("id"):
                continue
            workers.append(r)
        last = workers[0]["id"] if workers else None
        return {"ok": True, "workers": workers, "last_worker_id": last}

    def upsert_worker(self, w: dict[str, Any]) -> dict[str, Any]:
        rows = _read_json(WORKERS_PATH, [])
        if not isinstance(rows, list):
            rows = []
        wid = str(w.get("id") or "") or f"w_{uuid.uuid4().hex[:8]}"
        item = {
            "id": wid,
            "name": str(w.get("name") or "").strip(),
            "bank": str(w.get("bank") or "").strip(),
            "hourly_yen": int(w.get("hourly_yen") or 0),
        }
        if not item["name"]:
            return {"ok": False, "error": "missing_name"}
        out = []
        replaced = False
        for r in rows:
            if isinstance(r, dict) and r.get("id") == wid:
                out.append(item)
                replaced = True
            else:
                out.append(r)
        if not replaced:
            out.insert(0, item)
        _write_json(WORKERS_PATH, out)
        return {"ok": True, "id": wid}

    # --- work session ---
    def start_work(self, worker_id: str) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        self._working_worker_id = str(worker_id or "")
        self._private = False
        return {"ok": True}

    def toggle_private(self) -> dict[str, Any]:
        self._private = not self._private
        return {"ok": True, "in_private": self._private}

    # --- tags / placements / values ---
    def add_text_field(self, tag: str, page: int, x: float, y: float, font_size: int) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        if not t:
            return {"ok": False, "error": "missing_tag"}
        data = self._project.data
        tags = list(data.get("tags") or [])
        if t not in tags:
            tags.append(t)
        data["tags"] = tags
        placements = dict(data.get("placements") or {})
        placements[t] = {"page": int(page or 0), "x": float(x), "y": float(y), "font_size": int(font_size or 14)}
        data["placements"] = placements
        _write_json(self._project.path, data)
        return {"ok": True, "tag": t}

    def set_element_pos(self, tag: str, x: float, y: float) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        placements = dict(self._project.data.get("placements") or {})
        if t not in placements:
            placements[t] = {"page": 0, "x": float(x), "y": float(y), "font_size": 14}
        else:
            placements[t]["x"] = float(x)
            placements[t]["y"] = float(y)
        self._project.data["placements"] = placements
        _write_json(self._project.path, self._project.data)
        return {"ok": True}

    def get_element_info(self, tag: str) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        pl = (self._project.data.get("placements") or {}).get(t)
        if not isinstance(pl, dict):
            return {"ok": False, "error": "not_found"}
        page = int(pl.get("page") or 0)
        w, h = self._page_image_size(page)
        return {
            "ok": True,
            "page": page,
            "x": float(pl.get("x") or 0),
            "y": float(pl.get("y") or 0),
            "font_size": int(pl.get("font_size") or 14),
            "page_display_width": w,
            "page_display_height": h,
        }

    def set_value(self, tag: str, value: str) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        values = dict(self._project.data.get("values") or {})
        values[t] = str(value or "")
        self._project.data["values"] = values
        _write_json(self._project.path, self._project.data)
        return {"ok": True}

    # --- preview / export ---
    def _pdf_path(self) -> Path:
        if not self._project:
            raise RuntimeError("no project")
        pdf_name = str(self._project.data.get("pdf") or "template.pdf")
        return (self._project.path.parent / pdf_name).resolve()

    def _page_image_size(self, page_index: int) -> tuple[int, int]:
        # Compute expected image size at our DPI without rendering full image each time.
        try:
            if self._fitz_doc is not None:
                pi = int(page_index)
                if pi < 0:
                    pi = 0
                if pi >= int(self._fitz_doc.page_count):
                    pi = int(self._fitz_doc.page_count) - 1
                page = self._fitz_doc.load_page(pi)
                r = page.rect  # points
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

    def get_preview_png_base64(self, tag: str) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        placements = dict(self._project.data.get("placements") or {})
        pl = placements.get(t)
        page_index = int(pl.get("page") if isinstance(pl, dict) else 0) if placements else 0
        # Route to page renderer so cache/prefetch & PyMuPDF path applies.
        return self.get_preview_png_base64_page(page_index)

    def finish(self) -> dict[str, Any]:
        if not self._project:
            return {"ok": False, "error": "no_project"}
        try:
            pdf_in = self._pdf_path()
            reader = PdfReader(str(pdf_in))
            placements = dict(self._project.data.get("placements") or {})
            values = dict(self._project.data.get("values") or {})

            out_dir = self._project.path.parent / "exports"
            out_dir.mkdir(parents=True, exist_ok=True)
            out_pdf = out_dir / f"filled-{int(time.time())}.pdf"

            writer = PdfWriter()
            for pi, page in enumerate(reader.pages):
                w_pt = float(page.mediabox.width)
                h_pt = float(page.mediabox.height)
                w_px = int(round(w_pt / 72.0 * RENDER_DPI))
                h_px = int(round(h_pt / 72.0 * RENDER_DPI))

                import io

                packet = io.BytesIO()
                c = canvas.Canvas(packet, pagesize=(w_pt, h_pt))
                for k, p in placements.items():
                    if not isinstance(p, dict):
                        continue
                    if int(p.get("page") or 0) != pi:
                        continue
                    text = str(values.get(k) or "").replace("<br>", "\n")
                    if not text.strip():
                        continue
                    x_px = float(p.get("x") or 0)
                    y_px = float(p.get("y") or 0)
                    fs = int(p.get("font_size") or 14)
                    x_pt = x_px * 72.0 / RENDER_DPI
                    y_pt = (h_px - y_px) * 72.0 / RENDER_DPI
                    c.setFont("Helvetica", fs)
                    for line_idx, line in enumerate(text.splitlines() or [""]):
                        c.drawString(x_pt, y_pt - (fs + 2) * line_idx, line)
                c.save()
                packet.seek(0)
                overlay = PdfReader(packet).pages[0]
                page.merge_page(overlay)
                writer.add_page(page)

            with out_pdf.open("wb") as f:
                writer.write(f)

            out_zip = out_dir / f"filled-{int(time.time())}.zip"
            with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(out_pdf, arcname=out_pdf.name)
            return {"ok": True, "dir": str(out_dir.resolve()), "zip": str(out_zip.resolve())}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def main() -> None:
    _ensure_dirs()
    api = Api()
    ui_index = (UI_DIR / "index.html").resolve()
    if not ui_index.exists():
        raise RuntimeError(f"UI not found: {ui_index}")
    window = webview.create_window(
        "Input Studio",
        url=str(ui_index),
        js_api=api,
        width=1280,
        height=820,
        x=60,
        y=40,
        resizable=True,
    )
    webview.start(debug=False)


if __name__ == "__main__":
    main()



