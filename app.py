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
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import HexColor

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
        self._last_project_path: str | None = None
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
                "placements": {},  # fid -> {tag,page,x,y,font_size,...}
                "created_at": _now_iso(),
                "updated_at": _now_iso(),
            }
            proj_json = proj_dir / "project.json"
            _write_json(proj_json, data)
            self._last_project_path = str(proj_json.resolve())
            return {"ok": True, "path": str(proj_json)}
        except Exception as e:
            return {"ok": False, "errors": [str(e)]}

    def _ensure_project_loaded(self) -> bool:
        """Best-effort auto recovery when API is called before/after project is loaded."""
        if self._project:
            return True
        if not self._last_project_path:
            return False
        try:
            r = self.load_project(self._last_project_path)
            return bool(r.get("ok"))
        except Exception:
            return False

    def load_project(self, path: str) -> dict[str, Any]:
        try:
            p = Path(path).resolve()
            if not p.exists():
                return {"ok": False, "error": "not_found"}
            data = _read_json(p, None)
            if not isinstance(data, dict):
                return {"ok": False, "error": "invalid_json"}

            # ---- schema normalization / migration ----
            # Old schema: placements[tag] = {page,x,y,font_size,...}
            # New schema: placements[fid] = {tag, page,x,y,font_size,...}
            changed = False
            placements0 = data.get("placements")
            if isinstance(placements0, dict):
                newp: dict[str, Any] = {}
                for k, v in placements0.items():
                    if isinstance(v, dict) and "tag" in v:
                        # already new-style
                        newp[str(k)] = v
                        continue
                    if isinstance(v, dict):
                        fid = f"f_{uuid.uuid4().hex[:8]}"
                        nv = dict(v)
                        nv["tag"] = str(k)
                        newp[fid] = nv
                        changed = True
                # If we detected any old-style entries, migrate whole dict.
                if changed:
                    data["placements"] = newp
            else:
                data["placements"] = {}

            # Ensure tags list contains all placement tags (preserve order).
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

            self._project = LoadedProject(path=p, data=data)
            self._last_project_path = str(p)
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
            if changed:
                # Write back migrated/normalized schema so future loads are consistent.
                _write_json(self._project.path, self._project.data)
            return {
                "ok": True,
                "project": data.get("project") or p.parent.name,
                "tags": list(data.get("tags") or []),
                "values": dict(data.get("values") or {}),
                "placements": dict(data.get("placements") or {}),
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

    def _png_as_data_url(self, file_url: str) -> str | None:
        """Read a PNG file (given as file:// URL) and return base64 data URL. Used as fallback if WebView blocks file://."""
        try:
            from urllib.parse import urlparse, unquote

            parsed = urlparse(file_url)
            if parsed.scheme != "file":
                return None
            # file:// URL -> local path (handle Windows /C:/... form)
            p = unquote(parsed.path or "")
            if p.startswith("/") and len(p) >= 3 and p[2] == ":":
                p = p[1:]
            path = Path(p)
            data = path.read_bytes()
            return "data:image/png;base64," + base64.b64encode(data).decode("ascii")
        except Exception:
            return None

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

    def _invalidate_pages(self, pages: set[int] | None = None) -> None:
        """Invalidate cached preview PNGs for given pages (or all)."""
        try:
            if pages is None:
                self._page_cache.clear()
                d = self._cache_dir()
                for p in d.glob("page_*.png"):
                    try:
                        p.unlink()
                    except Exception:
                        pass
                return
            for pi in pages:
                try:
                    self._page_cache.pop(int(pi), None)
                except Exception:
                    pass
                try:
                    fp = self._cache_png_path(int(pi))
                    if fp.exists():
                        fp.unlink()
                except Exception:
                    pass
        except Exception:
            return

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
                        # Japanese text needs a JP-capable font; prefer Windows built-ins.
                        candidates = [
                            r"C:\Windows\Fonts\meiryo.ttc",
                            r"C:\Windows\Fonts\YuGothR.ttc",
                            r"C:\Windows\Fonts\YuGothM.ttc",
                            r"C:\Windows\Fonts\msgothic.ttc",
                            r"C:\Windows\Fonts\msmincho.ttc",
                            "arial.ttf",
                        ]
                        f = None
                        for fp in candidates:
                            try:
                                f = ImageFont.truetype(fp, sz)
                                break
                            except Exception:
                                f = None
                        if f is None:
                            raise RuntimeError("font_load_failed")
                    except Exception:
                        f = ImageFont.load_default()
                    font_cache[sz] = f
                    return f

            except Exception:
                font = lambda sz: None  # type: ignore

            def _hex_to_rgba(h: str) -> tuple[int, int, int, int]:
                s = (h or "").strip()
                if not s:
                    return (15, 23, 42, 255)
                if not s.startswith("#"):
                    s = "#" + s
                try:
                    if len(s) == 4:  # #rgb
                        r = int(s[1] * 2, 16)
                        g = int(s[2] * 2, 16)
                        b = int(s[3] * 2, 16)
                        return (r, g, b, 255)
                    if len(s) >= 7:
                        r = int(s[1:3], 16)
                        g = int(s[3:5], 16)
                        b = int(s[5:7], 16)
                        return (r, g, b, 255)
                except Exception:
                    pass
                return (15, 23, 42, 255)

            def _draw_text(draw2: Any, x: float, y: float, text: str, fs: int, fill: tuple[int, int, int, int], line_h: float, letter_s: float) -> None:
                fnt = font(fs)
                lines = text.split("\n")
                cy = float(y)
                for line in lines:
                    cx = float(x)
                    if letter_s and fnt is not None:
                        for ch in line:
                            draw2.text((cx, cy), ch, fill=fill, font=fnt)
                            try:
                                w = fnt.getlength(ch)  # type: ignore
                            except Exception:
                                try:
                                    w = draw2.textlength(ch, font=fnt)
                                except Exception:
                                    w = fs * 0.62
                            cx += float(w) + float(letter_s)
                    else:
                        draw2.text((cx, cy), line, fill=fill, font=fnt)
                    cy += float(fs) * float(line_h)

            placements = dict(self._project.data.get("placements") or {})
            values = dict(self._project.data.get("values") or {})
            for _, p in placements.items():
                if not isinstance(p, dict):
                    continue
                if int(p.get("page") or 0) != idx:
                    continue
                tag = str(p.get("tag") or "").strip()
                if not tag:
                    continue
                text = str(values.get(tag) or "").replace("<br>", "\n")
                if not text.strip():
                    continue
                x = float(p.get("x") or 0)
                y = float(p.get("y") or 0)
                fs = int(p.get("font_size") or 14)
                color = _hex_to_rgba(str(p.get("color") or "#0f172a"))
                line_h = float(p.get("line_height") or 1.2)
                letter_s = float(p.get("letter_spacing") or 0)
                _draw_text(draw, x, y, text, fs, color, line_h, letter_s)
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
        if not self._project and not self._ensure_project_loaded():
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
                return {
                    "ok": True,
                    "png": hit,
                    "png_data": self._png_as_data_url(hit),
                    "page_display_width": w,
                    "page_display_height": h,
                    "page_index": idx,
                }

            with self._render_lock:
                hit2 = self._cache_get(idx)
                if hit2:
                    w, h = self._page_image_size(idx)
                    return {
                        "ok": True,
                        "png": hit2,
                        "png_data": self._png_as_data_url(hit2),
                        "page_display_width": w,
                        "page_display_height": h,
                        "page_index": idx,
                    }
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
                "png_data": self._png_as_data_url(png),
                "page_display_width": w,
                "page_display_height": h,
                "page_index": idx,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_current_project(self) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            self._project.data["ui_mode"] = self._ui_mode
            self._project.data["updated_at"] = _now_iso()
            _write_json(self._project.path, self._project.data)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_project_as(self, name: str) -> dict[str, Any]:
        """
        Save as a new project (duplicate current project to a new folder) and load it.
        """
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            new_name = _safe_name(str(name or "").strip()) or "project"
            pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}-{new_name}"
            src_dir = self._project.path.parent
            dst_dir = PROJECTS_DIR / pid
            # Copy whole project folder, but skip exports
            shutil.copytree(src_dir, dst_dir, ignore=shutil.ignore_patterns("exports"))

            # Rewrite project.json with updated name/timestamps
            proj_json = dst_dir / self._project.path.name
            data = dict(self._project.data or {})
            data["project"] = new_name
            data["created_at"] = _now_iso()
            data["updated_at"] = _now_iso()
            _write_json(proj_json, data)

            # Load newly saved project
            self._last_project_path = str(proj_json.resolve())
            self.load_project(self._last_project_path)
            return {"ok": True, "path": self._last_project_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def append_pdf_to_project(self, pdf_path: str) -> dict[str, Any]:
        """
        Append another PDF to current project's template.pdf (merge pages).
        This produces a single combined PDF for the project.
        """
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            src = Path(pdf_path).resolve()
            if not src.exists():
                return {"ok": False, "error": "pdf_not_found"}

            dst_pdf = self._pdf_path()
            # Close renderer before touching the PDF file on Windows.
            try:
                if self._fitz_doc is not None:
                    self._fitz_doc.close()
            except Exception:
                pass
            self._fitz_doc = None
            self._fitz_pdf_path = None

            # Keep a copy of the added PDF inside project folder for traceability.
            try:
                src_dir = self._project.path.parent / "sources"
                src_dir.mkdir(parents=True, exist_ok=True)
                stamp = time.strftime("%Y%m%d-%H%M%S")
                dst_src = src_dir / f"{stamp}-{_safe_name(src.stem)}.pdf"
                shutil.copy2(src, dst_src)
            except Exception:
                pass

            # Merge existing template and new PDF into template.pdf
            reader_a = PdfReader(str(dst_pdf))
            reader_b = PdfReader(str(src))
            writer = PdfWriter()
            for pg in reader_a.pages:
                writer.add_page(pg)
            for pg in reader_b.pages:
                writer.add_page(pg)

            tmp = dst_pdf.with_suffix(".pdf.tmp")
            with open(tmp, "wb") as f:
                writer.write(f)

            # Optional backup
            try:
                bak = dst_pdf.with_name(f"template__bak_{int(time.time())}.pdf")
                shutil.copy2(dst_pdf, bak)
            except Exception:
                pass

            os.replace(tmp, dst_pdf)

            # Reload and reset caches
            try:
                if fitz is not None:
                    self._fitz_doc = fitz.open(str(dst_pdf))
                    self._fitz_pdf_path = str(dst_pdf)
                    self._page_count = max(1, int(self._fitz_doc.page_count))
                else:
                    self._page_count = max(1, int(len(PdfReader(str(dst_pdf)).pages)))
            except Exception:
                self._fitz_doc = None
                self._fitz_pdf_path = None
                self._page_count = 1

            self._page_cache.clear()
            self._invalidate_pages(None)

            self._project.data["updated_at"] = _now_iso()
            _write_json(self._project.path, self._project.data)
            return {"ok": True, "page_count": int(self._page_count)}
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
        # If there are no workers yet, seed a friendly default (prevents confusing empty UI).
        if not workers:
            workers = [{"id": "w1", "name": "作業者1", "bank": "", "hourly_yen": 0}]
            _write_json(WORKERS_PATH, workers)
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

    def delete_worker(self, worker_id: str) -> dict[str, Any]:
        try:
            wid = str(worker_id or "").strip()
            if not wid:
                return {"ok": False, "error": "missing_id"}
            rows = _read_json(WORKERS_PATH, [])
            if not isinstance(rows, list):
                rows = []
            out = [r for r in rows if not (isinstance(r, dict) and str(r.get("id") or "") == wid)]
            _write_json(WORKERS_PATH, out)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # --- work session ---
    def start_work(self, worker_id: str) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        self._working_worker_id = str(worker_id or "")
        self._private = False
        return {"ok": True}

    def toggle_private(self) -> dict[str, Any]:
        self._private = not self._private
        return {"ok": True, "in_private": self._private}

    # --- tags / placements / values ---
    def add_text_field(self, tag: str, page: int, x: float, y: float, font_size: int) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
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
        fid = f"f_{uuid.uuid4().hex[:8]}"
        placements[fid] = {
            "tag": t,
            "page": int(page or 0),
            "x": float(x),
            "y": float(y),
            "font_size": int(font_size or 14),
            "color": "#0f172a",
            "line_height": 1.2,
            "letter_spacing": 0,
        }
        data["placements"] = placements
        _write_json(self._project.path, data)
        self._invalidate_pages({int(page or 0)})
        return {"ok": True, "fid": fid, "tag": t}

    def set_element_pos(self, fid: str, x: float, y: float) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        f = str(fid or "").strip()
        if not f:
            return {"ok": False, "error": "missing_id"}
        placements = dict(self._project.data.get("placements") or {})
        if f not in placements or not isinstance(placements.get(f), dict):
            placements[f] = {"tag": "", "page": 0, "x": float(x), "y": float(y), "font_size": 14, "color": "#0f172a", "line_height": 1.2, "letter_spacing": 0}
        else:
            placements[f]["x"] = float(x)
            placements[f]["y"] = float(y)
        self._project.data["placements"] = placements
        _write_json(self._project.path, self._project.data)
        try:
            self._invalidate_pages({int(placements[f].get("page") or 0)})
        except Exception:
            self._invalidate_pages(None)
        return {"ok": True}

    def get_element_info(self, fid: str) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        f = str(fid or "").strip()
        pl = (self._project.data.get("placements") or {}).get(f)
        if not isinstance(pl, dict):
            return {"ok": False, "error": "not_found"}
        page = int(pl.get("page") or 0)
        w, h = self._page_image_size(page)
        return {
            "ok": True,
            "page": page,
            "tag": str(pl.get("tag") or ""),
            "x": float(pl.get("x") or 0),
            "y": float(pl.get("y") or 0),
            "font_size": int(pl.get("font_size") or 14),
            "page_display_width": w,
            "page_display_height": h,
        }

    def set_value(self, tag: str, value: str) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        t = str(tag or "").strip()
        values = dict(self._project.data.get("values") or {})
        values[t] = str(value or "")
        self._project.data["values"] = values
        _write_json(self._project.path, self._project.data)
        # Invalidate all pages that have placements using this tag.
        try:
            pages: set[int] = set()
            for _, pl in (self._project.data.get("placements") or {}).items():
                if isinstance(pl, dict) and str(pl.get("tag") or "").strip() == t:
                    pages.add(int(pl.get("page") or 0))
            self._invalidate_pages(pages if pages else None)
        except Exception:
            self._invalidate_pages(None)
        return {"ok": True}

    def update_placement(self, fid: str, patch: dict[str, Any]) -> dict[str, Any]:
        """Update style/position fields for a placement."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        f = str(fid or "").strip()
        if not f:
            return {"ok": False, "error": "missing_id"}
        placements = dict(self._project.data.get("placements") or {})
        pl = placements.get(f)
        if not isinstance(pl, dict):
            return {"ok": False, "error": "not_found"}
        if not isinstance(patch, dict):
            return {"ok": False, "error": "invalid_patch"}
        for k, v in patch.items():
            if k in ("x", "y"):
                pl[k] = float(v)
            elif k in ("page",):
                pl[k] = int(v)
            elif k in ("font_size",):
                pl[k] = int(v)
            elif k in ("color",):
                pl[k] = str(v)
            elif k in ("line_height",):
                pl[k] = float(v)
            elif k in ("letter_spacing",):
                pl[k] = float(v)
            elif k in ("tag",):
                pl[k] = str(v)
        placements[f] = pl
        self._project.data["placements"] = placements
        _write_json(self._project.path, self._project.data)
        self._invalidate_pages({int(pl.get("page") or 0)})
        return {"ok": True}

    def delete_elements(self, fids: list[str]) -> dict[str, Any]:
        """Delete specific elements (placements). Does not delete tag values unless unused."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        if not isinstance(fids, list):
            return {"ok": False, "error": "invalid_args"}
        data = self._project.data
        placements = dict(data.get("placements") or {})
        pages: set[int] = set()
        removed_tags: list[str] = []
        for fid in [str(x).strip() for x in fids if str(x).strip()]:
            pl = placements.pop(fid, None)
            if isinstance(pl, dict):
                pages.add(int(pl.get("page") or 0))
                removed_tags.append(str(pl.get("tag") or "").strip())
        data["placements"] = placements

        # Remove tags that are no longer used by any placement.
        still_used = {str(pl.get("tag") or "").strip() for pl in placements.values() if isinstance(pl, dict)}
        tags0 = [str(t).strip() for t in (data.get("tags") or []) if str(t).strip()]
        if removed_tags:
            data["tags"] = [t for t in tags0 if t in still_used]
            values = dict(data.get("values") or {})
            for t in list(values.keys()):
                if str(t).strip() and str(t).strip() not in still_used:
                    values.pop(t, None)
            data["values"] = values

        _write_json(self._project.path, data)
        self._invalidate_pages(pages if pages else None)
        return {"ok": True}

    def delete_tags(self, tags: list[str]) -> dict[str, Any]:
        """Delete tags and associated values/placements."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        if not isinstance(tags, list):
            return {"ok": False, "error": "invalid_args"}
        tset = {str(t).strip() for t in tags if str(t).strip()}
        if not tset:
            return {"ok": True}
        data = self._project.data
        old_tags = list(data.get("tags") or [])
        data["tags"] = [t for t in old_tags if t not in tset]
        values = dict(data.get("values") or {})
        placements = dict(data.get("placements") or {})
        pages: set[int] = set()
        for t in list(tset):
            values.pop(t, None)
        # Remove all placements that use these tags.
        for fid, pl in list(placements.items()):
            if isinstance(pl, dict) and str(pl.get("tag") or "").strip() in tset:
                pages.add(int(pl.get("page") or 0))
                placements.pop(fid, None)
        data["values"] = values
        data["placements"] = placements
        _write_json(self._project.path, data)
        self._invalidate_pages(pages if pages else None)
        return {"ok": True}

    def set_project_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Replace tags/values/placements in current project (for undo/redo & bulk ops)."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        if not isinstance(payload, dict):
            return {"ok": False, "error": "invalid_payload"}
        data = self._project.data
        tags = payload.get("tags")
        values = payload.get("values")
        placements = payload.get("placements")
        if isinstance(tags, list):
            data["tags"] = [str(t) for t in tags if str(t).strip()]
        if isinstance(values, dict):
            data["values"] = {str(k): str(v) for k, v in values.items()}
        if isinstance(placements, dict):
            data["placements"] = dict(placements)
        _write_json(self._project.path, data)
        self._invalidate_pages(None)
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
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        q = str(tag or "").strip()
        placements = dict(self._project.data.get("placements") or {})
        page_index = 0
        # Accept fid (new) or tag (legacy).
        if q in placements and isinstance(placements.get(q), dict):
            page_index = int(placements[q].get("page") or 0)
        else:
            for _, pl in placements.items():
                if isinstance(pl, dict) and str(pl.get("tag") or "").strip() == q:
                    page_index = int(pl.get("page") or 0)
                    break
        # Route to page renderer so cache/prefetch & PyMuPDF path applies.
        return self.get_preview_png_base64_page(page_index)

    def finish(self) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            pdf_in = self._pdf_path()
            reader = PdfReader(str(pdf_in))
            placements = dict(self._project.data.get("placements") or {})
            values = dict(self._project.data.get("values") or {})

            # Ensure Japanese-capable font for PDF export.
            # Use built-in CID font so we don't depend on external font files.
            try:
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont

                pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
                jp_font = "HeiseiKakuGo-W5"
            except Exception:
                jp_font = "Helvetica"

            import re

            def _needs_jp(s: str) -> bool:
                return bool(re.search(r"[\u3040-\u30ff\u3400-\u9fff\u3000-\u303f\uff00-\uffef]", s or ""))

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
                for _, p in placements.items():
                    if not isinstance(p, dict):
                        continue
                    if int(p.get("page") or 0) != pi:
                        continue
                    tag = str(p.get("tag") or "").strip()
                    if not tag:
                        continue
                    text = str(values.get(tag) or "").replace("<br>", "\n")
                    if not text.strip():
                        continue
                    x_px = float(p.get("x") or 0)
                    y_px = float(p.get("y") or 0)
                    fs = int(p.get("font_size") or 14)
                    color = str(p.get("color") or "#0f172a")
                    line_h = float(p.get("line_height") or 1.2)
                    letter_s = float(p.get("letter_spacing") or 0)
                    x_pt = x_px * 72.0 / RENDER_DPI
                    y_pt = (h_px - y_px) * 72.0 / RENDER_DPI
                    font_name = jp_font if _needs_jp(text) else "Helvetica"
                    c.setFont(font_name, fs)
                    try:
                        c.setFillColor(HexColor(color))
                    except Exception:
                        c.setFillColor(HexColor("#0f172a"))

                    def _draw_line_with_spacing(x0: float, y0: float, s: str) -> None:
                        if not letter_s:
                            c.drawString(x0, y0, s)
                            return
                        cx = x0
                        for ch in s:
                            c.drawString(cx, y0, ch)
                            try:
                                w = pdfmetrics.stringWidth(ch, font_name, fs)
                            except Exception:
                                w = fs * 0.62
                            cx += float(w) + float(letter_s) * 72.0 / RENDER_DPI

                    for line_idx, line in enumerate(text.splitlines() or [""]):
                        y_line = y_pt - (fs * line_h) * line_idx
                        _draw_line_with_spacing(x_pt, y_line, line)
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





