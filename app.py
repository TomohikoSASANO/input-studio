from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
import sys
import traceback
import time
import uuid
import zipfile
import threading
import ctypes
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
except Exception:  # pragma: no cover
    fitz = None


def _configure_pythonnet_coreclr() -> None:
    """
    Ensure Python.NET is loaded BEFORE importing pywebview.

    Policy:
    - Try netfx first (works on machines with .NET Framework and needs no dotnet root).
    - If netfx fails, try coreclr (.NET Desktop Runtime). This requires dotnet root.
    """
    base = None
    if hasattr(sys, "_MEIPASS"):
        # In PyInstaller one-folder, this points to the bundled runtime directory (usually `_internal`).
        base = Path(getattr(sys, "_MEIPASS"))  # type: ignore[arg-type]
    else:
        base = Path(__file__).resolve().parent

    rc = base / "coreclr.runtimeconfig.json"
    if not rc.exists():
        # Fallback: next to executable (or in _internal next to it)
        exe_dir = Path(sys.executable).resolve().parent
        cand = exe_dir / "coreclr.runtimeconfig.json"
        if cand.exists():
            rc = cand
        else:
            cand2 = exe_dir / "_internal" / "coreclr.runtimeconfig.json"
            if cand2.exists():
                rc = cand2

    # Import pythonnet lazily; if it's not bundled, raise a clear error.
    try:
        import pythonnet  # type: ignore
    except Exception as e:
        raise RuntimeError(f"pythonnet is missing or failed to import: {e}") from e

    # 1) Try .NET Framework (netfx) first.
    try:
        pythonnet.load("netfx")
        return
    except Exception:
        netfx_err = traceback.format_exc()

    # 2) Fallback to CoreCLR (.NET Desktop Runtime)
    def _find_dotnet_root() -> Path | None:
        # Prefer explicit env if present
        env = os.environ.get("DOTNET_ROOT") or os.environ.get("DOTNET_ROOT(x86)")
        if env:
            p = Path(env)
            if p.is_dir():
                return p
        # Common install paths
        for c in [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramW6432"),
        ]:
            if c:
                p = Path(c) / "dotnet"
                if p.is_dir():
                    return p
        la = os.environ.get("LocalAppData")
        if la:
            p = Path(la) / "Microsoft" / "dotnet"
            if p.is_dir():
                return p
        # dotnet on PATH
        try:
            import shutil as _shutil

            d = _shutil.which("dotnet")
            if d:
                return Path(d).resolve().parent
        except Exception:
            pass
        return None

    dotnet_root = _find_dotnet_root()
    if not dotnet_root:
        raise RuntimeError(
            "Failed to initialize .NET runtime.\n\n"
            "Tried netfx first, but it failed. Then tried coreclr, but DOTNET_ROOT could not be determined.\n\n"
            "Please install '.NET Desktop Runtime (x64)' and reboot.\n\n"
            f"netfx error:\n{netfx_err}"
        )

    # For pythonnet+coreclr we pass params directly (more reliable than env-only).
    if rc.exists():
        pythonnet.load("coreclr", runtime_config=str(rc), dotnet_root=str(dotnet_root))
    else:
        # As a last resort, try coreclr with discovered dotnet root (will pick latest runtime).
        pythonnet.load("coreclr", dotnet_root=str(dotnet_root))


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
    """
    Simple tategaki approximation:
    - splitlines() => columns (right to left)
    - each character is drawn downwards
    - ASCII chars are rotated 90deg for readability
    """
    c.setFont(font_name, fs_pt)

    # vertical step (down)
    leading = float(fs_pt) * float(line_h or 1.2)
    step_y = leading + float(letter_s_pt or 0.0)
    # column step (left)
    step_x = float(fs_pt) * 1.10 + float(letter_s_pt or 0.0)

    cols = str(text or "").splitlines() or [""]
    # baseline: convert top anchor to a usable baseline for drawString.
    # Slight extra downward shift keeps PDF output aligned with preview PNG text box.
    baseline_shift_em = PDF_BASELINE_SHIFT_EM
    try:
        ascent = float(pdfmetrics.getAscent(font_name) or 0) / 1000.0 * fs_pt
    except Exception:
        ascent = fs_pt * 0.8
    y0 = float(y_top_pt) - float(ascent) - (baseline_shift_em * float(fs_pt))

    for col_idx, col in enumerate(cols):
        cx = float(x_pt) - (col_idx * step_x)  # vertical-rl
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


ROOT = Path(__file__).resolve().parent
UI_DIR = ROOT / "ui"


def _default_local_data_dir() -> Path:
    """
    Store user data under LocalAppData so it survives app updates.
    Windows: %LOCALAPPDATA%/InputStudio/_local_data
    """
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if base:
        return Path(base) / "InputStudio" / "_local_data"
    return Path.home() / "AppData" / "Local" / "InputStudio" / "_local_data"


# Allow override for development/testing
LOCAL = Path(os.environ.get("INPUTSTUDIO_LOCAL_DIR") or _default_local_data_dir())
PROJECTS_DIR = LOCAL / "projects"
WORKERS_PATH = LOCAL / "workers.json"
ADMIN_SETTINGS_PATH = LOCAL / "admin_settings.json"

# Bundled font (for preview PNG + PDF export). Included via PyInstaller spec.
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", str(ROOT)))  # type: ignore[arg-type]
BUNDLED_FONT_PATH = (RESOURCE_ROOT / "assets" / "fonts" / "NotoSansJP-Regular.ttf").resolve()

# Render DPI for preview images and coordinate system in this app.
RENDER_DPI = 150
PDF_BASELINE_SHIFT_EM = 0.26
PDF_LETTER_SPACING_FACTOR = 0.72
DEFAULT_LETTER_SPACING = 1.2


def _ensure_dirs() -> None:
    _migrate_legacy_local_data_if_needed()
    LOCAL.mkdir(parents=True, exist_ok=True)
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)


def _legacy_local_dirs() -> list[Path]:
    exe_dir = Path(sys.executable).resolve().parent
    return [
        ROOT / "_local_data",
        exe_dir / "_local_data",
        exe_dir / "_internal" / "_local_data",
        ROOT / "_internal" / "_local_data",
    ]


def _dir_looks_nonempty(p: Path) -> bool:
    try:
        if not p.exists() or not p.is_dir():
            return False
        if (p / "admin_settings.json").exists() or (p / "workers.json").exists():
            return True
        d = p / "projects"
        return d.exists() and any(d.iterdir())
    except Exception:
        return False


def _migrate_legacy_local_data_if_needed() -> None:
    """
    One-time best-effort migration:
    If new LOCAL is empty but a legacy dir has data, copy it over.
    """
    # If user explicitly overrides local dir, do not auto-migrate.
    if os.environ.get("INPUTSTUDIO_LOCAL_DIR"):
        return
    try:
        dst = LOCAL
        dst.mkdir(parents=True, exist_ok=True)
        if _dir_looks_nonempty(dst):
            return
        for src in _legacy_local_dirs():
            if src.resolve() == dst.resolve():
                continue
            if not _dir_looks_nonempty(src):
                continue
            for item in src.iterdir():
                target = dst / item.name
                if target.exists():
                    continue
                if item.is_dir():
                    shutil.copytree(item, target, dirs_exist_ok=True)
                else:
                    shutil.copy2(item, target)
            return
    except Exception:
        return


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

            # writing_mode default
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
                        # Prefer bundled font (ensures preview matches PDF export).
                        candidates = []
                        try:
                            if BUNDLED_FONT_PATH and BUNDLED_FONT_PATH.exists():
                                candidates.append(str(BUNDLED_FONT_PATH))
                        except Exception:
                            pass
                        # Fallback: Windows built-ins.
                        candidates.extend(
                            [
                                r"C:\Windows\Fonts\meiryo.ttc",
                                r"C:\Windows\Fonts\YuGothR.ttc",
                                r"C:\Windows\Fonts\YuGothM.ttc",
                                r"C:\Windows\Fonts\msgothic.ttc",
                                r"C:\Windows\Fonts\msmincho.ttc",
                                "arial.ttf",
                            ]
                        )
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

            def _is_ascii(ch: str) -> bool:
                try:
                    return ord(ch) < 128
                except Exception:
                    return False

            def _draw_text_vertical(draw2: Any, x: float, y: float, text: str, fs: int, fill: tuple[int, int, int, int], line_h: float, letter_s: float) -> None:
                """
                Simple tategaki for preview PNG:
                - splitlines() => columns (right to left)
                - characters stacked downward
                - ASCII chars are rotated 90deg
                """
                fnt = font(fs)
                step_y = float(fs) * float(line_h) + float(letter_s or 0.0)
                step_x = float(fs) * 1.10 + float(letter_s or 0.0)

                cols = str(text or "").split("\n")
                if not cols:
                    cols = [""]

                # Use RGBA compositing for rotated ASCII.
                try:
                    from PIL import Image
                except Exception:
                    Image = None  # type: ignore

                for col_idx, col in enumerate(cols):
                    cx = float(x) - (col_idx * step_x)
                    cy = float(y)
                    for ch in str(col or ""):
                        if ch == "\r":
                            continue
                        if ch == " ":
                            cy += step_y
                            continue
                        if _is_ascii(ch) and Image is not None and fnt is not None:
                            # draw rotated glyph onto temporary image then paste
                            pad = int(max(2, fs * 0.2))
                            tmp = Image.new("RGBA", (fs + pad * 2, fs + pad * 2), (0, 0, 0, 0))
                            td = ImageDraw.Draw(tmp)
                            td.text((pad, pad), ch, fill=fill, font=fnt)
                            rot = tmp.rotate(90, expand=True)
                            try:
                                img0 = draw2.im  # PIL internal
                                img0.alpha_composite(rot, (int(cx), int(cy)))
                            except Exception:
                                try:
                                    draw2.bitmap((cx, cy), rot)
                                except Exception:
                                    draw2.text((cx, cy), ch, fill=fill, font=fnt)
                        else:
                            draw2.text((cx, cy), ch, fill=fill, font=fnt)
                        cy += step_y

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
                letter_s = float(p.get("letter_spacing") or DEFAULT_LETTER_SPACING)
                writing_mode = _norm_writing_mode(p.get("writing_mode"))
                if writing_mode == "vertical":
                    _draw_text_vertical(draw, x, y, text, fs, color, line_h, letter_s)
                else:
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

    def save_current_project(self, make_filled_pdf: bool = False) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            self._project.data["ui_mode"] = self._ui_mode
            self._project.data["updated_at"] = _now_iso()
            _write_json(self._project.path, self._project.data)
            filled_pdf = None
            pdf_path = None
            if bool(make_filled_pdf):
                out_dir = (self._project.path.parent / "exports").resolve()
                stamp = time.strftime("%Y%m%d-%H%M%S")
                proj = _safe_name(str(self._project.data.get("project") or "project"))
                who = _safe_name(str(self._working_worker_id or "worker"))
                out_pdf = out_dir / f"autosave-{proj}-{stamp}-{who}.pdf"
                self._export_filled_pdf(out_pdf)
                latest = (self._project.path.parent / "template_filled_latest.pdf").resolve()
                try:
                    shutil.copy2(out_pdf, latest)
                except Exception:
                    try:
                        self._export_filled_pdf(latest)
                    except Exception:
                        pass
                filled_pdf = str(latest if latest else out_pdf.resolve())
                pdf_path = str(out_pdf.resolve())
            return {
                "ok": True,
                "path": str(self._project.path),
                "project_dir": str(self._project.path.parent),
                "exports_dir": str((self._project.path.parent / "exports").resolve()),
                "pdf": pdf_path,
                "filled_pdf": filled_pdf,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_project_as(self, name: str, make_filled_pdf: bool = True) -> dict[str, Any]:
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
            filled_pdf = None
            pdf_path = None
            if bool(make_filled_pdf) and self._project:
                out_dir = (Path(self._last_project_path).resolve().parent / "exports").resolve()
                stamp = time.strftime("%Y%m%d-%H%M%S")
                proj = _safe_name(str(self._project.data.get("project") or "project"))
                who = _safe_name(str(self._working_worker_id or "worker"))
                out_pdf = out_dir / f"autosave-{proj}-{stamp}-{who}.pdf"
                self._export_filled_pdf(out_pdf)
                latest = (Path(self._last_project_path).resolve().parent / "template_filled_latest.pdf").resolve()
                try:
                    shutil.copy2(out_pdf, latest)
                except Exception:
                    try:
                        self._export_filled_pdf(latest)
                    except Exception:
                        pass
                filled_pdf = str(latest if latest else out_pdf.resolve())
                pdf_path = str(out_pdf.resolve())
            return {
                "ok": True,
                "path": self._last_project_path,
                "project_dir": str(Path(self._last_project_path).resolve().parent),
                "exports_dir": str((Path(self._last_project_path).resolve().parent / "exports").resolve()),
                "pdf": pdf_path,
                "filled_pdf": filled_pdf,
            }
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

    def copy_page_with_elements(self, page_index: int) -> dict[str, Any]:
        """Duplicate one page in template PDF and duplicate placements on that page."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            dst_pdf = self._pdf_path()
            reader = PdfReader(str(dst_pdf))
            total = len(reader.pages)
            if total <= 0:
                return {"ok": False, "error": "no_pages"}
            idx = max(0, min(total - 1, int(page_index or 0)))

            try:
                if self._fitz_doc is not None:
                    self._fitz_doc.close()
            except Exception:
                pass
            self._fitz_doc = None
            self._fitz_pdf_path = None

            writer = PdfWriter()
            for i, pg in enumerate(reader.pages):
                writer.add_page(pg)
                if i == idx:
                    writer.add_page(pg)

            tmp = dst_pdf.with_suffix(".pdf.tmp")
            with open(tmp, "wb") as f:
                writer.write(f)
            os.replace(tmp, dst_pdf)

            data = self._project.data
            placements = dict(data.get("placements") or {})
            out: dict[str, Any] = {}
            for fid, pl in placements.items():
                if not isinstance(pl, dict):
                    continue
                p = int(pl.get("page") or 0)
                base = dict(pl)
                if p > idx:
                    base["page"] = p + 1
                out[str(fid)] = base
                if p == idx:
                    nf = f"f_{uuid.uuid4().hex[:8]}"
                    cp = dict(pl)
                    cp["page"] = idx + 1
                    out[nf] = cp
            data["placements"] = out
            data["updated_at"] = _now_iso()
            _write_json(self._project.path, data)

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
                self._page_count = max(1, total + 1)

            self._invalidate_pages(None)
            return {
                "ok": True,
                "page_count": int(self._page_count),
                "page_index": int(idx + 1),
                "placements": dict(data.get("placements") or {}),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_page_from_project(self, page_index: int) -> dict[str, Any]:
        """Delete one page from template PDF and remove/shift related placements."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            dst_pdf = self._pdf_path()
            reader = PdfReader(str(dst_pdf))
            total = len(reader.pages)
            if total <= 1:
                return {"ok": False, "error": "cannot_delete_last_page"}
            idx = max(0, min(total - 1, int(page_index or 0)))

            try:
                if self._fitz_doc is not None:
                    self._fitz_doc.close()
            except Exception:
                pass
            self._fitz_doc = None
            self._fitz_pdf_path = None

            writer = PdfWriter()
            for i, pg in enumerate(reader.pages):
                if i == idx:
                    continue
                writer.add_page(pg)

            tmp = dst_pdf.with_suffix(".pdf.tmp")
            with open(tmp, "wb") as f:
                writer.write(f)
            os.replace(tmp, dst_pdf)

            data = self._project.data
            placements = dict(data.get("placements") or {})
            out: dict[str, Any] = {}
            for fid, pl in placements.items():
                if not isinstance(pl, dict):
                    continue
                p = int(pl.get("page") or 0)
                if p == idx:
                    continue
                item = dict(pl)
                if p > idx:
                    item["page"] = p - 1
                out[str(fid)] = item
            data["placements"] = out

            # Remove tags/values that became unused.
            still_used = {str(pl.get("tag") or "").strip() for pl in out.values() if isinstance(pl, dict)}
            tags0 = [str(t).strip() for t in (data.get("tags") or []) if str(t).strip()]
            data["tags"] = [t for t in tags0 if t in still_used]
            values = dict(data.get("values") or {})
            for t in list(values.keys()):
                if str(t).strip() not in still_used:
                    values.pop(t, None)
            data["values"] = values
            data["updated_at"] = _now_iso()
            _write_json(self._project.path, data)

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
                self._page_count = max(1, total - 1)

            self._invalidate_pages(None)
            next_idx = max(0, min(int(self._page_count) - 1, idx))
            return {
                "ok": True,
                "page_count": int(self._page_count),
                "page_index": int(next_idx),
                "tags": list(data.get("tags") or []),
                "values": dict(data.get("values") or {}),
                "placements": dict(data.get("placements") or {}),
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def reorder_pages(self, order: list[int]) -> dict[str, Any]:
        """Reorder template PDF pages and remap placement page indices."""
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            if not isinstance(order, list):
                return {"ok": False, "error": "invalid_order"}
            dst_pdf = self._pdf_path()
            reader = PdfReader(str(dst_pdf))
            total = len(reader.pages)
            if total <= 0:
                return {"ok": False, "error": "no_pages"}

            normalized: list[int] = []
            for x in order:
                try:
                    normalized.append(int(x))
                except Exception:
                    return {"ok": False, "error": "invalid_order"}
            if len(normalized) != total or len(set(normalized)) != total:
                return {"ok": False, "error": "invalid_order"}
            if min(normalized) < 0 or max(normalized) >= total:
                return {"ok": False, "error": "invalid_order"}

            try:
                if self._fitz_doc is not None:
                    self._fitz_doc.close()
            except Exception:
                pass
            self._fitz_doc = None
            self._fitz_pdf_path = None

            writer = PdfWriter()
            for old_idx in normalized:
                writer.add_page(reader.pages[int(old_idx)])
            tmp = dst_pdf.with_suffix(".pdf.tmp")
            with open(tmp, "wb") as f:
                writer.write(f)
            os.replace(tmp, dst_pdf)

            # remap placements page index: old -> new
            old_to_new = {int(old): int(new) for new, old in enumerate(normalized)}
            data = self._project.data
            placements = dict(data.get("placements") or {})
            for fid, pl in placements.items():
                if not isinstance(pl, dict):
                    continue
                old_page = int(pl.get("page") or 0)
                if old_page in old_to_new:
                    pl["page"] = old_to_new[old_page]
                    placements[fid] = pl
            data["placements"] = placements
            data["updated_at"] = _now_iso()
            _write_json(self._project.path, data)

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
                self._page_count = max(1, total)

            self._invalidate_pages(None)
            return {
                "ok": True,
                "page_count": int(self._page_count),
                "placements": dict(data.get("placements") or {}),
            }
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
        # Backward-compatible defaults (old settings files may not have these keys).
        if "default_font_size" not in s:
            s["default_font_size"] = 14
        if "view_zoom" not in s:
            s["view_zoom"] = 1.0
        return {"ok": True, "settings": s}

    def update_admin_settings(self, patch: dict[str, Any]) -> dict[str, Any]:
        """
        Merge/overwrite admin settings keys.
        Backward compatible: only adds optional keys; does not change existing meanings.
        """
        try:
            cur = _read_json(ADMIN_SETTINGS_PATH, {"ui_mode": "worker"})
            if not isinstance(cur, dict):
                cur = {"ui_mode": "worker"}
            if not isinstance(patch, dict):
                return {"ok": False, "error": "invalid_patch"}
            for k, v in patch.items():
                cur[str(k)] = v
            _write_json(ADMIN_SETTINGS_PATH, cur)
            return {"ok": True, "settings": cur}
        except Exception as e:
            return {"ok": False, "error": str(e)}

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
            "letter_spacing": DEFAULT_LETTER_SPACING,
            "writing_mode": "horizontal",
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
            placements[f] = {
                "tag": "",
                "page": 0,
                "x": float(x),
                "y": float(y),
                "font_size": 14,
                "color": "#0f172a",
                "line_height": 1.2,
                "letter_spacing": DEFAULT_LETTER_SPACING,
                "writing_mode": "horizontal",
            }
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
            elif k in ("writing_mode",):
                pl[k] = _norm_writing_mode(v)
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

    def _export_filled_pdf(self, out_pdf: Path) -> None:
        """Render current project values onto template.pdf and write to out_pdf."""
        if not self._project and not self._ensure_project_loaded():
            raise RuntimeError("no_project")
        assert self._project is not None

        pdf_in = self._pdf_path()
        reader = PdfReader(str(pdf_in))
        placements = dict(self._project.data.get("placements") or {})
        values = dict(self._project.data.get("values") or {})

        # Prefer bundled TTF so preview(Pillow) and export(PDF) match.
        bundled_font_name = "InputStudioFont"
        pdf_font = "Helvetica"
        try:
            if BUNDLED_FONT_PATH and BUNDLED_FONT_PATH.exists():
                pdfmetrics.registerFont(TTFont(bundled_font_name, str(BUNDLED_FONT_PATH)))
                pdf_font = bundled_font_name
        except Exception:
            pdf_font = "Helvetica"

        # Fallback Japanese CID font (if bundled font is unavailable)
        if pdf_font == "Helvetica":
            try:
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont

                pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
                pdf_font = "HeiseiKakuGo-W5"
            except Exception:
                pdf_font = "Helvetica"

        writer = PdfWriter()
        for pi, page in enumerate(reader.pages):
            # Use CropBox like preview, and match preview pixel rounding.
            media = page.mediabox
            crop = getattr(page, "cropbox", None) or media
            w_pt = float(media.width)
            h_pt = float(media.height)
            llx = float(getattr(crop, "lower_left", (0, 0))[0])
            lly = float(getattr(crop, "lower_left", (0, 0))[1])
            crop_w_pt = float(crop.width)
            crop_h_pt = float(crop.height)
            # Use the same pixel rounding as the preview renderer to avoid drift.
            try:
                crop_w_px, crop_h_px = self._page_image_size(pi)
            except Exception:
                crop_w_px = max(1, int(round(crop_w_pt / 72.0 * RENDER_DPI)))
                crop_h_px = max(1, int(round(crop_h_pt / 72.0 * RENDER_DPI)))

            import io

            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=(w_pt, h_pt))
            # Ensure overlay has at least one page.
            try:
                c.setFont("Helvetica", 1)
                c.setFillColor(HexColor("#ffffff"))
                c.drawString(-10000, -10000, " ")
            except Exception:
                pass

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
                fs_px = float(p.get("font_size") or 14)
                color = str(p.get("color") or "#0f172a")
                line_h = float(p.get("line_height") or 1.2)
                letter_s_px = float(p.get("letter_spacing") or DEFAULT_LETTER_SPACING)

                x_pt = llx + (x_px / float(crop_w_px)) * crop_w_pt
                y_top_pt = lly + crop_h_pt - ((y_px / float(crop_h_px)) * crop_h_pt)

                # Keep the same font family for all glyphs when possible
                # so preview(Pillow) and export(PDF) line wrapping stay close.
                font_name = pdf_font
                fs_pt = float(fs_px) * 72.0 / RENDER_DPI
                c.setFont(font_name, fs_pt)
                try:
                    c.setFillColor(HexColor(color))
                except Exception:
                    c.setFillColor(HexColor("#0f172a"))

                # baseline adjust (top-anchor in UI -> baseline in PDF)
                baseline_shift_em = PDF_BASELINE_SHIFT_EM
                try:
                    ascent = float(pdfmetrics.getAscent(font_name) or 0) / 1000.0 * fs_pt
                except Exception:
                    ascent = fs_pt * 0.8
                y_base0 = y_top_pt - ascent - (baseline_shift_em * fs_pt)
                letter_s_pt = float(letter_s_px) * 72.0 / RENDER_DPI * PDF_LETTER_SPACING_FACTOR

                def _draw_line_with_spacing(x0: float, y0: float, s: str) -> None:
                    if not letter_s_pt:
                        c.drawString(x0, y0, s)
                        return
                    cx = x0
                    for ch in s:
                        c.drawString(cx, y0, ch)
                        try:
                            w = pdfmetrics.stringWidth(ch, font_name, fs_pt)
                        except Exception:
                            w = fs_pt * 0.62
                        cx += float(w) + float(letter_s_pt)

                writing_mode = _norm_writing_mode(p.get("writing_mode"))
                if writing_mode == "vertical":
                    _draw_vertical_text(
                        c,
                        x_pt=x_pt,
                        y_top_pt=y_top_pt,
                        text=text,
                        font_name=font_name,
                        fs_pt=fs_pt,
                        line_h=line_h,
                        letter_s_pt=letter_s_pt,
                    )
                else:
                    for line_idx, line in enumerate(text.splitlines() or [""]):
                        y_line = y_base0 - (fs_pt * line_h) * line_idx
                        _draw_line_with_spacing(x_pt, y_line, line)

            c.save()
            packet.seek(0)
            overlay = PdfReader(packet).pages[0]
            page.merge_page(overlay)
            writer.add_page(page)

        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        with out_pdf.open("wb") as f:
            writer.write(f)

    def _render_report_pdf(self, out_pdf: Path, meta: dict[str, Any]) -> None:
        """
        Create a simple report PDF for sharing/administration.
        Contains:
        - worker info (name/bank/hourly)
        - project info
        - work time summary
        - counts (total tags / filled tags / placements)
        - (optional) tag/value list
        """
        if not self._project:
            raise RuntimeError("no_project")

        # Font (match preview as much as possible)
        font_name = "Helvetica"
        try:
            if BUNDLED_FONT_PATH and BUNDLED_FONT_PATH.exists():
                pdfmetrics.registerFont(TTFont("InputStudioFont", str(BUNDLED_FONT_PATH)))
                font_name = "InputStudioFont"
        except Exception:
            font_name = "Helvetica"
        if font_name == "Helvetica":
            try:
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont

                pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
                font_name = "HeiseiKakuGo-W5"
            except Exception:
                font_name = "Helvetica"

        def _fmt_sec(s: Any) -> str:
            try:
                sec = int(float(s))
            except Exception:
                sec = 0
            if sec < 0:
                sec = 0
            hh = sec // 3600
            mm = (sec % 3600) // 60
            ss = sec % 60
            return f"{hh:02d}:{mm:02d}:{ss:02d}"

        def _wrap(s: str, max_chars: int) -> list[str]:
            s = str(s or "")
            out: list[str] = []
            cur = ""
            for ch in s:
                if ch == "\r":
                    continue
                if ch == "\n":
                    out.append(cur)
                    cur = ""
                    continue
                cur += ch
                if len(cur) >= max_chars:
                    out.append(cur)
                    cur = ""
            if cur:
                out.append(cur)
            return out or [""]

        project_name = str(self._project.data.get("project") or "")
        project_path = str(self._project.path)
        values = dict(self._project.data.get("values") or {})
        tags = [str(t) for t in (self._project.data.get("tags") or []) if str(t).strip()]
        placements = dict(self._project.data.get("placements") or {})

        # counts
        def _val_for(t: str) -> str:
            return str(values.get(t) or "").replace("<br>", "\n")

        filled_count = int(meta.get("filled_count")) if str(meta.get("filled_count", "")).strip() else 0
        if filled_count <= 0:
            filled_count = sum(1 for t in tags if _val_for(t).strip())
        total_tags = int(meta.get("total_tags")) if str(meta.get("total_tags", "")).strip() else len(tags)
        empty_count = max(0, int(total_tags) - int(filled_count))
        placement_count = int(meta.get("placement_count")) if str(meta.get("placement_count", "")).strip() else len(placements)

        worker_id = str(meta.get("worker_id") or self._working_worker_id or "").strip()
        worker_name = str(meta.get("worker_name") or "").strip()
        bank = ""
        hourly_yen = 0
        try:
            rows = _read_json(WORKERS_PATH, [])
            if isinstance(rows, list):
                for r in rows:
                    if isinstance(r, dict) and str(r.get("id") or "").strip() == worker_id:
                        if not worker_name:
                            worker_name = str(r.get("name") or "").strip()
                        bank = str(r.get("bank") or "").strip()
                        try:
                            hourly_yen = int(r.get("hourly_yen") or 0)
                        except Exception:
                            hourly_yen = 0
                        break
        except Exception:
            pass

        start_iso = str(meta.get("start_iso") or "").strip()
        end_iso = str(meta.get("end_iso") or "").strip()
        duration_sec = meta.get("duration_sec", 0)
        private_sec = meta.get("private_sec", 0)

        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        c = canvas.Canvas(str(out_pdf), pagesize=A4)
        w, h = A4
        margin_x = 36
        y = h - 42

        def hline() -> None:
            nonlocal y
            y -= 6
            c.setStrokeColor(HexColor("#e5e7eb"))
            c.line(margin_x, y, w - margin_x, y)
            y -= 10

        c.setFillColor(HexColor("#0f172a"))
        c.setFont(font_name, 16)
        c.drawString(margin_x, y, "作業報告書")
        y -= 22
        c.setFont(font_name, 10)
        c.setFillColor(HexColor("#475569"))
        c.drawString(margin_x, y, f"生成日時: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        y -= 14
        hline()

        c.setFillColor(HexColor("#0f172a"))
        c.setFont(font_name, 11)
        c.drawString(margin_x, y, "■ 作業者")
        y -= 16
        c.setFont(font_name, 10)
        for lab, val in [
            ("作業者ID", worker_id or "-"),
            ("氏名", worker_name or "-"),
            ("銀行", bank or "-"),
            ("時給(参考)", f"{hourly_yen:,} 円/h" if hourly_yen else "-"),
        ]:
            c.setFillColor(HexColor("#334155"))
            c.drawString(margin_x, y, f"{lab}:")
            c.setFillColor(HexColor("#0f172a"))
            c.drawString(margin_x + 84, y, str(val))
            y -= 14
        y -= 2
        hline()

        c.setFillColor(HexColor("#0f172a"))
        c.setFont(font_name, 11)
        c.drawString(margin_x, y, "■ 案件")
        y -= 16
        c.setFont(font_name, 10)
        for lab, val in [
            ("案件名", project_name or "-"),
            ("プロジェクト", project_path or "-"),
        ]:
            c.setFillColor(HexColor("#334155"))
            c.drawString(margin_x, y, f"{lab}:")
            c.setFillColor(HexColor("#0f172a"))
            # wrap long path
            lines = _wrap(str(val), 60)
            c.drawString(margin_x + 84, y, lines[0])
            y -= 14
            for extra in lines[1:]:
                c.drawString(margin_x + 84, y, extra)
                y -= 14
        y -= 2
        hline()

        c.setFillColor(HexColor("#0f172a"))
        c.setFont(font_name, 11)
        c.drawString(margin_x, y, "■ 作業時間")
        y -= 16
        c.setFont(font_name, 10)
        for lab, val in [
            ("開始", start_iso or "-"),
            ("終了", end_iso or "-"),
            ("正味時間", _fmt_sec(duration_sec)),
            ("中断(私用)合計", _fmt_sec(private_sec)),
        ]:
            c.setFillColor(HexColor("#334155"))
            c.drawString(margin_x, y, f"{lab}:")
            c.setFillColor(HexColor("#0f172a"))
            c.drawString(margin_x + 84, y, str(val))
            y -= 14
        y -= 2
        hline()

        c.setFillColor(HexColor("#0f172a"))
        c.setFont(font_name, 11)
        c.drawString(margin_x, y, "■ 入力集計")
        y -= 16
        c.setFont(font_name, 10)
        for lab, val in [
            ("項目数", f"{int(total_tags)}"),
            ("入力済み", f"{int(filled_count)}"),
            ("未入力", f"{int(empty_count)}"),
            ("要素数(配置)", f"{int(placement_count)}"),
        ]:
            c.setFillColor(HexColor("#334155"))
            c.drawString(margin_x, y, f"{lab}:")
            c.setFillColor(HexColor("#0f172a"))
            c.drawString(margin_x + 84, y, str(val))
            y -= 14

        # Next pages: value list (optional)
        y -= 8
        c.showPage()
        c.setFont(font_name, 12)
        c.setFillColor(HexColor("#0f172a"))
        c.drawString(margin_x, h - 42, "入力内容（タグ別）")
        y = h - 66
        c.setFont(font_name, 9)
        c.setFillColor(HexColor("#334155"))
        c.drawString(margin_x, y, "タグ")
        c.drawString(margin_x + 180, y, "値（改行は / で表示、長い場合は省略）")
        y -= 10
        c.setStrokeColor(HexColor("#e5e7eb"))
        c.line(margin_x, y, w - margin_x, y)
        y -= 12

        def _one_line(v: str) -> str:
            s = str(v or "").replace("<br>", "\n")
            s = " / ".join([x.strip() for x in s.splitlines() if x.strip()])
            if len(s) > 80:
                return s[:77] + "..."
            return s

        for t in tags:
            if y < 64:
                c.showPage()
                c.setFont(font_name, 12)
                c.setFillColor(HexColor("#0f172a"))
                c.drawString(margin_x, h - 42, "入力内容（タグ別）")
                y = h - 66
                c.setFont(font_name, 9)
                c.setFillColor(HexColor("#334155"))
                c.drawString(margin_x, y, "タグ")
                c.drawString(margin_x + 180, y, "値（改行は / で表示、長い場合は省略）")
                y -= 10
                c.setStrokeColor(HexColor("#e5e7eb"))
                c.line(margin_x, y, w - margin_x, y)
                y -= 12
            c.setFont(font_name, 9)
            c.setFillColor(HexColor("#0f172a"))
            c.drawString(margin_x, y, t[:28])
            c.setFillColor(HexColor("#334155"))
            c.drawString(margin_x + 180, y, _one_line(values.get(t) or ""))
            y -= 12

        c.save()

    def finish(self, report_meta: dict[str, Any] | None = None) -> dict[str, Any]:
        if not self._project and not self._ensure_project_loaded():
            return {"ok": False, "error": "no_project"}
        try:
            out_dir = self._project.path.parent / "exports"
            out_dir.mkdir(parents=True, exist_ok=True)
            stamp = time.strftime("%Y%m%d-%H%M%S")
            proj = _safe_name(str(self._project.data.get("project") or "project"))
            who = _safe_name(str(self._working_worker_id or "worker"))
            base = f"{proj}-{stamp}-{who}"
            out_pdf = out_dir / f"{base}.pdf"
            self._export_filled_pdf(out_pdf)

            latest = (self._project.path.parent / "template_filled_latest.pdf").resolve()
            try:
                shutil.copy2(out_pdf, latest)
            except Exception:
                try:
                    self._export_filled_pdf(latest)
                except Exception:
                    pass

            # ---- bundle for sharing (PDF + project data in same folder) ----
            bundle_dir = out_dir / base
            bundle_dir.mkdir(parents=True, exist_ok=True)
            # Copy project.json and template.pdf so another PC can open/edit this project.
            try:
                shutil.copy2(self._project.path, bundle_dir / "project.json")
            except Exception:
                pass
            try:
                shutil.copy2(self._pdf_path(), bundle_dir / "template.pdf")
            except Exception:
                pass
            try:
                shutil.copy2(out_pdf, bundle_dir / out_pdf.name)
            except Exception:
                pass
            # Report PDF
            report_path = bundle_dir / "report.pdf"
            try:
                meta = report_meta if isinstance(report_meta, dict) else {}
                # augment meta with worker name if missing
                if self._working_worker_id and "worker_id" not in meta:
                    meta["worker_id"] = self._working_worker_id
                if self._project and "project" not in meta:
                    meta["project"] = str(self._project.data.get("project") or "")
                self._render_report_pdf(report_path, meta)
            except Exception:
                # don't fail finish on report errors
                report_path = None  # type: ignore

            # ZIP for sending (includes the bundle folder)
            out_zip = out_dir / f"{base}.zip"
            with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for fp in bundle_dir.rglob("*"):
                    if fp.is_file():
                        arc = str(fp.relative_to(out_dir))
                        z.write(fp, arcname=arc)
            return {
                "ok": True,
                "dir": str(out_dir.resolve()),
                "zip": str(out_zip.resolve()),
                "pdf": str(out_pdf.resolve()),
                "filled_pdf": str(latest if latest else out_pdf.resolve()),
                "bundle_dir": str(bundle_dir.resolve()),
                "report_pdf": str(report_path.resolve()) if report_path else None,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def reveal_in_explorer(self, path: str) -> dict[str, Any]:
        """
        Open Windows Explorer at the given path.
        - If path is a file: open folder and select it.
        - If path is a folder: open it.
        """
        try:
            p = Path(str(path or "").strip())
            if not p:
                return {"ok": False, "error": "missing_path"}
            p = p.expanduser()
            if not p.is_absolute():
                # treat as relative to current project folder if possible
                if self._project:
                    p = (self._project.path.parent / p).resolve()
                else:
                    p = p.resolve()

            # Prefer explorer /select for files
            if p.exists() and p.is_file():
                os.system(f'explorer /select,"{str(p)}"')
                return {"ok": True}
            # Folder open
            target = p if p.exists() else p.parent
            if target.exists():
                os.startfile(str(target))  # type: ignore[attr-defined]
                return {"ok": True}
            return {"ok": False, "error": "not_found"}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def main() -> None:
    _ensure_dirs()
    _configure_pythonnet_coreclr()
    import webview  # local import so pythonnet env is configured first

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
    # NOTE:
    # Closing-event confirmation caused intermittent crashes on some environments.
    # Keep shutdown path minimal/stable; saving is handled during normal operations.
    webview.start(debug=False)


if __name__ == "__main__":
    main()





