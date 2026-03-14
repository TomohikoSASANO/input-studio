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
import html
import re
import time
import asyncio
import zipfile
import threading
from datetime import datetime, timezone
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

app = FastAPI(title="PDF Template Builder Web API")

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
_MAX_UPLOAD_BYTES = int(os.environ.get("INPUTSTUDIO_MAX_UPLOAD_MB", "500")) * 1024 * 1024
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

_SITE_BASE_URL = str(
    os.environ.get("INPUTSTUDIO_SITE_BASE_URL", "https://pdf-input-studio.kanazawa-application-support.jp")
).rstrip("/")
_SEO_LOCALES = ["ja", "en", "zh", "hi", "es", "fr", "ar", "pt", "ru", "bn", "id", "ur", "de", "it", "tr", "vi", "ko", "fa", "th", "pl", "uk", "nl"]
_CASE_STATIC_LOCALES = {"ja", "en", "zh"}
_CASE_DYNAMIC_LOCALES = [loc for loc in _SEO_LOCALES if loc not in _CASE_STATIC_LOCALES]
_LP_STATIC_LOCALES = {"en"}
_LP_DYNAMIC_LOCALES = [loc for loc in _SEO_LOCALES if loc not in _LP_STATIC_LOCALES]
_LOCALE_CACHE: dict[str, dict[str, str]] = {}
_LANG_SYNC_TAG = '<script defer src="/lang-sync.js"></script>'


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


def _abs_url(path: str) -> str:
    if not path.startswith("/"):
        path = "/" + path
    return f"{_SITE_BASE_URL}{path}"


def _iso_mtime(path: Path) -> str:
    try:
        dt = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)
    except Exception:
        dt = datetime.now(timezone.utc)
    return dt.strftime("%Y-%m-%d")


def _inject_lang_sync(html_text: str) -> str:
    """Inject lang-sync script into HTML once."""
    text = str(html_text or "")
    if "/lang-sync.js" in text:
        return text
    lower = text.lower()
    idx = lower.rfind("</body>")
    if idx >= 0:
        return text[:idx] + _LANG_SYNC_TAG + "\n" + text[idx:]
    return text + "\n" + _LANG_SYNC_TAG + "\n"


def _html_response(body: str) -> Response:
    return Response(content=_inject_lang_sync(body), media_type="text/html")


def _locale_text(loc: str, key: str, fallback: str) -> str:
    """Read translated text from ui/locales/{loc}.json with fallback."""
    code = (loc or "").strip().lower()
    if not code:
        return fallback
    if code not in _LOCALE_CACHE:
        fp = ROOT / "ui" / "locales" / f"{code}.json"
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
            _LOCALE_CACHE[code] = data if isinstance(data, dict) else {}
        except Exception:
            _LOCALE_CACHE[code] = {}
    value = _LOCALE_CACHE.get(code, {}).get(key)
    if isinstance(value, str) and value.strip():
        return value.strip()
    return fallback


def _render_case_hub_html(lang: str) -> str:
    l = lang.lower()
    title = _locale_text(l, "top.nav.cases", "Case Studies")
    checklist = _locale_text(l, "top.nav.checklist", "Pre-submission Checklist")
    tagrules = _locale_text(l, "top.nav.tagrules", "Tag Design Rules")
    updates = _locale_text(l, "top.nav.updates", "Updates")
    guide = _locale_text(l, "top.nav.guideFull", "How-to Guide")
    home = _locale_text(l, "main.backToTop", "Back to top")
    safe_title = html.escape(title)
    safe_home = html.escape(home)
    safe_checklist = html.escape(checklist)
    safe_tagrules = html.escape(tagrules)
    safe_updates = html.escape(updates)
    safe_guide = html.escape(guide)
    return f"""<!doctype html>
<html lang="{html.escape(l)}">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{safe_title} | PDF Template Builder</title>
    <meta name="description" content="Practical case studies for PDF Template Builder: recurring form operations, batch invoice workflows, and team handoff processes." />
    <meta name="robots" content="index,follow" />
    <link rel="canonical" href="{html.escape(_abs_url(f"/case-studies-{l}.html"))}" />
    <style>
      :root {{ --bg:#f8f7ff; --card:#ffffff; --text:#1f2330; --muted:#5b6477; --line:#dfe3ef; --accent:#7c5cff; }}
      * {{ box-sizing: border-box; }}
      body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif; color:var(--text); background:var(--bg); }}
      .wrap {{ max-width: 1040px; margin:0 auto; padding:24px 16px 44px; }}
      .home {{ display:inline-block; margin-bottom:14px; color:var(--accent); text-decoration:none; font-weight:700; }}
      .card {{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:22px; box-shadow:0 8px 24px rgba(15,23,42,.06); }}
      h1 {{ margin:0 0 10px; font-size:30px; }}
      p {{ color:var(--muted); line-height:1.75; }}
      .grid {{ margin-top:12px; display:grid; grid-template-columns: repeat(auto-fit,minmax(280px,1fr)); gap:12px; }}
      .item {{ border:1px solid var(--line); border-radius:12px; padding:14px; background:#fff; }}
      .item h2 {{ margin:0 0 6px; font-size:18px; }}
      .item a {{ color:var(--accent); text-decoration:none; font-weight:700; }}
      .item a:hover {{ text-decoration:underline; }}
      .meta {{ font-size:12px; color:#667085; margin-top:6px; }}
      .nav {{ margin-top:20px; display:flex; flex-wrap:wrap; gap:8px; }}
      .nav a {{ text-decoration:none; color:var(--text); border:1px solid var(--line); border-radius:999px; padding:7px 12px; background:#fff; font-size:13px; }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <a class="home" href="/template-builder.html">← {safe_home}</a>
      <article class="card">
        <h1>{safe_title}</h1>
        <p>These practical examples show how teams run recurring document operations with reusable project ZIP workflows.</p>
        <div class="grid">
          <section class="item">
            <h2>Monthly Application Updates</h2>
            <p>How teams reuse project ZIP files and tag sync for monthly application forms.</p>
            <a href="/case-application-monthly-{html.escape(l)}.html">Read case study</a>
            <div class="meta">Use case: recurring forms</div>
          </section>
          <section class="item">
            <h2>Batch Invoice Processing</h2>
            <p>How back-office staff produce multiple invoice PDFs with fewer mistakes.</p>
            <a href="/case-invoice-batch-{html.escape(l)}.html">Read case study</a>
            <div class="meta">Use case: accounting workflow</div>
          </section>
          <section class="item">
            <h2>Team Handoff Workflow</h2>
            <p>How teams reduce rework during operator handoff with shared project ZIP operations.</p>
            <a href="/case-team-handoff-{html.escape(l)}.html">Read case study</a>
            <div class="meta">Use case: team collaboration</div>
          </section>
        </div>
        <div class="nav">
          <a href="/case-studies-en.html">Case Studies (EN)</a>
          <a href="/case-studies-zh.html">Case Studies (ZH)</a>
          <a href="/template-automation.html">Template Automation</a>
          <a href="/pdf-template-workflow.html">Team Workflow</a>
          <a href="/beginner-guide.html">{safe_guide}</a>
          <a href="/document-quality-checklist.html">{safe_checklist}</a>
          <a href="/tag-design-rules.html">{safe_tagrules}</a>
          <a href="/updates.html">{safe_updates}</a>
        </div>
      </article>
    </div>
  </body>
</html>
"""


def _render_case_detail_html(lang: str, page_key: str) -> str:
    l = lang.lower()
    title_map = {
        "application-monthly": "Case Study: Monthly Application Updates",
        "invoice-batch": "Case Study: Batch Invoice Processing",
        "team-handoff": "Case Study: Team Handoff Workflow",
    }
    desc_map = {
        "application-monthly": "How teams handle monthly application updates with reusable ZIP projects and tag-sync workflows.",
        "invoice-batch": "How back-office teams process multiple invoice PDFs efficiently using tag updates and page-control operations.",
        "team-handoff": "How teams reduce rework during operator handoff with project ZIP-based document operations and shared tag conventions.",
    }
    scenario_map = {
        "application-monthly": "A team submits similar application forms every month with updated dates, values, and applicant information.",
        "invoice-batch": "Finance teams create many invoice PDFs at month end for different clients and billing structures.",
        "team-handoff": "Ongoing PDF projects are handed from one operator to another without losing context.",
    }
    workflow_map = {
        "application-monthly": [
            "Open the previous month project ZIP",
            "Update monthly values in tag list (date, amounts, applicant fields)",
            "Verify page sequence and supporting attachments",
            "Run pre-submission checklist and export final PDF",
            "Save next-month reusable ZIP package",
        ],
        "invoice-batch": [
            "Start from a common invoice template ZIP",
            "Update client-specific tag values only",
            "Append supporting pages and reorder before export",
            "Run checklist validation for numbers and formatting",
            "Export per-client final PDFs",
        ],
        "team-handoff": [
            "Save a project ZIP at each handoff checkpoint",
            "Normalize tags using shared naming rules",
            "Share checklist and expected output format",
            "Next operator reopens ZIP, applies delta updates, and re-saves",
        ],
    }
    challenges_map = {
        "application-monthly": [
            "Frequent copy-paste mistakes and stale values from previous month",
            "Repeated fields across pages updated manually",
            "Handoff quality drops when the operator changes",
        ],
        "invoice-batch": [
            "Large amount of repetitive updates (client name, invoice ID, due date)",
            "Attachments and page order differ by client",
            "Operator-dependent steps cause quality inconsistency",
        ],
        "team-handoff": [
            "Hard to know what is already done",
            "Inconsistent tag naming across operators",
            "Confusion between draft and final outputs",
        ],
    }
    outcome_map = {
        "application-monthly": [
            "Lower update-miss rate for repeated fields",
            "Faster correction cycle when forms are returned",
            "More stable quality during operator handoff",
        ],
        "invoice-batch": [
            "Standardized invoice generation process",
            "Lower risk of copy/paste and attachment mistakes",
            "Less month-end workload pressure",
        ],
        "team-handoff": [
            "Clear project state visibility",
            "Less rework and verification overhead",
            "Stable quality across operator changes",
        ],
    }
    path_map = {
        "application-monthly": f"/case-application-monthly-{l}.html",
        "invoice-batch": f"/case-invoice-batch-{l}.html",
        "team-handoff": f"/case-team-handoff-{l}.html",
    }
    if page_key not in title_map:
        raise HTTPException(status_code=404, detail="Not found")
    safe_title = html.escape(title_map[page_key])
    safe_desc = html.escape(desc_map[page_key])
    safe_scenario = html.escape(scenario_map[page_key])
    home = html.escape(_locale_text(l, "top.nav.cases", "Case Studies"))
    scenario_lbl = "Scenario"
    challenge_lbl = "Challenges"
    workflow_lbl = "Workflow"
    outcome_lbl = "Outcome"
    rows_challenges = "\n".join(f"<li>{html.escape(v)}</li>" for v in challenges_map[page_key])
    rows_workflow = "\n".join(f"<li>{html.escape(v)}</li>" for v in workflow_map[page_key])
    rows_outcome = "\n".join(f"<li>{html.escape(v)}</li>" for v in outcome_map[page_key])
    return f"""<!doctype html>
<html lang="{html.escape(l)}">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{safe_title} | PDF Template Builder</title>
    <meta name="description" content="{safe_desc}" />
    <meta name="robots" content="index,follow" />
    <link rel="canonical" href="{html.escape(_abs_url(path_map[page_key]))}" />
    <style>
      :root {{ --bg:#f8f7ff; --card:#fff; --text:#1f2330; --muted:#5b6477; --line:#dfe3ef; --accent:#7c5cff; }}
      * {{ box-sizing: border-box; }}
      body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif; color:var(--text); background:var(--bg); }}
      .wrap {{ max-width:980px; margin:0 auto; padding:24px 16px 44px; }}
      .home {{ display:inline-block; margin-bottom:14px; color:var(--accent); text-decoration:none; font-weight:700; }}
      .card {{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:22px; box-shadow:0 8px 24px rgba(15,23,42,.06); }}
      h1 {{ margin:0 0 10px; font-size:30px; }} h2 {{ margin:20px 0 8px; font-size:20px; }} p, li {{ color:var(--muted); line-height:1.75; }}
      ul, ol {{ margin:8px 0 0 20px; padding:0; }}
      .nav {{ margin-top:20px; display:flex; flex-wrap:wrap; gap:8px; }}
      .nav a {{ text-decoration:none; color:var(--text); border:1px solid var(--line); border-radius:999px; padding:7px 12px; background:#fff; font-size:13px; }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <a class="home" href="/case-studies-{html.escape(l)}.html">← {home}</a>
      <article class="card">
        <h1>{safe_title}</h1>
        <p><strong>{scenario_lbl}:</strong> {safe_scenario}</p>
        <h2>{challenge_lbl}</h2>
        <ul>{rows_challenges}</ul>
        <h2>{workflow_lbl}</h2>
        <ol>{rows_workflow}</ol>
        <h2>{outcome_lbl}</h2>
        <ul>{rows_outcome}</ul>
        <div class="nav">
          <a href="/case-studies-en.html">Case Studies (EN)</a>
          <a href="/case-studies-zh.html">Case Studies (ZH)</a>
          <a href="/template-builder.html">Template Builder</a>
          <a href="/pricing.html">Pricing</a>
        </div>
      </article>
    </div>
  </body>
</html>
"""


def _render_template_automation_html(lang: str) -> str:
    l = lang.lower()
    guide = html.escape(_locale_text(l, "top.nav.guideFull", "How-to Guide"))
    updates = html.escape(_locale_text(l, "top.nav.updates", "Updates"))
    cases = html.escape(_locale_text(l, "top.nav.cases", "Case Studies"))
    checklist = html.escape(_locale_text(l, "top.nav.checklist", "Pre-submission Checklist"))
    tagrules = html.escape(_locale_text(l, "top.nav.tagrules", "Tag Design Rules"))
    home = html.escape(_locale_text(l, "main.backToTop", "Back to top"))
    title = "Document Template Automation | PDF Template Builder"
    description = "Document template automation for teams using reusable PDF templates, tag sync, and project ZIP handoff workflows."
    return f"""<!doctype html>
<html lang="{html.escape(l)}">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{html.escape(title)}</title>
    <meta name="description" content="{html.escape(description)}" />
    <meta name="robots" content="index,follow" />
    <link rel="canonical" href="{html.escape(_abs_url(f"/template-automation-{l}.html"))}" />
    <style>
      :root {{ --bg:#f7f8ff; --card:#fff; --text:#1f2330; --muted:#5c6578; --line:#dfe3ef; --accent:#5b56f0; }}
      * {{ box-sizing: border-box; }}
      body {{ margin: 0; background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; }}
      .wrap {{ max-width: 980px; margin: 0 auto; padding: 24px 16px 42px; }}
      .nav {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 14px; }}
      .nav a {{ text-decoration: none; color: var(--text); background: #fff; border: 1px solid var(--line); border-radius: 999px; padding: 7px 12px; font-size: 13px; }}
      .card {{ background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 20px; }}
      h1 {{ margin: 0 0 10px; font-size: 30px; }}
      h2 {{ margin: 20px 0 8px; font-size: 20px; }}
      p, li {{ color: var(--muted); line-height: 1.75; }}
      ul {{ margin: 8px 0 0 18px; }}
      .cta {{ margin-top: 12px; display: flex; flex-wrap: wrap; gap: 10px; }}
      .btn {{ text-decoration: none; border-radius: 10px; padding: 10px 14px; font-weight: 700; }}
      .btn--primary {{ background: var(--accent); color: #fff; }}
      .btn--soft {{ background: #fff; color: var(--text); border: 1px solid var(--line); }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="nav">
        <a href="/template-builder.html">Template Builder</a>
        <a href="/pdf-template-workflow-{html.escape(l)}.html">Team Workflow</a>
        <a href="/case-studies-{html.escape(l)}.html">{cases}</a>
        <a href="/">{home}</a>
      </div>
      <article class="card">
        <h1>Document Template Automation for Teams</h1>
        <p>
          PDF Template Builder helps teams automate repetitive document updates.
          Instead of editing the same fields many times, define reusable tags and update values in one place.
        </p>
        <h2>What gets automated</h2>
        <ul>
          <li>Repeated form fields across multiple pages and files</li>
          <li>Project-level handoff using ZIP bundles</li>
          <li>Combined workflows with PDF merge and split operations</li>
        </ul>
        <h2>Who benefits most</h2>
        <ul>
          <li>Operations teams handling recurring applications</li>
          <li>Back-office staff producing standardized documents</li>
          <li>Teams that need fast review-ready PDF outputs</li>
        </ul>
        <div class="cta">
          <a class="btn btn--primary" href="/?lang={html.escape(l)}">Open App</a>
          <a class="btn btn--soft" href="/template-builder.html">Back to Product Page</a>
        </div>
        <div class="nav" style="margin-top:16px;">
          <a href="/beginner-guide.html">{guide}</a>
          <a href="/document-quality-checklist.html">{checklist}</a>
          <a href="/tag-design-rules.html">{tagrules}</a>
          <a href="/updates.html">{updates}</a>
        </div>
      </article>
    </div>
  </body>
</html>
"""


def _render_template_workflow_html(lang: str) -> str:
    l = lang.lower()
    cases = html.escape(_locale_text(l, "top.nav.cases", "Case Studies"))
    home = html.escape(_locale_text(l, "main.backToTop", "Back to top"))
    title = "PDF Template Workflow for Teams | PDF Template Builder"
    description = "Learn a practical PDF template workflow for teams: build reusable templates, sync repeated values, and deliver consistent outputs faster."
    return f"""<!doctype html>
<html lang="{html.escape(l)}">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{html.escape(title)}</title>
    <meta name="description" content="{html.escape(description)}" />
    <meta name="robots" content="index,follow" />
    <link rel="canonical" href="{html.escape(_abs_url(f"/pdf-template-workflow-{l}.html"))}" />
    <style>
      :root {{ --bg:#f7f8ff; --card:#fff; --text:#1f2330; --muted:#5c6578; --line:#dfe3ef; --accent:#5b56f0; }}
      * {{ box-sizing: border-box; }}
      body {{ margin: 0; background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; }}
      .wrap {{ max-width: 980px; margin: 0 auto; padding: 24px 16px 42px; }}
      .nav {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 14px; }}
      .nav a {{ text-decoration: none; color: var(--text); background: #fff; border: 1px solid var(--line); border-radius: 999px; padding: 7px 12px; font-size: 13px; }}
      .card {{ background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 20px; }}
      h1 {{ margin: 0 0 10px; font-size: 30px; }}
      h2 {{ margin: 20px 0 8px; font-size: 20px; }}
      p, li {{ color: var(--muted); line-height: 1.75; }}
      ol {{ margin: 8px 0 0 18px; }}
      .cta {{ margin-top: 12px; display: flex; flex-wrap: wrap; gap: 10px; }}
      .btn {{ text-decoration: none; border-radius: 10px; padding: 10px 14px; font-weight: 700; }}
      .btn--primary {{ background: var(--accent); color: #fff; }}
      .btn--soft {{ background: #fff; color: var(--text); border: 1px solid var(--line); }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="nav">
        <a href="/template-builder.html">Template Builder</a>
        <a href="/template-automation-{html.escape(l)}.html">Template Automation</a>
        <a href="/case-studies-{html.escape(l)}.html">{cases}</a>
        <a href="/">{home}</a>
      </div>
      <article class="card">
        <h1>Team PDF Template Workflow</h1>
        <p>
          A practical workflow for recurring form operations. Keep templates reusable,
          reduce manual edits, and maintain consistent outputs across your team.
        </p>
        <h2>Recommended 5-step flow</h2>
        <ol>
          <li>Start from a source PDF and place reusable tags</li>
          <li>Define values once and sync repeated fields</li>
          <li>Use PDF merge/split to shape deliverables</li>
          <li>Save the whole job as project ZIP for handoff</li>
          <li>Reopen ZIP for updates and regenerate outputs quickly</li>
        </ol>
        <h2>Operational benefits</h2>
        <ol>
          <li>Lower rework for repetitive updates</li>
          <li>Faster onboarding with repeatable process</li>
          <li>More consistent final PDF quality</li>
        </ol>
        <div class="cta">
          <a class="btn btn--primary" href="/?lang={html.escape(l)}">Try the App</a>
          <a class="btn btn--soft" href="/template-builder.html">Product Overview</a>
        </div>
      </article>
    </div>
  </body>
</html>
"""


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
        try:
            return _html_response(index_path.read_text(encoding="utf-8"))
        except Exception:
            return FileResponse(str(index_path), media_type="text/html")
    return {"message": "PDF Template Builder Web API - UI not found"}


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


@app.get("/lang-sync.js")
async def lang_sync_js():
    """Keep selected language across internal page navigation."""
    js = r"""
(function () {
  var SUPPORTED = new Set(["ja","en","zh","hi","es","fr","ar","pt","ru","bn","id","ur","de","it","tr","vi","ko","fa","th","pl","uk","nl"]);
  var KEY = "inputstudio-locale";
  var current = "ja";

  try {
    var q = new URLSearchParams(window.location.search || "");
    var fromQuery = String(q.get("lang") || "").toLowerCase().trim();
    if (SUPPORTED.has(fromQuery)) {
      current = fromQuery;
      localStorage.setItem(KEY, current);
    } else {
      var fromStorage = String(localStorage.getItem(KEY) || "").toLowerCase().trim();
      if (SUPPORTED.has(fromStorage)) {
        current = fromStorage;
      } else {
        var fromHtmlLang = String(document.documentElement.lang || "").toLowerCase().trim();
        if (SUPPORTED.has(fromHtmlLang)) current = fromHtmlLang;
      }
    }
  } catch (e) {}

  try { document.documentElement.lang = current; } catch (e) {}

  function mappedPath(pathname, lang) {
    var p = String(pathname || "");
    var maps = [
      [/^\/case-studies(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/case-studies-" + lang + ".html"],
      [/^\/case-application-monthly(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/case-application-monthly-" + lang + ".html"],
      [/^\/case-invoice-batch(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/case-invoice-batch-" + lang + ".html"],
      [/^\/case-team-handoff(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/case-team-handoff-" + lang + ".html"],
      [/^\/template-automation(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/template-automation-" + lang + ".html"],
      [/^\/pdf-template-workflow(?:-(?:ja|en|zh|hi|es|fr|ar|pt|ru|bn|id|ur|de|it|tr|vi|ko|fa|th|pl|uk|nl))?\.html$/i, "/pdf-template-workflow-" + lang + ".html"]
    ];
    for (var i = 0; i < maps.length; i += 1) {
      if (maps[i][0].test(p)) return maps[i][1];
    }
    return p;
  }

  function rewrite(anchor) {
    if (!anchor || !anchor.getAttribute) return;
    var raw = anchor.getAttribute("href");
    if (!raw) return;
    if (raw[0] === "#" || raw.indexOf("mailto:") === 0 || raw.indexOf("tel:") === 0 || raw.indexOf("javascript:") === 0) return;
    var url;
    try { url = new URL(raw, window.location.href); } catch (e) { return; }
    if (url.origin !== window.location.origin) return;
    var explicitLang = String(url.searchParams.get("lang") || "").toLowerCase().trim();
    var hasExplicitLang = SUPPORTED.has(explicitLang);
    var targetLang = hasExplicitLang ? explicitLang : current;
    url.pathname = mappedPath(url.pathname, targetLang);
    if (url.pathname === "/") {
      if (!hasExplicitLang) url.searchParams.set("lang", targetLang);
    } else if (url.pathname.endsWith(".html")) {
      if (!hasExplicitLang) url.searchParams.set("lang", targetLang);
    }
    var out = url.pathname + (url.search || "") + (url.hash || "");
    anchor.setAttribute("href", out);
  }

  var links = document.querySelectorAll("a[href]");
  for (var i = 0; i < links.length; i += 1) rewrite(links[i]);
})();
"""
    response = Response(content=js, media_type="application/javascript")
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/robots.txt")
async def robots_txt():
    body = (
        "User-agent: *\n"
        "Allow: /\n"
        f"Sitemap: {_abs_url('/sitemap.xml')}\n"
    )
    return Response(content=body, media_type="text/plain")


def _case_static_file(stem: str, lang: str) -> str:
    if stem == "hub":
        return {"ja": "case-studies.html", "en": "case-studies-en.html", "zh": "case-studies-zh.html"}[lang]
    if stem == "application-monthly":
        return {"ja": "case-application-monthly.html", "en": "case-application-monthly-en.html", "zh": "case-application-monthly-zh.html"}[lang]
    if stem == "invoice-batch":
        return {"ja": "case-invoice-batch.html", "en": "case-invoice-batch-en.html", "zh": "case-invoice-batch-zh.html"}[lang]
    if stem == "team-handoff":
        return {"ja": "case-team-handoff.html", "en": "case-team-handoff-en.html", "zh": "case-team-handoff-zh.html"}[lang]
    raise HTTPException(status_code=404, detail="Not found")


def _lp_static_file(stem: str, lang: str) -> str:
    if lang != "en":
        raise HTTPException(status_code=404, detail="Not found")
    if stem == "template-automation":
        return "template-automation.html"
    if stem == "pdf-template-workflow":
        return "pdf-template-workflow.html"
    raise HTTPException(status_code=404, detail="Not found")


@app.get("/case-studies-{lang}.html")
async def case_studies_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _CASE_STATIC_LOCALES:
        fp = ui_dir / _case_static_file("hub", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_case_hub_html(code))


@app.get("/case-application-monthly-{lang}.html")
async def case_application_monthly_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _CASE_STATIC_LOCALES:
        fp = ui_dir / _case_static_file("application-monthly", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_case_detail_html(code, "application-monthly"))


@app.get("/case-invoice-batch-{lang}.html")
async def case_invoice_batch_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _CASE_STATIC_LOCALES:
        fp = ui_dir / _case_static_file("invoice-batch", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_case_detail_html(code, "invoice-batch"))


@app.get("/case-team-handoff-{lang}.html")
async def case_team_handoff_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _CASE_STATIC_LOCALES:
        fp = ui_dir / _case_static_file("team-handoff", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_case_detail_html(code, "team-handoff"))


@app.get("/template-automation-{lang}.html")
async def template_automation_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _LP_STATIC_LOCALES:
        fp = ui_dir / _lp_static_file("template-automation", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_template_automation_html(code))


@app.get("/pdf-template-workflow-{lang}.html")
async def pdf_template_workflow_by_lang(lang: str):
    code = str(lang or "").strip().lower()
    if code not in _SEO_LOCALES:
        raise HTTPException(status_code=404, detail="Not found")
    if code in _LP_STATIC_LOCALES:
        fp = ui_dir / _lp_static_file("pdf-template-workflow", code)
        if fp.exists():
            try:
                return _html_response(fp.read_text(encoding="utf-8"))
            except Exception:
                return FileResponse(str(fp), media_type="text/html")
    return _html_response(_render_template_workflow_html(code))


@app.get("/sitemap.xml")
async def sitemap_xml():
    page_paths = [
        "/",
        "/about.html",
        "/contact.html",
        "/faq.html",
        "/privacy.html",
        "/terms.html",
        "/solutions.html",
        "/application-form-filling.html",
        "/pdf-merge-split.html",
        "/global-search.html",
        "/pricing.html",
        "/template-builder.html",
        "/template-automation.html",
        "/pdf-template-workflow.html",
        "/beginner-guide.html",
        "/updates.html",
        "/document-quality-checklist.html",
        "/tag-design-rules.html",
        "/case-studies.html",
        "/case-application-monthly.html",
        "/case-invoice-batch.html",
        "/case-team-handoff.html",
        "/case-studies-en.html",
        "/case-application-monthly-en.html",
        "/case-invoice-batch-en.html",
        "/case-team-handoff-en.html",
        "/case-studies-zh.html",
        "/case-application-monthly-zh.html",
        "/case-invoice-batch-zh.html",
        "/case-team-handoff-zh.html",
    ]
    pages: list[tuple[str, str]] = []
    for path in page_paths:
        file_name = "index.html" if path == "/" else path.lstrip("/")
        fp = ui_dir / file_name
        pages.append((_abs_url(path), _iso_mtime(fp)))
    idx_lastmod = _iso_mtime(ui_dir / "index.html")
    for loc in _CASE_DYNAMIC_LOCALES:
        pages.append((_abs_url(f"/case-studies-{loc}.html"), idx_lastmod))
        pages.append((_abs_url(f"/case-application-monthly-{loc}.html"), idx_lastmod))
        pages.append((_abs_url(f"/case-invoice-batch-{loc}.html"), idx_lastmod))
        pages.append((_abs_url(f"/case-team-handoff-{loc}.html"), idx_lastmod))
    for loc in _LP_DYNAMIC_LOCALES:
        pages.append((_abs_url(f"/template-automation-{loc}.html"), idx_lastmod))
        pages.append((_abs_url(f"/pdf-template-workflow-{loc}.html"), idx_lastmod))
    for loc in _SEO_LOCALES:
        pages.append((f"{_abs_url('/')}?lang={loc}", _iso_mtime(ui_dir / "index.html")))
    items = "\n".join(
        f"<url><loc>{loc}</loc><lastmod>{lastmod}</lastmod></url>" for (loc, lastmod) in pages
    )
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        f"{items}\n"
        "</urlset>\n"
    )
    return Response(content=xml, media_type="application/xml")


def _ads_txt_line_from_client(client_id: str) -> str:
    """Build an ads.txt line from AdSense client ID if possible."""
    cid = str(client_id or "").strip()
    if not cid:
        return ""
    m = re.search(r"(?:ca-)?pub-(\d+)", cid, flags=re.IGNORECASE)
    if not m:
        return ""
    pub_id = m.group(1)
    # Google official seller line format
    return f"google.com, pub-{pub_id}, DIRECT, f08c47fec0942fa0"


@app.get("/ads.txt")
async def ads_txt():
    """
    Serve ads.txt for AdSense domain verification.
    Priority:
      1) INPUTSTUDIO_ADS_TXT_LINE (full custom line)
      2) Derived from INPUTSTUDIO_ADSENSE_CLIENT
    """
    line = str(os.environ.get("INPUTSTUDIO_ADS_TXT_LINE", "")).strip()
    if not line:
        line = _ads_txt_line_from_client(str(os.environ.get("INPUTSTUDIO_ADSENSE_CLIENT", "")).strip())
    if not line:
        raise HTTPException(status_code=404, detail="ads.txt not configured")
    return Response(content=line + "\n", media_type="text/plain")

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
                try:
                    return _html_response(index_path.read_text(encoding="utf-8"))
                except Exception:
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
            if ext == ".html":
                try:
                    return _html_response(full_path.read_text(encoding="utf-8"))
                except Exception:
                    return FileResponse(str(full_path), media_type=media_type)
            return FileResponse(str(full_path), media_type=media_type)
        
        raise HTTPException(status_code=404, detail="File not found")


if __name__ == "__main__":
    # Ensure directories exist
    _ensure_dirs()
    
    port = int(os.environ.get("PORT", 8001))  # Changed default to 8001 to avoid conflicts
    uvicorn.run(app, host="0.0.0.0", port=port)
