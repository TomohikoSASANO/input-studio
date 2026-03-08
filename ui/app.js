const $ = (sel) => document.querySelector(sel)
const DEFAULT_LETTER_SPACING = 1.2
const isWeb = () => !!(typeof window !== 'undefined' && window.__INPUTSTUDIO_WEB__)
const tr = (key, fallback, vars = {}) => {
  const fn = window.i18n?.t
  if (typeof fn !== "function") return String(fallback || key)
  const out = fn(key, vars)
  if (out === key && fallback) return String(fallback)
  return String(out)
}
const getLocaleSafe = () => {
  const fn = window.i18n?.getLocale
  if (typeof fn !== "function") return "ja"
  return fn() || "ja"
}
const LOCALE_OPTIONS = [
  { code: "ja", label: "日本語 / Japanese", flag: "jp" },
  { code: "en", label: "English / 英語", flag: "us" },
  { code: "zh", label: "中文 / Chinese", flag: "cn" },
  { code: "hi", label: "हिन्दी / Hindi", flag: "in" },
  { code: "es", label: "Español / Spanish", flag: "es" },
  { code: "fr", label: "Français / French", flag: "fr" },
  { code: "ar", label: "العربية / Arabic", flag: "sa" },
  { code: "pt", label: "Português / Portuguese", flag: "br" },
  { code: "ru", label: "Русский / Russian", flag: "ru" },
  { code: "bn", label: "বাংলা / Bengali", flag: "bd" },
  { code: "id", label: "Bahasa Indonesia / Indonesian", flag: "id" },
  { code: "ur", label: "اردو / Urdu", flag: "pk" },
  { code: "de", label: "Deutsch / German", flag: "de" },
  { code: "it", label: "Italiano / Italian", flag: "it" },
  { code: "tr", label: "Türkçe / Turkish", flag: "tr" },
  { code: "vi", label: "Tiếng Việt / Vietnamese", flag: "vi" },
  { code: "ko", label: "한국어 / Korean", flag: "kr" },
  { code: "fa", label: "فارسی / Persian", flag: "ir" },
  { code: "th", label: "ไทย / Thai", flag: "th" },
  { code: "pl", label: "Polski / Polish", flag: "pl" },
  { code: "uk", label: "Українська / Ukrainian", flag: "ua" },
  { code: "nl", label: "Nederlands / Dutch", flag: "nl" },
]
const getLocaleMeta = (locale) => {
  return LOCALE_OPTIONS.find((x) => x.code === String(locale || "").toLowerCase()) || LOCALE_OPTIONS[0]
}
function getLocaleFromQuery() {
  try {
    const q = new URLSearchParams(window.location.search || "")
    const v = String(q.get("lang") || "").trim().toLowerCase()
    if (!v) return null
    const hit = LOCALE_OPTIONS.find((x) => x.code === v)
    return hit ? hit.code : null
  } catch {
    return null
  }
}
function syncLocaleQuery(locale) {
  if (!isWeb()) return
  try {
    const url = new URL(window.location.href)
    url.searchParams.set("lang", String(locale || "ja").toLowerCase())
    window.history.replaceState({}, "", url.toString())
  } catch {}
}
let _viewportMetricsBound = false
function syncViewportMetrics() {
  const root = document.documentElement
  if (!root) return
  const vv = window.visualViewport
  const rawViewportH = Number(vv?.height || window.innerHeight || document.documentElement.clientHeight || 0)
  const viewportH = Math.max(1, Math.round(rawViewportH))
  let usableH = viewportH
  // Desktop app can report taller than work area; clamp to OS available height (taskbar excluded).
  if (!isWeb()) {
    const avail = Number(window.screen?.availHeight || 0)
    if (avail > 0) usableH = Math.min(usableH, Math.round(avail))
  }
  root.style.setProperty("--inputstudio-usable-vh", `${usableH}px`)
}

function bindViewportMetricsOnce() {
  if (_viewportMetricsBound) return
  _viewportMetricsBound = true
  syncViewportMetrics()
  window.addEventListener("resize", syncViewportMetrics, { passive: true })
  window.visualViewport?.addEventListener?.("resize", syncViewportMetrics, { passive: true })
}

const AD_UNLOCK_RULES = {
  zip_open: { cooldownMs: 5 * 60 * 1000, maxPerSession: 2 },
  pdf_append: { cooldownMs: 3 * 60 * 1000, maxPerSession: 3 },
}
const AD_SLOT_IDS = {
  gate: "adSlotGate",
  panel: "adSlotPanel",
  panelBottom: "adSlotPanelBottom",
}
const adRuntime = {
  scriptReady: false,
  scriptPromise: null,
}

function getAdConfig() {
  const raw = window.__INPUTSTUDIO_AD_CONFIG__ && typeof window.__INPUTSTUDIO_AD_CONFIG__ === "object"
    ? window.__INPUTSTUDIO_AD_CONFIG__
    : {}
  const adsense = raw.adsense && typeof raw.adsense === "object" ? raw.adsense : {}
  const slots = adsense.slots && typeof adsense.slots === "object" ? adsense.slots : {}
  const unlock = raw.unlock && typeof raw.unlock === "object" ? raw.unlock : {}
  return {
    enabled: !!raw.enabled && isWeb(),
    provider: String(raw.provider || "none").toLowerCase(),
    adsense: {
      client: String(adsense.client || "").trim(),
      slots: {
        gate: String(slots.gate || "").trim(),
        panel: String(slots.panel || "").trim(),
        panelBottom: String(slots.panelBottom || "").trim(),
        unlock: String(slots.unlock || "").trim(),
      },
    },
    unlock: {
      minSeconds: Math.max(0, Number(unlock.minSeconds || 3) || 3),
    },
  }
}

function getAdSlotFor(name) {
  const cfg = getAdConfig()
  return String(cfg.adsense?.slots?.[name] || "").trim()
}

async function ensureAdSenseScript() {
  const cfg = getAdConfig()
  if (!cfg.enabled || cfg.provider !== "adsense") return false
  if (!cfg.adsense.client) return false
  if (adRuntime.scriptReady) return true
  if (adRuntime.scriptPromise) return adRuntime.scriptPromise
  adRuntime.scriptPromise = new Promise((resolve) => {
    const src = `https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=${encodeURIComponent(cfg.adsense.client)}`
    const exists = Array.from(document.querySelectorAll("script")).find((s) => String(s.src || "").includes("pagead/js/adsbygoogle.js"))
    if (exists) {
      adRuntime.scriptReady = true
      resolve(true)
      return
    }
    const s = document.createElement("script")
    s.async = true
    s.src = src
    s.crossOrigin = "anonymous"
    s.onload = () => {
      adRuntime.scriptReady = true
      resolve(true)
    }
    s.onerror = () => resolve(false)
    document.head.appendChild(s)
  })
  return adRuntime.scriptPromise
}

function mountAdSenseInto(container, slotId) {
  const cfg = getAdConfig()
  if (!container || !slotId || !cfg.adsense.client) return false
  const adId = `ad-${slotId}-${Date.now().toString(36)}`
  container.innerHTML = `
    <ins id="${escapeHtml(adId)}"
      class="adsbygoogle inputstudioAd"
      style="display:block"
      data-ad-client="${escapeHtml(cfg.adsense.client)}"
      data-ad-slot="${escapeHtml(slotId)}"
      data-ad-format="auto"
      data-full-width-responsive="true"></ins>
  `
  try {
    ;(window.adsbygoogle = window.adsbygoogle || []).push({})
    return true
  } catch {
    return false
  }
}

async function refreshAdSlots() {
  const cfg = getAdConfig()
  if (!cfg.enabled || cfg.provider !== "adsense") return
  const ok = await ensureAdSenseScript()
  if (!ok) return
  for (const [slotName, domId] of Object.entries(AD_SLOT_IDS)) {
    const host = document.getElementById(domId)
    if (!host) continue
    const slotId = getAdSlotFor(slotName)
    if (!slotId) continue
    const live = host.querySelector(".adSlot__live")
    if (!live) continue
    if (live.dataset.liveMounted === "1") continue
    const mounted = mountAdSenseInto(live, slotId)
    if (mounted) live.dataset.liveMounted = "1"
  }
}

// --- Web demo mode (GitHub Pages) ------------------------------------------
// GitHub上で「実画面レビュー」を回すため、pywebviewが無い環境では
// 画面を動かせるモックAPIを注入する。
;(function ensureDemoApi() {
  // Desktop app (pywebview + WebView2) may not have window.pywebview at initial parse.
  // Detect desktop reliably and NEVER inject the demo mock there.
  try {
    // WebView2 exposes window.chrome.webview
    if (window.chrome && window.chrome.webview) return
    const host = String(window.location?.hostname || "")
    if (host === "127.0.0.1" || host === "localhost") return
  } catch {
    return
  }
  // If pywebview exists at all, assume desktop and DO NOT inject the mock.
  if (window.pywebview) return
  window.__INPUTSTUDIO_DEMO__ = true

  const demo = {
    projectName: "デモ案件：外國語書類一式",
    projectPath: "demo/project.json",
    pageCount: 58,
    uiMode: "worker",
    tags: [],
    values: {},
    placements: {}, // fid -> {tag,page,x,y,font_size,...}
    defaultFontSize: 14,
    viewZoom: 1.0,
  }

  const makeSvgDataUrl = (pageIndex) => {
    const w = 1240
    const h = 1754
    const p = pageIndex + 1
    const n = demo.pageCount
    const placed = Object.entries(demo.placements).filter(([, pl]) => Number(pl?.page || 0) === pageIndex)
    const overlay = placed
      .map(([, pl]) => {
        const tag = String(pl?.tag || "").trim()
        const v = String(demo.values[tag] || tag).replaceAll("<br>", "\n")
        const x = Math.max(24, Math.min(w - 24, Number(pl.x || 0)))
        const y = Math.max(24, Math.min(h - 24, Number(pl.y || 0)))
        const fs = Math.max(10, Math.min(36, Number(pl.font_size || 14)))
        const safe = v.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
        return `<text x="${x}" y="${y}" font-size="${fs}" fill="#0f172a" font-family="Arial, sans-serif">${safe}</text>`
      })
      .join("")

    const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
  <defs>
    <linearGradient id="bg" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#ffffff"/>
      <stop offset="1" stop-color="#f3f4ff"/>
    </linearGradient>
  </defs>
  <rect x="0" y="0" width="${w}" height="${h}" fill="url(#bg)"/>
  <rect x="40" y="40" width="${w - 80}" height="${h - 80}" fill="#ffffff" stroke="rgba(15,23,42,0.10)" stroke-width="2" rx="18"/>
  <text x="72" y="110" font-size="28" fill="rgba(15,23,42,0.75)" font-family="Arial, sans-serif">PDF Input Studio デモプレビュー</text>
  <text x="72" y="150" font-size="18" fill="rgba(15,23,42,0.55)" font-family="Arial, sans-serif">ページ ${p} / ${n}</text>
  <g opacity="0.18">
    <rect x="90" y="220" width="${w - 180}" height="${h - 320}" fill="none" stroke="#7c5cff" stroke-width="2" stroke-dasharray="10 10" rx="10"/>
  </g>
  ${overlay}
</svg>`
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  }

  const api = {
    async get_admin_settings() {
      return {
        ok: true,
        settings: {
          ui_mode: demo.uiMode,
          default_font_size: Number(demo.defaultFontSize || 14) || 14,
          view_zoom: Number(demo.viewZoom || 1.0) || 1.0,
        },
      }
    },
    async update_admin_settings(patch) {
      const p = patch && typeof patch === "object" ? patch : {}
      if (p.default_font_size != null) demo.defaultFontSize = Number(p.default_font_size || 14) || 14
      if (p.view_zoom != null) demo.viewZoom = Number(p.view_zoom || 1.0) || 1.0
      return {
        ok: true,
        settings: {
          ui_mode: demo.uiMode,
          default_font_size: Number(demo.defaultFontSize || 14) || 14,
          view_zoom: Number(demo.viewZoom || 1.0) || 1.0,
        },
      }
    },
    async get_workers() {
      return {
        ok: true,
        workers: [
          { id: "w_demo", name: "デモ作業者", bank: "" },
          { id: "w_demo2", name: "デモ作業者2", bank: "" },
        ],
        last_worker_id: "w_demo",
      }
    },
    async pick_project() {
      return { ok: true, path: demo.projectPath }
    },
    async pick_pdf() {
      return { ok: true, path: "demo.pdf" }
    },
    async create_project_from_pdf_simple() {
      demo.tags = []
      demo.values = {}
      demo.placements = {}
      return { ok: true, path: demo.projectPath }
    },
    async load_project() {
      return {
        ok: true,
        project: demo.projectName,
        tags: demo.tags,
        values: demo.values,
        placements: demo.placements,
        drop_dir: "demo/exports",
        ui_mode: demo.uiMode,
        page_count: demo.pageCount,
      }
    },
    async save_current_project() {
      return { ok: true }
    },
    async save_project_as(name) {
      demo.projectName = String(name || demo.projectName || "案件")
      demo.projectPath = "demo/project.json"
      return { ok: true, path: demo.projectPath }
    },
    async append_pdf_to_project() {
      demo.pageCount = Math.max(1, Number(demo.pageCount || 1) + 1)
      return { ok: true, page_count: demo.pageCount }
    },
    async copy_page_with_elements(page_index) {
      const idx = Math.max(0, Math.min(Math.max(1, Number(demo.pageCount || 1)) - 1, Number(page_index || 0)))
      const out = {}
      for (const [fid, pl] of Object.entries(demo.placements || {})) {
        if (!pl || typeof pl !== "object") continue
        const page = Number(pl.page || 0)
        const next = { ...pl }
        if (page > idx) next.page = page + 1
        out[fid] = next
        if (page === idx) {
          const nf = `f_${Date.now().toString(16)}_${Math.random().toString(16).slice(2, 8)}`
          out[nf] = { ...pl, page: idx + 1 }
        }
      }
      demo.placements = out
      demo.pageCount = Math.max(1, Number(demo.pageCount || 1) + 1)
      return { ok: true, page_count: demo.pageCount, page_index: idx + 1, placements: demo.placements }
    },
    async delete_page_from_project(page_index) {
      const total = Math.max(1, Number(demo.pageCount || 1))
      if (total <= 1) return { ok: false, error: "cannot_delete_last_page" }
      const idx = Math.max(0, Math.min(total - 1, Number(page_index || 0)))
      const out = {}
      for (const [fid, pl] of Object.entries(demo.placements || {})) {
        if (!pl || typeof pl !== "object") continue
        const page = Number(pl.page || 0)
        if (page === idx) continue
        out[fid] = { ...pl, page: page > idx ? page - 1 : page }
      }
      demo.placements = out
      demo.pageCount = Math.max(1, total - 1)
      return {
        ok: true,
        page_count: demo.pageCount,
        page_index: Math.max(0, Math.min(demo.pageCount - 1, idx)),
        tags: demo.tags,
        values: demo.values,
        placements: demo.placements,
      }
    },
    async reorder_pages(order) {
      const total = Math.max(1, Number(demo.pageCount || 1))
      if (!Array.isArray(order) || order.length !== total) return { ok: false, error: "invalid_order" }
      const norm = order.map((x) => Number(x))
      if (norm.some((x) => !Number.isFinite(x))) return { ok: false, error: "invalid_order" }
      const set = new Set(norm)
      if (set.size !== total) return { ok: false, error: "invalid_order" }
      const min = Math.min(...norm)
      const max = Math.max(...norm)
      if (min < 0 || max >= total) return { ok: false, error: "invalid_order" }
      const oldToNew = {}
      norm.forEach((oldIdx, newIdx) => {
        oldToNew[Number(oldIdx)] = Number(newIdx)
      })
      const out = {}
      for (const [fid, pl] of Object.entries(demo.placements || {})) {
        if (!pl || typeof pl !== "object") continue
        const oldPage = Number(pl.page || 0)
        out[fid] = { ...pl, page: Number(oldToNew[oldPage] ?? oldPage) }
      }
      demo.placements = out
      return { ok: true, page_count: demo.pageCount, placements: demo.placements }
    },
    async set_ui_mode(mode) {
      demo.uiMode = String(mode || "worker")
      return { ok: true }
    },
    async start_work() {
      return { ok: true }
    },
    async toggle_private() {
      return { ok: true, in_private: false }
    },
    async finish() {
      return { ok: true, dir: "demo/exports", zip: "demo.zip" }
    },
    async delete_worker() {
      return { ok: true }
    },
    async set_value(tag, value) {
      demo.values[String(tag)] = String(value ?? "")
      return { ok: true }
    },
    async add_text_field(tag, page, x, y, font_size) {
      const t = String(tag || "").trim()
      if (!t) return { ok: false, error: "missing_tag" }
      if (!demo.tags.includes(t)) demo.tags.push(t)
      const fid = `f_${Date.now().toString(16)}_${Math.random().toString(16).slice(2, 8)}`
      demo.placements[fid] = {
        tag: t,
        page: Number(page || 0),
        x: Number(x || 0),
        y: Number(y || 0),
        font_size: Number(font_size || 14),
        color: "#0f172a",
        line_height: 1.2,
        letter_spacing: DEFAULT_LETTER_SPACING,
      }
      return { ok: true, fid, tag: t }
    },
    async set_element_pos(fid, x, y) {
      const f = String(fid || "").trim()
      const pl = demo.placements[f] || { tag: "", page: 0, x: 0, y: 0, font_size: 14 }
      pl.x = Number(x || 0)
      pl.y = Number(y || 0)
      demo.placements[f] = pl
      return { ok: true }
    },
    async update_placement(fid, patch) {
      const f = String(fid || "").trim()
      const pl = demo.placements[f]
      if (!pl) return { ok: false, error: "not_found" }
      const p = patch && typeof patch === "object" ? patch : {}
      Object.assign(pl, p)
      demo.placements[f] = pl
      return { ok: true }
    },
    async delete_tags(tags) {
      const arr = Array.isArray(tags) ? tags.map((x) => String(x).trim()).filter(Boolean) : []
      for (const t of arr) {
        demo.tags = demo.tags.filter((k) => k !== t)
        delete demo.values[t]
        for (const [fid, pl] of Object.entries(demo.placements)) {
          if (pl && typeof pl === "object" && String(pl.tag || "").trim() === t) delete demo.placements[fid]
        }
      }
      return { ok: true }
    },
    async delete_elements(fids) {
      const arr = Array.isArray(fids) ? fids.map((x) => String(x).trim()).filter(Boolean) : []
      for (const fid of arr) delete demo.placements[fid]
      return { ok: true }
    },
    async set_project_payload(payload) {
      const p = payload && typeof payload === "object" ? payload : {}
      demo.tags = Array.isArray(p.tags) ? p.tags.map(String) : demo.tags
      demo.values = p.values && typeof p.values === "object" ? { ...p.values } : demo.values
      demo.placements = p.placements && typeof p.placements === "object" ? { ...p.placements } : demo.placements
      return { ok: true }
    },
    async get_element_info(fid) {
      const f = String(fid || "").trim()
      const pl = demo.placements[f]
      if (!pl) return { ok: false, error: "not_found" }
      return {
        ok: true,
        page: Number(pl.page || 0),
        tag: String(pl.tag || ""),
        x: Number(pl.x || 0),
        y: Number(pl.y || 0),
        font_size: Number(pl.font_size || 14),
        page_display_width: 1240,
        page_display_height: 1754,
      }
    },
    async get_preview_png_base64_page(page_index) {
      const idx = Math.max(0, Math.min(demo.pageCount - 1, Number(page_index || 0)))
      return {
        ok: true,
        png: makeSvgDataUrl(idx),
        page_display_width: 1240,
        page_display_height: 1754,
        page_index: idx,
      }
    },
    async get_preview_png_base64(tagOrFid) {
      const q = String(tagOrFid || "").trim()
      const pl = demo.placements[q]
      if (pl) return api.get_preview_png_base64_page(Number(pl.page || 0))
      for (const [, p] of Object.entries(demo.placements)) {
        if (p && typeof p === "object" && String(p.tag || "").trim() === q) return api.get_preview_png_base64_page(Number(p.page || 0))
      }
      return api.get_preview_png_base64_page(0)
    },
  }

  window.pywebview = { api }
})()

// いろんなママさんペルソナで“最大公約数”に寄せた設計（後述）:
// - 夜に作業する（暗めでも目が疲れない）
// - 片手でも押せる（大きいタップ領域、下に主要アクション）
// - “作業感”を減らす（柔らかい色、手ごたえ、気分が上がる演出）
// - 迷わない（次へだけで進む、今どこか常に見える）

const state = {
  projectPath: null,
  projectName: null,
  workers: [],
  workerId: null,
  appStage: "gate", // "gate" | "main"
  gate: {
    step: "choose", // "choose" | "worker" | "admin"
    password: "",
    error: "",
  },
  tags: [],
  idx: 0,
  values: {},
  placements: {},
  selectKeys: [],
  clipboard: null,
  undoStack: [],
  redoStack: [],
  working: false,
  inPrivate: false,
  timerStart: null,
  privateTotal: 0,
  dropDir: "",
  lastPreviewKey: null,
  justCompleted: false,
  designMode: false,
  designKey: null,
  pageW: 600,
  pageH: 800,
  designPos: null,
  uiMode: "worker", // "admin" | "worker"
  addMode: false,
  addDraftName: "",
  previewPageIndex: 0,
  history: [],
  lastSession: null,
  sessionStart: null,
  lastProjectDir: null,
  pageCount: 1,
  // タグ欄は常時表示（右上のON/OFFボタンは廃止）
  showTagPane: true,
  showPanel: true,
  pageLocked: false,
  lastFilledPdf: null,
  lastReportPdf: null,
  lastExportDir: null,
  defaultFontSize: 14,
  viewZoom: 1.0,
  viewBaseZoom: 1.0,
  viewPanX: 0,
  viewPanY: 0,
  locale: getLocaleSafe(),
  adLastShown: {},
  adSessionCounts: {},
  showPreviewHint: true,
  placePaletteOpen: false,
}

state.history = loadLocal("inputstudio-history", [])
state.lastSession = loadLocal("inputstudio-last-session", null)
state.lastProjectDir = loadLocal("inputstudio-last-dir", null)
state.showPanel = loadLocal("inputstudio-show-panel", true)
state.adLastShown = loadLocal("inputstudio-ad-last-shown", {})

function loadLocal(key, fallback) {
  try {
    const raw = localStorage.getItem(key)
    if (!raw) return fallback
    return JSON.parse(raw)
  } catch {
    return fallback
  }
}

function saveLocal(key, val) {
  try {
    localStorage.setItem(key, JSON.stringify(val))
  } catch {}
}

function getRenderedContentRect(imgEl, pageW, pageH) {
  // imgEl is sized to its container, with object-fit: contain.
  // We need the actual rendered content box to avoid coordinate drift.
  const r = imgEl.getBoundingClientRect()
  const cw = Math.max(1, r.width)
  const ch = Math.max(1, r.height)
  // Prefer actual intrinsic image aspect; fall back to logical page size.
  const iw = Math.max(1, Number(imgEl.naturalWidth || pageW || 1))
  const ih = Math.max(1, Number(imgEl.naturalHeight || pageH || 1))
  const s = Math.min(cw / iw, ch / ih)
  const dw = iw * s
  const dh = ih * s
  const dx = (cw - dw) / 2
  const dy = (ch - dh) / 2
  return {
    left: r.left + dx,
    top: r.top + dy,
    width: dw,
    height: dh,
  }
}

function clampNum(v, minV, maxV) {
  const n = Number(v)
  if (!Number.isFinite(n)) return minV
  return Math.max(minV, Math.min(maxV, n))
}

function applyPreviewTransform() {
  const sc = $("#previewScale")
  if (!sc) return
  const userZoom = clampNum(state.viewZoom || 1, 0.5, 3.0)
  const baseZoom = clampNum(state.viewBaseZoom || 1, 0.6, 1.0)
  const z = clampNum(userZoom * baseZoom, 0.4, 3.0)
  const host = $("#previewImg")
  if (host) {
    const r = host.getBoundingClientRect()
    const w = Math.max(1, Number(r.width || 0))
    const h = Math.max(1, Number(r.height || 0))
    if (z <= 1.001) {
      state.viewPanX = 0
      state.viewPanY = 0
    } else {
      const maxX = Math.max(0, ((w * z) - w) / 2)
      const maxY = Math.max(0, ((h * z) - h) / 2)
      state.viewPanX = clampNum(state.viewPanX || 0, -maxX, maxX)
      state.viewPanY = clampNum(state.viewPanY || 0, -maxY, maxY)
    }
  }
  const tx = Number(state.viewPanX || 0) || 0
  const ty = Number(state.viewPanY || 0) || 0
  sc.style.transform = `translate(${tx}px, ${ty}px) scale(${z})`
  const zi = $("#zoomIndicator")
  if (zi) zi.textContent = `${Math.round(userZoom * 100)}%`
}

function resetPreviewViewport({ zoom = 1.0 } = {}) {
  state.viewZoom = clampNum(zoom, 0.5, 3.0)
  state.viewBaseZoom = 1.0
  state.viewPanX = 0
  state.viewPanY = 0
}

function updatePreviewBaseZoom() {
  // Keep base zoom neutral; object-fit: contain already handles viewport fitting.
  state.viewBaseZoom = 1.0
}

function normalizeViewportAtFit() {
  const z = clampNum((Number(state.viewZoom || 1) || 1) * (Number(state.viewBaseZoom || 1) || 1), 0.4, 3.0)
  if (z <= 1.001) {
    state.viewPanX = 0
    state.viewPanY = 0
  }
}

function updatePreviewGuideHint() {
  const hint = $("#previewGuideHint")
  if (!hint) return
  const previewImg = $("#previewImg")
  if (!previewImg) {
    hint.style.display = "none"
    return
  }
  const paletteMode = !!state.placePaletteOpen
  const title = paletteMode
    ? "タグ名と値を入力して配置しよう"
    : "まずはPDFに欄（タグ）を置きましょう"
  const text = paletteMode
    ? "タグ名と値を入力して配置しよう。タグ一覧のタグをクリックすると既存タグを呼び出せます。同じタグはまとめて値を編集できます。"
    : "PDF上をダブルクリックしてタグ名と値を入力し、欄を配置できます。"
  hint.className = `emptyHint ${paletteMode ? "emptyHint--palette" : ""}`.trim()
  hint.innerHTML = `
    <div class="emptyHint__title">${escapeHtml(title)}</div>
    <div class="emptyHint__text">${escapeHtml(text)}</div>
    <div class="emptyHint__actions">
      ${paletteMode ? "" : `<button class="btn btn--primary" id="btnAddFromCenter">中央に欄を追加</button>`}
    </div>
  `
  hint.style.display = "block"

  const btnAddFromCenter = $("#btnAddFromCenter")
  if (btnAddFromCenter) btnAddFromCenter.onclick = () => {
    const x = Math.round(0.5 * state.pageW)
    const y = Math.round(0.5 * state.pageH)
    if (!x || !y) return
    openPlacePalette({ x, y })
  }
}

async function setViewZoom(nextZoom, { persist = true } = {}) {
  state.viewZoom = clampNum(nextZoom, 0.5, 3.0)
  normalizeViewportAtFit()
  applyPreviewTransform()
  drawOverlay()
  if (!persist) return
  try {
    await window.pywebview?.api?.update_admin_settings?.({ view_zoom: state.viewZoom })
  } catch {}
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj))
}

function snapshotProject() {
  return {
    tags: [...state.tags],
    values: deepClone(state.values || {}),
    placements: deepClone(state.placements || {}),
  }
}

async function applyProjectSnapshot(snap, { save = true } = {}) {
  state.tags = Array.isArray(snap?.tags) ? [...snap.tags] : []
  state.values = snap?.values && typeof snap.values === "object" ? deepClone(snap.values) : {}
  state.placements = snap?.placements && typeof snap.placements === "object" ? deepClone(snap.placements) : {}
  state.idx = Math.max(0, Math.min(state.idx, state.tags.length - 1))
  state.selectKeys = state.selectKeys.filter((fid) => state.placements?.[fid])
  if (save && window.pywebview?.api?.set_project_payload) {
    await window.pywebview.api.set_project_payload({ tags: state.tags, values: state.values, placements: state.placements })
    await window.pywebview.api.save_current_project?.(false)
  }
  render()
}

function pushUndo(beforeSnap) {
  state.undoStack.push(beforeSnap)
  if (state.undoStack.length > 60) state.undoStack.shift()
  state.redoStack = []
}

function isTextEditingTarget(el) {
  const t = (el?.tagName || "").toLowerCase()
  if (t === "textarea") return true
  if (t === "input") return true
  if (el?.isContentEditable) return true
  return false
}

function uniqueTag(base) {
  const clean = String(base || "").trim() || "tag"
  if (!state.tags.includes(clean)) return clean
  for (let i = 2; i < 9999; i++) {
    const t = `${clean}_${i}`
    if (!state.tags.includes(t)) return t
  }
  return `${clean}_${Date.now()}`
}

function isWideChar(ch) {
  const c = (ch || "").charCodeAt(0) || 0
  // Hiragana/Katakana/CJK, fullwidth forms, punctuation blocks
  if ((c >= 0x3040 && c <= 0x30ff) || (c >= 0x3400 && c <= 0x9fff) || (c >= 0xff00 && c <= 0xffef) || (c >= 0x3000 && c <= 0x303f))
    return true
  return false
}

function placementBoxOnPage(fid, pl) {
  const fs = Number(pl?.font_size || 14) || 14
  const lh = Number(pl?.line_height || 1.2) || 1.2
  const ls = Number(pl?.letter_spacing ?? DEFAULT_LETTER_SPACING) || DEFAULT_LETTER_SPACING
  const writingMode = String(pl?.writing_mode || "horizontal")
  const tag = String(pl?.tag || "").trim()
  const v = String((state.values?.[tag] || "")).replaceAll("<br>", "\n")
  const lines = (v ? v.split("\n") : []).filter((s) => s != null)
  const drawLines = lines.length ? lines : [tag || fid]

  const padX = Math.max(8, fs * 0.35)
  const padY = Math.max(6, fs * 0.25)

  let wPage = 42
  let hPage = 22
  if (writingMode === "vertical") {
    // splitlines => columns (right-to-left)
    const cols = drawLines
    const maxChars = Math.max(1, ...cols.map((s) => String(s || "").length))
    const colCount = Math.max(1, cols.length)
    const stepY = fs * lh + ls
    const stepX = fs * 1.10 + ls
    hPage = Math.max(22, maxChars * stepY + padY * 2)
    wPage = Math.max(32, colCount * stepX + padX * 2)
  } else {
    let maxW = 0
    for (const line of drawLines) {
      const s = String(line || "")
      let w = 0
      for (const ch of s) {
        w += (isWideChar(ch) ? fs * 1.0 : fs * 0.62)
      }
      if (s.length > 1) w += (s.length - 1) * ls
      maxW = Math.max(maxW, w)
    }
    wPage = Math.max(42, maxW + padX * 2)
    hPage = Math.max(22, drawLines.length * fs * lh + padY * 2)
  }

  const x = Number(pl?.x || 0)
  const y = Number(pl?.y || 0)
  return { x, y, w: wPage, h: hPage, padX, padY }
}

function toast(msg) {
  const el = $("#toast")
  el.textContent = msg
  el.style.display = "block"
  clearTimeout(el._t)
  el._t = setTimeout(() => (el.style.display = "none"), 2100)
}

function ensureSysDialogRoot() {
  let root = document.getElementById("sysDialogRoot")
  if (!root) {
    root = document.createElement("div")
    root.id = "sysDialogRoot"
    root.className = "sysDialog"
    root.style.display = "none"
    document.body.appendChild(root)
  }
  return root
}

async function openSysDialog({ title, message, type = "alert", defaultValue = "" }) {
  const root = ensureSysDialogRoot()
  return await new Promise((resolve) => {
    const close = (value) => {
      root.style.display = "none"
      root.innerHTML = ""
      resolve(value)
    }
    root.style.display = "block"
    root.innerHTML = `
      <div class="sysDialog__backdrop" id="sysDialogBackdrop"></div>
      <div class="sysDialog__card">
        <div class="sysDialog__title">${escapeHtml(String(title || tr("dialog.title", "確認")))}</div>
        <div class="sysDialog__body">${escapeHtml(String(message || ""))}</div>
        ${type === "prompt" ? `<input class="input" id="sysDialogInput" value="${escapeHtml(String(defaultValue || ""))}" />` : ""}
        <div class="row" style="justify-content:flex-end; margin-top:12px">
          ${type !== "alert" ? `<button class="btn btn--soft" id="sysDialogCancel">${escapeHtml(tr("dialog.cancel", "キャンセル"))}</button>` : ""}
          <button class="btn btn--primary" id="sysDialogOk">${escapeHtml(type === "alert" ? tr("dialog.ok", "OK") : tr("dialog.continue", "続行"))}</button>
        </div>
      </div>
    `
    const input = document.getElementById("sysDialogInput")
    if (input) {
      input.focus()
      try {
        const len = String(input.value || "").length
        input.setSelectionRange(len, len)
      } catch {}
      input.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter") {
          ev.preventDefault()
          close(String(input.value || ""))
        }
      })
    }
    const ok = document.getElementById("sysDialogOk")
    if (ok) ok.onclick = () => close(type === "prompt" ? String(input?.value || "") : true)
    const cancel = document.getElementById("sysDialogCancel")
    if (cancel) cancel.onclick = () => close(type === "prompt" ? null : false)
    const backdrop = document.getElementById("sysDialogBackdrop")
    if (backdrop) backdrop.onclick = () => close(type === "prompt" ? null : false)
  })
}

const uiAlert = async (msg, title = null) => openSysDialog({ title, message: msg, type: "alert" })
const uiConfirm = async (msg, title = null) => openSysDialog({ title, message: msg, type: "confirm" })
const uiPrompt = async (msg, defaultValue = "", title = null) => openSysDialog({ title, message: msg, type: "prompt", defaultValue })

function apiErrorMessage(result, fallback = "エラーが発生しました") {
  const code = String(result?.code || "").toUpperCase()
  const map = {
    RATE_LIMITED: tr("error.rateLimited", "アクセスが集中しています。少し待って再試行してください。"),
    UPLOAD_TOO_LARGE: tr("error.uploadTooLarge", "アップロードファイルが大きすぎます。サイズを下げて再試行してください。"),
    SERVER_BUSY: tr("error.serverBusy", "サーバーが混み合っています。少し待って再試行してください。"),
    INVALID_ORDER: tr("error.invalidOrder", "並び順データが不正です。並べ替えをやり直してください。"),
    METHOD_NOT_ALLOWED: tr("error.methodNotAllowed", "サーバー接続エラーです。サーバー再起動後に再試行してください。"),
  }
  if (code && map[code]) return map[code]
  return String(result?.error || fallback)
}

function shouldShowUnlockAd(action) {
  if (!isWeb()) return false
  const rule = AD_UNLOCK_RULES[action]
  if (!rule) return false
  const now = Date.now()
  const last = Number(state.adLastShown?.[action] || 0)
  const count = Number(state.adSessionCounts?.[action] || 0)
  if (count >= Number(rule.maxPerSession || 0)) return false
  return now - last >= Number(rule.cooldownMs || 0)
}

function recordUnlockAdShown(action) {
  const now = Date.now()
  state.adLastShown = state.adLastShown || {}
  state.adSessionCounts = state.adSessionCounts || {}
  state.adLastShown[action] = now
  state.adSessionCounts[action] = Number(state.adSessionCounts[action] || 0) + 1
  saveLocal("inputstudio-ad-last-shown", state.adLastShown)
}

function unlockHintBubble(action) {
  if (!isWeb()) return ""
  if (!AD_UNLOCK_RULES[action]) return ""
  return ` <span class="adHintBubble">${escapeHtml(tr("ad.unlock.badge", "広告を見て解放"))}</span>`
}

async function showUnlockAd(action) {
  if (!shouldShowUnlockAd(action)) return true
  const modal = $("#modal")
  if (!modal) return true
  const cfg = getAdConfig()
  const unlockSlotId = getAdSlotFor("unlock") || getAdSlotFor("gate")
  const showLiveUnlockAd = cfg.enabled && cfg.provider === "adsense" && !!unlockSlotId
  return await new Promise((resolve) => {
    let sec = Math.max(1, Number(cfg.unlock?.minSeconds || 3) || 3)
    const close = (ok) => {
      modal.style.display = "none"
      modal.innerHTML = ""
      resolve(!!ok)
    }
    modal.style.display = "block"
    modal.innerHTML = `
      <div class="modal__backdrop" id="adUnlockClose"></div>
      <div class="modal__card" style="max-width:420px; width:min(92vw,420px)">
        <div class="modal__title">${escapeHtml(tr("ad.unlock.title", "広告を表示して続行"))}</div>
        <div class="label" style="line-height:1.7">${escapeHtml(tr("ad.unlock.desc", "無料提供を継続するため、短い広告表示後にこの操作を実行できます。"))}</div>
        ${showLiveUnlockAd ? `<div class="adUnlockLive" id="adUnlockLive"></div>` : `<div class="adUnlockMock">AD</div>`}
        <div class="label" id="adUnlockTimer">${escapeHtml(tr("ad.unlock.wait", `${sec}秒後に続行できます`, { sec }))}</div>
        <div class="row" style="justify-content:flex-end; margin-top:12px">
          <button class="btn btn--soft" id="adUnlockCancel">${escapeHtml(tr("ad.unlock.cancel", "キャンセル"))}</button>
          <button class="btn btn--primary" id="adUnlockGo" disabled>${escapeHtml(tr("ad.unlock.continue", "広告を見て続行"))}</button>
        </div>
      </div>
    `
    if (showLiveUnlockAd) {
      ensureAdSenseScript().then((ok) => {
        if (!ok) return
        const live = document.getElementById("adUnlockLive")
        if (!live) return
        mountAdSenseInto(live, unlockSlotId)
      })
    }
    const btnGo = $("#adUnlockGo")
    const timerEl = $("#adUnlockTimer")
    const timer = setInterval(() => {
      sec -= 1
      if (sec <= 0) {
        clearInterval(timer)
        if (btnGo) btnGo.disabled = false
        if (timerEl) timerEl.textContent = tr("ad.unlock.ready", "続行できます")
        return
      }
      if (timerEl) timerEl.textContent = tr("ad.unlock.wait", `${sec}秒後に続行できます`, { sec })
    }, 1000)
    $("#adUnlockClose").onclick = () => {
      clearInterval(timer)
      close(false)
    }
    $("#adUnlockCancel").onclick = () => {
      clearInterval(timer)
      close(false)
    }
    if (btnGo) {
      btnGo.onclick = () => {
        clearInterval(timer)
        recordUnlockAdShown(action)
        close(true)
      }
    }
  })
}

function fmtTime(sec) {
  const s = Math.max(0, Math.floor(sec))
  const h = String(Math.floor(s / 3600)).padStart(2, "0")
  const m = String(Math.floor((s % 3600) / 60)).padStart(2, "0")
  // 要望: タイマーは「時:分」だけ表示（秒は不要）
  return `${h}:${m}`
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;")
}

function tipIcon(n, text) {
  return `<span class="tipIcon" data-tip="${escapeHtml(text)}">${n}</span>`
}

// Tooltip that never goes off-screen (replaces CSS-only tooltip).
let _tipFloatBound = false
let _previewFitBound = false
function bindTipFloatOnce() {
  if (_tipFloatBound) return
  _tipFloatBound = true

  const ensureEl = () => {
    let el = document.getElementById("tipFloat")
    if (!el) {
      el = document.createElement("div")
      el.id = "tipFloat"
      el.className = "tipFloat"
      el.style.display = "none"
      document.body.appendChild(el)
    }
    return el
  }

  let active = null
  const hide = () => {
    const el = document.getElementById("tipFloat")
    if (el) el.style.display = "none"
    active = null
  }
  const showFor = (target) => {
    const tip = target?.getAttribute?.("data-tip")
    if (!tip) return
    active = target
    const el = ensureEl()
    el.textContent = tip
    el.style.display = "block"

    const r = target.getBoundingClientRect()
    const br = el.getBoundingClientRect()
    const pad = 10
    const clamp = (v, a, b) => Math.min(Math.max(v, a), b)
    let left = r.left + r.width / 2 - br.width / 2
    left = clamp(left, pad, window.innerWidth - br.width - pad)
    let top = r.top - br.height - 10
    if (top < pad) top = r.bottom + 10
    top = clamp(top, pad, window.innerHeight - br.height - pad)
    el.style.left = `${Math.round(left)}px`
    el.style.top = `${Math.round(top)}px`
  }

  const findIcon = (ev) => ev?.target?.closest?.(".tipIcon")
  document.addEventListener("pointerover", (ev) => {
    const icon = findIcon(ev)
    if (!icon) return
    showFor(icon)
  })
  document.addEventListener("pointerout", (ev) => {
    const icon = findIcon(ev)
    if (!icon) return
    const rel = ev.relatedTarget
    if (rel && icon.contains(rel)) return
    hide()
  })
  document.addEventListener("focusin", (ev) => {
    const icon = findIcon(ev)
    if (!icon) return
    showFor(icon)
  })
  document.addEventListener("focusout", (ev) => {
    const icon = findIcon(ev)
    if (!icon) return
    hide()
  })
  window.addEventListener(
    "scroll",
    () => {
      if (!active) return
      showFor(active)
    },
    true
  )
  window.addEventListener("resize", () => {
    if (!active) return
    showFor(active)
  })
  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") hide()
  })
}

function bindPreviewFitOnce() {
  if (_previewFitBound) return
  _previewFitBound = true
  window.addEventListener("resize", () => {
    if (!state.projectPath) return
    updatePreviewBaseZoom()
    normalizeViewportAtFit()
    applyPreviewTransform()
    drawOverlay()
  })
}

async function showPage(pageIndex) {
  if (!state.projectPath) return
  const api = window.pywebview?.api
  if (!api || typeof api.get_preview_png_base64_page !== "function") return
  const idx = Math.max(0, Math.min((state.pageCount || 1) - 1, Number(pageIndex) || 0))
  state.previewPageIndex = idx
  state.pageLocked = true
  const my = ++pageReq
  const p0 = $("#pageIndicator")
  if (p0) p0.textContent = `${idx + 1} / ${state.pageCount || 1} …`
  let r = await api.get_preview_png_base64_page(idx)
  // Auto-recover when backend lost the project (WebView reload / timing / cache issues).
  if (r && !r.ok && r.error === "no_project" && state.projectPath && typeof api.load_project === "function") {
    try {
      await api.load_project(state.projectPath)
      r = await api.get_preview_png_base64_page(idx)
    } catch {}
  }
  if (my !== pageReq) return
  if (r && r.ok) {
    const img = $("#previewImg")
    if (img) {
      img.onload = () => {
        img.style.visibility = "visible"
        updatePreviewBaseZoom()
        normalizeViewportAtFit()
        applyPreviewTransform()
        drawOverlay()
      }
      img.onerror = () => {
        img.style.visibility = "hidden"
        toast("プレビュー画像の読み込みに失敗しました（パス/権限/文字コードの可能性）")
      }
      img.style.visibility = "hidden"
      img.src = r.png_data || r.png
    }
    // Align coordinate system to actual rendered image (rotation/aspect-safe)
    if (img && img.naturalWidth && img.naturalHeight) {
      state.pageW = img.naturalWidth
      state.pageH = img.naturalHeight
    } else {
      state.pageW = r.page_display_width || state.pageW
      state.pageH = r.page_display_height || state.pageH
      updatePreviewBaseZoom()
      normalizeViewportAtFit()
    }
    drawOverlay()
    const p = $("#pageIndicator")
    if (p) p.textContent = `${idx + 1} / ${state.pageCount || 1}`
  } else {
    const img = $("#previewImg")
    if (img) {
      img.src = ""
      img.style.visibility = "hidden"
    }
    toast(`ページ表示に失敗: ${r?.error || "unknown"}`)
  }
}

function calcNetSeconds() {
  if (!state.timerStart) return 0
  const now = Date.now()
  const base = (now - state.timerStart) / 1000
  return Math.max(0, base - state.privateTotal)
}

function filledCount() {
  let n = 0
  for (const k of state.tags) {
    const v = (state.values[k] || "").replaceAll("<br>", "").trim()
    if (v) n++
  }
  return n
}

async function updateTagValue(tag, rawText) {
  const raw = (rawText || "").replaceAll("\r\n", "\n")
  const val = raw.replaceAll("\n", "<br>")
  state.values[tag] = val
  await window.pywebview.api.set_value(tag, val)
  queuePreview(tag)
}

// NOTE: 右側（または下段）に常時表示するタグ一覧は廃止。
// タグ操作は「パレット上のタグ一覧」に一本化する。

function renderGate() {
  const err = String(state.gate?.error || "")
  const localeOptionsHtml = LOCALE_OPTIONS.map((opt) => {
    const sel = state.locale === opt.code ? "selected" : ""
    return `<option value="${escapeHtml(opt.code)}" ${sel}>${escapeHtml(opt.label)}</option>`
  }).join("")
  const localeFlag = getLocaleMeta(state.locale).flag

  $("#app").innerHTML = `
    <div class="bgBlobs" aria-hidden="true">
      <div class="blob b1"></div>
      <div class="blob b2"></div>
      <div class="blob b3"></div>
    </div>
    <div class="gate">
      <div class="gateCard">
        <div class="gateBrand">
          <div class="logo gateLogo" aria-hidden="true"></div>
          <div class="gateTitle">
            <div class="gateTitle__top">PDF Input Studio</div>
            <div class="gateTitle__sub">${escapeHtml(tr("brand.tagline", "PDFに文字を置いて、完成PDFを作る"))}</div>
          </div>
        </div>
        <div class="row gateLocaleRow" style="justify-content:flex-end; margin-top:8px">
          <label class="label gateLocaleLabel" for="gateLocale">${escapeHtml(tr("top.languageMixed", "言語 Language"))}</label>
          <span id="gateLocaleFlag" class="flagIcon flagIcon--${escapeHtml(localeFlag)}" aria-hidden="true"></span>
          <select id="gateLocale" class="input" style="width:220px; padding:8px 10px">
            ${localeOptionsHtml}
          </select>
        </div>

        <div class="gateActions">
          <button class="btn btn--primary" id="gateLoadPdf">${escapeHtml(tr("gate.loadPdf", "PDFを読み込む"))}</button>
          <button class="btn btn--soft" id="gateLoadProject">${escapeHtml(tr("gate.loadZip", "プロジェクトZIPを開く"))}${unlockHintBubble("zip_open")}</button>
        </div>
        <div class="label gateHint">${escapeHtml(tr("gate.hint", "PDFから新規作成　／　既存の案件（ZIP・PDF同梱）を開く"))}</div>
        <div class="gateGuide">
          <div class="gateGuide__title">${escapeHtml(tr("top.value.title", "このサイトでできること"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.value.audience", "対象: 申請書・帳票の入力担当者"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.value.benefit1", "価値1: タグ同期で同じ項目を一括更新"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.value.benefit2", "価値2: ZIPで案件を持ち運びしやすい"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.value.benefit3", "価値3: 入力からPDF出力までブラウザで完結"))}</div>
          <div class="gateGuide__title" style="margin-top:6px">${escapeHtml(tr("top.howto.title", "使い方（3ステップ）"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.howto.step1", "1. プロジェクトZIPを開く（またはPDFから新規作成）"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.howto.step2", "2. 項目を入力して必要に応じてページ操作"))}</div>
          <div class="gateGuide__item">${escapeHtml(tr("top.howto.step3", "3. PDFダウンロード / プロジェクト保存"))}</div>
        </div>
        <div class="adSlot adSlot--gate" id="adSlotGate">
          <div class="adSlot__label">${escapeHtml(tr("ad.label", "広告"))}</div>
          <div class="adSlot__title">${escapeHtml(tr("ad.sponsorTitle", "スポンサーからのお知らせ"))}</div>
          <div class="adSlot__body">${escapeHtml(tr("ad.sponsorBody", "ここにバナー広告が表示されます（実装準備中）"))}</div>
          <div class="adSlot__live" aria-label="ad slot gate"></div>
        </div>
        <div class="gateTrustNav gateTrustNav--footer" aria-label="site trust navigation">
          <a class="gateTrustNav__link" href="/global-search.html">${escapeHtml(tr("top.nav.global", "多言語検索ガイド"))}</a>
          <a class="gateTrustNav__link" href="/solutions.html">${escapeHtml(tr("top.nav.guide", "活用ガイド"))}</a>
          <a class="gateTrustNav__link" href="/application-form-filling.html">${escapeHtml(tr("top.nav.forms", "申請書/様式入力"))}</a>
          <a class="gateTrustNav__link" href="/pdf-merge-split.html">${escapeHtml(tr("top.nav.tools", "PDF結合/分割"))}</a>
          <a class="gateTrustNav__link" href="/about.html">${escapeHtml(tr("top.nav.about", "企業情報"))}</a>
          <a class="gateTrustNav__link" href="/contact.html">${escapeHtml(tr("top.nav.contact", "お問い合わせ"))}</a>
          <a class="gateTrustNav__link" href="/privacy.html">${escapeHtml(tr("top.nav.privacy", "プライバシーポリシー"))}</a>
          <a class="gateTrustNav__link" href="/terms.html">${escapeHtml(tr("top.nav.terms", "利用規約"))}</a>
          <a class="gateTrustNav__link" href="/faq.html">${escapeHtml(tr("top.nav.faq", "FAQ"))}</a>
        </div>
        ${err ? `<div class="gateError">${escapeHtml(err)}</div>` : ""}
        ${window.__INPUTSTUDIO_DEMO__ ? `<input type="file" id="gateDemoPdf" accept=".pdf,application/pdf" style="display:none" />` : ""}
      </div>
    </div>
    <div class="toast" id="toast"></div>
    <div class="modal" id="modal" style="display:none"></div>
  `
  refreshAdSlots()

  // bindings: PDFを読み込む
  const gateLocale = $("#gateLocale")
  if (gateLocale) {
    gateLocale.onchange = () => {
      const next = String(gateLocale.value || "ja")
      state.locale = window.i18n?.setLocale?.(next) || next
      syncLocaleQuery(state.locale)
      renderGate()
    }
  }

  const bLoadPdf = $("#gateLoadPdf")
  if (bLoadPdf) bLoadPdf.onclick = async () => {
    if (window.__INPUTSTUDIO_DEMO__) {
      const inp = $("#gateDemoPdf")
      if (inp) inp.click()
      return
    }
    try {
      const api = window.pywebview?.api
      const pick = api?.pick_pdf
      const createSimple = api?.create_project_from_pdf_simple
      if (!pick || !createSimple) {
        await uiAlert("PDFから新規作成する機能が見つかりません。最新版のアプリをご利用ください。")
        return
      }
      const r = await pick()
      if (!r?.ok) return
      toast(tr("gate.toastCreateProjectFromPdf", "PDFを読み込み、新規プロジェクトを作成します…"))
      const g = await createSimple(r.path)
      if (!g?.ok || !g.path) {
        await uiAlert((g?.errors || ["PDFをプロジェクト化できませんでした"]).join("\n"))
        return
      }
      const loaded = await api.load_project(g.path)
      if (!loaded?.ok) {
        await uiAlert("新規プロジェクトを開けませんでした")
        return
      }
      state.projectPath = g.path
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.placements = loaded.placements || {}
      state.pageCount = loaded.page_count || 1
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      state.uiMode = loaded.ui_mode || state.uiMode
      state.lastSession = { path: g.path, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      try {
        const dir = g.path.replace(/[/\\][^/\\]+$/, "")
        state.lastProjectDir = dir
        saveLocal("inputstudio-last-dir", dir)
      } catch {}
      state.working = false
      state.inPrivate = false
      state.timerStart = null
      state.privateTotal = 0
      state.appStage = "main"
      state.showPreviewHint = true
      state.placePaletteOpen = false
      resetPreviewViewport({ zoom: 1.0 })
      const started = await autoStartWorkIfPossible()
      toast(started ? "新規案件を作成し、作業タイマーを開始しました" : "PDFから新規プロジェクトを作成しました。必要に応じてタグを配置してください。")
      render()
      await queuePreview()
    } catch (e) {
      await uiAlert(`PDFから新規作成に失敗しました: ${e}`)
    }
  }

  // bindings: ZIPを読み込む（プロジェクト＝ZIP前提）
  const bLoadProject = $("#gateLoadProject")
  if (bLoadProject) bLoadProject.onclick = async () => {
    try {
      const okToProceed = await showUnlockAd("zip_open")
      if (!okToProceed) return
      const r = await window.pywebview.api.pick_project(
        window.__INPUTSTUDIO_WEB__ ? { zipOnly: true } : undefined
      )
      if (!r.ok) {
        if (r.error) toast(apiErrorMessage(r, r.error))
        return
      }
      toast(tr("gate.toastLoadingZip", "ZIPを読み込み中…"))
      try {
        const dir = r.path?.replace(/[/\\][^/\\]+$/, "")
        if (dir) {
          state.lastProjectDir = dir
          saveLocal("inputstudio-last-dir", dir)
        }
      } catch {}
      const loaded = await window.pywebview.api.load_project(r.path)
      if (!loaded.ok) return
      state.projectPath = r.path
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.placements = loaded.placements || {}
      state.pageCount = loaded.page_count || 1
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      state.uiMode = loaded.ui_mode || state.uiMode
      state.lastSession = { path: r.path, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      state.working = false
      state.inPrivate = false
      state.timerStart = null
      state.privateTotal = 0
      state.appStage = "main"
      state.showPreviewHint = true
      state.placePaletteOpen = false
      resetPreviewViewport({ zoom: 1.0 })
      const started = await autoStartWorkIfPossible()
      toast(started ? tr("gate.toastLoadedZipAndTimer", "ZIPを読み込み、作業タイマーを開始しました") : tr("gate.toastLoadedZip", "ZIPを読み込みました"))
      render()
      await queuePreview()
    } catch (e) {
      await uiAlert(`プロジェクトの読み込みに失敗しました: ${e}`)
    }
  }

  // Demo: ファイル選択後に読み込み
  const gateDemoPdf = $("#gateDemoPdf")
  if (gateDemoPdf) {
    gateDemoPdf.onchange = async () => {
      const file = gateDemoPdf.files?.[0]
      if (!file) return
      try {
        toast("PDFを読み込み中…")
        const buf = await file.arrayBuffer()
        const pdfjs = await import("https://cdn.jsdelivr.net/npm/pdfjs-dist@4.8.69/build/pdf.mjs")
        pdfjs.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.8.69/build/pdf.worker.mjs"
        const doc = await pdfjs.getDocument({ data: buf }).promise
        window.__demoPdfDoc = doc
        window.__demoPdfCache = new Map()
        const api = window.pywebview?.api
        if (api) {
          api.get_preview_png_base64_page = async (page_index) => {
            const idx = Math.max(0, Math.min(doc.numPages - 1, Number(page_index || 0)))
            const cache = window.__demoPdfCache
            if (cache.has(idx)) return cache.get(idx)
            const page = await doc.getPage(idx + 1)
            const scale = 150 / 72
            const vp = page.getViewport({ scale })
            const canvas = document.createElement("canvas")
            canvas.width = Math.floor(vp.width)
            canvas.height = Math.floor(vp.height)
            const ctx = canvas.getContext("2d")
            await page.render({ canvasContext: ctx, viewport: vp }).promise
            const blob = await new Promise((res) => canvas.toBlob(res, "image/png"))
            const url = URL.createObjectURL(blob)
            const out = { ok: true, png: url, page_display_width: canvas.width, page_display_height: canvas.height, page_index: idx }
            cache.set(idx, out)
            return out
          }
          api.get_preview_png_base64 = async (tag) => {
            const t = String(tag || "").trim()
            const pl = state.placements?.[t]
            const idx = pl ? Number(pl.page || 0) : 0
            return api.get_preview_png_base64_page(idx)
          }
        }
        state.projectPath = "demo:pdf"
        state.projectName = file.name
        state.pageCount = doc.numPages
        state.previewPageIndex = 0
        state.tags = []
        state.values = {}
        state.placements = {}
        state.appStage = "main"
        state.showPreviewHint = true
        state.placePaletteOpen = false
        toast(`PDFを読み込みました（${doc.numPages}ページ）`)
        render()
      } catch (e) {
        await uiAlert(`PDF読み込みに失敗しました: ${e}`)
      } finally {
        gateDemoPdf.value = ""
      }
    }
  }
}

function render() {
  if (state.appStage !== "main") {
    renderGate()
    return
  }
  const total = state.tags.length || 0
  const done = total ? filledCount() : 0
  const idx = state.idx
  const key = total ? state.tags[idx] : null
  const hasTags = !!key
  const val = key ? (state.values[key] || "") : ""
  const valText = val.replaceAll("<br>", "\n")
  const progress = total ? Math.round(((idx + 1) / total) * 100) : 0

  const isAdmin = state.uiMode === "admin"
  // 画面常設のタグ一覧は表示しない（パレットに統一）
  const showTagPane = false

  const left = `
    <div class="top">
      <div class="brand row spread" style="align-items:center">
        <div class="row" style="align-items:center; gap:12px">
          <div class="logo" aria-hidden="true"></div>
          <div class="brand__name">PDF Input Studio</div>
        </div>
        <button class="chip chip--soft" id="btnBackToGate">${escapeHtml(tr("main.backToTop", "トップページに戻る"))}</button>
      </div>
      ${window.__INPUTSTUDIO_DEMO__ ? `<input type="file" id="demoPdfFile" accept=".pdf,application/pdf" style="display:none" />` : ""}

      <div class="miniActions">
        <button class="chip chip--soft" id="btnOpen">${escapeHtml(tr("main.openZip", "プロジェクトZIPを開く"))}${unlockHintBubble("zip_open")}</button>
        ${isAdmin ? `<button class="chip chip--soft" id="btnOpenPdf">${escapeHtml(tr("main.newFromPdf", "PDFから新規"))}</button>` : ""}
        ${isAdmin ? `        <button class="chip" id="btnDesign">設計（統括）</button>` : ""}
        <button class="chip chip--soft" id="btnAppendPdf" ${state.projectPath ? "" : "disabled"}>${escapeHtml(tr("main.appendPdf", "PDF追加"))}${unlockHintBubble("pdf_append")}</button>
        <button class="chip chip--soft" id="btnReorderPdf" ${state.projectPath ? "" : "disabled"}>${escapeHtml(tr("main.reorderPdf", "PDF並べ替え"))}</button>
        <button class="chip chip--soft" id="btnCopyPageOp" ${state.projectPath ? "" : "disabled"}>${escapeHtml(tr("main.copyCurrentPage", "現在ページ複製"))}</button>
        <button class="chip chip--soft" id="btnDeletePageOp" ${state.projectPath ? "" : "disabled"}>${escapeHtml(tr("main.deleteCurrentPage", "現在ページ削除"))}</button>
        ${state.projectPath && !isWeb() ? `<button class="chip chip--soft" id="btnOpenSaved">保存先</button>` : ""}
        ${isAdmin ? `<button class="chip chip--soft" id="btnHistoryExport">履歴CSV</button>` : ""}
        ${isAdmin ? `<button class="chip chip--soft" id="btnHistoryReset">履歴リセット</button>` : ""}
      </div>
      <div class="adSlot adSlot--panel" id="adSlotPanel">
        <div class="adSlot__label">${escapeHtml(tr("ad.label", "広告"))}</div>
        <div class="adSlot__title">${escapeHtml(tr("ad.recommendTitle", "おすすめサービス"))}</div>
        <div class="adSlot__body">${escapeHtml(tr("ad.recommendBody", "ここに常時バナー広告が表示されます（実装準備中）"))}</div>
        <div class="adSlot__live" aria-label="ad slot panel"></div>
      </div>

    </div>

    ${
      hasTags
        ? `<div class="focus ${state.justCompleted ? "pop" : ""}">
            <div class="focus__head">
              <div class="focus__title">${escapeHtml(key)}</div>
              <div class="focus__meta">${total ? escapeHtml(tr("main.focusMeta", `${idx + 1}/${total} ・ Enterで次へ / Shift+Enterで改行`, { index: idx + 1, total })) : ""}</div>
            </div>

            <div class="focus__body">
              <textarea class="input textarea focus__input" id="val" placeholder="${escapeHtml(tr("main.inputPlaceholder", "ここに入力…"))}">${escapeHtml(valText)}</textarea>
              <div class="row spread" style="margin-top:10px">
                <div class="row">
                  <button class="btn btn--soft" id="btnPrev" ${idx <= 0 ? "disabled" : ""}>${escapeHtml(tr("main.prev", "戻る"))}</button>
                  <button class="btn btn--primary" id="btnNext" ${idx >= total - 1 ? "disabled" : ""}>${escapeHtml(tr("main.next", "次へ"))}</button>
                </div>
              </div>
              <div class="row" style="margin-top:14px; gap:8px">
                <button class="btn btn--soft" id="btnSave" ${state.projectPath ? "" : "disabled"} style="flex:1" title="${isWeb() ? escapeHtml(tr("main.savePdfTitle", "保存して完成PDFをダウンロードします。保存先を選択できます。")) : ""}">${isWeb() ? escapeHtml(tr("main.savePdf", "PDFダウンロード")) : escapeHtml(tr("main.saveOverwrite", "上書き保存"))}</button>
                <button class="btn btn--soft" id="btnSaveAs" ${state.projectPath ? "" : "disabled"} style="flex:1" title="${isWeb() ? escapeHtml(tr("main.saveProjectTitle", "プロジェクトをZIP（PDF同梱）で保存します。保存先を選択できます。")) : ""}">${isWeb() ? escapeHtml(tr("main.saveProject", "プロジェクトを保存")) : escapeHtml(tr("main.saveAs", "名前を付けて保存"))}</button>
              </div>
              <button class="btn btn--danger" id="btnFinish" style="width:100%; margin-top:8px; padding:12px 20px">${escapeHtml(tr("main.finish", "終了"))}</button>
            </div>
          </div>`
        : (state.projectPath ? `<div style="margin-top:12px">
            <div class="row" style="gap:8px; margin-bottom:8px">
              <button class="btn btn--soft" id="btnSave" style="flex:1" title="${isWeb() ? escapeHtml(tr("main.savePdfShortTitle", "保存して完成PDFをダウンロードします。")) : ""}">${isWeb() ? escapeHtml(tr("main.savePdf", "PDFダウンロード")) : escapeHtml(tr("main.saveOverwrite", "上書き保存"))}</button>
              <button class="btn btn--soft" id="btnSaveAs" style="flex:1" title="${isWeb() ? escapeHtml(tr("main.saveProjectShortTitle", "プロジェクトをZIP（PDF同梱）で保存します。")) : ""}">${isWeb() ? escapeHtml(tr("main.saveProject", "プロジェクトを保存")) : escapeHtml(tr("main.saveAs", "名前を付けて保存"))}</button>
            </div>
            <button class="btn btn--danger" id="btnFinish" style="width:100%; padding:12px 20px">${escapeHtml(tr("main.finish", "終了"))}</button>
          </div>` : "")
    }
    ${!isWeb() ? `<div class="glassBox" style="margin-top:10px">
      ${state.lastProjectDir ? `<div class="pathLine" title="${escapeHtml(state.lastProjectDir)}">前回開いたフォルダ: <span class="pathValue">${escapeHtml(state.lastProjectDir)}</span></div>` : ""}
      ${state.projectPath ? `<div class="pathLine" title="${escapeHtml(state.projectPath)}">保存先: <span class="pathValue">${escapeHtml(state.projectPath)}</span></div>` : ""}
      ${state.lastFilledPdf ? `<div class="pathLine" title="${escapeHtml(state.lastFilledPdf)}">提出PDF: <span class="pathValue">${escapeHtml(state.lastFilledPdf)}</span></div>` : ""}
    </div>` : ""}
    <div class="adSlot adSlot--panelBottom" id="adSlotPanelBottom" style="margin-top:auto">
      <div class="adSlot__label">${escapeHtml(tr("ad.label", "広告"))}</div>
      <div class="adSlot__title">${escapeHtml(tr("ad.recommendTitle", "おすすめサービス"))}</div>
      <div class="adSlot__body">${escapeHtml(tr("ad.recommendBody", "ここに常時バナー広告が表示されます（実装準備中）"))}</div>
      <div class="adSlot__live" aria-label="ad slot panel bottom"></div>
    </div>
  `

  const right = `
    ${
      state.projectPath
        ? `<div class="previewImg">
            <div class="previewScale" id="previewScale">
              <img id="previewImg" alt="preview" draggable="false" />
            </div>
            <div class="previewHud">
              <div class="previewHud__left">
                <span class="badge">ライブプレビュー</span>
              </div>
              <div class="previewHud__right">
                <button class="btn btn--soft" id="btnPrevPage">${escapeHtml(tr("main.prevPage", "前"))}</button>
                <button class="btn btn--soft" id="pageIndicator" title="ページ番号を入力して移動">${(state.previewPageIndex || 0) + 1} / ${state.pageCount || 1}</button>
                <button class="btn btn--soft" id="btnNextPage">${escapeHtml(tr("main.nextPage", "次"))}</button>
                <button class="btn btn--soft" id="btnZoomOut">−</button>
                <span class="badge" id="zoomIndicator">${Math.round((Number(state.viewZoom || 1) || 1) * 100)}%</span>
                <button class="btn btn--soft" id="btnZoomIn">＋</button>
                <button class="btn btn--soft" id="btnZoomReset">100%</button>
              </div>
            </div>
            <canvas id="confetti" class="confetti" aria-hidden="true"></canvas>
            <canvas id="overlay" class="overlay"></canvas>
            <div class="emptyHint mainGuidePopup" id="mainGuidePopup">
              <div class="emptyHint__title">まずはPDFに欄（タグ）を置きましょう</div>
              <div class="emptyHint__text">PDF上をダブルクリックしてタグ名と値を入力し、欄を配置できます。</div>
              <div class="emptyHint__actions">
                <button class="btn btn--primary" id="btnGuidePlace">中央に欄を追加</button>
              </div>
            </div>
          </div>`
        : `<div class="previewPlaceholder">プロジェクトZIPを開くとPDFがここに表示されます</div>`
    }
  `

  $("#app").innerHTML = `
    <div class="bgBlobs" aria-hidden="true">
      <div class="blob b1"></div>
      <div class="blob b2"></div>
      <div class="blob b3"></div>
    </div>
    <div class="layout ${state.showPanel ? "" : "layout--nopanel"}">
      <div class="panel">${left}</div>
      <div class="stage stage--nosplit">
        ${right}
      </div>
    </div>
    <div class="toast" id="toast"></div>
    <div class="modal" id="modal" style="display:none"></div>
  `
  refreshAdSlots()

  bind()
  queuePreview()
  tickTimerOnce()
  if (state.justCompleted) {
    burstConfetti()
    state.justCompleted = false
  }
}

function bind() {
  // Tooltips that never go off-screen
  bindTipFloatOnce()
  bindPreviewFitOnce()
  applyPreviewTransform()
  const btnGuidePlace = $("#btnGuidePlace")
  if (btnGuidePlace) btnGuidePlace.onclick = () => {
    const x = Math.round(0.5 * state.pageW)
    const y = Math.round(0.5 * state.pageH)
    if (!x || !y) return
    openPlacePalette({ x, y })
  }


  // Global hotkeys (selection / undo / copy-paste)
  document.onkeydown = async (ev) => {
    if (!state.projectPath) return
    if (isTextEditingTarget(ev.target)) return
    const k = ev.key
    const ctrl = ev.ctrlKey || ev.metaKey

    // Undo / Redo
    if (ctrl && (k === "z" || k === "Z")) {
      ev.preventDefault()
      if (ev.shiftKey) {
        const next = state.redoStack.pop()
        if (!next) return
        state.undoStack.push(snapshotProject())
        await applyProjectSnapshot(next)
        showPage(state.previewPageIndex || 0)
        return
      }
      const prev = state.undoStack.pop()
      if (!prev) return
      state.redoStack.push(snapshotProject())
      await applyProjectSnapshot(prev)
      showPage(state.previewPageIndex || 0)
      return
    }
    if (ctrl && (k === "y" || k === "Y")) {
      ev.preventDefault()
      const next = state.redoStack.pop()
      if (!next) return
      state.undoStack.push(snapshotProject())
      await applyProjectSnapshot(next)
      showPage(state.previewPageIndex || 0)
      return
    }

    // Copy / Paste (copies elements; values stay tag-synced)
    if (ctrl && (k === "c" || k === "C")) {
      if (!state.selectKeys.length) return
      ev.preventDefault()
      const fids = state.selectKeys.filter((fid) => state.placements?.[fid])
      const clip = { fids, placements: {} }
      for (const fid of fids) {
        clip.placements[fid] = { ...(state.placements?.[fid] || {}) }
      }
      state.clipboard = clip
      toast(`コピー: ${fids.length}件`)
      return
    }
    if (ctrl && (k === "v" || k === "V")) {
      if (!state.clipboard?.fids?.length) return
      ev.preventDefault()
      const before = snapshotProject()
      const pasted = []
      const offset = 18
      let n = 0
      const makeFid = () => `f_${Date.now().toString(16)}_${Math.random().toString(16).slice(2, 8)}`
      for (const src of state.clipboard.fids) {
        const pl = state.clipboard.placements?.[src]
        if (!pl) continue
        const newFid = makeFid()
        const tag = String(pl.tag || "").trim()
        if (tag && !state.tags.includes(tag)) state.tags.push(tag)
        state.placements[newFid] = { ...pl, x: Number(pl.x || 0) + offset * (n + 1), y: Number(pl.y || 0) + offset * (n + 1) }
        pasted.push(newFid)
        n++
      }
      if (!pasted.length) return
      state.selectKeys = pasted
      pushUndo(before)
      await window.pywebview.api.set_project_payload?.({ tags: state.tags, values: state.values, placements: state.placements })
      await window.pywebview.api.save_current_project?.(false)
      showPage(state.previewPageIndex || 0)
      render()
      toast(`貼り付け: ${pasted.length}件`)
      return
    }

    // Delete selected
    if (k === "Delete" || k === "Backspace") {
      if (!state.selectKeys.length) return
      ev.preventDefault()
      const before = snapshotProject()
      const del = [...state.selectKeys]
      for (const fid of del) delete state.placements[fid]
      state.selectKeys = []
      pushUndo(before)
      if (window.pywebview?.api?.delete_elements) await window.pywebview.api.delete_elements(del)
      else {
        // fallback: try to persist full payload
        await window.pywebview.api.set_project_payload?.({ tags: state.tags, values: state.values, placements: state.placements })
      }
      await window.pywebview.api.save_current_project?.(false)
      showPage(state.previewPageIndex || 0)
      render()
      toast(`削除: ${del.length}件`)
      return
    }

    // Arrow keys:
    // - Multi-page projects: page navigation with plain arrows
    // - Element nudge: Alt + arrows
    if (k === "ArrowLeft" || k === "ArrowRight" || k === "ArrowUp" || k === "ArrowDown") {
      const hasMultiPage = Number(state.pageCount || 1) > 1
      const plainArrow = !ev.altKey && !ctrl
      if (hasMultiPage && plainArrow) {
        ev.preventDefault()
        state.pageLocked = true
        const cur = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
        const delta = (k === "ArrowRight" || k === "ArrowDown") ? 1 : -1
        await showPage(cur + delta)
        return
      }
      if (!ev.altKey) return
      if (!state.selectKeys.length) return
      ev.preventDefault()
      const step = ev.shiftKey ? 10 : 1
      const dx = k === "ArrowLeft" ? -step : k === "ArrowRight" ? step : 0
      const dy = k === "ArrowUp" ? -step : k === "ArrowDown" ? step : 0
      const before = snapshotProject()
      const page = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
      for (const fid of state.selectKeys) {
        const pl = state.placements?.[fid]
        if (!pl) continue
        if (Number(pl.page || 0) !== page) continue
        pl.x = Math.max(0, Number(pl.x || 0) + dx)
        pl.y = Math.max(0, Number(pl.y || 0) + dy)
        state.placements[fid] = pl
      }
      pushUndo(before)
      await window.pywebview.api.set_project_payload?.({ tags: state.tags, values: state.values, placements: state.placements })
      await window.pywebview.api.save_current_project?.(false)
      drawOverlay()
      showPage(state.previewPageIndex || 0)
      return
    }
  }

  const btnOpen = $("#btnOpen")
  if (btnOpen) btnOpen.onclick = async () => {
    const okToProceed = await showUnlockAd("zip_open")
    if (!okToProceed) return
    const r = await window.pywebview.api.pick_project(
      window.__INPUTSTUDIO_WEB__ ? { zipOnly: true } : undefined
    )
    if (!r.ok) {
      if (r.error) toast(apiErrorMessage(r, r.error))
      return
    }
    try {
      const dir = r.path?.replace(/[/\\][^/\\]+$/, "")
      if (dir) {
        state.lastProjectDir = dir
        saveLocal("inputstudio-last-dir", dir)
      }
    } catch {}
    const loaded = await window.pywebview.api.load_project(r.path)
    if (!loaded.ok) return
    state.projectPath = r.path
    state.projectName = loaded.project
    state.tags = loaded.tags || []
    state.values = loaded.values || {}
    state.placements = loaded.placements || {}
    state.pageCount = loaded.page_count || 1
    state.idx = 0
    state.dropDir = loaded.drop_dir || ""
    state.uiMode = loaded.ui_mode || state.uiMode
    state.lastSession = { path: r.path, workerId: state.workerId, projectName: state.projectName }
    saveLocal("inputstudio-last-session", state.lastSession)
    state.working = false
    state.inPrivate = false
    state.timerStart = null
    state.privateTotal = 0
    const started = await autoStartWorkIfPossible()
    toast(started ? tr("main.toast.projectLoadedAndTimer", "案件を読み込み、作業タイマーを開始しました") : tr("main.toast.projectLoaded", "案件を読み込みました"))
    render()
    await queuePreview()
  }

  // --- Demo (GitHub Pages): load real PDF in browser, but use the same button ---
  const demoPdfFile = $("#demoPdfFile")
  const loadPdfInBrowser = async (file) => {
    toast("PDFを読み込み中…")
    const buf = await file.arrayBuffer()
    // dynamic import pdf.js as ESM from CDN
    const pdfjs = await import("https://cdn.jsdelivr.net/npm/pdfjs-dist@4.8.69/build/pdf.mjs")
    pdfjs.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.8.69/build/pdf.worker.mjs"
    const doc = await pdfjs.getDocument({ data: buf }).promise

    window.__demoPdfDoc = doc
    window.__demoPdfCache = new Map()

    const api = window.pywebview.api
    api.get_preview_png_base64_page = async (page_index) => {
      const idx = Math.max(0, Math.min(doc.numPages - 1, Number(page_index || 0)))
      const cache = window.__demoPdfCache
      if (cache.has(idx)) return cache.get(idx)
      const page = await doc.getPage(idx + 1)
      // Match desktop coordinate system (RENDER_DPI=150)
      const scale = 150 / 72
      const vp = page.getViewport({ scale })
      const canvas = document.createElement("canvas")
      canvas.width = Math.floor(vp.width)
      canvas.height = Math.floor(vp.height)
      const ctx = canvas.getContext("2d")
      await page.render({ canvasContext: ctx, viewport: vp }).promise
      const blob = await new Promise((res) => canvas.toBlob(res, "image/png"))
      const url = URL.createObjectURL(blob)
      const out = { ok: true, png: url, page_display_width: canvas.width, page_display_height: canvas.height, page_index: idx }
      cache.set(idx, out)
      return out
    }
    api.get_preview_png_base64 = async (tag) => {
      const t = String(tag || "").trim()
      const pl = state.placements?.[t]
      const idx = pl ? Number(pl.page || 0) : 0
      return api.get_preview_png_base64_page(idx)
    }

    state.projectPath = "demo:pdf"
    state.projectName = file.name
    state.pageCount = doc.numPages
    state.previewPageIndex = 0
    state.tags = []
    state.values = {}
    state.placements = {}
    toast(`PDFを読み込みました（${doc.numPages}ページ）`)
    render()
  }
  if (demoPdfFile) {
    demoPdfFile.onchange = async () => {
      const file = demoPdfFile.files?.[0]
      if (!file) return
      try {
        await loadPdfInBrowser(file)
      } catch (e) {
        await uiAlert(`PDF読み込みに失敗しました: ${e}`)
      } finally {
        demoPdfFile.value = ""
      }
    }
  }

  const btnOpenPdf = $("#btnOpenPdf")
  if (btnOpenPdf) btnOpenPdf.onclick = async () => {
    if (window.__INPUTSTUDIO_DEMO__ && demoPdfFile) {
      demoPdfFile.click()
      return
    }
    try {
      const api = window.pywebview?.api
      const pick = api?.pick_pdf
      const createSimple = api?.create_project_from_pdf_simple
      if (!pick || !createSimple) {
        await uiAlert("PDFから新規作成する機能が見つかりません。最新版またはバックエンドの create_project_from_pdf_simple/pick_pdf をご用意ください。")
        return
      }
      const r = await pick()
      if (!r?.ok) return
      toast(tr("gate.toastCreateProjectFromPdf", "PDFを読み込み、新規プロジェクトを作成します…"))
      const g = await createSimple(r.path)
      if (!g?.ok || !g.path) {
        await uiAlert((g?.errors || ["PDFをプロジェクト化できませんでした"]).join("\n"))
        return
      }
      const loaded = await api.load_project(g.path)
      if (!loaded?.ok) {
        await uiAlert("新規プロジェクトを開けませんでした")
        return
      }
      state.projectPath = g.path
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.placements = loaded.placements || {}
      state.pageCount = loaded.page_count || 1
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      state.uiMode = loaded.ui_mode || state.uiMode
      state.lastSession = { path: g.path, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      try {
        const dir = g.path.replace(/[/\\][^/\\]+$/, "")
        state.lastProjectDir = dir
        saveLocal("inputstudio-last-dir", dir)
      } catch {}
      state.working = false
      state.inPrivate = false
      state.timerStart = null
      state.privateTotal = 0
      const started = await autoStartWorkIfPossible()
      toast(started ? "新規案件を作成し、作業タイマーを開始しました" : "PDFから新規プロジェクトを作成しました。必要に応じてタグを配置してください。")
      render()
    } catch (e) {
      await uiAlert(`PDFから新規作成に失敗しました: ${e}`)
    }
  }

  const btnResume = $("#btnResume")
  if (btnResume && state.lastSession?.path) {
    btnResume.onclick = async () => {
      const p = state.lastSession.path
      const loaded = await window.pywebview.api.load_project(p)
      if (!loaded.ok) return toast("前回の案件を開けませんでした")
      state.projectPath = p
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.placements = loaded.placements || {}
      state.pageCount = loaded.page_count || 1
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      if (state.lastSession.workerId) state.workerId = state.lastSession.workerId
      state.working = false
      state.inPrivate = false
      state.timerStart = null
      state.privateTotal = 0
      const started = await autoStartWorkIfPossible()
      toast(started ? "前回の案件を読み込み、作業タイマーを開始しました" : "前回の案件を読み込みました")
      render()
      await queuePreview()
    }
  }

  // フォーム付きPDFを扱わない前提のため、自動作成機能は削除

  const btnBackToGate = $("#btnBackToGate")
  if (btnBackToGate) btnBackToGate.onclick = () => {
    state.appStage = "gate"
    state.gate = state.gate || { error: "" }
    render()
  }

  const btnDesign = $("#btnDesign")
  if (btnDesign) btnDesign.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    state.designMode = true
    // design mode operates on element id (fid)
    state.designKey = state.designKey || (Object.keys(state.placements || {})[0] || null)
    await openDesignModal()
  }

  const btnSave = $("#btnSave")
  if (btnSave) btnSave.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    try {
      toast(isWeb() ? "保存してPDFをダウンロードします…" : "保存中…")
      const r = await window.pywebview.api.save_current_project(true)
      if (!r?.ok) return toast(`保存に失敗: ${r?.error || "unknown"}`)
      state.lastSession = { path: state.projectPath, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      if (r?.filled_pdf) state.lastFilledPdf = r.filled_pdf
      if (r?.exports_dir) state.lastExportDir = r.exports_dir
      toast(isWeb() ? "保存しました。保存先を選んでください。" : "案件を保存しました（PDFも生成）")
      if (isWeb() && window.pywebview?.api?.download_filled_pdf) {
        const dl = await window.pywebview.api.download_filled_pdf()
        if (dl?.error === "cancelled") toast("保存をキャンセルしました")
        else if (!dl?.ok) toast(`PDFダウンロードに失敗: ${dl?.error || "unknown"}`)
        else toast("PDFを保存しました")
      }
      render()
    } catch (e) {
      toast(`保存に失敗しました: ${e}`)
    }
  }

  const btnSaveAs = $("#btnSaveAs")
  if (btnSaveAs) btnSaveAs.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    const name0 = String(state.projectName || "案件").trim() || "案件"
    try {
      if (isWeb()) {
        toast("プロジェクトを保存します…")
        if (window.pywebview?.api?.save_project_to_picker) {
          const r = await window.pywebview.api.save_project_to_picker(name0)
          if (r?.error === "cancelled") toast("保存をキャンセルしました")
          else if (!r?.ok) toast(`プロジェクト保存に失敗: ${r?.error || "unknown"}`)
          else toast("プロジェクトを保存しました")
        } else {
          toast("プロジェクト保存機能が見つかりません")
        }
        pulse()
        render()
        return
      }
      const name = await uiPrompt("名前を付けて保存（新しい案件名）", `${name0}-コピー`)
      if (!name) return
      const r = await window.pywebview.api.save_project_as(String(name), true)
      if (!r?.ok || !r.path) return toast(`保存に失敗: ${r?.error || "unknown"}`)
      const loaded = await window.pywebview.api.load_project(r.path)
      if (!loaded?.ok) return toast("保存した案件を開けませんでした")
      state.projectPath = r.path
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.placements = loaded.placements || {}
      state.pageCount = loaded.page_count || 1
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      state.uiMode = loaded.ui_mode || state.uiMode
      state.lastSession = { path: r.path, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      if (r?.filled_pdf) state.lastFilledPdf = r.filled_pdf
      if (r?.exports_dir) state.lastExportDir = r.exports_dir
      toast("名前を付けて保存しました（PDFも生成）")
      pulse()
      render()
      await queuePreview()
    } catch (e) {
      toast(`保存に失敗しました: ${e}`)
    }
  }

  const btnAppendPdf = $("#btnAppendPdf")
  if (btnAppendPdf) {
    btnAppendPdf.onclick = async () => {
      if (!state.projectPath) return toast("先に案件を開いてください")
      const okToProceed = await showUnlockAd("pdf_append")
      if (!okToProceed) return
      const api = window.pywebview?.api
      if (!api?.pick_pdf || !api?.append_pdf_to_project) return toast("PDF追加機能が見つかりません（最新版に更新してください）")
      const curr = Number(state.previewPageIndex || 0) || 0
      const r = await api.pick_pdf()
      if (!r?.ok || !r.path) return
      toast(tr("main.toast.appendProcessing", "PDFを追加して結合中…"))
      const a = await api.append_pdf_to_project(r.path)
      if (!a?.ok) return toast(`PDF追加に失敗: ${apiErrorMessage(a, "unknown")}`)
      state.pageCount = a.page_count || state.pageCount
      resetPreviewViewport({ zoom: 1.0 })
      render()
      await showPage(curr)
      toast(tr("main.toast.appendDone", `PDFを追加しました（合計 ${state.pageCount} ページ）`, { pages: state.pageCount }))
    }
  }
  const btnCopyPageOp = $("#btnCopyPageOp")
  if (btnCopyPageOp) {
    btnCopyPageOp.onclick = async () => {
      if (!state.projectPath) return toast("先に案件を開いてください")
      const api = window.pywebview?.api
      if (!api?.copy_page_with_elements) return toast("ページコピー機能が見つかりません（最新版に更新してください）")
      const curr = Number(state.previewPageIndex || 0) || 0
      const r = await api.copy_page_with_elements(curr)
      if (!r?.ok) return toast(`ページコピーに失敗: ${apiErrorMessage(r, "unknown")}`)
      state.pageCount = Number(r.page_count || state.pageCount) || state.pageCount
      state.placements = r.placements && typeof r.placements === "object" ? r.placements : state.placements
      render()
      await showPage(Number(r.page_index ?? (curr + 1)))
      toast(tr("main.toast.copyDone", "ページをコピーしました"))
    }
  }
  const btnDeletePageOp = $("#btnDeletePageOp")
  if (btnDeletePageOp) {
    btnDeletePageOp.onclick = async () => {
      if (!state.projectPath) return toast("先に案件を開いてください")
      const api = window.pywebview?.api
      if (!api?.delete_page_from_project) return toast("ページ削除機能が見つかりません（最新版に更新してください）")
      const curr = Number(state.previewPageIndex || 0) || 0
      const ok = await uiConfirm(tr("dialog.confirmDeletePage", "現在ページを削除します。配置済み要素も対象ページ分は削除されます。よろしいですか？"))
      if (!ok) return
      const r = await api.delete_page_from_project(curr)
      if (!r?.ok) return toast(`ページ削除に失敗: ${apiErrorMessage(r, "unknown")}`)
      state.pageCount = Number(r.page_count || state.pageCount) || state.pageCount
      state.tags = Array.isArray(r.tags) ? r.tags : state.tags
      state.values = r.values && typeof r.values === "object" ? r.values : state.values
      state.placements = r.placements && typeof r.placements === "object" ? r.placements : state.placements
      render()
      await showPage(Number(r.page_index ?? Math.max(0, curr - 1)))
      toast(tr("main.toast.deleteDone", "ページを削除しました"))
    }
  }
  const btnReorderPdf = $("#btnReorderPdf")
  if (btnReorderPdf) {
    btnReorderPdf.onclick = async () => {
      if (!state.projectPath) return toast("先に案件を開いてください")
      await openPageOpsModal()
    }
  }

  const btnOpenSaved = $("#btnOpenSaved")
  if (btnOpenSaved) btnOpenSaved.onclick = async () => {
    if (!state.projectPath) return
    const api = window.pywebview?.api
    // Preferred: open Explorer via backend
    if (api?.reveal_in_explorer) {
      const r = await api.reveal_in_explorer(state.projectPath)
      if (!r?.ok) toast(`開けませんでした: ${r?.error || "unknown"}`)
      return
    }
    // Fallback: copy path
    try {
      await navigator.clipboard.writeText(String(state.projectPath))
      toast("保存先パスをコピーしました")
    } catch {
      toast(String(state.projectPath))
    }
  }

  const btnOpenFilled = $("#btnOpenFilled")
  if (btnOpenFilled) btnOpenFilled.onclick = async () => {
    const p = state.lastFilledPdf || state.lastExportDir
    if (!p) return
    const api = window.pywebview?.api
    if (api?.reveal_in_explorer) {
      const r = await api.reveal_in_explorer(p)
      if (!r?.ok) toast(`開けませんでした: ${r?.error || "unknown"}`)
      return
    }
    try {
      await navigator.clipboard.writeText(String(p))
      toast("提出PDFパスをコピーしました")
    } catch {
      toast(String(p))
    }
  }

  const btnPrevPage = $("#btnPrevPage")
  if (btnPrevPage) btnPrevPage.onclick = () => {
    state.pageLocked = true
    showPage((state.previewPageIndex || 0) - 1)
  }
  const btnNextPage = $("#btnNextPage")
  if (btnNextPage) btnNextPage.onclick = () => {
    state.pageLocked = true
    showPage((state.previewPageIndex || 0) + 1)
  }
  const pageIndicator = $("#pageIndicator")
  if (pageIndicator) pageIndicator.onclick = async () => {
    const total = Math.max(1, Number(state.pageCount || 1))
    const cur = (Number(state.previewPageIndex || 0) || 0) + 1
    const raw = await uiPrompt(tr("dialog.gotoPage", `移動先ページを入力してください（1-${total}）`, { total }), String(cur))
    if (raw == null) return
    const n = Number(String(raw).trim())
    if (!Number.isFinite(n)) return toast("ページ番号が不正です")
    const idx = Math.max(1, Math.min(total, Math.floor(n))) - 1
    state.pageLocked = true
    await showPage(idx)
  }

  const btnZoomOut = $("#btnZoomOut")
  if (btnZoomOut) btnZoomOut.onclick = () => setViewZoom((Number(state.viewZoom || 1) || 1) - 0.1)
  const btnZoomIn = $("#btnZoomIn")
  if (btnZoomIn) btnZoomIn.onclick = () => setViewZoom((Number(state.viewZoom || 1) || 1) + 0.1)
  const btnZoomReset = $("#btnZoomReset")
  if (btnZoomReset)
    btnZoomReset.onclick = () => {
      state.viewPanX = 0
      state.viewPanY = 0
      setViewZoom(1.0)
    }

  const btnAdmin = $("#btnAdmin")
  if (btnAdmin) btnAdmin.onclick = async () => {
    const ok = await uiConfirm(tr("dialog.switchAdmin", "管理者モードに切り替えます（OCR/設計が表示されます）。よろしいですか？"))
    if (!ok) return
    const r = await window.pywebview.api.set_ui_mode("admin")
    if (!r.ok) return toast("切り替えに失敗しました")
    state.uiMode = "admin"
    toast("管理者モード")
    render()
  }

  const btnLock = $("#btnLock")
  if (btnLock) btnLock.onclick = async () => {
    const ok = await uiConfirm(tr("dialog.switchWorker", "入力者モードに切り替えます（OCR/設計を隠します）。よろしいですか？"))
    if (!ok) return
    const r = await window.pywebview.api.set_ui_mode("worker")
    if (!r.ok) return toast("切り替えに失敗しました")
    state.uiMode = "worker"
    state.designMode = false
    state.addMode = false
    toast("入力者モード")
    render()
  }

  const workerSelect = $("#workerSelect")
  if (workerSelect) {
    workerSelect.onchange = async (e) => {
      state.workerId = e.target.value
      saveLocal("inputstudio-last-worker", state.workerId)
      if (!state.working) {
        const started = await autoStartWorkIfPossible()
        if (started) {
          toast("作業タイマーを開始しました")
          render()
        }
      }
    }
  }

  const btnWorker = $("#btnWorker")
  if (btnWorker) btnWorker.onclick = () => openWorkerModal({ mode: "manage" })

  const historyExport = () => {
    if (!state.history.length) return toast("履歴がありません")
    const header = ["project","path","worker","start_iso","end_iso","duration_sec"].join(",")
    const rows = state.history.map((h) =>
      [h.projectName || "", h.projectPath || "", h.workerName || "", h.start, h.end, h.duration].map((s) =>
        `"${String(s || "").replace(/"/g, '""')}"`
      ).join(",")
    )
    const csv = [header, ...rows].join("\n")
    const blob = new Blob([csv], { type: "text/csv" })
    const a = document.createElement("a")
    a.href = URL.createObjectURL(blob)
    a.download = `inputstudio-history-${Date.now()}.csv`
    a.click()
    URL.revokeObjectURL(a.href)
  }

  const btnHistoryExport = $("#btnHistoryExport")
  if (btnHistoryExport) btnHistoryExport.onclick = historyExport
  const btnHistoryReset = $("#btnHistoryReset")
  if (btnHistoryReset) btnHistoryReset.onclick = async () => {
    const ok = await uiConfirm(tr("dialog.resetHistory", "作業履歴をリセットします（内部保存のみ削除、プロジェクトは残ります）。よろしいですか？"))
    if (!ok) return
    state.history = []
    saveLocal("inputstudio-history", state.history)
    toast("履歴をリセットしました")
  }

  const btnStart = $("#btnStart")
  if (btnStart) btnStart.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    if (!state.workerId) return toast("作業者を選んでください")
    const r = await window.pywebview.api.start_work(state.workerId)
    if (!r.ok) return toast("開始できませんでした")
    state.working = true
    state.inPrivate = false
    state.timerStart = Date.now()
    state.privateTotal = 0
    state.sessionStart = new Date().toISOString()
    state.lastSession = { path: state.projectPath, workerId: state.workerId, projectName: state.projectName }
    saveLocal("inputstudio-last-session", state.lastSession)
    pulse()
    toast("作業タイマーを開始しました")
    render()
  }

  const btnPrivate = $("#btnPrivate")
  if (btnPrivate) btnPrivate.onclick = async () => {
    const r = await window.pywebview.api.toggle_private()
    if (!r.ok) return
    if (!state.inPrivate) {
      state.inPrivate = true
      state._privateStart = Date.now()
      toast("作業タイマーを中断しました")
    } else {
      state.inPrivate = false
      state.privateTotal += (Date.now() - state._privateStart) / 1000
      toast("作業タイマーを再開しました")
      pulse()
    }
    render()
  }

  const btnFinish = $("#btnFinish")
  if (btnFinish) btnFinish.onclick = async () => {
    const ok = await uiConfirm(tr("dialog.finishWork", "勤務を終了して提出物（ZIP）を作成します。よろしいですか？"))
    if (!ok) return
    await pushValue()
    toast("提出物を作成中…")
    const endIso = new Date().toISOString()
    const meta = {
      worker_id: String(state.workerId || ""),
      worker_name: String((state.workers.find((w) => w.id === state.workerId) || {}).name || ""),
      start_iso: String(state.sessionStart || ""),
      end_iso: String(endIso),
      duration_sec: Number(calcNetSeconds() || 0),
      private_sec: Number(state.privateTotal || 0),
      total_tags: Number(state.tags?.length || 0),
      filled_count: Number(filledCount() || 0),
      placement_count: Number(Object.keys(state.placements || {}).length || 0),
      project_name: String(state.projectName || ""),
      project_path: String(state.projectPath || ""),
    }
    const r = await window.pywebview.api.finish(meta)
    if (!r.ok) return toast(`提出物の作成に失敗しました: ${r.error || "unknown"}`)
    if (r?.filled_pdf) state.lastFilledPdf = r.filled_pdf
    if (r?.dir) state.lastExportDir = r.dir
    if (r?.report_pdf) state.lastReportPdf = r.report_pdf
    state.working = false
    state.justCompleted = true
    const duration = calcNetSeconds()
    state.history = [
      ...state.history,
      {
        projectName: state.projectName,
        projectPath: state.projectPath,
        workerId: state.workerId,
        workerName: (state.workers.find((w) => w.id === state.workerId) || {}).name || "",
        start: state.sessionStart,
        end: endIso,
        duration,
      },
    ]
    saveLocal("inputstudio-history", state.history)
    state.sessionStart = null
    state.timerStart = null
    state.privateTotal = 0
    state.inPrivate = false
    render()
    openFinishModal(r)
    // 要望: 終了直後にエクスプローラーで保存階層を開く
    try {
      const api = window.pywebview?.api
      const target = r?.bundle_dir || r?.dir || r?.filled_pdf || ""
      if (api?.reveal_in_explorer && target) {
        await api.reveal_in_explorer(String(target))
      }
    } catch {}
  }

  const btnPrev = $("#btnPrev")
  if (btnPrev) btnPrev.onclick = async () => {
    await pushValue()
    state.pageLocked = false
    state.idx = Math.max(0, state.idx - 1)
    swipe("left")
    render()
  }
  const btnNext = $("#btnNext")
  if (btnNext) btnNext.onclick = async () => {
    const beforeEmpty = isCurrentEmpty()
    await pushValue()
    state.pageLocked = false
    const afterEmpty = isKeyEmpty(state.tags[state.idx])
    if (beforeEmpty && !afterEmpty) {
      state.justCompleted = true
    }
    state.idx = Math.min(state.tags.length - 1, state.idx + 1)
    swipe("right")
    render()
  }
  const btnNextEmpty = $("#btnNextEmpty")
  if (btnNextEmpty) btnNextEmpty.onclick = async () => {
    await pushValue()
    state.pageLocked = false
    for (let i = state.idx + 1; i < state.tags.length; i++) {
      const k = state.tags[i]
      if (isKeyEmpty(k)) {
        state.idx = i
        swipe("right")
        render()
        return
      }
    }
    toast("未入力はありません")
  }
  const btnClear = $("#btnClear")
  if (btnClear) btnClear.onclick = async () => {
    if (!state.tags.length) return
    const k = state.tags[state.idx]
    state.values[k] = ""
    await window.pywebview.api.set_value(k, "")
    pulse()
    render()
  }

  const btnAddField = $("#btnAddField")
  if (btnAddField) {
    btnAddField.onclick = async () => {
      if (!state.projectPath) return toast("先に案件を開いてください")
      const name = (await uiPrompt("追加する欄の名前（例：備考2 / メモ / 追記）", "")) || ""
      const n = name.trim()
      if (!n) return
      state.addDraftName = n
      state.addMode = true
      toast("プレビュー上をクリックして欄を置いてください")
      drawOverlay()
      enableOverlayPointer(true)
    }
  }

  const val = $("#val")
  if (val) {
    val.addEventListener("keydown", async (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault()
        if (state.idx < state.tags.length - 1) {
          if (btnNext?.onclick) await btnNext.onclick()
          // Enterで次へ進んだ直後、次項目の入力欄にフォーカスを戻す
          setTimeout(() => {
            const nextVal = $("#val")
            if (nextVal) {
              nextVal.focus()
              const len = String(nextVal.value || "").length
              try {
                nextVal.setSelectionRange(len, len)
              } catch {}
            }
          }, 0)
        }
      }
    })

    let t = null
    val.addEventListener("input", () => {
      clearTimeout(t)
      t = setTimeout(() => {
        pushValue(true)
      }, 180)
    })
  }

  // 追加モード：クリックで配置
  const ov = $("#overlay")
  if (ov) {
    // enable overlay interactions for selection/editing
    enableOverlayPointer(!!state.projectPath)
    const previewHost = ov.parentElement
    if (previewHost) {
      if (!previewHost.__inputstudioWheelBound) {
        previewHost.__inputstudioWheelBound = true
        previewHost.addEventListener("wheel", (ev) => {
          if (!state.projectPath) return
          ev.preventDefault()
          const withCtrl = !!(ev.ctrlKey || ev.metaKey)
          const step = withCtrl ? 0.2 : 0.1
          const dir = ev.deltaY > 0 ? -step : step
          setViewZoom((Number(state.viewZoom || 1) || 1) + dir)
        }, { passive: false })
      }
      // Fallback for environments where pointer middle-drag stops working after zoom.
      if (!previewHost.__inputstudioMiddlePanBound) {
        previewHost.__inputstudioMiddlePanBound = true
        let middlePanning = false
        let startX = 0
        let startY = 0
        let baseX = 0
        let baseY = 0
        const onMove = (ev) => {
          if (!middlePanning) return
          const dx = ev.clientX - startX
          const dy = ev.clientY - startY
          state.viewPanX = baseX + dx
          state.viewPanY = baseY + dy
          applyPreviewTransform()
          drawOverlay()
        }
        const onUp = () => {
          if (!middlePanning) return
          middlePanning = false
          window.removeEventListener("mousemove", onMove, true)
          window.removeEventListener("mouseup", onUp, true)
        }
        previewHost.addEventListener("mousedown", (ev) => {
          if (!state.projectPath || state.designMode) return
          if (ev.button !== 1) return
          middlePanning = true
          startX = ev.clientX
          startY = ev.clientY
          baseX = Number(state.viewPanX || 0) || 0
          baseY = Number(state.viewPanY || 0) || 0
          window.addEventListener("mousemove", onMove, true)
          window.addEventListener("mouseup", onUp, true)
          ev.preventDefault()
        }, { passive: false })
        previewHost.addEventListener("auxclick", (ev) => {
          if (ev.button === 1) ev.preventDefault()
        }, { passive: false })
      }
    }
    const toPageXY = (ev) => {
      const img = $("#previewImg")
      if (!img || !img.src) return null
      const box = getRenderedContentRect(img, state.pageW, state.pageH)
      const x0 = ev.clientX - box.left
      const y0 = ev.clientY - box.top
      if (x0 < 0 || y0 < 0 || x0 > box.width || y0 > box.height) return null
      const x = (x0 / box.width) * state.pageW
      const y = (y0 / box.height) * state.pageH
      return { x, y, box }
    }

    const hitTest = (pt) => {
      const page = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
      const keys = Object.keys(state.placements || {})
      // last keys = topmost (rough)
      for (let i = keys.length - 1; i >= 0; i--) {
        const fid = keys[i]
        const pl = state.placements?.[fid]
        if (!pl) continue
        if (Number(pl.page || 0) !== page) continue
        const b = placementBoxOnPage(fid, pl)
        const m = Math.max(10, (Number(pl.font_size || 14) || 14) * 0.35)
        if (pt.x >= b.x - m && pt.y >= b.y - m && pt.x <= b.x + b.w + m && pt.y <= b.y + b.h + m) {
          return fid
        }
      }
      return null
    }

    let dragging = false
    let dragStart = null
    let dragBase = null
    let dragUndo = null
    let clickTag = null
    let moved = false
    let panning = false
    let panStartX = 0
    let panStartY = 0
    let panBaseX = 0
    let panBaseY = 0

    // PDFをダブルクリック -> 配置パレット（作業者でも使える）
    ov.ondblclick = (ev) => {
      if (!state.projectPath) return
      // design mode のダブルクリックは既存の処理に任せる
      if (state.designMode) return
      const p = toPageXY(ev)
      if (!p) return
      ev.preventDefault()
      openPlacePalette({ x: p.x, y: p.y }, null)
    }

    ov.onpointerdown = (ev) => {
      if (!state.projectPath) return
      if (state.designMode) return
      if (ev.button === 1 || ev.altKey) {
        panning = true
        panStartX = ev.clientX
        panStartY = ev.clientY
        panBaseX = Number(state.viewPanX || 0) || 0
        panBaseY = Number(state.viewPanY || 0) || 0
        try {
          ov.setPointerCapture?.(ev.pointerId)
        } catch {}
        ev.preventDefault()
        return
      }
      if (state.addMode) return
      // ignore if starting on modal etc
      const p = toPageXY(ev)
      if (!p) return
      const t = hitTest(p)
      clickTag = t
      moved = false
      if (t) {
        if (ev.shiftKey) {
          if (state.selectKeys.includes(t)) state.selectKeys = state.selectKeys.filter((k) => k !== t)
          else state.selectKeys = [...state.selectKeys, t]
        } else {
          state.selectKeys = [t]
        }
        dragUndo = snapshotProject()
        dragging = true
        dragStart = { x: p.x, y: p.y }
        dragBase = {}
        for (const k of state.selectKeys) {
          const pl = state.placements?.[k]
          if (!pl) continue
          dragBase[k] = { x: Number(pl.x || 0), y: Number(pl.y || 0), page: Number(pl.page || 0), font_size: Number(pl.font_size || 14), color: pl.color, line_height: pl.line_height, letter_spacing: pl.letter_spacing }
        }
        ev.preventDefault()
        try {
          ov.setPointerCapture?.(ev.pointerId)
        } catch {}
        drawOverlay()
      } else {
        if (!ev.shiftKey) {
          state.selectKeys = []
          drawOverlay()
        }
      }
    }

    ov.onpointermove = (ev) => {
      if (panning) {
        const dx = ev.clientX - panStartX
        const dy = ev.clientY - panStartY
        state.viewPanX = panBaseX + dx
        state.viewPanY = panBaseY + dy
        applyPreviewTransform()
        drawOverlay()
        return
      }
      if (!dragging || !dragStart || !dragBase) return
      const p = toPageXY(ev)
      if (!p) return
      const dx = p.x - dragStart.x
      const dy = p.y - dragStart.y
      if (Math.abs(dx) + Math.abs(dy) > 1) moved = true
      for (const k of state.selectKeys) {
        const base = dragBase[k]
        if (!base) continue
        const pl = state.placements[k] || {}
        pl.x = Math.max(0, base.x + dx)
        pl.y = Math.max(0, base.y + dy)
        state.placements[k] = pl
      }
      drawOverlay()
    }

    ov.onpointerup = async () => {
      if (panning) {
        panning = false
        return
      }
      if (!dragging) {
        clickTag = null
        return
      }
      dragging = false
      // click (no move) -> open palette for selected
      if (clickTag && !moved) {
        const pl = state.placements?.[clickTag]
        if (pl) {
          openPlacePalette({ x: Number(pl.x || 0), y: Number(pl.y || 0) }, clickTag)
        }
        clickTag = null
        return
      }
      clickTag = null
      // commit drag
      if (dragUndo) {
        pushUndo(dragUndo)
      }
      dragUndo = null
      dragStart = null
      dragBase = null
      try {
        if (window.pywebview?.api?.set_project_payload) {
          await window.pywebview.api.set_project_payload({ tags: state.tags, values: state.values, placements: state.placements })
          await window.pywebview.api.save_current_project?.(false)
        } else {
          // fallback
          for (const k of state.selectKeys) {
            const pl = state.placements?.[k]
            if (pl) await window.pywebview.api.set_element_pos?.(k, pl.x, pl.y)
          }
        }
      } catch {}
      // refresh preview for current page
      showPage(state.previewPageIndex || 0)
    }

    ov.onpointercancel = () => {
      panning = false
      dragging = false
      dragUndo = null
      dragStart = null
      dragBase = null
      clickTag = null
    }

    // Prevent browser middle-click auto-scroll from stealing drag interactions.
    ov.onmousedown = (ev) => {
      if (ev.button === 1) ev.preventDefault()
    }
    ov.onauxclick = (ev) => {
      if (ev.button === 1) ev.preventDefault()
    }

    ov.onclick = async (ev) => {
      if (!state.addMode) return
      const p = toPageXY(ev)
      if (!p) return
      const x = p.x
      const y = p.y
      toast("欄を追加中…")
      const fs = Number(state.defaultFontSize || 14) || 14
      let r = await window.pywebview.api.add_text_field(state.addDraftName, state.previewPageIndex || 0, x, y, fs)
      // Recover if backend lost project context (rare, but observed)
      if (!r.ok && r.error === "no_project" && state.projectPath && window.pywebview.api.load_project) {
        try {
          await window.pywebview.api.load_project(state.projectPath)
          r = await window.pywebview.api.add_text_field(state.addDraftName, state.previewPageIndex || 0, x, y, fs)
        } catch {}
      }
      if (!r.ok) {
        state.addMode = false
        enableOverlayPointer(false)
        drawOverlay()
        await uiAlert(`追加に失敗: ${r.error || "unknown"}`)
        return
      }
      const fid = r.fid
      const tag = r.tag
      if (!state.tags.includes(tag)) state.tags.push(tag)
      if (state.values[tag] == null) state.values[tag] = ""
      state.placements[fid] = { tag, page: state.previewPageIndex || 0, x, y, font_size: fs, color: "#0f172a", line_height: 1.2, letter_spacing: DEFAULT_LETTER_SPACING }
      state.selectKeys = [fid]
      state.idx = state.tags.indexOf(tag)
      await window.pywebview.api.save_current_project(false)
      state.addMode = false
      enableOverlayPointer(false)
      toast(`追加しました：${tag}`)
      pulse()
      render()
    }
  }
}

function isKeyEmpty(k) {
  const v = (state.values[k] || "").replaceAll("<br>", "").trim()
  return !v
}
function isCurrentEmpty() {
  if (!state.tags.length) return true
  return isKeyEmpty(state.tags[state.idx])
}

async function pushValue() {
  if (!state.tags.length) return
  const key = state.tags[state.idx]
  const raw = ($("#val")?.value || "").replaceAll("\r\n", "\n")
  const value = raw.replaceAll("\n", "<br>")
  state.values[key] = value
  await window.pywebview.api.set_value(key, value)
  queuePreview(key)
}

let previewReq = 0
let pageReq = 0
async function queuePreview(key) {
  if (!state.projectPath) {
    const img = $("#previewImg")
    if (img) {
      img.src = ""
      img.style.visibility = "hidden"
    }
    return
  }

  // ページ固定中は、選択タグに関係なく現在ページを維持
  if (state.pageLocked && window.pywebview?.api?.get_preview_png_base64_page) {
    await showPage(state.previewPageIndex || 0)
    return
  }

  // タグが無い（=新規直後など）でも、まずは1ページ目を表示できるようにする
  if (!state.tags.length) {
    if (window.pywebview?.api?.get_preview_png_base64_page) {
      await showPage(0)
      return
    }
    return
  }

  const k = key || state.tags[state.idx]
  const my = ++previewReq
  let r = await window.pywebview.api.get_preview_png_base64(k)
  if (r && !r.ok && r.error === "no_project" && state.projectPath && window.pywebview?.api?.load_project) {
    try {
      await window.pywebview.api.load_project(state.projectPath)
      r = await window.pywebview.api.get_preview_png_base64(k)
    } catch {}
  }
  if (my !== previewReq) return
  if (r.ok) {
    const img = $("#previewImg")
    if (img) {
      img.onload = () => (img.style.visibility = "visible")
      img.onerror = () => {
        img.style.visibility = "hidden"
      }
      img.style.visibility = "hidden"
      img.src = r.png_data || r.png
    }
    if (img && img.naturalWidth && img.naturalHeight) {
      state.pageW = img.naturalWidth
      state.pageH = img.naturalHeight
    } else {
      state.pageW = r.page_display_width || state.pageW
      state.pageH = r.page_display_height || state.pageH
    }
    state.previewPageIndex = Number.isFinite(r.page_index) ? r.page_index : state.previewPageIndex
    const p = $("#pageIndicator")
    if (p) p.textContent = `${(state.previewPageIndex || 0) + 1} / ${state.pageCount || 1}`
    drawOverlay()
  } else {
    const img = $("#previewImg")
    if (img) {
      img.src = ""
      img.style.visibility = "hidden"
    }
    toast(`プレビュー取得に失敗: ${r?.error || "unknown"}`)
  }
}

async function loadWorkers() {
  const r = await window.pywebview.api.get_workers()
  if (!r.ok) return
  state.workers = r.workers || []
  const last = loadLocal("inputstudio-last-worker", null)
  if (last && state.workers.some((w) => w.id === last)) state.workerId = last
  else state.workerId = r.last_worker_id || (state.workers[0] ? state.workers[0].id : null)
}

async function autoStartWorkIfPossible() {
  // 要望: PDF/案件を読み込んだら自動で作業開始（押し忘れ防止）
  if (!state.projectPath) return false
  if (state.working) return true
  if (!state.workerId) {
    toast("作業者を選ぶと自動で作業タイマー開始します")
    return false
  }
  try {
    const api = window.pywebview?.api
    if (!api?.start_work) return false
    const r = await api.start_work(state.workerId)
    if (!r?.ok) return false
    state.working = true
    state.inPrivate = false
    state.timerStart = Date.now()
    state.privateTotal = 0
    state.sessionStart = new Date().toISOString()
    state.lastSession = { path: state.projectPath, workerId: state.workerId, projectName: state.projectName }
    saveLocal("inputstudio-last-session", state.lastSession)
    return true
  } catch {
    return false
  }
}

let _timerTicker = null
function ensureTimerTicker() {
  if (_timerTicker) return
  _timerTicker = setInterval(() => {
    if (state.appStage !== "main") return
    const el = document.querySelector(".status__time")
    if (!el) return
    el.textContent = fmtTime(calcNetSeconds())
  }, 1000)
}

function tickTimerOnce() {
  const el = $(".status__time")
  if (el) el.textContent = fmtTime(calcNetSeconds())
  ensureTimerTicker()
}

function openFinishModal(result) {
  const modal = $("#modal")
  if (!modal) return
  const dir = String(result?.dir || "")
  const zip = String(result?.zip || "")
  const pdf = String(result?.filled_pdf || "")
  const bundleDir = String(result?.bundle_dir || "")
  const reportPdf = String(result?.report_pdf || "")

  const close = () => {
    modal.style.display = "none"
    modal.innerHTML = ""
  }

  if (isWeb()) {
    modal.style.display = "block"
    modal.innerHTML = `
      <div class="modal__backdrop" id="modalClose"></div>
      <div class="modal__card" style="max-width:480px">
        <div class="modal__title">提出データを作成しました</div>
        <div class="label" style="margin-top:6px; line-height:1.7">
          完成PDFをダウンロードしました。メールに添付して送信してください。お疲れ様でした！
        </div>
        <div class="row" style="margin-top:14px; justify-content:flex-end; gap:10px">
          <button class="btn btn--primary" id="btnFinishDownload">PDFを再ダウンロード</button>
          <button class="btn btn--soft" id="btnFinishClose">閉じる</button>
        </div>
      </div>
    `
    $("#modalClose").onclick = close
    $("#btnFinishClose").onclick = close
    $("#btnFinishDownload").onclick = async () => {
      if (window.pywebview?.api?.download_filled_pdf) {
        await window.pywebview.api.download_filled_pdf()
        toast("ダウンロードを開始しました")
      }
    }
    return
  }

  modal.style.display = "block"
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card" style="max-width:680px">
      <div class="modal__title">提出データを作成しました</div>
      <div class="label" style="margin-top:6px; line-height:1.7">
        今回の作業成果物を生成しました。次のボタンを押し、出てきたデータをメールに添付して送信してください。お疲れ様でした！
      </div>
      <div class="field" style="margin-top:12px">
        <div class="label">フォルダ</div>
        <div class="pathLine" title="${escapeHtml(dir)}"><span class="pathValue">${escapeHtml(dir)}</span></div>
      </div>
      <div class="field" style="margin-top:10px">
        <div class="label">ZIP（添付）</div>
        <div class="pathLine" title="${escapeHtml(zip)}"><span class="pathValue">${escapeHtml(zip)}</span></div>
      </div>
      <div class="field" style="margin-top:10px">
        <div class="label">送付用フォルダ（PDF + project.json + template.pdf）</div>
        <div class="pathLine" title="${escapeHtml(bundleDir || dir)}"><span class="pathValue">${escapeHtml(bundleDir || dir)}</span></div>
      </div>
      <div class="field" style="margin-top:10px">
        <div class="label">PDF（確認用）</div>
        <div class="pathLine" title="${escapeHtml(pdf)}"><span class="pathValue">${escapeHtml(pdf)}</span></div>
      </div>
      <div class="field" style="margin-top:10px">
        <div class="label">報告書PDF</div>
        <div class="pathLine" title="${escapeHtml(reportPdf || "")}"><span class="pathValue">${escapeHtml(reportPdf || "（未生成）")}</span></div>
      </div>
      <div class="row" style="margin-top:14px; justify-content:flex-end">
        <button class="btn btn--primary" id="btnOpenAttachment">フォルダを開く</button>
        <button class="btn btn--soft" id="btnFinishClose">閉じる</button>
      </div>
    </div>
  `
  const closeEl = $("#modalClose")
  if (closeEl) closeEl.onclick = close
  const closeBtn = $("#btnFinishClose")
  if (closeBtn) closeBtn.onclick = close
  const openBtn = $("#btnOpenAttachment")
  if (openBtn)
    openBtn.onclick = async () => {
      const api = window.pywebview?.api
      const target = bundleDir || dir || zip || pdf
      if (!target) return
      if (api?.reveal_in_explorer) {
        const r = await api.reveal_in_explorer(target)
        if (!r?.ok) toast(`開けませんでした: ${r?.error || "unknown"}`)
        return
      }
      try {
        await navigator.clipboard.writeText(String(target))
        toast("パスをコピーしました")
      } catch {
        toast(String(target))
      }
    }
}

async function openPageOpsModal() {
  const modal = $("#modal")
  if (!modal) return
  const close = () => {
    modal.style.display = "none"
    modal.innerHTML = ""
  }
  modal.style.display = "block"
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card" style="width:min(1200px, calc(100vw - 40px)); max-width:1200px">
      <div class="modal__title">PDF並べ替え</div>
      <div class="label" style="line-height:1.7">現在ページ: ${(Number(state.previewPageIndex || 0) || 0) + 1} / ${Math.max(1, Number(state.pageCount || 1))}</div>
      <div class="label" style="margin-top:10px">下のカードをドラッグ&ドロップして直感的に並び替えできます。</div>
      <div class="pageOpsBoard" id="pageOpsBoard"></div>
      <div class="row" style="margin-top:14px; justify-content:flex-end">
        <button class="btn btn--primary" id="poApplyOrder" disabled>並び替えを保存</button>
        <button class="btn btn--soft" id="poClose">閉じる</button>
      </div>
    </div>
  `
  $("#modalClose").onclick = close
  $("#poClose").onclick = close

  const api = window.pywebview?.api
  const curr = Number(state.previewPageIndex || 0) || 0
  const totalPages = Math.max(1, Number(state.pageCount || 1) || 1)
  let order = Array.from({ length: totalPages }, (_, i) => i)
  let dragPos = -1
  const poApplyOrder = $("#poApplyOrder")
  const board = $("#pageOpsBoard")

  const renderBoard = async () => {
    if (!board) return
    const pageModels = await Promise.all(
      order.map(async (oldPageIdx, pos) => {
        const pr = await api.get_preview_png_base64_page(oldPageIdx)
        const src = pr?.ok ? String(pr.png || "") : ""
        const srcData = pr?.ok ? String(pr.png_data || "") : ""
        return {
          pos,
          oldPageIdx,
          src,
          srcData,
        }
      })
    )
    const modelByPos = new Map(pageModels.map((m) => [m.pos, m]))
    const cards = pageModels
      .map(({ pos, oldPageIdx, src, srcData }) => {
        const initialSrc = src || srcData || ""
        const hasImage = !!initialSrc
        const currentCls = oldPageIdx === curr ? " is-current" : ""
        return `
          <div class="pageCard${currentCls}" draggable="true" data-pos="${pos}" data-old="${oldPageIdx}">
            <div class="pageCard__thumb">${hasImage ? `<img class="pageCardImg" draggable="false" data-pos="${pos}" src="${escapeHtml(initialSrc)}" alt="page ${oldPageIdx + 1}" />` : "<div class=\"pageCard__noimg\">No Image</div>"}</div>
            <div class="pageCard__meta">
              <span class="badge">表示順 ${pos + 1}</span>
              <span class="badge badge--soft">元ページ ${oldPageIdx + 1}</span>
            </div>
          </div>
        `
      })
      .join("")
    board.innerHTML = cards

    board.querySelectorAll(".pageCardImg").forEach((img) => {
      img.addEventListener("error", () => {
        const pos = Number(img.getAttribute("data-pos") || "-1")
        const model = modelByPos.get(pos)
        if (!model) return
        const current = String(img.getAttribute("src") || "")
        if (model.srcData && current !== model.srcData) {
          img.setAttribute("src", model.srcData)
          return
        }
      })
    })

    board.querySelectorAll(".pageCard").forEach((el) => {
      el.addEventListener("dragstart", (ev) => {
        dragPos = Number(el.getAttribute("data-pos") || "-1")
        ev.dataTransfer?.setData("text/plain", String(dragPos))
        ev.dataTransfer.effectAllowed = "move"
      })
      el.addEventListener("dragover", (ev) => {
        ev.preventDefault()
        el.classList.add("is-over")
      })
      el.addEventListener("dragleave", () => {
        el.classList.remove("is-over")
      })
      el.addEventListener("drop", (ev) => {
        ev.preventDefault()
        el.classList.remove("is-over")
        const toPos = Number(el.getAttribute("data-pos") || "-1")
        const fromPos = Number(ev.dataTransfer?.getData("text/plain") || dragPos)
        if (fromPos < 0 || toPos < 0 || fromPos === toPos) return
        const next = [...order]
        const [moved] = next.splice(fromPos, 1)
        next.splice(toPos, 0, moved)
        order = next
        if (poApplyOrder) poApplyOrder.disabled = false
        renderBoard()
      })
    })
  }

  await renderBoard()

  if (poApplyOrder) {
    poApplyOrder.onclick = async () => {
      if (!api?.reorder_pages) return toast("ページ並び替え機能が見つかりません（最新版に更新してください）")
      const newCurrent = Math.max(0, order.indexOf(curr))
      toast("ページ順を保存中…")
      const r = await api.reorder_pages(order)
      if (!r?.ok) return toast(`ページ並び替えに失敗: ${apiErrorMessage(r, "unknown")}`)
      state.pageCount = Number(r.page_count || state.pageCount) || state.pageCount
      state.placements = r.placements && typeof r.placements === "object" ? r.placements : state.placements
      close()
      render()
      await showPage(newCurrent)
      toast("ページ順を更新しました")
    }
  }
}

function pulse() {
  document.body.classList.remove("pulse")
  void document.body.offsetWidth
  document.body.classList.add("pulse")
  setTimeout(() => document.body.classList.remove("pulse"), 420)
}

function swipe(dir) {
  document.body.dataset.swipe = dir
  setTimeout(() => (document.body.dataset.swipe = ""), 260)
}

function openWorkerModal(opts = {}) {
  const modal = $("#modal")
  const NEW = "__new__"
  const mode = String(opts.mode || "manage") // "manage" | "create"
  let editingId = mode === "create" ? NEW : state.workerId || (state.workers[0] ? state.workers[0].id : NEW)

  const close = () => {
    modal.style.display = "none"
    modal.innerHTML = ""
  }

  const renderModal = () => {
    const isNew = editingId === NEW
    const current = isNew ? {} : state.workers.find((w) => w.id === editingId) || {}
    modal.style.display = "block"
    modal.innerHTML = `
      <div class="modal__backdrop" id="modalClose"></div>
      <div class="modal__card">
        <div class="modal__title">作業者の登録</div>
        <div class="label">作業者を追加・編集できます（開始/終了の記録にも使います）。</div>

        ${
          mode === "manage"
            ? `<div class="row" style="margin-top:10px">
                <div class="field" style="flex:1">
                  <div class="label">一覧</div>
                  <select id="mPick">
                    <option value="${NEW}" ${isNew ? "selected" : ""}>（新規）</option>
                    ${state.workers.map((w) => `<option value="${escapeHtml(w.id)}" ${w.id === editingId ? "selected" : ""}>${escapeHtml(w.name)}</option>`).join("")}
                  </select>
                </div>
                <button class="btn btn--soft" id="mNew">新規</button>
              </div>`
            : `<div class="row" style="margin-top:10px">
                <div class="badge">新規登録</div>
                <span class="label">（既存一覧は表示しません）</span>
              </div>`
        }

        <div class="field" style="margin-top:10px">
          <div class="label">名前</div>
          <input class="input" id="mName" value="${escapeHtml(current.name || "")}" placeholder="例）作業者A">
        </div>
        <div class="field">
          <div class="label">振込先</div>
          <input class="input" id="mBank" value="${escapeHtml(current.bank || "")}" placeholder="○○銀行　普通　1234567　カナザワ　タロウ">
        </div>

        <div class="row spread" style="margin-top:14px">
          <button class="btn btn--soft" id="modalCancel">閉じる</button>
          <div class="row">
            ${mode === "manage" && !isNew && editingId ? `<button class="btn btn--danger" id="mDelete">削除</button>` : ""}
            <button class="btn btn--primary" id="modalSave">保存</button>
          </div>
        </div>
      </div>
    `

    $("#modalClose").onclick = close
    $("#modalCancel").onclick = close
    const pick = $("#mPick")
    if (mode === "manage" && pick) pick.onchange = (e) => {
      editingId = e.target.value
      renderModal()
    }
    const btnNew = $("#mNew")
    if (mode === "manage" && btnNew) btnNew.onclick = () => {
      editingId = NEW
      renderModal()
      $("#mName")?.focus?.()
    }
    const btnDel = $("#mDelete")
    if (mode === "manage" && btnDel) btnDel.onclick = async () => {
      const ok = await uiConfirm(tr("dialog.deleteWorker", "この作業者を削除しますか？"))
      if (!ok) return
      const r = await window.pywebview.api.delete_worker?.(String(editingId))
      if (!r?.ok) return toast(`削除に失敗: ${r?.error || "unknown"}`)
      await loadWorkers()
      editingId = state.workerId || (state.workers[0] ? state.workers[0].id : NEW)
      pulse()
      toast("削除しました")
      renderModal()
      render()
    }

    $("#modalSave").onclick = async () => {
      const w = {
        id: editingId === NEW ? null : editingId,
        name: $("#mName").value.trim(),
        bank: $("#mBank").value.trim(),
      }
      if (!w.name) return toast("名前を入れてください")
      const r = await window.pywebview.api.upsert_worker(w)
      if (!r.ok) return toast("保存できませんでした")
      await loadWorkers()
      state.workerId = r.id
      saveLocal("inputstudio-last-worker", state.workerId)
      editingId = state.workerId || NEW
      pulse()
      toast("保存しました")
      close()
      render()
    }
  }

  renderModal()
}

// confetti（軽量）
function burstConfetti() {
  const c = $("#confetti")
  if (!c) return
  const ctx = c.getContext("2d")
  const rect = c.getBoundingClientRect()
  c.width = Math.max(1, Math.floor(rect.width * devicePixelRatio))
  c.height = Math.max(1, Math.floor(rect.height * devicePixelRatio))
  ctx.scale(devicePixelRatio, devicePixelRatio)

  const parts = []
  const colors = ["#ff6aa2", "#7c5cff", "#5ad7ff", "#ffd36a", "#7cffb2"]
  for (let i = 0; i < 90; i++) {
    parts.push({
      x: rect.width * 0.5,
      y: rect.height * 0.2,
      vx: (Math.random() - 0.5) * 6,
      vy: Math.random() * -5 - 2,
      g: 0.18 + Math.random() * 0.08,
      s: 2 + Math.random() * 3,
      r: Math.random() * Math.PI,
      vr: (Math.random() - 0.5) * 0.2,
      c: colors[i % colors.length],
      a: 1,
    })
  }
  const t0 = performance.now()
  function step(t) {
    const dt = (t - t0) / 1000
    ctx.clearRect(0, 0, rect.width, rect.height)
    for (const p of parts) {
      p.vy += p.g
      p.x += p.vx
      p.y += p.vy
      p.r += p.vr
      p.a = Math.max(0, 1 - dt / 1.2)
      ctx.save()
      ctx.globalAlpha = p.a
      ctx.translate(p.x, p.y)
      ctx.rotate(p.r)
      ctx.fillStyle = p.c
      ctx.fillRect(-p.s, -p.s, p.s * 2, p.s * 2)
      ctx.restore()
    }
    if (dt < 1.2) requestAnimationFrame(step)
    else ctx.clearRect(0, 0, rect.width, rect.height)
  }
  requestAnimationFrame(step)
}

// ---- design mode ----
async function openDesignModal() {
  const modal = $("#modal")
  modal.style.display = "block"
  const allItems = Object.entries(state.placements || {})
    .map(([fid, pl]) => {
      const p = pl && typeof pl === "object" ? pl : {}
      const tag = String(p.tag || "").trim() || "(タグ未設定)"
      const page = Number(p.page || 0) + 1
      return { fid: String(fid), tag, page, label: `${tag}（p${page}）` }
    })
    .filter((x) => x.fid)
  if (!state.designKey || !state.placements?.[state.designKey]) {
    state.designKey = allItems[0]?.fid || null
  }
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card">
      <div class="modal__title">設計（統括）モード</div>
      <div class="label" style="margin-bottom:8px">タグを選んで、プレビュー上をクリックで移動。矢印で微調整。</div>

      <div class="row" style="margin-top:6px">
        <div class="field" style="flex:1">
          <div class="label">検索</div>
          <input class="input" id="dSearch" placeholder="例）氏名 / 住所 / 金額 …" />
        </div>
        <div class="field" style="width:140px">
          <div class="label">移動幅</div>
          <select id="dStep">
            <option value="1">1px</option>
            <option value="2" selected>2px</option>
            <option value="5">5px</option>
            <option value="10">10px</option>
          </select>
        </div>
      </div>

      <div class="row" style="margin-top:10px">
        <div class="field" style="flex:1">
          <div class="label">対象要素</div>
          <select id="dKey">
            ${allItems.map((it) => `<option value="${escapeHtml(it.fid)}" ${it.fid === state.designKey ? "selected" : ""}>${escapeHtml(it.label)}</option>`).join("")}
          </select>
        </div>
        <button class="btn btn--soft" id="dPrev">前</button>
        <button class="btn btn--soft" id="dNext">次</button>
        <button class="btn btn--soft" id="dFocus">表示</button>
      </div>

      <div class="row" style="margin-top:10px">
        <button class="btn btn--soft" id="dUp">↑</button>
        <button class="btn btn--soft" id="dLeft">←</button>
        <button class="btn btn--soft" id="dRight">→</button>
        <button class="btn btn--soft" id="dDown">↓</button>
        <span class="badge" id="dPos">x:- y:- ${tipIcon(5, "ここで配置中のタグ座標を確認・微調整できます。")}</span>
      </div>

      <div class="row spread" style="margin-top:14px">
        <button class="btn btn--soft" id="dClose">閉じる</button>
        <div class="row">
          <button class="btn btn--tint" id="dToggleOverlay">プレビューで移動: ON</button>
          <button class="btn btn--primary" id="dSave">保存</button>
        </div>
      </div>
    </div>
  `

  const close = () => {
    state.designMode = false
    modal.style.display = "none"
    drawOverlay()
  }
  $("#modalClose").onclick = close
  $("#dClose").onclick = close

  $("#dKey").onchange = async (e) => {
    state.designKey = e.target.value
    await focusDesignKey()
  }
  $("#dPrev").onclick = async () => {
    const ids = allItems.map((x) => x.fid)
    const i = Math.max(0, ids.indexOf(state.designKey) - 1)
    state.designKey = ids[i] || state.designKey
    $("#dKey").value = state.designKey
    await focusDesignKey()
  }
  $("#dNext").onclick = async () => {
    const ids = allItems.map((x) => x.fid)
    const i = Math.min(ids.length - 1, ids.indexOf(state.designKey) + 1)
    state.designKey = ids[i] || state.designKey
    $("#dKey").value = state.designKey
    await focusDesignKey()
  }
  $("#dFocus").onclick = async () => {
    await focusDesignKey()
  }
  $("#dSave").onclick = async () => {
    const r = await window.pywebview.api.save_current_project(false)
    if (!r.ok) {
      await uiAlert(`保存に失敗: ${r.error || "unknown"}`)
      return
    }
    toast("保存しました")
    pulse()
  }

  let overlayEnabled = true
  $("#dToggleOverlay").onclick = () => {
    overlayEnabled = !overlayEnabled
    $("#dToggleOverlay").textContent = `プレビューで移動: ${overlayEnabled ? "ON" : "OFF"}`
    const ov = $("#overlay")
    if (ov) ov.style.pointerEvents = overlayEnabled && state.designMode ? "auto" : "none"
    drawOverlay()
  }

  const nudge = async (dx, dy) => {
    const info = await window.pywebview.api.get_element_info(state.designKey)
    if (!info.ok) return toast("対象要素が見つかりません")
    const x = (info.x || 0) + dx
    const y = (info.y || 0) + dy
    await window.pywebview.api.set_element_pos(state.designKey, x, y)
    await focusDesignKey(false)
  }
  const step = () => Number($("#dStep")?.value || "2") || 2
  $("#dUp").onclick = () => nudge(0, -step())
  $("#dDown").onclick = () => nudge(0, step())
  $("#dLeft").onclick = () => nudge(-step(), 0)
  $("#dRight").onclick = () => nudge(step(), 0)

  // 検索（option絞り込み）
  const filterOptions = () => {
    const q = ($("#dSearch")?.value || "").trim().toLowerCase()
    const sel = $("#dKey")
    if (!sel) return
    const filtered = allItems.filter((it) => (q ? it.label.toLowerCase().includes(q) : true))
    sel.innerHTML = filtered
      .map((it) => `<option value="${escapeHtml(it.fid)}" ${it.fid === state.designKey ? "selected" : ""}>${escapeHtml(it.label)}</option>`)
      .join("")
  }
  $("#dSearch").addEventListener("input", () => {
    filterOptions()
  })

  // overlay click -> move
  const ov = $("#overlay")
  if (ov) {
    enableOverlayPointer(overlayEnabled && state.designMode)

    // “ドラッグで置ける” を追加
    let dragging = false
    let lastSent = 0
    const toXY = (ev) => {
      const img = $("#previewImg")
      if (!img || !img.src) return null
      // Use actual rendered content rect (object-fit: contain) to avoid drift.
      const box = getRenderedContentRect(img, state.pageW, state.pageH)
      const x0 = ev.clientX - box.left
      const y0 = ev.clientY - box.top
      if (x0 < 0 || y0 < 0 || x0 > box.width || y0 > box.height) return null
      const x = (x0 / box.width) * state.pageW
      const y = (y0 / box.height) * state.pageH
      return { x, y }
    }

    ov.onpointerdown = async (ev) => {
      if (!state.designMode || !overlayEnabled) return
      const p = toXY(ev)
      if (!p) return
      dragging = true
      ov.setPointerCapture?.(ev.pointerId)
      await window.pywebview.api.set_element_pos(state.designKey, p.x, p.y)
      state.designPos = { x: p.x, y: p.y }
      drawOverlay()
      pulse()
    }
    ov.onpointermove = async (ev) => {
      if (!dragging) return
      const now = Date.now()
      if (now - lastSent < 35) return
      lastSent = now
      const p = toXY(ev)
      if (!p) return
      await window.pywebview.api.set_element_pos(state.designKey, p.x, p.y)
      state.designPos = { x: p.x, y: p.y }
      drawOverlay()
    }
    ov.onpointerup = async () => {
      if (!dragging) return
      dragging = false
      await focusDesignKey(false)
    }

    // click（細かい置き直し）
    ov.onclick = async (ev) => {
      if (!state.designMode || !overlayEnabled) return
      const p = toXY(ev)
      if (!p) return
      const { x, y } = p
      await window.pywebview.api.set_element_pos(state.designKey, x, y)
      state.designPos = { x, y }
      await focusDesignKey(false)
      pulse()
    }

    // double click -> open palette (タグ/値/サイズ指定して配置)
    ov.ondblclick = (ev) => {
      if (!state.designMode || !overlayEnabled) return
      const p = toXY(ev)
      if (!p) return
      ev.preventDefault()
      openPlacePalette(p)
    }
  }

  await focusDesignKey()
}

function openPlacePalette(pt, editFid = null) {
  const modal = $("#modal")
  if (!modal) return
  const close = () => {
    modal.style.display = "none"
    modal.innerHTML = ""
  }
  const pageIdx = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
  const isEdit = !!editFid
  const currentPl = isEdit ? (state.placements?.[editFid] || {}) : {}
  const defaultFs = Number(currentPl.font_size || state.defaultFontSize || 14) || 14
  const curColor = String(currentPl.color || "#0f172a")
  const curLH = Number(currentPl.line_height || 1.2) || 1.2
  const curLS = Number(currentPl.letter_spacing ?? DEFAULT_LETTER_SPACING) || DEFAULT_LETTER_SPACING
  let writingMode = String(currentPl.writing_mode || "horizontal")
  const tagsOptions = state.tags.map((t) => `<option value="${escapeHtml(t)}"></option>`).join("")
  modal.style.display = "block"
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card modal__card--anchored" id="paletteCard" style="max-width:520px">
      <div class="modal__title">${isEdit ? "要素を編集" : "PDFに配置（ダブルクリック）"}</div>
      <div class="label">${isEdit ? "選択中の要素の値・見た目を調整します。" : "タグ一覧でクリックして選択するか、タグ名を直接入力して配置します。"}</div>
      <div class="field" style="margin-top:8px">
        <div class="label">タグ</div>
        <input class="input" id="pTag" placeholder="タグ名（一覧でクリック or 新規入力）" ${isEdit ? "disabled" : ""} />
      </div>
      <div class="field">
        <div class="label">値</div>
        <textarea class="textarea" id="pVal" rows="3" placeholder="ここに値を入力（Enterで改行）"></textarea>
      </div>
      <div class="row">
        <div class="field" style="width:160px">
          <div class="label">サイズ</div>
          <div class="spin">
            <input class="input" id="pSize" inputmode="numeric" value="${defaultFs}" />
            <div class="spin__btns">
              <button class="spin__btn" id="pSizeUp" type="button">▲</button>
              <button class="spin__btn" id="pSizeDown" type="button">▼</button>
            </div>
          </div>
        </div>
        <div class="field" style="width:200px">
          <div class="label">縦書き/横書き</div>
          <div class="row" style="gap:8px">
            <button class="btn btn--soft" id="pWmH" type="button">横書</button>
            <button class="btn btn--soft" id="pWmV" type="button">縦書</button>
          </div>
        </div>
        <div class="field" style="width:180px">
          <div class="label">色</div>
          <div class="colorPicker">
            <div class="swatches" id="pSwatches"></div>
            <input class="input" id="pColor" value="${escapeHtml(curColor)}" readonly />
          </div>
        </div>
        <div class="field" style="width:120px">
          <div class="label">ページ</div>
          <input class="input" id="pPage" inputmode="numeric" value="${pageIdx + 1}" />
        </div>
      </div>
      <div class="row">
        <div class="field" style="width:160px">
          <div class="label">行間</div>
          <div class="spin">
            <input class="input" id="pLineH" inputmode="decimal" value="${curLH}" />
            <div class="spin__btns">
              <button class="spin__btn" id="pLineHUp" type="button">▲</button>
              <button class="spin__btn" id="pLineHDown" type="button">▼</button>
            </div>
          </div>
        </div>
        <div class="field" style="width:160px">
          <div class="label">字間</div>
          <div class="spin">
            <input class="input" id="pLetterS" inputmode="decimal" value="${curLS}" />
            <div class="spin__btns">
              <button class="spin__btn" id="pLetterSUp" type="button">▲</button>
              <button class="spin__btn" id="pLetterSDown" type="button">▼</button>
            </div>
          </div>
        </div>
        <div class="field" style="flex:1">
          <div class="label">座標 (x,y)</div>
          <input class="input" id="pPos" value="${Math.round(pt.x)}, ${Math.round(pt.y)}" disabled />
        </div>
      </div>
      <div class="row" style="margin-top:10px; justify-content:flex-end">
        ${isEdit ? `<button class="btn btn--danger" id="pDelete">削除</button>` : ""}
        <button class="btn" id="pCancel">キャンセル</button>
        <button class="btn btn--primary" id="pSave">${isEdit ? "更新" : "配置"}</button>
      </div>
    </div>
    <div class="modal__card modal__card--anchored" id="tagCard" style="max-width:520px; width:520px">
      <div class="modal__title">タグ一覧（同期）</div>
      <div class="label">同じタグの値は、このプロジェクト内の全ページ・全要素で同期します。</div>
      <div class="field" style="margin-top:8px">
        <div class="label">検索</div>
        <input class="input" id="tagSearch" placeholder="例）氏名 / 住所 / 金額 …" />
      </div>
      <div class="badge badge--soft" style="margin-top:8px">タグ名をクリックで選択 → 値は即反映</div>
      <div class="tagPane" id="tagQuickPane" style="margin-top:10px; max-height: calc(100vh - 220px)"></div>
    </div>
    <div class="modal__card modal__card--anchored paletteGuideCard" id="paletteGuideCard" style="max-width:420px; width:min(420px, calc(100vw - 40px))">
      <div class="paletteGuideCard__title">配置のコツ</div>
      <div class="paletteGuideCard__text">タグ名と値を入力して配置しよう。タグ一覧のタグをクリックすると既存タグを呼び出せます。同じタグはまとめて値を編集できます。</div>
    </div>
  `
  const original = isEdit
    ? {
        fid: String(editFid),
        pl: { ...(state.placements?.[editFid] || {}) },
        val: String(state.values?.[String((state.placements?.[editFid] || {}).tag || "")] || ""),
      }
    : null
  let liveDirty = false
  let liveTimer = null

  const revertLive = async () => {
    if (!original) return
    try {
      const fid = original.fid
      const pl = original.pl || {}
      const x = Number(pl.x || 0)
      const y = Number(pl.y || 0)
      const page = Number(pl.page || 0)
      const tag = String(pl.tag || "").trim()
      const fontSize = Number(pl.font_size || 14) || 14
      const color = String(pl.color || "#0f172a")
      const lineH = Number(pl.line_height || 1.2) || 1.2
      const letterS = Number(pl.letter_spacing ?? DEFAULT_LETTER_SPACING) || DEFAULT_LETTER_SPACING
      const writingMode0 = String(pl.writing_mode || "horizontal")
      state.placements[fid] = { ...(state.placements?.[fid] || {}), tag, page, x, y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS }
      if (tag) state.values[tag] = String(original.val || "")
      if (window.pywebview?.api?.update_placement) {
        await window.pywebview.api.update_placement(fid, { tag, page, x, y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS, writing_mode: writingMode0 })
      } else {
        await window.pywebview.api.set_element_pos?.(fid, x, y)
      }
      if (tag) await window.pywebview.api.set_value?.(tag, String(original.val || ""))
      await showPage(page)
    } catch {}
  }

  const closeMaybe = async () => {
    if (isEdit && liveDirty) await revertLive()
    close()
  }

  $("#modalClose").onclick = closeMaybe
  $("#pCancel").onclick = closeMaybe
  const tagInput = $("#pTag")
  const valInput = $("#pVal")
  const sizeInput = $("#pSize")
  const wmHBtn = $("#pWmH")
  const wmVBtn = $("#pWmV")
  const colorInput = $("#pColor")
  const lineHInput = $("#pLineH")
  const letterSInput = $("#pLetterS")
  const pageInput = $("#pPage")
  const card = $("#paletteCard")
  const tagCard = $("#tagCard")
  const guideCard = $("#paletteGuideCard")
  const tagQuickPane = $("#tagQuickPane")
  const tagSearch = $("#tagSearch")

  const setWritingMode = (m) => {
    writingMode = (String(m) === "vertical") ? "vertical" : "horizontal"
    if (wmHBtn) wmHBtn.classList.toggle("is-selected", writingMode === "horizontal")
    if (wmVBtn) wmVBtn.classList.toggle("is-selected", writingMode === "vertical")
    // live preview: reflect into placement box sizing
    if (isEdit && editFid && state.placements?.[editFid]) {
      state.placements[editFid] = { ...(state.placements[editFid] || {}), writing_mode: writingMode }
      drawOverlay()
    }
  }
  if (wmHBtn) wmHBtn.onclick = () => setWritingMode("horizontal")
  if (wmVBtn) wmVBtn.onclick = () => setWritingMode("vertical")
  setWritingMode(writingMode)

  // 2つのパレットが重なる場合：操作している方を前面にする
  const bringToFront = (which) => {
    if (!card || !tagCard) return
    const top = 90
    const under = 89
    if (which === "tag") {
      tagCard.style.zIndex = String(top)
      card.style.zIndex = String(under)
    } else {
      card.style.zIndex = String(top)
      tagCard.style.zIndex = String(under)
    }
  }
  bringToFront("palette")
  if (card) {
    card.addEventListener("pointerdown", () => bringToFront("palette"), { passive: true })
    card.addEventListener("focusin", () => bringToFront("palette"))
  }
  if (tagCard) {
    tagCard.addEventListener("pointerdown", () => bringToFront("tag"), { passive: true })
    tagCard.addEventListener("focusin", () => bringToFront("tag"))
  }

  // 色パレット（選択式）
  const sw = $("#pSwatches")
  if (sw) {
    const colors = [
      "#0f172a",
      "#ffffff",
      "#141726",
      "#64748b",
      "#7c5cff",
      "#5ad7ff",
      "#ff6aa2",
      "#ff4d6d",
      "#22c55e",
      "#ffd36a",
      "#7cffb2",
    ]
    const norm = (s) => String(s || "").trim().toLowerCase()
    const applySelected = () => {
      const cur = norm(colorInput?.value)
      for (const el of sw.querySelectorAll(".swatch")) {
        el.classList.toggle("is-selected", norm(el.dataset.color) === cur)
      }
    }
    sw.innerHTML = ""
    colors.forEach((c) => {
      const b = document.createElement("button")
      b.type = "button"
      b.className = "swatch"
      b.dataset.color = c
      b.style.background = c
      b.onclick = () => {
        if (colorInput) {
          colorInput.value = c
          colorInput.dispatchEvent(new Event("input", { bubbles: true }))
        }
        applySelected()
      }
      sw.appendChild(b)
    })
    applySelected()
  }

  // 上下ボタンで微調整（現場では“数字→感覚”が作れないのでボタン中心に）
  const bindSpin = (inputEl, upEl, downEl, step, minV = null, maxV = null, digits = null) => {
    if (!inputEl) return
    const toNum = () => {
      const v = Number(String(inputEl.value || "").trim())
      return Number.isFinite(v) ? v : 0
    }
    const setNum = (v) => {
      let x = v
      if (typeof minV === "number") x = Math.max(minV, x)
      if (typeof maxV === "number") x = Math.min(maxV, x)
      if (typeof digits === "number") x = Number(x.toFixed(digits))
      inputEl.value = String(x)
    }
    const bump = (dir) => {
      setNum(toNum() + dir * step)
      inputEl.dispatchEvent(new Event("input", { bubbles: true }))
    }
    if (upEl) upEl.onclick = () => bump(+1)
    if (downEl) downEl.onclick = () => bump(-1)
  }
  bindSpin(sizeInput, $("#pSizeUp"), $("#pSizeDown"), 1, 6, 96, 0)
  bindSpin(lineHInput, $("#pLineHUp"), $("#pLineHDown"), 0.1, 0.6, 3.0, 1)
  bindSpin(letterSInput, $("#pLetterSUp"), $("#pLetterSDown"), 0.5, -5, 30, 1)

  // 編集中は「変更した瞬間にプレビューへ反映」する（微調整が主運用のため）
  const scheduleLive = () => {
    if (!isEdit) return
    liveDirty = true
    if (liveTimer) clearTimeout(liveTimer)
    liveTimer = setTimeout(async () => {
      try {
        const fid = String(editFid)
        const raw = (valInput?.value || "").replaceAll("\r\n", "\n")
        const val = raw.replaceAll("\n", "<br>")
        const fontSize = Number(sizeInput?.value || "14") || 14
        state.defaultFontSize = fontSize
        window.pywebview?.api?.update_admin_settings?.({ default_font_size: fontSize })
        const color = String(colorInput?.value || "#0f172a").trim() || "#0f172a"
        const lineH = Number(lineHInput?.value || "1.2") || 1.2
        const letterS = Number(letterSInput?.value || "0") || 0
        const page = Math.max(0, (Number(pageInput?.value || "1") || 1) - 1)
        const pl0 = state.placements?.[fid] || currentPl || {}
        const tag = String(tagInput?.value || pl0.tag || "").trim()
        const x = Number(pl0.x || 0)
        const y = Number(pl0.y || 0)
        state.placements[fid] = { ...(pl0 || {}), tag, page, x, y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS }
        if (tag) state.values[tag] = val
        if (window.pywebview?.api?.update_placement) {
          await window.pywebview.api.update_placement(fid, { tag, page, x, y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS })
        } else {
          await window.pywebview.api.set_element_pos?.(fid, x, y)
        }
        if (tag) await window.pywebview.api.set_value?.(tag, val)
        await showPage(page)
      } catch {}
    }, 120)
  }
  if (isEdit) {
    sizeInput?.addEventListener("input", scheduleLive)
    lineHInput?.addEventListener("input", scheduleLive)
    letterSInput?.addEventListener("input", scheduleLive)
    pageInput?.addEventListener("input", scheduleLive)
    colorInput?.addEventListener("input", scheduleLive)
    valInput?.addEventListener("input", scheduleLive)
  }

  // パレットを“要素に被らず”かつ“PDF表示領域内”に収める
  const positionPalette = () => {
    if (!card) return
    const img = $("#previewImg")
    if (!img || !img.parentElement) return
    const stage = img.parentElement.getBoundingClientRect() // PDF表示枠（白余白含む）
    const content = getRenderedContentRect(img, state.pageW, state.pageH) // 実PDF領域
    const pad = 10
    const margin = 12

    // 最大高さを枠に合わせる（はみ出し防止）
    const maxH = Math.max(240, Math.floor(stage.height - pad * 2))
    card.style.maxHeight = `${maxH}px`
    card.style.overflow = "auto"

    // anchor rect（編集時は要素サイズを推定、配置時はクリック点）
    let ax = content.left + (pt.x / state.pageW) * content.width
    let ay = content.top + (pt.y / state.pageH) * content.height
    let aw = 1
    let ah = 1
    if (isEdit && editFid) {
      const pl = state.placements?.[editFid] || {}
      const fs = Number(pl.font_size || 14) || 14
      const v = String((state.values?.[String(pl.tag || "")] || "")).replaceAll("<br>", "\n")
      const lines = v ? v.split("\n") : [String(pl.tag || "")]
      const longest = Math.max(...lines.map((s) => s.length), 1)
      const lh = Number(pl.line_height || 1.2) || 1.2
      const ls = Number(pl.letter_spacing ?? DEFAULT_LETTER_SPACING) || DEFAULT_LETTER_SPACING
      const wPage = Math.max(42, longest * (fs * 0.62 + ls))
      const hPage = Math.max(22, lines.length * fs * lh)
      ax = content.left + (Number(pl.x || pt.x) / state.pageW) * content.width
      ay = content.top + (Number(pl.y || pt.y) / state.pageH) * content.height
      aw = (wPage / state.pageW) * content.width
      ah = (hPage / state.pageH) * content.height
    }

    const rect = () => card.getBoundingClientRect()
    const cw = rect().width
    const ch = rect().height
    const el = { left: ax, top: ay, width: aw, height: ah }

    // Keep palette visible even when stage is partially outside viewport.
    const viewport = {
      left: pad,
      top: pad,
      right: Math.max(pad, window.innerWidth - pad),
      bottom: Math.max(pad, window.innerHeight - pad),
    }
    const bounds = {
      left: Math.max(stage.left + pad, viewport.left),
      top: Math.max(stage.top + pad, viewport.top),
      right: Math.min(stage.right - pad, viewport.right),
      bottom: Math.min(stage.bottom - pad, viewport.bottom),
    }
    if (bounds.right - bounds.left < 40 || bounds.bottom - bounds.top < 40) {
      bounds.left = viewport.left
      bounds.top = viewport.top
      bounds.right = viewport.right
      bounds.bottom = viewport.bottom
    }

    const fit = (l, t) =>
      l >= bounds.left &&
      t >= bounds.top &&
      l + cw <= bounds.right &&
      t + ch <= bounds.bottom

    const clamp = (l, t) => {
      const ll = Math.min(Math.max(l, bounds.left), bounds.right - cw)
      const tt = Math.min(Math.max(t, bounds.top), bounds.bottom - ch)
      return { l: ll, t: tt }
    }

    const candidates = [
      { l: el.left + el.width + margin, t: el.top }, // 右
      { l: el.left - cw - margin, t: el.top }, // 左
      { l: el.left, t: el.top + el.height + margin }, // 下
      { l: el.left, t: el.top - ch - margin }, // 上
    ]

    let pos = null
    for (const c of candidates) {
      if (fit(c.l, c.t)) {
        pos = c
        break
      }
    }
    if (!pos) pos = clamp(el.left + el.width + margin, el.top)
    else pos = clamp(pos.l, pos.t)

    // もし要素と重なりそうなら、少しずらす（最低限）
    const overlaps =
      pos.l < el.left + el.width &&
      pos.l + cw > el.left &&
      pos.t < el.top + el.height &&
      pos.t + ch > el.top
    if (overlaps) {
      const alt = clamp(el.left - cw - margin, el.top)
      pos = alt
    }

    card.style.left = `${Math.round(pos.l)}px`
    card.style.top = `${Math.round(pos.t)}px`
    // Place tagCard near paletteCard within stage (same size feeling)
    if (tagCard) {
      const r1 = card.getBoundingClientRect()
      const w2 = tagCard.getBoundingClientRect().width || 520
      const h2 = tagCard.getBoundingClientRect().height || 420
      const fit = (l, t) =>
        l >= bounds.left &&
        t >= bounds.top &&
        l + w2 <= bounds.right &&
        t + h2 <= bounds.bottom
      const clamp = (l, t) => {
        const ll = Math.min(Math.max(l, bounds.left), bounds.right - w2)
        const tt = Math.min(Math.max(t, bounds.top), bounds.bottom - h2)
        return { l: ll, t: tt }
      }
      const cands = [
        { l: r1.right + margin, t: r1.top },
        { l: r1.left - w2 - margin, t: r1.top },
        { l: r1.left, t: r1.bottom + margin },
        { l: r1.left, t: r1.top - h2 - margin },
      ]
      let p2 = null
      for (const c of cands) {
        if (fit(c.l, c.t)) {
          p2 = c
          break
        }
      }
      if (!p2) p2 = clamp(r1.right + margin, r1.top)
      else p2 = clamp(p2.l, p2.t)
      tagCard.style.left = `${Math.round(p2.l)}px`
      tagCard.style.top = `${Math.round(p2.t)}px`
    }
    // Keep guide popup at bottom-left inside visible stage area.
    if (guideCard) {
      const w3 = guideCard.getBoundingClientRect().width || 380
      const h3 = guideCard.getBoundingClientRect().height || 120
      const gLeft = Math.min(Math.max(bounds.left + 8, bounds.left), Math.max(bounds.left, bounds.right - w3 - 8))
      const gTop = Math.min(Math.max(bounds.top + 8, bounds.top), Math.max(bounds.top, bounds.bottom - h3 - 8))
      guideCard.style.left = `${Math.round(gLeft)}px`
      guideCard.style.top = `${Math.round(gTop)}px`
    }
  }
  requestAnimationFrame(() => {
    positionPalette()
    setTimeout(positionPalette, 0)
  })

  if (tagInput) tagInput.focus()
  if (isEdit) {
    try {
      tagInput.value = String((currentPl.tag || "")).trim()
      valInput.value = String((state.values?.[String(currentPl.tag || "")] || "")).replaceAll("<br>", "\n")
      pageInput.value = String((Number(currentPl.page || pageIdx) || 0) + 1)
    } catch {}
  }

  const normalizeText = (s) => String(s || "").replaceAll("<br>", "\n").trim().toLowerCase()
  const getFilteredTags = (qRaw) => {
    const q = normalizeText(qRaw)
    const tags = state.tags || []
    if (!q) return [...tags]
    return tags.filter((t) => {
      const tt = normalizeText(t)
      const vv = normalizeText(state.values?.[t] || "")
      return tt.includes(q) || vv.includes(q)
    })
  }

  // ---- Tag quick palette (edit values / select tag to place) ----
  const renderTagQuick = () => {
    if (!tagQuickPane) return
    const q = String(tagSearch?.value || "")
    const tags = getFilteredTags(q)
    const currentTag = String(tagInput?.value || "").trim()
    tagQuickPane.innerHTML = `
      <div class="badge">タグ一覧</div>
      <div class="badge badge--soft">${tags.length} 件</div>
      <div class="list" id="tagQuickList"></div>
    `
    const list = $("#tagQuickList")
    if (!list) return
    tags.forEach((t, i) => {
      const row = document.createElement("div")
      row.className = "row"
      row.style.alignItems = "center"
      row.style.gap = "10px"
      const v = String((state.values?.[t] || "")).replaceAll("<br>", "\n")
      row.innerHTML = `
        <button type="button" class="btn btn--soft tagNameBtn" data-tag="${escapeHtml(t)}">${escapeHtml(t)}</button>
        <input class="input" data-tag="${escapeHtml(t)}" placeholder="値…" value="${escapeHtml(v)}">
      `
      const name = row.querySelector(".tagNameBtn")
      const inp = row.querySelector("input")
      if (name) {
        name.onclick = () => {
          if (tagInput) tagInput.value = t
          if (valInput) valInput.value = String((state.values?.[t] || "")).replaceAll("<br>", "\n")
          // visually hint selection
          try {
            const chips = tagQuickPane.querySelectorAll("[data-tag]")
            chips.forEach((el) => el.classList.remove("is-selected"))
          } catch {}
        }
      }
      if (inp) {
        if (t === currentTag) inp.style.boxShadow = "0 0 0 5px rgba(124,92,255,.12)"
        let timer = null
        inp.addEventListener("input", () => {
          const raw = String(inp.value || "").replaceAll("\r\n", "\n")
          const val = raw.replaceAll("\n", "<br>")
          state.values[t] = val
          if (timer) clearTimeout(timer)
          timer = setTimeout(async () => {
            try {
              await window.pywebview.api.set_value(t, val)
              await showPage(state.previewPageIndex || 0)
            } catch {}
          }, 120)
        })
        inp.addEventListener("keydown", (ev) => {
          if (ev.key === "Enter" && !ev.shiftKey) {
            ev.preventDefault()
            const next = list.querySelectorAll("input")[i + 1]
            if (next) next.focus()
          }
        })
      }
      list.appendChild(row)
    })
  }
  if (tagSearch) {
    tagSearch.addEventListener("input", () => renderTagQuick())
  }
  renderTagQuick()

  const save = async () => {
    const tag = (tagInput?.value || "").trim()
    if (!tag) {
      await uiAlert("タグを入れてください")
      return
    }
    const raw = (valInput?.value || "").replaceAll("\r\n", "\n")
    const val = raw.replaceAll("\n", "<br>")
    const fontSize = Number(sizeInput?.value || "14") || 14
    state.defaultFontSize = fontSize
    window.pywebview?.api?.update_admin_settings?.({ default_font_size: fontSize })
    const color = String(colorInput?.value || "#0f172a").trim() || "#0f172a"
    const lineH = Number(lineHInput?.value || "1.2") || 1.2
    const letterS = Number(letterSInput?.value || "0") || 0
    const page = Math.max(0, (Number(pageInput?.value || "1") || 1) - 1)
    try {
      // Ensure tag exists in list (for tag pane)
      if (!state.tags.includes(tag)) state.tags.push(tag)

      let fid = isEdit ? String(editFid) : null
      if (!isEdit) {
        // Always create a new element (same tag can be placed multiple times).
        let r = await window.pywebview.api.add_text_field(tag, page, pt.x, pt.y, fontSize)
        if (!r.ok && r.error === "no_project" && state.projectPath && window.pywebview.api.load_project) {
          try {
            await window.pywebview.api.load_project(state.projectPath)
            r = await window.pywebview.api.add_text_field(tag, page, pt.x, pt.y, fontSize)
          } catch {}
        }
        if (!r.ok) {
          await uiAlert(`追加に失敗: ${r.error || "unknown"}`)
          return
        }
        fid = r.fid
        state.placements[fid] = { tag, page, x: pt.x, y: pt.y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS, writing_mode: writingMode }
        if (window.pywebview?.api?.update_placement) {
          await window.pywebview.api.update_placement(fid, { writing_mode: writingMode })
        }
      } else {
        // Update existing element
        if (!fid) {
          await uiAlert("要素IDが不明です")
          return
        }
        state.placements[fid] = { ...(state.placements[fid] || {}), tag, page, x: pt.x, y: pt.y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS, writing_mode: writingMode }
        if (window.pywebview?.api?.update_placement) {
          await window.pywebview.api.update_placement(fid, { tag, page, x: pt.x, y: pt.y, font_size: fontSize, color, line_height: lineH, letter_spacing: letterS, writing_mode: writingMode })
        } else {
          await window.pywebview.api.set_element_pos(fid, pt.x, pt.y)
        }
      }

      state.values[tag] = val
      let sv = await window.pywebview.api.set_value(tag, val)
      if (sv && sv.ok === false && sv.error === "no_project" && state.projectPath && window.pywebview.api.load_project) {
        try {
          await window.pywebview.api.load_project(state.projectPath)
          await window.pywebview.api.set_value(tag, val)
        } catch {}
      }
      state.selectKeys = fid ? [fid] : []
      state.idx = Math.max(0, state.tags.indexOf(tag))
      await window.pywebview.api.save_current_project(false)
      await showPage(page)
      render()
      close()
    } catch (e) {
      await uiAlert(`配置に失敗: ${e}`)
    }
  }
  $("#pSave").onclick = save
  const del = $("#pDelete")
  if (del) del.onclick = async () => {
    const ok = await uiConfirm(tr("dialog.deleteElement", "この要素を削除しますか？（Undoで戻せます）"))
    if (!ok) return
    const before = snapshotProject()
    const fid = String(editFid)
    delete state.placements[fid]
    state.selectKeys = state.selectKeys.filter((k) => k !== fid)
    pushUndo(before)
    if (window.pywebview?.api?.delete_elements) await window.pywebview.api.delete_elements?.([fid])
    else await window.pywebview.api.set_project_payload?.({ tags: state.tags, values: state.values, placements: state.placements })
    await window.pywebview.api.save_current_project?.(false)
    showPage(state.previewPageIndex || 0)
    render()
    close()
  }
  valInput?.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && ev.metaKey) {
      ev.preventDefault()
      save()
    }
  })
}

async function focusDesignKey(refreshPreview = true) {
  if (!state.designKey) return
  const info = await window.pywebview.api.get_element_info(state.designKey)
  if (info.ok) {
    const pos = $("#dPos")
    if (pos) pos.textContent = `x:${Math.round(info.x)} y:${Math.round(info.y)}`
    state.pageW = info.page_display_width || state.pageW
    state.pageH = info.page_display_height || state.pageH
    state.designPos = { x: info.x, y: info.y }
  }
  if (refreshPreview) await queuePreview(state.designKey)
  drawOverlay()
}

function drawOverlay() {
  const ov = $("#overlay")
  const img = $("#previewImg")
  if (!ov) return
  const ctx = ov.getContext("2d")
  const rect = ov.getBoundingClientRect()
  ov.width = Math.max(1, Math.floor(rect.width * devicePixelRatio))
  ov.height = Math.max(1, Math.floor(rect.height * devicePixelRatio))
  ctx.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0)
  ctx.clearRect(0, 0, rect.width, rect.height)

  const hasSelection = (state.selectKeys || []).length > 0
  if ((!state.designMode && !state.addMode && !hasSelection) || !img || !img.src) return

  // 画像の実描画領域（object-fit: contain の余白を除外）
  const box = getRenderedContentRect(img, state.pageW, state.pageH)
  const ox = box.left - rect.left
  const oy = box.top - rect.top
  const iw = box.width
  const ih = box.height

  // 枠
  ctx.save()
  ctx.strokeStyle = "rgba(124,92,255,.35)"
  ctx.lineWidth = 2
  ctx.strokeRect(ox + 1, oy + 1, Math.max(0, iw - 2), Math.max(0, ih - 2))
  ctx.restore()

  // 追加モードの案内
  if (state.addMode) {
    ctx.save()
    ctx.fillStyle = "rgba(15,23,42,.60)"
    ctx.strokeStyle = "rgba(255,255,255,.65)"
    ctx.lineWidth = 1
    const pad = 10
    const msg = `クリックで追加：${state.addDraftName}`
    ctx.font = "600 12px system-ui, -apple-system, Segoe UI, sans-serif"
    const tw = ctx.measureText(msg).width
    const x = ox + pad
    const y = oy + pad
    ctx.fillRect(x, y, tw + 18, 26)
    ctx.strokeRect(x, y, tw + 18, 26)
    ctx.fillStyle = "rgba(255,255,255,.92)"
    ctx.fillText(msg, x + 9, y + 17)
    ctx.restore()
    return
  }

  // 選択中要素の枠（作業者向け）
  if (hasSelection) {
    const page = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
    const selected = state.selectKeys.filter((t) => state.placements?.[t] && Number(state.placements[t].page || 0) === page)
    ctx.save()
    ctx.setLineDash([])
    for (const t of selected) {
      const pl = state.placements[t] || {}
      const b = placementBoxOnPage(t, pl)
      const x1 = (Number(pl.x || 0) / state.pageW) * iw + ox
      const y1 = (Number(pl.y || 0) / state.pageH) * ih + oy
      const w1 = (b.w / state.pageW) * iw
      const h1 = (b.h / state.pageH) * ih
      ctx.strokeStyle = "rgba(255,106,162,.95)"
      ctx.lineWidth = 2
      ctx.strokeRect(x1 - 2, y1 - 2, w1 + 4, h1 + 4)
      ctx.fillStyle = "rgba(255,106,162,.10)"
      ctx.fillRect(x1 - 2, y1 - 2, w1 + 4, h1 + 4)
      // ラベル
      ctx.font = "700 12px system-ui, -apple-system, Segoe UI, sans-serif"
      ctx.fillStyle = "rgba(15,23,42,.82)"
      ctx.fillText(String((state.placements?.[t]?.tag || t) || t), x1 + 4, y1 - 8)
    }
    ctx.restore()
  }

  // 座標（stateにキャッシュして“ヌルヌル”動かす）
  const p = state.designPos
  if (!p) return
  const x = (p.x / state.pageW) * iw + ox
  const y = (p.y / state.pageH) * ih + oy
  ctx.save()
  ctx.strokeStyle = "rgba(255,106,162,.9)"
  ctx.lineWidth = 2
  ctx.beginPath()
  ctx.moveTo(x - 12, y)
  ctx.lineTo(x + 12, y)
  ctx.moveTo(x, y - 12)
  ctx.lineTo(x, y + 12)
  ctx.stroke()
  ctx.fillStyle = "rgba(255,106,162,.10)"
  ctx.beginPath()
  ctx.arc(x, y, 12, 0, Math.PI * 2)
  ctx.fill()

  // ガイド線（統括が“気持ちよく揃えられる”）
  ctx.strokeStyle = "rgba(255,255,255,.55)"
  ctx.lineWidth = 1
  ctx.setLineDash([4, 6])
  ctx.beginPath()
  ctx.moveTo(ox, y)
  ctx.lineTo(ox + iw, y)
  ctx.moveTo(x, oy)
  ctx.lineTo(x, oy + ih)
  ctx.stroke()
  ctx.restore()
}

function enableOverlayPointer(on) {
  const ov = $("#overlay")
  if (!ov) return
  ov.classList.toggle("is-active", !!on)
  ov.style.pointerEvents = on ? "auto" : "none"
}

async function boot() {
  bindViewportMetricsOnce()
  try {
    await window.i18n?.ready
    const fromQuery = getLocaleFromQuery()
    if (fromQuery) {
      state.locale = window.i18n?.setLocale?.(fromQuery) || fromQuery
    } else {
      state.locale = getLocaleSafe()
      syncLocaleQuery(state.locale)
    }
  } catch {}
  try {
    await ensureAdSenseScript()
  } catch {}
  try {
    await loadWorkers()
  } catch (e) {
    // If worker fetch fails, still show the gate (so user can retry/relaunch).
    console.error("loadWorkers failed:", e)
    state.workers = []
    state.workerId = null
    state.gate = state.gate || { error: "" }
    state.gate.error = "起動に失敗しました（作業者一覧の取得）。アプリを再起動してください。"
  }
  try {
    const r = await window.pywebview.api.get_admin_settings?.()
    const s = r?.settings && typeof r.settings === "object" ? r.settings : {}
    state.defaultFontSize = Number(s.default_font_size || 14) || 14
    state.viewZoom = Number(s.view_zoom || 1.0) || 1.0
  } catch {
    state.defaultFontSize = 14
    state.viewZoom = 1.0
  }
  // 起動時は必ずゲート画面（PDF/プロジェクトの読み込み選択）へ
  state.appStage = "gate"
  state.gate = state.gate || { error: "" }
  state.uiMode = "admin"
  render()
}

let _booted = false
async function bootOnce() {
  if (_booted) return
  // In desktop(pywebview), DOMContentLoaded can fire before window.pywebview.api is injected.
  // If we boot too early, we crash and never boot again (because _booted becomes true).
  if (!window.pywebview || !window.pywebview.api) {
    bootOnce.__tries = (bootOnce.__tries || 0) + 1
    // Retry briefly; pywebviewready will also fire.
    if (bootOnce.__tries < 200) setTimeout(bootOnce, 50)
    return
  }
  try {
    await boot()
    _booted = true
  } catch (e) {
    console.error("boot failed:", e)
    _booted = false
    // Retry once API is ready; do not lock into blank screen.
    bootOnce.__tries = (bootOnce.__tries || 0) + 1
    if (bootOnce.__tries < 260) setTimeout(bootOnce, 200)
  }
}

// Desktop (pywebview) emits this event. Web demo (Pages) does not.
window.addEventListener("pywebviewready", bootOnce)
// Web demo entrypoint
window.addEventListener("DOMContentLoaded", bootOnce)
if (document.readyState !== "loading") bootOnce()

