const $ = (sel) => document.querySelector(sel)

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
  tags: [],
  idx: 0,
  values: {},
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
  showTagPane: true,
  showPanel: true,
  pageLocked: false,
}

state.history = loadLocal("inputstudio-history", [])
state.lastSession = loadLocal("inputstudio-last-session", null)
state.lastProjectDir = loadLocal("inputstudio-last-dir", null)
state.showTagPane = loadLocal("inputstudio-show-tagpane", true)
state.showPanel = loadLocal("inputstudio-show-panel", true)

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

function toast(msg) {
  const el = $("#toast")
  el.textContent = msg
  el.style.display = "block"
  clearTimeout(el._t)
  el._t = setTimeout(() => (el.style.display = "none"), 2100)
}

function fmtTime(sec) {
  const s = Math.max(0, Math.floor(sec))
  const h = String(Math.floor(s / 3600)).padStart(2, "0")
  const m = String(Math.floor((s % 3600) / 60)).padStart(2, "0")
  const ss = String(s % 60).padStart(2, "0")
  return `${h}:${m}:${ss}`
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
  const r = await api.get_preview_png_base64_page(idx)
  if (my !== pageReq) return
  if (r && r.ok) {
    const img = $("#previewImg")
    if (img) {
      img.style.visibility = "visible"
      img.onerror = () => (img.style.visibility = "hidden")
      img.src = r.png
    }
    state.pageW = r.page_display_width || state.pageW
    state.pageH = r.page_display_height || state.pageH
    drawOverlay()
    const p = $("#pageIndicator")
    if (p) p.textContent = `${idx + 1} / ${state.pageCount || 1}`
  } else {
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

function renderTagPane() {
  const pane = $("#tagPane")
  if (!pane) return
  if (!state.tags.length) {
    pane.innerHTML = ``
    return
  }
  const items = state.tags
  pane.innerHTML = `
    <div class="badge">タグ一覧（Enterで次へ）</div>
    <div class="badge badge--soft">${items.length} 件</div>
    <div class="row" style="gap:8px; margin:6px 0">
      <button class="btn btn--soft" id="btnTagExport">タグを書き出す</button>
      <button class="btn btn--soft" id="btnTagImport">タグを読み込む</button>
      <input type="file" id="tagImportFile" accept=".json,.txt" style="display:none" />
    </div>
    <div class="list" id="tagList"></div>
  `
  const list = $("#tagList")
  items.forEach((t, i) => {
    const row = document.createElement("div")
    row.className = "row"
    row.innerHTML = `
      <div class="minw120"><label>${escapeHtml(t)}</label></div>
      <input data-tag-edit="${escapeHtml(t)}" placeholder="値を入力…">
    `
    const inp = row.querySelector("input")
    const current = (state.values[t] || "").replaceAll("<br>", "\n")
    inp.value = current
    inp.addEventListener("keydown", async (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault()
        const inputs = Array.from(list.querySelectorAll("input[data-tag-edit]"))
        const idx = inputs.indexOf(inp)
        if (inputs[idx + 1]) inputs[idx + 1].focus()
      }
    })
    inp.addEventListener("input", () => {
      const inputs = Array.from(list.querySelectorAll("input[data-tag-edit]"))
      const idx = inputs.indexOf(inp)
      updateTagValue(t, inp.value)
    })
    list.appendChild(row)
  })

  const tagImportFile = $("#tagImportFile")
  const btnTagExport = $("#btnTagExport")
  const btnTagImport = $("#btnTagImport")
  if (btnTagExport) {
    btnTagExport.onclick = () => {
      const payload = { tags: state.tags, values: state.values }
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" })
      const a = document.createElement("a")
      a.href = URL.createObjectURL(blob)
      a.download = `inputstudio-tags-${Date.now()}.json`
      a.click()
      URL.revokeObjectURL(a.href)
      toast("タグをJSONで書き出しました")
    }
  }
  if (btnTagImport && tagImportFile) {
    btnTagImport.onclick = () => tagImportFile.click()
    tagImportFile.onchange = async (e) => {
      const file = e.target.files?.[0]
      if (!file) return
      const text = await file.text().catch(() => null)
      if (!text) return toast("読み込みに失敗しました")
      try {
        const data = JSON.parse(text)
        const tags = Array.isArray(data.tags) ? data.tags.map(String) : []
        const values = typeof data.values === "object" && data.values ? data.values : {}
        if (!tags.length) return toast("タグが見つかりませんでした")
        state.tags = Array.from(new Set([...state.tags, ...tags]))
        state.values = { ...state.values, ...values }
        toast(`タグを読み込みました (${tags.length}件)`)
        render()
      } catch {
        toast("JSONが不正です")
      }
    }
  }
}

function render() {
  const total = state.tags.length || 0
  const done = total ? filledCount() : 0
  const idx = state.idx
  const key = total ? state.tags[idx] : null
  const hasTags = !!key
  const val = key ? (state.values[key] || "") : ""
  const valText = val.replaceAll("<br>", "\n")
  const progress = total ? Math.round(((idx + 1) / total) * 100) : 0

  const isAdmin = state.uiMode === "admin"

  const modeChip = isAdmin
    ? `<button class="chip" id="btnLock">入力者モードにする</button>`
    : `<button class="chip" id="btnAdmin">管理者</button>`

  const left = `
    <div class="top">
      <div class="brand">
        <div class="logo" aria-hidden="true"></div>
        <div>
          <div class="brand__name">Input Studio</div>
        </div>
      </div>

      <div class="guide">
        <div class="guide__row">
          <span class="guide__step">1</span>
          <div style="flex:1">初めて：PDFを選んで新規プロジェクトを作成（フォーム検出なし。自分で欄を配置）</div>
          <button class="chip chip--soft" id="btnOpenPdf">PDFから新規</button>
        </div>
        <div class="guide__row">
          <span class="guide__step">2</span>
          <div style="flex:1">続きから：既存の案件（プロジェクト）を開いて編集/入力を再開</div>
          <button class="chip" id="btnOpen">案件を開く ${tipIcon(1, "PDF付きの案件ファイルを選択して開始します。")} </button>
          ${state.lastSession ? `<button class="chip chip--soft" id="btnResume">続きから</button>` : ""}
        </div>
        ${state.lastProjectDir ? `<div style="font-size:12px; color:var(--muted)">前回開いたフォルダ: ${escapeHtml(state.lastProjectDir)}</div>` : ""}
      </div>

      <div class="miniActions">
        ${isAdmin ? `<button class="chip" id="btnDesign">設計（統括）</button>` : ""}
        <button class="chip" id="btnSave">上書き保存</button>
        ${!isAdmin ? `<button class="chip chip--soft" id="btnMyHistory">自分の履歴</button>` : ""}
        ${modeChip}
        ${isAdmin ? `<button class="chip chip--soft" id="btnHistoryExport">履歴CSV</button>` : ""}
        ${isAdmin ? `<button class="chip chip--soft" id="btnHistoryReset">履歴リセット</button>` : ""}
      </div>

      <div class="glassBox">
        <div class="row spread">
          <div class="badge">${state.projectName ? `案件：${escapeHtml(state.projectName)}` : "案件未選択"}</div>
          <div class="badge">${done}/${total || 0} 完了</div>
        </div>

        <div class="row" style="margin-top:10px">
          <div class="field">
            <div class="label">作業者 ${tipIcon(2, "作業者を選ぶと、その人の入力・進捗でPDFが更新されます。")}</div>
            <select id="workerSelect">
              ${state.workers
                .map((w) => `<option value="${w.id}" ${w.id === state.workerId ? "selected" : ""}>${escapeHtml(w.name)}</option>`)
                .join("")}
            </select>
          </div>
          ${isAdmin ? `<button class="btn btn--soft" id="btnWorker">編集</button>` : ""}
        </div>

        <div class="row" style="margin-top:10px">
          <div class="field" style="flex:1">
            <div class="label">説明 / 注意 ${tipIcon(3, "作業の注意・締切・補足をここに記載。入力者はここを見ながら進めます。")}</div>
            <textarea id="infoBox" placeholder="作業の注意・締切・進め方などを記入" rows="3"></textarea>
          </div>
        </div>

        <div class="row spread" style="margin-top:12px">
          <div class="status">
            <div class="status__pill ${state.working ? (state.inPrivate ? "is-private" : "is-working") : ""}">
              ${state.working ? (state.inPrivate ? "私用中" : "作業中") : "待機中"}
            </div>
            <div class="status__time">${fmtTime(calcNetSeconds())}</div>
          </div>
          <div class="ring" style="--p:${progress}">
            <div class="ring__inner">${progress}%</div>
          </div>
        </div>

        <div class="row" style="margin-top:12px">
          <button class="btn btn--primary" id="btnStart" ${state.working ? "disabled" : ""}>開始</button>
          <button class="btn btn--tint" id="btnPrivate" ${state.working ? "" : "disabled"}>
            ${state.inPrivate ? "作業に戻る" : "私用"}
          </button>
          <button class="btn btn--danger" id="btnFinish" ${state.working ? "" : "disabled"}>終了</button>
        </div>
      </div>
    </div>

    ${
      hasTags
        ? `<div class="focus ${state.justCompleted ? "pop" : ""}">
            <div class="focus__head">
              <div class="focus__kicker">いま入力する</div>
              <div class="focus__title">${escapeHtml(key)}</div>
              <div class="focus__meta">${total ? `${idx + 1}/${total}` : ""}　・　Enterで次へ / Shift+Enterで改行</div>
            </div>

            <div class="focus__body">
              <textarea class="input textarea focus__input" id="val" placeholder="ここに入力…">${escapeHtml(valText)}</textarea>
              <div class="row spread" style="margin-top:10px">
                <div class="row">
                  <button class="btn btn--soft" id="btnPrev" ${idx <= 0 ? "disabled" : ""}>戻る</button>
                  <button class="btn btn--primary" id="btnNext" ${idx >= total - 1 ? "disabled" : ""}>次へ</button>
                  <button class="btn btn--soft" id="btnNextEmpty">未入力へ</button>
                  ${!isAdmin ? `<button class="btn btn--tint" id="btnAddField">欄を追加</button>` : ""}
                </div>
                <button class="btn btn--ghost" id="btnClear">クリア</button>
              </div>
            </div>
          </div>`
        : ""
    }
  `

  const right = `
    <div class="previewTop">
      <div class="badge">ライブプレビュー ${tipIcon(4, "入力した値がPDF画像に重ねて表示されます。ページ切替で別ページも確認できます。")}</div>
      <div class="badge badge--soft">入力が反映されます</div>
      <div class="row" style="margin-left:auto; gap:8px; flex:0 0 auto">
        <button class="btn btn--soft" id="btnPrevPage" ${state.projectPath ? "" : "disabled"}>前</button>
        <span class="badge" id="pageIndicator">${(state.previewPageIndex || 0) + 1} / ${state.pageCount || 1}</span>
        <button class="btn btn--soft" id="btnNextPage" ${state.projectPath ? "" : "disabled"}>次</button>
        <button class="btn btn--soft" id="btnToggleTagPane">${state.showTagPane ? "タグ欄:ON" : "タグ欄:OFF"}</button>
        <button class="btn btn--soft" id="btnTogglePanel">${state.showPanel ? "操作欄:ON" : "操作欄:OFF"}</button>
      </div>
    </div>
    ${
      state.projectPath
        ? `<div class="previewImg">
            <img id="previewImg" alt="preview" />
            <canvas id="confetti" class="confetti" aria-hidden="true"></canvas>
            <canvas id="overlay" class="overlay"></canvas>
            ${
              !hasTags
                ? `<div class="emptyHint">
                    <div class="emptyHint__title">まずはPDFに欄（タグ）を置きましょう</div>
                    <div class="emptyHint__text">PDF上をダブルクリック → タグ/値/サイズを入力して配置できます。</div>
                    <div class="emptyHint__actions">
                      <button class="btn btn--primary" id="btnAddFromCenter">中央に欄を追加</button>
                      <button class="btn btn--soft" id="btnHideHints">この案内を隠す</button>
                    </div>
                  </div>`
                : ""
            }
          </div>`
        : `<div class="previewPlaceholder">案件を開くとPDFがここに表示されます</div>`
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
      <div class="stage ${state.showTagPane ? "stage--split" : "stage--nosplit"}">
        ${right}
        <div class="tagPane ${state.showTagPane ? "" : "is-hidden"}" id="tagPane"></div>
      </div>
    </div>
    <div class="toast" id="toast"></div>
    <div class="modal" id="modal" style="display:none"></div>
  `

  renderTagPane()
  bind()
  queuePreview()
  tickTimerOnce()
  if (state.justCompleted) {
    burstConfetti()
    state.justCompleted = false
  }
}

function bind() {
  $("#btnOpen").onclick = async () => {
    const r = await window.pywebview.api.pick_project()
    if (!r.ok) return
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
    state.pageCount = loaded.page_count || 1
    if (!state.tags.length) {
      state.showTagPane = false
      saveLocal("inputstudio-show-tagpane", state.showTagPane)
    }
    state.idx = 0
    state.dropDir = loaded.drop_dir || ""
    state.uiMode = loaded.ui_mode || state.uiMode
    state.lastSession = { path: r.path, workerId: state.workerId, projectName: state.projectName }
    saveLocal("inputstudio-last-session", state.lastSession)
    state.working = false
    state.inPrivate = false
    state.timerStart = null
    state.privateTotal = 0
    toast("案件を読み込みました")
    render()
  }

  const btnOpenPdf = $("#btnOpenPdf")
  if (btnOpenPdf) btnOpenPdf.onclick = async () => {
    try {
      const api = window.pywebview?.api
      const pick = api?.pick_pdf
      const createSimple = api?.create_project_from_pdf_simple
      if (!pick || !createSimple) {
        alert("PDFから新規作成する機能が見つかりません。最新版またはバックエンドの create_project_from_pdf_simple/pick_pdf をご用意ください。")
        return
      }
      const r = await pick()
      if (!r?.ok) return
      toast("PDFを読み込み、新規プロジェクトを作成します…")
      const g = await createSimple(r.path)
      if (!g?.ok || !g.path) {
        return alert((g?.errors || ["PDFをプロジェクト化できませんでした"]).join("\n"))
      }
      const loaded = await api.load_project(g.path)
      if (!loaded?.ok) {
        return alert("新規プロジェクトを開けませんでした")
      }
      state.projectPath = g.path
      state.projectName = loaded.project
      state.tags = loaded.tags || []
      state.values = loaded.values || {}
      state.pageCount = loaded.page_count || 1
      if (!state.tags.length) {
        state.showTagPane = false
        saveLocal("inputstudio-show-tagpane", state.showTagPane)
      }
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
      toast("PDFから新規プロジェクトを作成しました。必要に応じてタグを配置してください。")
      render()
    } catch (e) {
      alert(`PDFから新規作成に失敗しました: ${e}`)
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
      state.pageCount = loaded.page_count || 1
      if (!state.tags.length) {
        state.showTagPane = false
        saveLocal("inputstudio-show-tagpane", state.showTagPane)
      }
      state.idx = 0
      state.dropDir = loaded.drop_dir || ""
      if (state.lastSession.workerId) state.workerId = state.lastSession.workerId
      state.working = false
      state.inPrivate = false
      state.timerStart = null
      state.privateTotal = 0
      toast("前回の案件を読み込みました")
      render()
    }
  }

  // フォーム付きPDFを扱わない前提のため、自動作成機能は削除

  const btnDesign = $("#btnDesign")
  if (btnDesign) btnDesign.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    state.designMode = true
    state.designKey = state.designKey || (state.tags[0] || null)
    await openDesignModal()
  }

  const btnSave = $("#btnSave")
  if (btnSave) btnSave.onclick = async () => {
    if (!state.projectPath) return toast("先に案件を開いてください")
    try {
      await window.pywebview.api.save_current_project()
      state.lastSession = { path: state.projectPath, workerId: state.workerId, projectName: state.projectName }
      saveLocal("inputstudio-last-session", state.lastSession)
      toast("案件を保存しました")
    } catch (e) {
      toast(`保存に失敗しました: ${e}`)
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
  const btnToggleTagPane = $("#btnToggleTagPane")
  if (btnToggleTagPane) btnToggleTagPane.onclick = () => {
    state.showTagPane = !state.showTagPane
    saveLocal("inputstudio-show-tagpane", state.showTagPane)
    render()
  }

  const btnTogglePanel = $("#btnTogglePanel")
  if (btnTogglePanel) btnTogglePanel.onclick = () => {
    state.showPanel = !state.showPanel
    saveLocal("inputstudio-show-panel", state.showPanel)
    render()
  }

  const btnAddFromCenter = $("#btnAddFromCenter")
  if (btnAddFromCenter) btnAddFromCenter.onclick = () => {
    const img = $("#previewImg")
    if (!img || !img.src) return
    const x = Math.round(0.5 * state.pageW)
    const y = Math.round(0.5 * state.pageH)
    if (!x || !y) return
    openPlacePalette({ x, y })
  }

  const btnHideHints = $("#btnHideHints")
  if (btnHideHints) btnHideHints.onclick = () => {
    const h = document.querySelector(".emptyHint")
    if (h) h.remove()
  }

  const btnAdmin = $("#btnAdmin")
  if (btnAdmin) btnAdmin.onclick = async () => {
    const ok = confirm("管理者モードに切り替えます（OCR/設計が表示されます）。よろしいですか？")
    if (!ok) return
    const r = await window.pywebview.api.set_ui_mode("admin")
    if (!r.ok) return toast("切り替えに失敗しました")
    state.uiMode = "admin"
    toast("管理者モード")
    render()
  }

  const btnLock = $("#btnLock")
  if (btnLock) btnLock.onclick = async () => {
    const ok = confirm("入力者モードに切り替えます（OCR/設計を隠します）。よろしいですか？")
    if (!ok) return
    const r = await window.pywebview.api.set_ui_mode("worker")
    if (!r.ok) return toast("切り替えに失敗しました")
    state.uiMode = "worker"
    state.designMode = false
    state.addMode = false
    toast("入力者モード")
    render()
  }

  $("#workerSelect").onchange = (e) => {
    state.workerId = e.target.value
  }

  const btnWorker = $("#btnWorker")
  if (btnWorker) btnWorker.onclick = () => openWorkerModal()

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
  if (btnHistoryReset) btnHistoryReset.onclick = () => {
    const ok = confirm("作業履歴をリセットします（内部保存のみ削除、プロジェクトは残ります）。よろしいですか？")
    if (!ok) return
    state.history = []
    saveLocal("inputstudio-history", state.history)
    toast("履歴をリセットしました")
  }

  $("#btnStart").onclick = async () => {
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
    toast("スタート！")
    render()
  }

  $("#btnPrivate").onclick = async () => {
    const r = await window.pywebview.api.toggle_private()
    if (!r.ok) return
    if (!state.inPrivate) {
      state.inPrivate = true
      state._privateStart = Date.now()
      toast("私用を開始")
    } else {
      state.inPrivate = false
      state.privateTotal += (Date.now() - state._privateStart) / 1000
      toast("作業に戻りました")
      pulse()
    }
    render()
  }

  $("#btnFinish").onclick = async () => {
    const ok = confirm("勤務を終了して提出物（ZIP）を作成します。よろしいですか？")
    if (!ok) return
    await pushValue()
    toast("提出物を作成中…")
    const r = await window.pywebview.api.finish()
    if (!r.ok) return toast("提出物の作成に失敗しました")
    state.working = false
    state.justCompleted = true
    const end = new Date().toISOString()
    const duration = calcNetSeconds()
    state.history = [
      ...state.history,
      {
        projectName: state.projectName,
        projectPath: state.projectPath,
        workerId: state.workerId,
        workerName: (state.workers.find((w) => w.id === state.workerId) || {}).name || "",
        start: state.sessionStart,
        end,
        duration,
      },
    ]
    saveLocal("inputstudio-history", state.history)
    state.sessionStart = null
    state.timerStart = null
    state.privateTotal = 0
    state.inPrivate = false
    alert(`提出物を作成しました。\n\nフォルダ: ${r.dir}\nZIP: ${r.zip}`)
    render()
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
      const name = prompt("追加する欄の名前（例：備考2 / メモ / 追記）", "") || ""
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
    // PDFをダブルクリック -> 配置パレット（作業者でも使える）
    ov.ondblclick = (ev) => {
      if (!state.projectPath) return
      // design mode のダブルクリックは既存の処理に任せる
      if (state.designMode) return
      const img = $("#previewImg")
      if (!img || !img.src) return
      const imgRect = img.getBoundingClientRect()
      const x0 = ev.clientX - imgRect.left
      const y0 = ev.clientY - imgRect.top
      if (x0 < 0 || y0 < 0 || x0 > imgRect.width || y0 > imgRect.height) return
      const x = (x0 / imgRect.width) * state.pageW
      const y = (y0 / imgRect.height) * state.pageH
      ev.preventDefault()
      openPlacePalette({ x, y })
    }

    ov.onclick = async (ev) => {
      if (!state.addMode) return
      const img = $("#previewImg")
      if (!img || !img.src) return
      const imgRect = img.getBoundingClientRect()
      const x0 = ev.clientX - imgRect.left
      const y0 = ev.clientY - imgRect.top
      if (x0 < 0 || y0 < 0 || x0 > imgRect.width || y0 > imgRect.height) return
      const x = (x0 / imgRect.width) * state.pageW
      const y = (y0 / imgRect.height) * state.pageH
      toast("欄を追加中…")
      const r = await window.pywebview.api.add_text_field(state.addDraftName, state.previewPageIndex || 0, x, y, 14)
      if (!r.ok) {
        state.addMode = false
        enableOverlayPointer(false)
        drawOverlay()
        return alert(`追加に失敗: ${r.error || "unknown"}`)
      }
      const tag = r.tag
      state.values[tag] = ""
      if (!state.tags.includes(tag)) state.tags.push(tag)
      state.idx = state.tags.indexOf(tag)
      await window.pywebview.api.save_current_project()
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
  const r = await window.pywebview.api.get_preview_png_base64(k)
  if (my !== previewReq) return
  if (r.ok) {
    const img = $("#previewImg")
    if (img) {
      img.style.visibility = "visible"
      img.onerror = () => {
        img.style.visibility = "hidden"
      }
      img.src = r.png
    }
    state.pageW = r.page_display_width || state.pageW
    state.pageH = r.page_display_height || state.pageH
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
  }
}

async function loadWorkers() {
  const r = await window.pywebview.api.get_workers()
  if (!r.ok) return
  state.workers = r.workers || []
  state.workerId = r.last_worker_id || (state.workers[0] ? state.workers[0].id : null)
}

function tickTimerOnce() {
  const el = $(".status__time")
  if (el) el.textContent = fmtTime(calcNetSeconds())
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

function openWorkerModal() {
  const modal = $("#modal")
  const current = state.workers.find((w) => w.id === state.workerId) || {}
  modal.style.display = "block"
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card">
      <div class="modal__title">作業者</div>
      <div class="field">
        <div class="label">名前</div>
        <input class="input" id="mName" value="${escapeHtml(current.name || "")}">
      </div>
      <div class="field">
        <div class="label">振込先</div>
        <input class="input" id="mBank" value="${escapeHtml(current.bank || "")}">
      </div>
      <div class="field">
        <div class="label">時給（円）</div>
        <input class="input" id="mHourly" type="number" value="${escapeHtml(current.hourly_yen || 0)}">
      </div>
      <div class="row spread" style="margin-top:14px">
        <button class="btn btn--soft" id="modalCancel">閉じる</button>
        <button class="btn btn--primary" id="modalSave">保存</button>
      </div>
    </div>
  `
  $("#modalClose").onclick = () => (modal.style.display = "none")
  $("#modalCancel").onclick = () => (modal.style.display = "none")
  $("#modalSave").onclick = async () => {
    const w = {
      id: current.id || null,
      name: $("#mName").value.trim(),
      bank: $("#mBank").value.trim(),
      hourly_yen: Number($("#mHourly").value || "0") || 0,
    }
    if (!w.name) return toast("名前を入れてください")
    const r = await window.pywebview.api.upsert_worker(w)
    if (!r.ok) return toast("保存できませんでした")
    await loadWorkers()
    state.workerId = r.id
    modal.style.display = "none"
    pulse()
    toast("保存しました")
    render()
  }
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
          <div class="label">対象タグ</div>
          <select id="dKey">
            ${state.tags.map((k) => `<option value="${k}" ${k === state.designKey ? "selected" : ""}>${escapeHtml(k)}</option>`).join("")}
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
    const i = Math.max(0, state.tags.indexOf(state.designKey) - 1)
    state.designKey = state.tags[i] || state.designKey
    $("#dKey").value = state.designKey
    await focusDesignKey()
  }
  $("#dNext").onclick = async () => {
    const i = Math.min(state.tags.length - 1, state.tags.indexOf(state.designKey) + 1)
    state.designKey = state.tags[i] || state.designKey
    $("#dKey").value = state.designKey
    await focusDesignKey()
  }
  $("#dFocus").onclick = async () => {
    await focusDesignKey()
  }
  $("#dSave").onclick = async () => {
    const r = await window.pywebview.api.save_current_project()
    if (!r.ok) return alert(`保存に失敗: ${r.error || "unknown"}`)
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
    sel.innerHTML = state.tags
      .filter((k) => (q ? k.toLowerCase().includes(q) : true))
      .map((k) => `<option value="${k}" ${k === state.designKey ? "selected" : ""}>${escapeHtml(k)}</option>`)
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
      const imgRect = img.getBoundingClientRect()
      const x0 = ev.clientX - imgRect.left
      const y0 = ev.clientY - imgRect.top
      if (x0 < 0 || y0 < 0 || x0 > imgRect.width || y0 > imgRect.height) return null
      const x = (x0 / imgRect.width) * state.pageW
      const y = (y0 / imgRect.height) * state.pageH
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

function openPlacePalette(pt) {
  const modal = $("#modal")
  if (!modal) return
  const close = () => {
    modal.style.display = "none"
    modal.innerHTML = ""
  }
  const pageIdx = Number.isFinite(state.previewPageIndex) ? state.previewPageIndex : 0
  const tagsOptions = state.tags.map((t) => `<option value="${escapeHtml(t)}"></option>`).join("")
  modal.style.display = "block"
  modal.innerHTML = `
    <div class="modal__backdrop" id="modalClose"></div>
    <div class="modal__card" style="max-width:520px">
      <div class="modal__title">PDFに配置（ダブルクリック）</div>
      <div class="label">タグを選ぶ or 新規作成し、値とサイズを指定して配置します。</div>
      <div class="field" style="margin-top:8px">
        <div class="label">タグ（既存 or 新規）</div>
        <input class="input" list="paletteTags" id="pTag" placeholder="例）氏名 / 住所" />
        <datalist id="paletteTags">${tagsOptions}</datalist>
      </div>
      <div class="field">
        <div class="label">値</div>
        <textarea class="textarea" id="pVal" rows="3" placeholder="ここに値を入力（Enterで改行）"></textarea>
      </div>
      <div class="row">
        <div class="field" style="width:120px">
          <div class="label">サイズ</div>
          <input class="input" id="pSize" inputmode="numeric" value="14" />
        </div>
        <div class="field" style="width:120px">
          <div class="label">ページ</div>
          <input class="input" id="pPage" inputmode="numeric" value="${pageIdx + 1}" />
        </div>
        <div class="field" style="flex:1">
          <div class="label">座標 (x,y)</div>
          <input class="input" id="pPos" value="${Math.round(pt.x)}, ${Math.round(pt.y)}" disabled />
        </div>
      </div>
      <div class="row" style="margin-top:10px; justify-content:flex-end">
        <button class="btn" id="pCancel">キャンセル</button>
        <button class="btn btn--primary" id="pSave">配置</button>
      </div>
    </div>
  `
  $("#modalClose").onclick = close
  $("#pCancel").onclick = close
  const tagInput = $("#pTag")
  const valInput = $("#pVal")
  const sizeInput = $("#pSize")
  const pageInput = $("#pPage")
  if (tagInput) tagInput.focus()

  const save = async () => {
    const tag = (tagInput?.value || "").trim()
    if (!tag) return alert("タグを入れてください")
    const raw = (valInput?.value || "").replaceAll("\r\n", "\n")
    const val = raw.replaceAll("\n", "<br>")
    const fontSize = Number(sizeInput?.value || "14") || 14
    const page = Math.max(0, (Number(pageInput?.value || "1") || 1) - 1)
    try {
      let tagId = tag
      const exists = state.tags.includes(tag)
      if (!exists) {
        const r = await window.pywebview.api.add_text_field(tag, page, pt.x, pt.y, fontSize)
        if (!r.ok) return alert(`追加に失敗: ${r.error || "unknown"}`)
        tagId = r.tag || tag
        if (!state.tags.includes(tagId)) state.tags.push(tagId)
      } else {
        await window.pywebview.api.set_element_pos(tag, pt.x, pt.y)
      }
      state.values[tagId] = val
      await window.pywebview.api.set_value(tagId, val)
      state.designKey = tagId
      state.idx = Math.max(0, state.tags.indexOf(tagId))
      await window.pywebview.api.save_current_project()
      await queuePreview(tagId)
      render()
      close()
    } catch (e) {
      alert(`配置に失敗: ${e}`)
    }
  }
  $("#pSave").onclick = save
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

  if ((!state.designMode && !state.addMode) || !img || !img.src) return

  // 画像が表示されている矩形（centered/letterbox を想定）
  const imgRect = img.getBoundingClientRect()
  const ox = imgRect.left - rect.left
  const oy = imgRect.top - rect.top
  const iw = imgRect.width
  const ih = imgRect.height

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
  await loadWorkers()
  try {
    const s = await window.pywebview.api.get_admin_settings()
    if (s && s.ok && s.settings) {
      state.uiMode = s.settings.ui_mode || state.uiMode
    }
  } catch (e) {}
  render()
}

window.addEventListener("pywebviewready", boot)

