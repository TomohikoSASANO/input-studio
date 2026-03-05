/**
 * Web API wrapper for Input Studio.
 * Replaces window.pywebview.api for web version.
 */

class WebAPI {
  constructor(baseUrl = '') {
    this.baseUrl = baseUrl;
    this.sessionId = null;
    this.projectId = null;
  }

  async _extractApiError(response, fallbackMessage) {
    let payload = null;
    try {
      payload = await response.json();
    } catch (_) {
      payload = null;
    }
    if (payload && typeof payload === 'object') {
      const code = payload.code || payload.error_code || null;
      const detail = payload.detail;
      if (Array.isArray(detail)) {
        const msg = detail.map((x) => (typeof x === 'object' && x?.msg) ? x.msg : String(x)).join('; ');
        return { code, message: msg || fallbackMessage };
      }
      if (typeof detail === 'string') {
        return { code, message: detail || fallbackMessage };
      }
      if (detail && typeof detail === 'object') {
        return { code: detail.code || code || null, message: detail.detail || fallbackMessage };
      }
    }
    return { code: null, message: fallbackMessage };
  }

  async init() {
    if (!this.sessionId) {
      const response = await fetch(`${this.baseUrl}/api/session`, {
        method: 'POST',
      });
      const data = await response.json();
      this.sessionId = data.session_id;
    }
    return { ok: true, session_id: this.sessionId };
  }

  async pick_pdf() {
    // In web version, trigger file input and return file path for compatibility.
    // app.js will then call create_project_from_pdf_simple(r.path) - we use _pendingPdfFile.
    const self = this;
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf';
      input.onchange = async function(e) {
        const file = e.target.files[0];
        if (file) {
          self._pendingPdfFile = file;
          resolve({ ok: true, path: file.name });
        } else {
          resolve({ ok: false });
        }
      };
      input.click();
    });
  }

  async pick_project(opts) {
    const self = this;
    const zipOnly = opts && opts.zipOnly === true;
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = zipOnly ? '.zip,application/zip' : '.zip,.json,.pdf,application/zip,application/json,application/pdf';
      input.multiple = !zipOnly;
      input.onchange = async function(e) {
        const files = Array.from(e.target.files || []);
        const zipFile = files.find((f) => f.name.toLowerCase().endsWith('.zip'));
        const jsonFile = zipOnly ? null : files.find((f) => f.name.toLowerCase().endsWith('.json'));
        if (zipFile) {
          try {
            const initResult = await self.init();
            if (!initResult?.session_id) {
              resolve({ ok: false, error: 'セッションの取得に失敗しました' });
              return;
            }
            const formData = new FormData();
            formData.append('zip_file', zipFile);
            const url = `${self.baseUrl}/api/upload-project-zip?session_id=${encodeURIComponent(self.sessionId)}`;
            const response = await fetch(url, { method: 'POST', body: formData });
            if (!response.ok) {
              let apiErr = await self._extractApiError(response, 'ZIPのアップロードに失敗しました');
              if (response.status === 405) apiErr = { code: 'METHOD_NOT_ALLOWED', message: 'サーバー接続エラー（Method Not Allowed）。サーバーを再起動してください。' };
              resolve({ ok: false, code: apiErr.code || null, error: apiErr.message });
              return;
            }
            const data = await response.json();
            if (data.ok && data.project_id) {
              self.projectId = data.project_id;
              resolve({ ok: true, path: data.path });
            } else {
              resolve({ ok: false, error: 'Invalid response' });
            }
          } catch (err) {
            resolve({ ok: false, error: 'Failed: ' + err.message });
          }
          return;
        }
        if (!jsonFile) {
          resolve({ ok: false, error: zipOnly ? 'ZIPファイル（PDF同梱）を選択してください' : 'project.json または .zip ファイルを選択してください' });
          return;
        }
        try {
          await self.init();
          const formData = new FormData();
          formData.append('project_file', jsonFile);
          const pdfFile = files.find((f) => f.name.toLowerCase().endsWith('.pdf'));
          if (pdfFile) formData.append('pdf_file', pdfFile);
          const uploadUrl = `${self.baseUrl}/api/projects/upload?session_id=${encodeURIComponent(self.sessionId)}`;
            
            const response = await fetch(uploadUrl, {
              method: 'POST',
              body: formData,
            });
            
            if (!response.ok) {
              const apiErr = await self._extractApiError(response, 'Failed to upload project');
              resolve({ ok: false, code: apiErr.code || null, error: apiErr.message });
              return;
            }
            
            const data = await response.json();
            if (data.ok && data.path) {
              // Extract project_id from path
              // Path format: C:\Users\...\projects\{project_id}\project.json
              // or: /path/to/projects/{project_id}/project.json
              const pathParts = data.path.split(/[/\\]/);
              let projectDir = null;
              for (let i = pathParts.length - 1; i >= 0; i--) {
                if (pathParts[i] === 'projects' && i + 1 < pathParts.length) {
                  projectDir = pathParts[i + 1];
                  break;
                }
              }
              if (!projectDir && pathParts.length >= 2) {
                projectDir = pathParts[pathParts.length - 2];
              }
              if (projectDir) {
                self.projectId = projectDir;
              }
              resolve({ ok: true, path: data.path });
            } else {
              resolve({ ok: false, error: 'Invalid response from server' });
            }
          } catch (err) {
            resolve({ ok: false, error: 'Failed to upload project: ' + err.message });
          }
      };
      input.click();
    });
  }

  async create_project_from_pdf_simple(pdfPathOrFile) {
    await this.init();
    
    // Handle both file object and path string (for compatibility)
    let pdfFile = pdfPathOrFile;
    if (typeof pdfPathOrFile === 'string') {
      // If path is provided, use pending file
      pdfFile = this._pendingPdfFile;
      if (!pdfFile) {
        return { ok: false, errors: ['PDF file not found. Please select a PDF file first.'] };
      }
    }
    
    if (!pdfFile) {
      return { ok: false, errors: ['PDF file is required'] };
    }
    
    const formData = new FormData();
    formData.append('pdf_file', pdfFile);
    // session_id must be in URL (Query) for FastAPI
    const url = `${this.baseUrl}/api/projects/create?session_id=${encodeURIComponent(this.sessionId)}`;

    const response = await fetch(url, { method: 'POST', body: formData });

    if (!response.ok) {
      let apiErr = await this._extractApiError(response, 'PDFの読み込みに失敗しました');
      if (response.status === 413) {
        apiErr = {
          code: 'UPLOAD_TOO_LARGE',
          message: 'PDFサイズが上限を超えています。サーバー管理者に「Nginx client_max_body_size」と「INPUTSTUDIO_MAX_UPLOAD_MB」を増やすよう依頼してください。',
        };
      } else if (response.status === 405) {
        apiErr = {
          code: 'METHOD_NOT_ALLOWED',
          message: 'サーバー接続エラー（Method Not Allowed）。サーバーを再起動してください。',
        };
      }
      return { ok: false, code: apiErr.code || null, errors: [apiErr.message] };
    }

    const data = await response.json();
    if (data.ok && data.path) {
      // Extract project_id from path
      // Path format: C:\Users\...\projects\{project_id}\project.json
      const pathParts = data.path.split(/[/\\]/);
      let projectDir = null;
      
      // Find 'projects' directory and get the next part as project_id
      for (let i = 0; i < pathParts.length; i++) {
        if (pathParts[i] === 'projects' && i + 1 < pathParts.length) {
          projectDir = pathParts[i + 1];
          break;
        }
      }
      
      // Fallback: use second-to-last part if 'projects' not found
      if (!projectDir && pathParts.length >= 2) {
        projectDir = pathParts[pathParts.length - 2];
      }
      
      if (projectDir) {
        this.projectId = projectDir;
      }
      // Clear pending file
      this._pendingPdfFile = null;
    }
    return data;
  }

  async load_project(path) {
    await this.init();
    
    // Extract project_id from path
    // Path format: C:\Users\...\projects\{project_id}\project.json
    // or: /path/to/projects/{project_id}/project.json
    const pathParts = path.split(/[/\\]/);
    let projectDir = null;
    
    // Find 'projects' directory and get the next part as project_id
    for (let i = 0; i < pathParts.length; i++) {
      if (pathParts[i] === 'projects' && i + 1 < pathParts.length) {
        projectDir = pathParts[i + 1];
        break;
      }
    }
    
    // Fallback: use second-to-last part if 'projects' not found
    if (!projectDir && pathParts.length >= 2) {
      projectDir = pathParts[pathParts.length - 2];
    }
    
    if (!projectDir) {
      return { ok: false, error: 'Could not extract project ID from path' };
    }
    
    this.projectId = projectDir;

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/load?session_id=${this.sessionId}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to load project');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }

    return await response.json();
  }

  async get_preview_png_base64_page(pageIndex) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/preview/${pageIndex}?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to get preview');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }

    return await response.json();
  }

  async get_preview_png_base64(tagOrFid) {
    // For web version, we need to find the page from tag/fid
    // This is a simplified version - you may need to enhance it
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    // Try page 0 first (can be enhanced to find correct page)
    return this.get_preview_png_base64_page(0);
  }

  async save_current_project(makeFilledPdf = false) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/save?session_id=${this.sessionId}&make_filled_pdf=${makeFilledPdf}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to save project');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }

    return await response.json();
  }

  async download_filled_pdf(filename) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/export?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to export PDF' };
    }

    const blob = await response.blob();
    const suggestedName = (filename || `${this.projectId}_filled`).replace(/\.pdf$/i, '') + '.pdf';

    if (typeof window.showSaveFilePicker === 'function') {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName,
          types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return { ok: true };
      } catch (e) {
        if (e.name === 'AbortError') return { ok: false, error: 'cancelled' };
        throw e;
      }
    }

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    return { ok: true };
  }

  async set_value(tag, value) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/values?session_id=${this.sessionId}&tag=${encodeURIComponent(tag)}&value=${encodeURIComponent(value)}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to set value' };
    }

    return await response.json();
  }

  async add_text_field(tag, page, x, y, fontSize) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/placements?session_id=${this.sessionId}&tag=${encodeURIComponent(tag)}&page=${page}&x=${x}&y=${y}&font_size=${fontSize}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to add text field' };
    }

    return await response.json();
  }

  async set_element_pos(fid, x, y) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/placements/${fid}/position?session_id=${this.sessionId}&x=${x}&y=${y}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to set element position' };
    }

    return await response.json();
  }

  async update_placement(fid, patch) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/placements/${fid}?session_id=${this.sessionId}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(patch),
      }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to update placement' };
    }

    return await response.json();
  }

  async get_element_info(fid) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/elements/${fid}?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to get element info' };
    }

    return await response.json();
  }

  async delete_elements(fids) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const fidsParam = fids.map(f => `fids=${encodeURIComponent(f)}`).join('&');
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/elements?session_id=${this.sessionId}&${fidsParam}`,
      { method: 'DELETE' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to delete elements' };
    }

    return await response.json();
  }

  async set_project_payload(payload) {
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/payload?session_id=${this.sessionId}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to set project payload' };
    }

    return await response.json();
  }

  async get_workers() {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/workers?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to get workers' };
    }

    return await response.json();
  }

  async get_admin_settings() {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/admin/settings?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to get admin settings' };
    }

    return await response.json();
  }

  async update_admin_settings(patch) {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/admin/settings?session_id=${this.sessionId}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(patch),
      }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to update admin settings' };
    }

    return await response.json();
  }

  async set_ui_mode(mode) {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/ui/mode?session_id=${this.sessionId}&mode=${encodeURIComponent(mode)}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to set UI mode' };
    }

    return await response.json();
  }

  async start_work(workerId) {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/work/start?session_id=${this.sessionId}&worker_id=${encodeURIComponent(workerId)}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to start work' };
    }

    return await response.json();
  }

  async toggle_private() {
    await this.init();
    
    const response = await fetch(
      `${this.baseUrl}/api/work/private/toggle?session_id=${this.sessionId}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to toggle private' };
    }

    return await response.json();
  }

  async finish(reportMeta) {
    // This endpoint needs to be added to the server
    await this.init();
    
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }

    // Export PDF
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/export?session_id=${this.sessionId}`
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to finish' };
    }

    // Download the PDF
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${this.projectId}_filled.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    return { ok: true };
  }

  async save_project_as(name, makeFilledPdf = true) {
    // This endpoint needs to be added to the server
    await this.init();
    
    return { ok: false, error: 'not_implemented' };
  }

  async upsert_worker(worker) {
    // This endpoint needs to be added to the server
    await this.init();
    
    return { ok: true, id: worker.id || 'w1' };
  }

  async delete_worker(workerId) {
    // This endpoint needs to be added to the server
    await this.init();
    
    return { ok: true };
  }

  async append_pdf_to_project(pdfPath) {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const pdfFile = this._pendingPdfFile;
    if (!pdfFile) {
      return { ok: false, error: 'PDF file not found. Please select a PDF file first.' };
    }
    const formData = new FormData();
    formData.append('pdf_file', pdfFile);
    const appendUrl = `${this.baseUrl}/api/projects/${this.projectId}/append-pdf?session_id=${encodeURIComponent(this.sessionId)}`;
    const response = await fetch(appendUrl, { method: 'POST', body: formData });
    if (!response.ok) {
      let apiErr = await this._extractApiError(response, 'PDF追加に失敗しました');
      if (response.status === 413) {
        apiErr = {
          code: 'UPLOAD_TOO_LARGE',
          message: 'PDFサイズが上限を超えています。サーバー管理者に「Nginx client_max_body_size」と「INPUTSTUDIO_MAX_UPLOAD_MB」を増やすよう依頼してください。',
        };
      }
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }
    this._pendingPdfFile = null;
    return await response.json();
  }

  async copy_page_with_elements(pageIndex) {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/copy-page?session_id=${encodeURIComponent(this.sessionId)}&page_index=${pageIndex}`,
      { method: 'POST' }
    );
    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to copy page');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }
    return await response.json();
  }

  async delete_page_from_project(pageIndex) {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/delete-page?session_id=${encodeURIComponent(this.sessionId)}&page_index=${pageIndex}`,
      { method: 'POST' }
    );
    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to delete page');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }
    return await response.json();
  }

  async reorder_pages(order) {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/reorder-pages?session_id=${encodeURIComponent(this.sessionId)}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ order }),
      }
    );
    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to reorder pages');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }
    return await response.json();
  }

  async get_project_json() {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/export-json?session_id=${encodeURIComponent(this.sessionId)}`
    );
    if (!response.ok) {
      const apiErr = await this._extractApiError(response, 'Failed to export project');
      return { ok: false, code: apiErr.code || null, error: apiErr.message };
    }
    return await response.json();
  }

  async save_project_to_picker(suggestedName) {
    await this.init();
    if (!this.projectId) {
      return { ok: false, error: 'no_project' };
    }
    const base = (suggestedName || 'project').replace(/[\\/:*?"<>|]/g, '_').replace(/\.(json|zip)$/i, '');
    const baseName = base + '.zip';
    let handle = null;
    if (typeof window.showSaveFilePicker === 'function') {
      try {
        handle = await window.showSaveFilePicker({
          suggestedName: baseName,
          types: [{ description: 'ZIP（PDF同梱）', accept: { 'application/zip': ['.zip'] } }],
        });
      } catch (e) {
        if (e.name === 'AbortError') return { ok: false, error: 'cancelled' };
        throw e;
      }
    }
    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/export-zip?session_id=${encodeURIComponent(this.sessionId)}`
    );
    if (!response.ok) {
      let errMsg = 'Failed to export project';
      try {
        const err = await response.json();
        const d = err.detail;
        if (Array.isArray(d)) errMsg = d.map(x => (typeof x === 'object' && x?.msg) ? x.msg : String(x)).join('; ');
        else if (typeof d === 'string') errMsg = d;
      } catch (_) {}
      return { ok: false, error: errMsg };
    }
    const blob = await response.blob();
    if (blob.size < 100) {
      return { ok: false, error: 'ZIPの取得に失敗しました。空のファイルが返されました。' };
    }
    if (handle) {
      try {
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return { ok: true };
      } catch (e) {
        throw e;
      }
    }
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = baseName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    return { ok: true };
  }
}

// Initialize web API if not in desktop mode
if (typeof window !== 'undefined') {
  // Check if we're in web mode (not desktop)
  const isWebMode = !window.chrome?.webview && !window.pywebview;
  
  if (isWebMode) {
    window.__INPUTSTUDIO_WEB__ = true;  // For app.js to hide desktop-only UI
    const webApi = new WebAPI();
    
    // Bind all methods so they work when called as: const fn = api.pick_pdf; fn()
    const boundApi = {};
    for (const key of Object.getOwnPropertyNames(Object.getPrototypeOf(webApi))) {
      if (key !== 'constructor' && typeof webApi[key] === 'function') {
        boundApi[key] = webApi[key].bind(webApi);
      }
    }
    
    window.pywebview = {
      api: boundApi,
    };
    
    // Defensive: verify api has required methods
    if (!window.pywebview.api || typeof window.pywebview.api.create_project_from_pdf_simple !== 'function') {
      console.error('WebAPI initialization failed: create_project_from_pdf_simple is missing');
    }
  }
}
