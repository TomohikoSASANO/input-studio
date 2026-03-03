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
    // In web version, trigger file input
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf';
      input.onchange = async (e) => {
        const file = e.target.files[0];
        if (file) {
          const result = await this.create_project_from_pdf_simple(file);
          resolve(result);
        } else {
          resolve({ ok: false });
        }
      };
      input.click();
    });
  }

  async pick_project() {
    // In web version, trigger file input
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';
      input.onchange = async (e) => {
        const file = e.target.files[0];
        if (file) {
          const text = await file.text();
          try {
            const data = JSON.parse(text);
            // For web version, we need to handle project loading differently
            // This is a simplified version
            resolve({ ok: true, path: file.name });
          } catch (err) {
            resolve({ ok: false, error: 'Invalid JSON' });
          }
        } else {
          resolve({ ok: false });
        }
      };
      input.click();
    });
  }

  async create_project_from_pdf_simple(pdfFile) {
    await this.init();
    
    const formData = new FormData();
    formData.append('pdf_file', pdfFile);
    formData.append('session_id', this.sessionId);

    const response = await fetch(`${this.baseUrl}/api/projects/create`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, errors: [error.detail || 'Failed to create project'] };
    }

    const data = await response.json();
    if (data.ok && data.path) {
      // Extract project_id from path
      const pathParts = data.path.split(/[/\\]/);
      const projectDir = pathParts[pathParts.length - 2];
      this.projectId = projectDir;
    }
    return data;
  }

  async load_project(path) {
    await this.init();
    
    // Extract project_id from path
    const pathParts = path.split(/[/\\]/);
    const projectDir = pathParts[pathParts.length - 2];
    this.projectId = projectDir;

    const response = await fetch(
      `${this.baseUrl}/api/projects/${this.projectId}/load?session_id=${this.sessionId}`,
      { method: 'POST' }
    );

    if (!response.ok) {
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to load project' };
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
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to get preview' };
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
      const error = await response.json();
      return { ok: false, error: error.detail || 'Failed to save project' };
    }

    return await response.json();
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
    // This endpoint needs to be added to the server
    await this.init();
    
    return { ok: false, error: 'not_implemented' };
  }
}

// Initialize web API if not in desktop mode
if (typeof window !== 'undefined') {
  // Check if we're in web mode (not desktop)
  const isWebMode = !window.chrome?.webview && !window.pywebview;
  
  if (isWebMode) {
    const webApi = new WebAPI();
    
    // Create window.pywebview.api for compatibility
    window.pywebview = {
      api: webApi,
    };
  }
}
