/**
 * connections.js — Connection Manager
 *
 * Manages Azure DevOps connections stored in localStorage.
 * Each connection:  { id, name, orgUrl, pat, active, createdAt, defaultProject?, patExpiry? }
 *
 * Improvement #1: PAT Expiry Warning Badge
 */

const ConnectionsModule = (() => {
  const STORAGE_KEY = 'ado_connections';

  // ─── Storage helpers ───────────────────────────────────────────

  function _load() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    } catch {
      return [];
    }
  }

  function _save(connections) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(connections));
  }

  // ─── Public API ────────────────────────────────────────────────

  function getAll() { return _load(); }

  function getActive() { return _load().filter(c => c.active !== false); }

  function getById(id) { return _load().find(c => c.id === id) || null; }

  function add(name, orgUrl, pat, defaultProject, patExpiry) {
    const connections = _load();
    const id = 'conn_' + Date.now();
    const conn = { id, name, orgUrl: orgUrl.replace(/\/$/, ''), pat, active: true, createdAt: new Date().toISOString() };
    if (defaultProject) conn.defaultProject = defaultProject;
    if (patExpiry) conn.patExpiry = patExpiry;
    connections.push(conn);
    _save(connections);
    return id;
  }

  function update(id, fields) {
    const connections = _load().map(c => {
      if (c.id !== id) return c;
      return {
        ...c,
        ...fields,
        orgUrl: fields.orgUrl ? fields.orgUrl.replace(/\/$/, '') : c.orgUrl,
      };
    });
    _save(connections);
  }

  function remove(id) {
    _save(_load().filter(c => c.id !== id));
  }

  function toggleActive(id) {
    const connections = _load().map(c => c.id === id ? { ...c, active: !c.active } : c);
    _save(connections);
    return connections.find(c => c.id === id);
  }

  /** Return PAT with only last 4 chars visible. */
  function maskPat(pat) {
    if (!pat || pat.length <= 4) return '****';
    return '•'.repeat(pat.length - 4) + pat.slice(-4);
  }

  /** Compute days remaining until PAT expiry. Returns null if no expiry set. */
  function _patDaysRemaining(patExpiry) {
    if (!patExpiry) return null;
    const expDate = new Date(patExpiry);
    if (isNaN(expDate.getTime())) return null;
    const msPerDay = 1000 * 60 * 60 * 24;
    return Math.ceil((expDate.getTime() - Date.now()) / msPerDay);
  }

  /** Build the expiry badge HTML for a connection. Returns '' if no expiry set. */
  function _expiryBadge(c) {
    const days = _patDaysRemaining(c.patExpiry);
    if (days === null) return '';
    if (days <= 0) {
      return '<span class="badge badge-danger"><i class="fa-solid fa-triangle-exclamation"></i> PAT Expired</span>';
    }
    if (days <= 14) {
      return '<span class="badge badge-warning"><i class="fa-solid fa-clock"></i> Expires in ' + days + 'd</span>';
    }
    return '';
  }

  // ─── Render ────────────────────────────────────────────────────

  function render(container) {
    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-plug"></i> Connections</div>
      <p class="page-subtitle">Manage your Azure DevOps organization connections.</p>

      <!-- Add / Edit form -->
      <div class="card mb-16" id="conn-form-card">
        <div class="card-header">
          <span class="card-title" id="conn-form-title">Add Connection</span>
        </div>
        <form id="conn-form" novalidate>
          <input type="hidden" id="conn-edit-id" />
          <div class="form-group">
            <label for="conn-name">Organization Name <span style="color:var(--color-danger)">*</span></label>
            <input type="text" id="conn-name" placeholder="My Org" required />
          </div>
          <div class="form-group">
            <label for="conn-url">Organization URL <span style="color:var(--color-danger)">*</span></label>
            <input type="url" id="conn-url" placeholder="https://dev.azure.com/myorg" required />
            <span class="form-hint">e.g., https://dev.azure.com/myorg</span>
          </div>
          <div class="form-group">
            <label for="conn-pat">Personal Access Token (PAT) <span style="color:var(--color-danger)">*</span></label>
            <input type="password" id="conn-pat" placeholder="••••••••••••••••" required autocomplete="new-password" />
            <span class="form-hint">PAT is stored in browser localStorage.</span>
          </div>
          <div class="form-group">
            <label for="conn-expiry">PAT Expiry Date <span class="text-muted text-sm">(optional)</span></label>
            <input type="date" id="conn-expiry" />
            <span class="form-hint">Optional: set a reminder when your PAT expires.</span>
          </div>
          <div class="form-group">
            <label for="conn-project">Default Project <span class="text-muted text-sm">(optional)</span></label>
            <select id="conn-project" disabled>
              <option value="">Enter URL &amp; PAT first…</option>
            </select>
            <span class="form-hint">Select a default project for this connection, or leave blank for no default.</span>
          </div>
          <div id="conn-form-error" class="form-error mb-8 hidden"></div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
            <button type="submit" class="btn btn-primary" id="conn-save-btn">
              <i class="fa-solid fa-floppy-disk"></i> Save Connection
            </button>
            <button type="button" class="btn btn-secondary" id="conn-test-btn">
              <i class="fa-solid fa-plug-circle-check"></i> Test Connection
            </button>
            <button type="button" class="btn btn-secondary hidden" id="conn-cancel-btn">
              <i class="fa-solid fa-xmark"></i> Cancel
            </button>
            <span id="conn-test-result" class="text-sm"></span>
          </div>
        </form>
      </div>

      <!-- Connection list -->
      <div class="section-title">Saved Connections</div>
      <div id="conn-list" class="connection-list"></div>
    `;

    _renderList(container);
    _bindFormEvents(container);
  }

  function _renderList(container) {
    const listEl = container.querySelector('#conn-list');
    const connections = _load();

    if (connections.length === 0) {
      listEl.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-plug-circle-xmark"></i>
          <p>No connections yet. Add one above to get started.</p>
        </div>`;
      return;
    }

    listEl.innerHTML = connections.map(c => `
      <div class="connection-item ${c.active === false ? 'inactive' : ''}" data-id="${c.id}">
        <div class="connection-icon"><i class="fa-brands fa-windows"></i></div>
        <div class="connection-info">
          <div class="connection-name">${_esc(c.name)} ${_expiryBadge(c)}</div>
          <div class="connection-url">${_esc(c.orgUrl)}</div>
          <div class="connection-pat font-mono">PAT: ${maskPat(c.pat)}</div>
          ${c.defaultProject ? `<div class="connection-url">Project: ${_esc(c.defaultProject)}</div>` : '<div class="connection-url text-muted">Project: All</div>'}
        </div>
        <div class="connection-actions">
          <label class="toggle-switch" title="${c.active !== false ? 'Active' : 'Inactive'}">
            <input type="checkbox" class="conn-toggle" data-id="${c.id}" ${c.active !== false ? 'checked' : ''} />
            <span class="toggle-slider"></span>
          </label>
          <button class="btn-icon conn-edit-btn" data-id="${c.id}" title="Edit">
            <i class="fa-solid fa-pen"></i>
          </button>
          <button class="btn-icon conn-delete-btn" data-id="${c.id}" title="Delete" style="color:var(--color-danger)">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </div>
    `).join('');

    // Bind list events
    listEl.querySelectorAll('.conn-toggle').forEach(el => {
      el.addEventListener('change', () => {
        toggleActive(el.dataset.id);
        App.refreshOrgSelector();
        _renderList(container);
      });
    });

    listEl.querySelectorAll('.conn-edit-btn').forEach(el => {
      el.addEventListener('click', () => _startEdit(container, el.dataset.id));
    });

    listEl.querySelectorAll('.conn-delete-btn').forEach(el => {
      el.addEventListener('click', () => {
        if (confirm('Delete this connection?')) {
          remove(el.dataset.id);
          App.refreshOrgSelector();
          App.showToast('Connection removed.', 'info');
          _renderList(container);
        }
      });
    });
  }

  function _startEdit(container, id) {
    const c = getById(id);
    if (!c) return;
    container.querySelector('#conn-form-title').textContent = 'Edit Connection';
    container.querySelector('#conn-edit-id').value = c.id;
    container.querySelector('#conn-name').value = c.name;
    container.querySelector('#conn-url').value = c.orgUrl;
    container.querySelector('#conn-pat').value = c.pat;
    container.querySelector('#conn-expiry').value = c.patExpiry || '';
    container.querySelector('#conn-cancel-btn').classList.remove('hidden');
    container.querySelector('#conn-form-card').scrollIntoView({ behavior: 'smooth' });
    // Populate project dropdown with saved defaultProject pre-selected
    _fetchProjectsForDropdown(container, c.orgUrl, c.pat, c.defaultProject || '');
  }

  function _resetForm(container) {
    container.querySelector('#conn-form-title').textContent = 'Add Connection';
    container.querySelector('#conn-edit-id').value = '';
    container.querySelector('#conn-name').value = '';
    container.querySelector('#conn-url').value = '';
    container.querySelector('#conn-pat').value = '';
    container.querySelector('#conn-expiry').value = '';
    container.querySelector('#conn-cancel-btn').classList.add('hidden');
    container.querySelector('#conn-form-error').classList.add('hidden');
    container.querySelector('#conn-test-result').textContent = '';
    const projSel = container.querySelector('#conn-project');
    projSel.innerHTML = '<option value="">Enter URL &amp; PAT first…</option>';
    projSel.disabled = true;
  }

  function _bindFormEvents(container) {
    const form        = container.querySelector('#conn-form');
    const saveBtn     = container.querySelector('#conn-save-btn');
    const testBtn     = container.querySelector('#conn-test-btn');
    const cancelBtn   = container.querySelector('#conn-cancel-btn');
    const errorEl     = container.querySelector('#conn-form-error');
    const testResult  = container.querySelector('#conn-test-result');

    cancelBtn.addEventListener('click', () => _resetForm(container));

    // Auto-fetch projects on blur when both URL and PAT have values
    function _tryFetchProjects() {
      const url = container.querySelector('#conn-url').value.trim();
      const pat = container.querySelector('#conn-pat').value.trim();
      if (url && pat) _fetchProjectsForDropdown(container, url, pat, '');
    }
    container.querySelector('#conn-url').addEventListener('blur', _tryFetchProjects);
    container.querySelector('#conn-pat').addEventListener('blur', _tryFetchProjects);

    testBtn.addEventListener('click', async () => {
      const url = container.querySelector('#conn-url').value.trim();
      const pat = container.querySelector('#conn-pat').value.trim();
      if (!url || !pat) { testResult.textContent = '⚠ Enter URL and PAT first.'; return; }

      testBtn.disabled = true;
      testResult.innerHTML = '<span class="spinner" style="width:14px;height:14px;border-width:2px;vertical-align:middle;"></span> Testing…';

      const res = await AzureApi.validateConnection(url, pat);
      testBtn.disabled = false;

      if (res.valid) {
        testResult.innerHTML = '<span style="color:var(--color-success)"><i class="fa-solid fa-circle-check"></i> Connected (' + res.projectCount + ' project' + (res.projectCount !== 1 ? 's' : '') + ')</span>';
        // Populate project dropdown after successful test
        const currentProject = container.querySelector('#conn-project').value;
        _fetchProjectsForDropdown(container, url, pat, currentProject);
      } else {
        testResult.innerHTML = '<span style="color:var(--color-danger)"><i class="fa-solid fa-circle-xmark"></i> ' + _esc(res.message) + '</span>';
      }
    });

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      errorEl.classList.add('hidden');

      const editId        = container.querySelector('#conn-edit-id').value;
      const name          = container.querySelector('#conn-name').value.trim();
      const url           = container.querySelector('#conn-url').value.trim();
      const pat           = container.querySelector('#conn-pat').value.trim();
      const patExpiry     = container.querySelector('#conn-expiry').value.trim();
      const defaultProject= container.querySelector('#conn-project').value;

      if (!name || !url || !pat) {
        errorEl.textContent = 'All fields are required.';
        errorEl.classList.remove('hidden');
        return;
      }

      saveBtn.disabled = true;
      saveBtn.innerHTML = '<span class="spinner" style="width:14px;height:14px;border-width:2px;vertical-align:middle;"></span> Saving…';

      const valid = await AzureApi.validateConnection(url, pat);
      saveBtn.disabled = false;
      saveBtn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Save Connection';

      if (!valid.valid) {
        errorEl.textContent = 'Validation failed: ' + valid.message;
        errorEl.classList.remove('hidden');
        return;
      }

      if (editId) {
        update(editId, { name, orgUrl: url, pat, defaultProject: defaultProject || undefined, patExpiry: patExpiry || undefined });
        App.showToast('Connection updated successfully.', 'success');
      } else {
        add(name, url, pat, defaultProject || undefined, patExpiry || undefined);
        App.showToast('Connection added successfully.', 'success');
      }

      App.refreshOrgSelector();
      _resetForm(container);
      _renderList(container);
    });
  }

  /**
   * Fetch projects from the API and populate the #conn-project dropdown.
   * @param {Element} container
   * @param {string} url       Organisation URL
   * @param {string} pat       PAT
   * @param {string} selected  Value to pre-select (may be empty string)
   */
  async function _fetchProjectsForDropdown(container, url, pat, selected) {
    const projSel = container.querySelector('#conn-project');
    projSel.disabled = true;
    projSel.innerHTML = '<option value="">Loading projects…</option>';

    const result = await AzureApi.getProjects(url, pat);

    if (result.error) {
      projSel.innerHTML = '<option value="">Could not load projects</option>';
      return;
    }

    const projects = result.value || [];
    projSel.innerHTML =
      '<option value="">All Projects (no default)</option>' +
      projects.map(p => '<option value="' + _esc(p.name) + '"' + (p.name === selected ? ' selected' : '') + '>' + _esc(p.name) + '</option>').join('');
    projSel.disabled = false;
  }

  /** Simple HTML-escape for rendering user data. */
  function _esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  return { getAll, getActive, getById, add, update, remove, toggleActive, maskPat, render };
})();
