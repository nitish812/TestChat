/**
 * workitems.js — Work Items
 *
 * Displays work items for a selected project with filtering,
 * pagination (50 per page), and a detail modal.
 */

const WorkItemsModule = (() => {
  const PAGE_SIZE = 50;

  let _items       = [];
  let _currentPage = 1;
  let _context     = null;   // { connId, project }

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${c.id}">${_esc(c.name)}</option>`
    ).join('');

    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-list-check"></i> Work Items</div>
      <p class="page-subtitle">Browse and track work items across your projects.</p>

      <!-- Selectors row -->
      <div class="filter-bar mb-16">
        <select id="wi-conn-select" style="min-width:160px">
          <option value="">Select Organization…</option>
          ${connOptions}
        </select>
        <select id="wi-project-select" style="min-width:180px" disabled>
          <option value="">Select Project…</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-load-btn" disabled>
          <i class="fa-solid fa-download"></i> Load
        </button>
      </div>

      <!-- Filters -->
      <div class="filter-bar mb-8" id="wi-filters" style="display:none">
        <select id="wi-type-filter">
          <option value="">All Types</option>
          <option value="Bug">Bug</option>
          <option value="Task">Task</option>
          <option value="User Story">User Story</option>
          <option value="Epic">Epic</option>
          <option value="Feature">Feature</option>
          <option value="Issue">Issue</option>
          <option value="Test Case">Test Case</option>
        </select>
        <select id="wi-state-filter">
          <option value="">All States</option>
          <option value="New">New</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
          <option value="Done">Done</option>
        </select>
        <select id="wi-priority-filter">
          <option value="">All Priorities</option>
          <option value="1">1 – Critical</option>
          <option value="2">2 – High</option>
          <option value="3">3 – Medium</option>
          <option value="4">4 – Low</option>
        </select>
        <input type="text" id="wi-search" placeholder="🔍 Search title…" style="min-width:160px" />
        <button class="btn btn-secondary btn-sm" id="wi-apply-filter-btn">
          <i class="fa-solid fa-filter"></i> Apply
        </button>
      </div>

      <div id="wi-content"></div>
    `;

    _bindSelectorEvents(container);

    // Pre-populate from params (e.g. navigated from a project card)
    if (params.connId) {
      const sel = container.querySelector('#wi-conn-select');
      sel.value = params.connId;
      await _loadProjects(container, params.connId);
      if (params.project) {
        const pSel = container.querySelector('#wi-project-select');
        pSel.value = params.project;
        if (pSel.value) {
          container.querySelector('#wi-load-btn').disabled = false;
          await _fetchItems(container);
        }
      }
    }
  }

  async function refresh(container) {
    if (_context) await _fetchItems(container);
  }

  // ─── Private ───────────────────────────────────────────────────

  function _bindSelectorEvents(container) {
    container.addEventListener('change', async e => {
      if (e.target.id === 'wi-conn-select') {
        const connId = e.target.value;
        container.querySelector('#wi-project-select').innerHTML = '<option value="">Loading…</option>';
        container.querySelector('#wi-project-select').disabled = true;
        container.querySelector('#wi-load-btn').disabled = true;
        if (connId) await _loadProjects(container, connId);
      }
      if (e.target.id === 'wi-project-select') {
        container.querySelector('#wi-load-btn').disabled = !e.target.value;
      }
    });

    container.addEventListener('click', async e => {
      if (e.target.closest('#wi-load-btn')) await _fetchItems(container);
      if (e.target.closest('#wi-apply-filter-btn')) _renderTable(container);
    });

    container.addEventListener('keydown', e => {
      if (e.key === 'Enter' && e.target.id === 'wi-search') _renderTable(container);
    });
  }

  async function _loadProjects(container, connId) {
    const conn = ConnectionsModule.getById(connId);
    if (!conn) return;

    const pSel = container.querySelector('#wi-project-select');
    const result = await AzureApi.getProjects(conn.orgUrl, conn.pat);
    if (result.error) { App.showToast(result.message, 'error'); return; }

    const projects = result.value || [];
    pSel.innerHTML = '<option value="">Select Project…</option>' +
      projects.map(p => `<option value="${_esc(p.name)}">${_esc(p.name)}</option>`).join('');
    pSel.disabled = false;
  }

  async function _fetchItems(container) {
    const connId  = container.querySelector('#wi-conn-select').value;
    const project = container.querySelector('#wi-project-select').value;
    if (!connId || !project) return;

    const conn = ConnectionsModule.getById(connId);
    if (!conn) return;

    _context = { connId, project, conn };
    _currentPage = 1;

    const content = container.querySelector('#wi-content');
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading work items…</p></div>';
    container.querySelector('#wi-filters').style.display = 'none';

    const filters = _getFilters(container);
    const result  = await AzureApi.getWorkItems(conn.orgUrl, project, conn.pat, filters);

    if (result.error) {
      content.innerHTML = `
        <div class="error-state">
          <i class="fa-solid fa-circle-exclamation"></i>
          <p>${_esc(result.message)}</p>
          ${result.cors ? '<p class="text-sm">Tip: CORS restrictions may apply. Check the README for proxy setup instructions.</p>' : ''}
        </div>`;
      return;
    }

    _items = result.value || [];
    container.querySelector('#wi-filters').style.display = '';
    _renderTable(container);
  }

  function _getFilters(container) {
    return {
      type:     container.querySelector('#wi-type-filter')?.value || '',
      state:    container.querySelector('#wi-state-filter')?.value || '',
      priority: container.querySelector('#wi-priority-filter')?.value || '',
    };
  }

  function _renderTable(container) {
    const search    = (container.querySelector('#wi-search')?.value || '').toLowerCase();
    const content   = container.querySelector('#wi-content');

    let filtered = _items;
    if (search) filtered = filtered.filter(w =>
      (w.fields['System.Title'] || '').toLowerCase().includes(search)
    );

    const total    = filtered.length;
    const pages    = Math.ceil(total / PAGE_SIZE) || 1;
    _currentPage   = Math.min(_currentPage, pages);
    const start    = (_currentPage - 1) * PAGE_SIZE;
    const pageItems= filtered.slice(start, start + PAGE_SIZE);

    if (filtered.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-list-check"></i>
          <p>No work items found.</p>
        </div>`;
      return;
    }

    content.innerHTML = `
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              <th>ID</th><th>Title</th><th>Type</th><th>State</th>
              <th>Assigned To</th><th>Priority</th>
              <th>Created</th><th>Updated</th>
            </tr>
          </thead>
          <tbody id="wi-tbody">
            ${pageItems.map(w => _rowHtml(w)).join('')}
          </tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}
    `;

    // Row click → open detail modal
    content.querySelectorAll('.wi-row').forEach(row => {
      row.addEventListener('click', () => _openDetail(row.dataset.id));
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });
  }

  function _rowHtml(w) {
    const f       = w.fields;
    const id      = f['System.Id']   || w.id;
    const title   = f['System.Title']|| '(no title)';
    const type    = f['System.WorkItemType'] || '';
    const state   = f['System.State'] || '';
    const assigned= (f['System.AssignedTo']?.displayName) || f['System.AssignedTo'] || '—';
    const priority= f['Microsoft.VSTS.Common.Priority'] || '—';
    const created = f['System.CreatedDate'] ? _fmtDate(f['System.CreatedDate']) : '—';
    const updated = f['System.ChangedDate'] ? _fmtDate(f['System.ChangedDate']) : '—';

    return `
      <tr class="wi-row" data-id="${id}" style="cursor:pointer">
        <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
        <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
        <td>${_typeBadge(type)}</td>
        <td>${_stateBadge(state)}</td>
        <td class="truncate" style="max-width:140px">${_esc(String(assigned))}</td>
        <td>${_priorityBadge(priority)}</td>
        <td class="text-muted text-sm">${created}</td>
        <td class="text-muted text-sm">${updated}</td>
      </tr>`;
  }

  async function _openDetail(id) {
    if (!_context) return;
    const { conn } = _context;

    App.openModal('Work Item #' + id, '<div class="loading-state"><div class="spinner"></div><p>Loading…</p></div>');

    const result = await AzureApi.getWorkItemDetail(conn.orgUrl, id, conn.pat);
    if (result.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i><p>${_esc(result.message)}</p></div>`);
      return;
    }

    const f = result.fields;
    App.updateModalBody(`
      <div class="detail-row"><span class="detail-label">ID</span><span class="detail-value font-mono">#${result.id}</span></div>
      <div class="detail-row"><span class="detail-label">Title</span><span class="detail-value">${_esc(f['System.Title'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">Type</span><span class="detail-value">${_typeBadge(f['System.WorkItemType'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">State</span><span class="detail-value">${_stateBadge(f['System.State'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">Assigned To</span><span class="detail-value">${_esc(String(f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || '—'))}</span></div>
      <div class="detail-row"><span class="detail-label">Priority</span><span class="detail-value">${_priorityBadge(f['Microsoft.VSTS.Common.Priority'] || '—')}</span></div>
      <div class="detail-row"><span class="detail-label">Area Path</span><span class="detail-value">${_esc(f['System.AreaPath'] || '—')}</span></div>
      <div class="detail-row"><span class="detail-label">Iteration</span><span class="detail-value">${_esc(f['System.IterationPath'] || '—')}</span></div>
      <div class="detail-row"><span class="detail-label">Created</span><span class="detail-value">${_esc(f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toLocaleString() : '—')}</span></div>
      <div class="detail-row"><span class="detail-label">Updated</span><span class="detail-value">${_esc(f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toLocaleString() : '—')}</span></div>
      ${f['System.Description'] ? `
      <div class="detail-row" style="flex-direction:column;gap:6px">
        <span class="detail-label">Description</span>
        <div style="font-size:.82rem;border:1px solid var(--border-color);padding:10px;border-radius:var(--radius-sm);background:var(--bg-table-alt);">
          ${f['System.Description']}
        </div>
      </div>` : ''}
    `);
  }

  function _paginationHtml(current, total) {
    if (total <= 1) return '';
    let html = '<div class="pagination">';
    html += `<button class="page-btn" data-page="${current - 1}" ${current === 1 ? 'disabled' : ''}><i class="fa-solid fa-chevron-left"></i></button>`;
    for (let p = 1; p <= total; p++) {
      if (total > 7 && p > 2 && p < total - 1 && Math.abs(p - current) > 1) {
        if (p === 3 || p === total - 2) html += '<span class="page-info">…</span>';
        continue;
      }
      html += `<button class="page-btn ${p === current ? 'active' : ''}" data-page="${p}">${p}</button>`;
    }
    html += `<button class="page-btn" data-page="${current + 1}" ${current === total ? 'disabled' : ''}><i class="fa-solid fa-chevron-right"></i></button>`;
    html += '</div>';
    return html;
  }

  // ─── Badge helpers ─────────────────────────────────────────────

  function _typeBadge(type) {
    const icons = {
      'Bug':        { icon: 'fa-bug',          cls: 'badge-danger' },
      'Task':       { icon: 'fa-thumbtack',    cls: 'badge-primary' },
      'User Story': { icon: 'fa-book',         cls: 'badge-success' },
      'Epic':       { icon: 'fa-bolt',         cls: 'badge-warning' },
      'Feature':    { icon: 'fa-star',         cls: 'badge-info' },
      'Issue':      { icon: 'fa-circle-exclamation', cls: 'badge-danger' },
      'Test Case':  { icon: 'fa-flask',        cls: 'badge-secondary' },
    };
    const d = icons[type] || { icon: 'fa-file', cls: 'badge-secondary' };
    return `<span class="badge ${d.cls}"><i class="fa-solid ${d.icon}"></i> ${_esc(type)}</span>`;
  }

  function _stateBadge(state) {
    const map = {
      'Active':   'badge-primary',
      'New':      'badge-info',
      'Resolved': 'badge-warning',
      'Closed':   'badge-secondary',
      'Done':     'badge-success',
    };
    return `<span class="badge ${map[state] || 'badge-secondary'}">${_esc(state)}</span>`;
  }

  function _priorityBadge(p) {
    const map = { 1: 'badge-danger', 2: 'badge-warning', 3: 'badge-primary', 4: 'badge-secondary' };
    return `<span class="badge ${map[p] || 'badge-secondary'}">${p}</span>`;
  }

  function _fmtDate(s) {
    const d = new Date(s);
    return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function _esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  return { render, refresh };
})();
