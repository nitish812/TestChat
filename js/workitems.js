/**
 * workitems.js — Work Items
 *
 * Displays work items for a selected project with filtering,
 * pagination (50 per page), and a detail modal.
 * Supports "Single Project" and "All Connections" modes.
 */

const WorkItemsModule = (() => {
  const PAGE_SIZE = 50;

  let _items       = [];
  let _currentPage = 1;
  let _context     = null;   // { connId, project }
  let _mode        = 'single'; // 'single' | 'all'
  let _allSummary  = null;   // { total, connCount, failCount }
  let _sortCol = null;
  let _sortDir = 'asc';

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];
    _mode = 'single';
    _allSummary = null;

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${c.id}">${_esc(c.name)}</option>`
    ).join('');

    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-list-check"></i> Work Items</div>
      <p class="page-subtitle">Browse and track work items across your projects.</p>

      <!-- Mode toggle -->
      <div class="filter-bar mb-8" id="wi-mode-bar">
        <label style="font-weight:600">View:</label>
        <button class="btn btn-primary btn-sm active" id="wi-mode-single" data-mode="single">Single Project</button>
        <button class="btn btn-secondary btn-sm" id="wi-mode-all" data-mode="all">All Connections</button>
      </div>

      <!-- Selectors row (single mode) -->
      <div class="filter-bar mb-16" id="wi-single-bar">
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

      <!-- All connections bar (all mode, hidden by default) -->
      <div class="filter-bar mb-16" id="wi-all-bar" style="display:none">
        <button class="btn btn-primary btn-sm" id="wi-load-all-btn">
          <i class="fa-solid fa-download"></i> Load All
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
        <button class="btn btn-secondary btn-sm" id="wi-export-btn">
          <i class="fa-solid fa-file-csv"></i> Export CSV
        </button>
        <button class="btn btn-primary btn-sm" id="wi-new-btn" style="display:none">
          <i class="fa-solid fa-plus"></i> New Work Item
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
    // Mode toggle
    container.querySelector('#wi-mode-single').addEventListener('click', () => {
      if (_mode === 'single') return;
      _mode = 'single';
      container.querySelector('#wi-mode-single').className = 'btn btn-primary btn-sm active';
      container.querySelector('#wi-mode-all').className = 'btn btn-secondary btn-sm';
      container.querySelector('#wi-single-bar').style.display = '';
      container.querySelector('#wi-all-bar').style.display = 'none';
      container.querySelector('#wi-filters').style.display = 'none';
      container.querySelector('#wi-content').innerHTML = '';
      _items = [];
      _allSummary = null;
    });

    container.querySelector('#wi-mode-all').addEventListener('click', () => {
      if (_mode === 'all') return;
      _mode = 'all';
      container.querySelector('#wi-mode-all').className = 'btn btn-primary btn-sm active';
      container.querySelector('#wi-mode-single').className = 'btn btn-secondary btn-sm';
      container.querySelector('#wi-single-bar').style.display = 'none';
      container.querySelector('#wi-all-bar').style.display = '';
      container.querySelector('#wi-filters').style.display = 'none';
      container.querySelector('#wi-content').innerHTML = '';
      _items = [];
      _allSummary = null;
    });

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
      if (e.target.closest('#wi-load-all-btn')) await _fetchAllItems(container);
      if (e.target.closest('#wi-apply-filter-btn')) _renderTable(container);
      if (e.target.closest('#wi-export-btn')) _exportCsv(_getCurrentFiltered(container));
      if (e.target.closest('#wi-new-btn')) _openCreateModal(container);
      if (e.target.closest('#wi-bulk-apply-btn')) _bulkApply(container);
      if (e.target.closest('#wi-bulk-export-btn')) _exportCsv(_getCheckedItems(container));
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
    container.querySelector('#wi-new-btn').style.display = '';
    _renderTable(container);
  }

  async function _fetchAllItems(container) {
    const activeConns = ConnectionsModule.getActive().filter(c => c.defaultProject);

    if (activeConns.length === 0) {
      container.querySelector('#wi-content').innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-plug-circle-xmark"></i>
          <p>No connections have a default project configured. <a href="#/connections" style="color:var(--color-primary)">Go to Connections</a> to set one.</p>
        </div>`;
      return;
    }

    _currentPage = 1;
    _context = null;

    const content = container.querySelector('#wi-content');
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading work items from all connections…</p></div>';
    container.querySelector('#wi-filters').style.display = 'none';

    const results = await Promise.allSettled(
      activeConns.map(conn =>
        AzureApi.getWorkItems(conn.orgUrl, conn.defaultProject, conn.pat, {})
          .then(res => ({ conn, res }))
      )
    );

    let allItems = [];
    let failCount = 0;

    for (const outcome of results) {
      if (outcome.status === 'rejected' || outcome.value.res.error) {
        failCount++;
        continue;
      }
      const { conn, res } = outcome.value;
      const items = (res.value || []).map(item => ({
        ...item,
        _connName: conn.name,
        _project: conn.defaultProject,
        _connId: conn.id,
      }));
      allItems = allItems.concat(items);
    }

    _items = allItems;
    _allSummary = { total: allItems.length, connCount: activeConns.length - failCount, failCount };
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
    const showOrgCol = _mode === 'all';

    let filtered = _items;
    if (search) filtered = filtered.filter(w =>
      (w.fields['System.Title'] || '').toLowerCase().includes(search)
    );

    // Apply type/state/priority filters
    const typeF     = container.querySelector('#wi-type-filter')?.value || '';
    const stateF    = container.querySelector('#wi-state-filter')?.value || '';
    const priorityF = container.querySelector('#wi-priority-filter')?.value || '';
    if (typeF)     filtered = filtered.filter(w => (w.fields['System.WorkItemType'] || '') === typeF);
    if (stateF)    filtered = filtered.filter(w => (w.fields['System.State'] || '') === stateF);
    if (priorityF) filtered = filtered.filter(w => String(w.fields['Microsoft.VSTS.Common.Priority'] || '') === priorityF);

    // Apply sorting
    if (_sortCol) {
      filtered = filtered.slice().sort((a, b) => {
        let va = _sortVal(a, _sortCol);
        let vb = _sortVal(b, _sortCol);
        if (va < vb) return _sortDir === 'asc' ? -1 : 1;
        if (va > vb) return _sortDir === 'asc' ? 1 : -1;
        return 0;
      });
    }

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

    const summaryBanner = (showOrgCol && _allSummary) ? `
      <div class="filter-bar mb-8" style="flex-wrap:wrap;gap:8px">
        <span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> ${_allSummary.total} item${_allSummary.total !== 1 ? 's' : ''} from ${_allSummary.connCount} connection${_allSummary.connCount !== 1 ? 's' : ''}</span>
        ${_allSummary.failCount ? `<span class="badge badge-warning"><i class="fa-solid fa-triangle-exclamation"></i> ${_allSummary.failCount} connection${_allSummary.failCount !== 1 ? 's' : ''} failed</span>` : ''}
      </div>` : '';

    // Column count for group headers colspan
    // checkbox + id + (org + project if all) + title + type + state + assigned + priority + created + updated + edit
    const totalCols = showOrgCol ? 12 : 10;

    const _th = (key, label) => {
      const active = _sortCol === key;
      const icon = active ? (_sortDir === 'asc' ? ' ▲' : ' ▼') : '';
      return `<th class="sortable-col${active ? ' sort-active' : ''}" data-sort="${key}" style="cursor:pointer;user-select:none">${label}${icon}<span class="resize-handle" data-resize="${key}"></span></th>`;
    };

    const orgCols = showOrgCol ? `${_th('org','Organization')}${_th('project','Project')}` : '';

    content.innerHTML = `
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table id="wi-table">
          <thead>
            <tr>
              <th style="width:32px"><input type="checkbox" id="wi-select-all" title="Select all" /></th>
              ${_th('id','ID')}${orgCols}${_th('title','Title')}${_th('type','Type')}${_th('state','State')}
              ${_th('assigned','Assigned To')}${_th('priority','Priority')}
              ${_th('created','Created')}${_th('updated','Updated')}
              <th>Edit</th>
            </tr>
          </thead>
          <tbody id="wi-tbody">
            ${pageItems.map(w => _rowHtml(w, showOrgCol)).join('')}
          </tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}
      <div id="wi-bulk-bar" class="bulk-bar" style="display:none">
        <span id="wi-bulk-count" class="text-sm"></span>
        <select id="wi-bulk-state">
          <option value="">Change State…</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
          <option value="Done">Done</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-bulk-apply-btn">Apply</button>
        <button class="btn btn-secondary btn-sm" id="wi-bulk-export-btn"><i class="fa-solid fa-file-csv"></i> Export Selected</button>
      </div>
    `;

    // Sort header click
    content.querySelectorAll('.sortable-col').forEach(th => {
      th.addEventListener('click', e => {
        if (e.target.classList.contains('resize-handle')) return;
        const col = th.dataset.sort;
        if (_sortCol === col) {
          _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          _sortCol = col;
          _sortDir = 'asc';
        }
        _renderTable(container);
      });
    });

    // Select-all checkbox
    const selectAll = content.querySelector('#wi-select-all');
    if (selectAll) {
      selectAll.addEventListener('change', () => {
        content.querySelectorAll('.wi-row-check').forEach(cb => { cb.checked = selectAll.checked; });
        _updateBulkBar(content);
      });
    }

    // Row checkboxes
    content.querySelectorAll('.wi-row-check').forEach(cb => {
      cb.addEventListener('change', () => _updateBulkBar(content));
    });

    // Row click → open detail modal (skip checkbox and edit columns)
    content.querySelectorAll('.wi-row').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.wi-row-check') || e.target.closest('.wi-edit-btn')) return;
        _openDetail(row.dataset.id, row.dataset.connId);
      });
      row.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          _openDetail(row.dataset.id, row.dataset.connId);
        }
      });
    });

    // Edit button
    content.querySelectorAll('.wi-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        _openEditModal(btn.dataset.id, btn.dataset.connId, container);
      });
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });

    _initResizableColumns(content.querySelector('#wi-table'));
  }

  function _rowHtml(w, showOrgCol) {
    const f       = w.fields;
    const id      = f['System.Id']   || w.id;
    const title   = f['System.Title']|| '(no title)';
    const type    = f['System.WorkItemType'] || '';
    const state   = f['System.State'] || '';
    const assigned= (f['System.AssignedTo']?.displayName) || f['System.AssignedTo'] || '—';
    const priority= f['Microsoft.VSTS.Common.Priority'] || '—';
    const created = f['System.CreatedDate'] ? _fmtDate(f['System.CreatedDate']) : '—';
    const updated = f['System.ChangedDate'] ? _fmtDate(f['System.ChangedDate']) : '—';

    const orgCells = showOrgCol
      ? `<td class="text-sm">${_esc(w._connName || '')}</td><td class="text-sm">${_esc(w._project || '')}</td>`
      : '';

    const connIdAttr = w._connId ? ` data-conn-id="${_esc(w._connId)}"` : '';

    return `
      <tr class="wi-row" data-id="${id}"${connIdAttr} tabindex="0" style="cursor:pointer">
        <td onclick="event.stopPropagation()"><input type="checkbox" class="wi-row-check" data-id="${id}" /></td>
        <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
        ${orgCells}
        <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
        <td>${_typeBadge(type)}</td>
        <td>${_stateBadge(state)}</td>
        <td class="truncate" style="max-width:140px">${_esc(String(assigned))}</td>
        <td>${_priorityBadge(priority)}</td>
        <td class="text-muted text-sm">${created}</td>
        <td class="text-muted text-sm">${updated}</td>
        <td onclick="event.stopPropagation()"><button class="btn-icon wi-edit-btn" data-id="${id}" data-conn-id="${_esc(w._connId || '')}" title="Edit"><i class="fa-solid fa-pen"></i></button></td>
      </tr>`;
  }

  async function _openDetail(id, connIdOverride) {
    // In "all" mode, the work item row carries a connId; otherwise fall back to _context
    let conn;
    if (connIdOverride) {
      conn = ConnectionsModule.getById(connIdOverride);
    }
    if (!conn && _context) {
      conn = _context.conn;
    }
    if (!conn) {
      // Fall back: find the item in _items and use its stored _connId
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      if (item && item._connId) {
        conn = ConnectionsModule.getById(item._connId) || null;
      }
    }
    if (!conn) return;

    App.openModal('Work Item #' + id, '<div class="loading-state"><div class="spinner"></div><p>Loading…</p></div>');

    const result = await AzureApi.getWorkItemDetail(conn.orgUrl, id, conn.pat);
    if (result.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i><p>${_esc(result.message)}</p></div>`);
      return;
    }

    const f = result.fields;
    const detailHtml = `
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
      ${result._links?.html?.href ? `<div class="detail-row"><span class="detail-label">Link</span><span class="detail-value"><a href="${_esc(result._links.html.href)}" target="_blank" rel="noopener">Open in Azure DevOps <i class="fa-solid fa-arrow-up-right-from-square"></i></a></span></div>` : ''}
      ${f['System.Description'] ? `
      <div class="detail-row" style="flex-direction:column;gap:6px">
        <span class="detail-label">Description</span>
        <div style="font-size:.82rem;border:1px solid var(--border-color);padding:10px;border-radius:var(--radius-sm);background:var(--bg-table-alt);">
          ${f['System.Description']}
        </div>
      </div>` : ''}
      <div id="wi-comments-section" style="margin-top:12px">
        <div class="section-title" style="font-size:.82rem;margin:8px 0 4px">Comments</div>
        <div id="wi-comments-list"><div class="spinner" style="width:14px;height:14px;border-width:2px;vertical-align:middle;display:inline-block"></div> Loading comments…</div>
        <div style="margin-top:8px;display:flex;gap:8px;align-items:flex-start;flex-direction:column">
          <textarea id="wi-new-comment" rows="2" style="width:100%;resize:vertical;font-size:.82rem;padding:6px;border:1px solid var(--border-color);border-radius:var(--radius-sm);background:var(--bg-input);color:var(--text-primary)" placeholder="Add a comment…"></textarea>
          <button class="btn btn-primary btn-sm" id="wi-add-comment-btn" data-id="${result.id}" data-conn-org="${_esc(conn.orgUrl)}" data-conn-pat="${_esc(conn.pat)}">Post Comment</button>
        </div>
      </div>
    `;
    App.updateModalBody(detailHtml);

    // Load comments async
    AzureApi.getWorkItemComments(conn.orgUrl, result.id, conn.pat).then(cr => {
      const listEl = document.getElementById('wi-comments-list');
      if (!listEl) return;
      if (cr.error || !cr.value || cr.value.length === 0) {
        listEl.innerHTML = '<span class="text-muted text-sm">No comments yet.</span>';
        return;
      }
      listEl.innerHTML = cr.value.map(c => `
        <div style="border-bottom:1px solid var(--border-color);padding:6px 0;font-size:.82rem">
          <strong>${_esc(c.createdBy?.displayName || '—')}</strong>
          <span class="text-muted text-sm" style="margin-left:8px">${c.createdDate ? new Date(c.createdDate).toLocaleString() : ''}</span>
          <div style="margin-top:4px">${c.text || ''}</div>
        </div>`).join('');
    });

    // Post comment handler
    setTimeout(() => {
      const btn = document.getElementById('wi-add-comment-btn');
      if (!btn) return;
      btn.addEventListener('click', async () => {
        const text = document.getElementById('wi-new-comment')?.value?.trim();
        if (!text) return;
        btn.disabled = true;
        const res = await AzureApi.addWorkItemComment(conn.orgUrl, result.id, conn.pat, text);
        btn.disabled = false;
        if (res.error) { App.showToast('Failed to post comment: ' + res.message, 'error'); return; }
        App.showToast('Comment posted.', 'success');
        document.getElementById('wi-new-comment').value = '';
        const cr2 = await AzureApi.getWorkItemComments(conn.orgUrl, result.id, conn.pat);
        const listEl2 = document.getElementById('wi-comments-list');
        if (listEl2 && !cr2.error && cr2.value?.length) {
          listEl2.innerHTML = cr2.value.map(c => `
            <div style="border-bottom:1px solid var(--border-color);padding:6px 0;font-size:.82rem">
              <strong>${_esc(c.createdBy?.displayName || '—')}</strong>
              <span class="text-muted text-sm" style="margin-left:8px">${c.createdDate ? new Date(c.createdDate).toLocaleString() : ''}</span>
              <div style="margin-top:4px">${c.text || ''}</div>
            </div>`).join('');
        }
      });
    }, 0);
  }

  function _sortVal(w, col) {
    const f = w.fields;
    switch (col) {
      case 'id':       return f['System.Id'] || 0;
      case 'title':    return (f['System.Title'] || '').toLowerCase();
      case 'type':     return (f['System.WorkItemType'] || '').toLowerCase();
      case 'state':    return (f['System.State'] || '').toLowerCase();
      case 'assigned': return (f['System.AssignedTo']?.displayName || String(f['System.AssignedTo'] || '')).toLowerCase();
      case 'priority': return f['Microsoft.VSTS.Common.Priority'] || 99;
      case 'created':  return f['System.CreatedDate'] || '';
      case 'updated':  return f['System.ChangedDate'] || '';
      case 'org':      return (w._connName || '').toLowerCase();
      case 'project':  return (w._project || '').toLowerCase();
      default:         return '';
    }
  }

  function _updateBulkBar(content) {
    const checked = content.querySelectorAll('.wi-row-check:checked');
    const bar = content.querySelector('#wi-bulk-bar');
    const countEl = content.querySelector('#wi-bulk-count');
    if (!bar) return;
    if (checked.length > 0) {
      bar.style.display = 'flex';
      if (countEl) countEl.textContent = `${checked.length} item${checked.length !== 1 ? 's' : ''} selected`;
    } else {
      bar.style.display = 'none';
    }
  }

  function _getCheckedItems(container) {
    const content = container.querySelector('#wi-content');
    const checked = content ? Array.from(content.querySelectorAll('.wi-row-check:checked')) : [];
    const ids = new Set(checked.map(cb => String(cb.dataset.id)));
    return _items.filter(w => ids.has(String(w.fields['System.Id'] || w.id)));
  }

  function _getCurrentFiltered(container) {
    const search    = (container.querySelector('#wi-search')?.value || '').toLowerCase();
    const typeF     = container.querySelector('#wi-type-filter')?.value || '';
    const stateF    = container.querySelector('#wi-state-filter')?.value || '';
    const priorityF = container.querySelector('#wi-priority-filter')?.value || '';
    let filtered = _items;
    if (search)    filtered = filtered.filter(w => (w.fields['System.Title'] || '').toLowerCase().includes(search));
    if (typeF)     filtered = filtered.filter(w => (w.fields['System.WorkItemType'] || '') === typeF);
    if (stateF)    filtered = filtered.filter(w => (w.fields['System.State'] || '') === stateF);
    if (priorityF) filtered = filtered.filter(w => String(w.fields['Microsoft.VSTS.Common.Priority'] || '') === priorityF);
    return filtered;
  }

  function _exportCsv(items) {
    if (!items || items.length === 0) { App.showToast('No items to export.', 'info'); return; }
    const headers = ['ID','Title','Type','State','Assigned To','Priority','Created','Updated','Organization','Project'];
    const rows = items.map(w => {
      const f = w.fields;
      return [
        f['System.Id'] || w.id,
        f['System.Title'] || '',
        f['System.WorkItemType'] || '',
        f['System.State'] || '',
        f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || '',
        f['Microsoft.VSTS.Common.Priority'] || '',
        f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toLocaleDateString() : '',
        f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toLocaleDateString() : '',
        w._connName || '',
        w._project || '',
      ].map(v => `"${String(v).replace(/"/g,'""')}"`).join(',');
    });
    const csv = [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `workitems_${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    App.showToast(`Exported ${items.length} items to CSV.`, 'success');
  }

  async function _bulkApply(container) {
    const content = container.querySelector('#wi-content');
    const checked = content ? Array.from(content.querySelectorAll('.wi-row-check:checked')) : [];
    const state = content?.querySelector('#wi-bulk-state')?.value;
    if (!state) { App.showToast('Select a state to apply.', 'warning'); return; }
    if (checked.length === 0) { App.showToast('No items selected.', 'warning'); return; }

    const btn = content.querySelector('#wi-bulk-apply-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Applying…'; }

    let successCount = 0;
    for (const cb of checked) {
      const id = cb.dataset.id;
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      const connId = item?._connId || _context?.connId;
      const conn = connId ? ConnectionsModule.getById(connId) : _context?.conn;
      if (!conn) continue;
      const res = await AzureApi.updateWorkItemState(conn.orgUrl, id, conn.pat, state);
      if (!res.error) {
        successCount++;
        if (item) item.fields['System.State'] = state;
      }
    }

    if (btn) { btn.disabled = false; btn.textContent = 'Apply'; }
    App.showToast(`Updated ${successCount} of ${checked.length} items to "${state}".`, successCount > 0 ? 'success' : 'error');
    _renderTable(container);
  }

  function _openCreateModal(container) {
    if (!_context) { App.showToast('Select a project first.', 'warning'); return; }
    App.openModal('New Work Item', `
      <div class="form-group">
        <label>Type</label>
        <select id="wi-create-type">
          <option value="Task">Task</option>
          <option value="Bug">Bug</option>
          <option value="User Story">User Story</option>
          <option value="Feature">Feature</option>
          <option value="Epic">Epic</option>
          <option value="Issue">Issue</option>
        </select>
      </div>
      <div class="form-group">
        <label>Title <span style="color:var(--color-danger)">*</span></label>
        <input type="text" id="wi-create-title" placeholder="Work item title…" />
      </div>
      <div class="form-group">
        <label>Description</label>
        <textarea id="wi-create-desc" rows="3" style="width:100%;resize:vertical;padding:6px;font-size:.85rem;border:1px solid var(--border-color);border-radius:var(--radius-sm);background:var(--bg-input);color:var(--text-primary)" placeholder="Optional description…"></textarea>
      </div>
      <div style="display:flex;gap:8px;margin-top:12px">
        <button class="btn btn-primary btn-sm" id="wi-create-submit-btn">Create</button>
        <button class="btn btn-secondary btn-sm" onclick="App.closeModal()">Cancel</button>
      </div>
      <div id="wi-create-error" class="form-error hidden" style="margin-top:8px"></div>
    `);
    setTimeout(() => {
      const btn = document.getElementById('wi-create-submit-btn');
      if (!btn) return;
      btn.addEventListener('click', async () => {
        const title = document.getElementById('wi-create-title')?.value?.trim();
        const type  = document.getElementById('wi-create-type')?.value;
        const desc  = document.getElementById('wi-create-desc')?.value?.trim();
        const errEl = document.getElementById('wi-create-error');
        if (!title) { if (errEl) { errEl.textContent = 'Title is required.'; errEl.classList.remove('hidden'); } return; }
        btn.disabled = true; btn.textContent = 'Creating…';
        const fields = { '/fields/System.Title': title };
        if (desc) fields['/fields/System.Description'] = desc;
        const res = await AzureApi.createWorkItem(_context.conn.orgUrl, _context.project, _context.conn.pat, type, fields);
        btn.disabled = false; btn.textContent = 'Create';
        if (res.error) {
          if (errEl) { errEl.textContent = res.message; errEl.classList.remove('hidden'); }
          return;
        }
        App.closeModal();
        App.showToast(`Work item #${res.id} created successfully.`, 'success');
        await _fetchItems(container);
      });
    }, 0);
  }

  function _openEditModal(id, connIdOverride, container) {
    const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
    if (!item) return;
    const connId = connIdOverride || item._connId || _context?.connId;
    const conn = connId ? ConnectionsModule.getById(connId) : _context?.conn;
    if (!conn) return;
    const f = item.fields;
    App.openModal(`Edit Work Item #${id}`, `
      <div class="form-group">
        <label>Title</label>
        <input type="text" id="wi-edit-title" value="${_esc(f['System.Title'] || '')}" />
      </div>
      <div class="form-group">
        <label>State</label>
        <select id="wi-edit-state">
          ${['New','Active','Resolved','Closed','Done'].map(s => `<option value="${s}"${f['System.State'] === s ? ' selected' : ''}>${s}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Priority</label>
        <select id="wi-edit-priority">
          ${[1,2,3,4].map(p => `<option value="${p}"${f['Microsoft.VSTS.Common.Priority'] == p ? ' selected' : ''}>${p}</option>`).join('')}
        </select>
      </div>
      <div style="display:flex;gap:8px;margin-top:12px">
        <button class="btn btn-primary btn-sm" id="wi-edit-save-btn">Save</button>
        <button class="btn btn-secondary btn-sm" onclick="App.closeModal()">Cancel</button>
      </div>
      <div id="wi-edit-error" class="form-error hidden" style="margin-top:8px"></div>
    `);
    setTimeout(() => {
      const btn = document.getElementById('wi-edit-save-btn');
      if (!btn) return;
      btn.addEventListener('click', async () => {
        const title    = document.getElementById('wi-edit-title')?.value?.trim();
        const state    = document.getElementById('wi-edit-state')?.value;
        const priority = document.getElementById('wi-edit-priority')?.value;
        const errEl    = document.getElementById('wi-edit-error');
        if (!title) { if (errEl) { errEl.textContent = 'Title is required.'; errEl.classList.remove('hidden'); } return; }
        btn.disabled = true; btn.textContent = 'Saving…';
        const fields = {
          '/fields/System.Title':    title,
          '/fields/System.State':    state,
          '/fields/Microsoft.VSTS.Common.Priority': parseInt(priority, 10),
        };
        const res = await AzureApi.updateWorkItemFields(conn.orgUrl, id, conn.pat, fields);
        btn.disabled = false; btn.textContent = 'Save';
        if (res.error) {
          if (errEl) { errEl.textContent = res.message; errEl.classList.remove('hidden'); }
          return;
        }
        // Update local item
        if (item) {
          item.fields['System.Title'] = title;
          item.fields['System.State'] = state;
          item.fields['Microsoft.VSTS.Common.Priority'] = parseInt(priority, 10);
        }
        App.closeModal();
        App.showToast('Work item updated.', 'success');
        _renderTable(container);
      });
    }, 0);
  }

  function _initResizableColumns(table) {
    if (!table) return;
    table.querySelectorAll('.resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', e => {
        e.stopPropagation();
        const th = handle.closest('th');
        if (!th) return;
        const startX = e.pageX;
        const startW = th.offsetWidth;
        function onMove(ev) { th.style.width = Math.max(40, startW + ev.pageX - startX) + 'px'; }
        function onUp() { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });
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
