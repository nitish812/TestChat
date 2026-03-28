/**
 * workitems.js — Work Items
 *
 * Displays work items for a selected project with filtering,
 * pagination (50 per page), and a detail modal.
 * Supports "Single Project" and "All Connections" modes.
 *
 * Features:
 *  - Collapsible org groups (All Connections mode)
 *  - Sortable columns (_sortCol / _sortDir)
 *  - Resizable columns
 *  - Keyboard navigation (ArrowUp/Down, Enter, Escape)
 *  - Breadcrumb navigation
 *  - Bulk actions (select-all, state change, export)
 *  - Export CSV
 *  - Inline editing (State, AssignedTo)
 *  - Work item comments in detail modal
 *  - Create new work item
 *  - Linked work items in detail modal
 */

const WorkItemsModule = (() => {
  const PAGE_SIZE = 50;

  let _items       = [];
  let _currentPage = 1;
  let _context     = null;   // { connId, project, conn }
  let _mode        = 'single'; // 'single' | 'all'
  let _allSummary  = null;   // { total, connCount, failCount }
  let _sortCol     = null;
  let _sortDir     = 'asc';
  let _collapsedGroups = new Set();

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];
    _mode = 'single';
    _allSummary = null;
    _sortCol = null;
    _sortDir = 'asc';
    _collapsedGroups = new Set();

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
        <button class="btn btn-secondary btn-sm" id="wi-export-csv-btn" title="Export to CSV">
          <i class="fa-solid fa-file-csv"></i> Export CSV
        </button>
        <button class="btn btn-primary btn-sm" id="wi-new-btn" title="Create new work item">
          <i class="fa-solid fa-plus"></i> New Work Item
        </button>
      </div>

      <!-- Breadcrumb -->
      <div id="wi-breadcrumb" class="breadcrumb mb-8" style="display:none"></div>

      <!-- Bulk toolbar -->
      <div id="wi-bulk-toolbar" class="bulk-toolbar" style="display:none">
        <span id="wi-bulk-count" class="text-sm"></span>
        <select id="wi-bulk-state">
          <option value="">Change state to…</option>
          <option value="New">New</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
          <option value="Done">Done</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-bulk-state-btn"><i class="fa-solid fa-pen"></i> Apply State</button>
        <button class="btn btn-secondary btn-sm" id="wi-bulk-export-btn"><i class="fa-solid fa-file-csv"></i> Export Selected</button>
        <button class="btn btn-secondary btn-sm" id="wi-bulk-clear-btn"><i class="fa-solid fa-xmark"></i> Clear</button>
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
    else if (_mode === 'all') await _fetchAllItems(container);
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
      container.querySelector('#wi-breadcrumb').style.display = 'none';
      container.querySelector('#wi-content').innerHTML = '';
      _items = [];
      _allSummary = null;
      _context = null;
    });

    container.querySelector('#wi-mode-all').addEventListener('click', () => {
      if (_mode === 'all') return;
      _mode = 'all';
      container.querySelector('#wi-mode-all').className = 'btn btn-primary btn-sm active';
      container.querySelector('#wi-mode-single').className = 'btn btn-secondary btn-sm';
      container.querySelector('#wi-single-bar').style.display = 'none';
      container.querySelector('#wi-all-bar').style.display = '';
      container.querySelector('#wi-filters').style.display = 'none';
      container.querySelector('#wi-breadcrumb').style.display = 'none';
      container.querySelector('#wi-content').innerHTML = '';
      _items = [];
      _allSummary = null;
      _context = null;
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
      if (e.target.closest('#wi-load-btn'))          await _fetchItems(container);
      if (e.target.closest('#wi-load-all-btn'))      await _fetchAllItems(container);
      if (e.target.closest('#wi-apply-filter-btn'))  _renderTable(container);
      if (e.target.closest('#wi-export-csv-btn'))    _exportCsv();
      if (e.target.closest('#wi-new-btn'))           _openCreateModal(container);
      if (e.target.closest('#wi-bulk-state-btn'))    await _bulkChangeState(container);
      if (e.target.closest('#wi-bulk-export-btn'))   _exportCsv(true);
      if (e.target.closest('#wi-bulk-clear-btn'))    _clearBulk(container);

      // Group header toggle
      const groupHdr = e.target.closest('.group-header-row');
      if (groupHdr) {
        const org = groupHdr.dataset.org;
        if (_collapsedGroups.has(org)) { _collapsedGroups.delete(org); } else { _collapsedGroups.add(org); }
        _renderTable(container);
        return;
      }

      // Inline edit
      const editBtn = e.target.closest('.wi-inline-edit-btn');
      if (editBtn) {
        _startInlineEdit(container, editBtn.closest('tr'));
        return;
      }
      const saveBtn = e.target.closest('.wi-inline-save-btn');
      if (saveBtn) {
        await _saveInlineEdit(container, saveBtn.closest('tr'));
        return;
      }
      const cancelBtn = e.target.closest('.wi-inline-cancel-btn');
      if (cancelBtn) {
        _cancelInlineEdit(container, cancelBtn.closest('tr'));
        return;
      }

      // Row click → open detail modal (skip if clicking controls)
      const row = e.target.closest('.wi-row');
      if (row && !e.target.closest('button') && !e.target.closest('input') && !e.target.closest('select')) {
        _openDetail(row.dataset.id, row.dataset.connId);
      }
    });

    container.addEventListener('keydown', e => {
      if (e.key === 'Enter' && e.target.id === 'wi-search') _renderTable(container);
    });

    // Select-all checkbox
    container.addEventListener('change', e => {
      if (e.target.id === 'wi-select-all') {
        const checked = e.target.checked;
        container.querySelectorAll('.wi-row-cb').forEach(cb => { cb.checked = checked; });
        _updateBulkToolbar(container);
      }
      if (e.target.classList.contains('wi-row-cb')) {
        _updateBulkToolbar(container);
      }
    });

    // Table keyboard navigation
    container.addEventListener('keydown', e => {
      const table = container.querySelector('table');
      if (!table) return;
      const focused = table.querySelector('.row-focused');
      if (!focused && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        const first = table.querySelector('tbody .wi-row');
        if (first) { first.classList.add('row-focused'); first.focus(); }
        e.preventDefault();
        return;
      }
      if (!focused) return;
      if (e.key === 'ArrowDown') {
        const next = focused.nextElementSibling;
        if (next && next.classList.contains('wi-row')) {
          focused.classList.remove('row-focused');
          next.classList.add('row-focused');
          next.focus();
        }
        e.preventDefault();
      } else if (e.key === 'ArrowUp') {
        const prev = focused.previousElementSibling;
        if (prev && prev.classList.contains('wi-row')) {
          focused.classList.remove('row-focused');
          prev.classList.add('row-focused');
          prev.focus();
        }
        e.preventDefault();
      } else if (e.key === 'Enter') {
        if (focused) _openDetail(focused.dataset.id, focused.dataset.connId);
        e.preventDefault();
      }
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
    container.querySelector('#wi-breadcrumb').style.display = 'none';

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
    _renderBreadcrumb(container, conn.name, project);
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
    container.querySelector('#wi-breadcrumb').style.display = 'none';

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
    _renderBreadcrumb(container, 'All Connections', null);
    _renderTable(container);
  }

  function _getFilters(container) {
    return {
      type:     container.querySelector('#wi-type-filter')?.value || '',
      state:    container.querySelector('#wi-state-filter')?.value || '',
      priority: container.querySelector('#wi-priority-filter')?.value || '',
    };
  }

  function _renderBreadcrumb(container, orgName, projectName) {
    const bc = container.querySelector('#wi-breadcrumb');
    if (!bc) return;
    let parts = `<a href="#/connections" class="bc-link">All Connections</a>`;
    if (orgName) parts += ` <span class="bc-sep">›</span> <span class="bc-cur">${_esc(orgName)}</span>`;
    if (projectName) parts += ` <span class="bc-sep">›</span> <span class="bc-cur">${_esc(projectName)}</span>`;
    bc.innerHTML = parts;
    bc.style.display = '';
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

    // Sort
    if (_sortCol) filtered = _sortItems(filtered, _sortCol, _sortDir);

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

    const orgCols = showOrgCol ? `
      <th class="sortable-th" data-col="org" style="cursor:pointer">Organization ${_sortIndicator('org')}</th>
      <th class="sortable-th" data-col="project" style="cursor:pointer">Project ${_sortIndicator('project')}</th>` : '';

    const tableId = 'wi-table-' + Date.now();

    let tableBody = '';
    if (showOrgCol) {
      // Group by org
      const groups = {};
      pageItems.forEach(w => {
        const key = w._connName || '(unknown)';
        if (!groups[key]) groups[key] = [];
        groups[key].push(w);
      });

      const colCount = showOrgCol ? 10 : 8;
      for (const [org, items] of Object.entries(groups)) {
        const collapsed = _collapsedGroups.has(org);
        tableBody += `
          <tr class="group-header-row" data-org="${_esc(org)}">
            <td colspan="${colCount}" style="font-weight:700;background:var(--bg-table-alt);padding:8px 14px;cursor:pointer">
              <i class="fa-solid fa-chevron-${collapsed ? 'right' : 'down'}" style="margin-right:6px;font-size:.75rem"></i>
              <i class="fa-brands fa-windows" style="margin-right:4px;color:var(--color-primary)"></i>
              ${_esc(org)} <span class="badge badge-secondary ml-4">${items.length}</span>
            </td>
          </tr>`;
        if (!collapsed) {
          items.forEach(w => { tableBody += _rowHtml(w, showOrgCol); });
        }
      }
    } else {
      pageItems.forEach(w => { tableBody += _rowHtml(w, showOrgCol); });
    }

    content.innerHTML = `
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table id="${tableId}" tabindex="0">
          <thead>
            <tr>
              <th><input type="checkbox" id="wi-select-all" title="Select all" /></th>
              <th class="sortable-th" data-col="id" style="cursor:pointer">ID ${_sortIndicator('id')}</th>
              ${orgCols}
              <th class="sortable-th" data-col="title" style="cursor:pointer">Title ${_sortIndicator('title')}</th>
              <th class="sortable-th" data-col="type" style="cursor:pointer">Type ${_sortIndicator('type')}</th>
              <th class="sortable-th" data-col="state" style="cursor:pointer">State ${_sortIndicator('state')}</th>
              <th class="sortable-th" data-col="assigned" style="cursor:pointer">Assigned To ${_sortIndicator('assigned')}</th>
              <th class="sortable-th" data-col="priority" style="cursor:pointer">Priority ${_sortIndicator('priority')}</th>
              <th class="sortable-th" data-col="created" style="cursor:pointer">Created ${_sortIndicator('created')}</th>
              <th class="sortable-th" data-col="updated" style="cursor:pointer">Updated ${_sortIndicator('updated')}</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="wi-tbody">
            ${tableBody}
          </tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}
    `;

    // Sortable column headers
    content.querySelectorAll('.sortable-th').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.col;
        if (_sortCol === col) { _sortDir = _sortDir === 'asc' ? 'desc' : 'asc'; }
        else { _sortCol = col; _sortDir = 'asc'; }
        _renderTable(container);
      });
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });

    // Resizable columns
    const table = content.querySelector('table');
    if (table) _initResizableColumns(table);

    _updateBulkToolbar(container);
  }

  function _sortItems(items, col, dir) {
    const mult = dir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let va, vb;
      const fa = a.fields, fb = b.fields;
      switch (col) {
        case 'id':       va = fa['System.Id'] || a.id || 0; vb = fb['System.Id'] || b.id || 0; return mult * (va - vb);
        case 'title':    va = (fa['System.Title'] || '').toLowerCase(); vb = (fb['System.Title'] || '').toLowerCase(); break;
        case 'type':     va = fa['System.WorkItemType'] || ''; vb = fb['System.WorkItemType'] || ''; break;
        case 'state':    va = fa['System.State'] || ''; vb = fb['System.State'] || ''; break;
        case 'assigned': va = String(fa['System.AssignedTo']?.displayName || fa['System.AssignedTo'] || ''); vb = String(fb['System.AssignedTo']?.displayName || fb['System.AssignedTo'] || ''); break;
        case 'priority': va = fa['Microsoft.VSTS.Common.Priority'] || 99; vb = fb['Microsoft.VSTS.Common.Priority'] || 99; return mult * (va - vb);
        case 'created':  va = fa['System.CreatedDate'] || ''; vb = fb['System.CreatedDate'] || ''; break;
        case 'updated':  va = fa['System.ChangedDate'] || ''; vb = fb['System.ChangedDate'] || ''; break;
        case 'org':      va = a._connName || ''; vb = b._connName || ''; break;
        case 'project':  va = a._project || ''; vb = b._project || ''; break;
        default: return 0;
      }
      if (va < vb) return -1 * mult;
      if (va > vb) return 1 * mult;
      return 0;
    });
  }

  function _sortIndicator(col) {
    if (_sortCol !== col) return '<span class="sort-indicator text-muted">⇅</span>';
    return _sortDir === 'asc'
      ? '<span class="sort-indicator" style="color:var(--color-primary)">▲</span>'
      : '<span class="sort-indicator" style="color:var(--color-primary)">▼</span>';
  }

  function _initResizableColumns(table) {
    const ths = table.querySelectorAll('th');
    ths.forEach(th => {
      if (th.querySelector('.col-resize-handle')) return;
      const handle = document.createElement('span');
      handle.className = 'col-resize-handle';
      handle.addEventListener('mousedown', e => {
        e.preventDefault();
        const startX = e.pageX;
        const startW = th.offsetWidth;
        function onMove(ev) { th.style.width = Math.max(40, startW + ev.pageX - startX) + 'px'; }
        function onUp() { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
      th.style.position = 'relative';
      th.appendChild(handle);
    });
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
      <tr class="wi-row" data-id="${id}"${connIdAttr} tabindex="0">
        <td><input type="checkbox" class="wi-row-cb" data-id="${id}" /></td>
        <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
        ${orgCells}
        <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
        <td data-field="type">${_typeBadge(type)}</td>
        <td data-field="state">${_stateBadge(state)}</td>
        <td class="truncate" style="max-width:140px" data-field="assigned">${_esc(String(assigned))}</td>
        <td>${_priorityBadge(priority)}</td>
        <td class="text-muted text-sm">${created}</td>
        <td class="text-muted text-sm">${updated}</td>
        <td>
          <button class="btn-icon wi-inline-edit-btn" title="Edit" data-id="${id}">
            <i class="fa-solid fa-pen-to-square"></i>
          </button>
        </td>
      </tr>`;
  }

  // ─── Inline editing ────────────────────────────────────────────

  function _startInlineEdit(container, row) {
    if (row.classList.contains('editing')) return;
    row.classList.add('editing');

    const stateCell    = row.querySelector('td[data-field="state"]');
    const assignedCell = row.querySelector('td[data-field="assigned"]');
    const currentState    = stateCell.querySelector('.badge')?.textContent?.trim() || '';
    const currentAssigned = assignedCell.textContent?.trim() || '';

    stateCell.innerHTML = `
      <select class="inline-edit-select" style="font-size:.78rem;padding:2px 4px">
        ${['New','Active','Resolved','Closed','Done'].map(s =>
          `<option value="${s}" ${s === currentState ? 'selected' : ''}>${s}</option>`
        ).join('')}
      </select>`;
    assignedCell.innerHTML = `<input type="text" class="inline-edit-input" value="${_esc(currentAssigned)}" style="font-size:.78rem;padding:2px 4px;width:120px" />`;

    const actCell = row.querySelector('.wi-inline-edit-btn').parentElement;
    actCell.innerHTML = `
      <button class="btn btn-sm btn-primary wi-inline-save-btn" data-id="${row.dataset.id}">Save</button>
      <button class="btn btn-sm btn-secondary wi-inline-cancel-btn" style="margin-left:4px">Cancel</button>
    `;
  }

  async function _saveInlineEdit(container, row) {
    const id = row.dataset.id;
    const connId = row.dataset.connId || _context?.connId;
    const conn = connId ? ConnectionsModule.getById(connId) : _context?.conn;
    if (!conn) { App.showToast('Cannot determine connection for this item.', 'error'); return; }

    const newState    = row.querySelector('.inline-edit-select')?.value || '';
    const newAssigned = row.querySelector('.inline-edit-input')?.value?.trim() || '';

    const fields = {};
    if (newState)    fields['System.State'] = newState;
    if (newAssigned) fields['System.AssignedTo'] = newAssigned;

    const result = await AzureApi.updateWorkItemFields(conn.orgUrl, id, fields, conn.pat);
    if (result.error) { App.showToast('Failed to save: ' + result.message, 'error'); return; }

    // Update in-memory item
    const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
    if (item) {
      if (newState)    item.fields['System.State'] = newState;
      if (newAssigned) item.fields['System.AssignedTo'] = { displayName: newAssigned };
    }

    App.showToast('Work item updated.', 'success');
    _renderTable(container);
  }

  function _cancelInlineEdit(container, row) {
    row.classList.remove('editing');
    _renderTable(container);
  }

  // ─── Bulk actions ──────────────────────────────────────────────

  function _updateBulkToolbar(container) {
    const checked = container.querySelectorAll('.wi-row-cb:checked');
    const toolbar = container.querySelector('#wi-bulk-toolbar');
    if (!toolbar) return;
    if (checked.length > 0) {
      toolbar.style.display = '';
      const cnt = container.querySelector('#wi-bulk-count');
      if (cnt) cnt.textContent = `${checked.length} selected`;
    } else {
      toolbar.style.display = 'none';
    }
  }

  async function _bulkChangeState(container) {
    const newState = container.querySelector('#wi-bulk-state')?.value;
    if (!newState) { App.showToast('Select a state first.', 'warning'); return; }
    const checkedIds = [...container.querySelectorAll('.wi-row-cb:checked')].map(cb => cb.dataset.id);
    if (!checkedIds.length) return;

    let done = 0;
    for (const id of checkedIds) {
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      const connId = item?._connId || _context?.connId;
      const conn = connId ? ConnectionsModule.getById(connId) : _context?.conn;
      if (!conn) continue;
      const res = await AzureApi.updateWorkItemState(conn.orgUrl, id, newState, conn.pat);
      if (!res.error) {
        if (item) item.fields['System.State'] = newState;
        done++;
      }
    }
    App.showToast(`Updated ${done} of ${checkedIds.length} item(s).`, done === checkedIds.length ? 'success' : 'warning');
    _renderTable(container);
  }

  function _clearBulk(container) {
    container.querySelectorAll('.wi-row-cb').forEach(cb => { cb.checked = false; });
    const all = container.querySelector('#wi-select-all');
    if (all) all.checked = false;
    _updateBulkToolbar(container);
  }

  // ─── Export CSV ────────────────────────────────────────────────

  function _exportCsv(selectedOnly = false) {
    let items = _items;
    if (selectedOnly) {
      const content = document.getElementById('wi-content') || document.querySelector('#wi-content');
      if (!content) return;
      const checkedIds = new Set([...content.querySelectorAll('.wi-row-cb:checked')].map(cb => cb.dataset.id));
      items = _items.filter(w => checkedIds.has(String(w.fields['System.Id'] || w.id)));
    }
    if (!items.length) { App.showToast('No items to export.', 'warning'); return; }

    const headers = ['ID','Title','Type','State','AssignedTo','Priority','Created','Updated','Org','Project'];
    const rows = items.map(w => {
      const f = w.fields;
      return [
        f['System.Id'] || w.id,
        f['System.Title'] || '',
        f['System.WorkItemType'] || '',
        f['System.State'] || '',
        String(f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || ''),
        f['Microsoft.VSTS.Common.Priority'] || '',
        f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toLocaleDateString() : '',
        f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toLocaleDateString() : '',
        w._connName || '',
        w._project || '',
      ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    });

    const csv = [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const date = new Date().toISOString().slice(0, 10);
    a.href = url; a.download = `workitems-${date}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ─── Create work item ──────────────────────────────────────────

  function _openCreateModal(container) {
    const conns = ConnectionsModule.getActive().filter(c => c.defaultProject || _context);
    if (!conns.length && !_context) {
      App.showToast('No active connection with a project. Load a project first.', 'warning');
      return;
    }

    const connSel = _context
      ? `<input type="hidden" id="wi-new-conn" value="${_esc(_context.connId)}" /><input type="hidden" id="wi-new-project" value="${_esc(_context.project)}" />`
      : `<div class="form-group"><label>Connection & Project</label>
           <select id="wi-new-conn-project" style="width:100%">
             ${conns.map(c => `<option value="${_esc(c.id)}|${_esc(c.defaultProject)}">${_esc(c.name)} — ${_esc(c.defaultProject)}</option>`).join('')}
           </select></div>`;

    App.openModal('New Work Item', `
      <form id="wi-create-form" style="display:flex;flex-direction:column;gap:12px">
        ${connSel}
        <div class="form-group">
          <label>Type</label>
          <select id="wi-new-type" style="width:100%">
            <option value="Task">Task</option>
            <option value="Bug">Bug</option>
            <option value="User Story">User Story</option>
            <option value="Epic">Epic</option>
            <option value="Feature">Feature</option>
            <option value="Issue">Issue</option>
          </select>
        </div>
        <div class="form-group">
          <label>Title <span style="color:var(--color-danger)">*</span></label>
          <input type="text" id="wi-new-title" placeholder="Enter work item title" required style="width:100%" />
        </div>
        <div class="form-group">
          <label>Description</label>
          <textarea id="wi-new-desc" rows="3" placeholder="Optional description" style="width:100%"></textarea>
        </div>
        <div id="wi-create-error" class="form-error hidden"></div>
        <div style="display:flex;gap:8px">
          <button type="submit" class="btn btn-primary" id="wi-create-submit-btn"><i class="fa-solid fa-plus"></i> Create</button>
          <button type="button" class="btn btn-secondary" id="wi-create-cancel-btn">Cancel</button>
        </div>
      </form>
    `);

    document.getElementById('wi-create-cancel-btn')?.addEventListener('click', () => App.closeModal());

    document.getElementById('wi-create-form')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const errEl = document.getElementById('wi-create-error');
      errEl.classList.add('hidden');

      let connId, project;
      if (_context) {
        connId = _context.connId; project = _context.project;
      } else {
        const cp = document.getElementById('wi-new-conn-project')?.value?.split('|') || [];
        connId = cp[0]; project = cp[1];
      }

      const conn = ConnectionsModule.getById(connId);
      if (!conn) { errEl.textContent = 'Invalid connection.'; errEl.classList.remove('hidden'); return; }

      const type  = document.getElementById('wi-new-type')?.value || 'Task';
      const title = document.getElementById('wi-new-title')?.value?.trim() || '';
      const desc  = document.getElementById('wi-new-desc')?.value?.trim() || '';
      if (!title) { errEl.textContent = 'Title is required.'; errEl.classList.remove('hidden'); return; }

      const btn = document.getElementById('wi-create-submit-btn');
      btn.disabled = true;
      btn.innerHTML = '<span class="spinner" style="width:14px;height:14px;border-width:2px;vertical-align:middle"></span> Creating…';

      const fields = { 'System.Title': title };
      if (desc) fields['System.Description'] = desc;
      const res = await AzureApi.createWorkItem(conn.orgUrl, project, type, fields, conn.pat);

      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-plus"></i> Create';

      if (res.error) { errEl.textContent = res.message; errEl.classList.remove('hidden'); return; }

      App.closeModal();
      App.showToast(`Work item #${res.id} created.`, 'success');
      // Add to list and re-render
      const newItem = { id: res.id, fields: res.fields, _connName: conn.name, _project: project, _connId: conn.id };
      _items.unshift(newItem);
      _renderTable(container);
    });
  }

  // ─── Detail modal ──────────────────────────────────────────────

  async function _openDetail(id, connIdOverride) {
    let conn;
    if (connIdOverride) {
      conn = ConnectionsModule.getById(connIdOverride);
    }
    if (!conn && _context) {
      conn = _context.conn;
    }
    if (!conn) {
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      if (item && item._connId) {
        conn = ConnectionsModule.getById(item._connId) || null;
      }
    }
    if (!conn) return;

    App.openModal('Work Item #' + id, '<div class="loading-state"><div class="spinner"></div><p>Loading…</p></div>');

    const [detailResult, relResult, commentsResult] = await Promise.all([
      AzureApi.getWorkItemDetail(conn.orgUrl, id, conn.pat),
      AzureApi.getWorkItemWithRelations(conn.orgUrl, id, conn.pat),
      AzureApi.getWorkItemComments(conn.orgUrl, id, conn.pat),
    ]);

    if (detailResult.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i><p>${_esc(detailResult.message)}</p></div>`);
      return;
    }

    const f = detailResult.fields;
    const relations = relResult.relations || [];

    // Group relations
    const relGroups = { Parent: [], Child: [], Related: [] };
    relations.forEach(r => {
      const rel = (r.attributes?.name || '').toLowerCase();
      const wiUrl = r.url || '';
      const wiId  = wiUrl.split('/').pop();
      if (rel === 'parent') relGroups.Parent.push(wiId);
      else if (rel === 'child') relGroups.Child.push(wiId);
      else relGroups.Related.push(wiId);
    });

    const relHtml = relations.length === 0 ? '' : `
      <div class="detail-row" style="flex-direction:column;gap:6px">
        <span class="detail-label">Linked Items</span>
        ${Object.entries(relGroups).map(([type, ids]) => ids.length === 0 ? '' : `
          <div style="font-size:.78rem">
            <strong>${type}:</strong> ${ids.map(wid => `<a href="#" style="color:var(--color-primary)">#${_esc(wid)}</a>`).join(', ')}
          </div>`).join('')}
      </div>`;

    // Comments
    const comments = commentsResult.error ? [] : (commentsResult.comments || commentsResult.value || []);
    const commentsHtml = `
      <div style="margin-top:12px;border-top:1px solid var(--border-color);padding-top:12px">
        <div style="font-weight:600;font-size:.85rem;margin-bottom:8px"><i class="fa-regular fa-comment"></i> Comments (${comments.length})</div>
        <div id="wi-detail-comments" style="display:flex;flex-direction:column;gap:8px;max-height:200px;overflow-y:auto">
          ${comments.length === 0 ? '<div class="text-muted text-sm">No comments yet.</div>' :
            comments.map(c => `
              <div style="background:var(--bg-table-alt);border-radius:var(--radius-sm);padding:8px;font-size:.78rem">
                <div style="font-weight:600;color:var(--text-secondary)">${_esc(c.createdBy?.displayName || 'Unknown')}
                  <span class="text-muted" style="font-weight:400"> · ${c.createdDate ? new Date(c.createdDate).toLocaleString() : ''}</span>
                </div>
                <div>${c.text || ''}</div>
              </div>`).join('')}
        </div>
        <form id="wi-comment-form" style="margin-top:10px;display:flex;gap:8px">
          <input type="text" id="wi-comment-text" placeholder="Add a comment…" style="flex:1;font-size:.82rem" />
          <button type="submit" class="btn btn-primary btn-sm"><i class="fa-solid fa-paper-plane"></i> Send</button>
        </form>
      </div>`;

    App.updateModalBody(`
      <div class="detail-row"><span class="detail-label">ID</span><span class="detail-value font-mono">#${detailResult.id}</span></div>
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
      ${relHtml}
      ${commentsHtml}
    `);

    // Bind comment form
    const commentForm = document.getElementById('wi-comment-form');
    if (commentForm) {
      commentForm.addEventListener('submit', async (ev) => {
        ev.preventDefault();
        const txt = document.getElementById('wi-comment-text')?.value?.trim();
        if (!txt) return;
        const res = await AzureApi.addWorkItemComment(conn.orgUrl, id, txt, conn.pat);
        if (res.error) { App.showToast('Failed to add comment: ' + res.message, 'error'); return; }
        App.showToast('Comment added.', 'success');
        document.getElementById('wi-comment-text').value = '';
        const feed = document.getElementById('wi-detail-comments');
        if (feed) {
          const div = document.createElement('div');
          div.style.cssText = 'background:var(--bg-table-alt);border-radius:var(--radius-sm);padding:8px;font-size:.78rem';
          div.innerHTML = `<div style="font-weight:600;color:var(--text-secondary)">You <span class="text-muted" style="font-weight:400"> · just now</span></div><div>${_esc(txt)}</div>`;
          feed.appendChild(div);
          feed.scrollTop = feed.scrollHeight;
          const noComments = feed.querySelector('.text-muted');
          if (noComments && noComments.textContent === 'No comments yet.') noComments.remove();
        }
      });
    }
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
