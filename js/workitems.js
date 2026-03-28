/**
 * workitems.js — Work Items
 *
 * Displays work items for a selected project with filtering,
 * pagination (50 per page), and a detail modal.
 * Supports "Single Project" and "All Connections" modes.
 *
 * Improvements:
 *  2.  Collapsible org groups (all mode)
 *  3.  Sortable columns
 *  4.  Resizable columns
 *  6.  Keyboard navigation
 *  7.  Breadcrumbs
 *  9.  Bulk actions
 * 10.  Export CSV
 * 11.  Inline edit
 * 12.  Comments in detail modal
 * 13.  Create work item
 * 14.  Linked work items in detail modal
 */

const WorkItemsModule = (() => {
  const PAGE_SIZE = 50;

  let _items       = [];
  let _currentPage = 1;
  let _context     = null;   // { connId, project, conn }
  let _mode        = 'single'; // 'single' | 'all'
  let _allSummary  = null;   // { total, connCount, failCount }
  let _sortCol     = '';
  let _sortDir     = 'asc';
  let _container   = null;

  // --- Public ---

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];
    _mode = 'single';
    _allSummary = null;
    _sortCol = '';
    _sortDir = 'asc';
    _container = container;

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${c.id}">${_esc(c.name)}</option>`
    ).join('');

    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-list-check"></i> Work Items</div>
      <p class="page-subtitle">Browse and track work items across your projects.</p>

      <div class="filter-bar mb-8" id="wi-mode-bar">
        <label style="font-weight:600">View:</label>
        <button class="btn btn-primary btn-sm active" id="wi-mode-single">Single Project</button>
        <button class="btn btn-secondary btn-sm" id="wi-mode-all">All Connections</button>
      </div>

      <div class="filter-bar mb-16" id="wi-single-bar">
        <select id="wi-conn-select" style="min-width:160px">
          <option value="">Select Organization...</option>
          ${connOptions}
        </select>
        <select id="wi-project-select" style="min-width:180px" disabled>
          <option value="">Select Project...</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-load-btn" disabled>
          <i class="fa-solid fa-download"></i> Load
        </button>
        <button class="btn btn-success btn-sm hidden" id="wi-new-btn">
          <i class="fa-solid fa-plus"></i> New Work Item
        </button>
      </div>

      <div class="filter-bar mb-16" id="wi-all-bar" style="display:none">
        <button class="btn btn-primary btn-sm" id="wi-load-all-btn">
          <i class="fa-solid fa-download"></i> Load All
        </button>
      </div>

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
          <option value="1">1 - Critical</option>
          <option value="2">2 - High</option>
          <option value="3">3 - Medium</option>
          <option value="4">4 - Low</option>
        </select>
        <input type="text" id="wi-search" placeholder="Search title..." style="min-width:160px" />
        <button class="btn btn-secondary btn-sm" id="wi-apply-filter-btn">
          <i class="fa-solid fa-filter"></i> Apply
        </button>
        <button class="btn btn-secondary btn-sm" id="wi-export-csv-btn" title="Export CSV">
          <i class="fa-solid fa-file-csv"></i> Export CSV
        </button>
      </div>

      <nav id="wi-breadcrumb" class="breadcrumb hidden"></nav>

      <div id="wi-bulk-bar" class="bulk-bar hidden">
        <span id="wi-bulk-count" style="font-weight:600"></span>
        <select id="wi-bulk-state" title="Change state">
          <option value="">Change State...</option>
          <option value="New">New</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
          <option value="Done">Done</option>
        </select>
        <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff;border:1px solid rgba(255,255,255,.4)" id="wi-bulk-apply-btn">
          <i class="fa-solid fa-check"></i> Apply
        </button>
        <button class="btn btn-sm" style="background:rgba(255,255,255,.2);color:#fff;border:1px solid rgba(255,255,255,.4)" id="wi-bulk-export-btn">
          <i class="fa-solid fa-file-csv"></i> Export Selected
        </button>
      </div>

      <div id="wi-content"></div>
    `;

    _bindSelectorEvents(container);

    if (params.connId) {
      const sel = container.querySelector('#wi-conn-select');
      sel.value = params.connId;
      await _loadProjects(container, params.connId);
      if (params.project) {
        const pSel = container.querySelector('#wi-project-select');
        pSel.value = params.project;
        if (pSel.value) {
          container.querySelector('#wi-load-btn').disabled = false;
          container.querySelector('#wi-new-btn').classList.remove('hidden');
          await _fetchItems(container);
        }
      }
    }
  }

  async function refresh(container) {
    if (_mode === 'all') {
      await _fetchAllItems(container);
    } else if (_context) {
      await _fetchItems(container);
    }
  }

  // --- Private ---

  function _bindSelectorEvents(container) {
    container.querySelector('#wi-mode-single').addEventListener('click', () => {
      if (_mode === 'single') return;
      _mode = 'single';
      container.querySelector('#wi-mode-single').className = 'btn btn-primary btn-sm active';
      container.querySelector('#wi-mode-all').className = 'btn btn-secondary btn-sm';
      container.querySelector('#wi-single-bar').style.display = '';
      container.querySelector('#wi-all-bar').style.display = 'none';
      container.querySelector('#wi-filters').style.display = 'none';
      container.querySelector('#wi-content').innerHTML = '';
      container.querySelector('#wi-breadcrumb').classList.add('hidden');
      container.querySelector('#wi-bulk-bar').classList.add('hidden');
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
      container.querySelector('#wi-breadcrumb').classList.add('hidden');
      container.querySelector('#wi-bulk-bar').classList.add('hidden');
      _items = [];
      _allSummary = null;
    });

    container.addEventListener('change', async e => {
      if (e.target.id === 'wi-conn-select') {
        const connId = e.target.value;
        container.querySelector('#wi-project-select').innerHTML = '<option value="">Loading...</option>';
        container.querySelector('#wi-project-select').disabled = true;
        container.querySelector('#wi-load-btn').disabled = true;
        container.querySelector('#wi-new-btn').classList.add('hidden');
        if (connId) await _loadProjects(container, connId);
      }
      if (e.target.id === 'wi-project-select') {
        const hasProject = !!e.target.value;
        container.querySelector('#wi-load-btn').disabled = !hasProject;
        container.querySelector('#wi-new-btn').classList.toggle('hidden', !hasProject);
      }
      if (e.target.id === 'wi-select-all') {
        container.querySelectorAll('.wi-row-check').forEach(cb => { cb.checked = e.target.checked; });
        _updateBulkBar(container);
      }
      if (e.target.classList.contains('wi-row-check')) {
        _updateBulkBar(container);
        const all = container.querySelectorAll('.wi-row-check');
        const selectAll = container.querySelector('#wi-select-all');
        if (selectAll) selectAll.checked = Array.from(all).every(cb => cb.checked);
      }
    });

    container.addEventListener('click', async e => {
      if (e.target.closest('#wi-load-btn'))         await _fetchItems(container);
      if (e.target.closest('#wi-load-all-btn'))     await _fetchAllItems(container);
      if (e.target.closest('#wi-apply-filter-btn')) { _currentPage = 1; _renderTable(container); }
      if (e.target.closest('#wi-export-csv-btn'))   _exportCSV(_getFilteredItems(container));
      if (e.target.closest('#wi-new-btn'))          _openCreateModal(container);
      if (e.target.closest('#wi-bulk-apply-btn'))   await _bulkChangeState(container);
      if (e.target.closest('#wi-bulk-export-btn'))  _exportCSV(_getSelectedItems(container));
    });

    container.addEventListener('keydown', e => {
      if (e.key === 'Enter' && e.target.id === 'wi-search') { _currentPage = 1; _renderTable(container); }
    });
  }

  async function _loadProjects(container, connId) {
    const conn = ConnectionsModule.getById(connId);
    if (!conn) return;
    const pSel = container.querySelector('#wi-project-select');
    const result = await AzureApi.getProjects(conn.orgUrl, conn.pat);
    if (result.error) { App.showToast(result.message, 'error'); return; }
    const projects = result.value || [];
    pSel.innerHTML = '<option value="">Select Project...</option>' +
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
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading work items...</p></div>';
    container.querySelector('#wi-filters').style.display = 'none';

    const filters = _getFilters(container);
    const result  = await AzureApi.getWorkItems(conn.orgUrl, project, conn.pat, filters);

    if (result.error) {
      content.innerHTML = `<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i>
        <p>${_esc(result.message)}</p>
        ${result.cors ? '<p class="text-sm">CORS restrictions may apply.</p>' : ''}</div>`;
      return;
    }

    _items = result.value || [];
    container.querySelector('#wi-filters').style.display = '';
    _updateBreadcrumb(container);
    _renderTable(container);
  }

  async function _fetchAllItems(container) {
    const activeConns = ConnectionsModule.getActive().filter(c => c.defaultProject);
    if (activeConns.length === 0) {
      container.querySelector('#wi-content').innerHTML = `<div class="empty-state">
        <i class="fa-solid fa-plug-circle-xmark"></i>
        <p>No connections have a default project configured.
          <a href="#/connections" style="color:var(--color-primary)">Go to Connections</a> to set one.</p>
        </div>`;
      return;
    }

    _currentPage = 1;
    _context = null;
    const content = container.querySelector('#wi-content');
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading work items from all connections...</p></div>';
    container.querySelector('#wi-filters').style.display = 'none';
    container.querySelector('#wi-breadcrumb').classList.add('hidden');

    const results = await Promise.allSettled(
      activeConns.map(conn =>
        AzureApi.getWorkItems(conn.orgUrl, conn.defaultProject, conn.pat, {}).then(res => ({ conn, res }))
      )
    );

    let allItems = [];
    let failCount = 0;
    for (const outcome of results) {
      if (outcome.status === 'rejected' || outcome.value.res.error) { failCount++; continue; }
      const { conn, res } = outcome.value;
      const items = (res.value || []).map(item => ({ ...item, _connName: conn.name, _project: conn.defaultProject, _connId: conn.id }));
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

  function _getFilteredItems(container) {
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

  function _getSelectedItems(container) {
    const checkedIds = new Set(Array.from(container.querySelectorAll('.wi-row-check:checked')).map(cb => cb.dataset.id));
    return _items.filter(w => checkedIds.has(String(w.fields['System.Id'] || w.id)));
  }

  function _updateBreadcrumb(container) {
    const bc = container.querySelector('#wi-breadcrumb');
    if (!bc) return;
    if (_mode === 'all' || !_context) { bc.classList.add('hidden'); return; }
    const { conn, project } = _context;
    bc.innerHTML = `
      <a href="#/connections">All Connections</a>
      <span class="breadcrumb-sep">&rsaquo;</span>
      <a href="#/projects">${_esc(conn.name)}</a>
      <span class="breadcrumb-sep">&rsaquo;</span>
      <span class="breadcrumb-current">${_esc(project)}</span>`;
    bc.classList.remove('hidden');
  }

  function _updateBulkBar(container) {
    const checked = container.querySelectorAll('.wi-row-check:checked').length;
    const bar = container.querySelector('#wi-bulk-bar');
    if (!bar) return;
    if (checked > 0) {
      bar.classList.remove('hidden');
      const el = container.querySelector('#wi-bulk-count');
      if (el) el.textContent = `${checked} selected`;
    } else {
      bar.classList.add('hidden');
    }
  }

  function _sortedItems(items) {
    if (!_sortCol) return items;
    const dir = _sortDir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let va = (a.fields && a.fields[_sortCol] != null) ? a.fields[_sortCol] : (a[_sortCol] || '');
      let vb = (b.fields && b.fields[_sortCol] != null) ? b.fields[_sortCol] : (b[_sortCol] || '');
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return -1 * dir;
      if (va > vb) return  1 * dir;
      return 0;
    });
  }

  function _sortColClass(col) {
    if (_sortCol !== col) return 'sortable';
    return _sortDir === 'asc' ? 'sortable sorted-asc' : 'sortable sorted-desc';
  }

  function _renderTable(container) {
    const content    = container.querySelector('#wi-content');
    const showOrgCol = _mode === 'all';
    let filtered = _sortedItems(_getFilteredItems(container));
    const total  = filtered.length;
    const pages  = Math.ceil(total / PAGE_SIZE) || 1;
    _currentPage = Math.min(_currentPage, pages);
    const start  = (_currentPage - 1) * PAGE_SIZE;
    const pageItems = filtered.slice(start, start + PAGE_SIZE);

    if (filtered.length === 0) {
      content.innerHTML = `<div class="empty-state"><i class="fa-solid fa-list-check"></i><p>No work items found.</p></div>`;
      return;
    }

    const summaryBanner = (showOrgCol && _allSummary) ? `
      <div class="filter-bar mb-8" style="flex-wrap:wrap;gap:8px">
        <span class="badge badge-success"><i class="fa-solid fa-circle-check"></i>
          ${_allSummary.total} item${_allSummary.total !== 1 ? 's' : ''} from
          ${_allSummary.connCount} connection${_allSummary.connCount !== 1 ? 's' : ''}</span>
        ${_allSummary.failCount ? `<span class="badge badge-warning">
          <i class="fa-solid fa-triangle-exclamation"></i>
          ${_allSummary.failCount} connection${_allSummary.failCount !== 1 ? 's' : ''} failed</span>` : ''}
      </div>` : '';

    const orgCols = showOrgCol
      ? `<th class="${_sortColClass('_connName')}" data-col="_connName">Organization</th>
         <th class="${_sortColClass('_project')}" data-col="_project">Project</th>`
      : '';

    // Build rows - grouped by org in all mode
    let rowsHtml = '';
    if (showOrgCol) {
      const groups = {};
      const groupOrder = [];
      pageItems.forEach(w => {
        const org = w._connName || 'Unknown';
        if (!groups[org]) { groups[org] = []; groupOrder.push(org); }
        groups[org].push(w);
      });
      const totalCols = 10 + 2; // checkbox + id + 2 org + title + type + state + assign + pri + created + updated + edit
      groupOrder.forEach(org => {
        rowsHtml += `<tr class="group-header-row" data-group="${_esc(org)}">
          <td colspan="${totalCols}"><span class="group-arrow">&#9660;</span>
            <i class="fa-brands fa-windows" style="margin-right:6px;color:var(--color-primary)"></i>
            ${_esc(org)} <span class="text-muted" style="font-weight:400">(${groups[org].length} items)</span>
          </td></tr>`;
        groups[org].forEach(w => { rowsHtml += _rowHtml(w, true, org); });
      });
    } else {
      rowsHtml = pageItems.map(w => _rowHtml(w, false, '')).join('');
    }

    content.innerHTML = `
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table tabindex="0" id="wi-table">
          <thead><tr>
            <th style="width:32px"><input type="checkbox" id="wi-select-all" title="Select all" /></th>
            <th class="${_sortColClass('System.Id')}" data-col="System.Id">ID</th>
            ${orgCols}
            <th class="${_sortColClass('System.Title')}" data-col="System.Title">Title</th>
            <th class="${_sortColClass('System.WorkItemType')}" data-col="System.WorkItemType">Type</th>
            <th class="${_sortColClass('System.State')}" data-col="System.State">State</th>
            <th class="${_sortColClass('System.AssignedTo')}" data-col="System.AssignedTo">Assigned To</th>
            <th class="${_sortColClass('Microsoft.VSTS.Common.Priority')}" data-col="Microsoft.VSTS.Common.Priority">Priority</th>
            <th class="${_sortColClass('System.CreatedDate')}" data-col="System.CreatedDate">Created</th>
            <th class="${_sortColClass('System.ChangedDate')}" data-col="System.ChangedDate">Updated</th>
            <th style="width:40px"></th>
          </tr></thead>
          <tbody id="wi-tbody">${rowsHtml}</tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}`;

    // Row clicks (open detail, not on checkbox or edit btn)
    content.querySelectorAll('.wi-row').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.wi-row-check') || e.target.closest('.wi-edit-btn')) return;
        _openDetail(row.dataset.id, row.dataset.connId);
      });
    });

    // Edit button
    content.querySelectorAll('.wi-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        const row  = btn.closest('tr');
        const id   = row.dataset.id;
        const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
        if (!item) return;
        const conn = _getConnForItem(item, row.dataset.connId);
        if (!conn) return;
        _startInlineEdit(row, item, conn, container);
      });
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });

    // Collapsible org group headers
    content.querySelectorAll('.group-header-row').forEach(headerRow => {
      headerRow.addEventListener('click', () => {
        const group = headerRow.dataset.group;
        const isCollapsed = headerRow.classList.toggle('group-collapsed');
        content.querySelectorAll(`.wi-row[data-group]`).forEach(row => {
          if (row.dataset.group === group) row.classList.toggle('hidden', isCollapsed);
        });
      });
    });

    // Sort headers
    const table = content.querySelector('#wi-table');
    if (table) {
      table.querySelectorAll('th[data-col]').forEach(th => {
        th.addEventListener('click', () => {
          const col = th.dataset.col;
          if (_sortCol === col) {
            _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
          } else {
            _sortCol = col;
            _sortDir = 'asc';
          }
          _renderTable(container);
        });
      });
      _initResizableColumns(table);
      _initKeyboardNav(table, container);
    }
  }

  // --- Inline Edit ---

  function _startInlineEdit(row, item, conn, container) {
    const stateCell  = row.querySelector('.wi-state-cell');
    const assignCell = row.querySelector('.wi-assign-cell');
    if (!stateCell || !assignCell) return;

    const currentState  = stateCell.dataset.state  || '';
    const currentAssign = assignCell.dataset.assigned || '';

    stateCell.innerHTML = `<select class="wi-inline-state" style="font-size:.78rem;padding:2px 6px;max-width:100px">
      ${['New','Active','Resolved','Closed','Done'].map(s =>
        `<option${s === currentState ? ' selected' : ''}>${s}</option>`
      ).join('')}
    </select>`;
    assignCell.innerHTML = `<input type="text" class="wi-inline-assign" value="${_esc(currentAssign)}"
      style="font-size:.78rem;padding:2px 6px;width:100%" />`;

    const editBtn = row.querySelector('.wi-edit-btn');
    editBtn.innerHTML = `
      <button class="btn btn-primary btn-sm wi-save-btn" style="padding:2px 8px;font-size:.72rem">Save</button>
      <button class="btn btn-secondary btn-sm wi-cancel-btn" style="padding:2px 8px;font-size:.72rem">Cancel</button>`;

    editBtn.querySelector('.wi-cancel-btn').addEventListener('click', e => {
      e.stopPropagation();
      _renderTable(container);
    });

    editBtn.querySelector('.wi-save-btn').addEventListener('click', async e => {
      e.stopPropagation();
      const newState  = stateCell.querySelector('.wi-inline-state').value;
      const newAssign = assignCell.querySelector('.wi-inline-assign').value.trim();
      const id = item.fields['System.Id'] || item.id;
      const fields = { 'System.State': newState };
      if (newAssign !== currentAssign) fields['System.AssignedTo'] = newAssign;
      const res = await AzureApi.updateWorkItemFields(conn.orgUrl, id, conn.pat, fields);
      if (res.error) {
        App.showToast(res.message, 'error');
      } else {
        App.showToast('Work item updated.', 'success');
        const wi = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
        if (wi) {
          wi.fields['System.State'] = newState;
          if (newAssign !== currentAssign) wi.fields['System.AssignedTo'] = newAssign;
        }
        _renderTable(container);
      }
    });
  }

  // --- Bulk state change ---

  async function _bulkChangeState(container) {
    const newState = container.querySelector('#wi-bulk-state')?.value;
    if (!newState) { App.showToast('Please select a state.', 'warning'); return; }
    const selected = _getSelectedItems(container);
    if (selected.length === 0) return;
    const applyBtn = container.querySelector('#wi-bulk-apply-btn');
    applyBtn.disabled = true;
    let successCount = 0;
    for (const item of selected) {
      const conn = _getConnForItem(item, item._connId);
      if (!conn) continue;
      const id = item.fields['System.Id'] || item.id;
      const res = await AzureApi.updateWorkItemState(conn.orgUrl, id, conn.pat, newState);
      if (!res.error) { item.fields['System.State'] = newState; successCount++; }
    }
    applyBtn.disabled = false;
    App.showToast(`Updated ${successCount} of ${selected.length} work items.`,
      successCount === selected.length ? 'success' : 'warning');
    _renderTable(container);
  }

  // --- Export CSV ---

  function _exportCSV(items) {
    const showOrg = _mode === 'all';
    const headers = ['ID', ...(showOrg ? ['Organization','Project'] : []),
      'Title','Type','State','Assigned To','Priority','Created','Updated'];
    const rows = items.map(w => {
      const f = w.fields;
      return [
        f['System.Id'] || w.id,
        ...(showOrg ? [w._connName || '', w._project || ''] : []),
        f['System.Title'] || '',
        f['System.WorkItemType'] || '',
        f['System.State'] || '',
        (f['System.AssignedTo']?.displayName) || f['System.AssignedTo'] || '',
        f['Microsoft.VSTS.Common.Priority'] || '',
        f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toLocaleDateString() : '',
        f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toLocaleDateString() : '',
      ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
    });
    const csv  = [headers.join(','), ...rows].join('\n');
    const date = new Date().toISOString().slice(0,10);
    const blob = new Blob([csv], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = `workitems-${date}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // --- Create Work Item ---

  function _openCreateModal(container) {
    if (!_context) return;
    App.openModal('New Work Item', `
      <form id="wi-create-form">
        <div class="form-group">
          <label>Title <span style="color:var(--color-danger)">*</span></label>
          <input type="text" id="wi-create-title" placeholder="Enter title..." required />
        </div>
        <div class="form-group">
          <label>Type</label>
          <select id="wi-create-type">
            <option value="Task">Task</option>
            <option value="Bug">Bug</option>
            <option value="User Story">User Story</option>
            <option value="Epic">Epic</option>
            <option value="Feature">Feature</option>
            <option value="Issue">Issue</option>
          </select>
        </div>
        <div class="form-group">
          <label>Assigned To</label>
          <input type="text" id="wi-create-assign" placeholder="user@domain.com" />
        </div>
        <div class="form-group">
          <label>Priority</label>
          <select id="wi-create-priority">
            <option value="1">1 - Critical</option>
            <option value="2">2 - High</option>
            <option value="3" selected>3 - Medium</option>
            <option value="4">4 - Low</option>
          </select>
        </div>
        <div class="form-group">
          <label>Description</label>
          <textarea id="wi-create-desc" rows="4" placeholder="Description..."></textarea>
        </div>
        <div id="wi-create-error" class="form-error mb-8 hidden"></div>
        <div style="display:flex;gap:8px">
          <button type="submit" class="btn btn-primary btn-sm" id="wi-create-submit">
            <i class="fa-solid fa-plus"></i> Create
          </button>
          <button type="button" class="btn btn-secondary btn-sm" id="wi-create-cancel">Cancel</button>
        </div>
      </form>`);

    const mb = document.getElementById('modal-body');
    mb.querySelector('#wi-create-cancel').addEventListener('click', () => App.closeModal());
    mb.querySelector('#wi-create-form').addEventListener('submit', async e => {
      e.preventDefault();
      const title    = mb.querySelector('#wi-create-title').value.trim();
      const type     = mb.querySelector('#wi-create-type').value;
      const assigned = mb.querySelector('#wi-create-assign').value.trim();
      const priority = mb.querySelector('#wi-create-priority').value;
      const desc     = mb.querySelector('#wi-create-desc').value.trim();
      const errEl    = mb.querySelector('#wi-create-error');
      if (!title) { errEl.textContent = 'Title is required.'; errEl.classList.remove('hidden'); return; }
      errEl.classList.add('hidden');
      const submitBtn = mb.querySelector('#wi-create-submit');
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="spinner" style="width:12px;height:12px;border-width:2px;vertical-align:middle"></span> Creating...';
      const fields = { 'System.Title': title, 'Microsoft.VSTS.Common.Priority': Number(priority) };
      if (assigned) fields['System.AssignedTo'] = assigned;
      if (desc)     fields['System.Description'] = desc;
      const { conn, project } = _context;
      const res = await AzureApi.createWorkItem(conn.orgUrl, project, type, conn.pat, fields);
      submitBtn.disabled = false;
      submitBtn.innerHTML = '<i class="fa-solid fa-plus"></i> Create';
      if (res.error) {
        errEl.textContent = res.message;
        errEl.classList.remove('hidden');
      } else {
        App.closeModal();
        App.showToast(`Work item #${res.id} created.`, 'success');
        await _fetchItems(container);
      }
    });
  }

  // --- Detail modal (comments + linked items) ---

  async function _openDetail(id, connIdOverride) {
    const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
    const conn = _getConnForItem(item || {}, connIdOverride);
    if (!conn) return;

    App.openModal('Work Item #' + id, '<div class="loading-state"><div class="spinner"></div><p>Loading...</p></div>');

    const [wiResult, commentsResult] = await Promise.all([
      AzureApi.getWorkItemWithRelations(conn.orgUrl, id, conn.pat),
      AzureApi.getWorkItemComments(conn.orgUrl, id, conn.pat),
    ]);

    if (wiResult.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i>
        <p>${_esc(wiResult.message)}</p></div>`);
      return;
    }

    const f = wiResult.fields;

    // Linked items
    const relations = wiResult.relations || [];
    const parents   = relations.filter(r => r.rel === 'System.LinkTypes.Hierarchy-Reverse');
    const children  = relations.filter(r => r.rel === 'System.LinkTypes.Hierarchy-Forward');
    const related   = relations.filter(r => r.rel === 'System.LinkTypes.Related');
    const relId = r => r.url.split('/').pop();
    const allLinkedIds = [...parents, ...children, ...related].map(relId).filter(Boolean);

    let linkedItems = {};
    if (allLinkedIds.length > 0) {
      const linked = await AzureApi.getWorkItemsByIds(conn.orgUrl, allLinkedIds, conn.pat);
      if (!linked.error) (linked.value || []).forEach(wi => { linkedItems[wi.id] = wi; });
    }

    const linkItemHtml = (r, label) => {
      const wiId = relId(r);
      const wi   = linkedItems[wiId];
      const title = wi ? _esc(wi.fields['System.Title'] || '(no title)') : `#${wiId}`;
      return `<div class="detail-link-row" data-linked-id="${_esc(wiId)}"
        style="cursor:pointer;display:flex;align-items:center;gap:8px;padding:3px 0">
        <span class="badge badge-secondary">${label}</span>
        <a style="color:var(--color-primary)">#${_esc(String(wiId))} ${title}</a>
      </div>`;
    };

    const linksHtml = (parents.length || children.length || related.length) ? `
      <div class="detail-row" style="flex-direction:column;gap:4px">
        <span class="detail-label">Links</span>
        ${parents.map(r  => linkItemHtml(r, 'Parent')).join('')}
        ${children.map(r => linkItemHtml(r, 'Child')).join('')}
        ${related.map(r  => linkItemHtml(r, 'Related')).join('')}
      </div>` : '';

    // Comments
    const comments = commentsResult.error ? [] : (commentsResult.comments || []);
    const commentsHtml = `
      <div class="detail-row" style="flex-direction:column;gap:6px">
        <span class="detail-label">Comments (${comments.length})</span>
        ${comments.length > 0 ? `<div class="wi-comments-list">
          ${comments.map(c => `<div class="wi-comment">
            <div class="wi-comment-author">${_esc(c.createdBy?.displayName || '?')}</div>
            <div class="wi-comment-text">${c.text || ''}</div>
            <div class="text-muted text-sm">${c.createdDate ? new Date(c.createdDate).toLocaleString() : ''}</div>
          </div>`).join('')}
        </div>` : '<div class="text-muted text-sm">No comments yet.</div>'}
        <div style="margin-top:8px">
          <textarea id="wi-comment-text" rows="3" placeholder="Add a comment..."
            style="width:100%;font-size:.82rem;padding:6px 8px;border:1px solid var(--border-color);
                   border-radius:var(--radius-sm);background:var(--bg-input);color:var(--text-primary)"></textarea>
          <button class="btn btn-primary btn-sm mt-4" id="wi-comment-submit">
            <i class="fa-solid fa-paper-plane"></i> Add Comment
          </button>
        </div>
      </div>`;

    App.updateModalBody(`
      <div class="detail-row"><span class="detail-label">ID</span><span class="detail-value font-mono">#${wiResult.id}</span></div>
      <div class="detail-row"><span class="detail-label">Title</span><span class="detail-value">${_esc(f['System.Title'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">Type</span><span class="detail-value">${_typeBadge(f['System.WorkItemType'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">State</span><span class="detail-value">${_stateBadge(f['System.State'] || '')}</span></div>
      <div class="detail-row"><span class="detail-label">Assigned To</span>
        <span class="detail-value">${_esc(String(f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || '?'))}</span></div>
      <div class="detail-row"><span class="detail-label">Priority</span><span class="detail-value">${_priorityBadge(f['Microsoft.VSTS.Common.Priority'] || '?')}</span></div>
      <div class="detail-row"><span class="detail-label">Area Path</span><span class="detail-value">${_esc(f['System.AreaPath'] || '?')}</span></div>
      <div class="detail-row"><span class="detail-label">Iteration</span><span class="detail-value">${_esc(f['System.IterationPath'] || '?')}</span></div>
      <div class="detail-row"><span class="detail-label">Created</span>
        <span class="detail-value">${_esc(f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toLocaleString() : '?')}</span></div>
      <div class="detail-row"><span class="detail-label">Updated</span>
        <span class="detail-value">${_esc(f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toLocaleString() : '?')}</span></div>
      ${f['System.Description'] ? `
      <div class="detail-row" style="flex-direction:column;gap:6px">
        <span class="detail-label">Description</span>
        <div style="font-size:.82rem;border:1px solid var(--border-color);padding:10px;
                    border-radius:var(--radius-sm);background:var(--bg-table-alt);">
          ${f['System.Description']}
        </div></div>` : ''}
      ${linksHtml}
      ${commentsHtml}`);

    const mb = document.getElementById('modal-body');
    mb.querySelectorAll('.detail-link-row').forEach(row => {
      row.addEventListener('click', () => _openDetail(row.dataset.linkedId, connIdOverride));
    });
    mb.querySelector('#wi-comment-submit')?.addEventListener('click', async () => {
      const text = mb.querySelector('#wi-comment-text')?.value?.trim();
      if (!text) return;
      const btn = mb.querySelector('#wi-comment-submit');
      btn.disabled = true;
      btn.innerHTML = '<span class="spinner" style="width:12px;height:12px;border-width:2px;vertical-align:middle"></span> Posting...';
      const res = await AzureApi.addWorkItemComment(conn.orgUrl, id, conn.pat, text);
      if (res.error) {
        App.showToast(res.message, 'error');
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Add Comment';
      } else {
        App.showToast('Comment added.', 'success');
        _openDetail(id, connIdOverride);
      }
    });
  }

  function _getConnForItem(item, connIdOverride) {
    if (connIdOverride) { const c = ConnectionsModule.getById(connIdOverride); if (c) return c; }
    if (_context && _context.conn) return _context.conn;
    if (item && item._connId) return ConnectionsModule.getById(item._connId) || null;
    return null;
  }

  // --- Resizable columns ---

  function _initResizableColumns(table) {
    table.querySelectorAll('th').forEach(th => {
      if (th.querySelector('.col-resize-handle')) return; // avoid duplicates
      const handle = document.createElement('span');
      handle.className = 'col-resize-handle';
      th.appendChild(handle);
      handle.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();
        const startX = e.pageX;
        const startW = th.offsetWidth;
        const onMove = ev => {
          const newW = Math.max(40, startW + (ev.pageX - startX));
          th.style.width = newW + 'px';
          th.style.minWidth = newW + 'px';
        };
        const onUp = () => {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });
  }

  // --- Keyboard navigation ---

  function _initKeyboardNav(table, container) {
    table.addEventListener('keydown', e => {
      if (!['ArrowUp','ArrowDown','Enter','Escape'].includes(e.key)) return;
      const rows = Array.from(table.querySelectorAll('.wi-row:not(.hidden)'));
      if (rows.length === 0) return;
      const focused = table.querySelector('.row-focused');
      let idx = focused ? rows.indexOf(focused) : -1;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        idx = Math.min(idx + 1, rows.length - 1);
        rows.forEach((r, i) => r.classList.toggle('row-focused', i === idx));
        rows[idx]?.scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        idx = Math.max(idx - 1, 0);
        rows.forEach((r, i) => r.classList.toggle('row-focused', i === idx));
        rows[idx]?.scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter' && focused) {
        _openDetail(focused.dataset.id, focused.dataset.connId);
      } else if (e.key === 'Escape') {
        App.closeModal();
      }
    });
  }

  // --- Pagination ---

  function _paginationHtml(current, total) {
    if (total <= 1) return '';
    let html = '<div class="pagination">';
    html += `<button class="page-btn" data-page="${current - 1}" ${current === 1 ? 'disabled' : ''}>
      <i class="fa-solid fa-chevron-left"></i></button>`;
    for (let p = 1; p <= total; p++) {
      if (total > 7 && p > 2 && p < total - 1 && Math.abs(p - current) > 1) {
        if (p === 3 || p === total - 2) html += '<span class="page-info">...</span>';
        continue;
      }
      html += `<button class="page-btn ${p === current ? 'active' : ''}" data-page="${p}">${p}</button>`;
    }
    html += `<button class="page-btn" data-page="${current + 1}" ${current === total ? 'disabled' : ''}>
      <i class="fa-solid fa-chevron-right"></i></button></div>`;
    return html;
  }

  // --- Row HTML ---

  function _rowHtml(w, showOrgCol, groupKey) {
    const f        = w.fields;
    const id       = f['System.Id']   || w.id;
    const title    = f['System.Title']|| '(no title)';
    const type     = f['System.WorkItemType'] || '';
    const state    = f['System.State'] || '';
    const assigned = (f['System.AssignedTo']?.displayName) || f['System.AssignedTo'] || '?';
    const priority = f['Microsoft.VSTS.Common.Priority'] || '?';
    const created  = f['System.CreatedDate'] ? _fmtDate(f['System.CreatedDate']) : '?';
    const updated  = f['System.ChangedDate'] ? _fmtDate(f['System.ChangedDate']) : '?';
    const orgCells = showOrgCol
      ? `<td class="text-sm">${_esc(w._connName || '')}</td><td class="text-sm">${_esc(w._project || '')}</td>`
      : '';
    const connIdAttr = w._connId ? ` data-conn-id="${_esc(w._connId)}"` : '';
    const groupAttr  = groupKey  ? ` data-group="${_esc(groupKey)}"` : '';
    return `<tr class="wi-row" data-id="${id}"${connIdAttr}${groupAttr} style="cursor:pointer">
      <td onclick="event.stopPropagation()"><input type="checkbox" class="wi-row-check" data-id="${id}" /></td>
      <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
      ${orgCells}
      <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
      <td>${_typeBadge(type)}</td>
      <td class="wi-state-cell" data-state="${_esc(state)}">${_stateBadge(state)}</td>
      <td class="truncate wi-assign-cell" style="max-width:140px" data-assigned="${_esc(String(assigned))}">${_esc(String(assigned))}</td>
      <td>${_priorityBadge(priority)}</td>
      <td class="text-muted text-sm">${created}</td>
      <td class="text-muted text-sm">${updated}</td>
      <td><button class="wi-edit-btn" title="Edit"><i class="fa-solid fa-pen-to-square"></i></button></td>
    </tr>`;
  }

  // --- Badge helpers ---

  function _typeBadge(type) {
    const icons = {
      'Bug':        { icon: 'fa-bug',                cls: 'badge-danger' },
      'Task':       { icon: 'fa-thumbtack',          cls: 'badge-primary' },
      'User Story': { icon: 'fa-book',               cls: 'badge-success' },
      'Epic':       { icon: 'fa-bolt',               cls: 'badge-warning' },
      'Feature':    { icon: 'fa-star',               cls: 'badge-info' },
      'Issue':      { icon: 'fa-circle-exclamation', cls: 'badge-danger' },
      'Test Case':  { icon: 'fa-flask',              cls: 'badge-secondary' },
    };
    const d = icons[type] || { icon: 'fa-file', cls: 'badge-secondary' };
    return `<span class="badge ${d.cls}"><i class="fa-solid ${d.icon}"></i> ${_esc(type)}</span>`;
  }

  function _stateBadge(state) {
    const map = { 'Active':'badge-primary','New':'badge-info','Resolved':'badge-warning','Closed':'badge-secondary','Done':'badge-success' };
    return `<span class="badge ${map[state] || 'badge-secondary'}">${_esc(state)}</span>`;
  }

  function _priorityBadge(p) {
    const map = { 1:'badge-danger', 2:'badge-warning', 3:'badge-primary', 4:'badge-secondary' };
    return `<span class="badge ${map[p] || 'badge-secondary'}">${p}</span>`;
  }

  function _fmtDate(s) {
    return new Date(s).toLocaleDateString(undefined, { month:'short', day:'numeric', year:'numeric' });
  }

  function _esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  return { render, refresh };
})();
