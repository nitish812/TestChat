/**
 * workitems.js — Work Items
 *
 * Features:
 *  - Single Project / All Connections modes
 *  - Collapsible org groups (Feature 2)
 *  - Sortable columns (Feature 3)
 *  - Resizable columns (Feature 4)
 *  - Keyboard navigation (Feature 6)
 *  - Breadcrumb nav (Feature 7)
 *  - Bulk actions (Feature 9)
 *  - Export CSV (Feature 10)
 *  - Inline edit (Feature 11)
 *  - Comments in modal (Feature 12)
 *  - Create work item (Feature 13)
 *  - Linked work items in modal (Feature 14)
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

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];
    _mode = 'single';
    _allSummary = null;
    _sortCol = '';
    _sortDir = 'asc';

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${c.id}">${_esc(c.name)}</option>`
    ).join('');

    container.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:4px">
        <div class="page-title" style="margin-bottom:0"><i class="fa-solid fa-list-check"></i> Work Items</div>
        <button class="btn btn-secondary btn-sm hidden" id="wi-new-btn">
          <i class="fa-solid fa-plus"></i> New Work Item
        </button>
      </div>
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
        <button class="btn btn-secondary btn-sm" id="wi-export-csv-btn">
          <i class="fa-solid fa-file-csv"></i> Export CSV
        </button>
      </div>

      <!-- Bulk action bar (hidden until checkboxes are checked) -->
      <div class="bulk-bar hidden" id="wi-bulk-bar">
        <span id="wi-bulk-count">0 selected</span>
        <select id="wi-bulk-state">
          <option value="">Change State…</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-bulk-apply-btn"><i class="fa-solid fa-check"></i> Apply</button>
        <button class="btn btn-secondary btn-sm" id="wi-bulk-export-btn"><i class="fa-solid fa-file-csv"></i> Export Selected</button>
      </div>

      <div id="wi-breadcrumb"></div>
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
      container.querySelector('#wi-content').innerHTML = '';
      container.querySelector('#wi-breadcrumb').innerHTML = '';
      container.querySelector('#wi-new-btn').classList.add('hidden');
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
      container.querySelector('#wi-breadcrumb').innerHTML = '';
      container.querySelector('#wi-new-btn').classList.add('hidden');
      _items = [];
      _allSummary = null;
    });

    container.addEventListener('change', async e => {
      if (e.target.id === 'wi-conn-select') {
        const connId = e.target.value;
        container.querySelector('#wi-project-select').innerHTML = '<option value="">Loading…</option>';
        container.querySelector('#wi-project-select').disabled = true;
        container.querySelector('#wi-load-btn').disabled = true;
        container.querySelector('#wi-new-btn').classList.add('hidden');
        if (connId) await _loadProjects(container, connId);
      }
      if (e.target.id === 'wi-project-select') {
        container.querySelector('#wi-load-btn').disabled = !e.target.value;
      }
    });

    container.addEventListener('click', async e => {
      if (e.target.closest('#wi-load-btn'))           await _fetchItems(container);
      if (e.target.closest('#wi-load-all-btn'))       await _fetchAllItems(container);
      if (e.target.closest('#wi-apply-filter-btn'))   _renderTable(container);
      if (e.target.closest('#wi-export-csv-btn'))     _exportCsv(container);
      if (e.target.closest('#wi-new-btn'))             _openCreateModal(container);
      if (e.target.closest('#wi-bulk-apply-btn'))     await _bulkApply(container);
      if (e.target.closest('#wi-bulk-export-btn'))    _exportCsv(container, true);
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
    container.querySelector('#wi-breadcrumb').innerHTML = '';

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
    // Show "+ New Work Item" button in single project mode
    container.querySelector('#wi-new-btn').classList.remove('hidden');
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
    container.querySelector('#wi-breadcrumb').innerHTML = '';

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
    _renderBreadcrumb(container, null, null);
    _renderTable(container);
  }

  function _getFilters(container) {
    return {
      type:     container.querySelector('#wi-type-filter')?.value || '',
      state:    container.querySelector('#wi-state-filter')?.value || '',
      priority: container.querySelector('#wi-priority-filter')?.value || '',
    };
  }

  // ─── Feature 7: Breadcrumb ─────────────────────────────────────

  function _renderBreadcrumb(container, orgName, projectName) {
    const el = container.querySelector('#wi-breadcrumb');
    if (!el) return;
    if (_mode === 'all') {
      el.innerHTML = `<nav class="breadcrumb">
        <span class="bc-link" id="bc-all-conns">All Connections</span>
      </nav>`;
      el.querySelector('#bc-all-conns').addEventListener('click', () => App.navigate('connections'));
    } else if (orgName && projectName) {
      el.innerHTML = `<nav class="breadcrumb">
        <span class="bc-link" id="bc-all-conns">All Connections</span>
        <span class="bc-sep">›</span>
        <span class="bc-link" id="bc-org">${_esc(orgName)}</span>
        <span class="bc-sep">›</span>
        <span>${_esc(projectName)}</span>
      </nav>`;
      el.querySelector('#bc-all-conns').addEventListener('click', () => App.navigate('connections'));
      el.querySelector('#bc-org').addEventListener('click', () => App.navigate('projects'));
    } else {
      el.innerHTML = '';
    }
  }

  // ─── Sorting helper ────────────────────────────────────────────

  function _sortItems(items) {
    if (!_sortCol) return items;
    const dir = _sortDir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let va = '', vb = '';
      if (_sortCol === 'id')       { va = a.fields['System.Id'] || 0; vb = b.fields['System.Id'] || 0; return dir * (va - vb); }
      if (_sortCol === 'title')    { va = a.fields['System.Title'] || ''; vb = b.fields['System.Title'] || ''; }
      if (_sortCol === 'type')     { va = a.fields['System.WorkItemType'] || ''; vb = b.fields['System.WorkItemType'] || ''; }
      if (_sortCol === 'state')    { va = a.fields['System.State'] || ''; vb = b.fields['System.State'] || ''; }
      if (_sortCol === 'assigned') { va = String(a.fields['System.AssignedTo']?.displayName || a.fields['System.AssignedTo'] || ''); vb = String(b.fields['System.AssignedTo']?.displayName || b.fields['System.AssignedTo'] || ''); }
      if (_sortCol === 'priority') { va = a.fields['Microsoft.VSTS.Common.Priority'] || 99; vb = b.fields['Microsoft.VSTS.Common.Priority'] || 99; return dir * (va - vb); }
      if (_sortCol === 'created')  { va = a.fields['System.CreatedDate'] || ''; vb = b.fields['System.CreatedDate'] || ''; }
      if (_sortCol === 'updated')  { va = a.fields['System.ChangedDate'] || ''; vb = b.fields['System.ChangedDate'] || ''; }
      if (_sortCol === 'org')      { va = a._connName || ''; vb = b._connName || ''; }
      if (_sortCol === 'project')  { va = a._project || ''; vb = b._project || ''; }
      return dir * va.toString().localeCompare(vb.toString());
    });
  }

  function _thSortClass(col) {
    if (_sortCol !== col) return 'sortable';
    return `sortable sort-${_sortDir}`;
  }

  // ─── Render table ─────────────────────────────────────────────

  function _renderTable(container) {
    const search    = (container.querySelector('#wi-search')?.value || '').toLowerCase();
    const content   = container.querySelector('#wi-content');
    const showOrgCol = _mode === 'all';

    let filtered = _items;
    if (search) filtered = filtered.filter(w =>
      (w.fields['System.Title'] || '').toLowerCase().includes(search)
    );

    const typeF     = container.querySelector('#wi-type-filter')?.value || '';
    const stateF    = container.querySelector('#wi-state-filter')?.value || '';
    const priorityF = container.querySelector('#wi-priority-filter')?.value || '';
    if (typeF)     filtered = filtered.filter(w => (w.fields['System.WorkItemType'] || '') === typeF);
    if (stateF)    filtered = filtered.filter(w => (w.fields['System.State'] || '') === stateF);
    if (priorityF) filtered = filtered.filter(w => String(w.fields['Microsoft.VSTS.Common.Priority'] || '') === priorityF);

    // Sort
    filtered = _sortItems(filtered);

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

    const orgCols = showOrgCol
      ? `<th class="${_thSortClass('org')}" data-col="org">Organization</th><th class="${_thSortClass('project')}" data-col="project">Project</th>`
      : '';

    content.innerHTML = `
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table tabindex="0" id="wi-table">
          <thead>
            <tr>
              <th><input type="checkbox" id="wi-select-all" title="Select all" /></th>
              <th class="${_thSortClass('id')}" data-col="id">ID</th>
              ${orgCols}
              <th class="${_thSortClass('title')}" data-col="title">Title</th>
              <th class="${_thSortClass('type')}" data-col="type">Type</th>
              <th class="${_thSortClass('state')}" data-col="state">State</th>
              <th class="${_thSortClass('assigned')}" data-col="assigned">Assigned To</th>
              <th class="${_thSortClass('priority')}" data-col="priority">Priority</th>
              <th class="${_thSortClass('created')}" data-col="created">Created</th>
              <th class="${_thSortClass('updated')}" data-col="updated">Updated</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="wi-tbody">
            ${_buildGroupedRows(pageItems, showOrgCol)}
          </tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}
    `;

    // Resizable columns (Feature 4)
    _initResizableColumns(content.querySelector('#wi-table'));

    // Feature 2: Collapsible group headers
    _bindGroupToggles(content);

    // Sort click on th (Feature 3)
    content.querySelectorAll('th[data-col]').forEach(th => {
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

    // Select all checkbox
    const selectAll = content.querySelector('#wi-select-all');
    if (selectAll) {
      selectAll.addEventListener('change', () => {
        content.querySelectorAll('.wi-row-cb').forEach(cb => { cb.checked = selectAll.checked; });
        _updateBulkBar(container);
      });
    }

    // Row checkboxes
    content.querySelectorAll('.wi-row-cb').forEach(cb => {
      cb.addEventListener('change', () => _updateBulkBar(container));
    });

    // Row click → open detail modal (only when not clicking checkbox or action buttons)
    content.querySelectorAll('.wi-row').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('input[type=checkbox]')) return;
        if (e.target.closest('.wi-edit-btn')) return;
        if (e.target.closest('.wi-save-btn')) return;
        if (e.target.closest('.wi-cancel-btn')) return;
        if (e.target.closest('select')) return;
        _openDetail(row.dataset.id, row.dataset.connId);
      });
    });

    // Inline edit buttons
    content.querySelectorAll('.wi-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        _startInlineEdit(content, btn.dataset.id);
      });
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });

    // Feature 6: Keyboard navigation
    const table = content.querySelector('#wi-table');
    if (table) {
      table.addEventListener('keydown', e => {
        const rows = Array.from(table.querySelectorAll('tbody .wi-row'));
        const focused = table.querySelector('tr.row-focused');
        let idx = focused ? rows.indexOf(focused) : -1;

        if (e.key === 'ArrowDown') {
          e.preventDefault();
          idx = Math.min(idx + 1, rows.length - 1);
          rows.forEach(r => r.classList.remove('row-focused'));
          if (rows[idx]) rows[idx].classList.add('row-focused');
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          idx = Math.max(idx - 1, 0);
          rows.forEach(r => r.classList.remove('row-focused'));
          if (rows[idx]) rows[idx].classList.add('row-focused');
        } else if (e.key === 'Enter' && focused) {
          _openDetail(focused.dataset.id, focused.dataset.connId);
        } else if (e.key === 'Escape') {
          App.closeModal();
        }
      });
    }
  }

  // ─── Feature 2: Collapsible org groups ─────────────────────────

  function _buildGroupedRows(items, showOrgCol) {
    if (!showOrgCol) {
      return items.map(w => _rowHtml(w, false)).join('');
    }

    // Group by org name
    const groups = {};
    const groupOrder = [];
    items.forEach(w => {
      const key = w._connName || '(unknown)';
      if (!groups[key]) { groups[key] = []; groupOrder.push(key); }
      groups[key].push(w);
    });

    let html = '';
    groupOrder.forEach(orgName => {
      const groupItems = groups[orgName];
      const groupId = 'grp_' + orgName.replace(/\W/g, '_');
      html += `<tr class="group-header" data-group="${groupId}">
        <td colspan="11">
          <i class="fa-solid fa-chevron-down" style="margin-right:6px;font-size:.75rem"></i>
          ${_esc(orgName)}
          <span class="badge badge-secondary" style="margin-left:8px">${groupItems.length}</span>
        </td>
      </tr>`;
      html += groupItems.map(w => `<tr class="wi-row wi-group-row ${groupId}"
          data-id="${w.fields['System.Id'] || w.id}"
          ${w._connId ? `data-conn-id="${_esc(w._connId)}"` : ''}
          style="cursor:pointer">
        ${_rowCells(w, true)}
      </tr>`).join('');
    });
    return html;
  }

  function _rowHtml(w, showOrgCol) {
    const id = w.fields['System.Id'] || w.id;
    const connIdAttr = w._connId ? ` data-conn-id="${_esc(w._connId)}"` : '';
    return `<tr class="wi-row" data-id="${id}"${connIdAttr} style="cursor:pointer">
      ${_rowCells(w, showOrgCol)}
    </tr>`;
  }

  function _rowCells(w, showOrgCol) {
    const f       = w.fields;
    const id      = f['System.Id'] || w.id;
    const title   = f['System.Title'] || '(no title)';
    const type    = f['System.WorkItemType'] || '';
    const state   = f['System.State'] || '';
    const assigned= String(f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || '—');
    const priority= f['Microsoft.VSTS.Common.Priority'] || '—';
    const created = f['System.CreatedDate'] ? _fmtDate(f['System.CreatedDate']) : '—';
    const updated = f['System.ChangedDate'] ? _fmtDate(f['System.ChangedDate']) : '—';

    const orgCells = showOrgCol
      ? `<td class="text-sm">${_esc(w._connName || '')}</td><td class="text-sm">${_esc(w._project || '')}</td>`
      : '';

    return `
      <td><input type="checkbox" class="wi-row-cb" data-id="${id}" /></td>
      <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
      ${orgCells}
      <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
      <td id="wi-type-cell-${id}">${_typeBadge(type)}</td>
      <td id="wi-state-cell-${id}">${_stateBadge(state)}</td>
      <td id="wi-assigned-cell-${id}" class="truncate" style="max-width:140px">${_esc(assigned)}</td>
      <td>${_priorityBadge(priority)}</td>
      <td class="text-muted text-sm">${created}</td>
      <td class="text-muted text-sm">${updated}</td>
      <td><button class="btn-icon wi-edit-btn" data-id="${id}" title="Inline edit"><i class="fa-solid fa-pen-to-square"></i></button></td>
    `;
  }

  // ─── Feature 4: Resizable columns ─────────────────────────────

  function _initResizableColumns(table) {
    if (!table) return;
    table.querySelectorAll('thead th').forEach(th => {
      // Avoid adding multiple handles
      if (th.querySelector('.col-resize-handle')) return;
      const handle = document.createElement('span');
      handle.className = 'col-resize-handle';
      th.appendChild(handle);

      let startX, startW;
      handle.addEventListener('mousedown', e => {
        e.stopPropagation();
        startX = e.pageX;
        startW = th.offsetWidth;
        const onMove = mv => { th.style.width = (startW + mv.pageX - startX) + 'px'; };
        const onUp = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });
  }

  // ─── Feature 2: Group toggle ──────────────────────────────────

  function _bindGroupToggles(container) {
    container.querySelectorAll('.group-header').forEach(hdr => {
      hdr.addEventListener('click', () => {
        const groupId = hdr.dataset.group;
        const icon = hdr.querySelector('i');
        const rows = container.querySelectorAll(`.${groupId}`);
        rows.forEach(row => {
          row.classList.toggle('group-collapsed');
        });
        const isCollapsed = rows.length > 0 && rows[0].classList.contains('group-collapsed');
        if (icon) icon.className = isCollapsed ? 'fa-solid fa-chevron-right' : 'fa-solid fa-chevron-down';
      });
    });
  }

  // ─── Feature 9: Bulk actions ──────────────────────────────────

  function _updateBulkBar(container) {
    const checked = container.querySelectorAll('.wi-row-cb:checked');
    const bar = container.querySelector('#wi-bulk-bar');
    const count = container.querySelector('#wi-bulk-count');
    if (!bar) return;
    if (checked.length > 0) {
      bar.classList.remove('hidden');
      count.textContent = `${checked.length} selected`;
    } else {
      bar.classList.add('hidden');
    }
  }

  function _getCheckedIds(container) {
    return Array.from(container.querySelectorAll('.wi-row-cb:checked')).map(cb => cb.dataset.id);
  }

  async function _bulkApply(container) {
    const ids = _getCheckedIds(container);
    const newState = container.querySelector('#wi-bulk-state')?.value;
    if (!ids.length || !newState) { App.showToast('Select items and a target state.', 'warning'); return; }

    // Determine connection for each item
    const btn = container.querySelector('#wi-bulk-apply-btn');
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner" style="width:12px;height:12px;border-width:2px;vertical-align:middle;"></span> Applying…';

    let ok = 0, fail = 0;
    for (const id of ids) {
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      if (!item) { fail++; continue; }
      const conn = item._connId ? ConnectionsModule.getById(item._connId) : (_context?.conn || null);
      const project = item._project || _context?.project;
      if (!conn || !project) { fail++; continue; }
      const res = await AzureApi.updateWorkItemState(conn.orgUrl, project, id, newState, conn.pat);
      if (res.error) fail++;
      else {
        ok++;
        item.fields['System.State'] = newState;
      }
    }

    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-check"></i> Apply';
    App.showToast(`Bulk update: ${ok} succeeded, ${fail} failed.`, fail > 0 ? 'warning' : 'success');
    _renderTable(container);
  }

  // ─── Feature 10: Export CSV ────────────────────────────────────

  function _exportCsv(container, selectedOnly = false) {
    const search    = (container.querySelector('#wi-search')?.value || '').toLowerCase();
    const typeF     = container.querySelector('#wi-type-filter')?.value || '';
    const stateF    = container.querySelector('#wi-state-filter')?.value || '';
    const priorityF = container.querySelector('#wi-priority-filter')?.value || '';
    const showOrgCol = _mode === 'all';

    let items = _items;
    if (search)    items = items.filter(w => (w.fields['System.Title'] || '').toLowerCase().includes(search));
    if (typeF)     items = items.filter(w => (w.fields['System.WorkItemType'] || '') === typeF);
    if (stateF)    items = items.filter(w => (w.fields['System.State'] || '') === stateF);
    if (priorityF) items = items.filter(w => String(w.fields['Microsoft.VSTS.Common.Priority'] || '') === priorityF);

    if (selectedOnly) {
      const selectedIds = new Set(_getCheckedIds(container));
      items = items.filter(w => selectedIds.has(String(w.fields['System.Id'] || w.id)));
    }

    const headers = ['ID','Title','Type','State','AssignedTo','Priority','Organization','Project','Created','Updated'];
    const rows = items.map(w => {
      const f = w.fields;
      return [
        f['System.Id'] || w.id,
        f['System.Title'] || '',
        f['System.WorkItemType'] || '',
        f['System.State'] || '',
        String(f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || ''),
        f['Microsoft.VSTS.Common.Priority'] || '',
        w._connName || (_context?.conn?.name || ''),
        w._project || (_context?.project || ''),
        f['System.CreatedDate'] ? new Date(f['System.CreatedDate']).toISOString().slice(0,10) : '',
        f['System.ChangedDate'] ? new Date(f['System.ChangedDate']).toISOString().slice(0,10) : '',
      ].map(v => `"${String(v).replace(/"/g, '""')}"`);
    });

    const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const date = new Date().toISOString().slice(0,10);
    const blob = new Blob([csv], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = `workitems-${date}.csv`; a.click();
    URL.revokeObjectURL(url);
  }

  // ─── Feature 11: Inline edit ──────────────────────────────────

  function _startInlineEdit(content, id) {
    const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
    if (!item) return;

    const stateCell    = content.querySelector(`#wi-state-cell-${id}`);
    const assignedCell = content.querySelector(`#wi-assigned-cell-${id}`);
    if (!stateCell || !assignedCell) return;

    const currentState    = item.fields['System.State'] || '';
    const currentAssigned = String(item.fields['System.AssignedTo']?.displayName || item.fields['System.AssignedTo'] || '');

    stateCell.innerHTML = `<select class="wi-inline-state" style="font-size:.78rem;padding:2px 4px">
      ${['New','Active','Resolved','Closed','Done'].map(s =>
        `<option value="${s}" ${s === currentState ? 'selected' : ''}>${s}</option>`
      ).join('')}
    </select>`;

    assignedCell.innerHTML = `<input type="text" class="wi-inline-assigned" value="${_esc(currentAssigned)}" style="font-size:.78rem;padding:2px 4px;width:120px" />`;

    // Replace edit button with Save/Cancel
    const editBtn = content.querySelector(`.wi-edit-btn[data-id="${id}"]`);
    if (editBtn) {
      editBtn.parentElement.innerHTML = `
        <button class="btn btn-primary btn-sm wi-save-btn" data-id="${id}" style="padding:2px 6px;font-size:.72rem">Save</button>
        <button class="btn btn-secondary btn-sm wi-cancel-btn" data-id="${id}" style="padding:2px 6px;font-size:.72rem">Cancel</button>
      `;
      content.querySelector(`.wi-save-btn[data-id="${id}"]`).addEventListener('click', async e => {
        e.stopPropagation();
        await _saveInlineEdit(content, id);
      });
      content.querySelector(`.wi-cancel-btn[data-id="${id}"]`).addEventListener('click', e => {
        e.stopPropagation();
        _renderTable(document.getElementById('page-container'));
      });
    }
  }

  async function _saveInlineEdit(content, id) {
    const stateCell    = content.querySelector(`#wi-state-cell-${id}`);
    const assignedCell = content.querySelector(`#wi-assigned-cell-${id}`);
    if (!stateCell || !assignedCell) return;

    const newState    = stateCell.querySelector('select')?.value;
    const newAssigned = assignedCell.querySelector('input')?.value?.trim();

    const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
    if (!item) return;

    const conn = item._connId ? ConnectionsModule.getById(item._connId) : (_context?.conn || null);
    const project = item._project || _context?.project;
    if (!conn || !project) { App.showToast('Cannot determine connection for this item.', 'error'); return; }

    const fields = {};
    if (newState)    fields['System.State']      = newState;
    if (newAssigned !== undefined) fields['System.AssignedTo'] = newAssigned;

    const res = await AzureApi.updateWorkItemFields(conn.orgUrl, project, id, fields, conn.pat);
    if (res.error) { App.showToast('Update failed: ' + res.message, 'error'); return; }

    // Update in-memory item
    if (newState)    item.fields['System.State']      = newState;
    if (newAssigned !== undefined) item.fields['System.AssignedTo'] = newAssigned;

    App.showToast('Work item updated.', 'success');
    const container = document.getElementById('page-container');
    _renderTable(container);
  }

  // ─── Feature 12: Comments in modal ────────────────────────────

  async function _loadComments(orgUrl, project, id, pat) {
    const result = await AzureApi.getWorkItemComments(orgUrl, project, id, pat);
    if (result.error) return null;
    return result;
  }

  function _commentsHtml(commentsData, orgUrl, project, id, pat) {
    const comments = commentsData?.comments || [];
    const count    = commentsData?.count || 0;
    const commentItems = comments.map(c => `
      <div class="comment-item">
        <div class="comment-meta">
          <strong>${_esc(c.createdBy?.displayName || '—')}</strong>
          <span style="margin-left:6px">${c.createdDate ? new Date(c.createdDate).toLocaleString() : ''}</span>
        </div>
        <div class="comment-body">${c.text || ''}</div>
      </div>`).join('');

    return `
      <div class="comments-section" id="wi-comments-section">
        <div class="section-title">Comments (${count})</div>
        <div id="wi-comments-list">${commentItems || '<p class="text-muted text-sm">No comments yet.</p>'}</div>
        <div style="margin-top:12px">
          <textarea id="wi-comment-input" placeholder="Add a comment…" rows="3" style="width:100%;margin-bottom:8px"></textarea>
          <button class="btn btn-primary btn-sm" id="wi-post-comment-btn" data-id="${id}">
            <i class="fa-solid fa-paper-plane"></i> Post
          </button>
        </div>
      </div>`;
  }

  // ─── Feature 14: Linked work items in modal ────────────────────

  async function _loadLinks(conn, id) {
    const result = await AzureApi._fetchLinkedWorkItem
      ? AzureApi._fetchLinkedWorkItem(conn.orgUrl, id, conn.pat)
      : _fetch_linked(conn, id);
    return result;
  }

  async function _fetch_linked(conn, id) {
    // Fetch work item with relations expanded
    try {
      const base = conn.orgUrl.replace(/\/$/, '');
      const url  = `${base}/_apis/wit/workitems/${id}?$expand=relations&api-version=7.0`;
      const res  = await fetch(url, {
        headers: { 'Authorization': 'Basic ' + btoa(':' + conn.pat), 'Content-Type': 'application/json' },
      });
      if (!res.ok) return null;
      return await res.json();
    } catch { return null; }
  }

  async function _resolveLinkedIds(conn, relatedIds) {
    if (!relatedIds.length) return [];
    try {
      const base = conn.orgUrl.replace(/\/$/, '');
      const url  = `${base}/_apis/wit/workitems?ids=${relatedIds.join(',')}&fields=System.Id,System.Title,System.State,System.WorkItemType&api-version=7.0`;
      const res  = await fetch(url, {
        headers: { 'Authorization': 'Basic ' + btoa(':' + conn.pat), 'Content-Type': 'application/json' },
      });
      if (!res.ok) return [];
      const data = await res.json();
      return data.value || [];
    } catch { return []; }
  }

  function _linksHtml(linksWithDetails) {
    if (!linksWithDetails.length) return '';
    const rows = linksWithDetails.map(l => `
      <div class="link-item" data-linked-id="${l.linkedId}">
        <span class="link-rel-badge">${_esc(l.relType)}</span>
        <span class="font-mono text-sm" style="color:var(--color-primary)">#${l.linkedId}</span>
        <span style="flex:1">${_esc(l.title || '—')}</span>
        ${l.state ? _stateBadge(l.state) : ''}
        <span class="text-muted text-sm">${_esc(l.type || '')}</span>
      </div>`).join('');
    return `<div class="links-section" id="wi-links-section">
      <div class="section-title">Linked Work Items (${linksWithDetails.length})</div>
      ${rows}
    </div>`;
  }

  // ─── Feature 13: Create work item ─────────────────────────────

  function _openCreateModal(container) {
    if (!_context) return;
    const { project } = _context;

    App.openModal('New Work Item', `
      <form id="wi-create-form" class="modal-form" novalidate>
        <div class="form-group">
          <label for="wi-new-title">Title <span style="color:var(--color-danger)">*</span></label>
          <input type="text" id="wi-new-title" placeholder="Work item title" required />
        </div>
        <div class="form-group">
          <label for="wi-new-type">Type</label>
          <select id="wi-new-type">
            <option value="Bug">Bug</option>
            <option value="Task">Task</option>
            <option value="User Story">User Story</option>
            <option value="Epic">Epic</option>
            <option value="Feature">Feature</option>
          </select>
        </div>
        <div class="form-group">
          <label for="wi-new-assigned">Assigned To</label>
          <input type="text" id="wi-new-assigned" placeholder="Display name or email" />
        </div>
        <div class="form-group">
          <label for="wi-new-desc">Description</label>
          <textarea id="wi-new-desc" rows="3" placeholder="Description…"></textarea>
        </div>
        <div class="form-group">
          <label for="wi-new-priority">Priority</label>
          <select id="wi-new-priority">
            <option value="1">1 – Critical</option>
            <option value="2">2 – High</option>
            <option value="3" selected>3 – Medium</option>
            <option value="4">4 – Low</option>
          </select>
        </div>
        <div id="wi-create-error" class="form-error hidden"></div>
        <div style="display:flex;gap:8px;margin-top:12px">
          <button type="submit" class="btn btn-primary btn-sm" id="wi-create-submit-btn">
            <i class="fa-solid fa-plus"></i> Create
          </button>
          <button type="button" class="btn btn-secondary btn-sm" id="wi-create-cancel-btn">Cancel</button>
        </div>
      </form>
    `);

    document.getElementById('wi-create-cancel-btn')?.addEventListener('click', () => App.closeModal());

    document.getElementById('wi-create-form')?.addEventListener('submit', async e => {
      e.preventDefault();
      const titleVal   = document.getElementById('wi-new-title')?.value.trim();
      const type       = document.getElementById('wi-new-type')?.value;
      const assignedTo = document.getElementById('wi-new-assigned')?.value.trim();
      const desc       = document.getElementById('wi-new-desc')?.value.trim();
      const priority   = document.getElementById('wi-new-priority')?.value;
      const errEl      = document.getElementById('wi-create-error');

      if (!titleVal) {
        errEl.textContent = 'Title is required.';
        errEl.classList.remove('hidden');
        return;
      }

      const submitBtn = document.getElementById('wi-create-submit-btn');
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="spinner" style="width:12px;height:12px;border-width:2px;vertical-align:middle;"></span> Creating…';

      const conn = _context?.conn;
      const project = _context?.project;
      const res = await AzureApi.createWorkItem(conn.orgUrl, project, type, titleVal, desc, assignedTo, priority, conn.pat);

      submitBtn.disabled = false;
      submitBtn.innerHTML = '<i class="fa-solid fa-plus"></i> Create';

      if (res.error) {
        errEl.textContent = 'Failed: ' + res.message;
        errEl.classList.remove('hidden');
        return;
      }

      App.closeModal();
      App.showToast('Work item created successfully.', 'success');
      await _fetchItems(container);
    });
  }

  // ─── Detail modal (Features 12, 14) ───────────────────────────

  async function _openDetail(id, connIdOverride) {
    let conn;
    if (connIdOverride) conn = ConnectionsModule.getById(connIdOverride);
    if (!conn && _context) conn = _context.conn;
    if (!conn) {
      const item = _items.find(w => String(w.fields['System.Id'] || w.id) === String(id));
      if (item && item._connId) conn = ConnectionsModule.getById(item._connId) || null;
    }
    if (!conn) return;

    const project = _context?.project ||
      _items.find(w => String(w.fields['System.Id'] || w.id) === String(id))?._project || '';

    App.openModal('Work Item #' + id, '<div class="loading-state"><div class="spinner"></div><p>Loading…</p></div>');

    // Fetch detail, links and comments in parallel
    const [result, linkedRaw, commentsData] = await Promise.all([
      AzureApi.getWorkItemDetail(conn.orgUrl, id, conn.pat),
      _fetch_linked(conn, id),
      project ? _loadComments(conn.orgUrl, project, id, conn.pat) : Promise.resolve(null),
    ]);

    if (result.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i><p>${_esc(result.message)}</p></div>`);
      return;
    }

    const f = result.fields;

    // Parse links
    let linksHtml = '';
    if (linkedRaw && linkedRaw.relations) {
      const relMap = {
        'System.LinkTypes.Hierarchy-Reverse': 'Parent',
        'System.LinkTypes.Hierarchy-Forward': 'Child',
        'System.LinkTypes.Related': 'Related',
      };
      const relItems = linkedRaw.relations
        .filter(r => r.rel && relMap[r.rel])
        .map(r => {
          const parts = (r.url || '').split('/');
          const linkedId = parts[parts.length - 1];
          return { linkedId, relType: relMap[r.rel] };
        })
        .filter(r => r.linkedId && !isNaN(r.linkedId));

      if (relItems.length) {
        const details = await _resolveLinkedIds(conn, relItems.map(r => r.linkedId));
        const detailMap = {};
        details.forEach(d => { detailMap[d.id] = d; });
        const enriched = relItems.map(r => {
          const d = detailMap[r.linkedId];
          return { ...r, title: d?.fields?.['System.Title'] || '', state: d?.fields?.['System.State'] || '', type: d?.fields?.['System.WorkItemType'] || '' };
        });
        linksHtml = _linksHtml(enriched);
      }
    }

    const commHtml = (project && commentsData) ? _commentsHtml(commentsData, conn.orgUrl, project, id, conn.pat) : '';

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
      ${linksHtml}
      ${commHtml}
    `);

    // Bind link clicks (Feature 14)
    document.querySelectorAll('.link-item[data-linked-id]').forEach(el => {
      el.addEventListener('click', () => _openDetail(el.dataset.linkedId, connIdOverride));
    });

    // Bind post comment (Feature 12)
    function _bindPostComment() {
      document.getElementById('wi-post-comment-btn')?.addEventListener('click', async () => {
        const text = document.getElementById('wi-comment-input')?.value?.trim();
        if (!text) return;
        const postBtn = document.getElementById('wi-post-comment-btn');
        postBtn.disabled = true;
        postBtn.textContent = 'Posting…';
        const res = await AzureApi.addWorkItemComment(conn.orgUrl, project, id, text, conn.pat);
        postBtn.disabled = false;
        postBtn.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Post';
        if (res.error) { App.showToast('Failed to post comment: ' + res.message, 'error'); return; }
        // Refresh comments section
        const newComments = await _loadComments(conn.orgUrl, project, id, conn.pat);
        const section = document.getElementById('wi-comments-section');
        if (section && newComments) {
          section.outerHTML = _commentsHtml(newComments, conn.orgUrl, project, id, conn.pat);
          _bindPostComment();
        }
      });
    }
    _bindPostComment();
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
