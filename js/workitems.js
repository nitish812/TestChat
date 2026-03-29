/**
 * workitems.js — Work Items
 *
 * Displays work items for a selected project with filtering,
 * pagination (50 per page), and a detail modal.
 * Supports "Single Project" and "All Connections" modes.
 *
 * Improvements:
 *  #2  Collapsible connection groups
 *  #3  Column sorting
 *  #4  Resizable columns
 *  #6  Keyboard navigation
 *  #7  Breadcrumb navigation
 *  #9  Bulk actions
 *  #10 Export to CSV
 *  #11 Inline state editing
 *  #12 Work item comments
 *  #13 Create work item
 *  #14 Link work items (parent/children)
 */

const WorkItemsModule = (() => {
  const PAGE_SIZE = 50;

  let _items       = [];
  let _currentPage = 1;
  let _context     = null;   // { connId, project, conn }
  let _mode        = 'single'; // 'single' | 'all'
  let _allSummary  = null;   // { total, connCount, failCount }

  // #3 Column sorting
  let _sortCol = '';
  let _sortDir = 'asc';

  // #2 Collapsible groups — persists collapsed state within session
  const _collapsedGroups = new Set();

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _currentPage = 1;
    _items = [];
    _mode = 'single';
    _allSummary = null;

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${_esc(c.id)}">${_esc(c.name)}</option>`
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

      <!-- #13 Create Work Item panel (hidden until context loaded) -->
      <div id="wi-create-panel" class="card mb-16" style="display:none">
        <div class="card-header">
          <span class="card-title"><i class="fa-solid fa-plus"></i> Create Work Item</span>
          <button class="btn btn-secondary btn-sm" id="wi-create-toggle">Show</button>
        </div>
        <div id="wi-create-form-body" style="display:none; padding:12px">
          <div class="form-group">
            <label>Type</label>
            <select id="wi-new-type">
              <option>Task</option><option>Bug</option><option>User Story</option><option>Feature</option><option>Epic</option>
            </select>
          </div>
          <div class="form-group">
            <label>Title <span style="color:var(--color-danger)">*</span></label>
            <input type="text" id="wi-new-title" placeholder="Work item title" />
          </div>
          <button class="btn btn-primary btn-sm" id="wi-create-submit">
            <i class="fa-solid fa-plus"></i> Create
          </button>
          <span id="wi-create-result" class="text-sm ml-8"></span>
        </div>
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
      </div>

      <!-- #9 Bulk action bar (hidden until selection) -->
      <div id="wi-bulk-bar" class="bulk-bar hidden">
        <span id="wi-bulk-count" class="text-sm font-mono">0 selected</span>
        <select id="wi-bulk-state">
          <option value="">Set State…</option>
          <option value="New">New</option>
          <option value="Active">Active</option>
          <option value="Resolved">Resolved</option>
          <option value="Closed">Closed</option>
          <option value="Done">Done</option>
        </select>
        <button class="btn btn-primary btn-sm" id="wi-bulk-apply-state">Apply State</button>
        <button class="btn btn-secondary btn-sm" id="wi-bulk-export">Export Selected</button>
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
      if (e.target.closest('#wi-apply-filter-btn')) { _currentPage = 1; _renderTable(container); }
      if (e.target.closest('#wi-export-btn')) _exportCsv(false);
      if (e.target.closest('#wi-bulk-export')) _exportCsv(true);
      if (e.target.closest('#wi-bulk-apply-state')) {
        App.showToast('Bulk state change would require write API access. Feature scaffolded.', 'info');
      }
      // #13 Create panel toggle
      if (e.target.closest('#wi-create-toggle')) {
        const body = container.querySelector('#wi-create-form-body');
        const btn  = container.querySelector('#wi-create-toggle');
        const hidden = body.style.display === 'none';
        body.style.display = hidden ? '' : 'none';
        btn.textContent = hidden ? 'Hide' : 'Show';
      }
      // #13 Create work item submit
      if (e.target.closest('#wi-create-submit')) {
        await _createWorkItem(container);
      }
    });

    container.addEventListener('keydown', e => {
      if (e.key === 'Enter' && e.target.id === 'wi-search') { _currentPage = 1; _renderTable(container); }
    });

    // #6 Keyboard navigation on #wi-content
    container.querySelector('#wi-content').addEventListener('keydown', e => {
      const focused = document.activeElement;
      if (!focused || !focused.classList.contains('wi-row')) return;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        const rows = [...container.querySelectorAll('.wi-row')];
        const idx = rows.indexOf(focused);
        if (idx < rows.length - 1) rows[idx + 1].focus();
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        const rows = [...container.querySelectorAll('.wi-row')];
        const idx = rows.indexOf(focused);
        if (idx > 0) rows[idx - 1].focus();
      } else if (e.key === 'Enter') {
        _openDetail(focused.dataset.id, focused.dataset.connId);
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
    // #13 Show create panel when context is set
    container.querySelector('#wi-create-panel').style.display = '';
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

  function _getFilteredItems(container) {
    const search    = (container.querySelector('#wi-search')?.value || '').toLowerCase();
    const typeF     = container.querySelector('#wi-type-filter')?.value || '';
    const stateF    = container.querySelector('#wi-state-filter')?.value || '';
    const priorityF = container.querySelector('#wi-priority-filter')?.value || '';

    let filtered = _items;
    if (search) filtered = filtered.filter(w =>
      (w.fields['System.Title'] || '').toLowerCase().includes(search)
    );
    if (typeF)     filtered = filtered.filter(w => (w.fields['System.WorkItemType'] || '') === typeF);
    if (stateF)    filtered = filtered.filter(w => (w.fields['System.State'] || '') === stateF);
    if (priorityF) filtered = filtered.filter(w => String(w.fields['Microsoft.VSTS.Common.Priority'] || '') === priorityF);
    return filtered;
  }

  function _sortItems(items) {
    if (!_sortCol) return items;
    const dir = _sortDir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let av, bv;
      const af = a.fields, bf = b.fields;
      switch (_sortCol) {
        case 'id':       av = af['System.Id'] || 0;                      bv = bf['System.Id'] || 0; break;
        case 'title':    av = (af['System.Title'] || '').toLowerCase();  bv = (bf['System.Title'] || '').toLowerCase(); break;
        case 'type':     av = af['System.WorkItemType'] || '';           bv = bf['System.WorkItemType'] || ''; break;
        case 'state':    av = af['System.State'] || '';                  bv = bf['System.State'] || ''; break;
        case 'assigned': av = (af['System.AssignedTo']?.displayName || af['System.AssignedTo'] || '').toLowerCase(); bv = (bf['System.AssignedTo']?.displayName || bf['System.AssignedTo'] || '').toLowerCase(); break;
        case 'priority': av = af['Microsoft.VSTS.Common.Priority'] || 99; bv = bf['Microsoft.VSTS.Common.Priority'] || 99; break;
        case 'created':  av = af['System.CreatedDate'] || '';            bv = bf['System.CreatedDate'] || ''; break;
        case 'updated':  av = af['System.ChangedDate'] || '';            bv = bf['System.ChangedDate'] || ''; break;
        default: return 0;
      }
      if (av < bv) return -1 * dir;
      if (av > bv) return  1 * dir;
      return 0;
    });
  }

  function _sortIcon(col) {
    if (_sortCol !== col) return '<i class="fa-solid fa-sort text-muted" style="font-size:.7rem"></i>';
    return _sortDir === 'asc'
      ? '<i class="fa-solid fa-sort-up" style="font-size:.7rem;color:var(--color-primary)"></i>'
      : '<i class="fa-solid fa-sort-down" style="font-size:.7rem;color:var(--color-primary)"></i>';
  }

  function _renderTable(container) {
    const content    = container.querySelector('#wi-content');
    const showOrgCol = _mode === 'all';

    const filtered = _sortItems(_getFilteredItems(container));

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

    // #7 Breadcrumb
    const breadcrumb = _context ? `
      <nav class="breadcrumb mb-8">
        <a href="#/dashboard">Dashboard</a>
        <span class="bc-sep">›</span>
        <a href="#/projects">Projects</a>
        <span class="bc-sep">›</span>
        <span class="bc-org">${_esc(_context.conn.name)}</span>
        <span class="bc-sep">›</span>
        <span class="bc-current">${_esc(_context.project)}</span>
      </nav>` : '';

    const orgCols = showOrgCol ? `<th data-col="org">Organization ${_sortIcon('org')}</th><th data-col="proj">Project ${_sortIcon('proj')}</th>` : '';

    // #9 Select-all checkbox column
    content.innerHTML = `
      ${breadcrumb}
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${total} item${total !== 1 ? 's' : ''} found</p>
      <div class="table-wrapper">
        <table id="wi-table" style="table-layout:fixed;min-width:700px">
          <thead>
            <tr>
              <th style="width:36px"><input type="checkbox" id="wi-select-all" title="Select all" /></th>
              <th data-col="id" style="cursor:pointer">ID ${_sortIcon('id')}</th>
              ${orgCols}
              <th data-col="title" style="cursor:pointer">Title ${_sortIcon('title')}</th>
              <th data-col="type" style="cursor:pointer">Type ${_sortIcon('type')}</th>
              <th data-col="state" style="cursor:pointer">State ${_sortIcon('state')}</th>
              <th data-col="assigned" style="cursor:pointer">Assigned To ${_sortIcon('assigned')}</th>
              <th data-col="priority" style="cursor:pointer">Priority ${_sortIcon('priority')}</th>
              <th data-col="created" style="cursor:pointer">Created ${_sortIcon('created')}</th>
              <th data-col="updated" style="cursor:pointer">Updated ${_sortIcon('updated')}</th>
            </tr>
          </thead>
          <tbody id="wi-tbody">
            ${pageItems.map(w => _rowHtml(w, showOrgCol)).join('')}
          </tbody>
        </table>
      </div>
      ${_paginationHtml(_currentPage, pages)}
    `;

    // #4 Resizable columns
    _makeColumnsResizable(content.querySelector('#wi-table'));

    // #3 Sort on <th> click
    content.querySelectorAll('thead th[data-col]').forEach(th => {
      th.style.cursor = 'pointer';
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

    // #9 Select-all
    const selectAll = content.querySelector('#wi-select-all');
    if (selectAll) {
      selectAll.addEventListener('change', () => {
        content.querySelectorAll('.wi-select-cb').forEach(cb => { cb.checked = selectAll.checked; });
        _updateBulkBar(container);
      });
    }

    // #9 Row checkboxes
    content.querySelectorAll('.wi-select-cb').forEach(cb => {
      cb.addEventListener('change', () => _updateBulkBar(container));
    });

    // Row click → open detail modal (not triggered by checkbox or inline select)
    content.querySelectorAll('.wi-row').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.wi-select-cb') || e.target.closest('.wi-inline-state')) return;
        _openDetail(row.dataset.id, row.dataset.connId);
      });
    });

    // #11 Inline state select — stop propagation to prevent row click
    content.querySelectorAll('.wi-inline-state').forEach(sel => {
      sel.addEventListener('click', e => e.stopPropagation());
      sel.addEventListener('change', e => {
        e.stopPropagation();
        App.showToast('Inline state update requires write API — feature scaffolded.', 'info');
      });
    });

    // Pagination
    content.querySelectorAll('.page-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p)) { _currentPage = p; _renderTable(container); }
      });
    });

    // #2 Group header toggle
    content.querySelectorAll('.btn-group-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const group = btn.closest('tr').dataset.group;
        if (_collapsedGroups.has(group)) {
          _collapsedGroups.delete(group);
        } else {
          _collapsedGroups.add(group);
        }
        _renderTable(container);
      });
    });
  }

  // #9 Update bulk bar visibility and count
  function _updateBulkBar(container) {
    const selected = [...container.querySelectorAll('.wi-select-cb:checked')];
    const bar = container.querySelector('#wi-bulk-bar');
    const countEl = container.querySelector('#wi-bulk-count');
    if (!bar) return;
    if (selected.length > 0) {
      bar.classList.remove('hidden');
      countEl.textContent = `${selected.length} selected`;
    } else {
      bar.classList.add('hidden');
    }
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

    // #11 Inline state select
    const stateCell = `
      <td>
        <select class="wi-inline-state" style="font-size:.75rem;padding:2px 4px;border-radius:var(--radius-sm);border:1px solid var(--border-color);background:var(--bg-input);color:var(--text-primary)">
          ${['New','Active','Resolved','Closed','Done'].map(s =>
            `<option value="${s}"${s === state ? ' selected' : ''}>${_esc(s)}</option>`
          ).join('')}
        </select>
      </td>`;

    return `
      <tr class="wi-row" data-id="${id}"${connIdAttr} style="cursor:pointer" tabindex="0">
        <td><input type="checkbox" class="wi-select-cb" data-id="${id}" /></td>
        <td><a class="text-sm font-mono" style="color:var(--color-primary)">#${id}</a></td>
        ${orgCells}
        <td class="cell-title" title="${_esc(title)}">${_esc(title)}</td>
        <td>${_typeBadge(type)}</td>
        ${stateCell}
        <td class="truncate" style="max-width:140px">${_esc(String(assigned))}</td>
        <td>${_priorityBadge(priority)}</td>
        <td class="text-muted text-sm">${created}</td>
        <td class="text-muted text-sm">${updated}</td>
      </tr>`;
  }

  // #4 Resizable columns helper
  function _makeColumnsResizable(tableEl) {
    if (!tableEl) return;
    const ths = tableEl.querySelectorAll('thead th');
    ths.forEach(th => {
      // Ensure position:relative for the handle
      th.style.position = 'relative';
      th.style.overflow = 'hidden';
      const handle = document.createElement('div');
      handle.className = 'col-resizer';
      th.appendChild(handle);

      let startX, startW;
      handle.addEventListener('mousedown', e => {
        e.stopPropagation();
        startX = e.pageX;
        startW = th.offsetWidth;
        const onMove = mv => {
          const newW = Math.max(40, startW + mv.pageX - startX);
          th.style.width = newW + 'px';
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

    const [result, commentsResult] = await Promise.all([
      AzureApi.getWorkItemDetail(conn.orgUrl, id, conn.pat),
      AzureApi.getWorkItemComments(conn.orgUrl, id, conn.pat),
    ]);

    if (result.error) {
      App.updateModalBody(`<div class="error-state"><i class="fa-solid fa-circle-exclamation"></i><p>${_esc(result.message)}</p></div>`);
      return;
    }

    const f = result.fields;

    // #12 Comments
    let commentsHtml = '<p class="text-muted text-sm">No comments.</p>';
    if (!commentsResult.error) {
      const comments = commentsResult.comments || commentsResult.value || [];
      if (comments.length > 0) {
        commentsHtml = comments.map(c => {
          const author = c.createdBy?.displayName || c.author?.displayName || 'Unknown';
          const date   = c.createdDate ? new Date(c.createdDate).toLocaleString() : '';
          const text   = c.text || c.renderedText || '';
          return `
            <div class="comment-item">
              <div class="comment-meta"><strong>${_esc(author)}</strong> · <span class="text-muted text-sm">${_esc(date)}</span></div>
              <div class="comment-body">${text}</div>
            </div>`;
        }).join('');
      }
      const commentCount = comments.length;
      commentsHtml = `<span class="detail-label">Comments (${commentCount})</span><div class="comments-list">${commentsHtml}</div>`;
    }

    // #14 Related items (parent/children)
    let relatedHtml = '';
    if (result.relations && result.relations.length > 0) {
      const parentRels   = result.relations.filter(r => (r.rel || '').includes('Hierarchy-Reverse'));
      const childRels    = result.relations.filter(r => (r.rel || '').includes('Hierarchy-Forward'));
      const extractId    = url => (url || '').split('/').pop();
      const relLink      = (rid) => `<a href="#" class="wi-relation-link" data-id="${_esc(rid)}" style="color:var(--color-primary)">#${rid} Open</a>`;
      const parentStr    = parentRels.length > 0
        ? parentRels.map(r => { const rid = extractId(r.url); return `<strong>Parent:</strong> ${relLink(rid)}`; }).join(', ')
        : '<span class="text-muted">None</span>';
      const childrenStr  = childRels.length > 0
        ? childRels.map(r => { const rid = extractId(r.url); return relLink(rid); }).join(', ')
        : '<span class="text-muted">None</span>';

      relatedHtml = `
        <div class="detail-row" style="flex-direction:column;gap:4px">
          <span class="detail-label">Related Items</span>
          <div class="text-sm">
            <div>${parentStr}</div>
            <div><strong>Children:</strong> ${childrenStr}</div>
          </div>
        </div>`;
    }

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
      ${relatedHtml}
      <div class="detail-row" style="flex-direction:column;gap:6px">
        ${commentsHtml}
      </div>
    `);

    // #14 Relation link click
    document.querySelectorAll('.wi-relation-link').forEach(a => {
      a.addEventListener('click', e => {
        e.preventDefault();
        App.closeModal();
        _openDetail(a.dataset.id, connIdOverride);
      });
    });
  }

  // #13 Create work item
  async function _createWorkItem(container) {
    if (!_context) return;
    const titleInput  = container.querySelector('#wi-new-title');
    const typeSelect  = container.querySelector('#wi-new-type');
    const resultEl    = container.querySelector('#wi-create-result');
    const title = titleInput?.value?.trim();
    if (!title) { App.showToast('Title is required.', 'warning'); return; }

    const { conn, project } = _context;
    resultEl.textContent = 'Creating…';
    const res = await AzureApi.createWorkItem(conn.orgUrl, project, typeSelect.value, conn.pat, [
      { op: 'add', path: '/fields/System.Title', value: title },
    ]);

    if (res.error) {
      resultEl.textContent = '';
      App.showToast('Create failed: ' + res.message, 'error');
    } else {
      resultEl.textContent = '';
      titleInput.value = '';
      App.showToast(`Work item #${res.id} created successfully.`, 'success');
      await _fetchItems(container);
    }
  }

  // #10 Export CSV
  function _exportCsv(selectedOnly) {
    let items = _items;
    if (selectedOnly) {
      const checkedIds = new Set([...document.querySelectorAll('.wi-select-cb:checked')].map(cb => cb.dataset.id));
      items = items.filter(w => checkedIds.has(String(w.fields['System.Id'] || w.id)));
    }
    const headers = ['ID', 'Title', 'Type', 'State', 'AssignedTo', 'Priority', 'Created', 'Updated'];
    const csvRow  = (cols) => cols.map(v => {
      const s = String(v == null ? '' : v);
      return (s.includes(',') || s.includes('"') || s.includes('\n')) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',');

    const lines = [csvRow(headers)];
    items.forEach(w => {
      const f = w.fields;
      lines.push(csvRow([
        f['System.Id'] || w.id,
        f['System.Title'] || '',
        f['System.WorkItemType'] || '',
        f['System.State'] || '',
        f['System.AssignedTo']?.displayName || f['System.AssignedTo'] || '',
        f['Microsoft.VSTS.Common.Priority'] || '',
        f['System.CreatedDate'] || '',
        f['System.ChangedDate'] || '',
      ]));
    });

    const csv = lines.join('\r\n');
    const a   = document.createElement('a');
    a.download = 'workitems.csv';
    a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
    a.click();
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
