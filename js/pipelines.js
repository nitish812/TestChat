/**
 * pipelines.js — Pipelines / Builds
 *
 * Fetches build definitions and recent runs from a selected project
 * and renders them in a table with color-coded status badges.
 *
 * Features:
 *  - Single Project / All Connections mode
 *  - Collapsible org groups (All Connections mode)
 *  - Sortable columns
 *  - Resizable columns
 *  - Keyboard navigation
 *  - Breadcrumb navigation
 */

const PipelinesModule = (() => {
  let _pipelines   = [];   // merged build definition + last run
  let _context     = null;
  let _mode        = 'single'; // 'single' | 'all'
  let _allSummary  = null;
  let _sortCol     = null;
  let _sortDir     = 'asc';
  let _collapsedGroups = new Set();

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _pipelines = [];
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
      <div class="page-title"><i class="fa-solid fa-rocket"></i> Pipelines</div>
      <p class="page-subtitle">Monitor build and release pipelines across your projects.</p>

      <!-- Mode toggle -->
      <div class="filter-bar mb-8" id="pipe-mode-bar">
        <label style="font-weight:600">View:</label>
        <button class="btn btn-primary btn-sm active" id="pipe-mode-single">Single Project</button>
        <button class="btn btn-secondary btn-sm" id="pipe-mode-all">All Connections</button>
      </div>

      <!-- Single project selectors -->
      <div class="filter-bar mb-16" id="pipe-single-bar">
        <select id="pipe-conn-select" style="min-width:160px">
          <option value="">Select Organization…</option>
          ${connOptions}
        </select>
        <select id="pipe-project-select" style="min-width:180px" disabled>
          <option value="">Select Project…</option>
        </select>
        <button class="btn btn-primary btn-sm" id="pipe-load-btn" disabled>
          <i class="fa-solid fa-download"></i> Load
        </button>
      </div>

      <!-- All connections bar -->
      <div class="filter-bar mb-16" id="pipe-all-bar" style="display:none">
        <button class="btn btn-primary btn-sm" id="pipe-load-all-btn">
          <i class="fa-solid fa-download"></i> Load All
        </button>
      </div>

      <!-- Filters (hidden until data loaded) -->
      <div class="filter-bar mb-8" id="pipe-filters" style="display:none">
        <input type="text" id="pipe-search" placeholder="🔍 Search pipelines…" style="min-width:160px" />
        <select id="pipe-status-filter">
          <option value="">All Statuses</option>
          <option value="succeeded">Succeeded</option>
          <option value="failed">Failed</option>
          <option value="inProgress">In Progress</option>
          <option value="cancelled">Cancelled</option>
          <option value="notStarted">Not Started</option>
        </select>
      </div>

      <!-- Breadcrumb -->
      <div id="pipe-breadcrumb" class="breadcrumb mb-8" style="display:none"></div>

      <div id="pipe-content"></div>
    `;

    _bindSelectorEvents(container);

    // Pre-populate from params (e.g. navigated from a project card)
    if (params.connId) {
      const sel = container.querySelector('#pipe-conn-select');
      sel.value = params.connId;
      await _loadProjects(container, params.connId);
      if (params.project) {
        const pSel = container.querySelector('#pipe-project-select');
        pSel.value = params.project;
        if (pSel.value) {
          container.querySelector('#pipe-load-btn').disabled = false;
          await _fetchPipelines(container);
        }
      }
    }
  }

  async function refresh(container) {
    if (_context) await _fetchPipelines(container);
    else if (_mode === 'all') await _fetchAllPipelines(container);
  }

  // ─── Private ───────────────────────────────────────────────────

  function _bindSelectorEvents(container) {
    // Mode toggle
    container.querySelector('#pipe-mode-single').addEventListener('click', () => {
      if (_mode === 'single') return;
      _mode = 'single';
      container.querySelector('#pipe-mode-single').className = 'btn btn-primary btn-sm active';
      container.querySelector('#pipe-mode-all').className = 'btn btn-secondary btn-sm';
      container.querySelector('#pipe-single-bar').style.display = '';
      container.querySelector('#pipe-all-bar').style.display = 'none';
      container.querySelector('#pipe-filters').style.display = 'none';
      container.querySelector('#pipe-breadcrumb').style.display = 'none';
      container.querySelector('#pipe-content').innerHTML = '';
      _pipelines = [];
      _context = null;
    });

    container.querySelector('#pipe-mode-all').addEventListener('click', () => {
      if (_mode === 'all') return;
      _mode = 'all';
      container.querySelector('#pipe-mode-all').className = 'btn btn-primary btn-sm active';
      container.querySelector('#pipe-mode-single').className = 'btn btn-secondary btn-sm';
      container.querySelector('#pipe-single-bar').style.display = 'none';
      container.querySelector('#pipe-all-bar').style.display = '';
      container.querySelector('#pipe-filters').style.display = 'none';
      container.querySelector('#pipe-breadcrumb').style.display = 'none';
      container.querySelector('#pipe-content').innerHTML = '';
      _pipelines = [];
      _context = null;
    });

    container.addEventListener('change', async e => {
      if (e.target.id === 'pipe-conn-select') {
        const connId = e.target.value;
        container.querySelector('#pipe-project-select').innerHTML = '<option value="">Loading…</option>';
        container.querySelector('#pipe-project-select').disabled = true;
        container.querySelector('#pipe-load-btn').disabled = true;
        if (connId) await _loadProjects(container, connId);
      }
      if (e.target.id === 'pipe-project-select') {
        container.querySelector('#pipe-load-btn').disabled = !e.target.value;
      }
      if (e.target.id === 'pipe-status-filter') _renderPipelineList(container);
    });

    container.addEventListener('click', async e => {
      if (e.target.closest('#pipe-load-btn'))     await _fetchPipelines(container);
      if (e.target.closest('#pipe-load-all-btn')) await _fetchAllPipelines(container);

      // Group header toggle
      const groupHdr = e.target.closest('.group-header-row');
      if (groupHdr) {
        const org = groupHdr.dataset.org;
        if (_collapsedGroups.has(org)) { _collapsedGroups.delete(org); } else { _collapsedGroups.add(org); }
        _renderPipelineList(container);
        return;
      }
    });

    container.addEventListener('input', e => {
      if (e.target.id === 'pipe-search') _renderPipelineList(container);
    });

    // Keyboard navigation
    container.addEventListener('keydown', e => {
      const table = container.querySelector('table');
      if (!table) return;
      const focused = table.querySelector('.row-focused');
      if (!focused && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        const first = table.querySelector('tbody tr:not(.group-header-row)');
        if (first) { first.classList.add('row-focused'); first.setAttribute('tabindex', '0'); first.focus(); }
        e.preventDefault();
        return;
      }
      if (!focused) return;
      if (e.key === 'ArrowDown') {
        let next = focused.nextElementSibling;
        while (next && next.classList.contains('group-header-row')) next = next.nextElementSibling;
        if (next) { focused.classList.remove('row-focused'); next.classList.add('row-focused'); next.setAttribute('tabindex', '0'); next.focus(); }
        e.preventDefault();
      } else if (e.key === 'ArrowUp') {
        let prev = focused.previousElementSibling;
        while (prev && prev.classList.contains('group-header-row')) prev = prev.previousElementSibling;
        if (prev) { focused.classList.remove('row-focused'); prev.classList.add('row-focused'); prev.setAttribute('tabindex', '0'); prev.focus(); }
        e.preventDefault();
      }
    });
  }

  async function _loadProjects(container, connId) {
    const conn = ConnectionsModule.getById(connId);
    if (!conn) return;
    const result = await AzureApi.getProjects(conn.orgUrl, conn.pat);
    if (result.error) { App.showToast(result.message, 'error'); return; }
    const pSel = container.querySelector('#pipe-project-select');
    pSel.innerHTML = '<option value="">Select Project…</option>' +
      (result.value || []).map(p => `<option value="${_esc(p.name)}">${_esc(p.name)}</option>`).join('');
    pSel.disabled = false;
  }

  async function _fetchPipelines(container) {
    const connId  = container.querySelector('#pipe-conn-select').value;
    const project = container.querySelector('#pipe-project-select').value;
    if (!connId || !project) return;

    const conn = ConnectionsModule.getById(connId);
    if (!conn) return;

    _context = { connId, project, conn };
    _pipelines = [];

    const content = container.querySelector('#pipe-content');
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading pipelines…</p></div>';
    container.querySelector('#pipe-filters').style.display = 'none';
    container.querySelector('#pipe-breadcrumb').style.display = 'none';

    const [defsResult, buildsResult] = await Promise.all([
      AzureApi.getBuildDefinitions(conn.orgUrl, project, conn.pat),
      AzureApi.getBuildRuns(conn.orgUrl, project, conn.pat),
    ]);

    if (defsResult.error && buildsResult.error) {
      content.innerHTML = `
        <div class="error-state">
          <i class="fa-solid fa-circle-exclamation"></i>
          <p>${_esc(defsResult.message)}</p>
          ${defsResult.cors ? '<p class="text-sm">CORS may be blocking requests. See README for details.</p>' : ''}
        </div>`;
      return;
    }

    const definitions = defsResult.value || [];
    const builds      = buildsResult.value || [];
    const lastBuildMap = {};
    builds.forEach(b => {
      const defId = b.definition?.id;
      if (defId && !lastBuildMap[defId]) lastBuildMap[defId] = b;
    });

    _pipelines = definitions.map(def => ({
      def,
      lastBuild: lastBuildMap[def.id] || null,
      _connName: conn.name,
      _project: project,
      _connId: conn.id,
    }));

    container.querySelector('#pipe-filters').style.display = '';
    _renderBreadcrumb(container, conn.name, project);
    _renderPipelineList(container);
  }

  async function _fetchAllPipelines(container) {
    const activeConns = ConnectionsModule.getActive().filter(c => c.defaultProject);

    if (activeConns.length === 0) {
      container.querySelector('#pipe-content').innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-plug-circle-xmark"></i>
          <p>No connections have a default project configured. <a href="#/connections" style="color:var(--color-primary)">Go to Connections</a> to set one.</p>
        </div>`;
      return;
    }

    _context = null;
    _pipelines = [];
    const content = container.querySelector('#pipe-content');
    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading pipelines from all connections…</p></div>';
    container.querySelector('#pipe-filters').style.display = 'none';

    const results = await Promise.allSettled(
      activeConns.map(conn =>
        Promise.all([
          AzureApi.getBuildDefinitions(conn.orgUrl, conn.defaultProject, conn.pat),
          AzureApi.getBuildRuns(conn.orgUrl, conn.defaultProject, conn.pat),
        ]).then(([defs, builds]) => ({ conn, defs, builds }))
      )
    );

    let allPipelines = [];
    let failCount = 0;

    for (const outcome of results) {
      if (outcome.status === 'rejected') { failCount++; continue; }
      const { conn, defs, builds } = outcome.value;
      if (defs.error) { failCount++; continue; }
      const definitions = defs.value || [];
      const buildsArr   = builds.value || [];
      const lastBuildMap = {};
      buildsArr.forEach(b => {
        const defId = b.definition?.id;
        if (defId && !lastBuildMap[defId]) lastBuildMap[defId] = b;
      });
      definitions.forEach(def => {
        allPipelines.push({
          def,
          lastBuild: lastBuildMap[def.id] || null,
          _connName: conn.name,
          _project: conn.defaultProject,
          _connId: conn.id,
        });
      });
    }

    _pipelines = allPipelines;
    _allSummary = { total: allPipelines.length, connCount: activeConns.length - failCount, failCount };
    container.querySelector('#pipe-filters').style.display = '';
    _renderBreadcrumb(container, 'All Connections', null);
    _renderPipelineList(container);
  }

  function _renderBreadcrumb(container, orgName, projectName) {
    const bc = container.querySelector('#pipe-breadcrumb');
    if (!bc) return;
    let parts = `<a href="#/connections" class="bc-link">All Connections</a>`;
    if (orgName) parts += ` <span class="bc-sep">›</span> <span class="bc-cur">${_esc(orgName)}</span>`;
    if (projectName) parts += ` <span class="bc-sep">›</span> <span class="bc-cur">${_esc(projectName)}</span>`;
    bc.innerHTML = parts;
    bc.style.display = '';
  }

  function _renderPipelineList(container) {
    const content   = container.querySelector('#pipe-content');
    const search    = (container.querySelector('#pipe-search')?.value || '').toLowerCase();
    const statusF   = container.querySelector('#pipe-status-filter')?.value || '';
    const showOrgCol = _mode === 'all';

    let filtered = _pipelines;

    if (search) filtered = filtered.filter(p =>
      (p.def.name || '').toLowerCase().includes(search)
    );

    if (statusF) filtered = filtered.filter(p => {
      const s = _buildStatus(p.lastBuild);
      return s.key === statusF;
    });

    // Sort
    if (_sortCol) filtered = _sortPipelines(filtered, _sortCol, _sortDir);

    if (filtered.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-rocket"></i>
          <p>No pipelines match your filter.</p>
        </div>`;
      return;
    }

    const summaryBanner = (showOrgCol && _allSummary) ? `
      <div class="filter-bar mb-8" style="flex-wrap:wrap;gap:8px">
        <span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> ${_allSummary.total} pipeline${_allSummary.total !== 1 ? 's' : ''} from ${_allSummary.connCount} connection${_allSummary.connCount !== 1 ? 's' : ''}</span>
        ${_allSummary.failCount ? `<span class="badge badge-warning"><i class="fa-solid fa-triangle-exclamation"></i> ${_allSummary.failCount} connection${_allSummary.failCount !== 1 ? 's' : ''} failed</span>` : ''}
      </div>` : '';

    const orgCols = showOrgCol ? `
      <th class="sortable-th" data-col="org" style="cursor:pointer">Organization ${_sortIndicator('org')}</th>
      <th class="sortable-th" data-col="project" style="cursor:pointer">Project ${_sortIndicator('project')}</th>` : '';

    let tableBody = '';
    if (showOrgCol) {
      const groups = {};
      filtered.forEach(p => {
        const key = p._connName || '(unknown)';
        if (!groups[key]) groups[key] = [];
        groups[key].push(p);
      });
      const colCount = 8;
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
          items.forEach(p => { tableBody += _pipelineRow(p, showOrgCol); });
        }
      }
    } else {
      filtered.forEach(p => { tableBody += _pipelineRow(p, showOrgCol); });
    }

    content.innerHTML = `
      ${summaryBanner}
      <p class="text-sm text-muted mb-8">${filtered.length} pipeline${filtered.length !== 1 ? 's' : ''}</p>
      <div class="table-wrapper">
        <table tabindex="0">
          <thead>
            <tr>
              ${orgCols}
              <th class="sortable-th" data-col="name" style="cursor:pointer">Pipeline ${_sortIndicator('name')}</th>
              <th class="sortable-th" data-col="status" style="cursor:pointer">Last Status ${_sortIndicator('status')}</th>
              <th class="sortable-th" data-col="branch" style="cursor:pointer">Branch ${_sortIndicator('branch')}</th>
              <th>Triggered By</th>
              <th>Duration</th>
              <th class="sortable-th" data-col="lastRun" style="cursor:pointer">Last Run ${_sortIndicator('lastRun')}</th>
            </tr>
          </thead>
          <tbody>
            ${tableBody}
          </tbody>
        </table>
      </div>`;

    // Sortable headers
    content.querySelectorAll('.sortable-th').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.col;
        if (_sortCol === col) { _sortDir = _sortDir === 'asc' ? 'desc' : 'asc'; }
        else { _sortCol = col; _sortDir = 'asc'; }
        _renderPipelineList(container);
      });
    });

    // Resizable columns
    const table = content.querySelector('table');
    if (table) _initResizableColumns(table);
  }

  function _sortPipelines(pipelines, col, dir) {
    const mult = dir === 'asc' ? 1 : -1;
    return [...pipelines].sort((a, b) => {
      let va, vb;
      switch (col) {
        case 'name':    va = a.def.name || ''; vb = b.def.name || ''; break;
        case 'status':  va = _buildStatus(a.lastBuild).label; vb = _buildStatus(b.lastBuild).label; break;
        case 'branch':  va = a.lastBuild?.sourceBranch || ''; vb = b.lastBuild?.sourceBranch || ''; break;
        case 'lastRun': va = a.lastBuild?.finishTime || a.lastBuild?.startTime || ''; vb = b.lastBuild?.finishTime || b.lastBuild?.startTime || ''; break;
        case 'org':     va = a._connName || ''; vb = b._connName || ''; break;
        case 'project': va = a._project || ''; vb = b._project || ''; break;
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
    table.querySelectorAll('th').forEach(th => {
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

  function _pipelineRow(entry, showOrgCol) {
    const def   = entry.def;
    const build = entry.lastBuild;
    const status= _buildStatus(build);

    const branch    = build?.sourceBranch?.replace('refs/heads/', '') || '—';
    const triggeredBy = build?.requestedFor?.displayName || build?.requestedBy?.displayName || '—';
    const duration  = build ? _duration(build.startTime, build.finishTime) : '—';
    const lastRun   = build?.finishTime
      ? _relativeTime(new Date(build.finishTime))
      : (build?.startTime ? _relativeTime(new Date(build.startTime)) : '—');

    const orgCells = showOrgCol
      ? `<td class="text-sm">${_esc(entry._connName || '')}</td><td class="text-sm">${_esc(entry._project || '')}</td>`
      : '';

    return `
      <tr>
        ${orgCells}
        <td>
          <div style="font-weight:600;font-size:.85rem">${_esc(def.name)}</div>
          <div class="text-muted text-sm">#${def.id}</div>
        </td>
        <td>
          <span class="badge ${status.cls}">
            ${status.icon} ${status.label}
          </span>
        </td>
        <td class="text-sm">${branch !== '—' ? `<i class="fa-solid fa-code-branch text-muted"></i> ${_esc(branch)}` : '—'}</td>
        <td class="text-sm truncate" style="max-width:120px">${_esc(triggeredBy)}</td>
        <td class="text-sm text-muted">${duration}</td>
        <td class="text-sm text-muted">${lastRun}</td>
      </tr>`;
  }

  // ─── Status helpers ────────────────────────────────────────────

  function _buildStatus(build) {
    if (!build) return { key: 'notStarted', label: 'No runs', cls: 'status-notstarted', icon: '⚪' };

    const result = (build.result || '').toLowerCase();
    const status = (build.status || '').toLowerCase();

    if (result === 'succeeded')        return { key: 'succeeded',  label: 'Succeeded',   cls: 'status-succeeded',  icon: '🟢' };
    if (result === 'failed')           return { key: 'failed',     label: 'Failed',       cls: 'status-failed',     icon: '🔴' };
    if (result === 'canceled')         return { key: 'cancelled',  label: 'Cancelled',    cls: 'status-cancelled',  icon: '⚪' };
    if (result === 'partiallySucceeded') return { key: 'failed',   label: 'Partial',      cls: 'status-inprogress', icon: '🟡' };
    if (status === 'inprogress')       return { key: 'inProgress', label: 'In Progress',  cls: 'status-inprogress', icon: '🟡' };
    if (status === 'notstarted')       return { key: 'notStarted', label: 'Not Started',  cls: 'status-notstarted', icon: '⚪' };

    return { key: 'unknown', label: build.status || 'Unknown', cls: 'status-unknown', icon: '⚪' };
  }

  function _duration(start, finish) {
    if (!start) return '—';
    const end  = finish ? new Date(finish) : new Date();
    const ms   = end - new Date(start);
    const secs = Math.floor(ms / 1000);
    if (secs < 60)   return secs + 's';
    const mins = Math.floor(secs / 60);
    if (mins < 60)   return `${mins}m ${secs % 60}s`;
    return `${Math.floor(mins / 60)}h ${mins % 60}m`;
  }

  function _relativeTime(date) {
    const diff = (Date.now() - date) / 1000;
    if (diff < 60)    return 'just now';
    if (diff < 3600)  return `${Math.floor(diff/60)}m ago`;
    if (diff < 86400) return `${Math.floor(diff/3600)}h ago`;
    return `${Math.floor(diff/86400)}d ago`;
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
