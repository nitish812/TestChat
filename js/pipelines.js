/**
 * pipelines.js — Pipelines / Builds
 *
 * Fetches build definitions and recent runs from a selected project
 * and renders them in a table with color-coded status badges.
 *
 * Features:
 *  - Collapsible org groups (Feature 2) — All Connections mode
 *  - Sortable columns (Feature 3)
 *  - Keyboard navigation (Feature 6)
 *  - Breadcrumb nav (Feature 7)
 */

const PipelinesModule = (() => {
  let _pipelines   = [];   // merged build definition + last run
  let _context     = null;
  let _sortCol     = '';
  let _sortDir     = 'asc';

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _pipelines = [];
    _sortCol = '';
    _sortDir = 'asc';

    const connections = ConnectionsModule.getActive();
    const connOptions = connections.map(c =>
      `<option value="${c.id}">${_esc(c.name)}</option>`
    ).join('');

    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-rocket"></i> Pipelines</div>
      <p class="page-subtitle">Monitor build and release pipelines across your projects.</p>

      <div class="filter-bar mb-16">
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

      <div id="pipe-breadcrumb"></div>
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
  }

  // ─── Private ───────────────────────────────────────────────────

  function _bindSelectorEvents(container) {
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
      if (e.target.closest('#pipe-load-btn')) await _fetchPipelines(container);
    });

    container.addEventListener('input', e => {
      if (e.target.id === 'pipe-search') _renderPipelineList(container);
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
    container.querySelector('#pipe-breadcrumb').innerHTML = '';

    // Fetch build definitions + latest builds in parallel
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

    // Map each definition to its last build run
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
      _connId: connId,
    }));

    container.querySelector('#pipe-filters').style.display = '';
    _renderBreadcrumb(container, conn.name, project);
    _renderPipelineList(container);
  }

  // ─── Feature 7: Breadcrumb ─────────────────────────────────────

  function _renderBreadcrumb(container, orgName, projectName) {
    const el = container.querySelector('#pipe-breadcrumb');
    if (!el) return;
    el.innerHTML = `<nav class="breadcrumb">
      <span class="bc-link" id="bc-all-conns">All Connections</span>
      <span class="bc-sep">›</span>
      <span class="bc-link" id="bc-org">${_esc(orgName)}</span>
      <span class="bc-sep">›</span>
      <span>${_esc(projectName)}</span>
    </nav>`;
    el.querySelector('#bc-all-conns').addEventListener('click', () => App.navigate('connections'));
    el.querySelector('#bc-org').addEventListener('click', () => App.navigate('projects'));
  }

  // ─── Feature 3: Sort helpers ────────────────────────────────────

  function _sortPipelines(items) {
    if (!_sortCol) return items;
    const dir = _sortDir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let va = '', vb = '';
      if (_sortCol === 'name')   { va = a.def.name || ''; vb = b.def.name || ''; }
      if (_sortCol === 'status') { va = _buildStatus(a.lastBuild).label; vb = _buildStatus(b.lastBuild).label; }
      if (_sortCol === 'branch') { va = a.lastBuild?.sourceBranch || ''; vb = b.lastBuild?.sourceBranch || ''; }
      if (_sortCol === 'triggered') { va = a.lastBuild?.requestedFor?.displayName || ''; vb = b.lastBuild?.requestedFor?.displayName || ''; }
      if (_sortCol === 'lastrun') {
        va = a.lastBuild?.finishTime || a.lastBuild?.startTime || '';
        vb = b.lastBuild?.finishTime || b.lastBuild?.startTime || '';
      }
      return dir * va.toString().localeCompare(vb.toString());
    });
  }

  function _thSortClass(col) {
    if (_sortCol !== col) return 'sortable';
    return `sortable sort-${_sortDir}`;
  }

  function _renderPipelineList(container) {
    const content   = container.querySelector('#pipe-content');
    const search    = (container.querySelector('#pipe-search')?.value || '').toLowerCase();
    const statusF   = container.querySelector('#pipe-status-filter')?.value || '';

    let filtered = _pipelines;

    if (search) filtered = filtered.filter(p =>
      (p.def.name || '').toLowerCase().includes(search)
    );

    if (statusF) filtered = filtered.filter(p => {
      const s = _buildStatus(p.lastBuild);
      return s.key === statusF;
    });

    // Sort (Feature 3)
    filtered = _sortPipelines(filtered);

    if (filtered.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-rocket"></i>
          <p>No pipelines match your filter.</p>
        </div>`;
      return;
    }

    content.innerHTML = `
      <p class="text-sm text-muted mb-8">${filtered.length} pipeline${filtered.length !== 1 ? 's' : ''}</p>
      <div class="table-wrapper">
        <table tabindex="0" id="pipe-table">
          <thead>
            <tr>
              <th class="${_thSortClass('name')}" data-col="name">Pipeline</th>
              <th class="${_thSortClass('status')}" data-col="status">Last Status</th>
              <th class="${_thSortClass('branch')}" data-col="branch">Branch</th>
              <th class="${_thSortClass('triggered')}" data-col="triggered">Triggered By</th>
              <th>Duration</th>
              <th class="${_thSortClass('lastrun')}" data-col="lastrun">Last Run</th>
            </tr>
          </thead>
          <tbody>
            ${filtered.map(p => _pipelineRow(p)).join('')}
          </tbody>
        </table>
      </div>`;

    // Resizable columns (Feature 4 subset for pipelines)
    _initResizableColumns(content.querySelector('#pipe-table'));

    // Sort click (Feature 3)
    content.querySelectorAll('th[data-col]').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.col;
        if (_sortCol === col) {
          _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          _sortCol = col;
          _sortDir = 'asc';
        }
        _renderPipelineList(container);
      });
    });

    // Keyboard navigation (Feature 6)
    const table = content.querySelector('#pipe-table');
    if (table) {
      table.addEventListener('keydown', e => {
        const rows = Array.from(table.querySelectorAll('tbody tr'));
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
        } else if (e.key === 'Escape') {
          App.closeModal();
        }
      });
    }
  }

  // ─── Feature 4: Resizable columns (reuse same impl) ────────────

  function _initResizableColumns(table) {
    if (!table) return;
    table.querySelectorAll('thead th').forEach(th => {
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

  function _pipelineRow(entry) {
    const def   = entry.def;
    const build = entry.lastBuild;
    const status= _buildStatus(build);

    const branch    = build?.sourceBranch?.replace('refs/heads/', '') || '—';
    const triggeredBy = build?.requestedFor?.displayName || build?.requestedBy?.displayName || '—';
    const duration  = build ? _duration(build.startTime, build.finishTime) : '—';
    const lastRun   = build?.finishTime
      ? _relativeTime(new Date(build.finishTime))
      : (build?.startTime ? _relativeTime(new Date(build.startTime)) : '—');

    return `
      <tr>
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
