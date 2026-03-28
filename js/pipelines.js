/**
 * pipelines.js — Pipelines / Builds
 *
 * Fetches build definitions and recent runs from a selected project
 * and renders them in a table with color-coded status badges.
 */

const PipelinesModule = (() => {
  let _pipelines   = [];   // merged build definition + last run
  let _context     = null;
  let _sortCol     = null;
  let _sortDir     = 'asc';

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, params = {}) {
    _pipelines = [];

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
    }));

    container.querySelector('#pipe-filters').style.display = '';
    _renderPipelineList(container);
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

    if (filtered.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-rocket"></i>
          <p>No pipelines match your filter.</p>
        </div>`;
      return;
    }

    // Apply sort
    if (_sortCol) {
      filtered = filtered.slice().sort((a, b) => {
        let va = _pipeSortVal(a, _sortCol);
        let vb = _pipeSortVal(b, _sortCol);
        if (va < vb) return _sortDir === 'asc' ? -1 : 1;
        if (va > vb) return _sortDir === 'asc' ? 1 : -1;
        return 0;
      });
    }

    const _th = (key, label) => {
      const active = _sortCol === key;
      const icon = active ? (_sortDir === 'asc' ? ' ▲' : ' ▼') : '';
      return `<th class="sortable-col${active ? ' sort-active' : ''}" data-sort="${key}" style="cursor:pointer;user-select:none">${label}${icon}</th>`;
    };

    content.innerHTML = `
      <p class="text-sm text-muted mb-8">${filtered.length} pipeline${filtered.length !== 1 ? 's' : ''}</p>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              ${_th('name','Pipeline')}${_th('status','Last Status')}${_th('branch','Branch')}
              ${_th('triggeredBy','Triggered By')}${_th('duration','Duration')}${_th('lastRun','Last Run')}
            </tr>
          </thead>
          <tbody>
            ${filtered.map(p => _pipelineRow(p)).join('')}
          </tbody>
        </table>
      </div>`;

    // Sort header click
    content.querySelectorAll('.sortable-col').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.sort;
        if (_sortCol === col) { _sortDir = _sortDir === 'asc' ? 'desc' : 'asc'; }
        else { _sortCol = col; _sortDir = 'asc'; }
        _renderPipelineList(container);
      });
    });

    // Keyboard nav on rows
    content.querySelectorAll('tbody tr').forEach(row => {
      row.setAttribute('tabindex', '0');
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

  function _pipeSortVal(entry, col) {
    const def = entry.def;
    const build = entry.lastBuild;
    switch (col) {
      case 'name':       return (def.name || '').toLowerCase();
      case 'status':     return _buildStatus(build).label.toLowerCase();
      case 'branch':     return (build?.sourceBranch || '').replace('refs/heads/', '').toLowerCase();
      case 'triggeredBy':return (build?.requestedFor?.displayName || build?.requestedBy?.displayName || '').toLowerCase();
      case 'duration':   return build ? (new Date(build.finishTime || Date.now()) - new Date(build.startTime || Date.now())) : 0;
      case 'lastRun':    return build?.finishTime || build?.startTime || '';
      default:           return '';
    }
  }

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
