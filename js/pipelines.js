/**
 * pipelines.js — Pipelines / Builds
 *
 * Fetches build definitions and recent runs from a selected project
 * and renders them in a table with color-coded status badges.
 *
 * Improvements:
 *  #2  Collapsible connection groups
 *  #3  Column sorting
 *  #4  Resizable columns
 *  #7  Breadcrumb navigation
 */

const PipelinesModule = (() => {
  let _pipelines   = [];   // merged build definition + last run
  let _context     = null;

  // #3 Column sorting
  let _sortCol = '';
  let _sortDir = 'asc';

  // #2 Collapsible groups — persists collapsed state within session
  const _collapsedGroups = new Set();

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

  function _sortPipelines(items) {
    if (!_sortCol) return items;
    const dir = _sortDir === 'asc' ? 1 : -1;
    return [...items].sort((a, b) => {
      let av, bv;
      switch (_sortCol) {
        case 'name':        av = (a.def.name || '').toLowerCase();                            bv = (b.def.name || '').toLowerCase(); break;
        case 'status':      av = _buildStatus(a.lastBuild).key;                               bv = _buildStatus(b.lastBuild).key; break;
        case 'branch':      av = (a.lastBuild?.sourceBranch || '').replace('refs/heads/', ''); bv = (b.lastBuild?.sourceBranch || '').replace('refs/heads/', ''); break;
        case 'triggeredBy': av = (a.lastBuild?.requestedFor?.displayName || '').toLowerCase(); bv = (b.lastBuild?.requestedFor?.displayName || '').toLowerCase(); break;
        case 'duration':    av = a.lastBuild?.startTime ? new Date(a.lastBuild.startTime).getTime() : 0; bv = b.lastBuild?.startTime ? new Date(b.lastBuild.startTime).getTime() : 0; break;
        case 'finished':    av = a.lastBuild?.finishTime || '';                               bv = b.lastBuild?.finishTime || ''; break;
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

    // #3 Apply sort
    filtered = _sortPipelines(filtered);

    if (filtered.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-rocket"></i>
          <p>No pipelines match your filter.</p>
        </div>`;
      return;
    }

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

    // #2 Group by connection name (single group in single-project mode)
    const groupName = _context ? _context.conn.name : 'Pipelines';
    const isCollapsed = _collapsedGroups.has(groupName);
    const rowsHtml = isCollapsed ? '' : filtered.map(p => _pipelineRow(p)).join('');
    const chevron  = isCollapsed ? 'fa-chevron-right' : 'fa-chevron-down';

    content.innerHTML = `
      ${breadcrumb}
      <p class="text-sm text-muted mb-8">${filtered.length} pipeline${filtered.length !== 1 ? 's' : ''}</p>
      <div class="table-wrapper">
        <table id="pipe-table" style="table-layout:fixed;min-width:600px">
          <thead>
            <tr>
              <th data-col="name" style="cursor:pointer">Pipeline ${_sortIcon('name')}</th>
              <th data-col="status" style="cursor:pointer">Last Status ${_sortIcon('status')}</th>
              <th data-col="branch" style="cursor:pointer">Branch ${_sortIcon('branch')}</th>
              <th data-col="triggeredBy" style="cursor:pointer">Triggered By ${_sortIcon('triggeredBy')}</th>
              <th data-col="duration" style="cursor:pointer">Duration ${_sortIcon('duration')}</th>
              <th data-col="finished" style="cursor:pointer">Last Run ${_sortIcon('finished')}</th>
            </tr>
          </thead>
          <tbody>
            <tr class="group-header" data-group="${_esc(groupName)}">
              <td colspan="6">
                <button class="btn-group-toggle">
                  <i class="fa-solid ${chevron}"></i>
                  ${_esc(groupName)} (${filtered.length} item${filtered.length !== 1 ? 's' : ''})
                </button>
              </td>
            </tr>
            ${rowsHtml}
          </tbody>
        </table>
      </div>`;

    // #4 Resizable columns
    _makeColumnsResizable(content.querySelector('#pipe-table'));

    // #3 Sort on <th> click
    content.querySelectorAll('thead th[data-col]').forEach(th => {
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

    // #2 Group toggle
    content.querySelectorAll('.btn-group-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const group = btn.closest('tr').dataset.group;
        if (_collapsedGroups.has(group)) {
          _collapsedGroups.delete(group);
        } else {
          _collapsedGroups.add(group);
        }
        _renderPipelineList(container);
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

  // #4 Resizable columns helper
  function _makeColumnsResizable(tableEl) {
    if (!tableEl) return;
    const ths = tableEl.querySelectorAll('thead th');
    ths.forEach(th => {
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
