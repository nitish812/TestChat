/**
 * projects.js — Projects Listing
 *
 * Fetches all projects from active connections and renders them
 * in a card-based grid with search and org filter.
 * Supports pinning/favoriting projects via ado_pinned_projects localStorage key.
 */

const ProjectsModule = (() => {
  let _projects = [];   // { project, connectionId, connectionName, orgUrl, pat }
  let _loading  = false;
  const PINNED_KEY = 'ado_pinned_projects';

  // ─── Pin helpers ───────────────────────────────────────────────

  function _getPinned() {
    try { return JSON.parse(localStorage.getItem(PINNED_KEY) || '[]'); } catch { return []; }
  }

  function _setPinned(pins) {
    localStorage.setItem(PINNED_KEY, JSON.stringify(pins));
  }

  function _isPinned(connId, projectName) {
    return _getPinned().some(p => p.connId === connId && p.name === projectName);
  }

  function _togglePin(connId, projectName) {
    let pins = _getPinned();
    const idx = pins.findIndex(p => p.connId === connId && p.name === projectName);
    if (idx >= 0) { pins.splice(idx, 1); } else { pins.push({ connId, name: projectName }); }
    _setPinned(pins);
  }

  // ─── Public ────────────────────────────────────────────────────

  async function render(container, selectedOrgId) {
    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-folder-open"></i> Projects</div>
      <p class="page-subtitle">All projects across your connected Azure DevOps organizations.</p>
      <div class="filter-bar">
        <input type="text" id="proj-search" placeholder="🔍 Search projects…" />
        <select id="proj-org-filter">
          <option value="">All Organizations</option>
        </select>
        <button class="btn btn-secondary btn-sm" id="proj-refresh-btn">
          <i class="fa-solid fa-rotate-right"></i> Refresh
        </button>
      </div>
      <div id="proj-grid" class="grid-3"></div>
    `;

    _bindFilters(container);

    if (_projects.length === 0) {
      await _fetchAll(container, selectedOrgId);
    } else {
      _populateOrgFilter(container);
      _applyFilter(container, selectedOrgId);
    }
  }

  async function refresh(container, selectedOrgId) {
    _projects = [];
    await _fetchAll(container, selectedOrgId);
  }

  // ─── Private ───────────────────────────────────────────────────

  async function _fetchAll(container, selectedOrgId) {
    if (_loading) return;
    _loading = true;

    const grid = container.querySelector('#proj-grid');
    if (!grid) { _loading = false; return; }

    grid.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Loading projects…</p></div>';

    const connections = ConnectionsModule.getActive();
    if (connections.length === 0) {
      grid.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-plug-circle-xmark"></i>
          <p>No active connections found.</p>
          <a href="#/connections" class="btn btn-primary mt-8"><i class="fa-solid fa-plug"></i> Add Connection</a>
        </div>`;
      _loading = false;
      return;
    }

    _projects = [];

    for (const conn of connections) {
      const result = await AzureApi.getProjects(conn.orgUrl, conn.pat);
      if (result.error) {
        App.showToast(`${conn.name}: ${result.message}`, 'error');
        continue;
      }
      const items = result.value || [];
      items.forEach(p => _projects.push({
        project: p,
        connectionId: conn.id,
        connectionName: conn.name,
        orgUrl: conn.orgUrl,
        pat: conn.pat,
      }));
    }

    _loading = false;
    _populateOrgFilter(container);
    _applyFilter(container, selectedOrgId);
  }

  function _populateOrgFilter(container) {
    const sel = container.querySelector('#proj-org-filter');
    if (!sel) return;
    const seen = new Set();
    sel.innerHTML = '<option value="">All Organizations</option>';
    _projects.forEach(p => {
      if (!seen.has(p.connectionId)) {
        seen.add(p.connectionId);
        const opt = document.createElement('option');
        opt.value = p.connectionId;
        opt.textContent = p.connectionName;
        sel.appendChild(opt);
      }
    });
  }

  function _applyFilter(container, selectedOrgId) {
    const grid   = container.querySelector('#proj-grid');
    const search = (container.querySelector('#proj-search')?.value || '').toLowerCase();
    const orgId  = container.querySelector('#proj-org-filter')?.value || selectedOrgId || '';

    if (!grid) return;

    let filtered = _projects;
    if (orgId)   filtered = filtered.filter(p => p.connectionId === orgId);
    if (search)  filtered = filtered.filter(p =>
      p.project.name.toLowerCase().includes(search) ||
      (p.project.description || '').toLowerCase().includes(search)
    );

    if (filtered.length === 0) {
      grid.innerHTML = `
        <div class="empty-state" style="grid-column:1/-1">
          <i class="fa-solid fa-folder-open"></i>
          <p>No projects match your filter.</p>
        </div>`;
      return;
    }

    const pinned   = filtered.filter(p => _isPinned(p.connectionId, p.project.name));
    const unpinned = filtered.filter(p => !_isPinned(p.connectionId, p.project.name));

    let html = '';
    if (pinned.length > 0) {
      html += `<div class="section-title pinned-section-title" style="grid-column:1/-1"><i class="fa-solid fa-star" style="color:var(--color-warning)"></i> Pinned Projects</div>`;
      html += pinned.map(p => _cardHtml(p)).join('');
      if (unpinned.length > 0) {
        html += `<div class="section-title" style="grid-column:1/-1">All Projects</div>`;
      }
    }
    html += unpinned.map(p => _cardHtml(p)).join('');

    grid.innerHTML = html;

    // Bind pin buttons
    grid.querySelectorAll('.proj-pin-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        _togglePin(btn.dataset.connId, btn.dataset.project);
        _applyFilter(container, selectedOrgId);
      });
    });

    // Bind quick-action buttons
    grid.querySelectorAll('.proj-wi-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        App.navigate('workitems', {
          connId:  btn.dataset.connId,
          project: btn.dataset.project,
        });
      });
    });

    grid.querySelectorAll('.proj-pipe-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        App.navigate('pipelines', {
          connId:  btn.dataset.connId,
          project: btn.dataset.project,
        });
      });
    });
  }

  function _cardHtml(entry) {
    const p = entry.project;
    const pinned = _isPinned(entry.connectionId, p.name);
    const visibility = p.visibility === 'public'
      ? '<span class="badge badge-success">Public</span>'
      : '<span class="badge badge-secondary">Private</span>';

    const lastUpdated = p.lastUpdateTime
      ? _relativeTime(new Date(p.lastUpdateTime))
      : '—';

    const desc = p.description
      ? _esc(p.description)
      : '<span class="text-muted">No description</span>';

    return `
      <div class="project-card">
        <div class="project-card-header">
          <div>
            <div class="project-name">${_esc(p.name)}</div>
            <div class="project-org"><i class="fa-brands fa-windows" style="font-size:.7rem"></i> ${_esc(entry.connectionName)}</div>
          </div>
          <div style="display:flex;gap:6px;align-items:center">
            ${visibility}
            <button class="btn-icon proj-pin-btn" data-conn-id="${entry.connectionId}" data-project="${_esc(p.name)}"
              title="${pinned ? 'Unpin project' : 'Pin project'}" style="color:${pinned ? 'var(--color-warning)' : 'var(--text-muted)'}">
              <i class="fa-${pinned ? 'solid' : 'regular'} fa-star"></i>
            </button>
          </div>
        </div>
        <div class="project-desc">${desc}</div>
        <div class="project-meta">
          <span><i class="fa-regular fa-clock"></i> ${lastUpdated}</span>
          <span><i class="fa-solid fa-circle" style="color:${_stateColor(p.state)};font-size:.5rem;vertical-align:middle"></i> ${p.state || 'wellFormed'}</span>
        </div>
        <div class="project-actions">
          <button class="btn btn-secondary btn-sm proj-wi-btn"
            data-conn-id="${entry.connectionId}" data-project="${_esc(p.name)}">
            <i class="fa-solid fa-list-check"></i> Work Items
          </button>
          <button class="btn btn-secondary btn-sm proj-pipe-btn"
            data-conn-id="${entry.connectionId}" data-project="${_esc(p.name)}">
            <i class="fa-solid fa-rocket"></i> Pipelines
          </button>
        </div>
      </div>`;
  }

  function _bindFilters(container) {
    // Debounce search
    let timer;
    container.addEventListener('input', e => {
      if (e.target.id !== 'proj-search') return;
      clearTimeout(timer);
      timer = setTimeout(() => _applyFilter(container), 250);
    });

    container.addEventListener('change', e => {
      if (e.target.id === 'proj-org-filter') _applyFilter(container);
    });

    container.addEventListener('click', e => {
      if (e.target.closest('#proj-refresh-btn')) refresh(container);
    });
  }

  function _stateColor(state) {
    if (!state || state === 'wellFormed') return 'var(--color-success)';
    if (state === 'createPending')        return 'var(--color-warning)';
    return 'var(--color-danger)';
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
