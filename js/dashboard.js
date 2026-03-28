/**
 * dashboard.js — Summary Dashboard
 *
 * Shows aggregate stats across all active connections:
 * total orgs, total projects, recent builds, health indicators.
 * Uses CSS-only charts (progress bars) — no external chart library.
 *
 * Features:
 *  - Draggable widgets with HTML5 DnD (order saved to ado_widget_order)
 *  - Date range filter (7/30/90 days + custom)
 *  - Export dashboard as PDF (window.print)
 */

const DashboardModule = (() => {
  const WIDGET_ORDER_KEY = 'ado_widget_order';
  let _dateRangeDays = null;   // null = all time, number = days
  let _dateFrom = null;
  let _dateTo   = null;

  // ─── Public ────────────────────────────────────────────────────

  async function render(container) {
    container.innerHTML = `
      <div class="page-title no-print">
        <i class="fa-solid fa-house"></i> Dashboard
        <button class="btn btn-secondary btn-sm" id="dash-print-btn" style="margin-left:auto" title="Export as PDF">
          <i class="fa-solid fa-print"></i> Export PDF
        </button>
      </div>
      <p class="page-subtitle no-print">Centralized overview of all your Azure DevOps organizations.</p>

      <!-- Date range filter -->
      <div class="filter-bar mb-16 no-print" id="dash-date-bar">
        <label style="font-weight:600">Range:</label>
        <button class="btn btn-secondary btn-sm dash-range-btn ${_dateRangeDays === null ? 'active' : ''}" data-days="">All Time</button>
        <button class="btn btn-secondary btn-sm dash-range-btn ${_dateRangeDays === 7  ? 'active' : ''}" data-days="7">Last 7 days</button>
        <button class="btn btn-secondary btn-sm dash-range-btn ${_dateRangeDays === 30 ? 'active' : ''}" data-days="30">Last 30 days</button>
        <button class="btn btn-secondary btn-sm dash-range-btn ${_dateRangeDays === 90 ? 'active' : ''}" data-days="90">Last 90 days</button>
        <span style="font-weight:600">Custom:</span>
        <input type="date" id="dash-from" value="${_dateFrom || ''}" style="width:auto" title="From date" />
        <input type="date" id="dash-to"   value="${_dateTo   || ''}" style="width:auto" title="To date" />
        <button class="btn btn-primary btn-sm" id="dash-custom-btn"><i class="fa-solid fa-filter"></i> Apply</button>
      </div>

      <div id="dash-content">
        <div class="loading-state"><div class="spinner spinner-lg"></div><p>Gathering data…</p></div>
      </div>
    `;

    _bindDateBar(container);
    container.querySelector('#dash-print-btn')?.addEventListener('click', () => window.print());
    await _loadData(container);
  }

  async function refresh(container) {
    await render(container);
  }

  // ─── Private ───────────────────────────────────────────────────

  function _bindDateBar(container) {
    container.querySelectorAll('.dash-range-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const days = btn.dataset.days === '' ? null : parseInt(btn.dataset.days, 10);
        _dateRangeDays = days;
        _dateFrom = null;
        _dateTo   = null;
        container.querySelector('#dash-from').value = '';
        container.querySelector('#dash-to').value   = '';
        container.querySelectorAll('.dash-range-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        _loadData(container);
      });
    });

    container.querySelector('#dash-custom-btn')?.addEventListener('click', () => {
      _dateRangeDays = null;
      _dateFrom = container.querySelector('#dash-from')?.value || null;
      _dateTo   = container.querySelector('#dash-to')?.value   || null;
      container.querySelectorAll('.dash-range-btn').forEach(b => b.classList.remove('active'));
      _loadData(container);
    });
  }

  function _buildDateFilter() {
    if (_dateRangeDays !== null) {
      const from = new Date();
      from.setDate(from.getDate() - _dateRangeDays);
      return { minTime: from.toISOString() };
    }
    if (_dateFrom || _dateTo) {
      const filter = {};
      if (_dateFrom) filter.minTime = new Date(_dateFrom).toISOString();
      if (_dateTo)   filter.maxTime = new Date(_dateTo + 'T23:59:59').toISOString();
      return filter;
    }
    return {};
  }

  async function _loadData(container) {
    const connections = ConnectionsModule.getActive();
    const content     = container.querySelector('#dash-content');

    if (connections.length === 0) {
      content.innerHTML = `
        <div class="empty-state">
          <i class="fa-solid fa-plug-circle-xmark"></i>
          <p>No active connections. Add a connection to see your dashboard.</p>
          <a href="#/connections" class="btn btn-primary mt-8">
            <i class="fa-solid fa-plug"></i> Add Connection
          </a>
        </div>`;
      return;
    }

    content.innerHTML = '<div class="loading-state"><div class="spinner spinner-lg"></div><p>Gathering data…</p></div>';

    // Fetch all org data in parallel
    const orgResults = await Promise.all(connections.map(conn => _fetchOrgData(conn)));

    const dateFilter = _buildDateFilter();

    // Aggregate totals
    let totalProjects = 0;
    let totalBuilds   = 0;
    let buildSucceeded= 0;
    let buildFailed   = 0;
    let buildInProg   = 0;
    const recentActivity = [];

    orgResults.forEach(org => {
      totalProjects += org.projectCount;
      // Apply date filter to builds
      const builds = _filterBuilds(org.builds, dateFilter);
      totalBuilds   += builds.length;
      builds.forEach(b => {
        const r = (b.result || '').toLowerCase();
        const s = (b.status || '').toLowerCase();
        if (r === 'succeeded') buildSucceeded++;
        else if (r === 'failed' || r === 'partiallysucceeded') buildFailed++;
        else if (s === 'inprogress') buildInProg++;
      });
      builds.slice(0, 3).forEach(b => {
        recentActivity.push({
          time:  b.finishTime || b.startTime,
          org:   org.name,
          text:  `Build <strong>${_esc(b.definition?.name || '—')}</strong> ${_esc(b.result || b.status || '?')} in project <em>${_esc(b.project?.name || '—')}</em>`,
          icon:  _buildIcon(b),
        });
      });
    });

    recentActivity.sort((a, b) => (b.time || '') > (a.time || '') ? 1 : -1);

    const buildsPct = (n) => totalBuilds > 0 ? Math.round(n / totalBuilds * 100) : 0;

    // Build widget HTML map
    const widgets = {
      stats: `
        <div class="dash-widget" id="dw-stats" draggable="true" data-widget="stats">
          <div class="drag-handle no-print" title="Drag to reorder"><i class="fa-solid fa-grip-vertical"></i></div>
          <div class="stats-grid">
            <div class="stat-card">
              <div class="stat-value" style="color:var(--color-primary)">${connections.length}</div>
              <div class="stat-label"><i class="fa-solid fa-plug"></i> Connected Orgs</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" style="color:var(--color-info)">${totalProjects}</div>
              <div class="stat-label"><i class="fa-solid fa-folder-open"></i> Total Projects</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" style="color:var(--color-success)">${buildSucceeded}</div>
              <div class="stat-label"><i class="fa-solid fa-circle-check"></i> Builds Succeeded</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" style="color:var(--color-danger)">${buildFailed}</div>
              <div class="stat-label"><i class="fa-solid fa-circle-xmark"></i> Builds Failed</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" style="color:var(--color-warning)">${buildInProg}</div>
              <div class="stat-label"><i class="fa-solid fa-spinner fa-spin"></i> In Progress</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" style="color:var(--text-secondary)">${totalBuilds}</div>
              <div class="stat-label"><i class="fa-solid fa-rocket"></i> Total Build Runs</div>
            </div>
          </div>
        </div>`,

      health: totalBuilds > 0 ? `
        <div class="dash-widget card mb-16" id="dw-health" draggable="true" data-widget="health">
          <div class="drag-handle no-print" title="Drag to reorder"><i class="fa-solid fa-grip-vertical"></i></div>
          <div class="card-header"><span class="card-title"><i class="fa-solid fa-chart-bar"></i> Build Health</span></div>
          <div style="margin-bottom:6px;font-size:.8rem;display:flex;gap:16px;flex-wrap:wrap">
            <span style="color:var(--color-success)">🟢 Succeeded: ${buildSucceeded} (${buildsPct(buildSucceeded)}%)</span>
            <span style="color:var(--color-danger)">🔴 Failed: ${buildFailed} (${buildsPct(buildFailed)}%)</span>
            <span style="color:var(--color-warning)">🟡 In Progress: ${buildInProg}</span>
          </div>
          <div style="display:flex;height:14px;border-radius:20px;overflow:hidden;gap:2px">
            ${buildSucceeded > 0 ? `<div style="flex:${buildSucceeded};background:var(--color-success)" title="Succeeded"></div>` : ''}
            ${buildFailed   > 0 ? `<div style="flex:${buildFailed};background:var(--color-danger)"  title="Failed"></div>` : ''}
            ${buildInProg   > 0 ? `<div style="flex:${buildInProg};background:var(--color-warning)" title="In Progress"></div>` : ''}
          </div>
        </div>` : '',

      orgActivity: `
        <div class="dash-widget" id="dw-org-activity" draggable="true" data-widget="orgActivity" style="display:grid;grid-template-columns:1fr 1fr;gap:16px" class="dash-cols">
          <div class="drag-handle no-print" title="Drag to reorder" style="grid-column:1/-1"><i class="fa-solid fa-grip-vertical"></i></div>
          <div class="card">
            <div class="card-header"><span class="card-title"><i class="fa-solid fa-heart-pulse"></i> Organization Health</span></div>
            <div style="display:flex;flex-direction:column;gap:10px">
              ${orgResults.map(org => _orgHealthHtml(org, dateFilter)).join('')}
            </div>
          </div>
          <div class="card">
            <div class="card-header"><span class="card-title"><i class="fa-solid fa-clock-rotate-left"></i> Recent Activity</span></div>
            ${recentActivity.length === 0 ? `
              <div class="empty-state" style="padding:20px">
                <i class="fa-regular fa-calendar-xmark"></i>
                <p>No recent activity found.</p>
              </div>` : `
            <div class="activity-feed">
              ${recentActivity.slice(0, 10).map(a => `
                <div class="activity-item">
                  <div class="activity-icon">${a.icon}</div>
                  <div class="activity-text">
                    ${a.text}
                    <div class="text-muted text-sm">${_esc(a.org)}</div>
                  </div>
                  <div class="activity-time">${a.time ? _relativeTime(new Date(a.time)) : ''}</div>
                </div>`).join('')}
            </div>`}
          </div>
        </div>`,

      orgs: `
        <div class="dash-widget" id="dw-orgs" draggable="true" data-widget="orgs">
          <div class="drag-handle no-print" title="Drag to reorder"><i class="fa-solid fa-grip-vertical"></i></div>
          <div class="section-title">Organizations</div>
          <div class="grid-3">
            ${orgResults.map(org => _orgDetailCard(org, dateFilter)).join('')}
          </div>
        </div>`,
    };

    // Apply saved widget order
    const savedOrder = _getWidgetOrder();
    const widgetKeys = savedOrder.length > 0
      ? [...savedOrder.filter(k => widgets[k]), ...Object.keys(widgets).filter(k => !savedOrder.includes(k))]
      : Object.keys(widgets);

    content.innerHTML = `<div id="dash-widget-container">${widgetKeys.map(k => widgets[k]).join('')}</div>`;

    // Init drag-and-drop
    _initWidgetDnD(content);
  }

  function _filterBuilds(builds, dateFilter) {
    if (!dateFilter.minTime && !dateFilter.maxTime) return builds;
    return builds.filter(b => {
      const t = b.finishTime || b.startTime;
      if (!t) return true;
      const time = new Date(t).getTime();
      if (dateFilter.minTime && time < new Date(dateFilter.minTime).getTime()) return false;
      if (dateFilter.maxTime && time > new Date(dateFilter.maxTime).getTime()) return false;
      return true;
    });
  }

  // ─── Drag and drop ─────────────────────────────────────────────

  function _getWidgetOrder() {
    try { return JSON.parse(localStorage.getItem(WIDGET_ORDER_KEY) || '[]'); } catch { return []; }
  }

  function _saveWidgetOrder(order) {
    localStorage.setItem(WIDGET_ORDER_KEY, JSON.stringify(order));
  }

  function _initWidgetDnD(content) {
    const container = content.querySelector('#dash-widget-container');
    if (!container) return;
    let dragSrc = null;

    container.querySelectorAll('.dash-widget').forEach(widget => {
      widget.addEventListener('dragstart', e => {
        dragSrc = widget;
        e.dataTransfer.effectAllowed = 'move';
        widget.style.opacity = '0.5';
      });
      widget.addEventListener('dragend', () => {
        widget.style.opacity = '';
        // Save order
        const order = [...container.querySelectorAll('.dash-widget')].map(w => w.dataset.widget);
        _saveWidgetOrder(order);
      });
      widget.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        widget.style.outline = '2px dashed var(--color-primary)';
      });
      widget.addEventListener('dragleave', () => {
        widget.style.outline = '';
      });
      widget.addEventListener('drop', e => {
        e.preventDefault();
        widget.style.outline = '';
        if (dragSrc && dragSrc !== widget) {
          const allWidgets = [...container.querySelectorAll('.dash-widget')];
          const srcIdx = allWidgets.indexOf(dragSrc);
          const dstIdx = allWidgets.indexOf(widget);
          if (srcIdx < dstIdx) {
            container.insertBefore(dragSrc, widget.nextSibling);
          } else {
            container.insertBefore(dragSrc, widget);
          }
        }
      });
    });
  }

  async function _fetchOrgData(conn) {
    const [projResult, buildsResult] = await Promise.all([
      AzureApi.getProjects(conn.orgUrl, conn.pat),
      _fetchRecentBuildsAllProjects(conn),
    ]);

    const projects = projResult.error ? [] : (projResult.value || []);

    return {
      conn,
      name:         conn.name,
      orgUrl:       conn.orgUrl,
      projectCount: projects.length,
      projects,
      builds:       buildsResult,
      error:        projResult.error ? projResult.message : null,
    };
  }

  async function _fetchRecentBuildsAllProjects(conn) {
    const projResult = await AzureApi.getProjects(conn.orgUrl, conn.pat);
    if (projResult.error) return [];

    const projects = (projResult.value || []).slice(0, 5);
    const buildLists = await Promise.all(
      projects.map(p => AzureApi.getBuildRuns(conn.orgUrl, p.name, conn.pat))
    );

    const builds = [];
    buildLists.forEach((res, i) => {
      if (!res.error) {
        (res.value || []).slice(0, 5).forEach(b => {
          b.project = projects[i];
          builds.push(b);
        });
      }
    });

    return builds;
  }

  function _orgHealthHtml(org, dateFilter) {
    const builds = _filterBuilds(org.builds, dateFilter);
    const health = _orgHealth({ ...org, builds });
    return `
      <div class="org-health-card">
        <div style="display:flex;align-items:center;justify-content:space-between">
          <div class="org-health-title">
            <span class="health-dot ${health.cls}"></span>${_esc(org.name)}
          </div>
          <span class="badge ${health.badgeCls}">${health.label}</span>
        </div>
        <div class="org-health-stats">
          <span><i class="fa-solid fa-folder-open"></i> ${org.projectCount} projects</span>
          <span><i class="fa-solid fa-rocket"></i> ${builds.length} recent runs</span>
        </div>
        ${org.error ? `<div class="text-sm" style="color:var(--color-danger)">${_esc(org.error)}</div>` : ''}
      </div>`;
  }

  function _orgDetailCard(org, dateFilter) {
    const builds = _filterBuilds(org.builds, dateFilter);
    const succeeded = builds.filter(b => (b.result || '').toLowerCase() === 'succeeded').length;
    const failed    = builds.filter(b => (b.result || '').toLowerCase() === 'failed').length;
    const total     = builds.length;
    const health    = _orgHealth({ ...org, builds });

    return `
      <div class="card">
        <div class="card-header">
          <span class="card-title">
            <span class="health-dot ${health.cls}"></span>${_esc(org.name)}
          </span>
          <span class="badge ${health.badgeCls}">${health.label}</span>
        </div>
        <div class="card-subtitle truncate">${_esc(org.orgUrl)}</div>
        ${org.error ? `<div class="text-sm mt-8" style="color:var(--color-danger)"><i class="fa-solid fa-triangle-exclamation"></i> ${_esc(org.error)}</div>` : `
        <div style="margin-top:10px;font-size:.82rem;display:flex;flex-direction:column;gap:6px">
          <div style="display:flex;justify-content:space-between">
            <span><i class="fa-solid fa-folder-open"></i> Projects</span>
            <strong>${org.projectCount}</strong>
          </div>
          <div style="display:flex;justify-content:space-between">
            <span><i class="fa-solid fa-circle-check" style="color:var(--color-success)"></i> Succeeded</span>
            <strong style="color:var(--color-success)">${succeeded}</strong>
          </div>
          <div style="display:flex;justify-content:space-between">
            <span><i class="fa-solid fa-circle-xmark" style="color:var(--color-danger)"></i> Failed</span>
            <strong style="color:var(--color-danger)">${failed}</strong>
          </div>
          ${total > 0 ? `
          <div class="progress-bar-wrap mt-4">
            <div class="progress-bar-fill fill-success" style="width:${Math.round(succeeded/total*100)}%"></div>
          </div>` : ''}
        </div>
        <div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap">
          <a href="#/projects" class="btn btn-secondary btn-sm"><i class="fa-solid fa-folder-open"></i> Projects</a>
          <a href="#/pipelines" class="btn btn-secondary btn-sm"><i class="fa-solid fa-rocket"></i> Pipelines</a>
        </div>`}
      </div>`;
  }

  function _orgHealth(org) {
    if (org.error) return { cls: 'health-error', badgeCls: 'badge-danger', label: 'Error' };
    const failed = org.builds.filter(b => (b.result || '').toLowerCase() === 'failed').length;
    const total  = org.builds.length;
    if (total === 0 || failed === 0)               return { cls: 'health-ok',      badgeCls: 'badge-success',   label: 'Healthy' };
    if (failed / total < 0.3)                      return { cls: 'health-warning', badgeCls: 'badge-warning',   label: 'Warning' };
    return { cls: 'health-error', badgeCls: 'badge-danger', label: 'Degraded' };
  }

  function _buildIcon(build) {
    const r = (build.result || '').toLowerCase();
    if (r === 'succeeded')  return '🟢';
    if (r === 'failed')     return '🔴';
    if (r === 'canceled')   return '⚪';
    return '🟡';
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
