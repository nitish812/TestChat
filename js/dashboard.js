/**
 * dashboard.js — Summary Dashboard
 *
 * Shows aggregate stats across all active connections:
 * total orgs, total projects, recent builds, health indicators.
 * Uses CSS-only charts (progress bars) — no external chart library.
 */

const DashboardModule = (() => {
  let _dateFrom = null;
  let _dateTo   = null;

  // ─── Public ────────────────────────────────────────────────────

  async function render(container) {
    container.innerHTML = `
      <div class="page-title"><i class="fa-solid fa-house"></i> Dashboard</div>
      <p class="page-subtitle">Centralized overview of all your Azure DevOps organizations.</p>
      <div class="filter-bar mb-8" style="flex-wrap:wrap">
        <label class="text-sm" style="align-self:center">From:</label>
        <input type="date" id="dash-date-from" style="max-width:160px" value="${_dateFrom || ''}" />
        <label class="text-sm" style="align-self:center">To:</label>
        <input type="date" id="dash-date-to" style="max-width:160px" value="${_dateTo || ''}" />
        <button class="btn btn-secondary btn-sm" id="dash-date-apply"><i class="fa-solid fa-filter"></i> Apply</button>
        <button class="btn btn-secondary btn-sm" id="dash-date-clear">Clear</button>
        <button class="btn btn-secondary btn-sm" id="dash-export-pdf" style="margin-left:auto"><i class="fa-solid fa-file-pdf"></i> Export PDF</button>
      </div>
      <div id="dash-content">
        <div class="loading-state"><div class="spinner spinner-lg"></div><p>Gathering data…</p></div>
      </div>
    `;

    container.querySelector('#dash-date-apply')?.addEventListener('click', () => {
      _dateFrom = container.querySelector('#dash-date-from')?.value || null;
      _dateTo   = container.querySelector('#dash-date-to')?.value || null;
      _loadData(container);
    });
    container.querySelector('#dash-date-clear')?.addEventListener('click', () => {
      _dateFrom = null; _dateTo = null;
      container.querySelector('#dash-date-from').value = '';
      container.querySelector('#dash-date-to').value = '';
      _loadData(container);
    });
    container.querySelector('#dash-export-pdf')?.addEventListener('click', () => window.print());

    await _loadData(container);
  }

  async function refresh(container) {
    await render(container);
  }

  // ─── Private ───────────────────────────────────────────────────

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

    // Fetch all org data in parallel
    const orgResults = await Promise.all(connections.map(conn => _fetchOrgData(conn)));

    // Aggregate totals
    let totalProjects = 0;
    let totalBuilds   = 0;
    let buildSucceeded= 0;
    let buildFailed   = 0;
    let buildInProg   = 0;
    const recentActivity = [];

    orgResults.forEach(org => {
      // Apply date range filter to builds if set
      let builds = org.builds;
      if (_dateFrom) builds = builds.filter(b => (b.finishTime || b.startTime || '') >= _dateFrom);
      if (_dateTo)   builds = builds.filter(b => (b.finishTime || b.startTime || '') <= (_dateTo + 'T23:59:59Z'));
      totalProjects += org.projectCount;
      totalBuilds   += builds.length;
      builds.forEach(b => {
        const r = (b.result || '').toLowerCase();
        const s = (b.status || '').toLowerCase();
        if (r === 'succeeded') buildSucceeded++;
        else if (r === 'failed' || r === 'partiallysucceeded') buildFailed++;
        else if (s === 'inprogress') buildInProg++;
      });
      // Add to activity feed
      builds.slice(0, 3).forEach(b => {
        recentActivity.push({
          time:  b.finishTime || b.startTime,
          org:   org.name,
          text:  `Build <strong>${_esc(b.definition?.name || '—')}</strong> ${_esc(b.result || b.status || '?')} in project <em>${_esc(b.project?.name || '—')}</em>`,
          icon:  _buildIcon(b),
        });
      });
    });

    // Sort activity feed by time (most recent first)
    recentActivity.sort((a, b) => (b.time || '') > (a.time || '') ? 1 : -1);

    // ── Render ──────────────────────────────────────────────────

    const buildsPct = (n) => totalBuilds > 0 ? Math.round(n / totalBuilds * 100) : 0;

    content.innerHTML = `
      <!-- Stats Row -->
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

      <!-- Build health bar -->
      ${totalBuilds > 0 ? `
      <div class="card mb-16">
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
      </div>
      ` : ''}

      <!-- Two-column: org health + activity -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;flex-wrap:wrap" class="dash-cols">
        <!-- Org Health -->
        <div class="card">
          <div class="card-header"><span class="card-title"><i class="fa-solid fa-heart-pulse"></i> Organization Health</span></div>
          <div style="display:flex;flex-direction:column;gap:10px">
            ${orgResults.map(org => _orgHealthHtml(org)).join('')}
          </div>
        </div>

        <!-- Recent Activity -->
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
      </div>

      <!-- Per-org detail cards -->
      <div class="section-title">Organizations</div>
      <div class="grid-3">
        ${orgResults.map(org => _orgDetailCard(org)).join('')}
      </div>
    `;

    // Responsive: stack columns on narrow screens
    const dashCols = content.querySelector('.dash-cols');
    if (dashCols) {
      const obs = new ResizeObserver(entries => {
        for (const entry of entries) {
          dashCols.style.gridTemplateColumns = entry.contentRect.width < 600 ? '1fr' : '1fr 1fr';
        }
      });
      obs.observe(dashCols.parentElement);
    }
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

  /** Fetch recent builds from first few projects (avoid excessive API calls) */
  async function _fetchRecentBuildsAllProjects(conn) {
    const projResult = await AzureApi.getProjects(conn.orgUrl, conn.pat);
    if (projResult.error) return [];

    const projects = (projResult.value || []).slice(0, 5); // limit to 5 projects
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

  function _orgHealthHtml(org) {
    const health = _orgHealth(org);
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
          <span><i class="fa-solid fa-rocket"></i> ${org.builds.length} recent runs</span>
        </div>
        ${org.error ? `<div class="text-sm" style="color:var(--color-danger)">${_esc(org.error)}</div>` : ''}
      </div>`;
  }

  function _orgDetailCard(org) {
    const succeeded = org.builds.filter(b => (b.result || '').toLowerCase() === 'succeeded').length;
    const failed    = org.builds.filter(b => (b.result || '').toLowerCase() === 'failed').length;
    const total     = org.builds.length;
    const health    = _orgHealth(org);

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
