/**
 * app.js — Main application entry point
 *
 * Handles:
 *  - Hash-based SPA routing  (#/dashboard, #/connections, etc.)
 *  - Global state
 *  - Sidebar/header interactions
 *  - Theme toggle (light / dark)
 *  - Toast notifications
 *  - Modal helper
 *  - Refresh button
 *  - Org selector (header dropdown)
 */

const App = (() => {
  // ─── State ─────────────────────────────────────────────────────
  let _currentRoute = null;
  let _currentParams= {};

  // ─── Init ──────────────────────────────────────────────────────

  function init() {
    _restoreTheme();
    _bindGlobalEvents();
    refreshOrgSelector();

    // Initial navigation
    const hash = window.location.hash || '#/dashboard';
    _handleHashChange(hash);

    // Listen for hash changes
    window.addEventListener('hashchange', () => _handleHashChange(window.location.hash));

    // Feature 1: Show PAT expiry toasts for active connections
    _checkPatExpiry();
  }

  /** Show toast notifications for PATs that are expired or expiring soon. */
  function _checkPatExpiry() {
    ConnectionsModule.getActive().forEach(c => {
      const info = ConnectionsModule.getExpiryInfo(c.patExpiry);
      if (!info) return;
      if (info.expired) {
        showToast(`"${c.name}": PAT has expired. Please update your Personal Access Token.`, 'error');
      } else if (info.daysLeft <= 7) {
        showToast(`"${c.name}": PAT expires in ${info.daysLeft} day${info.daysLeft !== 1 ? 's' : ''}. Consider renewing it.`, 'warning');
      }
    });
  }

  // ─── Routing ───────────────────────────────────────────────────

  function _handleHashChange(hash) {
    const [routePart] = hash.replace('#/', '').split('?');
    const route = routePart || 'dashboard';
    _navigate(route, {});
  }

  function navigate(route, params = {}) {
    // Set params before changing the hash so the hashchange handler can read them.
    _currentParams = params;
    window.location.hash = `#/${route}`;
  }

  function _navigate(route, params) {
    _currentRoute = route;
    if (params && Object.keys(params).length) _currentParams = params;

    // Update sidebar active state
    document.querySelectorAll('.nav-link').forEach(a => {
      a.classList.toggle('active', a.dataset.route === route);
    });

    const container = document.getElementById('page-container');
    if (!container) return;

    const p = _currentParams || {};

    switch (route) {
      case 'dashboard':
        DashboardModule.render(container);
        break;
      case 'connections':
        ConnectionsModule.render(container);
        break;
      case 'projects':
        ProjectsModule.render(container, _getSelectedOrgId());
        break;
      case 'workitems':
        WorkItemsModule.render(container, p);
        break;
      case 'pipelines':
        PipelinesModule.render(container, p);
        break;
      default:
        DashboardModule.render(container);
    }

    // Reset per-navigation params after use
    _currentParams = {};
  }

  function _getSelectedOrgId() {
    return document.getElementById('org-selector')?.value || '';
  }

  // ─── Org Selector ──────────────────────────────────────────────

  function refreshOrgSelector() {
    const sel = document.getElementById('org-selector');
    if (!sel) return;
    const selected = sel.value;
    sel.innerHTML = '<option value="">— All Organizations —</option>';
    ConnectionsModule.getActive().forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.id;
      opt.textContent = c.name;
      sel.appendChild(opt);
    });
    if (selected) sel.value = selected;
  }

  // ─── Global events ─────────────────────────────────────────────

  function _bindGlobalEvents() {
    // Sidebar toggle
    document.getElementById('sidebar-toggle')?.addEventListener('click', () => {
      document.body.classList.toggle('sidebar-collapsed');
      // On mobile, toggle open class instead
      if (window.innerWidth <= 768) {
        document.body.classList.remove('sidebar-collapsed');
        document.body.classList.toggle('sidebar-open');
      }
    });

    // Close mobile sidebar when clicking main
    document.getElementById('main-content')?.addEventListener('click', () => {
      if (document.body.classList.contains('sidebar-open')) {
        document.body.classList.remove('sidebar-open');
      }
    });

    // Theme toggle
    document.getElementById('theme-toggle')?.addEventListener('click', () => {
      const isDark = document.body.classList.toggle('theme-dark');
      document.body.classList.toggle('theme-light', !isDark);
      localStorage.setItem('ado_theme', isDark ? 'dark' : 'light');
      const icon = document.querySelector('#theme-toggle i');
      if (icon) icon.className = isDark ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
    });

    // Refresh button
    document.getElementById('refresh-btn')?.addEventListener('click', async () => {
      const btn = document.getElementById('refresh-btn');
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-rotate-right fa-spin"></i> Refreshing…';

      const container = document.getElementById('page-container');
      switch (_currentRoute) {
        case 'dashboard':   await DashboardModule.refresh(container);   break;
        case 'projects':    await ProjectsModule.refresh(container, _getSelectedOrgId()); break;
        case 'workitems':   await WorkItemsModule.refresh(container);   break;
        case 'pipelines':   await PipelinesModule.refresh(container);   break;
        default: break;
      }

      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-rotate-right"></i> Refresh';
    });

    // Org selector change → re-render current page if relevant
    document.getElementById('org-selector')?.addEventListener('change', () => {
      const container = document.getElementById('page-container');
      if (_currentRoute === 'projects') ProjectsModule.render(container, _getSelectedOrgId());
    });

    // Modal close
    document.getElementById('modal-close')?.addEventListener('click', closeModal);
    document.getElementById('modal-overlay')?.addEventListener('click', e => {
      if (e.target === document.getElementById('modal-overlay')) closeModal();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') closeModal();
    });
  }

  function _restoreTheme() {
    const saved = localStorage.getItem('ado_theme');
    if (saved === 'dark') {
      document.body.classList.remove('theme-light');
      document.body.classList.add('theme-dark');
      const icon = document.querySelector('#theme-toggle i');
      if (icon) icon.className = 'fa-solid fa-sun';
    }
  }

  // ─── Toast notifications ───────────────────────────────────────

  function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const icons = {
      success: '<i class="fa-solid fa-circle-check" style="color:var(--color-success)"></i>',
      error:   '<i class="fa-solid fa-circle-xmark" style="color:var(--color-danger)"></i>',
      warning: '<i class="fa-solid fa-triangle-exclamation" style="color:var(--color-warning)"></i>',
      info:    '<i class="fa-solid fa-circle-info" style="color:var(--color-primary)"></i>',
    };

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      <span class="toast-icon">${icons[type] || icons.info}</span>
      <span class="toast-msg">${_esc(message)}</span>
      <button class="toast-close" aria-label="Dismiss"><i class="fa-solid fa-xmark"></i></button>
    `;

    toast.querySelector('.toast-close').addEventListener('click', () => toast.remove());
    container.appendChild(toast);

    // Auto-dismiss after 5 s
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 5000);
  }

  // ─── Modal ─────────────────────────────────────────────────────

  function openModal(title, bodyHtml) {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML = bodyHtml;
    document.getElementById('modal-overlay').classList.remove('hidden');
  }

  function updateModalBody(html) {
    const body = document.getElementById('modal-body');
    if (body) body.innerHTML = html;
  }

  function closeModal() {
    document.getElementById('modal-overlay').classList.add('hidden');
  }

  // ─── Utility ───────────────────────────────────────────────────

  function _esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ─── Bootstrap ─────────────────────────────────────────────────
  document.addEventListener('DOMContentLoaded', init);

  return { navigate, refreshOrgSelector, showToast, openModal, updateModalBody, closeModal };
})();
