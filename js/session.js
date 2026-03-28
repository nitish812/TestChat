/**
 * session.js — Session Manager
 *
 * Persists and restores the user's navigation state so that
 * reloading the page (or returning later) resumes where they
 * left off — route, route params, org-selector value, and
 * sidebar collapsed state are all preserved in localStorage.
 *
 * Sessions expire after 24 hours of inactivity.
 */

const SessionManager = (() => {
  const STORAGE_KEY  = 'ado_session';
  const MAX_AGE_MS   = 24 * 60 * 60 * 1000; // 24 h

  /**
   * Save the current session state.
   * @param {{ route: string, params: object, orgSelectorValue: string, sidebarCollapsed: boolean }} state
   */
  function save(state) {
    try {
      const session = {
        route:             state.route || 'dashboard',
        params:            state.params || {},
        orgSelectorValue:  state.orgSelectorValue || '',
        sidebarCollapsed:  !!state.sidebarCollapsed,
        savedAt:           new Date().toISOString(),
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(session));
    } catch { /* ignore quota errors */ }
  }

  /**
   * Load the previously saved session, or null if none / expired.
   * @returns {{ route: string, params: object, orgSelectorValue: string, sidebarCollapsed: boolean } | null}
   */
  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;

      const session = JSON.parse(raw);

      // Expire stale sessions
      if (session.savedAt) {
        const age = Date.now() - new Date(session.savedAt).getTime();
        if (age > MAX_AGE_MS) {
          clear();
          return null;
        }
      }

      return session;
    } catch {
      return null;
    }
  }

  /** Remove the stored session. */
  function clear() {
    localStorage.removeItem(STORAGE_KEY);
  }

  return { save, load, clear };
})();
