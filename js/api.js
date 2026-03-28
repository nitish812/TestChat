/**
 * api.js — Azure DevOps REST API wrapper
 *
 * All fetch calls are made directly from the browser using Basic Auth
 * (Base64 encoded ":PAT" per Azure DevOps docs).
 *
 * Error types returned from every function:
 *   { error: true, status, message, cors? }
 */

const AzureApi = (() => {
  const API_VERSION = '7.0';

  /**
   * Build the Base64 Basic-Auth header from a PAT.
   * Azure DevOps expects `:PAT` (empty username).
   */
  function _authHeader(pat) {
    return 'Basic ' + btoa(':' + pat);
  }

  /**
   * Core fetch wrapper — returns parsed JSON or an error object.
   */
  async function _fetch(url, pat, options = {}) {
    try {
      const res = await fetch(url, {
        ...options,
        headers: {
          'Authorization': _authHeader(pat),
          'Content-Type': 'application/json',
          ...(options.headers || {}),
        },
      });

      if (res.status === 401) {
        return { error: true, status: 401, message: 'Unauthorized — check your PAT.' };
      }
      if (res.status === 403) {
        return { error: true, status: 403, message: 'Forbidden — PAT may lack required scopes.' };
      }
      if (res.status === 404) {
        return { error: true, status: 404, message: 'Not found — check your organization URL.' };
      }
      if (res.status === 429) {
        return { error: true, status: 429, message: 'Rate limited — please wait before retrying.' };
      }
      if (!res.ok) {
        return { error: true, status: res.status, message: `API error: ${res.statusText}` };
      }

      const data = await res.json();
      return data;
    } catch (err) {
      // TypeError is thrown for network errors AND CORS blocks
      if (err instanceof TypeError) {
        return {
          error: true,
          status: 0,
          cors: true,
          message:
            'Network error — the request was blocked (possible CORS restriction). ' +
            'Azure DevOps REST API may not allow browser requests from this origin. ' +
            'Consider using a CORS proxy or a local backend.',
        };
      }
      return { error: true, status: 0, message: err.message };
    }
  }

  /** Build a standard Azure DevOps API URL */
  function _url(orgUrl, path, params = {}) {
    const base = orgUrl.replace(/\/$/, '');
    const qs = new URLSearchParams({ 'api-version': API_VERSION, ...params }).toString();
    return `${base}/${path}?${qs}`;
  }

  // ─────────────────────────────────────────────────────────────
  // Public API
  // ─────────────────────────────────────────────────────────────

  /**
   * Validate a connection by fetching the project list.
   * Returns { valid: true, projectCount } or { valid: false, message }
   */
  async function validateConnection(orgUrl, pat) {
    const result = await _fetch(_url(orgUrl, '_apis/projects', { $top: 1 }), pat);
    if (result.error) return { valid: false, message: result.message };
    return { valid: true, projectCount: result.count };
  }

  /** List all projects for an org. */
  async function getProjects(orgUrl, pat) {
    return _fetch(_url(orgUrl, '_apis/projects', { $top: 200, stateFilter: 'all' }), pat);
  }

  /** List work items for a project via WIQL. */
  async function getWorkItems(orgUrl, project, pat, filters = {}) {
    const conditions = ['[System.TeamProject] = @project'];

    if (filters.type)       conditions.push(`[System.WorkItemType] = '${_esc(filters.type)}'`);
    if (filters.state)      conditions.push(`[System.State] = '${_esc(filters.state)}'`);
    if (filters.assignedTo) conditions.push(`[System.AssignedTo] = '${_esc(filters.assignedTo)}'`);
    if (filters.priority)   conditions.push(`[Microsoft.VSTS.Common.Priority] = ${parseInt(filters.priority, 10)}`);

    const wiql = `SELECT [System.Id] FROM WorkItems WHERE ${conditions.join(' AND ')} ORDER BY [System.ChangedDate] DESC`;

    const wiqlResult = await _fetch(
      _url(orgUrl, `${encodeURIComponent(project)}/_apis/wit/wiql`, { $top: 200 }),
      pat,
      { method: 'POST', body: JSON.stringify({ query: wiql }) }
    );

    if (wiqlResult.error) return wiqlResult;
    if (!wiqlResult.workItems || wiqlResult.workItems.length === 0) return { value: [] };

    // Fetch details in batches of 50
    const ids = wiqlResult.workItems.map(w => w.id);
    const fields = [
      'System.Id', 'System.Title', 'System.WorkItemType', 'System.State',
      'System.AssignedTo', 'Microsoft.VSTS.Common.Priority',
      'System.CreatedDate', 'System.ChangedDate', 'System.Description',
      'System.AreaPath', 'System.IterationPath', 'Microsoft.VSTS.Common.Severity',
    ].join(',');

    const batches = _chunk(ids, 50);
    const results = [];

    for (const batch of batches) {
      const data = await _fetch(
        _url(orgUrl, '_apis/wit/workitems', { ids: batch.join(','), fields }),
        pat
      );
      if (data.error) return data;
      results.push(...(data.value || []));
    }

    return { value: results };
  }

  /** Get a single work item by ID (with all fields). */
  async function getWorkItemDetail(orgUrl, id, pat) {
    return _fetch(_url(orgUrl, `_apis/wit/workitems/${id}`, { $expand: 'all' }), pat);
  }

  /** List build definitions (pipelines) for a project. */
  async function getBuildDefinitions(orgUrl, project, pat) {
    return _fetch(_url(orgUrl, `${encodeURIComponent(project)}/_apis/build/definitions`, { $top: 100 }), pat);
  }

  /** List recent build runs for a project. */
  async function getBuildRuns(orgUrl, project, pat, definitionId) {
    const params = { $top: 50 };
    if (definitionId) params.definitions = definitionId;
    return _fetch(_url(orgUrl, `${encodeURIComponent(project)}/_apis/build/builds`, params), pat);
  }

  /** List release pipelines for a project. */
  async function getReleaseDefinitions(orgUrl, project, pat) {
    // Release API uses vsrm subdomain
    const releaseOrgUrl = orgUrl.replace('dev.azure.com', 'vsrm.dev.azure.com');
    return _fetch(_url(releaseOrgUrl, `${encodeURIComponent(project)}/_apis/release/definitions`, { $top: 100 }), pat);
  }

  /** List recent releases for a project. */
  async function getReleases(orgUrl, project, pat) {
    const releaseOrgUrl = orgUrl.replace('dev.azure.com', 'vsrm.dev.azure.com');
    return _fetch(_url(releaseOrgUrl, `${encodeURIComponent(project)}/_apis/release/releases`, { $top: 20 }), pat);
  }

  /**
   * Update a work item's state via PATCH (JSON Patch document).
   * @param {string} orgUrl
   * @param {number|string} id  Work item ID
   * @param {string} state      New state value
   * @param {string} pat
   */
  async function updateWorkItemState(orgUrl, id, state, pat) {
    const body = JSON.stringify([
      { op: 'add', path: '/fields/System.State', value: state },
    ]);
    return _fetch(
      _url(orgUrl, `_apis/wit/workitems/${id}`),
      pat,
      { method: 'PATCH', headers: { 'Content-Type': 'application/json-patch+json' }, body }
    );
  }

  /**
   * Update arbitrary work item fields via PATCH (JSON Patch document).
   * @param {string} orgUrl
   * @param {number|string} id
   * @param {Object} fieldsMap  e.g. { 'System.State': 'Done', 'System.AssignedTo': 'user@org' }
   * @param {string} pat
   */
  async function updateWorkItemFields(orgUrl, id, fieldsMap, pat) {
    const ops = Object.entries(fieldsMap).map(([path, value]) => ({
      op: 'add', path: `/fields/${path}`, value,
    }));
    return _fetch(
      _url(orgUrl, `_apis/wit/workitems/${id}`),
      pat,
      { method: 'PATCH', headers: { 'Content-Type': 'application/json-patch+json' }, body: JSON.stringify(ops) }
    );
  }

  /**
   * Create a new work item via POST (JSON Patch document).
   * @param {string} orgUrl
   * @param {string} project
   * @param {string} type       e.g. 'Task', 'Bug', 'User Story'
   * @param {Object} fieldsMap  e.g. { 'System.Title': 'My task' }
   * @param {string} pat
   */
  async function createWorkItem(orgUrl, project, type, fieldsMap, pat) {
    const ops = Object.entries(fieldsMap).map(([path, value]) => ({
      op: 'add', path: `/fields/${path}`, value,
    }));
    return _fetch(
      _url(orgUrl, `${encodeURIComponent(project)}/_apis/wit/workitems/${encodeURIComponent('$' + type)}`),
      pat,
      { method: 'POST', headers: { 'Content-Type': 'application/json-patch+json' }, body: JSON.stringify(ops) }
    );
  }

  /**
   * Get comments for a work item.
   */
  async function getWorkItemComments(orgUrl, id, pat) {
    return _fetch(_url(orgUrl, `_apis/wit/workitems/${id}/comments`, { 'api-version': '7.0-preview.3' }), pat);
  }

  /**
   * Add a comment to a work item.
   */
  async function addWorkItemComment(orgUrl, id, text, pat) {
    return _fetch(
      _url(orgUrl, `_apis/wit/workitems/${id}/comments`, { 'api-version': '7.0-preview.3' }),
      pat,
      { method: 'POST', body: JSON.stringify({ text }) }
    );
  }

  /** Get a single work item with relations expanded. */
  async function getWorkItemWithRelations(orgUrl, id, pat) {
    return _fetch(_url(orgUrl, `_apis/wit/workitems/${id}`, { $expand: 'relations' }), pat);
  }

  // ─────────────────────────────────────────────────────────────
  // Helpers
  // ─────────────────────────────────────────────────────────────

  /** Escape single quotes in WIQL string values. */
  function _esc(s) { return String(s).replace(/'/g, "''"); }

  /** Split an array into fixed-size chunks. */
  function _chunk(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  return {
    validateConnection,
    getProjects,
    getWorkItems,
    getWorkItemDetail,
    getWorkItemWithRelations,
    getBuildDefinitions,
    getBuildRuns,
    getReleaseDefinitions,
    getReleases,
    updateWorkItemState,
    updateWorkItemFields,
    createWorkItem,
    getWorkItemComments,
    addWorkItemComment,
  };
})();
