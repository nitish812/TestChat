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

  /**
   * Feature 9/11 — Update a single field on a work item (state).
   * PATCH with json-patch+json content type.
   */
  async function updateWorkItemState(orgUrl, project, id, newState, pat) {
    return _fetch(
      _url(orgUrl, `${encodeURIComponent(project)}/_apis/wit/workitems/${id}`),
      pat,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json-patch+json' },
        body: JSON.stringify([{ op: 'add', path: '/fields/System.State', value: newState }]),
      }
    );
  }

  /**
   * Feature 11 — Update multiple fields on a work item.
   * fields: { 'System.State': 'Active', 'System.AssignedTo': 'Name' }
   */
  async function updateWorkItemFields(orgUrl, project, id, fields, pat) {
    const patch = Object.entries(fields).map(([k, v]) => ({ op: 'add', path: `/fields/${k}`, value: v }));
    return _fetch(
      _url(orgUrl, `${encodeURIComponent(project)}/_apis/wit/workitems/${id}`),
      pat,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json-patch+json' },
        body: JSON.stringify(patch),
      }
    );
  }

  /**
   * Feature 12 — Get comments for a work item.
   */
  async function getWorkItemComments(orgUrl, project, id, pat) {
    const base = orgUrl.replace(/\/$/, '');
    const url = `${base}/${encodeURIComponent(project)}/_apis/wit/workitems/${id}/comments?api-version=7.1-preview.3`;
    return _fetch(url, pat);
  }

  /**
   * Feature 12 — Add a comment to a work item.
   */
  async function addWorkItemComment(orgUrl, project, id, text, pat) {
    const base = orgUrl.replace(/\/$/, '');
    const url = `${base}/${encodeURIComponent(project)}/_apis/wit/workitems/${id}/comments?api-version=7.1-preview.3`;
    return _fetch(url, pat, {
      method: 'POST',
      body: JSON.stringify({ text }),
    });
  }

  /**
   * Feature 13 — Create a new work item.
   */
  async function createWorkItem(orgUrl, project, type, titleVal, description, assignedTo, priority, pat) {
    const patch = [
      { op: 'add', path: '/fields/System.Title', value: titleVal },
    ];
    if (description)  patch.push({ op: 'add', path: '/fields/System.Description', value: description });
    if (assignedTo)   patch.push({ op: 'add', path: '/fields/System.AssignedTo',   value: assignedTo });
    if (priority)     patch.push({ op: 'add', path: '/fields/Microsoft.VSTS.Common.Priority', value: parseInt(priority, 10) });

    const base = orgUrl.replace(/\/$/, '');
    const url = `${base}/${encodeURIComponent(project)}/_apis/wit/workitems/${encodeURIComponent('$' + type)}?api-version=${API_VERSION}`;
    return _fetch(url, pat, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json-patch+json' },
      body: JSON.stringify(patch),
    });
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
    updateWorkItemState,
    updateWorkItemFields,
    getWorkItemComments,
    addWorkItemComment,
    createWorkItem,
    getBuildDefinitions,
    getBuildRuns,
    getReleaseDefinitions,
    getReleases,
  };
})();
