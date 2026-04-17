const ADMIN_AUTH_SESSION_KEY = 'busbar.admin.auth.v1';
const fmtTimestampNO = new Intl.DateTimeFormat('no-NO', { dateStyle: 'short', timeStyle: 'short' });
const fmtNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
const fmtPercentNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 0,
  maximumFractionDigits: 2
});

const adminViewState = {
  users: [],
  projectRows: [],
  lineRows: [],
  projectColumnSort: { key: '', type: 'text', direction: 'asc' },
  lineColumnSort: { key: '', type: 'text', direction: 'asc' }
};

const $ = id => document.getElementById(id);

function normalizeApiBaseUrl(value){
  const raw = String(value || '').trim();
  if (!raw) return '';
  try{
    const parsed = new URL(raw, window.location.origin);
    if (!['http:', 'https:'].includes(parsed.protocol)) return '';
    const pathname = parsed.pathname.replace(/\/+$/, '');
    const suffix = pathname === '/' ? '' : pathname;
    return `${parsed.origin}${suffix}`;
  }catch(_err){
    return '';
  }
}

function isLocalDevelopmentHost(){
  const host = String(window.location.hostname || '').toLowerCase();
  return host === 'localhost' || host === '127.0.0.1' || host === '[::1]';
}

function resolveApiBaseUrl(){
  let fromStorage = '';
  try{
    fromStorage = localStorage.getItem('busbar.api.base') || '';
  }catch(_err){}

  let fromQuery = '';
  try{
    fromQuery = new URLSearchParams(window.location.search).get('apiBase') || '';
  }catch(_err){}

  const fromMeta = document.querySelector('meta[name="busbar-api-base"]')?.getAttribute('content') || '';
  const fromGlobal = typeof window.BUSBAR_API_BASE === 'string' ? window.BUSBAR_API_BASE : '';
  const normalized = normalizeApiBaseUrl(fromQuery || fromMeta || fromGlobal || fromStorage);

  if (fromQuery && normalized){
    try{
      localStorage.setItem('busbar.api.base', normalized);
    }catch(_err){}
  }
  if (normalized) return normalized;

  if (isLocalDevelopmentHost()){
    return 'http://localhost:5500';
  }
  return '';
}

const API_BASE_URL = resolveApiBaseUrl();

function buildApiUrl(path){
  const suffix = String(path || '').startsWith('/') ? String(path) : `/${String(path || '')}`;
  return API_BASE_URL ? `${API_BASE_URL}${suffix}` : suffix;
}

function isGithubPagesWithoutApiBase(){
  const host = String(window.location.hostname || '').toLowerCase();
  return host.endsWith('github.io') && !API_BASE_URL;
}

function appendApiBaseHint(errorText, status){
  if (!isGithubPagesWithoutApiBase()) return errorText;
  if (status !== 404 && status !== 405) return errorText;
  return `${errorText}. GitHub Pages krever at busbar-api-base peker til Render-backenden.`;
}

function compareNoText(left, right){
  return String(left || '').localeCompare(String(right || ''), 'no', {
    sensitivity: 'base',
    numeric: true
  });
}

function renderTablesFromState(){
  renderProjectsTable(getSortedProjectRows());
  renderLinesTable(getSortedLineRows());
}

function compareValuesByType(aValue, bValue, type = 'text'){
  if (type === 'number'){
    const left = Number(aValue);
    const right = Number(bValue);
    const leftSafe = Number.isFinite(left) ? left : Number.NEGATIVE_INFINITY;
    const rightSafe = Number.isFinite(right) ? right : Number.NEGATIVE_INFINITY;
    return leftSafe - rightSafe;
  }
  if (type === 'date'){
    const left = new Date(aValue || 0).getTime();
    const right = new Date(bValue || 0).getTime();
    const leftSafe = Number.isFinite(left) ? left : 0;
    const rightSafe = Number.isFinite(right) ? right : 0;
    return leftSafe - rightSafe;
  }
  return compareNoText(aValue, bValue);
}

function getSortedProjectRows(){
  const rows = Array.isArray(adminViewState.projectRows) ? [...adminViewState.projectRows] : [];
  const colSort = adminViewState.projectColumnSort;
  if (colSort?.key){
    rows.sort((a, b)=>{
      const cmp = compareValuesByType(a?.[colSort.key], b?.[colSort.key], colSort.type);
      return colSort.direction === 'desc' ? -cmp : cmp;
    });
    return rows;
  }
  return rows;
}

function getSortedLineRows(){
  const rows = Array.isArray(adminViewState.lineRows) ? [...adminViewState.lineRows] : [];
  const colSort = adminViewState.lineColumnSort;
  if (colSort?.key){
    rows.sort((a, b)=>{
      const cmp = compareValuesByType(a?.[colSort.key], b?.[colSort.key], colSort.type);
      return colSort.direction === 'desc' ? -cmp : cmp;
    });
    return rows;
  }
  return rows;
}

function setColumnSort(table, key, type){
  const normalizedTable = table === 'lines' ? 'lines' : 'projects';
  const normalizedType = type || 'text';
  const current = normalizedTable === 'lines'
    ? adminViewState.lineColumnSort
    : adminViewState.projectColumnSort;
  const next = { ...current };
  if (next.key === key){
    next.direction = next.direction === 'asc' ? 'desc' : 'asc';
  } else {
    next.key = key;
    next.type = normalizedType;
    next.direction = normalizedType === 'date' ? 'desc' : 'asc';
  }
  if (normalizedTable === 'lines'){
    adminViewState.lineColumnSort = next;
  } else {
    adminViewState.projectColumnSort = next;
  }
  updateColumnSortUi();
  renderTablesFromState();
}

function updateColumnSortUi(){
  const buttons = Array.from(document.querySelectorAll('.admin-col-sort-btn'));
  buttons.forEach(button=>{
    const table = button.dataset.table === 'lines' ? 'lines' : 'projects';
    const key = button.dataset.key || '';
    const state = table === 'lines' ? adminViewState.lineColumnSort : adminViewState.projectColumnSort;
    const active = Boolean(state?.key) && state.key === key;
    const direction = active ? state.direction : '';
    button.classList.toggle('is-active', active);
    button.classList.toggle('is-asc', active && direction === 'asc');
    button.classList.toggle('is-desc', active && direction === 'desc');
    const indicator = button.querySelector('.sort-indicator');
    if (indicator){
      indicator.textContent = active ? (direction === 'desc' ? '▼' : '▲') : '↕';
    }
    const th = button.closest('th');
    if (th){
      th.setAttribute('aria-sort', active ? (direction === 'desc' ? 'descending' : 'ascending') : 'none');
    }
  });
}

function readStoredAdminAuth(){
  try{
    const raw = sessionStorage.getItem(ADMIN_AUTH_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed.authHeader !== 'string' || typeof parsed.username !== 'string'){
      return null;
    }
    return {
      authHeader: parsed.authHeader,
      username: parsed.username
    };
  }catch(_err){
    return null;
  }
}

function writeStoredAdminAuth(authHeader, username){
  try{
    sessionStorage.setItem(ADMIN_AUTH_SESSION_KEY, JSON.stringify({ authHeader, username }));
  }catch(_err){}
}

function clearStoredAdminAuth(){
  try{
    sessionStorage.removeItem(ADMIN_AUTH_SESSION_KEY);
  }catch(_err){}
}

function encodeBasicAuth(username, password){
  const token = btoa(unescape(encodeURIComponent(`${username}:${password}`)));
  return `Basic ${token}`;
}

function formatTimestamp(value){
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '-';
  return fmtTimestampNO.format(d);
}

function toFiniteNumber(value){
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function formatAmount(value){
  const n = toFiniteNumber(value);
  if (!Number.isFinite(n)) return '-';
  return `${fmtNO.format(n)} NOK`;
}

function formatPercent(value){
  const n = toFiniteNumber(value);
  if (!Number.isFinite(n)) return '-';
  return `${fmtPercentNO.format(n * 100)} %`;
}

function resolveLineTotalExFreight(line){
  const totals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  return toFiniteNumber(totals.totalExMontasje);
}

function resolveLineDgRate(line){
  const totals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const inputs = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
    ? line.inputs
    : {};
  const direct = toFiniteNumber(inputs.marginRate ?? totals.marginRate);
  if (Number.isFinite(direct)){
    return direct > 1 ? direct / 100 : direct;
  }
  const material = toFiniteNumber(totals.material);
  const subtotal = toFiniteNumber(totals.subtotal);
  if (Number.isFinite(material) && Number.isFinite(subtotal) && subtotal > 0){
    return 1 - material / subtotal;
  }
  return NaN;
}

function aggregateProjectMetrics(project){
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  let projectTotalExFreight = 0;
  let materialSum = 0;
  let subtotalSum = 0;
  let hasSubtotalMaterial = false;
  let dgRateSum = 0;
  let dgRateCount = 0;

  lines.forEach(line => {
    const lineTotalExFreight = resolveLineTotalExFreight(line);
    if (Number.isFinite(lineTotalExFreight)){
      projectTotalExFreight += lineTotalExFreight;
    }

    const totals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
      ? line.totals
      : {};
    const material = toFiniteNumber(totals.material);
    const subtotal = toFiniteNumber(totals.subtotal);
    if (Number.isFinite(material) && Number.isFinite(subtotal) && subtotal > 0){
      materialSum += material;
      subtotalSum += subtotal;
      hasSubtotalMaterial = true;
    }

    const lineDgRate = resolveLineDgRate(line);
    if (Number.isFinite(lineDgRate)){
      dgRateSum += lineDgRate;
      dgRateCount += 1;
    }
  });

  let projectDgRate = NaN;
  if (hasSubtotalMaterial && subtotalSum > 0){
    projectDgRate = 1 - materialSum / subtotalSum;
  } else if (dgRateCount > 0){
    projectDgRate = dgRateSum / dgRateCount;
  }

  return {
    projectTotalExFreight,
    projectDgRate
  };
}

function renderSummaryFromRows(projectRows, lineRows){
  const summaryEl = $('adminSummary');
  if (!summaryEl) return;
  const projectEntries = Array.isArray(projectRows) ? projectRows : [];
  const lineEntries = Array.isArray(lineRows) ? lineRows : [];
  const users = new Set();
  projectEntries.forEach(row=>users.add(String(row?.email || '').trim().toLowerCase()));
  lineEntries.forEach(row=>users.add(String(row?.email || '').trim().toLowerCase()));
  users.delete('');
  users.delete('-');
  const userCount = users.size;
  const projectCount = projectEntries.length;
  const lineCount = lineEntries.length;
  summaryEl.innerHTML = `
    <div class="admin-summary-card">
      <span class="admin-summary-label">Brukere</span>
      <strong class="admin-summary-value">${userCount}</strong>
    </div>
    <div class="admin-summary-card">
      <span class="admin-summary-label">Prosjekter</span>
      <strong class="admin-summary-value">${projectCount}</strong>
    </div>
    <div class="admin-summary-card">
      <span class="admin-summary-label">Linjer</span>
      <strong class="admin-summary-value">${lineCount}</strong>
    </div>
  `;
}

function buildProjectRows(users){
  const rows = [];
  (Array.isArray(users) ? users : []).forEach(user => {
    const email = String(user?.email || '-');
    const projects = Array.isArray(user?.projects) ? user.projects : [];
    projects.forEach(project => {
      const metrics = aggregateProjectMetrics(project);
      const lineCount = Array.isArray(project?.lines) ? project.lines.length : 0;
      const projectTotalRaw = toFiniteNumber(metrics.projectTotalExFreight);
      const projectDgRaw = toFiniteNumber(metrics.projectDgRate);
      rows.push({
        email,
        name: project?.name || '-',
        customer: project?.customer || '-',
        contactPerson: project?.contactPerson || '-',
        lineCount,
        totalExFreightRaw: projectTotalRaw,
        totalExFreight: formatAmount(projectTotalRaw),
        dgPercentRaw: projectDgRaw,
        dgPercent: formatPercent(projectDgRaw),
        createdAtRaw: project?.createdAt || null,
        updatedAtRaw: project?.updatedAt || null,
        createdAt: formatTimestamp(project?.createdAt),
        updatedAt: formatTimestamp(project?.updatedAt)
      });
    });
  });
  return rows;
}

function buildLineRows(users){
  const rows = [];
  (Array.isArray(users) ? users : []).forEach(user => {
    const email = String(user?.email || '-');
    const projects = Array.isArray(user?.projects) ? user.projects : [];
    projects.forEach(project => {
      const projectName = project?.name || '-';
      const lines = Array.isArray(project?.lines) ? project.lines : [];
      lines.forEach(line => {
        const inputs = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
          ? line.inputs
          : {};
        const ampereRaw = toFiniteNumber(inputs?.ampere);
        const meterRaw = toFiniteNumber(inputs?.meter);
        const totalRaw = toFiniteNumber(resolveLineTotalExFreight(line));
        const dgRaw = toFiniteNumber(resolveLineDgRate(line));
        rows.push({
          email,
          projectName,
          lineNumber: line?.lineNumber || '-',
          series: inputs?.series || '-',
          ampereRaw,
          ampere: Number.isFinite(ampereRaw) ? ampereRaw : '-',
          meterRaw,
          meter: Number.isFinite(meterRaw) ? meterRaw : '-',
          totalExFreightRaw: totalRaw,
          totalExFreight: formatAmount(totalRaw),
          dgPercentRaw: dgRaw,
          dgPercent: formatPercent(dgRaw),
          createdAtRaw: line?.createdAt || null,
          updatedAtRaw: line?.updatedAt || line?.createdAt || null,
          updatedAt: formatTimestamp(line?.updatedAt || line?.createdAt)
        });
      });
    });
  });
  return rows;
}

function rebuildRowsFromUsers(){
  adminViewState.projectRows = buildProjectRows(adminViewState.users);
  adminViewState.lineRows = buildLineRows(adminViewState.users);
}

function renderProjectsTable(rowsInput){
  const tbody = document.querySelector('#adminProjectsTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="9">Ingen prosjekter er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.name)}</td>
      <td>${escapeHtml(row.customer)}</td>
      <td>${escapeHtml(row.contactPerson)}</td>
      <td>${row.lineCount}</td>
      <td>${escapeHtml(row.totalExFreight)}</td>
      <td>${escapeHtml(row.dgPercent)}</td>
      <td>${escapeHtml(row.createdAt)}</td>
      <td>${escapeHtml(row.updatedAt)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderLinesTable(rowsInput){
  const tbody = document.querySelector('#adminLinesTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="9">Ingen linjer er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.projectName)}</td>
      <td>${escapeHtml(String(row.lineNumber))}</td>
      <td>${escapeHtml(String(row.series))}</td>
      <td>${escapeHtml(String(row.ampere))}</td>
      <td>${escapeHtml(String(row.meter))}</td>
      <td>${escapeHtml(row.totalExFreight)}</td>
      <td>${escapeHtml(row.dgPercent)}</td>
      <td>${escapeHtml(row.updatedAt)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function escapeHtml(value){
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setAuthError(message){
  const el = $('adminAuthError');
  if (!el) return;
  el.textContent = message || '';
}

async function fetchAdminOverview(authHeader){
  const res = await fetch(buildApiUrl('/api/admin/project-overview'), {
    method: 'GET',
    headers: {
      Authorization: authHeader
    },
    cache: 'no-store'
  });

  if (res.status === 401){
    const err = new Error('Ugyldig admin brukernavn/passord.');
    err.code = 'AUTH';
    throw err;
  }
  if (!res.ok){
    let message = `Kunne ikke hente adminoversikt (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        message += `: ${data.error.trim()}`;
      }
    }catch(_err){}
    throw new Error(appendApiBaseHint(message, res.status));
  }
  return res.json();
}

async function loadOverview(authHeader){
  const refreshBtn = $('adminRefreshBtn');
  if (refreshBtn) refreshBtn.disabled = true;
  try{
    const data = await fetchAdminOverview(authHeader);
    adminViewState.users = Array.isArray(data?.users) ? data.users : [];
    rebuildRowsFromUsers();
    renderSummaryFromRows(adminViewState.projectRows, adminViewState.lineRows);
    renderTablesFromState();
    const generatedEl = $('adminGeneratedAt');
    if (generatedEl){
      generatedEl.textContent = `Sist oppdatert: ${formatTimestamp(data?.generatedAt)}`;
    }
    const dataCard = $('adminDataCard');
    if (dataCard) dataCard.hidden = false;
    setAuthError('');
  }catch(err){
    const dataCard = $('adminDataCard');
    if (dataCard) dataCard.hidden = true;
    throw err;
  }finally{
    if (refreshBtn) refreshBtn.disabled = false;
  }
}

async function handleAdminLoginSubmit(){
  const username = String($('adminUsername')?.value || '').trim();
  const password = String($('adminPassword')?.value || '');
  if (!username || !password){
    setAuthError('Fyll inn brukernavn og passord.');
    return;
  }
  const authHeader = encodeBasicAuth(username, password);
  try{
    await loadOverview(authHeader);
    writeStoredAdminAuth(authHeader, username);
    setAuthError('');
  }catch(err){
    clearStoredAdminAuth();
    setAuthError(err?.message || 'Innlogging feilet.');
  }
}

function handleAdminLogout(){
  clearStoredAdminAuth();
  const dataCard = $('adminDataCard');
  if (dataCard) dataCard.hidden = true;
  setAuthError('Logget ut.');
}

function bindUi(){
  const loginForm = $('adminLoginForm');
  if (loginForm){
    loginForm.addEventListener('submit', evt => {
      evt.preventDefault();
      void handleAdminLoginSubmit();
    });
  }
  const loginBtn = $('adminLoginBtn');
  if (loginBtn){
    loginBtn.addEventListener('click', evt => {
      evt.preventDefault();
      void handleAdminLoginSubmit();
    });
  }
  const refreshBtn = $('adminRefreshBtn');
  if (refreshBtn){
    refreshBtn.addEventListener('click', async () => {
      const stored = readStoredAdminAuth();
      if (!stored?.authHeader){
        setAuthError('Logg inn som admin for a laste data.');
        return;
      }
      try{
        await loadOverview(stored.authHeader);
      }catch(err){
        setAuthError(err?.message || 'Oppdatering feilet.');
      }
    });
  }
  document.addEventListener('click', evt=>{
    const button = evt.target?.closest?.('.admin-col-sort-btn');
    if (!button) return;
    const key = String(button.dataset.key || '').trim();
    if (!key) return;
    const table = button.dataset.table === 'lines' ? 'lines' : 'projects';
    const type = String(button.dataset.type || 'text').trim().toLowerCase();
    setColumnSort(table, key, type);
  });
  const logoutBtn = $('adminLogoutBtn');
  if (logoutBtn){
    logoutBtn.addEventListener('click', handleAdminLogout);
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  updateColumnSortUi();
  bindUi();
  const stored = readStoredAdminAuth();
  if (!stored) return;
  const usernameInput = $('adminUsername');
  if (usernameInput) usernameInput.value = stored.username || '';
  try{
    await loadOverview(stored.authHeader);
  }catch(err){
    clearStoredAdminAuth();
    setAuthError(err?.message || 'Kunne ikke validere admin-sesjon.');
  }
});
