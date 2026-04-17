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

const ADMIN_PROJECT_SORT_STORAGE_KEY = 'busbar.admin.project.sort.v1';
const ADMIN_LINE_SORT_STORAGE_KEY = 'busbar.admin.line.sort.v1';
const ADMIN_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
const adminViewState = {
  users: [],
  projectSort: 'date_newest',
  lineSort: 'date_newest'
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

function loadSortMode(storageKey, validModes, fallback){
  if (typeof localStorage === 'undefined') return fallback;
  try{
    const raw = localStorage.getItem(storageKey);
    return validModes.includes(raw) ? raw : fallback;
  }catch(_err){
    return fallback;
  }
}

function saveSortMode(storageKey, mode){
  if (typeof localStorage === 'undefined') return;
  try{
    localStorage.setItem(storageKey, mode);
  }catch(_err){}
}

function getSortableCreatedTimestamp(item){
  const created = new Date(item?.createdAt || item?.updatedAt || 0).getTime();
  return Number.isFinite(created) ? created : 0;
}

function compareNoText(left, right){
  return String(left || '').localeCompare(String(right || ''), 'no', {
    sensitivity: 'base',
    numeric: true
  });
}

function compareAdminProjectRowsForSort(a, b, mode = adminViewState.projectSort){
  if (mode === 'alpha_asc'){
    return compareNoText(a?.name, b?.name);
  }
  if (mode === 'alpha_desc'){
    return compareNoText(b?.name, a?.name);
  }
  const aTime = getSortableCreatedTimestamp(a);
  const bTime = getSortableCreatedTimestamp(b);
  if (mode === 'date_oldest'){
    return aTime - bTime;
  }
  return bTime - aTime;
}

function compareAdminLineRowsForSort(a, b, mode = adminViewState.lineSort){
  if (mode === 'alpha_asc'){
    return compareNoText(a?.lineNumber, b?.lineNumber);
  }
  if (mode === 'alpha_desc'){
    return compareNoText(b?.lineNumber, a?.lineNumber);
  }
  const aTime = getSortableCreatedTimestamp(a);
  const bTime = getSortableCreatedTimestamp(b);
  if (mode === 'date_oldest'){
    return aTime - bTime;
  }
  return bTime - aTime;
}

function updateSortControlValues(){
  const projectSelect = $('adminProjectSortSelect');
  const lineSelect = $('adminLineSortSelect');
  if (projectSelect && ADMIN_SORT_OPTIONS.includes(adminViewState.projectSort)){
    projectSelect.value = adminViewState.projectSort;
  }
  if (lineSelect && ADMIN_SORT_OPTIONS.includes(adminViewState.lineSort)){
    lineSelect.value = adminViewState.lineSort;
  }
}

function applyAdminSortModesFromStorage(){
  adminViewState.projectSort = loadSortMode(
    ADMIN_PROJECT_SORT_STORAGE_KEY,
    ADMIN_SORT_OPTIONS,
    'date_newest'
  );
  adminViewState.lineSort = loadSortMode(
    ADMIN_LINE_SORT_STORAGE_KEY,
    ADMIN_SORT_OPTIONS,
    'date_newest'
  );
  updateSortControlValues();
}

function renderTablesFromState(){
  renderProjectsTable(adminViewState.users);
  renderLinesTable(adminViewState.users);
}

function setAdminProjectSortMode(mode, options = {}){
  if (!ADMIN_SORT_OPTIONS.includes(mode)) return;
  adminViewState.projectSort = mode;
  if (options.persist !== false){
    saveSortMode(ADMIN_PROJECT_SORT_STORAGE_KEY, mode);
  }
  updateSortControlValues();
  if (options.render !== false){
    renderTablesFromState();
  }
}

function setAdminLineSortMode(mode, options = {}){
  if (!ADMIN_SORT_OPTIONS.includes(mode)) return;
  adminViewState.lineSort = mode;
  if (options.persist !== false){
    saveSortMode(ADMIN_LINE_SORT_STORAGE_KEY, mode);
  }
  updateSortControlValues();
  if (options.render !== false){
    renderTablesFromState();
  }
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

function renderSummary(totals){
  const summaryEl = $('adminSummary');
  if (!summaryEl) return;
  const userCount = Number(totals?.userCount || 0);
  const projectCount = Number(totals?.projectCount || 0);
  const lineCount = Number(totals?.lineCount || 0);
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

function renderProjectsTable(users){
  const tbody = document.querySelector('#adminProjectsTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = [];

  (Array.isArray(users) ? users : []).forEach(user => {
    const email = String(user?.email || '-');
    const projects = Array.isArray(user?.projects) ? user.projects : [];
    projects.forEach(project => {
      const metrics = aggregateProjectMetrics(project);
      rows.push({
        email,
        name: project?.name || '-',
        customer: project?.customer || '-',
        contactPerson: project?.contactPerson || '-',
        lineCount: Array.isArray(project?.lines) ? project.lines.length : 0,
        totalExFreight: formatAmount(metrics.projectTotalExFreight),
        dgPercent: formatPercent(metrics.projectDgRate),
        createdAtRaw: project?.createdAt || null,
        updatedAtRaw: project?.updatedAt || null,
        createdAt: formatTimestamp(project?.createdAt),
        updatedAt: formatTimestamp(project?.updatedAt)
      });
    });
  });

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="9">Ingen prosjekter er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.sort((a, b)=>compareAdminProjectRowsForSort({
    name: a.name,
    createdAt: a.createdAtRaw,
    updatedAt: a.updatedAtRaw
  }, {
    name: b.name,
    createdAt: b.createdAtRaw,
    updatedAt: b.updatedAtRaw
  }, adminViewState.projectSort));

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

function renderLinesTable(users){
  const tbody = document.querySelector('#adminLinesTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
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
        rows.push({
          email,
          projectName,
          lineNumber: line?.lineNumber || '-',
          series: inputs?.series || '-',
          ampere: Number.isFinite(Number(inputs?.ampere)) ? Number(inputs.ampere) : '-',
          meter: Number.isFinite(Number(inputs?.meter)) ? Number(inputs.meter) : '-',
          totalExFreight: formatAmount(resolveLineTotalExFreight(line)),
          dgPercent: formatPercent(resolveLineDgRate(line)),
          createdAtRaw: line?.createdAt || null,
          updatedAtRaw: line?.updatedAt || line?.createdAt || null,
          updatedAt: formatTimestamp(line?.updatedAt || line?.createdAt)
        });
      });
    });
  });

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="9">Ingen linjer er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.sort((a, b)=>compareAdminLineRowsForSort({
    lineNumber: a.lineNumber,
    createdAt: a.createdAtRaw,
    updatedAt: a.updatedAtRaw
  }, {
    lineNumber: b.lineNumber,
    createdAt: b.createdAtRaw,
    updatedAt: b.updatedAtRaw
  }, adminViewState.lineSort));

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
    renderSummary(data?.totals || {});
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
  const projectSortSelect = $('adminProjectSortSelect');
  if (projectSortSelect){
    projectSortSelect.addEventListener('change', ()=>{
      setAdminProjectSortMode(projectSortSelect.value);
    });
  }
  const lineSortSelect = $('adminLineSortSelect');
  if (lineSortSelect){
    lineSortSelect.addEventListener('change', ()=>{
      setAdminLineSortMode(lineSortSelect.value);
    });
  }
  const logoutBtn = $('adminLogoutBtn');
  if (logoutBtn){
    logoutBtn.addEventListener('click', handleAdminLogout);
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  applyAdminSortModesFromStorage();
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
