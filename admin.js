const ADMIN_AUTH_SESSION_KEY = 'busbar.admin.auth.v1';
const DASHBOARD_SIDEBAR_COLLAPSED_KEY = 'busbar.dashboard.sidebar.collapsed.v1';
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
  userRows: [],
  projectRows: [],
  lineRows: [],
  customers: [],
  customerRows: [],
  customerContactRows: [],
  userColumnSort: { key: '', type: 'text', direction: 'asc' },
  projectColumnSort: { key: '', type: 'text', direction: 'asc' },
  lineColumnSort: { key: '', type: 'text', direction: 'asc' },
  customerColumnSort: { key: '', type: 'text', direction: 'asc' },
  customerContactColumnSort: { key: '', type: 'text', direction: 'asc' }
};

const $ = id => document.getElementById(id);

function readDashboardSidebarCollapsed(){
  try{
    return localStorage.getItem(DASHBOARD_SIDEBAR_COLLAPSED_KEY) === '1';
  }catch(_err){
    return false;
  }
}

function persistDashboardSidebarCollapsed(collapsed){
  try{
    localStorage.setItem(DASHBOARD_SIDEBAR_COLLAPSED_KEY, collapsed ? '1' : '0');
  }catch(_err){}
}

function setDashboardSidebarCollapsed(collapsed, options = {}){
  const shell = $('dashboardShell');
  const toggle = $('dashboardSidebarToggle');
  const next = Boolean(collapsed);
  if (shell) shell.classList.toggle('is-sidebar-collapsed', next);
  if (toggle){
    toggle.setAttribute('aria-expanded', next ? 'false' : 'true');
    toggle.setAttribute('aria-label', next ? 'Utvid meny' : 'Minimer meny');
    toggle.title = next ? 'Utvid meny' : 'Minimer meny';
  }
  if (!options.skipPersist){
    persistDashboardSidebarCollapsed(next);
  }
}

function initDashboardSidebar(){
  const shell = $('dashboardShell');
  const toggle = $('dashboardSidebarToggle');
  if (!shell || !toggle) return;
  setDashboardSidebarCollapsed(readDashboardSidebarCollapsed(), { skipPersist: true });
  toggle.addEventListener('click', evt=>{
    evt.preventDefault();
    setDashboardSidebarCollapsed(!shell.classList.contains('is-sidebar-collapsed'));
  });
}

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
  let fromQuery = '';
  try{
    fromQuery = new URLSearchParams(window.location.search).get('apiBase') || '';
  }catch(_err){}
  const normalizedQuery = normalizeApiBaseUrl(fromQuery);
  if (fromQuery && normalizedQuery){
    try{
      localStorage.setItem('busbar.api.base', normalizedQuery);
    }catch(_err){}
    return normalizedQuery;
  }

  if (isLocalDevelopmentHost()){
    return window.location.origin;
  }

  let fromStorage = '';
  try{
    fromStorage = localStorage.getItem('busbar.api.base') || '';
  }catch(_err){}

  const fromMeta = document.querySelector('meta[name="busbar-api-base"]')?.getAttribute('content') || '';
  const fromGlobal = typeof window.BUSBAR_API_BASE === 'string' ? window.BUSBAR_API_BASE : '';
  const normalized = normalizeApiBaseUrl(fromMeta || fromGlobal || fromStorage);
  if (normalized) return normalized;

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
  renderUsersTable(getSortedUserRows());
  renderProjectsTable(getSortedProjectRows());
  renderLinesTable(getSortedLineRows());
  renderCustomersTable(getSortedCustomerRows());
  renderCustomerContactsTable(getSortedCustomerContactRows());
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

function getSortedUserRows(){
  const rows = Array.isArray(adminViewState.userRows) ? [...adminViewState.userRows] : [];
  const colSort = adminViewState.userColumnSort;
  if (colSort?.key){
    rows.sort((a, b)=>{
      const cmp = compareValuesByType(a?.[colSort.key], b?.[colSort.key], colSort.type);
      return colSort.direction === 'desc' ? -cmp : cmp;
    });
  }
  return rows;
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

function getSortedCustomerRows(){
  const rows = Array.isArray(adminViewState.customerRows) ? [...adminViewState.customerRows] : [];
  const colSort = adminViewState.customerColumnSort;
  if (colSort?.key){
    rows.sort((a, b)=>{
      const cmp = compareValuesByType(a?.[colSort.key], b?.[colSort.key], colSort.type);
      return colSort.direction === 'desc' ? -cmp : cmp;
    });
  }
  return rows;
}

function getSortedCustomerContactRows(){
  const rows = Array.isArray(adminViewState.customerContactRows) ? [...adminViewState.customerContactRows] : [];
  const colSort = adminViewState.customerContactColumnSort;
  if (colSort?.key){
    rows.sort((a, b)=>{
      const cmp = compareValuesByType(a?.[colSort.key], b?.[colSort.key], colSort.type);
      return colSort.direction === 'desc' ? -cmp : cmp;
    });
  }
  return rows;
}

function setColumnSort(table, key, type){
  const normalizedTable = table === 'users'
    ? 'users'
    : table === 'lines'
    ? 'lines'
    : table === 'customers'
      ? 'customers'
      : table === 'customerContacts'
        ? 'customerContacts'
        : 'projects';
  const normalizedType = type || 'text';
  const current = normalizedTable === 'users'
    ? adminViewState.userColumnSort
    : normalizedTable === 'lines'
    ? adminViewState.lineColumnSort
    : normalizedTable === 'customers'
      ? adminViewState.customerColumnSort
      : normalizedTable === 'customerContacts'
        ? adminViewState.customerContactColumnSort
        : adminViewState.projectColumnSort;
  const next = { ...current };
  if (next.key === key){
    next.direction = next.direction === 'asc' ? 'desc' : 'asc';
  } else {
    next.key = key;
    next.type = normalizedType;
    next.direction = normalizedType === 'date' ? 'desc' : 'asc';
  }
  if (normalizedTable === 'users'){
    adminViewState.userColumnSort = next;
  } else if (normalizedTable === 'lines'){
    adminViewState.lineColumnSort = next;
  } else if (normalizedTable === 'customers'){
    adminViewState.customerColumnSort = next;
  } else if (normalizedTable === 'customerContacts'){
    adminViewState.customerContactColumnSort = next;
  } else {
    adminViewState.projectColumnSort = next;
  }
  updateColumnSortUi();
  renderTablesFromState();
}

function updateColumnSortUi(){
  const buttons = Array.from(document.querySelectorAll('.admin-col-sort-btn'));
  buttons.forEach(button=>{
    const table = button.dataset.table === 'users'
      ? 'users'
      : button.dataset.table === 'lines'
      ? 'lines'
      : button.dataset.table === 'customers'
        ? 'customers'
        : button.dataset.table === 'customerContacts'
          ? 'customerContacts'
          : 'projects';
    const key = button.dataset.key || '';
    const state = table === 'users'
      ? adminViewState.userColumnSort
      : table === 'lines'
      ? adminViewState.lineColumnSort
      : table === 'customers'
        ? adminViewState.customerColumnSort
        : table === 'customerContacts'
          ? adminViewState.customerContactColumnSort
          : adminViewState.projectColumnSort;
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
  const userCount = Array.isArray(adminViewState.users) ? adminViewState.users.length : 0;
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

function buildUserRows(users){
  return (Array.isArray(users) ? users : []).map(user => {
    const profile = user?.profile && typeof user.profile === 'object' ? user.profile : {};
    const registered = user?.registered === true;
    const microsoftLinked = user?.microsoftLinked === true;
    const hasPassword = user?.hasPassword === true;
    const isAdmin = user?.isAdmin === true;
    const loginParts = [
      registered ? 'Registrert' : 'Kun prosjekt',
      microsoftLinked ? 'Microsoft' : '',
      hasPassword ? 'Passord' : '',
      isAdmin ? 'Admin' : ''
    ].filter(Boolean);
    return {
      email: String(user?.email || '-'),
      name: profile.name || '-',
      phone: profile.phone || '-',
      company: profile.company || '-',
      position: profile.position || '-',
      projectCount: Number(user?.projectCount || 0),
      lineCount: Number(user?.lineCount || 0),
      loginStatus: loginParts.join(' / ') || '-',
      updatedAtRaw: user?.authUpdatedAt || user?.updatedAt || null,
      updatedAt: formatTimestamp(user?.authUpdatedAt || user?.updatedAt),
      rawProfile: {
        name: profile.name || '',
        phone: profile.phone || '',
        company: profile.company || '',
        position: profile.position || ''
      }
    };
  });
}

function buildProjectRows(users){
  const rows = [];
  (Array.isArray(users) ? users : []).forEach(user => {
    const email = String(user?.email || '-');
    const profile = user?.profile && typeof user.profile === 'object' ? user.profile : {};
    const projects = Array.isArray(user?.projects) ? user.projects : [];
    projects.forEach(project => {
      const metrics = aggregateProjectMetrics(project);
      const lineCount = Array.isArray(project?.lines) ? project.lines.length : 0;
      const projectTotalRaw = toFiniteNumber(metrics.projectTotalExFreight);
      const projectDgRaw = toFiniteNumber(metrics.projectDgRate);
      rows.push({
        email,
        userName: profile.name || '-',
        userPhone: profile.phone || '-',
        userCompany: profile.company || '-',
        userPosition: profile.position || '-',
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
    const profile = user?.profile && typeof user.profile === 'object' ? user.profile : {};
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
          userName: profile.name || '-',
          userPhone: profile.phone || '-',
          userCompany: profile.company || '-',
          userPosition: profile.position || '-',
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
  adminViewState.userRows = buildUserRows(adminViewState.users);
  adminViewState.projectRows = buildProjectRows(adminViewState.users);
  adminViewState.lineRows = buildLineRows(adminViewState.users);
}

function buildCustomerRows(customers){
  const rows = [];
  (Array.isArray(customers) ? customers : []).forEach(customer => {
    rows.push({
      customer: customer?.name || '-',
      address: customer?.address || '-',
      postalPlace: customer?.postalPlace || '-',
      segment: customer?.segment || '-',
      customerResponsible: customer?.customerResponsible || '-',
      projectCount: Number.isFinite(Number(customer?.projectCount)) ? Number(customer.projectCount) : 0
    });
  });
  return rows;
}

function buildCustomerContactRows(customers){
  const rows = [];
  (Array.isArray(customers) ? customers : []).forEach(customer => {
    const contacts = Array.isArray(customer?.contacts) ? customer.contacts : [];
    contacts.forEach(contact => {
      rows.push({
        customer: customer?.name || '-',
        contactPerson: contact?.name || '-',
        phone: contact?.phone || '-',
        email: contact?.email || '-'
      });
    });
  });
  return rows;
}

function renderUsersTable(rowsInput){
  const tbody = document.querySelector('#adminUsersTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="9">Ingen brukere er registrert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.name)}</td>
      <td>${escapeHtml(row.phone)}</td>
      <td>${escapeHtml(row.company)}</td>
      <td>${escapeHtml(row.position)}</td>
      <td>${escapeHtml(String(row.projectCount))}</td>
      <td>${escapeHtml(String(row.lineCount))}</td>
      <td>${escapeHtml(row.loginStatus)}</td>
      <td class="admin-row-actions">
        <button type="button" class="btn alt btn-small" data-user-edit="true" data-email="${escapeHtml(row.email)}">Endre</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function renderProjectsTable(rowsInput){
  const tbody = document.querySelector('#adminProjectsTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];

  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="13">Ingen prosjekter er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.userName)}</td>
      <td>${escapeHtml(row.userPhone)}</td>
      <td>${escapeHtml(row.userCompany)}</td>
      <td>${escapeHtml(row.userPosition)}</td>
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
    tr.innerHTML = '<td colspan="13">Ingen linjer er synkronisert enna.</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.userName)}</td>
      <td>${escapeHtml(row.userPhone)}</td>
      <td>${escapeHtml(row.userCompany)}</td>
      <td>${escapeHtml(row.userPosition)}</td>
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

function renderCustomersTable(rowsInput){
  const tbody = document.querySelector('#adminCustomersTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="6">Ingen kunder er registrert.</td>';
    tbody.appendChild(tr);
    return;
  }
  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.customer)}</td>
      <td>${escapeHtml(row.address)}</td>
      <td>${escapeHtml(row.postalPlace)}</td>
      <td>${escapeHtml(row.segment)}</td>
      <td>${escapeHtml(String(row.projectCount))}</td>
      <td class="admin-row-actions">
        <button type="button" class="btn alt btn-small" data-customer-edit="true" data-customer="${escapeHtml(row.customer)}">Endre</button>
        <button type="button" class="btn danger btn-small" data-customer-delete="true" data-customer="${escapeHtml(row.customer)}">Slett</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function renderCustomerContactsTable(rowsInput){
  const tbody = document.querySelector('#adminCustomerContactsTable tbody');
  if (!tbody) return;
  tbody.textContent = '';
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="5">Ingen kontaktpersoner er registrert.</td>';
    tbody.appendChild(tr);
    return;
  }
  rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.contactPerson)}</td>
      <td>${escapeHtml(row.phone)}</td>
      <td>${escapeHtml(row.email)}</td>
      <td>${escapeHtml(row.customer)}</td>
      <td class="admin-row-actions">
        <button type="button" class="btn alt btn-small" data-customer-contact-edit="true" data-customer="${escapeHtml(row.customer)}" data-contact="${escapeHtml(row.contactPerson)}">Endre</button>
        <button type="button" class="btn danger btn-small" data-customer-contact-delete="true" data-customer="${escapeHtml(row.customer)}" data-contact="${escapeHtml(row.contactPerson)}">Slett</button>
      </td>
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

function setCustomerStatus(message){
  const el = $('adminCustomerStatus');
  if (!el) return;
  el.textContent = message || '';
}

function setUserStatus(message){
  const el = $('adminUserStatus');
  if (!el) return;
  el.textContent = message || '';
}

function getStoredAuthOrWarn(){
  const stored = readStoredAdminAuth();
  if (!stored?.authHeader){
    setCustomerStatus('Logg inn som admin først.');
    return null;
  }
  return stored;
}

function getStoredAdminAuthOrUserWarn(){
  const stored = readStoredAdminAuth();
  if (!stored?.authHeader){
    setUserStatus('Logg inn som admin først.');
    return null;
  }
  return stored;
}

function showAdminCustomerView(){
  const overview = $('adminDataCard');
  const customerCard = $('adminCustomerCard');
  if (overview) overview.hidden = true;
  if (customerCard) customerCard.hidden = false;
}

function showAdminOverviewView(){
  const overview = $('adminDataCard');
  const customerCard = $('adminCustomerCard');
  if (overview) overview.hidden = false;
  if (customerCard) customerCard.hidden = true;
}

function findUserRow(email){
  return adminViewState.userRows.find(row=>String(row.email || '').toLowerCase() === String(email || '').toLowerCase()) || null;
}

function fillUserForm(row){
  const form = $('adminUserForm');
  if (form) form.hidden = false;
  const values = {
    adminUserOriginalEmail: row?.email || '',
    adminUserEmail: row?.email === '-' ? '' : row?.email || '',
    adminUserName: row?.rawProfile?.name || (row?.name === '-' ? '' : row?.name || ''),
    adminUserPhone: row?.rawProfile?.phone || (row?.phone === '-' ? '' : row?.phone || ''),
    adminUserCompany: row?.rawProfile?.company || (row?.company === '-' ? '' : row?.company || ''),
    adminUserPosition: row?.rawProfile?.position || (row?.position === '-' ? '' : row?.position || '')
  };
  Object.entries(values).forEach(([id, value])=>{
    const el = $(id);
    if (el) el.value = value;
  });
  setUserStatus('');
  $('adminUserName')?.focus();
}

function clearUserForm(){
  ['adminUserOriginalEmail','adminUserEmail','adminUserName','adminUserPhone','adminUserCompany','adminUserPosition'].forEach(id=>{
    const el = $(id);
    if (el) el.value = '';
  });
  const form = $('adminUserForm');
  if (form) form.hidden = true;
  setUserStatus('');
}

function getUserFormPayload(){
  return {
    email: String($('adminUserEmail')?.value || '').trim(),
    profile: {
      name: String($('adminUserName')?.value || '').trim(),
      phone: String($('adminUserPhone')?.value || '').trim(),
      company: String($('adminUserCompany')?.value || '').trim(),
      position: String($('adminUserPosition')?.value || '').trim()
    }
  };
}

function clearCustomerForm(){
  ['adminCustomerOriginalCustomer','adminCustomerOriginalContact','adminCustomerName','adminCustomerAddress','adminCustomerPostalPlace','adminCustomerSegment','adminCustomerResponsible','adminCustomerContact','adminCustomerPhone','adminCustomerEmail'].forEach(id=>{
    const el = $(id);
    if (el) el.value = '';
  });
  setCustomerFormMode('new');
  setCustomerStatus('');
}

function setCustomerFormMode(mode){
  const isCustomerMode = mode === 'customer';
  const isContactMode = mode === 'contact';
  const customerInput = $('adminCustomerName');
  const addressInput = $('adminCustomerAddress');
  const postalInput = $('adminCustomerPostalPlace');
  const segmentInput = $('adminCustomerSegment');
  const responsibleInput = $('adminCustomerResponsible');
  const contactInput = $('adminCustomerContact');
  const phoneInput = $('adminCustomerPhone');
  const emailInput = $('adminCustomerEmail');
  if (customerInput) customerInput.disabled = isContactMode;
  if (addressInput) addressInput.disabled = isContactMode;
  if (postalInput) postalInput.disabled = isContactMode;
  if (segmentInput) segmentInput.disabled = isContactMode;
  if (responsibleInput) responsibleInput.disabled = isContactMode;
  if (contactInput) contactInput.disabled = isCustomerMode;
  if (phoneInput) phoneInput.disabled = isCustomerMode;
  if (emailInput) emailInput.disabled = isCustomerMode;
}

function fillCustomerForm(row){
  const customer = String(row?.customer || '').trim();
  const contact = String(row?.contactPerson || '').trim();
  const customerRecord = adminViewState.customers.find(item=>String(item?.name || '') === customer) || {};
  const values = {
    adminCustomerOriginalCustomer: customer,
    adminCustomerOriginalContact: contact,
    adminCustomerName: customer,
    adminCustomerAddress: row?.address === '-' ? '' : row?.address || customerRecord.address || '',
    adminCustomerPostalPlace: row?.postalPlace === '-' ? '' : row?.postalPlace || customerRecord.postalPlace || '',
    adminCustomerSegment: row?.segment === '-' ? '' : row?.segment || customerRecord.segment || '',
    adminCustomerResponsible: row?.customerResponsible === '-' ? '' : row?.customerResponsible || customerRecord.customerResponsible || '',
    adminCustomerContact: contact === '-' ? '' : contact,
    adminCustomerPhone: row?.phone === '-' ? '' : row?.phone || '',
    adminCustomerEmail: row?.email === '-' ? '' : row?.email || ''
  };
  Object.entries(values).forEach(([id, value])=>{
    const el = $(id);
    if (el) el.value = value;
  });
  setCustomerFormMode(contact ? 'contact' : 'customer');
  setCustomerStatus('');
}

function findCustomerRow(customer, contact){
  if (contact){
    return adminViewState.customerContactRows.find(row=>
      String(row.customer || '') === String(customer || '') &&
      String(row.contactPerson || '') === String(contact || '')
    ) || null;
  }
  return adminViewState.customerRows.find(row=>String(row.customer || '') === String(customer || '')) || null;
}

function getCustomerFormPayload(){
  return {
    originalCustomer: String($('adminCustomerOriginalCustomer')?.value || '').trim(),
    originalContactPerson: String($('adminCustomerOriginalContact')?.value || '').trim(),
    customer: String($('adminCustomerName')?.value || '').trim(),
    address: String($('adminCustomerAddress')?.value || '').trim(),
    postalPlace: String($('adminCustomerPostalPlace')?.value || '').trim(),
    segment: String($('adminCustomerSegment')?.value || '').trim(),
    customerResponsible: String($('adminCustomerResponsible')?.value || '').trim(),
    contactPerson: String($('adminCustomerContact')?.value || '').trim(),
    phone: String($('adminCustomerPhone')?.value || '').trim(),
    email: String($('adminCustomerEmail')?.value || '').trim()
  };
}

function getCustomerRecordByName(name){
  return adminViewState.customers.find(customer=>String(customer?.name || '') === String(name || '')) || null;
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

async function fetchAdminCustomerDatabase(authHeader){
  const res = await fetch(buildApiUrl('/api/admin/customer-database'), {
    method: 'GET',
    headers: { Authorization: authHeader },
    cache: 'no-store'
  });
  if (res.status === 401){
    const err = new Error('Ugyldig admin brukernavn/passord.');
    err.code = 'AUTH';
    throw err;
  }
  if (!res.ok){
    let message = `Kunne ikke hente kundedatabase (${res.status})`;
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

async function postAdminUserProfile(authHeader, body){
  const res = await fetch(buildApiUrl('/api/admin/users/profile'), {
    method: 'POST',
    headers: {
      Authorization: authHeader,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body || {})
  });
  if (!res.ok){
    let message = `Brukeroppdatering feilet (${res.status})`;
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

async function postAdminCustomerMutation(path, authHeader, body){
  const res = await fetch(buildApiUrl(path), {
    method: 'POST',
    headers: {
      Authorization: authHeader,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body || {})
  });
  if (!res.ok){
    let message = `Kundedatabase-operasjon feilet (${res.status})`;
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

async function loadCustomerDatabase(authHeader){
  const refreshBtn = $('adminCustomerRefreshBtn');
  if (refreshBtn) refreshBtn.disabled = true;
  try{
    const data = await fetchAdminCustomerDatabase(authHeader);
    adminViewState.customers = Array.isArray(data?.customers) ? data.customers : [];
    adminViewState.customerRows = buildCustomerRows(adminViewState.customers);
    adminViewState.customerContactRows = buildCustomerContactRows(adminViewState.customers);
    renderCustomersTable(getSortedCustomerRows());
    renderCustomerContactsTable(getSortedCustomerContactRows());
    const generatedEl = $('adminCustomerGeneratedAt');
    if (generatedEl){
      generatedEl.textContent = `Sist oppdatert: ${formatTimestamp(data?.generatedAt)}`;
    }
    setCustomerStatus('');
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

async function handleUserSave(){
  const stored = getStoredAdminAuthOrUserWarn();
  if (!stored) return;
  const payload = getUserFormPayload();
  if (!payload.email){
    setUserStatus('Epost må fylles ut.');
    return;
  }
  try{
    await postAdminUserProfile(stored.authHeader, payload);
    await loadOverview(stored.authHeader);
    clearUserForm();
    setUserStatus('Brukerinfo lagret.');
  }catch(err){
    setUserStatus(err?.message || 'Lagring av brukerinfo feilet.');
  }
}

function handleAdminLogout(){
  clearStoredAdminAuth();
  clearUserForm();
  const dataCard = $('adminDataCard');
  if (dataCard) dataCard.hidden = true;
  const customerCard = $('adminCustomerCard');
  if (customerCard) customerCard.hidden = true;
  setAuthError('Logget ut.');
}

async function handleCustomerSave(){
  const stored = getStoredAuthOrWarn();
  if (!stored) return;
  const payload = getCustomerFormPayload();
  if (!payload.customer){
    setCustomerStatus('Kunde må fylles ut.');
    return;
  }
  if (payload.originalContactPerson){
    const customerRecord = getCustomerRecordByName(payload.originalCustomer || payload.customer);
    payload.customer = payload.originalCustomer || payload.customer;
    payload.address = customerRecord?.address || '';
    payload.postalPlace = customerRecord?.postalPlace || '';
    payload.segment = customerRecord?.segment || '';
    payload.customerResponsible = customerRecord?.customerResponsible || '';
  } else if (payload.originalCustomer){
    payload.contactPerson = '';
    payload.phone = '';
    payload.email = '';
  }
  try{
    const result = await postAdminCustomerMutation('/api/admin/customer-database/upsert', stored.authHeader, payload);
    adminViewState.customers = Array.isArray(result?.customers) ? result.customers : [];
    adminViewState.customerRows = buildCustomerRows(adminViewState.customers);
    adminViewState.customerContactRows = buildCustomerContactRows(adminViewState.customers);
    renderCustomersTable(getSortedCustomerRows());
    renderCustomerContactsTable(getSortedCustomerContactRows());
    clearCustomerForm();
    setCustomerStatus(`Lagret. Oppdaterte ${Number(result?.updatedProjects || 0)} prosjekt(er).`);
    await loadOverview(stored.authHeader);
    showAdminCustomerView();
  }catch(err){
    setCustomerStatus(err?.message || 'Lagring feilet.');
  }
}

async function handleCustomerDelete(deleteContactOnly){
  const stored = getStoredAuthOrWarn();
  if (!stored) return;
  const payload = getCustomerFormPayload();
  if (!payload.customer){
    setCustomerStatus('Velg eller fyll inn kunde først.');
    return;
  }
  if (deleteContactOnly && !payload.originalContactPerson && !payload.contactPerson){
    setCustomerStatus('Velg kontaktperson først.');
    return;
  }
  const body = {
    customer: payload.originalCustomer || payload.customer,
    contactPerson: deleteContactOnly ? (payload.originalContactPerson || payload.contactPerson) : ''
  };
  const text = deleteContactOnly
    ? `Fjerne kontaktpersonen "${body.contactPerson}" fra "${body.customer}"? Feltene tømmes i alle prosjekter som bruker denne kontaktpersonen.`
    : `Fjerne kunden "${body.customer}"? Kunde-, adresse-, kontakt- og telefonfelt tømmes i alle prosjekter som bruker kunden.`;
  if (!window.confirm(text)) return;
  try{
    const result = await postAdminCustomerMutation('/api/admin/customer-database/delete', stored.authHeader, body);
    adminViewState.customers = Array.isArray(result?.customers) ? result.customers : [];
    adminViewState.customerRows = buildCustomerRows(adminViewState.customers);
    adminViewState.customerContactRows = buildCustomerContactRows(adminViewState.customers);
    renderCustomersTable(getSortedCustomerRows());
    renderCustomerContactsTable(getSortedCustomerContactRows());
    clearCustomerForm();
    setCustomerStatus(`Fjernet. Oppdaterte ${Number(result?.updatedProjects || 0)} prosjekt(er).`);
    await loadOverview(stored.authHeader);
    showAdminCustomerView();
  }catch(err){
    setCustomerStatus(err?.message || 'Sletting feilet.');
  }
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
  const customersBtn = $('adminCustomersBtn');
  if (customersBtn){
    customersBtn.addEventListener('click', async () => {
      const stored = readStoredAdminAuth();
      if (!stored?.authHeader){
        setAuthError('Logg inn som admin for å åpne kundedatabase.');
        return;
      }
      showAdminCustomerView();
      try{
        await loadCustomerDatabase(stored.authHeader);
      }catch(err){
        setCustomerStatus(err?.message || 'Kunne ikke laste kundedatabase.');
      }
    });
  }
  const overviewBtn = $('adminOverviewBtn');
  if (overviewBtn){
    overviewBtn.addEventListener('click', showAdminOverviewView);
  }
  const customerRefreshBtn = $('adminCustomerRefreshBtn');
  if (customerRefreshBtn){
    customerRefreshBtn.addEventListener('click', async () => {
      const stored = getStoredAuthOrWarn();
      if (!stored) return;
      try{
        await loadCustomerDatabase(stored.authHeader);
      }catch(err){
        setCustomerStatus(err?.message || 'Oppdatering feilet.');
      }
    });
  }
  document.addEventListener('click', evt=>{
    const button = evt.target?.closest?.('.admin-col-sort-btn');
    if (button){
      const key = String(button.dataset.key || '').trim();
      if (!key) return;
      const table = button.dataset.table === 'users'
        ? 'users'
        : button.dataset.table === 'lines'
        ? 'lines'
        : button.dataset.table === 'customers'
          ? 'customers'
          : button.dataset.table === 'customerContacts'
            ? 'customerContacts'
            : 'projects';
      const type = String(button.dataset.type || 'text').trim().toLowerCase();
      setColumnSort(table, key, type);
      return;
    }
    const editUserButton = evt.target?.closest?.('[data-user-edit]');
    if (editUserButton){
      const row = findUserRow(editUserButton.dataset.email);
      if (row) fillUserForm(row);
      return;
    }
    const editButton = evt.target?.closest?.('[data-customer-edit], [data-customer-contact-edit]');
    if (editButton){
      const row = findCustomerRow(editButton.dataset.customer, editButton.dataset.contact);
      if (row) fillCustomerForm(row);
      return;
    }
    const deleteCustomerButton = evt.target?.closest?.('[data-customer-delete]');
    if (deleteCustomerButton){
      fillCustomerForm({ customer: deleteCustomerButton.dataset.customer || '' });
      void handleCustomerDelete(false);
      return;
    }
    const deleteContactButton = evt.target?.closest?.('[data-customer-contact-delete]');
    if (deleteContactButton){
      const row = findCustomerRow(deleteContactButton.dataset.customer, deleteContactButton.dataset.contact);
      if (row) fillCustomerForm(row);
      void handleCustomerDelete(true);
    }
  });
  const customerForm = $('adminCustomerForm');
  if (customerForm){
    customerForm.addEventListener('submit', evt=>{
      evt.preventDefault();
      void handleCustomerSave();
    });
  }
  const userForm = $('adminUserForm');
  if (userForm){
    userForm.addEventListener('submit', evt=>{
      evt.preventDefault();
      void handleUserSave();
    });
  }
  const userClearBtn = $('adminUserClearBtn');
  if (userClearBtn) userClearBtn.addEventListener('click', clearUserForm);
  const clearBtn = $('adminCustomerClearBtn');
  if (clearBtn) clearBtn.addEventListener('click', clearCustomerForm);
  const logoutBtn = $('adminLogoutBtn');
  if (logoutBtn){
    logoutBtn.addEventListener('click', handleAdminLogout);
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  initDashboardSidebar();
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
