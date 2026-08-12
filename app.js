// v2 + v2.1, XCP-S/XCM, distribusjon/feeder, avtappingsbokser, ekspansjon-modal >30 m

const RAW_CSV_PATHS = [
  'data/busbar-webapp-embedded-v2.csv',
  'data/busbar-webapp-embedded-v2.1.csv',
  'data/busbar-webapp-embedded-v2.2.csv',
  'data/XAP_busbar_prisliste_ekstrakt_5W_clean.csv'
];
const XAP_SERIES = 'XAP-B';
const EPOXY_IP68_SERIES = 'RCP-IP68';
const COMPARISON_ELIGIBLE_SERIES = Object.freeze(['XCM','XCP-S']);
const ENABLE_XAP_COMPARISON = false; // Pauset inntil videre
const LOCKED_LEDERE_BY_SERIES = Object.freeze({
  XCM: '3F+N+PE',
  [XAP_SERIES]: '3F+N+PE',
  [EPOXY_IP68_SERIES]: '3F+N'
});
const CRT_FEED_ALLOWED_SERIES = Object.freeze(['XCP-S', XAP_SERIES]);
const seriesLockedLedereValue = series => LOCKED_LEDERE_BY_SERIES[series] || '';
const seriesLocksLedere = series => Boolean(seriesLockedLedereValue(series));
const seriesSupportsCrtFeed = series => CRT_FEED_ALLOWED_SERIES.includes(series);
const shouldCompareXap = series => ENABLE_XAP_COMPARISON && COMPARISON_ELIGIBLE_SERIES.includes(series);
const USD_TO_NOK_RATE = 10.95; // Dagens USD→NOK-kurs (2025-10-27)
let usdToNokRate = USD_TO_NOK_RATE;
const marketDataState = { snapshot: null };
const MARKET_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const MARKET_STATUS_DEFAULT = 'Oppdateres daglig';
const MARKET_STATUS_MANUAL = 'Oppdateres manuelt';
const marketTickerState = { timerId: null };
const DEFAULT_MARGIN_RATE = 0.20;
const DEFAULT_MATERIAL_MARGIN_RATE = 0.25;
const MAX_MARGIN_RATE = 0.95;
const MAX_AUTO_MONTERING_MARGIN_RATE = 0.99;
let lastCalc = null; // delsummer for live frakt-oppdatering
let lastCalcInput = null;
const AUTH_SESSION_KEY = 'busbar.auth.session.v1';
let authState = { loggedIn: false, username: '', token: '', profile: null, isAdmin: false };
const ADMIN_NAV_ALLOWED_EMAILS = Object.freeze(['hans.jakob.risnes@busbar.no']);
const LEGACY_PROJECTS_STORAGE_KEY = 'busbar.projects.v1';
const PROJECTS_STORAGE_KEY_PREFIX = 'busbar.projects.user.v2';
const PROJECT_SYNC_DEBOUNCE_MS = 800;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MICROSOFT_AUTH_DEFAULT_SCOPES = Object.freeze(['openid', 'profile', 'email']);
const MICROSOFT_GRAPH_CALENDAR_SCOPES = Object.freeze(['Calendars.ReadWrite']);
const MICROSOFT_GRAPH_MAIL_SCOPES = Object.freeze(['Mail.ReadWrite', 'Mail.Send', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared']);
const MICROSOFT_GRAPH_SHAREPOINT_SCOPES = Object.freeze(['Files.ReadWrite.All', 'Sites.Read.All']);
const MICROSOFT_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const PROJECT_MAILBOX_ADDRESS = 'prosjekt@busbar.no';
const SHAREPOINT_FOLDER_CONFIG = Object.freeze({
  'project-folders': {
    title: 'Prosjektmapper',
    statusId: 'projectFoldersStatus',
    listId: 'projectFoldersList',
    refreshBtnId: 'refreshProjectFoldersBtn',
    siteHost: 'mcselektrotavler.sharepoint.com',
    sitePath: '/sites/BCDokumentarkiv',
    folderPath: 'Drift/BUSBAR/Prosjekter/2026'
  },
  'supplier-folders': {
    title: 'Leverandørmapper',
    statusId: 'supplierFoldersStatus',
    listId: 'supplierFoldersList',
    refreshBtnId: 'refreshSupplierFoldersBtn',
    siteHost: 'mcselektrotavler.sharepoint.com',
    sitePath: '/sites/BCDokumentarkiv',
    folderPath: 'Drift/BUSBAR/Strømskinner'
  }
});
const PROJECT_FOLDER_TEMPLATE_NAME = '0-MAL';
const DASHBOARD_SIDEBAR_COLLAPSED_KEY = 'busbar.dashboard.sidebar.collapsed.v1';
const PROJECT_SORT_STORAGE_KEY = 'busbar.project.sort.v1';
const LINE_SORT_STORAGE_KEY = 'busbar.line.sort.v1';
const OFFER_SORT_STORAGE_KEY = 'busbar.offer.sort.v1';
const PROJECT_FLOW_STORAGE_KEY = 'busbar.project.flow.v1';
const PROJECT_FLOW_ALL_PROJECTS = '__all__';
const PROJECT_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
const LINE_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
const PROJECT_STATUS_OPTIONS = Object.freeze([
  { id: 'unresolved', label: 'Uavklart', tone: 'idle' },
  { id: 'won', label: 'Vunnet', tone: 'success' },
  { id: 'lost', label: 'Tapt', tone: 'danger' },
  { id: 'finished', label: 'Ferdig', tone: 'done' }
]);
const PROJECT_ARCHIVE_STATUS_IDS = Object.freeze(['lost', 'finished']);
const PROJECT_FLOW_PHASES = Object.freeze([
  { id: 'request', label: 'Forespørsel' },
  { id: 'offer', label: 'Tilbud' },
  { id: 'order', label: 'Ordre' },
  { id: 'engineering', label: 'Prosjektering' },
  { id: 'procurement', label: 'Innkjøp' },
  { id: 'production', label: 'Produksjon' },
  { id: 'delivery', label: 'Levering' },
  { id: 'finished', label: 'Ferdig' }
]);
const PROJECT_FLOW_ZOOM_LEVELS = Object.freeze([42, 56, 72, 96, 124]);
const PROJECT_FLOW_DEFAULT_ZOOM_INDEX = 2;
const projectSyncState = {
  timerId: null,
  inFlight: false,
  pending: false
};
const projectFolderStatusState = {
  loaded: false,
  loading: false,
  byProjectId: {}
};

function resetProjectFolderStatusState(){
  projectFolderStatusState.loaded = false;
  projectFolderStatusState.loading = false;
  projectFolderStatusState.byProjectId = {};
}

const dashboardState = {
  activePage: 'dashboard',
  sidebarCollapsed: false
};
const projectState = {
  currentProjectId: null,
  currentProject: '',
  currentCustomer: '',
  currentContact: '',
  currentLineNumber: '',
  projectHistory: [],
  customerHistory: [],
  contactHistory: [],
  customerDatabase: [],
  globalCustomerDatabaseLoaded: false,
  projects: [],
  projectOwnerEmails: [],
  expandedProjectId: null,
  projectSearchTerm: '',
  projectSort: 'date_newest',
  lineSort: 'date_newest',
  showArchive: false
};
const projectModalState = {
  mode: 'create',
  projectId: null,
  saveLineAfterCreate: false,
  pendingDetails: null,
  copySourceProjectId: null
};
const projectMarginModalState = {
  projectId: null
};
const projectStatusModalState = {
  projectId: null
};
const offerDetailsWarningState = {
  resolver: null
};
const offerListState = {
  searchTerm: '',
  sort: 'date_newest',
  statusByProjectId: {},
  loaded: false
};
const globalListState = {
  companySearchTerm: '',
  companySort: 'alpha_asc',
  contactSearchTerm: '',
  contactSort: 'alpha_asc'
};
let lastEmailPayload = null;
let pendingBoxItems = [];
let pendingSpecialElementItems = [];
let tapOffItemCounter = 0;
let specialElementItemCounter = 0;
let microsoftAuthConfigPromise = null;
let microsoftMsalClient = null;
let microsoftLastAccount = null;
const calendarViewState = {
  mode: 'month',
  cursor: new Date(),
  events: [],
  loadedStart: null,
  loadedEnd: null,
  formAttendees: [],
  datePickerCursor: new Date()
};
const emailViewState = {
  messages: [],
  selectedMessageId: ''
};
const sharePointFolderState = {};
const projectFlowState = {
  selectedProjectId: PROJECT_FLOW_ALL_PROJECTS,
  milestonesByProjectId: {},
  editingMilestoneId: '',
  editingProjectId: '',
  collapsedPhaseIds: new Set(),
  zoomIndex: PROJECT_FLOW_DEFAULT_ZOOM_INDEX,
  fitDayWidth: null,
  datePickerCursor: new Date(),
  datePickerTargetId: '',
  drag: null,
  linkDrag: null,
  taskColumnWidth: 260,
  suppressClickUntil: 0
};

// --- CSV ---
function parseCSVAuto(text){
  const lines = text.replace(/\r/g,'').split('\n').filter(x=>x.length);
  if (!lines.length) return [];
  const sep = (lines[0].split(';').length > lines[0].split(',').length) ? ';' : ',';
  const parseLine = (line)=>{
    const out=[]; let f='', q=false;
    for(let i=0;i<line.length;i++){
      const c=line[i];
      if(c === '"'){ if(q && line[i+1] === '"'){ f+='"'; i++; } else { q=!q; } continue; }
      if(!q && c === sep){ out.push(f); f=''; continue; }
      f += c;
    }
    out.push(f); return out;
  };
  const header = parseLine(lines[0]).map(h=>h.trim());
  return lines.slice(1)
    .map(parseLine)
    .filter(r=>r.some(x=>x && x.trim()!==''))
    .map(cols=>{
      const obj = Object.fromEntries(header.map((h,i)=>[h, cols[i]??'']));
      obj._cols = cols;
      return obj;
    });
}

// --- utils ---
const $ = id=>document.getElementById(id);
const hasDashboardUI = ()=>Boolean($('dashboardView') && $('projectList'));
const hasCalculatorUI = ()=>Boolean($('calcBtn') && $('series'));

function ensureFormModal(formId, titleText){
  const form = $(formId);
  if (!form) return null;
  const modalId = `${formId}Modal`;
  let modal = $(modalId);
  if (!modal){
    modal = document.createElement('div');
    modal.id = modalId;
    modal.className = 'form-modal-backdrop';
    modal.hidden = true;
    modal.innerHTML = `
      <div class="form-modal-panel" role="dialog" aria-modal="true" aria-labelledby="${modalId}Title">
        <div class="form-modal-header">
          <h3 id="${modalId}Title"></h3>
          <button type="button" class="form-modal-close" aria-label="Lukk">x</button>
        </div>
        <div class="form-modal-body"></div>
      </div>
    `;
    document.body.appendChild(modal);
    modal.addEventListener('click', evt=>{
      if (evt.target === modal) closeFormModal(formId);
    });
    modal.querySelector('.form-modal-close')?.addEventListener('click', ()=>closeFormModal(formId));
  }
  const title = modal.querySelector('.form-modal-header h3');
  if (title) title.textContent = titleText || '';
  const body = modal.querySelector('.form-modal-body');
  if (body && form.parentElement !== body){
    body.appendChild(form);
  }
  return modal;
}

function openFormModal(formId, titleText){
  const modal = ensureFormModal(formId, titleText);
  if (!modal) return;
  modal.hidden = false;
  document.body.classList.add('has-form-modal-open');
}

function closeFormModal(formId){
  const form = $(formId);
  if (form) form.hidden = true;
  const modal = $(`${formId}Modal`);
  if (modal) modal.hidden = true;
  if (!document.querySelector('.form-modal-backdrop:not([hidden])')){
    document.body.classList.remove('has-form-modal-open');
  }
}

function buildAppUrl(fileName, params = {}){
  const url = new URL(fileName, window.location.href);
  Object.entries(params).forEach(([key, value])=>{
    if (value === undefined || value === null || value === ''){
      return;
    }
    url.searchParams.set(key, String(value));
  });
  return url;
}

function goToDashboard(params = {}){
  window.location.href = buildAppUrl('index.html', params).toString();
}

function goToCalculator(params = {}){
  window.location.href = buildAppUrl('calculator.html', params).toString();
}

function updateMarketTickerVisibility(){
  const ticker = $('marketTicker');
  if (!ticker) return;
  const shouldShow = hasCalculatorUI() || dashboardState.activePage === 'dashboard';
  ticker.hidden = !shouldShow;
}

function setDashboardPage(page){
  const nextPage = String(page || 'projects');
  const pages = Array.from(document.querySelectorAll('[data-dashboard-page]'));
  if (pages.length){
    const target = pages.find(el=>el.dataset.dashboardPage === nextPage) || pages[0];
    dashboardState.activePage = target.dataset.dashboardPage || 'projects';
    pages.forEach(el=>{
      const active = el === target;
      el.hidden = !active;
      el.classList.toggle('is-active', active);
    });
  } else {
    dashboardState.activePage = nextPage;
  }
  const navItems = Array.from(document.querySelectorAll('[data-dashboard-nav], [data-dashboard-link]'));
  navItems.forEach(item=>{
    const itemPage = item.dataset.dashboardNav || item.dataset.dashboardLink || '';
    const active = itemPage === dashboardState.activePage;
    item.classList.toggle('is-active', active);
    item.setAttribute('aria-current', active ? 'page' : 'false');
  });
  updateMarketTickerVisibility();
  handleDashboardPageActivated(dashboardState.activePage);
}

function handleDashboardPageActivated(page){
  if (page === 'calendar'){
    loadCalendarEvents({ silent: true });
  } else if (page === 'email'){
    loadEmailMessages({ silent: true });
  } else if (page === 'offers'){
    loadOfferStatus({ silent: true });
  } else if (page === 'project-flow'){
    renderProjectFlowView();
  } else if (page === 'customers' || page === 'company-card'){
    loadGlobalCustomerDatabase({ silent: true });
  } else if (page === 'project-folders' || page === 'supplier-folders'){
    loadSharePointFolder(page, { silent: true });
  }
}

function resolveInitialDashboardPage(){
  const shell = $('dashboardShell');
  const configured = String(shell?.dataset?.dashboardActivePage || '').trim();
  if (configured) return configured;
  const params = new URLSearchParams(window.location.search);
  const view = String(params.get('view') || '').trim();
  if (view) return view;
  return hasCalculatorUI() ? 'calculator' : 'dashboard';
}

function readDashboardSidebarCollapsed(){
  if (typeof localStorage === 'undefined') return false;
  try{
    return localStorage.getItem(DASHBOARD_SIDEBAR_COLLAPSED_KEY) === '1';
  }catch(_err){
    return false;
  }
}

function persistDashboardSidebarCollapsed(collapsed){
  if (typeof localStorage === 'undefined') return;
  try{
    localStorage.setItem(DASHBOARD_SIDEBAR_COLLAPSED_KEY, collapsed ? '1' : '0');
  }catch(_err){}
}

function setDashboardSidebarCollapsed(collapsed, options = {}){
  const shell = $('dashboardShell');
  const toggle = $('dashboardSidebarToggle');
  const next = Boolean(collapsed);
  dashboardState.sidebarCollapsed = next;
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

function initDashboardShell(){
  const shell = $('dashboardShell');
  if (!shell) return;
  if (shell.dataset.dashboardShellBound === '1'){
    setDashboardPage(dashboardState.activePage);
    setDashboardSidebarCollapsed(dashboardState.sidebarCollapsed, { skipPersist: true });
    return;
  }
  shell.dataset.dashboardShellBound = '1';
  dashboardState.activePage = resolveInitialDashboardPage();
  dashboardState.sidebarCollapsed = readDashboardSidebarCollapsed();
  shell.addEventListener('click', evt=>{
    const toggleBtn = evt.target?.closest?.('#dashboardSidebarToggle');
    if (toggleBtn){
      evt.preventDefault();
      setDashboardSidebarCollapsed(!dashboardState.sidebarCollapsed);
      return;
    }
    const navBtn = evt.target?.closest?.('[data-dashboard-nav]');
    if (!navBtn) return;
    evt.preventDefault();
    setDashboardPage(navBtn.dataset.dashboardNav);
  });
  setDashboardSidebarCollapsed(dashboardState.sidebarCollapsed, { skipPersist: true });
  setDashboardPage(dashboardState.activePage);
}

initDashboardShell();

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
  return `${errorText}. GitHub Pages kjører kun statisk frontend. Sett <meta name="busbar-api-base" ...> til backend-URL.`;
}

function buildStaticAssetUrl(relativePath){
  try{
    return new URL(String(relativePath || ''), window.location.href).toString();
  }catch(_err){
    return String(relativePath || '');
  }
}

const round2 = n=>Math.round(n*100)/100;
const fmtNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtIntNO = new Intl.NumberFormat('no-NO', { maximumFractionDigits: 0 });
const fmtFxNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtTimestampNO = new Intl.DateTimeFormat('no-NO', { dateStyle: 'short', timeStyle: 'short' });
const fmtPercentNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
const fmtMarketPercentNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const MIN_MONTERING_TOTAL = 30000;
let currentMarginRate = DEFAULT_MATERIAL_MARGIN_RATE;
let currentMontasjeMarginRate = DEFAULT_MARGIN_RATE;
let currentEngineeringMarginRate = DEFAULT_MARGIN_RATE;
let currentOpphengMarginRate = DEFAULT_MARGIN_RATE;
let currentTapOffMarginRate = DEFAULT_MARGIN_RATE;
let currentDgModalTarget = 'material';
const toNum = x => {
  if (x===undefined || x===null) return NaN;
  const v = Number(String(x).replace(/\s/g,'').replace(',','.'));
  return Number.isFinite(v) ? v : NaN;
};
function pick(row, names){ for (const n of names){ if (n in row && row[n]!=='' && row[n]!==undefined) return row[n]; } return ''; }

function normalizeBoxItem(item){
  if (!item || typeof item !== 'object') return null;
  const existingId = String(item.id || item.tapOffGroupId || '').trim();
  const boxSel = String(item.boxSel || item.value || '').trim();
  const qtyRaw = Number(item.boxQty ?? item.qty ?? 0);
  const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
  const innmatRaw = toNum(item.innmatSum ?? item.innmat ?? 0);
  const innmatSum = Number.isFinite(innmatRaw) ? Math.max(0, round2(innmatRaw)) : 0;
  if (!boxSel || qty <= 0) return null;
  return { id: existingId || generateTapOffItemId(), boxSel, boxQty: qty, innmatSum };
}

function normalizeBoxItems(items, legacyBoxSel = '', legacyBoxQty = 0, legacyInnmatSum = 0){
  const out = [];
  if (Array.isArray(items)){
    items.forEach(item=>{
      const normalized = normalizeBoxItem(item);
      if (normalized) out.push(normalized);
    });
  }
  if (!out.length){
    const fallback = normalizeBoxItem({
      boxSel: legacyBoxSel,
      boxQty: legacyBoxQty,
      innmatSum: legacyInnmatSum
    });
    if (fallback) out.push(fallback);
  }
  return out;
}

function boxLabelFromSelection(value){
  const sel = $('boxSel');
  if (!sel || !value) return String(value || '');
  const match = [...sel.options].find(opt=>opt.value === value);
  return match ? String(match.textContent || value) : String(value);
}

function generateTapOffItemId(){
  tapOffItemCounter += 1;
  return `tapoff-${Date.now()}-${tapOffItemCounter}`;
}

function normalizeSpecialElementItem(item){
  if (!item || typeof item !== 'object') return null;
  const selection = String(item.selection || item.value || item.type || '').trim();
  const qtyRaw = Number(item.qty ?? item.quantity ?? 0);
  const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
  const unitSumRaw = toNum(item.unitSum ?? item.sum ?? item.elementSum ?? 0);
  const unitSum = Number.isFinite(unitSumRaw) ? Math.max(0, round2(unitSumRaw)) : 0;
  if (!selection || qty <= 0) return null;
  const existingId = String(item.id || item.specialElementGroupId || '').trim();
  return { id: existingId || generateSpecialElementItemId(), selection, qty, unitSum };
}

function normalizeSpecialElementItems(items){
  if (!Array.isArray(items)) return [];
  return items.map(normalizeSpecialElementItem).filter(Boolean);
}

function generateSpecialElementItemId(){
  specialElementItemCounter += 1;
  return `special-element-${Date.now()}-${specialElementItemCounter}`;
}

function specialElementLabelFromSelection(value){
  const sel = $('specialElementType');
  if (!sel || !value) return String(value || '');
  const match = [...sel.options].find(opt=>opt.value === value);
  return match ? String(match.textContent || value) : String(value);
}

function isSeparateTapOffBoxType(value){
  return ['plug_in_box', 'tap_off_box'].includes(String(value || '').trim().toLowerCase());
}

function isTapOffInnmatType(value){
  return ['plug_in_box_innmat', 'tap_off_box_innmat'].includes(String(value || '').trim().toLowerCase());
}

function isSeparateTapOffBoxBomLine(entry){
  const type = entry?.type || entry?.element_type || entry?.elementType;
  return isSeparateTapOffBoxType(type) || isTapOffInnmatType(type);
}

function isSeparateSpecialElementBomLine(entry){
  return Boolean(String(entry?.specialElementGroupId || '').trim());
}

function resolveBomLineSum(entry){
  const direct = Number(entry?.sum);
  if (Number.isFinite(direct)) return direct;
  const unit = Number(entry?.enhet ?? entry?.unit ?? entry?.unit_price);
  const qty = Number(entry?.antall ?? entry?.qty ?? entry?.quantity);
  if (Number.isFinite(unit) && Number.isFinite(qty)) return unit * qty;
  return 0;
}

function sumSeparateTapOffBoxTotal(bomList){
  return round2((Array.isArray(bomList) ? bomList : []).reduce((sum, entry)=>{
    if (!isSeparateTapOffBoxBomLine(entry)) return sum;
    return sum + resolveBomLineSum(entry);
  }, 0));
}

function resolveTapOffOfferRows(source = lastCalc){
  const bomList = Array.isArray(source?.bom)
    ? source.bom
    : (Array.isArray(lastEmailPayload?.bom) ? lastEmailPayload.bom : []);
  const groups = new Map();
  bomList.forEach((entry, index)=>{
    if (!isSeparateTapOffBoxBomLine(entry)) return;
    const groupId = String(entry.tapOffGroupId || entry.tapOffBoxSel || `tapoff-row-${index}`).trim();
    if (!groups.has(groupId)){
      groups.set(groupId, {
        id: groupId,
        label: '',
        cost: 0,
        qty: 0,
        dgRate: normalizeMarginRate(source?.tapOffMarginRate ?? currentTapOffMarginRate, DEFAULT_MARGIN_RATE)
      });
    }
    const group = groups.get(groupId);
    group.cost = round2(group.cost + resolveBomLineSum(entry));
    if (!entry.tapOffInnmatLine){
      group.label = boxLabelFromSelection(entry.tapOffBoxSel) || String(entry.type || 'Avtappingsboks');
      const qty = Number(entry.antall ?? entry.qty ?? entry.quantity);
      group.qty = Number.isFinite(qty) ? qty : group.qty;
    }
  });
  return Array.from(groups.values()).map(group=>{
    const pricing = calculateDgPricing(group.cost, group.dgRate);
    return {
      ...group,
      dg: pricing.dg,
      total: pricing.totalWithDg
    };
  });
}

function calculateTapOffOfferTotal(source = lastCalc){
  return round2(resolveTapOffOfferRows(source).reduce((sum, row)=>sum + (Number(row.total) || 0), 0));
}

function resolveSpecialElementOfferRows(source = lastCalc){
  const bomList = Array.isArray(source?.bom)
    ? source.bom
    : (Array.isArray(lastEmailPayload?.bom) ? lastEmailPayload.bom : []);
  return bomList
    .filter(isSeparateSpecialElementBomLine)
    .map((entry, index)=>{
      const cost = round2(resolveBomLineSum(entry));
      const dgRate = normalizeMarginRate(source?.tapOffMarginRate ?? currentTapOffMarginRate, DEFAULT_MARGIN_RATE);
      const pricing = calculateDgPricing(cost, dgRate);
      return {
        id: String(entry.specialElementGroupId || `special-element-row-${index}`),
        label: String(entry.type || specialElementLabelFromSelection(entry.specialElementSelection) || 'Spesialelement'),
        qty: Number(entry.antall ?? entry.qty ?? entry.quantity) || 0,
        cost,
        dgRate,
        dg: pricing.dg,
        total: pricing.totalWithDg
      };
    });
}

function calculateSpecialElementOfferTotal(source = lastCalc){
  return round2(resolveSpecialElementOfferRows(source).reduce((sum, row)=>sum + (Number(row.total) || 0), 0));
}

function renderTapOffOfferRows(){
  const container = $('tapOffOfferRows');
  if (!container) return;
  const rows = resolveTapOffOfferRows(lastCalc);
  container.innerHTML = '';
  container.hidden = rows.length === 0;
  rows.forEach(row=>{
    const line = document.createElement('div');
    line.className = 'totals-line tap-off-offer-row';
    line.innerHTML = `
      <div class="total-item"><strong>${row.label || 'Avtappingsboks'}${row.qty ? ` x ${fmtIntNO.format(row.qty)}` : ''}:</strong> <span>${fmtNO.format(row.cost)}</span></div>
      <div class="total-item margin-item tap-off-margin-item">
        <button type="button" class="btn alt margin-config-btn" data-tap-off-dg>Endre</button>
        <strong>DG ${fmtPercentNO.format(row.dgRate * 100)} %:</strong>
        <span>${fmtNO.format(row.dg)}</span>
      </div>
      <div class="total-item tap-off-total-item"><strong>Total:</strong> <span>${fmtNO.format(row.total)}</span> NOK eks. mva <button type="button" class="btn danger btn-small tap-off-delete-btn" data-delete-tap-off="${row.id}">Slett</button></div>
    `;
    container.appendChild(line);
  });
  const total = calculateTapOffOfferTotal(lastCalc);
  if (total > 0){
    const summary = document.createElement('div');
    summary.className = 'totals-line tap-off-offer-row tap-off-summary-row';
    summary.innerHTML = `
      <div class="total-item"><strong>Total avtappingsbokser:</strong></div>
      <div class="total-item"></div>
      <div class="total-item tap-off-total-item"><strong><span id="tapOffOfferTotal">${fmtNO.format(total)}</span></strong> NOK eks. mva</div>
    `;
    container.appendChild(summary);
  }
}

function renderSpecialElementOfferRows(){
  const container = $('specialElementOfferRows');
  if (!container) return;
  const rows = resolveSpecialElementOfferRows(lastCalc);
  container.innerHTML = '';
  container.hidden = rows.length === 0;
  rows.forEach(row=>{
    const line = document.createElement('div');
    line.className = 'totals-line tap-off-offer-row';
    line.innerHTML = `
      <div class="total-item"><strong>${row.label || 'Spesialelement'}${row.qty ? ` x ${fmtIntNO.format(row.qty)}` : ''}:</strong> <span>${fmtNO.format(row.cost)}</span></div>
      <div class="total-item margin-item tap-off-margin-item">
        <button type="button" class="btn alt margin-config-btn" data-special-element-dg>Endre</button>
        <strong>DG ${fmtPercentNO.format(row.dgRate * 100)} %:</strong>
        <span>${fmtNO.format(row.dg)}</span>
      </div>
      <div class="total-item tap-off-total-item"><strong>Total:</strong> <span>${fmtNO.format(row.total)}</span> NOK eks. mva <button type="button" class="btn danger btn-small tap-off-delete-btn" data-delete-special-element="${row.id}">Slett</button></div>
    `;
    container.appendChild(line);
  });
  const total = calculateSpecialElementOfferTotal(lastCalc);
  if (total > 0){
    const summary = document.createElement('div');
    summary.className = 'totals-line tap-off-offer-row tap-off-summary-row';
    summary.innerHTML = `
      <div class="total-item"><strong>Total spesialelementer:</strong></div>
      <div class="total-item"></div>
      <div class="total-item tap-off-total-item"><strong><span id="specialElementOfferTotal">${fmtNO.format(total)}</span></strong> NOK eks. mva</div>
    `;
    container.appendChild(summary);
  }
}

function syncSelectedAddonTotalToPayload(total){
  const safeTotal = round2(Number(total) || 0);
  if (lastCalc){
    lastCalc.selectedAddonTotal = safeTotal;
  }
  if (lastEmailPayload?.totals){
    lastEmailPayload.totals.selectedAddonTotal = safeTotal;
    lastEmailPayload.totals.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
    lastEmailPayload.totals.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
    lastEmailPayload.totals.tapOffMarginRate = normalizeMarginRate(lastCalc?.tapOffMarginRate ?? currentTapOffMarginRate, DEFAULT_MARGIN_RATE);
  }
}

function renderPendingBoxItems(){
  const row = $('boxItemsRow');
  const preview = $('boxItemsPreview');
  if (!row || !preview) return;
  if (!pendingBoxItems.length){
    row.hidden = true;
    preview.textContent = '';
    return;
  }
  row.hidden = false;
  const text = pendingBoxItems.map(item=>{
    const boxTxt = boxLabelFromSelection(item.boxSel);
    const qtyTxt = fmtIntNO.format(item.boxQty);
    const innmatTxt = fmtNO.format(item.innmatSum || 0);
    return `${boxTxt} · antall ${qtyTxt} · innmat ${innmatTxt} kr/stk`;
  }).join(' | ');
  preview.textContent = `Valgte bokser: ${text}`;
}

function renderPendingSpecialElementItems(){
  const row = $('specialElementItemsRow');
  const preview = $('specialElementItemsPreview');
  if (!row || !preview) return;
  if (!pendingSpecialElementItems.length){
    row.hidden = true;
    preview.textContent = '';
    return;
  }
  row.hidden = false;
  const text = pendingSpecialElementItems
    .map(item=>`${specialElementLabelFromSelection(item.selection)} · antall ${fmtIntNO.format(item.qty)} · ${fmtNO.format(item.unitSum)} kr/stk`)
    .join(' | ');
  preview.textContent = `Valgte spesialelementer: ${text}`;
}

function convertUsdToNok(value){
  if (!Number.isFinite(value)) return 0;
  return round2(value * usdToNokRate);
}

function deepClone(value){
  if (value === null || value === undefined) return value;
  if (typeof structuredClone === 'function'){
    try{
      return structuredClone(value);
    }catch(_err){
      /* fallback below */
    }
  }
  try{
    return JSON.parse(JSON.stringify(value));
  }catch(_err){
    return value;
  }
}

function normalizeMarginRate(value, fallback = DEFAULT_MARGIN_RATE, maxRate = MAX_MARGIN_RATE){
  const raw = Number(value);
  if (!Number.isFinite(raw)) return fallback;
  const asRate = raw > 1 ? raw / 100 : raw;
  if (!Number.isFinite(asRate)) return fallback;
  if (asRate < 0) return 0;
  const safeMaxRate = Number.isFinite(maxRate) && maxRate > 0 && maxRate < 1 ? maxRate : MAX_MARGIN_RATE;
  if (asRate >= 1) return safeMaxRate;
  return Math.min(safeMaxRate, asRate);
}

function marginFactorFromRate(rate){
  return 1 - normalizeMarginRate(rate);
}

function resolveMarginRateFromData({ totals, input } = {}){
  const fromInput = normalizeMarginRate(input?.marginRate, NaN);
  if (Number.isFinite(fromInput)) return fromInput;
  const fromTotals = normalizeMarginRate(totals?.marginRate, NaN);
  if (Number.isFinite(fromTotals)) return fromTotals;
  const material = Number(totals?.material);
  const subtotal = Number(totals?.subtotal);
  if (Number.isFinite(material) && Number.isFinite(subtotal) && subtotal > 0){
    return normalizeMarginRate(1 - material / subtotal, DEFAULT_MATERIAL_MARGIN_RATE);
  }
  return DEFAULT_MATERIAL_MARGIN_RATE;
}

function resolveDgRate(fromInput, fromTotals, fallback = DEFAULT_MARGIN_RATE){
  const inputRate = normalizeMarginRate(fromInput, NaN);
  if (Number.isFinite(inputRate)) return inputRate;
  const totalsRate = normalizeMarginRate(fromTotals, NaN);
  if (Number.isFinite(totalsRate)) return totalsRate;
  return fallback;
}

function resolveMontasjeDgRate(fromInput, fromTotals, fallback = DEFAULT_MARGIN_RATE){
  const inputRate = normalizeMarginRate(fromInput, NaN, MAX_AUTO_MONTERING_MARGIN_RATE);
  if (Number.isFinite(inputRate)) return inputRate;
  const totalsRate = normalizeMarginRate(fromTotals, NaN, MAX_AUTO_MONTERING_MARGIN_RATE);
  if (Number.isFinite(totalsRate)) return totalsRate;
  return fallback;
}

function updateDgLabel(labelId, rate, maxRate = MAX_MARGIN_RATE){
  const normalizedRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE, maxRate);
  const percentTxt = fmtPercentNO.format(normalizedRate * 100);
  const labelEl = $(labelId);
  if (labelEl){
    labelEl.textContent = `DG ${percentTxt} %:`;
  }
}

function updateMarginUI(){
  updateDgLabel('marginLabel', currentMarginRate);
  updateDgLabel('montasjeDgLabel', currentMontasjeMarginRate, MAX_AUTO_MONTERING_MARGIN_RATE);
  updateDgLabel('engineeringDgLabel', currentEngineeringMarginRate);
  updateDgLabel('opphengDgLabel', currentOpphengMarginRate);
  const marginBtn = $('marginConfigBtn');
  if (marginBtn){
    marginBtn.textContent = 'Endre';
  }
  const montasjeBtn = $('montasjeDgConfigBtn');
  if (montasjeBtn){
    montasjeBtn.textContent = 'Endre';
  }
  const engineeringBtn = $('engineeringDgConfigBtn');
  if (engineeringBtn){
    engineeringBtn.textContent = 'Endre';
  }
  const opphengBtn = $('opphengDgConfigBtn');
  if (opphengBtn){
    opphengBtn.textContent = 'Endre';
  }
}

function setCurrentMarginRate(rate){
  currentMarginRate = normalizeMarginRate(rate, DEFAULT_MATERIAL_MARGIN_RATE);
  updateMarginUI();
  return currentMarginRate;
}

function setCurrentMontasjeMarginRate(rate){
  currentMontasjeMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE, MAX_AUTO_MONTERING_MARGIN_RATE);
  updateMarginUI();
  return currentMontasjeMarginRate;
}

function setCurrentEngineeringMarginRate(rate){
  currentEngineeringMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
  updateMarginUI();
  return currentEngineeringMarginRate;
}

function setCurrentOpphengMarginRate(rate){
  currentOpphengMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
  updateMarginUI();
  return currentOpphengMarginRate;
}

function setCurrentTapOffMarginRate(rate){
  currentTapOffMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
  updateMarginUI();
  return currentTapOffMarginRate;
}

function calculateSelectedAddonTotal(calc){
  if (!calc) return { base: 0, total: 0 };
  const baseTotal = round2(Number(calc.totalExMontasje) || 0);
  const includeMontasje = Boolean($('includeMontasje')?.checked);
  const includeEngineering = Boolean($('includeEngineering')?.checked);
  const includeOppheng = Boolean($('includeOppheng')?.checked);
  const montasjeTotal = Number(calc.totalInclMontasje);
  const engineeringTotal = Number(calc.totalInclEngineering);
  const opphengTotal = Number(calc.totalInclOppheng ?? calc.total);
  let sum = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) sum += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) sum += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) sum += opphengTotal;
  sum += calculateTapOffOfferTotal(calc);
  sum += calculateSpecialElementOfferTotal(calc);
  return { base: baseTotal, total: round2(sum) };
}

function updateSelectedAddonTotalUI(){
  const totalEl = $('selectedAddonTotal');
  renderTapOffOfferRows();
  renderSpecialElementOfferRows();
  if (!totalEl) return;
  const sum = calculateSelectedAddonTotal(lastCalc);
  totalEl.textContent = fmtNO.format(Number.isFinite(sum.total) ? sum.total : 0);
  syncSelectedAddonTotalToPayload(sum.total);
}

function calculateDgPricing(baseCost, dgRate, maxRate = MAX_MARGIN_RATE){
  const base = round2(Number(baseCost) || 0);
  const normalizedRate = normalizeMarginRate(dgRate, DEFAULT_MARGIN_RATE, maxRate);
  const factor = 1 - normalizedRate;
  if (!(factor > 0)){
    throw new Error('DG-faktor må være større enn 0.');
  }
  const totalWithDg = round2(base / factor);
  const dg = round2(totalWithDg - base);
  return { base, dgRate: normalizedRate, dg, totalWithDg };
}

function calculateMontasjePricing(baseCost, dgRate){
  const regularPricing = calculateDgPricing(baseCost, dgRate, MAX_AUTO_MONTERING_MARGIN_RATE);
  if (regularPricing.totalWithDg >= MIN_MONTERING_TOTAL || regularPricing.base <= 0){
    return regularPricing;
  }

  const targetRatePercent = Math.ceil(((1 - (regularPricing.base / MIN_MONTERING_TOTAL)) * 100) - 1e-9);
  const boundedRatePercent = Math.max(0, Math.min(Math.round(MAX_AUTO_MONTERING_MARGIN_RATE * 100), targetRatePercent));
  const minPricing = calculateDgPricing(regularPricing.base, boundedRatePercent / 100, MAX_AUTO_MONTERING_MARGIN_RATE);

  return minPricing.totalWithDg > regularPricing.totalWithDg
    ? minPricing
    : regularPricing;
}

function calculateTotalsFromMaterial({
  material,
  marginRate,
  freightRate,
  montasjeCost = 0,
  montasjeMarginRate = DEFAULT_MARGIN_RATE,
  engineeringCost = 0,
  engineeringMarginRate = DEFAULT_MARGIN_RATE,
  opphengCost = 0,
  opphengMarginRate = DEFAULT_MARGIN_RATE
}){
  const normalizedMarginRate = normalizeMarginRate(marginRate, DEFAULT_MATERIAL_MARGIN_RATE);
  const factor = marginFactorFromRate(normalizedMarginRate);
  if (!(factor > 0)){
    throw new Error('DG-faktor må være større enn 0.');
  }
  const safeMaterial = round2(Number(material) || 0);
  const safeFreightRate = Number(freightRate);
  const appliedFreightRate = Number.isFinite(safeFreightRate) ? safeFreightRate : 0;
  const subtotal = round2(safeMaterial / factor);
  const margin = round2(subtotal - safeMaterial);
  const freight = round2(safeMaterial * appliedFreightRate);
  const totalExMontasje = round2(subtotal + freight);
  const montasjePricing = calculateMontasjePricing(montasjeCost, montasjeMarginRate);
  const engineeringPricing = calculateDgPricing(engineeringCost, engineeringMarginRate);
  const opphengPricing = calculateDgPricing(opphengCost, opphengMarginRate);
  const totalInclMontasje = round2(montasjePricing.totalWithDg);
  const totalInclEngineering = round2(engineeringPricing.totalWithDg);
  const totalInclOppheng = round2(opphengPricing.totalWithDg);
  const total = totalInclOppheng;
  return {
    material: safeMaterial,
    marginRate: normalizedMarginRate,
    marginFactor: factor,
    freightRate: appliedFreightRate,
    margin,
    montasjeMarginRate: montasjePricing.dgRate,
    montasjeMargin: montasjePricing.dg,
    montasjeTotalWithDg: montasjePricing.totalWithDg,
    engineeringMarginRate: engineeringPricing.dgRate,
    engineeringMargin: engineeringPricing.dg,
    engineeringTotalWithDg: engineeringPricing.totalWithDg,
    opphengMarginRate: opphengPricing.dgRate,
    opphengMargin: opphengPricing.dg,
    opphengTotalWithDg: opphengPricing.totalWithDg,
    subtotal,
    freight,
    totalExMontasje,
    totalInclMontasje,
    totalInclEngineering,
    totalInclOppheng,
    total
  };
}

function recalcLastTotalsFromCurrentRates(){
  if (!lastCalc) return;
  const rate = Number(document.getElementById('freightRate')?.value ?? lastCalcInput?.freightRate ?? 0.10);
  const recalculated = calculateTotalsFromMaterial({
    material: Number(lastCalc.material) || 0,
    marginRate: currentMarginRate,
    freightRate: rate,
    montasjeCost: Number(lastCalc.montasje?.cost) || 0,
    montasjeMarginRate: currentMontasjeMarginRate,
    engineeringCost: Number(lastCalc.engineering?.cost) || 0,
    engineeringMarginRate: currentEngineeringMarginRate,
    opphengCost: Number(lastCalc.oppheng?.cost) || 0,
    opphengMarginRate: currentOpphengMarginRate
  });
  const setText = (id, value)=>{
    const el = $(id);
    if (!el) return;
    el.textContent = fmtNO.format(Number(value) || 0);
  };
  setText('margin', recalculated.margin);
  setText('subtotal', recalculated.subtotal);
  setText('freight', recalculated.freight);
  setText('totalExMontasje', recalculated.totalExMontasje);
  setText('montasjeMargin', recalculated.montasjeMargin);
  setText('engineeringMargin', recalculated.engineeringMargin);
  setText('opphengMargin', recalculated.opphengMargin);
  setText('totalInclMontasje', recalculated.totalInclMontasje);
  setText('totalInclEngineering', recalculated.totalInclEngineering);
  setText('total', recalculated.total);
  currentMontasjeMarginRate = recalculated.montasjeMarginRate;
  updateMarginUI();
  Object.assign(lastCalc, {
    marginRate: recalculated.marginRate,
    marginFactor: recalculated.marginFactor,
    margin: recalculated.margin,
    montasjeMarginRate: recalculated.montasjeMarginRate,
    montasjeMargin: recalculated.montasjeMargin,
    engineeringMarginRate: recalculated.engineeringMarginRate,
    engineeringMargin: recalculated.engineeringMargin,
    opphengMarginRate: recalculated.opphengMarginRate,
    opphengMargin: recalculated.opphengMargin,
    tapOffMarginRate: currentTapOffMarginRate,
    subtotal: recalculated.subtotal,
    freight: recalculated.freight,
    totalExMontasje: recalculated.totalExMontasje,
    totalInclMontasje: recalculated.totalInclMontasje,
    totalInclEngineering: recalculated.totalInclEngineering,
    totalInclOppheng: recalculated.totalInclOppheng,
    total: recalculated.total
  });
  lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
  lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
  if (lastCalcInput){
    lastCalcInput.marginRate = recalculated.marginRate;
    lastCalcInput.freightRate = recalculated.freightRate;
    lastCalcInput.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastCalcInput.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastCalcInput.opphengMarginRate = recalculated.opphengMarginRate;
    lastCalcInput.tapOffMarginRate = currentTapOffMarginRate;
  }
  if (lastEmailPayload?.inputs){
    lastEmailPayload.inputs.marginRate = recalculated.marginRate;
    lastEmailPayload.inputs.freightRate = recalculated.freightRate;
    lastEmailPayload.inputs.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastEmailPayload.inputs.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastEmailPayload.inputs.opphengMarginRate = recalculated.opphengMarginRate;
    lastEmailPayload.inputs.tapOffMarginRate = currentTapOffMarginRate;
  }
  if (lastEmailPayload?.totals){
    lastEmailPayload.totals.marginRate = recalculated.marginRate;
    lastEmailPayload.totals.tapOffBoxTotal = lastCalc.tapOffBoxTotal || 0;
    lastEmailPayload.totals.margin = recalculated.margin;
    lastEmailPayload.totals.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastEmailPayload.totals.montasjeMargin = recalculated.montasjeMargin;
    lastEmailPayload.totals.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastEmailPayload.totals.engineeringMargin = recalculated.engineeringMargin;
    lastEmailPayload.totals.opphengMarginRate = recalculated.opphengMarginRate;
    lastEmailPayload.totals.tapOffMarginRate = currentTapOffMarginRate;
    lastEmailPayload.totals.opphengMargin = recalculated.opphengMargin;
    lastEmailPayload.totals.subtotal = recalculated.subtotal;
    lastEmailPayload.totals.freight = recalculated.freight;
    lastEmailPayload.totals.totalExMontasje = recalculated.totalExMontasje;
    lastEmailPayload.totals.totalInclMontasje = recalculated.totalInclMontasje;
    lastEmailPayload.totals.totalInclEngineering = recalculated.totalInclEngineering;
    lastEmailPayload.totals.totalInclOppheng = recalculated.totalInclOppheng;
    lastEmailPayload.totals.total = recalculated.total;
    lastEmailPayload.totals.tapOffOfferTotal = lastCalc.tapOffOfferTotal || 0;
    lastEmailPayload.totals.specialElementOfferTotal = lastCalc.specialElementOfferTotal || 0;
  }
  updateSelectedAddonTotalUI();
}

function getDgModalTitleByTarget(target){
  if (target === 'montasje') return 'Endre DG for montasje';
  if (target === 'engineering') return 'Endre DG for ingeniør';
  if (target === 'oppheng') return 'Endre DG for oppheng';
  if (target === 'tapoff') return 'Endre DG for avtappingsbokser';
  if (target === 'special') return 'Endre DG for spesialelementer';
  return 'Endre DG for material';
}

function getCurrentDgRateByTarget(target){
  if (target === 'montasje') return currentMontasjeMarginRate;
  if (target === 'engineering') return currentEngineeringMarginRate;
  if (target === 'oppheng') return currentOpphengMarginRate;
  if (target === 'tapoff') return currentTapOffMarginRate;
  if (target === 'special') return currentTapOffMarginRate;
  return currentMarginRate;
}

function setCurrentDgRateByTarget(target, rate){
  if (target === 'montasje'){
    return setCurrentMontasjeMarginRate(rate);
  }
  if (target === 'engineering'){
    return setCurrentEngineeringMarginRate(rate);
  }
  if (target === 'oppheng'){
    return setCurrentOpphengMarginRate(rate);
  }
  if (target === 'tapoff'){
    return setCurrentTapOffMarginRate(rate);
  }
  if (target === 'special'){
    return setCurrentTapOffMarginRate(rate);
  }
  return setCurrentMarginRate(rate);
}

function openMarginModal(target = 'material'){
  currentDgModalTarget = target;
  const modal = $('marginModal');
  if (!modal) return;
  const inputEl = $('marginPercentInput');
  const errorEl = $('marginError');
  const titleEl = $('marginTitle');
  if (errorEl) errorEl.textContent = '';
  if (titleEl){
    titleEl.textContent = getDgModalTitleByTarget(currentDgModalTarget);
  }
  if (inputEl){
    inputEl.value = String(round2(getCurrentDgRateByTarget(currentDgModalTarget) * 100));
    inputEl.focus();
    const len = inputEl.value.length;
    try{
      inputEl.setSelectionRange(0, len);
    }catch(_err){}
  }
  modal.style.display = 'flex';
}

function closeMarginModal(){
  const modal = $('marginModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('marginError');
  if (errorEl) errorEl.textContent = '';
}

function submitMarginModal(){
  const inputEl = $('marginPercentInput');
  const errorEl = $('marginError');
  const parsed = Number(String(inputEl?.value ?? '').trim().replace(',','.'));
  if (!Number.isFinite(parsed)){
    if (errorEl) errorEl.textContent = 'Oppgi en gyldig DG i prosent.';
    if (inputEl) inputEl.focus();
    return;
  }
  const nextRate = parsed / 100;
  if (!Number.isFinite(nextRate) || nextRate < 0 || nextRate > MAX_MARGIN_RATE){
    if (errorEl) errorEl.textContent = 'DG må være mellom 0 og 95 %.';
    if (inputEl) inputEl.focus();
    return;
  }
  setCurrentDgRateByTarget(currentDgModalTarget, nextRate);
  if (lastCalc){
    if (currentDgModalTarget === 'tapoff' || currentDgModalTarget === 'special'){
      lastCalc.tapOffMarginRate = currentTapOffMarginRate;
      lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
      lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
      if (lastCalcInput){
        lastCalcInput.tapOffMarginRate = currentTapOffMarginRate;
      }
      if (lastEmailPayload?.inputs){
        lastEmailPayload.inputs.tapOffMarginRate = currentTapOffMarginRate;
      }
      if (lastEmailPayload?.totals){
        lastEmailPayload.totals.tapOffMarginRate = currentTapOffMarginRate;
        lastEmailPayload.totals.tapOffOfferTotal = lastCalc.tapOffOfferTotal;
        lastEmailPayload.totals.specialElementOfferTotal = lastCalc.specialElementOfferTotal;
      }
      updateSelectedAddonTotalUI();
    } else {
      recalcLastTotalsFromCurrentRates();
    }
  }
  closeMarginModal();
}

// --- market data ---
function setMarketStatus(message, isError){
  const statusEl = $('marketStatus');
  if (!statusEl) return;
  const tone =
    isError === true ? 'warning' :
    isError === false ? 'neutral' :
    isError;
  statusEl.textContent = message || '';
  statusEl.classList.toggle('error', tone === 'warning' && Boolean(message));
  statusEl.classList.toggle('ok', tone === 'success' && Boolean(message));
}

function formatMarketTimestamp(value){
  if (!value) return '--';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '--';
  try{
    return fmtTimestampNO.format(date);
  }catch(_err){
    return date.toLocaleString('no-NO');
  }
}

function getMarketSnapshotTimestamp(snapshot){
  return snapshot?.fetchedAt || snapshot?.updatedAt || null;
}

function formatDateKeyInTimezone(date, timeZone){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  try{
    return new Intl.DateTimeFormat('en-CA', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      timeZone: timeZone || undefined
    }).format(value);
  }catch(_err){
    return new Intl.DateTimeFormat('en-CA', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    }).format(value);
  }
}

function isTimestampTodayInTimezone(value, timeZone){
  const dateKey = formatDateKeyInTimezone(value, timeZone);
  if (!dateKey) return false;
  return dateKey === formatDateKeyInTimezone(new Date(), timeZone);
}

function getMarketFreshness(snapshot){
  const schedule = (snapshot && snapshot.schedule && typeof snapshot.schedule === 'object') ? snapshot.schedule : null;
  const timezone = typeof schedule?.timezone === 'string' ? schedule.timezone.trim() : '';
  const lastSuccessAt = schedule?.lastSuccessAt || schedule?.lastRunAt || getMarketSnapshotTimestamp(snapshot);
  const lastAttemptAt = schedule?.lastAttemptAt || '';
  const lastError = typeof schedule?.lastError === 'string' ? schedule.lastError.trim() : '';

  const successMs = Date.parse(lastSuccessAt || '');
  const attemptMs = Date.parse(lastAttemptAt || '');
  const hasFreshError =
    Boolean(lastError) &&
    (
      !Number.isFinite(successMs) ||
      !Number.isFinite(attemptMs) ||
      attemptMs >= successMs
    );

  const updatedToday = isTimestampTodayInTimezone(lastSuccessAt, timezone);
  const isHealthyToday = updatedToday && !hasFreshError;

  return {
    isHealthyToday,
    hasFreshError
  };
}

function pickMarketAluminium(snapshot){
  if (snapshot && snapshot.aluminium && typeof snapshot.aluminium === 'object'){
    return snapshot.aluminium;
  }
  if (snapshot && snapshot.metals && snapshot.metals.aluminium && typeof snapshot.metals.aluminium === 'object'){
    return snapshot.metals.aluminium;
  }
  return {};
}

function normalizeFxPoint(rawPoint, fallbackSource){
  if (rawPoint && typeof rawPoint === 'object'){
    const rate = Number(rawPoint.rate);
    return {
      rate: Number.isFinite(rate) ? rate : NaN,
      date: rawPoint.date || '',
      source: rawPoint.source || fallbackSource || '',
      changes: rawPoint.changes && typeof rawPoint.changes === 'object' ? rawPoint.changes : {}
    };
  }
  const numericRate = Number(rawPoint);
  return {
    rate: Number.isFinite(numericRate) ? numericRate : NaN,
    date: '',
    source: fallbackSource || ''
  };
}

function pickFxData(snapshot){
  const fx = (snapshot && snapshot.fx && typeof snapshot.fx === 'object') ? snapshot.fx : {};
  const fallbackSource = fx.source || '';
  return {
    usd: normalizeFxPoint(fx.usdNok, fallbackSource),
    eur: normalizeFxPoint(fx.eurNok, fallbackSource)
  };
}

function buildFxMetaText(point){
  const pieces = [];
  if (point?.source) pieces.push(point.source);
  if (point?.date) pieces.push(point.date);
  return pieces.join(' · ');
}

function formatMarketChangeValue(change){
  const percent = Number(change?.percent);
  if (!Number.isFinite(percent)) return '--';
  const rate = Number(change?.rate);
  const rateText = Number.isFinite(rate) ? fmtFxNO.format(rate) : '--';
  const arrow = percent > 0 ? '↑' : percent < 0 ? '↓' : '→';
  return `${rateText} ${arrow} ${fmtMarketPercentNO.format(Math.abs(percent))} %`;
}

function setMarketChangeValue(elementId, change){
  const el = $(elementId);
  if (!el) return;
  const percent = Number(change?.percent);
  const direction = percent > 0 ? 'up' : percent < 0 ? 'down' : 'flat';
  el.textContent = formatMarketChangeValue(change);
  el.classList.toggle('is-up', Number.isFinite(percent) && direction === 'up');
  el.classList.toggle('is-down', Number.isFinite(percent) && direction === 'down');
  el.classList.toggle('is-flat', !Number.isFinite(percent) || direction === 'flat');
}

function applyMarketSnapshot(snapshot){
  if (!snapshot) return;
  marketDataState.snapshot = snapshot;
  const aluminium = pickMarketAluminium(snapshot);
  const fx = pickFxData(snapshot);
  const freshness = getMarketFreshness(snapshot);

  const alEl = $('marketAlPrice');
  const alPrice = Number(aluminium.price);
  if (alEl){
    alEl.textContent = Number.isFinite(alPrice) ? fmtNO.format(alPrice) : '--';
  }
  const alMetaEl = $('marketAlMeta');
  if (alMetaEl){
    const pieces = [];
    if (aluminium.notation){
      pieces.push(aluminium.notation);
    } else {
      const currency = aluminium.currency || 'USD';
      const unit = aluminium.unit || 't';
      pieces.push(`${currency}/${unit}`);
    }
    if (aluminium.symbol){
      pieces.push(aluminium.symbol);
    } else if (aluminium.source){
      pieces.push(aluminium.source);
    }
    alMetaEl.textContent = pieces.filter(Boolean).join(' · ');
  }

  const usdRate = fx.usd.rate;
  const usdEl = $('marketUsdNok');
  if (usdEl){
    usdEl.textContent = Number.isFinite(usdRate) ? fmtFxNO.format(usdRate) : '--';
  }
  const usdMetaEl = $('marketUsdMeta');
  if (usdMetaEl){
    usdMetaEl.textContent = buildFxMetaText(fx.usd) || 'Ingen data';
  }
  setMarketChangeValue('marketUsdWeek', fx.usd.changes?.week);
  setMarketChangeValue('marketUsdMonth', fx.usd.changes?.month);

  const eurRate = fx.eur.rate;
  const eurEl = $('marketEurNok');
  if (eurEl){
    eurEl.textContent = Number.isFinite(eurRate) ? fmtFxNO.format(eurRate) : '--';
  }
  const eurMetaEl = $('marketEurMeta');
  if (eurMetaEl){
    eurMetaEl.textContent = buildFxMetaText(fx.eur) || 'Ingen data';
  }
  setMarketChangeValue('marketEurWeek', fx.eur.changes?.week);
  setMarketChangeValue('marketEurMonth', fx.eur.changes?.month);

  const updatedEl = $('marketUpdated');
  if (updatedEl){
    updatedEl.textContent = formatMarketTimestamp(getMarketSnapshotTimestamp(snapshot));
    updatedEl.classList.toggle('market-updated-ok', freshness.isHealthyToday);
    updatedEl.classList.toggle('market-updated-warning', !freshness.isHealthyToday);
  }
  const manualMode = snapshot.mode === 'static';
  if (manualMode){
    setMarketStatus(MARKET_STATUS_MANUAL, false);
  } else if (freshness.isHealthyToday){
    setMarketStatus('Oppdatert i dag', 'success');
  } else if (freshness.hasFreshError){
    setMarketStatus('Oppdatering feilet, prøver igjen automatisk', 'warning');
  } else {
    setMarketStatus(MARKET_STATUS_DEFAULT, false);
  }
  updateUsdRateFromMarket(snapshot);
}

async function fetchMarketSnapshot(){
  const staticFallbackUrl = buildStaticAssetUrl('data/market-data.json');
  const sources = [
    buildApiUrl('/api/market-data'),
    staticFallbackUrl
  ];
  let lastErr = null;
  for (const url of sources){
    try{
      const res = await fetch(url, { cache:'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const payload = await res.json();
      if (url === staticFallbackUrl && payload && typeof payload === 'object'){
        if (!payload.mode) payload.mode = 'static';
      }
      return payload;
    }catch(err){
      lastErr = err;
    }
  }
  throw lastErr || new Error('Ingen markedsdatakilde svarte');
}

async function refreshMarketData(){
  if (!$('marketTicker')) return;
  setMarketStatus('Oppdaterer...', false);
  try{
    const payload = await fetchMarketSnapshot();
    applyMarketSnapshot(payload);
  }catch(err){
    console.warn('Kunne ikke hente markedsdata', err);
    setMarketStatus('Kunne ikke hente markedsdata', true);
  }
}

function initMarketDataTicker(){
  if (!$('marketTicker')) return;
  refreshMarketData();
  if (marketTickerState.timerId){
    clearInterval(marketTickerState.timerId);
  }
  marketTickerState.timerId = window.setInterval(refreshMarketData, MARKET_REFRESH_INTERVAL_MS);
}

function loadAuthFromSession(){
  if (typeof sessionStorage === 'undefined') return;
  try{
    const stored = sessionStorage.getItem(AUTH_SESSION_KEY);
    if (!stored) return;
    const parsed = JSON.parse(stored);
    if (!parsed || parsed.loggedIn !== true) return;
    const username = normalizeUserEmail(parsed.username);
    const token = typeof parsed.token === 'string' ? parsed.token : '';
    const profile = parsed.profile && typeof parsed.profile === 'object' ? parsed.profile : null;
    const isAdmin = parsed.isAdmin === true || ADMIN_NAV_ALLOWED_EMAILS.includes(username);
    if (!hasValidUserEmail(username)) return;
    if (!token) return;
    authState = { loggedIn: true, username, token, profile, isAdmin };
  }catch(err){
    console.warn('Kunne ikke lese innloggingsstatus', err);
  }
}

function persistAuthToSession(){
  if (typeof sessionStorage === 'undefined') return;
  try{
    if (authState.loggedIn){
      sessionStorage.setItem(AUTH_SESSION_KEY, JSON.stringify({
        loggedIn: true,
        username: authState.username || '',
        token: authState.token || '',
        profile: authState.profile || null,
        isAdmin: authState.isAdmin === true
      }));
      return;
    }
    sessionStorage.removeItem(AUTH_SESSION_KEY);
  }catch(err){
    console.warn('Kunne ikke lagre innloggingsstatus', err);
  }
}

function canEditGlobalCustomerData(){
  if (!authState.loggedIn) return false;
  return authState.isAdmin === true || ADMIN_NAV_ALLOWED_EMAILS.includes(normalizeUserEmail(authState.username));
}

function canAccessProjectMailbox(){
  return canEditGlobalCustomerData();
}

function projectMailboxGraphPath(path){
  const suffix = String(path || '');
  return `/users/${encodeURIComponent(PROJECT_MAILBOX_ADDRESS)}${suffix.startsWith('/') ? suffix : `/${suffix}`}`;
}

function updateEmailMailboxAccessUi(){
  const canAccess = canAccessProjectMailbox();
  const composeBtn = $('composeEmailBtn');
  const refreshBtn = $('refreshEmailBtn');
  const list = $('emailMessagesList');

  if (composeBtn){
    composeBtn.hidden = !canAccess;
    composeBtn.disabled = !canAccess;
  }
  if (refreshBtn){
    refreshBtn.disabled = !canAccess;
  }
  if (!canAccess){
    closeEmailComposeForm();
    emailViewState.selectedMessageId = '';
    renderEmailMessages([]);
    if (list) delete list.dataset.loaded;
    setGraphStatus(
      'emailStatus',
      authState.loggedIn
        ? `E-post er kun synlig for Owners. ${PROJECT_MAILBOX_ADDRESS} vises ikke for denne brukeren.`
        : 'Logg inn med Microsoft for å vise e-post.',
      authState.loggedIn ? '' : 'error'
    );
  }
}

function updateAuthUI(){
  const calcBtn = $('calcBtn');
  if (calcBtn) calcBtn.disabled = !authState.loggedIn;

  const loginBtn = $('loginBtn');
  const logoutBtn = $('logoutBtn');
  const userLabel = $('authUser');
  const adminBtn = $('adminPageBtn');
  const dashboardAdminNav = $('dashboardAdminNav');
  const dashboardShell = $('dashboardShell');
  const canOpenAdmin = canEditGlobalCustomerData();

  document.body.classList.toggle('is-auth-locked', !authState.loggedIn);
  if (loginBtn) loginBtn.hidden = authState.loggedIn;
  if (logoutBtn) logoutBtn.hidden = !authState.loggedIn;
  if (dashboardShell) dashboardShell.classList.toggle('is-authenticated', authState.loggedIn);
  if (adminBtn) {
    adminBtn.hidden = !canOpenAdmin;
    adminBtn.setAttribute('aria-disabled', canOpenAdmin ? 'false' : 'true');
    adminBtn.tabIndex = canOpenAdmin ? 0 : -1;
  }
  if (dashboardAdminNav) {
    dashboardAdminNav.hidden = !canOpenAdmin;
    dashboardAdminNav.setAttribute('aria-disabled', canOpenAdmin ? 'false' : 'true');
    dashboardAdminNav.tabIndex = canOpenAdmin ? 0 : -1;
  }
  if (userLabel){
    if (authState.loggedIn){
      userLabel.textContent = authState.username || 'Innlogget';
      userLabel.hidden = false;
    } else {
      userLabel.textContent = '';
      userLabel.hidden = true;
    }
  }

  const statusEl = $('status');
  if (statusEl){
    const loginMsg = 'Logg inn for \u00E5 beregne.';
    if (!authState.loggedIn){
      if (!statusEl.textContent){
        statusEl.textContent = loginMsg;
      }
    } else if (statusEl.textContent === loginMsg){
      statusEl.textContent = '';
    }
  }

  const newProjectBtn = $('newProjectBtn');
  if (newProjectBtn){
    newProjectBtn.disabled = !authState.loggedIn || projectState.showArchive === true;
  }
  const refreshProjectsBtn = $('refreshProjectsBtn');
  if (refreshProjectsBtn){
    refreshProjectsBtn.disabled = !authState.loggedIn;
  }
  const createProjectButtons = Array.from(document.querySelectorAll('button[data-action="create-project"]'));
  createProjectButtons.forEach(btn=>{
    btn.disabled = !authState.loggedIn || projectState.showArchive === true;
  });
  updateEmailMailboxAccessUi();
  renderGlobalCustomerViews();
  renderProjectFlowView();
}

function authHeaders(){
  if (!authState.loggedIn || !authState.token) return {};
  return { Authorization: `Bearer ${authState.token}` };
}

function defaultMicrosoftRedirectUri(){
  const origin = window.location.origin;
  const path = window.location.pathname || '/';
  const repoBase = '/busbar-webapp-embedded-v2/';
  if (path.startsWith(repoBase)){
    return `${origin}${repoBase}auth/callback`;
  }
  return `${origin}/auth/callback`;
}

async function loadMicrosoftAuthConfig(){
  if (!microsoftAuthConfigPromise){
    const configUrl = buildApiUrl('/api/auth/microsoft/config');
    const url = new URL(configUrl, window.location.href);
    url.searchParams.set('origin', window.location.origin);
    url.searchParams.set('path', window.location.pathname || '/');
    microsoftAuthConfigPromise = fetch(url.toString(), {
      cache: 'no-store'
    })
      .then(async res => {
        let payload = null;
        try{
          payload = await res.json();
        }catch(_err){}
        if (!res.ok){
          throw new Error(payload?.error || appendApiBaseHint('Kunne ikke hente Microsoft-konfigurasjon.', res.status));
        }
        if (!payload?.enabled || !payload?.clientId || !payload?.tenantId){
          throw new Error('Microsoft-innlogging er ikke konfigurert på serveren.');
        }
        return {
          clientId: String(payload.clientId),
          tenantId: String(payload.tenantId),
          authority: String(payload.authority || `https://login.microsoftonline.com/${payload.tenantId}`),
          redirectUri: String(payload.redirectUri || defaultMicrosoftRedirectUri()),
          scopes: Array.isArray(payload.scopes) && payload.scopes.length
            ? payload.scopes.map(scope=>String(scope)).filter(Boolean)
            : [...MICROSOFT_AUTH_DEFAULT_SCOPES]
        };
      })
      .catch(err => {
        microsoftAuthConfigPromise = null;
        throw new Error(err?.message || `Kunne ikke hente Microsoft-konfigurasjon fra ${configUrl}.`);
      });
  }
  return microsoftAuthConfigPromise;
}

async function getMicrosoftMsalClient(){
  if (!window.msal?.PublicClientApplication){
    throw new Error('Microsoft innloggingsbibliotek ble ikke lastet.');
  }
  const config = await loadMicrosoftAuthConfig();
  if (!microsoftMsalClient){
    microsoftMsalClient = new window.msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: config.authority,
        redirectUri: config.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
      }
    });
    if (typeof microsoftMsalClient.initialize === 'function'){
      await microsoftMsalClient.initialize();
    }
  }
  return { client: microsoftMsalClient, config };
}

async function exchangeMicrosoftToken(idToken){
  const res = await fetch(buildApiUrl('/api/auth/microsoft/session'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ idToken })
  });
  let payload = null;
  try{
    payload = await res.json();
  }catch(_err){}
  if (!res.ok){
    throw new Error(payload?.error || appendApiBaseHint('Microsoft-innlogging ble avvist av serveren.', res.status));
  }
  const username = normalizeUserEmail(payload?.email);
  const token = typeof payload?.token === 'string' ? payload.token : '';
  if (!hasValidUserEmail(username) || !token){
    throw new Error('Serveren returnerte ugyldig Microsoft-innlogging.');
  }
  return { username, token, profile: normalizeProfilePayload(payload?.profile), isAdmin: payload?.isAdmin === true };
}

async function getMicrosoftAccountForGraph(){
  const { client } = await getMicrosoftMsalClient();
  const accounts = typeof client.getAllAccounts === 'function' ? client.getAllAccounts() : [];
  const currentEmail = normalizeUserEmail(authState.username);
  const matching = accounts.find(account=>normalizeUserEmail(account.username) === currentEmail)
    || accounts.find(account=>normalizeUserEmail(account.idTokenClaims?.preferred_username) === currentEmail)
    || microsoftLastAccount
    || accounts[0]
    || null;
  if (matching && typeof client.setActiveAccount === 'function'){
    client.setActiveAccount(matching);
  }
  return matching;
}

async function acquireMicrosoftGraphToken(scopes){
  if (!authState.loggedIn){
    throw new Error('Logg inn med Microsoft for å hente data.');
  }
  const { client } = await getMicrosoftMsalClient();
  const account = await getMicrosoftAccountForGraph();
  if (!account){
    throw new Error('Fant ikke aktiv Microsoft-konto. Logg inn med Microsoft på nytt.');
  }
  const request = {
    scopes: Array.isArray(scopes) ? scopes : [],
    account
  };
  try{
    const result = await client.acquireTokenSilent(request);
    return result?.accessToken || '';
  }catch(err){
    const needsInteraction = window.msal?.InteractionRequiredAuthError
      && err instanceof window.msal.InteractionRequiredAuthError;
    if (!needsInteraction && !String(err?.errorCode || '').includes('interaction_required')){
      throw err;
    }
    const result = await client.acquireTokenPopup(request);
    return result?.accessToken || '';
  }
}

async function microsoftGraphRequest(path, scopes, options = {}){
  const accessToken = await acquireMicrosoftGraphToken(scopes);
  if (!accessToken){
    throw new Error('Microsoft returnerte ikke access token.');
  }
  const headers = {
      Authorization: `Bearer ${accessToken}`,
      Prefer: 'outlook.timezone="Europe/Oslo"'
  };
  const request = {
    method: options.method || 'GET',
    headers
  };
  if (options.body !== undefined){
    if (options.rawBody){
      request.body = options.body;
      if (options.contentType) headers['Content-Type'] = options.contentType;
    } else {
      headers['Content-Type'] = 'application/json';
      request.body = JSON.stringify(options.body);
    }
  }
  const res = await fetch(`${MICROSOFT_GRAPH_BASE_URL}${path}`, request);
  let payload = null;
  try{
    payload = await res.json();
  }catch(_err){}
  if (!res.ok){
    throw new Error(payload?.error?.message || `Microsoft Graph svarte ${res.status}.`);
  }
  return payload;
}

function splitEmailAddresses(value){
  return String(value || '')
    .split(/[;,]/)
    .map(item=>item.trim())
    .filter(Boolean);
}

function graphEmailRecipients(value){
  return splitEmailAddresses(value).map(address=>({
    emailAddress: { address }
  }));
}

function isLikelyEmailAddress(value){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function graphLocalDateTimeValue(value){
  const raw = String(value || '').trim();
  return raw ? raw.replace('T', 'T') : '';
}

function formatDateTimeLocalInput(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  const pad = number=>String(number).padStart(2, '0');
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}T${pad(value.getHours())}:${pad(value.getMinutes())}`;
}

function formatDateInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  const pad = number=>String(number).padStart(2, '0');
  return `${pad(value.getDate())}/${pad(value.getMonth() + 1)}/${value.getFullYear()}`;
}

function formatIsoDateInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  const pad = number=>String(number).padStart(2, '0');
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
}

function formatTimeInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  const pad = number=>String(number).padStart(2, '0');
  return `${pad(value.getHours())}:${pad(value.getMinutes())}`;
}

function combineLocalDateAndTimeValue(dateValue, timeValue){
  const parsedDate = parseCalendarDateInputValue(dateValue);
  const time = String(timeValue || '').trim();
  return parsedDate && time ? `${parsedDate}T${time}` : '';
}

function parseCalendarDateInputValue(value){
  const raw = String(value || '').trim();
  const match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return '';
  const [, day, month, year] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  if (
    date.getFullYear() !== Number(year) ||
    date.getMonth() !== Number(month) - 1 ||
    date.getDate() !== Number(day)
  ) return '';
  return `${year}-${month}-${day}`;
}

function syncCalendarNativeDatePicker(){
  const textInput = $('calendarEventStartDate');
  const picker = $('calendarEventDatePicker');
  if (!textInput || !picker) return;
  const parsed = parseCalendarDateInputValue(textInput.value);
  picker.value = parsed || '';
}

function dateFromCalendarInputValue(value){
  const parsed = parseCalendarDateInputValue(value);
  return parsed ? new Date(`${parsed}T00:00`) : null;
}

function closeCalendarDatePickerPopover(){
  const popover = $('calendarDatePickerPopover');
  if (popover) popover.hidden = true;
}

function renderCalendarDatePickerPopover(){
  const popover = $('calendarDatePickerPopover');
  if (!popover) return;
  const selected = dateFromCalendarInputValue($('calendarEventStartDate')?.value);
  const cursor = calendarViewState.datePickerCursor instanceof Date
    ? calendarViewState.datePickerCursor
    : selected || new Date();
  const monthStart = startOfMonth(cursor);
  const gridStart = startOfWeekMonday(monthStart);
  const today = startOfDay(new Date());
  popover.innerHTML = '';

  const head = document.createElement('div');
  head.className = 'calendar-date-picker-head';
  const title = document.createElement('div');
  title.className = 'calendar-date-picker-title';
  title.textContent = new Intl.DateTimeFormat('no-NO', { month: 'long', year: 'numeric' }).format(cursor);
  const nav = document.createElement('div');
  nav.className = 'calendar-date-picker-nav';
  const prev = document.createElement('button');
  prev.type = 'button';
  prev.setAttribute('aria-label', 'Forrige måned');
  prev.textContent = '‹';
  prev.addEventListener('click', ()=>{
    calendarViewState.datePickerCursor = addMonths(cursor, -1);
    renderCalendarDatePickerPopover();
  });
  const next = document.createElement('button');
  next.type = 'button';
  next.setAttribute('aria-label', 'Neste måned');
  next.textContent = '›';
  next.addEventListener('click', ()=>{
    calendarViewState.datePickerCursor = addMonths(cursor, 1);
    renderCalendarDatePickerPopover();
  });
  nav.append(prev, next);
  head.append(title, nav);
  popover.appendChild(head);

  const grid = document.createElement('div');
  grid.className = 'calendar-date-picker-grid';
  ['Ma', 'Ti', 'On', 'To', 'Fr', 'Lø', 'Sø'].forEach(label=>{
    const weekday = document.createElement('div');
    weekday.className = 'calendar-date-picker-weekday';
    weekday.textContent = label;
    grid.appendChild(weekday);
  });
  for (let index = 0; index < 42; index += 1){
    const day = addDays(gridStart, index);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'calendar-date-picker-day';
    btn.classList.toggle('is-outside', day.getMonth() !== cursor.getMonth());
    btn.classList.toggle('is-today', sameCalendarDay(day, today));
    btn.classList.toggle('is-selected', selected ? sameCalendarDay(day, selected) : false);
    btn.textContent = String(day.getDate());
    btn.addEventListener('click', ()=>{
      const input = $('calendarEventStartDate');
      if (input) input.value = formatDateInputValue(day);
      syncCalendarNativeDatePicker();
      closeCalendarDatePickerPopover();
    });
    grid.appendChild(btn);
  }
  popover.appendChild(grid);
}

function openCalendarDatePickerPopover(){
  const popover = $('calendarDatePickerPopover');
  if (!popover) return;
  const selected = dateFromCalendarInputValue($('calendarEventStartDate')?.value);
  calendarViewState.datePickerCursor = selected || new Date();
  renderCalendarDatePickerPopover();
  popover.hidden = false;
}

function formatCalendarDurationOption(hours){
  const normalized = Number(hours);
  if (!Number.isFinite(normalized)) return '';
  return normalized % 1 === 0 ? `${normalized} timer` : `${String(normalized).replace('.', ',')} time`;
}

function populateCalendarDurationOptions(){
  const select = $('calendarEventDuration');
  if (!select || select.options.length) return;
  for (let value = 0.5; value <= 24; value += 0.5){
    const option = document.createElement('option');
    option.value = String(value);
    option.textContent = formatCalendarDurationOption(value);
    if (value === 1) option.selected = true;
    select.appendChild(option);
  }
}

function populateCalendarTimeOptions(){
  const select = $('calendarEventStartTime');
  if (!select || select.options.length) return;
  for (let hour = 0; hour < 24; hour += 1){
    for (let minute = 0; minute < 60; minute += 15){
      const value = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    }
  }
}

function calendarDurationFromDates(start, end){
  const startDate = start instanceof Date ? start : new Date(start);
  const endDate = end instanceof Date ? end : new Date(end);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return '1';
  const diffHours = (endDate.getTime() - startDate.getTime()) / 3600000;
  const rounded = Math.min(24, Math.max(0.5, Math.round(diffHours * 2) / 2));
  return String(rounded);
}

function formatGraphDateTime(raw){
  const value = String(raw?.dateTime || raw || '').trim();
  if (!value) return '';
  const date = parseGraphDate(raw);
  if (!date || Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('no-NO', {
    weekday: 'short',
    day: '2-digit',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

function parseGraphDateTimeValue(value, timeZone){
  const rawValue = String(value || '').trim();
  if (!rawValue) return null;
  const normalized = rawValue.replace(/(\.\d{3})\d+/, '$1');
  if (/[zZ]$|[+-]\d{2}:?\d{2}$/.test(normalized)){
    const date = new Date(normalized);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const tz = String(timeZone || '').trim().toLowerCase();
  if (tz === 'utc' || tz === 'etc/utc' || tz === 'gmt'){
    const date = new Date(`${normalized}Z`);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?/);
  if (match){
    const [, year, month, day, hour, minute, second = '0', ms = '0'] = match;
    return new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
      Number(ms.padEnd(3, '0'))
    );
  }
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseGraphDate(raw){
  const value = String(raw?.dateTime || raw || '').trim();
  if (!value) return null;
  return parseGraphDateTimeValue(value, raw?.timeZone);
}

function startOfDay(date){
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function addDays(date, days){
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function addMonths(date, months){
  const copy = new Date(date);
  copy.setMonth(copy.getMonth() + months);
  return copy;
}

function startOfWeekMonday(date){
  const copy = startOfDay(date);
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + diff);
  return copy;
}

function startOfMonth(date){
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date){
  return new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59, 999);
}

function getIsoWeekNumber(date){
  const copy = startOfDay(date);
  copy.setDate(copy.getDate() + 3 - ((copy.getDay() + 6) % 7));
  const weekOne = new Date(copy.getFullYear(), 0, 4);
  return 1 + Math.round(((copy - weekOne) / 86400000 - 3 + ((weekOne.getDay() + 6) % 7)) / 7);
}

function calendarRangeForState(){
  if (calendarViewState.mode === 'month'){
    const start = startOfWeekMonday(startOfMonth(calendarViewState.cursor));
    const end = addDays(startOfWeekMonday(endOfMonth(calendarViewState.cursor)), 7);
    return { start, end };
  }
  if (calendarViewState.mode === 'week'){
    const start = startOfWeekMonday(calendarViewState.cursor);
    return { start, end: addDays(start, 7) };
  }
  const start = startOfDay(new Date());
  return { start, end: addDays(start, 14) };
}

function formatCalendarPeriodLabel(){
  if (calendarViewState.mode === 'month'){
    return new Intl.DateTimeFormat('no-NO', { month: 'long', year: 'numeric' }).format(calendarViewState.cursor);
  }
  if (calendarViewState.mode === 'week'){
    const start = startOfWeekMonday(calendarViewState.cursor);
    const end = addDays(start, 6);
    return `${new Intl.DateTimeFormat('no-NO', { day: '2-digit', month: 'short' }).format(start)} - ${new Intl.DateTimeFormat('no-NO', { day: '2-digit', month: 'short', year: 'numeric' }).format(end)}`;
  }
  return 'Kommende 14 dager';
}

function setGraphStatus(id, message, state = ''){
  const el = $(id);
  if (!el) return;
  el.textContent = message || '';
  el.classList.toggle('error', state === 'error');
  el.classList.toggle('ok', state === 'ok');
}

function renderCalendarEvents(events){
  const list = $('calendarEventsList');
  if (!list) return;
  list.innerHTML = '';
  if (!events.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = 'Ingen kommende kalenderhendelser funnet.';
    list.appendChild(empty);
    return;
  }
  events.forEach(event=>{
    const item = document.createElement('article');
    item.className = 'graph-item calendar-event-item';

    const time = document.createElement('div');
    time.className = 'graph-item-time';
    time.textContent = `${formatGraphDateTime(event.start)} - ${formatGraphDateTime(event.end)}`;

    const body = document.createElement('div');
    body.className = 'graph-item-body';
    const title = document.createElement('h3');
    title.textContent = event.subject || 'Uten tittel';
    const meta = document.createElement('p');
    meta.className = 'muted-text';
    const organizer = event.organizer?.emailAddress?.name || event.organizer?.emailAddress?.address || '';
    const location = event.location?.displayName || '';
    meta.textContent = [organizer ? `Arrangør: ${organizer}` : '', location ? `Sted: ${location}` : ''].filter(Boolean).join(' | ');
    body.appendChild(title);
    body.appendChild(meta);
    const actions = document.createElement('div');
    actions.className = 'graph-card-actions';
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn alt btn-small';
    editBtn.textContent = 'Rediger';
    editBtn.dataset.editCalendarEvent = event.id || '';
    actions.appendChild(editBtn);
    body.appendChild(actions);

    item.appendChild(time);
    item.appendChild(body);
    list.appendChild(item);
  });
}

function sameCalendarDay(left, right){
  return left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate();
}

function renderCalendarGrid(events){
  const grid = $('calendarGridView');
  if (!grid) return;
  grid.innerHTML = '';
  const mode = calendarViewState.mode;
  const range = calendarRangeForState();
  const days = [];
  const dayCount = mode === 'month' ? 42 : 7;
  for (let index = 0; index < dayCount; index += 1){
    days.push(addDays(range.start, index));
  }

  const weekdays = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag', 'Søndag'];
  const header = document.createElement('div');
  header.className = 'calendar-grid-header';
  if (mode === 'month' || mode === 'week'){
    const weekHeader = document.createElement('div');
    weekHeader.className = 'calendar-week-header-spacer';
    weekHeader.textContent = '';
    header.appendChild(weekHeader);
  }
  weekdays.forEach(day=>{
    const cell = document.createElement('div');
    cell.textContent = day;
    header.appendChild(cell);
  });
  grid.appendChild(header);

  const body = document.createElement('div');
  body.className = `calendar-grid calendar-grid-${mode}`;
  const today = startOfDay(new Date());
  days.forEach((day, index)=>{
    if ((mode === 'month' || mode === 'week') && index % 7 === 0){
      const week = document.createElement('div');
      week.className = 'calendar-week-number';
      week.textContent = `UKE ${getIsoWeekNumber(day)}`;
      body.appendChild(week);
    }
    const cell = document.createElement('section');
    cell.className = 'calendar-day-cell';
    cell.classList.toggle('is-outside-month', mode === 'month' && day.getMonth() !== calendarViewState.cursor.getMonth());
    cell.classList.toggle('is-today', sameCalendarDay(day, today));
    cell.classList.toggle('is-past-day', day < today);

    const heading = document.createElement('div');
    heading.className = 'calendar-day-heading';
    heading.textContent = new Intl.DateTimeFormat('no-NO', { day: '2-digit', month: mode === 'week' ? 'short' : undefined }).format(day);
    cell.appendChild(heading);

    const dayEvents = events
      .filter(event=>{
        const start = parseGraphDate(event.start);
        return start && sameCalendarDay(start, day);
      })
      .sort((a, b)=>(parseGraphDate(a.start)?.getTime() || 0) - (parseGraphDate(b.start)?.getTime() || 0));

    if (!dayEvents.length){
      const empty = document.createElement('div');
      empty.className = 'calendar-day-empty';
      empty.textContent = '';
      cell.appendChild(empty);
    }

    dayEvents.forEach(event=>{
      const item = document.createElement('button');
      item.className = 'calendar-grid-event';
      item.type = 'button';
      item.dataset.editCalendarEvent = event.id || '';
      const time = parseGraphDate(event.start);
      const timeLabel = time ? new Intl.DateTimeFormat('no-NO', { hour: '2-digit', minute: '2-digit' }).format(time) : '';
      item.textContent = [timeLabel, event.subject || 'Uten tittel'].filter(Boolean).join(' ');
      cell.appendChild(item);
    });

    body.appendChild(cell);
  });
  grid.appendChild(body);
}

function getCalendarContactEmailSuggestions(){
  const suggestions = [];
  const seen = new Set();
  flattenGlobalContacts().forEach(contact=>{
    const address = String(contact?.email || '').trim();
    if (!address) return;
    const key = address.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    suggestions.push({
      address,
      label: [contact.name, contact.customerName].filter(Boolean).join(' - ')
    });
  });
  return suggestions;
}

function renderCalendarAttendeeOptions(){
  const list = $('calendarAttendeeOptions');
  if (!list) return;
  list.innerHTML = '';
  getCalendarContactEmailSuggestions().forEach(item=>{
    const option = document.createElement('option');
    option.value = item.address;
    option.label = item.label || item.address;
    list.appendChild(option);
  });
}

async function ensureCalendarAttendeeSuggestions(){
  if (!authState.loggedIn || projectState.globalCustomerDatabaseLoaded){
    renderCalendarAttendeeOptions();
    return;
  }
  try{
    await loadGlobalCustomerDatabase({ silent: true });
  }catch(_err){
    renderCalendarAttendeeOptions();
  }
}

function setCalendarFormAttendees(attendees){
  const seen = new Set();
  calendarViewState.formAttendees = (Array.isArray(attendees) ? attendees : [])
    .map(value=>String(value || '').trim())
    .filter(Boolean)
    .filter(address=>{
      const key = address.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  renderCalendarAttendeesList();
}

function addCalendarFormAttendees(value){
  const next = [
    ...calendarViewState.formAttendees,
    ...splitEmailAddresses(value)
  ];
  setCalendarFormAttendees(next);
  const input = $('calendarAttendeeInput');
  if (input) input.value = '';
}

function renderCalendarAttendeesList(){
  const list = $('calendarAttendeesList');
  if (!list) return;
  list.innerHTML = '';
  if (!calendarViewState.formAttendees.length){
    const empty = document.createElement('span');
    empty.className = 'calendar-attendee-empty';
    empty.textContent = 'Ingen mottakere lagt til';
    list.appendChild(empty);
    return;
  }
  calendarViewState.formAttendees.forEach(address=>{
    const item = document.createElement('span');
    item.className = 'calendar-attendee-chip';
    const text = document.createElement('span');
    text.textContent = address;
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.setAttribute('aria-label', `Fjern ${address}`);
    remove.textContent = '×';
    remove.addEventListener('click', ()=>{
      setCalendarFormAttendees(calendarViewState.formAttendees.filter(itemAddress=>itemAddress !== address));
    });
    item.append(text, remove);
    list.appendChild(item);
  });
}

function resetCalendarEventForm(){
  const form = $('calendarEventForm');
  if (!form) return;
  form.reset();
  const idEl = $('calendarEventId');
  if (idEl) idEl.value = '';
  const deleteBtn = $('deleteCalendarEventBtn');
  if (deleteBtn) deleteBtn.hidden = true;
  const start = new Date();
  start.setMinutes(0, 0, 0);
  start.setHours(start.getHours() + 1);
  const end = new Date(start);
  end.setHours(end.getHours() + 1);
  const startDateEl = $('calendarEventStartDate');
  const startTimeEl = $('calendarEventStartTime');
  const durationEl = $('calendarEventDuration');
  if (startDateEl) startDateEl.value = formatDateInputValue(start);
  if (startTimeEl) startTimeEl.value = formatTimeInputValue(start);
  if (durationEl) durationEl.value = calendarDurationFromDates(start, end);
  const nativeDatePicker = $('calendarEventDatePicker');
  if (nativeDatePicker) nativeDatePicker.value = formatIsoDateInputValue(start);
  setCalendarFormAttendees([]);
}

function openCalendarEventForm(event = null){
  const form = $('calendarEventForm');
  if (!form) return;
  resetCalendarEventForm();
  const id = String(event?.id || '').trim();
  const idEl = $('calendarEventId');
  if (idEl) idEl.value = id;
  const subjectEl = $('calendarEventSubject');
  const locationEl = $('calendarEventLocation');
  const startDateEl = $('calendarEventStartDate');
  const startTimeEl = $('calendarEventStartTime');
  const durationEl = $('calendarEventDuration');
  const bodyEl = $('calendarEventBody');
  if (subjectEl) subjectEl.value = event?.subject || '';
  if (locationEl) locationEl.value = event?.location?.displayName || '';
  const start = event?.start ? parseGraphDate(event.start) : null;
  const end = event?.end ? parseGraphDate(event.end) : null;
  if (startDateEl && start) startDateEl.value = formatDateInputValue(start);
  if (startTimeEl && start) startTimeEl.value = formatTimeInputValue(start);
  if (durationEl && start && end) durationEl.value = calendarDurationFromDates(start, end);
  const nativeDatePicker = $('calendarEventDatePicker');
  if (nativeDatePicker && start) nativeDatePicker.value = formatIsoDateInputValue(start);
  setCalendarFormAttendees((Array.isArray(event?.attendees) ? event.attendees : [])
    .map(attendee=>attendee?.emailAddress?.address)
    .filter(Boolean));
  if (bodyEl) bodyEl.value = event?.bodyPreview || '';
  const deleteBtn = $('deleteCalendarEventBtn');
  if (deleteBtn) deleteBtn.hidden = !id;
  form.hidden = false;
  openFormModal('calendarEventForm', id ? 'Endre avtale' : 'Ny avtale');
  void ensureCalendarAttendeeSuggestions();
  if (subjectEl) subjectEl.focus();
}

function closeCalendarEventForm(){
  closeFormModal('calendarEventForm');
}

function getCalendarEventPayloadFromForm(){
  const subject = String($('calendarEventSubject')?.value || '').trim();
  const location = String($('calendarEventLocation')?.value || '').trim();
  const start = graphLocalDateTimeValue(combineLocalDateAndTimeValue($('calendarEventStartDate')?.value, $('calendarEventStartTime')?.value));
  const durationHours = Number($('calendarEventDuration')?.value || 0);
  const body = String($('calendarEventBody')?.value || '').trim();
  const attendeeInput = $('calendarAttendeeInput');
  const attendees = [
    ...calendarViewState.formAttendees,
    ...splitEmailAddresses(attendeeInput?.value)
  ].filter((address, index, array)=>array.findIndex(item=>item.toLowerCase() === address.toLowerCase()) === index);
  const invalidAttendee = attendees.find(address=>!isLikelyEmailAddress(address));
  if (!parseCalendarDateInputValue($('calendarEventStartDate')?.value)) throw new Error('Dato må skrives som DD/MM/ÅÅÅÅ.');
  if (!subject || !start || !durationHours) throw new Error('Fyll inn emne, startdato, starttid og varighet.');
  if (invalidAttendee) throw new Error(`Ugyldig e-postadresse: ${invalidAttendee}`);
  if (durationHours < 0.5 || durationHours > 24) throw new Error('Varighet må være mellom 0,5 og 24 timer.');
  const startDate = new Date(start);
  if (Number.isNaN(startDate.getTime())) throw new Error('Starttidspunktet er ugyldig.');
  const endDate = new Date(startDate.getTime() + durationHours * 3600000);
  const end = formatDateTimeLocalInput(endDate);
  return {
    subject,
    start: { dateTime: start, timeZone: 'Europe/Oslo' },
    end: { dateTime: end, timeZone: 'Europe/Oslo' },
    location: { displayName: location },
    body: { contentType: 'text', content: body },
    attendees: attendees.map(address=>({
      emailAddress: { address },
      type: 'required'
    }))
  };
}

async function saveCalendarEventFromForm(){
  const form = $('calendarEventForm');
  if (!form) return;
  const submitBtn = form.querySelector('button[type="submit"]');
  const id = String($('calendarEventId')?.value || '').trim();
  if (submitBtn) submitBtn.disabled = true;
  setGraphStatus('calendarStatus', id ? 'Lagrer avtale...' : 'Oppretter avtale...');
  try{
    const body = getCalendarEventPayloadFromForm();
    if (id){
      await microsoftGraphRequest(`/me/events/${encodeURIComponent(id)}`, MICROSOFT_GRAPH_CALENDAR_SCOPES, {
        method: 'PATCH',
        body
      });
    } else {
      await microsoftGraphRequest('/me/events', MICROSOFT_GRAPH_CALENDAR_SCOPES, {
        method: 'POST',
        body
      });
    }
    closeCalendarEventForm();
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    await loadCalendarEvents();
  }catch(err){
    console.warn('Kalenderlagring feilet', err);
    setGraphStatus('calendarStatus', err?.message || 'Kunne ikke lagre avtalen.', 'error');
  }finally{
    if (submitBtn) submitBtn.disabled = false;
  }
}

async function deleteCalendarEventFromForm(){
  const id = String($('calendarEventId')?.value || '').trim();
  if (!id) return;
  if (!window.confirm('Slette denne kalenderhendelsen?')) return;
  const btn = $('deleteCalendarEventBtn');
  if (btn) btn.disabled = true;
  setGraphStatus('calendarStatus', 'Sletter avtale...');
  try{
    await microsoftGraphRequest(`/me/events/${encodeURIComponent(id)}`, MICROSOFT_GRAPH_CALENDAR_SCOPES, {
      method: 'DELETE'
    });
    closeCalendarEventForm();
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    await loadCalendarEvents();
  }catch(err){
    console.warn('Kalendersletting feilet', err);
    setGraphStatus('calendarStatus', err?.message || 'Kunne ikke slette avtalen.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}

function updateCalendarViewControls(){
  const label = $('calendarPeriodLabel');
  if (label) label.textContent = formatCalendarPeriodLabel();
  const select = $('calendarModeSelect');
  if (select) select.value = calendarViewState.mode;
}

function renderCalendarView(){
  const list = $('calendarEventsList');
  const grid = $('calendarGridView');
  const isList = calendarViewState.mode === 'list';
  if (list) list.hidden = !isList;
  if (grid) grid.hidden = isList;
  updateCalendarViewControls();
  if (isList){
    renderCalendarEvents(calendarViewState.events);
  } else {
    renderCalendarGrid(calendarViewState.events);
  }
}

function renderEmailMessages(messages){
  const list = $('emailMessagesList');
  if (!list) return;
  list.innerHTML = '';
  emailViewState.messages = Array.isArray(messages) ? messages : [];
  if (emailViewState.selectedMessageId && !emailViewState.messages.some(message=>message.id === emailViewState.selectedMessageId)){
    emailViewState.selectedMessageId = '';
  }
  updateEmailMessageActions();
  if (!messages.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = 'Ingen e-poster funnet.';
    list.appendChild(empty);
    return;
  }
  messages.forEach(message=>{
    const item = document.createElement('button');
    item.type = 'button';
    item.className = `graph-item email-message-item${message.isRead ? '' : ' is-unread'}${message.id === emailViewState.selectedMessageId ? ' is-selected' : ''}`;
    item.dataset.emailMessageId = message.id || '';

    const marker = document.createElement('div');
    marker.className = 'email-read-marker';
    marker.setAttribute('aria-hidden', 'true');

    const body = document.createElement('div');
    body.className = 'graph-item-body';
    const title = document.createElement('h3');
    title.textContent = message.subject || '(Uten emne)';
    const meta = document.createElement('p');
    meta.className = 'muted-text';
    const from = message.from?.emailAddress?.name || message.from?.emailAddress?.address || 'Ukjent avsender';
    meta.textContent = `${from} | ${formatGraphDateTime(message.receivedDateTime)}`;
    const preview = document.createElement('p');
    preview.className = 'email-preview';
    preview.textContent = message.bodyPreview || '';
    body.appendChild(title);
    body.appendChild(meta);
    body.appendChild(preview);

    item.appendChild(marker);
    item.appendChild(body);
    list.appendChild(item);
  });
}

function getSelectedEmailMessage(){
  const id = String(emailViewState.selectedMessageId || '').trim();
  if (!id) return null;
  return emailViewState.messages.find(message=>message.id === id) || null;
}

function updateEmailMessageActions(){
  const actions = $('emailMessageActions');
  if (!actions) return;
  const message = getSelectedEmailMessage();
  actions.hidden = !message;
  const label = $('selectedEmailLabel');
  if (label){
    label.textContent = message ? `Valgt: ${message.subject || '(Uten emne)'}` : '';
  }
  const markBtn = $('markSelectedEmailReadBtn');
  if (markBtn){
    markBtn.textContent = message?.isRead ? 'Marker ulest' : 'Marker lest';
  }
}

function selectEmailMessage(id){
  const nextId = String(id || '').trim();
  emailViewState.selectedMessageId = emailViewState.selectedMessageId === nextId ? '' : nextId;
  renderEmailMessages(emailViewState.messages);
}

function openEmailComposeForm(){
  if (!canAccessProjectMailbox()){
    setGraphStatus('emailStatus', 'E-post er kun tilgjengelig for Owners.', 'error');
    return;
  }
  const form = $('emailComposeForm');
  if (!form) return;
  form.reset();
  form.hidden = false;
  openFormModal('emailComposeForm', 'Ny e-post');
  const input = $('emailToInput');
  if (input) input.focus();
}

function closeEmailComposeForm(){
  closeFormModal('emailComposeForm');
}

function getEmailComposePayload(){
  const toRecipients = graphEmailRecipients($('emailToInput')?.value);
  const ccRecipients = graphEmailRecipients($('emailCcInput')?.value);
  const subject = String($('emailSubjectInput')?.value || '').trim();
  const content = String($('emailBodyInput')?.value || '').trim();
  if (!toRecipients.length) throw new Error('Legg inn minst en mottaker.');
  if (!subject) throw new Error('Legg inn emne.');
  if (!content) throw new Error('Legg inn melding.');
  return {
    message: {
      subject,
      body: { contentType: 'Text', content },
      toRecipients,
      ccRecipients
    },
    saveToSentItems: true
  };
}

async function sendEmailFromForm(){
  if (!canAccessProjectMailbox()){
    setGraphStatus('emailStatus', 'E-post er kun tilgjengelig for Owners.', 'error');
    return;
  }
  const form = $('emailComposeForm');
  if (!form) return;
  const submitBtn = form.querySelector('button[type="submit"]');
  if (submitBtn) submitBtn.disabled = true;
  setGraphStatus('emailStatus', 'Sender e-post...');
  try{
    await microsoftGraphRequest(projectMailboxGraphPath('/sendMail'), MICROSOFT_GRAPH_MAIL_SCOPES, {
      method: 'POST',
      body: getEmailComposePayload()
    });
    closeEmailComposeForm();
    await loadEmailMessages();
    setGraphStatus('emailStatus', 'E-post sendt.', 'ok');
  }catch(err){
    console.warn('Sending av e-post feilet', err);
    setGraphStatus('emailStatus', err?.message || 'Kunne ikke sende e-post.', 'error');
  }finally{
    if (submitBtn) submitBtn.disabled = false;
  }
}

async function markSelectedEmailRead(){
  if (!canAccessProjectMailbox()){
    setGraphStatus('emailStatus', 'E-post er kun tilgjengelig for Owners.', 'error');
    return;
  }
  const message = getSelectedEmailMessage();
  if (!message?.id) return;
  const nextIsRead = !message.isRead;
  setGraphStatus('emailStatus', nextIsRead ? 'Markerer e-post som lest...' : 'Markerer e-post som ulest...');
  try{
    await microsoftGraphRequest(projectMailboxGraphPath(`/messages/${encodeURIComponent(message.id)}`), MICROSOFT_GRAPH_MAIL_SCOPES, {
      method: 'PATCH',
      body: { isRead: nextIsRead }
    });
    await loadEmailMessages();
  }catch(err){
    console.warn('Markering av e-post feilet', err);
    setGraphStatus('emailStatus', err?.message || 'Kunne ikke endre lest-status.', 'error');
  }
}

async function deleteSelectedEmail(){
  if (!canAccessProjectMailbox()){
    setGraphStatus('emailStatus', 'E-post er kun tilgjengelig for Owners.', 'error');
    return;
  }
  const message = getSelectedEmailMessage();
  if (!message?.id) return;
  if (!window.confirm('Slette valgt e-post?')) return;
  setGraphStatus('emailStatus', 'Sletter e-post...');
  try{
    await microsoftGraphRequest(projectMailboxGraphPath(`/messages/${encodeURIComponent(message.id)}`), MICROSOFT_GRAPH_MAIL_SCOPES, {
      method: 'DELETE'
    });
    emailViewState.selectedMessageId = '';
    await loadEmailMessages();
  }catch(err){
    console.warn('Sletting av e-post feilet', err);
    setGraphStatus('emailStatus', err?.message || 'Kunne ikke slette e-post.', 'error');
  }
}

function formatSharePointFileSize(size){
  const bytes = Number(size);
  if (!Number.isFinite(bytes) || bytes <= 0) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} kB`;
  return `${(bytes / (1024 * 1024)).toLocaleString('no-NO', { maximumFractionDigits: 1 })} MB`;
}

function getSharePointListState(page){
  if (!sharePointFolderState[page]) sharePointFolderState[page] = {};
  return sharePointFolderState[page];
}

function normalizeListSearchText(value){
  return String(value || '').trim().toLowerCase();
}

function getSharePointSortMode(page){
  const state = getSharePointListState(page);
  return PROJECT_SORT_OPTIONS.includes(state.sort) ? state.sort : 'alpha_asc';
}

function sharePointItemMatchesSearch(item, page){
  const term = normalizeListSearchText(getSharePointListState(page).searchTerm);
  if (!term) return true;
  const haystack = [
    item?.name,
    item?.folder ? 'mappe' : 'fil'
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

function compareSharePointItems(a, b, page){
  const folderCmp = Number(Boolean(b?.folder)) - Number(Boolean(a?.folder));
  if (folderCmp) return folderCmp;
  const mode = getSharePointSortMode(page);
  if (mode === 'alpha_desc'){
    return compareNoText(b?.name, a?.name);
  }
  if (mode === 'date_newest' || mode === 'date_oldest'){
    const aTime = new Date(a?.lastModifiedDateTime || 0).getTime() || 0;
    const bTime = new Date(b?.lastModifiedDateTime || 0).getTime() || 0;
    return mode === 'date_oldest' ? aTime - bTime : bTime - aTime;
  }
  return compareNoText(a?.name, b?.name);
}

function renderSharePointFolderItems(config, items, page){
  const list = $(config.listId);
  if (!list) return;
  list.innerHTML = '';
  const sortedItems = [...items]
    .filter(item=>sharePointItemMatchesSearch(item, page))
    .sort((a, b)=>compareSharePointItems(a, b, page));
  if (!sortedItems.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = normalizeListSearchText(getSharePointListState(page).searchTerm)
      ? 'Ingen filer eller mapper matcher søket.'
      : 'Ingen filer eller mapper funnet.';
    list.appendChild(empty);
    return;
  }
  sortedItems.forEach(item=>{
    const row = document.createElement('article');
    row.className = `sharepoint-item${item.folder ? ' is-folder' : ' is-file'}`;

    const link = document.createElement('a');
    link.className = 'sharepoint-item-link';
    link.href = item.webUrl || '#';
    link.target = '_blank';
    link.rel = 'noopener noreferrer';

    const icon = document.createElement('span');
    icon.className = 'sharepoint-item-icon';
    icon.textContent = item.folder ? 'Mappe' : 'Fil';

    const body = document.createElement('span');
    body.className = 'sharepoint-item-body';
    const name = document.createElement('strong');
    name.textContent = item.name || 'Uten navn';
    const meta = document.createElement('span');
    meta.className = 'muted-text';
    const modified = formatTimestampForSharePoint(item.lastModifiedDateTime);
    const size = item.file ? formatSharePointFileSize(item.size) : '';
    meta.textContent = [item.folder ? `${Number(item.folder.childCount || 0)} elementer` : size, modified ? `Endret ${modified}` : ''].filter(Boolean).join(' | ');
    body.append(name, meta);

    link.append(icon, body);
    row.appendChild(link);

    const actions = document.createElement('div');
    actions.className = 'sharepoint-item-actions';
    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'btn danger btn-small';
    deleteBtn.textContent = 'Slett';
    deleteBtn.dataset.deleteSharepointItem = item.id || '';
    deleteBtn.dataset.sharepointPage = page || '';
    deleteBtn.dataset.sharepointName = item.name || '';
    actions.appendChild(deleteBtn);
    row.appendChild(actions);

    list.appendChild(row);
  });
}

function formatTimestampForSharePoint(value){
  const raw = String(value || '').trim();
  if (!raw) return '';
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return '';
  return new Intl.DateTimeFormat('no-NO', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

async function resolveSharePointDrive(config){
  const sitePath = String(config.sitePath || '').replace(/^\/+/, '');
  const sitePayload = await microsoftGraphRequest(
    `/sites/${encodeURIComponent(config.siteHost)}:/${sitePath}`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES
  );
  const siteId = String(sitePayload?.id || '').trim();
  if (!siteId) throw new Error(`Fant ikke SharePoint-site for ${config.title}.`);

  const drivesPayload = await microsoftGraphRequest(
    `/sites/${encodeURIComponent(siteId)}/drives?$select=id,name,webUrl`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES
  );
  const drives = Array.isArray(drivesPayload?.value) ? drivesPayload.value : [];
  const drive = drives.find(item=>{
    const name = String(item?.name || '').trim().toLowerCase();
    const webUrl = String(item?.webUrl || '').toLowerCase();
    return ['delte dokumenter', 'shared documents', 'documents', 'dokumenter'].includes(name)
      || webUrl.includes('/delte%20dokumenter')
      || webUrl.includes('/shared%20documents');
  })
    || drives[0]
    || null;
  if (!drive?.id) throw new Error(`Fant ikke dokumentbibliotek for ${config.title}.`);
  return drive;
}

async function loadSharePointFolder(page, options = {}){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config) return;
  if (!authState.loggedIn){
    setGraphStatus(config.statusId, 'Logg inn med Microsoft for å vise SharePoint-mappen.', 'error');
    return;
  }
  const list = $(config.listId);
  if (!list) return;
  if (options.silent && list.dataset.loaded === '1') return;
  const btn = $(config.refreshBtnId);
  if (btn) btn.disabled = true;
  setGraphStatus(config.statusId, 'Henter SharePoint-mappe...');
  try{
    const drive = await resolveSharePointDrive(config);
    sharePointFolderState[page] = {
      ...getSharePointListState(page),
      driveId: drive.id,
      folderPath: String(config.folderPath || '')
    };
    const folderPath = String(config.folderPath || '').split('/').map(encodeURIComponent).join('/');
    const query = new URLSearchParams({
      '$select': 'id,name,webUrl,folder,file,size,lastModifiedDateTime',
      '$top': '200'
    });
    const payload = await microsoftGraphRequest(
      `/drives/${encodeURIComponent(drive.id)}/root:/${folderPath}:/children?${query.toString()}`,
      MICROSOFT_GRAPH_SHAREPOINT_SCOPES
    );
    const items = Array.isArray(payload?.value) ? payload.value : [];
    getSharePointListState(page).items = items;
    renderSharePointFolderItems(config, items, page);
    list.dataset.loaded = '1';
    setGraphStatus(config.statusId, `Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`, 'ok');
  }catch(err){
    console.warn('SharePoint-henting feilet', err);
    setGraphStatus(config.statusId, err?.message || 'Kunne ikke hente SharePoint-mappen.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}

async function ensureSharePointFolderState(page){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config) throw new Error('Ukjent SharePoint-mappe.');
  if (!sharePointFolderState[page]?.driveId){
    const drive = await resolveSharePointDrive(config);
    sharePointFolderState[page] = {
      ...getSharePointListState(page),
      driveId: drive.id,
      folderPath: String(config.folderPath || '')
    };
  }
  return sharePointFolderState[page];
}

async function createSharePointFolder(page){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config) return;
  const name = window.prompt(`Navn på ny mappe i ${config.title}:`);
  const folderName = String(name || '').trim();
  if (!folderName) return;
  setGraphStatus(config.statusId, 'Oppretter mappe...');
  try{
    const state = await ensureSharePointFolderState(page);
    const folderPath = String(state.folderPath || '').split('/').map(encodeURIComponent).join('/');
    await microsoftGraphRequest(
      `/drives/${encodeURIComponent(state.driveId)}/root:/${folderPath}:/children`,
      MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
      {
        method: 'POST',
        body: {
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename'
        }
      }
    );
    await loadSharePointFolder(page);
  }catch(err){
    console.warn('Oppretting av SharePoint-mappe feilet', err);
    setGraphStatus(config.statusId, err?.message || 'Kunne ikke opprette mappe.', 'error');
  }
}

function sanitizeSharePointFolderName(value, fallback = 'Prosjekt'){
  const raw = String(value || '').trim() || fallback;
  return raw
    .replace(/[~"#%&*:<>?/\\{|}\u0000-\u001F]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim()
    || fallback;
}

function formatProjectFolderName(project){
  const existingName = String(project?.projectFolderName || '').trim();
  if (existingName) return sanitizeSharePointFolderName(existingName);
  const number = String(project?.projectNumber || '').trim();
  const name = String(project?.name || '').trim() || 'Uten navn';
  return sanitizeSharePointFolderName([number, name].filter(Boolean).join(' - '));
}

function sharePointNameEquals(a, b){
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

async function getSharePointFolderChildren(driveId, itemId){
  const query = new URLSearchParams({
    '$select': 'id,name,webUrl,folder,file,size,lastModifiedDateTime',
    '$top': '200'
  });
  const payload = await microsoftGraphRequest(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?${query.toString()}`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES
  );
  return Array.isArray(payload?.value) ? payload.value : [];
}

async function createSharePointChildFolder(driveId, parentItemId, folderName, conflictBehavior = 'fail'){
  return microsoftGraphRequest(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentItemId)}/children`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
    {
      method: 'POST',
      body: {
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': conflictBehavior
      }
    }
  );
}

async function copySharePointFileToFolder(driveId, sourceItemId, targetFolderId, name){
  await microsoftGraphRequest(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(sourceItemId)}/copy`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
    {
      method: 'POST',
      body: {
        parentReference: { driveId, id: targetFolderId },
        name
      }
    }
  );
}

async function cloneSharePointFolderContents(driveId, sourceFolderId, targetFolderId){
  const children = await getSharePointFolderChildren(driveId, sourceFolderId);
  for (const child of children){
    const name = String(child?.name || '').trim();
    if (!name || !child?.id) continue;
    if (child.folder){
      const createdFolder = await createSharePointChildFolder(driveId, targetFolderId, name, 'rename');
      if (createdFolder?.id){
        await cloneSharePointFolderContents(driveId, child.id, createdFolder.id);
      }
    } else if (child.file){
      await copySharePointFileToFolder(driveId, child.id, targetFolderId, name);
    }
  }
}

async function getSharePointRootFolderItem(page){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config) throw new Error('Ukjent SharePoint-mappe.');
  const state = await ensureSharePointFolderState(page);
  const folderPath = String(state.folderPath || '').split('/').map(encodeURIComponent).join('/');
  const root = await microsoftGraphRequest(
    `/drives/${encodeURIComponent(state.driveId)}/root:/${folderPath}`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES
  );
  if (!root?.id) throw new Error(`Fant ikke rotmappen for ${config.title}.`);
  return { ...state, rootItem: root };
}

async function findProjectSharePointFolder(project){
  const state = await getSharePointRootFolderItem('project-folders');
  const folderName = formatProjectFolderName(project);
  const children = await getSharePointFolderChildren(state.driveId, state.rootItem.id);
  const folder = children.find(item=>item?.folder && sharePointNameEquals(item.name, folderName));
  return { ...state, folderName, folder: folder || null };
}

function getProjectFolderStatus(project){
  const id = String(project?.id || '').trim();
  return id ? projectFolderStatusState.byProjectId[id] || null : null;
}

function projectHasConfirmedFolder(project){
  return getProjectFolderStatus(project)?.exists === true;
}

async function refreshProjectFolderStatuses(options = {}){
  if (!authState.loggedIn || projectFolderStatusState.loading) return;
  const projects = Array.isArray(projectState.projects) ? projectState.projects : [];
  if (!projects.length) return;
  projectFolderStatusState.loading = true;
  try{
    const state = await getSharePointRootFolderItem('project-folders');
    const children = await getSharePointFolderChildren(state.driveId, state.rootItem.id);
    const folders = children.filter(item=>item?.folder);
    const next = {};
    let changedProjects = false;
    projects.forEach(project=>{
      const expectedName = formatProjectFolderName(project);
      const folder = folders.find(item=>sharePointNameEquals(item.name, expectedName));
      next[project.id] = {
        exists: Boolean(folder?.id),
        folderName: folder?.name || expectedName,
        webUrl: folder?.webUrl || ''
      };
      if (folder?.id){
        if (project.projectFolderName !== folder.name || project.projectFolderWebUrl !== folder.webUrl || project.projectFolderCreated !== true){
          project.projectFolderName = String(folder.name || expectedName).trim();
          project.projectFolderCreated = true;
          project.projectFolderWebUrl = String(folder.webUrl || '').trim();
          changedProjects = true;
        }
      }
    });
    projectFolderStatusState.byProjectId = next;
    projectFolderStatusState.loaded = true;
    if (changedProjects) saveProjectsToStorage();
    if (options.render !== false) renderProjectDashboard();
  }catch(err){
    console.warn('Kunne ikke hente prosjektmappestatus', err);
    projectFolderStatusState.loaded = true;
  }finally{
    projectFolderStatusState.loading = false;
  }
}

function ensureProjectFolderStatusesLoaded(){
  if (!authState.loggedIn || projectFolderStatusState.loaded || projectFolderStatusState.loading) return;
  void refreshProjectFolderStatuses();
}

async function ensureProjectOfferSharePointFolder(project){
  const state = await findProjectSharePointFolder(project);
  if (!state.folder?.id){
    throw new Error(`Prosjektmappe må opprettes før tilbud kan genereres: ${state.folderName}`);
  }
  const children = await getSharePointFolderChildren(state.driveId, state.folder.id);
  const existingOfferFolder = children.find(item=>item?.folder && sharePointNameEquals(item.name, 'Tilbud'));
  if (existingOfferFolder?.id){
    return { driveId: state.driveId, projectFolder: state.folder, offerFolder: existingOfferFolder };
  }
  const offerFolder = await createSharePointChildFolder(state.driveId, state.folder.id, 'Tilbud', 'fail');
  if (!offerFolder?.id) throw new Error('Kunne ikke opprette Tilbud-mappe i prosjektmappen.');
  return { driveId: state.driveId, projectFolder: state.folder, offerFolder };
}

async function uploadProjectOfferToSharePoint(project, generatedOffer, targetFolder = null){
  if (!generatedOffer?.blob) throw new Error('Tilbudsfilen mangler.');
  const { driveId, offerFolder } = targetFolder || await ensureProjectOfferSharePointFolder(project);
  const normalizedOfferFileName = String(generatedOffer.fileName || 'Tilbud.docx')
    .replace(/-rev(\d+)(?=(\.[^.]+)?$)/i, '-$1');
  const safeFileName = sanitizeSharePointFolderName(normalizedOfferFileName, 'Tilbud.docx');
  const fileName = encodeURIComponent(safeFileName);
  return microsoftGraphRequest(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(offerFolder.id)}:/${fileName}:/content`,
    MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
    {
      method: 'PUT',
      body: generatedOffer.blob,
      rawBody: true,
      contentType: generatedOffer.blob.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    }
  );
}

async function findProjectOfferSharePointFolder(project){
  const state = await findProjectSharePointFolder(project);
  if (!state.folder?.id){
    throw new Error(`Fant ikke prosjektmappe i SharePoint: ${state.folderName}`);
  }
  const children = await getSharePointFolderChildren(state.driveId, state.folder.id);
  const offerFolder = children.find(item=>item?.folder && sharePointNameEquals(item.name, 'Tilbud'));
  if (!offerFolder?.id){
    throw new Error(`Fant ikke Tilbud-mappe under prosjektmappen ${state.folderName}.`);
  }
  return { driveId: state.driveId, projectFolder: state.folder, offerFolder };
}

async function findLatestProjectOfferFile(project){
  const { driveId, offerFolder } = await findProjectOfferSharePointFolder(project);
  const children = await getSharePointFolderChildren(driveId, offerFolder.id);
  const files = children
    .filter(item=>item?.file && /\.(docx?|pdf)$/i.test(String(item.name || '')))
    .sort((a, b)=>new Date(b.lastModifiedDateTime || 0).getTime() - new Date(a.lastModifiedDateTime || 0).getTime());
  const latest = files[0] || null;
  if (!latest?.webUrl){
    throw new Error('Fant ingen Word- eller PDF-fil i Tilbud-mappen for prosjektet.');
  }
  return latest;
}

async function createProjectFolderFromTemplate(projectId, triggerBtn){
  const project = getProjectById(projectId);
  if (!project) return;
  if (!authState.loggedIn){
    showLoginModal();
    return;
  }
  const btn = triggerBtn && triggerBtn.tagName === 'BUTTON' ? triggerBtn : null;
  const originalText = btn?.textContent || '';
  if (btn){
    btn.disabled = true;
    btn.textContent = 'Oppretter...';
  }
  setGraphStatus('projectFoldersStatus', 'Oppretter prosjektmappe...');
  try{
    const state = await getSharePointRootFolderItem('project-folders');
    const children = await getSharePointFolderChildren(state.driveId, state.rootItem.id);
    const template = children.find(item=>
      item?.folder && String(item.name || '').trim().toLowerCase() === PROJECT_FOLDER_TEMPLATE_NAME.toLowerCase()
    );
    if (!template?.id){
      throw new Error(`Fant ikke malmappen ${PROJECT_FOLDER_TEMPLATE_NAME} under Prosjektmapper.`);
    }
    const folderName = formatProjectFolderName(project);
    const projectFolder = await createSharePointChildFolder(state.driveId, state.rootItem.id, folderName, 'fail');
    if (!projectFolder?.id) throw new Error('Prosjektmappen ble ikke opprettet.');
    await cloneSharePointFolderContents(state.driveId, template.id, projectFolder.id);
    project.projectFolderName = folderName;
    project.projectFolderCreated = true;
    project.projectFolderWebUrl = projectFolder.webUrl || project.projectFolderWebUrl || '';
    projectFolderStatusState.byProjectId[project.id] = {
      exists: true,
      folderName,
      webUrl: project.projectFolderWebUrl || ''
    };
    projectFolderStatusState.loaded = true;
    project.updatedAt = new Date().toISOString();
    saveProjectsToStorage();
    renderProjectDashboard();
    await loadSharePointFolder('project-folders');
    setGraphStatus('projectFoldersStatus', `Opprettet ${folderName}`, 'ok');
    if (projectFolder.webUrl && window.confirm('Prosjektmappe opprettet. Åpne mappen i SharePoint?')){
      window.open(projectFolder.webUrl, '_blank', 'noopener,noreferrer');
    }
  }catch(err){
    console.warn('Oppretting av prosjektmappe feilet', err);
    setGraphStatus('projectFoldersStatus', err?.message || 'Kunne ikke opprette prosjektmappe.', 'error');
    window.alert(err?.message || 'Kunne ikke opprette prosjektmappe.');
  }finally{
    if (btn){
      btn.disabled = false;
      btn.textContent = originalText || 'Opprett prosjektmappe';
    }
  }
}

async function openProjectSharePointFolder(projectId, triggerBtn){
  const project = getProjectById(projectId);
  if (!project) return;
  if (!authState.loggedIn){
    showLoginModal();
    return;
  }
  const btn = triggerBtn && triggerBtn.tagName === 'BUTTON' ? triggerBtn : null;
  const originalText = btn?.textContent || '';
  if (btn){
    btn.disabled = true;
    btn.textContent = 'Åpner...';
  }
  try{
    let webUrl = String(project.projectFolderWebUrl || '').trim();
    if (!webUrl){
      const state = await findProjectSharePointFolder(project);
      if (!state.folder?.id){
        throw new Error(`Fant ikke prosjektmappe i SharePoint: ${state.folderName}`);
      }
      webUrl = String(state.folder.webUrl || '').trim();
      project.projectFolderName = String(state.folder.name || state.folderName || '').trim();
      project.projectFolderCreated = true;
      project.projectFolderWebUrl = webUrl;
      projectFolderStatusState.byProjectId[project.id] = {
        exists: true,
        folderName: project.projectFolderName,
        webUrl
      };
      project.updatedAt = new Date().toISOString();
      saveProjectsToStorage();
      renderProjectDashboard();
    }
    if (!webUrl) throw new Error('Fant prosjektmappen, men SharePoint returnerte ingen åpne-lenke.');
    window.open(webUrl, '_blank', 'noopener,noreferrer');
  }catch(err){
    window.alert(String(err?.message || 'Kunne ikke åpne prosjektmappe.'));
  }finally{
    if (btn){
      btn.disabled = false;
      btn.textContent = originalText || 'Åpne prosjektmappe';
    }
  }
}

async function uploadSharePointFile(page, file){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config || !file) return;
  if (file.size > 250 * 1024 * 1024){
    setGraphStatus(config.statusId, 'Filen er for stor for enkel Graph-opplasting. Maks 250 MB.', 'error');
    return;
  }
  setGraphStatus(config.statusId, `Laster opp ${file.name}...`);
  try{
    const state = await ensureSharePointFolderState(page);
    const folderPath = String(state.folderPath || '').split('/').map(encodeURIComponent).join('/');
    const fileName = encodeURIComponent(file.name);
    await microsoftGraphRequest(
      `/drives/${encodeURIComponent(state.driveId)}/root:/${folderPath}/${fileName}:/content`,
      MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
      {
        method: 'PUT',
        body: file,
        rawBody: true,
        contentType: file.type || 'application/octet-stream'
      }
    );
    await loadSharePointFolder(page);
  }catch(err){
    console.warn('SharePoint-opplasting feilet', err);
    setGraphStatus(config.statusId, err?.message || 'Kunne ikke laste opp fil.', 'error');
  }
}

async function deleteSharePointItem(page, itemId, name){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config || !itemId) return;
  if (!window.confirm(`Slette ${name || 'valgt element'}?`)) return;
  setGraphStatus(config.statusId, 'Sletter element...');
  try{
    const state = await ensureSharePointFolderState(page);
    await microsoftGraphRequest(
      `/drives/${encodeURIComponent(state.driveId)}/items/${encodeURIComponent(itemId)}`,
      MICROSOFT_GRAPH_SHAREPOINT_SCOPES,
      { method: 'DELETE' }
    );
    await loadSharePointFolder(page);
  }catch(err){
    console.warn('SharePoint-sletting feilet', err);
    setGraphStatus(config.statusId, err?.message || 'Kunne ikke slette element.', 'error');
  }
}

async function loadCalendarEvents(options = {}){
  if (!authState.loggedIn){
    setGraphStatus('calendarStatus', 'Logg inn med Microsoft for å vise kalender.', 'error');
    return;
  }
  const list = $('calendarEventsList');
  const grid = $('calendarGridView');
  if (!list && !grid) return;
  const range = calendarRangeForState();
  if (
    options.silent
    && calendarViewState.loadedStart
    && calendarViewState.loadedEnd
    && calendarViewState.loadedStart <= range.start
    && calendarViewState.loadedEnd >= range.end
  ){
    renderCalendarView();
    return;
  }
  const btn = $('refreshCalendarBtn');
  if (btn) btn.disabled = true;
  setGraphStatus('calendarStatus', 'Henter kalender...');
  try{
    const query = new URLSearchParams({
      startDateTime: range.start.toISOString(),
      endDateTime: range.end.toISOString(),
      '$top': '200',
      '$orderby': 'start/dateTime',
      '$select': 'id,subject,start,end,location,organizer,isAllDay,showAs,webLink,bodyPreview'
    });
    const payload = await microsoftGraphRequest(`/me/calendarView?${query.toString()}`, MICROSOFT_GRAPH_CALENDAR_SCOPES);
    const events = Array.isArray(payload?.value) ? payload.value : [];
    calendarViewState.events = events;
    calendarViewState.loadedStart = range.start;
    calendarViewState.loadedEnd = range.end;
    renderCalendarView();
    setGraphStatus('calendarStatus', `Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`, 'ok');
  }catch(err){
    console.warn('Kalenderhenting feilet', err);
    setGraphStatus('calendarStatus', err?.message || 'Kunne ikke hente kalender.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}

async function loadEmailMessages(options = {}){
  const list = $('emailMessagesList');
  if (!list) return;
  if (!authState.loggedIn){
    renderEmailMessages([]);
    delete list.dataset.loaded;
    setGraphStatus('emailStatus', 'Logg inn med Microsoft for å vise e-post.', 'error');
    return;
  }
  if (!canAccessProjectMailbox()){
    renderEmailMessages([]);
    delete list.dataset.loaded;
    setGraphStatus('emailStatus', `E-post er kun synlig for Owners. ${PROJECT_MAILBOX_ADDRESS} vises ikke for denne brukeren.`);
    return;
  }
  if (options.silent && list.dataset.loaded === '1') return;
  const btn = $('refreshEmailBtn');
  if (btn) btn.disabled = true;
  setGraphStatus('emailStatus', `Henter e-post fra ${PROJECT_MAILBOX_ADDRESS}...`);
  try{
    const query = new URLSearchParams({
      '$top': '25',
      '$orderby': 'receivedDateTime desc',
      '$select': 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance,webLink'
    });
    const payload = await microsoftGraphRequest(projectMailboxGraphPath(`/mailFolders/inbox/messages?${query.toString()}`), MICROSOFT_GRAPH_MAIL_SCOPES);
    const messages = Array.isArray(payload?.value) ? payload.value : [];
    renderEmailMessages(messages);
    list.dataset.loaded = '1';
    setGraphStatus('emailStatus', `Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`, 'ok');
  }catch(err){
    console.warn('E-posthenting feilet', err);
    setGraphStatus('emailStatus', err?.message || 'Kunne ikke hente e-post.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}

async function handleMicrosoftLogin(){
  const errorEl = $('loginError');
  const microsoftBtn = $('microsoftLoginBtn');
  const loginSubmit = $('loginSubmit');
  const registerSubmit = $('registerSubmit');
  if (errorEl) errorEl.textContent = 'Åpner Microsoft-innlogging...';
  if (microsoftBtn) microsoftBtn.disabled = true;
  if (loginSubmit) loginSubmit.disabled = true;
  if (registerSubmit) registerSubmit.disabled = true;
  try{
    const { client, config } = await getMicrosoftMsalClient();
    const result = await client.loginPopup({
      scopes: config.scopes,
      prompt: 'select_account'
    });
    if (!result?.idToken){
      throw new Error('Microsoft returnerte ikke ID-token.');
    }
    microsoftLastAccount = result.account || null;
    if (result.account && typeof client.setActiveAccount === 'function'){
      client.setActiveAccount(result.account);
    }
    if (errorEl) errorEl.textContent = 'Validerer Microsoft-innlogging...';
    const auth = await exchangeMicrosoftToken(result.idToken);
    await completeAuth(auth, 'Microsoft-innlogging fullført. Henter prosjekter fra server...');
  }catch(err){
    console.warn('Microsoft-innlogging feilet', err);
    if (errorEl) errorEl.textContent = err?.message || 'Microsoft-innlogging feilet.';
  }finally{
    if (microsoftBtn) microsoftBtn.disabled = false;
    if (loginSubmit) loginSubmit.disabled = false;
    if (registerSubmit) registerSubmit.disabled = false;
  }
}

function clearAuthSession(){
  authState = { loggedIn: false, username: '', token: '', profile: null, isAdmin: false };
  resetProjectFolderStatusState();
  projectState.customerDatabase = [];
  projectState.globalCustomerDatabaseLoaded = false;
  renderGlobalCustomerViews();
  persistAuthToSession();
}

function normalizeProfilePayload(raw){
  return {
    name: String(raw?.name || '').trim(),
    phone: String(raw?.phone || '').trim(),
    company: String(raw?.company || '').trim(),
    position: String(raw?.position || '').trim()
  };
}

async function submitAuthRequest(mode, email, password, options = {}){
  const endpoint = mode === 'register' ? '/api/auth/register' : '/api/auth/login';
  const res = await fetch(buildApiUrl(endpoint), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      email,
      password,
      confirmPassword: options.confirmPassword,
      profile: options.profile
    })
  });
  let payload = null;
  try{
    payload = await res.json();
  }catch(_err){}
  if (!res.ok){
    const message = payload?.error || (mode === 'register' ? 'Kunne ikke opprette bruker.' : 'Kunne ikke logge inn.');
    throw new Error(appendApiBaseHint(message, res.status));
  }
  const username = normalizeUserEmail(payload?.email);
  const token = typeof payload?.token === 'string' ? payload.token : '';
  if (!hasValidUserEmail(username) || !token){
    throw new Error('Serveren returnerte ugyldig innloggingsdata.');
  }
  return { username, token, profile: normalizeProfilePayload(payload?.profile), isAdmin: payload?.isAdmin === true };
}

function showLoginModal(){
  const modal = $('loginModal');
  if (!modal) return;
  modal.style.display = 'flex';
  const usernameInput = $('loginUsername');
  const passwordInput = $('loginPassword');
  const errorEl = $('loginError');
  if (errorEl) errorEl.textContent = '';
  if (usernameInput){
    usernameInput.value = authState.username || '';
    try{
      const len = usernameInput.value.length;
      usernameInput.focus();
      usernameInput.setSelectionRange(len, len);
    }catch(_err){
      usernameInput.focus();
    }
  }
  if (passwordInput){
    passwordInput.value = '';
  }
}

function hideLoginModal(){
  const modal = $('loginModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('loginError');
  if (errorEl) errorEl.textContent = '';
}

function showRegisterModal(){
  const modal = $('registerModal');
  if (!modal) return;
  const loginUsername = $('loginUsername');
  const registerEmail = $('registerEmail');
  const errorEl = $('registerError');
  if (errorEl) errorEl.textContent = '';
  if (registerEmail && loginUsername && loginUsername.value){
    registerEmail.value = normalizeUserEmail(loginUsername.value);
  }
  hideLoginModal();
  modal.style.display = 'flex';
  const nameInput = $('registerName');
  if (nameInput) nameInput.focus();
}

function hideRegisterModal(){
  const modal = $('registerModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('registerError');
  if (errorEl) errorEl.textContent = '';
}

async function completeAuth(auth, statusMessage){
  const username = normalizeUserEmail(auth.username);
  authState = {
    loggedIn: true,
    username,
    token: auth.token,
    profile: auth.profile || null,
    isAdmin: auth.isAdmin === true || ADMIN_NAV_ALLOWED_EMAILS.includes(username)
  };
  persistAuthToSession();
  hideLoginModal();
  hideRegisterModal();
  projectState.projects = [];
  projectState.expandedProjectId = null;
  sortProjects();
  updateProjectHistories();
  renderProjectDashboard();
  updateProjectMetaDisplay();
  updateAuthUI();
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = statusMessage || 'Henter prosjekter fra server...';
  await syncProjectsForCurrentUser();
  if (statusEl) statusEl.textContent = '';
}

async function handleAuthSubmit(mode = 'login'){
  const passwordInput = $('loginPassword');
  const usernameInput = $('loginUsername');
  const errorEl = $('loginError');
  const password = passwordInput ? passwordInput.value : '';
  const username = normalizeUserEmail(usernameInput ? usernameInput.value : '');
  if (!hasValidUserEmail(username)){
    if (errorEl) errorEl.textContent = 'Brukernavn må være en gyldig e-postadresse.';
    if (usernameInput){
      usernameInput.focus();
      try{
        const len = usernameInput.value.length;
        usernameInput.setSelectionRange(len, len);
      }catch(_err){}
    }
    return;
  }
  if (!password){
    if (errorEl) errorEl.textContent = 'Fyll inn passord.';
    if (passwordInput) passwordInput.focus();
    return;
  }
  if (mode === 'register' && password.length < 4){
    if (errorEl) errorEl.textContent = 'Passord må være minst 4 tegn.';
    if (passwordInput) passwordInput.focus();
    return;
  }
  const loginSubmit = $('loginSubmit');
  const registerSubmit = $('registerSubmit');
  if (loginSubmit) loginSubmit.disabled = true;
  if (registerSubmit) registerSubmit.disabled = true;
  if (errorEl) errorEl.textContent = mode === 'register' ? 'Oppretter bruker...' : 'Logger inn...';
  try{
    const auth = await submitAuthRequest(mode, username, password);
    await completeAuth(auth, mode === 'register' ? 'Bruker opprettet. Henter prosjekter fra server...' : 'Henter prosjekter fra server...');
  }catch(err){
    if (errorEl) errorEl.textContent = err?.message || 'Innlogging feilet.';
    if (passwordInput) passwordInput.focus();
  }finally{
    if (loginSubmit) loginSubmit.disabled = false;
    if (registerSubmit) registerSubmit.disabled = false;
  }
}

async function handleLoginSubmit(){
  await handleAuthSubmit('login');
}

async function handleRegisterSubmit(){
  const fields = {
    name: $('registerName'),
    email: $('registerEmail'),
    phone: $('registerPhone'),
    company: $('registerCompany'),
    position: $('registerPosition'),
    password: $('registerPassword'),
    confirmPassword: $('registerConfirmPassword')
  };
  const errorEl = $('registerError');
  const profile = normalizeProfilePayload({
    name: fields.name?.value,
    phone: fields.phone?.value,
    company: fields.company?.value,
    position: fields.position?.value
  });
  const email = normalizeUserEmail(fields.email?.value || '');
  const password = fields.password?.value || '';
  const confirmPassword = fields.confirmPassword?.value || '';
  const required = [
    [fields.name, profile.name],
    [fields.email, email],
    [fields.phone, profile.phone],
    [fields.company, profile.company],
    [fields.position, profile.position],
    [fields.password, password],
    [fields.confirmPassword, confirmPassword]
  ];
  const missing = required.find(([, value])=>!value);
  if (missing){
    if (errorEl) errorEl.textContent = 'Fyll inn alle obligatoriske felt.';
    missing[0]?.focus?.();
    return;
  }
  if (!hasValidUserEmail(email)){
    if (errorEl) errorEl.textContent = 'E-post må være gyldig.';
    fields.email?.focus?.();
    return;
  }
  if (password.length < 4){
    if (errorEl) errorEl.textContent = 'Passord må være minst 4 tegn.';
    fields.password?.focus?.();
    return;
  }
  if (password !== confirmPassword){
    if (errorEl) errorEl.textContent = 'Passordene er ikke like.';
    fields.confirmPassword?.focus?.();
    return;
  }
  const registerCreateBtn = $('registerCreateBtn');
  if (registerCreateBtn) registerCreateBtn.disabled = true;
  if (errorEl) errorEl.textContent = 'Oppretter bruker...';
  try{
    const auth = await submitAuthRequest('register', email, password, { confirmPassword, profile });
    await completeAuth(auth, 'Bruker opprettet. Henter prosjekter fra server...');
  }catch(err){
    if (errorEl) errorEl.textContent = err?.message || 'Kunne ikke opprette bruker.';
  }finally{
    if (registerCreateBtn) registerCreateBtn.disabled = false;
  }
}

const loginBtn = $('loginBtn');
if (loginBtn){
  loginBtn.addEventListener('click', showLoginModal);
}
const logoutBtn = $('logoutBtn');
if (logoutBtn){
  logoutBtn.addEventListener('click', ()=>{
    if (projectSyncState.timerId){
      clearTimeout(projectSyncState.timerId);
      projectSyncState.timerId = null;
    }
    projectSyncState.pending = false;
    clearAuthSession();
    hideLoginModal();
    clearProjectOverviewForLoggedOutUser();
    updateAuthUI();
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Logg inn for \u00E5 beregne.';
  });
}
const loginCancel = $('loginCancel');
if (loginCancel){
  loginCancel.addEventListener('click', hideLoginModal);
}
const loginSubmit = $('loginSubmit');
if (loginSubmit){
  loginSubmit.addEventListener('click', handleLoginSubmit);
}
const microsoftLoginBtn = $('microsoftLoginBtn');
if (microsoftLoginBtn){
  microsoftLoginBtn.addEventListener('click', handleMicrosoftLogin);
}
const refreshCalendarBtn = $('refreshCalendarBtn');
if (refreshCalendarBtn){
  refreshCalendarBtn.addEventListener('click', ()=>loadCalendarEvents());
}
async function refreshProjectsFromToolbar(){
  if (!authState.loggedIn){
    setGraphStatus('projectsStatus', 'Logg inn for å oppdatere prosjekter.', 'error');
    return;
  }
  const btn = $('refreshProjectsBtn');
  if (btn) btn.disabled = true;
  setGraphStatus('projectsStatus', 'Oppdaterer prosjekter...');
  try{
    await syncProjectsForCurrentUser();
    setGraphStatus('projectsStatus', `Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`, 'ok');
  }catch(err){
    console.warn('Prosjektoppdatering feilet', err);
    setGraphStatus('projectsStatus', err?.message || 'Kunne ikke oppdatere prosjekter.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}
const refreshProjectsBtn = $('refreshProjectsBtn');
if (refreshProjectsBtn){
  refreshProjectsBtn.addEventListener('click', ()=>void refreshProjectsFromToolbar());
}
const newCalendarEventBtn = $('newCalendarEventBtn');
if (newCalendarEventBtn){
  newCalendarEventBtn.addEventListener('click', ()=>openCalendarEventForm());
}
populateCalendarDurationOptions();
populateCalendarTimeOptions();
const calendarEventForm = $('calendarEventForm');
if (calendarEventForm){
  calendarEventForm.addEventListener('submit', evt=>{
    evt.preventDefault();
    void saveCalendarEventFromForm();
  });
}
const addCalendarAttendeeBtn = $('addCalendarAttendeeBtn');
if (addCalendarAttendeeBtn){
  addCalendarAttendeeBtn.addEventListener('click', ()=>{
    addCalendarFormAttendees($('calendarAttendeeInput')?.value);
  });
}
const calendarAttendeeInput = $('calendarAttendeeInput');
if (calendarAttendeeInput){
  calendarAttendeeInput.addEventListener('keydown', evt=>{
    if (evt.key !== 'Enter') return;
    evt.preventDefault();
    addCalendarFormAttendees(calendarAttendeeInput.value);
  });
}
const calendarEventStartDate = $('calendarEventStartDate');
if (calendarEventStartDate){
  calendarEventStartDate.addEventListener('input', syncCalendarNativeDatePicker);
}
const calendarEventDatePicker = $('calendarEventDatePicker');
if (calendarEventDatePicker){
  calendarEventDatePicker.addEventListener('change', ()=>{
    const value = calendarEventDatePicker.value;
    const textInput = $('calendarEventStartDate');
    if (textInput && value) textInput.value = formatDateInputValue(new Date(`${value}T00:00`));
  });
}
const calendarEventDatePickerBtn = $('calendarEventDatePickerBtn');
if (calendarEventDatePickerBtn){
  calendarEventDatePickerBtn.addEventListener('click', evt=>{
    evt.stopPropagation();
    syncCalendarNativeDatePicker();
    const popover = $('calendarDatePickerPopover');
    if (popover && !popover.hidden){
      closeCalendarDatePickerPopover();
    } else {
      openCalendarDatePickerPopover();
    }
  });
}
const calendarDatePickerPopover = $('calendarDatePickerPopover');
if (calendarDatePickerPopover){
  calendarDatePickerPopover.addEventListener('click', evt=>evt.stopPropagation());
}
document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target : null;
  if (target?.closest?.('.calendar-date-input-row')) return;
  closeCalendarDatePickerPopover();
});
const cancelCalendarEventBtn = $('cancelCalendarEventBtn');
if (cancelCalendarEventBtn){
  cancelCalendarEventBtn.addEventListener('click', closeCalendarEventForm);
}
const deleteCalendarEventBtn = $('deleteCalendarEventBtn');
if (deleteCalendarEventBtn){
  deleteCalendarEventBtn.addEventListener('click', ()=>void deleteCalendarEventFromForm());
}
const calendarModeSelect = $('calendarModeSelect');
if (calendarModeSelect){
  calendarModeSelect.value = calendarViewState.mode;
  calendarModeSelect.addEventListener('change', ()=>{
    const mode = String(calendarModeSelect.value || 'month');
    if (!['list', 'month', 'week'].includes(mode)) return;
    calendarViewState.mode = mode;
    if (mode === 'list') calendarViewState.cursor = new Date();
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    void loadCalendarEvents();
  });
}
const calendarPrevBtn = $('calendarPrevBtn');
if (calendarPrevBtn){
  calendarPrevBtn.addEventListener('click', ()=>{
    if (calendarViewState.mode === 'month'){
      calendarViewState.cursor = addMonths(calendarViewState.cursor, -1);
    } else if (calendarViewState.mode === 'week'){
      calendarViewState.cursor = addDays(calendarViewState.cursor, -7);
    } else {
      calendarViewState.cursor = addDays(calendarViewState.cursor, -14);
    }
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    void loadCalendarEvents();
  });
}
const calendarTodayBtn = $('calendarTodayBtn');
if (calendarTodayBtn){
  calendarTodayBtn.addEventListener('click', ()=>{
    calendarViewState.cursor = new Date();
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    void loadCalendarEvents();
  });
}
const calendarNextBtn = $('calendarNextBtn');
if (calendarNextBtn){
  calendarNextBtn.addEventListener('click', ()=>{
    if (calendarViewState.mode === 'month'){
      calendarViewState.cursor = addMonths(calendarViewState.cursor, 1);
    } else if (calendarViewState.mode === 'week'){
      calendarViewState.cursor = addDays(calendarViewState.cursor, 7);
    } else {
      calendarViewState.cursor = addDays(calendarViewState.cursor, 14);
    }
    calendarViewState.loadedStart = null;
    calendarViewState.loadedEnd = null;
    void loadCalendarEvents();
  });
}
const refreshEmailBtn = $('refreshEmailBtn');
if (refreshEmailBtn){
  refreshEmailBtn.addEventListener('click', ()=>loadEmailMessages());
}
const refreshOffersBtn = $('refreshOffersBtn');
if (refreshOffersBtn){
  refreshOffersBtn.addEventListener('click', ()=>loadOfferStatus());
}
const offerSearchInput = $('offerSearchInput');
if (offerSearchInput){
  offerSearchInput.addEventListener('input', ()=>{
    setOfferSearchTerm(offerSearchInput.value);
  });
}
const offerSortSelect = $('offerSortSelect');
if (offerSortSelect){
  offerSortSelect.addEventListener('change', ()=>{
    setOfferSortMode(offerSortSelect.value);
  });
}
const composeEmailBtn = $('composeEmailBtn');
if (composeEmailBtn){
  composeEmailBtn.addEventListener('click', openEmailComposeForm);
}
const emailComposeForm = $('emailComposeForm');
if (emailComposeForm){
  emailComposeForm.addEventListener('submit', evt=>{
    evt.preventDefault();
    void sendEmailFromForm();
  });
}
const cancelEmailComposeBtn = $('cancelEmailComposeBtn');
if (cancelEmailComposeBtn){
  cancelEmailComposeBtn.addEventListener('click', closeEmailComposeForm);
}
const openSelectedEmailBtn = $('openSelectedEmailBtn');
if (openSelectedEmailBtn){
  openSelectedEmailBtn.addEventListener('click', ()=>{
    const message = getSelectedEmailMessage();
    if (message?.webLink) window.open(message.webLink, '_blank', 'noopener,noreferrer');
  });
}
const markSelectedEmailReadBtn = $('markSelectedEmailReadBtn');
if (markSelectedEmailReadBtn){
  markSelectedEmailReadBtn.addEventListener('click', ()=>void markSelectedEmailRead());
}
const deleteSelectedEmailBtn = $('deleteSelectedEmailBtn');
if (deleteSelectedEmailBtn){
  deleteSelectedEmailBtn.addEventListener('click', ()=>void deleteSelectedEmail());
}
const refreshProjectFoldersBtn = $('refreshProjectFoldersBtn');
if (refreshProjectFoldersBtn){
  refreshProjectFoldersBtn.addEventListener('click', ()=>loadSharePointFolder('project-folders'));
}
const projectFolderSearchInput = $('projectFolderSearchInput');
if (projectFolderSearchInput){
  projectFolderSearchInput.addEventListener('input', ()=>{
    setSharePointSearchTerm('project-folders', projectFolderSearchInput.value);
  });
}
const projectFolderSortSelect = $('projectFolderSortSelect');
if (projectFolderSortSelect){
  projectFolderSortSelect.addEventListener('change', ()=>{
    setSharePointSortMode('project-folders', projectFolderSortSelect.value);
  });
}
const newProjectFolderBtn = $('newProjectFolderBtn');
if (newProjectFolderBtn){
  newProjectFolderBtn.addEventListener('click', ()=>void createSharePointFolder('project-folders'));
}
const uploadProjectFolderFile = $('uploadProjectFolderFile');
if (uploadProjectFolderFile){
  uploadProjectFolderFile.addEventListener('change', evt=>{
    const file = evt.target?.files?.[0] || null;
    void uploadSharePointFile('project-folders', file).finally(()=>{
      uploadProjectFolderFile.value = '';
    });
  });
}
const refreshSupplierFoldersBtn = $('refreshSupplierFoldersBtn');
if (refreshSupplierFoldersBtn){
  refreshSupplierFoldersBtn.addEventListener('click', ()=>loadSharePointFolder('supplier-folders'));
}
const supplierFolderSearchInput = $('supplierFolderSearchInput');
if (supplierFolderSearchInput){
  supplierFolderSearchInput.addEventListener('input', ()=>{
    setSharePointSearchTerm('supplier-folders', supplierFolderSearchInput.value);
  });
}
const supplierFolderSortSelect = $('supplierFolderSortSelect');
if (supplierFolderSortSelect){
  supplierFolderSortSelect.addEventListener('change', ()=>{
    setSharePointSortMode('supplier-folders', supplierFolderSortSelect.value);
  });
}
const newSupplierFolderBtn = $('newSupplierFolderBtn');
if (newSupplierFolderBtn){
  newSupplierFolderBtn.addEventListener('click', ()=>void createSharePointFolder('supplier-folders'));
}
const uploadSupplierFolderFile = $('uploadSupplierFolderFile');
if (uploadSupplierFolderFile){
  uploadSupplierFolderFile.addEventListener('change', evt=>{
    const file = evt.target?.files?.[0] || null;
    void uploadSharePointFile('supplier-folders', file).finally(()=>{
      uploadSupplierFolderFile.value = '';
    });
  });
}
const refreshCompaniesBtn = $('refreshCompaniesBtn');
if (refreshCompaniesBtn){
  refreshCompaniesBtn.addEventListener('click', ()=>loadGlobalCustomerDatabase());
}
const companySearchInput = $('companySearchInput');
if (companySearchInput){
  companySearchInput.addEventListener('input', ()=>{
    setCompanySearchTerm(companySearchInput.value);
  });
}
const companySortSelect = $('companySortSelect');
if (companySortSelect){
  companySortSelect.addEventListener('change', ()=>{
    setCompanySortMode(companySortSelect.value);
  });
}
const refreshContactsBtn = $('refreshContactsBtn');
if (refreshContactsBtn){
  refreshContactsBtn.addEventListener('click', ()=>loadGlobalCustomerDatabase());
}
const contactSearchInput = $('contactSearchInput');
if (contactSearchInput){
  contactSearchInput.addEventListener('input', ()=>{
    setContactSearchTerm(contactSearchInput.value);
  });
}
const contactSortSelect = $('contactSortSelect');
if (contactSortSelect){
  contactSortSelect.addEventListener('change', ()=>{
    setContactSortMode(contactSortSelect.value);
  });
}
const addCompanyBtn = $('addCompanyBtn');
if (addCompanyBtn){
  addCompanyBtn.addEventListener('click', ()=>openCompanyEditForm());
}
const companyEditForm = $('companyEditForm');
if (companyEditForm){
  companyEditForm.addEventListener('submit', handleCompanyFormSubmit);
}
const cancelCompanyEditBtn = $('cancelCompanyEditBtn');
if (cancelCompanyEditBtn){
  cancelCompanyEditBtn.addEventListener('click', closeCompanyEditForm);
}
const addContactBtn = $('addContactBtn');
if (addContactBtn){
  addContactBtn.addEventListener('click', ()=>openContactEditForm());
}
const contactEditForm = $('contactEditForm');
if (contactEditForm){
  contactEditForm.addEventListener('submit', handleContactFormSubmit);
}
const cancelContactEditBtn = $('cancelContactEditBtn');
if (cancelContactEditBtn){
  cancelContactEditBtn.addEventListener('click', closeContactEditForm);
}
const registerSubmit = $('registerSubmit');
if (registerSubmit){
  registerSubmit.addEventListener('click', showRegisterModal);
}
const registerCancel = $('registerCancel');
if (registerCancel){
  registerCancel.addEventListener('click', hideRegisterModal);
}
const registerCreateBtn = $('registerCreateBtn');
if (registerCreateBtn){
  registerCreateBtn.addEventListener('click', handleRegisterSubmit);
}
['loginUsername','loginPassword'].forEach(id=>{
  const input = $(id);
  if (input){
    input.addEventListener('keydown', evt=>{
      if (evt.key === 'Enter'){
        evt.preventDefault();
        handleLoginSubmit();
      } else if (evt.key === 'Escape'){
        evt.preventDefault();
        hideLoginModal();
      }
    });
  }
});
['registerName','registerEmail','registerPhone','registerCompany','registerPosition','registerPassword','registerConfirmPassword'].forEach(id=>{
  const input = $(id);
  if (input){
    input.addEventListener('keydown', evt=>{
      if (evt.key === 'Enter'){
        evt.preventDefault();
        handleRegisterSubmit();
      } else if (evt.key === 'Escape'){
        evt.preventDefault();
        hideRegisterModal();
      }
    });
  }
});
const loginModal = $('loginModal');
if (loginModal){
  loginModal.addEventListener('click', evt=>{
    if (evt.target === loginModal){
      hideLoginModal();
    }
  });
}
function addToHistory(list, value){
  const trimmed = String(value||'').trim();
  if (!trimmed) return;
  const lower = trimmed.toLowerCase();
  const existingIdx = list.findIndex(entry=>entry.toLowerCase() === lower);
  if (existingIdx !== -1){
    list.splice(existingIdx,1);
  }
  list.push(trimmed);
  if (list.length > 20){
    list.splice(0, list.length - 20);
  }
}

function generateProjectId(){
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'){
    return crypto.randomUUID();
  }
  return `proj-${Date.now().toString(36)}-${Math.random().toString(16).slice(2, 10)}`;
}

function normalizeProjectStatus(value){
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return PROJECT_STATUS_OPTIONS[0].id;
  if (PROJECT_STATUS_OPTIONS.some(option=>option.id === raw)) return raw;
  const mapped = {
    uavklart: 'unresolved',
    vunnet: 'won',
    tapt: 'lost',
    ferdig: 'finished'
  };
  return mapped[raw] || PROJECT_STATUS_OPTIONS[0].id;
}

function getProjectStatusConfig(projectOrStatus){
  const status = typeof projectOrStatus === 'string'
    ? normalizeProjectStatus(projectOrStatus)
    : normalizeProjectStatus(projectOrStatus?.projectStatus);
  return PROJECT_STATUS_OPTIONS.find(option=>option.id === status) || PROJECT_STATUS_OPTIONS[0];
}

function projectIsArchived(project){
  return PROJECT_ARCHIVE_STATUS_IDS.includes(getProjectStatusConfig(project).id);
}

function normalizeProject(raw){
  if (!raw) return null;
  const fallback = new Date().toISOString();
  const selectedAddonConfig = normalizeSelectedAddonConfig(raw.selectedAddonConfig || null, null);
  return {
    id: raw.id || generateProjectId(),
    projectNumber: String(raw.projectNumber || raw.project_number || raw.offerNumber || raw.offer_number || '').trim(),
    name: String(raw.name || '').trim(),
    customer: String(raw.customer || '').trim(),
    contactPerson: String(raw.contactPerson || raw.contact || '').trim(),
    customerAddress: String(raw.customerAddress || raw.address || '').trim(),
    customerPostalPlace: String(raw.customerPostalPlace || raw.postalPlace || '').trim(),
    contactPhone: String(raw.contactPhone || raw.phone || '').trim(),
    projectResponsible: String(raw.projectResponsible || raw.projectOwner || raw.ownerName || '').trim(),
    projectOwnerEmail: normalizeUserEmail(raw.projectOwnerEmail || raw.ownerEmail || ''),
    projectFolderName: String(raw.projectFolderName || '').trim(),
    projectFolderCreated: raw.projectFolderCreated === true,
    projectFolderWebUrl: String(raw.projectFolderWebUrl || '').trim(),
    projectStatus: normalizeProjectStatus(raw.projectStatus || raw.status),
    createdAt: raw.createdAt || fallback,
    updatedAt: raw.updatedAt || fallback,
    selectedAddonConfig,
    lines: Array.isArray(raw.lines) ? raw.lines : []
  };
}

function getCurrentProjectResponsibleName(){
  return String(authState?.profile?.name || authState?.username || '').trim();
}

function getProjectResponsibleName(project){
  return String(project?.projectResponsible || getCurrentProjectResponsibleName()).trim();
}

function normalizeUserEmail(value){
  return String(value || '').trim().toLowerCase();
}

function hasValidUserEmail(value){
  return EMAIL_REGEX.test(String(value || '').trim());
}

function getCurrentUserEmail(){
  if (!authState || authState.loggedIn !== true) return '';
  const email = normalizeUserEmail(authState.username);
  if (!hasValidUserEmail(email)) return '';
  return email;
}

function getProjectsStorageKeyForEmail(email){
  const normalized = normalizeUserEmail(email);
  if (!hasValidUserEmail(normalized)) return '';
  return `${PROJECTS_STORAGE_KEY_PREFIX}:${normalized}`;
}

function readProjectsFromStorageKey(storageKey){
  if (!storageKey || typeof localStorage === 'undefined') return [];
  try{
    const stored = localStorage.getItem(storageKey);
    if (!stored) return [];
    const parsed = JSON.parse(stored);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeProject).filter(Boolean);
  }catch(err){
    console.warn(`Kunne ikke lese prosjekter fra ${storageKey}`, err);
    return [];
  }
}

function getLocalProjectStorageKeysForMigration(email){
  const keys = new Set();
  const storageKey = getProjectsStorageKeyForEmail(email);
  if (storageKey) keys.add(storageKey);
  keys.add(LEGACY_PROJECTS_STORAGE_KEY);

  if (typeof localStorage === 'undefined') return Array.from(keys);
  try{
    for (let i = 0; i < localStorage.length; i += 1){
      const key = localStorage.key(i);
      if (key && key.startsWith(`${PROJECTS_STORAGE_KEY_PREFIX}:`)){
        keys.add(key);
      }
    }
  }catch(err){
    console.warn('Kunne ikke skanne lokale prosjektlister', err);
  }

  return Array.from(keys);
}

function loadLocalProjectsForMigration(email){
  const all = [];
  getLocalProjectStorageKeysForMigration(email).forEach(storageKey=>{
    readProjectsFromStorageKey(storageKey).forEach(project=>all.push(project));
  });
  return all;
}

function cleanupMigratedProjectStorage(email){
  if (typeof localStorage === 'undefined') return;
  const keepKey = getProjectsStorageKeyForEmail(email);
  try{
    const keysToRemove = [];
    for (let i = 0; i < localStorage.length; i += 1){
      const key = localStorage.key(i);
      if (!key) continue;
      if (key === LEGACY_PROJECTS_STORAGE_KEY){
        keysToRemove.push(key);
      } else if (key.startsWith(`${PROJECTS_STORAGE_KEY_PREFIX}:`) && key !== keepKey){
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(key=>localStorage.removeItem(key));
  }catch(err){
    console.warn('Kunne ikke rydde migrerte lokale prosjektlister', err);
  }
}

function clearProjectOverviewForLoggedOutUser(){
  projectState.projects = [];
  projectState.projectOwnerEmails = [];
  projectState.expandedProjectId = null;
  updateProjectHistories();
  clearActiveProject();
  renderProjectDashboard();
  updateProjectMetaDisplay();
}

function canUseProjectSyncApi(){
  return !isGithubPagesWithoutApiBase();
}

function getProjectUpdateTimestamp(project){
  if (!project || typeof project !== 'object') return 0;
  const updated = new Date(project.updatedAt || project.createdAt || 0).getTime();
  if (!Number.isFinite(updated)) return 0;
  return updated;
}

function mergeProjectsByLatest(localProjects, remoteProjects){
  const merged = new Map();
  const add = project=>{
    const normalized = normalizeProject(project);
    if (!normalized) return;
    const key = normalized.id || `${normalized.name}|${normalized.customer}|${normalized.contactPerson}`;
    const existing = merged.get(key);
    if (!existing){
      merged.set(key, normalized);
      return;
    }
    const existingTs = getProjectUpdateTimestamp(existing);
    const candidateTs = getProjectUpdateTimestamp(normalized);
    if (candidateTs > existingTs){
      merged.set(key, normalized);
      return;
    }
    if (candidateTs === existingTs){
      const existingStatus = normalizeProjectStatus(existing.projectStatus);
      const candidateStatus = normalizeProjectStatus(project?.projectStatus || project?.status);
      if (candidateStatus !== 'unresolved' || existingStatus === 'unresolved'){
        merged.set(key, normalized);
      }
    }
  };
  (Array.isArray(localProjects) ? localProjects : []).forEach(add);
  (Array.isArray(remoteProjects) ? remoteProjects : []).forEach(add);
  return Array.from(merged.values());
}

async function fetchUserProjectsFromServer(email){
  const query = encodeURIComponent(email);
  const res = await fetch(buildApiUrl(`/api/user-projects?email=${query}`), {
    cache: 'no-store',
    headers: authHeaders()
  });
  if (res.status === 401 || res.status === 403){
    clearAuthSession();
    updateAuthUI();
  }
  if (!res.ok){
    let message = `Kunne ikke hente prosjekter fra server (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        message += `: ${data.error.trim()}`;
      }
    }catch(_err){}
    throw new Error(appendApiBaseHint(message, res.status));
  }
  const payload = await res.json();
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  return {
    updatedAt: typeof payload?.updatedAt === 'string' ? payload.updatedAt : null,
    ownerEmails: Array.isArray(payload?.ownerEmails) ? payload.ownerEmails.map(normalizeUserEmail).filter(hasValidUserEmail) : [],
    projects: projects.map(normalizeProject).filter(Boolean)
  };
}

async function pushUserProjectsToServer(email, projects){
  const res = await fetch(buildApiUrl('/api/user-projects/sync'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify({
      email,
      ownerEmails: authState.isAdmin === true ? projectState.projectOwnerEmails : undefined,
      projects: Array.isArray(projects) ? projects : []
    })
  });
  if (res.status === 401 || res.status === 403){
    clearAuthSession();
    updateAuthUI();
  }
  if (!res.ok){
    let message = `Kunne ikke synkronisere prosjekter (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        message += `: ${data.error.trim()}`;
      }
    }catch(_err){}
    throw new Error(appendApiBaseHint(message, res.status));
  }
  const payload = await res.json();
  const syncedProjects = Array.isArray(payload?.projects) ? payload.projects : [];
  projectState.projectOwnerEmails = Array.isArray(payload?.ownerEmails)
    ? payload.ownerEmails.map(normalizeUserEmail).filter(hasValidUserEmail)
    : projectState.projectOwnerEmails;
  return syncedProjects.map(normalizeProject).filter(Boolean);
}

async function fetchCustomerDatabaseFromServer(){
  const res = await fetch(buildApiUrl('/api/customer-database'), {
    cache: 'no-store',
    headers: authHeaders()
  });
  if (res.status === 401 || res.status === 403){
    clearAuthSession();
    updateAuthUI();
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
  const payload = await res.json();
  return Array.isArray(payload?.customers) ? payload.customers : [];
}

function setGlobalCustomerStatus(id, message, state = '') {
  const el = $(id);
  if (!el) return;
  el.textContent = message || '';
  el.classList.toggle('error', state === 'error');
  el.classList.toggle('ok', state === 'ok');
}

function normalizeGlobalCustomerPayload(payload) {
  return (Array.isArray(payload) ? payload : []).map(customer => ({
    id: String(customer?.id || ''),
    name: String(customer?.name || '').trim(),
    address: String(customer?.address || '').trim(),
    postalPlace: String(customer?.postalPlace || '').trim(),
    segment: String(customer?.segment || '').trim(),
    customerResponsible: String(customer?.customerResponsible || customer?.responsible || '').trim(),
    projectCount: Number.isFinite(Number(customer?.projectCount)) ? Number(customer.projectCount) : 0,
    contacts: (Array.isArray(customer?.contacts) ? customer.contacts : []).map(contact => ({
      id: String(contact?.id || ''),
      name: String(contact?.name || '').trim(),
      phone: String(contact?.phone || '').trim(),
      email: String(contact?.email || '').trim()
    })).filter(contact => contact.name || contact.phone || contact.email)
  })).filter(customer => customer.name);
}

function flattenGlobalContacts() {
  return normalizeGlobalCustomerPayload(projectState.customerDatabase).flatMap(customer => (
    customer.contacts.map(contact => ({
      ...contact,
      customerName: customer.name,
      customerAddress: customer.address,
      customerPostalPlace: customer.postalPlace,
      customerSegment: customer.segment,
      customerResponsible: customer.customerResponsible
    }))
  )).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'no', { sensitivity: 'base', numeric: true }));
}

function compareGlobalListItems(a, b, mode, nameAccessor){
  if (mode === 'alpha_desc') return compareNoText(nameAccessor(b), nameAccessor(a));
  if (mode === 'date_newest' || mode === 'date_oldest'){
    const aTime = new Date(a?.updatedAt || a?.createdAt || 0).getTime() || 0;
    const bTime = new Date(b?.updatedAt || b?.createdAt || 0).getTime() || 0;
    return mode === 'date_oldest' ? aTime - bTime : bTime - aTime;
  }
  return compareNoText(nameAccessor(a), nameAccessor(b));
}

function companyMatchesSearch(customer){
  const term = normalizeListSearchText(globalListState.companySearchTerm);
  if (!term) return true;
  const haystack = [
    customer?.name,
    customer?.address,
    customer?.postalPlace,
    customer?.segment,
    customer?.customerResponsible,
    ...(Array.isArray(customer?.contacts) ? customer.contacts.flatMap(contact=>[contact?.name, contact?.phone, contact?.email]) : [])
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

function contactMatchesSearch(contact){
  const term = normalizeListSearchText(globalListState.contactSearchTerm);
  if (!term) return true;
  const haystack = [
    contact?.name,
    contact?.phone,
    contact?.email,
    contact?.customerName,
    contact?.customerAddress,
    contact?.customerPostalPlace,
    contact?.customerSegment,
    contact?.customerResponsible
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

function getGlobalCustomerByName(name) {
  const key = normalizeLookupKey(name);
  return normalizeGlobalCustomerPayload(projectState.customerDatabase).find(customer => normalizeLookupKey(customer.name) === key) || null;
}

async function loadGlobalCustomerDatabase(options = {}) {
  if (!authState.loggedIn) {
    setGlobalCustomerStatus('companiesStatus', 'Logg inn for å vise firma.', 'error');
    setGlobalCustomerStatus('contactsStatus', 'Logg inn for å vise kontaktpersoner.', 'error');
    return;
  }
  if (options.silent && projectState.globalCustomerDatabaseLoaded) {
    renderGlobalCustomerViews();
    return;
  }
  setGlobalCustomerStatus('companiesStatus', 'Henter firma...');
  setGlobalCustomerStatus('contactsStatus', 'Henter kontaktpersoner...');
  try {
    projectState.customerDatabase = normalizeGlobalCustomerPayload(await fetchCustomerDatabaseFromServer());
    projectState.globalCustomerDatabaseLoaded = true;
    updateProjectHistories();
    renderGlobalCustomerViews();
    const stamp = new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' });
    setGlobalCustomerStatus('companiesStatus', `Oppdatert ${stamp}`, 'ok');
    setGlobalCustomerStatus('contactsStatus', `Oppdatert ${stamp}`, 'ok');
  } catch (err) {
    console.warn('Kunne ikke hente global kundedatabase', err);
    setGlobalCustomerStatus('companiesStatus', err?.message || 'Kunne ikke hente firma.', 'error');
    setGlobalCustomerStatus('contactsStatus', err?.message || 'Kunne ikke hente kontaktpersoner.', 'error');
  }
}

async function saveGlobalCustomerRecord(payload) {
  const res = await fetch(buildApiUrl('/api/customer-database/upsert'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify(payload || {})
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    throw new Error(data?.error || `Kunne ikke lagre (${res.status})`);
  }
  projectState.customerDatabase = normalizeGlobalCustomerPayload(data?.customers);
  projectState.globalCustomerDatabaseLoaded = true;
  updateProjectHistories();
  renderGlobalCustomerViews();
  return data;
}

async function deleteGlobalCustomerRecord(payload) {
  const res = await fetch(buildApiUrl('/api/customer-database/delete'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify(payload || {})
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    throw new Error(data?.error || `Kunne ikke slette (${res.status})`);
  }
  projectState.customerDatabase = normalizeGlobalCustomerPayload(data?.customers);
  projectState.globalCustomerDatabaseLoaded = true;
  updateProjectHistories();
  renderGlobalCustomerViews();
  return data;
}

function renderGlobalCustomerViews() {
  renderCompanyCardsList();
  renderContactPersonsList();
}

function appendGlobalDataDetails(parent, rows) {
  const details = document.createElement('div');
  details.className = 'global-data-details';
  rows.forEach(([label, value]) => {
    const row = document.createElement('p');
    row.className = 'global-data-detail';
    row.textContent = `${label}: ${String(value || '').trim() || '-'}`;
    details.appendChild(row);
  });
  parent.appendChild(details);
}

function renderCompanyCardsList() {
  const list = $('companiesList');
  const companyOptions = $('globalCompanyOptions');
  if (companyOptions) {
    companyOptions.innerHTML = '';
    normalizeGlobalCustomerPayload(projectState.customerDatabase).forEach(customer => {
      const option = document.createElement('option');
      option.value = customer.name;
      companyOptions.appendChild(option);
    });
  }
  if (!list) return;
  const customers = normalizeGlobalCustomerPayload(projectState.customerDatabase)
    .filter(companyMatchesSearch)
    .sort((a, b)=>compareGlobalListItems(a, b, globalListState.companySort, item=>item?.name));
  const canEdit = canEditGlobalCustomerData();
  const form = $('companyEditForm');
  if (form) form.hidden = !canEdit || form.dataset.editing !== '1';
  const addBtn = $('addCompanyBtn');
  if (addBtn) addBtn.hidden = !canEdit;
  const ownerHint = $('companiesOwnerHint');
  if (ownerHint) ownerHint.hidden = canEdit;
  list.innerHTML = '';
  if (!customers.length) {
    const empty = document.createElement('div');
    empty.className = 'global-data-empty';
    empty.textContent = normalizeListSearchText(globalListState.companySearchTerm)
      ? 'Ingen firma matcher søket.'
      : 'Ingen firma funnet i prosjektarkivet.';
    list.appendChild(empty);
    return;
  }
  customers.forEach(customer => {
    const item = document.createElement('article');
    item.className = 'global-data-card';
    const body = document.createElement('div');
    body.className = 'global-data-card-body';
    const title = document.createElement('h3');
    title.textContent = customer.name;
    const projectCount = Number.isFinite(Number(customer.projectCount)) ? Number(customer.projectCount) : 0;
    body.appendChild(title);
    appendGlobalDataDetails(body, [
      ['Adresse', customer.address],
      ['Postnummer og sted', customer.postalPlace],
      ['Kundesegment', customer.segment],
      ['Kundeansvarlig', customer.customerResponsible],
      ['Kontaktpersoner', customer.contacts.length],
      ['Prosjekter totalt', projectCount]
    ]);
    item.appendChild(body);
    if (canEdit) {
      const actions = document.createElement('div');
      actions.className = 'global-data-actions';
      const edit = document.createElement('button');
      edit.type = 'button';
      edit.className = 'btn alt btn-small';
      edit.textContent = 'Endre';
      edit.addEventListener('click', () => openCompanyEditForm(customer));
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'btn danger btn-small';
      remove.textContent = 'Slett';
      remove.addEventListener('click', () => handleDeleteCompany(customer));
      actions.append(edit, remove);
      item.appendChild(actions);
    }
    list.appendChild(item);
  });
}

function renderContactPersonsList() {
  const list = $('contactsList');
  if (!list) return;
  const contacts = flattenGlobalContacts()
    .filter(contactMatchesSearch)
    .sort((a, b)=>compareGlobalListItems(a, b, globalListState.contactSort, item=>item?.name));
  const canEdit = canEditGlobalCustomerData();
  const form = $('contactEditForm');
  if (form) form.hidden = !canEdit || form.dataset.editing !== '1';
  const addBtn = $('addContactBtn');
  if (addBtn) addBtn.hidden = !canEdit;
  const ownerHint = $('contactsOwnerHint');
  if (ownerHint) ownerHint.hidden = canEdit;
  list.innerHTML = '';
  if (!contacts.length) {
    const empty = document.createElement('div');
    empty.className = 'global-data-empty';
    empty.textContent = normalizeListSearchText(globalListState.contactSearchTerm)
      ? 'Ingen kontaktpersoner matcher søket.'
      : 'Ingen kontaktpersoner funnet i prosjektarkivet.';
    list.appendChild(empty);
    return;
  }
  contacts.forEach(contact => {
    const item = document.createElement('article');
    item.className = 'global-data-card';
    const body = document.createElement('div');
    body.className = 'global-data-card-body';
    const title = document.createElement('h3');
    title.textContent = contact.name || 'Uten navn';
    body.appendChild(title);
    appendGlobalDataDetails(body, [
      ['Firmanavn', contact.customerName],
      ['Telefon', contact.phone],
      ['E-post', contact.email]
    ]);
    item.appendChild(body);
    if (canEdit) {
      const actions = document.createElement('div');
      actions.className = 'global-data-actions';
      const edit = document.createElement('button');
      edit.type = 'button';
      edit.className = 'btn alt btn-small';
      edit.textContent = 'Endre';
      edit.addEventListener('click', () => openContactEditForm(contact));
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'btn danger btn-small';
      remove.textContent = 'Slett';
      remove.addEventListener('click', () => handleDeleteContact(contact));
      actions.append(edit, remove);
      item.appendChild(actions);
    }
    list.appendChild(item);
  });
}

function openCompanyEditForm(customer = null) {
  if (!canEditGlobalCustomerData()) return;
  const form = $('companyEditForm');
  if (!form) return;
  form.dataset.editing = '1';
  form.hidden = false;
  openFormModal('companyEditForm', customer ? 'Endre kunde' : 'Legg til kunde');
  $('companyOriginalName').value = customer?.name || '';
  $('companyNameInput').value = customer?.name || '';
  $('companyAddressInput').value = customer?.address || '';
  $('companyPostalPlaceInput').value = customer?.postalPlace || '';
  $('companySegmentInput').value = customer?.segment || '';
  $('companyResponsibleInput').value = customer?.customerResponsible || '';
  $('companyNameInput')?.focus();
}

function closeCompanyEditForm() {
  const form = $('companyEditForm');
  if (!form) return;
  form.dataset.editing = '0';
  closeFormModal('companyEditForm');
  form.reset();
}

function openContactEditForm(contact = null) {
  if (!canEditGlobalCustomerData()) return;
  const form = $('contactEditForm');
  if (!form) return;
  form.dataset.editing = '1';
  form.hidden = false;
  openFormModal('contactEditForm', contact ? 'Endre kontakt' : 'Legg til kontakt');
  $('contactOriginalCustomer').value = contact?.customerName || '';
  $('contactOriginalName').value = contact?.name || '';
  $('contactCompanyInput').value = contact?.customerName || '';
  $('contactNameInput').value = contact?.name || '';
  $('contactPhoneInput').value = contact?.phone || '';
  $('contactEmailInput').value = contact?.email || '';
  $('contactCompanyInput')?.focus();
}

function closeContactEditForm() {
  const form = $('contactEditForm');
  if (!form) return;
  form.dataset.editing = '0';
  closeFormModal('contactEditForm');
  form.reset();
}

async function handleCompanyFormSubmit(evt) {
  evt?.preventDefault?.();
  const originalCustomer = String($('companyOriginalName')?.value || '').trim();
  const customer = String($('companyNameInput')?.value || '').trim();
  const address = String($('companyAddressInput')?.value || '').trim();
  const postalPlace = String($('companyPostalPlaceInput')?.value || '').trim();
  const segment = String($('companySegmentInput')?.value || '').trim();
  const customerResponsible = String($('companyResponsibleInput')?.value || '').trim();
  if (!customer) {
    setGlobalCustomerStatus('companiesStatus', 'Firmanavn mangler.', 'error');
    return;
  }
  setGlobalCustomerStatus('companiesStatus', 'Lagrer...');
  try {
    await saveGlobalCustomerRecord({ originalCustomer, customer, address, postalPlace, segment, customerResponsible });
    closeCompanyEditForm();
    setGlobalCustomerStatus('companiesStatus', 'Lagret', 'ok');
  } catch (err) {
    setGlobalCustomerStatus('companiesStatus', err?.message || 'Kunne ikke lagre firma.', 'error');
  }
}

async function handleContactFormSubmit(evt) {
  evt?.preventDefault?.();
  const originalCustomer = String($('contactOriginalCustomer')?.value || '').trim();
  const originalContactPerson = String($('contactOriginalName')?.value || '').trim();
  const customer = String($('contactCompanyInput')?.value || '').trim();
  const contactPerson = String($('contactNameInput')?.value || '').trim();
  const phone = String($('contactPhoneInput')?.value || '').trim();
  const email = String($('contactEmailInput')?.value || '').trim();
  if (!customer || !contactPerson) {
    setGlobalCustomerStatus('contactsStatus', 'Firma og kontaktperson mangler.', 'error');
    return;
  }
  const existingCustomer = getGlobalCustomerByName(customer);
  setGlobalCustomerStatus('contactsStatus', 'Lagrer...');
  try {
    await saveGlobalCustomerRecord({
      originalCustomer: originalCustomer || customer,
      originalContactPerson,
      customer,
      address: existingCustomer?.address || '',
      postalPlace: existingCustomer?.postalPlace || '',
      segment: existingCustomer?.segment || '',
      customerResponsible: existingCustomer?.customerResponsible || '',
      contactPerson,
      phone,
      email
    });
    closeContactEditForm();
    setGlobalCustomerStatus('contactsStatus', 'Lagret', 'ok');
  } catch (err) {
    setGlobalCustomerStatus('contactsStatus', err?.message || 'Kunne ikke lagre kontaktperson.', 'error');
  }
}

async function handleDeleteCompany(customer) {
  if (!customer?.name) return;
  if (!window.confirm(`Slette firmaet "${customer.name}" globalt? Dette tømmer firmafeltene i prosjekter som bruker firmaet.`)) return;
  setGlobalCustomerStatus('companiesStatus', 'Sletter...');
  try {
    await deleteGlobalCustomerRecord({ customer: customer.name });
    setGlobalCustomerStatus('companiesStatus', 'Slettet', 'ok');
  } catch (err) {
    setGlobalCustomerStatus('companiesStatus', err?.message || 'Kunne ikke slette firma.', 'error');
  }
}

async function handleDeleteContact(contact) {
  if (!contact?.customerName || !contact?.name) return;
  if (!window.confirm(`Slette kontaktpersonen "${contact.name}" fra ${contact.customerName}?`)) return;
  setGlobalCustomerStatus('contactsStatus', 'Sletter...');
  try {
    await deleteGlobalCustomerRecord({ customer: contact.customerName, contactPerson: contact.name });
    setGlobalCustomerStatus('contactsStatus', 'Slettet', 'ok');
  } catch (err) {
    setGlobalCustomerStatus('contactsStatus', err?.message || 'Kunne ikke slette kontaktperson.', 'error');
  }
}

async function flushProjectSync(){
  if (!canUseProjectSyncApi()) return;
  const email = getCurrentUserEmail();
  if (!email) return;
  if (projectSyncState.inFlight){
    projectSyncState.pending = true;
    return;
  }
  projectSyncState.inFlight = true;
  try{
    const syncedProjects = await pushUserProjectsToServer(email, projectState.projects);
    if (syncedProjects.length){
      projectState.projects = mergeProjectsByLatest(projectState.projects, syncedProjects);
      sortProjects();
      updateProjectHistories();
      saveProjectsToStorage({ skipRemoteSync: true });
      renderProjectDashboard();
      updateProjectMetaDisplay();
    }
  }catch(err){
    console.warn('Kunne ikke synkronisere prosjekter mot server', err);
  }finally{
    projectSyncState.inFlight = false;
    if (projectSyncState.pending){
      projectSyncState.pending = false;
      queueProjectSync({ immediate: true });
    }
  }
}

function queueProjectSync(options = {}){
  if (!canUseProjectSyncApi()) return;
  const email = getCurrentUserEmail();
  if (!email) return;
  if (projectSyncState.timerId){
    clearTimeout(projectSyncState.timerId);
    projectSyncState.timerId = null;
  }
  if (options.immediate){
    void flushProjectSync();
    return;
  }
  projectSyncState.timerId = window.setTimeout(()=>{
    projectSyncState.timerId = null;
    void flushProjectSync();
  }, PROJECT_SYNC_DEBOUNCE_MS);
}

async function syncProjectsForCurrentUser(){
  if (!canUseProjectSyncApi()) return;
  const email = getCurrentUserEmail();
  if (!email) return;
  try{
    const remoteSnapshot = await fetchUserProjectsFromServer(email);
    const remoteProjects = Array.isArray(remoteSnapshot?.projects) ? remoteSnapshot.projects : [];
    projectState.projectOwnerEmails = Array.isArray(remoteSnapshot?.ownerEmails) && remoteSnapshot.ownerEmails.length
      ? remoteSnapshot.ownerEmails
      : [email];
    const hasAuthoritativeEmptyRemote = !!remoteSnapshot?.updatedAt && remoteProjects.length === 0;
    const localProjects = hasAuthoritativeEmptyRemote
      ? []
      : mergeProjectsByLatest(projectState.projects, loadLocalProjectsForMigration(email));
    const mergedProjects = hasAuthoritativeEmptyRemote
      ? []
      : mergeProjectsByLatest(localProjects, remoteProjects);
    projectState.projects = mergedProjects;
    sortProjects();
    updateProjectHistories();
    saveProjectsToStorage({ skipRemoteSync: true });
    renderProjectDashboard();
    updateProjectMetaDisplay();
    const syncedProjects = await pushUserProjectsToServer(email, projectState.projects);
    if (syncedProjects.length){
      projectState.projects = mergeProjectsByLatest(projectState.projects, syncedProjects);
      sortProjects();
      updateProjectHistories();
      saveProjectsToStorage({ skipRemoteSync: true });
      renderProjectDashboard();
      updateProjectMetaDisplay();
    }
    projectState.customerDatabase = normalizeGlobalCustomerPayload(await fetchCustomerDatabaseFromServer());
    projectState.globalCustomerDatabaseLoaded = true;
    updateProjectHistories();
    renderGlobalCustomerViews();
    cleanupMigratedProjectStorage(email);
  }catch(err){
    console.warn('Kunne ikke hente prosjekter fra server', err);
    const fallbackProjects = mergeProjectsByLatest(projectState.projects, loadLocalProjectsForMigration(email));
    projectState.projects = fallbackProjects;
    sortProjects();
    updateProjectHistories();
    saveProjectsToStorage({ skipRemoteSync: true });
    renderProjectDashboard();
    updateProjectMetaDisplay();
  }
}

function loadProjectsFromStorage(){
  const email = getCurrentUserEmail();
  const storageKey = getProjectsStorageKeyForEmail(email);
  if (!storageKey) return [];
  return readProjectsFromStorageKey(storageKey);
}

function saveProjectsToStorage(options = {}){
  const email = getCurrentUserEmail();
  const storageKey = getProjectsStorageKeyForEmail(email);
  if (!storageKey) return;
  if (typeof localStorage === 'undefined') return;
  try{
    localStorage.setItem(storageKey, JSON.stringify(projectState.projects));
    if (localStorage.getItem(LEGACY_PROJECTS_STORAGE_KEY)){
      localStorage.removeItem(LEGACY_PROJECTS_STORAGE_KEY);
    }
  }catch(err){
    console.warn('Kunne ikke lagre prosjekter', err);
  }
  if (!options.skipRemoteSync){
    queueProjectSync({ immediate: true });
  }
}

function loadSortMode(storageKey, validModes, fallback){
  if (typeof localStorage === 'undefined') return fallback;
  try{
    const raw = String(localStorage.getItem(storageKey) || '').trim();
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

function normalizeProjectSearchText(value){
  return String(value || '').trim().toLowerCase();
}

function projectMatchesSearch(project, rawSearchTerm = projectState.projectSearchTerm){
  const term = normalizeProjectSearchText(rawSearchTerm);
  if (!term) return true;
  const haystack = [
    project?.name,
    project?.customer,
    project?.contactPerson,
    getProjectResponsibleName(project),
    getProjectStatusConfig(project).label
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

function setProjectSearchTerm(value, options = {}){
  projectState.projectSearchTerm = normalizeProjectSearchText(value);
  const input = $('projectSearchInput');
  if (input && input.value !== value){
    input.value = value || '';
  }
  if (options.render !== false){
    renderProjectDashboard();
  }
}
const registerModal = $('registerModal');
if (registerModal){
  registerModal.addEventListener('click', evt=>{
    if (evt.target === registerModal){
      hideRegisterModal();
    }
  });
}

function compareProjectsForSort(a, b, mode = projectState.projectSort){
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

function compareLinesForSort(a, b, mode = projectState.lineSort){
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
  const projectSelect = $('projectSortSelect');
  const lineSelect = $('lineSortSelect');
  if (projectSelect && PROJECT_SORT_OPTIONS.includes(projectState.projectSort)){
    projectSelect.value = projectState.projectSort;
  }
  if (lineSelect && LINE_SORT_OPTIONS.includes(projectState.lineSort)){
    lineSelect.value = projectState.lineSort;
  }
}

function applyDashboardSortModesFromStorage(){
  projectState.projectSort = loadSortMode(
    PROJECT_SORT_STORAGE_KEY,
    PROJECT_SORT_OPTIONS,
    'date_newest'
  );
  projectState.lineSort = loadSortMode(
    LINE_SORT_STORAGE_KEY,
    LINE_SORT_OPTIONS,
    'date_newest'
  );
  offerListState.sort = loadSortMode(
    OFFER_SORT_STORAGE_KEY,
    PROJECT_SORT_OPTIONS,
    'date_newest'
  );
  updateSortControlValues();
  updateOfferControlValues();
}

function setProjectSortMode(mode, options = {}){
  if (!PROJECT_SORT_OPTIONS.includes(mode)) return;
  projectState.projectSort = mode;
  sortProjects();
  if (options.persist !== false){
    saveSortMode(PROJECT_SORT_STORAGE_KEY, mode);
  }
  updateSortControlValues();
  if (options.render !== false){
    renderProjectDashboard();
  }
}

function setLineSortMode(mode, options = {}){
  if (!LINE_SORT_OPTIONS.includes(mode)) return;
  projectState.lineSort = mode;
  if (options.persist !== false){
    saveSortMode(LINE_SORT_STORAGE_KEY, mode);
  }
  updateSortControlValues();
  if (options.render !== false){
    renderProjectDashboard();
  }
}

function sortProjects(){
  projectState.projects.sort((a,b)=>compareProjectsForSort(a, b, projectState.projectSort));
}

function getProjectById(id){
  if (!id) return null;
  return projectState.projects.find(p=>p.id === id) || null;
}

function hasActiveProject(){
  return Boolean(projectState.currentProjectId && projectState.currentProject && projectState.currentCustomer);
}

function updateProjectHistories(){
  projectState.projectHistory.length = 0;
  projectState.customerHistory.length = 0;
  projectState.contactHistory.length = 0;
  (Array.isArray(projectState.customerDatabase) ? projectState.customerDatabase : []).forEach(customer=>{
    addToHistory(projectState.customerHistory, customer?.name);
    (Array.isArray(customer?.contacts) ? customer.contacts : []).forEach(contact=>{
      addToHistory(projectState.contactHistory, contact?.name);
    });
  });
  projectState.projects.forEach(project=>{
    addToHistory(projectState.projectHistory, project.name);
    addToHistory(projectState.customerHistory, project.customer);
    addToHistory(projectState.contactHistory, project.contactPerson);
  });
}

function normalizeLookupKey(value){
  return String(value || '').trim().toLowerCase();
}

function getCustomerRecord(customerName){
  const customerKey = normalizeLookupKey(customerName);
  if (!customerKey) return null;
  const record = {
    name: '',
    address: '',
    postalPlace: '',
    contacts: new Map()
  };
  (Array.isArray(projectState.customerDatabase) ? projectState.customerDatabase : []).forEach(customer=>{
    if (normalizeLookupKey(customer?.name) !== customerKey) return;
    if (!record.name) record.name = customer.name || customerName;
    if (!record.address && customer.address) record.address = customer.address;
    if (!record.postalPlace && customer.postalPlace) record.postalPlace = customer.postalPlace;
    (Array.isArray(customer.contacts) ? customer.contacts : []).forEach(contact=>{
      const contactName = String(contact?.name || '').trim();
      if (!contactName) return;
      const contactKey = normalizeLookupKey(contactName);
      const existing = record.contacts.get(contactKey) || { name: contactName, phone: '', email: '' };
      if (!existing.phone && contact.phone) existing.phone = contact.phone;
      if (!existing.email && contact.email) existing.email = contact.email;
      record.contacts.set(contactKey, existing);
    });
  });
  projectState.projects.forEach(project=>{
    if (normalizeLookupKey(project.customer) !== customerKey) return;
    if (!record.name) record.name = project.customer || customerName;
    if (!record.address && project.customerAddress) record.address = project.customerAddress;
    if (!record.postalPlace && project.customerPostalPlace) record.postalPlace = project.customerPostalPlace;
    const contactName = String(project.contactPerson || '').trim();
    if (contactName){
      const contactKey = normalizeLookupKey(contactName);
      const existing = record.contacts.get(contactKey) || { name: contactName, phone: '', email: '' };
      if (!existing.phone && project.contactPhone) existing.phone = project.contactPhone;
      record.contacts.set(contactKey, existing);
    }
  });
  return record.name ? record : null;
}

function getContactsForCustomer(customerName){
  const record = getCustomerRecord(customerName);
  if (!record) return [];
  return Array.from(record.contacts.values()).map(contact=>contact.name).filter(Boolean);
}

function getKnownProjectDetails(customerName, contactPerson){
  const record = getCustomerRecord(customerName);
  const contact = record?.contacts.get(normalizeLookupKey(contactPerson));
  return {
    customerAddress: record?.address || '',
    customerPostalPlace: record?.postalPlace || '',
    contactPhone: contact?.phone || '',
    contactEmail: contact?.email || ''
  };
}

function setActiveProject(project){
  const nextProject = typeof project === 'string' ? getProjectById(project) : project;
  if (!nextProject){
    clearActiveProject();
    return;
  }
  projectState.currentProjectId = nextProject.id;
  projectState.currentProject = nextProject.name;
  projectState.currentCustomer = nextProject.customer;
  projectState.currentContact = nextProject.contactPerson || '';
  addToHistory(projectState.projectHistory, nextProject.name);
  addToHistory(projectState.customerHistory, nextProject.customer);
  addToHistory(projectState.contactHistory, nextProject.contactPerson);
  updateProjectMetaDisplay();
  updateAuthUI();
}

function clearActiveProject(){
  projectState.currentProjectId = null;
  projectState.currentProject = '';
  projectState.currentCustomer = '';
  projectState.currentContact = '';
  updateProjectMetaDisplay();
  updateAuthUI();
}

function createProject(projectName, customerName, contactPerson, details = {}){
  const now = new Date().toISOString();
  const project = {
    id: generateProjectId(),
    projectNumber: '',
    name: projectName,
    customer: customerName,
    contactPerson: contactPerson,
    customerAddress: String(details.customerAddress || '').trim(),
    customerPostalPlace: String(details.customerPostalPlace || '').trim(),
    contactPhone: String(details.contactPhone || '').trim(),
    projectResponsible: getCurrentProjectResponsibleName(),
    projectOwnerEmail: getCurrentUserEmail(),
    createdAt: now,
    updatedAt: now,
    selectedAddonConfig: normalizeSelectedAddonConfig(null, null),
    lines: []
  };
  projectState.projects.push(project);
  sortProjects();
  saveProjectsToStorage();
  addToHistory(projectState.projectHistory, project.name);
  addToHistory(projectState.customerHistory, project.customer);
  addToHistory(projectState.contactHistory, project.contactPerson);
  renderProjectDashboard();
  return project;
}

function copyProject(sourceProjectId, customerName, contactPerson, details = {}){
  const source = getProjectById(sourceProjectId);
  if (!source) return null;
  const now = new Date().toISOString();
  const project = {
    id: generateProjectId(),
    projectNumber: '',
    name: source.name || 'Uten navn',
    customer: String(customerName || '').trim(),
    contactPerson: String(contactPerson || '').trim(),
    customerAddress: String(details.customerAddress || '').trim(),
    customerPostalPlace: String(details.customerPostalPlace || '').trim(),
    contactPhone: String(details.contactPhone || '').trim(),
    projectResponsible: getCurrentProjectResponsibleName(),
    projectOwnerEmail: getCurrentUserEmail(),
    createdAt: now,
    updatedAt: now,
    selectedAddonConfig: normalizeSelectedAddonConfig(source.selectedAddonConfig || null, null),
    lines: (Array.isArray(source.lines) ? source.lines : []).map(line=>({
      ...deepClone(line),
      id: generateProjectId()
    }))
  };
  projectState.projects.push(project);
  sortProjects();
  saveProjectsToStorage();
  addToHistory(projectState.projectHistory, project.name);
  addToHistory(projectState.customerHistory, project.customer);
  addToHistory(projectState.contactHistory, project.contactPerson);
  renderProjectDashboard();
  return project;
}

function updateProject(projectId, updates){
  const target = getProjectById(projectId);
  if (!target) return null;
  target.name = updates.name;
  target.customer = updates.customer;
  target.contactPerson = String(updates.contactPerson || '').trim();
  target.customerAddress = String(updates.customerAddress || '').trim();
  target.customerPostalPlace = String(updates.customerPostalPlace || '').trim();
  target.contactPhone = String(updates.contactPhone || '').trim();
  target.updatedAt = new Date().toISOString();
  sortProjects();
  saveProjectsToStorage();
  renderProjectDashboard();
  if (projectState.currentProjectId === projectId){
    projectState.currentProject = target.name;
    projectState.currentCustomer = target.customer;
    projectState.currentContact = target.contactPerson || '';
    updateProjectMetaDisplay();
    updateAuthUI();
  }
  addToHistory(projectState.projectHistory, target.name);
  addToHistory(projectState.customerHistory, target.customer);
  addToHistory(projectState.contactHistory, target.contactPerson);
  return target;
}

function deleteProject(projectId){
  const target = getProjectById(projectId);
  if (!target) return;
  const projectName = target.name || 'Uten navn';
  const confirmed = window.confirm(
    `Er du sikker på at du vil slette prosjektet \"${projectName}\"? Alle linjer i prosjektet blir slettet.`,
  );
  if (!confirmed) return;

  projectState.projects = projectState.projects.filter(project=>project.id !== projectId);
  if (projectState.expandedProjectId === projectId){
    projectState.expandedProjectId = null;
  }
  if (projectState.currentProjectId === projectId){
    resetCalculatorForm({ preserveProject: false });
  }

  sortProjects();
  saveProjectsToStorage();
  updateProjectHistories();
  renderProjectDashboard();
  updateProjectMetaDisplay();
  updateProjectSubmitState();

  const statusEl = $('status');
  if (statusEl){
    statusEl.textContent = `Prosjekt \"${projectName}\" er slettet.`;
  }
}

function deleteProjectLine(projectId, lineId){
  const project = getProjectById(projectId);
  if (!project || !Array.isArray(project.lines)) return;
  const idx = project.lines.findIndex(line=>line.id === lineId);
  if (idx < 0) return;

  const line = project.lines[idx];
  const lineLabel = line?.lineNumber || 'uten linjenummer';
  const confirmed = window.confirm(
    `Er du sikker på at du vil slette linje \"${lineLabel}\" fra prosjektet \"${project.name || 'Uten navn'}\"?`,
  );
  if (!confirmed) return;

  project.lines.splice(idx, 1);
  project.updatedAt = new Date().toISOString();
  sortProjects();
  saveProjectsToStorage();
  renderProjectDashboard();
  setActiveProject(project);

  const lineInput = $('lineNumberInput');
  if (lineInput && (lineInput.value || '').trim().toLowerCase() === String(lineLabel).toLowerCase()){
    lineInput.value = '';
    projectState.currentLineNumber = '';
  }

  const statusEl = $('status');
  if (statusEl){
    statusEl.textContent = `Linje \"${lineLabel}\" er slettet.`;
  }
}

function updateProjectMetaDisplay(){
  const hasData = hasActiveProject();
  const nameNodes = document.querySelectorAll('[data-project-name]');
  const customerNodes = document.querySelectorAll('[data-project-customer]');
  const contactNodes = document.querySelectorAll('[data-project-contact]');
  const metaWrappers = document.querySelectorAll('[data-project-meta]');
  const editButtons = document.querySelectorAll('[data-project-edit]');
  nameNodes.forEach(el=>{
    el.textContent = projectState.currentProject || '';
  });
  customerNodes.forEach(el=>{
    el.textContent = projectState.currentCustomer || '';
  });
  contactNodes.forEach(el=>{
    el.textContent = projectState.currentContact || '';
  });
  metaWrappers.forEach(el=>{
    el.hidden = !hasData;
  });
  editButtons.forEach(btn=>{
    btn.hidden = !hasData;
    btn.disabled = !hasData;
  });
}

function hideSuggestions(listEl){
  if (listEl){
    listEl.hidden = true;
    listEl.innerHTML = '';
  }
}

function showSuggestions(listEl, items){
  if (!listEl) return;
  const entries = items.filter(Boolean);
  if (!entries.length){
    hideSuggestions(listEl);
    return;
  }
  const frag = document.createDocumentFragment();
  [...entries].reverse().forEach(value=>{
    const li = document.createElement('li');
    li.textContent = value;
    li.dataset.value = value;
    frag.appendChild(li);
  });
  listEl.innerHTML = '';
  listEl.appendChild(frag);
  listEl.hidden = false;
}

function updateProjectSubmitState(){
  const submit = $('projectSubmit');
  if (!submit) return;
  const projectVal = (($('projectNameInput')?.value) || '').trim();
  const customerVal = (($('customerNameInput')?.value) || '').trim();
  const contactVal = (($('contactPersonInput')?.value) || '').trim();
  submit.disabled = !(projectVal && customerVal && contactVal);
}

function openProjectModal(options = {}){
  projectModalState.mode = options.mode || 'create';
  projectModalState.projectId = options.projectId || null;
  projectModalState.saveLineAfterCreate = Boolean(options.saveLineAfterCreate);
  projectModalState.copySourceProjectId = options.copySourceProjectId || null;
  const modal = $('projectModal');
  if (!modal) return;
  modal.style.display = 'flex';
  const errorEl = $('projectError');
  if (errorEl) errorEl.textContent = '';
  const projectInput = $('projectNameInput');
  const customerInput = $('customerNameInput');
  const contactInput = $('contactPersonInput');
  const titleEl = $('projectTitle');
  const sourceProject = projectModalState.mode === 'copy'
    ? getProjectById(projectModalState.copySourceProjectId)
    : null;
  if (titleEl){
    titleEl.textContent = projectModalState.mode === 'edit'
      ? 'Oppdater prosjekt'
      : (projectModalState.mode === 'copy' ? 'Kopier prosjekt' : 'Nytt prosjekt');
  }
  if (projectInput){
    projectInput.disabled = projectModalState.mode === 'copy';
    if (projectModalState.mode === 'copy'){
      projectInput.value = sourceProject?.name || '';
    } else if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      projectInput.value = existing?.name || '';
    } else {
      projectInput.value = '';
    }
    if (projectModalState.mode !== 'copy') projectInput.focus();
    const len = projectInput.value.length;
    try{
      projectInput.setSelectionRange(len, len);
    }catch(_err){
      /* ignore selection errors */
    }
  }
  if (customerInput){
    if (projectModalState.mode === 'copy'){
      customerInput.value = sourceProject?.customer || '';
    } else if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      customerInput.value = existing?.customer || '';
    } else {
      customerInput.value = '';
    }
  }
  if (contactInput){
    if (projectModalState.mode === 'copy'){
      contactInput.value = sourceProject?.contactPerson || '';
    } else if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      contactInput.value = existing?.contactPerson || '';
    } else {
      contactInput.value = '';
    }
  }
  if (projectModalState.mode === 'copy' && customerInput) customerInput.focus();
  updateContactSuggestionsForCustomer();
  updateProjectSubmitState();
}

function closeProjectModal(){
  const modal = $('projectModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('projectError');
  if (errorEl) errorEl.textContent = '';
  hideSuggestions($('projectSuggestions'));
  hideSuggestions($('customerSuggestions'));
  hideSuggestions($('contactSuggestions'));
  const projectInput = $('projectNameInput');
  if (projectInput) projectInput.disabled = false;
  projectModalState.mode = 'create';
  projectModalState.projectId = null;
  projectModalState.saveLineAfterCreate = false;
  projectModalState.pendingDetails = null;
  projectModalState.copySourceProjectId = null;
}

function updateContactSuggestionsForCustomer(){
  const customerInput = $('customerNameInput');
  const contactInput = $('contactPersonInput');
  const customerName = customerInput ? customerInput.value.trim() : '';
  const knownContacts = getContactsForCustomer(customerName);
  if (contactInput && !knownContacts.some(name=>normalizeLookupKey(name) === normalizeLookupKey(contactInput.value))){
    const customerRecord = getCustomerRecord(customerName);
    if (!customerRecord && customerInput && customerInput.value.trim()){
      contactInput.value = '';
    }
  }
  hideSuggestions($('contactSuggestions'));
}

function persistProjectInfo(projectName, customerName, contactPerson, options = {}){
  const trimmedName = projectName.trim();
  const trimmedCustomer = customerName.trim();
  const trimmedContact = contactPerson.trim();
  const knownDetails = getKnownProjectDetails(trimmedCustomer, trimmedContact);
  const customerAddress = String(options.customerAddress ?? knownDetails.customerAddress ?? '').trim();
  const customerPostalPlace = String(options.customerPostalPlace ?? knownDetails.customerPostalPlace ?? '').trim();
  const contactPhone = String(options.contactPhone ?? knownDetails.contactPhone ?? '').trim();
  if (options.projectId){
    updateProject(options.projectId, {
      name: trimmedName,
      customer: trimmedCustomer,
      contactPerson: trimmedContact,
      customerAddress,
      customerPostalPlace,
      contactPhone
    });
    setActiveProject(options.projectId);
    return;
  }
  if (options.copySourceProjectId){
    const copied = copyProject(options.copySourceProjectId, trimmedCustomer, trimmedContact, {
      customerAddress,
      customerPostalPlace,
      contactPhone
    });
    if (copied) setActiveProject(copied);
    return;
  }
  const created = createProject(trimmedName, trimmedCustomer, trimmedContact, {
    customerAddress,
    customerPostalPlace,
    contactPhone
  });
  setActiveProject(created);
}

function formatProjectTimestamp(value){
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '-';
  try{
    return fmtTimestampNO.format(date);
  }catch(_err){
    return date.toLocaleString('no-NO');
  }
}

function resolveSelectedAddonFlag(value, fallback = true){
  if (typeof value === 'boolean') return value;
  if (value === 1 || value === '1' || value === 'true' || value === 'TRUE') return true;
  if (value === 0 || value === '0' || value === 'false' || value === 'FALSE') return false;
  return fallback;
}

function normalizeSelectedAddonConfig(config, fallbackConfig = null){
  const fallback = fallbackConfig || {};
  const includeMontasje = resolveSelectedAddonFlag(
    config?.includeMontasje,
    resolveSelectedAddonFlag(fallback.includeMontasje, true)
  );
  const includeEngineering = resolveSelectedAddonFlag(
    config?.includeEngineering,
    resolveSelectedAddonFlag(fallback.includeEngineering, true)
  );
  const includeOppheng = resolveSelectedAddonFlag(
    config?.includeOppheng,
    resolveSelectedAddonFlag(fallback.includeOppheng, true)
  );
  const showMontasje = resolveSelectedAddonFlag(
    config?.showMontasje,
    resolveSelectedAddonFlag(
      config?.includeMontasje,
      resolveSelectedAddonFlag(
        fallback.showMontasje,
        resolveSelectedAddonFlag(fallback.includeMontasje, false)
      )
    )
  );
  const showEngineering = resolveSelectedAddonFlag(
    config?.showEngineering,
    resolveSelectedAddonFlag(
      config?.includeEngineering,
      resolveSelectedAddonFlag(
        fallback.showEngineering,
        resolveSelectedAddonFlag(fallback.includeEngineering, false)
      )
    )
  );
  const showOppheng = resolveSelectedAddonFlag(
    config?.showOppheng,
    resolveSelectedAddonFlag(
      config?.includeOppheng,
      resolveSelectedAddonFlag(
        fallback.showOppheng,
        resolveSelectedAddonFlag(fallback.includeOppheng, false)
      )
    )
  );
  const includeUnitPrices = resolveSelectedAddonFlag(
    config?.includeUnitPrices,
    resolveSelectedAddonFlag(fallback.includeUnitPrices, false)
  );
  return {
    includeMontasje,
    includeEngineering,
    includeOppheng,
    showMontasje: includeMontasje && showMontasje,
    showEngineering: includeEngineering && showEngineering,
    showOppheng: includeOppheng && showOppheng,
    includeUnitPrices
  };
}

function getSelectedAddonConfig(line, fallbackConfig = null){
  const totals = line?.totals || {};
  const raw = line?.selectedAddonConfig || totals.selectedAddonConfig || null;
  return normalizeSelectedAddonConfig(raw, fallbackConfig);
}

function getProjectSelectedAddonConfig(project){
  return normalizeSelectedAddonConfig(project?.selectedAddonConfig || null, null);
}

function getOfferAddonCheckboxValuesFromUI(){
  return normalizeSelectedAddonConfig({
    includeMontasje: Boolean($('includeMontasje')?.checked),
    includeEngineering: Boolean($('includeEngineering')?.checked),
    includeOppheng: Boolean($('includeOppheng')?.checked),
    showMontasje: Boolean($('showMontasje')?.checked),
    showEngineering: Boolean($('showEngineering')?.checked),
    showOppheng: Boolean($('showOppheng')?.checked),
    includeUnitPrices: Boolean($('includeUnitPrices')?.checked)
  }, null);
}

function applyOfferAddonCheckboxConstraints(){
  const pairs = [
    { includeId: 'includeMontasje', showId: 'showMontasje' },
    { includeId: 'includeEngineering', showId: 'showEngineering' },
    { includeId: 'includeOppheng', showId: 'showOppheng' }
  ];
  pairs.forEach(pair=>{
    const includeEl = $(pair.includeId);
    const showEl = $(pair.showId);
    if (!showEl) return;
    const includeChecked = Boolean(includeEl?.checked);
    showEl.disabled = !includeChecked;
    if (!includeChecked){
      showEl.checked = false;
    }
  });
}

function applySelectedAddonCheckboxes(line){
  const config = getSelectedAddonConfig(line);
  const includeMontasje = $('includeMontasje');
  const includeEngineering = $('includeEngineering');
  const includeOppheng = $('includeOppheng');
  const showMontasje = $('showMontasje');
  const showEngineering = $('showEngineering');
  const showOppheng = $('showOppheng');
  const includeUnitPrices = $('includeUnitPrices');
  if (includeMontasje) includeMontasje.checked = config.includeMontasje;
  if (includeEngineering) includeEngineering.checked = config.includeEngineering;
  if (includeOppheng) includeOppheng.checked = config.includeOppheng;
  if (showMontasje) showMontasje.checked = config.showMontasje;
  if (showEngineering) showEngineering.checked = config.showEngineering;
  if (showOppheng) showOppheng.checked = config.showOppheng;
  if (includeUnitPrices) includeUnitPrices.checked = config.includeUnitPrices;
  applyOfferAddonCheckboxConstraints();
  return config;
}

function formatLineSummary(line){
  if (!line) return '';
  const input = line.inputs || {};
  const formatElementType = value=>{
    const key = String(value || '').trim();
    if (!key) return '';
    const labels = {
      board_feed: 'Tavleelement',
      end_feed_unit: 'Endetilførselsboks',
      crt_board_feed: 'Trafoelement',
      end_cover: 'Endelokk',
      none: 'Ingen'
    };
    return labels[key] || key;
  };
  const parts = [];
  if (input.series) parts.push(input.series);
  const amp = Number(input.ampere ?? input.amp ?? line.ampere);
  if (Number.isFinite(amp)) parts.push(`${fmtIntNO.format(amp)}A`);
  if (input.ledere) parts.push(input.ledere);
  const meter = Number(input.meter);
  if (Number.isFinite(meter)) parts.push(`${fmtIntNO.format(meter)} meter`);
  const v90h = Number(input.v90h ?? input.v90_h);
  const v90v = Number(input.v90v ?? input.v90_v);
  const totalAngles = (Number.isFinite(v90h) ? v90h : 0) + (Number.isFinite(v90v) ? v90v : 0);
  if (totalAngles > 0) parts.push(`${fmtIntNO.format(totalAngles)} vinkler`);
  const startEl = formatElementType(input.startEl);
  if (startEl) parts.push(startEl);
  const sluttEl = formatElementType(input.sluttEl);
  if (sluttEl) parts.push(sluttEl);
  const fbQty = Number(input.fbQty ?? input.fireBarrierQty);
  if (Number.isFinite(fbQty) && fbQty > 0) parts.push(`Brann: ${fmtIntNO.format(fbQty)}`);
  const boxItems = normalizeBoxItems(input.boxItems, input.boxSel, input.boxQty, input.boxInnmatSum);
  const boxQty = boxItems.reduce((sum, item)=>sum + Number(item.boxQty || 0), 0);
  if (boxQty > 0) parts.push(`Bokser: ${fmtIntNO.format(boxQty)}`);
  return parts.join(' | ') || 'Ingen detaljer lagret';
}

function formatLineUpdatedText(line){
  if (!line) return '';
  const stamp = line.updatedAt || line.createdAt;
  if (!stamp) return '';
  return `Oppdatert ${formatProjectTimestamp(stamp)}`;
}

function resolveLineDisplayTotalWithConfig(line, config){
  const totals = line?.totals || {};
  const baseTotal = Number(totals.totalExMontasje);
  if (!Number.isFinite(baseTotal)){
    const directTotal = Number(line?.selectedAddonTotal ?? line?.totals?.selectedAddonTotal);
    return Number.isFinite(directTotal) ? round2(directTotal) : NaN;
  }
  const flags = normalizeSelectedAddonConfig(config, getSelectedAddonConfig(line));
  const includeMontasje = flags.includeMontasje;
  const includeEngineering = flags.includeEngineering;
  const includeOppheng = flags.includeOppheng;
  const montasjeTotal = Number(totals.totalInclMontasje);
  const engineeringTotal = Number(totals.totalInclEngineering);
  const opphengTotal = Number(totals.totalInclOppheng ?? totals.total);
  const tapOffOfferTotal = Number(totals.tapOffOfferTotal);
  const specialElementOfferTotal = Number(totals.specialElementOfferTotal);
  let total = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
  if (Number.isFinite(tapOffOfferTotal)){
    total += tapOffOfferTotal;
  } else if (Array.isArray(line?.bom)){
    total += calculateTapOffOfferTotal({
      bom: line.bom,
      tapOffMarginRate: totals.tapOffMarginRate ?? line?.inputs?.tapOffMarginRate ?? DEFAULT_MARGIN_RATE
    });
  }
  if (Number.isFinite(specialElementOfferTotal)){
    total += specialElementOfferTotal;
  } else if (Array.isArray(line?.bom)){
    total += calculateSpecialElementOfferTotal({
      bom: line.bom,
      tapOffMarginRate: totals.tapOffMarginRate ?? line?.inputs?.tapOffMarginRate ?? DEFAULT_MARGIN_RATE
    });
  }
  return round2(total);
}

function resolveLineDisplayTotal(line){
  return resolveLineDisplayTotalWithConfig(line, getSelectedAddonConfig(line));
}

function setLineSelectedAddonConfig(line, config){
  if (!line) return normalizeSelectedAddonConfig(config, null);
  const normalized = normalizeSelectedAddonConfig(config, getSelectedAddonConfig(line));
  line.selectedAddonConfig = deepClone(normalized);
  if (!line.totals || typeof line.totals !== 'object'){
    line.totals = {};
  }
  line.totals.selectedAddonConfig = deepClone(normalized);
  const computedTotal = resolveLineDisplayTotalWithConfig(line, normalized);
  if (Number.isFinite(computedTotal)){
    line.selectedAddonTotal = computedTotal;
    line.totals.selectedAddonTotal = computedTotal;
  }
  return normalized;
}

function syncActiveCalculatorAddonConfig(project){
  if (!hasCalculatorUI()) return;
  if (!project || projectState.currentProjectId !== project.id) return;
  const lineInput = $('lineNumberInput');
  const currentLine = String(projectState.currentLineNumber || lineInput?.value || '').trim().toLowerCase();
  if (!currentLine) return;
  const activeLine = (Array.isArray(project.lines) ? project.lines : [])
    .find(entry=>String(entry.lineNumber || '').trim().toLowerCase() === currentLine);
  if (!activeLine) return;
  applySelectedAddonCheckboxes(activeLine);
  updateSelectedAddonTotalUI();
}

function applyProjectAddonCheckboxesToCalculator(project){
  if (!hasCalculatorUI()) return;
  if (!project) return;
  const config = getProjectSelectedAddonConfig(project);
  const includeMontasje = $('includeMontasje');
  const includeEngineering = $('includeEngineering');
  const includeOppheng = $('includeOppheng');
  const showMontasje = $('showMontasje');
  const showEngineering = $('showEngineering');
  const showOppheng = $('showOppheng');
  const includeUnitPrices = $('includeUnitPrices');
  if (includeMontasje) includeMontasje.checked = config.includeMontasje;
  if (includeEngineering) includeEngineering.checked = config.includeEngineering;
  if (includeOppheng) includeOppheng.checked = config.includeOppheng;
  if (showMontasje) showMontasje.checked = config.showMontasje;
  if (showEngineering) showEngineering.checked = config.showEngineering;
  if (showOppheng) showOppheng.checked = config.showOppheng;
  if (includeUnitPrices) includeUnitPrices.checked = config.includeUnitPrices;
  applyOfferAddonCheckboxConstraints();
  updateSelectedAddonTotalUI();
}

function updateLineSelectedAddonConfig(projectId, lineId, partialConfig){
  const project = getProjectById(projectId);
  if (!project || !Array.isArray(project.lines)) return;
  const line = project.lines.find(entry=>entry.id === lineId);
  if (!line) return;
  const next = normalizeSelectedAddonConfig(partialConfig, getSelectedAddonConfig(line));
  setLineSelectedAddonConfig(line, next);
  saveProjectsToStorage();
  renderProjectDashboard();
  syncActiveCalculatorAddonConfig(project);
}

function updateProjectSelectedAddonConfig(projectId, partialConfig){
  const project = getProjectById(projectId);
  if (!project) return;
  const current = getProjectSelectedAddonConfig(project);
  const next = normalizeSelectedAddonConfig(partialConfig, current);
  project.selectedAddonConfig = deepClone(next);
  const lines = Array.isArray(project.lines) ? project.lines : [];
  lines.forEach(line=>setLineSelectedAddonConfig(line, next));
  project.updatedAt = new Date().toISOString();
  saveProjectsToStorage();
  sortProjects();
  renderProjectDashboard();
  syncActiveCalculatorAddonConfig(project);
}

function resolveLineMaterialMarginRate(line){
  const totals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
    ? line.inputs
    : {};
  return resolveMarginRateFromData({ totals, input });
}

function resolveProjectMaterialMarginRate(project){
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  for (const line of lines){
    const rate = resolveLineMaterialMarginRate(line);
    if (Number.isFinite(rate)) return rate;
  }
  return DEFAULT_MATERIAL_MARGIN_RATE;
}

function getProjectMaterialMarginStats(project){
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  const rates = lines
    .map(resolveLineMaterialMarginRate)
    .filter(rate=>Number.isFinite(rate))
    .map(rate=>normalizeMarginRate(rate, DEFAULT_MATERIAL_MARGIN_RATE));
  const roundedKeys = new Set(rates.map(rate=>rate.toFixed(6)));
  const minRate = rates.length ? Math.min(...rates) : NaN;
  const maxRate = rates.length ? Math.max(...rates) : NaN;
  return {
    lineCount: lines.length,
    uniqueCount: roundedKeys.size,
    minRate,
    maxRate
  };
}

function formatProjectMarginSummary(project){
  const stats = getProjectMaterialMarginStats(project);
  if (!stats.lineCount) return 'Prosjektet har ingen linjer ennå.';
  if (stats.uniqueCount <= 1 && Number.isFinite(stats.maxRate)){
    return `Nåværende DG for prosjektet er ${fmtPercentNO.format(stats.maxRate * 100)} %.`;
  }
  if (Number.isFinite(stats.minRate) && Number.isFinite(stats.maxRate)){
    return `DG varierer mellom ${fmtPercentNO.format(stats.minRate * 100)} % og ${fmtPercentNO.format(stats.maxRate * 100)} %. Ny verdi overstyrer alle linjer.`;
  }
  return 'DG er ikke satt på alle linjer. Ny verdi overstyrer alle linjer.';
}

function formatProjectMarginBadgeText(project){
  const stats = getProjectMaterialMarginStats(project);
  if (!stats.lineCount){
    return 'DG prosjekt: -';
  }
  if (stats.uniqueCount <= 1 && Number.isFinite(stats.maxRate)){
    return `DG prosjekt: ${fmtPercentNO.format(stats.maxRate * 100)} %`;
  }
  if (Number.isFinite(stats.minRate) && Number.isFinite(stats.maxRate)){
    return `DG prosjekt: ${fmtPercentNO.format(stats.minRate * 100)}-${fmtPercentNO.format(stats.maxRate * 100)} %`;
  }
  return `DG prosjekt: ${fmtPercentNO.format(resolveProjectMaterialMarginRate(project) * 100)} %`;
}

function shouldUseWarningForProjectMargin(stats){
  if (!stats || !stats.lineCount) return false;
  if (!Number.isFinite(stats.minRate) || !Number.isFinite(stats.maxRate)) return true;
  const epsilon = 0.000001;
  return (
    Math.abs(stats.minRate - DEFAULT_MATERIAL_MARGIN_RATE) > epsilon ||
    Math.abs(stats.maxRate - DEFAULT_MATERIAL_MARGIN_RATE) > epsilon
  );
}

function openProjectMarginModal(projectId){
  const project = getProjectById(projectId);
  if (!project) return false;
  const modal = $('projectMarginModal');
  if (!modal) return false;
  const lineCount = Array.isArray(project.lines) ? project.lines.length : 0;
  if (!lineCount){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Prosjektet har ingen linjer å oppdatere.';
    return true;
  }
  const titleEl = $('projectMarginTitle');
  const currentEl = $('projectMarginCurrent');
  const inputEl = $('projectMarginPercentInput');
  const errorEl = $('projectMarginError');
  if (errorEl) errorEl.textContent = '';
  if (titleEl){
    titleEl.textContent = `Sett prosjekt-DG: ${project.name || 'Uten navn'}`;
  }
  if (currentEl){
    currentEl.textContent = formatProjectMarginSummary(project);
  }
  if (inputEl){
    const currentRate = resolveProjectMaterialMarginRate(project);
    inputEl.value = fmtPercentNO.format(currentRate * 100);
    inputEl.focus();
    const len = inputEl.value.length;
    try{
      inputEl.setSelectionRange(0, len);
    }catch(_err){}
  }
  projectMarginModalState.projectId = project.id;
  modal.style.display = 'flex';
  return true;
}

function closeProjectMarginModal(){
  const modal = $('projectMarginModal');
  if (!modal) return;
  modal.style.display = 'none';
  projectMarginModalState.projectId = null;
  const errorEl = $('projectMarginError');
  if (errorEl) errorEl.textContent = '';
}

function openProjectStatusModal(projectId){
  const project = getProjectById(projectId);
  if (!project) return;
  const modal = $('projectStatusModal');
  if (!modal) return;
  projectStatusModalState.projectId = project.id;
  modal.dataset.projectId = project.id;
  const nameEl = $('projectStatusProjectName');
  if (nameEl){
    const title = project.projectNumber
      ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
      : (project.name || 'Uten navn');
    nameEl.textContent = title;
  }
  const currentStatus = getProjectStatusConfig(project).id;
  modal.querySelectorAll('[data-project-status-option]').forEach(btn=>{
    const active = btn.getAttribute('data-project-status-option') === currentStatus;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
  });
  modal.style.display = 'flex';
}

function closeProjectStatusModal(){
  const modal = $('projectStatusModal');
  if (modal){
    modal.style.display = 'none';
    delete modal.dataset.projectId;
  }
  projectStatusModalState.projectId = null;
}

function setProjectStatus(projectId, status){
  const project = getProjectById(projectId);
  if (!project) return;
  const nextStatus = normalizeProjectStatus(status);
  project.projectStatus = nextStatus;
  project.updatedAt = new Date().toISOString();
  saveProjectsToStorage();
  sortProjects();
  if (PROJECT_ARCHIVE_STATUS_IDS.includes(nextStatus)){
    projectState.expandedProjectId = null;
  }
  const statusEl = $('projectsStatus');
  if (statusEl){
    statusEl.textContent = `Prosjektstatus endret til ${getProjectStatusConfig(nextStatus).label}.`;
    statusEl.classList.remove('error');
    statusEl.classList.add('ok');
  }
  closeProjectStatusModal();
  renderProjectDashboard();
}

function submitProjectMarginModal(){
  const projectId = projectMarginModalState.projectId;
  if (!projectId){
    closeProjectMarginModal();
    return;
  }
  const inputEl = $('projectMarginPercentInput');
  const errorEl = $('projectMarginError');
  const raw = String(inputEl?.value ?? '').trim().replace(',','.');
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)){
    if (errorEl) errorEl.textContent = 'Oppgi en gyldig DG i prosent.';
    if (inputEl) inputEl.focus();
    return;
  }
  const nextRate = parsed > 1 ? parsed / 100 : parsed;
  if (!Number.isFinite(nextRate) || nextRate < 0 || nextRate > MAX_MARGIN_RATE){
    if (errorEl) errorEl.textContent = 'DG må være mellom 0 og 95 %.';
    if (inputEl) inputEl.focus();
    return;
  }
  const result = applyProjectMarginRate(projectId, nextRate);
  const appliedPercent = fmtPercentNO.format(result.appliedRate * 100);
  const statusEl = $('status');
  if (statusEl){
    statusEl.textContent = `DG ${appliedPercent}% er satt på prosjektet. Oppdatert ${result.updatedLines} linje(r).`;
  }
  closeProjectMarginModal();
}

function applyProjectMarginRate(projectId, nextRate){
  const project = getProjectById(projectId);
  if (!project) return { updatedLines: 0, skippedLines: 0 };
  const normalizedRate = normalizeMarginRate(nextRate, DEFAULT_MATERIAL_MARGIN_RATE);
  const lines = Array.isArray(project.lines) ? project.lines : [];
  let updatedLines = 0;
  let skippedLines = 0;

  lines.forEach(line=>{
    const totals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
      ? line.totals
      : null;
    if (!totals){
      skippedLines += 1;
      return;
    }
    const material = Number(totals.material);
    if (!Number.isFinite(material) || material < 0){
      skippedLines += 1;
      return;
    }

    const inputs = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};
    const freightRate = Number(inputs.freightRate ?? totals.freightRate ?? 0.10);
    const montasjeCost = Number(totals?.montasje?.cost);
    const engineeringCost = Number(totals?.engineering?.cost);
    const opphengCost = Number(totals?.oppheng?.cost);
    const montasjeMarginRate = resolveMontasjeDgRate(inputs?.montasjeMarginRate, totals?.montasjeMarginRate, DEFAULT_MARGIN_RATE);
    const engineeringMarginRate = resolveDgRate(inputs?.engineeringMarginRate, totals?.engineeringMarginRate, DEFAULT_MARGIN_RATE);
    const opphengMarginRate = resolveDgRate(inputs?.opphengMarginRate, totals?.opphengMarginRate, DEFAULT_MARGIN_RATE);

    const recalculated = calculateTotalsFromMaterial({
      material,
      marginRate: normalizedRate,
      freightRate,
      montasjeCost: Number.isFinite(montasjeCost) ? montasjeCost : 0,
      montasjeMarginRate,
      engineeringCost: Number.isFinite(engineeringCost) ? engineeringCost : 0,
      engineeringMarginRate,
      opphengCost: Number.isFinite(opphengCost) ? opphengCost : 0,
      opphengMarginRate
    });

    totals.marginRate = recalculated.marginRate;
    totals.marginFactor = recalculated.marginFactor;
    totals.margin = recalculated.margin;
    totals.subtotal = recalculated.subtotal;
    totals.freightRate = recalculated.freightRate;
    totals.freight = recalculated.freight;
    totals.totalExMontasje = recalculated.totalExMontasje;
    totals.montasjeMarginRate = recalculated.montasjeMarginRate;
    totals.montasjeMargin = recalculated.montasjeMargin;
    totals.totalInclMontasje = recalculated.totalInclMontasje;
    totals.engineeringMarginRate = recalculated.engineeringMarginRate;
    totals.engineeringMargin = recalculated.engineeringMargin;
    totals.totalInclEngineering = recalculated.totalInclEngineering;
    totals.opphengMarginRate = recalculated.opphengMarginRate;
    totals.opphengMargin = recalculated.opphengMargin;
    totals.totalInclOppheng = recalculated.totalInclOppheng;
    totals.total = recalculated.total;

    if (inputs){
      inputs.marginRate = recalculated.marginRate;
      inputs.freightRate = recalculated.freightRate;
      line.inputs = inputs;
    }

    setLineSelectedAddonConfig(line, getSelectedAddonConfig(line, getProjectSelectedAddonConfig(project)));
    line.updatedAt = new Date().toISOString();
    updatedLines += 1;
  });

  project.updatedAt = new Date().toISOString();
  saveProjectsToStorage();
  sortProjects();
  renderProjectDashboard();
  syncActiveCalculatorAddonConfig(project);
  return { updatedLines, skippedLines, appliedRate: normalizedRate };
}

function promptAndApplyProjectMarginRate(projectId){
  if (openProjectMarginModal(projectId)){
    return;
  }
  const project = getProjectById(projectId);
  if (!project) return;
  const lineCount = Array.isArray(project.lines) ? project.lines.length : 0;
  if (!lineCount){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Prosjektet har ingen linjer å oppdatere.';
    return;
  }
  const currentRate = resolveProjectMaterialMarginRate(project);
  const defaultPercent = fmtPercentNO.format(currentRate * 100);
  const input = window.prompt(
    `Angi DG% for hele prosjektet "${project.name || 'Uten navn'}". Denne overstyrer DG på alle linjer.`,
    defaultPercent
  );
  if (input === null) return;

  const parsed = Number(String(input).replace(',', '.'));
  if (!Number.isFinite(parsed)){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Ugyldig DG-verdi.';
    return;
  }
  const nextRate = normalizeMarginRate(parsed > 1 ? parsed / 100 : parsed, DEFAULT_MATERIAL_MARGIN_RATE);
  const result = applyProjectMarginRate(projectId, nextRate);
  const appliedPercent = fmtPercentNO.format(result.appliedRate * 100);
  const statusEl = $('status');
  if (statusEl){
    statusEl.textContent = `DG ${appliedPercent}% er satt på prosjektet. Oppdatert ${result.updatedLines} linje(r).`;
  }
}

function buildAddonSelectorControl(config, options = {}){
  const normalized = normalizeSelectedAddonConfig(config, null);
  const wrapper = document.createElement('div');
  const extraClass = options.className ? ` ${options.className}` : '';
  wrapper.className = `addon-config-panel${extraClass}`;
  const checkboxDefs = [
    { includeKey: 'includeMontasje', showKey: 'showMontasje', label: 'Montasje' },
    { includeKey: 'includeEngineering', showKey: 'showEngineering', label: 'Ingeniør' },
    { includeKey: 'includeOppheng', showKey: 'showOppheng', label: 'Opphengsmateriell' }
  ];

  const buildSelectorGroup = (titleText, mode, keyName)=>{
    const group = document.createElement('div');
    group.className = 'addon-selectors';
    const title = document.createElement('span');
    title.className = 'addon-selectors-title';
    const strong = document.createElement('strong');
    strong.textContent = titleText;
    title.appendChild(strong);
    group.appendChild(title);

    checkboxDefs.forEach(def=>{
      const labelEl = document.createElement('label');
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = Boolean(normalized[def[keyName]]);
      input.dataset.addonField = def[keyName];
      input.dataset.addonMode = mode;
      if (mode === 'show' && !normalized[def.includeKey]){
        input.checked = false;
        input.disabled = true;
      }
      if (options.scope === 'project'){
        input.dataset.projectAddon = '1';
        input.dataset.projectId = options.projectId || '';
      } else if (options.scope === 'line'){
        input.dataset.lineAddon = '1';
        input.dataset.projectId = options.projectId || '';
        input.dataset.lineId = options.lineId || '';
      }
      labelEl.appendChild(input);
      labelEl.appendChild(document.createTextNode(` ${def.label}`));
      group.appendChild(labelEl);
    });

    return group;
  };

  const buildUnitGroup = ()=>{
    const unitGroup = document.createElement('div');
    unitGroup.className = 'addon-selectors';
    const unitTitle = document.createElement('span');
    unitTitle.className = 'addon-selectors-title';
    const unitStrong = document.createElement('strong');
    unitStrong.textContent = 'Inkluder enhetspriser:';
    unitTitle.appendChild(unitStrong);
    unitGroup.appendChild(unitTitle);
    const unitLabel = document.createElement('label');
    const unitInput = document.createElement('input');
    unitInput.type = 'checkbox';
    unitInput.checked = Boolean(normalized.includeUnitPrices);
    unitInput.dataset.addonField = 'includeUnitPrices';
    unitInput.dataset.addonMode = 'unit';
    if (options.scope === 'project'){
      unitInput.dataset.projectAddon = '1';
      unitInput.dataset.projectId = options.projectId || '';
    } else if (options.scope === 'line'){
      unitInput.dataset.lineAddon = '1';
      unitInput.dataset.projectId = options.projectId || '';
      unitInput.dataset.lineId = options.lineId || '';
    }
    unitLabel.appendChild(unitInput);
    unitLabel.appendChild(document.createTextNode(' Enhetspriser'));
    unitGroup.appendChild(unitLabel);
    return unitGroup;
  };

  const groups = Array.isArray(options.groups) && options.groups.length
    ? options.groups
    : ['include', 'show', 'unit'];
  groups.forEach(group=>{
    if (group === 'include') wrapper.appendChild(buildSelectorGroup('Inkluder i tilbud:', 'include', 'includeKey'));
    if (group === 'show') wrapper.appendChild(buildSelectorGroup('Synliggjør pris:', 'show', 'showKey'));
    if (group === 'unit') wrapper.appendChild(buildUnitGroup());
  });
  return wrapper;
}

function formatLineTotal(line){
  const total = resolveLineDisplayTotal(line);
  return Number.isFinite(total) ? `${fmtNO.format(total)} NOK` : 'Ingen sum';
}

function resolveLineSkinMaterialCost(line){
  const totalsMaterial = Number(line?.totals?.material);
  if (Number.isFinite(totalsMaterial)) return round2(totalsMaterial);
  const directMaterial = Number(line?.material);
  return Number.isFinite(directMaterial) ? round2(directMaterial) : NaN;
}

function formatLineSkinMaterialCost(line){
  const material = resolveLineSkinMaterialCost(line);
  return Number.isFinite(material) ? `Materiell skinne: ${fmtNO.format(material)} NOK` : 'Materiell skinne: -';
}

function sanitizeDownloadFileName(value, fallback = 'tilbud'){
  const raw = String(value || '').trim() || fallback;
  const cleaned = raw.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').replace(/\s+/g, ' ').trim();
  return cleaned || fallback;
}

function downloadBlob(blob, fileName){
  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = fileName || 'nedlasting';
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(()=>URL.revokeObjectURL(blobUrl), 1000);
}

function getFilenameFromContentDisposition(headerValue){
  const source = String(headerValue || '');
  if (!source) return '';
  const utfMatch = source.match(/filename\*=UTF-8''([^;]+)/i);
  if (utfMatch && utfMatch[1]){
    try{
      return decodeURIComponent(utfMatch[1]).trim();
    }catch(_err){
      return utfMatch[1].trim();
    }
  }
  const plainMatch = source.match(/filename="?([^";]+)"?/i);
  if (plainMatch && plainMatch[1]){
    return plainMatch[1].trim();
  }
  return '';
}

async function generateProjectOffer(project){
  const res = await fetch(buildApiUrl('/api/generate-offer'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify({ project })
  });
  if (res.status === 401 || res.status === 403){
    clearAuthSession();
    updateAuthUI();
  }
  if (!res.ok){
    let errorText = `Tilbudsgenerering feilet (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        errorText += `: ${data.error.trim()}`;
      }
    }catch(_jsonErr){
      try{
        const txt = await res.text();
        if (txt && txt.trim()) errorText += `: ${txt.trim()}`;
      }catch(_textErr){}
    }
    const err = new Error(appendApiBaseHint(errorText, res.status));
    err.status = res.status;
    throw err;
  }

  const blob = await res.blob();
  const headerName = String(res.headers.get('X-Offer-Filename') || '').trim()
    || getFilenameFromContentDisposition(res.headers.get('Content-Disposition'));
  const offerNumber = String(res.headers.get('X-Offer-Number') || '').trim();
  const revision = String(res.headers.get('X-Offer-Revision') || '').trim();
  const projectName = sanitizeDownloadFileName(project?.name || 'prosjekt');
  const fallbackName = `Tilbud-${projectName}${offerNumber ? `-${offerNumber}` : ''}${revision ? `-${revision}` : ''}.docx`;
  return {
    blob,
    fileName: headerName || fallbackName,
    offerNumber,
    revision
  };
}

async function generateLatestProjectOffer(project){
  const res = await fetch(buildApiUrl('/api/generate-offer-latest'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify({ project })
  });
  if (res.status === 401 || res.status === 403){
    clearAuthSession();
    updateAuthUI();
  }
  if (!res.ok){
    let errorText = `Kunne ikke åpne tilbud (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        errorText += `: ${data.error.trim()}`;
      }
    }catch(_jsonErr){}
    const err = new Error(appendApiBaseHint(errorText, res.status));
    err.status = res.status;
    throw err;
  }
  const blob = await res.blob();
  const headerName = String(res.headers.get('X-Offer-Filename') || '').trim()
    || getFilenameFromContentDisposition(res.headers.get('Content-Disposition'));
  const offerNumber = String(res.headers.get('X-Offer-Number') || '').trim();
  const revision = String(res.headers.get('X-Offer-Revision') || '').trim();
  const projectName = sanitizeDownloadFileName(project?.name || 'prosjekt');
  const fallbackName = `Tilbud-${projectName}${offerNumber ? `-${offerNumber}` : ''}${revision ? `-${revision}` : ''}.docx`;
  return {
    blob,
    fileName: headerName || fallbackName,
    offerNumber,
    revision
  };
}

function normalizeOfferSearchText(value){
  return String(value || '').trim().toLowerCase();
}

function getOfferStatusForProject(project){
  const id = String(project?.id || '').trim();
  return id ? offerListState.statusByProjectId[id] || null : null;
}

function buildOfferRows(){
  return projectState.projects.map(project=>({
    project,
    status: getOfferStatusForProject(project)
  }));
}

function offerRowMatchesSearch(row){
  const term = normalizeOfferSearchText(offerListState.searchTerm);
  if (!term) return true;
  const status = row.status || {};
  const project = row.project || {};
  const haystack = [
    project.name,
    project.customer,
    project.contactPerson,
    getProjectResponsibleName(project),
    project.projectNumber,
    status.offerNumber,
    status.revision !== null && status.revision !== undefined ? `-${status.revision}` : ''
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

function compareOfferRows(a, b){
  if (offerListState.sort === 'alpha_asc'){
    return compareNoText(a?.project?.name, b?.project?.name);
  }
  if (offerListState.sort === 'alpha_desc'){
    return compareNoText(b?.project?.name, a?.project?.name);
  }
  const aTime = getSortableCreatedTimestamp(a?.project);
  const bTime = getSortableCreatedTimestamp(b?.project);
  if (offerListState.sort === 'date_oldest') return aTime - bTime;
  return bTime - aTime;
}

function renderOffersList(){
  const list = $('offersList');
  if (!list) return;
  list.innerHTML = '';
  const rows = buildOfferRows()
    .filter(offerRowMatchesSearch)
    .sort(compareOfferRows);
  if (!rows.length){
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.innerHTML = '<p>Ingen tilbud matcher søket.</p>';
    list.appendChild(empty);
    return;
  }
  rows.forEach(row=>{
    const project = row.project;
    const status = row.status || {};
    const item = document.createElement('article');
    item.className = `offer-row${status.hasOffer ? '' : ' is-missing-offer'}`;

    const body = document.createElement('div');
    body.className = 'offer-row-body';
    const title = document.createElement('h3');
    title.textContent = project.projectNumber
      ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
      : project.name || 'Uten navn';
    body.appendChild(title);
    appendGlobalDataDetails(body, [
      ['Prosjektansvarlig', getProjectResponsibleName(project)],
      ['Kunde', project.customer],
      ['Kontaktperson', project.contactPerson],
      ['Tilbud', status.hasOffer ? `${status.offerNumber || '-'}-${status.revision ?? '-'}` : 'Ingen genererte tilbud'],
      ['Oppdatert', formatProjectTimestamp(project.updatedAt || project.createdAt)]
    ]);

    const actions = document.createElement('div');
    actions.className = 'offer-row-actions';
    const openWordBtn = document.createElement('button');
    openWordBtn.type = 'button';
    openWordBtn.className = 'btn';
    openWordBtn.dataset.openOfferWord = project.id;
    openWordBtn.textContent = 'Åpne Word';
    openWordBtn.disabled = !authState.loggedIn;
    const pdfBtn = document.createElement('button');
    pdfBtn.type = 'button';
    pdfBtn.className = 'btn alt';
    pdfBtn.textContent = 'PDF';
    pdfBtn.disabled = true;
    pdfBtn.title = 'PDF er ikke tilgjengelig før serveren får dokumentkonvertering.';
    actions.append(openWordBtn, pdfBtn);

    item.append(body, actions);
    list.appendChild(item);
  });
}

function updateOfferControlValues(){
  const input = $('offerSearchInput');
  if (input && input.value !== offerListState.searchTerm) input.value = offerListState.searchTerm;
  const select = $('offerSortSelect');
  if (select && PROJECT_SORT_OPTIONS.includes(offerListState.sort)) select.value = offerListState.sort;
}

function setOfferSearchTerm(value, options = {}){
  offerListState.searchTerm = normalizeOfferSearchText(value);
  updateOfferControlValues();
  if (options.render !== false) renderOffersList();
}

function setOfferSortMode(mode, options = {}){
  if (!PROJECT_SORT_OPTIONS.includes(mode)) return;
  offerListState.sort = mode;
  if (options.persist !== false) saveSortMode(OFFER_SORT_STORAGE_KEY, mode);
  updateOfferControlValues();
  if (options.render !== false) renderOffersList();
}

function setCompanySearchTerm(value){
  globalListState.companySearchTerm = normalizeListSearchText(value);
  renderCompanyCardsList();
}

function setCompanySortMode(mode){
  if (!PROJECT_SORT_OPTIONS.includes(mode)) return;
  globalListState.companySort = mode;
  renderCompanyCardsList();
}

function setContactSearchTerm(value){
  globalListState.contactSearchTerm = normalizeListSearchText(value);
  renderContactPersonsList();
}

function setContactSortMode(mode){
  if (!PROJECT_SORT_OPTIONS.includes(mode)) return;
  globalListState.contactSort = mode;
  renderContactPersonsList();
}

function setSharePointSearchTerm(page, value){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config) return;
  const state = getSharePointListState(page);
  state.searchTerm = normalizeListSearchText(value);
  renderSharePointFolderItems(config, Array.isArray(state.items) ? state.items : [], page);
}

function setSharePointSortMode(page, mode){
  const config = SHAREPOINT_FOLDER_CONFIG[page];
  if (!config || !PROJECT_SORT_OPTIONS.includes(mode)) return;
  const state = getSharePointListState(page);
  state.sort = mode;
  renderSharePointFolderItems(config, Array.isArray(state.items) ? state.items : [], page);
}

async function loadOfferStatus(options = {}){
  const list = $('offersList');
  if (!list) return;
  if (!authState.loggedIn){
    setGraphStatus('offersStatus', 'Logg inn for å vise tilbud.', 'error');
    return;
  }
  if (options.silent && offerListState.loaded){
    renderOffersList();
    return;
  }
  const btn = $('refreshOffersBtn');
  if (btn) btn.disabled = true;
  setGraphStatus('offersStatus', 'Henter tilbud...');
  try{
    const res = await fetch(buildApiUrl('/api/offer-status'), {
      headers: authHeaders()
    });
    if (res.status === 401 || res.status === 403){
      clearAuthSession();
      updateAuthUI();
    }
    if (!res.ok){
      let message = `Kunne ikke hente tilbud (${res.status})`;
      try{
        const data = await res.json();
        if (data?.error) message += `: ${data.error}`;
      }catch(_err){}
      throw new Error(appendApiBaseHint(message, res.status));
    }
    const payload = await res.json();
    const next = {};
    (Array.isArray(payload?.offers) ? payload.offers : []).forEach(item=>{
      const projectId = String(item?.projectId || '').trim();
      if (!projectId) return;
      next[projectId] = {
        offerNumber: String(item.offerNumber || '').trim(),
        revision: item.revision === null || item.revision === undefined ? null : Number(item.revision),
        hasOffer: item.hasOffer === true
      };
    });
    offerListState.statusByProjectId = next;
    offerListState.loaded = true;
    renderOffersList();
    setGraphStatus('offersStatus', `Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`, 'ok');
  }catch(err){
    console.warn('Tilbudshenting feilet', err);
    setGraphStatus('offersStatus', err?.message || 'Kunne ikke hente tilbud.', 'error');
  }finally{
    if (btn) btn.disabled = false;
  }
}

async function openLatestOfferForProject(projectId, triggerBtn){
  const project = getProjectById(projectId);
  if (!project) return;
  const btn = triggerBtn && triggerBtn.tagName === 'BUTTON' ? triggerBtn : null;
  const originalText = btn?.textContent || '';
  if (btn){
    btn.disabled = true;
    btn.textContent = 'Åpner...';
  }
  try{
    const latestFile = await findLatestProjectOfferFile(project);
    window.open(latestFile.webUrl, '_blank', 'noopener,noreferrer');
  }catch(err){
    window.alert(String(err?.message || err));
  }finally{
    if (btn){
      btn.disabled = false;
      btn.textContent = originalText || 'Åpne Word';
    }
  }
}

function getMissingOfferDetails(project){
  const checks = [
    ['Kunde', project?.customer],
    ['Adresse', project?.customerAddress],
    ['Postnummer og sted', project?.customerPostalPlace],
    ['Kontaktperson', project?.contactPerson],
    ['Telefon', project?.contactPhone]
  ];
  return checks
    .filter(([, value])=>!String(value || '').trim())
    .map(([label])=>label);
}

function closeOfferDetailsWarning(shouldGenerate){
  const modal = $('offerDetailsWarningModal');
  if (modal) modal.style.display = 'none';
  const resolver = offerDetailsWarningState.resolver;
  offerDetailsWarningState.resolver = null;
  if (resolver) resolver(Boolean(shouldGenerate));
}

function confirmGenerateOfferWithMissingDetails(missingFields){
  return new Promise(resolve=>{
    const modal = $('offerDetailsWarningModal');
    const list = $('offerDetailsMissingList');
    if (!modal || !list){
      resolve(window.confirm(`Prosjektet mangler: ${missingFields.join(', ')}.\nKontakt administrator for å legge til detaljer.\n\nGenerere likevel?`));
      return;
    }
    list.innerHTML = '';
    missingFields.forEach(field=>{
      const li = document.createElement('li');
      li.textContent = field;
      list.appendChild(li);
    });
    offerDetailsWarningState.resolver = resolve;
    modal.style.display = 'flex';
  });
}

async function requestGenerateProjectOffer(projectId, triggerBtn){
  const project = getProjectById(projectId);
  if (!project) return;
  const missingDetails = getMissingOfferDetails(project);
  if (missingDetails.length){
    const shouldContinue = await confirmGenerateOfferWithMissingDetails(missingDetails);
    if (!shouldContinue) return;
  }

  const buttonEl = triggerBtn && triggerBtn.tagName === 'BUTTON' ? triggerBtn : null;
  const originalText = buttonEl ? buttonEl.textContent : '';
  let failed = false;
  if (buttonEl){
    buttonEl.disabled = true;
    buttonEl.textContent = 'Genererer...';
  }

  try{
    const offerTarget = await ensureProjectOfferSharePointFolder(project);
    project.projectFolderName = String(offerTarget.projectFolder?.name || formatProjectFolderName(project)).trim();
    project.projectFolderCreated = true;
    project.projectFolderWebUrl = offerTarget.projectFolder?.webUrl || project.projectFolderWebUrl || '';
    projectFolderStatusState.byProjectId[project.id] = {
      exists: true,
      folderName: project.projectFolderName,
      webUrl: project.projectFolderWebUrl || ''
    };
    const generated = await generateProjectOffer(project);
    await uploadProjectOfferToSharePoint(project, generated, offerTarget);
    if (generated.offerNumber){
      project.projectNumber = generated.offerNumber;
    }
    project.updatedAt = new Date().toISOString();
    saveProjectsToStorage();
    offerListState.loaded = false;
    renderProjectDashboard();
    void loadOfferStatus({ silent: true });

    if (buttonEl){
      buttonEl.textContent = generated.offerNumber ? `Lagret ${generated.offerNumber}` : 'Lagret';
    }
  }catch(err){
    failed = true;
    if (buttonEl){
      buttonEl.textContent = 'Feil, prøv igjen';
    }
    window.alert(String(err?.message || err));
  }

  if (buttonEl){
    setTimeout(()=>{
      buttonEl.disabled = false;
      buttonEl.textContent = failed ? 'Generer tilbud' : (originalText || 'Generer tilbud');
    }, 1500);
  }
}

function readProjectFlowStore(){
  if (typeof localStorage === 'undefined') return {};
  try{
    const raw = localStorage.getItem(getProjectFlowStorageKey());
    const parsed = raw ? JSON.parse(raw) : {};
    if (!parsed || typeof parsed !== 'object') return {};
    return parsed;
  }catch(_err){
    return {};
  }
}

function persistProjectFlowStore(){
  if (typeof localStorage === 'undefined') return;
  try{
    localStorage.setItem(getProjectFlowStorageKey(), JSON.stringify(projectFlowState.milestonesByProjectId || {}));
  }catch(_err){}
}

function getProjectFlowStorageKey(){
  const email = getCurrentUserEmail();
  return email ? `${PROJECT_FLOW_STORAGE_KEY}.${email}` : `${PROJECT_FLOW_STORAGE_KEY}.local`;
}

function normalizeProjectFlowMilestones(items){
  return (Array.isArray(items) ? items : [])
    .map(item=>{
      const startDate = String(item?.startDate || item?.date || '').trim();
      const endDate = String(item?.endDate || item?.date || startDate).trim();
      return {
        id: String(item?.id || '').trim(),
        phaseId: PROJECT_FLOW_PHASES.some(phase=>phase.id === item?.phaseId) ? item.phaseId : PROJECT_FLOW_PHASES[0].id,
        title: String(item?.title || '').trim(),
        startDate,
        endDate,
        date: startDate,
        durationValue: Math.max(1, Number.parseInt(item?.durationValue, 10) || 1),
        durationUnit: ['days', 'weeks', 'months'].includes(item?.durationUnit) ? item.durationUnit : 'days',
        fileName: String(item?.fileName || '').trim(),
        drivenByTaskId: String(item?.drivenByTaskId || (item?.dependencyRelation === 'drivenBy' ? item?.dependencyTaskId : '') || '').trim(),
        drivesTaskId: String(item?.drivesTaskId || (item?.dependencyRelation === 'drives' ? item?.dependencyTaskId : '') || '').trim(),
        dependencyRelation: ['drives', 'drivenBy'].includes(item?.dependencyRelation) ? item.dependencyRelation : '',
        dependencyTaskId: String(item?.dependencyTaskId || '').trim(),
        createdAt: String(item?.createdAt || item?.updatedAt || '').trim(),
        updatedAt: String(item?.updatedAt || item?.createdAt || '').trim(),
        completed: Boolean(item?.completed)
      };
    })
    .filter(item=>{
      if (!item.id) return false;
      const start = parseProjectFlowDate(item.startDate);
      const end = parseProjectFlowDate(item.endDate);
      return Boolean(start && end && start <= end);
    });
}

function loadProjectFlowState(){
  const store = readProjectFlowStore();
  const normalized = {};
  Object.entries(store).forEach(([projectId, items])=>{
    const key = String(projectId || '').trim();
    if (!key) return;
    normalized[key] = normalizeProjectFlowMilestones(items);
  });
  projectFlowState.milestonesByProjectId = normalized;
}

function getProjectFlowMilestones(projectId){
  const key = String(projectId || '').trim();
  return normalizeProjectFlowMilestones(projectFlowState.milestonesByProjectId[key] || []);
}

function setProjectFlowMilestones(projectId, milestones){
  const key = String(projectId || '').trim();
  if (!key) return;
  projectFlowState.milestonesByProjectId[key] = normalizeProjectFlowMilestones(milestones);
  persistProjectFlowStore();
}

function createProjectFlowId(){
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `flow-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function parseProjectFlowDate(value){
  const raw = String(value || '').trim();
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const displayMatch = raw.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/);
  const match = isoMatch || displayMatch;
  if (!match) return null;
  const year = isoMatch ? Number(match[1]) : Number(match[3]);
  const month = isoMatch ? Number(match[2]) : Number(match[2]);
  const day = isoMatch ? Number(match[3]) : Number(match[1]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatProjectFlowDate(date){
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function formatProjectFlowDisplayDate(value){
  const date = parseProjectFlowDate(value);
  if (!date) return '-';
  return date.toLocaleDateString('no-NO', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function formatProjectFlowInputDate(value){
  const date = value instanceof Date ? value : parseProjectFlowDate(value);
  if (!date) return '';
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${date.getFullYear()}`;
}

function getProjectFlowDatePickerTarget(){
  const targetId = String(projectFlowState.datePickerTargetId || '').trim();
  return targetId ? $(targetId) : null;
}

function closeProjectFlowDatePickerPopover(){
  const popover = $('projectFlowDatePickerPopover');
  if (popover) popover.hidden = true;
  projectFlowState.datePickerTargetId = '';
}

function renderProjectFlowDatePickerPopover(){
  const popover = $('projectFlowDatePickerPopover');
  const targetInput = getProjectFlowDatePickerTarget();
  if (!popover || !targetInput) return;
  const selected = parseProjectFlowDate(targetInput.value);
  const cursor = projectFlowState.datePickerCursor instanceof Date
    ? projectFlowState.datePickerCursor
    : selected || new Date();
  const monthStart = startOfMonth(cursor);
  const gridStart = startOfWeekMonday(monthStart);
  const today = startOfDay(new Date());
  popover.innerHTML = '';

  const head = document.createElement('div');
  head.className = 'calendar-date-picker-head';
  const title = document.createElement('div');
  title.className = 'calendar-date-picker-title';
  title.textContent = new Intl.DateTimeFormat('no-NO', { month: 'long', year: 'numeric' }).format(cursor);
  const nav = document.createElement('div');
  nav.className = 'calendar-date-picker-nav';
  const prev = document.createElement('button');
  prev.type = 'button';
  prev.setAttribute('aria-label', 'Forrige måned');
  prev.textContent = '‹';
  prev.addEventListener('click', ()=>{
    projectFlowState.datePickerCursor = addMonths(cursor, -1);
    renderProjectFlowDatePickerPopover();
  });
  const next = document.createElement('button');
  next.type = 'button';
  next.setAttribute('aria-label', 'Neste måned');
  next.textContent = '›';
  next.addEventListener('click', ()=>{
    projectFlowState.datePickerCursor = addMonths(cursor, 1);
    renderProjectFlowDatePickerPopover();
  });
  nav.append(prev, next);
  head.append(title, nav);
  popover.appendChild(head);

  const grid = document.createElement('div');
  grid.className = 'calendar-date-picker-grid';
  ['Ma', 'Ti', 'On', 'To', 'Fr', 'Lø', 'Sø'].forEach(label=>{
    const weekday = document.createElement('div');
    weekday.className = 'calendar-date-picker-weekday';
    weekday.textContent = label;
    grid.appendChild(weekday);
  });
  for (let index = 0; index < 42; index += 1){
    const day = addDays(gridStart, index);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'calendar-date-picker-day';
    btn.classList.toggle('is-outside', day.getMonth() !== cursor.getMonth());
    btn.classList.toggle('is-today', sameCalendarDay(day, today));
    btn.classList.toggle('is-selected', selected ? sameCalendarDay(day, selected) : false);
    btn.textContent = String(day.getDate());
    btn.addEventListener('click', ()=>{
      targetInput.value = formatProjectFlowInputDate(day);
      if (targetInput.id === 'projectFlowEndDateInput'){
        updateProjectFlowDurationFromEndDate();
      } else {
        updateProjectFlowEndDateFromDuration();
      }
      closeProjectFlowDatePickerPopover();
    });
    grid.appendChild(btn);
  }
  popover.appendChild(grid);
}

function openProjectFlowDatePickerPopover(targetId){
  const targetInput = $(targetId);
  const popover = $('projectFlowDatePickerPopover');
  if (!targetInput || !popover) return;
  projectFlowState.datePickerTargetId = targetId;
  projectFlowState.datePickerCursor = parseProjectFlowDate(targetInput.value) || new Date();
  const targetRow = targetInput.closest('.calendar-date-input-row');
  if (targetRow && popover.parentElement !== targetRow){
    targetRow.appendChild(popover);
  }
  renderProjectFlowDatePickerPopover();
  popover.hidden = false;
}

function addProjectFlowDays(date, days){
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  copy.setDate(copy.getDate() + Number(days || 0));
  return copy;
}

function addProjectFlowDuration(startDate, value, unit){
  const amount = Math.max(1, Number.parseInt(value, 10) || 1);
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  if (unit === 'weeks'){
    return addProjectFlowDays(start, (amount * 7) - 1);
  }
  if (unit === 'months'){
    const end = new Date(start.getFullYear(), start.getMonth() + amount, start.getDate());
    return addProjectFlowDays(end, -1);
  }
  return addProjectFlowDays(start, amount - 1);
}

function getProjectFlowDurationFromDates(startDate, endDate){
  const days = Math.max(1, getProjectFlowDayDiff(startDate, endDate) + 1);
  if (days >= 28 && days % 30 === 0) return { value: days / 30, unit: 'months' };
  if (days >= 7 && days % 7 === 0) return { value: days / 7, unit: 'weeks' };
  return { value: days, unit: 'days' };
}

function getProjectFlowDayDiff(start, end){
  const startUtc = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
  const endUtc = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((endUtc - startUtc) / 86400000);
}

function getProjectFlowWeekNumber(date){
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  return Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
}

function getProjectFlowWeekKey(date){
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  return `${utc.getUTCFullYear()}-${String(getProjectFlowWeekNumber(date)).padStart(2, '0')}`;
}

function getProjectFlowWeekSpans(dates){
  const spans = [];
  dates.forEach(date=>{
    const key = getProjectFlowWeekKey(date);
    const weekNumber = getProjectFlowWeekNumber(date);
    const current = spans[spans.length - 1];
    if (current && current.key === key){
      current.days += 1;
    } else {
      spans.push({ key, weekNumber, days: 1 });
    }
  });
  return spans;
}

function isProjectFlowWeekStart(date, index){
  return index > 0 && date.getDay() === 1;
}

function getProjectFlowVisibleProjects(){
  const selectedId = String(projectFlowState.selectedProjectId || PROJECT_FLOW_ALL_PROJECTS).trim();
  if (selectedId && selectedId !== PROJECT_FLOW_ALL_PROJECTS){
    const project = projectState.projects.find(item=>item.id === selectedId);
    return project ? [project] : [];
  }
  return [...projectState.projects];
}

function getProjectFlowAllTasks(){
  return projectState.projects.flatMap(project=>{
    return getProjectFlowMilestones(project.id).map(task=>({
      ...task,
      projectId: project.id,
      projectName: project.name || 'Uten navn',
      projectNumber: project.projectNumber || ''
    }));
  });
}

function getProjectFlowDateKey(date){
  return formatProjectFlowDate(date);
}

function getProjectFlowRange(milestones){
  const today = new Date();
  const currentWeekStart = startOfWeekMonday(today);
  let start = currentWeekStart;
  const taskDates = milestones
    .flatMap(item=>[parseProjectFlowDate(item.startDate), parseProjectFlowDate(item.endDate)])
    .filter(Boolean);
  if (taskDates.length){
    const earliestTaskDate = new Date(Math.min(...taskDates.map(date=>date.getTime())));
    const earliestTaskWeekStart = startOfWeekMonday(earliestTaskDate);
    if (earliestTaskWeekStart < start) start = earliestTaskWeekStart;
  }
  const lastDayOfYear = new Date(today.getFullYear(), 11, 31);
  let end = addProjectFlowDays(startOfWeekMonday(lastDayOfYear), 6);
  const taskEndDates = taskDates;
  if (taskEndDates.length){
    const latestTaskEnd = new Date(Math.max(...taskEndDates.map(date=>date.getTime())));
    if (latestTaskEnd > end){
      end = addProjectFlowDays(startOfWeekMonday(latestTaskEnd), 6);
    }
  }
  return { start, end };
}

function buildProjectFlowDates(start, end){
  const dates = [];
  let cursor = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  while (cursor <= end && dates.length < 900){
    dates.push(new Date(cursor));
    cursor = addProjectFlowDays(cursor, 1);
  }
  return dates;
}

function syncProjectFlowTopScrollbar(scroller, topScrollbar){
  if (!scroller || !topScrollbar) return;
  let syncing = false;
  let middlePan = null;
  let middlePanFrame = 0;
  const applyMiddlePanScroll = ()=>{
    middlePanFrame = 0;
    if (!middlePan) return;
    syncing = true;
    scroller.scrollLeft = Math.max(0, Math.min(scroller.scrollWidth - scroller.clientWidth, middlePan.targetScrollLeft));
    topScrollbar.scrollLeft = scroller.scrollLeft;
    renderProjectFlowDependencyLines();
    syncing = false;
  };
  topScrollbar.addEventListener('scroll', ()=>{
    if (syncing) return;
    syncing = true;
    scroller.scrollLeft = topScrollbar.scrollLeft;
    renderProjectFlowDependencyLines();
    syncing = false;
  });
  scroller.addEventListener('scroll', ()=>{
    if (syncing) return;
    syncing = true;
    topScrollbar.scrollLeft = scroller.scrollLeft;
    renderProjectFlowDependencyLines();
    syncing = false;
  });
  scroller.addEventListener('mousedown', evt=>{
    if (evt.button !== 1) return;
    evt.preventDefault();
    middlePan = {
      startX: evt.clientX,
      startScrollLeft: scroller.scrollLeft,
      targetScrollLeft: scroller.scrollLeft
    };
    scroller.classList.add('is-middle-panning');
  });
  scroller.addEventListener('auxclick', evt=>{
    if (evt.button === 1) evt.preventDefault();
  });
  window.addEventListener('mousemove', evt=>{
    if (!middlePan) return;
    evt.preventDefault();
    middlePan.targetScrollLeft = middlePan.startScrollLeft - (evt.clientX - middlePan.startX);
    if (!middlePanFrame){
      middlePanFrame = window.requestAnimationFrame(applyMiddlePanScroll);
    }
  });
  window.addEventListener('mouseup', evt=>{
    if (evt.button !== 1 || !middlePan) return;
    if (middlePanFrame){
      window.cancelAnimationFrame(middlePanFrame);
      middlePanFrame = 0;
      applyMiddlePanScroll();
    }
    middlePan = null;
    scroller.classList.remove('is-middle-panning');
  });
}

function getProjectFlowSvgPoint(scroller, element, side = 'center'){
  if (!scroller || !(element instanceof Element)) return null;
  const scrollRect = scroller.getBoundingClientRect();
  const rect = element.getBoundingClientRect();
  const x = (side === 'right' ? rect.right : side === 'left' ? rect.left : rect.left + (rect.width / 2)) - scrollRect.left + scroller.scrollLeft;
  const y = rect.top + (rect.height / 2) - scrollRect.top + scroller.scrollTop;
  return { x, y };
}

function snapProjectFlowXToDateCenter(x){
  const taskWidth = getProjectFlowTaskColumnWidth();
  const dayWidth = getProjectFlowDayWidth();
  if (!Number.isFinite(x) || x < taskWidth) return x;
  const dayIndex = Math.max(0, Math.round((x - taskWidth - (dayWidth / 2)) / dayWidth));
  return taskWidth + (dayIndex * dayWidth) + (dayWidth / 2);
}

function snapProjectFlowXAwayFromPoint(originX, side = 'right', minDistance = 0){
  const dayWidth = getProjectFlowDayWidth();
  const direction = side === 'left' ? -1 : 1;
  const minimum = Math.max(8, Number(minDistance) || dayWidth * 0.55);
  let snapped = snapProjectFlowXToDateCenter(originX + (direction * minimum));
  const isOutward = direction > 0 ? snapped > originX : snapped < originX;
  if (!isOutward || Math.abs(snapped - originX) < Math.min(12, dayWidth * 0.25)){
    snapped = snapProjectFlowXToDateCenter(originX + (direction * (minimum + dayWidth)));
  }
  return snapped;
}

function getProjectFlowTaskBarObstacles(scroller, excludeTaskIds = []){
  if (!scroller) return [];
  const excluded = new Set((Array.isArray(excludeTaskIds) ? excludeTaskIds : []).map(id=>String(id || '')));
  const scrollRect = scroller.getBoundingClientRect();
  return Array.from(scroller.querySelectorAll('[data-project-flow-task-bar]'))
    .filter(bar=>!excluded.has(String(bar.getAttribute('data-project-flow-task-bar') || '')))
    .map(bar=>{
      const rect = bar.getBoundingClientRect();
      return {
        left: rect.left - scrollRect.left + scroller.scrollLeft - 8,
        right: rect.right - scrollRect.left + scroller.scrollLeft + 8,
        top: rect.top - scrollRect.top + scroller.scrollTop - 8,
        bottom: rect.bottom - scrollRect.top + scroller.scrollTop + 8
      };
    });
}

function projectFlowRangeIntersects(aStart, aEnd, bStart, bEnd){
  return Math.max(aStart, bStart) <= Math.min(aEnd, bEnd);
}

function projectFlowHorizontalRouteIsClear(y, x1, x2, obstacles){
  const minX = Math.min(x1, x2);
  const maxX = Math.max(x1, x2);
  return !(Array.isArray(obstacles) ? obstacles : []).some(obstacle=>
    obstacle.axis !== 'v'
    &&
    y >= obstacle.top
    && y <= obstacle.bottom
    && projectFlowRangeIntersects(minX, maxX, obstacle.left, obstacle.right)
  );
}

function projectFlowVerticalRouteIsClear(x, y1, y2, obstacles){
  const minY = Math.min(y1, y2);
  const maxY = Math.max(y1, y2);
  return !(Array.isArray(obstacles) ? obstacles : []).some(obstacle=>
    obstacle.axis !== 'h'
    &&
    x >= obstacle.left
    && x <= obstacle.right
    && projectFlowRangeIntersects(minY, maxY, obstacle.top, obstacle.bottom)
  );
}

function getProjectFlowXRouteCandidates(centerX){
  const dayWidth = getProjectFlowDayWidth();
  const base = snapProjectFlowXToDateCenter(centerX);
  const laneOffset = Math.max(8, Math.min(18, dayWidth * 0.22));
  return [
    base,
    base - laneOffset,
    base + laneOffset,
    base - (laneOffset * 2),
    base + (laneOffset * 2)
  ];
}

function findProjectFlowClearRouteY(preferredY, x1, x2, fromY, toY, obstacles){
  const rowStep = 52;
  const candidates = [preferredY, fromY + rowStep, fromY - rowStep, toY + rowStep, toY - rowStep];
  for (let offset = 2; offset <= 10; offset += 1){
    candidates.push(preferredY + (rowStep * offset), preferredY - (rowStep * offset));
  }
  const seen = new Set();
  for (const candidate of candidates){
    const y = Math.max(0, Math.round(candidate));
    if (seen.has(y)) continue;
    seen.add(y);
    if (projectFlowHorizontalRouteIsClear(y, x1, x2, obstacles)){
      return y;
    }
  }
  return preferredY;
}

function findProjectFlowClearRouteX(preferredX, x1, x2, fromY, toY, obstacles){
  const dayWidth = getProjectFlowDayWidth();
  const minX = Math.min(x1, x2);
  const maxX = Math.max(x1, x2);
  const candidates = [...getProjectFlowXRouteCandidates(preferredX)];
  for (let x = snapProjectFlowXToDateCenter(minX); x <= maxX; x += dayWidth){
    candidates.push(...getProjectFlowXRouteCandidates(x));
  }
  for (let offset = 1; offset <= 8; offset += 1){
    candidates.push(
      ...getProjectFlowXRouteCandidates(minX - (dayWidth * offset)),
      ...getProjectFlowXRouteCandidates(maxX + (dayWidth * offset))
    );
  }
  const seen = new Set();
  for (const candidate of candidates){
    const x = Math.round(candidate);
    if (seen.has(x)) continue;
    seen.add(x);
    if (projectFlowVerticalRouteIsClear(x, fromY, toY, obstacles)){
      return x;
    }
  }
  return snapProjectFlowXToDateCenter(preferredX);
}

function buildRoundedProjectFlowPath(points, radius = 10){
  const validPoints = points.filter(point=>Array.isArray(point) && point.length === 2);
  if (validPoints.length < 2) return '';
  const commands = [`M ${validPoints[0][0]} ${validPoints[0][1]}`];
  for (let index = 1; index < validPoints.length - 1; index += 1){
    const prev = validPoints[index - 1];
    const current = validPoints[index];
    const next = validPoints[index + 1];
    const v1 = { x: current[0] - prev[0], y: current[1] - prev[1] };
    const v2 = { x: next[0] - current[0], y: next[1] - current[1] };
    const len1 = Math.hypot(v1.x, v1.y);
    const len2 = Math.hypot(v2.x, v2.y);
    if (!len1 || !len2 || (v1.x && v1.y) || (v2.x && v2.y)){
      commands.push(`L ${current[0]} ${current[1]}`);
      continue;
    }
    const cornerRadius = Math.min(radius, len1 / 2, len2 / 2);
    const before = [
      current[0] - (v1.x / len1) * cornerRadius,
      current[1] - (v1.y / len1) * cornerRadius
    ];
    const after = [
      current[0] + (v2.x / len2) * cornerRadius,
      current[1] + (v2.y / len2) * cornerRadius
    ];
    commands.push(`L ${before[0]} ${before[1]}`);
    commands.push(`Q ${current[0]} ${current[1]} ${after[0]} ${after[1]}`);
  }
  const last = validPoints[validPoints.length - 1];
  commands.push(`L ${last[0]} ${last[1]}`);
  return commands.join(' ');
}

function drawProjectFlowDependencyPath(svg, from, to, className = '', fromSide = 'right', toSide = 'left', obstacles = []){
  if (!svg || !from || !to) return [];
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  const dayWidth = getProjectFlowDayWidth();
  const fromDir = fromSide === 'left' ? -1 : 1;
  const toDir = toSide === 'left' ? -1 : 1;
  const stub = Math.max(28, dayWidth * 0.55);
  const fromOut = snapProjectFlowXAwayFromPoint(from.x, fromSide, stub);
  const toIn = snapProjectFlowXAwayFromPoint(to.x, toSide, stub);
  const hasRoomBetweenStubs = fromDir === 1 && toDir === -1
    ? fromOut < toIn
    : fromDir === -1 && toDir === 1
      ? fromOut > toIn
      : Math.abs(fromOut - toIn) >= dayWidth;
  const points = hasRoomBetweenStubs
    ? (() => {
      const routeX = findProjectFlowClearRouteX((fromOut + toIn) / 2, fromOut, toIn, from.y, to.y, obstacles);
      return [
        [from.x, from.y],
        [fromOut, from.y],
        [routeX, from.y],
        [routeX, to.y],
        [toIn, to.y],
        [to.x, to.y]
      ];
    })()
    : (() => {
      const routeY = findProjectFlowClearRouteY(from.y + ((to.y - from.y) / 2), fromOut, toIn, from.y, to.y, obstacles);
      return [
        [from.x, from.y],
        [fromOut, from.y],
        [fromOut, routeY],
        [toIn, routeY],
        [toIn, to.y],
        [to.x, to.y]
      ];
    })();
  path.setAttribute('d', buildRoundedProjectFlowPath(points, 8));
  path.setAttribute('class', `project-flow-link-line ${className}`.trim());
  svg.appendChild(path);
  return points;
}

function getProjectFlowLineRouteObstacles(routes, projectId){
  const currentProjectId = String(projectId || '');
  return (Array.isArray(routes) ? routes : [])
    .filter(route=>route.projectId !== currentProjectId)
    .flatMap(route=>route.obstacles || []);
}

function buildProjectFlowLineRouteObstacles(points){
  const items = [];
  const validPoints = Array.isArray(points) ? points : [];
  for (let index = 1; index < validPoints.length; index += 1){
    const prev = validPoints[index - 1];
    const current = validPoints[index];
    if (!Array.isArray(prev) || !Array.isArray(current)) continue;
    if (Math.round(prev[1]) === Math.round(current[1])){
      items.push({
        axis: 'h',
        left: Math.min(prev[0], current[0]),
        right: Math.max(prev[0], current[0]),
        top: prev[1] - 7,
        bottom: prev[1] + 7
      });
    } else if (Math.round(prev[0]) === Math.round(current[0])){
      items.push({
        axis: 'v',
        left: prev[0] - 7,
        right: prev[0] + 7,
        top: Math.min(prev[1], current[1]),
        bottom: Math.max(prev[1], current[1])
      });
    }
  }
  return items;
}

function renderProjectFlowDependencyLines(){
  const root = $('projectFlowTimeline');
  const scroller = root?.querySelector?.('.project-flow-scroller');
  const svg = root?.querySelector?.('.project-flow-link-overlay');
  if (!scroller || !svg) return;
  svg.setAttribute('width', String(scroller.scrollWidth));
  svg.setAttribute('height', String(scroller.scrollHeight));
  svg.setAttribute('viewBox', `0 0 ${scroller.scrollWidth} ${scroller.scrollHeight}`);
  svg.innerHTML = '';
  const drawnEdges = new Set();
  const drawnRoutes = [];
  getProjectFlowVisibleProjects().forEach(project=>{
    getProjectFlowMilestones(project.id).forEach(task=>{
      const drivesTaskId = getProjectFlowRelationTaskId(task, 'drives');
      if (!drivesTaskId) return;
      const dependency = findProjectFlowTaskLocation(drivesTaskId);
      if (!dependency?.task) return;
      if (dependency.project.id !== project.id) return;
      const edgeKey = [task.id, dependency.task.id].sort().join('|');
      if (drawnEdges.has(edgeKey)) return;
      drawnEdges.add(edgeKey);
      const taskBar = scroller.querySelector(`[data-project-flow-task-bar="${projectFlowCssEscape(task.id)}"][data-project-id="${projectFlowCssEscape(project.id)}"]`);
      const dependencyBar = scroller.querySelector(`[data-project-flow-task-bar="${projectFlowCssEscape(dependency.task.id)}"][data-project-id="${projectFlowCssEscape(dependency.project.id)}"]`);
      if (!taskBar || !dependencyBar) return;
      const obstacles = [
        ...getProjectFlowTaskBarObstacles(scroller, [task.id, dependency.task.id]),
        ...getProjectFlowLineRouteObstacles(drawnRoutes, project.id)
      ];
      let points = [];
      points = drawProjectFlowDependencyPath(svg, getProjectFlowSvgPoint(scroller, taskBar, 'right'), getProjectFlowSvgPoint(scroller, dependencyBar, 'left'), '', 'right', 'left', obstacles);
      if (points.length){
        drawnRoutes.push({
          projectId: project.id,
          obstacles: buildProjectFlowLineRouteObstacles(points)
        });
      }
    });
  });
  const drag = projectFlowState.linkDrag;
  if (drag?.source){
    const sourceBar = scroller.querySelector(`[data-project-flow-task-bar="${projectFlowCssEscape(drag.source.taskId)}"][data-project-id="${projectFlowCssEscape(drag.source.projectId)}"]`);
    const fromSide = drag.source.type === 'drives' ? 'right' : 'left';
    const from = getProjectFlowSvgPoint(scroller, sourceBar, fromSide);
    const scrollRect = scroller.getBoundingClientRect();
    const to = {
      x: drag.x - scrollRect.left + scroller.scrollLeft,
      y: drag.y - scrollRect.top + scroller.scrollTop
    };
    const obstacles = getProjectFlowTaskBarObstacles(scroller, [drag.source.taskId]);
    drawProjectFlowDependencyPath(svg, from, to, 'is-preview', fromSide, fromSide === 'right' ? 'left' : 'right', obstacles);
  }
}

function scrollProjectFlowToCurrentWeek(scroller, rangeStart, topScrollbar = null){
  if (!scroller || !(rangeStart instanceof Date)) return;
  const currentWeekStart = startOfWeekMonday(new Date());
  const offsetDays = Math.max(0, getProjectFlowDayDiff(rangeStart, currentWeekStart));
  const dayWidth = getProjectFlowDayWidth();
  requestAnimationFrame(()=>{
    const nextScrollLeft = Math.max(0, (offsetDays * dayWidth) - 12);
    scroller.scrollLeft = nextScrollLeft;
    if (topScrollbar) topScrollbar.scrollLeft = nextScrollLeft;
  });
}

function getProjectFlowSelectedProject(){
  const selectedId = String(projectFlowState.selectedProjectId || '').trim();
  if (selectedId === PROJECT_FLOW_ALL_PROJECTS) return null;
  return projectState.projects.find(project=>project.id === selectedId) || projectState.projects[0] || null;
}

function updateProjectFlowProjectSelect(){
  const select = $('projectFlowProjectSelect');
  if (!select) return;
  const current = String(projectFlowState.selectedProjectId || '').trim();
  select.innerHTML = '';
  if (!projectState.projects.length){
    const option = document.createElement('option');
    option.value = '';
    option.textContent = 'Ingen prosjekter';
    select.appendChild(option);
    projectFlowState.selectedProjectId = '';
    select.disabled = true;
    return;
  }
  const allOption = document.createElement('option');
  allOption.value = PROJECT_FLOW_ALL_PROJECTS;
  allOption.textContent = 'Alle prosjekter';
  select.appendChild(allOption);
  projectState.projects.forEach(project=>{
    const option = document.createElement('option');
    option.value = project.id || '';
    option.textContent = project.projectNumber
      ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
      : project.name || 'Uten navn';
    select.appendChild(option);
  });
  const exists = current === PROJECT_FLOW_ALL_PROJECTS || projectState.projects.some(project=>project.id === current);
  projectFlowState.selectedProjectId = exists ? current : PROJECT_FLOW_ALL_PROJECTS;
  select.value = projectFlowState.selectedProjectId;
  select.disabled = false;
}

function populateProjectFlowTaskProjectSelect(selectedProjectId = '', options = {}){
  const select = $('projectFlowTaskProjectSelect');
  if (!select) return;
  select.innerHTML = '';
  const emptyOption = document.createElement('option');
  emptyOption.value = '';
  emptyOption.textContent = 'Velg prosjekt';
  select.appendChild(emptyOption);
  const excludedProjectIds = new Set((Array.isArray(options.excludeProjectIds) ? options.excludeProjectIds : []).map(id=>String(id || '')));
  const selectableProjects = projectState.projects.filter(project=>{
    if (!project?.id) return false;
    if (selectedProjectId && project.id === selectedProjectId) return true;
    return !excludedProjectIds.has(project.id);
  });
  selectableProjects.forEach(project=>{
    const option = document.createElement('option');
    option.value = project.id || '';
    option.textContent = project.projectNumber
      ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
      : project.name || 'Uten navn';
    select.appendChild(option);
  });
  select.value = selectableProjects.some(project=>project.id === selectedProjectId) ? selectedProjectId : '';
  select.disabled = !projectState.projects.length;
}

function populateProjectFlowPhaseSelect(){
  const select = $('projectFlowPhaseSelect');
  if (!select || select.options.length) return;
  PROJECT_FLOW_PHASES.forEach(phase=>{
    const option = document.createElement('option');
    option.value = phase.id;
    option.textContent = phase.label;
    select.appendChild(option);
  });
}

function getProjectFlowDependencyOptionLabel(task){
  const projectLabel = task.projectNumber
    ? `${task.projectNumber} - ${task.projectName}`
    : task.projectName;
  return `${projectLabel}: ${getProjectFlowPhaseLabel(task.phaseId)}`;
}

function populateProjectFlowDependencySelect(select, selectedTaskId = '', currentTaskId = '', relation = '', currentProjectId = ''){
  if (!select) return;
  select.innerHTML = '';
  const emptyOption = document.createElement('option');
  emptyOption.value = '';
  emptyOption.textContent = 'Ingen oppgave';
  select.appendChild(emptyOption);
  getProjectFlowAllTasks()
    .filter(task=>task.id !== currentTaskId)
    .filter(task=>!currentProjectId || task.projectId === currentProjectId)
    .filter(task=>!wouldCreateProjectFlowMutualDependency(currentTaskId, relation, task.id))
    .forEach(task=>{
      const option = document.createElement('option');
      option.value = task.id;
      option.textContent = getProjectFlowDependencyOptionLabel(task);
      select.appendChild(option);
    });
  select.value = selectedTaskId || '';
  if (select.value !== selectedTaskId) select.value = '';
}

function populateProjectFlowDependencyTaskSelect(milestone = null){
  const drivenBySelect = $('projectFlowDependencyRelationSelect');
  const drivesSelect = $('projectFlowDependencyTaskSelect');
  const currentTaskId = milestone?.id || '';
  const drivenBySelected = getProjectFlowRelationTaskId(milestone, 'drivenBy');
  const drivesSelected = getProjectFlowRelationTaskId(milestone, 'drives');
  const currentProjectId = String(milestone?.projectId || projectFlowState.editingProjectId || $('projectFlowTaskProjectSelect')?.value || '').trim();
  populateProjectFlowDependencySelect(drivenBySelect, drivenBySelected, currentTaskId, 'drivenBy', currentProjectId);
  populateProjectFlowDependencySelect(drivesSelect, drivesSelected, currentTaskId, 'drives', currentProjectId);
}

function setProjectFlowStatus(message, type = ''){
  const el = $('projectFlowStatus');
  if (!el) return;
  el.textContent = message || '';
  el.classList.toggle('ok', type === 'ok');
  el.classList.toggle('error', type === 'error');
}

function getProjectFlowDayWidth(){
  const fit = Number(projectFlowState.fitDayWidth);
  if (Number.isFinite(fit) && fit > 0) return fit;
  return PROJECT_FLOW_ZOOM_LEVELS[projectFlowState.zoomIndex] || PROJECT_FLOW_ZOOM_LEVELS[PROJECT_FLOW_DEFAULT_ZOOM_INDEX];
}

function getProjectFlowTaskColumnWidth(){
  const width = Number(projectFlowState.taskColumnWidth);
  return Number.isFinite(width) && width > 0 ? width : 260;
}

function measureProjectFlowTaskColumnWidth(labels){
  const values = (Array.isArray(labels) ? labels : [])
    .map(label=>String(label || '').trim())
    .filter(Boolean);
  let maxWidth = 0;
  const canvas = document.createElement('canvas');
  const context = canvas.getContext?.('2d');
  if (context){
    context.font = '700 16px Arial, sans-serif';
    values.forEach(label=>{
      maxWidth = Math.max(maxWidth, context.measureText(label).width);
    });
  } else {
    values.forEach(label=>{
      maxWidth = Math.max(maxWidth, label.length * 8);
    });
  }
  return Math.ceil(Math.max(260, Math.min(900, maxWidth + 96)));
}

function setProjectFlowZoomIndex(nextIndex){
  const clamped = Math.max(0, Math.min(PROJECT_FLOW_ZOOM_LEVELS.length - 1, Number(nextIndex)));
  projectFlowState.zoomIndex = Number.isFinite(clamped) ? clamped : PROJECT_FLOW_DEFAULT_ZOOM_INDEX;
  projectFlowState.fitDayWidth = null;
  renderProjectFlowView();
}

function getProjectFlowTasksByPhase(tasks){
  const grouped = new Map(PROJECT_FLOW_PHASES.map(phase=>[phase.id, []]));
  (Array.isArray(tasks) ? tasks : []).forEach(task=>{
    const phaseId = PROJECT_FLOW_PHASES.some(phase=>phase.id === task.phaseId) ? task.phaseId : PROJECT_FLOW_PHASES[0].id;
    const list = grouped.get(phaseId) || [];
    list.push(task);
    grouped.set(phaseId, list);
  });
  grouped.forEach(list=>{
    list.sort((a, b)=>String(a.startDate || '').localeCompare(String(b.startDate || ''), 'no'));
  });
  return grouped;
}

function getProjectFlowProjectLabel(project){
  if (!project) return 'Ukjent prosjekt';
  return project.projectNumber
    ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
    : project.name || 'Uten navn';
}

function getProjectFlowPhaseLabel(phaseId){
  return PROJECT_FLOW_PHASES.find(phase=>phase.id === phaseId)?.label || PROJECT_FLOW_PHASES[0].label;
}

function getProjectFlowTaskLabel(item, includeProject = false){
  if (includeProject){
    return item.projectNumber
      ? `${item.projectNumber} - ${item.projectName || 'Uten navn'}`
      : item.projectName || 'Uten navn';
  }
  return getProjectFlowPhaseLabel(item.phaseId);
}

function getProjectFlowTaskActivityTime(task){
  const explicit = new Date(task?.updatedAt || task?.createdAt || 0).getTime();
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  const end = parseProjectFlowDate(task?.endDate);
  if (end) return end.getTime();
  const start = parseProjectFlowDate(task?.startDate);
  return start ? start.getTime() : 0;
}

function formatProjectFlowStatusHours(milliseconds){
  return `${Math.max(1, Math.floor(milliseconds / 3600000))} t`;
}

function formatProjectFlowStatusDays(milliseconds){
  return `${Math.max(1, Math.floor(milliseconds / 86400000))} d`;
}

function getProjectFlowStatusForProject(project){
  const projectId = typeof project === 'object' ? project?.id : project;
  const tasks = getProjectFlowMilestones(projectId);
  const tasksByPhase = getProjectFlowTasksByPhase(tasks);
  const firstPhaseTasks = tasksByPhase.get(PROJECT_FLOW_PHASES[0].id) || [];
  if (!firstPhaseTasks.some(task=>task.completed)){
    const projectTime = new Date(project?.createdAt || 0).getTime();
    const untreatedAge = Number.isFinite(projectTime) && projectTime > 0 ? Date.now() - projectTime : 0;
    const stale = untreatedAge >= 24 * 60 * 60 * 1000;
    return {
      label: 'Ubehandlet',
      tone: stale ? 'danger' : 'idle',
      ageText: stale ? formatProjectFlowStatusHours(untreatedAge) : ''
    };
  }
  let lastCompletedPhase = PROJECT_FLOW_PHASES[0];
  for (const phase of PROJECT_FLOW_PHASES.slice(1)){
    const phaseTasks = tasksByPhase.get(phase.id) || [];
    if (!phaseTasks.length) break;
    const hasOpenTask = phaseTasks.some(task=>!task.completed);
    if (hasOpenTask){
      const latestActivity = Math.max(...phaseTasks.filter(task=>!task.completed).map(getProjectFlowTaskActivityTime));
      const age = Number.isFinite(latestActivity) && latestActivity > 0 ? Date.now() - latestActivity : 0;
      const stale = age >= 7 * 24 * 60 * 60 * 1000;
      return {
        label: phase.label,
        tone: stale ? 'danger' : 'warning',
        ageText: stale ? formatProjectFlowStatusDays(age) : ''
      };
    }
    lastCompletedPhase = phase;
  }
  return { label: lastCompletedPhase.label, tone: 'success' };
}

function setProjectFlowAllExpanded(expanded){
  if (expanded){
    projectFlowState.collapsedPhaseIds.clear();
  } else {
    projectFlowState.collapsedPhaseIds = new Set(PROJECT_FLOW_PHASES.map(phase=>phase.id));
  }
  renderProjectFlowView();
}

function setProjectFlowZoomToFit(){
  const root = $('projectFlowTimeline');
  const scroller = root?.querySelector?.('.project-flow-scroller');
  const grid = root?.querySelector?.('.project-flow-grid');
  const dayCount = Number(grid?.style?.getPropertyValue('--project-flow-days') || 0);
  if (!scroller || !dayCount){
    projectFlowState.fitDayWidth = null;
    renderProjectFlowView();
    return;
  }
  const taskWidth = getProjectFlowTaskColumnWidth();
  const available = Math.max(240, scroller.clientWidth - taskWidth - 24);
  projectFlowState.fitDayWidth = Math.max(26, Math.min(96, Math.floor(available / dayCount)));
  renderProjectFlowView();
}

function openProjectFlowMilestoneForm(milestone = null, options = {}){
  const form = $('projectFlowMilestoneForm');
  if (!form) return;
  populateProjectFlowPhaseSelect();
  const selectedProjectId = options.projectId || milestone?.projectId || '';
  populateProjectFlowTaskProjectSelect(selectedProjectId, {
    excludeProjectIds: options.excludeProjectIds
  });
  populateProjectFlowDependencyTaskSelect(milestone);
  projectFlowState.editingMilestoneId = milestone?.id || '';
  projectFlowState.editingProjectId = selectedProjectId;
  const idInput = $('projectFlowMilestoneId');
  const clickedDateInput = $('projectFlowClickedDateInput');
  const phaseSelect = $('projectFlowPhaseSelect');
  const dateInput = $('projectFlowDateInput');
  const endDateInput = $('projectFlowEndDateInput');
  const durationValueInput = $('projectFlowDurationValueInput');
  const durationUnitSelect = $('projectFlowDurationUnitSelect');
  const fileInput = $('projectFlowFileInput');
  const deleteBtn = $('deleteProjectFlowMilestoneBtn');
  const clickedDate = options.clickedDate || '';
  const originalStart = parseProjectFlowDate(milestone?.startDate || milestone?.date || '');
  const originalEnd = parseProjectFlowDate(milestone?.endDate || milestone?.date || '');
  const originalDuration = originalStart && originalEnd
    ? getProjectFlowDurationFromDates(originalStart, originalEnd)
    : { value: 1, unit: 'days' };
  const fallbackStart = clickedDate || milestone?.startDate || milestone?.date || formatProjectFlowDate(new Date());
  const parsedStart = parseProjectFlowDate(fallbackStart);
  const fallbackEnd = clickedDate && parsedStart
    ? formatProjectFlowDate(addProjectFlowDuration(parsedStart, milestone?.durationValue || originalDuration.value, milestone?.durationUnit || originalDuration.unit))
    : (milestone?.endDate || milestone?.date || fallbackStart);
  const parsedEnd = parseProjectFlowDate(fallbackEnd);
  const duration = parsedStart && parsedEnd
    ? getProjectFlowDurationFromDates(parsedStart, parsedEnd)
    : originalDuration;
  if (idInput) idInput.value = milestone?.id || '';
  if (clickedDateInput) clickedDateInput.value = clickedDate;
  if (phaseSelect) phaseSelect.value = milestone?.phaseId || options.phaseId || PROJECT_FLOW_PHASES[0].id;
  if (dateInput) dateInput.value = formatProjectFlowInputDate(fallbackStart);
  if (endDateInput) endDateInput.value = formatProjectFlowInputDate(fallbackEnd);
  if (durationValueInput) durationValueInput.value = String(milestone?.durationValue || duration.value || 1);
  if (durationUnitSelect) durationUnitSelect.value = milestone?.durationUnit || duration.unit || 'days';
  if (fileInput) fileInput.value = '';
  if (deleteBtn) deleteBtn.hidden = !milestone?.id;
  form.hidden = false;
  openFormModal('projectFlowMilestoneForm', milestone?.id ? 'Endre oppgave' : 'Ny oppgave');
  dateInput?.focus();
}

function updateProjectFlowEndDateFromDuration(){
  const startInput = $('projectFlowDateInput');
  const endInput = $('projectFlowEndDateInput');
  const valueInput = $('projectFlowDurationValueInput');
  const unitSelect = $('projectFlowDurationUnitSelect');
  const startDate = parseProjectFlowDate(startInput?.value || '');
  if (!startDate || !endInput) return;
  const endDate = addProjectFlowDuration(startDate, valueInput?.value || 1, unitSelect?.value || 'days');
  endInput.value = formatProjectFlowInputDate(endDate);
}

function updateProjectFlowDurationFromEndDate(){
  const startDate = parseProjectFlowDate($('projectFlowDateInput')?.value || '');
  const endDate = parseProjectFlowDate($('projectFlowEndDateInput')?.value || '');
  const valueInput = $('projectFlowDurationValueInput');
  const unitSelect = $('projectFlowDurationUnitSelect');
  if (!startDate || !endDate || endDate < startDate || !valueInput || !unitSelect) return;
  valueInput.value = String(getProjectFlowDayDiff(startDate, endDate) + 1);
  unitSelect.value = 'days';
}

function closeProjectFlowMilestoneForm(){
  const form = $('projectFlowMilestoneForm');
  if (!form) return;
  closeProjectFlowDatePickerPopover();
  closeFormModal('projectFlowMilestoneForm');
  form.reset();
  projectFlowState.editingMilestoneId = '';
  projectFlowState.editingProjectId = '';
}

function findProjectFlowTaskLocation(taskId){
  const wanted = String(taskId || '').trim();
  if (!wanted) return null;
  for (const project of projectState.projects){
    const tasks = getProjectFlowMilestones(project.id);
    const task = tasks.find(item=>item.id === wanted);
    if (task) return { project, task, tasks };
  }
  return null;
}

function shiftProjectFlowTaskStart(task, startDate){
  const parsedStart = parseProjectFlowDate(startDate);
  if (!parsedStart) return task;
  const durationValue = Math.max(1, Number.parseInt(task.durationValue, 10) || 1);
  const durationUnit = ['days', 'weeks', 'months'].includes(task.durationUnit) ? task.durationUnit : 'days';
  const endDate = formatProjectFlowDate(addProjectFlowDuration(parsedStart, durationValue, durationUnit));
  return {
    ...task,
    startDate: formatProjectFlowDate(parsedStart),
    endDate,
    date: formatProjectFlowDate(parsedStart),
    updatedAt: new Date().toISOString()
  };
}

function shiftProjectFlowTaskByDays(task, deltaDays){
  const startDate = parseProjectFlowDate(task?.startDate);
  const endDate = parseProjectFlowDate(task?.endDate);
  const delta = Number.parseInt(deltaDays, 10) || 0;
  if (!task || !startDate || !endDate || !delta) return task;
  return {
    ...task,
    startDate: formatProjectFlowDate(addProjectFlowDays(startDate, delta)),
    endDate: formatProjectFlowDate(addProjectFlowDays(endDate, delta)),
    date: formatProjectFlowDate(addProjectFlowDays(startDate, delta)),
    updatedAt: new Date().toISOString()
  };
}

function propagateProjectFlowDrivenTaskOffsets(projectId, taskId, moveDeltaDays = 0, durationDeltaDays = 0, visited = new Set()){
  const sourceKey = String(taskId || '');
  if (!projectId || !sourceKey || visited.has(sourceKey)) return;
  visited.add(sourceKey);
  const moveDelta = Number.parseInt(moveDeltaDays, 10) || 0;
  const durationDelta = Number.parseInt(durationDeltaDays, 10) || 0;
  const totalDelta = moveDelta + durationDelta;
  if (!totalDelta) return;
  const milestones = getProjectFlowMilestones(projectId);
  const drivenTasks = milestones.filter(item=>getProjectFlowRelationTaskId(item, 'drivenBy') === sourceKey);
  drivenTasks.forEach(drivenTask=>{
    const shiftedTask = shiftProjectFlowTaskByDays(drivenTask, totalDelta);
    if (shiftedTask === drivenTask) return;
    const latestMilestones = getProjectFlowMilestones(projectId);
    setProjectFlowMilestones(projectId, latestMilestones.map(item=>item.id === drivenTask.id ? shiftedTask : item));
    propagateProjectFlowDrivenTaskOffsets(projectId, drivenTask.id, totalDelta, 0, visited);
  });
}

function getProjectFlowDateChangeDeltas(previousTask, nextStartDate, nextEndDate){
  const previousStart = parseProjectFlowDate(previousTask?.startDate);
  const previousEnd = parseProjectFlowDate(previousTask?.endDate);
  const nextStart = parseProjectFlowDate(nextStartDate);
  const nextEnd = parseProjectFlowDate(nextEndDate);
  if (!previousStart || !previousEnd || !nextStart || !nextEnd) return { moveDelta: 0, durationDelta: 0 };
  const startDelta = getProjectFlowDayDiff(previousStart, nextStart);
  const endDelta = getProjectFlowDayDiff(previousEnd, nextEnd);
  const previousDuration = getProjectFlowDayDiff(previousStart, previousEnd) + 1;
  const nextDuration = getProjectFlowDayDiff(nextStart, nextEnd) + 1;
  const isPureMove = startDelta === endDelta && previousDuration === nextDuration;
  return {
    moveDelta: isPureMove ? startDelta : 0,
    durationDelta: isPureMove ? 0 : endDelta
  };
}

function applyProjectFlowDrivenByDependency(task){
  return task;
}

function applyProjectFlowDrivesDependency(projectId, task){
  return { projectId, task };
}

function getProjectFlowRelationTaskId(task, relation){
  if (!task) return '';
  if (relation === 'drivenBy'){
    return String(task.drivenByTaskId || (task.dependencyRelation === 'drivenBy' ? task.dependencyTaskId : '') || '').trim();
  }
  if (relation === 'drives'){
    return String(task.drivesTaskId || (task.dependencyRelation === 'drives' ? task.dependencyTaskId : '') || '').trim();
  }
  return '';
}

function setProjectFlowRelationOnTask(task, relation, taskId){
  const nextTaskId = String(taskId || '').trim();
  const next = { ...task, updatedAt: new Date().toISOString() };
  if (relation === 'drivenBy'){
    next.drivenByTaskId = nextTaskId;
  } else if (relation === 'drives'){
    next.drivesTaskId = nextTaskId;
  }
  if (next.drivenByTaskId){
    next.dependencyRelation = 'drivenBy';
    next.dependencyTaskId = next.drivenByTaskId;
  } else if (next.drivesTaskId){
    next.dependencyRelation = 'drives';
    next.dependencyTaskId = next.drivesTaskId;
  } else {
    next.dependencyRelation = '';
    next.dependencyTaskId = '';
  }
  return next;
}

function updateProjectFlowTaskDates(projectId, taskId, startDate, endDate){
  const parsedStart = parseProjectFlowDate(startDate);
  const parsedEnd = parseProjectFlowDate(endDate);
  if (!projectId || !taskId || !parsedStart || !parsedEnd || parsedEnd < parsedStart) return false;
  const milestones = getProjectFlowMilestones(projectId);
  const previousTask = milestones.find(item=>item.id === taskId);
  let updatedTask = null;
  const updated = milestones.map(item=>{
    if (item.id !== taskId) return item;
    const duration = getProjectFlowDurationFromDates(parsedStart, parsedEnd);
    updatedTask = {
      ...item,
      startDate: formatProjectFlowDate(parsedStart),
      endDate: formatProjectFlowDate(parsedEnd),
      date: formatProjectFlowDate(parsedStart),
      durationValue: duration.value,
      durationUnit: duration.unit,
      updatedAt: new Date().toISOString()
    };
    return updatedTask;
  });
  if (!updatedTask) return false;
  setProjectFlowMilestones(projectId, updated);
  const deltas = getProjectFlowDateChangeDeltas(previousTask, updatedTask.startDate, updatedTask.endDate);
  propagateProjectFlowDrivenTaskOffsets(projectId, taskId, deltas.moveDelta, deltas.durationDelta);
  return true;
}

function moveProjectFlowTaskByDays(projectId, taskId, deltaDays){
  const location = findProjectFlowTaskLocation(taskId);
  if (!location?.task || location.project.id !== projectId) return false;
  const startDate = parseProjectFlowDate(location.task.startDate);
  const endDate = parseProjectFlowDate(location.task.endDate);
  if (!startDate || !endDate) return false;
  return updateProjectFlowTaskDates(
    projectId,
    taskId,
    formatProjectFlowDate(addProjectFlowDays(startDate, deltaDays)),
    formatProjectFlowDate(addProjectFlowDays(endDate, deltaDays))
  );
}

function resizeProjectFlowTaskByDays(projectId, taskId, edge, deltaDays){
  const location = findProjectFlowTaskLocation(taskId);
  if (!location?.task || location.project.id !== projectId) return false;
  const startDate = parseProjectFlowDate(location.task.startDate);
  const endDate = parseProjectFlowDate(location.task.endDate);
  if (!startDate || !endDate) return false;
  let nextStart = startDate;
  let nextEnd = endDate;
  if (edge === 'start'){
    nextStart = addProjectFlowDays(startDate, deltaDays);
    if (nextStart > endDate) nextStart = endDate;
  } else if (edge === 'end'){
    nextEnd = addProjectFlowDays(endDate, deltaDays);
    if (nextEnd < startDate) nextEnd = startDate;
  } else {
    return false;
  }
  return updateProjectFlowTaskDates(projectId, taskId, formatProjectFlowDate(nextStart), formatProjectFlowDate(nextEnd));
}

function getProjectFlowReciprocalRelation(relation){
  if (relation === 'drivenBy') return 'drives';
  if (relation === 'drives') return 'drivenBy';
  return '';
}

function wouldCreateProjectFlowMutualDependency(taskId, relation, dependencyTaskId){
  const current = findProjectFlowTaskLocation(taskId);
  const dependency = findProjectFlowTaskLocation(dependencyTaskId);
  if (!dependency?.task) return false;
  const reciprocalRelation = getProjectFlowReciprocalRelation(relation);
  if (current?.task && getProjectFlowRelationTaskId(current.task, reciprocalRelation) === dependencyTaskId) return true;
  return getProjectFlowRelationTaskId(dependency.task, relation) === taskId;
}

function clearProjectFlowReciprocalDependency(taskId, relation, dependencyTaskId){
  const reciprocalRelation = getProjectFlowReciprocalRelation(relation);
  if (!taskId || !reciprocalRelation || !dependencyTaskId) return;
  const dependency = findProjectFlowTaskLocation(dependencyTaskId);
  if (!dependency?.project?.id || !dependency.task) return;
  if (getProjectFlowRelationTaskId(dependency.task, reciprocalRelation) !== taskId) return;
  const updated = dependency.tasks.map(item=>item.id === dependency.task.id
    ? setProjectFlowRelationOnTask(item, reciprocalRelation, '')
    : item
  );
  setProjectFlowMilestones(dependency.project.id, updated);
}

function setProjectFlowReciprocalDependency(taskId, relation, dependencyTaskId){
  const reciprocalRelation = getProjectFlowReciprocalRelation(relation);
  if (!taskId || !reciprocalRelation || !dependencyTaskId) return;
  const dependency = findProjectFlowTaskLocation(dependencyTaskId);
  if (!dependency?.project?.id || !dependency.task) return;
  const previousTaskId = getProjectFlowRelationTaskId(dependency.task, reciprocalRelation);
  if (previousTaskId && previousTaskId !== taskId){
    clearProjectFlowReciprocalDependency(dependencyTaskId, reciprocalRelation, previousTaskId);
  }
  const latest = findProjectFlowTaskLocation(dependencyTaskId) || dependency;
  const updated = latest.tasks.map(item=>item.id === latest.task.id
    ? setProjectFlowRelationOnTask(item, reciprocalRelation, taskId)
    : item
  );
  setProjectFlowMilestones(latest.project.id, updated);
}

function setProjectFlowTaskDependency(projectId, taskId, relation, dependencyTaskId){
  if (!projectId || !taskId || !dependencyTaskId || taskId === dependencyTaskId) return false;
  const dependency = findProjectFlowTaskLocation(dependencyTaskId);
  if (!dependency?.task) return false;
  if (dependency.project.id !== projectId){
    setProjectFlowStatus('Oppgaver kan bare kobles mot oppgaver i samme prosjekt.', 'error');
    return false;
  }
  if (wouldCreateProjectFlowMutualDependency(taskId, relation, dependencyTaskId)){
    setProjectFlowStatus('Oppgaver kan ikke styre og styres av hverandre samtidig.', 'error');
    return false;
  }
  const milestones = getProjectFlowMilestones(projectId);
  let nextTask = null;
  const previousTask = milestones.find(item=>item.id === taskId);
  const previousDependencyTaskId = getProjectFlowRelationTaskId(previousTask, relation);
  const updated = milestones.map(item=>{
    if (item.id !== taskId) return item;
    nextTask = setProjectFlowRelationOnTask(item, relation, dependencyTaskId);
    applyProjectFlowDrivenByDependency(nextTask);
    return nextTask;
  });
  if (!nextTask) return false;
  setProjectFlowMilestones(projectId, updated);
  if (previousDependencyTaskId && previousDependencyTaskId !== dependencyTaskId){
    clearProjectFlowReciprocalDependency(taskId, relation, previousDependencyTaskId);
  }
  setProjectFlowReciprocalDependency(taskId, relation, dependencyTaskId);
  applyProjectFlowDrivesDependency(projectId, nextTask);
  return true;
}

function connectProjectFlowTasks(source, target){
  if (!source || !target || source.taskId === target.taskId) return false;
  if (source.projectId !== target.projectId){
    setProjectFlowStatus('Oppgaver kan bare kobles mot oppgaver i samme prosjekt.', 'error');
    return false;
  }
  const sourceType = source.type;
  const targetType = target.type;
  if (sourceType === targetType) return false;
  if (!canConnectProjectFlowLinkHandles(source, target)){
    setProjectFlowStatus('Oppgaver kan ikke styre og styres av hverandre samtidig.', 'error');
    return false;
  }
  if (sourceType === 'drives' && targetType === 'drivenBy'){
    return setProjectFlowTaskDependency(source.projectId, source.taskId, 'drives', target.taskId);
  }
  if (sourceType === 'drivenBy' && targetType === 'drives'){
    return setProjectFlowTaskDependency(source.projectId, source.taskId, 'drivenBy', target.taskId);
  }
  return false;
}

function getProjectFlowLinkHandleData(handle){
  if (!(handle instanceof Element)) return null;
  return {
    projectId: handle.getAttribute('data-project-id') || '',
    taskId: handle.getAttribute('data-project-flow-link-task') || '',
    type: handle.getAttribute('data-project-flow-link') || ''
  };
}

function canConnectProjectFlowLinkHandles(source, target){
  if (!source || !target || source.taskId === target.taskId) return false;
  if (source.projectId !== target.projectId) return false;
  if (source.type === target.type) return false;
  const relation = source.type === 'drivenBy' ? 'drivenBy' : 'drives';
  return !wouldCreateProjectFlowMutualDependency(source.taskId, relation, target.taskId);
}

function updateProjectFlowLinkTargetHighlights(){
  const root = $('projectFlowTimeline');
  const drag = projectFlowState.linkDrag;
  const handles = Array.from(root?.querySelectorAll?.('[data-project-flow-link]') || []);
  handles.forEach(handle=>{
    const data = getProjectFlowLinkHandleData(handle);
    const valid = Boolean(drag?.source && canConnectProjectFlowLinkHandles(drag.source, data));
    handle.classList.toggle('is-valid-link-target', valid);
  });
}

function clearProjectFlowLinkTargetHighlights(){
  const root = $('projectFlowTimeline');
  Array.from(root?.querySelectorAll?.('[data-project-flow-link]') || []).forEach(handle=>{
    handle.classList.remove('is-valid-link-target');
  });
}

function projectFlowCssEscape(value){
  const raw = String(value || '');
  if (window.CSS?.escape) return window.CSS.escape(raw);
  return raw.replace(/["\\]/g, '\\$&');
}

function startProjectFlowLinkDrag(evt, handle){
  if (!(evt instanceof PointerEvent) || !(handle instanceof HTMLElement) || evt.button !== 0) return;
  const data = getProjectFlowLinkHandleData(handle);
  if (!data?.projectId || !data.taskId || !['drives', 'drivenBy'].includes(data.type)) return;
  evt.preventDefault();
  handle.setPointerCapture?.(evt.pointerId);
  projectFlowState.linkDrag = {
    pointerId: evt.pointerId,
    source: data,
    handle,
    x: evt.clientX,
    y: evt.clientY
  };
  handle.classList.add('is-linking');
  updateProjectFlowLinkTargetHighlights();
  renderProjectFlowDependencyLines();
}

function updateProjectFlowLinkDrag(evt){
  const drag = projectFlowState.linkDrag;
  if (!drag || evt.pointerId !== drag.pointerId) return;
  drag.x = evt.clientX;
  drag.y = evt.clientY;
  updateProjectFlowLinkTargetHighlights();
  renderProjectFlowDependencyLines();
  evt.preventDefault();
}

function finishProjectFlowLinkDrag(evt){
  const drag = projectFlowState.linkDrag;
  if (!drag || evt.pointerId !== drag.pointerId) return;
  drag.handle?.releasePointerCapture?.(evt.pointerId);
  drag.handle?.classList?.remove('is-linking');
  const targetEl = document.elementFromPoint(evt.clientX, evt.clientY)?.closest?.('[data-project-flow-link]');
  const target = getProjectFlowLinkHandleData(targetEl);
  projectFlowState.linkDrag = null;
  clearProjectFlowLinkTargetHighlights();
  if (connectProjectFlowTasks(drag.source, target)){
    renderProjectFlowView();
    setProjectFlowStatus('Oppgavekobling lagret.', 'ok');
  } else {
    renderProjectFlowDependencyLines();
  }
  projectFlowState.suppressClickUntil = Date.now() + 350;
}

function startProjectFlowTaskDrag(evt, bar){
  if (!(evt instanceof PointerEvent) || !(bar instanceof HTMLElement)) return;
  if (evt.button !== 0) return;
  const resizeTarget = evt.target instanceof Element ? evt.target.closest('[data-project-flow-resize]') : null;
  if (evt.target instanceof Element && !resizeTarget && evt.target.closest('input,[data-project-flow-delete]')) return;
  const projectId = bar.getAttribute('data-project-id') || '';
  const taskId = bar.getAttribute('data-project-flow-drag') || '';
  const mode = resizeTarget
    ? resizeTarget.getAttribute('data-project-flow-resize')
    : 'move';
  const location = findProjectFlowTaskLocation(taskId);
  if (!projectId || !taskId || !location?.task) return;
  const startDate = parseProjectFlowDate(location.task.startDate);
  const endDate = parseProjectFlowDate(location.task.endDate);
  if (!startDate || !endDate) return;
  evt.preventDefault();
  bar.setPointerCapture?.(evt.pointerId);
  bar.classList.add('is-dragging');
  projectFlowState.drag = {
    pointerId: evt.pointerId,
    projectId,
    taskId,
    mode: mode === 'start' || mode === 'end' ? mode : 'move',
    startX: evt.clientX,
    deltaDays: 0,
    moved: false,
    startDate: formatProjectFlowDate(startDate),
    endDate: formatProjectFlowDate(endDate),
    durationDays: getProjectFlowDayDiff(startDate, endDate) + 1,
    bar
  };
}

function updateProjectFlowTaskDrag(evt){
  const drag = projectFlowState.drag;
  if (!drag || evt.pointerId !== drag.pointerId) return;
  const dayWidth = getProjectFlowDayWidth();
  const movedPixels = evt.clientX - drag.startX;
  drag.moved = drag.moved || Math.abs(movedPixels) >= 5;
  let deltaDays = Math.round((evt.clientX - drag.startX) / dayWidth);
  if (drag.mode === 'start'){
    deltaDays = Math.min(Math.max(deltaDays, -(drag.durationDays - 1)), drag.durationDays - 1);
  } else if (drag.mode === 'end'){
    deltaDays = Math.max(deltaDays, -(drag.durationDays - 1));
  }
  drag.deltaDays = deltaDays;
  drag.bar?.style?.setProperty('--project-flow-drag-days', String(deltaDays));
  drag.bar?.style?.setProperty('--project-flow-resize-days', String(deltaDays));
  drag.bar?.classList?.toggle('is-resizing-start', drag.mode === 'start');
  drag.bar?.classList?.toggle('is-resizing-end', drag.mode === 'end');
  evt.preventDefault();
}

function finishProjectFlowTaskDrag(evt){
  const drag = projectFlowState.drag;
  if (!drag || evt.pointerId !== drag.pointerId) return;
  drag.bar?.releasePointerCapture?.(evt.pointerId);
  drag.bar?.classList?.remove('is-dragging', 'is-resizing-start', 'is-resizing-end');
  drag.bar?.style?.removeProperty('--project-flow-drag-days');
  drag.bar?.style?.removeProperty('--project-flow-resize-days');
  projectFlowState.drag = null;
  if (drag.moved) projectFlowState.suppressClickUntil = Date.now() + 800;
  if (!drag.deltaDays) return;
  const changed = drag.mode === 'move'
    ? moveProjectFlowTaskByDays(drag.projectId, drag.taskId, drag.deltaDays)
    : resizeProjectFlowTaskByDays(drag.projectId, drag.taskId, drag.mode, drag.deltaDays);
  if (changed){
    renderProjectFlowView();
    setProjectFlowStatus('Oppgave oppdatert.', 'ok');
  }
}

function handleProjectFlowMilestoneSave(evt){
  evt?.preventDefault?.();
  const projectIdFromForm = String($('projectFlowTaskProjectSelect')?.value || '').trim();
  const project = projectState.projects.find(item=>item.id === projectIdFromForm) || getProjectFlowSelectedProject() || projectState.projects[0] || null;
  if (!project?.id){
    setProjectFlowStatus('Velg et prosjekt først.', 'error');
    return;
  }
  const startDateInputValue = String($('projectFlowDateInput')?.value || '').trim();
  const durationValue = Math.max(1, Number.parseInt($('projectFlowDurationValueInput')?.value || '1', 10) || 1);
  const durationUnit = String($('projectFlowDurationUnitSelect')?.value || 'days').trim();
  const phaseId = String($('projectFlowPhaseSelect')?.value || '').trim();
  const title = getProjectFlowPhaseLabel(phaseId);
  const drivenByTaskId = String($('projectFlowDependencyRelationSelect')?.value || '').trim();
  const drivesTaskId = String($('projectFlowDependencyTaskSelect')?.value || '').trim();
  const fileInput = $('projectFlowFileInput');
  const fileName = fileInput?.files?.[0]?.name ? String(fileInput.files[0].name).trim() : '';
  const id = String($('projectFlowMilestoneId')?.value || '').trim() || createProjectFlowId();
  const parsedStart = parseProjectFlowDate(startDateInputValue);
  const parsedEnd = parsedStart ? addProjectFlowDuration(parsedStart, durationValue, durationUnit) : null;
  const endDate = parsedEnd ? formatProjectFlowDate(parsedEnd) : '';
  const startDate = parsedStart ? formatProjectFlowDate(parsedStart) : '';
  if (!parsedStart || !parsedEnd || !PROJECT_FLOW_PHASES.some(phase=>phase.id === phaseId)){
    setProjectFlowStatus('Fyll inn punkt, startdato og varighet.', 'error');
    return;
  }
  if (drivenByTaskId && drivesTaskId && drivenByTaskId === drivesTaskId){
    setProjectFlowStatus('Samme oppgave kan ikke både styre og styres av denne oppgaven.', 'error');
    return;
  }
  if (drivenByTaskId && wouldCreateProjectFlowMutualDependency(id, 'drivenBy', drivenByTaskId)){
    setProjectFlowStatus('Oppgaver kan ikke styre og styres av hverandre samtidig.', 'error');
    return;
  }
  if (drivesTaskId && wouldCreateProjectFlowMutualDependency(id, 'drives', drivesTaskId)){
    setProjectFlowStatus('Oppgaver kan ikke styre og styres av hverandre samtidig.', 'error');
    return;
  }
  const milestones = getProjectFlowMilestones(project.id);
  const existing = milestones.find(item=>item.id === id);
  const now = new Date().toISOString();
  const next = {
    id,
    phaseId,
    title,
    startDate,
    endDate,
    date: startDate,
    durationValue,
    durationUnit,
    fileName: fileName || existing?.fileName || '',
    drivenByTaskId,
    drivesTaskId,
    dependencyRelation: drivenByTaskId ? 'drivenBy' : (drivesTaskId ? 'drives' : ''),
    dependencyTaskId: drivenByTaskId || drivesTaskId,
    createdAt: existing?.createdAt || now,
    updatedAt: now,
    completed: Boolean(existing?.completed)
  };
  applyProjectFlowDrivenByDependency(next);
  const originalProjectId = String(projectFlowState.editingProjectId || '').trim();
  if (originalProjectId && originalProjectId !== project.id && id){
    const originalTasks = getProjectFlowMilestones(originalProjectId);
    const originalTask = originalTasks.find(item=>item.id === id);
    if (originalTask){
      setProjectFlowMilestones(originalProjectId, originalTasks.filter(item=>item.id !== id));
      clearProjectFlowReciprocalDependency(id, 'drivenBy', getProjectFlowRelationTaskId(originalTask, 'drivenBy'));
      clearProjectFlowReciprocalDependency(id, 'drives', getProjectFlowRelationTaskId(originalTask, 'drives'));
    }
  }
  const updated = existing
    ? milestones.map(item=>item.id === id ? next : item)
    : [...milestones, next];
  setProjectFlowMilestones(project.id, updated);
  if (existing && getProjectFlowRelationTaskId(existing, 'drivenBy') !== next.drivenByTaskId){
    clearProjectFlowReciprocalDependency(id, 'drivenBy', getProjectFlowRelationTaskId(existing, 'drivenBy'));
  }
  if (existing && getProjectFlowRelationTaskId(existing, 'drives') !== next.drivesTaskId){
    clearProjectFlowReciprocalDependency(id, 'drives', getProjectFlowRelationTaskId(existing, 'drives'));
  }
  if (next.drivenByTaskId){
    setProjectFlowReciprocalDependency(id, 'drivenBy', next.drivenByTaskId);
  }
  if (next.drivesTaskId){
    setProjectFlowReciprocalDependency(id, 'drives', next.drivesTaskId);
  }
  if (existing){
    const deltas = getProjectFlowDateChangeDeltas(existing, next.startDate, next.endDate);
    propagateProjectFlowDrivenTaskOffsets(project.id, id, deltas.moveDelta, deltas.durationDelta);
  }
  closeProjectFlowMilestoneForm();
  renderProjectFlowView();
  setProjectFlowStatus('Oppgave lagret.', 'ok');
}

function toggleProjectFlowMilestone(projectId, milestoneId, completed){
  const milestones = getProjectFlowMilestones(projectId);
  const updated = milestones.map(item=>item.id === milestoneId ? { ...item, completed: Boolean(completed), updatedAt: new Date().toISOString() } : item);
  setProjectFlowMilestones(projectId, updated);
  renderProjectFlowView();
}

function deleteProjectFlowMilestone(projectId, milestoneId){
  const milestones = getProjectFlowMilestones(projectId);
  const milestone = milestones.find(item=>item.id === milestoneId);
  if (!milestone) return;
  if (!window.confirm(`Slette oppgaven "${getProjectFlowPhaseLabel(milestone.phaseId)}"?`)) return;
  clearProjectFlowReciprocalDependency(milestoneId, 'drivenBy', getProjectFlowRelationTaskId(milestone, 'drivenBy'));
  clearProjectFlowReciprocalDependency(milestoneId, 'drives', getProjectFlowRelationTaskId(milestone, 'drives'));
  setProjectFlowMilestones(projectId, milestones.filter(item=>item.id !== milestoneId));
  renderProjectFlowView();
  setProjectFlowStatus('Oppgave slettet.', 'ok');
}

function deleteProjectFlowMilestoneFromForm(){
  const projectId = String(projectFlowState.editingProjectId || '').trim();
  const milestoneId = String(projectFlowState.editingMilestoneId || '').trim();
  if (!projectId || !milestoneId) return;
  closeProjectFlowMilestoneForm();
  deleteProjectFlowMilestone(projectId, milestoneId);
}

function renderProjectFlowView(){
  const root = $('projectFlowTimeline');
  if (!root) return;
  loadProjectFlowState();
  updateProjectFlowProjectSelect();
  populateProjectFlowPhaseSelect();
  root.innerHTML = '';
  const addBtn = $('addProjectFlowMilestoneBtn');
  const refreshBtn = $('refreshProjectFlowBtn');
  const visibleProjects = getProjectFlowVisibleProjects();
  if (addBtn) addBtn.disabled = !projectState.projects.length;
  if (refreshBtn) refreshBtn.disabled = false;
  if (!visibleProjects.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = 'Ingen prosjekter er tilgjengelige for prosjektflyt.';
    root.appendChild(empty);
    return;
  }

  const milestones = visibleProjects.flatMap(project=>{
    return getProjectFlowMilestones(project.id).map(task=>({
      ...task,
      projectId: project.id,
      projectName: project.name || 'Uten navn',
      projectNumber: project.projectNumber || ''
    }));
  });
  const { start, end } = getProjectFlowRange(milestones);
  const dates = buildProjectFlowDates(start, end);
  const todayKey = getProjectFlowDateKey(new Date());
  const sortedMilestones = [...milestones].sort((a, b)=>{
    const phaseA = PROJECT_FLOW_PHASES.findIndex(phase=>phase.id === a.phaseId);
    const phaseB = PROJECT_FLOW_PHASES.findIndex(phase=>phase.id === b.phaseId);
    if (phaseA !== phaseB) return phaseA - phaseB;
    return String(a.startDate || '').localeCompare(String(b.startDate || ''), 'no');
  });

  const title = document.createElement('div');
  title.className = 'project-flow-title';
  const heading = document.createElement('h3');
  const selectedProject = getProjectFlowSelectedProject();
  projectFlowState.taskColumnWidth = measureProjectFlowTaskColumnWidth([
    'Oppgave',
    'Ny oppgave',
    'Ingen oppgaver',
    ...PROJECT_FLOW_PHASES.map(phase=>phase.label),
    ...visibleProjects.map(project=>getProjectFlowProjectLabel(project)),
    ...sortedMilestones.map(item=>getProjectFlowTaskLabel(item, !selectedProject))
  ]);
  heading.textContent = selectedProject
    ? (selectedProject.projectNumber
      ? `${selectedProject.projectNumber} - ${selectedProject.name || 'Uten navn'}`
      : selectedProject.name || 'Uten navn')
    : 'Alle prosjekter';
  const meta = document.createElement('p');
  const completedCount = sortedMilestones.filter(item=>item.completed).length;
  meta.className = 'muted-text';
  meta.textContent = `${completedCount} av ${sortedMilestones.length} oppgaver fullført`;
  title.append(heading, meta);

  const scroller = document.createElement('div');
  scroller.className = 'project-flow-scroller';
  const topScrollbar = document.createElement('div');
  topScrollbar.className = 'project-flow-top-scrollbar';
  const topScrollbarInner = document.createElement('div');
  topScrollbarInner.className = 'project-flow-top-scrollbar-inner';
  topScrollbarInner.style.setProperty('--project-flow-days', String(dates.length));
  topScrollbarInner.style.setProperty('--project-flow-day-width', `${getProjectFlowDayWidth()}px`);
  topScrollbarInner.style.setProperty('--project-flow-task-width', `${getProjectFlowTaskColumnWidth()}px`);
  topScrollbar.appendChild(topScrollbarInner);
  const grid = document.createElement('div');
  grid.className = 'project-flow-grid';
  grid.style.setProperty('--project-flow-days', String(dates.length));
  grid.style.setProperty('--project-flow-day-width', `${getProjectFlowDayWidth()}px`);
  grid.style.setProperty('--project-flow-task-width', `${getProjectFlowTaskColumnWidth()}px`);

  const phaseHead = document.createElement('div');
  phaseHead.className = 'project-flow-phase-head';
  phaseHead.textContent = 'Oppgave';
  grid.appendChild(phaseHead);

  getProjectFlowWeekSpans(dates).forEach(span=>{
    const week = document.createElement('div');
    week.className = 'project-flow-week-head';
    week.style.gridColumn = `span ${span.days}`;
    week.textContent = `UKE ${span.weekNumber}`;
    grid.appendChild(week);
  });

  dates.forEach((date, index)=>{
    const day = document.createElement('div');
    day.className = 'project-flow-day-head';
    const dateKey = getProjectFlowDateKey(date);
    if (dateKey === todayKey) day.classList.add('is-today');
    if (isProjectFlowWeekStart(date, index)) day.classList.add('is-week-start');
    const dateLine = document.createElement('span');
    dateLine.className = 'project-flow-date-line';
    const month = document.createElement('span');
    month.textContent = date.toLocaleDateString('no-NO', { month: 'short' }).replace('.', '');
    const dayNumber = document.createElement('strong');
    dayNumber.textContent = String(date.getDate());
    dateLine.append(month, dayNumber);
    const weekday = document.createElement('span');
    weekday.className = 'project-flow-weekday';
    weekday.textContent = date.toLocaleDateString('no-NO', { weekday: 'short' }).replace('.', '');
    day.append(dateLine, weekday);
    grid.appendChild(day);
  });

  const renderGroups = selectedProject
    ? [{ id: selectedProject.id, label: getProjectFlowProjectLabel(selectedProject), tasks: sortedMilestones.filter(task=>task.projectId === selectedProject.id), showHeader: false, includeProjectInTask: false }]
    : [{
      id: PROJECT_FLOW_ALL_PROJECTS,
      label: 'Alle prosjekter',
      tasks: sortedMilestones,
      showHeader: false,
      includeProjectInTask: false,
      projectRows: visibleProjects.filter(project=>getProjectFlowMilestones(project.id).length)
    }];

  renderGroups.forEach(group=>{
    if (group.showHeader){
      const projectLabel = document.createElement('div');
      projectLabel.className = 'project-flow-project-label';
      projectLabel.textContent = group.label;
      grid.appendChild(projectLabel);
      dates.forEach((date, index)=>{
        const cell = document.createElement('div');
        cell.className = 'project-flow-project-cell';
        if (getProjectFlowDateKey(date) === todayKey) cell.classList.add('is-today');
        if (isProjectFlowWeekStart(date, index)) cell.classList.add('is-week-start');
        grid.appendChild(cell);
      });
    }

    const tasksByPhase = getProjectFlowTasksByPhase(group.tasks);
    PROJECT_FLOW_PHASES.forEach(phase=>{
      const phaseTasks = tasksByPhase.get(phase.id) || [];
      const collapsed = projectFlowState.collapsedPhaseIds.has(phase.id);
      const phaseLabel = document.createElement('button');
      phaseLabel.type = 'button';
      phaseLabel.className = 'project-flow-group-label';
      phaseLabel.dataset.projectFlowPhaseToggle = phase.id;
      phaseLabel.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
      phaseLabel.innerHTML = `<span class="project-flow-group-caret">${collapsed ? '+' : '-'}</span><span>${phase.label}</span><strong>${phaseTasks.length}</strong>`;
      grid.appendChild(phaseLabel);
      dates.forEach((date, index)=>{
        const dateKey = getProjectFlowDateKey(date);
        const cell = document.createElement('div');
        cell.className = 'project-flow-group-cell';
        if (dateKey === todayKey) cell.classList.add('is-today');
        if (isProjectFlowWeekStart(date, index)) cell.classList.add('is-week-start');
        grid.appendChild(cell);
      });
      if (collapsed) return;

      if (!phaseTasks.length && !group.projectRows?.length){
        const emptyLabel = document.createElement('div');
        emptyLabel.className = 'project-flow-task-label is-empty';
        emptyLabel.textContent = 'Ingen oppgaver';
        grid.appendChild(emptyLabel);
        dates.forEach((date, index)=>{
          const dateKey = getProjectFlowDateKey(date);
          const cell = document.createElement('div');
          cell.className = 'project-flow-cell is-empty';
          cell.dataset.projectFlowDateCell = '';
          cell.dataset.phaseId = phase.id;
          cell.dataset.projectId = selectedProject?.id || '';
          cell.dataset.date = dateKey;
          if (dateKey === todayKey) cell.classList.add('is-today');
          if (isProjectFlowWeekStart(date, index)) cell.classList.add('is-week-start');
          grid.appendChild(cell);
        });
        return;
      }

      const appendEmptyTaskRow = (labelText = 'Ny oppgave', projectId = selectedProject?.id || '', isProjectRow = false)=>{
        const emptyLabel = document.createElement('div');
        emptyLabel.className = 'project-flow-task-label is-empty';
        if (isProjectRow) emptyLabel.classList.add('is-project-row');
        emptyLabel.textContent = labelText;
        grid.appendChild(emptyLabel);
        dates.forEach((date, index)=>{
          const dateKey = getProjectFlowDateKey(date);
          const cell = document.createElement('div');
          cell.className = 'project-flow-cell is-empty';
          cell.dataset.projectFlowDateCell = '';
          cell.dataset.phaseId = phase.id;
          cell.dataset.projectId = projectId;
          cell.dataset.date = dateKey;
          if (dateKey === todayKey) cell.classList.add('is-today');
          if (isProjectFlowWeekStart(date, index)) cell.classList.add('is-week-start');
          grid.appendChild(cell);
        });
      };

      const rowsToRender = group.projectRows?.length
        ? group.projectRows.flatMap(project=>{
          const projectTasks = phaseTasks.filter(task=>task.projectId === project.id);
          return projectTasks.length
            ? projectTasks.map(task=>({ type: 'task', task, project }))
            : [{ type: 'empty', project }];
        })
        : phaseTasks.map(task=>({ type: 'task', task, project: null }));

      rowsToRender.forEach(row=>{
        if (row.type === 'empty'){
          appendEmptyTaskRow(getProjectFlowProjectLabel(row.project), row.project.id, true);
          return;
        }
        const item = row.task;
        const projectRowLabel = group.projectRows?.length
          ? getProjectFlowProjectLabel(row.project || projectState.projects.find(project=>project.id === item.projectId))
          : getProjectFlowTaskLabel(item, group.includeProjectInTask);
        const taskLabel = document.createElement('div');
        taskLabel.className = 'project-flow-task-label';
        taskLabel.dataset.projectFlowTaskRow = item.id;
        taskLabel.dataset.projectId = item.projectId;
        if (item.completed) taskLabel.classList.add('is-completed');
        const taskTitle = document.createElement('button');
        taskTitle.type = 'button';
        taskTitle.dataset.projectFlowEdit = item.id;
        taskTitle.dataset.projectId = item.projectId;
        taskTitle.textContent = projectRowLabel;
        const taskMeta = document.createElement('span');
        taskMeta.textContent = `${formatProjectFlowDisplayDate(item.startDate)} - ${formatProjectFlowDisplayDate(item.endDate)}`;
        taskLabel.append(taskTitle, taskMeta);
        if (item.fileName){
          const fileMeta = document.createElement('span');
          fileMeta.className = 'project-flow-task-file';
          fileMeta.textContent = `Fil: ${item.fileName}`;
          taskLabel.appendChild(fileMeta);
        }
        grid.appendChild(taskLabel);

        const startDate = parseProjectFlowDate(item.startDate);
        const endDate = parseProjectFlowDate(item.endDate);
        const startIndex = startDate ? getProjectFlowDayDiff(start, startDate) : -1;
        const endIndex = endDate ? getProjectFlowDayDiff(start, endDate) : startIndex;
        const spanDays = Math.max(1, endIndex - startIndex + 1);
        dates.forEach((date, index)=>{
          const dateKey = getProjectFlowDateKey(date);
          const cell = document.createElement('div');
          cell.className = 'project-flow-cell';
          cell.dataset.projectId = item.projectId;
          cell.dataset.phaseId = item.phaseId;
          cell.dataset.projectFlowRowTask = item.id;
          cell.dataset.date = dateKey;
          if (dateKey === todayKey) cell.classList.add('is-today');
          if (isProjectFlowWeekStart(date, index)) cell.classList.add('is-week-start');
          const dateIndex = getProjectFlowDayDiff(start, date);
          const isCovered = dateIndex >= startIndex && dateIndex <= endIndex;
          cell.dataset.projectFlowDateCell = isCovered ? item.id : '';
          if (isCovered) cell.classList.add('is-covered');
          if (isCovered && dateIndex !== startIndex) cell.classList.add('is-covered-continuation');
          if (dateIndex === startIndex){
            const bar = document.createElement('div');
            bar.className = 'project-flow-bar';
            bar.dataset.projectFlowDrag = item.id;
            bar.dataset.projectFlowTaskBar = item.id;
            bar.dataset.projectId = item.projectId;
            if (spanDays <= 1) bar.classList.add('is-single-day');
            if (item.completed) bar.classList.add('is-completed');
            bar.style.setProperty('--project-flow-span-days', String(spanDays));
            const resizeStart = document.createElement('span');
            resizeStart.className = 'project-flow-resize-handle is-start';
            resizeStart.dataset.projectFlowResize = 'start';
            resizeStart.setAttribute('aria-hidden', 'true');
            const check = document.createElement('input');
            check.type = 'checkbox';
            check.checked = item.completed;
            check.dataset.projectFlowToggle = item.id;
            check.dataset.projectId = item.projectId;
            check.title = 'Marker fullført';
            const text = document.createElement('button');
            text.type = 'button';
            text.className = 'project-flow-bar-title';
            text.dataset.projectFlowEdit = item.id;
            text.dataset.projectId = item.projectId;
            text.textContent = projectRowLabel;
            const remove = document.createElement('button');
            remove.type = 'button';
            remove.className = 'project-flow-delete';
            remove.dataset.projectFlowDelete = item.id;
            remove.dataset.projectId = item.projectId;
            remove.textContent = '×';
            remove.title = 'Slett oppgave';
            const resizeEnd = document.createElement('span');
            resizeEnd.className = 'project-flow-resize-handle is-end';
            resizeEnd.dataset.projectFlowResize = 'end';
            resizeEnd.setAttribute('aria-hidden', 'true');
            const linkIn = document.createElement('button');
            linkIn.type = 'button';
            linkIn.className = 'project-flow-link-handle is-driven-by';
            linkIn.dataset.projectFlowLink = 'drivenBy';
            linkIn.dataset.projectFlowLinkTask = item.id;
            linkIn.dataset.projectId = item.projectId;
            linkIn.title = 'Styres av';
            linkIn.setAttribute('aria-label', 'Styres av');
            const linkOut = document.createElement('button');
            linkOut.type = 'button';
            linkOut.className = 'project-flow-link-handle is-drives';
            linkOut.dataset.projectFlowLink = 'drives';
            linkOut.dataset.projectFlowLinkTask = item.id;
            linkOut.dataset.projectId = item.projectId;
            linkOut.title = 'Styrer';
            linkOut.setAttribute('aria-label', 'Styrer');
            if (spanDays <= 1){
              bar.append(linkIn, check, linkOut);
            } else {
              bar.append(linkIn, resizeStart, check, text, remove, resizeEnd, linkOut);
            }
            cell.appendChild(bar);
          }
          grid.appendChild(cell);
        });
      });
      appendEmptyTaskRow();
    });
  });

  scroller.appendChild(grid);
  const linkOverlay = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  linkOverlay.classList.add('project-flow-link-overlay');
  scroller.appendChild(linkOverlay);
  syncProjectFlowTopScrollbar(scroller, topScrollbar);
  root.append(title, topScrollbar, scroller);
  scrollProjectFlowToCurrentWeek(scroller, start, topScrollbar);
  requestAnimationFrame(renderProjectFlowDependencyLines);
  if (!milestones.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty project-flow-empty';
    empty.textContent = 'Ingen Gantt-oppgaver er opprettet ennå.';
    root.appendChild(empty);
  }
}

function renderProjectDashboard(){
  const listEl = $('projectList');
  if (!listEl) return;
  loadProjectFlowState();
  updateProjectArchiveUi();
  listEl.innerHTML = '';
  if (!projectState.projects.length){
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    const text = document.createElement('p');
    text.textContent = 'Ingen prosjekter er registrert ennå.';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn alt';
    btn.dataset.action = 'create-project';
    btn.textContent = 'Nytt prosjekt';
    btn.disabled = !authState.loggedIn || projectState.showArchive === true;
    empty.appendChild(text);
    empty.appendChild(btn);
    listEl.appendChild(empty);
    projectState.expandedProjectId = null;
    renderOffersList();
    renderProjectFlowView();
    return;
  }
  const archiveMode = projectState.showArchive === true;
  const projectsForView = projectState.projects.filter(project=>archiveMode ? projectIsArchived(project) : !projectIsArchived(project));
  const visibleProjects = projectsForView.filter(project=>projectMatchesSearch(project));
  ensureProjectFolderStatusesLoaded();
  if (!visibleProjects.length){
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    const text = document.createElement('p');
    if (!projectsForView.length){
      text.textContent = archiveMode ? 'Ingen prosjekter i arkivet.' : 'Ingen aktive prosjekter.';
    } else {
      text.textContent = 'Ingen prosjekter matcher søket.';
    }
    empty.appendChild(text);
    listEl.appendChild(empty);
    renderOffersList();
    renderProjectFlowView();
    return;
  }
  const frag = document.createDocumentFragment();
  visibleProjects.forEach(project=>{
    const expanded = projectState.expandedProjectId === project.id;
    const projectLines = Array.isArray(project.lines) ? project.lines : [];
    const projectTotal = round2(projectLines.reduce((sum, line)=>{
      const lineTotal = resolveLineDisplayTotal(line);
      return Number.isFinite(lineTotal) ? sum + lineTotal : sum;
    }, 0));
    const projectBusbarTotal = round2(projectLines.reduce((sum, line)=>{
      const busbarTotal = Number(line?.totals?.totalExMontasje);
      return Number.isFinite(busbarTotal) ? sum + busbarTotal : sum;
    }, 0));
    const projectSkinMaterialTotal = round2(projectLines.reduce((sum, line)=>{
      const material = resolveLineSkinMaterialCost(line);
      return Number.isFinite(material) ? sum + material : sum;
    }, 0));
    const projectMontasjeTotal = round2(projectLines.reduce((sum, line)=>{
      const montasjeTotal = Number(line?.totals?.totalInclMontasje);
      return Number.isFinite(montasjeTotal) ? sum + montasjeTotal : sum;
    }, 0));
    const row = document.createElement('section');
    row.className = 'project-row';
    row.dataset.projectRowId = project.id || '';
    if (expanded) row.classList.add('is-expanded');

    const head = document.createElement('div');
    head.className = 'project-row-head';
    const titleWrap = document.createElement('div');
    titleWrap.className = 'project-row-title';
    const title = document.createElement('h3');
    const projectTitle = project.name || 'Uten navn';
    title.textContent = project.projectNumber
      ? `${project.projectNumber} - ${projectTitle}`
      : projectTitle;
    const flowStatus = getProjectFlowStatusForProject(project);
    const titleRow = document.createElement('div');
    titleRow.className = 'project-title-row';
    const statusBadge = document.createElement('span');
    statusBadge.className = `project-flow-status-badge is-${flowStatus.tone}`;
    statusBadge.textContent = flowStatus.ageText
      ? `${flowStatus.label} - ${flowStatus.ageText}`
      : flowStatus.label;
    const projectStatus = getProjectStatusConfig(project);
    const projectStatusBtn = document.createElement('button');
    projectStatusBtn.type = 'button';
    projectStatusBtn.className = `project-status-badge is-${projectStatus.tone}`;
    projectStatusBtn.dataset.projectStatusEdit = project.id;
    projectStatusBtn.textContent = projectStatus.label;
    projectStatusBtn.title = archiveMode
      ? 'Endre prosjektstatus for å flytte prosjektet tilbake til prosjektoversikten.'
      : 'Endre prosjektstatus';
    titleRow.append(title, statusBadge, projectStatusBtn);
    const projectAddonConfig = getProjectSelectedAddonConfig(project);
    const createFieldText = (label, value)=>{
      const wrapper = document.createElement('span');
      wrapper.className = 'project-field-text';
      const labelEl = document.createElement('strong');
      labelEl.textContent = `${label}:`;
      const valueEl = document.createElement('span');
      valueEl.textContent = ` ${value || '-'}`;
      wrapper.append(labelEl, valueEl);
      return wrapper;
    };
    const createInfoRow = (label, value)=>{
      const infoRow = document.createElement('div');
      infoRow.className = 'project-info-row project-row-meta';
      infoRow.appendChild(createFieldText(label, value));
      return infoRow;
    };
    const responsible = createInfoRow('Prosjektansvarlig', getProjectResponsibleName(project) || '-');
    const customer = createInfoRow('Kunde', project.customer || '-');
    const contact = createInfoRow('Kontaktperson', project.contactPerson || '-');
    const created = createInfoRow('Opprettet', formatProjectTimestamp(project.createdAt));
    const lineCount = projectLines.length;
    const summary = createInfoRow('Linjer', String(lineCount));
    const marginBadge = document.createElement('p');
    marginBadge.className = 'project-margin-badge';
    marginBadge.textContent = formatProjectMarginBadgeText(project);
    const marginStats = getProjectMaterialMarginStats(project);
    if (marginStats.uniqueCount > 1){
      marginBadge.classList.add('is-mixed');
    }
    if (shouldUseWarningForProjectMargin(marginStats)){
      marginBadge.classList.add('is-warning');
    }

    const setMarginBtn = document.createElement('button');
    setMarginBtn.type = 'button';
    setMarginBtn.className = 'btn alt project-margin-btn';
    setMarginBtn.dataset.projectSetMargin = project.id;
    setMarginBtn.textContent = 'Endre';
    setMarginBtn.disabled = archiveMode || !projectLines.length;

    const marginRow = document.createElement('div');
    marginRow.className = 'project-margin-row';
    marginRow.appendChild(marginBadge);
    marginRow.appendChild(setMarginBtn);

    titleWrap.appendChild(titleRow);
    titleWrap.appendChild(marginRow);
    const infoStack = document.createElement('div');
    infoStack.className = 'project-info-column';
    infoStack.appendChild(responsible);
    infoStack.appendChild(customer);
    infoStack.appendChild(contact);
    infoStack.appendChild(created);
    infoStack.appendChild(summary);

    const actions = document.createElement('div');
    actions.className = 'project-row-actions';
    const hasProjectFolder = projectHasConfirmedFolder(project);
    const disableProjectActions = archiveMode;

    const detailBtn = document.createElement('button');
    detailBtn.type = 'button';
    detailBtn.className = 'btn alt project-detail-toggle-btn';
    detailBtn.dataset.projectDetail = project.id;
    detailBtn.textContent = expanded ? 'Skjul linjer' : 'Vis linjer';
    detailBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    const safeProjectId = String(project.id || '').replace(/[^A-Za-z0-9_-]/g,'');
    const detailId = `project-detail-${safeProjectId || Math.random().toString(36).slice(2)}`;
    detailBtn.setAttribute('aria-controls', detailId);

    const newLineBtn = document.createElement('button');
    newLineBtn.type = 'button';
    newLineBtn.className = 'btn';
    newLineBtn.dataset.projectNewline = project.id;
    newLineBtn.textContent = 'Ny linje';
    newLineBtn.disabled = disableProjectActions;

    const generateOfferBtn = document.createElement('button');
    generateOfferBtn.type = 'button';
    generateOfferBtn.className = 'btn';
    generateOfferBtn.dataset.projectGenerateOffer = project.id;
    generateOfferBtn.textContent = 'Generer tilbud';
    generateOfferBtn.disabled = disableProjectActions || !projectLines.length || !hasProjectFolder;
    if (!hasProjectFolder) generateOfferBtn.title = 'Prosjektmappe må opprettes før tilbud kan genereres.';

    const openFolderBtn = document.createElement('button');
    openFolderBtn.type = 'button';
    openFolderBtn.className = 'btn alt';
    openFolderBtn.dataset.projectOpenFolder = project.id;
    openFolderBtn.textContent = 'Åpne prosjektmappe';
    openFolderBtn.disabled = disableProjectActions || !authState.loggedIn || !hasProjectFolder;
    if (!hasProjectFolder) openFolderBtn.title = 'Fant ingen prosjektmappe på SharePoint-stien.';

    const createFolderBtn = document.createElement('button');
    createFolderBtn.type = 'button';
    createFolderBtn.className = 'btn alt';
    createFolderBtn.dataset.projectCreateFolder = project.id;
    createFolderBtn.textContent = 'Opprett prosjektmappe';
    createFolderBtn.disabled = disableProjectActions || !authState.loggedIn;

    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.className = 'btn alt';
    copyBtn.dataset.projectCopy = project.id;
    copyBtn.textContent = 'Kopier';
    copyBtn.disabled = disableProjectActions;

    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn alt';
    editBtn.dataset.projectCardEdit = project.id;
    editBtn.textContent = 'Endre';
    editBtn.disabled = disableProjectActions;

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'btn danger';
    deleteBtn.dataset.projectDelete = project.id;
    deleteBtn.textContent = 'Slett';
    deleteBtn.disabled = disableProjectActions;

    const actionButtons = document.createElement('div');
    actionButtons.className = 'project-action-buttons';
    actionButtons.appendChild(detailBtn);
    actionButtons.appendChild(newLineBtn);
    actionButtons.appendChild(generateOfferBtn);
    actionButtons.appendChild(openFolderBtn);
    actionButtons.appendChild(createFolderBtn);
    actionButtons.appendChild(copyBtn);
    actionButtons.appendChild(editBtn);
    actionButtons.appendChild(deleteBtn);

    const metricsStack = document.createElement('div');
    metricsStack.className = 'project-metrics-column';
    const createMetricText = (item)=>{
      if (!item.plain) return createFieldText(item.label, item.value);
      const wrapper = document.createElement('span');
      wrapper.className = 'project-field-text';
      const labelEl = document.createElement('span');
      labelEl.textContent = `${item.label}:`;
      const valueEl = document.createElement('span');
      valueEl.textContent = ` ${item.value || '-'}`;
      wrapper.append(labelEl, valueEl);
      return wrapper;
    };
    [
      { label: 'Total strømskinne', value: `${fmtNO.format(projectBusbarTotal)} NOK` },
      { label: 'Materiellkost', value: `${fmtNO.format(projectSkinMaterialTotal)} NOK`, className: 'is-muted-italic', plain: true },
      { label: 'Total montasje', value: `${fmtNO.format(projectMontasjeTotal)} NOK` },
      { label: 'Totalsum inkludert i tilbud', value: `${fmtNO.format(projectTotal)} NOK` }
    ].forEach(item=>{
      const metricRow = document.createElement('div');
      metricRow.className = `project-metric-row${item.className ? ` ${item.className}` : ''}`;
      metricRow.appendChild(createMetricText(item));
      metricsStack.appendChild(metricRow);
    });

    const addonStack = document.createElement('div');
    addonStack.className = 'project-action-addon-stack';
    ['include', 'show', 'unit'].forEach(group=>{
      const controlRow = document.createElement('div');
      controlRow.className = 'project-action-addon-row';
      controlRow.appendChild(buildAddonSelectorControl(projectAddonConfig, {
        className: 'project-inline-selectors',
        scope: 'project',
        projectId: project.id,
        groups: [group]
      }));
      if (archiveMode){
        controlRow.querySelectorAll('input, button, select, textarea').forEach(control=>{
          control.disabled = true;
        });
      }
      addonStack.appendChild(controlRow);
    });

    actions.appendChild(actionButtons);

    head.appendChild(titleWrap);
    head.appendChild(actions);
    row.appendChild(head);

    const bodyGrid = document.createElement('div');
    bodyGrid.className = 'project-body-grid';
    bodyGrid.appendChild(infoStack);
    bodyGrid.appendChild(metricsStack);
    bodyGrid.appendChild(addonStack);
    row.appendChild(bodyGrid);

    const detail = document.createElement('div');
    detail.className = 'project-detail';
    detail.id = detailId;
    detail.hidden = !expanded;

    const linesWrapper = document.createElement('div');
    linesWrapper.className = 'project-detail-lines';
    const lines = [...projectLines];
    lines.sort((a,b)=>compareLinesForSort(a, b, projectState.lineSort));
    if (!lines.length){
      const emptyLine = document.createElement('p');
      emptyLine.className = 'project-line-empty';
      emptyLine.textContent = 'Ingen lagrede linjer. Klikk «Ny linje» for å starte.';
      linesWrapper.appendChild(emptyLine);
    } else {
      lines.forEach(line=>{
        const lineWrap = document.createElement('div');
        lineWrap.className = 'project-line-item';
        const lineMain = document.createElement('div');
        lineMain.className = 'project-line-main';

        const lineBtn = document.createElement('button');
        lineBtn.type = 'button';
        lineBtn.className = 'project-line-row';
        lineBtn.dataset.lineEdit = line.id;
        lineBtn.dataset.projectId = project.id;
        lineBtn.disabled = archiveMode;

        const nameSpan = document.createElement('span');
        nameSpan.className = 'line-name';
        nameSpan.textContent = line.lineNumber || 'Uten linjenummer';

        const infoSpan = document.createElement('span');
        infoSpan.className = 'line-info';
        infoSpan.textContent = formatLineSummary(line);

        const totalSpan = document.createElement('span');
        totalSpan.className = 'line-total';
        totalSpan.textContent = formatLineTotal(line);

        const materialSpan = document.createElement('span');
        materialSpan.className = 'line-material';
        materialSpan.textContent = formatLineSkinMaterialCost(line);

        const updatedSpan = document.createElement('span');
        updatedSpan.className = 'line-updated';
        updatedSpan.textContent = formatLineUpdatedText(line);

        lineBtn.appendChild(nameSpan);
        lineBtn.appendChild(infoSpan);
        lineBtn.appendChild(materialSpan);
        lineBtn.appendChild(totalSpan);
        lineBtn.appendChild(updatedSpan);
        const lineAddonControl = buildAddonSelectorControl(
          getSelectedAddonConfig(line, projectAddonConfig),
          {
            className: 'line-addon-selectors',
            scope: 'line',
            projectId: project.id,
            lineId: line.id
          }
        );
        if (archiveMode){
          lineAddonControl.querySelectorAll('input, button, select, textarea').forEach(control=>{
            control.disabled = true;
          });
        }

        const lineActionButtons = document.createElement('div');
        lineActionButtons.className = 'line-action-buttons';

        const lineEditBtn = document.createElement('button');
        lineEditBtn.type = 'button';
        lineEditBtn.className = 'btn alt line-edit-btn';
        lineEditBtn.dataset.lineEdit = line.id;
        lineEditBtn.dataset.projectId = project.id;
        lineEditBtn.textContent = 'Endre';
        lineEditBtn.disabled = archiveMode;

        const lineDeleteBtn = document.createElement('button');
        lineDeleteBtn.type = 'button';
        lineDeleteBtn.className = 'btn danger line-delete-btn';
        lineDeleteBtn.dataset.lineDelete = line.id;
        lineDeleteBtn.dataset.projectId = project.id;
        lineDeleteBtn.textContent = 'Slett';
        lineDeleteBtn.disabled = archiveMode;

        lineMain.appendChild(lineBtn);
        lineMain.appendChild(lineAddonControl);
        lineWrap.appendChild(lineMain);
        lineActionButtons.appendChild(lineEditBtn);
        lineActionButtons.appendChild(lineDeleteBtn);
        lineWrap.appendChild(lineActionButtons);
        linesWrapper.appendChild(lineWrap);
      });
    }
    detail.appendChild(linesWrapper);
    row.appendChild(detail);

    frag.appendChild(row);
  });
  listEl.appendChild(frag);
  renderOffersList();
  renderProjectFlowView();
}

async function initProjectDashboard(){
  loadProjectFlowState();
  resetProjectFolderStatusState();
  projectState.projects = [];
  if (!getCurrentUserEmail()){
    projectState.projects = [];
  }
  applyDashboardSortModesFromStorage();
  sortProjects();
  updateProjectHistories();
  await syncProjectsForCurrentUser();
  if (hasDashboardUI()){
    showDashboardView({ clearSelection: true });
    applyDashboardQueryContext();
  }
  updateProjectMetaDisplay();
}

function applyDashboardQueryContext(){
  if (!hasDashboardUI()) return;
  const params = new URLSearchParams(window.location.search);
  const projectId = params.get('project');
  if (!projectId) return;
  const project = getProjectById(projectId);
  if (!project) return;
  setDashboardPage('projects');
  const shouldFocusProject = params.get('focusProject') === '1';
  if (shouldFocusProject && projectState.projectSearchTerm && !projectMatchesSearch(project)){
    setProjectSearchTerm('', { render: false });
  }
  setActiveProject(project);
  projectState.expandedProjectId = project.id;
  renderProjectDashboard();
  if (shouldFocusProject){
    setTimeout(()=>scrollProjectIntoView(project.id), 0);
  }

  if (typeof history !== 'undefined' && history.replaceState){
    const cleanUrl = new URL(window.location.href);
    cleanUrl.search = '';
    history.replaceState({}, '', cleanUrl.toString());
  }
}

function applyCalculatorQueryContext(){
  if (!hasCalculatorUI()) return;
  const params = new URLSearchParams(window.location.search);
  const projectId = params.get('project');
  const lineId = params.get('line');
  const newLine = params.get('newLine') === '1';
  if (!projectId) return;
  const project = getProjectById(projectId);
  if (!project) return;

  setActiveProject(project);
  projectState.expandedProjectId = project.id;
  if (lineId){
    openProjectLine(project.id, lineId);
  } else if (newLine){
    resetCalculatorForm({ preserveProject: true });
    applyProjectAddonCheckboxesToCalculator(project);
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Oppgi linjenummer for ny linje.';
    scrollPageToTop();
  }

  if (typeof history !== 'undefined' && history.replaceState){
    const cleanUrl = new URL(window.location.href);
    cleanUrl.search = '';
    history.replaceState({}, '', cleanUrl.toString());
  }
}

function selectProjectDetail(projectId){
  if (!projectId){
    projectState.expandedProjectId = null;
    renderProjectDashboard();
    return;
  }
  if (projectState.expandedProjectId === projectId){
    projectState.expandedProjectId = null;
    renderProjectDashboard();
    return;
  }
  const project = getProjectById(projectId);
  if (!project) return;
  projectState.expandedProjectId = projectId;
  setActiveProject(project);
  renderProjectDashboard();
}

function startNewLineForProject(projectId){
  const project = getProjectById(projectId);
  if (!project) return;
  if (!hasCalculatorUI()){
    goToCalculator({ project: project.id, newLine: '1' });
    return;
  }
  setActiveProject(project);
  if (projectState.expandedProjectId !== projectId){
    projectState.expandedProjectId = projectId;
    renderProjectDashboard();
  } else {
    renderProjectDashboard();
  }
  resetCalculatorForm({ preserveProject: true });
  applyProjectAddonCheckboxesToCalculator(project);
  showCalculatorView();
  scrollPageToTop();
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = 'Oppgi linjenummer for ny linje.';
}

function scrollPageToTop(){
  if (typeof window === 'undefined' || typeof window.scrollTo !== 'function') return;
  window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
}

function scrollProjectIntoView(projectId){
  if (!projectId || typeof document === 'undefined') return;
  const row = Array.from(document.querySelectorAll('[data-project-row-id]'))
    .find(el=>el.getAttribute('data-project-row-id') === String(projectId));
  if (!row || typeof row.scrollIntoView !== 'function') return;
  row.scrollIntoView({ block: 'start', inline: 'nearest', behavior: 'auto' });
}

function updateProjectArchiveUi(){
  const archiveMode = projectState.showArchive === true;
  const toggleBtn = $('projectArchiveToggleBtn');
  if (toggleBtn){
    toggleBtn.textContent = archiveMode ? 'Prosjekter' : 'Prosjekt arkiv';
    toggleBtn.setAttribute('aria-pressed', archiveMode ? 'true' : 'false');
  }
  const titleEl = document.querySelector('#dashboardView .dashboard-intro h2');
  if (titleEl){
    titleEl.textContent = archiveMode ? 'Prosjekt arkiv' : 'Prosjekter';
  }
  const introEl = document.querySelector('#dashboardView .dashboard-intro .muted-text');
  if (introEl){
    introEl.textContent = archiveMode
      ? 'Arkiverte prosjekter er skrivebeskyttet. Du kan fortsatt vise og skjule linjer.'
      : 'Velg et prosjekt, eller opprett nytt prosjekt. Prisverktøy åpnes på egen side.';
  }
  const newProjectBtn = $('newProjectBtn');
  if (newProjectBtn){
    newProjectBtn.disabled = !authState.loggedIn || archiveMode;
  }
}

function toggleProjectArchiveView(){
  projectState.showArchive = !projectState.showArchive;
  projectState.expandedProjectId = null;
  updateProjectArchiveUi();
  renderProjectDashboard();
}

function ensureOption(selectEl, value, label){
  if (!selectEl) return;
  if ([...selectEl.options].some(opt=>opt.value === value)) return;
  const opt = document.createElement('option');
  opt.value = value;
  opt.textContent = label || value;
  selectEl.appendChild(opt);
}

function applyInputsToCalculator(input){
  if (!input) return;
  const seriesEl = $('series');
  if (seriesEl){
    seriesEl.value = input.series || '';
  }
  refreshUIBySeries();
  const distEl = $('dist');
  if (distEl){
    let distValue = 'Nei';
    if (typeof input.dist === 'string'){
      distValue = input.dist;
    } else if (input.dist){
      distValue = 'Ja';
    }
    distEl.value = distValue;
  }
  const setNumberValue = (id, value)=>{
    const el = $(id);
    if (!el) return;
    if (value === undefined || value === null || Number.isNaN(value)){
      el.value = '';
      return;
    }
    el.value = String(value);
  };
  setNumberValue('meter', input.meter);
  setNumberValue('v90h', input.v90_h ?? input.v90h);
  setNumberValue('v90v', input.v90_v ?? input.v90v);
  setNumberValue('fbQty', input.fbQty);
  setNumberValue('boxQty', input.boxQty);
  setNumberValue('boxInnmatSum', input.boxInnmatSum);
  const ampEl = $('ampSelect');
  if (ampEl){
    const ampValue = Number(input.ampere);
    if (Number.isFinite(ampValue)){
      const ampStr = String(ampValue);
      ensureOption(ampEl, ampStr, ampStr);
      ampEl.value = ampStr;
    } else {
      ampEl.value = '';
    }
  }
  const ledereEl = $('ledere');
  if (ledereEl && !seriesLocksLedere(input.series || '')){
    ledereEl.value = input.ledere || '';
  }
  const startEl = $('startEl');
  if (startEl) startEl.value = input.startEl || '';
  const sluttEl = $('sluttEl');
  if (sluttEl){
    const value = input.sluttEl || '';
    if (!seriesSupportsCrtFeed(input.series) && value === 'crt_board_feed'){
      sluttEl.value = '';
    } else {
      sluttEl.value = value;
    }
  }
  const boxSel = $('boxSel');
  if (boxSel){
    const value = input.boxSel || '';
    if (value && ![...boxSel.options].some(opt=>opt.value === value)){
      ensureOption(boxSel, value, value);
    }
    boxSel.value = value;
  }
  pendingBoxItems = normalizeBoxItems(input.boxItems, input.boxSel, input.boxQty, input.boxInnmatSum);
  renderPendingBoxItems();
  pendingSpecialElementItems = normalizeSpecialElementItems(input.specialElementItems);
  const specialElementEl = $('specialElement');
  if (specialElementEl){
    specialElementEl.value = pendingSpecialElementItems.length || input.specialElement === 'Ja' ? 'Ja' : 'Nei';
  }
  pendingSpecialElementItems.forEach(item=>{
    ensureOption($('specialElementType'), item.selection, item.selection);
  });
  renderPendingSpecialElementItems();
  const freightSelect = $('freightRate');
  if (freightSelect){
    const freightValue = Number(input.freightRate);
    if (Number.isFinite(freightValue)){
      const value = freightValue.toFixed(2);
      ensureOption(freightSelect, value, `${Math.round(freightValue * 100)} % (lagret)`);
      freightSelect.value = value;
    }
  }
  setCurrentMarginRate(resolveMarginRateFromData({ input }));
  setCurrentMontasjeMarginRate(resolveMontasjeDgRate(input?.montasjeMarginRate, NaN, DEFAULT_MARGIN_RATE));
  setCurrentEngineeringMarginRate(resolveDgRate(input?.engineeringMarginRate, NaN, DEFAULT_MARGIN_RATE));
  setCurrentOpphengMarginRate(resolveDgRate(input?.opphengMarginRate, NaN, DEFAULT_MARGIN_RATE));
  setCurrentTapOffMarginRate(resolveDgRate(input?.tapOffMarginRate, NaN, DEFAULT_MARGIN_RATE));
  const montasjeInput = $('montasjeHourlyRate');
  const opphengInput = $('opphengRate');
  const rateToggle = $('rateToggle');
  if (input.montasjeSettings){
    const hourlyRate = Number(input.montasjeSettings.hourlyRate);
    if (montasjeInput && Number.isFinite(hourlyRate)){
      montasjeInput.value = String(hourlyRate);
    }
    if (opphengInput){
      const rate = Number(input.montasjeSettings.opphengRate);
      opphengInput.value = Number.isFinite(rate) ? String(rate) : '';
    }
    if (rateToggle){
      const custom = Number.isFinite(hourlyRate) && hourlyRate !== DEFAULT_HOURLY_RATE
        || Number(input.montasjeSettings.opphengRate) > 0;
      rateToggle.checked = custom;
    }
    const locked = !(rateToggle ? rateToggle.checked : true);
    if (montasjeInput){
      setInputLocked(montasjeInput, locked);
      if (locked){
        montasjeInput.value = String(DEFAULT_HOURLY_RATE);
      }
    }
    if (opphengInput){
      setInputLocked(opphengInput, locked);
      if (locked){
        opphengInput.value = '';
      }
    }
  }


  const engineeringInput = $('engineeringHourlyRate');
  const engineeringToggle = $('engineeringRateToggle');
  if (input.engineeringSettings){
    const hourlyRate = Number(input.engineeringSettings.hourlyRate);
    if (engineeringInput && Number.isFinite(hourlyRate)){
      engineeringInput.value = String(hourlyRate);
    }
    if (engineeringToggle){
      const custom = Number.isFinite(hourlyRate) && hourlyRate !== DEFAULT_ENGINEERING_HOURLY_RATE;
      engineeringToggle.checked = custom;
    }
    const locked = !(engineeringToggle ? engineeringToggle.checked : true);
    if (engineeringInput){
      setInputLocked(engineeringInput, locked);
      if (locked){
        engineeringInput.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
      }
    }
  } else {
    if (engineeringToggle){
      engineeringToggle.checked = false;
    }
    if (engineeringInput){
      engineeringInput.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
      setInputLocked(engineeringInput, true);
    }
  }
  updateMontasjePreview();
  updateTapOffConfigVisibility();
  updateSpecialElementConfigVisibility();
}

function applySavedTotalsToUI(line){
  if (!line || !line.totals) return;
  const totals = line.totals;
  applySelectedAddonCheckboxes(line);
  const savedMarginRate = resolveMarginRateFromData({ totals, input: line.inputs });
  const savedMontasjeMarginRate = resolveMontasjeDgRate(line.inputs?.montasjeMarginRate, totals?.montasjeMarginRate, DEFAULT_MARGIN_RATE);
  const savedEngineeringMarginRate = resolveDgRate(line.inputs?.engineeringMarginRate, totals?.engineeringMarginRate, DEFAULT_MARGIN_RATE);
  const savedOpphengMarginRate = resolveDgRate(line.inputs?.opphengMarginRate, totals?.opphengMarginRate, DEFAULT_MARGIN_RATE);
  const savedTapOffMarginRate = resolveDgRate(line.inputs?.tapOffMarginRate, totals?.tapOffMarginRate, savedMarginRate);
  setCurrentMarginRate(savedMarginRate);
  setCurrentMontasjeMarginRate(savedMontasjeMarginRate);
  setCurrentEngineeringMarginRate(savedEngineeringMarginRate);
  setCurrentOpphengMarginRate(savedOpphengMarginRate);
  setCurrentTapOffMarginRate(savedTapOffMarginRate);
  const setText = (id, value)=>{
    const el = $(id);
    if (!el) return;
    const num = Number(value);
    el.textContent = Number.isFinite(num) ? fmtNO.format(num) : '--';
  };
  setText('mat', totals.material);
  setText('margin', totals.margin);
  setText('subtotal', totals.subtotal);
  setText('freight', totals.freight);
  setText('totalExMontasje', totals.totalExMontasje);
  const montasjePricing = calculateMontasjePricing(totals.montasje?.cost, savedMontasjeMarginRate);
  const fallbackTotalInclMontasje = round2(Number(montasjePricing.totalWithDg || 0));
  setText('totalInclMontasje', totals.totalInclMontasje ?? fallbackTotalInclMontasje);
  const engineeringPricing = calculateDgPricing(totals.engineering?.cost, savedEngineeringMarginRate);
  const fallbackTotalInclEngineering = round2(Number(engineeringPricing.totalWithDg || 0));
  setText('totalInclEngineering', totals.totalInclEngineering ?? fallbackTotalInclEngineering);
  const hasSavedOpphengDg = Number.isFinite(Number(totals.totalInclOppheng))
    || Number.isFinite(Number(totals.opphengMargin))
    || Number.isFinite(Number(totals.opphengMarginRate))
    || Number.isFinite(Number(line.inputs?.opphengMarginRate));
  const effectiveOpphengRate = hasSavedOpphengDg ? savedOpphengMarginRate : 0;
  const opphengPricing = calculateDgPricing(totals.oppheng?.cost, effectiveOpphengRate);
  const fallbackOpphengTotal = round2(Number(opphengPricing.totalWithDg || 0));
  const resolvedOpphengTotal = Number.isFinite(Number(totals.totalInclOppheng))
    ? Number(totals.totalInclOppheng)
    : (totals.total ?? fallbackOpphengTotal);
  setText('total', resolvedOpphengTotal);
  const montasjeEl = $('montasje');
  if (montasjeEl && totals.montasje){
    const cost = Number(totals.montasje.cost);
    montasjeEl.textContent = Number.isFinite(cost) ? fmtNO.format(cost) : '--';
  }
  const opphengEl = $('oppheng');
  if (opphengEl && totals.oppheng){
    const cost = Number(totals.oppheng.cost);
    opphengEl.textContent = Number.isFinite(cost) ? fmtNO.format(cost) : '--';
  }
  const engineeringEl = $('engineering');
  if (engineeringEl){
    const cost = Number(totals.engineering?.cost);
    engineeringEl.textContent = Number.isFinite(cost) ? fmtNO.format(cost) : '--';
  }
  const montasjeCost = Number(totals.montasje?.cost);
  const montasjeMarginVal = Number.isFinite(montasjeCost) ? (totals.montasjeMargin ?? montasjePricing.dg) : NaN;
  setText('montasjeMargin', montasjeMarginVal);
  const engineeringCost = Number(totals.engineering?.cost);
  const engineeringMarginVal = Number.isFinite(engineeringCost) ? (totals.engineeringMargin ?? engineeringPricing.dg) : NaN;
  setText('engineeringMargin', engineeringMarginVal);
  const opphengCost = Number(totals.oppheng?.cost);
  const opphengMarginVal = Number.isFinite(opphengCost)
    ? (hasSavedOpphengDg ? (totals.opphengMargin ?? opphengPricing.dg) : 0)
    : NaN;
  setText('opphengMargin', opphengMarginVal);
  const montasjeDetailEl = $('montasjeDetail');
  if (montasjeDetailEl) montasjeDetailEl.textContent = totals.montasjeDetail || '';
  const opphengDetailEl = $('opphengDetail');
  if (opphengDetailEl) opphengDetailEl.textContent = totals.opphengDetail || '';
  const engineeringDetailEl = $('engineeringDetail');
  if (engineeringDetailEl) engineeringDetailEl.textContent = totals.engineeringDetail || '';
  const resultsEl = $('results');
  if (resultsEl) resultsEl.hidden = false;
  renderBomTable('bomTbl', Array.isArray(line.bom) ? line.bom : []);
  updateXapComparisonUI(null);
  lastCalc = deepClone(totals);
  if (lastCalc){
    lastCalc.bom = Array.isArray(line.bom) ? deepClone(line.bom) : [];
    lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
    lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
    lastCalc.lineNumber = line.lineNumber || '';
    lastCalc.marginRate = savedMarginRate;
    lastCalc.marginFactor = marginFactorFromRate(savedMarginRate);
    lastCalc.montasjeMarginRate = savedMontasjeMarginRate;
    lastCalc.engineeringMarginRate = savedEngineeringMarginRate;
    lastCalc.opphengMarginRate = savedOpphengMarginRate;
    lastCalc.tapOffMarginRate = savedTapOffMarginRate;
  }
  lastCalcInput = line.inputs ? deepClone(line.inputs) : null;
  if (lastCalcInput){
    lastCalcInput.marginRate = savedMarginRate;
    lastCalcInput.montasjeMarginRate = savedMontasjeMarginRate;
    lastCalcInput.engineeringMarginRate = savedEngineeringMarginRate;
    lastCalcInput.opphengMarginRate = savedOpphengMarginRate;
    lastCalcInput.tapOffMarginRate = savedTapOffMarginRate;
  }
  const totalsForPayload = deepClone(totals);
  if (totalsForPayload){
    totalsForPayload.marginRate = savedMarginRate;
    totalsForPayload.montasjeMarginRate = savedMontasjeMarginRate;
    totalsForPayload.engineeringMarginRate = savedEngineeringMarginRate;
    totalsForPayload.opphengMarginRate = savedOpphengMarginRate;
    totalsForPayload.tapOffMarginRate = savedTapOffMarginRate;
    totalsForPayload.tapOffOfferTotal = calculateTapOffOfferTotal({ bom: line.bom || [], tapOffMarginRate: savedTapOffMarginRate });
    totalsForPayload.specialElementOfferTotal = calculateSpecialElementOfferTotal({ bom: line.bom || [], tapOffMarginRate: savedTapOffMarginRate });
  }
  lastEmailPayload = {
    project: projectState.currentProject,
    customer: projectState.currentCustomer,
    lineNumber: line.lineNumber || '',
    inputs: line.inputs ? deepClone(line.inputs) : null,
    totals: totalsForPayload,
    bom: line.bom ? deepClone(line.bom) : []
  };
  if (lastEmailPayload.inputs){
    lastEmailPayload.inputs.marginRate = savedMarginRate;
    lastEmailPayload.inputs.montasjeMarginRate = savedMontasjeMarginRate;
    lastEmailPayload.inputs.engineeringMarginRate = savedEngineeringMarginRate;
    lastEmailPayload.inputs.opphengMarginRate = savedOpphengMarginRate;
    lastEmailPayload.inputs.tapOffMarginRate = savedTapOffMarginRate;
  }
  updateEngineeringPreview();
  const sendBtn = $('sendRequestBtn');
  if (sendBtn){
    sendBtn.disabled = false;
    sendBtn.textContent = 'Send forespørsel';
  }
  const saveBtn = $('saveLineBtn');
  if (saveBtn){
    saveBtn.disabled = false;
  }
  updateSelectedAddonTotalUI();
  markClean();
}

function openProjectLine(projectId, lineId){
  const project = getProjectById(projectId);
  if (!project) return;
  if (!hasCalculatorUI()){
    goToCalculator({ project: project.id, line: lineId });
    return;
  }
  projectState.expandedProjectId = projectId;
  renderProjectDashboard();
  const line = project.lines?.find(entry=>entry.id === lineId);
  if (!line) return;
  setActiveProject(project);
  showCalculatorView();
  if (line.inputs){
    applyInputsToCalculator(line.inputs);
  } else {
    resetCalculatorForm({ preserveProject: true });
  }
  projectState.currentLineNumber = line.lineNumber || '';
  const lineInput = $('lineNumberInput');
  if (lineInput){
    lineInput.value = line.lineNumber || '';
  }
  applySelectedAddonCheckboxes(line);
  if (line.totals){
    applySavedTotalsToUI(line);
    const statusEl = $('status');
    if (statusEl){
      statusEl.textContent = `Linje ${line.lineNumber || ''} er lastet. Beregn på nytt ved endringer.`;
    }
  } else {
    const statusEl = $('status');
    if (statusEl){
      statusEl.textContent = 'Linjen har ingen lagrede summer. Gjør endringer og beregn.';
    }
    const resultsEl = $('results');
    if (resultsEl) resultsEl.hidden = true;
    renderBomTable('bomTbl', []);
    updateXapComparisonUI(null);
  }
}

function showDashboardView(options = {}){
  if (!hasDashboardUI()){
    goToDashboard();
    return;
  }
  if (options.clearSelection){
    projectState.expandedProjectId = null;
  }
  if (options.forceDashboardPage || dashboardState.activePage === 'projects'){
    setDashboardPage('projects');
  }
  const dash = $('dashboardView');
  if (dash && dashboardState.activePage === 'projects') dash.hidden = false;
  renderProjectDashboard();
}

function showProjectOverview(projectId){
  const project = projectId ? getProjectById(projectId) : null;
  if (project){
    setActiveProject(project);
    projectState.expandedProjectId = project.id;
  } else {
    projectState.expandedProjectId = null;
  }
  if (!hasDashboardUI()){
    if (project){
      goToDashboard({ project: project.id, focusProject: '1' });
    } else {
      goToDashboard();
    }
    return;
  }
  showDashboardView({ clearSelection: !project });
  if (project){
    setTimeout(()=>scrollProjectIntoView(project.id), 0);
  }
}

function showCalculatorView(){
  if (!hasCalculatorUI()){
    if (projectState.currentProjectId){
      goToCalculator({ project: projectState.currentProjectId });
    } else {
      goToCalculator();
    }
    return;
  }
  const calc = $('calculatorView');
  if (calc) calc.hidden = false;
  updateProjectMetaDisplay();
}

function saveCurrentLineToProject(){
  if (!hasActiveProject()){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Opprett nytt prosjekt for \u00E5 lagre linjen.';
    openProjectModal({ mode: 'create', saveLineAfterCreate: true });
    return;
  }
  if (!lastCalc){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Kj\u00F8r en beregning f\u00F8r du lagrer.';
    return;
  }
  const lineInput = $('lineNumberInput');
  const lineValue = (lineInput?.value || projectState.currentLineNumber || '').trim();
  if (!lineValue){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Oppgi linjenummer f\u00F8r du lagrer.';
    if (lineInput) lineInput.focus();
    return;
  }
  const project = getProjectById(projectState.currentProjectId);
  if (!project) return;
  const now = new Date().toISOString();
  const selectedAddonFlags = getOfferAddonCheckboxValuesFromUI();
  const selectedAddonTotal = round2(calculateSelectedAddonTotal(lastCalc).total);
  const totalsSnapshot = deepClone(lastCalc) || {};
  totalsSnapshot.selectedAddonTotal = selectedAddonTotal;
  totalsSnapshot.selectedAddonConfig = deepClone(selectedAddonFlags);
  const entryData = {
    id: generateProjectId(),
    lineNumber: lineValue,
    createdAt: now,
    updatedAt: now,
    totals: totalsSnapshot,
    selectedAddonTotal,
    selectedAddonConfig: selectedAddonFlags,
    inputs: lastCalcInput ? deepClone(lastCalcInput) : null,
    bom: Array.isArray(lastEmailPayload?.bom) ? deepClone(lastEmailPayload.bom) : []
  };
  const normalized = lineValue.toLowerCase();
  const existingIdx = project.lines.findIndex(line=>String(line.lineNumber||'').toLowerCase() === normalized);
  let message = `Linje ${lineValue} lagret.`;
  if (existingIdx >= 0){
    const existing = project.lines[existingIdx];
    entryData.id = existing.id || entryData.id;
    entryData.createdAt = existing.createdAt || entryData.createdAt;
    project.lines[existingIdx] = entryData;
    message = `Linje ${lineValue} oppdatert.`;
  } else {
    project.lines.push(entryData);
  }
  project.updatedAt = now;
  saveProjectsToStorage();
  setActiveProject(project);
  projectState.currentLineNumber = '';
  if (lineInput){
    lineInput.value = '';
  }
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = message;
  showProjectOverview(project.id);
}

function resetCalculatorForm(options = {}){
  const { preserveProject = true } = options;
  ['series','dist','ampSelect','ledere','startEl','sluttEl','boxSel','specialElementType'].forEach(id=>{
    const el = $(id);
    if (el){
      el.value = '';
      el.disabled = false;
    }
  });
  ['meter','v90h','v90v','fbQty','boxQty','boxInnmatSum','specialElementQty','specialElementSum'].forEach(id=>{
    const el = $(id);
    if (el){
      el.value = '';
    }
  });
  pendingBoxItems = [];
  renderPendingBoxItems();
  pendingSpecialElementItems = [];
  const specialElementEl = $('specialElement');
  if (specialElementEl) specialElementEl.value = 'Nei';
  renderPendingSpecialElementItems();
  renderTapOffOfferRows();
  renderSpecialElementOfferRows();
  const lineNumberEl = $('lineNumberInput');
  if (lineNumberEl){
    lineNumberEl.value = '';
  }
  projectState.currentLineNumber = '';
  const res = document.getElementById('results');
  if (res) res.hidden = true;
  const st = document.getElementById('status');
  if (st){
    st.textContent = '';
  }
  const rateInput = $('montasjeHourlyRate');
  const rateToggle = $('rateToggle');
  if (rateToggle) rateToggle.checked = false;
  if (rateInput){
    rateInput.value = String(DEFAULT_HOURLY_RATE);
    setInputLocked(rateInput, true);
  }
  const opphengInput = $('opphengRate');
  if (opphengInput){
    opphengInput.value = '';
    opphengInput.dataset.appliedValue = '';
    opphengInput.dataset.userOverride = 'false';
    opphengInput.dataset.defaultRate = '';
    opphengInput.dataset.appliedAmp = '';
    setInputLocked(opphengInput, true);
  }
  const engineeringRateInput = $('engineeringHourlyRate');
  const engineeringRateToggle = $('engineeringRateToggle');
  if (engineeringRateToggle) engineeringRateToggle.checked = false;
  if (engineeringRateInput){
    engineeringRateInput.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
    setInputLocked(engineeringRateInput, true);
  }
  const includeMontasje = $('includeMontasje');
  const includeEngineering = $('includeEngineering');
  const includeOppheng = $('includeOppheng');
  const showMontasje = $('showMontasje');
  const showEngineering = $('showEngineering');
  const showOppheng = $('showOppheng');
  const includeUnitPrices = $('includeUnitPrices');
  if (includeMontasje) includeMontasje.checked = true;
  if (includeEngineering) includeEngineering.checked = true;
  if (includeOppheng) includeOppheng.checked = true;
  if (showMontasje) showMontasje.checked = false;
  if (showEngineering) showEngineering.checked = false;
  if (showOppheng) showOppheng.checked = false;
  if (includeUnitPrices) includeUnitPrices.checked = false;
  applyOfferAddonCheckboxConstraints();
  if (preserveProject && projectState.currentProjectId){
    const project = getProjectById(projectState.currentProjectId);
    if (project){
      applyProjectAddonCheckboxesToCalculator(project);
    }
  }
  refreshUIBySeries();
  updateTapOffConfigVisibility();
  updateSpecialElementConfigVisibility();
  setCurrentMarginRate(DEFAULT_MATERIAL_MARGIN_RATE);
  setCurrentMontasjeMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentEngineeringMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentOpphengMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentTapOffMarginRate(DEFAULT_MARGIN_RATE);
  updateMontasjePreview();
  updateXapComparisonUI(null);
  renderBomTable('bomTbl', []);
  if (!preserveProject){
    clearActiveProject();
  }
  lastCalc = null;
  lastCalcInput = null;
  lastEmailPayload = null;
  isDirty = false;
  const sendBtnReset = document.getElementById('sendRequestBtn');
  if (sendBtnReset){
    sendBtnReset.disabled = true;
    sendBtnReset.textContent = 'Send foresp\u00f8rsel';
  }
  updateSelectedAddonTotalUI();
}

function submitProjectModal(){
  const projectInput = $('projectNameInput');
  const customerInput = $('customerNameInput');
  const contactInput = $('contactPersonInput');
  const errorEl = $('projectError');
  const projectName = projectInput ? projectInput.value.trim() : '';
  const customerName = customerInput ? customerInput.value.trim() : '';
  const contactPerson = contactInput ? contactInput.value.trim() : '';
  if (!projectName || !customerName || !contactPerson){
    if (errorEl) errorEl.textContent = 'Fyll ut alle feltene.';
    updateProjectSubmitState();
    return;
  }
  const wasEditMode = projectModalState.mode === 'edit';
  const wasCopyMode = projectModalState.mode === 'copy';
  const shouldSaveLineAfterCreate = !wasEditMode && !wasCopyMode && projectModalState.saveLineAfterCreate;
  const knownDetails = getKnownProjectDetails(customerName, contactPerson);
  if (!wasEditMode && (!knownDetails.customerAddress || !knownDetails.contactPhone)){
    projectModalState.pendingDetails = {
      projectName,
      customerName,
      contactPerson,
      customerAddress: knownDetails.customerAddress || '',
      customerPostalPlace: knownDetails.customerPostalPlace || '',
      contactPhone: knownDetails.contactPhone || '',
      shouldSaveLineAfterCreate,
      copySourceProjectId: wasCopyMode ? projectModalState.copySourceProjectId : null
    };
    openProjectDetailsModal(projectModalState.pendingDetails);
    return;
  }
  if (wasEditMode && projectModalState.projectId){
    persistProjectInfo(projectName, customerName, contactPerson, {
      projectId: projectModalState.projectId,
      customerAddress: knownDetails.customerAddress,
      customerPostalPlace: knownDetails.customerPostalPlace,
      contactPhone: knownDetails.contactPhone
    });
  } else if (wasCopyMode && projectModalState.copySourceProjectId){
    persistProjectInfo(projectName, customerName, contactPerson, {
      copySourceProjectId: projectModalState.copySourceProjectId,
      customerAddress: knownDetails.customerAddress,
      customerPostalPlace: knownDetails.customerPostalPlace,
      contactPhone: knownDetails.contactPhone
    });
  } else {
    persistProjectInfo(projectName, customerName, contactPerson, knownDetails);
  }
  closeProjectModal();
  if (shouldSaveLineAfterCreate){
    saveCurrentLineToProject();
    return;
  }
  if (!wasEditMode){
    showProjectOverview(projectState.currentProjectId);
  }
}

function openProjectDetailsModal(details){
  const modal = $('projectDetailsModal');
  if (!modal) return;
  const projectModalEl = $('projectModal');
  if (projectModalEl) projectModalEl.style.display = 'none';
  const customerEl = $('projectDetailsCustomer');
  const addressEl = $('projectDetailsAddress');
  const postalPlaceEl = $('projectDetailsPostalPlace');
  const contactEl = $('projectDetailsContact');
  const phoneEl = $('projectDetailsPhone');
  const errorEl = $('projectDetailsError');
  if (customerEl) customerEl.value = details.customerName || '';
  if (addressEl) addressEl.value = details.customerAddress || '';
  if (postalPlaceEl) postalPlaceEl.value = details.customerPostalPlace || '';
  if (contactEl) contactEl.value = details.contactPerson || '';
  if (phoneEl) phoneEl.value = details.contactPhone || '';
  if (errorEl) errorEl.textContent = '';
  modal.style.display = 'flex';
  if (addressEl && !addressEl.value) addressEl.focus();
  else if (postalPlaceEl && !postalPlaceEl.value) postalPlaceEl.focus();
  else if (phoneEl) phoneEl.focus();
}

function closeProjectDetailsModal(options = {}){
  const modal = $('projectDetailsModal');
  if (modal) modal.style.display = 'none';
  const errorEl = $('projectDetailsError');
  if (errorEl) errorEl.textContent = '';
  if (!options.keepPending){
    projectModalState.pendingDetails = null;
  }
}

function submitProjectDetailsModal(){
  const pending = projectModalState.pendingDetails;
  if (!pending) return;
  const address = ($('projectDetailsAddress')?.value || '').trim();
  const postalPlace = ($('projectDetailsPostalPlace')?.value || '').trim();
  const phone = ($('projectDetailsPhone')?.value || '').trim();
  const errorEl = $('projectDetailsError');
  if (!address || !postalPlace || !phone){
    if (errorEl) errorEl.textContent = 'Fyll ut adresse, postnummer/sted og telefon.';
    return;
  }
  persistProjectInfo(pending.projectName, pending.customerName, pending.contactPerson, {
    copySourceProjectId: pending.copySourceProjectId || null,
    customerAddress: address,
    customerPostalPlace: postalPlace,
    contactPhone: phone
  });
  closeProjectDetailsModal();
  closeProjectModal();
  if (pending.shouldSaveLineAfterCreate){
    saveCurrentLineToProject();
    return;
  }
  showProjectOverview(projectState.currentProjectId);
}

function cancelProjectModal(){
  closeProjectModal();
}

const projectSubmit = $('projectSubmit');
if (projectSubmit){
  projectSubmit.addEventListener('click', submitProjectModal);
}
const projectCancel = $('projectCancel');
if (projectCancel){
  projectCancel.addEventListener('click', cancelProjectModal);
}
const projectModal = $('projectModal');
if (projectModal){
  projectModal.addEventListener('click', evt=>{
    if (evt.target === projectModal){
      cancelProjectModal();
    }
  });
}
const projectDetailsCancel = $('projectDetailsCancel');
if (projectDetailsCancel){
  projectDetailsCancel.addEventListener('click', ()=>closeProjectDetailsModal());
}
const projectDetailsSubmit = $('projectDetailsSubmit');
if (projectDetailsSubmit){
  projectDetailsSubmit.addEventListener('click', submitProjectDetailsModal);
}
['projectDetailsAddress','projectDetailsPostalPlace','projectDetailsPhone'].forEach(id=>{
  const input = $(id);
  if (!input) return;
  input.addEventListener('keydown', evt=>{
    if (evt.key === 'Enter'){
      evt.preventDefault();
      submitProjectDetailsModal();
    } else if (evt.key === 'Escape'){
      evt.preventDefault();
      closeProjectDetailsModal();
    }
  });
});
const projectDetailsModal = $('projectDetailsModal');
if (projectDetailsModal){
  projectDetailsModal.addEventListener('click', evt=>{
    if (evt.target === projectDetailsModal){
      closeProjectDetailsModal();
    }
  });
}
const offerDetailsCancel = $('offerDetailsCancel');
if (offerDetailsCancel){
  offerDetailsCancel.addEventListener('click', ()=>closeOfferDetailsWarning(false));
}
const offerDetailsContinue = $('offerDetailsContinue');
if (offerDetailsContinue){
  offerDetailsContinue.addEventListener('click', ()=>closeOfferDetailsWarning(true));
}
const offerDetailsWarningModal = $('offerDetailsWarningModal');
if (offerDetailsWarningModal){
  offerDetailsWarningModal.addEventListener('click', evt=>{
    if (evt.target === offerDetailsWarningModal){
      closeOfferDetailsWarning(false);
    }
  });
}

const projectMarginCancelBtn = $('projectMarginCancel');
if (projectMarginCancelBtn){
  projectMarginCancelBtn.addEventListener('click', closeProjectMarginModal);
}
const projectMarginSubmitBtn = $('projectMarginSubmit');
if (projectMarginSubmitBtn){
  projectMarginSubmitBtn.addEventListener('click', submitProjectMarginModal);
}
const projectMarginPercentInput = $('projectMarginPercentInput');
if (projectMarginPercentInput){
  projectMarginPercentInput.addEventListener('keydown', evt=>{
    if (evt.key === 'Enter'){
      evt.preventDefault();
      submitProjectMarginModal();
    } else if (evt.key === 'Escape'){
      evt.preventDefault();
      closeProjectMarginModal();
    }
  });
}
const projectMarginModal = $('projectMarginModal');
if (projectMarginModal){
  projectMarginModal.addEventListener('click', evt=>{
    if (evt.target === projectMarginModal){
      closeProjectMarginModal();
    }
  });
}

const projectStatusModal = $('projectStatusModal');
if (projectStatusModal){
  projectStatusModal.addEventListener('click', evt=>{
    if (evt.target === projectStatusModal){
      closeProjectStatusModal();
      return;
    }
  });
  projectStatusModal.querySelectorAll('[data-project-status-option]').forEach(optionBtn=>{
    optionBtn.addEventListener('click', evt=>{
      evt.preventDefault();
      evt.stopPropagation();
      const projectId = projectStatusModalState.projectId || projectStatusModal.dataset.projectId;
      if (!projectId) return;
      setProjectStatus(projectId, optionBtn.getAttribute('data-project-status-option'));
    });
  });
}

const projectStatusCancel = $('projectStatusCancel');
if (projectStatusCancel){
  projectStatusCancel.addEventListener('click', closeProjectStatusModal);
}

const newProjectBtn = $('newProjectBtn');
if (newProjectBtn){
  newProjectBtn.addEventListener('click', ()=>{
    if (!authState.loggedIn){
      showLoginModal();
      return;
    }
    openProjectModal({ mode: 'create' });
  });
}

const projectSortSelect = $('projectSortSelect');
if (projectSortSelect){
  projectSortSelect.addEventListener('change', ()=>{
    setProjectSortMode(projectSortSelect.value);
  });
}

const projectSearchInput = $('projectSearchInput');
if (projectSearchInput){
  projectSearchInput.addEventListener('input', ()=>{
    setProjectSearchTerm(projectSearchInput.value);
  });
}

const projectArchiveToggleBtn = $('projectArchiveToggleBtn');
if (projectArchiveToggleBtn){
  projectArchiveToggleBtn.addEventListener('click', toggleProjectArchiveView);
}

const lineSortSelect = $('lineSortSelect');
if (lineSortSelect){
  lineSortSelect.addEventListener('change', ()=>{
    setLineSortMode(lineSortSelect.value);
  });
}

const projectListEl = $('projectList');
if (projectListEl){
  projectListEl.addEventListener('click', evt=>{
    const target = evt.target.closest('button');
    if (!target) return;
    if (target.dataset.lineEdit){
      openProjectLine(target.dataset.projectId, target.dataset.lineEdit);
      return;
    }
    if (target.dataset.lineDelete){
      deleteProjectLine(target.dataset.projectId, target.dataset.lineDelete);
      return;
    }
    if (target.dataset.projectCardEdit){
      openProjectModal({ mode: 'edit', projectId: target.dataset.projectCardEdit });
      return;
    }
    if (target.dataset.projectStatusEdit){
      openProjectStatusModal(target.dataset.projectStatusEdit);
      return;
    }
    if (target.dataset.projectDelete){
      deleteProject(target.dataset.projectDelete);
      return;
    }
    if (target.dataset.projectDetail){
      selectProjectDetail(target.dataset.projectDetail);
      return;
    }
    if (target.dataset.projectNewline){
      startNewLineForProject(target.dataset.projectNewline);
      return;
    }
    if (target.dataset.projectGenerateOffer){
      requestGenerateProjectOffer(target.dataset.projectGenerateOffer, target);
      return;
    }
    if (target.dataset.projectCreateFolder){
      void createProjectFolderFromTemplate(target.dataset.projectCreateFolder, target);
      return;
    }
    if (target.dataset.projectOpenFolder){
      void openProjectSharePointFolder(target.dataset.projectOpenFolder, target);
      return;
    }
    if (target.dataset.projectCopy){
      openProjectModal({ mode: 'copy', copySourceProjectId: target.dataset.projectCopy });
      return;
    }
    if (target.dataset.projectSetMargin){
      promptAndApplyProjectMarginRate(target.dataset.projectSetMargin);
      return;
    }
    if (target.dataset.action === 'create-project'){
      if (!authState.loggedIn){
        showLoginModal();
        return;
      }
      openProjectModal({ mode: 'create' });
    }
  });
  projectListEl.addEventListener('change', evt=>{
    const target = evt.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.dataset.lineAddon){
      const projectId = target.dataset.projectId || '';
      const lineId = target.dataset.lineId || '';
      const field = target.dataset.addonField || '';
      if (!projectId || !lineId || !field) return;
      updateLineSelectedAddonConfig(projectId, lineId, { [field]: target.checked });
      return;
    }
    if (target.dataset.projectAddon){
      const projectId = target.dataset.projectId || '';
      const field = target.dataset.addonField || '';
      if (!projectId || !field) return;
      updateProjectSelectedAddonConfig(projectId, { [field]: target.checked });
    }
  });
}

const projectFlowProjectSelect = $('projectFlowProjectSelect');
if (projectFlowProjectSelect){
  projectFlowProjectSelect.addEventListener('change', ()=>{
    projectFlowState.selectedProjectId = projectFlowProjectSelect.value || '';
    closeProjectFlowMilestoneForm();
    renderProjectFlowView();
  });
}

const projectFlowTaskProjectSelect = $('projectFlowTaskProjectSelect');
if (projectFlowTaskProjectSelect){
  projectFlowTaskProjectSelect.addEventListener('change', ()=>{
    populateProjectFlowDependencyTaskSelect();
  });
}

const addProjectFlowMilestoneBtn = $('addProjectFlowMilestoneBtn');
if (addProjectFlowMilestoneBtn){
  addProjectFlowMilestoneBtn.addEventListener('click', ()=>{
    openProjectFlowMilestoneForm();
  });
}

const refreshProjectFlowBtn = $('refreshProjectFlowBtn');
if (refreshProjectFlowBtn){
  refreshProjectFlowBtn.addEventListener('click', ()=>{
    renderProjectFlowView();
    setProjectFlowStatus('Oppdatert.', 'ok');
  });
}

const projectFlowExpandAllBtn = $('projectFlowExpandAllBtn');
if (projectFlowExpandAllBtn){
  projectFlowExpandAllBtn.addEventListener('click', ()=>{
    setProjectFlowAllExpanded(true);
  });
}

const projectFlowCollapseAllBtn = $('projectFlowCollapseAllBtn');
if (projectFlowCollapseAllBtn){
  projectFlowCollapseAllBtn.addEventListener('click', ()=>{
    setProjectFlowAllExpanded(false);
  });
}

const projectFlowZoomInBtn = $('projectFlowZoomInBtn');
if (projectFlowZoomInBtn){
  projectFlowZoomInBtn.addEventListener('click', ()=>{
    setProjectFlowZoomIndex(projectFlowState.zoomIndex + 1);
  });
}

const projectFlowZoomOutBtn = $('projectFlowZoomOutBtn');
if (projectFlowZoomOutBtn){
  projectFlowZoomOutBtn.addEventListener('click', ()=>{
    setProjectFlowZoomIndex(projectFlowState.zoomIndex - 1);
  });
}

const projectFlowZoomFitBtn = $('projectFlowZoomFitBtn');
if (projectFlowZoomFitBtn){
  projectFlowZoomFitBtn.addEventListener('click', setProjectFlowZoomToFit);
}

const projectFlowMilestoneForm = $('projectFlowMilestoneForm');
if (projectFlowMilestoneForm){
  projectFlowMilestoneForm.addEventListener('submit', handleProjectFlowMilestoneSave);
}

const cancelProjectFlowMilestoneBtn = $('cancelProjectFlowMilestoneBtn');
if (cancelProjectFlowMilestoneBtn){
  cancelProjectFlowMilestoneBtn.addEventListener('click', closeProjectFlowMilestoneForm);
}

const deleteProjectFlowMilestoneBtn = $('deleteProjectFlowMilestoneBtn');
if (deleteProjectFlowMilestoneBtn){
  deleteProjectFlowMilestoneBtn.addEventListener('click', deleteProjectFlowMilestoneFromForm);
}

['projectFlowDateInput', 'projectFlowDurationValueInput', 'projectFlowDurationUnitSelect'].forEach(id=>{
  const el = $(id);
  if (el){
    el.addEventListener('change', updateProjectFlowEndDateFromDuration);
    el.addEventListener('input', updateProjectFlowEndDateFromDuration);
  }
});

const projectFlowEndDateInput = $('projectFlowEndDateInput');
if (projectFlowEndDateInput){
  projectFlowEndDateInput.addEventListener('change', updateProjectFlowDurationFromEndDate);
}

const projectFlowDrivenBySelect = $('projectFlowDependencyRelationSelect');
const projectFlowDrivesSelect = $('projectFlowDependencyTaskSelect');
if (projectFlowDrivenBySelect && projectFlowDrivesSelect){
  projectFlowDrivenBySelect.addEventListener('change', ()=>{
    populateProjectFlowDependencyTaskSelect({
      id: String($('projectFlowMilestoneId')?.value || '').trim(),
      projectId: String($('projectFlowTaskProjectSelect')?.value || '').trim(),
      drivenByTaskId: projectFlowDrivenBySelect.value,
      drivesTaskId: projectFlowDrivesSelect.value
    });
  });
  projectFlowDrivesSelect.addEventListener('change', ()=>{
    populateProjectFlowDependencyTaskSelect({
      id: String($('projectFlowMilestoneId')?.value || '').trim(),
      projectId: String($('projectFlowTaskProjectSelect')?.value || '').trim(),
      drivenByTaskId: projectFlowDrivenBySelect.value,
      drivesTaskId: projectFlowDrivesSelect.value
    });
  });
}

const projectFlowStartDatePickerBtn = $('projectFlowStartDatePickerBtn');
if (projectFlowStartDatePickerBtn){
  projectFlowStartDatePickerBtn.addEventListener('click', evt=>{
    evt.stopPropagation();
    const popover = $('projectFlowDatePickerPopover');
    if (popover && !popover.hidden && projectFlowState.datePickerTargetId === 'projectFlowDateInput'){
      closeProjectFlowDatePickerPopover();
    } else {
      openProjectFlowDatePickerPopover('projectFlowDateInput');
    }
  });
}

const projectFlowEndDatePickerBtn = $('projectFlowEndDatePickerBtn');
if (projectFlowEndDatePickerBtn){
  projectFlowEndDatePickerBtn.addEventListener('click', evt=>{
    evt.stopPropagation();
    const popover = $('projectFlowDatePickerPopover');
    if (popover && !popover.hidden && projectFlowState.datePickerTargetId === 'projectFlowEndDateInput'){
      closeProjectFlowDatePickerPopover();
    } else {
      openProjectFlowDatePickerPopover('projectFlowEndDateInput');
    }
  });
}

const projectFlowDatePickerPopover = $('projectFlowDatePickerPopover');
if (projectFlowDatePickerPopover){
  projectFlowDatePickerPopover.addEventListener('click', evt=>evt.stopPropagation());
}

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target : null;
  if (target?.closest?.('#projectFlowMilestoneForm .calendar-date-input-row')) return;
  closeProjectFlowDatePickerPopover();
});

const projectFlowTimelineEl = $('projectFlowTimeline');
if (projectFlowTimelineEl){
  projectFlowTimelineEl.addEventListener('pointerdown', evt=>{
    const linkHandle = evt.target instanceof Element ? evt.target.closest('[data-project-flow-link]') : null;
    if (linkHandle){
      startProjectFlowLinkDrag(evt, linkHandle);
      return;
    }
    const bar = evt.target instanceof Element ? evt.target.closest('[data-project-flow-drag]') : null;
    if (bar) startProjectFlowTaskDrag(evt, bar);
  });
  projectFlowTimelineEl.addEventListener('pointermove', evt=>{
    updateProjectFlowLinkDrag(evt);
    updateProjectFlowTaskDrag(evt);
  });
  projectFlowTimelineEl.addEventListener('pointerup', evt=>{
    finishProjectFlowLinkDrag(evt);
    finishProjectFlowTaskDrag(evt);
  });
  projectFlowTimelineEl.addEventListener('pointercancel', evt=>{
    finishProjectFlowLinkDrag(evt);
    finishProjectFlowTaskDrag(evt);
  });
  projectFlowTimelineEl.addEventListener('change', evt=>{
    const target = evt.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.dataset.projectFlowToggle) return;
    toggleProjectFlowMilestone(target.dataset.projectId || '', target.dataset.projectFlowToggle, target.checked);
  });
  projectFlowTimelineEl.addEventListener('click', evt=>{
    if (projectFlowState.drag) return;
    if (Date.now() < Number(projectFlowState.suppressClickUntil || 0)){
      evt.preventDefault();
      evt.stopPropagation();
      return;
    }
    const toggleTarget = evt.target instanceof Element ? evt.target.closest('[data-project-flow-toggle]') : null;
    if (toggleTarget) return;
    const phaseToggle = evt.target instanceof Element ? evt.target.closest('[data-project-flow-phase-toggle]') : null;
    if (phaseToggle){
      const phaseId = phaseToggle.getAttribute('data-project-flow-phase-toggle') || '';
      if (projectFlowState.collapsedPhaseIds.has(phaseId)){
        projectFlowState.collapsedPhaseIds.delete(phaseId);
      } else {
        projectFlowState.collapsedPhaseIds.add(phaseId);
      }
      renderProjectFlowView();
      return;
    }
    const deleteTarget = evt.target instanceof Element ? evt.target.closest('[data-project-flow-delete]') : null;
    if (deleteTarget){
      deleteProjectFlowMilestone(
        deleteTarget.getAttribute('data-project-id') || '',
        deleteTarget.getAttribute('data-project-flow-delete') || ''
      );
      return;
    }
    const editTarget = evt.target instanceof Element ? evt.target.closest('[data-project-flow-edit]') : null;
    if (editTarget){
      const projectId = editTarget.getAttribute('data-project-id') || '';
      const milestoneId = editTarget.getAttribute('data-project-flow-edit') || '';
      const milestone = getProjectFlowMilestones(projectId).find(item=>item.id === milestoneId);
      if (milestone) openProjectFlowMilestoneForm(milestone);
      return;
    }
    const dateCell = evt.target instanceof Element ? evt.target.closest('[data-project-flow-date-cell]') : null;
    if (dateCell){
      const projectId = dateCell.getAttribute('data-project-id') || '';
      const taskId = dateCell.getAttribute('data-project-flow-date-cell') || '';
      const clickedDate = dateCell.getAttribute('data-date') || '';
      const phaseId = dateCell.getAttribute('data-phase-id') || PROJECT_FLOW_PHASES[0].id;
      const task = getProjectFlowMilestones(projectId).find(item=>item.id === taskId);
      if (task){
        openProjectFlowMilestoneForm({ ...task, projectId }, { projectId, clickedDate, phaseId: task.phaseId });
      } else if (clickedDate){
        const excludeProjectIds = projectId
          ? []
          : projectState.projects
            .filter(project=>getProjectFlowMilestones(project.id).length)
            .map(project=>project.id);
        openProjectFlowMilestoneForm(null, { projectId, clickedDate, phaseId, excludeProjectIds });
      }
      return;
    }
  });
}

const calendarEventsListEl = $('calendarEventsList');
if (calendarEventsListEl){
  calendarEventsListEl.addEventListener('click', evt=>{
    const target = evt.target instanceof Element ? evt.target.closest('[data-edit-calendar-event]') : null;
    if (!target) return;
    const id = String(target.getAttribute('data-edit-calendar-event') || '').trim();
    const event = calendarViewState.events.find(item=>item.id === id);
    if (event) openCalendarEventForm(event);
  });
}

const calendarGridViewEl = $('calendarGridView');
if (calendarGridViewEl){
  calendarGridViewEl.addEventListener('click', evt=>{
    const target = evt.target instanceof Element ? evt.target.closest('[data-edit-calendar-event]') : null;
    if (!target) return;
    const id = String(target.getAttribute('data-edit-calendar-event') || '').trim();
    const event = calendarViewState.events.find(item=>item.id === id);
    if (event) openCalendarEventForm(event);
  });
}

const emailMessagesListEl = $('emailMessagesList');
if (emailMessagesListEl){
  emailMessagesListEl.addEventListener('click', evt=>{
    const target = evt.target instanceof Element ? evt.target.closest('[data-email-message-id]') : null;
    if (!target) return;
    selectEmailMessage(target.getAttribute('data-email-message-id') || '');
  });
}

const offersListEl = $('offersList');
if (offersListEl){
  offersListEl.addEventListener('click', evt=>{
    const target = evt.target instanceof Element ? evt.target.closest('[data-open-offer-word]') : null;
    if (!target) return;
    void openLatestOfferForProject(target.getAttribute('data-open-offer-word') || '', target);
  });
}

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target.closest('[data-delete-sharepoint-item]') : null;
  if (!target) return;
  const page = target.getAttribute('data-sharepoint-page') || '';
  const id = target.getAttribute('data-delete-sharepoint-item') || '';
  const name = target.getAttribute('data-sharepoint-name') || '';
  void deleteSharePointItem(page, id, name);
});

const saveLineBtn = $('saveLineBtn');
if (saveLineBtn){
  saveLineBtn.disabled = true;
  saveLineBtn.addEventListener('click', saveCurrentLineToProject);
}

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target.closest('[data-delete-tap-off]') : null;
  if (!target) return;
  const groupId = String(target.getAttribute('data-delete-tap-off') || '').trim();
  if (!groupId) return;
  const sourceBom = Array.isArray(lastCalc?.bom) ? lastCalc.bom : (Array.isArray(lastEmailPayload?.bom) ? lastEmailPayload.bom : []);
  const sourceGroupLine = sourceBom.find(entry=>String(entry.tapOffGroupId || '') === groupId && !entry.tapOffInnmatLine);
  const sourceBoxSel = String(sourceGroupLine?.tapOffBoxSel || '').trim();
  const removeBoxItem = items=>{
    const normalized = normalizeBoxItems(items);
    const byId = normalized.filter(item=>String(item.id || '') !== groupId);
    if (byId.length !== normalized.length) return byId;
    if (!sourceBoxSel) return byId;
    let removed = false;
    return normalized.filter(item=>{
      if (!removed && String(item.boxSel || '') === sourceBoxSel){
        removed = true;
        return false;
      }
      return true;
    });
  };
  pendingBoxItems = removeBoxItem(pendingBoxItems);
  if (lastCalcInput){
    lastCalcInput.boxItems = removeBoxItem(lastCalcInput.boxItems);
  }
  let refreshed = false;
  try{
    refreshed = refreshCalculatedBoxItems();
  }catch(err){
    const st = $('status');
    if (st) st.textContent = String(err.message || err);
  }
  if (!refreshed){
    if (lastCalc?.bom){
      lastCalc.bom = lastCalc.bom.filter(entry=>String(entry.tapOffGroupId || '') !== groupId);
      lastCalc.tapOffBoxTotal = sumSeparateTapOffBoxTotal(lastCalc.bom);
      lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
      renderBomTable('bomTbl', lastCalc.bom);
    }
    if (lastEmailPayload?.bom){
      lastEmailPayload.bom = lastEmailPayload.bom.filter(entry=>String(entry.tapOffGroupId || '') !== groupId);
      if (lastEmailPayload.totals){
        lastEmailPayload.totals.tapOffBoxTotal = sumSeparateTapOffBoxTotal(lastEmailPayload.bom);
        lastEmailPayload.totals.tapOffOfferTotal = calculateTapOffOfferTotal({
          bom: lastEmailPayload.bom,
          tapOffMarginRate: lastEmailPayload.totals.tapOffMarginRate ?? lastEmailPayload.inputs?.tapOffMarginRate ?? DEFAULT_MARGIN_RATE
        });
      }
    }
    updateSelectedAddonTotalUI();
  }
  renderPendingBoxItems();
  const st = $('status');
  if (st) st.textContent = 'Bokslinje slettet fra BOM og tilbudssum. Lagre linjen for å beholde endringen.';
  const saveBtn = $('saveLineBtn');
  if (saveBtn) saveBtn.disabled = false;
});

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target.closest('[data-tap-off-dg]') : null;
  if (!target) return;
  openMarginModal('tapoff');
});

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target.closest('[data-special-element-dg]') : null;
  if (!target) return;
  openMarginModal('special');
});

document.addEventListener('click', evt=>{
  const target = evt.target instanceof Element ? evt.target.closest('[data-delete-special-element]') : null;
  if (!target) return;
  const groupId = String(target.getAttribute('data-delete-special-element') || '').trim();
  if (!groupId) return;
  const removeItem = items=>normalizeSpecialElementItems(items)
    .filter(item=>String(item.id || '') !== groupId);
  pendingSpecialElementItems = removeItem(pendingSpecialElementItems);
  if (lastCalcInput){
    lastCalcInput.specialElementItems = removeItem(lastCalcInput.specialElementItems);
  }
  let refreshed = false;
  try{
    refreshed = refreshCalculatedBoxItems();
  }catch(err){
    const st = $('status');
    if (st) st.textContent = String(err.message || err);
  }
  if (!refreshed){
    if (lastCalc?.bom){
      lastCalc.bom = lastCalc.bom.filter(entry=>String(entry.specialElementGroupId || '') !== groupId);
      lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
      renderBomTable('bomTbl', lastCalc.bom);
    }
    if (lastEmailPayload?.bom){
      lastEmailPayload.bom = lastEmailPayload.bom.filter(entry=>String(entry.specialElementGroupId || '') !== groupId);
      if (lastEmailPayload.totals){
        lastEmailPayload.totals.specialElementOfferTotal = calculateSpecialElementOfferTotal({
          bom: lastEmailPayload.bom,
          tapOffMarginRate: lastEmailPayload.totals.tapOffMarginRate ?? lastEmailPayload.inputs?.tapOffMarginRate
        });
      }
    }
    updateSelectedAddonTotalUI();
  }
  renderPendingSpecialElementItems();
  const st = $('status');
  if (st) st.textContent = 'Spesialelement slettet fra BOM og tilbudssum. Lagre linjen for å beholde endringen.';
  const saveBtn = $('saveLineBtn');
  if (saveBtn) saveBtn.disabled = false;
});

const lineNumberInputEl = $('lineNumberInput');
if (lineNumberInputEl){
  lineNumberInputEl.addEventListener('input', ()=>{
    const value = lineNumberInputEl.value.trim();
    projectState.currentLineNumber = value;
    if (lastCalc){
      lastCalc.lineNumber = value;
    }
    if (lastEmailPayload){
      lastEmailPayload.lineNumber = value;
    }
  });
}

const resetBtn = $('resetBtn');
if (resetBtn){
  resetBtn.addEventListener('click', ()=>{
    resetCalculatorForm({ preserveProject: true });
  });
}

['includeMontasje','includeEngineering','includeOppheng'].forEach(id=>{
  const checkbox = $(id);
  if (!checkbox) return;
  checkbox.addEventListener('change', ()=>{
    applyOfferAddonCheckboxConstraints();
    updateSelectedAddonTotalUI();
  });
});
['showMontasje','showEngineering','showOppheng'].forEach(id=>{
  const checkbox = $(id);
  if (!checkbox) return;
  checkbox.addEventListener('change', applyOfferAddonCheckboxConstraints);
});
const includeUnitPricesEl = $('includeUnitPrices');
if (includeUnitPricesEl){
  includeUnitPricesEl.addEventListener('change', updateSelectedAddonTotalUI);
}
applyOfferAddonCheckboxConstraints();
updateSelectedAddonTotalUI();

const marginCancelBtn = $('marginCancel');
if (marginCancelBtn){
  marginCancelBtn.addEventListener('click', closeMarginModal);
}

const marginSubmitBtn = $('marginSubmit');
if (marginSubmitBtn){
  marginSubmitBtn.addEventListener('click', submitMarginModal);
}

const marginPercentInput = $('marginPercentInput');
if (marginPercentInput){
  marginPercentInput.addEventListener('keydown', evt=>{
    if (evt.key === 'Enter'){
      evt.preventDefault();
      submitMarginModal();
    } else if (evt.key === 'Escape'){
      evt.preventDefault();
      closeMarginModal();
    }
  });
}

const marginModal = $('marginModal');
if (marginModal){
  marginModal.addEventListener('click', evt=>{
    if (evt.target === marginModal){
      closeMarginModal();
    }
  });
}

function bindSuggestionBehaviour(inputId, listId, source){
  const input = $(inputId);
  const listEl = $(listId);
  if (!input || !listEl) return;
  const getSource = ()=>typeof source === 'function' ? source() : source;
  input.addEventListener('focus', ()=>{
    showSuggestions(listEl, getSource());
  });
  input.addEventListener('input', ()=>{
    hideSuggestions(listEl);
    updateProjectSubmitState();
    const errorEl = $('projectError');
    if (errorEl) errorEl.textContent = '';
    if (inputId === 'customerNameInput'){
      updateContactSuggestionsForCustomer();
    }
  });
  input.addEventListener('blur', ()=>{
    setTimeout(()=>hideSuggestions(listEl), 80);
  });
  input.addEventListener('keydown', evt=>{
    if (evt.key === 'Enter'){
      evt.preventDefault();
      const submitBtn = $('projectSubmit');
      if (submitBtn && !submitBtn.disabled){
        submitProjectModal();
      }
    } else if (evt.key === 'Escape'){
      evt.preventDefault();
      cancelProjectModal();
    }
  });
  listEl.addEventListener('mousedown', evt=>{
    if (evt.target && evt.target.matches('li[data-value]')){
      evt.preventDefault();
      const value = evt.target.dataset.value || evt.target.textContent || '';
      input.value = value;
      hideSuggestions(listEl);
      if (inputId === 'customerNameInput'){
        updateContactSuggestionsForCustomer();
      }
      updateProjectSubmitState();
    }
  });
}

bindSuggestionBehaviour('projectNameInput', 'projectSuggestions', projectState.projectHistory);
bindSuggestionBehaviour('customerNameInput', 'customerSuggestions', projectState.customerHistory);
bindSuggestionBehaviour('contactPersonInput', 'contactSuggestions', ()=>{
  const customerName = ($('customerNameInput')?.value || '').trim();
  return getContactsForCustomer(customerName);
});

const editProjectBtn = $('editProjectBtn');
if (editProjectBtn){
  editProjectBtn.addEventListener('click', ()=>{
    if (!projectState.currentProjectId) return;
    openProjectModal({ mode: 'edit', projectId: projectState.currentProjectId });
  });
}

document.addEventListener('keydown', evt=>{
  if (evt.key === 'Escape'){
    const loginModalEl = $('loginModal');
    if (loginModalEl && loginModalEl.style.display === 'flex'){
      hideLoginModal();
      return;
    }
    const marginModalEl = $('marginModal');
    if (marginModalEl && marginModalEl.style.display === 'flex'){
      closeMarginModal();
      return;
    }
    const projectMarginModalEl = $('projectMarginModal');
    if (projectMarginModalEl && projectMarginModalEl.style.display === 'flex'){
      closeProjectMarginModal();
      return;
    }
    const projectDetailsModalEl = $('projectDetailsModal');
    if (projectDetailsModalEl && projectDetailsModalEl.style.display === 'flex'){
      closeProjectDetailsModal();
      return;
    }
    const offerDetailsWarningModalEl = $('offerDetailsWarningModal');
    if (offerDetailsWarningModalEl && offerDetailsWarningModalEl.style.display === 'flex'){
      closeOfferDetailsWarning(false);
      return;
    }
    const projectModalEl = $('projectModal');
    if (projectModalEl && projectModalEl.style.display === 'flex'){
      cancelProjectModal();
    }
  }
});

loadAuthFromSession();
updateProjectMetaDisplay();
updateAuthUI();

const H = {
  code: ['code','Code','SKU','sku','produkt','Produkt'],
  price: ['price','Price','unit price','unit_price','Unit Price','UnitPrice','pris','Pris'],
  desc:  ['desc_text','description','Description','desc','tekst','Tekst'],
  amp:   ['ampere','Ampere','amp','Amp'],
  et2:   ['element_type_2','Element type 2','element type 2','H']
};

const DEFAULT_HOURLY_RATE = 700;
const DEFAULT_ENGINEERING_HOURLY_RATE = 929.9;
const ENGINEERING_HOURS_PER_METER = 0.30;
const ENGINEERING_HOURS_PER_TEN_ANGLES = 0.30;
const OPPHENG_RATE_TABLE = Object.freeze([
  { maxAmp: 1600, rate: 400, label: '160–1600A' },
  { maxAmp: 2500, rate: 500, label: '2000–2500A' },
  { maxAmp: Infinity, rate: 800, label: '3200–5000A' }
]);
const MONTASJE_TIME_TABLE = Object.freeze([
  { maxAmp: 250,   hoursPerMeter: 1.5, hoursPerAngle: 0.5,  label: '160–250A' },
  { maxAmp: 630,   hoursPerMeter: 2,   hoursPerAngle: 0.5,  label: '400–630A' },
  { maxAmp: 1600,  hoursPerMeter: 3,   hoursPerAngle: 0.75, label: '800–1600A' },
  { maxAmp: 2500,  hoursPerMeter: 4,   hoursPerAngle: 1,    label: '2000–2500A' },
  { maxAmp: Infinity, hoursPerMeter: 5, hoursPerAngle: 1.5, label: '3200–5000A' }
]);
const EPOXY_MONTASJE_TIME_TABLE = Object.freeze([
  { maxAmp: 1250, hoursPerMeter: 5, hoursPerAngle: 2, label: '1250A' },
  { maxAmp: 2500, hoursPerMeter: 6, hoursPerAngle: 3, label: '1600–2500A' },
  { maxAmp: Infinity, hoursPerMeter: 7, hoursPerAngle: 4, label: '3200–5000A' }
]);
const EXCLUDED_AMP_MIN = 160;
const EXCLUDED_AMP_MAX = 470;
const EPOXY_MONOBLOC_FACTOR_BY_AMP = Object.freeze({
  1250: 0.5,
  1600: 0.5,
  2000: 0.7,
  2500: 1,
  3200: 1,
  4000: 1.25,
  5000: 2
});
const EPOXY_MAIN_ELEMENT_TYPES = new Set([
  'board_feed',
  'crt_board_feed',
  'end_feed_unit',
  'end_cover',
  'straight_3m',
  'straight_3m_dist',
  'straight_1501_2000',
  'straight_1501_2000_dist',
  'straight_500_1000',
  'straight_500_1000_dist',
  'elbow_horizontal_90',
  'elbow_vertical_90'
]);

function sanitizeHourlyRate(value, fallback = DEFAULT_HOURLY_RATE){
  const raw = value ?? '';
  if (String(raw).trim()==='') return fallback;
  const n = toNum(raw);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(0, n);
}

function getOpphengRateRowForAmp(amp){
  const a = Number(amp);
  if (!Number.isFinite(a) || a <= 0) return null;
  for (const row of OPPHENG_RATE_TABLE){
    if (a <= row.maxAmp) return row;
  }
  return OPPHENG_RATE_TABLE[OPPHENG_RATE_TABLE.length - 1];
}

function sanitizeOpphengRate(value, fallback){
  const raw = value ?? '';
  const fallbackRate = Number.isFinite(fallback) ? fallback : 0;
  if (String(raw).trim()==='') return fallbackRate;
  const n = toNum(raw);
  if (!Number.isFinite(n)) return fallbackRate;
  return Math.max(0, n);
}

function calculateOpphengsmateriell({ meter, amp, ratePerPiece }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const pieceCount = Math.max(0, Math.ceil(totalMeters / 2));
  const ampValue = Number(amp);
  const profile = getOpphengRateRowForAmp(ampValue);
  const defaultRate = profile ? profile.rate : 0;
  const rate = sanitizeOpphengRate(ratePerPiece, defaultRate);
  const cost = round2(pieceCount * rate);
  return {
    cost,
    meters: totalMeters,
    pieceCount,
    ratePerPiece: rate,
    defaultRate,
    profile,
    isDefaultRate: rate === defaultRate,
    amp: Number.isFinite(ampValue) ? ampValue : null
  };
}

function getMontasjeProfileForAmp(amp, series){
  const a = Number(amp);
  if (!Number.isFinite(a) || a <= 0) return null;
  const table = series === EPOXY_IP68_SERIES ? EPOXY_MONTASJE_TIME_TABLE : MONTASJE_TIME_TABLE;
  for (const row of table){
    if (a <= row.maxAmp) return row;
  }
  return table[table.length - 1];
}

function calculateMontasje({ meter, angles, amp, series, hourlyRate }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const totalAngles = Math.max(0, Math.round(Number(angles) || 0));
  const rate = sanitizeHourlyRate(hourlyRate, DEFAULT_HOURLY_RATE);
  const ampValue = Number(amp);
  const profile = getMontasjeProfileForAmp(ampValue, series);
  if (!profile){
    return {
      cost: 0,
      meters: totalMeters,
      angles: totalAngles,
      hourlyRate: rate,
      totalHours: 0,
      profile: null,
      amp: Number.isFinite(ampValue) ? ampValue : null
    };
  }
  const hours = round2(totalMeters * profile.hoursPerMeter + totalAngles * profile.hoursPerAngle);
  const cost = round2(hours * rate);
  return {
    cost,
    meters: totalMeters,
    angles: totalAngles,
    hourlyRate: rate,
    totalHours: hours,
    profile,
    hoursPerMeter: profile.hoursPerMeter,
    hoursPerAngle: profile.hoursPerAngle,
    amp: Number.isFinite(ampValue) ? ampValue : null
  };
}

function calculateEngineering({ meter, angles, hourlyRate }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const totalAngles = Math.max(0, Math.round(Number(angles) || 0));
  const rate = sanitizeHourlyRate(hourlyRate, DEFAULT_ENGINEERING_HOURLY_RATE);
  const hoursFromMeters = totalMeters * ENGINEERING_HOURS_PER_METER;
  const hoursFromAngles = (totalAngles / 10) * ENGINEERING_HOURS_PER_TEN_ANGLES;
  const totalHours = round2(hoursFromMeters + hoursFromAngles);
  const cost = round2(totalHours * rate);
  return {
    cost,
    meters: totalMeters,
    angles: totalAngles,
    hourlyRate: rate,
    totalHours,
    hoursPerMeter: ENGINEERING_HOURS_PER_METER,
    hoursPerTenAngles: ENGINEERING_HOURS_PER_TEN_ANGLES
  };
}

function readMontasjeSettingsFromUI(){
  const hourlyRateInput = document.getElementById('montasjeHourlyRate');
  const hourlyRate = sanitizeHourlyRate(hourlyRateInput ? hourlyRateInput.value : DEFAULT_HOURLY_RATE, DEFAULT_HOURLY_RATE);
  if (hourlyRateInput){
    hourlyRateInput.value = String(hourlyRate);
  }
  const opphengRateInput = document.getElementById('opphengRate');
  let opphengRate = '';
  if (opphengRateInput){
    const defaultRate = toNum(opphengRateInput.dataset.defaultRate);
    const fallbackRate = Number.isFinite(defaultRate) ? defaultRate : 0;
    const raw = opphengRateInput.value;
    const sanitized = sanitizeOpphengRate(raw, fallbackRate);
    const isOverride = sanitized !== fallbackRate;
    opphengRateInput.dataset.appliedValue = String(sanitized);
    opphengRateInput.dataset.userOverride = isOverride ? 'true' : 'false';
    if (fallbackRate === 0 && !isOverride && raw.trim()===''){
      opphengRateInput.value = '';
    }else{
      opphengRateInput.value = (sanitized || sanitized === 0) ? String(sanitized) : '';
    }
    opphengRate = sanitized;
  }
  return { hourlyRate, opphengRate };
}

function readEngineeringSettingsFromUI(){
  const hourlyRateInput = document.getElementById('engineeringHourlyRate');
  const hourlyRate = sanitizeHourlyRate(
    hourlyRateInput ? hourlyRateInput.value : DEFAULT_ENGINEERING_HOURLY_RATE,
    DEFAULT_ENGINEERING_HOURLY_RATE
  );
  if (hourlyRateInput){
    hourlyRateInput.value = String(hourlyRate);
  }
  return { hourlyRate };
}

function formatMontasjeDetail(m){
  if (!m || !m.profile){
    return 'Montasje kalkuleres automatisk når ampere er valgt.';
  }
  const parts = [];
  parts.push(`${fmtIntNO.format(m.meters)} m × ${fmtNO.format(m.profile.hoursPerMeter)} t/m`);
  if (m.angles){
    parts.push(`${fmtIntNO.format(m.angles)} × ${fmtNO.format(m.profile.hoursPerAngle)} t`);
  }
  const basis = parts.join(' + ');
  const hoursTxt = fmtNO.format(m.totalHours);
  const rateTxt = fmtNO.format(m.hourlyRate);
  const ampTxt = Number.isFinite(m.amp) ? `${fmtIntNO.format(m.amp)}A` : '';
  const labelTxt = ampTxt ? `${ampTxt} (${m.profile.label})` : m.profile.label;
  return `Montasjegrunnlag (${labelTxt}): ${basis} = ${hoursTxt} t × ${rateTxt} kr/t`;
}

function formatOpphengDetail(o){
  if (!o){
    return 'Opphengsmateriell kalkuleres automatisk når ampere er valgt.';
  }
  if (!Number.isFinite(o.amp) || o.amp <= 0){
    return 'Opphengsmateriell kalkuleres automatisk når ampere er valgt.';
  }
  if (!o.meters){
    return 'Opphengsmateriell beregnes når meter er angitt.';
  }
  const metersTxt = fmtIntNO.format(o.meters);
  const piecesTxt = fmtIntNO.format(o.pieceCount || 0);
  const rateTxt = fmtNO.format(o.ratePerPiece);
  const costTxt = fmtNO.format(o.cost);
  const labelTxt = o.profile ? ` (${o.profile.label})` : '';
  return `Opphengsmateriell${labelTxt}: ${metersTxt} m gir ${piecesTxt} stk × ${rateTxt} kr/stk = ${costTxt} kr`;
}

function formatEngineeringDetail(e){
  if (!e){
    return 'Ingeniør beregnes automatisk fra meter og vinkler.';
  }
  const metersTxt = fmtIntNO.format(e.meters);
  const anglesTxt = fmtIntNO.format(e.angles);
  const perMeterTxt = fmtNO.format(e.hoursPerMeter);
  const perTenAnglesTxt = fmtNO.format(e.hoursPerTenAngles);
  const totalHoursTxt = fmtNO.format(e.totalHours);
  const rateTxt = fmtNO.format(e.hourlyRate);
  return `Ingeniørgrunnlag: ${metersTxt} m × ${perMeterTxt} t/m + (${anglesTxt} / 10) × ${perTenAnglesTxt} t = ${totalHoursTxt} t × ${rateTxt} kr/t`;
}

function setInputLocked(input, locked){
  if (!input) return;
  input.readOnly = locked;
  input.classList.toggle('locked', locked);
  input.setAttribute('aria-readonly', locked ? 'true' : 'false');
  input.dataset.locked = locked ? 'true' : 'false';
}

function updateMontasjePreview(){
  const meterEl = $('meter');
  const v90hEl = $('v90h');
  const v90vEl = $('v90v');
  const ampEl = $('ampSelect');
  const rateEl = $('montasjeHourlyRate');
  const opphengRateEl = $('opphengRate');
  const rateToggle = $('rateToggle');

  const meter = meterEl ? Number(meterEl.value || 0) : 0;
  const angles = (v90hEl ? Number(v90hEl.value || 0) : 0) + (v90vEl ? Number(v90vEl.value || 0) : 0);
  const amp = ampEl ? Number(ampEl.value || 0) : NaN;
  const series = $('series') ? $('series').value : '';
  const ratesUnlocked = rateToggle ? rateToggle.checked : true;
  const montasjeLocked = !ratesUnlocked;
  if (rateEl){
    setInputLocked(rateEl, montasjeLocked);
    if (montasjeLocked){
      rateEl.value = String(DEFAULT_HOURLY_RATE);
    }
  }
  const hourlyRate = rateEl ? rateEl.value : DEFAULT_HOURLY_RATE;

  const montasjePreview = calculateMontasje({ meter, angles, amp, series, hourlyRate });

  let opphengRateForCalc = 0;
  if (opphengRateEl){
    const opphengLocked = !ratesUnlocked;
    setInputLocked(opphengRateEl, opphengLocked);
    if (!opphengRateEl.dataset.userOverride){
      opphengRateEl.dataset.userOverride = 'false';
    }
    const hasAmp = Number.isFinite(amp) && amp > 0;
    const opphengProfile = getOpphengRateRowForAmp(amp);
    const defaultOpphengRate = hasAmp && opphengProfile ? opphengProfile.rate : 0;
    const ampKey = hasAmp ? String(amp) : '';
    const prevAmp = opphengRateEl.dataset.appliedAmp ?? '';
    const ampChanged = ampKey !== prevAmp;

    if (opphengLocked){
      opphengRateEl.dataset.userOverride = 'false';
    }

    if (ampChanged){
      opphengRateEl.dataset.appliedAmp = ampKey;
      if (opphengLocked || opphengRateEl.dataset.userOverride !== 'true'){
        if (hasAmp){
          opphengRateEl.value = defaultOpphengRate ? String(defaultOpphengRate) : '';
        }else{
          opphengRateEl.value = '';
        }
      }
    }

    if (opphengLocked){
      if (hasAmp){
        opphengRateEl.value = defaultOpphengRate ? String(defaultOpphengRate) : '';
      }else{
        opphengRateEl.value = '';
      }
    }

    opphengRateEl.placeholder = hasAmp ? (defaultOpphengRate ? String(defaultOpphengRate) : '') : '';

    const rawValue = opphengRateEl.value;
    const fallbackRate = hasAmp ? defaultOpphengRate : 0;
    const sanitizedRate = sanitizeOpphengRate(rawValue, fallbackRate);

    opphengRateEl.dataset.defaultRate = hasAmp ? String(defaultOpphengRate) : '';
    opphengRateEl.dataset.appliedValue = String(sanitizedRate);
    opphengRateEl.dataset.appliedAmp = ampKey;

    if (!hasAmp && opphengRateEl.dataset.userOverride !== 'true'){
      opphengRateEl.value = '';
    }
    if (!opphengLocked && ampChanged && opphengRateEl.dataset.userOverride !== 'true' && hasAmp){
      opphengRateEl.value = sanitizedRate || sanitizedRate === 0 ? String(sanitizedRate) : '';
    }

    opphengRateForCalc = sanitizedRate;
  }

  const opphengPreview = calculateOpphengsmateriell({ meter, amp, ratePerPiece: opphengRateForCalc });

  const labelEl = $('montasjeProfileLabel');
  const perMeterEl = $('montasjeHoursPerMeter');
  const perAngleEl = $('montasjeHoursPerAngle');
  const totalHoursEl = $('montasjeTotalHours');
  const costEl = $('montasjePreviewCost');
  const opphengRatePreviewEl = $('opphengPreviewRate');
  const opphengCostPreviewEl = $('opphengPreviewCost');
  const montasjeDetailEl = $('montasjeDetail');
  const opphengDetailEl = $('opphengDetail');

  const hasProfile = Boolean(montasjePreview.profile);

  if (labelEl){
    if (hasProfile){
      const ampTxt = Number.isFinite(montasjePreview.amp) ? `${fmtIntNO.format(montasjePreview.amp)}A` : '';
      labelEl.textContent = ampTxt ? `Strømskinne ${ampTxt} (${montasjePreview.profile.label})` : `Strømskinne ${montasjePreview.profile.label}`;
    }else{
      labelEl.textContent = 'Velg ampere for å hente montasjetider.';
    }
  }
  if (perMeterEl){
    perMeterEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.profile.hoursPerMeter)} t/m` : '–';
  }
  if (perAngleEl){
    perAngleEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.profile.hoursPerAngle)} t/vinkel` : '–';
  }
  if (totalHoursEl){
    totalHoursEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.totalHours)} t` : '–';
  }
  if (costEl){
    costEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.cost)} kr` : '–';
  }
  const hasOpphengAmp = Number.isFinite(opphengPreview.amp) && opphengPreview.amp > 0;
  if (opphengRatePreviewEl){
    opphengRatePreviewEl.textContent = hasOpphengAmp ? `${fmtNO.format(opphengPreview.ratePerPiece)} kr/stk` : '–';
  }
  if (opphengCostPreviewEl){
    opphengCostPreviewEl.textContent = hasOpphengAmp ? `${fmtNO.format(opphengPreview.cost)} kr` : '–';
  }
  if (montasjeDetailEl){
    montasjeDetailEl.textContent = formatMontasjeDetail(montasjePreview);
  }
  if (opphengDetailEl){
    opphengDetailEl.textContent = formatOpphengDetail(opphengPreview);
  }
  updateEngineeringPreview();
}

function updateEngineeringPreview(){
  const meterEl = $('meter');
  const v90hEl = $('v90h');
  const v90vEl = $('v90v');
  const rateEl = $('engineeringHourlyRate');
  const rateToggle = $('engineeringRateToggle');

  const meter = meterEl ? Number(meterEl.value || 0) : 0;
  const angles = (v90hEl ? Number(v90hEl.value || 0) : 0) + (v90vEl ? Number(v90vEl.value || 0) : 0);

  const ratesUnlocked = rateToggle ? rateToggle.checked : true;
  const locked = !ratesUnlocked;
  if (rateEl){
    setInputLocked(rateEl, locked);
    if (locked){
      rateEl.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
    }
  }
  const preview = calculateEngineering({
    meter,
    angles,
    hourlyRate: rateEl ? rateEl.value : DEFAULT_ENGINEERING_HOURLY_RATE
  });

  const profileEl = $('engineeringProfileLabel');
  const perMeterEl = $('engineeringHoursPerMeter');
  const perTenAnglesEl = $('engineeringHoursPerTenAngles');
  const totalHoursEl = $('engineeringTotalHours');
  const costEl = $('engineeringPreviewCost');
  const detailEl = $('engineeringDetail');

  if (profileEl){
    profileEl.textContent = 'Ingeniørtid gjelder alle serier og alle ampere.';
  }
  if (perMeterEl){
    perMeterEl.textContent = `${fmtNO.format(preview.hoursPerMeter)} t/m`;
  }
  if (perTenAnglesEl){
    perTenAnglesEl.textContent = `${fmtNO.format(preview.hoursPerTenAngles)} t / 10 vinkler`;
  }
  if (totalHoursEl){
    totalHoursEl.textContent = `${fmtNO.format(preview.totalHours)} t`;
  }
  if (costEl){
    costEl.textContent = `${fmtNO.format(preview.cost)} kr`;
  }
  if (detailEl){
    detailEl.textContent = formatEngineeringDetail(preview);
  }
}

// --- detect ---
function detectSeries(row){
  const code = String(pick(row,H.code)).toUpperCase();
  const d = String(pick(row,H.desc)).toUpperCase();
  if (Object.prototype.hasOwnProperty.call(row,'RCP IP68') || d.includes('RCP')) return EPOXY_IP68_SERIES;
  if (code.startsWith('XCM') || d.includes('XCM')) return 'XCM';
  if (code.startsWith('XCP') || d.includes('XCP')) return 'XCP-S';
  if (code.startsWith('XCA') || d.includes('XCA')) return 'XCP-S';
  return '';
}

function detectType(descRaw){
  const d0 = String(descRaw||'');
  const d  = d0.toLowerCase().replace(/\s+/g,' ');

  // feed
  if (d.includes('crt') && d.includes('feed')) return 'crt_board_feed';
  if ((d.includes('board-trans') || d.includes('board')) && d.includes('feed')) return 'board_feed';

  // XCM FEEDER (før alt generelt)
  if (/\bxcm\b/.test(d) && /\bfeeder\b/.test(d)){
    if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))        return 'xcm_feeder_3m';
    if (/l\s*=\s*1501\s*[-–]\s*2999\b/.test(d))        return 'xcm_feeder_1501_2999';
    if (/l\s*=\s*600\s*[-–]\s*1500\b/.test(d))         return 'xcm_feeder_600_1500';
  }

  // XCM DISTRIBUSJON
  if (/\bxcm\b/.test(d) && /straight\s*length/.test(d)){
    if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))        return 'straight_3m_dist';
    if (/l\s*=\s*1500\s*[-–]\s*2999\b/.test(d))        return 'xcm_dist_1500_2999';
    if (/l\s*=\s*1000\s*[-–]\s*1500\b/.test(d))        return 'xcm_dist_1000_1500';
  }

  // Generelt / XCP-S
  const isDistGen = /straight\s*length|\boutl(ets?)?\b/.test(d);
  if (/l\s*=\s*2001\s*[-–]\s*3000\b/.test(d))        return isDistGen ? 'straight_3m_dist' : 'straight_3m';
  if (/l\s*=\s*1001\s*[-–]\s*2000\b/.test(d))        return isDistGen ? 'straight_1501_2000_dist' : 'straight_1501_2000';
  if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))          return isDistGen ? 'straight_3m_dist' : 'straight_3m';
  if (/l\s*=\s*1501\s*[-–]\s*2000\b/.test(d))          return isDistGen ? 'straight_1501_2000_dist' : 'straight_1501_2000';
  if (/l\s*=\s*500\s*[-–]\s*1000\b/.test(d))           return isDistGen ? 'straight_500_1000_dist'  : 'straight_500_1000';

  // andre typer
  if (d.includes('horizontal') && (d.includes('elbow') || d.includes('90'))) return 'elbow_horizontal_90';
  if (d.includes('vertical')   && (d.includes('elbow') || d.includes('90'))) return 'elbow_vertical_90';
  if (/tap[\s-]*off\s*box/.test(d))  return 'tap_off_box';
  if (/plug[\s-]*in\s*box/.test(d))  return 'plug_in_box';
  if (/bolt[\s-]*on\s*box|b160\s*bolt/.test(d)) return 'bolt_on_box';
  if (/monobloc/.test(d) && /joint/.test(d)) return 'monobloc_joint_4c';
  if (/junction\s*kit\s*part\s*1/.test(d)) return 'junction_kit_part_1';
  if (/junction\s*kit\s*part\s*2/.test(d)) return 'junction_kit_part_2';
  if (/edge[^a-z0-9]*molds[^a-z0-9]*kit/.test(d)) return 'edge_molds_kit';
  if (/vert[^a-z0-9]*molds[^a-z0-9]*kit/.test(d)) return 'vert_molds_kit';
  if (/end\s*cover/.test(d))        return 'end_cover';
  if (/end\s*feed\s*unit/.test(d))  return 'end_feed_unit';
  if (/expansion/.test(d))          return 'expansion_unit';

  const isFire = /fire[\s-]*barrier|firebarrier|fire[\s-]*stop|firestop|brannbarrier|brannbarriere|brannelement|brann/.test(d);
  if (isFire){
    const ext = /(external|utvendig|ytter)/i.test(d);
    const int = /(internal|innvendig|inner)/i.test(d);
    if (ext && !int) return 'fire_barrier_kit_external';
    if (int && !ext) return 'fire_barrier_kit_internal';
    return 'fire_barrier_kit';
  }
  return '';
}

const XAP_TYPE_MAP = Object.freeze({
  'Straight length, 3 meters max.': 'straight_3m',
  'Straight length, 1 meter': 'straight_500_1000',
  'L horizontal elbow': 'elbow_horizontal_90',
  'L horizontal elbow, special angle': 'elbow_horizontal_90',
  'L vertical elbow': 'elbow_vertical_90',
  'L vertical elbow, special angle': 'elbow_vertical_90',
  'L veritcal elbow, special angle': 'elbow_vertical_90',
  'Expansion Joint': 'expansion_unit',
  'Switchboard/panel flange end': 'board_feed',
  'Transformer flange end': 'crt_board_feed',
  'End cap': 'end_cover',
  'End tap box - EMPTY': 'end_feed_unit',
  'End tap box - with 400A MCCB, 3P, 36kA': 'end_feed_unit',
  'End tap box - with 400A MCCB, 4P, 36kA': 'end_feed_unit',
  'End tap box - with 500A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 500A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 630A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 630A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 800A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 800A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 1000A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 1000A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 1250A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 1250A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 1600A MCCB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 1600A MCCB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 2000A ACB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 2000A ACB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 2500A ACB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 2500A ACB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 3200A ACB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 3200A ACB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 4000A ACB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 4000A ACB, 4P, 50kA': 'end_feed_unit',
  'End tap box - with 5000A ACB, 3P, 50kA': 'end_feed_unit',
  'End tap box - with 5000A ACB, 4P, 50kA': 'end_feed_unit',
  'Outlet for tap-off boxes': 'outlet_section',
  'Joint': 'joint',
  'Flexible set': 'flexible_set',
  'Tee horizontal offset': 'tee_horizontal_offset',
  'Tee vertical offset': 'tee_vertical_offset',
  'Z horizontal offset': 'z_horizontal_offset',
  'Z vertical offset': 'z_vertical_offset'
});

function isXAPRow(row){
  return Object.prototype.hasOwnProperty.call(row,'Beskrivelse')
      && Object.prototype.hasOwnProperty.call(row,'Pris_USD');
}

function parseXAPPriceUSD(raw){
  if (raw===undefined || raw===null) return 0;
  const txt = String(raw).trim();
  if (!txt || txt==='-' || txt==='$-') return 0;
  const cleaned = txt.replace(/\$/g,'').replace(/\s/g,'').replace(',','.');
  const num = Number(cleaned);
  return Number.isFinite(num) ? round2(num) : 0;
}

function parseXAPAmp(row, desc){
  const raw = row?.Ampere;
  if (raw){
    const cleaned = String(raw).replace(/[^\d.,]/g,'').replace(',', '.');
    const num = Number(cleaned);
    if (Number.isFinite(num)) return num;
  }
  const match = String(desc||'').match(/(\d{2,4})\s*A/i);
  if (match) return Number(match[1]);
  return NaN;
}

function normalizeXAPMaterial(value){
  if (!value && value !== 0) return '';
  const txt = String(value).trim();
  if (!txt) return '';
  if (/^al/i.test(txt)) return 'Al';
  if (/^cu/i.test(txt)) return 'Cu';
  return txt;
}

function extractEpoxyAmpFromDesc(desc){
  const match = String(desc || '').match(/(\d{3,4})\s*A\b/i);
  if (!match) return NaN;
  const amp = Number(match[1]);
  return Number.isFinite(amp) ? amp : NaN;
}

function parseEpoxyMonoblocAmp(desc){
  const d = String(desc || '').toUpperCase().replace(/\s+/g,'');
  if (d.includes('3XB160')) return 5000;
  if (d.includes('2XB190')) return 4000;
  if (d.includes('2XB160')) return 3200;
  if (d.includes('2XB120')) return 2500;
  if (d.includes('B210')) return 2000;
  if (d.includes('B160')) return 1600;
  if (d.includes('B120')) return 1250;
  return NaN;
}

function adaptXAPRow(row){
  const desc = String(row?.Beskrivelse || '').trim();
  if (!desc) return [];
  const parts = desc.split(' - ');
  const typeLabelRaw = parts.length > 1 ? parts.slice(1).join(' - ').trim() : '';
  const normalizedLabel = typeLabelRaw.replace(/\s+/g,' ').trim();
  const seriesPart = desc.split(/\s+/)[0] || XAP_SERIES;
  const series = seriesPart.replace(/[^A-Za-z0-9-]/g,'') || XAP_SERIES;
  const type = XAP_TYPE_MAP[normalizedLabel] || '';
  const ampere = parseXAPAmp(row, desc);
  const unit_price_usd = parseXAPPriceUSD(row?.Pris_USD);
  const unit_price = convertUsdToNok(unit_price_usd);
  const base = {
    code: desc,
    type,
    series,
    ampere,
    unit_price,
    unit_price_usd,
    _desc: desc,
    ledere: '3F+N+PE',
    ledermateriell: normalizeXAPMaterial(row?.Leder || '')
  };
  const entries = [base];
  if (type === 'straight_3m'){
    entries.push({ ...base, type: 'straight_3m_dist' });
  } else if (type === 'straight_500_1000'){
    entries.push({ ...base, type: 'straight_500_1000_dist' });
  }
  return entries;
}

// amp
function extractAmpGeneric(row){
  let src = (pick(row, H.desc) + ' ' + pick(row, H.code)).toUpperCase();
  let m = src.match(/(\d{2,4})\s*A\b/);
  if (m) return Number(m[1]);
  m = src.match(/\b(\d{2,4})\b/);
  if (m) return Number(m[1]);
  return NaN;
}
function deriveAmp(row){
  if (Number.isFinite(row.ampere)) return Number(row.ampere);
  const s = String((row._desc||'')+' '+(row.code||'')).toUpperCase();
  let m = s.match(/(\d{2,4})\s*A\b/); if (m) return Number(m[1]);
  m = s.match(/\b(\d{2,4})\b/);       if (m) return Number(m[1]);
  return NaN;
}

// rå → katalog
function adaptRawToCatalog(rawRows){
  const list = [];
  let currentEpoxyAmp = NaN;
  rawRows.forEach(r=>{
    if (isXAPRow(r)){
      const adapted = adaptXAPRow(r);
      adapted.filter(Boolean).forEach(item=>list.push(item));
      return;
    }
    const code = pick(r, H.code);
    const price = toNum(pick(r, H.price));
    const desc  = pick(r, H.desc);
    const series= detectSeries(r);
    let type    = detectType(desc);

    // brann tag
    const et2H  = (r._cols && r._cols.length>=8) ? r._cols[7] : '';
    const tag = String(et2H||'').replace(/\s+/g,'').toUpperCase();
    const tagAmp =
      tag==='B160'   ? 1250 :
      tag==='B190'   ? 1600 :
      tag==='B210'   ? 2000 :
      tag==='2XB160' ? 2500 :
      tag==='2XB190' ? 3200 :
      tag==='2XB210' ? 4000 :
      tag==='3XB160' ? 5000 : NaN;

    if (Number.isFinite(tagAmp)) {
      const d = String(desc||'').toLowerCase();
      const ext = /(external|utvendig|ytter)/.test(d);
      const int = /(internal|innvendig|inner)/.test(d);
      type = ext && !int ? 'fire_barrier_kit_external'
           : int && !ext ? 'fire_barrier_kit_internal'
           : 'fire_barrier_kit';
    } else if (!type && /fire|brann/i.test(String(desc||''))) {
      type = 'fire_barrier_kit';
    }

    let amp = Number.isFinite(tagAmp) ? tagAmp : extractAmpGeneric(r);
    if (series === EPOXY_IP68_SERIES){
      const explicitAmp = extractEpoxyAmpFromDesc(desc);
      if (Number.isFinite(explicitAmp) && EPOXY_MAIN_ELEMENT_TYPES.has(type)){
        currentEpoxyAmp = explicitAmp;
      }
      if (type === 'monobloc_joint_4c'){
        const monoblocAmp = parseEpoxyMonoblocAmp(desc);
        amp = Number.isFinite(currentEpoxyAmp) ? currentEpoxyAmp : (Number.isFinite(monoblocAmp) ? monoblocAmp : amp);
      } else if (
        type === 'junction_kit_part_1'
        || type === 'junction_kit_part_2'
        || type === 'edge_molds_kit'
        || type === 'vert_molds_kit'
        || type === 'fire_barrier_kit'
        || type === 'fire_barrier_kit_external'
        || type === 'fire_barrier_kit_internal'
      ){
        if (Number.isFinite(currentEpoxyAmp)) amp = currentEpoxyAmp;
        else if (Number.isFinite(explicitAmp)) amp = explicitAmp;
      } else if (Number.isFinite(explicitAmp)){
        amp = explicitAmp;
      }
    }

    list.push({ code, type, series, ampere: amp, unit_price: price, _desc: desc, _et2: et2H });
  });
  return list;
}

// match
function matchesLedere(row, ledere){
  if (!ledere) return true;
  if (row.ledere) return row.ledere === ledere;
  if (row.series === 'XCM') return true;
  if (row.series === EPOXY_IP68_SERIES) return true;
  const c = String(row.code||'').toUpperCase();
  const implied = c.includes('-3W') ? '3F+PE' : '3F+N+PE';
  return ledere === implied || !c;
}
function isThreeWireRow(row){
  return String(row?.code || '').toUpperCase().includes('-3W');
}
function hasConductorVariants(rows){
  return rows.some(isThreeWireRow);
}
function byTypeAmpSeries(rows, type, amp, series){ return rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp)); }
function byTypeAmpSeriesL(rows, type, amp, series, ledere){
  const exactAmp = rows.filter(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp));
  const exactMatch = exactAmp.find(r=>matchesLedere(r, ledere));
  if (exactMatch) return exactMatch;
  if (hasConductorVariants(exactAmp)) return null;
  if (exactAmp.length) return exactAmp[0];

  const sameTypeSeries = rows.filter(r=>r.type===type && r.series===series);
  const looseMatch = sameTypeSeries.find(r=>matchesLedere(r, ledere));
  if (looseMatch) return looseMatch;
  if (hasConductorVariants(sameTypeSeries)) return null;
  return sameTypeSeries[0] || null;
}
function byBoxAll(catalog, kind, amp, prefSeries){
  const list = catalog.filter(r=>r.type===kind && (amp ? Number(deriveAmp(r))===Number(amp) : true));
  if (!list.length) return null;
  const inSeries = list.find(r=>r.series===prefSeries);
  return inSeries || list[0];
}
function findByTypeSeriesAmp(rows, type, series, amp){
  let r = rows.find(x=>x.type===type && x.series===series && Number(deriveAmp(x))===Number(amp));
  if (r) return r;
  const cand = rows.filter(x=>x.type===type && x.series===series);
  r = cand.find(x=>Number(deriveAmp(x))===Number(amp));
  return r || cand[0] || null;
}
function getFireBarrier(rows, amp, series){
  if (series === EPOXY_IP68_SERIES && Number(amp) === 5000){
    const fb2500 = getFireBarrier(rows, 2500, series);
    if (fb2500 && Number.isFinite(fb2500.unit)){
      return { code: `${fb2500.code} x2`, unit: round2(fb2500.unit * 2) };
    }
  }
  const direct = rows.find(r=>r.type==='fire_barrier_kit' && r.series===series && Number(deriveAmp(r))===amp);
  if (direct) return { code: direct.code, unit: toNum(direct.unit_price) };
  const ext = rows.find(r=>r.type==='fire_barrier_kit_external' && r.series===series && Number(deriveAmp(r))===amp);
  const int = rows.find(r=>r.type==='fire_barrier_kit_internal' && r.series===series && Number(deriveAmp(r))===amp);
  if (!ext && !int) return null;
  const unit = (ext?toNum(ext.unit_price):0) + (int?toNum(int.unit_price):0);
  const code = [ext?.code, int?.code].filter(Boolean).join('+') || 'FIRE-BARRIER';
  return { code, unit };
}

function getUnitPriceFromRow(row){
  if (!row) return NaN;
  const unit = toNum(row.unit_price);
  return Number.isFinite(unit) && unit > 0 ? unit : NaN;
}

function getUnitPriceForType(catRows, series, amp, ledere, typeSpec){
  const row = findLengthVariant(catRows, series, amp, ledere, typeSpec)
    || byTypeAmpSeriesL(catRows, Array.isArray(typeSpec) ? typeSpec[0] : typeSpec, amp, series, ledere);
  return getUnitPriceFromRow(row);
}

function buildRawUnitPriceSnapshot(cat, input){
  const startType = String(input.startEl || '').trim();
  const endType = String(input.sluttEl || '').trim();
  const meterUnit = getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, [
    'straight_500_1000',
    'straight_500_1000_dist',
    'xcm_feeder_600_1500',
    'xcm_dist_1000_1500'
  ]);
  const vinkelVertikalUnit = getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, 'elbow_vertical_90');
  const vinkelHorisontalUnit = getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, 'elbow_horizontal_90');
  const startUnit = startType && startType !== 'none'
    ? getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, startType)
    : NaN;
  const endUnit = endType && endType !== 'none'
    ? getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, endType)
    : NaN;
  const fireBarrier = getFireBarrier(cat.rows, input.ampere, input.series);
  const expansionUnit = getUnitPriceForType(cat.rows, input.series, input.ampere, input.ledere, 'expansion_unit');
  const firstBoxItem = normalizeBoxItems(input.boxItems, input.boxSel, input.boxQty, input.boxInnmatSum)[0];
  let tapOffUnit = NaN;
  if (firstBoxItem?.boxSel){
    const [kind, ampStr] = String(firstBoxItem.boxSel || '').split('|');
    const row = byBoxAll(cat.catalog, kind, ampStr ? Number(ampStr) : undefined, input.series);
    tapOffUnit = getUnitPriceFromRow(row);
  }
  return {
    meter: Number.isFinite(meterUnit) ? round2(meterUnit) : null,
    vinkel: Number.isFinite(vinkelVertikalUnit) ? round2(vinkelVertikalUnit) : null,
    vinkelVertikal: Number.isFinite(vinkelVertikalUnit) ? round2(vinkelVertikalUnit) : null,
    vinkelHorisontal: Number.isFinite(vinkelHorisontalUnit) ? round2(vinkelHorisontalUnit) : null,
    tavleelement: Number.isFinite(startUnit) ? round2(startUnit) : null,
    sluttelement: Number.isFinite(endUnit) ? round2(endUnit) : null,
    brann: fireBarrier && Number.isFinite(fireBarrier.unit) ? round2(fireBarrier.unit) : null,
    ekspansjon: Number.isFinite(expansionUnit) ? round2(expansionUnit) : null,
    avtappingsboks: Number.isFinite(tapOffUnit) ? round2(tapOffUnit) : null
  };
}

// plan
function planSegments(m){
  let rem = Math.ceil(Number(m));
  const plan = { n3:0, n15_2000:0, n500_1000:0 };
  while(rem > 3){ plan.n3++; rem -= 3; }
  if (rem > 2){ plan.n3++; }
  else if (rem > 1){ plan.n15_2000++; }
  else if (rem > 0){ plan.n500_1000++; }
  return plan;
}

// lines
function makeLine(row, series, amp, ledere, qty){
  const unit = toNum(row.unit_price);
  if (!Number.isFinite(unit)) throw new Error('Ugyldig pris: '+row.code);
  return { code: row.code, type: row.type, series, ampere: Number(deriveAmp(row))||amp||'', ledere, antall: qty, enhet: unit, sum: round2(unit*qty) };
}
function makeCustomLine(code, type, series, amp, ledere, unit, qty){
  if (!Number.isFinite(unit)) throw new Error('Ugyldig pris: '+type);
  return { code, type, series, ampere: amp||'', ledere, antall: qty, enhet: unit, sum: round2(unit*qty) };
}
const preferTypes = (...types)=>{
  const seen = new Set();
  const list = [];
  types.forEach(t=>{
    if (!t || seen.has(t)) return;
    seen.add(t);
    list.push(t);
  });
  return list;
};
const toTypeArray = spec => Array.isArray(spec) ? spec.filter(Boolean) : (spec ? [spec] : []);
const primaryTypeLabel = spec => toTypeArray(spec)[0] || 'lengde';

// Hvilke type-navn tilsvarer 3m / ~2m / ~1m for valgt serie og distribusjon
function lengthTypes(series, dist){
  if (series==='XCM'){
    return dist
      ? {
          L3: preferTypes('straight_3m_dist','xcm_feeder_3m'),
          L2: preferTypes('xcm_dist_1500_2999','xcm_feeder_1501_2999'),
          L1: preferTypes('xcm_dist_1000_1500','xcm_feeder_600_1500')
        }
      : {
          L3: preferTypes('xcm_feeder_3m','straight_3m_dist'),
          L2: preferTypes('xcm_feeder_1501_2999','xcm_dist_1500_2999'),
          L1: preferTypes('xcm_feeder_600_1500','xcm_dist_1000_1500')
        };
  }
  // XCP-S
  return dist
    ? {
        L3: preferTypes('straight_3m_dist','straight_3m'),
        L2: preferTypes('straight_1501_2000_dist','straight_1501_2000'),
        L1: preferTypes('straight_500_1000_dist','straight_500_1000')
      }
    : {
        L3: preferTypes('straight_3m','straight_3m_dist'),
        L2: preferTypes('straight_1501_2000','straight_1501_2000_dist'),
        L1: preferTypes('straight_500_1000','straight_500_1000_dist')
      };
}

// Greedy plan (3,2,1) + fallback-regler
function planWithFallback(m, avail){
  // grunnplan
  let rem = Math.ceil(Number(m));
  let n3=0,n2=0,n1=0;
  while(rem>3){ n3++; rem-=3; }
  if (rem>2) n3++; else if (rem>1) n2++; else if (rem>0) n1++;

  // 1) Mangler 1m: bytt (1×3m + 1×1m) → (2×2m)
  if (!avail.L1 && avail.L2){
    while(n1>0 && n3>0){ n3--; n2+=2; n1--; }
  }
  // 2) Mangler 2m: bytt 1×2m → (1×3m + -1×1m)
  if (!avail.L2 && avail.L3 && avail.L1){
    while(n2>0 && n1>0){ n2--; n3++; n1--; }
  }
  // 3) Mangler 3m: bytt 1×3m → (1×2m + 1×1m) el. (3×1m) el. (2×2m - 1×1m)
  if (!avail.L3){
    if (avail.L2 && avail.L1){ while(n3>0){ n3--; n2++; n1++; } }
    else if (avail.L1){ while(n3>0){ n3--; n1+=3; } }
    else if (avail.L2){ while(n3>0 && n1>0){ n3--; n2+=2; n1--; } }
  }

  // 4) Hvis 2m fortsatt mangler: prøv 1×2m → 2×1m
  if (!avail.L2 && avail.L1){ while(n2>0){ n2--; n1+=2; } }

  // 5) Hvis 1m fortsatt mangler og finnes 2m: prøv å balansere med 3m
  if (!avail.L1 && avail.L2 && n1>0){
    // ikke mulig uten å låne 3m; forsøk en ekstra bytte hvis mulig
    if (n3>0){ n3--; n2+=2; n1--; }
  }

  // valider lengde
  const total = 3*n3 + 2*n2 + 1*n1;
  if (total !== Math.ceil(Number(m))) {
    throw new Error('Kan ikke finne lengdekombinasjon med tilgjengelige elementer.');
  }
  return {n3,n2,n1};
}

// Finn ut om type finnes for valgt konfig
function findLengthVariant(catRows, series, amp, ledere, typeSpec){
  const rows = catRows.filter(r=>r.series===series);
  for (const type of toTypeArray(typeSpec)){
    const row = byTypeAmpSeriesL(rows,type,amp,series,ledere);
    if (row) return row;
  }
  return null;
}

function hasType(catRows, series, amp, ledere, typeSpec){
  return Boolean(findLengthVariant(catRows, series, amp, ledere, typeSpec));
}

// pris
// streng match på amp, men fall tilbake til type+serie hvis amp mangler
function needAnyAmp(rows, type, amp, series){
  return byTypeAmpSeries(rows, type, amp, series)
      || rows.find(r => r.type===type && r.series===series); // ingen amp i CSV
}
function needAnyAmpL(rows, type, amp, series, ledere){
  const sameTypeSeries = rows.filter(r=>r.type===type && r.series===series);
  return byTypeAmpSeriesL(rows, type, amp, series, ledere)
      || sameTypeSeries.find(r => matchesLedere(r, ledere))
      || (!hasConductorVariants(sameTypeSeries) ? sameTypeSeries[0] : null);
}

function epoxyMonoblocFactor(amp){
  const factor = EPOXY_MONOBLOC_FACTOR_BY_AMP[Number(amp)];
  return Number.isFinite(factor) ? factor : NaN;
}

function findEpoxyRow(rows, type, amp, series){
  return byTypeAmpSeries(rows, type, amp, series)
      || rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp))
      || rows.find(r=>r.type===type && r.series===series)
      || null;
}

function addEpoxyIp68AutoBomLines(cat, input, bom, pf){
  if (input.series !== EPOXY_IP68_SERIES) return;
  const straightCount = Number(pf?.n3 || 0) + Number(pf?.n2 || 0) + Number(pf?.n1 || 0);
  const horizontalAngles = Math.max(0, Math.round(Number(input.v90_h) || 0));
  const verticalAngles = Math.max(0, Math.round(Number(input.v90_v) || 0));
  const monoblocCount = straightCount + horizontalAngles + verticalAngles;
  if (monoblocCount <= 0) return;

  const push = (row, qty)=>bom.push(makeLine(row, input.series, input.ampere, input.ledere, qty));

  const monoblocRow = findEpoxyRow(cat.rows, 'monobloc_joint_4c', input.ampere, input.series);
  if (!monoblocRow) throw new Error('Mangler 4C monobloc joint for valgt ampere.');
  push(monoblocRow, monoblocCount);

  const edgeMoldsCount = straightCount + horizontalAngles;
  if (edgeMoldsCount > 0){
    const edgeRow = findEpoxyRow(cat.rows, 'edge_molds_kit', input.ampere, input.series);
    if (!edgeRow) throw new Error('Mangler EDGE MOLDS KIT for valgt ampere.');
    push(edgeRow, edgeMoldsCount);
  }

  if (verticalAngles > 0){
    const vertRow = findEpoxyRow(cat.rows, 'vert_molds_kit', input.ampere, input.series);
    if (!vertRow) throw new Error('Mangler VERT MOLDS KIT for valgt ampere.');
    push(vertRow, verticalAngles);
  }

  const factor = epoxyMonoblocFactor(input.ampere);
  if (!Number.isFinite(factor)) throw new Error('Mangler støpemassefaktor for valgt ampere.');
  const junctionQty = Math.ceil(monoblocCount * factor);
  if (junctionQty <= 0) return;

  const junctionPart1 = findEpoxyRow(cat.rows, 'junction_kit_part_1', input.ampere, input.series);
  if (!junctionPart1) throw new Error('Mangler Junction KIT part 1 for valgt ampere.');
  push(junctionPart1, junctionQty);

  const junctionPart2 = findEpoxyRow(cat.rows, 'junction_kit_part_2', input.ampere, input.series);
  if (!junctionPart2) throw new Error('Mangler Junction KIT part 2 for valgt ampere.');
  push(junctionPart2, junctionQty);
}

function price(cat, input){
  const bom=[];
  const push = (r,q)=>bom.push(makeLine(r, input.series, input.ampere, input.ledere, q));
  const need = (type)=> byTypeAmpSeries(cat.rows, type, input.ampere, input.series);

  // Startelement
if (input.startEl === 'board_feed'){
  const bf = needAnyAmpL(cat.rows, 'board_feed', input.ampere, input.series, input.ledere);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.startEl === 'end_feed_unit'){
  const ef = needAnyAmpL(cat.rows, 'end_feed_unit', input.ampere, input.series, input.ledere);
  if (!ef) throw new Error(`Mangler end_feed_unit for ${input.series}.`);
  push(ef,1);
}

// Sluttelement
if (input.sluttEl === 'board_feed'){
  const bf = needAnyAmpL(cat.rows, 'board_feed', input.ampere, input.series, input.ledere);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.sluttEl === 'crt_board_feed'){
  if (!seriesSupportsCrtFeed(input.series)) throw new Error('Trafoelement er ikke tilgjengelig for valgt system.');
  const crt = needAnyAmpL(cat.rows, 'crt_board_feed', input.ampere, input.series, input.ledere);
  if (!crt) throw new Error(`Mangler crt_board_feed for ${input.series}.`);
  push(crt,1);
} else if (input.sluttEl === 'end_cover'){
  const ec = needAnyAmpL(cat.rows, 'end_cover', input.ampere, input.series, input.ledere);
  if (!ec) throw new Error(`Mangler end_cover for ${input.series}.`);
  push(ec,1);
}

  // ⬇ Lengder med fallback
const tmap = lengthTypes(input.series, input.dist);
const avail = {
  L3: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L3),
  L2: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L2),
  L1: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L1)
};
const pf = planWithFallback(input.meter, avail);

// legg til linjer etter plan
if (pf.n3){
  const r = findLengthVariant(cat.rows, input.series, input.ampere, input.ledere, tmap.L3);
  if(!r) throw new Error(`Mangler ${primaryTypeLabel(tmap.L3)}.`);
  push(r, pf.n3);
}
if (pf.n2){
  const r = findLengthVariant(cat.rows, input.series, input.ampere, input.ledere, tmap.L2);
  if(!r) throw new Error(`Mangler ${primaryTypeLabel(tmap.L2)}.`);
  push(r, pf.n2);
}
if (pf.n1){
  const r = findLengthVariant(cat.rows, input.series, input.ampere, input.ledere, tmap.L1);
  if(!r) throw new Error(`Mangler ${primaryTypeLabel(tmap.L1)}.`);
  push(r, pf.n1);
}

  // Vinkler
  if (input.v90_h){ const r=byTypeAmpSeriesL(cat.rows,'elbow_horizontal_90',input.ampere,input.series,input.ledere); if(!r) throw new Error('Mangler elbow_horizontal_90.'); push(r,input.v90_h); }
  if (input.v90_v){ const r=byTypeAmpSeriesL(cat.rows,'elbow_vertical_90'  ,input.ampere,input.series,input.ledere); if(!r) throw new Error('Mangler elbow_vertical_90.');   push(r,input.v90_v); }
  addEpoxyIp68AutoBomLines(cat, input, bom, pf);

  // Avtappingsbokser
  const boxItems = normalizeBoxItems(input.boxItems, input.boxSel, input.boxQty, input.boxInnmatSum);
  if (boxItems.length){
    boxItems.forEach(item=>{
      const [kind, ampStr] = String(item.boxSel || '').split('|');
      if (kind==='bolt_on_box' && input.series==='XCM') throw new Error('Bolt-on box kan ikke brukes på XCM.');
      const row = byBoxAll(cat.catalog, kind, ampStr ? Number(ampStr) : undefined, input.series);
      if (!row) throw new Error(`Mangler ${kind} ${ampStr||''}A.`);
      const line = makeLine(row, input.series, deriveAmp(row)||'', input.ledere, item.boxQty);
      const baseEnhet = Number(line.enhet || 0);
      const innmatSum = Number(item.innmatSum || 0);
      const tapOffGroupId = String(item.id || '').trim() || `tapoff-${bom.length + 1}-${String(kind || 'box')}-${String(ampStr || '')}`;
      line.tapOffBoxSel = item.boxSel;
      line.tapOffInnmatPerUnit = Number.isFinite(innmatSum) ? Math.max(0, innmatSum) : 0;
      line.tapOffInnmatTotal = round2(line.tapOffInnmatPerUnit * Number(item.boxQty || 0));
      line.tapOffBaseEnhet = Number.isFinite(baseEnhet) ? baseEnhet : 0;
      line.tapOffGroupId = tapOffGroupId;
      line.tapOffIncludesInnmatInSum = false;
      bom.push(line);
      if (line.tapOffInnmatPerUnit > 0 && line.tapOffInnmatTotal > 0){
        bom.push({
          code: `INNMAT-${line.code}`,
          type: `${kind}_innmat`,
          series: input.series,
          ampere: line.ampere,
          ledere: input.ledere,
          antall: item.boxQty,
          enhet: line.tapOffInnmatPerUnit,
          sum: line.tapOffInnmatTotal,
          tapOffGroupId,
          tapOffParentCode: line.code,
          tapOffBoxSel: item.boxSel,
          tapOffInnmatLine: true
        });
      }
    });
  }

  // Spesialelementer prises manuelt per element.
  const specialElementItems = normalizeSpecialElementItems(input.specialElementItems);
  specialElementItems.forEach(item=>{
    const selection = String(item.selection || '').trim();
    const label = specialElementLabelFromSelection(selection);
    const code = `SPECIAL-${selection.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toUpperCase()}`;
    const line = makeCustomLine(code, label, input.series, input.ampere, input.ledere, item.unitSum, item.qty);
    line.specialElementGroupId = item.id;
    line.specialElementSelection = selection;
    bom.push(line);
  });

  // Brann
  if (input.fbQty>0){
    const fb = getFireBarrier(cat.rows, input.ampere, input.series);
    if (!fb) throw new Error('Mangler fire barrier kit.');
    bom.push(makeCustomLine(fb.code, 'fire_barrier_kit', input.series, input.ampere, input.ledere, fb.unit, input.fbQty));
  }

  // Ekspansjon
  if (input.meter > 30 && input.expansionYes){
    const exp = byTypeAmpSeriesL(cat.rows,'expansion_unit',input.ampere,input.series,input.ledere);
    if (!exp) throw new Error(`Mangler expansion_unit for ${input.series} ${input.ampere}.`);
    push(exp,1);
  }

  const tapOffBoxTotal = sumSeparateTapOffBoxTotal(bom);
  const specialElementTotal = round2(bom.reduce((sum, entry)=>{
    if (!isSeparateSpecialElementBomLine(entry)) return sum;
    return sum + resolveBomLineSum(entry);
  }, 0));
  const material = round2(bom.reduce((sum, entry)=>{
    if (isSeparateTapOffBoxBomLine(entry) || isSeparateSpecialElementBomLine(entry)) return sum;
    return sum + resolveBomLineSum(entry);
  }, 0));
  const marginRate = normalizeMarginRate(input.marginRate, DEFAULT_MATERIAL_MARGIN_RATE);
  const rate     = Number(input.freightRate ?? 0.10);
  const montasje = calculateMontasje({
    meter: input.meter,
    angles: (input.v90_h || 0) + (input.v90_v || 0),
    amp: input.ampere,
    series: input.series,
    hourlyRate: input.montasjeSettings?.hourlyRate
  });
  const engineering = calculateEngineering({
    meter: input.meter,
    angles: (input.v90_h || 0) + (input.v90_v || 0),
    hourlyRate: input.engineeringSettings?.hourlyRate
  });
  const oppheng = calculateOpphengsmateriell({
    meter: input.meter,
    amp: input.ampere,
    ratePerPiece: input.montasjeSettings?.opphengRate
  });
  const totals = calculateTotalsFromMaterial({
    material,
    marginRate,
    freightRate: rate,
    montasjeCost: montasje.cost,
    montasjeMarginRate: input.montasjeMarginRate,
    engineeringCost: engineering.cost,
    engineeringMarginRate: input.engineeringMarginRate,
    opphengCost: oppheng.cost,
    opphengMarginRate: input.opphengMarginRate
  });
  const rawUnitPrices = buildRawUnitPriceSnapshot(cat, input);
  return {
    bom,
    material,
    rawUnitPrices,
    tapOffBoxTotal,
    specialElementTotal,
    marginRate: totals.marginRate,
    marginFactor: totals.marginFactor,
    margin: totals.margin,
    subtotal: totals.subtotal,
    freight: totals.freight,
    montasjeMarginRate: totals.montasjeMarginRate,
    montasjeMargin: totals.montasjeMargin,
    montasje,
    engineeringMarginRate: totals.engineeringMarginRate,
    engineeringMargin: totals.engineeringMargin,
    engineering,
    opphengMarginRate: totals.opphengMarginRate,
    opphengMargin: totals.opphengMargin,
    oppheng,
    totalExMontasje: totals.totalExMontasje,
    totalInclMontasje: totals.totalInclMontasje,
    totalInclEngineering: totals.totalInclEngineering,
    totalInclOppheng: totals.totalInclOppheng,
    total: totals.total
  };
}

function renderBomTable(tableId, bomList){
  const tbody = document.querySelector(`#${tableId} tbody`);
  if (!tbody) return;
  tbody.innerHTML = '';
  (bomList || []).forEach(b=>{
    const unitVal = Number(b.enhet);
    const sumVal = Number(b.sum);
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${b.code}</td><td>${b.type}</td><td>${b.series}</td><td>${b.ampere}</td><td>${b.lederes||b.ledere||''}</td><td>${b.antall}</td><td>${Number.isFinite(unitVal)?unitVal.toFixed(2):''}</td><td>${Number.isFinite(sumVal)?sumVal.toFixed(2):''}</td>`;
    tbody.appendChild(tr);
  });
}

function computeXapComparison(baseInput){
  const xapRows = catalog.filter(r=>r.series===XAP_SERIES);
  if (!xapRows.length) return { error: 'XAP-B-data mangler.' };
  const cmpInput = { ...baseInput, series: XAP_SERIES, ledere: '3F+N+PE' };
  const xapCat = { rows: xapRows, catalog };
  try{
    return price(xapCat, cmpInput);
  }catch(err){
    return { error: String(err && err.message ? err.message : err) };
  }
}

function updateXapComparisonUI(result){
  const line = $('xapCompareLine');
  const valueEl = $('xapComparisonValue');
  const currencyEl = $('xapComparisonCurrency');
  const wrapper = $('xapBomWrapper');
  if (!line || !valueEl || !wrapper) return;
  if (!result){
    line.hidden = true;
    if (currencyEl) currencyEl.hidden = false;
    valueEl.textContent = '';
    wrapper.hidden = true;
    renderBomTable('xapBomTbl', []);
    return;
  }
  line.hidden = false;
  const hasError = Boolean(result.error);
  if (currencyEl) currencyEl.hidden = hasError;
  if (hasError){
    valueEl.textContent = result.error;
    wrapper.hidden = true;
    renderBomTable('xapBomTbl', []);
    return;
  }
  valueEl.textContent = fmtNO.format(result.totalExMontasje || 0);
  wrapper.hidden = false;
  renderBomTable('xapBomTbl', result.bom || []);
}

function refreshCalculatedBoxItems(){
  if (!lastCalc || !lastCalcInput || !Array.isArray(catalog) || !catalog.length) return false;
  const input = {
    ...deepClone(lastCalcInput),
    boxItems: normalizeBoxItems(pendingBoxItems),
    specialElementItems: normalizeSpecialElementItems(pendingSpecialElementItems),
    boxQty: 0,
    boxSel: ''
  };
  const series = input.series || $('series')?.value || '';
  const rows = catalog.filter(r=>r.series === series);
  if (!series || !rows.length) return false;
  const out = price({ rows, catalog }, input);
  input.marginRate = out.marginRate;
  input.montasjeMarginRate = out.montasjeMarginRate;
  input.engineeringMarginRate = out.engineeringMarginRate;
  input.opphengMarginRate = out.opphengMarginRate;
  input.tapOffMarginRate = normalizeMarginRate(input.tapOffMarginRate ?? currentTapOffMarginRate, DEFAULT_MARGIN_RATE);
  lastCalcInput = deepClone(input);
  Object.assign(lastCalc, {
    material: out.material,
    tapOffBoxTotal: out.tapOffBoxTotal,
    specialElementTotal: out.specialElementTotal,
    marginRate: out.marginRate,
    marginFactor: out.marginFactor,
    margin: out.margin,
    subtotal: out.subtotal,
    freight: out.freight,
    totalExMontasje: out.totalExMontasje,
    totalInclMontasje: out.totalInclMontasje,
    totalInclEngineering: out.totalInclEngineering,
    totalInclOppheng: out.totalInclOppheng,
    total: out.total,
    rawUnitPrices: deepClone(out.rawUnitPrices || {}),
    tapOffMarginRate: input.tapOffMarginRate,
    bom: deepClone(out.bom)
  });
  lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
  lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
  renderBomTable('bomTbl', out.bom);
  [
    ['mat', out.material],
    ['margin', out.margin],
    ['subtotal', out.subtotal],
    ['freight', out.freight],
    ['totalExMontasje', out.totalExMontasje],
    ['montasje', out.montasje?.cost],
    ['montasjeMargin', out.montasjeMargin],
    ['totalInclMontasje', out.totalInclMontasje],
    ['engineering', out.engineering?.cost],
    ['engineeringMargin', out.engineeringMargin],
    ['totalInclEngineering', out.totalInclEngineering],
    ['oppheng', out.oppheng?.cost],
    ['opphengMargin', out.opphengMargin],
    ['total', out.totalInclOppheng]
  ].forEach(([id, value])=>{
    const el = $(id);
    const number = Number(value);
    if (el && Number.isFinite(number)) el.textContent = fmtNO.format(number);
  });
  if (lastEmailPayload){
    lastEmailPayload.inputs = deepClone(input);
    lastEmailPayload.bom = deepClone(out.bom);
    if (!lastEmailPayload.totals) lastEmailPayload.totals = {};
    Object.assign(lastEmailPayload.totals, {
      material: out.material,
      tapOffBoxTotal: out.tapOffBoxTotal,
      specialElementTotal: out.specialElementTotal,
      marginRate: out.marginRate,
      margin: out.margin,
      subtotal: out.subtotal,
      freight: out.freight,
      total: out.total,
      totalExMontasje: out.totalExMontasje,
      totalInclMontasje: out.totalInclMontasje,
      totalInclEngineering: out.totalInclEngineering,
      totalInclOppheng: out.totalInclOppheng,
      rawUnitPrices: deepClone(out.rawUnitPrices || {}),
      tapOffMarginRate: input.tapOffMarginRate,
      tapOffOfferTotal: lastCalc.tapOffOfferTotal,
      specialElementOfferTotal: lastCalc.specialElementOfferTotal
    });
  }
  updateSelectedAddonTotalUI();
  return true;
}

// --- app ---
let catalog=[];
const ampOptionsBySeries = new Map();
let isDirty = false;

function markDirty(){
  isDirty = true;
  const st = $('status');
  if (st) st.textContent = 'Beregn for å få inkludere endringer';
  const saveBtn = $('saveLineBtn');
  if (saveBtn) saveBtn.disabled = true;
}

function markClean(){
  isDirty = false;
  const st = $('status');
  if (st) st.textContent = 'OK';
  const saveBtn = $('saveLineBtn');
  if (saveBtn) saveBtn.disabled = false;
}

function applyUsdRateToCatalog(){
  if (!Array.isArray(catalog)) return;
  catalog.forEach(item=>{
    if (Number.isFinite(item.unit_price_usd)){
      item.unit_price = convertUsdToNok(item.unit_price_usd);
    }
  });
}

function rebuildAmpLookup(){
  ampOptionsBySeries.clear();
  if (!Array.isArray(catalog) || !catalog.length) return;
  const grouped = new Map();
  catalog.forEach(item=>{
    if (!item || !item.series) return;
    if (item.type && /box$/i.test(item.type)) return;
    const ampValue = Number(deriveAmp(item));
    if (!Number.isFinite(ampValue)) return;
    if (ampValue >= EXCLUDED_AMP_MIN && ampValue <= EXCLUDED_AMP_MAX) return;
    const key = String(item.series);
    if (!grouped.has(key)){
      grouped.set(key, new Set());
    }
    grouped.get(key).add(ampValue);
  });
  grouped.forEach((set, key)=>{
    const sorted = Array.from(set).sort((a,b)=>a-b);
    ampOptionsBySeries.set(key, sorted);
  });
}

function updateUsdRateFromMarket(snapshot){
  const fx = pickFxData(snapshot);
  const next = Number(fx?.usd?.rate);
  if (!Number.isFinite(next) || next <= 0) return;
  if (Math.abs(next - usdToNokRate) < 0.0005) return;
  usdToNokRate = next;
  applyUsdRateToCatalog();
  markDirty();
}

function updateTapOffConfigVisibility(){
  const series = $('series')?.value || '';
  const isEpoxySeries = series === EPOXY_IP68_SERIES;
  const distValue = $('dist')?.value || '';
  const showConfig = !isEpoxySeries && distValue === 'Ja';
  const configRow = $('tapOffConfigRow');
  if (configRow) configRow.hidden = !showConfig;
  if (!showConfig){
    pendingBoxItems = [];
    const boxSel = $('boxSel');
    const boxQty = $('boxQty');
    const boxInnmat = $('boxInnmatSum');
    if (boxSel) boxSel.value = '';
    if (boxQty) boxQty.value = '';
    if (boxInnmat) boxInnmat.value = '';
  }
  renderPendingBoxItems();
}

function updateSpecialElementConfigVisibility(){
  const showConfig = ($('specialElement')?.value || 'Nei') === 'Ja';
  const configRow = $('specialElementConfigRow');
  if (configRow) configRow.hidden = !showConfig;
  if (!showConfig){
    pendingSpecialElementItems = [];
    const typeEl = $('specialElementType');
    const qtyEl = $('specialElementQty');
    const sumEl = $('specialElementSum');
    if (typeEl) typeEl.value = '';
    if (qtyEl) qtyEl.value = '';
    if (sumEl) sumEl.value = '';
  }
  renderPendingSpecialElementItems();
}

window.addEventListener('DOMContentLoaded', async ()=>{
  initDashboardShell();
  await initProjectDashboard();
  initMarketDataTicker();
  if (!hasCalculatorUI()){
    return;
  }
  try{
    ['meter','v90h','v90v','fbQty','boxQty','boxInnmatSum','specialElementQty','specialElementSum'].forEach(id=>{
      const el = $(id);
      if (el) el.value = '';
    });
    pendingBoxItems = [];
    renderPendingBoxItems();
    pendingSpecialElementItems = [];
    renderPendingSpecialElementItems();

    const all = [];
    for (const p of RAW_CSV_PATHS){
      try{
        const res = await fetch(p,{cache:'no-store'}); if (!res.ok) continue;
        const txt = await res.text();
        all.push(...parseCSVAuto(txt));
      }catch{}
    }
    catalog = adaptRawToCatalog(all);
    applyUsdRateToCatalog();
    rebuildAmpLookup();

    const sendBtnInit = document.getElementById('sendRequestBtn');
    if (sendBtnInit){
      sendBtnInit.disabled = true;
    }

    const seriesEl = $('series');
    if (seriesEl){
      seriesEl.addEventListener('change', refreshUIBySeries);
    }
    const meterEl = $('meter');
    if (meterEl){
      meterEl.addEventListener('change', ()=>Math.ceil(Number(meterEl.value||0)));
      meterEl.addEventListener('blur', ()=>Math.ceil(Number(meterEl.value||0)));
    }

    const rateInput = $('montasjeHourlyRate');
    const opphengInput = $('opphengRate');
    const rateToggle = $('rateToggle');
    const engineeringRateInput = $('engineeringHourlyRate');
    const engineeringRateToggle = $('engineeringRateToggle');
    if (rateInput){
      if (!rateInput.value) rateInput.value = String(DEFAULT_HOURLY_RATE);
      const syncRate = ()=>{
        rateInput.value = String(sanitizeHourlyRate(rateInput.value));
        updateMontasjePreview();
        markDirty();
      };
      rateInput.addEventListener('input', ()=>{
        if (rateToggle && !rateToggle.checked) return;
        updateMontasjePreview();
        markDirty();
      });
      rateInput.addEventListener('change', syncRate);
      rateInput.addEventListener('blur', syncRate);
    }
    if (opphengInput){
      const sanitizeOpphengInput = (markDirtyAfter)=>{
        if (rateToggle && !rateToggle.checked) return;
        const raw = opphengInput.value;
        const defaultRate = toNum(opphengInput.dataset.defaultRate);
        const fallbackRate = Number.isFinite(defaultRate) ? defaultRate : 0;
        const sanitized = sanitizeOpphengRate(raw, fallbackRate);
        if (!(sanitized === fallbackRate && fallbackRate === 0 && raw.trim()==='')){
          opphengInput.value = (sanitized || sanitized === 0) ? String(sanitized) : '';
        }else{
          opphengInput.value = '';
        }
        opphengInput.dataset.appliedValue = String(sanitized);
        opphengInput.dataset.userOverride = sanitized !== fallbackRate ? 'true' : 'false';
        updateMontasjePreview();
        if (markDirtyAfter) markDirty();
      };
      opphengInput.addEventListener('input', ()=>{
        if (rateToggle && !rateToggle.checked) return;
        opphengInput.dataset.userOverride = 'true';
        updateMontasjePreview();
        markDirty();
      });
      opphengInput.addEventListener('change', ()=>sanitizeOpphengInput(true));
      opphengInput.addEventListener('blur', ()=>sanitizeOpphengInput(false));
    }
    if (rateToggle){
      rateToggle.checked = false;
      const applyRateLock = (markDirtyAfter)=>{
        const locked = !rateToggle.checked;
        setInputLocked(rateInput, locked);
        if (locked && rateInput){
          rateInput.value = String(DEFAULT_HOURLY_RATE);
        }
        if (opphengInput){
          setInputLocked(opphengInput, locked);
          if (locked){
            opphengInput.dataset.userOverride = 'false';
            opphengInput.value = '';
          }
        }
        updateMontasjePreview();
        if (markDirtyAfter) markDirty();
      };
      rateToggle.addEventListener('change', ()=>applyRateLock(true));
      applyRateLock(false);
    }else{
      setInputLocked(rateInput, false);
      if (opphengInput) setInputLocked(opphengInput, false);
    }
    if (engineeringRateInput){
      if (!engineeringRateInput.value) engineeringRateInput.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
      const syncEngineeringRate = ()=>{
        engineeringRateInput.value = String(sanitizeHourlyRate(engineeringRateInput.value, DEFAULT_ENGINEERING_HOURLY_RATE));
        updateEngineeringPreview();
        markDirty();
      };
      engineeringRateInput.addEventListener('input', ()=>{
        if (engineeringRateToggle && !engineeringRateToggle.checked) return;
        updateEngineeringPreview();
        markDirty();
      });
      engineeringRateInput.addEventListener('change', syncEngineeringRate);
      engineeringRateInput.addEventListener('blur', syncEngineeringRate);
    }
    if (engineeringRateToggle){
      engineeringRateToggle.checked = false;
      const applyEngineeringRateLock = (markDirtyAfter)=>{
        const locked = !engineeringRateToggle.checked;
        setInputLocked(engineeringRateInput, locked);
        if (locked && engineeringRateInput){
          engineeringRateInput.value = String(DEFAULT_ENGINEERING_HOURLY_RATE);
        }
        updateEngineeringPreview();
        if (markDirtyAfter) markDirty();
      };
      engineeringRateToggle.addEventListener('change', ()=>applyEngineeringRateLock(true));
      applyEngineeringRateLock(false);
    }else{
      setInputLocked(engineeringRateInput, false);
    }
    ['meter','v90h','v90v'].forEach(id=>{
      const el = $(id);
      if (!el) return;
      el.addEventListener('input', ()=>{ updateMontasjePreview(); markDirty(); });
      el.addEventListener('change', ()=>{ updateMontasjePreview(); markDirty(); });
    });
    const ampSelEl = $('ampSelect');
    if (ampSelEl){
      ampSelEl.addEventListener('change', updateMontasjePreview);
    }
    // Sett dist til 'Nei' hvis tom, både ved last og når feltet forlates
    const distEl = $('dist');
    if (distEl){
      const ensureDist = ()=>{ if (!distEl.value) distEl.value = 'Nei'; };
      ensureDist();
      distEl.addEventListener('blur', ensureDist);
      distEl.addEventListener('change', ()=>{
        updateTapOffConfigVisibility();
        markDirty();
      });
    }
    const specialElementEl = $('specialElement');
    if (specialElementEl){
      specialElementEl.value = 'Nei';
      specialElementEl.addEventListener('change', ()=>{
        updateSpecialElementConfigVisibility();
        markDirty();
      });
    }
    const addBoxItemBtn = $('addBoxItemBtn');
    if (addBoxItemBtn){
      addBoxItemBtn.addEventListener('click', ()=>{
        const series = $('series')?.value || '';
        if (series === EPOXY_IP68_SERIES) return;
        if (($('dist')?.value || '') !== 'Ja') return;
        const boxSelValue = String($('boxSel')?.value || '').trim();
        const boxQtyRaw = Number($('boxQty')?.value || 0);
        const boxQtyValue = Number.isFinite(boxQtyRaw) ? Math.max(0, Math.round(boxQtyRaw)) : 0;
        const boxInnmatRaw = toNum($('boxInnmatSum')?.value || 0);
        const boxInnmatValue = Number.isFinite(boxInnmatRaw) ? Math.max(0, round2(boxInnmatRaw)) : 0;
        if (!boxSelValue){
          const st = $('status');
          if (st) st.textContent = 'Velg type boks før du legger til.';
          return;
        }
        if (!boxQtyValue){
          const st = $('status');
          if (st) st.textContent = 'Angi antall bokser før du legger til.';
          return;
        }
        pendingBoxItems.push({
          id: generateTapOffItemId(),
          boxSel: boxSelValue,
          boxQty: boxQtyValue,
          innmatSum: boxInnmatValue
        });
        const boxSelEl = $('boxSel');
        const boxQtyEl = $('boxQty');
        const boxInnmatEl = $('boxInnmatSum');
        if (boxSelEl) boxSelEl.value = '';
        if (boxQtyEl) boxQtyEl.value = '';
        if (boxInnmatEl) boxInnmatEl.value = '';
        renderPendingBoxItems();
        if (lastCalc && lastCalcInput){
          try{
            refreshCalculatedBoxItems();
            const st = $('status');
            if (st) st.textContent = 'Boks lagt til i BOM og tilbudssum. Lagre linjen for å beholde endringen.';
            const saveBtn = $('saveLineBtn');
            if (saveBtn) saveBtn.disabled = false;
          }catch(err){
            const st = $('status');
            if (st) st.textContent = String(err.message || err);
          }
        } else {
          markDirty();
        }
      });
    }
    const addSpecialElementBtn = $('addSpecialElementBtn');
    if (addSpecialElementBtn){
      addSpecialElementBtn.addEventListener('click', ()=>{
        if (($('specialElement')?.value || 'Nei') !== 'Ja') return;
        const selection = String($('specialElementType')?.value || '').trim();
        const qtyRaw = Number($('specialElementQty')?.value || 0);
        const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
        const unitSumRaw = toNum($('specialElementSum')?.value || 0);
        const unitSum = Number.isFinite(unitSumRaw) ? Math.max(0, round2(unitSumRaw)) : 0;
        if (!selection){
          const st = $('status');
          if (st) st.textContent = 'Velg type element før du legger til.';
          return;
        }
        if (!qty){
          const st = $('status');
          if (st) st.textContent = 'Angi antall elementer før du legger til.';
          return;
        }
        if (!unitSum){
          const st = $('status');
          if (st) st.textContent = 'Angi sum per element før du legger til.';
          return;
        }
        pendingSpecialElementItems.push({
          id: generateSpecialElementItemId(),
          selection,
          qty,
          unitSum
        });
        const typeEl = $('specialElementType');
        const qtyEl = $('specialElementQty');
        const sumEl = $('specialElementSum');
        if (typeEl) typeEl.value = '';
        if (qtyEl) qtyEl.value = '';
        if (sumEl) sumEl.value = '';
        renderPendingSpecialElementItems();
        if (lastCalc && lastCalcInput){
          try{
            refreshCalculatedBoxItems();
            const st = $('status');
            if (st) st.textContent = 'Spesialelement lagt til i BOM og resultatsum. Lagre linjen for å beholde endringen.';
            const saveBtn = $('saveLineBtn');
            if (saveBtn) saveBtn.disabled = false;
          }catch(err){
            const st = $('status');
            if (st) st.textContent = String(err.message || err);
          }
        } else {
          markDirty();
        }
      });
    }
    refreshUIBySeries();
    updateTapOffConfigVisibility();
    updateSpecialElementConfigVisibility();
    setCurrentMarginRate(DEFAULT_MATERIAL_MARGIN_RATE);
    setCurrentMontasjeMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentEngineeringMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentOpphengMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentTapOffMarginRate(DEFAULT_MARGIN_RATE);
    applyCalculatorQueryContext();

    const marginConfigBtn = $('marginConfigBtn');
    if (marginConfigBtn){
      marginConfigBtn.addEventListener('click', ()=>openMarginModal('material'));
    }
    const montasjeDgConfigBtn = $('montasjeDgConfigBtn');
    if (montasjeDgConfigBtn){
      montasjeDgConfigBtn.addEventListener('click', ()=>openMarginModal('montasje'));
    }
    const engineeringDgConfigBtn = $('engineeringDgConfigBtn');
    if (engineeringDgConfigBtn){
      engineeringDgConfigBtn.addEventListener('click', ()=>openMarginModal('engineering'));
    }
    const opphengDgConfigBtn = $('opphengDgConfigBtn');
    if (opphengDgConfigBtn){
      opphengDgConfigBtn.addEventListener('click', ()=>openMarginModal('oppheng'));
    }

    const frSel = document.getElementById('freightRate');
    if (frSel){
      frSel.addEventListener('change', ()=>{
        if (!lastCalc) return;
        recalcLastTotalsFromCurrentRates();
      });
    }

    // Markér status som "Oppdater..." ved endringer i parametere
    const dirtySelectors = [
      '#series','#dist','#meter','#v90h','#v90v','#ampSelect','#ledere',
      '#startEl','#sluttEl','#fbQty','#boxQty','#boxSel','#boxInnmatSum',
      '#specialElement','#specialElementType','#specialElementQty','#specialElementSum'
    ];
    dirtySelectors.forEach(sel=>{
      const el = document.querySelector(sel);
      if (!el) return;
      if (el.dataset.markDirtyBound) return;
      el.addEventListener('input', markDirty);
      el.addEventListener('change', markDirty);
      el.dataset.markDirtyBound = '1';
    });

  function enhanceNumberSteppers() {
    const ids = ['meter','v90h','v90v','fbQty','boxQty','specialElementQty'];
    ids.forEach(id=>{
      const input = document.getElementById(id);
      if (!input || input.dataset.enhanced) return;

      input.setAttribute('min','0');
      input.setAttribute('step','1');
      input.placeholder = input.placeholder || '0';

      const wrap = document.createElement('div');
      wrap.className = 'stepper';

      const plus = document.createElement('button');
      plus.type = 'button';
      plus.className = 'btn-step';
      plus.textContent = '+';

      const minus = document.createElement('button');
      minus.type = 'button';
      minus.className = 'btn-step';
      minus.textContent = '-';

      const parent = input.parentNode;
      if (!parent) return;
      parent.insertBefore(wrap, input);
      wrap.appendChild(input);
      wrap.appendChild(plus);
      wrap.appendChild(minus);

      const clampInt = ()=>{
        const v = Math.max(0, Math.round(Number(input.value||0)));
        input.value = Number.isFinite(v) ? String(v) : '0';
      };

      minus.addEventListener('click', ()=>{
        const next = Math.max(0, Number(input.value||0) - 1);
        input.value = String(next);
        clampInt();
        input.dispatchEvent(new Event('change'));
      });
      plus.addEventListener('click', ()=>{
        const next = Math.max(0, Number(input.value||0) + 1);
        input.value = String(next);
        clampInt();
        input.dispatchEvent(new Event('change'));
      });

      input.addEventListener('input', clampInt);
      input.addEventListener('blur', clampInt);

      if (input.disabled){
        plus.disabled = true;
        minus.disabled = true;
      }

      input.dataset.enhanced = '1';
    });
  }

// Kall denne etter UI er bygd første gang
enhanceNumberSteppers();

    const statusEl = $('status');
    if (statusEl){
      statusEl.textContent = `CSV lastet (${catalog.length} varer)`;
    }
  }catch(e){
    const statusEl = $('status');
    if (statusEl){
      statusEl.textContent = 'Feil CSV: '+(e.message||e);
    }
  }
});

function refreshUIBySeries(){
  const series = $('series').value;
  const isEpoxySeries = series === EPOXY_IP68_SERIES;

  // Ledere låses for enkelte serier
  const ledereEl = $('ledere');
  if (ledereEl){
    Array.from(ledereEl.options).forEach(opt=>{
      if (opt.value === '3F+N' || opt.textContent.trim() === '3F+N'){
        opt.hidden = series === 'XCP-S';
        opt.disabled = series === 'XCP-S';
      }
    });
    if (seriesLocksLedere(series)){
      ledereEl.value = seriesLockedLedereValue(series);
      ledereEl.disabled = true;
    } else {
      ledereEl.disabled = false;
      if (series === 'XCP-S' && ledereEl.value === '3F+N'){
        ledereEl.value = '';
      }
      if (!ledereEl.value) ledereEl.value = '';
    }
  }

  // Startelement: skjul endetilførselsboks kun for Epoxy IP68
  const start = $('startEl');
  if (start){
    Array.from(start.options).forEach(opt=>{
      if (opt.value==='end_feed_unit') opt.hidden = isEpoxySeries;
    });
    if (isEpoxySeries && start.value==='end_feed_unit') start.value='';
  }

  // Sluttelement: skjul trafo for ikke-XCP-S, og endelokk kun for Epoxy IP68
  const slutt = $('sluttEl');
  const crtAllowed = seriesSupportsCrtFeed(series);
  Array.from(slutt.options).forEach(opt=>{
    if (opt.value==='crt_board_feed') opt.hidden = !crtAllowed;
    if (opt.value==='end_cover') opt.hidden = isEpoxySeries;
  });
  if (!crtAllowed && slutt.value==='crt_board_feed') slutt.value='';
  if (isEpoxySeries && slutt.value==='end_cover') slutt.value='';

  // Amp-valg
  const ampSelectEl = $('ampSelect');
  if (ampSelectEl){
    const previousValue = ampSelectEl.value;
    const ampList = ampOptionsBySeries.get(series) || [];
    const options = ampList.map(a=>`<option value="${a}">${a}</option>`).join('');
    ampSelectEl.innerHTML = '<option value="">Velg…</option>' + options;
    const prevExists = ampList.some(val=>String(val) === String(previousValue));
    ampSelectEl.value = prevExists ? previousValue : '';
    const disabled = !series || !ampList.length;
    ampSelectEl.disabled = disabled;
    ampSelectEl.setAttribute('aria-disabled', disabled ? 'true' : 'false');
    ampSelectEl.title = disabled
      ? (!series ? 'Velg system for å få ampere-listen.' : 'Ingen ampere funnet for valgt system.')
      : '';
  }

  // Bokser
  const boxes = catalog.filter(r=>['plug_in_box','tap_off_box','bolt_on_box'].includes(r.type));
  const labelOf = t => t==='plug_in_box'?'Plug-in box (plast)':t==='tap_off_box'?'Tap-off box (metall)':'Bolt-on box (metall)';
  const seen = new Set(); const opts = [];
  [...boxes.filter(b=>b.series===series), ...boxes.filter(b=>b.series!==series)].forEach(b=>{
    if (b.type==='bolt_on_box' && series==='XCM') return;
    const key = `${b.type}|${deriveAmp(b)||''}`;
    if (seen.has(key)) return; seen.add(key);
    const txt = `${deriveAmp(b)||''}A · ${labelOf(b.type)}`.replace(/^A · /,'');
    opts.push({v:`${b.type}|${deriveAmp(b)||''}`,t:txt});
  });
  opts.sort((a,b)=> (parseInt(a.t) || 1e9) - (parseInt(b.t) || 1e9) || String(a.t).localeCompare(b.t,'no'));
  $('boxSel').innerHTML = '<option value="">Velg...</option>'+opts.map(o=>`<option value="${o.v}">${o.t}</option>`).join('');

  // Lås distribusjon/avtappingsfelt kun for Epoxy IP68
  const distEl = $('dist');
  if (distEl){
    if (isEpoxySeries){
      distEl.value = 'Nei';
      distEl.disabled = true;
    } else {
      distEl.disabled = false;
      if (!distEl.value) distEl.value = 'Nei';
    }
  }
  const boxQtyEl = $('boxQty');
  if (boxQtyEl){
    boxQtyEl.disabled = isEpoxySeries;
    if (isEpoxySeries){
      boxQtyEl.value = '0';
    }
    const stepper = boxQtyEl.closest('.stepper');
    if (stepper){
      stepper.querySelectorAll('.btn-step').forEach(btn=>{ btn.disabled = isEpoxySeries; });
    }
  }
  const boxSelEl = $('boxSel');
  if (boxSelEl){
    boxSelEl.disabled = isEpoxySeries;
    if (isEpoxySeries){
      boxSelEl.value = '';
    }
  }
  const boxInnmatEl = $('boxInnmatSum');
  if (boxInnmatEl){
    boxInnmatEl.disabled = isEpoxySeries;
    if (isEpoxySeries){
      boxInnmatEl.value = '';
    }
  }
  const addBoxItemBtn = $('addBoxItemBtn');
  if (addBoxItemBtn){
    addBoxItemBtn.disabled = isEpoxySeries;
  }

  updateTapOffConfigVisibility();
  updateMontasjePreview();
}

// ekspansjons-modal
function askExpansionIfNeeded(meter){
  return new Promise(resolve=>{
    if (meter <= 30){ resolve(false); return; }
    const bd = $('expModal'); bd.style.display='flex';
    const yes = $('expYes'), no = $('expNo');
    const done = (v)=>{ bd.style.display='none'; yes.onclick=null; no.onclick=null; resolve(v); };
    yes.onclick = ()=>done(true);
    no.onclick  = ()=>done(false);
  });
}

async function sendCalculationEmail(payload){
  const res = await fetch(buildApiUrl('/api/send-calculation-email'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (!res.ok){
    let errorText = `Send foresp\u00f8rsel feilet (${res.status})`;
    try{
      const data = await res.json();
      if (data && typeof data.error === 'string' && data.error.trim()){
        errorText += `: ${data.error.trim()}`;
      }
    }catch(_jsonErr){
      try{
        const txt = await res.text();
        if (txt && txt.trim()) errorText += `: ${txt.trim()}`;
      }catch(_textErr){}
    }
    const err = new Error(appendApiBaseHint(errorText, res.status));
    err.status = res.status;
    throw err;
  }
}

// beregn
const calcBtn = $('calcBtn');
if (calcBtn){
  calcBtn.addEventListener('click', async ()=>{
  if (!authState.loggedIn){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Logg inn for \u00E5 beregne.';
    return;
  }
  try{
    if (!catalog.length) throw new Error('Ingen varer i katalog.');

    const series = $('series').value;
    const isEpoxySeries = series === EPOXY_IP68_SERIES;
    if (isEpoxySeries){
      $('dist').value = 'Nei';
    } else if (!$('dist').value){
      $('dist').value = 'Nei';
    }
    const dist   = isEpoxySeries ? false : ($('dist').value==='Ja');
    const meter  = Math.ceil(Number($('meter').value || 0));
    const v90_h  = Number($('v90h').value || 0);
    const v90_v  = Number($('v90v').value || 0);
    const ampSel = $('ampSelect').value;
    const ledere = seriesLocksLedere(series) ? seriesLockedLedereValue(series) : $('ledere').value;
    const startEl= $('startEl').value;
    const sluttEl= $('sluttEl').value;
    const fbQty  = Number($('fbQty').value || 0);
    const distValue = $('dist').value;
    const boxItems = (isEpoxySeries || distValue !== 'Ja')
      ? []
      : normalizeBoxItems(pendingBoxItems);
    const specialElementItems = ($('specialElement')?.value || 'Nei') === 'Ja'
      ? normalizeSpecialElementItems(pendingSpecialElementItems)
      : [];
    const boxQty = 0;
    const boxSel = '';

    if (!series) throw new Error('Velg system.');
    if (!meter) throw new Error('Angi meter (heltall).');
    if (!ampSel) throw new Error('Velg ampere.');
    if (!seriesLocksLedere(series) && !ledere) throw new Error('Velg ledere.');
    if (!startEl) throw new Error('Velg startelement.');
    if (!sluttEl) throw new Error('Velg sluttelement.');
    const lineNumberInputEl = $('lineNumberInput');
    const lineNumberValue = (lineNumberInputEl?.value || '').trim();
    if (!lineNumberValue){
      const statusEl = $('status');
      if (statusEl) statusEl.textContent = 'Oppgi linjenummer f\u00F8r beregning.';
      if (lineNumberInputEl) lineNumberInputEl.focus();
      return;
    }
    projectState.currentLineNumber = lineNumberValue;
    if (lineNumberInputEl){
      lineNumberInputEl.value = lineNumberValue;
    }

    const expansionYes = await askExpansionIfNeeded(meter);
    const amp = Number(ampSel);
    const rows = catalog.filter(r=>r.series===series);
    const freightRate = Number((document.getElementById('freightRate')?.value) || 0.10);
    const montasjeSettings = readMontasjeSettingsFromUI();
    const engineeringSettings = readEngineeringSettingsFromUI();

    const cat = { rows, catalog };
    const priceInput = {
      series, dist, meter, v90_h, v90_v, ampere: amp, ledere,
      startEl, sluttEl,
      fbQty, boxQty, boxSel, boxItems, specialElementItems,
      expansionYes, freightRate, marginRate: currentMarginRate,
      montasjeMarginRate: currentMontasjeMarginRate,
      engineeringMarginRate: currentEngineeringMarginRate,
      opphengMarginRate: currentOpphengMarginRate,
      tapOffMarginRate: currentTapOffMarginRate,
      montasjeSettings, engineeringSettings
    };
    const out = price(cat, priceInput);
    priceInput.marginRate = out.marginRate;
    priceInput.montasjeMarginRate = out.montasjeMarginRate;
    priceInput.engineeringMarginRate = out.engineeringMarginRate;
    priceInput.opphengMarginRate = out.opphengMarginRate;
    priceInput.tapOffMarginRate = currentTapOffMarginRate;
    setCurrentMarginRate(out.marginRate);
    setCurrentMontasjeMarginRate(out.montasjeMarginRate);
    setCurrentEngineeringMarginRate(out.engineeringMarginRate);
    setCurrentOpphengMarginRate(out.opphengMarginRate);
    lastCalcInput = deepClone(priceInput);
    let xapComparison = null;
    if (shouldCompareXap(series)){
      xapComparison = computeXapComparison(priceInput);
    }
    $('mat').textContent      = fmtNO.format(out.material);
    $('margin').textContent   = fmtNO.format(out.margin);
    $('subtotal').textContent = fmtNO.format(out.subtotal);
    $('freight').textContent  = fmtNO.format(out.freight);
    const montasjeEl = document.getElementById('montasje');
    if (montasjeEl) montasjeEl.textContent = fmtNO.format(out.montasje.cost);
    const montasjeMarginVal = round2(out.montasjeMargin ?? 0);
    const montasjeMarginEl = document.getElementById('montasjeMargin');
    if (montasjeMarginEl) montasjeMarginEl.textContent = fmtNO.format(montasjeMarginVal);
    const montasjeDetailText = formatMontasjeDetail(out.montasje);
    const montasjeDetailEl = document.getElementById('montasjeDetail');
    if (montasjeDetailEl) montasjeDetailEl.textContent = montasjeDetailText;
    const engineeringEl = document.getElementById('engineering');
    if (engineeringEl) engineeringEl.textContent = fmtNO.format(out.engineering.cost);
    const engineeringMarginVal = round2(out.engineeringMargin ?? 0);
    const engineeringMarginEl = document.getElementById('engineeringMargin');
    if (engineeringMarginEl) engineeringMarginEl.textContent = fmtNO.format(engineeringMarginVal);
    const engineeringDetailText = formatEngineeringDetail(out.engineering);
    const engineeringDetailEl = document.getElementById('engineeringDetail');
    if (engineeringDetailEl) engineeringDetailEl.textContent = engineeringDetailText;
    const opphengEl = document.getElementById('oppheng');
    if (opphengEl) opphengEl.textContent = fmtNO.format(out.oppheng.cost);
    const opphengMarginVal = round2(out.opphengMargin ?? 0);
    const opphengMarginEl = document.getElementById('opphengMargin');
    if (opphengMarginEl) opphengMarginEl.textContent = fmtNO.format(opphengMarginVal);
    const opphengDetailText = formatOpphengDetail(out.oppheng);
    const opphengDetailEl = document.getElementById('opphengDetail');
    if (opphengDetailEl) opphengDetailEl.textContent = opphengDetailText;
    const totalExEl = document.getElementById('totalExMontasje');
    if (totalExEl) totalExEl.textContent = fmtNO.format(out.totalExMontasje);
    const totalInclMontasjeVal = round2(out.totalInclMontasje ?? calculateDgPricing(out.montasje.cost, out.montasjeMarginRate).totalWithDg);
    const totalInclMontasjeEl = document.getElementById('totalInclMontasje');
    if (totalInclMontasjeEl) totalInclMontasjeEl.textContent = fmtNO.format(totalInclMontasjeVal);
    const totalInclEngineeringVal = round2(out.totalInclEngineering ?? calculateDgPricing(out.engineering.cost, out.engineeringMarginRate).totalWithDg);
    const totalInclEngineeringEl = document.getElementById('totalInclEngineering');
    if (totalInclEngineeringEl) totalInclEngineeringEl.textContent = fmtNO.format(totalInclEngineeringVal);
    const totalInclOpphengVal = round2(out.totalInclOppheng ?? calculateDgPricing(out.oppheng.cost, out.opphengMarginRate).totalWithDg);
    $('total').textContent = fmtNO.format(totalInclOpphengVal);
    const calcTimestamp = new Date().toISOString();
    lastCalc = {
      lineNumber: lineNumberValue,
      timestamp: calcTimestamp,
      material: out.material,
      tapOffBoxTotal: out.tapOffBoxTotal,
      specialElementTotal: out.specialElementTotal,
      marginRate: out.marginRate,
      marginFactor: out.marginFactor,
      margin: out.margin,
      montasjeMarginRate: out.montasjeMarginRate,
      montasjeMargin: montasjeMarginVal,
      engineeringMarginRate: out.engineeringMarginRate,
      engineeringMargin: engineeringMarginVal,
      opphengMarginRate: out.opphengMarginRate,
      opphengMargin: opphengMarginVal,
      tapOffMarginRate: currentTapOffMarginRate,
      subtotal: out.subtotal,
      freightRate,
      freight: out.freight,
      totalExMontasje: out.totalExMontasje,
      totalInclMontasje: totalInclMontasjeVal,
      totalInclEngineering: totalInclEngineeringVal,
      totalInclOppheng: totalInclOpphengVal,
      montasje: out.montasje,
      engineering: out.engineering,
      oppheng: out.oppheng,
      montasjeDetail: montasjeDetailText,
      engineeringDetail: engineeringDetailText,
      opphengDetail: opphengDetailText,
      total: totalInclOpphengVal,
      rawUnitPrices: deepClone(out.rawUnitPrices || {}),
      bom: deepClone(out.bom)
    };
    lastCalc.tapOffOfferTotal = calculateTapOffOfferTotal(lastCalc);
    lastCalc.specialElementOfferTotal = calculateSpecialElementOfferTotal(lastCalc);
    updateSelectedAddonTotalUI();
    markClean();

    renderBomTable('bomTbl', out.bom);
    document.getElementById('results').hidden = false;
    updateProjectMetaDisplay();
    updateXapComparisonUI(xapComparison);

    lastEmailPayload = {
      project: projectState.currentProject,
      customer: projectState.currentCustomer,
      lineNumber: lineNumberValue,
      inputs: deepClone(priceInput),
      totals: {
        material: out.material,
        tapOffBoxTotal: out.tapOffBoxTotal,
        specialElementTotal: out.specialElementTotal,
        marginRate: out.marginRate,
        margin: out.margin,
        montasjeMarginRate: out.montasjeMarginRate,
        montasjeMargin: montasjeMarginVal,
        engineeringMarginRate: out.engineeringMarginRate,
        engineeringMargin: engineeringMarginVal,
        opphengMarginRate: out.opphengMarginRate,
        opphengMargin: opphengMarginVal,
        subtotal: out.subtotal,
        freight: out.freight,
        total: totalInclOpphengVal,
        totalExMontasje: out.totalExMontasje,
        totalInclMontasje: totalInclMontasjeVal,
        totalInclEngineering: totalInclEngineeringVal,
        totalInclOppheng: totalInclOpphengVal,
        rawUnitPrices: deepClone(out.rawUnitPrices || {}),
        tapOffMarginRate: currentTapOffMarginRate,
        tapOffOfferTotal: calculateTapOffOfferTotal({ bom: out.bom, tapOffMarginRate: currentTapOffMarginRate }),
        specialElementOfferTotal: calculateSpecialElementOfferTotal({ bom: out.bom, tapOffMarginRate: currentTapOffMarginRate })
      },
      bom: out.bom
    };
    updateSelectedAddonTotalUI();

    document.getElementById('exportCsv').onclick = ()=>{
      const bomForExport = Array.isArray(lastEmailPayload?.bom) ? lastEmailPayload.bom : out.bom;
      const header = ['code','type','series','ampere','ledere','antall','enhet','sum'];
      const lines = [header.join(',')].concat(bomForExport.map(b=>[b.code,b.type,b.series,b.ampere,b.ledere,b.antall,b.enhet,b.sum].join(',')));
      const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'BOM.csv'; a.click();
    };
    const sendBtn = document.getElementById('sendRequestBtn');
    if (sendBtn){
      sendBtn.disabled = false;
      sendBtn.textContent = 'Send foresp\u00f8rsel';
      sendBtn.onclick = async ()=>{
        if (!lastEmailPayload){
          console.warn('Ingen beregning \u00e5 sende.');
          return;
        }
        const originalText = sendBtn.textContent;
        sendBtn.disabled = true;
        sendBtn.textContent = 'Sender...';
        try{
          await sendCalculationEmail(lastEmailPayload);
          sendBtn.textContent = 'Sent';
          const statusEl = $('status');
          if (statusEl) statusEl.textContent = 'Foresp\u00f8rsel sendt.';
        }catch(err){
          console.warn('E-postsending feilet', err);
          sendBtn.textContent = 'Feil, pr\u00f8v igjen';
          sendBtn.disabled = false;
          const statusEl = $('status');
          if (statusEl) statusEl.textContent = String(err.message||err);
          return;
        }
        setTimeout(()=>{
          sendBtn.textContent = originalText;
          sendBtn.disabled = false;
        }, 2000);
      };
    }
  }catch(err){
    $('status').textContent = String(err.message||err);
    lastEmailPayload = null;
    updateXapComparisonUI(null);
    const sendBtn = document.getElementById('sendRequestBtn');
    if (sendBtn){
      sendBtn.disabled = true;
      sendBtn.textContent = 'Send foresp\u00f8rsel';
    }
  }
  });
}
