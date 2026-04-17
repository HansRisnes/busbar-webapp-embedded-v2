// v2 + v2.1, XCP-S/XCM, distribusjon/feeder, avtappingsbokser, ekspansjon-modal >30 m

const RAW_CSV_PATHS = [
  'data/busbar-webapp-embedded-v2.csv',
  'data/busbar-webapp-embedded-v2.1.csv',
  'data/XAP_busbar_prisliste_ekstrakt_5W_clean.csv'
];
const XAP_SERIES = 'XAP-B';
const COMPARISON_ELIGIBLE_SERIES = Object.freeze(['XCM','XCP-S']);
const ENABLE_XAP_COMPARISON = false; // Pauset inntil videre
const LEDERE_LOCKED_SERIES = Object.freeze(['XCM', XAP_SERIES]);
const CRT_FEED_ALLOWED_SERIES = Object.freeze(['XCP-S', XAP_SERIES]);
const seriesLocksLedere = series => LEDERE_LOCKED_SERIES.includes(series);
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
const MAX_MARGIN_RATE = 0.95;
let lastCalc = null; // delsummer for live frakt-oppdatering
let lastCalcInput = null;
const AUTH_PASSWORD = 'busbar';
const AUTH_SESSION_KEY = 'busbar.auth.session.v1';
let authState = { loggedIn: false, username: '' };
const LEGACY_PROJECTS_STORAGE_KEY = 'busbar.projects.v1';
const PROJECTS_STORAGE_KEY_PREFIX = 'busbar.projects.user.v2';
const PROJECT_SYNC_DEBOUNCE_MS = 800;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PROJECT_SORT_STORAGE_KEY = 'busbar.project.sort.v1';
const LINE_SORT_STORAGE_KEY = 'busbar.line.sort.v1';
const PROJECT_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
const LINE_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
const projectSyncState = {
  timerId: null,
  inFlight: false,
  pending: false
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
  projects: [],
  expandedProjectId: null,
  projectSort: 'date_newest',
  lineSort: 'date_newest'
};
const projectModalState = {
  mode: 'create',
  projectId: null,
  saveLineAfterCreate: false
};
const projectMarginModalState = {
  projectId: null
};
let lastEmailPayload = null;

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
const fmtFxNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 4, maximumFractionDigits: 4 });
const fmtTimestampNO = new Intl.DateTimeFormat('no-NO', { dateStyle: 'short', timeStyle: 'short' });
const fmtPercentNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
let currentMarginRate = DEFAULT_MARGIN_RATE;
let currentMontasjeMarginRate = DEFAULT_MARGIN_RATE;
let currentEngineeringMarginRate = DEFAULT_MARGIN_RATE;
let currentOpphengMarginRate = DEFAULT_MARGIN_RATE;
let currentDgModalTarget = 'material';
const toNum = x => {
  if (x===undefined || x===null) return NaN;
  const v = Number(String(x).replace(/\s/g,'').replace(',','.'));
  return Number.isFinite(v) ? v : NaN;
};
function pick(row, names){ for (const n of names){ if (n in row && row[n]!=='' && row[n]!==undefined) return row[n]; } return ''; }

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

function normalizeMarginRate(value, fallback = DEFAULT_MARGIN_RATE){
  const raw = Number(value);
  if (!Number.isFinite(raw)) return fallback;
  const asRate = raw > 1 ? raw / 100 : raw;
  if (!Number.isFinite(asRate)) return fallback;
  if (asRate < 0) return 0;
  if (asRate >= 1) return MAX_MARGIN_RATE;
  return Math.min(MAX_MARGIN_RATE, asRate);
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
    return normalizeMarginRate(1 - material / subtotal, DEFAULT_MARGIN_RATE);
  }
  return DEFAULT_MARGIN_RATE;
}

function resolveDgRate(fromInput, fromTotals, fallback = DEFAULT_MARGIN_RATE){
  const inputRate = normalizeMarginRate(fromInput, NaN);
  if (Number.isFinite(inputRate)) return inputRate;
  const totalsRate = normalizeMarginRate(fromTotals, NaN);
  if (Number.isFinite(totalsRate)) return totalsRate;
  return fallback;
}

function updateDgLabel(labelId, rate){
  const normalizedRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
  const percentTxt = fmtPercentNO.format(normalizedRate * 100);
  const labelEl = $(labelId);
  if (labelEl){
    labelEl.textContent = `DG ${percentTxt} %:`;
  }
}

function updateMarginUI(){
  updateDgLabel('marginLabel', currentMarginRate);
  updateDgLabel('montasjeDgLabel', currentMontasjeMarginRate);
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
  currentMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
  updateMarginUI();
  return currentMarginRate;
}

function setCurrentMontasjeMarginRate(rate){
  currentMontasjeMarginRate = normalizeMarginRate(rate, DEFAULT_MARGIN_RATE);
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
  return { base: baseTotal, total: round2(sum) };
}

function updateSelectedAddonTotalUI(){
  const totalEl = $('selectedAddonTotal');
  if (!totalEl) return;
  const sum = calculateSelectedAddonTotal(lastCalc);
  totalEl.textContent = fmtNO.format(Number.isFinite(sum.total) ? sum.total : 0);
}

function calculateDgPricing(baseCost, dgRate){
  const base = round2(Number(baseCost) || 0);
  const normalizedRate = normalizeMarginRate(dgRate, DEFAULT_MARGIN_RATE);
  const factor = marginFactorFromRate(normalizedRate);
  if (!(factor > 0)){
    throw new Error('DG-faktor må være større enn 0.');
  }
  const totalWithDg = round2(base / factor);
  const dg = round2(totalWithDg - base);
  return { base, dgRate: normalizedRate, dg, totalWithDg };
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
  const normalizedMarginRate = normalizeMarginRate(marginRate, DEFAULT_MARGIN_RATE);
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
  const montasjePricing = calculateDgPricing(montasjeCost, montasjeMarginRate);
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
    subtotal: recalculated.subtotal,
    freight: recalculated.freight,
    totalExMontasje: recalculated.totalExMontasje,
    totalInclMontasje: recalculated.totalInclMontasje,
    totalInclEngineering: recalculated.totalInclEngineering,
    totalInclOppheng: recalculated.totalInclOppheng,
    total: recalculated.total
  });
  if (lastCalcInput){
    lastCalcInput.marginRate = recalculated.marginRate;
    lastCalcInput.freightRate = recalculated.freightRate;
    lastCalcInput.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastCalcInput.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastCalcInput.opphengMarginRate = recalculated.opphengMarginRate;
  }
  if (lastEmailPayload?.inputs){
    lastEmailPayload.inputs.marginRate = recalculated.marginRate;
    lastEmailPayload.inputs.freightRate = recalculated.freightRate;
    lastEmailPayload.inputs.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastEmailPayload.inputs.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastEmailPayload.inputs.opphengMarginRate = recalculated.opphengMarginRate;
  }
  if (lastEmailPayload?.totals){
    lastEmailPayload.totals.marginRate = recalculated.marginRate;
    lastEmailPayload.totals.margin = recalculated.margin;
    lastEmailPayload.totals.montasjeMarginRate = recalculated.montasjeMarginRate;
    lastEmailPayload.totals.montasjeMargin = recalculated.montasjeMargin;
    lastEmailPayload.totals.engineeringMarginRate = recalculated.engineeringMarginRate;
    lastEmailPayload.totals.engineeringMargin = recalculated.engineeringMargin;
    lastEmailPayload.totals.opphengMarginRate = recalculated.opphengMarginRate;
    lastEmailPayload.totals.opphengMargin = recalculated.opphengMargin;
    lastEmailPayload.totals.subtotal = recalculated.subtotal;
    lastEmailPayload.totals.freight = recalculated.freight;
    lastEmailPayload.totals.totalExMontasje = recalculated.totalExMontasje;
    lastEmailPayload.totals.totalInclMontasje = recalculated.totalInclMontasje;
    lastEmailPayload.totals.totalInclEngineering = recalculated.totalInclEngineering;
    lastEmailPayload.totals.totalInclOppheng = recalculated.totalInclOppheng;
    lastEmailPayload.totals.total = recalculated.total;
  }
  updateSelectedAddonTotalUI();
}

function getDgModalTitleByTarget(target){
  if (target === 'montasje') return 'Endre DG for montasje';
  if (target === 'engineering') return 'Endre DG for ingeniør';
  if (target === 'oppheng') return 'Endre DG for oppheng';
  return 'Endre DG for material';
}

function getCurrentDgRateByTarget(target){
  if (target === 'montasje') return currentMontasjeMarginRate;
  if (target === 'engineering') return currentEngineeringMarginRate;
  if (target === 'oppheng') return currentOpphengMarginRate;
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
    recalcLastTotalsFromCurrentRates();
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
      source: rawPoint.source || fallbackSource || ''
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

  const eurRate = fx.eur.rate;
  const eurEl = $('marketEurNok');
  if (eurEl){
    eurEl.textContent = Number.isFinite(eurRate) ? fmtFxNO.format(eurRate) : '--';
  }
  const eurMetaEl = $('marketEurMeta');
  if (eurMetaEl){
    eurMetaEl.textContent = buildFxMetaText(fx.eur) || 'Ingen data';
  }

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
    if (!hasValidUserEmail(username)) return;
    authState = { loggedIn: true, username };
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
        username: authState.username || ''
      }));
      return;
    }
    sessionStorage.removeItem(AUTH_SESSION_KEY);
  }catch(err){
    console.warn('Kunne ikke lagre innloggingsstatus', err);
  }
}

function updateAuthUI(){
  const calcBtn = $('calcBtn');
  if (calcBtn) calcBtn.disabled = !authState.loggedIn;

  const loginBtn = $('loginBtn');
  const logoutBtn = $('logoutBtn');
  const userLabel = $('authUser');
  const adminBtn = $('adminPageBtn');

  if (loginBtn) loginBtn.hidden = authState.loggedIn;
  if (logoutBtn) logoutBtn.hidden = !authState.loggedIn;
  if (adminBtn) adminBtn.hidden = !authState.loggedIn;
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
    newProjectBtn.disabled = !authState.loggedIn;
  }
  const createProjectButtons = Array.from(document.querySelectorAll('button[data-action="create-project"]'));
  createProjectButtons.forEach(btn=>{
    btn.disabled = !authState.loggedIn;
  });
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

function handleLoginSubmit(){
  const passwordInput = $('loginPassword');
  const usernameInput = $('loginUsername');
  const errorEl = $('loginError');
  const password = passwordInput ? passwordInput.value : '';
  if (password !== AUTH_PASSWORD){
    if (errorEl) errorEl.textContent = 'Feil passord.';
    if (passwordInput) passwordInput.focus();
    return;
  }
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
  authState = { loggedIn: true, username };
  persistAuthToSession();
  hideLoginModal();
  projectState.projects = loadProjectsFromStorage();
  projectState.expandedProjectId = null;
  sortProjects();
  updateProjectHistories();
  renderProjectDashboard();
  updateProjectMetaDisplay();
  updateAuthUI();
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = '';
  void syncProjectsForCurrentUser();
}

const loginBtn = $('loginBtn');
if (loginBtn){
  loginBtn.addEventListener('click', showLoginModal);
}
const logoutBtn = $('logoutBtn');
if (logoutBtn){
  logoutBtn.addEventListener('click', ()=>{
    authState = { loggedIn: false, username: '' };
    if (projectSyncState.timerId){
      clearTimeout(projectSyncState.timerId);
      projectSyncState.timerId = null;
    }
    projectSyncState.pending = false;
    persistAuthToSession();
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

function normalizeProject(raw){
  if (!raw) return null;
  const fallback = new Date().toISOString();
  const selectedAddonConfig = normalizeSelectedAddonConfig(raw.selectedAddonConfig || null, null);
  return {
    id: raw.id || generateProjectId(),
    name: String(raw.name || '').trim(),
    customer: String(raw.customer || '').trim(),
    contactPerson: String(raw.contactPerson || raw.contact || '').trim(),
    createdAt: raw.createdAt || fallback,
    updatedAt: raw.updatedAt || fallback,
    selectedAddonConfig,
    lines: Array.isArray(raw.lines) ? raw.lines : []
  };
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

function clearProjectOverviewForLoggedOutUser(){
  projectState.projects = [];
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
    if (candidateTs >= existingTs){
      merged.set(key, normalized);
    }
  };
  (Array.isArray(localProjects) ? localProjects : []).forEach(add);
  (Array.isArray(remoteProjects) ? remoteProjects : []).forEach(add);
  return Array.from(merged.values());
}

async function fetchUserProjectsFromServer(email){
  const query = encodeURIComponent(email);
  const res = await fetch(buildApiUrl(`/api/user-projects?email=${query}`), {
    cache: 'no-store'
  });
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
  return projects.map(normalizeProject).filter(Boolean);
}

async function pushUserProjectsToServer(email, projects){
  const res = await fetch(buildApiUrl('/api/user-projects/sync'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      email,
      projects: Array.isArray(projects) ? projects : []
    })
  });
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
  return syncedProjects.map(normalizeProject).filter(Boolean);
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
    await pushUserProjectsToServer(email, projectState.projects);
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
    const remoteProjects = await fetchUserProjectsFromServer(email);
    const mergedProjects = mergeProjectsByLatest(projectState.projects, remoteProjects);
    projectState.projects = mergedProjects;
    sortProjects();
    updateProjectHistories();
    saveProjectsToStorage({ skipRemoteSync: true });
    renderProjectDashboard();
    updateProjectMetaDisplay();
    queueProjectSync({ immediate: true });
  }catch(err){
    console.warn('Kunne ikke hente prosjekter fra server', err);
  }
}

function loadProjectsFromStorage(){
  const email = getCurrentUserEmail();
  const storageKey = getProjectsStorageKeyForEmail(email);
  if (!storageKey) return [];
  if (typeof localStorage === 'undefined') return [];
  try{
    const stored = localStorage.getItem(storageKey);
    if (!stored) return [];
    const parsed = JSON.parse(stored);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeProject).filter(Boolean);
  }catch(err){
    console.warn('Kunne ikke lese prosjekter', err);
    return [];
  }
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
    queueProjectSync();
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
  updateSortControlValues();
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
  projectState.projects.forEach(project=>{
    addToHistory(projectState.projectHistory, project.name);
    addToHistory(projectState.customerHistory, project.customer);
    addToHistory(projectState.contactHistory, project.contactPerson);
  });
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

function createProject(projectName, customerName, contactPerson){
  const now = new Date().toISOString();
  const project = {
    id: generateProjectId(),
    name: projectName,
    customer: customerName,
    contactPerson: contactPerson,
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

function updateProject(projectId, updates){
  const target = getProjectById(projectId);
  if (!target) return null;
  target.name = updates.name;
  target.customer = updates.customer;
  target.contactPerson = String(updates.contactPerson || '').trim();
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
  const modal = $('projectModal');
  if (!modal) return;
  modal.style.display = 'flex';
  const errorEl = $('projectError');
  if (errorEl) errorEl.textContent = '';
  const projectInput = $('projectNameInput');
  const customerInput = $('customerNameInput');
  const contactInput = $('contactPersonInput');
  const titleEl = $('projectTitle');
  if (titleEl){
    titleEl.textContent = projectModalState.mode === 'edit' ? 'Oppdater prosjekt' : 'Nytt prosjekt';
  }
  if (projectInput){
    if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      projectInput.value = existing?.name || '';
    } else {
      projectInput.value = '';
    }
    projectInput.focus();
    const len = projectInput.value.length;
    try{
      projectInput.setSelectionRange(len, len);
    }catch(_err){
      /* ignore selection errors */
    }
  }
  if (customerInput){
    if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      customerInput.value = existing?.customer || '';
    } else {
      customerInput.value = '';
    }
  }
  if (contactInput){
    if (projectModalState.mode === 'edit'){
      const existing = getProjectById(projectModalState.projectId);
      contactInput.value = existing?.contactPerson || '';
    } else {
      contactInput.value = '';
    }
  }
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
  projectModalState.mode = 'create';
  projectModalState.projectId = null;
  projectModalState.saveLineAfterCreate = false;
}

function persistProjectInfo(projectName, customerName, contactPerson, options = {}){
  const trimmedName = projectName.trim();
  const trimmedCustomer = customerName.trim();
  const trimmedContact = contactPerson.trim();
  if (options.projectId){
    updateProject(options.projectId, {
      name: trimmedName,
      customer: trimmedCustomer,
      contactPerson: trimmedContact
    });
    setActiveProject(options.projectId);
    return;
  }
  const created = createProject(trimmedName, trimmedCustomer, trimmedContact);
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
  return {
    includeMontasje,
    includeEngineering,
    includeOppheng,
    showMontasje: includeMontasje && showMontasje,
    showEngineering: includeEngineering && showEngineering,
    showOppheng: includeOppheng && showOppheng
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
    showOppheng: Boolean($('showOppheng')?.checked)
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
  if (includeMontasje) includeMontasje.checked = config.includeMontasje;
  if (includeEngineering) includeEngineering.checked = config.includeEngineering;
  if (includeOppheng) includeOppheng.checked = config.includeOppheng;
  if (showMontasje) showMontasje.checked = config.showMontasje;
  if (showEngineering) showEngineering.checked = config.showEngineering;
  if (showOppheng) showOppheng.checked = config.showOppheng;
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
  const boxQty = Number(input.boxQty);
  if (Number.isFinite(boxQty) && boxQty > 0) parts.push(`Bokser: ${fmtIntNO.format(boxQty)}`);
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
  let total = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
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
  if (includeMontasje) includeMontasje.checked = config.includeMontasje;
  if (includeEngineering) includeEngineering.checked = config.includeEngineering;
  if (includeOppheng) includeOppheng.checked = config.includeOppheng;
  if (showMontasje) showMontasje.checked = config.showMontasje;
  if (showEngineering) showEngineering.checked = config.showEngineering;
  if (showOppheng) showOppheng.checked = config.showOppheng;
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
  return DEFAULT_MARGIN_RATE;
}

function getProjectMaterialMarginStats(project){
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  const rates = lines
    .map(resolveLineMaterialMarginRate)
    .filter(rate=>Number.isFinite(rate))
    .map(rate=>normalizeMarginRate(rate, DEFAULT_MARGIN_RATE));
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
    Math.abs(stats.minRate - DEFAULT_MARGIN_RATE) > epsilon ||
    Math.abs(stats.maxRate - DEFAULT_MARGIN_RATE) > epsilon
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
  const normalizedRate = normalizeMarginRate(nextRate, DEFAULT_MARGIN_RATE);
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
    const montasjeMarginRate = resolveDgRate(inputs?.montasjeMarginRate, totals?.montasjeMarginRate, DEFAULT_MARGIN_RATE);
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
  const nextRate = normalizeMarginRate(parsed > 1 ? parsed / 100 : parsed, DEFAULT_MARGIN_RATE);
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

  wrapper.appendChild(buildSelectorGroup('Inkluder i tilbud:', 'include', 'includeKey'));
  wrapper.appendChild(buildSelectorGroup('Synliggjør pris:', 'show', 'showKey'));
  return wrapper;
}

function formatLineTotal(line){
  const total = resolveLineDisplayTotal(line);
  return Number.isFinite(total) ? `${fmtNO.format(total)} NOK` : 'Ingen sum';
}

function sanitizeDownloadFileName(value, fallback = 'tilbud'){
  const raw = String(value || '').trim() || fallback;
  const cleaned = raw.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').replace(/\s+/g, ' ').trim();
  return cleaned || fallback;
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
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ project })
  });
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
  const headerName = getFilenameFromContentDisposition(res.headers.get('Content-Disposition'));
  const offerNumber = String(res.headers.get('X-Offer-Number') || '').trim();
  const fallbackName = `Tilbud-${sanitizeDownloadFileName(project?.name || 'prosjekt')}${offerNumber ? `-${offerNumber}` : ''}.docx`;
  return {
    blob,
    fileName: headerName || fallbackName,
    offerNumber
  };
}

async function requestGenerateProjectOffer(projectId, triggerBtn){
  const project = getProjectById(projectId);
  if (!project) return;

  const buttonEl = triggerBtn && triggerBtn.tagName === 'BUTTON' ? triggerBtn : null;
  const originalText = buttonEl ? buttonEl.textContent : '';
  let failed = false;
  if (buttonEl){
    buttonEl.disabled = true;
    buttonEl.textContent = 'Genererer...';
  }

  try{
    const generated = await generateProjectOffer(project);
    const blobUrl = URL.createObjectURL(generated.blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = generated.fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(()=>URL.revokeObjectURL(blobUrl), 1000);

    if (buttonEl){
      buttonEl.textContent = generated.offerNumber ? `Generert ${generated.offerNumber}` : 'Generert';
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

function renderProjectDashboard(){
  const listEl = $('projectList');
  if (!listEl) return;
  listEl.innerHTML = '';
  if (!projectState.projects.length){
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    const text = document.createElement('p');
    text.textContent = 'Ingen prosjekter er registrert ennå.';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn';
    btn.dataset.action = 'create-project';
    btn.textContent = 'Opprett nytt prosjekt';
    btn.disabled = !authState.loggedIn;
    empty.appendChild(text);
    empty.appendChild(btn);
    listEl.appendChild(empty);
    projectState.expandedProjectId = null;
    return;
  }
  const frag = document.createDocumentFragment();
  projectState.projects.forEach(project=>{
    const expanded = projectState.expandedProjectId === project.id;
    const projectLines = Array.isArray(project.lines) ? project.lines : [];
    const projectTotal = round2(projectLines.reduce((sum, line)=>{
      const lineTotal = resolveLineDisplayTotal(line);
      return Number.isFinite(lineTotal) ? sum + lineTotal : sum;
    }, 0));
    const row = document.createElement('section');
    row.className = 'project-row';
    if (expanded) row.classList.add('is-expanded');

    const head = document.createElement('div');
    head.className = 'project-row-head';
    const titleWrap = document.createElement('div');
    titleWrap.className = 'project-row-title';
    const title = document.createElement('h3');
    title.textContent = project.name || 'Uten navn';
    const customer = document.createElement('p');
    customer.textContent = project.customer ? `Kunde: ${project.customer}` : 'Kunde: -';
    const contact = document.createElement('p');
    contact.textContent = project.contactPerson ? `Kontaktperson: ${project.contactPerson}` : 'Kontaktperson: -';
    const created = document.createElement('p');
    created.className = 'project-row-meta';
    created.textContent = `Opprettet: ${formatProjectTimestamp(project.createdAt)}`;
    const summary = document.createElement('p');
    summary.className = 'project-row-meta';
    const lineCount = projectLines.length;
    summary.textContent = `Linjer: ${lineCount} | Totalsum linjer: ${fmtNO.format(projectTotal)} NOK`;
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
    setMarginBtn.disabled = !projectLines.length;

    const marginRow = document.createElement('div');
    marginRow.className = 'project-margin-row';
    marginRow.appendChild(marginBadge);
    marginRow.appendChild(setMarginBtn);

    titleWrap.appendChild(title);
    titleWrap.appendChild(customer);
    titleWrap.appendChild(contact);
    titleWrap.appendChild(created);
    titleWrap.appendChild(summary);
    titleWrap.appendChild(marginRow);
    const projectAddonConfig = getProjectSelectedAddonConfig(project);
    const projectAddonControl = buildAddonSelectorControl(projectAddonConfig, {
      className: 'project-inline-selectors',
      scope: 'project',
      projectId: project.id
    });
    titleWrap.appendChild(projectAddonControl);

    const actions = document.createElement('div');
    actions.className = 'project-row-actions';

    const detailBtn = document.createElement('button');
    detailBtn.type = 'button';
    detailBtn.className = 'btn alt';
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

    const generateOfferBtn = document.createElement('button');
    generateOfferBtn.type = 'button';
    generateOfferBtn.className = 'btn';
    generateOfferBtn.dataset.projectGenerateOffer = project.id;
    generateOfferBtn.textContent = 'Generer tilbud';
    generateOfferBtn.disabled = !projectLines.length;

    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn alt';
    editBtn.dataset.projectCardEdit = project.id;
    editBtn.textContent = 'Endre';

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'btn danger';
    deleteBtn.dataset.projectDelete = project.id;
    deleteBtn.textContent = 'Slett';

    actions.appendChild(detailBtn);
    actions.appendChild(newLineBtn);
    actions.appendChild(generateOfferBtn);
    actions.appendChild(editBtn);
    actions.appendChild(deleteBtn);

    head.appendChild(titleWrap);
    head.appendChild(actions);
    row.appendChild(head);

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

        const nameSpan = document.createElement('span');
        nameSpan.className = 'line-name';
        nameSpan.textContent = line.lineNumber || 'Uten linjenummer';

        const infoSpan = document.createElement('span');
        infoSpan.className = 'line-info';
        infoSpan.textContent = formatLineSummary(line);

        const totalSpan = document.createElement('span');
        totalSpan.className = 'line-total';
        totalSpan.textContent = formatLineTotal(line);

        const updatedSpan = document.createElement('span');
        updatedSpan.className = 'line-updated';
        updatedSpan.textContent = formatLineUpdatedText(line);

        lineBtn.appendChild(nameSpan);
        lineBtn.appendChild(infoSpan);
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

        const lineActionButtons = document.createElement('div');
        lineActionButtons.className = 'line-action-buttons';

        const lineEditBtn = document.createElement('button');
        lineEditBtn.type = 'button';
        lineEditBtn.className = 'btn alt line-edit-btn';
        lineEditBtn.dataset.lineEdit = line.id;
        lineEditBtn.dataset.projectId = project.id;
        lineEditBtn.textContent = 'Endre';

        const lineDeleteBtn = document.createElement('button');
        lineDeleteBtn.type = 'button';
        lineDeleteBtn.className = 'btn danger line-delete-btn';
        lineDeleteBtn.dataset.lineDelete = line.id;
        lineDeleteBtn.dataset.projectId = project.id;
        lineDeleteBtn.textContent = 'Slett';

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
}

async function initProjectDashboard(){
  projectState.projects = loadProjectsFromStorage();
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
  setActiveProject(project);
  projectState.expandedProjectId = project.id;
  renderProjectDashboard();

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
    const lineInput = $('lineNumberInput');
    if (lineInput) lineInput.focus();
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
  const lineInput = $('lineNumberInput');
  if (lineInput){
    lineInput.focus();
  }
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = 'Oppgi linjenummer for ny linje.';
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
  setCurrentMontasjeMarginRate(resolveDgRate(input?.montasjeMarginRate, NaN, DEFAULT_MARGIN_RATE));
  setCurrentEngineeringMarginRate(resolveDgRate(input?.engineeringMarginRate, NaN, DEFAULT_MARGIN_RATE));
  setCurrentOpphengMarginRate(resolveDgRate(input?.opphengMarginRate, NaN, DEFAULT_MARGIN_RATE));
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
}

function applySavedTotalsToUI(line){
  if (!line || !line.totals) return;
  const totals = line.totals;
  applySelectedAddonCheckboxes(line);
  const savedMarginRate = resolveMarginRateFromData({ totals, input: line.inputs });
  const savedMontasjeMarginRate = resolveDgRate(line.inputs?.montasjeMarginRate, totals?.montasjeMarginRate, DEFAULT_MARGIN_RATE);
  const savedEngineeringMarginRate = resolveDgRate(line.inputs?.engineeringMarginRate, totals?.engineeringMarginRate, DEFAULT_MARGIN_RATE);
  const savedOpphengMarginRate = resolveDgRate(line.inputs?.opphengMarginRate, totals?.opphengMarginRate, DEFAULT_MARGIN_RATE);
  setCurrentMarginRate(savedMarginRate);
  setCurrentMontasjeMarginRate(savedMontasjeMarginRate);
  setCurrentEngineeringMarginRate(savedEngineeringMarginRate);
  setCurrentOpphengMarginRate(savedOpphengMarginRate);
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
  const montasjePricing = calculateDgPricing(totals.montasje?.cost, savedMontasjeMarginRate);
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
    lastCalc.lineNumber = line.lineNumber || '';
    lastCalc.marginRate = savedMarginRate;
    lastCalc.marginFactor = marginFactorFromRate(savedMarginRate);
    lastCalc.montasjeMarginRate = savedMontasjeMarginRate;
    lastCalc.engineeringMarginRate = savedEngineeringMarginRate;
    lastCalc.opphengMarginRate = savedOpphengMarginRate;
  }
  lastCalcInput = line.inputs ? deepClone(line.inputs) : null;
  if (lastCalcInput){
    lastCalcInput.marginRate = savedMarginRate;
    lastCalcInput.montasjeMarginRate = savedMontasjeMarginRate;
    lastCalcInput.engineeringMarginRate = savedEngineeringMarginRate;
    lastCalcInput.opphengMarginRate = savedOpphengMarginRate;
  }
  const totalsForPayload = deepClone(totals);
  if (totalsForPayload){
    totalsForPayload.marginRate = savedMarginRate;
    totalsForPayload.montasjeMarginRate = savedMontasjeMarginRate;
    totalsForPayload.engineeringMarginRate = savedEngineeringMarginRate;
    totalsForPayload.opphengMarginRate = savedOpphengMarginRate;
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
  const dash = $('dashboardView');
  if (dash) dash.hidden = false;
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
      goToDashboard({ project: project.id });
    } else {
      goToDashboard();
    }
    return;
  }
  showDashboardView({ clearSelection: !project });
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
  ['series','dist','ampSelect','ledere','startEl','sluttEl','boxSel'].forEach(id=>{
    const el = $(id);
    if (el){
      el.value = '';
      el.disabled = false;
    }
  });
  ['meter','v90h','v90v','fbQty','boxQty'].forEach(id=>{
    const el = $(id);
    if (el){
      el.value = '';
    }
  });
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
  if (includeMontasje) includeMontasje.checked = true;
  if (includeEngineering) includeEngineering.checked = true;
  if (includeOppheng) includeOppheng.checked = true;
  if (showMontasje) showMontasje.checked = false;
  if (showEngineering) showEngineering.checked = false;
  if (showOppheng) showOppheng.checked = false;
  applyOfferAddonCheckboxConstraints();
  if (preserveProject && projectState.currentProjectId){
    const project = getProjectById(projectState.currentProjectId);
    if (project){
      applyProjectAddonCheckboxesToCalculator(project);
    }
  }
  refreshUIBySeries();
  setCurrentMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentMontasjeMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentEngineeringMarginRate(DEFAULT_MARGIN_RATE);
  setCurrentOpphengMarginRate(DEFAULT_MARGIN_RATE);
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
  const shouldSaveLineAfterCreate = !wasEditMode && projectModalState.saveLineAfterCreate;
  if (wasEditMode && projectModalState.projectId){
    persistProjectInfo(projectName, customerName, contactPerson, { projectId: projectModalState.projectId });
  } else {
    persistProjectInfo(projectName, customerName, contactPerson);
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

const saveLineBtn = $('saveLineBtn');
if (saveLineBtn){
  saveLineBtn.disabled = true;
  saveLineBtn.addEventListener('click', saveCurrentLineToProject);
}

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
  input.addEventListener('focus', ()=>{
    showSuggestions(listEl, source);
  });
  input.addEventListener('input', ()=>{
    hideSuggestions(listEl);
    updateProjectSubmitState();
    const errorEl = $('projectError');
    if (errorEl) errorEl.textContent = '';
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
      updateProjectSubmitState();
    }
  });
}

bindSuggestionBehaviour('projectNameInput', 'projectSuggestions', projectState.projectHistory);
bindSuggestionBehaviour('customerNameInput', 'customerSuggestions', projectState.customerHistory);
bindSuggestionBehaviour('contactPersonInput', 'contactSuggestions', projectState.contactHistory);

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

const DEFAULT_HOURLY_RATE = 657.8;
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

function getMontasjeProfileForAmp(amp){
  const a = Number(amp);
  if (!Number.isFinite(a) || a <= 0) return null;
  for (const row of MONTASJE_TIME_TABLE){
    if (a <= row.maxAmp) return row;
  }
  return MONTASJE_TIME_TABLE[MONTASJE_TIME_TABLE.length - 1];
}

function calculateMontasje({ meter, angles, amp, hourlyRate }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const totalAngles = Math.max(0, Math.round(Number(angles) || 0));
  const rate = sanitizeHourlyRate(hourlyRate, DEFAULT_HOURLY_RATE);
  const ampValue = Number(amp);
  const profile = getMontasjeProfileForAmp(ampValue);
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
  const ratesUnlocked = rateToggle ? rateToggle.checked : true;
  const montasjeLocked = !ratesUnlocked;
  if (rateEl){
    setInputLocked(rateEl, montasjeLocked);
    if (montasjeLocked){
      rateEl.value = String(DEFAULT_HOURLY_RATE);
    }
  }
  const hourlyRate = rateEl ? rateEl.value : DEFAULT_HOURLY_RATE;

  const montasjePreview = calculateMontasje({ meter, angles, amp, hourlyRate });

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
  if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))          return isDistGen ? 'straight_3m_dist' : 'straight_3m';
  if (/l\s*=\s*1501\s*[-–]\s*2000\b/.test(d))          return isDistGen ? 'straight_1501_2000_dist' : 'straight_1501_2000';
  if (/l\s*=\s*500\s*[-–]\s*1000\b/.test(d))           return isDistGen ? 'straight_500_1000_dist'  : 'straight_500_1000';

  // andre typer
  if (d.includes('horizontal') && (d.includes('elbow') || d.includes('90'))) return 'elbow_horizontal_90';
  if (d.includes('vertical')   && (d.includes('elbow') || d.includes('90'))) return 'elbow_vertical_90';
  if (/tap[\s-]*off\s*box/.test(d))  return 'tap_off_box';
  if (/plug[\s-]*in\s*box/.test(d))  return 'plug_in_box';
  if (/bolt[\s-]*on\s*box|b160\s*bolt/.test(d)) return 'bolt_on_box';
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

    const amp = Number.isFinite(tagAmp) ? tagAmp : extractAmpGeneric(r);

    list.push({ code, type, series, ampere: amp, unit_price: price, _desc: desc, _et2: et2H });
  });
  return list;
}

// match
function matchesLedere(row, ledere){
  if (!ledere) return true;
  if (row.ledere) return row.ledere === ledere;
  if (row.series === 'XCM') return true;
  const c = String(row.code||'').toUpperCase();
  const implied = c.includes('-3W') ? '3F+PE' : '3F+N+PE';
  return ledere === implied || !c;
}
function byTypeAmpSeries(rows, type, amp, series){ return rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp)); }
function byTypeAmpSeriesL(rows, type, amp, series, ledere){
  const m = rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp) && matchesLedere(r, ledere));
  return m || rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp)) || rows.find(r=>r.type===type && r.series===series);
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
  const direct = rows.find(r=>r.type==='fire_barrier_kit' && r.series===series && Number(deriveAmp(r))===amp);
  if (direct) return { code: direct.code, unit: toNum(direct.unit_price) };
  const ext = rows.find(r=>r.type==='fire_barrier_kit_external' && r.series===series && Number(deriveAmp(r))===amp);
  const int = rows.find(r=>r.type==='fire_barrier_kit_internal' && r.series===series && Number(deriveAmp(r))===amp);
  if (!ext && !int) return null;
  const unit = (ext?toNum(ext.unit_price):0) + (int?toNum(int.unit_price):0);
  const code = [ext?.code, int?.code].filter(Boolean).join('+') || 'FIRE-BARRIER';
  return { code, unit };
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
    const row = findByTypeSeriesAmp(rows, type, series, amp) || byTypeAmpSeriesL(rows,type,amp,series,ledere);
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

function price(cat, input){
  const bom=[];
  const push = (r,q)=>bom.push(makeLine(r, input.series, input.ampere, input.ledere, q));
  const need = (type)=> byTypeAmpSeries(cat.rows, type, input.ampere, input.series);

  // Startelement
if (input.startEl === 'board_feed'){
  const bf = needAnyAmp(cat.rows, 'board_feed', input.ampere, input.series);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.startEl === 'end_feed_unit'){
  const ef = needAnyAmp(cat.rows, 'end_feed_unit', input.ampere, input.series);
  if (!ef) throw new Error(`Mangler end_feed_unit for ${input.series}.`);
  push(ef,1);
}

// Sluttelement
if (input.sluttEl === 'board_feed'){
  const bf = needAnyAmp(cat.rows, 'board_feed', input.ampere, input.series);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.sluttEl === 'crt_board_feed'){
  if (!seriesSupportsCrtFeed(input.series)) throw new Error('Trafoelement er ikke tilgjengelig for valgt system.');
  const crt = needAnyAmp(cat.rows, 'crt_board_feed', input.ampere, input.series);
  if (!crt) throw new Error(`Mangler crt_board_feed for ${input.series}.`);
  push(crt,1);
} else if (input.sluttEl === 'end_cover'){
  const ec = needAnyAmp(cat.rows, 'end_cover', input.ampere, input.series);
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

  // Avtappingsbokser
  if (input.boxQty>0){
    if (input.boxSel){
      const [kind, ampStr] = input.boxSel.split('|');
      if (kind==='bolt_on_box' && input.series==='XCM') throw new Error('Bolt-on box kan ikke brukes på XCM.');
      const row = byBoxAll(cat.catalog, kind, ampStr?Number(ampStr):undefined, input.series);
      if (!row) throw new Error(`Mangler ${kind} ${ampStr||''}A.`);
      bom.push(makeLine(row, input.series, deriveAmp(row)||'', input.ledere, input.boxQty));
    }else{
      const tryKinds = ['plug_in_box','tap_off_box','bolt_on_box'];
      let row = null;
      for (const k of tryKinds){
        if (k==='bolt_on_box' && input.series==='XCM') continue;
        row = byBoxAll(cat.catalog, k, undefined, input.series);
        if (row) break;
      }
      if (!row) throw new Error('Ingen avtappingsbokser funnet i data.');
      bom.push(makeLine(row, input.series, deriveAmp(row)||'', input.ledere, input.boxQty));
    }
  }

  // Brann
  if (input.fbQty>0){
    const fb = getFireBarrier(cat.rows, input.ampere, input.series);
    if (!fb) throw new Error('Mangler fire barrier kit.');
    bom.push(makeCustomLine(fb.code, 'fire_barrier_kit', input.series, input.ampere, input.ledere, fb.unit, input.fbQty));
  }

  // Ekspansjon
  if (input.meter > 30 && input.expansionYes){
    const exp = byTypeAmpSeries(cat.rows,'expansion_unit',input.ampere,input.series);
    if (!exp) throw new Error(`Mangler expansion_unit for ${input.series} ${input.ampere}.`);
    push(exp,1);
  }

  const material = round2(bom.reduce((s,x)=>s+x.sum,0));
  const marginRate = normalizeMarginRate(input.marginRate, DEFAULT_MARGIN_RATE);
  const rate     = Number(input.freightRate ?? 0.10);
  const montasje = calculateMontasje({
    meter: input.meter,
    angles: (input.v90_h || 0) + (input.v90_v || 0),
    amp: input.ampere,
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
  return {
    bom,
    material,
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
window.addEventListener('DOMContentLoaded', async ()=>{
  await initProjectDashboard();
  initMarketDataTicker();
  if (!hasCalculatorUI()){
    return;
  }
  try{
    ['meter','v90h','v90v','fbQty','boxQty'].forEach(id=>{
      const el = $(id);
      if (el) el.value = '';
    });

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
    }
    refreshUIBySeries();
    setCurrentMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentMontasjeMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentEngineeringMarginRate(DEFAULT_MARGIN_RATE);
    setCurrentOpphengMarginRate(DEFAULT_MARGIN_RATE);
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
      '#startEl','#sluttEl','#fbQty','#boxQty','#boxSel'
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
    const ids = ['meter','v90h','v90v','fbQty','boxQty'];
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

  // Ledere låses for enkelte serier
  const ledereEl = $('ledere');
  if (ledereEl){
    if (seriesLocksLedere(series)){
      ledereEl.value = '3F+N+PE';
      ledereEl.disabled = true;
    } else {
      ledereEl.disabled = false;
      if (!ledereEl.value) ledereEl.value = '';
    }
  }

  // Sluttelement: skjul trafo for ikke-XCP-S
  const slutt = $('sluttEl');
  const crtAllowed = seriesSupportsCrtFeed(series);
  Array.from(slutt.options).forEach(opt=>{
    if (opt.value==='crt_board_feed') opt.hidden = !crtAllowed;
  });
  if (!crtAllowed && slutt.value==='crt_board_feed') slutt.value='';

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
    if (!$('dist').value) { $('dist').value = 'Nei'; }
    const dist   = ($('dist').value==='Ja');
    const meter  = Math.ceil(Number($('meter').value || 0));
    const v90_h  = Number($('v90h').value || 0);
    const v90_v  = Number($('v90v').value || 0);
    const ampSel = $('ampSelect').value;
    const ledere = seriesLocksLedere(series) ? '3F+N+PE' : $('ledere').value;
    const startEl= $('startEl').value;
    const sluttEl= $('sluttEl').value;
    const fbQty  = Number($('fbQty').value || 0);
    const boxQty = Number($('boxQty').value || 0);
    const boxSel = $('boxSel').value;

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
      fbQty, boxQty, boxSel,
      expansionYes, freightRate, marginRate: currentMarginRate,
      montasjeMarginRate: currentMontasjeMarginRate,
      engineeringMarginRate: currentEngineeringMarginRate,
      opphengMarginRate: currentOpphengMarginRate,
      montasjeSettings, engineeringSettings
    };
    const out = price(cat, priceInput);
    priceInput.marginRate = out.marginRate;
    priceInput.montasjeMarginRate = out.montasjeMarginRate;
    priceInput.engineeringMarginRate = out.engineeringMarginRate;
    priceInput.opphengMarginRate = out.opphengMarginRate;
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
      marginRate: out.marginRate,
      marginFactor: out.marginFactor,
      margin: out.margin,
      montasjeMarginRate: out.montasjeMarginRate,
      montasjeMargin: montasjeMarginVal,
      engineeringMarginRate: out.engineeringMarginRate,
      engineeringMargin: engineeringMarginVal,
      opphengMarginRate: out.opphengMarginRate,
      opphengMargin: opphengMarginVal,
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
      total: totalInclOpphengVal
    };
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
        totalInclOppheng: totalInclOpphengVal
      },
      bom: out.bom
    };

    document.getElementById('exportCsv').onclick = ()=>{
      const header = ['code','type','series','ampere','ledere','antall','enhet','sum'];
      const lines = [header.join(',')].concat(out.bom.map(b=>[b.code,b.type,b.series,b.ampere,b.ledere,b.antall,b.enhet,b.sum].join(',')));
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
