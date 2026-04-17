require('dotenv').config();
const path = require('path');
const fs = require('fs/promises');
const express = require('express');
const nodemailer = require('nodemailer');
const AdmZip = require('adm-zip');
const { ClientSecretCredential } = require('@azure/identity');

const STATIC_DATA_DIR = path.resolve(__dirname, '..', 'data');
const RUNTIME_DATA_DIR = (() => {
  const raw = String(process.env.DATA_DIR || '').trim();
  if (!raw) return STATIC_DATA_DIR;
  return path.resolve(raw);
})();
const PRICE_DATA_DIR = (() => {
  const raw = String(process.env.PRICE_DATA_DIR || '').trim();
  if (!raw) return STATIC_DATA_DIR;
  return path.resolve(raw);
})();
const DEFAULT_MARKET_FILE = path.resolve(RUNTIME_DATA_DIR, 'market-data.json');
const OFFER_TEMPLATE_FILE = path.resolve(
  __dirname,
  'templates',
  'tilbud',
  'tilbud-stromskinner-template.docx'
);
const OFFER_COUNTER_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-sequence.json');
const OFFER_PROJECT_NUMBERS_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-project-numbers.json');
const OFFER_REVISIONS_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-revisions.json');
const PROJECT_ARCHIVE_FILE = path.resolve(RUNTIME_DATA_DIR, 'project-archive.json');
const OFFER_LINE_BLOCK_START_TOKEN = '__BUSBAR_LINE_BLOCK_START__';
const OFFER_LINE_BLOCK_END_TOKEN = '__BUSBAR_LINE_BLOCK_END__';
const OFFER_FIRE_BLOCK_START_TOKEN = '__BUSBAR_FIRE_BLOCK_START__';
const OFFER_FIRE_BLOCK_END_TOKEN = '__BUSBAR_FIRE_BLOCK_END__';
const OFFER_OPPHENG_BLOCK_START_TOKEN = '__BUSBAR_OPPHENG_BLOCK_START__';
const OFFER_OPPHENG_BLOCK_END_TOKEN = '__BUSBAR_OPPHENG_BLOCK_END__';
const OFFER_PRICE_SOURCE_FILES = [
  path.resolve(PRICE_DATA_DIR, 'busbar-webapp-embedded-v2.csv'),
  path.resolve(PRICE_DATA_DIR, 'busbar-webapp-embedded-v2.1.csv'),
  path.resolve(PRICE_DATA_DIR, 'busbar-webapp-embedded-v2.2.csv')
];
const MARKET_HTTP_TIMEOUT_MS = Number(process.env.MARKET_HTTP_TIMEOUT_MS || 7000);
const MARKET_LME_URL =
  process.env.MARKET_LME_URL ||
  'https://query1.finance.yahoo.com/v7/finance/quote?symbols=ALI%3DF';
const MARKET_USER_AGENT =
  process.env.MARKET_USER_AGENT || 'BusbarPricing/1.0 (+https://busbar.no)';
const NORGES_BANK_USD_NOK_URL =
  'https://data.norges-bank.no/api/data/EXR/B.USD.NOK.SP?format=sdmx-json';
const NORGES_BANK_EUR_NOK_URL =
  'https://data.norges-bank.no/api/data/EXR/B.EUR.NOK.SP?format=sdmx-json';
const MARKET_DATA_FILE = (()=>{
  const preferred = String(process.env.MARKET_DATA_FILE || '').trim();
  if (preferred) return path.resolve(preferred);

  // Backward compatibility with older env naming.
  const legacy = String(process.env.MARKET_STATIC_FILE || '').trim();
  if (legacy) return path.resolve(legacy);

  return DEFAULT_MARKET_FILE;
})();
const MARKET_DAILY_REFRESH_HOUR = (()=>{
  const parsed = Number(process.env.MARKET_DAILY_REFRESH_HOUR || 6);
  if (!Number.isInteger(parsed) || parsed < 0 || parsed > 23) return 6;
  return parsed;
})();
const MARKET_RETRY_DELAY_MINUTES = (()=>{
  const parsed = Number(process.env.MARKET_RETRY_DELAY_MINUTES || 60);
  if (!Number.isFinite(parsed) || parsed < 1 || parsed > 1440) return 60;
  return Math.round(parsed);
})();
const MARKET_TIMEZONE =
  process.env.MARKET_TIMEZONE ||
  Intl.DateTimeFormat().resolvedOptions().timeZone ||
  'local';
const MARKET_DAILY_REFRESH_LABEL = `${String(MARKET_DAILY_REFRESH_HOUR).padStart(2, '0')}:00`;

const marketScheduleState = {
  timerId: null,
  lastRunAt: null,
  lastAttemptAt: null,
  lastSuccessAt: null,
  nextRunAt: null,
  lastError: null,
  status: 'idle',
  retryCount: 0
};

let marketCache = { payload: null };
let offerNumberLock = Promise.resolve();
let projectArchiveLock = Promise.resolve();
let fireBarrierPriceIndexPromise = null;
let marketRefreshInFlight = null;

fs.access(MARKET_DATA_FILE).catch(err=>{
  console.warn(`[market-data] Datafil utilgjengelig (${MARKET_DATA_FILE}): ${err.message}`);
});

const app = express();
app.use(express.json({ limit: '10mb' }));

const DEFAULT_CORS_ALLOWED_ORIGINS = [
  'https://hansrisnes.github.io'
];
const DEFAULT_LOCAL_ORIGINS = [
  'http://localhost:3000',
  'http://127.0.0.1:3000',
  'http://localhost:4173',
  'http://127.0.0.1:4173',
  'http://localhost:5173',
  'http://127.0.0.1:5173',
  'http://localhost:5500',
  'http://127.0.0.1:5500'
];

function parseCsvEnv(value) {
  return String(value || '')
    .split(',')
    .map(entry => entry.trim())
    .filter(Boolean);
}

const corsAllowLocalhost = String(process.env.CORS_ALLOW_LOCALHOST || 'true').trim().toLowerCase() !== 'false';
const configuredCorsOrigins = parseCsvEnv(process.env.CORS_ALLOWED_ORIGINS);
const corsAllowedOrigins = new Set([
  ...(configuredCorsOrigins.length ? configuredCorsOrigins : DEFAULT_CORS_ALLOWED_ORIGINS),
  ...(corsAllowLocalhost ? DEFAULT_LOCAL_ORIGINS : [])
]);
const corsAllowAllOrigins = corsAllowedOrigins.has('*');

function isCorsOriginAllowed(origin) {
  if (!origin) return true;
  if (corsAllowAllOrigins) return true;
  return corsAllowedOrigins.has(origin);
}

app.use((req, res, next) => {
  const origin = String(req.headers.origin || '').trim();
  const allowOrigin = !origin || isCorsOriginAllowed(origin);
  if (!allowOrigin) {
    if (req.method === 'OPTIONS') {
      return res.status(403).json({ error: 'Origin er ikke tillatt av CORS' });
    }
    return res.status(403).json({ error: 'Origin er ikke tillatt av CORS' });
  }

  if (origin) {
    res.setHeader('Access-Control-Allow-Origin', corsAllowAllOrigins ? '*' : origin);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(204);
  }
  return next();
});

const ADMIN_USERNAME = safeString(process.env.ADMIN_USERNAME || 'admin');
const ADMIN_PASSWORD = safeString(process.env.ADMIN_PASSWORD || 'change-me-admin');

if (!process.env.ADMIN_USERNAME || !process.env.ADMIN_PASSWORD) {
  console.warn(
    '[admin] ADMIN_USERNAME eller ADMIN_PASSWORD mangler i miljøvariabler. Bruker midlertidige standardverdier.'
  );
}

const requiredEnv = [
  'OAUTH_TENANT_ID',
  'OAUTH_CLIENT_ID',
  'OAUTH_CLIENT_SECRET',
  'SMTP_USER',
  'MAIL_TO'
];

requiredEnv.forEach(key => {
  if (!process.env[key]) {
    console.warn(`Environment variable ${key} mangler. Tjenesten kan ikke sende e-post uten denne.`);
  }
});

function round2(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function toFiniteNumber(value) {
  if (typeof value === 'string') {
    const normalized = value.replace(/\s/g, '').replace(',', '.');
    const parsedFromString = Number(normalized);
    if (Number.isFinite(parsedFromString)) return parsedFromString;
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return NaN;
  return parsed;
}

function safeString(value) {
  return String(value || '').trim();
}

function formatNoCurrency(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount)) return '';
  return new Intl.NumberFormat('no-NO', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(round2(amount));
}

function formatNoInteger(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount)) return '';
  return new Intl.NumberFormat('no-NO', {
    maximumFractionDigits: 0
  }).format(Math.round(amount));
}

function formatNoIntegerUp(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount)) return '';
  return new Intl.NumberFormat('no-NO', {
    maximumFractionDigits: 0
  }).format(Math.ceil(amount));
}

function formatOfferDate(date = new Date()) {
  return new Intl.DateTimeFormat('no-NO', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    timeZone: 'Europe/Oslo'
  }).format(date);
}

function getOfferYear(date = new Date()) {
  return Number(new Intl.DateTimeFormat('en-CA', {
    year: 'numeric',
    timeZone: 'Europe/Oslo'
  }).format(date));
}

function sanitizeFileName(value, fallback = 'tilbud') {
  const raw = safeString(value) || fallback;
  const cleaned = raw.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').replace(/\s+/g, ' ').trim();
  return cleaned || fallback;
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function withOfferNumberLock(task) {
  const run = offerNumberLock.then(()=>task());
  offerNumberLock = run.catch(()=>{});
  return run;
}

async function readJsonFile(filePath, fallbackValue) {
  try {
    const raw = await fs.readFile(filePath, 'utf8');
    return JSON.parse(raw);
  } catch (err) {
    if (err && err.code === 'ENOENT') return fallbackValue;
    throw err;
  }
}

async function writeJsonFile(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), 'utf8');
}

function withProjectArchiveLock(task) {
  const run = projectArchiveLock.then(() => task());
  projectArchiveLock = run.catch(() => {});
  return run;
}

function normalizeEmail(value) {
  return safeString(value).toLowerCase();
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ''));
}

function toIsoTimestamp(value, fallback = new Date().toISOString()) {
  const raw = safeString(value);
  if (!raw) return fallback;
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return fallback;
  return parsed.toISOString();
}

function safeJsonClone(value, fallback) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (_err) {
    return fallback;
  }
}

function generateRecordId(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(16).slice(2, 10)}`;
}

function normalizeLineRecord(raw) {
  if (!raw || typeof raw !== 'object') return null;
  const now = new Date().toISOString();
  const selectedAddonTotal = toFiniteNumber(raw.selectedAddonTotal);
  const bom = Array.isArray(raw.bom) ? safeJsonClone(raw.bom, []) : [];
  const inputs = (raw.inputs && typeof raw.inputs === 'object')
    ? safeJsonClone(raw.inputs, {})
    : {};
  const totals = (raw.totals && typeof raw.totals === 'object')
    ? safeJsonClone(raw.totals, {})
    : {};
  const selectedAddonConfig = (raw.selectedAddonConfig && typeof raw.selectedAddonConfig === 'object')
    ? safeJsonClone(raw.selectedAddonConfig, {})
    : {};
  return {
    id: safeString(raw.id) || generateRecordId('line'),
    lineNumber: safeString(raw.lineNumber),
    createdAt: toIsoTimestamp(raw.createdAt, now),
    updatedAt: toIsoTimestamp(raw.updatedAt || raw.createdAt, now),
    inputs,
    totals,
    bom,
    selectedAddonConfig,
    selectedAddonTotal: Number.isFinite(selectedAddonTotal) ? round2(selectedAddonTotal) : null
  };
}

function normalizeProjectRecord(raw) {
  if (!raw || typeof raw !== 'object') return null;
  const now = new Date().toISOString();
  const linesRaw = Array.isArray(raw.lines) ? raw.lines : [];
  const lines = linesRaw.map(normalizeLineRecord).filter(Boolean);
  const selectedAddonConfig = (raw.selectedAddonConfig && typeof raw.selectedAddonConfig === 'object')
    ? safeJsonClone(raw.selectedAddonConfig, {})
    : {};
  return {
    id: safeString(raw.id) || generateRecordId('proj'),
    name: safeString(raw.name),
    customer: safeString(raw.customer),
    contactPerson: safeString(raw.contactPerson || raw.contact),
    createdAt: toIsoTimestamp(raw.createdAt, now),
    updatedAt: toIsoTimestamp(raw.updatedAt || raw.createdAt, now),
    selectedAddonConfig,
    lines
  };
}

function normalizeStoredUserRecord(email, raw) {
  const now = new Date().toISOString();
  const projectsRaw = Array.isArray(raw?.projects) ? raw.projects : [];
  const projects = projectsRaw.map(normalizeProjectRecord).filter(Boolean);
  return {
    email,
    updatedAt: toIsoTimestamp(raw?.updatedAt, now),
    projects
  };
}

async function readProjectArchive() {
  const stored = await readJsonFile(PROJECT_ARCHIVE_FILE, { users: {} });
  const usersRaw = (stored && typeof stored === 'object' && stored.users && typeof stored.users === 'object')
    ? stored.users
    : {};
  const users = {};
  Object.entries(usersRaw).forEach(([key, value]) => {
    const email = normalizeEmail(key || value?.email);
    if (!isValidEmail(email)) return;
    users[email] = normalizeStoredUserRecord(email, value);
  });
  return { users };
}

async function writeProjectArchive(state) {
  const users = (state && typeof state === 'object' && state.users && typeof state.users === 'object')
    ? state.users
    : {};
  await writeJsonFile(PROJECT_ARCHIVE_FILE, { users });
}

function parseBasicAuthHeader(headerValue) {
  const header = String(headerValue || '');
  if (!header.toLowerCase().startsWith('basic ')) return null;
  const encoded = header.slice(6).trim();
  if (!encoded) return null;
  try {
    const decoded = Buffer.from(encoded, 'base64').toString('utf8');
    const separatorIndex = decoded.indexOf(':');
    if (separatorIndex < 0) return null;
    return {
      username: decoded.slice(0, separatorIndex),
      password: decoded.slice(separatorIndex + 1)
    };
  } catch (_err) {
    return null;
  }
}

function requireAdminAuth(req, res, next) {
  const auth = parseBasicAuthHeader(req.headers.authorization);
  if (!auth || auth.username !== ADMIN_USERNAME || auth.password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Ugyldig admin-innlogging' });
  }
  return next();
}

function addDays(date, days) {
  const base = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(base.getTime())) return new Date();
  const copy = new Date(base);
  copy.setDate(copy.getDate() + Number(days || 0));
  return copy;
}

function allocateOfferNumberFromState(state, date = new Date()) {
  const nextState = (state && typeof state === 'object') ? state : {};
  const years = (nextState.years && typeof nextState.years === 'object')
    ? nextState.years
    : {};
  const year = String(getOfferYear(date));
  const previous = Number(years[year]);
  const next = (Number.isInteger(previous) && previous >= 1000) ? previous + 1 : 1001;
  years[year] = next;
  nextState.years = years;
  return `${year}-${next}`;
}

async function allocateOfferNumber(date = new Date()) {
  return withOfferNumberLock(async ()=>{
    const state = await readJsonFile(OFFER_COUNTER_FILE, { years: {} });
    const offerNumber = allocateOfferNumberFromState(state, date);
    await writeJsonFile(OFFER_COUNTER_FILE, state);
    return offerNumber;
  });
}

function resolveProjectOfferKey(project) {
  const explicitId = safeString(project?.id);
  if (explicitId) return `id:${explicitId}`;
  const name = safeString(project?.name).toLowerCase();
  const customer = safeString(project?.customer).toLowerCase();
  const contact = safeString(project?.contactPerson || project?.contact).toLowerCase();
  return `meta:${name}|${customer}|${contact}` || 'meta:unknown';
}

async function allocateOfferIdentity(project, date = new Date()) {
  return withOfferNumberLock(async ()=>{
    const [counterState, projectNumbersRaw, revisionsRaw] = await Promise.all([
      readJsonFile(OFFER_COUNTER_FILE, { years: {} }),
      readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {}),
      readJsonFile(OFFER_REVISIONS_FILE, {})
    ]);

    const projectNumbers = (projectNumbersRaw && typeof projectNumbersRaw === 'object')
      ? projectNumbersRaw
      : {};
    const revisions = (revisionsRaw && typeof revisionsRaw === 'object')
      ? revisionsRaw
      : {};
    const key = resolveProjectOfferKey(project);

    let offerNumber = safeString(projectNumbers[key]);
    if (!/^\d{4}-\d+$/.test(offerNumber)) {
      offerNumber = allocateOfferNumberFromState(counterState, date);
      projectNumbers[key] = offerNumber;
    }

    const previousRevision = Number(revisions[key]);
    const revision = (Number.isInteger(previousRevision) && previousRevision >= 0)
      ? previousRevision + 1
      : 0;
    revisions[key] = revision;

    await Promise.all([
      writeJsonFile(OFFER_COUNTER_FILE, counterState),
      writeJsonFile(OFFER_PROJECT_NUMBERS_FILE, projectNumbers),
      writeJsonFile(OFFER_REVISIONS_FILE, revisions)
    ]);

    return { offerNumber, revision };
  });
}

function resolveSelectedAddonFlag(value, fallback = true) {
  if (typeof value === 'boolean') return value;
  if (value === 1 || value === '1' || value === 'true' || value === 'TRUE') return true;
  if (value === 0 || value === '0' || value === 'false' || value === 'FALSE') return false;
  return fallback;
}

function resolveLineSelectedAddonConfig(line, lineTotals = {}) {
  const raw = line?.selectedAddonConfig || lineTotals?.selectedAddonConfig || null;
  const includeMontasje = resolveSelectedAddonFlag(raw?.includeMontasje, true);
  const includeEngineering = resolveSelectedAddonFlag(raw?.includeEngineering, true);
  const includeOppheng = resolveSelectedAddonFlag(raw?.includeOppheng, true);
  const showMontasje = resolveSelectedAddonFlag(
    raw?.showMontasje,
    resolveSelectedAddonFlag(raw?.includeMontasje, false)
  );
  const showEngineering = resolveSelectedAddonFlag(
    raw?.showEngineering,
    resolveSelectedAddonFlag(raw?.includeEngineering, false)
  );
  const showOppheng = resolveSelectedAddonFlag(
    raw?.showOppheng,
    resolveSelectedAddonFlag(raw?.includeOppheng, false)
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

function formatNoCurrencyWithKr(value) {
  const formatted = formatNoCurrency(value);
  return formatted ? `kr. ${formatted}` : '';
}

function parseCsvLine(line) {
  const values = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === ',' && !inQuotes) {
      values.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  values.push(current);
  return values;
}

function normalizeSeriesForFireLookup(rawSeries) {
  const normalized = safeString(rawSeries).toUpperCase().replace(/\s+/g, '');
  if (!normalized) return '';
  if (normalized.includes('XCP')) return 'XCP-S';
  if (normalized.includes('XCM')) return 'XCM';
  if (normalized.includes('RCP')) return 'RCP';
  return normalized;
}

function extractAmpFromFireBarrierDesc(desc) {
  const raw = safeString(desc).toUpperCase();
  const explicitAmp = raw.match(/(\d{2,4})\s*A\b/);
  if (explicitAmp) return Number(explicitAmp[1]);

  const tagMatch = raw.match(/\b(3XB160|2XB210|2XB190|2XB160|B210|B190|B160|H470|H380|H300|H245|H200|H160)\b/);
  if (!tagMatch) return NaN;
  const map = {
    B160: 1250,
    B190: 1600,
    B210: 2000,
    '2XB160': 2500,
    '2XB190': 3200,
    '2XB210': 4000,
    '3XB160': 5000,
    H160: 1250,
    H200: 1600,
    H245: 2000,
    H300: 2500,
    H380: 3200,
    H470: 4000
  };
  return Number(map[tagMatch[1]]);
}

function detectFireBarrierType(desc) {
  const raw = safeString(desc).toLowerCase();
  const hasExternal = /(external|ext\.|utvendig|ytter)/.test(raw);
  const hasInternal = /(internal|int\.|innvendig|inner)/.test(raw);
  if (hasExternal && !hasInternal) return 'external';
  if (hasInternal && !hasExternal) return 'internal';
  return 'direct';
}

function resolveFireBarrierUnitFromBom(line) {
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  if (!bom.length) return NaN;

  let direct = NaN;
  let external = NaN;
  let internal = NaN;
  bom.forEach(entry=>{
    const type = safeString(entry?.type).toLowerCase();
    if (!type.includes('fire_barrier')) return;
    const unit = toFiniteNumber(entry?.enhet ?? entry?.unit ?? entry?.unit_price);
    if (!Number.isFinite(unit)) return;
    if (type.includes('_external')) {
      external = Number.isFinite(external) ? external + unit : unit;
      return;
    }
    if (type.includes('_internal')) {
      internal = Number.isFinite(internal) ? internal + unit : unit;
      return;
    }
    direct = unit;
  });

  if (Number.isFinite(direct) && direct > 0) return direct;
  if (Number.isFinite(external) || Number.isFinite(internal)) {
    return round2((Number.isFinite(external) ? external : 0) + (Number.isFinite(internal) ? internal : 0));
  }
  return NaN;
}

async function loadFireBarrierPriceIndex() {
  const index = {};
  for (const filePath of OFFER_PRICE_SOURCE_FILES) {
    let raw;
    try {
      raw = await fs.readFile(filePath, 'utf8');
    } catch (err) {
      if (err && err.code === 'ENOENT') continue;
      console.warn(`[offer-template] Kunne ikke lese prisfil (${filePath}): ${err.message}`);
      continue;
    }

    const rows = raw.split(/\r?\n/).filter(Boolean);
    if (rows.length < 2) continue;
    const header = parseCsvLine(rows[0]).map(cell=>safeString(cell).toLowerCase());
    const descIdx = header.indexOf('desc_text');
    const priceIdx = header.indexOf('price');
    if (descIdx < 0 || priceIdx < 0) continue;

    for (let i = 1; i < rows.length; i += 1) {
      const columns = parseCsvLine(rows[i]);
      const desc = columns[descIdx] || '';
      if (!/fire\s*barri(?:er|e)\b|brann/i.test(desc)) continue;

      const amp = extractAmpFromFireBarrierDesc(desc);
      if (!Number.isFinite(amp)) continue;

      const price = toFiniteNumber(columns[priceIdx]);
      if (!Number.isFinite(price)) continue;

      const series = normalizeSeriesForFireLookup(desc);
      if (!series) continue;
      const type = detectFireBarrierType(desc);
      const ampKey = String(Math.round(amp));
      if (!index[series]) index[series] = {};
      if (!index[series][ampKey]) index[series][ampKey] = {};

      const previous = toFiniteNumber(index[series][ampKey][type]);
      if (!Number.isFinite(previous) || previous <= 0 || (price > 0 && price < previous)) {
        index[series][ampKey][type] = round2(price);
      }
    }
  }
  return index;
}

async function getFireBarrierPriceIndex() {
  if (!fireBarrierPriceIndexPromise) {
    fireBarrierPriceIndexPromise = loadFireBarrierPriceIndex().catch(err=>{
      fireBarrierPriceIndexPromise = null;
      throw err;
    });
  }
  return fireBarrierPriceIndexPromise;
}

function resolveFireBarrierUnitFromPriceIndex(priceIndex, series, amp) {
  const seriesKey = normalizeSeriesForFireLookup(series);
  const ampNum = toFiniteNumber(amp);
  if (!seriesKey || !Number.isFinite(ampNum)) return NaN;
  const ampKey = String(Math.round(ampNum));
  const row = priceIndex?.[seriesKey]?.[ampKey];
  if (!row || typeof row !== 'object') return NaN;

  const direct = toFiniteNumber(row.direct);
  if (Number.isFinite(direct) && direct > 0) return direct;

  const external = toFiniteNumber(row.external);
  const internal = toFiniteNumber(row.internal);
  if (Number.isFinite(external) || Number.isFinite(internal)) {
    return round2((Number.isFinite(external) ? external : 0) + (Number.isFinite(internal) ? internal : 0));
  }
  return NaN;
}

function resolveFireBarrierUnitPrice(line, priceIndex, input) {
  const fromBom = resolveFireBarrierUnitFromBom(line);
  if (Number.isFinite(fromBom) && fromBom > 0) return fromBom;
  const amp = toFiniteNumber(input?.ampere ?? input?.amp);
  const fromIndex = resolveFireBarrierUnitFromPriceIndex(priceIndex, input?.series, amp);
  if (Number.isFinite(fromIndex) && fromIndex > 0) return fromIndex;
  return NaN;
}

function resolveLineSelectedAddonTotal(line) {
  const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const baseTotal = toFiniteNumber(lineTotals.totalExMontasje);
  if (!Number.isFinite(baseTotal)) {
    const direct = toFiniteNumber(line?.selectedAddonTotal ?? line?.totals?.selectedAddonTotal);
    return Number.isFinite(direct) ? round2(direct) : NaN;
  }

  const selectedFlags = resolveLineSelectedAddonConfig(line, lineTotals);
  const includeMontasje = selectedFlags.includeMontasje;
  const includeEngineering = selectedFlags.includeEngineering;
  const includeOppheng = selectedFlags.includeOppheng;
  const montasjeTotal = toFiniteNumber(lineTotals.totalInclMontasje);
  const engineeringTotal = toFiniteNumber(lineTotals.totalInclEngineering);
  const opphengTotal = toFiniteNumber(lineTotals.totalInclOppheng ?? lineTotals.total);

  let total = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
  return round2(total);
}

function resolveLineMainVisibleTotal(line) {
  const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const baseTotal = toFiniteNumber(lineTotals.totalExMontasje);
  if (!Number.isFinite(baseTotal)) return NaN;

  const selectedFlags = resolveLineSelectedAddonConfig(line, lineTotals);
  const includeMontasje = selectedFlags.includeMontasje;
  const includeEngineering = selectedFlags.includeEngineering;
  const includeOppheng = selectedFlags.includeOppheng;
  const showMontasje = includeMontasje && selectedFlags.showMontasje;
  const showEngineering = includeEngineering && selectedFlags.showEngineering;
  const showOppheng = includeOppheng && selectedFlags.showOppheng;

  const montasjeTotal = toFiniteNumber(lineTotals.totalInclMontasje);
  const engineeringTotal = toFiniteNumber(lineTotals.totalInclEngineering);
  const opphengTotal = toFiniteNumber(lineTotals.totalInclOppheng ?? lineTotals.total);

  let total = baseTotal;
  if (includeMontasje && !showMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && !showEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && !showOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
  return round2(total);
}

function resolveLineOfferAmounts(line) {
  const includedTotal = resolveLineSelectedAddonTotal(line);
  const mainVisibleTotal = resolveLineMainVisibleTotal(line);
  let visibleAddonsTotal = NaN;
  if (Number.isFinite(includedTotal) && Number.isFinite(mainVisibleTotal)) {
    visibleAddonsTotal = round2(includedTotal - mainVisibleTotal);
  }
  return {
    includedTotal,
    mainVisibleTotal,
    visibleAddonsTotal
  };
}

function aggregateProjectOfferTotals(project) {
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  const totals = {
    material: 0,
    margin: 0,
    subtotal: 0,
    freight: 0,
    totalExMontasje: 0,
    montasje: 0,
    montasjeHours: 0,
    montasjeMargin: 0,
    totalInclMontasje: 0,
    engineering: 0,
    engineeringHours: 0,
    engineeringMargin: 0,
    totalInclEngineering: 0,
    oppheng: 0,
    opphengCount: 0,
    selectedAddonTotal: 0,
    offerIncludedTotal: 0,
    offerMainVisibleTotal: 0,
    offerVisibleAddonsTotal: 0
  };

  lines.forEach(line=>{
    const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
      ? line.totals
      : {};
    const add = (field, value)=>{
      const num = toFiniteNumber(value);
      if (Number.isFinite(num)) totals[field] += num;
    };
    add('material', lineTotals.material);
    add('margin', lineTotals.margin);
    add('subtotal', lineTotals.subtotal);
    add('freight', lineTotals.freight);
    add('totalExMontasje', lineTotals.totalExMontasje);
    add('montasje', lineTotals?.montasje?.cost);
    add('montasjeHours', lineTotals?.montasje?.totalHours);
    add('montasjeMargin', lineTotals.montasjeMargin);
    add('totalInclMontasje', lineTotals.totalInclMontasje);
    add('engineering', lineTotals?.engineering?.cost);
    add('engineeringHours', lineTotals?.engineering?.totalHours);
    add('engineeringMargin', lineTotals.engineeringMargin);
    add('totalInclEngineering', lineTotals.totalInclEngineering);
    add('oppheng', lineTotals?.oppheng?.cost ?? lineTotals.total);
    add('opphengCount', lineTotals?.oppheng?.pieceCount);
    const lineOfferAmounts = resolveLineOfferAmounts(line);
    add('selectedAddonTotal', lineOfferAmounts.includedTotal);
    add('offerIncludedTotal', lineOfferAmounts.includedTotal);
    add('offerMainVisibleTotal', lineOfferAmounts.mainVisibleTotal);
    add('offerVisibleAddonsTotal', lineOfferAmounts.visibleAddonsTotal);
  });

  return {
    lines,
    totals: Object.fromEntries(Object.entries(totals).map(([key, value])=>[key, round2(value)]))
  };
}

function normalizeElementLabel(rawValue) {
  const key = safeString(rawValue);
  const labels = {
    board_feed: 'Tavleelement',
    end_feed_unit: 'Endetilforselsboks',
    crt_board_feed: 'Trafoelement',
    end_cover: 'Endelokk',
    none: 'Ingen'
  };
  return labels[key] || key;
}

function collectProjectInputSummary(lines) {
  const pushUnique = (list, value)=>{
    const normalized = safeString(value);
    if (!normalized) return;
    if (!list.includes(normalized)) list.push(normalized);
  };

  const lineNumbers = [];
  const systems = [];
  const ampereValues = [];
  const ledereValues = [];
  const startElements = [];
  const sluttElements = [];
  let brannElementTotal = 0;
  let meterTotal = 0;
  let verticalAnglesTotal = 0;
  let horizontalAnglesTotal = 0;

  lines.forEach(line=>{
    const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};

    pushUnique(lineNumbers, line?.lineNumber);
    pushUnique(systems, input.series);
    pushUnique(ledereValues, input.ledere);
    pushUnique(startElements, normalizeElementLabel(input.startEl));
    pushUnique(sluttElements, normalizeElementLabel(input.sluttEl));

    const meter = toFiniteNumber(input.meter);
    if (Number.isFinite(meter)) meterTotal += meter;

    const verticalAngles = toFiniteNumber(input.v90_v ?? input.v90v);
    if (Number.isFinite(verticalAngles)) verticalAnglesTotal += verticalAngles;

    const horizontalAngles = toFiniteNumber(input.v90_h ?? input.v90h);
    if (Number.isFinite(horizontalAngles)) horizontalAnglesTotal += horizontalAngles;

    const brannQty = toFiniteNumber(input.fbQty ?? input.fireBarrierQty);
    if (Number.isFinite(brannQty)) brannElementTotal += brannQty;

    const ampNum = toFiniteNumber(input.ampere ?? input.amp);
    if (Number.isFinite(ampNum)) {
      pushUnique(ampereValues, String(Math.round(ampNum)));
    } else {
      pushUnique(ampereValues, input.ampere ?? input.amp);
    }
  });

  return {
    lineNumbers: lineNumbers.join(', '),
    systems: systems.join(', '),
    meterTotal: formatNoInteger(meterTotal),
    verticalAnglesTotal: formatNoInteger(verticalAnglesTotal),
    horizontalAnglesTotal: formatNoInteger(horizontalAnglesTotal),
    ampereValues: ampereValues.join(', '),
    ledereValues: ledereValues.join(', '),
    startElements: startElements.join(', '),
    sluttElements: sluttElements.join(', '),
    brannElementTotal: formatNoInteger(brannElementTotal)
  };
}

function buildOfferPlaceholderValues(project, offerNumber, offerDate, revision = 0) {
  const safeProject = (project && typeof project === 'object') ? project : {};
  const projectName = safeString(safeProject.name);
  const customer = safeString(safeProject.customer);
  const contactPerson = safeString(safeProject.contactPerson || safeProject.contact);
  const { lines, totals } = aggregateProjectOfferTotals(safeProject);
  const offerIncludedTotal = Number.isFinite(toFiniteNumber(totals.offerIncludedTotal))
    ? totals.offerIncludedTotal
    : (
      Number.isFinite(toFiniteNumber(totals.selectedAddonTotal))
        ? totals.selectedAddonTotal
        : totals.totalExMontasje
    );
  const offerMainVisibleTotal = Number.isFinite(toFiniteNumber(totals.offerMainVisibleTotal))
    ? totals.offerMainVisibleTotal
    : totals.totalExMontasje;
  const offerVisibleAddonsTotal = Number.isFinite(toFiniteNumber(totals.offerVisibleAddonsTotal))
    ? totals.offerVisibleAddonsTotal
    : (
      Number.isFinite(toFiniteNumber(offerIncludedTotal)) && Number.isFinite(toFiniteNumber(offerMainVisibleTotal))
        ? round2(offerIncludedTotal - offerMainVisibleTotal)
        : NaN
    );
  const inputSummary = collectProjectInputSummary(lines);
  const offerDatePlus30 = addDays(offerDate, 30);
  const revisionNumber = Number.isInteger(Number(revision)) ? Number(revision) : 0;

  const placeholders = {
    tilbud_nr: offerNumber,
    tilbudsdato: formatOfferDate(offerDate),
    dato: formatOfferDate(offerDate),
    dato30: formatOfferDate(offerDatePlus30),
    revisjon: String(revisionNumber),
    prosjektnavn: projectName,
    prosjekt: projectName,
    kunde: customer,
    customer: customer,
    kontaktperson: contactPerson,
    lss: '',
    linjer_start: '',
    lse: '',
    linjer_slutt: '',
    bss: '',
    bse: '',
    oss: '',
    ose: '',
    line_number: inputSummary.lineNumbers,
    linjenummer: inputSummary.lineNumbers,
    antall_linjer: String(lines.length),
    lnr: inputSummary.lineNumbers,
    sys: inputSummary.systems,
    mtr: inputSummary.meterTotal,
    vvk: inputSummary.verticalAnglesTotal,
    hvk: inputSummary.horizontalAnglesTotal,
    amp: inputSummary.ampereValues,
    led: inputSummary.ledereValues,
    ste: inputSummary.startElements,
    sle: inputSummary.sluttElements,
    bre: inputSummary.brannElementTotal,
    brt: '',
    brp: '',
    stv: formatNoCurrency(totals.selectedAddonTotal),
    tmo: formatNoCurrency(totals.totalInclMontasje),
    tin: formatNoCurrency(totals.totalInclEngineering),
    mtl: '',
    mtp: '',
    itl: '',
    itp: '',
    tol: '',
    top: '',
    tod: '',
    ttm: formatNoIntegerUp(totals.montasjeHours),
    tti: formatNoIntegerUp(totals.engineeringHours),
    aop: formatNoInteger(totals.opphengCount),
    timer_totalt_montasje: formatNoIntegerUp(totals.montasjeHours),
    timer_totalt_ingenior: formatNoIntegerUp(totals.engineeringHours),
    antall_oppheng: formatNoInteger(totals.opphengCount),
    material_nok: formatNoCurrency(totals.material),
    margin_nok: formatNoCurrency(totals.margin),
    subtotal_nok: formatNoCurrency(totals.subtotal),
    frakt_nok: formatNoCurrency(totals.freight),
    freight_nok: formatNoCurrency(totals.freight),
    total_ex_montasje_nok: formatNoCurrency(offerIncludedTotal),
    total_ex_montasje_hoved_nok: formatNoCurrency(offerMainVisibleTotal),
    total_ex_montasje_total_nok: formatNoCurrency(offerIncludedTotal),
    total_ex_montasje_tilvalg_nok: formatNoCurrency(offerVisibleAddonsTotal),
    offer_main_nok: formatNoCurrency(offerMainVisibleTotal),
    offer_total_nok: formatNoCurrency(offerIncludedTotal),
    offer_tilvalg_nok: formatNoCurrency(offerVisibleAddonsTotal),
    montasje_nok: formatNoCurrency(totals.montasje),
    montasje_margin_nok: formatNoCurrency(totals.montasjeMargin),
    total_incl_montasje_nok: formatNoCurrency(totals.totalInclMontasje),
    engineering_nok: formatNoCurrency(totals.engineering),
    engineering_margin_nok: formatNoCurrency(totals.engineeringMargin),
    total_incl_engineering_nok: formatNoCurrency(totals.totalInclEngineering),
    oppheng_nok: formatNoCurrency(totals.oppheng),
    selected_addon_total_nok: formatNoCurrency(offerIncludedTotal),
    total_valgte_nok: formatNoCurrency(offerIncludedTotal)
  };

  return placeholders;
}

async function buildOfferLinePlaceholderValues(project) {
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  const fireBarrierPriceIndex = await getFireBarrierPriceIndex().catch(err=>{
    console.warn(`[offer-template] Klarte ikke laste brannpriser: ${err.message}`);
    return {};
  });
  const linePlaceholderSets = [];
  const firePlaceholderSets = [];
  const opphengPlaceholderSets = [];

  lines.forEach((line, index)=>{
    const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};
    const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
      ? line.totals
      : {};
    const selectedAddonConfig = resolveLineSelectedAddonConfig(line, lineTotals);
    const lineOfferAmounts = resolveLineOfferAmounts(line);
    const showMontasje = selectedAddonConfig.includeMontasje && selectedAddonConfig.showMontasje;
    const showEngineering = selectedAddonConfig.includeEngineering && selectedAddonConfig.showEngineering;
    const showOppheng = selectedAddonConfig.includeOppheng && selectedAddonConfig.showOppheng;

    const ampNum = toFiniteNumber(input.ampere ?? input.amp);
    const amp = Number.isFinite(ampNum)
      ? String(Math.round(ampNum))
      : safeString(input.ampere ?? input.amp);
    const lineNumber = safeString(line?.lineNumber || String(index + 1));
    const brannQtyNum = toFiniteNumber(input.fbQty ?? input.fireBarrierQty);
    const brannQty = Number.isFinite(brannQtyNum) ? brannQtyNum : 0;
    const hasBrannElements = brannQty > 0;

    const montasjePrice = formatNoCurrencyWithKr(lineTotals.totalInclMontasje);
    const engineeringPrice = formatNoCurrencyWithKr(lineTotals.totalInclEngineering);
    const opphengCost = lineTotals?.oppheng?.cost ?? lineTotals.totalInclOppheng ?? lineTotals.total;
    const opphengPrice = formatNoCurrencyWithKr(opphengCost);
    const opphengCount = formatNoInteger(lineTotals?.oppheng?.pieceCount);
    const opphengDetail = opphengCount ? `- ${opphengCount} stk. Oppheng` : '';
    const montasjeHoursValue = formatNoIntegerUp(lineTotals?.montasje?.totalHours);
    const engineeringHoursValue = formatNoIntegerUp(lineTotals?.engineering?.totalHours);
    const montasjeHoursLabel = montasjeHoursValue ? `${montasjeHoursValue} timer totalt` : '';
    const engineeringHoursLabel = engineeringHoursValue ? `${engineeringHoursValue} timer totalt` : '';

    linePlaceholderSets.push({
      lnr: lineNumber,
      linjenummer: lineNumber,
      sys: safeString(input.series),
      mtr: formatNoInteger(input.meter),
      vvk: formatNoInteger(input.v90_v ?? input.v90v),
      hvk: formatNoInteger(input.v90_h ?? input.v90h),
      amp,
      led: safeString(input.ledere),
      ste: normalizeElementLabel(input.startEl),
      sle: normalizeElementLabel(input.sluttEl),
      avb: formatNoInteger(input.boxQty),
      bre: hasBrannElements
        ? `${formatNoInteger(brannQty)} stk. Branngjennomforing EI 60/90/120`
        : '',
      total_ex_montasje_nok: formatNoCurrency(lineOfferAmounts.mainVisibleTotal),
      stv: formatNoCurrency(lineOfferAmounts.mainVisibleTotal),
      stv_hoved: formatNoCurrency(lineOfferAmounts.mainVisibleTotal),
      stv_total: formatNoCurrency(lineOfferAmounts.includedTotal),
      stv_tilvalg: formatNoCurrency(lineOfferAmounts.visibleAddonsTotal),
      line_main_nok: formatNoCurrency(lineOfferAmounts.mainVisibleTotal),
      line_total_nok: formatNoCurrency(lineOfferAmounts.includedTotal),
      line_tilvalg_nok: formatNoCurrency(lineOfferAmounts.visibleAddonsTotal),
      tmo: formatNoCurrency(lineTotals.totalInclMontasje),
      tin: formatNoCurrency(lineTotals.totalInclEngineering),
      mtl: showMontasje ? 'Montasje' : '',
      mtp: showMontasje ? montasjePrice : '',
      ttm: showMontasje ? montasjeHoursLabel : '',
      itl: showEngineering ? 'Ingenior' : '',
      itp: showEngineering ? engineeringPrice : '',
      tti: showEngineering ? engineeringHoursLabel : '',
      tol: showOppheng ? 'Opphengsmateriell' : '',
      top: showOppheng ? opphengPrice : '',
      tod: showOppheng ? opphengDetail : '',
      aop: formatNoInteger(lineTotals?.oppheng?.pieceCount),
      timer_totalt_montasje: showMontasje ? montasjeHoursValue : '',
      timer_totalt_ingenior: showEngineering ? engineeringHoursValue : '',
      antall_oppheng: formatNoInteger(lineTotals?.oppheng?.pieceCount),
      brt: '',
      brp: '',
      bss: '',
      bse: '',
      oss: '',
      ose: ''
    });

    if (!selectedAddonConfig.includeOppheng && (opphengPrice || opphengDetail)) {
      opphengPlaceholderSets.push({
        lnr: lineNumber,
        tol: 'Opphengsmateriell',
        top: opphengPrice,
        tod: opphengDetail
      });
    }

    if (!hasBrannElements) {
      const fireUnitPrice = resolveFireBarrierUnitPrice(line, fireBarrierPriceIndex, input);
      const fireOfferPrice = Number.isFinite(fireUnitPrice) ? round2(fireUnitPrice / 0.8) : NaN;
      const fireAmpSuffix = amp ? ` - ${amp}A` : '';
      firePlaceholderSets.push({
        lnr: lineNumber,
        brt: `Branngjennomforing EI 60/90/120${fireAmpSuffix}`,
        brp: formatNoCurrency(fireOfferPrice)
      });
    }
  });

  return {
    linePlaceholderSets,
    firePlaceholderSets,
    opphengPlaceholderSets
  };
}

function replacePlaceholdersInXml(xml, placeholders) {
  const replaceDelimiterSplitPlaceholder = (input, key, escapedValue)=>{
    const escapedKey = escapeRegex(key);
    const nextTextNode = '((?:(?!<w:t)[\\s\\S])*?<w:t[^>]*>)';

    // Case 1: node1 has "...{{", node2 has "...key...", node3 has "}}..."
    const threeNodePattern = new RegExp(
      `(<w:t[^>]*>)([^<]*?)\\{\\{\\s*(</w:t>)${nextTextNode}([^<]*?)${escapedKey}([^<]*?)(</w:t>)${nextTextNode}\\s*\\}\\}([^<]*?)(</w:t>)`,
      'g'
    );

    // Case 2: node1 has "...{{", node2 has "...key}}..."
    const twoNodePatternA = new RegExp(
      `(<w:t[^>]*>)([^<]*?)\\{\\{\\s*(</w:t>)${nextTextNode}([^<]*?)${escapedKey}([^<]*?)\\s*\\}\\}([^<]*?)(</w:t>)`,
      'g'
    );

    // Case 3: node1 has "...{{key", node2 has "}}..."
    const twoNodePatternB = new RegExp(
      `(<w:t[^>]*>)([^<]*?)\\{\\{\\s*${escapedKey}([^<]*?)(</w:t>)${nextTextNode}\\s*\\}\\}([^<]*?)(</w:t>)`,
      'g'
    );

    // Keep text around placeholder and only replace the placeholder fragment itself.
    let output = input.replace(threeNodePattern, `$1$2${escapedValue}$3$4$5$6$7$8$9$10`);
    output = output.replace(twoNodePatternA, `$1$2${escapedValue}$3$4$5$6$7$8`);
    output = output.replace(twoNodePatternB, `$1$2${escapedValue}$3$4$5$6$7`);
    return output;
  };

  const replaceSplitThreeLetterPlaceholder = (input, key, escapedValue)=>{
    if (String(key).length !== 3) return input;
    const [a, b, c] = key.split('');
    const nextTextNode = '((?:(?!<w:t)[\\s\\S])*?<w:t[^>]*>)';

    // Case 1: {{s | t | e}} split over three separate w:t nodes.
    const threeNodePattern = new RegExp(
      `(<w:t[^>]*>)\\{\\{\\s*${escapeRegex(a)}\\s*(</w:t>)${nextTextNode}\\s*${escapeRegex(b)}\\s*(</w:t>)${nextTextNode}\\s*${escapeRegex(c)}\\s*\\}\\}(</w:t>)`,
      'g'
    );

    // Case 2: {{s | te}} split over two nodes.
    const twoNodePatternA = new RegExp(
      `(<w:t[^>]*>)\\{\\{\\s*${escapeRegex(a)}\\s*(</w:t>)${nextTextNode}\\s*${escapeRegex(b)}\\s*${escapeRegex(c)}\\s*\\}\\}(</w:t>)`,
      'g'
    );

    // Case 3: {{st | e}} split over two nodes.
    const twoNodePatternB = new RegExp(
      `(<w:t[^>]*>)\\{\\{\\s*${escapeRegex(a)}\\s*${escapeRegex(b)}\\s*(</w:t>)${nextTextNode}\\s*${escapeRegex(c)}\\s*\\}\\}(</w:t>)`,
      'g'
    );

    // Keep XML structure intact by writing value into first text node and emptying the rest.
    let output = input.replace(threeNodePattern, `$1${escapedValue}$2$3$4$5$6$7`);
    output = output.replace(twoNodePatternA, `$1${escapedValue}$2$3$4`);
    output = output.replace(twoNodePatternB, `$1${escapedValue}$2$3$4`);
    return output;
  };

  let output = String(xml);
  Object.entries(placeholders).forEach(([key, rawValue])=>{
    const escapedValue = escapeXml(rawValue ?? '');
    const pattern = new RegExp(`\\{\\{\\s*${escapeRegex(key)}\\s*\\}\\}`, 'g');
    output = output.replace(pattern, escapedValue);
    output = replaceDelimiterSplitPlaceholder(output, key, escapedValue);
    output = replaceSplitThreeLetterPlaceholder(output, key, escapedValue);
  });
  return output;
}

function expandRepeatBlock(xml, options) {
  const {
    startAliases,
    endAliases,
    startToken,
    endToken,
    placeholderSets
  } = options;
  const normalizedStartAliases = Array.isArray(startAliases) ? startAliases : [];
  const normalizedEndAliases = Array.isArray(endAliases) ? endAliases : [];
  if (!normalizedStartAliases.length || !normalizedEndAliases.length) return String(xml);

  const markerPlaceholders = {};
  normalizedStartAliases.forEach(alias=>{
    markerPlaceholders[alias] = startToken;
  });
  normalizedEndAliases.forEach(alias=>{
    markerPlaceholders[alias] = endToken;
  });
  const markerReplaced = replacePlaceholdersInXml(xml, markerPlaceholders);

  const tokenPattern = new RegExp(
    `${escapeRegex(startToken)}([\\s\\S]*?)${escapeRegex(endToken)}`,
    'g'
  );
  const rawStart = normalizedStartAliases.map(alias=>escapeRegex(alias)).join('|');
  const rawEnd = normalizedEndAliases.map(alias=>escapeRegex(alias)).join('|');
  const rawPattern = new RegExp(
    `\\{\\{\\s*(?:${rawStart})\\s*\\}\\}([\\s\\S]*?)\\{\\{\\s*(?:${rawEnd})\\s*\\}\\}`,
    'g'
  );

  const renderBlock = (blockTemplate)=>{
    if (!Array.isArray(placeholderSets) || placeholderSets.length === 0) return '';
    return placeholderSets
      .map(placeholders=>replacePlaceholdersInXml(blockTemplate, placeholders))
      .join('');
  };

  let expanded = markerReplaced.replace(tokenPattern, (_match, blockTemplate)=>renderBlock(blockTemplate));
  expanded = expanded.replace(rawPattern, (_match, blockTemplate)=>renderBlock(blockTemplate));

  const clearPlaceholders = {};
  normalizedStartAliases.forEach(alias=>{
    clearPlaceholders[alias] = '';
  });
  normalizedEndAliases.forEach(alias=>{
    clearPlaceholders[alias] = '';
  });
  expanded = replacePlaceholdersInXml(expanded, clearPlaceholders);

  const markerRegex = new RegExp(
    `\\{\\{\\s*(?:${rawStart}|${rawEnd})\\s*\\}\\}`,
    'g'
  );
  return expanded
    .split(startToken).join('')
    .split(endToken).join('')
    .replace(markerRegex, '');
}

function expandLineRepeatBlocks(xml, linePlaceholderSets) {
  return expandRepeatBlock(xml, {
    startAliases: ['lss', 'linjer_start'],
    endAliases: ['lse', 'linjer_slutt'],
    startToken: OFFER_LINE_BLOCK_START_TOKEN,
    endToken: OFFER_LINE_BLOCK_END_TOKEN,
    placeholderSets: linePlaceholderSets
  });
}

function expandFireRepeatBlocks(xml, firePlaceholderSets) {
  return expandRepeatBlock(xml, {
    startAliases: ['bss'],
    endAliases: ['bse'],
    startToken: OFFER_FIRE_BLOCK_START_TOKEN,
    endToken: OFFER_FIRE_BLOCK_END_TOKEN,
    placeholderSets: firePlaceholderSets
  });
}

function expandOpphengRepeatBlocks(xml, opphengPlaceholderSets) {
  return expandRepeatBlock(xml, {
    startAliases: ['oss'],
    endAliases: ['ose'],
    startToken: OFFER_OPPHENG_BLOCK_START_TOKEN,
    endToken: OFFER_OPPHENG_BLOCK_END_TOKEN,
    placeholderSets: opphengPlaceholderSets
  });
}

async function generateOfferDocxBuffer(project, offerNumber, offerDate, revision = 0) {
  await fs.access(OFFER_TEMPLATE_FILE);
  const placeholders = buildOfferPlaceholderValues(project, offerNumber, offerDate, revision);
  const {
    linePlaceholderSets,
    firePlaceholderSets,
    opphengPlaceholderSets
  } = await buildOfferLinePlaceholderValues(project);
  const zip = new AdmZip(OFFER_TEMPLATE_FILE);
  const entries = zip.getEntries().filter(entry=>
    !entry.isDirectory &&
    entry.entryName.startsWith('word/') &&
    entry.entryName.endsWith('.xml')
  );

  entries.forEach(entry=>{
    const xml = entry.getData().toString('utf8');
    const withExpandedLineBlocks = expandLineRepeatBlocks(xml, linePlaceholderSets);
    const withExpandedFireBlocks = expandFireRepeatBlocks(withExpandedLineBlocks, firePlaceholderSets);
    const withExpandedOpphengBlocks = expandOpphengRepeatBlocks(withExpandedFireBlocks, opphengPlaceholderSets);
    const replaced = replacePlaceholdersInXml(withExpandedOpphengBlocks, placeholders);
    if (replaced !== xml) {
      zip.updateFile(entry.entryName, Buffer.from(replaced, 'utf8'));
    }
  });

  return zip.toBuffer();
}

async function fetchWithTimeout(url, options = {}) {
  const {
    timeout = MARKET_HTTP_TIMEOUT_MS,
    headers = {},
    accept = 'application/json',
    ...rest
  } = options;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    return await fetch(url, {
      ...rest,
      headers: {
        'User-Agent': MARKET_USER_AGENT,
        Accept: accept,
        ...headers
      },
      signal: controller.signal
    });
  } finally {
    clearTimeout(timer);
  }
}

function isObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function parseRate(value) {
  const parsed = Number(String(value ?? '').trim().replace(',', '.'));
  if (!Number.isFinite(parsed)) return NaN;
  return parsed;
}

function normalizeIsoDate(value) {
  if (typeof value !== 'string') return '';
  const trimmed = value.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return '';
  const parsed = Date.parse(`${trimmed}T00:00:00Z`);
  if (Number.isNaN(parsed)) return '';
  return trimmed;
}

function normalizeIsoTimestamp(value) {
  if (typeof value !== 'string') return '';
  const parsed = Date.parse(value);
  if (Number.isNaN(parsed)) return '';
  return new Date(parsed).toISOString();
}

function extractLatestObservationFromSdmx(payload, pair) {
  const timeValues = payload?.data?.structure?.dimensions?.observation?.[0]?.values;
  const series = payload?.data?.dataSets?.[0]?.series;

  if (!Array.isArray(timeValues) || !isObject(series)) {
    throw new Error(`SDMX-respons for ${pair} mangler nodene som trengs`);
  }

  let latestIndex = -1;
  let latestRate = NaN;

  for (const seriesEntry of Object.values(series)) {
    const observations = seriesEntry?.observations;
    if (!isObject(observations)) continue;

    for (const [indexKey, rawValue] of Object.entries(observations)) {
      const idx = Number(indexKey);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const rawRate = Array.isArray(rawValue) ? rawValue[0] : rawValue;
      const parsedRate = parseRate(rawRate);
      if (idx > latestIndex && Number.isFinite(parsedRate)) {
        latestIndex = idx;
        latestRate = parsedRate;
      }
    }
  }

  if (latestIndex < 0 || !Number.isFinite(latestRate)) {
    throw new Error(`Fant ikke gyldig siste datapunkt for ${pair}`);
  }

  const date = normalizeIsoDate(timeValues?.[latestIndex]?.id);
  if (!date) {
    throw new Error(`Siste datapunkt for ${pair} mangler gyldig dato`);
  }

  return { rate: latestRate, date };
}

async function fetchNorgesBankRate(url, pair) {
  const res = await fetchWithTimeout(url, { accept: 'application/json' });
  if (!res.ok) {
    throw new Error(`Norges Bank ${pair} svarte ${res.status}`);
  }

  const body = await res.json();
  return extractLatestObservationFromSdmx(body, pair);
}

async function fetchFxRatesFromNorgesBank() {
  const [usd, eur] = await Promise.all([
    fetchNorgesBankRate(NORGES_BANK_USD_NOK_URL, 'USD/NOK'),
    fetchNorgesBankRate(NORGES_BANK_EUR_NOK_URL, 'EUR/NOK')
  ]);

  return {
    usdNok: {
      pair: 'USD/NOK',
      rate: usd.rate,
      date: usd.date,
      source: 'Norges Bank'
    },
    eurNok: {
      pair: 'EUR/NOK',
      rate: eur.rate,
      date: eur.date,
      source: 'Norges Bank'
    },
    source: 'Norges Bank'
  };
}

async function fetchLmeQuote() {
  const res = await fetchWithTimeout(MARKET_LME_URL, { accept: 'application/json' });
  if (!res.ok) {
    throw new Error(`LME price-endepunkt svarte ${res.status}`);
  }
  const data = await res.json();
  const quote = data?.quoteResponse?.result?.[0] || {};
  const price = Number(quote.regularMarketPrice);
  return {
    price: Number.isFinite(price) ? price : null,
    currency: quote.currency || 'USD',
    symbol: quote.symbol || 'ALI=F',
    source: 'Yahoo Finance',
    notation: `${quote.currency || 'USD'}/t`,
    unit: 't'
  };
}

function normalizeFxPoint(rawPoint, pair, fallbackSource) {
  if (isObject(rawPoint)) {
    const rate = parseRate(rawPoint.rate ?? rawPoint.value ?? rawPoint.last);
    const date = normalizeIsoDate(rawPoint.date ?? rawPoint.time ?? rawPoint.valueDate);
    if (Number.isFinite(rate) && rate > 0) {
      return {
        pair,
        rate,
        date,
        source: String(rawPoint.source || fallbackSource || 'Norges Bank')
      };
    }
  }

  const directRate = parseRate(rawPoint);
  if (Number.isFinite(directRate) && directRate > 0) {
    return {
      pair,
      rate: directRate,
      date: '',
      source: String(fallbackSource || 'Norges Bank')
    };
  }

  return {};
}

function normalizeFxSnapshot(rawFx) {
  if (!isObject(rawFx)) {
    return {
      usdNok: {},
      eurNok: {},
      source: 'Norges Bank'
    };
  }

  const fallbackSource = String(rawFx.source || 'Norges Bank');
  const usdRaw = rawFx.usdNokPoint ?? rawFx.usdNok;
  const eurRaw = rawFx.eurNokPoint ?? rawFx.eurNok;
  const usd = normalizeFxPoint(usdRaw, 'USD/NOK', fallbackSource);
  const eur = normalizeFxPoint(eurRaw, 'EUR/NOK', fallbackSource);

  if (Object.keys(usd).length === 0) {
    const fromLegacy = normalizeIsoDate(rawFx.usdNokDate);
    if (fromLegacy) usd.date = fromLegacy;
  }
  if (Object.keys(eur).length === 0) {
    const fromLegacy = normalizeIsoDate(rawFx.eurNokDate);
    if (fromLegacy) eur.date = fromLegacy;
  }

  return {
    usdNok: usd,
    eurNok: eur,
    source: fallbackSource
  };
}

function normalizeMarketPayload(rawPayload) {
  if (!isObject(rawPayload)) {
    return {
      aluminium: {},
      fx: normalizeFxSnapshot({}),
      updatedAt: new Date().toISOString()
    };
  }

  const updatedAt =
    normalizeIsoTimestamp(rawPayload.updatedAt) ||
    normalizeIsoTimestamp(rawPayload.fetchedAt) ||
    new Date().toISOString();

  return {
    aluminium: isObject(rawPayload.aluminium) ? rawPayload.aluminium : {},
    fx: normalizeFxSnapshot(rawPayload.fx),
    updatedAt
  };
}

function toDateKeyInTimezone(dateValue, timeZone = MARKET_TIMEZONE) {
  const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (Number.isNaN(date.getTime())) return '';
  try {
    return new Intl.DateTimeFormat('en-CA', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      timeZone
    }).format(date);
  } catch (_err) {
    return date.toISOString().slice(0, 10);
  }
}

function isMarketDataStaleForToday(now = new Date()) {
  const lastSuccessIso = normalizeIsoTimestamp(
    marketScheduleState.lastSuccessAt || marketCache.payload?.updatedAt || ''
  );
  if (!lastSuccessIso) return true;
  const todayKey = toDateKeyInTimezone(now, MARKET_TIMEZONE);
  const lastKey = toDateKeyInTimezone(lastSuccessIso, MARKET_TIMEZONE);
  if (!todayKey || !lastKey) return true;
  return todayKey !== lastKey;
}

function buildScheduleMeta() {
  return {
    type: 'daily',
    refreshAtLocal: MARKET_DAILY_REFRESH_LABEL,
    retryDelayMinutes: MARKET_RETRY_DELAY_MINUTES,
    timezone: MARKET_TIMEZONE,
    status: marketScheduleState.status,
    lastRunAt: marketScheduleState.lastRunAt,
    lastAttemptAt: marketScheduleState.lastAttemptAt,
    lastSuccessAt: marketScheduleState.lastSuccessAt,
    nextRunAt: marketScheduleState.nextRunAt,
    lastError: marketScheduleState.lastError,
    retryCount: marketScheduleState.retryCount
  };
}

function withSchedule(payload) {
  return {
    ...payload,
    mode: 'auto-daily',
    schedule: buildScheduleMeta()
  };
}

async function readMarketPayloadFromFile(filePath) {
  try {
    const raw = await fs.readFile(filePath, 'utf8');
    const parsed = JSON.parse(raw);
    return normalizeMarketPayload(parsed);
  } catch (err) {
    if (err && err.code === 'ENOENT') {
      return null;
    }
    console.warn(`[market-data] Kunne ikke lese datafil (${filePath}): ${err.message}`);
    return null;
  }
}

async function writeMarketPayloadToFile(filePath, payload) {
  try {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    const persisted = {
      aluminium: payload.aluminium || {},
      fx: payload.fx || {},
      updatedAt: payload.updatedAt || new Date().toISOString(),
      mode: 'auto-daily'
    };
    await fs.writeFile(filePath, JSON.stringify(persisted, null, 2), 'utf8');
  } catch (err) {
    console.warn(`[market-data] Kunne ikke lagre datafil (${filePath}): ${err.message}`);
  }
}

function computeNextDailyRefresh(from = new Date()) {
  const next = new Date(from);
  next.setHours(MARKET_DAILY_REFRESH_HOUR, 0, 0, 0);
  if (next.getTime() <= from.getTime()) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

async function refreshMarketDataNow(reason = 'manual') {
  const base = marketCache.payload || (await readMarketPayloadFromFile(MARKET_DATA_FILE)) || {
    aluminium: {},
    fx: normalizeFxSnapshot({}),
    updatedAt: new Date().toISOString()
  };

  const [fxResult, lmeResult] = await Promise.allSettled([
    fetchFxRatesFromNorgesBank(),
    fetchLmeQuote()
  ]);

  if (fxResult.status !== 'fulfilled') {
    const wrapped = new Error(`Kunne ikke hente valutakurser fra Norges Bank (${reason})`);
    wrapped.cause = fxResult.reason;
    throw wrapped;
  }

  const nextPayload = normalizeMarketPayload({
    aluminium: lmeResult.status === 'fulfilled' ? lmeResult.value : base.aluminium,
    fx: fxResult.value,
    updatedAt: new Date().toISOString()
  });

  marketCache.payload = nextPayload;
  marketScheduleState.lastRunAt = nextPayload.updatedAt;
  marketScheduleState.lastAttemptAt = nextPayload.updatedAt;
  marketScheduleState.lastSuccessAt = nextPayload.updatedAt;
  marketScheduleState.lastError = null;
  marketScheduleState.status = 'ok';
  marketScheduleState.retryCount = 0;

  await writeMarketPayloadToFile(MARKET_DATA_FILE, nextPayload);
  return nextPayload;
}

async function runScheduledMarketRefresh(reason) {
  if (marketRefreshInFlight) {
    return marketRefreshInFlight;
  }

  marketScheduleState.lastAttemptAt = new Date().toISOString();
  marketRefreshInFlight = (async ()=>{
    try {
      await refreshMarketDataNow(reason);
      console.log(`[market-data] Oppdatert automatisk (${reason})`);
      return true;
    } catch (err) {
      const details = err?.message || String(err);
      marketScheduleState.lastError = `${new Date().toISOString()} ${details}`;
      marketScheduleState.status = 'error';
      console.error(`[market-data] Automatisk oppdatering feilet (${reason})`, err);
      return false;
    } finally {
      marketRefreshInFlight = null;
    }
  })();

  return marketRefreshInFlight;
}

function scheduleMarketRefreshAt(targetDate, reason) {
  if (marketScheduleState.timerId) {
    clearTimeout(marketScheduleState.timerId);
    marketScheduleState.timerId = null;
  }

  marketScheduleState.nextRunAt = targetDate.toISOString();
  const delayMs = Math.max(1000, targetDate.getTime() - Date.now());

  marketScheduleState.timerId = setTimeout(async ()=>{
    await handleScheduledMarketRefresh(reason);
  }, delayMs);
}

function scheduleDailyMarketRefresh() {
  const nextRun = computeNextDailyRefresh();
  scheduleMarketRefreshAt(nextRun, `daily-${MARKET_DAILY_REFRESH_LABEL}`);
}

function scheduleRetryMarketRefresh() {
  const nextRun = new Date(Date.now() + MARKET_RETRY_DELAY_MINUTES * 60 * 1000);
  const retryNumber = Math.max(1, marketScheduleState.retryCount);
  scheduleMarketRefreshAt(nextRun, `retry-${retryNumber}`);
}

async function handleScheduledMarketRefresh(reason) {
  const succeeded = await runScheduledMarketRefresh(reason);
  if (succeeded) {
    scheduleDailyMarketRefresh();
    return;
  }

  marketScheduleState.retryCount += 1;
  scheduleRetryMarketRefresh();
}

async function refreshMarketDataIfStale(reason = 'on-demand') {
  if (!isMarketDataStaleForToday(new Date())) {
    return false;
  }

  console.log('[market-data] Data er eldre enn dagens dato, trigget automatisk oppdatering');
  await handleScheduledMarketRefresh(reason);
  return true;
}

async function ensureMarketDataLoaded() {
  if (marketCache.payload) {
    return marketCache.payload;
  }

  const fromFile = await readMarketPayloadFromFile(MARKET_DATA_FILE);
  if (fromFile) {
    marketCache.payload = fromFile;
    const fromFileUpdatedAt = normalizeIsoTimestamp(fromFile.updatedAt);
    if (fromFileUpdatedAt) {
      marketScheduleState.lastRunAt = fromFileUpdatedAt;
      marketScheduleState.lastSuccessAt = fromFileUpdatedAt;
      marketScheduleState.status = 'ok';
    }
    return marketCache.payload;
  }

  return refreshMarketDataNow('bootstrap');
}

async function initializeMarketDataAutomation() {
  await ensureMarketDataLoaded();
  const startupSucceeded = await runScheduledMarketRefresh('startup');
  if (startupSucceeded) {
    scheduleDailyMarketRefresh();
  } else {
    marketScheduleState.retryCount += 1;
    scheduleRetryMarketRefresh();
  }

  console.log(
    `[market-data] Automatisk valutaoppdatering aktivert: daglig ${MARKET_DAILY_REFRESH_LABEL} (${MARKET_TIMEZONE}), retry ${MARKET_RETRY_DELAY_MINUTES} min ved feil`
  );
}

function currentMarketPayloadForResponse() {
  if (!marketCache.payload) {
    return withSchedule(
      normalizeMarketPayload({
        aluminium: {},
        fx: {},
        updatedAt: new Date().toISOString()
      })
    );
  }
  return withSchedule(marketCache.payload);
}

function stopMarketScheduler() {
  if (marketScheduleState.timerId) {
    clearTimeout(marketScheduleState.timerId);
    marketScheduleState.timerId = null;
  }
}

const credential = new ClientSecretCredential(
  process.env.OAUTH_TENANT_ID,
  process.env.OAUTH_CLIENT_ID,
  process.env.OAUTH_CLIENT_SECRET
);

async function getAccessToken() {
  const scope = 'https://outlook.office365.com/.default';
  const token = await credential.getToken(scope);
  return token?.token;
}

async function sendMail({ subject, html }) {
  const accessToken = await getAccessToken();

  if (!accessToken) {
    throw new Error('Kunne ikke hente OAuth2 access token');
  }

  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false,
    auth: {
      type: 'OAuth2',
      user: process.env.SMTP_USER,
      clientId: process.env.OAUTH_CLIENT_ID,
      clientSecret: process.env.OAUTH_CLIENT_SECRET,
      tenantId: process.env.OAUTH_TENANT_ID,
      accessToken
    }
  });

  return transporter.sendMail({
    from: process.env.SMTP_USER,
    to: process.env.MAIL_TO,
    subject,
    html
  });
}

app.get('/api/health', async (_req, res) => {
  const now = new Date().toISOString();
  return res.json({
    ok: true,
    service: 'busbar-api',
    time: now,
    runtimeDataDir: RUNTIME_DATA_DIR
  });
});

app.get('/api/market-data', async (req, res) => {
  try {
    await ensureMarketDataLoaded();
    await refreshMarketDataIfStale('on-demand-stale');
    res.json(currentMarketPayloadForResponse());
  } catch (err) {
    console.error('Markedsdata feilet', err);
    res.status(502).json({ error: 'Kunne ikke hente markedsdata' });
  }
});

app.get('/api/user-projects', async (req, res) => {
  try {
    const email = normalizeEmail(req.query?.email);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }
    const archive = await readProjectArchive();
    const userRecord = archive.users[email] || {
      email,
      updatedAt: null,
      projects: []
    };
    return res.json({
      email,
      updatedAt: userRecord.updatedAt,
      projects: userRecord.projects
    });
  } catch (err) {
    console.error('Henting av brukerprosjekter feilet', err);
    return res.status(500).json({ error: 'Kunne ikke hente prosjekter' });
  }
});

app.post('/api/user-projects/sync', async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }
    if (!Array.isArray(req.body?.projects)) {
      return res.status(400).json({ error: 'Mangler prosjekter' });
    }
    if (req.body.projects.length > 2000) {
      return res.status(413).json({ error: 'For mange prosjekter i én synk' });
    }

    const normalizedProjects = req.body.projects
      .map(normalizeProjectRecord)
      .filter(Boolean);
    const updatedAt = new Date().toISOString();

    const nextUserRecord = await withProjectArchiveLock(async () => {
      const archive = await readProjectArchive();
      archive.users[email] = {
        email,
        updatedAt,
        projects: normalizedProjects
      };
      await writeProjectArchive(archive);
      return archive.users[email];
    });

    return res.json({
      email: nextUserRecord.email,
      updatedAt: nextUserRecord.updatedAt,
      projects: nextUserRecord.projects
    });
  } catch (err) {
    console.error('Synk av brukerprosjekter feilet', err);
    return res.status(500).json({ error: 'Kunne ikke synkronisere prosjekter' });
  }
});

app.get('/api/admin/project-overview', requireAdminAuth, async (req, res) => {
  try {
    const archive = await readProjectArchive();
    const users = Object.values(archive.users).map(user => {
      const projects = Array.isArray(user.projects) ? user.projects : [];
      const projectsWithCounts = projects.map(project => ({
        ...project,
        lineCount: Array.isArray(project.lines) ? project.lines.length : 0
      }));
      const lineCount = projectsWithCounts.reduce((sum, project) => {
        return sum + Number(project.lineCount || 0);
      }, 0);
      return {
        email: user.email,
        updatedAt: user.updatedAt,
        projectCount: projectsWithCounts.length,
        lineCount,
        projects: projectsWithCounts
      };
    });

    users.sort((a, b) => {
      const aTime = new Date(a.updatedAt || 0).getTime();
      const bTime = new Date(b.updatedAt || 0).getTime();
      return bTime - aTime;
    });

    const totals = users.reduce((acc, user) => {
      acc.userCount += 1;
      acc.projectCount += Number(user.projectCount || 0);
      acc.lineCount += Number(user.lineCount || 0);
      return acc;
    }, { userCount: 0, projectCount: 0, lineCount: 0 });

    return res.json({
      generatedAt: new Date().toISOString(),
      totals,
      users
    });
  } catch (err) {
    console.error('Henting av admin-oversikt feilet', err);
    return res.status(500).json({ error: 'Kunne ikke hente admin-oversikt' });
  }
});

app.post('/api/send-calculation-email', async (req, res) => {
  try {
    const { project, customer, totals, bom } = req.body || {};

    const subject = `Ny beregning: ${project || 'Uten prosjektnavn'}`;

    const htmlRows = (Array.isArray(bom) ? bom : []).map(item => `
      <tr>
        <td>${item.code || ''}</td>
        <td>${item.type || ''}</td>
        <td>${item.series || ''}</td>
        <td>${item.ampere || ''}</td>
        <td>${item.ledere || item.lederes || ''}</td>
        <td>${item.antall || ''}</td>
        <td>${item.enhet || ''}</td>
        <td>${item.sum || ''}</td>
      </tr>
    `).join('');

    const html = `
      <h1>Ny beregning</h1>
      <p><strong>Prosjekt:</strong> ${project || '-'}</p>
      <p><strong>Kunde:</strong> ${customer || '-'}</p>
      <p><strong>Total:</strong> ${totals?.total ?? '-'}</p>
      <p><strong>Material:</strong> ${totals?.material ?? '-'}</p>
      <p><strong>DG:</strong> ${totals?.margin ?? '-'}</p>
      <p><strong>Frakt:</strong> ${totals?.freight ?? '-'}</p>
      <p><strong>Subtotal:</strong> ${totals?.subtotal ?? '-'}</p>
      <h2>Materialliste</h2>
      <table border="1" cellpadding="4" cellspacing="0">
        <thead>
          <tr>
            <th>Code</th>
            <th>Type</th>
            <th>Serie</th>
            <th>Amp</th>
            <th>Ledere</th>
            <th>Antall</th>
            <th>Enhet</th>
            <th>Sum</th>
          </tr>
        </thead>
        <tbody>
          ${htmlRows || '<tr><td colspan="8">Ingen BOM</td></tr>'}
        </tbody>
      </table>
    `;

    await sendMail({ subject, html });
    res.status(204).end();
  } catch (err) {
    console.error('Send mail feilet', err);
    res.status(500).json({ error: 'Kunne ikke sende e-post' });
  }
});

app.post('/api/generate-offer', async (req, res) => {
  try {
    const project = (req.body && typeof req.body === 'object') ? req.body.project : null;
    if (!project || typeof project !== 'object') {
      return res.status(400).json({ error: 'Mangler prosjektdata' });
    }

    const now = new Date();
    const { offerNumber, revision } = await allocateOfferIdentity(project, now);
    const buffer = await generateOfferDocxBuffer(project, offerNumber, now, revision);
    const projectName = sanitizeFileName(project.name || 'prosjekt');
    const fileName = `Tilbud-${projectName}-${offerNumber}-rev${revision}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Offer-Number', offerNumber);
    res.setHeader('X-Offer-Revision', String(revision));
    res.status(200).send(buffer);
  } catch (err) {
    console.error('Tilbudsgenerering feilet', err);
    if (err && err.code === 'ENOENT') {
      return res.status(500).json({ error: 'Fant ikke tilbudsmalen i server/templates/tilbud' });
    }
    res.status(500).json({ error: 'Kunne ikke generere tilbud' });
  }
});

const staticDir = path.resolve(__dirname, '..');
app.use(express.static(staticDir));

const port = Number(process.env.PORT) || 5500;
const host = String(process.env.HOST || '0.0.0.0').trim() || '0.0.0.0';
app.listen(port, host, () => {
  console.log(`Mail service lytter pa ${host}:${port}`);
  console.log(`[runtime-data] ${RUNTIME_DATA_DIR}`);
  if (!corsAllowAllOrigins) {
    console.log(`[cors] Tillatte origins: ${Array.from(corsAllowedOrigins).join(', ')}`);
  } else {
    console.log('[cors] Tillater alle origins (*)');
  }
  initializeMarketDataAutomation().catch(err=>{
    console.error('[market-data] Init feilet', err);
    scheduleRetryMarketRefresh();
  });
});

process.on('SIGINT', ()=>{
  stopMarketScheduler();
  process.exit(0);
});

process.on('SIGTERM', ()=>{
  stopMarketScheduler();
  process.exit(0);
});
