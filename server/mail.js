require('dotenv').config();
const path = require('path');
const fs = require('fs/promises');
const crypto = require('crypto');
const { promisify } = require('util');
const express = require('express');
const nodemailer = require('nodemailer');
const AdmZip = require('adm-zip');
const { ClientSecretCredential } = require('@azure/identity');
const scryptAsync = promisify(crypto.scrypt);

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
const OFFER_TEMPLATE_DIR = path.resolve(
  __dirname,
  'templates',
  'tilbud'
);
const OFFER_TEMPLATE_FILE = (() => {
  const rawFile = String(process.env.OFFER_TEMPLATE_FILE || '').trim();
  if (rawFile) return path.resolve(rawFile);

  const rawName = String(process.env.OFFER_TEMPLATE_NAME || '').trim();
  if (rawName) return path.resolve(OFFER_TEMPLATE_DIR, rawName);

  return '';
})();
const OFFER_COUNTER_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-sequence.json');
const OFFER_PROJECT_NUMBERS_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-project-numbers.json');
const OFFER_REVISIONS_FILE = path.resolve(RUNTIME_DATA_DIR, 'offer-revisions.json');
const PROJECT_ARCHIVE_FILE = path.resolve(RUNTIME_DATA_DIR, 'project-archive.json');
const USER_AUTH_FILE = path.resolve(RUNTIME_DATA_DIR, 'user-auth.json');
const CUSTOMER_DATABASE_FILE = path.resolve(RUNTIME_DATA_DIR, 'customer-database.json');
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
let userAuthLock = Promise.resolve();
let customerDatabaseLock = Promise.resolve();
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
  res.setHeader(
    'Access-Control-Expose-Headers',
    'Content-Disposition, X-Offer-Number, X-Offer-Revision, X-Offer-Filename, X-Offer-Template, X-Offer-Template-Path'
  );
  if (req.method === 'OPTIONS') {
    return res.sendStatus(204);
  }
  return next();
});

const ADMIN_USERNAME = safeString(process.env.ADMIN_USERNAME || 'admin');
const ADMIN_PASSWORD = safeString(process.env.ADMIN_PASSWORD || 'admin1');
const AUTH_TOKEN_SECRET = safeString(
  process.env.AUTH_TOKEN_SECRET ||
  process.env.SESSION_SECRET ||
  process.env.ADMIN_PASSWORD ||
  'dev-auth-token-secret-change-me'
);
const AUTH_TOKEN_TTL_SECONDS = (() => {
  const parsed = Number(process.env.AUTH_TOKEN_TTL_SECONDS || 60 * 60 * 24 * 30);
  if (!Number.isFinite(parsed) || parsed < 300) return 60 * 60 * 24 * 30;
  return Math.round(parsed);
})();
const MICROSOFT_AUTH_TENANT_ID = safeString(
  process.env.MICROSOFT_AUTH_TENANT_ID ||
  process.env.MS_AUTH_TENANT_ID ||
  'e1b96c2a-c273-40b9-bb46-a2a7b570e133'
);
const MICROSOFT_AUTH_CLIENT_ID = safeString(
  process.env.MICROSOFT_AUTH_CLIENT_ID ||
  process.env.MS_AUTH_CLIENT_ID ||
  '48570b46-4211-46ac-b43e-6eb1451ad1a5'
);
const MICROSOFT_AUTH_REDIRECT_URI = safeString(
  process.env.MICROSOFT_AUTH_REDIRECT_URI ||
  'https://hansrisnes.github.io/busbar-webapp-embedded-v2/auth/callback'
);
const MICROSOFT_AUTH_ALLOWED_ORIGINS = new Set([
  ...parseCsvEnv(process.env.MICROSOFT_AUTH_ALLOWED_ORIGINS),
  'https://hansrisnes.github.io',
  'http://localhost:5500',
  'http://127.0.0.1:5500'
]);
const MICROSOFT_AUTH_SCOPES = parseCsvEnv(
  process.env.MICROSOFT_AUTH_SCOPES ||
  'openid,profile,email,Calendars.Read,Mail.Read,Files.Read.All,Sites.Read.All'
);
const MICROSOFT_AUTH_CLIENT_SECRET = safeString(
  process.env.MICROSOFT_AUTH_CLIENT_SECRET ||
  process.env.MS_AUTH_CLIENT_SECRET
);
const MICROSOFT_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MICROSOFT_AUTH_ISSUER = `https://login.microsoftonline.com/${MICROSOFT_AUTH_TENANT_ID}/v2.0`;
const MICROSOFT_AUTH_JWKS_URL = `https://login.microsoftonline.com/${MICROSOFT_AUTH_TENANT_ID}/discovery/v2.0/keys`;
const MICROSOFT_OWNER_CACHE_TTL_MS = 5 * 60 * 1000;
const ADMIN_OWNER_EMAILS = new Set([
  ...parseCsvEnv(process.env.ADMIN_OWNER_EMAILS),
  ...parseCsvEnv(process.env.MICROSOFT_AUTH_OWNER_EMAILS),
  'hans.jakob.risnes@busbar.no',
  'lars@busbar.no'
].map(normalizeEmail).filter(isValidEmail));
let microsoftJwksCache = { fetchedAt: 0, keys: [] };
let microsoftOwnerCache = { fetchedAt: 0, ownerIds: new Set() };
let microsoftOwnerSecretWarned = false;

if (!process.env.ADMIN_USERNAME || !process.env.ADMIN_PASSWORD) {
  console.warn(
    '[admin] ADMIN_USERNAME eller ADMIN_PASSWORD mangler i miljøvariabler. Bruker midlertidige standardverdier.'
  );
}
if (!process.env.AUTH_TOKEN_SECRET && !process.env.SESSION_SECRET) {
  console.warn('[auth] AUTH_TOKEN_SECRET mangler i miljøvariabler. Bruker midlertidig lokal fallback.');
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

function normalizeMarginRate(value, fallback = 0.20) {
  const raw = toFiniteNumber(value);
  if (!Number.isFinite(raw)) return fallback;
  const rate = raw > 1 ? raw / 100 : raw;
  if (!Number.isFinite(rate)) return fallback;
  if (rate < 0) return 0;
  if (rate >= 1) return 0.95;
  return rate;
}

function applyDgToCost(cost, rate) {
  const safeCost = round2(toFiniteNumber(cost) || 0);
  const dgRate = normalizeMarginRate(rate);
  const factor = 1 - dgRate;
  if (!(factor > 0)) return safeCost;
  return round2(safeCost / factor);
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

function formatNoFxRate(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount)) return '';
  return new Intl.NumberFormat('no-NO', {
    minimumFractionDigits: 4,
    maximumFractionDigits: 4
  }).format(amount);
}

function formatNoInteger(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount)) return '';
  return new Intl.NumberFormat('no-NO', {
    maximumFractionDigits: 0
  }).format(Math.round(amount));
}

function formatNoPositiveInteger(value) {
  const amount = toFiniteNumber(value);
  if (!Number.isFinite(amount) || amount <= 0) return '';
  return formatNoInteger(amount);
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
    return JSON.parse(raw.replace(/^\uFEFF/, ''));
  } catch (err) {
    if (err && err.code === 'ENOENT') return fallbackValue;
    throw err;
  }
}

async function writeJsonFile(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), 'utf8');
}

async function resolveOfferTemplateFile() {
  if (OFFER_TEMPLATE_FILE) {
    await fs.access(OFFER_TEMPLATE_FILE);
    return OFFER_TEMPLATE_FILE;
  }

  const entries = await fs.readdir(OFFER_TEMPLATE_DIR, { withFileTypes: true });
  const candidates = await Promise.all(
    entries
      .filter(entry=>
        entry.isFile() &&
        !entry.name.startsWith('~$') &&
        entry.name.toLowerCase().endsWith('.docx')
      )
      .map(async entry=>{
        const filePath = path.resolve(OFFER_TEMPLATE_DIR, entry.name);
        const stat = await fs.stat(filePath);
        return { filePath, mtimeMs: stat.mtimeMs };
      })
  );

  if (!candidates.length) {
    const err = new Error(`Fant ingen .docx-mal i ${OFFER_TEMPLATE_DIR}`);
    err.code = 'ENOENT';
    throw err;
  }

  candidates.sort((a, b)=>b.mtimeMs - a.mtimeMs);
  return candidates[0].filePath;
}

function withProjectArchiveLock(task) {
  const run = projectArchiveLock.then(() => task());
  projectArchiveLock = run.catch(() => {});
  return run;
}

function withUserAuthLock(task) {
  const run = userAuthLock.then(() => task());
  userAuthLock = run.catch(() => {});
  return run;
}

function withCustomerDatabaseLock(task) {
  const run = customerDatabaseLock.then(() => task());
  customerDatabaseLock = run.catch(() => {});
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

function normalizeLookupKey(value) {
  return safeString(value).toLowerCase();
}

function normalizePassword(value) {
  return String(value || '');
}

function isValidPassword(value) {
  const password = normalizePassword(value);
  return password.length >= 4 && password.length <= 200;
}

function normalizeUserProfile(raw) {
  return {
    name: safeString(raw?.name),
    phone: safeString(raw?.phone),
    company: safeString(raw?.company),
    position: safeString(raw?.position)
  };
}

function hasAnyUserProfileValue(profile) {
  const normalized = normalizeUserProfile(profile);
  return Boolean(normalized.name || normalized.phone || normalized.company || normalized.position);
}

function isCompleteUserProfile(profile) {
  return Boolean(
    safeString(profile?.name) &&
    safeString(profile?.phone) &&
    safeString(profile?.company) &&
    safeString(profile?.position)
  );
}

function base64UrlEncode(value) {
  return Buffer.from(value).toString('base64url');
}

function signAuthPayload(encodedPayload) {
  return crypto
    .createHmac('sha256', AUTH_TOKEN_SECRET)
    .update(encodedPayload)
    .digest('base64url');
}

function isFallbackAdminEmail(email) {
  return ADMIN_OWNER_EMAILS.has(normalizeEmail(email));
}

function isStoredMicrosoftOwner(userRecord) {
  return Boolean(userRecord?.microsoft?.isOwner);
}

function resolveUserIsAdmin(userRecord) {
  return isFallbackAdminEmail(userRecord?.email) || isStoredMicrosoftOwner(userRecord);
}

async function resolveUserIsAdminFresh(email) {
  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) return false;
  if (isFallbackAdminEmail(normalizedEmail)) return true;

  return withUserAuthLock(async () => {
    const store = await readUserAuthStore();
    const userRecord = store.users[normalizedEmail] || { email: normalizedEmail };
    if (resolveUserIsAdmin(userRecord)) return true;

    const microsoftOid = safeString(userRecord?.microsoft?.oid);
    if (!microsoftOid) return false;

    const isOwner = await isMicrosoftAppOwner(microsoftOid);
    if (!isOwner) return false;

    store.users[normalizedEmail] = {
      ...userRecord,
      email: normalizedEmail,
      microsoft: {
        ...userRecord.microsoft,
        oid: microsoftOid,
        isOwner: true
      },
      updatedAt: new Date().toISOString()
    };
    await writeUserAuthStore(store);
    return true;
  });
}

function createAuthToken(email, options = {}) {
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    email,
    isAdmin: Boolean(options.isAdmin),
    iat: now,
    exp: now + AUTH_TOKEN_TTL_SECONDS
  };
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signature = signAuthPayload(encodedPayload);
  return `${encodedPayload}.${signature}`;
}

function verifyAuthToken(token) {
  const raw = safeString(token);
  const parts = raw.split('.');
  if (parts.length !== 2 || !parts[0] || !parts[1]) return null;
  const expected = signAuthPayload(parts[0]);
  const givenBuffer = Buffer.from(parts[1]);
  const expectedBuffer = Buffer.from(expected);
  if (givenBuffer.length !== expectedBuffer.length) return null;
  if (!crypto.timingSafeEqual(givenBuffer, expectedBuffer)) return null;
  try {
    const payload = JSON.parse(Buffer.from(parts[0], 'base64url').toString('utf8'));
    const email = normalizeEmail(payload?.email);
    const exp = Number(payload?.exp);
    if (!isValidEmail(email) || !Number.isFinite(exp)) return null;
    if (exp < Math.floor(Date.now() / 1000)) return null;
    return { email, isAdmin: payload?.isAdmin === true };
  } catch (_err) {
    return null;
  }
}

async function getMicrosoftGraphAppToken() {
  if (!MICROSOFT_AUTH_CLIENT_SECRET) {
    if (!microsoftOwnerSecretWarned) {
      microsoftOwnerSecretWarned = true;
      console.warn('[microsoft-auth] MICROSOFT_AUTH_CLIENT_SECRET mangler. Entra Owner-sjekk er deaktivert.');
    }
    return '';
  }
  const graphCredential = new ClientSecretCredential(
    MICROSOFT_AUTH_TENANT_ID,
    MICROSOFT_AUTH_CLIENT_ID,
    MICROSOFT_AUTH_CLIENT_SECRET
  );
  const token = await graphCredential.getToken(MICROSOFT_GRAPH_SCOPE);
  return token?.token || '';
}

async function fetchMicrosoftGraphJson(url) {
  const accessToken = await getMicrosoftGraphAppToken();
  if (!accessToken) return null;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json'
    }
  });
  if (!res.ok) {
    throw new Error(`Microsoft Graph svarte ${res.status}`);
  }
  return res.json();
}

async function collectMicrosoftOwnerIdsForCollection(collectionName) {
  const filter = encodeURIComponent(`appId eq '${MICROSOFT_AUTH_CLIENT_ID}'`);
  const payload = await fetchMicrosoftGraphJson(
    `https://graph.microsoft.com/v1.0/${collectionName}?$filter=${filter}&$select=id,appId`
  );
  const objectId = Array.isArray(payload?.value) ? safeString(payload.value[0]?.id) : '';
  const ownerIds = new Set();
  if (!objectId) return ownerIds;

  let nextUrl = `https://graph.microsoft.com/v1.0/${collectionName}/${encodeURIComponent(objectId)}/owners?$select=id`;
  while (nextUrl) {
    const ownersPayload = await fetchMicrosoftGraphJson(nextUrl);
    (Array.isArray(ownersPayload?.value) ? ownersPayload.value : []).forEach(owner => {
      const ownerId = safeString(owner?.id);
      if (ownerId) ownerIds.add(ownerId);
    });
    nextUrl = safeString(ownersPayload?.['@odata.nextLink']);
  }
  return ownerIds;
}

async function getMicrosoftAuthAppOwnerIds() {
  const now = Date.now();
  if (microsoftOwnerCache.fetchedAt && now - microsoftOwnerCache.fetchedAt < MICROSOFT_OWNER_CACHE_TTL_MS) {
    return microsoftOwnerCache.ownerIds;
  }
  const ownerIds = new Set();
  const errors = [];
  for (const collectionName of ['applications', 'servicePrincipals']) {
    try {
      const ids = await collectMicrosoftOwnerIdsForCollection(collectionName);
      ids.forEach(id => ownerIds.add(id));
    } catch (err) {
      errors.push(`${collectionName}: ${err?.message || err}`);
    }
  }
  if (!ownerIds.size && errors.length) {
    throw new Error(errors.join('; '));
  }
  microsoftOwnerCache = { fetchedAt: now, ownerIds };
  return ownerIds;
}

async function isMicrosoftAppOwner(objectId) {
  const oid = safeString(objectId);
  if (!oid) return false;
  try {
    const ownerIds = await getMicrosoftAuthAppOwnerIds();
    return ownerIds.has(oid);
  } catch (err) {
    console.warn('[microsoft-auth] Kunne ikke sjekke Entra Owners', err?.message || err);
    return false;
  }
}

function decodeBase64UrlJson(value) {
  try {
    return JSON.parse(Buffer.from(safeString(value), 'base64url').toString('utf8'));
  } catch (_err) {
    return null;
  }
}

async function fetchMicrosoftJwks() {
  const now = Date.now();
  if (Array.isArray(microsoftJwksCache.keys) && microsoftJwksCache.keys.length && now - microsoftJwksCache.fetchedAt < 60 * 60 * 1000) {
    return microsoftJwksCache.keys;
  }
  const res = await fetch(MICROSOFT_AUTH_JWKS_URL);
  if (!res.ok) {
    throw new Error(`Kunne ikke hente Microsoft-nøkler (${res.status})`);
  }
  const payload = await res.json();
  const keys = Array.isArray(payload?.keys) ? payload.keys : [];
  microsoftJwksCache = { fetchedAt: now, keys };
  return keys;
}

async function verifyMicrosoftIdToken(idToken) {
  const raw = safeString(idToken);
  const parts = raw.split('.');
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) {
    return null;
  }
  const header = decodeBase64UrlJson(parts[0]);
  const payload = decodeBase64UrlJson(parts[1]);
  if (!header || !payload || header.alg !== 'RS256' || !header.kid) {
    return null;
  }
  const keys = await fetchMicrosoftJwks();
  const jwk = keys.find(key => key?.kid === header.kid);
  if (!jwk) {
    microsoftJwksCache = { fetchedAt: 0, keys: [] };
    const refreshed = await fetchMicrosoftJwks();
    const refreshedJwk = refreshed.find(key => key?.kid === header.kid);
    if (!refreshedJwk) return null;
    return verifyMicrosoftIdTokenWithJwk(raw, parts, payload, refreshedJwk);
  }
  return verifyMicrosoftIdTokenWithJwk(raw, parts, payload, jwk);
}

function verifyMicrosoftIdTokenWithJwk(rawToken, parts, payload, jwk) {
  const verifier = crypto.createVerify('RSA-SHA256');
  verifier.update(`${parts[0]}.${parts[1]}`);
  verifier.end();
  const publicKey = crypto.createPublicKey({ key: jwk, format: 'jwk' });
  const signatureValid = verifier.verify(publicKey, Buffer.from(parts[2], 'base64url'));
  if (!signatureValid) return null;

  const now = Math.floor(Date.now() / 1000);
  const exp = Number(payload.exp);
  const nbf = Number(payload.nbf || 0);
  const aud = safeString(payload.aud);
  const iss = safeString(payload.iss);
  const tid = safeString(payload.tid);
  if (aud !== MICROSOFT_AUTH_CLIENT_ID) return null;
  if (iss !== MICROSOFT_AUTH_ISSUER) return null;
  if (tid !== MICROSOFT_AUTH_TENANT_ID) return null;
  if (!Number.isFinite(exp) || exp < now) return null;
  if (Number.isFinite(nbf) && nbf > now + 300) return null;

  const email = normalizeEmail(payload.preferred_username || payload.email || payload.upn);
  if (!isValidEmail(email)) return null;
  return {
    email,
    oid: safeString(payload.oid),
    name: safeString(payload.name),
    tenantId: tid,
    rawToken
  };
}

function normalizeMicrosoftAuthOrigin(value) {
  const raw = safeString(value);
  if (!raw) return '';
  try {
    const parsed = new URL(raw);
    if (!['http:', 'https:'].includes(parsed.protocol)) return '';
    return parsed.origin;
  } catch (_err) {
    return '';
  }
}

function resolveMicrosoftRedirectUri(req) {
  const requestedOrigin = normalizeMicrosoftAuthOrigin(req.query?.origin || req.headers.origin);
  const fallback = MICROSOFT_AUTH_REDIRECT_URI;
  if (!requestedOrigin || !MICROSOFT_AUTH_ALLOWED_ORIGINS.has(requestedOrigin)) {
    return fallback;
  }
  if (requestedOrigin === 'https://hansrisnes.github.io') {
    return `${requestedOrigin}/busbar-webapp-embedded-v2/auth/callback`;
  }
  return `${requestedOrigin}/auth/callback`;
}

function getBearerToken(req) {
  const header = safeString(req.headers.authorization);
  if (!header.toLowerCase().startsWith('bearer ')) return '';
  return header.slice(7).trim();
}

async function requireUserAuth(req, res, next) {
  const auth = verifyAuthToken(getBearerToken(req));
  if (!auth) {
    return res.status(401).json({ error: 'Logg inn for å hente prosjekter' });
  }
  if (auth.isAdmin !== true) {
    try {
      auth.isAdmin = await resolveUserIsAdminFresh(auth.email);
    } catch (err) {
      console.warn('[auth] Kunne ikke oppdatere adminstatus fra brukerregister', err?.message || err);
      auth.isAdmin = isFallbackAdminEmail(auth.email);
    }
  }
  req.userAuth = auth;
  return next();
}

function requireMicrosoftOwnerAuth(req, res, next) {
  if (req.userAuth?.isAdmin === true) {
    return next();
  }
  return res.status(403).json({ error: 'Kun Owners i Entra kan redigere globale firmadata' });
}

async function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('base64url');
  const hash = await scryptAsync(normalizePassword(password), salt, 64);
  return `scrypt:${salt}:${Buffer.from(hash).toString('base64url')}`;
}

async function verifyPassword(password, storedHash) {
  const parts = safeString(storedHash).split(':');
  if (parts.length !== 3 || parts[0] !== 'scrypt' || !parts[1] || !parts[2]) return false;
  const expected = Buffer.from(parts[2], 'base64url');
  const actual = await scryptAsync(normalizePassword(password), parts[1], expected.length);
  const actualBuffer = Buffer.from(actual);
  if (actualBuffer.length !== expected.length) return false;
  return crypto.timingSafeEqual(actualBuffer, expected);
}

function normalizeUserAuthRecord(email, raw) {
  const now = new Date().toISOString();
  const microsoft = raw?.microsoft && typeof raw.microsoft === 'object'
    ? {
        oid: safeString(raw.microsoft.oid),
        tenantId: safeString(raw.microsoft.tenantId),
        isOwner: raw.microsoft.isOwner === true,
        linkedAt: toIsoTimestamp(raw.microsoft.linkedAt, now),
        lastLoginAt: toIsoTimestamp(raw.microsoft.lastLoginAt || raw.microsoft.linkedAt, now)
      }
    : null;
  return {
    email,
    profile: normalizeUserProfile(raw?.profile || raw),
    passwordHash: safeString(raw?.passwordHash),
    ...(microsoft?.oid ? { microsoft } : {}),
    createdAt: toIsoTimestamp(raw?.createdAt, now),
    updatedAt: toIsoTimestamp(raw?.updatedAt || raw?.createdAt, now)
  };
}

async function readUserAuthStore() {
  const stored = await readJsonFile(USER_AUTH_FILE, { users: {} });
  const usersRaw = (stored && typeof stored === 'object' && stored.users && typeof stored.users === 'object')
    ? stored.users
    : {};
  const users = {};
  Object.entries(usersRaw).forEach(([key, value]) => {
    const email = normalizeEmail(key || value?.email);
    if (!isValidEmail(email)) return;
    const user = normalizeUserAuthRecord(email, value);
    if (!user.passwordHash && !user.microsoft?.oid && !hasAnyUserProfileValue(user.profile)) return;
    users[email] = user;
  });
  return { users };
}

async function writeUserAuthStore(state) {
  const users = (state && typeof state === 'object' && state.users && typeof state.users === 'object')
    ? state.users
    : {};
  await writeJsonFile(USER_AUTH_FILE, { users });
}

function normalizeCustomerContact(raw) {
  const name = safeString(raw?.name || raw?.contactPerson || raw?.contact);
  const phone = safeString(raw?.phone || raw?.contactPhone);
  const email = safeString(raw?.email || raw?.contactEmail);
  if (!name && !phone && !email) return null;
  return {
    id: safeString(raw?.id) || generateRecordId('contact'),
    name,
    phone,
    email
  };
}

function normalizeCustomerRecord(raw) {
  const name = safeString(raw?.name || raw?.customer);
  if (!name) return null;
  const contactsRaw = Array.isArray(raw?.contacts) ? raw.contacts : [];
  const contactsByKey = new Map();
  contactsRaw.forEach(contactRaw => {
    const contact = normalizeCustomerContact(contactRaw);
    if (!contact) return;
    const key = normalizeLookupKey(contact.name || contact.id);
    if (!key) return;
    contactsByKey.set(key, contact);
  });
  return {
    id: safeString(raw?.id) || generateRecordId('customer'),
    name,
    address: safeString(raw?.address || raw?.customerAddress),
    postalPlace: safeString(raw?.postalPlace || raw?.customerPostalPlace),
    segment: safeString(raw?.segment || raw?.customerSegment),
    customerResponsible: safeString(raw?.customerResponsible || raw?.responsible),
    projectCount: Number.isFinite(Number(raw?.projectCount)) ? Number(raw.projectCount) : 0,
    contacts: Array.from(contactsByKey.values())
  };
}

async function readCustomerDatabase() {
  const stored = await readJsonFile(CUSTOMER_DATABASE_FILE, { customers: [] });
  const rawCustomers = Array.isArray(stored?.customers) ? stored.customers : [];
  const byKey = new Map();
  rawCustomers.forEach(raw => {
    const customer = normalizeCustomerRecord(raw);
    if (!customer) return;
    byKey.set(normalizeLookupKey(customer.name), customer);
  });
  return { customers: Array.from(byKey.values()) };
}

async function writeCustomerDatabase(state) {
  const customers = Array.isArray(state?.customers)
    ? state.customers.map(normalizeCustomerRecord).filter(Boolean)
    : [];
  await writeJsonFile(CUSTOMER_DATABASE_FILE, { customers });
}

function mergeCustomerIntoMap(map, customerInput) {
  const customer = normalizeCustomerRecord(customerInput);
  if (!customer) return;
  const key = normalizeLookupKey(customer.name);
  const existing = map.get(key) || {
    id: customer.id,
    name: customer.name,
    address: '',
    postalPlace: '',
    segment: '',
    customerResponsible: '',
    projectCount: 0,
    contacts: []
  };
  existing.name = existing.name || customer.name;
  if (!existing.address && customer.address) existing.address = customer.address;
  if (!existing.postalPlace && customer.postalPlace) existing.postalPlace = customer.postalPlace;
  if (!existing.segment && customer.segment) existing.segment = customer.segment;
  if (!existing.customerResponsible && customer.customerResponsible) existing.customerResponsible = customer.customerResponsible;
  const contactsByKey = new Map(existing.contacts.map(contact => [normalizeLookupKey(contact.name), contact]));
  customer.contacts.forEach(contact => {
    const contactKey = normalizeLookupKey(contact.name);
    if (!contactKey) return;
    const existingContact = contactsByKey.get(contactKey) || { id: contact.id, name: contact.name, phone: '', email: '' };
    if (!existingContact.phone && contact.phone) existingContact.phone = contact.phone;
    if (!existingContact.email && contact.email) existingContact.email = contact.email;
    contactsByKey.set(contactKey, existingContact);
  });
  existing.contacts = Array.from(contactsByKey.values());
  map.set(key, existing);
}

function getProjectTimestampMs(project) {
  const updated = Date.parse(safeString(project?.updatedAt));
  if (!Number.isNaN(updated)) return updated;
  const created = Date.parse(safeString(project?.createdAt));
  return Number.isNaN(created) ? 0 : created;
}

function mergeProjectsByLatestForServer(leftProjects, rightProjects) {
  const merged = new Map();
  const add = project => {
    const normalized = normalizeProjectRecord(project);
    if (!normalized) return;
    const key = safeString(normalized.id) || resolveProjectOfferKey(normalized);
    const existing = merged.get(key);
    if (!existing || getProjectTimestampMs(normalized) >= getProjectTimestampMs(existing)) {
      merged.set(key, normalized);
    }
  };
  (Array.isArray(leftProjects) ? leftProjects : []).forEach(add);
  (Array.isArray(rightProjects) ? rightProjects : []).forEach(add);
  return Array.from(merged.values());
}

async function removeOfferMetadataForProjects(projects) {
  const keys = (Array.isArray(projects) ? projects : [])
    .map(resolveProjectOfferKey)
    .filter(Boolean);
  if (!keys.length) return { removedProjectNumbers: 0, removedRevisions: 0 };
  const keySet = new Set(keys);
  const [projectNumbersRaw, revisionsRaw] = await Promise.all([
    readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {}),
    readJsonFile(OFFER_REVISIONS_FILE, {})
  ]);
  const projectNumbers = isObject(projectNumbersRaw) ? projectNumbersRaw : {};
  const revisions = isObject(revisionsRaw) ? revisionsRaw : {};
  let removedProjectNumbers = 0;
  let removedRevisions = 0;
  keySet.forEach(key => {
    if (Object.prototype.hasOwnProperty.call(projectNumbers, key)) {
      delete projectNumbers[key];
      removedProjectNumbers += 1;
    }
    if (Object.prototype.hasOwnProperty.call(revisions, key)) {
      delete revisions[key];
      removedRevisions += 1;
    }
  });
  await Promise.all([
    writeJsonFile(OFFER_PROJECT_NUMBERS_FILE, projectNumbers),
    writeJsonFile(OFFER_REVISIONS_FILE, revisions)
  ]);
  return { removedProjectNumbers, removedRevisions };
}

const PROJECT_TRANSFER_SOURCE_EMAIL = 'hans.jakob.risnes@mcselektrotavler.no';
const PROJECT_TRANSFER_TARGET_EMAIL = 'hans.jakob.risnes@busbar.no';

async function migrateProjectsBetweenUsersOnStartup(
  sourceEmail = PROJECT_TRANSFER_SOURCE_EMAIL,
  targetEmail = PROJECT_TRANSFER_TARGET_EMAIL
) {
  const source = normalizeEmail(sourceEmail);
  const target = normalizeEmail(targetEmail);
  if (!isValidEmail(source) || !isValidEmail(target) || source === target) {
    return { moved: 0, skipped: true };
  }

  return withProjectArchiveLock(async () => {
    const archive = await readProjectArchive();
    const sourceUser = archive.users[source];
    const sourceProjects = Array.isArray(sourceUser?.projects) ? sourceUser.projects : [];
    if (!sourceProjects.length) {
      if (archive.users[source]) {
        delete archive.users[source];
        await writeProjectArchive(archive);
      }
      return { moved: 0, source, target, sourceRemoved: Boolean(sourceUser) };
    }

    const targetUser = archive.users[target] || { email: target, updatedAt: null, projects: [] };
    const movedProjects = sourceProjects
      .map(project => normalizeProjectRecord({
        ...project,
        projectOwnerEmail: target
      }))
      .filter(Boolean);
    const mergedProjects = mergeProjectsByLatestForServer(targetUser.projects, movedProjects)
      .map(project => ({
        ...project,
        projectOwnerEmail: target
      }));

    archive.users[target] = {
      email: target,
      updatedAt: new Date().toISOString(),
      projects: mergedProjects
    };
    delete archive.users[source];
    await writeProjectArchive(archive);

    return {
      moved: movedProjects.length,
      source,
      target,
      totalTargetProjects: mergedProjects.length
    };
  });
}

async function buildMergedCustomerDatabase() {
  const [database, archive] = await Promise.all([
    readCustomerDatabase(),
    readProjectArchive()
  ]);
  const byKey = new Map();
  const projectKeysByCustomer = new Map();
  database.customers.forEach(customer => mergeCustomerIntoMap(byKey, customer));
  Object.values(archive.users || {}).forEach(user => {
    (Array.isArray(user.projects) ? user.projects : []).forEach(project => {
      const customerName = safeString(project.customer);
      if (!customerName) return;
      const customerKey = normalizeLookupKey(customerName);
      const projectKey = safeString(project.id || project.projectNumber || project.name || project.createdAt);
      if (projectKey) {
        const projectKeys = projectKeysByCustomer.get(customerKey) || new Set();
        projectKeys.add(projectKey);
        projectKeysByCustomer.set(customerKey, projectKeys);
      }
      mergeCustomerIntoMap(byKey, {
        name: customerName,
        address: project.customerAddress,
        postalPlace: project.customerPostalPlace,
        contacts: safeString(project.contactPerson)
          ? [{ name: project.contactPerson, phone: project.contactPhone }]
          : []
      });
    });
  });
  const customers = Array.from(byKey.values()).map(customer => ({
    ...customer,
    projectCount: projectKeysByCustomer.get(normalizeLookupKey(customer.name))?.size || 0
  })).sort((a, b) => a.name.localeCompare(b.name, 'no', {
    sensitivity: 'base',
    numeric: true
  }));
  return { customers };
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

function normalizeProjectStatusRecord(value) {
  const raw = safeString(value).toLowerCase();
  if (!raw) return 'unresolved';
  if (['unresolved', 'won', 'lost', 'finished'].includes(raw)) return raw;
  const mapped = {
    uavklart: 'unresolved',
    vunnet: 'won',
    tapt: 'lost',
    ferdig: 'finished'
  };
  return mapped[raw] || 'unresolved';
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
    projectNumber: safeString(raw.projectNumber || raw.project_number || raw.offerNumber || raw.offer_number),
    name: safeString(raw.name),
    customer: safeString(raw.customer),
    contactPerson: safeString(raw.contactPerson || raw.contact),
    customerAddress: safeString(raw.customerAddress || raw.address),
    customerPostalPlace: safeString(raw.customerPostalPlace || raw.postalPlace),
    contactPhone: safeString(raw.contactPhone || raw.phone),
    projectResponsible: safeString(raw.projectResponsible || raw.projectOwner || raw.ownerName),
    projectOwnerEmail: normalizeEmail(raw.projectOwnerEmail || raw.ownerEmail),
    projectFolderName: safeString(raw.projectFolderName),
    projectFolderCreated: raw.projectFolderCreated === true,
    projectFolderWebUrl: safeString(raw.projectFolderWebUrl),
    projectStatus: normalizeProjectStatusRecord(raw.projectStatus || raw.status),
    createdAt: toIsoTimestamp(raw.createdAt, now),
    updatedAt: toIsoTimestamp(raw.updatedAt || raw.createdAt, now),
    selectedAddonConfig,
    lines
  };
}

function buildVisibleProjectsForAuth(archive, auth) {
  const email = normalizeEmail(auth?.email);
  if (!auth?.isAdmin) {
    const user = archive.users[email] || { email, updatedAt: null, projects: [] };
    return {
      email,
      updatedAt: user.updatedAt,
      ownerEmails: [email],
      projects: (Array.isArray(user.projects) ? user.projects : []).map(project => ({
        ...project,
        projectOwnerEmail: email
      }))
    };
  }
  const projects = [];
  let updatedAt = null;
  Object.entries(archive.users || {}).forEach(([ownerEmailRaw, user]) => {
    const ownerEmail = normalizeEmail(ownerEmailRaw || user?.email);
    if (!isValidEmail(ownerEmail)) return;
    const userUpdatedAt = toIsoTimestamp(user?.updatedAt, '');
    if (userUpdatedAt && (!updatedAt || new Date(userUpdatedAt) > new Date(updatedAt))) {
      updatedAt = userUpdatedAt;
    }
    (Array.isArray(user?.projects) ? user.projects : []).forEach(project => {
      projects.push({
        ...project,
        projectOwnerEmail: ownerEmail
      });
    });
  });
  return {
    email,
    updatedAt,
    projects,
    ownerEmails: Object.keys(archive.users || {}).map(normalizeEmail).filter(isValidEmail)
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

function isValidOfferNumber(value) {
  return /^\d{4}-\d+$/.test(safeString(value));
}

function ensureProjectNumber(project, projectNumbers, counterState, date = new Date()) {
  if (!project || typeof project !== 'object') return '';
  const key = resolveProjectOfferKey(project);
  let projectNumber = safeString(project.projectNumber);
  const mappedNumber = safeString(projectNumbers?.[key]);

  if (!isValidOfferNumber(projectNumber) && isValidOfferNumber(mappedNumber)) {
    projectNumber = mappedNumber;
  }
  if (!isValidOfferNumber(projectNumber)) {
    projectNumber = allocateOfferNumberFromState(counterState, date);
  }

  project.projectNumber = projectNumber;
  if (projectNumbers && key) projectNumbers[key] = projectNumber;
  return projectNumber;
}

function ensureProjectNumbersForArchive(archive, projectNumbers, counterState, date = new Date()) {
  const stats = { users: 0, projects: 0, added: 0, preserved: 0 };
  const users = (archive && typeof archive === 'object' && archive.users && typeof archive.users === 'object')
    ? archive.users
    : {};
  Object.values(users).forEach(user=>{
    stats.users += 1;
    (Array.isArray(user?.projects) ? user.projects : []).forEach(project=>{
      stats.projects += 1;
      const before = safeString(project?.projectNumber);
      ensureProjectNumber(project, projectNumbers, counterState, date);
      if (isValidOfferNumber(before)) stats.preserved += 1;
      else stats.added += 1;
    });
  });
  return stats;
}

async function migrateProjectNumbersForAllUsersOnStartup() {
  return withOfferNumberLock(async ()=>{
    const [counterState, projectNumbersRaw] = await Promise.all([
      readJsonFile(OFFER_COUNTER_FILE, { years: {} }),
      readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {})
    ]);
    const projectNumbers = (projectNumbersRaw && typeof projectNumbersRaw === 'object')
      ? projectNumbersRaw
      : {};
    const stats = await withProjectArchiveLock(async ()=>{
      const archive = await readProjectArchive();
      const result = ensureProjectNumbersForArchive(archive, projectNumbers, counterState);
      await writeProjectArchive(archive);
      return result;
    });
    await Promise.all([
      writeJsonFile(OFFER_COUNTER_FILE, counterState),
      writeJsonFile(OFFER_PROJECT_NUMBERS_FILE, projectNumbers)
    ]);
    return stats;
  });
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

    let offerNumber = safeString(project?.projectNumber);
    if (!isValidOfferNumber(offerNumber)) offerNumber = safeString(projectNumbers[key]);
    if (!isValidOfferNumber(offerNumber)) {
      offerNumber = allocateOfferNumberFromState(counterState, date);
    }
    projectNumbers[key] = offerNumber;

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

async function getOfferStatusForProjects(projects) {
  const [projectNumbersRaw, revisionsRaw] = await Promise.all([
    readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {}),
    readJsonFile(OFFER_REVISIONS_FILE, {})
  ]);
  const projectNumbers = (projectNumbersRaw && typeof projectNumbersRaw === 'object')
    ? projectNumbersRaw
    : {};
  const revisions = (revisionsRaw && typeof revisionsRaw === 'object')
    ? revisionsRaw
    : {};
  return (Array.isArray(projects) ? projects : []).map(project=>{
    const key = resolveProjectOfferKey(project);
    const revisionValue = Number(revisions[key]);
    const revision = Number.isInteger(revisionValue) && revisionValue >= 0 ? revisionValue : null;
    const projectNumber = safeString(project?.projectNumber) || safeString(projectNumbers[key]);
    return {
      projectId: safeString(project?.id),
      projectKey: key,
      offerNumber: isValidOfferNumber(projectNumber) ? projectNumber : '',
      revision,
      hasOffer: revision !== null && isValidOfferNumber(projectNumber)
    };
  });
}

async function persistProjectNumberForUser(email, projectId, projectNumber) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedProjectId = safeString(projectId);
  const normalizedProjectNumber = safeString(projectNumber);
  if (!isValidEmail(normalizedEmail) || !normalizedProjectId || !isValidOfferNumber(normalizedProjectNumber)) return;
  await withProjectArchiveLock(async ()=>{
    const archive = await readProjectArchive();
    const user = archive.users[normalizedEmail];
    const projects = Array.isArray(user?.projects) ? user.projects : [];
    const project = projects.find(entry=>safeString(entry?.id) === normalizedProjectId);
    if (!project) return;
    project.projectNumber = normalizedProjectNumber;
    await writeProjectArchive(archive);
  });
}

function resolveWritableProjectOwnerEmail(project, auth) {
  const fallback = normalizeEmail(auth?.email);
  const requested = normalizeEmail(project?.projectOwnerEmail || project?.ownerEmail);
  if (auth?.isAdmin === true && isValidEmail(requested)) return requested;
  return fallback;
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
  const includeUnitPrices = resolveSelectedAddonFlag(raw?.includeUnitPrices, false);
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
  const tapOffOfferTotal = resolveTapOffOfferPriceTotal(line, line?.inputs);
  const specialElementOfferTotal = resolveSpecialElementOfferPriceTotal(line, line?.inputs);

  let total = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
  if (Number.isFinite(tapOffOfferTotal)) total += tapOffOfferTotal;
  if (Number.isFinite(specialElementOfferTotal)) total += specialElementOfferTotal;
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
    tapOffBoxTotal: 0,
    tapOffOfferTotal: 0,
    specialElementTotal: 0,
    specialElementOfferTotal: 0,
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
    const explicitTapOffBoxTotal = toFiniteNumber(lineTotals.tapOffBoxTotal);
    add(
      'tapOffBoxTotal',
      Number.isFinite(explicitTapOffBoxTotal)
        ? explicitTapOffBoxTotal
        : resolveTapOffBoxPriceTotal(line, line?.inputs)
    );
    add('tapOffOfferTotal', resolveTapOffOfferPriceTotal(line, line?.inputs));
    add('specialElementTotal', resolveSpecialElementCostTotal(line, line?.inputs));
    add('specialElementOfferTotal', resolveSpecialElementOfferPriceTotal(line, line?.inputs));
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

function resolveIpGradeFromSeries(rawSeries) {
  const series = safeString(rawSeries).toUpperCase();
  if (series === 'RCP-IP68') return 'IP68';
  if (series === 'XCM' || series === 'XCP-S' || series === 'XAP-B') return 'IP55';
  return '';
}

function resolveBoxLabelFromSelection(value) {
  const [kind, ampRaw] = safeString(value).split('|');
  const amp = safeString(ampRaw);
  const labels = {
    plug_in_box: 'Plug-in box (plast)',
    tap_off_box: 'Tap-off box (metall)',
    bolt_on_box: 'Bolt-on box (metall)'
  };
  const label = labels[kind] || safeString(kind);
  if (!amp && !label) return '';
  if (!amp) return label;
  if (!label) return `${amp}A`;
  return `${amp}A · ${label}`;
}

function isSeparateTapOffBoxType(value) {
  return ['plug_in_box', 'tap_off_box'].includes(safeString(value).toLowerCase());
}

function isTapOffInnmatType(value) {
  return ['plug_in_box_innmat', 'tap_off_box_innmat'].includes(safeString(value).toLowerCase());
}

function isSeparateTapOffBoxBomLine(entry) {
  const type = entry?.type || entry?.element_type || entry?.elementType;
  return isSeparateTapOffBoxType(type) || isTapOffInnmatType(type);
}

function isTapOffInnmatBomLine(entry) {
  return isTapOffInnmatType(entry?.type || entry?.element_type || entry?.elementType)
    || entry?.tapOffInnmatLine === true;
}

function resolveBomLineSum(entry) {
  const direct = toFiniteNumber(entry?.sum);
  if (Number.isFinite(direct)) return direct;
  const unit = toFiniteNumber(entry?.enhet ?? entry?.unit ?? entry?.unit_price);
  const qty = toFiniteNumber(entry?.antall ?? entry?.qty ?? entry?.quantity);
  if (Number.isFinite(unit) && Number.isFinite(qty)) return unit * qty;
  return 0;
}

function resolveTapOffInnmatTotalFromInputForBomEntry(entry, inputItems, usedIndexes) {
  const entryType = safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase();
  const entryAmp = toFiniteNumber(entry?.ampere);
  const entryQty = toFiniteNumber(entry?.antall ?? entry?.qty ?? entry?.quantity);

  for (let i = 0; i < inputItems.length; i += 1) {
    if (usedIndexes.has(i)) continue;
    const item = inputItems[i];
    const [itemType, itemAmpRaw] = safeString(item.boxSel).split('|');
    const itemAmp = toFiniteNumber(itemAmpRaw);
    const itemQty = toFiniteNumber(item.qty);
    const sameType = safeString(itemType).toLowerCase() === entryType;
    const sameAmp = !Number.isFinite(entryAmp) || !Number.isFinite(itemAmp) || Math.round(entryAmp) === Math.round(itemAmp);
    const sameQty = !Number.isFinite(entryQty) || !Number.isFinite(itemQty) || Math.round(entryQty) === Math.round(itemQty);
    if (!sameType || !sameAmp || !sameQty) continue;
    usedIndexes.add(i);
    const innmat = toFiniteNumber(item.innmatSum);
    const qty = Number.isFinite(entryQty) ? entryQty : itemQty;
    return Number.isFinite(innmat) && Number.isFinite(qty) ? innmat * qty : 0;
  }

  return 0;
}

function resolveTapOffBoxPriceTotal(line, input = {}) {
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  const inputItems = resolveTapOffItemsFromInput(input);
  const usedInputIndexes = new Set();
  const hasSeparateInnmatLines = bom.some(isTapOffInnmatBomLine);
  return round2(bom.reduce((sum, entry)=>{
    if (!isSeparateTapOffBoxBomLine(entry)) return sum;
    const baseSum = resolveBomLineSum(entry);
    if (isTapOffInnmatBomLine(entry)) return sum + baseSum;
    const includesInnmat = entry?.tapOffIncludesInnmatInSum === true;
    const explicitInnmatTotal = toFiniteNumber(entry?.tapOffInnmatTotal);
    if (includesInnmat) return sum + baseSum;
    if (hasSeparateInnmatLines) return sum + baseSum;
    if (Number.isFinite(explicitInnmatTotal) && explicitInnmatTotal > 0) {
      return sum + baseSum + explicitInnmatTotal;
    }
    return sum + baseSum + resolveTapOffInnmatTotalFromInputForBomEntry(entry, inputItems, usedInputIndexes);
  }, 0));
}

function resolveTapOffOfferPriceTotal(line, input = {}) {
  const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const explicit = toFiniteNumber(lineTotals.tapOffOfferTotal);
  if (Number.isFinite(explicit)) return round2(explicit);
  const cost = resolveTapOffBoxPriceTotal(line, input);
  if (!Number.isFinite(cost)) return NaN;
  const rate = lineTotals.tapOffMarginRate ?? input?.tapOffMarginRate ?? lineTotals.marginRate ?? input?.marginRate;
  return applyDgToCost(cost, rate);
}

function resolveTapOffItemsFromInput(input = {}) {
  const directItems = Array.isArray(input?.boxItems) ? input.boxItems : [];
  const normalizedDirect = directItems
    .map(item=>{
      const boxSel = safeString(item?.boxSel || item?.value || '');
      const qtyRaw = toFiniteNumber(item?.boxQty ?? item?.qty);
      const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
      const innmatRaw = toFiniteNumber(item?.innmatSum ?? item?.innmat);
      const innmatSum = Number.isFinite(innmatRaw) ? Math.max(0, innmatRaw) : 0;
      if (!boxSel || qty <= 0) return null;
      const ampRaw = toFiniteNumber(boxSel.split('|')[1]);
      const amp = Number.isFinite(ampRaw) && ampRaw > 0 ? Math.round(ampRaw) : null;
      return { boxSel, qty, amp, innmatSum };
    })
    .filter(Boolean);
  if (normalizedDirect.length) return normalizedDirect;

  const legacyBoxSel = safeString(input?.boxSel || '');
  const legacyQtyRaw = toFiniteNumber(input?.boxQty);
  if (legacyBoxSel && Number.isFinite(legacyQtyRaw) && legacyQtyRaw > 0) {
    const ampRaw = toFiniteNumber(legacyBoxSel.split('|')[1]);
    const amp = Number.isFinite(ampRaw) && ampRaw > 0 ? Math.round(ampRaw) : null;
    const innmatRaw = toFiniteNumber(input?.boxInnmatSum);
    const innmatSum = Number.isFinite(innmatRaw) ? Math.max(0, innmatRaw) : 0;
    return [{ boxSel: legacyBoxSel, qty: Math.round(legacyQtyRaw), amp, innmatSum }];
  }

  return [];
}

function resolveTapOffItemsFromLine(line, input = {}) {
  const fromInput = resolveTapOffItemsFromInput(input);
  if (fromInput.length) return fromInput;

  const bom = Array.isArray(line?.bom) ? line.bom : [];
  const types = new Set(['plug_in_box', 'tap_off_box', 'bolt_on_box']);
  const fromBom = bom
    .filter(entry=>{
      const type = safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase();
      if (isTapOffInnmatType(type) || entry?.tapOffInnmatLine === true) return false;
      const code = safeString(entry?.code).toLowerCase();
      const desc = safeString(entry?.desc || entry?.description || entry?.tekst).toLowerCase();
      const haystack = `${type} ${code} ${desc}`;
      return (
        types.has(type)
        || /\b(tap[\s-]*off|plug[\s-]*in|bolt[\s-]*on)\b/.test(haystack)
        || /\b(avtapping|avtappings|boks|box)\b/.test(haystack)
      );
    })
    .map(entry=>{
      const qtyRaw = toFiniteNumber(entry?.antall ?? entry?.qty);
      const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
      if (qty <= 0) return null;
      const ampRawDirect = toFiniteNumber(entry?.ampere);
      const ampFromCodeMatch = safeString(entry?.code).match(/(\d{2,4})\s*A/i);
      const ampFromCode = ampFromCodeMatch ? Number(ampFromCodeMatch[1]) : NaN;
      const amp = Number.isFinite(ampRawDirect)
        ? Math.round(ampRawDirect)
        : (Number.isFinite(ampFromCode) ? Math.round(ampFromCode) : null);
      return {
        boxSel: amp
          ? `${safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase()}|${amp}`
          : safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase(),
        code: safeString(entry?.code),
        qty,
        amp,
        innmatSum: 0
      };
    })
    .filter(Boolean);
  return fromBom;
}

function buildTapOffOfferText(line, input = {}) {
  const items = resolveTapOffItemsFromLine(line, input);
  if (!items.length) return '';
  return items.map(item=>{
    const label = resolveBoxLabelFromSelection(item.boxSel) || item.code || 'Avtappingsboks';
    const qtyTxt = formatNoInteger(item.qty) || String(item.qty);
    return `${label} · antall ${qtyTxt}`;
  }).join(' | ');
}

function formatTapOffOfferPricePlaceholder(value, qty = 0) {
  const amount = toFiniteNumber(value);
  const count = toFiniteNumber(qty);
  if (!Number.isFinite(amount) || amount <= 0) return '';
  if (Number.isFinite(count) && count <= 0) return '';
  return formatNoCurrencyWithKr(amount);
}

function resolveSpecialElementLabel(selection) {
  const labels = {
    phase_change: 'Faseendring',
    neutral_change: 'Nøytralendring',
    epoxy_metal_transition: 'Overgang til epoxy-/metallkapslet'
  };
  const key = safeString(selection);
  return labels[key] || key;
}

function resolveSpecialElementItemsFromLine(line, input = {}) {
  const directItems = Array.isArray(input?.specialElementItems) ? input.specialElementItems : [];
  const normalizedDirect = directItems
    .map(item=>{
      const selection = safeString(item?.selection || item?.value || item?.type);
      const qtyRaw = toFiniteNumber(item?.qty ?? item?.quantity);
      const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
      const unitSumRaw = toFiniteNumber(item?.unitSum ?? item?.sum ?? item?.elementSum);
      const unitSum = Number.isFinite(unitSumRaw) ? Math.max(0, unitSumRaw) : 0;
      if (!selection || qty <= 0) return null;
      return { selection, label: resolveSpecialElementLabel(selection), qty, unitSum };
    })
    .filter(Boolean);
  if (normalizedDirect.length) return normalizedDirect;

  const bom = Array.isArray(line?.bom) ? line.bom : [];
  return bom
    .filter(entry=>safeString(entry?.specialElementGroupId))
    .map(entry=>{
      const selection = safeString(entry?.specialElementSelection || entry?.type);
      const qtyRaw = toFiniteNumber(entry?.antall ?? entry?.qty ?? entry?.quantity);
      const qty = Number.isFinite(qtyRaw) ? Math.max(0, Math.round(qtyRaw)) : 0;
      const unitSumRaw = toFiniteNumber(entry?.enhet ?? entry?.unit ?? entry?.unit_price);
      const unitSum = Number.isFinite(unitSumRaw) ? Math.max(0, unitSumRaw) : 0;
      if (!selection || qty <= 0) return null;
      return {
        selection,
        label: safeString(entry?.type) || resolveSpecialElementLabel(selection),
        qty,
        unitSum
      };
    })
    .filter(Boolean);
}

function buildSpecialElementOfferText(line, input = {}) {
  const items = resolveSpecialElementItemsFromLine(line, input);
  if (!items.length) return '';
  return items.map(item=>{
    const qtyTxt = formatNoInteger(item.qty) || String(item.qty);
    return `${item.label || 'Spesialelement'} · antall ${qtyTxt}`;
  }).join(' | ');
}

function resolveSpecialElementCostTotal(line, input = {}) {
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  const bomItems = bom.filter(entry=>safeString(entry?.specialElementGroupId));
  if (bomItems.length) {
    return round2(bomItems.reduce((sum, entry)=>sum + resolveBomLineSum(entry), 0));
  }
  return round2(resolveSpecialElementItemsFromLine(line, input).reduce(
    (sum, item)=>sum + (Number(item.unitSum || 0) * Number(item.qty || 0)),
    0
  ));
}

function resolveSpecialElementOfferPriceTotal(line, input = {}) {
  const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const explicit = toFiniteNumber(lineTotals.specialElementOfferTotal);
  if (Number.isFinite(explicit)) return round2(explicit);
  const cost = resolveSpecialElementCostTotal(line, input);
  if (!Number.isFinite(cost)) return NaN;
  const rate = lineTotals.tapOffMarginRate ?? input?.tapOffMarginRate ?? lineTotals.marginRate ?? input?.marginRate;
  return applyDgToCost(cost, rate);
}

function resolveBomUnitByType(line, typeNames) {
  const types = new Set((Array.isArray(typeNames) ? typeNames : [typeNames])
    .map(type=>safeString(type).toLowerCase())
    .filter(Boolean));
  if (!types.size) return NaN;
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  const match = bom.find(entry=>{
    const type = safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase();
    if (!types.has(type)) return false;
    if (isTapOffInnmatBomLine(entry)) return false;
    const unit = toFiniteNumber(entry?.enhet ?? entry?.unit ?? entry?.unit_price);
    return Number.isFinite(unit) && unit > 0;
  });
  if (!match) return NaN;
  return toFiniteNumber(match.enhet ?? match.unit ?? match.unit_price);
}

function applyMaterialOfferUnitPrice(unitCost, input = {}, lineTotals = {}) {
  const unit = toFiniteNumber(unitCost);
  if (!Number.isFinite(unit) || unit <= 0) return NaN;
  const marginRate = normalizeMarginRate(lineTotals.marginRate ?? input.marginRate, 0.20);
  const factor = 1 - marginRate;
  if (!(factor > 0)) return NaN;
  const freightRaw = toFiniteNumber(input.freightRate ?? lineTotals.freightRate);
  const freightRate = Number.isFinite(freightRaw) ? Math.max(0, freightRaw) : 0;
  return round2((unit / factor) + (unit * freightRate));
}

function formatMaterialOfferUnitPrice(unitCost, input = {}, lineTotals = {}) {
  const price = applyMaterialOfferUnitPrice(unitCost, input, lineTotals);
  return Number.isFinite(price) ? formatNoCurrencyWithKr(price) : '';
}

const UNIT_PRICE_PLACEHOLDER_GROUPS = [
  ['enhetspris_meter', 'meter_enhetspris', 'mtr_enhetspris', 'mtr_enhetspris_nok'],
  ['enhetspris_vinkel', 'vinkel_enhetspris', 'vinkel_enhetspris_nok'],
  ['enhetspris_vinkel_vertikal', 'vertikal_vinkel_enhetspris', 'vvk_enhetspris', 'vvk_enhetspris_nok'],
  ['enhetspris_vinkel_horisontal', 'horisontal_vinkel_enhetspris', 'hvk_enhetspris', 'hvk_enhetspris_nok'],
  ['enhetspris_tavleelement', 'tavleelement_enhetspris', 'ste_enhetspris', 'ste_enhetspris_nok'],
  ['enhetspris_sluttelement', 'sluttelement_enhetspris', 'sle_enhetspris', 'sle_enhetspris_nok'],
  ['enhetspris_brann', 'brann_enhetspris', 'bre_enhetspris', 'bre_enhetspris_nok'],
  ['enhetspris_ekspansjon', 'ekspansjon_enhetspris', 'exp_enhetspris', 'exp_enhetspris_nok'],
  ['enhetspris_avtappingsboks', 'avtappingsboks_enhetspris', 'avb_enhetspris', 'avb_enhetspris_nok']
];

function buildEmptyUnitPricePlaceholders() {
  const keys = UNIT_PRICE_PLACEHOLDER_GROUPS.flat();
  return Object.fromEntries(keys.map(key=>[key, '']));
}

function buildAggregateUnitPricePlaceholders(linePlaceholderSets) {
  const output = buildEmptyUnitPricePlaceholders();
  const sets = Array.isArray(linePlaceholderSets) ? linePlaceholderSets : [];

  UNIT_PRICE_PLACEHOLDER_GROUPS.forEach(keys=>{
    const values = [];
    sets.forEach(placeholders=>{
      keys.forEach(key=>{
        const value = safeString(placeholders?.[key]);
        if (value && !values.includes(value)) values.push(value);
      });
    });
    const aggregateValue = values.join(' / ');
    keys.forEach(key=>{
      output[key] = aggregateValue;
    });
  });

  return output;
}

function hasPositiveQuantity(value) {
  const qty = toFiniteNumber(value);
  return Number.isFinite(qty) && qty > 0;
}

function buildUnitPricePlaceholders(line, input, lineTotals, fireBarrierPriceIndex) {
  const empty = buildEmptyUnitPricePlaceholders();
  const selectedAddonConfig = resolveLineSelectedAddonConfig(line, lineTotals);
  if (!selectedAddonConfig.includeUnitPrices) return empty;

  const rawUnitPrices = (lineTotals?.rawUnitPrices && typeof lineTotals.rawUnitPrices === 'object')
    ? lineTotals.rawUnitPrices
    : {};
  const rawUnit = key => {
    const value = toFiniteNumber(rawUnitPrices[key]);
    return Number.isFinite(value) && value > 0 ? value : NaN;
  };
  const startType = safeString(input?.startEl);
  const endType = safeString(input?.sluttEl);
  const tapOffTypes = ['plug_in_box', 'tap_off_box', 'bolt_on_box'];
  const fireUnit = resolveFireBarrierUnitPrice(line, fireBarrierPriceIndex, input);
  const meterQty = toFiniteNumber(input?.meter);
  const verticalAngleQty = toFiniteNumber(input?.v90_v ?? input?.v90v);
  const horizontalAngleQty = toFiniteNumber(input?.v90_h ?? input?.v90h);
  const brannQty = toFiniteNumber(input?.fbQty ?? input?.fireBarrierQty);
  const expansionQty = resolveExpansionQtyFromLine(line, input);
  const tapOffQty = resolveTapOffItemsFromLine(line, input).reduce((sum, item)=>sum + (toFiniteNumber(item?.qty) || 0), 0);
  const values = {
    meter: hasPositiveQuantity(meterQty) ? formatMaterialOfferUnitPrice(rawUnit('meter') || resolveBomUnitByType(line, [
      'straight_500_1000',
      'straight_500_1000_dist',
      'xcm_feeder_600_1500',
      'xcm_dist_1000_1500'
    ]), input, lineTotals) : '',
    vinkelVertikal: hasPositiveQuantity(verticalAngleQty)
      ? formatMaterialOfferUnitPrice(rawUnit('vinkelVertikal') || rawUnit('vinkel') || resolveBomUnitByType(line, 'elbow_vertical_90'), input, lineTotals)
      : '',
    vinkelHorisontal: hasPositiveQuantity(horizontalAngleQty)
      ? formatMaterialOfferUnitPrice(rawUnit('vinkelHorisontal') || resolveBomUnitByType(line, 'elbow_horizontal_90'), input, lineTotals)
      : '',
    tavleelement: startType && startType !== 'none'
      ? formatMaterialOfferUnitPrice(rawUnit('tavleelement') || resolveBomUnitByType(line, startType), input, lineTotals)
      : '',
    sluttelement: endType && endType !== 'none'
      ? formatMaterialOfferUnitPrice(rawUnit('sluttelement') || resolveBomUnitByType(line, endType), input, lineTotals)
      : '',
    brann: hasPositiveQuantity(brannQty) ? formatMaterialOfferUnitPrice(rawUnit('brann') || fireUnit, input, lineTotals) : '',
    ekspansjon: hasPositiveQuantity(expansionQty) ? formatMaterialOfferUnitPrice(rawUnit('ekspansjon') || resolveBomUnitByType(line, 'expansion_unit'), input, lineTotals) : '',
    avtappingsboks: hasPositiveQuantity(tapOffQty) ? formatMaterialOfferUnitPrice(rawUnit('avtappingsboks') || resolveBomUnitByType(line, tapOffTypes), input, lineTotals) : ''
  };

  return {
    ...empty,
    enhetspris_meter: values.meter,
    meter_enhetspris: values.meter,
    mtr_enhetspris: values.meter,
    mtr_enhetspris_nok: values.meter,
    enhetspris_vinkel: values.vinkelVertikal,
    vinkel_enhetspris: values.vinkelVertikal,
    vinkel_enhetspris_nok: values.vinkelVertikal,
    enhetspris_vinkel_vertikal: values.vinkelVertikal,
    vertikal_vinkel_enhetspris: values.vinkelVertikal,
    vvk_enhetspris: values.vinkelVertikal,
    vvk_enhetspris_nok: values.vinkelVertikal,
    enhetspris_vinkel_horisontal: values.vinkelHorisontal,
    horisontal_vinkel_enhetspris: values.vinkelHorisontal,
    hvk_enhetspris: values.vinkelHorisontal,
    hvk_enhetspris_nok: values.vinkelHorisontal,
    enhetspris_tavleelement: values.tavleelement,
    tavleelement_enhetspris: values.tavleelement,
    ste_enhetspris: values.tavleelement,
    ste_enhetspris_nok: values.tavleelement,
    enhetspris_sluttelement: values.sluttelement,
    sluttelement_enhetspris: values.sluttelement,
    sle_enhetspris: values.sluttelement,
    sle_enhetspris_nok: values.sluttelement,
    enhetspris_brann: values.brann,
    brann_enhetspris: values.brann,
    bre_enhetspris: values.brann,
    bre_enhetspris_nok: values.brann,
    enhetspris_ekspansjon: values.ekspansjon,
    ekspansjon_enhetspris: values.ekspansjon,
    exp_enhetspris: values.ekspansjon,
    exp_enhetspris_nok: values.ekspansjon,
    enhetspris_avtappingsboks: values.avtappingsboks,
    avtappingsboks_enhetspris: values.avtappingsboks,
    avb_enhetspris: values.avtappingsboks,
    avb_enhetspris_nok: values.avtappingsboks
  };
}

function resolveExpansionQtyFromBom(line) {
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  return bom.reduce((sum, entry)=>{
    const type = safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase();
    const code = safeString(entry?.code).toLowerCase();
    const desc = safeString(entry?.desc || entry?.description || entry?.tekst).toLowerCase();
    const looksLikeExpansion = type === 'expansion_unit' || /\bexpans/.test(type) || /\bexpans/.test(code) || /\bexpans/.test(desc);
    if (!looksLikeExpansion) return sum;
    const qty = toFiniteNumber(entry?.antall ?? entry?.qty ?? entry?.quantity);
    const normalizedQty = Number.isFinite(qty) ? qty : 1;
    return sum + normalizedQty;
  }, 0);
}

function resolveExpansionQtyFromLine(line, input = {}) {
  const fromBom = resolveExpansionQtyFromBom(line);
  if (fromBom > 0) return fromBom;
  const meter = toFiniteNumber(input?.meter);
  const expYes = Boolean(input?.expansionYes);
  if (expYes && Number.isFinite(meter) && meter > 30) return 1;
  return 0;
}

function buildOfferLineDebugSummary(line, input = {}) {
  const tapOffText = buildTapOffOfferText(line, input);
  const specialElementText = buildSpecialElementOfferText(line, input);
  const expansionQty = resolveExpansionQtyFromLine(line, input);
  const bom = Array.isArray(line?.bom) ? line.bom : [];
  return {
    lineNumber: safeString(line?.lineNumber),
    boxItemsCount: Array.isArray(input?.boxItems) ? input.boxItems.length : 0,
    bomBoxCount: resolveTapOffItemsFromLine(line, input).length,
    avbTekst: tapOffText,
    avbPris: resolveTapOffBoxPriceTotal(line, input),
    speTekst: specialElementText,
    spePris: resolveSpecialElementOfferPriceTotal(line, input),
    expansionQty,
    expansionBomRows: bom
      .filter(entry=>{
        const type = safeString(entry?.type || entry?.element_type || entry?.elementType).toLowerCase();
        const code = safeString(entry?.code).toLowerCase();
        const desc = safeString(entry?.desc || entry?.description || entry?.tekst).toLowerCase();
        return type === 'expansion_unit' || /\bexpans/.test(`${type} ${code} ${desc}`);
      })
      .map(entry=>({
        code: safeString(entry?.code),
        type: safeString(entry?.type || entry?.element_type || entry?.elementType),
        antall: entry?.antall ?? entry?.qty ?? entry?.quantity
      }))
  };
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
  const ipGrades = [];
  const tapOffTexts = [];
  const specialElementTexts = [];
  let brannElementTotal = 0;
  let tapOffTotal = 0;
  let tapOffPriceTotal = 0;
  let specialElementTotal = 0;
  let specialElementPriceTotal = 0;
  let expansionElementTotal = 0;
  let meterTotal = 0;
  let verticalAnglesTotal = 0;
  let horizontalAnglesTotal = 0;

  lines.forEach(line=>{
    const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};

    pushUnique(lineNumbers, line?.lineNumber);
    pushUnique(systems, input.series);
    pushUnique(ipGrades, resolveIpGradeFromSeries(input.series));
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
    const tapOffItems = resolveTapOffItemsFromLine(line, input);
    tapOffTotal += tapOffItems.reduce((sum, item)=>sum + Number(item.qty || 0), 0);
    tapOffPriceTotal += resolveTapOffOfferPriceTotal(line, input);
    const tapOffText = buildTapOffOfferText(line, input);
    pushUnique(tapOffTexts, tapOffText);
    const specialElementItems = resolveSpecialElementItemsFromLine(line, input);
    specialElementTotal += specialElementItems.reduce((sum, item)=>sum + Number(item.qty || 0), 0);
    specialElementPriceTotal += resolveSpecialElementOfferPriceTotal(line, input);
    pushUnique(specialElementTexts, buildSpecialElementOfferText(line, input));
    expansionElementTotal += resolveExpansionQtyFromLine(line, input);

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
    brannElementTotal: formatNoInteger(brannElementTotal),
    ipGrades: ipGrades.join(', '),
    expansionElementTotal: formatNoInteger(expansionElementTotal),
    tapOffTotal: formatNoInteger(tapOffTotal),
    tapOffPriceTotal: round2(tapOffPriceTotal),
    tapOffTexts: tapOffTexts.join(' | '),
    specialElementTotal: formatNoInteger(specialElementTotal),
    specialElementPriceTotal: round2(specialElementPriceTotal),
    specialElementTexts: specialElementTexts.join(' | ')
  };
}

function buildOfferPlaceholderValues(project, offerNumber, offerDate, revision = 0, userProfile = {}) {
  const safeProject = (project && typeof project === 'object') ? project : {};
  const projectName = safeString(safeProject.name);
  const customer = safeString(safeProject.customer);
  const contactPerson = safeString(safeProject.contactPerson || safeProject.contact);
  const customerAddress = safeString(safeProject.customerAddress || safeProject.address);
  const customerPostalPlace = safeString(safeProject.customerPostalPlace || safeProject.postalPlace);
  const contactPhone = safeString(safeProject.contactPhone || safeProject.phone);
  const profile = normalizeUserProfile(userProfile);
  const marketPayload = currentMarketPayloadForResponse();
  const usdNokRate = formatNoFxRate(marketPayload?.fx?.usdNok?.rate);
  const eurNokRate = formatNoFxRate(marketPayload?.fx?.eurNok?.rate);
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
  const tapOffPricePlaceholder = formatTapOffOfferPricePlaceholder(
    inputSummary.tapOffPriceTotal,
    inputSummary.tapOffTotal
  );
  const specialElementPricePlaceholder = formatTapOffOfferPricePlaceholder(
    inputSummary.specialElementPriceTotal,
    inputSummary.specialElementTotal
  );

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
    adresse: customerAddress,
    kunde_adresse: customerAddress,
    customer_address: customerAddress,
    ADRESSE: customerAddress,
    KUNDE_ADRESSE: customerAddress,
    postnummer_sted: customerPostalPlace,
    kunde_postnummer_sted: customerPostalPlace,
    postal_place: customerPostalPlace,
    POSTNUMMER_STED: customerPostalPlace,
    KUNDE_POSTNUMMER_STED: customerPostalPlace,
    telefon_kontaktperson: contactPhone,
    kontaktperson_telefon: contactPhone,
    contact_phone: contactPhone,
    TELEFON_KONTAKTPERSON: contactPhone,
    KONTAKTPERSON_TELEFON: contactPhone,
    bruker_navn: profile.name,
    bruker_epost: safeString(userProfile?.email),
    bruker_telefon: profile.phone,
    bruker_selskap: profile.company,
    bruker_stilling: profile.position,
    bruker_firma: profile.company,
    selger_navn: profile.name,
    selger_epost: safeString(userProfile?.email),
    selger_telefon: profile.phone,
    selger_selskap: profile.company,
    selger_stilling: profile.position,
    SELGER_NAVN: profile.name,
    SELGER_EPOST: safeString(userProfile?.email),
    SELGER_TELEFON: profile.phone,
    SELGER_SELSKAP: profile.company,
    SELGER_STILLING: profile.position,
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
    ipg: inputSummary.ipGrades,
    ip_grad: inputSummary.ipGrades,
    IP_GRAD: inputSummary.ipGrades,
    avb: inputSummary.tapOffTotal,
    avb_tekst: inputSummary.tapOffTexts,
    avtappingsbokser_tekst: inputSummary.tapOffTexts,
    AVTAPPINGSBOKSER_TEKST: inputSummary.tapOffTexts,
    avb_pris: tapOffPricePlaceholder,
    avb_pris_nok: tapOffPricePlaceholder,
    avb_sum: tapOffPricePlaceholder,
    avb_sum_nok: tapOffPricePlaceholder,
    avtappingsbokser_pris: tapOffPricePlaceholder,
    avtappingsbokser_pris_nok: tapOffPricePlaceholder,
    AVTAPPINGSBOKSER_PRIS: tapOffPricePlaceholder,
    spe: inputSummary.specialElementTotal,
    spe_tekst: inputSummary.specialElementTexts,
    spesialelement_tekst: inputSummary.specialElementTexts,
    spesialelementer_tekst: inputSummary.specialElementTexts,
    SPESIALELEMENT_TEKST: inputSummary.specialElementTexts,
    SPESIALELEMENTER_TEKST: inputSummary.specialElementTexts,
    spe_pris: specialElementPricePlaceholder,
    spe_pris_nok: specialElementPricePlaceholder,
    spe_sum: specialElementPricePlaceholder,
    spe_sum_nok: specialElementPricePlaceholder,
    spesialelement_pris: specialElementPricePlaceholder,
    spesialelement_pris_nok: specialElementPricePlaceholder,
    spesialelementer_pris: specialElementPricePlaceholder,
    spesialelementer_pris_nok: specialElementPricePlaceholder,
    SPESIALELEMENT_PRIS: specialElementPricePlaceholder,
    SPESIALELEMENTER_PRIS: specialElementPricePlaceholder,
    bre: inputSummary.brannElementTotal,
    exp: Number(inputSummary.expansionElementTotal) > 0
      ? `${inputSummary.expansionElementTotal} stk. Ekspansjonselement`
      : '',
    ekspansjonselement: Number(inputSummary.expansionElementTotal) > 0
      ? `${inputSummary.expansionElementTotal} stk. Ekspansjonselement`
      : '',
    EXPANSJONSELEMENT: Number(inputSummary.expansionElementTotal) > 0
      ? `${inputSummary.expansionElementTotal} stk. Ekspansjonselement`
      : '',
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
    tap_off_box_total_nok: formatNoCurrencyWithKr(totals.tapOffBoxTotal),
    tap_off_box_offer_total_nok: formatNoCurrencyWithKr(totals.tapOffOfferTotal),
    total_avtappingsbokser_nok: formatNoCurrencyWithKr(totals.tapOffOfferTotal),
    avb_total_nok: formatNoCurrencyWithKr(totals.tapOffOfferTotal),
    special_element_total_nok: formatNoCurrencyWithKr(totals.specialElementTotal),
    special_element_offer_total_nok: formatNoCurrencyWithKr(totals.specialElementOfferTotal),
    total_spesialelementer_nok: formatNoCurrencyWithKr(totals.specialElementOfferTotal),
    selected_addon_total_nok: formatNoCurrency(offerIncludedTotal),
    total_valgte_nok: formatNoCurrency(offerIncludedTotal),
    usd_nok_dagens: usdNokRate ? `USD ${usdNokRate}` : '',
    eur_nok_dagens: eurNokRate ? `EUR ${eurNokRate}` : '',
    USD_NOK_DAGENS: usdNokRate ? `USD ${usdNokRate}` : '',
    EUR_NOK_DAGENS: eurNokRate ? `EUR ${eurNokRate}` : ''
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
    const ipGrade = resolveIpGradeFromSeries(input.series);
    const expansionQty = resolveExpansionQtyFromLine(line, input);
    const expansionText = expansionQty > 0
      ? `${formatNoInteger(expansionQty)} stk. Ekspansjonselement`
      : '';

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
    const tapOffItems = resolveTapOffItemsFromLine(line, input);
    const tapOffText = buildTapOffOfferText(line, input);
    const tapOffTotalQty = tapOffItems.reduce((sum, item)=>sum + Number(item.qty || 0), 0);
    const tapOffPriceTotal = resolveTapOffOfferPriceTotal(line, input);
    const tapOffPricePlaceholder = formatTapOffOfferPricePlaceholder(tapOffPriceTotal, tapOffTotalQty);
    const specialElementItems = resolveSpecialElementItemsFromLine(line, input);
    const specialElementText = buildSpecialElementOfferText(line, input);
    const specialElementTotalQty = specialElementItems.reduce((sum, item)=>sum + Number(item.qty || 0), 0);
    const specialElementPriceTotal = resolveSpecialElementOfferPriceTotal(line, input);
    const specialElementPricePlaceholder = formatTapOffOfferPricePlaceholder(
      specialElementPriceTotal,
      specialElementTotalQty
    );
    const unitPricePlaceholders = buildUnitPricePlaceholders(line, input, lineTotals, fireBarrierPriceIndex);

    linePlaceholderSets.push({
      ...unitPricePlaceholders,
      lnr: lineNumber,
      linjenummer: lineNumber,
      sys: safeString(input.series),
      mtr: formatNoInteger(input.meter),
      vvk: formatNoPositiveInteger(input.v90_v ?? input.v90v),
      hvk: formatNoPositiveInteger(input.v90_h ?? input.v90h),
      amp,
      led: safeString(input.ledere),
      ste: normalizeElementLabel(input.startEl),
      sle: normalizeElementLabel(input.sluttEl),
      ipg: ipGrade,
      ip_grad: ipGrade,
      IP_GRAD: ipGrade,
      avb: formatNoPositiveInteger(tapOffTotalQty),
      avb_tekst: tapOffText,
      avtappingsbokser_tekst: tapOffText,
      AVTAPPINGSBOKSER_TEKST: tapOffText,
      avb_pris: tapOffPricePlaceholder,
      avb_pris_nok: tapOffPricePlaceholder,
      avb_sum: tapOffPricePlaceholder,
      avb_sum_nok: tapOffPricePlaceholder,
      avtappingsbokser_pris: tapOffPricePlaceholder,
      avtappingsbokser_pris_nok: tapOffPricePlaceholder,
      AVTAPPINGSBOKSER_PRIS: tapOffPricePlaceholder,
      spe: formatNoPositiveInteger(specialElementTotalQty),
      spe_tekst: specialElementText,
      spesialelement_tekst: specialElementText,
      spesialelementer_tekst: specialElementText,
      SPESIALELEMENT_TEKST: specialElementText,
      SPESIALELEMENTER_TEKST: specialElementText,
      spe_pris: specialElementPricePlaceholder,
      spe_pris_nok: specialElementPricePlaceholder,
      spe_sum: specialElementPricePlaceholder,
      spe_sum_nok: specialElementPricePlaceholder,
      spesialelement_pris: specialElementPricePlaceholder,
      spesialelement_pris_nok: specialElementPricePlaceholder,
      spesialelementer_pris: specialElementPricePlaceholder,
      spesialelementer_pris_nok: specialElementPricePlaceholder,
      SPESIALELEMENT_PRIS: specialElementPricePlaceholder,
      SPESIALELEMENTER_PRIS: specialElementPricePlaceholder,
      bre: hasBrannElements
        ? `${formatNoInteger(brannQty)} stk. Branngjennomforing EI 60/90/120`
        : '',
      exp: expansionText,
      ekspansjonselement: expansionText,
      EXPANSJONSELEMENT: expansionText,
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

  const replaceCharSplitPlaceholder = (input, key, escapedValue)=>{
    const normalizedKey = String(key || '');
    if (!normalizedKey) return input;
    const chars = [...normalizedKey];
    const bridge = '(?:\\s|</w:t>(?:(?!<w:t)[\\s\\S])*<w:t[^>]*>)*';
    const body = chars.map(ch=>`${escapeRegex(ch)}${bridge}`).join('');
    const open = `\\{${bridge}\\{${bridge}`;
    const close = `${bridge}\\}${bridge}\\}`;
    const pattern = new RegExp(`${open}${body}${close}`, 'g');
    return input.replace(pattern, escapedValue);
  };

  let output = String(xml);
  const orderedEntries = Object.entries(placeholders)
    .sort((a, b)=>String(b[0] || '').length - String(a[0] || '').length);
  orderedEntries.forEach(([key, rawValue])=>{
    const escapedValue = escapeXml(rawValue ?? '');
    const pattern = new RegExp(`\\{\\{\\s*${escapeRegex(key)}\\s*\\}\\}`, 'g');
    output = output.replace(pattern, escapedValue);
    output = replaceDelimiterSplitPlaceholder(output, key, escapedValue);
    output = replaceSplitThreeLetterPlaceholder(output, key, escapedValue);
    output = replaceCharSplitPlaceholder(output, key, escapedValue);
  });
  return output;
}

function hasUsablePlaceholderValue(placeholders, keys) {
  return keys.some(key=>safeString(placeholders?.[key]));
}

function xmlContainsPlaceholder(xml, key) {
  const source = String(xml || '');
  const directPattern = new RegExp(`\\{\\{\\s*${escapeRegex(key)}\\s*\\}\\}`, 'i');
  if (directPattern.test(source)) return true;

  const normalizedKey = String(key || '');
  if (!normalizedKey) return false;
  const bridge = '(?:\\s|</w:t>(?:(?!<w:t)[\\s\\S])*<w:t[^>]*>)*';
  const body = [...normalizedKey].map(ch=>`${escapeRegex(ch)}${bridge}`).join('');
  const open = `\\{${bridge}\\{${bridge}`;
  const close = `${bridge}\\}${bridge}\\}`;
  return new RegExp(`${open}${body}${close}`, 'i').test(source);
}

function getXmlElementRanges(xml, tagName) {
  const source = String(xml || '');
  const escapedTag = escapeRegex(tagName);
  const tagPattern = new RegExp(`<\\/?${escapedTag}\\b[^>]*>`, 'g');
  const ranges = [];
  const stack = [];
  let match;
  while ((match = tagPattern.exec(source)) !== null) {
    const token = match[0];
    if (token.startsWith(`</`)) {
      const start = stack.pop();
      if (start !== undefined && stack.length === 0) {
        ranges.push({ start, end: tagPattern.lastIndex });
      }
    } else if (!token.endsWith('/>')) {
      stack.push(match.index);
    }
  }
  return ranges;
}

function rangeInsideAny(range, outerRanges) {
  return outerRanges.some(outer=>range.start >= outer.start && range.end <= outer.end);
}

function removeRangesMatchingPlaceholders(xml, tagName, keys, options = {}) {
  let output = String(xml || '');
  const excludedRanges = Array.isArray(options.excludedRanges) ? options.excludedRanges : [];
  getXmlElementRanges(output, tagName).reverse().forEach(range=>{
    if (rangeInsideAny(range, excludedRanges)) return;
    const fragment = output.slice(range.start, range.end);
    if (!keys.some(key=>xmlContainsPlaceholder(fragment, key))) return;
    output = `${output.slice(0, range.start)}${output.slice(range.end)}`;
  });
  return output;
}

function removeXmlContainersForEmptyPlaceholders(xml, placeholders, specs) {
  let output = String(xml || '');
  specs.forEach(spec=>{
    const keys = Array.isArray(spec.keys) ? spec.keys : [];
    if (!keys.length || hasUsablePlaceholderValue(placeholders, keys)) return;
    if (spec.container === 'p') {
      output = removeRangesMatchingPlaceholders(output, 'w:tr', keys);
      const rowRanges = getXmlElementRanges(output, 'w:tr');
      output = removeRangesMatchingPlaceholders(output, 'w:p', keys, { excludedRanges: rowRanges });
      return;
    }
    const tagName = spec.container === 'tbl' ? 'w:tbl' : 'w:tr';
    output = removeRangesMatchingPlaceholders(output, tagName, keys);
  });
  getXmlElementRanges(output, 'w:tbl').reverse().forEach(range=>{
    const fragment = output.slice(range.start, range.end);
    if (/<w:tr\b/i.test(fragment)) return;
    output = `${output.slice(0, range.start)}${output.slice(range.end)}`;
  });
  return output;
}

function removeUnusedOfferPlaceholderContainers(xml, placeholders) {
  const specs = [
    { container: 'p', keys: ['ekspansjonselement', 'exp', 'EXPANSJONSELEMENT'] },
    { container: 'p', keys: ['bre'] },
    { container: 'p', keys: ['vvk'] },
    { container: 'p', keys: ['hvk'] },
    { container: 'p', keys: ['avb'] },
    { container: 'tr', keys: ['avb_tekst', 'avtappingsbokser_tekst', 'AVTAPPINGSBOKSER_TEKST', 'avb_pris', 'avb_pris_nok', 'avb_sum', 'avb_sum_nok', 'avtappingsbokser_pris', 'avtappingsbokser_pris_nok', 'AVTAPPINGSBOKSER_PRIS'] },
    { container: 'p', keys: ['spe'] },
    { container: 'tr', keys: ['spe_tekst', 'spesialelement_tekst', 'spesialelementer_tekst', 'SPESIALELEMENT_TEKST', 'SPESIALELEMENTER_TEKST', 'spe_pris', 'spe_pris_nok', 'spe_sum', 'spe_sum_nok', 'spesialelement_pris', 'spesialelement_pris_nok', 'spesialelementer_pris', 'spesialelementer_pris_nok', 'SPESIALELEMENT_PRIS', 'SPESIALELEMENTER_PRIS'] },
    { container: 'tr', keys: ['mtl', 'mtp'] },
    { container: 'p', keys: ['ttm', 'timer_totalt_montasje'] },
    { container: 'tr', keys: ['itl', 'itp'] },
    { container: 'p', keys: ['tti', 'timer_totalt_ingenior'] },
    { container: 'tr', keys: ['tol', 'top'] },
    { container: 'p', keys: ['tod'] }
  ];
  return removeXmlContainersForEmptyPlaceholders(xml, placeholders, specs);
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
      .map(placeholders=>replacePlaceholdersInXml(removeUnusedOfferPlaceholderContainers(blockTemplate, placeholders), placeholders))
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

async function generateOfferDocx(project, offerNumber, offerDate, revision = 0, userProfile = {}) {
  const offerTemplateFile = await resolveOfferTemplateFile();
  console.log(`[offer-template] Bruker mal: ${offerTemplateFile}`);
  const offerLines = Array.isArray(project?.lines) ? project.lines : [];
  const debugSummary = offerLines.map(line=>{
    const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};
    return buildOfferLineDebugSummary(line, input);
  });
  console.log(`[offer-template] Prosjekt "${safeString(project?.name)}" plassholderdata: ${JSON.stringify(debugSummary)}`);
  const {
    linePlaceholderSets,
    firePlaceholderSets,
    opphengPlaceholderSets
  } = await buildOfferLinePlaceholderValues(project);
  const placeholders = {
    ...buildOfferPlaceholderValues(project, offerNumber, offerDate, revision, userProfile),
    ...buildAggregateUnitPricePlaceholders(linePlaceholderSets)
  };
  const zip = new AdmZip(offerTemplateFile);
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
    const withoutUnusedPlaceholders = removeUnusedOfferPlaceholderContainers(withExpandedOpphengBlocks, placeholders);
    const replaced = replacePlaceholdersInXml(withoutUnusedPlaceholders, placeholders);
    if (replaced !== xml) {
      zip.updateFile(entry.entryName, Buffer.from(replaced, 'utf8'));
    }
  });

  return {
    buffer: zip.toBuffer(),
    templateFile: offerTemplateFile
  };
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

function extractObservationsFromSdmx(payload, pair) {
  const timeValues = payload?.data?.structure?.dimensions?.observation?.[0]?.values;
  const series = payload?.data?.dataSets?.[0]?.series;

  if (!Array.isArray(timeValues) || !isObject(series)) {
    throw new Error(`SDMX-respons for ${pair} mangler nodene som trengs`);
  }

  const observationsByDate = new Map();

  for (const seriesEntry of Object.values(series)) {
    const observations = seriesEntry?.observations;
    if (!isObject(observations)) continue;

    for (const [indexKey, rawValue] of Object.entries(observations)) {
      const idx = Number(indexKey);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const rawRate = Array.isArray(rawValue) ? rawValue[0] : rawValue;
      const parsedRate = parseRate(rawRate);
      const date = normalizeIsoDate(timeValues?.[idx]?.id);
      if (date && Number.isFinite(parsedRate)) {
        observationsByDate.set(date, { rate: parsedRate, date });
      }
    }
  }

  const observations = Array.from(observationsByDate.values())
    .sort((a, b) => a.date.localeCompare(b.date));
  if (!observations.length) {
    throw new Error(`Fant ikke gyldig siste datapunkt for ${pair}`);
  }

  return observations;
}

function subtractUtcDays(isoDate, days) {
  const parsed = Date.parse(`${isoDate}T00:00:00Z`);
  if (Number.isNaN(parsed)) return '';
  const next = new Date(parsed);
  next.setUTCDate(next.getUTCDate() - days);
  return next.toISOString().slice(0, 10);
}

function findObservationOnOrBefore(observations, targetDate) {
  let match = null;
  observations.forEach(point => {
    if (point.date <= targetDate) match = point;
  });
  return match;
}

function calculateFxChange(latest, comparePoint) {
  const current = Number(latest?.rate);
  const previous = Number(comparePoint?.rate);
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous <= 0) return null;
  return {
    rate: previous,
    percent: ((current - previous) / previous) * 100,
    fromDate: comparePoint.date,
    toDate: latest.date
  };
}

function buildFxChanges(observations, latest) {
  const weekTarget = subtractUtcDays(latest.date, 7);
  const monthTarget = subtractUtcDays(latest.date, 30);
  return {
    week: calculateFxChange(latest, findObservationOnOrBefore(observations, weekTarget)),
    month: calculateFxChange(latest, findObservationOnOrBefore(observations, monthTarget))
  };
}

function extractLatestObservationFromSdmx(payload, pair) {
  const observations = extractObservationsFromSdmx(payload, pair);
  const latest = observations[observations.length - 1];
  return {
    rate: latest.rate,
    date: latest.date,
    changes: buildFxChanges(observations, latest)
  };
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
      changes: usd.changes,
      source: 'Norges Bank'
    },
    eurNok: {
      pair: 'EUR/NOK',
      rate: eur.rate,
      date: eur.date,
      changes: eur.changes,
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
        changes: isObject(rawPoint.changes) ? rawPoint.changes : {},
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
  let offerTemplate = null;
  let runtimeData = {
    dir: RUNTIME_DATA_DIR,
    userAuthFile: USER_AUTH_FILE,
    userAuthFileExists: false,
    writable: false
  };
  try {
    const filePath = await resolveOfferTemplateFile();
    offerTemplate = {
      fileName: path.basename(filePath),
      filePath
    };
  } catch (err) {
    offerTemplate = {
      error: err?.message || 'Fant ikke tilbudsmal'
    };
  }
  try {
    await fs.mkdir(RUNTIME_DATA_DIR, { recursive: true });
    await fs.access(RUNTIME_DATA_DIR);
    runtimeData.writable = true;
  } catch (err) {
    runtimeData.writable = false;
    runtimeData.error = err?.message || 'Runtime data-mappen er ikke skrivbar';
  }
  try {
    await fs.access(USER_AUTH_FILE);
    runtimeData.userAuthFileExists = true;
  } catch (_err) {
    runtimeData.userAuthFileExists = false;
  }
  return res.json({
    ok: true,
    service: 'busbar-api',
    time: now,
    runtimeDataDir: RUNTIME_DATA_DIR,
    runtimeData,
    offerTemplate
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

app.get('/api/auth/microsoft/config', (req, res) => {
  if (!MICROSOFT_AUTH_TENANT_ID || !MICROSOFT_AUTH_CLIENT_ID) {
    return res.status(503).json({ enabled: false, error: 'Microsoft-innlogging er ikke konfigurert' });
  }
  const redirectUri = resolveMicrosoftRedirectUri(req);
  return res.json({
    enabled: true,
    tenantId: MICROSOFT_AUTH_TENANT_ID,
    clientId: MICROSOFT_AUTH_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${MICROSOFT_AUTH_TENANT_ID}`,
    redirectUri,
    scopes: MICROSOFT_AUTH_SCOPES
  });
});

app.post('/api/auth/microsoft/session', async (req, res) => {
  try {
    const microsoftUser = await verifyMicrosoftIdToken(req.body?.idToken);
    if (!microsoftUser) {
      return res.status(401).json({ error: 'Microsoft-innlogging kunne ikke valideres' });
    }
    const isOwner = await isMicrosoftAppOwner(microsoftUser.oid);

    const userRecord = await withUserAuthLock(async () => {
      const store = await readUserAuthStore();
      const existing = store.users[microsoftUser.email] || {};
      const now = new Date().toISOString();
      const existingProfile = normalizeUserProfile(existing.profile || {});
      const profile = {
        ...existingProfile,
        name: existingProfile.name || microsoftUser.name || microsoftUser.email,
        phone: existingProfile.phone || '',
        company: existingProfile.company || '',
        position: existingProfile.position || ''
      };
      store.users[microsoftUser.email] = {
        ...existing,
        email: microsoftUser.email,
        profile,
        passwordHash: safeString(existing.passwordHash),
        microsoft: {
          oid: microsoftUser.oid,
          tenantId: microsoftUser.tenantId,
          isOwner,
          linkedAt: existing?.microsoft?.linkedAt || now,
          lastLoginAt: now
        },
        createdAt: toIsoTimestamp(existing.createdAt, now),
        updatedAt: now
      };
      await writeUserAuthStore(store);
      return store.users[microsoftUser.email];
    });
    const isAdmin = resolveUserIsAdmin(userRecord);

    return res.json({
      email: userRecord.email,
      profile: userRecord.profile,
      isAdmin,
      token: createAuthToken(userRecord.email, { isAdmin })
    });
  } catch (err) {
    console.error('Microsoft-innlogging feilet', err);
    return res.status(500).json({ error: 'Kunne ikke fullføre Microsoft-innlogging' });
  }
});

app.post('/api/auth/register', async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    const password = normalizePassword(req.body?.password);
    const confirmPassword = normalizePassword(req.body?.confirmPassword);
    const profile = normalizeUserProfile(req.body?.profile || req.body);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }
    if (!isValidPassword(password)) {
      return res.status(400).json({ error: 'Passord må være minst 4 tegn' });
    }
    if (password !== confirmPassword) {
      return res.status(400).json({ error: 'Passordene er ikke like' });
    }
    if (!isCompleteUserProfile(profile)) {
      return res.status(400).json({ error: 'Fyll inn navn, telefon, selskap og stilling' });
    }

    const userRecord = await withUserAuthLock(async () => {
      const store = await readUserAuthStore();
      if (store.users[email]) {
        const err = new Error('Bruker finnes allerede');
        err.statusCode = 409;
        throw err;
      }
      const now = new Date().toISOString();
      store.users[email] = {
        email,
        profile,
        passwordHash: await hashPassword(password),
        createdAt: now,
        updatedAt: now
      };
      await writeUserAuthStore(store);
      return store.users[email];
    });

    return res.status(201).json({
      email: userRecord.email,
      profile: userRecord.profile,
      isAdmin: resolveUserIsAdmin(userRecord),
      token: createAuthToken(userRecord.email, { isAdmin: resolveUserIsAdmin(userRecord) })
    });
  } catch (err) {
    if (err?.statusCode === 409) {
      return res.status(409).json({ error: 'Bruker finnes allerede. Logg inn med passordet ditt.' });
    }
    console.error('Oppretting av bruker feilet', err);
    return res.status(500).json({ error: 'Kunne ikke opprette bruker' });
  }
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    const password = normalizePassword(req.body?.password);
    if (!isValidEmail(email) || !password) {
      return res.status(400).json({ error: 'Ugyldig e-post eller passord' });
    }

    const store = await readUserAuthStore();
    const userRecord = store.users[email];
    if (!userRecord || !(await verifyPassword(password, userRecord.passwordHash))) {
      return res.status(401).json({ error: 'Feil e-post eller passord' });
    }

    const isAdmin = resolveUserIsAdmin(userRecord);
    return res.json({
      email: userRecord.email,
      profile: userRecord.profile,
      isAdmin,
      token: createAuthToken(userRecord.email, { isAdmin })
    });
  } catch (err) {
    console.error('Innlogging feilet', err);
    return res.status(500).json({ error: 'Kunne ikke logge inn' });
  }
});

app.get('/api/user-projects', requireUserAuth, async (req, res) => {
  try {
    const email = normalizeEmail(req.query?.email);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }
    if (email !== req.userAuth.email && req.userAuth.isAdmin !== true) {
      return res.status(403).json({ error: 'Ingen tilgang til denne brukeren' });
    }
    const visibleRecord = await withOfferNumberLock(async ()=>{
      const [counterState, projectNumbersRaw] = await Promise.all([
        readJsonFile(OFFER_COUNTER_FILE, { years: {} }),
        readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {})
      ]);
      const projectNumbers = (projectNumbersRaw && typeof projectNumbersRaw === 'object')
        ? projectNumbersRaw
        : {};
      const archive = await withProjectArchiveLock(async ()=>{
        const state = await readProjectArchive();
        ensureProjectNumbersForArchive(state, projectNumbers, counterState);
        await writeProjectArchive(state);
        return state;
      });
      await Promise.all([
        writeJsonFile(OFFER_COUNTER_FILE, counterState),
        writeJsonFile(OFFER_PROJECT_NUMBERS_FILE, projectNumbers)
      ]);
      return buildVisibleProjectsForAuth(archive, req.userAuth);
    });
    return res.json({
      email: visibleRecord.email,
      updatedAt: visibleRecord.updatedAt,
      ownerEmails: visibleRecord.ownerEmails,
      isAdmin: req.userAuth.isAdmin === true,
      projects: visibleRecord.projects
    });
  } catch (err) {
    console.error('Henting av brukerprosjekter feilet', err);
    return res.status(500).json({ error: 'Kunne ikke hente prosjekter' });
  }
});

app.post('/api/user-projects/sync', requireUserAuth, async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }
    if (email !== req.userAuth.email && req.userAuth.isAdmin !== true) {
      return res.status(403).json({ error: 'Ingen tilgang til denne brukeren' });
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

    const nextUserRecord = await withOfferNumberLock(async () => {
      const [counterState, projectNumbersRaw] = await Promise.all([
        readJsonFile(OFFER_COUNTER_FILE, { years: {} }),
        readJsonFile(OFFER_PROJECT_NUMBERS_FILE, {})
      ]);
      const projectNumbers = (projectNumbersRaw && typeof projectNumbersRaw === 'object')
        ? projectNumbersRaw
        : {};
      const record = await withProjectArchiveLock(async () => {
        const archive = await readProjectArchive();
        normalizedProjects.forEach(project=>{
          ensureProjectNumber(project, projectNumbers, counterState);
        });
        if (req.userAuth.isAdmin === true) {
          const grouped = new Map();
          normalizedProjects.forEach(project => {
            const ownerEmail = isValidEmail(project.projectOwnerEmail) ? project.projectOwnerEmail : email;
            if (!grouped.has(ownerEmail)) grouped.set(ownerEmail, []);
            grouped.get(ownerEmail).push({
              ...project,
              projectOwnerEmail: ownerEmail
            });
          });
          grouped.forEach((projects, ownerEmail) => {
            archive.users[ownerEmail] = {
              email: ownerEmail,
              updatedAt,
              projects
            };
          });
          const visibleOwnerEmails = Array.isArray(req.body?.ownerEmails)
            ? req.body.ownerEmails.map(normalizeEmail).filter(isValidEmail)
            : [];
          visibleOwnerEmails.forEach(ownerEmail => {
            if (grouped.has(ownerEmail)) return;
            archive.users[ownerEmail] = {
              email: ownerEmail,
              updatedAt,
              projects: []
            };
          });
        } else {
          archive.users[email] = {
            email,
            updatedAt,
            projects: normalizedProjects.map(project => ({
              ...project,
              projectOwnerEmail: email
            }))
          };
        }
        await writeProjectArchive(archive);
        return buildVisibleProjectsForAuth(archive, req.userAuth);
      });
      await Promise.all([
        writeJsonFile(OFFER_COUNTER_FILE, counterState),
        writeJsonFile(OFFER_PROJECT_NUMBERS_FILE, projectNumbers)
      ]);
      return record;
    });

    return res.json({
      email: nextUserRecord.email,
      updatedAt: nextUserRecord.updatedAt,
      ownerEmails: nextUserRecord.ownerEmails,
      isAdmin: req.userAuth.isAdmin === true,
      projects: nextUserRecord.projects
    });
  } catch (err) {
    console.error('Synk av brukerprosjekter feilet', err);
    return res.status(500).json({ error: 'Kunne ikke synkronisere prosjekter' });
  }
});

async function upsertGlobalCustomerRecord(body) {
  const originalCustomer = safeString(body?.originalCustomer);
  const originalContactPerson = safeString(body?.originalContactPerson);
  const customerName = safeString(body?.customer);
  const address = safeString(body?.address);
  const postalPlace = safeString(body?.postalPlace);
  const segment = safeString(body?.segment);
  const customerResponsible = safeString(body?.customerResponsible || body?.responsible);
  const contactPerson = safeString(body?.contactPerson);
  const phone = safeString(body?.phone);
  const email = safeString(body?.email);
  const isContactUpdate = Boolean(originalContactPerson || contactPerson || phone || email);
  if (!customerName) {
    const err = new Error('Kunde mangler');
    err.statusCode = 400;
    throw err;
  }

  const result = await withCustomerDatabaseLock(async () => {
    const database = await readCustomerDatabase();
    const customers = database.customers;
    const findCustomerIndex = name => customers.findIndex(customer => normalizeLookupKey(customer.name) === normalizeLookupKey(name));
    let customerIndex = originalCustomer ? findCustomerIndex(originalCustomer) : -1;
    if (customerIndex < 0) customerIndex = findCustomerIndex(customerName);
    let customer = customerIndex >= 0 ? customers[customerIndex] : null;
    if (!customer) {
      customer = {
        id: generateRecordId('customer'),
        name: customerName,
        address,
        postalPlace,
        segment,
        customerResponsible,
        contacts: []
      };
      customers.push(customer);
    }
    customer.name = customerName;
    customer.address = address;
    customer.postalPlace = postalPlace;
    customer.segment = segment;
    customer.customerResponsible = customerResponsible;
    if (isContactUpdate) {
      const findContactIndex = name => customer.contacts.findIndex(contact => normalizeLookupKey(contact.name) === normalizeLookupKey(name));
      let contactIndex = originalContactPerson ? findContactIndex(originalContactPerson) : -1;
      if (contactIndex < 0) contactIndex = findContactIndex(contactPerson);
      let contact = contactIndex >= 0 ? customer.contacts[contactIndex] : null;
      if (!contact) {
        contact = { id: generateRecordId('contact'), name: contactPerson, phone, email };
        customer.contacts.push(contact);
      }
      contact.name = contactPerson;
      contact.phone = phone;
      contact.email = email;
    }
    await writeCustomerDatabase({ customers });

    let updatedProjects = 0;
    const archive = await readProjectArchive();
    Object.values(archive.users || {}).forEach(user => {
      (Array.isArray(user.projects) ? user.projects : []).forEach(project => {
        const customerMatches = originalCustomer
          ? normalizeLookupKey(project.customer) === normalizeLookupKey(originalCustomer)
          : false;
        if (!customerMatches) return;
        if (!isContactUpdate) {
          project.customer = customerName;
          project.customerAddress = address;
          project.customerPostalPlace = postalPlace;
        }
        if (isContactUpdate && normalizeLookupKey(project.contactPerson) === normalizeLookupKey(originalContactPerson || contactPerson)) {
          project.contactPerson = contactPerson;
          project.contactPhone = phone;
        }
        project.updatedAt = new Date().toISOString();
        updatedProjects += 1;
      });
    });
    if (originalCustomer) {
      await writeProjectArchive(archive);
    }
    return { updatedProjects };
  });

  const merged = await buildMergedCustomerDatabase();
  return {
    ...result,
    customers: merged.customers
  };
}

async function deleteGlobalCustomerRecord(body) {
  const customerName = safeString(body?.customer);
  const contactPerson = safeString(body?.contactPerson);
  if (!customerName) {
    const err = new Error('Kunde mangler');
    err.statusCode = 400;
    throw err;
  }

  const result = await withCustomerDatabaseLock(async () => {
    const database = await readCustomerDatabase();
    const customers = database.customers;
    const customerIndex = customers.findIndex(customer => normalizeLookupKey(customer.name) === normalizeLookupKey(customerName));
    if (customerIndex >= 0) {
      if (contactPerson) {
        customers[customerIndex].contacts = customers[customerIndex].contacts.filter(
          contact => normalizeLookupKey(contact.name) !== normalizeLookupKey(contactPerson)
        );
      } else {
        customers.splice(customerIndex, 1);
      }
      await writeCustomerDatabase({ customers });
    }

    let updatedProjects = 0;
    const archive = await readProjectArchive();
    Object.values(archive.users || {}).forEach(user => {
      (Array.isArray(user.projects) ? user.projects : []).forEach(project => {
        if (normalizeLookupKey(project.customer) !== normalizeLookupKey(customerName)) return;
        if (contactPerson) {
          if (normalizeLookupKey(project.contactPerson) !== normalizeLookupKey(contactPerson)) return;
          project.contactPerson = '';
          project.contactPhone = '';
        } else {
          project.customer = '';
          project.customerAddress = '';
          project.customerPostalPlace = '';
          project.contactPerson = '';
          project.contactPhone = '';
        }
        project.updatedAt = new Date().toISOString();
        updatedProjects += 1;
      });
    });
    await writeProjectArchive(archive);
    return { updatedProjects };
  });

  const merged = await buildMergedCustomerDatabase();
  return {
    ...result,
    customers: merged.customers
  };
}

app.get('/api/customer-database', requireUserAuth, async (_req, res) => {
  try {
    const database = await buildMergedCustomerDatabase();
    return res.json(database);
  } catch (err) {
    console.error('Henting av kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke hente kundedatabase' });
  }
});

app.post('/api/customer-database/upsert', requireUserAuth, requireMicrosoftOwnerAuth, async (req, res) => {
  try {
    const result = await upsertGlobalCustomerRecord(req.body);
    return res.json(result);
  } catch (err) {
    if (err?.statusCode === 400) {
      return res.status(400).json({ error: err.message });
    }
    console.error('Oppdatering av global kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke oppdatere kundedatabase' });
  }
});

app.post('/api/customer-database/delete', requireUserAuth, requireMicrosoftOwnerAuth, async (req, res) => {
  try {
    const result = await deleteGlobalCustomerRecord(req.body);
    return res.json(result);
  } catch (err) {
    if (err?.statusCode === 400) {
      return res.status(400).json({ error: err.message });
    }
    console.error('Sletting i global kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke slette fra kundedatabase' });
  }
});

app.get('/api/admin/customer-database', requireAdminAuth, async (_req, res) => {
  try {
    const database = await buildMergedCustomerDatabase();
    return res.json({
      generatedAt: new Date().toISOString(),
      customers: database.customers
    });
  } catch (err) {
    console.error('Henting av admin kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke hente kundedatabase' });
  }
});

app.post('/api/admin/customer-database/upsert', requireAdminAuth, async (req, res) => {
  try {
    const originalCustomer = safeString(req.body?.originalCustomer);
    const originalContactPerson = safeString(req.body?.originalContactPerson);
  const customerName = safeString(req.body?.customer);
    const address = safeString(req.body?.address);
    const postalPlace = safeString(req.body?.postalPlace);
    const segment = safeString(req.body?.segment);
    const customerResponsible = safeString(req.body?.customerResponsible || req.body?.responsible);
    const contactPerson = safeString(req.body?.contactPerson);
  const phone = safeString(req.body?.phone);
  const email = safeString(req.body?.email);
  const isContactUpdate = Boolean(originalContactPerson || contactPerson || phone || email);
    if (!customerName) {
      return res.status(400).json({ error: 'Kunde mangler' });
    }

    const result = await withCustomerDatabaseLock(async () => {
      const database = await readCustomerDatabase();
      const customers = database.customers;
      const findCustomerIndex = name => customers.findIndex(customer => normalizeLookupKey(customer.name) === normalizeLookupKey(name));
      let customerIndex = originalCustomer ? findCustomerIndex(originalCustomer) : -1;
      if (customerIndex < 0) customerIndex = findCustomerIndex(customerName);
      let customer = customerIndex >= 0 ? customers[customerIndex] : null;
      if (!customer) {
        customer = {
          id: generateRecordId('customer'),
          name: customerName,
          address,
          postalPlace,
          segment,
          customerResponsible,
          contacts: []
        };
        customers.push(customer);
      }
      customer.name = customerName;
      customer.address = address;
      customer.postalPlace = postalPlace;
      customer.segment = segment;
      customer.customerResponsible = customerResponsible;
      if (isContactUpdate) {
        const findContactIndex = name => customer.contacts.findIndex(contact => normalizeLookupKey(contact.name) === normalizeLookupKey(name));
        let contactIndex = originalContactPerson ? findContactIndex(originalContactPerson) : -1;
        if (contactIndex < 0) contactIndex = findContactIndex(contactPerson);
        let contact = contactIndex >= 0 ? customer.contacts[contactIndex] : null;
        if (!contact) {
          contact = { id: generateRecordId('contact'), name: contactPerson, phone, email };
          customer.contacts.push(contact);
        }
        contact.name = contactPerson;
        contact.phone = phone;
        contact.email = email;
      }
      await writeCustomerDatabase({ customers });

      let updatedProjects = 0;
      const archive = await readProjectArchive();
      Object.values(archive.users || {}).forEach(user => {
        (Array.isArray(user.projects) ? user.projects : []).forEach(project => {
          const customerMatches = originalCustomer
            ? normalizeLookupKey(project.customer) === normalizeLookupKey(originalCustomer)
            : false;
          if (!customerMatches) return;
          if (!isContactUpdate) {
            project.customer = customerName;
            project.customerAddress = address;
            project.customerPostalPlace = postalPlace;
          }
          if (isContactUpdate && normalizeLookupKey(project.contactPerson) === normalizeLookupKey(originalContactPerson || contactPerson)) {
            project.contactPerson = contactPerson;
            project.contactPhone = phone;
          }
          project.updatedAt = new Date().toISOString();
          updatedProjects += 1;
        });
      });
      if (originalCustomer) {
        await writeProjectArchive(archive);
      }
      return { updatedProjects };
    });

    const merged = await buildMergedCustomerDatabase();
    return res.json({
      ...result,
      customers: merged.customers
    });
  } catch (err) {
    console.error('Oppdatering av kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke oppdatere kundedatabase' });
  }
});

app.post('/api/admin/customer-database/delete', requireAdminAuth, async (req, res) => {
  try {
    const customerName = safeString(req.body?.customer);
    const contactPerson = safeString(req.body?.contactPerson);
    if (!customerName) {
      return res.status(400).json({ error: 'Kunde mangler' });
    }

    const result = await withCustomerDatabaseLock(async () => {
      const database = await readCustomerDatabase();
      const customers = database.customers;
      const customerIndex = customers.findIndex(customer => normalizeLookupKey(customer.name) === normalizeLookupKey(customerName));
      if (customerIndex >= 0) {
        if (contactPerson) {
          customers[customerIndex].contacts = customers[customerIndex].contacts.filter(
            contact => normalizeLookupKey(contact.name) !== normalizeLookupKey(contactPerson)
          );
        } else {
          customers.splice(customerIndex, 1);
        }
        await writeCustomerDatabase({ customers });
      }

      let updatedProjects = 0;
      const archive = await readProjectArchive();
      Object.values(archive.users || {}).forEach(user => {
        (Array.isArray(user.projects) ? user.projects : []).forEach(project => {
          if (normalizeLookupKey(project.customer) !== normalizeLookupKey(customerName)) return;
          if (contactPerson) {
            if (normalizeLookupKey(project.contactPerson) !== normalizeLookupKey(contactPerson)) return;
            project.contactPerson = '';
            project.contactPhone = '';
          } else {
            project.customer = '';
            project.customerAddress = '';
            project.customerPostalPlace = '';
            project.contactPerson = '';
            project.contactPhone = '';
          }
          project.updatedAt = new Date().toISOString();
          updatedProjects += 1;
        });
      });
      await writeProjectArchive(archive);
      return { updatedProjects };
    });

    const merged = await buildMergedCustomerDatabase();
    return res.json({
      ...result,
      customers: merged.customers
    });
  } catch (err) {
    console.error('Sletting i kundedatabase feilet', err);
    return res.status(500).json({ error: 'Kunne ikke slette fra kundedatabase' });
  }
});

app.get('/api/admin/project-overview', requireAdminAuth, async (req, res) => {
  try {
    const archive = await readProjectArchive();
    const authStore = await readUserAuthStore();
    const userEmails = new Set([
      ...Object.keys(archive.users || {}),
      ...Object.keys(authStore.users || {})
    ]);
    const users = Array.from(userEmails).map(email => {
      const user = archive.users[email] || { email, updatedAt: null, projects: [] };
      const authRecord = authStore.users[email] || null;
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
        profile: authRecord?.profile || null,
        registered: Boolean(authRecord),
        hasPassword: Boolean(authRecord?.passwordHash),
        microsoftLinked: Boolean(authRecord?.microsoft?.oid),
        isAdmin: resolveUserIsAdmin(authRecord || { email }),
        authUpdatedAt: authRecord?.updatedAt || null,
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

app.post('/api/admin/users/profile', requireAdminAuth, async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    const profile = normalizeUserProfile(req.body?.profile || req.body);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }

    const userRecord = await withUserAuthLock(async () => {
      const store = await readUserAuthStore();
      const existing = store.users[email] || {};
      const now = new Date().toISOString();
      store.users[email] = {
        ...existing,
        email,
        profile,
        passwordHash: safeString(existing.passwordHash),
        ...(existing.microsoft?.oid ? { microsoft: existing.microsoft } : {}),
        createdAt: toIsoTimestamp(existing.createdAt, now),
        updatedAt: now
      };
      await writeUserAuthStore(store);
      return store.users[email];
    });

    return res.json({
      email: userRecord.email,
      profile: userRecord.profile,
      updatedAt: userRecord.updatedAt
    });
  } catch (err) {
    console.error('Oppdatering av brukerprofil feilet', err);
    return res.status(500).json({ error: 'Kunne ikke oppdatere brukerprofil' });
  }
});

app.post('/api/admin/users/delete', requireAdminAuth, async (req, res) => {
  try {
    const email = normalizeEmail(req.body?.email);
    if (!isValidEmail(email)) {
      return res.status(400).json({ error: 'Ugyldig e-post' });
    }

    let deletedAuth = false;
    let deletedProjects = 0;
    let offerCleanup = { removedProjectNumbers: 0, removedRevisions: 0 };

    deletedAuth = await withUserAuthLock(async () => {
      const store = await readUserAuthStore();
      if (!store.users[email]) return false;
      delete store.users[email];
      await writeUserAuthStore(store);
      return true;
    });

    const projectsToDelete = await withProjectArchiveLock(async () => {
      const archive = await readProjectArchive();
      const user = archive.users[email];
      const projects = Array.isArray(user?.projects) ? user.projects : [];
      if (archive.users[email]) {
        delete archive.users[email];
        await writeProjectArchive(archive);
      }
      return projects;
    });
    deletedProjects = projectsToDelete.length;
    offerCleanup = await removeOfferMetadataForProjects(projectsToDelete);

    return res.json({
      email,
      deleted: deletedAuth || deletedProjects > 0,
      deletedAuth,
      deletedProjects,
      ...offerCleanup
    });
  } catch (err) {
    console.error('Sletting av bruker feilet', err);
    return res.status(500).json({ error: 'Kunne ikke slette bruker' });
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

app.post('/api/generate-offer', requireUserAuth, async (req, res) => {
  try {
    const project = (req.body && typeof req.body === 'object') ? req.body.project : null;
    if (!project || typeof project !== 'object') {
      return res.status(400).json({ error: 'Mangler prosjektdata' });
    }
    const authStore = await readUserAuthStore();
    const userRecord = authStore.users[req.userAuth.email];
    if (!userRecord) {
      return res.status(401).json({ error: 'Logg inn for å generere tilbud' });
    }

    const now = new Date();
    const { offerNumber, revision } = await allocateOfferIdentity(project, now);
    project.projectNumber = offerNumber;
    const projectOwnerEmail = resolveWritableProjectOwnerEmail(project, req.userAuth);
    await persistProjectNumberForUser(projectOwnerEmail, project.id, offerNumber);
    const generated = await generateOfferDocx(project, offerNumber, now, revision, {
      email: userRecord.email,
      ...userRecord.profile
    });
    const projectName = sanitizeFileName(project.name || 'prosjekt');
    const fileName = `Tilbud-${projectName}-${offerNumber}-${revision}.docx`;
    const encodedFileName = encodeURIComponent(fileName);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"; filename*=UTF-8''${encodedFileName}`);
    res.setHeader('X-Offer-Number', offerNumber);
    res.setHeader('X-Offer-Revision', String(revision));
    res.setHeader('X-Offer-Filename', fileName);
    res.setHeader('X-Offer-Template', path.basename(generated.templateFile));
    res.setHeader('X-Offer-Template-Path', generated.templateFile);
    res.status(200).send(generated.buffer);
  } catch (err) {
    console.error('Tilbudsgenerering feilet', err);
    if (err && err.code === 'ENOENT') {
      return res.status(500).json({ error: 'Fant ikke tilbudsmalen i server/templates/tilbud' });
    }
    res.status(500).json({ error: 'Kunne ikke generere tilbud' });
  }
});

app.get('/api/offer-status', requireUserAuth, async (req, res) => {
  try {
    const archive = await readProjectArchive();
    const visible = buildVisibleProjectsForAuth(archive, req.userAuth);
    const offers = await getOfferStatusForProjects(visible.projects);
    res.json({ offers });
  } catch (err) {
    console.error('Henting av tilbudsstatus feilet', err);
    res.status(500).json({ error: 'Kunne ikke hente tilbudsstatus' });
  }
});

app.post('/api/generate-offer-latest', requireUserAuth, async (req, res) => {
  try {
    const project = (req.body && typeof req.body === 'object') ? req.body.project : null;
    if (!project || typeof project !== 'object') {
      return res.status(400).json({ error: 'Mangler prosjektdata' });
    }
    const authStore = await readUserAuthStore();
    const userRecord = authStore.users[req.userAuth.email];
    if (!userRecord) {
      return res.status(401).json({ error: 'Logg inn for å åpne tilbud' });
    }

    const status = (await getOfferStatusForProjects([project]))[0];
    if (!status?.hasOffer) {
      return res.status(404).json({ error: 'Prosjektet har ikke et generert tilbud ennå' });
    }

    const revision = Number(status.revision);
    const offerNumber = status.offerNumber;
    project.projectNumber = offerNumber;
    const generated = await generateOfferDocx(project, offerNumber, new Date(), revision, {
      email: userRecord.email,
      ...userRecord.profile
    });
    const projectName = sanitizeFileName(project.name || 'prosjekt');
    const fileName = `Tilbud-${projectName}-${offerNumber}-${revision}.docx`;
    const encodedFileName = encodeURIComponent(fileName);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"; filename*=UTF-8''${encodedFileName}`);
    res.setHeader('X-Offer-Number', offerNumber);
    res.setHeader('X-Offer-Revision', String(revision));
    res.setHeader('X-Offer-Filename', fileName);
    res.setHeader('X-Offer-Template', path.basename(generated.templateFile));
    res.setHeader('X-Offer-Template-Path', generated.templateFile);
    res.status(200).send(generated.buffer);
  } catch (err) {
    console.error('Åpning av siste tilbud feilet', err);
    if (err && err.code === 'ENOENT') {
      return res.status(500).json({ error: 'Fant ikke tilbudsmalen i server/templates/tilbud' });
    }
    res.status(500).json({ error: 'Kunne ikke åpne siste tilbud' });
  }
});

const staticDir = path.resolve(__dirname, '..');
app.use(express.static(staticDir, {
  setHeaders(res, filePath) {
    const ext = path.extname(filePath).toLowerCase();
    if (['.html', '.js', '.css'].includes(ext)) {
      res.setHeader('Cache-Control', 'no-store');
    }
  }
}));

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
  migrateProjectNumbersForAllUsersOnStartup()
    .then(stats=>{
      console.log(`[project-number-migration] users=${stats.users} projects=${stats.projects} added=${stats.added} preserved=${stats.preserved}`);
    })
    .catch(err=>{
      console.error('[project-number-migration] Feilet', err);
    });
  migrateProjectsBetweenUsersOnStartup()
    .then(result=>{
      if (result?.moved) {
        console.log(`[project-transfer] Flyttet ${result.moved} prosjekt(er) fra ${result.source} til ${result.target}`);
      } else if (result?.sourceRemoved) {
        console.log(`[project-transfer] Fjernet tom kildebruker ${result.source}`);
      }
    })
    .catch(err=>{
      console.error('[project-transfer] Feilet', err);
    });
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
