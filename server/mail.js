require('dotenv').config();
const path = require('path');
const fs = require('fs/promises');
const express = require('express');
const nodemailer = require('nodemailer');
const AdmZip = require('adm-zip');
const { ClientSecretCredential } = require('@azure/identity');

const DEFAULT_MARKET_FILE = path.resolve(__dirname, '..', 'data', 'market-data.json');
const OFFER_TEMPLATE_FILE = path.resolve(
  __dirname,
  'templates',
  'tilbud',
  'tilbud-stromskinner-template.docx'
);
const OFFER_COUNTER_FILE = path.resolve(__dirname, '..', 'data', 'offer-sequence.json');
const OFFER_LINE_BLOCK_START_TOKEN = '__BUSBAR_LINE_BLOCK_START__';
const OFFER_LINE_BLOCK_END_TOKEN = '__BUSBAR_LINE_BLOCK_END__';
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

fs.access(MARKET_DATA_FILE).catch(err=>{
  console.warn(`[market-data] Datafil utilgjengelig (${MARKET_DATA_FILE}): ${err.message}`);
});

const app = express();
app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(204);
  }
  return next();
});

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

async function allocateOfferNumber(date = new Date()) {
  return withOfferNumberLock(async ()=>{
    const state = await readJsonFile(OFFER_COUNTER_FILE, { years: {} });
    const years = (state && typeof state === 'object' && state.years && typeof state.years === 'object')
      ? state.years
      : {};
    const year = String(getOfferYear(date));
    const previous = Number(years[year]);
    const next = (Number.isInteger(previous) && previous >= 1000) ? previous + 1 : 1001;
    years[year] = next;
    await writeJsonFile(OFFER_COUNTER_FILE, { years });
    return `${year}-${next}`;
  });
}

function resolveLineSelectedAddonTotal(line) {
  const direct = toFiniteNumber(line?.selectedAddonTotal ?? line?.totals?.selectedAddonTotal);
  if (Number.isFinite(direct)) return round2(direct);

  const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
    ? line.totals
    : {};
  const baseTotal = toFiniteNumber(lineTotals.totalExMontasje);
  if (!Number.isFinite(baseTotal)) return NaN;

  const savedFlags = line?.selectedAddonConfig || lineTotals.selectedAddonConfig || null;
  const includeMontasje = savedFlags ? Boolean(savedFlags.includeMontasje) : true;
  const includeEngineering = savedFlags ? Boolean(savedFlags.includeEngineering) : true;
  const includeOppheng = savedFlags ? Boolean(savedFlags.includeOppheng) : true;
  const montasjeTotal = toFiniteNumber(lineTotals.totalInclMontasje);
  const engineeringTotal = toFiniteNumber(lineTotals.totalInclEngineering);
  const opphengTotal = toFiniteNumber(lineTotals.total);

  let total = baseTotal;
  if (includeMontasje && Number.isFinite(montasjeTotal)) total += montasjeTotal;
  if (includeEngineering && Number.isFinite(engineeringTotal)) total += engineeringTotal;
  if (includeOppheng && Number.isFinite(opphengTotal)) total += opphengTotal;
  return round2(total);
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
    selectedAddonTotal: 0
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
    add('selectedAddonTotal', resolveLineSelectedAddonTotal(line));
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

function buildOfferPlaceholderValues(project, offerNumber, offerDate) {
  const safeProject = (project && typeof project === 'object') ? project : {};
  const projectName = safeString(safeProject.name);
  const customer = safeString(safeProject.customer);
  const contactPerson = safeString(safeProject.contactPerson || safeProject.contact);
  const { lines, totals } = aggregateProjectOfferTotals(safeProject);
  const inputSummary = collectProjectInputSummary(lines);

  const placeholders = {
    tilbud_nr: offerNumber,
    tilbudsdato: formatOfferDate(offerDate),
    dato: formatOfferDate(offerDate),
    prosjektnavn: projectName,
    prosjekt: projectName,
    kunde: customer,
    customer: customer,
    kontaktperson: contactPerson,
    lss: '',
    linjer_start: '',
    lse: '',
    linjer_slutt: '',
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
    stv: formatNoCurrency(totals.selectedAddonTotal),
    tmo: formatNoCurrency(totals.totalInclMontasje),
    tin: formatNoCurrency(totals.totalInclEngineering),
    top: formatNoCurrency(totals.oppheng),
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
    total_ex_montasje_nok: formatNoCurrency(totals.totalExMontasje),
    montasje_nok: formatNoCurrency(totals.montasje),
    montasje_margin_nok: formatNoCurrency(totals.montasjeMargin),
    total_incl_montasje_nok: formatNoCurrency(totals.totalInclMontasje),
    engineering_nok: formatNoCurrency(totals.engineering),
    engineering_margin_nok: formatNoCurrency(totals.engineeringMargin),
    total_incl_engineering_nok: formatNoCurrency(totals.totalInclEngineering),
    oppheng_nok: formatNoCurrency(totals.oppheng),
    selected_addon_total_nok: formatNoCurrency(totals.selectedAddonTotal),
    total_valgte_nok: formatNoCurrency(totals.selectedAddonTotal)
  };

  return placeholders;
}

function buildOfferLinePlaceholderValues(project) {
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  return lines.map((line, index)=>{
    const input = (line && typeof line === 'object' && line.inputs && typeof line.inputs === 'object')
      ? line.inputs
      : {};
    const lineTotals = (line && typeof line === 'object' && line.totals && typeof line.totals === 'object')
      ? line.totals
      : {};

    const ampNum = toFiniteNumber(input.ampere ?? input.amp);
    const amp = Number.isFinite(ampNum)
      ? String(Math.round(ampNum))
      : safeString(input.ampere ?? input.amp);

    return {
      lnr: safeString(line?.lineNumber || String(index + 1)),
      linjenummer: safeString(line?.lineNumber || String(index + 1)),
      sys: safeString(input.series),
      mtr: formatNoInteger(input.meter),
      vvk: formatNoInteger(input.v90_v ?? input.v90v),
      hvk: formatNoInteger(input.v90_h ?? input.v90h),
      amp,
      led: safeString(input.ledere),
      ste: normalizeElementLabel(input.startEl),
      sle: normalizeElementLabel(input.sluttEl),
      avb: formatNoInteger(input.boxQty),
      bre: formatNoInteger(input.fbQty ?? input.fireBarrierQty),
      stv: formatNoCurrency(resolveLineSelectedAddonTotal(line)),
      tmo: formatNoCurrency(lineTotals.totalInclMontasje),
      tin: formatNoCurrency(lineTotals.totalInclEngineering),
      top: formatNoCurrency(lineTotals?.oppheng?.cost ?? lineTotals.total),
      ttm: formatNoIntegerUp(lineTotals?.montasje?.totalHours),
      tti: formatNoIntegerUp(lineTotals?.engineering?.totalHours),
      aop: formatNoInteger(lineTotals?.oppheng?.pieceCount),
      timer_totalt_montasje: formatNoIntegerUp(lineTotals?.montasje?.totalHours),
      timer_totalt_ingenior: formatNoIntegerUp(lineTotals?.engineering?.totalHours),
      antall_oppheng: formatNoInteger(lineTotals?.oppheng?.pieceCount)
    };
  });
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

function expandLineRepeatBlocks(xml, linePlaceholderSets) {
  const markerReplaced = replacePlaceholdersInXml(xml, {
    linjer_start: OFFER_LINE_BLOCK_START_TOKEN,
    lss: OFFER_LINE_BLOCK_START_TOKEN,
    linjer_slutt: OFFER_LINE_BLOCK_END_TOKEN,
    lse: OFFER_LINE_BLOCK_END_TOKEN
  });

  const tokenPattern = new RegExp(
    `${escapeRegex(OFFER_LINE_BLOCK_START_TOKEN)}([\\s\\S]*?)${escapeRegex(OFFER_LINE_BLOCK_END_TOKEN)}`,
    'g'
  );
  const rawPattern = /\{\{\s*(?:lss|linjer_start)\s*\}\}([\s\S]*?)\{\{\s*(?:lse|linjer_slutt)\s*\}\}/g;
  const renderBlock = (blockTemplate)=>{
    if (!Array.isArray(linePlaceholderSets) || linePlaceholderSets.length === 0) return '';
    return linePlaceholderSets
      .map(linePlaceholders => replacePlaceholdersInXml(blockTemplate, linePlaceholders))
      .join('');
  };

  let expanded = markerReplaced.replace(tokenPattern, (_match, blockTemplate)=>renderBlock(blockTemplate));
  expanded = expanded.replace(rawPattern, (_match, blockTemplate)=>renderBlock(blockTemplate));

  // Remove any leftover markers so they are never visible in output.
  expanded = replacePlaceholdersInXml(expanded, {
    linjer_start: '',
    lss: '',
    linjer_slutt: '',
    lse: ''
  });
  return expanded
    .split(OFFER_LINE_BLOCK_START_TOKEN).join('')
    .split(OFFER_LINE_BLOCK_END_TOKEN).join('')
    .replace(/\{\{\s*(?:lss|linjer_start|lse|linjer_slutt)\s*\}\}/g, '');
}

async function generateOfferDocxBuffer(project, offerNumber, offerDate) {
  await fs.access(OFFER_TEMPLATE_FILE);
  const placeholders = buildOfferPlaceholderValues(project, offerNumber, offerDate);
  const linePlaceholderSets = buildOfferLinePlaceholderValues(project);
  const zip = new AdmZip(OFFER_TEMPLATE_FILE);
  const entries = zip.getEntries().filter(entry=>
    !entry.isDirectory &&
    entry.entryName.startsWith('word/') &&
    entry.entryName.endsWith('.xml')
  );

  entries.forEach(entry=>{
    const xml = entry.getData().toString('utf8');
    const withExpandedLineBlocks = expandLineRepeatBlocks(xml, linePlaceholderSets);
    const replaced = replacePlaceholdersInXml(withExpandedLineBlocks, placeholders);
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
  marketScheduleState.lastAttemptAt = new Date().toISOString();
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
  }
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

app.get('/api/market-data', async (req, res) => {
  try {
    await ensureMarketDataLoaded();
    res.json(currentMarketPayloadForResponse());
  } catch (err) {
    console.error('Markedsdata feilet', err);
    res.status(502).json({ error: 'Kunne ikke hente markedsdata' });
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
    const offerNumber = await allocateOfferNumber(now);
    const buffer = await generateOfferDocxBuffer(project, offerNumber, now);
    const projectName = sanitizeFileName(project.name || 'prosjekt');
    const fileName = `Tilbud-${projectName}-${offerNumber}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Offer-Number', offerNumber);
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
app.listen(port, () => {
  console.log(`Mail service lytter pa port ${port}`);
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
