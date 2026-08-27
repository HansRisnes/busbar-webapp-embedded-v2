import {
  DASHBOARD_SIDEBAR_COLLAPSED_KEY,
  PROJECT_FLOW_PHASES,
  PROJECT_STATUS_OPTIONS
} from './config.js';
import { $ } from './dom.js';
import { fmtIntNO, fmtNO, round2 } from './format.js';
import { readLocalText, writeLocalText } from './storage.js';

export const hasDashboardUI = () => Boolean($('dashboardView') && $('projectList'));
export const hasCalculatorUI = () => Boolean($('calcBtn') && $('series'));
const HIDDEN_DASHBOARD_PAGES = new Set(['project-flow-beta']);

function normalizeDashboardPage(page){
  const normalized = String(page || '').trim();
  return HIDDEN_DASHBOARD_PAGES.has(normalized) ? 'dashboard' : normalized;
}

export function buildAppUrl(fileName, params = {}){
  const url = new URL(fileName, window.location.href);
  Object.entries(params).forEach(([key, value])=>{
    if (value === undefined || value === null || value === ''){
      return;
    }
    url.searchParams.set(key, String(value));
  });
  return url;
}
export function goToDashboard(params = {}){
  window.location.href = buildAppUrl('index.html', params).toString();
}

export function goToCalculator(params = {}){
  window.location.href = buildAppUrl('calculator.html', params).toString();
}

export function updateMarketTickerVisibility(state){
  const ticker = $('marketTicker');
  if (!ticker) return;
  const shouldShow = hasCalculatorUI() || state.activePage === 'dashboard';
  ticker.hidden = !shouldShow;
}

function resolveInitialDashboardPage(){
  const shell = $('dashboardShell');
  const configured = String(shell?.dataset?.dashboardActivePage || '').trim();
  if (configured) return normalizeDashboardPage(configured);
  const params = new URLSearchParams(window.location.search);
  const view = String(params.get('view') || '').trim();
  if (view) return normalizeDashboardPage(view);
  return hasCalculatorUI() ? 'calculator' : 'dashboard';
}

function updateDashboardUrl(page, options = {}){
  if (typeof history === 'undefined' || !history.pushState || !history.replaceState) return;
  const normalized = String(page || '').trim();
  if (!normalized || hasCalculatorUI()) return;
  try{
    const url = new URL(window.location.href);
    const currentView = String(url.searchParams.get('view') || '').trim();
    url.searchParams.set('view', normalized);
    const nextUrl = url.toString();
    if (nextUrl === window.location.href || currentView === normalized) return;
    if (options.replace){
      history.replaceState({ view: normalized }, '', nextUrl);
    } else {
      history.pushState({ view: normalized }, '', nextUrl);
    }
  }catch(_err){}
}

function readDashboardSidebarCollapsed(){
  return readLocalText(DASHBOARD_SIDEBAR_COLLAPSED_KEY, '') === '1';
}

function persistDashboardSidebarCollapsed(collapsed){
  writeLocalText(DASHBOARD_SIDEBAR_COLLAPSED_KEY, collapsed ? '1' : '0');
}

export function setDashboardSidebarCollapsed(state, collapsed, options = {}){
  const shell = $('dashboardShell');
  const toggle = $('dashboardSidebarToggle');
  const next = Boolean(collapsed);
  state.sidebarCollapsed = next;
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
export function setDashboardPage(state, page, onPageActivated, options = {}){
  const nextPage = normalizeDashboardPage(page || 'projects');
  const pages = Array.from(document.querySelectorAll('[data-dashboard-page]'));
  if (pages.length){
    const target = pages.find(el=>el.dataset.dashboardPage === nextPage) || pages[0];
    state.activePage = target.dataset.dashboardPage || 'projects';
    pages.forEach(el=>{
      const active = el === target;
      el.hidden = !active;
      el.classList.toggle('is-active', active);
    });
  } else {
    state.activePage = nextPage;
  }
  const navItems = Array.from(document.querySelectorAll('[data-dashboard-nav], [data-dashboard-link]'));
  navItems.forEach(item=>{
    const itemPage = item.dataset.dashboardNav || item.dataset.dashboardLink || '';
    const active = itemPage === state.activePage;
    item.classList.toggle('is-active', active);
    item.setAttribute('aria-current', active ? 'page' : 'false');
  });
  updateMarketTickerVisibility(state);
  updateDashboardUrl(state.activePage, { replace: options.replaceUrl });
  if (typeof onPageActivated === 'function'){
    onPageActivated(state.activePage, {
      fromNavigation: options.fromNavigation !== false && !options.replaceUrl
    });
  }
}

export function initDashboardShell(state, onPageActivated){
  const shell = $('dashboardShell');
  if (!shell) return;
  if (shell.dataset.dashboardShellBound === '1'){
    setDashboardPage(state, state.activePage, onPageActivated, { replaceUrl: true });
    setDashboardSidebarCollapsed(state, state.sidebarCollapsed, { skipPersist: true });
    return;
  }
  shell.dataset.dashboardShellBound = '1';
  state.activePage = resolveInitialDashboardPage();
  state.sidebarCollapsed = readDashboardSidebarCollapsed();
  shell.addEventListener('click', evt=>{
    const toggleBtn = evt.target?.closest?.('#dashboardSidebarToggle');
    if (toggleBtn){
      evt.preventDefault();
      setDashboardSidebarCollapsed(state, !state.sidebarCollapsed);
      return;
    }
    const navBtn = evt.target?.closest?.('[data-dashboard-nav]');
    if (!navBtn) return;
    evt.preventDefault();
    setDashboardPage(state, navBtn.dataset.dashboardNav, onPageActivated, { fromNavigation: true });
  });
  setDashboardSidebarCollapsed(state, state.sidebarCollapsed, { skipPersist: true });
  setDashboardPage(state, state.activePage, onPageActivated, { replaceUrl: true });
  window.addEventListener('popstate', ()=>{
    const page = resolveInitialDashboardPage();
    setDashboardPage(state, page, onPageActivated, { replaceUrl: true });
  });
}

function getDashboardAllProjectLines(projects){
  return (Array.isArray(projects) ? projects : [])
    .flatMap(project=>Array.isArray(project.lines) ? project.lines : []);
}

function sumDashboardLineValue(lines, resolver){
  return round2((Array.isArray(lines) ? lines : []).reduce((sum, line)=>{
    const value = Number(resolver(line));
    return Number.isFinite(value) ? sum + value : sum;
  }, 0));
}

function sumDashboardIncludedLineValue(lines, includeKey, resolver){
  return sumDashboardLineValue(lines, line=>{
    const config = getDashboardLineAddonConfig(line);
    return config[includeKey] ? resolver(line) : 0;
  });
}

function getDashboardLineAddonConfig(line){
  const config = line?.selectedAddonConfig || line?.totals?.selectedAddonConfig || {};
  return {
    includeMontasje: config.includeMontasje !== false,
    includeEngineering: config.includeEngineering !== false,
    includeOppheng: config.includeOppheng !== false
  };
}

function getDashboardLineIncludedCost(line, resolveLineSkinMaterialCost){
  const totals = line?.totals || {};
  const config = getDashboardLineAddonConfig(line);
  let cost = Number(resolveLineSkinMaterialCost(line));
  if (!Number.isFinite(cost)) cost = 0;
  if (config.includeMontasje){
    const montasjeCost = Number(totals.montasje?.cost);
    if (Number.isFinite(montasjeCost)) cost += montasjeCost;
  }
  if (config.includeEngineering){
    const engineeringCost = Number(totals.engineering?.cost);
    if (Number.isFinite(engineeringCost)) cost += engineeringCost;
  }
  if (config.includeOppheng){
    const opphengCost = Number(totals.oppheng?.cost);
    if (Number.isFinite(opphengCost)) cost += opphengCost;
  }
  return round2(cost);
}

export function getDashboardTotals(projects, resolvers = {}){
  const lines = getDashboardAllProjectLines(projects);
  const resolveLineSkinMaterialCost = typeof resolvers.resolveLineSkinMaterialCost === 'function'
    ? resolvers.resolveLineSkinMaterialCost
    : (()=>0);
  const resolveLineDisplayTotal = typeof resolvers.resolveLineDisplayTotal === 'function'
    ? resolvers.resolveLineDisplayTotal
    : (line=>line?.totals?.totalInclMontasje);
  const getProjectStatusConfig = typeof resolvers.getProjectStatusConfig === 'function'
    ? resolvers.getProjectStatusConfig
    : (project=>({ id: project?.projectStatus || '' }));
  const finishedLines = (Array.isArray(projects) ? projects : [])
    .filter(project=>['won', 'finished'].includes(getProjectStatusConfig(project).id))
    .flatMap(project=>Array.isArray(project.lines) ? project.lines : []);
  const busbarTotal = sumDashboardLineValue(lines, line=>line?.totals?.totalExMontasje);
  const engineeringTotal = sumDashboardIncludedLineValue(lines, 'includeEngineering', line=>line?.totals?.totalInclEngineering);
  const engineeringCost = sumDashboardIncludedLineValue(lines, 'includeEngineering', line=>line?.totals?.engineering?.cost);
  const engineeringMargin = sumDashboardIncludedLineValue(lines, 'includeEngineering', line=>line?.totals?.engineeringMargin);
  const materialCost = sumDashboardLineValue(lines, line=>resolveLineSkinMaterialCost(line));
  const montasjeTotal = sumDashboardIncludedLineValue(lines, 'includeMontasje', line=>line?.totals?.totalInclMontasje);
  const opphengTotal = sumDashboardIncludedLineValue(lines, 'includeOppheng', line=>line?.totals?.totalInclOppheng ?? line?.totals?.total);
  const opphengCost = sumDashboardIncludedLineValue(lines, 'includeOppheng', line=>line?.totals?.oppheng?.cost);
  const opphengMargin = sumDashboardIncludedLineValue(lines, 'includeOppheng', line=>line?.totals?.opphengMargin);
  const montasjeCost = sumDashboardIncludedLineValue(lines, 'includeMontasje', line=>line?.totals?.montasje?.cost);
  const lineTotal = sumDashboardLineValue(lines, line=>resolveLineDisplayTotal(line));
  const lineIncludedCost = sumDashboardLineValue(lines, line=>getDashboardLineIncludedCost(line, resolveLineSkinMaterialCost));
  const finishedBusbarTotal = sumDashboardLineValue(finishedLines, line=>line?.totals?.totalExMontasje);
  const finishedEngineeringTotal = sumDashboardIncludedLineValue(finishedLines, 'includeEngineering', line=>line?.totals?.totalInclEngineering);
  const finishedEngineeringCost = sumDashboardIncludedLineValue(finishedLines, 'includeEngineering', line=>line?.totals?.engineering?.cost);
  const finishedMaterialCost = sumDashboardLineValue(finishedLines, line=>resolveLineSkinMaterialCost(line));
  const finishedMontasjeTotal = sumDashboardIncludedLineValue(finishedLines, 'includeMontasje', line=>line?.totals?.totalInclMontasje);
  const finishedOpphengTotal = sumDashboardIncludedLineValue(finishedLines, 'includeOppheng', line=>line?.totals?.totalInclOppheng ?? line?.totals?.total);
  const finishedOpphengCost = sumDashboardIncludedLineValue(finishedLines, 'includeOppheng', line=>line?.totals?.oppheng?.cost);
  const finishedMontasjeCost = sumDashboardIncludedLineValue(finishedLines, 'includeMontasje', line=>line?.totals?.montasje?.cost);
  const finishedLineTotal = sumDashboardLineValue(finishedLines, line=>resolveLineDisplayTotal(line));
  const finishedAllProfit = round2(
    Math.max(0, finishedBusbarTotal - finishedMaterialCost)
      + Math.max(0, finishedMontasjeTotal - finishedMontasjeCost)
      + Math.max(0, finishedEngineeringTotal - finishedEngineeringCost)
      + Math.max(0, finishedOpphengTotal - finishedOpphengCost)
  );
  return {
    projectCount: (Array.isArray(projects) ? projects : []).length,
    lineCount: lines.length,
    busbar: {
      total: busbarTotal,
      secondaryTotal: engineeringTotal,
      cost: materialCost,
      secondaryCost: engineeringCost,
      margin: Math.max(0, round2(busbarTotal - materialCost)),
      secondaryMargin: Math.max(0, round2(engineeringMargin)),
      realProfit: Math.max(0, round2(finishedBusbarTotal - finishedMaterialCost)),
      secondaryRealProfit: Math.max(0, round2(finishedEngineeringTotal - finishedEngineeringCost)),
      secondaryRealProfitTotal: finishedEngineeringTotal,
      realProfitTotal: finishedBusbarTotal
    },
    montasje: {
      total: montasjeTotal,
      secondaryTotal: opphengTotal,
      cost: montasjeCost,
      secondaryCost: opphengCost,
      margin: Math.max(0, round2(montasjeTotal - montasjeCost)),
      secondaryMargin: Math.max(0, round2(opphengMargin)),
      realProfit: Math.max(0, round2(finishedMontasjeTotal - finishedMontasjeCost)),
      secondaryRealProfit: Math.max(0, round2(finishedOpphengTotal - finishedOpphengCost)),
      secondaryRealProfitTotal: finishedOpphengTotal,
      realProfitTotal: finishedMontasjeTotal
    },
    engineering: {
      total: engineeringTotal,
      realProfitTotal: finishedEngineeringTotal
    },
    oppheng: {
      total: opphengTotal,
      realProfitTotal: finishedOpphengTotal
    },
    allTotal: lineTotal,
    allMargin: Math.max(0, round2(lineTotal - lineIncludedCost)),
    finishedAllProfit,
    finishedLineTotal
  };
}

const DASHBOARD_MONTH_NAMES = [
  'Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni',
  'Juli', 'August', 'September', 'Oktober', 'November', 'Desember'
];

function getDashboardProjectCreatedDate(project){
  const timestamp = new Date(project?.createdAt || '').getTime();
  return Number.isFinite(timestamp) ? new Date(timestamp) : null;
}

function setDashboardSelectOptions(select, options, selectedValue){
  if (!select) return selectedValue;
  const normalizedSelected = String(selectedValue || 'all');
  select.innerHTML = '';
  options.forEach(option=>{
    const element = document.createElement('option');
    element.value = String(option.value);
    element.textContent = option.label;
    select.appendChild(element);
  });
  const hasSelected = options.some(option=>String(option.value) === normalizedSelected);
  select.value = hasSelected ? normalizedSelected : 'all';
  return select.value;
}

function filterDashboardProjectsByCreatedAt(projects, year, month){
  const selectedYear = String(year || 'all');
  const selectedMonth = String(month || 'all');
  return (Array.isArray(projects) ? projects : []).filter(project=>{
    if (selectedYear === 'all' && selectedMonth === 'all') return true;
    const createdAt = getDashboardProjectCreatedDate(project);
    if (!createdAt) return false;
    if (selectedYear !== 'all' && String(createdAt.getFullYear()) !== selectedYear) return false;
    if (selectedMonth !== 'all' && String(createdAt.getMonth() + 1) !== selectedMonth) return false;
    return true;
  });
}

function prepareDashboardTotalFilters(state, projects){
  const yearSelect = $('dashboardTotalsYearFilter');
  const monthSelect = $('dashboardTotalsMonthFilter');
  const years = Array.from(new Set(
    (Array.isArray(projects) ? projects : [])
      .map(getDashboardProjectCreatedDate)
      .filter(Boolean)
      .map(date=>date.getFullYear())
  ))
    .sort((a, b)=>b - a)
    .map(year=>({ value: year, label: String(year) }));
  const selectedYear = setDashboardSelectOptions(yearSelect, [
    { value: 'all', label: 'Alle' },
    ...years
  ], state.totalsYear);
  const selectedMonth = setDashboardSelectOptions(monthSelect, [
    { value: 'all', label: 'Alle' },
    ...DASHBOARD_MONTH_NAMES.map((label, index)=>({ value: index + 1, label }))
  ], state.totalsMonth);
  state.totalsYear = selectedYear;
  state.totalsMonth = selectedMonth;
  return filterDashboardProjectsByCreatedAt(projects, selectedYear, selectedMonth);
}

function formatDashboardMoney(value){
  const number = Number(value);
  return Number.isFinite(number) ? `${fmtNO.format(round2(number))} NOK` : '0,00 NOK';
}

function formatDashboardPercent(value, total){
  const amount = Number(value);
  const base = Number(total);
  if (!Number.isFinite(amount) || !Number.isFinite(base) || base <= 0) return '0,0 %';
  return `${fmtNO.format(round2((amount / base) * 100))} %`;
}

function createDashboardMetric(label, value, percent, options = {}){
  const item = document.createElement('div');
  item.className = 'dashboard-total-metric';
  if (options.className){
    item.classList.add(options.className);
  }
  if (options.moduleRows){
    item.style.setProperty('--dashboard-total-metric-rows', String(options.moduleRows));
  }
  const labelEl = document.createElement('span');
  labelEl.textContent = label;
  const valueEl = document.createElement('strong');
  valueEl.textContent = value;
  item.append(labelEl, valueEl);
  const appendSecondary = ()=>{
    const secondaryLabelEl = document.createElement('span');
    secondaryLabelEl.className = 'dashboard-total-secondary-label';
    secondaryLabelEl.textContent = options.secondaryLabel || '';
    const secondaryValueEl = document.createElement('strong');
    secondaryValueEl.className = 'dashboard-total-secondary-value';
    secondaryValueEl.textContent = options.secondaryValue || '';
    item.append(secondaryLabelEl, secondaryValueEl);
    if (options.secondaryPercent){
      const secondaryPercentEl = document.createElement('small');
      secondaryPercentEl.className = 'dashboard-total-secondary-percent';
      secondaryPercentEl.textContent = options.secondaryPercent;
      item.appendChild(secondaryPercentEl);
    }
  };
  if ((options.secondaryLabel || options.secondaryValue) && !options.secondaryAfterPercent){
    appendSecondary();
  }
  if (percent || options.reservePercentSpace){
    const percentEl = document.createElement('small');
    percentEl.textContent = percent || '\u00a0';
    if (!percent) percentEl.setAttribute('aria-hidden', 'true');
    item.appendChild(percentEl);
  }
  if ((options.secondaryLabel || options.secondaryValue) && options.secondaryAfterPercent){
    appendSecondary();
  }
  return item;
}

export function renderDashboardTotalsWidget(state, projects, resolvers = {}){
  const root = $('dashboardTotalsContent');
  if (!root) return;
  const filteredProjects = prepareDashboardTotalFilters(state, projects);
  const totals = getDashboardTotals(filteredProjects, resolvers);
  const sections = [
    {
      totalLabel: 'Total strømskinne',
      secondaryTotalLabel: 'Total ingeniør',
      costLabel: 'Materiellkost',
      secondaryCostLabel: 'Ingeniørkost',
      marginLabel: 'Påslag strømskinne',
      secondaryMarginLabel: 'Påslag ingeniør',
      realProfitLabel: 'Total fortjeneste strømskinne',
      secondaryRealProfitLabel: 'Total fortjeneste ingeniør',
      values: totals.busbar
    },
    {
      totalLabel: 'Total montasje',
      secondaryTotalLabel: 'Total oppheng',
      costLabel: 'Montasjekost',
      secondaryCostLabel: 'Opphengskost',
      marginLabel: 'Påslag montasje',
      secondaryMarginLabel: 'Påslag oppheng',
      realProfitLabel: 'Total fortjeneste montasje',
      secondaryRealProfitLabel: 'Total fortjeneste oppheng',
      values: totals.montasje
    }
  ];
  root.innerHTML = '';
  const hero = document.createElement('div');
  hero.className = 'dashboard-total-hero';
  hero.style.setProperty('--dashboard-total-hero-rows', '3');
  const heroLabel = document.createElement('span');
  heroLabel.textContent = 'Totalsum inkludert i tilbud';
  const heroValue = document.createElement('strong');
  heroValue.textContent = formatDashboardMoney(totals.allTotal);
  const heroMeta = document.createElement('small');
  heroMeta.textContent = `${fmtIntNO.format(totals.projectCount)} prosjekter | ${totals.lineCount} linjer beregnet`;
  hero.append(heroLabel, heroValue, heroMeta);

  const marginHero = document.createElement('div');
  marginHero.className = 'dashboard-total-hero';
  marginHero.style.setProperty('--dashboard-total-hero-rows', '2');
  const marginHeroLabel = document.createElement('span');
  marginHeroLabel.textContent = 'Påslag inkludert i tilbud';
  const marginHeroValue = document.createElement('strong');
  marginHeroValue.textContent = formatDashboardMoney(totals.allMargin);
  const marginHeroMeta = document.createElement('small');
  marginHeroMeta.textContent = formatDashboardPercent(totals.allMargin, totals.allTotal);
  marginHero.append(marginHeroLabel, marginHeroValue, marginHeroMeta);

  const profitHero = document.createElement('div');
  profitHero.className = 'dashboard-total-hero';
  profitHero.style.setProperty('--dashboard-total-hero-rows', '2');
  const profitHeroLabel = document.createElement('span');
  profitHeroLabel.textContent = 'Total fortjeneste inkludert i tilbud';
  const profitHeroValue = document.createElement('strong');
  profitHeroValue.textContent = formatDashboardMoney(totals.finishedAllProfit);
  const profitHeroMeta = document.createElement('small');
  profitHeroMeta.textContent = formatDashboardPercent(totals.finishedAllProfit, totals.allTotal);
  profitHero.append(profitHeroLabel, profitHeroValue, profitHeroMeta);

  const heroStack = document.createElement('div');
  heroStack.className = 'dashboard-total-hero-stack';
  heroStack.append(hero, marginHero, profitHero);

  const sectionsWrap = document.createElement('div');
  sectionsWrap.className = 'dashboard-total-sections';
  sections.forEach(data=>{
    const section = document.createElement('section');
    section.className = 'dashboard-total-section';
    const metrics = document.createElement('div');
    metrics.className = 'dashboard-total-metrics';
    metrics.append(
      createDashboardMetric(data.totalLabel, formatDashboardMoney(data.values.total), '', {
        className: 'has-secondary-total',
        moduleRows: 3,
        secondaryLabel: data.secondaryTotalLabel,
        secondaryValue: formatDashboardMoney(data.values.secondaryTotal)
      }),
      createDashboardMetric(data.costLabel, formatDashboardMoney(data.values.cost), '', {
        className: 'has-secondary-total',
        moduleRows: 3,
        secondaryLabel: data.secondaryCostLabel,
        secondaryValue: formatDashboardMoney(data.values.secondaryCost)
      }),
      createDashboardMetric(data.marginLabel, formatDashboardMoney(data.values.margin), formatDashboardPercent(data.values.margin, data.values.total), {
        className: 'has-secondary-total',
        moduleRows: 4,
        secondaryAfterPercent: true,
        secondaryLabel: data.secondaryMarginLabel,
        secondaryValue: formatDashboardMoney(data.values.secondaryMargin),
        secondaryPercent: formatDashboardPercent(data.values.secondaryMargin, data.values.secondaryTotal)
      }),
      createDashboardMetric(data.realProfitLabel, formatDashboardMoney(data.values.realProfit), formatDashboardPercent(data.values.realProfit, data.values.realProfitTotal), {
        className: 'has-secondary-total',
        moduleRows: 4,
        secondaryAfterPercent: true,
        secondaryLabel: data.secondaryRealProfitLabel,
        secondaryValue: formatDashboardMoney(data.values.secondaryRealProfit),
        secondaryPercent: formatDashboardPercent(data.values.secondaryRealProfit, data.values.secondaryRealProfitTotal)
      })
    );
    section.append(metrics);
    sectionsWrap.appendChild(section);
  });
  root.append(heroStack, sectionsWrap);
}
function getDashboardProjectStatusCounts(projects, getProjectStatusConfig){
  const counts = new Map(PROJECT_STATUS_OPTIONS.map(option=>[option.id, 0]));
  const busbarTotals = new Map(PROJECT_STATUS_OPTIONS.map(option=>[option.id, 0]));
  (Array.isArray(projects) ? projects : []).forEach(project=>{
    const status = getProjectStatusConfig(project).id;
    counts.set(status, (counts.get(status) || 0) + 1);
    const lines = Array.isArray(project?.lines) ? project.lines : [];
    const total = sumDashboardLineValue(lines, line=>line?.totals?.totalExMontasje);
    busbarTotals.set(status, round2((busbarTotals.get(status) || 0) + total));
  });
  return { counts, busbarTotals };
}

export function renderDashboardProjectStatusWidget(projects, getProjectStatusConfig){
  const root = $('dashboardProjectStatusSummary');
  if (!root || typeof getProjectStatusConfig !== 'function') return;
  const { counts, busbarTotals } = getDashboardProjectStatusCounts(projects, getProjectStatusConfig);
  root.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'dashboard-status-buttons';
  PROJECT_STATUS_OPTIONS.forEach(option=>{
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `dashboard-status-button is-${option.tone}`;
    btn.dataset.dashboardProjectStatus = option.id;
    btn.innerHTML = `
      <span>${option.label}</span>
      <div class="dashboard-status-value-row">
        <strong>${fmtIntNO.format(counts.get(option.id) || 0)}</strong>
        <small>${formatDashboardMoney(busbarTotals.get(option.id) || 0)}</small>
      </div>
    `;
    list.appendChild(btn);
  });
  root.appendChild(list);
}

export function renderDashboardFlowStatusWidget(projects, getProjectFlowStatusForProject){
  const root = $('dashboardFlowStatusSummary');
  if (!root || typeof getProjectFlowStatusForProject !== 'function') return;
  const counts = new Map();
  (Array.isArray(projects) ? projects : []).forEach(project=>{
    const flowStatus = getProjectFlowStatusForProject(project);
    const key = `${flowStatus.label}|${flowStatus.tone}`;
    const item = counts.get(key) || { label: flowStatus.label, tone: flowStatus.tone, count: 0 };
    item.count += 1;
    counts.set(key, item);
  });
  const order = ['Ubehandlet', ...PROJECT_FLOW_PHASES.map(phase=>phase.label)];
  const items = Array.from(counts.values()).sort((a, b)=>{
    const aIndex = order.indexOf(a.label);
    const bIndex = order.indexOf(b.label);
    return (aIndex < 0 ? 999 : aIndex) - (bIndex < 0 ? 999 : bIndex);
  });
  root.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'dashboard-status-buttons';
  if (!items.length){
    const empty = document.createElement('div');
    empty.className = 'dashboard-status-project-empty';
    empty.textContent = 'Ingen prosjekter.';
    root.appendChild(empty);
    return;
  }
  items.forEach(item=>{
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `dashboard-status-button is-${item.tone}`;
    btn.dataset.dashboardFlowStatus = item.label;
    btn.innerHTML = `<span>${item.label}</span><strong>${fmtIntNO.format(item.count)}</strong>`;
    list.appendChild(btn);
  });
  list.classList.toggle('is-scrollable', items.length > 4);
  root.appendChild(list);
}

export function renderDashboardRecommendedActionsWidget(actions = []){
  const root = $('dashboardRecommendedActions');
  if (!root) return;
  root.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'dashboard-recommended-action-list';
  const source = Array.isArray(actions) ? actions : [];
  if (!source.length){
    const empty = document.createElement('div');
    empty.className = 'dashboard-status-project-empty';
    empty.textContent = 'Ingen anbefalte handlinger akkurat nÃ¥.';
    root.appendChild(empty);
    return;
  }
  source.forEach(action=>{
    const item = document.createElement('article');
    item.className = `dashboard-recommended-action is-${action.tone || 'idle'}`;
    const body = document.createElement('div');
    body.className = 'dashboard-recommended-action-body';
    const title = document.createElement('strong');
    title.textContent = action.title || 'Anbefalt handling';
    const meta = document.createElement('span');
    meta.textContent = action.meta || '';
    body.append(title, meta);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn alt btn-small';
    btn.dataset.dashboardRecommendedAction = action.id || '';
    btn.textContent = action.buttonText || 'Behandle';
    item.append(body, btn);
    list.appendChild(item);
  });
  root.appendChild(list);
}

export function renderDashboardEmailProjectSuggestionsWidget(suggestions = []){
  const root = $('dashboardEmailProjectSuggestions');
  if (!root) return;
  root.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'dashboard-email-project-suggestion-list';
  const source = Array.isArray(suggestions) ? suggestions : [];
  if (!source.length){
    const empty = document.createElement('div');
    empty.className = 'dashboard-status-project-empty';
    empty.textContent = 'Ingen e-posttrÃ¥der foreslÃ¥tt.';
    root.appendChild(empty);
    return;
  }
  source.forEach(suggestion=>{
    const item = document.createElement('article');
    item.className = 'dashboard-email-project-suggestion';
    item.dataset.dashboardOpenSuggestionEmail = suggestion.id || '';

    const body = document.createElement('div');
    body.className = 'dashboard-email-project-suggestion-body';
    const title = document.createElement('strong');
    title.textContent = suggestion.projectName || suggestion.subject || '(Uten emne)';
    const meta = document.createElement('span');
    const contact = suggestion.contactPerson || 'Ukjent kontakt';
    const email = suggestion.from ? ` (${suggestion.from})` : '';
    const company = suggestion.customer || '-';
    const received = suggestion.receivedLabel || '-';
    meta.textContent = `${contact}${email} | ${company} | ${received}`;
    const preview = document.createElement('p');
    preview.textContent = suggestion.preview || '';
    body.append(title, meta, preview);

    const actions = document.createElement('div');
    actions.className = 'dashboard-email-project-suggestion-actions';
    const createBtn = document.createElement('button');
    createBtn.type = 'button';
    createBtn.className = 'btn alt btn-small';
    createBtn.dataset.dashboardCreateProjectFromEmail = suggestion.id || '';
    createBtn.textContent = 'Nytt prosjekt';
    const dismiss = document.createElement('button');
    dismiss.type = 'button';
    dismiss.className = 'dashboard-email-project-dismiss';
    dismiss.dataset.dashboardDismissEmailProject = suggestion.id || '';
    dismiss.setAttribute('aria-label', 'Fjern forslag');
    dismiss.innerHTML = '&#10005;';
    actions.append(createBtn, dismiss);
    item.append(body, actions);
    list.appendChild(item);
  });
  root.appendChild(list);
}

export function getProjectsForDashboardStatus(projects, statusId, getProjectStatusConfig, normalizeProjectStatus){
  const normalized = String(statusId || '').trim();
  const source = Array.isArray(projects) ? projects : [];
  if (normalized === 'all') return source;
  const statusNormalizer = typeof normalizeProjectStatus === 'function' ? normalizeProjectStatus : value=>String(value || '');
  return source.filter(project=>getProjectStatusConfig(project).id === statusNormalizer(normalized));
}

export function renderDashboardProjectStatusList(projects, statusId, helpers = {}){
  const list = $('dashboardProjectStatusList');
  if (!list) return;
  const getProjectStatusConfig = helpers.getProjectStatusConfig;
  if (typeof getProjectStatusConfig !== 'function') return;
  const visibleProjects = getProjectsForDashboardStatus(
    projects,
    statusId,
    getProjectStatusConfig,
    helpers.normalizeProjectStatus
  );
  list.innerHTML = '';
  if (!visibleProjects.length){
    const empty = document.createElement('div');
    empty.className = 'dashboard-status-project-empty';
    empty.textContent = 'Ingen prosjekter med denne statusen.';
    list.appendChild(empty);
    return;
  }
  const compareProjectsForSort = typeof helpers.compareProjectsForSort === 'function'
    ? helpers.compareProjectsForSort
    : (()=>0);
  visibleProjects
    .slice()
    .sort((a,b)=>compareProjectsForSort(a, b, 'date_newest'))
    .forEach(project=>{
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'dashboard-status-project-item';
      btn.dataset.dashboardOpenProject = project.id || '';
      const title = project.projectNumber
        ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
        : (project.name || 'Uten navn');
      const status = getProjectStatusConfig(project);
      const lineCount = Array.isArray(project.lines) ? project.lines.length : 0;
      const titleEl = document.createElement('span');
      titleEl.className = 'dashboard-status-project-title';
      titleEl.textContent = title;
      const metaEl = document.createElement('span');
      metaEl.className = 'dashboard-status-project-meta';
      metaEl.textContent = `${project.customer || 'Uten kunde'} | ${fmtIntNO.format(lineCount)} linjer`;
      const statusEl = document.createElement('span');
      statusEl.className = `project-status-badge is-${status.tone}`;
      statusEl.textContent = status.label;
      btn.append(titleEl, metaEl, statusEl);
      list.appendChild(btn);
    });
}

function getProjectsForDashboardFlowStatus(projects, statusLabel, getProjectFlowStatusForProject){
  const normalized = String(statusLabel || '').trim();
  const source = Array.isArray(projects) ? projects : [];
  if (!normalized || typeof getProjectFlowStatusForProject !== 'function') return [];
  return source.filter(project=>getProjectFlowStatusForProject(project).label === normalized);
}

export function renderDashboardFlowStatusList(projects, statusLabel, helpers = {}){
  const list = $('dashboardProjectStatusList');
  if (!list) return;
  const getProjectFlowStatusForProject = helpers.getProjectFlowStatusForProject;
  if (typeof getProjectFlowStatusForProject !== 'function') return;
  const visibleProjects = getProjectsForDashboardFlowStatus(projects, statusLabel, getProjectFlowStatusForProject);
  list.innerHTML = '';
  if (!visibleProjects.length){
    const empty = document.createElement('div');
    empty.className = 'dashboard-status-project-empty';
    empty.textContent = 'Ingen prosjekter med denne prosjektflyt-statusen.';
    list.appendChild(empty);
    return;
  }
  const compareProjectsForSort = typeof helpers.compareProjectsForSort === 'function'
    ? helpers.compareProjectsForSort
    : (()=>0);
  visibleProjects
    .slice()
    .sort((a,b)=>compareProjectsForSort(a, b, 'date_newest'))
    .forEach(project=>{
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'dashboard-status-project-item';
      btn.dataset.dashboardOpenFlowProject = project.id || '';
      const title = project.projectNumber
        ? `${project.projectNumber} - ${project.name || 'Uten navn'}`
        : (project.name || 'Uten navn');
      const flowStatus = getProjectFlowStatusForProject(project);
      const taskCount = Number(helpers.getProjectFlowTaskCount?.(project.id) || 0);
      const titleEl = document.createElement('span');
      titleEl.className = 'dashboard-status-project-title';
      titleEl.textContent = title;
      const metaEl = document.createElement('span');
      metaEl.className = 'dashboard-status-project-meta';
      metaEl.textContent = `${project.customer || 'Uten kunde'} | ${fmtIntNO.format(taskCount)} oppgaver`;
      const statusEl = document.createElement('span');
      statusEl.className = `project-flow-status-badge is-${flowStatus.tone}`;
      statusEl.textContent = flowStatus.label;
      btn.append(titleEl, metaEl, statusEl);
      list.appendChild(btn);
    });
}

export function openDashboardProjectStatusModal(projects, statusId, helpers = {}){
  const modal = $('dashboardProjectStatusModal');
  if (!modal) return;
  const getProjectStatusConfig = helpers.getProjectStatusConfig;
  if (typeof getProjectStatusConfig !== 'function') return;
  const normalized = String(statusId || 'all').trim() || 'all';
  modal.dataset.modalMode = 'project-status';
  modal.dataset.statusId = normalized;
  delete modal.dataset.flowStatusLabel;
  const title = $('dashboardProjectStatusTitle');
  const subtitle = $('dashboardProjectStatusSubtitle');
  const openAllFlowBtn = $('dashboardProjectStatusOpenAllFlow');
  if (openAllFlowBtn) openAllFlowBtn.hidden = true;
  const visibleProjects = getProjectsForDashboardStatus(
    projects,
    normalized,
    getProjectStatusConfig,
    helpers.normalizeProjectStatus
  );
  const statusConfig = normalized === 'all' ? null : getProjectStatusConfig(normalized);
  if (title) title.textContent = statusConfig ? `Prosjekter: ${statusConfig.label}` : 'Alle prosjekter';
  if (subtitle) subtitle.textContent = `${fmtIntNO.format(visibleProjects.length)} prosjekter`;
  renderDashboardProjectStatusList(projects, normalized, helpers);
  modal.style.display = 'flex';
}

export function openDashboardFlowStatusModal(projects, statusLabel, helpers = {}){
  const modal = $('dashboardProjectStatusModal');
  if (!modal) return;
  const getProjectFlowStatusForProject = helpers.getProjectFlowStatusForProject;
  if (typeof getProjectFlowStatusForProject !== 'function') return;
  const label = String(statusLabel || '').trim();
  if (!label) return;
  modal.dataset.modalMode = 'project-flow-status';
  modal.dataset.flowStatusLabel = label;
  delete modal.dataset.statusId;
  const title = $('dashboardProjectStatusTitle');
  const subtitle = $('dashboardProjectStatusSubtitle');
  const visibleProjects = getProjectsForDashboardFlowStatus(projects, label, getProjectFlowStatusForProject);
  if (title) title.textContent = `Prosjektflyt: ${label}`;
  if (subtitle) subtitle.textContent = `${fmtIntNO.format(visibleProjects.length)} prosjekter`;
  const openAllFlowBtn = $('dashboardProjectStatusOpenAllFlow');
  if (openAllFlowBtn){
    openAllFlowBtn.hidden = false;
    openAllFlowBtn.dataset.dashboardOpenFlowAll = label;
  }
  renderDashboardFlowStatusList(projects, label, helpers);
  modal.style.display = 'flex';
}

export function closeDashboardProjectStatusModal(){
  const modal = $('dashboardProjectStatusModal');
  if (!modal) return;
  modal.style.display = 'none';
  delete modal.dataset.statusId;
  delete modal.dataset.flowStatusLabel;
  delete modal.dataset.modalMode;
  const openAllFlowBtn = $('dashboardProjectStatusOpenAllFlow');
  if (openAllFlowBtn){
    openAllFlowBtn.hidden = true;
    delete openAllFlowBtn.dataset.dashboardOpenFlowAll;
  }
}
