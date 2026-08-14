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
  if (configured) return configured;
  const params = new URLSearchParams(window.location.search);
  const view = String(params.get('view') || '').trim();
  if (view) return view;
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
  const nextPage = String(page || 'projects');
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

export function getDashboardTotals(projects, resolvers = {}){
  const lines = getDashboardAllProjectLines(projects);
  const resolveLineSkinMaterialCost = typeof resolvers.resolveLineSkinMaterialCost === 'function'
    ? resolvers.resolveLineSkinMaterialCost
    : (()=>0);
  const resolveLineDisplayTotal = typeof resolvers.resolveLineDisplayTotal === 'function'
    ? resolvers.resolveLineDisplayTotal
    : (line=>line?.totals?.totalInclMontasje);
  const busbarTotal = sumDashboardLineValue(lines, line=>line?.totals?.totalExMontasje);
  const materialCost = sumDashboardLineValue(lines, line=>resolveLineSkinMaterialCost(line));
  const montasjeTotal = sumDashboardLineValue(lines, line=>line?.totals?.totalInclMontasje);
  const montasjeCost = sumDashboardLineValue(lines, line=>line?.totals?.montasje?.cost);
  const lineTotal = sumDashboardLineValue(lines, line=>resolveLineDisplayTotal(line));
  return {
    lineCount: lines.length,
    busbar: {
      total: busbarTotal,
      cost: materialCost,
      margin: Math.max(0, round2(busbarTotal - materialCost))
    },
    montasje: {
      total: montasjeTotal,
      cost: montasjeCost,
      margin: Math.max(0, round2(montasjeTotal - montasjeCost))
    },
    allTotal: lineTotal
  };
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

function createDashboardMetric(label, value, percent){
  const item = document.createElement('div');
  item.className = 'dashboard-total-metric';
  const labelEl = document.createElement('span');
  labelEl.textContent = label;
  const valueEl = document.createElement('strong');
  valueEl.textContent = value;
  const percentEl = document.createElement('small');
  percentEl.textContent = percent;
  item.append(labelEl, valueEl, percentEl);
  return item;
}

export function renderDashboardTotalsWidget(state, projects, resolvers = {}){
  const root = $('dashboardTotalsContent');
  if (!root) return;
  const totals = getDashboardTotals(projects, resolvers);
  const activeTab = state.totalsTab === 'montasje' ? 'montasje' : 'busbar';
  const tabButtons = Array.from(document.querySelectorAll('[data-dashboard-total-tab]'));
  tabButtons.forEach(btn=>{
    const active = btn.dataset.dashboardTotalTab === activeTab;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-selected', active ? 'true' : 'false');
  });
  const data = activeTab === 'montasje'
    ? {
      totalLabel: 'Total montasje',
      costLabel: 'Montasjekost',
      marginLabel: 'Påslag montasje',
      values: totals.montasje
    }
    : {
      totalLabel: 'Total strømskinne',
      costLabel: 'Materiellkost',
      marginLabel: 'Påslag strømskinne',
      values: totals.busbar
    };
  root.innerHTML = '';
  const hero = document.createElement('div');
  hero.className = 'dashboard-total-hero';
  const heroLabel = document.createElement('span');
  heroLabel.textContent = 'Totalsum inkludert i tilbud';
  const heroValue = document.createElement('strong');
  heroValue.textContent = formatDashboardMoney(totals.allTotal);
  const heroMeta = document.createElement('small');
  heroMeta.textContent = `${totals.lineCount} linjer beregnet`;
  hero.append(heroLabel, heroValue, heroMeta);

  const metrics = document.createElement('div');
  metrics.className = 'dashboard-total-metrics';
  metrics.append(
    createDashboardMetric(data.totalLabel, formatDashboardMoney(data.values.total), '100 %'),
    createDashboardMetric(data.costLabel, formatDashboardMoney(data.values.cost), formatDashboardPercent(data.values.cost, data.values.total)),
    createDashboardMetric(data.marginLabel, formatDashboardMoney(data.values.margin), formatDashboardPercent(data.values.margin, data.values.total))
  );
  root.append(hero, metrics);
}

function getDashboardProjectStatusCounts(projects, getProjectStatusConfig){
  const counts = new Map(PROJECT_STATUS_OPTIONS.map(option=>[option.id, 0]));
  (Array.isArray(projects) ? projects : []).forEach(project=>{
    const status = getProjectStatusConfig(project).id;
    counts.set(status, (counts.get(status) || 0) + 1);
  });
  return counts;
}

export function renderDashboardProjectStatusWidget(projects, getProjectStatusConfig){
  const root = $('dashboardProjectStatusSummary');
  if (!root || typeof getProjectStatusConfig !== 'function') return;
  const counts = getDashboardProjectStatusCounts(projects, getProjectStatusConfig);
  root.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'dashboard-status-buttons';
  PROJECT_STATUS_OPTIONS.forEach(option=>{
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `dashboard-status-button is-${option.tone}`;
    btn.dataset.dashboardProjectStatus = option.id;
    btn.innerHTML = `<span>${option.label}</span><strong>${fmtIntNO.format(counts.get(option.id) || 0)}</strong>`;
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
    empty.textContent = 'Ingen anbefalte handlinger akkurat nå.';
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
    empty.textContent = 'Ingen e-posttråder foreslått.';
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
