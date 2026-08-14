import {
  PROJECT_ARCHIVE_STATUS_IDS,
  PROJECT_STATUS_OPTIONS
} from './config.js';
import {
  readLocalText,
  writeLocalText
} from './storage.js';
import { fmtNO, round2 } from './format.js';
import { $ } from './dom.js';

export function normalizeProjectStatus(value){
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

export function getProjectStatusConfig(projectOrStatus){
  const status = typeof projectOrStatus === 'string'
    ? normalizeProjectStatus(projectOrStatus)
    : normalizeProjectStatus(projectOrStatus?.projectStatus);
  return PROJECT_STATUS_OPTIONS.find(option=>option.id === status) || PROJECT_STATUS_OPTIONS[0];
}

export function projectIsArchived(project){
  return PROJECT_ARCHIVE_STATUS_IDS.includes(getProjectStatusConfig(project).id);
}

export function loadSortMode(storageKey, validModes, fallback){
  const raw = readLocalText(storageKey, '').trim();
  return validModes.includes(raw) ? raw : fallback;
}

export function saveSortMode(storageKey, mode){
  writeLocalText(storageKey, mode);
}

export function getSortableCreatedTimestamp(item){
  const created = new Date(item?.createdAt || item?.updatedAt || 0).getTime();
  return Number.isFinite(created) ? created : 0;
}

export function compareNoText(left, right){
  return String(left || '').localeCompare(String(right || ''), 'no', {
    sensitivity: 'base',
    numeric: true
  });
}

export function normalizeProjectSearchText(value){
  return String(value || '').trim().toLowerCase();
}

export function projectMatchesSearch(project, rawSearchTerm, helpers = {}){
  const term = normalizeProjectSearchText(rawSearchTerm);
  if (!term) return true;
  const getProjectResponsibleName = typeof helpers.getProjectResponsibleName === 'function'
    ? helpers.getProjectResponsibleName
    : (()=>'');
  const getStatusConfig = typeof helpers.getProjectStatusConfig === 'function'
    ? helpers.getProjectStatusConfig
    : getProjectStatusConfig;
  const haystack = [
    project?.name,
    project?.customer,
    project?.contactPerson,
    getProjectResponsibleName(project),
    getStatusConfig(project).label
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

export function compareProjectsForSort(a, b, mode = 'date_newest'){
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

export function compareLinesForSort(a, b, mode = 'date_newest'){
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

export function getProjectDisplayTitle(project){
  const projectTitle = project?.name || 'Uten navn';
  return project?.projectNumber
    ? `${project.projectNumber} - ${projectTitle}`
    : projectTitle;
}

export function calculateProjectTotals(project, resolvers = {}){
  const lines = Array.isArray(project?.lines) ? project.lines : [];
  const resolveLineDisplayTotal = typeof resolvers.resolveLineDisplayTotal === 'function'
    ? resolvers.resolveLineDisplayTotal
    : (()=>NaN);
  const resolveLineSkinMaterialCost = typeof resolvers.resolveLineSkinMaterialCost === 'function'
    ? resolvers.resolveLineSkinMaterialCost
    : (()=>NaN);
  return {
    lineCount: lines.length,
    total: round2(lines.reduce((sum, line)=>{
      const lineTotal = resolveLineDisplayTotal(line);
      return Number.isFinite(lineTotal) ? sum + lineTotal : sum;
    }, 0)),
    busbarTotal: round2(lines.reduce((sum, line)=>{
      const busbarTotal = Number(line?.totals?.totalExMontasje);
      return Number.isFinite(busbarTotal) ? sum + busbarTotal : sum;
    }, 0)),
    skinMaterialTotal: round2(lines.reduce((sum, line)=>{
      const material = resolveLineSkinMaterialCost(line);
      return Number.isFinite(material) ? sum + material : sum;
    }, 0)),
    montasjeTotal: round2(lines.reduce((sum, line)=>{
      const montasjeTotal = Number(line?.totals?.totalInclMontasje);
      return Number.isFinite(montasjeTotal) ? sum + montasjeTotal : sum;
    }, 0))
  };
}

export function createProjectFieldText(label, value){
  const wrapper = document.createElement('span');
  wrapper.className = 'project-field-text';
  const labelEl = document.createElement('strong');
  labelEl.textContent = `${label}:`;
  const valueEl = document.createElement('span');
  valueEl.textContent = ` ${value || '-'}`;
  wrapper.append(labelEl, valueEl);
  return wrapper;
}

export function createProjectInfoRow(label, value){
  const infoRow = document.createElement('div');
  infoRow.className = 'project-info-row project-row-meta';
  infoRow.appendChild(createProjectFieldText(label, value));
  return infoRow;
}

export function updateProjectArchiveUi(archiveMode, isLoggedIn){
  const archived = archiveMode === true;
  const toggleBtn = $('projectArchiveToggleBtn');
  if (toggleBtn){
    toggleBtn.textContent = archived ? 'Prosjekter' : 'Prosjekt arkiv';
    toggleBtn.setAttribute('aria-pressed', archived ? 'true' : 'false');
  }
  const titleEl = document.querySelector('#dashboardView .dashboard-intro h2');
  if (titleEl){
    titleEl.textContent = archived ? 'Prosjekt arkiv' : 'Prosjekter';
  }
  const introEl = document.querySelector('#dashboardView .dashboard-intro .muted-text');
  if (introEl){
    introEl.textContent = archived
      ? 'Arkiverte prosjekter er skrivebeskyttet. Du kan fortsatt vise og skjule linjer.'
      : 'Velg et prosjekt, eller opprett nytt prosjekt. Prisverktøy åpnes på egen side.';
  }
  const newProjectBtn = $('newProjectBtn');
  if (newProjectBtn){
    newProjectBtn.disabled = !isLoggedIn || archived;
  }
}

export function createProjectEmptyState(message, options = {}){
  const empty = document.createElement('div');
  empty.className = 'empty-state';
  const text = document.createElement('p');
  text.textContent = message || 'Ingen prosjekter er registrert ennå.';
  empty.appendChild(text);
  if (options.showCreateButton){
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn alt';
    btn.dataset.action = 'create-project';
    btn.textContent = 'Nytt prosjekt';
    btn.disabled = !options.isLoggedIn || options.archiveMode === true;
    empty.appendChild(btn);
  }
  return empty;
}

function createProjectMetricText(item){
  if (!item.plain) return createProjectFieldText(item.label, item.value);
  const wrapper = document.createElement('span');
  wrapper.className = 'project-field-text';
  const labelEl = document.createElement('span');
  labelEl.textContent = `${item.label}:`;
  const valueEl = document.createElement('span');
  valueEl.textContent = ` ${item.value || '-'}`;
  wrapper.append(labelEl, valueEl);
  return wrapper;
}

function requireHelper(helpers, name, fallback = null){
  const fn = helpers?.[name];
  return typeof fn === 'function' ? fn : fallback;
}

export function createProjectRow(project, options = {}){
  const archiveMode = options.archiveMode === true;
  const expanded = options.expanded === true;
  const isLoggedIn = options.isLoggedIn === true;
  const helpers = options.helpers || {};
  const projectLines = Array.isArray(project?.lines) ? project.lines : [];
  const projectTotals = calculateProjectTotals(project, {
    resolveLineDisplayTotal: helpers.resolveLineDisplayTotal,
    resolveLineSkinMaterialCost: helpers.resolveLineSkinMaterialCost
  });

  const getProjectFlowStatusForProject = requireHelper(helpers, 'getProjectFlowStatusForProject', ()=>({ label: '-', tone: 'idle' }));
  const getProjectSelectedAddonConfig = requireHelper(helpers, 'getProjectSelectedAddonConfig', ()=>({}));
  const getProjectResponsibleName = requireHelper(helpers, 'getProjectResponsibleName', ()=>'');
  const formatProjectTimestamp = requireHelper(helpers, 'formatProjectTimestamp', value=>String(value || '-'));
  const formatProjectMarginBadgeText = requireHelper(helpers, 'formatProjectMarginBadgeText', ()=>'DG prosjekt: -');
  const getProjectMaterialMarginStats = requireHelper(helpers, 'getProjectMaterialMarginStats', ()=>({ uniqueCount: 0 }));
  const shouldUseWarningForProjectMargin = requireHelper(helpers, 'shouldUseWarningForProjectMargin', ()=>false);
  const projectHasConfirmedFolder = requireHelper(helpers, 'projectHasConfirmedFolder', ()=>false);
  const buildAddonSelectorControl = requireHelper(helpers, 'buildAddonSelectorControl', ()=>document.createElement('span'));
  const getSelectedAddonConfig = requireHelper(helpers, 'getSelectedAddonConfig', ()=>({}));
  const formatLineSummary = requireHelper(helpers, 'formatLineSummary', ()=>'');
  const formatLineTotal = requireHelper(helpers, 'formatLineTotal', ()=>'');
  const formatLineSkinMaterialCost = requireHelper(helpers, 'formatLineSkinMaterialCost', ()=>'');
  const formatLineUpdatedText = requireHelper(helpers, 'formatLineUpdatedText', ()=>'');

  const row = document.createElement('section');
  row.className = 'project-row';
  row.dataset.projectRowId = project.id || '';
  if (expanded) row.classList.add('is-expanded');

  const head = document.createElement('div');
  head.className = 'project-row-head';
  const titleWrap = document.createElement('div');
  titleWrap.className = 'project-row-title';
  const title = document.createElement('h3');
  title.textContent = getProjectDisplayTitle(project);
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
  const infoStack = document.createElement('div');
  infoStack.className = 'project-info-column';
  infoStack.appendChild(createProjectInfoRow('Prosjektansvarlig', getProjectResponsibleName(project) || '-'));
  infoStack.appendChild(createProjectInfoRow('Kunde', project.customer || '-'));
  infoStack.appendChild(createProjectInfoRow('Kontaktperson', project.contactPerson || '-'));
  infoStack.appendChild(createProjectInfoRow('Opprettet', formatProjectTimestamp(project.createdAt)));
  infoStack.appendChild(createProjectInfoRow('Linjer', String(projectLines.length)));

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
  marginRow.append(marginBadge, setMarginBtn);

  titleWrap.append(titleRow, marginRow);

  const actions = document.createElement('div');
  actions.className = 'project-row-actions';
  const hasProjectFolder = projectHasConfirmedFolder(project);
  const disableProjectActions = archiveMode;

  const safeProjectId = String(project.id || '').replace(/[^A-Za-z0-9_-]/g,'');
  const detailId = `project-detail-${safeProjectId || Math.random().toString(36).slice(2)}`;

  const detailBtn = document.createElement('button');
  detailBtn.type = 'button';
  detailBtn.className = 'btn alt project-detail-toggle-btn';
  detailBtn.dataset.projectDetail = project.id;
  detailBtn.textContent = expanded ? 'Skjul linjer' : 'Vis linjer';
  detailBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
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
  openFolderBtn.disabled = disableProjectActions || !isLoggedIn || !hasProjectFolder;
  if (!hasProjectFolder) openFolderBtn.title = 'Fant ingen prosjektmappe på SharePoint-stien.';

  const createFolderBtn = document.createElement('button');
  createFolderBtn.type = 'button';
  createFolderBtn.className = 'btn alt';
  createFolderBtn.dataset.projectCreateFolder = project.id;
  createFolderBtn.textContent = 'Opprett prosjektmappe';
  createFolderBtn.disabled = disableProjectActions || !isLoggedIn;

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
  actionButtons.append(detailBtn, newLineBtn, generateOfferBtn, openFolderBtn, createFolderBtn, copyBtn, editBtn, deleteBtn);

  const metricsStack = document.createElement('div');
  metricsStack.className = 'project-metrics-column';
  [
    { label: 'Total strømskinne', value: `${fmtNO.format(projectTotals.busbarTotal)} NOK` },
    { label: 'Materiellkost', value: `${fmtNO.format(projectTotals.skinMaterialTotal)} NOK`, className: 'is-muted-italic', plain: true },
    { label: 'Total montasje', value: `${fmtNO.format(projectTotals.montasjeTotal)} NOK` },
    { label: 'Totalsum inkludert i tilbud', value: `${fmtNO.format(projectTotals.total)} NOK` }
  ].forEach(item=>{
    const metricRow = document.createElement('div');
    metricRow.className = `project-metric-row${item.className ? ` ${item.className}` : ''}`;
    metricRow.appendChild(createProjectMetricText(item));
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
  head.append(titleWrap, actions);
  row.appendChild(head);

  const bodyGrid = document.createElement('div');
  bodyGrid.className = 'project-body-grid';
  bodyGrid.append(infoStack, metricsStack, addonStack);
  row.appendChild(bodyGrid);

  const detail = document.createElement('div');
  detail.className = 'project-detail';
  detail.id = detailId;
  detail.hidden = !expanded;

  const linesWrapper = document.createElement('div');
  linesWrapper.className = 'project-detail-lines';
  const lines = [...projectLines];
  lines.sort((a,b)=>compareLinesForSort(a, b, options.lineSort));
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

      lineBtn.append(nameSpan, infoSpan, materialSpan, totalSpan, updatedSpan);
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

      lineMain.append(lineBtn, lineAddonControl);
      lineWrap.appendChild(lineMain);
      lineActionButtons.append(lineEditBtn, lineDeleteBtn);
      lineWrap.appendChild(lineActionButtons);
      linesWrapper.appendChild(lineWrap);
    });
  }
  detail.appendChild(linesWrapper);
  row.appendChild(detail);
  return row;
}

export function renderProjectsPage(options = {}){
  const {
    authState = {},
    callbacks = {},
    projectState = {},
    rowHelpers = {}
  } = options;
  const listEl = $('projectList');
  if (!listEl) return { rendered: false };
  callbacks.renderMainDashboard?.();
  callbacks.loadProjectFlowState?.();
  callbacks.updateProjectArchiveUi?.();
  listEl.innerHTML = '';
  const projects = Array.isArray(projectState.projects) ? projectState.projects : [];
  const archiveMode = projectState.showArchive === true;
  if (!projects.length){
    listEl.appendChild(createProjectEmptyState('Ingen prosjekter er registrert ennå.', {
      archiveMode,
      isLoggedIn: authState.loggedIn === true,
      showCreateButton: true
    }));
    projectState.expandedProjectId = null;
    callbacks.renderOffersList?.();
    callbacks.renderProjectFlowView?.();
    return { rendered: true, visibleCount: 0 };
  }
  const projectsForView = projects.filter(project=>archiveMode ? projectIsArchived(project) : !projectIsArchived(project));
  const projectMatchesSearch = typeof callbacks.projectMatchesSearch === 'function'
    ? callbacks.projectMatchesSearch
    : (()=>true);
  const visibleProjects = projectsForView.filter(project=>projectMatchesSearch(project));
  callbacks.ensureProjectFolderStatusesLoaded?.();
  if (!visibleProjects.length){
    const message = !projectsForView.length
      ? (archiveMode ? 'Ingen prosjekter i arkivet.' : 'Ingen aktive prosjekter.')
      : 'Ingen prosjekter matcher søket.';
    listEl.appendChild(createProjectEmptyState(message));
    callbacks.renderOffersList?.();
    callbacks.renderProjectFlowView?.();
    return { rendered: true, visibleCount: 0 };
  }
  const frag = document.createDocumentFragment();
  visibleProjects.forEach(project=>{
    frag.appendChild(createProjectRow(project, {
      archiveMode,
      expanded: projectState.expandedProjectId === project.id,
      isLoggedIn: authState.loggedIn === true,
      lineSort: projectState.lineSort,
      helpers: rowHelpers
    }));
  });
  listEl.appendChild(frag);
  callbacks.renderOffersList?.();
  callbacks.renderProjectFlowView?.();
  return { rendered: true, visibleCount: visibleProjects.length };
}
