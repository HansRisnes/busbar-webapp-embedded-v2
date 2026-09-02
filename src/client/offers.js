import {
  PROJECT_SORT_OPTIONS
} from './config.js';
import {
  $
} from './dom.js';
import {
  compareNoText,
  getSortableCreatedTimestamp
} from './projects.js';

export function normalizeOfferSearchText(value){
  return String(value || '').trim().toLowerCase();
}

function appendOfferDetails(parent, rows){
  rows.forEach(([label, value])=>{
    const line = document.createElement('p');
    line.className = 'offer-detail';
    const strong = document.createElement('strong');
    strong.className = 'offer-detail-label';
    strong.textContent = `${label}:`;
    const detail = document.createElement('span');
    detail.className = 'offer-detail-value';
    detail.textContent = ` ${value || '-'}`;
    line.append(strong, detail);
    parent.appendChild(line);
  });
}

export function getOfferStatusForProject(project, offerListState){
  const id = String(project?.id || '').trim();
  return id ? offerListState.statusByProjectId[id] || null : null;
}

export function buildOfferRows(projects, offerListState){
  return projects.map(project=>({
    project,
    status: getOfferStatusForProject(project, offerListState)
  }));
}

export function offerRowMatchesSearch(row, offerListState, helpers = {}){
  const term = normalizeOfferSearchText(offerListState.searchTerm);
  if (!term) return true;
  const status = row.status || {};
  const project = row.project || {};
  const getProjectResponsibleName = helpers.getProjectResponsibleName || (()=>'');
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

export function compareOfferRows(a, b, sort){
  if (sort === 'alpha_asc'){
    return compareNoText(a?.project?.name, b?.project?.name);
  }
  if (sort === 'alpha_desc'){
    return compareNoText(b?.project?.name, a?.project?.name);
  }
  const aTime = getSortableCreatedTimestamp(a?.project);
  const bTime = getSortableCreatedTimestamp(b?.project);
  if (sort === 'date_oldest') return aTime - bTime;
  return bTime - aTime;
}

export function updateOfferControlValues(offerListState){
  const input = $('offerSearchInput');
  if (input && input.value !== offerListState.searchTerm) input.value = offerListState.searchTerm;
  const select = $('offerSortSelect');
  if (select && PROJECT_SORT_OPTIONS.includes(offerListState.sort)) select.value = offerListState.sort;
}

export function renderOffersPage(options = {}){
  const {
    authState = {},
    offerListState,
    projects = [],
    helpers = {}
  } = options;
  const list = $('offersList');
  if (!list || !offerListState) return;
  list.innerHTML = '';
  const rows = buildOfferRows(projects, offerListState)
    .filter(row=>offerRowMatchesSearch(row, offerListState, helpers))
    .sort((a, b)=>compareOfferRows(a, b, offerListState.sort));
  if (!rows.length){
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.innerHTML = '<p>Ingen tilbud matcher søket.</p>';
    list.appendChild(empty);
    return;
  }
  const getProjectResponsibleName = helpers.getProjectResponsibleName || (()=>'');
  const formatProjectTimestamp = helpers.formatProjectTimestamp || (value=>value || '-');
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
    appendOfferDetails(body, [
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
