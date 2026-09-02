import {
  PROJECT_SORT_OPTIONS
} from './config.js';
import {
  $
} from './dom.js';
import {
  compareNoText
} from './projects.js';

export function normalizeListSearchText(value){
  return String(value || '').trim().toLowerCase();
}

export function compareGlobalListItems(a, b, mode, nameAccessor){
  if (mode === 'alpha_desc') return compareNoText(nameAccessor(b), nameAccessor(a));
  if (mode === 'date_newest' || mode === 'date_oldest'){
    const aTime = new Date(a?.updatedAt || a?.createdAt || 0).getTime() || 0;
    const bTime = new Date(b?.updatedAt || b?.createdAt || 0).getTime() || 0;
    return mode === 'date_oldest' ? aTime - bTime : bTime - aTime;
  }
  return compareNoText(nameAccessor(a), nameAccessor(b));
}

export function companyMatchesSearch(customer, searchTerm){
  const term = normalizeListSearchText(searchTerm);
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

export function contactMatchesSearch(contact, searchTerm){
  const term = normalizeListSearchText(searchTerm);
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

function appendGlobalDataDetails(parent, rows) {
  const details = document.createElement('div');
  details.className = 'global-data-details';
  rows.forEach(([label, value]) => {
    const row = document.createElement('p');
    row.className = 'global-data-detail';
    const labelEl = document.createElement('strong');
    labelEl.className = 'global-data-detail-label';
    labelEl.textContent = `${label}:`;
    const valueEl = document.createElement('span');
    valueEl.className = 'global-data-detail-value';
    valueEl.textContent = ` ${String(value || '').trim() || '-'}`;
    row.append(labelEl, valueEl);
    details.appendChild(row);
  });
  parent.appendChild(details);
}

export function renderCompanyCardsList(options = {}) {
  const {
    canEdit = false,
    customers = [],
    globalListState,
    callbacks = {}
  } = options;
  const list = $('companiesList');
  const companyOptions = $('globalCompanyOptions');
  if (companyOptions) {
    companyOptions.innerHTML = '';
    customers.forEach(customer => {
      const option = document.createElement('option');
      option.value = customer.name;
      companyOptions.appendChild(option);
    });
  }
  if (!list || !globalListState) return;
  const filtered = customers
    .filter(customer=>companyMatchesSearch(customer, globalListState.companySearchTerm))
    .sort((a, b)=>compareGlobalListItems(a, b, globalListState.companySort, item=>item?.name));
  const form = $('companyEditForm');
  if (form) form.hidden = !canEdit || form.dataset.editing !== '1';
  const addBtn = $('addCompanyBtn');
  if (addBtn) addBtn.hidden = !canEdit;
  const ownerHint = $('companiesOwnerHint');
  if (ownerHint) ownerHint.hidden = canEdit;
  list.innerHTML = '';
  if (!filtered.length) {
    const empty = document.createElement('div');
    empty.className = 'global-data-empty';
    empty.textContent = normalizeListSearchText(globalListState.companySearchTerm)
      ? 'Ingen firma matcher søket.'
      : 'Ingen firma funnet i prosjektarkivet.';
    list.appendChild(empty);
    return;
  }
  filtered.forEach(customer => {
    const item = document.createElement('article');
    item.className = 'global-data-card';
    item.dataset.customerProjects = customer.name;
    item.addEventListener('click', event=>{
      if (event.target instanceof Element && event.target.closest('button')) return;
      callbacks.openCustomerProjects?.(customer);
    });
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
      edit.addEventListener('click', () => callbacks.openCompanyEditForm?.(customer));
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'btn danger btn-small';
      remove.textContent = 'Slett';
      remove.addEventListener('click', () => callbacks.handleDeleteCompany?.(customer));
      actions.append(edit, remove);
      item.appendChild(actions);
    }
    list.appendChild(item);
  });
}

export function renderContactPersonsList(options = {}) {
  const {
    canEdit = false,
    contacts = [],
    globalListState,
    callbacks = {}
  } = options;
  const list = $('contactsList');
  if (!list || !globalListState) return;
  const filtered = contacts
    .filter(contact=>contactMatchesSearch(contact, globalListState.contactSearchTerm))
    .sort((a, b)=>compareGlobalListItems(a, b, globalListState.contactSort, item=>item?.name));
  const form = $('contactEditForm');
  if (form) form.hidden = !canEdit || form.dataset.editing !== '1';
  const addBtn = $('addContactBtn');
  if (addBtn) addBtn.hidden = !canEdit;
  const ownerHint = $('contactsOwnerHint');
  if (ownerHint) ownerHint.hidden = canEdit;
  list.innerHTML = '';
  if (!filtered.length) {
    const empty = document.createElement('div');
    empty.className = 'global-data-empty';
    empty.textContent = normalizeListSearchText(globalListState.contactSearchTerm)
      ? 'Ingen kontaktpersoner matcher søket.'
      : 'Ingen kontaktpersoner funnet i prosjektarkivet.';
    list.appendChild(empty);
    return;
  }
  filtered.forEach(contact => {
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
      edit.addEventListener('click', () => callbacks.openContactEditForm?.(contact));
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'btn danger btn-small';
      remove.textContent = 'Slett';
      remove.addEventListener('click', () => callbacks.handleDeleteContact?.(contact));
      actions.append(edit, remove);
      item.appendChild(actions);
    }
    list.appendChild(item);
  });
}

export function normalizeGlobalSortMode(mode, fallback = 'alpha_asc'){
  return PROJECT_SORT_OPTIONS.includes(mode) ? mode : fallback;
}
