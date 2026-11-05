import { buildApiUrl } from './api.js';
import { $, closeFormModal, openFormModal } from './dom.js';

const PHASE_LABELS = {
  new: 'Ny',
  contacted: 'Kontaktet',
  qualified: 'Kvalifisert',
  offer_sent: 'Tilbud sendt',
  won: 'Vunnet',
  lost: 'Tapt'
};

const FIELDS = [
  'companyName', 'contactName', 'contactEmail', 'contactPhone', 'projectNeed',
  'salesPhase', 'nextActivity', 'followUpDate', 'responsibleEmail',
  'estimatedValue', 'decisionDate', 'projectPhase', 'customerType',
  'leadSource', 'probability', 'comments'
];

const state = { leads: [], users: [], customers: [], canAssignOthers: false, editingId: '', requestId: 0, formCompanyValue: '', formContactValue: '' };
let access = null;
let datePickerCursor = new Date();
let datePickerField = 'followUpDate';

const EXTRA_SUGGESTIONS = {
  projectPhase: ['Prosjektering', 'Anbud', 'Gjennomføring'],
  customerType: ['Elektroentreprenør', 'Rådgiver', 'Byggherre'],
  leadSource: ['Nettside', 'LinkedIn', 'Anbefaling', 'Oppsøkende salg']
};

export function getLeadsForDashboard() {
  return state.leads;
}

function fieldId(name) {
  return name === 'leadSource' ? 'leadSource' : `lead${name[0].toUpperCase()}${name.slice(1)}`;
}

function field(name) {
  return $(fieldId(name));
}

function setStatus(message, error = false) {
  const el = $('leadsStatus');
  if (!el) return;
  el.textContent = message;
  el.classList.toggle('error', error);
  el.classList.toggle('ok', !error && Boolean(message));
}

function formatDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return '-';
  const date = new Date(`${value}T12:00:00`);
  return Number.isNaN(date.getTime()) ? '-' : new Intl.DateTimeFormat('no-NO', {
    day: '2-digit', month: '2-digit', year: 'numeric'
  }).format(date);
}

function todayKey() {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
}

function parseFollowUpDate(value) {
  const match = String(value || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return '';
  const iso = `${match[3]}-${match[2]}-${match[1]}`;
  const date = new Date(`${iso}T00:00:00Z`);
  return Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== iso ? '' : iso;
}

function displayFollowUpDate(iso) {
  const match = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : '';
}

function findCustomer(name) {
  const key = String(name || '').trim().toLocaleLowerCase('no');
  return state.customers.find(customer => customer.name.toLocaleLowerCase('no') === key);
}

function showLeadSuggestions(list, values, query) {
  if (!list) return;
  const term = String(query || '').trim().toLocaleLowerCase('no');
  const sorted = [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'no'));
  const matches = sorted.filter(value => value.toLocaleLowerCase('no').startsWith(term))
    .concat(sorted.filter(value => !value.toLocaleLowerCase('no').startsWith(term) && value.toLocaleLowerCase('no').includes(term)))
    .slice(0, 12);
  list.replaceChildren();
  matches.forEach(value => {
    const item = document.createElement('li');
    item.textContent = value;
    item.dataset.value = value;
    list.appendChild(item);
  });
  list.hidden = matches.length === 0;
}

function hideLeadSuggestions(list) {
  if (list) list.hidden = true;
}

function refreshCustomerSuggestions() {
  showLeadSuggestions($('leadCustomerSuggestions'), state.customers.map(customer => customer.name), field('companyName')?.value);
}

function refreshContactSuggestions() {
  const customer = findCustomer(field('companyName')?.value);
  showLeadSuggestions($('leadContactSuggestions'), (customer?.contacts || []).map(contact => contact.name), field('contactName')?.value);
}

function updateLeadCompany(value) {
  field('companyName').value = value;
  if (value !== state.formCompanyValue) {
    field('contactName').value = '';
    field('contactEmail').value = '';
    field('contactPhone').value = '';
    state.formCompanyValue = value;
    state.formContactValue = '';
  }
  hideLeadSuggestions($('leadContactSuggestions'));
}

function updateLeadContact(value) {
  field('contactName').value = value;
  if (value !== state.formContactValue) {
    field('contactEmail').value = '';
    field('contactPhone').value = '';
    state.formContactValue = value;
  }
  const contact = findCustomer(field('companyName').value)?.contacts?.find(item =>
    item.name.toLocaleLowerCase('no') === value.trim().toLocaleLowerCase('no'));
  if (contact) {
    field('contactEmail').value = contact.email || '';
    field('contactPhone').value = contact.phone || '';
  }
}

function closeLeadDatePicker() {
  const popover = $('leadDatePickerPopover');
  if (popover) popover.hidden = true;
}

function renderLeadDatePicker() {
  const popover = $('leadDatePickerPopover');
  if (!popover) return;
  popover.replaceChildren();
  const selected = parseFollowUpDate(field(datePickerField)?.value);
  const first = new Date(datePickerCursor.getFullYear(), datePickerCursor.getMonth(), 1);
  const start = new Date(first);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  const head = document.createElement('div');
  head.className = 'calendar-date-picker-head';
  addText(head, 'calendar-date-picker-title', new Intl.DateTimeFormat('no-NO', { month: 'long', year: 'numeric' }).format(first));
  const nav = document.createElement('div');
  nav.className = 'calendar-date-picker-nav';
  for (const [label, offset] of [['Forrige måned', -1], ['Neste måned', 1]]) {
    const button = document.createElement('button');
    button.type = 'button';
    button.setAttribute('aria-label', label);
    button.textContent = offset < 0 ? '‹' : '›';
    button.addEventListener('click', () => {
      datePickerCursor = new Date(first.getFullYear(), first.getMonth() + offset, 1);
      renderLeadDatePicker();
    });
    nav.appendChild(button);
  }
  head.appendChild(nav);
  popover.appendChild(head);
  const grid = document.createElement('div');
  grid.className = 'calendar-date-picker-grid';
  ['Ma', 'Ti', 'On', 'To', 'Fr', 'Lø', 'Sø'].forEach(label => addText(grid, 'calendar-date-picker-weekday', label));
  for (let day = 0; day < 42; day += 1) {
    const current = new Date(start);
    current.setDate(start.getDate() + day);
    const iso = `${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}-${String(current.getDate()).padStart(2, '0')}`;
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'calendar-date-picker-day';
    button.classList.toggle('is-outside', current.getMonth() !== first.getMonth());
    button.classList.toggle('is-today', iso === todayKey());
    button.classList.toggle('is-selected', iso === selected);
    button.textContent = String(current.getDate());
    button.addEventListener('click', () => {
      field(datePickerField).value = displayFollowUpDate(iso);
      field(datePickerField).setCustomValidity('');
      closeLeadDatePicker();
    });
    grid.appendChild(button);
  }
  popover.appendChild(grid);
}

function openLeadDatePicker(name = 'followUpDate') {
  const popover = $('leadDatePickerPopover');
  const trigger = $(`${fieldId(name)}PickerBtn`);
  if (!popover || !trigger) return;
  datePickerField = name;
  const selected = parseFollowUpDate(field(name)?.value);
  datePickerCursor = selected ? new Date(`${selected}T12:00:00`) : new Date();
  renderLeadDatePicker();
  const rect = trigger.getBoundingClientRect();
  const width = Math.min(300, window.innerWidth - 16);
  popover.style.width = `${width}px`;
  popover.style.left = `${Math.max(8, Math.min(rect.right - width, window.innerWidth - width - 8))}px`;
  popover.style.top = `${rect.bottom + 310 <= window.innerHeight ? rect.bottom + 8 : Math.max(8, rect.top - 310)}px`;
  popover.hidden = false;
}

function addText(parent, className, value, tag = 'span') {
  const el = document.createElement(tag);
  el.className = className;
  el.textContent = value;
  parent.appendChild(el);
  return el;
}

function getResponsibleName(lead) {
  return state.users.find(user => user.email === lead.responsibleEmail)?.name
    || lead.responsibleName
    || lead.responsibleEmail
    || '-';
}

function renderLeads() {
  const root = $('leadsList');
  const summary = $('leadsSummary');
  if (!root || !summary) return;
  root.replaceChildren();
  if (!access?.isLoggedIn()) {
    summary.textContent = '';
    addText(root, 'leads-empty', 'Logg inn for å se leads.', 'p');
    return;
  }
  const query = String($('leadSearchInput')?.value || '').trim().toLocaleLowerCase('no');
  const phase = $('leadPhaseFilter')?.value || 'all';
  const matches = state.leads
    .filter(lead => phase === 'all' || lead.salesPhase === phase)
    .filter(lead => !query || [lead.companyName, lead.projectNeed, lead.contactName, lead.contactEmail, lead.nextActivity]
      .some(value => String(value || '').toLocaleLowerCase('no').includes(query)))
    .sort((a, b) => a.followUpDate.localeCompare(b.followUpDate) || b.createdAt.localeCompare(a.createdAt));
  const overdueCount = state.leads.filter(lead => !lead.followUpCompleted && lead.followUpDate < todayKey()).length;
  summary.textContent = `${state.leads.length} ${state.leads.length === 1 ? 'lead' : 'leads'} · ${overdueCount} med passert oppfølging`;
  if (!matches.length) {
    addText(root, 'leads-empty', state.leads.length ? 'Ingen leads samsvarer med filtrene.' : 'Ingen leads registrert.', 'p');
    return;
  }
  matches.forEach(lead => {
    const card = document.createElement('article');
    card.className = 'lead-item';
    card.dataset.leadId = lead.id;
    const open = document.createElement('button');
    open.type = 'button';
    open.className = 'lead-item-open';
    open.dataset.leadOpen = lead.id;
    const heading = document.createElement('span');
    heading.className = 'lead-item-heading';
    addText(heading, 'lead-item-title', lead.projectNeed, 'strong');
    addText(heading, `lead-phase is-${lead.salesPhase}`, PHASE_LABELS[lead.salesPhase] || 'Ny');
    open.appendChild(heading);
    const detailsRoot = document.createElement('span');
    detailsRoot.className = 'lead-item-details';
    const details = [
      ['Bedriftsnavn', lead.companyName],
      ['Kontaktperson', lead.contactName],
      ['Neste aktivitet', lead.nextActivity],
      ['Oppfølging', formatDate(lead.followUpDate)],
      ['Ansvarlig', getResponsibleName(lead)],
      ['Estimert verdi', lead.estimatedValue == null ? '-' : `${new Intl.NumberFormat('no-NO', { maximumFractionDigits: 0 }).format(lead.estimatedValue)} NOK`]
    ];
    details.forEach(([label, value]) => {
      const item = document.createElement('span');
      item.className = 'lead-item-detail';
      addText(item, 'lead-item-label', `${label}:`, 'strong');
      addText(item, 'lead-item-value', ` ${String(value || '').trim() || '-'}`);
      if (label === 'Oppfølging' && !lead.followUpCompleted && lead.followUpDate < todayKey()) {
        item.classList.add('is-overdue');
      }
      detailsRoot.appendChild(item);
    });
    open.appendChild(detailsRoot);
    if (lead.calendarSyncError) {
      const warning = addText(open, 'lead-calendar-warning', 'Kalender ikke synkronisert');
      warning.title = lead.calendarSyncError;
    }
    const actions = document.createElement('label');
    actions.className = 'lead-item-actions';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'lead-follow-up-checkbox';
    checkbox.dataset.leadComplete = lead.id;
    checkbox.checked = lead.followUpCompleted === true;
    checkbox.disabled = lead.canEdit === false;
    checkbox.setAttribute('aria-label', `Oppfølging fullført: ${lead.projectNeed}`);
    actions.appendChild(checkbox);
    card.append(actions, open);
    root.appendChild(card);
  });
}

async function request(path, body) {
  const response = await fetch(buildApiUrl(path), {
    method: body === undefined ? 'GET' : 'POST',
    headers: {
      ...(body === undefined ? {} : { 'Content-Type': 'application/json' }),
      ...access.getAuthHeaders()
    },
    ...(body === undefined ? {} : { body: JSON.stringify(body) })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || 'Forespørselen feilet');
  return payload;
}

function applyPayload(payload) {
  state.leads = Array.isArray(payload.leads) ? payload.leads : [];
  state.users = Array.isArray(payload.users) ? payload.users : [];
  state.customers = Array.isArray(payload.customers) ? payload.customers : [];
  state.canAssignOthers = payload.canAssignOthers === true;
  renderLeads();
  access?.onChanged?.();
}

export function clearLeadsPage() {
  state.requestId += 1;
  state.leads = [];
  state.users = [];
  state.customers = [];
  state.editingId = '';
  closeLeadDatePicker();
  closeFormModal('leadForm');
  if ($('newLeadBtn')) $('newLeadBtn').disabled = true;
  setStatus('');
  renderLeads();
  access?.onChanged?.();
}

export async function loadLeadsPage() {
  if (!access || !$('leadsView')) return;
  if (!access.isLoggedIn()) {
    clearLeadsPage();
    return;
  }
  const requestId = ++state.requestId;
  if ($('newLeadBtn')) $('newLeadBtn').disabled = false;
  setStatus('Henter leads...');
  try {
    const payload = await request('/api/leads');
    if (requestId !== state.requestId || !access.isLoggedIn()) return;
    applyPayload(payload);
    setStatus(`Oppdatert ${new Date().toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' })}`);
  } catch (error) {
    if (requestId !== state.requestId) return;
    setStatus(error.message || 'Kunne ikke hente leads', true);
  }
}

function populateResponsible(lead) {
  const select = field('responsibleEmail');
  if (!select) return;
  select.replaceChildren();
  const users = [...new Map(state.users.map(user => [user.email.toLowerCase(), user])).values()];
  const current = lead?.responsibleEmail || access.getCurrentUserEmail();
  if (current && !users.some(user => user.email === current)) {
    users.push({ email: current, name: lead?.responsibleName || current });
  }
  users.forEach(user => {
    const option = document.createElement('option');
    option.value = user.email;
    option.textContent = user.name || user.email;
    select.appendChild(option);
  });
  select.value = current;
  select.disabled = !state.canAssignOthers || lead?.canEdit === false;
}

function setOptionalSelectValue(name, value) {
  const select = field(name);
  if (!select) return;
  select.querySelector('option[data-legacy-value]')?.remove();
  const selected = value == null ? '' : String(value);
  if (selected && !Array.from(select.options).some(option => option.value === selected)) {
    const option = document.createElement('option');
    option.value = selected;
    option.textContent = name === 'probability'
      ? `${selected} %`
      : new Intl.NumberFormat('no-NO', { maximumFractionDigits: 2 }).format(Number(selected));
    option.dataset.legacyValue = '';
    select.appendChild(option);
  }
  select.value = selected;
}

function openLeadForm(id = '') {
  const form = $('leadForm');
  if (!form || !access?.isLoggedIn()) return false;
  const lead = id ? state.leads.find(item => item.id === id) : null;
  if (id && !lead) return false;
  state.editingId = lead?.id || '';
  form.reset();
  const lookupName = Math.random().toString(36).slice(2);
  field('companyName').name = `lead-company-${lookupName}`;
  field('contactName').name = `lead-contact-${lookupName}`;
  Object.keys(EXTRA_SUGGESTIONS).forEach(name => {
    field(name).name = `lead-${name}-${lookupName}`;
    hideLeadSuggestions($(`${fieldId(name)}Suggestions`));
  });
  closeLeadDatePicker();
  $('leadId').value = state.editingId;
  const activity = field('nextActivity');
  activity.querySelector('[data-legacy-activity]')?.remove();
  if (lead?.nextActivity && !Array.from(activity.options).some(option => option.value === lead.nextActivity)) {
    const option = document.createElement('option');
    option.value = lead.nextActivity;
    option.textContent = lead.nextActivity;
    option.dataset.legacyActivity = '';
    activity.appendChild(option);
  }
  FIELDS.forEach(name => {
    const input = field(name);
    if (!input) return;
    if (name === 'estimatedValue' || name === 'probability') {
      setOptionalSelectValue(name, lead?.[name]);
    } else {
      input.value = name === 'followUpDate' || name === 'decisionDate'
        ? displayFollowUpDate(lead?.[name])
        : lead?.[name] ?? (name === 'salesPhase' ? 'new' : '');
    }
    input.disabled = lead?.canEdit === false;
  });
  state.formCompanyValue = field('companyName').value;
  state.formContactValue = field('contactName').value;
  hideLeadSuggestions($('leadCustomerSuggestions'));
  hideLeadSuggestions($('leadContactSuggestions'));
  populateResponsible(lead);
  const hasExtras = Boolean(lead && ['estimatedValue', 'decisionDate', 'projectPhase', 'customerType', 'leadSource', 'probability', 'comments']
    .some(name => lead[name] !== null && lead[name] !== ''));
  form.querySelector('.lead-extra-fields').open = hasExtras;
  $('leadFormError').hidden = true;
  $('leadFormError').textContent = '';
  $('saveLeadBtn').hidden = lead?.canEdit === false;
  $('deleteLeadBtn').hidden = !lead || lead.canEdit === false;
  form.hidden = false;
  $('leadFormModal')?.classList.remove('lead-modal-anchored');
  openFormModal('leadForm', lead ? (lead.canEdit ? 'Endre lead' : 'Lead') : 'Ny lead');
  const modal = $('leadFormModal');
  const body = modal?.querySelector('.form-modal-body');
  if (body && !body.dataset.leadScrollBound) {
    body.addEventListener('scroll', closeLeadDatePicker);
    body.dataset.leadScrollBound = '1';
  }
  const top = modal?.querySelector('.form-modal-panel')?.getBoundingClientRect().top;
  if (Number.isFinite(top)) {
    modal.style.setProperty('--lead-modal-top', `${top}px`);
    modal.classList.add('lead-modal-anchored');
  }
  field('companyName')?.focus();
  if (lead?.canEdit !== false) refreshCustomerSuggestions();
  return true;
}

export function openLeadById(id) {
  return openLeadForm(id);
}

async function saveLead(event) {
  event.preventDefault();
  const form = $('leadForm');
  const dateInput = field('followUpDate');
  const decisionInput = field('decisionDate');
  dateInput.setCustomValidity(parseFollowUpDate(dateInput.value) ? '' : 'Oppgi dato som DD/MM/ÅÅÅÅ');
  decisionInput.setCustomValidity(!decisionInput.value.trim() || parseFollowUpDate(decisionInput.value)
    ? '' : 'Oppgi dato som DD/MM/ÅÅÅÅ');
  if (!form?.reportValidity()) return;
  const lead = state.leads.find(item => item.id === state.editingId);
  if (lead && !lead.canEdit) return;
  const body = { id: state.editingId, updatedAt: lead?.updatedAt || '' };
  FIELDS.forEach(name => { body[name] = field(name)?.value ?? ''; });
  body.followUpDate = parseFollowUpDate(dateInput.value);
  body.decisionDate = parseFollowUpDate(decisionInput.value);
  const save = $('saveLeadBtn');
  const requester = access.getCurrentUserEmail();
  save.disabled = true;
  $('leadFormError').hidden = true;
  try {
    const payload = await request('/api/leads/upsert', body);
    if (!access.isLoggedIn() || access.getCurrentUserEmail() !== requester) return;
    state.requestId += 1;
    applyPayload(payload);
    closeLeadDatePicker();
    closeFormModal('leadForm');
    const saved = state.leads.find(item => item.id === (lead?.id || payload.leads.at(-1)?.id));
    setStatus(saved?.calendarSyncError ? 'Lagret, men kalender ikke synkronisert' : 'Lagret', Boolean(saved?.calendarSyncError));
  } catch (error) {
    $('leadFormError').textContent = error.message || 'Kunne ikke lagre lead';
    $('leadFormError').hidden = false;
  } finally {
    save.disabled = false;
  }
}

async function deleteLead() {
  const lead = state.leads.find(item => item.id === state.editingId);
  if (!lead?.canEdit || !window.confirm(`Slette lead for ${lead.companyName}?`)) return;
  const button = $('deleteLeadBtn');
  const requester = access.getCurrentUserEmail();
  button.disabled = true;
  $('leadFormError').hidden = true;
  try {
    const payload = await request('/api/leads/delete', { id: lead.id, updatedAt: lead.updatedAt });
    if (!access.isLoggedIn() || access.getCurrentUserEmail() !== requester) return;
    state.requestId += 1;
    applyPayload(payload);
    closeLeadDatePicker();
    closeFormModal('leadForm');
    setStatus('Leaden er slettet');
  } catch (error) {
    $('leadFormError').textContent = error.message || 'Kunne ikke slette lead';
    $('leadFormError').hidden = false;
  } finally {
    button.disabled = false;
  }
}

async function toggleLeadCompleted(input) {
  const lead = state.leads.find(item => item.id === input.dataset.leadComplete);
  if (!lead?.canEdit) return;
  input.disabled = true;
  try {
    const payload = await request('/api/leads/complete', {
      id: lead.id, updatedAt: lead.updatedAt, completed: input.checked
    });
    state.requestId += 1;
    applyPayload(payload);
    const updated = state.leads.find(item => item.id === lead.id);
    setStatus(updated?.calendarSyncError ? 'Oppdatert, men kalender ikke synkronisert' : 'Oppfølging oppdatert', Boolean(updated?.calendarSyncError));
  } catch (error) {
    input.checked = !input.checked;
    input.disabled = false;
    setStatus(error.message || 'Kunne ikke oppdatere oppfølging', true);
  }
}

export function initLeadsPage(dependencies) {
  if (!$('leadsView') || $('leadsView').dataset.leadsBound === '1') return;
  access = dependencies;
  $('leadsView').dataset.leadsBound = '1';
  $('newLeadBtn')?.addEventListener('click', () => openLeadForm());
  $('refreshLeadsBtn')?.addEventListener('click', loadLeadsPage);
  $('leadSearchInput')?.addEventListener('input', renderLeads);
  $('leadPhaseFilter')?.addEventListener('change', renderLeads);
  $('leadsList')?.addEventListener('change', event => {
    if (event.target.matches('[data-lead-complete]')) void toggleLeadCompleted(event.target);
  });
  $('leadsList')?.addEventListener('click', event => {
    const button = event.target.closest('[data-lead-open]');
    if (button) openLeadForm(button.dataset.leadOpen);
  });
  $('leadCompanyName')?.addEventListener('focus', refreshCustomerSuggestions);
  $('leadCompanyName')?.addEventListener('input', () => {
    updateLeadCompany(field('companyName').value);
    refreshCustomerSuggestions();
  });
  $('leadCompanyName')?.addEventListener('blur', () => setTimeout(() => hideLeadSuggestions($('leadCustomerSuggestions')), 80));
  $('leadCustomerSuggestions')?.addEventListener('mousedown', event => {
    const item = event.target.closest('li[data-value]');
    if (!item) return;
    event.preventDefault();
    updateLeadCompany(item.dataset.value);
    hideLeadSuggestions($('leadCustomerSuggestions'));
    field('contactName').focus();
  });
  $('leadContactName')?.addEventListener('focus', refreshContactSuggestions);
  $('leadContactName')?.addEventListener('input', () => {
    updateLeadContact(field('contactName').value);
    refreshContactSuggestions();
  });
  $('leadContactName')?.addEventListener('blur', () => setTimeout(() => hideLeadSuggestions($('leadContactSuggestions')), 80));
  $('leadContactSuggestions')?.addEventListener('mousedown', event => {
    const item = event.target.closest('li[data-value]');
    if (!item) return;
    event.preventDefault();
    updateLeadContact(item.dataset.value);
    hideLeadSuggestions($('leadContactSuggestions'));
  });
  Object.entries(EXTRA_SUGGESTIONS).forEach(([name, values]) => {
    const input = field(name);
    const list = $(`${fieldId(name)}Suggestions`);
    const refresh = () => showLeadSuggestions(list, values, input.value);
    input?.addEventListener('focus', refresh);
    input?.addEventListener('input', refresh);
    input?.addEventListener('blur', () => setTimeout(() => hideLeadSuggestions(list), 80));
    list?.addEventListener('mousedown', event => {
      const item = event.target.closest('li[data-value]');
      if (!item) return;
      event.preventDefault();
      input.value = item.dataset.value;
      hideLeadSuggestions(list);
    });
  });
  $('leadFollowUpDate')?.addEventListener('input', () => field('followUpDate').setCustomValidity(''));
  $('leadDecisionDate')?.addEventListener('input', () => field('decisionDate').setCustomValidity(''));
  $('leadFollowUpDatePickerBtn')?.addEventListener('click', () => openLeadDatePicker('followUpDate'));
  $('leadDecisionDatePickerBtn')?.addEventListener('click', () => openLeadDatePicker('decisionDate'));
  document.addEventListener('pointerdown', event => {
    const popover = $('leadDatePickerPopover');
    if (popover?.hidden || popover?.contains(event.target)
      || $('leadFollowUpDatePickerBtn')?.contains(event.target)
      || $('leadDecisionDatePickerBtn')?.contains(event.target)) return;
    closeLeadDatePicker();
  });
  document.addEventListener('keydown', event => {
    if (event.key !== 'Escape' || $('leadDatePickerPopover')?.hidden !== false) return;
    event.stopPropagation();
    event.preventDefault();
    closeLeadDatePicker();
  }, true);
  $('leadForm')?.addEventListener('submit', saveLead);
  $('cancelLeadBtn')?.addEventListener('click', () => { closeLeadDatePicker(); closeFormModal('leadForm'); });
  const popover = $('leadDatePickerPopover');
  if (popover) document.body.appendChild(popover);
  $('deleteLeadBtn')?.addEventListener('click', deleteLead);
  clearLeadsPage();
}
