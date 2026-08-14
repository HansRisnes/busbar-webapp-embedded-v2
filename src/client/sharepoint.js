import {
  PROJECT_SORT_OPTIONS
} from './config.js';
import {
  $
} from './dom.js';
import {
  normalizeListSearchText
} from './global-lists.js';
import {
  compareNoText
} from './projects.js';

export function getSharePointSortMode(state){
  return PROJECT_SORT_OPTIONS.includes(state?.sort) ? state.sort : 'alpha_asc';
}

export function sharePointItemMatchesSearch(item, state){
  const term = normalizeListSearchText(state?.searchTerm);
  if (!term) return true;
  const haystack = [
    item?.name,
    item?.folder ? 'mappe' : 'fil'
  ].map(value=>String(value || '').toLowerCase()).join(' ');
  return haystack.includes(term);
}

export function compareSharePointItems(a, b, state){
  const folderCmp = Number(Boolean(b?.folder)) - Number(Boolean(a?.folder));
  if (folderCmp) return folderCmp;
  const mode = getSharePointSortMode(state);
  if (mode === 'alpha_desc'){
    return compareNoText(b?.name, a?.name);
  }
  if (mode === 'date_newest' || mode === 'date_oldest'){
    const aTime = new Date(a?.lastModifiedDateTime || 0).getTime() || 0;
    const bTime = new Date(b?.lastModifiedDateTime || 0).getTime() || 0;
    return mode === 'date_oldest' ? aTime - bTime : bTime - aTime;
  }
  return compareNoText(a?.name, b?.name);
}

export function formatTimestampForSharePoint(value){
  const raw = String(value || '').trim();
  if (!raw) return '';
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return '';
  return new Intl.DateTimeFormat('no-NO', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

export function renderSharePointFolderItems(options = {}){
  const {
    config,
    formatSharePointFileSize = (()=>''),
    items = [],
    page = '',
    state = {}
  } = options;
  const list = $(config?.listId);
  if (!list) return;
  list.innerHTML = '';
  const sortedItems = [...items]
    .filter(item=>sharePointItemMatchesSearch(item, state))
    .sort((a, b)=>compareSharePointItems(a, b, state));
  if (!sortedItems.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = normalizeListSearchText(state.searchTerm)
      ? 'Ingen filer eller mapper matcher søket.'
      : 'Ingen filer eller mapper funnet.';
    list.appendChild(empty);
    return;
  }
  sortedItems.forEach(item=>{
    const row = document.createElement('article');
    row.className = `sharepoint-item${item.folder ? ' is-folder' : ' is-file'}`;

    const link = document.createElement('a');
    link.className = 'sharepoint-item-link';
    link.href = item.webUrl || '#';
    link.target = '_blank';
    link.rel = 'noopener noreferrer';

    const icon = document.createElement('span');
    icon.className = 'sharepoint-item-icon';
    icon.textContent = item.folder ? 'Mappe' : 'Fil';

    const body = document.createElement('span');
    body.className = 'sharepoint-item-body';
    const name = document.createElement('strong');
    name.textContent = item.name || 'Uten navn';
    const meta = document.createElement('span');
    meta.className = 'muted-text';
    const modified = formatTimestampForSharePoint(item.lastModifiedDateTime);
    const size = item.file ? formatSharePointFileSize(item.size) : '';
    meta.textContent = [item.folder ? `${Number(item.folder.childCount || 0)} elementer` : size, modified ? `Endret ${modified}` : ''].filter(Boolean).join(' | ');
    body.append(name, meta);

    link.append(icon, body);
    row.appendChild(link);

    const actions = document.createElement('div');
    actions.className = 'sharepoint-item-actions';
    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'btn danger btn-small';
    deleteBtn.textContent = 'Slett';
    deleteBtn.dataset.deleteSharepointItem = item.id || '';
    deleteBtn.dataset.sharepointPage = page || '';
    deleteBtn.dataset.sharepointName = item.name || '';
    actions.appendChild(deleteBtn);
    row.appendChild(actions);

    list.appendChild(row);
  });
}
