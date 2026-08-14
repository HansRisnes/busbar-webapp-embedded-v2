import {
  readLocalText,
  writeLocalText
} from './storage.js';

export function normalizeApiBaseUrl(value){
  const raw = String(value || '').trim();
  if (!raw) return '';
  try{
    const parsed = new URL(raw, window.location.origin);
    if (!['http:', 'https:'].includes(parsed.protocol)) return '';
    const pathname = parsed.pathname.replace(/\/+$/, '');
    const suffix = pathname === '/' ? '' : pathname;
    return `${parsed.origin}${suffix}`;
  }catch(_err){
    return '';
  }
}

export function isLocalDevelopmentHost(){
  const host = String(window.location.hostname || '').toLowerCase();
  return host === 'localhost' || host === '127.0.0.1' || host === '[::1]';
}

export function resolveApiBaseUrl(){
  let fromQuery = '';
  try{
    fromQuery = new URLSearchParams(window.location.search).get('apiBase') || '';
  }catch(_err){}
  const normalizedQuery = normalizeApiBaseUrl(fromQuery);
  if (fromQuery && normalizedQuery){
    writeLocalText('busbar.api.base', normalizedQuery);
    return normalizedQuery;
  }

  if (isLocalDevelopmentHost()){
    return window.location.origin;
  }

  const fromStorage = readLocalText('busbar.api.base', '');
  const fromMeta = document.querySelector('meta[name="busbar-api-base"]')?.getAttribute('content') || '';
  const fromGlobal = typeof window.BUSBAR_API_BASE === 'string' ? window.BUSBAR_API_BASE : '';
  const normalized = normalizeApiBaseUrl(fromMeta || fromGlobal || fromStorage);
  return normalized || '';
}

export const API_BASE_URL = resolveApiBaseUrl();

export function buildApiUrl(path){
  const suffix = String(path || '').startsWith('/') ? String(path) : `/${String(path || '')}`;
  return API_BASE_URL ? `${API_BASE_URL}${suffix}` : suffix;
}

export function isGithubPagesWithoutApiBase(){
  const host = String(window.location.hostname || '').toLowerCase();
  return host.endsWith('github.io') && !API_BASE_URL;
}

export function appendApiBaseHint(errorText, status){
  if (!isGithubPagesWithoutApiBase()) return errorText;
  if (status !== 404 && status !== 405) return errorText;
  return `${errorText}. GitHub Pages kjører kun statisk frontend. Sett <meta name="busbar-api-base" ...> til backend-URL.`;
}
