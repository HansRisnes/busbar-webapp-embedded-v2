import {
  MICROSOFT_GRAPH_BASE_URL,
  PROJECT_MAILBOX_ADDRESS
} from './config.js';

export function projectMailboxGraphPath(path){
  const suffix = String(path || '');
  return `/users/${encodeURIComponent(PROJECT_MAILBOX_ADDRESS)}${suffix.startsWith('/') ? suffix : `/${suffix}`}`;
}

export async function requestMicrosoftGraph(acquireToken, path, scopes, options = {}){
  const accessToken = await acquireToken(scopes);
  if (!accessToken){
    throw new Error('Microsoft returnerte ikke access token.');
  }
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Prefer: 'outlook.timezone="Europe/Oslo"'
  };
  const request = {
    method: options.method || 'GET',
    headers
  };
  if (options.body !== undefined){
    if (options.rawBody){
      request.body = options.body;
      if (options.contentType) headers['Content-Type'] = options.contentType;
    } else {
      headers['Content-Type'] = 'application/json';
      request.body = JSON.stringify(options.body);
    }
  }
  const res = await fetch(`${MICROSOFT_GRAPH_BASE_URL}${path}`, request);
  let payload = null;
  try{
    payload = await res.json();
  }catch(_err){}
  if (!res.ok){
    throw new Error(payload?.error?.message || `Microsoft Graph svarte ${res.status}.`);
  }
  return payload;
}
