import {
  $
} from './dom.js';

export function getSelectedEmailMessage(emailViewState){
  const id = String(emailViewState?.selectedMessageId || '').trim();
  if (!id) return null;
  return (emailViewState.messages || []).find(message=>message.id === id) || null;
}

export function updateEmailMessageActions(emailViewState){
  const message = getSelectedEmailMessage(emailViewState);
  document.querySelectorAll('[data-email-action="mark-read"]').forEach(markBtn=>{
    markBtn.textContent = message?.isRead ? 'Marker ulest' : 'Marker lest';
  });
}

function normalizeEmailPreviewText(value){
  return String(value || '')
    .replace(/[\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]/g, '')
    .replace(/[^\S\r\n]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function htmlToEmailPreviewText(value){
  const template = document.createElement('template');
  template.innerHTML = String(value || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '');
  return normalizeEmailPreviewText(template.content.textContent || '');
}

function removeQuotedEmailThread(text){
  const lines = normalizeEmailPreviewText(text).split(/\r?\n/);
  const kept = [];
  const replyHeaderPatterns = [
    /^-{2,}\s*original message\s*-{2,}$/i,
    /^-{2,}\s*opprinnelig melding\s*-{2,}$/i,
    /^_{6,}$/,
    /^from:\s+/i,
    /^fra:\s+/i,
    /^sent:\s+/i,
    /^sendt:\s+/i,
    /^to:\s+/i,
    /^til:\s+/i,
    /^subject:\s+/i,
    /^emne:\s+/i,
    /^on .+ wrote:$/i,
    /^den .+ skrev .*:$/i
  ];
  for (const line of lines){
    const trimmed = line.trim();
    if (replyHeaderPatterns.some(pattern=>pattern.test(trimmed))) break;
    if (/^>+/.test(trimmed)) break;
    kept.push(line);
  }
  return kept.join('\n');
}

function removeEmailPreviewNoise(text){
  const noisePatterns = [
    /^cid:.+/i,
    /^image\d*\.(png|jpe?g|gif|svg)$/i,
    /^\[cid:.+\]$/i,
    /^external sender/i,
    /^ekstern avsender/i,
    /^you don['’]t often get email from\b.*\blearn why this is important\.?$/i,
    /^get outlook for/i,
    /^sendt fra (min|outlook|iphone|android)/i,
    /^sent from (my|outlook|iphone|android)/i,
    /^--\s*$/
  ];
  const signatureStartPatterns = [
    /^med vennlig hilsen[,]?$/i,
    /^best regards[,]?$/i,
    /^kind regards[,]?$/i,
    /^mvh[,]?$/i
  ];
  const kept = [];
  for (const line of normalizeEmailPreviewText(text).split(/\r?\n/)){
    const trimmed = line.trim();
    if (signatureStartPatterns.some(pattern=>pattern.test(trimmed))) break;
    kept.push(trimmed);
  }
  return kept
    .filter(line=>line && !noisePatterns.some(pattern=>pattern.test(line)))
    .filter(line=>/[A-Za-zÆØÅæøå0-9]/.test(line))
    .map(line=>line.replace(/[•●▪◆◇■□▲▼►◄★☆✓✔✕✖]+/g, '').trim())
    .filter(Boolean)
    .join(' ');
}

function getEmailBodyContentText(body){
  const content = String(body?.content || '').trim();
  if (!content) return '';
  return body?.contentType === 'html'
    ? htmlToEmailPreviewText(content)
    : normalizeEmailPreviewText(content);
}

export function getEmailPreviewText(message){
  const sourceText = getEmailBodyContentText(message?.uniqueBody)
    || getEmailBodyContentText(message?.body)
    || normalizeEmailPreviewText(message?.bodyPreview || '');
  return removeEmailPreviewNoise(removeQuotedEmailThread(sourceText));
}

export function renderEmailMessages(messages, emailViewState, helpers = {}){
  const list = $('emailMessagesList');
  if (!list || !emailViewState) return;
  list.innerHTML = '';
  const sortedMessages = (Array.isArray(messages) ? [...messages] : []).sort((a, b)=>{
    const aTime = new Date(a?.receivedDateTime || 0).getTime() || 0;
    const bTime = new Date(b?.receivedDateTime || 0).getTime() || 0;
    return bTime - aTime;
  });
  emailViewState.messages = sortedMessages;
  if (emailViewState.selectedMessageId && !emailViewState.messages.some(message=>message.id === emailViewState.selectedMessageId)){
    emailViewState.selectedMessageId = '';
  }
  updateEmailMessageActions(emailViewState);
  if (!sortedMessages.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = 'Ingen e-poster funnet.';
    list.appendChild(empty);
    return;
  }
  const formatGraphDateTime = helpers.formatGraphDateTime || (value=>value || '');
  const formatEmailMessageMeta = helpers.formatEmailMessageMeta || (message=>{
    const from = message.from?.emailAddress?.name || message.from?.emailAddress?.address || 'Ukjent avsender';
    return `${from} | ${formatGraphDateTime(message.receivedDateTime)}`;
  });
  sortedMessages.forEach(message=>{
    const item = document.createElement('article');
    item.className = `graph-item email-message-item${message.isRead ? '' : ' is-unread'}${message.id === emailViewState.selectedMessageId ? ' is-selected' : ''}`;
    item.dataset.emailMessageId = message.id || '';
    item.tabIndex = 0;
    item.setAttribute('role', 'button');

    const marker = document.createElement('div');
    marker.className = 'email-read-marker';
    marker.setAttribute('aria-hidden', 'true');

    const body = document.createElement('div');
    body.className = 'graph-item-body';
    const title = document.createElement('h3');
    title.textContent = message.subject || '(Uten emne)';
    const meta = document.createElement('p');
    meta.className = 'muted-text';
    meta.textContent = formatEmailMessageMeta(message);
    const preview = document.createElement('p');
    preview.className = 'email-preview';
    preview.textContent = getEmailPreviewText(message);
    body.appendChild(title);
    body.appendChild(meta);
    body.appendChild(preview);

    item.appendChild(marker);
    item.appendChild(body);
    if (message.id === emailViewState.selectedMessageId){
      const actions = document.createElement('div');
      actions.className = 'graph-inline-actions email-message-actions';

      const openBtn = document.createElement('button');
      openBtn.type = 'button';
      openBtn.className = 'btn alt';
      openBtn.dataset.emailAction = 'open';
      openBtn.textContent = 'Åpne';

      const markBtn = document.createElement('button');
      markBtn.type = 'button';
      markBtn.className = 'btn alt';
      markBtn.dataset.emailAction = 'mark-read';
      markBtn.textContent = message.isRead ? 'Marker ulest' : 'Marker lest';

      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'btn danger';
      deleteBtn.dataset.emailAction = 'delete';
      deleteBtn.textContent = 'Slett';

      actions.append(openBtn, markBtn, deleteBtn);
      item.appendChild(actions);
    }
    list.appendChild(item);
  });
}

export function selectEmailMessage(id, emailViewState, renderMessages){
  const nextId = String(id || '').trim();
  emailViewState.selectedMessageId = emailViewState.selectedMessageId === nextId ? '' : nextId;
  renderMessages(emailViewState.messages);
}
