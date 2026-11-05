const SALES_PHASES = new Set(['new', 'contacted', 'qualified', 'offer_sent', 'won', 'lost']);

function leadError(message, statusCode = 400) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function text(value, maxLength = 500) {
  const result = String(value ?? '').trim();
  if (result.length > maxLength) throw leadError(`Feltet kan ikke overstige ${maxLength} tegn`);
  return result;
}

function email(value) {
  return text(value, 254).toLowerCase();
}

function date(value, required = false) {
  const result = text(value, 10);
  if (!result && !required) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(result)) throw leadError('Oppgi en gyldig dato');
  const parsed = new Date(`${result}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== result) {
    throw leadError('Oppgi en gyldig dato');
  }
  return result;
}

function optionalNumber(value, max) {
  if (value === '' || value === null || value === undefined) return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > max) throw leadError('Oppgi et gyldig tall');
  return parsed;
}

function normalizeLeadInput(raw, existing, actorEmail, canAssignOthers, registeredEmails) {
  const companyName = text(raw?.companyName, 200);
  const projectNeed = text(raw?.projectNeed, 500);
  const nextActivity = text(raw?.nextActivity, 500);
  const followUpDate = date(raw?.followUpDate, true);
  const salesPhase = text(raw?.salesPhase, 30) || 'new';
  const contactName = text(raw?.contactName, 200);
  const contactEmail = email(raw?.contactEmail);
  const contactPhone = text(raw?.contactPhone, 80);
  if (!companyName || !projectNeed || !nextActivity || !contactName || !contactEmail || !contactPhone) {
    throw leadError('Alle feltene i hoveddelen må fylles ut');
  }
  if (!SALES_PHASES.has(salesPhase)) throw leadError('Ugyldig salgsfase');

  const responsibleEmail = email(raw?.responsibleEmail || existing?.responsibleEmail || actorEmail);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(responsibleEmail)) throw leadError('Ugyldig ansvarlig');
  if (!canAssignOthers && responsibleEmail !== actorEmail && responsibleEmail !== existing?.responsibleEmail) {
    throw leadError('Du kan bare tildele leads til deg selv', 403);
  }
  if (canAssignOthers && responsibleEmail !== existing?.responsibleEmail && !registeredEmails.has(responsibleEmail)) {
    throw leadError('Ansvarlig må være en registrert bruker');
  }
  if (contactEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contactEmail)) {
    throw leadError('Ugyldig e-postadresse');
  }

  const previousUpdate = Date.parse(existing?.updatedAt || '') || 0;
  const now = new Date(Math.max(Date.now(), previousUpdate + 1)).toISOString();
  return {
    id: existing?.id || '',
    companyName,
    contactName,
    contactEmail,
    contactPhone,
    projectNeed,
    salesPhase,
    nextActivity,
    followUpDate,
    responsibleEmail,
    estimatedValue: optionalNumber(raw?.estimatedValue, 1e12),
    decisionDate: date(raw?.decisionDate),
    projectPhase: text(raw?.projectPhase, 100),
    customerType: text(raw?.customerType, 100),
    leadSource: text(raw?.leadSource, 100),
    probability: optionalNumber(raw?.probability, 100),
    comments: text(raw?.comments, 5000),
    followUpCompleted: raw?.followUpCompleted === true || (existing?.followUpCompleted === true && raw?.followUpCompleted !== false),
    calendarEventId: existing?.calendarEventId || '',
    calendarOwnerEmail: existing?.calendarOwnerEmail || '',
    calendarSyncError: existing?.calendarSyncError || '',
    createdByEmail: existing?.createdByEmail || actorEmail,
    createdAt: existing?.createdAt || now,
    updatedAt: now
  };
}

function canViewLead(lead, auth) {
  return auth.canViewAllProjects === true
    || lead.responsibleEmail === auth.email
    || lead.createdByEmail === auth.email;
}

function canEditLead(lead, auth) {
  return auth.canEditAllProjects === true
    || lead.responsibleEmail === auth.email
    || lead.createdByEmail === auth.email;
}

module.exports = { normalizeLeadInput, canViewLead, canEditLead };
