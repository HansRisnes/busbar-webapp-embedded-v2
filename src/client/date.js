const pad2 = value => String(value).padStart(2, '0');

export function formatDateTimeLocalInput(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  return `${value.getFullYear()}-${pad2(value.getMonth() + 1)}-${pad2(value.getDate())}T${pad2(value.getHours())}:${pad2(value.getMinutes())}`;
}

export function formatDateInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  return `${pad2(value.getDate())}/${pad2(value.getMonth() + 1)}/${value.getFullYear()}`;
}

export function formatIsoDateInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  return `${value.getFullYear()}-${pad2(value.getMonth() + 1)}-${pad2(value.getDate())}`;
}

export function formatTimeInputValue(date){
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return '';
  return `${pad2(value.getHours())}:${pad2(value.getMinutes())}`;
}

export function parseCalendarDateInputValue(value){
  const raw = String(value || '').trim();
  const match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return '';
  const [, day, month, year] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  if (
    date.getFullYear() !== Number(year) ||
    date.getMonth() !== Number(month) - 1 ||
    date.getDate() !== Number(day)
  ) return '';
  return `${year}-${month}-${day}`;
}

export function combineLocalDateAndTimeValue(dateValue, timeValue){
  const parsedDate = parseCalendarDateInputValue(dateValue);
  const time = String(timeValue || '').trim();
  return parsedDate && time ? `${parsedDate}T${time}` : '';
}

export function dateFromCalendarInputValue(value){
  const parsed = parseCalendarDateInputValue(value);
  return parsed ? new Date(`${parsed}T00:00`) : null;
}

export function parseGraphDateTimeValue(value, timeZone){
  const rawValue = String(value || '').trim();
  if (!rawValue) return null;
  const normalized = rawValue.replace(/(\.\d{3})\d+/, '$1');
  if (/[zZ]$|[+-]\d{2}:?\d{2}$/.test(normalized)){
    const date = new Date(normalized);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const tz = String(timeZone || '').trim().toLowerCase();
  if (tz === 'utc' || tz === 'etc/utc' || tz === 'gmt'){
    const date = new Date(`${normalized}Z`);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?/);
  if (match){
    const [, year, month, day, hour, minute, second = '0', ms = '0'] = match;
    return new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
      Number(ms.padEnd(3, '0'))
    );
  }
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? null : date;
}

export function parseGraphDate(raw){
  const value = String(raw?.dateTime || raw || '').trim();
  if (!value) return null;
  return parseGraphDateTimeValue(value, raw?.timeZone);
}

export function formatGraphDateTime(raw){
  const value = String(raw?.dateTime || raw || '').trim();
  if (!value) return '';
  const date = parseGraphDate(raw);
  if (!date || Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('no-NO', {
    weekday: 'short',
    day: '2-digit',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

export function startOfDay(date){
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

export function addDays(date, days){
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

export function addMonths(date, months){
  const copy = new Date(date);
  copy.setMonth(copy.getMonth() + months);
  return copy;
}

export function startOfWeekMonday(date){
  const copy = startOfDay(date);
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + diff);
  return copy;
}

export function startOfMonth(date){
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

export function endOfMonth(date){
  return new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59, 999);
}

export function getIsoWeekNumber(date){
  const copy = startOfDay(date);
  copy.setDate(copy.getDate() + 3 - ((copy.getDay() + 6) % 7));
  const weekOne = new Date(copy.getFullYear(), 0, 4);
  return 1 + Math.round(((copy - weekOne) / 86400000 - 3 + ((weekOne.getDay() + 6) % 7)) / 7);
}

export function sameCalendarDay(left, right){
  return left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate();
}

export function getCalendarDayDiff(start, end){
  const startUtc = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
  const endUtc = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((endUtc - startUtc) / 86400000);
}

export function parseProjectFlowDate(value){
  const raw = String(value || '').trim();
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const displayMatch = raw.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/);
  const match = isoMatch || displayMatch;
  if (!match) return null;
  const year = isoMatch ? Number(match[1]) : Number(match[3]);
  const month = Number(match[2]);
  const day = isoMatch ? Number(match[3]) : Number(match[1]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return Number.isNaN(date.getTime()) ? null : date;
}

export function formatProjectFlowDate(date){
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

export function formatProjectFlowDisplayDate(value){
  const date = parseProjectFlowDate(value);
  if (!date) return '-';
  return date.toLocaleDateString('no-NO', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export function formatProjectFlowInputDate(value){
  const date = value instanceof Date ? value : parseProjectFlowDate(value);
  if (!date) return '';
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
}

export function addProjectFlowDays(date, days){
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  copy.setDate(copy.getDate() + Number(days || 0));
  return copy;
}

export function addProjectFlowDuration(startDate, value, unit){
  const amount = Math.max(1, Number.parseInt(value, 10) || 1);
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  if (unit === 'weeks'){
    return addProjectFlowDays(start, (amount * 7) - 1);
  }
  if (unit === 'months'){
    const end = new Date(start.getFullYear(), start.getMonth() + amount, start.getDate());
    return addProjectFlowDays(end, -1);
  }
  return addProjectFlowDays(start, amount - 1);
}

export function getProjectFlowDayDiff(start, end){
  const startUtc = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
  const endUtc = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((endUtc - startUtc) / 86400000);
}

export function getProjectFlowWeekNumber(date){
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  return Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
}

export function getProjectFlowWeekKey(date){
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  return `${utc.getUTCFullYear()}-${pad2(getProjectFlowWeekNumber(date))}`;
}
