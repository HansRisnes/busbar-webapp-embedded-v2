import {
  $
} from './dom.js';

function getFallbackCalendarVisualKind(event, helpers = {}){
  const helperKind = String(helpers.getCalendarEventVisualKind?.(event) || '').trim();
  if (helperKind) return helperKind;
  const subject = String(event?.subject || '').trim().toLowerCase();
  const categories = Array.isArray(event?.categories)
    ? event.categories.map(item=>String(item || '').trim().toLowerCase())
    : [];
  if (categories.includes('to-do fullført')) return 'todo-completed';
  if (categories.includes('to-do') || subject.startsWith('to-do:') || subject.startsWith('todo:')) return 'todo';
  if (categories.includes('busbar prosjekt')) return 'project';
  return '';
}

function applyCalendarVisualKind(item, visualKind){
  const kind = String(visualKind || '').trim();
  if (!kind) return;
  item.dataset.calendarKind = kind;
  const colors = {
    project: {
      border: 'rgba(254,190,16,0.86)',
      background: 'rgba(254,190,16,0.72)'
    },
    todo: {
      border: 'rgba(248,113,113,0.86)',
      background: 'rgba(220,38,38,0.72)'
    },
    'todo-completed': {
      border: 'rgba(74,222,128,0.86)',
      background: 'rgba(34,197,94,0.72)'
    }
  };
  const color = colors[kind];
  if (!color) return;
  item.style.setProperty('border', `1px solid ${color.border}`, 'important');
  item.style.setProperty('background', color.background, 'important');
  item.style.setProperty('color', '#fff', 'important');
}

export function renderCalendarEvents(events, helpers = {}){
  const list = $('calendarEventsList');
  if (!list) return;
  list.innerHTML = '';
  if (!events.length){
    const empty = document.createElement('div');
    empty.className = 'graph-empty';
    empty.textContent = 'Ingen kommende kalenderhendelser funnet.';
    list.appendChild(empty);
    return;
  }
  const formatGraphDateTime = helpers.formatGraphDateTime || (value=>value || '');
  const getCalendarEventVisualKind = helpers.getCalendarEventVisualKind || (()=>'');
  const isCalendarTodoEvent = helpers.isCalendarTodoEvent || (()=>false);
  const isCalendarTodoCompletedEvent = helpers.isCalendarTodoCompletedEvent || (()=>false);
  events.forEach(event=>{
    const visualKind = getFallbackCalendarVisualKind(event, helpers);
    const isTodo = visualKind === 'todo' || visualKind === 'todo-completed' || isCalendarTodoEvent(event);
    const item = document.createElement('article');
    item.className = 'graph-item calendar-event-item';
    item.classList.toggle('is-todo', isTodo);
    item.classList.toggle('is-todo-completed', visualKind === 'todo-completed' || isCalendarTodoCompletedEvent(event));
    applyCalendarVisualKind(item, visualKind);
    item.dataset.calendarSubject = event.subject || '';

    const time = document.createElement('div');
    time.className = 'graph-item-time';
    time.textContent = `${formatGraphDateTime(event.start)} - ${formatGraphDateTime(event.end)}`;

    const body = document.createElement('div');
    body.className = 'graph-item-body';
    const title = document.createElement('h3');
    title.textContent = event.subject || 'Uten tittel';
    const meta = document.createElement('p');
    meta.className = 'muted-text';
    const organizer = event.organizer?.emailAddress?.name || event.organizer?.emailAddress?.address || '';
    const location = event.location?.displayName || '';
    meta.textContent = [organizer ? `Arrangør: ${organizer}` : '', location ? `Sted: ${location}` : ''].filter(Boolean).join(' | ');
    body.appendChild(title);
    body.appendChild(meta);
    const actions = document.createElement('div');
    actions.className = 'graph-card-actions';
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'btn alt btn-small';
    editBtn.textContent = 'Rediger';
    editBtn.dataset.editCalendarEvent = event.id || '';
    actions.appendChild(editBtn);
    body.appendChild(actions);

    item.appendChild(time);
    item.appendChild(body);
    list.appendChild(item);
  });
}

export function renderCalendarGrid(events, calendarViewState, helpers = {}){
  const grid = $('calendarGridView');
  if (!grid) return;
  grid.innerHTML = '';
  const mode = calendarViewState.mode;
  const range = helpers.calendarRangeForState();
  const days = [];
  const dayCount = mode === 'month' ? 42 : 7;
  for (let index = 0; index < dayCount; index += 1){
    days.push(helpers.addDays(range.start, index));
  }

  const weekdays = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag', 'Søndag'];
  const header = document.createElement('div');
  header.className = 'calendar-grid-header';
  if (mode === 'month' || mode === 'week'){
    const weekHeader = document.createElement('div');
    weekHeader.className = 'calendar-week-header-spacer';
    weekHeader.textContent = '';
    header.appendChild(weekHeader);
  }
  weekdays.forEach(day=>{
    const cell = document.createElement('div');
    cell.textContent = day;
    header.appendChild(cell);
  });
  grid.appendChild(header);

  const body = document.createElement('div');
  body.className = `calendar-grid calendar-grid-${mode}`;
  const today = helpers.startOfDay(new Date());
  const projectFlowSegments = [];
  const projectFlowRowLanes = new Map();
  const addProjectFlowSegment = (event, startIndex, endIndex, lane)=>{
    const item = document.createElement('button');
    item.className = 'calendar-grid-event is-project-flow-span';
    applyCalendarVisualKind(item, 'project');
    item.type = 'button';
    item.dataset.editCalendarEvent = event.id || '';
    item.style.gridRow = String(Math.floor(startIndex / 7) + 1);
    item.style.gridColumn = `${(startIndex % 7) + 2} / ${(endIndex % 7) + 3}`;
    item.style.setProperty('--calendar-event-lane', String(lane));
    item.textContent = event.subject || 'Uten tittel';
    body.appendChild(item);
  };

  events
    .filter(event=>helpers.getCalendarEventProjectFlowTaskId(event))
    .sort((a, b)=>{
      const left = helpers.getCalendarEventDisplayDates(a);
      const right = helpers.getCalendarEventDisplayDates(b);
      const leftDuration = left ? helpers.getCalendarDayDiff(left.startDay, left.endDay) : 0;
      const rightDuration = right ? helpers.getCalendarDayDiff(right.startDay, right.endDay) : 0;
      if (rightDuration !== leftDuration) return rightDuration - leftDuration;
      return (left?.startDay?.getTime() || 0) - (right?.startDay?.getTime() || 0);
    })
    .forEach(event=>{
      const dates = helpers.getCalendarEventDisplayDates(event);
      if (!dates) return;
      const visibleStart = days.findIndex(day=>helpers.sameCalendarDay(day, dates.startDay));
      const visibleEnd = days.findIndex(day=>helpers.sameCalendarDay(day, dates.endDay));
      const startIndex = visibleStart >= 0
        ? visibleStart
        : (dates.startDay < days[0] ? 0 : -1);
      const endIndex = visibleEnd >= 0
        ? visibleEnd
        : (dates.endDay > days[days.length - 1] ? days.length - 1 : -1);
      if (startIndex < 0 || endIndex < 0 || endIndex < startIndex) return;
      for (let rowStart = startIndex; rowStart <= endIndex; rowStart = (Math.floor(rowStart / 7) + 1) * 7){
        const rowEnd = Math.min(endIndex, (Math.floor(rowStart / 7) * 7) + 6);
        const weekRow = Math.floor(rowStart / 7);
        const lane = projectFlowSegments.filter(segment=>
          segment.weekRow === weekRow
          && rowStart <= segment.endIndex
          && rowEnd >= segment.startIndex
        ).length;
        projectFlowSegments.push({ weekRow, startIndex: rowStart, endIndex: rowEnd, lane });
        projectFlowRowLanes.set(weekRow, Math.max(projectFlowRowLanes.get(weekRow) || 0, lane + 1));
        addProjectFlowSegment(event, rowStart, rowEnd, lane);
      }
    });

  days.forEach((day, index)=>{
    if ((mode === 'month' || mode === 'week') && index % 7 === 0){
      const week = document.createElement('div');
      week.className = 'calendar-week-number';
      week.textContent = `UKE ${helpers.getIsoWeekNumber(day)}`;
      week.style.gridColumn = '1';
      week.style.gridRow = String(Math.floor(index / 7) + 1);
      body.appendChild(week);
    }
    const cell = document.createElement('section');
    cell.className = 'calendar-day-cell';
    cell.style.gridColumn = String((index % 7) + 2);
    cell.style.gridRow = String(Math.floor(index / 7) + 1);
    cell.classList.toggle('is-outside-month', mode === 'month' && day.getMonth() !== calendarViewState.cursor.getMonth());
    cell.classList.toggle('is-today', helpers.sameCalendarDay(day, today));
    cell.classList.toggle('is-past-day', day < today);

    const heading = document.createElement('div');
    heading.className = 'calendar-day-heading';
    heading.textContent = new Intl.DateTimeFormat('no-NO', { day: '2-digit', month: mode === 'week' ? 'short' : undefined }).format(day);
    const rowLaneCount = projectFlowRowLanes.get(Math.floor(index / 7)) || 0;
    if (rowLaneCount){
      heading.style.marginBottom = `${6 + (rowLaneCount * 31)}px`;
    }
    cell.appendChild(heading);

    const dayEvents = events
      .filter(event=>!helpers.getCalendarEventProjectFlowTaskId(event))
      .filter(event=>{
        const start = helpers.parseGraphDate(event.start);
        return start && helpers.sameCalendarDay(start, day);
      })
      .sort((a, b)=>(helpers.parseGraphDate(a.start)?.getTime() || 0) - (helpers.parseGraphDate(b.start)?.getTime() || 0));

    if (!dayEvents.length){
      const empty = document.createElement('div');
      empty.className = 'calendar-day-empty';
      empty.textContent = '';
      cell.appendChild(empty);
    }

    dayEvents.forEach(event=>{
      const visualKind = getFallbackCalendarVisualKind(event, helpers);
      const isTodo = visualKind === 'todo' || visualKind === 'todo-completed' || Boolean(helpers.isCalendarTodoEvent?.(event));
      const item = document.createElement('button');
      item.className = 'calendar-grid-event';
      item.classList.toggle('is-todo', isTodo);
      item.classList.toggle('is-todo-completed', visualKind === 'todo-completed' || Boolean(helpers.isCalendarTodoCompletedEvent?.(event)));
      applyCalendarVisualKind(item, visualKind);
      item.dataset.calendarSubject = event.subject || '';
      item.type = 'button';
      item.dataset.editCalendarEvent = event.id || '';
      const time = helpers.parseGraphDate(event.start);
      const timeLabel = time ? new Intl.DateTimeFormat('no-NO', { hour: '2-digit', minute: '2-digit' }).format(time) : '';
      item.textContent = [timeLabel, event.subject || 'Uten tittel'].filter(Boolean).join(' ');
      cell.appendChild(item);
    });

    body.appendChild(cell);
  });
  grid.appendChild(body);
}

export function updateCalendarViewControls(calendarViewState, formatCalendarPeriodLabel){
  const label = $('calendarPeriodLabel');
  if (label) label.textContent = formatCalendarPeriodLabel();
  const select = $('calendarModeSelect');
  if (select) select.value = calendarViewState.mode;
}

export function renderCalendarView(calendarViewState, helpers = {}){
  const list = $('calendarEventsList');
  const grid = $('calendarGridView');
  const isList = calendarViewState.mode === 'list';
  if (list) list.hidden = !isList;
  if (grid) grid.hidden = isList;
  updateCalendarViewControls(calendarViewState, helpers.formatCalendarPeriodLabel);
  if (isList){
    renderCalendarEvents(calendarViewState.events, helpers);
  } else {
    renderCalendarGrid(calendarViewState.events, calendarViewState, helpers);
  }
}
