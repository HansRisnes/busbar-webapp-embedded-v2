import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const userEmail = 'flow-test@busbar.no';
const responsibleEmail = 'flow-responsible@busbar.no';

function dateKey(daysFromToday){
  const date = new Date();
  date.setDate(date.getDate() + daysFromToday);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function initialProject(){
  return {
    id: 'flow-test-project',
    projectNumber: 'TEST-001',
    name: 'Scrolltest',
    projectOwnerEmail: responsibleEmail,
    projectOwnerName: 'Ansvarlig',
    projectStatus: 'unresolved',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    lines: [],
    projectFlowMilestones: [{
      id: 'flow-test-task',
      phaseId: 'request',
      title: 'Foresporsel',
      startDate: dateKey(-90),
      endDate: dateKey(-88),
      durationValue: 3,
      durationUnit: 'days',
      calendarEventId: 'existing-event',
      calendarOwnerEmail: responsibleEmail,
      completed: false,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }]
  };
}

for (const viewportWidth of [1100, 1440, 1920]){
test(`project flow keeps both scroll axes at ${viewportWidth}px`, async t=>{
  let projects = [initialProject()];
  let projectSyncCount = 0;
  let calendarSyncCount = 0;
  const calendarSyncActions = [];
  let delayNextProjectReadMs = 0;
  const server = createServer(async (request, response)=>{
    const url = new URL(request.url, 'http://localhost');
    response.setHeader('Content-Type', 'application/json; charset=utf-8');
    if (url.pathname === '/api/user-projects' && request.method === 'GET'){
      if (delayNextProjectReadMs){
        const delay = delayNextProjectReadMs;
        delayNextProjectReadMs = 0;
        await new Promise(resolveDelay=>setTimeout(resolveDelay, delay));
      }
      response.end(JSON.stringify({
        updatedAt: new Date().toISOString(),
        projects,
        ownerEmails: [userEmail, responsibleEmail],
        canViewAllProjects: true,
        canEditAllProjects: true
      }));
      return;
    }
    if (url.pathname === '/api/user-projects/sync' && request.method === 'POST'){
      let body = '';
      for await (const chunk of request) body += chunk;
      projects = JSON.parse(body).projects;
      projectSyncCount++;
      response.end(JSON.stringify({
        projects,
        ownerEmails: [userEmail, responsibleEmail],
        canViewAllProjects: true,
        canEditAllProjects: true
      }));
      return;
    }
    if (url.pathname === '/api/project-flow/calendar-sync' && request.method === 'POST'){
      let body = '';
      for await (const chunk of request) body += chunk;
      const payload = JSON.parse(body);
      calendarSyncCount++;
      calendarSyncActions.push(payload.action);
      response.end(JSON.stringify(payload.action === 'delete'
        ? { deleted: true }
        : { calendarEventId: 'flow-test-calendar-event', calendarOwnerEmail: responsibleEmail }));
      return;
    }
    if (url.pathname === '/api/customer-database'){
      response.end(JSON.stringify({ customers: [] }));
      return;
    }
    if (url.pathname.startsWith('/api/')){
      response.statusCode = 404;
      response.end('{}');
      return;
    }
    const file = resolve(root, `.${url.pathname === '/' ? '/index.html' : url.pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.statusCode = 403;
      response.end('{}');
      return;
    }
    try{
      const content = await readFile(file);
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml' };
      response.setHeader('Content-Type', mime[extname(file)] || 'application/octet-stream');
      response.end(content);
    }catch{
      response.statusCode = 404;
      response.end('{}');
    }
  });
  await new Promise(resolveListening=>server.listen(0, '127.0.0.1', resolveListening));
  t.after(()=>{
    server.closeAllConnections();
    return new Promise(resolveClose=>server.close(resolveClose));
  });

  const browser = await chromium.launch({ headless: true });
  t.after(()=>browser.close());
  const page = await browser.newPage({ viewport: { width: viewportWidth, height: 900 } });
  await page.addInitScript(email=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true,
      username: email,
      token: 'test-token',
      canViewAllProjects: true,
      canEditAllProjects: true
    }));
  }, userEmail);
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=project-flow`);
  await page.locator('#projectFlowProjectSelect option[value="flow-test-project"]').waitFor({ state: 'attached' });
  await page.locator('#projectFlowProjectSelect').selectOption('flow-test-project');
  await page.locator('[data-project-flow-task-bar="flow-test-task"]').waitFor();

  async function position(){
    return page.evaluate(()=>{
      const scroller = document.querySelector('.project-flow-scroller');
      const topScrollbar = document.querySelector('.project-flow-top-scrollbar');
      const firstDay = document.querySelector('.project-flow-day-head');
      return {
        left: scroller.scrollLeft,
        top: scroller.scrollTop,
        barLeft: topScrollbar.scrollLeft,
        windowY: window.scrollY,
        firstDayX: firstDay?.getBoundingClientRect().left - scroller.getBoundingClientRect().left,
        dayWidth: firstDay?.getBoundingClientRect().width
      };
    });
  }

  async function setPosition(){
    return page.evaluate(async ()=>{
      await new Promise(resolveFrame=>requestAnimationFrame(()=>requestAnimationFrame(resolveFrame)));
      const scroller = document.querySelector('.project-flow-scroller');
      const topScrollbar = document.querySelector('.project-flow-top-scrollbar');
      scroller.scrollLeft = 550;
      topScrollbar.scrollLeft = 550;
      scroller.scrollTop = 150;
      window.scrollTo(0, 300);
      await new Promise(resolveFrame=>requestAnimationFrame(()=>requestAnimationFrame(resolveFrame)));
      const firstDay = document.querySelector('.project-flow-day-head');
      return {
        left: scroller.scrollLeft,
        top: scroller.scrollTop,
        barLeft: topScrollbar.scrollLeft,
        windowY: window.scrollY,
        firstDayX: firstDay?.getBoundingClientRect().left - scroller.getBoundingClientRect().left,
        dayWidth: firstDay?.getBoundingClientRect().width
      };
    });
  }

  async function assertPosition(expected){
    await page.waitForTimeout(250);
    const actual = await position();
    for (const key of ['left', 'top', 'barLeft', 'windowY', 'firstDayX', 'dayWidth']){
      assert.equal(actual[key], expected[key], `${key} changed after project flow update`);
    }
  }

  async function waitForServerProject(predicate){
    const started = Date.now();
    while (!predicate(projects[0])){
      assert.ok(Date.now() - started < 3000, 'project sync did not finish');
      await new Promise(resolveWait=>setTimeout(resolveWait, 20));
    }
  }

  const editPosition = await setPosition();
  assert.ok(editPosition.left > 0 && editPosition.windowY > 0, 'test must scroll away from the default view');
  await page.locator('[data-project-flow-edit="flow-test-task"]').first().evaluate(element=>element.click());
  await page.locator('#projectFlowDurationValueInput').fill('4');
  await page.locator('#projectFlowMilestoneForm button[type="submit"]').click();
  await waitForServerProject(project=>project?.projectFlowMilestones?.some(task=>task.id === 'flow-test-task' && task.durationValue === 4));
  await assertPosition(editPosition);

  const createPosition = await setPosition();
  await page.locator('#addProjectFlowMilestoneBtn').evaluate(element=>element.click());
  await page.locator('#projectFlowTaskProjectSelect').selectOption('flow-test-project');
  await page.locator('#projectFlowPhaseSelect').selectOption('offer');
  await page.locator('#projectFlowDateInput').fill(dateKey(-30).split('-').reverse().join('/'));
  await page.locator('#projectFlowDurationValueInput').fill('2');
  await page.locator('#projectFlowMilestoneForm button[type="submit"]').click();
  await waitForServerProject(project=>project?.projectFlowMilestones?.length === 2);
  await page.waitForTimeout(250);
  const createResult = await page.evaluate(()=>({
    count: document.querySelectorAll('[data-project-flow-task-bar]').length,
    status: document.querySelector('#projectFlowStatus')?.textContent,
    modalOpen: !document.querySelector('#projectFlowMilestoneFormModal')?.hidden,
    valid: document.querySelector('#projectFlowMilestoneForm')?.checkValidity(),
    selectedProject: document.querySelector('#projectFlowProjectSelect')?.value,
    timelineText: document.querySelector('#projectFlowTimeline')?.textContent?.slice(0, 160)
  }));
  assert.equal(createResult.count, 2, JSON.stringify({ ...createResult, syncedProjects: projects.map(p=>({ id: p.id, tasks: p.projectFlowMilestones?.length })), projectSyncCount }));
  await assertPosition(createPosition);

  const deletePosition = await setPosition();
  await page.locator('[data-project-flow-edit="flow-test-task"]').first().evaluate(element=>element.click());
  page.once('dialog', dialog=>dialog.accept());
  await page.locator('#deleteProjectFlowMilestoneBtn').click();
  await waitForServerProject(project=>project?.projectFlowMilestones?.length === 1 && project.projectFlowMilestones[0].id !== 'flow-test-task');
  await page.waitForTimeout(250);
  const deleteResult = await page.evaluate(()=>({
    remaining: document.querySelectorAll('[data-project-flow-task-bar="flow-test-task"]').length,
    status: document.querySelector('#projectFlowStatus')?.textContent,
    modalOpen: !document.querySelector('#projectFlowMilestoneFormModal')?.hidden
  }));
  assert.equal(deleteResult.remaining, 0, JSON.stringify({ ...deleteResult, syncedProjects: projects.map(p=>({ id: p.id, tasks: p.projectFlowMilestones?.map(task=>task.id) })), projectSyncCount }));
  await assertPosition(deletePosition);
  assert.equal(calendarSyncCount, 0, 'calendar must remain unchanged while project flow is open');
  await page.locator('#projectFlowZoomInBtn').evaluate(element=>element.click());
  await page.waitForTimeout(100);
  const zoomPosition = await position();
  assert.ok(Math.abs(zoomPosition.left - deletePosition.left) > 10, 'zoom should return to the current week');

  delayNextProjectReadMs = 300;
  await page.locator('#refreshProjectFlowBtn').evaluate(element=>element.click());
  await page.waitForTimeout(40);
  const duringRefresh = await setPosition();
  await page.waitForTimeout(450);
  await assertPosition(duringRefresh);

  assert.ok(projectSyncCount >= 3, `expected background project syncs, got ${projectSyncCount}`);
  assert.deepEqual(calendarSyncActions.sort(), ['delete', 'upsert']);
  assert.equal(await page.evaluate(()=>localStorage.getItem('busbar.project.flow.calendar.pending.v1.flow-test@busbar.no')), null);
  assert.equal(projects[0].projectFlowMilestones[0].calendarEventId, 'flow-test-calendar-event');

  const remainingTaskId = projects[0].projectFlowMilestones[0].id;
  await page.locator(`[data-project-flow-edit="${remainingTaskId}"]`).first().evaluate(element=>element.click());
  await page.locator('#projectFlowDurationValueInput').fill('3');
  await page.locator('#projectFlowMilestoneForm button[type="submit"]').click();
  await waitForServerProject(project=>project?.projectFlowMilestones?.[0]?.durationValue === 3);
  assert.equal(calendarSyncCount, 2, 'calendar sync should still be deferred before leaving the page');
  await page.locator('[data-dashboard-nav="dashboard"]').evaluate(element=>element.click());
  await page.waitForFunction(()=>localStorage.getItem('busbar.project.flow.calendar.pending.v1.flow-test@busbar.no') === null);
  assert.equal(calendarSyncCount, 3, 'leaving project flow should send the pending calendar change');

  await page.locator('[data-dashboard-nav="project-flow"]').evaluate(element=>element.click());
  await page.waitForFunction(()=>document.querySelector('#projectFlowStatus')?.textContent?.includes('Oppdatert.'));
  await page.locator('#projectFlowProjectSelect').selectOption('flow-test-project');
  await page.locator(`[data-project-flow-edit="${remainingTaskId}"]`).first().evaluate(element=>element.click());
  await page.locator('#projectFlowDurationValueInput').fill('4');
  await page.locator('#projectFlowMilestoneForm button[type="submit"]').click();
  await waitForServerProject(project=>project?.projectFlowMilestones?.[0]?.durationValue === 4);
  assert.equal(calendarSyncCount, 3, 'calendar sync should remain deferred until browser refresh');
  await page.reload();
  await page.waitForFunction(()=>localStorage.getItem('busbar.project.flow.calendar.pending.v1.flow-test@busbar.no') === null);
  assert.equal(calendarSyncCount, 4, 'refresh should send the last calendar change');
});
}
