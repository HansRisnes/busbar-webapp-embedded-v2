import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import { buildDashboardFlowAlerts } from '../src/client/dashboard.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('dashboard alerts include overdue and upcoming unfinished project flow tasks', ()=>{
  const projects = [{
    id: 'project-1',
    projectNumber: '2026-1054',
    name: 'Test',
    projectFlowMilestones: [
      { id: 'overdue', phaseId: 'request', endDate: '2026-11-01', completed: false },
      { id: 'today', phaseId: 'request', endDate: '2026-11-03', completed: false },
      { id: 'last-day', phaseId: 'request', endDate: '2026-11-10', completed: false },
      { id: 'later', phaseId: 'request', endDate: '2026-11-11', completed: false },
      { id: 'completed', phaseId: 'request', endDate: '2026-10-30', completed: true },
      { id: 'invalid', phaseId: 'request', endDate: 'not-a-date', completed: false }
    ]
  }];

  const alerts = buildDashboardFlowAlerts(projects, new Date(2026, 10, 3, 14));
  assert.deepEqual(alerts.map(item=>item.taskId), ['overdue', 'today', 'last-day']);
  assert.deepEqual(alerts.map(item=>item.daysUntil), [-2, 0, 7]);
  assert.ok(alerts.every(item=>item.projectTitle === '2026-1054 - Test'));
});

test('dashboard alerts include only unfinished leads due within seven days', ()=>{
  const leads = [
    { id: 'late', companyName: 'Kunde A', projectNeed: 'Ny strømskinne', nextActivity: 'Oppfølging', followUpDate: '2026-11-02' },
    { id: 'soon', companyName: 'Kunde B', nextActivity: 'Møte', followUpDate: '2026-11-10' },
    { id: 'later', companyName: 'Kunde C', followUpDate: '2026-11-12' },
    { id: 'done', companyName: 'Kunde D', followUpDate: '2026-11-01', followUpCompleted: true }
  ];
  const alerts = buildDashboardFlowAlerts([], new Date(2026, 10, 4, 12), leads);
  assert.deepEqual(alerts.map(item=>item.leadId), ['late', 'soon']);
  assert.deepEqual(alerts.map(item=>item.daysUntil), [-2, 6]);
  assert.ok(alerts.every(item=>item.kind === 'lead'));
  assert.equal(alerts[0].projectTitle, 'Ny strømskinne');
});

test('dashboard alert strip opens the due-task list and its project flow', async t=>{
  const dateKey = offset=>{
    const date = new Date();
    date.setHours(12, 0, 0, 0);
    date.setDate(date.getDate() + offset);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  };
  const projects = [{
    id: 'flow-1', name: 'Varseltest', projectNumber: '2026-1054',
    projectOwnerEmail: 'alerts@busbar.no', lines: [],
    createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    projectFlowMilestones: [
      { id: 'late', phaseId: 'request', startDate: dateKey(-2), endDate: dateKey(-1), completed: false },
      { id: 'soon', phaseId: 'offer', startDate: dateKey(1), endDate: dateKey(3), completed: false },
      { id: 'done', phaseId: 'order', startDate: dateKey(-2), endDate: dateKey(-1), completed: true }
    ]
  }];
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects, ownerEmails: ['alerts@busbar.no'], updatedAt: new Date().toISOString() }
        : { customers: [] }));
      return;
    }
    const file = resolve(root, `.${pathname === '/' ? '/index.html' : pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.statusCode = 403;
      response.end();
      return;
    }
    try{
      const content = await readFile(file);
      response.setHeader('Content-Type', `${({ '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.ttf': 'font/ttf' })[extname(file)] || 'application/octet-stream'}; charset=utf-8`);
      response.end(content);
    }catch{
      response.statusCode = 404;
      response.end();
    }
  });
  await new Promise(resolveListening=>server.listen(0, '127.0.0.1', resolveListening));
  t.after(()=>{
    server.closeAllConnections();
    return new Promise(resolveClose=>server.close(resolveClose));
  });
  const browser = await chromium.launch({ headless: true });
  t.after(()=>browser.close());
  const page = await browser.newPage();
  await page.addInitScript(()=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true, username: 'alerts@busbar.no', token: 'test-token'
    }));
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=dashboard`);
  await page.waitForFunction(()=>document.querySelectorAll('.dashboard-flow-alert-cell').length === 2);
  assert.match(await page.locator('.dashboard-flow-alert-cell').first().textContent(), /2026-1054 - Varseltest/);
  assert.match(await page.locator('.dashboard-flow-alert-cell').first().textContent(), /Forfalt for 1 dag siden/);
  assert.match(await page.locator('.dashboard-flow-alert-cell').last().textContent(), /Frist om 3 dager/);
  assert.ok(await page.locator('.dashboard-flow-alert-cell').first().evaluate(el=>el.classList.contains('is-overdue')));
  assert.ok(await page.locator('.dashboard-flow-alert-cell').last().evaluate(el=>el.classList.contains('is-upcoming')));
  const positions = await page.evaluate(()=>({
    strip: document.querySelector('#dashboardFlowAlerts').getBoundingClientRect().bottom,
    grid: document.querySelector('.main-dashboard-grid').getBoundingClientRect().top
  }));
  assert.ok(positions.strip <= positions.grid);
  const checkButtonPlacement = async()=>{
    const geometry = await page.evaluate(()=>{
      const rect = selector=>document.querySelector(selector).getBoundingClientRect();
      return {
        strip: rect('#dashboardFlowAlerts'),
        cells: rect('.dashboard-flow-alert-cells'),
        button: rect('[data-dashboard-flow-alerts-open]')
      };
    });
    if (page.viewportSize().width <= 760) assert.ok(geometry.button.bottom <= geometry.cells.top);
    else assert.ok(geometry.button.left >= geometry.cells.right - 1);
    assert.ok(geometry.button.right <= geometry.strip.right);
  };
  await checkButtonPlacement();
  if (process.env.SAVE_DASHBOARD_SCREENSHOTS === '1'){
    await page.locator('#dashboardFlowAlerts').screenshot({ path: join(tmpdir(), 'busbar-alert-strip-desktop.png') });
  }
  await page.locator('.dashboard-flow-alert-cell').first().click();
  await page.waitForFunction(()=>document.querySelector('#projectFlowProjectSelect')?.value === 'flow-1');

  await page.setViewportSize({ width: 390, height: 800 });
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=dashboard`);
  await page.waitForFunction(()=>document.querySelectorAll('.dashboard-flow-alert-cell').length === 2);
  await checkButtonPlacement();
  if (process.env.SAVE_DASHBOARD_SCREENSHOTS === '1'){
    await page.locator('#dashboardFlowAlerts').screenshot({ path: join(tmpdir(), 'busbar-alert-strip-mobile.png') });
  }
  await page.getByRole('button', { name: 'Vis alle' }).click();
  assert.equal(await page.locator('#dashboardFlowAlertsList button').count(), 2);
  const overdueDue = await page.locator('#dashboardFlowAlertsList .dashboard-flow-alert-due').first().textContent();
  const upcomingDue = await page.locator('#dashboardFlowAlertsList .dashboard-flow-alert-due').last().textContent();
  assert.match(overdueDue, /^Forfalt for 1 dag siden · \d{2}\.\d{2}\.\d{4}$/);
  assert.match(upcomingDue, /^Frist om 3 dager · \d{2}\.\d{2}\.\d{4}$/);
  await page.locator('#dashboardFlowAlertsList button').first().click();
  await page.waitForFunction(()=>document.querySelector('#projectFlowProjectSelect')?.value === 'flow-1');
  assert.equal(await page.locator('#dashboardFlowAlertsModal').evaluate(el=>el.style.display), 'none');
});
