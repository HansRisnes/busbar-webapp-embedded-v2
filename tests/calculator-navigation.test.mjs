import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('dashboard and calculator restore the sidebar before the app script loads', async t=>{
  let delayProjectSync = false;
  let delayCsvResponse = false;
  let projectSyncCompletedAt = 0;
  let firstCsvRequestedAt = 0;
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      if (delayProjectSync && pathname === '/api/user-projects/sync'){
        await new Promise(resolveDelay=>setTimeout(resolveDelay, 600));
        projectSyncCompletedAt = Date.now();
      }
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects: [], ownerEmails: ['admin@busbar.no'], updatedAt: new Date().toISOString() }
        : { customers: [] }));
      return;
    }
    if (pathname.endsWith('.csv')){
      if (delayProjectSync && !firstCsvRequestedAt) firstCsvRequestedAt = Date.now();
      if (delayCsvResponse) await new Promise(resolveDelay=>setTimeout(resolveDelay, 600));
    }
    const file = resolve(root, `.${pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.writeHead(403).end();
      return;
    }
    try{
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.ttf': 'font/ttf', '.csv': 'text/csv' };
      response.setHeader('Content-Type', mime[extname(file)] || 'application/octet-stream');
      response.end(await readFile(file));
    }catch{
      response.writeHead(404).end();
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
      loggedIn: true,
      username: 'admin@busbar.no',
      token: 'test-token',
      isAdmin: true
    }));
    if (localStorage.getItem('busbar.dashboard.sidebar.collapsed.v1') === null) {
      localStorage.setItem('busbar.dashboard.sidebar.collapsed.v1', '0');
    }
  });
  const base = `http://127.0.0.1:${server.address().port}`;

  for (const [path, expectedCollapsed] of [
    ['/index.html?view=dashboard', false],
    ['/calculator.html', true]
  ]){
    if (expectedCollapsed) {
      await page.evaluate(()=>localStorage.setItem('busbar.dashboard.sidebar.collapsed.v1', '1'));
    }
    let releaseScript;
    const scriptGate = new Promise(resolveGate=>{ releaseScript = resolveGate; });
    await page.route('**/app.js', async route=>{
      await scriptGate;
      await route.continue();
    });
    try{
      await page.goto(`${base}${path}`, { waitUntil: 'commit' });
      await page.locator('#dashboardAdminNav').waitFor({ state: 'visible' });
      const shellState = await page.evaluate(()=>({
        collapsed: document.getElementById('dashboardShell').classList.contains('is-sidebar-collapsed'),
        locked: document.body.classList.contains('is-auth-locked'),
        adminHidden: document.getElementById('dashboardAdminNav').hidden,
        username: document.getElementById('authUser').textContent,
        loginHidden: document.getElementById('loginBtn').hidden,
        logoutHidden: document.getElementById('logoutBtn').hidden
      }));
      assert.deepEqual(shellState, {
        collapsed: expectedCollapsed,
        locked: false,
        adminHidden: false,
        username: 'admin@busbar.no',
        loginHidden: true,
        logoutHidden: false
      });
      const sidebarStyle = await page.locator('.dashboard-sidebar').evaluate(el=>({
        width: Math.round(el.getBoundingClientRect().width),
        transition: getComputedStyle(el).transitionDuration,
        labelOpacity: getComputedStyle(el.querySelector('.sidebar-nav-label')).opacity
      }));
      assert.equal(sidebarStyle.width, expectedCollapsed ? 72 : 240);
      assert.equal(sidebarStyle.transition, '0s');
      assert.equal(sidebarStyle.labelOpacity, expectedCollapsed ? '0' : '1');
      if (expectedCollapsed){
        assert.equal(await page.locator('#meter').evaluate(el=>el.parentElement.classList.contains('stepper')), true);
        assert.equal(await page.locator('#dist').inputValue(), 'Nei');
        assert.equal(await page.locator('#montasjeHoursPerMeter').textContent(), '–');
        assert.equal(await page.locator('#activeProjectName').textContent(), 'Ingen prosjekt valgt');
        assert.equal(await page.locator('#opphengRate').getAttribute('placeholder'), null);
        assert.equal(await page.locator('#opphengRate').inputValue(), '');
        for (const [id, rate] of [['montasjeHourlyRate', '890'], ['engineeringHourlyRate', '1200']]){
          assert.equal(await page.locator(`#${id}`).inputValue(), rate);
          assert.equal(await page.locator(`#${id}`).evaluate(el=>el.readOnly && el.classList.contains('locked')), true);
        }
      }
    } finally {
      releaseScript();
    }
    await page.waitForLoadState('load');
    await page.waitForFunction(()=>!document.body.classList.contains('is-shell-booting'));
    await page.unroute('**/app.js');
  }

  for (const view of [
    'projects', 'project-flow', 'offers', 'leads', 'email', 'calendar',
    'customers', 'company-card', 'busbar-folders', 'project-folders', 'supplier-folders'
  ]){
    let releaseScript;
    const scriptGate = new Promise(resolveGate=>{ releaseScript = resolveGate; });
    await page.route('**/app.js', async route=>{
      await scriptGate;
      await route.continue();
    });
    try{
      await page.goto(`${base}/index.html?view=${view}`, { waitUntil: 'commit' });
      await page.locator(`[data-dashboard-nav="${view}"].is-active`).waitFor();
      assert.equal(await page.locator('[data-dashboard-nav="dashboard"]').getAttribute('aria-current'), 'false');
      assert.equal(await page.locator('[data-dashboard-nav].is-active').count(), 1);
    }finally{
      releaseScript();
    }
    await page.waitForLoadState('load');
    await page.unroute('**/app.js');
  }

  const expectedSidebarLinks = await page.locator('[data-dashboard-nav], [data-dashboard-link]')
    .evaluateAll(items=>items.map(item=>item.dataset.dashboardNav || item.dataset.dashboardLink));
  await page.evaluate(()=>sessionStorage.setItem('busbar.admin.auth.v1', JSON.stringify({
    authHeader: 'Basic test', username: 'Admin Test'
  })));
  await page.goto(`${base}/admin.html`);
  const adminSidebarLinks = await page.locator('[data-dashboard-link]')
    .evaluateAll(items=>items.map(item=>item.dataset.dashboardLink));
  assert.deepEqual(adminSidebarLinks, [...expectedSidebarLinks, 'admin']);
  assert.equal(await page.locator('.sidebar-brand').getAttribute('href'), 'https://busbar.no');
  assert.equal(await page.locator('#authUser').textContent(), 'Admin Test');
  assert.equal(await page.locator('#logoutBtn').isVisible(), true);
  assert.equal(await page.locator('#dashboardShell').evaluate(el=>el.classList.contains('is-sidebar-collapsed')), true);
  let releaseAdminScript;
  const adminScriptGate = new Promise(resolveGate=>{ releaseAdminScript = resolveGate; });
  await page.route('**/admin.js', async route=>{
    await adminScriptGate;
    await route.continue();
  });
  try{
    await page.reload({ waitUntil: 'commit' });
    await page.locator('#authUser').getByText('Admin Test').waitFor({ state: 'attached' });
    assert.equal(await page.locator('#authUser').textContent(), 'Admin Test');
    assert.equal(await page.locator('#authUser').evaluate(el=>el.hidden), false);
    assert.equal(await page.locator('#dashboardShell').evaluate(el=>el.classList.contains('is-sidebar-collapsed')), true);
  }finally{
    releaseAdminScript();
  }
  await page.waitForLoadState('load');
  await page.unroute('**/admin.js');
  await page.locator('#logoutBtn').click();
  assert.equal(await page.locator('#loginBtn').isVisible(), true);
  assert.equal(await page.locator('#authUser').isVisible(), false);
  await page.locator('#loginBtn').click();
  assert.equal(await page.locator('#adminUsername').evaluate(el=>document.activeElement === el), true);

  delayProjectSync = true;
  delayCsvResponse = true;
  await page.goto(`${base}/calculator.html`);
  const initialCalculatorState = await page.evaluate(()=>( {
    meterStepper: document.getElementById('meter').parentElement.classList.contains('stepper'),
    ampDisabled: document.getElementById('ampSelect').disabled,
    ampText: document.getElementById('ampSelect').selectedOptions[0]?.textContent,
    distribution: document.getElementById('dist').value,
    montageHours: document.getElementById('montasjeHoursPerMeter').textContent,
    engineeringHours: document.getElementById('engineeringHoursPerMeter').textContent
  }));
  assert.equal(initialCalculatorState.meterStepper, true);
  assert.equal(initialCalculatorState.ampDisabled, true);
  assert.equal(initialCalculatorState.ampText, 'Velg…');
  assert.equal(initialCalculatorState.distribution, 'Nei');
  assert.equal(initialCalculatorState.montageHours, '–');
  assert.notEqual(initialCalculatorState.engineeringHours, '-');
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  await page.locator('#meter + [data-step="plus"]').click();
  assert.equal(await page.locator('#meter').inputValue(), '1');
  await page.locator('#rateToggle').check();
  await page.locator('#engineeringRateToggle').check();
  assert.equal(await page.locator('#montasjeHourlyRate').evaluate(el=>el.readOnly), false);
  assert.equal(await page.locator('#engineeringHourlyRate').evaluate(el=>el.readOnly), false);
  assert.ok(firstCsvRequestedAt > 0);
  assert.ok(projectSyncCompletedAt > 0);
  assert.ok(firstCsvRequestedAt < projectSyncCompletedAt,
    'Price files should load while project synchronization is pending');
});
