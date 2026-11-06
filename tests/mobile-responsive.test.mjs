import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const email = 'mobile-test@busbar.no';

async function setup(t){
  let projects = [{
    id: 'mobile-project', name: 'Mobilprosjekt med et langt navn for å kontrollere linjebryting',
    customer: 'Kunde med et langt firmanavn AS', contactPerson: 'Testperson',
    projectOwnerEmail: email, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    lines: [{ id: 'line-1', lineNumber: 'L-1', name: 'En lang prosjektlinje', totals: { totalExMontasje: 1000 } }],
    projectFlowMilestones: [{ id: 'task-1', phaseId: 'request', startDate: '2026-11-06', endDate: '2026-11-09', completed: false }]
  }];
  const calendarRequests = [];
  const server = createServer(async(request, response)=>{
    const url = new URL(request.url, 'http://localhost');
    if (url.pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      if (url.pathname === '/api/user-projects/sync' && request.method === 'POST'){
        let body = '';
        for await (const chunk of request) body += chunk;
        projects = JSON.parse(body).projects;
      }
      let payload = { customers: [], contacts: [] };
      if (url.pathname.startsWith('/api/user-projects')) payload = { projects, ownerEmails: [email], updatedAt: new Date().toISOString() };
      if (url.pathname === '/api/auth/microsoft/config') payload = { enabled: true, clientId: 'test', tenantId: 'test' };
      if (url.pathname === '/api/market-data') payload = {
        aluminium: { price: 2800, currency: 'USD', symbol: 'MANUAL', source: 'Manuell', unit: 't', updatedAt: new Date().toISOString() },
        fx: { usdNok: { rate: 10, date: '2026-11-06' } }, updatedAt: new Date().toISOString()
      };
      response.end(JSON.stringify(payload));
      return;
    }
    const file = resolve(root, `.${url.pathname === '/' ? '/index.html' : url.pathname}`);
    if (!file.startsWith(`${root}${sep}`)){ response.writeHead(403).end(); return; }
    try{
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.csv': 'text/csv', '.ttf': 'font/ttf', '.png': 'image/png' };
      response.setHeader('Content-Type', mime[extname(file)] || 'application/octet-stream');
      response.end(await readFile(file));
    } catch { response.writeHead(404).end(); }
  });
  await new Promise(done=>server.listen(0, '127.0.0.1', done));
  t.after(()=>{ server.closeAllConnections(); return new Promise(done=>server.close(done)); });
  const browser = await chromium.launch({ headless: true });
  t.after(()=>browser.close());
  const base = `http://127.0.0.1:${server.address().port}`;
  async function pageFor(width, loggedIn = true){
    const page = await browser.newPage({ viewport: { width, height: 844 }, isMobile: width <= 760, hasTouch: width <= 760 });
    await page.route('**/assets/vendor/msal-browser.min.js', route=>route.fulfill({
      contentType: 'text/javascript', body: `window.msal = { PublicClientApplication: class {
        getAllAccounts(){ return [{ username: '${email}' }]; }
        setActiveAccount(){}
        async acquireTokenSilent(){ return { accessToken: 'test-graph-token' }; }
      }};`
    }));
    await page.route('https://graph.microsoft.com/**', async route=>{
      const url = new URL(route.request().url());
      if (url.pathname.endsWith('/calendarView')) calendarRequests.push(url.searchParams.get('startDateTime'));
      await route.fulfill({ json: { value: url.pathname.endsWith('/calendarView') ? [{
        id: 'event-1', subject: 'Mobilavtale med kunde', start: { dateTime: '2026-11-06T10:00:00' }, end: { dateTime: '2026-11-06T11:00:00' }
      }] : [] } });
    });
    if (loggedIn) await page.addInitScript(({ email })=>{
      sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({ loggedIn: true, username: email, token: 'test-token' }));
    }, { email });
    return page;
  }
  return { base, pageFor, projects: ()=>projects, calendarRequests };
}

async function fits(page, label){
  const size = await page.evaluate(()=>({ width: innerWidth, scroll: document.documentElement.scrollWidth,
    outside: [...document.querySelectorAll('body *')].filter(el=>{
      const box = el.getBoundingClientRect();
      return box.width && box.right > innerWidth + 1 && getComputedStyle(el).position !== 'absolute'
        && !el.closest('.project-flow-scroller, .calendar-grid-view, .admin-table-wrap');
    }).slice(0, 8).map(el=>`${el.tagName}.${el.className}#${el.id}`)
  }));
  assert.ok(size.scroll <= size.width + 1, `${label}: page overflows ${JSON.stringify(size)}`);
  if (process.env.SAVE_MOBILE_SCREENSHOTS === '1' && ['projects 390', 'flow list', 'calendar agenda', 'calculator results'].includes(label)){
    await page.screenshot({ path: join(tmpdir(), `busbar-mobile-${label.replaceAll(' ', '-')}.png`), fullPage: true });
  }
}

async function checkMobileHeader(page){
  const header = page.locator('.mobile-app-header');
  assert.equal(await header.isVisible(), true);
  assert.equal(await header.locator('#dashboardSidebarToggle').isVisible(), true);
  assert.equal(await header.locator('img[alt="Busbar"]').isVisible(), true);
  assert.equal(await header.locator('img').evaluate(img=>img.complete && img.naturalWidth > 0), true);
  const active = page.locator('.sidebar-nav .is-active');
  assert.equal(await header.locator('.mobile-page-title').textContent(), await active.locator('.sidebar-nav-label').textContent());
  assert.equal(await header.locator('svg').count(), 0, 'the page icon is omitted from the header');
  const toggleBox = await header.locator('#dashboardSidebarToggle').boundingBox();
  const titleBox = await header.locator('.mobile-header-page').boundingBox();
  const logoBox = await header.locator('.mobile-header-brand').boundingBox();
  assert.ok(titleBox.x >= toggleBox.x + toggleBox.width + 16, 'page name has spacing after the menu button');
  assert.ok(logoBox.x >= titleBox.x + titleBox.width + 16, 'logo is to the right of the page name');
  assert.equal(Math.round(toggleBox.x), 12, 'menu button retains its original left inset');
  assert.equal(await page.evaluate(()=>innerWidth) - Math.round(logoBox.x + logoBox.width), 20, 'logo keeps its adjusted right inset');
  const boxes = await header.locator(':scope > *, .mobile-header-page > *').evaluateAll(items=>items.map(item=>{
    const r = item.getBoundingClientRect();
    return { left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: r.width, viewport: innerWidth };
  }));
  for (const box of boxes) assert.ok(box.width > 0 && box.left >= 0 && box.right <= box.viewport && box.top >= 0 && box.bottom <= 64, 'header content fits');
}

async function settleMenu(page){
  await page.locator('.dashboard-sidebar').evaluate(element=>Promise.all(
    element.getAnimations().map(animation=>animation.finished.catch(()=>{}))
  ));
}

test('mobile menu, pages and forms fit 320 to 760 px, preserving desktop preference', async t=>{
  const { base, pageFor } = await setup(t);
  for (const width of [320, 360, 390, 760]){
    const page = await pageFor(width);
    const errors = [];
    page.on('pageerror', err=>errors.push(err.message));
    await page.goto(`${base}/index.html?view=projects`);
    await page.locator('#projectList .project-row').waitFor();
    await fits(page, `projects ${width}`);
    const toggle = page.locator('#dashboardSidebarToggle');
    assert.equal(await toggle.getAttribute('aria-expanded'), 'false');
    assert.equal(await page.locator('.sidebar-nav').isVisible(), false);
    assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false);
    await checkMobileHeader(page);
    const preference = await page.evaluate(()=>localStorage.getItem('busbar.dashboard.sidebar.collapsed.v1'));
    await toggle.tap();
    await settleMenu(page);
    assert.equal(await page.locator('#logoutBtn').isVisible(), true);
    const sidebar = await page.locator('.dashboard-sidebar').boundingBox();
    assert.equal(sidebar.x, 0, 'expanded menu is on the left');
    assert.equal(sidebar.y, 64, 'expanded menu starts below the header');
    assert.equal(sidebar.height, 844 - 64);
    await checkMobileHeader(page);
    if (width === 390 && process.env.SAVE_MOBILE_SCREENSHOTS === '1'){
      await page.screenshot({ path: join(tmpdir(), 'busbar-mobile-sidebar-open.png') });
    }
    await page.keyboard.press('Escape');
    await settleMenu(page);
    assert.equal(await toggle.getAttribute('aria-expanded'), 'false');
    await page.locator('.project-detail-toggle-btn').tap();
    await fits(page, `project lines ${width}`);
    await page.locator('#newProjectBtn').tap();
    await page.locator('#projectNameInput').fill('Test mobil');
    await fits(page, `project form ${width}`);
    await page.locator('#projectFormModal .form-modal-close').tap();
    for (const view of ['dashboard', 'offers', 'leads', 'email', 'customers', 'company-card', 'busbar-folders', 'project-folders', 'supplier-folders']){
      await toggle.tap();
      await settleMenu(page);
      await page.locator(`[data-dashboard-nav="${view}"]`).tap();
      assert.equal(await page.locator('.sidebar-nav').isVisible(), true, 'page selection keeps the sidebar open');
      await checkMobileHeader(page);
      await toggle.tap();
      await settleMenu(page);
      assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false, 'minimized sidebar has no icon rail');
      await checkMobileHeader(page);
      await fits(page, `${view} ${width}`);
      if (view === 'leads'){
        const [button, form] = ['newLeadBtn', 'leadForm'];
        await page.locator(`#${button}`).tap();
        await page.locator(`#${form}Modal`).waitFor({ state: 'visible' });
        await fits(page, `${view} form ${width}`);
        const rect = await page.locator(`#${form}Modal .form-modal-panel`).boundingBox();
        assert.ok(rect.x >= 0 && rect.x + rect.width <= width, `${view} modal fits horizontally`);
        assert.ok(rect.y >= 0 && rect.y + rect.height <= 844, `${view} modal fits vertically`);
        await page.locator(`#${form}Modal .form-modal-close`).tap();
      }
    }
    if (width === 390){
      await toggle.tap();
      await page.locator('[data-dashboard-link="calculator"]').tap();
      await page.waitForURL(/calculator\.html/);
      await page.locator('#status').getByText(/CSV lastet/).waitFor();
      await checkMobileHeader(page);
      assert.equal(await page.locator('.dashboard-sidebar').isVisible(), true, 'sidebar state survives page navigation');
      await toggle.tap();
      await page.reload();
      await page.locator('#status').getByText(/CSV lastet/).waitFor();
      await checkMobileHeader(page);
      assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false, 'minimized state survives reload');
    }
    assert.equal(await page.evaluate(()=>localStorage.getItem('busbar.dashboard.sidebar.collapsed.v1')), preference);
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.waitForFunction(()=>!document.querySelector('.sidebar-nav').inert);
    assert.equal(await page.locator('.sidebar-nav').isVisible(), true);
    assert.equal(await page.locator('.sidebar-nav').evaluate(el=>el.inert), false);
    assert.equal(await toggle.getAttribute('aria-expanded'), 'true');
    assert.deepEqual(errors, []);
    await page.close();
  }
});

test('mobile login remains reachable on index, calculator and admin', async t=>{
  const { base, pageFor } = await setup(t);
  for (const file of ['index.html', 'calculator.html', 'admin.html']){
    const page = await pageFor(390, false);
    await page.goto(`${base}/${file}`);
    await page.waitForFunction(()=>!document.body.classList.contains('is-shell-booting'));
    await checkMobileHeader(page);
    await page.locator('#dashboardSidebarToggle').tap();
    assert.equal(await page.locator('#loginBtn').isVisible(), true, file);
    await page.locator('#loginBtn').tap();
    const form = page.locator(file === 'admin.html' ? '#adminLoginForm' : '#loginModal');
    assert.equal(await form.isVisible(), true, file);
    await fits(page, file);
    await page.close();
  }
});

test('mobile sidebar fades and uses the PC toggle icon and timing with either motion preference', async t=>{
  const { base, pageFor } = await setup(t);
  const page = await pageFor(390);
  await page.goto(`${base}/index.html?view=projects`);
  await page.locator('#projectList .project-row').waitFor();
  await page.evaluate(()=>document.fonts.ready);
  const toggle = page.locator('#dashboardSidebarToggle');
  await toggle.focus();
  assert.equal(await toggle.getAttribute('title'), 'Utvid meny');
  const closedIcon = await toggle.evaluate(el=>getComputedStyle(el.querySelector('.sidebar-collapse-icon'), '::before').transform);
  await page.keyboard.press('Enter');
  await settleMenu(page);
  assert.equal(await toggle.getAttribute('title'), 'Minimer meny');
  const openIcon = await toggle.evaluate(el=>getComputedStyle(el.querySelector('.sidebar-collapse-icon'), '::before').transform);
  assert.notEqual(openIcon, closedIcon, 'menu icon changes to a close icon');
  await page.keyboard.press('Space');
  await settleMenu(page);
  assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false);

  for (const reducedMotion of ['no-preference', 'reduce']){
   await page.emulateMedia({ reducedMotion });
   for (const opening of [true, false]){
    const samples = await page.evaluate(async()=>{
      const sidebar = document.querySelector('.dashboard-sidebar');
      const backdrop = document.querySelector('.mobile-menu-backdrop');
      const sample = ()=>({ opacity: Number(getComputedStyle(sidebar).opacity), backdrop: Number(getComputedStyle(backdrop).opacity),
        x: sidebar.getBoundingClientRect().x, visible: getComputedStyle(sidebar).visibility });
      const frames = [sample()];
      document.getElementById('dashboardSidebarToggle').click();
      const start = performance.now();
      do {
        await new Promise(requestAnimationFrame);
        frames.push(sample());
      } while (performance.now() - start < 350);
      return frames;
    });
    assert.ok(samples.some(frame=>frame.opacity > 0 && frame.opacity < 1), 'sidebar visibly fades');
    assert.ok(samples.some(frame=>frame.backdrop > 0 && frame.backdrop < 1), 'backdrop visibly fades');
    assert.ok(samples.some(frame=>frame.x > -240 && frame.x < 0), 'sidebar slides smoothly');
    assert.equal(samples.at(-1).opacity, opening ? 1 : 0);
    assert.equal(samples.at(-1).visible, opening ? 'visible' : 'hidden');
   }
  }
  await toggle.evaluate(button=>{ button.click(); button.click(); button.click(); });
  await settleMenu(page);
  assert.equal(await toggle.getAttribute('aria-expanded'), 'true');
  await page.locator('.mobile-menu-backdrop').click({ position: { x: 300, y: 100 } });
  await settleMenu(page);
  assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false);
  assert.equal(await page.locator('.sidebar-nav').evaluate(el=>el.inert), true);
  await checkMobileHeader(page);
  await toggle.click();
  await settleMenu(page);
  const mobileIcon = await toggle.evaluate(el=>{
    const style = getComputedStyle(el.querySelector('.sidebar-collapse-icon'), '::before');
    return { transform: style.transform, duration: style.transitionDuration, easing: style.transitionTimingFunction };
  });
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.waitForFunction(()=>document.getElementById('dashboardSidebarToggle').closest('.sidebar-top-row'));
  await page.waitForFunction(expected=>getComputedStyle(document.querySelector('.sidebar-collapse-icon'), '::before').transform === expected, mobileIcon.transform);
  assert.deepEqual(await toggle.evaluate(el=>{
    const style = getComputedStyle(el.querySelector('.sidebar-collapse-icon'), '::before');
    return { transform: style.transform, duration: style.transitionDuration, easing: style.transitionTimingFunction };
  }), mobileIcon, 'mobile and PC share the same expanded icon and animation');
  await toggle.click();
  await page.waitForFunction(expected=>getComputedStyle(document.querySelector('.sidebar-collapse-icon'), '::before').transform === expected, closedIcon);
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForFunction(()=>document.getElementById('dashboardSidebarToggle').closest('.mobile-app-header'));
  await toggle.click();
  await settleMenu(page);
  assert.equal(await page.locator('.dashboard-sidebar').isVisible(), false);
});

test('mobile calendar uses agenda with working period navigation and editing', async t=>{
  const { base, pageFor, calendarRequests } = await setup(t);
  const page = await pageFor(390);
  await page.goto(`${base}/index.html?view=calendar`);
  await page.locator('.calendar-event-item').waitFor();
  assert.equal(await page.locator('#calendarModeSelect').inputValue(), 'list');
  await fits(page, 'calendar agenda');
  const first = calendarRequests.at(-1);
  await page.locator('#calendarNextBtn').tap();
  await page.waitForFunction(()=>document.querySelector('#calendarPeriodLabel').textContent !== 'Kommende 14 dager');
  await page.locator('#calendarStatus').filter({ hasText: 'Oppdatert' }).waitFor();
  assert.equal(new Date(calendarRequests.at(-1)) - new Date(first), 14 * 86400000);
  await page.locator('[data-edit-calendar-event="event-1"]').tap();
  await page.locator('#calendarEventFormModal').waitFor({ state: 'visible' });
  await fits(page, 'calendar edit');
  const modal = await page.locator('#calendarEventFormModal .form-modal-panel').boundingBox();
  assert.ok(modal.x >= 0 && modal.x + modal.width <= 390);
  await page.locator('#calendarEventFormModal .form-modal-close').tap();
  await page.locator('#calendarModeSelect').selectOption('month');
  await page.locator('#calendarGridView').waitFor({ state: 'visible' });
  await fits(page, 'calendar month scroll container');
});

test('mobile flow edits dates and completion in shared tasks, with optional touch-safe timeline', async t=>{
  const { base, pageFor, projects } = await setup(t);
  const page = await pageFor(390);
  await page.goto(`${base}/index.html?view=project-flow`);
  await page.locator('.project-flow-list-task').waitFor();
  await fits(page, 'flow list');
  await page.locator('[data-project-flow-edit="task-1"]').tap();
  await page.locator('#projectFlowMilestoneFormModal').waitFor({ state: 'visible' });
  await fits(page, 'flow edit form');
  await page.locator('#projectFlowEndDateInput').fill('10/11/2026');
  await page.locator('#projectFlowEndDateInput').dispatchEvent('change');
  await page.locator('#projectFlowMilestoneForm button[type="submit"]').tap();
  await page.locator('#projectFlowMilestoneFormModal').waitFor({ state: 'hidden' });
  await page.waitForFunction(()=>document.querySelector('.project-flow-list-task').textContent.includes('10/11/2026'));
  await page.locator('[data-project-flow-toggle="task-1"]').check();
  await page.waitForFunction(()=>document.querySelector('.project-flow-list-task').classList.contains('is-completed'));
  await page.waitForTimeout(800); // project sync is debounced
  assert.equal(projects()[0].projectFlowMilestones[0].completed, true);
  assert.equal(projects()[0].projectFlowMilestones[0].endDate, '2026-11-10');
  await page.locator('#projectFlowTimelineBtn').tap();
  await page.locator('.project-flow-scroller').waitFor();
  await fits(page, 'flow timeline');
  const bar = page.locator('[data-project-flow-drag="task-1"]');
  await bar.dispatchEvent('pointerdown', { pointerType: 'touch', pointerId: 1, clientX: 100, clientY: 100 });
  assert.equal(await page.locator('.project-flow-bar.is-dragging').count(), 0);
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.locator('.project-flow-list').waitFor({ state: 'hidden' });
  assert.equal(await page.locator('.project-flow-scroller').isVisible(), true);
});

test('mobile calculator calculates, downloads BOM and saves a line', async t=>{
  const { base, pageFor, projects } = await setup(t);
  const page = await pageFor(360);
  await page.goto(`${base}/calculator.html?project=mobile-project&newLine=1`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor();
  await fits(page, 'calculator form');
  await page.locator('.mobile-calculator-section > summary').first().tap();
  await fits(page, 'calculator technical section');
  await page.locator('#series').selectOption('XAP-B');
  await page.locator('#ampSelect').selectOption('500');
  await page.locator('#meter').fill('1');
  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('end_cover');
  await page.locator('#lineNumberInput').fill('Mobil-2');
  await page.locator('#calcBtn').tap();
  await page.locator('#status').getByText('OK', { exact: true }).waitFor();
  await fits(page, 'calculator results');
  const resultLayout = await page.evaluate(()=>{
    const groups = [...document.querySelectorAll('.totals-selectors > *')].map(el=>el.getBoundingClientRect());
    const check = document.getElementById('includeMontasje').getBoundingClientRect();
    return { noOverlap: groups.every((rect, index)=>!index || rect.top >= groups[index - 1].bottom), checkWidth: check.width };
  });
  assert.ok(resultLayout.noOverlap, 'offer options and totals must not overlap');
  assert.ok(resultLayout.checkWidth >= 22, 'offer checkboxes must be visible and tappable');
  await page.locator('#includeMontasje').uncheck();
  await page.locator('#includeMontasje').check();
  assert.ok(await page.locator('#bomTbl td[data-label]').count());
  const [download] = await Promise.all([page.waitForEvent('download'), page.locator('#exportCsv').tap()]);
  assert.equal(download.suggestedFilename(), 'BOM.csv');
  assert.equal(await download.failure(), null);
  await page.locator('#saveLineBtn').tap();
  await page.waitForURL(/index\.html/);
  await page.locator('.project-row').waitFor();
  assert.ok(projects()[0].lines.some(line=>line.lineNumber === 'Mobil-2'));
  await fits(page, 'saved project');
});

test('desktop geometry matches the committed app and collapsed sidebar survives mobile use', async t=>{
  const { base, pageFor } = await setup(t);
  const baselineApp = execFileSync('git', ['show', 'HEAD:app.js'], { cwd: root });
  const selectors = {
    'index.html?view=dashboard': ['.dashboard-sidebar', '.main-dashboard-grid', '.dashboard-total-widget', '.dashboard-actions-widget'],
    'index.html?view=projects': ['.dashboard-sidebar', '.project-row', '.project-body-grid', '.project-action-buttons'],
    'calculator.html': ['.dashboard-sidebar', '#calculatorView', '.row', '.montasje-box', '.row-actions'],
    'admin.html': ['.dashboard-sidebar', '#adminLoginForm']
  };
  async function geometry(page, path){
    await page.goto(`${base}/${path}`);
    await page.waitForFunction(()=>!document.body.classList.contains('is-shell-booting'));
    if (path.includes('projects')) await page.locator('.project-row').waitFor();
    if (path === 'calculator.html') await page.locator('#status').getByText(/CSV lastet/).waitFor();
    await page.evaluate(()=>document.fonts.ready);
    return page.evaluate(selectors=>selectors.map(selector=>{
      const el = document.querySelector(selector);
      const r = el.getBoundingClientRect();
      return { selector, x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height) };
    }), selectors[path]);
  }
  for (const path of Object.keys(selectors)){
    const current = await pageFor(1440);
    const baseline = await pageFor(1440);
    const file = path.split('?')[0];
    const baselineHtml = execFileSync('git', ['show', `HEAD:${file}`], { cwd: root });
    await baseline.route(`**/${file}*`, route=>route.fulfill({ contentType: 'text/html', body: baselineHtml }));
    if (file !== 'admin.html') await baseline.route('**/app.js', route=>route.fulfill({ contentType: 'text/javascript', body: baselineApp }));
    assert.deepEqual(await geometry(current, path), await geometry(baseline, path), path);
    await current.close();
    await baseline.close();
  }
  const page = await pageFor(1440);
  await page.goto(`${base}/index.html?view=projects`);
  await page.locator('#projectList .project-row').waitFor();
  await page.locator('#dashboardSidebarToggle').click();
  assert.equal(await page.evaluate(()=>localStorage.getItem('busbar.dashboard.sidebar.collapsed.v1')), '1');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForFunction(()=>document.querySelector('#dashboardSidebarToggle').getAttribute('aria-label') === 'Utvid meny'
    && document.querySelector('#dashboardSidebarToggle').closest('.mobile-app-header'));
  await page.locator('#dashboardSidebarToggle').click();
  assert.equal(await page.locator('#logoutBtn').isVisible(), true);
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.waitForFunction(()=>document.querySelector('#dashboardSidebarToggle').getAttribute('aria-label') === 'Utvid meny');
  assert.equal(await page.locator('#dashboardSidebarToggle').getAttribute('aria-expanded'), 'false');
  assert.equal(await page.locator('.sidebar-nav').evaluate(el=>el.inert), false);
  assert.equal(await page.evaluate(()=>localStorage.getItem('busbar.dashboard.sidebar.collapsed.v1')), '1');
});
