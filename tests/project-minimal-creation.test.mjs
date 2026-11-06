import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('name-only projects support lines and the first direct calculation creates and saves one', async t=>{
  let projects = [];
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      if (pathname === '/api/user-projects/sync' && request.method === 'POST'){
        let body = '';
        for await (const chunk of request) body += chunk;
        const byId = new Map(projects.map(project=>[project.id, project]));
        const incomingProjects = JSON.parse(body).projects;
        for (const project of incomingProjects){
          const existing = byId.get(project.id);
          if (!existing || Date.parse(project.updatedAt) > Date.parse(existing.updatedAt)
            || (project.updatedAt === existing.updatedAt && project.lines.length >= existing.lines.length)){
            byId.set(project.id, project);
          }
        }
        projects = Array.from(byId.values());
      }
      if (pathname === '/api/market-data'){
        response.end(JSON.stringify({
          aluminium: { price: 2800, currency: 'USD', symbol: 'MANUAL', source: 'Manuell', unit: 't', updatedAt: new Date().toISOString() },
          fx: { usdNok: { rate: 10, date: '2026-11-05' } },
          updatedAt: new Date(Date.now() - 60_000).toISOString()
        }));
      } else {
        response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
          ? { projects, ownerEmails: ['project-test@busbar.no'], updatedAt: new Date().toISOString() }
          : { customers: [] }));
      }
      return;
    }
    const file = resolve(root, `.${pathname === '/' ? '/index.html' : pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.writeHead(403).end();
      return;
    }
    try{
      const content = await readFile(file);
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.csv': 'text/csv', '.svg': 'image/svg+xml', '.ttf': 'font/ttf' };
      response.setHeader('Content-Type', `${mime[extname(file)] || 'application/octet-stream'}; charset=utf-8`);
      response.end(content);
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
      loggedIn: true, username: 'project-test@busbar.no', token: 'test-token'
    }));
  });
  const base = `http://127.0.0.1:${server.address().port}`;
  await page.goto(`${base}/index.html?view=projects`);
  await page.locator('#newProjectBtn').click();
  assert.equal(await page.locator('#projectSuggestions').count(), 0);
  assert.equal(await page.locator('#projectSubmit').isDisabled(), true);
  await page.locator('#projectNameInput').fill('Prosjekt uten kunde');
  assert.equal(await page.locator('#projectSubmit').isEnabled(), true);
  await page.locator('#projectNameInput').press('Enter');
  await page.locator('.project-row', { hasText: 'Prosjekt uten kunde' }).waitFor();
  assert.equal(await page.locator('#projectDetailsModal').isVisible(), false);
  assert.equal(projects.length, 1);
  assert.equal(projects[0].customer, '');
  assert.equal(projects[0].contactPerson, '');

  await page.goto(`${base}/calculator.html?project=${encodeURIComponent(projects[0].id)}&newLine=1`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  await page.waitForFunction(()=>{
    const meta = document.querySelector('#projectMeta');
    return meta && !meta.hidden && meta.querySelector('[data-project-name]')?.textContent === 'Prosjekt uten kunde';
  });
  await page.locator('#series').selectOption('XAP-B');
  await page.locator('#ampSelect').selectOption('500');
  await page.locator('#meter').fill('1');
  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('end_cover');
  await page.locator('#lineNumberInput').fill('L-1');
  await page.locator('#calcBtn').click();
  await page.locator('#status').getByText('OK', { exact: true }).waitFor();
  await page.locator('#saveLineBtn').click();
  await page.waitForURL(/index\.html/);
  await page.locator('.project-row', { hasText: 'Prosjekt uten kunde' }).getByText('Linjer: 1').waitFor();
  assert.equal(projects[0].lines.length, 1);
  assert.equal(projects[0].lines[0].lineNumber, 'L-1');

  page.on('dialog', dialog=>dialog.dismiss());
  const generateOffer = page.locator(`[data-project-generate-offer="${projects[0].id}"]`);
  await generateOffer.evaluate(button=>{ button.disabled = false; });
  await generateOffer.click();
  await page.locator('#offerDetailsWarningModal').waitFor({ state: 'visible' });
  assert.match(await page.locator('#offerDetailsMissingList').textContent(), /Kunde/);
  assert.equal(await page.getByText('Generer likevel').count(), 0);
  await page.locator('#offerDetailsContinue').click();
  await page.locator('#projectDetailsModal').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#projectDetailsCustomer').isEnabled(), true);
  assert.equal(await page.locator('#projectDetailsContact').isEnabled(), true);
  await page.locator('#projectDetailsCustomer').fill('Ny kunde AS');
  await page.locator('#projectDetailsContact').fill('Kari Kontakt');
  await page.locator('#projectDetailsAddress').fill('Testveien 1');
  await page.locator('#projectDetailsPostalPlace').fill('0123 Oslo');
  await page.locator('#projectDetailsPhone').fill('12345678');
  await page.locator('#projectDetailsSubmit').click();
  await page.locator('.project-row', { hasText: 'Prosjekt uten kunde' }).getByText('Kunde: Ny kunde AS').waitFor();
  assert.equal(projects[0].contactPerson, 'Kari Kontakt');

  await page.goto(`${base}/calculator.html?project=${encodeURIComponent(projects[0].id)}`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  await page.locator('#activeCustomerName').getByText('Ny kunde AS').waitFor();
  assert.equal(new URL(page.url()).searchParams.get('project'), projects[0].id);
  let releaseAppScript;
  const appScriptGate = new Promise(resolveGate=>{ releaseAppScript = resolveGate; });
  await page.route('**/app.js', async route=>{
    await appScriptGate;
    await route.continue();
  });
  try{
    await page.reload({ waitUntil: 'commit' });
    await page.locator('#activeCustomerName').getByText('Ny kunde AS').waitFor();
    assert.match(await page.locator('#activeProjectName').textContent(), /Prosjekt uten kunde/);
  }finally{
    releaseAppScript();
  }
  await page.waitForLoadState('load');
  await page.unroute('**/app.js');
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  assert.equal(await page.locator('#activeCustomerName').textContent(), 'Ny kunde AS');
  assert.equal(await page.locator('#activeProjectBar').evaluate(el=>el.closest('.card-header')?.classList.contains('calculator-card-header')), true);
  assert.equal(await page.locator('#activeProjectBar button').count(), 0);
  assert.equal(await page.getByText('Åpne prosjektoversikt').count(), 0);
  await page.setViewportSize({ width: 390, height: 844 });
  const headerBounds = await page.locator('.calculator-card-header').evaluate(el=>{
    const header = el.getBoundingClientRect();
    const project = el.querySelector('#activeProjectBar').getBoundingClientRect();
    return { headerRight: header.right, projectRight: project.right, projectLeft: project.left };
  });
  assert.ok(headerBounds.projectRight <= headerBounds.headerRight + 1);
  assert.ok(headerBounds.projectLeft >= 0);
  await page.setViewportSize({ width: 1280, height: 720 });

  await page.goto(`${base}/calculator.html`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  await page.locator('#activeProjectBar').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#activeProjectName').textContent(), 'Ingen prosjekt valgt');
  await page.locator('#series').selectOption('XAP-B');
  await page.locator('#ampSelect').selectOption('500');
  await page.locator('#meter').fill('1');
  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('end_cover');
  await page.locator('#lineNumberInput').fill('DIREKTE-1');
  await page.locator('#calcBtn').click();
  await page.locator('#projectForm').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#projectSuggestions').count(), 0);
  assert.equal(await page.locator('#projectSubmit').textContent(), 'Opprett prosjekt og beregn');
  assert.equal(projects.length, 1);
  await page.locator('#customerNameInput').fill('Ny');
  await page.locator('#customerSuggestions li', { hasText: 'Ny kunde AS' }).waitFor();
  await page.locator('#customerNameInput').fill('');
  await page.locator('#projectNameInput').fill('Direkte prosjekt');
  await page.locator('#projectSubmit').click();
  await page.locator('#status').getByText('Linje DIREKTE-1 lagret.', { exact: true }).waitFor();
  const syncDeadline = Date.now() + 5000;
  while (!projects.some(project=>project.name === 'Direkte prosjekt' && project.lines?.length === 1)
    && Date.now() < syncDeadline){
    await new Promise(resolveDelay=>setTimeout(resolveDelay, 50));
  }
  assert.equal(new URL(page.url()).pathname, '/calculator.html');
  assert.match(await page.locator('#activeProjectName').textContent(), /Direkte prosjekt/);
  assert.equal(projects.length, 2);
  const directProject = projects.find(project=>project.name === 'Direkte prosjekt');
  assert.equal(directProject.customer, '');
  assert.equal(directProject.lines.length, 1);
  assert.equal(directProject.lines[0].lineNumber, 'DIREKTE-1');

  await page.locator('#lineNumberInput').fill('DIREKTE-2');
  await page.locator('#calcBtn').click();
  await page.locator('#status').getByText('OK', { exact: true }).waitFor();
  assert.equal(projects.find(project=>project.name === 'Direkte prosjekt').lines.length, 1,
    'Later calculations must not be saved automatically');
});
