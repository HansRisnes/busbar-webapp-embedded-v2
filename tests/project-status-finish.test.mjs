import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { spawn } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer as createHttpServer } from 'node:http';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('project requires a completed Fakturert task before becoming finished', async t=>{
  const dataDir = await mkdtemp(join(tmpdir(), 'busbar-finish-test-'));
  const secret = 'finish-test-secret';
  const email = 'seller@busbar.no';
  await writeFile(join(dataDir, 'user-auth.json'), JSON.stringify({
    users: { [email]: { email, appRole: 'standard', profile: { name: 'Seller' } } }
  }));
  const portServer = createServer();
  await new Promise(resolveListening=>portServer.listen(0, '127.0.0.1', resolveListening));
  const port = portServer.address().port;
  await new Promise(resolveClose=>portServer.close(resolveClose));

  const child = spawn(process.execPath, ['server/mail.js'], {
    cwd: root,
    env: { ...process.env, DATA_DIR: dataDir, PORT: String(port), AUTH_TOKEN_SECRET: secret },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  let log = '';
  child.stdout.on('data', chunk=>{ log += chunk.toString(); });
  child.stderr.on('data', chunk=>{ log += chunk.toString(); });
  t.after(async()=>{
    child.kill();
    await rm(dataDir, { recursive: true, force: true });
  });

  const base = `http://127.0.0.1:${port}`;
  let ready = false;
  for (let attempt = 0; attempt < 120; attempt++){
    if (child.exitCode !== null) break;
    try{
      const response = await fetch(`${base}/api/market-data`);
      if (response.ok){ ready = true; break; }
    }catch{}
    await delay(250);
  }
  assert.ok(ready, `Project API did not start: ${log}`);

  const now = Math.floor(Date.now() / 1000);
  const payload = Buffer.from(JSON.stringify({ email, iat: now, exp: now + 3600 })).toString('base64url');
  const token = `${payload}.${createHmac('sha256', secret).update(payload).digest('base64url')}`;
  const sync = async project=>{
    const response = await fetch(`${base}/api/user-projects/sync`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, projects: [project] })
    });
    return { status: response.status, data: await response.json() };
  };
  const project = {
    id: 'finish-test-project', name: 'Test', projectOwnerEmail: email,
    projectStatus: 'unresolved', createdAt: new Date().toISOString(), lines: [],
    projectFlowMilestones: []
  };
  assert.equal((await sync(project)).status, 200, log);
  project.projectStatus = 'finished';
  const withoutInvoice = await sync(project);
  assert.equal(withoutInvoice.status, 409, log);
  assert.match(withoutInvoice.data.error, /Fakturert/);

  project.projectFlowMilestones = [{
    id: 'invoice', phaseId: 'finished', startDate: '2026-11-05', endDate: '2026-11-05', completed: false
  }];
  assert.equal((await sync(project)).status, 409, log);
  project.projectFlowMilestones[0].completed = true;
  project.projectFlowMilestones.push({
    id: 'second-invoice', phaseId: 'finished', startDate: '2026-11-06', endDate: '2026-11-06', completed: false
  });
  assert.equal((await sync(project)).status, 409, log);
  project.projectFlowMilestones[1].completed = true;
  assert.equal((await sync(project)).status, 200, log);

  project.projectFlowMilestones = [];
  assert.equal((await sync(project)).status, 200, 'Already finished projects must remain syncable');
});

test('project status dialog unlocks Ferdig only after Fakturert is checked', async t=>{
  const email = 'status-test@busbar.no';
  const project = (id, completed)=>({
    id, name: id, customer: 'Test', projectOwnerEmail: email,
    projectStatus: 'unresolved', createdAt: '2026-11-05T00:00:00.000Z', lines: [],
    projectFlowMilestones: completed === null ? [] : [{
      id: `${id}-invoice`, phaseId: 'finished', startDate: '2026-11-05', endDate: '2026-11-05', completed
    }]
  });
  let projects = [project('without-invoice', null), project('invoiced', true)];
  const server = createHttpServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      if (pathname === '/api/user-projects/sync' && request.method === 'POST'){
        let body = '';
        for await (const chunk of request) body += chunk;
        projects = JSON.parse(body).projects;
      }
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects, ownerEmails: [email], updatedAt: new Date().toISOString() }
        : { customers: [] }));
      return;
    }
    const file = resolve(root, `.${pathname === '/' ? '/index.html' : pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.writeHead(403).end();
      return;
    }
    try{
      response.setHeader('Content-Type', {
        '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml'
      }[extname(file)] || 'application/octet-stream');
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
  await page.addInitScript(userEmail=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true, username: userEmail, token: 'test-token'
    }));
  }, email);
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=projects`);
  await page.locator('[data-project-status-edit="without-invoice"]').click();
  assert.equal(await page.locator('[data-project-status-option="finished"]').isDisabled(), true);
  assert.equal(await page.locator('#projectStatusRequirement').count(), 0);
  await page.locator('#projectStatusCancel').click();
  await page.locator('[data-project-status-edit="invoiced"]').click();
  assert.equal(await page.locator('[data-project-status-option="finished"]').isEnabled(), true);
  await page.locator('[data-project-status-option="finished"]').click();
  assert.equal(projects.find(item=>item.id === 'invoiced').projectStatus, 'finished');
});
