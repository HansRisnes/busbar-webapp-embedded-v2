import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function token(email, secret) {
  const now = Math.floor(Date.now() / 1000);
  const payload = Buffer.from(JSON.stringify({ email, iat: now, exp: now + 3600 })).toString('base64url');
  const signature = createHmac('sha256', secret).update(payload).digest('base64url');
  return `${payload}.${signature}`;
}

async function availablePort() {
  const server = createServer();
  await new Promise(resolveListening => server.listen(0, '127.0.0.1', resolveListening));
  const port = server.address().port;
  await new Promise(resolveClose => server.close(resolveClose));
  return port;
}

test('leads persist and respect administrator, insight and standard access', async t => {
  const dataDir = await mkdtemp(join(tmpdir(), 'busbar-leads-test-'));
  const secret = 'leads-test-secret';
  const port = await availablePort();
  const users = {
    'owner@busbar.no': { email: 'owner@busbar.no', appRole: 'administrator', profile: { name: 'Owner' } },
    'owner@mcselektrotavler.no': { email: 'owner@mcselektrotavler.no', appRole: 'administrator', profile: { name: 'Owner' } },
    'owner2@busbar.no': { email: 'owner2@busbar.no', appRole: 'administrator', profile: { name: 'Owner 2' } },
    'seller@busbar.no': { email: 'seller@busbar.no', appRole: 'standard', profile: { name: 'Selger' } },
    'viewer@busbar.no': { email: 'viewer@busbar.no', appRole: 'insight', profile: { name: 'Innsyn' } }
  };
  await writeFile(join(dataDir, 'user-auth.json'), JSON.stringify({ users }));

  const child = spawn(process.execPath, ['server/mail.js'], {
    cwd: root,
    env: {
      ...process.env,
      DATA_DIR: dataDir,
      PORT: String(port),
      AUTH_TOKEN_SECRET: secret,
      ADMIN_OWNER_EMAILS: '',
      MICROSOFT_AUTH_CLIENT_SECRET: ' ',
      MS_AUTH_CLIENT_SECRET: ' '
    },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  let log = '';
  child.stdout.on('data', chunk => { log += chunk.toString(); });
  child.stderr.on('data', chunk => { log += chunk.toString(); });
  t.after(async () => {
    child.kill();
    await rm(dataDir, { recursive: true, force: true });
  });

  const base = `http://127.0.0.1:${port}`;
  let ready = false;
  for (let i = 0; i < 120; i++) {
    if (child.exitCode !== null) break;
    try {
      const response = await fetch(`${base}/api/market-data`);
      if (response.ok) { ready = true; break; }
    } catch {}
    await delay(250);
  }
  assert.ok(ready, `Leads API did not start: ${log}`);

  const api = async (email, path, body) => {
    const response = await fetch(`${base}${path}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: {
        Authorization: `Bearer ${token(email, secret)}`,
        ...(body === undefined ? {} : { 'Content-Type': 'application/json' })
      },
      ...(body === undefined ? {} : { body: JSON.stringify(body) })
    });
    return { status: response.status, data: await response.json() };
  };
  const input = {
    companyName: 'Testbedrift AS',
    contactName: 'Kari Kontakt',
    contactEmail: 'kari@example.no',
    contactPhone: '12345678',
    projectNeed: 'Strømskinner til nytt bygg',
    salesPhase: 'new',
    nextActivity: 'Oppfølging',
    followUpDate: '2026-12-01',
    estimatedValue: 450000,
    probability: 30
  };

  assert.equal((await fetch(`${base}/api/leads`)).status, 401);
  assert.equal((await api('seller@busbar.no', '/api/leads/upsert', { ...input, followUpDate: '2026-02-30' })).status, 400);
  assert.equal((await api('seller@busbar.no', '/api/leads/upsert', { ...input, nextActivity: '' })).status, 400);
  assert.equal((await api('seller@busbar.no', '/api/leads/upsert', { ...input, contactPhone: '' })).status, 400);
  assert.equal((await api('seller@busbar.no', '/api/leads/upsert', {
    ...input, responsibleEmail: 'owner@busbar.no'
  })).status, 403);

  const created = await api('seller@busbar.no', '/api/leads/upsert', input);
  assert.equal(created.status, 200, log);
  assert.equal(created.data.leads.length, 1);
  const lead = created.data.leads[0];
  assert.equal(lead.responsibleEmail, 'seller@busbar.no');
  assert.equal(lead.createdByEmail, 'seller@busbar.no');
  assert.equal(lead.contactName, 'Kari Kontakt');
  assert.equal(lead.estimatedValue, 450000);
  assert.equal(lead.canEdit, true);
  assert.equal(created.data.canAssignOthers, false);
  const customers = JSON.parse(await readFile(join(dataDir, 'customer-database.json'), 'utf8')).customers;
  assert.equal(customers.find(item => item.name === 'Testbedrift AS').contacts[0].email, 'kari@example.no');
  assert.ok(created.data.customers.some(item => item.name === 'Testbedrift AS'));
  assert.match(lead.calendarSyncError, /MICROSOFT_AUTH_CLIENT_SECRET/);

  const ownerList = await api('owner@busbar.no', '/api/leads');
  assert.equal(ownerList.status, 200);
  assert.equal(ownerList.data.leads.length, 1);
  assert.equal(ownerList.data.canAssignOthers, true);
  assert.equal(ownerList.data.leads[0].responsibleName, 'Selger');
  assert.ok(ownerList.data.users.some(user => user.email === 'owner2@busbar.no'));
  assert.ok(ownerList.data.users.some(user => user.email === 'lars@busbar.no'));
  assert.equal(ownerList.data.users.filter(user => user.name === 'Owner').length, 1);
  assert.ok(ownerList.data.users.some(user => user.email === 'owner@busbar.no'));
  assert.equal(created.data.users.length, 1);

  const insightList = await api('viewer@busbar.no', '/api/leads');
  assert.equal(insightList.data.leads.length, 1);
  assert.equal(insightList.data.leads[0].canEdit, false);
  assert.equal(insightList.data.canAssignOthers, false);
  assert.equal((await api('viewer@busbar.no', '/api/leads/delete', {
    id: lead.id, updatedAt: lead.updatedAt
  })).status, 403);
  assert.equal((await api('viewer@busbar.no', '/api/leads/upsert', {
    ...lead, nextActivity: 'Ikke tillatt', updatedAt: lead.updatedAt
  })).status, 403);
  assert.equal((await api('viewer@busbar.no', '/api/leads/complete', {
    id: lead.id, updatedAt: lead.updatedAt, completed: true
  })).status, 403);

  const completed = await api('seller@busbar.no', '/api/leads/complete', {
    id: lead.id, updatedAt: lead.updatedAt, completed: true
  });
  assert.equal(completed.status, 200);
  assert.equal(completed.data.leads[0].followUpCompleted, true);
  const activeLead = completed.data.leads[0];

  const ownerCreated = await api('owner@busbar.no', '/api/leads/upsert', {
    ...input, companyName: 'Kun for eier'
  });
  assert.equal(ownerCreated.status, 200);
  assert.equal((await api('seller@busbar.no', '/api/leads')).data.leads.length, 1);
  assert.equal((await api('viewer@busbar.no', '/api/leads')).data.leads.length, 2);

  const reassigned = await api('owner@busbar.no', '/api/leads/upsert', {
    ...activeLead, responsibleEmail: 'owner2@busbar.no', updatedAt: activeLead.updatedAt
  });
  assert.equal(reassigned.status, 200);
  assert.equal(reassigned.data.leads.find(item => item.id === lead.id).responsibleEmail, 'owner2@busbar.no');
  assert.equal((await api('owner2@busbar.no', '/api/leads')).data.leads.length, 2);
  assert.notEqual(reassigned.data.leads.find(item => item.id === lead.id).updatedAt, activeLead.updatedAt);
  assert.equal((await api('seller@busbar.no', '/api/leads')).data.leads.length, 1);
  assert.equal((await api('owner@busbar.no', '/api/leads/upsert', {
    ...lead, nextActivity: 'Utdatert endring', updatedAt: lead.updatedAt
  })).status, 409);

  const persisted = JSON.parse(await readFile(join(dataDir, 'leads.json'), 'utf8'));
  assert.equal(persisted.leads.length, 2);
  assert.equal(persisted.leads.find(item => item.id === lead.id).responsibleEmail, 'owner2@busbar.no');
  const deleted = await api('owner@busbar.no', '/api/leads/delete', {
    id: lead.id, updatedAt: reassigned.data.leads.find(item => item.id === lead.id).updatedAt
  });
  assert.equal(deleted.status, 200);
  assert.equal(deleted.data.leads.length, 1);
});
