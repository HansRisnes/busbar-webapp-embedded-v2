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

function token(email, secret){
  const now = Math.floor(Date.now() / 1000);
  const payload = Buffer.from(JSON.stringify({ email, iat: now, exp: now + 3600 })).toString('base64url');
  const signature = createHmac('sha256', secret).update(payload).digest('base64url');
  return `${payload}.${signature}`;
}

async function availablePort(){
  const server = createServer();
  await new Promise(resolveListening=>server.listen(0, '127.0.0.1', resolveListening));
  const port = server.address().port;
  await new Promise(resolveClose=>server.close(resolveClose));
  return port;
}

test('manual aluminium price is persisted and owner-only', async t=>{
  const dataDir = await mkdtemp(join(tmpdir(), 'busbar-market-test-'));
  const secret = 'market-price-test-secret';
  const port = await availablePort();
  const fxUpdatedAt = new Date().toISOString();
  await writeFile(join(dataDir, 'market-data.json'), JSON.stringify({
    aluminium: { price: 3035, currency: 'USD', source: 'Manuell', unit: 't' },
    fx: { usdNok: { rate: 10, date: '2026-10-28' }, eurNok: { rate: 11, date: '2026-10-28' } },
    updatedAt: fxUpdatedAt
  }));
  await writeFile(join(dataDir, 'user-auth.json'), JSON.stringify({ users: {
    'owner@busbar.no': { email: 'owner@busbar.no', appRole: 'administrator', profile: { name: 'Owner' } },
    'user@busbar.no': { email: 'user@busbar.no', appRole: 'standard', profile: { name: 'User' } }
  } }));

  const child = spawn(process.execPath, ['server/mail.js'], {
    cwd: root,
    env: {
      ...process.env,
      DATA_DIR: dataDir,
      MARKET_DATA_FILE: join(dataDir, 'market-data.json'),
      PORT: String(port),
      AUTH_TOKEN_SECRET: secret,
      ADMIN_OWNER_EMAILS: '',
      MARKET_TIMEZONE: 'Europe/Oslo'
    },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  let log = '';
  child.stdout.on('data', chunk=>{ log += chunk.toString(); });
  child.stderr.on('data', chunk=>{ log += chunk.toString(); });
  t.after(async ()=>{
    child.kill();
    await rm(dataDir, { recursive: true, force: true });
  });

  const url = `http://127.0.0.1:${port}/api/market-data`;
  let ready = false;
  for (let i = 0; i < 60; i++){
    try{
      const response = await fetch(url);
      if (response.ok){ ready = true; break; }
    }catch{}
    await delay(250);
  }
  assert.ok(ready, `Market API did not start: ${log}`);
  const schedule = (await (await fetch(url)).json()).schedule;
  const nextRunLocal = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Europe/Oslo', hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
  }).format(new Date(schedule.nextRunAt));
  assert.equal(nextRunLocal, '06:00');

  const put = (authorization, price)=>fetch(`${url}/aluminium`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json', ...(authorization ? { Authorization: authorization } : {}) },
    body: JSON.stringify({ price })
  });
  assert.equal((await put('', 3200)).status, 401);
  assert.equal((await put(`Bearer ${token('user@busbar.no', secret)}`, 3200)).status, 403);
  assert.equal((await put(`Bearer ${token('owner@busbar.no', secret)}`, -1)).status, 400);

  const response = await put(`Bearer ${token('owner@busbar.no', secret)}`, 3200);
  assert.equal(response.status, 200, log);
  const result = await response.json();
  assert.equal(result.aluminium.price, 3200);
  assert.equal(result.aluminium.source, 'Manuell');
  assert.equal(result.updatedAt, fxUpdatedAt);
  assert.ok(Date.parse(result.aluminium.updatedAt) > Date.parse(fxUpdatedAt));
  const persisted = JSON.parse(await readFile(join(dataDir, 'market-data.json'), 'utf8'));
  assert.equal(persisted.aluminium.price, 3200);
  assert.equal(persisted.aluminium.updatedAt, result.aluminium.updatedAt);
});
