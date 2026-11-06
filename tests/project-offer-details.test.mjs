import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('offer generation requires complete customer and contact details', async t=>{
  const dataDir = await mkdtemp(join(tmpdir(), 'busbar-offer-details-test-'));
  const secret = 'offer-details-test-secret';
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
  assert.ok(ready, `Offer API did not start: ${log}`);

  const now = Math.floor(Date.now() / 1000);
  const payload = Buffer.from(JSON.stringify({ email, iat: now, exp: now + 3600 })).toString('base64url');
  const token = `${payload}.${createHmac('sha256', secret).update(payload).digest('base64url')}`;
  const response = await fetch(`${base}/api/generate-offer`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ project: { id: 'minimal', name: 'Navn alene', customer: '', contactPerson: '' } })
  });
  assert.equal(response.status, 400, log);
  const data = await response.json();
  assert.match(data.error, /Kunde/);
  assert.match(data.error, /Kontaktperson/);
  assert.match(data.error, /Telefon/);
});
