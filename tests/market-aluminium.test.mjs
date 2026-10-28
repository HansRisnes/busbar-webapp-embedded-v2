import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('manual aluminium price turns white until the next FX update', async t=>{
  let fxUpdatedAt = new Date(Date.now() - 60_000).toISOString();
  let aluminium = { price: 3035, currency: 'USD', source: 'Manuell', unit: 't' };
  const snapshot = ()=>({
    aluminium,
    fx: {
      usdNok: { rate: 10, date: '2026-10-28', changes: { week: { rate: 9.8, percent: 2.04 }, month: { rate: 9.6, percent: 4.17 } } },
      eurNok: { rate: 11, date: '2026-10-28', changes: { week: { rate: 11.2, percent: -1.79 }, month: { rate: 11.3, percent: -2.65 } } }
    },
    updatedAt: fxUpdatedAt
  });
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname === '/api/market-data/aluminium' && request.method === 'PUT'){
      if (request.headers.authorization !== 'Bearer test-token'){
        response.statusCode = 401;
        response.end(JSON.stringify({ error: 'Unauthorized' }));
        return;
      }
      let body = '';
      for await (const chunk of request) body += chunk;
      const { price } = JSON.parse(body);
      aluminium = { ...aluminium, price, updatedAt: new Date().toISOString() };
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(snapshot()));
      return;
    }
    if (pathname === '/api/market-data'){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(snapshot()));
      return;
    }
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects: [], ownerEmails: ['owner@busbar.no'], isAdmin: true }
        : { customers: [] }));
      return;
    }
    const file = resolve(root, `.${pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.statusCode = 403;
      response.end();
      return;
    }
    try{
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.csv': 'text/csv' };
      response.setHeader('Content-Type', `${mime[extname(file)] || 'application/octet-stream'}; charset=utf-8`);
      response.end(await readFile(file));
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
  const context = await browser.newContext();
  await context.route('https://www.lme.com/**', route=>route.fulfill({ status: 200, body: 'LME' }));
  const page = await context.newPage();
  await page.addInitScript(()=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true,
      username: 'owner@busbar.no',
      token: 'test-token',
      isAdmin: true
    }));
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/calculator.html`);
  const priceButton = page.locator('#marketAlPrice');
  await priceButton.waitFor();
  await page.waitForFunction(()=>document.querySelector('#marketAlPrice')?.classList.contains('is-stale'));
  await page.setViewportSize({ width: 2400, height: 900 });
  for (const id of ['marketUsdWeek', 'marketUsdMonth', 'marketEurWeek', 'marketEurMonth']){
    const change = page.locator(`#${id}`);
    assert.match(await change.locator('.market-change-rate').textContent(), /\d/);
    assert.match(await change.locator('.market-change-percent').textContent(), /%/);
    assert.notEqual(await change.locator('.market-change-percent').evaluate(el=>getComputedStyle(el).display), 'none');
  }
  await page.setViewportSize({ width: 390, height: 844 });
  for (const id of ['marketUsdWeek', 'marketUsdMonth', 'marketEurWeek', 'marketEurMonth']){
    assert.equal(await page.locator(`#${id} .market-change-percent`).evaluate(el=>getComputedStyle(el).display), 'none');
  }
  await page.setViewportSize({ width: 1280, height: 720 });
  assert.equal(await priceButton.isEnabled(), true);
  assert.equal(await priceButton.evaluate(el=>getComputedStyle(el).color), 'rgb(254, 190, 16)');

  const popupPromise = context.waitForEvent('page');
  await priceButton.click();
  const popup = await popupPromise;
  await popup.waitForLoadState('domcontentloaded');
  assert.match(popup.url(), /lme\.com\/metals\/non-ferrous\/lme-aluminium/);
  await page.locator('#marketAluminiumInput').fill('3200');
  await page.locator('#marketAluminiumForm [type="submit"]').click();
  await page.waitForFunction(()=>!document.querySelector('#marketAlPrice')?.classList.contains('is-stale'));
  assert.equal(await priceButton.isEnabled(), false);
  assert.equal(await priceButton.evaluate(el=>getComputedStyle(el).color), 'rgb(241, 240, 240)');
  assert.equal(aluminium.price, 3200);

  fxUpdatedAt = new Date(Date.now() + 1000).toISOString();
  await page.reload();
  await page.waitForFunction(()=>document.querySelector('#marketAlPrice')?.classList.contains('is-stale'));
  assert.equal(await priceButton.isEnabled(), true);

  await page.goto(`http://127.0.0.1:${server.address().port}/index.html`);
  await page.waitForFunction(()=>document.querySelector('#marketAlPrice')?.classList.contains('is-stale'));
  assert.equal(await page.locator('#marketAlPrice').isEnabled(), true);
  await page.setViewportSize({ width: 390, height: 844 });
  await page.locator('#marketAlPrice').click();
  const modal = page.locator('#marketAluminiumFormModal .form-modal-panel');
  await modal.waitFor();
  const bounds = await modal.boundingBox();
  assert.ok(bounds.x >= 0 && bounds.x + bounds.width <= 390);
  assert.ok(bounds.y >= 0 && bounds.y + bounds.height <= 844);
  await page.keyboard.press('Escape');
  await page.locator('#marketAluminiumFormModal').waitFor({ state: 'hidden' });
});
