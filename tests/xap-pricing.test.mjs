import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import {
  adjustXapUsdPrice,
  getCurrentAluminiumPrice,
  isXapIndexedType,
  XAP_ALUMINIUM_REFERENCE_USD_PER_TON
} from '../src/client/xap-pricing.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('XAP busway and flexibles use the supplier aluminium adjustment', ()=>{
  assert.equal(XAP_ALUMINIUM_REFERENCE_USD_PER_TON, 2562);
  assert.equal(adjustXapUsdPrice(1000, 'straight_3m', 2800), 1074.32);
  assert.equal(adjustXapUsdPrice(1000, 'straight_1501_2000', 2800), 1074.32);
  assert.equal(adjustXapUsdPrice(1000, 'flexible_set', 2300), 918.19);
  assert.equal(adjustXapUsdPrice(1000, 'elbow_horizontal_90', 2562), 1000);
  assert.equal(adjustXapUsdPrice(1000, 'board_feed', 2800), 1074.32);
  assert.equal(isXapIndexedType('outlet_section'), false);
  assert.equal(adjustXapUsdPrice(1000, 'outlet_section', 2800), 1000);
  assert.equal(adjustXapUsdPrice(1000, 'joint', 2800), 1000);
  assert.ok(Number.isNaN(adjustXapUsdPrice(1000, 'straight_3m', null)));
});

test('only a manual price entered after the FX update may index XAP prices', ()=>{
  const fxUpdatedAt = '2026-10-28T10:00:00Z';
  const quote = {
    price: 2800,
    currency: 'USD',
    symbol: 'MANUAL',
    source: 'Manuell',
    unit: 't',
    updatedAt: '2026-10-28T10:01:00Z'
  };
  assert.equal(getCurrentAluminiumPrice(quote, fxUpdatedAt), 2800);
  assert.equal(getCurrentAluminiumPrice({ ...quote, source: 'Yahoo Finance (COMEX)' }, fxUpdatedAt), null);
  assert.equal(getCurrentAluminiumPrice({ ...quote, updatedAt: '' }, fxUpdatedAt), null);
  assert.equal(getCurrentAluminiumPrice({ ...quote, updatedAt: '2026-10-28T09:59:00Z' }, fxUpdatedAt), null);
  assert.equal(getCurrentAluminiumPrice({ ...quote, price: 0 }, fxUpdatedAt), null);
});

test('calculator applies the current manual quote to XAP busway before NOK conversion', async t=>{
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname === '/api/market-data'){
      response.setHeader('Content-Type', 'application/json');
      const fxUpdatedAt = new Date(Date.now() - 60_000).toISOString();
      response.end(JSON.stringify({
        aluminium: {
          price: 2800,
          currency: 'USD',
          symbol: 'MANUAL',
          source: 'Manuell',
          unit: 't',
          updatedAt: new Date().toISOString()
        },
        fx: { usdNok: { rate: 10, date: '2026-10-28' } },
        updatedAt: fxUpdatedAt
      }));
      return;
    }
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects: [], ownerEmails: ['xap@busbar.no'] }
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
  const page = await browser.newPage();
  await page.addInitScript(()=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true,
      username: 'xap@busbar.no',
      token: 'test-token'
    }));
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/calculator.html`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });
  await page.getByText('Manuell', { exact: false }).first().waitFor();
  assert.equal(await page.locator('#series option[value="XAP-B"]').count(), 1);
  await page.locator('#series').selectOption('XAP-B');
  assert.equal(await page.locator('#ledere').inputValue(), '3F+N+PE');
  assert.equal(await page.locator('#ledere').isDisabled(), true);
  assert.ok(await page.locator('#ampSelect option[value="400"]').count());
  await page.locator('#ampSelect').selectOption('500');
  await page.locator('#meter').fill('1');
  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('end_cover');
  await page.locator('#lineNumberInput').fill('XAP-TEST');
  await page.locator('#calcBtn').click();
  const row = page.locator('#bomTbl tr', { hasText: 'XAP-B 500A, 4W - Straight length, 1 meter' }).first();
  await page.waitForTimeout(300);
  assert.ok(await row.count(), `XAP calculation failed: ${await page.locator('#status').textContent()} | ${await page.locator('body').innerText().then(text=>text.slice(-500))}`);
  const expected = Math.round(adjustXapUsdPrice(78.54, 'straight_500_1000', 2800) * 10 * 100) / 100;
  const amount = Number((await row.locator('td').last().textContent()).replace(/\s/g, '').replace(',', '.'));
  assert.equal(amount, expected);
  assert.equal(await page.locator('#xapCompareLine').isVisible(), false);
  const bomQty = async label=>{
    const row = page.locator('#bomTbl tbody tr', { hasText: label }).first();
    assert.ok(await row.count(), `BOM-rad mangler: ${label}`);
    return Number((await row.locator('td').nth(5).textContent()).trim());
  };
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 3);
  assert.equal(await page.locator('#bomTbl tbody tr', { hasText: 'Outlet for tap-off boxes' }).count(), 0);

  await page.locator('#dist').selectOption('Ja');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Outlet for tap-off boxes'), 2);
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 3);

  await page.locator('#meter').fill('2');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 2 meters'), 1);
  const twoMeterRow = page.locator('#bomTbl tbody tr', { hasText: 'XAP-B 500A, 4W - Straight length, 2 meters' }).first();
  const twoMeterUnit = Number((await twoMeterRow.locator('td').nth(6).textContent()).trim());
  assert.equal(twoMeterUnit, Math.round(adjustXapUsdPrice(78.54 * 2, 'straight_1501_2000', 2800) * 10 * 100) / 100);
  assert.equal(await bomQty('XAP-B 500A, 4W - Outlet for tap-off boxes'), 3);
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 3);

  await page.locator('#meter').fill('3');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 2 meters'), 1);
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 1 meter'), 1);
  assert.equal(await page.locator('#bomTbl tbody tr', { hasText: 'Straight length, 3 meters max.' }).count(), 0);
  assert.equal(await bomQty('XAP-B 500A, 4W - Outlet for tap-off boxes'), 5);
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 4);

  await page.locator('#meter').fill('5');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 2 meters'), 2);
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 1 meter'), 1);
  assert.equal(await bomQty('XAP-B 500A, 4W - Outlet for tap-off boxes'), 8);
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 5);

  await page.locator('#meter').fill('10');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 2 meters'), 5);
  assert.equal(await page.locator('#bomTbl tbody tr', { hasText: 'Straight length, 3 meters max.' }).count(), 0);
  assert.equal(await bomQty('XAP-B 500A, 4W - Outlet for tap-off boxes'), 15);

  await page.locator('#meter').fill('11');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 3 meters max.'), 3);
  assert.equal(await bomQty('XAP-B 500A, 4W - Straight length, 2 meters'), 1);

  await page.locator('#v90h').fill('1');
  await page.locator('#calcBtn').click();
  assert.equal(await bomQty('XAP-B 500A, 4W - Joint'), 7);
  await page.locator('#v90h').fill('0');
  await page.locator('#dist').selectOption('Nei');
  await page.locator('#calcBtn').click();
  assert.equal(await page.locator('#bomTbl tbody tr', { hasText: 'Outlet for tap-off boxes' }).count(), 0);

  await page.locator('#startEl').selectOption('end_feed_unit');
  await page.locator('#calcBtn').click();
  assert.match(await page.locator('#status').textContent(), /Mangler pris for XAP-B endetilf/);
  assert.equal(await page.locator('#results').isVisible(), false);
  assert.equal(await page.locator('#saveLineBtn').isDisabled(), true);

  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('crt_board_feed');
  await page.locator('#calcBtn').click();
  assert.match(await page.locator('#status').textContent(), /Mangler pris for XAP-B trafoelement/);

  await page.locator('#sluttEl').selectOption('end_cover');
  await page.locator('#ampSelect').selectOption('3200');
  await page.locator('#meter').fill('3');
  await page.locator('#calcBtn').click();
  assert.match(await page.locator('#status').textContent(), /^OK$/);
  assert.equal(await bomQty('XAP-B 3200A, 4W - Straight length, 2 meters'), 1);
  assert.equal(await bomQty('XAP-B 3200A, 4W - Straight length, 1 meter'), 1);
  assert.equal(await bomQty('XAP-B 3200A, 4W - Joint'), 4);

  await page.locator('#ampSelect').selectOption('400');
  await page.locator('#meter').fill('1');
  await page.locator('#calcBtn').click();
  assert.match(await page.locator('#status').textContent(), /^OK$/);
  assert.ok(await page.locator('#bomTbl tr', { hasText: 'XAP-B 400A, 4W - Straight length, 1 meter' }).count());

  await page.locator('#ampSelect').selectOption('5000');
  await page.locator('#meter').fill('11');
  await page.locator('#v90v').fill('1');
  await page.locator('#calcBtn').click();
  assert.match(await page.locator('#status').textContent(), /^OK$/);
  const fiveThousandThreeMeter = page.locator('#bomTbl tbody tr', { hasText: 'XAP-B 5000A, 4W - Straight length, 3 meters max.' }).first();
  const threeMeterUnit = Number((await fiveThousandThreeMeter.locator('td').nth(6).textContent()).trim());
  assert.equal(threeMeterUnit, Math.round(adjustXapUsdPrice(1626.24, 'straight_3m', 2800) * 10 * 100) / 100);
  assert.equal(await bomQty('XAP-B 5000A, 4W - L vertical elbow'), 1);

  await page.locator('#v90v').fill('0');
  await page.locator('#dist').selectOption('Ja');
  await page.locator('#addBoxItemBtn').click();
  assert.match(await page.locator('#status').textContent(), /Mangler pris for avtappingsbokser til XAP-B/);
});
