import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('new CSV variants load and retain the matching conductor prices', async t=>{
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects: [], ownerEmails: ['catalog@busbar.no'] }
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
      username: 'catalog@busbar.no',
      token: 'test-token'
    }));
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/calculator.html`);
  await page.locator('#status').getByText(/CSV lastet/).waitFor({ timeout: 15000 });

  await page.locator('#series').selectOption('XCP-S');
  const xcpAmps = await page.locator('#ampSelect option').allTextContents();
  for (const amp of ['630', '800', '1000']) assert.ok(xcpAmps.includes(amp), `${amp}A missing from XCP-S`);

  await page.locator('#ampSelect').selectOption('630');
  await page.locator('#meter').fill('1');
  await page.locator('#startEl').selectOption('board_feed');
  await page.locator('#sluttEl').selectOption('board_feed');
  await page.locator('#lineNumberInput').fill('CSV-TEST');
  for (const [ledere, expectedCode] of [
    ['3F+PE', '64281000P-3W'],
    ['3F+N+PE', '64281000P']
  ]){
    await page.locator('#ledere').selectOption(ledere);
    await page.locator('#calcBtn').click();
    await page.locator('#bomTbl').getByText(expectedCode, { exact: true }).first().waitFor();
    assert.ok((await page.locator('#bomTbl').textContent()).includes(expectedCode));
  }

  const readAmount = async selector=>Number((await page.locator(selector).textContent()).replace(/\s/g, '').replace(',', '.'));
  assert.equal(await page.locator('#montasjeHourlyRate').inputValue(), '890');
  assert.equal(await readAmount('#montasje'), 1780);
  assert.equal(await readAmount('#montasjeMargin'), 445);
  assert.equal(await readAmount('#montasjeStartupCost'), 20000);
  assert.equal(await readAmount('#totalInclMontasje'), 22225);
  const includedTotalBefore = await readAmount('#selectedAddonTotal');
  assert.equal(await page.locator('.montasje-startup-cost').evaluate(el=>getComputedStyle(el).fontStyle), 'italic');
  await page.locator('#montasjeDgConfigBtn').click();
  await page.locator('#marginPercentInput').fill('30');
  await page.locator('#marginSubmit').click();
  assert.equal(await readAmount('#montasjeMargin'), 762.86);
  assert.equal(await readAmount('#montasjeStartupCost'), 20000);
  assert.equal(await readAmount('#totalInclMontasje'), 22542.86);
  await page.locator('#montasjeDgConfigBtn').click();
  await page.locator('#marginPercentInput').fill('20');
  await page.locator('#marginSubmit').click();
  await page.locator('#rateToggle').check();
  await page.locator('#montasjeHourlyRate').fill('0');
  await page.locator('#calcBtn').click();
  assert.equal(await readAmount('#montasje'), 0);
  assert.equal(await readAmount('#montasjeStartupCost'), 0);
  assert.equal(await readAmount('#totalInclMontasje'), 0);
  assert.equal(includedTotalBefore - await readAmount('#selectedAddonTotal'), 22225);
  await page.locator('#rateToggle').uncheck();
  await page.locator('#calcBtn').click();

  await page.locator('#fbQty').fill('1');
  await page.locator('#calcBtn').click();
  const fireBarrierRow = page.locator('#bomTbl tr', { hasText: '672EFB01' }).first();
  await fireBarrierRow.waitFor();
  assert.match(await fireBarrierRow.textContent(), /3252\.05/);
  await page.locator('#fbQty').fill('0');

  await page.locator('#series').selectOption('RCP-IP68');
  const rcpAmps = await page.locator('#ampSelect option').allTextContents();
  for (const amp of ['630', '800', '1000']) assert.ok(rcpAmps.includes(amp), `${amp}A missing from RCP-IP68`);
  const epoxyConductors = await page.locator('#ledere option:not([disabled])').allTextContents();
  assert.ok(epoxyConductors.includes('3F+N'));
  assert.ok(epoxyConductors.includes('3F+N+PE'));
  assert.ok(!epoxyConductors.includes('3F'));

  await page.locator('#ampSelect').selectOption('1250');
  await page.locator('#ledere').selectOption('3F+N');
  await page.locator('#calcBtn').click();
  await page.locator('#bomTbl').getByText('60281004R', { exact: true }).first().waitFor();
  const epoxyBom = await page.locator('#bomTbl').textContent();
  for (const code of ['6028EJ11R', '6MLD4E03R', '65200001R', '65200002R']){
    assert.ok(epoxyBom.includes(code), `${code} missing from epoxy BOM`);
  }

  await page.locator('#meter').fill('7');
  const readBomLine = async code=>{
    const row = page.locator('#bomTbl tbody tr', { has: page.locator('td', { hasText: code }) }).first();
    await row.waitFor();
    const cells = await row.locator('td').allTextContents();
    return { count: Number(cells[5]), price: Number(cells[6]) };
  };
  const variants = [
    ['630', '3F+N', '6028EJ11R', 2320.88, '6MLD4E01R', 3683.80, 1456.64],
    ['630', '3F+N+PE', '6028EJ21R', 2658.36, '6MLD5E01R', 3918.10, 1456.64],
    ['800', '3F+N', '6028EJ11R', 2320.88, '6MLD4E02R', 3559.50, 1456.64],
    ['800', '3F+N+PE', '6028EJ21R', 2658.36, '6MLD5E02R', 3680.05, 1456.64],
    ['1000', '3F+N', '6028EJ11R', 2320.88, '6MLD4E02R', 3559.50, 1456.64],
    ['1000', '3F+N+PE', '6028EJ21R', 2658.36, '6MLD5E02R', 3680.05, 1456.64],
    ['1250', '3F+N', '6028EJ11R', 2320.88, '6MLD4E03R', 3524.10, 1456.64],
    ['1250', '3F+N+PE', '6028EJ21R', 2658.36, '6MLD5E03R', 3644.65, 1456.64],
    ['1600', '3F+N', '6028EJ12R', 2928.36, '6MLD4E05R', 3793.05, 1456.64],
    ['1600', '3F+N+PE', '6028EJ22R', 3315.28, '6MLD5E05R', 3913.55, 1456.64],
    ['2000', '3F+N', '6028EJ14R', 3711.84, '6MLD4E07R', 4079.30, 2184.96],
    ['2000', '3F+N+PE', '6028EJ24R', 4188.52, '6MLD5E07R', 4206.60, 2184.96],
    ['2500', '3F+N', '6039EJ11R', 5974.84, '6MLD4E11R', 4893.65, 2913.28],
    ['2500', '3F+N+PE', '6039EJ21R', 7019.88, '6MLD5E11R', 5130.95, 2913.28],
    ['3200', '3F+N', '6039EJ12R', 5797.64, '6MLD4E13R', 5463.15, 2913.28],
    ['3200', '3F+N+PE', '6039EJ22R', 6482.88, '6MLD5E13R', 5690.70, 3641.60],
    ['4000', '3F+N', '6039EJ13R', 6828.20, '6MLD4E14R', 5537.00, 3641.60],
    ['4000', '3F+N+PE', '6039EJ23R', 7526.72, '6MLD5E14R', 5656.05, 3641.60],
    ['5000', '3F+N', '6039EJ11R', 11949.68, '6MLD4E11R', 9787.30, 5098.24],
    ['5000', '3F+N+PE', '6039EJ21R', 14039.76, '6MLD5E11R', 10261.90, 5098.24]
  ];
  for (const [amp, ledere, monoblocCode, monoblocPrice, moldCode, moldPrice, junctionPrice] of variants){
    await page.locator('#ampSelect').selectOption(amp);
    await page.locator('#ledere').selectOption(ledere);
    await page.locator('#calcBtn').click();
    const monobloc = await readBomLine(monoblocCode);
    const mold = await readBomLine(moldCode);
    const junction = await readBomLine('65200001R');
    assert.equal(monobloc.price, monoblocPrice, `${amp}A ${ledere} monobloc price`);
    assert.equal(mold.price, moldPrice, `${amp}A ${ledere} mold price`);
    assert.equal(junction.price, junctionPrice, `${amp}A ${ledere} junction price`);
    assert.equal(mold.count, monobloc.count, `${amp}A ${ledere} mold count`);
    if (Number(amp) <= 1600){
      assert.ok(monobloc.count >= 3);
      assert.equal(junction.count, Math.ceil(monobloc.count * 0.5), `${amp}A ${ledere} junction count`);
    }
  }

  await page.locator('#ampSelect').selectOption('800');
  await page.locator('#ledere').selectOption('3F+N+PE');
  await page.locator('#v90v').fill('1');
  await page.locator('#calcBtn').click();
  assert.equal((await readBomLine('6MLD5R02R')).price, 1468.70);
  assert.equal((await readBomLine('65200002R')).price, 2532.10);
  assert.equal(await page.locator('#bomTbl tbody tr', { hasText: '6MLD4R02R' }).count(), 0);

  await page.locator('#ampSelect').selectOption('3200');
  await page.locator('#calcBtn').click();
  assert.equal((await readBomLine('6MLD5R13R')).price, 1922.80);
  assert.equal((await readBomLine('65200002R')).price, 6330.25);
});
