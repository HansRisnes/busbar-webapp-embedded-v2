import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('dashboard report exports calculated and offered hours by project status', async t=>{
  const server = createServer(async (request, response)=>{
    if (request.url === '/'){
      response.setHeader('Content-Type', 'text/html');
      response.end('<!doctype html><title>Report test</title>');
      return;
    }
    const path = resolve(root, `.${new URL(request.url, 'http://localhost').pathname}`);
    if (!path.startsWith(`${root}${sep}`)){
      response.writeHead(403).end();
      return;
    }
    try{
      response.setHeader('Content-Type', 'text/javascript');
      response.end(await readFile(path));
    }catch{
      response.writeHead(404).end();
    }
  });
  await new Promise(resolveListening=>server.listen(0, '127.0.0.1', resolveListening));
  t.after(()=>new Promise(resolveClose=>server.close(resolveClose)));

  const browser = await chromium.launch({ headless: true });
  t.after(()=>browser.close());
  const page = await browser.newPage();
  await page.goto(`http://127.0.0.1:${server.address().port}/`);
  const result = await page.evaluate(async ()=>{
    const { exportDashboardTotalsReport } = await import('/src/client/dashboard.js');
    URL.createObjectURL = blob=>{
      window.reportBlob = blob;
      return 'blob:report-test';
    };
    HTMLAnchorElement.prototype.click = function(){};
    const line = (engineering, montasje, includeEngineering, includeMontasje)=>({
      totals: {
        engineering: { totalHours: engineering },
        montasje: { totalHours: montasje }
      },
      selectedAddonConfig: { includeEngineering, includeMontasje }
    });
    const projects = [
      { projectNumber: '1', name: 'Åpent', projectStatus: 'unresolved', createdAt: '2026-01-02', lines: [line(2, 3, false, false)] },
      { projectNumber: '2', name: 'Vunnet', projectStatus: 'won', createdAt: '2026-01-03', lines: [line(4, 5, true, false)] },
      { projectNumber: '3', name: 'Tapt', projectStatus: 'lost', createdAt: '2026-01-04', lines: [line(6, 7, false, true)] },
      { projectNumber: '4', name: 'Ferdig', projectStatus: 'finished', createdAt: '2026-02-04', lines: [line(8, 9, true, true)] }
    ];
    exportDashboardTotalsReport({ totalsYear: '2026', totalsMonth: '1' }, projects, {
      getProjectStatusConfig: project=>({ id: project.projectStatus, label: project.projectStatus })
    });
    const spreadsheet = 'urn:schemas-microsoft-com:office:spreadsheet';
    const readHourSheet = async ()=>{
      const xml = new DOMParser().parseFromString(await window.reportBlob.text(), 'application/xml');
      const sheet = [...xml.getElementsByTagNameNS(spreadsheet, 'Worksheet')]
        .find(node=>node.getAttributeNS(spreadsheet, 'Name') === 'Timer');
      return {
        xmlError: xml.getElementsByTagName('parsererror').length,
        rows: [...sheet.getElementsByTagNameNS(spreadsheet, 'Row')].map(row=>
          [...row.getElementsByTagNameNS(spreadsheet, 'Data')].map(cell=>cell.textContent))
      };
    };
    const january = await readHourSheet();
    exportDashboardTotalsReport({ totalsYear: '2026', totalsMonth: 'all' }, projects, {
      getProjectStatusConfig: project=>({ id: project.projectStatus, label: project.projectStatus })
    });
    return {
      january,
      fullYear: await readHourSheet()
    };
  });
  assert.equal(result.january.xmlError, 0);
  assert.equal(result.fullYear.xmlError, 0);
  const row = label=>result.january.rows.find(cells=>cells[0] === label);
  assert.deepEqual(row('Alle prosjekter'), ['Alle prosjekter', '12', '4', '15', '7']);
  assert.deepEqual(row('Uavklart'), ['Uavklart', '2', '0', '3', '0']);
  assert.deepEqual(row('Vunnet'), ['Vunnet', '4', '4', '5', '0']);
  assert.deepEqual(row('Tapt'), ['Tapt', '6', '0', '7', '7']);
  assert.deepEqual(row('Ferdig'), ['Ferdig', '0', '0', '0', '0']);
  assert.deepEqual(row('Vunnet og ferdig'), ['Vunnet og ferdig', '4', '4', '5', '0']);
  assert.deepEqual(row('1'), ['1', 'Åpent', 'unresolved', '2', '0', '3', '0']);
  assert.equal(row('4'), undefined);
  const fullYearRow = label=>result.fullYear.rows.find(cells=>cells[0] === label);
  assert.deepEqual(fullYearRow('Alle prosjekter'), ['Alle prosjekter', '20', '12', '24', '16']);
  assert.deepEqual(fullYearRow('Ferdig'), ['Ferdig', '8', '8', '9', '9']);
  assert.deepEqual(fullYearRow('Vunnet og ferdig'), ['Vunnet og ferdig', '12', '12', '14', '9']);
});
