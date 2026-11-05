import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('leads page creates, filters, edits and deletes a lead', async t => {
  let leads = [];
  let editable = true;
  let ownerView = false;
  const payload = () => ({
    leads: leads.map(lead => ({ ...lead, responsibleName: 'Selger', canEdit: editable })),
    users: ownerView
      ? [{ email: 'seller@busbar.no', name: 'Selger' }, { email: 'owner2@busbar.no', name: 'Annen eier' }]
      : [{ email: 'seller@busbar.no', name: 'Selger' }],
    customers: [{ name: 'Bygg AS', contacts: [{ name: 'Kari Kontakt', email: 'kari@example.no', phone: '12345678' }] }],
    canAssignOthers: ownerView
  });
  const server = createServer(async (request, response) => {
    const url = new URL(request.url, 'http://localhost');
    if (url.pathname.startsWith('/api/')) {
      if (url.pathname === '/api/leads/upsert' || url.pathname === '/api/leads/delete' || url.pathname === '/api/leads/complete') {
        let body = '';
        for await (const chunk of request) body += chunk;
        const input = JSON.parse(body);
        if (url.pathname.endsWith('/delete')) {
          leads = leads.filter(lead => lead.id !== input.id);
        } else if (url.pathname.endsWith('/complete')) {
          leads = leads.map(lead => lead.id === input.id ? { ...lead, followUpCompleted: input.completed, updatedAt: new Date().toISOString() } : lead);
        } else {
          const lead = { ...input, id: input.id || 'lead-1', createdAt: '2026-11-03T10:00:00.000Z', updatedAt: new Date().toISOString() };
          leads = input.id ? leads.map(item => item.id === input.id ? lead : item) : [...leads, lead];
        }
      }
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(url.pathname.startsWith('/api/leads') ? payload() : { projects: [], customers: [] }));
      return;
    }
    const file = resolve(root, `.${url.pathname === '/' ? '/index.html' : url.pathname}`);
    if (!file.startsWith(`${root}${sep}`)) {
      response.statusCode = 403;
      response.end();
      return;
    }
    try {
      const content = await readFile(file);
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.ttf': 'font/ttf' };
      response.setHeader('Content-Type', `${mime[extname(file)] || 'application/octet-stream'}; charset=utf-8`);
      response.end(content);
    } catch {
      response.statusCode = 404;
      response.end();
    }
  });
  await new Promise(resolveListening => server.listen(0, '127.0.0.1', resolveListening));
  t.after(() => {
    server.closeAllConnections();
    return new Promise(resolveClose => server.close(resolveClose));
  });
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.addInitScript(() => {
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true, username: 'seller@busbar.no', token: 'test-token'
    }));
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=leads`);
  await page.locator('#leadsStatus').getByText(/Oppdatert/).waitFor();
  assert.equal(await page.locator('#leadsList .lead-item').count(), 0);
  await page.locator('#newLeadBtn').click();
  assert.equal(await page.locator('#leadCompanyName').getAttribute('autocomplete'), 'off');
  assert.equal(await page.locator('#leadContactName').getAttribute('autocomplete'), 'off');
  assert.equal(await page.locator('#leadCompanyName').getAttribute('type'), 'search');
  const firstCompanyLookupName = await page.locator('#leadCompanyName').getAttribute('name');
  const modalTop = () => page.locator('#leadFormModal .form-modal-panel').evaluate(el => el.getBoundingClientRect().top);
  const initialTop = await modalTop();
  await page.locator('#leadCompanyName').fill('Byg');
  await page.locator('#leadCustomerSuggestions li').getByText('Bygg AS').click();
  assert.equal(await page.locator('#leadCompanyName').inputValue(), 'Bygg AS');
  await page.locator('#leadContactName').fill('Kar');
  await page.locator('#leadContactSuggestions li').getByText('Kari Kontakt').click();
  assert.equal(await page.locator('#leadContactName').inputValue(), 'Kari Kontakt');
  assert.equal(await page.locator('#leadContactEmail').inputValue(), 'kari@example.no');
  assert.equal(await page.locator('#leadContactPhone').inputValue(), '12345678');
  await page.locator('#leadProjectNeed').fill('Strømskinner til ny hall');
  await page.locator('#leadNextActivity').selectOption('Oppfølging');
  await page.locator('#leadFollowUpDate').fill('20/11/2026');
  await page.locator('#leadFollowUpDatePickerBtn').click();
  assert.equal(await page.locator('#leadDatePickerPopover').isVisible(), true);
  await page.keyboard.press('Escape');
  await page.locator('.lead-extra-fields summary').click();
  assert.ok(Math.abs((await modalTop()) - initialTop) < 1, 'Opening optional fields must not move the modal top');
  assert.deepEqual(await page.locator('#leadEstimatedValue option').evaluateAll(options =>
    options.map(option => option.value).filter(Boolean)),
  ['50000', '75000', '100000', '200000', '300000', '400000', '500000', '600000', '700000', '800000', '900000', '1000000']);
  assert.deepEqual(await page.locator('#leadProbability option').evaluateAll(options =>
    options.map(option => option.value).filter(Boolean)),
  ['0', '10', '20', '30', '40', '50', '60', '70', '80', '90', '100']);
  await page.locator('#leadEstimatedValue').selectOption('500000');
  await page.locator('#leadProbability').selectOption('30');
  await page.locator('#leadProjectPhase').fill('An');
  await page.locator('#leadProjectPhaseSuggestions li').getByText('Anbud').click();
  await page.locator('#leadCustomerType').fill('Elektro');
  await page.locator('#leadCustomerTypeSuggestions li').getByText('Elektroentreprenør').click();
  await page.locator('#leadSource').fill('Nett');
  await page.locator('#leadSourceSuggestions li').getByText('Nettside').click();
  await page.locator('#leadDecisionDate').fill('25/12/2026');
  await page.locator('#leadDecisionDatePickerBtn').click();
  assert.equal(await page.locator('#leadDatePickerPopover .calendar-date-picker-day.is-selected').innerText(), '25');
  await page.locator('#leadDatePickerPopover .calendar-date-picker-day:not(.is-outside)').getByText('26', { exact: true }).click();
  assert.equal(await page.locator('#leadDecisionDate').inputValue(), '26/12/2026');
  await page.locator('#leadComments').fill('Oppfølging planlagt.');
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-extra-desktop.png'), fullPage: true });
  }
  await page.locator('#saveLeadBtn').click();
  await page.locator('#leadsList .lead-item').waitFor();
  assert.equal(await page.locator('.lead-item-title').innerText(), 'Strømskinner til ny hall');
  assert.doesNotMatch(await page.locator('#leadsList .lead-item').innerText(), /Fulgt opp/);
  const checkboxAndTitle = await page.locator('.lead-item').evaluate(el => ({
    checkbox: el.querySelector('.lead-follow-up-checkbox').getBoundingClientRect(),
    title: el.querySelector('.lead-item-title').getBoundingClientRect()
  }));
  assert.ok(checkboxAndTitle.checkbox.right < checkboxAndTitle.title.left);
  assert.match(await page.locator('#leadsList .lead-item').innerText(), /Bedriftsnavn:.*Bygg AS/);
  assert.match(await page.locator('#leadsList .lead-item').innerText(), /Kontaktperson:.*Kari Kontakt/);
  assert.equal(await page.locator('.lead-item-title').evaluate(el => getComputedStyle(el).fontSize), '24px');
  assert.equal(await page.locator('.lead-item-label').first().evaluate(el => getComputedStyle(el).fontSize), '16px');
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-desktop.png'), fullPage: true });
  }
  assert.match(await page.locator('#leadsSummary').innerText(), /1 lead/);
  assert.match(await page.locator('#leadsList .lead-item').innerText(), /Oppfølging/);
  assert.match(await page.locator('#leadsList .lead-item').innerText(), /500\s?000 NOK/);
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForFunction(() => parseFloat(getComputedStyle(document.body).paddingLeft) <= 15);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth + 1), false);
  const mobileCard = await page.locator('.lead-item').evaluate(el => ({
    card: el.getBoundingClientRect().right,
    details: el.querySelector('.lead-item-details').getBoundingClientRect().right,
    actions: el.querySelector('.lead-item-actions').getBoundingClientRect().right
  }));
  assert.ok(mobileCard.details <= mobileCard.card && mobileCard.actions <= mobileCard.card);
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-card-mobile.png'), fullPage: true });
  }
  await page.setViewportSize({ width: 1440, height: 900 });
  assert.equal(leads[0].estimatedValue, '500000');
  assert.equal(leads[0].probability, '30');
  assert.equal(leads[0].decisionDate, '2026-12-26');
  assert.equal(leads[0].projectPhase, 'Anbud');
  assert.equal(leads[0].leadSource, 'Nettside');
  leads[0].estimatedValue = 450000;
  leads[0].probability = 25;
  await page.locator('#refreshLeadsBtn').click();
  await page.locator('#leadsList .lead-item').click();
  assert.equal(await page.locator('#leadEstimatedValue').inputValue(), '450000');
  assert.equal(await page.locator('#leadProbability').inputValue(), '25');
  assert.equal(await page.locator('#leadDecisionDate').inputValue(), '26/12/2026');
  await page.locator('#cancelLeadBtn').click();
  const soon = new Date();
  soon.setDate(soon.getDate() + 2);
  leads[0].followUpDate = `${soon.getFullYear()}-${String(soon.getMonth() + 1).padStart(2, '0')}-${String(soon.getDate()).padStart(2, '0')}`;
  await page.locator('[data-dashboard-nav="dashboard"]').click();
  const leadAlert = page.locator('.dashboard-flow-alert-cell[data-dashboard-flow-alert-lead="lead-1"]');
  await leadAlert.waitFor();
  assert.equal(await leadAlert.evaluate(el => el.getBoundingClientRect().width), 240);
  assert.equal(await leadAlert.locator('.dashboard-flow-alert-lead-heading strong').innerText(), 'Strømskinner til ny hall');
  assert.equal(await leadAlert.locator('.dashboard-flow-alert-activity').innerText(), 'Neste aktivitet: Oppfølging');
  assert.equal(await leadAlert.locator('span').last().innerText(), 'Frist om 2 dager');
  assert.doesNotMatch(await leadAlert.innerText(), /Bygg AS/);
  assert.equal(await leadAlert.locator('.dashboard-flow-alert-lead-heading').evaluate(el => el.scrollHeight > el.clientHeight + 1), false);
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.locator('#dashboardFlowAlerts').screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-alert-desktop.png') });
  }
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForFunction(() => parseFloat(getComputedStyle(document.body).paddingLeft) <= 15);
  assert.equal(await leadAlert.evaluate(el => el.getBoundingClientRect().width), 240);
  assert.equal(await leadAlert.locator('.dashboard-flow-alert-lead-heading').evaluate(el => el.scrollHeight > el.clientHeight + 1), false);
  assert.equal(await leadAlert.locator(':scope > span').last().evaluate(el => el.scrollHeight > el.clientHeight + 1), false);
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.locator('#dashboardFlowAlerts').screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-alert-mobile.png') });
  }
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.getByRole('button', { name: 'Vis alle' }).click();
  const leadAlertListItem = page.locator('#dashboardFlowAlertsList [data-dashboard-flow-alert-lead="lead-1"]');
  assert.equal(await leadAlertListItem.locator('.dashboard-flow-alert-lead-heading strong').innerText(), 'Strømskinner til ny hall');
  assert.equal(await leadAlertListItem.locator('.dashboard-flow-alert-activity').innerText(), 'Neste aktivitet: Oppfølging');
  assert.doesNotMatch(await leadAlertListItem.innerText(), /Bygg AS/);
  await page.locator('#dashboardFlowAlertsClose').click();
  await leadAlert.click();
  await page.locator('#leadFormModal:not([hidden])').waitFor();
  assert.equal(await page.locator('#leadCompanyName').inputValue(), 'Bygg AS');
  await page.locator('#cancelLeadBtn').click();
  await page.locator('.lead-follow-up-checkbox').check();
  await page.waitForFunction(() => document.querySelector('.lead-follow-up-checkbox')?.checked === true);
  assert.equal(leads[0].followUpCompleted, true);

  await page.locator('#leadSearchInput').fill('finnes ikke');
  assert.equal(await page.locator('#leadsList .lead-item').count(), 0);
  await page.locator('#leadSearchInput').fill('Bygg');
  assert.equal(await page.locator('#leadsList .lead-item').count(), 1);
  await page.locator('#leadPhaseFilter').selectOption('qualified');
  assert.equal(await page.locator('#leadsList .lead-item').count(), 0);
  await page.locator('#leadPhaseFilter').selectOption('all');

  await page.locator('#leadsList .lead-item').click();
  const editTop = await modalTop();
  await page.locator('.lead-extra-fields summary').click();
  assert.ok(Math.abs((await modalTop()) - editTop) < 1, 'Collapsing optional fields must not move the modal top');
  await page.locator('#leadSalesPhase').selectOption('qualified');
  await page.locator('#saveLeadBtn').click();
  await page.locator('#leadsList .lead-phase.is-qualified').waitFor();
  await page.locator('#leadsList .lead-item').click();
  page.once('dialog', dialog => dialog.accept());
  await page.locator('#deleteLeadBtn').click();
  await page.locator('#leadsList .leads-empty').waitFor();
  assert.equal(leads.length, 0);

  ownerView = true;
  await page.locator('#refreshLeadsBtn').click();
  await page.locator('#newLeadBtn').click();
  assert.notEqual(await page.locator('#leadCompanyName').getAttribute('name'), firstCompanyLookupName);
  assert.equal(await page.locator('#leadResponsibleEmail option').count(), 2);
  assert.equal(await page.locator('#leadResponsibleEmail').isDisabled(), false);
  await page.locator('#leadResponsibleEmail').selectOption('owner2@busbar.no');
  assert.equal(await page.locator('#leadResponsibleEmail').inputValue(), 'owner2@busbar.no');
  await page.locator('#cancelLeadBtn').click();
  ownerView = false;

  editable = false;
  leads = [{
    id: 'read-only-lead', companyName: 'Innsyn AS', projectNeed: 'Strømskinner',
    nextActivity: 'Send oppfølging', followUpDate: '2026-11-20', salesPhase: 'new',
    estimatedValue: null, createdAt: '2026-11-03T10:00:00.000Z', updatedAt: '2026-11-03T10:00:00.000Z'
  }];
  await page.locator('#leadSearchInput').fill('');
  await page.locator('#refreshLeadsBtn').click();
  await page.locator('#leadsList .lead-item').click();
  assert.equal(await page.locator('#leadCompanyName').isDisabled(), true);
  assert.equal(await page.locator('#saveLeadBtn').isVisible(), false);
  assert.equal(await page.locator('#deleteLeadBtn').isVisible(), false);
  await page.locator('#cancelLeadBtn').click();

  await page.setViewportSize({ width: 390, height: 844 });
  await page.locator('#leadsList .lead-item').click();
  const mobileTop = await modalTop();
  await page.locator('.lead-extra-fields summary').click();
  assert.ok(Math.abs((await modalTop()) - mobileTop) < 1, 'Mobile optional fields must not move the modal top');
  const mobileBottom = await page.locator('#leadFormModal .form-modal-panel').evaluate(el => el.getBoundingClientRect().bottom);
  assert.ok(mobileBottom <= 844, 'Expanded modal must remain inside the mobile viewport');
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > innerWidth + 1);
  assert.equal(overflow, false, 'Leads page should fit a mobile viewport');
  assert.equal(await page.locator('#leadFormModal').isVisible(), true);
  await page.locator('#cancelLeadBtn').scrollIntoViewIfNeeded();
  assert.equal(await page.locator('#cancelLeadBtn').isVisible(), true);
  if (process.env.LEADS_SCREENSHOT_DIR) {
    await page.screenshot({ path: resolve(process.env.LEADS_SCREENSHOT_DIR, 'leads-mobile.png'), fullPage: true });
  }
});
