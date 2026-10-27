import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('dashboard widgets keep their intended order at desktop and mobile widths', async t=>{
  const server = createServer(async (request, response)=>{
    const url = new URL(request.url, 'http://localhost');
    if (url.pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(url.pathname.startsWith('/api/user-projects')
        ? { updatedAt: new Date().toISOString(), projects: [], ownerEmails: ['layout@busbar.no'] }
        : { customers: [] }));
      return;
    }
    const file = resolve(root, `.${url.pathname === '/' ? '/index.html' : url.pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.statusCode = 403;
      response.end();
      return;
    }
    try{
      const content = await readFile(file);
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.ttf': 'font/ttf' };
      response.setHeader('Content-Type', `${mime[extname(file)] || 'application/octet-stream'}; charset=utf-8`);
      response.end(content);
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
  for (const width of [1600, 1400, 1381, 1200, 1000, 390]){
    const page = await browser.newPage({ viewport: { width, height: 900 } });
    await page.addInitScript(()=>{
      sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
        loggedIn: true,
        username: 'layout@busbar.no',
        token: 'test-token'
      }));
    });
    await page.goto(`http://127.0.0.1:${server.address().port}/index.html?view=dashboard`);
    await page.locator('#dashboardProjectStatusSummary').waitFor();
    await page.evaluate(()=>document.fonts.ready);
    assert.equal(await page.getByText('Tomt vindu', { exact: true }).count(), 0);
    if (width === 1600){
      assert.equal(await page.locator('#dashboardRecommendedActions').textContent(),
        'Ingen anbefalte handlinger akkurat nå.');
      assert.equal(await page.locator('#dashboardEmailProjectSuggestions').textContent(),
        'Ingen e-posttråder foreslått.');
      const scrollbars = await page.evaluate(()=>{
        const styles = className=>{
          const element = document.createElement('div');
          element.className = className;
          document.body.appendChild(element);
          const result = {
            colors: getComputedStyle(element).scrollbarColor,
            width: getComputedStyle(element, '::-webkit-scrollbar').width,
            track: getComputedStyle(element, '::-webkit-scrollbar-track').backgroundColor,
            thumb: getComputedStyle(element, '::-webkit-scrollbar-thumb').backgroundColor
          };
          element.remove();
          return result;
        };
        return {
          actions: styles('dashboard-recommended-action-groups'),
          suggestions: styles('dashboard-email-project-suggestion-list')
        };
      });
      assert.deepEqual(scrollbars.actions, scrollbars.suggestions);
    }
    const boxes = await page.evaluate(()=>{
      const box = selector=>{
        const rect = document.querySelector(selector).getBoundingClientRect();
        return { left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom };
      };
      return {
        total: box('.dashboard-total-widget'),
        actions: box('.dashboard-actions-widget'),
        status: box('.dashboard-status-widget'),
        flow: box('.dashboard-flow-status-widget'),
        todo: box('.dashboard-todo-widget'),
        suggestions: box('.dashboard-email-project-widget'),
        copilot: box('.dashboard-copilot-widget')
      };
    });
    assert.equal(boxes.actions.left, boxes.total.left, `actions not below totals at ${width}px`);
    assert.ok(boxes.actions.top >= boxes.total.bottom, `actions overlap totals at ${width}px`);
    assert.ok(boxes.suggestions.top >= boxes.total.bottom, `suggestions overlap totals at ${width}px`);
    assert.ok(Math.abs((boxes.actions.bottom - boxes.actions.top) - (boxes.suggestions.bottom - boxes.suggestions.top)) < 1,
      `actions and suggestions have different heights at ${width}px`);
    assert.ok(boxes.flow.top >= boxes.status.bottom, `flow overlaps projects at ${width}px`);
    assert.ok(boxes.todo.top >= boxes.flow.bottom, `to-do overlaps flow at ${width}px`);
    if (width > 1380){
      assert.equal(boxes.suggestions.top, boxes.actions.top);
      assert.equal(boxes.suggestions.right, boxes.total.right);
      assert.ok(boxes.suggestions.right - boxes.suggestions.left > boxes.actions.right - boxes.actions.left,
        'suggestions should span four columns, actions three');
      assert.ok(boxes.status.left > boxes.total.right, 'right rail must be to the right of totals');
      assert.ok(boxes.status.right - boxes.status.left < boxes.actions.right - boxes.actions.left,
        'right rail should span two columns');
      assert.equal(boxes.status.left, boxes.flow.left);
      assert.equal(boxes.flow.left, boxes.todo.left);
      assert.equal(boxes.status.top, boxes.total.top);
      assert.ok(Math.abs(boxes.actions.bottom - boxes.todo.bottom) < 1,
        `actions should end level with to-do at ${width}px`);
      assert.ok(Math.abs(boxes.suggestions.bottom - boxes.todo.bottom) < 1,
        `suggestions should end level with to-do at ${width}px`);
      assert.equal(boxes.copilot.left, boxes.todo.left);
      assert.ok(boxes.copilot.top >= boxes.todo.bottom);
      const todoHeader = await page.evaluate(()=>{
        const header = document.querySelector('.dashboard-todo-widget .dashboard-widget-header').getBoundingClientRect();
        const button = document.querySelector('.dashboard-todo-widget .dashboard-widget-header .btn').getBoundingClientRect();
        return { header, button };
      });
      assert.ok(todoHeader.button.right <= todoHeader.header.right, `to-do button clipped at ${width}px`);
      const exportHeader = await page.evaluate(()=>{
        const header = document.querySelector('.dashboard-total-widget .dashboard-widget-header').getBoundingClientRect();
        const button = document.querySelector('#dashboardTotalsExport').getBoundingClientRect();
        return { header, button };
      });
      assert.ok(exportHeader.button.right <= exportHeader.header.right, `export button clipped at ${width}px`);
    } else if (width > 760){
      assert.equal(boxes.suggestions.top, boxes.actions.top);
      assert.ok(Math.abs((boxes.actions.right - boxes.actions.left) - (boxes.suggestions.right - boxes.suggestions.left)) < 1,
        `tablet cards should have equal widths at ${width}px`);
      assert.ok(boxes.status.top >= boxes.actions.bottom, `status stack must follow actions at ${width}px`);
    } else {
      assert.ok(boxes.suggestions.top >= boxes.actions.bottom, `suggestions must follow actions at ${width}px`);
      assert.ok(boxes.status.top >= boxes.suggestions.bottom, `status stack must follow suggestions at ${width}px`);
    }
    if (width <= 760){
      const mobileContent = await page.evaluate(()=>{
        const bounds = selector=>{
          const rect = document.querySelector(selector).getBoundingClientRect();
          return { top: rect.top, bottom: rect.bottom, right: rect.right };
        };
        return {
          hero: bounds('.dashboard-total-hero-stack'),
          sections: bounds('.dashboard-total-sections'),
          total: bounds('.dashboard-total-widget'),
          todoButton: bounds('.dashboard-todo-widget .dashboard-widget-header .btn'),
          todo: bounds('.dashboard-todo-widget')
        };
      });
      assert.ok(mobileContent.sections.top >= mobileContent.hero.bottom, 'total sections overlap hero on mobile');
      assert.ok(mobileContent.total.bottom >= mobileContent.sections.bottom, 'total sections are clipped on mobile');
      assert.ok(mobileContent.todoButton.bottom <= mobileContent.todo.bottom, 'to-do button is clipped on mobile');
    }
    if (process.env.SAVE_DASHBOARD_SCREENSHOTS === '1'){
      await page.screenshot({ path: join(tmpdir(), `busbar-dashboard-layout-${width}.png`), fullPage: true });
    }
    await page.close();
  }
});
