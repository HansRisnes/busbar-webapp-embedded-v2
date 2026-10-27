import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

test('project search renders only matches and defers hidden lines', async t=>{
  const apiProjects = Array.from({ length: 68 }, (_, index)=>({
    id: `project-${index}`,
    name: index < 3 ? `Kongsberg ${index}` : `Other ${index}`,
    customer: 'Test',
    projectOwnerEmail: 'search@busbar.no',
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z',
    lines: Array.from({ length: 15 }, (_, line)=>({ id: `line-${index}-${line}`, lineNumber: `L-${line}` }))
  }));
  const server = createServer(async (request, response)=>{
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname.startsWith('/api/')){
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify(pathname.startsWith('/api/user-projects')
        ? { projects: apiProjects, ownerEmails: ['search@busbar.no'], updatedAt: '2026-01-01T00:00:00.000Z' }
        : { customers: [] }));
      return;
    }
    if (pathname === '/'){
      response.setHeader('Content-Type', 'text/html; charset=utf-8');
      response.end('<!doctype html><html><body><div id="projectList"></div></body></html>');
      return;
    }
    const file = resolve(root, `.${pathname}`);
    if (!file.startsWith(`${root}${sep}`)){
      response.statusCode = 403;
      response.end();
      return;
    }
    try{
      const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
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
  await page.goto(`http://127.0.0.1:${server.address().port}/`);
  const result = await page.evaluate(async()=>{
    const { renderProjectsPage } = await import('/src/client/projects.js');
    const projects = Array.from({ length: 68 }, (_, index)=>({
      id: `project-${index}`,
      name: index < 3 ? `Kongsberg ${index}` : `Other ${index}`,
      customer: 'Test',
      lines: Array.from({ length: 15 }, (_, line)=>({ id: `line-${index}-${line}`, lineNumber: `L-${line}` }))
    }));
    const state = { projects, projectSearchTerm: 'kong', expandedProjectId: null, showArchive: false };
    const calls = { dashboard: 0, flowState: 0, archive: 0, folders: 0, offers: 0, flow: 0 };
    const callbacks = {
      renderMainDashboard: ()=>calls.dashboard++,
      loadProjectFlowState: ()=>calls.flowState++,
      updateProjectArchiveUi: ()=>calls.archive++,
      ensureProjectFolderStatusesLoaded: ()=>calls.folders++,
      renderOffersList: ()=>calls.offers++,
      renderProjectFlowView: ()=>calls.flow++,
      projectMatchesSearch: project=>project.name.toLowerCase().includes(state.projectSearchTerm)
    };
    const search = ()=>renderProjectsPage({ projectState: state, callbacks, searchOnly: true });
    const first = search();
    const collapsed = {
      count: document.querySelectorAll('.project-row').length,
      hiddenDetails: document.querySelectorAll('.project-detail[hidden]').length,
      lineItems: document.querySelectorAll('.project-line-item').length,
      toggleText: document.querySelector('.project-detail-toggle-btn')?.textContent
    };
    state.expandedProjectId = projects[0].id;
    search();
    const expanded = {
      lineItems: document.querySelectorAll('.project-line-item').length,
      visibleDetails: document.querySelectorAll('.project-detail:not([hidden])').length,
      toggleText: document.querySelector('.project-detail-toggle-btn')?.textContent
    };
    const searchCalls = { ...calls };
    renderProjectsPage({ projectState: state, callbacks });
    const fullRenderCalls = { ...calls };
    state.expandedProjectId = null;
    state.projectSearchTerm = '';
    const allProjects = search();
    return {
      visibleCount: first.visibleCount,
      collapsed,
      expanded,
      searchCalls,
      fullRenderCalls,
      allProjects: {
        visibleCount: allProjects.visibleCount,
        rows: document.querySelectorAll('.project-row').length,
        lineItems: document.querySelectorAll('.project-line-item').length
      }
    };
  });

  assert.equal(result.visibleCount, 3);
  assert.deepEqual(result.collapsed, { count: 3, hiddenDetails: 3, lineItems: 0, toggleText: 'Vis linjer' });
  assert.deepEqual(result.expanded, { lineItems: 15, visibleDetails: 1, toggleText: 'Skjul linjer' });
  assert.deepEqual(result.searchCalls, { dashboard: 0, flowState: 0, archive: 0, folders: 0, offers: 0, flow: 0 });
  assert.deepEqual(result.fullRenderCalls, { dashboard: 1, flowState: 1, archive: 1, folders: 1, offers: 1, flow: 1 });
  assert.deepEqual(result.allProjects, { visibleCount: 68, rows: 68, lineItems: 0 });

  const appPage = await browser.newPage();
  await appPage.addInitScript(()=>{
    sessionStorage.setItem('busbar.auth.session.v1', JSON.stringify({
      loggedIn: true,
      username: 'search@busbar.no',
      token: 'test-token'
    }));
  });
  await appPage.goto(`http://127.0.0.1:${server.address().port}/index.html?view=projects`);
  await appPage.waitForFunction(()=>document.querySelectorAll('#projectList .project-row').length === 68);
  assert.equal(await appPage.locator('#projectList .project-line-item').count(), 0);
  await appPage.locator('#projectSearchInput').pressSequentially('kong');
  await appPage.waitForFunction(()=>document.querySelectorAll('#projectList .project-row').length === 3);
  await appPage.locator('#projectList .project-detail-toggle-btn').first().click();
  assert.equal(await appPage.locator('#projectList .project-line-item').count(), 15);
  await appPage.close();
});
