// Mobile navigation is separate from the persisted desktop sidebar preference.
const media = window.matchMedia('(max-width: 760px)');
const shell = document.getElementById('dashboardShell');
const sidebar = shell?.querySelector('.dashboard-sidebar');
const toggle = document.getElementById('dashboardSidebarToggle');
const nav = sidebar?.querySelector('.sidebar-nav');

if (sidebar && toggle && nav){
  const header = document.createElement('header');
  header.className = 'mobile-app-header';
  header.setAttribute('aria-label', 'Busbar og gjeldende side');
  const brand = document.createElement('a');
  brand.className = 'mobile-header-brand';
  brand.href = sidebar.querySelector('.sidebar-brand')?.href || 'https://busbar.no';
  brand.setAttribute('aria-label', 'Gå til busbar.no');
  const logo = document.createElement('img');
  logo.src = new URL('../../assets/Busbar_original_hvit.png', import.meta.url).href;
  logo.alt = 'Busbar';
  brand.append(logo);
  const page = document.createElement('div');
  page.className = 'mobile-header-page';
  const title = document.createElement('span');
  title.className = 'mobile-page-title';
  page.append(title);
  header.append(page, brand);
  document.body.prepend(header);
  const togglePosition = document.createComment('Desktop sidebar toggle');
  toggle.before(togglePosition);
  const backdrop = document.createElement('button');
  backdrop.type = 'button';
  backdrop.className = 'mobile-menu-backdrop';
  backdrop.setAttribute('aria-label', 'Lukk meny');
  backdrop.tabIndex = -1;
  backdrop.setAttribute('aria-hidden', 'true');
  document.body.append(backdrop);
  nav.id ||= 'mobileNavigation';
  const mobileOpenKey = 'busbar.mobile.sidebar.open.v1';
  let open = false;
  try { open = sessionStorage.getItem(mobileOpenKey) === '1'; } catch {}

  function syncTitle(){
    const active = nav.querySelector('.is-active');
    title.textContent = active?.querySelector('.sidebar-nav-label')?.textContent
      || (document.getElementById('calcBtn') ? 'Prisverktøy' : 'Dashbord');
  }
  function sync(){
    if (media.matches && toggle.parentElement !== header) header.prepend(toggle);
    else if (!media.matches && toggle.parentElement !== togglePosition.parentElement) togglePosition.after(toggle);
    const mobileOpen = media.matches && open;
    document.body.classList.toggle('is-mobile-menu-open', mobileOpen);
    backdrop.inert = !mobileOpen;
    nav.inert = media.matches && !open;
    if (media.matches){
      toggle.setAttribute('aria-controls', nav.id);
      toggle.setAttribute('aria-expanded', String(open));
      toggle.setAttribute('aria-label', open ? 'Minimer meny' : 'Utvid meny');
      toggle.title = open ? 'Minimer meny' : 'Utvid meny';
    }
    syncTitle();
  }
  function close(restoreFocus = false){
    open = false;
    try { sessionStorage.setItem(mobileOpenKey, '0'); } catch {}
    sync();
    if (restoreFocus) toggle.focus();
  }
  // Capture prevents a mobile click from changing the saved desktop width.
  toggle.addEventListener('click', event=>{
    if (!media.matches) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    open = !open;
    try { sessionStorage.setItem(mobileOpenKey, open ? '1' : '0'); } catch {}
    sync();
  }, true);
  backdrop.addEventListener('click', ()=>close(true));
  nav.addEventListener('click', event=>{
    // Page selection keeps the sidebar expanded, as on desktop.
    if (media.matches && event.target.closest('#loginBtn, #logoutBtn')) close();
  });
  document.addEventListener('keydown', event=>{
    if (!media.matches || !open) return;
    if (event.key === 'Escape'){
      event.preventDefault();
      close(true);
    } else if (event.key === 'Tab'){
      const items = [toggle, brand, ...nav.querySelectorAll('a[href], button:not(:disabled)')]
        .filter(element=>element.getClientRects().length && !element.closest('[hidden]'));
      const first = items[0];
      const last = items.at(-1);
      if (event.shiftKey && document.activeElement === first){
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last){
        event.preventDefault();
        first.focus();
      }
    }
  });
  media.addEventListener('change', ()=>{
    if (!media.matches){
      const collapsed = shell.classList.contains('is-sidebar-collapsed');
      toggle.removeAttribute('aria-controls');
      toggle.setAttribute('aria-label', collapsed ? 'Utvid meny' : 'Minimer meny');
      toggle.setAttribute('aria-expanded', String(!collapsed));
      toggle.title = toggle.getAttribute('aria-label');
    }
    sync();
  });
  new MutationObserver(syncTitle).observe(nav, { subtree: true, attributes: true, attributeFilter: ['class'] });
  new MutationObserver(()=>{
    if (media.matches && (toggle.getAttribute('aria-expanded') !== String(open)
      || toggle.getAttribute('aria-label') !== (open ? 'Minimer meny' : 'Utvid meny'))) sync();
  }).observe(toggle, { attributes: true, attributeFilter: ['aria-expanded', 'aria-label'] });
  sync();
  // The application initializes its desktop sidebar after module loading.
  window.addEventListener('load', sync, { once: true });
}

// Keep technical calculation sections available without a very long mobile form.
document.querySelectorAll('#calculatorView .montasje-box').forEach(fieldset=>{
  const details = document.createElement('details');
  details.className = 'mobile-calculator-section';
  const summary = document.createElement('summary');
  summary.textContent = fieldset.querySelector('legend')?.textContent || 'Detaljer';
  fieldset.before(details);
  details.append(summary, fieldset);
  const sync = ()=>{ details.open = !media.matches; };
  media.addEventListener('change', sync);
  sync();
});

const resultActions = document.querySelector('#results .result-actions');
if (resultActions){
  const originalPosition = document.createComment('Result actions on desktop');
  resultActions.before(originalPosition);
  const sync = ()=>{
    if (media.matches) document.querySelector('#results .results-heading').after(resultActions);
    else originalPosition.after(resultActions);
  };
  media.addEventListener('change', sync);
  sync();
}

// Label each BOM cell so the same table can be presented as cards on a phone.
document.querySelectorAll('#bomTbl, #xapBomTbl').forEach(table=>{
  const headings = [...table.querySelectorAll('thead th')].map(cell=>cell.textContent);
  const labelCells = ()=>{
    table.querySelectorAll('tbody tr').forEach(row=>{
      [...row.cells].forEach((cell, index)=>{ cell.dataset.label = headings[index] || ''; });
    });
  };
  new MutationObserver(labelCells).observe(table.tBodies[0], { childList: true, subtree: true });
  labelCells();
});
