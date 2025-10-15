// v2 + v2.1, XCP-S/XCM, distribusjon/feeder, avtappingsbokser, ekspansjon-modal >30 m

const RAW_CSV_PATHS = [
  'data/busbar-webapp-embedded-v2.csv',
  'data/busbar-webapp-embedded-v2.1.csv'
];
let lastCalc = null; // delsummer for live frakt-oppdatering
const AUTH_PASSWORD = 'busbar';
let authState = { loggedIn: false, username: '' };
const projectState = {
  currentProject: '',
  currentCustomer: '',
  projectHistory: [],
  customerHistory: []
};
let projectModalResolve = null;

// --- CSV ---
function parseCSVAuto(text){
  const lines = text.replace(/\r/g,'').split('\n').filter(x=>x.length);
  if (!lines.length) return [];
  const sep = (lines[0].split(';').length > lines[0].split(',').length) ? ';' : ',';
  const parseLine = (line)=>{
    const out=[]; let f='', q=false;
    for(let i=0;i<line.length;i++){
      const c=line[i];
      if(c === '"'){ if(q && line[i+1] === '"'){ f+='"'; i++; } else { q=!q; } continue; }
      if(!q && c === sep){ out.push(f); f=''; continue; }
      f += c;
    }
    out.push(f); return out;
  };
  const header = parseLine(lines[0]).map(h=>h.trim());
  return lines.slice(1)
    .map(parseLine)
    .filter(r=>r.some(x=>x && x.trim()!==''))
    .map(cols=>{
      const obj = Object.fromEntries(header.map((h,i)=>[h, cols[i]??'']));
      obj._cols = cols;
      return obj;
    });
}

// --- utils ---
const $ = id=>document.getElementById(id);
const round2 = n=>Math.round(n*100)/100;
const fmtNO = new Intl.NumberFormat('no-NO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtIntNO = new Intl.NumberFormat('no-NO', { maximumFractionDigits: 0 });
const toNum = x => {
  if (x===undefined || x===null) return NaN;
  const v = Number(String(x).replace(/\s/g,'').replace(',','.'));
  return Number.isFinite(v) ? v : NaN;
};
function pick(row, names){ for (const n of names){ if (n in row && row[n]!=='' && row[n]!==undefined) return row[n]; } return ''; }

function updateAuthUI(){
  const calcBtn = $('calcBtn');
  if (calcBtn) calcBtn.disabled = !authState.loggedIn;

  const loginBtn = $('loginBtn');
  const logoutBtn = $('logoutBtn');
  const userLabel = $('authUser');

  if (loginBtn) loginBtn.hidden = authState.loggedIn;
  if (logoutBtn) logoutBtn.hidden = !authState.loggedIn;
  if (userLabel){
    if (authState.loggedIn){
      userLabel.textContent = authState.username || 'Innlogget';
      userLabel.hidden = false;
    } else {
      userLabel.textContent = '';
      userLabel.hidden = true;
    }
  }

  const statusEl = $('status');
  if (statusEl){
    if (!authState.loggedIn){
      if (!statusEl.textContent){
        statusEl.textContent = 'Logg inn for \u00E5 beregne.';
      }
    } else if (statusEl.textContent === 'Logg inn for \u00E5 beregne.'){
      statusEl.textContent = '';
    }
  }
}

function showLoginModal(){
  const modal = $('loginModal');
  if (!modal) return;
  modal.style.display = 'flex';
  const usernameInput = $('loginUsername');
  const passwordInput = $('loginPassword');
  const errorEl = $('loginError');
  if (errorEl) errorEl.textContent = '';
  if (usernameInput){
    usernameInput.value = authState.username || '';
    try{
      const len = usernameInput.value.length;
      usernameInput.focus();
      usernameInput.setSelectionRange(len, len);
    }catch(_err){
      usernameInput.focus();
    }
  }
  if (passwordInput){
    passwordInput.value = '';
  }
}

function hideLoginModal(){
  const modal = $('loginModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('loginError');
  if (errorEl) errorEl.textContent = '';
}

function handleLoginSubmit(){
  const passwordInput = $('loginPassword');
  const usernameInput = $('loginUsername');
  const errorEl = $('loginError');
  const password = passwordInput ? passwordInput.value : '';
  if (password !== AUTH_PASSWORD){
    if (errorEl) errorEl.textContent = 'Feil passord.';
    if (passwordInput) passwordInput.focus();
    return;
  }
  const username = usernameInput ? usernameInput.value.trim() : '';
  if (!username || !username.includes('@')){
    if (errorEl) errorEl.textContent = 'Brukernavn må være en e-postadresse.';
    if (usernameInput){
      usernameInput.focus();
      try{
        const len = usernameInput.value.length;
        usernameInput.setSelectionRange(len, len);
      }catch(_err){}
    }
    return;
  }
  authState = { loggedIn: true, username };
  hideLoginModal();
  updateAuthUI();
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = '';
}

const loginBtn = $('loginBtn');
if (loginBtn){
  loginBtn.addEventListener('click', showLoginModal);
}
const logoutBtn = $('logoutBtn');
if (logoutBtn){
  logoutBtn.addEventListener('click', ()=>{
    authState = { loggedIn: false, username: '' };
    hideLoginModal();
   updateAuthUI();
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Logg inn for \u00E5 beregne.';
  });
}
const loginCancel = $('loginCancel');
if (loginCancel){
  loginCancel.addEventListener('click', hideLoginModal);
}
const loginSubmit = $('loginSubmit');
if (loginSubmit){
  loginSubmit.addEventListener('click', handleLoginSubmit);
}
['loginUsername','loginPassword'].forEach(id=>{
  const input = $(id);
  if (input){
    input.addEventListener('keydown', evt=>{
      if (evt.key === 'Enter'){
        evt.preventDefault();
        handleLoginSubmit();
      } else if (evt.key === 'Escape'){
        evt.preventDefault();
        hideLoginModal();
      }
    });
  }
});
const loginModal = $('loginModal');
if (loginModal){
  loginModal.addEventListener('click', evt=>{
    if (evt.target === loginModal){
      hideLoginModal();
    }
  });
}
function addToHistory(list, value){
  const trimmed = String(value||'').trim();
  if (!trimmed) return;
  const lower = trimmed.toLowerCase();
  const existingIdx = list.findIndex(entry=>entry.toLowerCase() === lower);
  if (existingIdx !== -1){
    list.splice(existingIdx,1);
  }
  list.push(trimmed);
  if (list.length > 20){
    list.splice(0, list.length - 20);
  }
}

function updateProjectMetaDisplay(){
  const meta = $('projectMeta');
  const nameEl = $('projectNameDisplay');
  const customerEl = $('customerNameDisplay');
  const editBtn = $('editProjectBtn');
  const hasData = Boolean(projectState.currentProject && projectState.currentCustomer);
  if (nameEl) nameEl.textContent = projectState.currentProject;
  if (customerEl) customerEl.textContent = projectState.currentCustomer;
  if (meta) meta.hidden = !hasData;
  if (editBtn){
    editBtn.hidden = !hasData;
    editBtn.disabled = !hasData;
  }
}

function hideSuggestions(listEl){
  if (listEl){
    listEl.hidden = true;
    listEl.innerHTML = '';
  }
}

function showSuggestions(listEl, items){
  if (!listEl) return;
  const entries = items.filter(Boolean);
  if (!entries.length){
    hideSuggestions(listEl);
    return;
  }
  const frag = document.createDocumentFragment();
  [...entries].reverse().forEach(value=>{
    const li = document.createElement('li');
    li.textContent = value;
    li.dataset.value = value;
    frag.appendChild(li);
  });
  listEl.innerHTML = '';
  listEl.appendChild(frag);
  listEl.hidden = false;
}

function updateProjectSubmitState(){
  const submit = $('projectSubmit');
  if (!submit) return;
  const projectVal = (($('projectNameInput')?.value) || '').trim();
  const customerVal = (($('customerNameInput')?.value) || '').trim();
  submit.disabled = !(projectVal && customerVal);
}

function openProjectModal(){
  const modal = $('projectModal');
  if (!modal) return;
  modal.style.display = 'flex';
  const errorEl = $('projectError');
  if (errorEl) errorEl.textContent = '';
  const projectInput = $('projectNameInput');
  const customerInput = $('customerNameInput');
  if (projectInput){
    projectInput.value = projectState.currentProject || '';
    projectInput.focus();
    const len = projectInput.value.length;
    try{
      projectInput.setSelectionRange(len, len);
    }catch(_err){
      /* ignore selection errors */
    }
  }
  if (customerInput){
    customerInput.value = projectState.currentCustomer || '';
  }
  updateProjectSubmitState();
}

function closeProjectModal(){
  const modal = $('projectModal');
  if (!modal) return;
  modal.style.display = 'none';
  const errorEl = $('projectError');
  if (errorEl) errorEl.textContent = '';
  hideSuggestions($('projectSuggestions'));
  hideSuggestions($('customerSuggestions'));
}

function persistProjectInfo(projectName, customerName){
  projectState.currentProject = projectName;
  projectState.currentCustomer = customerName;
  addToHistory(projectState.projectHistory, projectName);
  addToHistory(projectState.customerHistory, customerName);
  updateProjectMetaDisplay();
  const statusEl = $('status');
  if (statusEl && statusEl.textContent === 'Oppgi prosjektnavn og kunde.'){
    statusEl.textContent = '';
  }
}

function resolveProjectModal(result){
  if (projectModalResolve){
    const resolve = projectModalResolve;
    projectModalResolve = null;
    resolve(result);
  }
}

function submitProjectModal(){
  const projectInput = $('projectNameInput');
  const customerInput = $('customerNameInput');
  const errorEl = $('projectError');
  const projectName = projectInput ? projectInput.value.trim() : '';
  const customerName = customerInput ? customerInput.value.trim() : '';
  if (!projectName || !customerName){
    if (errorEl) errorEl.textContent = 'Fyll ut begge feltene.';
    updateProjectSubmitState();
    return;
  }
  persistProjectInfo(projectName, customerName);
  closeProjectModal();
  resolveProjectModal({ projectName, customerName });
}

function cancelProjectModal(){
  closeProjectModal();
  if (!projectState.currentProject || !projectState.currentCustomer){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Oppgi prosjektnavn og kunde.';
  }
  resolveProjectModal(null);
}

function ensureProjectInfo(){
  if (projectState.currentProject && projectState.currentCustomer){
    return Promise.resolve({
      projectName: projectState.currentProject,
      customerName: projectState.currentCustomer
    });
  }
  const statusEl = $('status');
  if (statusEl) statusEl.textContent = 'Oppgi prosjektnavn og kunde.';
  return new Promise(resolve=>{
    projectModalResolve = resolve;
    openProjectModal();
  });
}

const projectSubmit = $('projectSubmit');
if (projectSubmit){
  projectSubmit.addEventListener('click', submitProjectModal);
}
const projectCancel = $('projectCancel');
if (projectCancel){
  projectCancel.addEventListener('click', cancelProjectModal);
}
const projectModal = $('projectModal');
if (projectModal){
  projectModal.addEventListener('click', evt=>{
    if (evt.target === projectModal){
      cancelProjectModal();
    }
  });
}

function bindSuggestionBehaviour(inputId, listId, source){
  const input = $(inputId);
  const listEl = $(listId);
  if (!input || !listEl) return;
  input.addEventListener('focus', ()=>{
    showSuggestions(listEl, source);
  });
  input.addEventListener('input', ()=>{
    hideSuggestions(listEl);
    updateProjectSubmitState();
    const errorEl = $('projectError');
    if (errorEl) errorEl.textContent = '';
  });
  input.addEventListener('blur', ()=>{
    setTimeout(()=>hideSuggestions(listEl), 80);
  });
  input.addEventListener('keydown', evt=>{
    if (evt.key === 'Enter'){
      evt.preventDefault();
      const submitBtn = $('projectSubmit');
      if (submitBtn && !submitBtn.disabled){
        submitProjectModal();
      }
    } else if (evt.key === 'Escape'){
      evt.preventDefault();
      cancelProjectModal();
    }
  });
  listEl.addEventListener('mousedown', evt=>{
    if (evt.target && evt.target.matches('li[data-value]')){
      evt.preventDefault();
      const value = evt.target.dataset.value || evt.target.textContent || '';
      input.value = value;
      hideSuggestions(listEl);
      updateProjectSubmitState();
    }
  });
}

bindSuggestionBehaviour('projectNameInput', 'projectSuggestions', projectState.projectHistory);
bindSuggestionBehaviour('customerNameInput', 'customerSuggestions', projectState.customerHistory);

const editProjectBtn = $('editProjectBtn');
if (editProjectBtn){
  editProjectBtn.addEventListener('click', ()=>{
    projectModalResolve = null;
    openProjectModal();
  });
}

document.addEventListener('keydown', evt=>{
  if (evt.key === 'Escape'){
    const loginModalEl = $('loginModal');
    if (loginModalEl && loginModalEl.style.display === 'flex'){
      hideLoginModal();
      return;
    }
    const projectModalEl = $('projectModal');
    if (projectModalEl && projectModalEl.style.display === 'flex'){
      cancelProjectModal();
    }
  }
});

updateProjectMetaDisplay();
updateAuthUI();

const H = {
  code: ['code','Code','SKU','sku','produkt','Produkt'],
  price: ['price','Price','unit price','unit_price','Unit Price','UnitPrice','pris','Pris'],
  desc:  ['desc_text','description','Description','desc','tekst','Tekst'],
  amp:   ['ampere','Ampere','amp','Amp'],
  et2:   ['element_type_2','Element type 2','element type 2','H']
};

const DEFAULT_HOURLY_RATE = 700;
const OPPHENG_RATE_TABLE = Object.freeze([
  { maxAmp: 1600, rate: 400, label: '160–1600A' },
  { maxAmp: 2500, rate: 500, label: '2000–2500A' },
  { maxAmp: Infinity, rate: 800, label: '3200–5000A' }
]);
const MONTASJE_TIME_TABLE = Object.freeze([
  { maxAmp: 250,   hoursPerMeter: 1.5, hoursPerAngle: 0.5,  label: '160–250A' },
  { maxAmp: 630,   hoursPerMeter: 2,   hoursPerAngle: 0.5,  label: '400–630A' },
  { maxAmp: 1600,  hoursPerMeter: 3,   hoursPerAngle: 0.75, label: '800–1600A' },
  { maxAmp: 2500,  hoursPerMeter: 4,   hoursPerAngle: 1,    label: '2000–2500A' },
  { maxAmp: Infinity, hoursPerMeter: 5, hoursPerAngle: 1.5, label: '3200–5000A' }
]);

function sanitizeHourlyRate(value){
  const raw = value ?? '';
  if (String(raw).trim()==='') return DEFAULT_HOURLY_RATE;
  const n = toNum(raw);
  if (!Number.isFinite(n)) return DEFAULT_HOURLY_RATE;
  return Math.max(0, n);
}

function getOpphengRateRowForAmp(amp){
  const a = Number(amp);
  if (!Number.isFinite(a) || a <= 0) return null;
  for (const row of OPPHENG_RATE_TABLE){
    if (a <= row.maxAmp) return row;
  }
  return OPPHENG_RATE_TABLE[OPPHENG_RATE_TABLE.length - 1];
}

function sanitizeOpphengRate(value, fallback){
  const raw = value ?? '';
  const fallbackRate = Number.isFinite(fallback) ? fallback : 0;
  if (String(raw).trim()==='') return fallbackRate;
  const n = toNum(raw);
  if (!Number.isFinite(n)) return fallbackRate;
  return Math.max(0, n);
}

function calculateOpphengsmateriell({ meter, amp, ratePerMeter }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const ampValue = Number(amp);
  const profile = getOpphengRateRowForAmp(ampValue);
  const defaultRate = profile ? profile.rate : 0;
  const rate = sanitizeOpphengRate(ratePerMeter, defaultRate);
  const cost = round2(totalMeters * rate);
  return {
    cost,
    meters: totalMeters,
    ratePerMeter: rate,
    defaultRate,
    profile,
    isDefaultRate: rate === defaultRate,
    amp: Number.isFinite(ampValue) ? ampValue : null
  };
}

function getMontasjeProfileForAmp(amp){
  const a = Number(amp);
  if (!Number.isFinite(a) || a <= 0) return null;
  for (const row of MONTASJE_TIME_TABLE){
    if (a <= row.maxAmp) return row;
  }
  return MONTASJE_TIME_TABLE[MONTASJE_TIME_TABLE.length - 1];
}

function calculateMontasje({ meter, angles, amp, hourlyRate }){
  const totalMeters = Math.max(0, Math.ceil(Number(meter) || 0));
  const totalAngles = Math.max(0, Math.round(Number(angles) || 0));
  const rate = sanitizeHourlyRate(hourlyRate);
  const ampValue = Number(amp);
  const profile = getMontasjeProfileForAmp(ampValue);
  if (!profile){
    return {
      cost: 0,
      meters: totalMeters,
      angles: totalAngles,
      hourlyRate: rate,
      totalHours: 0,
      profile: null,
      amp: Number.isFinite(ampValue) ? ampValue : null
    };
  }
  const hours = round2(totalMeters * profile.hoursPerMeter + totalAngles * profile.hoursPerAngle);
  const cost = round2(hours * rate);
  return {
    cost,
    meters: totalMeters,
    angles: totalAngles,
    hourlyRate: rate,
    totalHours: hours,
    profile,
    hoursPerMeter: profile.hoursPerMeter,
    hoursPerAngle: profile.hoursPerAngle,
    amp: Number.isFinite(ampValue) ? ampValue : null
  };
}

function readMontasjeSettingsFromUI(){
  const hourlyRateInput = document.getElementById('montasjeHourlyRate');
  const hourlyRate = sanitizeHourlyRate(hourlyRateInput ? hourlyRateInput.value : DEFAULT_HOURLY_RATE);
  if (hourlyRateInput){
    hourlyRateInput.value = String(hourlyRate);
  }
  const opphengRateInput = document.getElementById('opphengRate');
  let opphengRate = '';
  if (opphengRateInput){
    const defaultRate = toNum(opphengRateInput.dataset.defaultRate);
    const fallbackRate = Number.isFinite(defaultRate) ? defaultRate : 0;
    const raw = opphengRateInput.value;
    const sanitized = sanitizeOpphengRate(raw, fallbackRate);
    const isOverride = sanitized !== fallbackRate;
    opphengRateInput.dataset.appliedValue = String(sanitized);
    opphengRateInput.dataset.userOverride = isOverride ? 'true' : 'false';
    if (fallbackRate === 0 && !isOverride && raw.trim()===''){
      opphengRateInput.value = '';
    }else{
      opphengRateInput.value = (sanitized || sanitized === 0) ? String(sanitized) : '';
    }
    opphengRate = sanitized;
  }
  return { hourlyRate, opphengRate };
}

function formatMontasjeDetail(m){
  if (!m || !m.profile){
    return 'Montasje kalkuleres automatisk når ampere er valgt.';
  }
  const parts = [];
  parts.push(`${fmtIntNO.format(m.meters)} m × ${fmtNO.format(m.profile.hoursPerMeter)} t/m`);
  if (m.angles){
    parts.push(`${fmtIntNO.format(m.angles)} × ${fmtNO.format(m.profile.hoursPerAngle)} t`);
  }
  const basis = parts.join(' + ');
  const hoursTxt = fmtNO.format(m.totalHours);
  const rateTxt = fmtNO.format(m.hourlyRate);
  const ampTxt = Number.isFinite(m.amp) ? `${fmtIntNO.format(m.amp)}A` : '';
  const labelTxt = ampTxt ? `${ampTxt} (${m.profile.label})` : m.profile.label;
  return `Montasjegrunnlag (${labelTxt}): ${basis} = ${hoursTxt} t × ${rateTxt} kr/t`;
}

function formatOpphengDetail(o){
  if (!o){
    return 'Opphengsmateriell kalkuleres automatisk når ampere er valgt.';
  }
  if (!Number.isFinite(o.amp) || o.amp <= 0){
    return 'Opphengsmateriell kalkuleres automatisk når ampere er valgt.';
  }
  if (!o.meters){
    return 'Opphengsmateriell beregnes når meter er angitt.';
  }
  const metersTxt = fmtIntNO.format(o.meters);
  const rateTxt = fmtNO.format(o.ratePerMeter);
  const costTxt = fmtNO.format(o.cost);
  const labelTxt = o.profile ? ` (${o.profile.label})` : '';
  return `Opphengsmateriell${labelTxt}: ${metersTxt} m × ${rateTxt} kr/m = ${costTxt} kr`;
}

function setInputLocked(input, locked){
  if (!input) return;
  input.readOnly = locked;
  input.classList.toggle('locked', locked);
  input.setAttribute('aria-readonly', locked ? 'true' : 'false');
  input.dataset.locked = locked ? 'true' : 'false';
}

function updateMontasjePreview(){
  const meterEl = $('meter');
  const v90hEl = $('v90h');
  const v90vEl = $('v90v');
  const ampEl = $('ampSelect');
  const rateEl = $('montasjeHourlyRate');
  const opphengRateEl = $('opphengRate');
  const rateToggle = $('rateToggle');

  const meter = meterEl ? Number(meterEl.value || 0) : 0;
  const angles = (v90hEl ? Number(v90hEl.value || 0) : 0) + (v90vEl ? Number(v90vEl.value || 0) : 0);
  const amp = ampEl ? Number(ampEl.value || 0) : NaN;
  const ratesUnlocked = rateToggle ? rateToggle.checked : true;
  const montasjeLocked = !ratesUnlocked;
  if (rateEl){
    setInputLocked(rateEl, montasjeLocked);
    if (montasjeLocked){
      rateEl.value = String(DEFAULT_HOURLY_RATE);
    }
  }
  const hourlyRate = rateEl ? rateEl.value : DEFAULT_HOURLY_RATE;

  const montasjePreview = calculateMontasje({ meter, angles, amp, hourlyRate });

  let opphengRateForCalc = 0;
  if (opphengRateEl){
    const opphengLocked = !ratesUnlocked;
    setInputLocked(opphengRateEl, opphengLocked);
    if (!opphengRateEl.dataset.userOverride){
      opphengRateEl.dataset.userOverride = 'false';
    }
    const hasAmp = Number.isFinite(amp) && amp > 0;
    const opphengProfile = getOpphengRateRowForAmp(amp);
    const defaultOpphengRate = hasAmp && opphengProfile ? opphengProfile.rate : 0;
    const ampKey = hasAmp ? String(amp) : '';
    const prevAmp = opphengRateEl.dataset.appliedAmp ?? '';
    const ampChanged = ampKey !== prevAmp;

    if (opphengLocked){
      opphengRateEl.dataset.userOverride = 'false';
    }

    if (ampChanged){
      opphengRateEl.dataset.appliedAmp = ampKey;
      if (opphengLocked || opphengRateEl.dataset.userOverride !== 'true'){
        if (hasAmp){
          opphengRateEl.value = defaultOpphengRate ? String(defaultOpphengRate) : '';
        }else{
          opphengRateEl.value = '';
        }
      }
    }

    if (opphengLocked){
      if (hasAmp){
        opphengRateEl.value = defaultOpphengRate ? String(defaultOpphengRate) : '';
      }else{
        opphengRateEl.value = '';
      }
    }

    opphengRateEl.placeholder = hasAmp ? (defaultOpphengRate ? String(defaultOpphengRate) : '') : '';

    const rawValue = opphengRateEl.value;
    const fallbackRate = hasAmp ? defaultOpphengRate : 0;
    const sanitizedRate = sanitizeOpphengRate(rawValue, fallbackRate);

    opphengRateEl.dataset.defaultRate = hasAmp ? String(defaultOpphengRate) : '';
    opphengRateEl.dataset.appliedValue = String(sanitizedRate);
    opphengRateEl.dataset.appliedAmp = ampKey;

    if (!hasAmp && opphengRateEl.dataset.userOverride !== 'true'){
      opphengRateEl.value = '';
    }
    if (!opphengLocked && ampChanged && opphengRateEl.dataset.userOverride !== 'true' && hasAmp){
      opphengRateEl.value = sanitizedRate || sanitizedRate === 0 ? String(sanitizedRate) : '';
    }

    opphengRateForCalc = sanitizedRate;
  }

  const opphengPreview = calculateOpphengsmateriell({ meter, amp, ratePerMeter: opphengRateForCalc });

  const labelEl = $('montasjeProfileLabel');
  const perMeterEl = $('montasjeHoursPerMeter');
  const perAngleEl = $('montasjeHoursPerAngle');
  const totalHoursEl = $('montasjeTotalHours');
  const costEl = $('montasjePreviewCost');
  const opphengRatePreviewEl = $('opphengPreviewRate');
  const opphengCostPreviewEl = $('opphengPreviewCost');
  const montasjeDetailEl = $('montasjeDetail');
  const opphengDetailEl = $('opphengDetail');

  const hasProfile = Boolean(montasjePreview.profile);

  if (labelEl){
    if (hasProfile){
      const ampTxt = Number.isFinite(montasjePreview.amp) ? `${fmtIntNO.format(montasjePreview.amp)}A` : '';
      labelEl.textContent = ampTxt ? `Strømskinne ${ampTxt} (${montasjePreview.profile.label})` : `Strømskinne ${montasjePreview.profile.label}`;
    }else{
      labelEl.textContent = 'Velg ampere for å hente montasjetider.';
    }
  }
  if (perMeterEl){
    perMeterEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.profile.hoursPerMeter)} t/m` : '–';
  }
  if (perAngleEl){
    perAngleEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.profile.hoursPerAngle)} t/vinkel` : '–';
  }
  if (totalHoursEl){
    totalHoursEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.totalHours)} t` : '–';
  }
  if (costEl){
    costEl.textContent = hasProfile ? `${fmtNO.format(montasjePreview.cost)} kr` : '–';
  }
  const hasOpphengAmp = Number.isFinite(opphengPreview.amp) && opphengPreview.amp > 0;
  if (opphengRatePreviewEl){
    opphengRatePreviewEl.textContent = hasOpphengAmp ? `${fmtNO.format(opphengPreview.ratePerMeter)} kr/m` : '–';
  }
  if (opphengCostPreviewEl){
    opphengCostPreviewEl.textContent = hasOpphengAmp ? `${fmtNO.format(opphengPreview.cost)} kr` : '–';
  }
  if (montasjeDetailEl){
    montasjeDetailEl.textContent = formatMontasjeDetail(montasjePreview);
  }
  if (opphengDetailEl){
    opphengDetailEl.textContent = formatOpphengDetail(opphengPreview);
  }
}

// --- detect ---
function detectSeries(row){
  const code = String(pick(row,H.code)).toUpperCase();
  const d = String(pick(row,H.desc)).toUpperCase();
  if (code.startsWith('XCM') || d.includes('XCM')) return 'XCM';
  if (code.startsWith('XCP') || d.includes('XCP')) return 'XCP-S';
  if (code.startsWith('XCA') || d.includes('XCA')) return 'XCP-S';
  return '';
}

function detectType(descRaw){
  const d0 = String(descRaw||'');
  const d  = d0.toLowerCase().replace(/\s+/g,' ');

  // feed
  if (d.includes('crt') && d.includes('feed')) return 'crt_board_feed';
  if ((d.includes('board-trans') || d.includes('board')) && d.includes('feed')) return 'board_feed';

  // XCM FEEDER (før alt generelt)
  if (/\bxcm\b/.test(d) && /\bfeeder\b/.test(d)){
    if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))        return 'xcm_feeder_3m';
    if (/l\s*=\s*1501\s*[-–]\s*2999\b/.test(d))        return 'xcm_feeder_1501_2999';
    if (/l\s*=\s*600\s*[-–]\s*1500\b/.test(d))         return 'xcm_feeder_600_1500';
  }

  // XCM DISTRIBUSJON
  if (/\bxcm\b/.test(d) && /straight\s*length/.test(d)){
    if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))        return 'straight_3m_dist';
    if (/l\s*=\s*1500\s*[-–]\s*2999\b/.test(d))        return 'xcm_dist_1500_2999';
    if (/l\s*=\s*1000\s*[-–]\s*1500\b/.test(d))        return 'xcm_dist_1000_1500';
  }

  // Generelt / XCP-S
  const isDistGen = /straight\s*length|\boutl(ets?)?\b/.test(d);
  if (/l\s*=\s*3\s*m\b|l\s*=\s*3m\b/.test(d))          return isDistGen ? 'straight_3m_dist' : 'straight_3m';
  if (/l\s*=\s*1501\s*[-–]\s*2000\b/.test(d))          return isDistGen ? 'straight_1501_2000_dist' : 'straight_1501_2000';
  if (/l\s*=\s*500\s*[-–]\s*1000\b/.test(d))           return isDistGen ? 'straight_500_1000_dist'  : 'straight_500_1000';

  // andre typer
  if (d.includes('horizontal') && (d.includes('elbow') || d.includes('90'))) return 'elbow_horizontal_90';
  if (d.includes('vertical')   && (d.includes('elbow') || d.includes('90'))) return 'elbow_vertical_90';
  if (/tap[\s-]*off\s*box/.test(d))  return 'tap_off_box';
  if (/plug[\s-]*in\s*box/.test(d))  return 'plug_in_box';
  if (/bolt[\s-]*on\s*box|b160\s*bolt/.test(d)) return 'bolt_on_box';
  if (/end\s*cover/.test(d))        return 'end_cover';
  if (/end\s*feed\s*unit/.test(d))  return 'end_feed_unit';
  if (/expansion/.test(d))          return 'expansion_unit';

  const isFire = /fire[\s-]*barrier|firebarrier|fire[\s-]*stop|firestop|brannbarrier|brannbarriere|brannelement|brann/.test(d);
  if (isFire){
    const ext = /(external|utvendig|ytter)/i.test(d);
    const int = /(internal|innvendig|inner)/i.test(d);
    if (ext && !int) return 'fire_barrier_kit_external';
    if (int && !ext) return 'fire_barrier_kit_internal';
    return 'fire_barrier_kit';
  }
  return '';
}

// amp
function extractAmpGeneric(row){
  let src = (pick(row, H.desc) + ' ' + pick(row, H.code)).toUpperCase();
  let m = src.match(/(\d{2,4})\s*A\b/);
  if (m) return Number(m[1]);
  m = src.match(/\b(\d{2,4})\b/);
  if (m) return Number(m[1]);
  return NaN;
}
function deriveAmp(row){
  if (Number.isFinite(row.ampere)) return Number(row.ampere);
  const s = String((row._desc||'')+' '+(row.code||'')).toUpperCase();
  let m = s.match(/(\d{2,4})\s*A\b/); if (m) return Number(m[1]);
  m = s.match(/\b(\d{2,4})\b/);       if (m) return Number(m[1]);
  return NaN;
}

// rå → katalog
function adaptRawToCatalog(rawRows){
  return rawRows.map(r=>{
    const code = pick(r, H.code);
    const price = toNum(pick(r, H.price));
    const desc  = pick(r, H.desc);
    const series= detectSeries(r);
    let type    = detectType(desc);

    // brann tag
    const et2H  = (r._cols && r._cols.length>=8) ? r._cols[7] : '';
    const tag = String(et2H||'').replace(/\s+/g,'').toUpperCase();
    const tagAmp =
      tag==='B160'   ? 1250 :
      tag==='B190'   ? 1600 :
      tag==='B210'   ? 2000 :
      tag==='2XB160' ? 2500 :
      tag==='2XB190' ? 3200 :
      tag==='2XB210' ? 4000 :
      tag==='3XB160' ? 5000 : NaN;

    if (Number.isFinite(tagAmp)) {
      const d = String(desc||'').toLowerCase();
      const ext = /(external|utvendig|ytter)/.test(d);
      const int = /(internal|innvendig|inner)/.test(d);
      type = ext && !int ? 'fire_barrier_kit_external'
           : int && !ext ? 'fire_barrier_kit_internal'
           : 'fire_barrier_kit';
    } else if (!type && /fire|brann/i.test(String(desc||''))) {
      type = 'fire_barrier_kit';
    }

    const amp = Number.isFinite(tagAmp) ? tagAmp : extractAmpGeneric(r);

    return { code, type, series, ampere: amp, unit_price: price, _desc: desc, _et2: et2H };
  })

}

// match
function matchesLedere(row, ledere){
  if (row.series === 'XCM') return true;
  const c = String(row.code||'').toUpperCase();
  const implied = c.includes('-3W') ? '3F+PE' : '3F+N+PE';
  return ledere === implied || !c;
}
function byTypeAmpSeries(rows, type, amp, series){ return rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp)); }
function byTypeAmpSeriesL(rows, type, amp, series, ledere){
  const m = rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp) && matchesLedere(r, ledere));
  return m || rows.find(r=>r.type===type && r.series===series && Number(deriveAmp(r))===Number(amp)) || rows.find(r=>r.type===type && r.series===series);
}
function byBoxAll(catalog, kind, amp, prefSeries){
  const list = catalog.filter(r=>r.type===kind && (amp ? Number(deriveAmp(r))===Number(amp) : true));
  if (!list.length) return null;
  const inSeries = list.find(r=>r.series===prefSeries);
  return inSeries || list[0];
}
function findByTypeSeriesAmp(rows, type, series, amp){
  let r = rows.find(x=>x.type===type && x.series===series && Number(deriveAmp(x))===Number(amp));
  if (r) return r;
  const cand = rows.filter(x=>x.type===type && x.series===series);
  r = cand.find(x=>Number(deriveAmp(x))===Number(amp));
  return r || cand[0] || null;
}
function getFireBarrier(rows, amp, series){
  const direct = rows.find(r=>r.type==='fire_barrier_kit' && r.series===series && Number(deriveAmp(r))===amp);
  if (direct) return { code: direct.code, unit: toNum(direct.unit_price) };
  const ext = rows.find(r=>r.type==='fire_barrier_kit_external' && r.series===series && Number(deriveAmp(r))===amp);
  const int = rows.find(r=>r.type==='fire_barrier_kit_internal' && r.series===series && Number(deriveAmp(r))===amp);
  if (!ext && !int) return null;
  const unit = (ext?toNum(ext.unit_price):0) + (int?toNum(int.unit_price):0);
  const code = [ext?.code, int?.code].filter(Boolean).join('+') || 'FIRE-BARRIER';
  return { code, unit };
}

// plan
function planSegments(m){
  let rem = Math.ceil(Number(m));
  const plan = { n3:0, n15_2000:0, n500_1000:0 };
  while(rem > 3){ plan.n3++; rem -= 3; }
  if (rem > 2){ plan.n3++; }
  else if (rem > 1){ plan.n15_2000++; }
  else if (rem > 0){ plan.n500_1000++; }
  return plan;
}

// lines
function makeLine(row, series, amp, ledere, qty){
  const unit = toNum(row.unit_price);
  if (!Number.isFinite(unit)) throw new Error('Ugyldig pris: '+row.code);
  return { code: row.code, type: row.type, series, ampere: Number(deriveAmp(row))||amp||'', ledere, antall: qty, enhet: unit, sum: round2(unit*qty) };
}
function makeCustomLine(code, type, series, amp, ledere, unit, qty){
  if (!Number.isFinite(unit)) throw new Error('Ugyldig pris: '+type);
  return { code, type, series, ampere: amp||'', ledere, antall: qty, enhet: unit, sum: round2(unit*qty) };
}
// Hvilke type-navn tilsvarer 3m / ~2m / ~1m for valgt serie og distribusjon
function lengthTypes(series, dist){
  if (series==='XCM'){
    return dist
      ? {L3:'straight_3m_dist', L2:'xcm_dist_1500_2999',  L1:'xcm_dist_1000_1500'}
      : {L3:'xcm_feeder_3m',    L2:'xcm_feeder_1501_2999',L1:'xcm_feeder_600_1500'};
  }
  // XCP-S
  return dist
    ? {L3:'straight_3m_dist', L2:'straight_1501_2000_dist', L1:'straight_500_1000_dist'}
    : {L3:'straight_3m',      L2:'straight_1501_2000',      L1:'straight_500_1000'};
}

// Greedy plan (3,2,1) + fallback-regler
function planWithFallback(m, avail){
  // grunnplan
  let rem = Math.ceil(Number(m));
  let n3=0,n2=0,n1=0;
  while(rem>3){ n3++; rem-=3; }
  if (rem>2) n3++; else if (rem>1) n2++; else if (rem>0) n1++;

  // 1) Mangler 1m: bytt (1×3m + 1×1m) → (2×2m)
  if (!avail.L1 && avail.L2){
    while(n1>0 && n3>0){ n3--; n2+=2; n1--; }
  }
  // 2) Mangler 2m: bytt 1×2m → (1×3m + -1×1m)
  if (!avail.L2 && avail.L3 && avail.L1){
    while(n2>0 && n1>0){ n2--; n3++; n1--; }
  }
  // 3) Mangler 3m: bytt 1×3m → (1×2m + 1×1m) el. (3×1m) el. (2×2m - 1×1m)
  if (!avail.L3){
    if (avail.L2 && avail.L1){ while(n3>0){ n3--; n2++; n1++; } }
    else if (avail.L1){ while(n3>0){ n3--; n1+=3; } }
    else if (avail.L2){ while(n3>0 && n1>0){ n3--; n2+=2; n1--; } }
  }

  // 4) Hvis 2m fortsatt mangler: prøv 1×2m → 2×1m
  if (!avail.L2 && avail.L1){ while(n2>0){ n2--; n1+=2; } }

  // 5) Hvis 1m fortsatt mangler og finnes 2m: prøv å balansere med 3m
  if (!avail.L1 && avail.L2 && n1>0){
    // ikke mulig uten å låne 3m; forsøk en ekstra bytte hvis mulig
    if (n3>0){ n3--; n2+=2; n1--; }
  }

  // valider lengde
  const total = 3*n3 + 2*n2 + 1*n1;
  if (total !== Math.ceil(Number(m))) {
    throw new Error('Kan ikke finne lengdekombinasjon med tilgjengelige elementer.');
  }
  return {n3,n2,n1};
}

// Finn ut om type finnes for valgt konfig
function hasType(catRows, series, amp, ledere, type){
  const rows = catRows.filter(r=>r.series===series);
  const r = findByTypeSeriesAmp(rows, type, series, amp) || byTypeAmpSeriesL(rows,type,amp,series,ledere);
  return !!r;
}

// pris
// streng match på amp, men fall tilbake til type+serie hvis amp mangler
function needAnyAmp(rows, type, amp, series){
  return byTypeAmpSeries(rows, type, amp, series)
      || rows.find(r => r.type===type && r.series===series); // ingen amp i CSV
}

function price(cat, input){
  const bom=[];
  const push = (r,q)=>bom.push(makeLine(r, input.series, input.ampere, input.ledere, q));
  const need = (type)=> byTypeAmpSeries(cat.rows, type, input.ampere, input.series);

  // Startelement
if (input.startEl === 'board_feed'){
  const bf = needAnyAmp(cat.rows, 'board_feed', input.ampere, input.series);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.startEl === 'end_feed_unit'){
  const ef = needAnyAmp(cat.rows, 'end_feed_unit', input.ampere, input.series);
  if (!ef) throw new Error(`Mangler end_feed_unit for ${input.series}.`);
  push(ef,1);
}

// Sluttelement
if (input.sluttEl === 'board_feed'){
  const bf = needAnyAmp(cat.rows, 'board_feed', input.ampere, input.series);
  if (!bf) throw new Error(`Mangler board_feed for ${input.series}.`);
  push(bf,1);
} else if (input.sluttEl === 'crt_board_feed'){
  if (input.series !== 'XCP-S') throw new Error('Trafoelement er kun for XCP-S.');
  const crt = needAnyAmp(cat.rows, 'crt_board_feed', input.ampere, input.series);
  if (!crt) throw new Error(`Mangler crt_board_feed for ${input.series}.`);
  push(crt,1);
} else if (input.sluttEl === 'end_cover'){
  const ec = needAnyAmp(cat.rows, 'end_cover', input.ampere, input.series);
  if (!ec) throw new Error(`Mangler end_cover for ${input.series}.`);
  push(ec,1);
}

  // ⬇ Lengder med fallback
const tmap = lengthTypes(input.series, input.dist);
const avail = {
  L3: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L3),
  L2: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L2),
  L1: hasType(cat.rows, input.series, input.ampere, input.ledere, tmap.L1)
};
const pf = planWithFallback(input.meter, avail);

// legg til linjer etter plan
if (pf.n3){
  const r = findByTypeSeriesAmp(cat.rows, tmap.L3, input.series, input.ampere) || byTypeAmpSeriesL(cat.rows,tmap.L3,input.ampere,input.series,input.ledere);
  if(!r) throw new Error(`Mangler ${tmap.L3}.`);
  push(r, pf.n3);
}
if (pf.n2){
  const r = findByTypeSeriesAmp(cat.rows, tmap.L2, input.series, input.ampere) || byTypeAmpSeriesL(cat.rows,tmap.L2,input.ampere,input.series,input.ledere);
  if(!r) throw new Error(`Mangler ${tmap.L2}.`);
  push(r, pf.n2);
}
if (pf.n1){
  const r = findByTypeSeriesAmp(cat.rows, tmap.L1, input.series, input.ampere) || byTypeAmpSeriesL(cat.rows,tmap.L1,input.ampere,input.series,input.ledere);
  if(!r) throw new Error(`Mangler ${tmap.L1}.`);
  push(r, pf.n1);
}

  // Vinkler
  if (input.v90_h){ const r=byTypeAmpSeriesL(cat.rows,'elbow_horizontal_90',input.ampere,input.series,input.ledere); if(!r) throw new Error('Mangler elbow_horizontal_90.'); push(r,input.v90_h); }
  if (input.v90_v){ const r=byTypeAmpSeriesL(cat.rows,'elbow_vertical_90'  ,input.ampere,input.series,input.ledere); if(!r) throw new Error('Mangler elbow_vertical_90.');   push(r,input.v90_v); }

  // Avtappingsbokser
  if (input.boxQty>0){
    if (input.boxSel){
      const [kind, ampStr] = input.boxSel.split('|');
      if (kind==='bolt_on_box' && input.series==='XCM') throw new Error('Bolt-on box kan ikke brukes på XCM.');
      const row = byBoxAll(cat.catalog, kind, ampStr?Number(ampStr):undefined, input.series);
      if (!row) throw new Error(`Mangler ${kind} ${ampStr||''}A.`);
      bom.push(makeLine(row, input.series, deriveAmp(row)||'', input.ledere, input.boxQty));
    }else{
      const tryKinds = ['plug_in_box','tap_off_box','bolt_on_box'];
      let row = null;
      for (const k of tryKinds){
        if (k==='bolt_on_box' && input.series==='XCM') continue;
        row = byBoxAll(cat.catalog, k, undefined, input.series);
        if (row) break;
      }
      if (!row) throw new Error('Ingen avtappingsbokser funnet i data.');
      bom.push(makeLine(row, input.series, deriveAmp(row)||'', input.ledere, input.boxQty));
    }
  }

  // Brann
  if (input.fbQty>0){
    const fb = getFireBarrier(cat.rows, input.ampere, input.series);
    if (!fb) throw new Error('Mangler fire barrier kit.');
    bom.push(makeCustomLine(fb.code, 'fire_barrier_kit', input.series, input.ampere, input.ledere, fb.unit, input.fbQty));
  }

  // Ekspansjon
  if (input.meter > 30 && input.expansionYes){
    const exp = byTypeAmpSeries(cat.rows,'expansion_unit',input.ampere,input.series);
    if (!exp) throw new Error(`Mangler expansion_unit for ${input.series} ${input.ampere}.`);
    push(exp,1);
  }

  const material = round2(bom.reduce((s,x)=>s+x.sum,0));
  const margin   = round2(material*0.20);
  const subtotal = round2(material+margin);
  const rate     = Number(input.freightRate ?? 0.10);
  const freight  = round2(subtotal * rate);
  const montasje = calculateMontasje({
    meter: input.meter,
    angles: (input.v90_h || 0) + (input.v90_v || 0),
    amp: input.ampere,
    hourlyRate: input.montasjeSettings?.hourlyRate
  });
  const oppheng = calculateOpphengsmateriell({
    meter: input.meter,
    amp: input.ampere,
    ratePerMeter: input.montasjeSettings?.opphengRate
  });
  const totalExMontasje = round2(subtotal + freight);
  const total    = round2(totalExMontasje + montasje.cost + oppheng.cost);
  return { bom, material, margin, subtotal, freight, montasje, oppheng, totalExMontasje, total };
}

// --- app ---
let catalog=[];
let isDirty = false;

function markDirty(){
  isDirty = true;
  const st = $('status');
  if (st) st.textContent = 'Beregn for å få inkludere endringer';
}

function markClean(){
  isDirty = false;
  const st = $('status');
  if (st) st.textContent = 'OK';
}
window.addEventListener('DOMContentLoaded', async ()=>{
  try{
    ['meter','v90h','v90v','fbQty','boxQty'].forEach(id => $(id).value = '');

    const all = [];
    for (const p of RAW_CSV_PATHS){
      try{
        const res = await fetch(p,{cache:'no-store'}); if (!res.ok) continue;
        const txt = await res.text();
        all.push(...parseCSVAuto(txt));
      }catch{}
    }
    catalog = adaptRawToCatalog(all);

    $('series').addEventListener('change', refreshUIBySeries);
    $('meter').addEventListener('change', ()=>Math.ceil(Number($('meter').value||0)));
    $('meter').addEventListener('blur',  ()=>Math.ceil(Number($('meter').value||0)));

    const rateInput = $('montasjeHourlyRate');
    const opphengInput = $('opphengRate');
    const rateToggle = $('rateToggle');
    if (rateInput){
      if (!rateInput.value) rateInput.value = String(DEFAULT_HOURLY_RATE);
      const syncRate = ()=>{
        rateInput.value = String(sanitizeHourlyRate(rateInput.value));
        updateMontasjePreview();
        markDirty();
      };
      rateInput.addEventListener('input', ()=>{
        if (rateToggle && !rateToggle.checked) return;
        updateMontasjePreview();
        markDirty();
      });
      rateInput.addEventListener('change', syncRate);
      rateInput.addEventListener('blur', syncRate);
    }
    if (opphengInput){
      const sanitizeOpphengInput = (markDirtyAfter)=>{
        if (rateToggle && !rateToggle.checked) return;
        const raw = opphengInput.value;
        const defaultRate = toNum(opphengInput.dataset.defaultRate);
        const fallbackRate = Number.isFinite(defaultRate) ? defaultRate : 0;
        const sanitized = sanitizeOpphengRate(raw, fallbackRate);
        if (!(sanitized === fallbackRate && fallbackRate === 0 && raw.trim()==='')){
          opphengInput.value = (sanitized || sanitized === 0) ? String(sanitized) : '';
        }else{
          opphengInput.value = '';
        }
        opphengInput.dataset.appliedValue = String(sanitized);
        opphengInput.dataset.userOverride = sanitized !== fallbackRate ? 'true' : 'false';
        updateMontasjePreview();
        if (markDirtyAfter) markDirty();
      };
      opphengInput.addEventListener('input', ()=>{
        if (rateToggle && !rateToggle.checked) return;
        opphengInput.dataset.userOverride = 'true';
        updateMontasjePreview();
        markDirty();
      });
      opphengInput.addEventListener('change', ()=>sanitizeOpphengInput(true));
      opphengInput.addEventListener('blur', ()=>sanitizeOpphengInput(false));
    }
    if (rateToggle){
      rateToggle.checked = false;
      const applyRateLock = (markDirtyAfter)=>{
        const locked = !rateToggle.checked;
        setInputLocked(rateInput, locked);
        if (locked && rateInput){
          rateInput.value = String(DEFAULT_HOURLY_RATE);
        }
        if (opphengInput){
          setInputLocked(opphengInput, locked);
          if (locked){
            opphengInput.dataset.userOverride = 'false';
            opphengInput.value = '';
          }
        }
        updateMontasjePreview();
        if (markDirtyAfter) markDirty();
      };
      rateToggle.addEventListener('change', ()=>applyRateLock(true));
      applyRateLock(false);
    }else{
      setInputLocked(rateInput, false);
      if (opphengInput) setInputLocked(opphengInput, false);
    }
    ['meter','v90h','v90v'].forEach(id=>{
      const el = $(id);
      if (!el) return;
      el.addEventListener('input', ()=>{ updateMontasjePreview(); markDirty(); });
      el.addEventListener('change', ()=>{ updateMontasjePreview(); markDirty(); });
    });
    const ampSelEl = $('ampSelect');
    if (ampSelEl){
      ampSelEl.addEventListener('change', updateMontasjePreview);
    }
    // Sett dist til 'Nei' hvis tom, både ved last og når feltet forlates
    const distEl = $('dist');
    if (distEl){
      const ensureDist = ()=>{ if (!distEl.value) distEl.value = 'Nei'; };
      ensureDist();
      distEl.addEventListener('blur', ensureDist);
    }
    refreshUIBySeries();

    const frSel = document.getElementById('freightRate');
    if (frSel){
      frSel.addEventListener('change', ()=>{
        if (!lastCalc) return;
        const rate = Number(frSel.value || 0.10);
        const freight = round2((lastCalc.subtotal || 0) * rate);
        const totalEx = round2((lastCalc.subtotal || 0) + freight);
        const montasjeCost = round2(lastCalc.montasje?.cost || 0);
        const opphengCost = round2(lastCalc.oppheng?.cost || 0);
        const totalInclMontasje = round2(totalEx + montasjeCost);
        const total = round2(totalInclMontasje + opphengCost);
        const freightEl = $('freight');
        if (freightEl) freightEl.textContent = fmtNO.format(freight);
        const totalExEl = $('totalExMontasje');
        if (totalExEl) totalExEl.textContent = fmtNO.format(totalEx);
        const totalInclEl = document.getElementById('totalInclMontasje');
        if (totalInclEl) totalInclEl.textContent = fmtNO.format(totalInclMontasje);
        const totalEl = $('total');
        if (totalEl) totalEl.textContent = fmtNO.format(total);
        lastCalc.freight = freight;
        lastCalc.totalExMontasje = totalEx;
        lastCalc.totalInclMontasje = totalInclMontasje;
        lastCalc.total = total;
      });
    }

    // Markér status som "Oppdater..." ved endringer i parametere
    const dirtySelectors = [
      '#series','#dist','#meter','#v90h','#v90v','#ampSelect','#ledere',
      '#startEl','#sluttEl','#fbQty','#boxQty','#boxSel'
    ];
    dirtySelectors.forEach(sel=>{
      const el = document.querySelector(sel);
      if (!el) return;
      if (el.dataset.markDirtyBound) return;
      el.addEventListener('input', markDirty);
      el.addEventListener('change', markDirty);
      el.dataset.markDirtyBound = '1';
    });

  function enhanceNumberSteppers() {
  // Finn alle number-felt vi bruker
  const ids = ['meter','v90h','v90v','fbQty','boxQty'];
  ids.forEach(id=>{
    const input = document.getElementById(id);
    if (!input || input.dataset.enhanced) return;

    // Krav: heltall >= 0
    input.setAttribute('min','0');
    input.setAttribute('step','1');
    input.placeholder = input.placeholder || '0';

    // Pakk inn i stepper
    const wrap = document.createElement('div');
    wrap.className = 'stepper';
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);

    wrap.appendChild(input);

const plus  = document.createElement('button');
plus.type = 'button';
plus.className = 'btn-step';
plus.textContent = '+';

const minus = document.createElement('button');
minus.type = 'button';
minus.className = 'btn-step';
minus.textContent = '–';

wrap.appendChild(plus);
wrap.appendChild(minus);

    const clampInt = () => {
      const v = Math.max(0, Math.round(Number(input.value||0)));
      input.value = isFinite(v) ? String(v) : '0';
    };

    minus.addEventListener('click', ()=>{ input.value = String(Math.max(0, (Number(input.value||0)-1))); clampInt(); input.dispatchEvent(new Event('change')); });
    plus .addEventListener('click', ()=>{ input.value = String(Math.max(0, (Number(input.value||0)+1))); clampInt(); input.dispatchEvent(new Event('change')); });

    input.addEventListener('input', clampInt);
    input.addEventListener('blur',  clampInt);

    input.dataset.enhanced = '1';
  });
}

// Kall denne etter UI er bygd første gang
enhanceNumberSteppers();

    $('status').textContent = `CSV lastet (${catalog.length} varer)`;
  }catch(e){
    $('status').textContent = 'Feil CSV: '+(e.message||e);
  }
});

function refreshUIBySeries(){
  const series = $('series').value;

  // XCM låser ledere
  if (series==='XCM'){ $('ledere').value='3F+N+PE'; $('ledere').disabled=true; }
  else { $('ledere').disabled=false; if(!$('ledere').value) $('ledere').value=''; }

  // Sluttelement: skjul trafo for ikke-XCP-S
  const slutt = $('sluttEl');
  Array.from(slutt.options).forEach(opt=>{
    if (opt.value==='crt_board_feed') opt.hidden = (series!=='XCP-S');
  });
  if (series!=='XCP-S' && slutt.value==='crt_board_feed') slutt.value='';

  // Amp-valg
  const amps = Array.from(new Set(catalog.filter(r=>r.series===series && !/box$/.test(r.type)).map(r=>Number(deriveAmp(r))))).filter(Number.isFinite).sort((a,b)=>a-b);
  $('ampSelect').innerHTML = '<option value="">Velg…</option>' + amps.map(a=>`<option value="${a}">${a}</option>`).join('');

  // Bokser
  const boxes = catalog.filter(r=>['plug_in_box','tap_off_box','bolt_on_box'].includes(r.type));
  const labelOf = t => t==='plug_in_box'?'Plug-in box (plast)':t==='tap_off_box'?'Tap-off box (metall)':'Bolt-on box (metall)';
  const seen = new Set(); const opts = [];
  [...boxes.filter(b=>b.series===series), ...boxes.filter(b=>b.series!==series)].forEach(b=>{
    if (b.type==='bolt_on_box' && series==='XCM') return;
    const key = `${b.type}|${deriveAmp(b)||''}`;
    if (seen.has(key)) return; seen.add(key);
    const txt = `${deriveAmp(b)||''}A · ${labelOf(b.type)}`.replace(/^A · /,'');
    opts.push({v:`${b.type}|${deriveAmp(b)||''}`,t:txt});
  });
  opts.sort((a,b)=> (parseInt(a.t) || 1e9) - (parseInt(b.t) || 1e9) || String(a.t).localeCompare(b.t,'no'));
  $('boxSel').innerHTML = '<option value="">Velg...</option>'+opts.map(o=>`<option value="${o.v}">${o.t}</option>`).join('');

  updateMontasjePreview();
}

// ekspansjons-modal
function askExpansionIfNeeded(meter){
  return new Promise(resolve=>{
    if (meter <= 30){ resolve(false); return; }
    const bd = $('expModal'); bd.style.display='flex';
    const yes = $('expYes'), no = $('expNo');
    const done = (v)=>{ bd.style.display='none'; yes.onclick=null; no.onclick=null; resolve(v); };
    yes.onclick = ()=>done(true);
    no.onclick  = ()=>done(false);
  });
}

// beregn
$('calcBtn').addEventListener('click', async ()=>{
  if (!authState.loggedIn){
    const statusEl = $('status');
    if (statusEl) statusEl.textContent = 'Logg inn for \u00E5 beregne.';
    return;
  }
  try{
    if (!catalog.length) throw new Error('Ingen varer i katalog.');

    const series = $('series').value;
    if (!$('dist').value) { $('dist').value = 'Nei'; }
    const dist   = ($('dist').value==='Ja');
    const meter  = Math.ceil(Number($('meter').value || 0));
    const v90_h  = Number($('v90h').value || 0);
    const v90_v  = Number($('v90v').value || 0);
    const ampSel = $('ampSelect').value;
    const ledere = series==='XCM' ? '3F+N+PE' : $('ledere').value;
    const startEl= $('startEl').value;
    const sluttEl= $('sluttEl').value;
    const fbQty  = Number($('fbQty').value || 0);
    const boxQty = Number($('boxQty').value || 0);
    const boxSel = $('boxSel').value;

    if (!series) throw new Error('Velg system.');
    if (!meter) throw new Error('Angi meter (heltall).');
    if (!ampSel) throw new Error('Velg ampere.');
    if (series!=='XCM' && !ledere) throw new Error('Velg ledere.');
    if (!startEl) throw new Error('Velg startelement.');
    if (!sluttEl) throw new Error('Velg sluttelement.');

    if (!projectState.currentProject || !projectState.currentCustomer){
      const info = await ensureProjectInfo();
      if (!info) return;
    }

    const expansionYes = await askExpansionIfNeeded(meter);
    const amp = Number(ampSel);
    const rows = catalog.filter(r=>r.series===series);
    const freightRate = Number((document.getElementById('freightRate')?.value) || 0.10);
    const montasjeSettings = readMontasjeSettingsFromUI();

    const cat = { rows, catalog };
    const out = price(cat, {
      series, dist, meter, v90_h, v90_v, ampere: amp, ledere,
      startEl, sluttEl,
      fbQty, boxQty, boxSel,
      expansionYes, freightRate, montasjeSettings
    });
    $('mat').textContent      = fmtNO.format(out.material);
    $('margin').textContent   = fmtNO.format(out.margin);
    $('subtotal').textContent = fmtNO.format(out.subtotal);
    $('freight').textContent  = fmtNO.format(out.freight);
    const montasjeEl = document.getElementById('montasje');
    if (montasjeEl) montasjeEl.textContent = fmtNO.format(out.montasje.cost);
    const montasjeMarginVal = round2(out.montasje.cost * 0.20);
    const montasjeMarginEl = document.getElementById('montasjeMargin');
    if (montasjeMarginEl) montasjeMarginEl.textContent = fmtNO.format(montasjeMarginVal);
    const montasjeDetailText = formatMontasjeDetail(out.montasje);
    const montasjeDetailEl = document.getElementById('montasjeDetail');
    if (montasjeDetailEl) montasjeDetailEl.textContent = montasjeDetailText;
    const opphengEl = document.getElementById('oppheng');
    if (opphengEl) opphengEl.textContent = fmtNO.format(out.oppheng.cost);
    const opphengDetailText = formatOpphengDetail(out.oppheng);
    const opphengDetailEl = document.getElementById('opphengDetail');
    if (opphengDetailEl) opphengDetailEl.textContent = opphengDetailText;
    const totalExEl = document.getElementById('totalExMontasje');
    if (totalExEl) totalExEl.textContent = fmtNO.format(out.totalExMontasje);
    const totalInclMontasjeVal = round2(out.totalExMontasje + out.montasje.cost);
    const totalInclMontasjeEl = document.getElementById('totalInclMontasje');
    if (totalInclMontasjeEl) totalInclMontasjeEl.textContent = fmtNO.format(totalInclMontasjeVal);
    $('total').textContent    = fmtNO.format(out.total);
    lastCalc = {
      material: out.material,
      margin: out.margin,
      subtotal: out.subtotal,
      freight: out.freight,
      totalExMontasje: out.totalExMontasje,
      totalInclMontasje: totalInclMontasjeVal,
      montasje: out.montasje,
      oppheng: out.oppheng,
      montasjeDetail: montasjeDetailText,
      opphengDetail: opphengDetailText,
      total: out.total
    };
    markClean();

    const tbody = document.querySelector('#bomTbl tbody');
    tbody.innerHTML = '';
    out.bom.forEach(b=>{
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${b.code}</td><td>${b.type}</td><td>${b.series}</td><td>${b.ampere}</td><td>${b.lederes||b.ledere||''}</td><td>${b.antall}</td><td>${b.enhet.toFixed(2)}</td><td>${b.sum.toFixed(2)}</td>`;
      tbody.appendChild(tr);
    });
    document.getElementById('results').hidden = false;
    updateProjectMetaDisplay();

    document.getElementById('exportCsv').onclick = ()=>{
      const header = ['code','type','series','ampere','ledere','antall','enhet','sum'];
      const lines = [header.join(',')].concat(out.bom.map(b=>[b.code,b.type,b.series,b.ampere,b.ledere,b.antall,b.enhet,b.sum].join(',')));
      const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'BOM.csv'; a.click();
    };
  }catch(err){
    $('status').textContent = String(err.message||err);
  }
  
// Nullstill
document.getElementById('resetBtn').addEventListener('click', ()=>{
  // tøm felter
  ['series','dist','ampSelect','ledere','startEl','sluttEl','boxSel'].forEach(id=>{ const el=$(id); if(el){ el.value=''; el.disabled=false; }});
  ['meter','v90h','v90v','fbQty','boxQty'].forEach(id=>{ const el=$(id); if(el){ el.value=''; }});

  // skjul resultat
  const res = document.getElementById('results');
  if (res) res.hidden = true;

  // nullstill status
  const st = document.getElementById('status');
  if (st) st.textContent = '';

  // bygg UI på nytt
  const rateInput = $('montasjeHourlyRate');
  const rateToggle = $('rateToggle');
  if (rateToggle) rateToggle.checked = false;
  if (rateInput){
    rateInput.value = String(DEFAULT_HOURLY_RATE);
    setInputLocked(rateInput, true);
  }
  const opphengInput = $('opphengRate');
  if (opphengInput){
    opphengInput.value = '';
    opphengInput.dataset.appliedValue = '';
    opphengInput.dataset.userOverride = 'false';
    opphengInput.dataset.defaultRate = '';
    opphengInput.dataset.appliedAmp = '';
    setInputLocked(opphengInput, true);
  }
  refreshUIBySeries();
  updateMontasjePreview();
  projectState.currentProject = '';
  projectState.currentCustomer = '';
  updateProjectMetaDisplay();
  updateAuthUI();
  lastCalc = null;
  isDirty = false;
});

});

