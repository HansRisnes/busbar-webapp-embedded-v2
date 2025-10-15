// v2 + v2.1, XCP-S/XCM, distribusjon/feeder, avtappingsbokser, ekspansjon-modal >30 m

const RAW_CSV_PATHS = [
  'data/busbar-webapp-embedded-v2.csv',
  'data/busbar-webapp-embedded-v2.1.csv'
];
let lastCalc = null;

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
const toNum = x => {
  if (x===undefined || x===null) return NaN;
  const v = Number(String(x).replace(/\s/g,'').replace(',','.'));
  return Number.isFinite(v) ? v : NaN;
};
function pick(row, names){ for (const n of names){ if (n in row && row[n]!=='' && row[n]!==undefined) return row[n]; } return ''; }

const H = {
  code: ['code','Code','SKU','sku','produkt','Produkt'],
  price: ['price','Price','unit price','unit_price','Unit Price','UnitPrice','pris','Pris'],
  desc:  ['desc_text','description','Description','desc','tekst','Tekst'],
  amp:   ['ampere','Ampere','amp','Amp'],
  et2:   ['element_type_2','Element type 2','element type 2','H']
};

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
  const total    = round2(subtotal + freight);
  return { bom, material, margin, subtotal, freight, total };
}

// --- app ---
let catalog=[];
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
    refreshUIBySeries();

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

  // oppdater totaler når frakt endres
const frSel = document.getElementById('freightRate');
if (frSel){
  frSel.addEventListener('change', ()=>{
    if (!lastCalc) return;
    const rate = Number(frSel.value || 0.10);
    const freight = round2(lastCalc.subtotal * rate);
    const total   = round2(lastCalc.subtotal + freight);
    document.getElementById('freight').textContent = fmtNO.format(freight);
    document.getElementById('total').textContent   = fmtNO.format(total);
  });
}

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
  const labelOf = t => t==='plug_in_box'?'Plug-in box':t==='tap_off_box'?'Tap-off box':'B160 bolt-on box';
  const seen = new Set(); const opts = [];
  [...boxes.filter(b=>b.series===series), ...boxes.filter(b=>b.series!==series)].forEach(b=>{
    if (b.type==='bolt_on_box' && series==='XCM') return;
    const key = `${b.type}|${deriveAmp(b)||''}`;
    if (seen.has(key)) return; seen.add(key);
    const txt = `${deriveAmp(b)||''}A · ${labelOf(b.type)}`.replace(/^A · /,'');
    opts.push({v:`${b.type}|${deriveAmp(b)||''}`,t:txt});
  });
  opts.sort((a,b)=>String(a.t).localeCompare(b.t,'no'));
  $('boxSel').innerHTML = '<option value="">Auto</option>'+opts.map(o=>`<option value="${o.v}">${o.t}</option>`).join('');
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
  try{
    if (!catalog.length) throw new Error('Ingen varer i katalog.');

    const series = $('series').value;
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

    const expansionYes = await askExpansionIfNeeded(meter);
    const amp = Number(ampSel);
    const rows = catalog.filter(r=>r.series===series);
    const freightRate = Number((document.getElementById('freightRate')?.value) || 0.10);

    const cat = { rows, catalog };
    const out = price(cat, {
      series, dist, meter, v90_h, v90_v, ampere: amp, ledere,
      startEl, sluttEl,
      fbQty, boxQty, boxSel,
      expansionYes, freightRate
    });

    $('mat').textContent      = fmtNO.format(out.material);
    $('margin').textContent   = fmtNO.format(out.margin);
    $('subtotal').textContent = fmtNO.format(out.subtotal);
    $('freight').textContent  = fmtNO.format(out.freight);
    $('total').textContent    = fmtNO.format(out.total);
    lastCalc = { material: out.material, margin: out.margin, subtotal: out.subtotal };
    $('status').textContent = 'OK';

    const tbody = document.querySelector('#bomTbl tbody');
    tbody.innerHTML = '';
    out.bom.forEach(b=>{
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${b.code}</td><td>${b.type}</td><td>${b.series}</td><td>${b.ampere}</td><td>${b.lederes||b.ledere||''}</td><td>${b.antall}</td><td>${b.enhet.toFixed(2)}</td><td>${b.sum.toFixed(2)}</td>`;
      tbody.appendChild(tr);
    });
    document.getElementById('results').hidden = false;

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
  refreshUIBySeries();
});

});
