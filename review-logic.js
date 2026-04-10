// Variant Review Tool - dynamic file upload with column mapping
(function() {

const rvColors = ["#ffc7ce","#ffeb9c","#c6efce","#bdd7ee","#e2b5f4","#ffd966","#9fc5e8","#f4cccc","#b6d7a8","#cfe2f3","#d9ead3","#fce5cd","#d9d9d9","#ea9999","#a4c2f4","#b4a7d6","#d5a6bd","#93c47d","#76a5af","#f9cb9c","#ead1dc","#d0e0e3","#fff2cc","#e6b8a2","#c9daf8","#f4f1de"];

const STORAGE_KEY = "orchard_clean_v5_final";
const FLAGS_KEY = "orchard_flags_v1";
const NOTES_KEY = "orchard_notes_v1";
const BACKUP_KEY = "orchard_backup";

let rvStorageAvailable = true;
function rvStorageGet(key) { try { return JSON.parse(localStorage.getItem(key)); } catch(e) { rvStorageAvailable = false; return null; } }
function rvStorageSet(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch(e) { rvStorageAvailable = false; } }
try { localStorage.setItem("__rvtest__", "1"); if (localStorage.getItem("__rvtest__") !== "1") throw 0; localStorage.removeItem("__rvtest__"); } catch(e) { rvStorageAvailable = false; }

let touched = rvStorageGet(STORAGE_KEY) || {};
let flags = rvStorageGet(FLAGS_KEY) || {};
let notes = rvStorageGet(NOTES_KEY) || {};
let activeRow = null;
let rvRows = [];

// ── Uploaded file data ──
let uploadedHeaders = [];
let uploadedRows = [];
let dbHeaders = [];
let dbRows = [];
let dbSet = null; // Set of "site||name" strings from database file

const $ = id => document.getElementById(id);

// ── File upload handling ──
let rvPendingCb = null;

$('rv-file-upload').addEventListener('change', e => {
  if (!e.target.files[0]) return;
  const file = e.target.files[0];
  $('rv-file-name').textContent = file.name;
  readFile(file, (headers, rows) => {
    uploadedHeaders = headers;
    uploadedRows = rows;
    populateMapping();
    showPreview();
  });
});

$('rv-db-upload').addEventListener('change', e => {
  if (!e.target.files[0]) return;
  $('rv-db-name').textContent = e.target.files[0].name;
  readFile(e.target.files[0], (headers, rows) => {
    dbHeaders = headers;
    dbRows = rows;
    populateDbMapping();
    $('rv-db-mapping').style.display = '';
  });
});

function readFile(file, cb) {
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();
  if (ext === 'csv') {
    reader.onload = e => {
      const { headers, rows } = parseCSV(e.target.result);
      cb(headers, rows);
    };
    reader.readAsText(file);
  } else {
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      if (wb.SheetNames.length === 1) {
        loadSheet(wb, wb.SheetNames[0], cb);
      } else {
        showRvSheetPicker(wb, cb);
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

function loadSheet(wb, sheetName, cb) {
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const filtered = data.filter(r => r.some(c => String(c).trim() !== ''));
  if (filtered.length < 2) { alert('Sheet "' + sheetName + '" has no data.'); return; }
  cb(filtered[0].map(String), filtered.slice(1).map(r => r.map(String)));
}

function showRvSheetPicker(wb, cb) {
  // Create a simple picker modal
  let picker = document.getElementById('rv-sheet-picker');
  if (!picker) {
    picker = document.createElement('div');
    picker.id = 'rv-sheet-picker';
    picker.className = 'modal-overlay show';
    document.body.appendChild(picker);
  }
  let html = '<div class="modal"><h3>Select a Sheet</h3><div id="rv-sheet-picker-list">';
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const rows = data.filter(r => r.some(c => String(c).trim() !== ''));
    const count = Math.max(0, rows.length - 1);
    html += '<button class="btn btn-ghost" style="width:100%;text-align:left;margin-bottom:4px;justify-content:space-between;" data-sheet="' + name.replace(/"/g, '&quot;') + '">' +
      '<span>' + (typeof escHtml === 'function' ? escHtml(name) : name) + '</span><span class="text-muted small">' + count + ' rows</span></button>';
  });
  html += '</div><div class="modal-actions"><button class="btn btn-ghost" id="rv-sheet-picker-cancel">Cancel</button></div></div>';
  picker.innerHTML = html;
  picker.style.display = 'flex';

  picker.querySelector('#rv-sheet-picker-cancel').addEventListener('click', () => { picker.style.display = 'none'; });
  picker.querySelectorAll('[data-sheet]').forEach(btn => {
    btn.addEventListener('click', () => {
      picker.style.display = 'none';
      loadSheet(wb, btn.dataset.sheet, cb);
    });
  });
}

function parseCSV(text) {
  const headers = [], rows = [];
  const fields = [];
  let cur = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"' && text[i+1] === '"') { cur += '"'; i++; }
      else if (ch === '"') inQ = false;
      else cur += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') { fields.push(cur); cur = ''; }
      else if (ch === '\n' || (ch === '\r' && text[i+1] === '\n')) {
        fields.push(cur); cur = '';
        if (ch === '\r') i++;
        if (!headers.length) headers.push(...fields.splice(0));
        else rows.push(fields.splice(0));
      } else cur += ch;
    }
  }
  if (cur || fields.length) { fields.push(cur); if (!headers.length) headers.push(...fields); else rows.push(fields.splice(0)); }
  return { headers, rows };
}

// ── Populate column mapping dropdowns ──
function populateMapping() {
  $('rv-mapping').style.display = '';
  const ids = ['rv-map-site','rv-map-block','rv-map-crop','rv-map-variety','rv-map-rootstock','rv-map-type','rv-map-acreage','rv-map-plantcount'];
  ids.forEach(id => {
    const sel = $(id);
    if (!sel) return;
    sel.innerHTML = '<option value="">-- skip --</option>';
    uploadedHeaders.forEach((h, i) => { sel.innerHTML += '<option value="' + i + '">' + h + '</option>'; });
  });

  // Smart auto-detection with keyword matching
  const rules = {
    'rv-map-site': ['site', 'farm', 'ranch', 'grower', 'site name', 'site group'],
    'rv-map-block': ['block', 'location', 'name', 'location name', 'block name', 'section'],
    'rv-map-crop': ['crop', 'commodity', 'crop type'],
    'rv-map-variety': ['variety', 'cultivar', 'variety name', 'sub-variety', 'clone'],
    'rv-map-rootstock': ['rootstock', 'root stock', 'root'],
    'rv-map-type': ['type', 'location type', 'site type', 'block type'],
    'rv-map-acreage': ['acreage', 'acres', 'acre', 'area', 'length'],
    'rv-map-plantcount': ['plant count', 'plantcount', 'tree count', 'treecount', 'trees', 'plants', 'count']
  };

  // First pass: exact match (header equals keyword)
  uploadedHeaders.forEach((h, i) => {
    const l = h.toLowerCase().trim();
    Object.keys(rules).forEach(id => {
      if (rules[id].includes(l)) trySelect(id, i);
    });
  });

  // Second pass: partial match (header contains keyword)
  uploadedHeaders.forEach((h, i) => {
    const l = h.toLowerCase().trim();
    Object.keys(rules).forEach(id => {
      rules[id].forEach(keyword => {
        if (l.includes(keyword)) trySelect(id, i);
      });
    });
  });

  // Show setup actions + default data button
  $('rv-setup-actions').style.display = '';
  if (typeof rawData !== 'undefined' && rawData.length > 0) $('rv-btn-default').style.display = '';

  // Count how many were auto-mapped
  let mapped = 0;
  ids.forEach(id => { const s = $(id); if (s && s.value) mapped++; });
  if (mapped > 0) {
    // Highlight unmapped fields
    ids.forEach(id => {
      const s = $(id);
      if (!s) return;
      if (!s.value) s.style.borderColor = '#f59e0b';
      else s.style.borderColor = '';
    });
  }
}

function trySelect(id, val) { const s = $(id); if (s && !s.value) s.value = val; }

function populateDbMapping() {
  ['rv-dbmap-site','rv-dbmap-block'].forEach(id => {
    const sel = $(id);
    sel.innerHTML = '<option value="">-- select --</option>';
    dbHeaders.forEach((h, i) => { sel.innerHTML += '<option value="' + i + '">' + h + '</option>'; });
  });
  // Smart auto-detect
  const dbRules = {
    'rv-dbmap-site': ['site', 'farm', 'ranch', 'grower', 'site name'],
    'rv-dbmap-block': ['block', 'location', 'name', 'location name', 'block name']
  };
  dbHeaders.forEach((h, i) => {
    const l = h.toLowerCase().trim();
    Object.keys(dbRules).forEach(id => {
      dbRules[id].forEach(kw => { if (l.includes(kw)) trySelect(id, i); });
    });
  });
}

function showPreview() {
  const wrap = $('rv-preview-wrap');
  const tbl = $('rv-preview-table');
  wrap.style.display = '';
  let html = '<thead><tr>';
  uploadedHeaders.forEach(h => { html += `<th>${esc(h)}</th>`; });
  html += '</tr></thead><tbody>';
  uploadedRows.slice(0, 5).forEach(r => {
    html += '<tr>';
    r.forEach(c => { html += `<td>${esc(c)}</td>`; });
    html += '</tr>';
  });
  html += '</tbody>';
  tbl.innerHTML = html;
}

// ── Build the review from uploaded data ──
$('rv-btn-build').addEventListener('click', () => {
  const siteCol = +$('rv-map-site').value;
  const blockCol = +$('rv-map-block').value;
  const cropCol = +$('rv-map-crop').value;
  const varietyCol = +$('rv-map-variety').value;

  if (isNaN(siteCol) || $('rv-map-site').value === '') { alert('Select a Site column'); return; }
  if (isNaN(blockCol) || $('rv-map-block').value === '') { alert('Select a Block/Name column'); return; }

  const rootCol = $('rv-map-rootstock').value !== '' ? +$('rv-map-rootstock').value : -1;
  const acreCol = $('rv-map-acreage').value !== '' ? +$('rv-map-acreage').value : -1;
  const plantCol = $('rv-map-plantcount').value !== '' ? +$('rv-map-plantcount').value : -1;
  const hasCrop = $('rv-map-crop').value !== '';
  const hasVariety = $('rv-map-variety').value !== '';

  // Build database set if a db file was uploaded
  buildDbSet();

  // Normalize uploaded rows into rawData-like format
  const data = uploadedRows.map(r => {
    const site = (r[siteCol] || '').trim();
    const block = (r[blockCol] || '').trim();
    let crop = '', variety = '';

    if (hasCrop && hasVariety && cropCol === varietyCol) {
      // Combined "Crop-Variety" column (e.g., "CHERRIES-CHELAN")
      const val = (r[cropCol] || '').trim();
      const idx = val.indexOf('-');
      if (idx >= 0) { crop = val.substring(0, idx).trim(); variety = val.substring(idx + 1).trim(); }
      else { crop = val; variety = ''; }
    } else {
      if (hasCrop) crop = (r[cropCol] || '').trim();
      if (hasVariety) variety = (r[varietyCol] || '').trim();
    }

    const rootstock = rootCol >= 0 ? (r[rootCol] || '').trim() : '';
    const acreage = acreCol >= 0 ? (r[acreCol] || '').trim() : '';
    const plantCount = plantCol >= 0 ? (r[plantCol] || '').trim() : '';

    return { site, block, crop, variety, rootstock, acreage, plantCount };
  });

  // Group by site+block to find duplicates
  rvRows = [];
  const groupCounts = {};
  data.forEach(d => {
    const key = d.site + '||' + d.block;
    groupCounts[key] = (groupCounts[key] || 0) + 1;
  });

  const seenCounts = {};
  const colorMap = {};
  let colorIdx = 0;

  data.forEach(d => {
    const locId = d.site + '||' + d.block;
    if (!colorMap[locId]) colorMap[locId] = rvColors[colorIdx++ % rvColors.length];
    seenCounts[locId] = (seenCounts[locId] || 0) + 1;

    const isFirst = seenCounts[locId] === 1 && groupCounts[locId] > 1;
    const isUnique = groupCounts[locId] === 1;
    const locNameFull = d.variety ? d.block + ' - ' + d.variety : d.block;

    if (isFirst) {
      rvRows.push({
        data: [locNameFull, d.site, d.crop, d.variety, d.rootstock, 'BLOCK', d.acreage, d.plantCount],
        color: colorMap[locId],
        autoMark: true,
        uniqueId: `top-${d.site}-${d.block}-${d.variety}-${seenCounts[locId]}`
      });
    }

    if (!isFirst && !isUnique) {
      // Check if already in database
      const inDb = dbSet ? dbSet.has((d.site + '||' + locNameFull).toLowerCase().replace(/\s*-\s*/g,'-')) : false;
      if (!inDb) {
        rvRows.push({
          data: [locNameFull, d.site, d.crop, d.variety, d.rootstock, 'BLOCK', d.acreage, d.plantCount],
          color: colorMap[locId],
          autoMark: false,
          uniqueId: `mid-${d.site}-${d.block}-${d.variety}-${seenCounts[locId]}`
        });
      }
    }
  });

  // Switch to review UI
  $('rv-setup').style.display = 'none';
  $('rv-review-ui').style.display = '';
  initReviewUI();
});

// ── Load default hardcoded data ──
$('rv-btn-default').addEventListener('click', () => {
  if (typeof rawData === 'undefined' || !rawData.length) return;

  const dbNorm = typeof dbEntries !== 'undefined' ? new Set([...dbEntries].map(s => s.replace(/\s*-\s*/g, '-'))) : new Set();
  function dbHas(key) { return dbNorm.has(key.replace(/\s*-\s*/g, '-')); }

  rvRows = [];
  const groupCounts = {};
  rawData.forEach(r => { const key = r[0] + '||' + r[1]; groupCounts[key] = (groupCounts[key] || 0) + 1; });
  const seenCounts = {};
  const colorMap = {};
  let colorIdx = 0;

  rawData.forEach(r => {
    const locId = r[0] + '||' + r[1];
    if (!colorMap[locId]) colorMap[locId] = rvColors[colorIdx++ % rvColors.length];
    seenCounts[locId] = (seenCounts[locId] || 0) + 1;
    const isFirst = seenCounts[locId] === 1 && groupCounts[locId] > 1;
    const isUnique = groupCounts[locId] === 1;
    const variety = r[6], cropType = r[5];
    const locNameFull = r[1] + ' - ' + variety;
    const isAPPLES = cropType.toLowerCase().includes('apple');
    const isVR6State = r[0] === 'VALLEY ROZ 6 - STATE';
    const siteBase = isVR6State ? 'VALLEY ROZ 6 STATE' : r[0].replace(/-([A-Za-z])$/, '$1');
    const sep = isVR6State ? ' - ' : '-';

    if (isFirst) {
      rvRows.push({ data: [locNameFull, r[0], cropType, variety, r[9], 'BLOCK', r[10], r[11]], color: colorMap[locId], autoMark: true, uniqueId: `top-${r[0]}-${r[1]}-${variety}-${seenCounts[locId]}`, oldUniqueId: `top-${r[0]}-${r[1]}-${variety}` });
    }
    if (!isFirst && !isUnique) {
      const site63 = siteBase + sep + '63';
      const siteCrop = siteBase + sep + (isAPPLES ? '10' : '30');
      if (!dbHas(site63 + '||' + locNameFull)) {
        rvRows.push({ data: [locNameFull, site63, cropType, variety, r[9], 'BLOCK', r[10], r[11]], color: colorMap[locId], autoMark: false, uniqueId: `mid-v63-${r[0]}-${r[1]}-${variety}-${seenCounts[locId]}`, oldUniqueId: `mid-v63-${r[0]}-${r[1]}-${variety}` });
      }
      if (!dbHas(siteCrop + '||' + locNameFull)) {
        rvRows.push({ data: [locNameFull, siteCrop, cropType, variety, r[9], 'BLOCK', r[10], r[11]], color: colorMap[locId], autoMark: false, uniqueId: `mid-vcrop-${r[0]}-${r[1]}-${variety}-${seenCounts[locId]}`, oldUniqueId: `mid-vcrop-${r[0]}-${r[1]}-${variety}` });
      }
    }
  });

  $('rv-setup').style.display = 'none';
  $('rv-review-ui').style.display = '';
  initReviewUI();
});

// ── Back to setup ──
$('rv-btn-back').addEventListener('click', () => {
  $('rv-review-ui').style.display = 'none';
  $('rv-setup').style.display = '';
});

// ── Build database set from uploaded db file ──
function buildDbSet() {
  if (!dbRows.length) { dbSet = null; return; }
  const siteCol = +$('rv-dbmap-site').value;
  const nameCol = +$('rv-dbmap-block').value;
  if (isNaN(siteCol) || isNaN(nameCol) || $('rv-dbmap-site').value === '' || $('rv-dbmap-block').value === '') { dbSet = null; return; }
  dbSet = new Set();
  dbRows.forEach(r => {
    const key = ((r[siteCol] || '') + '||' + (r[nameCol] || '')).toLowerCase().replace(/\s*-\s*/g, '-');
    if (key !== '||') dbSet.add(key);
  });
}

function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// ════════════════════════════════════════════════════════════════════
// ── Review UI logic (same as before, runs after data is built) ──
// ════════════════════════════════════════════════════════════════════

function saveState() { rvStorageSet(STORAGE_KEY, touched); rvStorageSet(FLAGS_KEY, flags); rvStorageSet(NOTES_KEY, notes); }
function saveBackup() { rvStorageSet(BACKUP_KEY, { touched, flags, notes, time: Date.now() }); }

const completionTimestamps = [];
let lastKnownCount = 0;
const BREAK_THRESHOLD = 30000;
let onBreak = false;

function formatETA(seconds) {
  if (seconds < 60) return '< 1 min remaining';
  if (seconds < 3600) return `~${Math.ceil(seconds / 60)} min remaining`;
  const h = Math.floor(seconds / 3600), m = Math.ceil((seconds % 3600) / 60);
  return m > 0 ? `~${h}h ${m}m remaining` : `~${h}h remaining`;
}

function getActiveElapsed() {
  let active = 0;
  for (let i = 1; i < completionTimestamps.length; i++) {
    const gap = completionTimestamps[i].time - completionTimestamps[i - 1].time;
    if (gap < BREAK_THRESHOLD) active += gap;
  }
  return active / 1000;
}

setInterval(() => {
  if (completionTimestamps.length === 0) return;
  const idle = Date.now() - completionTimestamps[completionTimestamps.length - 1].time;
  if (idle >= BREAK_THRESHOLD && !onBreak) {
    onBreak = true;
    const el = $('rv-eta');
    if (el) { el.textContent = 'Paused — ETA resumes when you continue'; el.style.color = '#b07800'; }
  }
}, 5000);

function updateStat() {
  let count = 0, total = 0;
  document.querySelectorAll('#rv-tbody td[data-uid]').forEach(td => {
    if (td.dataset.auto === 'true') return;
    total++;
    if (touched[td.dataset.uid] === true) count++;
  });
  const pct = total > 0 ? (count / total * 100) : 0;
  const label = rvStorageAvailable ? 'Session Saved' : '⚠️ NOT PERSISTING';
  const now = Date.now();
  if (count > lastKnownCount) {
    onBreak = false;
    completionTimestamps.push({ time: now, count });
    if (completionTimestamps.length > 60) completionTimestamps.shift();
  }
  lastKnownCount = count;
  let etaText = '';
  const remaining = total - count;
  if (remaining > 0 && completionTimestamps.length >= 2) {
    const elapsed = getActiveElapsed();
    const completed = completionTimestamps[completionTimestamps.length - 1].count - completionTimestamps[0].count;
    if (elapsed > 0 && completed > 0) etaText = formatETA(remaining / (completed / elapsed));
  } else if (remaining > 0) { etaText = 'Estimating time...'; }
  else { etaText = 'Complete!'; }

  $('rv-stat').textContent = `✅ ${count} / ${total} | ${label}`;
  $('rv-bar').style.width = pct + '%';
  $('rv-pct').textContent = pct.toFixed(1) + '%';
  $('rv-detail').textContent = `${count} / ${total} cells completed`;
  const etaEl = $('rv-eta');
  etaEl.textContent = etaText;
  etaEl.style.color = onBreak ? '#b07800' : '#666';
}

function handleCell(td, val, cellId, ri, tr, manual = true) {
  if (touched[cellId] === undefined) touched[cellId] = td.dataset.auto === 'true';
  if (manual) {
    touched[cellId] = true;
    if (activeRow && activeRow !== tr) activeRow.classList.remove('active-row');
    activeRow = tr; tr.classList.add('active-row');
    if (navigator.clipboard) navigator.clipboard.writeText(val);
    td.classList.add('copied');
    setTimeout(() => td.classList.remove('copied'), 600);
    saveState();
  }
  touched[cellId] ? td.classList.add('touched') : td.classList.remove('touched');
  updateRowBtn(tr);
  updateStat();
}

function updateRowBtn(tr) {
  const allDone = Array.from(tr.querySelectorAll('td[data-uid]')).every(el => touched[el.dataset.uid] === true);
  const btn = tr.querySelector('.row-btn');
  if (btn) allDone ? btn.classList.add('done') : btn.classList.remove('done');
}

function toggleRow(ri, tr) {
  const tds = tr.querySelectorAll('td[data-uid]');
  const allTouched = Array.from(tds).every(td => touched[td.dataset.uid] === true);
  const target = !allTouched;
  tds.forEach(td => { touched[td.dataset.uid] = target; });
  tds.forEach(td => { target ? td.classList.add('touched') : td.classList.remove('touched'); });
  updateRowBtn(tr);
  saveState();
  updateStat();
}

// ── Context menu / flag / note ──
let ctxTarget = null;
const ctxMenu = $('rv-ctx');

function applyFlagAndNote(td, uid) {
  if (flags[uid]) { td.classList.add('flagged'); $('rv-ctx-flag').textContent = '✅ Unmark Red'; }
  else { td.classList.remove('flagged'); $('rv-ctx-flag').textContent = '🔴 Mark Red'; }
  notes[uid] ? td.classList.add('has-note') : td.classList.remove('has-note');
  td.title = notes[uid] || '';
}

document.addEventListener('contextmenu', e => {
  const td = e.target.closest('#rv-tbody td[data-uid]');
  if (!td) return;
  e.preventDefault();
  ctxTarget = td;
  $('rv-ctx-flag').textContent = flags[td.dataset.uid] ? '✅ Unmark Red' : '🔴 Mark Red';
  $('rv-ctx-clear-note').style.display = notes[td.dataset.uid] ? '' : 'none';
  ctxMenu.style.display = 'block';
  ctxMenu.style.left = Math.min(e.clientX, window.innerWidth - 170) + 'px';
  ctxMenu.style.top = Math.min(e.clientY, window.innerHeight - 100) + 'px';
});
document.addEventListener('click', () => { ctxMenu.style.display = 'none'; });

const noteTooltip = $('rv-tooltip');
document.addEventListener('mouseover', e => { const td = e.target.closest('#rv-tbody td.has-note'); if (td && notes[td.dataset.uid]) { noteTooltip.textContent = notes[td.dataset.uid]; noteTooltip.style.display = 'block'; } });
document.addEventListener('mousemove', e => { if (noteTooltip.style.display === 'none') return; noteTooltip.style.left = Math.min(e.clientX + 14, window.innerWidth - noteTooltip.offsetWidth - 8) + 'px'; noteTooltip.style.top = Math.min(e.clientY + 14, window.innerHeight - noteTooltip.offsetHeight - 8) + 'px'; });
document.addEventListener('mouseout', e => { if (e.target.closest('td.has-note')) noteTooltip.style.display = 'none'; });

$('rv-ctx-flag').addEventListener('click', () => { if (!ctxTarget) return; flags[ctxTarget.dataset.uid] = !flags[ctxTarget.dataset.uid]; applyFlagAndNote(ctxTarget, ctxTarget.dataset.uid); saveState(); ctxMenu.style.display = 'none'; });

const noteOverlay = $('rv-note-overlay');
const noteText = $('rv-note-text');
$('rv-ctx-note').addEventListener('click', () => { if (!ctxTarget) return; noteText.value = notes[ctxTarget.dataset.uid] || ''; noteOverlay.classList.add('show'); noteText.focus(); ctxMenu.style.display = 'none'; });
$('rv-ctx-clear-note').addEventListener('click', () => { if (!ctxTarget) return; delete notes[ctxTarget.dataset.uid]; applyFlagAndNote(ctxTarget, ctxTarget.dataset.uid); saveState(); ctxMenu.style.display = 'none'; });
$('rv-note-save').addEventListener('click', () => { if (!ctxTarget) return; const v = noteText.value.trim(); if (v) notes[ctxTarget.dataset.uid] = v; else delete notes[ctxTarget.dataset.uid]; applyFlagAndNote(ctxTarget, ctxTarget.dataset.uid); saveState(); noteOverlay.classList.remove('show'); });
$('rv-note-cancel').addEventListener('click', () => { noteOverlay.classList.remove('show'); });
noteOverlay.addEventListener('click', e => { if (e.target === noteOverlay) noteOverlay.classList.remove('show'); });

// ── Sort ──
let sortCol = -1, sortAsc = true;
function sortRows(col) {
  if (sortCol === col) sortAsc = !sortAsc; else { sortCol = col; sortAsc = true; }
  rvRows.sort((a, b) => {
    let va = a.data[col], vb = b.data[col];
    const na = parseFloat(va.replace(/,/g, '')), nb = parseFloat(vb.replace(/,/g, ''));
    let cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : va.localeCompare(vb);
    return sortAsc ? cmp : -cmp;
  });
  $('rv-tbody').innerHTML = '';
  buildTable();
  applySearch($('rv-search').value);
  document.querySelectorAll('#rv-main-table thead th[data-col]').forEach(th => {
    th.classList.remove('sort-active');
    const a = th.querySelector('.sort-arrow'); if (a) a.textContent = '▲▼';
  });
  const th = document.querySelector(`#rv-main-table thead th[data-col="${col}"]`);
  if (th) { th.classList.add('sort-active'); const a = th.querySelector('.sort-arrow'); if (a) a.textContent = sortAsc ? '▲' : '▼'; }
}

// ── Build table ──
function buildTable() {
  const tbody = $('rv-tbody');
  rvRows.forEach((rowObj, ri) => {
    const tr = document.createElement('tr');
    const actionTd = document.createElement('td');
    const btn = document.createElement('button');
    btn.className = 'row-btn'; btn.textContent = '✔';
    btn.onclick = () => toggleRow(ri, tr);
    actionTd.appendChild(btn);
    tr.appendChild(actionTd);

    rowObj.data.forEach((val, ci) => {
      const td = document.createElement('td');
      td.textContent = val;
      td.dataset.uid = `${rowObj.uniqueId}-${ci}`;
      td.dataset.auto = rowObj.autoMark;
      if (ci === 0) { td.classList.add('loc-name'); td.style.background = rowObj.color; }
      else { td.style.background = ri % 2 === 0 ? '#ffffff' : '#f7f7f7'; }

      td._lastClick = 0; td._wasMarked = false;
      td.onclick = () => {
        const now = Date.now();
        if (now - td._lastClick < 300) {
          td._lastClick = 0;
          if (td._wasMarked) { touched[td.dataset.uid] = false; td.classList.remove('touched'); updateRowBtn(tr); saveState(); updateStat(); }
        } else {
          td._wasMarked = touched[td.dataset.uid] === true;
          td._lastClick = now;
          handleCell(td, val, td.dataset.uid, ri, tr, true);
        }
      };
      tr.appendChild(td);

      // Migrate old keys
      if (rowObj.oldUniqueId) {
        const oldUid = `${rowObj.oldUniqueId}-${ci}`, newUid = td.dataset.uid;
        if (touched[newUid] === undefined && touched[oldUid] !== undefined) touched[newUid] = touched[oldUid];
        if (flags[newUid] === undefined && flags[oldUid] !== undefined) flags[newUid] = flags[oldUid];
        if (notes[newUid] === undefined && notes[oldUid] !== undefined) notes[newUid] = notes[oldUid];
      }
      if (touched[td.dataset.uid] === undefined) touched[td.dataset.uid] = rowObj.autoMark === true;
      touched[td.dataset.uid] ? td.classList.add('touched') : td.classList.remove('touched');
      applyFlagAndNote(td, td.dataset.uid);
    });
    updateRowBtn(tr);
    tbody.appendChild(tr);
  });
}

function applySearch(query) {
  const q = query.trim().toLowerCase().replace(/\s+/g, '');
  $('rv-tbody').querySelectorAll('tr').forEach(tr => {
    let match = false;
    tr.querySelectorAll('td[data-uid]').forEach(td => {
      td.classList.remove('search-hit');
      if (q && td.textContent.toLowerCase().replace(/\s+/g, '').includes(q)) { td.classList.add('search-hit'); match = true; }
    });
    tr.style.display = q && !match ? 'none' : '';
  });
}

$('rv-search').addEventListener('input', e => applySearch(e.target.value));

function initReviewUI() {
  const backup = rvStorageGet(BACKUP_KEY);
  if (backup) {
    const mc = Object.keys(touched).length, bc = Object.keys(backup.touched || {}).length;
    if (mc < bc) {
      Object.entries(backup.touched || {}).forEach(([k, v]) => { if (touched[k] === undefined) touched[k] = v; });
      Object.entries(backup.flags || {}).forEach(([k, v]) => { if (flags[k] === undefined) flags[k] = v; });
      Object.entries(backup.notes || {}).forEach(([k, v]) => { if (notes[k] === undefined) notes[k] = v; });
    }
  }
  $('rv-tbody').innerHTML = '';
  buildTable();
  saveState();
  saveBackup();
  updateStat();

  document.querySelectorAll('#rv-main-table thead th[data-col]').forEach(th => {
    if (th.querySelector('.sort-arrow')) return; // already initialized
    const arrow = document.createElement('span');
    arrow.className = 'sort-arrow'; arrow.textContent = '▲▼';
    th.appendChild(arrow);
    th.addEventListener('click', () => sortRows(parseInt(th.dataset.col)));
  });

  if (!rvStorageAvailable) { $('rv-stat').textContent = '⚠️ Storage unavailable!'; $('rv-stat').style.color = '#ff6666'; }
}

window.addEventListener('beforeunload', () => { saveState(); saveBackup(); });
setInterval(saveBackup, 30000);

// Public init
window.rvInit = function() {
  if (typeof populateToolSheetSelects === 'function') populateToolSheetSelects();
};

// Public bridge: load data from global import
window.rvLoadSheetData = function(headers, rows) {
  uploadedHeaders = headers;
  uploadedRows = rows.map(r => r.map(c => String(c)));
  $('rv-file-name').textContent = '(from imported sheet)';
  populateMapping();
  showPreview();
};

window.rvLoadDbData = function(headers, rows) {
  dbHeaders = headers;
  dbRows = rows.map(r => r.map(c => String(c)));
  $('rv-db-name').textContent = '(from imported sheet)';
  populateDbMapping();
  $('rv-db-mapping').style.display = '';
};

})();
