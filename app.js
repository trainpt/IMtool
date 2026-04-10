// ═══════════════════════════════════════════
// Table Merge Tool
// Uses data-store.js for all persistence
// ═══════════════════════════════════════════

const HEADERS = ['Location','Sites','Crop','Variety','Location Type','Planted Date','Acreage','Plant Count'];
const RANCH_COL = 1;
const ACREAGE_COL = 6;
const PLANTCOUNT_COL = 7;
const SKIP_FOR_DONE = new Set([RANCH_COL]);
const NUMERIC_COLS = new Set([ACREAGE_COL, PLANTCOUNT_COL]);

let currentOrg = null;
let currentRanch = null;
let copiedCells = {};
let sortCol = -1, sortAsc = true;

// ═══ Init ═══
function initMergePage() {
  renderMergeSidebar();
  renderMergeMain();
}

// ═══ Sidebar ═══
function renderMergeSidebar() {
  const container = document.getElementById('mergeOrgList');
  if (!container) return;
  const store = getStore();
  container.innerHTML = '';

  Object.keys(store.orgs || {}).sort().forEach(orgName => {
    const org = store.orgs[orgName];
    const isOpen = currentOrg === orgName;

    // Org header
    const header = document.createElement('div');
    header.className = 'sb-org-header';
    header.innerHTML = '<span class="arrow' + (isOpen ? ' open' : '') + '">&#9654;</span>' +
      '<span class="name">' + escHtml(orgName) + '</span>' +
      '<span class="del" title="Archive org">x</span>';
    header.addEventListener('click', e => {
      if (e.target.classList.contains('del')) { archiveOrg(orgName); return; }
      toggleMergeOrg(orgName);
    });
    container.appendChild(header);

    // Ranch list
    const list = document.createElement('div');
    list.className = 'sb-list' + (isOpen ? ' open' : '');

    const ranches = Object.keys(org.ranches || {}).sort((a, b) => {
      const ad = isRanchComplete(org.ranches[a]), bd = isRanchComplete(org.ranches[b]);
      if (ad !== bd) return ad ? 1 : -1;
      return a.localeCompare(b);
    });

    ranches.forEach(rName => {
      const ranch = org.ranches[rName];
      const done = isRanchComplete(ranch);
      const item = document.createElement('div');
      item.className = 'sb-item' + (currentOrg === orgName && currentRanch === rName ? ' active' : '') + (done ? ' done' : '');
      item.title = rName;
      item.innerHTML = '<span class="name">' + escHtml(rName) + ' (' + ranch.rows.length + ')' + (done ? ' \u2713' : '') + '</span>' +
        '<span class="del" title="Archive">x</span>';
      item.addEventListener('click', e => {
        if (e.target.classList.contains('del')) { archiveRanch(orgName, rName); return; }
        selectRanch(orgName, rName);
      });
      item.addEventListener('contextmenu', e => {
        e.preventDefault();
        showRanchCtxMenu(e.clientX, e.clientY, orgName, rName);
      });
      list.appendChild(item);
    });

    // Add ranch input
    const addRow = document.createElement('div');
    addRow.className = 'sidebar-input-row';
    addRow.innerHTML = '<input class="input-field" style="flex:1;padding:4px 6px;font-size:11px;background:var(--sidebar-hover);color:#ccc;border-color:var(--border-sidebar);" placeholder="New ranch...">' +
      '<button class="btn btn-primary btn-sm" style="padding:4px 8px;">+</button>';
    addRow.querySelector('button').addEventListener('click', () => {
      const input = addRow.querySelector('input');
      const name = input.value.trim();
      if (!name) return;
      const s = getStore();
      if (!s.orgs[orgName]) return;
      if (s.orgs[orgName].ranches[name]) { alert('Ranch exists.'); return; }
      s.orgs[orgName].ranches[name] = { rows: [], progress: {} };
      saveStore(s);
      input.value = '';
      currentOrg = orgName; currentRanch = name;
      renderMergeSidebar(); renderMergeMain();
    });
    addRow.querySelector('input').addEventListener('keydown', e => {
      if (e.key === 'Enter') addRow.querySelector('button').click();
    });
    list.appendChild(addRow);
    container.appendChild(list);
  });

  // Quick import from imported sheets
  const sheets = getSheetNames();
  if (sheets.length > 0) {
    const section = document.createElement('div');
    section.style.cssText = 'border-top:1px solid var(--border-sidebar);padding:6px 0;margin-top:4px;';
    section.innerHTML = '<div class="sidebar-label">Import from Sheets</div>';
    sheets.forEach(name => {
      const s = getSheet(name);
      const item = document.createElement('div');
      item.className = 'sb-item';
      item.title = 'Load "' + name + '" into Table Merge';
      item.innerHTML = '<span class="name">' + escHtml(name) + '</span><span class="count">' + (s?s.rowCount:'') + '</span>';
      item.addEventListener('click', () => {
        if (!currentOrg) { alert('Select or create an org first.'); return; }
        const sheet = getSheet(name);
        if (sheet) processSheet([sheet.headers, ...sheet.rows]);
      });
      section.appendChild(item);
    });
    container.appendChild(section);
  }

  // New org button
  const newOrgBtn = document.getElementById('newOrgBtn');
  if (newOrgBtn && !newOrgBtn._bound) {
    newOrgBtn._bound = true;
    newOrgBtn.addEventListener('click', () => {
      const input = document.getElementById('newOrgName');
      const name = input.value.trim();
      if (!name) return;
      if (createOrg(name)) { currentOrg = name; input.value = ''; renderMergeSidebar(); }
      else alert('Org exists.');
    });
  }
}

function toggleMergeOrg(orgName) {
  currentOrg = currentOrg === orgName ? null : orgName;
  currentRanch = null;
  renderMergeSidebar(); renderMergeMain();
}

function selectRanch(orgName, ranchName) {
  currentOrg = orgName; currentRanch = ranchName;
  sortCol = -1; sortAsc = true;
  renderMergeSidebar(); renderMergeMain();
}

function isRanchComplete(ranch) {
  if (!ranch || !ranch.rows.length) return false;
  return ranch.rows.every((row, ri) => {
    const req = row.reduce((cols, c, ci) => { if (c !== '' && !SKIP_FOR_DONE.has(ci)) cols.push(ci); return cols; }, []);
    const done = new Set((ranch.progress || {})[ri] || []);
    return req.every(ci => done.has(ci));
  });
}

// ═══ Main Table ═══
function renderMergeMain() {
  const thead = document.getElementById('mergeThead');
  const tbody = document.getElementById('mergeTbody');
  const empty = document.getElementById('mergeEmpty');
  const bc = document.getElementById('mergeBreadcrumb');
  const toolbar = document.getElementById('mergeToolbar');
  thead.innerHTML = ''; tbody.innerHTML = ''; copiedCells = {};

  // Toolbar button visibility
  const hasOrg = !!currentOrg;
  const hasRanch = !!currentRanch;
  ['mergeUploadBtn','mergeHeaderBtn','mergePasteBtn'].forEach(id => document.getElementById(id).style.display = hasOrg ? '' : 'none');
  ['mergeDoneAllBtn','mergeResetBtn','mergeClearBtn'].forEach(id => document.getElementById(id).style.display = hasRanch ? '' : 'none');

  if (!hasOrg || !hasRanch) {
    bc.innerHTML = hasOrg ? escHtml(currentOrg) + ' — select a ranch' : 'Select a project';
    empty.style.display = ''; hideStatsBar(); return;
  }

  bc.innerHTML = '<span style="color:var(--text-3)">' + escHtml(currentOrg) + ' ›</span> ' + escHtml(currentRanch);
  const store = getStore();
  const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
  if (!ranch || !ranch.rows.length) { empty.style.display = ''; hideStatsBar(); return; }
  empty.style.display = 'none';

  // Auto-mark None/0/empty as done
  let autoMarked = false;
  ranch.rows.forEach((row, ri) => {
    const prog = new Set((ranch.progress || {})[ri] || []);
    row.forEach((cell, ci) => {
      if ((cell === 'None' || cell === '0' || cell === '') && !prog.has(ci)) { prog.add(ci); autoMarked = true; }
    });
    if (!ranch.progress) ranch.progress = {};
    ranch.progress[ri] = [...prog];
  });
  if (autoMarked) saveStore(store);

  // Render headers
  HEADERS.forEach((h, ci) => {
    const th = document.createElement('th');
    let arrow = '<span style="margin-left:3px;font-size:8px;opacity:0.3;">\u25B2</span>';
    if (sortCol === ci) { arrow = '<span style="margin-left:3px;font-size:8px;">' + (sortAsc ? '\u25B2' : '\u25BC') + '</span>'; th.classList.add('sort-active'); }
    th.innerHTML = escHtml(h) + arrow;
    th.addEventListener('click', () => { if (sortCol === ci) sortAsc = !sortAsc; else { sortCol = ci; sortAsc = true; } renderMergeMain(); });
    thead.appendChild(th);
  });

  // Sort
  const indices = ranch.rows.map((_, i) => i);
  if (sortCol >= 0) {
    indices.sort((a, b) => {
      let va = ranch.rows[a][sortCol] || '', vb = ranch.rows[b][sortCol] || '';
      if (NUMERIC_COLS.has(sortCol)) { va = parseFloat(va) || 0; vb = parseFloat(vb) || 0; return sortAsc ? va - vb : vb - va; }
      return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
  }

  // Render rows
  const progress = ranch.progress || {};
  indices.forEach(ri => {
    const row = ranch.rows[ri];
    copiedCells[ri] = new Set(progress[ri] || []);
    const nonEmpty = row.filter((c, i) => c !== '' && !SKIP_FOR_DONE.has(i)).length;
    const tr = document.createElement('tr');

    row.forEach((cell, ci) => {
      const td = document.createElement('td');
      td.textContent = cell;
      if (copiedCells[ri].has(ci)) td.classList.add('cell-highlighted');

      let clickCount = 0;
      td.addEventListener('click', () => {
        if (cell === '') return;
        clickCount++;
        if (clickCount === 1) {
          setTimeout(() => {
            if (clickCount === 1) {
              navigator.clipboard.writeText(cell).then(() => {
                const prev = document.querySelector('#mergeTbody .cell-clicked');
                if (prev) prev.classList.remove('cell-clicked');
                td.classList.add('cell-highlighted', 'cell-clicked');
                copiedCells[ri].add(ci);
                saveMergeProgress();
                checkMergeRowDone(tr, ri, row, nonEmpty);
              });
            } else {
              td.classList.remove('cell-highlighted', 'cell-clicked');
              copiedCells[ri].delete(ci);
              saveMergeProgress();
              checkMergeRowDone(tr, ri, row, nonEmpty);
            }
            clickCount = 0;
          }, 200);
        }
      });
      tr.appendChild(td);
    });

    tr.addEventListener('contextmenu', e => { e.preventDefault(); showRowCtxMenu(e.clientX, e.clientY, ri); });
    checkMergeRowDone(tr, ri, row, nonEmpty);
    tbody.appendChild(tr);
  });

  renderMergeStats();
}

function checkMergeRowDone(tr, ri, row, nonEmpty) {
  const done = [...copiedCells[ri]].filter(i => row[i] !== '' && !SKIP_FOR_DONE.has(i)).length;
  const blockCell = tr.querySelector('td');
  if (done >= nonEmpty && nonEmpty > 0) {
    tr.classList.add('row-done');
    if (!blockCell.querySelector('.badge')) { blockCell.insertAdjacentHTML('beforeend', '<span class="badge">\u2713 Done</span>'); }
  } else {
    tr.classList.remove('row-done');
    const badge = blockCell?.querySelector('.badge'); if (badge) badge.remove();
  }
}

function saveMergeProgress() {
  if (!currentOrg || !currentRanch) return;
  const store = getStore();
  const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
  if (!ranch) return;
  const obj = {};
  for (const key in copiedCells) obj[key] = [...copiedCells[key]];
  ranch.progress = obj;
  saveStore(store);
  renderMergeSidebar();
  renderMergeStats();
}

function renderMergeStats() {
  if (!currentOrg || !currentRanch) { hideStatsBar(); return; }
  const store = getStore();
  const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
  if (!ranch || !ranch.rows.length) { hideStatsBar(); return; }
  let total = ranch.rows.length, done = 0;
  ranch.rows.forEach((row, ri) => {
    const req = row.reduce((cols, c, ci) => { if (c !== '' && !SKIP_FOR_DONE.has(ci)) cols.push(ci); return cols; }, []);
    const d = new Set((ranch.progress || {})[ri] || []);
    if (req.every(ci => d.has(ci))) done++;
  });
  showStatsBar(total, done);
}

// ═══ Toolbar Actions ═══
function initMergeToolbarEvents() {
  document.getElementById('mergeDoneAllBtn').addEventListener('click', () => {
    if (!currentOrg || !currentRanch) return;
    const store = getStore();
    const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
    if (!ranch) return;
    ranch.rows.forEach((row, ri) => {
      ranch.progress[ri] = row.map((_, ci) => ci).filter(ci => row[ci] !== '');
    });
    saveStore(store); renderMergeMain(); renderMergeSidebar();
  });

  document.getElementById('mergeResetBtn').addEventListener('click', () => {
    if (!currentRanch || !confirm('Reset progress for "' + currentRanch + '"?')) return;
    const store = getStore();
    const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
    if (!ranch) return;
    ranch.progress = {};
    saveStore(store); renderMergeMain(); renderMergeSidebar();
  });

  document.getElementById('mergeClearBtn').addEventListener('click', () => {
    if (!currentRanch || !confirm('Archive all data from "' + currentRanch + '"?')) return;
    const store = getStore();
    const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
    if (!ranch || !ranch.rows.length) return;
    archiveItem('rows', currentRanch, JSON.parse(JSON.stringify(ranch)), currentOrg);
    ranch.rows = []; ranch.progress = {};
    saveStore(store); renderMergeMain(); renderMergeSidebar();
  });

  // Header setup
  document.getElementById('mergeHeaderBtn').addEventListener('click', () => {
    const box = document.getElementById('mergeHeaderSetup');
    box.style.display = box.style.display === 'none' ? '' : 'none';
  });
  document.getElementById('mergeHeaderCancel').addEventListener('click', () => { document.getElementById('mergeHeaderSetup').style.display = 'none'; });
  document.getElementById('mergeHeaderConfirm').addEventListener('click', () => {
    const text = document.getElementById('mergeHeaderArea').value.trim();
    if (!text || !currentOrg) return;
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) return;
    const fileHeaders = lines[0].split('\t').map(h => h.trim());
    document.getElementById('mergeHeaderSetup').style.display = 'none';
    document.getElementById('mergeHeaderArea').value = '';
    showMergeColumnMapper(fileHeaders);
  });

  // Paste
  document.getElementById('mergePasteBtn').addEventListener('click', () => {
    const box = document.getElementById('mergePasteBox');
    box.style.display = box.style.display === 'none' ? '' : 'none';
  });
  document.getElementById('mergePasteCancel').addEventListener('click', () => { document.getElementById('mergePasteBox').style.display = 'none'; });
  document.getElementById('mergePasteConfirm').addEventListener('click', () => {
    const text = document.getElementById('mergePasteArea').value.trim();
    if (!text || !currentOrg) return;
    const store = getStore();
    const mapping = store.orgs[currentOrg]?.columnMapping;
    if (!mapping) { alert('Set up header mapping first.'); return; }
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    const parsed = lines.map(l => l.split('\t').map(c => c.trim()));
    document.getElementById('mergePasteBox').style.display = 'none';
    document.getElementById('mergePasteArea').value = '';
    importMergeRows(parsed, mapping);
  });

  // File upload
  document.getElementById('mergeFileInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file || !currentOrg) return;
    e.target.value = '';
    try {
      const result = await readUploadedFile(file);
      if (result.sheets.length === 1) {
        processSheet([result.sheets[0].headers, ...result.sheets[0].rows]);
      } else {
        // Multi-sheet: show picker (reuse global import approach)
        alert('Multi-sheet file: use Import Data button for multi-sheet files. Importing first sheet.');
        processSheet([result.sheets[0].headers, ...result.sheets[0].rows]);
      }
    } catch(err) { alert('Error: ' + err.message); }
  });

  // Paste listener (Ctrl+V on page)
  document.addEventListener('paste', e => {
    if (!document.getElementById('page-merge').classList.contains('active')) return;
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    if (!currentOrg) return;
    const store = getStore();
    const mapping = store.orgs[currentOrg]?.columnMapping;
    if (!mapping) return;
    e.preventDefault();
    let parsed = null;
    const html = e.clipboardData.getData('text/html');
    if (html) { const doc = new DOMParser().parseFromString(html, 'text/html'); const rows = doc.querySelectorAll('tr'); if (rows.length) { parsed = []; rows.forEach(tr => { const cells = []; tr.querySelectorAll('td,th').forEach(c => cells.push(c.textContent.trim())); if (cells.length) parsed.push(cells); }); } }
    if (!parsed) { const text = e.clipboardData.getData('text/plain'); if (text) parsed = text.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t').map(c => c.trim())); }
    if (parsed && parsed.length) importMergeRows(parsed, mapping);
  });
}

// ═══ Column Mapping ═══
function showMergeColumnMapper(fileHeaders) {
  const mapper = document.getElementById('mergeColumnMapper');
  let html = '<div class="panel-box"><h3 style="font-size:14px;margin-bottom:10px;">Map columns</h3>' +
    '<p class="text-muted small" style="margin-bottom:8px;">Headers: ' + fileHeaders.join(', ') + '</p>';
  HEADERS.forEach((h, i) => {
    html += '<div style="display:flex;gap:8px;align-items:center;margin-bottom:4px;"><label style="width:100px;font-size:12px;font-weight:500;">' + h + '</label><select class="input-field" data-target="' + i + '" style="flex:1;">';
    html += '<option value="-1">(skip)</option>';
    fileHeaders.forEach((fh, fi) => { html += '<option value="' + fi + '"' + (autoMatch(h, fh) ? ' selected' : '') + '>' + escHtml(fh) + '</option>'; });
    html += '</select></div>';
  });
  html += '<div style="margin-top:10px;display:flex;gap:6px;"><button class="btn btn-primary" id="mapperConfirm">Save Mapping</button><button class="btn btn-ghost" id="mapperCancel">Cancel</button></div></div>';
  mapper.innerHTML = html;
  mapper.style.display = '';
  document.getElementById('mapperConfirm').addEventListener('click', () => {
    const mapping = [];
    mapper.querySelectorAll('select').forEach(sel => mapping.push(parseInt(sel.value)));
    const store = getStore();
    if (!store.orgs[currentOrg]) return;
    store.orgs[currentOrg].columnMapping = mapping;
    store.orgs[currentOrg].fileHeaders = fileHeaders;
    saveStore(store);
    mapper.style.display = 'none';
    updateMergeMapping();
  });
  document.getElementById('mapperCancel').addEventListener('click', () => { mapper.style.display = 'none'; });
}

function autoMatch(target, source) {
  const t = target.toLowerCase().replace(/[^a-z]/g, '');
  const s = source.toLowerCase().replace(/[^a-z]/g, '');
  if (t === s) return true;
  const aliases = { location:['location','blockname','block','name','blockid','locationname'], sites:['sites','site','ranch','ranchname','farm','sitename'], crop:['crop','croptype','commodity'], variety:['variety','varietyname'], locationtype:['locationtype','type','blocktype','sitetype'], planteddate:['planteddate','planted','dateplanted','plantdate'], acreage:['acreage','acres','area','length'], plantcount:['plantcount','treecount','trees','count','plants'] };
  return aliases[t]?.includes(s) || false;
}

function updateMergeMapping() {
  const el = document.getElementById('mergeMapping');
  if (!currentOrg) { el.innerHTML = ''; return; }
  const store = getStore();
  const org = store.orgs[currentOrg];
  if (org?.columnMapping) {
    const fh = (org.fileHeaders || []).filter(h => h).slice(0, 6);
    const extra = (org.fileHeaders || []).filter(h => h).length > 6 ? ' +' + ((org.fileHeaders || []).filter(h=>h).length - 6) + ' more' : '';
    el.innerHTML = '<span class="text-muted small">Mapped: ' + fh.join(', ') + extra + ' <a href="#" onclick="clearMergeMapping();return false;" style="color:var(--red);">Reset</a></span>';
  } else {
    el.innerHTML = '<span style="color:var(--amber);font-size:11px;">No headers set</span>';
  }
}

function clearMergeMapping() {
  if (!currentOrg || !confirm('Clear header mapping?')) return;
  const store = getStore();
  delete store.orgs[currentOrg].columnMapping;
  delete store.orgs[currentOrg].fileHeaders;
  saveStore(store);
  updateMergeMapping();
}

// ═══ Import Rows ═══
function processSheet(sheetRows) {
  const filtered = sheetRows.filter(r => r.some(c => String(c).trim() !== ''));
  if (filtered.length < 2) { alert('No data rows.'); return; }
  const store = getStore();
  const mapping = store.orgs[currentOrg]?.columnMapping;
  if (mapping) {
    importMergeRows(filtered.slice(1), mapping);
  } else {
    showMergeColumnMapper(filtered[0].map(h => String(h).trim()));
  }
}

function importMergeRows(dataRows, mapping) {
  const newRows = dataRows.map(sr => HEADERS.map((_, hi) => {
    const fi = mapping[hi];
    if (fi < 0 || fi >= sr.length) return '';
    return String(sr[fi]).trim();
  })).filter(r => r.some(c => c));
  if (!newRows.length) { alert('No rows.'); return; }

  // Safety: swap acreage/plant count if both > 0 and acreage > plant count
  newRows.forEach(row => {
    const a = parseFloat(row[ACREAGE_COL]), p = parseFloat(row[PLANTCOUNT_COL]);
    if (!isNaN(a) && !isNaN(p) && a > 0 && p > 0 && a > p) { row[ACREAGE_COL] = row[PLANTCOUNT_COL]; row[PLANTCOUNT_COL] = String(a); }
  });

  // Group by ranch
  const grouped = {};
  newRows.forEach(row => { const r = row[RANCH_COL] || 'Unknown'; if (!grouped[r]) grouped[r] = []; grouped[r].push(row); });

  // Show import destination modal
  showImportDestModal(grouped);
}

function showImportDestModal(grouped) {
  const store = getStore();
  const existing = Object.keys(store.orgs[currentOrg]?.ranches || {}).sort();
  const incoming = Object.keys(grouped).sort();
  const allTargets = [...new Set([...existing, ...incoming])].sort();
  const list = document.getElementById('importDestList');
  list.innerHTML = '';

  incoming.forEach(srcName => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:8px;margin-bottom:6px;padding:8px;background:var(--bg-alt);border-radius:var(--radius);';
    const label = document.createElement('div');
    label.innerHTML = '<strong style="font-size:12px;">' + escHtml(srcName) + '</strong><br><span class="text-muted small">' + grouped[srcName].length + ' row(s)</span>';
    const select = document.createElement('select');
    select.className = 'input-field'; select.style.flex = '1'; select.dataset.src = srcName;
    select.innerHTML = '<option value="__keep__">Keep as "' + escHtml(srcName) + '"</option>';
    // New incoming first
    allTargets.filter(r => r !== srcName && !existing.includes(r)).forEach(r => { select.innerHTML += '<option value="' + escHtml(r) + '">Merge into "' + escHtml(r) + '" (new)</option>'; });
    existing.forEach(r => { if (r !== srcName) select.innerHTML += '<option value="' + escHtml(r) + '">Merge into "' + escHtml(r) + '"</option>'; });
    row.appendChild(label); row.appendChild(select);
    list.appendChild(row);
  });

  document.getElementById('importDestModal').classList.add('show');
  document.getElementById('importDestCancel').onclick = () => document.getElementById('importDestModal').classList.remove('show');
  document.getElementById('importDestConfirm').onclick = () => {
    const store = getStore();
    let firstNew = null;
    document.querySelectorAll('#importDestList select').forEach(sel => {
      const srcName = sel.dataset.src, target = sel.value;
      const rows = grouped[srcName]; if (!rows) return;
      const destName = target === '__keep__' ? srcName : target;
      const isMerge = target !== '__keep__';
      rows.forEach(r => { r[RANCH_COL] = destName; });
      if (!store.orgs[currentOrg].ranches[destName]) store.orgs[currentOrg].ranches[destName] = { rows: [], progress: {} };
      const ranch = store.orgs[currentOrg].ranches[destName];
      const startIdx = ranch.rows.length;
      rows.forEach((row, i) => {
        ranch.rows.push(row);
        if (isMerge) { ranch.progress[startIdx + i] = []; }
        else { const allCols = []; row.forEach((c, ci) => { if (c) allCols.push(ci); }); ranch.progress[startIdx + i] = allCols; }
      });
      if (!firstNew && isMerge) firstNew = destName;
    });
    saveStore(store);
    document.getElementById('importDestModal').classList.remove('show');
    if (firstNew) currentRanch = firstNew;
    else if (!currentRanch) { const r = Object.keys(store.orgs[currentOrg].ranches).sort(); if (r.length) currentRanch = r[0]; }
    renderMergeSidebar(); renderMergeMain();
  };
}

// ═══ Context Menus ═══
function showRowCtxMenu(x, y, ri) {
  const menu = document.getElementById('ctxMenu');
  menu.innerHTML = '<div data-action="move">Move to another ranch</div><div data-action="delete">Archive row</div>';
  menu.style.display = 'block'; menu.style.left = x + 'px'; menu.style.top = y + 'px';
  menu.onclick = e => {
    menu.style.display = 'none';
    if (e.target.dataset.action === 'move') openMoveRowModal(ri);
    else if (e.target.dataset.action === 'delete') archiveRow(ri);
  };
}

function showRanchCtxMenu(x, y, orgName, ranchName) {
  const menu = document.getElementById('ctxMenu');
  menu.innerHTML = '<div data-action="rename">Rename</div><div data-action="merge">Merge into another</div>';
  menu.style.display = 'block'; menu.style.left = x + 'px'; menu.style.top = y + 'px';
  menu.onclick = e => {
    menu.style.display = 'none';
    if (e.target.dataset.action === 'rename') renameRanch(orgName, ranchName);
    else if (e.target.dataset.action === 'merge') mergeRanchPrompt(orgName, ranchName);
  };
}

document.addEventListener('click', () => { document.getElementById('ctxMenu').style.display = 'none'; });

// ═══ Row/Ranch Operations ═══
function archiveOrg(name) {
  if (!confirm('Archive "' + name + '"?')) return;
  const store = getStore();
  archiveItem('org', name, store.orgs[name]);
  delete store.orgs[name];
  saveStore(store);
  if (currentOrg === name) { currentOrg = null; currentRanch = null; }
  renderMergeSidebar(); renderMergeMain();
}

function archiveRanch(orgName, ranchName) {
  if (!confirm('Archive "' + ranchName + '"?')) return;
  const store = getStore();
  archiveItem('ranch', ranchName, store.orgs[orgName].ranches[ranchName], orgName);
  delete store.orgs[orgName].ranches[ranchName];
  saveStore(store);
  if (currentRanch === ranchName) currentRanch = null;
  renderMergeSidebar(); renderMergeMain();
}

function archiveRow(ri) {
  if (!confirm('Archive this row?')) return;
  const store = getStore();
  const ranch = store.orgs[currentOrg]?.ranches?.[currentRanch];
  if (!ranch) return;
  const row = ranch.rows.splice(ri, 1)[0];
  archiveItem('row', currentRanch, { row, progress: ranch.progress[ri] || [] }, currentOrg);
  const np = {};
  Object.keys(ranch.progress).forEach(k => { const ki = parseInt(k); if (ki < ri) np[ki] = ranch.progress[ki]; else if (ki > ri) np[ki-1] = ranch.progress[ki]; });
  ranch.progress = np;
  saveStore(store); renderMergeMain(); renderMergeSidebar();
}

function openMoveRowModal(ri) {
  const store = getStore();
  const select = document.getElementById('moveRowTarget');
  select.innerHTML = '';
  Object.keys(store.orgs).sort().forEach(orgName => {
    Object.keys(store.orgs[orgName].ranches).sort().forEach(rName => {
      if (orgName === currentOrg && rName === currentRanch) return;
      select.innerHTML += '<option value="' + escHtml(orgName) + '|||' + escHtml(rName) + '">' + escHtml(orgName) + ' > ' + escHtml(rName) + '</option>';
    });
  });
  if (!select.options.length) { alert('No other ranches.'); return; }
  document.getElementById('moveRowModal').classList.add('show');
  document.getElementById('moveRowCancel').onclick = () => document.getElementById('moveRowModal').classList.remove('show');
  document.getElementById('moveRowConfirm').onclick = () => {
    const [tOrg, tRanch] = select.value.split('|||');
    const store = getStore();
    const src = store.orgs[currentOrg].ranches[currentRanch];
    const dst = store.orgs[tOrg].ranches[tRanch];
    const row = src.rows.splice(ri, 1)[0];
    row[RANCH_COL] = tRanch;
    dst.rows.push(row); dst.progress[dst.rows.length - 1] = src.progress[ri] || [];
    const np = {};
    Object.keys(src.progress).forEach(k => { const ki = parseInt(k); if (ki < ri) np[ki] = src.progress[ki]; else if (ki > ri) np[ki-1] = src.progress[ki]; });
    src.progress = np;
    saveStore(store);
    document.getElementById('moveRowModal').classList.remove('show');
    renderMergeSidebar(); renderMergeMain();
  };
}

function renameRanch(orgName, ranchName) {
  const newName = prompt('Rename "' + ranchName + '" to:', ranchName);
  if (!newName || !newName.trim() || newName.trim() === ranchName) return;
  const store = getStore();
  if (store.orgs[orgName].ranches[newName.trim()]) { alert('Name exists. Use Merge.'); return; }
  store.orgs[orgName].ranches[newName.trim()] = store.orgs[orgName].ranches[ranchName];
  delete store.orgs[orgName].ranches[ranchName];
  store.orgs[orgName].ranches[newName.trim()].rows.forEach(r => { r[RANCH_COL] = newName.trim(); });
  saveStore(store);
  if (currentRanch === ranchName) currentRanch = newName.trim();
  renderMergeSidebar(); renderMergeMain();
}

function mergeRanchPrompt(orgName, srcName) {
  const store = getStore();
  const others = Object.keys(store.orgs[orgName].ranches).filter(r => r !== srcName).sort();
  if (!others.length) { alert('No other ranches.'); return; }
  const target = prompt('Merge "' + srcName + '" into which ranch?\n\n' + others.join('\n'));
  if (!target || !others.includes(target)) return;
  const src = store.orgs[orgName].ranches[srcName];
  const dst = store.orgs[orgName].ranches[target];
  src.rows.forEach((row, i) => { row[RANCH_COL] = target; dst.rows.push(row); dst.progress[dst.rows.length - 1] = src.progress[i] || []; });
  delete store.orgs[orgName].ranches[srcName];
  saveStore(store);
  if (currentRanch === srcName) currentRanch = target;
  renderMergeSidebar(); renderMergeMain();
}

// ═══ Init events (called once) ═══
initMergeToolbarEvents();
