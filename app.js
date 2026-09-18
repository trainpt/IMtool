// ═══════════════════════════════════════════
// Table Merge
//
// Upload (or paste) a sheet, map its columns onto the eight canonical ones,
// then click cells as you copy them across so you can see what's left. Data
// lives in this module for the session only — there are no projects, ranches
// or saved progress, and nothing is written to the browser.
// ═══════════════════════════════════════════

const HEADERS = ['Location','Sites','Crop','Variety','Location Type','Planted Date','Acreage','Plant Count'];
const RANCH_COL = 1;
const ACREAGE_COL = 6;
const PLANTCOUNT_COL = 7;
const SKIP_FOR_DONE = new Set([RANCH_COL]);
const NUMERIC_COLS = new Set([ACREAGE_COL, PLANTCOUNT_COL]);

// ─── Session state ───
let mergeRows = [];        // [[...8 cols]]
let mergeProgress = {};    // rowIdx → [colIdx…] marked copied
let mergeMapping = null;   // [srcColIdx per canonical column], -1 = skip
let mergeSrcHeaders = [];  // the uploaded file's own header row
let mergeFileName = '';
let copiedCells = {};
let sortCol = -1, sortAsc = true;

function initMergePage() { renderMergeMain(); }

// ═══ Main Table ═══
function renderMergeMain() {
  const thead = document.getElementById('mergeThead');
  const tbody = document.getElementById('mergeTbody');
  const empty = document.getElementById('mergeEmpty');
  const bc = document.getElementById('mergeBreadcrumb');
  if (!thead) return;
  thead.innerHTML = ''; tbody.innerHTML = ''; copiedCells = {};

  const hasRows = mergeRows.length > 0;
  ['mergeDoneAllBtn','mergeResetBtn','mergeClearBtn'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = hasRows ? '' : 'none';
  });

  if (!hasRows) {
    bc.textContent = '';
    empty.style.display = ''; hideStatsBar();
    updateMergeMapping();
    return;
  }
  bc.textContent = mergeFileName ? mergeFileName + ' · ' + mergeRows.length + ' rows' : mergeRows.length + ' rows';
  empty.style.display = 'none';

  // Cells that carry nothing to copy (None / 0 / blank) count as done up front.
  mergeRows.forEach((row, ri) => {
    const prog = new Set(mergeProgress[ri] || []);
    row.forEach((cell, ci) => {
      if (cell === 'None' || cell === '0' || cell === '') prog.add(ci);
    });
    mergeProgress[ri] = [...prog];
  });

  HEADERS.forEach((h, ci) => {
    const th = document.createElement('th');
    let arrow = '<span style="margin-left:3px;font-size:8px;opacity:0.3;">▲</span>';
    if (sortCol === ci) { arrow = '<span style="margin-left:3px;font-size:8px;">' + (sortAsc ? '▲' : '▼') + '</span>'; th.classList.add('sort-active'); }
    th.innerHTML = escHtml(h) + arrow;
    th.addEventListener('click', () => { if (sortCol === ci) sortAsc = !sortAsc; else { sortCol = ci; sortAsc = true; } renderMergeMain(); });
    thead.appendChild(th);
  });

  const indices = mergeRows.map((_, i) => i);
  if (sortCol >= 0) {
    indices.sort((a, b) => {
      let va = mergeRows[a][sortCol] || '', vb = mergeRows[b][sortCol] || '';
      if (NUMERIC_COLS.has(sortCol)) { va = parseFloat(va) || 0; vb = parseFloat(vb) || 0; return sortAsc ? va - vb : vb - va; }
      return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
  }

  indices.forEach(ri => {
    const row = mergeRows[ri];
    copiedCells[ri] = new Set(mergeProgress[ri] || []);
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
        if (clickCount !== 1) return;
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
      });
      tr.appendChild(td);
    });

    tr.addEventListener('contextmenu', e => { e.preventDefault(); showRowCtxMenu(e.clientX, e.clientY, ri); });
    checkMergeRowDone(tr, ri, row, nonEmpty);
    tbody.appendChild(tr);
  });

  updateMergeMapping();
  renderMergeStats();
}

function checkMergeRowDone(tr, ri, row, nonEmpty) {
  const done = [...copiedCells[ri]].filter(i => row[i] !== '' && !SKIP_FOR_DONE.has(i)).length;
  const blockCell = tr.querySelector('td');
  if (done >= nonEmpty && nonEmpty > 0) {
    tr.classList.add('row-done');
    if (!blockCell.querySelector('.badge')) blockCell.insertAdjacentHTML('beforeend', '<span class="badge">✓ Done</span>');
  } else {
    tr.classList.remove('row-done');
    const badge = blockCell?.querySelector('.badge'); if (badge) badge.remove();
  }
}

function saveMergeProgress() {
  const obj = {};
  for (const key in copiedCells) obj[key] = [...copiedCells[key]];
  mergeProgress = obj;
  renderMergeStats();
}

function renderMergeStats() {
  if (!mergeRows.length) { hideStatsBar(); return; }
  let total = mergeRows.length, done = 0;
  mergeRows.forEach((row, ri) => {
    const req = row.reduce((cols, c, ci) => { if (c !== '' && !SKIP_FOR_DONE.has(ci)) cols.push(ci); return cols; }, []);
    const d = new Set(mergeProgress[ri] || []);
    if (req.every(ci => d.has(ci))) done++;
  });
  showStatsBar(total, done);
}

// ═══ Toolbar ═══
function initMergeToolbarEvents() {
  document.getElementById('mergeDoneAllBtn').addEventListener('click', () => {
    mergeRows.forEach((row, ri) => {
      mergeProgress[ri] = row.map((_, ci) => ci).filter(ci => row[ci] !== '');
    });
    renderMergeMain();
  });

  document.getElementById('mergeResetBtn').addEventListener('click', () => {
    if (!confirm('Reset all copy progress?')) return;
    mergeProgress = {};
    renderMergeMain();
  });

  document.getElementById('mergeClearBtn').addEventListener('click', () => {
    if (!confirm('Clear all ' + mergeRows.length + ' rows?')) return;
    mergeRows = []; mergeProgress = {}; mergeFileName = '';
    sortCol = -1; sortAsc = true;
    renderMergeMain();
  });

  // Header mapping
  document.getElementById('mergeHeaderBtn').addEventListener('click', () => {
    const box = document.getElementById('mergeHeaderSetup');
    box.style.display = box.style.display === 'none' ? '' : 'none';
  });
  document.getElementById('mergeHeaderCancel').addEventListener('click', () => { document.getElementById('mergeHeaderSetup').style.display = 'none'; });
  document.getElementById('mergeHeaderConfirm').addEventListener('click', () => {
    const text = document.getElementById('mergeHeaderArea').value.trim();
    if (!text) return;
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) return;
    document.getElementById('mergeHeaderSetup').style.display = 'none';
    document.getElementById('mergeHeaderArea').value = '';
    showMergeColumnMapper(lines[0].split('\t').map(h => h.trim()));
  });

  // Paste
  document.getElementById('mergePasteBtn').addEventListener('click', () => {
    const box = document.getElementById('mergePasteBox');
    box.style.display = box.style.display === 'none' ? '' : 'none';
  });
  document.getElementById('mergePasteCancel').addEventListener('click', () => { document.getElementById('mergePasteBox').style.display = 'none'; });
  document.getElementById('mergePasteConfirm').addEventListener('click', () => {
    const text = document.getElementById('mergePasteArea').value.trim();
    if (!text) return;
    if (!mergeMapping) { alert('Set up header mapping first (Headers), or upload a file.'); return; }
    const parsed = text.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t').map(c => c.trim()));
    document.getElementById('mergePasteBox').style.display = 'none';
    document.getElementById('mergePasteArea').value = '';
    importMergeRows(parsed, mergeMapping);
  });

  // Upload
  document.getElementById('mergeFileInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = '';
    try {
      const result = await readUploadedFile(file);
      if (!result.sheets.length) { alert('No sheets found in that file.'); return; }
      mergeFileName = file.name;
      if (result.sheets.length > 1) {
        const names = result.sheets.map((s, i) => (i + 1) + '. ' + s.name).join('\n');
        const pick = prompt('This file has ' + result.sheets.length + ' sheets. Which one?\n\n' + names, '1');
        const idx = Math.max(0, Math.min(result.sheets.length - 1, (parseInt(pick, 10) || 1) - 1));
        processSheet([result.sheets[idx].headers, ...result.sheets[idx].rows]);
      } else {
        processSheet([result.sheets[0].headers, ...result.sheets[0].rows]);
      }
    } catch (err) { alert('Error: ' + err.message); }
  });

  // Ctrl+V straight onto the page
  document.addEventListener('paste', e => {
    if (!document.getElementById('page-merge').classList.contains('active')) return;
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    if (!mergeMapping) return;
    e.preventDefault();
    let parsed = null;
    const html = e.clipboardData.getData('text/html');
    if (html) {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const rows = doc.querySelectorAll('tr');
      if (rows.length) {
        parsed = [];
        rows.forEach(tr => { const cells = []; tr.querySelectorAll('td,th').forEach(c => cells.push(c.textContent.trim())); if (cells.length) parsed.push(cells); });
      }
    }
    if (!parsed) {
      const text = e.clipboardData.getData('text/plain');
      if (text) parsed = text.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t').map(c => c.trim()));
    }
    if (parsed && parsed.length) importMergeRows(parsed, mergeMapping);
  });
}

// ═══ Column Mapping ═══
function showMergeColumnMapper(fileHeaders) {
  const mapper = document.getElementById('mergeColumnMapper');
  let html = '<div class="panel-box"><h3 style="font-size:14px;margin-bottom:10px;">Map columns</h3>' +
    '<p class="text-muted small" style="margin-bottom:8px;">Headers: ' + escHtml(fileHeaders.join(', ')) + '</p>';
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
    mergeMapping = mapping;
    mergeSrcHeaders = fileHeaders;
    mapper.style.display = 'none';
    if (pendingSheetRows) { const rows = pendingSheetRows; pendingSheetRows = null; importMergeRows(rows, mapping); }
    else updateMergeMapping();
  });
  document.getElementById('mapperCancel').addEventListener('click', () => { mapper.style.display = 'none'; pendingSheetRows = null; });
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
  if (!el) return;
  if (mergeMapping) {
    const fh = mergeSrcHeaders.filter(h => h).slice(0, 6);
    const extra = mergeSrcHeaders.filter(h => h).length > 6 ? ' +' + (mergeSrcHeaders.filter(h => h).length - 6) + ' more' : '';
    el.innerHTML = '<span class="text-muted small">Mapped: ' + escHtml(fh.join(', ')) + extra +
      ' <a href="#" onclick="clearMergeMapping();return false;" style="color:var(--red);">Reset</a></span>';
  } else {
    el.innerHTML = '<span style="color:var(--amber);font-size:11px;">No headers set</span>';
  }
}

function clearMergeMapping() {
  if (!confirm('Clear header mapping?')) return;
  mergeMapping = null; mergeSrcHeaders = [];
  updateMergeMapping();
}

// ═══ Import ═══
let pendingSheetRows = null;   // data rows waiting on the user to confirm a mapping

function processSheet(sheetRows) {
  const filtered = sheetRows.filter(r => r.some(c => String(c).trim() !== ''));
  if (filtered.length < 2) { alert('No data rows.'); return; }
  if (mergeMapping) {
    importMergeRows(filtered.slice(1), mergeMapping);
  } else {
    // Hold the rows so the mapper can import them the moment it's confirmed.
    pendingSheetRows = filtered.slice(1);
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

  // Acreage and Plant Count get swapped often enough to be worth correcting:
  // acreage above plant count is almost always the two columns reversed.
  newRows.forEach(row => {
    const a = parseFloat(row[ACREAGE_COL]), p = parseFloat(row[PLANTCOUNT_COL]);
    if (!isNaN(a) && !isNaN(p) && a > 0 && p > 0 && a > p) { row[ACREAGE_COL] = row[PLANTCOUNT_COL]; row[PLANTCOUNT_COL] = String(a); }
  });

  const startIdx = mergeRows.length;
  newRows.forEach((row, i) => {
    mergeRows.push(row);
    mergeProgress[startIdx + i] = [];
  });
  sortCol = -1; sortAsc = true;
  renderMergeMain();
}

// ═══ Row context menu ═══
function showRowCtxMenu(x, y, ri) {
  const menu = document.getElementById('ctxMenu');
  menu.innerHTML = '<div data-action="delete">Remove row</div>';
  menu.style.display = 'block'; menu.style.left = x + 'px'; menu.style.top = y + 'px';
  menu.onclick = e => {
    menu.style.display = 'none';
    if (e.target.dataset.action === 'delete') removeMergeRow(ri);
  };
}

document.addEventListener('click', () => {
  const m = document.getElementById('ctxMenu');
  if (m) m.style.display = 'none';
});

function removeMergeRow(ri) {
  if (!confirm('Remove this row?')) return;
  mergeRows.splice(ri, 1);
  // Progress is keyed by row index, so everything after the gap shifts down.
  const np = {};
  Object.keys(mergeProgress).forEach(k => {
    const ki = parseInt(k);
    if (ki < ri) np[ki] = mergeProgress[ki];
    else if (ki > ri) np[ki - 1] = mergeProgress[ki];
  });
  mergeProgress = np;
  renderMergeMain();
}

// ═══ Init events (called once) ═══
initMergeToolbarEvents();
