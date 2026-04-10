// Data Progress Tracker - multi-sheet tracker with column reorder, resize, hide
(function() {

const TRK_STORAGE  = "dpt_touched_v1";
const TRK_FLAGS    = "dpt_flags_v1";
const TRK_NOTES    = "dpt_notes_v1";
const TRK_BACKUP   = "dpt_backup_v1";
const TRK_SHEETS   = "dpt_sheets_v2";
const TRK_SHEETS_FULL = "dpt_sheets_full_v1";
const TRK_ACTIVE   = "dpt_active_v1";

let trkStorageOk = true;
function sGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch(e) { trkStorageOk = false; return null; } }
function sSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch(e) { trkStorageOk = false; } }
try { localStorage.setItem("__trktest__","1"); if (localStorage.getItem("__trktest__")!=="1") throw 0; localStorage.removeItem("__trktest__"); } catch(e) { trkStorageOk = false; }

let trkTouched = sGet(TRK_STORAGE) || {};
let trkFlags   = sGet(TRK_FLAGS) || {};
let trkNotes   = sGet(TRK_NOTES) || {};

// Multi-sheet: { key: { name, headers, rows, colOrder, hiddenCols, colWidths } }
let trkSheets = {};
let trkActiveSheet = null;

// Active sheet refs
let trkHeaders = [];
let trkRows = [];
let trkColOrder = [];
let trkHiddenCols = new Set();  // set of original col indices
let trkColWidths = {};          // { colIdx: px }
let trkActiveRow = null;
let trkSortCol = -1, trkSortAsc = true;
let trkBuilt = false;

const $ = id => document.getElementById(id);
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// ══════════════════════════════════════════
// ── ETA ──
// ══════════════════════════════════════════
let trkETAData = {};
function getETA() {
  if (!trkActiveSheet) return { timestamps: [], lastCount: 0, onBreak: false };
  if (!trkETAData[trkActiveSheet]) trkETAData[trkActiveSheet] = { timestamps: [], lastCount: 0, onBreak: false };
  return trkETAData[trkActiveSheet];
}
const TRK_BREAK = 30000;
function trkFormatETA(sec) {
  if (sec < 60) return '< 1 min remaining';
  if (sec < 3600) return `~${Math.ceil(sec / 60)} min remaining`;
  const h = Math.floor(sec / 3600), m = Math.ceil((sec % 3600) / 60);
  return m > 0 ? `~${h}h ${m}m remaining` : `~${h}h remaining`;
}
function trkActiveElapsed() {
  const eta = getETA(); let active = 0;
  for (let i = 1; i < eta.timestamps.length; i++) {
    const gap = eta.timestamps[i].time - eta.timestamps[i - 1].time;
    if (gap < TRK_BREAK) active += gap;
  }
  return active / 1000;
}
setInterval(() => {
  const eta = getETA();
  if (eta.timestamps.length === 0) return;
  const idle = Date.now() - eta.timestamps[eta.timestamps.length - 1].time;
  if (idle >= TRK_BREAK && !eta.onBreak) {
    eta.onBreak = true;
    const el = $('trk-eta');
    if (el) { el.textContent = 'Paused — ETA resumes when you continue'; el.style.color = '#b07800'; }
  }
}, 5000);

// ══════════════════════════════════════════
// ── Save / Backup ──
// ══════════════════════════════════════════
function trkSave() { sSet(TRK_STORAGE, trkTouched); sSet(TRK_FLAGS, trkFlags); sSet(TRK_NOTES, trkNotes); trkSaveSheetsMeta(); }
function trkBackup() { sSet(TRK_BACKUP, { touched: trkTouched, flags: trkFlags, notes: trkNotes, time: Date.now() }); }
function trkSaveSheetsMeta() {
  const meta = {};
  const full = {};
  Object.keys(trkSheets).forEach(k => {
    const s = trkSheets[k];
    const hidden = s.hiddenCols instanceof Set ? [...s.hiddenCols] : [...(s.hiddenCols || [])];
    const rids = trkGetRids(s.rows || []);
    meta[k] = { name: s.name, colOrder: s.colOrder, hiddenCols: hidden, colWidths: s.colWidths || {}, rids };
    full[k] = {
      name: s.name,
      headers: s.headers,
      rows: s.rows,
      rids,
      colOrder: s.colOrder,
      hiddenCols: hidden,
      colWidths: s.colWidths || {},
      orgName: s.orgName || ''
    };
  });
  sSet(TRK_SHEETS, meta);
  sSet(TRK_SHEETS_FULL, full);
  if (trkActiveSheet) sSet(TRK_ACTIVE, trkActiveSheet);
}

function trkRestoreFromLocal() {
  const full = sGet(TRK_SHEETS_FULL);
  if (!full || typeof full !== 'object') return false;
  const keys = Object.keys(full);
  if (keys.length === 0) return false;

  // Raw restore only. trkTouched/trkFlags/trkNotes are already loaded from
  // localStorage at module init and are the freshest source of progress truth —
  // do NOT merge from the session store here (it may hold older values that
  // would overwrite more recent marks).
  keys.forEach(k => {
    const s = full[k];
    if (!s || !Array.isArray(s.headers) || !Array.isArray(s.rows)) return;
    const rows = s.rows.map(r => Array.isArray(r) ? r.slice() : r);
    trkTagRows(rows, s.rids);
    trkSheets[k] = {
      name: s.name,
      headers: s.headers,
      rows: rows,
      colOrder: s.colOrder || s.headers.map((_, i) => i),
      hiddenCols: new Set(s.hiddenCols || []),
      colWidths: s.colWidths || {},
      orgName: s.orgName || ''
    };
  });

  const savedActive = sGet(TRK_ACTIVE);
  const activeKey = (savedActive && trkSheets[savedActive]) ? savedActive : Object.keys(trkSheets)[0];
  if (activeKey) {
    trkSwitchToSheet(activeKey);
    return true;
  }
  return false;
}
function makeSheetKey(name) { return 'sheet_' + String(name).trim().replace(/[^a-zA-Z0-9]/g, '_').toLowerCase(); }

// Remove any stale progress entries (touched/flags/notes) keyed to `sheetKey`.
// Used when a fresh load is happening — prevents old marks from a previous
// sheet with the same key from auto-striking the new data.
function trkPurgeSheetProgress(sheetKey) {
  if (!sheetKey) return;
  const prefix = `trk-${sheetKey}-`;
  Object.keys(trkTouched).forEach(k => { if (k.startsWith(prefix)) delete trkTouched[k]; });
  Object.keys(trkFlags).forEach(k => { if (k.startsWith(prefix)) delete trkFlags[k]; });
  Object.keys(trkNotes).forEach(k => { if (k.startsWith(prefix)) delete trkNotes[k]; });
}

// Assign stable IDs to rows so progress follows them through sorts
function trkTagRows(rows, savedRids) {
  rows.forEach((r, i) => {
    if (savedRids && savedRids[i] !== undefined) {
      r._rid = savedRids[i];
    } else if (r._rid === undefined) {
      r._rid = i;
    }
  });
  return rows;
}
function trkGetRids(rows) { return rows.map(r => r._rid !== undefined ? r._rid : 0); }
function trkRowId(row) { return row._rid !== undefined ? row._rid : 0; }
function trkCellUid(sheetKey, row, ci) { return `trk-${sheetKey}-${trkRowId(row)}-${ci}`; }

// ══════════════════════════════════════════
// ── Persistent session save/load (unified data store) ──
// ══════════════════════════════════════════

function trkGetSessionStore() {
  const store = typeof getStore === 'function' ? getStore() : {};
  if (!store.trackerSaves) store.trackerSaves = {};
  return store;
}

// Save all current tracker sheets + progress to an org
function trkSaveSession(orgName) {
  if (!orgName) return;
  orgName = orgName.trim();
  const store = trkGetSessionStore();
  if (!store.trackerSaves) store.trackerSaves = {};
  if (!store.trackerSaves[orgName]) store.trackerSaves[orgName] = {};

  // Save each loaded sheet
  Object.keys(trkSheets).forEach(key => {
    const s = trkSheets[key];
    const hidden = s.hiddenCols instanceof Set ? [...s.hiddenCols] : [...(s.hiddenCols || [])];
    // Collect progress for this sheet
    const prefix = `trk-${key}-`;
    const touched = {}, flags = {}, notes = {};
    Object.keys(trkTouched).forEach(k => { if (k.startsWith(prefix)) touched[k] = trkTouched[k]; });
    Object.keys(trkFlags).forEach(k => { if (k.startsWith(prefix)) flags[k] = trkFlags[k]; });
    Object.keys(trkNotes).forEach(k => { if (k.startsWith(prefix)) notes[k] = trkNotes[k]; });

    store.trackerSaves[orgName][key] = {
      name: s.name, headers: s.headers, rows: s.rows, rids: trkGetRids(s.rows),
      colOrder: s.colOrder, hiddenCols: hidden, colWidths: s.colWidths || {},
      touched, flags, notes,
      savedAt: new Date().toISOString()
    };
  });

  if (typeof saveStore === 'function') saveStore(store);
}

// Load a saved session from an org (returns array of sheet objects)
function trkLoadSession(orgName) {
  const store = trkGetSessionStore();
  const sessions = store.trackerSaves?.[orgName];
  if (!sessions) return [];
  return Object.keys(sessions).map(key => ({ key, ...sessions[key] }));
}

// Delete a saved session sheet
function trkDeleteSavedSheet(orgName, key) {
  const store = trkGetSessionStore();
  if (store.trackerSaves?.[orgName]?.[key]) {
    delete store.trackerSaves[orgName][key];
    if (Object.keys(store.trackerSaves[orgName]).length === 0) delete store.trackerSaves[orgName];
    if (typeof saveStore === 'function') saveStore(store);
  }
}

// Get all org names that have tracker saves
function trkSavedOrgNames() {
  const store = trkGetSessionStore();
  const names = new Set();
  Object.keys(store.trackerSaves || {}).forEach(n => {
    if (Object.keys(store.trackerSaves[n]).length > 0) names.add(n);
  });
  return [...names].sort();
}

// Populate the saved sessions UI on setup screen
function trkPopulateSavedSessions() {
  const orgSel = $('trk-saved-org');
  const list = $('trk-saved-list');

  // Populate org dropdown
  const savedOrgs = trkSavedOrgNames();
  const allOrgs = typeof getAllOrgNames === 'function' ? getAllOrgNames() : [];
  const orgs = [...new Set([...savedOrgs, ...allOrgs])].sort();
  orgSel.innerHTML = '<option value="">-- Select organization --</option>';
  orgs.forEach(n => {
    const hasSaves = savedOrgs.includes(n);
    orgSel.innerHTML += '<option value="' + n.replace(/"/g, '&quot;') + '">' + esc(n) + (hasSaves ? ' ●' : '') + '</option>';
  });

  // Also populate the "save to org" dropdown
  const saveSel = $('trk-save-org');
  if (saveSel) {
    saveSel.innerHTML = '<option value="">-- Select --</option>';
    orgs.forEach(n => {
      saveSel.innerHTML += '<option value="' + n.replace(/"/g, '&quot;') + '">' + esc(n) + '</option>';
    });
  }

  list.innerHTML = '';
  trkRenderSavedList();
}

function trkRenderSavedList() {
  const orgName = $('trk-saved-org').value;
  const list = $('trk-saved-list');
  if (!orgName) { list.innerHTML = '<div class="trk-saved-empty">Select an organization to see saved sessions.</div>'; return; }

  const sessions = trkLoadSession(orgName);
  if (sessions.length === 0) {
    list.innerHTML = '<div class="trk-saved-empty">No saved sessions for this organization.</div>';
    return;
  }

  list.innerHTML = '';
  sessions.forEach(s => {
    // Calculate progress
    const hidden = new Set(s.hiddenCols || []);
    let done = 0, total = 0;
    (s.rows || []).forEach((row, ri) => {
      (s.colOrder || s.headers.map((_, i) => i)).forEach(ci => {
        if (hidden.has(ci)) return;
        total++;
        if (s.touched?.[`trk-${s.key}-${ri}-${ci}`]) done++;
      });
    });
    const pct = total > 0 ? Math.round(done / total * 100) : 0;
    const date = s.savedAt ? new Date(s.savedAt).toLocaleDateString() : '';

    const item = document.createElement('div');
    item.className = 'trk-saved-item';
    item.innerHTML =
      '<span class="trk-saved-item-name">' + esc(s.name) + '</span>' +
      '<span class="trk-saved-item-meta">' + (s.rows?.length || 0) + ' rows · ' + date + '</span>' +
      '<span class="trk-saved-item-pct' + (pct >= 100 ? ' complete' : '') + '">' + pct + '%</span>' +
      '<span class="trk-saved-item-del" title="Delete saved session">×</span>';

    item.querySelector('.trk-saved-item-name').addEventListener('click', () => trkRestoreSession(orgName, s));
    item.querySelector('.trk-saved-item-del').addEventListener('click', e => {
      e.stopPropagation();
      if (!confirm('Delete saved session "' + s.name + '"?')) return;
      trkDeleteSavedSheet(orgName, s.key);
      trkRenderSavedList();
    });
    list.appendChild(item);
  });

  // "Load All" button if multiple sessions
  if (sessions.length > 1) {
    const loadAll = document.createElement('button');
    loadAll.className = 'btn btn-primary btn-sm'; loadAll.style.marginTop = '8px';
    loadAll.textContent = 'Load All (' + sessions.length + ' sheets)';
    loadAll.addEventListener('click', () => {
      sessions.forEach(s => trkRestoreSession(orgName, s, true));
      if (Object.keys(trkSheets).length > 0) {
        trkSwitchToSheet(Object.keys(trkSheets)[0]);
        $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
      }
    });
    list.appendChild(loadAll);
  }
}

function trkRestoreSession(orgName, session, silent) {
  const key = session.key;
  trkSheets[key] = {
    name: session.name,
    headers: session.headers,
    rows: trkTagRows(session.rows.map(r => r.map(c => String(c))), session.rids),
    colOrder: session.colOrder || session.headers.map((_, i) => i),
    hiddenCols: new Set(session.hiddenCols || []),
    colWidths: session.colWidths || {},
    orgName: orgName
  };
  // Restore progress
  if (session.touched) Object.assign(trkTouched, session.touched);
  if (session.flags) Object.assign(trkFlags, session.flags);
  if (session.notes) Object.assign(trkNotes, session.notes);
  trkSaveSheetsMeta();
  trkSave();

  if (!silent) {
    trkSwitchToSheet(key);
    $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
  }
}

$('trk-saved-org').addEventListener('change', trkRenderSavedList);

// ══════════════════════════════════════════
// ── File upload / parse ──
// ══════════════════════════════════════════
$('trk-file-upload').addEventListener('change', e => {
  if (!e.target.files[0]) return;
  const file = e.target.files[0];
  $('trk-file-name').textContent = file.name;
  const sheetName = file.name.replace(/\.\w+$/, '');
  trkReadFile(file, sheetName);
});

function trkReadFile(file, fallbackName) {
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();
  if (ext === 'csv') {
    reader.onload = e => {
      const csv = trkParseCSV(e.target.result);
      const allRows = [csv.headers, ...csv.rows].map(r => r.map(c => String(c)));
      const parsed = trkSmartParse(allRows);
      trkShowSetup(fallbackName, parsed.headers, parsed.rows);
    };
    reader.readAsText(file);
  } else {
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellStyles: true });
      if (wb.SheetNames.length === 1) {
        trkLoadXlsSheet(wb, wb.SheetNames[0], (h, r) => trkShowSetup(wb.SheetNames[0], h, r));
      } else {
        trkSheetPicker(wb);
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

function trkLoadXlsSheet(wb, name, cb) {
  const ws = wb.Sheets[name];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Detect yellow-highlighted rows by reading cell background colors
  const yellowRows = new Set();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= Math.min(range.s.c + 3, range.e.c); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell || !cell.s) continue;
      const fill = cell.s.fgColor || cell.s.bgColor || (cell.s.fill && (cell.s.fill.fgColor || cell.s.fill.bgColor));
      if (fill) {
        const rgb = fill.rgb || '';
        const theme = fill.theme;
        // Detect yellow: RGB hex starting with FF or FE for yellow shades, or common yellow theme
        if (/^FF(FF|EF|EB|D9|E5|C7)/i.test(rgb) || /^FFFF/i.test(rgb) || rgb === 'FFFFFF00') {
          yellowRows.add(r);
          break;
        }
      }
    }
  }

  const nonBlank = data.filter(r => r.some(c => String(c).trim() !== ''));
  if (nonBlank.length < 2) { alert('Sheet "' + name + '" has no data.'); return; }

  // Map original row indices to track which nonBlank rows were yellow
  let origIdx = 0;
  const yellowFlags = [];
  data.forEach((row, dataIdx) => {
    if (row.some(c => String(c).trim() !== '')) {
      yellowFlags.push(yellowRows.has(dataIdx));
      origIdx++;
    }
  });

  const allRows = nonBlank.map(r => r.map(String));
  const { headers, rows } = trkSmartParse(allRows, yellowFlags);
  cb(headers, rows);
}

function trkSheetPicker(wb) {
  let picker = document.getElementById('trk-sheet-picker');
  if (!picker) { picker = document.createElement('div'); picker.id = 'trk-sheet-picker'; picker.className = 'modal-overlay show'; document.body.appendChild(picker); }
  const sheetInfo = [];
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const rows = data.filter(r => r.some(c => String(c).trim() !== ''));
    sheetInfo.push({ name, count: Math.max(0, rows.length - 1) });
  });
  let html = '<div class="modal" style="max-width:480px;"><h3>Select Sheets</h3>';
  html += '<div style="margin-bottom:10px;display:flex;align-items:center;gap:8px;"><label style="font-size:12px;cursor:pointer;display:flex;align-items:center;gap:4px;"><input type="checkbox" id="trk-picker-all"> <strong>Select All</strong></label><span class="text-muted small" id="trk-picker-count">0 selected</span></div>';
  html += '<div style="max-height:360px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius);margin-bottom:12px;">';
  sheetInfo.forEach((s, i) => {
    html += '<label class="trk-picker-row" style="display:flex;align-items:center;gap:8px;padding:8px 12px;cursor:pointer;border-bottom:1px solid var(--border-light);"><input type="checkbox" class="trk-picker-cb" data-idx="' + i + '" data-sheet="' + s.name.replace(/"/g, '&quot;') + '"><span style="flex:1;">' + esc(s.name) + '</span><span class="text-muted small">' + s.count + ' rows</span></label>';
  });
  html += '</div>';
  html += '<label style="display:flex;align-items:center;gap:6px;font-size:12px;margin-bottom:12px;cursor:pointer;"><input type="checkbox" id="trk-picker-hide-blank" checked> Auto-hide blank columns</label>';
  html += '<div class="modal-actions"><button class="btn btn-ghost" id="trk-picker-cancel">Cancel</button><button class="btn btn-primary" id="trk-picker-load" disabled>Load Selected</button></div></div>';
  picker.innerHTML = html;
  picker.style.display = 'flex';

  const checkboxes = picker.querySelectorAll('.trk-picker-cb');
  const selectAll = picker.querySelector('#trk-picker-all');
  const loadBtn = picker.querySelector('#trk-picker-load');
  const countEl = picker.querySelector('#trk-picker-count');
  function updateCount() {
    const n = picker.querySelectorAll('.trk-picker-cb:checked').length;
    countEl.textContent = n + ' selected'; loadBtn.disabled = n === 0;
    selectAll.checked = n === checkboxes.length; selectAll.indeterminate = n > 0 && n < checkboxes.length;
  }
  selectAll.addEventListener('change', () => { checkboxes.forEach(cb => { cb.checked = selectAll.checked; }); updateCount(); });
  checkboxes.forEach(cb => cb.addEventListener('change', updateCount));
  picker.querySelector('#trk-picker-cancel').addEventListener('click', () => { picker.style.display = 'none'; });

  loadBtn.addEventListener('click', () => {
    const selected = Array.from(picker.querySelectorAll('.trk-picker-cb:checked')).map(cb => cb.dataset.sheet);
    const hideBlank = picker.querySelector('#trk-picker-hide-blank').checked;
    picker.style.display = 'none';
    if (selected.length === 0) return;

    // Stage all selected sheets for the build step
    trkStagedMulti = [];
    selected.forEach(name => {
      trkLoadXlsSheet(wb, name, (headers, rows) => {
        const hidden = new Set();
        if (hideBlank) findBlankCols(headers, rows).forEach(ci => hidden.add(ci));
        trkStagedMulti.push({ name, headers, rows: trkTagRows(rows.map(r => r.map(c => String(c)))), hiddenCols: hidden });
      });
    });

    // Show setup screen with name/org fields; prefill name from file
    const combinedName = selected.length === 1 ? selected[0] : selected.join(', ');
    $('trk-staging-name').value = combinedName;
    $('trk-preview-wrap').style.display = '';
    $('trk-setup-actions').style.display = 'flex';
    $('trk-col-list').style.display = 'none'; // no column reorder for multi
    $('trk-preview-count').textContent = '(' + trkStagedMulti.length + ' sheets, ' + trkStagedMulti.reduce((s, sh) => s + sh.rows.length, 0) + ' total rows)';

    // Show a summary preview instead of a table
    let html = '<thead><tr><th>Sheet</th><th>Rows</th><th>Columns</th></tr></thead><tbody>';
    trkStagedMulti.forEach(sh => {
      html += '<tr><td>' + esc(sh.name) + '</td><td>' + sh.rows.length + '</td><td>' + sh.headers.length + '</td></tr>';
    });
    html += '</tbody>';
    $('trk-preview-table').innerHTML = html;

    // Hide blank prompt for multi
    $('trk-blank-prompt').style.display = 'none';
  });
}

function trkSmartParse(allRows, yellowFlags) {
  if (allRows.length < 2) return { headers: allRows[0] || [], rows: allRows.slice(1) };
  const colCount = Math.max(...allRows.map(r => r.length));
  function headerScore(row) {
    let filled = 0;
    for (let i = 0; i < colCount; i++) { const v = String(row[i] || '').trim(); if (v && v.length > 0 && v.length < 80) filled++; }
    return filled;
  }
  const searchLimit = Math.min(6, allRows.length);
  let bestIdx = 0, bestScore = 0;
  for (let i = 0; i < searchLimit; i++) { const score = headerScore(allRows[i]); if (score > bestScore) { bestScore = score; bestIdx = i; } }
  const headers = allRows[bestIdx].map(c => String(c).trim());
  const dataRows = allRows.slice(bestIdx + 1);
  // Compute average cell length for "normal" data rows (skip first 3 after header for calibration)
  // Then filter out description/instruction rows that have unusually long average cell text.
  function avgCellLen(row) {
    const filled = row.filter(c => String(c).trim() !== '');
    if (filled.length === 0) return 0;
    return filled.reduce((sum, c) => sum + String(c).length, 0) / filled.length;
  }

  // Calculate median avg cell length from rows 3+ onward (the "real" data)
  const sampleRows = dataRows.slice(3, 30);
  let dataMedian = 15; // reasonable default
  if (sampleRows.length > 0) {
    const avgs = sampleRows.map(avgCellLen).filter(a => a > 0).sort((a, b) => a - b);
    if (avgs.length > 0) dataMedian = avgs[Math.floor(avgs.length / 2)];
  }
  // Threshold: description rows have avg cell length much higher than data rows
  const descThreshold = Math.max(40, dataMedian * 3);

  // First pass: find description rows (high avg cell length in first 4 rows after header)
  const descFlags = [];
  for (let ri = 0; ri < Math.min(4, dataRows.length); ri++) {
    const row = dataRows[ri];
    const nonEmpty = row.filter(c => String(c).trim() !== '');
    descFlags[ri] = (nonEmpty.length >= 2 && avgCellLen(row) > descThreshold);
  }
  // Detect example rows: use yellow highlighting if available, otherwise heuristic
  let lastDescIdx = -1;
  for (let ri = 0; ri < descFlags.length; ri++) { if (descFlags[ri]) lastDescIdx = ri; }

  // Build set of yellow data row indices (offset by bestIdx+1 since dataRows starts after header)
  const yellowDataRows = new Set();
  if (yellowFlags) {
    dataRows.forEach((row, ri) => {
      const origFlagIdx = bestIdx + 1 + ri;
      if (yellowFlags[origFlagIdx]) yellowDataRows.add(ri);
    });
  }

  // Fallback heuristic if no yellow info: check if first row after descriptions has high avg cell length
  let exampleIdx = -1;
  if (yellowDataRows.size === 0 && lastDescIdx >= 0 && lastDescIdx + 1 < dataRows.length) {
    const candidate = dataRows[lastDescIdx + 1];
    const candidateAvg = avgCellLen(candidate);
    if (candidateAvg > dataMedian * 1.5 && candidateAvg > 20) {
      exampleIdx = lastDescIdx + 1;
    }
  }

  // Determine the median fill count for real data rows (how many cells are non-empty)
  // Sample from middle of data to avoid header/footer contamination
  const startSample = exampleIdx >= 0 ? exampleIdx + 1 : 0;
  const fillCounts = dataRows.slice(startSample, startSample + 30)
    .map(r => r.filter(c => String(c).trim() !== '').length)
    .filter(n => n > 1)
    .sort((a, b) => a - b);
  const medianFill = fillCounts.length > 0 ? fillCounts[Math.floor(fillCounts.length / 2)] : colCount;
  // Rows must fill at least half as many columns as typical data rows (minimum 2)
  const minFillForData = Math.max(2, Math.ceil(medianFill * 0.5));

  const filtered = dataRows.filter((row, ri) => {
    const nonEmpty = row.filter(c => String(c).trim() !== '');
    // Skip completely empty rows
    if (nonEmpty.length === 0) return false;
    // Skip rows with only 1 non-empty cell that's long (instructions)
    if (nonEmpty.length <= 1 && String(nonEmpty[0] || '').length > 100) return false;
    // Skip rows where any cell is > 200 chars
    if (row.some(c => String(c).length > 200)) return false;
    // Skip rows that are a duplicate of the header
    const isHeaderDupe = row.every((c, i) => String(c).trim() === (headers[i] || ''));
    if (isHeaderDupe && nonEmpty.length > 2) return false;
    // Skip description rows
    if (ri < 4 && descFlags[ri]) return false;
    // Skip yellow-highlighted rows (examples in spreadsheet)
    if (yellowDataRows.has(ri)) return false;
    // Skip the example row right after descriptions (fallback heuristic)
    if (ri === exampleIdx && ri < 5) return false;
    // Skip rows that don't fill enough columns compared to real data
    // (catches footer content like phone numbers, addresses, app names, notes)
    if (nonEmpty.length < minFillForData) return false;
    return true;
  });
  // Strip trailing empty columns (columns where header is empty and all data is empty)
  let lastUsedCol = headers.length - 1;
  while (lastUsedCol > 0) {
    const h = (headers[lastUsedCol] || '').trim();
    const hasData = filtered.some(r => (r[lastUsedCol] || '').trim() !== '');
    if (h || hasData) break;
    lastUsedCol--;
  }
  const trimmedHeaders = headers.slice(0, lastUsedCol + 1);
  const trimmedRows = filtered.map(r => r.slice(0, lastUsedCol + 1));

  return { headers: trimmedHeaders, rows: trimmedRows };
}

function trkParseCSV(text) {
  const headers = [], rows = [], fields = [];
  let cur = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) { if (ch === '"' && text[i+1] === '"') { cur += '"'; i++; } else if (ch === '"') inQ = false; else cur += ch; }
    else { if (ch === '"') inQ = true; else if (ch === ',') { fields.push(cur); cur = ''; } else if (ch === '\n' || (ch === '\r' && text[i+1] === '\n')) { fields.push(cur); cur = ''; if (ch === '\r') i++; if (!headers.length) headers.push(...fields.splice(0)); else rows.push(fields.splice(0)); } else cur += ch; }
  }
  if (cur || fields.length) { fields.push(cur); if (!headers.length) headers.push(...fields); else rows.push(fields.splice(0)); }
  return { headers, rows };
}

// ── Blank column detection ──
function findBlankCols(headers, rows) {
  const blank = [];
  for (let ci = 0; ci < headers.length; ci++) {
    const allEmpty = rows.every(r => !String(r[ci] || '').trim());
    if (allEmpty) blank.push(ci);
  }
  return blank;
}

// ══════════════════════════════════════════
// ── Multi-sheet management ──
// ══════════════════════════════════════════
function trkRemoveSheet(key) {
  if (!trkSheets[key]) return;
  if (!confirm('Remove "' + trkSheets[key].name + '" from the tracker?')) return;
  delete trkSheets[key]; delete trkETAData[key]; trkSaveSheetsMeta();
  if (trkActiveSheet === key) {
    const remaining = Object.keys(trkSheets);
    if (remaining.length > 0) trkSwitchToSheet(remaining[0]);
    else { trkActiveSheet = null; $('trk-main').style.display = 'none'; $('trk-setup').style.display = ''; }
  } else trkRenderSheetTabs();
}

function trkSwitchToSheet(key) {
  if (!trkSheets[key]) return;
  trkActiveSheet = key;
  const sheet = trkSheets[key];
  trkHeaders = sheet.headers;
  trkRows = trkTagRows(sheet.rows);
  trkColOrder = sheet.colOrder;
  trkHiddenCols = sheet.hiddenCols instanceof Set ? sheet.hiddenCols : new Set(sheet.hiddenCols || []);
  trkColWidths = sheet.colWidths || {};
  trkActiveRow = null; trkSortCol = -1; trkSortAsc = true;
  trkRenderSheetTabs();
  trkBuildReview();
}

function trkRenderSheetTabs() {
  const bar = $('trk-sheet-tabs'); bar.innerHTML = '';
  Object.keys(trkSheets).forEach(key => {
    const sheet = trkSheets[key];
    const tab = document.createElement('div');
    tab.className = 'trk-sheet-tab' + (key === trkActiveSheet ? ' active' : '');
    tab.onclick = () => { if (key !== trkActiveSheet) trkSwitchToSheet(key); };
    const label = document.createElement('span'); label.className = 'trk-sheet-tab-label';
    label.textContent = sheet.name; label.title = sheet.name + ' (' + sheet.rows.length + ' rows)';
    const mini = document.createElement('span'); mini.className = 'trk-sheet-tab-pct';
    const { done, total } = trkSheetProgress(key);
    const pct = total > 0 ? Math.round(done / total * 100) : 0;
    mini.textContent = pct + '%'; if (pct >= 100) mini.classList.add('complete');
    const close = document.createElement('span'); close.className = 'trk-sheet-tab-close';
    close.textContent = '×'; close.title = 'Remove sheet';
    close.onclick = (e) => { e.stopPropagation(); trkRemoveSheet(key); };
    tab.appendChild(label); tab.appendChild(mini); tab.appendChild(close); bar.appendChild(tab);
  });
  const addBtn = document.createElement('div'); addBtn.className = 'trk-sheet-tab trk-sheet-tab-add';
  addBtn.textContent = '+ Add Sheet'; addBtn.title = 'Add another sheet to this org';
  addBtn.onclick = () => trkOpenQuickAdd();
  bar.appendChild(addBtn);
}

// ══════════════════════════════════════════
// ── Quick Add Sheet modal ──
// ══════════════════════════════════════════
function trkCurrentOrg() {
  return Object.values(trkSheets).find(s => s.orgName)?.orgName || '';
}

function trkOpenQuickAdd() {
  const org = trkCurrentOrg();
  const overlay = $('trk-quickadd-overlay');
  if (!overlay) return;
  $('trk-quickadd-org').textContent = org ? '→ ' + org : '';
  trkRenderQuickAddList();
  $('trk-quickadd-file').value = '';
  overlay.classList.add('show');
}

function trkCloseQuickAdd() {
  const overlay = $('trk-quickadd-overlay');
  if (overlay) overlay.classList.remove('show');
}

function trkRenderQuickAddList() {
  const list = $('trk-quickadd-list');
  if (!list) return;
  const names = typeof getSheetNames === 'function' ? getSheetNames() : [];
  if (names.length === 0) {
    list.innerHTML = '<div class="trk-saved-empty">No imported sheets available. Use Import in the top bar or upload below.</div>';
    return;
  }
  list.innerHTML = '';
  names.forEach(n => {
    const sheet = typeof getSheet === 'function' ? getSheet(n) : null;
    const key = makeSheetKey(n);
    const alreadyLoaded = !!trkSheets[key];
    const item = document.createElement('div');
    item.className = 'trk-saved-item';
    item.style.cursor = alreadyLoaded ? 'default' : 'pointer';
    item.style.opacity = alreadyLoaded ? '0.5' : '1';
    item.innerHTML =
      '<span class="trk-saved-item-name">' + esc(n) + '</span>' +
      '<span class="trk-saved-item-meta">' + (sheet ? sheet.rowCount + ' rows' : '') + '</span>' +
      (alreadyLoaded ? '<span class="trk-saved-item-pct">Already added</span>' : '<span class="trk-saved-item-pct" style="background:var(--accent-soft);color:var(--accent);">Add</span>');
    if (!alreadyLoaded && sheet) {
      item.addEventListener('click', () => {
        // Route through the exact same entry point the full setup uses.
        trkCloseQuickAdd();
        trkGotoSetupWithOrg();
        if (typeof window.trkLoadSheetData === 'function') {
          window.trkLoadSheetData(sheet.headers, sheet.rows, n);
        }
      });
    }
    list.appendChild(item);
  });
}

// Switch the tracker to the setup screen and prefill the current org so the
// user lands on the same staging flow as a normal manual load.
function trkGotoSetupWithOrg() {
  $('trk-main').style.display = 'none';
  $('trk-setup').style.display = '';
  trkPopulateSavedSessions();
  const org = trkCurrentOrg();
  if (org) {
    const orgSel = $('trk-save-org');
    if (orgSel) orgSel.value = org;
    const newOrgInput = $('trk-save-new-org');
    if (newOrgInput) newOrgInput.value = '';
  }
}

// Wire up buttons once
(function trkBindQuickAdd() {
  const overlay = document.getElementById('trk-quickadd-overlay');
  if (!overlay) return;
  const closeBtn = document.getElementById('trk-quickadd-close');
  if (closeBtn) closeBtn.addEventListener('click', trkCloseQuickAdd);
  overlay.addEventListener('click', e => { if (e.target === overlay) trkCloseQuickAdd(); });

  const fileInput = document.getElementById('trk-quickadd-file');
  if (fileInput) {
    fileInput.addEventListener('change', e => {
      const file = e.target.files[0];
      if (!file) return;
      // Hand off to the existing file pipeline — identical parsing, preview,
      // blank-column detection, etc. Org is prefilled by trkGotoSetupWithOrg().
      trkCloseQuickAdd();
      trkGotoSetupWithOrg();
      $('trk-file-name').textContent = file.name;
      const sheetName = file.name.replace(/\.\w+$/, '');
      trkReadFile(file, sheetName);
      e.target.value = '';
    });
  }

  const advBtn = document.getElementById('trk-quickadd-advanced');
  if (advBtn) {
    advBtn.addEventListener('click', () => {
      trkCloseQuickAdd();
      trkGotoSetupWithOrg();
    });
  }
})();

function trkSheetProgress(key) {
  const sheet = trkSheets[key]; if (!sheet) return { done: 0, total: 0 };
  const hidden = sheet.hiddenCols instanceof Set ? sheet.hiddenCols : new Set(sheet.hiddenCols || []);
  let done = 0, total = 0;
  sheet.rows.forEach((row, ri) => {
    sheet.colOrder.forEach(ci => {
      if (hidden.has(ci)) return;
      total++;
      if (trkTouched[trkCellUid(key, row, ci)] === true) done++;
    });
  });
  return { done, total };
}

// ══════════════════════════════════════════
// ── Setup screen ──
// ══════════════════════════════════════════
let trkStagingHeaders = [], trkStagingRows = [], trkStagingColOrder = [], trkStagingName = '';
let trkStagedMulti = null; // for multi-sheet picker: [{ name, headers, rows, hiddenCols }]

function trkShowSetup(name, headers, rows) {
  trkStagedMulti = null; // clear multi-stage if switching to single
  trkStagingName = name || 'Untitled';
  trkStagingHeaders = headers; trkStagingRows = rows;
  trkStagingColOrder = headers.map((_, i) => i);
  $('trk-preview-wrap').style.display = '';
  $('trk-setup-actions').style.display = 'flex';
  $('trk-col-list').style.display = '';
  $('trk-preview-count').textContent = '(' + rows.length + ' rows)';
  $('trk-staging-name').value = trkStagingName;

  // Blank column detection prompt
  const blankCols = findBlankCols(headers, rows);
  const blankWrap = $('trk-blank-prompt');
  if (blankCols.length > 0) {
    blankWrap.style.display = '';
    blankWrap.innerHTML = '<label style="display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer;"><input type="checkbox" id="trk-hide-blank-cb" checked> Auto-hide ' + blankCols.length + ' empty column' + (blankCols.length > 1 ? 's' : '') + ' <span class="text-muted small">(' + blankCols.map(ci => headers[ci] || 'Col ' + (ci+1)).join(', ') + ')</span></label>';
  } else {
    blankWrap.style.display = 'none'; blankWrap.innerHTML = '';
  }

  trkRenderColList(); trkRenderPreview();
}

function trkRenderColList() {
  const list = $('trk-col-list-inner'); list.innerHTML = '';
  trkStagingColOrder.forEach((ci, pos) => {
    const chip = document.createElement('div'); chip.className = 'trk-col-chip'; chip.draggable = true; chip.dataset.pos = pos;
    chip.innerHTML = '<span class="trk-col-grip">&#9776;</span> ' + esc(trkStagingHeaders[ci]);
    chip.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', pos); chip.classList.add('dragging'); });
    chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
    chip.addEventListener('dragover', e => { e.preventDefault(); chip.classList.add('drag-over'); });
    chip.addEventListener('dragleave', () => chip.classList.remove('drag-over'));
    chip.addEventListener('drop', e => {
      e.preventDefault(); chip.classList.remove('drag-over');
      const from = parseInt(e.dataTransfer.getData('text/plain')); if (from === pos) return;
      const item = trkStagingColOrder.splice(from, 1)[0]; trkStagingColOrder.splice(pos, 0, item);
      trkRenderColList(); trkRenderPreview();
    });
    list.appendChild(chip);
  });
}

function trkRenderPreview() {
  let html = '<thead><tr>';
  trkStagingColOrder.forEach(ci => { html += '<th>' + esc(trkStagingHeaders[ci]) + '</th>'; });
  html += '</tr></thead><tbody>';
  trkStagingRows.slice(0, 5).forEach(r => {
    html += '<tr>'; trkStagingColOrder.forEach(ci => { html += '<td>' + esc(r[ci] || '') + '</td>'; }); html += '</tr>';
  });
  html += '</tbody>';
  $('trk-preview-table').innerHTML = html;
}

$('trk-btn-build').addEventListener('click', () => {
  // Determine org — required
  const orgSel = $('trk-save-org');
  const newOrg = $('trk-save-new-org').value.trim();
  const orgName = newOrg || (orgSel ? orgSel.value : '') || '';
  if (!orgName) {
    alert('Please select or create an organization before starting.');
    (orgSel || $('trk-save-new-org')).focus();
    return;
  }

  // Load any existing saved sheets for this org first
  if (orgName) {
    const existingSessions = trkLoadSession(orgName);
    existingSessions.forEach(s => {
      if (!trkSheets[s.key]) trkRestoreSession(orgName, s, true);
    });
  }

  // Multi-sheet path
  if (trkStagedMulti && trkStagedMulti.length > 0) {
    // If only 1 sheet staged, use the renamed name from the input field
    const userRename = ($('trk-staging-name').value || '').trim();
    const existingSessionKeys = new Set(
      orgName ? trkLoadSession(orgName).map(s => s.key) : []
    );
    trkStagedMulti.forEach((sh, idx) => {
      const sheetName = (trkStagedMulti.length === 1 && userRename) ? userRename : sh.name;
      const key = makeSheetKey(sheetName);
      // Fresh load → wipe any stale progress tied to this key unless there's
      // an existing saved session we want to preserve.
      if (!existingSessionKeys.has(key)) trkPurgeSheetProgress(key);
      const savedMeta = sGet(TRK_SHEETS) || {};
      const saved = savedMeta[key];
      const colOrder = saved?.colOrder?.length === sh.headers.length ? saved.colOrder : sh.headers.map((_, i) => i);
      trkSheets[key] = { name: sheetName, headers: sh.headers, rows: trkTagRows(sh.rows), colOrder, hiddenCols: sh.hiddenCols || new Set(), colWidths: saved?.colWidths || {}, orgName };
    });
    const firstName = (trkStagedMulti.length === 1 && userRename) ? userRename : trkStagedMulti[0].name;
    const firstKey = makeSheetKey(firstName);
    trkSaveSheetsMeta();
    trkSwitchToSheet(trkSheets[firstKey] ? firstKey : Object.keys(trkSheets)[0]);
    $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
    if (orgName) trkSaveSession(orgName);
    trkStagedMulti = null;
    return;
  }

  // Single-sheet path
  if (!trkStagingHeaders.length || !trkStagingRows.length) { alert('Load data first'); return; }
  const name = ($('trk-staging-name').value || '').trim() || trkStagingName;
  const key = makeSheetKey(name);
  const hidden = new Set();
  const hideBlankCb = document.getElementById('trk-hide-blank-cb');
  if (hideBlankCb && hideBlankCb.checked) {
    findBlankCols(trkStagingHeaders, trkStagingRows).forEach(ci => hidden.add(ci));
  }

  // Before adding, load any existing saved sheets for this org that aren't already loaded
  const existingSessions = orgName ? trkLoadSession(orgName) : [];
  if (orgName) {
    existingSessions.forEach(s => {
      if (!trkSheets[s.key]) trkRestoreSession(orgName, s, true);
    });
  }

  // Fresh load → wipe stale progress tied to this key unless there's a saved
  // session for it we want to preserve.
  const hasSavedSession = existingSessions.some(s => s.key === key);
  if (!hasSavedSession) trkPurgeSheetProgress(key);

  trkSheets[key] = { name, headers: trkStagingHeaders, rows: trkTagRows(trkStagingRows.map(r => r.map(c => String(c)))), colOrder: [...trkStagingColOrder], hiddenCols: hidden, colWidths: {}, orgName };
  // Ensure all loaded sheets share the same org
  if (orgName) Object.values(trkSheets).forEach(s => { if (!s.orgName) s.orgName = orgName; });
  trkSaveSheetsMeta(); trkSwitchToSheet(key);
  $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
  trkStagingHeaders = []; trkStagingRows = []; trkStagingColOrder = [];

  // Auto-save all sheets to org
  if (orgName) trkSaveSession(orgName);
});

$('trk-btn-back').addEventListener('click', () => {
  $('trk-main').style.display = 'none';
  $('trk-setup').style.display = '';
  if (typeof trkInit === 'function') trkInit();
});

// ══════════════════════════════════════════
// ── Stats / Progress ──
// ══════════════════════════════════════════
function trkUpdateStat() {
  if (!trkActiveSheet) return;
  let count = 0, total = 0;
  document.querySelectorAll('#trk-tbody td[data-uid]').forEach(td => { total++; if (trkTouched[td.dataset.uid] === true) count++; });
  const pct = total > 0 ? (count / total * 100) : 0;
  const label = trkStorageOk ? 'Session Saved' : '⚠️ NOT PERSISTING';
  const eta = getETA(); const now = Date.now();
  if (count > eta.lastCount) { eta.onBreak = false; eta.timestamps.push({ time: now, count }); if (eta.timestamps.length > 60) eta.timestamps.shift(); }
  eta.lastCount = count;
  let etaText = ''; const remaining = total - count;
  if (remaining > 0 && eta.timestamps.length >= 2) {
    const elapsed = trkActiveElapsed(); const completed = eta.timestamps[eta.timestamps.length - 1].count - eta.timestamps[0].count;
    if (elapsed > 0 && completed > 0) etaText = trkFormatETA(remaining / (completed / elapsed));
  } else if (remaining > 0) etaText = 'Estimating time...';
  else etaText = 'Complete!';
  $('trk-stat').textContent = `✅ ${count} / ${total} | ${label}`;
  $('trk-bar').style.width = pct + '%'; $('trk-pct').textContent = pct.toFixed(1) + '%';
  $('trk-detail').textContent = `${count} / ${total} cells completed`;
  const etaEl = $('trk-eta'); etaEl.textContent = etaText; etaEl.style.color = eta.onBreak ? '#b07800' : '#666';
  trkUpdateTabPcts();
}

function trkUpdateTabPcts() {
  const tabs = $('trk-sheet-tabs'); if (!tabs) return;
  tabs.querySelectorAll('.trk-sheet-tab:not(.trk-sheet-tab-add)').forEach(tab => {
    const label = tab.querySelector('.trk-sheet-tab-label'); const pctEl = tab.querySelector('.trk-sheet-tab-pct');
    if (!label || !pctEl) return;
    const key = Object.keys(trkSheets).find(k => trkSheets[k].name === label.textContent); if (!key) return;
    const { done, total } = trkSheetProgress(key);
    const pct = total > 0 ? Math.round(done / total * 100) : 0;
    pctEl.textContent = pct + '%'; pct >= 100 ? pctEl.classList.add('complete') : pctEl.classList.remove('complete');
  });
}

// ══════════════════════════════════════════
// ── Cell interaction ──
// ══════════════════════════════════════════
function trkHandleCell(td, val, uid, tr, manual) {
  if (trkTouched[uid] === undefined) trkTouched[uid] = false;
  if (manual) {
    trkTouched[uid] = true;
    if (trkActiveRow && trkActiveRow !== tr) trkActiveRow.classList.remove('active-row');
    trkActiveRow = tr; tr.classList.add('active-row');
    // Mark last-clicked cell
    const prev = document.querySelector('#trk-tbody td.trk-last-click');
    if (prev) prev.classList.remove('trk-last-click');
    td.classList.add('trk-last-click');
    if (navigator.clipboard) navigator.clipboard.writeText(val);
    td.classList.add('copied'); setTimeout(() => td.classList.remove('copied'), 600);
    trkSave();
  }
  trkTouched[uid] ? td.classList.add('touched') : td.classList.remove('touched');
  trkUpdateRowBtn(tr); trkUpdateStat();
}
function trkUpdateRowBtn(tr) {
  const allDone = Array.from(tr.querySelectorAll('td[data-uid]')).every(el => trkTouched[el.dataset.uid] === true);
  const btn = tr.querySelector('.trk-row-btn');
  if (btn) allDone ? btn.classList.add('done') : btn.classList.remove('done');
  allDone ? tr.classList.add('trk-row-done') : tr.classList.remove('trk-row-done');
}
let trkLastToggledRowIdx = -1;

function trkToggleRow(tr, e) {
  const rows = Array.from($('trk-tbody').querySelectorAll('tr'));
  const thisIdx = rows.indexOf(tr);

  if (e && e.shiftKey && trkLastToggledRowIdx >= 0) {
    // Range toggle: from last toggled to this row
    const minR = Math.min(trkLastToggledRowIdx, thisIdx);
    const maxR = Math.max(trkLastToggledRowIdx, thisIdx);
    // Use the target state of the clicked row
    const tds = tr.querySelectorAll('td[data-uid]');
    const allDone = Array.from(tds).every(td => trkTouched[td.dataset.uid] === true);
    const target = !allDone;
    for (let i = minR; i <= maxR; i++) {
      const r = rows[i];
      r.querySelectorAll('td[data-uid]').forEach(td => {
        trkTouched[td.dataset.uid] = target;
        target ? td.classList.add('touched') : td.classList.remove('touched');
      });
      trkUpdateRowBtn(r);
    }
    trkSave(); trkUpdateStat();
  } else {
    // Single row toggle
    const tds = tr.querySelectorAll('td[data-uid]');
    const allDone = Array.from(tds).every(td => trkTouched[td.dataset.uid] === true);
    const target = !allDone;
    tds.forEach(td => { trkTouched[td.dataset.uid] = target; target ? td.classList.add('touched') : td.classList.remove('touched'); });
    trkUpdateRowBtn(tr); trkSave(); trkUpdateStat();
  }
  trkLastToggledRowIdx = thisIdx;
}

// ══════════════════════════════════════════
// ── Context menu / Flags / Notes ──
// ══════════════════════════════════════════
let trkCtxTarget = null;
const trkCtxMenu = $('trk-ctx');
function trkApplyFlagNote(td, uid) {
  if (trkFlags[uid]) td.classList.add('flagged'); else td.classList.remove('flagged');
  trkNotes[uid] ? td.classList.add('has-note') : td.classList.remove('has-note');
  td.title = ''; // use custom tooltip only, not native
}
document.addEventListener('contextmenu', e => {
  const td = e.target.closest('#trk-tbody td[data-uid]'); if (!td) return;
  e.preventDefault(); trkCtxTarget = td;
  const totalTargets = trkSelected.size + (trkSelected.size > 0 && !trkSelected.has(trkCellKey(parseInt(td.dataset.ri), parseInt(td.dataset.ci))) ? 1 : 0) || 1;
  const suffix = totalTargets > 1 ? ' (' + totalTargets + ' cells)' : '';
  $('trk-ctx-flag').textContent = (trkFlags[td.dataset.uid] ? '✅ Unmark Red' : '🔴 Mark Red') + suffix;
  $('trk-ctx-note').textContent = 'Add Note' + suffix;
  $('trk-ctx-clear-note').style.display = trkNotes[td.dataset.uid] ? '' : 'none';
  trkCtxMenu.style.display = 'block';
  trkCtxMenu.style.left = Math.min(e.clientX, window.innerWidth - 170) + 'px';
  trkCtxMenu.style.top = Math.min(e.clientY, window.innerHeight - 100) + 'px';
});
document.addEventListener('click', () => { if (trkCtxMenu) trkCtxMenu.style.display = 'none'; });
$('trk-ctx-edit').addEventListener('click', () => {
  if (!trkCtxTarget) return;
  const ri = parseInt(trkCtxTarget.dataset.ri);
  const ci = parseInt(trkCtxTarget.dataset.ci);
  if (!isNaN(ri) && !isNaN(ci)) trkStartInlineEdit(trkCtxTarget, ri, ci);
  trkCtxMenu.style.display = 'none';
});
// Helper: get all target UIDs (selected cells + the right-clicked cell)
function trkCtxTargets() {
  const uids = new Set();
  // Always include the right-clicked cell
  if (trkCtxTarget) uids.add(trkCtxTarget.dataset.uid);
  // Include all selected cells
  if (trkSelected.size > 0) {
    trkGetSelectedUids().forEach(({ uid }) => uids.add(uid));
  }
  return [...uids];
}

function trkRefreshFlagNotes() {
  document.querySelectorAll('#trk-tbody td[data-uid]').forEach(td => trkApplyFlagNote(td, td.dataset.uid));
}

$('trk-ctx-flag').addEventListener('click', () => {
  const uids = trkCtxTargets();
  if (uids.length === 0) return;
  const anyUnflagged = uids.some(uid => !trkFlags[uid]);
  uids.forEach(uid => { trkFlags[uid] = anyUnflagged; });
  trkRefreshFlagNotes(); trkSave(); trkCtxMenu.style.display = 'none'; trkClearSelection();
});
const trkNoteOverlay = $('trk-note-overlay'), trkNoteText = $('trk-note-text');
$('trk-ctx-note').addEventListener('click', () => {
  const uids = trkCtxTargets();
  if (uids.length === 0) return;
  trkNoteText.value = trkNotes[uids[0]] || '';
  trkNoteText._targetUids = uids;
  trkNoteOverlay.classList.add('show'); trkNoteText.focus(); trkCtxMenu.style.display = 'none';
});
$('trk-ctx-clear-note').addEventListener('click', () => {
  const uids = trkCtxTargets();
  if (uids.length === 0) return;
  uids.forEach(uid => { delete trkNotes[uid]; });
  trkRefreshFlagNotes(); trkSave(); trkCtxMenu.style.display = 'none'; trkClearSelection();
});
$('trk-note-save').addEventListener('click', () => {
  const uids = trkNoteText._targetUids || (trkCtxTarget ? [trkCtxTarget.dataset.uid] : []);
  if (uids.length === 0) return;
  const v = trkNoteText.value.trim();
  uids.forEach(uid => { if (v) trkNotes[uid] = v; else delete trkNotes[uid]; });
  trkRefreshFlagNotes(); trkSave(); trkNoteOverlay.classList.remove('show'); trkClearSelection();
});
$('trk-note-cancel').addEventListener('click', () => trkNoteOverlay.classList.remove('show'));
trkNoteOverlay.addEventListener('click', e => { if (e.target === trkNoteOverlay) trkNoteOverlay.classList.remove('show'); });

const trkTooltip = $('trk-tooltip');
document.addEventListener('mouseover', e => { const td = e.target.closest('#trk-tbody td.has-note'); if (td && trkNotes[td.dataset.uid]) { trkTooltip.textContent = trkNotes[td.dataset.uid]; trkTooltip.style.display = 'block'; } });
document.addEventListener('mousemove', e => { if (trkTooltip.style.display === 'none') return; trkTooltip.style.left = Math.min(e.clientX + 14, window.innerWidth - trkTooltip.offsetWidth - 8) + 'px'; trkTooltip.style.top = Math.min(e.clientY + 14, window.innerHeight - trkTooltip.offsetHeight - 8) + 'px'; });
document.addEventListener('mouseout', e => { if (e.target.closest('#trk-tbody td.has-note')) trkTooltip.style.display = 'none'; });

// ══════════════════════════════════════════
// ── Sort ──
// ══════════════════════════════════════════
function trkSortRows(colPos) {
  const ci = trkVisibleCols()[colPos];
  if (trkSortCol === colPos) trkSortAsc = !trkSortAsc; else { trkSortCol = colPos; trkSortAsc = true; }
  trkRows.sort((a, b) => {
    let va = a[ci] || '', vb = b[ci] || '';
    const na = parseFloat(va.replace(/,/g, '')), nb = parseFloat(vb.replace(/,/g, ''));
    let cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : va.localeCompare(vb);
    return trkSortAsc ? cmp : -cmp;
  });
  trkRenderTable(); trkApplySearch($('trk-search').value);
  document.querySelectorAll('#trk-main-table thead th[data-col]').forEach(th => {
    th.classList.remove('sort-active'); const a = th.querySelector('.sort-arrow'); if (a) a.textContent = '▲▼';
  });
  const th = document.querySelector('#trk-main-table thead th[data-col="' + colPos + '"]');
  if (th) { th.classList.add('sort-active'); const a = th.querySelector('.sort-arrow'); if (a) a.textContent = trkSortAsc ? '▲' : '▼'; }
}

// ══════════════════════════════════════════
// ── Column visibility helpers ──
// ══════════════════════════════════════════
function trkVisibleCols() { return trkColOrder.filter(ci => !trkHiddenCols.has(ci)); }

function trkHideCol(ci) {
  trkHiddenCols.add(ci);
  if (trkSheets[trkActiveSheet]) { trkSheets[trkActiveSheet].hiddenCols = trkHiddenCols; }
  trkSaveSheetsMeta(); trkRenderTable(); trkUpdateStat(); trkRenderHiddenMenu();
}

function trkShowCol(ci) {
  trkHiddenCols.delete(ci);
  if (trkSheets[trkActiveSheet]) { trkSheets[trkActiveSheet].hiddenCols = trkHiddenCols; }
  trkSaveSheetsMeta(); trkRenderTable(); trkUpdateStat(); trkRenderHiddenMenu();
}

function trkShowAllCols() {
  trkHiddenCols.clear();
  if (trkSheets[trkActiveSheet]) { trkSheets[trkActiveSheet].hiddenCols = trkHiddenCols; }
  trkSaveSheetsMeta(); trkRenderTable(); trkUpdateStat(); trkRenderHiddenMenu();
}

function trkRenderHiddenMenu() {
  const wrap = $('trk-hidden-cols');
  if (trkHiddenCols.size === 0) { wrap.style.display = 'none'; return; }
  wrap.style.display = '';
  let html = '<span class="trk-hidden-label">Hidden:</span>';
  trkColOrder.forEach(ci => {
    if (!trkHiddenCols.has(ci)) return;
    html += '<span class="trk-hidden-chip" data-ci="' + ci + '">' + esc(trkHeaders[ci] || 'Col ' + (ci+1)) + ' <span class="trk-hidden-chip-x">×</span></span>';
  });
  html += '<span class="trk-hidden-chip trk-hidden-show-all">Show All</span>';
  wrap.innerHTML = html;
  wrap.querySelectorAll('.trk-hidden-chip[data-ci]').forEach(chip => {
    chip.addEventListener('click', () => trkShowCol(parseInt(chip.dataset.ci)));
  });
  wrap.querySelector('.trk-hidden-show-all').addEventListener('click', trkShowAllCols);
}

// ══════════════════════════════════════════
// ── Auto-fit columns to screen ──
// ══════════════════════════════════════════
function trkAutoFitCols() {
  const table = $('trk-main-table'); if (!table) return;
  const wrap = table.closest('.table-wrap'); if (!wrap) return;
  const visCols = trkVisibleCols();
  const available = wrap.clientWidth - 40; // minus checkbox col
  const perCol = Math.max(60, Math.floor(available / visCols.length));
  visCols.forEach(ci => { trkColWidths[ci] = perCol; });
  if (trkSheets[trkActiveSheet]) trkSheets[trkActiveSheet].colWidths = trkColWidths;
  trkSaveSheetsMeta();
  trkApplyColWidths();
}

function trkApplyColWidths() {
  const ths = $('trk-thead').querySelectorAll('th[data-ci]');
  ths.forEach(th => {
    const ci = parseInt(th.dataset.ci);
    const w = trkColWidths[ci];
    if (w) { th.style.width = w + 'px'; th.style.minWidth = w + 'px'; th.style.maxWidth = w + 'px'; }
    else { th.style.width = ''; th.style.minWidth = ''; th.style.maxWidth = ''; }
  });
  // Apply to body cells too
  $('trk-tbody').querySelectorAll('tr').forEach(tr => {
    const tds = tr.querySelectorAll('td[data-ci]');
    tds.forEach(td => {
      const ci = parseInt(td.dataset.ci);
      const w = trkColWidths[ci];
      if (w) { td.style.width = w + 'px'; td.style.minWidth = w + 'px'; td.style.maxWidth = w + 'px'; }
      else { td.style.width = ''; td.style.minWidth = ''; td.style.maxWidth = ''; }
    });
  });
}

// ══════════════════════════════════════════
// ── Column resize (drag border) ──
// ══════════════════════════════════════════
let trkResizing = null; // { th, ci, startX, startW }

function trkInitResize(th, ci) {
  const handle = document.createElement('div');
  handle.className = 'trk-resize-handle';
  th.style.position = 'relative';
  th.appendChild(handle);
  handle.addEventListener('mousedown', e => {
    e.preventDefault(); e.stopPropagation();
    trkResizing = { th, ci, startX: e.clientX, startW: th.offsetWidth };
    document.body.style.cursor = 'col-resize';
    document.body.classList.add('trk-resizing');
  });
}

document.addEventListener('mousemove', e => {
  if (!trkResizing) return;
  const diff = e.clientX - trkResizing.startX;
  const newW = Math.max(40, trkResizing.startW + diff);
  trkResizing.th.style.width = newW + 'px';
  trkResizing.th.style.minWidth = newW + 'px';
  trkResizing.th.style.maxWidth = newW + 'px';
  // Also resize body cells in this column
  const ci = trkResizing.ci;
  $('trk-tbody').querySelectorAll('td[data-ci="' + ci + '"]').forEach(td => {
    td.style.width = newW + 'px'; td.style.minWidth = newW + 'px'; td.style.maxWidth = newW + 'px';
  });
});

document.addEventListener('mouseup', () => {
  if (!trkResizing) return;
  const ci = trkResizing.ci;
  trkColWidths[ci] = trkResizing.th.offsetWidth;
  if (trkSheets[trkActiveSheet]) trkSheets[trkActiveSheet].colWidths = trkColWidths;
  trkSaveSheetsMeta();
  trkResizing = null;
  document.body.style.cursor = '';
  document.body.classList.remove('trk-resizing');
});

// ══════════════════════════════════════════
// ── Render table ──
// ══════════════════════════════════════════
function trkRenderTable() {
  if (!trkActiveSheet) return;
  const sheetKey = trkActiveSheet;
  const visCols = trkVisibleCols();

  // Header
  const thead = $('trk-thead');
  thead.innerHTML = '<th class="trk-action-th">#</th>';
  visCols.forEach((ci, pos) => {
    const th = document.createElement('th');
    th.dataset.col = pos;
    th.dataset.ci = ci;
    const w = trkColWidths[ci];
    if (w) { th.style.width = w + 'px'; th.style.minWidth = w + 'px'; th.style.maxWidth = w + 'px'; }

    // Label
    const labelSpan = document.createElement('span');
    labelSpan.className = 'trk-th-label';
    labelSpan.textContent = trkHeaders[ci];
    th.appendChild(labelSpan);

    // Sort arrow
    const arrow = document.createElement('span');
    arrow.className = 'sort-arrow'; arrow.textContent = '▲▼';
    th.appendChild(arrow);

    // Hide button
    const hideBtn = document.createElement('span');
    hideBtn.className = 'trk-th-hide'; hideBtn.textContent = '×'; hideBtn.title = 'Hide column';
    hideBtn.addEventListener('click', e => { e.stopPropagation(); trkHideCol(ci); });
    th.appendChild(hideBtn);

    th.addEventListener('click', e => {
      if (e.target.closest('.trk-th-hide') || e.target.closest('.trk-resize-handle')) return;
      trkSortRows(pos);
    });

    // Column drag reorder
    th.draggable = true;
    th.addEventListener('dragstart', e => {
      if (trkResizing) { e.preventDefault(); return; }
      e.dataTransfer.setData('text/plain', pos); th.classList.add('dragging');
    });
    th.addEventListener('dragend', () => th.classList.remove('dragging'));
    th.addEventListener('dragover', e => { e.preventDefault(); th.classList.add('drag-over'); });
    th.addEventListener('dragleave', () => th.classList.remove('drag-over'));
    th.addEventListener('drop', e => {
      e.preventDefault(); th.classList.remove('drag-over');
      const fromPos = parseInt(e.dataTransfer.getData('text/plain'));
      const toPos = pos; if (fromPos === toPos) return;
      // Map visible positions back to trkColOrder indices
      const fromCi = visCols[fromPos], toCi = visCols[toPos];
      const fromIdx = trkColOrder.indexOf(fromCi), toIdx = trkColOrder.indexOf(toCi);
      if (fromIdx < 0 || toIdx < 0) return;
      trkColOrder.splice(fromIdx, 1); trkColOrder.splice(toIdx, 0, fromCi);
      if (trkSheets[sheetKey]) trkSheets[sheetKey].colOrder = trkColOrder;
      trkSaveSheetsMeta(); trkRenderTable();
    });

    // Resize handle
    trkInitResize(th, ci);

    thead.appendChild(th);
  });

  // Body
  const tbody = $('trk-tbody'); tbody.innerHTML = '';
  trkRows.forEach((row, ri) => {
    const tr = document.createElement('tr');
    const actionTd = document.createElement('td');
    actionTd.className = 'trk-action-cell';
    const rowNum = document.createElement('span');
    rowNum.className = 'trk-row-num'; rowNum.textContent = ri + 1;
    const btn = document.createElement('button');
    btn.className = 'trk-row-btn'; btn.innerHTML = '&#10003;'; btn.onclick = (e) => trkToggleRow(tr, e);
    actionTd.appendChild(rowNum); actionTd.appendChild(btn); tr.appendChild(actionTd);

    visCols.forEach((ci, pos) => {
      const val = row[ci] || '';
      const td = document.createElement('td');
      td.textContent = val;
      td.dataset.uid = trkCellUid(sheetKey, row, ci);
      td.dataset.ci = ci;
      td.dataset.ri = ri;
      const w = trkColWidths[ci];
      if (w) { td.style.width = w + 'px'; td.style.minWidth = w + 'px'; td.style.maxWidth = w + 'px'; }
      td.style.background = ri % 2 === 0 ? '#ffffff' : '#f7f7f7';

      td._lastClick = 0; td._wasMarked = false;
      td.addEventListener('click', (e) => {
        // Ctrl/Shift+click = selection mode
        if (e.ctrlKey || e.metaKey || e.shiftKey) {
          e.preventDefault();
          trkToggleSelect(td, ri, ci, e);
          return;
        }
        // Normal click = copy + mark done (with double-click undo)
        const now = Date.now();
        if (now - td._lastClick < 300) {
          td._lastClick = 0;
          if (td._wasMarked) { trkTouched[td.dataset.uid] = false; td.classList.remove('touched'); trkUpdateRowBtn(tr); trkSave(); trkUpdateStat(); }
        } else {
          td._wasMarked = trkTouched[td.dataset.uid] === true;
          td._lastClick = now;
          trkClearSelection();
          trkHandleCell(td, val, td.dataset.uid, tr, true);
        }
        // Always remember this cell as anchor for shift+click
        trkLastSelectedTd = td;
      });
      // dblclick reserved for undo-mark-done (handled via _lastClick above)
      // Auto-mark empty cells as done
      if (!val.trim()) { trkTouched[td.dataset.uid] = true; }
      if (trkTouched[td.dataset.uid]) td.classList.add('touched');
      if (trkSelected.has(trkCellKey(ri, ci))) td.classList.add('trk-selected');
      trkApplyFlagNote(td, td.dataset.uid);
      tr.appendChild(td);
    });
    trkUpdateRowBtn(tr); tbody.appendChild(tr);
  });

  trkRenderHiddenMenu();
}

// ── Search ──
function trkApplySearch(query) {
  const q = query.trim().toLowerCase().replace(/\s+/g, '');
  $('trk-tbody').querySelectorAll('tr').forEach(tr => {
    let match = false;
    tr.querySelectorAll('td[data-uid]').forEach(td => {
      td.classList.remove('search-hit');
      if (q && td.textContent.toLowerCase().replace(/\s+/g, '').includes(q)) { td.classList.add('search-hit'); match = true; }
    });
    tr.style.display = q && !match ? 'none' : '';
  });
}
$('trk-search').addEventListener('input', e => trkApplySearch(e.target.value));

// ══════════════════════════════════════════
// ── Build review ──
// ══════════════════════════════════════════
function trkBuildReview() {
  const backup = sGet(TRK_BACKUP);
  if (backup) {
    const mc = Object.keys(trkTouched).length, bc = Object.keys(backup.touched || {}).length;
    if (mc < bc) {
      Object.entries(backup.touched || {}).forEach(([k, v]) => { if (trkTouched[k] === undefined) trkTouched[k] = v; });
      Object.entries(backup.flags || {}).forEach(([k, v]) => { if (trkFlags[k] === undefined) trkFlags[k] = v; });
      Object.entries(backup.notes || {}).forEach(([k, v]) => { if (trkNotes[k] === undefined) trkNotes[k] = v; });
    }
  }
  const sheet = trkSheets[trkActiveSheet];
  const visCols = trkVisibleCols();
  $('trk-breadcrumb').textContent = sheet.name + ' (' + trkRows.length + ' rows, ' + visCols.length + ' cols)';
  $('trk-search').value = '';
  trkRenderTable(); trkSave(); trkBackup(); trkUpdateStat();
  trkBuilt = true;
  if (!trkStorageOk) { $('trk-stat').textContent = '⚠️ Storage unavailable!'; $('trk-stat').style.color = '#ff6666'; }
}

// ── Reset ──
$('trk-btn-reset').addEventListener('click', () => {
  if (!trkActiveSheet) return;
  const sheet = trkSheets[trkActiveSheet];
  if (!confirm('Reset all progress for "' + sheet.name + '"? This cannot be undone.')) return;
  const prefix = `trk-${trkActiveSheet}-`;
  Object.keys(trkTouched).forEach(k => { if (k.startsWith(prefix)) delete trkTouched[k]; });
  Object.keys(trkFlags).forEach(k => { if (k.startsWith(prefix)) delete trkFlags[k]; });
  Object.keys(trkNotes).forEach(k => { if (k.startsWith(prefix)) delete trkNotes[k]; });
  trkSave(); trkBackup(); trkRenderTable(); trkUpdateStat();
});

// ── Auto-fit button ──
$('trk-btn-autofit').addEventListener('click', trkAutoFitCols);
$('trk-btn-undo').addEventListener('click', trkUndo);

// ── Selection action buttons ──
function trkGetSelectedUids() {
  const uids = [];
  trkSelected.forEach(key => {
    const [ri, ci] = key.split('-').map(Number);
    const row = trkRows[ri];
    if (row) uids.push({ ri, ci, row, uid: trkCellUid(trkActiveSheet, row, ci) });
  });
  return uids;
}

// Mark selected as done
$('trk-sel-done').addEventListener('click', (e) => {
  e.stopPropagation();
  const items = trkGetSelectedUids();
  if (items.length === 0) return;
  items.forEach(({ uid }) => { trkTouched[uid] = true; });
  trkSave(); trkRenderTable(); trkUpdateStat(); trkClearSelection();
});

// Mark selected as not done
$('trk-sel-undone').addEventListener('click', (e) => {
  e.stopPropagation();
  const items = trkGetSelectedUids();
  if (items.length === 0) return;
  items.forEach(({ uid }) => { trkTouched[uid] = false; });
  trkSave(); trkRenderTable(); trkUpdateStat(); trkClearSelection();
});

// Flag selected red (toggle)
$('trk-sel-flag').addEventListener('click', (e) => {
  e.stopPropagation();
  const items = trkGetSelectedUids();
  if (items.length === 0) return;
  const anyUnflagged = items.some(({ uid }) => !trkFlags[uid]);
  items.forEach(({ uid }) => { trkFlags[uid] = anyUnflagged; });
  trkSave();
  // Apply visually without full re-render to preserve selection view
  document.querySelectorAll('#trk-tbody td[data-uid]').forEach(td => {
    trkApplyFlagNote(td, td.dataset.uid);
  });
  trkClearSelection();
});

// Add note to all selected
$('trk-sel-note').addEventListener('click', (e) => {
  e.stopPropagation();
  const items = trkGetSelectedUids();
  if (items.length === 0) return;
  const note = prompt('Add note to ' + items.length + ' selected cell(s):', '');
  if (note === null) return;
  items.forEach(({ uid }) => {
    if (note.trim()) trkNotes[uid] = note.trim();
    else delete trkNotes[uid];
  });
  trkSave();
  document.querySelectorAll('#trk-tbody td[data-uid]').forEach(td => {
    trkApplyFlagNote(td, td.dataset.uid);
  });
  trkClearSelection();
});

// Edit selected cell values
$('trk-sel-edit').addEventListener('click', () => {
  if (trkSelected.size === 0) return;
  const newVal = prompt('Set ' + trkSelected.size + ' selected cell(s) to:', '');
  if (newVal === null) return;
  const changes = [];
  trkSelected.forEach(key => {
    const [ri, ci] = key.split('-').map(Number);
    const oldVal = trkRows[ri]?.[ci] || '';
    if (oldVal !== newVal) changes.push({ ri, ci, oldVal, newVal });
  });
  if (changes.length > 0) trkApplyEdits(changes, 'Edit ' + changes.length + ' selected cells');
  trkClearSelection();
});

// Escape to clear selection
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && trkSelected.size > 0) trkClearSelection();
  // Ctrl+Z for undo
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
    const active = document.activeElement;
    if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT')) return;
    if (trkUndoStack.length > 0) { e.preventDefault(); trkUndo(); }
  }
});

// ── Save button ──
$('trk-btn-save').addEventListener('click', () => {
  // Determine org to save to
  let orgName = null;
  // Check if any loaded sheet already has an org
  Object.values(trkSheets).forEach(s => { if (s.orgName) orgName = s.orgName; });
  if (!orgName) {
    // Prompt
    orgName = prompt('Save to which organization?');
    if (!orgName) return;
    // Assign to all loaded sheets
    Object.values(trkSheets).forEach(s => { s.orgName = orgName; });
  }
  trkSaveSession(orgName);
  const el = $('trk-stat');
  const prev = el.textContent;
  el.textContent = '💾 Saved to ' + orgName;
  setTimeout(() => { el.textContent = prev; }, 2000);
});

// ══════════════════════════════════════════
// ── Bulk Edit ──
// ══════════════════════════════════════════
const bulkOverlay = $('trk-bulk-overlay');

function trkPopulateBulkCols() {
  const visCols = trkVisibleCols();
  ['trk-bulk-col', 'trk-set-col', 'trk-clear-col'].forEach(id => {
    const sel = $(id); if (!sel) return;
    const hasAll = id === 'trk-bulk-col';
    sel.innerHTML = hasAll ? '<option value="all">All columns</option>' : '';
    visCols.forEach(ci => {
      sel.innerHTML += '<option value="' + ci + '">' + esc(trkHeaders[ci] || 'Col ' + (ci+1)) + '</option>';
    });
  });
}

// Open
$('trk-btn-bulkedit').addEventListener('click', () => {
  if (!trkActiveSheet) return;
  trkPopulateBulkCols();
  $('trk-bulk-find').value = ''; $('trk-bulk-replace-val').value = '';
  $('trk-bulk-preview').innerHTML = ''; $('trk-set-preview').innerHTML = ''; $('trk-clear-preview').innerHTML = '';
  $('trk-bulk-apply-btn').disabled = true; $('trk-set-apply-btn').disabled = true; $('trk-clear-apply-btn').disabled = true;
  bulkOverlay.classList.add('show');
});

// Close
['trk-bulk-close', 'trk-set-close', 'trk-clear-close'].forEach(id => {
  $(id).addEventListener('click', () => bulkOverlay.classList.remove('show'));
});
bulkOverlay.addEventListener('click', e => { if (e.target === bulkOverlay) bulkOverlay.classList.remove('show'); });

// Tab switching
document.querySelectorAll('.trk-bulk-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.trk-bulk-tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    $('trk-bulk-replace').style.display = tab.dataset.mode === 'replace' ? '' : 'none';
    $('trk-bulk-setcol').style.display = tab.dataset.mode === 'setcol' ? '' : 'none';
    $('trk-bulk-clear').style.display = tab.dataset.mode === 'clear' ? '' : 'none';
  });
});

// Populate unique value picker when a column is selected
function trkUpdateFindPicker() {
  const colVal = $('trk-bulk-col').value;
  const picker = $('trk-bulk-find-pick');
  if (colVal === 'all') {
    picker.style.display = 'none';
    return;
  }
  const ci = parseInt(colVal);
  const unique = new Map(); // value -> count
  trkRows.forEach(row => {
    const v = (row[ci] || '').trim();
    if (v) unique.set(v, (unique.get(v) || 0) + 1);
  });
  // Sort by count descending
  const sorted = [...unique.entries()].sort((a, b) => b[1] - a[1]);
  picker.innerHTML = '<option value="">-- Pick a value (' + sorted.length + ' unique) --</option>';
  sorted.forEach(([val, cnt]) => {
    picker.innerHTML += '<option value="' + val.replace(/"/g, '&quot;') + '">' + esc(val) + ' (' + cnt + ')</option>';
  });
  picker.style.display = '';
}

$('trk-bulk-col').addEventListener('change', trkUpdateFindPicker);

$('trk-bulk-find-pick').addEventListener('change', () => {
  const val = $('trk-bulk-find-pick').value;
  if (val) {
    $('trk-bulk-find').value = val;
    $('trk-bulk-exact').checked = true;
  }
});

// Show/hide conditional inputs
$('trk-set-filter').addEventListener('change', () => {
  $('trk-set-filter-val').style.display = $('trk-set-filter').value === 'equals' ? '' : 'none';
});
$('trk-clear-action').addEventListener('change', () => {
  $('trk-clear-text').style.display = $('trk-clear-action').value === 'remove-text' ? '' : 'none';
});

// ── Undo stack ──
const trkUndoStack = []; // [{ changes: [{ ri, ci, oldVal, newVal }], label }]
const TRK_MAX_UNDO = 30;

function trkApplyEdits(changes, label) {
  if (changes.length === 0) return;
  // Save reverse for undo
  const undo = changes.map(c => ({ ri: c.ri, ci: c.ci, oldVal: c.newVal, newVal: c.oldVal || c.oldVal }));
  trkUndoStack.push({ changes: changes.map(c => ({ ri: c.ri, ci: c.ci, oldVal: c.oldVal, newVal: c.newVal })), label: label || 'Edit' });
  if (trkUndoStack.length > TRK_MAX_UNDO) trkUndoStack.shift();
  trkUpdateUndoBtn();
  changes.forEach(({ ri, ci, newVal }) => { trkRows[ri][ci] = newVal; });
  if (trkSheets[trkActiveSheet]) trkSheets[trkActiveSheet].rows = trkRows;
  trkRenderTable(); trkUpdateStat();
}

function trkUndo() {
  if (trkUndoStack.length === 0) return;
  const last = trkUndoStack.pop();
  last.changes.forEach(({ ri, ci, oldVal }) => { trkRows[ri][ci] = oldVal; });
  if (trkSheets[trkActiveSheet]) trkSheets[trkActiveSheet].rows = trkRows;
  trkRenderTable(); trkUpdateStat(); trkUpdateUndoBtn();
}

function trkUpdateUndoBtn() {
  const btn = $('trk-btn-undo');
  if (!btn) return;
  if (trkUndoStack.length > 0) {
    btn.style.display = ''; btn.title = 'Undo: ' + trkUndoStack[trkUndoStack.length - 1].label;
  } else {
    btn.style.display = 'none';
  }
}

// ── Cell selection ──
let trkSelected = new Set(); // set of "ri-ci" keys
let trkLastSelectedTd = null; // for shift-click range

function trkCellKey(ri, ci) { return ri + '-' + ci; }

function trkClearSelection() {
  trkSelected.clear();
  document.querySelectorAll('#trk-tbody td.trk-selected').forEach(td => td.classList.remove('trk-selected'));
  trkUpdateSelectionInfo();
}

function trkToggleSelect(td, ri, ci, e) {
  const key = trkCellKey(ri, ci);

  if (e.shiftKey && trkLastSelectedTd) {
    // Range select: every cell between anchor and this cell
    const lastRi = parseInt(trkLastSelectedTd.dataset.ri);
    const lastCi = parseInt(trkLastSelectedTd.dataset.ci);
    const minR = Math.min(lastRi, ri), maxR = Math.max(lastRi, ri);
    const visCols = trkVisibleCols();
    const lastPos = visCols.indexOf(lastCi), thisPos = visCols.indexOf(ci);
    const minPos = Math.min(lastPos >= 0 ? lastPos : 0, thisPos >= 0 ? thisPos : 0);
    const maxPos = Math.max(lastPos >= 0 ? lastPos : visCols.length - 1, thisPos >= 0 ? thisPos : visCols.length - 1);
    const selectedCols = visCols.slice(minPos, maxPos + 1);
    // Add all cells in the rectangle (don't clear existing selection)
    for (let r = minR; r <= maxR; r++) {
      selectedCols.forEach(c => { trkSelected.add(trkCellKey(r, c)); });
    }
    trkApplySelectionClasses();
  } else if (e.ctrlKey || e.metaKey) {
    // Toggle individual cell, keep existing selection
    if (trkSelected.has(key)) trkSelected.delete(key);
    else trkSelected.add(key);
    td.classList.toggle('trk-selected');
    trkLastSelectedTd = td; // update anchor for next shift+click
  } else {
    // Plain click with no modifier shouldn't reach here (handled in main click)
    trkClearSelection();
    trkSelected.add(key);
    td.classList.add('trk-selected');
    trkLastSelectedTd = td;
  }

  trkUpdateSelectionInfo();
}

function trkApplySelectionClasses() {
  document.querySelectorAll('#trk-tbody td[data-uid]').forEach(td => {
    const ri = parseInt(td.dataset.ri);
    const ci = parseInt(td.dataset.ci);
    if (isNaN(ri) || isNaN(ci)) return;
    trkSelected.has(trkCellKey(ri, ci)) ? td.classList.add('trk-selected') : td.classList.remove('trk-selected');
  });
  trkUpdateSelectionInfo();
}

function trkUpdateSelectionInfo() {
  const info = $('trk-sel-info');
  const actions = $('trk-sel-actions');
  if (!info) return;
  if (trkSelected.size > 0) {
    info.style.display = '';
    info.textContent = trkSelected.size + ' cell' + (trkSelected.size > 1 ? 's' : '') + ' selected';
    if (actions) actions.style.display = '';
  } else {
    info.style.display = 'none';
    if (actions) actions.style.display = 'none';
  }
}

// ── Inline cell edit (double-click) ──
function trkStartInlineEdit(td, ri, ci) {
  if (td.querySelector('input')) return; // already editing
  const oldVal = trkRows[ri][ci] || '';
  const input = document.createElement('input');
  input.type = 'text'; input.value = oldVal;
  input.className = 'trk-inline-edit';
  td.textContent = '';
  td.appendChild(input);
  input.focus(); input.select();

  function commit() {
    const newVal = input.value;
    td.textContent = newVal;
    if (newVal !== oldVal) {
      trkApplyEdits([{ ri, ci, oldVal, newVal }], 'Edit cell R' + (ri+1));
    }
  }
  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
    if (e.key === 'Escape') { input.value = oldVal; input.blur(); }
  });
}

// ── Find & Replace ──
let bulkPendingChanges = [];

$('trk-bulk-find-btn').addEventListener('click', () => {
  const find = $('trk-bulk-find').value;
  if (!find) { $('trk-bulk-preview').innerHTML = '<div style="padding:8px;color:var(--text-2);">Enter text to find.</div>'; return; }
  const replace = $('trk-bulk-replace-val').value;
  const caseSens = $('trk-bulk-case').checked;
  const exact = $('trk-bulk-exact').checked;
  const colFilter = $('trk-bulk-col').value;
  const visCols = trkVisibleCols();

  bulkPendingChanges = [];
  trkRows.forEach((row, ri) => {
    visCols.forEach(ci => {
      if (colFilter !== 'all' && ci !== parseInt(colFilter)) return;
      const val = row[ci] || '';
      let matches = false, newVal = val;
      if (exact) {
        matches = caseSens ? val === find : val.toLowerCase() === find.toLowerCase();
        if (matches) newVal = replace;
      } else {
        const flags = caseSens ? 'g' : 'gi';
        const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(escaped, flags);
        if (re.test(val)) { matches = true; newVal = val.replace(re, replace); }
      }
      if (matches && newVal !== val) bulkPendingChanges.push({ ri, ci, oldVal: val, newVal });
    });
  });

  const preview = $('trk-bulk-preview');
  if (bulkPendingChanges.length === 0) {
    preview.innerHTML = '<div style="padding:8px;color:var(--text-2);">No matches found.</div>';
    $('trk-bulk-apply-btn').disabled = true; $('trk-bulk-apply-btn').textContent = 'Apply (0 changes)';
  } else {
    let html = '';
    bulkPendingChanges.slice(0, 50).forEach(c => {
      html += '<div class="trk-bulk-preview-row"><span class="trk-bulk-label">R' + (c.ri+1) + '</span><span class="trk-bulk-old">' + esc(c.oldVal) + '</span><span>→</span><span class="trk-bulk-new">' + esc(c.newVal) + '</span></div>';
    });
    if (bulkPendingChanges.length > 50) html += '<div style="padding:4px 10px;color:var(--text-2);">...and ' + (bulkPendingChanges.length - 50) + ' more</div>';
    preview.innerHTML = html;
    $('trk-bulk-apply-btn').disabled = false;
    $('trk-bulk-apply-btn').textContent = 'Apply (' + bulkPendingChanges.length + ' changes)';
  }
});

$('trk-bulk-apply-btn').addEventListener('click', () => {
  if (bulkPendingChanges.length === 0) return;
  trkApplyEdits(bulkPendingChanges, 'Find & Replace (' + bulkPendingChanges.length + ')');
  bulkPendingChanges = [];
  $('trk-bulk-preview').innerHTML = '<div style="padding:8px;color:var(--green);">Done! Use Undo to revert.</div>';
  $('trk-bulk-apply-btn').disabled = true; $('trk-bulk-apply-btn').textContent = 'Apply (0 changes)';
});

// ── Set Column Value ──
let setPendingChanges = [];

$('trk-set-preview-btn').addEventListener('click', () => {
  const ci = parseInt($('trk-set-col').value);
  if (isNaN(ci)) return;
  const newVal = $('trk-set-val').value;
  const filter = $('trk-set-filter').value;
  const filterVal = $('trk-set-filter-val').value;

  setPendingChanges = [];
  trkRows.forEach((row, ri) => {
    const cur = row[ci] || '';
    let match = false;
    if (filter === 'any') match = true;
    else if (filter === 'empty') match = cur.trim() === '';
    else if (filter === 'notempty') match = cur.trim() !== '';
    else if (filter === 'equals') match = cur === filterVal;
    if (match && cur !== newVal) setPendingChanges.push({ ri, ci, oldVal: cur, newVal });
  });

  const preview = $('trk-set-preview');
  if (setPendingChanges.length === 0) {
    preview.innerHTML = '<div style="padding:8px;color:var(--text-2);">No rows match.</div>';
    $('trk-set-apply-btn').disabled = true; $('trk-set-apply-btn').textContent = 'Apply (0 changes)';
  } else {
    let html = '';
    setPendingChanges.slice(0, 50).forEach(c => {
      html += '<div class="trk-bulk-preview-row"><span class="trk-bulk-label">R' + (c.ri+1) + '</span><span class="trk-bulk-old">' + esc(c.oldVal || '(empty)') + '</span><span>→</span><span class="trk-bulk-new">' + esc(c.newVal || '(empty)') + '</span></div>';
    });
    if (setPendingChanges.length > 50) html += '<div style="padding:4px 10px;color:var(--text-2);">...and ' + (setPendingChanges.length - 50) + ' more</div>';
    preview.innerHTML = html;
    $('trk-set-apply-btn').disabled = false;
    $('trk-set-apply-btn').textContent = 'Apply (' + setPendingChanges.length + ' changes)';
  }
});

$('trk-set-apply-btn').addEventListener('click', () => {
  if (setPendingChanges.length === 0) return;
  trkApplyEdits(setPendingChanges, 'Set column (' + setPendingChanges.length + ')');
  setPendingChanges = [];
  $('trk-set-preview').innerHTML = '<div style="padding:8px;color:var(--green);">Done! Use Undo to revert.</div>';
  $('trk-set-apply-btn').disabled = true; $('trk-set-apply-btn').textContent = 'Apply (0 changes)';
});

// ── Clear / Remove ──
let clearPendingChanges = [];

$('trk-clear-preview-btn').addEventListener('click', () => {
  const ci = parseInt($('trk-clear-col').value);
  if (isNaN(ci)) return;
  const action = $('trk-clear-action').value;
  const removeText = $('trk-clear-text').value;

  clearPendingChanges = [];
  trkRows.forEach((row, ri) => {
    const cur = row[ci] || '';
    let newVal = cur;
    if (action === 'clear-col') newVal = '';
    else if (action === 'remove-text' && removeText) newVal = cur.split(removeText).join('');
    else if (action === 'trim') newVal = cur.trim();
    if (newVal !== cur) clearPendingChanges.push({ ri, ci, oldVal: cur, newVal });
  });

  const preview = $('trk-clear-preview');
  if (clearPendingChanges.length === 0) {
    preview.innerHTML = '<div style="padding:8px;color:var(--text-2);">No changes needed.</div>';
    $('trk-clear-apply-btn').disabled = true; $('trk-clear-apply-btn').textContent = 'Apply (0 changes)';
  } else {
    let html = '';
    clearPendingChanges.slice(0, 50).forEach(c => {
      html += '<div class="trk-bulk-preview-row"><span class="trk-bulk-label">R' + (c.ri+1) + '</span><span class="trk-bulk-old">' + esc(c.oldVal) + '</span><span>→</span><span class="trk-bulk-new">' + esc(c.newVal || '(empty)') + '</span></div>';
    });
    if (clearPendingChanges.length > 50) html += '<div style="padding:4px 10px;color:var(--text-2);">...and ' + (clearPendingChanges.length - 50) + ' more</div>';
    preview.innerHTML = html;
    $('trk-clear-apply-btn').disabled = false;
    $('trk-clear-apply-btn').textContent = 'Apply (' + clearPendingChanges.length + ' changes)';
  }
});

$('trk-clear-apply-btn').addEventListener('click', () => {
  if (clearPendingChanges.length === 0) return;
  trkApplyEdits(clearPendingChanges, 'Clear/Remove (' + clearPendingChanges.length + ')');
  clearPendingChanges = [];
  $('trk-clear-preview').innerHTML = '<div style="padding:8px;color:var(--green);">Done! Use Undo to revert.</div>';
  $('trk-clear-apply-btn').disabled = true; $('trk-clear-apply-btn').textContent = 'Apply (0 changes)';
});

// ══════════════════════════════════════════
// ── Filter done/not done ──
// ══════════════════════════════════════════
$('trk-filter-done').addEventListener('change', () => {
  trkApplyDoneFilter();
});

function trkApplyDoneFilter() {
  const filter = $('trk-filter-done').value;
  const search = $('trk-search').value;
  if (filter === 'all' && !search) {
    $('trk-tbody').querySelectorAll('tr').forEach(tr => { tr.style.display = ''; });
    return;
  }
  const sheetKey = trkActiveSheet;
  $('trk-tbody').querySelectorAll('tr').forEach(tr => {
    const tds = tr.querySelectorAll('td[data-uid]');
    if (tds.length === 0) { tr.style.display = ''; return; }
    const allDone = Array.from(tds).every(td => trkTouched[td.dataset.uid] === true);
    let show = true;
    if (filter === 'done') show = allDone;
    else if (filter === 'notdone') show = !allDone;
    // Also respect search
    if (show && search) {
      const q = search.trim().toLowerCase().replace(/\s+/g, '');
      let match = false;
      tds.forEach(td => { if (td.textContent.toLowerCase().replace(/\s+/g, '').includes(q)) match = true; });
      show = match;
    }
    tr.style.display = show ? '' : 'none';
  });
}

// Override search to also respect filter
$('trk-search').removeEventListener('input', () => {});
$('trk-search').addEventListener('input', () => trkApplyDoneFilter());

// ══════════════════════════════════════════
// ── Download CSV ──
// ══════════════════════════════════════════
$('trk-btn-download').addEventListener('click', () => {
  if (!trkActiveSheet) return;
  const sheetKey = trkActiveSheet;
  const sheet = trkSheets[sheetKey];
  const visCols = trkVisibleCols();

  // Build CSV with status column
  const headerRow = [...visCols.map(ci => trkHeaders[ci]), 'Status'].map(csvEsc).join(',');
  const dataLines = trkRows.map(row => {
    const allDone = visCols.every(ci => trkTouched[trkCellUid(sheetKey, row, ci)] === true);
    const cells = visCols.map(ci => csvEsc(row[ci] || ''));
    cells.push(allDone ? 'DONE' : 'NOT DONE');
    return cells.join(',');
  });

  const csv = [headerRow, ...dataLines].join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (sheet.name || 'tracker') + '-' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
});

function csvEsc(val) {
  const s = String(val);
  if (s.includes(',') || s.includes('"') || s.includes('\n')) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

// ══════════════════════════════════════════
// ── Import Done ──
// ══════════════════════════════════════════
let trkDoneImportedHeaders = [];
let trkDoneImportedRows = [];
let trkDonePendingMatches = [];

const doneOverlay = $('trk-done-overlay');

$('trk-import-done-file').addEventListener('change', e => {
  if (!e.target.files[0] || !trkActiveSheet) return;
  const file = e.target.files[0];
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();

  if (ext === 'csv') {
    reader.onload = ev => {
      const { headers, rows } = trkParseCSV(ev.target.result);
      trkDoneImportedHeaders = headers;
      trkDoneImportedRows = rows;
      trkShowDoneModal();
    };
    reader.readAsText(file);
  } else {
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: 'array', cellStyles: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const nonBlank = data.filter(r => r.some(c => String(c).trim() !== ''));
      if (nonBlank.length < 2) { alert('No data found.'); return; }
      trkDoneImportedHeaders = nonBlank[0].map(String);
      trkDoneImportedRows = nonBlank.slice(1).map(r => r.map(String));
      trkShowDoneModal();
    };
    reader.readAsArrayBuffer(file);
  }
  e.target.value = '';
});

function trkShowDoneModal() {
  const trkCol = $('trk-done-trk-col');
  const visCols = trkVisibleCols();
  trkCol.innerHTML = '';
  visCols.forEach(ci => {
    // Show unique count per column to help user pick the right one
    const unique = new Set();
    trkRows.forEach(row => { const v = (row[ci] || '').trim(); if (v) unique.add(v); });
    trkCol.innerHTML += '<option value="' + ci + '">' + esc(trkHeaders[ci] || 'Col ' + (ci+1)) + ' (' + unique.size + ' unique)</option>';
  });

  const impCol = $('trk-done-imp-col');
  impCol.innerHTML = '';
  trkDoneImportedHeaders.forEach((h, i) => {
    const unique = new Set();
    trkDoneImportedRows.forEach(r => { const v = (r[i] || '').trim(); if (v) unique.add(v); });
    impCol.innerHTML += '<option value="' + i + '">' + esc(h) + ' (' + unique.size + ' unique)</option>';
  });

  // Auto-select: prefer columns with similar unique counts and matching names
  let bestScore = -1;
  visCols.forEach(ci => {
    const trkH = (trkHeaders[ci] || '').toLowerCase().trim();
    const trkUnique = new Set();
    trkRows.forEach(row => { const v = (row[ci] || '').trim().toLowerCase(); if (v) trkUnique.add(v); });

    trkDoneImportedHeaders.forEach((impH, impIdx) => {
      const impHN = impH.toLowerCase().trim();
      const impUnique = new Set();
      trkDoneImportedRows.forEach(r => { const v = (r[impIdx] || '').trim().toLowerCase(); if (v) impUnique.add(v); });

      // Score: name similarity + overlap of actual values
      let score = 0;
      if (trkH === impHN) score += 10;
      else if (trkH.includes(impHN) || impHN.includes(trkH)) score += 5;
      // Count how many imported values exist in the tracker column
      let overlap = 0;
      impUnique.forEach(v => { if (trkUnique.has(v)) overlap++; });
      score += overlap;

      if (score > bestScore) {
        bestScore = score;
        trkCol.value = ci;
        impCol.value = impIdx;
      }
    });
  });

  $('trk-done-preview').innerHTML = '';
  $('trk-done-apply-btn').disabled = true;
  $('trk-done-apply-btn').textContent = 'Mark Done (0)';
  doneOverlay.classList.add('show');
}

$('trk-done-close').addEventListener('click', () => doneOverlay.classList.remove('show'));
doneOverlay.addEventListener('click', e => { if (e.target === doneOverlay) doneOverlay.classList.remove('show'); });

$('trk-done-preview-btn').addEventListener('click', () => {
  const trkCi = parseInt($('trk-done-trk-col').value);
  const impCi = parseInt($('trk-done-imp-col').value);
  if (isNaN(trkCi) || isNaN(impCi)) return;

  // Build set of imported values (normalized)
  const importedVals = new Set();
  trkDoneImportedRows.forEach(r => {
    const v = String(r[impCi] || '').trim().toLowerCase();
    if (v) importedVals.add(v);
  });

  const sheetKey = trkActiveSheet;
  const visCols = trkVisibleCols();
  trkDonePendingMatches = [];
  let alreadyDone = 0;
  trkRows.forEach((row, ri) => {
    const val = String(row[trkCi] || '').trim().toLowerCase();
    if (val && importedVals.has(val)) {
      const allDone = visCols.every(ci => trkTouched[trkCellUid(sheetKey, row, ci)] === true);
      if (allDone) { alreadyDone++; return; }
      // Show multiple column values for context
      const display = visCols.slice(0, 3).map(ci => row[ci] || '').filter(v => v).join(' | ');
      trkDonePendingMatches.push({ ri, row, matchVal: row[trkCi], displayVal: display });
    }
  });

  const preview = $('trk-done-preview');
  const matchedTotal = trkDonePendingMatches.length + alreadyDone;
  let html = '<div style="padding:6px 10px;font-size:11px;color:var(--text-2);border-bottom:1px solid var(--border-light);">' +
    'Imported: ' + importedVals.size + ' unique values · Matched: ' + matchedTotal + ' rows' +
    (alreadyDone > 0 ? ' (' + alreadyDone + ' already done)' : '') + '</div>';

  if (trkDonePendingMatches.length === 0) {
    html += '<div style="padding:8px;color:var(--text-2);">No new rows to mark.</div>';
    preview.innerHTML = html;
    $('trk-done-apply-btn').disabled = true;
    $('trk-done-apply-btn').textContent = 'Mark Done (0)';
  } else {
    trkDonePendingMatches.slice(0, 50).forEach(m => {
      html += '<div class="trk-bulk-preview-row"><span class="trk-bulk-label">R' + (m.ri + 1) + '</span><span><strong>' + esc(m.matchVal) + '</strong></span><span class="text-muted" style="font-size:10px;">' + esc(m.displayVal) + '</span></div>';
    });
    if (trkDonePendingMatches.length > 50) html += '<div style="padding:4px 10px;color:var(--text-2);">...and ' + (trkDonePendingMatches.length - 50) + ' more</div>';
    preview.innerHTML = html;
    $('trk-done-apply-btn').disabled = false;
    $('trk-done-apply-btn').textContent = 'Mark Done (' + trkDonePendingMatches.length + ' rows)';
  }
});

$('trk-done-apply-btn').addEventListener('click', () => {
  if (trkDonePendingMatches.length === 0) return;
  const sheetKey = trkActiveSheet;
  const visCols = trkVisibleCols();

  trkDonePendingMatches.forEach(m => {
    visCols.forEach(ci => {
      trkTouched[trkCellUid(sheetKey, m.row, ci)] = true;
    });
  });

  trkSave(); trkRenderTable(); trkUpdateStat();
  const count = trkDonePendingMatches.length;
  trkDonePendingMatches = [];
  $('trk-done-preview').innerHTML = '<div style="padding:8px;color:var(--green);">' + count + ' rows marked as done!</div>';
  $('trk-done-apply-btn').disabled = true;
  $('trk-done-apply-btn').textContent = 'Mark Done (0)';
});

// ── Persistence ──
window.addEventListener('beforeunload', () => {
  trkSave(); trkBackup();
  // Auto-save to org on close
  const orgs = new Set();
  Object.values(trkSheets).forEach(s => { if (s.orgName) orgs.add(s.orgName); });
  orgs.forEach(org => trkSaveSession(org));
});
setInterval(trkBackup, 30000);

// ══════════════════════════════════════════
// ── Public API ──
// ══════════════════════════════════════════
window.trkLoadSheetData = function(headers, rows, name) {
  if (!name) name = 'Sheet ' + (Object.keys(trkSheets).length + 1);
  const allRows = [headers, ...rows].map(r => r.map(c => String(c)));
  const parsed = trkSmartParse(allRows);
  trkShowSetup(name, parsed.headers, parsed.rows);
  $('trk-setup').style.display = ''; $('trk-main').style.display = 'none';
};

let trkRestored = false;
window.trkInit = function() {
  trkPopulateSavedSessions();
  if (!trkRestored && Object.keys(trkSheets).length === 0) {
    trkRestored = true;
    trkRestoreFromLocal();
  }
};

// Debug exposure — safe to leave on; lets you inspect state via DevTools console.
window.trkDebug = {
  get touched() { return trkTouched; },
  get flags() { return trkFlags; },
  get notes() { return trkNotes; },
  get sheets() { return trkSheets; },
  get activeSheet() { return trkActiveSheet; },
  auditRows() {
    const out = [];
    document.querySelectorAll('#trk-tbody tr').forEach((tr, i) => {
      const tds = tr.querySelectorAll('td[data-uid]');
      if (tds.length === 0) return;
      const doneCount = Array.from(tds).filter(td => trkTouched[td.dataset.uid] === true).length;
      const rowDone = tr.classList.contains('trk-row-done');
      const btn = tr.querySelector('.trk-row-btn');
      const btnDone = btn ? btn.classList.contains('done') : null;
      out.push({ row: i + 1, total: tds.length, doneCount, rowDone, btnDone, btnPresent: !!btn });
    });
    return out;
  }
};

})();
