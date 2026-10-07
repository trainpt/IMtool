// Data Progress Tracker - multi-sheet tracker with column reorder, resize, hide
(function() {

const TRK_STORAGE  = "dpt_touched_v1";
const TRK_FLAGS    = "dpt_flags_v1";
const TRK_NOTES    = "dpt_notes_v1";
const TRK_BACKUP   = "dpt_backup_v1";
const TRK_SHEETS   = "dpt_sheets_v2";
const TRK_SHEETS_FULL = "dpt_sheets_full_v1";
const TRK_ACTIVE   = "dpt_active_v1";

// Session-scoped, not localStorage — nothing this app holds outlives the page
// (see data-store.js). Row progress is kept for the session and reset on
// reload; trkStorageOk stays true so the "progress won't be saved" warnings
// don't fire on every load.
let trkStorageOk = true;
function sGet(k) { try { return JSON.parse(sessionPrefs.getItem(k)); } catch(e) { return null; } }
function sSet(k, v) { try { sessionPrefs.setItem(k, JSON.stringify(v)); } catch(e) {} }

// Skip-yellow preference. trkSmartParse and #trkSkipYellowToggle both reach for
// these; until now neither existed, so the `typeof === 'function'` guards fell
// through to a hardcoded `true` and the checkbox did nothing.
const TRK_SKIP_YELLOW = "dpt_skip_yellow_v1";
function getImportSkipYellow() {
  const v = sGet(TRK_SKIP_YELLOW);
  return v === null ? true : !!v;   // default ON, as before
}
function setImportSkipYellow(on) { sSet(TRK_SKIP_YELLOW, !!on); }

let trkTouched = sGet(TRK_STORAGE) || {};
let trkFlags   = sGet(TRK_FLAGS) || {};
let trkNotes   = sGet(TRK_NOTES) || {};

// Multi-sheet: { key: { name, headers, rows, colOrder, hiddenCols, colWidths, refRanges } }
let trkSheets = {};
let trkActiveSheet = null;

// Full grid of every sheet in the last-imported workbook, kept so a Reference
// Range can point at anything the tracked table left behind — a lookup block
// sitting off to the side, or a table on another tab. trkSmartParse keeps only
// the main table; without this the rest of the file is gone after import.
let trkWorkbookGrids = {};   // sheet name → [[cell strings]]
let trkWorkbookName = '';

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
      colWidths: s.colWidths || {}
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
      colWidths: s.colWidths || {}
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
      trkCaptureWorkbook(wb, file.name);
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
  // Header score = the CONTIGUOUS run of usable cells from the first filled
  // column, not the raw fill count.
  //
  // Raw counting breaks on any sheet that carries a second table off to the
  // side: on the implementation workbook's jobs tab a lookup table starts at
  // column M on row 3, so that data row fills 12 columns to the real header's
  // 11 and wins — every column title then comes out as a job record. A run
  // stops at the gap (8 vs 11), so the header wins again. Ties keep the
  // earliest row, which is what every well-formed sheet wants anyway.
  const usable = v => { const s = String(v || '').trim(); return s.length > 0 && s.length < 80; };
  const isNumericCell = v => { const s = String(v || '').trim(); return /\d/.test(s) && /^-?[\d.,]+%?$/.test(s); };
  function headerScore(row) {
    let first = -1;
    for (let i = 0; i < colCount; i++) { if (usable(row[i])) { first = i; break; } }
    if (first < 0) return 0;
    let run = 0;
    for (let i = first; i < colCount; i++) { if (usable(row[i])) run++; else break; }
    // A row that is mostly numbers is a record, not a heading — this stops an
    // ID-led data row taking a tie off a short header.
    const vals = [];
    for (let i = first; i < first + run; i++) vals.push(String(row[i]).trim());
    const numeric = vals.filter(isNumericCell).length;
    if (run >= 2 && numeric / run > 0.5) return 0;
    return run;
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
  // Yellow marks EXAMPLE rows in a blank template — a handful at the top. When
  // most of the sheet is yellow it is colour-coding, not examples, and dropping
  // it would throw the data away: the jobs tab of the implementation workbook
  // is 57 highlighted rows out of 73, which left 3. Above this share the
  // highlight is ignored regardless of the toggle.
  const YELLOW_MAJORITY = 0.3;
  const yellowIsFormatting = dataRows.length > 0 &&
    (yellowDataRows.size / dataRows.length) > YELLOW_MAJORITY;
  if (yellowIsFormatting) yellowDataRows.clear();

  // Fallback heuristic if no yellow info: check if first row after descriptions has high avg cell length
  let exampleIdx = -1;
  if (!yellowIsFormatting && yellowDataRows.size === 0 && lastDescIdx >= 0 && lastDescIdx + 1 < dataRows.length) {
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
    // Skip yellow-highlighted rows (examples in spreadsheet) — unless the user has
    // disabled this in the Import modal. When their real data is yellow-highlighted,
    // they need to opt out so the rows aren't dropped.
    const skipYellow = (typeof getImportSkipYellow === 'function') ? getImportSkipYellow() : true;
    if (skipYellow && yellowDataRows.has(ri)) return false;
    // Skip the example row right after descriptions (fallback heuristic) — also gated
    // by the skipYellow toggle so users can opt out of all example-row removal.
    if (skipYellow && ri === exampleIdx && ri < 5) return false;
    // Skip rows that don't fill enough columns compared to real data
    // (catches footer content like phone numbers, addresses, app names, notes)
    if (nonEmpty.length < minFillForData) return false;
    return true;
  });
  // Trim to the tracked table's own width.
  //
  // Stopping at "the last column holding any data" drags neighbouring tables in
  // with it: the jobs tab's earning-code block at M–P and its wage table at
  // R–S both ended up as unnamed columns of the job list. A separate block is
  // always divided from the main one by an EMPTY column, so walk right from the
  // header's run and stop at the first gap. A column that touches the run is
  // kept even with no header — those are real columns someone forgot to name
  // (Pack Size's notes, Equipment's "Remove").
  const colHasAnything = c =>
    String(headers[c] || '').trim() !== '' || filtered.some(r => String(r[c] || '').trim() !== '');
  let firstCol = 0;
  while (firstCol < colCount && !usable(headers[firstCol])) firstCol++;
  if (firstCol >= colCount) firstCol = 0;
  let lastUsedCol = firstCol;
  while (lastUsedCol + 1 < colCount && usable(headers[lastUsedCol + 1])) lastUsedCol++;
  while (lastUsedCol + 1 < colCount && colHasAnything(lastUsedCol + 1)) lastUsedCol++;
  if (lastUsedCol < 0) lastUsedCol = 0;
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
  trkRenderRefPanel();        // ranges are per-sheet
  trkRefreshSuggestions(true); // and so are the proposals
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
function trkOpenQuickAdd() {
  const overlay = $('trk-quickadd-overlay');
  if (!overlay) return;
  trkRenderQuickAddList();
  $('trk-quickadd-file').value = '';
  overlay.classList.add('show');
}

function trkCloseQuickAdd() {
  const overlay = $('trk-quickadd-overlay');
  if (overlay) overlay.classList.remove('show');
}

// The shared sheet library is gone — every tool takes its own upload now — so
// there is nothing to quick-add from. Point the user at the upload instead.
function trkRenderQuickAddList() {
  const list = $('trk-quickadd-list');
  if (!list) return;
  list.innerHTML = '<div class="trk-saved-empty">Upload a file below to get started.</div>';
}

// Switch the tracker back to the setup screen.
function trkGotoSetupWithOrg() {
  $('trk-main').style.display = 'none';
  $('trk-setup').style.display = '';
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
  // Multi-sheet path
  if (trkStagedMulti && trkStagedMulti.length > 0) {
    // If only 1 sheet staged, use the renamed name from the input field
    const userRename = ($('trk-staging-name').value || '').trim();
    trkStagedMulti.forEach(sh => {
      const sheetName = (trkStagedMulti.length === 1 && userRename) ? userRename : sh.name;
      const key = makeSheetKey(sheetName);
      trkPurgeSheetProgress(key);
      const savedMeta = sGet(TRK_SHEETS) || {};
      const saved = savedMeta[key];
      const colOrder = saved?.colOrder?.length === sh.headers.length ? saved.colOrder : sh.headers.map((_, i) => i);
      trkSheets[key] = { name: sheetName, headers: sh.headers, rows: trkTagRows(sh.rows), colOrder, hiddenCols: sh.hiddenCols || new Set(), colWidths: saved?.colWidths || {} };
    });
    const firstName = (trkStagedMulti.length === 1 && userRename) ? userRename : trkStagedMulti[0].name;
    const firstKey = makeSheetKey(firstName);
    trkSaveSheetsMeta();
    trkSwitchToSheet(trkSheets[firstKey] ? firstKey : Object.keys(trkSheets)[0]);
    $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
    trkStagedMulti = null;
    return;
  }

  // Single-sheet path
  if (!trkStagingHeaders.length || !trkStagingRows.length) { alert('Load a file first.'); return; }
  const name = ($('trk-staging-name').value || '').trim() || trkStagingName;
  const key = makeSheetKey(name);
  const hidden = new Set();
  const hideBlankCb = document.getElementById('trk-hide-blank-cb');
  if (hideBlankCb && hideBlankCb.checked) {
    findBlankCols(trkStagingHeaders, trkStagingRows).forEach(ci => hidden.add(ci));
  }
  trkPurgeSheetProgress(key);

  trkSheets[key] = { name, headers: trkStagingHeaders, rows: trkTagRows(trkStagingRows.map(r => r.map(c => String(c)))), colOrder: [...trkStagingColOrder], hiddenCols: hidden, colWidths: {} };
  trkSaveSheetsMeta(); trkSwitchToSheet(key);
  $('trk-setup').style.display = 'none'; $('trk-main').style.display = 'flex';
  trkStagingHeaders = []; trkStagingRows = []; trkStagingColOrder = [];
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
  const pasteEl = $('trk-ctx-paste');
  if (pasteEl) pasteEl.textContent = '📋 Paste' + (trkSelected.size > 1 ? ' (' + trkSelected.size + ' cells)' : '');
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

// ── Bulk / single paste into selected cells ──
// Apply text to all currently selected cells (or to a fallback uid if none selected)
function trkPasteIntoTargets(text, fallbackUid) {
  if (text == null) return;
  text = String(text).replace(/\r?\n$/, '');
  let targets = [];
  if (trkSelected.size > 0) {
    trkSelected.forEach(key => {
      const [ri, ci] = key.split('-').map(Number);
      targets.push({ ri, ci });
    });
  } else if (fallbackUid) {
    const td = document.querySelector('#trk-tbody td[data-uid="' + fallbackUid + '"]');
    if (td) targets.push({ ri: parseInt(td.dataset.ri), ci: parseInt(td.dataset.ci) });
  }
  if (targets.length === 0) return;
  const changes = [];
  targets.forEach(({ ri, ci }) => {
    const oldVal = (trkRows[ri] && trkRows[ri][ci]) || '';
    if (oldVal !== text) changes.push({ ri, ci, oldVal, newVal: text });
  });
  if (changes.length > 0) {
    trkApplyEdits(changes, 'Paste into ' + changes.length + ' cell' + (changes.length > 1 ? 's' : ''));
  }
  trkClearSelection();
}

// Hidden textarea that captures the browser's native paste event.
// This avoids both the Clipboard API permission/secure-context limits and
// the fact that <td> elements aren't focusable (so paste never fires on them).
const trkPasteCapture = document.createElement('textarea');
trkPasteCapture.setAttribute('aria-hidden', 'true');
trkPasteCapture.tabIndex = -1;
trkPasteCapture.style.cssText = 'position:fixed;left:-9999px;top:-9999px;width:10px;height:10px;opacity:0;';
document.body.appendChild(trkPasteCapture);

let trkPasteFallbackUid = null;
trkPasteCapture.addEventListener('paste', e => {
  e.preventDefault();
  const text = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
  const fb = trkPasteFallbackUid; trkPasteFallbackUid = null;
  if (text) trkPasteIntoTargets(text, fb);
  trkPasteCapture.value = '';
  setTimeout(() => { try { trkPasteCapture.blur(); } catch(_) {} }, 0);
});

// Ctrl+V / Cmd+V: redirect focus to the hidden textarea so the browser
// delivers the paste event there, then we read it synchronously.
document.addEventListener('keydown', e => {
  if (!((e.ctrlKey || e.metaKey) && (e.key === 'v' || e.key === 'V') && !e.shiftKey && !e.altKey)) return;
  const activePage = document.querySelector('.page.active');
  if (!activePage || activePage.id !== 'page-tracker') return;
  const active = document.activeElement;
  if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT' || active.isContentEditable)) return;
  if (trkSelected.size === 0) return;
  // Don't preventDefault — we *want* the native paste to fire on the textarea
  trkPasteFallbackUid = null;
  trkPasteCapture.value = '';
  trkPasteCapture.focus();
});

// Right-click "Paste": try Clipboard API, then fall back to prompt
const trkCtxPaste = $('trk-ctx-paste');
if (trkCtxPaste) {
  trkCtxPaste.addEventListener('click', async () => {
    const fallbackUid = trkCtxTarget ? trkCtxTarget.dataset.uid : null;
    trkCtxMenu.style.display = 'none';
    let text = null;
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        text = await navigator.clipboard.readText();
      }
    } catch (_) { text = null; }
    if (text == null || text === '') {
      const v = prompt('Paste value to apply to selected cell(s):', '');
      if (v === null) return;
      text = v;
    }
    trkPasteIntoTargets(text, fallbackUid);
  });
}

// ── Save button ──
// Progress is written to the session store on every click already; this just
// confirms it, since there's no longer anywhere else to save to.
$('trk-btn-save').addEventListener('click', () => {
  trkSave();
  const el = $('trk-stat');
  const prev = el.textContent;
  el.textContent = '💾 Progress saved for this session';
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
  if (last.kind === 'rows') {
    // Put them back where they were, lowest index first.
    last.removed.slice().sort((a, b) => a.index - b.index)
      .forEach(r => trkRows.splice(Math.min(r.index, trkRows.length), 0, r.row));
    if (trkSheets[trkActiveSheet]) trkSheets[trkActiveSheet].rows = trkRows;
    trkBuildReview(); trkUpdateUndoBtn();
    if (typeof trkRefreshSuggestions === 'function') trkRefreshSuggestions(true);
    return;
  }
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

{
  const dd = $('trk-dedupe-file');
  if (dd) dd.addEventListener('change', e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    if (!trkActiveSheet) { alert('Load a sheet first.'); return; }
    trkReadAnySheet(file, imp => {
      if (!imp || !imp.rows.length) { alert('No data found in that file.'); return; }
      trkDedupeAgainst(imp);
    });
  });
}

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

// ══════════════════════════════════════════
// ── Remove duplicates against another file ──
//
// "Here is what is already in the system — take those off my list." Reads any
// sheet, works out by itself which column pairs up with which, and offers to
// drop the rows that already exist. No column pickers unless the guess is
// wrong. Removal is a single undo away.
// ══════════════════════════════════════════
function trkReadAnySheet(file, cb) {
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();
  if (ext === 'csv') {
    reader.onload = ev => {
      const { headers, rows } = trkParseCSV(ev.target.result);
      cb({ headers: headers, rows: rows, name: file.name });
    };
    reader.readAsText(file);
    return;
  }
  reader.onload = ev => {
    const wb = XLSX.read(ev.target.result, { type: 'array' });
    // Take whichever sheet carries the most rows — the reference export is
    // almost never the tiny legend tab.
    let bestName = wb.SheetNames[0], bestRows = null, bestCount = -1;
    wb.SheetNames.forEach(n => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' });
      const nb = data.filter(r => r.some(c => String(c).trim() !== ''));
      if (nb.length > bestCount) { bestCount = nb.length; bestName = n; bestRows = nb; }
    });
    if (!bestRows || bestRows.length < 2) { alert('No data found in that file.'); return; }
    const sliced = trkSmartParse(bestRows.map(r => r.map(String)), null);
    cb({ headers: sliced.headers, rows: sliced.rows, name: file.name + ' [' + bestName + ']' });
  };
  reader.readAsArrayBuffer(file);
}

// Best (tracked column ↔ imported column) pairing, scored on how many of the
// tracked sheet's values actually appear in the other file.
function trkBestDedupeMatch(imp) {
  const norm = v => String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim();
  let best = null;
  const visible = trkVisibleCols();
  visible.forEach(ci => {
    const trkVals = new Set();
    trkRows.forEach(r => { const v = norm(r[ci]); if (v) trkVals.add(v); });
    if (trkVals.size < 2) return;
    imp.headers.forEach((h, ii) => {
      const impVals = new Set();
      imp.rows.forEach(r => { const v = norm(r[ii]); if (v) impVals.add(v); });
      if (!impVals.size) return;
      let overlap = 0;
      trkVals.forEach(v => { if (impVals.has(v)) overlap++; });
      if (!overlap) return;
      // Two columns can hold the same values (an export where every site is
      // also its own group), so a header that looks related breaks the tie
      // toward the one a person would have picked.
      const a = norm(trkHeaders[ci]), b = norm(h);
      const nameBonus = (a && b && a === b) ? 0.15
        : (a && b && (a.indexOf(b) >= 0 || b.indexOf(a) >= 0)) ? 0.08 : 0;
      const share = overlap / trkVals.size;
      const score = share + nameBonus;
      if (!best || score > best.score) {
        best = { ci: ci, ii: ii, overlap: overlap, share: share, score: score, impVals: impVals };
      }
    });
  });
  return best;
}

function trkDedupeAgainst(imp) {
  const m = trkBestDedupeMatch(imp);
  if (!m) {
    alert('Nothing in "' + imp.name + '" matches this sheet, so there is nothing to remove.\n\n' +
          'Check it is the right file — the values have to line up with one of your columns.');
    return;
  }
  const norm = v => String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim();
  const hits = [];
  trkRows.forEach((r, ri) => { const v = norm(r[m.ci]); if (v && m.impVals.has(v)) hits.push(ri); });
  if (!hits.length) { alert('No rows in this sheet appear in "' + imp.name + '".'); return; }
  trkShowDedupeModal(imp, m, hits);
}

function trkShowDedupeModal(imp, m, hits) {
  let ov = $('trk-dedupe-overlay');
  if (!ov) { ov = document.createElement('div'); ov.id = 'trk-dedupe-overlay'; ov.className = 'modal-overlay'; document.body.appendChild(ov); }
  const sample = hits.slice(0, 6).map(ri =>
    '<li>' + esc(String(trkRows[ri][m.ci] || '').slice(0, 60)) + '</li>').join('');
  const colOpts = trkVisibleCols().map(ci =>
    '<option value="' + ci + '"' + (ci === m.ci ? ' selected' : '') + '>' +
    esc(trkHeaders[ci] || 'Col ' + (ci + 1)) + '</option>').join('');
  const impOpts = imp.headers.map((h, i) =>
    '<option value="' + i + '"' + (i === m.ii ? ' selected' : '') + '>' + esc(h || 'Col ' + (i + 1)) + '</option>').join('');
  ov.innerHTML =
    '<div class="modal" style="max-width:520px;"><h3>Remove rows already in that file</h3>' +
    '<div class="trk-dd-hit"><b>' + hits.length + '</b> of ' + trkRows.length +
      ' row' + (trkRows.length === 1 ? '' : 's') + ' already exist in <b>' + esc(imp.name) + '</b>' +
      '<div class="text-muted small" style="margin-top:3px;">matched <b>' + esc(trkHeaders[m.ci] || 'column') +
      '</b> against <b>' + esc(imp.headers[m.ii] || 'column') + '</b> &middot; ' +
      Math.round(m.share * 100) + '% of this column\'s values</div></div>' +
    '<ul class="ts-confirm-list">' + sample + (hits.length > 6 ? '<li>… and ' + (hits.length - 6) + ' more</li>' : '') + '</ul>' +
    '<details class="trk-dd-adv"><summary>Matched the wrong columns?</summary>' +
      '<div class="trk-inline" style="margin-top:8px;">' +
        '<select class="input-field input-sm" id="trk-dd-trk">' + colOpts + '</select>' +
        '<span class="text-muted small">against</span>' +
        '<select class="input-field input-sm" id="trk-dd-imp">' + impOpts + '</select>' +
      '</div></details>' +
    '<div class="modal-actions"><button class="btn btn-ghost" id="trk-dd-cancel">Cancel</button>' +
    '<button class="btn btn-danger" id="trk-dd-go">Remove ' + hits.length + ' rows</button></div></div>';
  ov.classList.add('show');
  ov.style.display = 'flex';
  const close = () => { ov.classList.remove('show'); ov.style.display = 'none'; };
  ov.querySelector('#trk-dd-cancel').addEventListener('click', close);
  ov.addEventListener('click', e => { if (e.target === ov) close(); });

  const recompute = () => {
    const ci = +ov.querySelector('#trk-dd-trk').value;
    const ii = +ov.querySelector('#trk-dd-imp').value;
    const norm = v => String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim();
    const vals = new Set();
    imp.rows.forEach(r => { const v = norm(r[ii]); if (v) vals.add(v); });
    const next = [];
    trkRows.forEach((r, ri) => { const v = norm(r[ci]); if (v && vals.has(v)) next.push(ri); });
    hits = next;
    ov.querySelector('#trk-dd-go').textContent = 'Remove ' + hits.length + ' rows';
    ov.querySelector('#trk-dd-go').disabled = !hits.length;
  };
  ov.querySelector('#trk-dd-trk').addEventListener('change', recompute);
  ov.querySelector('#trk-dd-imp').addEventListener('change', recompute);
  ov.querySelector('#trk-dd-go').addEventListener('click', () => {
    trkRemoveRows(hits, 'Removed ' + hits.length + ' already in ' + imp.name);
    close();
  });
}

// Splice rows out, remembering enough to put them back. Progress is keyed by
// each row's stable _rid, so a restored row brings its ticks with it.
function trkRemoveRows(indexes, label) {
  if (!indexes || !indexes.length) return;
  const sorted = indexes.slice().sort((a, b) => a - b);
  const removed = sorted.map(i => ({ index: i, row: trkRows[i] }));
  for (let k = sorted.length - 1; k >= 0; k--) trkRows.splice(sorted[k], 1);
  const sheet = trkSheets[trkActiveSheet];
  if (sheet) sheet.rows = trkRows;
  trkUndoStack.push({ kind: 'rows', removed: removed, label: label || 'Remove rows' });
  if (trkUndoStack.length > TRK_MAX_UNDO) trkUndoStack.shift();
  trkUpdateUndoBtn();
  trkBuildReview();
  trkRefreshSuggestions(true);
}

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
window.addEventListener('beforeunload', () => { trkSave(); trkBackup(); });
setInterval(trkBackup, 30000);

// ══════════════════════════════════════════
// ── Reference Ranges ──
//
// A workbook rarely holds one tidy table. The implementation workbook's jobs
// tab carries three blocks side by side: the job list in A–K, an earning-code
// lookup in M–P, and a state minimum-wage table in R3:S5. trkSmartParse keeps
// the first and drops the rest, which is right for tracking but throws away
// exactly the lookups you need while filling cells in.
//
// A Reference Range pins any rectangle of the workbook — on this sheet or any
// other — beside the grid, and can be used as a key→value lookup to fill a
// tracked column. Fills go through trkApplyEdits, so Undo and Save behave
// the same as a hand edit.
// ══════════════════════════════════════════

function trkCaptureWorkbook(wb, fileName) {
  trkWorkbookGrids = {};
  trkWorkbookName = fileName || '';
  (wb.SheetNames || []).forEach(name => {
    try {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
      trkWorkbookGrids[name] = aoa.map(r => (r || []).map(c => String(c == null ? '' : c)));
    } catch (err) { /* a sheet we can't read simply isn't referenceable */ }
  });
}

function trkColLetter(n) {
  let s = '', x = n + 1;
  while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); }
  return s;
}
function trkColNum(s) {
  let n = 0;
  for (const ch of String(s).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function trkRangeLabel(box) {
  return trkColLetter(box.c0) + (box.r0 + 1) + ':' + trkColLetter(box.c1) + (box.r1 + 1);
}

// Accepts "Sheet!R3:S5", "R3:S5", "R3", "R:S", "S", "3:5", "3".
// Open-ended forms are clamped to the sheet's real extent.
function trkParseA1(text, defaultSheet) {
  let s = String(text || '').trim();
  if (!s) return null;
  let sheet = defaultSheet;
  const bang = s.lastIndexOf('!');
  if (bang >= 0) {
    sheet = s.slice(0, bang).replace(/^'|'$/g, '').trim();
    s = s.slice(bang + 1).trim();
  }
  const grid = trkWorkbookGrids[sheet];
  if (!grid) return null;
  const maxRow = Math.max(0, grid.length - 1);
  const maxCol = Math.max(0, grid.reduce((m, r) => Math.max(m, (r || []).length), 0) - 1);
  const cell = /^([A-Za-z]+)(\d+)$/;
  const parts = s.split(':').map(p => p.trim());
  const one = p => {
    const m = cell.exec(p);
    if (m) return { c: trkColNum(m[1]), r: +m[2] - 1 };
    if (/^[A-Za-z]+$/.test(p)) return { c: trkColNum(p), r: null };
    if (/^\d+$/.test(p)) return { c: null, r: +p - 1 };
    return null;
  };
  const a = one(parts[0]);
  if (!a) return null;
  const b = parts.length > 1 ? one(parts[1]) : a;
  if (!b) return null;
  const box = {
    sheet: sheet,
    c0: Math.min(a.c == null ? 0 : a.c, b.c == null ? maxCol : b.c),
    c1: Math.max(a.c == null ? maxCol : a.c, b.c == null ? maxCol : b.c),
    r0: Math.min(a.r == null ? 0 : a.r, b.r == null ? maxRow : b.r),
    r1: Math.max(a.r == null ? maxRow : a.r, b.r == null ? maxRow : b.r)
  };
  box.c0 = Math.max(0, Math.min(box.c0, maxCol));
  box.c1 = Math.max(0, Math.min(box.c1, maxCol));
  box.r0 = Math.max(0, Math.min(box.r0, maxRow));
  box.r1 = Math.max(0, Math.min(box.r1, maxRow));
  return box;
}

// Rectangles of data separated by at least one fully empty column. The first
// group is the tracked table itself; the rest are what this feature is for.
function trkFindBlocks(sheetName) {
  const grid = trkWorkbookGrids[sheetName];
  if (!grid || !grid.length) return [];
  const width = grid.reduce((m, r) => Math.max(m, (r || []).length), 0);
  const colHas = [];
  for (let c = 0; c < width; c++) colHas[c] = grid.some(r => r && String(r[c] || '').trim() !== '');
  const groups = [];
  let c = 0;
  while (c < width) {
    if (!colHas[c]) { c++; continue; }
    const start = c;
    while (c < width && colHas[c]) c++;
    groups.push({ c0: start, c1: c - 1 });
  }
  return groups.map(g => {
    let r0 = -1, r1 = -1;
    grid.forEach((row, ri) => {
      if (!row) return;
      for (let cc = g.c0; cc <= g.c1; cc++) {
        if (String(row[cc] || '').trim() !== '') { if (r0 < 0) r0 = ri; r1 = ri; break; }
      }
    });
    return { sheet: sheetName, c0: g.c0, c1: g.c1, r0: r0, r1: r1 };
  }).filter(b => b.r0 >= 0);
}

function trkReadRange(box) {
  const grid = trkWorkbookGrids[box.sheet];
  if (!grid) return [];
  const out = [];
  for (let r = box.r0; r <= box.r1; r++) {
    const row = grid[r] || [];
    const line = [];
    for (let c = box.c0; c <= box.c1; c++) line.push(String(row[c] == null ? '' : row[c]).trim());
    if (line.some(v => v !== '')) out.push(line);
  }
  return out;
}

function trkRefList() {
  const sh = trkSheets[trkActiveSheet];
  if (!sh) return [];
  if (!sh.refRanges) sh.refRanges = [];
  return sh.refRanges;
}
function trkRuleList() {
  const sh = trkSheets[trkActiveSheet];
  if (!sh) return [];
  if (!sh.fillRules) sh.fillRules = [];
  return sh.fillRules;
}

// ─── Key extraction ───
// Lookup keys rarely sit in a cell by themselves. "Break (AZ)" has to become
// "AZ" before it can meet "Arizona (AZ)" — and the same trick, applied to both
// sides, is what makes one lookup serve a dozen different naming habits.
const TRK_EXTRACTORS = [
  { id: 'whole',      label: 'whole cell' },
  { id: 'parens',     label: 'text in (brackets)' },
  { id: 'brackets',   label: 'text in [brackets]' },
  { id: 'first',      label: 'first word' },
  { id: 'last',       label: 'last word' },
  { id: 'beforeDash', label: 'text before a dash' },
  { id: 'afterDash',  label: 'text after a dash' },
  { id: 'custom',     label: 'custom pattern…' }
];
function trkExtract(v, mode, custom) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  let m;
  switch (mode) {
    case 'parens':     m = s.match(/\(([^)]*)\)/);   return m ? m[1].trim() : '';
    case 'brackets':   m = s.match(/\[([^\]]*)\]/);  return m ? m[1].trim() : '';
    case 'first':      return (s.split(/\s+/)[0] || '');
    case 'last':       { const p = s.split(/\s+/); return p[p.length - 1] || ''; }
    case 'beforeDash': m = s.split(/\s*[-–—]\s*/); return m.length > 1 ? m[0].trim() : s;
    case 'afterDash':  m = s.match(/[-–—]\s*(.+)$/); return m ? m[1].trim() : '';
    case 'custom':
      if (!custom) return s;
      try {
        const re = new RegExp(custom, 'i');
        const hit = s.match(re);
        return hit ? String(hit[1] != null ? hit[1] : hit[0]).trim() : '';
      } catch (e) { return ''; }
    default: return s;
  }
}

// Works out every cell the rule would write, and buckets whatever it cannot
// resolve so those can be typed in. Precedence: a per-value entry you typed
// beats a per-key entry, which beats the lookup range.
function trkPlanFill(rule) {
  const data = trkReadRange(rule.box);
  const norm = v => rule.loose
    ? String(v || '').toLowerCase().replace(/[^a-z0-9]/g, '')
    : String(v || '').trim();
  const map = new Map();
  data.forEach(row => {
    const k = norm(trkExtract(row[rule.keyCol], rule.keyExtract, rule.keyCustom));
    if (k && !map.has(k)) map.set(k, String(row[rule.valCol] == null ? '' : row[rule.valCol]).trim());
  });
  const isNew = rule.destCi === '__new__';
  const destCi = isNew ? trkHeaders.length : Number(rule.destCi);
  const manual = rule.manual || {};
  const perRow = rule.perRow || {};
  const hits = [], groups = new Map();
  let skippedFilled = 0;
  trkRows.forEach((row, ri) => {
    const raw = String(row[rule.srcCi] == null ? '' : row[rule.srcCi]).trim();
    if (!raw) return;
    const ex = trkExtract(raw, rule.srcExtract, rule.srcCustom);
    const k = norm(ex);
    const cur = isNew ? '' : String(row[destCi] == null ? '' : row[destCi]).trim();
    if (rule.blanksOnly && cur !== '') { skippedFilled++; return; }
    let val = null;
    if (perRow[raw] != null && perRow[raw] !== '') val = perRow[raw];
    else if (k && manual[k] != null && manual[k] !== '') val = manual[k];
    else if (k && map.has(k)) val = map.get(k);
    if (val != null) {
      if (cur !== val) hits.push({ ri: ri, ci: destCi, oldVal: isNew ? '' : (row[destCi] == null ? '' : row[destCi]), newVal: val });
      return;
    }
    // Unresolved. Group by the extracted key when there is one, otherwise by
    // the cell itself — so every "(CA)" job shares one box while "Stand By"
    // and "Lunch" each get their own.
    const gid = k ? 'k\u0000' + k : 'r\u0000' + raw;
    const g = groups.get(gid) || { kind: k ? 'key' : 'row', key: k, raw: raw, label: k ? ex : raw, rows: [], samples: [] };
    g.rows.push(ri);
    if (g.samples.length < 4 && g.samples.indexOf(raw) < 0) g.samples.push(raw);
    groups.set(gid, g);
  });
  return {
    hits: hits, keys: map.size, isNew: isNew, destCi: destCi, skippedFilled: skippedFilled,
    groups: [...groups.values()].sort((a, b) => b.rows.length - a.rows.length || a.label.localeCompare(b.label))
  };
}

// ─── Panel ───
function trkRenderRefPanel() {
  const wrap = $('trk-ref-panel');
  if (!wrap) return;
  const list = trkRefList();
  const rules = trkRuleList();
  const btn = $('trk-btn-ref');
  if (btn) btn.style.display = Object.keys(trkWorkbookGrids).length ? '' : 'none';
  if (!list.length && !rules.length) { wrap.style.display = 'none'; wrap.innerHTML = ''; return; }
  wrap.style.display = '';
  let html = '';
  if (rules.length) {
    html += '<div class="trk-rule-bar"><span class="text-muted small">Fill rules</span>' +
      rules.map(r => {
        const dest = r.destCi === '__new__' ? (r.destName || 'new column') : (trkHeaders[r.destCi] || 'column ' + (Number(r.destCi) + 1));
        return '<span class="trk-rule-chip" title="' + esc((trkHeaders[r.srcCi] || '?') + ' → ' + dest) + '">' +
          '<b>' + esc(r.name) + '</b>' +
          '<span class="text-muted small">&rarr; ' + esc(dest) + '</span>' +
          (r.lastRun != null ? '<span class="text-muted small">(' + r.lastRun + ')</span>' : '') +
          '<button class="btn btn-primary btn-sm trk-rule-run" data-id="' + r.id + '">Run</button>' +
          '<button class="btn btn-ghost btn-sm trk-rule-edit" data-id="' + r.id + '" title="Edit">&#9998;</button>' +
          '<button class="btn btn-ghost btn-sm trk-rule-del" data-id="' + r.id + '" title="Delete rule">&times;</button>' +
        '</span>';
      }).join('') + '</div>';
  }
  list.forEach(rr => {
    const data = trkReadRange(rr.box);
    const cols = data.length ? data[0].length : 0;
    html += '<div class="trk-ref-card">' +
      '<div class="trk-ref-head">' +
        '<b>' + esc(rr.name) + '</b>' +
        '<code class="trk-ref-ref">' + esc(rr.box.sheet + '!' + trkRangeLabel(rr.box)) + '</code>' +
        '<span class="text-muted small">' + cols + ' col' + (cols === 1 ? '' : 's') +
          ' &times; ' + data.length + ' row' + (data.length === 1 ? '' : 's') + '</span>' +
        '<span style="margin-left:auto;"></span>' +
        (cols >= 2 ? '<button class="btn btn-primary btn-sm trk-ref-fill" data-id="' + rr.id + '">Fill&hellip;</button>' : '') +
        '<button class="btn btn-ghost btn-sm trk-ref-del" data-id="' + rr.id + '" title="Remove this range">&times;</button>' +
      '</div>' +
      '<div class="trk-ref-body"><table class="data-table trk-ref-table"><tbody>';
    data.slice(0, 40).forEach(row => {
      html += '<tr>' + row.map(v =>
        '<td class="trk-ref-cell" title="Click to copy">' + esc(v) + '</td>').join('') + '</tr>';
    });
    html += '</tbody></table>' +
      (data.length > 40 ? '<div class="text-muted small" style="padding:4px 6px;">… and ' + (data.length - 40) + ' more rows</div>' : '') +
      '</div></div>';
  });
  wrap.innerHTML = html;
  wrap.querySelectorAll('.trk-ref-cell').forEach(td => td.addEventListener('click', () => {
    const t = (td.textContent || '').trim();
    if (t && navigator.clipboard) navigator.clipboard.writeText(t).then(() => {
      td.classList.add('trk-ref-copied');
      setTimeout(() => td.classList.remove('trk-ref-copied'), 700);
    }, () => {});
  }));
  wrap.querySelectorAll('.trk-ref-del').forEach(b => b.addEventListener('click', () => {
    const l = trkRefList();
    const i = l.findIndex(x => x.id === b.dataset.id);
    if (i >= 0) l.splice(i, 1);
    trkRenderRefPanel();
  }));
  wrap.querySelectorAll('.trk-ref-fill').forEach(b => b.addEventListener('click', () => {
    const draft = trkNewRule(b.dataset.id);
    if (draft) trkShowRuleEditor(draft, true);
  }));
  wrap.querySelectorAll('.trk-rule-run').forEach(b => b.addEventListener('click', () => {
    const r = trkRuleList().find(x => x.id === b.dataset.id);
    if (!r) return;
    if (!trkRunRule(r)) alert('Nothing to fill — every target cell already holds the right value.');
    trkRenderRefPanel();
  }));
  wrap.querySelectorAll('.trk-rule-edit').forEach(b => b.addEventListener('click', () => {
    const r = trkRuleList().find(x => x.id === b.dataset.id);
    if (r) trkShowRuleEditor(r, false);
  }));
  wrap.querySelectorAll('.trk-rule-del').forEach(b => b.addEventListener('click', () => {
    const l = trkRuleList();
    const i = l.findIndex(x => x.id === b.dataset.id);
    if (i >= 0) l.splice(i, 1);
    trkRenderRefPanel();
  }));
}

// ─── Add-range modal ───
function trkShowAddRangeModal() {
  const grids = Object.keys(trkWorkbookGrids);
  if (!grids.length) { alert('Reference ranges come from an uploaded Excel workbook. Load one on the Setup screen first.'); return; }
  const sh = trkSheets[trkActiveSheet];
  const own = (sh && sh.name && trkWorkbookGrids[sh.name]) ? sh.name : grids[0];

  let ov = $('trk-range-overlay');
  const blocks = [];
  grids.forEach(sn => {
    trkFindBlocks(sn).forEach((b, i) => {
      // The first block of the tracked sheet IS the tracked table — skip it.
      if (sn === own && i === 0) return;
      blocks.push(b);
    });
  });
  // Nearest first: this sheet's own side-blocks before other tabs'.
  blocks.sort((a, b) => (a.sheet === own ? 0 : 1) - (b.sheet === own ? 0 : 1));

  let html = '<div class="modal" style="max-width:640px;"><h3>Add reference range</h3>' +
    '<p class="text-muted small">Blocks found outside the tracked table. Pick one, or type any range &mdash; ' +
    '<code>R3:S5</code>, <code>Sheet!R3:S5</code>, a column like <code>S</code>, or rows like <code>3:5</code>.</p>' +
    '<div class="trk-range-blocks">';
  if (!blocks.length) html += '<div class="text-muted small" style="padding:8px;">No separate blocks found — type a range below.</div>';
  blocks.slice(0, 12).forEach((b, i) => {
    const preview = trkReadRange(b).slice(0, 2)
      .map(r => r.slice(0, 4).map(v => esc(v.length > 26 ? v.slice(0, 26) + '…' : v)).join(' <span class="text-muted">|</span> '))
      .join('<br>');
    html += '<label class="trk-range-block">' +
      '<input type="radio" name="trk-range-pick" value="' + i + '">' +
      '<div><code>' + esc(b.sheet + '!' + trkRangeLabel(b)) + '</code> ' +
      '<span class="text-muted small">' + (b.c1 - b.c0 + 1) + ' cols &times; ' + (b.r1 - b.r0 + 1) + ' rows</span>' +
      '<div class="trk-range-prev">' + preview + '</div></div></label>';
  });
  html += '</div>' +
    '<div style="display:flex;gap:8px;align-items:center;margin:10px 0;flex-wrap:wrap;">' +
      '<label class="trk-range-block" style="flex:1;min-width:260px;">' +
        '<input type="radio" name="trk-range-pick" value="custom">' +
        '<div style="flex:1;"><span class="text-muted small">Type a range</span>' +
        '<div style="display:flex;gap:6px;margin-top:4px;">' +
        '<select class="input-field input-sm" id="trk-range-sheet">' +
          grids.map(s => '<option' + (s === own ? ' selected' : '') + '>' + esc(s) + '</option>').join('') +
        '</select>' +
        '<input type="text" class="input-field input-sm" id="trk-range-text" placeholder="R3:S5" style="width:120px;">' +
        '</div></div></label>' +
    '</div>' +
    '<div style="display:flex;gap:8px;align-items:center;margin-bottom:12px;">' +
      '<span class="text-muted small">Name</span>' +
      '<input type="text" class="input-field" id="trk-range-name" placeholder="e.g. Min Wage" style="flex:1;">' +
    '</div>' +
    '<div id="trk-range-err" class="text-muted small" style="color:#dc2626;display:none;margin-bottom:8px;"></div>' +
    '<div class="modal-actions"><button class="btn btn-ghost" id="trk-range-cancel">Cancel</button>' +
    '<button class="btn btn-primary" id="trk-range-add">Add</button></div></div>';

  if (!ov) { ov = document.createElement('div'); ov.id = 'trk-range-overlay'; ov.className = 'modal-overlay'; document.body.appendChild(ov); }
  ov.innerHTML = html;
  ov.classList.add('show');
  ov.style.display = 'flex';
  const close = () => { ov.classList.remove('show'); ov.style.display = 'none'; };
  ov.querySelector('#trk-range-cancel').addEventListener('click', close);
  ov.addEventListener('click', e => { if (e.target === ov) close(); });
  const nameBox = ov.querySelector('#trk-range-name');
  const textBox = ov.querySelector('#trk-range-text');
  if (textBox) textBox.addEventListener('focus', () => {
    const r = ov.querySelector('input[value="custom"]'); if (r) r.checked = true;
  });
  ov.querySelector('#trk-range-add').addEventListener('click', () => {
    const err = ov.querySelector('#trk-range-err');
    const pick = ov.querySelector('input[name="trk-range-pick"]:checked');
    if (!pick) { err.textContent = 'Pick a block, or type a range.'; err.style.display = ''; return; }
    let box;
    if (pick.value === 'custom') {
      box = trkParseA1(textBox.value, ov.querySelector('#trk-range-sheet').value);
      if (!box) { err.textContent = 'Could not read that range. Try R3:S5, S, or 3:5.'; err.style.display = ''; return; }
    } else {
      box = blocks[+pick.value];
    }
    if (!trkReadRange(box).length) { err.textContent = 'That range is empty.'; err.style.display = ''; return; }
    trkRefList().push({
      id: 'r' + Date.now() + Math.random().toString(36).slice(2, 6),
      name: (nameBox.value || '').trim() || (box.sheet + '!' + trkRangeLabel(box)),
      box: box
    });
    close();
    trkRenderRefPanel();
  });
}

// ─── Fill Rule editor ───
// One dialog covers the whole job: where the keys come from, how to pull them
// out of the text, what to write, and a typed-in value for everything that
// doesn't resolve. Save it under a name and the next sheet is one click.
// Opening on "column 1, whole cell" is useless — it compares job IDs against
// state names and reports 0 fills with 53 boxes to type in. So try every
// (tracked column × extraction × lookup key column × extraction) pairing and
// open on whichever actually matches the most rows. On this workbook that
// lands on Job Name → text in (brackets) → Arizona (AZ), with no setup at all.
function trkAutoConfigure(rule) {
  const data = trkReadRange(rule.box);
  if (!data.length || !trkHeaders.length) return rule;
  const width = data[0].length;
  const modes = TRK_EXTRACTORS.filter(x => x.id !== 'custom').map(x => x.id);
  const norm = v => String(v || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const sample = trkRows.slice(0, 200);
  let best = null;
  for (let keyCol = 0; keyCol < width; keyCol++) {
    for (const keyMode of modes) {
      const keys = new Set();
      data.forEach(r => { const k = norm(trkExtract(r[keyCol], keyMode)); if (k) keys.add(k); });
      if (!keys.size) continue;
      for (let ci = 0; ci < trkHeaders.length; ci++) {
        for (const srcMode of modes) {
          let hit = 0;
          for (let i = 0; i < sample.length; i++) {
            const k = norm(trkExtract(sample[i][ci], srcMode));
            if (k && keys.has(k)) hit++;
          }
          if (!hit) continue;
          if (!best || hit > best.hit || (hit === best.hit && keys.size > best.keyCount)) {
            best = { hit: hit, keyCol: keyCol, keyMode: keyMode, ci: ci, srcMode: srcMode, keyCount: keys.size };
          }
        }
      }
    }
  }
  if (best) {
    rule.srcCi = best.ci;
    rule.srcExtract = best.srcMode;
    rule.keyCol = best.keyCol;
    rule.keyExtract = best.keyMode;
    rule.valCol = width > 1 ? (best.keyCol === 0 ? 1 : 0) : 0;
    rule.autoHit = best.hit;
  }
  // Write into the first column that is completely empty — that is almost
  // always the one waiting to be filled (here: Rates).
  let emptyCi = -1;
  for (let ci = 0; ci < trkHeaders.length; ci++) {
    if (ci === rule.srcCi) continue;
    if (trkRows.every(r => String(r[ci] == null ? '' : r[ci]).trim() === '')) { emptyCi = ci; break; }
  }
  rule.destCi = emptyCi >= 0 ? emptyCi : (trkHeaders.length ? '__new__' : '__new__');
  return rule;
}

// ─── Auto-suggest ───
// Nobody should have to pick columns and extraction modes by hand. On import,
// try every lookup block in the workbook against every tracked column with
// every extraction, and keep the pairings that actually match rows. What comes
// back is a finished proposal — "fill Rates from the wage table, 14 rows" —
// with one button. The editor is there if you want to argue with it.
function trkSuggestFills() {
  if (!trkHeaders.length || !trkRows.length) return [];
  const modes = TRK_EXTRACTORS.filter(x => x.id !== 'custom').map(x => x.id);
  const norm = v => String(v || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const sample = trkRows.slice(0, 250);

  // Extract once per (column, mode); matching is then just set lookups.
  const srcKeys = [], emptyCol = [];
  for (let ci = 0; ci < trkHeaders.length; ci++) {
    emptyCol[ci] = trkRows.every(r => String(r[ci] == null ? '' : r[ci]).trim() === '');
    srcKeys[ci] = {};
    if (emptyCol[ci]) continue;
    modes.forEach(m => { srcKeys[ci][m] = sample.map(r => norm(trkExtract(r[ci], m))); });
  }
  const targets = [];
  for (let ci = 0; ci < trkHeaders.length; ci++) if (emptyCol[ci]) targets.push(ci);

  if (!targets.length) return [];   // nothing empty to fill — propose nothing

  const ownSheet = (trkSheets[trkActiveSheet] || {}).name;
  const cands = [];
  Object.keys(trkWorkbookGrids).forEach(sn => {
    trkFindBlocks(sn).forEach((b, bi) => {
      if (sn === ownSheet && bi === 0) return;          // that's the tracked table
      const data = trkReadRange(b);
      if (data.length < 2 || !data[0] || data[0].length < 2) return;
      // A lookup is a short reference list. A 2893-row sheet joined on site
      // name is a dataset, and its "value" column is somebody else's data —
      // that is how "fill Miles with a location name" got proposed.
      if (data.length > 200) return;
      const w = data[0].length;
      for (let keyCol = 0; keyCol < w; keyCol++) {
        for (const keyMode of modes) {
          const keys = new Set();
          let numericKeys = 0;
          data.forEach(r => {
            const k = norm(trkExtract(r[keyCol], keyMode));
            if (!k) return;
            if (!keys.has(k) && /^\d+$/.test(k)) numericKeys++;
            keys.add(k);
          });
          if (keys.size < 2) continue;
          // Bare numbers match each other by accident far too easily.
          if (numericKeys / keys.size > 0.5) continue;
          for (let ci = 0; ci < trkHeaders.length; ci++) {
            if (emptyCol[ci]) continue;
            for (const srcMode of modes) {
              const arr = srcKeys[ci][srcMode];
              const seen = new Set();
              let hit = 0;
              for (let i = 0; i < arr.length; i++) {
                if (arr[i] && keys.has(arr[i])) { hit++; seen.add(arr[i]); }
              }
              // How much of the lookup actually gets used is what separates a
              // real join from a coincidence. "Pay Style = Hourly" hitting one
              // key of nineteen matched 55 rows and meant nothing; the wage
              // table matched 3 keys of 3 across 14 rows and meant everything.
              // Three distinct keys is the floor. On two, "CA→41 / AZ→41"
              // scores a perfect coverage and means nothing.
              const coverage = seen.size / keys.size;
              if (seen.size < 3 || coverage < 0.6 || hit < 5) continue;
              cands.push({ sheet: sn, box: b, keyCol: keyCol, keyMode: keyMode, ci: ci, srcMode: srcMode,
                hit: hit, distinct: seen.size, keyCount: keys.size, coverage: coverage,
                score: seen.size * coverage });
            }
          }
        }
      }
    });
  });

  // One proposal per lookup block, best first. The destination is pre-set to
  // the first empty column but stays a dropdown in the banner: which column a
  // rate belongs in is the one thing that cannot be read off the data, so it
  // is offered rather than guessed at silently.
  const usedBlock = new Set(), out = [];
  cands.sort((a, b) => b.score - a.score || b.hit - a.hit);
  for (const s of cands) {
    const bid = s.sheet + '!' + trkRangeLabel(s.box);
    if (usedBlock.has(bid)) continue;
    const dest = targets[Math.min(out.length, targets.length - 1)];
    if (dest == null) break;
    usedBlock.add(bid);
    const proposal = {
      id: 's' + Math.random().toString(36).slice(2, 8),
      box: s.box, rangeLabel: bid,
      srcCi: s.ci, srcExtract: s.srcMode, srcCustom: '',
      keyCol: s.keyCol, keyExtract: s.keyMode, keyCustom: '',
      valCol: s.keyCol === 0 ? 1 : 0,
      destCi: dest, destName: '',
      blanksOnly: true, loose: true, manual: {}, perRow: {},
      name: 'Fill ' + (trkHeaders[dest] || 'column ' + (dest + 1)),
      matched: s.hit, keyCount: s.keyCount, coverage: s.coverage
    };
    // Only offer it if it would actually write something.
    if (!trkPlanFill(proposal).hits.length) { usedBlock.delete(bid); continue; }
    out.push(proposal);
    if (out.length >= 3) break;
  }
  return out;
}

let trkSuggestions = null;
let trkSuggestDismissed = false;

function trkRefreshSuggestions(force) {
  if (force) trkSuggestDismissed = false;
  trkSuggestions = (Object.keys(trkWorkbookGrids).length && !trkSuggestDismissed) ? trkSuggestFills() : [];
  trkRenderSuggestions();
}

function trkRenderSuggestions() {
  const bar = $('trk-suggest');
  if (!bar) return;
  const list = (trkSuggestions || []).filter(s => {
    const p = trkPlanFill(s);
    s.pending = p.hits.length;
    s.leftover = p.groups.length;
    return p.hits.length > 0;
  });
  if (!list.length) { bar.style.display = 'none'; bar.innerHTML = ''; return; }
  bar.style.display = '';
  const total = list.reduce((n, s) => n + s.pending, 0);
  bar.innerHTML =
    '<div class="trk-sg-head"><b>&#10022; ' + list.length + ' fill' + (list.length === 1 ? '' : 's') +
      ' found</b> <span class="text-muted small">' + total + ' cell' + (total === 1 ? '' : 's') +
      ' ready &mdash; nothing to set up</span>' +
      '<span style="margin-left:auto;"></span>' +
      (list.length > 1 ? '<button class="btn btn-primary btn-sm" id="trk-sg-all">Fill all</button>' : '') +
      '<button class="btn btn-ghost btn-sm" id="trk-sg-dismiss" title="Hide these">&times;</button></div>' +
    list.map(s => {
      const destName = s.destCi === '__new__' ? (s.destName || 'new column') : (trkHeaders[s.destCi] || 'column');
      const ex = TRK_EXTRACTORS.find(x => x.id === s.srcExtract);
      const eg = [];
      for (let ri = 0; ri < trkRows.length && eg.length < 2; ri++) {
        const raw = String(trkRows[ri][s.srcCi] || '').trim();
        const k = trkExtract(raw, s.srcExtract, '');
        if (raw && k && !eg.some(e => e.k === k)) eg.push({ raw: raw, k: k });
      }
      const data = trkReadRange(s.box);
      const width = data.length ? data[0].length : 1;
      let valOpts = '';
      for (let i = 0; i < width; i++) {
        if (i === s.keyCol) continue;
        const sm = (data.find(r => r[i]) || [])[i] || '';
        valOpts += '<option value="' + i + '"' + (i === s.valCol ? ' selected' : '') + '>' +
          esc(sm.length > 22 ? sm.slice(0, 22) + '…' : (sm || trkColLetter(s.box.c0 + i))) + '</option>';
      }
      const destOpts = trkHeaders.map((h, i) =>
        '<option value="' + i + '"' + (i === s.destCi ? ' selected' : '') + '>' +
        esc(h || 'column ' + (i + 1)) + '</option>').join('');
      return '<div class="trk-sg-row">' +
        '<div class="trk-sg-txt">' +
          '<span class="text-muted">matched</span> <b>' + esc(trkHeaders[s.srcCi] || 'column') + '</b>' +
          (s.srcExtract !== 'whole' ? ' <span class="text-muted">(' + esc(ex ? ex.label : s.srcExtract) + ')</span>' : '') +
          ' <span class="text-muted">against</span> <code>' + esc(s.rangeLabel) + '</code>' +
          ' <span class="text-muted">&middot; ' + s.pending + ' rows</span>' +
          (eg.length ? '<div class="trk-sg-eg">' + eg.map(e =>
            esc(e.raw.length > 28 ? e.raw.slice(0, 28) + '…' : e.raw) + ' &rarr; <b>' + esc(e.k) + '</b>').join(' &nbsp;·&nbsp; ') +
            (s.leftover ? ' &nbsp;·&nbsp; <span class="text-muted">' + s.leftover + ' left to type in</span>' : '') + '</div>' : '') +
        '</div>' +
        '<span class="trk-sg-pick"><span class="text-muted small">put</span>' +
          '<select class="input-field input-sm trk-sg-val" data-id="' + s.id + '">' + valOpts + '</select>' +
          '<span class="text-muted small">in</span>' +
          '<select class="input-field input-sm trk-sg-dest" data-id="' + s.id + '">' + destOpts + '</select></span>' +
        '<button class="btn btn-primary btn-sm trk-sg-run" data-id="' + s.id + '">Fill</button>' +
        '<button class="btn btn-ghost btn-sm trk-sg-edit" data-id="' + s.id + '">Review…</button>' +
      '</div>';
    }).join('');

  const find = id => (trkSuggestions || []).find(x => x.id === id);
  bar.querySelectorAll('.trk-sg-dest').forEach(sel => sel.addEventListener('change', () => {
    const s = find(sel.dataset.id);
    if (!s) return;
    s.destCi = +sel.value;
    s.name = 'Fill ' + (trkHeaders[s.destCi] || 'column ' + (s.destCi + 1));
    trkRenderSuggestions();
  }));
  bar.querySelectorAll('.trk-sg-val').forEach(sel => sel.addEventListener('change', () => {
    const s = find(sel.dataset.id);
    if (!s) return;
    s.valCol = +sel.value;
    trkRenderSuggestions();
  }));
  bar.querySelectorAll('.trk-sg-run').forEach(b => b.addEventListener('click', () => {
    const s = find(b.dataset.id);
    if (!s) return;
    trkAcceptSuggestion(s);
  }));
  bar.querySelectorAll('.trk-sg-edit').forEach(b => b.addEventListener('click', () => {
    const s = find(b.dataset.id);
    if (!s) return;
    trkPinRange(s.box, s.name);
    trkShowRuleEditor(s, true);
  }));
  const all = $('trk-sg-all');
  if (all) all.addEventListener('click', () => { list.slice().forEach(s => trkAcceptSuggestion(s, true)); trkRefreshSuggestions(); });
  const dis = $('trk-sg-dismiss');
  if (dis) dis.addEventListener('click', () => { trkSuggestDismissed = true; trkSuggestions = []; trkRenderSuggestions(); });
}

// Applying a suggestion also pins its lookup and keeps it as a rule, so the
// leftovers can be typed in later and the whole thing re-run on another sheet.
function trkAcceptSuggestion(s, quiet) {
  trkPinRange(s.box, s.name);
  const list = trkRuleList();
  if (!list.some(r => r.id === s.id)) list.push(s);
  const ok = trkRunRule(s);
  if (!quiet) {
    trkRefreshSuggestions();
    trkRenderRefPanel();
    if (ok && s.leftover) trkShowRuleEditor(s, false);
  }
  return ok;
}

function trkPinRange(box, name) {
  const list = trkRefList();
  const label = box.sheet + '!' + trkRangeLabel(box);
  if (list.some(r => (r.box.sheet + '!' + trkRangeLabel(r.box)) === label)) return;
  list.push({ id: 'r' + Math.random().toString(36).slice(2, 8), name: name || label, box: box });
}

function trkNewRule(rangeId) {
  const rr = trkRefList().find(x => x.id === rangeId);
  if (!rr) return null;
  const data = trkReadRange(rr.box);
  const width = data.length ? data[0].length : 1;
  const rule = {
    id: 'f' + Date.now() + Math.random().toString(36).slice(2, 6),
    name: rr.name, rangeId: rr.id, box: rr.box,
    srcCi: 0, srcExtract: 'whole', srcCustom: '',
    keyCol: 0, keyExtract: 'whole', keyCustom: '',
    valCol: Math.min(1, width - 1),
    destCi: trkHeaders.length ? 0 : '__new__', destName: '',
    blanksOnly: true, loose: true,
    manual: {}, perRow: {}, lastRun: null
  };
  trkAutoConfigure(rule);
  // A range named after its own address makes a poor rule name.
  if (!rr.name || /^.+![A-Z]+\d+:[A-Z]+\d+$/.test(rr.name)) {
    const dest = rule.destCi === '__new__' ? 'new column' : (trkHeaders[rule.destCi] || 'column');
    rule.name = 'Fill ' + dest;
  }
  return rule;
}

function trkShowRuleEditor(rule, isDraft) {
  const data = trkReadRange(rule.box);
  if (!data.length) { alert('That reference range is empty.'); return; }
  const width = data[0].length;
  const extOpts = (sel) => TRK_EXTRACTORS.map(x =>
    '<option value="' + x.id + '"' + (x.id === sel ? ' selected' : '') + '>' + esc(x.label) + '</option>').join('');
  const rangeColOpts = (sel) => {
    let o = '';
    for (let i = 0; i < width; i++) {
      const s = (data.find(r => r[i]) || [])[i] || '';
      o += '<option value="' + i + '"' + (i === sel ? ' selected' : '') + '>' +
        esc(trkColLetter(rule.box.c0 + i) + ' — ' + (s.length > 24 ? s.slice(0, 24) + '…' : s)) + '</option>';
    }
    return o;
  };
  const trackedOpts = (sel, withNew) => {
    let o = trkHeaders.map((h, i) => '<option value="' + i + '"' + (String(i) === String(sel) ? ' selected' : '') + '>' +
      esc(h || '(column ' + (i + 1) + ')') + '</option>').join('');
    if (withNew) o += '<option value="__new__"' + (sel === '__new__' ? ' selected' : '') + '>+ new column…</option>';
    return o;
  };

  let ov = $('trk-fill-overlay');
  if (!ov) { ov = document.createElement('div'); ov.id = 'trk-fill-overlay'; ov.className = 'modal-overlay'; document.body.appendChild(ov); }
  ov.innerHTML =
    '<div class="modal trk-rule-modal"><h3>' + (isDraft ? 'New fill rule' : 'Edit rule') + '</h3>' +
    '<div class="trk-fill-grid">' +
      '<label>Read tracked column</label>' +
        '<div class="trk-inline"><select class="input-field" id="trk-r-src">' + trackedOpts(rule.srcCi, false) + '</select>' +
        '<select class="input-field" id="trk-r-srcex">' + extOpts(rule.srcExtract) + '</select>' +
        '<input type="text" class="input-field trk-pat" id="trk-r-srcpat" placeholder="pattern" value="' + esc(rule.srcCustom || '') + '"></div>' +
      '<label>Match lookup column</label>' +
        '<div class="trk-inline"><select class="input-field" id="trk-r-key">' + rangeColOpts(rule.keyCol) + '</select>' +
        '<select class="input-field" id="trk-r-keyex">' + extOpts(rule.keyExtract) + '</select>' +
        '<input type="text" class="input-field trk-pat" id="trk-r-keypat" placeholder="pattern" value="' + esc(rule.keyCustom || '') + '"></div>' +
      '<label>Write lookup column</label><select class="input-field" id="trk-r-val">' + rangeColOpts(rule.valCol) + '</select>' +
      '<label>Into tracked column</label>' +
        '<div class="trk-inline"><select class="input-field" id="trk-r-dest">' + trackedOpts(rule.destCi, true) + '</select>' +
        '<input type="text" class="input-field" id="trk-r-destname" placeholder="new column name" value="' + esc(rule.destName || '') + '" style="display:none;"></div>' +
    '</div>' +
    '<div class="trk-inline" style="margin:8px 0 10px;">' +
      '<label class="trk-check"><input type="checkbox" id="trk-r-blanks"' + (rule.blanksOnly ? ' checked' : '') + '> only fill blanks</label>' +
      '<label class="trk-check" title="Ignores case, spaces and punctuation when comparing."><input type="checkbox" id="trk-r-loose"' + (rule.loose ? ' checked' : '') + '> loose match</label>' +
    '</div>' +
    '<div id="trk-r-preview" class="trk-fill-preview"></div>' +
    '<div id="trk-r-unmatched"></div>' +
    '<div class="trk-inline" style="margin:10px 0 4px;">' +
      '<span class="text-muted small">Rule name</span>' +
      '<input type="text" class="input-field" id="trk-r-name" value="' + esc(rule.name || '') + '" style="flex:1;">' +
    '</div>' +
    '<div class="modal-actions">' +
      '<button class="btn btn-ghost" id="trk-r-cancel">Cancel</button>' +
      '<button class="btn btn-ghost" id="trk-r-save">Save rule</button>' +
      '<button class="btn btn-primary" id="trk-r-apply">Save &amp; apply</button>' +
    '</div></div>';
  ov.classList.add('show');
  ov.style.display = 'flex';
  const close = () => { ov.classList.remove('show'); ov.style.display = 'none'; };
  ov.querySelector('#trk-r-cancel').addEventListener('click', close);
  ov.addEventListener('click', e => { if (e.target === ov) close(); });

  const q = id => ov.querySelector(id);
  function harvest() {
    rule.srcCi = +q('#trk-r-src').value;
    rule.srcExtract = q('#trk-r-srcex').value;
    rule.srcCustom = q('#trk-r-srcpat').value.trim();
    rule.keyCol = +q('#trk-r-key').value;
    rule.keyExtract = q('#trk-r-keyex').value;
    rule.keyCustom = q('#trk-r-keypat').value.trim();
    rule.valCol = +q('#trk-r-val').value;
    rule.destCi = q('#trk-r-dest').value === '__new__' ? '__new__' : +q('#trk-r-dest').value;
    rule.destName = q('#trk-r-destname').value.trim();
    rule.blanksOnly = q('#trk-r-blanks').checked;
    rule.loose = q('#trk-r-loose').checked;
    // The name follows the destination until you type one of your own —
    // otherwise changing the target leaves a rule called "Fill Earning Code"
    // that writes into Rates.
    if (rule.nameTouched) {
      rule.name = q('#trk-r-name').value.trim() || rule.name;
    } else {
      const dest = rule.destCi === '__new__'
        ? (rule.destName || 'new column')
        : (trkHeaders[rule.destCi] || 'column ' + (Number(rule.destCi) + 1));
      rule.name = 'Fill ' + dest;
      const nb = q('#trk-r-name');
      if (nb && nb.value !== rule.name) nb.value = rule.name;
    }
    q('#trk-r-srcpat').style.display = rule.srcExtract === 'custom' ? '' : 'none';
    q('#trk-r-keypat').style.display = rule.keyExtract === 'custom' ? '' : 'none';
    q('#trk-r-destname').style.display = rule.destCi === '__new__' ? '' : 'none';
  }

  function refresh() {
    harvest();
    const p = trkPlanFill(rule);
    // What the extraction is doing, on real rows — the thing you actually
    // need to see to trust it.
    // Show rows where the extraction WORKED first — three "→ nothing" lines
    // look like a broken rule even when it is filling fine.
    const good = [], bad = [];
    for (let ri = 0; ri < trkRows.length && good.length < 3; ri++) {
      const raw = String(trkRows[ri][rule.srcCi] == null ? '' : trkRows[ri][rule.srcCi]).trim();
      if (!raw) continue;
      const ex = trkExtract(raw, rule.srcExtract, rule.srcCustom);
      const bucket = ex ? good : bad;
      if (bucket.some(s => s.raw === raw) || bucket.length >= 3) continue;
      bucket.push({ raw: raw, ex: ex });
    }
    const shown = good.concat(bad).slice(0, 3);
    // Nothing matching, and a long list of leftovers, almost always means the
    // wrong column or the wrong extraction — say so rather than handing over
    // fifty boxes to type into.
    // Nothing to do because it is already done is not the same as nothing
    // lining up — only the second deserves a warning.
    const allDone = p.hits.length === 0 && p.skippedFilled > 0;
    const looksWrong = p.hits.length === 0 && !allDone && p.groups.length > 5 &&
      !Object.keys(rule.manual).length && !Object.keys(rule.perRow).length;
    q('#trk-r-preview').innerHTML =
      '<b>' + p.hits.length + '</b> cell' + (p.hits.length === 1 ? '' : 's') + ' will be filled' +
      ' &middot; <b>' + p.keys + '</b> key' + (p.keys === 1 ? '' : 's') + ' in the lookup' +
      (p.skippedFilled ? ' &middot; <span class="text-muted">' + p.skippedFilled + ' already filled, left alone</span>' : '') +
      '<div class="trk-fill-sample">' + shown.map(s =>
        esc(s.raw.length > 34 ? s.raw.slice(0, 34) + '…' : s.raw) + ' &rarr; <b>' +
        (s.ex ? esc(s.ex) : '<span class="text-muted">nothing</span>') + '</b>').join('<br>') + '</div>' +
      (looksWrong ? '<div class="trk-warn">Nothing lines up. Check <b>Read tracked column</b> ' +
        'and the two <b>extraction</b> dropdowns &mdash; the values above should look like the lookup&rsquo;s keys.</div>' : '') +
      (allDone ? '<div class="trk-note">Already done &mdash; those ' + p.skippedFilled +
        ' rows hold the right value. Untick <b>only fill blanks</b> to write over them.</div>' : '');

    const box = q('#trk-r-unmatched');
    if (!p.groups.length) {
      box.innerHTML = '<div class="text-muted small" style="margin-top:8px;">Everything resolved &mdash; nothing left to type in.</div>';
      return;
    }
    let rowsLeft = 0; p.groups.forEach(g => rowsLeft += g.rows.length);
    const CAP = 60;
    const shownGroups = p.groups.slice(0, CAP);
    box.innerHTML =
      '<div class="trk-unmatched-head">' + p.groups.length + ' value' + (p.groups.length === 1 ? '' : 's') +
        ' found no match (' + rowsLeft + ' row' + (rowsLeft === 1 ? '' : 's') + ') &mdash; type what they should get:' +
        '<button class="btn btn-ghost btn-sm" id="trk-r-fillall" title="Put the same value in every box below">Set all…</button>' +
      '</div>' +
      (p.groups.length > CAP ? '<div class="text-muted small" style="margin-bottom:4px;">Showing the ' + CAP +
        ' most common; fix the columns above or use <b>Set all…</b> if this list is longer than you expected.</div>' : '') +
      '<div class="trk-unmatched">' + shownGroups.map(g => {
        const cur = g.kind === 'key' ? (rule.manual[g.key] || '') : (rule.perRow[g.raw] || '');
        const sub = g.kind === 'key'
          ? '<div class="trk-um-sub">' + esc(g.samples.join(' · ').slice(0, 90)) + (g.samples.length >= 4 ? ' …' : '') + '</div>'
          : '';
        return '<div class="trk-um-row">' +
          '<div class="trk-um-label"><b>' + esc(g.label || '(blank)') + '</b>' +
            '<span class="text-muted small"> — ' + g.rows.length + ' row' + (g.rows.length === 1 ? '' : 's') + '</span>' + sub + '</div>' +
          '<input type="text" class="input-field trk-um-input" data-kind="' + g.kind + '" ' +
            'data-key="' + esc(g.key) + '" data-raw="' + esc(g.raw) + '" value="' + esc(cur) + '" placeholder="value">' +
          '</div>';
      }).join('') + '</div>';

    box.querySelectorAll('.trk-um-input').forEach(inp => {
      inp.addEventListener('change', () => {
        const v = inp.value.trim();
        if (inp.dataset.kind === 'key') {
          if (v) rule.manual[inp.dataset.key] = v; else delete rule.manual[inp.dataset.key];
        } else {
          if (v) rule.perRow[inp.dataset.raw] = v; else delete rule.perRow[inp.dataset.raw];
        }
        refresh();
      });
    });
    const all = q('#trk-r-fillall');
    if (all) all.addEventListener('click', () => {
      const v = prompt('Value for all ' + p.groups.length + ' unmatched entries:');
      if (v == null || !v.trim()) return;
      p.groups.forEach(g => {
        if (g.kind === 'key') rule.manual[g.key] = v.trim(); else rule.perRow[g.raw] = v.trim();
      });
      refresh();
    });
  }

  ['#trk-r-src', '#trk-r-srcex', '#trk-r-key', '#trk-r-keyex', '#trk-r-val',
   '#trk-r-dest', '#trk-r-blanks', '#trk-r-loose'].forEach(s => q(s).addEventListener('change', refresh));
  ['#trk-r-srcpat', '#trk-r-keypat'].forEach(s => q(s).addEventListener('input', refresh));
  q('#trk-r-name').addEventListener('input', () => { rule.nameTouched = true; });
  q('#trk-r-destname').addEventListener('input', refresh);
  refresh();

  function persist() {
    harvest();
    const list = trkRuleList();
    const i = list.findIndex(r => r.id === rule.id);
    if (i >= 0) list[i] = rule; else list.push(rule);
  }
  q('#trk-r-save').addEventListener('click', () => { persist(); close(); trkRenderRefPanel(); });
  q('#trk-r-apply').addEventListener('click', () => {
    persist();
    const ok = trkRunRule(rule);
    close();
    if (!ok) alert('Nothing to fill — every target cell already holds the right value.');
    trkRenderRefPanel();
  });
}

function trkRunRule(rule) {
  const p = trkPlanFill(rule);
  if (!p.hits.length) return false;
  if (p.isNew) {
    const nm = (rule.destName || '').trim() || rule.name || 'New column';
    const sheet = trkSheets[trkActiveSheet];
    trkHeaders.push(nm);
    trkRows.forEach(r => { while (r.length < trkHeaders.length) r.push(''); });
    if (sheet) {
      sheet.headers = trkHeaders;
      sheet.colOrder = (sheet.colOrder || []).concat([trkHeaders.length - 1]);
      sheet.rows = trkRows;
      trkColOrder = sheet.colOrder;
    }
    // The rule now points at a real column, so re-running won't add another.
    rule.destCi = trkHeaders.length - 1;
  }
  trkApplyEdits(p.hits, 'Fill: ' + (rule.name || 'rule'));
  rule.lastRun = p.hits.length;
  if (p.isNew) trkBuildReview();
  return true;
}

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
  if (!trkRestored && Object.keys(trkSheets).length === 0) {
    trkRestored = true;
    trkRestoreFromLocal();
  }
  // Sync the tracker-page yellow-skip checkbox to the persisted setting and wire it
  // so toggling it here writes the same key the import modal reads.
  { const rb = document.getElementById('trk-btn-ref');
    if (rb && !rb._wired) { rb.addEventListener('click', trkShowAddRangeModal); rb._wired = true; } }
  trkRenderRefPanel();
  trkRefreshSuggestions();
  const cb = document.getElementById('trkSkipYellowToggle');
  if (cb) {
    if (typeof getImportSkipYellow === 'function') cb.checked = getImportSkipYellow();
    if (!cb._wired) {
      cb.addEventListener('change', () => {
        if (typeof setImportSkipYellow === 'function') setImportSkipYellow(cb.checked);
      });
      cb._wired = true;
    }
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
