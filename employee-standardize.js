// ═══════════════════════════════════════════════════════════════════════
// PT 3.0 Employee Standardize & Dedupe — takes a PickTrace Implementation
// Data Template (the "Employees" sheet of new hires to add), cross-references
// an existing 3.0 employee export to suppress duplicates, maps every column
// onto the 3.0 Employee Bulk Create template, snaps values to the template's
// dropdowns, and exports upload-ready xlsx files (pristine template copies so
// all data validations survive). Duplicate detection is corroborated:
//   • CONFIDENT (excluded)  → exact Alt ID, exact SSN, or Name + DOB match
//   • POSSIBLE  (kept/flag) → name-only match with nothing else to confirm
// Rows the client highlighted with a fill color can be left out (or kept
// exclusively) before dedupe — see highlight-filter.js.
// Mirrors the Legacy Employee Migration module's UI / interactions.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');
  const noStar = h => norm(h || '').replace(/\*$/, '').trim();
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // ─── State ───
  let srcData = null;     // { headers, rows, idx, fileName, sheetName } — Data Template Employees sheet
  let srcWb = null;       // last-loaded Data Template workbook (kept so the user can re-pick the sheet)
  let srcFileName = '';
  let dbIndex = null;     // dedupe index built from the existing 3.0 export
  let dbFileName = '';
  let dbCount = 0;
  let tplData = null;     // { headers, dropdowns, employerList, rawBuffer, fileName }
  let outHeaders = null;  // output column headers (from template DATA ENTRY)
  let srcColForOut = null;// per-output-column source index (or -1)
  let formattedRows = null;// [[...]] kept (non-duplicate) rows, parallel to srcKept
  let srcKept = null;      // source rows that survived dedupe (parallel to formattedRows)
  let dupConfident = [];   // [{ kind:'db'|'internal', name, altId, by, existing }] — excluded duplicates
  let dupPossible = [];    // [{ ri, name, altId, existing }] — kept but flagged (ri into formattedRows)
  let columnFills = {};    // colIdx → { val, mode:'all'|'blank' }
  let esSmartFixMap = {};  // "colIdx||UPPERVALUE" → chosen dropdown value (applied Smart Fixes)
  let removedRows = new Set();
  let selCells = new Set();
  let selAnchor = null;
  let previewOrder = [];
  let pvMode = 'all';      // preview rows: 'all' | 'fix' | 'chg' | 'any' (IMWorkbench.filterBar)
  let pvColsOnly = false;  // preview: only the columns with a flag or a change
  let hasCellEdits = false;// a preview cell was hand-edited since the last build
  let hlScan = null;       // IMHighlight scan of the Data Template sheet's fill colors
  let hlState = { mode: 'off', keys: new Set() }; // highlight filter (off / exclude / only)
  // Bulk Update: the template comes pre-filled with the org's employees (keyed
  // by Employee ID). Those rows are the starting point; the Data Template, if
  // loaded, is matched onto them, and only employees that changed export.
  let target = 'create';   // 'create' | 'update'
  let baseRows = null;     // update: the template's employees as PickTrace has them (parallel to formattedRows)
  let fromSrc = null;      // update: Set "ri|ci" of cells the Data Template supplied
  let unmatched = [];      // update: Data Template rows with no employee to update [{ name, altId }]
  let initialized = false;

  // Pane visibility is owned by showBuild() in index.html's router — the old
  // Employees sub-tabs are gone (Standardize & Dedupe now sits under Template
  // Standardize → Employees, Legacy Migration under Legacy → 3.0 → Employees).
  function attachModeSwitcher() {}

  // 'MM/DD/YYYY' / Date / ISO → 'YYYY-MM-DD'. Non-matching values pass through.
  function normDate(s) {
    if (s instanceof Date && !isNaN(s)) {
      return s.getFullYear() + '-' +
        String(s.getMonth() + 1).padStart(2, '0') + '-' +
        String(s.getDate()).padStart(2, '0');
    }
    const t = String(s == null ? '' : s).trim();
    if (!t) return '';
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      let y = m[3];
      if (y.length === 2) y = (+y >= 70 ? '19' : '20') + y;
      return y + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
    }
    // Bare Excel date serial that slipped through (e.g. a DOB read as "38454").
    // 5–6 digit integers only, so 4-digit years pass through untouched. Epoch
    // 1899-12-30 (UTC) accounts for Excel's 1900 leap-year bug.
    if (/^\d{5,6}(\.0+)?$/.test(t)) {
      const n = Math.floor(parseFloat(t));
      if (n >= 10000 && n <= 80000) {
        const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
        if (!isNaN(d.getTime())) {
          return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
        }
      }
    }
    return t;
  }

  function buildCaseMap(arr) {
    const m = new Map();
    (arr || []).forEach(v => {
      const k = String(v).toUpperCase().trim();
      if (k && !m.has(k)) m.set(k, v);
    });
    return m;
  }

  // ─── Smart-match helpers (Smart Fixes panel) ───
  function looseKey(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function pluralKey(s) {
    return String(s == null ? '' : s).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean)
      .map(w => (w.length > 3 && w.endsWith('s')) ? w.slice(0, -1) : w).join('|');
  }
  function buildMatchMaps(vals) {
    const canon = new Map(), loose = new Map(), plural = new Map();
    (vals || []).forEach(v => {
      const u = String(v).toUpperCase().trim();
      if (u && !canon.has(u)) canon.set(u, v);
      const lk = looseKey(v); if (lk) { const a = loose.get(lk) || []; if (!a.includes(v)) a.push(v); loose.set(lk, a); }
      const pk = pluralKey(v); if (pk) { const a = plural.get(pk) || []; if (!a.includes(v)) a.push(v); plural.set(pk, a); }
    });
    return { canon, loose, plural };
  }
  function confidentMatch(value, maps) {
    const uniq = m => (m && m.length === 1) ? m[0] : null;
    return uniq(maps.loose.get(looseKey(value))) || uniq(maps.plural.get(pluralKey(value))) || null;
  }

  function chunkArr(a, n) {
    const out = [];
    for (let i = 0; i < a.length; i += n) out.push(a.slice(i, i + n));
    return out;
  }

  // ─── Dedupe normalizers ───
  const normName = s => String(s == null ? '' : s).toUpperCase().trim().replace(/\s+/g, ' ');
  const normSsn  = s => String(s == null ? '' : s).replace(/\D/g, '');
  const normId   = s => String(s == null ? '' : s).toUpperCase().trim();
  // A "real" SSN for matching purposes: 9 digits, not an all-zero / placeholder
  // run (e.g. 000-00-0558, 000000000). Placeholder SSNs must NOT collapse many
  // distinct people into one duplicate.
  function ssnKey(s) {
    const d = normSsn(s);
    if (d.length !== 9) return '';
    if (/^0+$/.test(d)) return '';
    if (d.slice(0, 5) === '00000') return '';
    return d;
  }
  // True when the new hire and a same-name reference have a CONFLICTING strong
  // identifier (different Alt ID, SSN, or DOB) — i.e. provably different people
  // who merely share a name. Used to suppress false "possible duplicate" flags.
  // If an identifier is missing on either side it can't rule them out.
  function strongConflict(aKey, sKey, dob, ref) {
    if (aKey && ref.aKey && aKey !== ref.aKey) return true;
    if (sKey && ref.sKey && sKey !== ref.sKey) return true;
    if (dob && ref.dob && dob !== ref.dob) return true;
    return false;
  }
  const nameTokens = s => normName(s).split(' ').filter(Boolean);
  // Two names "agree" when they share ≥2 distinct tokens — keeps real matches
  // where a middle name was added/dropped (JANE DOE ≈ JANE MARIE
  // DOE, JOHN SMITH ≈ JOHN PAUL SMITH) but rejects a coincidental
  // ID collision with a totally different person (ALEX X. SAMPLE ≠
  // Test User). Guards against recycled / test Alt IDs and SSNs.
  function nameAgree(aTok, bTok) {
    if (!aTok.length || !bTok.length) return false;
    const B = new Set(bTok);
    let shared = 0;
    new Set(aTok).forEach(t => { if (B.has(t)) shared++; });
    return shared >= 2;
  }

  // ─── CSV / xlsx parsing (raw text for CSV so dates aren't coerced) ───
  function parseCsvText(text) {
    text = String(text).replace(/^﻿/, '');
    const rows = [];
    let i = 0, f = '', row = [], q = false;
    while (i < text.length) {
      const c = text[i];
      if (q) {
        if (c === '"') {
          if (text[i + 1] === '"') { f += '"'; i += 2; continue; }
          q = false; i++; continue;
        }
        f += c; i++; continue;
      }
      if (c === '"') { q = true; i++; continue; }
      if (c === ',') { row.push(f); f = ''; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; i++; continue; }
      f += c; i++;
    }
    if (f.length || row.length) { row.push(f); rows.push(row); }
    return rows;
  }

  // Read a workbook (xlsx/xls) → { wb, aoaForSheet(name) }. CSV → single aoa.
  // withStyles keeps cell fills so the highlight filter can read them.
  function readWorkbook(file, withStyles) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      const isCsv = /\.csv$/i.test(file.name);
      r.onload = e => {
        try {
          if (isCsv) {
            resolve({ csv: true, aoa: parseCsvText(e.target.result) });
          } else {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellStyles: !!withStyles });
            resolve({ csv: false, wb });
          }
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      if (isCsv) r.readAsText(file); else r.readAsArrayBuffer(file);
    });
  }

  // Find the header row in an array-of-arrays by scanning the first ~15 rows
  // for the one containing a "First Name" cell (the Data Template carries a
  // banner + section-header rows above the real columns).
  function findHeaderRow(aoa) {
    const limit = Math.min(aoa.length, 15);
    for (let r = 0; r < limit; r++) {
      const row = aoa[r] || [];
      const hasFirst = row.some(c => noStar(c) === 'first name');
      const hasLast = row.some(c => noStar(c) === 'last name');
      if (hasFirst && hasLast) return r;
    }
    return -1;
  }

  // Read whole DROP-DOWN INPUTS sheet → Map<normHeader, values[]> (verbatim,
  // keeps trailing "ghost spaces" on employer names).
  function readDropdowns(wb) {
    const m = new Map();
    const dropName = wb.SheetNames.find(n => /drop.?down/i.test(n));
    if (!dropName) return m;
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[dropName], { header: 1, defval: '' });
    if (raw.length < 2) return m;
    const dh = raw[0].map(h => String(h == null ? '' : h).trim());
    dh.forEach((h, ci) => {
      if (!h) return;
      const k = norm(h);
      const arr = m.get(k) || [];
      for (let r = 1; r < raw.length; r++) {
        const v = String(raw[r][ci] != null ? raw[r][ci] : '');
        if (v.trim() && arr.indexOf(v) < 0) arr.push(v);
      }
      if (arr.length) m.set(k, arr);
    });
    return m;
  }

  // Per-sheet hidden flags (wb.Workbook.Sheets[i].Hidden: 0 visible, 1 hidden, 2 very-hidden).
  function sheetHiddenFlags(wb) {
    const flags = {};
    const arr = wb && wb.Workbook && wb.Workbook.Sheets;
    if (Array.isArray(arr)) wb.SheetNames.forEach((n, i) => { flags[n] = !!(arr[i] && arr[i].Hidden); });
    return flags;
  }

  // Scan every sheet: detect the employee header row and count new-hire rows.
  // Returns [{ name, aoa, hr, rowCount, hasHeader, hidden }].
  function scanEmployeeSheets(wb) {
    const hidden = sheetHiddenFlags(wb);
    return wb.SheetNames.map(name => {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false });
      const hr = findHeaderRow(aoa);
      let rowCount = 0;
      if (hr >= 0) {
        const H = (aoa[hr] || []).map(h => noStar(h));
        const fi = H.indexOf('first name'), li = H.indexOf('last name');
        rowCount = aoa.slice(hr + 1).filter(row => {
          if (!row) return false;
          const fn = fi >= 0 ? String(row[fi] == null ? '' : row[fi]).trim() : '';
          const ln = li >= 0 ? String(row[li] == null ? '' : row[li]).trim() : '';
          return fn || ln;
        }).length;
      }
      return { name, aoa, hr, rowCount, hasHeader: hr >= 0, hidden: !!hidden[name] };
    });
  }

  // Best default among header-bearing sheets: most rows, then visible, then a
  // name that mentions "employee".
  function bestEmployeeSheet(sheets) {
    return sheets.slice().sort((a, b) =>
      (b.rowCount - a.rowCount) ||
      ((a.hidden ? 1 : 0) - (b.hidden ? 1 : 0)) ||
      ((/employee/i.test(b.name) ? 1 : 0) - (/employee/i.test(a.name) ? 1 : 0))
    )[0];
  }

  // ─── Load handlers ───
  // 1) Data Template — the implementation "Employees" sheet of new hires.
  // Auto-uses the sheet that actually carries employee data; when a workbook has
  // more than one candidate (or none with data), the user picks the sheet.
  function handleSrcFile(file) {
    IMOutName.suggestClient(file.name);
    readWorkbook(file, true).then(d => {
      if (d.csv) {
        srcWb = null; srcFileName = file.name;
        const hr = findHeaderRow(d.aoa);
        if (hr < 0) { alert('Could not find the employee header row ("First Name" / "Last Name") in this CSV.'); return; }
        loadSrcSheet({ name: '', aoa: d.aoa, hr }, file.name);
        return;
      }
      const wb = d.wb;
      srcWb = wb; srcFileName = file.name;
      const all = scanEmployeeSheets(wb);
      const withHeader = all.filter(s => s.hasHeader);
      if (!withHeader.length) {
        alert('Could not find an employee header row ("First Name" / "Last Name") in any sheet of this Data Template. Use the sheet picker if your header uses different names.');
        openSrcSheetPicker(all, file.name);
        return;
      }
      const withData = withHeader.filter(s => s.rowCount > 0);
      if (withData.length === 1) { loadSrcSheet(withData[0], file.name); return; }      // unambiguous
      if (withData.length === 0 && withHeader.length === 1) { loadSrcSheet(withHeader[0], file.name); return; }
      openSrcSheetPicker(all, file.name); // ambiguous → let the user choose
    }).catch(err => alert('Failed to read Data Template: ' + (err && err.message ? err.message : err)));
  }

  // Build srcData from a chosen sheet ({ name, aoa, hr }) and render the meta
  // line with a "Change sheet" affordance when the workbook has more sheets.
  // srcData.allRows keeps every employee row; srcData.rows is what the
  // highlight filter lets through (the same rows when the filter is off).
  function loadSrcSheet(info, fileName) {
    const aoa = info.aoa, hr = info.hr, sheetName = info.name || '';
    const headers = (aoa[hr] || []).map(h => String(h == null ? '' : h).trim());
    const idx = {};
    headers.forEach((h, i) => { const k = noStar(h); if (k && idx[k] == null) idx[k] = i; });
    const fi = idx['first name'], li = idx['last name'];
    const cell = (row, i) => i != null ? String(row[i] == null ? '' : row[i]).trim() : '';
    // aoa index of each employee row — the highlight scan needs the real sheet row.
    const aoaIdxs = [];
    for (let r = hr + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (row && (cell(row, fi) || cell(row, li))) aoaIdxs.push(r);
    }
    const allRows = aoaIdxs.map(r => aoa[r]);
    srcData = { headers, rows: allRows, allRows, idx, fileName, sheetName, headerRow: hr };
    const ws = srcWb && sheetName ? srcWb.Sheets[sheetName] : null;
    hlScan = ws && window.IMHighlight
      ? IMHighlight.scan(srcWb, ws, aoaIdxs, headers, i => (cell(allRows[i], fi) + ' ' + cell(allRows[i], li)).trim())
      : null;
    hlState = { mode: 'off', keys: new Set() };
    $('es-src-name').textContent = fileName + (sheetName ? '  ·  ' + sheetName : '');
    applyHighlight();
    renderSrcMeta();
    renderHighlight();
    maybeRun();
  }

  function renderSrcMeta() {
    if (!srcData) return;
    const n = srcData.allRows.length;
    const notUsed = n - srcData.rows.length;
    const canPick = !!(srcWb && srcWb.SheetNames && srcWb.SheetNames.length > 1);
    $('es-src-meta').innerHTML = n + ' employee row' + (n === 1 ? '' : 's') +
      (notUsed ? '  ·  <b style="color:#b45309;">' + notUsed + ' not used (highlight filter)</b>' : '') +
      '  ·  header at row ' + (srcData.headerRow + 1) +
      (canPick ? '  ·  <button class="btn btn-ghost btn-sm" id="es-src-changesheet" style="padding:1px 8px;">Change sheet</button>' : '');
    if (canPick) {
      const b = $('es-src-changesheet');
      if (b) b.addEventListener('click', () => openSrcSheetPicker(scanEmployeeSheets(srcWb), srcData.fileName));
    }
  }

  // ─── Highlight filter (rows the client marked with a fill color) ───
  function applyHighlight() {
    if (!srcData) return;
    const keep = window.IMHighlight ? IMHighlight.filter(hlScan, hlState) : null;
    srcData.rows = keep ? keep.map(i => srcData.allRows[i]) : srcData.allRows;
  }

  function renderHighlight() {
    if (!window.IMHighlight) return;
    IMHighlight.render($('es-section-highlight'), hlScan, hlState, onHighlightChange,
      { noun: 'employee', labelHeader: 'Employee' });
  }

  // A change that alters which rows are used rebuilds the preview from the
  // source (like re-uploading it), so confirm before discarding work.
  function onHighlightChange(next, rowsChanged) {
    const edited = !!formattedRows && (hasCellEdits || removedRows.size > 0 ||
      Object.keys(columnFills).length > 0 || Object.keys(esSmartFixMap).length > 0);
    if (rowsChanged && edited && !window.confirm(
      'Changing the highlight filter rebuilds the preview from the Data Template.\n\n' +
      'Your bulk fills, Smart Fixes, removed rows and cell edits will be reset. Continue?')) {
      renderHighlight(); // put the controls back the way they were
      return;
    }
    hlState = next;
    applyHighlight();
    renderSrcMeta();
    renderHighlight();
    if (rowsChanged) maybeRun();
  }

  // Sheet picker modal — lists every sheet with its detected employee-row count;
  // sheets without a First/Last Name header are shown disabled.
  function openSrcSheetPicker(sheets, fileName) {
    const prior = document.getElementById('es-sheet-picker');
    if (prior) prior.remove();
    const overlay = document.createElement('div');
    overlay.id = 'es-sheet-picker';
    overlay.className = 'cmp-export-modal';
    overlay.style.display = 'flex';
    const inner = document.createElement('div');
    inner.className = 'cmp-export-modal-inner';
    const best = bestEmployeeSheet(sheets.filter(s => s.hasHeader));
    const items = sheets.map(s => {
      const disabled = !s.hasHeader;
      const sub = s.hasHeader
        ? (s.rowCount + ' employee' + (s.rowCount === 1 ? '' : 's') + (s.hidden ? ' · hidden' : '') + (best && s.name === best.name ? ' · best guess' : ''))
        : 'no First/Last Name header';
      return '<div class="cmp-org-list-item' + (disabled ? ' es-sheet-disabled' : '') + '" data-sheet="' + escHtml(s.name) + '"' +
        (disabled ? ' style="opacity:.5;cursor:not-allowed;"' : '') + '>' +
        '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
        '<span class="cmp-org-counts">' + escHtml(sub) + '</span></div>';
    }).join('');
    inner.innerHTML = '<h3>Select the employee sheet</h3>' +
      '<p>Pick the sheet in <code>' + escHtml(fileName) + '</code> that holds the new-hire rows:</p>' +
      '<div class="cmp-org-list">' + items + '</div>' +
      '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="es-sheet-cancel">Cancel</button></div>';
    overlay.appendChild(inner);
    document.body.appendChild(overlay);
    inner.querySelectorAll('.cmp-org-list-item').forEach(el => {
      if (el.classList.contains('es-sheet-disabled')) return;
      el.addEventListener('click', () => {
        overlay.remove();
        const s = sheets.find(x => x.name === el.dataset.sheet);
        if (s) loadSrcSheet(s, fileName);
      });
    });
    inner.querySelector('#es-sheet-cancel').addEventListener('click', () => overlay.remove());
    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  }

  // 2) Existing employees — a 3.0 DATA ENTRY export → dedupe index.
  function handleDbFile(file) {
    readWorkbook(file).then(d => {
      let aoa;
      if (d.csv) {
        aoa = d.aoa;
      } else {
        const wb = d.wb;
        const sn = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
        aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: false });
      }
      const filtered = aoa.filter(row => row && row.some(c => c != null && String(c).trim() !== ''));
      if (!filtered.length) { alert('Existing employees file has no rows.'); return; }
      const headers = filtered[0].map(h => String(h == null ? '' : h).trim());
      const idx = {};
      headers.forEach((h, i) => { const k = noStar(h); if (k && idx[k] == null) idx[k] = i; });
      if (idx['first name'] == null || idx['last name'] == null) {
        alert('Existing employees file must have First Name / Last Name columns (a 3.0 employee export).');
        return;
      }
      const g = (row, key) => { const i = idx[key]; return i == null ? '' : String(row[i] == null ? '' : row[i]).trim(); };
      const byAlt = new Map(), bySsn = new Map(), byNameDob = new Map(), byName = new Map();
      let n = 0;
      filtered.slice(1).forEach(row => {
        const fn = g(row, 'first name'), ln = g(row, 'last name');
        if (!fn && !ln) return;
        n++;
        const empId = g(row, 'employee id');
        const altId = g(row, 'alt id');
        const ssn = g(row, 'ssn');
        const dob = normDate(g(row, 'date of birth'));
        const nameK = normName(fn + ' ' + ln);
        const display = (fn + ' ' + ln).trim() + (empId ? ' (ID ' + empId + ')' : '');
        const a = normId(altId);
        const sk = ssnKey(ssn);
        // aKey/sKey/dob let a name-only match be ruled out: if the new hire's
        // Alt ID / SSN / DOB conflicts with this record, they're different
        // people despite sharing a name (very common with JOSE LOPEZ etc.).
        const rec = { display, empId, altId, name: (fn + ' ' + ln).trim(), aKey: a, sKey: sk, dob };
        if (a && !byAlt.has(a)) byAlt.set(a, rec);
        if (sk && !bySsn.has(sk)) bySsn.set(sk, rec);
        if (nameK && dob) {
          const k = nameK + '|' + dob;
          if (!byNameDob.has(k)) byNameDob.set(k, rec);
        }
        if (nameK && !byName.has(nameK)) byName.set(nameK, rec);
      });
      dbIndex = { byAlt, bySsn, byNameDob, byName };
      dbCount = n;
      dbFileName = file.name;
      $('es-db-name').textContent = file.name;
      $('es-db-meta').textContent = n + ' existing employee' + (n === 1 ? '' : 's') + ' indexed';
      maybeRun();
    }).catch(err => alert('Failed to read existing employees file: ' + (err && err.message ? err.message : err)));
  }

  // 3) Bulk Create or Bulk Update template — output format + dropdowns.
  // The Update template starts with Employee ID and carries every employee
  // the org has, so it also sets the target and becomes the rows to edit.
  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      try {
        const buf = e.target.result;
        // A plain read: cellStyles throws on PickTrace-generated workbooks.
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
        const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
        const dataWs = wb.Sheets[dataName];
        if (!dataWs) { alert('Could not find DATA ENTRY sheet in this template.'); return; }
        const dataAoa = XLSX.utils.sheet_to_json(dataWs, { header: 1, defval: '', raw: false });
        let headers = dataAoa.length ? dataAoa[0].map(h => String(h == null ? '' : h).trim()) : [];
        while (headers.length && headers[headers.length - 1] === '') headers.pop();
        const isUpdate = noStar(headers[0]) === 'employee id' && headers.some(h => noStar(h) === 'first name');
        const isCreate = noStar(headers[0]) === 'first name' && noStar(headers[4]) === 'employer';
        if (!isUpdate && !isCreate) {
          alert('This does not look like a 3.0 Employee Bulk Create or Bulk Update template (expected First Name… and Employer…, or Employee ID first).');
          return;
        }
        const dropdowns = readDropdowns(wb);
        const employerList = dropdowns.get('employer') || [];
        // The Update template's employees: every row with an Employee ID.
        const rows = isUpdate
          ? dataAoa.slice(1).filter(row => row && String(row[0] == null ? '' : row[0]).trim())
              .map(row => headers.map((h, c) => {
                const v = String(row[c] == null ? '' : row[c]).trim();
                return v && DATE_COLS.has(noStar(h)) ? normDate(v) : v;
              }))
          : [];
        tplData = { headers, dropdowns, employerList, rawBuffer: buf, fileName: file.name, kind: isUpdate ? 'update' : 'create', rows };
        setTarget(tplData.kind);
        $('es-tpl-name').textContent = file.name;
        $('es-tpl-meta').textContent = (isUpdate ? 'Bulk Update · ' + rows.length + ' employee' + (rows.length === 1 ? '' : 's') + '  ·  ' : 'Bulk Create  ·  ') +
          headers.length + ' columns  ·  ' + employerList.length + ' employer' + (employerList.length === 1 ? '' : 's');
        maybeRun();
      } catch (err) {
        alert('Failed to read template: ' + (err && err.message ? err.message : err));
      }
    };
    r.onerror = () => alert('Failed to read template file.');
    r.readAsArrayBuffer(file);
  }

  // ─── Column mapping (Data Template header → output column) ───
  // Source-header aliases that don't match the output header by name.
  const SRC_ALIASES = {
    'alt id (payroll id #)': 'alt id',
    'payroll id #': 'alt id',
    'payroll id': 'alt id',
    'hourly rate': 'compensation'
  };
  // Output columns we never auto-fill from the source even if a name collides.
  const DATE_COLS = new Set(['date of birth', 'hire date', 'start date', 'work authorization expiry date']);

  // For each output column, the source column index that feeds it (or -1).
  function buildColMap() {
    srcColForOut = outHeaders.map(h => {
      const target = noStar(h);
      // Look for a source header that maps (directly or via alias) to this output column.
      let found = -1;
      srcData.headers.forEach((sh, si) => {
        if (found >= 0) return;
        let k = noStar(sh);
        if (SRC_ALIASES[k]) k = SRC_ALIASES[k];
        if (k === target) found = si;
      });
      return found;
    });
  }

  // ─── Allowed-value snapping ───
  const COUNTRY_ALIASES = {
    'usa': 'US', 'u.s.': 'US', 'u.s.a.': 'US', 'united states': 'US',
    'united states of america': 'US', 'america': 'US',
    'mexico': 'MX', 'méxico': 'MX', 'canada': 'CA'
  };
  // Common language names → the codes PickTrace's Language dropdown uses
  // (en-US / es-MX). Fixes "Unsupported 'Language Preference'" when the source
  // has "ENGLISH" / "SPANISH" instead of the code.
  const LANGUAGE_ALIASES = {
    'english': 'en-US', 'en': 'en-US', 'en-us': 'en-US', 'eng': 'en-US',
    'spanish': 'es-MX', 'espanol': 'es-MX', 'español': 'es-MX', 'es': 'es-MX', 'es-mx': 'es-MX', 'spa': 'es-MX'
  };
  function snapToOption(val, cm, aliases) {
    const v = String(val == null ? '' : val).trim();
    if (!v || !cm.size) return v;
    const hit = cm.get(v.toUpperCase());
    if (hit != null) return hit;
    if (aliases) {
      const a = aliases[v.toLowerCase()];
      if (a) { const h2 = cm.get(String(a).toUpperCase()); if (h2 != null) return h2; }
    }
    return v;
  }
  // The alias table (if any) that applies to a given output column.
  function aliasesFor(header) {
    const h = noStar(header);
    if (/country$/i.test(header)) return COUNTRY_ALIASES;
    if (h === 'language preference' || h === 'language') return LANGUAGE_ALIASES;
    return null;
  }

  function colOptions(header) {
    const h = noStar(header);
    const dd = tplData && tplData.dropdowns;
    if (dd) {
      if (dd.get(h)) return dd.get(h);
      const alias = { 'language preference': 'language', 'h2a employee': 'h2a' };
      if (alias[h] && dd.get(alias[h])) return dd.get(alias[h]);
    }
    if (h === 'employer') return (tplData && tplData.employerList) || [];
    if (h === 'gender') return ['MALE', 'FEMALE'];
    if (h === 'h2a employee' || h === 'h2a contract') return ['TRUE', 'FALSE'];
    if (h === 'language preference') return ['en-US', 'es-MX'];
    return [];
  }

  // ─── Build: map → dedupe → snap ───
  function buildFormattedRows() {
    if (!tplData || (!srcData && target !== 'update')) return; // existing-employees export (dbIndex) is optional
    outHeaders = tplData.headers.slice();
    if (target === 'update') buildUpdateRows();
    else buildCreateRows();
    snapAndFill();
  }

  // ─── Bulk Update rows ───
  // Every employee the Update template lists, as PickTrace has them. A Data
  // Template, when loaded, is matched onto them — by Employee ID if it has
  // one, else Alt ID or SSN (with the name agreeing), else name + birth date
  // — and its non-empty cells overwrite. Rows it can't place are listed:
  // they're new people, for a Bulk Create.
  function buildUpdateRows() {
    baseRows = tplData.rows.map(r => r.slice());
    formattedRows = tplData.rows.map(r => r.slice());
    srcKept = null;
    fromSrc = new Set();
    unmatched = [];
    dupConfident = [];
    dupPossible = [];
    if (!srcData) { srcColForOut = outHeaders.map(() => -1); return; }
    buildColMap();
    const at = k => outHeaders.findIndex(h => noStar(h) === k);
    const idI = at('employee id'), fnI = at('first name'), lnI = at('last name'), altI = at('alt id'), ssnI = at('ssn'), dobI = at('date of birth');
    const byId = new Map(), byAlt = new Map(), bySsn = new Map(), byNameDob = new Map();
    baseRows.forEach((r, i) => {
      const nk = normName(r[fnI] + ' ' + r[lnI]);
      if (idI >= 0 && r[idI]) byId.set(normId(r[idI]), i);
      if (altI >= 0 && r[altI] && !byAlt.has(normId(r[altI]))) byAlt.set(normId(r[altI]), i);
      const sk = ssnI >= 0 ? ssnKey(r[ssnI]) : '';
      if (sk && !bySsn.has(sk)) bySsn.set(sk, i);
      if (nk && dobI >= 0 && r[dobI] && !byNameDob.has(nk + '|' + r[dobI])) byNameDob.set(nk + '|' + r[dobI], i);
    });
    const get = (row, si) => si < 0 ? '' : String(row[si] == null ? '' : row[si]).trim();
    const nameOk = (i, tok) => nameAgree(tok, nameTokens(baseRows[i][fnI] + ' ' + baseRows[i][lnI]));
    srcData.rows.forEach(srcRow => {
      const v = outHeaders.map((h, c) => {
        const x = get(srcRow, srcColForOut[c]);
        return x && DATE_COLS.has(noStar(h)) ? normDate(x) : x;
      });
      const name = ((fnI >= 0 ? v[fnI] : '') + ' ' + (lnI >= 0 ? v[lnI] : '')).trim();
      const tok = nameTokens(name), nk = normName(name);
      let bi = idI >= 0 && v[idI] ? byId.get(normId(v[idI])) : undefined;
      if (bi == null && altI >= 0 && v[altI] && byAlt.has(normId(v[altI])) && nameOk(byAlt.get(normId(v[altI])), tok)) bi = byAlt.get(normId(v[altI]));
      if (bi == null && ssnI >= 0 && ssnKey(v[ssnI]) && bySsn.has(ssnKey(v[ssnI])) && nameOk(bySsn.get(ssnKey(v[ssnI])), tok)) bi = bySsn.get(ssnKey(v[ssnI]));
      if (bi == null && nk && dobI >= 0 && v[dobI]) bi = byNameDob.get(nk + '|' + v[dobI]);
      if (bi == null) { unmatched.push({ name: name || '(no name)', altId: altI >= 0 ? v[altI] : '' }); return; }
      v.forEach((x, c) => {
        if (!x || c === idI) return;
        formattedRows[bi][c] = x;
        fromSrc.add(bi + '|' + c);
      });
    });
  }

  // The value an Update file carries for a cell: '<CLEAR>' where PickTrace
  // has a value and you emptied it (a blank cell would leave it unchanged).
  const CLEAR = '<CLEAR>';
  function updateCell(ri, ci) {
    const now = formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci]);
    return !now && baseRows && baseRows[ri] && baseRows[ri][ci] ? CLEAR : now;
  }
  // An Update row worth sending: something differs from what PickTrace has.
  function rowChanged(ri) {
    if (!baseRows || !baseRows[ri]) return true;
    const t = v => String(v == null ? '' : v).trim().toUpperCase(); // a case-only snap isn't a change
    return formattedRows[ri].some((v, ci) => t(v) !== t(baseRows[ri][ci]));
  }

  function setTarget(t) {
    target = t === 'update' ? 'update' : 'create';
    document.querySelectorAll('input[name="es-target"]').forEach(r => { r.checked = r.value === target; });
    document.querySelectorAll('.es-create-only').forEach(el => { el.style.display = target === 'update' ? 'none' : ''; });
    const src = $('es-src-req');
    if (src) src.textContent = target === 'update' ? '(optional)' : '(required)';
    const empty = $('es-empty');
    if (empty) empty.textContent = target === 'update'
      ? 'Upload the Bulk Update template — its employees are the rows to edit. A Data Template is optional: it fills in changes for the employees it matches.'
      : 'Upload the Data Template and the Bulk Create template (Existing Employees is optional) to start.';
  }

  function buildCreateRows() {
    baseRows = null; fromSrc = null; unmatched = [];
    buildColMap();
    const get = (row, si) => si < 0 ? '' : String(row[si] == null ? '' : row[si]).trim();
    const altOut = outHeaders.findIndex(h => noStar(h) === 'alt id');
    const ssnOut = outHeaders.findIndex(h => noStar(h) === 'ssn');
    const dobOut = outHeaders.findIndex(h => noStar(h) === 'date of birth');
    const fnOut = outHeaders.findIndex(h => noStar(h) === 'first name');
    const lnOut = outHeaders.findIndex(h => noStar(h) === 'last name');

    formattedRows = [];
    srcKept = [];
    dupConfident = [];
    dupPossible = [];

    // Running index of rows ALREADY accepted into the creation list, so the
    // list itself never contains the same person twice (same corroborated
    // match rules as the DB check, applied row-to-row within the template).
    const seenAlt = new Map(), seenSsn = new Map(), seenNameDob = new Map(), seenName = new Map();

    srcData.rows.forEach(srcRow => {
      const r = new Array(outHeaders.length).fill('');
      for (let c = 0; c < outHeaders.length; c++) {
        const si = srcColForOut[c];
        if (si < 0) continue;
        let v = get(srcRow, si);
        if (v && DATE_COLS.has(noStar(outHeaders[c]))) v = normDate(v);
        r[c] = v;
      }

      const altId = altOut >= 0 ? r[altOut] : '';
      const ssn = ssnOut >= 0 ? r[ssnOut] : '';
      const dob = dobOut >= 0 ? r[dobOut] : '';
      const fn = fnOut >= 0 ? r[fnOut] : '';
      const ln = lnOut >= 0 ? r[lnOut] : '';
      const nameK = normName(fn + ' ' + ln);
      const dob_k = nameK && dob ? nameK + '|' + dob : '';
      const dispName = (fn + ' ' + ln).trim();
      const a = normId(altId);
      const sk = ssnKey(ssn);
      const tok = nameTokens(dispName);

      // Resolve a confident match (DB first, then within-file). A single-ID
      // match (Alt ID / SSN) is only confident when the NAME also agrees —
      // otherwise it's a recycled/test ID colliding with a different person,
      // so we keep the row and flag it for review instead of dropping it.
      // Name + DOB needs no extra name check (the name already matches).
      // Returns { kind, by, existing } for a confident dup, a { conflict }
      // marker for an ID collision to review, or null.
      function classify() {
        // existing DB (only when an existing-employees export was uploaded)
        if (a && dbIndex && dbIndex.byAlt.has(a)) {
          const m = dbIndex.byAlt.get(a);
          if (nameAgree(tok, nameTokens(m.name))) return { kind: 'db', by: 'Alt ID ' + altId, existing: m.display };
          return { conflict: { by: 'Alt ID ' + altId + ' — name differs', existing: m.display } };
        }
        if (sk && dbIndex && dbIndex.bySsn.has(sk)) {
          const m = dbIndex.bySsn.get(sk);
          if (nameAgree(tok, nameTokens(m.name))) return { kind: 'db', by: 'SSN', existing: m.display };
          return { conflict: { by: 'SSN — name differs', existing: m.display } };
        }
        if (dob_k && dbIndex && dbIndex.byNameDob.has(dob_k)) {
          const m = dbIndex.byNameDob.get(dob_k);
          // SSN is the tie-breaker: same name + DOB but a DIFFERENT SSN means
          // two different people, so don't auto-exclude — keep and flag.
          if (sk && m.sKey && sk !== m.sKey)
            return { conflict: { by: 'Name + DOB — SSN differs', existing: m.display } };
          return { kind: 'db', by: 'Name + DOB', existing: m.display };
        }
        // within the creation list
        if (a && seenAlt.has(a)) {
          const s = seenAlt.get(a);
          if (nameAgree(tok, nameTokens(s.name))) return { kind: 'internal', by: 'Alt ID ' + altId, existing: s.disp };
          return { conflict: { by: 'Alt ID ' + altId + ' — name differs', existing: s.disp + ' — earlier in this file' } };
        }
        if (sk && seenSsn.has(sk)) {
          const s = seenSsn.get(sk);
          if (nameAgree(tok, nameTokens(s.name))) return { kind: 'internal', by: 'SSN', existing: s.disp };
          return { conflict: { by: 'SSN — name differs', existing: s.disp + ' — earlier in this file' } };
        }
        if (dob_k && seenNameDob.has(dob_k)) {
          const s = seenNameDob.get(dob_k);
          if (sk && s.sKey && sk !== s.sKey)
            return { conflict: { by: 'Name + DOB — SSN differs', existing: s.disp + ' — earlier in this file' } };
          return { kind: 'internal', by: 'Name + DOB', existing: s.disp };
        }
        return null;
      }
      const cls = classify();
      if (cls && cls.kind) {
        dupConfident.push({ kind: cls.kind, name: dispName, altId, by: cls.by, existing: cls.existing });
        return; // confident duplicate — excluded from the export
      }

      // ── Accept the row ──
      const ri = formattedRows.length;
      formattedRows.push(r);
      srcKept.push(srcRow);
      const disp = dispName + (altId ? ' (Alt ID ' + altId + ')' : '');
      if (a && !seenAlt.has(a)) seenAlt.set(a, { disp, name: dispName });
      if (sk && !seenSsn.has(sk)) seenSsn.set(sk, { disp, name: dispName });
      if (dob_k && !seenNameDob.has(dob_k)) seenNameDob.set(dob_k, { disp, name: dispName, sKey: sk });

      // Flag for review (kept in export): an ID collision with a different
      // name, a Name+DOB match whose SSN contradicts it, or a name-only match
      // that nothing rules out. cls.conflict.by carries the exact reason.
      if (cls && cls.conflict) {
        dupPossible.push({ ri, name: dispName, altId, reason: cls.conflict.by, sev: 'conflict', existing: cls.conflict.existing });
      } else if (nameK) {
        let possibleRef = '';
        const dbRef = dbIndex ? dbIndex.byName.get(nameK) : null;
        if (dbRef && !strongConflict(a, sk, dob, dbRef)) possibleRef = dbRef.display;
        else {
          const fileRef = seenName.get(nameK);
          if (fileRef && !strongConflict(a, sk, dob, fileRef)) possibleRef = fileRef.disp + ' — earlier in this file';
        }
        if (possibleRef) dupPossible.push({ ri, name: dispName, altId, reason: 'Name only', sev: 'name', existing: possibleRef });
      }
      if (nameK && !seenName.has(nameK)) seenName.set(nameK, { disp, name: dispName, aKey: a, sKey: sk, dob });
    });
  }

  // Both targets: snap to the dropdowns, then your Smart Fixes and fills.
  function snapAndFill() {
    // Snap dropdown-backed columns to the template's allowed values.
    for (let ci = 0; ci < outHeaders.length; ci++) {
      const opts = colOptions(outHeaders[ci]);
      if (!opts || !opts.length) continue;
      const aliases = aliasesFor(outHeaders[ci]);
      const cm = buildCaseMap(opts);
      formattedRows.forEach(r => { if (r[ci]) r[ci] = snapToOption(r[ci], cm, aliases); });
    }

    // Apply accepted Smart Fixes (bulk snap of off-list dropdown values the user
    // resolved in the Smart Fixes panel). After the automatic snap, before the
    // explicit bulk fills (which always win).
    if (Object.keys(esSmartFixMap).length) {
      for (let ci = 0; ci < outHeaders.length; ci++) {
        formattedRows.forEach(r => {
          const v = r[ci];
          if (!v) return;
          const hit = esSmartFixMap[ci + '||' + String(v).toUpperCase().trim()];
          if (hit) r[ci] = hit;
        });
      }
    }

    // Re-apply sticky bulk fills.
    Object.keys(columnFills).forEach(k => {
      const ci = +k, { val, mode } = columnFills[k];
      formattedRows.forEach((r, i) => {
        if (removedRows.has(i)) return;
        if (mode === 'blank') { if (!r[ci]) r[ci] = val; }
        else r[ci] = val;
      });
    });
  }

  function runStandardize() {
    if (!tplData || (!srcData && target !== 'update')) {
      alert(target === 'update' ? 'Upload the Bulk Update template first.'
        : 'Upload the Data Template and the Bulk Create template first. (Existing Employees is optional.)');
      return;
    }
    columnFills = {};
    removedRows = new Set();
    esSmartFixMap = {};
    hasCellEdits = false;
    selCells = new Set(); selAnchor = null;
    buildFormattedRows();
    renderDuplicates();
    renderPossible();
    renderUnmatched();
    renderEmployersToCreate();
    renderSmartFixes();
    renderBulkEdit();
    renderPreview();
    updateSummary();
    $('es-empty').style.display = 'none';
  }

  // Update: Data Template rows that match no employee in the template.
  function renderUnmatched() {
    const sec = $('es-section-unmatched');
    if (!sec) return;
    if (target !== 'update' || !unmatched.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('es-unmatched-title').textContent = 'Not in PickTrace (' + unmatched.length + ')';
    $('es-unmatched-table').innerHTML = '<thead><tr><th>Employee</th><th>Alt ID</th></tr></thead><tbody>' +
      unmatched.map(u => '<tr><td>' + escHtml(u.name) + '</td><td>' + escHtml(u.altId || '') + '</td></tr>').join('') + '</tbody>';
  }

  function maybeRun() {
    // Existing Employees (dbIndex) is optional; so is the Data Template for an Update.
    const ready = !!(tplData && (srcData || target === 'update'));
    $('es-run').disabled = !ready;
    if (ready) runStandardize();
  }

  // ─── Duplicates (confident, excluded) ───
  function renderDuplicates() {
    const sec = $('es-section-dups');
    const tbl = $('es-dups-table');
    if (!sec || !tbl) return;
    if (!dupConfident.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const inDb = dupConfident.filter(d => d.kind === 'db').length;
    const internal = dupConfident.length - inDb;
    const title = $('es-dups-title');
    if (title) title.textContent = 'Duplicates Excluded (' + dupConfident.length + ')' +
      (internal ? '  ·  ' + inDb + ' in database, ' + internal + ' within file' : '');
    let html = '<thead><tr><th>New employee</th><th>Alt ID</th><th>Type</th>' +
      '<th>Matched on</th><th>Duplicate of</th></tr></thead><tbody>';
    dupConfident.forEach(d => {
      const type = d.kind === 'db'
        ? '<span style="color:#dc2626;font-weight:600;">In database</span>'
        : '<span style="color:#7c3aed;font-weight:600;">Within file</span>';
      html += '<tr>' +
        '<td><b>' + escHtml(d.name || '—') + '</b></td>' +
        '<td>' + escHtml(d.altId || '') + '</td>' +
        '<td>' + type + '</td>' +
        '<td><span style="color:#15803d;font-weight:600;">' + escHtml(d.by) + '</span></td>' +
        '<td>' + escHtml(d.existing) + '</td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
  }

  // ─── Possible duplicates (name-only, kept + flagged) ───
  function renderPossible() {
    const sec = $('es-section-possible');
    const tbl = $('es-possible-table');
    if (!sec || !tbl) return;
    const live = dupPossible.filter(d => !removedRows.has(d.ri));
    if (!live.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const title = $('es-possible-title');
    if (title) title.textContent = 'Possible Duplicates (' + live.length + ')';
    let html = '<thead><tr><th>New employee</th><th>Alt ID</th>' +
      '<th>Matched on</th><th>Existing employee</th><th></th></tr></thead><tbody>';
    live.forEach(d => {
      const color = d.sev === 'conflict' ? '#dc2626' : '#b45309';
      const reason = '<span style="color:' + color + ';font-weight:600;">' + escHtml(d.reason || 'Review') + '</span>';
      html += '<tr>' +
        '<td><b>' + escHtml(d.name || '—') + '</b></td>' +
        '<td>' + escHtml(d.altId || '') + '</td>' +
        '<td>' + reason + '</td>' +
        '<td>' + escHtml(d.existing) + '</td>' +
        '<td><button class="btn btn-ghost btn-sm es-possible-drop" data-ri="' + d.ri +
          '" title="Exclude this row from the export">Exclude</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.es-possible-drop').forEach(btn => {
      btn.addEventListener('click', e => {
        removedRows.add(+e.currentTarget.dataset.ri);
        renderPossible();
        renderPreview();
        updateSummary();
      });
    });
  }

  // ─── Smart Fixes — off-list dropdown values with a "Set to" picker ───
  // Scans every dropdown-backed output column (except Employer, which has its
  // own panel) for values that aren't an allowed option. A confident match
  // (case / spacing / singular↔plural) is pre-selected; the rest ask the user.
  function computeSmartFixes() {
    if (!formattedRows || !tplData) return [];
    const out = [];
    for (let ci = 0; ci < outHeaders.length; ci++) {
      const h = noStar(outHeaders[ci]);
      if (h === 'employer') continue; // covered by "Employers to Create"
      const opts = colOptions(outHeaders[ci]);
      if (!opts || !opts.length) continue;
      const maps = buildMatchMaps(opts);
      const aliases = aliasesFor(outHeaders[ci]);
      const seen = new Map();
      formattedRows.forEach((r, i) => {
        if (removedRows.has(i)) return;
        const val = r[ci];
        if (!val) return;
        const u = String(val).toUpperCase().trim();
        if (maps.canon.has(u)) return;                              // already valid
        if (aliases && aliases[String(val).toLowerCase().trim()]) return; // auto-snapped by alias
        if (esSmartFixMap[ci + '||' + u]) return;                   // already applied
        const sug = confidentMatch(val, maps);
        const suggestion = (sug && String(sug).toUpperCase().trim() !== u) ? sug : '';
        const e = seen.get(u) || { colIdx: ci, header: outHeaders[ci], from: val, count: 0, options: opts, suggestion };
        e.count++;
        seen.set(u, e);
      });
      seen.forEach(e => out.push(e));
    }
    out.sort((a, b) => (b.suggestion ? 1 : 0) - (a.suggestion ? 1 : 0) || b.count - a.count);
    return out;
  }

  function renderSmartFixes() {
    const sec = $('es-section-smartfix');
    if (!sec) return;
    const fixes = computeSmartFixes();
    if (!fixes.length) { sec.style.display = 'none'; $('es-smartfix-table').innerHTML = ''; return; }
    sec.style.display = '';
    let totalRows = 0, needPick = 0;
    fixes.forEach(f => { totalRows += f.count; if (!f.suggestion) needPick++; });
    const t = $('es-smartfix-title');
    if (t) t.textContent = 'Smart Fixes (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + totalRows + ' rows' +
      (needPick ? ' — ' + needPick + ' need your choice' : '') + ')';
    let html = '<thead><tr><th>Column</th><th>Value in your data</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    fixes.forEach(f => {
      const opts = '<option value="">— pick a value —</option>' +
        f.options.slice().sort().map(o => '<option' + (f.suggestion && o === f.suggestion ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('');
      html += '<tr>' +
        '<td>' + escHtml(noStar(f.header)) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td><select class="es-smartfix-pick input-field" style="min-width:180px;">' + opts + '</select>' +
          (f.suggestion ? ' <span class="text-muted small">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') +
        '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm es-smartfix-apply" data-ci="' + f.colIdx + '" data-from="' + escHtml(f.from) + '">Apply</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    $('es-smartfix-table').innerHTML = html;
    $('es-smartfix-table').querySelectorAll('.es-smartfix-apply').forEach(btn => btn.addEventListener('click', e => {
      const tr = e.currentTarget.closest('tr');
      const to = tr.querySelector('.es-smartfix-pick').value.trim();
      if (!to) { alert('Pick a value to set "' + e.currentTarget.dataset.from + '" to first.'); return; }
      applySmartFix(+e.currentTarget.dataset.ci, e.currentTarget.dataset.from, to);
    }));
  }
  function reRenderAfterFix() {
    buildFormattedRows();
    renderDuplicates(); renderPossible(); renderEmployersToCreate();
    renderSmartFixes(); renderBulkEdit(); renderPreview(); updateSummary();
  }
  function applySmartFix(ci, from, to) {
    esSmartFixMap[ci + '||' + String(from).toUpperCase().trim()] = to;
    reRenderAfterFix();
  }
  function applyAllSmartFixes() {
    const tbl = $('es-smartfix-table');
    if (!tbl) return;
    const picks = [];
    tbl.querySelectorAll('.es-smartfix-apply').forEach(btn => {
      const tr = btn.closest('tr');
      const to = tr.querySelector('.es-smartfix-pick').value.trim();
      if (to) picks.push({ ci: +btn.dataset.ci, from: btn.dataset.from, to });
    });
    if (!picks.length) { alert('No values chosen yet. Pick a target for at least one row (suggested rows are pre-selected).'); return; }
    picks.forEach(p => { esSmartFixMap[p.ci + '||' + String(p.from).toUpperCase().trim()] = p.to; });
    reRenderAfterFix();
  }

  // ─── Employers to create (in column but not in template dropdown) ───
  function employerValid(v) {
    if (!v) return false;
    if (!tplData || !tplData.employerList || !tplData.employerList.length) return true;
    return buildCaseMap(tplData.employerList).has(String(v).toUpperCase().trim());
  }

  function getEmployersToCreate() {
    const m = new Map();
    if (!formattedRows) return m;
    const ei = outHeaders.findIndex(h => noStar(h) === 'employer');
    if (ei < 0) return m;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const v = r[ei];
      if (!v || employerValid(v)) return;
      const e = m.get(v) || { count: 0 };
      e.count++;
      m.set(v, e);
    });
    return m;
  }

  function renderEmployersToCreate() {
    const sec = $('es-section-create');
    const tbl = $('es-create-table');
    if (!sec || !tbl) return;
    const m = getEmployersToCreate();
    if (!m.size) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const title = $('es-create-title');
    if (title) title.textContent = 'Employers to Create in 3.0 (' + m.size + ')';
    let html = '<thead><tr><th>Employer (not in template dropdown)</th><th>Employees</th></tr></thead><tbody>';
    [...m.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([name, info]) => {
      html += '<tr><td><b>' + escHtml(name) + '</b></td><td>' + info.count + '</td></tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
  }

  // ─── Bulk-edit columns ───
  function renderBulkEdit() {
    if (!formattedRows || !tplData) return;
    const sec = $('es-section-bulk');
    const tbl = $('es-bulk-table');
    sec.style.display = '';
    let html = '<thead><tr><th>Column</th><th>Current bulk value</th>' +
      '<th>Pick</th><th>Manual override</th><th></th><th></th></tr></thead><tbody>';
    outHeaders.forEach((h, idx) => {
      const opts = colOptions(h);
      const cur = columnFills[idx];
      const optsHtml = opts.length
        ? '<option value="">— pick —</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (cur && o === cur.val ? ' selected' : '') +
            '>' + escHtml(o) + (/\s$/.test(o) ? ' ␣(trailing space)' : '') + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const status = cur
        ? '<span style="color:#15803d;font-weight:600;">✓ ' + escHtml(String(cur.val).slice(0, 40) ||
            (target === 'update' ? 'cleared — ' + CLEAR : '(cleared)')) +
          '</span> <span class="text-muted small">(' + (cur.mode === 'blank' ? 'blanks only' : 'all rows') + ')</span>' +
          ' <button class="btn btn-ghost btn-sm es-bulk-undo" data-idx="' + idx + '" title="Drop this fill and put the column back">Undo</button>'
        : '<span class="text-muted small">—</span>';
      const req = /\*$/.test(h);
      html += '<tr>' +
        '<td><b' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="es-bulk-pick input-field" data-idx="' + idx + '" style="min-width:180px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="es-bulk-override input-field" data-idx="' + idx + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm es-bulk-apply" data-idx="' + idx + '" title="Overwrite this column for every row">Apply all</button> ' +
        '<button class="btn btn-success btn-sm es-bulk-fillblank" data-idx="' + idx + '" title="Fill only the empty cells in this column">Fill blanks</button></td>' +
        '<td><button class="btn btn-ghost btn-sm es-bulk-clear" data-idx="' + idx + '" title="Empty this column for all rows">Clear</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    wireBulkEditHandlers();
  }

  function applyColumnFill(idx, val, mode) {
    columnFills[idx] = { val, mode: mode === 'blank' ? 'blank' : 'all' };
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (mode === 'blank') { if (!r[idx]) r[idx] = val; }
      else r[idx] = val;
    });
    renderBulkEdit();
    if (noStar(outHeaders[idx]) === 'employer') renderEmployersToCreate();
    renderPreview();
    updateSummary();
  }

  // Clear sticks like any fill (an empty value for every row), so a rebuild
  // doesn't bring the column back; in an Update it exports as <CLEAR>.
  // Undo drops the fill and rebuilds the column from its source.
  function clearColumnFill(idx) {
    columnFills[idx] = { val: '', mode: 'all' };
    formattedRows.forEach((r, ri) => { if (!removedRows.has(ri)) r[idx] = ''; });
    renderBulkEdit();
    if (noStar(outHeaders[idx]) === 'employer') renderEmployersToCreate();
    renderPreview();
    updateSummary();
  }

  function wireBulkEditHandlers() {
    const tbl = $('es-bulk-table');
    if (!tbl) return;
    const grab = e => {
      const tr = e.target.closest('tr');
      const pick = tr.querySelector('.es-bulk-pick').value;
      const ovr = tr.querySelector('.es-bulk-override').value.replace(/^\s+/, '');
      return ovr || pick;
    };
    tbl.querySelectorAll('.es-bulk-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'all');
      });
    });
    tbl.querySelectorAll('.es-bulk-fillblank').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'blank');
      });
    });
    tbl.querySelectorAll('.es-bulk-clear').forEach(btn => {
      btn.addEventListener('click', e => clearColumnFill(+e.target.dataset.idx));
    });
    tbl.querySelectorAll('.es-bulk-undo').forEach(btn => {
      btn.addEventListener('click', e => { delete columnFills[+e.target.dataset.idx]; reRenderAfterFix(); });
    });
    tbl.querySelectorAll('.es-bulk-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        const v = e.target.value;
        if (v) applyColumnFill(+e.target.dataset.idx, v, 'all');
      });
    });
  }

  // ─── Multi-cell selection ───
  const ckey = (ri, ci) => ri + ':' + ci;

  function selRangeTo(ri, ci) {
    if (!selAnchor) selAnchor = { ri, ci };
    const pa = previewOrder.indexOf(selAnchor.ri);
    const pb = previewOrder.indexOf(ri);
    if (pa < 0 || pb < 0) { selCells = new Set([ckey(ri, ci)]); return; }
    const r0 = Math.min(pa, pb), r1 = Math.max(pa, pb);
    const c0 = Math.min(selAnchor.ci, ci), c1 = Math.max(selAnchor.ci, ci);
    const s = new Set();
    for (let p = r0; p <= r1; p++)
      for (let c = c0; c <= c1; c++) s.add(ckey(previewOrder[p], c));
    selCells = s;
  }

  function updateSelBar() {
    const bar = $('es-sel-bar');
    if (!bar) return;
    const n = selCells.size;
    bar.style.display = n ? 'flex' : 'none';
    const c = $('es-sel-count');
    if (c) c.textContent = n + ' cell' + (n === 1 ? '' : 's') + ' selected';
  }

  function paintSelection() {
    const tbl = $('es-preview-table');
    if (!tbl) return;
    tbl.querySelectorAll('td.es-cell').forEach(td => {
      if (selCells.has(ckey(+td.dataset.ri, +td.dataset.ci))) {
        td.style.outline = '2px solid #2563eb';
        td.style.outlineOffset = '-2px';
      } else {
        td.style.outline = '';
        td.style.outlineOffset = '';
      }
    });
    updateSelBar();
  }

  function clearSelection() {
    selCells = new Set();
    paintSelection();
  }

  function setSelectedCells(val) {
    if (!formattedRows || !selCells.size) return;
    let employerTouched = false;
    const ei = outHeaders.findIndex(h => noStar(h) === 'employer');
    selCells.forEach(k => {
      const i = k.indexOf(':');
      const ri = +k.slice(0, i), ci = +k.slice(i + 1);
      if (!formattedRows[ri]) return;
      formattedRows[ri][ci] = val;
      hasCellEdits = true;
      if (ci === ei) employerTouched = true;
    });
    renderPreview();
    if (employerTouched) renderEmployersToCreate();
    updateSummary();
  }

  // ─── Preview ───
  // "Was … → now … · why" when the cell differs from its source row, for the
  // preview's change tick (see IMWorkbench.changeNote). Hand edits and fills
  // write straight into formattedRows, so they share one reason.
  function changeOf(ri, ci) {
    // Update: measured against what PickTrace has for that employee.
    if (target === 'update' && baseRows) {
      const was = String(baseRows[ri][ci] == null ? '' : baseRows[ri][ci]);
      const now = String(formattedRows[ri][ci] == null ? '' : formattedRows[ri][ci]);
      const fill = columnFills[ci];
      const why = now === CLEAR || (!now && was) ? 'Cleared — exported as ' + CLEAR
        : fill && fill.val === now ? 'Column fill you applied'
        : esSmartFixMap[ci + '||' + was.toUpperCase()] === now ? 'Smart Fix you applied'
        : fromSrc && fromSrc.has(ri + '|' + ci) ? 'From the Data Template'
        : 'Your edit';
      return IMWorkbench.changeNote(was, now === '' && was ? CLEAR : now, why);
    }
    const si = srcColForOut ? srcColForOut[ci] : -1, src = srcKept ? srcKept[ri] : null;
    const was = src && si >= 0 ? String(src[si] == null ? '' : src[si]).trim() : '';
    const now = formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci]);
    let why = 'Changed after import (fill, Smart Fix or your edit)';
    if (esSmartFixMap[ci + '||' + was.toUpperCase()] === now) why = 'Smart Fix you applied';
    else if (DATE_COLS.has(noStar(outHeaders[ci])) && normDate(was) === now) why = 'Date written as YYYY-MM-DD';
    return IMWorkbench.changeNote(was, now, why);
  }

  function renderPreview() {
    if (!formattedRows || !outHeaders) return;
    const sec = $('es-section-preview');
    const tbl = $('es-preview-table');
    sec.style.display = '';
    const possibleSet = new Set(dupPossible.map(d => d.ri));
    const reqIdx = outHeaders.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    const vis = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    // Every row, drawn as you scroll (IMWorkbench.rows) — these lists can run
    // to tens of thousands of employees. A row needs a fix when a required (*)
    // cell is empty or it may duplicate someone; it's changed when any cell
    // differs from the source.
    const info = new Map(), counts = { all: vis.length, fix: 0, chg: 0, any: 0 };
    vis.forEach(ri => {
      const f = reqIdx.filter(i => !formattedRows[ri][i]), c = [];
      for (let i = 0; i < outHeaders.length; i++) if (changeOf(ri, i)) c.push(i);
      const fix = f.length > 0 || possibleSet.has(ri);
      info.set(ri, { f, c, fix });
      if (fix) counts.fix++;
      if (c.length) counts.chg++;
      if (fix || c.length) counts.any++;
    });
    const show = vis.filter(ri => IMWorkbench.rowMatch(pvMode, info.get(ri).fix, info.get(ri).c.length));
    let cols = outHeaders.map((_, i) => i);
    if (pvColsOnly) {
      const keep = new Set(outHeaders.map((h, i) => /^(first|last) name/i.test(h) ? i : -1).filter(i => i >= 0));
      show.forEach(ri => { const x = info.get(ri); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      cols = cols.filter(i => keep.has(i));
    }
    const visSet = new Set(vis);
    selCells.forEach(k => { if (!visSet.has(+k.slice(0, k.indexOf(':')))) selCells.delete(k); });
    const head = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>' + cols.map(i =>
      '<th' + (/\*$/.test(outHeaders[i]) ? ' style="color:#dc2626;"' : '') + '>' + escHtml(outHeaders[i]) + '</th>').join('') + '</tr></thead>';
    const rowHtml = ri => {
      const row = formattedRows[ri];
      const flag = possibleSet.has(ri) ? ' title="Possible duplicate (name match)" style="background:#fffbeb;"' : '';
      let html = '<tr' + flag + '><td style="text-align:center;padding:0;">' +
        '<button class="es-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      cols.forEach(i => {
        const cs = (/\*$/.test(outHeaders[i]) && !row[i]) ? ' style="background:#fee2e2;color:#7f1d1d;"' : '';
        const note = changeOf(ri, i);
        html += '<td class="es-cell' + (note ? ' is-chg' : '') + '" contenteditable="true" spellcheck="false" ' +
          'data-ri="' + ri + '" data-ci="' + i + '"' + cs + (note ? ' title="' + escHtml(note) + '"' : '') + '>' + escHtml(row[i]) + '</td>';
      });
      return html + '</tr>';
    };
    // Shift-click ranges follow the order on screen, sorted or not.
    IMWorkbench.rows(tbl, { head: head, ids: show, row: rowHtml, cols: cols.length + 1,
      after: () => { previewOrder = IMWorkbench.rowOrder(tbl); paintSelection(); },
      keys: [null].concat(cols.map(i => outHeaders[i])), cell: (ri, p) => formattedRows[ri][cols[p - 1]],
      widths: [28].concat(cols.map(i => IMWorkbench.fitWidth(outHeaders[i], show.map(ri => formattedRows[ri][i])))),
      empty: pvMode === 'fix' ? 'No row needs a fix ✓' : 'Nothing here has been changed.' });
    const hint = sec.querySelector('.cmp-sites-hint');
    let ht = IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly, shownCols: cols.length, totalCols: outHeaders.length }) +
      '<span class="wb-filter-note">' + (show.length === vis.length ? 'All ' + vis.length : show.length + ' of ' + vis.length) +
      ' rows · editable — click a cell, type, then click away (or press Enter) to save. Empty required (*) cells are red; <span style="background:#fffbeb;">amber rows</span> are possible duplicates. Click <b>×</b> to drop a row from the export.</span>';
    if (removedRows.size) ht += ' <button class="btn btn-ghost btn-sm" id="es-restore-rows">' +
      'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    if (hint) hint.innerHTML = ht;
    // Rows come and go as you scroll, and the bar is redrawn every render —
    // their handlers live on the table / section, once.
    if (!tbl._esRowsWired) {
      tbl._esRowsWired = true;
      tbl.addEventListener('click', e => {
        const rm = e.target.closest('.es-row-remove');
        if (!rm) return;
        e.stopPropagation();
        removedRows.add(+rm.dataset.ri);
        renderPreview();
        renderPossible();
        renderEmployersToCreate();
        updateSummary();
      });
      sec.addEventListener('change', e => {
        if (e.target.classList.contains('wb-rowfilter')) { pvMode = e.target.value; IMWorkbench.rowsTop(tbl); renderPreview(); }
        else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
      });
      sec.addEventListener('click', e => {
        if (!e.target.closest('#es-restore-rows')) return;
        removedRows.clear();
        renderPreview();
        renderPossible();
        renderEmployersToCreate();
        updateSummary();
      });
    }
    paintSelection();
    updateExportButton();
  }

  // ─── Summary ───
  function updateSummary() {
    if (!tplData || !formattedRows || (!srcData && target !== 'update')) { $('es-summary').style.display = 'none'; return; }
    if (target === 'update') {
      const kept = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
      const changed = kept.filter(rowChanged).length;
      let clears = 0;
      kept.forEach(ri => formattedRows[ri].forEach((_, ci) => { if (updateCell(ri, ci) === CLEAR) clears++; }));
      const sum = $('es-summary');
      sum.style.display = '';
      sum.innerHTML =
        '<div class="cmp-stat"><b>' + kept.length + '</b> employees in the Update template</div>' +
        '<div class="cmp-stat"><b>' + changed + '</b> with a change — these export</div>' +
        (clears ? '<div class="cmp-stat"><b>' + clears + '</b> cell' + (clears === 1 ? '' : 's') + ' cleared (' + escHtml(CLEAR) + ')</div>' : '') +
        (unmatched.length ? '<div class="cmp-stat cmp-warn"><b>' + unmatched.length + '</b> Data Template row' + (unmatched.length === 1 ? '' : 's') + ' not in PickTrace</div>' : '');
      if (wb) wb.refresh();
      return;
    }
    const total = formattedRows.length;
    const visible = formattedRows.filter((_, i) => !removedRows.has(i)).length;
    const excluded = dupConfident.length;
    const excludedDb = dupConfident.filter(d => d.kind === 'db').length;
    const excludedInternal = excluded - excludedDb;
    const possible = dupPossible.filter(d => !removedRows.has(d.ri)).length;
    const reqIdx = outHeaders.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    let emptyReq = 0;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      reqIdx.forEach(i => { if (!r[i]) emptyReq++; });
    });
    const toCreate = getEmployersToCreate().size;
    let limit = parseInt($('es-batch').value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;
    const files = visible ? Math.ceil(visible / limit) : 0;
    const hlOut = srcData.allRows.length - srcData.rows.length;
    const sum = $('es-summary');
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visible + '</b> to create' +
        (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + total + ')</span>' : '') + '</div>' +
      (hlOut ? '<div class="cmp-stat cmp-warn"><b>' + hlOut + '</b> ' +
        (hlState.mode === 'only' ? 'not highlighted — skipped' : 'left out by highlight') + '</div>' : '') +
      '<div class="cmp-stat cmp-warn"><b>' + excluded + '</b> duplicates excluded' +
        (excludedInternal ? ' <span class="text-muted small">(' + excludedDb + ' in DB, ' + excludedInternal + ' within file)</span>' : '') + '</div>' +
      (possible ? '<div class="cmp-stat cmp-warn"><b>' + possible + '</b> possible duplicate' + (possible === 1 ? '' : 's') + '</div>' : '') +
      '<div class="cmp-stat"><b>' + outHeaders.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (toCreate ? '<div class="cmp-stat cmp-warn"><b>' + toCreate + '</b> employer' + (toCreate === 1 ? '' : 's') + ' to create in 3.0</div>' : '') +
      '<div class="cmp-stat"><b>' + files + '</b> output file' + (files === 1 ? '' : 's') + '</div>';
    if (wb) wb.refresh();
  }

  // ─── Workbench: one button per section (workbench.js) ───
  let wb = null;
  function mountWorkbench() {
    if (!window.IMWorkbench) return;
    const vis = () => formattedRows ? formattedRows.filter((_, i) => !removedRows.has(i)) : [];
    wb = IMWorkbench.mount({
      pane: 'cmp-mode-empstd', results: 'es-results', preview: 'es-section-preview',
      ready: () => !!(tplData && formattedRows && (srcData || target === 'update')),
      sections: [
        { key: 'required', group: 'fix', unit: 'column', name: 'Required blanks', desc: 'Fill required (*) columns — or bulk-edit any column',
          ids: ['es-section-bulk'],
          count: () => { const rows = vis(); return outHeaders.filter((h, i) => /\*$/.test(h) && rows.some(r => !r[i])).length; } },
        { key: 'dropdown', group: 'fix', unit: 'value', name: 'Not in dropdown', desc: 'Pick a listed value for anything PickTrace would reject',
          ids: ['es-section-smartfix'], count: () => computeSmartFixes().length },
        { key: 'build', group: 'fix', unit: 'employer', name: 'Build in PickTrace', desc: 'Employers to create in 3.0, then re-download the template',
          ids: ['es-section-create'], count: () => getEmployersToCreate().size },
        { key: 'possible', group: 'fix', unit: 'row', name: 'Possible duplicates', desc: 'Kept, but they match someone by name — check each one',
          ids: ['es-section-possible'], count: () => dupPossible.filter(d => !removedRows.has(d.ri)).length },
        { key: 'dups', group: 'check', unit: 'employee', name: 'Duplicates excluded', desc: 'Already in PickTrace, or earlier in this file — left out',
          ids: ['es-section-dups'], count: () => dupConfident.length },
        { key: 'unmatched', group: 'check', unit: 'row', name: 'Not in PickTrace', desc: 'Update: Data Template rows that match no employee — they need a Bulk Create',
          ids: ['es-section-unmatched'], count: () => (target === 'update' ? unmatched.length : 0) },
        { key: 'highlight', group: 'check', unit: 'row', name: 'Highlighted rows', desc: 'Rows the client marked with a fill colour — leave out or keep only',
          ids: ['es-section-highlight'], count: () => hlScan ? hlScan.rowColors.filter(c => c.length).length : 0 }
      ]
    });
  }

  function updateExportButton() {
    const btn = $('es-export');
    const visible = exportRowList().length;
    btn.disabled = !(tplData && (srcData || target === 'update') && visible); // Existing Employees optional
    if (window.IMDebug) IMDebug.refresh('employee-standardize');
  }

  // ─── Debug dump ───
  function collectDebug() {
    if (!tplData || !formattedRows || !outHeaders || (!srcData && target !== 'update')) return null;
    const H = outHeaders;
    const keptIdxs = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    const keptRows = keptIdxs.map(i => formattedRows[i]);
    const colToSrc = {};
    (srcColForOut || []).forEach((si, i) => { colToSrc[i] = si; });

    const cols = H.map((h, i) => {
      const si = srcColForOut ? srcColForOut[i] : -1;
      return { index: i, header: h, required: /\*$/.test(h), srcIndex: si,
        match: si >= 0 ? 'mapped' : 'unmapped',
        sample: si >= 0 && srcData && srcData.rows[0] ? srcData.rows[0][si] : null };
    });

    const emptyReqCols = H.map((h, i) => {
      if (!/\*$/.test(h)) return null;
      let n = 0; keptIdxs.forEach(ri => { if (!formattedRows[ri][i]) n++; });
      return n ? { column: h, emptyRows: n } : null;
    }).filter(Boolean);
    const toCreate = getEmployersToCreate();
    const sfPending = computeSmartFixes();
    const livePossible = dupPossible.filter(d => !removedRows.has(d.ri));

    const asked = [
      IMDebug.ask('empty-required', 'required', 'Required columns with empty cells', {
        count: emptyReqCols.length, blocksExport: true, items: emptyReqCols }),
      IMDebug.ask('employers-to-create', 'dropdown', 'Employers not in the template dropdown', {
        count: toCreate.size, blocksExport: true,
        detail: 'The upload is rejected with "Employer not found" unless these exist in PickTrace 3.0 first.',
        items: [...toCreate.entries()].map(([n, info]) => ({ employer: n, rows: info.count })) }),
      IMDebug.ask('possible-duplicates', 'data', 'Possible duplicates kept but flagged', {
        count: livePossible.length, blocksExport: false,
        detail: 'Matched on name or a conflicting field — kept in the export, for you to confirm.',
        items: livePossible.map(d => ({ rowIndex: d.ri, name: d.name, altId: d.altId, reason: d.reason, severity: d.sev, existing: d.existing })) }),
      IMDebug.ask('smart-fixes', 'smartfix', 'Off-list values awaiting a Smart Fix choice', {
        count: sfPending.length, blocksExport: false,
        items: sfPending.map(f => ({ column: f.header, value: f.from, rows: f.count, suggestion: f.suggestion || null })) })
    ];

    const upd = target === 'update';
    const prov = IMDebug.deriveProvenance({
      headers: H, rows: formattedRows, srcRows: upd ? baseRows : srcKept, colToSrc: upd ? H.reduce((o, _, i) => { o[i] = i; return o; }, {}) : colToSrc,
      fills: columnFills, smartFixes: esSmartFixMap });
    const sent = upd ? keptIdxs.filter(rowChanged) : keptIdxs;
    const out = IMDebug.output(H, upd ? exportRowList() : keptRows, {
      note: upd ? 'Bulk Update: only employees with a change export; an emptied value is written as ' + CLEAR + '. exportedRowIndexes[n] is the template row behind exported row n.'
        : 'exportedRowIndexes[n] is the formattedRows index behind exported row n. Export is batched — see settings.batchSize.' });
    out.exportedRowIndexes = sent;

    let batch = parseInt($('es-batch') ? $('es-batch').value : '', 10);
    if (!Number.isFinite(batch) || batch < 1) batch = 5000;

    return {
      settings: { target: upd ? 'UPDATE' : 'CREATE', batchSize: batch, outputFiles: sent.length ? Math.ceil(sent.length / batch) : 0,
        update: upd ? { templateEmployees: baseRows ? baseRows.length : 0, changed: sent.length,
          dataTemplateCellsApplied: fromSrc ? fromSrc.size : 0, notInPickTrace: unmatched } : null },
      inputs: [
        IMDebug.file('source data (Data Template Employees sheet)', srcData, { workbookFileName: srcFileName || null }),
        IMDebug.file('personalized bulk Employees ' + (upd ? 'Update' : 'Create') + ' template', tplData, {
          dropdowns: tplData.dropdowns ? [...tplData.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : [],
          employerList: tplData.employerList || null }),
        { role: 'existing employees in PickTrace 3.0', loaded: !!dbIndex,
          fileName: dbFileName || null, rowCount: dbCount || null,
          note: 'Used to exclude confident duplicates and flag possible ones.' }
      ],
      mapping: IMDebug.mapping(cols, srcData ? srcData.headers : [], { srcRows: srcData ? srcData.rows : [] }),
      derived: [
        { kind: 'duplicate classification',
          note: 'Confident duplicates are excluded from the export; possible ones are kept and flagged.',
          excludedConfident: dupConfident.length,
          excludedFoundInDatabase: dupConfident.filter(d => d.kind === 'db').length,
          excludedFoundWithinFile: dupConfident.filter(d => d.kind === 'internal').length,
          flaggedPossible: livePossible.length }
      ],
      fills: Object.keys(columnFills).map(i => ({
        column: H[+i], value: columnFills[i].val, scope: columnFills[i].mode === 'blank' ? 'blank cells only' : 'all rows', kind: 'column fill'
      })).concat(Object.keys(esSmartFixMap).map(k => {
        const sep = k.indexOf('||');
        return { column: H[+k.slice(0, sep)], from: k.slice(sep + 2), value: esSmartFixMap[k], kind: 'smart fix' };
      })),
      edits: { cells: [], names: [],
        note: 'This module edits cells in place in the preview grid rather than through a sticky override map; see provenance for cells that no longer match their source.' },
      dropped: {
        highlightFilter: window.IMHighlight ? IMHighlight.describe(hlScan, hlState) : null,
        removedByHand: { count: removedRows.size, rowIndexes: [...removedRows].slice(0, 500) },
        duplicatesExcluded: {
          count: dupConfident.length,
          items: dupConfident.slice(0, 500).map(d => ({ foundIn: d.kind === 'db' ? 'existing 3.0 export' : 'this file',
            name: d.name, altId: d.altId, matchedOn: d.by, existing: d.existing }))
        }
      },
      asked: asked,
      output: out,
      provenance: prov
    };
  }

  // ─── Export (batched, pristine template per file) ───
  // The rows that go out: every kept row for Create; for Update, only the
  // employees that changed, with <CLEAR> where you emptied a value.
  function exportRowList() {
    if (!formattedRows) return [];
    const idx = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    if (target !== 'update') return idx.map(i => formattedRows[i].slice());
    return idx.filter(rowChanged).map(i => formattedRows[i].map((_, ci) => updateCell(i, ci)));
  }

  async function doExport() {
    if (!formattedRows || !tplData) return;
    const exportRows = exportRowList();
    if (!exportRows.length) { alert(target === 'update' ? 'No employee has a change to send yet.' : 'No employees to export.'); return; }
    const toCreate = getEmployersToCreate();
    if (toCreate.size) {
      const names = [...toCreate.entries()].sort((a, b) => b[1].count - a[1].count)
        .map(([n, info]) => '  • ' + n + ' (' + info.count + ')').join('\n');
      const ok = window.confirm(
        toCreate.size + ' employer' + (toCreate.size === 1 ? '' : 's') +
        ' in this export are not in the template\'s Employer dropdown:\n\n' + names +
        '\n\nCreate these employers in PickTrace 3.0 first, or the upload will be ' +
        'rejected with "Employer not found".\n\nExport anyway?');
      if (!ok) return;
    }
    let limit = parseInt($('es-batch').value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;
    const chunks = chunkArr(exportRows, limit);
    const N = chunks.length;

    for (let i = 0; i < N; i++) {
      // Each file patches the template's own zip (IMXlsxTemplate): only the
      // DATA ENTRY rows change, so its dropdown validations survive — and an
      // Update template's pre-filled employees are replaced by these rows.
      let blob;
      try {
        blob = await IMXlsxTemplate.build({ rawBuffer: tplData.rawBuffer, headers: outHeaders, rows: chunks[i] });
      } catch (err) {
        alert('Could not write the file: ' + (err && err.message ? err.message : err));
        return;
      }
      // "Acme Farms - Employees - 2026-10-07 (1 of 3).xlsx" (see IMOutName in utils.js).
      const name = IMOutName.fileName(target === 'update' ? 'Employees Update' : 'Employees', N === 1 ? '' : ' (' + (i + 1) + ' of ' + N + ')');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      if (i < N - 1) await sleep(300);
    }
  }

  // ─── Reset ───
  function reset() {
    srcData = null; srcWb = null; srcFileName = ''; dbIndex = null; dbFileName = ''; dbCount = 0; tplData = null;
    IMOutName.forget();
    esSmartFixMap = {};
    outHeaders = null; srcColForOut = null;
    formattedRows = null; srcKept = null; baseRows = null; fromSrc = null; unmatched = [];
    dupConfident = []; dupPossible = [];
    setTarget('create');
    { const u = $('es-section-unmatched'); if (u) u.style.display = 'none'; }
    columnFills = {}; removedRows = new Set(); hasCellEdits = false;
    hlScan = null; hlState = { mode: 'off', keys: new Set() };
    selCells = new Set(); selAnchor = null; previewOrder = []; pvMode = 'all'; pvColsOnly = false;
    { const b = $('es-sel-bar'); if (b) b.style.display = 'none'; }
    $('es-src-name').textContent = 'No file selected';
    $('es-db-name').textContent = 'No file selected';
    $('es-tpl-name').textContent = 'No file selected';
    $('es-src-meta').textContent = '';
    $('es-db-meta').textContent = '';
    $('es-tpl-meta').textContent = '';
    $('es-src-file').value = ''; $('es-db-file').value = ''; $('es-tpl-file').value = '';
    ['es-section-highlight', 'es-section-dups', 'es-section-possible', 'es-section-create', 'es-section-smartfix', 'es-section-bulk', 'es-section-preview'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    $('es-summary').style.display = 'none';
    $('es-empty').style.display = '';
    $('es-run').disabled = true;
    $('es-export').disabled = true;
    if (wb) wb.reset();
    if (window.IMDebug) IMDebug.refresh('employee-standardize');
  }

  // ─── Inline cell editing ───
  function commitCellEdit(td) {
    if (!formattedRows) return;
    const ri = +td.dataset.ri, ci = +td.dataset.ci;
    if (!formattedRows[ri]) return;
    const val = td.textContent.trim();
    if (formattedRows[ri][ci] === val) return;
    formattedRows[ri][ci] = val;
    hasCellEdits = true;
    const req = /\*$/.test(outHeaders[ci]);
    td.style.background = (req && !val) ? '#fee2e2' : '';
    td.style.color = (req && !val) ? '#7f1d1d' : '';
    if (noStar(outHeaders[ci]) === 'employer') renderEmployersToCreate();
    updateSummary();
    updateExportButton();
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    initialized = true;
    mountWorkbench();
    if (window.IMDebug) {
      IMDebug.register('employee-standardize', {
        label: 'Employee Standardize',
        ready: () => !!(tplData && formattedRows && outHeaders && (srcData || target === 'update')),
        collect: collectDebug
      });
      IMDebug.wire('es-debug', 'employee-standardize');
    }
    attachModeSwitcher();
    $('es-src-file').addEventListener('change', e => {
      if (e.target.files[0]) handleSrcFile(e.target.files[0]);
      e.target.value = '';
    });
    $('es-db-file').addEventListener('change', e => {
      if (e.target.files[0]) handleDbFile(e.target.files[0]);
      e.target.value = '';
    });
    $('es-tpl-file').addEventListener('change', e => {
      if (e.target.files[0]) handleTplFile(e.target.files[0]);
      e.target.value = '';
    });
    const batch = $('es-batch');
    if (batch) batch.addEventListener('input', () => { if (formattedRows) updateSummary(); });
    // Create / Update. A template of the other kind can't be written as this
    // one — its columns differ — so a mismatch asks for the right template.
    document.querySelectorAll('input[name="es-target"]').forEach(r => r.addEventListener('change', e => {
      if (tplData && tplData.kind && tplData.kind !== e.target.value) {
        alert('The template loaded is a Bulk ' + (tplData.kind === 'update' ? 'Update' : 'Create') + ' template. Load the Bulk ' +
          (e.target.value === 'update' ? 'Update template (Employee ID first, pre-filled with your employees)' : 'Create template') + ' to switch.');
        setTarget(tplData.kind);
        return;
      }
      setTarget(e.target.value);
      formattedRows = null;
      maybeRun();
    }));
    setTarget(target);
    $('es-run').addEventListener('click', runStandardize);
    $('es-export').addEventListener('click', () => { doExport(); });
    $('es-reset').addEventListener('click', reset);
    { const sfa = $('es-smartfix-apply-all'); if (sfa) sfa.addEventListener('click', applyAllSmartFixes); }

    const ptbl = $('es-preview-table');
    if (ptbl) {
      ptbl.addEventListener('focusout', e => {
        const td = e.target.closest ? e.target.closest('td.es-cell') : null;
        if (td) commitCellEdit(td);
      });
      ptbl.addEventListener('keydown', e => {
        if (e.key === 'Escape' && selCells.size) { clearSelection(); return; }
        const td = e.target.closest ? e.target.closest('td.es-cell') : null;
        if (td && e.key === 'Enter') { e.preventDefault(); td.blur(); }
      });
      ptbl.addEventListener('mousedown', e => {
        const td = e.target.closest ? e.target.closest('td.es-cell') : null;
        if (!td) return;
        const ri = +td.dataset.ri, ci = +td.dataset.ci;
        if (e.shiftKey) {
          e.preventDefault();
          selRangeTo(ri, ci);
          paintSelection();
        } else if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          const k = ckey(ri, ci);
          if (selCells.has(k)) selCells.delete(k); else selCells.add(k);
          selAnchor = { ri, ci };
          paintSelection();
        } else {
          selAnchor = { ri, ci };
          if (selCells.size) clearSelection();
        }
      });
    }
    const selApply = $('es-sel-apply');
    if (selApply) {
      selApply.addEventListener('click', () => setSelectedCells($('es-sel-val').value));
      $('es-sel-clear').addEventListener('click', () => setSelectedCells(''));
      $('es-sel-deselect').addEventListener('click', clearSelection);
      $('es-sel-val').addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); setSelectedCells(e.target.value); }
      });
    }
  }

  window.empStdInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
