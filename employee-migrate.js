// ═══════════════════════════════════════════════════════════════════════
// Legacy → 3.0 Employee Migration — converts a Legacy employee CSV export
// into the 3.0 Employee Bulk Create xlsx. Keeps only valid employees
// (Is Valid = true), resolves each employee's Employer* from a separate
// Legacy employer/contractor CSV (Contractor ID → employer Name), maps the
// fixed Legacy→3.0 column schema, surfaces Contractor IDs with no employer
// match in an "Unresolved Employers" picklist, and splits large workforces
// into multiple upload-ready files (configurable max per file). Each output
// file is a pristine copy of the template so every dropdown / data
// validation / style is preserved. Mirrors the Template Standardize module.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // 3.0 Employee Bulk Create — DATA ENTRY columns, exact order. Required
  // columns end with '*' (First Name*, Last Name*, Date of Birth*, Employer*).
  const EMP_HEADERS = [
    'First Name*', 'Middle Name', 'Last Name*', 'Date of Birth*', 'Employer*',
    'Alt ID', 'SSN', 'Gender', 'Hire Date', 'Start Date', 'Language Preference',
    'H2A Employee', 'H2A Contract', 'Employee Group', 'Crew', 'Title',
    'Compensation', 'Phone Number', 'Email', 'Physical Address 1',
    'Physical Address 2', 'Physical Address City', 'Physical Address State',
    'Physical Address Zip Code', 'Physical Address Country', 'Mailing Address 1',
    'Mailing Address 2', 'Mailing Address City', 'Mailing Address State',
    'Mailing Address Zip Code', 'Mailing Address Country',
    'Emergency Contact Name', 'Emergency Contact Relationship',
    'Emergency Phone Number'
  ];

  // ─── State ───
  let empData = null;          // { headers, rows, fileName, idx } — Legacy employee CSV
  let employerMap = null;      // Map<string ContractorID, name> — ACTIVE contractors
  let fullContractorMap = null;// Map<string ContractorID, name> — full (archived+unarchived), optional
  let restoredArchived = new Set(); // archived ContractorIDs the user chose to keep
  let archivedStats = {};      // cid → { count, name } for the Archived panel
  let completedEmployers = new Set(); // employer names ticked off the "to create" checklist
  let migratedAltIds = null;   // Set<altId> of employees already in 3.0 (optional)
  let alreadyMigratedCount = 0;// per-build count of skips due to cross-reference
  let tplData = null;          // { headers, employerList[], rawBuffer, fileName }
  let formattedRows = null;    // [[...34]] mapped rows (post Is-Valid filter)
  let srcKept = null;          // source rows that survived the filter (parallel to formattedRows)
  let employerOverrides = {};  // ContractorID → chosen Employer (sticky across rebuilds)
  let columnFills = {};        // colIdx → { val, mode:'all'|'blank' } (sticky bulk fill)
  let removedRows = new Set(); // formattedRows indexes excluded from preview + export
  let selCells = new Set();    // "ri:ci" keys of multi-selected preview cells
  let selAnchor = null;        // { ri, ci } anchor for shift-range selection
  let previewOrder = [];       // ri's in displayed order (for shift-range math)
  let initialized = false;

  // ─── Mode-pill switcher (manages only this module's pane) ───
  function attachModeSwitcher() {
    // Page visibility is now owned by the top-level pill switcher
    // (template-standardize.js → "employees" tab) and the employee sub-toggle
    // (employee-standardize.js → .emp-subtab). Nothing to wire here.
  }

  // 'MM/DD/YYYY' → 'YYYY-MM-DD' (PickTrace bulk-template date format).
  // Anything that doesn't match is passed through trimmed.
  function normDate(s) {
    if (s instanceof Date && !isNaN(s)) {
      return s.getFullYear() + '-' +
        String(s.getMonth() + 1).padStart(2, '0') + '-' +
        String(s.getDate()).padStart(2, '0');
    }
    const t = String(s == null ? '' : s).trim();
    if (!t) return '';
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); // already ISO
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); // M/D/YY or M/D/YYYY
    if (m) {
      let y = m[3];
      if (y.length === 2) y = (+y >= 70 ? '19' : '20') + y;
      return y + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
    }
    return t;
  }

  // Case-insensitive index of the template's Employer dropdown so a Legacy
  // employer name matches the template literal regardless of casing.
  function buildCaseMap(arr) {
    const m = new Map();
    (arr || []).forEach(v => {
      const k = String(v).toUpperCase().trim();
      if (k && !m.has(k)) m.set(k, v);
    });
    return m;
  }

  function chunkArr(a, n) {
    const out = [];
    for (let i = 0; i < a.length; i += n) out.push(a.slice(i, i + n));
    return out;
  }

  // ─── File parsing ───
  // Minimal RFC-4180 CSV text parser (quoted fields, escaped "" quotes,
  // CRLF/LF). Keeps every value EXACTLY as written — critical for dates like
  // "02/13/2002" which SheetJS would otherwise auto-coerce to a date serial
  // or a re-formatted string before normDate sees it.
  function parseCsvText(text) {
    text = String(text).replace(/^﻿/, ''); // strip BOM
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

  // Reads a CSV as raw text (no date coercion) or an xlsx/xls via XLSX.
  function parseCsvFile(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      const isCsv = /\.csv$/i.test(file.name);
      r.onload = e => {
        try {
          let aoa;
          if (isCsv) {
            aoa = parseCsvText(e.target.result);
          } else {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
          }
          const filtered = aoa.filter(row => row && row.some(c => c != null && String(c).trim() !== ''));
          if (!filtered.length) return reject(new Error('No data rows found.'));
          const headers = filtered[0].map(h => String(h == null ? '' : h).trim());
          if (headers.length) headers[0] = headers[0].replace(/^﻿/, ''); // strip BOM
          resolve({ headers, rows: filtered.slice(1), fileName: file.name });
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      if (isCsv) r.readAsText(file); else r.readAsArrayBuffer(file);
    });
  }

  // Read the whole DROP-DOWN INPUTS sheet into Map<normalizedHeader, values[]>.
  // The sheet's header row literally names each column ("Employer", "Gender",
  // "Title", "Crew", "Employee Group", "Physical Address Country", "Mailing
  // Address Country", "H2A Contract"…), so columns map straight to the 3.0
  // schema by name (same approach as Template Standardize's parseTemplate).
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
        // Keep dropdown values VERBATIM — incl. trailing "ghost spaces".
        // PickTrace stores some re-created employers with a trailing space
        // (e.g. "Johan Hernandez FLC "); trimming it ⇒ "Employer not found".
        const v = String(raw[r][ci] != null ? raw[r][ci] : '');
        if (v.trim() && arr.indexOf(v) < 0) arr.push(v);
      }
      if (arr.length) m.set(k, arr);
    });
    return m;
  }

  // ─── Load handlers ───
  function handleEmpFile(file) {
    parseCsvFile(file).then(d => {
      const idx = {};
      d.headers.forEach((h, i) => { const k = norm(h); if (k && idx[k] == null) idx[k] = i; });
      if (idx['is active'] == null || idx['contractor id'] == null) {
        alert('Employee file must have "Is Active" and "Contractor ID" columns.');
        return;
      }
      empData = { headers: d.headers, rows: d.rows, fileName: file.name, idx };
      $('em-emp-name').textContent = file.name;
      $('em-emp-meta').textContent = d.rows.length + ' rows · ' + d.headers.length + ' columns';
      maybeRun();
    }).catch(err => alert('Failed to read employee file: ' + (err && err.message ? err.message : err)));
  }

  function handleEmployerFile(file) {
    parseCsvFile(file).then(d => {
      const idI = d.headers.findIndex(h => norm(h) === 'id');
      const nameI = d.headers.findIndex(h => norm(h) === 'name');
      if (idI < 0 || nameI < 0) {
        alert('Employer file must have "ID" and "Name" columns.');
        return;
      }
      const m = new Map();
      d.rows.forEach(r => {
        const id = String(r[idI] == null ? '' : r[idI]).trim();
        const nm = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (id) m.set(id, nm);
      });
      employerMap = m;
      $('em-employer-name').textContent = file.name;
      $('em-employer-meta').textContent = m.size + ' active contractor' + (m.size === 1 ? '' : 's');
      maybeRun();
    }).catch(err => alert('Failed to read active contractors file: ' + (err && err.message ? err.message : err)));
  }

  // Optional full (archived + unarchived) contractor list. Contractors in here
  // but NOT in the active list ⇒ archived ⇒ their employees are dropped.
  function handleFullContractorFile(file) {
    parseCsvFile(file).then(d => {
      const idI = d.headers.findIndex(h => norm(h) === 'id');
      const nameI = d.headers.findIndex(h => norm(h) === 'name');
      if (idI < 0 || nameI < 0) {
        alert('Full contractor file must have "ID" and "Name" columns.');
        return;
      }
      const m = new Map();
      d.rows.forEach(r => {
        const id = String(r[idI] == null ? '' : r[idI]).trim();
        const nm = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (id) m.set(id, nm);
      });
      fullContractorMap = m;
      $('em-fullcontractor-name').textContent = file.name;
      $('em-fullcontractor-meta').textContent = m.size + ' contractor' + (m.size === 1 ? '' : 's') + ' (full list)';
      maybeRun();
    }).catch(err => alert('Failed to read full contractor file: ' + (err && err.message ? err.message : err)));
  }

  // Optional: a 3.0 employees export. We extract Alt ID and skip any Legacy
  // employee whose Alt ID is already present (avoids duplicate uploads).
  function handleAlreadyMigratedFile(file) {
    parseCsvFile(file).then(d => {
      const altI = d.headers.findIndex(h => norm(h) === 'alt id');
      if (altI < 0) {
        alert('Already-in-3.0 file must have an "Alt ID" column to cross-reference.');
        return;
      }
      const s = new Set();
      d.rows.forEach(r => {
        const v = String(r[altI] == null ? '' : r[altI]).trim();
        if (v) s.add(v);
      });
      migratedAltIds = s;
      $('em-migrated-name').textContent = file.name;
      $('em-migrated-meta').textContent = s.size + ' Alt ID' + (s.size === 1 ? '' : 's') + ' to skip';
      if (empData && employerMap && tplData) runMigrate();
    }).catch(err => alert('Failed to read 3.0 employees file: ' + (err && err.message ? err.message : err)));
  }

  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      try {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
        const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
        const dataWs = wb.Sheets[dataName];
        if (!dataWs) { alert('Could not find DATA ENTRY sheet in this template.'); return; }
        const dataAoa = XLSX.utils.sheet_to_json(dataWs, { header: 1, defval: '' });
        const headers = dataAoa.length ? dataAoa[0].map(h => String(h == null ? '' : h).trim()) : [];
        // Drop trailing empty header cells — report/template files often carry
        // hundreds of blank styled columns (A1…ZZ1) that aren't real columns.
        while (headers.length && headers[headers.length - 1] === '') headers.pop();
        // Headers carry a trailing '*' on required columns ("First Name*",
        // "Employer*"); norm() lowercases/trims but keeps the '*', so strip it
        // before validating.
        const noStar = h => norm(h || '').replace(/\*$/, '').trim();
        if (noStar(headers[0]) !== 'first name' || noStar(headers[4]) !== 'employer') {
          alert('This does not look like the 3.0 Employee Bulk Create template (expected First Name… and Employer… columns).');
          return;
        }
        const dropdowns = readDropdowns(wb);
        const employerList = dropdowns.get('employer') || [];
        tplData = { headers, dropdowns, employerList, rawBuffer: buf, fileName: file.name };
        const ghost = employerList.filter(v => /\s$/.test(v)).length;
        $('em-tpl-name').textContent = file.name;
        $('em-tpl-meta').textContent = headers.length + ' columns · ' +
          employerList.length + ' employer' + (employerList.length === 1 ? '' : 's') +
          ' (' + ghost + ' with trailing space)';
        maybeRun();
      } catch (err) {
        alert('Failed to read template: ' + (err && err.message ? err.message : err));
      }
    };
    r.onerror = () => alert('Failed to read template file.');
    r.readAsArrayBuffer(file);
  }

  // ─── Index helper for the employee source ───
  function eIdx(key) { return empData && empData.idx[key] != null ? empData.idx[key] : -1; }
  function srcCid(srcRow) {
    const i = eIdx('contractor id');
    return (i < 0 ? '' : String(srcRow[i] == null ? '' : srcRow[i]).trim()) || '(blank)';
  }

  // ─── Transformation ───
  // Filter (Is Valid = true) + fixed Legacy→3.0 column map. Builds
  // formattedRows and the parallel srcKept array IN THE SAME PASS — the
  // unresolved-employers panel relies on formattedRows[i] ↔ srcKept[i].
  function buildFormattedRowsRaw() {
    if (!empData) return null;
    const get = (row, key) => { const i = eIdx(key); return i < 0 ? '' : String(row[i] == null ? '' : row[i]).trim(); };
    const iaI = eIdx('is active');
    const altI = eIdx('alt id');
    srcKept = [];
    archivedStats = {};
    alreadyMigratedCount = 0;
    const out = [];
    empData.rows.forEach(srcRow => {
      // Active filter: Is Active column is the sole determinant. Skip if not
      // literally "true" (false / blank both excluded).
      const ia = iaI < 0 ? '' : String(srcRow[iaI] == null ? '' : srcRow[iaI]).trim().toLowerCase();
      if (ia !== 'true') return;
      // Cross-reference: skip employees already migrated to 3.0 (joined on
      // Alt ID). Blank Alt IDs can't be cross-referenced and pass through.
      if (migratedAltIds && altI >= 0) {
        const aid = String(srcRow[altI] == null ? '' : srcRow[altI]).trim();
        if (aid && migratedAltIds.has(aid)) { alreadyMigratedCount++; return; }
      }

      // Classify the contractor: active (in active list) → resolve name;
      // archived (in full list but not active) → drop unless restored;
      // unknown (in neither) → keep with blank Employer (Unresolved panel).
      const cid = get(srcRow, 'contractor id');
      const activeName = employerMap ? employerMap.get(cid) : undefined;
      let employer = '';
      if (activeName !== undefined) {
        employer = activeName || '';
      } else if (fullContractorMap && fullContractorMap.has(cid)) {
        const nm = fullContractorMap.get(cid) || '';
        const st = archivedStats[cid] || { count: 0, name: nm };
        st.count++; archivedStats[cid] = st;
        if (!restoredArchived.has(cid)) return; // archived → removed
        employer = nm;                          // archived but restored
      }

      const r = new Array(EMP_HEADERS.length).fill('');
      r[0]  = get(srcRow, 'first name');
      r[1]  = get(srcRow, 'middle name');
      r[2]  = get(srcRow, 'last name');
      r[3]  = normDate(get(srcRow, 'birth date'));
      r[4]  = employer; // Employer*
      r[5]  = get(srcRow, 'alt id');
      r[6]  = get(srcRow, 'ssn');
      const male = get(srcRow, 'is male').toLowerCase();
      r[7]  = male === 'true' ? 'MALE' : (male === 'false' ? 'FEMALE' : '');
      r[8]  = normDate(get(srcRow, 'hired date')); // Hire Date ← Hired Date
      // r[9] Start Date — intentionally blank (see plan: open items)
      r[17] = get(srcRow, 'mobile phone #');
      r[18] = get(srcRow, 'email');
      r[19] = get(srcRow, 'address 1');
      r[20] = get(srcRow, 'address 2');
      r[21] = get(srcRow, 'city');
      r[22] = get(srcRow, 'state');
      r[23] = get(srcRow, 'postal code');
      r[24] = get(srcRow, 'country');
      out.push(r);
      srcKept.push(srcRow);
    });

    // Canonicalize Employer* to the template's dropdown literal when it
    // matches (case-insensitive). A resolved name that ISN'T in the template's
    // Employer dropdown is KEPT in the column (per request) — it surfaces in
    // both the "Employers to Create in 3.0" list and the Unresolved Employers
    // panel so it can be created in PickTrace or remapped before upload.
    if (tplData && tplData.employerList && tplData.employerList.length) {
      const cm = buildCaseMap(tplData.employerList);
      out.forEach(r => {
        if (r[4]) r[4] = cm.get(String(r[4]).toUpperCase().trim()) || r[4];
      });
    }

    // Snap every other dropdown-backed column to the template's allowed
    // values (case-insensitive; country aliases like USA→US). Unmatched
    // values are kept as-is so they stay visible and bulk-fixable.
    for (let ci = 0; ci < EMP_HEADERS.length; ci++) {
      if (ci === 4) continue; // Employer handled above
      const opts = colOptions(EMP_HEADERS[ci]);
      if (!opts || !opts.length) continue;
      const isCountry = /country$/i.test(EMP_HEADERS[ci]);
      const cm = buildCaseMap(opts);
      out.forEach(r => {
        if (r[ci]) r[ci] = snapToOption(r[ci], cm, isCountry);
      });
    }
    return out;
  }

  const COUNTRY_ALIASES = {
    'usa': 'US', 'u.s.': 'US', 'u.s.a.': 'US', 'united states': 'US',
    'united states of america': 'US', 'america': 'US',
    'mexico': 'MX', 'méxico': 'MX', 'canada': 'CA'
  };
  // cm = buildCaseMap(options): Map<UPPER, canonicalOption>. Returns the
  // template's literal when the value matches (case-insensitively, or via a
  // country alias); otherwise the original value untouched.
  function snapToOption(val, cm, isCountry) {
    const v = String(val == null ? '' : val).trim();
    if (!v || !cm.size) return v;
    const hit = cm.get(v.toUpperCase());
    if (hit != null) return hit;
    if (isCountry) {
      const a = COUNTRY_ALIASES[v.toLowerCase()];
      if (a) { const h2 = cm.get(String(a).toUpperCase()); if (h2 != null) return h2; }
    }
    return v;
  }

  function rebuildFormattedRows() {
    formattedRows = buildFormattedRowsRaw();
    if (!formattedRows) return;
    // Re-apply sticky per-Contractor-ID employer overrides (trusted as-is,
    // like Template Standardize's manualFills).
    formattedRows.forEach((r, i) => {
      const cid = srcCid(srcKept[i]);
      if (employerOverrides[cid]) r[4] = employerOverrides[cid];
    });
    // Re-apply sticky column bulk fills. mode 'all' overwrites the whole
    // column; mode 'blank' fills only empty cells (per non-removed row).
    Object.keys(columnFills).forEach(k => {
      const ci = +k, { val, mode } = columnFills[k];
      formattedRows.forEach((r, i) => {
        if (removedRows.has(i)) return;
        if (mode === 'blank') { if (!r[ci]) r[ci] = val; }
        else r[ci] = val;
      });
    });
  }

  function runMigrate() {
    if (!empData || !employerMap || !tplData) {
      alert('Upload the employee CSV, the employer CSV, and the 3.0 template first.');
      return;
    }
    employerOverrides = {};
    columnFills = {};
    removedRows = new Set();
    restoredArchived = new Set();
    completedEmployers = new Set();
    selCells = new Set(); selAnchor = null;
    rebuildFormattedRows();
    renderArchived();
    renderUnresolved();
    renderBulkEdit();
    renderPreview();
    updateSummary();
    $('em-empty').style.display = 'none';
  }

  function maybeRun() {
    const ready = !!(empData && employerMap && tplData);
    $('em-run').disabled = !ready;
    if (ready) runMigrate();
  }

  // ─── Archived contractors (removed, restorable) ───
  function archivedRemovedCount() {
    let n = 0;
    Object.keys(archivedStats).forEach(cid => {
      if (!restoredArchived.has(cid)) n += archivedStats[cid].count;
    });
    return n;
  }

  function renderArchived() {
    const sec = $('em-section-archived');
    const tbl = $('em-archived-table');
    if (!sec || !tbl) return;
    const cids = Object.keys(archivedStats);
    if (!cids.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const removed = archivedRemovedCount();
    const title = $('em-archived-title');
    if (title) title.textContent = 'Archived Contractors — ' + removed +
      ' employee' + (removed === 1 ? '' : 's') + ' removed';
    let html = '<thead><tr><th>Contractor ID</th><th>Contractor</th>' +
      '<th>Employees</th><th>Status</th><th></th></tr></thead><tbody>';
    cids.sort((a, b) => archivedStats[b].count - archivedStats[a].count).forEach(cid => {
      const st = archivedStats[cid];
      const on = restoredArchived.has(cid);
      const status = on
        ? '<span style="color:#15803d;font-weight:600;">✓ Restored (migrating)</span>'
        : '<span style="color:#b45309;font-weight:600;">⦸ Removed</span>';
      html += '<tr>' +
        '<td><b>' + escHtml(cid) + '</b></td>' +
        '<td>' + escHtml(st.name || '(no name)') + '</td>' +
        '<td>' + st.count + '</td>' +
        '<td>' + status + '</td>' +
        '<td><button class="btn btn-sm ' + (on ? 'btn-ghost' : 'btn-success') +
          ' em-arch-toggle" data-cid="' + escHtml(cid) + '">' +
          (on ? 'Remove' : 'Restore') + '</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.em-arch-toggle').forEach(btn => {
      btn.addEventListener('click', e => {
        const cid = e.target.dataset.cid;
        if (restoredArchived.has(cid)) restoredArchived.delete(cid);
        else restoredArchived.add(cid);
        rebuildFormattedRows();
        renderArchived();
        renderUnresolved();
        renderBulkEdit();
        renderPreview();
        updateSummary();
      });
    });
  }

  // ─── Unresolved employers ───
  // True if v is a valid 3.0 employer (in the template's Employer dropdown).
  // When the template has no Employer dropdown we can't validate → treat any
  // non-blank value as acceptable.
  function employerValid(v) {
    if (!v) return false;
    if (!tplData || !tplData.employerList || !tplData.employerList.length) return true;
    return buildCaseMap(tplData.employerList).has(String(v).toUpperCase().trim());
  }

  // Distinct employer names sitting in the column that are NOT valid 3.0
  // employers — i.e. they must be created in PickTrace before upload.
  function getEmployersToCreate() {
    const m = new Map(); // employerName → { count, cids:Set }
    if (!formattedRows) return m;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const v = r[4];
      if (!v || employerValid(v)) return;
      const e = m.get(v) || { count: 0, cids: new Set() };
      e.count++;
      e.cids.add(srcCid(srcKept[ri]));
      m.set(v, e);
    });
    return m;
  }

  function getUnresolved() {
    const m = new Map(); // cid → { count, sampleName, raw }
    if (!formattedRows || !empData) return m;
    const fnI = eIdx('first name'), lnI = eIdx('last name');
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (employerValid(r[4])) return; // valid 3.0 employer — nothing to resolve
      const src = srcKept[ri];
      const cid = srcCid(src);
      const e = m.get(cid) ||
        { count: 0, sampleName: '', raw: (employerMap && employerMap.get(cid)) ||
            (fullContractorMap && fullContractorMap.get(cid)) || '' };
      e.count++;
      if (!e.sampleName) {
        const fn = fnI < 0 ? '' : String(src[fnI] == null ? '' : src[fnI]).trim();
        const ln = lnI < 0 ? '' : String(src[lnI] == null ? '' : src[lnI]).trim();
        e.sampleName = (fn + ' ' + ln).trim();
      }
      m.set(cid, e);
    });
    return m;
  }

  // Employers still needing creation (checklist items not yet ticked off).
  function pendingEmployersToCreate() {
    const out = [];
    getEmployersToCreate().forEach((info, name) => {
      if (!completedEmployers.has(name)) out.push([name, info]);
    });
    return out;
  }

  function renderEmployersToCreate() {
    const sec = $('em-section-create');
    const tbl = $('em-create-table');
    if (!sec || !tbl) return;
    const m = getEmployersToCreate();
    if (!m.size) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const done = [...m.keys()].filter(nm => completedEmployers.has(nm)).length;
    const title = $('em-create-title');
    if (title) title.textContent = 'Employers to Create in 3.0 (' +
      (m.size - done) + ' to create' + (done ? ' · ' + done + ' done' : '') + ')';
    let html = '<thead><tr><th></th><th>Employer to create</th><th>Employees</th>' +
      '<th>Legacy Contractor IDs</th></tr></thead><tbody>';
    [...m.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([name, info]) => {
      const isDone = completedEmployers.has(name);
      const mark = isDone
        ? '<span style="color:#15803d;font-weight:700;">&#10003;</span>'
        : '<span style="color:#9ca3af;">&#9744;</span>';
      const nameStyle = isDone ? ' style="text-decoration:line-through;color:#15803d;"' : '';
      html += '<tr class="em-create-row" data-emp="' + escHtml(name) + '" ' +
        'style="cursor:pointer;"' + (isDone ? ' ' : '') +
        ' title="Click to copy this name and mark it created in 3.0">' +
        '<td style="text-align:center;">' + mark + '</td>' +
        '<td><b' + nameStyle + '>' + escHtml(name) + '</b></td>' +
        '<td>' + info.count + '</td>' +
        '<td>' + escHtml([...info.cids].sort((x, y) => x - y).join(', ')) + '</td></tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.em-create-row').forEach(tr => {
      tr.addEventListener('click', () => {
        const name = tr.dataset.emp;
        if (completedEmployers.has(name)) completedEmployers.delete(name);
        else completedEmployers.add(name);
        if (navigator.clipboard) navigator.clipboard.writeText(name).catch(() => {});
        renderEmployersToCreate();
        updateSummary();
      });
    });
  }

  function renderUnresolved() {
    renderEmployersToCreate();
    if (!formattedRows || !tplData) return;
    const sec = $('em-section-unresolved');
    const tbl = $('em-unresolved-table');
    const m = getUnresolved();
    if (!m.size) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const title = $('em-unresolved-title');
    if (title) title.textContent = 'Unresolved Employers (' + m.size + ')';
    const list = tplData.employerList || [];
    let html = '<thead><tr><th>Contractor ID</th><th>Affected rows</th>' +
      '<th>Sample employee</th><th>Why</th><th>Pick Employer</th>' +
      '<th>Manual override</th><th></th></tr></thead><tbody>';
    [...m.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([cid, info]) => {
      const why = info.raw
        ? 'Mapped to “' + escHtml(info.raw) + '” — not in template dropdown'
        : 'No employer record for this Contractor ID';
      const sel = employerOverrides[cid] || '';
      const opts = list.length
        ? '<option value="">— pick —</option>' +
          list.map(o => '<option value="' + escHtml(o) + '"' + (o === sel ? ' selected' : '') +
            '>' + escHtml(o) + (/\s$/.test(o) ? ' ␣(trailing space)' : '') + '</option>').join('')
        : '<option value="">(template has no employer dropdown)</option>';
      html += '<tr>' +
        '<td><b>' + escHtml(cid) + '</b></td>' +
        '<td><span style="color:#dc2626;font-weight:600;">⚠ ' + info.count + '</span></td>' +
        '<td>' + escHtml(info.sampleName || '—') + '</td>' +
        '<td>' + why + '</td>' +
        '<td><select class="em-emp-pick input-field" data-cid="' + escHtml(cid) + '" style="min-width:200px;">' + opts + '</select></td>' +
        '<td><input type="text" class="em-emp-override input-field" data-cid="' + escHtml(cid) + '" placeholder="Manual override" style="width:200px;"></td>' +
        '<td><button class="btn btn-primary btn-sm em-emp-apply" data-cid="' + escHtml(cid) + '">Apply</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    wireUnresolvedHandlers();
  }

  function applyEmployer(cid, val) {
    if (!val) return;
    employerOverrides[cid] = val; // sticky — survives rebuilds
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (srcCid(srcKept[ri]) === cid) r[4] = val;
    });
    renderPreview();
    renderUnresolved();
    updateSummary();
  }

  function wireUnresolvedHandlers() {
    const tbl = $('em-unresolved-table');
    if (!tbl) return;
    tbl.querySelectorAll('.em-emp-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const tr = e.target.closest('tr');
        const cid = e.target.dataset.cid;
        // Don't trim the dropdown pick — it may carry a deliberate trailing
        // "ghost space" that must be preserved to match the 3.0 employer.
        // The manual override likewise keeps a trailing space (leading only
        // stripped) so a ghost-spaced name can be typed in by hand.
        const pick = tr.querySelector('.em-emp-pick').value;
        const ovr = tr.querySelector('.em-emp-override').value.replace(/^\s+/, '');
        const val = ovr || pick;
        if (!val) { alert('Pick an Employer or type a manual override first.'); return; }
        applyEmployer(cid, val);
      });
    });
    // Instant apply on dropdown pick (intentional, like Template Standardize).
    tbl.querySelectorAll('.em-emp-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        const v = e.target.value; // verbatim — keep any trailing ghost space
        if (v) applyEmployer(e.target.dataset.cid, v);
      });
    });
  }

  // ─── Bulk-edit columns ───
  // Allowed values for a 3.0 column: first the template's DROP-DOWN INPUTS
  // sheet (matched by column name), then inline-list fallbacks for the few
  // validations the 3.0 template defines inline (Gender / Language / H2A).
  function colOptions(header) {
    const h = norm(header).replace(/\*$/, '').trim();
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

  function renderBulkEdit() {
    if (!formattedRows || !tplData) return;
    const sec = $('em-section-bulk');
    const tbl = $('em-bulk-table');
    sec.style.display = '';
    let html = '<thead><tr><th>Column</th><th>Current bulk value</th>' +
      '<th>Pick</th><th>Manual override</th><th></th><th></th></tr></thead><tbody>';
    EMP_HEADERS.forEach((h, idx) => {
      const opts = colOptions(h);
      const cur = columnFills[idx];
      const optsHtml = opts.length
        ? '<option value="">— pick —</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (cur && o === cur.val ? ' selected' : '') +
            '>' + escHtml(o) + (/\s$/.test(o) ? ' ␣(trailing space)' : '') + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const status = cur
        ? '<span style="color:#15803d;font-weight:600;">✓ ' + escHtml(String(cur.val).slice(0, 40) || '(blank)') +
          '</span> <span class="text-muted small">(' + (cur.mode === 'blank' ? 'blanks only' : 'all rows') + ')</span>'
        : '<span class="text-muted small">—</span>';
      const req = /\*$/.test(h);
      html += '<tr>' +
        '<td><b' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="em-bulk-pick input-field" data-idx="' + idx + '" style="min-width:180px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="em-bulk-override input-field" data-idx="' + idx + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm em-bulk-apply" data-idx="' + idx + '" title="Overwrite this column for every row">Apply all</button> ' +
        '<button class="btn btn-success btn-sm em-bulk-fillblank" data-idx="' + idx + '" title="Fill only the empty cells in this column">Fill blanks</button></td>' +
        '<td><button class="btn btn-ghost btn-sm em-bulk-clear" data-idx="' + idx + '" title="Empty this column for all rows">Clear</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    wireBulkEditHandlers();
  }

  // mode 'all' overwrites every non-removed row; mode 'blank' fills only the
  // empty cells (e.g. "give every missing Employer this value").
  function applyColumnFill(idx, val, mode) {
    columnFills[idx] = { val, mode: mode === 'blank' ? 'blank' : 'all' }; // sticky
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (mode === 'blank') { if (!r[idx]) r[idx] = val; }
      else r[idx] = val;
    });
    renderBulkEdit();
    if (idx === 4) renderUnresolved(); // Employer column affects unresolved panel
    renderPreview();
    updateSummary();
  }

  function clearColumnFill(idx) {
    delete columnFills[idx];
    formattedRows.forEach((r, ri) => { if (!removedRows.has(ri)) r[idx] = ''; });
    renderBulkEdit();
    if (idx === 4) renderUnresolved();
    renderPreview();
    updateSummary();
  }

  function wireBulkEditHandlers() {
    const tbl = $('em-bulk-table');
    if (!tbl) return;
    const grab = e => {
      const tr = e.target.closest('tr');
      // Pick verbatim (keeps trailing ghost spaces); override keeps a trailing
      // space too (leading whitespace only is stripped).
      const pick = tr.querySelector('.em-bulk-pick').value;
      const ovr = tr.querySelector('.em-bulk-override').value.replace(/^\s+/, '');
      return ovr || pick;
    };
    tbl.querySelectorAll('.em-bulk-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'all');
      });
    });
    tbl.querySelectorAll('.em-bulk-fillblank').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'blank');
      });
    });
    tbl.querySelectorAll('.em-bulk-clear').forEach(btn => {
      btn.addEventListener('click', e => clearColumnFill(+e.target.dataset.idx));
    });
    tbl.querySelectorAll('.em-bulk-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        const v = e.target.value; // verbatim — keep any trailing ghost space
        if (v) applyColumnFill(+e.target.dataset.idx, v, 'all');
      });
    });
  }

  // ─── Multi-cell selection (Shift = range, Ctrl/Cmd = toggle) ───
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
    const bar = $('em-sel-bar');
    if (!bar) return;
    const n = selCells.size;
    bar.style.display = n ? 'flex' : 'none';
    const c = $('em-sel-count');
    if (c) c.textContent = n + ' cell' + (n === 1 ? '' : 's') + ' selected';
  }

  function paintSelection() {
    const tbl = $('em-preview-table');
    if (!tbl) return;
    tbl.querySelectorAll('td.em-cell').forEach(td => {
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

  // Write one value into every selected cell at once (Apply / Clear contents).
  function setSelectedCells(val) {
    if (!formattedRows || !selCells.size) return;
    let employerTouched = false;
    selCells.forEach(k => {
      const i = k.indexOf(':');
      const ri = +k.slice(0, i), ci = +k.slice(i + 1);
      if (!formattedRows[ri]) return;
      formattedRows[ri][ci] = val;
      if (ci === 4) employerTouched = true;
    });
    renderPreview();           // re-renders + re-paints the kept selection
    if (employerTouched) renderUnresolved();
    updateSummary();
  }

  // ─── Preview ───
  function renderPreview() {
    if (!empData || !tplData || !formattedRows) return;
    const sec = $('em-section-preview');
    const tbl = $('em-preview-table');
    sec.style.display = '';
    let html = '<thead><tr>' +
      '<th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    EMP_HEADERS.forEach(h => {
      const req = /\*$/.test(h);
      html += '<th' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</th>';
    });
    html += '</tr></thead><tbody>';
    const reqIdx = EMP_HEADERS.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    const isProblem = ri => reqIdx.some(i => !formattedRows[ri][i]);
    const vis = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    // ALWAYS show every row missing a required cell so nothing problematic is
    // hidden from the export. Complete rows are capped for performance.
    const CLEAN_CAP = 50;
    let cleanShown = 0;
    const show = [];
    vis.forEach(ri => {
      if (isProblem(ri)) show.push(ri);
      else if (cleanShown < CLEAN_CAP) { show.push(ri); cleanShown++; }
    });
    const problemCount = vis.reduce((n, ri) => n + (isProblem(ri) ? 1 : 0), 0);
    const cleanTotal = vis.length - problemCount;
    // Track displayed order for shift-range math; drop selections whose row
    // is no longer visible (e.g. removed) so the toolbar count stays honest.
    previewOrder = show.slice();
    const visSet = new Set(vis);
    selCells.forEach(k => { if (!visSet.has(+k.slice(0, k.indexOf(':')))) selCells.delete(k); });
    show.forEach(ri => {
      const row = formattedRows[ri];
      html += '<tr><td style="text-align:center;padding:0;">' +
        '<button class="em-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      EMP_HEADERS.forEach((h, i) => {
        const empty = !row[i];
        const cs = (/\*$/.test(h) && empty) ? ' style="background:#fee2e2;color:#7f1d1d;"' : '';
        html += '<td class="em-cell" contenteditable="true" spellcheck="false" ' +
          'data-ri="' + ri + '" data-ci="' + i + '"' + cs + '>' + escHtml(row[i]) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    const hint = sec.querySelector('.cmp-sites-hint');
    let ht = 'Editable — click a cell, type, then click away (or press Enter) to save. Empty required (*) cells are red. Click <b>×</b> to drop a row from the export.';
    if (problemCount) {
      ht = '<b style="color:#dc2626;">' + problemCount + ' row' + (problemCount === 1 ? '' : 's') +
        ' missing a required (*) cell</b> — all shown below' +
        (cleanTotal > cleanShown ? ', plus first ' + cleanShown + ' of ' + cleanTotal + ' complete rows' : '') +
        '. ' + ht;
    } else if (cleanTotal > cleanShown) {
      ht = 'All required cells filled. Showing first ' + cleanShown + ' of ' + cleanTotal + ' rows. ' + ht;
    }
    if (removedRows.size) ht += ' &nbsp; <button class="btn btn-ghost btn-sm" id="em-restore-rows">' +
      'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    if (hint) hint.innerHTML = ht;
    tbl.querySelectorAll('.em-row-remove').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        removedRows.add(+e.currentTarget.dataset.ri);
        renderPreview();
        renderUnresolved();
        updateSummary();
      });
    });
    const rb = $('em-restore-rows');
    if (rb) rb.addEventListener('click', () => {
      removedRows.clear();
      renderPreview();
      renderUnresolved();
      updateSummary();
    });
    paintSelection();
    updateExportButton();
  }

  // ─── Summary ───
  function updateSummary() {
    if (!empData || !tplData) { $('em-summary').style.display = 'none'; return; }
    const total = formattedRows ? formattedRows.length : 0;
    const visible = formattedRows ? formattedRows.filter((_, i) => !removedRows.has(i)).length : 0;
    const archived = archivedRemovedCount();
    const skipped = empData.rows.length - total - archived - alreadyMigratedCount; // Is Active filter only
    const reqIdx = EMP_HEADERS.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    let emptyReq = 0;
    if (formattedRows) formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      reqIdx.forEach(i => { if (!r[i]) emptyReq++; });
    });
    const unresolved = getUnresolved().size;
    const toCreate = pendingEmployersToCreate().length;
    let limit = parseInt($('em-batch').value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;
    const files = visible ? Math.ceil(visible / limit) : 0;
    const sum = $('em-summary');
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visible + '</b> employees' +
        (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + total + ')</span>' : '') + '</div>' +
      '<div class="cmp-stat"><b>' + skipped + '</b> skipped (Is Active &ne; true)</div>' +
      (alreadyMigratedCount ? '<div class="cmp-stat"><b>' + alreadyMigratedCount + '</b> skipped (already in 3.0)</div>' : '') +
      (archived ? '<div class="cmp-stat cmp-warn"><b>' + archived + '</b> removed (archived contractor)</div>' : '') +
      '<div class="cmp-stat"><b>' + EMP_HEADERS.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (unresolved ? '<div class="cmp-stat cmp-warn"><b>' + unresolved + '</b> unresolved employer' + (unresolved === 1 ? '' : 's') + '</div>' : '') +
      (toCreate ? '<div class="cmp-stat cmp-warn"><b>' + toCreate + '</b> employer' + (toCreate === 1 ? '' : 's') + ' to create in 3.0</div>' : '') +
      '<div class="cmp-stat"><b>' + files + '</b> output file' + (files === 1 ? '' : 's') + '</div>';
  }

  function updateExportButton() {
    const btn = $('em-export');
    const visible = formattedRows ? formattedRows.filter((_, i) => !removedRows.has(i)).length : 0;
    btn.disabled = !(empData && employerMap && tplData && visible);
  }

  // ─── Export (batched, pristine template per file) ───
  async function doExport() {
    if (!formattedRows || !tplData) return;
    const exportRows = formattedRows.filter((_, i) => !removedRows.has(i)).map(r => r.slice());
    if (!exportRows.length) { alert('No employees to export.'); return; }
    // Warn if employers still need to be created in PickTrace 3.0 — uploading
    // before they exist will be rejected with "Employer not found".
    const pending = pendingEmployersToCreate();
    if (pending.length) {
      const names = pending.sort((a, b) => b[1].count - a[1].count)
        .map(([n, info]) => '  • ' + n + ' (' + info.count + ')').join('\n');
      const ok = window.confirm(
        pending.length + ' employer' + (pending.length === 1 ? '' : 's') +
        ' in this export are not yet created in PickTrace 3.0:\n\n' + names +
        '\n\nCreate these employers in PickTrace first (click each in the ' +
        '"Employers to Create" list to copy + check it off), or the upload ' +
        'will be rejected with "Employer not found".\n\nExport anyway?');
      if (!ok) return;
    }
    let limit = parseInt($('em-batch').value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;
    const chunks = chunkArr(exportRows, limit);
    const N = chunks.length;
    const origName = (tplData.fileName || '3.0_Employee_bulk_create.xlsx').trim();
    const dotIdx = origName.lastIndexOf('.');
    const base = dotIdx > 0 ? origName.substring(0, dotIdx) : origName;
    const ext = dotIdx > 0 ? origName.substring(dotIdx) : '.xlsx';

    for (let i = 0; i < N; i++) {
      const rows = chunks[i];
      // Re-read the ORIGINAL template buffer for every file so each workbook
      // keeps shared strings, the DROP-DOWN INPUTS sheet, and all data
      // validations intact (same recipe as Template Standardize's doExport).
      const wb = XLSX.read(new Uint8Array(tplData.rawBuffer), {
        type: 'array', cellStyles: true, cellDates: true, sheetStubs: true
      });
      const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
      const ws = wb.Sheets[dataName];
      if (!ws) { alert('Template missing DATA ENTRY sheet — cannot export.'); return; }

      // Wipe existing data rows (keep header row 0); leave sheet-level
      // properties (!cols, !merges, !dataValidation, …) intact.
      const range = ws['!ref']
        ? XLSX.utils.decode_range(ws['!ref'])
        : { s: { r: 0, c: 0 }, e: { r: 0, c: EMP_HEADERS.length - 1 } };
      for (let r = 1; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const ref = XLSX.utils.encode_cell({ r, c });
          if (ws[ref]) delete ws[ref];
        }
      }

      // Write chunk rows starting at row 1. All values as text — employee
      // data is names / ISO dates / SSN / phone / Alt ID and must match the
      // template's list validations literally and keep any leading zeros.
      const colCount = Math.max(range.e.c + 1, EMP_HEADERS.length);
      rows.forEach((row, ri) => {
        for (let c = 0; c < colCount; c++) {
          const val = row[c];
          if (val == null || val === '') continue;
          ws[XLSX.utils.encode_cell({ r: ri + 1, c })] = { v: String(val), t: 's' };
        }
      });

      ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: Math.max(0, rows.length), c: colCount - 1 }
      });

      const name = N === 1
        ? base + ' — filled' + ext
        : base + ' — filled (' + (i + 1) + ' of ' + N + ')' + ext;
      // bookSST:true forces shared-string output — that is the only writer
      // path where SheetJS emits xml:space="preserve", which is required to
      // keep trailing "ghost spaces" on employer names (otherwise the reader
      // trims them and PickTrace reports "Employer not found").
      XLSX.writeFile(wb, name, { cellStyles: true, bookSST: true });
      if (i < N - 1) await sleep(300); // let each download settle
    }
  }

  // ─── Reset ───
  function reset() {
    empData = null; employerMap = null; fullContractorMap = null;
    migratedAltIds = null; alreadyMigratedCount = 0; tplData = null;
    formattedRows = null; srcKept = null; employerOverrides = {};
    columnFills = {};
    removedRows = new Set();
    restoredArchived = new Set(); archivedStats = {};
    completedEmployers = new Set();
    selCells = new Set(); selAnchor = null; previewOrder = [];
    { const b = $('em-sel-bar'); if (b) b.style.display = 'none'; }
    $('em-emp-name').textContent = 'No file selected';
    $('em-employer-name').textContent = 'No file selected';
    $('em-fullcontractor-name').textContent = 'No file selected';
    $('em-migrated-name').textContent = 'No file selected';
    $('em-tpl-name').textContent = 'No file selected';
    $('em-emp-meta').textContent = '';
    $('em-employer-meta').textContent = '';
    $('em-fullcontractor-meta').textContent = '';
    $('em-migrated-meta').textContent = '';
    $('em-tpl-meta').textContent = '';
    $('em-emp-file').value = ''; $('em-employer-file').value = '';
    $('em-fullcontractor-file').value = ''; $('em-migrated-file').value = '';
    $('em-tpl-file').value = '';
    ['em-section-archived', 'em-section-create', 'em-section-unresolved', 'em-section-bulk', 'em-section-preview'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    $('em-summary').style.display = 'none';
    $('em-empty').style.display = '';
    $('em-run').disabled = true;
    $('em-export').disabled = true;
  }

  // ─── Inline cell editing ───
  // Commits a contenteditable preview cell back into formattedRows (the
  // canonical state, so the edit persists through re-renders and export).
  function commitCellEdit(td) {
    if (!formattedRows) return;
    const ri = +td.dataset.ri, ci = +td.dataset.ci;
    if (!formattedRows[ri]) return;
    const val = td.textContent.trim();
    if (formattedRows[ri][ci] === val) return;
    formattedRows[ri][ci] = val;
    const req = /\*$/.test(EMP_HEADERS[ci]);
    td.style.background = (req && !val) ? '#fee2e2' : '';
    td.style.color = (req && !val) ? '#7f1d1d' : '';
    if (ci === 4) renderUnresolved(); // Employer edited — refresh unresolved panel
    updateSummary();
    updateExportButton();
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    initialized = true;
    attachModeSwitcher();
    $('em-emp-file').addEventListener('change', e => {
      if (e.target.files[0]) handleEmpFile(e.target.files[0]);
      e.target.value = '';
    });
    $('em-employer-file').addEventListener('change', e => {
      if (e.target.files[0]) handleEmployerFile(e.target.files[0]);
      e.target.value = '';
    });
    $('em-fullcontractor-file').addEventListener('change', e => {
      if (e.target.files[0]) handleFullContractorFile(e.target.files[0]);
      e.target.value = '';
    });
    $('em-migrated-file').addEventListener('change', e => {
      if (e.target.files[0]) handleAlreadyMigratedFile(e.target.files[0]);
      e.target.value = '';
    });
    $('em-tpl-file').addEventListener('change', e => {
      if (e.target.files[0]) handleTplFile(e.target.files[0]);
      e.target.value = '';
    });
    const batch = $('em-batch');
    if (batch) batch.addEventListener('input', () => { if (formattedRows) updateSummary(); });
    $('em-run').addEventListener('click', runMigrate);
    $('em-export').addEventListener('click', () => { doExport(); });
    $('em-reset').addEventListener('click', reset);
    const copyBtn = $('em-create-copy');
    if (copyBtn) copyBtn.addEventListener('click', () => {
      const names = [...getEmployersToCreate().keys()];
      if (!names.length) { alert('No employers to create.'); return; }
      navigator.clipboard.writeText(names.join('\n')).then(
        () => alert('Copied ' + names.length + ' employer name' + (names.length === 1 ? '' : 's') + '.'),
        () => alert('Copy failed.'));
    });
    // Delegated inline-edit handlers bound ONCE to the persistent preview
    // table element (renderPreview only replaces its innerHTML).
    const ptbl = $('em-preview-table');
    if (ptbl) {
      ptbl.addEventListener('focusout', e => {
        const td = e.target.closest ? e.target.closest('td.em-cell') : null;
        if (td) commitCellEdit(td);
      });
      ptbl.addEventListener('keydown', e => {
        if (e.key === 'Escape' && selCells.size) { clearSelection(); return; }
        const td = e.target.closest ? e.target.closest('td.em-cell') : null;
        if (td && e.key === 'Enter') { e.preventDefault(); td.blur(); }
      });
      // Shift = rectangular range, Ctrl/Cmd = toggle. preventDefault on a
      // modified click stops the caret so it acts as pure selection; a plain
      // click clears the multi-selection and edits the single cell normally.
      ptbl.addEventListener('mousedown', e => {
        const td = e.target.closest ? e.target.closest('td.em-cell') : null;
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
    const selApply = $('em-sel-apply');
    if (selApply) {
      selApply.addEventListener('click', () => setSelectedCells($('em-sel-val').value));
      $('em-sel-clear').addEventListener('click', () => setSelectedCells(''));
      $('em-sel-deselect').addEventListener('click', clearSelection);
      $('em-sel-val').addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); setSelectedCells(e.target.value); }
      });
    }
  }

  window.empMigInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
