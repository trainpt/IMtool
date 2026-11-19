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

  // ─── Diagnostics for the Debug Dump ───
  // Every message this tool shows you, and every error thrown while it is on
  // screen, is kept here, so a dump taken after something breaks (even before
  // Build Preview has ever worked) still says what happened and where.
  const MODULE_FILE = 'employee-migrate.js';
  const events = [];
  function logEvent(kind, msg, err) {
    events.push({
      at: new Date().toISOString(),
      kind: kind,
      message: String(msg == null ? '' : msg).slice(0, 2000),
      stack: err && err.stack ? String(err.stack).split('\n').slice(0, 15).join('\n') : null
    });
    if (events.length > 200) events.shift();
    if (window.IMDebug) IMDebug.refresh('employee-migrate');
  }
  // Shadows window.alert inside this module so every message is recorded.
  function alert(msg) { logEvent('message', msg); window.alert(msg); }
  // A caught error: recorded with its stack, then shown.
  function fail(what, err, more) {
    const text = what + ': ' + (err && err.message ? err.message : err);
    logEvent('error', text, err);
    window.alert(text + (more || ''));
  }
  const paneShown = () => { const p = $('cmp-mode-legacy'); return !!(p && p.offsetParent !== null); };
  window.addEventListener('error', e => {
    const err = e.error;
    const here = String(e.filename || '').indexOf(MODULE_FILE) >= 0 ||
      !!(err && err.stack && String(err.stack).indexOf(MODULE_FILE) >= 0);
    // "ResizeObserver loop …" is the browser noting a resize it will finish next frame — not an error.
    if (/ResizeObserver loop/.test(e.message || '')) return;
    if (here || paneShown()) logEvent('uncaught error', e.message || (err && err.message), err);
  });
  window.addEventListener('unhandledrejection', e => {
    const err = e.reason;
    const here = !!(err && err.stack && String(err.stack).indexOf(MODULE_FILE) >= 0);
    if (here || paneShown()) logEvent('uncaught error (promise)', err && err.message ? err.message : err, err);
  });

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
  let existingIdx = null;      // { byNameDob, bySsn, byName } over the same 3.0 export (optional)
  let alreadyMigratedCount = 0;// per-build count of skips due to cross-reference
  let skippedExisting = [];    // [{ name, by, existing }] — rows left out as already in 3.0
  let tplData = null;          // { headers, employerList[], rawBuffer, fileName }
  let formattedRows = null;    // [[...34]] mapped rows (post Is-Active filter)
  let srcKept = null;          // source rows that survived the filter (parallel to formattedRows)
  let srcIdxKept = null;       // empData.rows index behind each formattedRows row
  let employerOverrides = {};  // ContractorID → chosen Employer (sticky across rebuilds)
  let employerIdsAssumed = false; // the contractors file had no ID column — IDs numbered by date created
  let contractorInfo = {};     // ContractorID → { phone, email } from the contractors file
  let employerDetails = {};    // ContractorID → what the client told us about a new employer (name, type, address…)
  let columnFills = {};        // colIdx → { val, mode:'all'|'blank' } (sticky bulk fill)
  let smartFixMap = {};        // "colIdx||UPPERVALUE" → template value (sticky Smart Fixes)
  let cellEdits = {};          // "srcIdx|colIdx" → value — hand edits, kept through rebuilds
  let removedSrc = new Set();  // empData.rows indexes you dropped (kept through rebuilds)
  let removedRows = new Set(); // formattedRows indexes excluded from preview + export (from removedSrc)
  let dupPossible = [];        // [{ ri, name, reason, existing }] — kept, but worth a look
  // Legacy writes its own employee ID into the name: "doe {1234}". Whether
  // 3.0 keeps it is the client's call, so it's a switch (on = as Legacy has it).
  let keepNameCode = true;
  let codeAnswered = false;    // the client already said whether the Legacy number stays in last names
  let altIdMode = null;        // null: Legacy's Alt ID as is · 'legacy' | 'altOrLegacy' | 'alt' | 'none' (the client's answer)
  let pvMode = 'all';          // preview rows: 'all' | 'fix' | 'chg' | 'any' (IMWorkbench.filterBar)
  let pvColsOnly = false;      // preview: only columns with a flag or a change
  let selCells = new Set();    // "ri:ci" keys of multi-selected preview cells
  let selAnchor = null;        // { ri, ci } anchor for shift-range selection
  let previewOrder = [];       // ri's in displayed order (for shift-range math)
  let initialized = false;

  // Output column → the Legacy column it is copied from (the fixed field map).
  const SRC_KEYS = {
    0: 'first name', 1: 'middle name', 2: 'last name', 3: 'birth date',
    5: 'alt id', 6: 'ssn', 8: 'hired date', 17: 'mobile phone #', 18: 'email',
    19: 'address 1', 20: 'address 2', 21: 'city', 22: 'state', 23: 'postal code', 24: 'country'
  };
  const NAME_CODE = /\s*\{\d+\}\s*$/;   // Legacy's " {3937}" employee-ID suffix
  const NAME_COLS = [0, 1, 2];
  const stripCode = v => String(v == null ? '' : v).replace(NAME_CODE, '').trim();

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
        // (e.g. "Sample Crew FLC "); trimming it ⇒ "Employer not found".
        const v = String(raw[r][ci] != null ? raw[r][ci] : '');
        if (v.trim() && arr.indexOf(v) < 0) arr.push(v);
      }
      if (arr.length) m.set(k, arr);
    });
    return m;
  }

  // ─── Load handlers ───
  function handleEmpFile(file) {
    return parseCsvFile(file).then(d => {
      const idx = {};
      d.headers.forEach((h, i) => { const k = norm(h); if (k && idx[k] == null) idx[k] = i; });
      if (idx['is active'] == null || idx['contractor id'] == null) {
        alert('Employee file must have "Is Active" and "Contractor ID" columns.');
        return;
      }
      empData = { headers: d.headers, rows: d.rows, fileName: file.name, idx };
      // Edits and removed rows point at this file's rows — a new file starts clean.
      cellEdits = {}; removedSrc = new Set();
      const coded = d.rows.filter(r => NAME_CODE.test(String(r[idx['last name']] == null ? '' : r[idx['last name']]))).length;
      if (coded) logEvent('note', coded + ' of ' + d.rows.length + ' last names carry a Legacy {code}.');
      $('em-emp-name').textContent = file.name;
      $('em-emp-meta').textContent = d.rows.length + ' rows · ' + d.headers.length + ' columns';
      maybeRun();
    }).catch(err => fail('Failed to read employee file', err));
  }

  function handleEmployerFile(file) {
    return parseCsvFile(file).then(d => {
      const col = (...ks) => d.headers.findIndex(h => ks.includes(norm(h)));
      const idI = col('id', 'contractor id'), nameI = col('name');
      const phoneI = col('phone', 'phone #', 'phone number'), emailI = col('email');
      const createdI = col('created on', 'created');
      if (nameI < 0) { alert('The contractors file needs a "Name" column.'); return; }
      const g = (r, i) => i < 0 ? '' : String(r[i] == null ? '' : r[i]).trim();
      const rows = d.rows.filter(r => g(r, nameI));
      const m = new Map();
      contractorInfo = {};
      employerIdsAssumed = idI < 0;
      if (idI >= 0) {
        rows.forEach(r => { const id = g(r, idI); if (id) { m.set(id, g(r, nameI)); contractorInfo[id] = { phone: g(r, phoneI), email: g(r, emailI) }; } });
      } else {
        // Legacy's contractors export has no ID column. Its IDs count up as
        // contractors are created, so number them by "Created On" (oldest = #1).
        // That's an assumption — it's shown in the Employers panels and the
        // client is asked to confirm the names against each group's people.
        rows.map((r, i) => ({ r, i, t: Date.parse(g(r, createdI)) || 0 }))
          .sort((a, b) => a.t - b.t || a.i - b.i)
          .forEach((x, k) => { const id = String(k + 1); m.set(id, g(x.r, nameI)); contractorInfo[id] = { phone: g(x.r, phoneI), email: g(x.r, emailI) }; });
        logEvent('note', 'Contractors file has no ID column — numbered by Created On: ' + [...m.entries()].map(([id, n]) => '#' + id + ' ' + n).join(', '));
      }
      employerMap = m;
      $('em-employer-name').textContent = file.name;
      $('em-employer-meta').textContent = m.size + ' contractor' + (m.size === 1 ? '' : 's') +
        (employerIdsAssumed ? ' · no ID column — numbered by date created (check under Employers)' : '');
      maybeRun();
    }).catch(err => fail('Failed to read active contractors file', err));
  }

  // Employer names compare without punctuation or a company suffix:
  // "Acme Farm-n-Pack, LLC" is the template's "Acme Farm-N-Pack".
  const employerKey = s => String(s == null ? '' : s).toUpperCase().replace(/&/g, ' AND ')
    .replace(/\b(LLC|L L C|INC|INCORPORATED|CO|CORP|CORPORATION|LTD|COMPANY|LP|LLP)\b\.?/g, ' ').replace(/[^A-Z0-9]+/g, '');
  function looseEmployer(name) {
    const list = (tplData && tplData.employerList) || [];
    const k = employerKey(name);
    if (!k) return null;
    const hits = list.filter(e => employerKey(e) === k);
    return hits.length === 1 ? hits[0] : null;
  }

  // Optional full (archived + unarchived) contractor list. Contractors in here
  // but NOT in the active list ⇒ archived ⇒ their employees are dropped.
  function handleFullContractorFile(file) {
    return parseCsvFile(file).then(d => {
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
    }).catch(err => fail('Failed to read full contractor file', err));
  }

  // Optional: a 3.0 employees export. We extract Alt ID and skip any Legacy
  // employee whose Alt ID is already present (avoids duplicate uploads).
  function handleAlreadyMigratedFile(file) {
    return parseCsvFile(file).then(d => {
      const col = (...names) => d.headers.findIndex(h => names.includes(norm(h).replace(/\*$/, '').trim()));
      const altI = col('alt id');
      const fnI = col('first name'), lnI = col('last name');
      const dobI = col('date of birth', 'birth date', 'dob');
      const ssnI = col('ssn', 'social security number');
      if (altI < 0 && (fnI < 0 || lnI < 0)) {
        alert('Already-in-3.0 file needs an "Alt ID" column, or "First Name" and "Last Name", to cross-reference.');
        return;
      }
      const g = (r, i) => i < 0 ? '' : String(r[i] == null ? '' : r[i]).trim();
      const s = new Set();
      const idx = { byNameDob: new Map(), bySsn: new Map(), byName: new Map() };
      d.rows.forEach(r => {
        const v = g(r, altI);
        if (v) s.add(v);
        const name = (g(r, fnI) + ' ' + g(r, lnI)).trim();
        const nk = nameKey(name);
        if (!nk) return;
        const who = name + (v ? ' (Alt ID ' + v + ')' : '');
        const dob = normDate(g(r, dobI));
        if (dob) idx.byNameDob.set(nk + '|' + dob, who);
        const sk = ssnKey(g(r, ssnI));
        if (sk) idx.bySsn.set(sk, who);
        if (!idx.byName.has(nk)) idx.byName.set(nk, who);
      });
      migratedAltIds = s;
      existingIdx = (fnI >= 0 && lnI >= 0) ? idx : null;
      $('em-migrated-name').textContent = file.name;
      $('em-migrated-meta').textContent = d.rows.length + ' employee' + (d.rows.length === 1 ? '' : 's') + ' · matched on ' +
        [altI >= 0 ? 'Alt ID' : '', existingIdx && dobI >= 0 ? 'name + birth date' : '', existingIdx && ssnI >= 0 ? 'SSN' : '']
          .filter(Boolean).join(', ');
      maybeRun();
    }).catch(err => fail('Failed to read 3.0 employees file', err));
  }

  function handleTplFile(file) {
    // Resolves once the file is read (Load everything places files in turn).
    return new Promise(done => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
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
        fail('Failed to read template', err);
      }
    };
    r.onerror = () => { alert('Failed to read template file.'); done(); };
    r.onloadend = () => done();
    r.readAsArrayBuffer(file);
    });
  }

  // Name / SSN keys for duplicate matching: case, accents, punctuation and
  // Legacy's {code} don't count; SSNs compare on their digits.
  function nameKey(s) {
    return stripCode(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
      .replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function ssnKey(s) {
    const d = String(s == null ? '' : s).replace(/\D/g, '');
    return d.length >= 7 && !/^0+$/.test(d) && !/^(\d)\1+$/.test(d) ? d : '';
  }

  // ─── Index helper for the employee source ───
  function eIdx(key) { return empData && empData.idx[key] != null ? empData.idx[key] : -1; }
  function srcCid(srcRow) {
    const i = eIdx('contractor id');
    return (i < 0 ? '' : String(srcRow[i] == null ? '' : srcRow[i]).trim()) || '(blank)';
  }

  // ─── Transformation ───
  // Filter (Is Active = true) + fixed Legacy→3.0 column map. Builds
  // formattedRows and the parallel srcKept / srcIdxKept arrays IN THE SAME
  // PASS — the panels rely on formattedRows[i] ↔ srcKept[i].
  function buildFormattedRowsRaw() {
    if (!empData) return null;
    const get = (row, key) => { const i = eIdx(key); return i < 0 ? '' : String(row[i] == null ? '' : row[i]).trim(); };
    const iaI = eIdx('is active');
    const altI = eIdx('alt id');
    srcKept = [];
    srcIdxKept = [];
    archivedStats = {};
    alreadyMigratedCount = 0;
    skippedExisting = [];
    const out = [];
    empData.rows.forEach((srcRow, srcIdx) => {
      // Active filter: Is Active column is the sole determinant. Skip if not
      // literally "true" (false / blank both excluded).
      const ia = iaI < 0 ? '' : String(srcRow[iaI] == null ? '' : srcRow[iaI]).trim().toLowerCase();
      if (ia !== 'true') return;
      // Cross-reference: skip employees already in 3.0 — on Alt ID, or on
      // name + birth date (Legacy leaves most Alt IDs blank).
      const dispName = (stripCode(get(srcRow, 'first name')) + ' ' + stripCode(get(srcRow, 'last name'))).trim();
      if (migratedAltIds && altI >= 0) {
        const aid = String(srcRow[altI] == null ? '' : srcRow[altI]).trim();
        if (aid && migratedAltIds.has(aid)) {
          alreadyMigratedCount++;
          skippedExisting.push({ name: dispName, by: 'Alt ID ' + aid, existing: '' });
          return;
        }
      }
      if (existingIdx) {
        const nk = nameKey(dispName), dob = normDate(get(srcRow, 'birth date'));
        const hit = nk && dob ? existingIdx.byNameDob.get(nk + '|' + dob) : null;
        if (hit) {
          alreadyMigratedCount++;
          skippedExisting.push({ name: dispName, by: 'Name + birth date', existing: hit });
          return;
        }
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
      // Alt ID — what 3.0's payroll export carries. altIdMode is the client's answer.
      const legacyId = get(srcRow, 'id'), altId = get(srcRow, 'alt id');
      r[5]  = altIdMode === 'legacy' ? legacyId : altIdMode === 'altOrLegacy' ? (altId || legacyId) : altIdMode === 'none' ? '' : altId;
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
      if (!keepNameCode) NAME_COLS.forEach(ci => { r[ci] = stripCode(r[ci]); });
      out.push(r);
      srcKept.push(srcRow);
      srcIdxKept.push(srcIdx);
    });

    // Canonicalize Employer* to the template's dropdown literal when it
    // matches (case-insensitive). A resolved name that ISN'T in the template's
    // Employer dropdown is KEPT in the column (per request) — it surfaces in
    // both the "Employers to Create in 3.0" list and the Unresolved Employers
    // panel so it can be created in PickTrace or remapped before upload.
    if (tplData && tplData.employerList && tplData.employerList.length) {
      const cm = buildCaseMap(tplData.employerList);
      out.forEach(r => {
        if (r[4]) r[4] = cm.get(String(r[4]).toUpperCase().trim()) || looseEmployer(r[4]) || r[4];
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

  // Raw rows, then (in this order) Smart Fixes, employer overrides, column
  // fills and hand edits. Removed rows and edits are keyed by the Legacy row
  // behind them, so restoring a contractor or flipping a switch (which changes
  // which rows exist) never moves them onto someone else.
  function rebuildFormattedRows() {
    formattedRows = buildFormattedRowsRaw();
    if (!formattedRows) return;
    removedRows = new Set();
    srcIdxKept.forEach((si, ri) => { if (removedSrc.has(si)) removedRows.add(ri); });
    if (Object.keys(smartFixMap).length) {
      formattedRows.forEach(r => {
        for (let ci = 0; ci < r.length; ci++) {
          if (!r[ci]) continue;
          const hit = smartFixMap[ci + '||' + String(r[ci]).toUpperCase().trim()];
          if (hit != null) r[ci] = hit;
        }
      });
    }
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
    formattedRows.forEach((r, ri) => {
      for (let ci = 0; ci < r.length; ci++) {
        const e = cellEdits[srcIdxKept[ri] + '|' + ci];
        if (e != null) r[ci] = e;
      }
    });
    findPossibleDuplicates();
  }

  function removeRow(ri) {
    removedRows.add(ri);
    if (srcIdxKept && srcIdxKept[ri] != null) removedSrc.add(srcIdxKept[ri]);
  }
  function restoreRows(list) {
    (list || [...removedRows]).forEach(ri => {
      removedRows.delete(ri);
      if (srcIdxKept && srcIdxKept[ri] != null) removedSrc.delete(srcIdxKept[ri]);
    });
  }

  // Kept, but worth a look: the same SSN, or the same name + birth date, twice
  // in this file; or someone in the 3.0 export with the same SSN or name
  // (rows matching on name + birth date there were already left out).
  function findPossibleDuplicates() {
    dupPossible = [];
    if (!formattedRows) return;
    const seenSsn = new Map(), seenNd = new Map();
    const disp = r => [stripCode(r[0]), stripCode(r[2])].filter(Boolean).join(' ');
    formattedRows.forEach((r, ri) => {
      const name = disp(r), nk = nameKey(name), sk = ssnKey(r[6]);
      const nd = nk && r[3] ? nk + '|' + r[3] : '';
      let reason = '', existing = '';
      if (sk && seenSsn.has(sk)) { reason = 'Same SSN as another row in this file'; existing = seenSsn.get(sk); }
      else if (nd && seenNd.has(nd)) { reason = 'Same name + birth date as another row in this file'; existing = seenNd.get(nd); }
      else if (existingIdx && sk && existingIdx.bySsn.has(sk)) { reason = 'SSN already in 3.0'; existing = existingIdx.bySsn.get(sk); }
      else if (existingIdx && nk && existingIdx.byName.has(nk)) { reason = 'Name already in 3.0 (birth date differs or missing)'; existing = existingIdx.byName.get(nk); }
      if (reason) dupPossible.push({ ri, name, reason, existing });
      const who = name + ' (row ' + (ri + 1) + ')';
      if (sk && !seenSsn.has(sk)) seenSsn.set(sk, who);
      if (nd && !seenNd.has(nd)) seenNd.set(nd, who);
    });
  }

  // A crash while building used to surface as "Failed to read <the file you
  // just loaded>", because each file handler's catch swallowed it. Catch it
  // here instead, record the stack, and say it was the build.
  // The contractor list is optional: without it every Contractor ID lands in
  // the Employers panel, where you pick the 3.0 employer for each one.
  const ready = () => !!(empData && tplData);
  function runMigrate() {
    if (!ready()) {
      alert('Upload the Legacy employee CSV and the 3.0 employee template first.');
      return;
    }
    try { runMigrateInner(); }
    catch (err) {
      fail('Build Preview failed', err,
        '\n\nUse Debug Dump to save what happened, and send it over.');
    }
    if (window.IMDebug) IMDebug.refresh('employee-migrate');
  }

  // Your choices (employers, fills, Smart Fixes, edits, removed rows) stay
  // through a re-run — loading another file shouldn't undo them. Reset clears
  // them; a new employee CSV clears the ones tied to its rows.
  function runMigrateInner() {
    selCells = new Set(); selAnchor = null;
    rebuildFormattedRows();
    renderAll();
  }

  function renderAll() {
    renderArchived();
    renderUnresolved();
    renderSmartFixes();
    renderPossible();
    renderExisting();
    renderNames();
    renderBulkEdit();
    renderPreview();
    updateSummary();
  }
  // After any change that alters rows: rebuild, then redraw everything.
  function rebuildAndRender() {
    rebuildFormattedRows();
    renderAll();
  }

  function maybeRun() {
    const ok = ready();
    if (ok) runMigrate();
    else if (window.IMDebug) IMDebug.refresh('employee-migrate');
    if (window.IMLegacyPage) IMLegacyPage.refresh();
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
        rebuildAndRender();
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
      if (employerOverrides[srcCid(srcKept[ri])]) return; // you (or the client) named one — it's under Build in PickTrace if it's new
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

  // What the client told us about a new employer, for the Build in PickTrace list.
  function detailCells(cids) {
    const d = cids.map(c => employerDetails[c]).find(Boolean);
    if (!d) return '<td colspan="3" class="text-muted small">No details yet — the client is asked for them</td>';
    const addr = [d.address1, d.address2, d.city, [d.state, d.zip].filter(Boolean).join(' '), d.country].filter(Boolean).join(', ');
    return '<td>' + escHtml(d.type || '—') + (d.type === 'Contractor' && d.flc ? '<div class="text-muted small">FLC: ' + escHtml(d.flc) + '</div>' : '') + '</td>' +
      '<td>' + escHtml(addr || '—') + '</td><td>' + escHtml([d.phone, d.email].filter(Boolean).join(' · ') || '—') + '</td>';
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
      '<th>Legacy Contractor IDs</th><th>Type</th><th>Address</th><th>Phone / email</th></tr></thead><tbody>';
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
        '<td>' + escHtml([...info.cids].sort((x, y) => x - y).join(', ')) + '</td>' + detailCells([...info.cids]) + '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    const note = $('em-create-note') || (() => { const p = document.createElement('p'); p.id = 'em-create-note'; p.className = 'cmp-sites-hint'; tbl.parentNode.before(p); return p; })();
    note.innerHTML = employerIdsAssumed && employerMap
      ? '<b>Check:</b> the contractors file has no ID column, so Legacy contractor numbers were matched by the order they were created — ' +
        [...employerMap.entries()].map(([id, n]) => '#' + escHtml(id) + ' ' + escHtml(n)).join(' · ') + '. The client is asked to confirm each name against its people.'
      : '';
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
        : employerMap ? 'No employer record for this Contractor ID'
        : 'No contractor list loaded — pick its 3.0 employer';
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
    renderAll();
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
    // A whole-column fill replaces earlier hand edits in that column.
    if (mode !== 'blank') Object.keys(cellEdits).forEach(k => { if (+k.split('|')[1] === idx) delete cellEdits[k]; });
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (mode === 'blank') { if (!r[idx]) r[idx] = val; }
      else r[idx] = val;
    });
    renderAll();
  }

  // Empties the column for good: kept as a blank fill so a rebuild doesn't
  // bring the Legacy values back.
  function clearColumnFill(idx) {
    columnFills[idx] = { val: '', mode: 'all' };
    Object.keys(cellEdits).forEach(k => { if (+k.split('|')[1] === idx) delete cellEdits[k]; });
    formattedRows.forEach((r, ri) => { if (!removedRows.has(ri)) r[idx] = ''; });
    renderAll();
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
    selCells.forEach(k => {
      const i = k.indexOf(':');
      const ri = +k.slice(0, i), ci = +k.slice(i + 1);
      if (!formattedRows[ri]) return;
      formattedRows[ri][ci] = val;
      cellEdits[srcIdxKept[ri] + '|' + ci] = val;
    });
    renderAll();               // re-renders + re-paints the kept selection
  }

  // ─── Preview ───
  // Red: a required (*) cell is empty. Amber: a value the template's dropdown
  // doesn't list (PickTrace would reject it). Corner tick: changed from what
  // Legacy had — hover for was → now and why.
  const FLAG_STYLE = { req: 'background:#fee2e2;color:#7f1d1d;', off: 'background:#fffbeb;color:#92400e;' };
  const FLAG_TIP = { req: 'Required — empty', off: 'Not in the template’s dropdown — see Not in dropdown' };
  let optCache = null;   // colIdx → Set<UPPER option> | null, per render
  function optSet(ci) {
    if (!optCache) optCache = {};
    if (!(ci in optCache)) {
      const o = colOptions(EMP_HEADERS[ci]);
      optCache[ci] = o && o.length ? new Set(o.map(v => String(v).toUpperCase().trim())) : null;
    }
    return optCache[ci];
  }
  function cellFlag(ri, ci) {
    const v = formattedRows[ri][ci];
    if (!v) return /\*$/.test(EMP_HEADERS[ci]) ? 'req' : '';
    const s = optSet(ci);
    return s && !s.has(String(v).toUpperCase().trim()) ? 'off' : '';
  }

  // "Was … → now … · why" when the cell differs from what Legacy had.
  function changeOf(ri, ci) {
    const src = srcKept ? srcKept[ri] : null;
    if (!src) return '';
    const key = SRC_KEYS[ci];
    const now = formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci]);
    const edit = cellEdits[srcIdxKept[ri] + '|' + ci];
    if (ci === 4) {   // Employer: resolved from Contractor ID, not copied
      const cid = srcCid(src);
      if (edit != null) return IMWorkbench.changeNote('', now, 'Your edit') || '';
      if (employerOverrides[cid] === now) return 'Contractor ID ' + cid + ' → “' + now + '” · Employer you picked';
      const f = columnFills[4];
      if (f && f.val === now) return 'Contractor ID ' + cid + ' → “' + now + '” · Column fill you applied';
      return '';
    }
    if (ci === 7) {   // Gender from "Is Male"
      const i = eIdx('is male'), was = i < 0 ? '' : String(src[i] == null ? '' : src[i]).trim().toLowerCase();
      const derived = was === 'true' ? 'MALE' : was === 'false' ? 'FEMALE' : '';
      return derived.toUpperCase() === now.toUpperCase().trim() ? '' :
        IMWorkbench.changeNote(derived, now, edit != null ? 'Your edit' : 'Changed after import');
    }
    if (!key) {
      return now ? IMWorkbench.changeNote('', now, edit != null ? 'Your edit'
        : columnFills[ci] && columnFills[ci].val === now ? 'Column fill you applied' : 'Filled after import') : '';
    }
    const i = eIdx(key);
    const was = i < 0 ? '' : String(src[i] == null ? '' : src[i]).trim();
    let why = 'Changed after import';
    if (edit != null) why = 'Your edit';
    else if (columnFills[ci] && columnFills[ci].val === now) why = 'Column fill you applied';
    else if (smartFixMap[ci + '||' + was.toUpperCase()] === now) why = 'Smart Fix you applied';
    else if (NAME_COLS.includes(ci) && stripCode(was) === now) why = 'Legacy {code} taken out';
    else if ((ci === 3 || ci === 8) && normDate(was) === now) return '';   // only the date format changed
    else if (/country$/i.test(EMP_HEADERS[ci])) why = 'Matched to the template’s country';
    return IMWorkbench.changeNote(was, now, why);
  }

  function renderPreview() {
    if (!empData || !tplData || !formattedRows) return;
    const sec = $('em-section-preview');
    const tbl = $('em-preview-table');
    sec.style.display = '';
    optCache = null;
    const possibleSet = new Set(dupPossible.map(d => d.ri));
    const vis = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    // Every row, drawn as you scroll (IMWorkbench.rows). A row needs a fix
    // when a cell is red or amber, or it may duplicate someone.
    const info = new Map(), counts = { all: vis.length, fix: 0, chg: 0, any: 0 };
    vis.forEach(ri => {
      const f = [], c = [];
      for (let i = 0; i < EMP_HEADERS.length; i++) {
        if (cellFlag(ri, i)) f.push(i);
        if (changeOf(ri, i)) c.push(i);
      }
      const fix = f.length > 0 || possibleSet.has(ri);
      info.set(ri, { f, c, fix });
      if (fix) counts.fix++;
      if (c.length) counts.chg++;
      if (fix || c.length) counts.any++;
    });
    const show = vis.filter(ri => IMWorkbench.rowMatch(pvMode, info.get(ri).fix, info.get(ri).c.length));
    let cols = EMP_HEADERS.map((_, i) => i);
    if (pvColsOnly) {
      const keep = new Set([0, 2]);
      show.forEach(ri => { const x = info.get(ri); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      cols = cols.filter(i => keep.has(i));
    }
    const visSet = new Set(vis);
    selCells.forEach(k => { if (!visSet.has(+k.slice(0, k.indexOf(':')))) selCells.delete(k); });
    const head = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>' + cols.map(i =>
      '<th' + (/\*$/.test(EMP_HEADERS[i]) ? ' style="color:#dc2626;"' : '') + '>' + escHtml(EMP_HEADERS[i]) + '</th>').join('') + '</tr></thead>';
    const rowHtml = ri => {
      const row = formattedRows[ri];
      const dup = possibleSet.has(ri) ? dupPossible.find(d => d.ri === ri) : null;
      let html = '<tr' + (dup ? ' title="Possible duplicate — ' + escHtml(dup.reason) + '" style="background:#fffbeb;"' : '') +
        '><td style="text-align:center;padding:0;">' +
        '<button class="em-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      cols.forEach(i => {
        const fl = cellFlag(ri, i), note = changeOf(ri, i);
        const tip = [FLAG_TIP[fl], note].filter(Boolean).join(' · ');
        html += '<td class="em-cell' + (note ? ' is-chg' : '') + '" contenteditable="true" spellcheck="false" ' +
          'data-ri="' + ri + '" data-ci="' + i + '"' + (fl ? ' style="' + FLAG_STYLE[fl] + '"' : '') +
          (tip ? ' title="' + escHtml(tip) + '"' : '') + '>' + escHtml(row[i]) + '</td>';
      });
      return html + '</tr>';
    };
    // Shift-click ranges follow the order on screen, sorted or not.
    IMWorkbench.rows(tbl, { head: head, ids: show, row: rowHtml, cols: cols.length + 1,
      after: () => { previewOrder = IMWorkbench.rowOrder(tbl); paintSelection(); },
      keys: [null].concat(cols.map(i => EMP_HEADERS[i])), cell: (ri, p) => formattedRows[ri][cols[p - 1]],
      widths: [28].concat(cols.map(i => IMWorkbench.fitWidth(EMP_HEADERS[i], show.map(ri => formattedRows[ri][i])))),
      empty: pvMode === 'fix' ? 'No row needs a fix ✓' : 'Nothing here has been changed.' });
    const hint = sec.querySelector('.cmp-sites-hint');
    let ht = IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly, shownCols: cols.length, totalCols: EMP_HEADERS.length }) +
      '<span class="wb-filter-note">' + (show.length === vis.length ? 'All ' + vis.length : show.length + ' of ' + vis.length) +
      ' rows · editable — click a cell, type, then click away (or press Enter) to save. <b>Shift+click</b> selects a range, <b>Ctrl/Cmd+click</b> single cells. ' +
      '<span style="background:#fffbeb;">Amber rows</span> are possible duplicates. Click <b>×</b> to drop a row from the export.</span>';
    if (removedRows.size) ht += ' <button class="btn btn-ghost btn-sm" id="em-restore-rows">' +
      'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    if (hint) hint.innerHTML = ht;
    // Rows come and go as you scroll, and the bar is redrawn every render —
    // their handlers live on the table / section, once.
    if (!tbl._emRowsWired) {
      tbl._emRowsWired = true;
      tbl.addEventListener('click', e => {
        const rm = e.target.closest('.em-row-remove');
        if (!rm) return;
        e.stopPropagation();
        removeRow(+rm.dataset.ri);
        renderAll();
      });
      sec.addEventListener('change', e => {
        if (e.target.classList.contains('wb-rowfilter')) { pvMode = e.target.value; IMWorkbench.rowsTop(tbl); renderPreview(); }
        else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
      });
      sec.addEventListener('click', e => {
        if (!e.target.closest('#em-restore-rows')) return;
        restoreRows();
        renderAll();
      });
    }
    paintSelection();
    updateExportButton();
  }

  // ─── Not in dropdown (Smart Fixes) ───
  // Every dropdown-backed column except Employer (it has its own panel): the
  // values the template doesn't list, each with a picker. A close match
  // (case, spacing, punctuation, one inside the other) is suggested.
  const squash = s => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  function suggestOption(val, opts) {
    const q = squash(val);
    if (!q) return '';
    const exact = opts.find(o => squash(o) === q);
    if (exact) return exact;
    const inside = opts.filter(o => { const s = squash(o); return s && (s.indexOf(q) >= 0 || q.indexOf(s) >= 0); });
    return inside.length === 1 ? inside[0] : '';
  }
  function computeSmartFixes() {
    if (!formattedRows || !tplData) return [];
    const out = [];
    for (let ci = 0; ci < EMP_HEADERS.length; ci++) {
      if (ci === 4) continue;
      const opts = colOptions(EMP_HEADERS[ci]);
      if (!opts || !opts.length) continue;
      const set = new Set(opts.map(o => String(o).toUpperCase().trim()));
      const seen = new Map();
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri) || !r[ci]) return;
        const u = String(r[ci]).toUpperCase().trim();
        if (set.has(u)) return;
        const e = seen.get(u) || { colIdx: ci, header: EMP_HEADERS[ci], from: r[ci], count: 0, options: opts, suggestion: suggestOption(r[ci], opts) };
        e.count++;
        seen.set(u, e);
      });
      seen.forEach(e => out.push(e));
    }
    out.sort((a, b) => (b.suggestion ? 1 : 0) - (a.suggestion ? 1 : 0) || b.count - a.count);
    return out;
  }
  let smartFixOpts = new Map();
  function renderSmartFixes() {
    const sec = $('em-section-smartfix'), tbl = $('em-smartfix-table');
    if (!sec || !tbl) return;
    const fixes = computeSmartFixes();
    if (!fixes.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    let rows = 0, needPick = 0;
    fixes.forEach(f => { rows += f.count; if (!f.suggestion) needPick++; });
    $('em-smartfix-title').textContent = 'Not in dropdown (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + rows + ' row' + (rows === 1 ? '' : 's') +
      (needPick ? ' — ' + needPick + ' need your choice' : '') + ')';
    const optsByCol = new Map();
    let html = '<thead><tr><th>Column</th><th>Value in Legacy</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    fixes.forEach(f => {
      optsByCol.set(f.colIdx, f.options);
      html += '<tr><td>' + escHtml(f.header.replace(/\*$/, '')) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td>' + IMWorkbench.pickButton(f.suggestion || '', ' data-ci="' + f.colIdx + '" style="min-width:180px;"', { kind: 'fix', cls: 'em-smartfix-pick input-field' }) +
          (f.suggestion ? ' <span class="text-muted small">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') + '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm em-smartfix-apply" data-ci="' + f.colIdx + '" data-from="' + escHtml(f.from) + '">Apply</button></td></tr>';
    });
    tbl.innerHTML = html + '</tbody>';
    smartFixOpts = optsByCol;
    IMWorkbench.wirePicks(tbl, ci => smartFixOpts.get(ci), { kind: 'fix' });
    if (!tbl._emWired) {
      tbl._emWired = true;
      tbl.addEventListener('click', e => {
        const b = e.target.closest('.em-smartfix-apply');
        if (!b) return;
        const pick = b.closest('tr').querySelector('.em-smartfix-pick');
        const to = pick ? String(pick.value || '').trim() : '';
        if (!to) { alert('Pick a value to set "' + b.dataset.from + '" to first.'); return; }
        smartFixMap[b.dataset.ci + '||' + String(b.dataset.from).toUpperCase().trim()] = to;
        rebuildAndRender();
      });
    }
  }
  function applyAllSmartFixes() {
    const tbl = $('em-smartfix-table');
    if (!tbl) return;
    let n = 0;
    tbl.querySelectorAll('.em-smartfix-apply').forEach(b => {
      const pick = b.closest('tr').querySelector('.em-smartfix-pick');
      const to = pick ? String(pick.value || '').trim() : '';
      if (!to) return;
      smartFixMap[b.dataset.ci + '||' + String(b.dataset.from).toUpperCase().trim()] = to;
      n++;
    });
    if (!n) { alert('No values chosen yet. Pick a target for at least one row (suggested ones are pre-filled).'); return; }
    rebuildAndRender();
  }

  // ─── Possible duplicates ───
  function renderPossible() {
    const sec = $('em-section-possible'), tbl = $('em-possible-table');
    if (!sec || !tbl) return;
    const live = dupPossible.filter(d => !removedRows.has(d.ri));
    if (!live.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    $('em-possible-title').textContent = 'Possible duplicates (' + live.length + ')';
    let html = '<thead><tr><th>Employee</th><th>Why</th><th>Matches</th><th></th></tr></thead><tbody>';
    live.forEach(d => {
      html += '<tr><td><b>' + escHtml(d.name || '—') + '</b></td>' +
        '<td><span style="color:#b45309;font-weight:600;">' + escHtml(d.reason) + '</span></td>' +
        '<td>' + escHtml(d.existing || '') + '</td>' +
        '<td><button class="btn btn-ghost btn-sm em-possible-drop" data-ri="' + d.ri + '" title="Leave this row out of the export">Exclude</button></td></tr>';
    });
    tbl.innerHTML = html + '</tbody>';
    if (!tbl._emWired) {
      tbl._emWired = true;
      tbl.addEventListener('click', e => {
        const b = e.target.closest('.em-possible-drop');
        if (!b) return;
        removeRow(+b.dataset.ri);
        renderAll();
      });
    }
  }

  // ─── Already in 3.0 (left out) ───
  function renderExisting() {
    const sec = $('em-section-existing'), tbl = $('em-existing-table');
    if (!sec || !tbl) return;
    if (!skippedExisting.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    $('em-existing-title').textContent = 'Already in 3.0 — left out (' + skippedExisting.length + ')';
    tbl.innerHTML = '<thead><tr><th>Legacy employee</th><th>Matched on</th><th>In 3.0 as</th></tr></thead><tbody>' +
      skippedExisting.map(s => '<tr><td>' + escHtml(s.name) + '</td><td>' + escHtml(s.by) + '</td><td>' + escHtml(s.existing) + '</td></tr>').join('') +
      '</tbody>';
  }

  // ─── Legacy codes in names ───
  function namesWithCode() {
    if (!srcKept) return 0;
    const i = eIdx('last name'), f = eIdx('first name');
    return srcKept.filter(r => NAME_CODE.test(String(i < 0 ? '' : r[i] || '')) || NAME_CODE.test(String(f < 0 ? '' : r[f] || ''))).length;
  }
  function renderNames() {
    const sec = $('em-section-names'), body = $('em-names-body');
    if (!sec || !body) return;
    const n = namesWithCode();
    if (!n) { sec.style.display = 'none'; body.innerHTML = ''; return; }
    sec.style.display = '';
    const li = eIdx('last name');
    const sample = srcKept.map(r => String(li < 0 ? '' : r[li] || '')).find(v => NAME_CODE.test(v)) || '';
    body.innerHTML = '<p class="cmp-sites-hint"><b>' + n + '</b> of ' + srcKept.length + ' names end in a Legacy code like <code>' +
      escHtml(sample) + '</code> &mdash; the employee’s Legacy ID. Keep it as part of the name in 3.0, or take it out ' +
      '(the client decides; Alt ID is what 3.0’s payroll export carries).</p>' +
      '<label class="cmp-export-option"><input type="checkbox" id="em-keep-code"' + (keepNameCode ? ' checked' : '') + '> ' +
      '<span>Keep the Legacy code in names (<code>' + escHtml(sample) + '</code>)</span></label>' +
      (keepNameCode ? '' : '<p class="cmp-sites-hint" style="color:#15803d;">✓ Taken out &mdash; <code>' + escHtml(stripCode(sample)) + '</code></p>');
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
    const limit = batchLimit();
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
    if (wb) wb.refresh();
  }

  function updateExportButton() {
    if (window.IMLegacyPage) IMLegacyPage.refresh();
    if (window.IMDebug) IMDebug.refresh('employee-migrate');
  }

  // ─── The one Legacy page (legacy-page.js) ───
  // This engine's share of the page: its section counts, its files for
  // Export all, and how "Upload all files" recognises its files. `wb` stands
  // in for the shared workbench, so the rest of this module calls
  // wb.refresh() / wb.open() as before.
  const wb = {
    refresh: () => { if (window.IMLegacyPage) IMLegacyPage.refresh(); },
    open: k => { if (window.IMLegacyPage) IMLegacyPage.open(k); },
    reset: () => {}
  };
  const visRows = () => formattedRows ? formattedRows.filter((_, i) => !removedRows.has(i)) : [];
  const COUNTS = {
    required: () => { const rows = visRows(); return EMP_HEADERS.filter((h, i) => /\*$/.test(h) && rows.some(r => !r[i])).length; },
    employers: () => getUnresolved().size,
    dropdown: () => computeSmartFixes().length,
    build: () => pendingEmployersToCreate().length,
    possible: () => dupPossible.filter(d => !removedRows.has(d.ri)).length,
    names: () => keepNameCode ? namesWithCode() : 0,
    archived: () => archivedRemovedCount(),
    existing: () => skippedExisting.length
  };
  function registerPage() {
    if (!window.IMLegacyPage) return;
    IMLegacyPage.register('em', {
      ready: () => ready(),
      built: () => !!(ready() && formattedRows),
      rows: () => visRows().length,
      count: key => COUNTS[key] ? COUNTS[key]() : 0,
      run: () => runMigrate(),
      reset: () => reset(),
      exportParts: limit => exportParts(limit),
      questions: () => questions(),
      applyAnswers: a => applyAnswers(a),
      debugReady: () => !!(empData || employerMap || fullContractorMap || migratedAltIds || tplData || events.length),
      drop: {
        kinds: { emp: 'Legacy employee export', tpl: '3.0 Employees template', employer: 'Active contractors',
          fullcontractor: 'All contractors (archived + active)', migrated: 'Employees already in 3.0' },
        order: ['tpl', 'employer', 'fullcontractor', 'migrated', 'emp'],
        classify: i => {
          const H = (...k) => IMLegacyDrop.has(i, ...k), A = (...k) => IMLegacyDrop.hasAny(i, ...k);
          if (i.isTemplate) return i.norm[0] === 'first name' && A('employer') ? 'tpl' : null;
          if (H('contractor id') && A('is active', 'is valid')) return 'emp';
          if (H('first name', 'last name') && !A('contractor id')) return 'migrated';   // a 3.0 export
          if (H('id', 'name') && !A('first name') && i.norm.filter(Boolean).length <= 6) return 'contractor';   // settled below
          // Legacy's contractors export: Name, Phone, Email, Created On — no ID column (numbered on load).
          if (H('name') && A('created on', 'last modified') && A('phone', 'email') && !A('first name', 'contractor id', 'site_id', 'virtual', 'address1', 'job')) return 'contractor';
          return null;
        },
        // Contractor lists: one is the active list; with two, the bigger file
        // is the full (archived + active) one.
        arrange: infos => {
          const cs = infos.filter(i => i.key === 'contractor').sort((a, b) => b.size - a.size);
          cs.forEach(i => { i.key = null; i.guess = 'employer'; });
          if (cs.length === 1) cs[0].key = 'employer';
          else if (cs.length >= 2) { cs[0].key = 'fullcontractor'; cs[1].key = 'employer'; }
        },
        place: (key, file) => ({
          emp: handleEmpFile, tpl: handleTplFile, employer: handleEmployerFile,
          fullcontractor: handleFullContractorFile, migrated: handleAlreadyMigratedFile
        })[key](file)
      }
    });
  }
  // One "Max rows per file" for the whole page.
  function batchLimit() {
    const n = parseInt(($('lm-batch') || $('em-batch') || {}).value, 10);
    return Number.isFinite(n) && n >= 1 ? n : 5000;
  }

  // ─── Debug dump ───
  // The Legacy columns this module reads. A missing one never stops the
  // build; it just leaves its 3.0 column blank, so the dump lists each one
  // and whether the uploaded file has it.
  const EMP_KEYS = ['is active', 'contractor id', 'first name', 'middle name', 'last name',
    'birth date', 'alt id', 'ssn', 'is male', 'hired date', 'mobile phone #', 'email',
    'address 1', 'address 2', 'city', 'state', 'postal code', 'country'];
  function columnCheck() {
    if (!empData || !empData.idx) return null;
    return EMP_KEYS.map(k => ({ column: k, found: empData.idx[k] != null,
      sourceIndex: empData.idx[k] != null ? empData.idx[k] : null,
      sourceHeader: empData.idx[k] != null ? empData.headers[empData.idx[k]] : null }));
  }
  // The output is written by position, so the template's header row has to
  // line up with EMP_HEADERS column for column.
  function headerCheck() {
    if (!tplData || !tplData.headers) return null;
    const diffs = [];
    for (let i = 0; i < Math.max(EMP_HEADERS.length, tplData.headers.length); i++) {
      const want = EMP_HEADERS[i] || null, got = tplData.headers[i] || null;
      if (norm(want || '') !== norm(got || '')) diffs.push({ index: i, expected: want, template: got });
    }
    return { matches: !diffs.length, expectedCount: EMP_HEADERS.length,
      templateCount: tplData.headers.length, differences: diffs };
  }

  function collectDebug() {
    const diagnostics = {
      note: 'events = every message the tool showed and every error thrown while this tool was open, oldest first.',
      events: events.slice(),
      legacyEmployeeColumns: columnCheck(),
      templateHeaders: headerCheck(),
      onScreen: {
        summary: $('em-summary') && $('em-summary').style.display !== 'none' ? $('em-summary').innerText : null,
        buildPreviewEnabled: !!($('lm-run') && !$('lm-run').disabled),
        exportEnabled: !!($('lm-export') && !$('lm-export').disabled)
      }
    };
    // Not built yet (a file missing, or the build itself failed): dump what
    // is loaded so the problem can still be seen.
    if (!ready() || !formattedRows) {
      const missing = [];
      if (!empData) missing.push('Legacy employee CSV');
      if (!tplData) missing.push('3.0 employee template');
      return partialDump(diagnostics, missing.length ? 'Still missing: ' + missing.join(', ') + '.'
        : 'All required files are loaded but there is no preview: the build failed. See diagnostics.events.');
    }
    // If the full dump trips over bad data, keep everything else.
    try { return fullDump(diagnostics); }
    catch (err) {
      const d = partialDump(diagnostics, 'The preview is built, but the full dump failed — see collectorError.');
      d.collectorError = { message: err && err.message ? err.message : String(err), stack: err && err.stack ? err.stack : null };
      return d;
    }
  }

  function debugInputs() {
    return [
      IMDebug.file('legacy employee export', empData),
      { role: 'active contractor list', loaded: !!employerMap,
        contractorCount: employerMap ? employerMap.size : null },
      { role: 'full contractor list (archived + active)', loaded: !!fullContractorMap,
        contractorCount: fullContractorMap ? fullContractorMap.size : null,
        note: 'Optional. Without it, archived contractors cannot be told apart from unknown ones.' },
      { role: 'employees already in PickTrace 3.0', loaded: !!migratedAltIds,
        altIdCount: migratedAltIds ? migratedAltIds.size : null,
        note: 'Optional. Joined on Alt ID to skip employees already migrated.' },
      IMDebug.file('personalized bulk Employees template', tplData, {
        employerList: tplData ? tplData.employerList || null : null,
        dropdowns: tplData && tplData.dropdowns ? [...tplData.dropdowns.entries()].map(([k, v]) =>
          ({ column: k, count: v.length, values: v.slice(0, 300) })) : [] })
    ];
  }

  function partialDump(diagnostics, reason) {
    return {
      state: { built: !!formattedRows, reason: reason },
      diagnostics: diagnostics,
      inputs: debugInputs(),
      asked: [],
      output: null
    };
  }

  function fullDump(diagnostics) {
    const H = EMP_HEADERS;
    const keptIdxs = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    const keptRows = keptIdxs.map(i => formattedRows[i]);

    // This module uses a fixed Legacy→3.0 field map rather than auto-matching,
    // so the mapping section reports that fixed map by source header name.
    const FIXED_MAP = SRC_KEYS;
    const colToSrc = {};
    const cols = H.map((h, i) => {
      const key = FIXED_MAP[i];
      const si = key ? eIdx(key) : -1;
      colToSrc[i] = si;
      let note = null, match = 'unmapped';
      if (key) { match = si >= 0 ? 'fixed legacy field map ("' + key + '")' : 'fixed map — source column "' + key + '" not present'; }
      else if (i === 4) { match = 'derived'; note = 'Employer* is resolved from Contractor ID via the contractor list, then canonicalized to the template dropdown.'; }
      else if (i === 7) { match = 'derived'; note = 'Gender derived from the legacy "Is Male" boolean.'; }
      else if (i === 9) { match = 'left blank'; note = 'Start Date is intentionally blank.'; }
      else note = 'Not produced by the legacy map — fill it from the Column Fill panel if PickTrace needs it.';
      return { index: i, header: h, required: /\*$/.test(h), srcIndex: si, match, note,
        sample: si >= 0 && empData.rows[0] ? empData.rows[0][si] : null };
    });

    const emptyReqCols = H.map((h, i) => {
      if (!/\*$/.test(h)) return null;
      let n = 0; keptIdxs.forEach(ri => { if (!formattedRows[ri][i]) n++; });
      return n ? { column: h, emptyRows: n } : null;
    }).filter(Boolean);
    const unresolved = getUnresolved();
    const pending = pendingEmployersToCreate();
    const archivedRemoved = archivedRemovedCount();

    const asked = [
      IMDebug.ask('empty-required', 'required', 'Required columns with empty cells', {
        count: emptyReqCols.length, blocksExport: true, items: emptyReqCols }),
      IMDebug.ask('unresolved-employers', 'data', 'Contractor IDs that resolved to no Employer', {
        count: unresolved.size, blocksExport: true,
        detail: 'These rows export with a blank Employer*. Map each Contractor ID, or the upload is rejected.',
        items: [...unresolved.entries()].map(([cid, info]) => ({ contractorId: cid,
          rows: info.count, sampleEmployee: info.sampleName || null,
          legacyName: info.raw || null })) }),
      IMDebug.ask('employers-to-create', 'dropdown', 'Employers not yet in PickTrace 3.0', {
        count: pending.length, blocksExport: true,
        detail: 'Create these in 3.0 first, or tick them off the checklist once created.',
        items: pending.map(([name, info]) => ({ employer: name,
          rows: info && info.count != null ? info.count : null })) }),
      IMDebug.ask('archived-contractors', 'data', 'Archived contractors whose employees were dropped', {
        count: Object.keys(archivedStats).filter(cid => !restoredArchived.has(cid)).length,
        blocksExport: false,
        detail: archivedRemoved + ' employee rows removed. Restore a contractor to keep its employees.',
        items: Object.keys(archivedStats).filter(cid => !restoredArchived.has(cid))
          .map(cid => ({ contractorId: cid, name: archivedStats[cid].name, rows: archivedStats[cid].count })) }),
      IMDebug.ask('smart-fixes', 'smartfix', 'Values not in a template dropdown, awaiting a Smart Fix', {
        blocksExport: false, items: computeSmartFixes().map(x => ({ column: x.header, value: x.from, rows: x.count, suggestion: x.suggestion || null })) }),
      IMDebug.ask('possible-duplicates', 'data', 'Possible duplicates kept but flagged', {
        blocksExport: false, items: dupPossible.filter(d => !removedRows.has(d.ri)).map(d => ({ row: d.ri, name: d.name, reason: d.reason, matches: d.existing })) })
    ];

    const prov = IMDebug.deriveProvenance({
      headers: H, rows: formattedRows, srcRows: srcKept, colToSrc: colToSrc, fills: columnFills });
    const out = IMDebug.output(H, keptRows, {
      note: 'exportedRowIndexes[n] is the formattedRows index behind exported row n. Export is batched — see settings.batchSize.' });
    out.exportedRowIndexes = keptIdxs;

    const batch = batchLimit();
    const skippedInactive = empData.rows.length - formattedRows.length - archivedRemoved - alreadyMigratedCount;

    return {
      settings: { batchSize: batch, outputFiles: keptRows.length ? Math.ceil(keptRows.length / batch) : 0,
        legacyCodeInNames: keepNameCode ? 'kept' : 'taken out', namesCarryingCode: namesWithCode(),
        contractorListLoaded: !!employerMap },
      state: { built: true },
      diagnostics: diagnostics,
      inputs: debugInputs(),
      mapping: IMDebug.mapping(cols, empData.headers, {
        srcRows: empData.rows,
        note: 'This module does not auto-match columns — it applies a fixed Legacy→3.0 field map.' }),
      derived: [
        { kind: 'employer resolution',
          note: 'Contractor ID → active contractor name → canonicalized to the template\'s Employer dropdown spelling.',
          resolvedFromActiveList: employerMap ? employerMap.size : 0,
          overriddenByHand: Object.keys(employerOverrides).length,
          overrides: Object.keys(employerOverrides).map(cid => ({ contractorId: cid, employer: employerOverrides[cid] })),
          unresolved: unresolved.size },
        { kind: 'gender', note: '"Is Male" = true → MALE, false → FEMALE, blank → blank.' },
        { kind: 'date normalization', note: 'Birth Date and Hired Date run through normDate().' },
        { kind: 'dropdown snapping',
          note: 'Every dropdown-backed column is snapped to the template\'s spelling (country aliases like USA→US included); unmatched values are kept visible.' }
      ],
      fills: Object.keys(columnFills).map(i => ({
        column: H[+i], value: columnFills[i].val,
        scope: columnFills[i].mode === 'blank' ? 'blank cells only' : 'all rows', kind: 'column fill' })),
      edits: { cells: Object.keys(cellEdits).slice(0, 500).map(k => { const [si, ci] = k.split('|'); return { legacyRow: +si + 2, column: H[+ci], value: cellEdits[k] }; }),
        smartFixes: Object.keys(smartFixMap).map(k => { const i = k.indexOf('||'); return { column: H[+k.slice(0, i)], from: k.slice(i + 2), to: smartFixMap[k] }; }),
        note: 'legacyRow = the row number in the Legacy employee CSV (header = row 1).' },
      dropped: {
        removedByHand: { count: removedRows.size, rowIndexes: [...removedRows].slice(0, 500) },
        skippedInactive: { count: skippedInactive, reason: 'Is Active was not literally "true".' },
        alreadyInPickTrace: { count: alreadyMigratedCount, reason: 'Matched in the 3.0 export on Alt ID or name + birth date.', rows: skippedExisting.slice(0, 500) },
        archivedContractor: {
          count: archivedRemoved,
          restoredByHand: [...restoredArchived],
          reason: 'Contractor is archived and was not restored.' }
      },
      asked: asked,
      output: out,
      provenance: prov
    };
  }

  // ─── Export: this engine's files for "Export all" (legacy-page.js) ───
  // Nothing blocks. Flagged cells are highlighted; the column's header note
  // says what it needs, and a cell's own question ("Which employer is Legacy
  // contractor #3?") sits on that cell. Questions about a whole column (the
  // {code} in names, missing Alt IDs) are header notes. Things only we can do
  // (create an employer) go to `team`.
  function askFor(ci, flag) {
    const h = EMP_HEADERS[ci].replace(/\*$/, '');
    if (ci === 4 && flag === 'req') return { title: 'Tell us each person’s employer',
      detail: 'Legacy only kept a contractor number. Pick the employer for each person — each cell’s note says which Legacy contractor it was.' };
    if (ci === 4) return { title: 'Check these employer names', detail: 'They aren’t set up in 3.0 yet — make sure each is spelled the way you want it.' };
    if (ci === 2 && flag === 'off') return { title: 'Check possible duplicates', detail: 'These people may be listed twice (or already be in 3.0). Same person, or someone else?' };
    return flag === 'req' ? { title: 'Fill in ' + h, detail: '' }
      : { title: 'Pick a valid ' + h, detail: 'These values aren’t options in 3.0 — pick one from the list below.' };
  }
  function noteFor(ri, ci, flag) {
    const r = formattedRows[ri];
    if (ci === 4 && flag === 'req') {
      const cid = srcCid(srcKept[ri]);
      const raw = (employerMap && employerMap.get(cid)) || (fullContractorMap && fullContractorMap.get(cid)) || '';
      return 'Legacy contractor #' + cid + (raw ? ' (“' + raw + '”)' : '');
    }
    if (ci === 2 && flag === 'off') { const d = dupPossible.find(x => x.ri === ri); return d ? d.reason + (d.existing ? ': ' + d.existing : '') : ''; }
    if (flag === 'off') return '“' + r[ci] + '” isn’t an option in 3.0.';
    return '';
  }

  // ─── Questions for the client (legacy-form.js) — and their answers back ───
  // People are keyed by their row in the Legacy employee CSV, contractors by
  // their Legacy Contractor ID, so answers land on the same people when the
  // same Legacy file is loaded again.
  const people = (n) => n + ' ' + (n === 1 ? 'person' : 'people');
  function questions() {
    optCache = null;
    const Q = [];
    if (!formattedRows) return Q;
    const kept = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    const who = r => [stripCode(r[0]), stripCode(r[2])].filter(Boolean).join(' ');
    // 1. Employers. Contractor numbers with no name at all → "who is it?".
    //    Employers that aren't in 3.0 yet → a card per employer with what
    //    3.0's "Create New Employer" asks for.
    const un = getUnresolved();
    const nameless = [...un.entries()].filter(([cid, info]) => !info.raw && !employerOverrides[cid]);
    if (nameless.length) Q.push({ id: 'em.employer', type: 'pick', other: true, where: 'Employees · ' + people(nameless.reduce((n, [, x]) => n + x.count, 0)),
      title: 'Who is the employer for each of these groups?',
      help: 'Legacy only kept a contractor number for each person. Pick their employer in 3.0 — or type the name if it’s a new one.',
      options: (tplData && tplData.employerList) || [],
      items: nameless.sort((a, b) => b[1].count - a[1].count).map(([cid, info]) => ({ key: cid,
        label: 'Legacy contractor #' + cid, sub: people(info.count) + ', e.g. ' + stripCode(info.sampleName || '') })) });
    const toMake = new Map();   // cid → { name, count, samples }
    kept.forEach(ri => {
      const r = formattedRows[ri];
      if (!r[4] || employerValid(r[4])) return;
      const cid = srcCid(srcKept[ri]);
      if (employerDetails[cid]) return;          // the client already told us about this one
      const e = toMake.get(cid) || toMake.set(cid, { name: r[4], count: 0, samples: [] }).get(cid);
      e.count++;
      if (e.samples.length < 3) e.samples.push(who(r));
    });
    if (toMake.size) {
      Q.push({ id: 'em.newemployers', type: 'cards', where: 'Employers · ' + toMake.size + ' new',
        title: 'Tell us about ' + (toMake.size === 1 ? 'this employer' : 'these employers'),
        help: (toMake.size === 1 ? 'It isn’t' : 'They aren’t') + ' set up in PickTrace 3.0 yet, so we need ' + (toMake.size === 1 ? 'its' : 'their') + ' details. ' +
          (employerIdsAssumed ? 'We matched Legacy’s contractor numbers to names by the order they were created — if the people listed don’t work for that employer, fix the name. ' : '') +
          'Each employer has its own tab. Everything marked Required is needed to create them in 3.0.',
        // Laid out like 3.0's Create New Employer: who they are, contact, payroll.
        // Address parts sit in section 'addr' — not drawn as fields (the address
        // box fills them) but they count toward "required left".
        sections: [{ key: 'who', label: 'Who they are', hint: 'Their name and what kind of employer they are' },
          { key: 'contact', label: 'Contact', hint: 'Paste the address and it fills the boxes' },
          { key: 'pay', label: 'Payroll', hint: 'When their pay cycle starts and how many days it runs' },
          { key: 'more', label: 'Optional', hint: 'Only if you use them' }],
        address: 'contact',
        fields: [
          { key: 'name', label: 'Employer name', required: true, section: 'who' },
          { key: 'type', label: 'Employer type', type: 'buttons', options: ['Regular', 'Internal contractor', 'Contractor'], required: true, section: 'who', full: true },
          { key: 'flc', label: 'FLC organization', required: true, section: 'who', showIf: { field: 'type', values: ['Contractor'] } },
          { key: 'ein', label: 'Employer Identification Number (EIN)', section: 'who' },
          { key: 'phone', label: 'Phone number', type: 'tel', required: true, section: 'contact', placeholder: '(555) 555-0100' },
          { key: 'email', label: 'Email', type: 'email', required: true, section: 'contact' },
          { key: 'paystart', label: 'Pay cycle start date', type: 'date', required: true, section: 'pay' },
          { key: 'payDays', label: 'Pay cycle length (days)', type: 'number', required: true, section: 'pay', placeholder: '7' },
          { key: 'commission', label: 'Commission %', type: 'number', section: 'more', showIf: { field: 'type', values: ['Internal contractor', 'Contractor'] } },
          { key: 'altid', label: 'Employer Alt ID', section: 'more' },
          { key: 'address1', label: 'Street', required: true, section: 'addr' }, { key: 'city', label: 'City', required: true, section: 'addr' },
          { key: 'state', label: 'State', required: true, section: 'addr' }, { key: 'zip', label: 'Zip', required: true, section: 'addr' },
          { key: 'country', label: 'Country', required: true, section: 'addr' }
        ],
        items: [...toMake.entries()].sort((a, b) => b[1].count - a[1].count).map(([cid, e]) => {
          const d = employerDetails[cid] || {}, ci = contractorInfo[cid] || {};
          return { key: cid, label: e.name, sub: 'Legacy contractor #' + cid + ' · ' + people(e.count) + ', e.g. ' + e.samples.join(', '),
            values: Object.assign({ name: e.name, country: 'US', payDays: '7', phone: ci.phone || '', email: ci.email || '' }, d) };
        }) });
    }
    // 2. Other required columns that are empty, one question per column.
    EMP_HEADERS.forEach((h, ci) => {
      if (!/\*$/.test(h) || ci === 4) return;
      const miss = kept.filter(ri => !formattedRows[ri][ci]);
      if (!miss.length) return;
      Q.push({ id: 'em.req.' + ci, type: 'pick', other: true, where: 'Employees · ' + people(miss.length),
        title: 'What is the ' + h.replace(/\*$/, '').toLowerCase() + ' for ' + (miss.length === 1 ? 'this person' : 'these people') + '?',
        help: 'Legacy didn’t have it, and 3.0 needs it.', placeholder: ci === 3 ? 'YYYY-MM-DD' : '',
        options: colOptions(h), items: miss.map(ri => ({ key: String(srcIdxKept[ri]), label: who(formattedRows[ri]) || '(no name)' })) });
    });
    // 3. Possible duplicates.
    const dups = dupPossible.filter(d => !removedRows.has(d.ri));
    if (dups.length) Q.push({ id: 'em.dups', type: 'pick', where: 'Employees · ' + people(dups.length),
      title: 'Are these the same person as someone else?',
      help: 'They match another person on SSN or name. Tell us if each one is a duplicate.',
      extra: [{ value: 'same', label: 'Same person — leave this one out' }, { value: 'different', label: 'Different person — keep both' }],
      items: dups.map(d => ({ key: String(srcIdxKept[d.ri]), label: d.name, sub: d.reason + (d.existing ? ' — ' + d.existing : '') })) });
    // 4. Values that aren't options in 3.0 — one question per column.
    const byCol = new Map();
    computeSmartFixes().forEach(f => (byCol.get(f.colIdx) || byCol.set(f.colIdx, []).get(f.colIdx)).push(f));
    byCol.forEach((list, ci) => Q.push({ id: 'em.fix.' + ci, type: 'pick', where: 'Employees · ' + list[0].header.replace(/\*$/, ''),
      title: 'Which ' + list[0].header.replace(/\*$/, '') + ' do you mean?',
      help: 'These values from Legacy aren’t options in 3.0. Pick the one each should be.',
      options: list[0].options.slice().sort(),
      items: list.map(f => ({ key: String(f.from).toUpperCase().trim(), label: '“' + f.from + '”', sub: people(f.count), value: f.suggestion || '' })) }));
    // 5–6. Names and Alt IDs, shown on real employees so it's clear which
    //      number is which: Legacy writes its own ID into the last name, and
    //      a few people also have a separate Alt ID.
    const li = eIdx('last name'), fi = eIdx('first name'), ii = eIdx('id'), ai = eIdx('alt id');
    const gs = (r, i) => i < 0 ? '' : String(r[i] == null ? '' : r[i]).trim();
    const coded = srcKept.filter((r, k) => !removedRows.has(k) && NAME_CODE.test(gs(r, li)));
    const withAlt = srcKept.filter((r, k) => !removedRows.has(k) && gs(r, ai));
    const samples = [].concat(withAlt.filter(r => NAME_CODE.test(gs(r, li))).slice(0, 1), coded.filter(r => !gs(r, ai)).slice(0, 2), withAlt.slice(0, 1))
      .filter((r, k, arr) => arr.indexOf(r) === k).slice(0, 3);
    const exRows = samples.map(r => [gs(r, fi), gs(r, li).replace(/(\{\d+\})$/, '`$1`'), gs(r, ii) || '—', gs(r, ai) || '— none —']);
    const examples = { columns: ['First name', 'Last name in Legacy', 'Legacy ID', 'Alt ID in Legacy'], rows: exRows };
    // One question, Alt ID first (it's the number payroll gets), then whether
    // the Legacy number still belongs in the last name. The example table gets
    // a column per part showing what the pick does to those same people.
    const parts = [];
    if (kept.length && altIdMode == null) {
      const nAlt = kept.filter(ri => gs(srcKept[ri], ai)).length;
      // Two different people: one with only a Legacy ID, one who also has an Alt ID.
      const sL = samples.find(r => !gs(r, ai)) || srcKept[kept[0]], sA = withAlt.find(r => r !== sL);
      parts.push({ key: 'altid', column: 'Alt ID in 3.0',
        title: 'Which number should be each employee’s Alt ID?',
        help: 'Only ' + people(nAlt) + ' have an Alt ID in Legacy; everyone has a Legacy ID.',
        choices: [{ value: 'legacy', label: 'Their Legacy ID', note: 'e.g. ' + gs(sL, fi) + ' → ' + (gs(sL, ii) || '—'), cells: samples.map(r => gs(r, ii)) },
          { value: 'altOrLegacy', label: 'Their Alt ID from Legacy when they have one, otherwise their Legacy ID',
            note: sA ? 'e.g. ' + gs(sA, fi) + ' → ' + gs(sA, ai) + ', and ' + gs(sL, fi) + ' → ' + (gs(sL, ii) || '—') : '', cells: samples.map(r => gs(r, ai) || gs(r, ii)) },
          { value: 'alt', label: 'Only the Alt ID from Legacy', note: 'The other ' + people(kept.length - nAlt) + ' get no Alt ID', cells: samples.map(r => gs(r, ai)) },
          { value: 'none', label: 'No Alt ID for anyone', note: 'Leave it empty — you can add them in 3.0 later', cells: samples.map(() => '') }] });
    }
    if (keepNameCode && coded.length && !codeAnswered) {
      const s0 = samples.find(r => NAME_CODE.test(gs(r, li))) || coded[0];
      parts.push({ key: 'code', column: 'Last name in 3.0',
        title: 'Take the Legacy number out of their last names?',
        help: 'In Legacy, ' + people(coded.length) + ' have their Legacy ID added to the end of their last name. In 3.0 the number travels to payroll in the Alt ID, so it isn’t needed in the name.',
        choices: [{ value: 'remove', label: 'Take the number out', note: 'Last name becomes “' + stripCode(gs(s0, li)) + '”', cells: samples.map(r => stripCode(gs(r, li))) },
          { value: 'keep', label: 'Keep it as it is', note: 'Last name stays “' + gs(s0, li) + '”', cells: samples.map(r => gs(r, li)) }] });
    }
    if (parts.length) {
      Q.push({ id: 'em.ids', type: 'choices', where: 'Employees · ' + people(kept.length),
        title: parts.length > 1 ? 'Employee Alt IDs and last names' : parts[0].title,
        help: 'In PickTrace 3.0 the Alt ID is what goes out to your payroll in the payroll export, so it should be the number your payroll system knows each person by. ' +
          'Here are a few of your employees as they are in Legacy — the blue columns show what they’ll be in 3.0 with your picks:',
        examples, parts });
    }
    return Q;
  }

  function applyAnswers(A) {
    const out = [];
    let n = 0;
    Object.entries(A['em.employer'] || {}).forEach(([cid, v]) => { if (v) { employerOverrides[cid] = v; n++; } });
    // New employers' details. A changed name renames that contractor's people's Employer.
    let det = 0;
    // "none", "NA", "-" typed into a box is the client saying there isn't one:
    // the box is left empty in "Employers to create" rather than written out.
    const NO_VALUE = /^(n\/?a|none|null|nil|unknown|tbd|-+|—|\?+|x)$/i;
    let blanked = 0;
    Object.entries(A['em.newemployers'] || {}).forEach(([cid, d]) => {
      if (!d) return;
      d = Object.assign({}, d);
      Object.keys(d).forEach(k => { if (typeof d[k] === 'string' && NO_VALUE.test(d[k].trim())) { d[k] = ''; blanked++; } });
      employerDetails[cid] = Object.assign({}, employerDetails[cid] || {}, d);
      if (d.name && d.name.trim()) employerOverrides[cid] = d.name.trim();
      det++;
    });
    if (det) out.push('New employers: details for ' + det + (blanked ? ' (' + blanked + ' “none” / “NA” left empty)' : ''));
    if (n) out.push('Employers: ' + n + ' contractor' + (n === 1 ? '' : 's') + ' matched');
    Object.keys(A).filter(k => k.indexOf('em.req.') === 0).forEach(k => {
      const ci = +k.slice(7); let c = 0;
      Object.entries(A[k] || {}).forEach(([si, v]) => { if (v) { cellEdits[si + '|' + ci] = v; c++; } });
      if (c) out.push(EMP_HEADERS[ci].replace(/\*$/, '') + ': ' + c + ' filled in');
    });
    let same = 0;
    Object.entries(A['em.dups'] || {}).forEach(([si, v]) => { if (v === 'same') { removedSrc.add(+si); same++; } });
    if (A['em.dups']) out.push('Possible duplicates: ' + same + ' left out');
    Object.keys(A).filter(k => k.indexOf('em.fix.') === 0).forEach(k => {
      const ci = k.slice(7); let c = 0;
      Object.entries(A[k] || {}).forEach(([u, v]) => { if (v) { smartFixMap[ci + '||' + u] = v; c++; } });
      if (c) out.push('Fixed ' + c + ' value' + (c === 1 ? '' : 's') + ' not in a dropdown');
    });
    // em.ids = { altid, code } (one question); em.code / em.altid from older pages.
    const ids = A['em.ids'] || {};
    const code = ids.code ? { choice: ids.code } : A['em.code'];
    if (code && code.choice) { keepNameCode = code.choice === 'keep'; codeAnswered = true; out.push('Legacy number in last names: ' + (keepNameCode ? 'kept' : 'taken out')); }
    const alt = ids.altid ? { choice: ids.altid } : A['em.altid'];
    if (alt && typeof alt === 'object' && alt.choice) {
      altIdMode = alt.choice;
      out.push('Alt ID: ' + ({ legacy: 'their Legacy ID', altOrLegacy: 'Legacy Alt ID, else Legacy ID', alt: 'Legacy Alt ID only', none: 'left empty' })[altIdMode]);
    } else if (alt) out.push('Alt ID — the client says: “' + String(alt).trim() + '”');
    rebuildAndRender();
    return out;
  }

  async function exportParts(limit) {
    optCache = null;
    const kept = formattedRows.map((_, i) => i).filter(i => !removedRows.has(i));
    const origName = (tplData.fileName || '3.0_Employee_bulk_create.xlsx').trim();
    const dot = origName.lastIndexOf('.');
    const base = dot > 0 ? origName.substring(0, dot) : origName;
    const nameOf = (chunk, total) => base + ' — filled' + (total > 1 ? ' (' + (chunk + 1) + ' of ' + total + ')' : '') + '.xlsx';
    const dups = new Set(dupPossible.filter(d => !removedRows.has(d.ri)).map(d => d.ri));
    const who = r => [stripCode(r[0]), stripCode(r[2])].filter(Boolean).join(' ');
    // Whole-column questions: no highlight, a header note and a read-me line.
    const headerOnly = [];
    if (keepNameCode && namesWithCode()) {
      const li = eIdx('last name');
      const sample = srcKept.map(r => String(li < 0 ? '' : r[li] || '')).find(v => NAME_CODE.test(v)) || '';
      headerOnly.push({ ci: 2, count: namesWithCode(), title: 'Keep the Legacy number in last names?',
        detail: 'Last names end in the person’s Legacy ID, e.g. “' + sample + '”. Should we keep the number, or take it out?' });
    }
    const noAlt = kept.filter(ri => !formattedRows[ri][5]).length;
    if (noAlt) headerOnly.push({ ci: 5, count: noAlt, title: 'Do these people need an Alt ID?',
      detail: 'They have no Alt ID in Legacy. The Alt ID is what 3.0’s payroll export uses — should they get one, and what should it be?' });
    const a = IMLegacyDrop.annotate({
      label: 'Employees', H: EMP_HEADERS, total: kept.length, limit, file: nameOf,
      unit: ['person', 'people'],
      options: ci => colOptions(EMP_HEADERS[ci]),
      cells: pos => {
        const ri = kept[pos], out = [];
        for (let ci = 0; ci < EMP_HEADERS.length; ci++) {
          const flag = cellFlag(ri, ci);
          if (flag) out.push({ ci, flag, note: noteFor(ri, ci, flag) });
        }
        if (dups.has(ri) && !out.some(c => c.ci === 2)) out.push({ ci: 2, flag: 'off', note: noteFor(ri, 2, 'off') });
        return out;
      },
      nameOf: pos => who(formattedRows[kept[pos]]),
      ask: askFor,
      headerOnly
    });
    const chunks = chunkArr(kept.map(ri => formattedRows[ri].slice()), limit);
    const files = [];
    // Each file patches the template's own zip (IMXlsxTemplate): only the
    // DATA ENTRY rows change, so every dropdown, style and the header row ship
    // as PickTrace made them; inline strings keep an employer's trailing space.
    for (let i = 0; i < chunks.length; i++) {
      const blob = await IMXlsxTemplate.build({ rawBuffer: tplData.rawBuffer, headers: EMP_HEADERS, rows: chunks[i],
        marks: a.chunks[i].marks, notes: a.chunks[i].notes, noteAuthor: 'PickTrace' });
      files.push({ label: 'Employees', name: nameOf(i, chunks.length), data: new Uint8Array(await blob.arrayBuffer()) });
    }
    const team = [];
    const toMake = pendingEmployersToCreate().map(([n]) => n);
    if (toMake.length) {
      team.push('Create employer' + (toMake.length === 1 ? '' : 's') + ' in 3.0 (details in “Employers to create”): ' + IMLegacyDrop.examples(toMake, 10));
      // One row per employer to create, with whatever the client has told us —
      // laid out like 3.0's "Create New Employer" screen, top to bottom.
      const H = ['Employer Name', 'Employer Type', 'FLC Organization', 'Employer Identification Number', 'Address 1', 'Address 2', 'City', 'State',
        'Zip Code', 'Country', 'Phone Number', 'Email', 'Pay Cycle Start Date', 'Pay Cycle Duration (days)', 'Commission %', 'Employer Alt ID',
        'Legacy Contractor IDs', 'Employees'];
      const rows = pendingEmployersToCreate().map(([name, info]) => {
        const cids = [...info.cids].sort((x, y) => x - y);
        const d = cids.map(c => employerDetails[c]).find(Boolean) || {};
        const ci = cids.map(c => contractorInfo[c]).find(Boolean) || {};
        return [d.name || name, d.type || '', d.type === 'Contractor' ? d.flc || '' : '', d.ein || '', d.address1 || '', d.address2 || '', d.city || '',
          d.state || '', d.zip || '', d.country || '', d.phone || ci.phone || '', d.email || ci.email || '', d.paystart || '', d.payDays || '',
          d.commission || '', d.altid || '', cids.join(', '), info.count];
      });
      const ws = XLSX.utils.aoa_to_sheet([H].concat(rows));
      ws['!cols'] = H.map(h => ({ wch: Math.max(12, Math.min(34, h.length + 4)) }));
      const wbOut = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbOut, ws, 'Employers to create');
      files.unshift({ label: 'Employers to create', name: 'Employers to create in 3.0.xlsx', data: new Uint8Array(XLSX.write(wbOut, { type: 'array', bookType: 'xlsx' })) });
    }
    logEvent('export', 'Prepared ' + kept.length + ' employees in ' + files.length + ' file' + (files.length === 1 ? '' : 's') + '.');
    return { files, items: a.items, team };
  }

  // ─── Reset ───
  function reset() {
    logEvent('reset', 'Reset clicked — files and choices cleared.');
    empData = null; employerMap = null; fullContractorMap = null;
    migratedAltIds = null; alreadyMigratedCount = 0; tplData = null;
    formattedRows = null; srcKept = null; srcIdxKept = null; employerOverrides = {};
    employerIdsAssumed = false; contractorInfo = {}; employerDetails = {};
    columnFills = {}; smartFixMap = {}; cellEdits = {};
    removedRows = new Set(); removedSrc = new Set();
    existingIdx = null; skippedExisting = []; dupPossible = [];
    keepNameCode = true; codeAnswered = false; altIdMode = null; pvMode = 'all'; pvColsOnly = false;
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
    ['em-section-archived', 'em-section-create', 'em-section-unresolved', 'em-section-bulk', 'em-section-preview',
     'em-section-smartfix', 'em-section-possible', 'em-section-existing', 'em-section-names'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    $('em-summary').style.display = 'none';
    if (window.IMDebug) IMDebug.refresh('employee-migrate');
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
    cellEdits[srcIdxKept[ri] + '|' + ci] = val;   // kept through rebuilds
    // Repaint just this cell — redrawing the grid here would throw away the
    // cell you just tabbed into.
    const fl = cellFlag(ri, ci), note = changeOf(ri, ci);
    td.setAttribute('style', FLAG_STYLE[fl] || '');
    td.classList.toggle('is-chg', !!note);
    td.title = [FLAG_TIP[fl], note].filter(Boolean).join(' · ');
    findPossibleDuplicates();
    renderUnresolved();
    renderSmartFixes();
    renderPossible();
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
    { const batch = $('lm-batch'); if (batch) batch.addEventListener('input', () => { if (formattedRows) updateSummary(); }); }
    // Build, Export all, Reset and Debug Dump belong to the whole page (legacy-page.js).
    if (window.IMDebug) {
      IMDebug.register('employee-migrate', {
        label: 'Employee Migrate (Legacy → 3.0)',
        // Any file loaded, or anything logged, is enough to dump: the point is
        // to see what went wrong even when the build never ran.
        ready: () => !!(empData || employerMap || fullContractorMap || migratedAltIds ||
          tplData || events.length),
        collect: collectDebug
      });
    }
    registerPage();
    const sfAll = $('em-smartfix-apply-all');
    if (sfAll) sfAll.addEventListener('click', applyAllSmartFixes);
    // The keep-the-code switch is redrawn with its panel; listen on the panel.
    const names = $('em-section-names');
    if (names) names.addEventListener('change', e => {
      if (e.target.id !== 'em-keep-code') return;
      keepNameCode = e.target.checked;
      logEvent('setting', 'Legacy code in names: ' + (keepNameCode ? 'kept' : 'taken out'));
      if (formattedRows) rebuildAndRender();
    });
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
