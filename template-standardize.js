// ═══════════════════════════════════════════════════════════════════════
// Template Standardize — formats source rows (Excel/CSV or Org Data) into
// a personalized PickTrace bulk Create or Update template. Auto-maps
// columns by name; surfaces Sites and Crop & Variety values that aren't
// in the template's dropdown so the user can create them in PickTrace
// before bulk uploading. Required (*) cells must all be filled; the
// "Empty Required Columns" picklist drives single-click bulk-fills.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');

  // Today as YYYY-MM-DD (PickTrace's bulk-template format).
  function todayYMD() {
    const d = new Date();
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  }
  // True if the given string parses to a date strictly BEFORE today (local time).
  // Unparseable strings return false (we don't penalize unrecognized formats).
  // Critically, parses YYYY-MM-DD and MM/DD/YYYY strings as LOCAL midnight —
  // `new Date("2026-05-05")` interprets ISO-style strings as UTC, which in any
  // timezone west of UTC becomes "yesterday local" and gets flagged as past.
  function isDateBeforeToday(s) {
    if (!s) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const str = String(s).trim();
    let d = null;
    let m;
    if ((m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) {
      d = new Date(+m[1], +m[2] - 1, +m[3]);
    } else if ((m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) {
      d = new Date(+m[3], +m[1] - 1, +m[2]);
    } else {
      d = new Date(str);
    }
    if (!d || isNaN(d.getTime())) return false;
    d.setHours(0, 0, 0, 0);
    return d.getTime() < today.getTime();
  }

  // Bulk template column schemas (matches PickTrace's templates exactly).
  const CREATE_HEADERS = [
    'Site*','Name*','Alt ID','Location Type*','Crop & Variety*','Planted At','Acreage','Length',
    'Plant Count','Start Date*','Location Group','Lot Number','Rootstock','Plant/Seed Seller',
    'Plant/Seed Producer','Clone/Subvariety','Growing Manager','Production Status','Training Style',
    'Trellis Type','Mulch Type','Row Direction','Organic Status','Organic Certifier','Wet Date',
    'Germination Date','Planted Date','Grafting Date','Production Start','Organic Certification Date',
    'Stand Count','Row/Bed Count','Post Count','Percent Covered','Row Spacing, in.','Plant Spacing, in.',
    'Post Spacing, in.','Bed Width, in.','Custom Data 1','Custom Data 2','Custom Data 3'
  ];
  const UPDATE_HEADERS = CREATE_HEADERS.concat(['Is Archived']);

  // ─── State ───
  let srcData = null;          // { headers, rows, fileName, sheetName }
  let tplData = null;          // { headers, rows, dropdowns, fileName, rawBuffer }
  let target = 'create';
  let mapping = {};            // bulkColIdx → srcColIdx (-1 = unmapped)
  let formattedRows = null;    // [[...]] mapped + filled rows (canonical state)
  let manualFills = {};        // bulkColIdx → user-supplied value (re-applied after rebuilds)
  let removedRows = new Set(); // formattedRows indexes excluded from preview + export
  let siteAddresses = {};      // site value → address string (reference only; survives re-renders)
  let dateFill = { mm: '', dd: '', range: 'first' }; // year-only → full-date expansion
  let existingData = null;        // { keys: Set<siteName>, fileName, count } — locations already in PickTrace
  let existingRowIdxs = new Set();// formattedRows indexes that already exist in PickTrace (dropped)
  let cropSplit = true;           // split " - Subvariety" out of Crop & Variety → Clone/Subvariety
  let nameOverrides = {};         // formattedRows index → renamed Name* (resolves collisions, sticky)
  let cellOverrides = {};         // "rowIdx|colIdx" → value — sticky per-cell edits from the editable preview (all columns EXCEPT Name*, which uses nameOverrides)
  let smartFixMap = {};           // "colIdx||UPPERVALUE" → canonical dropdown value — sticky bulk fixes the user applied from the Smart Fixes panel
  let previewIssuesOnly = false;  // preview toggle: show only rows with an errored cell
  let collisionKeys = new Set();  // locKeys flagged as collisions (set in rebuildFormattedRows)
  let existingTakenKeys = new Set(); // locKeys whose name is already used by a DIFFERENT block in PickTrace
  let initialized = false;

  // A row is excluded from preview + export if the user removed it OR it already
  // exists in PickTrace (cross-referenced from the uploaded existing-locations file).
  function excluded(i) { return removedRows.has(i) || existingRowIdxs.has(i); }

  // Location identity key — case/whitespace-insensitive Site + Name (block).
  const locNorm = s => String(s == null ? '' : s).toUpperCase().trim().replace(/\s+/g, ' ');
  function locKey(site, name) { return locNorm(site) + '||' + locNorm(name); }
  function findColIdxByNames(headers, names) {
    const noStar = h => String(h || '').toLowerCase().replace(/\*+$/, '').trim();
    for (const want of names) {
      const i = headers.findIndex(h => noStar(h) === want);
      if (i >= 0) return i;
    }
    return -1;
  }

  // Historical date columns whose cells may carry only a year / year range.
  // Start Date* is intentionally excluded (it's the required activation date).
  const DATE_FILL_HEADERS = ['Wet Date', 'Germination Date', 'Planted Date',
    'Grafting Date', 'Production Start', 'Organic Certification Date'];
  function dateFillIdxs() {
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    return DATE_FILL_HEADERS.map(h => bulkHeaders.indexOf(h)).filter(i => i >= 0);
  }

  // The full set of bulk columns that hold dates (Start Date* + the historical
  // date columns). Used to coerce Excel date serials read from the source.
  function dateColIdxs() {
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const names = ['Start Date*'].concat(DATE_FILL_HEADERS);
    return names.map(h => bulkHeaders.indexOf(h)).filter(i => i >= 0);
  }

  // Excel stores dates as serial day-counts; when a source column is date-
  // formatted, XLSX hands us the raw serial (e.g. 46213) instead of a date, so
  // the cell would otherwise show "46213". Convert a 5–7-digit serial in a date
  // column to 'YYYY-MM-DD'. Excel's epoch is 1900-01-01 but it wrongly counts
  // 1900 as a leap year, so anchoring at 1899-12-30 (UTC, to avoid TZ drift)
  // accounts for the phantom 1900-02-29. Non-serial values (real date strings,
  // 4-digit years, blanks, text) are returned untouched.
  function excelSerialToYMD(serial) {
    const n = Math.floor(Number(serial));
    if (!isFinite(n)) return null;
    const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    if (isNaN(d.getTime())) return null;
    return d.getUTCFullYear() + '-' +
      String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
      String(d.getUTCDate()).padStart(2, '0');
  }
  function coerceExcelDate(v) {
    const s = String(v == null ? '' : v).trim();
    // Only 5–7 digit integers → leaves 4-digit years, year ranges (have a dash),
    // real date strings, and text alone. 10000 ≈ 1927, 2958465 = 9999-12-31.
    if (!/^\d{5,7}$/.test(s)) return s;
    const n = parseInt(s, 10);
    if (n < 10000 || n > 2958465) return s;
    return excelSerialToYMD(n) || s;
  }

  function activeSchemaHeaders() { return target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS; }

  // ─── Smart-match helpers (Smart Fixes panel) ───
  // Collapse to alphanumerics only — catches spacing/punctuation/case diffs
  // ("Crop & Variety" ↔ "Crop and Variety", "Almond -Other" ↔ "Almond-Other").
  function looseKey(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
  // Token-wise singular/plural key — depluralizes each word (>3 chars ending in
  // 's') so "Almonds-Other" and "Almond-Other" collapse to the same key. The
  // same normalization is applied to both sides, so exact linguistic accuracy
  // doesn't matter — only that the two sides agree.
  function pluralKey(s) {
    return String(s == null ? '' : s).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean)
      .map(w => (w.length > 3 && w.endsWith('s')) ? w.slice(0, -1) : w).join('|');
  }
  // Build { canon:Map<UPPER,val>, loose:Map<key,[vals]>, plural:Map<key,[vals]> }
  // for a dropdown value list.
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
  // Given a value already known to be off-list (not a case-insensitive exact
  // dropdown hit), return the single confident canonical match, or null when
  // there's no match or it's ambiguous (2+ candidates).
  function confidentMatch(value, maps) {
    const uniq = m => (m && m.length === 1) ? m[0] : null;
    const byLoose = uniq(maps.loose.get(looseKey(value)));
    if (byLoose) return byLoose;
    const byPlural = uniq(maps.plural.get(pluralKey(value)));
    if (byPlural) return byPlural;
    return null;
  }
  // Bulk columns that carry a template dropdown (indices into the active schema).
  function dropdownColIdxs() {
    const bulkHeaders = activeSchemaHeaders();
    return bulkHeaders.map((h, i) => dropdownValuesFor(h) ? i : -1).filter(i => i >= 0);
  }
  // Scan every dropdown-backed column for off-list values. Each becomes one row
  // in the Smart Fixes panel with a dropdown to choose the target value —
  // pre-selected to the confident match when there is one (case/spacing/plural),
  // left as "— pick —" when there isn't (e.g. "EQUIPMENT" has no auto match, so
  // the user chooses a valid Location Type). Returns
  // [{ colIdx, header, from, count, options, suggestion }], suggested first.
  function computeSmartFixes() {
    if (!formattedRows || !tplData) return [];
    const bulkHeaders = activeSchemaHeaders();
    const out = [];
    dropdownColIdxs().forEach(ci => {
      const vals = dropdownValuesFor(bulkHeaders[ci]);
      const maps = buildMatchMaps(vals);
      const seen = new Map(); // UPPERVALUE → { from, count, options, suggestion }
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        const val = r[ci];
        if (!val) return;
        const u = String(val).toUpperCase().trim();
        if (maps.canon.has(u)) return;                 // already an exact (case-insensitive) hit
        if (smartFixMap[ci + '||' + u]) return;         // already applied
        const sug = confidentMatch(val, maps);
        const suggestion = (sug && String(sug).toUpperCase().trim() !== u) ? sug : '';
        const e = seen.get(u) || { colIdx: ci, header: bulkHeaders[ci], from: val, count: 0, options: vals, suggestion };
        e.count++;
        seen.set(u, e);
      });
      seen.forEach(e => out.push(e));
    });
    // Confident suggestions first, then most-rows-first.
    out.sort((a, b) => (b.suggestion ? 1 : 0) - (a.suggestion ? 1 : 0) || b.count - a.count);
    return out;
  }

  // Template dropdown values for a bulk column, or null if that column carries
  // no list. tplData.dropdowns is keyed by norm(header) with the trailing '*'
  // absent (the DROP-DOWN INPUTS sheet labels columns "Site", "Crop & Variety",
  // "Location Type", "Production Status", …).
  function dropdownValuesFor(bulkHeader) {
    if (!tplData || !tplData.dropdowns) return null;
    const key = norm(bulkHeader).replace(/\*+$/, '').trim();
    const set = tplData.dropdowns.get(key);
    return set && set.size ? [...set] : null;
  }

  // ─── Mode-pill switcher ───
  function attachModeSwitcher() {
    document.querySelectorAll('.cmp-mode-pill').forEach(btn => {
      btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        document.querySelectorAll('.cmp-mode-pill').forEach(b => b.classList.toggle('cmp-mode-active', b === btn));
        $('cmp-mode-compare').style.display = mode === 'compare' ? '' : 'none';
        $('cmp-mode-standardize').style.display = mode === 'standardize' ? '' : 'none';
        // Employee Migration + Standardize/Dedupe now share one top-level tab
        // ("Employees") with an internal sub-toggle (see employee-standardize.js).
        const emp = $('cmp-mode-employees'); if (emp) emp.style.display = mode === 'employees' ? '' : 'none';
      });
    });
  }

  // ─── Header-row detection ───
  // Soto Bros / PickTrace implementation templates have 3-5 rows of metadata
  // (sheet title, "Instructions", long description, "Basics") before the real
  // column header row. Score every candidate row by:
  //   - count of non-empty cells (more = better)
  //   - presence of canonical PickTrace template keywords (heavy bonus)
  //   - short, header-like cell text (small bonus per cell)
  //   - long descriptive sentences (penalty — those are instruction rows)
  // Pick the highest-scoring row in the first ~15 rows.
  const TEMPLATE_KEYWORDS = [
    'site','name','block','location','crop','variety','plant',
    'acreage','acres','archived','date','type','count','group',
    'alt id','lot number','seller','producer','rootstock','planted',
    'organic','training','trellis','mulch','spacing','custom data',
    'bed width','stand count','row spacing','plant count'
  ];
  function detectHeaderRow(aoa) {
    if (!aoa || aoa.length < 2) return 0;
    let bestRow = 0, bestScore = -Infinity;
    const limit = Math.min(aoa.length, 20);
    for (let r = 0; r < limit; r++) {
      const row = aoa[r] || [];
      const nonEmpty = row.filter(c => c != null && String(c).trim() !== '');
      if (nonEmpty.length < 2) continue;
      let score = nonEmpty.length * 2;
      nonEmpty.forEach(c => {
        const s = String(c).trim().toLowerCase();
        if (s.length < 40) score += 1;
        if (s.length > 80) score -= 4;
        TEMPLATE_KEYWORDS.forEach(kw => { if (s.includes(kw)) score += 3; });
      });
      if (score > bestScore) { bestScore = score; bestRow = r; }
    }
    return bestRow;
  }

  function isInstructionRow(row) {
    const nonEmpty = (row || []).filter(c => c != null && String(c).trim() !== '');
    if (!nonEmpty.length) return true;
    const longCells = nonEmpty.filter(c => String(c).length > 80);
    if (longCells.length >= Math.ceil(nonEmpty.length / 2)) return true;
    return nonEmpty.some(c => {
      const s = String(c).trim().toLowerCase();
      return s.startsWith('required.') || s.startsWith('please ') ||
             s.startsWith('this field') || s.startsWith('this tab') ||
             s.startsWith('the unique identifier') || s.startsWith('an alternate id') ||
             s.includes('used to collect') || s.includes('used in reporting');
    });
  }

  // Slice a sheet's AOA into clean { headers, rows } using the auto-detected
  // header row. Trims metadata above and instruction rows below.
  function sliceSheetAoa(aoa) {
    const hIdx = detectHeaderRow(aoa);
    const headers = (aoa[hIdx] || []).map(h => String(h == null ? '' : h).trim());
    const rows = aoa.slice(hIdx + 1)
      .filter(r => !isInstructionRow(r))
      .map(r => r.map(c => c == null ? '' : String(c).trim()));
    return { headers, rows, headerRow: hIdx };
  }

  // ─── Sheet picker for multi-sheet workbooks ───
  // Shows a modal listing every sheet that has at least one row of data.
  // Returns a Promise that resolves with the picked sheet (or rejects on cancel).
  function pickSheet(sheets, fileName) {
    return new Promise((resolve, reject) => {
      // Tear down any existing picker.
      const prior = document.getElementById('ts-sheet-picker');
      if (prior) prior.remove();

      const overlay = document.createElement('div');
      overlay.id = 'ts-sheet-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';

      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner';
      inner.innerHTML =
        '<h3>Select sheet</h3>' +
        '<p>The file <code>' + escHtml(fileName) + '</code> has multiple sheets. Pick the one containing the source data:</p>' +
        '<div class="cmp-org-list" id="ts-sheet-picker-list"></div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="ts-sheet-picker-cancel">Cancel</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);

      const list = inner.querySelector('#ts-sheet-picker-list');
      sheets.forEach((s, i) => {
        const item = document.createElement('div');
        item.className = 'cmp-org-list-item';
        item.innerHTML =
          '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
          '<span class="cmp-org-counts">' + (s.rows ? s.rows.length : 0) + ' rows · ' +
          (s.headers ? s.headers.length : 0) + ' cols</span>';
        item.addEventListener('click', () => {
          overlay.remove();
          resolve(s);
        });
        list.appendChild(item);
      });

      const cancel = () => { overlay.remove(); reject(new Error('cancelled')); };
      inner.querySelector('#ts-sheet-picker-cancel').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
    });
  }

  // ─── File parsing ───
  // Always reads via XLSX directly so we can run the header-row detector on
  // the raw AOA. PickTrace implementation templates (Soto Bros, Wonderful, etc.)
  // have multi-row metadata above the real column header — readUploadedFile
  // would treat row 0 as the header and miss everything.
  function readSrcFile(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          const sheets = wb.SheetNames.map(n => {
            const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' });
            const filtered = aoa.filter(row => row && row.some(c => c != null && String(c).trim() !== ''));
            if (!filtered.length) return null;
            const sliced = sliceSheetAoa(filtered);
            if (!sliced.headers.filter(h => h).length) return null;
            return {
              name: n,
              headers: sliced.headers,
              rows: sliced.rows,
              headerRow: sliced.headerRow
            };
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({
            headers: s.headers, rows: s.rows, fileName: file.name,
            sheetName: s.name, headerRow: s.headerRow
          });
          if (sheets.length === 1) return finish(sheets[0]);
          pickSheet(sheets, file.name).then(finish).catch(reject);
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
  }

  function parseTemplate(buf, fileName) {
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
    const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
    const dropName = wb.SheetNames.find(n => /drop.?down/i.test(n));
    const dataWs = wb.Sheets[dataName];
    if (!dataWs) return null;
    const dataAoa = XLSX.utils.sheet_to_json(dataWs, { header: 1, defval: '' });
    const headers = dataAoa.length ? dataAoa[0].map(h => String(h || '').trim()) : [];
    const dropdowns = new Map();
    if (dropName) {
      const dropWs = wb.Sheets[dropName];
      const raw = XLSX.utils.sheet_to_json(dropWs, { header: 1, defval: '' });
      if (raw.length >= 2) {
        const dh = raw[0].map(h => String(h || '').trim());
        dh.forEach((h, ci) => {
          if (!h) return;
          const k = norm(h);
          const set = dropdowns.get(k) || new Set();
          raw.slice(1).forEach(r => {
            const v = String(r[ci] != null ? r[ci] : '').trim();
            if (v) set.add(v);
          });
          if (set.size) dropdowns.set(k, set);
        });
      }
    }
    return { headers, dropdowns, fileName, rawBuffer: buf, sheetNames: wb.SheetNames };
  }

  // ─── Auto-mapping ───
  // Match each bulk header to a source-data column by:
  //   1. Exact case-insensitive header match
  //   2. Asterisk-stripped match ("Site*" → "site")
  //   3. Common alias map (display_name → Name*, plant_name → Crop & Variety*, etc.)
  const ALIASES = {
    'site*':              ['site','sites','site_name','site name','grower','farm'],
    'name*':              ['name','block','location name','location_name','display_name','display name'],
    'alt id':             ['alt id','altid','code'],
    'location type*':     ['location type','type'],
    'crop & variety*':    ['crop & variety','crop and variety','plant_name','plant name','crop'],
    'planted at':         ['planted at','planted_at'],
    'acreage':            ['acreage','acres','acre','hectares'],
    'length':             ['length','length_meters'],
    'plant count':        ['plant count','plant_count','plant quantity','tree count','trees','plants','quantity'],
    'start date*':        ['start date','start_date'],
    'planted date':       ['planted date','planted_date'],
    'is archived':        ['is archived','is_archived','archived'],
    // Clone/Subvariety must NOT pick up "Variety". Specific aliases ensure
    // Phase 3 catches "Clone" first; STRICT_HEADERS below blocks Phase 4
    // from ever matching Variety against this column via substring.
    'clone/subvariety':   ['clone','subvariety','sub variety','sub_variety','sub-variety']
  };

  // Bulk headers where the substring fallback (Phase 4) is unsafe — a source
  // header could match by accident (e.g., "Variety" matching "Clone/Subvariety"
  // because "clonesubvariety".includes("variety") is true). For these we
  // accept ONLY exact-match + loose-alphanumeric + alias hits.
  const STRICT_HEADERS = new Set(['clone/subvariety']);
  // Loose normalization — strip all non-alphanumeric chars so "Plant/Seed Seller",
  // "Plant Seed Seller", and "plant-seed_seller" all collapse to "plantseedseller".
  const normLoose = h => norm(h).replace(/[^a-z0-9]/g, '');

  function autoMap(bulkHeaders, srcHeaders) {
    const srcIdx = {};         // norm(header) → idx (case+trim+lowercase only)
    const srcIdxLoose = {};    // normLoose(header) → idx (alphanumeric only)
    srcHeaders.forEach((h, i) => {
      const k = norm(h);
      const kLoose = normLoose(h);
      if (k && srcIdx[k] == null) srcIdx[k] = i;
      if (kLoose && srcIdxLoose[kLoose] == null) srcIdxLoose[kLoose] = i;
    });
    const out = {};
    bulkHeaders.forEach((bh, bi) => {
      const bn = norm(bh);
      const bnPlain = bn.replace(/\*$/, '').trim();
      const bnLoose = normLoose(bh);              // includes asterisk-strip via the regex
      const bnLoosePlain = normLoose(bh.replace(/\*$/, ''));
      // 1. Exact match (with or without asterisk).
      if (srcIdx[bn] != null) { out[bi] = srcIdx[bn]; return; }
      if (srcIdx[bnPlain] != null) { out[bi] = srcIdx[bnPlain]; return; }
      // 2. Loose alphanumeric match — catches "Plant/Seed Seller" ↔ "Plant Seed Seller",
      //    "Crop & Variety*" ↔ "Crop and Variety", "Row Spacing, in." ↔ "Row Spacing in".
      if (srcIdxLoose[bnLoose] != null) { out[bi] = srcIdxLoose[bnLoose]; return; }
      if (srcIdxLoose[bnLoosePlain] != null) { out[bi] = srcIdxLoose[bnLoosePlain]; return; }
      // 3. Aliases.
      const aliases = ALIASES[bn] || ALIASES[bnPlain];
      if (aliases) {
        for (const a of aliases) {
          if (srcIdx[a] != null) { out[bi] = srcIdx[a]; return; }
          const aLoose = normLoose(a);
          if (srcIdxLoose[aLoose] != null) { out[bi] = srcIdxLoose[aLoose]; return; }
        }
      }
      // Strict bulk headers don't fall through to substring matching — too
      // many false positives ("variety" matching "clonesubvariety", etc.).
      if (STRICT_HEADERS.has(bn) || STRICT_HEADERS.has(bnPlain)) {
        out[bi] = -1;
        return;
      }
      // 4. Substring fallback on the loose form. Skip *_id columns.
      for (const k of Object.keys(srcIdxLoose)) {
        if (!k || k.endsWith('id') && /id$/.test(k) && k.length <= bnLoosePlain.length) continue;
        if (k === bnLoosePlain || k.includes(bnLoosePlain) || bnLoosePlain.includes(k)) {
          if (k.length < 4 || bnLoosePlain.length < 4) continue;
          out[bi] = srcIdxLoose[k];
          return;
        }
      }
      out[bi] = -1;
    });
    return out;
  }

  // ─── Render ───
  function renderMapping() {
    if (!srcData || !tplData) return;
    const sec = $('ts-section-mapping');
    const tbl = $('ts-mapping-table');
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const srcOpts = ['<option value="-1">— (leave empty) —</option>']
      .concat(srcData.headers.map((h, i) => '<option value="' + i + '">' + escHtml(h) + '</option>'))
      .join('');
    let html = '<thead><tr><th>Template column</th><th>Source column</th><th>Sample value</th><th></th></tr></thead><tbody>';
    bulkHeaders.forEach((bh, bi) => {
      const required = /\*$/.test(bh);
      const sel = mapping[bi] != null ? mapping[bi] : -1;
      const sample = sel >= 0 && srcData.rows[0] ? String(srcData.rows[0][sel] || '').trim() : '';
      html += '<tr>' +
        '<td>' + (required ? '<b>' + escHtml(bh) + '</b>' : escHtml(bh)) + '</td>' +
        '<td><select class="ts-map-select input-field" data-bi="' + bi + '" style="min-width:240px;">' +
          srcOpts.replace('value="' + sel + '"', 'value="' + sel + '" selected') +
        '</select></td>' +
        '<td><span class="text-muted small">' + escHtml(sample.slice(0, 60)) + '</span></td>' +
        '<td><button class="btn btn-ghost btn-sm ts-map-clear" data-bi="' + bi + '" title="Clear all values in this column (header stays)">Clear</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.ts-map-select').forEach(sel => {
      sel.addEventListener('change', e => {
        const bi = +e.target.dataset.bi;
        mapping[bi] = +e.target.value;
        rebuildFormattedRows();
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      });
    });
    tbl.querySelectorAll('.ts-map-clear').forEach(btn => {
      btn.addEventListener('click', e => {
        const bi = +e.target.dataset.bi;
        clearColumn(bi);
      });
    });
  }

  // Find a source-column index by checking exact then loose-alphanumeric
  // matches. `excludeContains` lets the caller reject specific source headers
  // (e.g., when looking for "Variety", exclude headers containing "clone" or
  // "subvariety" so we don't pick up Clone/Subvariety by mistake).
  function findSrcColByNames(names, excludeContains) {
    if (!srcData) return -1;
    const skip = (excludeContains || []).map(s => s.toLowerCase());
    const idx = {};
    const idxLoose = {};
    srcData.headers.forEach((h, i) => {
      const headerLower = String(h || '').toLowerCase();
      if (skip.some(s => headerLower.includes(s))) return; // never consider excluded headers
      const k = norm(h);
      const kLoose = normLoose(h);
      if (k && idx[k] == null) idx[k] = i;
      if (kLoose && idxLoose[kLoose] == null) idxLoose[kLoose] = i;
    });
    for (const n of names) {
      const k = norm(n);
      if (idx[k] != null) return idx[k];
      const kLoose = normLoose(n);
      if (idxLoose[kLoose] != null) return idxLoose[kLoose];
    }
    return -1;
  }

  function buildFormattedRowsRaw() {
    if (!srcData || !tplData) return null;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const isArchivedIdx = bulkHeaders.indexOf('Is Archived');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    // Detect a separate Crop + Variety pair in the source. If both exist we
    // combine them as "Crop-Variety" for the Crop & Variety* cell. Clone /
    // Subvariety is a different field — exclude it from variety detection.
    const cropSrcIdx = findSrcColByNames(['crop','crop type','crop name','plant_name','plant name']);
    const varSrcIdx = findSrcColByNames(['variety','varietal'], ['clone','subvariety','sub variety','sub-variety']);
    const synthesizeCv = cvIdx >= 0 && cropSrcIdx >= 0 && varSrcIdx >= 0 && cropSrcIdx !== varSrcIdx;
    const dateIdxs = dateColIdxs();
    const acreageIdx = bulkHeaders.indexOf('Acreage');

    return srcData.rows.map(srcRow => {
      const out = new Array(bulkHeaders.length).fill('');
      bulkHeaders.forEach((bh, bi) => {
        const si = mapping[bi];
        if (si != null && si >= 0 && srcRow[si] != null) {
          const v = String(srcRow[si]).trim();
          if (v) out[bi] = v;
        }
      });
      // Acreage of 0 → blank (a plot with 0 acres shouldn't carry a literal 0).
      if (acreageIdx >= 0 && out[acreageIdx] !== '') {
        const n = parseFloat(String(out[acreageIdx]).replace(/[, ]/g, ''));
        if (!isNaN(n) && n === 0) out[acreageIdx] = '';
      }
      if (synthesizeCv) {
        const c = String(srcRow[cropSrcIdx] || '').trim();
        const v = String(srcRow[varSrcIdx] || '').trim();
        if (c && v) out[cvIdx] = c + '-' + v;
        else if (c) out[cvIdx] = c;
        else if (v) out[cvIdx] = v;
      }
      // Convert Excel date serials in date columns to YYYY-MM-DD so dates read
      // as dates, not raw numbers (e.g. 46213 → 2026-07-10).
      dateIdxs.forEach(ci => { if (out[ci]) out[ci] = coerceExcelDate(out[ci]); });
      if (target === 'update' && isArchivedIdx >= 0 && !out[isArchivedIdx]) {
        out[isArchivedIdx] = 'FALSE';
      }
      return out;
    });
  }

  // Rebuild formatted rows from source + mapping, then re-apply any manual
  // fills the user has applied. Called when the inputs change (mapping edited,
  // target switched, source/template re-uploaded). NOT called by renderPreview
  // — that function just displays the canonical formattedRows state.
  function rebuildFormattedRows() {
    formattedRows = buildFormattedRowsRaw();
    if (!formattedRows) return;
    Object.entries(manualFills).forEach(([idx, val]) => {
      const i = +idx;
      formattedRows.forEach(r => { if (!r[i]) r[i] = val; });
    });
    // Expand year-only / year-range date cells to full YYYY-MM-DD using the
    // operator's chosen month/day. Re-derived from raw source on every rebuild,
    // so changing the month/day re-expands cleanly (non-destructive).
    if (dateFill.mm && dateFill.dd) {
      const idxs = dateFillIdxs();
      formattedRows.forEach(r => {
        idxs.forEach(ci => {
          if (isYearOnlyDate(r[ci])) r[ci] = expandYearOnlyDate(r[ci], dateFill.mm, dateFill.dd, dateFill.range);
        });
      });
    }
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    // Split a "Crop-Variety - Subvariety" value: keep "Crop-Variety" in the
    // Crop & Variety cell and move the part after the first " - " into
    // Clone/Subvariety (only when that cell is empty). PickTrace's location
    // upload rejects a subvariety embedded in Crop & Variety.
    if (cropSplit) {
      const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
      const cloneIdx = bulkHeaders.indexOf('Clone/Subvariety');
      if (cvIdx >= 0) {
        formattedRows.forEach(r => {
          const v = String(r[cvIdx] == null ? '' : r[cvIdx]);
          const m = v.match(/^(.*?)\s+-\s+(.+)$/); // split on the FIRST " - "
          if (m) {
            r[cvIdx] = m[1].trim();
            if (cloneIdx >= 0 && !String(r[cloneIdx] || '').trim()) r[cloneIdx] = m[2].trim();
          }
        });
      }
    }
    const siteIdx = bulkHeaders.indexOf('Site*');
    const nameIdx = bulkHeaders.indexOf('Name*');
    const cvIdxR = bulkHeaders.indexOf('Crop & Variety*');
    // Re-apply sticky Name* overrides (collision resolutions) FIRST, so a
    // renamed row no longer counts as a collision or a duplicate.
    if (nameIdx >= 0) {
      Object.keys(nameOverrides).forEach(k => {
        const ri = +k;
        if (formattedRows[ri] && nameOverrides[k]) formattedRows[ri][nameIdx] = nameOverrides[k];
      });
    }
    // Apply accepted Smart Fixes (bulk snap of off-list values to the template
    // dropdown, e.g. "Almonds-Other" → "Almond-Other"). Keyed by column + upper
    // value, applied to every matching cell. User cell edits (below) still win.
    if (Object.keys(smartFixMap).length) {
      const dropIdxs = dropdownColIdxs();
      formattedRows.forEach(r => {
        dropIdxs.forEach(ci => {
          const v = r[ci];
          if (!v) return;
          const hit = smartFixMap[ci + '||' + String(v).toUpperCase().trim()];
          if (hit) r[ci] = hit;
        });
      });
    }
    // Re-apply sticky per-cell edits from the editable preview. Applied AFTER
    // mapping / crop-split / date-fill / name-overrides and BEFORE collision +
    // existing detection so edited Site/Crop/etc. feed those computations.
    // Never touches Name* — that column is owned by nameOverrides above.
    Object.keys(cellOverrides).forEach(k => {
      const sep = k.indexOf('|');
      const ri = +k.slice(0, sep), ci = +k.slice(sep + 1);
      if (ci === nameIdx) return;
      if (formattedRows[ri]) formattedRows[ri][ci] = cellOverrides[k];
    });
    const rowNameKey = ri => locKey(formattedRows[ri][siteIdx], formattedRows[ri][nameIdx]);
    const rowBlockKey = ri => rowNameKey(ri) + '||' + locNorm(cvIdxR >= 0 ? formattedRows[ri][cvIdxR] : '');
    // Count how many source rows share each Site + Name (a within-file collision
    // protects against being silently dropped by a name-only existing file).
    const withinCounts = new Map();
    if (siteIdx >= 0 && nameIdx >= 0) formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const s = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
      const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!s || !n) return;
      const k = rowNameKey(ri);
      withinCounts.set(k, (withinCounts.get(k) || 0) + 1);
    });
    // Cross-reference: drop a row ONLY when the SAME block (Site + Name + Crop &
    // Variety) already exists in PickTrace — so "SGS / Golden Delicious" is
    // recognized as already there, while "SGS / Granny Smith" (a different
    // block reusing the name) is NOT dropped and gets flagged as a collision.
    // If the existing file has no Crop column we fall back to Site + Name, but
    // never auto-drop a within-file collision (we can't tell the blocks apart).
    existingRowIdxs = new Set();
    if (existingData && existingData.keys && existingData.keys.size && siteIdx >= 0 && nameIdx >= 0) {
      const useBlock = !!(existingData.hasCrop && existingData.blockKeys);
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri)) return;
        const s = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
        const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
        if (!s || !n) return;
        const match = useBlock
          ? existingData.blockKeys.has(rowBlockKey(ri))
          : (existingData.keys.has(rowNameKey(ri)) && (withinCounts.get(rowNameKey(ri)) || 0) <= 1);
        if (match) existingRowIdxs.add(ri);
      });
    }
    // Collisions among the rows that will actually be CREATED (not removed, not
    // already-in-PickTrace): (a) a Site + Name shared by 2+ such rows, or
    // (b) a Site + Name that PickTrace already uses for a DIFFERENT block.
    collisionKeys = new Set();
    existingTakenKeys = new Set();
    if (siteIdx >= 0 && nameIdx >= 0) {
      const counts = new Map();
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri) || existingRowIdxs.has(ri)) return;
        const s = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
        const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
        if (!s || !n) return;
        const kk = rowNameKey(ri);
        counts.set(kk, (counts.get(kk) || 0) + 1);
        if (existingData && existingData.keys && existingData.keys.has(kk)) {
          collisionKeys.add(kk);
          existingTakenKeys.add(kk);
        }
      });
      counts.forEach((c, kk) => { if (c >= 2) collisionKeys.add(kk); });
    }
  }

  // Groups the to-be-created rows that fall under a flagged collision name
  // (computed in rebuildFormattedRows). Returns Map<locKey, [rowIdx,…]>. A group
  // can have a single row when the name is taken by an existing PickTrace block
  // (existingTakenKeys) — that row still can't load under its name.
  function computeCollisions() {
    const out = new Map();
    if (!formattedRows || !collisionKeys.size) return out;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*');
    const nameIdx = bulkHeaders.indexOf('Name*');
    if (siteIdx < 0 || nameIdx < 0) return out;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri) || existingRowIdxs.has(ri)) return;
      const site = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
      const name = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!site || !name) return;
      const k = locKey(site, name);
      if (!collisionKeys.has(k)) return;
      const arr = out.get(k) || [];
      arr.push(ri);
      out.set(k, arr);
    });
    return out;
  }
  function collisionRowSet() {
    const s = new Set();
    computeCollisions().forEach(arr => arr.forEach(ri => s.add(ri)));
    return s;
  }

  function renderPreview() {
    if (!srcData || !tplData || !formattedRows) return;
    renderExisting();
    renderCollisions();
    renderSmartFixes();
    renderDateFill();
    const sec = $('ts-section-preview');
    const tbl = $('ts-preview-table');
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const collisionIdxs = collisionRowSet();
    const nameColIdx = bulkHeaders.indexOf('Name*');
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    const dateIdxSet = new Set(dateFillIdxs());
    // Per-column dropdown metadata, built once: the value list for the cell's
    // <datalist> and a caseless set used to flag off-list values amber.
    const colDrop = bulkHeaders.map(h => {
      const vals = dropdownValuesFor(h);
      if (!vals) return null;
      return { vals, lower: new Set(vals.map(v => String(v).toLowerCase().trim())) };
    });
    let html = '<thead><tr>';
    // First column is the row-action column (X to remove the row).
    html += '<th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    bulkHeaders.forEach((h, i) => {
      const req = /\*$/.test(h);
      const listed = colDrop[i] ? ' title="Dropdown from template — pick a value or type your own"' : '';
      html += '<th' + (req ? ' style="color:#dc2626;"' : '') + listed + '>' + escHtml(h) +
        (colDrop[i] ? ' <span class="ts-col-listed" title="Has a template dropdown">&#9662;</span>' : '') + '</th>';
    });
    html += '</tr></thead><tbody>';
    // A row "needs attention" if any cell would be flagged in the preview:
    // empty required, off-list dropdown value, name collision, past Start Date,
    // or a year-only date.
    const rowHasIssue = ri => {
      const row = formattedRows[ri];
      for (let i = 0; i < bulkHeaders.length; i++) {
        const val = row[i] == null ? '' : String(row[i]);
        if (/\*$/.test(bulkHeaders[i]) && !val) return true;
        if (i === startDateIdx && val && isDateBeforeToday(val)) return true;
        if (dateIdxSet.has(i) && isYearOnlyDate(val)) return true;
        if (i === nameColIdx && collisionIdxs.has(ri)) return true;
        const d = colDrop[i];
        if (d && val && !d.lower.has(val.toLowerCase().trim())) return true;
      }
      return false;
    };
    const allVisible = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const issueCount = allVisible.filter(rowHasIssue).length;
    if (previewIssuesOnly && !issueCount) previewIssuesOnly = false; // nothing to filter to
    const visibleRowIdxs = previewIssuesOnly ? allVisible.filter(rowHasIssue) : allVisible;
    // No cap in "needs attention" mode (safety ceiling 2000); 50 in normal mode.
    const CAP = previewIssuesOnly ? 2000 : 50;
    const limit = Math.min(visibleRowIdxs.length, CAP);
    for (let r = 0; r < limit; r++) {
      const ri = visibleRowIdxs[r];
      const row = formattedRows[ri];
      html += '<tr>';
      html += '<td style="text-align:center;padding:0;">' +
        '<button class="ts-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button>' +
        '</td>';
      bulkHeaders.forEach((h, i) => {
        const req = /\*$/.test(h);
        const val = row[i] == null ? '' : String(row[i]);
        const empty = !val;
        const d = colDrop[i];
        const inList = !!(d && val && d.lower.has(val.toLowerCase().trim()));
        let bg = '', title = '';
        if (req && empty) { bg = '#fee2e2'; title = 'Required — must be filled before export'; }
        else if (i === startDateIdx && val && isDateBeforeToday(val)) {
          bg = '#fef3c7'; title = 'Start Date must be today (' + todayYMD() + ') or later';
        } else if (dateIdxSet.has(i) && isYearOnlyDate(val)) {
          bg = '#fef3c7'; title = 'Year only — set a month/day in Year-Only Dates above to make this a full YYYY-MM-DD date';
        } else if (i === nameColIdx && collisionIdxs.has(ri)) {
          bg = '#fee2e2'; title = 'Name collision — another block at this Site shares this Name. Rename it here (or in Name Collisions above), or PickTrace will drop one.';
        } else if (d && val && !inList) {
          bg = '#fef3c7'; title = 'Off-list — "' + val + '" isn’t in the template dropdown for ' + h + '. Kept as-is (a new value to create in PickTrace), or pick a listed value.';
        }
        const cellStyle = bg ? ' style="background:' + bg + ';"' : '';
        const fieldColor = bg ? ('color:' + (bg === '#fee2e2' ? '#7f1d1d' : '#7c2d12') + ';') : '';
        if (d) {
          // Real <select> combo: always shows the full template list; keeps the
          // current off-list value as a "(current)" option; "Type custom…" swaps
          // the cell to a free-text input for a brand-new value.
          let opts = '<option value=""' + (empty ? ' selected' : '') + '>— blank —</option>';
          if (val && !inList) opts += '<option value="' + escHtml(val) + '" selected>' + escHtml(val) + '  (current)</option>';
          d.vals.forEach(v => {
            const s = (inList && v.toLowerCase().trim() === val.toLowerCase().trim()) ? ' selected' : '';
            opts += '<option value="' + escHtml(v) + '"' + s + '>' + escHtml(v) + '</option>';
          });
          opts += '<option value="__ts_custom__">✎ Type custom…</option>';
          html += '<td class="ts-cell"' + cellStyle + '>' +
            '<select class="ts-cell-select" data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') +
            (fieldColor ? ' style="' + fieldColor + '"' : '') + '>' + opts + '</select>' +
            '</td>';
        } else {
          html += '<td class="ts-cell"' + cellStyle + '>' +
            '<input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') +
            (fieldColor ? ' style="' + fieldColor + '"' : '') +
            ' value="' + escHtml(val) + '">' +
            '</td>';
        }
      });
      html += '</tr>';
    }
    html += '</tbody>';
    tbl.innerHTML = html;
    const hint = sec.querySelector('.cmp-sites-hint');
    // "Needs attention" toggle — lets the user see EVERY errored row (no cap),
    // not just the first 50. Disabled when there are no issues.
    const toggle = '<label class="ts-issues-toggle" style="display:inline-flex;align-items:center;gap:6px;margin-right:12px;font-weight:600;' +
      (issueCount ? '' : 'opacity:.5;') + '">' +
      '<input type="checkbox" id="ts-issues-only"' + (previewIssuesOnly ? ' checked' : '') + (issueCount ? '' : ' disabled') + '>' +
      'Show only rows needing attention' + (issueCount ? ' (' + issueCount + ')' : ' (0)') + '</label>';
    let hintText = toggle;
    if (previewIssuesOnly) {
      hintText += 'Showing ' + Math.min(limit, visibleRowIdxs.length) + ' of ' + issueCount + ' flagged row' + (issueCount === 1 ? '' : 's') +
        (visibleRowIdxs.length > limit ? ' (capped at ' + limit + ')' : '') + '. ';
    } else if (visibleRowIdxs.length > limit) {
      hintText += 'Showing first ' + limit + ' of ' + visibleRowIdxs.length + ' rows. ';
    }
    hintText += 'Every cell is editable — <span class="ts-col-listed">&#9662;</span> columns are template dropdowns. ' +
      'Red = empty required; amber = off-list / needs attention. Click <b>×</b> to drop a row.';
    if (removedRows.size) {
      hintText += ' &nbsp; <button class="btn btn-ghost btn-sm" id="ts-restore-rows">' +
        'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    }
    hint.innerHTML = hintText;
    const issuesOnly = $('ts-issues-only');
    if (issuesOnly) issuesOnly.addEventListener('change', e => { previewIssuesOnly = !!e.target.checked; renderPreview(); });
    tbl.querySelectorAll('.ts-row-remove').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        const ri = +e.currentTarget.dataset.ri;
        removedRows.add(ri);
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      });
    });
    const restoreBtn = $('ts-restore-rows');
    if (restoreBtn) restoreBtn.addEventListener('click', () => {
      removedRows.clear();
      renderPreview();
      renderRequired();
      renderSitesAndCvToCreate();
      updateSummary();
    });
    // Editable-cell handlers, delegated on the persistent <table> so we attach
    // them once (not once per cell) and they survive innerHTML re-renders.
    if (!tbl._tsCellWired) {
      tbl._tsCellWired = true;
      tbl.addEventListener('change', e => {
        const sel = e.target.closest('.ts-cell-select');
        if (sel) {
          const ri = +sel.dataset.ri, ci = +sel.dataset.ci;
          // "Type custom…" swaps the cell to a free-text input in place; the real
          // value gets committed by the input's own change event below.
          if (sel.value === '__ts_custom__') { swapSelectToInput(sel, ri, ci); return; }
          setCell(ri, ci, sel.value);
          recomputeAfterEdit();
          return;
        }
        const inp = e.target.closest('.ts-cell-input');
        if (inp) {
          setCell(+inp.dataset.ri, +inp.dataset.ci, inp.value);
          recomputeAfterEdit();
        }
      });
    }
    updateExportButton();
  }

  // Route a preview cell edit into the correct sticky store + the live grid.
  // Name* flows through nameOverrides (keeps the Name Collisions panel in sync);
  // every other column flows through cellOverrides. Empty values are stored too
  // (except Name*, which clears its override) so a deliberate clear sticks across
  // rebuilds instead of snapping back to the mapped source value.
  function setCell(ri, ci, val) {
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const nameIdx = bulkHeaders.indexOf('Name*');
    const v = val == null ? '' : String(val);
    if (ci === nameIdx) {
      if (v.trim()) nameOverrides[ri] = v.trim(); else delete nameOverrides[ri];
    } else {
      cellOverrides[ri + '|' + ci] = v;
    }
    if (formattedRows[ri]) formattedRows[ri][ci] = v;
  }

  // Replace a listed-column <select> with a free-text input in place, so the
  // user can enter a brand-new value not in the template list. The input's own
  // change event commits it (and re-renders the cell back to a select).
  function swapSelectToInput(sel, ri, ci) {
    const td = sel.closest('td');
    if (!td) return;
    const cur = formattedRows[ri] ? (formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci])) : '';
    td.innerHTML = '<input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + ci +
      '" placeholder="Type a value…" value="' + escHtml(cur) + '">';
    const inp = td.querySelector('input');
    if (inp) { inp.focus(); inp.select(); }
  }

  // After any preview edit: rebuild (re-applies sticky edits, recomputes
  // collisions + existing-in-PickTrace), then refresh the dependent panels.
  function recomputeAfterEdit() {
    rebuildFormattedRows();
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
  }

  // ─── Year-only date expansion control ───
  function renderDateFill() {
    const sec = $('ts-section-datefill');
    if (!sec || !formattedRows) return;
    sec.style.display = '';
    const idxs = dateFillIdxs();
    const visRows = formattedRows.filter((_, i) => !excluded(i));
    const n = countYearOnlyDates(visRows, idxs);
    const cnt = $('ts-datefill-count');
    if (cnt) cnt.innerHTML = n
      ? '<b style="color:#b45309;">' + n + '</b> year-only date cell' + (n === 1 ? '' : 's') + ' remaining'
      : '<span style="color:#15803d;font-weight:600;">✓ No year-only date cells</span>';
    // Reflect current settings in the inputs.
    const mmEl = $('ts-datefill-mm'), ddEl = $('ts-datefill-dd'), rgEl = $('ts-datefill-range');
    if (mmEl && document.activeElement !== mmEl) mmEl.value = dateFill.mm;
    if (ddEl && document.activeElement !== ddEl) ddEl.value = dateFill.dd;
    if (rgEl) rgEl.value = dateFill.range;
  }

  function applyDateFill() {
    const mm = ($('ts-datefill-mm').value || '').trim();
    const dd = ($('ts-datefill-dd').value || '').trim();
    const range = $('ts-datefill-range').value || 'first';
    const mi = parseInt(mm, 10), di = parseInt(dd, 10);
    if (!(mi >= 1 && mi <= 12) || !(di >= 1 && di <= 31)) {
      alert('Enter a valid month (1–12) and day (1–31) first.');
      return;
    }
    dateFill = { mm, dd, range };
    rebuildFormattedRows();
    renderRequired();
    renderSitesAndCvToCreate();
    renderPreview();
    updateSummary();
  }

  function clearDateFill() {
    dateFill = { mm: '', dd: '', range: ($('ts-datefill-range') ? $('ts-datefill-range').value : 'first') };
    if ($('ts-datefill-mm')) $('ts-datefill-mm').value = '';
    if ($('ts-datefill-dd')) $('ts-datefill-dd').value = '';
    rebuildFormattedRows();
    renderRequired();
    renderSitesAndCvToCreate();
    renderPreview();
    updateSummary();
  }

  function renderSitesAndCvToCreate() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx   = bulkHeaders.indexOf('Crop & Variety*');
    const tplSites = tplData.dropdowns.get('site') || new Set();
    const tplCvs   = tplData.dropdowns.get('crop & variety') || new Set();
    const tplSitesMap = buildCaseMap(tplSites);
    const tplCvsMap   = buildCaseMap(tplCvs);

    const sites = new Map(), cvs = new Map();
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      const s = r[siteIdx];
      if (s && !inDropdown(tplSitesMap, s)) sites.set(s, (sites.get(s) || 0) + 1);
      const v = r[cvIdx];
      if (v && !inDropdown(tplCvsMap, v))   cvs.set(v, (cvs.get(v) || 0) + 1);
    });

    const renderList = (sectionId, tableId, titleId, label, m, withAddress) => {
      const sec = $(sectionId);
      if (!m.size) { sec.style.display = 'none'; return; }
      sec.style.display = '';
      $(titleId).textContent = label + ' (' + m.size + ')';
      const sorted = [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
      let html = '<thead><tr><th>Value</th><th>Row count</th>' +
        (withAddress ? '<th>Address</th>' : '') + '</tr></thead><tbody>';
      if (withAddress) {
        html += '<tr class="ts-addr-allrow">' +
          '<td colspan="2"><b>Apply one address to every site</b>' +
            '<div class="text-muted small">Optional — for your reference when creating these Sites in PickTrace.</div></td>' +
          '<td><div class="ts-addr-all-wrap">' +
            '<input type="text" id="ts-site-addr-all" class="input-field" placeholder="123 Main St, City, ST">' +
            '<button class="btn btn-primary btn-sm" id="ts-site-addr-apply-all" type="button">Apply to all</button>' +
          '</div></td></tr>';
      }
      sorted.forEach(([val, cnt]) => {
        html += '<tr><td class="ts-copy-cell" title="Click to copy">' + escHtml(val) +
          '</td><td>' + cnt + '</td>';
        if (withAddress) {
          html += '<td><input type="text" class="ts-site-addr input-field" data-site="' +
            escHtml(val) + '" value="' + escHtml(siteAddresses[val] || '') +
            '" placeholder="Address (optional)"></td>';
        }
        html += '</tr>';
      });
      html += '</tbody>';
      $(tableId).innerHTML = html;
      if (withAddress) wireSitesAddressHandlers(tableId);
    };
    renderList('ts-section-sites-create', 'ts-sites-create-table', 'ts-sites-create-title', 'Sites to Create in PickTrace', sites, true);
    renderList('ts-section-cv-create',    'ts-cv-create-table',    'ts-cv-create-title',    'Crops & Varieties to Create in PickTrace', cvs, false);
    return { sitesCount: sites.size, cvsCount: cvs.size };
  }

  // Per-row address inputs + "Apply to all". Addresses are reference-only
  // (keyed by site value in siteAddresses) so they survive the frequent
  // renderSitesAndCvToCreate() rebuilds triggered by Apply/Clear/remove.
  function wireSitesAddressHandlers(tableId) {
    const tbl = $(tableId);
    if (!tbl) return;
    tbl.querySelectorAll('.ts-site-addr').forEach(inp => {
      inp.addEventListener('input', () => {
        siteAddresses[inp.dataset.site] = inp.value;
      });
    });
    const allInp = tbl.querySelector('#ts-site-addr-all');
    const allBtn = tbl.querySelector('#ts-site-addr-apply-all');
    if (allInp && allBtn) {
      allBtn.addEventListener('click', () => {
        const v = allInp.value.trim();
        if (!v) { alert('Type an address first.'); return; }
        tbl.querySelectorAll('.ts-site-addr').forEach(inp => {
          inp.value = v;
          siteAddresses[inp.dataset.site] = v;
        });
      });
    }
  }

  // ─── Smart Fixes panel — every off-list value gets a dropdown to choose the
  // target template value (suggested match pre-selected; the rest ask the user). ───
  function renderSmartFixes() {
    const sec = $('ts-section-smartfix');
    if (!sec) return;
    const fixes = computeSmartFixes();
    if (!fixes.length) { sec.style.display = 'none'; $('ts-smartfix-table').innerHTML = ''; return; }
    sec.style.display = '';
    let totalRows = 0, needPick = 0;
    fixes.forEach(f => { totalRows += f.count; if (!f.suggestion) needPick++; });
    const t = $('ts-smartfix-title');
    if (t) t.textContent = 'Smart Fixes (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + totalRows + ' rows' +
      (needPick ? ' — ' + needPick + ' need your choice' : '') + ')';
    let html = '<thead><tr><th>Column</th><th>Value in your data</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    fixes.forEach(f => {
      const opts = '<option value="">— pick a value —</option>' +
        f.options.slice().sort().map(o => '<option' + (f.suggestion && o === f.suggestion ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('');
      html += '<tr>' +
        '<td>' + escHtml(f.header) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td><select class="ts-smartfix-pick input-field" style="min-width:200px;">' + opts + '</select>' +
          (f.suggestion ? ' <span class="text-muted small" title="Confident match — spacing, punctuation or singular/plural">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') +
        '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm ts-smartfix-apply" data-ci="' + f.colIdx + '" data-from="' + escHtml(f.from) + '">Apply</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    $('ts-smartfix-table').innerHTML = html;
    $('ts-smartfix-table').querySelectorAll('.ts-smartfix-apply').forEach(btn => btn.addEventListener('click', e => {
      const tr = e.currentTarget.closest('tr');
      const to = tr.querySelector('.ts-smartfix-pick').value.trim();
      if (!to) { alert('Pick a value to set "' + e.currentTarget.dataset.from + '" to first.'); return; }
      applySmartFix(+e.currentTarget.dataset.ci, e.currentTarget.dataset.from, to);
    }));
    return fixes.length;
  }
  function applySmartFix(ci, from, to) {
    smartFixMap[ci + '||' + String(from).toUpperCase().trim()] = to;
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderSitesAndCvToCreate(); updateSummary();
  }
  // Apply every Smart Fixes row that currently has a value chosen in its
  // dropdown (suggested ones are pre-selected; the user picks the rest).
  function applyAllSmartFixes() {
    const tbl = $('ts-smartfix-table');
    if (!tbl) return;
    const picks = [];
    tbl.querySelectorAll('.ts-smartfix-apply').forEach(btn => {
      const tr = btn.closest('tr');
      const to = tr.querySelector('.ts-smartfix-pick').value.trim();
      if (to) picks.push({ ci: +btn.dataset.ci, from: btn.dataset.from, to });
    });
    if (!picks.length) { alert('No values chosen yet. Pick a target for at least one row (suggested rows are pre-selected).'); return; }
    picks.forEach(p => { smartFixMap[p.ci + '||' + String(p.from).toUpperCase().trim()] = p.to; });
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderSitesAndCvToCreate(); updateSummary();
  }

  function renderRequired() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const sec = $('ts-section-required');
    sec.style.display = '';
    const requiredCols = bulkHeaders
      .map((h, i) => ({ header: h, idx: i, req: /\*$/.test(h) }))
      .filter(x => x.req);

    // Split into two buckets: rows that still need attention vs rows that
    // are already filled. Filled rows stay visible (with a "Change value"
    // affordance) so the user can re-pick or clear at any time.
    const needsAction = [];
    const filled = [];
    requiredCols.forEach(col => {
      let empty = 0;
      const sample = new Set();
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        if (!r[col.idx]) empty++;
        else if (sample.size < 1 && r[col.idx]) sample.add(r[col.idx]);
      });
      const fillVal = manualFills[col.idx] || (sample.size ? [...sample][0] : '');
      const entry = { ...col, empty, fillVal };
      if (empty > 0) needsAction.push(entry); else filled.push(entry);
    });

    const renderRow = (col) => {
      const dropKey = norm(col.header).replace(/\*$/, '');
      const opts = tplData.dropdowns.get(dropKey)
        || tplData.dropdowns.get(dropKey + '*')
        || tplData.dropdowns.get(norm(col.header))
        || new Set();
      const optsHtml = opts.size
        ? '<option value="">— pick —</option>' +
          [...opts].sort().map(o => '<option' + (o === col.fillVal ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const status = col.empty > 0
        ? '<span style="color:#dc2626;font-weight:600;">⚠ ' + col.empty + ' empty</span>'
        : '<span style="color:#15803d;">✓ filled' + (col.fillVal ? ' (e.g. ' + escHtml(String(col.fillVal).slice(0, 30)) + ')' : '') + '</span>';
      const rowStyle = col.empty > 0 ? '' : ' style="background:var(--bg-sunken);"';
      // Date-aware quick fill — Start Date* gets a "Today" button that
      // fills every visible row with today's YYYY-MM-DD value.
      const isStartDate = /^start\s*date/i.test(col.header.replace(/\*$/, ''));
      const todayBtn = isStartDate
        ? '<button class="btn btn-warning btn-sm ts-req-today" data-idx="' + col.idx + '" title="Fill all rows with today\'s date (' + todayYMD() + ')">Today</button>'
        : '';
      return '<tr' + rowStyle + '>' +
        '<td><b>' + escHtml(col.header) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="ts-req-pick input-field" data-idx="' + col.idx + '" style="min-width:200px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="ts-req-override input-field" data-idx="' + col.idx + '" placeholder="Manual override" style="width:200px;"></td>' +
        '<td>' + todayBtn + '</td>' +
        '<td><button class="btn btn-primary btn-sm ts-req-apply" data-idx="' + col.idx + '">Apply</button></td>' +
        '<td><button class="btn btn-ghost btn-sm ts-req-clear" data-idx="' + col.idx + '" title="Clear all values in this column (header stays)">Clear</button></td>' +
        '</tr>';
    };

    let html = '<thead><tr><th>Column</th><th>Status</th><th>Pick from dropdown</th><th>Manual override</th><th>Quick fill</th><th></th><th></th></tr></thead>';
    if (needsAction.length) {
      html += '<tbody><tr><td colspan="7" style="background:#fee2e2;color:#7f1d1d;font-weight:600;padding:6px 10px;">' +
        '⚠ Action needed (' + needsAction.length + ')</td></tr>' +
        needsAction.map(renderRow).join('') + '</tbody>';
    }
    if (filled.length) {
      html += '<tbody><tr><td colspan="7" style="background:#dcfce7;color:#14532d;font-weight:600;padding:6px 10px;">' +
        '✓ Already filled — change or clear if needed (' + filled.length + ')</td></tr>' +
        filled.map(renderRow).join('') + '</tbody>';
    }
    $('ts-required-table').innerHTML = html;
    wireRequiredHandlers();
  }

  // ─── Apply / Clear handlers (extracted so renderRequired stays readable) ───
  function applyFill(idx, val, overwrite) {
    if (!val) return 0;
    let filled = 0;
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      if (overwrite || !r[idx]) { r[idx] = val; filled++; }
    });
    manualFills[idx] = val;
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
    return filled;
  }
  function clearColumn(idx) {
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      r[idx] = '';
    });
    delete manualFills[idx];
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
  }
  function wireRequiredHandlers() {
    const tbl = $('ts-required-table');
    if (!tbl) return;
    tbl.querySelectorAll('.ts-req-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        const tr = e.target.closest('tr');
        const pick = tr.querySelector('.ts-req-pick').value.trim();
        const ovr  = tr.querySelector('.ts-req-override').value.trim();
        const val = ovr || pick;
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        // Apply OVERWRITES — that way the user can change their mind and
        // re-apply a different value to the same column.
        applyFill(idx, val, true);
      });
    });
    tbl.querySelectorAll('.ts-req-clear').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        clearColumn(idx);
      });
    });
    tbl.querySelectorAll('.ts-req-today').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        applyFill(idx, todayYMD(), true);
      });
    });
    // Real-time auto-apply ONLY for the dropdown picker (instant, intentional).
    // The text override is hands-off until the user clicks Apply — typing no
    // longer auto-fills mid-keystroke.
    tbl.querySelectorAll('.ts-req-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        const idx = +e.target.dataset.idx;
        const v = e.target.value.trim();
        if (v) applyFill(idx, v, true);
      });
    });
  }

  function updateSummary() {
    if (!srcData || !tplData) { $('ts-summary').style.display = 'none'; return; }
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const reqIdxs = bulkHeaders.map((h, i) => ({ h, i, req: /\*$/.test(h) })).filter(x => x.req);
    let emptyReq = 0;
    if (formattedRows) reqIdxs.forEach(({ i }) => {
      formattedRows.forEach((r, ri) => { if (!excluded(ri) && !r[i]) emptyReq++; });
    });
    const sites = renderSitesAndCvToCreate() || { sitesCount: 0, cvsCount: 0 };
    const visibleCount = formattedRows ? formattedRows.filter((_, i) => !excluded(i)).length : 0;
    const totalCount = formattedRows ? formattedRows.length : 0;
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    let pastDates = 0;
    if (formattedRows && startDateIdx >= 0) {
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        if (isDateBeforeToday(r[startDateIdx])) pastDates++;
      });
    }
    const sum = $('ts-summary');
    sum.style.display = '';
    const existingCount = existingRowIdxs.size;
    let collisionRows = 0; computeCollisions().forEach(arr => collisionRows += arr.length);
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visibleCount + '</b> rows' +
        (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + totalCount + ')</span>' : '') +
        '</div>' +
      (existingCount ? '<div class="cmp-stat cmp-warn"><b>' + existingCount + '</b> already in PickTrace (dropped)</div>' : '') +
      (collisionRows ? '<div class="cmp-stat cmp-warn"><b>' + collisionRows + '</b> name collision rows</div>' : '') +
      '<div class="cmp-stat"><b>' + bulkHeaders.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (sites.sitesCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.sitesCount + '</b> sites to create</div>' : '') +
      (sites.cvsCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.cvsCount + '</b> C&V to create</div>' : '') +
      (pastDates ? '<div class="cmp-stat cmp-warn"><b>' + pastDates + '</b> Start Dates before today</div>' : '');
  }

  // Build a case-insensitive index of a dropdown Set so values like
  // "Beckstoffer Vineyards" match the template's "BECKSTOFFER VINEYARDS".
  // Returns Map<upperTrimmed, canonicalValue>. Use .has() / .get() instead of
  // the raw Set's case-sensitive .has().
  function buildCaseMap(set) {
    const m = new Map();
    if (!set) return m;
    set.forEach(v => {
      const k = String(v).toUpperCase().trim();
      if (k && !m.has(k)) m.set(k, v);
    });
    return m;
  }
  function inDropdown(map, value) {
    if (!value || !map.size) return false;
    return map.has(String(value).toUpperCase().trim());
  }

  function getExportBlockers() {
    if (!srcData || !tplData || !formattedRows) return null;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const reqIdxs = bulkHeaders
      .map((h, i) => /\*$/.test(h) ? i : -1)
      .filter(i => i >= 0);
    let emptyReq = 0;
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      reqIdxs.forEach(i => { if (!r[i]) emptyReq++; });
    });
    const tplSites = tplData.dropdowns.get('site') || new Set();
    const tplCvs   = tplData.dropdowns.get('crop & variety') || new Set();
    const tplSitesMap = buildCaseMap(tplSites);
    const tplCvsMap   = buildCaseMap(tplCvs);
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    const checkSites = tplSites.size > 0;
    const checkCvs   = tplCvs.size > 0;
    const unknownSiteVals = new Set();
    const unknownCvVals = new Set();
    let pastStartDates = 0;
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      if (checkSites && r[siteIdx] && !inDropdown(tplSitesMap, r[siteIdx])) unknownSiteVals.add(r[siteIdx]);
      if (checkCvs   && r[cvIdx]   && !inDropdown(tplCvsMap,   r[cvIdx]))   unknownCvVals.add(r[cvIdx]);
      if (startDateIdx >= 0 && isDateBeforeToday(r[startDateIdx])) pastStartDates++;
    });
    const collisionGroups = computeCollisions();
    let nameCollisions = 0; collisionGroups.forEach(arr => nameCollisions += arr.length);
    const collisionSamples = [];
    collisionGroups.forEach((arr, k) => {
      const r0 = formattedRows[arr[0]];
      collisionSamples.push(String(r0[siteIdx] || '') + ' / ' + String(r0[bulkHeaders.indexOf('Name*')] || '') + ' (×' + arr.length + ')');
    });
    return { emptyReq, unknownSites: [...unknownSiteVals], unknownCvs: [...unknownCvVals], pastStartDates, nameCollisions, collisionSamples };
  }

  function updateExportButton() {
    // Button stays ENABLED whenever there's data — clicking will surface a
    // confirmation when there are blockers, instead of leaving the user stuck.
    const btn = $('ts-export');
    btn.disabled = !(srcData && tplData && formattedRows && formattedRows.length);
    const b = getExportBlockers();
    if (!b) { btn.title = ''; return; }
    const reasons = [];
    if (b.emptyReq > 0) reasons.push(b.emptyReq + ' empty required cells');
    if (b.unknownSites.length) reasons.push(b.unknownSites.length + ' sites not in template dropdown');
    if (b.unknownCvs.length) reasons.push(b.unknownCvs.length + ' Crop & Variety values not in template dropdown');
    if (b.nameCollisions > 0) reasons.push(b.nameCollisions + ' rows with duplicate Site + Name (collision)');
    if (b.pastStartDates > 0) reasons.push(b.pastStartDates + ' Start Date values are before today');
    btn.title = reasons.length
      ? 'Click to review issues before exporting: ' + reasons.join(', ') + '.'
      : 'Ready to export.';
  }

  // ─── Custom export-confirm modal ───
  // Replaces the native window.confirm() which can't be styled and renders
  // raw blocker text. Reuses the .cmp-export-modal classes already defined
  // in styles.css so the look matches Block Compare's gate dialog.
  function showExportConfirm(b) {
    return new Promise(resolve => {
      const prior = document.getElementById('ts-export-confirm');
      if (prior) prior.remove();

      const overlay = document.createElement('div');
      overlay.id = 'ts-export-confirm';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';

      const sections = [];
      if (b.emptyReq > 0) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.emptyReq + ' empty required cells</div>' +
          '<div class="ts-confirm-issue-body">PickTrace will reject rows missing values for required (*) columns. Fill them via the Empty Required Columns panel.</div>' +
          '</div>'
        );
      }
      if (b.unknownSites.length) {
        const sample = b.unknownSites.slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = b.unknownSites.length > 8 ? '<li class="ts-confirm-more">… and ' + (b.unknownSites.length - 8) + ' more</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unknownSites.length + ' Site value(s) not in your template dropdown</div>' +
          '<ul class="ts-confirm-list">' + sample + more + '</ul>' +
          '<div class="ts-confirm-issue-body">Create these Sites in PickTrace, re-download your personalized template, and re-upload here.</div>' +
          '</div>'
        );
      }
      if (b.unknownCvs.length) {
        const sample = b.unknownCvs.slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = b.unknownCvs.length > 8 ? '<li class="ts-confirm-more">… and ' + (b.unknownCvs.length - 8) + ' more</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unknownCvs.length + ' Crop &amp; Variety value(s) not in your template dropdown</div>' +
          '<ul class="ts-confirm-list">' + sample + more + '</ul>' +
          '<div class="ts-confirm-issue-body">Create these in PickTrace and re-download the template.</div>' +
          '</div>'
        );
      }
      if (b.nameCollisions > 0) {
        const sample = (b.collisionSamples || []).slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = (b.collisionSamples || []).length > 8 ? '<li class="ts-confirm-more">… and ' + (b.collisionSamples.length - 8) + ' more</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.nameCollisions + ' row(s) share a Site + Name (' + (b.collisionSamples || []).length + ' collision' + ((b.collisionSamples || []).length === 1 ? '' : 's') + ')</div>' +
          '<div class="ts-confirm-issue-body">PickTrace requires a unique name within a site — one block in each pair will be <b>silently dropped</b> on upload. Rename one in the <b>Name Collisions</b> panel first.<ul class="ts-confirm-list">' + sample + more + '</ul></div>' +
          '</div>'
        );
      }
      if (b.pastStartDates > 0) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.pastStartDates + ' Start Date value(s) are before today (' + todayYMD() + ')</div>' +
          '<div class="ts-confirm-issue-body">PickTrace requires Start Date to be today or later. Use the <b>Today</b> button on the Start Date row.</div>' +
          '</div>'
        );
      }

      overlay.innerHTML =
        '<div class="cmp-export-modal-inner ts-confirm-modal">' +
          '<h3>Export with these issues?</h3>' +
          '<div class="ts-confirm-body">' + sections.join('') + '</div>' +
          '<div class="cmp-export-modal-actions">' +
            '<button class="btn btn-warning" id="ts-conf-ok">Export anyway</button>' +
            '<button class="btn btn-ghost" id="ts-conf-cancel">Cancel &amp; fix</button>' +
          '</div>' +
        '</div>';
      document.body.appendChild(overlay);

      const close = result => { overlay.remove(); resolve(result); };
      overlay.querySelector('#ts-conf-ok').addEventListener('click', () => close(true));
      overlay.querySelector('#ts-conf-cancel').addEventListener('click', () => close(false));
      overlay.addEventListener('click', e => { if (e.target === overlay) close(false); });
    });
  }

  // ─── Run / Export ───
  function runFormat() {
    if (!srcData || !tplData) { alert('Upload source data + a personalized bulk template first.'); return; }
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    mapping = autoMap(bulkHeaders, srcData.headers);
    // Reset manual fills + removed rows on a fresh format — they're tied
    // to the prior mapping/target context.
    manualFills = {};
    removedRows = new Set();
    cellOverrides = {};
    smartFixMap = {};
    previewIssuesOnly = false;
    rebuildFormattedRows();
    renderMapping();
    renderPreview();
    renderRequired();
    updateSummary();
    $('ts-empty').style.display = 'none';
  }

  function doExport() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const acreageIdx = bulkHeaders.indexOf('Acreage');
    // Case-canonicalize EVERY dropdown-backed column (Location Type, Production
    // Status, Organic Status, …) — not just Site / Crop & Variety. PickTrace is
    // case-sensitive on dropdown values, so "FIELD" is rejected where the list
    // has "Field" ("Invalid location type: FIELD"). Map upper-cased value →
    // the template's exact casing.
    const dropMaps = bulkHeaders.map(h => buildCaseMap(dropdownValuesFor(h)));
    const exportRows = formattedRows
      .filter((_, ri) => !excluded(ri))
      .map(row => {
        const r = row.slice();
        dropMaps.forEach((m, i) => {
          if (m.size && r[i]) {
            const k = String(r[i]).toUpperCase().trim();
            if (m.has(k)) r[i] = m.get(k);
          }
        });
        // Acreage of 0 exports as blank (per request — a plot with 0 acres
        // shouldn't carry a literal 0).
        if (acreageIdx >= 0 && r[acreageIdx] !== '' && r[acreageIdx] != null) {
          const n = parseFloat(String(r[acreageIdx]).replace(/[, ]/g, ''));
          if (!isNaN(n) && n === 0) r[acreageIdx] = '';
        }
        return r;
      });

    // Open the original template workbook in place so we keep:
    //   – every cell's font / fill / border / number format
    //   – column widths, frozen panes, sheet order
    //   – the DROP-DOWN INPUTS sheet (verbatim, so dropdown lookups still work)
    //   – any data validations (xlsx-js-style preserves them on roundtrip)
    // …then overwrite the DATA ENTRY rows with the user's formatted data.
    const wb = XLSX.read(new Uint8Array(tplData.rawBuffer), {
      type: 'array', cellStyles: true, cellDates: true, sheetStubs: true
    });
    const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[dataName];
    if (!ws) { alert('Template missing DATA ENTRY sheet — cannot export.'); return; }

    // Wipe existing data rows (keep the header row at row 0). We delete only
    // the cell objects, leaving sheet-level properties (!cols, !merges,
    // !dataValidation, !autofilter, etc.) intact.
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: bulkHeaders.length - 1 } };
    for (let r = 1; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) delete ws[ref];
      }
    }

    // Write the formatted rows starting at row 1 (right under the header).
    const colCount = Math.max(range.e.c + 1, bulkHeaders.length);
    exportRows.forEach((row, ri) => {
      for (let c = 0; c < colCount; c++) {
        const val = row[c];
        if (val == null || val === '') continue;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c });
        // Preserve numbers as numbers (so PickTrace's numeric validations work),
        // booleans for Is Archived, everything else as text.
        if (typeof val === 'number') ws[ref] = { v: val, t: 'n' };
        else if (typeof val === 'string' && /^-?\d+(\.\d+)?$/.test(val) &&
                 (bulkHeaders[c] === 'Acreage' || bulkHeaders[c] === 'Plant Count' ||
                  bulkHeaders[c] === 'Length' || bulkHeaders[c] === 'Stand Count' ||
                  bulkHeaders[c] === 'Row/Bed Count' || bulkHeaders[c] === 'Post Count')) {
          ws[ref] = { v: parseFloat(val), t: 'n' };
        } else {
          ws[ref] = { v: String(val), t: 's' };
        }
      }
    });

    // Update !ref to span the new range. Header at row 0; data rows 1..N.
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(0, exportRows.length), c: colCount - 1 }
    });

    // Filename: original template name + " — filled" before the extension.
    // Keeps the user from having to overwrite the source template.
    const origName = (tplData.fileName || 'bulk-template.xlsx').trim();
    const dotIdx = origName.lastIndexOf('.');
    const base = dotIdx > 0 ? origName.substring(0, dotIdx) : origName;
    const ext = dotIdx > 0 ? origName.substring(dotIdx) : '.xlsx';
    const exportName = base + ' — filled' + ext;

    XLSX.writeFile(wb, exportName, { cellStyles: true });
  }

  // ─── Load handlers ───
  function handleSrcFile(file) {
    readSrcFile(file).then(data => {
      srcData = data;
      $('ts-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      const hrowNote = data.headerRow > 0 ? ' (skipped ' + data.headerRow + ' metadata row' + (data.headerRow === 1 ? '' : 's') + ' above)' : '';
      $('ts-src-meta').textContent = data.rows.length + ' rows · ' + data.headers.length + ' columns · sheet: ' + data.sheetName + hrowNote;
      formattedRows = null;
      $('ts-run').disabled = !(srcData && tplData);
      // Auto-run if both already in.
      if (srcData && tplData) runFormat();
    }).catch(err => {
      if (err && err.message === 'cancelled') return; // user dismissed the picker — quiet
      alert('Failed to read file: ' + (err && err.message ? err.message : err));
    });
  }
  // Optional: a PickTrace existing-locations export (bulk Update template /
  // ALL RECORDS). Build a Site + Name index so rows already in PickTrace are
  // dropped from the export (avoids RESOURCE-DUPLICATE).
  function handleExistingFile(file) {
    readSrcFile(file).then(data => {
      const siteI = findColIdxByNames(data.headers, ['site', 'sites', 'grower', 'farm']);
      const nameI = findColIdxByNames(data.headers, ['name', 'block']);
      const cvI = findColIdxByNames(data.headers, ['crop & variety', 'crop and variety', 'crop variety', 'crop']);
      const acI = findColIdxByNames(data.headers, ['acreage', 'acres', 'acre']);
      if (siteI < 0 || nameI < 0) {
        alert('Existing-locations file must have Site and Name columns (a PickTrace bulk Update / locations export).');
        return;
      }
      // keys = Site+Name (is this name taken?). blockKeys = Site+Name+Crop&Variety
      // (is THIS exact block already there?). byName = Site+Name → existing
      // blocks under that name, for showing the conflict.
      const keys = new Set();
      const blockKeys = new Set();
      const byName = new Map();
      data.rows.forEach(r => {
        const s = r[siteI], n = r[nameI];
        if (!(String(s == null ? '' : s).trim() || String(n == null ? '' : n).trim())) return;
        const nk = locKey(s, n);
        const cv = cvI >= 0 ? String(r[cvI] == null ? '' : r[cvI]).trim() : '';
        const ac = acI >= 0 ? String(r[acI] == null ? '' : r[acI]).trim() : '';
        keys.add(nk);
        blockKeys.add(nk + '||' + locNorm(cv));
        const arr = byName.get(nk) || [];
        arr.push({ cv, ac });
        byName.set(nk, arr);
      });
      existingData = { keys, blockKeys, byName, hasCrop: cvI >= 0, fileName: file.name, count: keys.size };
      $('ts-existing-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('ts-existing-meta').textContent = keys.size + ' existing location' + (keys.size === 1 ? '' : 's') + ' indexed' +
        (cvI >= 0 ? ' (block-level)' : ' (name-only — no Crop column)');
      if (srcData && tplData) {
        rebuildFormattedRows();
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      }
    }).catch(err => {
      if (err && err.message === 'cancelled') return;
      alert('Failed to read existing-locations file: ' + (err && err.message ? err.message : err));
    });
  }

  function renderExisting() {
    const sec = $('ts-section-existing');
    if (!sec) return;
    const idxs = formattedRows ? [...existingRowIdxs] : [];
    if (!existingData || !idxs.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*'), nameIdx = bulkHeaders.indexOf('Name*');
    const t = $('ts-existing-title');
    if (t) t.textContent = 'Already in PickTrace — dropped (' + idxs.length + ')';
    let html = '<thead><tr><th>Site</th><th>Name</th></tr></thead><tbody>';
    idxs.slice(0, 200).forEach(ri => {
      const r = formattedRows[ri];
      html += '<tr><td>' + escHtml(r[siteIdx]) + '</td><td>' + escHtml(r[nameIdx]) + '</td></tr>';
    });
    html += '</tbody>';
    $('ts-existing-table').innerHTML = html;
  }

  // ─── Name collisions (same Site + Name, different blocks) ───
  function renderCollisions() {
    const sec = $('ts-section-collisions');
    if (!sec) return;
    const groups = formattedRows ? computeCollisions() : new Map();
    if (!groups.size) { sec.style.display = 'none'; $('ts-collisions-table').innerHTML = ''; return; }
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*');
    const nameIdx = bulkHeaders.indexOf('Name*');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    const acIdx = bulkHeaders.indexOf('Acreage');
    let total = 0; groups.forEach(arr => total += arr.length);
    const t = $('ts-collisions-title');
    if (t) t.textContent = 'Name Collisions (' + groups.size + ' name' + (groups.size === 1 ? '' : 's') + ', ' + total + ' to fix)';
    let html = '<thead><tr><th>Source</th><th>Site</th><th>Name</th><th>Crop &amp; Variety</th><th>Acreage</th>' +
      '<th>New name</th><th></th></tr></thead><tbody>';
    [...groups.entries()].forEach(([k, arr]) => {
      // When the name is already used by a DIFFERENT block in PickTrace, show
      // that existing block first (read-only) so it's clear the name is taken.
      if (existingTakenKeys.has(k) && existingData && existingData.byName && existingData.byName.get(k)) {
        existingData.byName.get(k).forEach(ex => {
          const r0 = formattedRows[arr[0]];
          html += '<tr style="background:var(--bg-sunken,#f3f4f6);color:#6b7280;">' +
            '<td><b>In PickTrace</b></td>' +
            '<td>' + escHtml(r0[siteIdx]) + '</td>' +
            '<td><b>' + escHtml(r0[nameIdx]) + '</b></td>' +
            '<td>' + escHtml(ex.cv) + '</td>' +
            '<td>' + escHtml(ex.ac) + '</td>' +
            '<td colspan="2"><i>already exists — name is taken</i></td>' +
            '</tr>';
        });
      }
      arr.forEach(ri => {
        const r = formattedRows[ri];
        html += '<tr>' +
          '<td>New</td>' +
          '<td>' + escHtml(r[siteIdx]) + '</td>' +
          '<td><b style="color:#dc2626;">' + escHtml(r[nameIdx]) + '</b></td>' +
          '<td>' + escHtml(cvIdx >= 0 ? r[cvIdx] : '') + '</td>' +
          '<td>' + escHtml(acIdx >= 0 ? r[acIdx] : '') + '</td>' +
          '<td><input type="text" class="ts-coll-name input-field" data-ri="' + ri + '" placeholder="' + escHtml(r[nameIdx]) + '" style="width:140px;"></td>' +
          '<td><button class="btn btn-primary btn-sm ts-coll-apply" data-ri="' + ri + '">Rename</button></td>' +
          '</tr>';
      });
    });
    html += '</tbody>';
    $('ts-collisions-table').innerHTML = html;
    const tbl = $('ts-collisions-table');
    tbl.querySelectorAll('.ts-coll-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const tr = e.target.closest('tr');
        const ri = +e.target.dataset.ri;
        const val = tr.querySelector('.ts-coll-name').value.trim();
        if (!val) { alert('Type a new name first.'); return; }
        applyRename(ri, val);
      });
    });
  }

  function applyRename(ri, newName) {
    nameOverrides[ri] = newName;
    rebuildFormattedRows();
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
  }

  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      const parsed = parseTemplate(e.target.result, file.name);
      if (!parsed || !parsed.headers.length) {
        alert('Could not find DATA ENTRY sheet in this template.');
        return;
      }
      tplData = parsed;
      $('ts-tpl-name').textContent = file.name;
      $('ts-tpl-meta').textContent =
        parsed.headers.length + ' columns · ' +
        (parsed.dropdowns.get('site') ? parsed.dropdowns.get('site').size : 0) + ' sites · ' +
        (parsed.dropdowns.get('crop & variety') ? parsed.dropdowns.get('crop & variety').size : 0) + ' C&V values';
      // Detect target by header count if user hasn't picked.
      const hasArch = parsed.headers.some(h => norm(h) === 'is archived');
      const radioVal = document.querySelector('input[name="ts-target"]:checked').value;
      if (hasArch && radioVal === 'create') {
        document.querySelector('input[name="ts-target"][value="update"]').checked = true;
        target = 'update';
      } else if (!hasArch && radioVal === 'update') {
        document.querySelector('input[name="ts-target"][value="create"]').checked = true;
        target = 'create';
      }
      formattedRows = null;
      $('ts-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    };
    r.readAsArrayBuffer(file);
  }

  // ─── Org-Data import — reuses Block Compare's aggregator if available ───
  function importFromOrgData() {
    if (typeof aggregateOrgData !== 'function' || typeof getAllOrgNames !== 'function') {
      alert('Org Data store unavailable. Reload the page.');
      return;
    }
    const names = getAllOrgNames();
    if (!names.length) { alert('No saved orgs found.'); return; }
    const pick = prompt('Type the org name to import:\n\n' + names.join('\n'));
    if (!pick) return;
    const found = names.find(n => n.toLowerCase() === pick.trim().toLowerCase());
    if (!found) { alert('Org "' + pick + '" not found.'); return; }
    const aggregated = aggregateOrgData(found);
    if (!aggregated || !aggregated.rows.length) { alert('No rows for "' + found + '".'); return; }
    srcData = {
      headers: aggregated.headers,
      rows: aggregated.rows,
      fileName: '(Org Data: ' + found + ')',
      sheetName: '(synthesized)'
    };
    $('ts-src-name').textContent = '(Org Data: ' + found + ')';
    $('ts-src-meta').textContent = aggregated.rows.length + ' rows · ' + aggregated.headers.length + ' columns';
    formattedRows = null;
    $('ts-run').disabled = !(srcData && tplData);
    if (srcData && tplData) runFormat();
  }

  // ─── Reset ───
  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    manualFills = {}; removedRows = new Set(); siteAddresses = {};
    smartFixMap = {}; previewIssuesOnly = false;
    dateFill = { mm: '', dd: '', range: 'first' };
    existingData = null; existingRowIdxs = new Set();
    nameOverrides = {}; cellOverrides = {}; collisionKeys = new Set(); existingTakenKeys = new Set();
    $('ts-src-name').textContent = 'No file selected';
    $('ts-tpl-name').textContent = 'No file selected';
    $('ts-src-meta').textContent = '';
    $('ts-tpl-meta').textContent = '';
    { const en = $('ts-existing-name'); if (en) en.textContent = 'No file selected'; }
    { const em = $('ts-existing-meta'); if (em) em.textContent = ''; }
    { const ef = $('ts-existing-file'); if (ef) ef.value = ''; }
    $('ts-src-file').value = ''; $('ts-tpl-file').value = '';
    ['ts-section-mapping','ts-section-smartfix','ts-section-sites-create','ts-section-cv-create','ts-section-required','ts-section-existing','ts-section-collisions','ts-section-datefill','ts-section-preview']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    $('ts-summary').style.display = 'none';
    $('ts-empty').style.display = '';
    $('ts-run').disabled = true;
    $('ts-export').disabled = true;
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    initialized = true;
    attachModeSwitcher();
    $('ts-src-file').addEventListener('change', e => {
      if (e.target.files[0]) handleSrcFile(e.target.files[0]);
      e.target.value = '';
    });
    $('ts-tpl-file').addEventListener('change', e => {
      if (e.target.files[0]) handleTplFile(e.target.files[0]);
      e.target.value = '';
    });
    $('ts-src-from-org').addEventListener('click', importFromOrgData);
    { const ef = $('ts-existing-file'); if (ef) ef.addEventListener('change', e => { if (e.target.files[0]) handleExistingFile(e.target.files[0]); e.target.value = ''; }); }
    { const cs = $('ts-crop-split'); if (cs) cs.addEventListener('change', e => {
        cropSplit = !!e.target.checked;
        if (srcData && tplData) { rebuildFormattedRows(); renderPreview(); renderRequired(); renderSitesAndCvToCreate(); updateSummary(); }
      }); }
    { const a = $('ts-datefill-apply'); if (a) a.addEventListener('click', applyDateFill); }
    { const c = $('ts-datefill-clear'); if (c) c.addEventListener('click', clearDateFill); }
    document.querySelectorAll('input[name="ts-target"]').forEach(r => {
      r.addEventListener('change', e => {
        target = e.target.value;
        if (srcData && tplData) runFormat();
      });
    });
    $('ts-run').addEventListener('click', runFormat);
    $('ts-export').addEventListener('click', async () => {
      const b = getExportBlockers();
      if (b && (b.emptyReq > 0 || b.unknownSites.length || b.unknownCvs.length || b.nameCollisions > 0 || b.pastStartDates > 0)) {
        const proceed = await showExportConfirm(b);
        if (!proceed) return;
      }
      doExport();
    });
    $('ts-reset').addEventListener('click', reset);
    { const sfa = $('ts-smartfix-apply-all'); if (sfa) sfa.addEventListener('click', applyAllSmartFixes); }
    $('ts-sites-copy').addEventListener('click', () => {
      const tbl = $('ts-sites-create-table');
      const rows = [...tbl.querySelectorAll('tbody tr')]
        .filter(tr => !tr.classList.contains('ts-addr-allrow'))
        .map(tr => {
          const val  = tr.children[0] ? tr.children[0].textContent.trim() : '';
          const cnt  = tr.children[1] ? tr.children[1].textContent.trim() : '';
          const aInp = tr.querySelector('.ts-site-addr');
          return val + '\t' + cnt + '\t' + (aInp ? aInp.value.trim() : '');
        });
      if (!rows.length) return;
      navigator.clipboard.writeText('Site\tRow count\tAddress\n' + rows.join('\n'))
        .then(() => alert('Copied ' + rows.length + ' sites.'), () => alert('Copy failed.'));
    });
    // Click-to-copy any cell (mirrors the Compare page). The Empty Required
    // Columns section is form-only, and the address control row has no value
    // worth copying, so both are excluded.
    $('ts-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      if (!td) return;
      if (td.closest('#ts-section-required')) return;
      if (td.closest('tr.ts-addr-allrow')) return;
      if (e.target.closest('input, button, label, select, textarea, a')) return;
      const text = (td.textContent || '').trim();
      if (!text) return;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
          .then(() => td.classList.toggle('cmp-copied'), () => {});
      }
    });
    $('ts-cv-copy').addEventListener('click', () => {
      const tbl = $('ts-cv-create-table');
      const rows = [...tbl.querySelectorAll('tbody tr')]
        .map(tr => [...tr.children].map(td => td.textContent).join('\t'));
      if (!rows.length) return;
      navigator.clipboard.writeText('Crop & Variety\tRow count\n' + rows.join('\n'))
        .then(() => alert('Copied ' + rows.length + ' values.'), () => alert('Copy failed.'));
    });
  }

  // Init when the parent page is opened.
  window.tsInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
