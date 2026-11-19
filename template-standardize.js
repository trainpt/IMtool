// ═══════════════════════════════════════════════════════════════════════
// Template Standardize — formats source rows (Excel/CSV) into
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
  let siteAddresses = {};      // new site → its own address ("street, city, ST zip"); survives re-renders
  // Sites file — the new sites, built into a Sites upload next to this one (see "Sites file" below).
  let siteTpl = null;          // the org's Sites template (parseTemplate); none → no Sites file
  let siteFill = {};           // Sites column key (norm, no *) → value for every new site
  let siteAddrAll = '';        // address for every new site
  let siteAddrGroup = {};      // ranch (site name before " - ") → address
  let siteEdits = {};          // "site|column key" → value typed in the Sites preview
  let pvView = 'locations';    // which file the preview shows: 'locations' | 'sites'
  let siteGrid = [];           // site names in the Sites preview's row order, for its edits
  let dateFill = { mm: '', dd: '', range: 'first' }; // year-only → full-date expansion
  let existingData = null;        // { keys: Set<siteName>, fileName, count } — locations already in PickTrace
  let builtOnce = false;          // Format Sheet pressed — from then on a file swap rebuilds by itself
  let existingRowIdxs = new Set();// formattedRows indexes that already exist in PickTrace (dropped)
  let cropSplit = true;           // split " - Subvariety" out of Crop & Variety → Clone/Subvariety
  let nameOverrides = {};         // formattedRows index → renamed Name* (resolves collisions, sticky)
  let cellOverrides = {};         // "rowIdx|colIdx" → value — sticky per-cell edits from the editable preview (all columns EXCEPT Name*, which uses nameOverrides)
  let smartFixMap = {};           // "colIdx||UPPERVALUE" → canonical dropdown value — sticky bulk fixes the user applied from the Smart Fixes panel
  let siteFillOk = new Set();     // "<first row>|<site>" of each Blank Site group you confirmed copying down
  let siteGaps = [];              // Blank Site groups found on the last rebuild (see rebuildFormattedRows)
  let smartFixOpts = new Map();   // colIdx → allowed values, for the Smart Fix pickers opened on click
  let pvMode = 'all';             // preview rows: 'all' | 'fix' | 'chg' | 'any' (IMWorkbench.filterBar)
  let pvColsOnly = false;         // preview: only the columns with a flag or a change
  let pvFocus = null;             // { ci } — the preview shows just the rows with this column empty
  let collisionKeys = new Set();  // locKeys flagged as collisions (set in rebuildFormattedRows)
  let existingTakenKeys = new Set(); // locKeys whose name is already used by a DIFFERENT block in PickTrace
  let mergeDupes = true;          // collapse duplicate Site + Name rows into one location
  let mergedRowIdxs = new Set();  // formattedRows indexes folded into a survivor (dropped from preview + export)
  let mergeGroups = new Map();    // locKey → { keep, folded: [ri…], conflicts: Map<colIdx,[values]> }
  let unmergedKeys = new Set();   // locKeys the user chose to keep expanded (sticky across rebuilds)
  let initialized = false;

  // A row is excluded from preview + export if the user removed it, it already
  // exists in PickTrace (cross-referenced from the uploaded existing-locations
  // file), or it was folded into a surviving duplicate.
  function excluded(i) { return removedRows.has(i) || existingRowIdxs.has(i) || mergedRowIdxs.has(i); }

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

  // ─── Numeric columns ───
  // PickTrace's numeric columns must hold bare numbers, but source sheets carry
  // the unit inline ("20.30 ac", "3500 Vines", "1,240 acres"), which would
  // otherwise export as text and fail the numeric validation. Pull the first
  // number out, ignoring grouping commas and any unit prefix or suffix. Returns
  // null when there's no number at all.
  function parseNum(v) {
    const s = String(v == null ? '' : v).replace(/,/g, '').trim();
    if (!s) return null;
    const m = s.match(/-?\d+(?:\.\d+)?/);
    if (!m) return null;
    const n = parseFloat(m[0]);
    return isFinite(n) ? n : null;
  }
  // Bulk columns that must reach PickTrace as a bare number.
  const NUMERIC_HEADERS = new Set(['Acreage', 'Length', 'Plant Count', 'Stand Count',
    'Row/Bed Count', 'Post Count', 'Percent Covered', 'Row Spacing, in.',
    'Plant Spacing, in.', 'Post Spacing, in.', 'Bed Width, in.']);
  // Acreage alone blanks out at 0 — a plot with 0 acres shouldn't carry a
  // literal 0. Every other numeric column keeps a legitimate zero.
  const BLANK_ON_ZERO = new Set(['Acreage']);
  // Canonical numeric cell: the bare number as a string, or '' when it's absent
  // (or zero, for the columns that blank at zero).
  function numericCell(v, header) {
    const n = parseNum(v);
    if (n == null) return '';
    if (n === 0 && BLANK_ON_ZERO.has(header)) return '';
    return String(n);
  }
  // Indexes of the numeric columns in the active schema.
  function numericColIdxs() {
    const bulkHeaders = activeSchemaHeaders();
    return bulkHeaders.map((h, i) => NUMERIC_HEADERS.has(h) ? i : -1).filter(i => i >= 0);
  }

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
    return bulkHeaders.map((h, i) => dropdownValuesFor(h) && restricted(h) ? i : -1).filter(i => i >= 0);
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
    // New sites going out in the Sites file count as listed: uploaded first,
    // they exist by the time the Locations file loads.
    // Sites the existing export shows are real too, listed or not.
    if (key === 'site' && (siteFileOn() || existingData)) {
      const all = [...knownSites()].concat(siteFileOn() ? [...newSitesMap().keys()] : []);
      return all.length ? all : null;
    }
    return set && set.size ? [...set] : null;
  }

  // A column the template's DROP-DOWN INPUTS names but leaves empty: the org
  // has none set up yet (Rootstock, Clone/Subvariety, …), and PickTrace only
  // takes listed values there — anything in it must be created first. Site
  // and Crop & Variety have flows of their own (Sites file, Build in PickTrace).
  function emptyListFor(bulkHeader) {
    if (!tplData || !tplData.listed) return false;
    const key = norm(bulkHeader).replace(/\*+$/, '').trim();
    if (key === 'site' || key === 'crop & variety') return false;
    return restricted(bulkHeader) && tplData.listed.has(key) && !(tplData.dropdowns.get(key) || new Set()).size;
  }

  // ─── Which columns PickTrace restricts ───
  // The template puts a list validation on the DATA ENTRY columns that only
  // take listed values (Site, Location Type, Crop & Variety, Production
  // Status, Training Style, Trellis Type, Row Direction, Organic Status). Its
  // DROP-DOWN INPUTS sheet also lists free-text columns — Rootstock,
  // Clone/Subvariety, Location Group, Mulch Type… — with the values the org
  // already uses; those take any value. Site and Crop & Variety are always
  // checked; when the file can't say (no validations read), every listed
  // column counts as restricted.
  function restricted(bulkHeader) {
    const key = norm(bulkHeader).replace(/\*+$/, '').trim();
    if (key === 'site' || key === 'crop & variety') return true;
    return !(tplData && tplData.restricted) || tplData.restricted.has(key);
  }
  // SheetJS doesn't read validations, so they come from the sheet XML: the
  // worksheet whose list validations point at the DROP-DOWN INPUTS sheet.
  // → Set of header keys, or null when there's nothing to read.
  async function readRestricted(buf, headers) {
    if (!window.IMXlsxTemplate) return null;
    try {
      for (const e of IMXlsxTemplate.zipRead(buf)) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name)) continue;
        const xml = await IMXlsxTemplate.entryText(e);
        if (!/<dataValidation\b/.test(xml) || !/DROP-?DOWN/i.test(xml)) continue;
        const set = new Set();
        (xml.match(/<dataValidation\b[^>]*>/g) || []).forEach(tag => {
          if (!/type="list"/.test(tag)) return;
          ((tag.match(/sqref="([^"]*)"/) || [])[1] || '').split(/\s+/).forEach(ref => {
            const col = (ref.match(/^([A-Z]+)/) || [])[1];
            const h = col ? headers[XLSX.utils.decode_col(col)] : '';
            if (h) set.add(norm(h).replace(/\*+$/, '').trim());
          });
        });
        return set.size ? set : null;
      }
    } catch (err) { /* unreadable — fall back to every listed column */ }
    return null;
  }
  // Column index → Map(value → rows) for those columns, case variants as one.
  function unlistedValues() {
    const out = new Map();
    if (!formattedRows || !tplData) return out;
    activeSchemaHeaders().forEach((h, ci) => {
      if (!emptyListFor(h)) return;
      const m = new Map(), first = new Map();
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        const v = String(r[ci] == null ? '' : r[ci]).trim();
        if (!v) return;
        const k = first.get(v.toUpperCase()) || v;
        first.set(v.toUpperCase(), k);
        m.set(k, (m.get(k) || 0) + 1);
      });
      if (m.size) out.set(ci, m);
    });
    return out;
  }
  function renderListsCreate() {
    const sec = $('ts-section-lists-create');
    if (!sec) return 0;
    const lists = unlistedValues();
    if (!lists.size) { sec.style.display = 'none'; return 0; }
    sec.style.display = '';
    const H = activeSchemaHeaders();
    let n = 0, html = '<thead><tr><th>Column</th><th>Value</th><th>Row count</th></tr></thead><tbody>';
    lists.forEach((m, ci) => {
      const vals = [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
      n += vals.length;
      vals.forEach(([v, c], k) => {
        html += '<tr>' + (k === 0 ? '<td rowspan="' + vals.length + '" data-copy="' + escHtml(H[ci]) + '"><b>' + escHtml(H[ci]) + '</b>' +
          '<div class="text-muted small">' + vals.length + ' value' + (vals.length === 1 ? '' : 's') + ' to create</div>' +
          '<button type="button" class="btn btn-ghost btn-sm ts-list-drop" data-ci="' + ci + '" title="Export this column empty instead (Column Mapping can bring it back)">Leave column out</button></td>' : '') +
          '<td class="ts-copy-cell" title="Click to copy">' + escHtml(v) + '</td><td>' + c + '</td></tr>';
      });
    });
    $('ts-lists-create-title').textContent = 'Lists your org hasn’t set up (' + n + ' value' + (n === 1 ? '' : 's') + ')';
    $('ts-lists-create-table').innerHTML = html + '</tbody>';
    return n;
  }

  // Pane visibility is owned by showBuild() in index.html's router — this
  // module no longer switches modes. Kept as a no-op so the init call below
  // (and the matching ones in the other builder modules) stay uniform.
  function attachModeSwitcher() {}

  // ─── Header-row detection ───
  // PickTrace implementation templates have 3-5 rows of metadata
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
  // A "(Detailed Version)" tab names Site in the banner row and carries
  // metadata rows under the header — IMSheetLayout (utils.js) undoes both.
  // The data ends at the template's "*add/delete rows as necessary" footer;
  // what follows is PickTrace's contact block, not locations.
  const FOOTER_RE = /^\*?\s*add\s*\/\s*delete rows/i;
  // `rowAt[i]` is data row i's index in `aoa` (for reading its cell styles).
  function sliceSheetAoa(aoa) {
    const hIdx = detectHeaderRow(aoa);
    const lay = IMSheetLayout.resolve(aoa, hIdx);
    let body = aoa.slice(hIdx + 1).map((r, k) => ({ r, at: hIdx + 1 + k }));
    const end = body.findIndex(x => (x.r || []).some(c => FOOTER_RE.test(String(c == null ? '' : c).trim())));
    if (end >= 0) body = body.slice(0, end);
    const kept = body.filter(x => !isInstructionRow(x.r) && !lay.isMeta(x.r));
    const rows = kept.map(x => x.r.map(cellText));
    return { headers: lay.headers, rows, headerRow: hIdx, rowAt: kept.map(x => x.at) };
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
  // the raw AOA. PickTrace implementation templates
  // have multi-row metadata above the real column header — readUploadedFile
  // would treat row 0 as the header and miss everything.
  function readSrcFile(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const buf = new Uint8Array(e.target.result);
          const wb = XLSX.read(buf, { type: 'array' });
          // A second read with cell styles, only if a sheet carries column
          // instructions (column-directives.js) — to see its red columns.
          let styled;
          const styledWb = () => {
            if (styled === undefined) {
              try { styled = XLSX.read(buf, { type: 'array', cellStyles: true }); } catch (err) { styled = null; }
            }
            return styled;
          };
          const sheets = wb.SheetNames.map(n => {
            const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' });
            const orig = [];
            const filtered = aoa.filter((row, i) => {
              const keep = row && row.some(c => c != null && String(c).trim() !== '');
              if (keep) orig.push(i);
              return keep;
            });
            if (!filtered.length) return null;
            const sliced = sliceSheetAoa(filtered);
            if (!sliced.headers.filter(h => h).length) return null;
            const sheet = {
              name: n,
              headers: sliced.headers,
              rows: sliced.rows,
              headerRow: sliced.headerRow
            };
            return IMDirectives.attach(sheet, { styledWb, sheetName: n, ws: wb.Sheets[n],
              headerAt: orig[sliced.headerRow], rowAt: sliced.rowAt.map(at => orig[at]) });
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({
            headers: s.headers, rows: s.rows, fileName: file.name,
            sheetName: s.name, headerRow: s.headerRow, directives: s.directives || null
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
    const listed = new Set(); // every list the sheet names, empty or not
    if (dropName) {
      const dropWs = wb.Sheets[dropName];
      const raw = XLSX.utils.sheet_to_json(dropWs, { header: 1, defval: '' });
      (raw[0] || []).forEach(h => { if (String(h || '').trim()) listed.add(norm(String(h).trim())); });
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
    return { headers, dropdowns, listed, fileName, rawBuffer: buf, sheetNames: wb.SheetNames };
  }

  // ─── Template identification ───
  // This module builds the Locations bulk template only — Sites and Employees
  // each have their own tool. Loading the wrong template silently "works":
  // rows are written positionally under whatever header row the file carries,
  // so locations data poured into the Sites template puts Location Type under
  // Employer*, Crop & Variety under Address1*, Acreage under City* and Start
  // Date under Country*, and the result fails validation on upload. Identify
  // the file on load so that never reaches an export.
  //
  // Signature columns are matched asterisk- and case-insensitively. Sites and
  // Locations both have Name*, so each signature includes columns the other
  // template doesn't have.
  const TEMPLATE_KINDS = [
    { kind: 'locations', label: 'Locations',
      need: ['site', 'name', 'location type', 'crop & variety'] },
    { kind: 'sites', label: 'Sites',
      need: ['site type', 'employer', 'address1'] },
    { kind: 'employees', label: 'Employees',
      need: ['first name', 'last name', 'date of birth'] }
  ];
  // → { kind, label }; kind 'unknown' when no signature matches.
  function identifyTemplate(headers) {
    const have = new Set((headers || []).map(h => norm(h).replace(/\*+$/, '').trim()).filter(Boolean));
    for (const t of TEMPLATE_KINDS) if (t.need.every(n => have.has(n))) return t;
    return { kind: 'unknown', label: 'unrecognized' };
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
    'clone/subvariety':   ['clone','subvariety','sub variety','sub_variety','sub-variety'],
    // Implementation Data Template names (the Detailed Version spells Mulch
    // "Multch" and calls the custom fields "Loc Data").
    'mulch type':         ['mulch type','multch type','mulch'],
    'custom data 1':      ['loc data 1'],
    'custom data 2':      ['loc data 2'],
    'custom data 3':      ['loc data 3'],
    'production status':  ['producing'],
    'production start':   ['producing start year','production start year']
  };

  // Bulk headers where the substring fallback (Phase 4) is unsafe — a source
  // header could match by accident (e.g., "Variety" matching "Clone/Subvariety"
  // because "clonesubvariety".includes("variety") is true). For these we
  // accept ONLY exact-match + loose-alphanumeric + alias hits.
  const STRICT_HEADERS = new Set(['clone/subvariety']);
  // Loose normalization — strip all non-alphanumeric chars so "Plant/Seed Seller",
  // "Plant Seed Seller", and "plant-seed_seller" all collapse to "plantseedseller".
  const normLoose = h => norm(h).replace(/[^a-z0-9]/g, '');

  // Names a client-annotated source may use for a field that the aliases
  // above deliberately leave out (see column-directives.js). Consulted only
  // for a column instruction, never for an ordinary header.
  const DIRECTIVE_EXTRAS = { 'clone/subvariety': ['crop sub variety', 'crop subvariety'] };

  // Mapping for the loaded source: its column instructions first when it
  // carries them, otherwise the plain header auto-map.
  function mapSource(bulkHeaders) {
    if (!IMDirectives.on(srcData)) return autoMap(bulkHeaders, srcData.headers);
    return IMDirectives.buildMapping({
      src: srcData, bulkHeaders, srcHeaders: srcData.headers, autoMap,
      aliases: ALIASES, extras: DIRECTIVE_EXTRAS, prefix: 'location',
      hasData: c => srcData.rows.some(r => r && r[c] != null && String(r[c]).trim() !== '')
    }).mapping;
  }

  function autoMap(bulkHeaders, srcHeaders) {
    const srcIdx = {};         // norm(header) → idx (case+trim+lowercase only)
    const srcIdxLoose = {};    // normLoose(header) → idx (alphanumeric only)
    srcHeaders.forEach((h, i) => {
      const k = norm(h);
      const kLoose = normLoose(h);
      if (k && srcIdx[k] == null) srcIdx[k] = i;
      if (kLoose && srcIdxLoose[kLoose] == null) srcIdxLoose[kLoose] = i;
    });
    // A header carrying a bracketed note ("Plant Quantity (Tree/Vine Quantity)")
    // is also indexed by the words outside the brackets — consulted after the
    // exact and loose matches, alongside the aliases.
    const srcIdxBare = {};
    srcHeaders.forEach((h, i) => {
      if (!/\(/.test(String(h || ''))) return;
      const kb = normLoose(String(h).replace(/\([^)]*\)/g, ''));
      if (kb && srcIdxBare[kb] == null) srcIdxBare[kb] = i;
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
      // 3. Aliases — and both, against bracket-stripped source headers.
      const aliases = ALIASES[bn] || ALIASES[bnPlain];
      if (aliases) {
        for (const a of aliases) {
          if (srcIdx[a] != null) { out[bi] = srcIdx[a]; return; }
          const aLoose = normLoose(a);
          if (srcIdxLoose[aLoose] != null) { out[bi] = srcIdxLoose[aLoose]; return; }
        }
      }
      for (const k of [bnLoosePlain].concat((aliases || []).map(normLoose))) {
        if (srcIdxBare[k] != null) { out[bi] = srcIdxBare[k]; return; }
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
    IMDirectives.renderBanner(sec, srcData, bulkHeaders, runFormat);
    const srcOpts = ['<option value="-1">— (leave empty) —</option>']
      .concat(srcData.headers.map((h, i) => '<option value="' + i + '">' + escHtml(h) + '</option>'))
      .join('');
    // Separate Crop Type + Variety columns are joined into Crop & Variety by
    // buildFormattedRowsRaw, whatever that row's dropdown says — so the row
    // says so, with a joined sample, instead of showing a dropdown that does nothing.
    const cropI = findSrcColByNames(['crop', 'crop type', 'crop name', 'plant_name', 'plant name']);
    const varI = findSrcColByNames(['variety', 'varietal'], ['clone', 'subvariety', 'sub variety', 'sub-variety']);
    const joined = cropI >= 0 && varI >= 0 && cropI !== varI;
    const cell = (r, i) => String(r && r[i] != null ? r[i] : '').trim();
    let html = '<thead><tr><th>Template column</th><th>Source column</th><th>Sample value</th><th></th></tr></thead><tbody>';
    bulkHeaders.forEach((bh, bi) => {
      const required = /\*$/.test(bh);
      const sel = mapping[bi] != null ? mapping[bi] : -1;
      const isCv = joined && bh === 'Crop & Variety*';
      const r0 = srcData.rows[0];
      const sample = isCv ? [cell(r0, cropI), cell(r0, varI)].filter(Boolean).join('-')
        : sel >= 0 && r0 ? cell(r0, sel) : '';
      html += '<tr>' +
        '<td>' + (required ? '<b>' + escHtml(bh) + '</b>' : escHtml(bh)) + '</td>' +
        (isCv
          ? '<td><span class="ts-map-joined">' + escHtml(srcData.headers[cropI]) + ' + ' + escHtml(srcData.headers[varI]) + ' (joined)</span></td>'
          : '<td><select class="ts-map-select input-field" data-bi="' + bi + '" style="min-width:240px;">' +
            srcOpts.replace('value="' + sel + '"', 'value="' + sel + '" selected') + '</select></td>') +
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
      if (IMDirectives.ignoredCol(srcData, i)) return;     // the client marked it to skip
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
    const numIdxs = numericColIdxs();
    // The Detailed Version's "Producing" values (Full Production / Not Full
    // Production / No) → the dropdown's Full / Partial / None — the same rule
    // Bulk Template Standardize applies (dtInternals.autoMatch).
    const psIdx = bulkHeaders.indexOf('Production Status');
    const psList = psIdx >= 0 ? dropdownValuesFor('Production Status') : null;
    const dt = window.dtInternals || null;

    return srcData.rows.map(srcRow => {
      const out = new Array(bulkHeaders.length).fill('');
      bulkHeaders.forEach((bh, bi) => {
        const si = mapping[bi];
        if (si != null && si >= 0 && srcRow[si] != null) {
          const v = String(srcRow[si]).trim();
          if (v) out[bi] = v;
        }
      });
      // Strip units off every numeric column so they land as bare numbers
      // ("20.30 ac" → "20.3", "3500 Vines" → "3500"); 0 acres blanks entirely.
      numIdxs.forEach(ci => { if (out[ci] !== '') out[ci] = numericCell(out[ci], bulkHeaders[ci]); });
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
      if (psList && dt && out[psIdx] && !psList.some(o => o.toUpperCase() === out[psIdx].toUpperCase())) {
        const to = dt.autoMatch('Production Status', out[psIdx], psList);
        if (to) out[psIdx] = to;
      }
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
    // Blank Site under a row that has one: the client wrote each site once for
    // its group. Offered in Smart Fixes and copied down only for the groups
    // you confirm — before column fills, so a confirmed site wins over a
    // blanket fill. Keyed by first row + site, so a remap never re-applies an
    // old confirmation to a different group.
    siteGaps = [];
    const heads = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteI = heads.indexOf('Site*'), nameI = heads.indexOf('Name*');
    if (siteI >= 0) {
      let last = '', cur = null;
      formattedRows.forEach((r, ri) => {
        if (r[siteI]) { last = r[siteI]; cur = null; return; }
        if (!last) return;
        if (!cur) { cur = { start: ri, site: last, rows: [], names: [] }; siteGaps.push(cur); }
        cur.rows.push(ri);
        if (nameI >= 0 && r[nameI] && cur.names.length < 3) cur.names.push(r[nameI]);
      });
      siteGaps.forEach(g => {
        g.id = g.start + '|' + g.site;
        g.applied = siteFillOk.has(g.id);
        if (g.applied) g.rows.forEach(ri => { formattedRows[ri][siteI] = g.site; });
      });
    }
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
    // Collapse duplicate Site + Name rows into one location. Runs BEFORE
    // existing-detection and collision-detection so folded rows never read as
    // collisions and never get counted against the existing-locations file.
    computeMerge(siteIdx, nameIdx);
    const rowNameKey = ri => locKey(formattedRows[ri][siteIdx], formattedRows[ri][nameIdx]);
    const rowBlockKey = ri => rowNameKey(ri) + '||' + locNorm(cvIdxR >= 0 ? formattedRows[ri][cvIdxR] : '');
    // Count how many source rows share each Site + Name (a within-file collision
    // protects against being silently dropped by a name-only existing file).
    const withinCounts = new Map();
    if (siteIdx >= 0 && nameIdx >= 0) formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri) || mergedRowIdxs.has(ri)) return;
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
    // A Bulk Update is made of locations PickTrace already has, so nothing is
    // dropped for being there (nor flagged below for taking its name).
    if (target !== 'update' && existingData && existingData.keys && existingData.keys.size && siteIdx >= 0 && nameIdx >= 0) {
      const useBlock = !!(existingData.hasCrop && existingData.blockKeys);
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri) || mergedRowIdxs.has(ri)) return;
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
        if (excluded(ri)) return;
        const s = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
        const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
        if (!s || !n) return;
        const kk = rowNameKey(ri);
        counts.set(kk, (counts.get(kk) || 0) + 1);
        if (target !== 'update' && existingData && existingData.keys && existingData.keys.has(kk)) {
          collisionKeys.add(kk);
          existingTakenKeys.add(kk);
        }
      });
      counts.forEach((c, kk) => { if (c >= 2) collisionKeys.add(kk); });
    }
  }

  // ─── Duplicate collapsing ───
  // Implementation Data Templates routinely list one row per (unnamed) block
  // inside a ranch, so the same Site + Name repeats with only Acreage varying —
  // a Locations tab of 6,368 rows describing 537 real locations is typical.
  // PickTrace needs exactly one row per location, so fold each duplicate group
  // into a single survivor (the first row in source order):
  //   • Acreage   → the largest value in the group
  //   • all other → the first non-empty value in the group
  // A column where two rows hold DIFFERENT non-empty values is recorded as a
  // conflict and surfaced in the Merged Duplicates panel rather than silently
  // resolved. Site and Name are the group key, so they're never touched.
  //
  // Rows with a blank Site or Name never merge (there's no identity to merge
  // on) and keys in unmergedKeys are left expanded, which is what the panel's
  // Unmerge button toggles. Fully re-derived on every rebuild, like cropSplit
  // and dateFill, so turning the feature off restores every folded row.
  function computeMerge(siteIdx, nameIdx) {
    mergedRowIdxs = new Set();
    mergeGroups = new Map();
    if (!mergeDupes || !formattedRows || siteIdx < 0 || nameIdx < 0) return;
    const bulkHeaders = activeSchemaHeaders();
    const acreageIdx = bulkHeaders.indexOf('Acreage');

    // Bucket every still-live row by Site + Name, preserving source order.
    const buckets = new Map();
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const s = String(r[siteIdx] == null ? '' : r[siteIdx]).trim();
      const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!s || !n) return;
      const k = locKey(s, n);
      const arr = buckets.get(k) || [];
      arr.push(ri);
      buckets.set(k, arr);
    });

    buckets.forEach((idxs, k) => {
      if (idxs.length < 2 || unmergedKeys.has(k)) return;
      const keep = idxs[0];
      const survivor = formattedRows[keep];
      const conflicts = new Map();
      for (let c = 0; c < bulkHeaders.length; c++) {
        if (c === siteIdx || c === nameIdx) continue;
        if (c === acreageIdx) {
          let max = null;
          idxs.forEach(ri => {
            const n = parseNum(formattedRows[ri][c]);
            if (n != null && (max == null || n > max)) max = n;
          });
          survivor[c] = (max == null || max === 0) ? '' : String(max);
          continue;
        }
        const seen = [];
        idxs.forEach(ri => {
          const v = String(formattedRows[ri][c] == null ? '' : formattedRows[ri][c]).trim();
          if (v && !seen.includes(v)) seen.push(v);
        });
        survivor[c] = seen.length ? seen[0] : '';
        if (seen.length > 1) conflicts.set(c, seen);
      }
      const folded = idxs.slice(1);
      folded.forEach(ri => mergedRowIdxs.add(ri));
      mergeGroups.set(k, { keep, folded, conflicts });
    });
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
      if (excluded(ri)) return;
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
    renderMerged();
    renderCollisions();
    renderSmartFixes();
    renderDateFill();
    const sec = $('ts-section-preview');
    const tbl = $('ts-preview-table');
    sec.style.display = '';
    if (pvView === 'sites' && !(siteFileOn() && newSitesMap().size)) pvView = 'locations';
    if (pvView === 'sites') { renderSitesPreview(sec, tbl); updateExportButton(); return; }
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
    // What the preview flags in a cell — red: empty required or a name
    // collision; amber: past Start Date, year-only date, off-list dropdown value.
    const RED = '#fee2e2', AMBER = '#fef3c7';
    const unlisted = new Set(bulkHeaders.map((h, i) => emptyListFor(h) ? i : -1).filter(i => i >= 0));
    const flagOf = (ri, i) => {
      const val = formattedRows[ri][i] == null ? '' : String(formattedRows[ri][i]);
      const h = bulkHeaders[i], d = colDrop[i];
      if (val && unlisted.has(i)) return { bg: AMBER, title: 'Your org has no ' + h + ' options set up yet — create "' + val + '" in PickTrace first (see Build in PickTrace), or leave the column out.' };
      if (/\*$/.test(h) && !val) return { bg: RED, title: 'Required — must be filled before export' };
      if (i === startDateIdx && val && isDateBeforeToday(val)) return { bg: AMBER, title: 'Start Date must be today (' + todayYMD() + ') or later' };
      if (dateIdxSet.has(i) && isYearOnlyDate(val)) return { bg: AMBER, title: 'Year only — set a month/day in Year-Only Dates above to make this a full YYYY-MM-DD date' };
      if (i === nameColIdx && collisionIdxs.has(ri)) return { bg: RED, title: 'Name collision — another block at this Site shares this Name. Rename it here (or in Name Collisions above), or PickTrace will drop one.' };
      if (d && val && restricted(h) && !d.lower.has(val.toLowerCase().trim())) return { bg: AMBER, title: 'Off-list — "' + val + '" isn’t in the template dropdown for ' + h + '. Kept as-is (a new value to create in PickTrace), or pick a listed value.' };
      return null;
    };
    // Every kept row with the columns it has flagged (f) and changed (c), for
    // the filter's counts and "only columns with something to see".
    const allVisible = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const info = new Map();
    const counts = { all: allVisible.length, fix: 0, chg: 0, any: 0 };
    allVisible.forEach(ri => {
      const f = [], c = [];
      for (let i = 0; i < bulkHeaders.length; i++) {
        if (flagOf(ri, i)) f.push(i);
        if (changeOf(ri, i)) c.push(i);
      }
      info.set(ri, { f, c });
      if (f.length) counts.fix++;
      if (c.length) counts.chg++;
      if (f.length || c.length) counts.any++;
    });
    if (pvFocus && pvFocus.ci >= bulkHeaders.length) pvFocus = null;
    const shown = pvFocus
      ? allVisible.filter(ri => !formattedRows[ri][pvFocus.ci])
      : allVisible.filter(ri => IMWorkbench.rowMatch(pvMode, info.get(ri).f.length, info.get(ri).c.length));
    let cols = bulkHeaders.map((_, i) => i);
    if (pvColsOnly) {
      const keep = new Set([bulkHeaders.indexOf('Site*'), nameColIdx].filter(i => i >= 0));
      if (pvFocus) keep.add(pvFocus.ci);
      shown.forEach(ri => { const x = info.get(ri); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      cols = cols.filter(i => keep.has(i));
    }
    // First column is the row-action column (× to remove the row).
    let head = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    cols.forEach(i => {
      const h = bulkHeaders[i], req = /\*$/.test(h);
      const listed = colDrop[i] ? ' title="Dropdown from template — pick a value or type your own"' : '';
      head += '<th' + (req ? ' style="color:#dc2626;"' : '') + listed + '>' + escHtml(h) +
        (colDrop[i] ? ' <span class="ts-col-listed" title="Has a template dropdown">&#9662;</span>' : '') + '</th>';
    });
    head += '</tr></thead>';
    const rowHtml = ri => {
      const row = formattedRows[ri];
      let html = '<tr><td style="text-align:center;padding:0;">' +
        '<button class="ts-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button>' +
        '</td>';
      cols.forEach(i => {
        const val = row[i] == null ? '' : String(row[i]);
        const d = colDrop[i], flag = flagOf(ri, i);
        let title = flag ? flag.title : '';
        const cellStyle = flag ? ' style="background:' + flag.bg + ';"' : '';
        const fieldColor = flag ? ('color:' + (flag.bg === RED ? '#7f1d1d' : '#7c2d12') + ';') : '';
        const note = changeOf(ri, i);
        if (note) title = title ? title + '\n' + note : note;
        const tdOpen = '<td class="ts-cell' + (note ? ' is-chg' : '') + '"' + cellStyle + '>';
        if (d) {
          // A dropdown cell is a light button until clicked; then it becomes
          // the full template-list <select> (current off-list value kept as a
          // "(current)" option, "Type custom…" for a brand-new value) — see
          // IMWorkbench.wirePicks. One full <select> per cell froze big files.
          html += tdOpen + IMWorkbench.pickButton(val, ' data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') +
            (fieldColor ? ' style="' + fieldColor + '"' : '')) + '</td>';
        } else {
          // Text until clicked / tabbed to, then an input (IMWorkbench.wireTextCells).
          html += tdOpen + IMWorkbench.textCell(val, ' data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') +
            (fieldColor ? ' style="' + fieldColor + '"' : '')) + '</td>';
        }
      });
      return html + '</tr>';
    };
    const widths = [28].concat(cols.map(i => IMWorkbench.fitWidth(bulkHeaders[i], shown.map(ri => formattedRows[ri][i]))));
    IMWorkbench.rows(tbl, { head: head, ids: shown, row: rowHtml, cols: cols.length + 1, widths: widths,
      keys: [null].concat(cols.map(i => 'loc:' + bulkHeaders[i])), cell: (ri, p) => formattedRows[ri][cols[p - 1]],
      empty: pvFocus ? 'No row is missing ' + bulkHeaders[pvFocus.ci] + ' any more ✓' :
        pvMode === 'fix' ? 'No row needs a fix ✓' : 'Nothing here has been changed.' });
    wirePreview(sec, tbl);
    const hint = sec.querySelector('.cmp-sites-hint');
    let hintText = viewPills() + IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly,
      shownCols: cols.length, totalCols: bulkHeaders.length,
      focus: pvFocus ? 'Rows with ' + bulkHeaders[pvFocus.ci] + ' empty (' + shown.length + ')' : '' });
    hintText += '<span class="wb-filter-note">' + (shown.length === allVisible.length ? 'All ' + allVisible.length : shown.length + ' of ' + allVisible.length) +
      ' rows · every cell is editable — <span class="ts-col-listed">&#9662;</span> columns are template dropdowns. Click <b>×</b> to drop a row.</span>';
    if (removedRows.size) {
      hintText += ' <button class="btn btn-ghost btn-sm" id="ts-restore-rows">' +
        'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    }
    hint.innerHTML = hintText;
    updateExportButton();
  }

  // The preview's handlers — on the section (its filter bar, view switch and
  // Restore are redrawn every render) and on the <table> (rows come and go as
  // you scroll) — attached once. Cell edits go to whichever file is showing.
  function wirePreview(sec, tbl) {
    IMWorkbench.wirePicks(tbl, ci => pvView === 'sites' ? siteDropFor(ci) : dropdownValuesFor(activeSchemaHeaders()[ci]), '__ts_custom__');
    IMWorkbench.wireTextCells(tbl);
    if (!sec._tsFilterWired) {
      sec._tsFilterWired = true;
      sec.addEventListener('change', e => {
        if (e.target.classList.contains('wb-rowfilter')) { pvMode = e.target.value; IMWorkbench.rowsTop(tbl); renderPreview(); }
        else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
      });
      sec.addEventListener('click', e => {
        const view = e.target.closest('[data-view]');
        if (view) { pvView = view.dataset.view; pvFocus = null; IMWorkbench.rowsTop(tbl); renderPreview(); return; }
        if (e.target.closest('.wb-focus-clear')) { pvFocus = null; IMWorkbench.rowsTop(tbl); renderPreview(); return; }
        if (e.target.closest('#ts-restore-rows')) {
          removedRows.clear();
          renderPreview();
          renderRequired();
          renderSitesAndCvToCreate();
          updateSummary();
        }
      });
    }
    if (tbl._tsCellWired) return;
    tbl._tsCellWired = true;
    tbl.addEventListener('click', e => {
      const rm = e.target.closest('.ts-row-remove');
      if (!rm) return;
      e.stopPropagation();
      removedRows.add(+rm.dataset.ri);
      renderPreview();
      renderRequired();
      renderSitesAndCvToCreate();
      updateSummary();
    });
    tbl.addEventListener('change', e => {
      const el = e.target.closest('.ts-cell-select, .ts-cell-input');
      if (!el) return;
      const ri = +el.dataset.ri, ci = +el.dataset.ci;
      if (pvView === 'sites') {
        const s = siteGrid[ri];
        if (s == null) return;
        if (el.value === '__ts_custom__') {
          const cur = (buildSiteFile().rows[ri] || [])[ci] || '';
          el.closest('td').innerHTML = '<input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + ci +
            '" placeholder="Type a value…" value="' + escHtml(cur) + '">';
          const inp = tbl.querySelector('input[data-ri="' + ri + '"][data-ci="' + ci + '"]');
          if (inp) { inp.focus(); inp.select(); }
          return;
        }
        siteEdits[s + '|' + siteKey(siteTpl.headers[ci])] = el.value;
        refreshSites();
        return;
      }
      // "Type custom…" swaps the cell to a free-text input in place; the real
      // value gets committed by the input's own change event.
      if (el.classList.contains('ts-cell-select') && el.value === '__ts_custom__') { swapSelectToInput(el, ri, ci); return; }
      setCell(ri, ci, el.value);
      recomputeAfterEdit();
    });
  }

  // Locations | Sites — which file the preview shows, once there's a Sites file.
  function viewPills() {
    if (!siteFileOn()) return '';
    const f = buildSiteFile();
    if (!f.rows.length) return '';
    const bad = f.rows.filter(r => f.headers.some((h, i) => /\*$/.test(h) && !r[i])).length;
    const nLoc = formattedRows.filter((_, i) => !excluded(i)).length;
    return '<span class="wb-views" role="group" aria-label="File shown">' +
      '<button type="button" class="btn btn-sm ts-domain-pill' + (pvView === 'locations' ? ' ts-domain-active' : '') + '" data-view="locations">Locations <b>' + nLoc + '</b></button>' +
      '<button type="button" class="btn btn-sm ts-domain-pill' + (pvView === 'sites' ? ' ts-domain-active' : '') + '" data-view="sites">Sites <b>' + f.rows.length + '</b>' +
        (bad ? ' <span class="dt-pill-warn" title="' + bad + ' site' + (bad === 1 ? '' : 's') + ' still missing a required value">⚠ ' + bad + '</span>' : '') + '</button></span>';
  }

  // The Sites file in the preview: same grid, filter and widths as Locations.
  // Name comes from the Locations Site column, so it's changed there, not here.
  function renderSitesPreview(sec, tbl) {
    const RED = '#fee2e2', AMBER = '#fef3c7';
    const f = buildSiteFile(), H = f.headers;
    siteGrid = f.names;
    const lower = H.map((_, i) => { const l = siteDropFor(i); return l ? new Set(l.map(v => String(v).toLowerCase().trim())) : null; });
    const flagOf = (si, i) => {
      const v = f.rows[si][i] == null ? '' : String(f.rows[si][i]);
      if (/\*$/.test(H[i]) && !v) return { bg: RED, title: 'Required — PickTrace rejects the site without it' };
      if (lower[i] && v && !lower[i].has(v.toLowerCase().trim())) return { bg: AMBER, title: 'Off-list — "' + v + '" isn’t in the Sites template dropdown for ' + H[i] + '.' };
      return null;
    };
    const filled = (si, i) => f.prov[si][i] && f.prov[si][i] !== 'name';
    const all = f.rows.map((_, si) => si), info = new Map(), counts = { all: all.length, fix: 0, chg: 0, any: 0 };
    all.forEach(si => {
      const fl = [], c = [];
      H.forEach((h, i) => { if (!h) return; if (flagOf(si, i)) fl.push(i); if (filled(si, i)) c.push(i); });
      info.set(si, { f: fl, c });
      if (fl.length) counts.fix++;
      if (c.length) counts.chg++;
      if (fl.length || c.length) counts.any++;
    });
    const shown = all.filter(si => IMWorkbench.rowMatch(pvMode, info.get(si).f.length, info.get(si).c.length));
    let cols = H.map((h, i) => h ? i : -1).filter(i => i >= 0);
    if (pvColsOnly) {
      const keep = new Set([f.nameI].filter(i => i >= 0));
      shown.forEach(si => { const x = info.get(si); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      cols = cols.filter(i => keep.has(i));
    }
    const head = '<thead><tr>' + cols.map(i => '<th' + (/\*$/.test(H[i]) ? ' style="color:#dc2626;"' : '') + '>' + escHtml(H[i]) +
      (lower[i] ? ' <span class="ts-col-listed">&#9662;</span>' : '') + '</th>').join('') + '</tr></thead>';
    const rowHtml = si => '<tr>' + cols.map(i => {
      const v = f.rows[si][i] == null ? '' : String(f.rows[si][i]);
      if (i === f.nameI) {
        return '<td class="ts-cell ts-cell-ro" title="From the Site column of the Locations file — change it there">' + escHtml(v) + '</td>';
      }
      const flag = flagOf(si, i), note = siteNote(f.prov[si][i], f.ranches.get(f.names[si]));
      const title = [flag ? flag.title : '', note].filter(Boolean).join('\n');
      const attrs = ' data-ri="' + si + '" data-ci="' + i + '"' + (title ? ' title="' + escHtml(title) + '"' : '') +
        (flag ? ' style="color:' + (flag.bg === RED ? '#7f1d1d' : '#7c2d12') + ';"' : '');
      return '<td class="ts-cell' + (note ? ' is-chg' : '') + '"' + (flag ? ' style="background:' + flag.bg + ';"' : '') + '>' +
        (lower[i] ? IMWorkbench.pickButton(v, attrs) : IMWorkbench.textCell(v, attrs)) + '</td>';
    }).join('') + '</tr>';
    IMWorkbench.rows(tbl, { head: head, ids: shown, row: rowHtml, cols: cols.length,
      widths: cols.map(i => IMWorkbench.fitWidth(H[i], shown.map(si => f.rows[si][i]))),
      keys: cols.map(i => 'site:' + H[i]), cell: (si, p) => f.rows[si][cols[p]],
      empty: pvMode === 'fix' ? 'Every new site is complete ✓' : 'Nothing filled in yet.' });
    wirePreview(sec, tbl);
    sec.querySelector('.cmp-sites-hint').innerHTML = viewPills() +
      IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly, shownCols: cols.length, totalCols: H.filter(Boolean).length }) +
      '<span class="wb-filter-note">' + (shown.length === all.length ? 'All ' + all.length : shown.length + ' of ' + all.length) +
      ' new sites · Name comes from the Locations Site column; every other cell is editable. Fill whole columns under <b>Sites file</b> and <b>Site addresses</b>.</span>';
  }

  // Route a preview cell edit into the correct sticky store + the live grid.
  // Name* flows through nameOverrides (keeps the Name Collisions panel in sync);
  // every other column flows through cellOverrides. Empty values are stored too
  // (except Name*, which clears its override) so a deliberate clear sticks across
  // rebuilds instead of snapping back to the mapped source value.
  // "Was … → now … · why" when the cell differs from its source, for the
  // preview's change tick (see IMWorkbench.changeNote). Crop & Variety is
  // compared with the source's Crop Type + Variety joined the same way.
  let cvSrc = { data: null };
  function changeOf(ri, ci) {
    const heads = activeSchemaHeaders(), h = heads[ci];
    const raw = (si) => si != null && si >= 0 && srcData.rows[ri] ? String(srcData.rows[ri][si] == null ? '' : srcData.rows[ri][si]).trim() : '';
    let was = raw(mapping[ci]);
    if (h === 'Crop & Variety*') {
      if (cvSrc.data !== srcData) {
        cvSrc = { data: srcData, c: findSrcColByNames(['crop', 'crop type', 'crop name', 'plant_name', 'plant name']),
          v: findSrcColByNames(['variety', 'varietal'], ['clone', 'subvariety', 'sub variety', 'sub-variety']) };
      }
      if (cvSrc.c >= 0 && cvSrc.v >= 0 && cvSrc.c !== cvSrc.v) {
        const c = raw(cvSrc.c), v = raw(cvSrc.v);
        was = c && v ? c + '-' + v : c || v;
      }
    }
    const now = formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci]);
    let why = 'Cleaned up by the tool';
    if (cellOverrides[ri + '|' + ci] != null) why = 'Your edit';
    else if (h === 'Name*' && nameOverrides[ri]) why = 'Renamed by you';
    else if (h === 'Site*' && siteGaps.some(g => g.applied && g.rows.indexOf(ri) >= 0)) why = 'Copied from the row above (you confirmed)';
    else if (smartFixMap[ci + '||' + was.toUpperCase()] === now) why = 'Smart Fix you applied';
    else if (!was && manualFills[ci] === now) why = 'Column fill you applied';
    else if (h === 'Crop & Variety*' || h === 'Clone/Subvariety') why = 'Subvariety split out of Crop & Variety';
    else if (dateColIdxs().indexOf(ci) >= 0) why = /^\d{4}$/.test(was) ? 'Year expanded to a full date' : 'Date written as YYYY-MM-DD';
    else if (NUMERIC_HEADERS.has(h)) why = 'Unit or formatting removed';
    else if (h === 'Production Status') why = 'Matched to the dropdown';
    return IMWorkbench.changeNote(was, now, why);
  }

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
    const tplSites = knownSites();
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

    // The new sites' addresses live in the Site addresses panel; this list is
    // for creating them by hand — so it steps aside once the Sites file does it.
    const renderList = (sectionId, tableId, titleId, label, m) => {
      const sec = $(sectionId);
      if (!m.size) { sec.style.display = 'none'; return; }
      sec.style.display = '';
      $(titleId).textContent = label + ' (' + m.size + ')';
      const sorted = [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
      let html = '<thead><tr><th>Value</th><th>Row count</th></tr></thead><tbody>';
      sorted.forEach(([val, cnt]) => {
        html += '<tr><td class="ts-copy-cell" title="Click to copy">' + escHtml(val) + '</td><td>' + cnt + '</td></tr>';
      });
      $(tableId).innerHTML = html + '</tbody>';
    };
    renderList('ts-section-sites-create', 'ts-sites-create-table', 'ts-sites-create-title', 'Sites to Create in PickTrace',
      siteFileOn() ? new Map() : sites);
    renderCvList(cvs, tplCvs);
    return { sitesCount: sites.size, cvsCount: cvs.size, listsCount: renderListsCreate() };
  }

  // Missing Crop & Variety values, split back into the two things PickTrace
  // creates separately: the crop, then its varieties. Splits at the FIRST
  // hyphen, as Bulk Template Standardize does — variety names carry hyphens
  // ("Sample 1234-1"), crop names don't. Case variants merge.
  function splitCv(v) {
    const s = String(v == null ? '' : v).trim(), i = s.indexOf('-');
    return i > 0 ? { crop: s.slice(0, i).trim(), variety: s.slice(i + 1).trim() } : { crop: s, variety: '' };
  }
  // value → count  →  [{ crop, isNew, vars: [{ variety, count }] }], crops and
  // varieties A→Z. isNew: the template lists no variety at all for that crop.
  function groupCropVarieties(cvs, tplCvs) {
    const known = new Set([...(tplCvs || [])].map(v => splitCv(v).crop.toUpperCase()));
    const crops = new Map();
    cvs.forEach((count, value) => {
      const p = splitCv(value), ck = p.crop.toUpperCase(), vk = p.variety.toUpperCase();
      if (!crops.has(ck)) crops.set(ck, { crop: p.crop, isNew: !known.has(ck), vars: new Map() });
      const g = crops.get(ck);
      const e = g.vars.get(vk) || { variety: p.variety, count: 0 };
      e.count += count;
      g.vars.set(vk, e);
    });
    return [...crops.values()].sort((a, b) => a.crop.localeCompare(b.crop)).map(g => ({
      crop: g.crop, isNew: g.isNew,
      vars: [...g.vars.values()].sort((a, b) => a.variety.localeCompare(b.variety))
    }));
  }

  // Crop | Variety | Row count, each crop written once with its cell spanning
  // its varieties — the order you create them in PickTrace.
  let cvGroups = [];   // what the table shows; Copy List writes the same rows
  function renderCvList(cvs, tplCvs) {
    const sec = $('ts-section-cv-create');
    cvGroups = groupCropVarieties(cvs, tplCvs);
    if (!cvGroups.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    let nVar = 0, nNew = 0;
    cvGroups.forEach(g => { nVar += g.vars.length; if (g.isNew) nNew++; });
    $('ts-cv-create-title').textContent = 'Crops & Varieties to Create in PickTrace (' +
      cvGroups.length + ' crop' + (cvGroups.length === 1 ? '' : 's') + (nNew ? ', ' + nNew + ' new' : '') +
      ' · ' + nVar + ' variet' + (nVar === 1 ? 'y' : 'ies') + ')';
    let html = '<thead><tr><th>Crop</th><th>Variety</th><th>Row count</th></tr></thead><tbody>';
    cvGroups.forEach(g => {
      g.vars.forEach((v, i) => {
        html += '<tr' + (i === 0 ? ' class="ts-cv-first"' : '') + '>';
        if (i === 0) {
          html += '<td class="ts-cv-crop" rowspan="' + g.vars.length + '" data-copy="' + escHtml(g.crop) + '" title="Click to copy">' +
            escHtml(g.crop) + ' <span class="dt-build-tag' + (g.isNew ? ' dt-build-tag-new">new crop' : '">crop exists') + '</span></td>';
        }
        html += '<td class="ts-copy-cell" data-copy="' + escHtml(v.variety) + '" title="Click to copy">' +
          (v.variety ? escHtml(v.variety) : '<span class="text-muted">(no variety)</span>') + '</td>' +
          '<td>' + v.count + '</td></tr>';
      });
    });
    $('ts-cv-create-table').innerHTML = html + '</tbody>';
  }

  // ─── Sites file ───
  // Sites the client file names that aren't in the Locations template's Site
  // list must exist in PickTrace before their locations can load. With the
  // org's Sites template loaded, this tool builds that upload too — one row
  // per new site:
  //   Name                 the Locations Site value
  //   Address1/City/State/Zip(/Country)  split out of the most specific address
  //                        you gave: the site's own, its ranch's (the name
  //                        before " - "), or the one for every new site
  //   any other column     your fill for every new site; a one-option list
  //                        (Employer, often) fills itself
  //   any cell             your edit in the Sites preview, over all of the above
  // Export then downloads both files in one .zip, and those sites count as
  // listed in the Locations Site column (dropdownValuesFor) — the Sites file
  // is uploaded first.
  function siteFileOn() { return !!(siteTpl && tplData); }
  // Sites PickTrace already has: the Locations template's Site list, plus any
  // Site in the existing-locations export (a list can lag behind the account).
  function knownSites() {
    const set = new Set(tplData && tplData.dropdowns.get('site') || []);
    if (existingData && existingData.sites) existingData.sites.forEach(x => set.add(x));
    return set;
  }
  const siteKey = h => norm(h).replace(/\*+$/, '').trim();
  // Each new site's ranch: the name before " - " — or before a bare "-" when
  // another site spells that ranch with " - " ("Example-Block 2" joins
  // "Example - Block 1"); otherwise the whole name. → Map name → ranch.
  function ranchMap(names) {
    const known = new Map();
    names.forEach(s => { const i = s.indexOf(' - '); if (i > 0) { const r = s.slice(0, i).trim(); if (!known.has(r.toUpperCase())) known.set(r.toUpperCase(), r); } });
    const m = new Map();
    names.forEach(s => {
      const i = s.indexOf(' - '), j = s.indexOf('-');
      const head = (i > 0 ? s.slice(0, i) : j > 0 ? s.slice(0, j) : '').trim().toUpperCase();
      m.set(s, known.get(head) || s.trim());
    });
    return m;
  }
  // New site → its row count. Judged against the Locations template's own
  // Site list (never the widened one, or the Sites file would vouch for
  // itself); case variants count as one site, first spelling kept.
  function newSitesMap() {
    const out = new Map();
    if (!formattedRows || !tplData) return out;
    const si = activeSchemaHeaders().indexOf('Site*');
    if (si < 0) return out;
    const known = buildCaseMap(knownSites());
    const first = new Map();
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      const s = String(r[si] == null ? '' : r[si]).trim();
      if (!s || inDropdown(known, s)) return;
      const k = first.get(s.toUpperCase()) || s;
      first.set(s.toUpperCase(), k);
      out.set(k, (out.get(k) || 0) + 1);
    });
    return out;
  }
  function newSiteNames() { return [...newSitesMap().keys()].sort((a, b) => a.localeCompare(b)); }
  // The address a site gets, and where it came from.
  function siteAddrFor(s, ranch) {
    if (siteAddresses[s]) return { text: siteAddresses[s], from: 'site' };
    const g = siteAddrGroup[ranch];
    if (g) return { text: g, from: 'ranch' };
    return siteAddrAll ? { text: siteAddrAll, from: 'all' } : { text: '', from: '' };
  }
  function parseAddr(text) {
    if (!text) return { address1: '', city: '', state: '', zip: '', country: '' };
    return window.IMAddress ? IMAddress.parse(text) : { address1: text, city: '', state: '', zip: '', country: '' };
  }
  const addrComplete = p => !!(p.address1 && p.city && p.state && p.zip);
  function siteDropFor(ci) {
    if (!siteTpl) return null;
    const set = siteTpl.dropdowns.get(siteKey(siteTpl.headers[ci] || ''));
    return set && set.size ? [...set] : null;
  }
  // → { headers, names, rows, prov, nameI } — prov[r][c] says where a cell
  // came from: name · addr:site|ranch|all · fill · only · edit.
  function buildSiteFile() {
    if (!siteFileOn()) return { headers: [], names: [], rows: [], prov: [], nameI: -1 };
    const H = siteTpl.headers, names = newSiteNames(), ranches = ranchMap(names);
    const at = k => H.findIndex(h => siteKey(h) === k);
    const nameI = at('name');
    const addrI = ['address1', 'city', 'state', 'zip', 'country'].map(k => [k, at(k)]).filter(x => x[1] >= 0);
    const lists = H.map((_, i) => siteDropFor(i));
    const rows = [], prov = [];
    names.forEach(s => {
      const row = H.map(() => ''), pv = H.map(() => '');
      if (nameI >= 0) { row[nameI] = s; pv[nameI] = 'name'; }
      const a = siteAddrFor(s, ranches.get(s)), p = parseAddr(a.text);
      addrI.forEach(([k, i]) => { if (p[k]) { row[i] = p[k]; pv[i] = 'addr:' + a.from; } });
      H.forEach((h, i) => {
        if (row[i] || !h) return;
        const f = siteFill[siteKey(h)];
        if (f) { row[i] = f; pv[i] = 'fill'; }
        else if (lists[i] && lists[i].length === 1) { row[i] = lists[i][0]; pv[i] = 'only'; }
      });
      H.forEach((h, i) => {
        const e = siteEdits[s + '|' + siteKey(h)];
        if (e != null && i !== nameI) { row[i] = e; pv[i] = 'edit'; }
      });
      rows.push(row); prov.push(pv);
    });
    return { headers: H, names, rows, prov, nameI, ranches };
  }
  function siteNote(pv, ranch) {
    return pv === 'addr:site' ? 'From the address you gave this site'
      : pv === 'addr:ranch' ? 'From the address for ' + ranch
      : pv === 'addr:all' ? 'From the address for every new site'
      : pv === 'fill' ? 'Your fill for every new site'
      : pv === 'only' ? 'The only option in the Sites template'
      : pv === 'edit' ? 'Your edit' : '';
  }
  // The Sites file's rows as written: dropdown values in the template's
  // casing, everything as text (a Zip keeps its leading zero).
  function siteRowsForExport() {
    const f = buildSiteFile();
    const maps = f.headers.map((_, i) => buildCaseMap(siteDropFor(i)));
    return f.rows.map(row => row.map((v, i) => {
      const s = v == null ? '' : String(v);
      const k = s.toUpperCase().trim();
      return maps[i].size && maps[i].has(k) ? maps[i].get(k) : s;
    }));
  }
  // Button counts: sites still missing a non-address required value (all of
  // them until the Sites template is in), and sites without a full address.
  const ADDR_KEYS = new Set(['address1', 'address2', 'city', 'state', 'zip']);
  function siteDetailNeeds() {
    const n = newSitesMap().size;
    if (!n || !siteFileOn()) return n;
    const f = buildSiteFile();
    return f.rows.filter(r => f.headers.some((h, i) => /\*$/.test(h) && !ADDR_KEYS.has(siteKey(h)) && !r[i])).length;
  }
  function siteAddrNeeds() {
    if (siteFileOn()) {
      const f = buildSiteFile();
      return f.rows.filter(r => f.headers.some((h, i) => /\*$/.test(h) && ADDR_KEYS.has(siteKey(h)) && !r[i])).length;
    }
    const names = newSiteNames(), ranches = ranchMap(names);
    return names.filter(s => !addrComplete(parseAddr(siteAddrFor(s, ranches.get(s)).text))).length;
  }

  function renderSiteFile() {
    const sec = $('ts-section-sitefile');
    if (!sec) return;
    const n = formattedRows ? newSitesMap().size : 0;
    if (!n) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('ts-sitefile-title').textContent = 'Sites file — ' + n + ' new site' + (n === 1 ? '' : 's');
    let html;
    if (!siteTpl) {
      html = '<p class="cmp-sites-hint">' + n + ' Site value' + (n === 1 ? ' isn’t' : 's aren’t') + ' in your Locations template, so ' +
        (n === 1 ? 'it doesn’t' : 'they don’t') + ' exist in PickTrace yet. Add your org’s personalized <b>Sites</b> template and the tool builds the Sites upload for ' +
        (n === 1 ? 'it' : 'them') + ' too &mdash; Export then downloads both files in one .zip. Upload the Sites file first, then Locations.</p>' +
        '<div class="ts-sf-actions"><label class="btn btn-primary btn-sm" for="ts-sitetpl-file">&#8679; Upload Sites template</label>' +
        '<span class="text-muted small">Addresses can go in now, under <b>Site addresses</b> &mdash; they carry over.</span></div>';
    } else {
      const f = buildSiteFile();
      html = '<p class="cmp-sites-hint">Built from <b>' + escHtml(siteTpl.fileName) + '</b>: one row per new site, its Name from the Locations Site column. ' +
        'Fill a column for every new site here, or change one site in the preview (<b>Sites</b> view). ' +
        'Export downloads both files in one .zip &mdash; <b>upload the Sites file first</b>, then Locations.</p>';
      html += '<div class="table-wrap"><table class="data-table"><thead><tr><th>Column</th><th>Status</th><th>Fill every new site</th></tr></thead><tbody>';
      f.headers.forEach((h, i) => {
        const k = siteKey(h);
        if (!h || k === 'name' || ADDR_KEYS.has(k)) return;
        const req = /\*$/.test(h), list = siteDropFor(i), cur = siteFill[k] || '';
        const emptyN = f.rows.filter(r => !r[i]).length;
        const only = !!(list && list.length === 1);
        const status = emptyN
          ? '<span style="color:' + (req ? '#dc2626' : 'var(--text-3)') + ';font-weight:600;">' + (req ? '⚠ ' : '') + emptyN + ' empty</span>'
          : '<span style="color:#15803d;">✓ ' + (only && !cur ? 'only option — filled' : 'filled') + '</span>';
        const ctl = list
          ? '<select class="input-field ts-sf-fill" data-k="' + escHtml(k) + '"><option value="">— pick —</option>' +
              list.slice().sort().map(v => '<option' + (v === cur ? ' selected' : '') + '>' + escHtml(v) + '</option>').join('') + '</select>'
          : '<input type="text" class="input-field ts-sf-fill" data-k="' + escHtml(k) + '" value="' + escHtml(cur) + '" placeholder="Value for every new site">';
        html += '<tr><td><b' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</b></td><td>' + status + '</td>' +
          '<td><div class="ts-req-ctl">' + ctl + (cur ? '<button type="button" class="btn btn-ghost btn-sm ts-sf-clear" data-k="' + escHtml(k) + '">Clear</button>' : '') +
          '</div></td></tr>';
      });
      html += '</tbody></table></div>';
    }
    $('ts-sitefile-body').innerHTML = html;
  }

  function renderSiteAddr() {
    const sec = $('ts-section-siteaddr');
    if (!sec) return;
    const names = formattedRows ? newSiteNames() : [];
    if (!names.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const groups = new Map();
    const ranches = ranchMap(names);
    names.forEach(s => { const g = ranches.get(s); groups.set(g, (groups.get(g) || 0) + 1); });
    const checked = text => {
      if (!text) return '';
      const p = parseAddr(text);
      if (addrComplete(p)) return '<span class="ts-addr-ok" title="Address1: ' + escHtml(p.address1) + '">✓ ' + escHtml(p.city + ', ' + p.state + ' ' + p.zip) + '</span>';
      const miss = ['address1', 'city', 'state', 'zip'].filter(k => !p[k]).map(k => k === 'address1' ? 'street' : k);
      return '<span class="ts-addr-bad">⚠ no ' + miss.join(', ') + ' found</span>';
    };
    const field = (scope, key, val, ph) => '<input type="text" class="input-field ts-addr-in" data-scope="' + scope + '"' +
      (key != null ? ' data-key="' + escHtml(key) + '"' : '') + ' value="' + escHtml(val) + '" placeholder="' + escHtml(ph) + '">';
    let html = '<p class="cmp-sites-hint">Type an address as <b>street, city, ST zip</b> &mdash; it’s split into Address1, City, State and Zip. ' +
      'The most specific one wins: a site’s own address, then its ranch’s, then the one for every site.' +
      (siteTpl ? '' : ' These carry into the Sites file once you add the Sites template.') + '</p>';
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Every new site (' + names.length + ')</div>' +
      '<div class="ts-addr-row">' + field('all', null, siteAddrAll, '123 Main St, City, ST 12345') + checked(siteAddrAll) + '</div></div>';
    // Per ranch only when the names actually group (several sites to a ranch).
    if (groups.size < names.length) {
      html += '<div class="ts-addr-block"><div class="ts-addr-label">Per ranch &mdash; the name before “ - ” (' + groups.size + ')</div>' +
        '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Ranch</th><th>Sites</th><th>Address</th></tr></thead><tbody>';
      [...groups.keys()].sort((a, b) => a.localeCompare(b)).forEach(g => {
        const v = siteAddrGroup[g] || '';
        html += '<tr><td>' + escHtml(g) + '</td><td>' + groups.get(g) + '</td><td><div class="ts-addr-row">' +
          field('ranch', g, v, siteAddrAll ? siteAddrAll + ' (every site)' : 'street, city, ST zip') + checked(v) + '</div></td></tr>';
      });
      html += '</tbody></table></div></div>';
    }
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Each site (' + names.length + ')</div>' +
      '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Site</th><th>Address</th></tr></thead><tbody>';
    names.forEach(s => {
      const own = siteAddresses[s] || '', inh = siteAddrGroup[ranches.get(s)] || siteAddrAll;
      html += '<tr><td>' + escHtml(s) + '</td><td><div class="ts-addr-row">' +
        field('site', s, own, inh ? inh + ' (inherited)' : 'street, city, ST zip') + checked(own || inh) + '</div></td></tr>';
    });
    $('ts-siteaddr-body').innerHTML = html + '</tbody></table></div></div>';
  }

  // After a Sites-file change: the lists, the preview and the counts.
  function refreshSites() {
    if (!formattedRows) return;
    renderSitesAndCvToCreate();
    renderPreview();
    updateSummary();
  }

  function handleSiteTplFile(file) {
    const r = new FileReader();
    r.onload = e => useSiteTemplate(parseTemplate(e.target.result, file.name), file.name);
    r.readAsArrayBuffer(file);
  }
  function useSiteTemplate(parsed, fileName) {
    if (!parsed || !parsed.headers.length) { alert('Could not find a DATA ENTRY sheet in this template.'); return; }
    const kind = identifyTemplate(parsed.headers);
    if (kind.kind !== 'sites') {
      alert('That isn’t the Sites bulk template' + (kind.kind === 'unknown' ? '' : ' — it’s the ' + kind.label + ' one') +
        '.\n\nThe Sites template’s DATA ENTRY sheet has Site Type, Employer and Address1 columns.');
      return;
    }
    siteTpl = parsed;
    $('ts-sitetpl-name').textContent = fileName;
    $('ts-sitetpl-meta').textContent = 'Sites template · ' + parsed.headers.filter(Boolean).length + ' columns';
    if (!formattedRows) { if (wb) wb.refresh(); return; }
    refreshAll(); // the new sites now count as listed everywhere
    if (newSitesMap().size && wb) wb.open('sitefile');
  }

  // ─── Smart Fixes panel — every off-list value gets a dropdown to choose the
  // target template value (suggested match pre-selected; the rest ask the user). ───
  // Blank Site groups, at the top of Smart Fixes. Nothing copies until you
  // click Copy down (or Copy down all); Undo puts the blanks back.
  function renderSiteGaps() {
    const el = $('ts-sitegaps');
    if (!el) return;
    if (!siteGaps.length) { el.innerHTML = ''; el.style.display = 'none'; return; }
    const pending = siteGaps.filter(g => !g.applied);
    let n = 0;
    pending.forEach(g => { n += g.rows.length; });
    const span = rs => rs.length === 1 ? 'Row ' + (rs[0] + 1) : 'Rows ' + (rs[0] + 1) + '–' + (rs[rs.length - 1] + 1);
    el.innerHTML = '<div class="ts-sitegaps-head"><b>Blank Site</b> — ' +
      (pending.length
        ? n + ' row' + (n === 1 ? ' has' : 's have') + ' no Site, right under a row that does. Copy each group\'s site down? Nothing changes until you confirm.'
        : 'every group is copied down.') +
      (pending.length > 1 ? ' <button type="button" class="btn btn-primary btn-sm" data-act="all">Copy down all</button>' : '') + '</div>' +
      '<div class="table-wrap"><table class="data-table"><thead><tr><th>Data rows</th><th>Locations</th><th>Site to copy down</th><th></th></tr></thead><tbody>' +
      siteGaps.map(g => '<tr><td>' + span(g.rows) + ' <span class="text-muted small">(' + g.rows.length + ')</span></td>' +
        '<td>' + escHtml(g.names.join(', ') + (g.rows.length > g.names.length ? ', …' : '')) + '</td>' +
        '<td><b>' + escHtml(g.site) + '</b></td><td>' +
        (g.applied
          ? '<span style="color:#15803d;">✓ Copied</span> <button type="button" class="btn btn-ghost btn-sm" data-act="undo" data-id="' + escHtml(g.id) + '">Undo</button>'
          : '<button type="button" class="btn btn-primary btn-sm" data-act="apply" data-id="' + escHtml(g.id) + '">Copy down</button>') +
        '</td></tr>').join('') + '</tbody></table></div>';
    el.style.display = '';
  }

  function renderSmartFixes() {
    const sec = $('ts-section-smartfix');
    if (!sec) return;
    const fixes = computeSmartFixes();
    // Blank Site renders on its own — it has its own button (and panel).
    renderSiteGaps();
    if (!fixes.length) { sec.style.display = 'none'; $('ts-smartfix-table').innerHTML = ''; return; }
    sec.style.display = '';
    let totalRows = 0, needPick = 0;
    fixes.forEach(f => { totalRows += f.count; if (!f.suggestion) needPick++; });
    const t = $('ts-smartfix-title');
    if (t) t.textContent = 'Smart Fixes (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + totalRows + ' rows' +
      (needPick ? ' — ' + needPick + ' need your choice' : '') + ')';
    let html = '<thead><tr><th>Column</th><th>Value in your data</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    // Each picker is a button until clicked (IMWorkbench.wirePicks): a full
    // <select> per row listed every allowed value 171 times over on a big file.
    const optsByCol = new Map(fixes.map(f => [f.colIdx, f.options]));
    fixes.forEach(f => {
      html += '<tr>' +
        '<td>' + escHtml(f.header) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td>' + IMWorkbench.pickButton(f.suggestion || '', ' data-ci="' + f.colIdx + '" style="min-width:200px;"',
          { kind: 'fix', cls: 'ts-smartfix-pick input-field' }) +
          (f.suggestion ? ' <span class="text-muted small" title="Confident match — spacing, punctuation or singular/plural">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') +
        '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm ts-smartfix-apply" data-ci="' + f.colIdx + '" data-from="' + escHtml(f.from) + '">Apply</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    $('ts-smartfix-table').innerHTML = html;
    smartFixOpts = optsByCol;
    IMWorkbench.wirePicks($('ts-smartfix-table'), ci => smartFixOpts.get(ci), { kind: 'fix' });
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
      // A column with blanks starts at "— pick —" so choosing its first value
      // — or its ONLY value, like a one-option Location Type — is a change and
      // fills them. Showing a sample there made the pick a no-op.
      const fillVal = manualFills[col.idx] || (empty ? '' : (sample.size ? [...sample][0] : ''));
      const entry = { ...col, empty, fillVal };
      if (empty > 0) needsAction.push(entry); else filled.push(entry);
    });

    const renderRow = (col) => {
      const dropKey = norm(col.header).replace(/\*$/, '');
      const opts = tplData.dropdowns.get(dropKey)
        || tplData.dropdowns.get(dropKey + '*')
        || tplData.dropdowns.get(norm(col.header))
        || new Set();
      // A column the template lists (Location Type, Site, …) gets a picker;
      // one it doesn't (Name, Start Date) just the text box.
      const optsHtml = '<option value="">— pick —</option>' +
        [...opts].sort().map(o => '<option' + (o === col.fillVal ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('');
      const status = col.empty > 0
        ? '<span style="color:#dc2626;font-weight:600;">⚠ ' + col.empty + ' empty</span>' +
          ' <button type="button" class="btn btn-ghost btn-sm ts-req-show" data-idx="' + col.idx + '" title="Show just these rows in the preview">Show ' +
          (col.empty === 1 ? 'the row' : 'them') + '</button>'
        : '<span style="color:#15803d;">✓ filled' + (col.fillVal ? ' (e.g. ' + escHtml(String(col.fillVal).slice(0, 30)) + ')' : '') + '</span>';
      const rowStyle = col.empty > 0 ? '' : ' style="background:var(--bg-sunken);"';
      // Date-aware quick fill — Start Date* gets a "Today" button that
      // fills every visible row with today's YYYY-MM-DD value.
      const isStartDate = /^start\s*date/i.test(col.header.replace(/\*$/, ''));
      const todayBtn = isStartDate
        ? '<button class="btn btn-warning btn-sm ts-req-today" data-idx="' + col.idx + '" title="Fill all rows with today\'s date (' + todayYMD() + ')">Today</button>'
        : '';
      // A column with blanks fills just its blanks; a filled column's value
      // is replaced in every row ("change if needed").
      const mode = col.empty > 0 ? 'empty' : 'all';
      return '<tr' + rowStyle + ' data-mode="' + mode + '">' +
        '<td><b>' + escHtml(col.header) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><div class="ts-req-ctl">' +
          (opts.size ? '<select class="ts-req-pick input-field" data-idx="' + col.idx + '" title="' + (mode === 'empty' ? 'Fills the empty cells' : 'Replaces the value in every row') + '">' + optsHtml + '</select>' : '') +
          '<input type="text" class="ts-req-override input-field" data-idx="' + col.idx + '" placeholder="' +
            (opts.size ? 'or type a value' : isStartDate ? 'YYYY-MM-DD' : 'Type a value') + '">' +
          todayBtn +
          '<button class="btn btn-primary btn-sm ts-req-apply" data-idx="' + col.idx + '">' + (mode === 'empty' ? 'Fill ' + col.empty + ' empty' : 'Apply to all') + '</button>' +
          '<button class="btn btn-ghost btn-sm ts-req-clear" data-idx="' + col.idx + '" title="Clear all values in this column (header stays)">Clear</button>' +
        '</div></td>' +
        '</tr>';
    };

    let html = '<thead><tr><th>Column</th><th>Status</th><th>Fill</th></tr></thead>';
    if (needsAction.length) {
      html += '<tbody><tr><td colspan="3" style="background:#fee2e2;color:#7f1d1d;font-weight:600;padding:6px 10px;">' +
        '⚠ Action needed (' + needsAction.length + ')</td></tr>' +
        needsAction.map(renderRow).join('') + '</tbody>';
    }
    if (filled.length) {
      html += '<tbody><tr><td colspan="3" style="background:#dcfce7;color:#14532d;font-weight:600;padding:6px 10px;">' +
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
        const pk = tr.querySelector('.ts-req-pick'), pick = pk ? pk.value.trim() : '';
        const ovr  = tr.querySelector('.ts-req-override').value.trim();
        const val = ovr || pick;
        if (!val) { alert('Pick or type a value first.'); return; }
        // A column with blanks fills only the blanks (the rest keep their
        // source values); a filled column is changed in every row.
        applyFill(idx, val, tr.dataset.mode === 'all');
      });
    });
    tbl.querySelectorAll('.ts-req-show').forEach(btn => {
      btn.addEventListener('click', e => {
        pvFocus = { ci: +e.currentTarget.dataset.idx };
        IMWorkbench.rowsTop($('ts-preview-table'));
        renderPreview();
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
        if (v) applyFill(idx, v, e.target.closest('tr').dataset.mode === 'all');
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
      (mergedRowIdxs.size ? '<div class="cmp-stat"><b>' + mergedRowIdxs.size + '</b> duplicate rows merged <span class="text-muted small">(' + mergeGroups.size + ' locations)</span></div>' : '') +
      (existingCount ? '<div class="cmp-stat cmp-warn"><b>' + existingCount + '</b> already in PickTrace (dropped)</div>' : '') +
      (collisionRows ? '<div class="cmp-stat cmp-warn"><b>' + collisionRows + '</b> name collision rows</div>' : '') +
      '<div class="cmp-stat"><b>' + bulkHeaders.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (sites.sitesCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.sitesCount + '</b> ' +
        (siteFileOn() ? 'new sites → Sites file' : 'sites to create') + '</div>' : '') +
      (sites.cvsCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.cvsCount + '</b> C&V to create</div>' : '') +
      (sites.listsCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.listsCount + '</b> list values to set up</div>' : '') +
      (pastDates ? '<div class="cmp-stat cmp-warn"><b>' + pastDates + '</b> Start Dates before today</div>' : '');
    // With a Sites file the new sites are handled there, not built by hand.
    lastSum = { pastDates: pastDates, toCreate: (siteFileOn() ? 0 : sites.sitesCount) + sites.cvsCount + (sites.listsCount || 0), collisionRows: collisionRows };
    renderSiteFile();
    renderSiteAddr();
    if (wb) wb.refresh();
  }

  // ─── Workbench: one button per section (workbench.js) ───
  let wb = null;
  let lastSum = { pastDates: 0, toCreate: 0, collisionRows: 0 };  // from updateSummary, for the button counts
  function mountWorkbench() {
    if (!window.IMWorkbench) return;
    const kept = () => formattedRows ? formattedRows.filter((_, i) => !excluded(i)) : [];
    wb = IMWorkbench.mount({
      pane: 'ts-loc-wrap', results: 'ts-results', preview: 'ts-section-preview',
      ready: () => !!(srcData && tplData && formattedRows),
      // A Sites panel shows the Sites file in the preview; any other, Locations.
      onOpen: key => {
        const v = (key === 'sitefile' || key === 'siteaddr') && siteFileOn() && newSitesMap().size ? 'sites' : 'locations';
        if (v !== pvView) { pvView = v; pvFocus = null; IMWorkbench.rowsTop($('ts-preview-table')); renderPreview(); }
      },
      sections: [
        // A Start Date before today needs fixing too (Today, in the same panel),
        // so its column counts here alongside the blank required ones.
        { key: 'required', group: 'fix', unit: 'column', name: 'Required blanks', desc: 'Fill required (*) columns; set any Start Date before today',
          ids: ['ts-section-required'],
          count: () => {
            const rows = kept();
            return activeSchemaHeaders().filter((h, i) => /\*$/.test(h) &&
              (rows.some(r => !r[i]) || (h === 'Start Date*' && lastSum.pastDates))).length;
          } },
        { key: 'dropdown', group: 'fix', unit: 'value', name: 'Not in dropdown', desc: 'Pick a listed value for anything PickTrace would reject',
          ids: ['ts-section-smartfix'], count: () => computeSmartFixes().length },
        { key: 'build', group: 'fix', unit: 'value', name: 'Build in PickTrace', desc: 'Sites, crops, varieties and list values to create, then re-download the template',
          ids: ['ts-section-sites-create', 'ts-section-cv-create', 'ts-section-lists-create'], count: () => lastSum.toCreate },
        // New sites go out in a Sites file of their own, built from the Sites template.
        { key: 'sitefile', group: 'fix', unit: 'site', name: 'Sites file', desc: 'New sites: add the Sites template, then Site Type, Employer, Country — exported with Locations',
          ids: ['ts-section-sitefile'], count: () => siteDetailNeeds() },
        { key: 'siteaddr', group: 'fix', unit: 'site', name: 'Site addresses', desc: 'Street, city, state and zip for each new site — for all, per ranch, or one by one',
          ids: ['ts-section-siteaddr'], count: () => siteAddrNeeds() },
        { key: 'fillsite', group: 'fix', unit: 'row', name: 'Blank Site', desc: 'Locations with no Site under one that has it — copy it down?',
          ids: ['ts-sitegaps'], count: () => siteGaps.filter(g => !g.applied).reduce((n, g) => n + g.rows.length, 0) },
        { key: 'collisions', group: 'fix', unit: 'row', name: 'Name collisions', desc: 'Blocks sharing a Site + Name — PickTrace keeps only one',
          ids: ['ts-section-collisions'], count: () => lastSum.collisionRows },
        { key: 'dates', group: 'fix', unit: 'cell', name: 'Year-only dates', desc: 'Dates written as just a year — pick a month and day',
          ids: ['ts-section-datefill'], count: () => countYearOnlyDates(kept(), dateFillIdxs()) },
        { key: 'mapping', group: 'check', unit: 'column', name: 'Column mapping', desc: 'Which source column fills each template column',
          ids: ['ts-section-mapping'], count: () => activeSchemaHeaders().filter((_, i) => mapping[i] != null && mapping[i] >= 0).length },
        { key: 'merged', group: 'check', unit: 'row', name: 'Merged duplicates', desc: 'Repeated Site + Name rows folded into one location',
          ids: ['ts-section-merged'], count: () => mergedRowIdxs.size },
        { key: 'existing', group: 'check', unit: 'row', name: 'Already in PickTrace', desc: 'Rows dropped because the existing export has them',
          ids: ['ts-section-existing'], count: () => existingRowIdxs.size }
      ]
    });
  }

  // Build a case-insensitive index of a dropdown Set so values like
  // "Example Vineyards" match the template's "EXAMPLE VINEYARDS".
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
    const tplSites = knownSites();
    const tplCvs   = tplData.dropdowns.get('crop & variety') || new Set();
    const tplSitesMap = buildCaseMap(tplSites);
    const tplCvsMap   = buildCaseMap(tplCvs);
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    const checkSites = tplSites.size > 0 && !siteFileOn(); // the Sites file brings the new ones
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
    // The Sites file's empty required cells — PickTrace rejects those sites, and
    // then every location under them.
    let siteFile = null;
    if (siteFileOn()) {
      const sf = buildSiteFile(), cols = new Set();
      let empty = 0;
      sf.rows.forEach(r => sf.headers.forEach((h, i) => { if (/\*$/.test(h) && !r[i]) { empty++; cols.add(h); } }));
      if (sf.rows.length) siteFile = { sites: sf.rows.length, emptyReq: empty, cols: [...cols] };
    }
    const unlisted = [];
    unlistedValues().forEach((m, ci) => { let rows = 0; m.forEach(c => { rows += c; }); unlisted.push({ column: bulkHeaders[ci], values: m.size, rows }); });
    return { emptyReq, unknownSites: [...unknownSiteVals], unknownCvs: [...unknownCvVals], pastStartDates, nameCollisions, collisionSamples, siteFile, unlisted };
  }

  function updateExportButton() {
    // Button stays ENABLED whenever there's data — clicking will surface a
    // confirmation when there are blockers, instead of leaving the user stuck.
    const btn = $('ts-export');
    btn.disabled = !(srcData && tplData && formattedRows && formattedRows.length);
    if (window.IMDebug) IMDebug.refresh('locations-standardize');
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
          '<div class="ts-confirm-issue-body">Create these Sites in PickTrace, re-download your personalized template, and re-upload here &mdash; or add your <b>Sites template</b> and the tool exports a Sites file for them alongside this one.</div>' +
          '</div>'
        );
      }
      if (b.unknownCvs.length) {
        // One line per crop with its varieties: "Grapes (new crop) — Variety A, Variety B".
        const groups = groupCropVarieties(new Map(b.unknownCvs.map(v => [v, 1])), tplData.dropdowns.get('crop & variety'));
        const sample = groups.slice(0, 8).map(g => {
          const names = g.vars.map(v => v.variety || '(no variety)');
          return '<li><b>' + escHtml(g.crop) + '</b>' + (g.isNew ? ' (new crop)' : '') + ' — ' +
            escHtml(names.slice(0, 6).join(', ')) + (names.length > 6 ? ', +' + (names.length - 6) + ' more' : '') + '</li>';
        }).join('');
        const more = groups.length > 8 ? '<li class="ts-confirm-more">… and ' + (groups.length - 8) + ' more crops</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unknownCvs.length + ' Crop &amp; Variety value(s) across ' +
            groups.length + ' crop' + (groups.length === 1 ? '' : 's') + ' not in your template dropdown</div>' +
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
      if (b.unlisted && b.unlisted.length) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unlisted.map(u => escHtml(u.column) + ': ' + u.values + ' value' + (u.values === 1 ? '' : 's') + ' on ' + u.rows + ' row' + (u.rows === 1 ? '' : 's')).join(' · ') +
            ' — lists your org hasn’t set up</div>' +
          '<div class="ts-confirm-issue-body">These are dropdowns in PickTrace with no options in your org yet, so it can reject the rows that use them. Create the values first (<b>Build in PickTrace</b> lists them), or leave the column out.</div>' +
          '</div>'
        );
      }
      if (b.siteFile && b.siteFile.emptyReq > 0) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.siteFile.emptyReq + ' empty required cell' + (b.siteFile.emptyReq === 1 ? '' : 's') + ' in the Sites file (' +
            b.siteFile.cols.map(escHtml).join(', ') + ')</div>' +
          '<div class="ts-confirm-issue-body">PickTrace rejects a site missing these &mdash; and then every location under it. Fill them under <b>Sites file</b> and <b>Site addresses</b>.</div>' +
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
    mapping = mapSource(bulkHeaders);
    // Reset manual fills + removed rows on a fresh format — they're tied
    // to the prior mapping/target context.
    manualFills = {};
    removedRows = new Set();
    cellOverrides = {};
    smartFixMap = {};
    siteFillOk = new Set();
    pvMode = 'all'; pvColsOnly = false; pvFocus = null;
    builtOnce = true;
    rebuildFormattedRows();
    renderMapping();
    renderPreview();
    renderRequired();
    updateSummary();
    $('ts-empty').style.display = 'none';
    $('ts-run').classList.remove('is-ready');
  }

  function doExport() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const numIdxs = numericColIdxs();
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
        // Numeric columns export as bare numbers (never "20.30 ac" / "3500
        // Vines"), and 0 acres exports as blank. Re-applied here so manual cell
        // edits are covered too.
        numIdxs.forEach(ci => {
          if (r[ci] !== '' && r[ci] != null) r[ci] = numericCell(r[ci], bulkHeaders[ci]);
        });
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

    // Each value goes under the template's own header of that name — the
    // header row is kept as the template has it, and PickTrace adds columns
    // between versions (Internal Commission), so placing by position could
    // shift every column after an inserted one. A column the template's header
    // row lacks keeps its position only where that cell of the header is blank.
    const tplHead = [];
    for (let c = 0; c <= range.e.c; c++) { const h = ws[XLSX.utils.encode_cell({ r: 0, c })]; tplHead[c] = h ? norm(String(h.v)) : ''; }
    const missing = [];
    const at = bulkHeaders.map((h, i) => {
      const j = tplHead.indexOf(norm(h));
      if (j >= 0) return j;
      if (!tplHead[i]) return i;
      missing.push(h);
      return -1;
    });
    if (missing.length && !confirm('The template’s header row has no column for: ' + missing.join(', ') + '.\n\nExport without those columns?')) return;
    const colCount = Math.max(range.e.c + 1, ...at.map(j => j + 1));
    exportRows.forEach((row, ri) => {
      for (let c = 0; c < bulkHeaders.length; c++) {
        const val = row[c];
        if (val == null || val === '' || at[c] < 0) continue;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c: at[c] });
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

    // "Acme Farms - Locations - 2026-10-07.xlsx" (see IMOutName in utils.js);
    // an Update template says so, since it's uploaded somewhere else.
    const exportName = IMOutName.fileName(target === 'update' ? 'Locations Update' : 'Locations');

    // bookSST:true is NOT optional. Without it SheetJS writes string cells as
    // t="str", which OOXML defines as a cached formula result rather than
    // literal text; strict readers hand back nothing and the upload is
    // rejected with "unexpected column header". (Sites goes further and
    // patches the template zip directly — see xlsx-template-writer.js — which
    // also keeps the <dataValidations> SheetJS drops on a round-trip. This
    // export still writes numeric columns as numbers, so it has not been moved
    // over yet.)
    const siteRows = siteFileOn() ? siteRowsForExport() : [];
    if (!siteRows.length) { XLSX.writeFile(wb, exportName, { cellStyles: true, bookSST: true }); return; }
    // With new sites: both files in one .zip folder (as Bulk's Download All),
    // the Sites file written by patching its template like the Sites tool does.
    const folder = IMOutName.folderName();
    const locBytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true, bookSST: true }));
    IMXlsxTemplate.build({ rawBuffer: siteTpl.rawBuffer, headers: siteTpl.headers, rows: siteRows })
      .then(blob => blob.arrayBuffer())
      .then(buf => {
        const zip = IMXlsxTemplate.bundle([
          { name: folder + '/' + IMOutName.fileName('Sites'), data: new Uint8Array(buf) },
          { name: folder + '/' + exportName, data: locBytes }]);
        const url = URL.createObjectURL(zip);
        const a = document.createElement('a');
        a.href = url;
        a.download = folder + '.zip';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
      })
      .catch(err => alert('Could not write the Sites file: ' + (err && err.message ? err.message : err)));
  }

  // ─── Load handlers ───
  function handleSrcFile(file) {
    readSrcFile(file).then(data => {
      srcData = data;
      IMOutName.suggestClient(file.name);
      $('ts-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      const hrowNote = data.headerRow > 0 ? ' (skipped ' + data.headerRow + ' metadata row' + (data.headerRow === 1 ? '' : 's') + ' above)' : '';
      $('ts-src-meta').textContent = data.rows.length + ' rows · ' + data.headers.length + ' columns · sheet: ' + data.sheetName + hrowNote;
      formattedRows = null;
      // The first build waits for Format Sheet, so the optional files can go in
      // first; after that a new file rebuilds straight away.
      if (srcData && tplData && builtOnce) runFormat(); else showReady();
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
      const sites = new Set(); // every Site the account already has (it has locations)
      data.rows.forEach(r => {
        const s = r[siteI], n = r[nameI];
        if (String(s == null ? '' : s).trim()) sites.add(String(s).trim());
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
      existingData = { keys, blockKeys, byName, sites, hasCrop: cvI >= 0, fileName: file.name, count: keys.size };
      $('ts-existing-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('ts-existing-meta').textContent = keys.size + ' existing location' + (keys.size === 1 ? '' : 's') + ' indexed' +
        (cvI >= 0 ? ' (block-level)' : ' (name-only — no Crop column)');
      if (formattedRows) {
        rebuildFormattedRows();
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      } else showReady();
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

  // ─── Merged duplicates (same Site + Name, folded into one location) ───
  function renderMerged() {
    const sec = $('ts-section-merged');
    if (!sec) return;
    const cb = $('ts-merge-dupes');
    if (cb) cb.checked = mergeDupes;
    // The panel is the only place the toggle lives, so keep it visible whenever
    // there's something to merge — including when the user has turned it off.
    const groups = mergeGroups;
    const hasDupes = groups.size > 0 || unmergedKeys.size > 0 || !mergeDupes;
    if (!formattedRows || !hasDupes) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const bulkHeaders = activeSchemaHeaders();
    const siteIdx = bulkHeaders.indexOf('Site*');
    const nameIdx = bulkHeaders.indexOf('Name*');
    const acIdx = bulkHeaders.indexOf('Acreage');
    const t = $('ts-merged-title');
    if (t) {
      t.textContent = mergeDupes
        ? 'Merged Duplicates (' + groups.size + ' name' + (groups.size === 1 ? '' : 's') +
          ' collapsed, ' + mergedRowIdxs.size + ' row' + (mergedRowIdxs.size === 1 ? '' : 's') + ' folded)'
        : 'Merged Duplicates (off)';
    }
    if (!groups.size) { $('ts-merged-table').innerHTML = ''; return; }
    let html = '<thead><tr><th>Site</th><th>Name</th><th>Rows merged</th>' +
      '<th>Acreage kept</th><th>Conflicting columns</th><th></th></tr></thead><tbody>';
    [...groups.entries()].slice(0, 200).forEach(([k, g]) => {
      const r = formattedRows[g.keep];
      const conf = [...g.conflicts.entries()]
        .map(([ci, vals]) => escHtml(bulkHeaders[ci]) + ' <span class="text-muted small">(kept ' +
          escHtml(vals[0]) + ' of ' + vals.length + ')</span>')
        .join(', ');
      html += '<tr>' +
        '<td>' + escHtml(r[siteIdx]) + '</td>' +
        '<td><b>' + escHtml(r[nameIdx]) + '</b></td>' +
        '<td>' + (g.folded.length + 1) + ' &rarr; 1</td>' +
        '<td>' + escHtml(acIdx >= 0 ? r[acIdx] : '') + '</td>' +
        '<td>' + (conf || '<span class="text-muted small">none</span>') + '</td>' +
        '<td><button class="btn btn-ghost btn-sm ts-merge-undo" data-key="' + escHtml(k) + '">Unmerge</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    $('ts-merged-table').innerHTML = html;
    $('ts-merged-table').querySelectorAll('.ts-merge-undo').forEach(btn => {
      btn.addEventListener('click', e => {
        unmergedKeys.add(e.target.dataset.key);
        refreshAll();
      });
    });
    const restore = $('ts-merged-restore');
    if (restore) {
      restore.style.display = unmergedKeys.size ? '' : 'none';
      restore.textContent = 'Re-merge ' + unmergedKeys.size + ' unmerged name' + (unmergedKeys.size === 1 ? '' : 's');
    }
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

  // Re-derive everything from source and repaint every panel. Used by any
  // control that changes what the export will contain.
  function refreshAll() {
    if (!srcData || !tplData) return;
    rebuildFormattedRows();
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
  }

  function applyRename(ri, newName) {
    nameOverrides[ri] = newName;
    refreshAll();
  }

  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      const parsed = parseTemplate(e.target.result, file.name);
      if (!parsed || !parsed.headers.length) {
        alert('Could not find DATA ENTRY sheet in this template.');
        return;
      }
      // Wrong-template guard. A positively-identified OTHER template is
      // refused outright — there's no case where formatting locations into it
      // is correct. An unrecognized header row only warns, so a future or
      // customized Locations template that misses the signature can still be
      // used deliberately.
      const kind = identifyTemplate(parsed.headers);
      // The Sites template has a slot of its own here (new sites → Sites file).
      if (kind.kind === 'sites') { useSiteTemplate(parsed, file.name); return; }
      // A locations EXPORT has the template's DATA ENTRY sheet but no
      // DROP-DOWN INPUTS: nothing to check Site, Location Type or Crop & Variety
      // against, and its header row isn't the template's (it adds Internal
      // Commission), so rows written under it would land a column off. It
      // belongs in the Existing locations slot.
      if (kind.kind === 'locations' && !parsed.sheetNames.some(n => /drop.?down/i.test(n))) {
        const free = !existingData;
        alert('That’s a locations export, not the personalized bulk template — it has no DROP-DOWN INPUTS sheet, so there would be no Site, ' +
          'Location Type or Crop & Variety lists to pick from or check against, and its columns don’t line up with the upload.\n\n' +
          (free ? 'It’s been added as Existing locations (rows already in PickTrace are dropped). ' : '') +
          'For slot 2, download the bulk Locations template from PickTrace (the one with the DROP-DOWN INPUTS sheet).');
        if (free) handleExistingFile(file);
        return;
      }
      if (kind.kind === 'employees') {
        alert('Wrong template — this is the ' + kind.label + ' bulk template.\n\n' +
          'This tool builds the LOCATIONS template. Loading a ' + kind.label +
          ' template writes location data under the wrong headers' +
          (kind.kind === 'sites'
            ? ' — Location Type lands in Employer*, Crop & Variety in Address1*, ' +
              'Acreage in City* and Start Date in Country*'
            : '') +
          ', and the upload fails validation.\n\n' +
          'Load the Locations bulk template instead' +
          (kind.kind === 'sites' ? ', or use the Sites Standardize tool for sites.' : '.'));
        return;
      }
      if (kind.kind === 'unknown') {
        const ok = confirm('This doesn’t look like a PickTrace Locations bulk template — ' +
          'the DATA ENTRY header row is missing Site, Name, Location Type or Crop & Variety.\n\n' +
          'Rows are written positionally, so if the header row is wrong every column ' +
          'lands in the wrong place.\n\nLoad it anyway?');
        if (!ok) return;
      }
      tplData = parsed;
      readRestricted(e.target.result, parsed.headers).then(set => {
        parsed.restricted = set;
        if (tplData === parsed && formattedRows) refreshAll();
      });
      $('ts-tpl-name').textContent = file.name;
      $('ts-tpl-meta').innerHTML =
        (kind.kind === 'locations' ? 'Locations template · ' : '') +
        parsed.headers.length + ' columns · ' +
        (parsed.dropdowns.get('site') ? parsed.dropdowns.get('site').size : 0) + ' sites · ' +
        (parsed.dropdowns.get('crop & variety') ? parsed.dropdowns.get('crop & variety').size : 0) + ' C&V values' +
        (kind.kind === 'locations' ? '' : ' <span class="cmp-warn">⚠ not a Locations template</span>');
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
      if (srcData && tplData && builtOnce) runFormat(); else showReady();
    };
    r.readAsArrayBuffer(file);
  }

  // Before the first build: the Format button, and what's still missing.
  function showReady() {
    $('ts-run').disabled = !(srcData && tplData);
    $('ts-run').classList.toggle('is-ready', !!(srcData && tplData && !formattedRows));
    if (formattedRows) return;
    const empty = $('ts-empty');
    empty.style.display = '';
    empty.innerHTML = srcData && tplData
      ? '<b>Ready.</b> Add any optional files (existing locations, the Sites template), then press <b>&#9654; Format Sheet</b>.'
      : 'Upload a source file AND a personalized bulk template' + (srcData || tplData ? ' — ' + (srcData ? 'the template' : 'the source file') + ' is still missing.' : ' to start.');
    if (wb) wb.refresh();
  }

  // ─── Reset ───
  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    IMOutName.forget();
    manualFills = {}; removedRows = new Set(); siteAddresses = {};
    siteTpl = null; siteFill = {}; siteAddrAll = ''; siteAddrGroup = {}; siteEdits = {}; pvView = 'locations'; siteGrid = [];
    builtOnce = false;
    { const n = $('ts-sitetpl-name'); if (n) n.textContent = 'No file selected'; }
    { const m = $('ts-sitetpl-meta'); if (m) m.textContent = ''; }
    smartFixMap = {}; pvMode = 'all'; pvColsOnly = false; pvFocus = null; siteFillOk = new Set(); siteGaps = [];
    dateFill = { mm: '', dd: '', range: 'first' };
    existingData = null; existingRowIdxs = new Set();
    nameOverrides = {}; cellOverrides = {}; collisionKeys = new Set(); existingTakenKeys = new Set();
    mergeDupes = true; mergedRowIdxs = new Set(); mergeGroups = new Map(); unmergedKeys = new Set();
    { const md = $('ts-merge-dupes'); if (md) md.checked = true; }
    $('ts-src-name').textContent = 'No file selected';
    $('ts-tpl-name').textContent = 'No file selected';
    $('ts-src-meta').textContent = '';
    $('ts-tpl-meta').textContent = '';
    { const en = $('ts-existing-name'); if (en) en.textContent = 'No file selected'; }
    { const em = $('ts-existing-meta'); if (em) em.textContent = ''; }
    { const ef = $('ts-existing-file'); if (ef) ef.value = ''; }
    $('ts-src-file').value = ''; $('ts-tpl-file').value = '';
    ['ts-section-mapping','ts-section-smartfix','ts-section-sites-create','ts-section-sitefile','ts-section-siteaddr','ts-section-cv-create','ts-section-required','ts-section-existing','ts-section-merged','ts-section-collisions','ts-section-datefill','ts-section-preview']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    $('ts-summary').style.display = 'none';
    $('ts-export').disabled = true;
    lastSum = { pastDates: 0, toCreate: 0, collisionRows: 0 };
    if (wb) wb.reset();
    showReady();
    if (window.IMDebug) IMDebug.refresh('locations-standardize');
  }

  // ─── Debug dump ───
  function collectDebug() {
    if (!srcData || !tplData || !formattedRows) return null;
    const H = activeSchemaHeaders();
    const b = getExportBlockers() || {};
    const siteIdx = H.indexOf('Site*');
    const nameIdx = H.indexOf('Name*');
    const cvIdx = H.indexOf('Crop & Variety*');
    const keptIdxs = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const keptRows = keptIdxs.map(i => formattedRows[i]);

    const cols = H.map((h, i) => {
      const si = mapping[i] != null ? mapping[i] : -1;
      return { index: i, header: h, required: /\*$/.test(h), srcIndex: si,
        match: si >= 0 ? 'mapped' : 'unmapped',
        sample: si >= 0 && srcData.rows[0] ? srcData.rows[0][si] : null };
    });

    const asked = [];
    const emptyReqCols = H.map((h, i) => {
      if (!/\*$/.test(h)) return null;
      let n = 0; keptIdxs.forEach(ri => { if (!formattedRows[ri][i]) n++; });
      return n ? { column: h, emptyRows: n } : null;
    }).filter(Boolean);
    asked.push(IMDebug.ask('empty-required', 'required', 'Required columns with empty cells', {
      count: emptyReqCols.length, blocksExport: true,
      detail: (b.emptyReq || 0) + ' empty required cells across ' + keptIdxs.length + ' exported rows.',
      items: emptyReqCols }));
    asked.push(IMDebug.ask('unknown-sites', 'dropdown', 'Site values not in the template dropdown', {
      count: (b.unknownSites || []).length, blocksExport: true,
      detail: 'These sites must exist in PickTrace before the locations can be created.',
      items: b.unknownSites || [] }));
    asked.push(IMDebug.ask('unknown-crop-varieties', 'dropdown', 'Crop & Variety values not in the template dropdown', {
      count: (b.unknownCvs || []).length, blocksExport: true, items: b.unknownCvs || [] }));
    asked.push(IMDebug.ask('name-collisions', 'collision', 'Rows sharing Site + Name', {
      count: b.nameCollisions || 0, blocksExport: true, items: b.collisionSamples || [] }));
    asked.push(IMDebug.ask('past-start-dates', 'data', 'Start Date values before today', {
      count: b.pastStartDates || 0, blocksExport: false,
      detail: 'PickTrace treats a past start date as already-planted.' }));
    asked.push(IMDebug.ask('unlisted-values', 'dropdown', 'Values in columns whose template list is empty (org has none set up)', {
      count: (b.unlisted || []).length, blocksExport: true, items: b.unlisted || [] }));
    const sfPending = computeSmartFixes();
    asked.push(IMDebug.ask('smart-fixes', 'smartfix', 'Off-list values awaiting a Smart Fix choice', {
      count: sfPending.length, blocksExport: false,
      items: sfPending.map(f => ({ column: f.header, value: f.from, rows: f.count, suggestion: f.suggestion || null })) }));

    const merges = [];
    mergeGroups.forEach((g, k) => {
      if (unmergedKeys.has(k)) return;
      const conflicts = [];
      if (g.conflicts) g.conflicts.forEach((vals, ci) => conflicts.push({ column: H[ci], values: vals }));
      merges.push({ key: k, keptRowIndex: g.keep, foldedRowIndexes: g.folded,
        site: formattedRows[g.keep] ? formattedRows[g.keep][siteIdx] : null,
        name: formattedRows[g.keep] ? formattedRows[g.keep][nameIdx] : null,
        conflictingColumns: conflicts });
    });

    const prov = IMDebug.deriveProvenance({
      headers: H, rows: formattedRows, srcRows: srcData.rows, colToSrc: mapping,
      fills: manualFills, cellEdits: cellOverrides, nameEdits: nameOverrides,
      smartFixes: smartFixMap, nameCol: nameIdx });
    const out = IMDebug.output(H, keptRows, {
      note: 'exportedRowIndexes[n] is the formattedRows index behind exported row n — the key used by provenance.codes, edits, merges and dropped.' });
    out.exportedRowIndexes = keptIdxs;

    return {
      settings: {
        target: target,
        targetSchema: target === 'update' ? 'UPDATE' : 'CREATE',
        cropSplitEnabled: cropSplit,
        mergeDuplicatesEnabled: mergeDupes,
        dateFill: dateFill
      },
      inputs: [
        IMDebug.file('source data', srcData),
        IMDebug.file('personalized bulk Locations template', tplData, {
          dropdowns: tplData.dropdowns ? [...tplData.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : [] }),
        IMDebug.file('existing locations in PickTrace', existingData, {
          indexedKeys: existingData ? existingData.count : null }),
        IMDebug.file('personalized bulk Sites template (new sites)', siteTpl, {
          dropdowns: siteTpl ? [...siteTpl.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : [] })
      ],
      mapping: IMDebug.mapping(cols, srcData.headers, { srcRows: srcData.rows }),
      derived: [
        { kind: 'crop/variety split', enabled: cropSplit,
          note: 'When on, " - Subvariety" is split off Crop & Variety into Clone/Subvariety.' },
        { kind: 'year-only date expansion', settings: dateFill,
          note: 'Year-only Start Dates are expanded to a full date using the month/day above.' },
        { kind: 'sites file', enabled: siteFileOn(), newSites: newSitesMap().size,
          note: 'New sites (not in the Locations template Site list) built into a Sites upload: address = own > ranch > every site, then fills, one-option lists, edits.',
          fills: siteFill, addressForEverySite: siteAddrAll || null,
          addressByRanch: siteAddrGroup, addressBySite: siteAddresses, edits: siteEdits,
          output: siteFileOn() ? (() => { const sf = buildSiteFile(); return { headers: sf.headers, rows: sf.rows.slice(0, 500), provenance: sf.prov.slice(0, 500) }; })() : null }
      ],
      fills: Object.keys(manualFills).map(i => ({ column: H[+i], value: manualFills[i], kind: 'column fill' }))
        .concat(Object.keys(smartFixMap).map(k => {
          const sep = k.indexOf('||');
          return { column: H[+k.slice(0, sep)], from: k.slice(sep + 2), value: smartFixMap[k], kind: 'smart fix' };
        })),
      edits: {
        cells: Object.keys(cellOverrides).map(k => {
          const sep = k.indexOf('|');
          return { rowIndex: +k.slice(0, sep), column: H[+k.slice(sep + 1)], value: cellOverrides[k] };
        }),
        names: Object.keys(nameOverrides).map(ri => ({ rowIndex: +ri, value: nameOverrides[ri] }))
      },
      merges: merges,
      dropped: {
        removedByHand: { count: removedRows.size, rowIndexes: [...removedRows].slice(0, 500) },
        mergedIntoAnotherRow: { count: mergedRowIdxs.size, rowIndexes: [...mergedRowIdxs].slice(0, 500) },
        alreadyInPickTrace: {
          count: existingRowIdxs.size,
          source: existingData ? existingData.fileName : null,
          samples: [...existingRowIdxs].slice(0, 200).map(ri => ({
            site: formattedRows[ri][siteIdx], name: formattedRows[ri][nameIdx], cropVariety: formattedRows[ri][cvIdx] }))
        }
      },
      asked: asked,
      output: out,
      provenance: prov
    };
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    initialized = true;
    mountWorkbench();
    if (window.IMDebug) {
      IMDebug.register('locations-standardize', {
        label: 'Locations Standardize',
        ready: () => !!(srcData && tplData && formattedRows),
        collect: collectDebug
      });
      IMDebug.wire('ts-debug', 'locations-standardize');
    }
    attachModeSwitcher();
    $('ts-src-file').addEventListener('change', e => {
      if (e.target.files[0]) handleSrcFile(e.target.files[0]);
      e.target.value = '';
    });
    $('ts-tpl-file').addEventListener('change', e => {
      if (e.target.files[0]) handleTplFile(e.target.files[0]);
      e.target.value = '';
    });
    { const ef = $('ts-existing-file'); if (ef) ef.addEventListener('change', e => { if (e.target.files[0]) handleExistingFile(e.target.files[0]); e.target.value = ''; }); }
    { const sf = $('ts-sitetpl-file'); if (sf) sf.addEventListener('change', e => { if (e.target.files[0]) handleSiteTplFile(e.target.files[0]); e.target.value = ''; }); }
    // Lists your org hasn't set up: leave a column out (unmapped — Column
    // Mapping can bring it back), or copy Column / Value / Rows for PickTrace.
    { const lt = $('ts-lists-create-table'); if (lt) lt.addEventListener('click', e => {
        const b = e.target.closest('.ts-list-drop');
        if (!b) return;
        mapping[+b.dataset.ci] = -1;
        renderMapping();
        refreshAll();
      }); }
    { const lc = $('ts-lists-copy'); if (lc) lc.addEventListener('click', () => {
        const H = activeSchemaHeaders(), lines = [];
        unlistedValues().forEach((m, ci) => m.forEach((c, v) => lines.push(H[ci] + '\t' + v + '\t' + c)));
        if (!lines.length) return;
        navigator.clipboard.writeText('Column\tValue\tRows\n' + lines.join('\n'))
          .then(() => alert('Copied ' + lines.length + ' values.'), () => alert('Copy failed.'));
      }); }
    // Sites file: a column fill for every new site (a pick, or text on change).
    { const body = $('ts-sitefile-body'); if (body) {
      body.addEventListener('change', e => {
        const el = e.target.closest('.ts-sf-fill');
        if (!el) return;
        const v = el.value.trim();
        if (v) siteFill[el.dataset.k] = v; else delete siteFill[el.dataset.k];
        refreshSites();
      });
      body.addEventListener('click', e => {
        const b = e.target.closest('.ts-sf-clear');
        if (!b) return;
        delete siteFill[b.dataset.k];
        refreshSites();
      });
    } }
    // Site addresses: every site / per ranch / one site — kept on change.
    { const body = $('ts-siteaddr-body'); if (body) body.addEventListener('change', e => {
        const el = e.target.closest('.ts-addr-in');
        if (!el) return;
        const v = el.value.trim(), k = el.dataset.key;
        if (el.dataset.scope === 'all') siteAddrAll = v;
        else if (el.dataset.scope === 'ranch') { if (v) siteAddrGroup[k] = v; else delete siteAddrGroup[k]; }
        else { if (v) siteAddresses[k] = v; else delete siteAddresses[k]; }
        refreshSites();
      }); }
    { const cs = $('ts-crop-split'); if (cs) cs.addEventListener('change', e => {
        cropSplit = !!e.target.checked;
        refreshAll();
      }); }
    { const md = $('ts-merge-dupes'); if (md) md.addEventListener('change', e => {
        mergeDupes = !!e.target.checked;
        if (mergeDupes) unmergedKeys = new Set();  // re-enabling starts clean
        refreshAll();
      }); }
    { const mr = $('ts-merged-restore'); if (mr) mr.addEventListener('click', () => {
        unmergedKeys = new Set();
        refreshAll();
      }); }
    { const a = $('ts-datefill-apply'); if (a) a.addEventListener('click', applyDateFill); }
    { const c = $('ts-datefill-clear'); if (c) c.addEventListener('click', clearDateFill); }
    document.querySelectorAll('input[name="ts-target"]').forEach(r => {
      r.addEventListener('change', e => {
        target = e.target.value;
        if (srcData && tplData && builtOnce) runFormat();
      });
    });
    $('ts-run').addEventListener('click', runFormat);
    $('ts-export').addEventListener('click', async () => {
      const b = getExportBlockers();
      if (b && (b.emptyReq > 0 || b.unknownSites.length || b.unknownCvs.length || b.nameCollisions > 0 || b.pastStartDates > 0 ||
          (b.siteFile && b.siteFile.emptyReq > 0) || (b.unlisted && b.unlisted.length))) {
        const proceed = await showExportConfirm(b);
        if (!proceed) return;
      }
      doExport();
    });
    $('ts-reset').addEventListener('click', reset);
    { const sfa = $('ts-smartfix-apply-all'); if (sfa) sfa.addEventListener('click', applyAllSmartFixes); }
    $('ts-sites-copy').addEventListener('click', () => {
      const tbl = $('ts-sites-create-table');
      const ranches = ranchMap(newSiteNames());
      const rows = [...tbl.querySelectorAll('tbody tr')]
        .map(tr => {
          const val  = tr.children[0] ? tr.children[0].textContent.trim() : '';
          const cnt  = tr.children[1] ? tr.children[1].textContent.trim() : '';
          return val + '\t' + cnt + '\t' + siteAddrFor(val, ranches.get(val)).text;
        });
      if (!rows.length) return;
      navigator.clipboard.writeText('Site\tRow count\tAddress\n' + rows.join('\n'))
        .then(() => alert('Copied ' + rows.length + ' sites.'), () => alert('Copy failed.'));
    });
    // Click-to-copy any cell (mirrors the Compare page). The Empty Required
    // Columns section is form-only, so it's excluded.
    $('ts-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      if (!td) return;
      if (td.closest('#ts-section-required')) return;
      if (e.target.closest('input, button, label, select, textarea, a')) return;
      // A cell can name what it copies (the crop without its "new crop" tag).
      const text = (td.dataset.copy != null ? td.dataset.copy : td.textContent || '').trim();
      if (!text) return;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
          .then(() => td.classList.toggle('cmp-copied'), () => {});
      }
    });
    // Same shape as the table: Crop | Variety | Row count, the crop written
    // once above its varieties, so it pastes into a sheet the way it reads.
    // Blank Site (in Smart Fixes): copy a group's site down, undo it, or all.
    { const sg = $('ts-sitegaps'); if (sg) sg.addEventListener('click', e => {
      const btn = e.target.closest('button[data-act]');
      if (!btn) return;
      if (btn.dataset.act === 'all') siteGaps.forEach(g => siteFillOk.add(g.id));
      else if (btn.dataset.act === 'apply') siteFillOk.add(btn.dataset.id);
      else siteFillOk.delete(btn.dataset.id);
      rebuildFormattedRows(); renderPreview(); renderRequired(); renderSitesAndCvToCreate(); updateSummary();
    }); }
    $('ts-cv-copy').addEventListener('click', () => {
      if (!cvGroups.length) return;
      const lines = ['Crop\tVariety\tRow count'];
      cvGroups.forEach(g => g.vars.forEach((v, i) => lines.push((i === 0 ? g.crop : '') + '\t' + v.variety + '\t' + v.count)));
      navigator.clipboard.writeText(lines.join('\n'))
        .then(() => alert('Copied ' + cvGroups.length + ' crop(s) and ' + (lines.length - 1) + ' variet' +
          (lines.length === 2 ? 'y' : 'ies') + '.'), () => alert('Copy failed.'));
    });
  }

  // Init when the parent page is opened.
  window.tsInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
