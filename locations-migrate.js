// ═══════════════════════════════════════════════════════════════════════
// Legacy → 3.0 Sites & Locations Migration — one combined module that turns
// a Legacy PickTrace export pair (pt_locations.csv + a sites export) into the
// two 3.0 bulk-create workbooks: Sites and Locations.
//
// Legacy stores a location's Site, Crop and Location Type as opaque numeric
// IDs (site_id / crop_id / location_type_id) with no names in the export, so
// this module resolves them three ways, in order of trust:
//   1. site_id  → the sites export's ID → Site column (a hard join).
//   2. crop_id / location_type_id → an optional Legacy lookup CSV (ID → Name),
//      then snapped to the template's dropdown literal.
//   3. Otherwise a suggestion inferred from each ID's own location names, which
//      the user confirms once in the ID Mapping panel (sticky across rebuilds).
//
// Archived rows (is_archived = true) are dropped and counted in the summary.
//
// Output is written by patching the ORIGINAL template zip in place (see
// "XLSX writing at the zip level" below) rather than round-tripping through
// SheetJS. The bundled xlsx-js-style build does not preserve <dataValidations>
// and rebuilds styles.xml / workbook.xml from its own model, which silently
// stripped all 8 dropdown validations and most of the header styling from this
// template. Patching the zip keeps every part byte-for-byte except the DATA
// ENTRY sheet's <sheetData>.
//
// Mirrors the structure of employee-migrate.js (Legacy → 3.0 Employee
// Migration) — note that module still exports via SheetJS and therefore still
// loses validations; this one does not.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  // Site names are referenced across two files (Sites file → Locations' Site*),
  // so runs of spaces are folded to one in both: a double space in a Legacy
  // site name would otherwise hinge on PickTrace keeping it.
  const cleanName = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  // ─── Diagnostics for the Debug Dump ───
  // Every message this tool shows you, and every error thrown while it is on
  // screen, is kept here, so a dump taken after something breaks (even before
  // Build Preview has ever worked) still says what happened and where.
  const MODULE_FILE = 'locations-migrate.js';
  const events = [];
  function logEvent(kind, msg, err) {
    events.push({
      at: new Date().toISOString(),
      kind: kind,
      message: String(msg == null ? '' : msg).slice(0, 2000),
      stack: err && err.stack ? String(err.stack).split('\n').slice(0, 15).join('\n') : null
    });
    if (events.length > 200) events.shift();
    if (window.IMDebug) IMDebug.refresh('locations-migrate');
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

  // 3.0 Locations Bulk Create — DATA ENTRY columns, exact order.
  const LOC_HEADERS = [
    'Site*', 'Name*', 'Alt ID', 'Location Type*', 'Crop & Variety*', 'Planted At',
    'Acreage', 'Length', 'Plant Count', 'Start Date*', 'Location Group',
    'Lot Number', 'Rootstock', 'Plant/Seed Seller', 'Plant/Seed Producer',
    'Clone/Subvariety', 'Growing Manager', 'Production Status', 'Training Style',
    'Trellis Type', 'Mulch Type', 'Row Direction', 'Organic Status',
    'Organic Certifier', 'Wet Date', 'Germination Date', 'Planted Date',
    'Grafting Date', 'Production Start', 'Organic Certification Date',
    'Stand Count', 'Row/Bed Count', 'Post Count', 'Percent Covered',
    'Row Spacing, in.', 'Plant Spacing, in.', 'Post Spacing, in.', 'Bed Width, in.',
    'Custom Data 1', 'Custom Data 2', 'Custom Data 3'
  ];
  // (Internal Commission used to close this list. It's a column of the
  // locations EXPORT, not of the bulk template, so it never had a slot to go in.)

  // 3.0 Sites Bulk Create — same schema sites-standardize.js targets.
  const SITE_HEADERS = ['Name*', 'Alt ID', 'Site Type*', 'Employer*', 'Address1*',
    'Address2', 'City*', 'State*', 'Zip*', 'Country*'];

  // Column indexes used often enough to name.
  const L_SITE = 0, L_NAME = 1, L_ALT = 2, L_TYPE = 3, L_CROP = 4;
  const S_NAME = 0, S_ALT = 1;

  const HECTARES_TO_ACRES = 2.4710538147;
  const METERS_TO_FEET = 3.2808398950;

  // ─── Legacy crop_id preset ───────────────────────────────────────────
  // Legacy exports crop_id as a bare integer with no name anywhere in the
  // file. These IDs are assigned per tenant, so the same number means a
  // different crop in a different org — applying this table blindly elsewhere
  // would silently mis-map an entire migration.
  //
  // The names below were recovered from this tenant's own Legacy Locations
  // grid and cross-checked against all 220 unarchived rows: every crop_id
  // resolved to exactly one name with zero conflicts. `target` is the 3.0
  // Crop & Variety literal; it is only used when the uploaded template
  // actually offers it, otherwise the value is re-derived from `name`.
  //
  // NOTHING here is applied unless PRESET_FINGERPRINT matches the loaded file
  // (see presetMatches) — on any other export the preset stays inert and the
  // mapping panel behaves normally.
  const CROP_PRESET = {
    id: 'legacy-crop-map',
    label: 'Legacy crop map (recovered from this tenant\'s Locations grid)',
    crops: {
      '1':  { name: 'blueberries',      target: 'Blueberries-Other' },
      '2':  { name: 'strawberries',     target: 'Strawberries-Other' },
      '7':  { name: 'citrus',           target: 'Citrus-Other' },
      '9':  { name: 'apples',           target: 'Apples-Other' },
      '10': { name: 'cherries',         target: 'Cherries-Other' },
      '24': { name: 'office',           target: 'Other-Other' },
      '27': { name: 'shop',             target: 'Other-Other' },
      '29': { name: 'pears',            target: 'Pears-Other' },
      '34': { name: 'general',          target: 'Various Crops-Other' },
      '35': { name: 'lettuce',          target: 'Lettuce-Other' },
      '36': { name: 'broccoli',         target: 'Broccoli-Other' },
      '37': { name: 'cauliflower',      target: 'Cauliflower-Other' },
      '38': { name: 'cabbage',          target: 'Cabbage-Other' },
      '39': { name: 'spinach',          target: 'Spinach-Other' },
      '40': { name: 'garlic',           target: 'Garlic-Other' },
      '41': { name: 'kale',             target: 'Green Kale-Other' },
      '44': { name: 'berries',          target: 'Summer Berries-Other' },
      '45': { name: 'vegetables',       target: 'Veggies-Other' },
      '46': { name: 'bell peppers',     target: 'Bell Peppers-Other' },
      '47': { name: 'oranges',          target: 'Oranges-Other' },
      '48': { name: 'grapefruit',       target: 'Grapefruits-Other' },
      '51': { name: 'brussels sprouts', target: 'Brussel Sprouts-Other' },
      '55': { name: 'grape leaves',     target: 'Grapes-Other' }
    },
    // crop_ids that exist only on archived rows, so they never reach an export
    // and were never observable in the grid. Listed so the coverage check below
    // doesn't treat them as evidence of a different tenant.
    archivedOnly: ['4', '16', '19', '21', '30', '32'],
    // Site IDs this tenant's locations reference.
    siteIds: ['1302', '2242', '2244', '343', '4', '5705', '5930'],
    // Content spot-checks: crop_id → a word that must appear in that ID's own
    // location text. Cheap to satisfy for the real file, essentially
    // impossible to satisfy by coincidence in someone else's export.
    probes: {
      '36': 'broccoli', '38': 'cabbage', '37': 'coliflor',
      '2': 'fresa', '40': 'ajo', '35': 'lettuce'
    },
    minProbes: 4
  };

  // True when the loaded locations file is demonstrably this tenant's export.
  // Requires: every site_id known, every crop_id known, and at least
  // minProbes content spot-checks passing.
  function presetMatches() {
    if (!locData) return null;
    const known = new Set(Object.keys(CROP_PRESET.crops).concat(CROP_PRESET.archivedOnly));
    const sites = new Set(), crops = new Set();
    const corpus = {};
    locData.rows.forEach(row => {
      const s = locGet(row, 'site_id'); if (s) sites.add(s);
      const c = locGet(row, 'crop_id'); if (c) crops.add(c);
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
      const txt = [locGet(row, 'display_name'), locGet(row, 'description'),
                   locGet(row, 'plant_name')].filter(Boolean).join(' ').toLowerCase();
      corpus[c] = (corpus[c] || '') + ' ' + txt;
    });
    const siteOk = [...sites].every(s => CROP_PRESET.siteIds.indexOf(s) >= 0);
    const cropOk = [...crops].every(c => known.has(c));
    let probes = 0, probesPossible = 0;
    Object.keys(CROP_PRESET.probes).forEach(cid => {
      if (corpus[cid] == null) return;
      probesPossible++;
      if (corpus[cid].indexOf(CROP_PRESET.probes[cid]) >= 0) probes++;
    });
    const ok = siteOk && cropOk && probes >= Math.min(CROP_PRESET.minProbes, probesPossible) && probesPossible > 0;
    return { ok, siteOk, cropOk, probes, probesPossible,
             unknownCrops: [...crops].filter(c => !known.has(c)),
             unknownSites: [...sites].filter(s => CROP_PRESET.siteIds.indexOf(s) < 0) };
  }

  // Resolve a preset entry against the CURRENT template's dropdown. The stored
  // target wins when the template offers it; otherwise fall back to matching on
  // the legacy name so the preset still helps on a differently-configured
  // template.
  function presetValueFor(cid) {
    const entry = CROP_PRESET.crops[cid];
    if (!entry) return null;
    const opts = locOptions('Crop & Variety*');
    if (!opts.length) return null;
    const cm = buildCaseMap(opts);
    const direct = cm.get(String(entry.target).toUpperCase().trim());
    if (direct != null) return direct;
    const s = suggestFromName(entry.name, opts);
    return s ? s.value : null;
  }

  function applyCropPreset() {
    let n = 0, missed = [];
    collectIds('crop').forEach((info, cid) => {
      const v = presetValueFor(cid);
      if (v) { cropMap[cid] = cropPart(v); n++; }
      else missed.push(cid + (CROP_PRESET.crops[cid] ? ' (' + CROP_PRESET.crops[cid].name + ')' : ''));
    });
    presetApplied = n > 0;
    return { n, missed };
  }

  // ─── State ───
  let locData = null;        // { headers, rows, idx, fileName } — pt_locations.csv
  let sitesData = null;      // { headers, rows, idx, fileName, byId:Map } — sites export
  let cropLookup = null;     // Map<crop_id, legacy crop name> — optional CSV
  let typeLookup = null;     // Map<location_type_id, legacy type name> — optional CSV
  let migratedSites = null;  // Set<UPPER site name> already in 3.0 — optional
  let locTpl = null;         // { headers, dropdowns, rawBuffer, fileName }
  let siteTpl = null;        // { headers, dropdowns, rawBuffer, fileName }

  let cropMap = {};          // crop_id → the crop's 3.0 name, e.g. 'Blueberries' (sticky)
  let varAnswered = new Set(); // varieties the client already named (their answers page) — not asked again
  let varietyMap = {};       // UPPER Legacy plant_name → the 3.0 variety name (sticky; default: the Legacy name tidied)
  let typeMap = {};          // location_type_id → chosen 'Location Type*' (sticky)
  let siteMap = {};          // legacy site_id → chosen 'Site*' (sticky)
  // 'crop|id' / 'type|id' the tool filled in by itself from location names —
  // used so the export isn't blocked, but counted until you confirm them.
  let guessed = new Set();
  let presetApplied = false; // the Legacy crop preset is currently in effect

  let locRows = null;        // [[...42]] mapped location rows
  let locSrc = null;         // parallel source rows
  let siteRows = null;       // [[...10]] mapped site rows
  let siteSrc = null;        // parallel source rows

  let locFills = {}, siteFills = {};       // colIdx → { val, mode }
  let locRemoved = new Set(), siteRemoved = new Set();   // row indexes, rebuilt from the keys below
  let locKeys = [], siteKeys = [];         // the Legacy ID behind each row (location_id / site ID)
  let locRemovedK = new Set(), siteRemovedK = new Set(); // rows you dropped, by Legacy ID
  let locEdits = {}, siteEdits = {};       // "legacyId|colIdx" → value — hand edits, kept through rebuilds
  let smartFixMap = {};                    // "grid|colIdx||UPPERVALUE" → template value (sticky)
  let archivedCount = 0;                   // dropped by is_archived
  let alreadyMigratedCount = 0;            // dropped by the 3.0 cross-reference
  let skippedSiteNames = [];               // sites left out as already in 3.0
  let settingsLeft = [];                   // locations on a Legacy settings site (e.g. “Master Employees and Settings”) — never migrated
  let refOnly = true;                      // Sites: only sites the locations use
  let domain = 'locations';                // which preview/bulk pane is showing
  let selCells = new Set(), selAnchor = null, previewOrder = [];
  let completedSites = new Set();          // sites ticked off the "to create" list
  let completedBuild = new Set();          // "kind|value" ticked off the crops / types list
  let pvMode = 'all', pvColsOnly = false;  // preview filter (IMWorkbench.filterBar)
  // Site addresses. Legacy's sites export carries none, and 3.0 needs a full
  // one per site. The most specific wins: one you typed for the site, then the
  // client's address sheet, then the ranch's, then the one for every site.
  let siteAddrAll = '';                    // "street, city, ST zip" for every site
  let siteAddrGroup = {};                  // ranch (name before " - ") → address
  let siteAddrOwn = {};                    // site name → address you typed
  let addrSheet = null;                    // { fileName, bySite: Map<UPPER name, parts>, unmatched:[], rows }
  let initialized = false;

  // ─── Small helpers ───
  function chunkArr(a, n) {
    const out = [];
    for (let i = 0; i < a.length; i += n) out.push(a.slice(i, i + n));
    return out;
  }

  function buildCaseMap(arr) {
    const m = new Map();
    (arr || []).forEach(v => {
      const k = String(v).toUpperCase().trim();
      if (k && !m.has(k)) m.set(k, v);
    });
    return m;
  }

  // 'MM/DD/YYYY' (Legacy) → 'YYYY-MM-DD' (bulk-template date format). Anything
  // that doesn't match a known shape passes through trimmed.
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
    return t;
  }

  // Legacy writes 0 into hectares / length_meters / quantity when the value was
  // never captured, so a 0 is "unknown", not "zero" — leave those blank rather
  // than exporting a misleading 0 acreage.
  function numOrBlank(v, factor) {
    const n = parseFloat(String(v == null ? '' : v).trim());
    if (!Number.isFinite(n) || n === 0) return '';
    const out = factor ? n * factor : n;
    return String(Math.round(out * 10000) / 10000);
  }

  // ─── CSV / workbook parsing ───
  // Minimal RFC-4180 parser. Keeps every value EXACTLY as written so Legacy
  // dates like "12/05/2019" reach normDate() untouched (SheetJS would coerce
  // them to a serial or re-format them first).
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

  function parseAnyFile(file) {
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
          if (headers.length) headers[0] = headers[0].replace(/^﻿/, '').replace(/^#/, '');
          const idx = {};
          headers.forEach((h, i) => { const k = norm(h); if (k && idx[k] == null) idx[k] = i; });
          resolve({ headers, rows: filtered.slice(1), fileName: file.name, idx });
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      if (isCsv) r.readAsText(file); else r.readAsArrayBuffer(file);
    });
  }

  // Read the DROP-DOWN INPUTS sheet into Map<normalizedHeader, values[]>. The
  // sheet's header row names each column ("Site", "Location Type", "Crop &
  // Variety"…), so columns map straight onto the 3.0 schema by name.
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
        // Verbatim — a trailing "ghost space" on a name is meaningful; trimming
        // it makes PickTrace reject the row as "not found".
        const v = String(raw[r][ci] != null ? raw[r][ci] : '');
        if (v.trim() && arr.indexOf(v) < 0) arr.push(v);
      }
      if (arr.length) m.set(k, arr);
    });
    return m;
  }

  function readTemplate(file, kind) {
    // Resolves once the file is read (Load everything places files in turn).
    return new Promise(done => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
        const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
        const ws = wb.Sheets[dataName];
        if (!ws) { alert('Could not find a DATA ENTRY sheet in this template.'); return; }
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const headers = aoa.length ? aoa[0].map(h => String(h == null ? '' : h).trim()) : [];
        while (headers.length && headers[headers.length - 1] === '') headers.pop();
        const noStar = h => norm(h || '').replace(/\*$/, '').trim();
        const dropdowns = readDropdowns(wb);
        if (kind === 'loc') {
          if (noStar(headers[0]) !== 'site' || noStar(headers[1]) !== 'name') {
            alert('This does not look like the 3.0 Locations Bulk Create template (expected Site… and Name… columns).');
            return;
          }
          locTpl = { headers, dropdowns, rawBuffer: buf, fileName: file.name };
          const sites = dropdowns.get('site') || [];
          const crops = dropdowns.get('crop & variety') || [];
          const types = dropdowns.get('location type') || [];
          $('lm-loctpl-name').textContent = file.name;
          $('lm-loctpl-meta').textContent = headers.length + ' columns · ' + sites.length +
            ' site' + (sites.length === 1 ? '' : 's') + ' · ' + types.length + ' location type' +
            (types.length === 1 ? '' : 's') + ' · ' + crops.length + ' crop' + (crops.length === 1 ? '' : 's');
        } else {
          if (noStar(headers[0]) !== 'name') {
            alert('This does not look like the 3.0 Sites Bulk Create template (expected a Name… first column).');
            return;
          }
          siteTpl = { headers, dropdowns, rawBuffer: buf, fileName: file.name };
          const types = dropdowns.get('site type') || [];
          const emps = dropdowns.get('employer') || [];
          $('lm-sitetpl-name').textContent = file.name;
          $('lm-sitetpl-meta').textContent = headers.length + ' columns · ' + types.length +
            ' site type' + (types.length === 1 ? '' : 's') + ' · ' + emps.length +
            ' employer' + (emps.length === 1 ? '' : 's');
        }
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

  // ─── Load handlers ───
  function handleLocFile(file) {
    return parseAnyFile(file).then(d => {
      if (d.idx['site_id'] == null || d.idx['location_id'] == null) {
        alert('Legacy locations file must have "site_id" and "location_id" columns.');
        return;
      }
      locData = d;
      $('lm-loc-name').textContent = file.name;
      const arch = d.rows.filter(r => String(r[d.idx['is_archived']] || '').trim().toLowerCase() === 'true').length;
      $('lm-loc-meta').textContent = d.rows.length + ' rows · ' + arch + ' archived';
      maybeRun();
    }).catch(err => fail('Failed to read locations file', err));
  }

  function handleSitesFile(file) {
    return parseAnyFile(file).then(d => {
      // The Legacy sites export names its columns "ID" and "Site " (note the
      // trailing space) — norm() handles the space; accept "name" too.
      const idI = d.headers.findIndex(h => norm(h) === 'id');
      let nameI = d.headers.findIndex(h => norm(h) === 'site');
      if (nameI < 0) nameI = d.headers.findIndex(h => norm(h) === 'name');
      if (idI < 0 || nameI < 0) {
        alert('Sites file must have "ID" and "Site" (or "Name") columns.');
        return;
      }
      const byId = new Map();
      d.rows.forEach(r => {
        const id = String(r[idI] == null ? '' : r[idI]).trim();
        const nm = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (id) byId.set(id, cleanName(nm));
      });
      sitesData = Object.assign(d, { byId, idI, nameI });
      $('lm-sites-name').textContent = file.name;
      $('lm-sites-meta').textContent = byId.size + ' site' + (byId.size === 1 ? '' : 's');
      maybeRun();
    }).catch(err => fail('Failed to read sites file', err));
  }

  // Optional Legacy lookup: ID → Name, for crops or location types.
  function handleLookupFile(file, kind) {
    return parseAnyFile(file).then(d => {
      const idI = d.headers.findIndex(h => /^(id|.*_id)$/.test(norm(h)));
      const nameI = d.headers.findIndex(h => /^(name|display_name|description|title)$/.test(norm(h)));
      if (idI < 0 || nameI < 0) {
        alert('Lookup file must have an ID column and a Name column.');
        return;
      }
      const m = new Map();
      d.rows.forEach(r => {
        const id = String(r[idI] == null ? '' : r[idI]).trim();
        const nm = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (id) m.set(id, nm);
      });
      if (kind === 'crop') {
        cropLookup = m;
        $('lm-croplk-name').textContent = file.name;
        $('lm-croplk-meta').textContent = m.size + ' crop' + (m.size === 1 ? '' : 's');
      } else {
        typeLookup = m;
        $('lm-typelk-name').textContent = file.name;
        $('lm-typelk-meta').textContent = m.size + ' location type' + (m.size === 1 ? '' : 's');
      }
      maybeRun();
    }).catch(err => fail('Failed to read lookup file', err));
  }

  // Optional: a 3.0 sites export. Any Legacy site whose name already appears
  // here is skipped so the Sites upload doesn't create duplicates.
  function handleMigratedSitesFile(file) {
    return parseAnyFile(file).then(d => {
      let nameI = d.headers.findIndex(h => norm(h).replace(/\*$/, '') === 'name');
      if (nameI < 0) nameI = d.headers.findIndex(h => norm(h) === 'site');
      if (nameI < 0) {
        alert('Already-in-3.0 sites file must have a "Name" or "Site" column.');
        return;
      }
      const s = new Set();
      d.rows.forEach(r => {
        const v = String(r[nameI] == null ? '' : r[nameI]).trim().toUpperCase();
        if (v) s.add(v);
      });
      migratedSites = s;
      $('lm-migrated-name').textContent = file.name;
      $('lm-migrated-meta').textContent = s.size + ' site' + (s.size === 1 ? '' : 's') + ' to skip';
      maybeRun();
    }).catch(err => fail('Failed to read 3.0 sites file', err));
  }

  // ─── Template option access ───
  function tplOptions(tpl, header) {
    if (!tpl || !tpl.dropdowns) return [];
    const h = norm(header).replace(/\*$/, '').trim();
    return tpl.dropdowns.get(h) || [];
  }
  function locOptions(header) { return tplOptions(locTpl, header); }
  function siteOptions(header) { return tplOptions(siteTpl, header); }

  function snapToOption(val, cm) {
    const v = String(val == null ? '' : val).trim();
    if (!v || !cm.size) return v;
    const hit = cm.get(v.toUpperCase());
    return hit != null ? hit : v;
  }

  // ─── ID → template value suggestion ───
  // Legacy exports carry no crop or location-type names, so when no lookup CSV
  // is supplied we infer a suggestion from the text each ID's own locations use
  // (display_name / description / plant_name). The template's Crop & Variety
  // values are "Commodity-Variety" pairs, so we score the COMMODITY prefix and
  // then prefer that group's "-Other" entry.
  //
  // This is a hint, not an answer — the panel always requires confirmation, and
  // low-confidence guesses are labelled as such.
  function optionGroups(options) {
    const groups = new Map();
    (options || []).forEach(o => {
      const g = String(o).split('-')[0].trim();
      if (!g) return;
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(o);
    });
    return groups;
  }

  function suggestFromTexts(texts, options) {
    const groups = optionGroups(options);
    if (!groups.size || !texts.length) return null;
    const blob = texts.map(s => ' ' + String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ');
    let best = null, bestHits = 0, bestLen = 0;
    groups.forEach((list, g) => {
      const gl = g.toLowerCase();
      // Match the group name and its simple singular/plural forms as whole words.
      const forms = new Set([gl, gl.replace(/ies$/, 'y'), gl.replace(/es$/, ''), gl.replace(/s$/, '')]);
      let hits = 0;
      blob.forEach(b => {
        for (const f of forms) {
          if (f.length > 2 && b.indexOf(' ' + f + ' ') >= 0) { hits++; return; }
        }
      });
      if (!hits) return;
      // Longer group names win ties so "Bell Peppers" beats "Peppers".
      if (hits > bestHits || (hits === bestHits && gl.length > bestLen)) {
        bestHits = hits; bestLen = gl.length; best = g;
      }
    });
    if (!best) return null;
    const list = groups.get(best);
    const other = list.find(o => /-other$/i.test(o));
    return { value: other || list[0], confidence: bestHits / blob.length };
  }

  // A lookup name ("broccoli") is a far stronger signal than scraped location
  // text, so try an exact/word match against the option groups first.
  function suggestFromName(name, options) {
    if (!name) return null;
    const cm = buildCaseMap(options);
    const direct = cm.get(String(name).toUpperCase().trim());
    if (direct != null) return { value: direct, confidence: 1 };
    return suggestFromTexts([name], options);
  }

  // The Legacy ID a location row is kept under (removed rows, edits).
  function locKeyOf(row, srcIdx) { return locGet(row, 'location_id') || ('row' + (srcIdx + 2)); }
  function locIdx(key) { return locData && locData.idx[key] != null ? locData.idx[key] : -1; }
  function locGet(row, key) {
    const i = locIdx(key);
    return i < 0 ? '' : String(row[i] == null ? '' : row[i]).trim();
  }

  // Distinct legacy IDs in the surviving rows, with the corroborating text used
  // for suggestions. kind = 'crop' | 'type'.
  function collectIds(kind) {
    const out = new Map(); // id → { count, texts:[], samples:[] }
    if (!locData) return out;
    const key = kind === 'crop' ? 'crop_id' : 'location_type_id';
    locData.rows.forEach(row => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
      const id = locGet(row, key) || '(blank)';
      const e = out.get(id) || { count: 0, texts: [], samples: [] };
      e.count++;
      const dn = locGet(row, 'display_name');
      const de = locGet(row, 'description');
      const pn = locGet(row, 'plant_name');
      const txt = [dn, de, pn].filter(Boolean).join(' | ');
      if (txt) e.texts.push(txt);
      if (e.samples.length < 3 && dn) e.samples.push(dn);
      out.set(id, e);
    });
    return out;
  }

  // ─── Crop & Variety = crop (from crop_id) + variety (from plant_name) ───
  // Legacy keeps the crop as a number (crop_id) and the variety as text
  // (plant_name). 3.0 wants one "Crop-Variety" value, so each part is settled
  // on its own: cropMap names each crop_id ("Blueberries"); each Legacy variety
  // name becomes a 3.0 variety (varietyMap, default: the Legacy name tidied).
  // No variety in Legacy → "Other"; no crop at all → "Other-Other", the usual
  // 3.0 placeholder now that the field is required.
  const cropPart = s => String(s == null ? '' : s).split('-')[0].trim();
  const NO_CROP = new Set(['', '-1', '0', '(blank)']);
  const varietyKey = p => String(p == null ? '' : p).trim().toUpperCase();
  // "BLUE CROP" → "Blue Crop"; short codes ("NB", "MBG") stay as they are.
  function tidyVariety(p) {
    const s = String(p == null ? '' : p).trim().replace(/\s+/g, ' ');
    if (!s) return 'Other';
    return s.split(' ').map(w => w.length <= 3 && /^[A-Z0-9]+$/.test(w) ? w : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
  }
  function varietyFor(plant) {
    const k = varietyKey(plant);
    return varietyMap[k] != null && String(varietyMap[k]).trim() ? String(varietyMap[k]).trim() : tidyVariety(plant);
  }
  function composeCv(row) {
    const cid = locGet(row, 'crop_id');
    if (NO_CROP.has(cid)) return 'Other-Other';
    const crop = cropPart(cropMap[cid] || '');
    return crop ? crop + '-' + varietyFor(locGet(row, 'plant_name')) : '';
  }
  // Crop names the template already has (the part before "-"), plus Other.
  function cropNameOptions() {
    const seen = new Map();
    locOptions('Crop & Variety*').forEach(o => { const c = cropPart(o); if (c && !seen.has(c.toUpperCase())) seen.set(c.toUpperCase(), c); });
    if (!seen.has('OTHER')) seen.set('OTHER', 'Other');
    return [...seen.values()].sort((a, b) => a.localeCompare(b));
  }
  // Look-alike Legacy variety names: the same words once harvest words
  // ("MACHINE", "HAND") are set aside — BLUE CROP / BLUE CROP MACHINE.
  const varietyBase = p => varietyKey(p).replace(/\b(MACHINE|MACH|MECH|MECHANICAL|HAND)\b/g, ' ').replace(/\s+/g, ' ').trim();
  // → [{ base, members: [{ key, plant, count, crops:Set }] }], biggest first, among rows still exported.
  function varietyGroups() {
    const by = new Map();
    if (!locRows) return [];
    locRows.forEach((r, ri) => {
      if (locRemoved.has(ri)) return;
      const cid = locGet(locSrc[ri], 'crop_id');
      if (NO_CROP.has(cid)) return;
      const plant = locGet(locSrc[ri], 'plant_name'), key = varietyKey(plant), base = varietyBase(plant) || key;
      const g = by.get(base) || by.set(base, { base, members: new Map() }).get(base);
      const m = g.members.get(key) || g.members.set(key, { key, plant, count: 0, crops: new Set() }).get(key);
      m.count++; m.crops.add(cid);
    });
    return [...by.values()].map(g => ({ base: g.base, members: [...g.members.values()].sort((a, b) => b.count - a.count) }))
      .sort((a, b) => b.members.length - a.members.length || b.members.reduce((n, m) => n + m.count, 0) - a.members.reduce((n, m) => n + m.count, 0));
  }
  function renderVarieties() {
    const sec = $('lm-section-crop');
    if (!sec) return;
    let box = $('lm-variety-box');
    if (!box) { box = document.createElement('div'); box.id = 'lm-variety-box'; sec.appendChild(box); }
    const groups = varietyGroups();
    if (!groups.length) { box.innerHTML = ''; return; }
    let html = '<h4 style="margin:18px 0 4px;">Varieties</h4><p class="cmp-sites-hint">Each Legacy variety name (<code>plant_name</code>) and the 3.0 variety it becomes — ' +
      'the Crop &amp; Variety value is <b>Crop-Variety</b>. Look-alike names are grouped: give them the same name to combine them. ' +
      'Blocks with no variety get <b>Other</b>.</p><div class="table-wrap"><table class="data-table"><thead><tr><th>Legacy variety</th><th>Blocks</th><th>3.0 variety</th></tr></thead><tbody>';
    groups.forEach(g => {
      if (g.members.length > 1) html += '<tr><td colspan="3" style="background:var(--bg-sunken,#f3f4f6);font-size:12px;color:var(--text-2);">Look alike: ' +
        g.members.map(m => escHtml(m.plant || '(none)')).join(' · ') + '</td></tr>';
      g.members.forEach(m => {
        const cur = varietyMap[m.key] != null ? varietyMap[m.key] : '';
        html += '<tr><td>' + escHtml(m.plant || '(no variety in Legacy)') + '</td><td>' + m.count + '</td><td>' +
          '<input type="text" class="input-field lm-var-in" data-key="' + escHtml(m.key) + '" value="' + escHtml(cur) + '" placeholder="' + escHtml(tidyVariety(m.plant)) + '" style="min-width:200px;"></td></tr>';
      });
    });
    box.innerHTML = html + '</tbody></table></div>';
  }

  function suggestionFor(kind, id, info) {
    const opts = kind === 'crop' ? locOptions('Crop & Variety*') : locOptions('Location Type*');
    if (!opts.length) return null;
    const lk = kind === 'crop' ? cropLookup : typeLookup;
    if (lk && lk.has(id)) {
      const s = suggestFromName(lk.get(id), opts);
      if (s) return { value: kind === 'crop' ? cropPart(s.value) : s.value, confidence: s.confidence, via: 'lookup: ' + lk.get(id) };
    }
    const s = suggestFromTexts(info.texts, opts);
    if (s) return { value: kind === 'crop' ? cropPart(s.value) : s.value, confidence: s.confidence, via: 'inferred from location names' };
    return null;
  }

  // ─── Locations transform ───
  // Legacy site IDs marked Virtual = true in the sites export.
  function settingsSiteIds() {
    const out = new Set();
    if (!sitesData) return out;
    const vi = sitesData.idx['virtual'];
    if (vi == null) return out;
    sitesData.rows.forEach(r => { if (String(r[vi] == null ? '' : r[vi]).trim().toLowerCase() === 'true') out.add(String(r[sitesData.idI] == null ? '' : r[sitesData.idI]).trim()); });
    return out;
  }
  function buildLocRowsRaw() {
    if (!locData) return null;
    locSrc = [];
    locKeys = [];
    archivedCount = 0;
    settingsLeft = [];
    const out = [];
    const siteCase = new Map(locOptions('Site*').map(o => [cleanName(o).toUpperCase(), o]));
    const settings = settingsSiteIds();
    locData.rows.forEach((row, srcIdx) => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') { archivedCount++; return; }
      const sid = locGet(row, 'site_id');
      // A Legacy settings site (Virtual = true, e.g. “Master Employees and
      // Settings”) holds configuration, not fields — its blocks don't move to 3.0.
      if (settings.has(sid)) { settingsLeft.push(locGet(row, 'display_name') + ' (' + sitesData.byId.get(sid) + ')'); return; }
      const r = new Array(LOC_HEADERS.length).fill('');
      // An explicit Site mapping wins; otherwise fall back to the name the
      // sites export gives for this site_id.
      const siteName = cleanName(siteMap[sid] ||
        (sitesData && sitesData.byId.has(sid) ? sitesData.byId.get(sid) : ''));
      // An existing 3.0 site keeps the template's exact spelling (spaces and all).
      r[L_SITE] = siteCase.get(siteName.toUpperCase()) || siteName;
      r[L_NAME] = locGet(row, 'display_name');
      r[L_ALT] = locGet(row, 'location_id');
      r[L_TYPE] = typeMap[locGet(row, 'location_type_id') || '(blank)'] || '';
      r[L_CROP] = composeCv(row);   // crop (crop_id) + variety (plant_name)
      r[6] = numOrBlank(locGet(row, 'hectares'), HECTARES_TO_ACRES);   // Acreage
      r[7] = numOrBlank(locGet(row, 'length_meters'), METERS_TO_FEET); // Length
      r[8] = numOrBlank(locGet(row, 'quantity'));                      // Plant Count
      r[9] = todayIso();                                               // Start Date* — 3.0 wants the upload day
      r[26] = normDate(locGet(row, 'planted_date'));                   // Planted Date
      // Legacy packs a cost code into info as "CC2=<code>" — keep the code,
      // drop the constant key so Custom Data 1 holds the code alone.
      r[38] = locGet(row, 'info').replace(/^CC\d*\s*=\s*/i, '');
      r[39] = locGet(row, 'description');
      r[40] = locGet(row, 'plant_name');
      out.push(r);
      locSrc.push(row);
      locKeys.push(locKeyOf(row, srcIdx));
    });

    // Snap every dropdown-backed column to the template's literal where it
    // matches case-insensitively; unmatched values stay visible and fixable.
    for (let ci = 0; ci < LOC_HEADERS.length; ci++) {
      const opts = locOptions(LOC_HEADERS[ci]);
      if (!opts.length) continue;
      const cm = buildCaseMap(opts);
      out.forEach(r => { if (r[ci]) r[ci] = snapToOption(r[ci], cm); });
    }
    return out;
  }

  // ─── Sites transform ───
  function referencedSiteIds() {
    const s = new Set();
    if (!locData) return s;
    locData.rows.forEach((row, i) => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
      // A block left out (by you, or the client's Remove) doesn't keep its
      // site in: a site whose blocks are all left out isn't moved either.
      if (locRemovedK.has(locKeyOf(row, i))) return;
      const id = locGet(row, 'site_id');
      if (id) s.add(id);
    });
    return s;
  }

  function buildSiteRowsRaw() {
    if (!sitesData) return null;
    siteSrc = [];
    siteKeys = [];
    skippedSiteNames = [];
    alreadyMigratedCount = 0;
    const refs = refOnly ? referencedSiteIds() : null;
    const out = [];
    const idI = sitesData.idI, nameI = sitesData.nameI;
    const a1I = sitesData.idx['address1'] != null ? sitesData.idx['address1'] : sitesData.idx['address'];
    const a2I = sitesData.idx['address2'];
    const cityI = sitesData.idx['city'];
    const stI = sitesData.idx['state'];
    const zipI = sitesData.idx['zip'] != null ? sitesData.idx['zip'] : sitesData.idx['postal code'];
    const ctryI = sitesData.idx['country'];
    const infoI = sitesData.idx['info'];
    const virtI = sitesData.idx['virtual'];
    const g = (row, i) => (i == null || i < 0) ? '' : String(row[i] == null ? '' : row[i]).trim();

    sitesData.rows.forEach(row => {
      const id = g(row, idI);
      const name = cleanName(g(row, nameI));
      if (!id || !name) return;
      // Legacy seeds every tenant with sentinel rows (ID -1 "Global", -2
      // "Error") that are not real sites.
      if (+id < 0) return;
      if (refs && !refs.has(id)) return;
      // Legacy "virtual" sites are config containers (the Master Section rows
      // that hold Employees / Jobs / Pay Groups / Pack Styles), not places, so
      // they are never migrated as 3.0 Sites. This condition used to read
      // `=== 'true' && !refs`, which had it exactly backwards and let the
      // containers through whenever referenced-only was on.
      if (String(g(row, virtI)).toLowerCase() === 'true') return;
      if (migratedSites && migratedSites.has(name.toUpperCase())) { alreadyMigratedCount++; skippedSiteNames.push(name); return; }
      const r = new Array(SITE_HEADERS.length).fill('');
      r[S_NAME] = name;
      // Prefer the CC1 cost code as Alt ID (it is the tenant's own identifier);
      // fall back to the Legacy row ID so Alt ID is never blank.
      const cc = g(row, infoI).replace(/^CC\d*\s*=\s*/i, '');
      r[S_ALT] = (cc && cc !== '-1') ? cc : id;
      r[4] = g(row, a1I);
      r[5] = g(row, a2I);
      r[6] = g(row, cityI);
      r[7] = g(row, stI);
      r[8] = g(row, zipI);
      r[9] = g(row, ctryI);
      out.push(r);
      siteSrc.push(row);
      siteKeys.push(id);
    });

    for (let ci = 0; ci < SITE_HEADERS.length; ci++) {
      const opts = siteOptions(SITE_HEADERS[ci]);
      if (!opts.length) continue;
      const cm = buildCaseMap(opts);
      out.forEach(r => { if (r[ci]) r[ci] = snapToOption(r[ci], cm); });
    }
    return out;
  }

  // Every cell a step below changes gets a reason, for the preview's
  // was → now hover: locWhy / siteWhy["ri|ci"] = why. locRaw / siteRaw hold
  // the rows as mapped from Legacy, before any of it.
  let locRaw = null, siteRaw = null, locWhy = {}, siteWhy = {};
  const whyOf = grid => grid === 'site' ? siteWhy : locWhy;
  function applyFills(rows, fills, removed, grid) {
    if (!rows) return;
    const why = whyOf(grid);
    Object.keys(fills).forEach(k => {
      const ci = +k, { val, mode } = fills[k];
      rows.forEach((r, i) => {
        if (removed.has(i)) return;
        if (mode === 'blank' && r[ci]) return;
        if (r[ci] !== val) why[i + '|' + ci] = val ? 'Column fill you applied' : 'Column cleared';
        r[ci] = val;
      });
    });
  }

  // Raw rows, then (in this order) Smart Fixes, column fills, site addresses
  // and the Sites template's only-option values, the Start Date fix, and your
  // hand edits. Removed rows and edits are keyed by the Legacy ID, so a
  // rebuild never moves them onto another row.
  function rebuild() {
    locRows = buildLocRowsRaw();
    siteRows = buildSiteRowsRaw();
    locRaw = locRows ? locRows.map(r => r.slice()) : null;
    siteRaw = siteRows ? siteRows.map(r => r.slice()) : null;
    locWhy = {}; siteWhy = {};
    locRemoved = keyIdx(locKeys, locRemovedK);
    siteRemoved = keyIdx(siteKeys, siteRemovedK);
    applySmartFixes(locRows, 'loc');
    applySmartFixes(siteRows, 'site');
    applyFills(locRows, locFills, locRemoved, 'loc');
    applyFills(siteRows, siteFills, siteRemoved, 'site');
    applySiteDetails();
    applyEdits(locRows, locKeys, locEdits, 'loc');
    applyEdits(siteRows, siteKeys, siteEdits, 'site');
  }
  function keyIdx(keys, set) {
    const out = new Set();
    keys.forEach((k, i) => { if (set.has(k)) out.add(i); });
    return out;
  }
  function applyEdits(rows, keys, edits, grid) {
    if (!rows) return;
    const why = whyOf(grid);
    rows.forEach((r, ri) => {
      for (let ci = 0; ci < r.length; ci++) {
        const e = edits[keys[ri] + '|' + ci];
        if (e != null) { r[ci] = e; why[ri + '|' + ci] = 'Your edit'; }
      }
    });
  }
  function applySmartFixes(rows, grid) {
    if (!rows || !Object.keys(smartFixMap).length) return;
    const why = whyOf(grid);
    rows.forEach((r, ri) => {
      for (let ci = 0; ci < r.length; ci++) {
        if (!r[ci]) continue;
        const hit = smartFixMap[grid + '|' + ci + '||' + String(r[ci]).toUpperCase().trim()];
        if (hit != null) { r[ci] = hit; why[ri + '|' + ci] = 'Smart Fix you applied'; }
      }
    });
  }

  function ready() {
    return !!(locData && sitesData && locTpl);
  }

  // A crash while building used to surface as "Failed to read <the file you
  // just loaded>", because each file handler's catch swallowed it. Catch it
  // here instead, record the stack, and say it was the build.
  function runMigrate() {
    if (!ready()) {
      alert('Upload the Legacy locations CSV, the sites export, and the 3.0 Locations template first.');
      return;
    }
    try { runMigrateInner(); }
    catch (err) {
      fail('Build Preview failed', err,
        '\n\nUse Debug Dump to save what happened, and send it over.');
    }
    if (window.IMDebug) IMDebug.refresh('locations-migrate');
  }

  // Your choices (mappings, fills, addresses, edits, removed rows) stay
  // through a re-run — loading another file shouldn't undo them. Reset clears them.
  function runMigrateInner() {
    selCells = new Set(); selAnchor = null;

    // Pre-fill the two required resolved columns so a build is usable
    // immediately rather than emitting a file with 220 blank required cells.
    const fp = presetMatches();
    if (fp && fp.ok) applyCropPreset();
    autoFillCrops();
    autoFillTypes();

    rebuild();
    renderAll();
    $('lm-empty').style.display = 'none';
  }

  function maybeRun() {
    const ok = ready();
    if (ok) runMigrate();
    else if (window.IMDebug) IMDebug.refresh('locations-migrate');
    if (window.IMLegacyPage) IMLegacyPage.refresh();
  }

  // Fill any crop_id the preset didn't cover, using the lookup CSV (if given)
  // or the inferred suggestion. Never overwrites an existing choice.
  function autoFillCrops() {
    collectIds('crop').forEach((info, cid) => {
      if (cropMap[cid]) return;
      const s = suggestionFor('crop', cid, info);
      if (!s) return;
      cropMap[cid] = s.value;
      if (!/^lookup/.test(s.via)) guessed.add('crop|' + cid);
    });
  }

  // Legacy leaves location_type_id at -1 on virtually every row, so there is
  // nothing to infer per-ID. Fall back to a single sensible default from the
  // template so Location Type* is never empty; the panel and the Bulk Edit
  // column both override it.
  const TYPE_DEFAULT_ORDER = ['Block', 'Field', 'Section', 'Ranch', 'Facility'];
  function autoFillTypes() {
    const opts = locOptions('Location Type*');
    if (!opts.length) return;
    const cm = buildCaseMap(opts);
    let fallback = null;
    for (const cand of TYPE_DEFAULT_ORDER) {
      const hit = cm.get(cand.toUpperCase());
      if (hit != null) { fallback = hit; break; }
    }
    if (fallback == null) fallback = opts[0];
    collectIds('type').forEach((info, id) => {
      if (typeMap[id]) return;
      // A guess read off a few location names (22 of 369 saying "Field")
      // isn't worth more than the default; a lookup file's name is.
      const s = suggestionFor('type', id, info);
      const strong = s && (/^lookup/.test(s.via) || s.confidence >= 0.5);
      typeMap[id] = strong ? s.value : fallback;
      if (!(s && /^lookup/.test(s.via))) guessed.add('type|' + id);
    });
  }

  function renderAll() {
    renderPresetBanner();
    renderSiteMap();
    renderIdMap('type');
    renderIdMap('crop');
    renderVarieties();
    renderSitesToCreate();
    renderOtherToCreate();
    renderSmartFixes();
    renderSiteFile();
    renderSiteAddr();
    renderCollisions();
    renderLeftOut();
    renderBulkEdit();
    renderPreview();
    updateSummary();
  }
  function rebuildAndRender() { rebuild(); renderAll(); }

  // ─── Legacy crop preset banner ───
  function renderPresetBanner() {
    const el = $('lm-preset-banner');
    if (!el) return;
    const fp = presetMatches();
    if (!fp) { el.style.display = 'none'; return; }
    el.style.display = '';
    if (fp.ok) {
      const mapped = [...collectIds('crop').keys()].filter(c => cropMap[c]).length;
      const total = collectIds('crop').size;
      el.style.background = '#ecfdf5'; el.style.borderColor = '#6ee7b7';
      el.innerHTML =
        '<b style="color:#15803d;">&#10003; ' + escHtml(CROP_PRESET.label) + '</b> &mdash; ' +
        'this file matches the tenant fingerprint (' + fp.probes + '/' + fp.probesPossible +
        ' content checks, all site_ids and crop_ids known), so <b>' + mapped + ' of ' + total +
        '</b> crop IDs were filled automatically. Every one is still editable below.' +
        (presetApplied ? ' <button class="btn btn-ghost btn-sm" id="lm-preset-clear">Clear preset</button>' : '');
      const cb = $('lm-preset-clear');
      if (cb) cb.onclick = () => {
        cropMap = {}; presetApplied = false;
        rebuild(); renderAll();
      };
    } else {
      el.style.background = '#fffbeb'; el.style.borderColor = '#fcd34d';
      const why = [];
      if (!fp.siteOk) why.push('unrecognised site_id' + (fp.unknownSites.length === 1 ? '' : 's') +
        ' ' + fp.unknownSites.slice(0, 6).join(', '));
      if (!fp.cropOk) why.push('unrecognised crop_id' + (fp.unknownCrops.length === 1 ? '' : 's') +
        ' ' + fp.unknownCrops.slice(0, 6).join(', '));
      if (fp.probes < CROP_PRESET.minProbes) why.push('only ' + fp.probes + ' of ' +
        fp.probesPossible + ' content checks passed');
      el.innerHTML =
        '<b style="color:#b45309;">Legacy crop preset not applied.</b> This export does not match ' +
        'the tenant the preset was recovered from (' + escHtml(why.join('; ')) + '), and crop IDs ' +
        'mean different crops in different orgs &mdash; applying it here would mis-map your data. ' +
        'Map the crop IDs below, or upload a crops lookup CSV. ' +
        '<button class="btn btn-ghost btn-sm" id="lm-preset-force">Apply preset anyway</button>';
      const fb = $('lm-preset-force');
      if (fb) fb.onclick = () => {
        if (!window.confirm('The preset was recovered from a different tenant\'s export. ' +
          'Applying it here will very likely assign the wrong crop to every location.\n\nApply anyway?')) return;
        applyCropPreset(); rebuild(); renderAll();
      };
    }
  }

  // ─── Site mapping panel ───
  // Legacy site_id → the 3.0 Site*. Pre-filled from the sites export when that
  // name is already a value in the template's Site dropdown; otherwise the name
  // is shown as-is and needs an explicit pick (or gets created in 3.0).
  function collectSiteIds() {
    const out = new Map(); // site_id → { count, legacyName, samples:[] }
    if (!locData) return out;
    locData.rows.forEach((row, srcIdx) => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
      if (locRemovedK.has(locKeyOf(row, srcIdx))) return;   // left out — nothing to map
      const id = locGet(row, 'site_id') || '(blank)';
      const e = out.get(id) ||
        { count: 0, legacyName: (sitesData && sitesData.byId.get(id)) || '', samples: [] };
      e.count++;
      const dn = locGet(row, 'display_name');
      if (e.samples.length < 3 && dn) e.samples.push(dn);
      out.set(id, e);
    });
    return out;
  }

  function renderSiteMap() {
    const sec = $('lm-section-sitemap');
    const tbl = $('lm-sitemap-table');
    if (!sec || !tbl) return;
    const ids = collectSiteIds();
    if (!ids.size || !locTpl) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    // Pick from the template's Sites plus the new ones the Sites file creates.
    const opts = siteChoices();
    const unresolved = siteMapProblems(ids);
    const title = $('lm-sitemap-title');
    if (title) title.textContent = 'Site Mapping (' + ids.size + ' legacy site' +
      (ids.size === 1 ? '' : 's') +
      (unresolved ? ' · ' + unresolved + ' to resolve' : ' · all resolved') + ')';

    let html = '<thead><tr><th>site_id</th><th>Locations</th><th>Legacy name</th>' +
      '<th>Status</th><th>3.0 Site*</th><th>Manual override</th><th></th></tr></thead><tbody>';
    [...ids.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([id, info]) => {
      const cur = siteMap[id] || '';
      const eff = cur || info.legacyName;
      const st = siteStatus(eff);
      const status = !eff
        ? '<span style="color:#dc2626;font-weight:600;">&#9888; not in the sites export</span> <span class="text-muted small">&mdash; usually a site deleted in Legacy. Pick or type its 3.0 Site, or leave its locations out.</span>'
        : st === 'tpl'
          ? '<span style="color:#15803d;font-weight:600;">&#10003; in template</span>'
          : st === 'new'
            ? '<span style="color:#15803d;font-weight:600;">&#10003; new &mdash; in the Sites file</span>'
            : '<span style="color:#b45309;font-weight:600;">&#9888; not in PickTrace &mdash; create it first (Build in PickTrace)</span>';
      const optsHtml = opts.length
        ? '<option value="">&mdash; keep legacy name &mdash;</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (o === cur ? ' selected' : '') +
            '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no 3.0 Sites yet &mdash; type one)</option>';
      html += '<tr>' +
        '<td><b>' + escHtml(id) + '</b></td>' +
        '<td>' + info.count + '</td>' +
        '<td title="' + escHtml(info.legacyName) + '">' +
          escHtml(info.legacyName.length > 44 ? info.legacyName.slice(0, 44) + '…' : (info.legacyName || '—')) +
          '<div class="text-muted small">' + escHtml(info.samples.join(', ')) + '</div></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="lm-site-pick input-field" data-id="' + escHtml(id) + '" style="min-width:200px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="lm-site-override input-field" data-id="' + escHtml(id) + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm lm-site-apply" data-id="' + escHtml(id) + '">Apply</button>' +
          (!eff ? ' <button class="btn btn-ghost btn-sm lm-site-leave" data-id="' + escHtml(id) + '" title="Drop every location on this site_id from the export (restorable under Left out)">Leave out</button>' : '') + '</td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;

    tbl.querySelectorAll('.lm-site-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const tr = e.currentTarget.closest('tr');
        const pick = tr.querySelector('.lm-site-pick').value;
        const ovr = tr.querySelector('.lm-site-override').value.replace(/^\s+/, '');
        const val = ovr || pick;
        if (!val) { alert('Pick a Site or type a manual override first.'); return; }
        siteMap[e.currentTarget.dataset.id] = val;
        rebuildAndRender();
      });
    });
    tbl.querySelectorAll('.lm-site-leave').forEach(btn => {
      btn.addEventListener('click', e => {
        const sid = e.currentTarget.dataset.id;
        locData.rows.forEach((row, i) => { if ((locGet(row, 'site_id') || '(blank)') === sid) locRemovedK.add(locKeyOf(row, i)); });
        logEvent('setting', 'Left out the locations on Legacy site_id ' + sid + '.');
        rebuildAndRender();
      });
    });
    tbl.querySelectorAll('.lm-site-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        if (e.target.value) { siteMap[e.target.dataset.id] = e.target.value; rebuildAndRender(); }
      });
    });
  }

  // ─── ID Mapping panel (crop_id / location_type_id) ───
  function renderIdMap(kind) {
    const sec = $('lm-section-' + kind);
    const tbl = $('lm-' + kind + '-table');
    if (!sec || !tbl) return;
    const ids = collectIds(kind);
    if (!ids.size || !locTpl) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const map = kind === 'crop' ? cropMap : typeMap;
    const opts = kind === 'crop' ? cropNameOptions() : locOptions('Location Type*');
    const label = kind === 'crop' ? 'Crop' : 'Location Type';
    const unmapped = [...ids.keys()].filter(id => !map[id] || guessed.has(kind + '|' + id)).length;
    const title = $('lm-' + kind + '-title');
    if (title) title.textContent = (kind === 'crop' ? 'Crop' : 'Location Type') +
      ' Mapping (' + ids.size + ' legacy ID' + (ids.size === 1 ? '' : 's') +
      (unmapped ? ' · ' + unmapped + ' unmapped' : ' · all mapped') + ')';

    let html = '<thead><tr><th>Legacy ID</th><th>Rows</th><th>Sample locations</th>' +
      '<th>Suggestion</th><th>3.0 ' + label + '</th><th>Manual override</th><th></th></tr></thead><tbody>';
    [...ids.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([id, info]) => {
      const cur = map[id] || '';
      const sug = cur ? null : suggestionFor(kind, id, info);
      const isGuess = cur && guessed.has(kind + '|' + id);
      const sugHtml = isGuess
        ? '<span style="color:#b45309;font-weight:600;">&#9888; guessed: ' + escHtml(cur) + '</span> <span class="text-muted small">— the tool picked it; confirm or change</span>'
        : cur
        ? '<span style="color:#15803d;font-weight:600;">&#10003; mapped</span>'
        : sug
          ? '<span style="color:' + (sug.confidence >= 0.5 ? '#15803d' : '#b45309') + ';">' +
            escHtml(sug.value) + '</span> <span class="text-muted small">(' +
            Math.round(sug.confidence * 100) + '% · ' + escHtml(sug.via) + ')</span>'
          : '<span class="text-muted small">no guess &mdash; pick one</span>';
      const optsHtml = opts.length
        ? '<option value="">&mdash; pick &mdash;</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (o === cur ? ' selected' : '') +
            '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(your template lists none yet &mdash; type it under Manual override)</option>';
      html += '<tr>' +
        '<td><b>' + escHtml(id) + '</b></td>' +
        '<td>' + info.count + '</td>' +
        '<td class="text-muted small">' + escHtml(info.samples.join(', ')) + '</td>' +
        '<td>' + sugHtml + '</td>' +
        '<td><select class="lm-id-pick input-field" data-kind="' + kind + '" data-id="' + escHtml(id) +
          '" style="min-width:220px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="lm-id-override input-field" data-kind="' + kind +
          '" data-id="' + escHtml(id) + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td>' + (sug && !cur
          ? '<button class="btn btn-success btn-sm lm-id-accept" data-kind="' + kind +
            '" data-id="' + escHtml(id) + '" data-val="' + escHtml(sug.value) + '">Accept</button> '
          : isGuess
          ? '<button class="btn btn-success btn-sm lm-id-accept" data-kind="' + kind +
            '" data-id="' + escHtml(id) + '" data-val="' + escHtml(cur) + '">Confirm</button> '
          : '') +
        '<button class="btn btn-primary btn-sm lm-id-apply" data-kind="' + kind +
          '" data-id="' + escHtml(id) + '">Apply</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    wireIdMapHandlers(tbl);

    const acceptAll = $('lm-' + kind + '-accept-all');
    if (acceptAll) acceptAll.onclick = () => {
      let n = 0;
      collectIds(kind).forEach((info, id) => {
        if (map[id]) { if (guessed.delete(kind + '|' + id)) n++; return; }
        const s = suggestionFor(kind, id, info);
        if (s) { map[id] = s.value; n++; }
      });
      if (!n) { alert('Nothing to accept — every ID is already mapped and confirmed.'); return; }
      rebuild(); renderAll();
    };
  }

  function setIdMapping(kind, id, val) {
    if (!val) return;
    (kind === 'crop' ? cropMap : typeMap)[id] = kind === 'crop' ? cropPart(val) : val;
    guessed.delete(kind + '|' + id);
    rebuild();
    renderAll();
  }

  function wireIdMapHandlers(tbl) {
    tbl.querySelectorAll('.lm-id-accept').forEach(btn => {
      btn.addEventListener('click', e => {
        const b = e.currentTarget;
        setIdMapping(b.dataset.kind, b.dataset.id, b.dataset.val);
      });
    });
    tbl.querySelectorAll('.lm-id-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const b = e.currentTarget, tr = b.closest('tr');
        const pick = tr.querySelector('.lm-id-pick').value;
        const ovr = tr.querySelector('.lm-id-override').value.replace(/^\s+/, '');
        const val = ovr || pick;
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        setIdMapping(b.dataset.kind, b.dataset.id, val);
      });
    });
    tbl.querySelectorAll('.lm-id-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        if (e.target.value) setIdMapping(e.target.dataset.kind, e.target.dataset.id, e.target.value);
      });
    });
  }

  // ═══ Sites file: addresses and the other required columns ═══════════════
  // Site rows' address columns, by the key IMAddress.parse returns.
  const ADDR_COLS = [['address1', 4], ['address2', 5], ['city', 6], ['state', 7], ['zip', 8], ['country', 9]];
  const ADDR_CI = new Set(ADDR_COLS.map(x => x[1]));
  const SITE_ADDR_PARTS = ADDR_COLS;   // [part name, column] — for the client's answers
  // What clients type instead of an answer. DROP_WORD across a whole address
  // means "don't move this site"; NO_VALUE in a box means "nothing here".
  const DROP_WORD = /^(delete|deleted|del|remove|removed|leave (it )?out|not needed|no longer used|not used|don'?t (need|use|move)( it)?|archive|archived)\.?$/i;
  const NO_VALUE = /^(n\/?a|none|null|nil|unknown|tbd|-+|—|\?+|x)$/i;
  function parseAddr(text) {
    if (!text) return null;
    return window.IMAddress ? IMAddress.parse(text)
      : { address1: String(text), city: '', state: '', zip: '', country: '' };
  }
  const addrComplete = p => !!(p && p.address1 && p.city && p.state && p.zip);
  // Ranch = the part of a site name before " - ", when several sites share it.
  function ranchMap(names) {
    const known = new Map();
    names.forEach(s => { const i = s.indexOf(' - '); if (i > 0) { const r = s.slice(0, i).trim(); if (!known.has(r.toUpperCase())) known.set(r.toUpperCase(), r); } });
    const m = new Map();
    names.forEach(s => {
      const i = s.indexOf(' - ');
      const head = (i > 0 ? s.slice(0, i) : '').trim().toUpperCase();
      m.set(s, known.get(head) || s.trim());
    });
    return m;
  }
  // The address a site gets, and where it came from (most specific wins).
  function siteAddrFor(name, ranch) {
    const u = String(name || '').toUpperCase().trim();
    if (siteAddrOwn[name]) return { parts: parseAddr(siteAddrOwn[name]), from: 'site', text: siteAddrOwn[name] };
    if (addrSheet && addrSheet.bySite.has(u)) return { parts: addrSheet.bySite.get(u), from: 'sheet', text: '' };
    if (ranch && siteAddrGroup[ranch]) return { parts: parseAddr(siteAddrGroup[ranch]), from: 'ranch', text: siteAddrGroup[ranch] };
    if (siteAddrAll) return { parts: parseAddr(siteAddrAll), from: 'all', text: siteAddrAll };
    return { parts: null, from: '', text: '' };
  }
  const ADDR_WHY = { site: 'The address you typed for this site', sheet: 'From the client’s address sheet',
    ranch: 'The address for its ranch', all: 'The address for every site' };
  function siteNames() {
    return siteRows ? siteRows.map(r => String(r[S_NAME] || '')).filter(Boolean) : [];
  }
  // After fills: addresses, then any column whose template list has exactly
  // one value (e.g. the org's one Employer). Your own address and the client's
  // sheet replace what Legacy had; the ranch / every-site ones only fill blanks.
  function applySiteDetails() {
    if (!siteRows) return;
    const ranches = ranchMap(siteNames());
    siteRows.forEach((r, ri) => {
      if (siteRemoved.has(ri)) return;
      const a = siteAddrFor(r[S_NAME], ranches.get(r[S_NAME]));
      if (!a.parts) return;
      const over = a.from === 'site' || a.from === 'sheet';
      ADDR_COLS.forEach(([k, ci]) => {
        const v = a.parts[k];
        if (!v || (r[ci] && !over) || r[ci] === v) return;
        r[ci] = v;
        siteWhy[ri + '|' + ci] = ADDR_WHY[a.from];
      });
    });
    SITE_HEADERS.forEach((h, ci) => {
      if (ci === S_NAME) return;
      const o = siteOptions(h);
      if (o.length !== 1) return;
      siteRows.forEach((r, ri) => {
        if (r[ci] || siteRemoved.has(ri)) return;
        r[ci] = o[0];
        siteWhy[ri + '|' + ci] = 'The only option in the Sites template';
      });
    });
  }

  // ─── Start Date ───
  const L_START = 9;
  function todayIso() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  // ─── Name collisions ───
  // Locations sharing Site + Name, and sites sharing a Name: PickTrace keeps
  // only one of each, so rename one in the preview (or drop it with ×).
  function collisionGroups() {
    const out = [];
    const group = (rows, removed, keyOf, grid) => {
      const m = new Map();
      (rows || []).forEach((r, ri) => {
        if (removed.has(ri)) return;
        const k = keyOf(r);
        if (!k) return;
        (m.get(k) || m.set(k, []).get(k)).push(ri);
      });
      m.forEach((list) => { if (list.length > 1) out.push({ grid, rows: list }); });
    };
    group(locRows, locRemoved, r => r[L_NAME] ? (String(r[L_SITE] || '') + '|' + r[L_NAME]).toUpperCase().trim() : '', 'loc');
    if (siteTpl) group(siteRows, siteRemoved, r => String(r[S_NAME] || '').toUpperCase().trim(), 'site');
    return out;
  }
  function collisionSets() {
    const loc = new Set(), site = new Set();
    collisionGroups().forEach(g => g.rows.forEach(ri => (g.grid === 'site' ? site : loc).add(ri)));
    return { loc, site };
  }
  function renderCollisions() {
    const sec = $('lm-section-collisions'), tbl = $('lm-collisions-table');
    if (!sec || !tbl) return;
    const groups = collisionGroups();
    if (!groups.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    let n = 0; groups.forEach(g => { n += g.rows.length; });
    $('lm-collisions-title').textContent = 'Name collisions (' + n + ' rows in ' + groups.length + ' group' + (groups.length === 1 ? '' : 's') + ')';
    tbl.innerHTML = '<thead><tr><th>File</th><th>Site</th><th>Name</th><th>Rows</th><th>Legacy IDs</th></tr></thead><tbody>' +
      groups.map(g => {
        const rows = g.grid === 'site' ? siteRows : locRows, keys = g.grid === 'site' ? siteKeys : locKeys, r0 = rows[g.rows[0]];
        return '<tr><td>' + (g.grid === 'site' ? 'Sites' : 'Locations') + '</td>' +
          '<td>' + escHtml(g.grid === 'site' ? '' : r0[L_SITE]) + '</td>' +
          '<td><b>' + escHtml(g.grid === 'site' ? r0[S_NAME] : r0[L_NAME]) + '</b></td>' +
          '<td>' + g.rows.length + '</td><td>' + escHtml(g.rows.map(ri => keys[ri]).join(', ')) + '</td></tr>';
      }).join('') + '</tbody>';
  }

  // ─── Left out ───
  function renderLeftOut() {
    const sec = $('lm-section-left'), body = $('lm-left-body');
    if (!sec || !body) return;
    const parts = [];
    if (archivedCount) parts.push('<li><b>' + archivedCount + '</b> archived location' + (archivedCount === 1 ? '' : 's') + ' (<code>is_archived</code> = true).</li>');
    if (settingsLeft.length) parts.push('<li><b>' + settingsLeft.length + '</b> location' + (settingsLeft.length === 1 ? '' : 's') + ' on a Legacy settings site, which isn’t a farm or ranch: ' + escHtml(settingsLeft.slice(0, 10).join(', ')) + (settingsLeft.length > 10 ? '…' : '') + '</li>');
    if (skippedSiteNames.length) parts.push('<li><b>' + skippedSiteNames.length + '</b> site' + (skippedSiteNames.length === 1 ? '' : 's') +
      ' already in 3.0 (the "Sites already in 3.0" file): ' + escHtml(skippedSiteNames.slice(0, 40).join(', ')) + (skippedSiteNames.length > 40 ? '…' : '') + '</li>');
    if (locRemoved.size) parts.push('<li><b>' + locRemoved.size + '</b> location' + (locRemoved.size === 1 ? '' : 's') + ' you removed (× in the preview).</li>');
    if (siteRemoved.size) parts.push('<li><b>' + siteRemoved.size + '</b> site' + (siteRemoved.size === 1 ? '' : 's') + ' you removed.</li>');
    if (refOnly) parts.push('<li>Sites no location points at, Legacy’s virtual sites (settings containers) and its <code>Global</code> / <code>Error</code> rows aren’t migrated.</li>');
    if (!parts.length) { sec.style.display = 'none'; body.innerHTML = ''; return; }
    sec.style.display = '';
    body.innerHTML = '<ul class="cmp-sites-hint" style="margin:0;padding-left:18px;">' + parts.join('') + '</ul>' +
      ((locRemoved.size || siteRemoved.size) ? '<button type="button" class="btn btn-ghost btn-sm" id="lm-restore-all">Restore the rows you removed</button>' : '');
  }
  function leftOutCount() { return archivedCount + settingsLeft.length + skippedSiteNames.length + locRemoved.size + siteRemoved.size; }

  // ─── Sites file: Site Type, Employer, Country… for every site ───
  function siteDetailNeeds() {
    if (!siteRows || !siteTpl) return 0;
    return siteRows.filter((r, ri) => !siteRemoved.has(ri) &&
      SITE_HEADERS.some((h, ci) => /\*$/.test(h) && !ADDR_CI.has(ci) && !r[ci])).length;
  }
  function siteAddrNeeds() {
    if (!siteRows || !siteTpl) return 0;
    return siteRows.filter((r, ri) => !siteRemoved.has(ri) &&
      ADDR_COLS.some(([, ci]) => /\*$/.test(SITE_HEADERS[ci]) && !r[ci])).length;
  }
  function renderSiteFile() {
    const sec = $('lm-section-sitefile'), body = $('lm-sitefile-body');
    if (!sec || !body) return;
    const n = siteRows ? siteRows.filter((_, i) => !siteRemoved.has(i)).length : 0;
    if (!n) { sec.style.display = 'none'; body.innerHTML = ''; return; }
    sec.style.display = '';
    if (!siteTpl) {
      body.innerHTML = '<p class="cmp-sites-hint">' + n + ' Legacy site' + (n === 1 ? ' is' : 's are') + ' referenced. Add the org’s ' +
        'personalized <b>Sites</b> template and they go out in a Sites file of their own &mdash; Export then downloads it with ' +
        'Locations in one .zip. Upload the Sites file first.</p>' +
        '<label class="btn btn-primary btn-sm" for="lm-sitetpl-file">&#8679; Upload Sites template</label>';
      return;
    }
    let html = '<p class="cmp-sites-hint">Legacy’s sites carry no Site Type or Employer. Set a value for every site here, or change one site in the preview (<b>Sites</b> view). Addresses are under <b>Site addresses</b>.</p>' +
      '<div class="table-wrap"><table class="data-table"><thead><tr><th>Column</th><th>Status</th><th>Every site</th></tr></thead><tbody>';
    SITE_HEADERS.forEach((h, ci) => {
      if (ci === S_NAME || (ADDR_CI.has(ci) && ci !== 9)) return;
      const req = /\*$/.test(h), list = siteOptions(h), cur = siteFills[ci] ? siteFills[ci].val : '';
      const empty = siteRows.filter((r, ri) => !siteRemoved.has(ri) && !r[ci]).length;
      const status = empty
        ? '<span style="color:' + (req ? '#dc2626' : 'var(--text-3)') + ';font-weight:600;">' + (req ? '⚠ ' : '') + empty + ' empty</span>'
        : '<span style="color:#15803d;">✓ ' + (list.length === 1 && !cur ? 'only option — filled' : 'filled') + '</span>';
      const ctl = list.length
        ? '<select class="input-field lm-sf-fill" data-ci="' + ci + '"><option value="">— pick —</option>' +
            list.slice().sort().map(v => '<option' + (v === cur ? ' selected' : '') + '>' + escHtml(v) + '</option>').join('') + '</select>'
        : '<input type="text" class="input-field lm-sf-fill" data-ci="' + ci + '" value="' + escHtml(cur) + '" placeholder="Value for every site">';
      html += '<tr><td><b' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</b></td><td>' + status + '</td>' +
        '<td><div class="ts-req-ctl">' + ctl + (cur ? ' <button type="button" class="btn btn-ghost btn-sm lm-sf-clear" data-ci="' + ci + '">Clear</button>' : '') + '</div></td></tr>';
    });
    body.innerHTML = html + '</tbody></table></div>';
  }

  // ─── Site addresses: every site / per ranch / each site, or the client's sheet ───
  function renderSiteAddr() {
    const sec = $('lm-section-siteaddr'), body = $('lm-siteaddr-body');
    if (!sec || !body) return;
    const names = siteRows ? siteRows.filter((_, i) => !siteRemoved.has(i)).map(r => String(r[S_NAME] || '')).filter(Boolean) : [];
    if (!names.length) { sec.style.display = 'none'; body.innerHTML = ''; return; }
    sec.style.display = '';
    const ranches = ranchMap(names), groups = new Map();
    names.forEach(s => { const g = ranches.get(s); groups.set(g, (groups.get(g) || 0) + 1); });
    const shown = p => {
      if (!p) return '';
      if (addrComplete(p)) return '<span class="ts-addr-ok" title="Address1: ' + escHtml(p.address1) + '">✓ ' + escHtml(p.city + ', ' + p.state + ' ' + p.zip) + '</span>';
      const miss = ['address1', 'city', 'state', 'zip'].filter(k => !p[k]).map(k => k === 'address1' ? 'street' : k);
      return '<span class="ts-addr-bad">⚠ no ' + miss.join(', ') + ' found</span>';
    };
    const field = (scope, key, val, ph) => '<input type="text" class="input-field ts-addr-in lm-addr-in" data-scope="' + scope + '"' +
      (key != null ? ' data-key="' + escHtml(key) + '"' : '') + ' value="' + escHtml(val) + '" placeholder="' + escHtml(ph) + '">';
    let html = '<p class="cmp-sites-hint">Legacy’s sites export has no addresses, and 3.0 needs a full one per site. Give one address for ' +
      '<b>every site</b> (the main office), one per <b>ranch</b>, or one for <b>each site</b> &mdash; typed as <b>street, city, ST zip</b> &mdash; ' +
      'or load the address sheet the client filled in (under More files). The most specific wins: a site’s own, then the sheet, then its ranch’s, then every site’s.' +
      (siteTpl ? '' : ' These go into the Sites file once you add the Sites template.') + '</p>';
    if (addrSheet) {
      html += '<div class="ts-addr-block"><div class="ts-addr-label">Client’s address sheet &mdash; ' + escHtml(addrSheet.fileName) + '</div>' +
        '<p class="cmp-sites-hint">' + addrSheet.bySite.size + ' site' + (addrSheet.bySite.size === 1 ? '' : 's') + ' matched' +
        (addrSheet.unmatched.length ? ' · <b style="color:#b45309;">' + addrSheet.unmatched.length + ' row' + (addrSheet.unmatched.length === 1 ? '' : 's') +
          ' matched no site:</b> ' + escHtml(addrSheet.unmatched.slice(0, 15).join('; ')) : '') + '</p></div>';
    }
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Every site (' + names.length + ')</div>' +
      '<div class="ts-addr-row">' + field('all', null, siteAddrAll, '123 Main St, City, ST 12345') + shown(parseAddr(siteAddrAll)) + '</div></div>';
    if (groups.size < names.length) {
      html += '<div class="ts-addr-block"><div class="ts-addr-label">Per ranch &mdash; the name before “ - ” (' + groups.size + ')</div>' +
        '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Ranch</th><th>Sites</th><th>Address</th></tr></thead><tbody>';
      [...groups.keys()].sort((a, b) => a.localeCompare(b)).forEach(g => {
        const v = siteAddrGroup[g] || '';
        html += '<tr><td>' + escHtml(g) + '</td><td>' + groups.get(g) + '</td><td><div class="ts-addr-row">' +
          field('ranch', g, v, siteAddrAll ? siteAddrAll + ' (every site)' : 'street, city, ST zip') + shown(parseAddr(v)) + '</div></td></tr>';
      });
      html += '</tbody></table></div></div>';
    }
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Each site (' + names.length + ')</div>' +
      '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Site</th><th>Address</th></tr></thead><tbody>';
    names.slice().sort((a, b) => a.localeCompare(b)).forEach(s => {
      const own = siteAddrOwn[s] || '', a = siteAddrFor(s, ranches.get(s));
      const ph = a.from === 'sheet' ? [a.parts.address1, a.parts.city, a.parts.state + ' ' + a.parts.zip].filter(x => x && x.trim()).join(', ') + ' (sheet)'
        : a.text ? a.text + ' (inherited)' : 'street, city, ST zip';
      html += '<tr><td>' + escHtml(s) + '</td><td><div class="ts-addr-row">' + field('site', s, own, ph) + shown(a.parts) + '</div></td></tr>';
    });
    body.innerHTML = html + '</tbody></table></div></div>';
  }

  // The client's filled-in address sheet: any CSV / xlsx with a site-name
  // column and either address parts (Address/Street, City, State, Zip,
  // Country) or one full-address column. Rows match sites by name — exactly,
  // then by the number before " - " ("01 - North Farm" ↔ "01"), then by
  // the name after it. A single row with no site name is taken as the
  // address for every site.
  function handleAddrFile(file) {
    return parseAnyFile(file).then(d => {
      const col = (...ks) => d.headers.findIndex(h => ks.includes(norm(h).replace(/\*$/, '').replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim()));
      const nameI = col('site', 'site name', 'name', 'ranch', 'ranch name', 'location', 'farm');
      const a1I = col('address1', 'address 1', 'street', 'street address', 'address line 1');
      const a2I = col('address2', 'address 2', 'address line 2', 'suite', 'apt');
      const cityI = col('city', 'town');
      const stI = col('state', 'st', 'province');
      const zipI = col('zip', 'zip code', 'zipcode', 'postal code', 'postal');
      const ctryI = col('country');
      const fullI = col('address', 'full address', 'mailing address', 'physical address');
      if (nameI < 0 || (a1I < 0 && fullI < 0)) {
        alert('The address sheet needs a site-name column (Site / Name / Ranch) and either address columns ' +
          '(Address1, City, State, Zip) or one full Address column.\n\nColumns found: ' + d.headers.filter(Boolean).join(', '));
        return;
      }
      const g = (r, i) => i < 0 ? '' : String(r[i] == null ? '' : r[i]).trim();
      const names = siteNames();
      const exact = new Map(names.map(n => [n.toUpperCase().trim(), n]));
      const byNum = new Map(), byTail = new Map();
      names.forEach(n => {
        const i = n.indexOf(' - ');
        if (i > 0) { byNum.set(n.slice(0, i).trim().toUpperCase(), n); byTail.set(n.slice(i + 3).trim().toUpperCase(), n); }
      });
      const bySite = new Map(), unmatched = [];
      let everySite = null;
      d.rows.forEach(r => {
        const raw = g(r, nameI);
        let parts;
        if (a1I >= 0 && g(r, a1I)) {
          parts = { address1: g(r, a1I), address2: g(r, a2I), city: g(r, cityI), state: g(r, stI).toUpperCase(), zip: g(r, zipI), country: g(r, ctryI) };
          if (!parts.country && parts.state) { const p = parseAddr('x, x, ' + parts.state + ' 00000'); parts.country = p ? p.country : ''; }
        } else if (fullI >= 0 && g(r, fullI)) {
          parts = parseAddr(g(r, fullI));
          if (parts && a2I >= 0) parts.address2 = g(r, a2I);
        } else return;     // no address on this row
        if (!raw) { if (!everySite) everySite = parts; return; }
        const u = raw.toUpperCase();
        const num = (raw.match(/^\s*(\d+)\b/) || [])[1];
        const hit = exact.get(u) || byNum.get(u) || (num && byNum.get(num)) ||
          byTail.get(u) || byTail.get(u.replace(/^\s*\d+\s*-\s*/, ''));
        if (hit) bySite.set(hit.toUpperCase().trim(), parts);
        else unmatched.push(raw);
      });
      if (everySite && !siteAddrAll) {
        siteAddrAll = [everySite.address1, everySite.city, (everySite.state + ' ' + everySite.zip).trim()].filter(Boolean).join(', ');
      }
      addrSheet = { fileName: file.name, bySite, unmatched, rows: d.rows.length };
      $('lm-addr-name').textContent = file.name;
      $('lm-addr-meta').textContent = bySite.size + ' site' + (bySite.size === 1 ? '' : 's') + ' matched' +
        (unmatched.length ? ' · ' + unmatched.length + ' unmatched' : '');
      logEvent('note', 'Address sheet: ' + bySite.size + ' matched, ' + unmatched.length + ' unmatched.');
      if (locRows) { rebuildAndRender(); if (wb) wb.open('siteaddr'); }
      else maybeRun();
    }).catch(err => fail('Failed to read the address sheet', err));
  }

  // ═══ Not in dropdown (Smart Fixes) ═══════════════════════════════════════
  // Values a template's dropdown doesn't list, in both files. Site, Location
  // Type and Crop & Variety have panels of their own, so they're left out here.
  const squash = s => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  function suggestOption(val, opts) {
    const q = squash(val);
    if (!q) return '';
    const exact = opts.find(o => squash(o) === q);
    if (exact) return exact;
    const inside = opts.filter(o => { const s = squash(o); return s && (s.indexOf(q) >= 0 || q.indexOf(s) >= 0); });
    return inside.length === 1 ? inside[0] : '';
  }
  const OWN_PANEL = new Set([L_SITE, L_TYPE, L_CROP]);
  function computeSmartFixes() {
    const out = [];
    const scan = (rows, removed, headers, optsFor, grid) => {
      if (!rows) return;
      headers.forEach((h, ci) => {
        if (grid === 'loc' && OWN_PANEL.has(ci)) return;
        if (grid === 'site' && ci === S_NAME) return;
        const opts = optsFor(h);
        if (!opts.length) return;
        const set = new Set(opts.map(o => String(o).toUpperCase().trim()));
        const seen = new Map();
        rows.forEach((r, ri) => {
          if (removed.has(ri) || !r[ci]) return;
          const u = String(r[ci]).toUpperCase().trim();
          if (set.has(u)) return;
          const e = seen.get(u) || { grid, colIdx: ci, header: h, from: r[ci], count: 0, options: opts, suggestion: suggestOption(r[ci], opts) };
          e.count++;
          seen.set(u, e);
        });
        seen.forEach(e => out.push(e));
      });
    };
    scan(locRows, locRemoved, LOC_HEADERS, locOptions, 'loc');
    if (siteTpl) scan(siteRows, siteRemoved, SITE_HEADERS, siteOptions, 'site');
    out.sort((a, b) => (b.suggestion ? 1 : 0) - (a.suggestion ? 1 : 0) || b.count - a.count);
    return out;
  }
  let smartFixOpts = new Map();
  function renderSmartFixes() {
    const sec = $('lm-section-smartfix'), tbl = $('lm-smartfix-table');
    if (!sec || !tbl) return;
    const fixes = computeSmartFixes();
    if (!fixes.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    let rows = 0, need = 0;
    fixes.forEach(f => { rows += f.count; if (!f.suggestion) need++; });
    $('lm-smartfix-title').textContent = 'Not in dropdown (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + rows + ' row' + (rows === 1 ? '' : 's') +
      (need ? ' — ' + need + ' need your choice' : '') + ')';
    smartFixOpts = new Map();
    let html = '<thead><tr><th>File</th><th>Column</th><th>Value</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    fixes.forEach((f, k) => {
      smartFixOpts.set(k, f.options);
      html += '<tr><td>' + (f.grid === 'site' ? 'Sites' : 'Locations') + '</td><td>' + escHtml(f.header.replace(/\*$/, '')) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td>' + IMWorkbench.pickButton(f.suggestion || '', ' data-ci="' + k + '" style="min-width:180px;"', { kind: 'fix', cls: 'lm-smartfix-pick input-field' }) +
          (f.suggestion ? ' <span class="text-muted small">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') + '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm lm-smartfix-apply" data-key="' + escHtml(f.grid + '|' + f.colIdx + '||' + String(f.from).toUpperCase().trim()) +
          '" data-from="' + escHtml(f.from) + '">Apply</button></td></tr>';
    });
    tbl.innerHTML = html + '</tbody>';
    // data-ci on the picker is the fix's row number here, which keys its options.
    IMWorkbench.wirePicks(tbl, k => smartFixOpts.get(k), { kind: 'fix' });
  }
  function applySmartFixRows(buttons) {
    let n = 0;
    buttons.forEach(b => {
      const pick = b.closest('tr').querySelector('.lm-smartfix-pick');
      const to = pick ? String(pick.value || '').trim() : '';
      if (!to) return;
      smartFixMap[b.dataset.key] = to;
      n++;
    });
    return n;
  }

  // ─── Crops / location types to create ───
  // Values in Crop & Variety* / Location Type* the Locations template doesn't
  // list — they have to exist in PickTrace before the upload (an empty list
  // means the org has none yet).
  function otherToCreate() {
    const out = [];
    if (!locRows) return out;
    [[L_CROP, 'Crop & Variety'], [L_TYPE, 'Location Type']].forEach(([ci, label]) => {
      const set = new Set(locOptions(LOC_HEADERS[ci]).map(o => String(o).toUpperCase().trim()));
      const m = new Map();
      locRows.forEach((r, ri) => {
        if (locRemoved.has(ri) || !r[ci]) return;
        const u = String(r[ci]).toUpperCase().trim();
        if (set.has(u)) return;
        const e = m.get(u) || { kind: label, value: r[ci], count: 0 };
        e.count++;
        m.set(u, e);
      });
      m.forEach(e => out.push(e));
    });
    return out;
  }
  function pendingOtherToCreate() { return otherToCreate().filter(e => !completedBuild.has(e.kind + '|' + e.value)); }
  function renderOtherToCreate() {
    const sec = $('lm-section-cv-create'), tbl = $('lm-cv-create-table');
    if (!sec || !tbl) return;
    const list = otherToCreate();
    if (!list.length) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const done = list.filter(e => completedBuild.has(e.kind + '|' + e.value)).length;
    $('lm-cv-create-title').textContent = 'Crops & location types to create in 3.0 (' + (list.length - done) + ' to create' + (done ? ' · ' + done + ' done' : '') + ')';
    tbl.innerHTML = '<thead><tr><th></th><th>Kind</th><th>Value</th><th>Locations</th></tr></thead><tbody>' +
      list.map(e => {
        const k = e.kind + '|' + e.value, on = completedBuild.has(k);
        return '<tr class="lm-cv-row" data-k="' + escHtml(k) + '" data-v="' + escHtml(e.value) + '" style="cursor:pointer;" title="Click to copy and mark it created in 3.0">' +
          '<td style="text-align:center;">' + (on ? '<span style="color:#15803d;font-weight:700;">&#10003;</span>' : '<span style="color:#9ca3af;">&#9744;</span>') + '</td>' +
          '<td>' + escHtml(e.kind) + '</td><td><b' + (on ? ' style="text-decoration:line-through;color:#15803d;"' : '') + '>' + escHtml(e.value) + '</b></td>' +
          '<td>' + e.count + '</td></tr>';
      }).join('') + '</tbody>';
  }

  // ─── Sites to create in 3.0 ───
  // 'tpl' = in the Locations template's Site list (exists in PickTrace);
  // 'new' = created by the Sites file this tool builds; '' = neither — it has
  // to be created in PickTrace before the Locations upload. An empty Site list
  // means the org has no sites yet, so nothing counts as 'tpl'.
  function newSiteSet() {
    const s = new Set();
    if (!siteTpl || !siteRows) return s;
    siteRows.forEach((r, i) => { if (!siteRemoved.has(i) && r[S_NAME]) s.add(String(r[S_NAME]).toUpperCase().trim()); });
    return s;
  }
  function siteStatus(v) {
    if (!v) return '';
    const u = String(v).toUpperCase().trim();
    if (buildCaseMap(locOptions('Site*')).has(u)) return 'tpl';
    return newSiteSet().has(u) ? 'new' : '';
  }
  function siteValid(v) { return !!siteStatus(v); }
  function siteChoices() {
    const out = locOptions('Site*').slice();
    const have = new Set(out.map(o => String(o).toUpperCase().trim()));
    if (siteTpl && siteRows) siteRows.forEach((r, i) => {
      const n = r[S_NAME];
      if (n && !siteRemoved.has(i) && !have.has(String(n).toUpperCase().trim())) { out.push(n); have.add(String(n).toUpperCase().trim()); }
    });
    return out.sort((a, b) => a.localeCompare(b));
  }
  // site_ids whose Site* is blank (not in the sites export, nothing picked).
  function siteMapProblems(ids) {
    return [...(ids || collectSiteIds()).entries()].filter(([id, info]) => !(siteMap[id] || info.legacyName)).length;
  }

  function getSitesToCreate() {
    const m = new Map(); // site name → { count, ids:Set }
    if (!locRows) return m;
    locRows.forEach((r, ri) => {
      if (locRemoved.has(ri)) return;
      const v = r[L_SITE];
      if (!v || siteValid(v)) return;
      const e = m.get(v) || { count: 0, ids: new Set() };
      e.count++;
      e.ids.add(locGet(locSrc[ri], 'site_id'));
      m.set(v, e);
    });
    return m;
  }

  function pendingSitesToCreate() {
    const out = [];
    getSitesToCreate().forEach((info, name) => {
      if (!completedSites.has(name)) out.push([name, info]);
    });
    return out;
  }

  function renderSitesToCreate() {
    const sec = $('lm-section-create');
    const tbl = $('lm-create-table');
    if (!sec || !tbl) return;
    const m = getSitesToCreate();
    if (!m.size) { sec.style.display = 'none'; tbl.innerHTML = ''; return; }
    sec.style.display = '';
    const done = [...m.keys()].filter(n => completedSites.has(n)).length;
    const title = $('lm-create-title');
    if (title) title.textContent = 'Sites to Create in 3.0 (' + (m.size - done) +
      ' to create' + (done ? ' · ' + done + ' done' : '') + ')';
    let html = '<thead><tr><th></th><th>Site to create</th><th>Locations</th>' +
      '<th>Legacy site_id</th></tr></thead><tbody>';
    [...m.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([name, info]) => {
      const isDone = completedSites.has(name);
      const mark = isDone
        ? '<span style="color:#15803d;font-weight:700;">&#10003;</span>'
        : '<span style="color:#9ca3af;">&#9744;</span>';
      const st = isDone ? ' style="text-decoration:line-through;color:#15803d;"' : '';
      html += '<tr class="lm-create-row" data-site="' + escHtml(name) + '" style="cursor:pointer;" ' +
        'title="Click to copy this name and mark it created in 3.0">' +
        '<td style="text-align:center;">' + mark + '</td>' +
        '<td><b' + st + '>' + escHtml(name) + '</b></td>' +
        '<td>' + info.count + '</td>' +
        '<td>' + escHtml([...info.ids].join(', ')) + '</td></tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.lm-create-row').forEach(tr => {
      tr.addEventListener('click', () => {
        const n = tr.dataset.site;
        if (completedSites.has(n)) completedSites.delete(n); else completedSites.add(n);
        if (navigator.clipboard) navigator.clipboard.writeText(n).catch(() => {});
        renderSitesToCreate();
        updateSummary();
      });
    });
  }

  // ─── Domain accessors (which pane the preview / bulk-edit act on) ───
  function curHeaders() { return domain === 'sites' ? SITE_HEADERS : LOC_HEADERS; }
  function curRows() { return domain === 'sites' ? siteRows : locRows; }
  function curFills() { return domain === 'sites' ? siteFills : locFills; }
  function curRemoved() { return domain === 'sites' ? siteRemoved : locRemoved; }
  function curOptions(h) { return domain === 'sites' ? siteOptions(h) : locOptions(h); }

  // ─── Bulk-edit columns ───
  function renderBulkEdit() {
    const sec = $('lm-section-bulk');
    const tbl = $('lm-bulk-table');
    if (!sec || !tbl || !curRows()) return;
    sec.style.display = '';
    const fills = curFills();
    let html = '<thead><tr><th>Column</th><th>Current bulk value</th><th>Pick</th>' +
      '<th>Manual override</th><th></th><th></th></tr></thead><tbody>';
    curHeaders().forEach((h, idx) => {
      const opts = curOptions(h);
      const cur = fills[idx];
      const optsHtml = opts.length
        ? '<option value="">&mdash; pick &mdash;</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (cur && o === cur.val ? ' selected' : '') +
            '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no template dropdown &mdash; use manual override)</option>';
      const status = cur
        ? '<span style="color:#15803d;font-weight:600;">&#10003; ' +
          escHtml(String(cur.val).slice(0, 40) || '(blank)') + '</span> ' +
          '<span class="text-muted small">(' + (cur.mode === 'blank' ? 'blanks only' : 'all rows') + ')</span>'
        : '<span class="text-muted small">&mdash;</span>';
      const req = /\*$/.test(h);
      html += '<tr>' +
        '<td><b' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="lm-bulk-pick input-field" data-idx="' + idx + '" style="min-width:180px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="lm-bulk-override input-field" data-idx="' + idx + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm lm-bulk-apply" data-idx="' + idx + '" title="Overwrite this column for every row">Apply all</button> ' +
        '<button class="btn btn-success btn-sm lm-bulk-fillblank" data-idx="' + idx + '" title="Fill only the empty cells in this column">Fill blanks</button></td>' +
        '<td><button class="btn btn-ghost btn-sm lm-bulk-clear" data-idx="' + idx + '" title="Empty this column for all rows">Clear</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    wireBulkEditHandlers(tbl);
  }

  // Fills are sticky (they re-apply after every rebuild). A whole-column fill
  // replaces earlier hand edits in that column; Clear keeps the column empty.
  function curEdits() { return domain === 'sites' ? siteEdits : locEdits; }
  function dropEditsIn(edits, idx) {
    Object.keys(edits).forEach(k => { if (+k.slice(k.lastIndexOf('|') + 1) === idx) delete edits[k]; });
  }
  function applyColumnFill(idx, val, mode) {
    curFills()[idx] = { val, mode: mode === 'blank' ? 'blank' : 'all' };
    if (mode !== 'blank') dropEditsIn(curEdits(), idx);
    rebuildAndRender();
  }

  function clearColumnFill(idx) {
    curFills()[idx] = { val: '', mode: 'all' };
    dropEditsIn(curEdits(), idx);
    rebuildAndRender();
  }

  function wireBulkEditHandlers(tbl) {
    const grab = e => {
      const tr = e.target.closest('tr');
      const pick = tr.querySelector('.lm-bulk-pick').value;
      const ovr = tr.querySelector('.lm-bulk-override').value.replace(/^\s+/, '');
      return ovr || pick;
    };
    tbl.querySelectorAll('.lm-bulk-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'all');
      });
    });
    tbl.querySelectorAll('.lm-bulk-fillblank').forEach(btn => {
      btn.addEventListener('click', e => {
        const val = grab(e);
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        applyColumnFill(+e.target.dataset.idx, val, 'blank');
      });
    });
    tbl.querySelectorAll('.lm-bulk-clear').forEach(btn => {
      btn.addEventListener('click', e => clearColumnFill(+e.target.dataset.idx));
    });
    tbl.querySelectorAll('.lm-bulk-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        if (e.target.value) applyColumnFill(+e.target.dataset.idx, e.target.value, 'all');
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
    for (let p = r0; p <= r1; p++) for (let c = c0; c <= c1; c++) s.add(ckey(previewOrder[p], c));
    selCells = s;
  }

  function updateSelBar() {
    const bar = $('lm-sel-bar');
    if (!bar) return;
    const n = selCells.size;
    bar.style.display = n ? 'flex' : 'none';
    const c = $('lm-sel-count');
    if (c) c.textContent = n + ' cell' + (n === 1 ? '' : 's') + ' selected';
  }

  function paintSelection() {
    const tbl = $('lm-preview-table');
    if (!tbl) return;
    tbl.querySelectorAll('td.lm-cell').forEach(td => {
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

  function clearSelection() { selCells = new Set(); paintSelection(); }

  function setSelectedCells(val) {
    const rows = curRows();
    if (!rows || !selCells.size) return;
    const keys = domain === 'sites' ? siteKeys : locKeys, edits = curEdits();
    selCells.forEach(k => {
      const i = k.indexOf(':');
      const ri = +k.slice(0, i), ci = +k.slice(i + 1);
      if (rows[ri]) edits[keys[ri] + '|' + ci] = val;
    });
    rebuildAndRender();
  }

  // ─── Preview ───
  // Red: a required (*) cell is empty, or the name clashes. Amber: a value
  // PickTrace doesn't have yet (not in the template's list, a site not in
  // PickTrace or the Sites file, a crop to create) or a Start Date before
  // today. Corner tick: changed after mapping — hover for was → now and why.
  const FLAG_STYLE = { req: 'background:#fee2e2;color:#7f1d1d;', off: 'background:#fffbeb;color:#92400e;' };
  let pvCache = null;   // per render: option sets + collisions
  function pvSets() {
    if (pvCache) return pvCache;
    const opt = {};
    const setFor = (grid, ci) => {
      const k = grid + ci;
      if (!(k in opt)) {
        const o = grid === 'site' ? siteOptions(SITE_HEADERS[ci]) : locOptions(LOC_HEADERS[ci]);
        opt[k] = o.length ? new Set(o.map(v => String(v).toUpperCase().trim())) : null;
      }
      return opt[k];
    };
    pvCache = { setFor, coll: collisionSets() };
    return pvCache;
  }
  // → [flag, tip] for one cell of the grid on show.
  function cellFlag(grid, ri, ci) {
    const rows = grid === 'site' ? siteRows : locRows, H = grid === 'site' ? SITE_HEADERS : LOC_HEADERS;
    const v = rows[ri][ci], c = pvSets();
    const nameCi = grid === 'site' ? S_NAME : L_NAME;
    if (ci === nameCi && c.coll[grid].has(ri)) return ['req', grid === 'site' ? 'Another site has this Name — PickTrace keeps only one'
      : 'Another location at this Site has this Name — PickTrace keeps only one'];
    if (!v) return /\*$/.test(H[ci]) ? ['req', 'Required — empty'] : ['', ''];
    if (grid === 'loc') {
      if (ci === L_SITE) return siteValid(v) ? ['', ''] : ['off', 'Not in PickTrace and not in the Sites file — create it first (Build in PickTrace)'];
      if (ci === L_CROP || ci === L_TYPE) {
        const s = c.setFor('loc', ci);
        if (!(s && s.has(String(v).toUpperCase().trim()))) return ['off', 'Not in the template — create it in PickTrace first (Build in PickTrace)'];
        // Filled in by the tool, not by you: worth a look until confirmed.
        const kind = ci === L_CROP ? 'crop' : 'type';
        const id = locGet(locSrc[ri], kind === 'crop' ? 'crop_id' : 'location_type_id') || '(blank)';
        return guessed.has(kind + '|' + id) && !locWhy[ri + '|' + ci]
          ? ['off', 'Guessed by the tool — confirm it under ' + (kind === 'crop' ? 'Crops' : 'Location types')] : ['', ''];
      }
    }
    const s = c.setFor(grid, ci);
    return s && !s.has(String(v).toUpperCase().trim()) ? ['off', 'Not in the template’s dropdown — see Not in dropdown'] : ['', ''];
  }
  function changeOf(grid, ri, ci) {
    const raw = grid === 'site' ? siteRaw : locRaw, rows = grid === 'site' ? siteRows : locRows;
    if (!raw || !raw[ri]) return '';
    const why = whyOf(grid)[ri + '|' + ci];
    return why ? IMWorkbench.changeNote(raw[ri][ci], rows[ri][ci], why) : '';
  }

  function renderPreview() {
    const rows = curRows();
    const sec = $('lm-section-preview');
    const tbl = $('lm-preview-table');
    if (!sec || !tbl || !rows) return;
    sec.style.display = '';
    pvCache = null;
    const grid = domain === 'sites' ? 'site' : 'loc';
    const headers = curHeaders();
    const removed = curRemoved();
    const keys = grid === 'site' ? siteKeys : locKeys;
    // The Employees pill is the page's (legacy-page.js), not this grid's.
    document.querySelectorAll('.lm-domain-pill:not([data-domain="employees"])').forEach(b => {
      b.classList.toggle('ts-domain-active', b.dataset.domain === domain);
      b.textContent = (b.dataset.domain === 'sites' ? '🏢 Sites' : '📍 Locations') + ' (' +
        ((b.dataset.domain === 'sites' ? siteRows : locRows) || []).filter((_, i) => !(b.dataset.domain === 'sites' ? siteRemoved : locRemoved).has(i)).length + ')';
    });

    const vis = rows.map((_, i) => i).filter(i => !removed.has(i));
    const info = new Map(), counts = { all: vis.length, fix: 0, chg: 0, any: 0 };
    vis.forEach(ri => {
      const f = [], c = [];
      for (let i = 0; i < headers.length; i++) {
        if (cellFlag(grid, ri, i)[0]) f.push(i);
        if (changeOf(grid, ri, i)) c.push(i);
      }
      info.set(ri, { f, c });
      if (f.length) counts.fix++;
      if (c.length) counts.chg++;
      if (f.length || c.length) counts.any++;
    });
    const show = vis.filter(ri => IMWorkbench.rowMatch(pvMode, info.get(ri).f.length, info.get(ri).c.length));
    let cols = headers.map((_, i) => i);
    if (pvColsOnly) {
      const keep = new Set(grid === 'site' ? [S_NAME] : [L_SITE, L_NAME]);
      show.forEach(ri => { const x = info.get(ri); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      cols = cols.filter(i => keep.has(i));
    }
    const visSet = new Set(vis);
    selCells.forEach(k => { if (!visSet.has(+k.slice(0, k.indexOf(':')))) selCells.delete(k); });
    const head = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>' + cols.map(i =>
      '<th' + (/\*$/.test(headers[i]) ? ' style="color:#dc2626;"' : '') + '>' + escHtml(headers[i]) + '</th>').join('') + '</tr></thead>';
    const rowHtml = ri => {
      const row = rows[ri];
      let html = '<tr><td style="text-align:center;padding:0;">' +
        '<button class="lm-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export (Legacy ID ' + escHtml(keys[ri]) + ')" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      cols.forEach(i => {
        const [fl, tip] = cellFlag(grid, ri, i), note = changeOf(grid, ri, i);
        const t = [tip, note].filter(Boolean).join(' · ');
        html += '<td class="lm-cell' + (note ? ' is-chg' : '') + '" contenteditable="true" spellcheck="false" ' +
          'data-ri="' + ri + '" data-ci="' + i + '"' + (fl ? ' style="' + FLAG_STYLE[fl] + '"' : '') +
          (t ? ' title="' + escHtml(t) + '"' : '') + '>' + escHtml(row[i]) + '</td>';
      });
      return html + '</tr>';
    };
    IMWorkbench.rows(tbl, { head: head, ids: show, row: rowHtml, cols: cols.length + 1,
      after: () => { previewOrder = IMWorkbench.rowOrder(tbl); paintSelection(); },
      keys: [null].concat(cols.map(i => grid + ':' + headers[i])), cell: (ri, p) => rows[ri][cols[p - 1]],
      widths: [28].concat(cols.map(i => IMWorkbench.fitWidth(headers[i], show.map(ri => rows[ri][i])))),
      empty: pvMode === 'fix' ? 'No row needs a fix ✓' : 'Nothing here has been changed.' });

    const hint = sec.querySelector('.cmp-sites-hint');
    let ht = IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly, shownCols: cols.length, totalCols: headers.length }) +
      '<span class="wb-filter-note">' + (show.length === vis.length ? 'All ' + vis.length : show.length + ' of ' + vis.length) +
      ' ' + (grid === 'site' ? 'sites' : 'locations') + ' · editable — click a cell, type, then click away (or press Enter). ' +
      '<b>Shift+click</b> selects a range, <b>Ctrl/Cmd+click</b> single cells. Click <b>×</b> to drop a row.</span>';
    if (removed.size) ht += ' <button class="btn btn-ghost btn-sm" id="lm-restore-rows">Restore ' +
      removed.size + ' removed row' + (removed.size === 1 ? '' : 's') + '</button>';
    if (hint) hint.innerHTML = ht;

    // Rows come and go as you scroll, and the bar is redrawn every render —
    // their handlers live on the table / section, once.
    if (!tbl._lmRowsWired) {
      tbl._lmRowsWired = true;
      tbl.addEventListener('click', e => {
        const rm = e.target.closest('.lm-row-remove');
        if (!rm) return;
        e.stopPropagation();
        const ri = +rm.dataset.ri;
        (domain === 'sites' ? siteRemovedK : locRemovedK).add((domain === 'sites' ? siteKeys : locKeys)[ri]);
        rebuildAndRender();
      });
      sec.addEventListener('change', e => {
        if (e.target.classList.contains('wb-rowfilter')) { pvMode = e.target.value; IMWorkbench.rowsTop(tbl); renderPreview(); }
        else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
      });
      sec.addEventListener('click', e => {
        if (!e.target.closest('#lm-restore-rows')) return;
        if (domain === 'sites') siteRemovedK.clear(); else locRemovedK.clear();
        rebuildAndRender();
      });
    }
    paintSelection();
    updateExportButton();
  }

  // ─── Summary ───
  function emptyRequired(rows, headers, removed) {
    if (!rows) return 0;
    const reqIdx = headers.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    let n = 0;
    rows.forEach((r, ri) => {
      if (removed.has(ri)) return;
      reqIdx.forEach(i => { if (!r[i]) n++; });
    });
    return n;
  }

  // Which required columns are incomplete, and by how many rows.
  function requiredGaps(rows, headers, removed) {
    if (!rows || !rows.length) return [];
    const out = [];
    headers.forEach((h, i) => {
      if (!/\*$/.test(h)) return;
      let n = 0;
      rows.forEach((r, ri) => { if (!removed.has(ri) && !r[i]) n++; });
      if (n) out.push({ header: h, idx: i, rows: n });
    });
    return out;
  }

  function updateSummary() {
    const sum = $('lm-summary');
    if (!sum) return;
    if (!locData || !locTpl) { sum.style.display = 'none'; return; }
    const locVis = locRows ? locRows.filter((_, i) => !locRemoved.has(i)).length : 0;
    const siteVis = siteRows ? siteRows.filter((_, i) => !siteRemoved.has(i)).length : 0;
    const locEmpty = emptyRequired(locRows, LOC_HEADERS, locRemoved);
    const siteEmpty = emptyRequired(siteRows, SITE_HEADERS, siteRemoved);
    const unmappedCrop = [...collectIds('crop').keys()].filter(id => !cropMap[id]).length;
    const unmappedType = [...collectIds('type').keys()].filter(id => !typeMap[id]).length;
    const toCreate = pendingSitesToCreate().length;
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + locVis + '</b> locations' +
        (locRemoved.size ? ' <span class="text-muted small">(' + locRemoved.size + ' removed)</span>' : '') + '</div>' +
      '<div class="cmp-stat"><b>' + siteVis + '</b> sites' +
        (refOnly ? ' <span class="text-muted small">(referenced only)</span>' : '') + '</div>' +
      '<div class="cmp-stat"><b>' + archivedCount + '</b> skipped (archived)</div>' +
      (alreadyMigratedCount ? '<div class="cmp-stat"><b>' + alreadyMigratedCount + '</b> skipped (already in 3.0)</div>' : '') +
      (unmappedType ? '<div class="cmp-stat cmp-warn"><b>' + unmappedType + '</b> unmapped location type ID' + (unmappedType === 1 ? '' : 's') + '</div>' : '') +
      (unmappedCrop ? '<div class="cmp-stat cmp-warn"><b>' + unmappedCrop + '</b> unmapped crop ID' + (unmappedCrop === 1 ? '' : 's') + '</div>' : '') +
      (toCreate ? '<div class="cmp-stat cmp-warn"><b>' + toCreate + '</b> site' + (toCreate === 1 ? '' : 's') + ' to create in 3.0</div>' : '') +
      (locEmpty ? '<div class="cmp-stat cmp-warn"><b>' + locEmpty + '</b> empty required cells (locations)</div>' : '') +
      (siteEmpty ? '<div class="cmp-stat cmp-warn"><b>' + siteEmpty + '</b> empty required cells (sites)</div>' : '');
    if (wb) wb.refresh();
  }

  // ─── The one Legacy page (legacy-page.js) ───
  // This engine's share of the page: its section counts, what opening one of
  // its sections does to the preview, its files for Export all, and how
  // "Upload all files" recognises its files. `wb` stands in for the shared
  // workbench, so the rest of this module calls wb.refresh() / wb.open() as before.
  const wb = {
    refresh: () => { if (window.IMLegacyPage) IMLegacyPage.refresh(); },
    open: k => { if (window.IMLegacyPage) IMLegacyPage.open(k); },
    reset: () => {}
  };
  const reqCols = (rows, H, removed) => rows ? H.filter((h, i) => /\*$/.test(h) && rows.some((r, ri) => !removed.has(ri) && !r[i])).length : 0;
  const COUNTS = {
    sitemap: () => siteMapProblems(),
    crops: () => [...collectIds('crop').keys()].filter(id => !NO_CROP.has(id) && (!cropMap[id] || guessed.has('crop|' + id))).length,
    types: () => [...collectIds('type').keys()].filter(id => !typeMap[id] || guessed.has('type|' + id)).length,
    sitefile: () => siteTpl ? siteDetailNeeds() : (siteRows ? siteRows.length : 0),
    siteaddr: () => siteAddrNeeds(),
    required: () => reqCols(locRows, LOC_HEADERS, locRemoved) + (siteTpl ? reqCols(siteRows, SITE_HEADERS, siteRemoved) : 0),
    dropdown: () => computeSmartFixes().length,
    build: () => pendingSitesToCreate().length + pendingOtherToCreate().length,
    collisions: () => { let n = 0; collisionGroups().forEach(g => { n += g.rows.length; }); return n; },
    left: () => leftOutCount()
  };
  function registerPage() {
    if (!window.IMLegacyPage) return;
    IMLegacyPage.register('lm', {
      ready: () => ready(),
      built: () => !!(ready() && locRows),
      rows: () => locRows ? locRows.filter((_, i) => !locRemoved.has(i)).length : 0,
      domain: () => domain,
      count: key => COUNTS[key] ? COUNTS[key]() : 0,
      // A Sites panel shows the Sites file in the preview; a Locations one, Locations.
      onOpen: key => {
        const v = (key === 'sitefile' || key === 'siteaddr') && siteRows && siteRows.length ? 'sites'
          : (key === 'collisions' || key === 'left') ? domain : 'locations';
        if (v !== domain) { domain = v; clearSelection(); IMWorkbench.rowsTop($('lm-preview-table')); renderBulkEdit(); renderPreview(); }
      },
      run: () => runMigrate(),
      reset: () => reset(),
      exportParts: limit => exportParts(limit),
      questions: () => questions(),
      applyAnswers: a => applyAnswers(a),
      debugReady: () => !!(locData || sitesData || locTpl || siteTpl || cropLookup || addrSheet || typeLookup || migratedSites || events.length),
      log: logEvent,
      fail: fail,
      drop: {
        kinds: { loc: 'Legacy locations export', sites: 'Legacy sites export', loctpl: '3.0 Locations template',
          sitetpl: '3.0 Sites template', migrated: 'Sites already in 3.0', addr: 'Site addresses from the client',
          croplk: 'Crops lookup (ID → name)', typelk: 'Location types lookup (ID → name)' },
        order: ['loctpl', 'sitetpl', 'croplk', 'typelk', 'migrated', 'sites', 'addr', 'loc'],
        classify: i => {
          const H = (...k) => IMLegacyDrop.has(i, ...k), A = (...k) => IMLegacyDrop.hasAny(i, ...k);
          if (i.isTemplate) {
            if (i.norm[0] === 'site' && i.norm[1] === 'name') return 'loctpl';
            if (i.norm[0] === 'name' && A('site type')) return 'sitetpl';
            return null;
          }
          if (H('site_id', 'location_id')) return 'loc';
          if (H('id') && A('site', 'name') && A('virtual', 'created', 'last modified') && !A('job')) return 'sites';
          if (A('site type') && A('name') && !A('first name')) return 'migrated';       // a 3.0 sites export
          if (A('site', 'site name', 'name', 'ranch', 'ranch name', 'location', 'farm') &&
              A('address', 'address1', 'address 1', 'street', 'street address') && !A('first name', 'contractor id')) return 'addr';
          // An ID → Name list is a crops or a location-types lookup (or a
          // contractor list) — the headers can't say which, so it's asked.
          if (i.norm.filter(Boolean).length <= 4 && A('id') && A('name', 'display_name', 'description', 'title') && !i.guess) i.guess = 'croplk';
          return null;
        },
        place: (key, file) => ({
          loc: handleLocFile, sites: handleSitesFile, migrated: handleMigratedSitesFile, addr: handleAddrFile,
          loctpl: f => readTemplate(f, 'loc'), sitetpl: f => readTemplate(f, 'site'),
          croplk: f => handleLookupFile(f, 'crop'), typelk: f => handleLookupFile(f, 'type')
        })[key](file)
      }
    });
  }

  function updateExportButton() {
    if (window.IMLegacyPage) IMLegacyPage.refresh();
    if (window.IMDebug) IMDebug.refresh('locations-migrate');
  }

  // ─── Debug dump ───
  // This module emits TWO grids from one source, so `output` covers Locations
  // (the primary) and `sitesOutput` covers the Sites companion file.
  // The Legacy columns this module reads ("a|b" = either name). A missing one
  // never stops the build; it just leaves its 3.0 column blank, so the dump
  // lists each one and whether the uploaded file has it.
  const LOC_KEYS = ['site_id', 'location_id', 'is_archived', 'display_name', 'location_type_id',
    'crop_id', 'hectares', 'length_meters', 'quantity', 'planted_date', 'info', 'description', 'plant_name'];
  const SITE_KEYS = ['id', 'site|name', 'address1|address', 'address2', 'city', 'state',
    'zip|postal code', 'country', 'info', 'virtual'];
  function columnCheck(data, keys) {
    if (!data || !data.idx) return null;
    return keys.map(k => {
      const hit = k.split('|').find(a => data.idx[a] != null);
      return { column: k.replace(/\|/g, ' or '), found: !!hit,
        sourceIndex: hit ? data.idx[hit] : null, sourceHeader: hit ? data.headers[data.idx[hit]] : null };
    });
  }
  // The output is written by position, so the template's header row has to
  // line up with LOC_HEADERS / SITE_HEADERS column for column.
  function headerCheck(tpl, expected) {
    if (!tpl || !tpl.headers) return null;
    const diffs = [];
    for (let i = 0; i < Math.max(expected.length, tpl.headers.length); i++) {
      const want = expected[i] || null, got = tpl.headers[i] || null;
      if (norm(want || '') !== norm(got || '')) diffs.push({ index: i, expected: want, template: got });
    }
    return { matches: !diffs.length, expectedCount: expected.length,
      templateCount: tpl.headers.length, differences: diffs };
  }
  function tplDropdowns(tpl) {
    return tpl && tpl.dropdowns ? [...tpl.dropdowns.entries()].map(([k, v]) => ({ column: k, count: v.length, values: [...v].slice(0, 300) })) : [];
  }

  function collectDebug() {
    const diagnostics = {
      note: 'events = every message the tool showed and every error thrown while this tool was open, oldest first.',
      events: events.slice(),
      legacyLocationColumns: columnCheck(locData, LOC_KEYS),
      legacySiteColumns: columnCheck(sitesData, SITE_KEYS),
      locationsTemplateHeaders: headerCheck(locTpl, LOC_HEADERS),
      sitesTemplateHeaders: headerCheck(siteTpl, SITE_HEADERS),
      onScreen: {
        summary: $('lm-summary') && $('lm-summary').style.display !== 'none' ? $('lm-summary').innerText : null,
        buildPreviewEnabled: !!($('lm-run') && !$('lm-run').disabled),
        exportEnabled: !!($('lm-export') && !$('lm-export').disabled)
      }
    };
    // Not built yet (a file missing, or the build itself failed): dump what
    // is loaded so the problem can still be seen.
    if (!ready() || !locRows) {
      const missing = [];
      if (!locData) missing.push('Legacy locations CSV');
      if (!sitesData) missing.push('Legacy sites export');
      if (!locTpl) missing.push('3.0 Locations template');
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

  function partialDump(diagnostics, reason) {
    return {
      settings: { sitesScope: refOnly ? 'only sites referenced by the locations' : 'every site in the export' },
      state: { built: !!locRows, reason: reason },
      diagnostics: diagnostics,
      inputs: [
        IMDebug.file('legacy locations export (pt_locations.csv)', locData),
        IMDebug.file('legacy sites export', sitesData),
        { role: 'crop id → legacy crop name lookup', loaded: !!cropLookup, entries: cropLookup ? cropLookup.size : null },
        { role: 'location_type_id → legacy type name lookup', loaded: !!typeLookup, entries: typeLookup ? typeLookup.size : null },
        { role: 'sites already in PickTrace 3.0', loaded: !!migratedSites, entries: migratedSites ? migratedSites.size : null },
        IMDebug.file('personalized bulk Locations template', locTpl, { dropdowns: tplDropdowns(locTpl) }),
        IMDebug.file('personalized bulk Sites template', siteTpl, { dropdowns: tplDropdowns(siteTpl) })
      ],
      asked: [],
      output: null
    };
  }

  function fullDump(diagnostics) {
    const locKept = locRows.map((_, i) => i).filter(i => !locRemoved.has(i));
    const siteKept = siteRows ? siteRows.map((_, i) => i).filter(i => !siteRemoved.has(i)) : [];

    const emptyReqFor = (rows, H, removed) => H.map((h, i) => {
      if (!/\*$/.test(h)) return null;
      let n = 0;
      (rows || []).forEach((r, ri) => { if (!removed.has(ri) && !r[i]) n++; });
      return n ? { column: h, emptyRows: n } : null;
    }).filter(Boolean);

    const cropIds = collectIds('crop');
    const typeIds = collectIds('type');
    const unmappedCrop = [...cropIds.entries()].filter(([id]) => !cropMap[id]);
    const unmappedType = [...typeIds.entries()].filter(([id]) => !typeMap[id]);
    const pendingSites = pendingSitesToCreate();

    const asked = [
      IMDebug.ask('unmapped-location-types', 'data', 'Legacy location_type_id values with no 3.0 Location Type chosen', {
        count: unmappedType.length, blocksExport: true,
        detail: 'Rows carrying these IDs export with an empty Location Type*.',
        items: unmappedType.map(([id, info]) => ({ legacyId: id, rows: info.count, samples: info.samples })) }),
      IMDebug.ask('unmapped-crops', 'data', 'Legacy crop_id values with no Crop & Variety chosen', {
        count: unmappedCrop.length, blocksExport: true,
        detail: 'Rows carrying these IDs export with an empty Crop & Variety*.',
        items: unmappedCrop.map(([id, info]) => ({ legacyId: id, rows: info.count, samples: info.samples })) }),
      IMDebug.ask('sites-to-create', 'dropdown', 'Sites referenced by locations but not in PickTrace 3.0', {
        count: pendingSites.length, blocksExport: true,
        detail: 'Create these sites first — a location cannot be uploaded before its site exists.',
        items: pendingSites.map(([name, info]) => ({ site: name, locations: info.count, legacySiteIds: [...info.ids] })) }),
      IMDebug.ask('empty-required-locations', 'required', 'Required Location columns with empty cells', {
        count: emptyReqFor(locRows, LOC_HEADERS, locRemoved).length, blocksExport: true,
        items: emptyReqFor(locRows, LOC_HEADERS, locRemoved) }),
      IMDebug.ask('empty-required-sites', 'required', 'Required Site columns with empty cells', {
        count: emptyReqFor(siteRows, SITE_HEADERS, siteRemoved).length, blocksExport: true,
        items: emptyReqFor(siteRows, SITE_HEADERS, siteRemoved) }),
      IMDebug.ask('site-ids-without-name', 'data', 'Legacy site_ids not in the sites export and not mapped', {
        count: siteMapProblems(), blocksExport: true,
        items: [...collectSiteIds().entries()].filter(([id, info]) => !(siteMap[id] || info.legacyName))
          .map(([id, info]) => ({ legacySiteId: id, locations: info.count, samples: info.samples })) }),
      IMDebug.ask('guessed-mappings', 'data', 'Crop / Location Type mappings the tool guessed and you have not confirmed', {
        blocksExport: false, items: [...guessed].map(k => { const [kind, id] = k.split('|'); return { kind, legacyId: id, guessedValue: (kind === 'crop' ? cropMap : typeMap)[id] || null }; }) }),
      IMDebug.ask('crops-types-to-create', 'dropdown', 'Crop & Variety / Location Type values the template does not list', {
        blocksExport: false, items: pendingOtherToCreate() }),
      IMDebug.ask('smart-fixes', 'smartfix', 'Values not in a template dropdown, awaiting a Smart Fix', {
        blocksExport: false, items: computeSmartFixes().map(x => ({ file: x.grid, column: x.header, value: x.from, rows: x.count, suggestion: x.suggestion || null })) }),
      IMDebug.ask('site-addresses', 'required', 'Sites missing part of their address', {
        count: siteAddrNeeds(), blocksExport: !!siteTpl }),
      IMDebug.ask('name-collisions', 'collision', 'Locations sharing Site + Name / sites sharing Name', {
        blocksExport: false, items: collisionGroups().map(g => ({ file: g.grid, legacyIds: g.rows.map(ri => (g.grid === 'site' ? siteKeys : locKeys)[ri]) })) })
    ];

    const locOut = IMDebug.output(LOC_HEADERS, locKept.map(i => locRows[i]), {
      note: 'Locations grid. exportedRowIndexes[n] is the locRows index behind exported row n.' });
    locOut.exportedRowIndexes = locKept;
    const siteOut = IMDebug.output(SITE_HEADERS, siteKept.map(i => siteRows[i]), {
      note: 'Sites companion grid. exportedRowIndexes[n] is the siteRows index behind exported row n.' });
    siteOut.exportedRowIndexes = siteKept;

    return {
      settings: {
        activePane: domain,
        sitesScope: refOnly ? 'only sites referenced by the locations' : 'every site in the export',
        legacyCropPresetApplied: presetApplied,
        startDate: 'always the upload day (' + todayIso() + '); the Legacy planted date is in Planted Date',
        addressForEverySite: siteAddrAll || null, addressByRanch: siteAddrGroup, addressBySite: siteAddrOwn,
        clientAddressSheet: addrSheet ? { fileName: addrSheet.fileName, rows: addrSheet.rows, matched: addrSheet.bySite.size, unmatched: addrSheet.unmatched } : null
      },
      state: { built: true },
      diagnostics: diagnostics,
      inputs: [
        IMDebug.file('legacy locations export (pt_locations.csv)', locData),
        IMDebug.file('legacy sites export', sitesData),
        { role: 'crop id → legacy crop name lookup', loaded: !!cropLookup, entries: cropLookup ? cropLookup.size : null },
        { role: 'location_type_id → legacy type name lookup', loaded: !!typeLookup, entries: typeLookup ? typeLookup.size : null },
        { role: 'sites already in PickTrace 3.0', loaded: !!migratedSites, entries: migratedSites ? migratedSites.size : null },
        IMDebug.file('personalized bulk Locations template', locTpl, {
          dropdowns: tplDropdowns(locTpl) }),
        IMDebug.file('personalized bulk Sites template', siteTpl, {
          dropdowns: tplDropdowns(siteTpl) })
      ],
      mapping: {
        note: 'This module applies a fixed legacy→3.0 field map plus three ID lookup tables, rather than auto-matching headers. ' +
              'The lookup tables are under `derived`.',
        locationColumns: LOC_HEADERS.map((h, i) => ({ index: i, header: h, required: /\*$/.test(h) })),
        siteColumns: SITE_HEADERS.map((h, i) => ({ index: i, header: h, required: /\*$/.test(h) }))
      },
      derived: [
        { kind: 'location type lookup', note: 'legacy location_type_id → the 3.0 Location Type you chose.',
          resolved: Object.keys(typeMap).length, unresolved: unmappedType.length,
          table: [...typeIds.entries()].map(([id, info]) => ({ legacyId: id, rows: info.count,
            samples: info.samples, mappedTo: typeMap[id] || null })) },
        { kind: 'crop lookup', note: 'legacy crop_id → the 3.0 Crop & Variety you chose.',
          resolved: Object.keys(cropMap).length, unresolved: unmappedCrop.length,
          legacyPresetApplied: presetApplied,
          table: [...cropIds.entries()].map(([id, info]) => ({ legacyId: id, rows: info.count,
            samples: info.samples, mappedTo: cropMap[id] || null })) },
        { kind: 'site lookup', note: 'legacy site_id → the 3.0 Site name.',
          resolved: Object.keys(siteMap).length,
          table: Object.keys(siteMap).slice(0, 500).map(id => ({ legacySiteId: id, mappedTo: siteMap[id] })) }
      ],
      fills: Object.keys(locFills).map(i => ({ grid: 'locations', column: LOC_HEADERS[+i], value: locFills[i].val,
          scope: locFills[i].mode === 'blank' ? 'blank cells only' : 'all rows', kind: 'column fill' }))
        .concat(Object.keys(siteFills).map(i => ({ grid: 'sites', column: SITE_HEADERS[+i], value: siteFills[i].val,
          scope: siteFills[i].mode === 'blank' ? 'blank cells only' : 'all rows', kind: 'column fill' }))),
      edits: {
        locations: Object.keys(locEdits).slice(0, 500).map(k => { const i = k.lastIndexOf('|'); return { legacyId: k.slice(0, i), column: LOC_HEADERS[+k.slice(i + 1)], value: locEdits[k] }; }),
        sites: Object.keys(siteEdits).slice(0, 500).map(k => { const i = k.lastIndexOf('|'); return { legacyId: k.slice(0, i), column: SITE_HEADERS[+k.slice(i + 1)], value: siteEdits[k] }; }),
        smartFixes: Object.keys(smartFixMap).map(k => ({ key: k, to: smartFixMap[k] })),
        note: 'Edits are keyed by Legacy ID (location_id / site ID), so they survive rebuilds.' },
      dropped: {
        archived: { count: archivedCount, reason: 'is_archived was true in the legacy export.' },
        onSettingsSite: { count: settingsLeft.length, reason: 'On a Legacy Virtual (settings) site such as “Master Employees and Settings” — never migrated.', locations: settingsLeft.slice(0, 200) },
        alreadyInPickTrace: { count: alreadyMigratedCount, reason: 'Site already present in the 3.0 cross-reference file.' },
        locationsRemovedByHand: { count: locRemoved.size, rowIndexes: [...locRemoved].slice(0, 500) },
        sitesRemovedByHand: { count: siteRemoved.size, rowIndexes: [...siteRemoved].slice(0, 500) },
        sitesNotReferenced: refOnly
          ? { note: 'Sites scope is "referenced only" — sites no location points at are not emitted.' }
          : null,
        sitesTickedOffAsCreated: [...completedSites]
      },
      asked: asked,
      output: locOut,
      sitesOutput: siteOut,
      provenance: IMDebug.deriveProvenance({
        headers: LOC_HEADERS, rows: locRows, srcRows: locSrc, colToSrc: {}, fills: locFills }),
      sitesProvenance: IMDebug.deriveProvenance({
        headers: SITE_HEADERS, rows: siteRows || [], srcRows: siteSrc, colToSrc: {}, fills: siteFills })
    };
  }

  // ═══ Export ═══════════════════════════════════════════════════════════
  // Each file patches its template's zip (IMXlsxTemplate, xlsx-template-writer.js
  // — lifted from this module's own copy): only DATA ENTRY's rows change, so
  // every dropdown, style and the header row ship as PickTrace made them.
  function outName(tpl, fallback, suffix) {
    const orig = (tpl && tpl.fileName ? tpl.fileName : fallback).trim();
    const dot = orig.lastIndexOf('.');
    const base = dot > 0 ? orig.substring(0, dot) : orig;
    const ext = dot > 0 ? orig.substring(dot) : '.xlsx';
    return base + ' — filled' + (suffix || '') + ext;
  }

  // ─── Export: this engine's files for "Export all" (legacy-page.js) ───
  // Nothing blocks. Every flagged cell is highlighted; the client reads what
  // each column needs on its header's note, and a cell's own question
  // ("Legacy site #28 … which site?") on that cell. The same asks feed the
  // read-me. Things only we can do (create a site, upload order) go to `team`.
  // What a column needs, said to the client.
  function askFor(grid, ci, flag) {
    const h = (grid === 'site' ? SITE_HEADERS : LOC_HEADERS)[ci].replace(/\*$/, '');
    if (grid === 'loc') {
      if (ci === L_SITE && flag === 'req') return { title: 'Tell us which site these blocks are on',
        detail: 'Their site was archived in Legacy, so they have no site yet. Pick a site for each one — or tell us to leave it out.' };
      if (ci === L_SITE) return { title: 'Check the site for these blocks',
        detail: 'In Legacy they sit on a settings site, not a real farm or ranch. Pick the site each one belongs to — or tell us to leave it out.' };
      if (ci === L_NAME) return { title: 'Rename blocks that share a name', detail: 'Two blocks at the same site have the same name, and 3.0 keeps only one.' };
      if (ci === L_CROP && flag === 'req') return { title: 'Tell us the crop and variety of each block',
        detail: 'Type it in the red cell, e.g. “Blueberry - Duke”. Each cell’s note shows the variety name Legacy had.' };
      if (ci === L_CROP) return { title: 'Check the crop and variety', detail: 'We filled these in for you — change any that are wrong.' };
      if (ci === L_TYPE) return { title: 'Check the location type', detail: 'Legacy didn’t have one, so we used the type shown. Change any that should be different.' };
    } else if (ci === S_NAME) return { title: 'Rename sites that share a name', detail: 'Two sites have the same name, and 3.0 keeps only one.' };
    return flag === 'req' ? { title: 'Fill in ' + h, detail: '' }
      : { title: 'Pick a valid ' + h, detail: 'These values aren’t options in 3.0 — pick one from the list below.' };
  }
  // A cell's own note, when it differs from its neighbours'.
  function noteFor(grid, ri, ci, flag) {
    if (grid === 'loc') {
      const src = locSrc[ri];
      if (ci === L_SITE && flag === 'req') return 'Legacy site #' + locGet(src, 'site_id') + ' (archived)';
      if (ci === L_CROP && flag === 'req') { const v = locGet(src, 'plant_name'); return v ? 'Legacy variety: ' + v : ''; }
      if (ci === L_SITE && flag === 'off') return 'Now on “' + locRows[ri][ci] + '”';
      if (flag === 'off' && ci !== L_TYPE && ci !== L_CROP) return '“' + locRows[ri][ci] + '” isn’t an option in 3.0.';
      return '';
    }
    if (flag === 'off' && !ADDR_CI.has(ci)) return '“' + siteRows[ri][ci] + '” isn’t an option in 3.0.';
    return '';
  }
  // The values a column's dropdown offers, for the read-me.
  function optionsFor(grid, ci) {
    if (grid === 'loc') return ci === L_SITE ? siteChoices() : locOptions(LOC_HEADERS[ci]);
    return siteOptions(SITE_HEADERS[ci]);
  }

  // ─── Questions for the client (legacy-form.js) — and their answers back ───
  // Everything the client has to settle, as plain questions with the choices
  // they can make. Answers come back keyed by Legacy ID (location_id / site
  // ID) or by the value they're about, so they land on the same rows when
  // the same Legacy files are loaded again.
  const LEAVE_OUT = { value: '__leave__', label: 'Leave it out — don’t move it to 3.0' };
  const plural = (n, one, many) => n + ' ' + (n === 1 ? one : many);
  // What a client would recognise a Legacy block by.
  function blockFacts(src) {
    const out = [], ha = parseFloat(locGet(src, 'hectares'));
    const v = locGet(src, 'plant_name'), cid = locGet(src, 'crop_id');
    if (v || cid) out.push(['Variety', (v || 'none') + (cid ? ' (Legacy crop #' + cid + ')' : '')]);
    if (ha > 0) out.push(['Size', (Math.round(ha * 2.47105 * 100) / 100) + ' acres']);
    const pd = locGet(src, 'planted_date');
    if (pd) out.push(['Planted', pd.slice(0, 10)]);
    const d = locGet(src, 'description') || locGet(src, 'info');
    if (d) out.push(['Notes', d.length > 80 ? d.slice(0, 78) + '…' : d]);
    out.push(['Legacy ID', locGet(src, 'location_id')]);
    return out;
  }
  function questions() {
    pvCache = null;
    const Q = [];
    if (!locRows) return Q;
    const kept = locRows.map((_, i) => i).filter(i => !locRemoved.has(i));
    // 1. Blocks whose Legacy site was archived.
    const orphan = kept.filter(ri => !locRows[ri][L_SITE]);
    if (orphan.length) Q.push({ id: 'lm.orphan', type: 'pick', where: 'Blocks · ' + plural(orphan.length, 'block', 'blocks'),
      title: 'Which site is each of these blocks on?',
      help: 'Their site was archived in Legacy, so they don’t have one yet. Pick the site for each — or choose “Leave it out”.',
      options: siteChoices(), extra: [LEAVE_OUT],
      layout: 'detail', answerLabel: 'Site in 3.0', groupPick: 'Same site for all',
      items: orphan.slice().sort((x, y) => (+locGet(locSrc[x], 'site_id') || 0) - (+locGet(locSrc[y], 'site_id') || 0)).map(ri => {
        const src = locSrc[ri], sid = locGet(src, 'site_id'), same = orphan.filter(r => locGet(locSrc[r], 'site_id') === sid).length;
        return { key: locKeys[ri], label: locRows[ri][L_NAME],
          group: 'Legacy site #' + sid + ' (archived)' + (same > 1 ? ' — ' + same + ' blocks' : ''),
          facts: blockFacts(src) };
      }) });
    // 3. Crops — one answer per Legacy crop number, shown with its varieties.
    const cropIds = [...collectIds('crop').entries()].filter(([id]) => !NO_CROP.has(id));
    const askCrop = cropIds.filter(([id]) => !cropMap[id] || guessed.has('crop|' + id));
    const varsOf = id => [...new Set(kept.filter(ri => locGet(locSrc[ri], 'crop_id') === id).map(ri => locGet(locSrc[ri], 'plant_name')).filter(Boolean))];
    if (askCrop.length) Q.push({ id: 'lm.cropname', type: 'pick', other: true, where: 'Blocks · ' + plural(askCrop.length, 'crop', 'crops'),
      title: 'What crop is each of these?', layout: 'detail', answerLabel: 'Crop in 3.0',
      help: 'Legacy kept each crop as a number. Its variety names are shown to help you tell which crop it is. Not a crop (an office, a shed)? Choose “Other”.',
      placeholder: 'e.g. Blueberries', options: cropNameOptions(),
      items: askCrop.sort((a, b) => b[1].count - a[1].count).map(([id, info]) => {
        const v = varsOf(id);
        const rows = kept.filter(ri => locGet(locSrc[ri], 'crop_id') === id);
        const siteCount = new Map();
        rows.forEach(ri => { const st = locRows[ri][L_SITE] || '(no site yet)'; siteCount.set(st, (siteCount.get(st) || 0) + 1); });
        const sites = [...siteCount.entries()].sort((a, b) => b[1] - a[1]);
        return { key: id, label: 'Legacy crop #' + id, value: cropMap[id] ? cropPart(cropMap[id]) : '',
          facts: [['Blocks', plural(info.count, 'block', 'blocks')],
            ['Varieties', v.length ? v : ['none recorded']],
            ['On sites', sites.slice(0, 4).map(([n, c]) => n + ' (' + c + ')').concat(sites.length > 4 ? ['+' + (sites.length - 4) + ' more'] : [])],
            ['For example', info.samples.slice(0, 3)]] };
      }) });
    // 4. Variety names — written the way they'll appear; look-alikes grouped so
    // they can be given one name (BLUE CROP / BLUE CROP MACHINE → "Blue Crop").
    const groups = varietyGroups();
    if (groups.length) {
      const items = [];   // varieties the client already named aren't asked again
      groups.forEach(g => g.members.forEach(m => varAnswered.has(m.key) || items.push({ key: m.key, label: m.plant || '(no variety in Legacy)',
        sub: plural(m.count, 'block', 'blocks'), value: varietyFor(m.plant),
        group: g.members.length > 1 ? 'These look alike: ' + g.members.map(x => x.plant || '(none)').join(' · ') : 'On its own' })));
      if (items.length) Q.push({ id: 'lm.variety', type: 'pick', other: true, where: 'Blocks · ' + plural(items.length, 'variety', 'varieties'),
        title: 'How should your varieties be named in 3.0?',
        help: 'This is how each Legacy variety name will appear. Change any you like. To combine look-alikes, give them the same name — e.g. “Blue Crop” for both BLUE CROP and BLUE CROP MACHINE. Blocks with no variety get “Other”.',
        placeholder: 'Variety name', items });
    }
    // A dropdown's everyday values first (they become the buttons), then the
    // org's others — its template may carry test values like "a01-audit-site".
    const everydayFirst = (opts, common) => {
      const out = common.map(c => opts.find(o => o.toUpperCase() === c.toUpperCase())).filter(Boolean);
      opts.slice().sort((x, y) => x.localeCompare(y)).forEach(o => { if (out.indexOf(o) < 0) out.push(o); });
      return out;
    };
    // 4. Location type, block by block: blocks grouped by site; select rows or
    //    a whole site, then press a type.
    const gTypes = [...collectIds('type').keys()].filter(id => guessed.has('type|' + id) && typeMap[id]);
    if (gTypes.length) {
      const ids = new Set(gTypes);
      const rows = kept.filter(ri => ids.has(locGet(locSrc[ri], 'location_type_id') || '(blank)'));
      const used = [...new Set(rows.map(ri => locRows[ri][L_TYPE]))];
      // The everyday types first as buttons; the org's other types under "More…".
      const opts = locOptions('Location Type*');
      const quick = everydayFirst(opts, ['Block', 'Field', 'Section', 'Ranch', 'Row', 'Facility', 'Packing House', 'Office', 'Shop', 'Warehouse', 'Nursery', 'Yard']);
      if (rows.length) Q.push({ id: 'lm.types', type: 'assign', where: 'Blocks · ' + plural(rows.length, 'block', 'blocks'),
        title: 'What type is each block?',
        help: 'Legacy didn’t record a type, so every block starts as “' + used.join(' / ') + '”. Only change what’s different: open a site, ' +
          'click blocks to select them (Ctrl/⌘+click adds or takes away one, Shift+click a range) — or a site’s box for the whole site — then press a type, or Remove to leave them out of 3.0.',
        options: opts, quick, groupLabel: 'Site',
        remove: { label: 'Remove', help: 'Leave the selected blocks out — they won’t be moved to 3.0' },
        items: rows.map(ri => ({ key: locKeys[ri], label: locRows[ri][L_NAME], group: locRows[ri][L_SITE] || '(no site yet)', value: locRows[ri][L_TYPE] })) });
    }
    // 6. Site details: one address for every site, or one each; Site Type
    //    (and anything else required) as buttons for all of them at once.
    if (siteTpl && siteRows) {
      const sk = siteRows.map((_, i) => i).filter(i => !siteRemoved.has(i));
      const cols = SITE_HEADERS.map((h, ci) => ci).filter(ci => ci !== S_NAME && ci !== 5 && /\*$/.test(SITE_HEADERS[ci]) && sk.some(ri => !siteRows[ri][ci]));
      const rows = sk.filter(ri => cols.some(ci => !siteRows[ri][ci]));
      const choiceCols = cols.filter(ci => !ADDR_CI.has(ci));
      if (rows.length) Q.push({ id: 'lm.sites', type: 'places', where: 'Sites · ' + plural(rows.length, 'site', 'sites'),
        title: 'Your sites’ details',
        help: '3.0 needs an address for each site. If they all share one (your main office, say), enter it once.',
        address: true,
        choices: choiceCols.map(ci => ({ key: String(ci), label: SITE_HEADERS[ci].replace(/\*$/, ''),
          options: everydayFirst(siteOptions(SITE_HEADERS[ci]), ['Ranch', 'Farm', 'Field', 'Packing House', 'Office', 'Shop', 'Warehouse', 'Facility', 'Nursery', 'Yard', 'Housing']) })),
        items: rows.map(ri => {
          const r = siteRows[ri], v = {};
          choiceCols.forEach(ci => { v[String(ci)] = r[ci] || ''; });
          SITE_ADDR_PARTS.forEach(([k, ci]) => { v[k] = r[ci] || ''; });
          return { key: siteKeys[ri], label: r[S_NAME], values: v };
        }) });
    }
    // 7. Values that aren't options in 3.0 — one question per column.
    const fixes = computeSmartFixes(), byCol = new Map();
    fixes.forEach(f => { const k = f.grid + '.' + f.colIdx; (byCol.get(k) || byCol.set(k, []).get(k)).push(f); });
    byCol.forEach((list, k) => Q.push({ id: 'lm.fix.' + k, type: 'pick', where: (list[0].grid === 'site' ? 'Sites' : 'Blocks') + ' · ' + list[0].header.replace(/\*$/, ''),
      title: 'Which ' + list[0].header.replace(/\*$/, '') + ' do you mean?',
      help: 'These values from Legacy aren’t options in 3.0. Pick the one each should be.',
      options: list[0].options.slice().sort(),
      items: list.map(f => ({ key: String(f.from).toUpperCase().trim(), label: '“' + f.from + '”', sub: plural(f.count, 'row', 'rows'), value: f.suggestion || '' })) }));
    // 8. Name clashes — a new name for each.
    const clash = [];
    collisionGroups().forEach(g => g.rows.forEach(ri => {
      const site = g.grid === 'site';
      clash.push({ key: (site ? 'site:' : 'loc:') + (site ? siteKeys : locKeys)[ri],
        label: site ? siteRows[ri][S_NAME] : locRows[ri][L_NAME] + ' — at ' + locRows[ri][L_SITE], sub: 'Shares its name with another ' + (site ? 'site' : 'block') });
    }));
    if (clash.length) Q.push({ id: 'lm.rename', type: 'pick', other: true, where: 'Names',
      title: 'These share a name — what should each be called?',
      help: '3.0 keeps only one of each name. Give each a name of its own (keeping one as it is is fine).', items: clash });
    return Q;
  }

  // Apply the client's answers (from IMLegacyForm.read) → lines saying what changed.
  function applyAnswers(A) {
    const out = [];
    const pick = id => A[id] || {};
    ['lm.orphan', 'lm.offsite'].forEach(id => {
      let site = 0, left = 0;
      Object.entries(pick(id)).forEach(([key, v]) => {
        if (!v) return;
        if (v === '__leave__') { locRemovedK.add(key); left++; }
        else { locEdits[key + '|' + L_SITE] = v; site++; }
      });
      if (site || left) out.push((id === 'lm.orphan' ? 'Archived-site blocks' : 'Settings-site blocks') + ': ' + site + ' given a site' + (left ? ', ' + left + ' left out' : ''));
    });
    const crop = pick('lm.crop');
    if (Object.keys(crop).length) {
      let n = 0;
      locRows.forEach((r, ri) => {
        const v = crop[locGet(locSrc[ri], 'plant_name')];
        if (!v || !(!r[L_CROP] || cellFlag('loc', ri, L_CROP)[0] === 'off')) return;
        locEdits[locKeys[ri] + '|' + L_CROP] = v; n++;
      });
      out.push('Crop & variety: ' + n + ' blocks filled in');
    }
    let named = 0;
    Object.entries(pick('lm.cropname')).forEach(([id, v]) => { if (v) { cropMap[id] = cropPart(v); guessed.delete('crop|' + id); named++; } });
    if (named) out.push('Crops: ' + named + ' named');
    let varRenamed = 0;
    Object.entries(pick('lm.variety')).forEach(([k, v]) => { if (!v || !v.trim()) return; varAnswered.add(k); if (v.trim() !== tidyVariety(k)) varRenamed++; varietyMap[k] = v.trim(); });
    if (Object.keys(pick('lm.variety')).length) out.push('Varieties: ' + varRenamed + ' renamed');
    // Location types, block by block. Every guessed type counts as confirmed.
    const types = pick('lm.types');
    if (Object.keys(types).length) {
      let changed = 0, removed = 0;
      const byKey = new Map(locKeys.map((k, ri) => [k, ri]));
      Object.entries(types).forEach(([key, v]) => {
        const ri = byKey.get(key);
        if (!v || ri == null) return;
        if (v === '__leave__') { locRemovedK.add(key); removed++; return; }   // “Remove”: not moved to 3.0
        if (v !== locRows[ri][L_TYPE]) { locEdits[key + '|' + L_TYPE] = v; changed++; }
      });
      [...collectIds('type').keys()].forEach(id => guessed.delete('type|' + id));
      out.push('Location types: confirmed' + (changed ? ', ' + changed + ' block' + (changed === 1 ? '' : 's') + ' changed' : '') + (removed ? ', ' + removed + ' removed' : ''));
    }
    const t = A['lm.type'];
    if (t && t.choice) {
      const ids = [...collectIds('type').keys()];
      if (t.choice === 'keep') ids.forEach(id => guessed.delete('type|' + id));
      else if (t.choice === 'change' && t.value) ids.forEach(id => { typeMap[id] = t.value; guessed.delete('type|' + id); });
      out.push('Location type: ' + (t.choice === 'keep' ? 'confirmed' : t.choice === 'change' ? 'changed to “' + t.value + '”' : 'needs a closer look — “' + (t.value || '') + '”'));
    }
    const sites = pick('lm.sites');
    let cells = 0, siteLeft = 0, blocksLeft = 0, countries = 0;
    // Address parts come back by name (address1, city…), other columns by number.
    const partCol = Object.fromEntries(SITE_ADDR_PARTS);
    Object.entries(sites).forEach(([key, vals]) => {
      if (key.indexOf('__') === 0) return;          // the form's own keys (__mode, __all__)
      vals = vals || {};
      // "Leave it out" on the page, or "delete" / "remove" typed into the
      // address: the site isn't moved, and neither are any blocks still on it.
      const addr = ['address1', 'city', 'state', 'zip'].map(k => String(vals[k] || '').trim()).filter(Boolean);
      if (vals.__leave || (addr.length && addr.every(v => DROP_WORD.test(v)))) {
        siteRemovedK.add(key); siteLeft++;
        if (locData) locData.rows.forEach((row, i) => {
          const lk = locKeyOf(row, i);
          if ((locGet(row, 'site_id') || '(blank)') === key && !locRemovedK.has(lk)) { locRemovedK.add(lk); blocksLeft++; }
        });
        return;
      }
      Object.entries(vals).forEach(([k, v]) => {
        const ci = partCol[k] != null ? partCol[k] : (/^\d+$/.test(k) ? +k : null);
        v = String(v == null ? '' : v).trim();
        // "n/a", "none", "-": nothing to write — the cell stays as it was.
        if (ci == null || !v || NO_VALUE.test(v) || DROP_WORD.test(v)) return;
        siteEdits[key + '|' + ci] = v; cells++;
      });
      // A US state with no country: the country is the US (as the address sheet does).
      if (!String(vals.country || '').trim() && vals.state) {
        const p = parseAddr('x, x, ' + String(vals.state).trim().toUpperCase() + ' 00000');
        if (p && p.country) { siteEdits[key + '|' + partCol.country] = p.country; countries++; }
      }
    });
    // A typo'd city ("Anytwon") among sites sharing its zip, where most spell it
    // one way ("Anytown"): the way most of them spell it.
    const cityFix = [];
    {
      const cityCi = partCol.city, zipCi = partCol.zip, byZip = new Map();
      Object.keys(sites).forEach(key => {
        const city = siteEdits[key + '|' + cityCi], zip = String(siteEdits[key + '|' + zipCi] || '').slice(0, 5);
        if (!city || !/^\d{5}$/.test(zip)) return;
        (byZip.get(zip) || byZip.set(zip, []).get(zip)).push(key);
      });
      const lev = (a, b) => {
        let p = Array.from({ length: b.length + 1 }, (_, j) => j);
        for (let i = 1; i <= a.length; i++) {
          const c = [i];
          for (let j = 1; j <= b.length; j++) c[j] = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
          p = c;
        }
        return p[b.length];
      };
      byZip.forEach(keys => {
        const n = new Map();
        keys.forEach(k => { const c = siteEdits[k + '|' + cityCi]; n.set(c, (n.get(c) || 0) + 1); });
        const top = [...n.entries()].sort((a, b) => b[1] - a[1])[0];
        if (!top || top[1] < 2) return;
        keys.forEach(k => {
          const c = siteEdits[k + '|' + cityCi];
          if (c === top[0] || n.get(c) * 2 > top[1]) return;
          const d = lev(c.toLowerCase(), top[0].toLowerCase());
          if (d === 0 || d > (c.length < 8 ? 1 : 2)) return;
          siteEdits[k + '|' + cityCi] = top[0];
          cityFix.push(c + ' → ' + top[0]);
        });
      });
    }
    if (cells) out.push('Site details: ' + cells + ' cells filled in' + (countries ? ' (country ' + countries + ' from the state)' : '') +
      (cityFix.length ? '; city spelling fixed to match the same zip: ' + [...new Set(cityFix)].join(', ') : ''));
    if (siteLeft) out.push('Sites left out: ' + siteLeft + ' (the client marked them delete / leave out)' + (blocksLeft ? ', with ' + plural(blocksLeft, 'block', 'blocks') + ' still on them' : ''));
    Object.keys(A).filter(k => k.indexOf('lm.fix.') === 0).forEach(k => {
      const [grid, ci] = k.slice(7).split('.');
      let n = 0;
      Object.entries(A[k] || {}).forEach(([u, v]) => { if (v) { smartFixMap[grid + '|' + ci + '||' + u] = v; n++; } });
      if (n) out.push('Fixed ' + n + ' value' + (n === 1 ? '' : 's') + ' not in a dropdown');
    });
    let renamed = 0;
    Object.entries(pick('lm.rename')).forEach(([k, v]) => {
      if (!v) return;
      const i = k.indexOf(':'), grid = k.slice(0, i), key = k.slice(i + 1);
      if (grid === 'site') siteEdits[key + '|' + S_NAME] = v; else locEdits[key + '|' + L_NAME] = v;
      renamed++;
    });
    if (renamed) out.push('Renamed ' + renamed);
    rebuildAndRender();
    return out;
  }

  async function exportParts(limit) {
    pvCache = null;
    const locKept = locRows.map((_, i) => i).filter(i => !locRemoved.has(i));
    const siteKept = siteRows && siteTpl ? siteRows.map((_, i) => i).filter(i => !siteRemoved.has(i)) : [];
    const files = [], items = [];
    const one = async (grid, kept, tpl, H, label, fallback) => {
      const rows = grid === 'site' ? siteRows : locRows;
      const nameOf = (chunk, total) => outName(tpl, fallback, total > 1 ? ' (' + (chunk + 1) + ' of ' + total + ')' : '');
      const a = IMLegacyDrop.annotate({
        label, H, total: kept.length, limit, file: nameOf,
        unit: grid === 'site' ? ['site', 'sites'] : ['block', 'blocks'],
        options: ci => optionsFor(grid, ci),
        cells: pos => {
          const ri = kept[pos], out = [];
          H.forEach((h, ci) => {
            const [flag] = cellFlag(grid, ri, ci);
            if (flag) out.push({ ci, flag, note: noteFor(grid, ri, ci, flag) });
          });
          return out;
        },
        nameOf: pos => rows[kept[pos]][grid === 'site' ? S_NAME : L_NAME],
        ask: (ci, flag) => askFor(grid, ci, flag)
      });
      items.push(...a.items);
      const chunks = chunkArr(kept.map(ri => rows[ri].slice()), limit);
      for (let i = 0; i < chunks.length; i++) {
        const blob = await IMXlsxTemplate.build({ rawBuffer: tpl.rawBuffer, headers: H, rows: chunks[i],
          marks: a.chunks[i].marks, notes: a.chunks[i].notes, noteAuthor: 'PickTrace' });
        files.push({ label, name: nameOf(i, chunks.length), data: new Uint8Array(await blob.arrayBuffer()) });
      }
    };
    // Sites first — they must exist in 3.0 before the locations that use them.
    if (siteKept.length) await one('site', siteKept, siteTpl, SITE_HEADERS, 'Sites', '3.0_Sites_bulk_create.xlsx');
    await one('loc', locKept, locTpl, LOC_HEADERS, 'Locations', '3.0_Locations_bulk_create.xlsx');
    const team = [];
    const toMake = pendingSitesToCreate().map(([n]) => n), other = pendingOtherToCreate();
    if (toMake.length) team.push('Create site' + (toMake.length === 1 ? '' : 's') + ' in 3.0: ' + IMLegacyDrop.examples(toMake, 10));
    ['Crop & Variety', 'Location Type'].forEach(k => {
      const v = other.filter(e => e.kind === k).map(e => e.value);
      if (v.length) team.push('Create ' + k + ' in 3.0: ' + IMLegacyDrop.examples(v, 10));
    });
    if (siteKept.length) team.push('Upload the Sites file before the Locations file.');
    return { files, items, team };
  }

  // ─── Reset ───
  function reset() {
    logEvent('reset', 'Reset clicked — files and choices cleared.');
    locData = null; sitesData = null; cropLookup = null; typeLookup = null;
    migratedSites = null; locTpl = null; siteTpl = null;
    cropMap = {}; varietyMap = {}; varAnswered = new Set(); typeMap = {}; siteMap = {}; presetApplied = false; guessed = new Set();
    locRows = null; locSrc = null; siteRows = null; siteSrc = null;
    locFills = {}; siteFills = {};
    locRemoved = new Set(); siteRemoved = new Set();
    locKeys = []; siteKeys = []; locRemovedK = new Set(); siteRemovedK = new Set();
    locEdits = {}; siteEdits = {}; smartFixMap = {};
    locRaw = null; siteRaw = null; locWhy = {}; siteWhy = {};
    siteAddrAll = ''; siteAddrGroup = {}; siteAddrOwn = {}; addrSheet = null;
    archivedCount = 0; alreadyMigratedCount = 0; skippedSiteNames = [];
    completedSites = new Set(); completedBuild = new Set();
    selCells = new Set(); selAnchor = null; previewOrder = [];
    pvMode = 'all'; pvColsOnly = false;
    domain = 'locations';
    { const b = $('lm-sel-bar'); if (b) b.style.display = 'none'; }
    [['lm-loc', 'No file selected'], ['lm-sites', 'No file selected'],
     ['lm-croplk', 'No file selected'], ['lm-typelk', 'No file selected'],
     ['lm-migrated', 'No file selected'], ['lm-loctpl', 'No file selected'],
     ['lm-sitetpl', 'No file selected'], ['lm-addr', 'No file selected']].forEach(([p, txt]) => {
      const n = $(p + '-name'); if (n) n.textContent = txt;
      const m = $(p + '-meta'); if (m) m.textContent = '';
      const f = $(p + '-file'); if (f) f.value = '';
    });
    ['lm-section-sitemap', 'lm-section-type', 'lm-section-crop', 'lm-section-create',
     'lm-section-bulk', 'lm-section-preview', 'lm-section-cv-create', 'lm-section-smartfix',
     'lm-section-sitefile', 'lm-section-siteaddr', 'lm-section-collisions',
     'lm-section-left'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    { const b = $('lm-preset-banner'); if (b) b.style.display = 'none'; }
    $('lm-summary').style.display = 'none';
    $('lm-empty').style.display = '';
    if (window.IMDebug) IMDebug.refresh('locations-migrate');
  }

  // ─── Inline cell editing ───
  function commitCellEdit(td) {
    const rows = curRows();
    if (!rows) return;
    const ri = +td.dataset.ri, ci = +td.dataset.ci;
    if (!rows[ri]) return;
    const val = td.textContent.trim();
    if (rows[ri][ci] === val) return;
    rows[ri][ci] = val;
    const grid = domain === 'sites' ? 'site' : 'loc';
    curEdits()[(grid === 'site' ? siteKeys : locKeys)[ri] + '|' + ci] = val;   // kept through rebuilds
    whyOf(grid)[ri + '|' + ci] = 'Your edit';
    // Repaint just this cell — redrawing the grid here would throw away the
    // cell you just tabbed into. The panels and counts catch up.
    pvCache = null;
    const [fl, tip] = cellFlag(grid, ri, ci), note = changeOf(grid, ri, ci);
    td.setAttribute('style', FLAG_STYLE[fl] || '');
    td.classList.toggle('is-chg', !!note);
    td.title = [tip, note].filter(Boolean).join(' · ');
    renderSiteMap();
    renderSitesToCreate();
    renderOtherToCreate();
    renderSmartFixes();
    renderCollisions();
    renderSiteFile();
    updateSummary();
    updateExportButton();
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    if (!$('lm-loc-file')) return; // markup not present yet
    initialized = true;

    $('lm-loc-file').addEventListener('change', e => {
      if (e.target.files[0]) handleLocFile(e.target.files[0]); e.target.value = '';
    });
    $('lm-sites-file').addEventListener('change', e => {
      if (e.target.files[0]) handleSitesFile(e.target.files[0]); e.target.value = '';
    });
    $('lm-croplk-file').addEventListener('change', e => {
      if (e.target.files[0]) handleLookupFile(e.target.files[0], 'crop'); e.target.value = '';
    });
    $('lm-typelk-file').addEventListener('change', e => {
      if (e.target.files[0]) handleLookupFile(e.target.files[0], 'type'); e.target.value = '';
    });
    $('lm-migrated-file').addEventListener('change', e => {
      if (e.target.files[0]) handleMigratedSitesFile(e.target.files[0]); e.target.value = '';
    });
    $('lm-loctpl-file').addEventListener('change', e => {
      if (e.target.files[0]) readTemplate(e.target.files[0], 'loc'); e.target.value = '';
    });
    $('lm-sitetpl-file').addEventListener('change', e => {
      if (e.target.files[0]) readTemplate(e.target.files[0], 'site'); e.target.value = '';
    });

    const ro = $('lm-refonly');
    if (ro) ro.addEventListener('change', e => {
      refOnly = e.target.checked;
      if (ready()) { rebuild(); renderAll(); }
    });
    const batch = $('lm-batch');
    if (batch) batch.addEventListener('input', () => { if (locRows) updateSummary(); });

    // Build, Export all, Reset and Debug Dump belong to the whole page (legacy-page.js).
    if (window.IMDebug) {
      IMDebug.register('locations-migrate', {
        label: 'Locations & Sites Migrate (Legacy → 3.0)',
        // Any file loaded, or anything logged, is enough to dump: the point is
        // to see what went wrong even when the build never ran.
        ready: () => !!(locData || sitesData || locTpl || siteTpl || cropLookup || addrSheet ||
          typeLookup || migratedSites || events.length),
        collect: collectDebug
      });
    }

    registerPage();
    { const af = $('lm-addr-file'); if (af) af.addEventListener('change', e => {
        if (e.target.files[0]) handleAddrFile(e.target.files[0]); e.target.value = '';
      }); }
    // The panels are redrawn on every render, so their controls are handled
    // here, once, on the section that holds them.
    const on = (id, ev, fn) => { const el = $(id); if (el) el.addEventListener(ev, fn); };
    on('lm-section-crop', 'change', e => {
      const el = e.target.closest('.lm-var-in');
      if (!el) return;
      const v = el.value.trim();
      if (v) varietyMap[el.dataset.key] = v; else delete varietyMap[el.dataset.key];
      rebuildAndRender();
    });
    on('lm-section-sitefile', 'change', e => {
      const el = e.target.closest('.lm-sf-fill');
      if (!el) return;
      const ci = +el.dataset.ci, v = el.value.replace(/^\s+/, '');
      if (v) siteFills[ci] = { val: v, mode: 'all' }; else delete siteFills[ci];
      dropEditsIn(siteEdits, ci);
      rebuildAndRender();
    });
    on('lm-section-sitefile', 'click', e => {
      const b = e.target.closest('.lm-sf-clear');
      if (!b) return;
      delete siteFills[+b.dataset.ci];
      rebuildAndRender();
    });
    on('lm-section-siteaddr', 'change', e => {
      const el = e.target.closest('.lm-addr-in');
      if (!el) return;
      const v = el.value.trim(), k = el.dataset.key;
      if (el.dataset.scope === 'all') siteAddrAll = v;
      else if (el.dataset.scope === 'ranch') { if (v) siteAddrGroup[k] = v; else delete siteAddrGroup[k]; }
      else { if (v) siteAddrOwn[k] = v; else delete siteAddrOwn[k]; }
      rebuildAndRender();
    });
    on('lm-section-smartfix', 'click', e => {
      const one = e.target.closest('.lm-smartfix-apply');
      const all = e.target.closest('#lm-smartfix-apply-all');
      if (!one && !all) return;
      const n = applySmartFixRows(one ? [one] : [...$('lm-smartfix-table').querySelectorAll('.lm-smartfix-apply')]);
      if (!n) { alert(one ? 'Pick a value to set "' + one.dataset.from + '" to first.' : 'No values chosen yet — pick a target for at least one row.'); return; }
      rebuildAndRender();
    });
    on('lm-section-cv-create', 'click', e => {
      const tr = e.target.closest('.lm-cv-row');
      if (!tr) return;
      const k = tr.dataset.k;
      if (completedBuild.has(k)) completedBuild.delete(k); else completedBuild.add(k);
      if (navigator.clipboard) navigator.clipboard.writeText(tr.dataset.v).catch(() => {});
      renderOtherToCreate();
      updateSummary();
    });
    on('lm-section-left', 'click', e => {
      if (!e.target.closest('#lm-restore-all')) return;
      locRemovedK.clear(); siteRemovedK.clear();
      rebuildAndRender();
    });
    const copyBtn = $('lm-create-copy');
    if (copyBtn) copyBtn.addEventListener('click', () => {
      const names = [...getSitesToCreate().keys()];
      if (!names.length) { alert('No sites to create.'); return; }
      navigator.clipboard.writeText(names.join('\n')).then(
        () => alert('Copied ' + names.length + ' site name' + (names.length === 1 ? '' : 's') + '.'),
        () => alert('Copy failed.'));
    });

    // Locations / Sites preview toggle.
    document.querySelectorAll('.lm-domain-pill:not([data-domain="employees"])').forEach(btn => {
      btn.addEventListener('click', () => {
        domain = btn.dataset.domain;
        clearSelection();
        IMWorkbench.rowsTop($('lm-preview-table'));
        renderBulkEdit();
        renderPreview();
        if (wb) wb.refresh();
      });
    });

    // Delegated inline-edit handlers bound ONCE to the persistent table element
    // (renderPreview only replaces its innerHTML).
    const ptbl = $('lm-preview-table');
    if (ptbl) {
      ptbl.addEventListener('focusout', e => {
        const td = e.target.closest ? e.target.closest('td.lm-cell') : null;
        if (td) commitCellEdit(td);
      });
      ptbl.addEventListener('keydown', e => {
        if (e.key === 'Escape' && selCells.size) { clearSelection(); return; }
        const td = e.target.closest ? e.target.closest('td.lm-cell') : null;
        if (td && e.key === 'Enter') { e.preventDefault(); td.blur(); }
      });
      ptbl.addEventListener('mousedown', e => {
        const td = e.target.closest ? e.target.closest('td.lm-cell') : null;
        if (!td) return;
        const ri = +td.dataset.ri, ci = +td.dataset.ci;
        if (e.shiftKey) {
          e.preventDefault(); selRangeTo(ri, ci); paintSelection();
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
    const selApply = $('lm-sel-apply');
    if (selApply) {
      selApply.addEventListener('click', () => setSelectedCells($('lm-sel-val').value));
      $('lm-sel-clear').addEventListener('click', () => setSelectedCells(''));
      $('lm-sel-deselect').addEventListener('click', clearSelection);
      $('lm-sel-val').addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); setSelectedCells(e.target.value); }
      });
    }
  }

  window.locMigInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
