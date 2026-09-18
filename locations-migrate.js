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
    'Custom Data 1', 'Custom Data 2', 'Custom Data 3', 'Internal Commission'
  ];

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
    id: 'legacy-sfc',
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
      if (v) { cropMap[cid] = v; n++; }
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

  let cropMap = {};          // crop_id → chosen 'Crop & Variety*' (sticky)
  let typeMap = {};          // location_type_id → chosen 'Location Type*' (sticky)
  let siteMap = {};          // legacy site_id → chosen 'Site*' (sticky)
  let presetApplied = false; // the Legacy crop preset is currently in effect

  let locRows = null;        // [[...42]] mapped location rows
  let locSrc = null;         // parallel source rows
  let siteRows = null;       // [[...10]] mapped site rows
  let siteSrc = null;        // parallel source rows

  let locFills = {}, siteFills = {};       // colIdx → { val, mode }
  let locRemoved = new Set(), siteRemoved = new Set();
  let archivedCount = 0;                   // dropped by is_archived
  let alreadyMigratedCount = 0;            // dropped by the 3.0 cross-reference
  let refOnly = true;                      // Sites: only sites the locations use
  let domain = 'locations';                // which preview/bulk pane is showing
  let selCells = new Set(), selAnchor = null, previewOrder = [];
  let completedSites = new Set();          // sites ticked off the "to create" list
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
    const r = new FileReader();
    r.onload = e => {
      try {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
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
        alert('Failed to read template: ' + (err && err.message ? err.message : err));
      }
    };
    r.onerror = () => alert('Failed to read template file.');
    r.readAsArrayBuffer(file);
  }

  // ─── Load handlers ───
  function handleLocFile(file) {
    parseAnyFile(file).then(d => {
      if (d.idx['site_id'] == null || d.idx['location_id'] == null) {
        alert('Legacy locations file must have "site_id" and "location_id" columns.');
        return;
      }
      locData = d;
      $('lm-loc-name').textContent = file.name;
      const arch = d.rows.filter(r => String(r[d.idx['is_archived']] || '').trim().toLowerCase() === 'true').length;
      $('lm-loc-meta').textContent = d.rows.length + ' rows · ' + arch + ' archived';
      maybeRun();
    }).catch(err => alert('Failed to read locations file: ' + (err && err.message ? err.message : err)));
  }

  function handleSitesFile(file) {
    parseAnyFile(file).then(d => {
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
        if (id) byId.set(id, nm);
      });
      sitesData = Object.assign(d, { byId, idI, nameI });
      $('lm-sites-name').textContent = file.name;
      $('lm-sites-meta').textContent = byId.size + ' site' + (byId.size === 1 ? '' : 's');
      maybeRun();
    }).catch(err => alert('Failed to read sites file: ' + (err && err.message ? err.message : err)));
  }

  // Optional Legacy lookup: ID → Name, for crops or location types.
  function handleLookupFile(file, kind) {
    parseAnyFile(file).then(d => {
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
    }).catch(err => alert('Failed to read lookup file: ' + (err && err.message ? err.message : err)));
  }

  // Optional: a 3.0 sites export. Any Legacy site whose name already appears
  // here is skipped so the Sites upload doesn't create duplicates.
  function handleMigratedSitesFile(file) {
    parseAnyFile(file).then(d => {
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
    }).catch(err => alert('Failed to read 3.0 sites file: ' + (err && err.message ? err.message : err)));
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

  function suggestionFor(kind, id, info) {
    const opts = kind === 'crop' ? locOptions('Crop & Variety*') : locOptions('Location Type*');
    if (!opts.length) return null;
    const lk = kind === 'crop' ? cropLookup : typeLookup;
    if (lk && lk.has(id)) {
      const s = suggestFromName(lk.get(id), opts);
      if (s) return { value: s.value, confidence: s.confidence, via: 'lookup: ' + lk.get(id) };
    }
    const s = suggestFromTexts(info.texts, opts);
    if (s) return { value: s.value, confidence: s.confidence, via: 'inferred from location names' };
    return null;
  }

  // ─── Locations transform ───
  function buildLocRowsRaw() {
    if (!locData) return null;
    locSrc = [];
    archivedCount = 0;
    const out = [];
    const siteCase = buildCaseMap(locOptions('Site*'));
    locData.rows.forEach(row => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') { archivedCount++; return; }
      const r = new Array(LOC_HEADERS.length).fill('');
      const sid = locGet(row, 'site_id');
      // An explicit Site mapping wins; otherwise fall back to the name the
      // sites export gives for this site_id.
      const siteName = siteMap[sid] ||
        (sitesData && sitesData.byId.has(sid) ? sitesData.byId.get(sid) : '');
      r[L_SITE] = snapToOption(siteName, siteCase);
      r[L_NAME] = locGet(row, 'display_name');
      r[L_ALT] = locGet(row, 'location_id');
      r[L_TYPE] = typeMap[locGet(row, 'location_type_id') || '(blank)'] || '';
      r[L_CROP] = cropMap[locGet(row, 'crop_id') || '(blank)'] || '';
      r[6] = numOrBlank(locGet(row, 'hectares'), HECTARES_TO_ACRES);   // Acreage
      r[7] = numOrBlank(locGet(row, 'length_meters'), METERS_TO_FEET); // Length
      r[8] = numOrBlank(locGet(row, 'quantity'));                      // Plant Count
      r[9] = normDate(locGet(row, 'planted_date'));                    // Start Date*
      r[26] = normDate(locGet(row, 'planted_date'));                   // Planted Date
      // Legacy packs a cost code into info as "CC2=<code>" — keep the code,
      // drop the constant key so Custom Data 1 holds the code alone.
      r[38] = locGet(row, 'info').replace(/^CC\d*\s*=\s*/i, '');
      r[39] = locGet(row, 'description');
      r[40] = locGet(row, 'plant_name');
      out.push(r);
      locSrc.push(row);
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
    locData.rows.forEach(row => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
      const id = locGet(row, 'site_id');
      if (id) s.add(id);
    });
    return s;
  }

  function buildSiteRowsRaw() {
    if (!sitesData) return null;
    siteSrc = [];
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
      const name = g(row, nameI);
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
      if (migratedSites && migratedSites.has(name.toUpperCase())) { alreadyMigratedCount++; return; }
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
    });

    for (let ci = 0; ci < SITE_HEADERS.length; ci++) {
      const opts = siteOptions(SITE_HEADERS[ci]);
      if (!opts.length) continue;
      const cm = buildCaseMap(opts);
      out.forEach(r => { if (r[ci]) r[ci] = snapToOption(r[ci], cm); });
    }
    return out;
  }

  function applyFills(rows, fills, removed) {
    if (!rows) return;
    Object.keys(fills).forEach(k => {
      const ci = +k, { val, mode } = fills[k];
      rows.forEach((r, i) => {
        if (removed.has(i)) return;
        if (mode === 'blank') { if (!r[ci]) r[ci] = val; }
        else r[ci] = val;
      });
    });
  }

  function rebuild() {
    locRows = buildLocRowsRaw();
    siteRows = buildSiteRowsRaw();
    applyFills(locRows, locFills, locRemoved);
    applyFills(siteRows, siteFills, siteRemoved);
  }

  function ready() {
    return !!(locData && sitesData && locTpl);
  }

  function runMigrate() {
    if (!ready()) {
      alert('Upload the Legacy locations CSV, the sites export, and the 3.0 Locations template first.');
      return;
    }
    locFills = {}; siteFills = {};
    locRemoved = new Set(); siteRemoved = new Set();
    completedSites = new Set();
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
    $('lm-run').disabled = !ok;
    if (ok) runMigrate();
  }

  // Fill any crop_id the preset didn't cover, using the lookup CSV (if given)
  // or the inferred suggestion. Never overwrites an existing choice.
  function autoFillCrops() {
    collectIds('crop').forEach((info, cid) => {
      if (cropMap[cid]) return;
      const s = suggestionFor('crop', cid, info);
      if (s) cropMap[cid] = s.value;
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
      const s = suggestionFor('type', id, info);
      typeMap[id] = s ? s.value : fallback;
    });
  }

  function renderAll() {
    renderPresetBanner();
    renderSiteMap();
    renderIdMap('type');
    renderIdMap('crop');
    renderSitesToCreate();
    renderBulkEdit();
    renderPreview();
    updateSummary();
  }

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
    locData.rows.forEach(row => {
      if (locGet(row, 'is_archived').toLowerCase() === 'true') return;
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
    const opts = locOptions('Site*');
    const cm = buildCaseMap(opts);
    const unresolved = [...ids.entries()].filter(([id, info]) => {
      const v = siteMap[id] || info.legacyName;
      return !v || (opts.length && !cm.has(String(v).toUpperCase().trim()));
    }).length;
    const title = $('lm-sitemap-title');
    if (title) title.textContent = 'Site Mapping (' + ids.size + ' legacy site' +
      (ids.size === 1 ? '' : 's') +
      (unresolved ? ' · ' + unresolved + ' not in the template' : ' · all resolved') + ')';

    let html = '<thead><tr><th>site_id</th><th>Locations</th><th>Legacy name</th>' +
      '<th>Status</th><th>3.0 Site*</th><th>Manual override</th><th></th></tr></thead><tbody>';
    [...ids.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([id, info]) => {
      const cur = siteMap[id] || '';
      const eff = cur || info.legacyName;
      const inTpl = eff && (!opts.length || cm.has(String(eff).toUpperCase().trim()));
      const status = !eff
        ? '<span style="color:#dc2626;font-weight:600;">&#9888; no name for this site_id</span>'
        : inTpl
          ? '<span style="color:#15803d;font-weight:600;">&#10003; in template</span>'
          : '<span style="color:#b45309;font-weight:600;">&#9888; not in template &mdash; will be rejected</span>';
      const optsHtml = opts.length
        ? '<option value="">&mdash; keep legacy name &mdash;</option>' +
          opts.map(o => '<option value="' + escHtml(o) + '"' + (o === cur ? ' selected' : '') +
            '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(template has no Site dropdown)</option>';
      html += '<tr>' +
        '<td><b>' + escHtml(id) + '</b></td>' +
        '<td>' + info.count + '</td>' +
        '<td title="' + escHtml(info.legacyName) + '">' +
          escHtml(info.legacyName.length > 44 ? info.legacyName.slice(0, 44) + '…' : (info.legacyName || '—')) +
          '<div class="text-muted small">' + escHtml(info.samples.join(', ')) + '</div></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="lm-site-pick input-field" data-id="' + escHtml(id) + '" style="min-width:200px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="lm-site-override input-field" data-id="' + escHtml(id) + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm lm-site-apply" data-id="' + escHtml(id) + '">Apply</button></td>' +
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
        rebuild(); renderAll();
      });
    });
    tbl.querySelectorAll('.lm-site-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        if (e.target.value) { siteMap[e.target.dataset.id] = e.target.value; rebuild(); renderAll(); }
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
    const opts = kind === 'crop' ? locOptions('Crop & Variety*') : locOptions('Location Type*');
    const label = kind === 'crop' ? 'Crop &amp; Variety' : 'Location Type';
    const unmapped = [...ids.keys()].filter(id => !map[id]).length;
    const title = $('lm-' + kind + '-title');
    if (title) title.textContent = (kind === 'crop' ? 'Crop' : 'Location Type') +
      ' Mapping (' + ids.size + ' legacy ID' + (ids.size === 1 ? '' : 's') +
      (unmapped ? ' · ' + unmapped + ' unmapped' : ' · all mapped') + ')';

    let html = '<thead><tr><th>Legacy ID</th><th>Rows</th><th>Sample locations</th>' +
      '<th>Suggestion</th><th>3.0 ' + label + '</th><th>Manual override</th><th></th></tr></thead><tbody>';
    [...ids.entries()].sort((a, b) => b[1].count - a[1].count).forEach(([id, info]) => {
      const cur = map[id] || '';
      const sug = cur ? null : suggestionFor(kind, id, info);
      const sugHtml = cur
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
        : '<option value="">(template has no ' + label + ' dropdown)</option>';
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
        if (map[id]) return;
        const s = suggestionFor(kind, id, info);
        if (s) { map[id] = s.value; n++; }
      });
      if (!n) { alert('No unmapped IDs have a suggestion to accept.'); return; }
      rebuild(); renderAll();
    };
  }

  function setIdMapping(kind, id, val) {
    if (!val) return;
    (kind === 'crop' ? cropMap : typeMap)[id] = val;
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

  // ─── Sites to create in 3.0 ───
  function siteValid(v) {
    if (!v) return false;
    const opts = locOptions('Site*');
    if (!opts.length) return true; // can't validate without a dropdown
    return buildCaseMap(opts).has(String(v).toUpperCase().trim());
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

  function applyColumnFill(idx, val, mode) {
    const rows = curRows(), removed = curRemoved();
    curFills()[idx] = { val, mode: mode === 'blank' ? 'blank' : 'all' };
    rows.forEach((r, ri) => {
      if (removed.has(ri)) return;
      if (mode === 'blank') { if (!r[idx]) r[idx] = val; }
      else r[idx] = val;
    });
    renderBulkEdit();
    if (domain === 'locations' && idx === L_SITE) renderSitesToCreate();
    renderPreview();
    updateSummary();
  }

  function clearColumnFill(idx) {
    const rows = curRows(), removed = curRemoved();
    delete curFills()[idx];
    rows.forEach((r, ri) => { if (!removed.has(ri)) r[idx] = ''; });
    renderBulkEdit();
    if (domain === 'locations' && idx === L_SITE) renderSitesToCreate();
    renderPreview();
    updateSummary();
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
    let siteTouched = false;
    selCells.forEach(k => {
      const i = k.indexOf(':');
      const ri = +k.slice(0, i), ci = +k.slice(i + 1);
      if (!rows[ri]) return;
      rows[ri][ci] = val;
      if (domain === 'locations' && ci === L_SITE) siteTouched = true;
    });
    renderPreview();
    if (siteTouched) renderSitesToCreate();
    updateSummary();
  }

  // ─── Preview ───
  function renderPreview() {
    const rows = curRows();
    const sec = $('lm-section-preview');
    const tbl = $('lm-preview-table');
    if (!sec || !tbl || !rows) return;
    sec.style.display = '';
    const headers = curHeaders();
    const removed = curRemoved();

    document.querySelectorAll('.lm-domain-pill').forEach(b =>
      b.classList.toggle('ts-domain-active', b.dataset.domain === domain));

    let html = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    headers.forEach(h => {
      const req = /\*$/.test(h);
      html += '<th' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</th>';
    });
    html += '</tr></thead><tbody>';

    const reqIdx = headers.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    const isProblem = ri => reqIdx.some(i => !rows[ri][i]);
    const vis = rows.map((_, i) => i).filter(i => !removed.has(i));
    // Always show every row missing a required cell; cap the complete ones.
    const CLEAN_CAP = 50;
    let cleanShown = 0;
    const show = [];
    vis.forEach(ri => {
      if (isProblem(ri)) show.push(ri);
      else if (cleanShown < CLEAN_CAP) { show.push(ri); cleanShown++; }
    });
    const problemCount = vis.reduce((n, ri) => n + (isProblem(ri) ? 1 : 0), 0);
    const cleanTotal = vis.length - problemCount;
    previewOrder = show.slice();
    const visSet = new Set(vis);
    selCells.forEach(k => { if (!visSet.has(+k.slice(0, k.indexOf(':')))) selCells.delete(k); });

    show.forEach(ri => {
      const row = rows[ri];
      html += '<tr><td style="text-align:center;padding:0;">' +
        '<button class="lm-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      headers.forEach((h, i) => {
        const empty = !row[i];
        const cs = (/\*$/.test(h) && empty) ? ' style="background:#fee2e2;color:#7f1d1d;"' : '';
        html += '<td class="lm-cell" contenteditable="true" spellcheck="false" ' +
          'data-ri="' + ri + '" data-ci="' + i + '"' + cs + '>' + escHtml(row[i]) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;

    const hint = sec.querySelector('.cmp-sites-hint');
    let ht = 'Editable &mdash; click a cell, type, then click away (or press Enter) to save. ' +
      '<b>Shift+click</b> selects a rectangular range; <b>Ctrl/Cmd+click</b> toggles cells. ' +
      'Empty required (*) cells are red. Click <b>&times;</b> to drop a row from the export.';
    if (problemCount) {
      ht = '<b style="color:#dc2626;">' + problemCount + ' row' + (problemCount === 1 ? '' : 's') +
        ' missing a required (*) cell</b> &mdash; all shown below' +
        (cleanTotal > cleanShown ? ', plus first ' + cleanShown + ' of ' + cleanTotal + ' complete rows' : '') +
        '. ' + ht;
    } else if (cleanTotal > cleanShown) {
      ht = 'All required cells filled. Showing first ' + cleanShown + ' of ' + cleanTotal + ' rows. ' + ht;
    }
    if (removed.size) ht += ' &nbsp; <button class="btn btn-ghost btn-sm" id="lm-restore-rows">Restore ' +
      removed.size + ' removed row' + (removed.size === 1 ? '' : 's') + '</button>';
    if (hint) hint.innerHTML = ht;

    tbl.querySelectorAll('.lm-row-remove').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        curRemoved().add(+e.currentTarget.dataset.ri);
        renderPreview(); renderSitesToCreate(); updateSummary();
      });
    });
    const rb = $('lm-restore-rows');
    if (rb) rb.addEventListener('click', () => {
      curRemoved().clear();
      renderPreview(); renderSitesToCreate(); updateSummary();
    });
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
  }

  function updateExportButton() {
    const btn = $('lm-export');
    if (!btn) return;
    const locVis = locRows ? locRows.filter((_, i) => !locRemoved.has(i)).length : 0;
    btn.disabled = !(ready() && locVis);
  }

  // ═══ XLSX writing at the zip level ═══════════════════════════════════
  // An .xlsx is a zip of XML parts. Reading one into SheetJS and writing it
  // back re-generates every part from SheetJS's own model, which drops
  // anything the model doesn't cover — on this template that meant all 8
  // <dataValidations> (the dropdowns) and most of styles.xml.
  //
  // Instead we treat the template as the source of truth and patch it: every
  // entry is copied through with its ORIGINAL compressed bytes (never
  // re-encoded), and only the DATA ENTRY sheet's <sheetData> is rewritten.
  // Cells are written as inline strings so sharedStrings.xml is untouched too,
  // and xml:space="preserve" keeps any trailing "ghost space" on a name.

  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Parse the central directory. Keeps each entry's raw (still-compressed)
  // bytes so untouched parts can be re-emitted verbatim.
  function zipRead(buffer) {
    const u8 = new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    const floor = Math.max(0, u8.length - 22 - 65535);
    for (let i = u8.length - 22; i >= floor; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Template is not a valid .xlsx (no zip end record).');
    const count = dv.getUint16(eocd + 10, true);
    let off = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(off, true) !== 0x02014b50) throw new Error('Corrupt zip central directory.');
      const method = dv.getUint16(off + 10, true);
      const mtime = dv.getUint16(off + 12, true);
      const mdate = dv.getUint16(off + 14, true);
      const crc = dv.getUint32(off + 16, true);
      const csize = dv.getUint32(off + 20, true);
      const usize = dv.getUint32(off + 24, true);
      const nlen = dv.getUint16(off + 28, true);
      const elen = dv.getUint16(off + 30, true);
      const clen = dv.getUint16(off + 32, true);
      const lho = dv.getUint32(off + 42, true);
      const name = dec.decode(u8.subarray(off + 46, off + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true);
      const lelen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lelen;
      entries.push({ name, method, crc, usize, mtime, mdate, data: u8.subarray(start, start + csize) });
      off += 46 + nlen + elen + clen;
    }
    return entries;
  }

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  async function entryText(e) {
    return new TextDecoder().decode(e.method === 0 ? e.data : await inflateRaw(e.data));
  }
  // Deflate the replaced part when the platform offers CompressionStream
  // (a 220-row sheet is ~150 KB stored vs ~15 KB deflated); fall back to
  // stored (method 0), which is equally valid, if it is missing or fails.
  async function replaceEntry(e, text) {
    const data = new TextEncoder().encode(text);
    const crc = crc32(data);
    let out = data, method = 0;
    if (typeof CompressionStream === 'function') {
      try {
        const s = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
        const z = new Uint8Array(await new Response(s).arrayBuffer());
        if (z.length < data.length) { out = z; method = 8; }
      } catch (err) { /* keep stored */ }
    }
    return { name: e.name, method, crc, usize: data.length,
             mtime: e.mtime, mdate: e.mdate, data: out };
  }

  function zipWrite(entries) {
    const enc = new TextEncoder();
    const parts = [], centrals = [];
    let offset = 0;
    entries.forEach(e => {
      const nb = enc.encode(e.name);
      const lh = new Uint8Array(30 + nb.length);
      const l = new DataView(lh.buffer);
      l.setUint32(0, 0x04034b50, true);
      l.setUint16(4, 20, true);
      l.setUint16(6, 0, true);
      l.setUint16(8, e.method, true);
      l.setUint16(10, e.mtime, true);
      l.setUint16(12, e.mdate, true);
      l.setUint32(14, e.crc, true);
      l.setUint32(18, e.data.length, true);
      l.setUint32(22, e.usize, true);
      l.setUint16(26, nb.length, true);
      l.setUint16(28, 0, true);
      lh.set(nb, 30);
      parts.push(lh, e.data);

      const ch = new Uint8Array(46 + nb.length);
      const c = new DataView(ch.buffer);
      c.setUint32(0, 0x02014b50, true);
      c.setUint16(4, 20, true);
      c.setUint16(6, 20, true);
      c.setUint16(8, 0, true);
      c.setUint16(10, e.method, true);
      c.setUint16(12, e.mtime, true);
      c.setUint16(14, e.mdate, true);
      c.setUint32(16, e.crc, true);
      c.setUint32(20, e.data.length, true);
      c.setUint32(24, e.usize, true);
      c.setUint16(28, nb.length, true);
      c.setUint32(42, offset, true);
      ch.set(nb, 46);
      centrals.push(ch);
      offset += lh.length + e.data.length;
    });
    let cdSize = 0;
    centrals.forEach(c => { cdSize += c.length; });
    const eocd = new Uint8Array(22);
    const e2 = new DataView(eocd.buffer);
    e2.setUint32(0, 0x06054b50, true);
    e2.setUint16(8, entries.length, true);
    e2.setUint16(10, entries.length, true);
    e2.setUint32(12, cdSize, true);
    e2.setUint32(16, offset, true);
    return new Blob(parts.concat(centrals, [eocd]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const xmlEsc = s => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // Strip characters XML 1.0 forbids outright — Legacy free-text fields
    // occasionally carry stray control bytes that would corrupt the sheet.
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

  function colRef(n) {
    let s = '';
    n += 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  // Locate the DATA ENTRY worksheet part by walking workbook.xml → rels.
  async function findSheetPart(entries) {
    const byName = {};
    entries.forEach(e => { byName[e.name] = e; });
    const wbXml = await entryText(byName['xl/workbook.xml']);
    const relsXml = await entryText(byName['xl/_rels/workbook.xml.rels']);
    const sheets = [...wbXml.matchAll(/<sheet\b[^>]*\/?>/g)].map(m => m[0]);
    let rid = null;
    for (const s of sheets) {
      const nm = (s.match(/name="([^"]*)"/) || [])[1] || '';
      if (/data.?entry/i.test(nm)) { rid = (s.match(/r:id="([^"]+)"/) || [])[1]; break; }
    }
    if (!rid && sheets.length) rid = (sheets[0].match(/r:id="([^"]+)"/) || [])[1];
    if (!rid) throw new Error('Could not find a DATA ENTRY sheet in the template.');
    const rel = [...relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)]
      .map(m => m[0]).find(r => (r.match(/Id="([^"]+)"/) || [])[1] === rid);
    if (!rel) throw new Error('Template worksheet relationship is missing.');
    let target = (rel.match(/Target="([^"]+)"/) || [])[1] || '';
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const path = 'xl/' + target;
    if (!byName[path]) throw new Error('Template worksheet part not found: ' + path);
    return byName[path];
  }

  // Build <sheetData>: the template's own header row verbatim, then our rows
  // as inline strings.
  function buildSheetData(headerRowXml, rows, colCount) {
    const out = ['<sheetData>', headerRowXml];
    rows.forEach((row, ri) => {
      const r = ri + 2; // row 1 is the header
      let cells = '';
      for (let c = 0; c < colCount; c++) {
        const v = row[c];
        if (v == null || v === '') continue;
        cells += '<c r="' + colRef(c) + r + '" t="inlineStr"><is><t xml:space="preserve">' +
          xmlEsc(v) + '</t></is></c>';
      }
      if (cells) out.push('<row r="' + r + '">' + cells + '</row>');
    });
    out.push('</sheetData>');
    return out.join('');
  }

  async function writeWorkbook(tpl, headers, rows, fileName) {
    const entries = zipRead(tpl.rawBuffer);
    const sheet = await findSheetPart(entries);
    let xml = await entryText(sheet);

    const sd = xml.match(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>|<sheetData\b[^>]*\/>/);
    if (!sd) throw new Error('Template worksheet has no <sheetData> element.');
    const headerRow = (sd[0].match(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/) || [''])[0];
    const colCount = headers.length;
    const next = buildSheetData(headerRow, rows, colCount);
    xml = xml.slice(0, sd.index) + next + xml.slice(sd.index + sd[0].length);
    // Keep <dimension> honest; Excel repairs the file if it disagrees badly.
    xml = xml.replace(/<dimension\b[^>]*\/>|<dimension\b[^>]*>[\s\S]*?<\/dimension>/,
      '<dimension ref="A1:' + colRef(colCount - 1) + Math.max(1, rows.length + 1) + '"></dimension>');

    const patchedSheet = await replaceEntry(sheet, xml);
    const patched = entries.map(e => e.name === sheet.name ? patchedSheet : e);
    const blob = zipWrite(patched);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return true;
  }

  function outName(tpl, fallback, suffix) {
    const orig = (tpl && tpl.fileName ? tpl.fileName : fallback).trim();
    const dot = orig.lastIndexOf('.');
    const base = dot > 0 ? orig.substring(0, dot) : orig;
    const ext = dot > 0 ? orig.substring(dot) : '.xlsx';
    return base + ' — filled' + (suffix || '') + ext;
  }

  async function doExport() {
    if (!locRows || !locTpl) return;
    const locOut = locRows.filter((_, i) => !locRemoved.has(i)).map(r => r.slice());
    const siteOut = siteRows ? siteRows.filter((_, i) => !siteRemoved.has(i)).map(r => r.slice()) : [];
    if (!locOut.length) { alert('No locations to export.'); return; }

    // HARD BLOCK. A required column that is empty makes the whole upload
    // fail, so this is an error, not a warning to click through — the previous
    // dismissible confirm let a file out with 220 blank Location Type* and
    // Crop & Variety* cells.
    const gaps = requiredGaps(locRows, LOC_HEADERS, locRemoved);
    if (gaps.length) {
      alert('Cannot export — required columns are incomplete:\n\n' +
        gaps.map(g => '  • ' + g.header + ': ' + g.rows + ' row' +
          (g.rows === 1 ? '' : 's') + ' empty').join('\n') +
        '\n\nPickTrace rejects the upload outright when a required (*) column is blank. ' +
        'Fill them in the mapping panels or Bulk Edit, then export again.');
      return;
    }
    const siteGaps = (siteTpl && siteOut.length)
      ? requiredGaps(siteRows, SITE_HEADERS, siteRemoved) : [];
    if (siteGaps.length) {
      alert('Cannot export — the Sites sheet has incomplete required columns:\n\n' +
        siteGaps.map(g => '  • ' + g.header + ': ' + g.rows + ' row' +
          (g.rows === 1 ? '' : 's') + ' empty').join('\n') +
        '\n\nLegacy carries no Site Type, Employer or address parts, so fill these ' +
        'from Bulk Edit (switch the preview to Sites first).');
      return;
    }

    // Sites not present in the template's Site list stay a warning: they may
    // legitimately be created in PickTrace before the upload runs.
    const pending = pendingSitesToCreate();
    if (pending.length) {
      const names = pending.sort((a, b) => b[1].count - a[1].count)
        .map(([n, info]) => '  • ' + n + ' (' + info.count + ' locations)').join('\n');
      if (!window.confirm(
        pending.length + ' site' + (pending.length === 1 ? '' : 's') +
        ' referenced by these locations are not in the template\'s Site list:\n\n' + names +
        '\n\nMap them in the Site Mapping panel, or create them in PickTrace and re-download ' +
        'the Locations template — otherwise the upload is rejected with "Site not found".' +
        '\n\nExport anyway?')) return;
    }

    let limit = parseInt($('lm-batch').value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;

    try {
      // Sites first — they must exist in 3.0 before the locations that use them.
      if (siteTpl && siteOut.length) {
        const chunks = chunkArr(siteOut, limit);
        for (let i = 0; i < chunks.length; i++) {
          const suffix = chunks.length === 1 ? '' : ' (' + (i + 1) + ' of ' + chunks.length + ')';
          await writeWorkbook(siteTpl, SITE_HEADERS, chunks[i],
            outName(siteTpl, '3.0_Sites_bulk_create.xlsx', suffix));
          await sleep(300);
        }
      }
      const chunks = chunkArr(locOut, limit);
      for (let i = 0; i < chunks.length; i++) {
        const suffix = chunks.length === 1 ? '' : ' (' + (i + 1) + ' of ' + chunks.length + ')';
        await writeWorkbook(locTpl, LOC_HEADERS, chunks[i],
          outName(locTpl, '3.0_Locations_bulk_create.xlsx', suffix));
        if (i < chunks.length - 1) await sleep(300);
      }
    } catch (err) {
      alert('Export failed: ' + (err && err.message ? err.message : err) +
        '\n\nNothing was written. The template file may be an unsupported .xls ' +
        '(only .xlsx can be patched without losing its dropdowns).');
    }
  }

  // ─── Reset ───
  function reset() {
    locData = null; sitesData = null; cropLookup = null; typeLookup = null;
    migratedSites = null; locTpl = null; siteTpl = null;
    cropMap = {}; typeMap = {}; siteMap = {}; presetApplied = false;
    locRows = null; locSrc = null; siteRows = null; siteSrc = null;
    locFills = {}; siteFills = {};
    locRemoved = new Set(); siteRemoved = new Set();
    archivedCount = 0; alreadyMigratedCount = 0;
    completedSites = new Set();
    selCells = new Set(); selAnchor = null; previewOrder = [];
    domain = 'locations';
    { const b = $('lm-sel-bar'); if (b) b.style.display = 'none'; }
    [['lm-loc', 'No file selected'], ['lm-sites', 'No file selected'],
     ['lm-croplk', 'No file selected'], ['lm-typelk', 'No file selected'],
     ['lm-migrated', 'No file selected'], ['lm-loctpl', 'No file selected'],
     ['lm-sitetpl', 'No file selected']].forEach(([p, txt]) => {
      const n = $(p + '-name'); if (n) n.textContent = txt;
      const m = $(p + '-meta'); if (m) m.textContent = '';
      const f = $(p + '-file'); if (f) f.value = '';
    });
    ['lm-section-sitemap', 'lm-section-type', 'lm-section-crop', 'lm-section-create',
     'lm-section-bulk', 'lm-section-preview'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    { const b = $('lm-preset-banner'); if (b) b.style.display = 'none'; }
    $('lm-summary').style.display = 'none';
    $('lm-empty').style.display = '';
    $('lm-run').disabled = true;
    $('lm-export').disabled = true;
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
    const req = /\*$/.test(curHeaders()[ci]);
    td.style.background = (req && !val) ? '#fee2e2' : '';
    td.style.color = (req && !val) ? '#7f1d1d' : '';
    if (domain === 'locations' && ci === L_SITE) renderSitesToCreate();
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

    $('lm-run').addEventListener('click', runMigrate);
    $('lm-export').addEventListener('click', () => { doExport(); });
    $('lm-reset').addEventListener('click', reset);

    const copyBtn = $('lm-create-copy');
    if (copyBtn) copyBtn.addEventListener('click', () => {
      const names = [...getSitesToCreate().keys()];
      if (!names.length) { alert('No sites to create.'); return; }
      navigator.clipboard.writeText(names.join('\n')).then(
        () => alert('Copied ' + names.length + ' site name' + (names.length === 1 ? '' : 's') + '.'),
        () => alert('Copy failed.'));
    });

    // Locations / Sites preview toggle.
    document.querySelectorAll('.lm-domain-pill').forEach(btn => {
      btn.addEventListener('click', () => {
        domain = btn.dataset.domain;
        clearSelection();
        renderBulkEdit();
        renderPreview();
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
