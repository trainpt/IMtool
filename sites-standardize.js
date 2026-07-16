// ═══════════════════════════════════════════════════════════════════════
// Sites Standardize — formats a source Excel/CSV (or an implementation
// template's "Sites" tab) into a personalized PickTrace bulk SITES Create
// template. Sibling to template-standardize.js (which handles Locations);
// both live under the Template Standardize page's Locations/Sites toggle.
//
// Sites schema (10 cols): Name*, Alt ID, Site Type*, Employer*, Address1*,
// Address2, City*, State*, Zip*, Country*. The source usually packs the whole
// address into ONE "Address" cell, so this module parses it into
// Address1/City/State/Zip/Country (City is a best-effort guess, flagged amber).
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');
  const normLoose = h => norm(h).replace(/[^a-z0-9]/g, '');
  const locNorm = s => String(s == null ? '' : s).toUpperCase().trim().replace(/\s+/g, ' ');

  const SITE_HEADERS = ['Name*', 'Alt ID', 'Site Type*', 'Employer*', 'Address1*',
    'Address2', 'City*', 'State*', 'Zip*', 'Country*'];

  // Source-column aliases → bulk Sites column.
  const SITE_ALIASES = {
    'name*':      ['name', 'site name', 'site', 'sites', 'ranch', 'location name'],
    'alt id':     ['alt id', 'altid', 'code'],
    'site type*': ['site type', 'type'],
    'employer*':  ['employer', 'grower', 'company', 'client'],
    'address1*':  ['address1', 'address 1', 'address', 'street', 'address line 1', 'addr'],
    'address2':   ['address2', 'address 2', 'address line 2', 'suite', 'unit'],
    'city*':      ['city', 'town'],
    'state*':     ['state', 'province'],
    'zip*':       ['zip', 'zipcode', 'zip code', 'postal', 'postal code', 'postal_code'],
    'country*':   ['country', 'nation']
  };

  const US_STATES = new Set(['AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','DC']);
  const CA_PROVINCES = new Set(['BC','QC','ON','AB','MB','NB','NL','NS','NT','NU','PE','SK','YT']);
  const STREET_SUFFIX = new Set(['st','street','rd','road','ave','av','avenue','blvd','boulevard','dr','drive','ln','lane','way','hwy','highway','ct','court','pl','place','ter','terrace','cir','circle','pkwy','parkway','route','rte','trl','trail','loop','row','sq','square','pike','fwy','expy','plaza','box','apt','ste','spc']);

  // ─── State ───
  let srcData = null;             // { headers, rows, fileName, sheetName }
  let tplData = null;             // { headers, dropdowns, fileName, rawBuffer }
  let mapping = {};               // bulkColIdx → srcColIdx (-1 = unmapped)
  let formattedRows = null;       // [[...]] canonical mapped + filled rows
  let manualFills = {};           // bulkColIdx → value re-applied after rebuilds
  let removedRows = new Set();    // formattedRows indexes dropped from preview + export
  let existingData = null;        // { keys:Set<nameKey>, byName:Map, fileName, count }
  let existingRowIdxs = new Set();
  let nameOverrides = {};         // rowIdx → renamed Name* (collision fix, sticky)
  let cellOverrides = {};         // "row|col" → value (sticky preview edits, all cols except Name*)
  let smartFixMap = {};           // "colIdx||UPPERVALUE" → canonical dropdown value (applied Smart Fixes)
  let previewIssuesOnly = false;  // preview toggle: show only rows with an errored cell
  let guessedCityRows = new Set();// rowIdxs whose City* was machine-guessed from a combined address
  let collisionKeys = new Set();
  let existingTakenKeys = new Set();
  let initialized = false;

  function excluded(i) { return removedRows.has(i) || existingRowIdxs.has(i); }
  function nameKeyOf(name) { return locNorm(name); }

  // ─── Address parsing ───
  // "22759 S. MERCEY SPRINGS RD. LOS BANOS, CA 93635" →
  //   { address1:'22759 S. MERCEY SPRINGS RD.', city:'LOS BANOS',
  //     state:'CA', zip:'93635', country:'US', cityGuessed:true }
  function parseAddress(raw) {
    const out = { address1: '', city: '', state: '', zip: '', country: '', cityGuessed: false };
    let s = String(raw == null ? '' : raw).trim().replace(/\s+/g, ' ');
    if (!s) return out;
    // ZIP — US 5(-4) or Canadian A1A 1A1.
    let m = s.match(/[,\s]\s*(\d{5}(?:-\d{4})?)\s*$/);
    if (m) { out.zip = m[1]; s = s.slice(0, m.index).trim(); }
    else {
      m = s.match(/[,\s]\s*([A-Za-z]\d[A-Za-z]\s*\d[A-Za-z]\d)\s*$/);
      if (m) { out.zip = m[1].toUpperCase().replace(/\s+/g, ' '); s = s.slice(0, m.index).trim(); }
    }
    // STATE — trailing 2-letter, accepted when comma-separated OR a known code.
    m = s.match(/(,?)\s*([A-Za-z]{2})\s*$/);
    if (m) {
      const st = m[2].toUpperCase();
      if (m[1] === ',' || US_STATES.has(st) || CA_PROVINCES.has(st)) {
        out.state = st;
        s = s.slice(0, s.length - m[0].length).trim();
      }
    }
    s = s.replace(/,\s*$/, '').trim();
    out.country = out.state ? (CA_PROVINCES.has(out.state) ? 'CA' : 'US') : '';
    // Remaining = street + city. A comma cleanly separates them; otherwise guess
    // the city as everything after the last street-suffix token.
    if (s.indexOf(',') >= 0) {
      const idx = s.lastIndexOf(',');
      out.address1 = s.slice(0, idx).trim();
      out.city = s.slice(idx + 1).trim();
    } else {
      const toks = s.split(' ');
      let cut = -1;
      for (let i = 0; i < toks.length; i++) {
        const t = toks[i].replace(/\./g, '').toLowerCase();
        if (STREET_SUFFIX.has(t)) cut = i;
      }
      if (cut >= 0 && cut < toks.length - 1) {
        out.address1 = toks.slice(0, cut + 1).join(' ');
        out.city = toks.slice(cut + 1).join(' ');
        out.cityGuessed = true;
      } else {
        out.address1 = s;
      }
    }
    return out;
  }

  // ─── Header-row detection (mirrors template-standardize) ───
  const TEMPLATE_KEYWORDS = ['site', 'name', 'type', 'employer', 'grower', 'address',
    'city', 'state', 'zip', 'country', 'code', 'group', 'alt id'];
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
      return s.startsWith('required.') || s.startsWith('please ') || s.startsWith('optional.') ||
             s.startsWith('this field') || s.startsWith('this tab') || s.includes('used to collect');
    });
  }
  function sliceSheetAoa(aoa) {
    const hIdx = detectHeaderRow(aoa);
    const headers = (aoa[hIdx] || []).map(h => String(h == null ? '' : h).trim());
    const rows = aoa.slice(hIdx + 1)
      .filter(r => !isInstructionRow(r))
      .map(r => r.map(c => c == null ? '' : String(c).trim()));
    return { headers, rows, headerRow: hIdx };
  }

  // ─── File parsing ───
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
            return { name: n, headers: sliced.headers, rows: sliced.rows };
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({ headers: s.headers, rows: s.rows, fileName: file.name, sheetName: s.name });
          if (sheets.length === 1) return finish(sheets[0]);
          pickSheet(sheets, file.name).then(finish).catch(reject);
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
  }
  function pickSheet(sheets, fileName) {
    return new Promise((resolve, reject) => {
      const prior = document.getElementById('tss-sheet-picker');
      if (prior) prior.remove();
      const overlay = document.createElement('div');
      overlay.id = 'tss-sheet-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner';
      inner.innerHTML = '<h3>Select sheet</h3><p>Pick the sheet with the <b>Sites</b> data in <code>' +
        escHtml(fileName) + '</code>:</p><div class="cmp-org-list" id="tss-sheet-picker-list"></div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="tss-sheet-picker-cancel">Cancel</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
      const list = inner.querySelector('#tss-sheet-picker-list');
      sheets.forEach(s => {
        const item = document.createElement('div');
        item.className = 'cmp-org-list-item';
        item.innerHTML = '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
          '<span class="cmp-org-counts">' + (s.rows ? s.rows.length : 0) + ' rows · ' + (s.headers ? s.headers.length : 0) + ' cols</span>';
        item.addEventListener('click', () => { overlay.remove(); resolve(s); });
        list.appendChild(item);
      });
      const cancel = () => { overlay.remove(); reject(new Error('cancelled')); };
      inner.querySelector('#tss-sheet-picker-cancel').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
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
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[dropName], { header: 1, defval: '' });
      if (raw.length >= 2) {
        raw[0].map(h => String(h || '').trim()).forEach((h, ci) => {
          if (!h) return;
          const k = norm(h);
          const set = dropdowns.get(k) || new Set();
          raw.slice(1).forEach(r => { const v = String(r[ci] != null ? r[ci] : '').trim(); if (v) set.add(v); });
          if (set.size) dropdowns.set(k, set);
        });
      }
    }
    return { headers, dropdowns, fileName, rawBuffer: buf, sheetNames: wb.SheetNames };
  }

  // Is this template actually a Sites bulk template? (Name* + Site Type* + no Crop.)
  function looksLikeSitesTemplate(headers) {
    const set = new Set((headers || []).map(h => norm(h).replace(/\*$/, '')));
    return set.has('site type') && (set.has('employer') || set.has('address1') || set.has('city'));
  }

  // ─── Auto-mapping ───
  function autoMap(bulkHeaders, srcHeaders) {
    const srcIdx = {}, srcIdxLoose = {};
    srcHeaders.forEach((h, i) => {
      const k = norm(h), kL = normLoose(h);
      if (k && srcIdx[k] == null) srcIdx[k] = i;
      if (kL && srcIdxLoose[kL] == null) srcIdxLoose[kL] = i;
    });
    const out = {};
    bulkHeaders.forEach((bh, bi) => {
      const bn = norm(bh), bnPlain = bn.replace(/\*$/, '').trim();
      const bnLoose = normLoose(bh), bnLoosePlain = normLoose(bh.replace(/\*$/, ''));
      if (srcIdx[bn] != null) { out[bi] = srcIdx[bn]; return; }
      if (srcIdx[bnPlain] != null) { out[bi] = srcIdx[bnPlain]; return; }
      if (srcIdxLoose[bnLoose] != null) { out[bi] = srcIdxLoose[bnLoose]; return; }
      if (srcIdxLoose[bnLoosePlain] != null) { out[bi] = srcIdxLoose[bnLoosePlain]; return; }
      const aliases = SITE_ALIASES[bn] || SITE_ALIASES[bnPlain];
      if (aliases) {
        for (const a of aliases) {
          if (srcIdx[a] != null) { out[bi] = srcIdx[a]; return; }
          const aL = normLoose(a);
          if (srcIdxLoose[aL] != null) { out[bi] = srcIdxLoose[aL]; return; }
        }
      }
      // Substring fallback (skip short + *_id columns). City/State/Zip/Country
      // usually have no source column and stay -1 (filled by the address parser).
      for (const k of Object.keys(srcIdxLoose)) {
        if (!k || (k.endsWith('id') && /id$/.test(k) && k.length <= bnLoosePlain.length)) continue;
        // Don't let a numbered target ("address2") grab the un-numbered base
        // source column ("address") — that's what Address1* already maps to.
        if (/\d$/.test(bnLoosePlain) && bnLoosePlain.replace(/\d+$/, '') === k) continue;
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

  function dropdownValuesFor(bulkHeader) {
    if (!tplData || !tplData.dropdowns) return null;
    const key = norm(bulkHeader).replace(/\*+$/, '').trim();
    const set = tplData.dropdowns.get(key);
    return set && set.size ? [...set] : null;
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
    const byLoose = uniq(maps.loose.get(looseKey(value)));
    if (byLoose) return byLoose;
    const byPlural = uniq(maps.plural.get(pluralKey(value)));
    if (byPlural) return byPlural;
    return null;
  }
  function dropdownColIdxs() {
    return SITE_HEADERS.map((h, i) => dropdownValuesFor(h) ? i : -1).filter(i => i >= 0);
  }
  function computeSmartFixes() {
    if (!formattedRows || !tplData) return [];
    const out = [];
    dropdownColIdxs().forEach(ci => {
      const vals = dropdownValuesFor(SITE_HEADERS[ci]);
      const maps = buildMatchMaps(vals);
      const seen = new Map();
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        const val = r[ci];
        if (!val) return;
        const u = String(val).toUpperCase().trim();
        if (maps.canon.has(u)) return;
        if (smartFixMap[ci + '||' + u]) return;
        const sug = confidentMatch(val, maps);
        const suggestion = (sug && String(sug).toUpperCase().trim() !== u) ? sug : '';
        const e = seen.get(u) || { colIdx: ci, header: SITE_HEADERS[ci], from: val, count: 0, options: vals, suggestion };
        e.count++;
        seen.set(u, e);
      });
      seen.forEach(e => out.push(e));
    });
    out.sort((a, b) => (b.suggestion ? 1 : 0) - (a.suggestion ? 1 : 0) || b.count - a.count);
    return out;
  }
  function buildCaseMap(set) {
    const m = new Map();
    if (!set) return m;
    set.forEach(v => { const k = String(v).toUpperCase().trim(); if (k && !m.has(k)) m.set(k, v); });
    return m;
  }
  function inDropdown(map, value) {
    if (!value || !map.size) return false;
    return map.has(String(value).toUpperCase().trim());
  }

  // ─── Build formatted rows ───
  function buildFormattedRowsRaw() {
    if (!srcData || !tplData) return null;
    const a1 = SITE_HEADERS.indexOf('Address1*');
    const ci = SITE_HEADERS.indexOf('City*');
    const si = SITE_HEADERS.indexOf('State*');
    const zi = SITE_HEADERS.indexOf('Zip*');
    const coi = SITE_HEADERS.indexOf('Country*');
    guessedCityRows = new Set();
    return srcData.rows.map((srcRow, ri) => {
      const out = new Array(SITE_HEADERS.length).fill('');
      SITE_HEADERS.forEach((bh, bi) => {
        const si2 = mapping[bi];
        if (si2 != null && si2 >= 0 && srcRow[si2] != null) {
          const v = String(srcRow[si2]).trim();
          if (v) out[bi] = v;
        }
      });
      // Parse a combined address (only when City/State/Zip aren't already mapped
      // in full). Fills only the empty target cells so explicit columns win.
      if (a1 >= 0 && out[a1] && !(out[ci] && out[si] && out[zi])) {
        const p = parseAddress(out[a1]);
        if (p.address1) out[a1] = p.address1;
        if (ci >= 0 && !out[ci] && p.city) { out[ci] = p.city; if (p.cityGuessed) guessedCityRows.add(ri); }
        if (si >= 0 && !out[si] && p.state) out[si] = p.state;
        if (zi >= 0 && !out[zi] && p.zip) out[zi] = p.zip;
        if (coi >= 0 && !out[coi] && p.country) out[coi] = p.country;
      }
      return out;
    });
  }

  function rebuildFormattedRows() {
    formattedRows = buildFormattedRowsRaw();
    if (!formattedRows) return;
    Object.entries(manualFills).forEach(([idx, val]) => {
      const i = +idx;
      formattedRows.forEach(r => { if (!r[i]) r[i] = val; });
    });
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const cityIdx = SITE_HEADERS.indexOf('City*');
    // Sticky Name* renames first (collision resolutions).
    Object.keys(nameOverrides).forEach(k => {
      const ri = +k;
      if (formattedRows[ri] && nameOverrides[k]) formattedRows[ri][nameIdx] = nameOverrides[k];
    });
    // Apply accepted Smart Fixes (bulk snap of off-list values to the dropdown).
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
    // Sticky per-cell preview edits (all columns except Name*). A user edit to a
    // guessed City clears the "guessed" flag so it's no longer amber.
    Object.keys(cellOverrides).forEach(k => {
      const sep = k.indexOf('|');
      const ri = +k.slice(0, sep), ci = +k.slice(sep + 1);
      if (ci === nameIdx) return;
      if (formattedRows[ri]) formattedRows[ri][ci] = cellOverrides[k];
      if (ci === cityIdx) guessedCityRows.delete(ri);
    });

    // Cross-reference an existing PickTrace sites export: drop rows whose Name
    // already exists; and flag names already taken as collisions.
    existingRowIdxs = new Set();
    const withinCounts = new Map();
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!n) return;
      const k = nameKeyOf(n);
      withinCounts.set(k, (withinCounts.get(k) || 0) + 1);
    });
    if (existingData && existingData.keys && existingData.keys.size) {
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri)) return;
        const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
        if (!n) return;
        const k = nameKeyOf(n);
        if (existingData.keys.has(k) && (withinCounts.get(k) || 0) <= 1) existingRowIdxs.add(ri);
      });
    }
    // Collisions among rows that will actually be created.
    collisionKeys = new Set();
    existingTakenKeys = new Set();
    const counts = new Map();
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri) || existingRowIdxs.has(ri)) return;
      const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!n) return;
      const k = nameKeyOf(n);
      counts.set(k, (counts.get(k) || 0) + 1);
      if (existingData && existingData.keys && existingData.keys.has(k)) {
        collisionKeys.add(k); existingTakenKeys.add(k);
      }
    });
    counts.forEach((c, k) => { if (c >= 2) collisionKeys.add(k); });
  }

  function computeCollisions() {
    const out = new Map();
    if (!formattedRows || !collisionKeys.size) return out;
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri) || existingRowIdxs.has(ri)) return;
      const n = String(r[nameIdx] == null ? '' : r[nameIdx]).trim();
      if (!n) return;
      const k = nameKeyOf(n);
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

  // ─── Mapping panel ───
  function renderMapping() {
    if (!srcData || !tplData) return;
    const sec = $('tss-section-mapping');
    sec.style.display = '';
    const srcOpts = ['<option value="-1">— (leave empty) —</option>']
      .concat(srcData.headers.map((h, i) => '<option value="' + i + '">' + escHtml(h) + '</option>')).join('');
    let html = '<thead><tr><th>Template column</th><th>Source column</th><th>Sample value</th></tr></thead><tbody>';
    SITE_HEADERS.forEach((bh, bi) => {
      const required = /\*$/.test(bh);
      const sel = mapping[bi] != null ? mapping[bi] : -1;
      const sample = sel >= 0 && srcData.rows[0] ? String(srcData.rows[0][sel] || '').trim() : '';
      html += '<tr><td>' + (required ? '<b>' + escHtml(bh) + '</b>' : escHtml(bh)) + '</td>' +
        '<td><select class="tss-map-select input-field" data-bi="' + bi + '" style="min-width:220px;">' +
        srcOpts.replace('value="' + sel + '"', 'value="' + sel + '" selected') + '</select></td>' +
        '<td><span class="text-muted small">' + escHtml(sample.slice(0, 60)) + '</span></td></tr>';
    });
    html += '</tbody>';
    $('tss-mapping-table').innerHTML = html;
    $('tss-mapping-table').querySelectorAll('.tss-map-select').forEach(s => {
      s.addEventListener('change', e => {
        mapping[+e.target.dataset.bi] = +e.target.value;
        rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
      });
    });
  }

  // ─── Editable preview (mirrors template-standardize's combo grid) ───
  function renderPreview() {
    if (!srcData || !tplData || !formattedRows) return;
    renderExisting(); renderCollisions(); renderSmartFixes();
    const sec = $('tss-section-preview');
    const tbl = $('tss-preview-table');
    sec.style.display = '';
    const collisionIdxs = collisionRowSet();
    const nameColIdx = SITE_HEADERS.indexOf('Name*');
    const cityIdx = SITE_HEADERS.indexOf('City*');
    const colDrop = SITE_HEADERS.map(h => {
      const vals = dropdownValuesFor(h);
      return vals ? { vals, lower: new Set(vals.map(v => v.toLowerCase().trim())) } : null;
    });
    let html = '<thead><tr><th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    SITE_HEADERS.forEach((h, i) => {
      const req = /\*$/.test(h);
      html += '<th' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) +
        (colDrop[i] ? ' <span class="ts-col-listed">&#9662;</span>' : '') + '</th>';
    });
    html += '</tr></thead><tbody>';
    // A row "needs attention" if any cell would be flagged: empty required,
    // off-list dropdown value, name collision, or a guessed City.
    const rowHasIssue = ri => {
      const row = formattedRows[ri];
      for (let i = 0; i < SITE_HEADERS.length; i++) {
        const val = row[i] == null ? '' : String(row[i]);
        if (/\*$/.test(SITE_HEADERS[i]) && !val) return true;
        if (i === nameColIdx && collisionIdxs.has(ri)) return true;
        if (i === cityIdx && guessedCityRows.has(ri) && val) return true;
        const d = colDrop[i];
        if (d && val && !d.lower.has(val.toLowerCase().trim())) return true;
      }
      return false;
    };
    const allVisible = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const issueCount = allVisible.filter(rowHasIssue).length;
    if (previewIssuesOnly && !issueCount) previewIssuesOnly = false;
    const visibleRowIdxs = previewIssuesOnly ? allVisible.filter(rowHasIssue) : allVisible;
    const CAP = previewIssuesOnly ? 2000 : 50;
    const limit = Math.min(visibleRowIdxs.length, CAP);
    for (let r = 0; r < limit; r++) {
      const ri = visibleRowIdxs[r];
      const row = formattedRows[ri];
      html += '<tr><td style="text-align:center;padding:0;">' +
        '<button class="tss-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      SITE_HEADERS.forEach((h, i) => {
        const req = /\*$/.test(h);
        const val = row[i] == null ? '' : String(row[i]);
        const empty = !val;
        const d = colDrop[i];
        const inList = !!(d && val && d.lower.has(val.toLowerCase().trim()));
        let bg = '', title = '';
        if (req && empty) { bg = '#fee2e2'; title = 'Required — must be filled before export'; }
        else if (i === nameColIdx && collisionIdxs.has(ri)) {
          bg = '#fee2e2'; title = 'Name collision — another site shares this Name. Rename it here (or in Name Collisions above); PickTrace requires unique site names.';
        } else if (i === cityIdx && guessedCityRows.has(ri) && val) {
          bg = '#fef3c7'; title = 'City is a best-effort guess from the combined address — verify it.';
        } else if (d && val && !inList) {
          bg = '#fef3c7'; title = 'Off-list — "' + val + '" isn’t in the template dropdown for ' + h + '. Kept as-is, or pick a listed value.';
        }
        const cellStyle = bg ? ' style="background:' + bg + ';"' : '';
        const fieldColor = bg ? ('color:' + (bg === '#fee2e2' ? '#7f1d1d' : '#7c2d12') + ';') : '';
        if (d) {
          let opts = '<option value=""' + (empty ? ' selected' : '') + '>— blank —</option>';
          if (val && !inList) opts += '<option value="' + escHtml(val) + '" selected>' + escHtml(val) + '  (current)</option>';
          d.vals.forEach(v => {
            const s = (inList && v.toLowerCase().trim() === val.toLowerCase().trim()) ? ' selected' : '';
            opts += '<option value="' + escHtml(v) + '"' + s + '>' + escHtml(v) + '</option>';
          });
          opts += '<option value="__ts_custom__">✎ Type custom…</option>';
          html += '<td class="ts-cell"' + cellStyle + '><select class="ts-cell-select" data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') + (fieldColor ? ' style="' + fieldColor + '"' : '') + '>' + opts + '</select></td>';
        } else {
          html += '<td class="ts-cell"' + cellStyle + '><input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + i + '"' +
            (title ? ' title="' + escHtml(title) + '"' : '') + (fieldColor ? ' style="' + fieldColor + '"' : '') +
            ' value="' + escHtml(val) + '"></td>';
        }
      });
      html += '</tr>';
    }
    html += '</tbody>';
    tbl.innerHTML = html;
    const hint = sec.querySelector('.cmp-sites-hint');
    if (hint) {
      const toggle = '<label class="ts-issues-toggle" style="display:inline-flex;align-items:center;gap:6px;margin-right:12px;font-weight:600;' +
        (issueCount ? '' : 'opacity:.5;') + '"><input type="checkbox" id="tss-issues-only"' + (previewIssuesOnly ? ' checked' : '') +
        (issueCount ? '' : ' disabled') + '>Show only rows needing attention' + (issueCount ? ' (' + issueCount + ')' : ' (0)') + '</label>';
      let t = toggle;
      if (previewIssuesOnly) t += 'Showing ' + Math.min(limit, visibleRowIdxs.length) + ' of ' + issueCount + ' flagged row' + (issueCount === 1 ? '' : 's') + (visibleRowIdxs.length > limit ? ' (capped at ' + limit + ')' : '') + '. ';
      else if (visibleRowIdxs.length > limit) t += 'Showing first ' + limit + ' of ' + visibleRowIdxs.length + ' rows. ';
      t += 'Every cell is editable — <span class="ts-col-listed">&#9662;</span> columns are template dropdowns. Red = empty required; amber = off-list or a guessed City. Click <b>×</b> to drop a row.';
      if (removedRows.size) t += ' &nbsp; <button class="btn btn-ghost btn-sm" id="tss-restore-rows">Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
      hint.innerHTML = t;
      const io = $('tss-issues-only');
      if (io) io.addEventListener('change', e => { previewIssuesOnly = !!e.target.checked; renderPreview(); });
    }
    tbl.querySelectorAll('.tss-row-remove').forEach(btn => btn.addEventListener('click', e => {
      e.stopPropagation();
      removedRows.add(+e.currentTarget.dataset.ri);
      renderPreview(); renderRequired(); renderToCreate(); updateSummary();
    }));
    const restoreBtn = $('tss-restore-rows');
    if (restoreBtn) restoreBtn.addEventListener('click', () => {
      removedRows.clear(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
    });
    if (!tbl._tssCellWired) {
      tbl._tssCellWired = true;
      tbl.addEventListener('change', e => {
        const sel = e.target.closest('.ts-cell-select');
        if (sel) {
          const ri = +sel.dataset.ri, ci = +sel.dataset.ci;
          if (sel.value === '__ts_custom__') { swapSelectToInput(sel, ri, ci); return; }
          setCell(ri, ci, sel.value); recomputeAfterEdit(); return;
        }
        const inp = e.target.closest('.ts-cell-input');
        if (inp) { setCell(+inp.dataset.ri, +inp.dataset.ci, inp.value); recomputeAfterEdit(); }
      });
    }
    updateExportButton();
  }

  function setCell(ri, ci, val) {
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const v = val == null ? '' : String(val);
    if (ci === nameIdx) { if (v.trim()) nameOverrides[ri] = v.trim(); else delete nameOverrides[ri]; }
    else cellOverrides[ri + '|' + ci] = v;
    if (formattedRows[ri]) formattedRows[ri][ci] = v;
    if (ci === SITE_HEADERS.indexOf('City*')) guessedCityRows.delete(ri);
  }
  function swapSelectToInput(sel, ri, ci) {
    const td = sel.closest('td');
    if (!td) return;
    const cur = formattedRows[ri] ? (formattedRows[ri][ci] == null ? '' : String(formattedRows[ri][ci])) : '';
    td.innerHTML = '<input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + ci + '" placeholder="Type a value…" value="' + escHtml(cur) + '">';
    const inp = td.querySelector('input');
    if (inp) { inp.focus(); inp.select(); }
  }
  function recomputeAfterEdit() {
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }

  // ─── Smart Fixes panel — dropdown per off-list value (suggestion pre-selected) ───
  function renderSmartFixes() {
    const sec = $('tss-section-smartfix');
    if (!sec) return;
    const fixes = computeSmartFixes();
    if (!fixes.length) { sec.style.display = 'none'; $('tss-smartfix-table').innerHTML = ''; return; }
    sec.style.display = '';
    let totalRows = 0, needPick = 0;
    fixes.forEach(f => { totalRows += f.count; if (!f.suggestion) needPick++; });
    const t = $('tss-smartfix-title');
    if (t) t.textContent = 'Smart Fixes (' + fixes.length + ' value' + (fixes.length === 1 ? '' : 's') + ', ' + totalRows + ' rows' +
      (needPick ? ' — ' + needPick + ' need your choice' : '') + ')';
    let html = '<thead><tr><th>Column</th><th>Value in your data</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    fixes.forEach(f => {
      const opts = '<option value="">— pick a value —</option>' +
        f.options.slice().sort().map(o => '<option' + (f.suggestion && o === f.suggestion ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('');
      html += '<tr><td>' + escHtml(f.header) + '</td>' +
        '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(f.from) + '</span></td>' +
        '<td><select class="tss-smartfix-pick input-field" style="min-width:200px;">' + opts + '</select>' +
          (f.suggestion ? ' <span class="text-muted small">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') +
        '</td>' +
        '<td>' + f.count + '</td>' +
        '<td><button class="btn btn-primary btn-sm tss-smartfix-apply" data-ci="' + f.colIdx + '" data-from="' + escHtml(f.from) + '">Apply</button></td></tr>';
    });
    html += '</tbody>';
    $('tss-smartfix-table').innerHTML = html;
    $('tss-smartfix-table').querySelectorAll('.tss-smartfix-apply').forEach(btn => btn.addEventListener('click', e => {
      const tr = e.currentTarget.closest('tr');
      const to = tr.querySelector('.tss-smartfix-pick').value.trim();
      if (!to) { alert('Pick a value to set "' + e.currentTarget.dataset.from + '" to first.'); return; }
      applySmartFix(+e.currentTarget.dataset.ci, e.currentTarget.dataset.from, to);
    }));
  }
  function applySmartFix(ci, from, to) {
    smartFixMap[ci + '||' + String(from).toUpperCase().trim()] = to;
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  function applyAllSmartFixes() {
    const tbl = $('tss-smartfix-table');
    if (!tbl) return;
    const picks = [];
    tbl.querySelectorAll('.tss-smartfix-apply').forEach(btn => {
      const tr = btn.closest('tr');
      const to = tr.querySelector('.tss-smartfix-pick').value.trim();
      if (to) picks.push({ ci: +btn.dataset.ci, from: btn.dataset.from, to });
    });
    if (!picks.length) { alert('No values chosen yet. Pick a target for at least one row (suggested rows are pre-selected).'); return; }
    picks.forEach(p => { smartFixMap[p.ci + '||' + String(p.from).toUpperCase().trim()] = p.to; });
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }

  // ─── Employers / Site Types to create ───
  function renderToCreate() {
    if (!formattedRows || !tplData) return { p1: 0, p2: 0 };
    const empIdx = SITE_HEADERS.indexOf('Employer*');
    const typeIdx = SITE_HEADERS.indexOf('Site Type*');
    const empMap = buildCaseMap(tplData.dropdowns.get('employer'));
    const typeMap = buildCaseMap(tplData.dropdowns.get('site type'));
    const emps = new Map(), types = new Map();
    formattedRows.forEach((r, ri) => {
      if (excluded(ri)) return;
      const e = r[empIdx];
      if (e && empMap.size && !inDropdown(empMap, e)) emps.set(e, (emps.get(e) || 0) + 1);
      const t = r[typeIdx];
      if (t && typeMap.size && !inDropdown(typeMap, t)) types.set(t, (types.get(t) || 0) + 1);
    });
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
      html += '</tbody>';
      $(tableId).innerHTML = html;
    };
    renderList('tss-section-emp-create', 'tss-emp-create-table', 'tss-emp-create-title', 'Employers to Create / Reconcile in PickTrace', emps);
    renderList('tss-section-type-create', 'tss-type-create-table', 'tss-type-create-title', 'Site Types to Create / Reconcile in PickTrace', types);
    return { p1: emps.size, p2: types.size };
  }

  // ─── Empty required columns ───
  function renderRequired() {
    if (!formattedRows || !tplData) return;
    const sec = $('tss-section-required');
    sec.style.display = '';
    const requiredCols = SITE_HEADERS.map((h, i) => ({ header: h, idx: i, req: /\*$/.test(h) })).filter(x => x.req);
    const needsAction = [], filled = [];
    requiredCols.forEach(col => {
      let empty = 0; const sample = new Set();
      formattedRows.forEach((r, ri) => {
        if (excluded(ri)) return;
        if (!r[col.idx]) empty++;
        else if (sample.size < 1 && r[col.idx]) sample.add(r[col.idx]);
      });
      const fillVal = manualFills[col.idx] || (sample.size ? [...sample][0] : '');
      const entry = { ...col, empty, fillVal };
      (empty > 0 ? needsAction : filled).push(entry);
    });
    const renderRow = (col) => {
      const opts = tplData.dropdowns.get(norm(col.header).replace(/\*$/, '')) || new Set();
      const optsHtml = opts.size
        ? '<option value="">— pick —</option>' + [...opts].sort().map(o => '<option' + (o === col.fillVal ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const status = col.empty > 0
        ? '<span style="color:#dc2626;font-weight:600;">⚠ ' + col.empty + ' empty</span>'
        : '<span style="color:#15803d;">✓ filled' + (col.fillVal ? ' (e.g. ' + escHtml(String(col.fillVal).slice(0, 30)) + ')' : '') + '</span>';
      const rowStyle = col.empty > 0 ? '' : ' style="background:var(--bg-sunken);"';
      return '<tr' + rowStyle + '><td><b>' + escHtml(col.header) + '</b></td><td>' + status + '</td>' +
        '<td><select class="tss-req-pick input-field" data-idx="' + col.idx + '" style="min-width:180px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="tss-req-override input-field" data-idx="' + col.idx + '" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm tss-req-apply" data-idx="' + col.idx + '">Apply</button></td>' +
        '<td><button class="btn btn-ghost btn-sm tss-req-clear" data-idx="' + col.idx + '" title="Clear all values in this column">Clear</button></td></tr>';
    };
    let html = '<thead><tr><th>Column</th><th>Status</th><th>Pick from dropdown</th><th>Manual override</th><th></th><th></th></tr></thead>';
    if (needsAction.length) html += '<tbody><tr><td colspan="6" style="background:#fee2e2;color:#7f1d1d;font-weight:600;padding:6px 10px;">⚠ Action needed (' + needsAction.length + ')</td></tr>' + needsAction.map(renderRow).join('') + '</tbody>';
    if (filled.length) html += '<tbody><tr><td colspan="6" style="background:#dcfce7;color:#14532d;font-weight:600;padding:6px 10px;">✓ Already filled — change or clear if needed (' + filled.length + ')</td></tr>' + filled.map(renderRow).join('') + '</tbody>';
    $('tss-required-table').innerHTML = html;
    wireRequiredHandlers();
  }
  function applyFill(idx, val) {
    if (!val) return;
    formattedRows.forEach((r, ri) => { if (!excluded(ri)) r[idx] = val; });
    manualFills[idx] = val;
    renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  function clearColumn(idx) {
    formattedRows.forEach((r, ri) => { if (!excluded(ri)) r[idx] = ''; });
    delete manualFills[idx];
    renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  function wireRequiredHandlers() {
    const tbl = $('tss-required-table');
    if (!tbl) return;
    tbl.querySelectorAll('.tss-req-apply').forEach(btn => btn.addEventListener('click', e => {
      const idx = +e.target.dataset.idx, tr = e.target.closest('tr');
      const val = tr.querySelector('.tss-req-override').value.trim() || tr.querySelector('.tss-req-pick').value.trim();
      if (!val) { alert('Pick a value or type a manual override first.'); return; }
      applyFill(idx, val);
    }));
    tbl.querySelectorAll('.tss-req-clear').forEach(btn => btn.addEventListener('click', e => clearColumn(+e.target.dataset.idx)));
    tbl.querySelectorAll('.tss-req-pick').forEach(sel => sel.addEventListener('change', e => { const v = e.target.value.trim(); if (v) applyFill(+e.target.dataset.idx, v); }));
  }

  // ─── Existing-in-PickTrace panel ───
  function renderExisting() {
    const sec = $('tss-section-existing');
    if (!sec) return;
    if (!existingData || !existingRowIdxs.size) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const typeIdx = SITE_HEADERS.indexOf('Site Type*');
    const rows = [...existingRowIdxs].slice(0, 200);
    let html = '<thead><tr><th>Name</th><th>Site Type</th></tr></thead><tbody>';
    rows.forEach(ri => {
      const r = formattedRows[ri];
      html += '<tr><td>' + escHtml(r[nameIdx]) + '</td><td>' + escHtml(r[typeIdx] || '') + '</td></tr>';
    });
    html += '</tbody>';
    $('tss-existing-table').innerHTML = html;
    const t = $('tss-existing-title');
    if (t) t.textContent = 'Already in PickTrace — dropped (' + existingRowIdxs.size + ')';
  }

  // ─── Name collisions ───
  function renderCollisions() {
    const sec = $('tss-section-collisions');
    if (!sec) return;
    const groups = computeCollisions();
    if (!groups.size) { sec.style.display = 'none'; $('tss-collisions-table').innerHTML = ''; return; }
    sec.style.display = '';
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const typeIdx = SITE_HEADERS.indexOf('Site Type*');
    let total = 0; groups.forEach(arr => total += arr.length);
    const t = $('tss-collisions-title');
    if (t) t.textContent = 'Name Collisions (' + groups.size + ' name' + (groups.size === 1 ? '' : 's') + ', ' + total + ' to fix)';
    let html = '<thead><tr><th>Source</th><th>Name</th><th>Site Type</th><th>New name</th><th></th></tr></thead><tbody>';
    [...groups.entries()].forEach(([k, arr]) => {
      if (existingTakenKeys.has(k)) {
        const r0 = formattedRows[arr[0]];
        html += '<tr style="background:var(--bg-sunken,#f3f4f6);color:#6b7280;"><td><b>In PickTrace</b></td>' +
          '<td><b>' + escHtml(r0[nameIdx]) + '</b></td><td></td><td colspan="2"><i>already exists — name is taken</i></td></tr>';
      }
      arr.forEach(ri => {
        const r = formattedRows[ri];
        html += '<tr><td>New</td><td><b style="color:#dc2626;">' + escHtml(r[nameIdx]) + '</b></td>' +
          '<td>' + escHtml(r[typeIdx] || '') + '</td>' +
          '<td><input type="text" class="tss-coll-name input-field" data-ri="' + ri + '" placeholder="' + escHtml(r[nameIdx]) + '" style="width:160px;"></td>' +
          '<td><button class="btn btn-primary btn-sm tss-coll-apply" data-ri="' + ri + '">Rename</button></td></tr>';
      });
    });
    html += '</tbody>';
    $('tss-collisions-table').innerHTML = html;
    $('tss-collisions-table').querySelectorAll('.tss-coll-apply').forEach(btn => btn.addEventListener('click', e => {
      const tr = e.target.closest('tr'), ri = +e.target.dataset.ri;
      const val = tr.querySelector('.tss-coll-name').value.trim();
      if (!val) { alert('Type a new name first.'); return; }
      nameOverrides[ri] = val;
      rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
    }));
  }

  // ─── Summary + export gating ───
  function updateSummary() {
    if (!srcData || !tplData) { $('tss-summary').style.display = 'none'; return; }
    const reqIdxs = SITE_HEADERS.map((h, i) => ({ h, i, req: /\*$/.test(h) })).filter(x => x.req);
    let emptyReq = 0;
    if (formattedRows) reqIdxs.forEach(({ i }) => formattedRows.forEach((r, ri) => { if (!excluded(ri) && !r[i]) emptyReq++; }));
    const tc = renderToCreate() || { p1: 0, p2: 0 };
    const visibleCount = formattedRows ? formattedRows.filter((_, i) => !excluded(i)).length : 0;
    const totalCount = formattedRows ? formattedRows.length : 0;
    let collisionRows = 0; computeCollisions().forEach(arr => collisionRows += arr.length);
    const sum = $('tss-summary');
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visibleCount + '</b> sites' + (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + totalCount + ')</span>' : '') + '</div>' +
      (existingRowIdxs.size ? '<div class="cmp-stat cmp-warn"><b>' + existingRowIdxs.size + '</b> already in PickTrace (dropped)</div>' : '') +
      (collisionRows ? '<div class="cmp-stat cmp-warn"><b>' + collisionRows + '</b> name collision rows</div>' : '') +
      '<div class="cmp-stat"><b>' + SITE_HEADERS.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (tc.p1 ? '<div class="cmp-stat cmp-warn"><b>' + tc.p1 + '</b> employers to reconcile</div>' : '') +
      (tc.p2 ? '<div class="cmp-stat cmp-warn"><b>' + tc.p2 + '</b> site types to reconcile</div>' : '');
  }

  function getExportBlockers() {
    if (!srcData || !tplData || !formattedRows) return null;
    const reqIdxs = SITE_HEADERS.map((h, i) => /\*$/.test(h) ? i : -1).filter(i => i >= 0);
    let emptyReq = 0;
    formattedRows.forEach((r, ri) => { if (excluded(ri)) return; reqIdxs.forEach(i => { if (!r[i]) emptyReq++; }); });
    // Off-list values in any dropdown-backed column.
    const dropCols = SITE_HEADERS.map((h, i) => ({ h, i, map: buildCaseMap(tplData.dropdowns.get(norm(h).replace(/\*$/, ''))) })).filter(x => x.map.size);
    const unknownByCol = [];
    dropCols.forEach(({ h, i, map }) => {
      const vals = new Set();
      formattedRows.forEach((r, ri) => { if (!excluded(ri) && r[i] && !inDropdown(map, r[i])) vals.add(r[i]); });
      if (vals.size) unknownByCol.push({ header: h, values: [...vals] });
    });
    const collisionGroups = computeCollisions();
    let nameCollisions = 0; const collisionSamples = [];
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    collisionGroups.forEach(arr => { nameCollisions += arr.length; collisionSamples.push(String(formattedRows[arr[0]][nameIdx] || '') + ' (×' + arr.length + ')'); });
    return { emptyReq, unknownByCol, nameCollisions, collisionSamples };
  }
  function hasBlockers(b) { return b && (b.emptyReq > 0 || b.unknownByCol.length || b.nameCollisions > 0); }

  function updateExportButton() {
    const btn = $('tss-export');
    btn.disabled = !(srcData && tplData && formattedRows && formattedRows.length);
    const b = getExportBlockers();
    if (!b) { btn.title = ''; return; }
    const reasons = [];
    if (b.emptyReq > 0) reasons.push(b.emptyReq + ' empty required cells');
    b.unknownByCol.forEach(c => reasons.push(c.values.length + ' ' + c.header.replace(/\*$/, '') + ' value(s) not in dropdown'));
    if (b.nameCollisions > 0) reasons.push(b.nameCollisions + ' duplicate site names');
    btn.title = reasons.length ? 'Click to review before exporting: ' + reasons.join(', ') + '.' : 'Ready to export.';
  }

  function showExportConfirm(b) {
    return new Promise(resolve => {
      const prior = document.getElementById('tss-export-confirm');
      if (prior) prior.remove();
      const overlay = document.createElement('div');
      overlay.id = 'tss-export-confirm';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const sections = [];
      if (b.emptyReq > 0) sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + b.emptyReq + ' empty required cells</div><div class="ts-confirm-issue-body">PickTrace will reject sites missing values for required (*) columns. Fill them via the Empty Required Columns panel.</div></div>');
      b.unknownByCol.forEach(c => {
        const sample = c.values.slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = c.values.length > 8 ? '<li class="ts-confirm-more">… and ' + (c.values.length - 8) + ' more</li>' : '';
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + c.values.length + ' ' + escHtml(c.header) + ' value(s) not in your template dropdown</div><ul class="ts-confirm-list">' + sample + more + '</ul><div class="ts-confirm-issue-body">Reconcile these in PickTrace (create them, or fix the spelling to match), then re-download the template.</div></div>');
      });
      if (b.nameCollisions > 0) {
        const sample = (b.collisionSamples || []).slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + b.nameCollisions + ' site(s) share a Name</div><div class="ts-confirm-issue-body">PickTrace requires unique site names — one of each pair will be <b>silently dropped</b>. Rename one in the <b>Name Collisions</b> panel first.<ul class="ts-confirm-list">' + sample + '</ul></div></div>');
      }
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner';
      inner.innerHTML = '<h3>Export anyway?</h3><div class="ts-confirm-issues">' + sections.join('') + '</div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="tss-conf-cancel">Cancel</button><button class="btn btn-primary" id="tss-conf-ok">Export anyway</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
      const close = v => { overlay.remove(); resolve(v); };
      inner.querySelector('#tss-conf-ok').addEventListener('click', () => close(true));
      inner.querySelector('#tss-conf-cancel').addEventListener('click', () => close(false));
      overlay.addEventListener('click', e => { if (e.target === overlay) close(false); });
    });
  }

  // ─── Export ───
  function doExport() {
    if (!formattedRows || !tplData) return;
    // Canonicalize dropdown-backed columns to the template's exact casing.
    const caseMaps = SITE_HEADERS.map(h => buildCaseMap(tplData.dropdowns.get(norm(h).replace(/\*$/, ''))));
    const exportRows = formattedRows.filter((_, ri) => !excluded(ri)).map(row => {
      const r = row.slice();
      caseMaps.forEach((m, i) => { if (m.size && r[i]) { const k = String(r[i]).toUpperCase().trim(); if (m.has(k)) r[i] = m.get(k); } });
      return r;
    });
    const wb = XLSX.read(new Uint8Array(tplData.rawBuffer), { type: 'array', cellStyles: true, cellDates: true, sheetStubs: true });
    const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[dataName];
    if (!ws) { alert('Template missing DATA ENTRY sheet — cannot export.'); return; }
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: SITE_HEADERS.length - 1 } };
    for (let r = 1; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) delete ws[ref];
      }
    }
    const colCount = Math.max(range.e.c + 1, SITE_HEADERS.length);
    exportRows.forEach((row, ri) => {
      for (let c = 0; c < colCount; c++) {
        const val = row[c];
        if (val == null || val === '') continue;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c });
        // Keep Zip as text so leading zeros survive (e.g. 07094).
        ws[ref] = { v: String(val), t: 's' };
      }
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, exportRows.length), c: colCount - 1 } });
    const origName = (tplData.fileName || 'sites-template.xlsx').trim();
    const dotIdx = origName.lastIndexOf('.');
    const base = dotIdx > 0 ? origName.substring(0, dotIdx) : origName;
    const ext = dotIdx > 0 ? origName.substring(dotIdx) : '.xlsx';
    XLSX.writeFile(wb, base + ' — filled' + ext, { cellStyles: true });
  }

  // ─── File load handlers ───
  function runFormat() {
    if (!srcData || !tplData) { alert('Upload source data + a personalized Sites template first.'); return; }
    mapping = autoMap(SITE_HEADERS, srcData.headers);
    manualFills = {}; removedRows = new Set(); cellOverrides = {}; nameOverrides = {};
    smartFixMap = {}; previewIssuesOnly = false;
    rebuildFormattedRows();
    renderMapping(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
    $('tss-empty').style.display = 'none';
  }
  function handleSrcFile(file) {
    readSrcFile(file).then(data => {
      srcData = data;
      $('tss-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('tss-src-meta').textContent = data.rows.length + ' rows · ' + data.headers.length + ' columns';
      $('tss-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    }).catch(err => { if (err && err.message === 'cancelled') return; alert('Failed to read source: ' + (err && err.message ? err.message : err)); });
  }
  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      const parsed = parseTemplate(e.target.result, file.name);
      if (!parsed || !parsed.headers.length) { alert('Could not find a DATA ENTRY sheet in this template.'); return; }
      tplData = parsed;
      $('tss-tpl-name').textContent = file.name;
      const warn = looksLikeSitesTemplate(parsed.headers) ? '' : ' <span class="cmp-warn">⚠ doesn’t look like a Sites template</span>';
      $('tss-tpl-meta').innerHTML = parsed.headers.length + ' columns · ' +
        (parsed.dropdowns.get('site type') ? parsed.dropdowns.get('site type').size : 0) + ' site types · ' +
        (parsed.dropdowns.get('employer') ? parsed.dropdowns.get('employer').size : 0) + ' employers' + warn;
      $('tss-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    };
    r.readAsArrayBuffer(file);
  }
  function handleExistingFile(file) {
    readSrcFile(file).then(data => {
      const nameI = data.headers.findIndex(h => { const k = norm(h).replace(/\*$/, ''); return k === 'name' || k === 'site name' || k === 'site' || k === 'sites'; });
      if (nameI < 0) { alert('Existing-sites file must have a Name (or Site) column.'); return; }
      const typeI = data.headers.findIndex(h => norm(h).replace(/\*$/, '') === 'site type');
      const keys = new Set(), byName = new Map();
      data.rows.forEach(r => {
        const n = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (!n) return;
        const k = nameKeyOf(n);
        keys.add(k);
        const arr = byName.get(k) || [];
        arr.push({ type: typeI >= 0 ? String(r[typeI] || '').trim() : '' });
        byName.set(k, arr);
      });
      existingData = { keys, byName, fileName: file.name, count: keys.size };
      $('tss-existing-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('tss-existing-meta').textContent = keys.size + ' existing site' + (keys.size === 1 ? '' : 's') + ' indexed';
      if (srcData && tplData) { rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary(); }
    }).catch(err => { if (err && err.message === 'cancelled') return; alert('Failed to read existing-sites file: ' + (err && err.message ? err.message : err)); });
  }
  function importFromOrgData() {
    if (typeof aggregateOrgData !== 'function' || typeof getAllOrgNames !== 'function') { alert('Org Data store unavailable. Reload the page.'); return; }
    const names = getAllOrgNames();
    if (!names.length) { alert('No saved orgs found.'); return; }
    const pick = prompt('Type the org name to import:\n\n' + names.join('\n'));
    if (!pick) return;
    const found = names.find(n => n.toLowerCase() === pick.trim().toLowerCase());
    if (!found) { alert('Org "' + pick + '" not found.'); return; }
    const aggregated = aggregateOrgData(found);
    if (!aggregated || !aggregated.rows.length) { alert('No rows for "' + found + '".'); return; }
    srcData = { headers: aggregated.headers, rows: aggregated.rows, fileName: '(Org Data: ' + found + ')', sheetName: '(synthesized)' };
    $('tss-src-name').textContent = '(Org Data: ' + found + ')';
    $('tss-src-meta').textContent = aggregated.rows.length + ' rows · ' + aggregated.headers.length + ' columns';
    $('tss-run').disabled = !(srcData && tplData);
    if (srcData && tplData) runFormat();
  }

  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    manualFills = {}; removedRows = new Set(); existingData = null; existingRowIdxs = new Set();
    nameOverrides = {}; cellOverrides = {}; smartFixMap = {}; previewIssuesOnly = false; guessedCityRows = new Set(); collisionKeys = new Set(); existingTakenKeys = new Set();
    $('tss-src-name').textContent = 'No file selected';
    $('tss-tpl-name').textContent = 'No file selected';
    $('tss-src-meta').textContent = ''; $('tss-tpl-meta').textContent = '';
    { const en = $('tss-existing-name'); if (en) en.textContent = 'No file selected'; }
    { const em = $('tss-existing-meta'); if (em) em.textContent = ''; }
    ['tss-src-file', 'tss-tpl-file', 'tss-existing-file'].forEach(id => { const el = $(id); if (el) el.value = ''; });
    ['tss-section-mapping', 'tss-section-smartfix', 'tss-section-emp-create', 'tss-section-type-create', 'tss-section-required', 'tss-section-existing', 'tss-section-collisions', 'tss-section-preview']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    $('tss-summary').style.display = 'none';
    $('tss-empty').style.display = '';
    $('tss-run').disabled = true; $('tss-export').disabled = true;
  }

  // ─── Init + domain toggle ───
  function init() {
    if (initialized) return;
    if (!$('tss-src-file')) return; // markup not present yet
    initialized = true;
    $('tss-src-file').addEventListener('change', e => { if (e.target.files[0]) handleSrcFile(e.target.files[0]); e.target.value = ''; });
    $('tss-tpl-file').addEventListener('change', e => { if (e.target.files[0]) handleTplFile(e.target.files[0]); e.target.value = ''; });
    { const ef = $('tss-existing-file'); if (ef) ef.addEventListener('change', e => { if (e.target.files[0]) handleExistingFile(e.target.files[0]); e.target.value = ''; }); }
    { const og = $('tss-src-from-org'); if (og) og.addEventListener('click', importFromOrgData); }
    $('tss-run').addEventListener('click', runFormat);
    $('tss-export').addEventListener('click', async () => {
      const b = getExportBlockers();
      if (hasBlockers(b)) { const ok = await showExportConfirm(b); if (!ok) return; }
      doExport();
    });
    $('tss-reset').addEventListener('click', reset);
    { const sfa = $('tss-smartfix-apply-all'); if (sfa) sfa.addEventListener('click', applyAllSmartFixes); }
    // Click-to-copy in the results area.
    $('tss-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      if (!td || td.closest('#tss-section-required') || e.target.closest('input, button, select')) return;
      const text = (td.textContent || '').trim();
      if (text && navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(() => td.classList.toggle('cmp-copied'), () => {});
    });
    // Domain toggle (Locations / Sites) — shared control at the top of the page.
    document.querySelectorAll('.ts-domain-pill').forEach(btn => {
      btn.addEventListener('click', () => {
        const d = btn.dataset.domain;
        document.querySelectorAll('.ts-domain-pill').forEach(b => b.classList.toggle('ts-domain-active', b === btn));
        const loc = $('ts-loc-wrap'), site = $('ts-site-wrap');
        if (loc) loc.style.display = d === 'locations' ? '' : 'none';
        if (site) site.style.display = d === 'sites' ? '' : 'none';
      });
    });
  }

  window.tssInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
