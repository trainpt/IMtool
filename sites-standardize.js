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
//
// Reference addresses: some implementation workbooks don't repeat the address
// on every row. They put a bare state code ("CA", "AZ") in the Address column
// and list each state's one real address ONCE, off to the side — often in a
// column with no header at all. scanReferenceAddresses() finds those, the
// Reference Addresses panel lets you confirm/correct the street-vs-city split,
// and every row carrying that code is filled from the confirmed entry.
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
  // Renaming one half of a colliding pair makes the other half unique, which
  // would drop the whole group out of computeCollisions() before the second row
  // could be dealt with. The panel therefore renders from this history — every
  // group that has EVER collided this run — not from the live collision set.
  let collisionHistory = new Map();  // original nameKey → { base, rows:Set<rowIdx> }
  let collisionDismissed = new Set();// nameKeys the user has finished with
  let groupColIdx = -1;              // source column holding the grower / site group
  let groupColPicked = false;        // true once the user overrides the detected column
  let completedGroups = new Set();   // site groups ticked off as CREATED in PickTrace
  let assignedSites = new Map();     // group → Set(site name) already assigned to it
  let groupWorkSel = '';             // group open in the assignment worklist
  let groupWorkFilter = '';          // its filter box
  // The Site Groups tab can run on its own file — you may want to fix a
  // grouping months after the sites went up, with no upload to build.
  let groupSrcOwn = null;            // { headers, rows, fileName, sheetName }
  let groupExistingOwn = null;       // { keys:Set<nameKey>, fileName, count }
  let groupNameColIdx = -1;          // site-name column, when running on its own file
  let groupNameColPicked = false;
  let refAddrMap = {};            // STATE CODE → reference-address entry (sticky; see scanReferenceAddresses)
  let refAddrExtras = [];         // additional full-address cells found for a code already claimed
  let xrefCodeCounts = new Map(); // STATE CODE → how many rows carry it as a bare Address value
  let xrefRows = new Map();       // formattedRows index → STATE CODE it was filled from
  let provMarks = {};             // "row|col" → provenance code, for the debug dump
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
      // A route number belongs to the street, not to the city: in
      // "12000 S. Hwy. 99 Fairview" the suffix is "Hwy." but the split goes
      // after "99", or the city comes out as "99 Fairview". Keep at least one
      // token back for the city itself.
      while (cut >= 0 && cut + 1 < toks.length - 1 && /^\d+[A-Za-z]?$/.test(toks[cut + 1])) cut++;
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

  // ─── Reference addresses ───
  // Workbooks that share one address across many sites put a bare state code in
  // the Address column and spell the real address out once, somewhere else on
  // the sheet — frequently in a column past the last header, which is why the
  // scan walks every cell of every row rather than the mapped columns.

  // "CA" / "AZ" / "BC" on its own — not part of a longer string.
  function bareStateCode(v) {
    const s = String(v == null ? '' : v).trim().toUpperCase();
    if (!/^[A-Z]{2}$/.test(s)) return '';
    return (US_STATES.has(s) || CA_PROVINCES.has(s)) ? s : '';
  }
  // A cell that reads like a complete street address ending in STATE + ZIP.
  function looksLikeFullAddress(v) {
    const s = String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
    if (s.length < 10 || s.length > 200) return '';
    const m = s.match(/[,\s]([A-Za-z]{2})[,\s]+(\d{5}(?:-\d{4})?|[A-Za-z]\d[A-Za-z]\s?\d[A-Za-z]\d)$/);
    if (!m) return '';
    const code = m[1].toUpperCase();
    if (!US_STATES.has(code) && !CA_PROVINCES.has(code)) return '';
    // Needs a street number ahead of the state, or it's a city/state/zip stub.
    if (!/\d/.test(s.slice(0, m.index))) return '';
    return code;
  }
  // How much to trust the street/city split parseAddress produced.
  function splitConfidence(p) {
    if (!p.city) return 'check';
    if (!p.cityGuessed) return 'clean';
    const toks = p.city.split(' ').filter(Boolean);
    if (/\d/.test(p.city) || toks.length > 3) return 'check';
    return 'guessed';
  }
  function makeRefEntry(code, raw, origin) {
    const p = parseAddress(raw);
    return {
      code, raw, origin,
      address1: p.address1 || '',
      city: p.city || '',
      state: p.state || code,
      zip: p.zip || '',
      country: p.country || (CA_PROVINCES.has(code) ? 'CA' : 'US'),
      confidence: splitConfidence(p),
      confirmed: false,
      edited: false,
      manual: !raw
    };
  }
  // Tally the bare state codes sitting in whichever source column feeds
  // Address1*. No codes → this workbook spells its addresses out per row and
  // the whole reference mechanism stays off.
  function countStateCodes() {
    const counts = new Map();
    if (!srcData) return counts;
    const a1 = SITE_HEADERS.indexOf('Address1*');
    const sc = mapping[a1];
    if (sc == null || sc < 0) return counts;
    srcData.rows.forEach(r => {
      const code = bareStateCode(r[sc]);
      if (code) counts.set(code, (counts.get(code) || 0) + 1);
    });
    return counts;
  }
  // Seeds refAddrMap. User edits are sticky, so an entry already confirmed or
  // hand-edited is left exactly as it is on a re-scan.
  function scanReferenceAddresses() {
    refAddrExtras = [];
    xrefCodeCounts = countStateCodes();
    if (!srcData || !xrefCodeCounts.size) return;
    const a1 = SITE_HEADERS.indexOf('Address1*');
    const addrCol = mapping[a1] != null ? mapping[a1] : -1;
    const found = new Map();
    srcData.rows.forEach((row, ri) => {
      for (let ci = 0; ci < row.length; ci++) {
        if (ci === addrCol) continue;              // that column holds the codes
        const code = looksLikeFullAddress(row[ci]);
        if (!code) continue;
        const sheetRow = (srcData.rowNums && srcData.rowNums[ri] != null) ? srcData.rowNums[ri] + 1 : ri + 1;
        const origin = {
          sheet: srcData.sheetName || null,
          cell: colLetter(ci) + sheetRow,
          column: srcData.headers[ci] ? srcData.headers[ci] : '(no header)',
          dataRowIndex: ri
        };
        const raw = String(row[ci]).trim().replace(/\s+/g, ' ');
        if (found.has(code)) { refAddrExtras.push({ code, raw, origin }); continue; }
        found.set(code, { raw, origin });
      }
    });
    // Every code actually used gets a panel row — including ones with no
    // address anywhere, so they can be typed in rather than silently skipped.
    xrefCodeCounts.forEach((_, code) => {
      const hit = found.get(code);
      const prior = refAddrMap[code];
      if (prior && (prior.confirmed || prior.edited)) {
        if (hit && !prior.raw) { prior.raw = hit.raw; prior.origin = hit.origin; }
        return;
      }
      refAddrMap[code] = hit
        ? makeRefEntry(code, hit.raw, hit.origin)
        : makeRefEntry(code, '', null);
    });
    // Drop entries for codes this mapping no longer uses.
    Object.keys(refAddrMap).forEach(code => { if (!xrefCodeCounts.has(code)) delete refAddrMap[code]; });
  }
  function refEntryUsable(e) { return !!(e && (e.address1 || e.city || e.zip)); }
  function refNeedsAttention(e) { return !!(e && !e.confirmed && (e.confidence === 'check' || !refEntryUsable(e))); }
  function refAddrOutstanding() {
    return Object.keys(refAddrMap).filter(c => refNeedsAttention(refAddrMap[c]));
  }
  // A split the parser flagged as wrong-looking, waved through without anyone
  // touching the fields. Not a blocker — the operator may know better — but it
  // rides on every row using that code, so it stays visible in the panel, in
  // the export dialog and in the debug dump instead of disappearing on a click.
  function refConfirmedUnfixed(e) { return !!(e && e.confirmed && !e.edited && e.confidence === 'check'); }
  function refAddrConfirmedUnfixed() {
    return Object.keys(refAddrMap).filter(c => refConfirmedUnfixed(refAddrMap[c]));
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
  // `indexed` is [{ row, i }] where i is the 0-based row index in the original
  // sheet. Carrying it through means the debug dump and the Reference Addresses
  // panel can name a real cell ("Site!J3") instead of a post-filter offset.
  function sliceSheetAoa(indexed) {
    const aoa = indexed.map(x => x.row);
    const hIdx = detectHeaderRow(aoa);
    const headers = (aoa[hIdx] || []).map(h => String(h == null ? '' : h).trim());
    const kept = indexed.slice(hIdx + 1).filter(x => !isInstructionRow(x.row));
    return {
      headers,
      rows: kept.map(x => x.row.map(c => c == null ? '' : String(c).trim())),
      rowNums: kept.map(x => x.i),          // 0-based sheet row per data row
      headerRow: indexed[hIdx] ? indexed[hIdx].i : hIdx
    };
  }
  function colLetter(n) {
    let s = '';
    let x = n + 1;
    while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); }
    return s;
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
            const indexed = aoa.map((row, i) => ({ row, i }))
              .filter(x => x.row && x.row.some(c => c != null && String(c).trim() !== ''));
            if (!indexed.length) return null;
            const sliced = sliceSheetAoa(indexed);
            if (!sliced.headers.filter(h => h).length) return null;
            return { name: n, headers: sliced.headers, rows: sliced.rows, rowNums: sliced.rowNums, headerRow: sliced.headerRow };
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({ headers: s.headers, rows: s.rows, rowNums: s.rowNums,
            headerRow: s.headerRow, fileName: file.name, sheetName: s.name });
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
    xrefRows = new Map();
    provMarks = {};
    return srcData.rows.map((srcRow, ri) => {
      const out = new Array(SITE_HEADERS.length).fill('');
      SITE_HEADERS.forEach((bh, bi) => {
        const si2 = mapping[bi];
        if (si2 != null && si2 >= 0 && srcRow[si2] != null) {
          const v = String(srcRow[si2]).trim();
          if (v) out[bi] = v;
        }
      });
      // A bare state code in Address1* means the real address lives in the
      // reference table, not on this row. Address1 holds only the code, so it
      // is replaced outright; the rest fill only where empty (a real source
      // column still wins).
      const code = a1 >= 0 ? bareStateCode(out[a1]) : '';
      const ref = code ? refAddrMap[code] : null;
      if (code && refEntryUsable(ref)) {
        const mark = 'xref:' + code;
        out[a1] = ref.address1 || '';
        provMarks[ri + '|' + a1] = ref.address1 ? mark : 'empty';
        if (ci >= 0 && !out[ci] && ref.city) { out[ci] = ref.city; provMarks[ri + '|' + ci] = mark; }
        if (si >= 0 && !out[si]) { out[si] = ref.state || code; provMarks[ri + '|' + si] = mark; }
        if (zi >= 0 && !out[zi] && ref.zip) { out[zi] = ref.zip; provMarks[ri + '|' + zi] = mark; }
        if (coi >= 0 && !out[coi] && ref.country) { out[coi] = ref.country; provMarks[ri + '|' + coi] = mark; }
        xrefRows.set(ri, code);
      } else if (a1 >= 0 && out[a1] && !(out[ci] && out[si] && out[zi])) {
        // Parse a combined address (only when City/State/Zip aren't already
        // mapped in full). Fills only the empty target cells.
        const p = parseAddress(out[a1]);
        if (p.address1) { out[a1] = p.address1; provMarks[ri + '|' + a1] = 'parse:address'; }
        if (ci >= 0 && !out[ci] && p.city) { out[ci] = p.city; provMarks[ri + '|' + ci] = 'parse:address'; if (p.cityGuessed) guessedCityRows.add(ri); }
        if (si >= 0 && !out[si] && p.state) { out[si] = p.state; provMarks[ri + '|' + si] = 'parse:address'; }
        if (zi >= 0 && !out[zi] && p.zip) { out[zi] = p.zip; provMarks[ri + '|' + zi] = 'parse:address'; }
        if (coi >= 0 && !out[coi] && p.country) { out[coi] = p.country; provMarks[ri + '|' + coi] = 'parse:address'; }
        if (code) xrefRows.set(ri, code); // code with no usable reference entry
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
    // Fold this pass's collisions into the sticky history the panel renders
    // from, so resolving half a pair doesn't hide the other half.
    recordCollisions();
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
        // Re-point Address1* and the reference table has to be rebuilt around
        // the new column (edited entries survive).
        scanReferenceAddresses();
        rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
      });
    });
  }

  // ─── Editable preview (mirrors template-standardize's combo grid) ───
  function renderPreview() {
    if (!srcData || !tplData || !formattedRows) return;
    renderExisting(); renderCollisions(); renderRefAddresses(); renderSmartFixes();
    renderGroupsTab();
    const sec = $('tss-section-preview');
    const tbl = $('tss-preview-table');
    sec.style.display = '';
    const collisionIdxs = collisionRowSet();
    const nameColIdx = SITE_HEADERS.indexOf('Name*');
    const cityIdx = SITE_HEADERS.indexOf('City*');
    // Address cells filled from a reference entry you haven't vetted yet.
    const addrCols = new Set(['Address1*', 'City*', 'State*', 'Zip*', 'Country*']
      .map(h => SITE_HEADERS.indexOf(h)).filter(i => i >= 0));
    const refPending = ri => {
      const code = xrefRows.get(ri);
      return code && refNeedsAttention(refAddrMap[code]) ? code : '';
    };
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
        if (addrCols.has(i) && refPending(ri)) return true;
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
        } else if (addrCols.has(i) && refPending(ri)) {
          bg = '#fef3c7';
          title = 'Filled from the reference address for "' + refPending(ri) +
            '" — that entry still needs your check in the Reference Addresses panel above. ' +
            'Fixing it there updates every row using this code.';
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
      t += 'Every cell is editable — <span class="ts-col-listed">&#9662;</span> columns are template dropdowns. Red = empty required; amber = off-list, a guessed City, or an unconfirmed reference address. Click <b>×</b> to drop a row.';
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

  // ─── Reference Addresses panel ───
  // One row per bare state code the Address column uses. Shows what was found
  // and where, the street/city split it produced, and how much to trust that
  // split. Editing any field confirms the entry; "Looks right" confirms it
  // as-is. Both push straight back through every row carrying the code.
  const REF_FIELDS = [
    { key: 'address1', label: 'Address1*', width: '190px' },
    { key: 'city',     label: 'City*',     width: '130px' },
    { key: 'state',    label: 'State*',    width: '60px'  },
    { key: 'zip',      label: 'Zip*',      width: '80px'  },
    { key: 'country',  label: 'Country*',  width: '70px'  }
  ];
  function refStatusCell(e) {
    if (!refEntryUsable(e)) {
      return '<span style="color:#dc2626;font-weight:600;">&#9888; no address found</span>' +
        '<div class="text-muted small">Type one in — these rows export with an empty Address1.</div>';
    }
    if (refConfirmedUnfixed(e)) {
      return '<span style="color:#b45309;font-weight:600;">&#10003; confirmed &mdash; but the split was flagged</span>' +
        '<div class="text-muted small">You waved this through without editing it. Check <b>Address1</b> and <b>City</b> above.</div>';
    }
    if (e.confirmed) return '<span style="color:#15803d;font-weight:600;">&#10003; confirmed</span>';
    if (e.confidence === 'check') {
      return '<span style="color:#b45309;font-weight:600;">&#9888; check this split</span>' +
        '<div class="text-muted small">The city looks wrong — fix Address1/City, then it applies to every row.</div>';
    }
    if (e.confidence === 'guessed') return '<span style="color:#b45309;">~ auto-split</span>';
    return '<span style="color:#15803d;">&#10003; parsed</span>';
  }
  function renderRefAddresses() {
    const sec = $('tss-section-refaddr');
    if (!sec) return;
    const codes = Object.keys(refAddrMap).sort();
    if (!codes.length) { sec.style.display = 'none'; $('tss-refaddr-table').innerHTML = ''; return; }
    sec.style.display = '';
    const totalRows = codes.reduce((n, c) => n + (xrefCodeCounts.get(c) || 0), 0);
    const outstanding = refAddrOutstanding().length;
    const t = $('tss-refaddr-title');
    if (t) t.textContent = 'Reference Addresses (' + codes.length + ' code' + (codes.length === 1 ? '' : 's') +
      ', ' + totalRows + ' rows' + (outstanding ? ' — ' + outstanding + ' need your check' : '') + ')';
    const hint = sec.querySelector('.cmp-sites-hint');
    if (hint) {
      hint.innerHTML = 'Your Address column holds state codes, not addresses — the real addresses were found elsewhere on the sheet. ' +
        'Confirm each split below and it fills <b>Address1 / City / State / Zip / Country</b> on every row carrying that code. ' +
        'Amber = the street/city guess is uncertain.';
    }
    let html = '<thead><tr><th>Code</th><th>Found in</th>' +
      REF_FIELDS.map(f => '<th>' + escHtml(f.label) + '</th>').join('') +
      '<th>Rows</th><th>Status</th><th></th></tr></thead><tbody>';
    codes.forEach(code => {
      const e = refAddrMap[code];
      const rows = xrefCodeCounts.get(code) || 0;
      const bad = refNeedsAttention(e);
      const tint = bad ? ' style="background:#fffbeb;"' : '';
      const origin = e.origin
        ? '<code>' + escHtml((e.origin.sheet ? e.origin.sheet + '!' : '') + e.origin.cell) + '</code>' +
          '<div class="text-muted small">column: ' + escHtml(e.origin.column) + '</div>'
        : '<span class="text-muted small">not found on the sheet</span>';
      const raw = e.raw
        ? '<div class="text-muted small" style="max-width:240px;">raw: ' + escHtml(e.raw) + '</div>' : '';
      html += '<tr' + tint + '><td><b>' + escHtml(code) + '</b></td>' +
        '<td>' + origin + raw + '</td>' +
        REF_FIELDS.map(f => '<td><input type="text" class="tss-ref-input input-field" data-code="' + escHtml(code) +
          '" data-key="' + f.key + '" style="width:' + f.width + ';" value="' + escHtml(e[f.key] || '') + '"></td>').join('') +
        '<td><b>' + rows + '</b></td>' +
        '<td>' + refStatusCell(e) + '</td>' +
        '<td><button class="btn ' + (bad ? 'btn-primary' : 'btn-ghost') + ' btn-sm tss-ref-ok" data-code="' +
          escHtml(code) + '">' +
          (e.confirmed ? 'Confirmed' : (e.confidence === 'check' ? 'Confirm anyway' : 'Looks right')) +
        '</button></td></tr>';
    });
    html += '</tbody>';
    $('tss-refaddr-table').innerHTML = html;

    if (refAddrExtras.length) {
      const extra = refAddrExtras.slice(0, 10).map(x =>
        '<li><b>' + escHtml(x.code) + '</b> &mdash; <code>' + escHtml((x.origin.sheet ? x.origin.sheet + '!' : '') + x.origin.cell) +
        '</code> ' + escHtml(x.raw) + '</li>').join('');
      const note = $('tss-refaddr-extra');
      if (note) {
        note.style.display = '';
        note.innerHTML = '<div class="text-muted small">Other addresses found for codes already covered above (not used — edit the row above if one of these is the right one):<ul>' +
          extra + (refAddrExtras.length > 10 ? '<li>… and ' + (refAddrExtras.length - 10) + ' more</li>' : '') + '</ul></div>';
      }
    } else {
      const note = $('tss-refaddr-extra');
      if (note) { note.style.display = 'none'; note.innerHTML = ''; }
    }

    const tbl = $('tss-refaddr-table');
    tbl.querySelectorAll('.tss-ref-input').forEach(inp => inp.addEventListener('change', ev => {
      const el = ev.currentTarget;
      const e = refAddrMap[el.dataset.code];
      if (!e) return;
      e[el.dataset.key] = el.value.trim();
      e.edited = true;
      e.confirmed = true;
      applyRefAddresses();
    }));
    tbl.querySelectorAll('.tss-ref-ok').forEach(btn => btn.addEventListener('click', ev => {
      const e = refAddrMap[ev.currentTarget.dataset.code];
      if (!e) return;
      if (!refEntryUsable(e)) { alert('Type an address for ' + e.code + ' first — there is nothing to confirm yet.'); return; }
      e.confirmed = true;
      applyRefAddresses();
    }));
    // "Confirm all" deliberately skips entries the parser flagged as a bad
    // split. Sweeping those up is how a wrong street/city reaches every row
    // using the code — those need a look, one at a time.
    const all = $('tss-refaddr-confirm-all');
    if (all) {
      const pending = codes.filter(c => confirmableInBulk(refAddrMap[c]));
      const held = codes.filter(c => refEntryUsable(refAddrMap[c]) && !refAddrMap[c].confirmed &&
        refAddrMap[c].confidence === 'check');
      all.style.display = pending.length ? '' : 'none';
      all.textContent = 'Confirm the ' + pending.length + ' clean one' + (pending.length === 1 ? '' : 's');
      all.title = held.length
        ? held.length + ' flagged split' + (held.length === 1 ? '' : 's') + ' (' + held.join(', ') +
          ') are left out on purpose — check each one yourself.'
        : 'Confirm every reference address whose split parsed cleanly.';
    }
  }
  function confirmableInBulk(e) {
    return !!(e && refEntryUsable(e) && !e.confirmed && e.confidence !== 'check');
  }
  function applyRefAddresses() {
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  function confirmAllRefAddresses() {
    const pending = Object.keys(refAddrMap).filter(c => confirmableInBulk(refAddrMap[c]));
    if (!pending.length) return;
    pending.forEach(c => { refAddrMap[c].confirmed = true; });
    applyRefAddresses();
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

  // ─── Site Groups ───
  // The bulk Sites template has no Site Group column, and groups can't be bulk
  // created — PickTrace files every uploaded site under a group of its own
  // name, so a source workbook's grower-client grouping is simply lost on
  // import. This panel is the worklist for rebuilding it by hand: each group,
  // the sites that belong to it, and a tick once you've done it.
  const GROUP_ALIASES = ['site group', 'sitegroup', 'group', 'grower', 'grower client',
    'growerclient', 'client', 'customer', 'ranch group', 'ranchgroup'];
  const NAME_ALIASES = ['site name', 'sitename', 'name', 'site', 'sites', 'ranch', 'location name'];
  // Whichever file the tab is running on: its own upload, else the Sites tab's.
  // `nameAt` is the important difference — on the Sites tab it returns the FINAL
  // name (collision renames included), which is what exists in PickTrace.
  function groupDataset() {
    if (groupSrcOwn) {
      const keys = groupExistingOwn && groupExistingOwn.keys;
      return {
        own: true, headers: groupSrcOwn.headers, rows: groupSrcOwn.rows,
        fileName: groupSrcOwn.fileName, sheetName: groupSrcOwn.sheetName,
        nameAt: ri => groupNameColIdx < 0 ? ''
          : String((groupSrcOwn.rows[ri] || [])[groupNameColIdx] == null ? ''
            : groupSrcOwn.rows[ri][groupNameColIdx]).trim(),
        skip: () => false,
        isExisting: ri => {
          if (!keys || !keys.size || groupNameColIdx < 0) return false;
          const n = String((groupSrcOwn.rows[ri] || [])[groupNameColIdx] || '').trim();
          return !!n && keys.has(nameKeyOf(n));
        }
      };
    }
    if (srcData && formattedRows) {
      const nameIdx = SITE_HEADERS.indexOf('Name*');
      const includeExisting = !!(existingData && existingData.keys && existingData.keys.size);
      return {
        own: false, headers: srcData.headers, rows: srcData.rows,
        fileName: srcData.fileName, sheetName: srcData.sheetName,
        nameAt: ri => String((formattedRows[ri] || [])[nameIdx] == null ? '' : formattedRows[ri][nameIdx]).trim(),
        skip: ri => removedRows.has(ri) || (existingRowIdxs.has(ri) && !includeExisting),
        isExisting: ri => existingRowIdxs.has(ri)
      };
    }
    return null;
  }
  function detectGroupCol() {
    const ds = groupDataset();
    if (!ds) return -1;
    // On the Sites tab, never pick a column that already feeds a template column.
    const used = ds.own ? new Set()
      : new Set(Object.keys(mapping).map(k => mapping[k]).filter(i => i >= 0));
    let hit = -1;
    ds.headers.forEach((h, i) => {
      if (hit >= 0 || used.has(i)) return;
      if (GROUP_ALIASES.indexOf(norm(h)) >= 0 || GROUP_ALIASES.indexOf(normLoose(h)) >= 0) hit = i;
    });
    return hit;
  }
  function detectNameCol() {
    const ds = groupDataset();
    if (!ds || !ds.own) return -1;
    let hit = -1;
    ds.headers.forEach((h, i) => {
      if (hit >= 0 || i === groupColIdx) return;
      if (NAME_ALIASES.indexOf(norm(h)) >= 0 || NAME_ALIASES.indexOf(normLoose(h)) >= 0) hit = i;
    });
    return hit;
  }
  // Which rows count toward a group. With an existing-sites export loaded we
  // also know about the rows dropped for already being in PickTrace — those are
  // live sites that still need assigning, so they belong on the list. Rows
  // removed by hand are off the job and stay out either way.
  function siteGroupScope() {
    const ds = groupDataset();
    if (!ds) return 'nothing loaded';
    if (ds.own) return (groupExistingOwn && groupExistingOwn.keys.size)
      ? 'every row in ' + ds.fileName + ', existing sites flagged'
      : 'every row in ' + ds.fileName;
    return (existingData && existingData.keys && existingData.keys.size)
      ? 'exported + already in PickTrace'
      : 'exported only';
  }
  function collectSiteGroups() {
    const out = new Map(); // group → { sites:[names], created, existing }
    const ds = groupDataset();
    if (!ds || groupColIdx < 0) return out;
    ds.rows.forEach((row, ri) => {
      if (ds.skip(ri)) return;
      const g = String(row[groupColIdx] == null ? '' : row[groupColIdx]).trim();
      if (!g) return;
      const e = out.get(g) || { sites: [], created: 0, existing: 0 };
      const nm = ds.nameAt(ri);
      if (nm) e.sites.push(nm);
      if (ds.isExisting(ri)) e.existing++; else e.created++;
      out.set(g, e);
    });
    return out;
  }
  function copyText(text, okMsg) {
    if (!navigator.clipboard || !navigator.clipboard.writeText) return;
    navigator.clipboard.writeText(text).then(() => { if (okMsg) flashGroupNote(okMsg); }, () => {});
  }
  // The two group panels sit far apart on the page, so a confirmation goes to
  // both — whichever one you're looking at shows it.
  function flashGroupNote(msg) {
    const els = ['tss-sitegroups-note', 'tss-groupwork-note'].map($).filter(Boolean);
    if (!els.length) return;
    els.forEach(el => { el.textContent = msg; el.style.display = ''; });
    clearTimeout(flashGroupNote._t);
    flashGroupNote._t = setTimeout(() => { els.forEach(el => { el.style.display = 'none'; }); }, 2500);
  }
  // Everything the Site Groups tab owns: its own summary bar, its empty state,
  // and the two panels. Called on every Sites rebuild AND on entering the tab.
  function renderGroupsTab() {
    const ds = groupDataset();
    if (!groupColPicked) groupColIdx = detectGroupCol();
    if (ds && ds.own && !groupNameColPicked) groupNameColIdx = detectNameCol();
    renderSiteGroups();
    renderGroupWork();
    renderGroupSummary();
    renderGroupLink();
    // Offer the Sites tab's data only when it has some and we're not already on it.
    const useBtn = $('tssg-use-sites');
    if (useBtn) useBtn.style.display = (groupSrcOwn && srcData && formattedRows) ? '' : 'none';
    const has = !!ds && collectSiteGroups().size > 0;
    const empty = $('tss-group-empty');
    if (empty) {
      empty.style.display = has ? 'none' : '';
      empty.innerHTML = !ds
        ? 'Choose a <b>source data</b> file above &mdash; or load one on the <b>Sites</b> tab and it will be read from there.'
        : (groupColIdx < 0
            ? 'No grower / site group column found in <code>' + escHtml(ds.fileName || 'that file') +
              '</code>. If it has one under another name, pick it with the <b>Group column</b> selector.'
            : (ds.own && groupNameColIdx < 0
                ? 'Found the group column, but not a site-name column. Pick one with the <b>Site name column</b> selector.'
                : 'No site group values found in that column.'));
    }
  }
  function renderGroupSummary() {
    const sum = $('tss-group-summary');
    if (!sum) return;
    const groups = collectSiteGroups();
    if (!groups.size) { sum.style.display = 'none'; return; }
    let siteTotal = 0, siteDone = 0;
    groups.forEach((info, k) => { const p = groupProgress(k, info); siteTotal += p.total; siteDone += p.done; });
    const toCreate = [...groups.keys()].filter(k => !completedGroups.has(k)).length;
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + groups.size + '</b> site groups</div>' +
      (toCreate ? '<div class="cmp-stat cmp-warn"><b>' + toCreate + '</b> still to create</div>'
                : '<div class="cmp-stat"><b>all</b> groups created</div>') +
      '<div class="cmp-stat"><b>' + siteTotal + '</b> sites to assign</div>' +
      (siteDone < siteTotal
        ? '<div class="cmp-stat cmp-warn"><b>' + (siteTotal - siteDone) + '</b> still to assign</div>'
        : '<div class="cmp-stat"><b>done</b> — every site assigned</div>') +
      '<div class="cmp-stat"><b>' + escHtml(siteGroupScope()) + '</b></div>';
  }
  // A short pointer on the Sites tab, so the work isn't invisible from there.
  function renderGroupLink() {
    const sec = $('tss-section-grouplink');
    if (!sec) return;
    // Only speak for the Sites tab's own data — if the Site Groups tab is
    // running on a different file, this pointer would be describing it.
    if (groupSrcOwn || !srcData || !formattedRows) { sec.style.display = 'none'; return; }
    const groups = collectSiteGroups();
    if (!groups.size) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    let siteTotal = 0, siteDone = 0;
    groups.forEach((info, k) => { const p = groupProgress(k, info); siteTotal += p.total; siteDone += p.done; });
    const left = [...groups.keys()].filter(k => !completedGroups.has(k)).length;
    const t = $('tss-grouplink-title');
    if (t) t.textContent = 'Site Groups (' + groups.size + ')';
    const hint = sec.querySelector('.cmp-sites-hint');
    if (hint) {
      hint.innerHTML = 'Your source groups these sites under <b>' + groups.size + '</b> grower client' +
        (groups.size === 1 ? '' : 's') + ', but the Sites template has no column for it &mdash; groups ' +
        '<b>cannot be bulk created</b> and are lost on import. ' +
        '<b>' + left + '</b> group' + (left === 1 ? '' : 's') + ' to create and <b>' + (siteTotal - siteDone) +
        '</b> of ' + siteTotal + ' sites to assign by hand: see the <b>Site Groups</b> tab above.';
    }
  }

  function renderSiteGroups() {
    const sec = $('tss-section-sitegroups');
    if (!sec) return;
    const ds = groupDataset();
    if (!ds) { sec.style.display = 'none'; return; }
    const groups = collectSiteGroups();
    // Hide only when nothing was found AND the user hasn't touched the picker —
    // once they have, the panel must stay up or the picker becomes unreachable.
    if (groupColIdx < 0 && !groups.size && !groupColPicked) { sec.style.display = 'none'; return; }
    sec.style.display = '';

    const entries = [...groups.entries()].sort((a, b) => b[1].sites.length - a[1].sites.length);
    const done = entries.filter(([g]) => completedGroups.has(g)).length;
    const t = $('tss-sitegroups-title');
    if (t) t.textContent = 'Site Groups — create & assign in PickTrace (' +
      (entries.length - done) + ' to do' + (done ? ' · ' + done + ' done' : '') + ')';

    // Column pickers — neither column is ever exported, so they aren't in
    // Column Mapping and get their own selectors here. The site-name picker
    // only matters when this tab is running on its own file; on the Sites tab
    // the name comes from the formatted rows, renames and all.
    const opts = (sel) => '<option value="-1">— none —</option>' + ds.headers.map((h, i) =>
      '<option value="' + i + '"' + (i === sel ? ' selected' : '') + '>' +
      escHtml(h || '(column ' + colLetter(i) + ')') + '</option>').join('');
    const pick = $('tss-sitegroups-col');
    if (pick) pick.innerHTML = opts(groupColIdx);
    const namePick = $('tss-sitegroups-namecol');
    const nameWrap = $('tss-sitegroups-namecol-wrap');
    if (nameWrap) nameWrap.style.display = ds.own ? '' : 'none';
    if (namePick && ds.own) namePick.innerHTML = opts(groupNameColIdx);

    const hint = sec.querySelector('.cmp-sites-hint');
    if (hint) {
      hint.innerHTML = 'Site groups <b>cannot be bulk created</b> — the Sites template has no column for them, so ' +
        'PickTrace files every uploaded site under a group of its own name. Rebuild them by hand here: ' +
        '<b>click a row</b> to copy the group name and tick it off, and <b>&#10697; sites</b> to copy that group\'s ' +
        'site names one per line for assigning. Counting <b>' + escHtml(siteGroupScope()) + '</b>' +
        (siteGroupScope() === 'exported only'
          ? ' — load your existing-sites export to also count sites already in PickTrace.' : '.');
    }

    if (!entries.length) {
      $('tss-sitegroups-table').innerHTML = '<tbody><tr><td class="text-muted">' +
        (groupColIdx < 0
          ? 'No group column selected — pick the one holding your grower / site group above.'
          : 'No site group values found in that column.') + '</td></tr></tbody>';
      return;
    }
    let html = '<thead><tr><th style="width:28px;" title="Group created in PickTrace"></th><th>Site group</th>' +
      '<th>Sites</th><th>Already in PickTrace</th><th>Assigned</th><th></th><th></th></tr></thead><tbody>';
    entries.forEach(([g, info]) => {
      const isDone = completedGroups.has(g);
      const mark = isDone ? '<span style="color:#15803d;font-weight:700;">&#10003;</span>'
                          : '<span style="color:#9ca3af;">&#9744;</span>';
      const st = isDone ? ' style="text-decoration:line-through;color:#15803d;"' : '';
      const p = groupProgress(g, info);
      const prog = p.done === 0
        ? '<span class="text-muted">0 / ' + p.total + '</span>'
        : '<span style="color:' + (p.done === p.total ? '#15803d;font-weight:600' : '#b45309') + ';">' +
          p.done + ' / ' + p.total + (p.done === p.total ? ' &#10003;' : '') + '</span>';
      html += '<tr class="tss-group-row" data-group="' + escHtml(g) + '" style="cursor:pointer;" ' +
        'title="Click to copy this group name and tick it off as created">' +
        '<td style="text-align:center;">' + mark + '</td>' +
        '<td><b' + st + '>' + escHtml(g) + '</b></td>' +
        '<td>' + info.sites.length + '</td>' +
        '<td>' + (info.existing ? info.existing : '<span class="text-muted">—</span>') + '</td>' +
        '<td>' + prog + '</td>' +
        '<td><button class="btn btn-primary btn-sm tss-group-open" data-group="' + escHtml(g) +
          '" title="Open this group\'s sites as a checklist">Assign sites &rarr;</button></td>' +
        '<td><button class="btn btn-ghost btn-sm tss-group-copy" data-group="' + escHtml(g) +
          '" title="Copy this group\'s site names, one per line">&#10697; ' + info.sites.length + '</button></td></tr>';
    });
    html += '</tbody>';
    $('tss-sitegroups-table').innerHTML = html;

    const tbl = $('tss-sitegroups-table');
    tbl.querySelectorAll('.tss-group-row').forEach(tr => tr.addEventListener('click', e => {
      if (e.target.closest('button')) return;      // the copy button handles itself
      const g = tr.dataset.group;
      if (completedGroups.has(g)) completedGroups.delete(g); else completedGroups.add(g);
      copyText(g);
      renderSiteGroups(); renderGroupSummary(); renderGroupLink(); updateSummary();
    }));
    tbl.querySelectorAll('.tss-group-copy').forEach(btn => btn.addEventListener('click', e => {
      e.stopPropagation();
      const g = e.currentTarget.dataset.group;
      const info = collectSiteGroups().get(g);
      if (!info || !info.sites.length) return;
      copyText(info.sites.join('\n'), 'Copied ' + info.sites.length + ' site name' +
        (info.sites.length === 1 ? '' : 's') + ' for "' + g + '".');
    }));
    tbl.querySelectorAll('.tss-group-open').forEach(btn => btn.addEventListener('click', e => {
      e.stopPropagation();
      openGroupWork(e.currentTarget.dataset.group);
    }));
    const copyAll = $('tss-sitegroups-copy-all');
    if (copyAll) copyAll.onclick = () => {
      const lines = entries.map(([g, info]) => g + '\t' + info.sites.length).join('\n');
      copyText(lines, 'Copied ' + entries.length + ' group names with their site counts.');
    };
  }

  // ─── Site-group assignment worklist ───
  // Creating the group is one click; putting 261 sites into it is the actual
  // job. This is that job as a checklist: one chip per site, ticked as you go,
  // so you can stop halfway and still know where you were.
  function assignedFor(group) {
    let s = assignedSites.get(group);
    if (!s) { s = new Set(); assignedSites.set(group, s); }
    return s;
  }
  function groupProgress(group, info) {
    const done = assignedFor(group);
    const total = info ? info.sites.length : 0;
    let n = 0;
    if (info) info.sites.forEach(s => { if (done.has(s)) n++; });
    return { done: n, total: total, pct: total ? Math.round(n / total * 100) : 0 };
  }
  function renderGroupWork() {
    const sec = $('tss-section-groupwork');
    if (!sec) return;
    const groups = collectSiteGroups();
    if (!groups.size) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const names = [...groups.keys()].sort((a, b) => groups.get(b).sites.length - groups.get(a).sites.length);
    if (!groupWorkSel || !groups.has(groupWorkSel)) groupWorkSel = names[0];
    const info = groups.get(groupWorkSel);
    const done = assignedFor(groupWorkSel);
    const prog = groupProgress(groupWorkSel, info);

    const sel = $('tss-groupwork-sel');
    if (sel) {
      sel.innerHTML = names.map(n => {
        const p = groupProgress(n, groups.get(n));
        const label = n + '  (' + (p.total - p.done) + ' of ' + p.total + ' left)';
        return '<option value="' + escHtml(n) + '"' + (n === groupWorkSel ? ' selected' : '') + '>' +
          escHtml(label) + '</option>';
      }).join('');
    }
    const t = $('tss-groupwork-title');
    if (t) t.textContent = 'Assign Sites to a Site Group — ' + groupWorkSel +
      ' (' + (prog.total - prog.done) + ' of ' + prog.total + ' still to add)';
    const bar = $('tss-groupwork-bar');
    if (bar) bar.style.width = prog.pct + '%';

    const f = groupWorkFilter.toLowerCase();
    const shown = info.sites.filter(s => !f || s.toLowerCase().indexOf(f) >= 0);
    const cnt = $('tss-groupwork-count');
    if (cnt) cnt.textContent = f
      ? shown.length + ' of ' + info.sites.length + ' match · ' + prog.done + ' added'
      : prog.done + ' of ' + prog.total + ' added (' + prog.pct + '%)';

    const list = $('tss-groupwork-list');
    if (list) {
      list.innerHTML = shown.length
        ? shown.map(s => '<button type="button" class="tss-site-chip' + (done.has(s) ? ' is-done' : '') +
            '" data-site="' + escHtml(s) + '" title="Click to copy this name and tick it off">' +
            escHtml(s) + '</button>').join('')
        : '<span class="text-muted small">No sites match that filter.</span>';
    }
  }
  // Toggling a chip updates it in place — re-rendering 261 chips on every click
  // would throw away the scroll position and the filter box's focus.
  function toggleAssigned(group, site, chipEl) {
    const set = assignedFor(group);
    if (set.has(site)) set.delete(site); else { set.add(site); copyText(site); }
    if (chipEl) chipEl.classList.toggle('is-done', set.has(site));
    const info = collectSiteGroups().get(group);
    const prog = groupProgress(group, info);
    const bar = $('tss-groupwork-bar');
    if (bar) bar.style.width = prog.pct + '%';
    const cnt = $('tss-groupwork-count');
    if (cnt && !groupWorkFilter) cnt.textContent = prog.done + ' of ' + prog.total + ' added (' + prog.pct + '%)';
    const t = $('tss-groupwork-title');
    if (t) t.textContent = 'Assign Sites to a Site Group — ' + group +
      ' (' + (prog.total - prog.done) + ' of ' + prog.total + ' still to add)';
    renderSiteGroups(); renderGroupSummary(); renderGroupLink();
    updateSummary();
  }
  function openGroupWork(group) {
    groupWorkSel = group;
    groupWorkFilter = '';
    const fb = $('tss-groupwork-filter');
    if (fb) fb.value = '';
    renderGroupWork();
    const sec = $('tss-section-groupwork');
    if (sec && sec.scrollIntoView) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
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
  // Two sites sharing a Name are usually still distinguishable in the source —
  // a different state, a different cost code — the Sites schema just has no
  // column to carry it. Find the source column that tells them apart and
  // propose it as a suffix, instead of asking for 16 invented names. Rows that
  // nothing distinguishes are real duplicates and get called that.
  function collisionDiscriminator(rowIdxs) {
    if (!srcData || !rowIdxs || rowIdxs.length < 2) return null;
    const width = srcData.rows.reduce((m, r) => Math.max(m, (r || []).length), srcData.headers.length);
    let best = null;
    for (let ci = 0; ci < width; ci++) {
      const vals = rowIdxs.map(ri => {
        const r = srcData.rows[ri] || [];
        return String(r[ci] == null ? '' : r[ci]).trim();
      });
      if (vals.some(v => !v)) continue;                                   // needs a value on every row
      if (new Set(vals.map(v => v.toUpperCase())).size !== vals.length) continue; // must be distinct
      const maxLen = Math.max.apply(null, vals.map(v => v.length));
      if (maxLen > 20) continue;                                          // too long to be a suffix
      if (!best || maxLen < best.maxLen) {
        best = { colIdx: ci, header: srcData.headers[ci] || '(no header)', values: vals, maxLen };
      }
    }
    return best;
  }
  // The row's name as the SOURCE spells it. Suggestions build on this rather
  // than on the current cell, so applying one twice can't yield "WEST-FIELD AZ AZ".
  function srcNameOf(ri) {
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const sc = mapping[nameIdx];
    if (sc != null && sc >= 0 && srcData && srcData.rows[ri]) {
      const v = String(srcData.rows[ri][sc] == null ? '' : srcData.rows[ri][sc]).trim();
      if (v) return v;
    }
    return (formattedRows && formattedRows[ri]) ? String(formattedRows[ri][nameIdx] || '').trim() : '';
  }
  // Fold the live collisions into the sticky history. Called on every rebuild,
  // so a group survives the rename that resolves half of it.
  function recordCollisions() {
    computeCollisions().forEach((arr, k) => {
      const e = collisionHistory.get(k) || { base: srcNameOf(arr[0]), rows: new Set() };
      arr.forEach(ri => e.rows.add(ri));
      collisionHistory.set(k, e);
    });
  }
  function collisionSuggestions() {
    const out = new Map(); // nameKey → { base, rowIdxs, disc, names: Map<ri, suggestedName> }
    const colliding = collisionRowSet();
    collisionHistory.forEach((e, k) => {
      if (collisionDismissed.has(k)) {
        // A dismissed group comes back if it starts colliding again — e.g. the
        // row that was dropped to resolve it gets restored.
        const live = [...e.rows].some(ri => !removedRows.has(ri) && !existingRowIdxs.has(ri) && colliding.has(ri));
        if (!live) return;
        collisionDismissed.delete(k);
      }
      const rowIdxs = [...e.rows].sort((a, b) => a - b);
      const disc = collisionDiscriminator(rowIdxs);
      const names = new Map();
      if (disc) rowIdxs.forEach((ri, i) => names.set(ri, (srcNameOf(ri) + ' ' + disc.values[i]).trim()));
      out.set(k, { base: e.base, rowIdxs, disc, names });
    });
    return out;
  }
  // A group is done once none of its surviving rows still collides.
  function collisionGroupState(rowIdxs) {
    const colliding = collisionRowSet();
    const live = rowIdxs.filter(ri => !removedRows.has(ri) && !existingRowIdxs.has(ri));
    const dropped = rowIdxs.filter(ri => removedRows.has(ri));
    return { live, dropped, unresolved: live.filter(ri => colliding.has(ri)) };
  }
  function applyAllCollisionSuggestions() {
    let n = 0;
    collisionSuggestions().forEach(g => {
      g.names.forEach((v, ri) => { if (!removedRows.has(ri)) { nameOverrides[ri] = v; n++; } });
    });
    if (!n) { alert('No suggestions available — none of these collisions can be told apart by a source column.'); return; }
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  // Drop a site from the export entirely, straight from this panel.
  function dropCollisionRow(ri) {
    removedRows.add(ri);
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }
  function restoreCollisionRow(ri) {
    removedRows.delete(ri);
    rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
  }

  function renderCollisions() {
    const sec = $('tss-section-collisions');
    if (!sec) return;
    const sugg = collisionSuggestions();
    if (!sugg.size) { sec.style.display = 'none'; $('tss-collisions-table').innerHTML = ''; return; }
    sec.style.display = '';
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const typeIdx = SITE_HEADERS.indexOf('Site Type*');

    let openGroups = 0, openRows = 0, suggestable = 0, identical = 0;
    sugg.forEach(g => {
      const st = collisionGroupState(g.rowIdxs);
      if (st.unresolved.length) { openGroups++; openRows += st.unresolved.length; }
      if (g.disc) { g.names.forEach((v, ri) => { if (!removedRows.has(ri) && nameOverrides[ri] !== v) suggestable++; }); }
      else if (st.unresolved.length) identical++;
    });
    const t = $('tss-collisions-title');
    if (t) t.textContent = 'Name Collisions (' + openGroups + ' name' + (openGroups === 1 ? '' : 's') +
      ', ' + openRows + ' to fix)' + (openGroups === 0 ? ' — all resolved' : '');
    const hint = sec.querySelector('.cmp-sites-hint');
    if (hint) {
      hint.innerHTML = 'PickTrace requires unique site names &mdash; it silently drops one of each duplicate pair. ' +
        'Each group stays listed until you are done with it, so you can rename <b>both</b> halves. ' +
        (suggestable ? 'Where a source column tells the rows apart (a different state, a different cost code), a name is suggested. ' : '') +
        (identical ? '<b>' + identical + '</b> group' + (identical === 1 ? ' is' : 's are') +
          ' identical in every source column &mdash; those are true duplicates, so <b>Drop</b> one. ' : '') +
        'Use <b>Drop</b> to remove a site from the export entirely.';
    }
    const btnAll = $('tss-coll-apply-all');
    if (btnAll) {
      btnAll.style.display = suggestable ? '' : 'none';
      btnAll.textContent = 'Apply all ' + suggestable + ' suggestion' + (suggestable === 1 ? '' : 's');
    }

    let html = '<thead><tr><th>Source</th><th>Name now</th><th>Site Type</th><th>Told apart by</th>' +
      '<th>New name</th><th></th><th></th></tr></thead><tbody>';
    [...sugg.entries()].forEach(([k, g]) => {
      const st = collisionGroupState(g.rowIdxs);
      const done = st.unresolved.length === 0;
      // Renaming one half makes the other unique, so the group is technically
      // resolved while its sibling still carries the bare name. Say so rather
      // than flashing a green "resolved" that hides work you meant to finish.
      const pending = g.disc
        ? g.rowIdxs.filter(ri => !removedRows.has(ri) && nameOverrides[ri] == null).length : 0;
      const tone = done ? (pending ? '#fef3c7' : '#dcfce7') : '#fee2e2';
      const ink  = done ? (pending ? '#7c2d12' : '#14532d') : '#7f1d1d';
      html += '<tr><td colspan="7" style="background:' + tone + ';color:' + ink +
        ';font-weight:600;padding:6px 10px;">' +
        (done ? (pending ? '&#10003; ' : '&#10003; ') : '&#9888; ') + escHtml(g.base) +
        ' — ' + g.rowIdxs.length + ' row' + (g.rowIdxs.length === 1 ? '' : 's') +
        (st.dropped.length ? ', ' + st.dropped.length + ' dropped' : '') +
        (done
          ? (pending
              ? ' — no longer colliding, but ' + pending + ' row' + (pending === 1 ? '' : 's') +
                ' still ' + (pending === 1 ? 'has' : 'have') + ' an unapplied suggestion'
              : ' — resolved')
          : ', ' + st.unresolved.length + ' still colliding') +
        (done ? '<button class="btn btn-ghost btn-sm tss-coll-dismiss" data-k="' + escHtml(k) +
          '" style="margin-left:10px;">Dismiss</button>' : '') +
        '</td></tr>';
      if (existingTakenKeys.has(k)) {
        html += '<tr style="background:var(--bg-sunken,#f3f4f6);color:#6b7280;"><td><b>In PickTrace</b></td>' +
          '<td><b>' + escHtml(g.base) + '</b></td><td></td><td colspan="4"><i>already exists — this name is taken</i></td></tr>';
      }
      g.rowIdxs.forEach((ri, i) => {
        const r = formattedRows[ri];
        if (!r) return;
        const isDropped = removedRows.has(ri);
        const renamed = nameOverrides[ri] != null;
        const suggested = g.names.get(ri) || '';
        const stillColliding = st.unresolved.indexOf(ri) >= 0;
        if (isDropped) {
          html += '<tr style="color:#9ca3af;"><td>Dropped</td>' +
            '<td><s>' + escHtml(srcNameOf(ri)) + '</s></td><td colspan="4"><i>removed from the export</i></td>' +
            '<td><button class="btn btn-ghost btn-sm tss-coll-restore" data-ri="' + ri + '">Restore</button></td></tr>';
          return;
        }
        const by = g.disc
          ? escHtml(g.disc.header) + ' = <b>' + escHtml(g.disc.values[i]) + '</b>'
          : '<span style="color:#b45309;">nothing — identical rows</span>';
        const nameCell = stillColliding
          ? '<b style="color:#dc2626;">' + escHtml(r[nameIdx]) + '</b>'
          : '<span style="color:#15803d;">' + escHtml(r[nameIdx]) + (renamed ? ' &#10003;' : '') + '</span>';
        // Once a row has been renamed its box starts empty — re-offering the
        // suggestion would invite appending the suffix a second time.
        const boxVal = renamed ? '' : suggested;
        html += '<tr><td>New</td><td>' + nameCell + '</td>' +
          '<td>' + escHtml(r[typeIdx] || '') + '</td>' +
          '<td class="small">' + by + '</td>' +
          '<td><input type="text" class="tss-coll-name input-field" data-ri="' + ri + '" placeholder="' +
            escHtml(r[nameIdx]) + '" value="' + escHtml(boxVal) + '" style="width:180px;"></td>' +
          '<td><button class="btn ' + (boxVal ? 'btn-primary' : 'btn-ghost') + ' btn-sm tss-coll-apply" data-ri="' + ri +
            '">' + (renamed ? 'Rename again' : 'Rename') + '</button></td>' +
          '<td><button class="btn btn-ghost btn-sm tss-coll-drop" data-ri="' + ri +
            '" title="Remove this site from the export entirely">Drop</button></td></tr>';
      });
    });
    html += '</tbody>';
    $('tss-collisions-table').innerHTML = html;
    const tbl = $('tss-collisions-table');
    tbl.querySelectorAll('.tss-coll-apply').forEach(btn => btn.addEventListener('click', e => {
      const ri = +e.currentTarget.dataset.ri;
      const tr = e.currentTarget.closest('tr');
      const val = tr.querySelector('.tss-coll-name').value.trim();
      if (!val) { alert('Type a new name first.'); return; }
      nameOverrides[ri] = val;
      rebuildFormattedRows(); renderPreview(); renderRequired(); renderToCreate(); updateSummary();
    }));
    tbl.querySelectorAll('.tss-coll-drop').forEach(btn => btn.addEventListener('click', e =>
      dropCollisionRow(+e.currentTarget.dataset.ri)));
    tbl.querySelectorAll('.tss-coll-restore').forEach(btn => btn.addEventListener('click', e =>
      restoreCollisionRow(+e.currentTarget.dataset.ri)));
    tbl.querySelectorAll('.tss-coll-dismiss').forEach(btn => btn.addEventListener('click', e => {
      collisionDismissed.add(e.currentTarget.dataset.k);
      renderCollisions();
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
    const refPend = refAddrOutstanding();
    let refPendRows = 0; refPend.forEach(c => refPendRows += (xrefCodeCounts.get(c) || 0));
    const sum = $('tss-summary');
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visibleCount + '</b> sites' + (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + totalCount + ')</span>' : '') + '</div>' +
      (refPendRows ? '<div class="cmp-stat cmp-warn"><b>' + refPendRows + '</b> rows on an unconfirmed reference address</div>' : '') +
      (existingRowIdxs.size ? '<div class="cmp-stat cmp-warn"><b>' + existingRowIdxs.size + '</b> already in PickTrace (dropped)</div>' : '') +
      (collisionRows ? '<div class="cmp-stat cmp-warn"><b>' + collisionRows + '</b> name collision rows</div>' : '') +
      '<div class="cmp-stat"><b>' + SITE_HEADERS.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (tc.p1 ? '<div class="cmp-stat cmp-warn"><b>' + tc.p1 + '</b> employers to reconcile</div>' : '') +
      (tc.p2 ? '<div class="cmp-stat cmp-warn"><b>' + tc.p2 + '</b> site types to reconcile</div>' : '') +
      (function () {
        const g = collectSiteGroups();
        if (!g.size) return '';
        const left = [...g.keys()].filter(k => !completedGroups.has(k)).length;
        let siteTotal = 0, siteDone = 0;
        g.forEach((info, k) => { const p = groupProgress(k, info); siteTotal += p.total; siteDone += p.done; });
        return (left
          ? '<div class="cmp-stat cmp-warn"><b>' + left + '</b> site group' + (left === 1 ? '' : 's') + ' to create by hand</div>'
          : '<div class="cmp-stat"><b>' + g.size + '</b> site groups created</div>') +
          (siteDone < siteTotal
            ? '<div class="cmp-stat cmp-warn"><b>' + (siteTotal - siteDone) + '</b> of ' + siteTotal + ' sites still to assign</div>'
            : '<div class="cmp-stat"><b>' + siteTotal + '</b> sites assigned to groups</div>');
      })();
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
    // Reference addresses still awaiting a look, and how many rows ride on each.
    const refPendingCodes = refAddrOutstanding().map(code => ({
      code,
      rows: xrefCodeCounts.get(code) || 0,
      reason: refEntryUsable(refAddrMap[code]) ? 'street/city split unverified' : 'no address found',
      address1: refAddrMap[code].address1 || '',
      city: refAddrMap[code].city || ''
    })).filter(x => x.rows > 0);
    // Flagged splits that were confirmed without anyone editing them.
    const refUnfixedCodes = refAddrConfirmedUnfixed().map(code => ({
      code,
      rows: xrefCodeCounts.get(code) || 0,
      address1: refAddrMap[code].address1 || '',
      city: refAddrMap[code].city || '',
      raw: refAddrMap[code].raw || ''
    })).filter(x => x.rows > 0);
    return { emptyReq, unknownByCol, nameCollisions, collisionSamples, refPendingCodes, refUnfixedCodes };
  }
  function hasBlockers(b) {
    return b && (b.emptyReq > 0 || b.unknownByCol.length || b.nameCollisions > 0 ||
      (b.refPendingCodes && b.refPendingCodes.length > 0) ||
      (b.refUnfixedCodes && b.refUnfixedCodes.length > 0));
  }

  function updateExportButton() {
    const btn = $('tss-export');
    btn.disabled = !(srcData && tplData && formattedRows && formattedRows.length);
    if (window.IMDebug) IMDebug.refresh('sites-standardize');
    const b = getExportBlockers();
    if (!b) { btn.title = ''; return; }
    const reasons = [];
    if (b.emptyReq > 0) reasons.push(b.emptyReq + ' empty required cells');
    b.unknownByCol.forEach(c => reasons.push(c.values.length + ' ' + c.header.replace(/\*$/, '') + ' value(s) not in dropdown'));
    if (b.nameCollisions > 0) reasons.push(b.nameCollisions + ' duplicate site names');
    (b.refPendingCodes || []).forEach(r => reasons.push('reference address for ' + r.code + ' unconfirmed (' + r.rows + ' rows)'));
    (b.refUnfixedCodes || []).forEach(r => reasons.push(r.code + ' address split flagged but never edited (' + r.rows + ' rows)'));
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
      (b.refUnfixedCodes || []).forEach(r => {
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ The "' + escHtml(r.code) +
          '" address split was flagged, confirmed, and never edited — ' + r.rows + ' row' + (r.rows === 1 ? '' : 's') +
          ' use it</div><ul class="ts-confirm-list"><li>From: <b>' + escHtml(r.raw) + '</b></li>' +
          '<li>Address1: <b>' + escHtml(r.address1 || '(empty)') + '</b></li>' +
          '<li>City: <b>' + escHtml(r.city || '(empty)') + '</b></li></ul>' +
          '<div class="ts-confirm-issue-body">The parser could not tell where the street ended. If that City looks wrong, ' +
          'fix it in <b>Reference Addresses</b> — it is wrong on all ' + r.rows + ' rows.</div></div>');
      });
      (b.refPendingCodes || []).forEach(r => {
        const split = r.address1 || r.city
          ? '<ul class="ts-confirm-list"><li>Address1: <b>' + escHtml(r.address1 || '(empty)') + '</b></li><li>City: <b>' + escHtml(r.city || '(empty)') + '</b></li></ul>'
          : '';
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ Reference address for "' + escHtml(r.code) +
          '" is unconfirmed — ' + r.rows + ' row' + (r.rows === 1 ? '' : 's') + ' use it</div>' + split +
          '<div class="ts-confirm-issue-body">' + escHtml(r.reason.charAt(0).toUpperCase() + r.reason.slice(1)) +
          '. Every one of these rows exports the same address, so a wrong split is wrong ' + r.rows +
          ' times. Check it in the <b>Reference Addresses</b> panel.</div></div>');
      });
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
  // The filled file is the TEMPLATE with a new <sheetData>, patched at the zip
  // level (see xlsx-template-writer.js). Round-tripping through SheetJS instead
  // rewrites every part from its own model: the dropdowns disappear and string
  // cells come out as t="str" — which OOXML defines as a cached formula result,
  // not literal text. PickTrace reads those back as nothing and rejects the
  // file with "unexpected column header", even though the grid is correct.
  function exportRowsForFile() {
    // Canonicalize dropdown-backed columns to the template's exact casing.
    const caseMaps = SITE_HEADERS.map(h => buildCaseMap(tplData.dropdowns.get(norm(h).replace(/\*$/, ''))));
    return formattedRows.filter((_, ri) => !excluded(ri)).map(row => {
      const r = row.slice();
      caseMaps.forEach((m, i) => {
        if (m.size && r[i]) { const k = String(r[i]).toUpperCase().trim(); if (m.has(k)) r[i] = m.get(k); }
      });
      // Everything is written as text, so Zip keeps any leading zero (07094).
      return r.map(v => (v == null ? '' : String(v)));
    });
  }
  function exportFileName() {
    const origName = (tplData.fileName || 'sites-template.xlsx').trim();
    const dotIdx = origName.lastIndexOf('.');
    const base = dotIdx > 0 ? origName.substring(0, dotIdx) : origName;
    const ext = dotIdx > 0 ? origName.substring(dotIdx) : '.xlsx';
    return base + ' — filled' + ext;
  }
  function doExport() {
    if (!formattedRows || !tplData) return;
    if (!window.IMXlsxTemplate) {
      alert('The template writer (xlsx-template-writer.js) did not load — cannot export a valid file.');
      return;
    }
    const rows = exportRowsForFile();
    IMXlsxTemplate.write({
      rawBuffer: tplData.rawBuffer,
      headers: SITE_HEADERS,
      rows: rows,
      fileName: exportFileName()
    }).catch(err => {
      alert('Could not write the filled template: ' + (err && err.message ? err.message : err));
    });
  }

  // ─── File load handlers ───
  function runFormat() {
    if (!srcData || !tplData) { alert('Upload source data + a personalized Sites template first.'); return; }
    mapping = autoMap(SITE_HEADERS, srcData.headers);
    manualFills = {}; removedRows = new Set(); cellOverrides = {}; nameOverrides = {};
    smartFixMap = {}; previewIssuesOnly = false; refAddrMap = {};
    collisionHistory = new Map(); collisionDismissed = new Set();
    if (!groupSrcOwn) resetGroupWork();
    scanReferenceAddresses();
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
  // ─── Site Groups tab: its own file handling ───
  function resetGroupWork() {
    completedGroups = new Set(); assignedSites = new Map();
    groupWorkSel = ''; groupWorkFilter = '';
    groupColPicked = false; groupColIdx = -1;
    groupNameColPicked = false; groupNameColIdx = -1;
    const fb = $('tss-groupwork-filter'); if (fb) fb.value = '';
  }
  function handleGroupSrcFile(file) {
    readSrcFile(file).then(data => {
      groupSrcOwn = data;
      resetGroupWork();
      $('tssg-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('tssg-src-meta').textContent = data.rows.length + ' rows · ' + data.headers.length + ' columns';
      renderGroupsTab();
    }).catch(err => {
      if (err && err.message === 'cancelled') return;
      alert('Failed to read that file: ' + (err && err.message ? err.message : err));
    });
  }
  function handleGroupExistingFile(file) {
    readSrcFile(file).then(data => {
      const nameI = data.headers.findIndex(h => {
        const k = norm(h).replace(/\*$/, '');
        return k === 'name' || k === 'site name' || k === 'site' || k === 'sites';
      });
      if (nameI < 0) { alert('That file needs a Name (or Site) column.'); return; }
      const keys = new Set();
      data.rows.forEach(r => {
        const n = String(r[nameI] == null ? '' : r[nameI]).trim();
        if (n) keys.add(nameKeyOf(n));
      });
      groupExistingOwn = { keys, fileName: file.name, count: keys.size };
      $('tssg-existing-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('tssg-existing-meta').textContent = keys.size + ' existing site' + (keys.size === 1 ? '' : 's') + ' indexed';
      renderGroupsTab();
    }).catch(err => {
      if (err && err.message === 'cancelled') return;
      alert('Failed to read that file: ' + (err && err.message ? err.message : err));
    });
  }
  function resetGroupTab() {
    groupSrcOwn = null; groupExistingOwn = null;
    resetGroupWork();
    $('tssg-src-name').textContent = 'No file selected';
    $('tssg-src-meta').textContent = '';
    $('tssg-existing-name').textContent = 'No file selected';
    $('tssg-existing-meta').textContent = 'Marks which sites are already in PickTrace.';
    ['tssg-src-file', 'tssg-existing-file'].forEach(id => { const el = $(id); if (el) el.value = ''; });
    renderGroupsTab();
  }

  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    manualFills = {}; removedRows = new Set(); existingData = null; existingRowIdxs = new Set();
    nameOverrides = {}; cellOverrides = {}; smartFixMap = {}; previewIssuesOnly = false; guessedCityRows = new Set(); collisionKeys = new Set(); existingTakenKeys = new Set();
    refAddrMap = {}; refAddrExtras = []; xrefCodeCounts = new Map(); xrefRows = new Map(); provMarks = {};
    collisionHistory = new Map(); collisionDismissed = new Set();
    // The Site Groups tab keeps its own Reset. Only clear its progress when it
    // is riding on this tab's data — not when it has a file of its own.
    if (!groupSrcOwn) resetGroupWork();
    $('tss-src-name').textContent = 'No file selected';
    $('tss-tpl-name').textContent = 'No file selected';
    $('tss-src-meta').textContent = ''; $('tss-tpl-meta').textContent = '';
    { const en = $('tss-existing-name'); if (en) en.textContent = 'No file selected'; }
    { const em = $('tss-existing-meta'); if (em) em.textContent = ''; }
    ['tss-src-file', 'tss-tpl-file', 'tss-existing-file'].forEach(id => { const el = $(id); if (el) el.value = ''; });
    ['tss-section-mapping', 'tss-section-refaddr', 'tss-section-smartfix', 'tss-section-emp-create',
     'tss-section-type-create', 'tss-section-grouplink', 'tss-section-required', 'tss-section-existing',
     'tss-section-collisions', 'tss-section-preview',
     'tss-section-sitegroups', 'tss-section-groupwork', 'tss-group-summary']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    { const ge = $('tss-group-empty'); if (ge) { ge.style.display = ''; ge.innerHTML =
        'Load your source file on the <b>Sites</b> tab first &mdash; site groups are read from it.'; } }
    $('tss-summary').style.display = 'none';
    $('tss-empty').style.display = '';
    { const rx = $('tss-refaddr-extra'); if (rx) { rx.style.display = 'none'; rx.innerHTML = ''; } }
    $('tss-run').disabled = true; $('tss-export').disabled = true;
    if (window.IMDebug) IMDebug.refresh('sites-standardize');
  }

  // ─── Debug dump ───
  // Everything that went in, every decision the module made on the way to the
  // export grid, and every item it surfaced for you — resolved or not.
  function collectDebug() {
    if (!srcData || !tplData || !formattedRows) return null;
    const a1 = SITE_HEADERS.indexOf('Address1*');
    const nameIdx = SITE_HEADERS.indexOf('Name*');
    const b = getExportBlockers() || {};
    const keptIdxs = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const keptRows = keptIdxs.map(i => formattedRows[i]);

    // How autoMap landed on each source column, re-derived for readability.
    const cols = SITE_HEADERS.map((h, i) => {
      const si = mapping[i] != null ? mapping[i] : -1;
      const hdr = si >= 0 ? norm(srcData.headers[si] || '') : '';
      const bn = norm(h), bnPlain = bn.replace(/\*$/, '').trim();
      let match = 'unmapped';
      if (si >= 0) {
        if (hdr === bn || hdr === bnPlain) match = 'exact header';
        else if ((SITE_ALIASES[bn] || SITE_ALIASES[bnPlain] || []).some(a => norm(a) === hdr)) match = 'alias';
        else if (normLoose(srcData.headers[si] || '') === normLoose(h.replace(/\*$/, ''))) match = 'loose header';
        else match = 'substring fallback';
      }
      let note = null;
      if (si < 0 && ['City*', 'State*', 'Zip*', 'Country*'].indexOf(h) >= 0) {
        note = 'No source column. Filled by the address parser or the reference-address cross-reference.';
      }
      return { index: i, header: h, required: /\*$/.test(h), srcIndex: si, match, note,
        sample: si >= 0 && srcData.rows[0] ? srcData.rows[0][si] : null };
    });

    const refTable = Object.keys(refAddrMap).sort().map(code => {
      const e = refAddrMap[code];
      return {
        code, rowsUsingCode: xrefCodeCounts.get(code) || 0,
        foundAt: e.origin ? ((e.origin.sheet ? e.origin.sheet + '!' : '') + e.origin.cell) : null,
        foundInColumn: e.origin ? e.origin.column : null,
        raw: e.raw || null,
        split: { address1: e.address1, city: e.city, state: e.state, zip: e.zip, country: e.country },
        splitConfidence: e.confidence,
        confirmedByUser: !!e.confirmed,
        editedByUser: !!e.edited,
        confirmedWithoutFixing: refConfirmedUnfixed(e),
        usable: refEntryUsable(e),
        outstanding: refNeedsAttention(e)
      };
    });

    const asked = [];
    if (refTable.length) {
      const unfixed = refAddrConfirmedUnfixed();
      asked.push(IMDebug.ask('reference-addresses', 'xref',
        'Reference addresses to confirm (Address column holds state codes, not addresses)', {
          count: refAddrOutstanding().length + unfixed.length,
          blocksExport: false,
          detail: refTable.length + ' state code(s) in use; each confirmed entry fills Address1/City/State/Zip/Country on every row carrying that code.' +
            (unfixed.length ? ' ' + unfixed.length + ' flagged split(s) (' + unfixed.join(', ') +
              ') were confirmed without being edited — see confirmedWithoutFixing.' : ''),
          items: refTable
        }));
    }
    const emptyReqCols = SITE_HEADERS.map((h, i) => {
      if (!/\*$/.test(h)) return null;
      let n = 0; keptIdxs.forEach(ri => { if (!formattedRows[ri][i]) n++; });
      return n ? { column: h, emptyRows: n } : null;
    }).filter(Boolean);
    asked.push(IMDebug.ask('empty-required', 'required', 'Required columns with empty cells', {
      count: emptyReqCols.length, blocksExport: true,
      detail: (b.emptyReq || 0) + ' empty required cells across ' + keptIdxs.length + ' exported rows.',
      items: emptyReqCols
    }));
    asked.push(IMDebug.ask('off-list-values', 'dropdown', 'Values not present in a template dropdown', {
      count: (b.unknownByCol || []).length, blocksExport: true,
      detail: 'Each of these must be created in PickTrace, or respelled to match the template.',
      items: (b.unknownByCol || []).map(c => ({ column: c.header, values: c.values }))
    }));
    const collGroups = [];
    collisionSuggestions().forEach((g, k) => {
      collGroups.push({
        name: formattedRows[g.rowIdxs[0]] ? formattedRows[g.rowIdxs[0]][nameIdx] : k,
        rowIndexes: g.rowIdxs,
        toldApartBy: g.disc ? { sourceColumn: g.disc.header, sourceColumnIndex: g.disc.colIdx, values: g.disc.values } : null,
        suggestedNames: g.disc ? g.rowIdxs.map(ri => g.names.get(ri)) : null,
        verdict: g.disc ? 'distinguishable — rename with the suggested suffix'
                        : 'identical in every source column — a true duplicate, drop one'
      });
    });
    asked.push(IMDebug.ask('name-collisions', 'collision', 'Sites sharing a Name', {
      count: b.nameCollisions || 0, blocksExport: true,
      detail: 'PickTrace silently drops one of each duplicate pair. ' +
        collGroups.filter(g => g.toldApartBy).length + ' of ' + collGroups.length +
        ' group(s) can be told apart by a source column and have a suggested rename.',
      items: collGroups
    }));
    const sfPending = computeSmartFixes();
    asked.push(IMDebug.ask('smart-fixes', 'smartfix', 'Off-list values awaiting a Smart Fix choice', {
      count: sfPending.length, blocksExport: false,
      items: sfPending.map(f => ({ column: f.header, value: f.from, rows: f.count, suggestion: f.suggestion || null }))
    }));
    const groups = collectSiteGroups();
    if (groups.size) {
      const pending = [...groups.keys()].filter(g => !completedGroups.has(g));
      asked.push(IMDebug.ask('site-groups', 'data', 'Site groups to create and assign by hand in PickTrace', {
        count: pending.length, blocksExport: false,
        detail: 'The Sites template has no Site Group column and groups cannot be bulk created — PickTrace files ' +
          'every uploaded site under a group of its own name, so this grouping is lost on import. Counting ' +
          siteGroupScope() + '.',
        items: [...groups.entries()]
          .sort((a, b) => b[1].sites.length - a[1].sites.length)
          .map(([g, info]) => {
            const p = groupProgress(g, info);
            const set = assignedFor(g);
            return { group: g, sites: info.sites.length, newSites: info.created,
              alreadyInPickTrace: info.existing,
              groupCreated: completedGroups.has(g),
              sitesAssigned: p.done, sitesStillToAssign: p.total - p.done,
              stillToAssign: info.sites.filter(s => !set.has(s)).slice(0, 500),
              siteNames: info.sites.slice(0, 500) };
          })
      }));
    }

    const fills = Object.keys(manualFills).map(i => ({
      column: SITE_HEADERS[+i], value: manualFills[i], scope: 'all exported rows',
      kind: 'column fill'
    })).concat(Object.keys(smartFixMap).map(k => {
      const sep = k.indexOf('||');
      return { column: SITE_HEADERS[+k.slice(0, sep)], from: k.slice(sep + 2), value: smartFixMap[k], kind: 'smart fix' };
    }));

    const prov = IMDebug.deriveProvenance({
      headers: SITE_HEADERS, rows: formattedRows, srcRows: srcData.rows,
      colToSrc: mapping, fills: manualFills, cellEdits: cellOverrides,
      nameEdits: nameOverrides, smartFixes: smartFixMap, marks: provMarks, nameCol: nameIdx
    });
    const out = IMDebug.output(SITE_HEADERS, keptRows, {
      note: 'Rows in export order. Dropdown-backed columns are re-cased to the template\'s exact spelling on export. ' +
        'exportedRowIndexes[n] is the formattedRows index behind exported row n — the same key used by provenance.codes, ' +
        'edits, and dropped.',
      exportFileName: (tplData.fileName || 'sites-template.xlsx').replace(/(\.[^.]+)$/, ' — filled$1')
    });
    out.exportedRowIndexes = keptIdxs;

    return {
      inputs: [
        IMDebug.file('source data (implementation workbook / export)', srcData, {
          stateCodesInAddressColumn: [...xrefCodeCounts.entries()].map(([c, n]) => ({ code: c, rows: n })),
          addressSourceColumn: mapping[a1] != null && mapping[a1] >= 0
            ? { index: mapping[a1], name: srcData.headers[mapping[a1]] } : null
        }),
        IMDebug.file('personalized bulk Sites template', tplData, {
          dropdowns: tplData.dropdowns ? [...tplData.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : [],
          sheetNames: tplData.sheetNames || null
        }),
        IMDebug.file('existing sites in PickTrace', existingData, {
          indexedNames: existingData ? existingData.count : null,
          note: 'Used to drop rows whose Name already exists and to flag taken names.'
        })
      ],
      mapping: IMDebug.mapping(cols, srcData.headers, {
        srcRows: srcData.rows,
        note: 'Columns with no source are filled by the address parser, the reference-address cross-reference, or a manual fill. ' +
              'sourceColumnsNotUsed includes unheadered columns that carry data — that is where reference addresses hide.'
      }),
      derived: [
        { kind: 'reference-address cross-reference',
          note: 'Address cells holding only a state code were replaced with that state\'s reference address.',
          rowsFilled: [...xrefRows.keys()].filter(ri => !excluded(ri)).length,
          table: refTable,
          alsoFoundButUnused: refAddrExtras.map(x => ({ code: x.code, raw: x.raw, foundAt: (x.origin.sheet ? x.origin.sheet + '!' : '') + x.origin.cell })) },
        { kind: 'site groups (not exported)',
          note: 'Read from a source column that has no destination in the Sites template. Listed for manual ' +
            'creation + assignment in PickTrace; never written to the file.',
          sourceColumn: groupColIdx >= 0 ? { index: groupColIdx, name: srcData.headers[groupColIdx] || '(no header)' } : null,
          columnChosenByUser: groupColPicked,
          scope: siteGroupScope(),
          groupCount: collectSiteGroups().size,
          markedDone: [...completedGroups] },
        { kind: 'combined-address parse',
          note: 'Rows whose Address cell held a full address were split into Address1/City/State/Zip/Country.',
          rowsWithGuessedCity: [...guessedCityRows].filter(ri => !excluded(ri)).length,
          guessedCitySamples: [...guessedCityRows].filter(ri => !excluded(ri)).slice(0, 20)
            .map(ri => ({ rowIndex: ri, name: formattedRows[ri][nameIdx],
              address1: formattedRows[ri][a1], city: formattedRows[ri][SITE_HEADERS.indexOf('City*')] })) }
      ],
      fills: fills,
      edits: {
        cells: Object.keys(cellOverrides).map(k => {
          const sep = k.indexOf('|');
          const ri = +k.slice(0, sep), ci = +k.slice(sep + 1);
          const si = mapping[ci];
          return { rowIndex: ri, column: SITE_HEADERS[ci], value: cellOverrides[k],
            sourceValue: (si != null && si >= 0 && srcData.rows[ri]) ? srcData.rows[ri][si] : null };
        }),
        names: Object.keys(nameOverrides).map(ri => ({
          rowIndex: +ri, value: nameOverrides[ri],
          sourceValue: (mapping[nameIdx] >= 0 && srcData.rows[+ri]) ? srcData.rows[+ri][mapping[nameIdx]] : null
        }))
      },
      dropped: {
        removedByHand: { count: removedRows.size, rowIndexes: [...removedRows].slice(0, 500) },
        alreadyInPickTrace: {
          count: existingRowIdxs.size,
          source: existingData ? existingData.fileName : null,
          names: [...existingRowIdxs].slice(0, 500).map(ri => formattedRows[ri][nameIdx])
        }
      },
      asked: asked,
      output: out,
      provenance: prov
    };
  }

  // ─── Init + domain toggle ───
  function init() {
    if (initialized) return;
    if (!$('tss-src-file')) return; // markup not present yet
    initialized = true;
    $('tss-src-file').addEventListener('change', e => { if (e.target.files[0]) handleSrcFile(e.target.files[0]); e.target.value = ''; });
    $('tss-tpl-file').addEventListener('change', e => { if (e.target.files[0]) handleTplFile(e.target.files[0]); e.target.value = ''; });
    { const ef = $('tss-existing-file'); if (ef) ef.addEventListener('change', e => { if (e.target.files[0]) handleExistingFile(e.target.files[0]); e.target.value = ''; }); }
    $('tss-run').addEventListener('click', runFormat);
    $('tss-export').addEventListener('click', async () => {
      const b = getExportBlockers();
      if (hasBlockers(b)) { const ok = await showExportConfirm(b); if (!ok) return; }
      doExport();
    });
    $('tss-reset').addEventListener('click', reset);
    { const sfa = $('tss-smartfix-apply-all'); if (sfa) sfa.addEventListener('click', applyAllSmartFixes); }
    { const rc = $('tss-refaddr-confirm-all'); if (rc) rc.addEventListener('click', confirmAllRefAddresses); }
    { const ca = $('tss-coll-apply-all'); if (ca) ca.addEventListener('click', applyAllCollisionSuggestions); }
    { const gc = $('tss-sitegroups-col'); if (gc) gc.addEventListener('change', e => {
        groupColIdx = +e.target.value;
        groupColPicked = true;
        // A different column means a different set of groups — the old
        // checklist no longer refers to anything.
        completedGroups = new Set(); assignedSites = new Map(); groupWorkSel = '';
        renderGroupsTab(); updateSummary();
      }); }
    { const nc = $('tss-sitegroups-namecol'); if (nc) nc.addEventListener('change', e => {
        groupNameColIdx = +e.target.value;
        groupNameColPicked = true;
        assignedSites = new Map(); groupWorkSel = '';
        renderGroupsTab(); updateSummary();
      }); }
    { const gf2 = $('tssg-src-file'); if (gf2) gf2.addEventListener('change', e => {
        if (e.target.files[0]) handleGroupSrcFile(e.target.files[0]); e.target.value = ''; }); }
    { const ge2 = $('tssg-existing-file'); if (ge2) ge2.addEventListener('change', e => {
        if (e.target.files[0]) handleGroupExistingFile(e.target.files[0]); e.target.value = ''; }); }
    { const gr2 = $('tssg-reset'); if (gr2) gr2.addEventListener('click', resetGroupTab); }
    { const us = $('tssg-use-sites'); if (us) us.addEventListener('click', () => {
        groupSrcOwn = null; groupExistingOwn = null;
        resetGroupWork();
        $('tssg-src-name').textContent = 'No file selected';
        $('tssg-src-meta').textContent = '';
        renderGroupsTab();
      }); }
    // ─── Site-group assignment worklist ───
    { const gs = $('tss-groupwork-sel'); if (gs) gs.addEventListener('change', e => {
        groupWorkSel = e.target.value; groupWorkFilter = '';
        const fb = $('tss-groupwork-filter'); if (fb) fb.value = '';
        renderGroupWork();
      }); }
    { const gf = $('tss-groupwork-filter'); if (gf) gf.addEventListener('input', e => {
        groupWorkFilter = e.target.value.trim(); renderGroupWork();
      }); }
    { const gl = $('tss-groupwork-list'); if (gl) gl.addEventListener('click', e => {
        const chip = e.target.closest('.tss-site-chip');
        if (chip) toggleAssigned(groupWorkSel, chip.dataset.site, chip);
      }); }
    { const ga = $('tss-groupwork-all'); if (ga) ga.addEventListener('click', () => {
        const info = collectSiteGroups().get(groupWorkSel);
        if (!info) return;
        const set = assignedFor(groupWorkSel);
        info.sites.forEach(s => set.add(s));
        renderGroupsTab(); updateSummary();
      }); }
    { const gn = $('tss-groupwork-none'); if (gn) gn.addEventListener('click', () => {
        assignedSites.set(groupWorkSel, new Set());
        renderGroupsTab(); updateSummary();
      }); }
    { const gr = $('tss-groupwork-copy-remaining'); if (gr) gr.addEventListener('click', () => {
        const info = collectSiteGroups().get(groupWorkSel);
        if (!info) return;
        const set = assignedFor(groupWorkSel);
        const left = info.sites.filter(s => !set.has(s));
        if (!left.length) { flashGroupNote('Nothing left — every site in "' + groupWorkSel + '" is ticked off.'); return; }
        copyText(left.join('\n'), 'Copied the ' + left.length + ' site' + (left.length === 1 ? '' : 's') +
          ' still to add to "' + groupWorkSel + '".');
      }); }
    if (window.IMDebug) {
      IMDebug.register('sites-standardize', {
        label: 'Sites Standardize',
        ready: () => !!(srcData && tplData && formattedRows),
        collect: collectDebug
      });
      IMDebug.wire('tss-debug', 'sites-standardize');
    }
    // Click-to-copy in the results area.
    $('tss-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      // Site Groups rows own their click (copy + tick off), so the generic
      // copy-the-cell handler must not fire on them too.
      if (!td || td.closest('#tss-section-required') || td.closest('#tss-section-sitegroups') ||
          e.target.closest('input, button, select')) return;
      const text = (td.textContent || '').trim();
      if (text && navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(() => td.classList.toggle('cmp-copied'), () => {});
    });
    // Locations / Sites is now the entity picker in the context strip, handled
    // by showBuild() in index.html's router — nothing to wire here.
  }

  window.tssInit = init;
  // showBuild() calls this when you switch to the Site Groups tab — its panels
  // live in a pane the Sites renders can't assume is visible.
  window.tssRenderGroups = function () { renderGroupsTab(); };
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
