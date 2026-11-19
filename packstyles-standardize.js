// ═══════════════════════════════════════════════════════════════════════
// Pack Styles Standardize — formats a source Excel/CSV (usually the "Pack
// Styles" tab of a client's Implementation Data Template) into a personalized
// PickTrace bulk PACK STYLES Create template. Sibling to sites-standardize.js;
// lives under Template Standardize's Pack Styles entity.
//
// The uploaded template is the schema: its DATA ENTRY header row and its
// DROP-DOWN INPUTS lists are read at upload, so nothing but the aliases below
// is hard-coded. The value rules are shared with Bulk Template Standardize
// (window.dtInternals, data-template.js) so both tools treat a cell the same:
//   splitWeight  "1 lb" → Weight 1 + Unit LB
//   autoMatch    "Conventional" → CONV, "Clamshell" → Clamshells
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const str = v => String(v == null ? '' : v).trim();
  // Lowercase alphanumerics only — drops the '*' required marker, spacing, punctuation.
  const normLoose = h => String(h == null ? '' : h).toLowerCase().replace(/[^a-z0-9]/g, '');
  const isReq = h => /\*\s*$/.test(h);
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');

  // Source-column aliases, keyed by the template column's loose name. The
  // Implementation tab's "Unit Type (for Piecework Pay)" holds the piecework
  // unit (Box), not the weight unit Unit takes (LB / OZ / G), so it is
  // deliberately absent — and there is no substring fallback to pick it up.
  const ALIASES = {
    packstylename: ['name', 'packstyle', 'packname', 'packagename'],
    croptype: ['crop', 'cropname', 'commodity'],
    count: ['packcount'],
    packtype: ['packagetype', 'container'],
    unit: ['weightunit', 'uom']
  };
  const NUMERIC = ['weight', 'count'];
  const PACK_SHEET = /^pack ?(styles?|sizes?|types?)$/i;
  const FOOTER_RE = /^\*?\s*add\s*\/\s*delete rows/i;
  const NA_RE = /^n\/?a$/i;
  const BUILD_CAP = 8;

  // ─── State ───
  let srcData = null;            // { headers, rows, fileName, sheetName, headerRow, naRows }
  let tplData = null;            // { headers, lists: Map(loose header → [values]), fileName, rawBuffer, sheetName, sheetNames }
  let mapping = {};              // template colIdx → source colIdx (-1 = unmapped)
  let formattedRows = null;      // rows in template column order, parallel to srcData.rows
  let manualFills = {};          // colIdx → value written into that column's blanks
  let removedRows = new Set();   // formattedRows indexes dropped from preview + export
  let cellOverrides = {};        // "ri|ci" → value (sticky preview edits)
  let smartFixMap = {};          // "ci||UPPER VALUE" → template value
  let pvMode = 'all';      // preview rows: 'all' | 'fix' | 'chg' | 'any' (IMWorkbench.filterBar)
  let pvColsOnly = false;  // preview: only the columns with a flag or a change
  let autoMatched = [];          // [{ ci, from, to, rows }] values matched to the one plausible entry
  let weightSplitRows = 0;       // rows whose Weight carried its unit
  let provMarks = {};            // "ri|ci" → provenance code, for the debug dump
  let initialized = false;

  const H = () => (tplData ? tplData.headers : []);
  const colOf = key => H().findIndex(h => normLoose(h) === key);
  const nameCol = () => colOf('packstylename');
  // The template's list for a column: an array (possibly EMPTY — the org has
  // no options yet), or null when the column has no dropdown at all.
  function listFor(ci) {
    if (!tplData) return null;
    const k = normLoose(H()[ci]);
    return tplData.lists.has(k) ? tplData.lists.get(k) : null;
  }
  function excluded(ri) { return removedRows.has(ri); }

  // ─── Header-row detection (mirrors sites-standardize) ───
  const HEADER_KEYWORDS = ['name', 'crop', 'pack', 'weight', 'size', 'unit', 'count', 'type',
    'upc', 'sku', 'label', 'method', 'gtin'];
  function detectHeaderRow(aoa) {
    let bestRow = 0, bestScore = -Infinity;
    for (let r = 0; r < Math.min(aoa.length, 20); r++) {
      const cells = (aoa[r] || []).filter(c => str(c) !== '');
      if (cells.length < 2) continue;
      let score = cells.length * 2;
      cells.forEach(c => {
        const s = str(c).toLowerCase();
        if (s.length < 40) score += 1;
        if (s.length > 80) score -= 4;
        HEADER_KEYWORDS.forEach(kw => { if (s.includes(kw)) score += 3; });
      });
      if (score > bestScore) { bestScore = score; bestRow = r; }
    }
    return bestRow;
  }
  function isInstructionRow(row) {
    const cells = (row || []).map(str).filter(Boolean);
    if (!cells.length) return true;
    if (cells.filter(c => c.length > 80).length >= Math.ceil(cells.length / 2)) return true;
    return cells.some(c => {
      const s = c.toLowerCase();
      return s.startsWith('required.') || s.startsWith('please ') || s.startsWith('optional.') ||
             s.startsWith('this field') || s.startsWith('this tab') || s.includes('used to collect');
    });
  }
  // Header, then data rows up to the template's "*add/delete rows" footer.
  // A placeholder row ("n/a" in every cell — the client has no pack styles)
  // is counted, not kept.
  function sliceSheet(aoa) {
    const hIdx = detectHeaderRow(aoa);
    const headers = (aoa[hIdx] || []).map(str);
    const rows = [];
    let naRows = 0;
    for (let r = hIdx + 1; r < aoa.length; r++) {
      const row = (aoa[r] || []).map(str);
      if (!row.some(Boolean)) continue;
      if (row.some(c => FOOTER_RE.test(c))) break;
      if (isInstructionRow(row)) continue;
      if (!row.some(c => c && !NA_RE.test(c))) { naRows++; continue; }
      rows.push(row);
    }
    return { headers, rows, headerRow: hIdx, naRows };
  }

  // ─── File parsing ───
  // Cell by cell: numbers raw, so a 12-digit UPC isn't reduced to the
  // "1.23457E+11" it displays as; dates as displayed, so a Size Excel took for
  // a date ("9/12") reads as the client typed it. (sheet_to_json can do one or
  // the other, and this SheetJS build ignores its rawNumbers option.)
  function sheetAoa(ws) {
    if (!ws || !ws['!ref']) return [];
    const range = XLSX.utils.decode_range(ws['!ref']);
    const aoa = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        row.push(!cell || cell.v == null ? '' : cell.t === 'd' ? (cell.w || String(cell.v)) : cell.v);
      }
      aoa.push(row);
    }
    return aoa;
  }
  function readSrcFile(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
          const sheets = wb.SheetNames.map(n => {
            const aoa = sheetAoa(wb.Sheets[n]);
            if (!aoa.some(row => (row || []).some(c => str(c)))) return null;
            const s = sliceSheet(aoa);
            if (!s.headers.filter(Boolean).length) return null;
            return Object.assign({ name: n }, s);
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({ headers: s.headers, rows: s.rows, headerRow: s.headerRow,
            naRows: s.naRows, fileName: file.name, sheetName: s.name });
          // A whole Implementation Data Template: its Pack Styles tab is the one.
          const pack = sheets.filter(s => PACK_SHEET.test(s.name.trim()));
          if (sheets.length === 1) return finish(sheets[0]);
          if (pack.length === 1) return finish(pack[0]);
          pickSheet(sheets, file.name).then(finish).catch(reject);
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
  }
  function pickSheet(sheets, fileName) {
    return new Promise((resolve, reject) => {
      const prior = $('tps-sheet-picker');
      if (prior) prior.remove();
      const overlay = document.createElement('div');
      overlay.id = 'tps-sheet-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner';
      inner.innerHTML = '<h3>Select sheet</h3><p>Pick the sheet with the <b>Pack Styles</b> data in <code>' +
        escHtml(fileName) + '</code>:</p><div class="cmp-org-list" id="tps-sheet-picker-list"></div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="tps-sheet-picker-cancel">Cancel</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
      const list = inner.querySelector('#tps-sheet-picker-list');
      sheets.forEach(s => {
        const item = document.createElement('div');
        item.className = 'cmp-org-list-item';
        item.innerHTML = '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
          '<span class="cmp-org-counts">' + s.rows.length + ' rows · ' + s.headers.length + ' cols</span>';
        item.addEventListener('click', () => { overlay.remove(); resolve(s); });
        list.appendChild(item);
      });
      const cancel = () => { overlay.remove(); reject(new Error('cancelled')); };
      inner.querySelector('#tps-sheet-picker-cancel').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
    });
  }

  function parseTemplate(buf, fileName) {
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
    const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
    const dropName = wb.SheetNames.find(n => /drop.?down/i.test(n));
    if (!wb.Sheets[dataName]) return null;
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[dataName], { header: 1, defval: '' });
    const headers = (aoa[0] || []).map(str);
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    const lists = new Map();
    if (dropName) {
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[dropName], { header: 1, defval: '' });
      (raw[0] || []).forEach((h, c) => {
        const k = normLoose(h);
        if (!k) return;
        const vals = [];
        raw.slice(1).forEach(r => { const v = str((r || [])[c]); if (v && vals.indexOf(v) < 0) vals.push(v); });
        // Empty lists are kept: a column the org has NO options for yet is a
        // finding, and it must be distinguishable from a column with no list.
        lists.set(k, vals);
      });
    }
    return { headers, lists, fileName, rawBuffer: buf, sheetName: dataName, sheetNames: wb.SheetNames };
  }
  function looksLikePackTemplate(headers) {
    const s = new Set((headers || []).map(normLoose));
    return s.has('packstylename') && s.has('croptype');
  }

  // ─── Auto-mapping ───
  // Exact (loose) names first for every column, then aliases — so an alias
  // can never take a source column another template column names exactly.
  function autoMap(tplHeaders, srcHeaders) {
    const idx = new Map();
    srcHeaders.forEach((h, i) => { const k = normLoose(h); if (k && !idx.has(k)) idx.set(k, i); });
    const out = {}, taken = new Set();
    tplHeaders.forEach((h, ci) => {
      const k = normLoose(h);
      if (idx.has(k)) { out[ci] = idx.get(k); taken.add(idx.get(k)); } else out[ci] = -1;
    });
    tplHeaders.forEach((h, ci) => {
      if (out[ci] >= 0) return;
      const hit = (ALIASES[normLoose(h)] || []).find(a => idx.has(a) && !taken.has(idx.get(a)));
      if (hit) { out[ci] = idx.get(hit); taken.add(idx.get(hit)); }
    });
    return out;
  }
  function matchKind(ci) {
    const si = mapping[ci] != null ? mapping[ci] : -1;
    if (si < 0) return 'unmapped';
    const k = normLoose(H()[ci]), sk = normLoose(srcData.headers[si]);
    if (k === sk) return 'exact header';
    if ((ALIASES[k] || []).indexOf(sk) >= 0) return 'alias';
    return 'chosen by you';
  }

  // ─── Build formatted rows ───
  // Order: source → weight split / number cleanup → column fills (blanks) →
  // Smart Fixes → dropdown pass (exact casing, then the one plausible entry)
  // → hand edits, which always win.
  function rebuildFormattedRows() {
    if (!srcData || !tplData) { formattedRows = null; return; }
    const heads = H();
    const dt = window.dtInternals || null;
    const wi = colOf('weight'), ui = colOf('unit');
    const numIdx = NUMERIC.map(colOf).filter(i => i >= 0);
    provMarks = {};
    weightSplitRows = 0;
    formattedRows = srcData.rows.map((sr, ri) => {
      const row = heads.map((h, ci) => {
        const si = mapping[ci];
        return si != null && si >= 0 ? str(sr[si]) : '';
      });
      if (wi >= 0 && dt && dt.splitWeight) {
        const w = dt.splitWeight(row[wi]);
        if (w) {
          row[wi] = w.weight;
          provMarks[ri + '|' + wi] = 'parse:weight';
          if (ui >= 0 && !row[ui]) { row[ui] = w.unit; provMarks[ri + '|' + ui] = 'parse:weight'; }
          weightSplitRows++;
        }
      }
      numIdx.forEach(i => { if (/^-?\d+\.0+$/.test(row[i])) row[i] = String(parseFloat(row[i])); });
      return row;
    });

    Object.keys(manualFills).forEach(k => {
      const ci = +k;
      formattedRows.forEach(r => { if (!r[ci]) r[ci] = manualFills[k]; });
    });

    formattedRows.forEach(r => {
      r.forEach((v, ci) => {
        if (!v) return;
        const hit = smartFixMap[ci + '||' + v.toUpperCase()];
        if (hit != null) r[ci] = hit;
      });
    });

    const tally = new Map();
    heads.forEach((h, ci) => {
      const list = listFor(ci);
      if (!list || !list.length) return;
      const caseMap = new Map(list.map(v => [v.toUpperCase(), v]));
      formattedRows.forEach(r => {
        const v = r[ci];
        if (!v) return;
        if (caseMap.has(v.toUpperCase())) { r[ci] = caseMap.get(v.toUpperCase()); return; }
        const to = dt && dt.autoMatch ? dt.autoMatch(h, v, list) : '';
        if (!to) return;
        r[ci] = to;
        const key = ci + '\u0000' + v + '\u0000' + to;
        tally.set(key, (tally.get(key) || 0) + 1);
      });
    });
    autoMatched = [...tally.entries()].map(([k, n]) => {
      const p = k.split('\u0000');
      return { ci: +p[0], from: p[1], to: p[2], rows: n };
    });

    Object.keys(cellOverrides).forEach(k => {
      const sep = k.indexOf('|');
      const ri = +k.slice(0, sep), ci = +k.slice(sep + 1);
      if (formattedRows[ri]) formattedRows[ri][ci] = cellOverrides[k];
    });
  }

  // ─── Findings ───
  function emptyRequired() {
    const out = [];
    H().forEach((h, ci) => {
      if (!isReq(h)) return;
      let n = 0;
      formattedRows.forEach((r, ri) => { if (!excluded(ri) && !r[ci]) n++; });
      if (n) out.push({ ci, header: h, empty: n });
    });
    return out;
  }
  // Values not in a column's (non-empty) list — respell via Smart Fixes, or create.
  function offList() {
    const out = [];
    H().forEach((h, ci) => {
      const list = listFor(ci);
      if (!list || !list.length) return;
      const up = new Set(list.map(v => v.toUpperCase()));
      const m = new Map();
      formattedRows.forEach((r, ri) => {
        const v = r[ci];
        if (excluded(ri) || !v || up.has(v.toUpperCase())) return;
        m.set(v, (m.get(v) || 0) + 1);
      });
      if (m.size) out.push({ ci, header: h, values: [...m.entries()].map(([value, count]) => ({ value, count })), allowed: list });
    });
    return out;
  }
  // What has to exist in PickTrace before this file will import: crop types
  // the org doesn't have, and any value in a column whose list is still empty
  // (Size, Product ID … — the org has none yet). Case variants merge.
  function toCreate() {
    const out = [];
    H().forEach((h, ci) => {
      const list = listFor(ci);
      if (!list) return;
      if (list.length && normLoose(h) !== 'croptype') return;
      const up = new Set(list.map(v => v.toUpperCase()));
      const m = new Map();
      formattedRows.forEach((r, ri) => {
        const v = r[ci];
        if (excluded(ri) || !v || up.has(v.toUpperCase())) return;
        const k = v.toUpperCase();
        const e = m.get(k) || { value: v, count: 0 };
        e.count++;
        m.set(k, e);
      });
      if (m.size) out.push({ ci, header: h, emptyList: !list.length, values: [...m.values()] });
    });
    return out;
  }
  // Pack Style Names used by more than one kept row (case-insensitive).
  function duplicates() {
    const ni = nameCol();
    const groups = new Map();
    if (ni < 0) return groups;
    formattedRows.forEach((r, ri) => {
      if (excluded(ri) || !r[ni]) return;
      const k = r[ni].toUpperCase().replace(/\s+/g, ' ');
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(ri);
    });
    [...groups.keys()].forEach(k => { if (groups.get(k).length < 2) groups.delete(k); });
    return groups;
  }

  // ─── Mapping panel ───
  function renderMapping() {
    $('tps-section-mapping').style.display = '';
    const opts = ['<option value="-1">— (leave empty) —</option>']
      .concat(srcData.headers.map((h, i) => h ? '<option value="' + i + '">' + escHtml(h) + '</option>' : '')).join('');
    let html = '<thead><tr><th>Template column</th><th>Source column</th><th>Sample value</th></tr></thead><tbody>';
    H().forEach((h, ci) => {
      const sel = mapping[ci] != null ? mapping[ci] : -1;
      const sample = sel >= 0 ? (srcData.rows.map(r => str(r[sel])).find(Boolean) || '') : '';
      html += '<tr><td>' + (isReq(h) ? '<b>' + escHtml(h) + '</b>' : escHtml(h)) + '</td>' +
        '<td><select class="tps-map-select input-field" data-ci="' + ci + '" style="min-width:220px;">' +
        opts.replace('value="' + sel + '"', 'value="' + sel + '" selected') + '</select></td>' +
        '<td><span class="text-muted small">' + escHtml(sample.slice(0, 60)) + '</span></td></tr>';
    });
    $('tps-mapping-table').innerHTML = html + '</tbody>';
  }

  // ─── Build-in-PickTrace banner (same look as Bulk Template Standardize) ───
  function renderBuild() {
    const el = $('tps-build');
    const list = toCreate();
    if (!list.length) { el.style.display = 'none'; el.innerHTML = ''; return; }
    let vals = 0;
    list.forEach(c => { vals += c.values.length; });
    let html = '<div class="dt-build-head"><div class="dt-build-title">Build these in PickTrace first</div>' +
      '<div class="dt-build-sub"><b>' + plural(vals, 'value') + '</b> across <b>' + plural(list.length, 'column') +
      '</b> have no match in your template, so rows using them are rejected on import. Create them in PickTrace, ' +
      'then re-download the template — or, if a crop is only spelled differently, map it in <b>Smart Fixes</b>.</div></div>' +
      '<div class="dt-build-cols">';
    list.forEach(c => {
      const vs = c.values.slice().sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
      const hidden = vs.length - BUILD_CAP;
      const what = normLoose(c.header) === 'croptype' ? plural(vs.length, 'crop') : plural(vs.length, 'value');
      html += '<div class="dt-build-col"><div class="dt-build-col-head"><span class="dt-build-col-name">' +
        escHtml(c.header) + (c.emptyList ? ' <span class="dt-build-tag">no options yet</span>' : '') + '</span>' +
        '<span class="dt-build-col-n">' + what + '</span></div><ul class="dt-build-vals">' +
        vs.map((v, i) => '<li' + (i >= BUILD_CAP ? ' class="dt-build-extra"' : '') + '>' +
          '<span class="dt-build-val">' + escHtml(v.value) + '</span>' +
          '<span class="dt-build-rows">' + plural(v.count, 'row') + '</span></li>').join('') + '</ul>' +
        (hidden > 0 ? '<button type="button" class="dt-build-toggle" data-more="' + hidden + '">Show ' + hidden + ' more</button>' : '') +
        '</div>';
    });
    el.innerHTML = html + '</div>';
    el.style.display = '';
  }

  // ─── Matched to the dropdown (informational) ───
  function renderMatched() {
    const sec = $('tps-section-matched');
    if (!autoMatched.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('tps-matched-title').textContent = 'Matched to the Dropdown (' + autoMatched.length + ')';
    $('tps-matched-table').innerHTML = '<thead><tr><th>Column</th><th>In your data</th><th>Written as</th><th>Rows</th></tr></thead><tbody>' +
      autoMatched.map(m => '<tr><td>' + escHtml(H()[m.ci]) + '</td><td>' + escHtml(m.from) + '</td><td><b>' +
        escHtml(m.to) + '</b></td><td>' + m.rows + '</td></tr>').join('') + '</tbody>';
  }

  // ─── Smart Fixes ───
  function renderSmartFixes() {
    const sec = $('tps-section-smartfix');
    const off = offList();
    if (!off.length) { sec.style.display = 'none'; $('tps-smartfix-table').innerHTML = ''; return; }
    sec.style.display = '';
    const dt = window.dtInternals || null;
    let n = 0, rowsHit = 0;
    let html = '<thead><tr><th>Column</th><th>Value in your data</th><th>Set to</th><th>Rows</th><th></th></tr></thead><tbody>';
    off.forEach(c => {
      c.values.forEach(v => {
        n++; rowsHit += v.count;
        const sug = dt && dt.suggestFor ? dt.suggestFor(v.value, c.allowed) : '';
        const opts = '<option value="">— pick a value —</option>' + c.allowed.map(o =>
          '<option' + (o === sug ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('');
        html += '<tr><td>' + escHtml(c.header) + '</td>' +
          '<td><span style="color:#7c2d12;background:#fef3c7;padding:1px 6px;border-radius:3px;">' + escHtml(v.value) + '</span></td>' +
          '<td><select class="tps-sf-pick input-field" style="min-width:200px;">' + opts + '</select>' +
          (sug ? ' <span class="text-muted small">suggested</span>' : ' <span style="color:#b45309;font-weight:600;" class="small">needs a choice</span>') + '</td>' +
          '<td>' + v.count + '</td>' +
          '<td><button class="btn btn-primary btn-sm tps-sf-apply" data-ci="' + c.ci + '" data-from="' + escHtml(v.value) + '">Apply</button></td></tr>';
      });
    });
    $('tps-smartfix-title').textContent = 'Smart Fixes (' + plural(n, 'value') + ', ' + plural(rowsHit, 'row') + ')';
    $('tps-smartfix-table').innerHTML = html + '</tbody>';
  }
  function applySmartFix(ci, from, to) {
    smartFixMap[ci + '||' + String(from).toUpperCase()] = to;
    refresh();
  }

  // ─── Empty required columns ───
  function renderRequired() {
    const sec = $('tps-section-required');
    const reqCols = H().map((h, ci) => ({ h, ci })).filter(x => isReq(x.h));
    if (!reqCols.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const empties = new Map(emptyRequired().map(e => [e.ci, e.empty]));
    const row = c => {
      const list = listFor(c.ci) || [];
      const cur = manualFills[c.ci] || '';
      const opts = list.length
        ? '<option value="">— pick —</option>' + list.map(o => '<option' + (o === cur ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const n = empties.get(c.ci) || 0;
      const status = n
        ? '<span class="dt-bad">⚠ ' + n + ' empty</span>'
        : '<span class="dt-ok">✓ filled</span>' + (cur ? ' <span class="text-muted small">(blanks filled with ' + escHtml(cur) + ')</span>' : '');
      return '<tr><td><b>' + escHtml(c.h) + '</b></td><td>' + status + '</td>' +
        '<td><select class="tps-req-pick input-field" style="min-width:180px;">' + opts + '</select></td>' +
        '<td><input type="text" class="tps-req-text input-field" placeholder="Manual override" style="width:180px;"></td>' +
        '<td><button class="btn btn-primary btn-sm tps-req-apply" data-ci="' + c.ci + '" title="Write into every empty cell of this column">Fill blanks</button></td>' +
        '<td>' + (cur ? '<button class="btn btn-ghost btn-sm tps-req-undo" data-ci="' + c.ci + '">Undo fill</button>' : '') + '</td></tr>';
    };
    $('tps-required-table').innerHTML = '<thead><tr><th>Column</th><th>Status</th><th>Pick from dropdown</th>' +
      '<th>Manual override</th><th></th><th></th></tr></thead><tbody>' + reqCols.map(row).join('') + '</tbody>';
  }

  // ─── Duplicate names ───
  function renderDuplicates() {
    const sec = $('tps-section-dups');
    const groups = duplicates();
    if (!groups.size) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const ni = nameCol(), ci = colOf('croptype');
    let rowsN = 0;
    groups.forEach(g => { rowsN += g.length; });
    $('tps-dups-title').textContent = 'Duplicate Pack Style Names (' + plural(groups.size, 'name') + ', ' + plural(rowsN, 'row') + ')';
    let html = '<thead><tr><th>Pack Style Name</th><th>Rows</th><th>Crop Type of each</th></tr></thead><tbody>';
    groups.forEach(g => {
      html += '<tr><td class="ts-copy-cell">' + escHtml(formattedRows[g[0]][ni]) + '</td><td>' + g.length + '</td><td>' +
        escHtml(g.map(ri => (ci >= 0 && formattedRows[ri][ci]) || '(blank)').join(' · ')) + '</td></tr>';
    });
    $('tps-dups-table').innerHTML = html + '</tbody>';
  }

  // ─── Editable preview (mirrors sites-standardize's grid) ───
  function renderPreview() {
    const sec = $('tps-section-preview');
    const tbl = $('tps-preview-table');
    sec.style.display = '';
    const heads = H();
    const ni = nameCol();
    const dupRows = new Set();
    duplicates().forEach(g => g.forEach(ri => dupRows.add(ri)));
    const cols = heads.map((h, ci) => {
      const list = listFor(ci);
      return { list, up: list ? new Set(list.map(v => v.toUpperCase())) : null };
    });
    // Why a cell is flagged, or '' — red for a hard stop, amber for a check.
    const flagOf = (ri, ci) => {
      const h = heads[ci], v = formattedRows[ri][ci] || '', c = cols[ci];
      if (isReq(h) && !v) return { bad: true, why: 'Required — must be filled before export' };
      if (ci === ni && dupRows.has(ri)) return { bad: true, why: 'Another row uses this Pack Style Name — rename or remove one.' };
      if (c.list && v && !c.list.length) return { bad: false, why: 'Your org has no ' + h + ' options yet — create "' + v + '" in PickTrace.' };
      if (c.list && v && !c.up.has(v.toUpperCase())) return { bad: false, why: '"' + v + '" isn’t in the template dropdown for ' + h + '.' };
      return null;
    };
    // Every kept row with its flagged (f) and changed (c) columns, for the
    // filter's counts and "only columns with something to see".
    const kept = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const info = new Map(), counts = { all: kept.length, fix: 0, chg: 0, any: 0 };
    kept.forEach(ri => {
      const f = [], c = [];
      heads.forEach((h, ci) => { if (flagOf(ri, ci)) f.push(ci); if (changeOf(ri, ci)) c.push(ci); });
      info.set(ri, { f, c });
      if (f.length) counts.fix++;
      if (c.length) counts.chg++;
      if (f.length || c.length) counts.any++;
    });
    const visible = kept.filter(ri => IMWorkbench.rowMatch(pvMode, info.get(ri).f.length, info.get(ri).c.length));
    let show = heads.map((_, ci) => ci);
    if (pvColsOnly) {
      const keep = new Set([ni].filter(i => i >= 0));
      visible.forEach(ri => { const x = info.get(ri); x.f.forEach(i => keep.add(i)); x.c.forEach(i => keep.add(i)); });
      show = show.filter(ci => keep.has(ci));
    }

    const head = '<thead><tr><th style="width:28px;">&nbsp;</th>' + show.map(ci =>
      '<th' + (isReq(heads[ci]) ? ' style="color:#dc2626;"' : '') + '>' + escHtml(heads[ci]) +
      (cols[ci].list && cols[ci].list.length ? ' <span class="ts-col-listed">&#9662;</span>' : '') + '</th>').join('') + '</tr></thead>';
    const rowHtml = ri => {
      const row = formattedRows[ri];
      let html = '<tr><td style="text-align:center;padding:0;"><button class="tps-row-remove" data-ri="' + ri +
        '" title="Remove this row from preview + export" style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button></td>';
      show.forEach(ci => {
        const v = row[ci] || '';
        const f = flagOf(ri, ci);
        const bg = f ? (f.bad ? '#fee2e2' : '#fef3c7') : '';
        const tdStyle = bg ? ' style="background:' + bg + ';"' : '';
        const fStyle = bg ? ' style="color:' + (f.bad ? '#7f1d1d' : '#7c2d12') + ';"' : '';
        const note = changeOf(ri, ci);
        const tip = [f ? f.why : '', note].filter(Boolean).join('\n');
        const attrs = ' data-ri="' + ri + '" data-ci="' + ci + '"' + (tip ? ' title="' + escHtml(tip) + '"' : '') + fStyle;
        const list = cols[ci].list;
        // Dropdown cells are a light button until clicked (IMWorkbench.wirePicks);
        // text cells are text until clicked (IMWorkbench.wireTextCells).
        html += '<td class="ts-cell' + (note ? ' is-chg' : '') + '"' + tdStyle + '>' +
          (list && list.length ? IMWorkbench.pickButton(v, attrs) : IMWorkbench.textCell(v, attrs)) + '</td>';
      });
      return html + '</tr>';
    };
    IMWorkbench.rows(tbl, { head: head, ids: visible, row: rowHtml, cols: show.length + 1,
      keys: [null].concat(show.map(ci => heads[ci])), cell: (ri, p) => formattedRows[ri][show[p - 1]],
      widths: [28].concat(show.map(ci => IMWorkbench.fitWidth(heads[ci], visible.map(ri => formattedRows[ri][ci])))),
      empty: pvMode === 'fix' ? 'No row needs a fix ✓' : 'Nothing here has been changed.' });
    IMWorkbench.wirePicks(tbl, ci => listFor(ci), '__tps_custom__');
    IMWorkbench.wireTextCells(tbl);

    let t = IMWorkbench.filterBar({ mode: pvMode, counts: counts, colsOnly: pvColsOnly, shownCols: show.length, totalCols: heads.length });
    t += '<span class="wb-filter-note">' + (visible.length === kept.length ? 'All ' + kept.length : visible.length + ' of ' + kept.length) +
      ' rows · every cell is editable — <span class="ts-col-listed">&#9662;</span> columns are template dropdowns. Red = empty required or a duplicate name; amber = not in the dropdown, or nothing to pick from yet. Click <b>×</b> to drop a row.</span>';
    if (removedRows.size) t += ' <button class="btn btn-ghost btn-sm" id="tps-restore-rows">Restore ' + plural(removedRows.size, 'removed row') + '</button>';
    $('tps-preview-hint').innerHTML = t;
  }

  // "Was … → now … · why" when the cell differs from its source cell, for the
  // preview's change tick (see IMWorkbench.changeNote).
  function changeOf(ri, ci) {
    const si = mapping[ci];
    const was = si != null && si >= 0 ? str(srcData.rows[ri][si]) : '';
    const now = formattedRows[ri][ci] || '';
    let why = 'Cleaned up by the tool';
    if (cellOverrides[ri + '|' + ci] != null) why = 'Your edit';
    else if (!was && manualFills[ci] === now) why = 'Column fill you applied';
    else if (smartFixMap[ci + '||' + was.toUpperCase()] === now) why = 'Smart Fix you applied';
    else if (provMarks[ri + '|' + ci] === 'parse:weight') why = 'Weight split from its unit';
    else if (autoMatched.some(m => m.ci === ci && m.from === was && m.to === now)) why = 'Matched to the dropdown';
    return IMWorkbench.changeNote(was, now, why);
  }

  function setCell(ri, ci, val) {
    cellOverrides[ri + '|' + ci] = val == null ? '' : String(val);
    refresh();
  }

  // ─── Summary + export gating ───
  function exportBlockers() {
    if (!formattedRows) return null;
    let emptyReq = 0;
    emptyRequired().forEach(e => { emptyReq += e.empty; });
    let dupRows = 0;
    const dupSamples = [];
    duplicates().forEach(g => { dupRows += g.length; dupSamples.push(formattedRows[g[0]][nameCol()] + ' (×' + g.length + ')'); });
    return { emptyReq, emptyCols: emptyRequired(), off: offList(), build: toCreate(), dupRows, dupSamples };
  }
  function hasBlockers(b) {
    return !!(b && (b.emptyReq || b.off.length || b.build.length || b.dupRows));
  }
  function updateSummary() {
    const sum = $('tps-summary');
    const kept = formattedRows.filter((_, i) => !excluded(i)).length;
    const b = exportBlockers();
    let offN = 0;
    b.off.forEach(c => { offN += c.values.length; });
    let buildN = 0;
    b.build.forEach(c => { buildN += c.values.length; });
    sum.innerHTML =
      '<span class="cmp-stat"><b>' + kept + '</b> pack style' + (kept === 1 ? '' : 's') + (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed)</span>' : '') + '</span>' +
      (srcData.naRows ? '<span class="cmp-stat">' + plural(srcData.naRows, 'n/a placeholder row') + ' skipped</span>' : '') +
      (weightSplitRows ? '<span class="cmp-stat">' + plural(weightSplitRows, 'weight') + ' split into Weight + Unit</span>' : '') +
      (autoMatched.length ? '<span class="cmp-stat">' + plural(autoMatched.length, 'value') + ' matched to the dropdown</span>' : '') +
      (b.emptyReq ? '<span class="cmp-stat cmp-warn">⚠ ' + b.emptyReq + ' empty required cells</span>' : '') +
      (offN ? '<span class="cmp-stat cmp-warn">⚠ ' + plural(offN, 'value') + ' not in a dropdown</span>' : '') +
      (buildN ? '<span class="cmp-stat cmp-warn">⚠ ' + buildN + ' to build in PickTrace</span>' : '') +
      (b.dupRows ? '<span class="cmp-stat cmp-warn">⚠ ' + b.dupRows + ' rows share a name</span>' : '');
    sum.style.display = '';
    const btn = $('tps-export');
    btn.disabled = !kept;
    const reasons = [];
    if (b.emptyReq) reasons.push(b.emptyReq + ' empty required cells');
    if (offN) reasons.push(offN + ' value(s) not in a dropdown');
    if (buildN) reasons.push(buildN + ' value(s) to build in PickTrace');
    if (b.dupRows) reasons.push(b.dupRows + ' rows sharing a name');
    btn.title = reasons.length ? 'Click to review before exporting: ' + reasons.join(', ') + '.' : 'Ready to export.';
  }

  function showExportConfirm(b) {
    return new Promise(resolve => {
      const prior = $('tps-export-confirm');
      if (prior) prior.remove();
      const overlay = document.createElement('div');
      overlay.id = 'tps-export-confirm';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const list = items => '<ul class="ts-confirm-list">' + items.slice(0, 8).map(x => '<li>' + escHtml(x) + '</li>').join('') +
        (items.length > 8 ? '<li class="ts-confirm-more">… and ' + (items.length - 8) + ' more</li>' : '') + '</ul>';
      const issue = (head, body) => '<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + head + '</div>' + body + '</div>';
      const sections = [];
      if (b.emptyReq) sections.push(issue(b.emptyReq + ' empty required cells',
        list(b.emptyCols.map(e => e.header + ' — ' + e.empty + ' empty')) +
        '<div class="ts-confirm-issue-body">PickTrace rejects pack styles missing a required (*) value. Fill them in <b>Empty Required Columns</b>.</div>'));
      b.build.forEach(c => sections.push(issue(plural(c.values.length, c.header.replace(/\*$/, '') + ' value') + ' to build in PickTrace',
        list(c.values.map(v => v.value)) + '<div class="ts-confirm-issue-body">Create them in PickTrace, then re-download the template.</div>')));
      b.off.filter(c => normLoose(c.header) !== 'croptype').forEach(c => sections.push(issue(
        plural(c.values.length, c.header.replace(/\*$/, '') + ' value') + ' not in your template dropdown',
        list(c.values.map(v => v.value)) + '<div class="ts-confirm-issue-body">These are written as they are and PickTrace will reject those rows. Pick a listed value in <b>Smart Fixes</b>.</div>')));
      if (b.dupRows) sections.push(issue(b.dupRows + ' rows share a Pack Style Name',
        list(b.dupSamples) + '<div class="ts-confirm-issue-body">Rename or remove duplicates so each pack style is created once.</div>'));
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner ts-confirm-modal';
      inner.innerHTML = '<h3>Export anyway?</h3><div class="ts-confirm-issues">' + sections.join('') + '</div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="tps-conf-cancel">Cancel</button>' +
        '<button class="btn btn-primary" id="tps-conf-ok">Export anyway</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
      const close = v => { overlay.remove(); resolve(v); };
      inner.querySelector('#tps-conf-ok').addEventListener('click', () => close(true));
      inner.querySelector('#tps-conf-cancel').addEventListener('click', () => close(false));
      overlay.addEventListener('click', e => { if (e.target === overlay) close(false); });
    });
  }

  // ─── Export ───
  // Patched into the template at the zip level (xlsx-template-writer.js), so
  // DATA ENTRY's dropdowns and the header row survive byte-for-byte.
  // "Acme Farms - Pack Styles - 2026-10-07.xlsx" (see IMOutName in utils.js).
  function exportFileName() { return IMOutName.fileName('Pack Styles'); }
  function doExport() {
    if (!window.IMXlsxTemplate) {
      alert('The template writer (xlsx-template-writer.js) did not load — cannot export a valid file.');
      return;
    }
    const rows = formattedRows.filter((_, ri) => !excluded(ri)).map(r => r.map(v => (v == null ? '' : String(v))));
    IMXlsxTemplate.write({ rawBuffer: tplData.rawBuffer, headers: H(), rows: rows, fileName: exportFileName() })
      .catch(err => alert('Could not write the filled template: ' + (err && err.message ? err.message : err)));
  }

  // ─── Render all ───
  function refresh() {
    if (!srcData || !tplData) return;
    rebuildFormattedRows();
    renderBuild(); renderMatched(); renderSmartFixes(); renderRequired(); renderDuplicates(); renderPreview();
    updateSummary();
    $('tps-empty').style.display = 'none';
    if (wb) wb.refresh();
    if (window.IMDebug) IMDebug.refresh('packstyles-standardize');
  }

  // ─── Workbench: one button per section (workbench.js) ───
  let wb = null;
  function mountWorkbench() {
    if (!window.IMWorkbench) return;
    const values = list => list.reduce((n, c) => n + c.values.length, 0);
    wb = IMWorkbench.mount({
      pane: 'ts-pack-wrap', results: 'tps-results', preview: 'tps-section-preview',
      ready: () => !!(srcData && tplData && formattedRows),
      sections: [
        { key: 'required', group: 'fix', unit: 'column', name: 'Required blanks', desc: 'Fill required (*) columns the source left empty',
          ids: ['tps-section-required'], count: () => emptyRequired().length },
        { key: 'dropdown', group: 'fix', unit: 'value', name: 'Not in dropdown', desc: 'Pick a listed value for anything PickTrace would reject',
          ids: ['tps-section-smartfix'], count: () => values(offList()) },
        { key: 'build', group: 'fix', unit: 'value', name: 'Build in PickTrace', desc: 'Crop types and options to create, then re-download the template',
          ids: ['tps-build'], count: () => values(toCreate()) },
        { key: 'dups', group: 'fix', unit: 'name', name: 'Duplicate names', desc: 'Pack Style Names used on more than one row',
          ids: ['tps-section-dups'], count: () => duplicates().size },
        { key: 'mapping', group: 'check', unit: 'column', name: 'Column mapping', desc: 'Which source column fills each template column',
          ids: ['tps-section-mapping'], count: () => H().filter((_, ci) => mapping[ci] >= 0).length },
        { key: 'matched', group: 'check', unit: 'value', name: 'Matched to dropdown', desc: 'Values written as their dropdown entry (Conventional → CONV)',
          ids: ['tps-section-matched'], count: () => autoMatched.length }
      ]
    });
  }
  function runFormat() {
    if (!srcData || !tplData) { alert('Upload source data + a personalized Pack Styles template first.'); return; }
    mapping = autoMap(H(), srcData.headers);
    manualFills = {}; removedRows = new Set(); cellOverrides = {}; smartFixMap = {}; pvMode = 'all'; pvColsOnly = false;
    renderMapping();
    refresh();
  }

  // ─── File load handlers ───
  function handleSrcFile(file) {
    readSrcFile(file).then(data => {
      srcData = data;
      IMOutName.suggestClient(file.name);
      $('tps-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      $('tps-src-meta').textContent = plural(data.rows.length, 'row') + ' · ' + plural(data.headers.filter(Boolean).length, 'column') +
        (data.naRows ? ' · ' + data.naRows + ' n/a row(s) skipped' : '');
      $('tps-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    }).catch(err => {
      if (err && err.message === 'cancelled') return;
      alert('Failed to read source: ' + (err && err.message ? err.message : err));
    });
  }
  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      const parsed = parseTemplate(e.target.result, file.name);
      if (!parsed || !parsed.headers.length) { alert('Could not find a DATA ENTRY sheet in this template.'); return; }
      tplData = parsed;
      $('tps-tpl-name').textContent = file.name;
      const crops = parsed.lists.get('croptype');
      $('tps-tpl-meta').innerHTML = plural(parsed.headers.length, 'column') + ' · ' +
        plural(crops ? crops.length : 0, 'crop type') +
        (looksLikePackTemplate(parsed.headers) ? '' : ' <span class="cmp-warn">⚠ doesn’t look like a Pack Styles template</span>');
      $('tps-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    };
    r.readAsArrayBuffer(file);
  }

  // ─── Reset ───
  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    manualFills = {}; removedRows = new Set(); cellOverrides = {}; smartFixMap = {};
    pvMode = 'all'; pvColsOnly = false; autoMatched = []; weightSplitRows = 0; provMarks = {};
    IMOutName.forget();
    $('tps-src-name').textContent = 'No file selected';
    $('tps-tpl-name').textContent = 'No file selected';
    $('tps-src-meta').textContent = ''; $('tps-tpl-meta').textContent = '';
    ['tps-section-mapping', 'tps-section-matched', 'tps-section-smartfix', 'tps-section-required',
     'tps-section-dups', 'tps-section-preview', 'tps-summary', 'tps-build']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    $('tps-empty').style.display = '';
    if (wb) wb.reset();
    $('tps-run').disabled = true; $('tps-export').disabled = true;
    if (window.IMDebug) IMDebug.refresh('packstyles-standardize');
  }

  // ─── Debug dump ───
  function collectDebug() {
    if (!srcData || !tplData || !formattedRows) return null;
    const heads = H();
    const b = exportBlockers();
    const keptIdxs = formattedRows.map((_, i) => i).filter(i => !excluded(i));
    const cols = heads.map((h, ci) => {
      const si = mapping[ci] != null ? mapping[ci] : -1;
      return { index: ci, header: h, required: isReq(h), srcIndex: si, match: matchKind(ci), note: null,
        sample: si >= 0 && srcData.rows[0] ? srcData.rows[0][si] : null };
    });
    const asked = [
      IMDebug.ask('empty-required', 'required', 'Required columns with empty cells', {
        count: b.emptyCols.length, blocksExport: true,
        items: b.emptyCols.map(e => ({ column: e.header, emptyRows: e.empty }))
      }),
      IMDebug.ask('to-build', 'dropdown', 'Values to create in PickTrace (crop types the org lacks; columns with no options yet)', {
        count: b.build.length, blocksExport: true,
        items: b.build.map(c => ({ column: c.header, listIsEmpty: c.emptyList, values: c.values }))
      }),
      IMDebug.ask('off-list-values', 'dropdown', 'Values not present in a template dropdown', {
        count: b.off.length, blocksExport: true,
        detail: 'Respell via Smart Fixes, or create them in PickTrace and re-download the template.',
        items: b.off.map(c => ({ column: c.header, values: c.values, allowed: c.allowed }))
      }),
      IMDebug.ask('duplicate-names', 'collision', 'Rows sharing a Pack Style Name', {
        count: duplicates().size, blocksExport: false,
        items: [...duplicates().values()].map(g => ({ name: formattedRows[g[0]][nameCol()], rowIndexes: g }))
      })
    ];
    const fills = Object.keys(manualFills).map(i => ({ column: heads[+i], value: manualFills[i], kind: 'column fill', scope: 'blank cells' }))
      .concat(Object.keys(smartFixMap).map(k => {
        const sep = k.indexOf('||');
        return { column: heads[+k.slice(0, sep)], from: k.slice(sep + 2), value: smartFixMap[k], kind: 'smart fix' };
      }));
    const prov = IMDebug.deriveProvenance({
      headers: heads, rows: formattedRows, srcRows: srcData.rows, colToSrc: mapping,
      fills: manualFills, cellEdits: cellOverrides, smartFixes: smartFixMap, marks: provMarks
    });
    const out = IMDebug.output(heads, keptIdxs.map(i => formattedRows[i]), {
      note: 'exportedRowIndexes[n] is the formattedRows index behind exported row n — the same key used by provenance.codes and edits.',
      exportFileName: exportFileName()
    });
    out.exportedRowIndexes = keptIdxs;
    return {
      inputs: [
        IMDebug.file('source data (implementation workbook / export)', srcData, { naPlaceholderRowsSkipped: srcData.naRows || 0 }),
        IMDebug.file('personalized bulk Pack Styles template', tplData, {
          dropdowns: [...tplData.lists.entries()].map(([k, v]) => ({ column: k, values: v })),
          sheetNames: tplData.sheetNames || null
        })
      ],
      mapping: IMDebug.mapping(cols, srcData.headers, { srcRows: srcData.rows }),
      derived: [
        { kind: 'weight split', note: '"1 lb" in Weight → Weight 1 + Unit LB (a Unit already on the row wins).', rows: weightSplitRows },
        { kind: 'matched to a dropdown entry', note: 'Off-list values with exactly one plausible entry (Conventional → CONV, Clamshell → Clamshells).',
          items: autoMatched.map(m => ({ column: heads[m.ci], from: m.from, to: m.to, rows: m.rows })) }
      ],
      fills: fills,
      edits: {
        cells: Object.keys(cellOverrides).map(k => {
          const sep = k.indexOf('|');
          const ri = +k.slice(0, sep), ci = +k.slice(sep + 1);
          const si = mapping[ci];
          return { rowIndex: ri, column: heads[ci], value: cellOverrides[k],
            sourceValue: si != null && si >= 0 && srcData.rows[ri] ? srcData.rows[ri][si] : null };
        })
      },
      dropped: {
        removedByHand: { count: removedRows.size, rowIndexes: [...removedRows].slice(0, 500) },
        naPlaceholderRows: srcData.naRows || 0
      },
      asked: asked,
      output: out,
      provenance: prov
    };
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    if (!$('tps-src-file')) return; // markup not present yet
    initialized = true;
    mountWorkbench();
    $('tps-src-file').addEventListener('change', e => { if (e.target.files[0]) handleSrcFile(e.target.files[0]); e.target.value = ''; });
    $('tps-tpl-file').addEventListener('change', e => { if (e.target.files[0]) handleTplFile(e.target.files[0]); e.target.value = ''; });
    $('tps-run').addEventListener('click', runFormat);
    $('tps-export').addEventListener('click', async () => {
      const b = exportBlockers();
      if (hasBlockers(b)) { const ok = await showExportConfirm(b); if (!ok) return; }
      doExport();
    });
    $('tps-reset').addEventListener('click', reset);

    $('tps-mapping-table').addEventListener('change', e => {
      const sel = e.target.closest('.tps-map-select');
      if (!sel) return;
      mapping[+sel.dataset.ci] = +sel.value;
      refresh();
    });
    $('tps-build').addEventListener('click', e => {
      const btn = e.target.closest('.dt-build-toggle');
      if (!btn) return;
      const open = btn.closest('.dt-build-col').classList.toggle('dt-build-open');
      btn.textContent = open ? 'Show less' : 'Show ' + btn.dataset.more + ' more';
    });
    $('tps-smartfix-table').addEventListener('click', e => {
      const btn = e.target.closest('.tps-sf-apply');
      if (!btn) return;
      const to = btn.closest('tr').querySelector('.tps-sf-pick').value;
      if (!to) { alert('Pick a value to set "' + btn.dataset.from + '" to first.'); return; }
      applySmartFix(+btn.dataset.ci, btn.dataset.from, to);
    });
    $('tps-smartfix-apply-all').addEventListener('click', () => {
      let n = 0;
      $('tps-smartfix-table').querySelectorAll('.tps-sf-apply').forEach(btn => {
        const to = btn.closest('tr').querySelector('.tps-sf-pick').value;
        if (!to) return;
        smartFixMap[btn.dataset.ci + '||' + String(btn.dataset.from).toUpperCase()] = to;
        n++;
      });
      if (!n) { alert('No values chosen yet — pick a target for at least one row.'); return; }
      refresh();
    });
    $('tps-required-table').addEventListener('click', e => {
      const btn = e.target.closest('button');
      if (!btn) return;
      const ci = +btn.dataset.ci;
      if (btn.classList.contains('tps-req-undo')) { delete manualFills[ci]; refresh(); return; }
      if (!btn.classList.contains('tps-req-apply')) return;
      const tr = btn.closest('tr');
      const val = str(tr.querySelector('.tps-req-text').value) || str(tr.querySelector('.tps-req-pick').value);
      if (!val) { alert('Pick a value or type a manual override first.'); return; }
      manualFills[ci] = val;
      refresh();
    });
    const grid = $('tps-preview-table');
    grid.addEventListener('change', e => {
      const sel = e.target.closest('.ts-cell-select');
      if (sel) {
        if (sel.value === '__tps_custom__') {
          const td = sel.closest('td');
          const ri = +sel.dataset.ri, ci = +sel.dataset.ci;
          td.innerHTML = '<input class="ts-cell-input" type="text" data-ri="' + ri + '" data-ci="' + ci +
            '" placeholder="Type a value…" value="' + escHtml(formattedRows[ri][ci] || '') + '">';
          const inp = td.querySelector('input');
          inp.focus(); inp.select();
          return;
        }
        setCell(+sel.dataset.ri, +sel.dataset.ci, sel.value);
        return;
      }
      const inp = e.target.closest('.ts-cell-input');
      if (inp) setCell(+inp.dataset.ri, +inp.dataset.ci, inp.value);
    });
    grid.addEventListener('click', e => {
      const btn = e.target.closest('.tps-row-remove');
      if (!btn) return;
      removedRows.add(+btn.dataset.ri);
      refresh();
    });
    $('tps-preview-hint').addEventListener('change', e => {
      if (e.target.classList.contains('wb-rowfilter')) { pvMode = e.target.value; IMWorkbench.rowsTop(grid); renderPreview(); }
      else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
    });
    $('tps-preview-hint').addEventListener('click', e => {
      if (e.target.id !== 'tps-restore-rows') return;
      removedRows.clear();
      refresh();
    });
    if (window.IMDebug) {
      IMDebug.register('packstyles-standardize', {
        label: 'Pack Styles Standardize',
        ready: () => !!(srcData && tplData && formattedRows),
        collect: collectDebug
      });
      IMDebug.wire('tps-debug', 'packstyles-standardize');
    }
    // Click-to-copy in the result tables (not the editable grid or the panels with controls).
    $('tps-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      if (!td || td.closest('#tps-section-preview') || td.closest('#tps-section-mapping') ||
          td.closest('#tps-section-required') || e.target.closest('input, button, select')) return;
      const text = (td.textContent || '').trim();
      if (text && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => td.classList.toggle('cmp-copied'), () => {});
      }
    });
  }

  window.tpsInit = init;
  // Headless surface for testing the pipeline without the DOM.
  window.tpsInternals = {
    sheetAoa: sheetAoa, sliceSheet: sliceSheet, parseTemplate: parseTemplate, autoMap: autoMap,
    load: function (src, tpl) { srcData = src; tplData = tpl; mapping = autoMap(tpl.headers, src.headers); rebuildFormattedRows(); },
    rows: function () { return formattedRows; },
    findings: function () { return { autoMatched, weightSplitRows, off: offList(), build: toCreate(), dups: duplicates(), emptyReq: emptyRequired(), mapping }; }
  };
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
