// ═══════════════════════════════════════════════════════════════════════
// Column Directives — a client-annotated source sheet. Some clients return
// the Implementation Data Template with its columns repurposed and add a row
// under the headers saying what each one really holds, e.g.
//   "Job Notes-Site"                  → this column feeds Site
//   "Job Number-Location ALT ID"      → this column feeds Alt ID
//   "Old Code-IGNORE THIS COLUMN"     → skip it
//   "Rootstock Notes"                 → no target; just describes the column
// and mark the columns to skip with a red fill.
//
// Nothing here runs unless that row is found, so an ordinary source file is
// read and mapped exactly as before. When it IS found:
//   - the row is left out of the data,
//   - its targets are mapped first, red / IGNORE columns are never mapped,
//   - the module's own auto-map fills the rest from the remaining columns,
//   - a column with only a description is matched by that description,
//   - cells holding only a dash ("-") read as blank.
// The mapping panel says what was applied and can switch it all off again.
//
// Usage (per module):
//   IMDirectives.attach(sheet, { styledWb, sheetName, ws, headerAt, rowAt })  // at read time
//   if (IMDirectives.on(srcData)) mapping = IMDirectives.buildMapping({...}).mapping;
//   IMDirectives.renderBanner(sectionEl, srcData, bulkHeaders, onToggle);
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\*+$/, '').trim();
  const loose = s => norm(s).replace(/[^a-z0-9]/g, '');

  // Field names a directive may point at — both bulk templates plus the
  // Implementation Data Template's own labels. Only used to RECOGNIZE the row,
  // so a data row whose values happen to contain a dash ("Example Ranch -
  // Block 1", "PEARS-SAMPLE") never qualifies.
  const KNOWN = new Set([
    'Site', 'Sites', 'Site Group', 'Site Groups', 'Name', 'Location Name', 'Site Name',
    'Alt ID', 'Location Alt ID', 'Site Alt ID', 'Location Type', 'Site Type',
    'Crop & Variety', 'Crop', 'Crop Type', 'Variety', 'Clone/Subvariety', 'Subvariety', 'Sub Variety',
    'Planted At', 'Acreage', 'Length', 'Plant Count', 'Start Date', 'Location Group', 'Lot Number',
    'Rootstock', 'Plant/Seed Seller', 'Plant/Seed Producer', 'Growing Manager', 'Production Status',
    'Training Style', 'Trellis Type', 'Mulch Type', 'Row Direction', 'Organic Status',
    'Organic Certifier', 'Wet Date', 'Germination Date', 'Planted Date', 'Grafting Date',
    'Production Start', 'Organic Certification Date', 'Stand Count', 'Row/Bed Count', 'Post Count',
    'Percent Covered', 'Row Spacing', 'Plant Spacing', 'Post Spacing', 'Bed Width',
    'Custom Data 1', 'Custom Data 2', 'Custom Data 3', 'Loc Data 1', 'Loc Data 2', 'Loc Data 3',
    'Employer', 'Address', 'Address1', 'Address2', 'City', 'State', 'Zip', 'Country'
  ].map(loose));

  const IGNORE_RE = /\bignore\b/i;
  const DASH_ONLY = /^[-–—]+$/;

  // "Job Notes-Site" → { desc:'Job Notes', target:'Site' }.
  // Splits at the LAST dash, so a description may carry dashes of its own.
  // "Legacy block code-" has an empty target; "Rootstock Notes" has none.
  function splitCell(v) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return null;
    if (IGNORE_RE.test(s)) return { desc: s.replace(/-?\s*ignore\b.*$/i, '').trim(), target: '', ignore: true };
    const i = s.lastIndexOf('-');
    if (i < 0) return { desc: s, target: '', ignore: false };
    return { desc: s.slice(0, i).trim(), target: s.slice(i + 1).trim(), ignore: false };
  }

  // A directive row names at least two known fields as "description-Field",
  // or marks a column IGNORE alongside at least one.
  function isDirectiveRow(row) {
    let known = 0, ignores = 0;
    (row || []).forEach(v => {
      const d = splitCell(v);
      if (!d) return;
      if (d.ignore) { ignores++; return; }
      if (d.desc && d.target && KNOWN.has(loose(d.target))) known++;
    });
    return known >= 2 || (ignores >= 1 && known >= 1);
  }

  // Index (into `rows`) of the directive row among the first few data rows, or -1.
  function find(rows) {
    const lim = Math.min((rows || []).length, 3);
    for (let i = 0; i < lim; i++) if (isDirectiveRow(rows[i])) return i;
    return -1;
  }

  // Columns (0-based from the sheet's first used column) whose cell in any of
  // the given rows has a red fill. `wb` is a cellStyles read of the file;
  // `origin` is the used range's top-left in the PLAIN read the rows came from
  // — a cellStyles read stretches the range over styled empty cells, so its
  // own origin can't be trusted to line up.
  function redCols(wb, sheetName, origin, aoaIdxs, width) {
    const out = new Set();
    const ws = wb && wb.Sheets[sheetName];
    if (!ws || !window.IMHighlight || !IMHighlight.fillHex) return out;
    const dense = Array.isArray(ws['!data']);
    const cellAt = (r, c) => dense ? (ws['!data'][r] || [])[c] : ws[XLSX.utils.encode_cell({ r, c })];
    aoaIdxs.forEach(ai => {
      if (ai == null || ai < 0) return;
      for (let c = 0; c < width; c++) {
        const cell = cellAt(origin.r + ai, origin.c + c);
        if (!cell || !cell.s) continue;
        const hex = IMHighlight.fillHex(cell.s, wb);
        if (hex && /\bred\b/i.test(IMHighlight.colorName(hex))) out.add(c);
      }
    });
    return out;
  }

  // Called by a module's file reader on each sheet it slices. When the sheet
  // carries a directive row, records it and switches `sheet.rows` (and
  // `sheet.rowNums`, if the module keeps them) to the cleaned rows. Both the
  // cleaned and the original versions are kept so the panel toggle can swap.
  //   o.styledWb  — () → the file read with cellStyles, or null. Called only
  //                 when a directive row is found: values always come from the
  //                 module's plain read, which a cellStyles read would change.
  //   o.sheetName, o.ws — the sheet, and its worksheet in the plain read
  //   o.headerAt  — the header row's index in sheet_to_json({ header: 1 }) output
  //   o.rowAt     — the same index for each of sheet.rows
  function attach(sheet, o) {
    if (!sheet || !sheet.rows) return sheet;
    const at = find(sheet.rows);
    if (at < 0) return sheet;
    const dirRow = sheet.rows[at];
    const width = Math.max(sheet.headers.length, dirRow.length);
    let red = new Set();
    try {
      const swb = o.styledWb ? o.styledWb() : null;
      const origin = o.ws && o.ws['!ref'] ? XLSX.utils.decode_range(o.ws['!ref']).s : { r: 0, c: 0 };
      red = redCols(swb, o.sheetName, origin, [o.headerAt, o.rowAt ? o.rowAt[at] : null], width);
    } catch (err) { /* no fills readable — IGNORE text still applies */ }
    const cols = [];
    for (let c = 0; c < width; c++) {
      const d = splitCell(dirRow[c]) || { desc: '', target: '', ignore: false };
      if (red.has(c)) d.ignore = true;
      d.red = red.has(c);
      cols.push(d);
    }
    const ignored = new Set();
    cols.forEach((d, c) => { if (d.ignore) ignored.add(c); });
    const keep = (_, i) => i !== at;
    const rowsOn = sheet.rows.filter(keep).map(r => r.map(v => DASH_ONLY.test(String(v == null ? '' : v).trim()) ? '' : v));
    sheet.directives = {
      enabled: true,
      cols, ignored,
      sheetRow: o.rowAt && o.rowAt[at] != null
        ? o.rowAt[at] + 1 + (o.ws && o.ws['!ref'] ? XLSX.utils.decode_range(o.ws['!ref']).s.r : 0) : null,
      rowsOn, rowsOff: sheet.rows,
      rowNumsOn: sheet.rowNums ? sheet.rowNums.filter(keep) : null,
      rowNumsOff: sheet.rowNums || null,
      notes: null
    };
    sheet.rows = rowsOn;
    if (sheet.rowNums) sheet.rowNums = sheet.directives.rowNumsOn;
    return sheet;
  }

  function on(src) { return !!(src && src.directives && src.directives.enabled); }

  // Swap a loaded source between the cleaned rows and the file as-is.
  function setEnabled(src, enabled) {
    const d = src && src.directives;
    if (!d) return;
    d.enabled = !!enabled;
    src.rows = d.enabled ? d.rowsOn : d.rowsOff;
    if (d.rowNumsOff) src.rowNums = d.enabled ? d.rowNumsOn : d.rowNumsOff;
  }

  // Is source column c ignored by the directives currently in effect?
  function ignoredCol(src, c) { return on(src) && src.directives.ignored.has(c); }

  // Best template column for a field name, or null. Exact header beats a loose
  // match beats the module's aliases (in their listed order) beats `extras`;
  // a name only matched after dropping `prefix` ("Location ALT ID" → "ALT ID")
  // ranks below all of those. No substring guessing — a directive names a field.
  function resolveName(text, bulkHeaders, aliases, opts) {
    const o = opts || {};
    const s = norm(text);
    if (!s) return null;
    const variants = [{ s, add: 0 }];
    if (o.prefix) {
      const re = new RegExp('^' + o.prefix + '\\s+', 'i');
      const t = s.replace(re, '').trim();
      if (t && t !== s) variants.push({ s: t, add: 100 });
    }
    let best = null;
    bulkHeaders.forEach((bh, bi) => {
      const full = String(bh || '').trim().toLowerCase();
      const plain = norm(bh);
      const list = (aliases[full] || aliases[plain] || []).concat((o.extras || {})[plain] || []);
      variants.forEach(v => {
        const vl = loose(v.s);
        let rank = -1;
        if (v.s === plain) rank = 0;
        else if (vl && vl === loose(plain)) rank = 1;
        else {
          const k = list.findIndex(a => norm(a) === v.s || loose(a) === vl);
          if (k >= 0) rank = 2 + k;
        }
        if (rank >= 0 && (!best || rank + v.add < best.rank)) best = { bi, rank: rank + v.add };
      });
    });
    return best;
  }

  // The mapping (bulkColIdx → srcColIdx, -1 = unmapped) for an annotated source.
  //   autoMap  — the module's own auto-mapper, run on whatever is left
  //   aliases  — the module's alias table (keyed like its autoMap's)
  //   extras   — { plainBulkHeader: [names] } matched only from directives
  //   prefix   — word a target may lead with ('location' / 'site')
  //   hasData  — c → true when source column c holds any value; empty
  //              columns never win a mapping in an annotated file
  function buildMapping(o) {
    const d = o.src.directives;
    const bulk = o.bulkHeaders;
    const out = {};
    bulk.forEach((_, bi) => { out[bi] = -1; });
    const claimed = new Set();
    d.cols.forEach((x, c) => { if (!x.ignore && x.target) claimed.add(c); });
    const notes = { applied: [], byDesc: [], unmatched: [], lost: [] };
    const ropts = { extras: o.extras, prefix: o.prefix };

    // 1. Targets. Two columns naming the same field: the closer name wins.
    const best = {};
    d.cols.forEach((x, c) => {
      if (x.ignore || !x.target) return;
      const r = resolveName(x.target, bulk, o.aliases, ropts);
      if (!r) { notes.unmatched.push(c); return; }
      const prev = best[r.bi];
      if (!prev || r.rank < prev.rank) {
        if (prev) notes.lost.push({ c: prev.c, bi: r.bi });
        best[r.bi] = { c, rank: r.rank };
      } else notes.lost.push({ c, bi: r.bi });
    });
    Object.keys(best).forEach(bi => { out[bi] = best[bi].c; notes.applied.push({ c: best[bi].c, bi: +bi }); });

    // 2. The module's usual auto-map over the columns nothing has claimed.
    const masked = o.srcHeaders.map((h, c) =>
      (d.ignored.has(c) || claimed.has(c) || (o.hasData && !o.hasData(c))) ? '' : h);
    const auto = o.autoMap(bulk, masked);
    bulk.forEach((_, bi) => { if (out[bi] < 0 && auto[bi] != null && auto[bi] >= 0) out[bi] = auto[bi]; });

    // 3. Descriptions, for columns still unused, into fields still empty.
    const used = new Set(Object.keys(out).map(k => out[k]).filter(c => c >= 0));
    d.cols.forEach((x, c) => {
      if (x.ignore || x.target || !x.desc || used.has(c) || claimed.has(c)) return;
      if (o.hasData && !o.hasData(c)) return;
      const r = resolveName(x.desc, bulk, o.aliases, ropts);
      if (r && out[r.bi] < 0) { out[r.bi] = c; used.add(c); notes.byDesc.push({ c, bi: r.bi }); }
    });

    d.notes = notes;
    return { mapping: out, notes };
  }

  // The note at the top of a module's Column Mapping panel. Shown only for an
  // annotated source; the checkbox switches the directives off (or back on)
  // and hands control to `onToggle`, which re-runs the module's format.
  function renderBanner(sec, src, bulkHeaders, onToggle) {
    if (!sec) return;
    let box = sec.querySelector('.im-dir-banner');
    const d = src && src.directives;
    if (!d) { if (box) box.remove(); return; }
    if (!box) {
      box = document.createElement('div');
      box.className = 'cmp-slot-hint im-dir-banner';
      box.style.marginBottom = '10px';
      const wrap = sec.querySelector('.table-wrap');
      sec.insertBefore(box, wrap || null);
    }
    const hdr = c => {
      const h = src.headers[c];
      return '<b>' + escHtml(h || ('Column ' + (c + 1))) + '</b>';
    };
    const where = d.sheetRow ? 'row ' + d.sheetRow : 'the row under the headers';
    let html = '<label style="display:flex;gap:6px;align-items:center;font-weight:600;margin-bottom:4px;">' +
      '<input type="checkbox" class="im-dir-toggle"' + (d.enabled ? ' checked' : '') + '> ' +
      'Use the column instructions in ' + escHtml(where) + '</label>';
    if (!d.enabled) {
      html += '<div>Off &mdash; the file is read as-is: that row is treated as data and columns map by header only.</div>';
    } else {
      const n = d.notes || { applied: [], byDesc: [], unmatched: [], lost: [] };
      const pairs = arr => arr.map(x => hdr(x.c) + ' &rarr; ' + escHtml(bulkHeaders[x.bi])).join(', ');
      const lines = [];
      if (n.applied.length) lines.push('Mapped as instructed: ' + pairs(n.applied) + '.');
      if (n.byDesc.length) lines.push('Matched by description: ' + n.byDesc.map(x =>
        hdr(x.c) + ' (&ldquo;' + escHtml(d.cols[x.c].desc) + '&rdquo;) &rarr; ' + escHtml(bulkHeaders[x.bi])).join(', ') + '.');
      const ign = [...d.ignored].sort((a, b) => a - b);
      if (ign.length) lines.push('Ignored (red or IGNORE): ' + ign.map(hdr).join(', ') + '.');
      if (n.unmatched.length) lines.push('No column for it in this template: ' + n.unmatched.map(c =>
        hdr(c) + ' (&ldquo;' + escHtml(d.cols[c].target) + '&rdquo;)').join(', ') + '.');
      if (n.lost.length) lines.push('Also pointed at a field already taken: ' + pairs(n.lost) + ' &mdash; not mapped.');
      lines.push('That row is left out of the data, and cells holding only &ldquo;-&rdquo; read as blank.');
      html += lines.map(l => '<div>' + l + '</div>').join('');
    }
    box.innerHTML = html;
    box.querySelector('.im-dir-toggle').addEventListener('change', e => {
      setEnabled(src, e.target.checked);
      if (onToggle) onToggle();
    });
  }

  window.IMDirectives = { attach, on, setEnabled, ignoredCol, buildMapping, renderBanner, resolveName, find };
})();
