// ═══════════════════════════════════════════════════════════════════════
// Shared Debug Dump — one JSON shape for every builder module.
//
// Each module registers a `collect()` that returns the sections below; this
// file wraps them in a common envelope, parks the result on window.imDebug and
// downloads it. The point is a complete audit of a single run: every file that
// went in, how each output column was matched to a source column, which values
// were DERIVED rather than copied, every bulk fill / smart fix / hand edit,
// every row dropped and why, and — the part that is otherwise invisible — the
// list of things the tool asked you to resolve and whether you resolved them.
//
// `provenance` closes the loop: per output cell, where that value came from.
// Modules that track derivation explicitly (Sites' address cross-reference)
// pass exact marks; for the rest deriveProvenance() reconstructs the origin by
// replaying mapping → fills → smart fixes → edits against the source row.
//
// Registration:
//   IMDebug.register('sites-standardize', {
//     label: 'Sites Standardize',
//     ready: () => !!formattedRows,
//     collect: () => ({ inputs, mapping, derived, fills, edits, dropped, asked, output, provenance })
//   });
//   IMDebug.wire('tss-debug', 'sites-standardize');
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const VERSION = 1;
  const SAMPLE_ROWS = 5;      // raw input rows kept per file
  const OUTPUT_ROW_CAP = 2000;// output rows serialized in full
  const PROV_ROW_CAP = 5000;  // rows given a compact provenance code list
  const TOP_VALUES = 8;       // distinct values listed per output column

  const registry = new Map();

  // ─── Provenance vocabulary ───
  // Codes are strings so they can carry a payload: "src:3", "xref:AZ".
  const PROV_LEGEND = {
    'src:<n>':     'Copied verbatim from source column <n> (see mapping.columns).',
    'parse:<what>':'Derived by parsing another cell on the same row (e.g. parse:address).',
    'xref:<what>': 'Resolved through a cross-reference lookup (e.g. xref:AZ = the reference address for state AZ).',
    'lookup:<what>':'Resolved through an ID → name lookup table.',
    'fill':        'Bulk column fill you applied from a panel.',
    'edit':        'Hand edit you made in the preview grid.',
    'rename':      'Name override you applied to resolve a collision.',
    'smartfix':    'Snapped to a template dropdown value via Smart Fixes.',
    'const':       'Constant written by the module (not from your data).',
    'derived':     'Computed by the module; exact origin not tracked.',
    'empty':       'Never filled. If the column is required this is a blocker.'
  };

  function todayStr() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') +
           '-' + String(d.getDate()).padStart(2, '0');
  }
  function str(v) { return v == null ? '' : String(v); }
  function trimmed(v) { return str(v).trim(); }

  // ─── Section normalizers ───────────────────────────────────────────────
  // Every module hands back loosely-shaped objects; these give the JSON a
  // stable schema so two dumps from different modules can be diffed.

  // One uploaded file. `data` is whatever the module keeps for it.
  function file(role, data, extra) {
    if (!data) return { role: role, loaded: false };
    const headers = Array.isArray(data.headers) ? data.headers.map(str) : null;
    const rows = Array.isArray(data.rows) ? data.rows : null;
    const out = {
      role: role,
      loaded: true,
      fileName: data.fileName || null,
      sheetName: data.sheetName || null,
      headerRow: data.headerRow != null ? data.headerRow : null,
      headers: headers,
      columnCount: headers ? headers.length : null,
      rowCount: rows ? rows.length : (data.count != null ? data.count : null)
    };
    if (rows && headers) {
      // Keep the first few rows as {header: value} so the dump is readable
      // without cross-referencing the header array by index. Cells past the
      // last header still matter (that is where Sites hides its reference
      // addresses), so they come back under a synthetic "col<N>" key.
      out.sampleRows = rows.slice(0, SAMPLE_ROWS).map(r => {
        const o = {};
        const width = Math.max(headers.length, (r || []).length);
        for (let i = 0; i < width; i++) {
          const key = headers[i] ? headers[i] : ('col' + i + ' (no header)');
          const v = trimmed(r ? r[i] : '');
          if (v) o[key] = v;
        }
        return o;
      });
      const widest = rows.reduce((m, r) => Math.max(m, (r || []).length), 0);
      if (widest > headers.length) {
        out.cellsPastHeaderRow = {
          note: 'Source rows are wider than the header row — these columns have no header but do carry data.',
          widestRow: widest,
          extraColumns: widest - headers.length
        };
      }
    }
    if (extra) Object.keys(extra).forEach(k => { out[k] = extra[k]; });
    return out;
  }

  // Column mapping. `cols` = [{ index, header, required, srcIndex, match, sample }]
  // `extra.srcRows` (optional) filters "source columns not used" down to the
  // ones that actually carry data — without it a wide sheet reports dozens of
  // empty trailing columns and the real orphans get lost.
  function mapping(cols, srcHeaders, extra) {
    const srcRows = extra && extra.srcRows ? extra.srcRows : null;
    const used = new Set();
    const columns = (cols || []).map(c => {
      const si = c.srcIndex != null ? c.srcIndex : -1;
      if (si >= 0) used.add(si);
      return {
        index: c.index,
        header: c.header,
        required: !!c.required,
        source: si >= 0
          ? { index: si, name: srcHeaders && srcHeaders[si] != null ? str(srcHeaders[si]) : ('col' + si) }
          : null,
        match: c.match || (si >= 0 ? 'mapped' : 'unmapped'),
        note: c.note || null,
        sample: c.sample != null ? trimmed(c.sample) : null
      };
    });
    // A source sheet can be far wider than its header row, so walk the data
    // too — a column with no header but real values (Sites' reference
    // addresses) is exactly the kind of orphan worth reporting.
    const width = srcRows
      ? srcRows.reduce((m, r) => Math.max(m, (r || []).length), (srcHeaders || []).length)
      : (srcHeaders || []).length;
    const unusedSource = [];
    for (let i = 0; i < width; i++) {
      if (used.has(i)) continue;
      const name = str((srcHeaders || [])[i]);
      const samples = [];
      if (srcRows) {
        for (let r = 0; r < srcRows.length && samples.length < 3; r++) {
          const v = trimmed((srcRows[r] || [])[i]);
          if (v && samples.indexOf(v) < 0) samples.push(v);
        }
        if (!name && !samples.length) continue;   // genuinely empty — not an orphan
      } else if (!name) {
        continue;
      }
      unusedSource.push({ index: i, name: name || '(no header)', hasData: samples.length > 0,
        sampleValues: srcRows ? samples : undefined });
    }
    const out = {
      columns: columns,
      mappedCount: columns.filter(c => c.source).length,
      unmappedCount: columns.filter(c => !c.source).length,
      sourceColumnsNotUsed: unusedSource
    };
    if (extra) Object.keys(extra).forEach(k => { if (k !== 'srcRows') out[k] = extra[k]; });
    return out;
  }

  // One thing the tool put in front of the user. `status` is the whole point:
  // it separates "the tool noticed and you fixed it" from "still outstanding".
  function ask(id, kind, title, opts) {
    const o = opts || {};
    const count = o.count != null ? o.count : (Array.isArray(o.items) ? o.items.length : 0);
    return {
      id: id,
      kind: kind,                       // required | dropdown | collision | smartfix | xref | data
      title: title,
      status: o.status || (count > 0 ? 'outstanding' : 'resolved'),
      count: count,
      blocksExport: !!o.blocksExport,
      detail: o.detail || null,
      items: Array.isArray(o.items) ? o.items.slice(0, 200) : undefined,
      itemsTruncated: Array.isArray(o.items) && o.items.length > 200 ? o.items.length - 200 : undefined
    };
  }

  // Final grid + per-column statistics.
  function output(headers, rows, opts) {
    const o = opts || {};
    const hs = (headers || []).map(str);
    const rs = rows || [];
    const columnStats = hs.map((h, i) => {
      let filled = 0;
      const counts = new Map();
      rs.forEach(r => {
        const v = trimmed(r ? r[i] : '');
        if (!v) return;
        filled++;
        counts.set(v, (counts.get(v) || 0) + 1);
      });
      const top = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, TOP_VALUES)
        .map(([v, n]) => ({ value: v, rows: n }));
      return {
        index: i,
        header: h,
        required: /\*$/.test(h),
        filled: filled,
        empty: rs.length - filled,
        distinctValues: counts.size,
        topValues: top
      };
    });
    return {
      rowCount: rs.length,
      columnCount: hs.length,
      headers: hs,
      columnStats: columnStats,
      rowsTruncated: rs.length > OUTPUT_ROW_CAP ? rs.length - OUTPUT_ROW_CAP : 0,
      rows: rs.slice(0, OUTPUT_ROW_CAP).map(r => {
        const o2 = {};
        hs.forEach((h, i) => { const v = trimmed(r ? r[i] : ''); if (v) o2[h] = v; });
        return o2;
      }),
      exportFileName: o.exportFileName || null,
      note: o.note || null
    };
  }

  // ─── Provenance ────────────────────────────────────────────────────────
  // Reconstructs each output cell's origin from the state a module already
  // keeps. `marks` lets a module override any cell it derived deliberately.
  //
  //   rows        [[...]]  final output rows
  //   srcRows     [[...]]  the source row behind each output row (parallel)
  //   colToSrc    colIdx → source column index (-1 = none)
  //   fills       colIdx → value (or { val })
  //   cellEdits   "ri|ci" → value
  //   nameEdits   ri → value      (paired with nameCol)
  //   smartFixes  "ci||UPPERVALUE" → value
  //   marks       "ri|ci" → code  (exact, wins over everything but edits)
  function deriveProvenance(opts) {
    const o = opts || {};
    const rows = o.rows || [];
    const srcRows = o.srcRows || null;
    const colToSrc = o.colToSrc || {};
    const fills = o.fills || {};
    const cellEdits = o.cellEdits || {};
    const nameEdits = o.nameEdits || {};
    const smartFixes = o.smartFixes || {};
    const marks = o.marks || {};
    const nameCol = o.nameCol != null ? o.nameCol : -1;
    const headers = o.headers || [];
    const width = headers.length || (rows[0] ? rows[0].length : 0);

    const fillVal = ci => {
      const f = fills[ci];
      if (f == null) return null;
      return typeof f === 'object' ? (f.val != null ? str(f.val) : null) : str(f);
    };
    // Reverse smart-fix table: a canonical value, per column, that Smart Fixes
    // could have produced. Only consulted when the cell no longer matches its
    // source, so a value that happens to equal the source stays "src:".
    const sfByCol = {};
    Object.keys(smartFixes).forEach(k => {
      const sep = k.indexOf('||');
      if (sep < 0) return;
      const ci = k.slice(0, sep);
      (sfByCol[ci] = sfByCol[ci] || new Set()).add(str(smartFixes[k]).toUpperCase().trim());
    });

    const codes = {};
    const tally = {};       // grouped by code family: "src:*", "xref:*", "fill" …
    const tallyExact = {};  // every distinct code, payload included
    const bump = c => {
      const key = c.indexOf(':') > 0 ? c.slice(0, c.indexOf(':') + 1) + '*' : c;
      tally[key] = (tally[key] || 0) + 1;
      tallyExact[c] = (tallyExact[c] || 0) + 1;
    };
    const limit = Math.min(rows.length, PROV_ROW_CAP);

    for (let ri = 0; ri < limit; ri++) {
      const row = rows[ri] || [];
      const src = srcRows ? srcRows[ri] : null;
      const line = [];
      for (let ci = 0; ci < width; ci++) {
        const val = trimmed(row[ci]);
        let code;
        if (cellEdits[ri + '|' + ci] != null) code = 'edit';
        else if (ci === nameCol && nameEdits[ri] != null) code = 'rename';
        else if (marks[ri + '|' + ci]) code = marks[ri + '|' + ci];
        else if (!val) code = 'empty';
        else {
          const si = colToSrc[ci] != null ? colToSrc[ci] : -1;
          const sv = (src && si >= 0) ? trimmed(src[si]) : null;
          if (sv != null && sv !== '' && sv === val) code = 'src:' + si;
          else if (sfByCol[ci] && sfByCol[ci].has(val.toUpperCase())) code = 'smartfix';
          else if (fillVal(ci) != null && fillVal(ci) === val) code = 'fill';
          else if (sv != null && sv !== '' && sv !== val) code = 'derived';
          else code = 'derived';
        }
        line.push(code);
        bump(code);
      }
      codes[ri] = line;
    }

    // Rows worth reading cell-by-cell: anything that is not a plain copy.
    const FLAG_CAP = 500;
    const flagged = [];
    let flaggedTotal = 0;
    Object.keys(codes).forEach(k => {
      const line = codes[k];
      if (!line.some(c => c === 'empty' || c === 'edit' || c === 'rename' ||
                          c === 'smartfix' || c === 'fill' || c.indexOf('xref:') === 0 ||
                          c.indexOf('parse:') === 0 || c.indexOf('lookup:') === 0 || c === 'derived')) return;
      flaggedTotal++;
      if (flagged.length >= FLAG_CAP) return;
      const ri = +k;
      const cells = {};
      headers.forEach((h, ci) => { cells[h] = { value: trimmed((rows[ri] || [])[ci]), from: line[ci] }; });
      flagged.push({ rowIndex: ri, cells: cells });
    });

    return {
      legend: PROV_LEGEND,
      note: 'codes[<output row index>] lists one origin code per output column, in header order. ' +
            'tally groups codes by family; tallyExact keeps each code with its payload.',
      rowsCovered: limit,
      rowsTruncated: rows.length > limit ? rows.length - limit : 0,
      tally: tally,
      tallyExact: tallyExact,
      codes: codes,
      flaggedRowCount: flaggedTotal,
      flaggedRowsShown: flagged.length,
      flaggedRowsTruncated: flaggedTotal > flagged.length ? flaggedTotal - flagged.length : 0,
      flaggedRows: flagged
    };
  }

  // ─── Envelope + download ───────────────────────────────────────────────
  function build(id) {
    const entry = registry.get(id);
    if (!entry) return null;
    let payload;
    try {
      payload = entry.collect();
    } catch (err) {
      return {
        module: id,
        label: entry.label || id,
        generatedAt: new Date().toISOString(),
        collectorError: { message: err && err.message ? err.message : String(err), stack: err && err.stack ? err.stack : null }
      };
    }
    if (!payload) return null;
    const asked = payload.asked || [];
    return Object.assign({
      tool: 'IMtool debug dump',
      version: VERSION,
      module: id,
      label: entry.label || id,
      generatedAt: new Date().toISOString(),
      summary: {
        inputsLoaded: (payload.inputs || []).filter(i => i && i.loaded).length,
        outputRows: payload.output ? payload.output.rowCount : null,
        outstanding: asked.filter(a => a.status === 'outstanding').length,
        resolved: asked.filter(a => a.status === 'resolved').length,
        blockingExport: asked.filter(a => a.status === 'outstanding' && a.blocksExport).length
      }
    }, payload);
  }

  function dump(id) {
    const entry = registry.get(id);
    if (!entry) { console.warn('[IMDebug] no module registered as "' + id + '"'); return; }
    if (entry.ready && !entry.ready()) {
      alert('Nothing to dump yet — load your files and format the sheet first.');
      return;
    }
    const data = build(id);
    if (!data) { alert('Nothing to dump yet — load your files and format the sheet first.'); return; }
    window.imDebug = data;
    console.log('[IMDebug] ' + (entry.label || id) + ' dump on window.imDebug', data);
    let text;
    try {
      text = JSON.stringify(data, null, 2);
    } catch (err) {
      alert('Could not serialize the debug dump: ' + (err && err.message ? err.message : err) +
            '\n\nIt is still on window.imDebug for console inspection.');
      return;
    }
    const blob = new Blob([text], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = id + '-debug-' + todayStr() + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  function register(id, spec) {
    if (!id || !spec || typeof spec.collect !== 'function') return;
    registry.set(id, spec);
    refresh(id);
  }

  // Keep a wired button's enabled state in step with the module's readiness.
  const wired = new Map(); // id → [buttonId…]
  function wire(btnId, id) {
    const btn = document.getElementById(btnId);
    if (!btn || btn._imDebugWired) return;
    btn._imDebugWired = true;
    btn.addEventListener('click', () => dump(id));
    const list = wired.get(id) || [];
    if (!list.includes(btnId)) list.push(btnId);
    wired.set(id, list);
    refresh(id);
  }
  function refresh(id) {
    const entry = registry.get(id);
    const list = wired.get(id) || [];
    list.forEach(btnId => {
      const btn = document.getElementById(btnId);
      if (!btn) return;
      const ok = !!(entry && (!entry.ready || entry.ready()));
      btn.disabled = !ok;
      btn.title = ok
        ? 'Download a JSON audit of this run: every input, how each column was mapped, what was derived, what you filled or edited, what was dropped, and what is still outstanding.'
        : 'Load your files and format the sheet first.';
    });
  }

  window.IMDebug = {
    register: register,
    wire: wire,
    dump: dump,
    build: build,
    refresh: refresh,
    file: file,
    mapping: mapping,
    ask: ask,
    output: output,
    deriveProvenance: deriveProvenance,
    PROV_LEGEND: PROV_LEGEND
  };
})();
