// ═══════════════════════════════════════════════════════════════════════
// Shared Highlight Filter — reads the cell fill colors a client used to mark
// rows in a source workbook (e.g. a red First Name = "don't add this person")
// and lets the user leave those rows out, or keep ONLY those rows.
//
// A row carries a color when ANY of its cells has that fill, so a highlight on
// one cell (the first name) drops or keeps the whole row. White / no-fill and
// header rows are ignored. Only real cell fills are seen — conditional
// formatting colors are not stored on the cell and can't be read.
//
// Usage (per module):
//   const wb = XLSX.read(buf, { type: 'array', cellStyles: true });
//   scan  = IMHighlight.scan(wb, ws, aoaRowIdxs, headers, i => label);
//   state = IMHighlight.newState();                 // { mode:'off', keys:Set }
//   kept  = IMHighlight.filter(scan, state);        // data-row indexes to use
//   IMHighlight.render(sectionEl, scan, state, (next, rowsChanged) => {...},
//                      { noun: 'employee' });
// `aoaRowIdxs[i]` is data row i's index in sheet_to_json({ header: 1 }) output.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const AFFECTED_CAP = 300;   // rows listed under "Show the rows…"
  const SAMPLE_LABELS = 4;    // example names per color

  // ─── Color math ───
  function hexToRgb(hex) {
    const n = parseInt(hex, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function rgbToHex(r, g, b) {
    return [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0'))
      .join('').toUpperCase();
  }
  function rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h;
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    return [h * 60, s, l];
  }
  function hslToRgb(h, s, l) {
    if (!s) return [l * 255, l * 255, l * 255];
    const hue = t => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    h /= 360;
    return [hue(h + 1 / 3) * 255, hue(h) * 255, hue(h - 1 / 3) * 255];
  }
  // Excel's tint: negative darkens, positive lightens (applied to HSL lightness).
  function applyTint(hex, tint) {
    if (!tint) return hex;
    const [h, s, l] = rgbToHsl(...hexToRgb(hex));
    const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    return rgbToHex(...hslToRgb(h, s, l2));
  }
  function isWhite(hex) {
    return hexToRgb(hex).every(v => v >= 0xFA);
  }

  // Plain-words name so the panel reads "Red", "Light yellow", not just a hex.
  function colorName(hex) {
    const [h, s, l] = rgbToHsl(...hexToRgb(hex));
    if (s < 0.12) return l > 0.85 ? 'Light gray' : (l < 0.2 ? 'Black' : 'Gray');
    let base;
    if (h < 15 || h >= 335) base = 'Red';
    else if (h < 40) base = l < 0.35 ? 'Brown' : 'Orange';
    else if (h < 65) base = 'Yellow';
    else if (h < 160) base = 'Green';
    else if (h < 195) base = 'Cyan';
    else if (h < 255) base = 'Blue';
    else if (h < 290) base = 'Purple';
    else base = 'Pink';
    if (l >= 0.75) return 'Light ' + base.toLowerCase();
    if (l <= 0.3) return 'Dark ' + base.toLowerCase();
    return base;
  }

  // Cell style → fill hex (6 chars, upper) or '' when unfilled / white.
  // SheetJS only resolves theme colors with a non-zero index, so theme fills
  // (Excel's "Accent 2, lighter 60%") are resolved here from the workbook theme.
  function fillHex(style, wb) {
    const s = style && (style.fill || style);
    if (!s || !s.patternType || s.patternType === 'none') return '';
    const c = s.fgColor || s.bgColor;
    if (!c) return '';
    let hex = '';
    if (c.rgb) hex = String(c.rgb).slice(-6).toUpperCase();
    else if (c.theme != null) {
      const scheme = wb && wb.Themes && wb.Themes.themeElements && wb.Themes.themeElements.clrScheme;
      const base = scheme && scheme[c.theme] && scheme[c.theme].rgb;
      if (base) hex = applyTint(String(base).slice(-6).toUpperCase(), c.tint || 0);
    }
    if (!/^[0-9A-F]{6}$/.test(hex) || isWhite(hex)) return '';
    return hex;
  }

  function colLetter(n) {
    let s = '';
    let x = n + 1;
    while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); }
    return s;
  }

  // ─── Scan ───
  // Returns null when the sheet has no highlighted data rows (CSV, plain sheets).
  function scan(wb, ws, aoaRowIdxs, headers, labelFn) {
    if (!wb || !ws || !ws['!ref'] || !aoaRowIdxs || !aoaRowIdxs.length) return null;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const dense = Array.isArray(ws['!data']);
    const cellAt = (r, c) => dense ? (ws['!data'][r] || [])[c] : ws[XLSX.utils.encode_cell({ r, c })];
    const byHex = new Map();
    const rowColors = [];
    const sheetRows = [];
    aoaRowIdxs.forEach((ai, i) => {
      const r = range.s.r + ai;
      sheetRows.push(r + 1);
      const seen = new Set();
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = cellAt(r, c);
        if (!cell || !cell.s) continue;
        const hex = fillHex(cell.s, wb);
        if (!hex) continue;
        let e = byHex.get(hex);
        if (!e) { e = { hex, name: colorName(hex), rows: [], cols: new Map(), samples: [] }; byHex.set(hex, e); }
        const colName = (headers && headers[c] && String(headers[c]).trim()) || ('Column ' + colLetter(c));
        e.cols.set(colName, (e.cols.get(colName) || 0) + 1);
        if (!seen.has(hex)) {
          seen.add(hex);
          e.rows.push(i);
          if (e.samples.length < SAMPLE_LABELS) {
            const lb = labelFn ? String(labelFn(i) || '').trim() : '';
            if (lb) e.samples.push(lb);
          }
        }
      }
      rowColors.push([...seen]);
    });
    if (!byHex.size) return null;
    const colors = [...byHex.values()].sort((a, b) => b.rows.length - a.rows.length);
    return { colors, rowColors, sheetRows, total: aoaRowIdxs.length, labelFn: labelFn || null };
  }

  function newState() { return { mode: 'off', keys: new Set() }; }
  function cloneState(s) { return { mode: s.mode, keys: new Set(s.keys) }; }
  function isActive(st) { return !!(st && st.mode !== 'off' && st.keys.size); }

  // Data-row indexes to keep under the given state (all rows when inactive).
  function filter(sc, st) {
    if (!sc) return null;
    const all = sc.rowColors.map((_, i) => i);
    if (!isActive(st)) return all;
    return all.filter(i => {
      const hit = sc.rowColors[i].some(h => st.keys.has(h));
      return st.mode === 'only' ? hit : !hit;
    });
  }
  // Data-row indexes NOT used under the state.
  function dropped(sc, st) {
    if (!sc || !isActive(st)) return [];
    const keep = new Set(filter(sc, st));
    return sc.rowColors.map((_, i) => i).filter(i => !keep.has(i));
  }
  function keptSignature(sc, st) { return (filter(sc, st) || []).join(','); }

  function swatch(hex) {
    return '<span style="display:inline-block;width:14px;height:14px;border-radius:3px;vertical-align:-2px;' +
      'border:1px solid rgba(0,0,0,.25);background:#' + hex + ';"></span>';
  }
  function checkedNames(sc, st) {
    return sc.colors.filter(c => st.keys.has(c.hex)).map(c => c.name).join(', ');
  }

  // One-line status of what the filter is doing, for the panel and summaries.
  function statusText(sc, st, noun) {
    const n = sc.total;
    if (st.mode === 'off') return 'Off — all ' + plural(n, noun) + ' are used.';
    if (!st.keys.size) return 'Tick one or more colors below to ' + (st.mode === 'only' ? 'keep only' : 'leave out') + ' those rows.';
    const out = dropped(sc, st).length;
    if (st.mode === 'exclude') return 'Leaving out ' + out + ' of ' + plural(n, noun) + ' (rows with ' + checkedNames(sc, st) + ').';
    return 'Using only ' + (n - out) + ' of ' + plural(n, noun) + ' (rows with ' + checkedNames(sc, st) + '); ' + out + ' skipped.';
  }

  // ─── Panel ───
  // `sec` is an empty <section class="cmp-section">; its content is owned here.
  // onChange(nextState, rowsChanged) — rowsChanged is false when the change has
  // no effect on which rows are used (e.g. ticking a color while mode is Off),
  // so the caller can skip a rebuild. The caller re-renders after applying.
  function render(sec, sc, st, onChange, opts) {
    if (!sec) return;
    if (!sc) { sec.style.display = 'none'; sec.innerHTML = ''; return; }
    const o = opts || {};
    const noun = o.noun || 'row';
    const name = (sec.id || 'hl') + '-mode';
    sec.style.display = '';
    const modeOpt = (val, label) =>
      '<label style="display:flex;align-items:center;gap:6px;cursor:pointer;">' +
      '<input type="radio" name="' + name + '" value="' + val + '"' + (st.mode === val ? ' checked' : '') + '> ' + label + '</label>';
    let html = '<header class="cmp-section-head"><h3>Highlighted Rows (' + plural(sc.colors.length, 'color') + ' found)</h3></header>' +
      '<p class="cmp-sites-hint">Cell fill colors found in the source rows. A row has a color when <b>any</b> of its cells is filled with it ' +
      '(a highlighted first name counts for the whole row). Choose a mode, then tick the colors it applies to. ' +
      'Changing this rebuilds the preview from the source file.</p>' +
      '<div style="display:flex;gap:18px;align-items:center;flex-wrap:wrap;margin:6px 0 10px;font-size:13px;">' +
      modeOpt('off', 'Off — use every row') +
      modeOpt('exclude', 'Leave out rows with the ticked colors') +
      modeOpt('only', 'Only use rows with the ticked colors') +
      '</div>' +
      '<div class="table-wrap"><table class="data-table"><thead><tr><th style="width:28px;"></th><th>Color</th>' +
      '<th>' + noun.charAt(0).toUpperCase() + noun.slice(1) + 's</th><th>Found in</th><th>Examples</th></tr></thead><tbody>';
    sc.colors.forEach(c => {
      const cols = [...c.cols.entries()].sort((a, b) => b[1] - a[1])
        .map(([h, k]) => escHtml(h) + (c.cols.size > 1 ? ' <span class="text-muted small">(' + k + ')</span>' : '')).join(', ');
      html += '<tr>' +
        '<td style="text-align:center;"><input type="checkbox" class="hl-key" data-hex="' + c.hex + '"' + (st.keys.has(c.hex) ? ' checked' : '') + '></td>' +
        '<td style="white-space:nowrap;">' + swatch(c.hex) + ' <b>' + escHtml(c.name) + '</b> <span class="text-muted small">#' + c.hex + '</span></td>' +
        '<td>' + c.rows.length + '</td>' +
        '<td>' + cols + '</td>' +
        '<td>' + escHtml(c.samples.join(', ')) + (c.rows.length > c.samples.length ? ', &hellip;' : '') + '</td>' +
        '</tr>';
    });
    html += '</tbody></table></div>';
    const active = isActive(st);
    const color = active ? 'var(--amber-text,#92400e)' : 'var(--text-muted,#6b7280)';
    html += '<div style="margin-top:8px;font-size:13px;font-weight:600;color:' + color + ';">' + escHtml(statusText(sc, st, noun)) + '</div>';
    const out = dropped(sc, st);
    if (out.length) {
      const verb = st.mode === 'only' ? 'skipped (not highlighted with a ticked color)' : 'left out';
      html += '<details style="margin-top:6px;"><summary style="cursor:pointer;font-size:13px;">Show the ' + plural(out.length, 'row') + ' ' + verb + '</summary>' +
        '<div class="table-wrap" style="max-height:320px;overflow:auto;margin-top:6px;"><table class="data-table"><thead><tr>' +
        '<th>Sheet row</th><th>' + escHtml(o.labelHeader || 'Name') + '</th><th>Colors</th></tr></thead><tbody>';
      out.slice(0, AFFECTED_CAP).forEach(i => {
        html += '<tr><td>' + sc.sheetRows[i] + '</td><td>' + escHtml(sc.labelFn ? sc.labelFn(i) : '') + '</td>' +
          '<td>' + (sc.rowColors[i].map(swatch).join(' ') || '<span class="text-muted small">none</span>') + '</td></tr>';
      });
      html += '</tbody></table></div>' +
        (out.length > AFFECTED_CAP ? '<div class="text-muted small">First ' + AFFECTED_CAP + ' shown.</div>' : '') +
        '</details>';
    }
    sec.innerHTML = html;

    const fire = next => onChange(next, keptSignature(sc, next) !== keptSignature(sc, st));
    sec.querySelectorAll('input[name="' + name + '"]').forEach(r => r.addEventListener('change', e => {
      const next = cloneState(st);
      next.mode = e.target.value;
      fire(next);
    }));
    sec.querySelectorAll('.hl-key').forEach(cb => cb.addEventListener('change', e => {
      const next = cloneState(st);
      if (e.target.checked) next.keys.add(e.target.dataset.hex); else next.keys.delete(e.target.dataset.hex);
      fire(next);
    }));
  }

  // Plain object for the debug dump.
  function describe(sc, st) {
    if (!sc) return { colorsFound: 0, mode: 'off' };
    const out = dropped(sc, st);
    return {
      mode: st.mode,
      active: isActive(st),
      colorsFound: sc.colors.length,
      colors: sc.colors.map(c => ({ hex: c.hex, name: c.name, rows: c.rows.length, ticked: st.keys.has(c.hex),
        foundIn: [...c.cols.entries()].map(([column, cells]) => ({ column, cells })) })),
      rowsNotUsed: out.length,
      items: out.slice(0, 500).map(i => ({ sheetRow: sc.sheetRows[i], label: sc.labelFn ? sc.labelFn(i) : null, colors: sc.rowColors[i] }))
    };
  }

  window.IMHighlight = { scan, filter, dropped, isActive, newState, cloneState, render, describe, statusText, colorName };
})();
