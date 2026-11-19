// ═══════════════════════════════════════════════════════════════════════
// Workbench layout for the Template Standardize tools — the same shape as
// Bulk Template Standardize (data-template.js): the upload slots fold into a
// one-line strip once the sheet is built, the preview is pinned to the top of
// the scroll area, and every other section becomes a button (its name, what
// it's for, and how many items wait in it) that opens its panel, one at a
// time. The look is shared too: the same dt-nav / dt-panel / dt-files-strip
// classes in styles.css.
//
// A module keeps rendering its sections exactly as before. mount() moves them
// into panels once; refresh() recounts and re-shows, and is called after every
// render (each module calls it from its summary update).
//
//   const wb = IMWorkbench.mount({
//     pane: 'ts-loc-wrap',             // holds the tool's .cmp-uploads
//     results: 'ts-results',           // holds the tool's sections
//     preview: 'ts-section-preview',   // pinned above the buttons
//     ready: () => !!formattedRows,    // false → upload slots open, no buttons yet
//     sections: [{ key, group: 'fix' | 'check', name, desc, unit,
//                  ids: ['ts-section-…', …], count: () => n }]
//   });
//   wb.refresh();          // after a render
//   wb.open('required');   // e.g. from a flagged preview cell
//   wb.reset();            // with the tool's Reset
//   onOpen: key => …       // optional — you clicked a section's button
// ═══════════════════════════════════════════════════════════════════════
window.IMWorkbench = (function () {
  'use strict';

  const GROUP_LABEL = { fix: 'Fix before upload', check: 'Check' };
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const plural = (n, w) => n + ' ' + (n === 1 ? w : w.replace(/y$/, 'ie') + 's');

  function mount(o) {
    const $ = id => document.getElementById(id);
    const results = $(o.results), pane = $(o.pane), pv = $(o.preview);
    if (!results || !pane || !pv) return null;
    if (results._wb) return results._wb;
    const uploads = pane.querySelector('.cmp-uploads');

    // Files strip, just above the upload slots it stands in for.
    const strip = document.createElement('div');
    strip.className = 'dt-files-strip';
    strip.style.display = 'none';
    if (uploads) uploads.parentNode.insertBefore(strip, uploads);

    // Preview first (pinned), then the buttons, then the open panel.
    pv.classList.add('wb-pv');
    results.insertBefore(pv, results.firstChild);
    const head = pv.querySelector('.cmp-section-head');
    if (head) {
      const tools = document.createElement('div');
      tools.className = 'wb-pv-tools';
      tools.innerHTML = '<button type="button" class="btn btn-ghost btn-sm wb-pv-toggle" aria-expanded="true">Hide grid</button>';
      head.appendChild(tools);
      const legend = document.createElement('div');
      legend.className = 'dt-pv-legend wb-pv-legend';
      legend.innerHTML = '<span class="dt-lg dt-lg-req">Required, empty or a clashing name</span>' +
        '<span class="dt-lg dt-lg-off">Not in the dropdown, or worth a check</span>' +
        '<span class="dt-lg dt-lg-chg">Changed &mdash; hover a cell for was &rarr; now</span>';
      head.after(legend);
    }
    const nav = document.createElement('nav');
    nav.className = 'dt-nav';
    nav.setAttribute('aria-label', 'Sections');
    nav.style.display = 'none';
    const panels = document.createElement('div');
    panels.className = 'dt-panels';
    panels.style.display = 'none';
    pv.after(nav);
    nav.after(panels);
    o.sections.forEach(s => {
      const p = document.createElement('div');
      p.className = 'dt-panel';
      p.dataset.panel = s.key;
      (s.ids || []).forEach(id => { const el = $(id); if (el) p.appendChild(el); });
      panels.appendChild(p);
    });
    const empty = document.createElement('div');
    empty.className = 'dt-panel-empty';
    empty.style.display = 'none';
    panels.appendChild(empty);
    const sp = split({ preview: pv, side: [nav, panels], grid: pv.querySelector('.table-wrap'), tools: pv.querySelector('.wb-pv-tools') });

    const st = { active: null, picked: false, filesOpen: null, pvHidden: false };

    nav.addEventListener('click', e => {
      const b = e.target.closest('.dt-nav-btn');
      if (!b) return;
      st.active = b.dataset.sec; st.picked = true;
      refresh();
      if (o.onOpen) o.onOpen(st.active);
    });
    strip.addEventListener('click', e => {
      if (!e.target.closest('.wb-files-toggle')) return;
      st.filesOpen = !!uploads && uploads.style.display === 'none';
      refresh();
    });
    pv.addEventListener('click', e => {
      if (!e.target.closest('.wb-pv-toggle')) return;
      st.pvHidden = !st.pvHidden;
      refresh();
    });

    // One chip per upload slot: its label, and the file in it (or not).
    function renderStrip(ready) {
      if (!uploads) return;
      const open = st.filesOpen === null ? !ready : st.filesOpen;
      uploads.style.display = open ? '' : 'none';
      const chips = [];
      let any = false;
      uploads.querySelectorAll('.cmp-slot').forEach(sl => {
        const lab = sl.querySelector('.cmp-slot-label'), fn = sl.querySelector('.cmp-filename');
        if (!lab || !fn) return;
        const name = [...lab.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('')
          .replace(/^\s*\d+\.\s+/, '').trim();   // "1. Source" → "Source"; "3.0 Sites" stays
        const file = (fn.textContent || '').trim();
        const loaded = !!file && !/^no (file|export|template)/i.test(file);
        if (loaded) any = true;
        // An optional slot left empty isn't worth a chip.
        if (!loaded && sl.classList.contains('cmp-slot-opt')) return;
        chips.push('<span class="dt-file-chip ' + (loaded ? 'is-ready' : 'is-empty') + '" title="' + esc(file) + '"><b>' +
          esc(name) + '</b> ' + esc(loaded ? file : 'not loaded') + '</span>');
      });
      if (!any) { strip.style.display = 'none'; strip.innerHTML = ''; return; }
      strip.innerHTML = chips.join('') + '<button type="button" class="btn btn-ghost btn-sm wb-files-toggle" aria-expanded="' + open + '">' +
        (open ? 'Fold file slots' : 'Change files') + '</button>';
      strip.style.display = '';
    }

    function refresh() {
      const ready = !!o.ready();
      renderStrip(ready);
      pv.classList.toggle('is-collapsed', st.pvHidden);
      { const t = pv.querySelector('.wb-pv-toggle');
        if (t) { t.textContent = st.pvHidden ? 'Show grid' : 'Hide grid'; t.setAttribute('aria-expanded', String(!st.pvHidden)); } }
      sp.show(ready);
      if (!ready) { nav.style.display = 'none'; panels.style.display = 'none'; return; }

      const counts = {};
      o.sections.forEach(s => { let n = 0; try { n = +s.count() || 0; } catch (err) { n = 0; } counts[s.key] = n; });
      // Until you pick one, the open section is the first with something waiting.
      if (!st.picked || !st.active || !o.sections.some(s => s.key === st.active)) {
        const first = o.sections.find(s => s.group === 'fix' && counts[s.key]) || o.sections.find(s => counts[s.key]);
        st.active = (first || o.sections[0]).key;
      }
      nav.innerHTML = ['fix', 'check'].map(g => {
        const list = o.sections.filter(s => s.group === g);
        if (!list.length) return '';
        return '<div class="dt-nav-group"><div class="dt-nav-label">' + GROUP_LABEL[g] + '</div><div class="dt-nav-row">' +
          list.map(s => {
            const n = counts[s.key], on = s.key === st.active;
            return '<button type="button" class="dt-nav-btn is-' + g + (on ? ' is-on' : '') + (n ? '' : ' is-clear') +
              '" data-sec="' + s.key + '" aria-pressed="' + on + '">' +
              '<span class="dt-nav-top"><span class="dt-nav-name">' + esc(s.name) + '</span>' +
              '<span class="dt-nav-count' + (n ? '' : ' is-zero') + '">' + (n ? n : '✓ 0') + '</span></span>' +
              '<span class="dt-nav-desc">' + esc(s.desc) + '</span>' +
              (n && s.unit ? '<span class="dt-nav-where">' + esc(plural(n, s.unit)) + '</span>' : '') + '</button>';
          }).join('') + '</div></div>';
      }).join('');
      nav.style.display = '';
      panels.style.display = '';
      let filled = false;
      panels.querySelectorAll('.dt-panel').forEach(p => {
        const on = p.dataset.panel === st.active;
        p.style.display = on ? '' : 'none';
        if (on) filled = [...p.children].some(ch => ch.style.display !== 'none' && (ch.tagName === 'SECTION' || ch.innerHTML.trim()));
      });
      const sec = o.sections.find(s => s.key === st.active);
      empty.style.display = filled ? 'none' : '';
      empty.textContent = filled ? '' : '✓ Nothing in ' + sec.name +
        (sec.group === 'fix' ? ' — nothing here will stop the upload.' : ' — nothing to review.');
    }

    // Open a section and bring its buttons into view beside / under the preview.
    function open(key) {
      st.active = key; st.picked = true;
      refresh();
      if (o.onOpen) o.onOpen(key);
      sp.reveal(nav);
    }

    function reset() { st.active = null; st.picked = false; st.filesOpen = null; st.pvHidden = false; refresh(); }

    const api = { refresh: refresh, open: open, reset: reset };
    results._wb = api;
    return api;
  }

  // ─── Side by side / stacked ───
  // Side by side: the preview on the left, the buttons and the open section on
  // the right, a drag handle between them setting the split. Stacked: the
  // preview over the buttons, the handle setting the grid's height. One choice
  // for every tool (bulk and single), held in sessionPrefs like the app's other
  // settings — nothing is written to the browser. Side by side is the default
  // on a wide window; a narrow one always stacks (styles.css).
  const PREF_LAYOUT = 'imtool_wb_layout', PREF_LEFT = 'imtool_wb_left';
  const pref = k => { try { return window.sessionPrefs ? sessionPrefs.getItem(k) : null; } catch (e) { return null; } };
  const keep = (k, v) => { try { if (window.sessionPrefs) sessionPrefs.setItem(k, v); } catch (e) { /* session-only nicety */ } };
  const LEFT_DEFAULT = 58;
  let layout = pref(PREF_LAYOUT) || (window.innerWidth >= 1280 ? 'side' : 'stack');
  let left = +pref(PREF_LEFT) || LEFT_DEFAULT;
  const splits = [];
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

  function applyLayout() {
    splits.forEach(sp => sp.apply());
  }

  //   const sp = IMWorkbench.split({ preview, side: [nav, panels], grid, tools });
  //   sp.show(visible); sp.reveal(nav); sp.isSide()
  function split(o) {
    const wrap = document.createElement('div');
    wrap.className = 'wb-split';
    wrap.style.display = 'none';
    const main = document.createElement('div'); main.className = 'wb-split-main';
    const handle = document.createElement('div'); handle.className = 'wb-split-handle';
    handle.setAttribute('role', 'separator');
    handle.tabIndex = 0;
    const side = document.createElement('div'); side.className = 'wb-split-side';
    o.preview.parentNode.insertBefore(wrap, o.preview);
    main.appendChild(o.preview);
    o.side.forEach(el => { if (el) side.appendChild(el); });
    wrap.appendChild(main); wrap.appendChild(handle); wrap.appendChild(side);

    if (o.tools) {
      const tg = document.createElement('div');
      tg.className = 'wb-layout';
      tg.setAttribute('role', 'group');
      tg.setAttribute('aria-label', 'Layout');
      tg.innerHTML = '<button type="button" data-layout="side" title="Preview and fixes next to each other">Side by side</button>' +
        '<button type="button" data-layout="stack" title="Preview above the fixes">Stacked</button>';
      o.tools.insertBefore(tg, o.tools.firstChild);
      tg.addEventListener('click', e => {
        const b = e.target.closest('button[data-layout]');
        if (!b) return;
        layout = b.dataset.layout;
        keep(PREF_LAYOUT, layout);
        applyLayout();
      });
    }

    // A narrow window drops side by side to stacked in CSS; ask the browser.
    const sideNow = () => layout === 'side' && getComputedStyle(wrap).display === 'grid';

    function apply() {
      wrap.classList.toggle('is-side', layout === 'side');
      wrap.classList.toggle('is-stack', layout !== 'side');
      wrap.style.setProperty('--wb-left', left + '%');
      // Side by side wants the whole window, not the page's reading width.
      for (let p = wrap.parentElement; p; p = p.parentElement) {
        if (p.classList && p.classList.contains('cmp-page')) p.classList.toggle('wb-wide', layout === 'side');
      }
      wrap.querySelectorAll('.wb-layout button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.layout === layout)));
      const isSide = layout === 'side';
      handle.setAttribute('aria-orientation', isSide ? 'vertical' : 'horizontal');
      handle.setAttribute('aria-label', isSide ? 'Resize preview and fixes — drag, or use ← →' : 'Resize the preview grid — drag, or use ↑ ↓');
      handle.title = isSide ? 'Drag to resize · double-click to reset' : 'Drag to change the grid height · double-click to reset';
    }

    // Side by side, the pinned preview fits the window: its grid gets the
    // page's scroll height less the preview's own head (title, legend, filter
    // bar, rail), so its last rows and its foot stay on screen. The head's
    // height doesn't depend on the grid's, so re-measuring settles at once.
    // The grid on show. A page can hold more than one (Legacy's Locations /
    // Employees switch); measuring a hidden one (height 0) made fit() chase
    // its own tail — every redraw resized the preview, which resized it again.
    const gridNow = () => [...o.preview.querySelectorAll('.table-wrap')].find(g => g.offsetParent !== null) || o.grid;
    function fit() {
      const sc = document.getElementById('page-compare');
      if (!gridNow() || !sc || !o.preview.offsetHeight) return;
      const px = Math.max(160, sc.clientHeight - (o.preview.offsetHeight - gridNow().offsetHeight) - 16) + 'px';
      if (wrap.style.getPropertyValue('--wb-fit') !== px) wrap.style.setProperty('--wb-fit', px);
    }
    // fit() changes the grid's height, which the observer sees again; doing it
    // on the next frame keeps that from piling up into a "ResizeObserver loop".
    if (window.ResizeObserver) { let fitQueued = 0; new ResizeObserver(() => { if (!fitQueued) fitQueued = requestAnimationFrame(() => { fitQueued = 0; fit(); }); }).observe(o.preview); }
    window.addEventListener('resize', fit);

    handle.addEventListener('pointerdown', e => {
      if (!gridNow()) return;
      e.preventDefault();
      handle.setPointerCapture(e.pointerId);
      handle.classList.add('is-dragging');
      const r = wrap.getBoundingClientRect(), y0 = e.clientY, h0 = gridNow().getBoundingClientRect().height;
      const move = ev => {
        if (sideNow()) {
          left = clamp((ev.clientX - r.left) / r.width * 100, 30, 80);
          splits.forEach(sp => sp.wrap.style.setProperty('--wb-left', left + '%'));
        } else {
          gridNow().style.height = clamp(h0 + ev.clientY - y0, 120, window.innerHeight * 0.8) + 'px';
        }
      };
      const up = () => {
        handle.removeEventListener('pointermove', move);
        handle.removeEventListener('pointerup', up);
        handle.classList.remove('is-dragging');
        keep(PREF_LEFT, String(Math.round(left)));
      };
      handle.addEventListener('pointermove', move);
      handle.addEventListener('pointerup', up);
    });
    handle.addEventListener('keydown', e => {
      if (!gridNow()) return;
      if (sideNow() && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
        left = clamp(left + (e.key === 'ArrowLeft' ? -2 : 2), 30, 80);
        keep(PREF_LEFT, String(left));
        applyLayout();
        e.preventDefault();
      } else if (!sideNow() && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
        const h = gridNow().getBoundingClientRect().height;
        gridNow().style.height = clamp(h + (e.key === 'ArrowUp' ? -24 : 24), 120, window.innerHeight * 0.8) + 'px';
        e.preventDefault();
      }
    });
    handle.addEventListener('dblclick', () => {
      if (sideNow()) { left = LEFT_DEFAULT; keep(PREF_LEFT, String(left)); applyLayout(); }
      else if (gridNow()) gridNow().style.height = '';
    });

    // Bring the fixes into view: side by side, the top of the split (preview
    // and buttons side by side); stacked, the button row just under the pinned preview.
    function reveal(nav) {
      const sc = document.getElementById('page-compare');
      if (!sc || !sc.scrollTo) return;
      const target = sideNow() ? wrap : (nav || side);
      const offset = sideNow() ? 0 : o.preview.offsetHeight + 8;
      const top = sc.scrollTop + target.getBoundingClientRect().top - sc.getBoundingClientRect().top - offset;
      sc.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
    }

    const sp = { wrap: wrap, apply: apply, reveal: reveal, isSide: sideNow,
      show: v => { wrap.style.display = v ? '' : 'none'; } };
    splits.push(sp);
    apply();
    return sp;
  }

  // ─── Dropdown cells on demand ───
  // An editable preview's dropdown cell renders as a light button showing its
  // value, and becomes the real <select> only when you click it (or Tab to it
  // and press Enter) — it then opens at once. Rendering every cell as a full
  // <select> put each template's whole Site / Crop & Variety list in every row:
  // half a million <option>s on a 1,000-row file, which froze the page.
  // The same goes for the Smart Fixes tables, whose every row listed every
  // allowed value. The button carries the current choice in its `value`, so
  // code that reads `.value` (Apply, Apply all) works on the button or on the
  // <select> that replaces it alike; its data-* attributes carry over too.
  //   pickButton(val, attrs, conf)         → button HTML (attrs: data-*, title, style)
  //   wirePicks(table, optionsFor, conf)   once per table and kind;
  //     optionsFor(ci, button) → [values]
  //   conf (both): { kind: 'cell' (default) | 'fix', custom: token for "Type custom…",
  //                  blank: the empty choice's label, sort: options A→Z (fix only, default on) }
  //     cell — becomes a .ts-cell-select; "— blank —", keeps an off-list value as "(current)"
  //     fix  — becomes a Smart Fix picker with the button's classes; "— pick a value —"
  function pickButton(val, attrs, conf) {
    const fix = conf && conf.kind === 'fix';
    const blank = (conf && conf.blank) || (fix ? '— pick a value —' : '— blank —');
    const label = val ? esc(val) : '<span class="ts-cell-pick-blank">' + esc(blank) + '</span>';
    return '<button type="button" class="' + (fix ? 'wb-fix-pick ' + esc(conf.cls || '') : 'ts-cell-pick') +
      '" value="' + esc(val) + '" data-val="' + esc(val) + '"' + (attrs || '') + '>' + label + '</button>';
  }
  function wirePicks(table, optionsFor, conf) {
    if (typeof conf === 'string') conf = { custom: conf };
    conf = conf || {};
    const fix = conf.kind === 'fix';
    const flag = fix ? '_wbFixPicks' : '_wbPicks';
    if (!table || table[flag]) return;
    table[flag] = true;
    table.addEventListener('click', e => {
      const b = e.target.closest(fix ? '.wb-fix-pick' : '.ts-cell-pick');
      if (!b) return;
      const val = b.dataset.val || '', ci = +b.dataset.ci;
      let vals = (optionsFor(ci, b) || []).slice();
      if (fix && conf.sort !== false) vals.sort();
      const same = v => String(v).toLowerCase().trim() === val.toLowerCase().trim();
      const inList = !!val && vals.some(same);
      const blank = conf.blank || (fix ? '— pick a value —' : '— blank —');
      let opts = '<option value=""' + (inList ? '' : ' selected') + '>' + esc(blank) + '</option>';
      if (!fix && val && !inList) opts += '<option value="' + esc(val) + '" selected>' + esc(val) + '  (current)</option>';
      vals.forEach(v => { opts += '<option value="' + esc(v) + '"' + (inList && same(v) ? ' selected' : '') + '>' + esc(v) + '</option>'; });
      if (!fix) opts += '<option value="' + esc(conf.custom || '') + '">✎ Type custom…</option>';
      const sel = document.createElement('select');
      sel.className = fix ? (b.className.replace(/\bwb-fix-pick\b/, '').trim()) : 'ts-cell-select';
      Object.keys(b.dataset).forEach(k => { if (k !== 'val') sel.dataset[k] = b.dataset[k]; });
      if (b.title) sel.title = b.title;
      if (b.getAttribute('style')) sel.setAttribute('style', b.getAttribute('style'));
      sel.innerHTML = opts;
      b.replaceWith(sel);
      sel.focus();
      try { if (sel.showPicker) sel.showPicker(); } catch (err) { /* the focused select still opens on Space / Alt+↓ */ }
    });
  }

  // ─── Every row, drawn as you scroll ───
  // The preview holds every row of the file, but only the stretch in and near
  // view is in the page; a spacer row stands in for the rest at the measured
  // row height, so the scrollbar is true to the whole file. Scrolling redraws
  // only once the view leaves the drawn stretch. Rows come and go, so a tool's
  // handlers must be delegated on the <table>.
  //   IMWorkbench.rows(table, { head, ids, row, cols, empty, widths, keys, cell })
  //     head  '<thead>…</thead>'     ids   row indexes, in the tool's order
  //     row   (ri, pos) → '<tr>…</tr>'   cols  column count (spacers, empty note)
  //     empty text when ids is empty
  //     after   called after every draw (e.g. to repaint a selection)
  //     widths  px per column (fitWidth) — fixes the layout, so columns don't
  //             jump as rows with longer values scroll in, and a redraw
  //             doesn't re-measure every cell
  //     keys    a stable key per header cell (null = neither sorts nor resizes)
  //             — click a header to sort (A→Z, Z→A, off), drag its right edge
  //             to resize (double-click: back to fit); both are kept per table
  //             by key, across renders and filters
  //     cell    (ri, pos) → the value a column sorts by
  //   Call it again after any change — the scroll and the sort stay.
  //   → the row indexes in display order (sorted); also IMWorkbench.rowOrder(table)
  //   IMWorkbench.rowsTop(table)  → back to the first row
  //   IMWorkbench.fitWidth(header, values) → px for a column holding them
  // A wide grid also gets a horizontal scroll rail pinned to the bottom of the
  // window (rail()), so sideways scrolling never waits on reaching its foot.
  const OVER = 30;   // rows drawn beyond each edge of the view
  const MARGIN = 10; // redraw once the view comes this close to the drawn edge
  function rows(table, o) {
    if (!table) return [];
    const st = table._wbRows || (table._wbRows = { h: 0, from: -1, to: -1, gen: 0 });
    st.o = o;
    st.ids = sortedIds(table, o);
    const sc = table.parentElement; // the preview's scroll box (.table-wrap / .dt-pv-wrap)
    if (sc && st.sc !== sc) {
      st.sc = sc;
      sc.addEventListener('scroll', () => {
        if (st.raf) return;
        st.raf = requestAnimationFrame(() => { st.raf = 0; drawRows(table, false); });
      }, { passive: true });
    }
    const own = table._wbColW || {};
    const w = o.widths && o.widths.map((px, p) => { const k = o.keys && o.keys[p]; return k != null && own[k] ? own[k] : px; });
    table.style.tableLayout = w ? 'fixed' : '';
    table.style.width = w ? w.reduce((a, b) => a + b, 0) + 'px' : '';
    table.innerHTML = (w ? '<colgroup>' + w.map(px => '<col style="width:' + px + 'px">').join('') + '</colgroup>' : '') +
      o.head + '<tbody></tbody>';
    markHead(table, o, !!w);
    wireHead(table);
    st.from = st.to = -1;
    drawRows(table, true);
    rail(table);
    return st.ids;
  }
  function rowOrder(table) { return table && table._wbRows ? table._wbRows.ids || [] : []; }

  // Sorted by the column you picked: numbers as numbers, text A→Z with digits
  // in order ("Block 2" before "Block 10"), blanks last either way.
  const collate = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  const asNum = v => /^-?[\d,]*\.?\d+$/.test(v) ? parseFloat(v.replace(/,/g, '')) : NaN;
  function sortedIds(table, o) {
    const s = table._wbSort;
    const p = s && o.keys && o.cell ? o.keys.indexOf(s.key) : -1;
    if (p < 0) return o.ids;
    const val = new Map();
    o.ids.forEach(ri => { const v = o.cell(ri, p); val.set(ri, v == null ? '' : String(v).trim()); });
    return o.ids.slice().sort((a, b) => {
      const x = val.get(a), y = val.get(b);
      if (!x || !y) return (x ? 0 : 1) - (y ? 0 : 1);
      const nx = asNum(x), ny = asNum(y);
      return s.dir * (!isNaN(nx) && !isNaN(ny) ? nx - ny : collate.compare(x, y));
    });
  }
  // Header cells with a key sort on click and carry a resize grip.
  function markHead(table, o, resizable) {
    if (!o.keys || !table.tHead || !table.tHead.rows[0]) return;
    const s = table._wbSort;
    [...table.tHead.rows[0].cells].forEach((th, p) => {
      const k = o.keys[p];
      if (k == null) return;
      th.classList.add('wb-th');
      th.dataset.wbKey = k;
      th.dataset.wbPos = p;
      if (!th.title) th.title = 'Click to sort' + (resizable ? ' · drag the right edge to resize' : '');
      if (s && s.key === k) {
        th.classList.add('sort-active');
        th.dataset.sortDir = s.dir > 0 ? 'asc' : 'desc';
        th.setAttribute('aria-sort', s.dir > 0 ? 'ascending' : 'descending');
      }
      if (resizable) th.insertAdjacentHTML('beforeend', '<span class="wb-col-grip" aria-hidden="true"></span>');
    });
  }
  function wireHead(table) {
    if (table._wbHeadWired) return;
    table._wbHeadWired = true;
    table.addEventListener('click', e => {
      const th = e.target.closest('th.wb-th');
      if (!th || e.target.closest('.wb-col-grip') || !table.tHead || !table.tHead.contains(th)) return;
      const k = th.dataset.wbKey, s = table._wbSort;
      table._wbSort = !s || s.key !== k ? { key: k, dir: 1 } : s.dir > 0 ? { key: k, dir: -1 } : null;
      rowsTop(table);
      rows(table, table._wbRows.o);
    });
    table.addEventListener('pointerdown', e => {
      const g = e.target.closest('.wb-col-grip');
      if (!g) return;
      e.preventDefault();
      e.stopPropagation();
      const th = g.closest('th'), k = th.dataset.wbKey;
      const col = table.querySelectorAll('colgroup col')[+th.dataset.wbPos];
      if (!col) return;
      const x0 = e.clientX, w0 = parseFloat(col.style.width) || th.offsetWidth;
      const t0 = parseFloat(table.style.width) || table.offsetWidth;
      g.setPointerCapture(e.pointerId);
      document.body.classList.add('wb-resizing');
      const move = ev => {
        const w = Math.max(48, Math.round(w0 + ev.clientX - x0));
        col.style.width = w + 'px';
        table.style.width = (t0 - w0 + w) + 'px';
      };
      const up = () => {
        g.removeEventListener('pointermove', move);
        g.removeEventListener('pointerup', up);
        document.body.classList.remove('wb-resizing');
        (table._wbColW || (table._wbColW = {}))[k] = parseFloat(col.style.width);
        railSize(table);
      };
      g.addEventListener('pointermove', move);
      g.addEventListener('pointerup', up);
    });
    table.addEventListener('dblclick', e => {
      const g = e.target.closest('.wb-col-grip');
      if (!g) return;
      if (table._wbColW) delete table._wbColW[g.closest('th').dataset.wbKey];
      rows(table, table._wbRows.o);
    });
  }

  // The rail: a slim scrollbar after the grid, position: sticky to the bottom
  // of the page's scroll, kept in step with the grid both ways. While the
  // grid's foot is in view it sits right under it; otherwise it rides the
  // bottom of the window. The grid's own horizontal bar is hidden (CSS) so
  // there's only ever the one.
  function rail(table) {
    const sc = table._wbRows && table._wbRows.sc;
    if (!sc || !sc.after) return;
    if (!sc._wbRail) {
      const r = document.createElement('div');
      r.className = 'wb-hrail';
      r.setAttribute('aria-hidden', 'true'); // the grid itself still scrolls by keyboard
      r.innerHTML = '<div></div>';
      sc.after(r);
      sc.classList.add('wb-hrail-host');
      sc._wbRail = r;
      sc.addEventListener('scroll', () => { if (r.scrollLeft !== sc.scrollLeft) r.scrollLeft = sc.scrollLeft; }, { passive: true });
      r.addEventListener('scroll', () => { if (sc.scrollLeft !== r.scrollLeft) sc.scrollLeft = r.scrollLeft; }, { passive: true });
      if (window.ResizeObserver) new ResizeObserver(() => railSize(table)).observe(sc);
    }
    railSize(table);
  }
  function railSize(table) {
    const st = table._wbRows, sc = st && st.sc, r = sc && sc._wbRail;
    if (!r) return;
    const over = sc.scrollWidth - sc.clientWidth;
    r.style.display = over > 1 ? '' : 'none';
    if (over <= 1) return;
    r.firstChild.style.width = (r.clientWidth + over) + 'px';
    r.scrollLeft = sc.scrollLeft;
  }

  // Wide enough for the longest value (or the header, set in small caps),
  // within reason; a longer value is cut off in view, never in the export.
  function fitWidth(header, values, min, max) {
    let n = String(header == null ? '' : header).length * 1.15;
    for (const v of values) { const l = v == null ? 0 : String(v).length; if (l > n) n = l; }
    return Math.round(Math.min(max || 340, Math.max(min || 104, n * 7.2 + 34)));
  }
  function rowsTop(table) {
    const st = table && table._wbRows;
    if (st && st.sc) st.sc.scrollTop = 0;
  }
  function drawRows(table, force) {
    const st = table._wbRows, o = st.o, ids = st.ids, tb = table.tBodies[0];
    if (!tb) return;
    const n = ids.length;
    if (!n) {
      tb.innerHTML = '<tr><td class="wb-rows-none" colspan="' + o.cols + '">' + esc(o.empty || 'No rows to show.') + '</td></tr>';
      st.from = st.to = 0;
      return;
    }
    const h = st.h || 36, sc = st.sc;
    const head = table.tHead ? table.tHead.offsetHeight : 0;
    const top = sc ? Math.max(0, sc.scrollTop - head) : 0, vh = sc && sc.clientHeight ? sc.clientHeight : 800;
    const first = Math.min(n - 1, Math.floor(top / h)), last = Math.min(n, Math.ceil((top + vh) / h));
    if (!force && (first >= st.from + MARGIN || st.from === 0) && (last <= st.to - MARGIN || st.to === n)) return;
    const from = Math.max(0, first - OVER), to = Math.min(n, last + OVER);
    const rowsHtml = (a, b) => { let s = ''; for (let i = a; i < b; i++) s += o.row(ids[i], i); return s; };
    const sized = (tr, px) => { tr.firstChild.style.height = px + 'px'; };
    if (!force && st.gt && from < st.to && to > st.from) {
      // Overlaps what's drawn: drop the rows leaving the stretch, add the ones
      // entering it — a scroll step touches a few rows, not all of them.
      // A cell being typed in would vanish with its row — commit it first
      // (its change handler may redraw on its own; then this draw is moot).
      const act = document.activeElement;
      if (act && tb.contains(act) && act.blur) {
        const ri = act.dataset ? act.dataset.ri : null, pos = ri == null ? -1 : ids.indexOf(+ri);
        if (pos < from || pos >= to) { const gen = st.gen; act.blur(); if (st.gen !== gen) return; }
      }
      for (let k = st.from; k < from; k++) st.gt.nextElementSibling.remove();
      for (let k = to; k < st.to; k++) st.gb.previousElementSibling.remove();
      if (from < st.from) st.gt.insertAdjacentHTML('afterend', rowsHtml(from, st.from));
      if (to > st.to) st.gb.insertAdjacentHTML('beforebegin', rowsHtml(st.to, to));
      sized(st.gt, from * h); sized(st.gb, (n - to) * h);
      st.from = from; st.to = to; st.gen++;
      if (o.after) o.after();
      return;
    }
    const act = document.activeElement;
    if (!force && act && tb.contains(act) && act.blur) { const gen = st.gen; act.blur(); if (st.gen !== gen) return; }
    // Two spacer rows always bracket the drawn stretch (0 px when at an end).
    const gap = px => '<tr class="wb-gap" aria-hidden="true"><td colspan="' + o.cols + '" style="height:' + px + 'px"></td></tr>';
    tb.innerHTML = gap(from * h) + rowsHtml(from, to) + gap((n - to) * h);
    st.gt = tb.firstElementChild; st.gb = tb.lastElementChild;
    st.from = from; st.to = to; st.gen++;
    if (o.after) o.after();
    // Measure the real row height once rows are on screen, then redraw at it.
    if (!force) return;
    let tot = 0, cnt = 0;
    for (const tr of tb.rows) { if (!tr.classList.contains('wb-gap')) { tot += tr.offsetHeight; cnt++; } }
    const real = cnt && tot ? tot / cnt : 0;
    if (real && Math.abs(real - h) > 0.5) { st.h = real; drawRows(table, true); }
  }

  // ─── Preview filter ───
  // One bar for every tool: which rows (all · needs a fix · changed · either),
  // "only columns with something to see", and — when a panel sent you to
  // specific rows ("Show the 1 empty row") — a chip naming them, × to clear.
  //   filterBar({ mode, counts: { all, fix, chg, any }, colsOnly, shownCols, totalCols, focus })
  //   rowMatch(mode, fix, chg) → bool
  // Controls carry .wb-rowfilter / .wb-colfilter / .wb-focus-clear for the
  // tool's delegated handlers.
  const ROW_FILTERS = [['all', 'All rows'], ['fix', 'Needs a fix'], ['chg', 'Changed'], ['any', 'Needs a fix or changed']];
  function filterBar(o) {
    const c = o.counts || {};
    const rowsCtl = o.focus
      ? '<span class="wb-focus">' + esc(o.focus) + '<button type="button" class="wb-focus-clear" title="Show every row again" aria-label="Clear">&times;</button></span>'
      : '<select class="input-field wb-rowfilter" aria-label="Rows to show">' + ROW_FILTERS.map(([k, l]) =>
          '<option value="' + k + '"' + (o.mode === k ? ' selected' : '') + '>' + l + ' (' + (c[k] || 0) + ')</option>').join('') + '</select>';
    return '<span class="wb-filter">' + rowsCtl +
      '<label class="wb-colfilter-wrap" title="Hide columns where none of the rows shown has a flag or a change">' +
        '<input type="checkbox" class="wb-colfilter"' + (o.colsOnly ? ' checked' : '') + '> Only columns with something to see' +
        (o.colsOnly && o.totalCols ? ' <span class="wb-colfilter-n">(' + o.shownCols + ' of ' + o.totalCols + ')</span>' : '') +
      '</label></span>';
  }
  function rowMatch(mode, fix, chg) {
    return mode === 'fix' ? !!fix : mode === 'chg' ? !!chg : mode === 'any' ? !!(fix || chg) : true;
  }

  // ─── Text cells on demand ───
  // A free-text preview cell is plain text until you click it or Tab to it;
  // then it becomes the .ts-cell-input it stands for, focused, so the tool's
  // delegated change handler commits the edit as before. Thousands of live
  // <input>s made every scroll redraw cost ~90 ms; text costs a fraction.
  //   textCell(val, attrs)  → HTML (attrs: data-ri / data-ci, title, style)
  //   wireTextCells(table)  once per table
  function textCell(val, attrs) {
    return '<span class="ts-cell-text" tabindex="0" data-val="' + esc(val) + '"' + (attrs || '') + '>' + esc(val) + '</span>';
  }
  function wireTextCells(table) {
    if (!table || table._wbText) return;
    table._wbText = true;
    const open = t => {
      const inp = document.createElement('input');
      inp.type = 'text';
      inp.className = 'ts-cell-input';
      inp.value = t.dataset.val || '';
      Object.keys(t.dataset).forEach(k => { if (k !== 'val') inp.dataset[k] = t.dataset[k]; });
      if (t.title) inp.title = t.title;
      if (t.getAttribute('style')) inp.setAttribute('style', t.getAttribute('style'));
      t.replaceWith(inp);
      inp.focus();
      try { inp.setSelectionRange(inp.value.length, inp.value.length); } catch (err) { /* not every input type allows it */ }
    };
    // Tab lands on it (focusin); a click opens it even where focus events
    // don't fire (a page in the background).
    table.addEventListener('focusin', e => { const t = e.target.closest('.ts-cell-text'); if (t) open(t); });
    table.addEventListener('click', e => { const t = e.target.closest('.ts-cell-text'); if (t) open(t); });
    table.addEventListener('keydown', e => {
      // Enter in an open cell commits it.
      if (e.key !== 'Enter' || !e.target.classList.contains('ts-cell-input')) return;
      e.preventDefault();
      e.target.blur();
    });
  }

  // "Was X → now Y · why" for a changed cell, or '' when the cell holds what
  // the source said (a change of case alone — matching a dropdown's spelling —
  // doesn't count).
  function changeNote(was, now, why) {
    const a = String(was == null ? '' : was).trim(), b = String(now == null ? '' : now).trim();
    if (a === b || a.toUpperCase() === b.toUpperCase()) return '';
    return 'Was ' + (a ? '"' + a + '"' : '(blank)') + ' → now ' + (b ? '"' + b + '"' : '(blank)') + (why ? ' · ' + why : '');
  }

  return { mount: mount, split: split, changeNote: changeNote, pickButton: pickButton, wirePicks: wirePicks,
    rows: rows, rowsTop: rowsTop, rowOrder: rowOrder, fitWidth: fitWidth, filterBar: filterBar, rowMatch: rowMatch,
    textCell: textCell, wireTextCells: wireTextCells };
})();
