// ═══════════════════════════════════════════════════════════════════════
// "Load everything at once" for the Legacy → 3.0 tools — the Bulk Template
// Standardize idea, one tool at a time: pick (or drop) every file together,
// each is identified by its header row and goes into its slot. Anything that
// can't be told apart lands in a short list with a "What is it?" picker.
//
//   IMLegacyDrop.mount({
//     host: 'lm-drop',                       // the element to fill
//     kinds: { loc: 'Legacy locations', … }, // slot key → label, for the picker
//     order: ['loctpl', 'sitetpl', 'loc', …],// placement order (templates first)
//     classify: info => key | null,          // info = { name, headers, norm, sheetNames, isTemplate, rows }
//     arrange: infos => void,                // optional: settle ties across files (sets info.key)
//     elsewhere: info => 'why' | null,       // optional: a file for another tool — skipped, not asked about
//     place: (key, file) => Promise,         // load the file into that slot
//     answers: file => Promise<message>,     // optional: a .html page the client answered
//     accept: '.csv,.xlsx,.xls,.html',     // optional: the file picker's filter
//     log: (kind, msg) => void               // optional: the module's event log
//   });
// The file's slot handler does the real reading; this only peeks at headers.
// ═══════════════════════════════════════════════════════════════════════
window.IMLegacyDrop = (function () {
  'use strict';

  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const normH = h => String(h == null ? '' : h).replace(/^﻿/, '').trim().toLowerCase()
    .replace(/^#/, '').replace(/\*$/, '').trim();

  // Header row + sheet names, cheaply. A workbook with a DATA ENTRY sheet and
  // a DROP-DOWN INPUTS sheet is a personalized template; its headers come
  // from DATA ENTRY. Anything else: the first sheet's first non-empty row.
  function peek(file) {
    return new Promise(resolve => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const wb = /\.csv$/i.test(file.name)
            ? XLSX.read(String(e.target.result), { type: 'string', sheetRows: 6 })
            : XLSX.read(new Uint8Array(e.target.result), { type: 'array', sheetRows: 6 });
          const names = wb.SheetNames || [];
          const de = names.find(n => /data.?entry/i.test(n));
          const isTemplate = !!de && names.some(n => /drop.?down/i.test(n));
          const ws = wb.Sheets[de || names[0]];
          const aoa = ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false }) : [];
          const head = (aoa.find(row => row && row.some(c => String(c).trim() !== '')) || []).map(h => String(h == null ? '' : h).trim());
          resolve({ file, name: file.name, headers: head, norm: head.map(normH), sheetNames: names, isTemplate, size: file.size });
        } catch (err) {
          resolve({ file, name: file.name, headers: [], norm: [], sheetNames: [], isTemplate: false, size: file.size, error: err.message });
        }
      };
      r.onerror = () => resolve({ file, name: file.name, headers: [], norm: [], sheetNames: [], isTemplate: false, size: file.size, error: 'unreadable' });
      if (/\.csv$/i.test(file.name)) r.readAsText(file); else r.readAsArrayBuffer(file);
    });
  }

  function mount(o) {
    const host = document.getElementById(o.host);
    if (!host || host._lgDrop) return null;
    host._lgDrop = true;
    const inputId = o.host + '-input';
    host.innerHTML =
      '<div class="cmp-slot-label">Load everything at once</div>' +
      '<label class="btn btn-primary cmp-upload-btn lg-drop-btn">' +
        '<span>&#8679; Upload all files</span>' +
        '<input type="file" id="' + inputId + '" accept="' + (o.accept || '.csv,.xlsx,.xls') + '" multiple hidden></label>' +
      '<div class="cmp-slot-hint">Pick (or drag here) every file at once &mdash; each one is recognised by its column headers and goes into its slot below. ' +
        'Files it can’t place are listed here so you can say what they are.' + (o.answers ? ' The page the client sends back with their answers goes here too.' : '') + '</div>' +
      '<div class="lg-drop-result" aria-live="polite"></div>';
    const out = host.querySelector('.lg-drop-result');
    let pending = [];   // files waiting for a "What is it?" answer

    function renderResult(placed, skipped) {
      let html = '';
      if (placed.length) html += '<ul class="lg-drop-list">' + placed.map(p =>
        '<li class="is-ok">✓ <b>' + esc(p.name) + '</b> → ' + esc(o.kinds[p.key]) + '</li>').join('') + '</ul>';
      if (pending.length) {
        const opts = '<option value="">Not used here — skip it</option>' +
          Object.keys(o.kinds).map(k => '<option value="' + k + '">' + esc(o.kinds[k]) + '</option>').join('');
        html += '<div class="lg-drop-ask"><b>What is ' + (pending.length === 1 ? 'this file' : 'each of these') + '?</b><ul class="lg-drop-list">' +
          pending.map((p, i) => '<li class="is-ask"><span>' + esc(p.name) +
            (p.headers.length ? ' <span class="text-muted small">(' + esc(p.headers.filter(Boolean).slice(0, 6).join(', ')) + (p.headers.length > 6 ? '…' : '') + ')</span>' : '') +
            '</span> <select class="input-field lg-drop-pick" data-i="' + i + '">' +
            opts.replace('value="' + (p.guess || '§') + '"', 'value="' + (p.guess || '§') + '" selected') + '</select></li>').join('') +
          '</ul><button type="button" class="btn btn-primary btn-sm lg-drop-place">Place ' + (pending.length === 1 ? 'it' : 'them') + '</button></div>';
      }
      if (skipped.length) html += '<div class="cmp-slot-hint"><b>Skipped:</b><br>' + skipped.map(esc).join('<br>') + '</div>';
      out.innerHTML = html;
    }

    async function placeAll(list) {
      // The tool's own order: templates first, so exports meet their lists.
      const rank = k => { const i = (o.order || []).indexOf(k); return i < 0 ? 99 : i; };
      list.sort((a, b) => rank(a.key) - rank(b.key));
      for (const p of list) {
        try { await o.place(p.key, p.file); }
        catch (err) { if (o.log) o.log('error', 'Could not load ' + p.name + ' as ' + o.kinds[p.key] + ': ' + (err && err.message ? err.message : err)); }
      }
    }

    async function handle(files) {
      let list = Array.from(files || []);
      if (!list.length) return;
      // A page the client answered (.html) isn't a slot file: it's handed to
      // o.answers, after the other files are placed (they may be what it needs).
      const pages = o.answers ? list.filter(f => /\.html?$/i.test(f.name)) : [];
      list = list.filter(f => pages.indexOf(f) < 0);
      if (!list.length) {
        out.innerHTML = '';
        for (const f of pages) out.insertAdjacentHTML('beforeend', '<ul class="lg-drop-list"><li class="is-ok">' + esc(await o.answers(f)) + '</li></ul>');
        return;
      }
      out.innerHTML = '<div class="cmp-slot-hint">Reading ' + list.length + ' file' + (list.length === 1 ? '' : 's') + '…</div>';
      const infos = await Promise.all(list.map(peek));
      infos.forEach(i => {
        i.key = i.error ? null : o.classify(i);
        i.away = !i.key && !i.error && o.elsewhere ? o.elsewhere(i) : null;
      });
      if (o.arrange) o.arrange(infos);
      // Two files for one slot: keep the first, ask about the other.
      const taken = new Set();
      infos.forEach(i => {
        if (!i.key) return;
        if (taken.has(i.key)) { i.guess = i.key; i.key = null; return; }
        taken.add(i.key);
      });
      const placed = infos.filter(i => i.key);
      pending = infos.filter(i => !i.key && !i.away);
      const away = infos.filter(i => i.away).map(i => i.name + ' — ' + i.away);
      if (o.log) o.log('note', 'Load everything: ' + infos.map(i => i.name + ' → ' + (i.key ? o.kinds[i.key] : i.away ? 'skipped (' + i.away + ')' : 'asked what it is')).join('; '));
      renderResult(placed, away);
      await placeAll(placed);
      for (const f of pages) out.insertAdjacentHTML('beforeend', '<ul class="lg-drop-list"><li class="is-ok">' + esc(await o.answers(f)) + '</li></ul>');
    }

    host.querySelector('#' + inputId).addEventListener('change', e => { handle(e.target.files); e.target.value = ''; });
    host.addEventListener('dragover', e => { e.preventDefault(); host.classList.add('is-over'); });
    host.addEventListener('dragleave', e => { if (!host.contains(e.relatedTarget)) host.classList.remove('is-over'); });
    host.addEventListener('drop', e => {
      e.preventDefault();
      host.classList.remove('is-over');
      if (e.dataTransfer && e.dataTransfer.files) handle(e.dataTransfer.files);
    });
    host.addEventListener('click', async e => {
      if (!e.target.closest('.lg-drop-place')) return;
      const picks = [], skipped = [];
      host.querySelectorAll('.lg-drop-pick').forEach(sel => {
        const p = pending[+sel.dataset.i];
        if (!p) return;
        if (sel.value) picks.push(Object.assign(p, { key: sel.value })); else skipped.push(p.name);
      });
      pending = [];
      renderResult(picks, skipped);
      await placeAll(picks);
    });
    return { reset() { pending = []; out.innerHTML = ''; } };
  }

  // Download a Blob under a file name.
  function save(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  // "Up to 8 examples, …" for a step's question.
  const examples = (list, n) => { const u = [...new Set(list.filter(Boolean))]; return u.slice(0, n || 8).join(', ') + (u.length > (n || 8) ? ', …' : ''); };

  // ─── Highlights + notes for one exported file set ───
  // Turns a tool's flagged cells into what the client sees:
  //   · every flagged cell coloured (red = fill this in, amber = check this)
  //   · a hover note on the column header saying what the column needs, and on
  //     each cell whose question is its own ("Its Legacy site (#28) was archived…")
  //   · one read-me item per column
  // When every flagged cell in a column asks the same thing, the header note
  // carries it and the cells stay quiet — one note, not 369 copies.
  //   annotate({ file: (chunk, total) → name, label, unit: ['block', 'blocks'], H, total, limit,
  //              cells: pos → [{ ci, flag, note }], nameOf: pos → text,
  //              ask: (ci, flag) → { title, detail } (or a string),
  //              options: ci → [values the column's dropdown offers],
  //              headerOnly: [{ ci, title, detail, count }] })   // questions with no highlight
  //   → { chunks: [{ marks, notes }], items }
  const colLetter = ci => { let s = '', n = ci + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  function annotate(o) {
    const nChunks = Math.max(1, Math.ceil(o.total / o.limit));
    const chunks = Array.from({ length: nChunks }, () => ({ marks: new Map(), notes: new Map() }));
    const byCol = new Map();   // "ci|flag" → [{ pos, note }]
    for (let pos = 0; pos < o.total; pos++) {
      (o.cells(pos) || []).forEach(c => {
        const k = c.ci + '|' + c.flag;
        (byCol.get(k) || byCol.set(k, []).get(k)).push({ pos, note: c.note || '' });
        const ch = chunks[Math.floor(pos / o.limit)], at = pos % o.limit;
        const m = ch.marks.get(at) || new Map();
        if (m.get(c.ci) !== 'req') m.set(c.ci, c.flag);
        ch.marks.set(at, m);
      });
    }
    const headerNote = (ci, text) => chunks.forEach(ch => {
      const m = ch.notes.get(-1) || new Map();
      m.set(ci, m.has(ci) ? m.get(ci) + '\n\n' + text : text);
      ch.notes.set(-1, m);
    });
    const askOf = (ci, flag) => { const a = o.ask(ci, flag); return typeof a === 'string' ? { title: a, detail: '' } : a; };
    const base = (ci, flag) => ({ file: o.label, unit: o.unit || ['row', 'rows'], column: o.H[ci].replace(/\*$/, ''), col: colLetter(ci),
      flag, options: o.options ? (o.options(ci) || []) : [] });
    const items = [];
    // Red first, then amber; within each, left to right.
    [...byCol.keys()].sort((a, b) => {
      const [ca, fa] = a.split('|'), [cb, fb] = b.split('|');
      return (fa === fb ? 0 : fa === 'req' ? -1 : 1) || (+ca - +cb);
    }).forEach(k => {
      const ci = +k.split('|')[0], flag = k.split('|')[1];
      const list = byCol.get(k);
      const notes = [...new Set(list.map(x => x.note))];
      const same = notes.length <= 1;
      const a = askOf(ci, flag);
      headerNote(ci, (flag === 'req' ? 'Red' : 'Amber') + ' cells: ' + a.title + (a.detail ? '\n' + a.detail : '') +
        (same ? (notes[0] ? '\n' + notes[0] : '') : '\nHover each highlighted cell for its own note.'));
      if (!same) list.forEach(x => {
        if (!x.note) return;
        const ch = chunks[Math.floor(x.pos / o.limit)], at = x.pos % o.limit;
        const m = ch.notes.get(at) || new Map();
        m.set(ci, m.has(ci) ? m.get(ci) + '\n' + x.note : x.note);
        ch.notes.set(at, m);
      });
      items.push(Object.assign(base(ci, flag), { title: a.title, detail: a.detail + (same && notes[0] ? (a.detail ? ' ' : '') + notes[0] : ''),
        // Rows are listed when they differ — or when there are few enough to name.
        count: list.length, rows: same && list.length > 12 ? []
          : list.map(x => ({ row: (x.pos % o.limit) + 2, name: o.nameOf(x.pos), note: same ? '' : x.note })) }));
    });
    (o.headerOnly || []).forEach(h => {
      headerNote(h.ci, 'Question: ' + h.title + (h.detail ? '\n' + h.detail : ''));
      items.push(Object.assign(base(h.ci, 'ask'), { title: h.title, detail: h.detail || '', count: h.count, rows: [], options: [] }));
    });
    return { chunks, items };
  }


  // Header tests the tools share.
  const has = (info, ...ks) => ks.every(k => info.norm.includes(k));
  const hasAny = (info, ...ks) => ks.some(k => info.norm.includes(k));

  return { mount: mount, peek: peek, has: has, hasAny: hasAny, save: save, examples: examples,
    annotate: annotate };
})();
