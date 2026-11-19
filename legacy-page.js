// ═══════════════════════════════════════════════════════════════════════
// Legacy → 3.0, one page: sites, locations and employees together.
//
// The two engines stay where they were — locations-migrate.js (sites +
// locations) and employee-migrate.js (employees) — each still reading its
// own files, keeping its own state and drawing its own panels. This file
// makes them one page:
//   · one "Upload all files" that hands each file to the engine that reads it
//   · one workbench: every section button from both, one pinned preview with
//     a Locations | Sites | Employees switch
//   · one Build, one Export all (one .zip + a read-me for the client),
//     one Reset, one Debug Dump
//
// Each engine registers on init:
//   IMLegacyPage.register('lm' | 'em', {
//     ready(), built(), rows(),          // files enough to build / has a preview / rows kept
//     count(sectionKey) → n,             // its share of a section button's count
//     onOpen(sectionKey),                // optional: a section of its was opened
//     run(), reset(),
//     exportParts(limit) → Promise<{ files: [{ label, name, data }], items, team: [lines] }>,
//     drop: { kinds, order, classify, arrange?, place }
//   });
// and calls IMLegacyPage.refresh() after it renders.
// ═══════════════════════════════════════════════════════════════════════
window.IMLegacyPage = (function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const apis = {};
  let wb = null, drop = null, view = 'lm';

  // Section buttons. `who` says whose panels they show: lm, em, or both.
  const SECTIONS = [
    { key: 'sitemap', who: 'lm', group: 'fix', unit: 'site ID', name: 'Site mapping', desc: 'Legacy site_ids with no site name — pick or type the 3.0 Site', ids: ['lm-section-sitemap'] },
    { key: 'crops', who: 'lm', group: 'fix', unit: 'crop ID', name: 'Crops', desc: 'Legacy crop_ids → the 3.0 Crop & Variety each becomes', ids: ['lm-section-crop'] },
    { key: 'types', who: 'lm', group: 'fix', unit: 'type ID', name: 'Location types', desc: 'Legacy location_type_ids → the 3.0 Location Type', ids: ['lm-section-type'] },
    { key: 'sitefile', who: 'lm', group: 'fix', unit: 'site', name: 'Sites file', desc: 'Site Type, Employer, Country for every site', ids: ['lm-section-sitefile'] },
    { key: 'siteaddr', who: 'lm', group: 'fix', unit: 'site', name: 'Site addresses', desc: 'One for every site, per ranch, per site — or the client’s sheet', ids: ['lm-section-siteaddr'] },
    { key: 'employers', who: 'em', group: 'fix', unit: 'contractor', name: 'Employers', desc: 'Legacy Contractor IDs with no 3.0 employer yet', ids: ['em-section-unresolved'] },
    { key: 'required', who: 'both', group: 'fix', unit: 'column', name: 'Required blanks', desc: 'Fill required (*) columns — or bulk-edit any column', ids: ['lm-section-bulk', 'em-section-bulk'] },
    { key: 'dropdown', who: 'both', group: 'fix', unit: 'value', name: 'Not in dropdown', desc: 'Pick a listed value for anything PickTrace would reject', ids: ['lm-section-smartfix', 'em-section-smartfix'] },
    { key: 'build', who: 'both', group: 'fix', unit: 'value', name: 'Build in PickTrace', desc: 'Sites, crops, types and employers to create first', ids: ['lm-section-create', 'lm-section-cv-create', 'em-section-create'] },
    { key: 'collisions', who: 'lm', group: 'fix', unit: 'row', name: 'Name collisions', desc: 'Locations sharing a Site + Name, sites sharing a Name', ids: ['lm-section-collisions'] },
    { key: 'possible', who: 'em', group: 'fix', unit: 'row', name: 'Possible duplicates', desc: 'Employees matching someone on SSN or name — check each', ids: ['em-section-possible'] },
    { key: 'names', who: 'em', group: 'check', unit: 'name', name: 'Legacy codes in names', desc: 'Names ending in {1234} — keep the code or take it out', ids: ['em-section-names'] },
    { key: 'archived', who: 'em', group: 'check', unit: 'employee', name: 'Archived contractors', desc: 'Employees of archived contractors — left out unless restored', ids: ['em-section-archived'] },
    { key: 'existing', who: 'em', group: 'check', unit: 'employee', name: 'Employees already in 3.0', desc: 'Matched in the 3.0 export — left out', ids: ['em-section-existing'] },
    { key: 'left', who: 'lm', group: 'check', unit: 'row', name: 'Left out', desc: 'Locations / sites archived, already in 3.0, or removed by you', ids: ['lm-section-left'] }
  ];
  const count = (who, key) => ['lm', 'em'].filter(k => who === 'both' || who === k)
    .reduce((n, k) => n + (apis[k] && apis[k].built() ? (+apis[k].count(key) || 0) : 0), 0);

  function register(key, api) {
    apis[key] = api;
    if (apis.lm && apis.em) mount();
  }

  // ─── The preview switch ───
  function setView(v) {
    view = v === 'employees' ? 'em' : 'lm';
    const pv = $('lm-section-preview');
    if (pv) pv.classList.toggle('lg-view-em', view === 'em');
    paintPills();
  }
  function paintPills() {
    const dom = apis.lm && apis.lm.domain ? apis.lm.domain() : 'locations';
    document.querySelectorAll('.lm-domain-pill').forEach(b => {
      const d = b.dataset.domain;
      b.classList.toggle('ts-domain-active', view === 'em' ? d === 'employees' : d === dom);
      if (d === 'employees') b.textContent = '👤 Employees' + (apis.em && apis.em.built() ? ' (' + apis.em.rows() + ')' : '');
    });
  }

  function mount() {
    if (wb || !window.IMWorkbench) return;
    wb = IMWorkbench.mount({
      pane: 'cmp-mode-legacy', results: 'lm-results', preview: 'lm-section-preview',
      ready: () => ['lm', 'em'].some(k => apis[k].built()),
      onOpen: key => {
        const s = SECTIONS.find(x => x.key === key);
        if (!s) return;
        if (s.who === 'em') setView('employees');
        else if (s.who === 'lm') { setView('lm'); if (apis.lm.onOpen) apis.lm.onOpen(key); }
      },
      sections: SECTIONS.map(s => ({ key: s.key, group: s.group, unit: s.unit, name: s.name, desc: s.desc, ids: s.ids,
        count: () => count(s.who, s.key) }))
    });
    // Employees pill: the engines' own pills (Locations / Sites) are handled
    // by locations-migrate.js; this one, and leaving it, are handled here.
    document.querySelectorAll('.lm-domain-pill').forEach(b => b.addEventListener('click', () => {
      setView(b.dataset.domain === 'employees' ? 'employees' : 'lm');
      if (wb) wb.refresh();
    }));

    // One "Upload all files" for both engines. Keys are "lm:loc", "em:emp"…
    if (window.IMLegacyDrop) {
      const kinds = {}, order = [];
      ['lm', 'em'].forEach(k => {
        const d = apis[k].drop;
        Object.keys(d.kinds).forEach(x => { kinds[k + ':' + x] = d.kinds[x]; });
        (d.order || []).forEach(x => order.push(k + ':' + x));
      });
      drop = IMLegacyDrop.mount({
        host: 'lg-drop', kinds, order,
        classify: info => {
          for (const k of ['lm', 'em']) {
            const key = apis[k].drop.classify(info);
            if (key) return k + ':' + key;
            if (info.guess && !/:/.test(info.guess)) info.guess = k + ':' + info.guess;
          }
          return null;
        },
        // Settle ties within one engine (its contractor lists): strip the
        // prefix, let it decide, put the prefix back.
        arrange: infos => ['lm', 'em'].forEach(k => {
          const d = apis[k].drop;
          if (!d.arrange) return;
          const mine = infos.filter(i => i.key && i.key.indexOf(k + ':') === 0);
          mine.forEach(i => { i.key = i.key.slice(k.length + 1); });
          d.arrange(mine);
          mine.forEach(i => {
            if (i.key) i.key = k + ':' + i.key;
            if (i.guess && !/:/.test(i.guess)) i.guess = k + ':' + i.guess;
          });
        }),
        elsewhere: i => IMLegacyDrop.hasAny(i, 'job') && IMLegacyDrop.hasAny(i, 'rate', 'pay type', 'piece rate')
          ? 'a jobs export, not part of Legacy → 3.0' : null,
        place: (key, file) => { const k = key.slice(0, 2); return apis[k].drop.place(key.slice(3), file); },
        answers: file => applyAnswersFile(file),
        accept: '.csv,.xlsx,.xls,.html',
        log: (kind, msg) => { if (apis.lm.log) apis.lm.log(kind, msg); }
      });
    }

    $('lm-run').addEventListener('click', () => ['lm', 'em'].forEach(k => { if (apis[k].ready()) apis[k].run(); }));
    $('lm-export').addEventListener('click', () => {
      exportAll().catch(err => { if (apis.lm.fail) apis.lm.fail('Export failed', err); else alert('Export failed: ' + (err && err.message)); });
    });
    $('lm-reset').addEventListener('click', () => {
      ['lm', 'em'].forEach(k => apis[k].reset());
      if (drop) drop.reset();
      showComments([]);
      setView('lm');
      if (wb) wb.reset();
      refresh();
    });
    if (window.IMDebug) {
      // One dump for the page: both engines' dumps side by side.
      IMDebug.register('legacy-migrate', {
        label: 'Legacy → 3.0 (sites, locations, employees)',
        ready: () => ['lm', 'em'].some(k => apis[k].debugReady && apis[k].debugReady()),
        collect: () => {
          const a = IMDebug.build('locations-migrate'), b = IMDebug.build('employee-migrate');
          return {
            inputs: [].concat((a && a.inputs) || [], (b && b.inputs) || []),
            asked: [].concat((a && a.asked) || [], (b && b.asked) || []),
            output: (a && a.output) || (b && b.output) || null,
            sitesAndLocations: a, employees: b
          };
        }
      });
      IMDebug.wire('lm-debug', 'legacy-migrate');
    }
    setView('lm');
    refresh();
  }

  // After any engine renders: buttons, the empty state, the switch, counts.
  function refresh() {
    if (!apis.lm || !apis.em) return;
    const anyReady = apis.lm.ready() || apis.em.ready();
    const anyBuilt = apis.lm.built() || apis.em.built();
    const run = $('lm-run'), exp = $('lm-export'), empty = $('lm-empty');
    if (run) run.disabled = !anyReady;
    if (exp) exp.disabled = !((apis.lm.built() && apis.lm.rows()) || (apis.em.built() && apis.em.rows()));
    if (empty) empty.style.display = anyBuilt ? 'none' : '';
    // Only one engine built: show its grid.
    if (view === 'em' && !apis.em.built() && apis.lm.built()) setView('lm');
    if (view === 'lm' && !apis.lm.built() && apis.em.built()) setView('employees');
    paintPills();
    if (wb) wb.refresh();
    if (window.IMDebug) IMDebug.refresh('legacy-migrate');
  }
  function open(key) { if (wb) wb.open(key); }

  // ─── Export all ───
  // Every built engine's files, numbered in upload order (Sites, Locations,
  // Employees), plus — when anything is still open — "0 - Read me first.html"
  // for the client. Nothing blocks: one heads-up, then the download.
  async function exportAll() {
    let limit = parseInt(($('lm-batch') || {}).value, 10);
    if (!Number.isFinite(limit) || limit < 1) limit = 5000;
    const built = ['lm', 'em'].filter(k => apis[k].built() && apis[k].rows());
    if (!built.length) { alert('Nothing to export yet.'); return; }
    const parts = [];
    for (const k of built) parts.push(await apis[k].exportParts(limit));
    const questions = [].concat(...built.map(k => apis[k].questions ? apis[k].questions() : []));
    const team = [].concat(...parts.map(p => p.team || []));
    const files = [].concat(...parts.map(p => p.files));
    if (questions.length || team.length) {
      const msg = (questions.length
        ? questions.length + ' question' + (questions.length === 1 ? '' : 's') + ' for the client — in “0 - Send to the client.html”:\n' +
          questions.map((q, i) => '  ' + (i + 1) + '. ' + q.title).join('\n') +
          '\n\nThey answer in that page and send back the file it saves. Drop that file on “Upload all files” here to apply their answers.'
        : '') +
        (team.length ? '\n\nBefore you upload (your side):\n' + team.map(t => '  • ' + t).join('\n') : '') +
        '\n\nExport now?';
      if (!window.confirm(msg)) return;
    }
    const date = today();
    const out = files.map((f, i) => ({ name: (i + 1) + ' - ' + f.label + ' - ' + f.name, data: f.data }));
    if (questions.length) out.unshift({ name: '0 - Send to the client.html', data: new TextEncoder().encode(IMLegacyForm.html({
      meta: {
        id: date + '-' + Math.random().toString(36).slice(2, 10),
        created: new Date().toISOString(),
        title: 'A few questions to finish your move to PickTrace 3.0',
        intro: 'We’ve moved your sites, blocks and employees over from Legacy. There ' + (questions.length === 1 ? 'is 1 thing' : 'are ' + questions.length + ' things') +
          ' only you can tell us — it takes a few minutes. Answer one at a time, then click “Save my answers” and send us the file it downloads.',
        answerName: 'PickTrace answers - ' + date + '.html'
      },
      questions
    })) });
    if (out.length === 1) { IMLegacyDrop.save(new Blob([out[0].data]), files[0].name); return; }
    const folder = 'Legacy to 3.0 - ' + date;
    IMLegacyDrop.save(IMXlsxTemplate.bundle(out.map(f => ({ name: folder + '/' + f.name, data: f.data }))), folder + '.zip');
    if (apis.lm.log) apis.lm.log('export', 'Exported ' + files.map(f => f.label).join(', ') + (questions.length ? ' and ' + questions.length + ' questions for the client' : '') + '.');
  }
  function today() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  // ─── The client's answers, back in ───
  // A page from "Send to the client", saved with their answers. Each engine
  // applies its own (keyed by Legacy IDs), so load the same Legacy files first.
  // Comments saved by the client page under answers.__notes: { questionId: { text, skip } }.
  function clientComments(r) {
    const notes = (r.answers && r.answers.__notes) || {};
    return r.questions.filter(q => notes[q.id] && String(notes[q.id].text || '').trim()).map(q => ({
      title: q.title, where: q.where || '', text: String(notes[q.id].text).trim(), skip: !!notes[q.id].skip
    }));
  }
  function showComments(list, fileName) {
    let box = $('lg-comments');
    if (!list.length) { if (box) box.remove(); return; }
    if (!box) {
      box = document.createElement('div');
      box.id = 'lg-comments';
      box.style.cssText = 'margin:12px 0;padding:14px 16px;border:1px solid #f0c36d;border-left:4px solid #f0a500;border-radius:8px;background:#fffaf0';
      const host = $('lg-drop');
      if (host) host.insertAdjacentElement('afterend', box); else document.body.appendChild(box);
    }
    const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const sk = list.filter(c => c.skip).length;
    box.innerHTML = '<div style="font-weight:600;margin:0 0 8px">Comments from the client (' + list.length + ')' + (sk ? ' — ' + sk + ' question' + (sk === 1 ? '' : 's') + ' skipped' : '') +
      ' <span style="font-weight:400;color:#6b7785">from ' + esc(fileName) + '</span></div>' +
      '<ul style="margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:8px">' + list.map(c =>
        '<li style="background:#fff;border:1px solid #eadfc4;border-radius:6px;padding:8px 12px">' +
        (c.skip ? '<span style="display:inline-block;font-size:11px;font-weight:700;color:#fff;background:#f0a500;border-radius:999px;padding:1px 8px;margin-right:6px">Skipped</span>' : '') +
        '<b>' + esc(c.title) + '</b>' + (c.where ? ' <span style="color:#6b7785;font-size:12px">' + esc(c.where) + '</span>' : '') +
        '<div style="margin-top:4px;font-style:italic">“' + esc(c.text) + '”</div></li>').join('') + '</ul>';
  }
  async function applyAnswersFile(file) {
    const text = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result)); fr.onerror = () => rej(fr.error); fr.readAsText(file); });
    const r = window.IMLegacyForm ? IMLegacyForm.read(text) : null;
    if (!r) return file.name + ' — not a PickTrace answers page';
    if (!Object.keys(r.answers).length) return file.name + ' — no answers in it yet';
    const needs = ['lm', 'em'].filter(k => Object.keys(r.answers).some(id => id.indexOf(k + '.') === 0));
    const missing = needs.filter(k => !apis[k].built());
    if (missing.length) {
      alert('These answers are for ' + missing.map(k => k === 'lm' ? 'sites & locations' : 'employees').join(' and ') +
        ', but those files aren’t loaded. Load the same Legacy files and templates first, then drop the answers file again.');
      return file.name + ' — load the Legacy files first';
    }
    // The client's comments (and the questions they skipped), shown under the upload box.
    const comments = clientComments(r);
    showComments(comments, file.name);
    const lines = [];
    needs.forEach(k => {
      const mine = {};
      Object.keys(r.answers).forEach(id => { if (id.indexOf(k + '.') === 0) mine[id] = r.answers[id]; });
      lines.push(...apis[k].applyAnswers(mine));
    });
    if (apis.lm.log) apis.lm.log('answers', 'Applied the client’s answers (' + file.name + '): ' + lines.join('; '));
    if (comments.length && apis.lm.log) apis.lm.log('answers', 'Client comments: ' + comments.map(c => (c.skip ? '[skipped] ' : '') + c.title + ': ' + c.text).join(' | '));
    refresh();
    alert('The client’s answers are in:\n\n' + (lines.length ? lines.map(l => '  • ' + l).join('\n') : '  (no changes)') +
      (comments.length ? '\n\nThey left ' + comments.length + ' comment' + (comments.length === 1 ? '' : 's') + (comments.some(c => c.skip) ? ', and skipped ' + comments.filter(c => c.skip).length + ' question' + (comments.filter(c => c.skip).length === 1 ? '' : 's') : '') + ':\n' +
        comments.map(c => '  • ' + (c.skip ? '[Skipped] ' : '') + c.title + ' — “' + c.text + '”').join('\n') + '\n\nThey’re listed under the upload box too.' : '') +
      '\n\nCheck the preview, then Export all again for clean files.');
    return file.name + ' — client’s answers applied (' + lines.length + ' change' + (lines.length === 1 ? '' : 's') + ')';
  }

  return { register, refresh, open, setView };
})();
