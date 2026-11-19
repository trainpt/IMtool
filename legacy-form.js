// ═══════════════════════════════════════════════════════════════════════
// The client's answer page for Legacy → 3.0: one self-contained HTML file.
//
// The client opens it, answers one question at a time (Back / Next, a
// progress bar), and clicks "Save my answers": that downloads a copy of the
// same page with their answers inside. They send it back; you open it to see
// what they chose, and drop it on "Upload all files" — IMLegacyForm.read()
// pulls the answers out and each engine applies them.
//
//   IMLegacyForm.html({ meta: { id, title, org, intro, answerName }, questions: [...] }) → string
//   IMLegacyForm.read(htmlText) → { meta, questions, answers, saved } | null
//
// Question types (each { id, type, title, help, where, ... }):
//   pick    — items [{ key, label, sub, value, group }], options, other (may type a new
//             value), placeholder, extra [{ value, label }]          → { key: value }
//   choice  — choices [{ value, label, note }], examples { columns, rows } → { choice }
//   choices — several choice questions on one page: parts [{ key, title, help, column,
//             choices [{ value, label, note, cells [per example row] }] }], examples;
//             each part adds a column to the example table showing what its pick does
//                                                                  → { partKey: value }
//   text    — placeholder                                          → 'string'
//   assign  — items [{ key, label, group, value }], options, quick [the few options shown
//             as buttons], groupLabel — select rows (or whole groups), press a type
//                                                                  → { key: value }
//   places  — items [{ key, label, values }], choices [{ key, label, options }] (e.g. Site
//             type), address: true — one address for all, or one each
//                                                                  → { key: { choiceKey, address1, city, … } }
//   cards   — items [{ key, label, sub, values }], fields [{ key, label, type, options,
//             required, section, showIf }], address: true, sections [{ key, label, copy }]
//                                                                  → { key: { field: value } }
// Address boxes: a paste box that splits a whole address into Street, Suite,
// City, State, Zip and Country — all shown as their own fields. Nothing is
// looked up or sent anywhere; the split happens in the page.
// ═══════════════════════════════════════════════════════════════════════
window.IMLegacyForm = (function () {
  'use strict';

  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  // JSON inside <script>: "</script>" or "<!--" in the data must not end it.
  const safeJson = v => JSON.stringify(v).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');

  // PickTrace 3.0 look: navy sidebar, white top bar, grey page, blue actions,
  // amber marker on the current question. One centred column; the bottom bar
  // tracks every item on the page so nothing is missed.
  const CSS = [
    ':root{color-scheme:light;--navy:#0e2a47;--ink:#1c2b3a;--ink2:#3a4654;--muted:#6b7785;--line:#dde3ec;--soft:#eef2f7;--bg:#f4f6f9;',
    '--accent:#1a6fba;--accent2:#2887d4;--accent-bg:#e9f2fb;--amber:#f0a500;--ok:#1e8449;--okc:#27ae60;--ok-bg:#e6f5ec;--warn:#9a5b00;--warn-bg:#fff4dc;--bad:#c0392b;--bad-bg:#fdecea;',
    '--side:248px;--top:56px;--col:1120px}',
    '*{box-sizing:border-box}html,body{margin:0}body{background:var(--bg);color:var(--ink2);font:15px/1.5 "Segoe UI",-apple-system,Roboto,"Open Sans",Arial,sans-serif;-webkit-font-smoothing:antialiased}',
    ':focus-visible{outline:2px solid var(--accent);outline-offset:2px}@media (prefers-reduced-motion:reduce){*{transition:none!important}}',
    // sidebar
    '.side{position:fixed;top:0;left:0;bottom:0;width:var(--side);background:var(--navy);color:#fff;display:flex;flex-direction:column;z-index:20}',
    '.side .logo{height:var(--top);display:flex;align-items:center;padding:0 20px;border-bottom:1px solid rgba(255,255,255,.08)}.side .logo img{height:22px;display:block}',
    '.side-l{padding:18px 20px 8px;font-size:12px;font-weight:600;color:rgba(255,255,255,.5)}',
    '.side nav{flex:1;overflow-y:auto;padding:0 0 12px;scrollbar-width:thin;scrollbar-color:rgba(255,255,255,.18) transparent}',
    '.sn{position:relative;display:flex;gap:11px;align-items:flex-start;width:100%;text-align:left;border:0;border-radius:0;background:none;color:rgba(255,255,255,.8);padding:9px 16px 9px 20px;font-size:13.5px;line-height:1.35}',
    '.sn:hover{background:rgba(255,255,255,.06);color:#fff}.sn.on{background:rgba(255,255,255,.12);color:#fff;font-weight:600}.sn.on::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--amber)}',
    '.sn .num{flex:none;width:22px;height:22px;border-radius:50%;display:grid;place-items:center;font-size:11px;font-weight:700;border:1.5px solid rgba(255,255,255,.35);margin-top:-1px}',
    '.sn .num.ok{background:var(--okc);border-color:var(--okc);color:#fff}.sn small{display:block;font-size:12px;font-weight:400;color:rgba(255,255,255,.5);margin-top:1px}',
    '.side-f{padding:14px 20px;border-top:1px solid rgba(255,255,255,.08);font-size:12px;color:rgba(255,255,255,.55);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    // top bar
    '.tbar{position:sticky;top:0;z-index:15;height:var(--top);margin-left:var(--side);background:#fff;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:16px;padding:0 28px}',
    '.tt{flex:1;min-width:0;font-size:15px;font-weight:600;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.tt span{font-weight:400;color:var(--muted);margin-left:8px}',
    '.count{font-size:13px;color:var(--muted);white-space:nowrap;font-variant-numeric:tabular-nums}',
    '.bar{height:6px;width:140px;background:var(--soft);border-radius:99px;overflow:hidden}.bar>i{display:block;height:100%;background:var(--okc);border-radius:99px;transition:width .25s}',
    // page column
    '.main{margin-left:var(--side);padding:32px 32px 112px}.inner{max-width:var(--col);margin:0 auto}',
    '.qh{margin:0 0 20px}.qn{font-size:13px;font-weight:600;color:var(--accent);margin:0 0 4px}',
    'h1{font-size:28px;line-height:1.2;margin:0 0 8px;color:var(--navy);font-weight:600;letter-spacing:-.01em}h2{font-size:20px;line-height:1.3;margin:0 0 8px;color:var(--navy);font-weight:600}h3{font-size:15px;margin:0;color:var(--ink);font-weight:600}',
    '.help{color:var(--muted);margin:0;max-width:75ch}.qh .help{font-size:15.5px}',
    // how-to strip: a real sequence, so numbered
    '.how{list-style:none;display:grid;grid-template-columns:repeat(3,1fr);gap:0;margin:18px 0 0;padding:0;background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden}',
    '.how li{display:flex;gap:10px;align-items:flex-start;padding:12px 16px;font-size:13.5px;color:var(--ink2);border-left:1px solid var(--line)}.how li:first-child{border-left:0}',
    '.how .n{flex:none;width:22px;height:22px;border-radius:50%;background:var(--accent);color:#fff;display:grid;place-items:center;font-size:12px;font-weight:700}',
    '.panel{background:#fff;border:1px solid var(--line);border-radius:8px;padding:20px 22px}.panel+.panel{margin-top:14px}',
    // controls
    'select,input[type=text],input[type=search],input[type=date],input[type=tel],input[type=email],input[type=number],textarea{width:100%;font:inherit;font-size:14px;height:38px;padding:6px 11px;border:1px solid #c9d2de;border-radius:6px;background:#fff;color:var(--ink);transition:border-color .15s,box-shadow .15s}',
    'textarea{height:auto;min-height:110px;resize:vertical}select:focus,input:focus,textarea:focus{outline:0;border-color:var(--accent);box-shadow:0 0 0 3px rgba(26,111,186,.16)}',
    'input::placeholder{color:#9aa6b4}',
    'button{font:inherit;font-size:14px;border-radius:6px;padding:8px 15px;border:1px solid #c9d2de;background:#fff;color:var(--ink);cursor:pointer;transition:background .15s,border-color .15s,color .15s}',
    'button:hover{border-color:var(--accent);color:var(--accent)}button.primary{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}button.primary:hover{background:var(--accent2);border-color:var(--accent2);color:#fff}',
    'button:disabled{opacity:.45;cursor:default}button.link{border:0;background:none;color:var(--accent);padding:2px 0;font-weight:600}button.small{font-size:13px;padding:5px 11px}',
    'button.danger{color:var(--bad);border-color:#efb7af}button.danger:hover{background:var(--bad-bg);color:var(--bad);border-color:var(--bad)}button.big{font-size:16px;padding:11px 22px}',
    '.seg{display:flex;flex-wrap:wrap;gap:6px;align-items:center}.seg button{border-radius:999px;padding:6px 16px}.seg button.on{background:var(--accent);border-color:var(--accent);color:#fff}',
    '.seg select{height:34px;border-radius:999px;width:auto!important;flex:none;max-width:160px}',
    // pick: tiles
    '.items{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:12px}',
    '.grp{background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden;min-width:0}.grp.all{grid-column:1/-1}',
    '.grp-h{padding:8px 14px;font-size:13px;font-weight:600;color:var(--ink2);background:var(--soft);border-bottom:1px solid var(--line)}',
    '.grp-b{display:grid;grid-template-columns:1fr}.grp.all .grp-b{grid-template-columns:repeat(auto-fill,minmax(340px,1fr))}',
    '.item{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.1fr);gap:12px;align-items:center;padding:9px 14px;border-top:1px solid var(--soft);min-width:0;scroll-margin:90px}.grp-b>.item:first-child{border-top:0}',
    '.grp.all .item{border-top:0;border-bottom:1px solid var(--soft)}.item.done{box-shadow:inset 3px 0 0 var(--okc)}',
    '.item .lab{min-width:0}.item .lab b{display:block;font-weight:600;font-size:14px;color:var(--ink);overflow-wrap:anywhere}.item .lab span{display:block;font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    // pick, detail layout
    '.dg{background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden}.dg+.dg{margin-top:16px}',
    '.dg-h{display:flex;flex-wrap:wrap;gap:12px 24px;align-items:center;justify-content:space-between;padding:12px 22px;background:var(--soft);border-bottom:1px solid var(--line)}',
    '.dg-h h3{font-size:14px;color:var(--ink2)}.dg-all{display:flex;gap:10px;align-items:center;font-size:13px;font-weight:600;color:var(--ink2)}.dg-all select,.dg-all input{width:280px}',
    '.drow{display:grid;grid-template-columns:minmax(0,1fr) 320px;gap:28px;align-items:start;padding:18px 22px;border-top:1px solid var(--soft);scroll-margin:110px}.dg-h+.drow,.dg>.drow:first-child{border-top:0}',
    '.drow.done{box-shadow:inset 4px 0 0 var(--okc)}.dt{display:block;font-size:16px;font-weight:600;color:var(--navy)}.dsub{margin:2px 0 0;font-size:13px;color:var(--muted)}',
    '.facts{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px 24px;margin:10px 0 0}.facts div{min-width:0}.facts dt{font-size:12px;color:var(--muted)}.facts dd{margin:1px 0 0;font-size:14px;color:var(--ink);overflow-wrap:anywhere}',
    '.facts div:has(.chips){grid-column:1/-1}.chips{display:flex;flex-wrap:wrap;gap:5px;margin-top:3px}.chip{font-size:12.5px;padding:2px 9px;border-radius:999px;background:var(--accent-bg);color:var(--accent)}',
    '.dans{display:flex;flex-direction:column;gap:6px;position:sticky;top:calc(var(--top) + 16px)}.dans>span{font-size:13px;font-weight:600;color:var(--ink)}',
    '@media (max-width:900px){.drow{grid-template-columns:1fr;gap:14px}.dg-all select,.dg-all input{width:100%}}',
    // choice
    '.choices{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:10px}.choice{display:flex;gap:10px;align-items:flex-start;padding:13px 15px;border:1px solid var(--line);border-radius:8px;cursor:pointer;background:#fff}',
    '.choice:hover{border-color:var(--accent)}.choice input{margin-top:3px;width:16px;height:16px;flex:none;accent-color:var(--accent)}.choice.on{border-color:var(--accent);background:var(--accent-bg);box-shadow:inset 0 0 0 1px var(--accent)}.choice .note{font-size:13px;color:var(--muted);margin-top:2px}',
    '.cpart{margin:26px 0 0;scroll-margin:90px}.cpart h3{display:flex;gap:10px;align-items:center;margin:0 0 4px;font-size:17px;color:var(--navy)}.cpart .num{display:grid;place-items:center;width:26px;height:26px;border-radius:50%;background:var(--accent);color:#fff;font-size:13px;flex:none}.cpart .help{margin:0 0 12px 36px}',
    '.ex{background:#fff;border:1px solid var(--line);border-radius:8px;overflow-x:auto}.ex table,.ptable{border-collapse:collapse;width:100%;font-size:14px}',
    '.ex th,.ptable th{padding:10px 12px;text-align:left;font-size:11px;font-weight:600;letter-spacing:.5px;text-transform:uppercase;color:var(--muted);border-bottom:2px solid var(--accent);white-space:nowrap;background:#fff}',
    '.ex td{padding:10px 12px;border-bottom:1px solid var(--soft)}.ex tr:last-child td{border-bottom:0}.ex tbody tr:nth-child(even) td{background:#f8fafc}.ex code{background:var(--warn-bg);color:var(--warn);padding:1px 5px;border-radius:4px}',
    '.ex th.res{color:var(--accent)}.ex td.res{background:var(--accent-bg)!important;color:var(--ink)}.ex .wait{color:var(--muted);font-style:italic}',
    // assign
    '.abar{position:sticky;top:calc(var(--top) + 10px);z-index:5;display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:12px 14px;margin:0 0 12px;background:#fff;border:1px solid var(--line);border-radius:8px;box-shadow:0 6px 18px rgba(14,42,71,.08)}',
    '.abar .sel{font-weight:700;color:var(--accent);min-width:100px}.abar .seg{flex:1}.abar>span:nth-child(2){font-size:13px;color:var(--muted)}',
    '.atools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 12px}.atools input{flex:1 1 260px;width:auto;border-radius:999px;padding-left:16px}',
    // Columns fill top-to-bottom (like a newspaper), so the order on screen is
    // the order Shift+click ranges use. An open site spans the full width.
    '.aglist{display:flex;flex-direction:column;gap:6px}',
    '.ag{border:1px solid var(--line);border-radius:8px;overflow:hidden;background:#fff;min-width:0}.ag-rm{margin-top:8px}',
    '.ag-h{display:flex;align-items:center;gap:10px;padding:10px 14px;cursor:pointer;user-select:none}.ag-h:hover{background:#f8fafc}.ag.open>.ag-h{background:var(--soft);border-bottom:1px solid var(--line)}',
    '.ag-h .nm{font-weight:600;color:var(--ink);flex:1;min-width:0}.ag-h .sum{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end}.ag-h .tw{color:var(--muted);width:14px;text-align:center}',
    '.tp{font-size:12px;padding:1px 8px;border-radius:999px;background:var(--soft);color:var(--ink2);white-space:nowrap}.tp.chg{background:var(--warn-bg);color:var(--warn)}.tp.rm{background:var(--bad-bg);color:var(--bad)}',
    '.ck{width:18px;height:18px;border:2px solid #9aa8b8;border-radius:4px;display:inline-grid;place-items:center;flex:none;font-size:11px;color:#fff;background:#fff}',
    '.ck.on{background:var(--accent);border-color:var(--accent)}.ck.part{background:#8dbbe4;border-color:#8dbbe4}',
    '.ar-g{display:flex;flex-direction:column;padding:6px 8px 8px 30px}.ar-g .ar{border-bottom:1px solid var(--soft);border-radius:0}.ar-g .ar:last-child{border-bottom:0}',
    '.ar{display:flex;align-items:flex-start;gap:9px;padding:6px 8px;border-radius:6px;cursor:pointer;user-select:none;font-size:13.5px;min-width:0}.ar:hover{background:#f3f6fa}.ar.on{background:var(--accent-bg)}',
    '.ar .ck{margin-top:1px}.ar .nm{flex:1;min-width:0;overflow-wrap:anywhere;color:var(--ink)}.pill{font-size:12px;padding:1px 9px;border-radius:999px;background:var(--soft);color:var(--ink2);white-space:nowrap}.pill.chg{background:var(--warn-bg);color:var(--warn)}.pill.rm{background:var(--bad-bg);color:var(--bad)}',
    '.arr{display:flex;align-items:center;gap:10px;padding:6px 14px;border-top:1px solid var(--soft);color:var(--muted);font-size:13.5px}.arr .nm{flex:1;min-width:0;text-decoration:line-through}.note-s{color:var(--muted);font-size:12px}.ag-rm>.ag-h{background:var(--bad-bg)}',
    // address
    '.lbl{font-size:13px;font-weight:600;color:var(--ink);margin:0 0 6px}',
    '.addr-in{border-style:dashed!important;background:#f8fafc!important}.addr-in:focus{background:#fff!important}',
    '.parts{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:12px;margin-top:12px}.parts [data-pt]{grid-column:span 3}',
    '.parts [data-pt=address1]{grid-column:span 8}.parts [data-pt=address2]{grid-column:span 4}.parts [data-pt=city]{grid-column:span 5}.parts [data-pt=state],.parts [data-pt=country]{grid-column:span 2}',
    // places
    '.mode{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:0 0 14px}.mode button{text-align:left;padding:14px 16px;border:1px solid var(--line);border-radius:8px;background:#fff}',
    '.mode button b{display:block;font-size:15px;color:var(--ink)}.mode button span{font-size:13px;color:var(--muted)}.mode button.on{border-color:var(--accent);background:var(--accent-bg);box-shadow:inset 0 0 0 1px var(--accent)}',
    '.blk{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px 20px;margin:0 0 14px}.blk>h3{margin:0 0 4px}.blk>.help{margin:0 0 12px}',
    '.twrap{margin:0 -20px -16px}.ptable{table-layout:fixed}.ptable thead th{position:sticky;top:var(--top);z-index:2}',
    '.ptable td{padding:4px 3px;border-bottom:1px solid var(--soft);vertical-align:middle}.ptable tr:last-child td{border-bottom:0}.ptable td:first-child,.ptable th:first-child{padding-left:20px}.ptable td:last-child,.ptable th:last-child{padding-right:20px}',
    '.ptable tr{scroll-margin:120px}.ptable td.nm{font-weight:600;font-size:13.5px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.ptable input,.ptable select{height:32px;font-size:13px;padding:3px 8px}',
    '.ptable td.miss input{border-color:#efb7af;background:#fffafa}.ptable tr.full td.nm{box-shadow:inset 3px 0 0 var(--okc)}',
    '.ptable .w-nm{width:25%}.ptable .w-su{width:9%}.ptable .w-ci{width:13%}.ptable .w-st{width:6%}.ptable .w-zip{width:8%}.ptable .w-co{width:7%}.ptable .w-ty{width:13%}',
    // cards (employers): tabs, then one form laid out like a settings page
    '.tabs{display:flex;flex-wrap:wrap;gap:0 26px;margin:0 0 16px;border-bottom:1px solid var(--line)}.tabs button{border:0;border-bottom:3px solid transparent;border-radius:0;margin-bottom:-1px;padding:9px 0;background:none;color:var(--muted);font-size:15px}',
    '.tabs button:hover{color:var(--accent)}.tabs button.on{border-bottom-color:var(--accent);color:var(--accent);font-weight:600}.tstat{margin-left:8px;font-size:11px;font-weight:700;padding:1px 7px;border-radius:999px}',
    '.tstat.left{background:var(--bad-bg);color:var(--bad)}.tstat.done{background:var(--ok-bg);color:var(--ok)}.tab-nav{display:flex;justify-content:space-between;margin-top:14px}',
    '.ecard{background:#fff;border:1px solid var(--line);border-radius:8px}',
    '.ecard-h{display:flex;gap:14px;align-items:center;padding:18px 24px;border-bottom:1px solid var(--line)}.ecard-h .t{flex:1;min-width:0}.ecard-h b{display:block;font-size:20px;font-weight:600;color:var(--navy)}.ecard-h .t span{font-size:13px;color:var(--muted)}',
    '.status{font-size:12px;font-weight:700;padding:4px 11px;border-radius:999px;white-space:nowrap}.status.left{background:var(--bad-bg);color:var(--bad)}.status.done{background:var(--ok-bg);color:var(--ok)}',
    '.sec{display:grid;grid-template-columns:210px minmax(0,1fr);gap:28px;padding:20px 24px;border-top:1px solid var(--soft)}.ecard-h+.sec{border-top:0}',
    '.sec-h h3{font-size:15px;color:var(--ink)}.sec-h p{margin:3px 0 0;font-size:13px;color:var(--muted)}',
    '.fgrid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:14px 14px;grid-auto-flow:row dense}.sec-b>.addr+.fgrid{margin-top:14px}',
    '.fld{display:block;min-width:0;grid-column:span 6}.fld.full{grid-column:1/-1}.fld .lt{display:flex;align-items:baseline;gap:8px;font-size:13px;color:var(--ink);margin:0 0 5px;font-weight:600}',
    '.fld .need,.lbl .need{color:var(--bad);font-weight:600;font-size:11.5px}.fld .need.ok,.lbl .need.ok{color:var(--ok)}.fld .opt{color:var(--muted);font-weight:400;font-size:11.5px}',
    '.fld.miss input,.fld.miss select{border-color:#efb7af;background:#fffafa}',
    // intro / finish
    '.list{list-style:none;margin:0;padding:0;background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden}.list li{display:flex;gap:14px;align-items:center;padding:13px 18px;border-top:1px solid var(--soft)}.list li:first-child{border-top:0}',
    '.dot{flex:none;width:26px;height:26px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:700;color:var(--muted);border:1.5px solid #c9d2de;background:#fff}.dot.ok{background:var(--okc);border-color:var(--okc);color:#fff}',
    '.list .t{flex:1;min-width:0}.list .t b{display:block;font-weight:600;color:var(--ink)}.list .t span{font-size:13px;color:var(--muted)}',
    '.note-box{margin:16px 0 0;padding:12px 16px;border-left:3px solid var(--accent);border-radius:0 8px 8px 0;background:var(--accent-bg);font-size:14px;color:var(--ink);max-width:75ch}.saved{font-size:13px;color:var(--muted)}',
    // comment / skip box
    '.cmt{margin:20px 0 0;background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px 20px;scroll-margin:120px}.cmt.closed{background:none;border:0;padding:0}',
    '.cmt.warn{border-color:#f0c36d;background:#fffaf0;box-shadow:inset 4px 0 0 var(--amber)}.cmt.sk{border-color:#f0c36d;background:#fffaf0}',
    '.cmt-h{font-size:14px;color:var(--ink);margin:0 0 12px;display:flex;flex-wrap:wrap;gap:8px 12px;align-items:center}.cmt-l{display:block;font-size:13px;font-weight:600;color:var(--ink);margin:0 0 6px}',
    '.cmt textarea{min-height:76px}.cmt-a{display:flex;gap:12px;align-items:center;margin-top:10px}.cmt-n{font-size:12.5px;color:var(--muted);margin:6px 0 0}',
    '.sn.lock{opacity:.45;cursor:not-allowed}.sn .num.sk,.dot.sk{background:var(--amber);border-color:var(--amber);color:#fff}.lk{font-size:12px;color:var(--muted)}',
    '.lq{display:block;margin-top:4px;font-size:13px;color:var(--ink2);font-style:italic}',
    // bottom bar: back · page tracker · next
    '.foot{position:fixed;left:var(--side);right:0;bottom:0;z-index:18;background:#fff;border-top:1px solid var(--line);box-shadow:0 -6px 18px rgba(14,42,71,.06)}',
    '.fin{max-width:var(--col);margin:0 auto;padding:12px 32px;display:flex;gap:18px;align-items:center;box-sizing:content-box}.fin .sp{flex:1}',
    '.trk{flex:1;min-width:0;display:flex;flex-direction:column;gap:6px}.trk-t{display:flex;gap:12px;align-items:baseline;font-size:13px;color:var(--ink2)}.trk-t b{color:var(--ink)}',
    '.trk-t button{margin-left:auto}.tk{display:flex;gap:3px;height:8px}.tki{flex:1;min-width:3px;max-width:48px;border:0;padding:0;border-radius:2px;background:#d6dde7;cursor:pointer}',
    '.tki:hover{background:#aebbcc}.tki.ok{background:var(--okc)}.tki.ok:hover{background:#1f9a54}.trk-t .all{color:var(--ok);font-weight:600}',
    '@media (max-width:1100px){.how{grid-template-columns:1fr}.how li{border-left:0;border-top:1px solid var(--line)}.how li:first-child{border-top:0}.sec{grid-template-columns:1fr;gap:12px}}',
    '@media (max-width:900px){:root{--side:0px}.side{display:none}.tbar{padding:0 14px}.bar{width:70px}.main{padding:20px 14px 150px}.fin{padding:10px 14px;flex-wrap:wrap}.trk{order:-1;flex-basis:100%}',
    '.twrap{overflow-x:auto}.ptable thead th{position:static}.item,.mode{grid-template-columns:1fr}.fld,.parts [data-pt]{grid-column:1/-1!important}h1{font-size:23px}.tt span{display:none}.ecard-h,.sec{padding-left:16px;padding-right:16px}}',
    '@media print{.foot,.side,.tbar{display:none}.main{margin:0}}'
  ].join('');
  // PickTrace wordmark (white), for the navy sidebar.
  const LOGO = 'data:image/png;base64,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';

  // Runs inside the client's page. Plain ES2017, no libraries.
  function app() {
    const spec = JSON.parse(document.getElementById('imt-spec').textContent);
    const store = document.getElementById('imt-answers');
    const Q = spec.questions;
    let saved = { saved: 0, answers: {} };
    try { saved = JSON.parse(store.textContent || '{}') || saved; } catch (e) { /* none yet */ }
    let ans = saved.answers || {};
    const KEY = 'imt-answers-' + spec.meta.id;
    try {
      const local = JSON.parse(localStorage.getItem(KEY) || 'null');
      if (local && local.saved > (saved.saved || 0)) ans = local.answers || ans;
    } catch (e) { /* storage blocked — the Save button still works */ }
    // Start from what the tool already knows.
    Q.forEach(q => {
      const a = ans[q.id] || (ans[q.id] = (q.type === 'choice' || q.type === 'text') ? ans[q.id] : {});
      if (q.type === 'pick' || q.type === 'assign') q.items.forEach(it => { if (a[it.key] == null && it.value) a[it.key] = it.value; });
      if (q.type === 'places' || q.type === 'cards') q.items.forEach(it => { a[it.key] = Object.assign({}, it.values || {}, a[it.key] || {}); });
      if (q.type === 'choice' && !ans[q.id] && q.value) ans[q.id] = { choice: q.value };
    });
    let step = 0;          // 0 intro · 1..n questions · n+1 finish
    let dirty = false;
    const n = Q.length;
    const $app = document.getElementById('app');
    const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const $ = (sel, root) => (root || document).querySelector(sel);
    const $$ = (sel, root) => Array.prototype.slice.call((root || document).querySelectorAll(sel));
    const cssq = v => (window.CSS && CSS.escape) ? CSS.escape(v) : String(v).replace(/(["\\])/g, '\\$1');
    const cur = () => Q[step - 1];

    const short = t => String(t || '').length > 58 ? String(t).slice(0, 56).trim() + '…' : String(t || '');
    function paintChrome() {
      paintTracker();
      const q0 = cur();
      if (blocked && q0 && done(q0)) { blocked = false; paintCmt(); }
      const dc = doneCount(), nav = document.getElementById('imt-nav');
      if (nav) nav.innerHTML = '<button class="sn' + (step === 0 ? ' on' : '') + '" data-go="0"><span class="num">i</span><span>Start here</span></button>' +
        Q.map((q, i) => { const lock = i + 1 > reach(), sk = skipped(q);
          return '<button class="sn' + (step === i + 1 ? ' on' : '') + (lock ? ' lock' : '') + '" data-go="' + (i + 1) + '"' + (lock ? ' title="Finish or skip the questions before this one first"' : '') + '>' +
            '<span class="num' + (done(q) ? ' ok' : sk ? ' sk' : '') + '">' + (done(q) ? '✓' : sk ? '!' : i + 1) + '</span>' +
            '<span>' + esc(short(q.title)) + '<small>' + (sk ? 'Skipped — comment sent' : esc(String(q.where || '').split(' · ').pop())) + '</small></span></button>'; }).join('') +
        '<button class="sn' + (step > n ? ' on' : '') + (reach() <= n ? ' lock' : '') + '" data-go="' + (n + 1) + '"><span class="num">⬇</span><span>Save &amp; send</span></button>';
      const sc = Q.filter(skipped).length, t = dc + ' of ' + n + ' answered' + (sc ? ', ' + sc + ' skipped' : '');
      const c = document.getElementById('imt-count'), b = document.getElementById('imt-bar'), ft = document.getElementById('imt-foot');
      if (c) c.textContent = t;
      if (b) b.style.width = Math.round(dc / Math.max(1, n) * 100) + '%';
      if (ft) ft.textContent = '✓ Your answers save as you go';
    }
    function keep() {
      dirty = true;
      try { localStorage.setItem(KEY, JSON.stringify({ saved: Date.now(), answers: ans })); } catch (e) { /* fine */ }
      const s = $('.saved');
      if (s) s.textContent = 'Saved on this computer — click “Save my answers” to send them.';
      paintChrome();
    }
    // ─── Addresses ───────────────────────────────────────────────────────
    const PARTS = ['address1', 'address2', 'city', 'state', 'zip', 'country'];
    const PART_LABEL = { address1: 'Street', address2: 'Suite / unit', city: 'City', state: 'State', zip: 'Zip', country: 'Country' };
    const STATES = { ALABAMA: 'AL', ALASKA: 'AK', ARIZONA: 'AZ', ARKANSAS: 'AR', CALIFORNIA: 'CA', COLORADO: 'CO', CONNECTICUT: 'CT', DELAWARE: 'DE',
      'DISTRICT OF COLUMBIA': 'DC', FLORIDA: 'FL', GEORGIA: 'GA', HAWAII: 'HI', IDAHO: 'ID', ILLINOIS: 'IL', INDIANA: 'IN', IOWA: 'IA', KANSAS: 'KS',
      KENTUCKY: 'KY', LOUISIANA: 'LA', MAINE: 'ME', MARYLAND: 'MD', MASSACHUSETTS: 'MA', MICHIGAN: 'MI', MINNESOTA: 'MN', MISSISSIPPI: 'MS',
      MISSOURI: 'MO', MONTANA: 'MT', NEBRASKA: 'NE', NEVADA: 'NV', 'NEW HAMPSHIRE': 'NH', 'NEW JERSEY': 'NJ', 'NEW MEXICO': 'NM', 'NEW YORK': 'NY',
      'NORTH CAROLINA': 'NC', 'NORTH DAKOTA': 'ND', OHIO: 'OH', OKLAHOMA: 'OK', OREGON: 'OR', PENNSYLVANIA: 'PA', 'RHODE ISLAND': 'RI',
      'SOUTH CAROLINA': 'SC', 'SOUTH DAKOTA': 'SD', TENNESSEE: 'TN', TEXAS: 'TX', UTAH: 'UT', VERMONT: 'VT', VIRGINIA: 'VA', WASHINGTON: 'WA',
      'WEST VIRGINIA': 'WV', WISCONSIN: 'WI', WYOMING: 'WY', 'PUERTO RICO': 'PR', ALBERTA: 'AB', 'BRITISH COLUMBIA': 'BC', MANITOBA: 'MB',
      'NEW BRUNSWICK': 'NB', 'NEWFOUNDLAND AND LABRADOR': 'NL', 'NOVA SCOTIA': 'NS', ONTARIO: 'ON', 'PRINCE EDWARD ISLAND': 'PE', QUEBEC: 'QC',
      SASKATCHEWAN: 'SK', 'NORTHWEST TERRITORIES': 'NT', NUNAVUT: 'NU', YUKON: 'YT' };
    const CODES = {};
    Object.keys(STATES).forEach(k => { CODES[STATES[k]] = 1; });
    const CA_PROV = { AB: 1, BC: 1, MB: 1, NB: 1, NL: 1, NS: 1, ON: 1, PE: 1, QC: 1, SK: 1, NT: 1, NU: 1, YT: 1 };
    const stateCode = s => { const u = String(s || '').trim().toUpperCase().replace(/\./g, ''); return CODES[u] ? u : (STATES[u] || ''); };
    const COUNTRY = { US: 'US', USA: 'US', 'UNITED STATES': 'US', 'UNITED STATES OF AMERICA': 'US', CANADA: 'CA', MEXICO: 'MX', 'MÉXICO': 'MX' };
    const SUFFIX = /^(ST|STREET|AVE|AV|AVENUE|RD|ROAD|DR|DRIVE|LN|LANE|BLVD|BOULEVARD|WAY|HWY|HIGHWAY|CT|COURT|PL|PLACE|PKWY|PARKWAY|CIR|CIRCLE|TRL|TRAIL|TER|TERRACE|RTE|ROUTE|LOOP|ROW|RUN|PIKE|SQ|EXPY|FWY|BOX)$/;
    const DIRS = /^(N|S|E|W|NE|NW|SE|SW|NORTH|SOUTH|EAST|WEST)$/;
    // "123 Main St, Suite 4, Anytown, IL 62701, USA" or the same with no commas.
    function parseAddress(raw) {
      const out = { address1: '', address2: '', city: '', state: '', zip: '', country: '' };
      const parts = String(raw || '').replace(/\s*[\r\n]+\s*/g, ', ').split(',').map(x => x.trim().replace(/\s+/g, ' ')).filter(Boolean);
      if (!parts.length) return out;
      const last = parts[parts.length - 1].toUpperCase().replace(/\./g, '');
      if (COUNTRY[last] && parts.length > 1) { out.country = COUNTRY[last]; parts.pop(); }
      // "… ST 12345" at the end (state may be spelled out).
      let tail = parts.pop();
      let m = tail.match(/^(.*?)[\s,]*\b([A-Za-z][A-Za-z .]*?)\.?\s+(\d{5}(?:-\d{4})?|[A-Za-z]\d[A-Za-z] ?\d[A-Za-z]\d)$/);
      if (m && !stateCode(m[2])) {
        // The state may be the last word only ("Anytown IL 62701" read "Anytown IL" as the state).
        const w = m[2].split(' ');
        if (stateCode(w[w.length - 1])) m = [m[0], (m[1] + ' ' + w.slice(0, -1).join(' ')).trim(), w[w.length - 1], m[3]];
        else m = null;
      }
      let rest = tail;
      if (m) { out.state = stateCode(m[2]); out.zip = m[3].toUpperCase(); rest = m[1]; }
      else if (stateCode(tail) && parts.length) { out.state = stateCode(tail); rest = ''; }
      else if (/^\d{5}(-\d{4})?$/.test(tail)) { out.zip = tail; rest = ''; }
      if (rest) parts.push(rest);
      if (parts.length >= 2) {
        out.city = parts.pop();
        out.address1 = parts.shift();
        out.address2 = parts.join(', ');
      } else if (parts.length === 1) {
        // No commas: the street ends at its last street word ("Ave", "Rd"…),
        // plus a direction or route number after it; the rest is the city.
        const toks = parts[0].split(' ');
        let cut = -1;
        toks.forEach((t, i) => { if (i > 0 && SUFFIX.test(t.toUpperCase().replace(/\./g, ''))) cut = i; });
        while (cut >= 0 && cut + 1 < toks.length - 1 && (DIRS.test(toks[cut + 1].toUpperCase()) || /^\d+[A-Za-z]?$/.test(toks[cut + 1]))) cut++;
        if (cut >= 0 && cut < toks.length - 1 && (out.state || out.zip)) { out.address1 = toks.slice(0, cut + 1).join(' '); out.city = toks.slice(cut + 1).join(' '); }
        else out.address1 = parts[0];
      }
      if (!out.country && out.state) out.country = CA_PROV[out.state] ? 'CA' : 'US';
      return out;
    }
    const complete = p => !!(p && p.address1 && p.city && p.state && p.zip);
    // Where an address box reads and writes: get(key) → parts, set(key, parts),
    // need(key) → the parts that are required (default: street, city, state, zip).
    const addrIO = {};
    const NEED_DEFAULT = ['address1', 'city', 'state', 'zip'];
    const partTag = (k, p, need) => need.indexOf(k) < 0 ? '<span class="opt">Optional</span>'
      : String(p[k] || '').trim() ? '<span class="need ok">✓ Done</span>' : '<span class="need">Required</span>';
    // A paste box on top, and every part as its own field — always on show.
    // Nothing is looked up online: the paste is split right here.
    function addrBox(scope, key, label) {
      const io = addrIO[scope], p = io.get(key) || {}, need = (io.need && io.need(key)) || NEED_DEFAULT;
      return '<div class="addr" data-scope="' + esc(scope) + '" data-key="' + esc(key) + '">' +
        (label ? '<div class="lbl">' + esc(label) + '</div>' : '') +
        '<input type="text" class="addr-in" placeholder="Paste the full address here — e.g. 123 Main St, Anytown, IL 62701 — it fills in the boxes below" autocomplete="off" value="">' +
        '<div class="parts">' + PARTS.map(k => '<label class="fld' + (need.indexOf(k) >= 0 && !String(p[k] || '').trim() ? ' miss' : '') + '" data-pt="' + k + '"><span class="lt"><span>' + PART_LABEL[k] + '</span>' + partTag(k, p, need) + '</span>' +
          '<input type="text" data-part="' + k + '" value="' + esc(p[k] || '') + '" autocomplete="off"></label>').join('') + '</div></div>';
    }
    function paintAddr(box) {
      const io = addrIO[box.dataset.scope], p = io.get(box.dataset.key) || {}, need = (io.need && io.need(box.dataset.key)) || NEED_DEFAULT;
      PARTS.forEach(k => {
        const l = $('[data-pt="' + k + '"]', box);
        if (!l) return;
        const i = $('input', l);
        if (document.activeElement !== i) i.value = p[k] || '';
        l.classList.toggle('miss', need.indexOf(k) >= 0 && !String(p[k] || '').trim());
        const lt = $('.lt', l);
        if (lt) lt.innerHTML = '<span>' + PART_LABEL[k] + '</span>' + partTag(k, p, need);
      });
      if (box.tagName === 'TR') box.classList.toggle('full', complete(p));
    }
    function setAddr(box, p) {
      const io = addrIO[box.dataset.scope];
      io.set(box.dataset.key, PARTS.reduce((o, k) => { o[k] = p[k] || ''; return o; }, {}));
      paintAddr(box);
      keep();
      if (io.after) io.after(box.dataset.key);
    }
    // A pasted (or typed + Enter) address: split it and fill the boxes. A
    // part the paste didn't have keeps what was there (e.g. Country).
    const clean = t => String(t || '').replace(/\s*[\r\n]+\s*/g, ', ').replace(/\s+/g, ' ').trim();
    function resolveText(box, text) {
      text = clean(text);
      if (!text) return;
      const io = addrIO[box.dataset.scope], old = io.get(box.dataset.key) || {}, p = parseAddress(text);
      if (!p.country && old.country) p.country = old.country;
      setAddr(box, p);
      // Repaint skips the box being typed in — the Street box often is.
      const si = $('[data-part="address1"]', box);
      if (si) si.value = p.address1 || '';
    }
    function resolve(input) {
      if (!input.value.trim()) return;
      resolveText(input.closest('.addr'), input.value);
      input.value = '';
    }

    // ─── Answered? ───────────────────────────────────────────────────────
    const shown = (f, v) => !f.showIf || (f.showIf.values || []).indexOf((v || {})[f.showIf.field] || '') >= 0;
    const cardLeft = (q, v) => q.fields.filter(f => f.required && shown(f, v) && !String((v || {})[f.key] || '').trim()).length;
    function done(q) {
      const a = ans[q.id];
      if (q.type === 'pick' || q.type === 'assign') return !!a && q.items.every(it => a[it.key] && String(a[it.key]).trim());
      if (q.type === 'places') return !!a && q.items.every(it => { const v = a[it.key] || {}; return complete(v) && (q.choices || []).every(c => v[c.key]); });
      if (q.type === 'cards') return !!a && q.items.every(it => !cardLeft(q, a[it.key]));
      if (q.type === 'choice') return !!(a && a.choice);
      if (q.type === 'choices') return !!a && q.parts.every(p => a[p.key]);
      return !!(a && String(a).trim());
    }
    const doneCount = () => Q.filter(done).length;
    // Comments for PickTrace, per question: { text, skip }. Saved with the answers
    // under "__notes" — the IM Tool shows them when the answers are loaded.
    const NOTES = '__notes';
    const noteOf = q => (ans[NOTES] || {})[q.id] || null;
    const setNote = (q, patch) => { const all = ans[NOTES] = ans[NOTES] || {}; all[q.id] = Object.assign({ text: '', skip: false }, all[q.id] || {}, patch); };
    const skipped = q => !done(q) && !!(noteOf(q) && noteOf(q).skip && String(noteOf(q).text || '').trim());
    const settled = q => done(q) || skipped(q);
    // The furthest page they may open: up to the first question not yet settled.
    const reach = () => { const i = Q.findIndex(q => !settled(q)); return i < 0 ? n + 1 : i + 1; };
    let blocked = false;     // pressed Next on an unfinished question
    let cmtOpen = false;     // opened the comment box

    function control(kind, val, opts, attrs, other, placeholder, extra) {
      const list = (extra || []).concat((opts || []).map(o => ({ value: o, label: o })));
      if (kind === 'date') return '<input type="date" value="' + esc(val) + '" ' + attrs + '>';
      if (kind === 'tel' || kind === 'email' || kind === 'number') return '<input type="' + kind + '" value="' + esc(val) + '"' + (kind === 'number' ? ' min="0" step="any"' : '') + (placeholder ? ' placeholder="' + esc(placeholder) + '"' : '') + ' ' + attrs + '>';
      if ((kind === 'select' || list.length) && !other) {
        return '<select ' + attrs + '><option value="">— choose —</option>' + list.map(o => '<option value="' + esc(o.value) + '"' +
          (o.value === val ? ' selected' : '') + '>' + esc(o.label) + '</option>').join('') + '</select>';
      }
      const id = 'dl-' + Math.random().toString(36).slice(2);
      return '<input type="text" list="' + id + '" value="' + esc(val) + '" placeholder="' + esc(placeholder || (list.length ? 'Choose or type…' : 'Type here…')) + '" ' + attrs + '>' +
        (list.length ? '<datalist id="' + id + '">' + list.map(o => '<option value="' + esc(o.value) + '">' + (o.label !== o.value ? esc(o.label) : '') + '</option>').join('') + '</datalist>' : '');
    }
    // A few options as buttons; the rest in a "More…" list.
    function seg(options, val, attrs, max) {
      const opts = options || [], few = opts.slice(0, max || 6), more = opts.slice(max || 6);
      const inMore = val && few.indexOf(val) < 0;
      return '<div class="seg" ' + attrs + '>' + few.map(o => '<button type="button" data-v="' + esc(o) + '"' + (o === val ? ' class="on"' : '') + '>' + esc(o) + '</button>').join('') +
        (more.length || inMore ? '<select data-more="1" style="width:auto"><option value="">More…</option>' + (inMore && more.indexOf(val) < 0 ? '<option selected>' + esc(val) + '</option>' : '') +
          more.map(o => '<option' + (o === val ? ' selected' : '') + '>' + esc(o) + '</option>').join('') + '</select>' : '') + '</div>';
    }

    // ─── assign: select rows or whole groups, press a type ───
    // "Remove" — leave the row out of 3.0. Stored as this value.
    const REMOVED = '__leave__';
    const asg = { sel: {}, open: {}, filter: '', last: null };
    // Rows still in play (not removed) that match the search, by group, in order.
    function assignGroups(q) {
      const a = ans[q.id] || {}, f = asg.filter.toLowerCase();
      const groups = [], byG = {};
      q.items.forEach(it => {
        if (a[it.key] === REMOVED) return;
        const g = it.group || '';
        if (f && String(it.label).toLowerCase().indexOf(f) < 0 && g.toLowerCase().indexOf(f) < 0) return;
        if (!byG[g]) { byG[g] = []; groups.push(g); }
        byG[g].push(it);
      });
      // "01 - …", "02 - …", "15 - …" in number order; "(no site yet)" last.
      const num = g => { const m = /^\s*(\d+)/.exec(g); return m ? +m[1] : Infinity; };
      groups.sort((x, y) => (x.charAt(0) === '(') - (y.charAt(0) === '(') || num(x) - num(y) || x.localeCompare(y));
      return groups.map(g => ({ g, its: byG[g] }));
    }
    // Is every block in the list (as searched) selected?
    function allPicked(q) { const its = [].concat.apply([], assignGroups(q).map(x => x.its)); return its.length > 0 && its.every(it => asg.sel[it.key]); }
    function assignBody(q) {
      const a = ans[q.id] || {};
      const count = Object.keys(asg.sel).length;
      const def = q.items[0] ? q.items[0].value : '';
      const removed = q.items.filter(it => a[it.key] === REMOVED);
      let html = '<div class="abar"><span class="sel">' + (count ? count + ' selected' : 'Select below') + '</span>' +
        '<span style="color:var(--muted);font-size:14px">Set to:</span>' + seg(q.quick || q.options, '', 'data-aset="1"', 6) +
        (q.remove ? '<button class="small danger" data-aremove="1" title="' + esc(q.remove.help || '') + '">✕ ' + esc(q.remove.label || 'Remove') + '</button>' : '') +
        (count ? '<button class="small" data-aclear="1">Clear</button>' : '') + '</div>' +
        '<div class="atools"><input type="search" data-afilter="1" placeholder="Find by name — e.g. “packing”, “office”" value="' + esc(asg.filter) + '">' +
        '<button class="small" data-aall="1">' + (allPicked(q) ? (asg.filter ? 'Deselect all matching' : 'Deselect all') : (asg.filter ? 'Select all matching' : 'Select all')) + '</button><button class="small" data-aexpand="1">' + (Object.keys(asg.open).length ? 'Collapse all' : 'Expand all') + '</button></div>';
      html += '<div class="aglist">';
      assignGroups(q).forEach(({ g, its }) => {
        const n1 = its.filter(it => asg.sel[it.key]).length;
        const tally = {};
        its.forEach(it => { const v = a[it.key] || '—'; tally[v] = (tally[v] || 0) + 1; });
        const gone = removed.filter(it => (it.group || '') === g).length;
        const sum = Object.keys(tally).map(k => '<span class="tp' + (k !== def ? ' chg' : '') + '">' + tally[k] + ' ' + esc(k) + '</span>').join('') + (gone ? '<span class="tp rm">' + gone + ' removed</span>' : '');
        const open = !!(asg.open[g] || asg.filter);
        html += '<div class="ag' + (open ? ' open' : '') + '" data-g="' + esc(g) + '"><div class="ag-h" title="Click to open · Ctrl/⌘+click or the box to select the site · Shift+click for a range of sites">' +
          '<span class="ck' + (n1 === its.length ? ' on' : n1 ? ' part' : '') + '" data-gsel="1">' + (n1 ? '✓' : '') + '</span>' +
          '<span class="nm">' + esc(g || (q.groupLabel || 'Other')) + '</span><span class="sum">' + sum + '</span><span class="tw">' + (open ? '▾' : '▸') + '</span></div>' +
          (open ? '<div class="ar-g">' + its.map(it => '<div class="ar' + (asg.sel[it.key] ? ' on' : '') + '" data-k="' + esc(it.key) + '"><span class="ck' + (asg.sel[it.key] ? ' on' : '') + '">' + (asg.sel[it.key] ? '✓' : '') + '</span>' +
            '<span class="nm">' + esc(it.label) + '</span><span class="pill' + (a[it.key] !== def ? ' chg' : '') + '">' + esc(a[it.key] || '—') + '</span></div>').join('') + '</div>' : '') + '</div>';
      });
      html += '</div>';
      // Removed rows: out of the list above, back with one click.
      if (removed.length) {
        html += '<div class="ag ag-rm"><div class="ag-h" data-rtoggle="1"><span class="nm">Removed — won’t be moved to 3.0 (' + removed.length + ')</span>' +
          '<button class="small" data-rall="1">Restore all</button><span class="tw">' + (asg.showRemoved ? '▾' : '▸') + '</span></div>' +
          (asg.showRemoved ? removed.map(it => '<div class="arr"><span class="nm">' + esc(it.label) + (it.group ? ' <span class="note-s">· ' + esc(it.group) + '</span>' : '') + '</span>' +
            '<button class="small" data-restore="' + esc(it.key) + '">Restore</button></div>').join('') : '') + '</div>';
      }
      return html;
    }
    function repaintAssign() {
      const q = cur(), body = $('.qbody');
      if (!q || q.type !== 'assign' || !body) return;
      const f = $('[data-afilter]'), had = document.activeElement === f, pos = f ? f.selectionStart : 0;
      body.innerHTML = assignBody(q);
      if (had) { const f2 = $('[data-afilter]'); f2.focus(); try { f2.setSelectionRange(pos, pos); } catch (e) { /* fine */ } }
    }

    // ─── places: one address for every site, or one each ───
    addrIO.places = {
      get: key => { const q = cur(), a = ans[q.id] || {}; if (key === '__all__') return a.__all__ || {}; return a[key] || {}; },
      set: (key, p) => {
        const q = cur(), a = ans[q.id] = ans[q.id] || {};
        if (key === '__all__') { a.__all__ = p; q.items.forEach(it => { a[it.key] = Object.assign({}, a[it.key] || {}, p); }); }
        else a[key] = Object.assign({}, a[key] || {}, p);
      }
    };
    function placesBody(q) {
      const a = ans[q.id] = ans[q.id] || {};
      const mode = a.__mode || '';
      let html = '<div class="mode"><button type="button" data-pmode="all"' + (mode === 'all' ? ' class="on"' : '') + '><b>One address for every site</b><span>e.g. your main office — enter it once</span></button>' +
        '<button type="button" data-pmode="each"' + (mode === 'each' ? ' class="on"' : '') + '><b>Each site has its own address</b><span>Enter an address per site</span></button></div>';
      (q.choices || []).forEach(c => {
        const vals = q.items.map(it => (a[it.key] || {})[c.key] || '');
        const same = vals.every(v => v === vals[0]) ? vals[0] : '';
        html += '<div class="blk"><div class="lbl">' + esc(c.label) + ' for every site' + (mode === 'each' ? ' (change single sites below)' : '') + '</div>' + seg(c.options, same, 'data-pall="' + esc(c.key) + '"', 8) + '</div>';
      });
      if (mode === 'all') html += '<div class="blk">' + addrBox('places', '__all__', 'Address for every site') + '</div>';
      if (mode === 'each') {
        const W = { address2: 'w-su', city: 'w-ci', state: 'w-st', zip: 'w-zip', country: 'w-co' };
        const filled = q.items.filter(it => complete(a[it.key])).length;
        html += '<div class="blk"><h3>Sites (' + filled + ' of ' + q.items.length + ' complete)</h3><p class="help">Paste a full address into a site’s Street box and it fills that row. Tab moves to the next box.</p>' +
          '<div class="twrap"><table class="ptable"><thead><tr><th class="w-nm">Site</th>' + PARTS.map(k => '<th' + (W[k] ? ' class="' + W[k] + '"' : '') + '>' + PART_LABEL[k] + '</th>').join('') +
          (q.choices || []).map(c => '<th class="w-ty">' + esc(c.label) + '</th>').join('') + '</tr></thead><tbody>' +
          q.items.map(it => {
            const p = a[it.key] || {};
            return '<tr class="addr' + (complete(p) ? ' full' : '') + '" data-scope="places" data-key="' + esc(it.key) + '"><td class="nm" title="' + esc(it.label) + '">' + esc(it.label) + '</td>' +
              PARTS.map(k => '<td data-pt="' + k + '"' + (NEED_DEFAULT.indexOf(k) >= 0 && !String(p[k] || '').trim() ? ' class="miss"' : '') + '><input type="text" data-part="' + k + '" value="' + esc(p[k] || '') + '" autocomplete="off" aria-label="' + esc(it.label + ' ' + PART_LABEL[k]) + '"' +
                (k === 'address1' ? ' placeholder="Paste full address"' : '') + '></td>').join('') +
              (q.choices || []).map(c => '<td>' + control('select', p[c.key] || '', c.options, 'data-pone="' + esc(c.key) + '" data-pk="' + esc(it.key) + '" aria-label="' + esc(c.label) + '"') + '</td>').join('') + '</tr>';
          }).join('') + '</tbody></table></div></div>';
      }
      if (!mode) html += '<p class="note-box">Choose one of the two options above to enter the address' + ((q.choices || []).length ? ', and pick the ' + q.choices.map(c => c.label.toLowerCase()).join(' and ') : '') + '.</p>';
      return html;
    }

    // ─── cards: one small form per item, in sections ───
    addrIO.cards = {
      get: key => (ans[cur().id] || {})[key] || {},
      set: (key, p) => { const a = ans[cur().id] = ans[cur().id] || {}; a[key] = Object.assign({}, a[key] || {}, p); },
      after: key => paintCard(key),
      // The card's address parts marked required (section 'addr' in the spec).
      need: () => cur().fields.filter(f => f.section === 'addr' && f.required).map(f => f.key)
    };
    function cardHtml(q, it) {
      const v = (ans[q.id] || {})[it.key] || {}, left = cardLeft(q, v);
      const field = f => {
        const val = v[f.key] || '', miss = f.required && !String(val).trim();
        const ctl = f.type === 'buttons' ? seg(f.options, val, 'data-cbtn="' + esc(f.key) + '"', 6)
          : control(f.type || 'text', val, f.options, 'data-cf="' + esc(f.key) + '"', false, f.placeholder);
        // A button group must not sit in a <label>: a click anywhere in a label
        // "clicks" its first control — which picked "Regular" by itself.
        const tag = f.type === 'buttons' ? 'div' : 'label';
        return '<' + tag + ' class="fld' + (miss ? ' miss' : '') + (f.full || f.type === 'buttons' ? ' full' : '') + '" data-f="' + esc(f.key) + '"' + (f.type === 'buttons' ? ' role="group" aria-label="' + esc(f.label) + '"' : '') + (shown(f, v) ? '' : ' style="display:none"') + '><span class="lt"><span>' + esc(f.label) + '</span>' +
          (f.required ? (miss ? '<span class="need">Required</span>' : '<span class="need ok">✓ Done</span>') : '<span class="opt">Optional</span>') + '</span>' + ctl + '</' + tag + '>';
      };
      return '<div class="ecard" data-card="' + esc(it.key) + '"><div class="ecard-h"><div class="t"><b>' + esc(v.name || it.label) + '</b>' + (it.sub ? '<span>' + esc(it.sub) + '</span>' : '') + '</div>' +
        '<span class="status ' + (left ? 'left">' + left + ' required left' : 'done">✓ Complete') + '</span></div>' +
        (q.sections || [{ key: '' }]).map(sct => {
          const fs = q.fields.filter(f => (f.section || '') === sct.key);
          const addr = q.address && sct.key === q.address;
          if (!fs.length && !addr) return '';
          return '<div class="sec"><div class="sec-h"><h3>' + esc(sct.label || '') + '</h3>' + (sct.hint ? '<p>' + esc(sct.hint) + '</p>' : '') + '</div><div class="sec-b">' +
            (addr ? addrBox('cards', it.key, 'Address') : '') + (fs.length ? '<div class="fgrid">' + fs.map(field).join('') + '</div>' : '') + '</div></div>';
        }).join('') + '</div>';
    }
    let cardTab = 0;
    const tabStatus = (q, it) => { const l = cardLeft(q, (ans[q.id] || {})[it.key]); return '<span class="tstat ' + (l ? 'left">' + l + ' left' : 'done">✓') + '</span>'; };
    function paintCard(key) {
      const q = cur(), card = $('.ecard[data-card="' + cssq(key) + '"]');
      if (!q || !card) return;
      const v = (ans[q.id] || {})[key] || {}, left = cardLeft(q, v);
      const tab = $('[data-ctab="' + q.items.findIndex(it => it.key === key) + '"]');
      const item = q.items.find(it => it.key === key);
      if (tab && item) tab.innerHTML = esc(v.name || item.label) + tabStatus(q, item);
      const st = $('.status', card);
      st.className = 'status ' + (left ? 'left' : 'done');
      st.textContent = left ? left + ' required left' : '✓ Complete';
      q.fields.forEach(f => {
        const el = $('[data-f="' + cssq(f.key) + '"]', card);
        if (!el) return;
        el.style.display = shown(f, v) ? '' : 'none';
        const miss = !!(f.required && !String(v[f.key] || '').trim());
        el.classList.toggle('miss', miss);
        const tag = $('.need', el);
        if (tag) { tag.classList.toggle('ok', !miss); tag.textContent = miss ? 'Required' : '✓ Done'; }
      });
    }

    // The three pages with more than one thing to do get a 1-2-3.
    const HOW = {
      assign: ['Open a site — or tick its box to take every block on it', 'Click blocks to select more (Shift+click a range, Ctrl+click one at a time)', 'Press the type they should be, or Remove to leave them out of 3.0'],
      places: ['Choose one address for every site, or one each', 'Pick each site’s type', 'Paste each address into Street — it fills in the rest'],
      cards: ['Open each employer’s tab', 'Fill in everything marked Required', 'The tab turns green when that employer is done']
    };

    // ─── Page tracker: one segment per item on the page ───
    function trackOf(q) {
      const a = ans[q.id] || {};
      if (q.type === 'pick') return q.items.map(it => ({ label: it.label, done: !!(a[it.key] && String(a[it.key]).trim()) }));
      if (q.type === 'places') return q.items.map(it => ({ label: it.label, done: complete(a[it.key]) && (q.choices || []).every(c => (a[it.key] || {})[c.key]) }));
      if (q.type === 'cards') return q.items.map(it => ({ label: (a[it.key] || {}).name || it.label, done: !cardLeft(q, a[it.key]) }));
      if (q.type === 'choices') return q.parts.map(p => ({ label: p.title, done: !!a[p.key] }));
      if (q.type === 'assign') return null;
      return [{ label: q.title, done: done(q) }];
    }
    function paintTracker() {
      const el = $('.trk'), q = cur();
      if (!el || !q) return;
      const t = trackOf(q);
      if (!t) {                       // block types: a tally instead
        const a = ans[q.id] || {}, def = q.items[0] ? q.items[0].value : '';
        const rm = q.items.filter(it => a[it.key] === REMOVED).length, ch = q.items.filter(it => a[it.key] !== REMOVED && a[it.key] !== def).length;
        el.innerHTML = '<div class="trk-t"><span><b>' + q.items.length + '</b> blocks — <b>' + ch + '</b> changed, <b>' + rm + '</b> removed. Everything else stays “' + esc(def) + '”.</span></div>';
        return;
      }
      const d = t.filter(x => x.done).length;
      el.innerHTML = '<div class="trk-t"><span><b>' + d + ' of ' + t.length + '</b> done on this page</span>' +
        (d === t.length ? '<span class="all">✓ All done here</span>' : '<button type="button" class="small" data-nextempty="1">Go to next empty</button>') + '</div>' +
        (t.length > 1 ? '<div class="tk">' + t.map((x, i) => '<button type="button" class="tki' + (x.done ? ' ok' : '') + '" data-jump="' + i + '" title="' + esc(x.label) + (x.done ? ' — done' : ' — not done yet') + '"></button>').join('') + '</div>' : '');
    }
    // Take them to item i: its tab, row or box, with the first empty field focused.
    function jump(i) {
      const q = cur(), a = ans[q.id] || {};
      let el = null;
      if (q.type === 'cards') {
        cardTab = i; rerenderBody(); paintTracker();
        el = $('.ecard .miss input, .ecard .miss select') || $('.ecard .addr-in');
        const card = $('.ecard');
        if (card) card.scrollIntoView({ block: 'start', behavior: 'smooth' });
        if (el) el.focus({ preventScroll: true });
        return;
      }
      if (q.type === 'pick') el = $('[data-k="' + cssq(q.items[i].key) + '"]');
      if (q.type === 'places') {
        if (a.__mode !== 'each') el = $('.addr .miss input') || $('.addr-in') || $('[data-pmode]');
        else { const row = $('tr.addr[data-key="' + cssq(q.items[i].key) + '"]'); el = row && ($('td.miss input', row) || $('select', row)); }
      }
      if (q.type === 'choices') { const p = $$('.cpart')[i]; el = p && $('input', p); }
      if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); }
    }

    // A few items, each with what we know about it, and room to answer.
    function factHtml(fc) {
      const v = Array.isArray(fc[1]) ? '<span class="chips">' + fc[1].map(x => '<span class="chip">' + esc(x) + '</span>').join('') + '</span>' : esc(fc[1]);
      return '<div><dt>' + esc(fc[0]) + '</dt><dd>' + v + '</dd></div>';
    }
    function detailBody(q, a) {
      const groups = [];
      q.items.forEach(it => { const g = it.group || ''; const last = groups[groups.length - 1]; if (last && last.g === g) last.its.push(it); else groups.push({ g, its: [it] }); });
      const ctl = (val, attrs) => control(q.other ? 'text' : 'select', val || '', q.options, attrs, q.other, q.placeholder, q.extra);
      return groups.map(gr => {
        const vals = gr.its.map(it => a[it.key] || ''), same = vals.every(v => v === vals[0]) ? vals[0] : '';
        const head = gr.g ? '<div class="dg-h"><h3>' + esc(gr.g) + '</h3>' +
          (q.groupPick && gr.its.length > 1 ? '<label class="dg-all"><span>' + esc(q.groupPick) + '</span>' + ctl(same, 'data-kgroup="' + esc(gr.g) + '"') + '</label>' : '') + '</div>' : '';
        return '<section class="dg">' + head + gr.its.map(it =>
          '<div class="drow item' + (a[it.key] ? ' done' : '') + '"><div class="dmain"><b class="dt">' + esc(it.label) + '</b>' +
            (it.sub ? '<p class="dsub">' + esc(it.sub) + '</p>' : '') +
            (it.facts && it.facts.length ? '<dl class="facts">' + it.facts.map(factHtml).join('') + '</dl>' : '') + '</div>' +
            '<label class="dans"><span>' + esc(q.answerLabel || 'Answer') + '</span>' + ctl(a[it.key], 'data-k="' + esc(it.key) + '"') + '</label></div>').join('') + '</section>';
      }).join('');
    }

    function questionHtml(q) {
      const a = ans[q.id];
      let body = '';
      if (q.type === 'pick' && q.layout === 'detail') body = detailBody(q, a || {});
      else if (q.type === 'pick') {
        const groups = [];
        q.items.forEach(it => { const g = it.group || ''; const last = groups[groups.length - 1]; if (last && last.g === g) last.its.push(it); else groups.push({ g, its: [it] }); });
        const tile = it => '<div class="item' + (a && a[it.key] ? ' done' : '') + '"><div class="lab"><b title="' + esc(it.label) + '">' + esc(it.label) + '</b>' +
          (it.sub ? '<span title="' + esc(it.sub) + '">' + esc(it.sub) + '</span>' : '') + '</div>' +
          control(q.other ? 'text' : 'select', a ? a[it.key] : '', q.options, 'data-k="' + esc(it.key) + '"', q.other, q.placeholder, q.extra) + '</div>';
        body = '<div class="items">' + groups.map(gr => {
          const big = gr.its.length > 3 || !gr.g;
          return '<div class="grp' + (big ? ' all' : '') + '">' + (gr.g ? '<div class="grp-h" title="' + esc(gr.g) + '">' + esc(gr.g) + '</div>' : '') + '<div class="grp-b">' + gr.its.map(tile).join('') + '</div></div>';
        }).join('') + '</div>';
      } else if (q.type === 'assign') body = assignBody(q);
      else if (q.type === 'places') body = placesBody(q);
      else if (q.type === 'cards') {
        // One tab per item; only the open one's form is on screen.
        if (cardTab >= q.items.length) cardTab = 0;
        body = (q.items.length > 1 ? '<div class="tabs" role="tablist">' + q.items.map((it, i) => '<button type="button" role="tab" data-ctab="' + i + '" aria-selected="' + (i === cardTab) + '"' +
          (i === cardTab ? ' class="on"' : '') + '>' + esc(((ans[q.id] || {})[it.key] || {}).name || it.label) + tabStatus(q, it) + '</button>').join('') + '</div>' : '') +
          '<div class="cards">' + cardHtml(q, q.items[cardTab]) + '</div>' +
          (q.items.length > 1 ? '<div class="tab-nav">' + (cardTab > 0 ? '<button type="button" class="small" data-ctab="' + (cardTab - 1) + '">← ' + esc(q.items[cardTab - 1].label) + '</button>' : '<span></span>') +
            (cardTab < q.items.length - 1 ? '<button type="button" class="small" data-ctab="' + (cardTab + 1) + '">' + esc(q.items[cardTab + 1].label) + ' →</button>' : '') + '</div>' : '');
      }
      else if (q.type === 'choices') {
        // One table of real examples; each part adds a column with what its pick gives.
        const av = a || {}, ex = q.examples || { columns: [], rows: [] };
        const cols = ex.columns.concat(q.parts.map(p => p.column));
        const cell = c => String(c == null ? '' : c).split('`').map((x, i) => i % 2 ? '<code>' + esc(x) + '</code>' : esc(x)).join('');
        body = '<div class="ex"><table><thead><tr>' + cols.map((c, i) => '<th' + (i >= ex.columns.length ? ' class="res"' : '') + '>' + esc(c) + '</th>').join('') + '</tr></thead><tbody>' +
          ex.rows.map((r, ri) => '<tr>' + r.map(c => '<td>' + cell(c) + '</td>').join('') + q.parts.map(p => {
            const ch = p.choices.find(c => c.value === av[p.key]);
            return '<td class="res">' + (ch ? '<b>' + cell((ch.cells || [])[ri] || '— empty —') + '</b>' : '<span class="wait">pick below</span>') + '</td>';
          }).join('') + '</tr>').join('') + '</tbody></table></div>' +
          q.parts.map((p, pi) => '<div class="cpart"><h3><span class="num">' + (pi + 1) + '</span>' + esc(p.title) + '</h3>' + (p.help ? '<p class="help">' + esc(p.help) + '</p>' : '') +
            '<div class="choices">' + p.choices.map(c => {
              const on = av[p.key] === c.value;
              return '<label class="choice' + (on ? ' on' : '') + '"><input type="radio" name="ch-' + esc(p.key) + '" data-cpart="' + esc(p.key) + '" value="' + esc(c.value) + '"' + (on ? ' checked' : '') + '><div><div>' + esc(c.label) + '</div>' +
                (c.note ? '<div class="note">' + esc(c.note) + '</div>' : '') + '</div></label>';
            }).join('') + '</div></div>').join('');
      }
      else if (q.type === 'choice') {
        const ex = q.examples;
        body = (ex ? '<div class="ex"><table><thead><tr>' + ex.columns.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr></thead><tbody>' +
          ex.rows.map(r => '<tr>' + r.map(c => '<td>' + String(c == null ? '' : c).split('`').map((s, i) => i % 2 ? '<code>' + esc(s) + '</code>' : esc(s)).join('') + '</td>').join('') + '</tr>').join('') + '</tbody></table></div>' : '') +
          '<div class="choices">' + q.choices.map(c => {
            const on = a && a.choice === c.value;
            return '<label class="choice' + (on ? ' on' : '') + '"><input type="radio" name="ch" value="' + esc(c.value) + '"' + (on ? ' checked' : '') + '><div><div>' + esc(c.label) + '</div>' +
              (c.note ? '<div class="note">' + esc(c.note) + '</div>' : '') + '</div></label>';
          }).join('') + '</div>';
      } else body = '<textarea data-text="1" placeholder="' + esc(q.placeholder || 'Type your answer…') + '">' + esc(a || '') + '</textarea>';
      const how = HOW[q.type];
      return '<div class="q"><div class="qh"><div class="qn step">Question ' + step + ' of ' + n + '</div><h1>' + esc(q.title) + '</h1>' +
        (q.help ? '<p class="help">' + esc(q.help) + '</p>' : '') +
        (how ? '<ol class="how">' + how.map((t, i) => '<li><span class="n">' + (i + 1) + '</span><span>' + esc(t) + '</span></li>').join('') + '</ol>' : '') + '</div>' +
        '<div class="qbody">' + body + '</div></div>';
    }

    function cmtHtml(q) {
      const nt = noteOf(q) || { text: '', skip: false }, text = nt.text || '';
      const box = (lbl, ph) => '<label class="cmt-l">' + lbl + '</label><textarea data-note="1" rows="3" placeholder="' + esc(ph) + '">' + esc(text) + '</textarea>';
      if (skipped(q)) return '<div class="cmt sk" id="cmt"><div class="cmt-h"><b>You skipped this question.</b> We’ll read your comment when we load your answers. Answer it any time and it counts as answered.</div>' +
        box('Your comment for PickTrace', '') + '<div class="cmt-a"><button type="button" class="small" data-unskip="1">Undo skip</button></div></div>';
      if (blocked && !done(q)) {
        const t = trackOf(q), left = t ? t.filter(x => !x.done).length : 1;
        return '<div class="cmt warn" id="cmt"><div class="cmt-h"><b>This question isn’t finished' + (t && t.length > 1 ? ' — ' + left + ' of ' + t.length + ' still empty' : '') + '.</b> Fill it in to continue' +
          (t && t.length > 1 ? ' <button type="button" class="small" data-nextempty="1">Go to next empty</button>' : '') + '</div>' +
          box('Can’t finish it? Tell us why, and skip it for now', 'e.g. “We don’t know the address for site 12 yet — will send next week.”') +
          '<div class="cmt-a"><button type="button" class="primary" data-skip="1"' + (text.trim() ? '' : ' disabled') + '>Skip with this comment</button><span class="cmt-n">A comment is needed to skip.</span></div></div>';
      }
      if (cmtOpen || text.trim()) return '<div class="cmt" id="cmt">' + box('Comment for PickTrace (optional)', 'Anything we should know about this question?') + '<p class="cmt-n">We’ll see this when we load your answers.</p></div>';
      return '<div class="cmt closed" id="cmt"><button type="button" class="link" data-cmtopen="1">💬 Add a comment for PickTrace</button></div>';
    }
    function paintCmt() { const q = cur(), el = $('#cmt'); if (q && el) el.outerHTML = cmtHtml(q); }

    function render() {
      let main = '';
      const listHtml = () => '<ul class="list">' + Q.map((q, i) => { const sk = skipped(q), nt = noteOf(q), lock = i + 1 > reach();
        return '<li><span class="dot' + (done(q) ? ' ok' : sk ? ' sk' : '') + '">' + (done(q) ? '✓' : sk ? '!' : i + 1) + '</span><div class="t"><b>' + esc(q.title) + '</b>' +
          '<span>' + (sk ? 'Skipped' : esc(q.where || '')) + '</span>' + (nt && String(nt.text || '').trim() ? '<q class="lq">' + esc(nt.text) + '</q>' : '') + '</div>' +
          (lock ? '<span class="lk">Locked</span>' : '<button class="link" data-go="' + (i + 1) + '">' + (done(q) ? 'Change' : sk ? 'Answer now' : 'Answer') + '</button>') + '</li>'; }).join('') + '</ul>';
      if (step === 0) {
        main = '<div class="qh"><div class="qn step">Before you start</div><h1>' + esc(spec.meta.title || 'A few questions') + '</h1><p class="help">' + esc(spec.meta.intro || '') + '</p></div>' + listHtml() +
          '<p class="note-box">Your answers are kept on this computer as you go. When you’re finished, click <b>Save my answers</b> and send us the file it downloads.</p>';
      } else if (step <= n) main = questionHtml(Q[step - 1]) + cmtHtml(Q[step - 1]);
      else {
        const left = Q.filter(q => !settled(q)), sks = Q.filter(skipped);
        main = '<div class="qh"><div class="qn step">Last step</div><h1>' + (left.length ? 'Almost there' : 'All done — thank you') + '</h1>' +
          '<p class="help">' + (left.length ? left.length + ' question' + (left.length === 1 ? ' still needs' : 's still need') + ' an answer. You can send what you have now and finish the rest later.'
            : 'Click below to save your answers, then send us the file it downloads.' + (sks.length ? ' You skipped ' + sks.length + ' — we’ll read your comment' + (sks.length === 1 ? '' : 's') + ' and follow up.' : '')) + '</p></div>' + listHtml() +
          '<p style="margin-top:22px"><button class="primary big" data-save="1">⬇ Save my answers</button></p><p class="saved">' + (dirty ? 'Saved on this computer — click “Save my answers” to send them.' : '') + '</p>';
      }
      $app.innerHTML = main +
        '<div class="foot"><div class="fin">' + (step > 0 ? '<button data-back="1">Back</button>' : '') +
        (step >= 1 && step <= n ? '<div class="trk"></div>' : '<span class="sp"></span>') +
        (step <= n ? '<button class="primary" data-next="1">' + (step === 0 ? 'Start' : step === n ? 'Finish' : 'Next question') + '</button>' : '<button class="primary" data-save="1">⬇ Save my answers</button>') +
        '</div></div>';
      paintChrome();
      window.scrollTo(0, 0);
    }
    function rerenderBody() {
      const q = cur(), body = $('.qbody');
      if (!q || !body) return;
      const y = window.scrollY;
      body.innerHTML = questionHtml(q).replace(/^[\s\S]*?<div class="qbody">/, '').replace(/<\/div><\/div>$/, '');
      window.scrollTo(0, y);
    }

    // ─── Events ──────────────────────────────────────────────────────────
    $app.addEventListener('input', e => {
      const t = e.target, q = cur();
      if (!q) return;
      if (t.dataset.part) {
        const box = t.closest('.addr'), io = addrIO[box.dataset.scope], p = Object.assign({}, io.get(box.dataset.key) || {});
        p[t.dataset.part] = t.dataset.part === 'state' ? (stateCode(t.value) || t.value.toUpperCase()) : t.value;
        io.set(box.dataset.key, p); paintAddr(box); keep(); if (io.after) io.after(box.dataset.key);
        return;
      }
      if (t.dataset.afilter) { asg.filter = t.value; repaintAssign(); return; }
      if (t.dataset.note) {
        setNote(q, { text: t.value });
        const sb = $('[data-skip]'); if (sb) sb.disabled = !t.value.trim();
        keep();
        return;
      }
      if (t.dataset.kgroup != null) {
        const a = ans[q.id] = ans[q.id] || {};
        q.items.forEach(it => {
          if ((it.group || '') !== t.dataset.kgroup) return;
          a[it.key] = t.value;
          const el = $('[data-k="' + cssq(it.key) + '"]');
          if (el) { el.value = t.value; const row = el.closest('.item'); if (row) row.classList.toggle('done', !!t.value.trim()); }
        });
        keep();
        return;
      }
      if (t.dataset.k != null) { (ans[q.id] = ans[q.id] || {})[t.dataset.k] = t.value; const it = t.closest('.item'); if (it) it.classList.toggle('done', !!t.value.trim()); }
      else if (t.dataset.cf) { const card = t.closest('.ecard'), a = ans[q.id] = ans[q.id] || {}; (a[card.dataset.card] = a[card.dataset.card] || {})[t.dataset.cf] = t.value; paintCard(card.dataset.card); }
      else if (t.dataset.pone) { const a = ans[q.id] = ans[q.id] || {}; (a[t.dataset.pk] = a[t.dataset.pk] || {})[t.dataset.pone] = t.value; }
      else if (t.dataset.cpart) { (ans[q.id] = ans[q.id] || {})[t.dataset.cpart] = t.value; keep(); rerenderBody(); return; }
      else if (t.type === 'radio') { ans[q.id] = { choice: t.value }; $$('.choice').forEach(c => c.classList.toggle('on', c.contains(t))); }
      else if (t.dataset.text) ans[q.id] = t.value;
      else return;
      keep();
    });
    $app.addEventListener('change', e => {
      const t = e.target;
      if (t.dataset.more) {           // "More…" list inside a button group
        const b = document.createElement('button');
        b.dataset.v = t.value;
        t.closest('.seg').appendChild(b);
        b.click(); b.remove();
        return;
      }
      if (t.matches('select, input[type=date]')) t.dispatchEvent(new Event('input', { bubbles: true }));
    });
    // Paste a whole address into the paste box — or into the Street box (with
    // commas or several lines) — and it's split into the boxes at once.
    $app.addEventListener('paste', e => {
      const t = e.target;
      if (!t.classList) return;
      const inPaste = t.classList.contains('addr-in'), inStreet = t.dataset && t.dataset.part === 'address1';
      if (!inPaste && !inStreet) return;
      const text = e.clipboardData && e.clipboardData.getData('text');
      if (!text) return;
      if (inStreet && !/[,\r\n]/.test(text.trim())) return;   // just a street — paste as normal
      e.preventDefault();
      resolveText(t.closest('.addr'), text);
    });
    // A whole address typed into Street ("…, Anytown, IL 62701") splits on
    // Enter or on leaving the box; a plain street stays as typed.
    const looksFull = v => /,/.test(v) || /\b[A-Za-z]{2}\s+\d{5}(-\d{4})?\s*$/.test(v);
    function splitStreet(t) {
      if (!looksFull(t.value)) return false;
      resolveText(t.closest('.addr'), t.value);
      return true;
    }
    $app.addEventListener('keydown', e => {
      const t = e.target;
      if (e.key !== 'Enter' || !t.classList) return;
      if (t.classList.contains('addr-in')) { e.preventDefault(); resolve(t); }
      else if (t.dataset.part === 'address1' && splitStreet(t)) e.preventDefault();
    });
    $app.addEventListener('focusout', e => {
      const t = e.target;
      if (!t.classList) return;
      if (t.classList.contains('addr-in') && t.value.trim()) resolve(t);
      else if (t.dataset.part === 'address1') splitStreet(t);
    });
    $app.addEventListener('click', e => {
      const q = cur();
      // assign: rows, groups, set
      const row = e.target.closest('.ar');
      if (row && q && q.type === 'assign') {
        // Click or Ctrl/⌘+click: add or take away one row. Shift+click: the
        // range from the last row clicked — selected, or deselected when the
        // row you Shift+click is already selected.
        const k = row.dataset.k, keys = $$('.ar').map(r => r.dataset.k);
        if (e.shiftKey && asg.last && keys.indexOf(asg.last) >= 0) {
          const on = !asg.sel[k];
          const [i, j] = [keys.indexOf(asg.last), keys.indexOf(k)].sort((x, y) => x - y);
          keys.slice(i, j + 1).forEach(x => { if (on) asg.sel[x] = 1; else delete asg.sel[x]; });
          document.getSelection && document.getSelection().removeAllRanges();
        } else if (asg.sel[k]) delete asg.sel[k]; else asg.sel[k] = 1;
        asg.last = k;
        repaintAssign();
        return;
      }
      // assign: the Removed section — show / restore
      const rb = e.target.closest('[data-restore], [data-rall]');
      if (rb && q && q.type === 'assign') {
        const a = ans[q.id] = ans[q.id] || {};
        q.items.forEach(it => { if ((rb.dataset.rall || rb.dataset.restore === it.key) && a[it.key] === REMOVED) a[it.key] = it.value || ''; });
        keep(); repaintAssign();
        return;
      }
      if (e.target.closest('[data-rtoggle]') && q && q.type === 'assign') { asg.showRemoved = !asg.showRemoved; repaintAssign(); return; }
      // assign: a site's header. The box, Ctrl/⌘+click: select or drop the
      // whole site. Shift+click: every site from the last one clicked to this
      // one — selected, or dropped if this one is already selected. Anywhere
      // else on the header: open / close it.
      const gh = e.target.closest('.ag-h');
      if (gh && q && q.type === 'assign') {
        const g = gh.parentNode.dataset.g;
        if (e.target.closest('[data-gsel]') || e.ctrlKey || e.metaKey || e.shiftKey) {
          const groups = assignGroups(q), gi = groups.findIndex(x => x.g === g);
          const isOn = x => x.its.every(it => asg.sel[it.key]);
          let span = [groups[gi]], on = !isOn(groups[gi]);
          if (e.shiftKey && asg.lastG != null) {
            const li = groups.findIndex(x => x.g === asg.lastG);
            if (li >= 0) span = groups.slice(Math.min(li, gi), Math.max(li, gi) + 1);
          }
          span.forEach(x => x.its.forEach(it => { if (on) asg.sel[it.key] = 1; else delete asg.sel[it.key]; }));
          asg.lastG = g;
          if (document.getSelection) document.getSelection().removeAllRanges();
        } else if (asg.open[g]) delete asg.open[g]; else asg.open[g] = 1;
        repaintAssign();
        return;
      }
      const b = e.target.closest('button');
      if (!b) return;
      const segEl = b.closest('.seg');
      if (segEl && b.dataset.v != null && q) {
        const v = b.dataset.v;
        if (segEl.dataset.aset) {
          const keys = Object.keys(asg.sel);
          if (!keys.length) { alert('Select blocks first — click them, or a site’s box to select the whole site.'); return; }
          const a = ans[q.id] = ans[q.id] || {};
          keys.forEach(k => { a[k] = v; });
          asg.sel = {};
          keep(); repaintAssign();
          return;
        }
        if (segEl.dataset.pall) {
          const a = ans[q.id] = ans[q.id] || {};
          q.items.forEach(it => { (a[it.key] = a[it.key] || {})[segEl.dataset.pall] = v; });
          keep(); rerenderBody();
          return;
        }
        if (segEl.dataset.cbtn) {
          const card = b.closest('.ecard'), a = ans[q.id] = ans[q.id] || {};
          (a[card.dataset.card] = a[card.dataset.card] || {})[segEl.dataset.cbtn] = v;
          $$('button', segEl).forEach(x => x.classList.toggle('on', x.dataset.v === v));
          keep(); paintCard(card.dataset.card);
          return;
        }
      }
      if (b.dataset.aremove) {
        const keys = Object.keys(asg.sel);
        if (!keys.length) { alert('Select the blocks to remove first.'); return; }
        const a = ans[q.id] = ans[q.id] || {};
        keys.forEach(k => { a[k] = REMOVED; });
        asg.sel = {};
        keep(); repaintAssign();
        return;
      }
      if (b.dataset.aclear) { asg.sel = {}; repaintAssign(); return; }
      // Every block in the list — open sites or not — that matches the search; never removed ones.
      if (b.dataset.aall) { const off = allPicked(q); assignGroups(q).forEach(x => x.its.forEach(it => { if (off) delete asg.sel[it.key]; else asg.sel[it.key] = 1; })); repaintAssign(); return; }
      if (b.dataset.aexpand) { if (Object.keys(asg.open).length) asg.open = {}; else q.items.forEach(it => { asg.open[it.group || ''] = 1; }); repaintAssign(); return; }
      if (b.dataset.pmode) { const a = ans[q.id] = ans[q.id] || {}; a.__mode = b.dataset.pmode; keep(); rerenderBody(); return; }
      if (b.dataset.ctab != null) { cardTab = +b.dataset.ctab; rerenderBody(); return; }
      if (b.dataset.jump != null) { jump(+b.dataset.jump); return; }
      if (b.dataset.cmtopen) { cmtOpen = true; paintCmt(); const ta = $('#cmt textarea'); if (ta) ta.focus(); return; }
      if (b.dataset.skip) { const nt = noteOf(q); if (!nt || !String(nt.text || '').trim()) { const ta = $('#cmt textarea'); if (ta) ta.focus(); return; } setNote(q, { skip: true }); keep(); goStep(step + 1); return; }
      if (b.dataset.unskip) { setNote(q, { skip: false }); blocked = false; keep(); render(); return; }
      if (b.dataset.nextempty) { const t = trackOf(q) || [], i = t.findIndex(x => !x.done); if (i >= 0) jump(i); return; }
      if (b.dataset.next) next();
      else if (b.dataset.back) goStep(Math.max(0, step - 1));
      else if (b.dataset.go) goStep(+b.dataset.go);
      else if (b.dataset.save) save();
    });

    // The download is this page, with the answers written into it.
    function goStep(target) {
      blocked = false; cmtOpen = false;
      if (target > reach()) { step = reach(); blocked = step <= n; }
      else step = target;
      render();
    }
    function next() {
      const q = cur();
      if (q && !settled(q)) { blocked = true; paintCmt(); const c = $('#cmt'); if (c) c.scrollIntoView({ block: 'center', behavior: 'smooth' }); return; }
      goStep(step + 1);
    }
    function save() {
      const open = Q.filter(q => !settled(q));
      if (open.length && !confirm(open.length + ' question' + (open.length === 1 ? ' isn’t' : 's aren’t') + ' answered or skipped yet. Save what you have so far? You can finish the rest later.')) return;
      const stamp = { saved: Date.now(), answers: ans };
      store.textContent = JSON.stringify(stamp).replace(/</g, '\\u003c');
      // Copy the page without the drawn parts — on a copy, so the live page
      // (and the button just clicked) is never pulled out from under the click.
      const copy = document.documentElement.cloneNode(true);
      ['#app', '#imt-nav'].forEach(sel => { const el = copy.querySelector(sel); if (el) el.innerHTML = ''; });
      const html = '<!doctype html>\n' + copy.outerHTML;
      const blob = new Blob([html], { type: 'text/html' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = spec.meta.answerName || 'PickTrace answers.html';
      document.body.appendChild(a); a.click(); a.remove();
      try { localStorage.setItem(KEY, JSON.stringify(stamp)); } catch (e) { /* fine */ }
      dirty = false;
      const s = $('.saved');
      if (s) s.textContent = 'Downloaded “' + a.download + '” — send that file back to us.';
    }
    // Sidebar and top bar only — everything in #app has its own handler above.
    document.addEventListener('click', e => {
      if (!e.target.closest('.side, .tbar')) return;
      const b = e.target.closest('[data-go], [data-save]');
      if (!b) return;
      if (b.dataset.save) save();
      else goStep(+b.dataset.go);
    });
    render();
  }

  function html(spec) {
    const meta = spec.meta || {};
    return '<!doctype html>\n<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>' + esc(meta.title || 'A few questions') + '</title><style>' + CSS + '</style></head><body>' +
      '<aside class="side"><div class="logo"><img src="' + LOGO + '" alt="PickTrace"></div><div class="side-l">Your questions</div><nav id="imt-nav"></nav><div class="side-f" id="imt-foot"></div></aside>' +
      // The org name only when it adds something (older pages repeated "PickTrace 3.0" here).
      '<header class="tbar"><div class="tt">' + esc(meta.title || 'A few questions') +
        (meta.org && String(meta.title || '').indexOf(meta.org) < 0 ? '<span>' + esc(meta.org) + '</span>' : '') + '</div>' +
      '<span class="count" id="imt-count"></span><div class="bar"><i id="imt-bar"></i></div><button class="primary" data-save="1" id="imt-save">⬇ Save my answers</button></header>' +
      '<main class="main"><div class="inner" id="app"><noscript>Please open this file in a web browser (Chrome, Edge, Safari or Firefox).</noscript></div></main>' +
      '<script type="application/json" id="imt-spec">' + safeJson(spec) + '</script>' +
      '<script type="application/json" id="imt-answers">{}</script>' +
      '<script>(' + app.toString() + ')();</script></body></html>';
  }

  // Pull the questions and answers back out of a (filled-in) page.
  function read(text) {
    const grab = id => {
      const m = String(text).match(new RegExp('<script type="application/json" id="' + id + '">([\\s\\S]*?)</script>'));
      if (!m) return null;
      try { return JSON.parse(m[1]); } catch (e) { return null; }
    };
    const spec = grab('imt-spec'), stored = grab('imt-answers');
    if (!spec || !spec.meta) return null;
    return { meta: spec.meta, questions: spec.questions || [], answers: (stored && stored.answers) || {}, saved: (stored && stored.saved) || 0 };
  }

  return { html: html, read: read };
})();
