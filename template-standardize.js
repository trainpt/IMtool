// ═══════════════════════════════════════════════════════════════════════
// Template Standardize — formats source rows (Excel/CSV or Org Data) into
// a personalized PickTrace bulk Create or Update template. Auto-maps
// columns by name; surfaces Sites and Crop & Variety values that aren't
// in the template's dropdown so the user can create them in PickTrace
// before bulk uploading. Required (*) cells must all be filled; the
// "Empty Required Columns" picklist drives single-click bulk-fills.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const norm = h => String(h == null ? '' : h).trim().toLowerCase().replace(/^#/, '');

  // Today as YYYY-MM-DD (PickTrace's bulk-template format).
  function todayYMD() {
    const d = new Date();
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  }
  // True if the given string parses to a date strictly BEFORE today (local time).
  // Unparseable strings return false (we don't penalize unrecognized formats).
  // Critically, parses YYYY-MM-DD and MM/DD/YYYY strings as LOCAL midnight —
  // `new Date("2026-05-05")` interprets ISO-style strings as UTC, which in any
  // timezone west of UTC becomes "yesterday local" and gets flagged as past.
  function isDateBeforeToday(s) {
    if (!s) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const str = String(s).trim();
    let d = null;
    let m;
    if ((m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) {
      d = new Date(+m[1], +m[2] - 1, +m[3]);
    } else if ((m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) {
      d = new Date(+m[3], +m[1] - 1, +m[2]);
    } else {
      d = new Date(str);
    }
    if (!d || isNaN(d.getTime())) return false;
    d.setHours(0, 0, 0, 0);
    return d.getTime() < today.getTime();
  }

  // Bulk template column schemas (matches PickTrace's templates exactly).
  const CREATE_HEADERS = [
    'Site*','Name*','Alt ID','Location Type*','Crop & Variety*','Planted At','Acreage','Length',
    'Plant Count','Start Date*','Location Group','Lot Number','Rootstock','Plant/Seed Seller',
    'Plant/Seed Producer','Clone/Subvariety','Growing Manager','Production Status','Training Style',
    'Trellis Type','Mulch Type','Row Direction','Organic Status','Organic Certifier','Wet Date',
    'Germination Date','Planted Date','Grafting Date','Production Start','Organic Certification Date',
    'Stand Count','Row/Bed Count','Post Count','Percent Covered','Row Spacing, in.','Plant Spacing, in.',
    'Post Spacing, in.','Bed Width, in.','Custom Data 1','Custom Data 2','Custom Data 3'
  ];
  const UPDATE_HEADERS = CREATE_HEADERS.concat(['Is Archived']);

  // ─── State ───
  let srcData = null;          // { headers, rows, fileName, sheetName }
  let tplData = null;          // { headers, rows, dropdowns, fileName, rawBuffer }
  let target = 'create';
  let mapping = {};            // bulkColIdx → srcColIdx (-1 = unmapped)
  let formattedRows = null;    // [[...]] mapped + filled rows (canonical state)
  let manualFills = {};        // bulkColIdx → user-supplied value (re-applied after rebuilds)
  let removedRows = new Set(); // formattedRows indexes excluded from preview + export
  let initialized = false;

  // ─── Mode-pill switcher ───
  function attachModeSwitcher() {
    document.querySelectorAll('.cmp-mode-pill').forEach(btn => {
      btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        document.querySelectorAll('.cmp-mode-pill').forEach(b => b.classList.toggle('cmp-mode-active', b === btn));
        $('cmp-mode-compare').style.display = mode === 'compare' ? '' : 'none';
        $('cmp-mode-standardize').style.display = mode === 'standardize' ? '' : 'none';
      });
    });
  }

  // ─── Header-row detection ───
  // Soto Bros / PickTrace implementation templates have 3-5 rows of metadata
  // (sheet title, "Instructions", long description, "Basics") before the real
  // column header row. Score every candidate row by:
  //   - count of non-empty cells (more = better)
  //   - presence of canonical PickTrace template keywords (heavy bonus)
  //   - short, header-like cell text (small bonus per cell)
  //   - long descriptive sentences (penalty — those are instruction rows)
  // Pick the highest-scoring row in the first ~15 rows.
  const TEMPLATE_KEYWORDS = [
    'site','name','block','location','crop','variety','plant',
    'acreage','acres','archived','date','type','count','group',
    'alt id','lot number','seller','producer','rootstock','planted',
    'organic','training','trellis','mulch','spacing','custom data',
    'bed width','stand count','row spacing','plant count'
  ];
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
      return s.startsWith('required.') || s.startsWith('please ') ||
             s.startsWith('this field') || s.startsWith('this tab') ||
             s.startsWith('the unique identifier') || s.startsWith('an alternate id') ||
             s.includes('used to collect') || s.includes('used in reporting');
    });
  }

  // Slice a sheet's AOA into clean { headers, rows } using the auto-detected
  // header row. Trims metadata above and instruction rows below.
  function sliceSheetAoa(aoa) {
    const hIdx = detectHeaderRow(aoa);
    const headers = (aoa[hIdx] || []).map(h => String(h == null ? '' : h).trim());
    const rows = aoa.slice(hIdx + 1)
      .filter(r => !isInstructionRow(r))
      .map(r => r.map(c => c == null ? '' : String(c).trim()));
    return { headers, rows, headerRow: hIdx };
  }

  // ─── Sheet picker for multi-sheet workbooks ───
  // Shows a modal listing every sheet that has at least one row of data.
  // Returns a Promise that resolves with the picked sheet (or rejects on cancel).
  function pickSheet(sheets, fileName) {
    return new Promise((resolve, reject) => {
      // Tear down any existing picker.
      const prior = document.getElementById('ts-sheet-picker');
      if (prior) prior.remove();

      const overlay = document.createElement('div');
      overlay.id = 'ts-sheet-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';

      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner';
      inner.innerHTML =
        '<h3>Select sheet</h3>' +
        '<p>The file <code>' + escHtml(fileName) + '</code> has multiple sheets. Pick the one containing the source data:</p>' +
        '<div class="cmp-org-list" id="ts-sheet-picker-list"></div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="ts-sheet-picker-cancel">Cancel</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);

      const list = inner.querySelector('#ts-sheet-picker-list');
      sheets.forEach((s, i) => {
        const item = document.createElement('div');
        item.className = 'cmp-org-list-item';
        item.innerHTML =
          '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
          '<span class="cmp-org-counts">' + (s.rows ? s.rows.length : 0) + ' rows · ' +
          (s.headers ? s.headers.length : 0) + ' cols</span>';
        item.addEventListener('click', () => {
          overlay.remove();
          resolve(s);
        });
        list.appendChild(item);
      });

      const cancel = () => { overlay.remove(); reject(new Error('cancelled')); };
      inner.querySelector('#ts-sheet-picker-cancel').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
    });
  }

  // ─── File parsing ───
  // Always reads via XLSX directly so we can run the header-row detector on
  // the raw AOA. PickTrace implementation templates (Soto Bros, Wonderful, etc.)
  // have multi-row metadata above the real column header — readUploadedFile
  // would treat row 0 as the header and miss everything.
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
            return {
              name: n,
              headers: sliced.headers,
              rows: sliced.rows,
              headerRow: sliced.headerRow
            };
          }).filter(Boolean);
          if (!sheets.length) return reject(new Error('No sheets with usable headers.'));
          const finish = s => resolve({
            headers: s.headers, rows: s.rows, fileName: file.name,
            sheetName: s.name, headerRow: s.headerRow
          });
          if (sheets.length === 1) return finish(sheets[0]);
          pickSheet(sheets, file.name).then(finish).catch(reject);
        } catch (err) { reject(err); }
      };
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
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
      const dropWs = wb.Sheets[dropName];
      const raw = XLSX.utils.sheet_to_json(dropWs, { header: 1, defval: '' });
      if (raw.length >= 2) {
        const dh = raw[0].map(h => String(h || '').trim());
        dh.forEach((h, ci) => {
          if (!h) return;
          const k = norm(h);
          const set = dropdowns.get(k) || new Set();
          raw.slice(1).forEach(r => {
            const v = String(r[ci] != null ? r[ci] : '').trim();
            if (v) set.add(v);
          });
          if (set.size) dropdowns.set(k, set);
        });
      }
    }
    return { headers, dropdowns, fileName, rawBuffer: buf, sheetNames: wb.SheetNames };
  }

  // ─── Auto-mapping ───
  // Match each bulk header to a source-data column by:
  //   1. Exact case-insensitive header match
  //   2. Asterisk-stripped match ("Site*" → "site")
  //   3. Common alias map (display_name → Name*, plant_name → Crop & Variety*, etc.)
  const ALIASES = {
    'site*':              ['site','sites','site_name','site name','grower','farm'],
    'name*':              ['name','block','display_name','display name'],
    'alt id':             ['alt id','altid','code'],
    'location type*':     ['location type','type'],
    'crop & variety*':    ['crop & variety','crop and variety','plant_name','plant name','crop'],
    'planted at':         ['planted at','planted_at'],
    'acreage':            ['acreage','acres','acre','hectares'],
    'length':             ['length','length_meters'],
    'plant count':        ['plant count','plant_count','tree count','trees','plants','quantity'],
    'start date*':        ['start date','start_date'],
    'planted date':       ['planted date','planted_date'],
    'is archived':        ['is archived','is_archived','archived'],
    // Clone/Subvariety must NOT pick up "Variety". Specific aliases ensure
    // Phase 3 catches "Clone" first; STRICT_HEADERS below blocks Phase 4
    // from ever matching Variety against this column via substring.
    'clone/subvariety':   ['clone','subvariety','sub variety','sub_variety','sub-variety']
  };

  // Bulk headers where the substring fallback (Phase 4) is unsafe — a source
  // header could match by accident (e.g., "Variety" matching "Clone/Subvariety"
  // because "clonesubvariety".includes("variety") is true). For these we
  // accept ONLY exact-match + loose-alphanumeric + alias hits.
  const STRICT_HEADERS = new Set(['clone/subvariety']);
  // Loose normalization — strip all non-alphanumeric chars so "Plant/Seed Seller",
  // "Plant Seed Seller", and "plant-seed_seller" all collapse to "plantseedseller".
  const normLoose = h => norm(h).replace(/[^a-z0-9]/g, '');

  function autoMap(bulkHeaders, srcHeaders) {
    const srcIdx = {};         // norm(header) → idx (case+trim+lowercase only)
    const srcIdxLoose = {};    // normLoose(header) → idx (alphanumeric only)
    srcHeaders.forEach((h, i) => {
      const k = norm(h);
      const kLoose = normLoose(h);
      if (k && srcIdx[k] == null) srcIdx[k] = i;
      if (kLoose && srcIdxLoose[kLoose] == null) srcIdxLoose[kLoose] = i;
    });
    const out = {};
    bulkHeaders.forEach((bh, bi) => {
      const bn = norm(bh);
      const bnPlain = bn.replace(/\*$/, '').trim();
      const bnLoose = normLoose(bh);              // includes asterisk-strip via the regex
      const bnLoosePlain = normLoose(bh.replace(/\*$/, ''));
      // 1. Exact match (with or without asterisk).
      if (srcIdx[bn] != null) { out[bi] = srcIdx[bn]; return; }
      if (srcIdx[bnPlain] != null) { out[bi] = srcIdx[bnPlain]; return; }
      // 2. Loose alphanumeric match — catches "Plant/Seed Seller" ↔ "Plant Seed Seller",
      //    "Crop & Variety*" ↔ "Crop and Variety", "Row Spacing, in." ↔ "Row Spacing in".
      if (srcIdxLoose[bnLoose] != null) { out[bi] = srcIdxLoose[bnLoose]; return; }
      if (srcIdxLoose[bnLoosePlain] != null) { out[bi] = srcIdxLoose[bnLoosePlain]; return; }
      // 3. Aliases.
      const aliases = ALIASES[bn] || ALIASES[bnPlain];
      if (aliases) {
        for (const a of aliases) {
          if (srcIdx[a] != null) { out[bi] = srcIdx[a]; return; }
          const aLoose = normLoose(a);
          if (srcIdxLoose[aLoose] != null) { out[bi] = srcIdxLoose[aLoose]; return; }
        }
      }
      // Strict bulk headers don't fall through to substring matching — too
      // many false positives ("variety" matching "clonesubvariety", etc.).
      if (STRICT_HEADERS.has(bn) || STRICT_HEADERS.has(bnPlain)) {
        out[bi] = -1;
        return;
      }
      // 4. Substring fallback on the loose form. Skip *_id columns.
      for (const k of Object.keys(srcIdxLoose)) {
        if (!k || k.endsWith('id') && /id$/.test(k) && k.length <= bnLoosePlain.length) continue;
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

  // ─── Render ───
  function renderMapping() {
    if (!srcData || !tplData) return;
    const sec = $('ts-section-mapping');
    const tbl = $('ts-mapping-table');
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const srcOpts = ['<option value="-1">— (leave empty) —</option>']
      .concat(srcData.headers.map((h, i) => '<option value="' + i + '">' + escHtml(h) + '</option>'))
      .join('');
    let html = '<thead><tr><th>Template column</th><th>Source column</th><th>Sample value</th><th></th></tr></thead><tbody>';
    bulkHeaders.forEach((bh, bi) => {
      const required = /\*$/.test(bh);
      const sel = mapping[bi] != null ? mapping[bi] : -1;
      const sample = sel >= 0 && srcData.rows[0] ? String(srcData.rows[0][sel] || '').trim() : '';
      html += '<tr>' +
        '<td>' + (required ? '<b>' + escHtml(bh) + '</b>' : escHtml(bh)) + '</td>' +
        '<td><select class="ts-map-select input-field" data-bi="' + bi + '" style="min-width:240px;">' +
          srcOpts.replace('value="' + sel + '"', 'value="' + sel + '" selected') +
        '</select></td>' +
        '<td><span class="text-muted small">' + escHtml(sample.slice(0, 60)) + '</span></td>' +
        '<td><button class="btn btn-ghost btn-sm ts-map-clear" data-bi="' + bi + '" title="Clear all values in this column (header stays)">Clear</button></td>' +
        '</tr>';
    });
    html += '</tbody>';
    tbl.innerHTML = html;
    tbl.querySelectorAll('.ts-map-select').forEach(sel => {
      sel.addEventListener('change', e => {
        const bi = +e.target.dataset.bi;
        mapping[bi] = +e.target.value;
        rebuildFormattedRows();
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      });
    });
    tbl.querySelectorAll('.ts-map-clear').forEach(btn => {
      btn.addEventListener('click', e => {
        const bi = +e.target.dataset.bi;
        clearColumn(bi);
      });
    });
  }

  // Find a source-column index by checking exact then loose-alphanumeric
  // matches. `excludeContains` lets the caller reject specific source headers
  // (e.g., when looking for "Variety", exclude headers containing "clone" or
  // "subvariety" so we don't pick up Clone/Subvariety by mistake).
  function findSrcColByNames(names, excludeContains) {
    if (!srcData) return -1;
    const skip = (excludeContains || []).map(s => s.toLowerCase());
    const idx = {};
    const idxLoose = {};
    srcData.headers.forEach((h, i) => {
      const headerLower = String(h || '').toLowerCase();
      if (skip.some(s => headerLower.includes(s))) return; // never consider excluded headers
      const k = norm(h);
      const kLoose = normLoose(h);
      if (k && idx[k] == null) idx[k] = i;
      if (kLoose && idxLoose[kLoose] == null) idxLoose[kLoose] = i;
    });
    for (const n of names) {
      const k = norm(n);
      if (idx[k] != null) return idx[k];
      const kLoose = normLoose(n);
      if (idxLoose[kLoose] != null) return idxLoose[kLoose];
    }
    return -1;
  }

  function buildFormattedRowsRaw() {
    if (!srcData || !tplData) return null;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const isArchivedIdx = bulkHeaders.indexOf('Is Archived');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    // Detect a separate Crop + Variety pair in the source. If both exist we
    // combine them as "Crop-Variety" for the Crop & Variety* cell. Clone /
    // Subvariety is a different field — exclude it from variety detection.
    const cropSrcIdx = findSrcColByNames(['crop','crop type','crop name','plant_name','plant name']);
    const varSrcIdx = findSrcColByNames(['variety','varietal'], ['clone','subvariety','sub variety','sub-variety']);
    const synthesizeCv = cvIdx >= 0 && cropSrcIdx >= 0 && varSrcIdx >= 0 && cropSrcIdx !== varSrcIdx;

    return srcData.rows.map(srcRow => {
      const out = new Array(bulkHeaders.length).fill('');
      bulkHeaders.forEach((bh, bi) => {
        const si = mapping[bi];
        if (si != null && si >= 0 && srcRow[si] != null) {
          const v = String(srcRow[si]).trim();
          if (v) out[bi] = v;
        }
      });
      if (synthesizeCv) {
        const c = String(srcRow[cropSrcIdx] || '').trim();
        const v = String(srcRow[varSrcIdx] || '').trim();
        if (c && v) out[cvIdx] = c + '-' + v;
        else if (c) out[cvIdx] = c;
        else if (v) out[cvIdx] = v;
      }
      if (target === 'update' && isArchivedIdx >= 0 && !out[isArchivedIdx]) {
        out[isArchivedIdx] = 'FALSE';
      }
      return out;
    });
  }

  // Rebuild formatted rows from source + mapping, then re-apply any manual
  // fills the user has applied. Called when the inputs change (mapping edited,
  // target switched, source/template re-uploaded). NOT called by renderPreview
  // — that function just displays the canonical formattedRows state.
  function rebuildFormattedRows() {
    formattedRows = buildFormattedRowsRaw();
    if (!formattedRows) return;
    Object.entries(manualFills).forEach(([idx, val]) => {
      const i = +idx;
      formattedRows.forEach(r => { if (!r[i]) r[i] = val; });
    });
  }

  function renderPreview() {
    if (!srcData || !tplData || !formattedRows) return;
    const sec = $('ts-section-preview');
    const tbl = $('ts-preview-table');
    sec.style.display = '';
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    let html = '<thead><tr>';
    // First column is the row-action column (X to remove the row).
    html += '<th style="width:28px;text-align:center;color:#9ca3af;">&nbsp;</th>';
    bulkHeaders.forEach(h => {
      const req = /\*$/.test(h);
      html += '<th' + (req ? ' style="color:#dc2626;"' : '') + '>' + escHtml(h) + '</th>';
    });
    html += '</tr></thead><tbody>';
    const visibleRowIdxs = formattedRows
      .map((_, i) => i)
      .filter(i => !removedRows.has(i));
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    const limit = Math.min(visibleRowIdxs.length, 50);
    for (let r = 0; r < limit; r++) {
      const ri = visibleRowIdxs[r];
      const row = formattedRows[ri];
      html += '<tr>';
      html += '<td style="text-align:center;padding:0;">' +
        '<button class="ts-row-remove" data-ri="' + ri + '" title="Remove this row from preview + export" ' +
        'style="all:unset;cursor:pointer;color:#9ca3af;font-size:14px;line-height:1;padding:2px 6px;">&times;</button>' +
        '</td>';
      bulkHeaders.forEach((h, i) => {
        const req = /\*$/.test(h);
        const empty = !row[i];
        let cellStyle = '';
        if (req && empty) cellStyle = ' style="background:#fee2e2;color:#7f1d1d;"';
        else if (i === startDateIdx && row[i] && isDateBeforeToday(row[i])) {
          cellStyle = ' style="background:#fef3c7;color:#7c2d12;" title="Start Date must be today (' + todayYMD() + ') or later"';
        }
        html += '<td' + cellStyle + '>' + escHtml(row[i]) + '</td>';
      });
      html += '</tr>';
    }
    html += '</tbody>';
    tbl.innerHTML = html;
    const hint = sec.querySelector('.cmp-sites-hint');
    let hintText = 'Empty required cells are highlighted in red. Click <b>×</b> at the start of a row to drop it from the preview + export.';
    if (visibleRowIdxs.length > limit) {
      hintText = 'Showing first ' + limit + ' of ' + visibleRowIdxs.length + ' rows. ' + hintText;
    }
    if (removedRows.size) {
      hintText += ' &nbsp; <button class="btn btn-ghost btn-sm" id="ts-restore-rows">' +
        'Restore ' + removedRows.size + ' removed row' + (removedRows.size === 1 ? '' : 's') + '</button>';
    }
    hint.innerHTML = hintText;
    tbl.querySelectorAll('.ts-row-remove').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        const ri = +e.currentTarget.dataset.ri;
        removedRows.add(ri);
        renderPreview();
        renderRequired();
        renderSitesAndCvToCreate();
        updateSummary();
      });
    });
    const restoreBtn = $('ts-restore-rows');
    if (restoreBtn) restoreBtn.addEventListener('click', () => {
      removedRows.clear();
      renderPreview();
      renderRequired();
      renderSitesAndCvToCreate();
      updateSummary();
    });
    updateExportButton();
  }

  function renderSitesAndCvToCreate() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx   = bulkHeaders.indexOf('Crop & Variety*');
    const tplSites = tplData.dropdowns.get('site') || new Set();
    const tplCvs   = tplData.dropdowns.get('crop & variety') || new Set();
    const tplSitesMap = buildCaseMap(tplSites);
    const tplCvsMap   = buildCaseMap(tplCvs);

    const sites = new Map(), cvs = new Map();
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      const s = r[siteIdx];
      if (s && !inDropdown(tplSitesMap, s)) sites.set(s, (sites.get(s) || 0) + 1);
      const v = r[cvIdx];
      if (v && !inDropdown(tplCvsMap, v))   cvs.set(v, (cvs.get(v) || 0) + 1);
    });

    const renderList = (sectionId, tableId, titleId, label, m) => {
      const sec = $(sectionId);
      if (!m.size) { sec.style.display = 'none'; return; }
      sec.style.display = '';
      $(titleId).textContent = label + ' (' + m.size + ')';
      const sorted = [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
      let html = '<thead><tr><th>Value</th><th>Row count</th></tr></thead><tbody>';
      sorted.forEach(([val, cnt]) => {
        html += '<tr><td>' + escHtml(val) + '</td><td>' + cnt + '</td></tr>';
      });
      html += '</tbody>';
      $(tableId).innerHTML = html;
    };
    renderList('ts-section-sites-create', 'ts-sites-create-table', 'ts-sites-create-title', 'Sites to Create in PickTrace', sites);
    renderList('ts-section-cv-create',    'ts-cv-create-table',    'ts-cv-create-title',    'Crops & Varieties to Create in PickTrace', cvs);
    return { sitesCount: sites.size, cvsCount: cvs.size };
  }

  function renderRequired() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const sec = $('ts-section-required');
    sec.style.display = '';
    const requiredCols = bulkHeaders
      .map((h, i) => ({ header: h, idx: i, req: /\*$/.test(h) }))
      .filter(x => x.req);

    // Split into two buckets: rows that still need attention vs rows that
    // are already filled. Filled rows stay visible (with a "Change value"
    // affordance) so the user can re-pick or clear at any time.
    const needsAction = [];
    const filled = [];
    requiredCols.forEach(col => {
      let empty = 0;
      const sample = new Set();
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri)) return;
        if (!r[col.idx]) empty++;
        else if (sample.size < 1 && r[col.idx]) sample.add(r[col.idx]);
      });
      const fillVal = manualFills[col.idx] || (sample.size ? [...sample][0] : '');
      const entry = { ...col, empty, fillVal };
      if (empty > 0) needsAction.push(entry); else filled.push(entry);
    });

    const renderRow = (col) => {
      const dropKey = norm(col.header).replace(/\*$/, '');
      const opts = tplData.dropdowns.get(dropKey)
        || tplData.dropdowns.get(dropKey + '*')
        || tplData.dropdowns.get(norm(col.header))
        || new Set();
      const optsHtml = opts.size
        ? '<option value="">— pick —</option>' +
          [...opts].sort().map(o => '<option' + (o === col.fillVal ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no template dropdown — use manual override)</option>';
      const status = col.empty > 0
        ? '<span style="color:#dc2626;font-weight:600;">⚠ ' + col.empty + ' empty</span>'
        : '<span style="color:#15803d;">✓ filled' + (col.fillVal ? ' (e.g. ' + escHtml(String(col.fillVal).slice(0, 30)) + ')' : '') + '</span>';
      const rowStyle = col.empty > 0 ? '' : ' style="background:var(--bg-sunken);"';
      // Date-aware quick fill — Start Date* gets a "Today" button that
      // fills every visible row with today's YYYY-MM-DD value.
      const isStartDate = /^start\s*date/i.test(col.header.replace(/\*$/, ''));
      const todayBtn = isStartDate
        ? '<button class="btn btn-warning btn-sm ts-req-today" data-idx="' + col.idx + '" title="Fill all rows with today\'s date (' + todayYMD() + ')">Today</button>'
        : '';
      return '<tr' + rowStyle + '>' +
        '<td><b>' + escHtml(col.header) + '</b></td>' +
        '<td>' + status + '</td>' +
        '<td><select class="ts-req-pick input-field" data-idx="' + col.idx + '" style="min-width:200px;">' + optsHtml + '</select></td>' +
        '<td><input type="text" class="ts-req-override input-field" data-idx="' + col.idx + '" placeholder="Manual override" style="width:200px;"></td>' +
        '<td>' + todayBtn + '</td>' +
        '<td><button class="btn btn-primary btn-sm ts-req-apply" data-idx="' + col.idx + '">Apply</button></td>' +
        '<td><button class="btn btn-ghost btn-sm ts-req-clear" data-idx="' + col.idx + '" title="Clear all values in this column (header stays)">Clear</button></td>' +
        '</tr>';
    };

    let html = '<thead><tr><th>Column</th><th>Status</th><th>Pick from dropdown</th><th>Manual override</th><th>Quick fill</th><th></th><th></th></tr></thead>';
    if (needsAction.length) {
      html += '<tbody><tr><td colspan="7" style="background:#fee2e2;color:#7f1d1d;font-weight:600;padding:6px 10px;">' +
        '⚠ Action needed (' + needsAction.length + ')</td></tr>' +
        needsAction.map(renderRow).join('') + '</tbody>';
    }
    if (filled.length) {
      html += '<tbody><tr><td colspan="7" style="background:#dcfce7;color:#14532d;font-weight:600;padding:6px 10px;">' +
        '✓ Already filled — change or clear if needed (' + filled.length + ')</td></tr>' +
        filled.map(renderRow).join('') + '</tbody>';
    }
    $('ts-required-table').innerHTML = html;
    wireRequiredHandlers();
  }

  // ─── Apply / Clear handlers (extracted so renderRequired stays readable) ───
  function applyFill(idx, val, overwrite) {
    if (!val) return 0;
    let filled = 0;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (overwrite || !r[idx]) { r[idx] = val; filled++; }
    });
    manualFills[idx] = val;
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
    return filled;
  }
  function clearColumn(idx) {
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      r[idx] = '';
    });
    delete manualFills[idx];
    renderPreview();
    renderRequired();
    renderSitesAndCvToCreate();
    updateSummary();
  }
  function wireRequiredHandlers() {
    const tbl = $('ts-required-table');
    if (!tbl) return;
    tbl.querySelectorAll('.ts-req-apply').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        const tr = e.target.closest('tr');
        const pick = tr.querySelector('.ts-req-pick').value.trim();
        const ovr  = tr.querySelector('.ts-req-override').value.trim();
        const val = ovr || pick;
        if (!val) { alert('Pick a value or type a manual override first.'); return; }
        // Apply OVERWRITES — that way the user can change their mind and
        // re-apply a different value to the same column.
        applyFill(idx, val, true);
      });
    });
    tbl.querySelectorAll('.ts-req-clear').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        clearColumn(idx);
      });
    });
    tbl.querySelectorAll('.ts-req-today').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = +e.target.dataset.idx;
        applyFill(idx, todayYMD(), true);
      });
    });
    // Real-time auto-apply ONLY for the dropdown picker (instant, intentional).
    // The text override is hands-off until the user clicks Apply — typing no
    // longer auto-fills mid-keystroke.
    tbl.querySelectorAll('.ts-req-pick').forEach(sel => {
      sel.addEventListener('change', e => {
        const idx = +e.target.dataset.idx;
        const v = e.target.value.trim();
        if (v) applyFill(idx, v, true);
      });
    });
  }

  function updateSummary() {
    if (!srcData || !tplData) { $('ts-summary').style.display = 'none'; return; }
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const reqIdxs = bulkHeaders.map((h, i) => ({ h, i, req: /\*$/.test(h) })).filter(x => x.req);
    let emptyReq = 0;
    if (formattedRows) reqIdxs.forEach(({ i }) => {
      formattedRows.forEach((r, ri) => { if (!removedRows.has(ri) && !r[i]) emptyReq++; });
    });
    const sites = renderSitesAndCvToCreate() || { sitesCount: 0, cvsCount: 0 };
    const visibleCount = formattedRows ? formattedRows.filter((_, i) => !removedRows.has(i)).length : 0;
    const totalCount = formattedRows ? formattedRows.length : 0;
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    let pastDates = 0;
    if (formattedRows && startDateIdx >= 0) {
      formattedRows.forEach((r, ri) => {
        if (removedRows.has(ri)) return;
        if (isDateBeforeToday(r[startDateIdx])) pastDates++;
      });
    }
    const sum = $('ts-summary');
    sum.style.display = '';
    sum.innerHTML =
      '<div class="cmp-stat"><b>' + visibleCount + '</b> rows' +
        (removedRows.size ? ' <span class="text-muted small">(' + removedRows.size + ' removed of ' + totalCount + ')</span>' : '') +
        '</div>' +
      '<div class="cmp-stat"><b>' + bulkHeaders.length + '</b> template columns</div>' +
      (emptyReq ? '<div class="cmp-stat cmp-warn"><b>' + emptyReq + '</b> empty required cells</div>' : '') +
      (sites.sitesCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.sitesCount + '</b> sites to create</div>' : '') +
      (sites.cvsCount ? '<div class="cmp-stat cmp-warn"><b>' + sites.cvsCount + '</b> C&V to create</div>' : '') +
      (pastDates ? '<div class="cmp-stat cmp-warn"><b>' + pastDates + '</b> Start Dates before today</div>' : '');
  }

  // Build a case-insensitive index of a dropdown Set so values like
  // "Beckstoffer Vineyards" match the template's "BECKSTOFFER VINEYARDS".
  // Returns Map<upperTrimmed, canonicalValue>. Use .has() / .get() instead of
  // the raw Set's case-sensitive .has().
  function buildCaseMap(set) {
    const m = new Map();
    if (!set) return m;
    set.forEach(v => {
      const k = String(v).toUpperCase().trim();
      if (k && !m.has(k)) m.set(k, v);
    });
    return m;
  }
  function inDropdown(map, value) {
    if (!value || !map.size) return false;
    return map.has(String(value).toUpperCase().trim());
  }

  function getExportBlockers() {
    if (!srcData || !tplData || !formattedRows) return null;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const reqIdxs = bulkHeaders
      .map((h, i) => /\*$/.test(h) ? i : -1)
      .filter(i => i >= 0);
    let emptyReq = 0;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      reqIdxs.forEach(i => { if (!r[i]) emptyReq++; });
    });
    const tplSites = tplData.dropdowns.get('site') || new Set();
    const tplCvs   = tplData.dropdowns.get('crop & variety') || new Set();
    const tplSitesMap = buildCaseMap(tplSites);
    const tplCvsMap   = buildCaseMap(tplCvs);
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx = bulkHeaders.indexOf('Crop & Variety*');
    const startDateIdx = bulkHeaders.indexOf('Start Date*');
    const checkSites = tplSites.size > 0;
    const checkCvs   = tplCvs.size > 0;
    const unknownSiteVals = new Set();
    const unknownCvVals = new Set();
    let pastStartDates = 0;
    formattedRows.forEach((r, ri) => {
      if (removedRows.has(ri)) return;
      if (checkSites && r[siteIdx] && !inDropdown(tplSitesMap, r[siteIdx])) unknownSiteVals.add(r[siteIdx]);
      if (checkCvs   && r[cvIdx]   && !inDropdown(tplCvsMap,   r[cvIdx]))   unknownCvVals.add(r[cvIdx]);
      if (startDateIdx >= 0 && isDateBeforeToday(r[startDateIdx])) pastStartDates++;
    });
    return { emptyReq, unknownSites: [...unknownSiteVals], unknownCvs: [...unknownCvVals], pastStartDates };
  }

  function updateExportButton() {
    // Button stays ENABLED whenever there's data — clicking will surface a
    // confirmation when there are blockers, instead of leaving the user stuck.
    const btn = $('ts-export');
    btn.disabled = !(srcData && tplData && formattedRows && formattedRows.length);
    const b = getExportBlockers();
    if (!b) { btn.title = ''; return; }
    const reasons = [];
    if (b.emptyReq > 0) reasons.push(b.emptyReq + ' empty required cells');
    if (b.unknownSites.length) reasons.push(b.unknownSites.length + ' sites not in template dropdown');
    if (b.unknownCvs.length) reasons.push(b.unknownCvs.length + ' Crop & Variety values not in template dropdown');
    if (b.pastStartDates > 0) reasons.push(b.pastStartDates + ' Start Date values are before today');
    btn.title = reasons.length
      ? 'Click to review issues before exporting: ' + reasons.join(', ') + '.'
      : 'Ready to export.';
  }

  // ─── Custom export-confirm modal ───
  // Replaces the native window.confirm() which can't be styled and renders
  // raw blocker text. Reuses the .cmp-export-modal classes already defined
  // in styles.css so the look matches Block Compare's gate dialog.
  function showExportConfirm(b) {
    return new Promise(resolve => {
      const prior = document.getElementById('ts-export-confirm');
      if (prior) prior.remove();

      const overlay = document.createElement('div');
      overlay.id = 'ts-export-confirm';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';

      const sections = [];
      if (b.emptyReq > 0) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.emptyReq + ' empty required cells</div>' +
          '<div class="ts-confirm-issue-body">PickTrace will reject rows missing values for required (*) columns. Fill them via the Empty Required Columns panel.</div>' +
          '</div>'
        );
      }
      if (b.unknownSites.length) {
        const sample = b.unknownSites.slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = b.unknownSites.length > 8 ? '<li class="ts-confirm-more">… and ' + (b.unknownSites.length - 8) + ' more</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unknownSites.length + ' Site value(s) not in your template dropdown</div>' +
          '<ul class="ts-confirm-list">' + sample + more + '</ul>' +
          '<div class="ts-confirm-issue-body">Create these Sites in PickTrace, re-download your personalized template, and re-upload here.</div>' +
          '</div>'
        );
      }
      if (b.unknownCvs.length) {
        const sample = b.unknownCvs.slice(0, 8).map(v => '<li>' + escHtml(v) + '</li>').join('');
        const more = b.unknownCvs.length > 8 ? '<li class="ts-confirm-more">… and ' + (b.unknownCvs.length - 8) + ' more</li>' : '';
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.unknownCvs.length + ' Crop &amp; Variety value(s) not in your template dropdown</div>' +
          '<ul class="ts-confirm-list">' + sample + more + '</ul>' +
          '<div class="ts-confirm-issue-body">Create these in PickTrace and re-download the template.</div>' +
          '</div>'
        );
      }
      if (b.pastStartDates > 0) {
        sections.push(
          '<div class="ts-confirm-issue">' +
          '<div class="ts-confirm-issue-head">⚠ ' + b.pastStartDates + ' Start Date value(s) are before today (' + todayYMD() + ')</div>' +
          '<div class="ts-confirm-issue-body">PickTrace requires Start Date to be today or later. Use the <b>Today</b> button on the Start Date row.</div>' +
          '</div>'
        );
      }

      overlay.innerHTML =
        '<div class="cmp-export-modal-inner ts-confirm-modal">' +
          '<h3>Export with these issues?</h3>' +
          '<div class="ts-confirm-body">' + sections.join('') + '</div>' +
          '<div class="cmp-export-modal-actions">' +
            '<button class="btn btn-warning" id="ts-conf-ok">Export anyway</button>' +
            '<button class="btn btn-ghost" id="ts-conf-cancel">Cancel &amp; fix</button>' +
          '</div>' +
        '</div>';
      document.body.appendChild(overlay);

      const close = result => { overlay.remove(); resolve(result); };
      overlay.querySelector('#ts-conf-ok').addEventListener('click', () => close(true));
      overlay.querySelector('#ts-conf-cancel').addEventListener('click', () => close(false));
      overlay.addEventListener('click', e => { if (e.target === overlay) close(false); });
    });
  }

  // ─── Run / Export ───
  function runFormat() {
    if (!srcData || !tplData) { alert('Upload source data + a personalized bulk template first.'); return; }
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    mapping = autoMap(bulkHeaders, srcData.headers);
    // Reset manual fills + removed rows on a fresh format — they're tied
    // to the prior mapping/target context.
    manualFills = {};
    removedRows = new Set();
    rebuildFormattedRows();
    renderMapping();
    renderPreview();
    renderRequired();
    updateSummary();
    $('ts-empty').style.display = 'none';
  }

  function doExport() {
    if (!formattedRows || !tplData) return;
    const bulkHeaders = target === 'update' ? UPDATE_HEADERS : CREATE_HEADERS;
    const siteIdx = bulkHeaders.indexOf('Site*');
    const cvIdx   = bulkHeaders.indexOf('Crop & Variety*');
    const tplSitesMap = buildCaseMap(tplData.dropdowns.get('site'));
    const tplCvsMap   = buildCaseMap(tplData.dropdowns.get('crop & variety'));
    // Drop removed rows and canonicalize Site / Crop & Variety casing so the
    // values match the template's dropdown literally (PickTrace is strict).
    const exportRows = formattedRows
      .filter((_, ri) => !removedRows.has(ri))
      .map(row => {
        const r = row.slice();
        if (siteIdx >= 0 && r[siteIdx]) {
          const k = String(r[siteIdx]).toUpperCase().trim();
          if (tplSitesMap.has(k)) r[siteIdx] = tplSitesMap.get(k);
        }
        if (cvIdx >= 0 && r[cvIdx]) {
          const k = String(r[cvIdx]).toUpperCase().trim();
          if (tplCvsMap.has(k)) r[cvIdx] = tplCvsMap.get(k);
        }
        return r;
      });

    // Open the original template workbook in place so we keep:
    //   – every cell's font / fill / border / number format
    //   – column widths, frozen panes, sheet order
    //   – the DROP-DOWN INPUTS sheet (verbatim, so dropdown lookups still work)
    //   – any data validations (xlsx-js-style preserves them on roundtrip)
    // …then overwrite the DATA ENTRY rows with the user's formatted data.
    const wb = XLSX.read(new Uint8Array(tplData.rawBuffer), {
      type: 'array', cellStyles: true, cellDates: true, sheetStubs: true
    });
    const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[dataName];
    if (!ws) { alert('Template missing DATA ENTRY sheet — cannot export.'); return; }

    // Wipe existing data rows (keep the header row at row 0). We delete only
    // the cell objects, leaving sheet-level properties (!cols, !merges,
    // !dataValidation, !autofilter, etc.) intact.
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: bulkHeaders.length - 1 } };
    for (let r = 1; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (ws[ref]) delete ws[ref];
      }
    }

    // Write the formatted rows starting at row 1 (right under the header).
    const colCount = Math.max(range.e.c + 1, bulkHeaders.length);
    exportRows.forEach((row, ri) => {
      for (let c = 0; c < colCount; c++) {
        const val = row[c];
        if (val == null || val === '') continue;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c });
        // Preserve numbers as numbers (so PickTrace's numeric validations work),
        // booleans for Is Archived, everything else as text.
        if (typeof val === 'number') ws[ref] = { v: val, t: 'n' };
        else if (typeof val === 'string' && /^-?\d+(\.\d+)?$/.test(val) &&
                 (bulkHeaders[c] === 'Acreage' || bulkHeaders[c] === 'Plant Count' ||
                  bulkHeaders[c] === 'Length' || bulkHeaders[c] === 'Stand Count' ||
                  bulkHeaders[c] === 'Row/Bed Count' || bulkHeaders[c] === 'Post Count')) {
          ws[ref] = { v: parseFloat(val), t: 'n' };
        } else {
          ws[ref] = { v: String(val), t: 's' };
        }
      }
    });

    // Update !ref to span the new range. Header at row 0; data rows 1..N.
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(0, exportRows.length), c: colCount - 1 }
    });

    // Filename: original template name + " — filled" before the extension.
    // Keeps the user from having to overwrite the source template.
    const origName = (tplData.fileName || 'bulk-template.xlsx').trim();
    const dotIdx = origName.lastIndexOf('.');
    const base = dotIdx > 0 ? origName.substring(0, dotIdx) : origName;
    const ext = dotIdx > 0 ? origName.substring(dotIdx) : '.xlsx';
    const exportName = base + ' — filled' + ext;

    XLSX.writeFile(wb, exportName, { cellStyles: true });
  }

  // ─── Load handlers ───
  function handleSrcFile(file) {
    readSrcFile(file).then(data => {
      srcData = data;
      $('ts-src-name').textContent = file.name + ' [' + data.sheetName + ']';
      const hrowNote = data.headerRow > 0 ? ' (skipped ' + data.headerRow + ' metadata row' + (data.headerRow === 1 ? '' : 's') + ' above)' : '';
      $('ts-src-meta').textContent = data.rows.length + ' rows · ' + data.headers.length + ' columns · sheet: ' + data.sheetName + hrowNote;
      formattedRows = null;
      $('ts-run').disabled = !(srcData && tplData);
      // Auto-run if both already in.
      if (srcData && tplData) runFormat();
    }).catch(err => {
      if (err && err.message === 'cancelled') return; // user dismissed the picker — quiet
      alert('Failed to read file: ' + (err && err.message ? err.message : err));
    });
  }
  function handleTplFile(file) {
    const r = new FileReader();
    r.onload = e => {
      const parsed = parseTemplate(e.target.result, file.name);
      if (!parsed || !parsed.headers.length) {
        alert('Could not find DATA ENTRY sheet in this template.');
        return;
      }
      tplData = parsed;
      $('ts-tpl-name').textContent = file.name;
      $('ts-tpl-meta').textContent =
        parsed.headers.length + ' columns · ' +
        (parsed.dropdowns.get('site') ? parsed.dropdowns.get('site').size : 0) + ' sites · ' +
        (parsed.dropdowns.get('crop & variety') ? parsed.dropdowns.get('crop & variety').size : 0) + ' C&V values';
      // Detect target by header count if user hasn't picked.
      const hasArch = parsed.headers.some(h => norm(h) === 'is archived');
      const radioVal = document.querySelector('input[name="ts-target"]:checked').value;
      if (hasArch && radioVal === 'create') {
        document.querySelector('input[name="ts-target"][value="update"]').checked = true;
        target = 'update';
      } else if (!hasArch && radioVal === 'update') {
        document.querySelector('input[name="ts-target"][value="create"]').checked = true;
        target = 'create';
      }
      formattedRows = null;
      $('ts-run').disabled = !(srcData && tplData);
      if (srcData && tplData) runFormat();
    };
    r.readAsArrayBuffer(file);
  }

  // ─── Org-Data import — reuses Block Compare's aggregator if available ───
  function importFromOrgData() {
    if (typeof aggregateOrgData !== 'function' || typeof getAllOrgNames !== 'function') {
      alert('Org Data store unavailable. Reload the page.');
      return;
    }
    const names = getAllOrgNames();
    if (!names.length) { alert('No saved orgs found.'); return; }
    const pick = prompt('Type the org name to import:\n\n' + names.join('\n'));
    if (!pick) return;
    const found = names.find(n => n.toLowerCase() === pick.trim().toLowerCase());
    if (!found) { alert('Org "' + pick + '" not found.'); return; }
    const aggregated = aggregateOrgData(found);
    if (!aggregated || !aggregated.rows.length) { alert('No rows for "' + found + '".'); return; }
    srcData = {
      headers: aggregated.headers,
      rows: aggregated.rows,
      fileName: '(Org Data: ' + found + ')',
      sheetName: '(synthesized)'
    };
    $('ts-src-name').textContent = '(Org Data: ' + found + ')';
    $('ts-src-meta').textContent = aggregated.rows.length + ' rows · ' + aggregated.headers.length + ' columns';
    formattedRows = null;
    $('ts-run').disabled = !(srcData && tplData);
    if (srcData && tplData) runFormat();
  }

  // ─── Reset ───
  function reset() {
    srcData = null; tplData = null; formattedRows = null; mapping = {};
    manualFills = {}; removedRows = new Set();
    $('ts-src-name').textContent = 'No file selected';
    $('ts-tpl-name').textContent = 'No file selected';
    $('ts-src-meta').textContent = '';
    $('ts-tpl-meta').textContent = '';
    $('ts-src-file').value = ''; $('ts-tpl-file').value = '';
    ['ts-section-mapping','ts-section-sites-create','ts-section-cv-create','ts-section-required','ts-section-preview']
      .forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    $('ts-summary').style.display = 'none';
    $('ts-empty').style.display = '';
    $('ts-run').disabled = true;
    $('ts-export').disabled = true;
  }

  // ─── Init ───
  function init() {
    if (initialized) return;
    initialized = true;
    attachModeSwitcher();
    $('ts-src-file').addEventListener('change', e => {
      if (e.target.files[0]) handleSrcFile(e.target.files[0]);
      e.target.value = '';
    });
    $('ts-tpl-file').addEventListener('change', e => {
      if (e.target.files[0]) handleTplFile(e.target.files[0]);
      e.target.value = '';
    });
    $('ts-src-from-org').addEventListener('click', importFromOrgData);
    document.querySelectorAll('input[name="ts-target"]').forEach(r => {
      r.addEventListener('change', e => {
        target = e.target.value;
        if (srcData && tplData) runFormat();
      });
    });
    $('ts-run').addEventListener('click', runFormat);
    $('ts-export').addEventListener('click', async () => {
      const b = getExportBlockers();
      if (b && (b.emptyReq > 0 || b.unknownSites.length || b.unknownCvs.length || b.pastStartDates > 0)) {
        const proceed = await showExportConfirm(b);
        if (!proceed) return;
      }
      doExport();
    });
    $('ts-reset').addEventListener('click', reset);
    $('ts-sites-copy').addEventListener('click', () => {
      const tbl = $('ts-sites-create-table');
      const rows = [...tbl.querySelectorAll('tbody tr')]
        .map(tr => [...tr.children].map(td => td.textContent).join('\t'));
      if (!rows.length) return;
      navigator.clipboard.writeText('Site\tRow count\n' + rows.join('\n'))
        .then(() => alert('Copied ' + rows.length + ' sites.'), () => alert('Copy failed.'));
    });
    $('ts-cv-copy').addEventListener('click', () => {
      const tbl = $('ts-cv-create-table');
      const rows = [...tbl.querySelectorAll('tbody tr')]
        .map(tr => [...tr.children].map(td => td.textContent).join('\t'));
      if (!rows.length) return;
      navigator.clipboard.writeText('Crop & Variety\tRow count\n' + rows.join('\n'))
        .then(() => alert('Copied ' + rows.length + ' values.'), () => alert('Copy failed.'));
    });
  }

  // Init when the parent page is opened.
  window.tsInit = init;
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
