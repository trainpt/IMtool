// ── Column Editor ─────────────────────────────────────────────────────────────
(function() {
  // Target file data
  let tgtHeaders = [], tgtRows = [], tgtOriginal = [];
  let ceSelCol = -1;
  let ceSelRows = new Set();
  let ceChangedRows = new Set();  // rows that were modified by Apply
  let ceCopyCols = new Set();     // columns selected for copy

  const ceEl = id => document.getElementById(id);

  // ── CSV parser ──
  function ceParseCSV(text) {
    const headers = [], rows = [];
    const lines = [];
    let cur = '', inQ = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"' && text[i+1] === '"') { cur += '"'; i++; }
        else if (ch === '"') inQ = false;
        else cur += ch;
      } else {
        if (ch === '"') inQ = true;
        else if (ch === ',') { lines.push(cur); cur = ''; }
        else if (ch === '\n' || (ch === '\r' && text[i+1] === '\n')) {
          lines.push(cur); cur = '';
          if (ch === '\r') i++;
          if (lines.length > 0) {
            if (!headers.length) headers.push(...lines.splice(0));
            else rows.push(lines.splice(0));
          }
        } else cur += ch;
      }
    }
    if (cur || lines.length) { lines.push(cur); if (!headers.length) headers.push(...lines); else rows.push(lines.splice(0)); }
    const max = Math.max(headers.length, ...rows.map(r => r.length));
    while (headers.length < max) headers.push('Col ' + headers.length);
    rows.forEach(r => { while (r.length < max) r.push(''); });
    return { headers, rows };
  }

  // ── File loading (CSV + Excel) ──
  ceEl('ce-file-target').addEventListener('change', e => {
    if (!e.target.files[0]) return;
    const file = e.target.files[0];
    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'csv') {
      const reader = new FileReader();
      reader.onload = ev => {
        const parsed = ceParseCSV(ev.target.result);
        ceLoadData(parsed.headers, parsed.rows, file.name);
      };
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = ev => {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        if (wb.SheetNames.length === 1) {
          ceLoadExcelSheet(wb, wb.SheetNames[0], file.name);
        } else {
          ceShowSheetPicker(wb, file.name);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  });

  function ceLoadExcelSheet(wb, sheetName, fileName) {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const filtered = data.filter(r => r.some(c => String(c).trim() !== ''));
    if (filtered.length < 2) { alert('Sheet "' + sheetName + '" has no data.'); return; }
    ceLoadData(filtered[0].map(String), filtered.slice(1).map(r => r.map(String)), fileName + ' [' + sheetName + ']');
  }

  function ceLoadData(headers, rows, label) {
    tgtHeaders = headers; tgtRows = rows;
    ceEl('ce-target-name').textContent = label + ' (' + tgtRows.length + ' rows)';
    ceSelRows = new Set(tgtRows.map((_, i) => i));
    ceSelCol = -1;
    ceShowControls();
  }

  function ceShowSheetPicker(wb, fileName) {
    let picker = document.getElementById('ce-sheet-picker');
    if (!picker) { picker = document.createElement('div'); picker.id = 'ce-sheet-picker'; picker.className = 'modal-overlay show'; document.body.appendChild(picker); }
    let html = '<div class="modal"><h3>Select a Sheet</h3>';
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const count = Math.max(0, data.filter(r => r.some(c => String(c).trim() !== '')).length - 1);
      const eName = typeof escHtml === 'function' ? escHtml(name) : name;
      html += '<button class="btn btn-ghost" style="width:100%;text-align:left;margin-bottom:4px;justify-content:space-between;" data-sheet="' + name.replace(/"/g, '&quot;') + '">' +
        '<span>' + eName + '</span><span class="text-muted small">' + count + ' rows</span></button>';
    });
    html += '<div class="modal-actions"><button class="btn btn-ghost" id="ce-sheet-picker-cancel">Cancel</button></div></div>';
    picker.innerHTML = html; picker.style.display = 'flex';
    picker.querySelector('#ce-sheet-picker-cancel').addEventListener('click', () => { picker.style.display = 'none'; });
    picker.querySelectorAll('[data-sheet]').forEach(btn => {
      btn.addEventListener('click', () => { picker.style.display = 'none'; ceLoadExcelSheet(wb, btn.dataset.sheet, fileName); });
    });
  }

  // ── Mode toggle ──
  ceEl('ce-mode').addEventListener('change', () => {
    const isAuto = ceEl('ce-mode').value === 'auto';
    ceEl('ce-source-group').style.display = isAuto ? '' : 'none';
    ceEl('ce-separator-group').style.display = isAuto ? '' : 'none';
    ceEl('ce-manual-group').style.display = isAuto ? 'none' : '';
  });

  // ── Populate dropdowns ──
  function ceUpdateDropdowns() {
    const colSel = ceEl('ce-col-select');
    const srcSel = ceEl('ce-source-col');
    const prev = colSel.value;
    colSel.innerHTML = '<option value="">-- select --</option>';
    srcSel.innerHTML = '<option value="">-- select --</option>';
    tgtHeaders.forEach((h, i) => {
      const label = h || ('Col ' + i);
      colSel.innerHTML += `<option value="${i}">${label}</option>`;
      srcSel.innerHTML += `<option value="${i}">${label}</option>`;
    });
    colSel.value = prev;
  }

  function ceShowControls() {
    ceEl('ce-controls').style.display = 'flex';
    ceEl('ce-table-wrap').style.display = '';
    ceEl('ce-footer').style.display = 'flex';
    ceUpdateDropdowns();
    ceRenderTable();
    ceUpdateFooter();
  }

  // ── Extract value from source cell ──
  function ceExtractValue(sourceVal, sepMode) {
    if (!sourceVal) return '';
    if (sepMode === 'full') return sourceVal.trim();
    if (sepMode === 'last-word') {
      const parts = sourceVal.trim().split(/\s+/);
      return parts[parts.length - 1] || '';
    }
    if (sepMode === 'right') {
      const idx = sourceVal.lastIndexOf('-');
      if (idx >= 0) return sourceVal.substring(idx + 1).trim();
      return sourceVal.trim();
    }
    if (sepMode === 'left') {
      const idx = sourceVal.indexOf('-');
      if (idx >= 0) return sourceVal.substring(0, idx).trim();
      return sourceVal.trim();
    }
    return sourceVal.trim();
  }

  // ── Get append text for a target row ──
  function ceGetAppendText(ri) {
    if (ceEl('ce-mode').value === 'manual') {
      return ceEl('ce-append-text').value;
    }
    const srcCol = +ceEl('ce-source-col').value;
    if (isNaN(srcCol) || srcCol < 0) return '';
    const raw = tgtRows[ri][srcCol] || '';
    const extracted = ceExtractValue(raw, ceEl('ce-separator').value);
    return extracted ? ' - ' + extracted : '';
  }

  // ── Table rendering ──
  function ceEsc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  function ceRenderTable() {
    const thead = ceEl('ce-thead');
    let hh = '<tr><th class="ce-chk"><input type="checkbox" id="ce-hdr-chk" checked></th>';
    tgtHeaders.forEach((h, i) => {
      const selCls = i === ceSelCol ? ' ce-sel-col' : '';
      const copyCls = ceCopyCols.has(i) ? ' ce-copy-col' : '';
      hh += `<th class="${selCls}${copyCls}" data-cecol="${i}"><input type="checkbox" class="ce-col-chk" data-cci="${i}" ${ceCopyCols.has(i)?'checked':''}>${ceEsc(h || 'Col '+i)}</th>`;
    });
    thead.innerHTML = hh + '</tr>';

    const tbody = ceEl('ce-tbody');
    let bh = '';
    tgtRows.forEach((row, ri) => {
      const s = ceSelRows.has(ri);
      const changed = ceChangedRows.has(ri);
      bh += `<tr class="${s ? 'ce-row-sel' : ''}${changed ? ' ce-row-changed' : ''}" data-cerow="${ri}">`;
      bh += `<td class="ce-chk"><input type="checkbox" data-ceri="${ri}" ${s ? 'checked' : ''}></td>`;
      row.forEach((cell, ci) => {
        bh += `<td class="${ci === ceSelCol ? 'ce-tgt' : ''}">${ceEsc(cell)}</td>`;
      });
      bh += '</tr>';
    });
    tbody.innerHTML = bh;

    ceEl('ce-hdr-chk').addEventListener('change', e => {
      if (e.target.checked) tgtRows.forEach((_, i) => ceSelRows.add(i));
      else ceSelRows.clear();
      ceRenderTable(); ceUpdateFooter();
    });
    tbody.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', e => {
        const ri = +e.target.dataset.ceri;
        if (e.target.checked) ceSelRows.add(ri); else ceSelRows.delete(ri);
        e.target.closest('tr').classList.toggle('ce-row-sel', e.target.checked);
        ceUpdateFooter();
      });
    });
    // Column header click = set target column; checkbox = toggle copy column
    thead.querySelectorAll('th[data-cecol]').forEach(th => {
      th.addEventListener('click', e => {
        if (e.target.classList.contains('ce-col-chk')) return; // let checkbox handle itself
        ceSelCol = +th.dataset.cecol;
        ceEl('ce-col-select').value = ceSelCol;
        ceRenderTable();
      });
    });
    thead.querySelectorAll('.ce-col-chk').forEach(cb => {
      cb.addEventListener('change', e => {
        e.stopPropagation();
        const ci = +e.target.dataset.cci;
        if (e.target.checked) ceCopyCols.add(ci); else ceCopyCols.delete(ci);
        e.target.closest('th').classList.toggle('ce-copy-col', e.target.checked);
      });
    });
  }

  function ceUpdateFooter() {
    ceEl('ce-foot-info').textContent = tgtRows.length + ' rows loaded';
    ceEl('ce-foot-sel').textContent = ceSelRows.size + ' of ' + tgtRows.length + ' selected';
  }

  ceEl('ce-col-select').addEventListener('change', e => {
    ceSelCol = e.target.value === '' ? -1 : +e.target.value;
    ceRenderTable();
  });

  ceEl('ce-btn-sel-all').addEventListener('click', () => {
    ceSelRows = new Set(tgtRows.map((_, i) => i)); ceRenderTable(); ceUpdateFooter();
  });
  ceEl('ce-btn-sel-none').addEventListener('click', () => {
    ceSelRows.clear(); ceRenderTable(); ceUpdateFooter();
  });
  ceEl('ce-btn-invert').addEventListener('click', () => {
    const n = new Set();
    tgtRows.forEach((_, i) => { if (!ceSelRows.has(i)) n.add(i); });
    ceSelRows = n; ceRenderTable(); ceUpdateFooter();
  });

  function ceSkip(val, appendText, mode) {
    if (mode === 'hyphen') return /\s-\s/.test(val);
    if (mode === 'custom') return val.includes(appendText.trim());
    return false;
  }

  function ceStatus(msg, type) {
    const bar = ceEl('ce-status');
    bar.style.display = '';
    bar.textContent = msg;
    bar.style.background = type === 'warn' ? '#fff3cd' : type === 'success' ? '#c6efce' : '#cce5ff';
    bar.style.color = type === 'warn' ? '#856404' : type === 'success' ? '#155724' : '#004085';
    bar.style.borderColor = type === 'warn' ? '#e0c84a' : type === 'success' ? '#70ad47' : '#b8daff';
  }

  function ceToast(msg) {
    let t = document.getElementById('ce-toast');
    if (!t) { t = document.createElement('div'); t.id = 'ce-toast'; t.className = 'ce-toast'; document.body.appendChild(t); }
    t.textContent = msg; t.style.display = 'block';
    setTimeout(() => t.style.display = 'none', 2000);
  }

  // ── Preview ──
  ceEl('ce-btn-preview').addEventListener('click', () => {
    const isAuto = ceEl('ce-mode').value === 'auto';
    if (!isAuto && !ceEl('ce-append-text').value) { ceStatus('Enter text to append.', 'warn'); return; }
    if (isAuto && ceEl('ce-source-col').value === '') { ceStatus('Select a source column.', 'warn'); return; }
    if (ceSelCol < 0) { ceStatus('Select a target column first.', 'warn'); return; }
    if (ceSelRows.size === 0) { ceStatus('Select at least one row.', 'warn'); return; }

    const skipMode = ceEl('ce-skip-mode').value;
    let willChange = 0, willSkip = 0, noValue = 0;
    const trs = ceEl('ce-tbody').querySelectorAll('tr');
    trs.forEach(tr => {
      const ri = +tr.dataset.cerow;
      if (!ceSelRows.has(ri)) return;
      const cell = tr.children[ceSelCol + 1];
      const val = tgtRows[ri][ceSelCol];
      const appendText = ceGetAppendText(ri);
      if (!appendText) {
        noValue++;
        cell.classList.remove('ce-will-change', 'ce-has-suffix');
        return;
      }
      if (ceSkip(val, appendText, skipMode)) {
        cell.classList.add('ce-has-suffix'); cell.classList.remove('ce-will-change');
        willSkip++;
      } else {
        cell.classList.add('ce-will-change'); cell.classList.remove('ce-has-suffix');
        cell.textContent = val + appendText;
        willChange++;
      }
    });
    let msg = `Preview: ${willChange} will change, ${willSkip} skipped.`;
    if (noValue > 0) msg += ` ${noValue} no source value.`;
    ceStatus(msg, 'info');
    ceEl('ce-btn-apply').disabled = false;
  });

  // ── Apply ──
  ceEl('ce-btn-apply').addEventListener('click', () => {
    if (ceSelCol < 0) return;
    const isAuto = ceEl('ce-mode').value === 'auto';
    if (!isAuto && !ceEl('ce-append-text').value) return;
    const skipMode = ceEl('ce-skip-mode').value;
    tgtOriginal = tgtRows.map(r => [...r]);
    ceChangedRows = new Set();
    let changed = 0;
    tgtRows.forEach((row, ri) => {
      if (!ceSelRows.has(ri)) return;
      const appendText = ceGetAppendText(ri);
      if (!appendText) return;
      if (ceSkip(row[ceSelCol], appendText, skipMode)) return;
      row[ceSelCol] = row[ceSelCol] + appendText;
      ceChangedRows.add(ri);
      changed++;
    });
    ceRenderTable();
    ceStatus(`Applied: ${changed} cells updated. Use "Hide Unchanged" to filter.`, 'success');
    ceEl('ce-btn-apply').disabled = true;
    ceEl('ce-btn-undo').disabled = false;
  });

  // ── Undo ──
  ceEl('ce-btn-undo').addEventListener('click', () => {
    if (!tgtOriginal.length) return;
    tgtRows = tgtOriginal.map(r => [...r]);
    tgtOriginal = [];
    ceChangedRows = new Set();
    ceRenderTable();
    ceStatus('Undo complete.', 'info');
    ceEl('ce-btn-undo').disabled = true;
  });

  // ── Copy Columns ──
  ceEl('ce-btn-copy').addEventListener('click', () => {
    // Determine which columns to copy: checked columns, or fall back to target column
    let cols = [...ceCopyCols].sort((a,b) => a - b);
    if (cols.length === 0) {
      if (ceSelCol < 0) { ceStatus('Check column headers to copy, or select a target column.', 'warn'); return; }
      cols = [ceSelCol];
    }

    // Determine which rows: selected rows, filtered by "changed only" if checked
    let indices = ceSelRows.size > 0
      ? [...ceSelRows].sort((a,b) => a - b)
      : tgtRows.map((_, i) => i);
    if (ceEl('ce-copy-changed').checked) {
      indices = indices.filter(i => ceChangedRows.has(i));
    }

    if (indices.length === 0) { ceStatus('No rows to copy.', 'warn'); return; }

    // Build tab-separated rows for Excel paste
    const lines = indices.map(i => cols.map(c => tgtRows[i][c]).join('\t'));
    const text = lines.join('\n');

    navigator.clipboard.writeText(text).then(() => ceToast(`Copied ${indices.length} rows x ${cols.length} cols`))
    .catch(() => {
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta);
      ta.select(); document.execCommand('copy');
      document.body.removeChild(ta);
      ceToast(`Copied ${indices.length} rows x ${cols.length} cols`);
    });
  });

  // ── Hide Unchanged / Show All ──
  ceEl('ce-btn-hide-unchanged').addEventListener('click', () => {
    const trs = ceEl('ce-tbody').querySelectorAll('tr');
    let hidden = 0;
    trs.forEach(tr => {
      const ri = +tr.dataset.cerow;
      if (!ceChangedRows.has(ri)) {
        tr.classList.add('ce-row-hidden');
        hidden++;
      }
    });
    ceStatus(`Showing ${tgtRows.length - hidden} changed rows, ${hidden} hidden.`, 'info');
  });

  ceEl('ce-btn-show-all').addEventListener('click', () => {
    ceEl('ce-tbody').querySelectorAll('tr.ce-row-hidden').forEach(tr => tr.classList.remove('ce-row-hidden'));
    ceStatus('All rows visible.', 'info');
  });

  // Public bridge: load from global import
  window.ceLoadSheetData = function(headers, rows) {
    tgtHeaders = headers.map(h => String(h));
    tgtRows = rows.map(r => r.map(c => String(c)));
    ceEl('ce-target-name').textContent = '(from imported sheet) (' + tgtRows.length + ' rows)';
    ceSelRows = new Set(tgtRows.map((_, i) => i));
    ceSelCol = -1;
    ceShowControls();
  };
})();
