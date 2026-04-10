// ═══════════════════════════════════════════
// Shared Utilities
// ═══════════════════════════════════════════

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function escJs(s) {
  return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
}

// CSV parser (proper quoted-field support)
function parseCSV(text) {
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
  if (cur || lines.length) {
    lines.push(cur);
    if (!headers.length) headers.push(...lines);
    else rows.push(lines.splice(0));
  }
  const max = Math.max(headers.length, ...rows.map(r => r.length));
  while (headers.length < max) headers.push('Col ' + headers.length);
  rows.forEach(r => { while (r.length < max) r.push(''); });
  return { headers: headers.map(h => h.trim()), rows };
}

// Read a file (CSV or Excel) and return {headers, rows, sheetNames?, workbook?}
function readUploadedFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'csv') {
      reader.onload = e => {
        const result = parseCSV(e.target.result);
        resolve({ sheets: [{ name: file.name.replace(/\.csv$/i, ''), headers: result.headers, rows: result.rows }] });
      };
      reader.readAsText(file);
    } else {
      reader.onload = e => {
        const d = new Uint8Array(e.target.result);
        const wb = XLSX.read(d, { type: 'array' });
        const sheets = wb.SheetNames.map(name => {
          const ws = wb.Sheets[name];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          const filtered = raw.filter(r => r.some(c => String(c).trim() !== ''));
          if (filtered.length < 2) return null;
          return {
            name,
            headers: filtered[0].map(h => String(h).trim()),
            rows: filtered.slice(1).map(r => r.map(c => String(c).trim()))
          };
        }).filter(Boolean);
        resolve({ sheets });
      };
      reader.readAsArrayBuffer(file);
    }
  });
}
