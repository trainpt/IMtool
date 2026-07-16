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

// ── Year-only date expansion ──
// PickTrace needs full dates (YYYY-MM-DD). Some date columns (e.g. Planted
// Date) carry only a year ("2012") or a year range ("2010-2011"). These
// helpers detect those and expand them to a full date using an operator-chosen
// month/day, KEEPING each row's year. rangePref ('first'|'last') picks which
// year of a range to keep. Anything that is already a full date — or not a
// year/range at all — is returned untouched (so the pass is idempotent).
function isYearOnlyDate(value) {
  const s = String(value == null ? '' : value).trim();
  return /^\d{4}$/.test(s) || /^\d{4}\s*(?:-|–|—|\/|to)\s*\d{4}$/i.test(s);
}
function expandYearOnlyDate(value, mm, dd, rangePref) {
  const s = String(value == null ? '' : value).trim();
  if (!s) return s;
  const mi = parseInt(String(mm), 10), di = parseInt(String(dd), 10);
  if (!(mi >= 1 && mi <= 12) || !(di >= 1 && di <= 31)) return s; // invalid MM/DD → leave as-is
  const MM = String(mi).padStart(2, '0'), DD = String(di).padStart(2, '0');
  let year = null;
  let m = s.match(/^(\d{4})$/);
  if (m) year = m[1];
  else {
    m = s.match(/^(\d{4})\s*(?:-|–|—|\/|to)\s*(\d{4})$/i);
    if (m) year = (String(rangePref) === 'last') ? m[2] : m[1];
  }
  if (!year) return s;
  return year + '-' + MM + '-' + DD;
}
function countYearOnlyDates(rows, colIdxs) {
  let n = 0;
  (rows || []).forEach(r => { (colIdxs || []).forEach(ci => { if (isYearOnlyDate(r[ci])) n++; }); });
  return n;
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
