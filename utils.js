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

// ── Output file names ──
// Every filled template downloads as "<Client> - <Type> - <YYYY-MM-DD>.xlsx",
// or "<Type> - <YYYY-MM-DD>.xlsx" when no client is known. The client lives in
// the Client box in the context strip (#ctxClient), shared by every builder.
// A source file whose name carries the client ("Acme Farms - PickTrace
// Implementation Data Template.xlsx") fills the box — until you type in it,
// after which your value sticks for the session.
window.IMOutName = (function () {
  const box = () => document.getElementById('ctxClient');
  const today = () => {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  };
  // Windows-safe: no \ / : * ? " < > | or control characters. (A trailing dot
  // is harmless here — the date always ends the name.)
  const safe = s => String(s == null ? '' : s).replace(/[\\/:*?"<>|\x00-\x1F]/g, '')
    .replace(/\s+/g, ' ').trim();

  // The client in an Implementation Data Template's file name: everything
  // before "[PickTrace] [year] Implementation …" or "Data Template". '' when
  // the name doesn't follow that shape (a PickTrace export, "template (15)").
  function clientFrom(fileName) {
    const base = String(fileName == null ? '' : fileName).replace(/\.[a-z0-9]+$/i, '').replace(/\s*\(\d+\)\s*$/, '');
    const m = base.match(/^(.*?)[\s\-–—_]*(?:pick\s*trace\s*)?(?:\d{4}(?:\s*-\s*\d{2,4})?\s*)?(?:implementation\b|data\s*template\b).*$/i);
    return m ? safe(m[1].replace(/[\s\-–—_,]+$/, '')) : '';
  }
  function client() { const b = box(); return b ? safe(b.value) : ''; }
  // A newly loaded source file offers its client; a name you typed wins.
  function suggestClient(fileName) {
    const b = box(), c = clientFrom(fileName);
    if (b && c && b.dataset.typed !== '1') b.value = c;
  }
  // A tool's Reset drops an auto-filled client, never one you typed.
  function forget() {
    const b = box();
    if (b && b.dataset.typed !== '1') b.value = '';
  }
  // "Acme Farms - Employees - 2026-10-07 (1 of 3).xlsx"
  function fileName(type, suffix) {
    return [client(), safe(type), today()].filter(Boolean).join(' - ') + (suffix || '') + '.xlsx';
  }
  // The folder Download All bundles everything into.
  function folderName() {
    return [client(), 'PickTrace Uploads', today()].filter(Boolean).join(' - ');
  }
  return { clientFrom, client, suggestClient, forget, fileName, folderName };
})();

// ── Detailed Implementation-template layouts ──
// "Locations (Detailed Version)" differs from the plain tab in two ways the
// readers must undo:
//   - a field whose header cell is an instruction sentence ("Please indicate
//     the 'Site' …") is named by the short banner cell above it ("Site");
//   - a label column headed "Location Fields -->" names the metadata rows
//     under the header (Adv. Filter, Expected Value Type, Preset Values,
//     Tooltip). Any row with a value in that column is metadata, not data.
// `rows` are the sheet's rows, `h` the index of the header row among them.
// Returns { headers, isMeta(row) }; on a plain tab, headers come back
// unchanged and isMeta is always false.
window.IMSheetLayout = (function () {
  const s = v => String(v == null ? '' : v).trim();
  const INSTRUCTION = /^(please|required|optional|enter|select)\b/i;
  function resolve(rows, h) {
    const headers = (rows[h] || []).map(s);
    const banner = h > 0 ? (rows[h - 1] || []).map(s) : [];
    headers.forEach((v, i) => {
      if ((v.length > 40 || INSTRUCTION.test(v)) && banner[i] && banner[i].length <= 30) headers[i] = banner[i];
    });
    const labelCol = headers.findIndex(v => /(-+>|→)\s*$/.test(v));
    if (labelCol >= 0) headers[labelCol] = '';
    return { headers, isMeta: row => labelCol >= 0 && s((row || [])[labelCol]) !== '' };
  }
  return { resolve };
})();

// A cell value as text. Numbers keep Excel's 15 significant digits: a pasted
// formula result stores binary float noise (16983.149999999998) that Excel
// displays as 16983.15. Whole numbers pass through untouched.
function cellText(c) {
  if (c == null) return '';
  if (typeof c === 'number' && isFinite(c) && !Number.isInteger(c)) return String(Number(c.toPrecision(15)));
  return String(c).trim();
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
            rows: filtered.slice(1).map(r => r.map(cellText))
          };
        }).filter(Boolean);
        resolve({ sheets });
      };
      reader.readAsArrayBuffer(file);
    }
  });
}
