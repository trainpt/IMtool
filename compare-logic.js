// Block Compare Tool — reconcile DB vs Add/Remove files, export updated XLSX
(function() {

// ─── State ───
let dbData = null;       // { headers, rows, fileName, sheetName }
let addData = null;      // { headers, rows, fileName, sheetName }
let removeData = null;   // { headers, rows, fileName, sheetName }
let diffResult = null;   // { newAdds, conflicts, removes, skipped, unchanged }
let dbCols = null;       // { name, site, altId, cropVar, acreage, plantCount, locType, archived }
let templateData = null; // { sites: [strings], cropVarieties: [strings], fileName }
let updateTemplateData = null; // bulk UPDATE template (PickTrace export) — supplies the base row for every TO UPDATE row.
let sitesMasterData = null; // { allLongNames: Set, archivedLongNames: Set, fileName, totalRows }
let cmpDateFill = { mm: '', dd: '', range: 'first' }; // year-only → full-date expansion (export)
let initialized = false;

const $ = id => document.getElementById(id);

// Historical date columns that may carry only a year / year range. Matched by
// normHeader (lowercased, trailing '*' stripped). Start Date* is excluded.
const CMP_DATE_COLS = new Set(['wet date', 'germination date', 'planted date',
  'grafting date', 'production start', 'organic certification date']);
function cmpYearOnlyDateCount() {
  if (!diffResult || !dbData || !dbData.headers) return 0;
  const idxs = dbData.headers.map((h, i) => CMP_DATE_COLS.has(normHeader(h)) ? i : -1).filter(i => i >= 0);
  if (!idxs.length) return 0;
  let n = 0;
  (diffResult.newAdds || []).forEach(a => idxs.forEach(ci => { if (isYearOnlyDate(a.row[ci])) n++; }));
  return n;
}

// ─── Helpers ───
function extractCode(s) {
  const m = String(s == null ? '' : s).trim().match(/^([A-Za-z0-9&]+)/);
  return m ? m[1].toUpperCase() : '';
}
function isNumericCode(c) { return /^\d+$/.test(c); }
function isAlphaCode(c) { return /[A-Za-z]/.test(c); }

function normalizeNumber(v) {
  if (v == null) return '';
  const s = String(v).replace(/[, ]/g, '').trim();
  if (s === '') return '';
  const n = parseFloat(s);
  return isNaN(n) ? s.toUpperCase() : n;
}

function valuesEqual(a, b) {
  const na = normalizeNumber(a), nb = normalizeNumber(b);
  if (typeof na === 'number' && typeof nb === 'number') return na === nb;
  return String(a == null ? '' : a).trim().toUpperCase() === String(b == null ? '' : b).trim().toUpperCase();
}

// Strip trailing '*' (PickTrace marks required columns as "Site*", "Name*").
function normHeader(h) {
  return String(h || '').toLowerCase().replace(/\*+$/, '').trim();
}

function findColIdx(headers, candidates) {
  if (!headers) return -1;
  const lows = headers.map(normHeader);
  for (let i = 0; i < lows.length; i++) {
    for (const c of candidates) {
      if (lows[i] === c) return i;
    }
  }
  for (let i = 0; i < lows.length; i++) {
    for (const c of candidates) {
      if (lows[i].includes(c)) return i;
    }
  }
  return -1;
}

// Parse "CODE(NAME)" or "CODE (NAME)" or "NAME (CODE)" into { code, name }.
// Heuristic: code is whichever side is shorter, all-uppercase / all-numeric / no spaces.
function parseCodeName(s) {
  const str = String(s == null ? '' : s).trim();
  if (!str) return { code: '', name: '' };
  // Closing paren is OPTIONAL — handles common typos like "MADERA CITRUS VENTURES (AUSTMAD".
  const m = str.match(/^([^()]+?)\s*\(\s*([^()]+?)\s*\)?\s*$/);
  if (!m) {
    // No parens — treat the entire string as the NAME. PickTrace exports many
    // sites as plain names (e.g. "CENTRAL", "DESERT", "WAYNE FLAGG") with no
    // embedded code. Misclassifying a short all-caps name as a code corrupts
    // downstream long-name extraction and Alt ID suggestions.
    return { code: '', name: str };
  }
  const a = m[1].trim(), b = m[2].trim();
  const looksCode = s => /^[A-Z0-9&]+$/.test(s) && s.length <= 10;
  if (looksCode(a) && !looksCode(b)) return { code: a, name: b };
  if (!looksCode(a) && looksCode(b)) return { code: b, name: a };
  return { code: a, name: b };
}

// Build canonical block string "CODE (NAME)" with single space.
function canonBlock(s) {
  const { code, name } = parseCodeName(s);
  if (code && name) return code + ' (' + name + ')';
  return String(s || '').trim();
}

// Auto-detect by content: scan rows, return index of first column whose values look numeric-coded.
function detectColByContent(rows, predicate, maxRows) {
  if (!rows || !rows.length) return -1;
  const sampleN = Math.min(rows.length, maxRows || 30);
  const ncols = rows[0].length;
  let best = -1, bestScore = 0;
  for (let c = 0; c < ncols; c++) {
    let score = 0;
    for (let r = 0; r < sampleN; r++) {
      const code = extractCode(rows[r][c]);
      if (predicate(code)) score++;
    }
    if (score > bestScore) { bestScore = score; best = c; }
  }
  return bestScore > 0 ? best : -1;
}

function todayStr() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return d.getFullYear() + '_' + pad(d.getMonth() + 1) + '_' + pad(d.getDate());
}

// ─── Bulk upload template helpers ───
// Parse a PickTrace bulk template. Captures DATA ENTRY column headers (target schema)
// AND the DROP-DOWN INPUTS values for every column.
function parseTemplateWorkbook(wb, fileName) {
  // 1. DATA ENTRY headers — the authoritative output schema for the export.
  const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
  const dataWs = wb.Sheets[dataName];
  let dataEntryHeaders = [];
  if (dataWs) {
    const raw = XLSX.utils.sheet_to_json(dataWs, { header: 1, defval: '' });
    if (raw.length) dataEntryHeaders = raw[0].map(h => String(h || '').trim()).filter(h => h);
  }

  // 2. DROP-DOWN INPUTS — every column becomes a Set<string> of valid values.
  // Keyed by normHeader so we can look up by data-entry column name.
  const dropdowns = new Map();
  const dropName = wb.SheetNames.find(n => /drop.?down/i.test(n)) || wb.SheetNames[1];
  if (dropName) {
    const dropWs = wb.Sheets[dropName];
    const raw = XLSX.utils.sheet_to_json(dropWs, { header: 1, defval: '' });
    if (raw.length >= 2) {
      const headers = raw[0].map(h => String(h || '').trim());
      headers.forEach((h, ci) => {
        if (!h) return;
        const key = normHeader(h);
        const set = dropdowns.get(key) || new Set();
        raw.slice(1).forEach(r => {
          const v = String(r[ci] != null ? r[ci] : '').trim();
          if (v) set.add(v);
        });
        if (set.size) dropdowns.set(key, set);
      });
    }
  }

  // Convenience extracts.
  const sites = dropdowns.has('site') ? [...dropdowns.get('site')] : [];
  const cvs   = dropdowns.has('crop & variety') ? [...dropdowns.get('crop & variety')] : [];

  // Pre-split the Crop & Variety list into prefix / suffix sets for the split export columns.
  const cropPrefixes = new Set();
  const varietySuffixes = new Set();
  const cvPairs = new Set();
  cvs.forEach(s => {
    const dash = s.indexOf('-');
    if (dash > 0) {
      const c = s.substring(0, dash).trim();
      const v = s.substring(dash + 1).trim();
      cropPrefixes.add(c);
      varietySuffixes.add(v);
      cvPairs.add(c + '|' + v);
    }
  });

  if (!dataEntryHeaders.length && !sites.length && !cvs.length) return null;
  return {
    dataEntryHeaders,
    sites,
    cropVarieties: cvs,
    dropdowns,
    cropPrefixes,
    varietySuffixes,
    cvPairs,
    fileName
  };
}

// Parse the bulk UPDATE template — same column structure as the create template
// PLUS an "Is Archived" column at the end. Contains every existing PickTrace
// location with its current values; used as the base row for TO UPDATE entries
// so we preserve all 42 columns of pre-existing data and only overlay the
// fields that actually changed.
//
// Indexed by composite key: stripCommas(Site) + Name code + Crop & Variety.
function parseUpdateTemplate(wb, fileName) {
  const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[dataName];
  if (!ws) return null;
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!raw.length) return null;
  const headers = raw[0].map(h => String(h || '').trim());
  const rows = raw.slice(1).filter(r => r.some(c => c != null && String(c).trim() !== ''));
  const siteIdx = headers.findIndex(h => /^site\*?$/i.test(h));
  const nameIdx = headers.findIndex(h => /^name\*?$/i.test(h));
  const cvIdx   = headers.findIndex(h => /^crop\s*&\s*variety\*?$/i.test(h));
  if (siteIdx < 0 || nameIdx < 0 || cvIdx < 0) return null;
  const byKey = new Map();
  let dups = 0;
  rows.forEach((r, i) => {
    const k = makeUpdateKey(r[siteIdx], r[nameIdx], r[cvIdx]);
    if (!k) return;
    if (byKey.has(k)) { dups++; return; } // first wins
    byKey.set(k, { row: r, rowIndex: i });
  });
  // Header lookup keyed by normHeader for column-by-column overlay.
  const headerIdx = new Map();
  headers.forEach((h, i) => headerIdx.set(normHeader(h), i));
  return { headers, rows, byKey, headerIdx, siteIdx, nameIdx, cvIdx, fileName, dupCount: dups };
}

// Build an updateTemplateData-equivalent object from the loaded DB itself.
// Valid when the DB is a complete ALL RECORDS PickTrace export — same column
// structure (Site* / Name* / Crop & Variety* / Is Archived) as the bulk update
// template. Lets us run the TO UPDATE / DATA ENTRY (UPDATE) flow without
// requiring a redundant slot-6 upload.
function buildUpdateTemplateFromDb() {
  if (!dbData || !dbData.headers || !dbData.rows.length) return null;
  const headers = dbData.headers.slice();
  const rows = dbData.rows.slice();
  const siteIdx = headers.findIndex(h => /^site\*?$/i.test(String(h || '').trim()));
  const nameIdx = headers.findIndex(h => /^name\*?$/i.test(String(h || '').trim()));
  const cvIdx   = headers.findIndex(h => /^crop\s*&\s*variety\*?$/i.test(String(h || '').trim()));
  if (siteIdx < 0 || nameIdx < 0 || cvIdx < 0) return null;
  const byKey = new Map();
  let dups = 0;
  rows.forEach((r, i) => {
    const k = makeUpdateKey(r[siteIdx], r[nameIdx], r[cvIdx]);
    if (!k) return;
    if (byKey.has(k)) { dups++; return; }
    byKey.set(k, { row: r, rowIndex: i });
  });
  const headerIdx = new Map();
  headers.forEach((h, i) => headerIdx.set(normHeader(h), i));
  return { headers, rows, byKey, headerIdx, siteIdx, nameIdx, cvIdx, fileName: '(derived from database)', dupCount: dups, _fromDb: true };
}

// Composite key: Site (commas stripped, upper) | Name code (extracted) | Crop & Variety (upper, trimmed).
function makeUpdateKey(site, name, cv) {
  const s = stripSiteCommas(String(site == null ? '' : site)).toUpperCase();
  const n = extractCode(String(name == null ? '' : name));
  const c = String(cv == null ? '' : cv).trim().toUpperCase();
  if (!s || !n || !c) return '';
  return s + '||' + n + '||' + c;
}

// Parse the Sites master list (CSV/XLSX). Expected columns include:
//   Group, Group ID, Site, Site ID, Location, Location ID, Location Archived
// Returns { allLongNames, archivedLongNames, fileName, totalRows }
//   - allLongNames: Set of long-form Site names found anywhere in the file
//   - archivedLongNames: Set of long-form Site names where ALL their locations are archived
function parseSitesMaster(headers, rows, fileName) {
  const siteCol = headers.findIndex(h => /^site$/i.test(String(h || '').trim()));
  const archCol = headers.findIndex(h => /location\s+archived|^archived$/i.test(String(h || '').trim()));
  if (siteCol < 0) return null;

  // Per-site tally keyed by long name. Each entry also tracks the embedded code
  // when the master CSV stores Site like "AUSTMAD(...)" or "MADERA... (AUSTMAD".
  const perSite = new Map(); // longNameUpper -> { longName, code, total, archived }
  rows.forEach(r => {
    const siteRaw = String(r[siteCol] || '').trim();
    if (!siteRaw) return;
    const parts = parseCodeName(siteRaw);
    const ln = stripSiteCommas(parts.name || siteRaw);
    if (!ln) return;
    const key = ln.toUpperCase();
    if (!perSite.has(key)) perSite.set(key, { longName: ln, code: parts.code || '', total: 0, archived: 0 });
    const e = perSite.get(key);
    if (parts.code && !e.code) e.code = parts.code;
    e.total++;
    if (archCol >= 0) {
      const v = String(r[archCol] || '').trim().toUpperCase();
      if (v === 'TRUE' || v === '1' || v === 'YES') e.archived++;
    }
  });

  const allLongNames = new Set();
  const archivedLongNames = new Set();
  const allCodes = new Set();
  const archivedCodes = new Set();
  perSite.forEach((e, key) => {
    allLongNames.add(key);
    const fullyArchived = e.total > 0 && e.archived === e.total;
    if (fullyArchived) archivedLongNames.add(key);
    if (e.code) {
      const codeUp = e.code.toUpperCase();
      allCodes.add(codeUp);
      if (fullyArchived) archivedCodes.add(codeUp);
    }
  });
  return { allLongNames, archivedLongNames, allCodes, archivedCodes, fileName, totalRows: rows.length };
}

// Look up the canonical Site string given the parsed grower {code, name}.
// Templates list sites in formats like "AARON FARMS INC. (AAROFAR)" or "WAYNE FLAGG" or "3FLGCITR (3 FLAGS CITRUS, LLC.)".
function lookupCanonicalSite(grower) {
  if (!templateData || !templateData.sites.length) return '';
  const code = (grower.code || '').toUpperCase();
  const name = (grower.name || '').toUpperCase();
  for (const s of templateData.sites) {
    const parts = parseCodeName(s);
    const sCode = (parts.code || '').toUpperCase();
    const sName = (parts.name || s).toUpperCase();
    // Match if either the code or the name lines up
    if (code && sCode && sCode === code) return s;
    if (name && sName === name) return s;
    if (code && sName === code) return s;        // template stored bare code
    if (name && sCode === name) return s;        // template stored bare name as "code"
  }
  return '';
}

// Look up a Crop & Variety string. Returns one of:
//   1. The exact value already learned from a matched DB row, or
//   2. A template value where the variety part matches the variety code (e.g. "RIO" -> "Grapefruit-Rio Red"),
//   3. A combination of learned Comm-name + learned Variety-name if both seen separately,
//   4. '' if nothing fits.
// strict=true: only return template-validated values; '' if no dropdown match.
//   Used when WRITING to the export so PickTrace doesn't reject the row.
// strict=false (default): falls back to a best-effort "Crop-Variety" guess for
//   display purposes (e.g., the "Crops & Varieties to Create" panel).
function lookupCanonicalCropVariety(comm, variety, learnedCV, perComm, perVariety, strict) {
  const key = (comm + '|' + variety).toUpperCase();
  if (learnedCV.has(key)) return learnedCV.get(key);

  const cropName = perComm.get((comm || '').toUpperCase()) || '';
  const varName  = perVariety.get((variety || '').toUpperCase()) || '';

  // If we have both learned parts, build & validate.
  if (cropName && varName) {
    const candidate = cropName + '-' + varName;
    if (templateData) {
      const hit = templateData.cropVarieties.find(s => s.toUpperCase() === candidate.toUpperCase());
      if (hit) return hit;
    }
    return candidate;
  }

  // Template fuzzy-search by variety code as a prefix of the variety part.
  if (templateData && variety) {
    const vUp = variety.toUpperCase();
    const candidates = templateData.cropVarieties.filter(s => {
      const dash = s.indexOf('-');
      if (dash < 0) return false;
      const varPart = s.substring(dash + 1).toUpperCase();
      return varPart.startsWith(vUp) || varPart.split(/\s+/)[0] === vUp;
    });
    if (cropName) {
      const refined = candidates.filter(s => s.toUpperCase().startsWith(cropName.toUpperCase() + '-'));
      if (refined.length === 1) return refined[0];
    }
    if (candidates.length === 1) return candidates[0];
  }
  // STRICT: nothing matched the template — return blank so the export cell is empty.
  // The user creates the missing C&V in PickTrace then re-uploads the template;
  // the next run's fuzzy-search will find it.
  if (strict) return '';
  // Non-strict fallback for the in-app "Crops & Varieties to Create" panel only.
  if (cropName && variety) return cropName + '-' + variety;
  if (comm && variety) return comm + '-' + variety;
  return '';
}

// ─── File reading ───
function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      reader.onload = e => {
        const wb = XLSX.read(e.target.result, { type: 'string' });
        resolve(wb);
      };
      reader.readAsText(file);
    } else {
      reader.onload = e => {
        const d = new Uint8Array(e.target.result);
        resolve(XLSX.read(d, { type: 'array' }));
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

function workbookToSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const filtered = raw.filter(r => r.some(c => String(c).trim() !== ''));
  if (filtered.length < 2) return null;
  return {
    name: sheetName,
    headers: filtered[0].map(h => String(h).trim()),
    rows: filtered.slice(1).map(r => r.map(c => (c == null ? '' : String(c).trim())))
  };
}

// Sheet picker modal — returns chosen sheet object (headers, rows, name) via callback.
function pickSheetIfNeeded(wb, fileName, cb) {
  const sheets = wb.SheetNames
    .map(n => workbookToSheet(wb, n))
    .filter(Boolean);
  if (sheets.length === 0) { alert('No data sheets found in ' + fileName); return; }
  if (sheets.length === 1) { cb(sheets[0]); return; }

  let picker = $('cmp-sheet-picker');
  if (!picker) {
    picker = document.createElement('div');
    picker.id = 'cmp-sheet-picker';
    picker.className = 'modal-overlay show';
    document.body.appendChild(picker);
  }
  let html = '<div class="modal"><h3>Select sheet from ' + escHtml(fileName) + '</h3><div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px;">';
  sheets.forEach(s => {
    html += '<button class="btn btn-ghost cmp-sheet-pick" data-name="' + escHtml(s.name) + '" style="justify-content:space-between;">' +
      '<span>' + escHtml(s.name) + '</span><span class="text-muted small">' + s.rows.length + ' rows</span></button>';
  });
  html += '</div><div class="modal-actions"><button class="btn btn-ghost" id="cmp-sheet-cancel">Cancel</button></div></div>';
  picker.innerHTML = html;
  picker.style.display = 'flex';

  picker.querySelector('#cmp-sheet-cancel').addEventListener('click', () => { picker.style.display = 'none'; });
  picker.querySelectorAll('.cmp-sheet-pick').forEach(btn => {
    btn.addEventListener('click', () => {
      picker.style.display = 'none';
      const chosen = sheets.find(s => s.name === btn.dataset.name);
      if (chosen) cb(chosen);
    });
  });
}

// ─── Import from saved Org Data + Data Progress sessions ───
// The 8-column canonical schema used by the Org Data tree (see app.js HEADERS):
//   [Location, Sites, Crop, Variety, Location Type, Planted Date, Acreage, Plant Count]
// Maps directly to a subset of the bulk update template's 42-column schema.
// Other 30+ columns (Alt ID, Length, Start Date*, custom fields, etc.) are
// emitted blank — the user fills required (*) ones via the Empty Required
// Columns picklist after running the comparison.
const BULK_UPDATE_HEADERS = [
  'Site*','Name*','Alt ID','Location Type*','Crop & Variety*','Planted At','Acreage','Length',
  'Plant Count','Start Date*','Location Group','Lot Number','Rootstock','Plant/Seed Seller',
  'Plant/Seed Producer','Clone/Subvariety','Growing Manager','Production Status','Training Style',
  'Trellis Type','Mulch Type','Row Direction','Organic Status','Organic Certifier','Wet Date',
  'Germination Date','Planted Date','Grafting Date','Production Start','Organic Certification Date',
  'Stand Count','Row/Bed Count','Post Count','Percent Covered','Row Spacing, in.','Plant Spacing, in.',
  'Post Spacing, in.','Bed Width, in.','Custom Data 1','Custom Data 2','Custom Data 3','Is Archived'
];

function ordRowFromOrgRow(row, headers) {
  const idx = {};
  const keyOrder = [];
  (headers || []).forEach((h, i) => {
    const k = String(h || '').trim().toLowerCase().replace(/^#/, '');
    if (k && idx[k] == null) { idx[k] = i; keyOrder.push(k); }
  });
  // Two-phase fuzzy getter — used for fields where substring fallback is safe
  // (e.g. "Block Name", "Block Code" both legitimately contain "block").
  const get = (...substrs) => {
    for (const s of substrs) {
      const i = idx[s];
      if (i != null) {
        const v = row[i];
        if (v != null && String(v).trim() !== '') return String(v).trim();
      }
    }
    for (const s of substrs) {
      for (const k of keyOrder) {
        if (k === s) continue;
        if (k.includes(s)) {
          const v = row[idx[k]];
          if (v != null && String(v).trim() !== '') return String(v).trim();
        }
      }
    }
    return '';
  };
  // Exact-only variant — for fields where substring matching would catch
  // numeric-ID columns ("site_id", "crop_id", "location_type_id") instead of
  // the human-readable column we want. Per user spec for pt_locations.csv,
  // the only columns that matter are display_name, plant_name, planted_date,
  // is_archived — the *_id columns must be ignored.
  const getExact = (...keys) => {
    for (const k of keys) {
      const i = idx[k];
      if (i != null) {
        const v = row[i];
        if (v != null && String(v).trim() !== '') return String(v).trim();
      }
    }
    return '';
  };

  // PickTrace's pt_locations.csv embeds the site name inside display_name
  // before " - ". E.g.:
  //   "Maricopa 400 - 2022 Devlp Clementine TW"  → Site="Maricopa 400", Name="2022 Devlp Clementine TW"
  //   "Thermal Plaza - Producing Lemon"          → Site="Thermal Plaza", Name="Producing Lemon"
  //   "FAN ROOM"                                 → Site=(empty), Name="FAN ROOM"
  // Split on the FIRST " - " so multi-dash names like
  // "Lemon Hill - 2018 Devlp Avocado" still work cleanly.
  let siteFromDisplay = '';
  let nameFromDisplay = '';
  const displayName = getExact('display_name', 'display name');
  if (displayName) {
    const dashIdx = displayName.indexOf(' - ');
    if (dashIdx > 0) {
      siteFromDisplay = displayName.slice(0, dashIdx).trim();
      nameFromDisplay = displayName.slice(dashIdx + 3).trim();
    } else {
      nameFromDisplay = displayName;
    }
  }

  const out = new Array(BULK_UPDATE_HEADERS.length).fill('');
  // Site*: prefer an explicit site/grower column; otherwise pull from the
  // dash-prefix in display_name; never grab "site_id".
  out[0]  = getExact('sites', 'site', 'site_name', 'site name', 'grower', 'farm') || siteFromDisplay;
  // Name*: prefer the display_name's post-dash portion (or full display_name
  // when there's no dash); fall back to fuzzy "name"/"block" for other formats.
  out[1]  = nameFromDisplay || get('name', 'block');
  out[2]  = get('alt id', 'altid', 'code');
  // Location Type*: exact only — never match "location_type_id".
  out[3]  = getExact('location type', 'type');
  // Crop & Variety*: prefer plant_name (PickTrace's crop column), then fall
  // back to a synthesized crop-variety pair. crop_id/variety_id are ignored.
  let cv = getExact('plant_name', 'plant name', 'crop & variety', 'crop and variety');
  if (!cv) {
    const crop = getExact('crop', 'comm');
    const variety = getExact('variety');
    cv = (crop && variety) ? (crop + '-' + variety) : (crop || variety);
  }
  out[4]  = cv;
  out[5]  = getExact('planted at', 'planted_at');
  out[6]  = getExact('acreage', 'acres', 'acre', 'hectares');
  out[7]  = getExact('length', 'length_meters');
  out[8]  = getExact('plant count', 'plant_count', 'tree count', 'trees', 'plants', 'quantity');
  out[26] = getExact('planted date', 'planted_date');
  const archV = getExact('is archived', 'is_archived', 'archived');
  out[41] = archV ? archV.toUpperCase() : 'FALSE';
  return out;
}

// Scan every saved data source for an org and build a numeric-ID → site-name
// lookup table. PickTrace exports use site_id (a number); the user's saved
// sites OCR or any 2-column ID/Name sheet provides the translation.
function buildOrgSitesIdMap(store, orgName) {
  const map = new Map();
  const consider = (rows, headers) => {
    if (!rows || !rows.length || !headers || !headers.length) return;
    let idCol = -1, nameCol = -1;
    headers.forEach((h, i) => {
      const k = String(h || '').trim().toLowerCase().replace(/^#/, '');
      if (idCol < 0 && (k === 'id' || k === 'site_id' || k === 'site id' || k === '#')) idCol = i;
      if (nameCol < 0 && (k === 'site' || k === 'sites' || k === 'site name' || k === 'site_name')) nameCol = i;
    });
    if (idCol < 0 || nameCol < 0) return;
    rows.forEach(r => {
      const id = String((r || [])[idCol] || '').trim();
      const name = String((r || [])[nameCol] || '').trim();
      if (id && name && /^-?\d+$/.test(id)) map.set(id, name);
    });
  };
  Object.values((store.trackerSaves || {})[orgName] || {}).forEach(s => consider(s.rows, s.headers));
  ((store.ocrSaves || {})[orgName] || []).forEach(item => consider(item.rows, item.headers));
  Object.values(store.sheets || {}).forEach(s => { if (s && s.orgName === orgName) consider(s.rows, s.headers); });
  return map;
}

// Walk every saved data source for an org and produce a flat array of rows in
// the bulk update template's 42-column schema.
function aggregateOrgData(orgName) {
  if (typeof getStore !== 'function') {
    console.warn('[Block Compare] data-store.js not loaded — cannot import.');
    return null;
  }
  const store = getStore();
  const rows = [];
  const seen = new Set(); // dedupe by Site|Name|CV — the same block can appear in multiple sources

  // Build numeric site_id → site_name lookup from any auxiliary sheets/OCR
  // saves the user has stored under this org. PickTrace's pt_locations export
  // uses numeric site_ids (12, 13, 14, …) which are meaningless without the
  // sites table to translate them.
  const sitesIdMap = buildOrgSitesIdMap(store, orgName);
  if (sitesIdMap.size) {
    console.log('[Block Compare][import] sites-id lookup built — ' + sitesIdMap.size + ' entries (e.g. ' +
      [...sitesIdMap.entries()].slice(0, 3).map(([k, v]) => k + '→' + v).join(', ') + ')');
  } else {
    console.warn('[Block Compare][import] No sites-id lookup found for "' + orgName + '". ' +
      'Sites will be exported as numeric IDs.\n' +
      'Fix: save a 2-column sheet under this org (Data Progress) with column headers ' +
      'like "Site ID" and "Site Name" mapping each numeric site_id to its name. ' +
      'The importer scans every saved sheet/OCR/tracker session for those headers ' +
      'and uses the mapping automatically.');
  }
  // Track unmapped numeric site_ids so we can surface them at the end.
  const unmappedIds = new Set();

  let dropped = 0;
  const push = (rawRow, headers, sourceLabel) => {
    if (!rawRow) return;
    const mapped = ordRowFromOrgRow(rawRow, headers || []);
    // Translate numeric Site* via the sites-id lookup if available; otherwise
    // record the unmapped ID so we can list them at the end.
    if (mapped[0] && /^-?\d+$/.test(mapped[0])) {
      if (sitesIdMap.has(mapped[0])) {
        mapped[0] = sitesIdMap.get(mapped[0]);
      } else {
        unmappedIds.add(mapped[0]);
      }
    }
    // Strip the "<site_id><display_name>" concatenation from Name* if present.
    // PickTrace's location_id is "13FANROOM"; if Name* came from that field it'll
    // have the site ID prefix glued on. Drop the prefix when it matches a known site_id.
    if (mapped[1]) {
      const m = String(mapped[1]).match(/^(\d+)(.+)$/);
      if (m && sitesIdMap.has(m[1])) {
        mapped[1] = m[2].trim();
      }
    }
    const hasContent = mapped.slice(0, 41).some(v => v != null && String(v).trim() !== '');
    if (!hasContent) { dropped++; return; }
    const key = (mapped[0] || '').toUpperCase() + '||' + (mapped[1] || '').toUpperCase() + '||' + (mapped[4] || '').toUpperCase();
    if (key !== '||||' && seen.has(key)) return;
    if (key !== '||||') seen.add(key);
    rows.push(mapped);
  };

  // 1. Org Data tree: org → ranches → rows (HEADERS schema from app.js).
  // The RANCH KEY is the actual site name (e.g. "CENTRAL"); the Sites column
  // inside each row may be empty, garbage, or a stray site ID. Override the
  // Sites column with the ranch key before mapping so Site* gets a real name.
  const HEADERS = ['Location','Sites','Crop','Variety','Location Type','Planted Date','Acreage','Plant Count'];
  const SITES_IDX = HEADERS.indexOf('Sites');
  const org = (store.orgs || {})[orgName];
  let orgRanchRows = 0;
  if (org && org.ranches) {
    Object.entries(org.ranches).forEach(([rName, ranch]) => {
      let cleanedRanchName = String(rName || '').trim();
      const isNumericRanch = /^-?\d+$/.test(cleanedRanchName);
      // Translate a numeric ranch key (site_id) to a real site name if a
      // lookup is available; otherwise drop it so Site* stays empty and the
      // Empty Required Columns picklist surfaces it for the user to fill once.
      if (isNumericRanch) {
        cleanedRanchName = sitesIdMap.has(cleanedRanchName)
          ? sitesIdMap.get(cleanedRanchName)
          : '';
      }
      orgRanchRows += (ranch.rows || []).length;
      (ranch.rows || []).forEach(r => {
        const cloned = (r || []).slice();
        const existing = String(cloned[SITES_IDX] || '').trim();
        // Only stamp the ranch name into the Sites column if we have a
        // meaningful, non-numeric ranch name AND the row doesn't already
        // carry a longer/more specific Sites value.
        if (cleanedRanchName && (!existing || existing.length < cleanedRanchName.length)) {
          cloned[SITES_IDX] = cleanedRanchName;
        }
        push(cloned, HEADERS, 'org-ranch');
      });
    });
  }

  // 2. Tracker saves: each saved session has its own headers.
  const trkSessions = (store.trackerSaves || {})[orgName] || {};
  let trkRows = 0;
  Object.entries(trkSessions).forEach(([key, sess]) => {
    const n = (sess.rows || []).length;
    trkRows += n;
    console.log('[Block Compare][import] tracker session "' + key + '" — headers=', sess.headers, 'rowCount=' + n);
    (sess.rows || []).forEach(r => push(r, sess.headers || [], 'tracker'));
  });

  // 3. OCR saves under this org.
  const ocrItems = (store.ocrSaves || {})[orgName] || [];
  let ocrRows = 0;
  ocrItems.forEach((item, ii) => {
    const n = (item.rows || []).length;
    ocrRows += n;
    console.log('[Block Compare][import] ocr save #' + ii + ' "' + (item.name || '') + '" — headers=', item.headers, 'rowCount=' + n);
    (item.rows || []).forEach(r => push(r, item.headers || [], 'ocr'));
  });

  // 4. Imported sheets where the orgName matches.
  let sheetRows = 0;
  Object.entries(store.sheets || {}).forEach(([sName, sheet]) => {
    if (sheet && sheet.orgName === orgName) {
      const n = (sheet.rows || []).length;
      sheetRows += n;
      console.log('[Block Compare][import] sheet "' + sName + '" — headers=', sheet.headers, 'rowCount=' + n);
      (sheet.rows || []).forEach(r => push(r, sheet.headers || [], 'sheet'));
    }
  });

  console.log('[Block Compare][import] aggregate for "' + orgName + '": orgRanch=' + orgRanchRows +
    ', tracker=' + trkRows + ', ocr=' + ocrRows + ', sheets=' + sheetRows +
    ' → kept=' + rows.length + ' (dropped ' + dropped + ' empty-after-map)');

  if (unmappedIds.size) {
    const sortedIds = [...unmappedIds].sort((a, b) => Number(a) - Number(b));
    console.warn('[Block Compare][import] ' + sortedIds.length + ' unmapped site_id(s) — Site* shown as raw ID for these: ' +
      sortedIds.join(', '));
  }

  return { headers: BULK_UPDATE_HEADERS.slice(), rows, unmappedIds: [...unmappedIds] };
}

function summarizeOrgSources(orgName) {
  const store = getStore();
  const org = (store.orgs || {})[orgName];
  let ranchCount = 0, ranchRows = 0;
  if (org && org.ranches) {
    const rNames = Object.keys(org.ranches);
    ranchCount = rNames.length;
    rNames.forEach(rn => { ranchRows += (org.ranches[rn].rows || []).length; });
  }
  const trkSessions = Object.keys((store.trackerSaves || {})[orgName] || {}).length;
  const ocrCount = ((store.ocrSaves || {})[orgName] || []).length;
  const sheetCount = Object.values(store.sheets || {}).filter(s => s && s.orgName === orgName).length;
  return { ranchCount, ranchRows, trkSessions, ocrCount, sheetCount };
}

function openOrgImportPicker() {
  const modal = $('cmp-org-import-modal');
  const list = $('cmp-org-import-list');
  if (!modal || !list) return;
  const names = (typeof getAllOrgNames === 'function') ? getAllOrgNames() : [];
  if (!names.length) {
    list.innerHTML = '<div class="cmp-org-list-empty">No saved org data found. Add data via the <b>Org Data</b> tab or the <b>Data Progress</b> tab first.</div>';
  } else {
    list.innerHTML = names.map(n => {
      const s = summarizeOrgSources(n);
      const parts = [];
      if (s.ranchCount) parts.push(s.ranchCount + ' ranches (' + s.ranchRows + ' rows)');
      if (s.trkSessions) parts.push(s.trkSessions + ' tracker sessions');
      if (s.ocrCount) parts.push(s.ocrCount + ' OCR saves');
      if (s.sheetCount) parts.push(s.sheetCount + ' sheets');
      const sub = parts.length ? parts.join(' &middot; ') : 'no data';
      return '<div class="cmp-org-list-item" data-org="' + escHtml(n) + '">' +
        '<span class="cmp-org-name">' + escHtml(n) + '</span>' +
        '<span class="cmp-org-counts">' + sub + '</span>' +
        '</div>';
    }).join('');
  }
  modal.style.display = 'flex';
}

function closeOrgImportPicker() {
  const modal = $('cmp-org-import-modal');
  if (modal) modal.style.display = 'none';
}

function importOrgIntoDb(orgName) {
  const aggregated = aggregateOrgData(orgName);
  if (!aggregated || !aggregated.rows.length) {
    alert('No usable rows after mapping "' + orgName + '" — open the browser console (F12) to see what each source returned. ' +
      'Most likely the saved sheet uses headers we don\'t recognize. ' +
      'Look for "[Block Compare][import]" lines.');
    return;
  }
  const fileLabel = '(Org Data: ' + orgName + ')';
  dbData = {
    headers: aggregated.headers,
    rows: aggregated.rows,
    fileName: fileLabel,
    sheetName: '(synthesized)',
    _fromOrgImport: true
  };
  $('cmp-db-name').textContent = fileLabel;
  renderDbMeta();
  const meta = $('cmp-db-meta');
  if (meta) {
    const prior = meta.querySelector('.cmp-warn-banner:not(.cmp-archive-warn)');
    if (prior) prior.remove();
  }
  // Surface the unmapped-site-IDs warning in the panel so the user sees it
  // without needing the console. This is the most common reason for "sites
  // and varieties don't align" in the Sites-to-Create panel.
  if (aggregated.unmappedIds && aggregated.unmappedIds.length && meta) {
    const head = aggregated.unmappedIds.slice(0, 10).join(', ');
    const tail = aggregated.unmappedIds.length > 10 ? ', …' : '';
    const banner = '<div class="cmp-warn-banner">&#9888; <b>' +
      aggregated.unmappedIds.length + ' numeric site_id(s)</b> couldn\'t be translated to site names: ' +
      '<code>' + escHtml(head + tail) + '</code>. ' +
      'Save a 2-column sheet under this org (via <b>Data Progress</b>) with column headers ' +
      '<code>Site ID</code> + <code>Site Name</code> mapping each ID to its name; the importer ' +
      'will pick it up automatically next time. Until then, Site* exports as the raw numeric ID.</div>';
    meta.insertAdjacentHTML('afterbegin', banner);
  }
  updateRunButton();
  closeOrgImportPicker();
  console.log('[Block Compare] Imported ' + aggregated.rows.length + ' rows from org "' + orgName + '" into Database.');
}

// ─── File-type fingerprinting ───
// Each upload slot has an expected shape. After parsing, we sniff the headers
// + sheet names to detect the file's actual type. If it doesn't match the slot,
// surface a warning under the slot so the user can spot a swapped upload before
// running the comparison.
const FILE_TYPES = {
  db:             { num: 1, label: 'Database' },
  add:            { num: 2, label: 'Locations to be added' },
  rm:             { num: 3, label: 'Locations to be removed' },
  tpl:            { num: 4, label: 'Bulk Create Template' },
  'sites-master': { num: 5, label: 'Sites export' },
  updt:           { num: 6, label: 'Bulk Update Template' }
};

function fingerprintFile(headers, sheetNames) {
  const h = (headers || []).map(s => normHeader(s));
  const has = name => h.includes(normHeader(name));
  const hasAny = (...names) => names.some(n => has(n));
  const sheetHas = re => (sheetNames || []).some(s => re.test(s));
  const cnt = h.length;

  // PickTrace bulk templates have starred required columns ("Site*", "Name*").
  // The presence of "is archived" as a header column is the giveaway between
  // CREATE (no Is Archived) and UPDATE (Is Archived at the end).
  if (sheetHas(/data.?entry/i) && has('site*') && has('name*')) {
    return has('is archived') ? 'updt' : 'tpl';
  }
  // Sites export — has Site + Site ID + Location + Location ID.
  if (has('site') && has('site id') && (has('location') || has('location id'))) return 'sites-master';
  // Locations to be removed — small file, Site + Location only, no Acreage.
  if (cnt <= 6 && has('site') && has('location') && !hasAny('acreage', 'acres')) return 'rm';
  // Locations to be added — has Block + Grower/Site + separate Comm + Variety.
  if (hasAny('block', 'block code') && hasAny('grower', 'site') && has('comm') && has('variety')) return 'add';
  // Database — Name/Block + Acreage/Acres, no DATA ENTRY sheet.
  if (!sheetHas(/data.?entry/i) && hasAny('name', 'block') && hasAny('acreage', 'acres')) return 'db';
  return null;
}

const SLOT_META_ID = {
  db: 'cmp-db-meta', add: 'cmp-add-meta', rm: 'cmp-rm-meta',
  tpl: 'cmp-tpl-meta', updt: 'cmp-updt-meta', 'sites-master': 'cmp-sites-mfile-meta'
};

// Slots whose file shape is structurally identical and therefore interchangeable
// at the fingerprint level. The Database (slot 1) and the Bulk Update Template
// (slot 6) are both PickTrace "ALL RECORDS" exports of the same DATA ENTRY
// sheet — there's no structural way to tell them apart.
const COMPATIBLE_SLOTS = {
  db: new Set(['updt']),
  updt: new Set(['db'])
};

function flagFileType(target, headers, sheetNames, wb) {
  // If a workbook is provided, prefer the DATA ENTRY sheet's headers over
  // whichever sheet was picked (the create template's DATA ENTRY is empty,
  // so picking falls through to DROP-DOWN INPUTS). If there's no DATA ENTRY
  // sheet at all, fall back to the first sheet that has data — this matters
  // when the wrong file is uploaded to a template slot (e.g., dropping the
  // Add file into slot 4) so the warning banner can still fingerprint it.
  let h = headers;
  if (wb && wb.SheetNames) {
    const dataEntry = wb.SheetNames.find(n => /data.?entry/i.test(n));
    if (dataEntry && wb.Sheets[dataEntry]) {
      try {
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[dataEntry], { header: 1, defval: '' });
        if (raw.length && raw[0] && raw[0].length) {
          h = raw[0].map(c => String(c == null ? '' : c).trim()).filter(Boolean);
        }
      } catch (e) { /* fall back below */ }
    }
    if (!h || !h.length) {
      for (const n of wb.SheetNames) {
        const ws = wb.Sheets[n];
        if (!ws) continue;
        try {
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (raw.length && raw[0] && raw[0].length) {
            h = raw[0].map(c => String(c == null ? '' : c).trim()).filter(Boolean);
            break;
          }
        } catch (e) {}
      }
    }
  }
  const detected = fingerprintFile(h, sheetNames);
  const el = $(SLOT_META_ID[target]);
  if (!el) return;
  // Strip any prior file-type banner so re-uploads show the latest verdict only.
  // The data-completeness banner (.cmp-archive-warn) is left alone — it's set by
  // renderDbMeta and represents a different concern.
  const prior = el.querySelector('.cmp-warn-banner:not(.cmp-archive-warn)');
  if (prior) prior.remove();
  if (!detected || detected === target) return; // match — nothing to flag
  const compat = COMPATIBLE_SLOTS[target];
  if (compat && compat.has(detected)) return; // structurally identical to expected — accept
  const det = FILE_TYPES[detected];
  const exp = FILE_TYPES[target];
  const msg = '<div class="cmp-warn-banner">&#9888; This looks like a <b>' + escHtml(det.label) +
    '</b> (slot ' + det.num + '), not a <b>' + escHtml(exp.label) + '</b> (slot ' + exp.num +
    '). Verify before running, or move it to slot ' + det.num + '.</div>';
  el.insertAdjacentHTML('afterbegin', msg);
}

function loadFileInto(file, target) {
  if (target === 'sites-master') {
    readUploadedFile(file).then(result => {
      const sheet = result.sheets && result.sheets[0];
      if (!sheet) { alert('Could not read sheet from sites file.'); return; }
      const parsed = parseSitesMaster(sheet.headers, sheet.rows, file.name);
      if (!parsed) {
        alert('Sites file is missing a "Site" column.');
        return;
      }
      sitesMasterData = parsed;
      $('cmp-sites-mfile-name').textContent = file.name;
      $('cmp-sites-mfile-meta').innerHTML =
        parsed.allLongNames.size + ' unique sites &middot; ' +
        parsed.archivedLongNames.size + ' fully-archived';
      flagFileType('sites-master', sheet.headers, result.sheets.map(s => s.name));
    }).catch(err => alert('Failed to read sites file: ' + err.message));
    return;
  }
  if (target === 'tpl') {
    // Keep the raw bytes so we can re-open the template fresh on each export
    // and preserve all original formatting + data validations.
    const reader = new FileReader();
    reader.onload = e => {
      const buf = e.target.result;
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
      const parsed = parseTemplateWorkbook(wb, file.name);
      if (!parsed || (!parsed.sites.length && !parsed.cropVarieties.length && !parsed.dataEntryHeaders.length)) {
        // Don't silently bail — surface the file in the slot and let the
        // fingerprint banner tell the user what they actually uploaded.
        $('cmp-tpl-name').textContent = file.name;
        $('cmp-tpl-meta').innerHTML = '<span class="cmp-warn">&#9888; Not a Bulk Create Template — no DATA ENTRY/DROP-DOWN INPUTS sheets found.</span>';
        flagFileType('tpl', null, wb.SheetNames, wb);
        return;
      }
      parsed.rawBuffer = buf;
      templateData = parsed;
      $('cmp-tpl-name').textContent = file.name;
      $('cmp-tpl-meta').innerHTML =
        parsed.dataEntryHeaders.length + ' columns &middot; ' +
        parsed.sites.length + ' sites &middot; ' +
        parsed.cropVarieties.length + ' crop/variety values';
      flagFileType('tpl', parsed.dataEntryHeaders, wb.SheetNames, wb);
      updateRunButton();
    };
    reader.readAsArrayBuffer(file);
    return;
  }
  if (target === 'updt') {
    const reader = new FileReader();
    reader.onload = e => {
      const buf = e.target.result;
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      const parsed = parseUpdateTemplate(wb, file.name);
      if (!parsed) {
        // Surface the file + fingerprint banner instead of an easy-to-miss alert.
        $('cmp-updt-name').textContent = file.name;
        $('cmp-updt-meta').innerHTML = '<span class="cmp-warn">&#9888; Not a Bulk Update Template — missing DATA ENTRY sheet or required columns (Site / Name / Crop &amp; Variety).</span>';
        flagFileType('updt', null, wb.SheetNames, wb);
        return;
      }
      parsed.rawBuffer = buf;
      updateTemplateData = parsed;
      $('cmp-updt-name').textContent = file.name;
      $('cmp-updt-meta').innerHTML =
        parsed.rows.length + ' rows &middot; ' +
        parsed.headers.length + ' columns &middot; ' +
        parsed.byKey.size + ' unique keys' +
        (parsed.dupCount ? ' &middot; <span class="cmp-warn">' + parsed.dupCount + ' duplicate key(s) — first wins</span>' : '');
      flagFileType('updt', parsed.headers, wb.SheetNames, wb);
      updateRunButton();
      // Recompute matches against any existing comparison and re-render.
      if (diffResult) {
        computeUpdateMatches();
        renderResults();
      }
    };
    reader.readAsArrayBuffer(file);
    return;
  }
  readFile(file).then(wb => {
    pickSheetIfNeeded(wb, file.name, sheet => {
      const data = { headers: sheet.headers, rows: sheet.rows, fileName: file.name, sheetName: sheet.name };
      if (target === 'db') {
        dbData = data;
        $('cmp-db-name').textContent = file.name;
        renderDbMeta();
      } else if (target === 'add') {
        addData = data;
        $('cmp-add-name').textContent = file.name;
        renderFileMeta('add', data);
      } else if (target === 'rm') {
        removeData = data;
        $('cmp-rm-name').textContent = file.name;
        renderFileMeta('rm', data);
      }
      flagFileType(target, sheet.headers, wb.SheetNames, wb);
      updateRunButton();
    });
  }).catch(err => alert('Failed to read file: ' + err.message));
}

function renderDbMeta() {
  if (!dbData) { $('cmp-db-meta').textContent = ''; return; }
  // Detect canonical PickTrace-style DB columns
  const H = dbData.headers;
  dbCols = {
    name:       findColIdx(H, ['name', 'block', 'block name', 'location', 'location name']),
    site:       findColIdx(H, ['site', 'site name', 'grower', 'farm']),
    altId:      findColIdx(H, ['alt id', 'altid', 'block code', 'code']),
    cropVar:    findColIdx(H, ['crop & variety', 'crop and variety', 'crop&variety']),
    acreage:    findColIdx(H, ['acreage', 'acres', 'acre', 'area']),
    plantCount: findColIdx(H, ['plant count', 'plantcount', 'tree count', 'treecount', 'trees', 'plants']),
    locType:    findColIdx(H, ['location type', 'type']),
    archived:   findColIdx(H, ['is archived', 'archived']),
    startDate:  findColIdx(H, ['start date', 'start_date'])
  };
  if (dbCols.name < 0) dbCols.name = detectColByContent(dbData.rows, isNumericCode);

  const meta = $('cmp-db-meta');
  if (dbCols.name < 0) {
    meta.innerHTML = '<span class="cmp-warn">&#9888; Could not identify a Block/Name column. Comparison may fail.</span>';
  } else {
    const lbl = (k, idx) => idx >= 0 ? '<b>' + escHtml(dbData.headers[idx]) + '</b>' : '<i>(none)</i>';
    meta.innerHTML = dbData.rows.length + ' rows &middot; Block col: ' + lbl('name', dbCols.name) +
      ' &middot; Grower col: ' + lbl('site', dbCols.site);
  }

  // Detect a non-ALL-RECORDS export. PickTrace exports identical column structure
  // whether the user filtered to active or selected ALL RECORDS — the only
  // difference is the presence of archived rows. Without them we can't:
  //   - route Adds that match an archived block into TO UNARCHIVE,
  //   - skip Removes that target already-archived blocks.
  // When this happens, reveal slot 6 (otherwise hidden) and add a warning so
  // the user knows the upload is incomplete.
  const slot6 = $('cmp-slot-updt');
  // Org-data imports don't carry archive flags — those records live only in
  // a PickTrace ALL RECORDS export. Skip the partial-export warning + the
  // required-slot-6 reveal when the DB came from the Org Data import flow.
  if (dbData._fromOrgImport) {
    if (slot6) slot6.style.display = 'none';
    if (updateTemplateData && updateTemplateData._fromDb) {
      updateTemplateData = null;
      $('cmp-updt-name').textContent = 'No file selected';
      $('cmp-updt-meta').innerHTML = '';
    }
    return;
  }
  if (dbCols.archived >= 0 && dbData.rows.length > 0) {
    const archivedCount = dbData.rows.filter(r => isArchived(r)).length;
    if (archivedCount === 0) {
      // Partial export — show slot 6 so the user can supply archived rows.
      // Clear any auto-derived updateTemplateData from a previous full DB.
      const warn = '<div class="cmp-warn-banner cmp-archive-warn">&#9888; <b>Warning:</b> no archived rows in this database. ' +
        'This looks like a partial export. Re-export with <b>Records = ALL RECORDS</b> in PickTrace, ' +
        'or upload the bulk update template into <b>slot 6</b> (now visible below) to supply the missing archived rows.</div>';
      meta.insertAdjacentHTML('afterbegin', warn);
      if (slot6) slot6.style.display = '';
      if (updateTemplateData && updateTemplateData._fromDb) {
        updateTemplateData = null;
        $('cmp-updt-name').textContent = 'No file selected';
        $('cmp-updt-meta').innerHTML = '';
      }
    } else {
      // Full ALL RECORDS export — hide slot 6 and auto-derive updateTemplateData
      // from the DB itself so the downstream TO UPDATE / DATA ENTRY (UPDATE) flow
      // works without a separate upload (the DB is structurally identical to the
      // bulk update template).
      if (slot6) slot6.style.display = 'none';
      // Drop any user-uploaded slot-6 file that was tied to the previous DB.
      if (updateTemplateData && !updateTemplateData._fromDb) {
        $('cmp-updt-name').textContent = 'No file selected';
        $('cmp-updt-meta').innerHTML = '';
      }
      const built = buildUpdateTemplateFromDb();
      if (built) {
        updateTemplateData = built;
      }
    }
  } else if (slot6) {
    // No archived column at all (probably the wrong file in slot 1) — keep slot 6 hidden.
    slot6.style.display = 'none';
  }
}

function renderFileMeta(which, data) {
  const id = 'cmp-' + which + '-meta';
  $(id).innerHTML = data.rows.length + ' rows &middot; ' + data.headers.length + ' cols';
}

function updateRunButton() {
  // Org-data imports + a template alone are a valid scenario: the user wants
  // to format the imported rows into the create/update template's column
  // structure and export. No Add/Remove files are required for that path.
  const hasAddOrRm = !!addData || !!removeData;
  const orgImportOnly = !!dbData && !!dbData._fromOrgImport && (!!templateData || !!updateTemplateData);
  const canRun = !!dbData && dbCols && dbCols.name >= 0 && (hasAddOrRm || orgImportOnly);
  $('cmp-run').disabled = !canRun;
  $('cmp-empty').style.display = dbData ? 'none' : '';
}

// ─── Comparison ───
// Returns true if a DB row is archived (Is Archived = TRUE) — excluded from matching.
function isArchived(r) {
  if (!dbCols || dbCols.archived < 0) return false;
  const v = String(r[dbCols.archived] || '').trim().toUpperCase();
  return v === 'TRUE' || v === '1' || v === 'YES';
}

function buildDbIndex() {
  const idx = new Map();
  dbData.rows.forEach((r, i) => {
    if (isArchived(r)) return;
    const code = extractCode(r[dbCols.name]);
    if (!code) return;
    if (!idx.has(code)) idx.set(code, { rowIndex: i, values: r });
  });
  return idx;
}

// Identify all relevant Add columns.
function detectAddColumns() {
  const H = addData.headers;
  let blockCol = findColIdx(H, ['block', 'block code', 'block name']);
  let growerCol = findColIdx(H, ['grower', 'site', 'site name', 'farm']);
  if (blockCol < 0) blockCol = detectColByContent(addData.rows, isNumericCode);
  if (growerCol < 0) growerCol = detectColByContent(addData.rows, c => isAlphaCode(c) && !isNumericCode(c));
  return {
    block:    blockCol,
    grower:   growerCol,
    district: findColIdx(H, ['district']),
    comm:     findColIdx(H, ['comm', 'commodity', 'crop']),
    variety:  findColIdx(H, ['variety', 'cultivar']),
    acres:    findColIdx(H, ['acres', 'acreage', 'acre']),
    trees:    findColIdx(H, ['trees', 'plant count', 'tree count'])
  };
}

// For a single Add row, normalize column-swap by content (numeric vs alpha).
// Returns a row array where block / grower are guaranteed correct.
function normalizeAddRow(row, addCols) {
  if (addCols.block < 0 || addCols.grower < 0) return row;
  const bCode = extractCode(row[addCols.block]);
  const gCode = extractCode(row[addCols.grower]);
  const bLooksGrower = isAlphaCode(bCode) && !isNumericCode(bCode);
  const gLooksBlock = isNumericCode(gCode);
  if (bLooksGrower && gLooksBlock) {
    const copy = row.slice();
    [copy[addCols.block], copy[addCols.grower]] = [copy[addCols.grower], copy[addCols.block]];
    return copy;
  }
  return row;
}

// Extract the meaningful fields from an Add row into a normalized object.
function extractAddFields(row, addCols) {
  const blockRaw = addCols.block >= 0 ? String(row[addCols.block] || '').trim() : '';
  const growerRaw = addCols.grower >= 0 ? String(row[addCols.grower] || '').trim() : '';
  const grower = parseCodeName(growerRaw);
  let acres = addCols.acres >= 0 ? String(row[addCols.acres] || '').trim() : '';
  let trees = addCols.trees >= 0 ? String(row[addCols.trees] || '').trim() : '';

  // Defensive correction: Plant Count is always > Acreage in real data.
  // If Acres > Trees and both are positive, they're swapped — fix it.
  const ac = parseFloat(String(acres).replace(/[, ]/g, ''));
  const tr = parseFloat(String(trees).replace(/[, ]/g, ''));
  let acresTreesSwapped = false;
  if (!isNaN(ac) && !isNaN(tr) && ac > 0 && tr > 0 && ac > tr) {
    [acres, trees] = [trees, acres];
    acresTreesSwapped = true;
  }

  return {
    blockRaw,
    growerRaw,
    blockCanonical: canonBlock(blockRaw),
    blockCode: extractCode(blockRaw),
    growerName: stripSiteCommas(grower.name),
    growerCode: grower.code,
    comm:    addCols.comm    >= 0 ? String(row[addCols.comm]    || '').trim() : '',
    variety: addCols.variety >= 0 ? String(row[addCols.variety] || '').trim() : '',
    acres,
    trees,
    acresTreesSwapped
  };
}

// Build Crop & Variety lookup maps from matched rows:
//   learned     : "comm|variety" -> full "Crop-Variety" string from the DB
//   perComm     : "comm"         -> "Crop" name (left side of dash)
//   perVariety  : "variety"      -> "Variety" name (right side of dash)
function learnCropVarietyMap(dbIndex, addCols) {
  const learned = new Map();
  const perComm = new Map();
  const perVariety = new Map();
  if (!addData || dbCols.cropVar < 0) return { learned, perComm, perVariety };
  addData.rows.forEach(rawRow => {
    const row = normalizeAddRow(rawRow, addCols);
    const f = extractAddFields(row, addCols);
    if (!f.blockCode || !f.comm || !f.variety) return;
    const dbRow = dbIndex.get(f.blockCode);
    if (!dbRow) return;
    const cv = String(dbRow.values[dbCols.cropVar] || '').trim();
    if (!cv) return;
    const key = (f.comm + '|' + f.variety).toUpperCase();
    if (!learned.has(key)) learned.set(key, cv);
    const dash = cv.indexOf('-');
    if (dash > 0) {
      const cropName = cv.substring(0, dash).trim();
      const varName = cv.substring(dash + 1).trim();
      const cu = f.comm.toUpperCase(), vu = f.variety.toUpperCase();
      if (!perComm.has(cu)) perComm.set(cu, cropName);
      if (!perVariety.has(vu)) perVariety.set(vu, varName);
    }
  });
  return { learned, perComm, perVariety };
}

// Resolve the Site value for an Add row. Prefer the bulk-template canonical form.
function resolveSite(fields, baseRow) {
  const canonical = lookupCanonicalSite({ code: fields.growerCode, name: fields.growerName });
  if (canonical) return canonical;
  if (fields.growerName) return fields.growerName;
  if (baseRow && dbCols.site >= 0) return baseRow[dbCols.site] || '';
  return fields.growerCode || '';
}

// Build a DB-shaped row for an Add entry. Only fills columns we know how to populate.
// For conflicts (existing dbRow), uses dbRow as the base and overrides the known fields.
function buildDbRowFromAdd(fields, cvMaps, baseRow) {
  const { learned, perComm, perVariety } = cvMaps;
  const out = baseRow ? baseRow.slice() : new Array(dbData.headers.length).fill('');
  if (dbCols.name >= 0)       out[dbCols.name]       = fields.blockCanonical;
  if (dbCols.site >= 0)       out[dbCols.site]       = resolveSite(fields, baseRow);
  if (dbCols.altId >= 0)      out[dbCols.altId]      = fields.growerCode || (baseRow ? baseRow[dbCols.altId] : '');
  if (dbCols.acreage >= 0)    out[dbCols.acreage]    = fields.acres;
  if (dbCols.plantCount >= 0) out[dbCols.plantCount] = fields.trees;
  if (dbCols.cropVar >= 0 && fields.comm && fields.variety) {
    // Strict: only template-validated values reach the export. Unmapped combos
    // leave the cell blank — surfaced in the "Crops & Varieties to Create" panel.
    const cv = lookupCanonicalCropVariety(fields.comm, fields.variety, learned, perComm, perVariety, true);
    if (cv) out[dbCols.cropVar] = cv;
    else if (!baseRow) out[dbCols.cropVar] = '';
  }
  if (!baseRow) {
    if (dbCols.locType >= 0 && !out[dbCols.locType]) out[dbCols.locType] = 'Block';
    if (dbCols.archived >= 0 && !out[dbCols.archived]) out[dbCols.archived] = 'FALSE';
    if (dbCols.startDate >= 0 && !out[dbCols.startDate]) out[dbCols.startDate] = todayIsoDate();
  }
  return out;
}

// YYYY-MM-DD format (matches PickTrace's Start Date format).
function todayIsoDate() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

// Compare a DB row against an Add fields-object. Returns list of DB column indices
// that count as a TRUE conflict — the ones the user must resolve manually.
//
// Auto-applied (NEVER conflicts, always overwritten by buildDbRowFromAdd):
//   - Acreage   (always take new)
//   - Plant Count (always take new)
//   - Crop & Variety (take dropdown-canonical synthesized value)
//   - Empty DB cells where Add has a value (just fill it in)
//
// Real conflicts only fire when the DB has a non-empty value AND it disagrees
// with a non-empty incoming value on:
//   - Site  (different grower name)
//   - Name  (different block name)
//   - Alt ID (different code)
function diffDbVsAdd(dbRow, fields /*, cvMaps */) {
  const diffs = [];
  const compareIfBoth = (idx, newVal) => {
    if (idx < 0) return;
    const dbVal = dbRow[idx] != null ? String(dbRow[idx]).trim() : '';
    const incomingVal = newVal != null ? String(newVal).trim() : '';
    if (dbVal === '') return;       // auto-fill empty DB cells — not a conflict
    if (incomingVal === '') return; // incoming empty — nothing to compare
    if (!valuesEqual(dbVal, incomingVal)) diffs.push(idx);
  };

  // Block name: compare canonical forms (strips parens-spacing differences).
  if (dbCols.name >= 0) {
    const dbName = canonBlock(dbRow[dbCols.name]);
    const newName = fields.blockCanonical;
    if (dbName && newName && !valuesEqual(dbName, newName)) {
      diffs.push(dbCols.name);
    }
  }

  // Site: compare LONG NAMES only (strips "CODE(NAME)" vs "NAME (CODE)" vs bare "NAME" formatting).
  if (dbCols.site >= 0 && fields.growerName) {
    const dbSiteLong = siteLongName(dbRow[dbCols.site]);
    if (dbSiteLong && !valuesEqual(dbSiteLong, fields.growerName)) {
      diffs.push(dbCols.site);
    }
  }

  if (fields.growerCode) compareIfBoth(dbCols.altId, fields.growerCode);
  // Acreage / Plant Count / Crop & Variety: never compared — always auto-applied.
  return diffs;
}

function processAdds(dbIndex) {
  const newAdds = [];
  const conflicts = [];
  const silentUpdates = [];     // auto-applied changes that don't need user review
  const toUnarchive = [];       // Add row matches an ARCHIVED DB row → user must unarchive in PickTrace
  let unchanged = 0;
  let dupAddRows = 0;
  let acresTreesSwapped = 0;
  if (!addData) return { newAdds, conflicts, silentUpdates, toUnarchive, unchanged, dupAddRows, acresTreesSwapped };

  const addCols = detectAddColumns();
  if (addCols.block < 0) {
    alert('Could not identify a Block column in the Add file. Aborting Add processing.');
    return { newAdds, conflicts, unchanged, dupAddRows };
  }

  const cvMaps = learnCropVarietyMap(dbIndex, addCols);
  const compositeIdx = buildDbCompositeIndex();
  const siteCodeIdx = buildDbSiteCodeIndex();
  const archivedSiteCodeIdx = buildArchivedSiteCodeIndex();

  // Dedupe Add file by block code — keep the LAST occurrence (assumes later rows are corrections).
  const lastByCode = new Map();
  addData.rows.forEach((rawRow, i) => {
    const row = normalizeAddRow(rawRow, addCols);
    const fields = extractAddFields(row, addCols);
    if (!fields.blockCode) return;
    if (fields.acresTreesSwapped) acresTreesSwapped++;
    if (lastByCode.has(fields.blockCode)) dupAddRows++;
    lastByCode.set(fields.blockCode, { fields, addRowIndex: i });
  });

  lastByCode.forEach(({ fields, addRowIndex }) => {
    const cropVar = lookupCanonicalCropVariety(
      fields.comm, fields.variety, cvMaps.learned, cvMaps.perComm, cvMaps.perVariety);
    const expectedComposite = makeCompositeKey({
      code: fields.blockCode,
      siteName: fields.growerName,
      cropVar,
      acreage: fields.acres,
      plantCount: fields.trees
    });
    // 1. Exact composite match → unchanged.
    if (compositeIdx.has(expectedComposite)) {
      unchanged++;
      return;
    }
    // 2. Same Site + Code → either a true conflict (Name/Site/Alt ID disagrees)
    //    or a silent auto-update (only Acreage/PlantCount/CropVariety/empty cells differ).
    const sk = makeSiteCodeKey({ code: fields.blockCode, siteName: fields.growerName });
    const siteHits = siteCodeIdx.get(sk);
    if (siteHits && siteHits.length) {
      siteHits.forEach(hit => {
        const diffFields = diffDbVsAdd(hit.values, fields);
        const newValues = buildDbRowFromAdd(fields, cvMaps, hit.values);
        if (diffFields.length === 0) {
          // No "real" disagreement on Site/Name/Alt ID — auto-apply.
          // If newValues differ from dbValues somewhere (acreage, plant count, etc.),
          // that's a silent update; otherwise truly unchanged.
          let anyDiff = false;
          for (let i = 0; i < dbData.headers.length; i++) {
            const a = String(hit.values[i] == null ? '' : hit.values[i]).trim();
            const b = String(newValues[i] == null ? '' : newValues[i]).trim();
            if (!valuesEqual(a, b)) { anyDiff = true; break; }
          }
          if (anyDiff) {
            silentUpdates.push({
              blockCode: fields.blockCode,
              dbRowIndex: hit.rowIndex,
              dbValues: hit.values.slice(),
              newValues,
              fields,
              addRowIndex
            });
          } else {
            unchanged++;
          }
        } else {
          conflicts.push({
            blockCode: fields.blockCode,
            dbRowIndex: hit.rowIndex,
            dbValues: hit.values.slice(),
            newValues,
            diffFields,
            fields,
            addRowIndex
          });
        }
      });
      return;
    }
    // 3. Add references an ARCHIVED DB row → divert to "TO UNARCHIVE" bucket.
    //    These rows are NOT written to DATA ENTRY (would create duplicates in PickTrace);
    //    instead they're surfaced as a checklist for the user to manually unarchive.
    const archivedHits = archivedSiteCodeIdx.get(sk);
    if (archivedHits && archivedHits.length) {
      archivedHits.forEach(hit => {
        const newRow = buildDbRowFromAdd(fields, cvMaps, hit.values);
        toUnarchive.push({
          blockCode: fields.blockCode,
          dbRowIndex: hit.rowIndex,
          dbValues: hit.values.slice(),
          newRow,
          fields,
          addRowIndex
        });
      });
      return;
    }

    // 4. Otherwise → genuinely new (different Site or no DB entry at all).
    const newRow = buildDbRowFromAdd(fields, cvMaps, null);
    newAdds.push({
      blockCode: fields.blockCode,
      row: newRow,
      fields,
      addRowIndex,
      previouslyArchived: false
    });
  });

  return { newAdds, conflicts, silentUpdates, toUnarchive, unchanged, dupAddRows, acresTreesSwapped, cvMaps };
}

// Build a multi-index of DB rows: blockCode -> [{ rowIndex, values }, ...]
// Skips archived rows (Is Archived = TRUE) so they don't participate in matching.
function buildDbMultiIndex() {
  const m = new Map();
  dbData.rows.forEach((r, i) => {
    if (isArchived(r)) return;
    const c = extractCode(r[dbCols.name]);
    if (!c) return;
    if (!m.has(c)) m.set(c, []);
    m.get(c).push({ rowIndex: i, values: r });
  });
  return m;
}

// ─── Composite-identity helpers ───
// Two DB rows with the same block code can be different physical blocks
// (different Site, different acreage, etc.). Identity = composite of fields.
function siteLongName(siteVal) {
  const parts = parseCodeName(siteVal);
  const raw = parts.name || String(siteVal == null ? '' : siteVal).trim();
  return stripSiteCommas(raw);
}

// Normalize a Site name to the office's canonical convention:
//   - Strip ALL commas        ("CROWN FARMING, INC."  → "CROWN FARMING INC.")
//   - Strip TRAILING periods  ("CROWN FARMING INC."   → "CROWN FARMING INC")
//   - Internal periods preserved (initials like "B. KIRSCHENMANN" stay intact)
//   - Collapse internal whitespace + trim edges
//
// Both sides of every comparison (Add Grower, DB Site, template Sites, sites
// master CSV) flow through this helper, so trailing-period mismatches like
// "B. KIRSCHENMANN FARMS, INC." vs "B. KIRSCHENMANN FARMS, INC" reconcile.
function stripSiteCommas(s) {
  return String(s == null ? '' : s)
    .replace(/,/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\.+$/, '')
    .trim();
}

function normVal(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function normNumVal(v) {
  if (v == null || v === '') return '';
  const s = String(v).replace(/[, ]/g, '').trim();
  if (s === '') return '';
  const n = parseFloat(s);
  return isNaN(n) ? s.toUpperCase() : String(n);
}

function makeCompositeKey(parts) {
  return [
    normVal(parts.code),
    normVal(parts.siteName),
    normVal(parts.cropVar),
    normNumVal(parts.acreage),
    normNumVal(parts.plantCount)
  ].join('||');
}

function makeSiteCodeKey(parts) {
  return normVal(parts.siteName) + '||' + normVal(parts.code);
}

function dbRowComposite(r) {
  return makeCompositeKey({
    code: extractCode(r[dbCols.name]),
    siteName: siteLongName(r[dbCols.site]),
    cropVar: r[dbCols.cropVar],
    acreage: r[dbCols.acreage],
    plantCount: r[dbCols.plantCount]
  });
}

function dbRowSiteCode(r) {
  return makeSiteCodeKey({
    code: extractCode(r[dbCols.name]),
    siteName: siteLongName(r[dbCols.site])
  });
}

function buildDbCompositeIndex() {
  const m = new Map();
  dbData.rows.forEach((r, i) => {
    if (isArchived(r)) return;
    const k = dbRowComposite(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push({ rowIndex: i, values: r });
  });
  return m;
}

function buildDbSiteCodeIndex() {
  const m = new Map();
  dbData.rows.forEach((r, i) => {
    if (isArchived(r)) return;
    const code = extractCode(r[dbCols.name]);
    if (!code) return;
    const k = dbRowSiteCode(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push({ rowIndex: i, values: r });
  });
  return m;
}

// Same as buildDbSiteCodeIndex but ONLY archived rows — for re-activation lookups.
function buildArchivedSiteCodeIndex() {
  const m = new Map();
  dbData.rows.forEach((r, i) => {
    if (!isArchived(r)) return;
    const code = extractCode(r[dbCols.name]);
    if (!code) return;
    const k = dbRowSiteCode(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push({ rowIndex: i, values: r });
  });
  return m;
}

function processRemoves(dbIndex, addCovered) {
  const removes = [];
  const skipped = [];
  if (!removeData) return { removes, skipped };
  // addCovered: Set of codes that the Add file is already handling. Removes that
  // target the same code should be treated as a no-op (the Add represents an UPDATE,
  // not a delete). Otherwise we'd lose the user's update.
  addCovered = addCovered || new Set();

  let locCol = findColIdx(removeData.headers, ['location', 'block', 'block name']);
  let siteCol = findColIdx(removeData.headers, ['site', 'grower', 'site name']);
  if (locCol < 0) locCol = detectColByContent(removeData.rows, isNumericCode);
  if (siteCol < 0 && locCol !== 0) siteCol = 0;

  if (locCol < 0) {
    alert('Could not identify a Location column in the Remove file.');
    return { removes, skipped };
  }

  // Two-pass parse:
  //   detail rows  = Site + specific Location (e.g., "33650 (DEL REY MORO BLD #1)")
  //   summary rows = Site + "{N}" — count-summary; semantics depend on whether
  //                  detail rows exist for the same Site.
  const detailsBySite = new Map();    // siteNormUpper -> [{ row, locRaw, siteRaw }]
  const summariesBySite = new Map();  // siteNormUpper -> { n, siteRaw, locRaw }

  removeData.rows.forEach(row => {
    const locRaw = row[locCol] != null ? String(row[locCol]).trim() : '';
    const siteRaw = siteCol >= 0 ? String(row[siteCol] || '').trim() : '';
    const sNorm = normVal(siteLongName(siteRaw));

    if (/^\{\d+\}$/.test(locRaw)) {
      const n = parseInt(locRaw.slice(1, -1), 10);
      summariesBySite.set(sNorm, { n, siteRaw, locRaw });
      return;
    }
    if (!locRaw) {
      skipped.push({ site: siteRaw, location: locRaw, reason: 'empty Location cell' });
      return;
    }
    if (!detailsBySite.has(sNorm)) detailsBySite.set(sNorm, []);
    detailsBySite.get(sNorm).push({ locRaw, siteRaw });
  });

  // Match by (Site long name + Block code). Falls back to code-only when no Site provided.
  const siteCodeIdx = buildDbSiteCodeIndex();
  const codeIdx = buildDbMultiIndex();
  const archivedSiteCodeIdx = buildArchivedSiteCodeIndex();

  // ─── Pass 1: process detail rows ───
  detailsBySite.forEach((rows, sNorm) => {
    rows.forEach(({ locRaw, siteRaw }) => {
      const code = extractCode(locRaw);
      if (!code) {
        skipped.push({ site: siteRaw, location: locRaw, reason: 'could not extract code' });
        return;
      }
      // If the Add file already covers this code, the Remove is part of an
      // update pattern (delete-then-recreate) — skip the remove.
      if (addCovered.has(code)) {
        skipped.push({ site: siteRaw, location: locRaw, reason: 'superseded by Add file entry — treated as update' });
        return;
      }
      const siteName = siteLongName(siteRaw);
      let hits = null;
      let matchKind = '';
      if (siteName) {
        hits = siteCodeIdx.get(makeSiteCodeKey({ code, siteName }));
        if (hits && hits.length) matchKind = 'site+code';
      }
      if (!hits || !hits.length) {
        const codeHits = codeIdx.get(code);
        if (codeHits && codeHits.length) {
          hits = codeHits;
          matchKind = 'code-only';
        }
      }
      if (!hits || !hits.length) {
        // Last check: does it match an ARCHIVED DB row? If so, no action needed.
        const siteName = siteLongName(siteRaw);
        const archHits = siteName ? archivedSiteCodeIdx.get(makeSiteCodeKey({ code, siteName })) : null;
        if (archHits && archHits.length) {
          skipped.push({ site: siteRaw, location: locRaw, reason: 'already archived in DB — no action needed' });
        } else {
          skipped.push({ site: siteRaw, location: locRaw, reason: 'no DB match for site+code ' + code });
        }
        return;
      }
      hits.forEach(h => {
        removes.push({
          blockCode: code,
          dbRowIndex: h.rowIndex,
          dbValues: h.values.slice(),
          site: siteRaw,
          location: locRaw,
          matchKind
        });
      });
    });
  });

  // ─── Pass 2: process {N} summaries ───
  summariesBySite.forEach((summary, sNorm) => {
    const detailRows = detailsBySite.get(sNorm) || [];
    if (detailRows.length > 0) {
      // Count-check: if {N} doesn't match the detailed-row count, surface a warning.
      if (detailRows.length !== summary.n) {
        skipped.push({
          site: summary.siteRaw,
          location: summary.locRaw,
          reason: 'count mismatch: file says {' + summary.n + '} but ' + detailRows.length + ' detailed remove rows present for this Site'
        });
      }
      // Detail rows already processed in pass 1 — nothing more to do.
      return;
    }

    // No detail rows under this Site: treat {N} as "remove all DB rows for this Site".
    const matchingDb = [];
    dbData.rows.forEach((r, i) => {
      const sn = normVal(siteLongName(r[dbCols.site]));
      if (sn === sNorm) matchingDb.push({ rowIndex: i, values: r });
    });
    if (matchingDb.length === 0) {
      skipped.push({
        site: summary.siteRaw,
        location: summary.locRaw,
        reason: '{' + summary.n + '} for site with no DB rows matching ' + summary.siteRaw
      });
      return;
    }
    if (matchingDb.length !== summary.n) {
      skipped.push({
        site: summary.siteRaw,
        location: summary.locRaw,
        reason: 'count mismatch: file says {' + summary.n + '} but DB has ' + matchingDb.length + ' rows for this Site (removing all)'
      });
    }
    matchingDb.forEach(h => {
      removes.push({
        blockCode: extractCode(h.values[dbCols.name]) || '*',
        dbRowIndex: h.rowIndex,
        dbValues: h.values.slice(),
        site: summary.siteRaw,
        location: '{all for site}',
        matchKind: 'site-bulk'
      });
    });
  });

  return { removes, skipped };
}

function runCompare() {
  if (!dbData || !dbCols || dbCols.name < 0) { alert('Upload a DB file first.'); return; }
  const dbIndex = buildDbIndex();

  const adds = processAdds(dbIndex);

  // Org-import + no Add file scenario: treat every imported DB row as a
  // synthetic "new add" so the UI counters and the Sites/Crops-to-Create
  // panels reflect what the export will actually emit. Without this the
  // user sees "0 new (DATA ENTRY)" even though all 438 rows will be in the export.
  if (dbData._fromOrgImport && !addData) {
    const cropVarIdx = dbCols.cropVar;
    dbData.rows.forEach((r, ri) => {
      const blockRaw = dbCols.name >= 0 ? String(r[dbCols.name] || '').trim() : '';
      const growerRaw = dbCols.site >= 0 ? String(r[dbCols.site] || '').trim() : '';
      const blockCode = extractCode(blockRaw);
      const grower = parseCodeName(growerRaw);
      const cvRaw = cropVarIdx >= 0 ? String(r[cropVarIdx] || '').trim() : '';
      let comm = '', variety = '';
      if (cvRaw && cvRaw.indexOf('-') > 0) {
        const dash = cvRaw.indexOf('-');
        comm = cvRaw.slice(0, dash).trim();
        variety = cvRaw.slice(dash + 1).trim();
      }
      adds.newAdds.push({
        blockCode,
        row: r.slice(),
        fields: {
          blockRaw,
          growerRaw,
          blockCanonical: blockRaw,
          blockCode,
          growerName: stripSiteCommas(grower.name || growerRaw),
          growerCode: grower.code || '',
          comm,
          variety,
          acres: dbCols.acreage >= 0 ? String(r[dbCols.acreage] || '') : '',
          trees: dbCols.plantCount >= 0 ? String(r[dbCols.plantCount] || '') : '',
          acresTreesSwapped: false
        },
        addRowIndex: ri,
        previouslyArchived: false
      });
    });
    console.log('[Block Compare] Org-import + no Add — synthesized ' + adds.newAdds.length + ' new-add entries from DB rows.');
  }
  // Codes the Add file is handling (as updates / new / unarchive). Any Remove
  // instruction for the same code represents the "delete-then-recreate" update
  // pattern — the Add already covers it via the TO UPDATE sheet, so the Remove
  // is suppressed from TO ARCHIVE.
  const addCovered = new Set();
  if (addData) {
    const addCols = detectAddColumns();
    addData.rows.forEach(rawRow => {
      const row = normalizeAddRow(rawRow, addCols);
      const f = extractAddFields(row, addCols);
      if (f.blockCode) addCovered.add(f.blockCode);
    });
  }
  const rms = processRemoves(dbIndex, addCovered);

  diffResult = {
    newAdds: adds.newAdds,
    conflicts: adds.conflicts,
    silentUpdates: adds.silentUpdates || [],
    toUnarchive: adds.toUnarchive || [],
    unchanged: adds.unchanged,
    dupAddRows: adds.dupAddRows || 0,
    acresTreesSwapped: adds.acresTreesSwapped || 0,
    autoSwappedDbRows: 0,
    removes: rms.removes,
    skipped: rms.skipped,
    sitesToCreate: computeSitesToCreate(adds.newAdds, adds.toUnarchive),
    cropVarsToCreate: computeCropVarietiesToCreate(adds.newAdds, adds.silentUpdates, adds.conflicts, adds.cvMaps),
    cvMaps: adds.cvMaps
  };
  diffResult.acresReversed = detectReversedAcresInDb();
  console.log('[Block Compare] reversed-acres detected:', diffResult.acresReversed.length,
              '| sample:', diffResult.acresReversed.slice(0, 3));
  // Auto-swap immediately — Plant Count is ALWAYS > Acreage in real data, so any
  // reversal is fixed without requiring a user click. The same swap mirrors into
  // every snapshot bucket so the preview, removes, and export all stay consistent.
  if (diffResult.acresReversed.length > 0 && dbCols.acreage >= 0 && dbCols.plantCount >= 0) {
    const list = diffResult.acresReversed;
    const reversedSet = new Set(list.map(x => x.dbRowIndex));
    const doSwap = arr => {
      const t = arr[dbCols.acreage];
      arr[dbCols.acreage] = arr[dbCols.plantCount];
      arr[dbCols.plantCount] = t;
    };
    list.forEach(item => { const r = dbData.rows[item.dbRowIndex]; if (r) doSwap(r); });
    (diffResult.silentUpdates || []).forEach(s => { if (reversedSet.has(s.dbRowIndex)) doSwap(s.dbValues); });
    (diffResult.conflicts     || []).forEach(c => { if (reversedSet.has(c.dbRowIndex)) doSwap(c.dbValues); });
    (diffResult.removes       || []).forEach(rm => { if (reversedSet.has(rm.dbRowIndex)) doSwap(rm.dbValues); });
    diffResult.autoSwappedDbRows = list.length;
    diffResult.acresReversed = [];
    console.log('[Block Compare] Auto-swapped Acreage <-> Plant Count on ' + list.length + ' rows.');
  }

  computeUpdateMatches();
  renderResults();
  $('cmp-export').disabled = false;
  $('cmp-debug').disabled = false;

  // Always expose the latest diff on window for console inspection.
  window.cmpDiffResult = diffResult;
  window.cmpDbData = dbData;
  window.cmpAddData = addData;
  window.cmpRemoveData = removeData;
  window.cmpTemplateData = templateData;
  window.cmpUpdateTemplateData = updateTemplateData;
  console.log('[Block Compare] diffResult on window.cmpDiffResult — click "Debug Dump" to download full JSON.');
}

// ─── Rendering ───
function renderResults() {
  const r = diffResult;
  const sum = $('cmp-summary');
  sum.style.display = '';
  const sitesToCreateCount = (r.sitesToCreate || []).length;
  const cropVarsToCreateCount = (r.cropVarsToCreate || []).length;
  const reversedCount = (r.acresReversed || []).length;
  const swappedCount = r.autoSwappedDbRows || 0;
  const toUnarchiveCount = (r.toUnarchive || []).length;
  const toUpdateCount = (r.silentUpdates || []).length + (r.conflicts || []).length;
  sum.innerHTML =
    '<div class="cmp-stat"><b>' + r.newAdds.length + '</b> new (DATA ENTRY)</div>' +
    '<div class="cmp-stat"><b>' + toUpdateCount + '</b> to update</div>' +
    '<div class="cmp-stat"><b>' + r.conflicts.length + '</b> conflicts</div>' +
    (toUnarchiveCount > 0 ? '<div class="cmp-stat cmp-warn"><b>' + toUnarchiveCount + '</b> to unarchive (already in PickTrace)</div>' : '') +
    '<div class="cmp-stat"><b>' + r.unchanged + '</b> unchanged</div>' +
    '<div class="cmp-stat"><b>' + r.removes.length + '</b> to archive</div>' +
    '<div class="cmp-stat"><b>' + r.skipped.length + '</b> skipped</div>' +
    (sitesToCreateCount > 0 ? '<div class="cmp-stat cmp-warn"><b>' + sitesToCreateCount + '</b> Sites need creation in PickTrace</div>' : '') +
    (cropVarsToCreateCount > 0 ? '<div class="cmp-stat cmp-warn"><b>' + cropVarsToCreateCount + '</b> Crops &amp; Varieties need creation in PickTrace</div>' : '') +
    (r.dupAddRows > 0 ? '<div class="cmp-stat cmp-warn"><b>' + r.dupAddRows + '</b> duplicate Add rows merged</div>' : '') +
    (r.acresTreesSwapped > 0 ? '<div class="cmp-stat cmp-warn"><b>' + r.acresTreesSwapped + '</b> Add-file Acres/Trees auto-swapped</div>' : '') +
    (reversedCount > 0
      ? '<div class="cmp-stat cmp-warn-strong"><b>&#9888; ' + reversedCount + '</b> rows have Acreage &gt; Plant Count <button class="btn btn-warning btn-sm" id="cmp-auto-swap" style="margin-left:8px;">Auto-swap all</button></div>'
      : '') +
    (swappedCount > 0
      ? '<div class="cmp-stat" style="color:var(--green-text);">&#10003; auto-swapped ' + swappedCount + ' Acreage/Plant Count rows</div>'
      : '') +
    (templateData ? '<div class="cmp-stat" style="color:var(--green-text);">&#10003; template loaded</div>' : '');

  // Wire up the auto-swap button (re-rendered each runCompare).
  const swapBtn = document.getElementById('cmp-auto-swap');
  if (swapBtn) swapBtn.addEventListener('click', applyAcresSwap);

  renderAdds();
  renderConflicts();
  renderRemoves();
  renderRequiredDefaults();
  renderCmpDateFill();
  renderSitesToCreate();
  renderCropVarsToCreate();
  renderUpdateMissing();
  renderArchiveMissing();
  renderSkipped();
}

// ─── Year-only date expansion control (applied at export) ───
function renderCmpDateFill() {
  const sec = $('cmp-section-datefill');
  if (!sec) return;
  const n = cmpYearOnlyDateCount();
  // Only surface the control when there are year-only date cells (or a setting
  // is already active), so it stays out of the way otherwise.
  if (!n && !(cmpDateFill.mm && cmpDateFill.dd)) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  const cnt = $('cmp-datefill-count');
  if (cnt) cnt.innerHTML = n
    ? '<b style="color:#b45309;">' + n + '</b> year-only date cell' + (n === 1 ? '' : 's') + ' in New Adds'
    : '<span style="color:#15803d;font-weight:600;">&#10003; Will expand on export</span>';
  const mmEl = $('cmp-datefill-mm'), ddEl = $('cmp-datefill-dd'), rgEl = $('cmp-datefill-range');
  if (mmEl && document.activeElement !== mmEl) mmEl.value = cmpDateFill.mm;
  if (ddEl && document.activeElement !== ddEl) ddEl.value = cmpDateFill.dd;
  if (rgEl) rgEl.value = cmpDateFill.range;
}

function applyCmpDateFill() {
  const mm = ($('cmp-datefill-mm').value || '').trim();
  const dd = ($('cmp-datefill-dd').value || '').trim();
  const range = $('cmp-datefill-range').value || 'first';
  const mi = parseInt(mm, 10), di = parseInt(dd, 10);
  if (!(mi >= 1 && mi <= 12) || !(di >= 1 && di <= 31)) {
    alert('Enter a valid month (1–12) and day (1–31) first.');
    return;
  }
  cmpDateFill = { mm, dd, range };
  renderCmpDateFill();
}

function clearCmpDateFill() {
  cmpDateFill = { mm: '', dd: '', range: ($('cmp-datefill-range') ? $('cmp-datefill-range').value : 'first') };
  if ($('cmp-datefill-mm')) $('cmp-datefill-mm').value = '';
  if ($('cmp-datefill-dd')) $('cmp-datefill-dd').value = '';
  renderCmpDateFill();
}

function renderCropVarsToCreate() {
  const sec = $('cmp-section-cropvars');
  const list = (diffResult.cropVarsToCreate || []);
  if (!list.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  $('cmp-cropvars-title').textContent = 'Crops & Varieties to Create in PickTrace (' + list.length + ')';
  let html = '<thead><tr><th>Suggested name</th><th>Comm code</th><th>Variety code</th><th>Blocks</th><th>Example block names</th></tr></thead><tbody>';
  list.forEach(c => {
    html += '<tr>' +
      '<td><b>' + escHtml(c.value) + '</b></td>' +
      '<td><code>' + escHtml(c.comm || '') + '</code></td>' +
      '<td><code>' + escHtml(c.variety || '') + '</code></td>' +
      '<td>' + c.count + '</td>' +
      '<td>' + c.examples.map(e => escHtml(e)).join('<br>') + '</td>' +
      '</tr>';
  });
  html += '</tbody>';
  $('cmp-cropvars-table').innerHTML = html;
}

function applyManualCropVariety(comm, variety, value) {
  if (!value || dbCols.cropVar < 0) return 0;
  const cu = String(comm || '').toUpperCase();
  const vu = String(variety || '').toUpperCase();
  let filled = 0;
  const fill = (row, fields) => {
    if (!fields) return;
    if (String(fields.comm || '').toUpperCase() !== cu) return;
    if (String(fields.variety || '').toUpperCase() !== vu) return;
    const cur = row[dbCols.cropVar];
    if (cur != null && String(cur).trim() !== '') return; // only fill empties
    row[dbCols.cropVar] = value;
    filled++;
  };
  (diffResult.newAdds || []).forEach(a => fill(a.row, a.fields));
  (diffResult.silentUpdates || []).forEach(s => fill(s.newValues, s.fields));
  (diffResult.conflicts || []).forEach(c => fill(c.newValues, c.fields));
  return filled;
}

function copyCropVarsList() {
  const list = (diffResult && diffResult.cropVarsToCreate) || [];
  if (!list.length) return;
  const lines = list.map(c => c.value + '\t' + (c.comm || '') + '\t' + (c.variety || '') + '\t' + c.count + ' blocks');
  const text = 'Crop & Variety\tComm code\tVariety code\tBlocks\n' + lines.join('\n');
  navigator.clipboard.writeText(text).then(
    () => alert('Copied ' + list.length + ' Crop & Variety values to clipboard.'),
    () => alert('Copy failed — clipboard access denied.')
  );
}

// Detect required (*) columns with empty cells, grouped (for Crop & Variety
// the grouping is by Comm+Variety code combo so the user can resolve each
// combo with a single dropdown choice; for other columns it's one group).
function detectEmptyRequiredColumns() {
  if (!dbData || !dbCols) return [];
  const reqCols = [];
  dbData.headers.forEach((h, i) => {
    if (/\*$/.test(String(h || '').trim())) reqCols.push({ idx: i, header: h });
  });
  const skipKeys = new Set(['name', 'is archived']);
  // buckets: each entry references the actual row by reference so we can mutate.
  // For TO UPDATE rows we also carry the matched update-template row so the
  // emptiness check can see PickTrace's existing value.
  const buckets = [
    ...(diffResult.newAdds || []).map(a => ({ row: a.row, fields: a.fields, tplRow: null })),
    ...(diffResult.silentUpdates || []).map(s => ({ row: s.newValues, fields: s.fields, tplRow: s._updateTplRow || null })),
    ...(diffResult.conflicts || []).map(c => ({ row: c.newValues, fields: c.fields, tplRow: c._updateTplRow || null }))
  ];
  const tplHeaderIdx = updateTemplateData ? updateTemplateData.headerIdx : null;
  return reqCols.filter(c => !skipKeys.has(normHeader(c.header)))
    .map(col => {
      const norm = normHeader(col.header);
      const isCV = norm === 'crop & variety';
      const emptyRows = buckets.filter(b => {
        const v = b.row[col.idx];
        if (v != null && String(v).trim() !== '') return false;
        // Cell is empty in the input. If a matched update-template row provides
        // a value at the corresponding header, treat it as filled.
        if (b.tplRow && tplHeaderIdx && tplHeaderIdx.has(norm)) {
          const tv = b.tplRow[tplHeaderIdx.get(norm)];
          if (tv != null && String(tv).trim() !== '') return false;
        }
        return true;
      });
      if (emptyRows.length === 0) return null;
      const groups = [];
      if (isCV) {
        const byCombo = new Map();
        emptyRows.forEach(b => {
          const comm = (b.fields && b.fields.comm) || '?';
          const variety = (b.fields && b.fields.variety) || '?';
          const key = comm + '|' + variety;
          if (!byCombo.has(key)) byCombo.set(key, { key, comm, variety, label: comm + ' + ' + variety, rows: [] });
          byCombo.get(key).rows.push(b);
        });
        byCombo.forEach(g => groups.push({
          key: g.key, label: g.label, comm: g.comm, variety: g.variety,
          count: g.rows.length, rows: g.rows
        }));
        groups.sort((a, b) => a.label.localeCompare(b.label));
      } else {
        groups.push({ key: 'all', label: 'all rows', count: emptyRows.length, rows: emptyRows });
      }
      return { idx: col.idx, header: col.header, norm, isCV, totalEmpty: emptyRows.length, groups };
    })
    .filter(Boolean);
}

// Get dropdown options (array of strings) for a column, sourced from the loaded template.
function getDropdownOptions(colHeader) {
  if (!templateData) return null;
  const norm = normHeader(colHeader);
  if (norm === 'crop & variety' && templateData.cropVarieties) return templateData.cropVarieties;
  if (norm === 'site' && templateData.sites) return templateData.sites;
  if (templateData.dropdowns && templateData.dropdowns.has(norm)) {
    return [...templateData.dropdowns.get(norm)].sort((a, b) => a.localeCompare(b));
  }
  return null;
}

let _groupRowsRegistry = []; // index -> rows (Apply button references via data-grp)

// ─── Bulk UPDATE template matching ───
// For every row that will land in TO UPDATE (silent updates + taken-new conflicts),
// look up the corresponding bulk-update-template row by composite key. Attach the
// matched row to each entry (or null) and surface the unmatched count separately.
function computeUpdateMatches() {
  if (!diffResult) return;
  const all = [
    ...(diffResult.silentUpdates || []),
    ...(diffResult.conflicts || [])
  ];
  let matched = 0, unmatched = 0;
  diffResult._unmatchedUpdates = [];
  all.forEach(u => {
    if (!updateTemplateData) { u._updateTplRow = null; return; }
    const dbRow = u.dbValues || [];
    const site = dbCols.site >= 0 ? dbRow[dbCols.site] : '';
    const name = dbCols.name >= 0 ? dbRow[dbCols.name] : '';
    const cv   = dbCols.cropVar >= 0 ? dbRow[dbCols.cropVar] : '';
    const key  = makeUpdateKey(site, name, cv);
    const hit  = key ? updateTemplateData.byKey.get(key) : null;
    if (hit) {
      u._updateTplRow = hit.row;
      matched++;
    } else {
      u._updateTplRow = null;
      unmatched++;
      diffResult._unmatchedUpdates.push({
        site: stripSiteCommas(String(site || '')),
        name: String(name || '').trim(),
        cv:   String(cv || '').trim(),
        key
      });
    }
  });
  diffResult._updateMatchedCount = matched;
  diffResult._updateUnmatchedCount = unmatched;

  // ─── Archive matching (same composite key) ─────────────────────────────
  // We dedupe by dbRowIndex so multi-instance DB rows only match once.
  let aMatched = 0, aUnmatched = 0;
  const seenIdx = new Set();
  diffResult._unmatchedArchives = [];
  (diffResult.removes || []).forEach(rm => {
    if (seenIdx.has(rm.dbRowIndex)) { rm._archiveTplRow = null; return; }
    seenIdx.add(rm.dbRowIndex);
    if (!updateTemplateData) { rm._archiveTplRow = null; return; }
    const dbRow = rm.dbValues || [];
    const site = dbCols.site >= 0 ? dbRow[dbCols.site] : '';
    const name = dbCols.name >= 0 ? dbRow[dbCols.name] : '';
    const cv   = dbCols.cropVar >= 0 ? dbRow[dbCols.cropVar] : '';
    const key  = makeUpdateKey(site, name, cv);
    const hit  = key ? updateTemplateData.byKey.get(key) : null;
    if (hit) {
      rm._archiveTplRow = hit.row;
      aMatched++;
    } else {
      rm._archiveTplRow = null;
      aUnmatched++;
      diffResult._unmatchedArchives.push({
        site: stripSiteCommas(String(site || '')),
        name: String(name || '').trim(),
        cv:   String(cv || '').trim(),
        key
      });
    }
  });
  diffResult._archiveMatchedCount = aMatched;
  diffResult._archiveUnmatchedCount = aUnmatched;
}

function renderUpdateMissing() {
  const sec = $('cmp-section-updt-missing');
  if (!sec) return;
  if (!diffResult) { sec.style.display = 'none'; return; }
  const total = (diffResult.silentUpdates || []).length + (diffResult.conflicts || []).length;
  if (total === 0) { sec.style.display = 'none'; return; }

  if (!updateTemplateData) {
    sec.style.display = '';
    $('cmp-updt-missing-title').textContent =
      'Bulk Update Template Required (' + total + ' update' + (total === 1 ? '' : 's') + ' pending)';
    $('cmp-updt-missing-body').innerHTML =
      '<div class="cmp-help">Upload the latest bulk UPDATE template (PickTrace export) above. ' +
      'TO UPDATE rows will be sourced from that file so all 42 columns of existing data are preserved.</div>';
    return;
  }

  const unmatched = diffResult._unmatchedUpdates || [];
  if (unmatched.length === 0) { sec.style.display = 'none'; return; }

  sec.style.display = '';
  $('cmp-updt-missing-title').textContent =
    'Updates Not Found in Bulk Update Template (' + unmatched.length + ')';
  let html =
    '<div class="cmp-help">These updates could not be matched in the uploaded update template ' +
    'by Site + Name + Crop &amp; Variety. They will be flagged in the export under ' +
    '<b>UNMATCHED UPDATES</b> for manual review. Re-download a fresh template from PickTrace ' +
    'and re-upload to resolve.</div>' +
    '<div class="table-wrap"><table class="data-table"><thead><tr>' +
    '<th>Site</th><th>Name</th><th>Crop &amp; Variety</th><th>Composite Key</th>' +
    '</tr></thead><tbody>';
  unmatched.forEach(u => {
    html += '<tr>' +
      '<td>' + escHtml(u.site) + '</td>' +
      '<td>' + escHtml(u.name) + '</td>' +
      '<td>' + escHtml(u.cv) + '</td>' +
      '<td><code>' + escHtml(u.key) + '</code></td>' +
      '</tr>';
  });
  html += '</tbody></table></div>';
  $('cmp-updt-missing-body').innerHTML = html;
}

function renderArchiveMissing() {
  const sec = $('cmp-section-arch-missing');
  if (!sec) return;
  if (!diffResult) { sec.style.display = 'none'; return; }
  const totalArch = (diffResult.removes || []).length;
  if (totalArch === 0) { sec.style.display = 'none'; return; }

  if (!updateTemplateData) {
    sec.style.display = '';
    $('cmp-arch-missing-title').textContent =
      'Bulk Update Template Required for Archive Flags (' + totalArch + ' archive' + (totalArch === 1 ? '' : 's') + ' pending)';
    $('cmp-arch-missing-body').innerHTML =
      '<div class="cmp-help">Upload the bulk UPDATE template above to enable the ' +
      '<code>DATA ENTRY (UPDATE)</code> sheet, which contains the archive rows pre-flagged ' +
      'with <code>Is Archived = TRUE</code>. If you skip this, that sheet will not be emitted ' +
      'and the export will fall back to the legacy red <code>TO ARCHIVE</code> checklist only.</div>';
    return;
  }

  const unmatched = diffResult._unmatchedArchives || [];
  if (unmatched.length === 0) { sec.style.display = 'none'; return; }

  sec.style.display = '';
  $('cmp-arch-missing-title').textContent =
    'Archives Not Found in Bulk Update Template (' + unmatched.length + ')';
  let html =
    '<div class="cmp-help">These archive rows could not be matched in the uploaded update template ' +
    'by Site + Name + Crop &amp; Variety. They will be exported under <b>UNMATCHED ARCHIVES</b> ' +
    '(red) for manual review. Re-download a fresh template from PickTrace and re-upload to resolve.</div>' +
    '<div class="table-wrap"><table class="data-table"><thead><tr>' +
    '<th>Site</th><th>Name</th><th>Crop &amp; Variety</th><th>Composite Key</th>' +
    '</tr></thead><tbody>';
  unmatched.forEach(u => {
    html += '<tr>' +
      '<td>' + escHtml(u.site) + '</td>' +
      '<td>' + escHtml(u.name) + '</td>' +
      '<td>' + escHtml(u.cv) + '</td>' +
      '<td><code>' + escHtml(u.key) + '</code></td>' +
      '</tr>';
  });
  html += '</tbody></table></div>';
  $('cmp-arch-missing-body').innerHTML = html;
}

function renderRequiredDefaults() {
  const sec = $('cmp-section-defaults');
  const list = detectEmptyRequiredColumns();
  diffResult._emptyRequired = list;
  _groupRowsRegistry = [];
  if (!list.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  const totalCells = list.reduce((s, c) => s + c.totalEmpty, 0);
  $('cmp-defaults-title').textContent = 'Empty Required Columns (' + list.length + ' columns / ' + totalCells + ' cells)';
  const cvMaps = diffResult.cvMaps || { learned: new Map(), perComm: new Map(), perVariety: new Map() };
  let html = '<thead><tr><th>Column</th><th>Group</th><th>Empty cells</th><th>Value (dropdown)</th><th>Manual override</th><th></th></tr></thead><tbody>';
  list.forEach(col => {
    const opts = getDropdownOptions(col.header);
    col.groups.forEach(grp => {
      const grpId = _groupRowsRegistry.length;
      _groupRowsRegistry.push({ colIdx: col.idx, rows: grp.rows });

      // For Crop & Variety groups, compute the best fuzzy match against the
      // currently-loaded template (strict mode) and pre-select it. After a
      // template re-upload, any newly-created C&V values will auto-populate.
      let preSelect = '';
      if (col.isCV && grp.comm && grp.variety) {
        preSelect = lookupCanonicalCropVariety(
          grp.comm, grp.variety, cvMaps.learned, cvMaps.perComm, cvMaps.perVariety, true) || '';
      }

      html += '<tr>' +
        '<td><b>' + escHtml(col.header) + '</b></td>' +
        '<td>' + (col.isCV ? '<code>' + escHtml(grp.label) + '</code>' : escHtml(grp.label)) + '</td>' +
        '<td>' + grp.count + '</td>';
      if (opts && opts.length) {
        let optHtml = '<option value="">-- choose --</option>';
        opts.forEach(o => {
          const sel = (preSelect && o === preSelect) ? ' selected' : '';
          optHtml += '<option value="' + escHtml(o) + '"' + sel + '>' + escHtml(o) + '</option>';
        });
        html += '<td><select class="cmp-default-input input-field" style="min-width:260px;">' + optHtml + '</select></td>';
      } else {
        html += '<td><input type="text" class="cmp-default-input input-field" placeholder="Type a value..." style="width:260px;"></td>';
      }
      // Manual override input — overrides the dropdown selection on Apply.
      html += '<td><input type="text" class="cmp-default-override input-field" placeholder="Manual override (optional)" style="width:240px;"></td>';
      html += '<td><button class="btn btn-primary btn-sm cmp-default-apply" data-grp="' + grpId + '">Apply</button></td>';
      html += '</tr>';
    });
  });
  html += '</tbody>';
  $('cmp-defaults-table').innerHTML = html;
}

function applyValueToGroup(grpId, value) {
  const v = value == null ? '' : String(value).trim();
  if (v === '') return 0;
  const reg = _groupRowsRegistry[grpId];
  if (!reg) return 0;
  let filled = 0;
  reg.rows.forEach(b => {
    const cur = b.row[reg.colIdx];
    if (cur == null || String(cur).trim() === '') {
      b.row[reg.colIdx] = v;
      filled++;
    }
  });
  return filled;
}

function renderSitesToCreate() {
  const sec = $('cmp-section-sites');
  const list = (diffResult.sitesToCreate || []);
  if (!list.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';

  const unarchiveCount = list.filter(s => s.action === 'unarchive').length;
  const createCount    = list.filter(s => s.action === 'create-new').length;
  const unknownCount   = list.filter(s => s.action === 'unknown').length;
  let title = 'Sites to Set Up in PickTrace (' + list.length + ')';
  const subParts = [];
  if (unarchiveCount > 0) subParts.push(unarchiveCount + ' unarchive');
  if (createCount > 0)    subParts.push(createCount + ' create new');
  if (unknownCount > 0)   subParts.push(unknownCount + ' unknown (no Sites file)');
  if (subParts.length) title += ' — ' + subParts.join(', ');
  $('cmp-sites-title').textContent = title;

  const actionLabel = a => {
    if (a === 'unarchive')  return '<span class="cmp-action cmp-action-un">Unarchive</span>';
    if (a === 'create-new') return '<span class="cmp-action cmp-action-new">Create new</span>';
    return '<span class="cmp-action cmp-action-unk">Unknown</span>';
  };

  const exactToggle = $('cmp-sites-exact');
  const useExact = !!(exactToggle && exactToggle.checked);
  let html = '<thead><tr><th>Action</th><th>Site name</th><th>Suggested Alt ID</th><th>New blocks</th><th>Example block names</th></tr></thead><tbody>';
  list.forEach(s => {
    const siteCell = useExact ? (s.rawDisplay || s.longName) : s.longName;
    const examples = useExact ? (s.rawExamples || s.examples) : s.examples;
    html += '<tr>' +
      '<td>' + actionLabel(s.action) + '</td>' +
      '<td><b>' + escHtml(siteCell) + '</b></td>' +
      '<td><code>' + escHtml(s.code || '') + '</code></td>' +
      '<td>' + s.count + '</td>' +
      '<td>' + examples.map(e => escHtml(e)).join('<br>') + '</td>' +
      '</tr>';
  });
  html += '</tbody>';
  $('cmp-sites-table').innerHTML = html;
}

function copySitesList() {
  const list = (diffResult && diffResult.sitesToCreate) || [];
  if (!list.length) return;
  const exact = $('cmp-sites-exact');
  const useExact = !!(exact && exact.checked);
  const actionLbl = a => a === 'unarchive' ? 'Unarchive' : a === 'create-new' ? 'Create new' : 'Unknown';
  const lines = list.map(s =>
    actionLbl(s.action) + '\t' + (useExact ? (s.rawDisplay || s.longName) : s.longName) + '\t' + (s.code || '') + '\t' + s.count + ' new blocks'
  );
  const text = 'Action\tSite name\tAlt ID\tNew blocks\n' + lines.join('\n');
  navigator.clipboard.writeText(text).then(
    () => alert('Copied ' + list.length + ' Sites to clipboard.'),
    () => alert('Copy failed — clipboard access denied.')
  );
}

// Scan rows that will end up in the export for reversed Acreage / Plant Count.
// Plant Count is always > Acreage in real data; if Acreage > Plant Count the values
// are flipped in the source DB. Returns [{ dbRowIndex, acres, plants, name, site }].
function detectReversedAcresInDb() {
  const out = [];
  if (!dbData || !dbCols || dbCols.acreage < 0 || dbCols.plantCount < 0) return out;
  const removeIdxSet = new Set();
  (diffResult.removes || []).forEach(r => removeIdxSet.add(r.dbRowIndex));
  const silentMap = new Map();
  (diffResult.silentUpdates || []).forEach(s => silentMap.set(s.dbRowIndex, s.newValues));
  const conflictMap = new Map();
  (diffResult.conflicts || []).forEach(c => conflictMap.set(c.dbRowIndex, c));

  dbData.rows.forEach((r, i) => {
    if (removeIdxSet.has(i)) return;
    if (isArchived(r)) return;
    const effective = silentMap.has(i) ? silentMap.get(i) : r;
    const acres = parseFloat(String(effective[dbCols.acreage] != null ? effective[dbCols.acreage] : '').replace(/[, ]/g, ''));
    const plants = parseFloat(String(effective[dbCols.plantCount] != null ? effective[dbCols.plantCount] : '').replace(/[, ]/g, ''));
    if (isNaN(acres) || isNaN(plants)) return;
    if (acres > 0 && plants > 0 && acres > plants) {
      out.push({
        dbRowIndex: i,
        acres,
        plants,
        name: r[dbCols.name] || '',
        site: r[dbCols.site] || ''
      });
    }
  });
  return out;
}

// Apply the auto-swap: for every reversed row, swap Acreage <-> Plant Count
// in dbData.rows AND in any silentUpdates / conflicts that target that row.
function applyAcresSwap() {
  const list = (diffResult && diffResult.acresReversed) || [];
  if (!list.length) return;
  const swap = (arr, a, b) => { const t = arr[a]; arr[a] = arr[b]; arr[b] = t; };

  list.forEach(item => {
    const r = dbData.rows[item.dbRowIndex];
    if (r) swap(r, dbCols.acreage, dbCols.plantCount);
  });
  // Mirror the swap in any computed-row buckets that snapshot DB values.
  (diffResult.silentUpdates || []).forEach(s => {
    if (list.some(x => x.dbRowIndex === s.dbRowIndex)) {
      swap(s.dbValues,  dbCols.acreage, dbCols.plantCount);
      // newValues are already from Add file — don't re-swap.
    }
  });
  (diffResult.conflicts || []).forEach(c => {
    if (list.some(x => x.dbRowIndex === c.dbRowIndex)) {
      swap(c.dbValues, dbCols.acreage, dbCols.plantCount);
    }
  });
  // Also swap remove rows that had reversed values, so the TO ARCHIVE sheet shows the corrected form.
  (diffResult.removes || []).forEach(rm => {
    if (list.some(x => x.dbRowIndex === rm.dbRowIndex)) {
      swap(rm.dbValues, dbCols.acreage, dbCols.plantCount);
    }
  });

  diffResult.autoSwappedDbRows = (diffResult.autoSwappedDbRows || 0) + list.length;
  diffResult.acresReversed = []; // cleared
  renderResults();
}

// Compute unique Crop & Variety values from new-adds + silent-updates + conflicts
// that aren't in the bulk template's Crop & Variety dropdown — the user must create
// these in PickTrace before uploading. Mirrors computeSitesToCreate.
function computeCropVarietiesToCreate(newAdds, silentUpdates, conflicts, cvMaps) {
  if (!dbCols || dbCols.cropVar < 0) return [];
  const templateSet = new Set();
  if (templateData && templateData.cropVarieties) {
    templateData.cropVarieties.forEach(s => {
      const u = String(s || '').trim().toUpperCase();
      if (u) templateSet.add(u);
    });
  }
  const maps = cvMaps || { learned: new Map(), perComm: new Map(), perVariety: new Map() };
  const need = new Map(); // valueUpper -> { value, count, comm, variety, examples }

  const consider = (row, fields) => {
    if (!row) return;
    let cv = row[dbCols.cropVar];
    cv = cv != null ? String(cv).trim() : '';
    // Strict-mode lookup left this cell blank → synthesize a best-guess name
    // (non-strict) so the user sees what we'd suggest creating in PickTrace.
    if (!cv && fields && fields.comm && fields.variety) {
      cv = lookupCanonicalCropVariety(fields.comm, fields.variety,
        maps.learned, maps.perComm, maps.perVariety, false);
    }
    if (!cv) return;
    const upper = cv.toUpperCase();
    if (templateSet.has(upper)) return; // already in dropdown — no action needed
    if (!need.has(upper)) {
      need.set(upper, {
        value: cv,
        count: 0,
        comm: (fields && fields.comm) || '',
        variety: (fields && fields.variety) || '',
        examples: []
      });
    }
    const entry = need.get(upper);
    entry.count++;
    const blockName = row[dbCols.name] || (fields && fields.blockCanonical) || '';
    if (entry.examples.length < 3 && blockName) entry.examples.push(blockName);
  };

  (newAdds || []).forEach(a => consider(a.row, a.fields));
  (silentUpdates || []).forEach(s => consider(s.newValues, s.fields));
  (conflicts || []).forEach(c => consider(c.newValues, c.fields));
  return [...need.values()].sort((a, b) => a.value.localeCompare(b.value));
}

// Compute the unique Site names from New Adds that don't exist in the bulk template's
// active Site dropdown. If a Sites Master List is loaded, classify each as either
// 'unarchive' (exists in master but archived in PickTrace) or 'create-new'
// (doesn't exist in PickTrace at all).
function computeSitesToCreate(newAdds, toUnarchive) {
  const buckets = [];
  (newAdds   || []).forEach(a => buckets.push({ entry: a, kind: 'new' }));
  (toUnarchive || []).forEach(a => buckets.push({ entry: a, kind: 'unarchive' }));
  if (!buckets.length) return [];
  const templateSet = new Set();
  if (templateData && templateData.sites) {
    templateData.sites.forEach(s => {
      const ln = siteLongName(s);
      if (ln) templateSet.add(ln.toUpperCase());
    });
  }
  const need = new Map(); // longNameUpper -> { longName, code, count, examples, action, kinds }
  buckets.forEach(({ entry: a, kind }) => {
    const longName = a.fields && a.fields.growerName
      ? a.fields.growerName
      : siteLongName(a.row ? a.row[dbCols.site] : (a.dbValues ? a.dbValues[dbCols.site] : ''));
    if (!longName) return;
    const upper = longName.toUpperCase();
    // The bulk create template's dropdown lists every site PickTrace knows about
    // regardless of archive status, so it's NOT a reliable "site is active" signal.
    // Only treat it as a shortcut when the Sites Master CSV isn't loaded — when
    // it IS loaded, the master is authoritative and we fall through to its check.
    if (!sitesMasterData && templateSet.has(upper)) return;

    // Sites master CSV check — authoritative source for what's in PickTrace.
    // Match strategies (any one is enough):
    //   1. Long-name exact match
    //   2. Embedded grower code match (handles CSV typos like "MADERA... (AUSTMAD")
    //   3. Word-prefix match — handles "DAVID PETERS BYPASS TRUST" vs CSV's "DAVID PETERS"
    //      (requires CSV to have ≥2 words to avoid single-word false positives)
    let action;
    if (sitesMasterData) {
      const codeUpper = (a.fields && a.fields.growerCode) ? a.fields.growerCode.toUpperCase() : '';
      const inMasterByName = sitesMasterData.allLongNames.has(upper);
      const inMasterByCode = codeUpper && sitesMasterData.allCodes && sitesMasterData.allCodes.has(codeUpper);
      const archivedByName = sitesMasterData.archivedLongNames.has(upper);
      const archivedByCode = codeUpper && sitesMasterData.archivedCodes && sitesMasterData.archivedCodes.has(codeUpper);

      let inMasterByPrefix = false, archivedByPrefix = false;
      if (!inMasterByName && !inMasterByCode) {
        const addWords = upper.split(/\s+/);
        sitesMasterData.allLongNames.forEach(masterUpper => {
          if (inMasterByPrefix) return;
          const masterWords = masterUpper.split(/\s+/);
          if (masterWords.length < 2) return;                  // safety guard
          if (addWords.length < masterWords.length) return;
          for (let j = 0; j < masterWords.length; j++) {
            if (addWords[j] !== masterWords[j]) return;
          }
          inMasterByPrefix = true;
          if (sitesMasterData.archivedLongNames.has(masterUpper)) archivedByPrefix = true;
        });
      }

      const inMaster = inMasterByName || inMasterByCode || inMasterByPrefix;
      const fullyArchived = archivedByName || archivedByCode || archivedByPrefix;
      // Active in master (has at least one un-archived location) → user doesn't
      // need to create or unarchive the site itself. New blocks just slot in;
      // archived blocks just need block-level unarchive in PickTrace.
      if (inMaster && !fullyArchived) return;
      if (inMaster && fullyArchived) action = 'unarchive';
      else if (kind === 'unarchive') {
        // Block exists archived in DB but the site isn't in the master — the
        // master CSV is incomplete or out of date. Surface as 'unarchive' so
        // the user knows to verify/unarchive the site in PickTrace.
        action = 'unarchive';
      }
      else action = 'create-new';
    } else {
      action = 'unknown';
    }

    if (!need.has(upper)) {
      // rawDisplay = the verbatim Add-file Grower string (e.g.
      // "MARDA15(15 AC MURCOTTS)") so the panel can show the un-stripped form
      // when the user toggles "exact names". Falls back to the parsed long name.
      const rawDisplay = (a.fields && a.fields.growerRaw) || longName;
      need.set(upper, {
        longName,
        rawDisplay,
        code: (a.fields && a.fields.growerCode) || '',
        count: 0,
        examples: [],
        rawExamples: [],
        action
      });
    }
    const e = need.get(upper);
    e.count++;
    if (e.examples.length < 3) {
      const exampleName = a.row ? a.row[dbCols.name]
        : (a.dbValues ? a.dbValues[dbCols.name] : '');
      e.examples.push(exampleName || (a.fields && a.fields.blockCanonical) || '');
      e.rawExamples.push((a.fields && a.fields.blockRaw) || exampleName || '');
    }
  });
  return [...need.values()].sort((a, b) => {
    // Sort: unarchive first, then create-new, then unknown — alphabetical inside each.
    const order = { 'unarchive': 0, 'create-new': 1, 'unknown': 2 };
    if (order[a.action] !== order[b.action]) return order[a.action] - order[b.action];
    return a.longName.localeCompare(b.longName);
  });
}

// Preview spec mirrors the bulk template — Crop & Variety stays as ONE column.
function previewSpec() {
  const out = [];
  if (dbCols.site >= 0)       out.push({ type: 'col', idx: dbCols.site,       label: dbData.headers[dbCols.site] });
  if (dbCols.name >= 0)       out.push({ type: 'col', idx: dbCols.name,       label: dbData.headers[dbCols.name] });
  if (dbCols.altId >= 0)      out.push({ type: 'col', idx: dbCols.altId,      label: dbData.headers[dbCols.altId] });
  if (dbCols.locType >= 0)    out.push({ type: 'col', idx: dbCols.locType,    label: dbData.headers[dbCols.locType] });
  if (dbCols.cropVar >= 0)    out.push({ type: 'col', idx: dbCols.cropVar,    label: dbData.headers[dbCols.cropVar] });
  if (dbCols.acreage >= 0)    out.push({ type: 'col', idx: dbCols.acreage,    label: dbData.headers[dbCols.acreage] });
  if (dbCols.plantCount >= 0) out.push({ type: 'col', idx: dbCols.plantCount, label: dbData.headers[dbCols.plantCount] });
  return out;
}

function renderCell(row, spec) {
  return row[spec.idx] != null ? String(row[spec.idx]) : '';
}

function renderAdds() {
  const sec = $('cmp-section-adds');
  if (diffResult.newAdds.length === 0) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  $('cmp-adds-title').textContent = 'New Adds (' + diffResult.newAdds.length + ')';

  const specs = previewSpec();
  let html = '<thead><tr><th style="width:32px;"></th>';
  specs.forEach(s => { html += '<th>' + escHtml(s.label) + '</th>'; });
  html += '</tr></thead><tbody>';
  diffResult.newAdds.forEach((a, i) => {
    html += '<tr><td><input type="checkbox" class="cmp-add-check" data-idx="' + i + '" checked></td>';
    specs.forEach(s => { html += '<td>' + escHtml(renderCell(a.row, s)) + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody>';
  $('cmp-adds-table').innerHTML = html;
}

function renderConflicts() {
  const sec = $('cmp-section-conflicts');
  if (diffResult.conflicts.length === 0) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  $('cmp-conflicts-title').textContent = 'Conflicts (' + diffResult.conflicts.length + ')';

  const specs = previewSpec();
  let html = '<thead><tr><th style="width:170px;">Resolution</th><th style="width:60px;">Source</th>';
  specs.forEach(s => { html += '<th>' + escHtml(s.label) + '</th>'; });
  html += '</tr></thead><tbody>';

  diffResult.conflicts.forEach((c, i) => {
    const diffSet = new Set(c.diffFields);
    html += '<tr class="cmp-conf-db" data-conf="' + i + '">' +
      '<td rowspan="2" class="cmp-conf-cell">' +
        '<div class="cmp-radio-group">' +
          '<label><input type="radio" name="cmp-conf-' + i + '" value="db" checked> Keep DB</label>' +
          '<label><input type="radio" name="cmp-conf-' + i + '" value="new"> Take New</label>' +
          '<label><input type="radio" name="cmp-conf-' + i + '" value="skip"> Skip</label>' +
        '</div>' +
      '</td>' +
      '<td class="cmp-source-db">DB</td>';
    specs.forEach(s => {
      const cls = diffSet.has(s.idx) ? ' class="cmp-cell-diff"' : '';
      html += '<td' + cls + '>' + escHtml(renderCell(c.dbValues, s)) + '</td>';
    });
    html += '</tr>';
    html += '<tr class="cmp-conf-new"><td class="cmp-source-new">New</td>';
    specs.forEach(s => {
      const cls = diffSet.has(s.idx) ? ' class="cmp-cell-diff"' : '';
      html += '<td' + cls + '>' + escHtml(renderCell(c.newValues, s)) + '</td>';
    });
    html += '</tr>';
  });
  html += '</tbody>';
  $('cmp-conflicts-table').innerHTML = html;
}

function renderRemoves() {
  const sec = $('cmp-section-removes');
  if (diffResult.removes.length === 0) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  $('cmp-removes-title').textContent = 'Removes (' + diffResult.removes.length + ')';

  const specs = previewSpec();
  let html = '<thead><tr><th style="width:32px;"></th>';
  specs.forEach(s => { html += '<th>' + escHtml(s.label) + '</th>'; });
  html += '</tr></thead><tbody>';
  diffResult.removes.forEach((r, i) => {
    html += '<tr><td><input type="checkbox" class="cmp-rm-check" data-idx="' + i + '" checked></td>';
    specs.forEach(s => { html += '<td>' + escHtml(renderCell(r.dbValues, s)) + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody>';
  $('cmp-removes-table').innerHTML = html;
}

function renderSkipped() {
  const sec = $('cmp-section-skipped');
  if (diffResult.skipped.length === 0) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  $('cmp-skipped-title').textContent = 'Skipped (' + diffResult.skipped.length + ')';
  let html = '<ul class="cmp-skipped-ul">';
  diffResult.skipped.forEach(s => {
    html += '<li><span class="cmp-skipped-loc">' + escHtml(s.site || '—') + ' / ' + escHtml(s.location || '—') + '</span>' +
      ' <span class="cmp-skipped-reason">' + escHtml(s.reason) + '</span></li>';
  });
  html += '</ul>';
  $('cmp-skipped-list').innerHTML = html;
}

// ─── Export ───
function exportXlsx() {
  if (!diffResult || !dbData) { alert('Run a comparison first.'); return; }

  // Collect user choices
  const removeIdxSet = new Set();
  document.querySelectorAll('.cmp-rm-check').forEach(cb => {
    if (cb.checked) {
      const i = +cb.dataset.idx;
      removeIdxSet.add(diffResult.removes[i].dbRowIndex);
    }
  });

  const addIdxSet = new Set();
  document.querySelectorAll('.cmp-add-check').forEach(cb => {
    if (cb.checked) addIdxSet.add(+cb.dataset.idx);
  });

  // Conflict resolutions: dbRowIndex -> 'db' | 'new' | 'skip'
  const conflictRes = {};
  diffResult.conflicts.forEach((c, i) => {
    const r = document.querySelector('input[name="cmp-conf-' + i + '"]:checked');
    conflictRes[c.dbRowIndex] = r ? r.value : 'db';
  });

  // ─── DATA ENTRY = truly-new blocks (Add file entries with codes that don't
  //     exist in the DB at all). Untouched DB rows, silent updates, and
  //     conflicts are NOT in DATA ENTRY — updates go to the separate TO UPDATE sheet.
  //     SPECIAL CASE: if the DB came from an Org Data import and there's no
  //     Add file, treat every imported row as if it were a new add — that's
  //     the "format my org data into the create template" workflow.
  const outDb = [];
  if (dbData && dbData._fromOrgImport && !addData) {
    (dbData.rows || []).forEach(r => outDb.push(r));
  } else {
    diffResult.newAdds.forEach((a, i) => {
      if (addIdxSet.has(i)) outDb.push(a.row);
    });
  }

  // Decide output schema: template if loaded, else DB.
  // The bulk upload does NOT process "Is Archived" — drop it unconditionally.
  const useTemplate = !!(templateData && templateData.dataEntryHeaders.length && templateData.rawBuffer);
  const baseHeaders = (useTemplate ? templateData.dataEntryHeaders : dbData.headers)
    .filter(h => normHeader(h) !== 'is archived' && normHeader(h) !== 'archived');

  // Replace the single "Crop & Variety*" column with two columns: "Crop*" and "Variety*".
  // This is what the user wants in the EXPORT (mirrors the in-app preview split).
  const outHeaders = [];
  // Spec: each entry { src: dbIdx-or-(-1), split: 'crop'|'variety'|null }
  const outSpec = [];

  const dbHeaderMap = new Map();
  dbData.headers.forEach((h, i) => dbHeaderMap.set(normHeader(h), i));

  baseHeaders.forEach(h => {
    const norm = normHeader(h);
    const dbIdx = dbHeaderMap.has(norm) ? dbHeaderMap.get(norm) : -1;
    const required = /\*$/.test(h);
    outHeaders.push(h);
    outSpec.push({ src: dbIdx, dropKey: norm, required });
  });

  // ─── TO UPDATE schema = bulk update template = bulk create template + "Is Archived"
  //     appended at the end. Required (*) columns are identical to create, so the same
  //     empty-required-cell flow validates both. Is Archived flows through from the DB.
  const updateHeaders = outHeaders.slice();
  const updateSpec = outSpec.slice();
  updateHeaders.push('Is Archived');
  updateSpec.push({
    src: (dbCols && dbCols.archived >= 0) ? dbCols.archived : -1,
    dropKey: 'is archived',
    required: false
  });

  // Dropdown matching against the template:
  //   - OPTIONAL columns: non-matching values are blanked.
  //   - REQUIRED columns (* in header — Site*, Name*, Location Type*, Crop & Variety*, Start Date*):
  //     non-matching values are KEPT so the row stays usable; you can fix them in Excel
  //     before uploading. Blanking a required field would break the row anyway.
  const dropdowns = useTemplate ? templateData.dropdowns : null;
  let droppedValues = 0;
  let keptUnmatched = 0;

  const validateCell = (sp, val) => {
    if (!val) return '';
    if (!useTemplate) return val;
    const set = dropdowns && dropdowns.get(sp.dropKey);
    if (!set) return val; // No dropdown for this column — pass through.
    if (set.has(val)) return val;
    if (sp.required) { keptUnmatched++; return val; }
    droppedValues++;
    return '';
  };

  // Helper: emit a numeric value for cells that should be Excel numbers (not text).
  const isNumericCol = sp => sp.dropKey === 'acreage' || sp.dropKey === 'plant count' ||
                              sp.dropKey === 'length' || sp.dropKey === 'stand count' ||
                              sp.dropKey === 'row/bed count' || sp.dropKey === 'post count' ||
                              sp.dropKey === 'percent covered' ||
                              sp.dropKey === 'row spacing, in.' || sp.dropKey === 'plant spacing, in.' ||
                              sp.dropKey === 'post spacing, in.' || sp.dropKey === 'bed width, in.';
  const toNumOrText = (v) => {
    if (v == null || v === '') return '';
    const n = parseFloat(String(v).replace(/[, ]/g, ''));
    return isNaN(n) ? String(v) : n;
  };

  // The "keep exact names" checkbox — when on, DATA ENTRY (CREATE) writes the
  // Site and Name cells using the Add file's verbatim value (e.g.
  //   Site: "BOMAPUB(PUBLIC PROPERTIES, INC.)"
  //   Name: "23300(BLK#1 RUSH TI)"
  // ) instead of the stripped/canonicalized form. Pulled from a.fields per row.
  const earlyOpts = window.cmpExportOpts || {};
  const keepExact = !!earlyOpts.keepExactNames;
  const outRows = outDb.map((row, ri) => {
    // Find the matching newAdds entry for this DATA ENTRY (CREATE) row so we
    // can pull growerRaw/blockRaw when needed. outDb is built from
    // diffResult.newAdds in the same iteration order.
    const a = (diffResult.newAdds || []).filter((_, i) => addIdxSet.has(i))[ri] || null;
    return outSpec.map(sp => {
      if (sp.src < 0) return '';
      let v;
      if (keepExact && a && sp.dropKey === 'site' && a.fields && a.fields.growerRaw) {
        v = a.fields.growerRaw;
      } else if (keepExact && a && sp.dropKey === 'name' && a.fields && a.fields.blockRaw) {
        v = a.fields.blockRaw;
      } else {
        const raw = row[sp.src];
        v = raw != null ? String(raw).trim() : '';
        if (sp.dropKey === 'site') v = stripSiteCommas(v);
        v = validateCell(sp, v);
      }
      if (isNumericCol(sp)) return toNumOrText(v);
      return v;
    });
  });

  // ─── Year-only date expansion: turn bare years / year ranges in the date
  //     columns (e.g. Planted Date "2012", "2010-2011") into full YYYY-MM-DD
  //     using the operator's chosen month/day, keeping each row's year.
  if (cmpDateFill && cmpDateFill.mm && cmpDateFill.dd) {
    const dfIdxs = outHeaders.map((h, i) => CMP_DATE_COLS.has(normHeader(h)) ? i : -1).filter(i => i >= 0);
    outRows.forEach(cells => {
      dfIdxs.forEach(ci => {
        if (isYearOnlyDate(cells[ci])) cells[ci] = expandYearOnlyDate(cells[ci], cmpDateFill.mm, cmpDateFill.dd, cmpDateFill.range);
      });
    });
  }

  // ─── Final safety-net pass: swap Acreage <-> Plant Count on any output row
  //     where Acreage > Plant Count (Plant Count is always greater in real data).
  const acColIdx = outHeaders.findIndex(h => normHeader(h) === 'acreage');
  const pcColIdx = outHeaders.findIndex(h => normHeader(h) === 'plant count');
  let exportSwapped = 0;
  if (acColIdx >= 0 && pcColIdx >= 0) {
    outRows.forEach(cells => {
      const ac = typeof cells[acColIdx] === 'number' ? cells[acColIdx] : parseFloat(String(cells[acColIdx] || '').replace(/[, ]/g, ''));
      const pc = typeof cells[pcColIdx] === 'number' ? cells[pcColIdx] : parseFloat(String(cells[pcColIdx] || '').replace(/[, ]/g, ''));
      if (!isNaN(ac) && !isNaN(pc) && ac > 0 && pc > 0 && ac > pc) {
        const tmp = cells[acColIdx];
        cells[acColIdx] = cells[pcColIdx];
        cells[pcColIdx] = tmp;
        exportSwapped++;
      }
    });
  }
  if (exportSwapped > 0) {
    console.log('[Block Compare] Export-time safety swap on ' + exportSwapped + ' rows.');
  }

  if (droppedValues > 0) {
    console.warn('[Block Compare] ' + droppedValues + ' optional cell value(s) blanked (not in template dropdowns).');
  }
  if (keptUnmatched > 0) {
    console.warn('[Block Compare] ' + keptUnmatched + ' required cell value(s) kept despite not matching the template dropdown — fix manually before upload.');
  }

  const wb = XLSX.utils.book_new();
  // Sheets are stored here as they are built and appended at the end in a fixed
  // user-facing order (DATA ENTRY (CREATE) → DATA ENTRY (UPDATE) → TO UPDATE →
  // TO ARCHIVE → ...).
  const builtSheets = {};
  const dataSheet = XLSX.utils.aoa_to_sheet([outHeaders].concat(outRows));
  builtSheets['DATA ENTRY (CREATE)'] = dataSheet;

  // ─── TO UPDATE sheet — silent updates + (taken-new) conflicts.
  //     Uses the bulk UPDATE template (= create columns + "Is Archived").
  //     Each row is sourced from the matched bulk-update-template row (preserving
  //     all 42 columns of pre-existing PickTrace data). Only the cells where the
  //     Add file's new value differs from the DB are overlaid + highlighted yellow.
  //     Rows whose composite key isn't found in the update template go to a
  //     separate UNMATCHED UPDATES sheet (red) for manual review.
  const updateAcColIdx = updateHeaders.findIndex(h => normHeader(h) === 'acreage');
  const updatePcColIdx = updateHeaders.findIndex(h => normHeader(h) === 'plant count');
  const tplHeaderIdx = updateTemplateData ? updateTemplateData.headerIdx : null;

  // Overlay set — only these columns are allowed to receive Add-file values
  // on top of the bulk update template's existing data. Identity columns
  // (Site / Name / Crop & Variety — already part of the match key) and
  // synthesized columns (Alt ID) are never overlaid: their template value
  // wins so PickTrace's authoritative formatting is preserved.
  const OVERLAY_COLS = new Set(['acreage', 'plant count']);

  const buildUpdateCells = (newRowVals, dbRowVals, tplRow) => {
    const cells = updateSpec.map(sp => {
      const isArchived = sp.dropKey === 'is archived';
      const fromTemplate = !!(tplRow && tplHeaderIdx && tplHeaderIdx.has(sp.dropKey));
      const eligibleOverlay = !isArchived && sp.src >= 0 && OVERLAY_COLS.has(sp.dropKey);

      // Base value: template (matched) or DB (fallback for unmatched / no-template-uploaded).
      let baseVal = '';
      if (fromTemplate) {
        baseVal = tplRow[tplHeaderIdx.get(sp.dropKey)];
      } else if (sp.src >= 0) {
        baseVal = dbRowVals[sp.src];
      }

      // Overlay only the explicit update columns.
      let finalVal = baseVal;
      if (eligibleOverlay) {
        const dbVal = dbRowVals[sp.src];
        const newVal = newRowVals[sp.src];
        if (newVal != null && !valuesEqual(dbVal, newVal)) {
          finalVal = newVal;
        }
      }

      let v = finalVal != null ? String(finalVal).trim() : '';
      // Site comma/period stripping is reserved for non-template-sourced rows.
      // Template-sourced Site values must be passed through verbatim because that
      // string IS PickTrace's identity for the row — any rewrite breaks the upload.
      if (sp.dropKey === 'site' && !fromTemplate) v = stripSiteCommas(v);
      // Skip dropdown validation for template-sourced values — PickTrace is the source of truth.
      if (!isArchived && !fromTemplate) v = validateCell(sp, v);
      if (isNumericCol(sp)) return toNumOrText(v);
      return v;
    });
    if (updateAcColIdx >= 0 && updatePcColIdx >= 0) {
      const ac = parseFloat(String(cells[updateAcColIdx] || '').replace(/[, ]/g, ''));
      const pc = parseFloat(String(cells[updatePcColIdx] || '').replace(/[, ]/g, ''));
      if (!isNaN(ac) && !isNaN(pc) && ac > 0 && pc > 0 && ac > pc) {
        const tmp = cells[updateAcColIdx]; cells[updateAcColIdx] = cells[updatePcColIdx]; cells[updatePcColIdx] = tmp;
      }
    }
    // Yellow highlight on overlay-eligible columns where Add ≠ DB (i.e. real updates).
    const changed = new Set();
    updateSpec.forEach((sp, i) => {
      if (!OVERLAY_COLS.has(sp.dropKey)) return;
      if (sp.src < 0) return;
      const dbRaw = dbRowVals[sp.src];
      const newRaw = newRowVals[sp.src];
      const dbv = dbRaw != null ? String(dbRaw).trim() : '';
      const newv = newRaw != null ? String(newRaw).trim() : '';
      if (newv && !valuesEqual(dbv, newv)) changed.add(i);
    });
    return { cells, changed };
  };

  const matchedUpdates = [];   // [{ cells, changed }]
  const unmatchedUpdates = []; // [{ cells, changed }]
  const pushUpdate = (newRowVals, dbRowVals, tplRow) => {
    const result = buildUpdateCells(newRowVals, dbRowVals, tplRow);
    if (tplRow) matchedUpdates.push(result);
    else if (updateTemplateData) unmatchedUpdates.push(result);
    else matchedUpdates.push(result); // no update template uploaded — fall back to single-sheet behavior
  };

  (diffResult.silentUpdates || []).forEach(s => pushUpdate(s.newValues, s.dbValues, s._updateTplRow || null));
  (diffResult.conflicts || []).forEach(c => {
    const res = conflictRes[c.dbRowIndex];
    if (res === 'new') pushUpdate(c.newValues, c.dbValues, c._updateTplRow || null);
  });

  // Composite-key column indexes within updateHeaders — these are the columns
  // the bulk-update-template lookup keys on. Highlighting them in unmatched
  // rows tells the user exactly which fields failed to match.
  const keyColIdxs = [
    updateHeaders.findIndex(h => normHeader(h) === 'site*'),
    updateHeaders.findIndex(h => normHeader(h) === 'name*'),
    updateHeaders.findIndex(h => normHeader(h) === 'crop & variety*')
  ].filter(i => i >= 0);

  const writeStyledSheet = (sheetName, rowsArr, headerHex) => {
    if (!rowsArr.length) return;
    const isUnmatched = sheetName.indexOf('UNMATCHED') === 0;
    // Unmatched sheets get a "Mismatch Reason" column prepended so the issue is
    // visible at a glance. Composite-key columns are highlighted in a brighter red.
    const REASON = 'No match in bulk update template by Site + Name + Crop & Variety. Verify those three fields.';
    const sheetHeaders = isUnmatched ? ['Mismatch Reason'].concat(updateHeaders) : updateHeaders;
    const sheetRows = rowsArr.map(r => isUnmatched ? [REASON].concat(r.cells) : r.cells);
    const sheet = XLSX.utils.aoa_to_sheet([sheetHeaders].concat(sheetRows));
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headerFill = { patternType: 'solid', fgColor: { rgb: headerHex }, bgColor: { rgb: headerHex } };
    const yellowFill = { patternType: 'solid', fgColor: { rgb: 'FFFFEB9C' }, bgColor: { rgb: 'FFFFEB9C' } };
    const redFill    = { patternType: 'solid', fgColor: { rgb: 'FFFCA5A5' }, bgColor: { rgb: 'FFFCA5A5' } };
    const keyFill    = { patternType: 'solid', fgColor: { rgb: 'FFEF4444' }, bgColor: { rgb: 'FFEF4444' } };
    const reasonFill = { patternType: 'solid', fgColor: { rgb: 'FFB91C1C' }, bgColor: { rgb: 'FFB91C1C' } };
    // After the Reason column shift, key columns move right by 1.
    const highlightKeyCols = isUnmatched ? new Set(keyColIdxs.map(i => i + 1)) : new Set(keyColIdxs);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const ref = XLSX.utils.encode_cell({ r: R, c: C });
        if (!sheet[ref]) sheet[ref] = { v: '', t: 's' };
        if (R === 0) {
          sheet[ref].s = { fill: headerFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
        } else if (isUnmatched) {
          if (C === 0) {
            sheet[ref].s = { fill: reasonFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
          } else if (highlightKeyCols.has(C)) {
            sheet[ref].s = { fill: keyFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
          } else {
            sheet[ref].s = { fill: redFill, font: { bold: true, color: { rgb: 'FF7F1D1D' } } };
          }
        } else {
          const changed = rowsArr[R - 1].changed;
          if (changed && changed.has(C)) {
            sheet[ref].s = { fill: yellowFill, font: { bold: true, color: { rgb: 'FF7C2D12' } } };
          }
        }
      }
    }
    builtSheets[sheetName] = sheet;
  };

  // Gate: when "Include updates" is on, the user can opt to drop the redundant
  // separate sheets (default: dropped). Unmatched sheets ALWAYS emit.
  const exportOpts = window.cmpExportOpts || {};
  if (exportOpts.keepSeparateUpdateSheet !== false) {
    writeStyledSheet('TO UPDATE', matchedUpdates, 'FFB45309'); // amber/brown header
  }
  writeStyledSheet('UNMATCHED UPDATES', unmatchedUpdates, 'FFB91C1C'); // red header — always emit

  // ─── DATA ENTRY (UPDATE) sheet — bulk-update-uploadable archive flags + (optional) update rows.
  //     Built only when the bulk update template is uploaded — needs it as the row base.
  //     Default: archive rows only (Is Archived=TRUE, red row). Optional: also include update rows
  //     (Is Archived passed through from template, yellow on Acreage/Plant Count overlay cells).
  if (updateTemplateData) {
    const isArchivedColIdx = updateHeaders.findIndex(h => normHeader(h) === 'is archived');
    const archCheckedIdx = removeIdxSet; // user-checked archives
    const matchedArchiveRows = [];
    const unmatchedArchiveRows = [];
    const seenArchKey = new Set();

    (diffResult.removes || []).forEach(rm => {
      if (!archCheckedIdx.has(rm.dbRowIndex)) return;
      if (seenArchKey.has(rm.dbRowIndex)) return;
      seenArchKey.add(rm.dbRowIndex);

      const tplRow = rm._archiveTplRow || null;
      const dbRow = rm.dbValues || [];

      const cells = updateSpec.map(sp => {
        const isArch = sp.dropKey === 'is archived';
        const fromTpl = !!(tplRow && tplHeaderIdx && tplHeaderIdx.has(sp.dropKey));
        let val = '';
        if (fromTpl) val = tplRow[tplHeaderIdx.get(sp.dropKey)];
        else if (sp.src >= 0) val = dbRow[sp.src];
        let v = val != null ? String(val).trim() : '';
        if (sp.dropKey === 'site' && !fromTpl) v = stripSiteCommas(v);
        if (isArch) v = 'TRUE'; // FORCE archive flag
        if (!isArch && !fromTpl) v = validateCell(sp, v);
        if (isNumericCol(sp)) return toNumOrText(v);
        return v;
      });
      // Acreage <-> Plant Count safety swap on archive rows too.
      if (updateAcColIdx >= 0 && updatePcColIdx >= 0) {
        const ac = parseFloat(String(cells[updateAcColIdx] || '').replace(/[, ]/g, ''));
        const pc = parseFloat(String(cells[updatePcColIdx] || '').replace(/[, ]/g, ''));
        if (!isNaN(ac) && !isNaN(pc) && ac > 0 && pc > 0 && ac > pc) {
          const tmp = cells[updateAcColIdx]; cells[updateAcColIdx] = cells[updatePcColIdx]; cells[updatePcColIdx] = tmp;
        }
      }
      const row = { cells, changed: new Set(), isArchive: true };
      if (tplRow) matchedArchiveRows.push(row);
      else unmatchedArchiveRows.push(row);
    });

    // Also include update rows if the user opted in via the modal checkbox.
    const includeUpdates = !!(window.cmpExportOpts && window.cmpExportOpts.includeUpdatesInBulkSheet);
    const combinedRows = matchedArchiveRows.slice();
    if (includeUpdates) {
      matchedUpdates.forEach(u => combinedRows.push({ cells: u.cells, changed: u.changed, isArchive: false }));
    }

    // Write DATA ENTRY (UPDATE) — red row tint on archive rows + Is Archived=TRUE cell highlighted;
    // yellow only on overlay-eligible cells of update rows.
    if (combinedRows.length > 0) {
      const sheet = XLSX.utils.aoa_to_sheet([updateHeaders].concat(combinedRows.map(r => r.cells)));
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const headerFill = { patternType: 'solid', fgColor: { rgb: 'FFB45309' }, bgColor: { rgb: 'FFB45309' } };
      const yellowFill = { patternType: 'solid', fgColor: { rgb: 'FFFFEB9C' }, bgColor: { rgb: 'FFFFEB9C' } };
      const redFill    = { patternType: 'solid', fgColor: { rgb: 'FFFCA5A5' }, bgColor: { rgb: 'FFFCA5A5' } };
      const trueFill   = { patternType: 'solid', fgColor: { rgb: 'FFB91C1C' }, bgColor: { rgb: 'FFB91C1C' } };
      for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
          const ref = XLSX.utils.encode_cell({ r: R, c: C });
          if (!sheet[ref]) sheet[ref] = { v: '', t: 's' };
          if (R === 0) {
            sheet[ref].s = { fill: headerFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
          } else {
            const r = combinedRows[R - 1];
            if (r.isArchive) {
              if (C === isArchivedColIdx) {
                sheet[ref].s = { fill: trueFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
              } else {
                sheet[ref].s = { fill: redFill, font: { color: { rgb: 'FF7F1D1D' } } };
              }
            } else if (r.changed && r.changed.has(C)) {
              sheet[ref].s = { fill: yellowFill, font: { bold: true, color: { rgb: 'FF7C2D12' } } };
            }
          }
        }
      }
      builtSheets['DATA ENTRY (UPDATE)'] = sheet;
    }

    // UNMATCHED ARCHIVES sheet (red) for archive rows that didn't match the bulk update template.
    writeStyledSheet('UNMATCHED ARCHIVES', unmatchedArchiveRows, 'FFB91C1C');
  }

  // ─── "TO ARCHIVE" sheet — every removed row, in red, so the user can
  //     manually archive each one in PickTrace. NOT for re-upload — checklist only.
  const archiveRows = [];
  const seenArchiveIdx = new Set();
  (diffResult.removes || []).forEach(rm => {
    if (!removeIdxSet.has(rm.dbRowIndex)) return;     // user unchecked it
    if (seenArchiveIdx.has(rm.dbRowIndex)) return;    // dedup multi-instance
    seenArchiveIdx.add(rm.dbRowIndex);
    const cells = outSpec.map(sp => {
      if (sp.src < 0) return '';
      const v = rm.dbValues[sp.src];
      let s = v != null ? String(v).trim() : '';
      if (sp.dropKey === 'site') s = stripSiteCommas(s);
      if (isNumericCol(sp)) return toNumOrText(s);
      return s;
    });
    // Apply the same Acreage <-> Plant Count safety swap to archive rows.
    if (acColIdx >= 0 && pcColIdx >= 0) {
      const ac = parseFloat(String(cells[acColIdx] || '').replace(/[, ]/g, ''));
      const pc = parseFloat(String(cells[pcColIdx] || '').replace(/[, ]/g, ''));
      if (!isNaN(ac) && !isNaN(pc) && ac > 0 && pc > 0 && ac > pc) {
        const tmp = cells[acColIdx];
        cells[acColIdx] = cells[pcColIdx];
        cells[pcColIdx] = tmp;
      }
    }
    archiveRows.push(cells);
  });

  // ─── "TO UNARCHIVE" sheet — Add rows that match an archived DB row.
  //     These should NOT be uploaded as new (would create duplicates in PickTrace);
  //     the user manually unarchives them in PickTrace. Tinted amber to distinguish.
  const unarchiveList = (diffResult.toUnarchive || []);
  if (unarchiveList.length > 0) {
    const unarchiveRows = unarchiveList.map(u => outSpec.map(sp => {
      if (sp.src < 0) return '';
      const v = u.newRow[sp.src];
      let s = v != null ? String(v).trim() : '';
      if (sp.dropKey === 'site') s = stripSiteCommas(s);
      if (isNumericCol(sp)) return toNumOrText(s);
      return s;
    }));
    const unarchiveSheet = XLSX.utils.aoa_to_sheet([outHeaders].concat(unarchiveRows));
    const range = XLSX.utils.decode_range(unarchiveSheet['!ref']);
    const amberFill   = { patternType: 'solid', fgColor: { rgb: 'FFFFEB9C' }, bgColor: { rgb: 'FFFFEB9C' } };
    const headerFill2 = { patternType: 'solid', fgColor: { rgb: 'FF9C5700' }, bgColor: { rgb: 'FF9C5700' } };
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const ref = XLSX.utils.encode_cell({ r: R, c: C });
        if (!unarchiveSheet[ref]) unarchiveSheet[ref] = { v: '', t: 's' };
        if (R === 0) {
          unarchiveSheet[ref].s = { fill: headerFill2, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
        } else {
          unarchiveSheet[ref].s = { fill: amberFill, font: { color: { rgb: 'FF9C5700' } } };
        }
      }
    }
    builtSheets['TO UNARCHIVE'] = unarchiveSheet;
  }

  if (archiveRows.length > 0 && exportOpts.keepSeparateArchiveSheet !== false) {
    const archiveSheet = XLSX.utils.aoa_to_sheet([outHeaders].concat(archiveRows));
    const range = XLSX.utils.decode_range(archiveSheet['!ref']);
    const redFill    = { patternType: 'solid', fgColor: { rgb: 'FFFFC7CE' }, bgColor: { rgb: 'FFFFC7CE' } };
    const headerFill = { patternType: 'solid', fgColor: { rgb: 'FFC00000' }, bgColor: { rgb: 'FFC00000' } };
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const ref = XLSX.utils.encode_cell({ r: R, c: C });
        if (!archiveSheet[ref]) archiveSheet[ref] = { v: '', t: 's' };
        if (R === 0) {
          archiveSheet[ref].s = { fill: headerFill, font: { bold: true, color: { rgb: 'FFFFFFFF' } } };
        } else {
          archiveSheet[ref].s = { fill: redFill, font: { color: { rgb: 'FF9C0006' } } };
        }
      }
    }
    builtSheets['TO ARCHIVE'] = archiveSheet;
  }

  // ─── Append sheets in the user's preferred order ───
  // DATA ENTRY (CREATE) and DATA ENTRY (UPDATE) sit next to each other, then
  // TO UPDATE and TO ARCHIVE, then the supporting unmatched/unarchive sheets.
  const SHEET_ORDER = [
    'DATA ENTRY (CREATE)',
    'DATA ENTRY (UPDATE)',
    'TO UPDATE',
    'TO ARCHIVE',
    'TO UNARCHIVE',
    'UNMATCHED UPDATES',
    'UNMATCHED ARCHIVES'
  ];
  SHEET_ORDER.forEach(name => {
    if (builtSheets[name]) XLSX.utils.book_append_sheet(wb, builtSheets[name], name);
  });

  // Drop-down inputs sheet (if a template is loaded) — copied verbatim so the
  // resulting file is still self-contained for PickTrace validation.
  if (useTemplate) {
    try {
      const tplWb = XLSX.read(new Uint8Array(templateData.rawBuffer), { type: 'array', cellStyles: true });
      const dropName = tplWb.SheetNames.find(n => /drop.?down/i.test(n));
      if (dropName) {
        XLSX.utils.book_append_sheet(wb, tplWb.Sheets[dropName], dropName);
      }
    } catch (e) { /* non-fatal: just ship without the dropdown sheet */ }
  }

  const fname = 'locations-updated-' + todayStr() + '.xlsx';
  XLSX.writeFile(wb, fname, { cellStyles: true });
}

// ─── Debug dump ───
// Serializes the full diff state plus all input metadata into a JSON file
// the user can download and inspect.
function debugDump() {
  if (!diffResult) { alert('Run a comparison first.'); return; }

  const summarizeArchived = () => {
    if (!dbCols || dbCols.archived < 0) return null;
    let archived = 0, active = 0;
    dbData.rows.forEach(r => { if (isArchived(r)) archived++; else active++; });
    return { archived, active };
  };

  const expandRow = (r) => {
    if (!r) return null;
    const out = {};
    dbData.headers.forEach((h, i) => { out[h] = r[i] != null ? r[i] : ''; });
    return out;
  };

  const dump = {
    meta: {
      generatedAt: new Date().toISOString(),
      dbFileName: dbData ? dbData.fileName : null,
      dbRowCount: dbData ? dbData.rows.length : 0,
      dbHeaders: dbData ? dbData.headers : null,
      dbColMap: dbCols,
      dbArchivedSummary: summarizeArchived(),
      addFileName: addData ? addData.fileName : null,
      addRowCount: addData ? addData.rows.length : 0,
      addHeaders: addData ? addData.headers : null,
      removeFileName: removeData ? removeData.fileName : null,
      removeRowCount: removeData ? removeData.rows.length : 0,
      removeHeaders: removeData ? removeData.headers : null,
      templateFileName: templateData ? templateData.fileName : null,
      templateLoaded: !!templateData,
      templateColumnCount: templateData ? templateData.dataEntryHeaders.length : 0,
      templateSiteCount: templateData ? templateData.sites.length : 0,
      templateCropVarietyCount: templateData ? templateData.cropVarieties.length : 0
    },
    summary: {
      newAdds: diffResult.newAdds.length,
      conflicts: diffResult.conflicts.length,
      silentUpdates: (diffResult.silentUpdates || []).length,
      toUnarchive: (diffResult.toUnarchive || []).length,
      sitesToCreate: (diffResult.sitesToCreate || []).length,
      unchanged: diffResult.unchanged,
      removes: diffResult.removes.length,
      skipped: diffResult.skipped.length,
      dupAddRows: diffResult.dupAddRows || 0,
      acresTreesSwapped: diffResult.acresTreesSwapped || 0
    },
    toUnarchive: (diffResult.toUnarchive || []).map(u => ({
      blockCode: u.blockCode,
      dbRowIndex: u.dbRowIndex,
      addRowIndex: u.addRowIndex,
      fields: u.fields,
      dbRow: expandRow(u.dbValues),
      newRow: expandRow(u.newRow)
    })),
    sitesToCreate: diffResult.sitesToCreate || [],
    newAdds: diffResult.newAdds.map(a => ({
      blockCode: a.blockCode,
      addRowIndex: a.addRowIndex,
      previouslyArchived: !!a.previouslyArchived,
      fields: a.fields,
      row: expandRow(a.row)
    })),
    conflicts: diffResult.conflicts.map(c => ({
      blockCode: c.blockCode,
      dbRowIndex: c.dbRowIndex,
      addRowIndex: c.addRowIndex,
      diffFieldNames: (c.diffFields || []).map(i => dbData.headers[i]),
      fields: c.fields,
      dbRow: expandRow(c.dbValues),
      newRow: expandRow(c.newValues)
    })),
    silentUpdates: (diffResult.silentUpdates || []).map(s => {
      // Per-row diff: which columns changed?
      const changes = [];
      dbData.headers.forEach((h, i) => {
        const a = String(s.dbValues[i] == null ? '' : s.dbValues[i]).trim();
        const b = String(s.newValues[i] == null ? '' : s.newValues[i]).trim();
        if (!valuesEqual(a, b)) changes.push({ column: h, before: a, after: b });
      });
      return {
        blockCode: s.blockCode,
        dbRowIndex: s.dbRowIndex,
        addRowIndex: s.addRowIndex,
        fields: s.fields,
        changes,
        dbRow: expandRow(s.dbValues),
        newRow: expandRow(s.newValues)
      };
    }),
    removes: diffResult.removes.map(r => ({
      blockCode: r.blockCode,
      dbRowIndex: r.dbRowIndex,
      site: r.site,
      location: r.location,
      matchKind: r.matchKind,
      dbRow: expandRow(r.dbValues)
    })),
    skipped: diffResult.skipped.slice()
  };

  // Expose on window for live console inspection.
  window.cmpDebug = dump;
  console.log('[Block Compare] Debug dump on window.cmpDebug', dump);

  const blob = new Blob([JSON.stringify(dump, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'block-compare-debug-' + todayStr() + '.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 100);
}

// ─── Export modal ───
// Always shows on Export click. Builds a dynamic message based on which gates
// (if any) fired, and exposes the "include updates in DATA ENTRY (UPDATE)" toggle.
function requestExport() {
  const updateCount = ((diffResult && diffResult.silentUpdates) || []).length +
                      ((diffResult && diffResult.conflicts) || []).length;
  const archiveCount = ((diffResult && diffResult.removes) || []).length;
  const unmatchedUpd = (diffResult && diffResult._updateUnmatchedCount) || 0;
  const unmatchedArc = (diffResult && diffResult._archiveUnmatchedCount) || 0;
  const cropVars = (diffResult && diffResult.cropVarsToCreate) || [];

  const m = $('cmp-export-modal');
  const msgs = [];
  let gate = 'none';        // 'updt' | 'tpl' | 'none' — which file the re-upload button should route to
  let needsAcknowledge = false;

  // Bulk update template missing while updates or archives exist.
  if (!updateTemplateData && (updateCount > 0 || archiveCount > 0)) {
    msgs.push(
      'Bulk update template not uploaded. ' + updateCount + ' update' + (updateCount === 1 ? '' : 's') + ' and ' +
      archiveCount + ' archive' + (archiveCount === 1 ? '' : 's') + ' are pending. Without the template the ' +
      'DATA ENTRY (UPDATE) sheet (with archive flags) will NOT be emitted, and TO UPDATE rows will fall back ' +
      'to DB-row data. Upload the bulk update template to enable the full bulk-update sheet.'
    );
    gate = 'updt';
    needsAcknowledge = true;
  }

  // Updates didn't match the uploaded template.
  if (updateTemplateData && unmatchedUpd > 0) {
    msgs.push(
      unmatchedUpd + ' update' + (unmatchedUpd === 1 ? '' : 's') + ' could not be matched in the bulk update ' +
      'template by Site + Name + Crop & Variety. They will be exported under UNMATCHED UPDATES (red).'
    );
    gate = 'updt';
    needsAcknowledge = true;
  }

  // Archives didn't match the uploaded template.
  if (updateTemplateData && unmatchedArc > 0) {
    msgs.push(
      unmatchedArc + ' archive' + (unmatchedArc === 1 ? '' : 's') + ' could not be matched in the bulk update ' +
      'template. They will be exported under UNMATCHED ARCHIVES (red).'
    );
    gate = gate === 'none' ? 'updt' : gate;
    needsAcknowledge = true;
  }

  // Crop & Variety values still missing in the create template.
  if (cropVars.length) {
    msgs.push(
      cropVars.length + ' Crop & Variety value' + (cropVars.length === 1 ? '' : 's') +
      ' still need to be created in PickTrace before this export will upload cleanly. ' +
      'Re-upload the create template after creating those values, or acknowledge and export anyway with empty cells.'
    );
    if (gate === 'none') gate = 'tpl';
    needsAcknowledge = true;
  }

  // Build modal content.
  $('cmp-export-modal-title').textContent = needsAcknowledge ? 'Confirm Export — Issues Detected' : 'Confirm Export';
  $('cmp-export-modal-msg').innerHTML =
    (msgs.length ? '<ul style="margin:8px 0 0 16px;padding:0;"><li>' + msgs.join('</li><li>') + '</li></ul>'
                 : '<span class="text-muted">All checks passed. Pick options below and export.</span>');

  // Show the bulk update template availability summary near the checkbox.
  const cb = $('cmp-modal-include-updates');
  if (cb) {
    cb.disabled = !updateTemplateData;
    cb.parentElement.style.opacity = updateTemplateData ? '1' : '0.5';
    cb.parentElement.title = updateTemplateData ? '' : 'Upload the bulk update template to enable this option.';
  }

  // Re-upload button: only show if we actually need a file from the user.
  //   - Bulk update template missing → show "Upload bulk update template"
  //   - Otherwise, only show for the create-template C&V gate ("Re-upload bulk template")
  //   - Hide entirely once the bulk update template is imported and no other upload is needed.
  const btn = $('cmp-modal-tpl-btn');
  const lbl = $('cmp-modal-tpl-btn-label');
  let showBtn = false;
  if (!updateTemplateData) {
    if (lbl) lbl.innerHTML = '&#8679; Upload bulk update template';
    showBtn = true;
  } else if (cropVars.length) {
    if (lbl) lbl.innerHTML = '&#8679; Re-upload bulk template';
    showBtn = true;
  }
  if (btn) btn.style.display = showBtn ? '' : 'none';
  m.dataset.gate = gate;

  // Button states: when there are gate issues, expose the "Acknowledge & Export anyway" path;
  // otherwise show the plain Export button.
  $('cmp-modal-export').style.display       = needsAcknowledge ? 'none' : '';
  $('cmp-modal-export-anyway').style.display = needsAcknowledge ? '' : 'none';

  m.style.display = 'flex';
}
function closeExportModal() {
  $('cmp-export-modal').style.display = 'none';
  const f = $('cmp-modal-tpl-file');
  if (f) f.value = '';
}

// ─── Reset ───
function resetAll() {
  dbData = null; addData = null; removeData = null; diffResult = null;
  dbCols = null; templateData = null; updateTemplateData = null;
  $('cmp-db-name').textContent = 'No file selected';
  $('cmp-add-name').textContent = 'No file selected';
  $('cmp-rm-name').textContent = 'No file selected';
  $('cmp-tpl-name').textContent = 'No file selected';
  $('cmp-updt-name').textContent = 'No file selected';
  $('cmp-db-meta').innerHTML = '';
  $('cmp-add-meta').innerHTML = '';
  $('cmp-rm-meta').innerHTML = '';
  $('cmp-tpl-meta').innerHTML = '';
  $('cmp-updt-meta').innerHTML = '';
  $('cmp-sites-mfile-name').textContent = 'No file selected';
  $('cmp-sites-mfile-meta').innerHTML = '';
  sitesMasterData = null;
  // Slot 6 is conditionally revealed when the DB is incomplete — hide on reset.
  const _slot6 = $('cmp-slot-updt');
  if (_slot6) _slot6.style.display = 'none';
  cmpDateFill = { mm: '', dd: '', range: 'first' };
  $('cmp-summary').style.display = 'none';
  ['cmp-section-adds','cmp-section-conflicts','cmp-section-removes','cmp-section-defaults','cmp-section-datefill','cmp-section-sites','cmp-section-cropvars','cmp-section-updt-missing','cmp-section-arch-missing','cmp-section-skipped'].forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
  $('cmp-empty').style.display = '';
  $('cmp-export').disabled = true;
  $('cmp-debug').disabled = true;
  $('cmp-run').disabled = true;
  ['cmp-db-file','cmp-add-file','cmp-rm-file','cmp-tpl-file','cmp-updt-file','cmp-sites-file'].forEach(id => { const el = $(id); if (el) el.value = ''; });
}

// ─── Init / wiring ───
function cmpInit() {
  if (initialized) return;
  initialized = true;

  // All file inputs clear their value after selection so the user can
  // re-select the same file (after a reset, or to retry a load). Without this,
  // the browser suppresses the change event when the chosen file is unchanged.
  $('cmp-db-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'db');
    e.target.value = '';
  });
  // Org-import button + modal wiring
  const dbFromOrgBtn = $('cmp-db-from-org');
  if (dbFromOrgBtn) dbFromOrgBtn.addEventListener('click', openOrgImportPicker);
  const orgImportCancel = $('cmp-org-import-cancel');
  if (orgImportCancel) orgImportCancel.addEventListener('click', closeOrgImportPicker);
  const orgImportList = $('cmp-org-import-list');
  if (orgImportList) {
    orgImportList.addEventListener('click', e => {
      const item = e.target.closest('.cmp-org-list-item');
      if (!item) return;
      importOrgIntoDb(item.dataset.org);
    });
  }
  // Click-outside-to-close on the org import modal
  const orgImportModal = $('cmp-org-import-modal');
  if (orgImportModal) {
    orgImportModal.addEventListener('click', e => {
      if (e.target === orgImportModal) closeOrgImportPicker();
    });
  }
  $('cmp-add-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'add');
    e.target.value = '';
  });
  $('cmp-rm-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'rm');
    e.target.value = '';
  });
  $('cmp-tpl-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'tpl');
    e.target.value = '';
  });
  $('cmp-updt-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'updt');
    e.target.value = '';
  });
  // Re-upload via the Updates Not Found section
  const updtMissingFile = $('cmp-updt-missing-file');
  if (updtMissingFile) {
    updtMissingFile.addEventListener('change', e => {
      if (!e.target.files[0]) return;
      loadFileInto(e.target.files[0], 'updt');
      e.target.value = '';
    });
  }
  // Re-upload via the Archives Not Found section
  const archMissingFile = $('cmp-arch-missing-file');
  if (archMissingFile) {
    archMissingFile.addEventListener('change', e => {
      if (!e.target.files[0]) return;
      loadFileInto(e.target.files[0], 'updt');
      e.target.value = '';
    });
  }
  $('cmp-sites-copy').addEventListener('click', copySitesList);
  // Toggle "exact names" — re-render the Sites to Set Up table.
  const sitesExact = $('cmp-sites-exact');
  if (sitesExact) sitesExact.addEventListener('change', () => {
    if (diffResult) renderSitesToCreate();
  });
  $('cmp-cropvars-copy').addEventListener('click', copyCropVarsList);
  $('cmp-cropvars-tpl-file').addEventListener('change', e => {
    if (!e.target.files[0]) return;
    loadFileInto(e.target.files[0], 'tpl');
    setTimeout(() => { if (templateData) runCompare(); }, 300);
    e.target.value = '';
  });
  $('cmp-sites-file').addEventListener('change', e => {
    if (e.target.files[0]) loadFileInto(e.target.files[0], 'sites-master');
    e.target.value = '';
  });
  // Apply-value buttons in the "Empty Required Columns" panel — manual override
  // takes precedence over the dropdown selection.
  document.getElementById('cmp-section-defaults').addEventListener('click', e => {
    const btn = e.target.closest('.cmp-default-apply');
    if (!btn) return;
    const grpId = +btn.dataset.grp;
    const tr = btn.closest('tr');
    const dropdown = tr.querySelector('.cmp-default-input');
    const override = tr.querySelector('.cmp-default-override');
    const overrideVal = (override && override.value || '').trim();
    const dropdownVal = (dropdown && dropdown.value || '').trim();
    const v = overrideVal || dropdownVal;
    if (!v) { alert('Pick a value or type a manual override first.'); return; }
    const filled = applyValueToGroup(grpId, v);
    btn.textContent = '✓ Filled ' + filled;
    btn.disabled = true;
    btn.classList.remove('btn-primary');
    btn.classList.add('btn-ghost');
    setTimeout(() => renderResults(), 200);
  });

  // Click-to-copy any cell in any preview table — once clicked, the cell stays
  // green as a "done" marker. Clicking again re-copies and toggles green off.
  // The Empty Required Columns section is form-only, so it's excluded.
  document.getElementById('cmp-results').addEventListener('click', e => {
    const td = e.target.closest('.cmp-section .data-table td, #cmp-sites-table td');
    if (!td) return;
    if (td.closest('#cmp-section-defaults')) return; // form section — no copy/green
    if (e.target.closest('input, button, label, select, textarea, a')) return;
    const text = (td.textContent || '').trim();
    if (!text) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        if (td.classList.contains('cmp-copied')) {
          td.classList.remove('cmp-copied');
        } else {
          td.classList.add('cmp-copied');
        }
      }, () => {});
    }
  });

  $('cmp-run').addEventListener('click', runCompare);
  $('cmp-export').addEventListener('click', requestExport);
  { const a = $('cmp-datefill-apply'); if (a) a.addEventListener('click', applyCmpDateFill); }
  { const c = $('cmp-datefill-clear'); if (c) c.addEventListener('click', clearCmpDateFill); }

  // Export gate modal handlers
  const captureExportOpts = () => {
    const includeUpdates = !!$('cmp-modal-include-updates').checked;
    window.cmpExportOpts = {
      includeUpdatesInBulkSheet: includeUpdates,
      // The "keep separate" toggles only apply when updates are bundled into
      // DATA ENTRY (UPDATE). Otherwise the separate sheets always emit (current behavior).
      keepSeparateUpdateSheet:  includeUpdates ? !!$('cmp-modal-keep-update').checked  : true,
      keepSeparateArchiveSheet: includeUpdates ? !!$('cmp-modal-keep-archive').checked : true,
      // When on, DATA ENTRY (CREATE) writes Site and Name (Block) cells using
      // the Add file's verbatim value — no comma/period stripping, no template
      // canonicalization. Other transforms (numeric conversion, swap, validation)
      // are unaffected.
      keepExactNames: !!$('cmp-modal-keep-exact-names').checked
    };
  };
  // Reveal sub-options only when the master "include updates" toggle is on.
  $('cmp-modal-include-updates').addEventListener('change', e => {
    const subs = $('cmp-modal-suboptions');
    if (subs) subs.style.display = e.target.checked ? '' : 'none';
    // Reset sub-checkboxes when the master is turned off.
    if (!e.target.checked) {
      $('cmp-modal-keep-update').checked = false;
      $('cmp-modal-keep-archive').checked = false;
    }
  });
  $('cmp-modal-cancel').addEventListener('click', closeExportModal);
  $('cmp-modal-export-anyway').addEventListener('click', () => {
    captureExportOpts();
    closeExportModal();
    exportXlsx();
  });
  $('cmp-modal-export').addEventListener('click', () => {
    captureExportOpts();
    closeExportModal();
    exportXlsx();
  });
  $('cmp-modal-tpl-file').addEventListener('change', e => {
    if (!e.target.files[0]) return;
    const gate = $('cmp-export-modal').dataset.gate || 'tpl';
    closeExportModal();
    if (gate === 'updt') {
      // Modal asked for the bulk UPDATE template — route it correctly.
      loadFileInto(e.target.files[0], 'updt');
      // After it loads, prompt the user to click Export again — matches need
      // to recompute against the new template before TO UPDATE rows are emitted.
    } else {
      loadFileInto(e.target.files[0], 'tpl');
      // Auto-rerun comparison once the new (create) template loads, so missing C&Vs resolve.
      setTimeout(() => {
        if (templateData) runCompare();
      }, 300);
    }
  });
  $('cmp-debug').addEventListener('click', debugDump);
  $('cmp-reset').addEventListener('click', resetAll);

  // Master select-all checkboxes
  $('cmp-adds-master').addEventListener('change', e => {
    document.querySelectorAll('.cmp-add-check').forEach(cb => { cb.checked = e.target.checked; });
  });
  $('cmp-rm-master').addEventListener('change', e => {
    document.querySelectorAll('.cmp-rm-check').forEach(cb => { cb.checked = e.target.checked; });
  });

  // Bulk conflict actions
  $('cmp-conf-all-db').addEventListener('click', () => {
    document.querySelectorAll('input[type=radio][name^="cmp-conf-"][value="db"]').forEach(r => { r.checked = true; });
  });
  $('cmp-conf-all-new').addEventListener('click', () => {
    document.querySelectorAll('input[type=radio][name^="cmp-conf-"][value="new"]').forEach(r => { r.checked = true; });
  });
}

// Expose init for switchPage()
window.cmpInit = cmpInit;

})();
