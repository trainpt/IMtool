// ═══════════════════════════════════════════════════════════════════════
// Bulk Template Standardize — takes a set of PickTrace account exports and
// fills every personalized bulk-create template in one pass, instead of
// running Template Standardize once per entity.
//
// Each domain is independent: one export + its own template → one filled
// download. Templates keep the house shape those tabs already assume —
// a DATA ENTRY sheet with headers on row 1 and a DROP-DOWN INPUTS sheet —
// and are filled the same way (raw buffer re-read, DATA ENTRY rows replaced,
// everything else verbatim so validations and dropdowns survive).
//
// Values are passed through RAW so they match the DROP-DOWN INPUTS lists
// (es-MX, TRUE/FALSE, MALE/FEMALE, "Tomatoes-Other" left whole). The only
// normalisation is dates → YYYY-MM-DD and snapping a value to a dropdown
// entry's exact casing. Anything still off-list is reported, not rewritten.
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const escHtml = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  const str = v => String(v == null ? '' : v).trim();
  // Loose key: lowercase alphanumerics only — drops the '*' required marker,
  // trailing spaces, punctuation and casing.
  const normLoose = h => String(h == null ? '' : h).toLowerCase().replace(/[^a-z0-9]/g, '');

  function isTrue(v) {
    const s = str(v).toLowerCase();
    return s === 'true' || s === 'yes' || s === 'y' || s === '1';
  }

  // Any date shape PickTrace exports → 'YYYY-MM-DD'.
  // '2001-09-02T00:00:00-04:00' is sliced, never passed through new Date(),
  // so the -04:00 offset can't roll the day backwards.
  function toYMD(v) {
    if (v instanceof Date && !isNaN(v)) {
      return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    }
    const t = str(v);
    if (!t) return '';
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      let y = m[3];
      if (y.length === 2) y = (+y >= 70 ? '19' : '20') + y;
      return y + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
    }
    if (/^\d{5}(\.\d+)?$/.test(t)) { // bare Excel serial
      const n = Math.floor(parseFloat(t));
      if (n >= 10000 && n <= 80000) {
        const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
        if (!isNaN(d.getTime())) {
          return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
        }
      }
    }
    return t;
  }

  // ═══ Messy-roster handling ═══
  // A customer-supplied list isn't a PickTrace export: it carries one "Name"
  // column instead of First/Middle/Last, one "ADRESS" blob instead of the
  // address fields, and misspelled headers. Rather than bolt a mapping UI on,
  // these derive VIRTUAL source columns named exactly as the template's, so
  // the ordinary exact-name resolver picks them up unchanged.

  const ROSTER_NAME_KEYS = ['name', 'fullname', 'employeename', 'employee', 'nombre'];
  const ROSTER_ADDR_KEYS = ['adress', 'address', 'addres', 'homeaddress', 'physicaladdress',
    'streetaddress', 'adress1', 'domicilio'];

  function titleCaseName(v) {
    return str(v).replace(/\s+/g, ' ').toLowerCase()
      .replace(/(^|[\s'-])([a-z])/g, (m, sep, ch) => sep + ch.toUpperCase());
  }

  // Western split, as chosen: first token, last token, everything between is
  // the middle. "MARIA M. CISNEROS" → Maria / M. / Cisneros.
  function splitName(full) {
    const t = str(full).replace(/\s+/g, ' ').split(' ').filter(Boolean);
    if (!t.length) return { first: '', middle: '', last: '' };
    if (t.length === 1) return { first: titleCaseName(t[0]), middle: '', last: '' };
    return {
      first: titleCaseName(t[0]),
      middle: t.slice(1, -1).map(titleCaseName).join(' '),
      last: titleCaseName(t[t.length - 1])
    };
  }

  const US_STATES = new Set(['AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','DC']);
  const STREET_SUFFIX = new Set(['st','street','rd','road','ave','av','avenue','blvd','boulevard','dr','drive','ln','lane','way','hwy','highway','ct','court','pl','place','ter','terrace','cir','circle','pkwy','parkway','route','rte','trl','trail','loop','row','sq','square','pike','fwy','expy','plaza']);
  // Secondary-unit markers belong with the street line, not the city.
  const UNIT_MARK = new Set(['spc','unit','apt','ste','suite','bldg','lot','trlr','rm','#']);
  // Spanish street generics LEAD the name ("Calle Pino"), so the street ends
  // one token after them rather than at a trailing suffix.
  const ES_STREET = new Set(['calle','avenida','camino','paseo','via','circulo','callejon']);

  // "1513 BERING AVENUE THERMAL CA, 92274" → street / city / state / zip.
  // Sibling of parseAddress() in sites-standardize.js, extended for the shapes
  // rosters actually contain: PO boxes, SPC/Unit numbers, glued "30THERMAL".
  function parseRosterAddress(raw) {
    const out = { a1: '', city: '', state: '', zip: '', country: '', flag: '' };
    let s = str(raw).replace(/\s+/g, ' ');
    if (!s) return out;

    let m = s.match(/[,\s]\s*(\d{5}(?:-\d{4})?)\s*$/);
    if (m) { out.zip = m[1]; s = s.slice(0, m.index).trim(); }
    m = s.match(/(,?)\s*([A-Za-z]{2})\s*$/);
    if (m && (m[1] === ',' || US_STATES.has(m[2].toUpperCase()))) {
      out.state = m[2].toUpperCase();
      s = s.slice(0, s.length - m[0].length).trim();
    }
    out.country = out.state ? 'US' : '';
    s = s.replace(/,\s*$/, '').trim();
    s = s.replace(/(\d)([A-Za-z]{3,})/g, '$1 $2'); // "Unit 30THERMAL"

    m = s.match(/^(?:P\.?\s*O\.?\s*BOX|P\.?\s*BOX|PO\s*BOX|POBOX)\s*#?\s*(\d+)\s*(.*)$/i);
    if (m) {
      out.a1 = 'PO Box ' + m[1];
      out.city = m[2].trim();
      if (!out.city) out.flag = 'no city';
      return out;
    }
    if (s.indexOf(',') >= 0) {
      const i = s.lastIndexOf(',');
      out.a1 = s.slice(0, i).trim();
      out.city = s.slice(i + 1).trim();
      return out;
    }

    const toks = s.split(' ');
    const bare = t => t.replace(/[.#]/g, '').toLowerCase();
    let cut = -1;
    for (let i = 0; i < toks.length; i++) if (STREET_SUFFIX.has(bare(toks[i]))) cut = i;
    if (cut < 0) {
      for (let i = 0; i < toks.length - 1; i++) if (ES_STREET.has(bare(toks[i]))) { cut = i + 1; break; }
    }
    if (cut < 0) { out.a1 = s; out.flag = 'city not identified'; return out; }

    let end = cut;
    while (end + 1 < toks.length) {
      const nx = toks[end + 1];
      if (UNIT_MARK.has(bare(nx)) || /^#?\d+[a-z]?$/i.test(nx)) end++; else break;
    }
    out.a1 = toks.slice(0, end + 1).join(' ');
    out.city = toks.slice(end + 1).join(' ');
    if (!out.city) out.flag = 'no city';
    return out;
  }

  // PickTrace: "Phone Number should start with 1 or 52 and be followed by 10
  // digits". A bare US 10-digit number gets the 1; anything that can't be made
  // to fit is left alone and reported rather than mangled.
  function normPhone(v) {
    const s = str(v);
    if (!s) return { value: '', ok: true };
    const d = s.replace(/\D/g, '');
    if (d.length === 10) return { value: '1' + d, ok: true };
    if (d.length === 11 && d.charAt(0) === '1') return { value: d, ok: true };
    if (d.length === 12 && d.slice(0, 2) === '52') return { value: d, ok: true };
    return { value: s, ok: false };
  }

  function normSsn(v) {
    const s = str(v);
    const d = s.replace(/\D/g, '');
    return d.length === 9 ? d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5) : s;
  }

  // Append virtual columns for whatever the roster packs into one field.
  // Returns the effective source plus a report of what was derived.
  function deriveColumns(dom, src) {
    if (dom.key !== 'employees') return { headers: src.headers, rows: src.rows, derived: [], consumed: [], flags: [] };
    const keys = src.headers.map(normLoose);
    const has = k => keys.indexOf(k) >= 0;
    const headers = src.headers.slice();
    const rows = src.rows.map(r => r.slice());
    const derived = [], consumed = [], flags = [];

    if (!has('firstname') && !has('lastname')) {
      const ni = keys.findIndex(k => ROSTER_NAME_KEYS.indexOf(k) >= 0);
      if (ni >= 0) {
        const base = headers.length;
        headers.push('First Name', 'Middle Name', 'Last Name');
        rows.forEach(r => {
          const p = splitName(r[ni]);
          r[base] = p.first; r[base + 1] = p.middle; r[base + 2] = p.last;
        });
        consumed.push(ni);
        derived.push({ domain: dom.label, from: src.headers[ni], into: 'First Name · Middle Name · Last Name',
          how: 'name split, title-cased', sample: sampleOf(rows, base, base + 2) });
      }
    }

    if (!has('physicaladdress1')) {
      const ai = keys.findIndex(k => ROSTER_ADDR_KEYS.indexOf(k) >= 0);
      if (ai >= 0) {
        const base = headers.length;
        headers.push('Physical Address 1', 'Physical Address City', 'Physical Address State', 'Physical Address Zip Code');
        rows.forEach((r, ri) => {
          const a = parseRosterAddress(r[ai]);
          r[base] = a.a1; r[base + 1] = a.city; r[base + 2] = a.state; r[base + 3] = a.zip;
          if (a.flag && str(r[ai])) flags.push({ domain: dom.label, reason: a.flag, detail: str(r[ai]), row: ri });
        });
        consumed.push(ai);
        derived.push({ domain: dom.label, from: src.headers[ai], into: 'Physical Address 1 · City · State · Zip Code',
          how: 'address parsed', sample: sampleOf(rows, base, base + 3) });
      }
    }

    return { headers: headers, rows: rows, derived: derived, consumed: consumed, flags: flags };
  }

  function sampleOf(rows, from, to) {
    for (const r of rows) {
      const parts = [];
      for (let c = from; c <= to; c++) if (str(r[c])) parts.push(str(r[c]));
      if (parts.length) return parts.join(' · ');
    }
    return '';
  }

  // Dice coefficient over character bigrams — tolerant of the misspellings
  // rosters carry ("ADRESS"), strict enough not to pair unrelated columns.
  function similarity(a, b) {
    if (a === b) return 1;
    if (a.length < 2 || b.length < 2) return 0;
    const grams = s => { const m = new Map(); for (let i = 0; i < s.length - 1; i++) { const g = s.slice(i, i + 2); m.set(g, (m.get(g) || 0) + 1); } return m; };
    const ga = grams(a), gb = grams(b);
    let hit = 0, na = 0, nb = 0;
    ga.forEach(v => { na += v; });
    gb.forEach(v => { nb += v; });
    ga.forEach((v, g) => { if (gb.has(g)) hit += Math.min(v, gb.get(g)); });
    return (2 * hit) / (na + nb);
  }
  const FUZZY_MIN = 0.72;

  function todayYMD() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  // Whole years between a YYYY-MM-DD birth date and today. null when unparseable.
  const MIN_AGE = 12;
  function ageFrom(ymd) {
    const m = String(ymd == null ? '' : ymd).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const now = new Date();
    let a = now.getFullYear() - (+m[1]);
    const mo = now.getMonth() + 1, d = now.getDate();
    if (mo < (+m[2]) || (mo === (+m[2]) && d < (+m[3]))) a--;
    return a;
  }

  // ─── Domains ───
  // Export and template schemas line up almost exactly, so mapping is an
  // exact name match per column; `aliases` covers the handful that differ.
  const DOMAINS = [
    {
      key: 'employees', label: 'Employees', needsTemplate: true,
      archived: 'Archived',
      aliases: { 'Emergency Phone Number': ['Emergency Contact Phone Number'] },
      dates: ['Date of Birth', 'Hire Date', 'Start Date'],
      numeric: ['Compensation'],
      phones: ['Phone Number', 'Emergency Phone Number'],
      // Columns that stand or fall together. PickTrace treats a partly-filled
      // address as an error, so nothing may be written into one of these
      // unless a sibling already carries data.
      groups: [
        ['Physical Address 1', 'Physical Address 2', 'Physical Address City',
          'Physical Address State', 'Physical Address Zip Code', 'Physical Address Country'],
        ['Mailing Address 1', 'Mailing Address 2', 'Mailing Address City',
          'Mailing Address State', 'Mailing Address Zip Code', 'Mailing Address Country']
      ],
      hint: 'Bulk employee create template. Dropdowns drive Employer, Crew, Gender, Title, Language and H2A validation.'
    },
    {
      key: 'locations', label: 'Locations', needsTemplate: true,
      archived: 'Is Archived',
      aliases: {},
      dates: ['Planted At', 'Start Date', 'Wet Date', 'Germination Date', 'Planted Date',
        'Grafting Date', 'Production Start', 'Organic Certification Date'],
      numeric: ['Acreage', 'Length', 'Plant Count', 'Stand Count', 'Row/Bed Count', 'Post Count',
        'Percent Covered', 'Row Spacing, in.', 'Plant Spacing, in.', 'Post Spacing, in.', 'Bed Width, in.'],
      hint: 'Bulk locations create template. The export is a near 1:1 match — every column but Is Archived carries over.'
    },
    {
      key: 'sites', label: 'Sites', needsTemplate: true,
      archived: 'Location Archived',
      dedupe: 'Name*',
      aliases: { 'Name*': ['Site'] },
      dates: [],
      numeric: [],
      hint: 'Bulk sites create template. The export is one row per location, so it is de-duplicated down to unique site names.'
    },
    {
      key: 'jobs', label: 'Jobs', needsTemplate: false,
      archived: 'Archived',
      aliases: {},
      dates: [],
      numeric: ['Duration Value', 'Hourly Rate', 'Minimum Wage', 'Premium Rate'],
      dropCols: ['Archived'],
      hint: 'No bulk-create template exists for jobs, so this is written as a plain DATA ENTRY sheet with the export\'s own columns.'
    }
  ];
  const DOMAIN_BY_KEY = {};
  DOMAINS.forEach(d => { DOMAIN_BY_KEY[d.key] = d; });

  // DROP-DOWN INPUTS headers that don't match their DATA ENTRY column name.
  const DROPDOWN_ALIAS = { h2aemployee: 'h2a', languagepreference: 'language' };

  // ─── State ───
  let tpls = {};   // key → { fileName, rawBuffer, sheetName, headers, dropdowns }
  let srcs = {};   // key → { fileName, sheetName, headers, rows }
  let built = null;// key → { headers, rows, ... }
  let previewKey = null;
  let migrate = true;  // rewrite off-list values to a single-option destination list
  let setToday = true; // stamp Locations' Start Date with today
  let ageFixes = {};   // "<domain>|<source row index>" → corrected YYYY-MM-DD
  let columnFills = {};// domain → colIdx → { val, mode:'all'|'blank' }  (Column Fill panel)
  let smartFix = {};   // domain → "<colIdx>||<UPPER VALUE>" → canonical dropdown value
  let initialized = false;

  // ═══ Reading ═══
  // One entry point for every upload: works out whether the file is a
  // template or an export, and which domain it belongs to.
  function readAnyFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      reader.onerror = () => reject(new Error('could not read file'));

      if (ext === 'csv') {
        reader.onload = e => {
          const p = parseCSV(e.target.result);
          if (!p.headers.length) { reject(new Error('no header row')); return; }
          resolve({
            kind: 'export', key: detectExport(p.headers),
            data: { headers: p.headers, rows: p.rows, fileName: file.name, sheetName: '' }
          });
        };
        reader.readAsText(file);
        return;
      }

      reader.onload = e => {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
        // Neither the sheet name nor the presence of DROP-DOWN INPUTS tells a
        // template from an export — PickTrace names the locations export's
        // sheet DATA ENTRY, and the employees export ships its own
        // DROP-DOWN INPUTS. The one reliable difference is records: a template
        // is a blank form, an export carries rows.
        const dataName = wb.SheetNames.find(n => /data.?entry/i.test(n));
        const dataAoa = dataName ? XLSX.utils.sheet_to_json(wb.Sheets[dataName], { header: 1, defval: '' }) : null;
        const dataRows = dataAoa ? dataAoa.slice(1).filter(r => r.some(c => str(c) !== '')) : [];
        if (dataName && !dataRows.length) {
          const headers = (dataAoa[0] || []).map(str).filter(Boolean);
          resolve({
            kind: 'template', key: detectTemplate(headers),
            data: {
              fileName: file.name, rawBuffer: buf, sheetName: dataName,
              headers: headers, dropdowns: parseDropdowns(wb)
            }
          });
          return;
        }
        // DATA ENTRY first when present, and never the reference sheet — the
        // employees export's DROP-DOWN INPUTS has 249 rows of its own.
        const order = wb.SheetNames.filter(n => !/drop.?down/i.test(n));
        if (dataName) order.sort((a, b) => (a === dataName ? -1 : b === dataName ? 1 : 0));
        for (const name of order) {
          const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' })
            .filter(r => r.some(c => str(c) !== ''));
          if (aoa.length >= 2) {
            resolve({
              kind: 'export', key: detectExport(aoa[0].map(str)),
              data: {
                headers: aoa[0].map(str),
                rows: aoa.slice(1).map(r => r.map(c => (c instanceof Date ? c : str(c)))),
                fileName: file.name, sheetName: name
              }
            });
            return;
          }
        }
        reject(new Error('no sheet with data rows'));
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function parseDropdowns(wb) {
    const map = new Map();
    const name = wb.SheetNames.find(n => /drop.?down/i.test(n));
    if (!name) return map;
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    (aoa[0] || []).forEach((h, c) => {
      const key = normLoose(h);
      if (!key) return;
      const vals = [];
      for (let r = 1; r < aoa.length; r++) {
        const v = str((aoa[r] || [])[c]);
        if (v && vals.indexOf(v) < 0) vals.push(v);
      }
      // Empty lists are kept: a column the destination org has NO options for
      // is a real finding, and it must be distinguishable from no list at all.
      map.set(key, vals);
    });
    return map;
  }

  function dropdownFor(dropdowns, header) {
    if (!dropdowns) return null;
    const k = normLoose(header);
    return dropdowns.get(k) || dropdowns.get(DROPDOWN_ALIAS[k] || '') || null;
  }

  // Templates and exports for the same domain share column names, so the
  // DATA ENTRY sheet is what marks a file as a template (see readAnyFile).
  function detectTemplate(headers) {
    const h = new Set(headers.map(normLoose));
    if (h.has('cropvariety') && h.has('locationtype')) return 'locations';
    if (h.has('firstname') && h.has('employer') && h.has('dateofbirth')) return 'employees';
    if (h.has('sitetype') && h.has('address1') && h.has('name')) return 'sites';
    return null;
  }
  function detectExport(headers) {
    const h = new Set(headers.map(normLoose));
    if (h.has('cropvariety') || (h.has('acreage') && h.has('plantcount'))) return 'locations';
    if (h.has('group') && h.has('site') && h.has('location') && (h.has('groupid') || h.has('locationid'))) return 'sites';
    if (h.has('firstname') && h.has('lastname')) return 'employees';
    if (h.has('name') && h.has('category') && (h.has('isproductive') || h.has('paystyle'))) return 'jobs';
    return null;
  }

  // ═══ Mapping ═══
  function mapDomain(dom) {
    const src = srcs[dom.key];
    if (!src) return null;
    const tpl = tpls[dom.key] || null;
    if (dom.needsTemplate && !tpl) return null;

    // Without a template the export's own columns are the schema.
    const drop = new Set((dom.dropCols || []).map(normLoose));
    const headers = tpl ? tpl.headers : src.headers.filter(h => str(h) && !drop.has(normLoose(h)));
    const dropdowns = tpl ? tpl.dropdowns : new Map();

    // A messy roster gets virtual columns (name split, address parsed) named
    // exactly as the template's, so the resolver below needs no special case.
    const der = deriveColumns(dom, src);
    const eff = { headers: der.headers, rows: der.rows };

    // Resolve each output column to a source column: exact name, then alias,
    // then a reported fuzzy guess. Exact and alias run to completion first so
    // a fuzzy guess can never steal a column an exact match wants.
    const srcIdx = new Map();
    eff.headers.forEach((h, i) => { const k = normLoose(h); if (k && !srcIdx.has(k)) srcIdx.set(k, i); });
    const claimedSrc = new Set(der.consumed);
    const colSrc = headers.map(h => {
      const k = normLoose(h);
      if (srcIdx.has(k)) { claimedSrc.add(srcIdx.get(k)); return srcIdx.get(k); }
      const al = dom.aliases[h] || dom.aliases[h.replace(/\*\s*$/, '').trim()];
      if (al) for (const a of al) {
        const ak = normLoose(a);
        if (srcIdx.has(ak)) { claimedSrc.add(srcIdx.get(ak)); return srcIdx.get(ak); }
      }
      return -1;
    });
    const fuzzy = [];
    headers.forEach((h, ci) => {
      if (colSrc[ci] >= 0) return;
      const k = normLoose(h).replace(/\*$/, '');
      let best = -1, bestScore = 0;
      eff.headers.forEach((sh, si) => {
        if (claimedSrc.has(si) || !str(sh)) return;
        const sc = similarity(k, normLoose(sh));
        if (sc > bestScore) { bestScore = sc; best = si; }
      });
      if (best >= 0 && bestScore >= FUZZY_MIN) {
        colSrc[ci] = best;
        claimedSrc.add(best);
        fuzzy.push({ domain: dom.label, column: h, from: eff.headers[best], score: Math.round(bestScore * 100) });
      }
    });

    const dateSet = new Set((dom.dates || []).map(normLoose));
    const archivedIdx = srcIdx.has(normLoose(dom.archived)) ? srcIdx.get(normLoose(dom.archived)) : -1;
    const dedupeCol = dom.dedupe ? headers.findIndex(h => normLoose(h) === normLoose(dom.dedupe)) : -1;

    const dobCol = headers.findIndex(h => normLoose(h) === 'dateofbirth');
    const ssnCol = headers.findIndex(h => normLoose(h) === 'ssn');
    const phoneCols = (dom.phones || []).map(p => headers.findIndex(h => normLoose(h) === normLoose(p))).filter(i => i >= 0);
    const phoneBad = [];

    const rows = [], rowSrc = [], dropped = [], seen = new Set();
    eff.rows.forEach((sr, si0) => {
      if (archivedIdx >= 0 && isTrue(sr[archivedIdx])) {
        dropped.push({ domain: dom.label, reason: dom.archived + ' = true', detail: describeRow(headers, colSrc, sr) });
        return;
      }
      const row = headers.map((h, ci) => {
        const si = colSrc[ci];
        if (si < 0) return '';
        const raw = sr[si];
        return dateSet.has(normLoose(h)) ? toYMD(raw) : str(raw);
      });
      if (!row.some(v => v !== '')) return;
      if (ssnCol >= 0 && row[ssnCol]) row[ssnCol] = normSsn(row[ssnCol]);
      phoneCols.forEach(pi => {
        if (!row[pi]) return;
        const p = normPhone(row[pi]);
        if (p.ok) row[pi] = p.value;
        else phoneBad.push({ domain: dom.label, column: headers[pi], value: row[pi] });
      });
      if (dedupeCol >= 0) {
        const k = row[dedupeCol].toUpperCase();
        if (!k || seen.has(k)) return;
        seen.add(k);
      }
      // A corrected birth date you entered in the Under Minimum Age panel wins
      // over the export's, and is keyed to the source row so it survives rebuilds.
      if (dobCol >= 0) {
        const fix = ageFixes[dom.key + '|' + si0];
        if (fix) row[dobCol] = fix;
      }
      rows.push(row);
      rowSrc.push(si0);
    });

    // Locations get today's date as their Start Date — a fresh import starts
    // now, not on whatever season date the source account recorded.
    const overrides = [];
    if (setToday && dom.key === 'locations') {
      const sdCol = headers.findIndex(h => normLoose(h) === 'startdate');
      if (sdCol >= 0 && rows.length) {
        const was = [];
        rows.forEach(r => { if (r[sdCol] && was.indexOf(r[sdCol]) < 0) was.push(r[sdCol]); });
        const today = todayYMD();
        rows.forEach(r => { r[sdCol] = today; });
        overrides.push({ domain: dom.label, column: headers[sdCol], from: was.length ? was : ['(blank)'], to: today });
      }
    }

    // Sticky Smart Fixes — an off-list value the operator mapped to a real
    // template value, re-applied to every matching cell in that column.
    const sfx = smartFix[dom.key] || {};
    if (Object.keys(sfx).length) {
      rows.forEach(r => {
        for (let ci = 0; ci < headers.length; ci++) {
          const v = r[ci];
          if (!v) continue;
          const hit = sfx[ci + '||' + String(v).toUpperCase().trim()];
          if (hit != null) r[ci] = hit;
        }
      });
    }

    // Sticky column fills from the Column Fill panel. Applied before the
    // dropdown pass so a filled value is case-snapped like any other.
    const fills = columnFills[dom.key] || {};
    const filledCols = [];
    Object.keys(fills).forEach(k => {
      const ci = Number(k), f = fills[k];
      if (!(ci >= 0 && ci < headers.length)) return;
      let n = 0;
      rows.forEach(r => {
        if (f.mode === 'blank' && r[ci] !== '') return;
        if (r[ci] !== f.val) n++;
        r[ci] = f.val;
      });
      filledCols.push({ domain: dom.label, column: headers[ci], value: f.val === '' ? '(cleared)' : f.val,
        mode: f.mode === 'blank' ? 'blanks only' : 'all rows', rows: n });
    });

    // Snap values to the dropdown's exact casing (false → FALSE), then handle
    // whatever is still off-list.
    //
    // When the destination list offers exactly ONE option, an off-list value is
    // migrated to it — the employees are moving orgs, so an export that says
    // Employer "Old Grower Co., LLC" belongs in a template whose
    // only allowed Employer is "New Grower Co., LLC". Every rewrite is reported, and the
    // Migrate toggle turns this off.
    //
    // An EMPTY destination list is different: the org has no options at all for
    // that column (Crew, H2A Contract), so there's nothing to migrate to and
    // the values have to be created in PickTrace first.
    const mismatches = [], remapped = [], noOptions = [];
    headers.forEach((h, ci) => {
      const list = dropdownFor(dropdowns, h);
      if (!list) return;
      const used = [];
      rows.forEach(r => { if (r[ci] && used.indexOf(r[ci]) < 0) used.push(r[ci]); });
      if (!list.length) {
        // Enumerate each distinct value with its row count — this list is a
        // build checklist for PickTrace, so it must not be summarised away.
        used.forEach(v => {
          noOptions.push({ domain: dom.label, column: h, value: v, count: rows.filter(r => r[ci] === v).length });
        });
        return;
      }
      const caseMap = new Map();
      list.forEach(v => { const k = v.toUpperCase(); if (!caseMap.has(k)) caseMap.set(k, v); });
      const off = [];
      rows.forEach(r => {
        const v = r[ci];
        if (!v) return;
        const k = v.toUpperCase();
        if (caseMap.has(k)) { r[ci] = caseMap.get(k); return; }
        if (off.indexOf(v) < 0) off.push(v);
      });
      if (!off.length) return;
      if (migrate && list.length === 1) {
        rows.forEach(r => { if (r[ci] && r[ci] !== list[0]) r[ci] = list[0]; });
        remapped.push({ domain: dom.label, column: h, from: off, to: list[0] });
      } else {
        // ci + domain key travel with the finding so the Smart Fixes panel can
        // write a sticky correction back to the exact column it came from.
        mismatches.push({ domain: dom.label, key: dom.key, ci: ci, column: h, values: off, allowed: list,
          counts: off.map(v => rows.filter(r => r[ci] === v).length) });
      }
    });

    // A column with no data at all, whose dropdown offers exactly one choice,
    // can only be that choice — fill it rather than leaving a required gap.
    //
    // But only on rows where the column's field group already has data. A
    // Country stamped onto a row with no address is what made PickTrace reject
    // 20 rows with "Incomplete Physical Address"; those rows must stay blank.
    const groupPeers = ci => {
      const k = normLoose(headers[ci]);
      const g = (dom.groups || []).find(grp => grp.some(x => normLoose(x) === k));
      if (!g) return null;
      const set = new Set(g.map(normLoose));
      return headers.map((h, i) => (i !== ci && set.has(normLoose(h)) ? i : -1)).filter(i => i >= 0);
    };
    const autofilled = [];
    headers.forEach((h, ci) => {
      if (rows.some(r => r[ci] !== '')) return;
      const list = dropdownFor(dropdowns, h);
      if (!list || list.length !== 1) return;
      const peers = groupPeers(ci);
      let n = 0;
      rows.forEach(r => {
        if (peers && !peers.some(pi => r[pi] !== '')) return;
        r[ci] = list[0];
        n++;
      });
      if (n) autofilled.push({ domain: dom.label, column: h, value: list[0], rows: n, total: rows.length });
    });

    // Required (*) columns still blank after all that.
    const required = [];
    headers.forEach((h, ci) => {
      if (!/\*\s*$/.test(h)) return;
      let empty = 0;
      rows.forEach(r => { if (r[ci] === '') empty++; });
      if (empty) required.push({ domain: dom.label, column: h, empty: empty, total: rows.length });
    });

    // Source columns holding data that no output column claims. Columns the
    // derivation swallowed (Name, ADRESS) are claimed, not orphaned, and the
    // virtual columns it produced are never reported as source columns.
    const claimed = new Set(colSrc.filter(i => i >= 0));
    der.consumed.forEach(i => claimed.add(i));
    if (archivedIdx >= 0) claimed.add(archivedIdx);
    const unmapped = [];
    src.headers.forEach((h, i) => {
      if (!str(h) || claimed.has(i)) return;
      const vals = [];
      for (const r of eff.rows) {
        const v = str(r[i]);
        if (v && vals.indexOf(v) < 0) { vals.push(v); if (vals.length >= 4) break; }
      }
      if (vals.length) unmapped.push({ domain: dom.label, column: h, sample: vals });
    });

    // Anyone below the minimum working age — almost always a typo'd birth year.
    // Listed for correction in the panel; never silently altered or removed.
    const underAge = [];
    if (dobCol >= 0) {
      const nameCols = headers.map((h, i) => ({ i, k: normLoose(h) }))
        .filter(x => x.k === 'firstname' || x.k === 'lastname').map(x => x.i);
      rows.forEach((r, ri) => {
        const a = ageFrom(r[dobCol]);
        if (a === null || a >= MIN_AGE) return;
        const name = nameCols.map(i => r[i]).filter(Boolean).join(' ') || '(row ' + (ri + 1) + ')';
        underAge.push({ domain: dom.label, key: dom.key + '|' + rowSrc[ri], name: name, dob: r[dobCol], age: a });
      });
    }

    return { headers: headers, colSrc: colSrc, rows: rows, rowSrc: rowSrc, dropped: dropped,
      mismatches: mismatches, remapped: remapped.concat(overrides), noOptions: noOptions,
      autofilled: autofilled, required: required, unmapped: unmapped, underAge: underAge,
      derived: der.derived, addrFlags: der.flags, fuzzy: fuzzy, filledCols: filledCols,
      phoneBad: phoneBad, dropdowns: dropdowns, key: dom.key, hasTemplate: !!tpl };
  }

  function describeRow(headers, colSrc, sr) {
    for (let i = 0; i < headers.length; i++) {
      if (colSrc[i] >= 0 && str(sr[colSrc[i]])) return str(sr[colSrc[i]]);
    }
    return '(row)';
  }

  // ═══ Build ═══
  function build() {
    const out = {};
    DOMAINS.forEach(d => { const m = mapDomain(d); if (m) out[d.key] = m; });
    built = out;
    if (!previewKey || !built[previewKey]) previewKey = Object.keys(built)[0] || null;
    render();
  }

  // ═══ Render ═══
  function render() {
    const keys = Object.keys(built || {});
    $('dt-empty').style.display = keys.length ? 'none' : '';
    renderSummary(); renderPreview(); renderBuildBanner(); renderUnderAge();
    renderSmartFix(); renderFill();
    renderList('fills', collect('filledCols'));
    renderList('derived', collect('derived'));
    renderList('fuzzy', collect('fuzzy'));
    renderList('addrflag', collect('addrFlags'));
    renderList('phonebad', collect('phoneBad'));
    renderList('mismatch', collect('mismatches'));
    renderList('remap', collect('remapped'));
    renderList('nooptions', collect('noOptions'));
    renderList('autofill', collect('autofilled'));
    renderList('required', collect('required'));
    renderList('unmapped', collect('unmapped'));
    renderList('dropped', collect('dropped'));
    DOMAINS.forEach(d => {
      const btn = $('dt-dl-' + d.key);
      if (btn) btn.disabled = !(built && built[d.key] && built[d.key].rows.length);
    });
    $('dt-dl-all').disabled = !keys.some(k => built[k].rows.length);
  }

  function collect(field) {
    const out = [];
    Object.keys(built || {}).forEach(k => { out.push(...(built[k][field] || [])); });
    return out;
  }

  function renderSummary() {
    const el = $('dt-summary');
    const keys = Object.keys(built || {});
    if (!keys.length) { el.style.display = 'none'; return; }
    let html = DOMAINS.filter(d => built[d.key]).map(d =>
      '<span class="cmp-stat"><b>' + built[d.key].rows.length + '</b> ' + escHtml(d.label.toLowerCase()) + '</span>').join('');
    const waiting = DOMAINS.filter(d => !built[d.key] && (srcs[d.key] || tpls[d.key])).map(d =>
      d.label + ' (' + (!srcs[d.key] ? 'no export' : 'no template') + ')');
    if (waiting.length) html += '<span class="cmp-stat cmp-warn">waiting: ' + escHtml(waiting.join(', ')) + '</span>';
    const bad = collect('mismatches').length;
    if (bad) html += '<span class="cmp-stat cmp-warn">⚠ ' + bad + ' column(s) with off-dropdown values</span>';
    const mig = collect('remapped').length;
    if (mig) html += '<span class="cmp-stat">' + mig + ' column(s) migrated</span>';
    const none = collect('noOptions').length;
    if (none) html += '<span class="cmp-stat cmp-warn">⚠ ' + none + ' value(s) to build in PickTrace</span>';
    const young = collect('underAge').length;
    if (young) html += '<span class="cmp-stat cmp-warn">⚠ ' + young + ' under age ' + MIN_AGE + '</span>';
    el.innerHTML = html;
    el.style.display = '';
  }

  function renderPreview() {
    const keys = DOMAINS.map(d => d.key).filter(k => built && built[k]);
    const sec = $('dt-section-preview');
    if (!keys.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    if (keys.indexOf(previewKey) < 0) previewKey = keys[0];
    $('dt-preview-pills').innerHTML = keys.map(k =>
      '<button type="button" class="btn btn-sm ts-domain-pill' + (k === previewKey ? ' ts-domain-active' : '') +
      '" data-key="' + k + '">' + escHtml(DOMAIN_BY_KEY[k].label) + ' <b>' + built[k].rows.length + '</b></button>').join('');

    const b = built[previewKey], CAP = 200;
    let html = '<thead><tr><th>#</th>' + b.headers.map((h, ci) =>
      '<th' + (b.colSrc[ci] < 0 ? ' class="cmp-warn"' : '') + '>' + escHtml(h) + '</th>').join('') + '</tr></thead><tbody>';
    b.rows.slice(0, CAP).forEach((r, i) => {
      html += '<tr><td>' + (i + 1) + '</td>' + r.map(v => '<td>' + escHtml(v) + '</td>').join('') + '</tr>';
    });
    $('dt-preview-table').innerHTML = html + '</tbody>';
    const dom = DOMAIN_BY_KEY[previewKey];
    const dest = b.hasTemplate ? '"' + tpls[previewKey].sheetName + '" in ' + tpls[previewKey].fileName : 'a generated DATA ENTRY sheet';
    $('dt-preview-note').textContent =
      (b.rows.length > CAP ? 'Showing first ' + CAP + ' of ' + b.rows.length + ' rows — all are exported. ' : '') +
      b.rows.length + ' row' + (b.rows.length === 1 ? '' : 's') + ' → ' + dest + ', from row 2. ' +
      'Greyed-amber headers have no source column.' + (dom.needsTemplate ? '' : ' No template needed.');
  }

  // Standing checklist of what has to exist in PickTrace before these files
  // will import. Deliberately does NOT block the download.
  //
  // Grouped domain → column → values rather than one flat list: a run of 24
  // crop varieties otherwise buries the two H2A contracts that actually gate
  // the employee upload. Values sort biggest-first and long lists collapse.
  const BUILD_CAP = 8;
  // Only the columns worth acting on get banner space. Everything else
  // (Rootstock, Mulch Type, Growing Manager, …) is optional metadata that
  // buried the four blockers — it stays in the Needs Building table below.
  const BUILD_COLUMNS = ['h2acontract', 'crew', 'site', 'cropvariety'];
  // Rendered as crop → varieties instead of a flat list.
  const BUILD_NESTED = ['cropvariety'];

  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');

  // "Blackberries-Orus 4663-1" → crop "Blackberries", variety "Orus 4663-1".
  // Split on the FIRST hyphen: crop names here never contain one, but variety
  // names do, so splitting on the last would produce "Blackberries-Orus 4663".
  function splitCrop(v) {
    const i = String(v).indexOf('-');
    return i > 0
      ? { crop: v.slice(0, i).trim(), variety: v.slice(i + 1).trim() }
      : { crop: String(v).trim(), variety: '' };
  }

  function renderBuildBanner() {
    const el = $('dt-build-banner');
    if (!el) return;
    const all = collect('noOptions');
    const list = all.filter(n => BUILD_COLUMNS.indexOf(normLoose(n.column)) >= 0);
    if (!list.length) { el.style.display = 'none'; el.innerHTML = ''; return; }

    const byDomain = new Map();
    list.forEach(n => {
      if (!byDomain.has(n.domain)) byDomain.set(n.domain, new Map());
      const cols = byDomain.get(n.domain);
      if (!cols.has(n.column)) cols.set(n.column, []);
      cols.get(n.column).push({ value: n.value, count: n.count });
    });
    let totalCols = 0;
    byDomain.forEach(cols => { totalCols += cols.size; });
    const other = all.length - list.length;

    let html = '<div class="dt-build-head">' +
      '<div class="dt-build-title">Build these in PickTrace first</div>' +
      '<div class="dt-build-sub"><b>' + plural(list.length, 'value') + '</b> across <b>' +
      plural(totalCols, 'column') + '</b> have no option in the destination template, so rows using them are ' +
      'rejected on import. Create them in PickTrace, then re-download the template. The files still download either way.' +
      (other ? ' <span class="dt-build-aside">' + plural(other, 'other value') +
        ' in optional columns are listed in Needs Building below.</span>' : '') +
      '</div></div>';

    byDomain.forEach((cols, domain) => {
      let vals = 0, rows = 0;
      cols.forEach(vs => { vals += vs.length; vs.forEach(v => { rows += v.count; }); });
      html += '<section class="dt-build-group">' +
        '<header class="dt-build-group-head">' +
        '<span class="dt-build-domain">' + escHtml(domain) + '</span>' +
        '<span class="dt-build-tally">' + plural(cols.size, 'column') + ' · ' +
        plural(vals, 'value') + ' to create · ' + plural(rows, 'row') + ' affected</span>' +
        '</header><div class="dt-build-cols">';
      cols.forEach((vs, col) => {
        html += (BUILD_NESTED.indexOf(normLoose(col)) >= 0 ? nestedCard : flatCard)(col, vs);
      });
      html += '</div></section>';
    });

    el.innerHTML = html;
    el.style.display = '';
  }

  function flatCard(col, vs) {
    vs.sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
    const hidden = vs.length - BUILD_CAP;
    return '<div class="dt-build-col">' +
      '<div class="dt-build-col-head"><span class="dt-build-col-name">' + escHtml(col) + '</span>' +
      '<span class="dt-build-col-n">' + plural(vs.length, 'value') + '</span></div>' +
      '<ul class="dt-build-vals">' +
      vs.map((v, i) =>
        '<li' + (i >= BUILD_CAP ? ' class="dt-build-extra"' : '') + '>' +
        '<span class="dt-build-val">' + escHtml(v.value) + '</span>' +
        '<span class="dt-build-rows">' + plural(v.count, 'row') + '</span></li>').join('') +
      '</ul>' +
      (hidden > 0 ? '<button type="button" class="dt-build-toggle" data-more="' + hidden + '">Show ' + hidden + ' more</button>' : '') +
      '</div>';
  }

  // Crop & Variety, grouped so you can see every variety that belongs to one
  // crop together. Crops order by total rows; varieties within a crop likewise.
  function nestedCard(col, vs) {
    const crops = new Map();
    vs.forEach(v => {
      const s = splitCrop(v.value);
      if (!crops.has(s.crop)) crops.set(s.crop, { rows: 0, vars: [] });
      const g = crops.get(s.crop);
      g.rows += v.count;
      g.vars.push({ name: s.variety || '(no variety)', count: v.count, full: v.value });
    });
    const ordered = [...crops.entries()].sort((a, b) => b[1].rows - a[1].rows || a[0].localeCompare(b[0]));
    return '<div class="dt-build-col dt-build-col-wide">' +
      '<div class="dt-build-col-head"><span class="dt-build-col-name">' + escHtml(col) + '</span>' +
      '<span class="dt-build-col-n">' + plural(ordered.length, 'crop') + ' · ' + plural(vs.length, 'variety').replace('varietys', 'varieties') + '</span></div>' +
      '<div class="dt-build-crops">' +
      ordered.map(([crop, g]) => {
        g.vars.sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
        return '<div class="dt-build-crop">' +
          '<div class="dt-build-crop-head"><span class="dt-build-crop-name">' + escHtml(crop) + '</span>' +
          '<span class="dt-build-crop-n">' + g.vars.length + ' · ' + plural(g.rows, 'row') + '</span></div>' +
          '<ul class="dt-build-vals">' +
          g.vars.map(v => '<li title="' + escHtml(v.full) + '">' +
            '<span class="dt-build-val">' + escHtml(v.name) + '</span>' +
            '<span class="dt-build-rows">' + plural(v.count, 'row') + '</span></li>').join('') +
          '</ul></div>';
      }).join('') +
      '</div></div>';
  }

  function renderUnderAge() {
    const sec = $('dt-section-underage');
    if (!sec) return;
    const list = collect('underAge');
    const fixed = Object.keys(ageFixes).length;
    if (!list.length && !fixed) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('dt-underage-title').textContent = list.length
      ? 'Under Minimum Age (' + list.length + ')'
      : 'Under Minimum Age — all corrected';
    $('dt-underage-table').innerHTML = list.length
      ? '<thead><tr><th>Domain</th><th>Employee</th><th>Age</th><th>Date of Birth</th></tr></thead><tbody>' +
        list.map(u => '<tr><td>' + escHtml(u.domain) + '</td><td>' + escHtml(u.name) + '</td>' +
          '<td class="cmp-warn">' + u.age + '</td>' +
          '<td><input type="date" class="input-field dt-age-row" data-key="' + escHtml(u.key) +
          '" value="' + escHtml(u.dob) + '" style="width:150px;"></td></tr>').join('') + '</tbody>'
      : '<thead><tr><th>Status</th></tr></thead><tbody><tr><td>' + fixed +
        ' corrected birth date' + (fixed === 1 ? '' : 's') + ' applied. Clear them to review again.</td></tr></tbody>';
  }

  // ─── Smart Fixes ───
  // Every off-list value gets a picker of the template's allowed values, with
  // the closest match pre-selected. Mirrors template-standardize.js.
  function suggestFor(value, allowed) {
    const v = normLoose(value);
    let best = '', score = 0;
    allowed.forEach(a => { const s = similarity(v, normLoose(a)); if (s > score) { score = s; best = a; } });
    return score >= 0.6 ? best : '';
  }

  function renderSmartFix() {
    const sec = $('dt-section-smartfix');
    if (!sec) return;
    const list = collect('mismatches');
    if (!list.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    let n = 0;
    list.forEach(m => { n += m.values.length; });
    $('dt-smartfix-title').textContent = 'Smart Fixes (' + n + ')';
    let html = '<thead><tr><th>Domain</th><th>Column</th><th>Value in your data</th><th>Rows</th>' +
      '<th>Write instead</th><th></th></tr></thead><tbody>';
    list.forEach(m => {
      m.values.forEach((v, vi) => {
        const sug = suggestFor(v, m.allowed);
        const opts = '<option value="">— leave as-is —</option>' +
          m.allowed.map(a => '<option' + (a === sug ? ' selected' : '') + '>' + escHtml(a) + '</option>').join('');
        html += '<tr><td>' + escHtml(m.domain) + '</td><td>' + escHtml(m.column) + '</td>' +
          '<td class="cmp-warn">' + escHtml(v) + '</td><td>' + (m.counts ? m.counts[vi] : '') + '</td>' +
          '<td><select class="input-field dt-sf-pick" data-key="' + escHtml(m.key) + '" data-ci="' + m.ci +
          '" data-from="' + escHtml(v) + '" style="min-width:190px;">' + opts + '</select></td>' +
          '<td><button class="btn btn-primary btn-sm dt-sf-apply">Apply</button></td></tr>';
      });
    });
    $('dt-smartfix-table').innerHTML = html + '</tbody>';
  }

  // ─── Column Fill ───
  // Bulk-set or clear any column of the previewed domain. Required columns that
  // are still empty are listed first; everything else sits behind a toggle.
  function renderFill() {
    const sec = $('dt-section-fill');
    if (!sec) return;
    const b = built && previewKey ? built[previewKey] : null;
    if (!b || !b.rows.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const dd = b.dropdowns || new Map();
    const fills = columnFills[previewKey] || {};

    const cols = b.headers.map((h, ci) => {
      let empty = 0, sample = '';
      b.rows.forEach(r => { if (r[ci] === '') empty++; else if (!sample) sample = r[ci]; });
      return { h: h, ci: ci, empty: empty, sample: sample, req: /\*\s*$/.test(h) };
    });
    // Every column is actionable. Required-and-empty still leads the table, but
    // nothing is collapsed — use the filter box to reach a specific column.
    const shown = cols.filter(c => {
      if (fillFilter && c.h.toLowerCase().indexOf(fillFilter) < 0) return false;
      if (fillScope === 'req') return c.req;
      if (fillScope === 'empty') return c.empty > 0;
      if (fillScope === 'filled') return c.empty === 0;
      return true;
    });
    const need = shown.filter(c => c.req && c.empty > 0);
    const rest = shown.filter(c => !(c.req && c.empty > 0));
    const reqEmpty = cols.filter(c => c.req && c.empty > 0).length;

    $('dt-fill-title').textContent = 'Column Fill — ' + DOMAIN_BY_KEY[previewKey].label +
      ' · ' + shown.length + ' of ' + cols.length + ' columns' +
      (reqEmpty ? ' · ' + reqEmpty + ' required still empty' : '');

    const rowFor = c => {
      const list = dropdownFor(dd, c.h) || [];
      const cur = fills[c.ci] ? fills[c.ci].val : '';
      const opts = list.length
        ? '<option value="">— pick a value —</option>' +
          list.map(o => '<option' + (o === cur ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no dropdown — type an override)</option>';
      const status = c.empty === 0
        ? '<span class="dt-ok">✓ all ' + b.rows.length + '</span>'
        : (c.empty === b.rows.length
          ? '<span class="' + (c.req ? 'dt-bad' : 'cmp-warn') + '">' + (c.req ? '⚠ ' : '') + 'all empty</span>'
          : '<span class="cmp-warn">' + c.empty + ' empty</span>');
      return '<tr data-ci="' + c.ci + '"><td><b>' + escHtml(c.h) + '</b></td><td>' + status + '</td>' +
        '<td class="dt-fill-sample">' + escHtml(String(c.sample).slice(0, 28)) + '</td>' +
        '<td><select class="input-field dt-fill-pick" style="min-width:170px;">' + opts + '</select></td>' +
        '<td><input type="text" class="input-field dt-fill-text" placeholder="Manual override" value="' +
        escHtml(list.length ? '' : cur) + '" style="width:160px;"></td>' +
        '<td><button class="btn btn-primary btn-sm dt-fill-all">All rows</button></td>' +
        '<td><button class="btn btn-ghost btn-sm dt-fill-blank" title="Only write where the cell is empty">Blanks</button></td>' +
        '<td><button class="btn btn-ghost btn-sm dt-fill-clear" title="Empty this column">Clear</button></td></tr>';
    };
    let html = '<thead><tr><th>Column</th><th>Status</th><th>Example</th><th>Pick from dropdown</th>' +
      '<th>Manual override</th><th colspan="3">Apply to</th></tr></thead>';
    if (!shown.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band">No column matches this filter.</td></tr></tbody>';
    }
    if (need.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band dt-fill-band-bad">⚠ Required and still empty (' +
        need.length + ')</td></tr>' + need.map(rowFor).join('') + '</tbody>';
    }
    if (rest.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band">' +
        (need.length ? 'Every other column' : 'Columns') + ' (' + rest.length + ')</td></tr>' +
        rest.map(rowFor).join('') + '</tbody>';
    }
    $('dt-fill-table').innerHTML = html;
  }
  let fillFilter = '';   // Column Fill search box, lowercased
  let fillScope = 'all'; // all | req | empty | filled

  const LIST_SPECS = {
    mismatch: { title: 'Off-Dropdown Values', cols: ['Domain', 'Column', 'Value(s) not in the list', 'Allowed'],
      row: m => [m.domain, m.column, m.values.slice(0, 6).join(' · ') + (m.values.length > 6 ? ' … +' + (m.values.length - 6) : ''),
        m.allowed.slice(0, 4).join(' · ') + (m.allowed.length > 4 ? ' … +' + (m.allowed.length - 4) : '')] },
    fills: { title: 'Column Fills Applied', cols: ['Domain', 'Column', 'Value', 'Applied to', 'Rows changed'],
      row: f => [f.domain, f.column, f.value, f.mode, f.rows] },
    derived: { title: 'Derived Columns', cols: ['Domain', 'Source column', 'Split into', 'How', 'Example'],
      row: d => [d.domain, d.from, d.into, d.how, d.sample] },
    fuzzy: { title: 'Fuzzy Column Matches', cols: ['Domain', 'Template column', 'Matched to', 'Confidence'],
      row: f => [f.domain, f.column, f.from, f.score + '%'] },
    addrflag: { title: 'Addresses Needing a Look', cols: ['Domain', 'Problem', 'Address'],
      row: a => [a.domain, a.reason, a.detail] },
    remap: { title: 'Migrated Values', cols: ['Domain', 'Column', 'Was', 'Written as'],
      row: r => [r.domain, r.column, r.from.join(' · '), r.to] },
    nooptions: { title: 'Needs Building in PickTrace', cols: ['Domain', 'Column', 'Value to create', 'Rows'],
      row: n => [n.domain, n.column, n.value, n.count] },
    autofill: { title: 'Auto-Filled Columns', cols: ['Domain', 'Column', 'Value', 'Rows'],
      row: a => [a.domain, a.column, a.value, a.rows + ' of ' + a.total] },
    phonebad: { title: 'Phone Numbers PickTrace Will Reject', cols: ['Domain', 'Column', 'Value'],
      row: p => [p.domain, p.column, p.value] },
    required: { title: 'Empty Required Columns', cols: ['Domain', 'Column', 'Empty'],
      row: q => [q.domain, q.column, q.empty + ' of ' + q.total] },
    unmapped: { title: 'Unmapped Source Columns', cols: ['Domain', 'Column', 'Sample values'],
      row: u => [u.domain, u.column, u.sample.join(' · ')] },
    dropped: { title: 'Dropped Rows', cols: ['Domain', 'Reason', 'Row'], row: d => [d.domain, d.reason, d.detail] }
  };

  function renderList(name, list) {
    const sec = $('dt-section-' + name);
    if (!sec) return;
    if (!list.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const spec = LIST_SPECS[name];
    $('dt-' + name + '-title').textContent = spec.title + ' (' + list.length + ')';
    $('dt-' + name + '-table').innerHTML =
      '<thead><tr>' + spec.cols.map(c => '<th>' + escHtml(c) + '</th>').join('') + '</tr></thead><tbody>' +
      list.map(item => '<tr>' + spec.row(item).map(c => '<td>' + escHtml(c) + '</td>').join('') + '</tr>').join('') +
      '</tbody>';
  }

  // ═══ Export ═══
  // Mirrors template-standardize.js: re-read the template so styles, column
  // widths, DROP-DOWN INPUTS and data validations come along untouched, then
  // replace only the DATA ENTRY rows below the header.
  function download(key) {
    const b = built && built[key];
    if (!b || !b.rows.length) return;
    const dom = DOMAIN_BY_KEY[key];
    const numeric = new Set((dom.numeric || []).map(normLoose));

    let wb, ws, outName;
    if (b.hasTemplate) {
      const tpl = tpls[key];
      wb = XLSX.read(new Uint8Array(tpl.rawBuffer), { type: 'array', cellStyles: true, cellDates: true, sheetStubs: true });
      ws = wb.Sheets[tpl.sheetName];
      if (!ws) { alert('Template is missing its DATA ENTRY sheet — cannot export.'); return; }
      const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: b.headers.length - 1 } };
      for (let r = 1; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const ref = XLSX.utils.encode_cell({ r: r, c: c });
          if (ws[ref]) delete ws[ref];
        }
      }
      outName = withSuffix(tpl.fileName);
    } else {
      wb = XLSX.utils.book_new();
      ws = XLSX.utils.aoa_to_sheet([b.headers]);
      XLSX.utils.book_append_sheet(wb, ws, 'DATA ENTRY');
      outName = withSuffix((srcs[key] && srcs[key].fileName) || dom.label + '.xlsx');
    }

    const colCount = b.headers.length;
    b.rows.forEach((row, ri) => {
      for (let c = 0; c < colCount; c++) {
        const val = row[c];
        if (val == null || val === '') continue;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c: c });
        // Numbers stay numbers where PickTrace expects them; everything else
        // is text so leading-zero IDs and ZIPs survive.
        if (numeric.has(normLoose(b.headers[c])) && /^-?\d+(\.\d+)?$/.test(String(val))) {
          ws[ref] = { v: parseFloat(val), t: 'n' };
        } else {
          ws[ref] = { v: String(val), t: 's' };
        }
      }
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: b.rows.length, c: colCount - 1 } });
    XLSX.writeFile(wb, outName, { cellStyles: true });
  }

  function withSuffix(fileName) {
    const n = str(fileName) || 'template.xlsx';
    const dot = n.lastIndexOf('.');
    const base = dot > 0 ? n.slice(0, dot) : n;
    return base + ' — filled.xlsx';
  }

  function downloadAll() {
    // Sequential, slightly staggered — browsers drop back-to-back saves.
    const keys = DOMAINS.map(d => d.key).filter(k => built && built[k] && built[k].rows.length);
    keys.forEach((k, i) => setTimeout(() => download(k), i * 400));
  }

  function stash(obj, key) { if (!obj[key]) obj[key] = {}; return obj[key]; }
  function applySmartFix(key, ci, from, to) {
    stash(smartFix, key)[ci + '||' + String(from).toUpperCase().trim()] = to;
    build();
  }
  function setFill(ci, val, mode) {
    if (!previewKey) return;
    stash(columnFills, previewKey)[Number(ci)] = { val: val, mode: mode };
    build();
  }

  // Pre-export confirm, mirroring sites-standardize.js. Off-list values are
  // written exactly as they are, so this is the last place to catch them.
  function exportBlockers() {
    const out = { off: [], req: [], build: [] };
    Object.keys(built || {}).forEach(k => {
      const b = built[k];
      if (!b.rows.length) return;
      out.off.push(...b.mismatches);
      out.req.push(...b.required);
      out.build.push(...b.noOptions);
    });
    return out;
  }
  function hasBlockers(x) { return !!(x.off.length || x.req.length || x.build.length); }

  function confirmExport(x) {
    return new Promise(resolve => {
      const prior = $('dt-export-confirm');
      if (prior) prior.remove();
      const overlay = document.createElement('div');
      overlay.id = 'dt-export-confirm';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const sections = [];
      if (x.off.length) {
        let n = 0;
        x.off.forEach(m => { n += m.values.length; });
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + n +
          ' value(s) not in a template dropdown</div><ul class="ts-confirm-list">' +
          x.off.slice(0, 8).map(m => '<li><b>' + escHtml(m.column) + '</b>: ' +
            escHtml(m.values.slice(0, 4).join(', ')) + '</li>').join('') +
          '</ul><div class="ts-confirm-issue-body">These are written exactly as they are and PickTrace will reject those rows. Resolve them in <b>Smart Fixes</b> first.</div></div>');
      }
      if (x.req.length) {
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + x.req.length +
          ' required column(s) with empty cells</div><ul class="ts-confirm-list">' +
          x.req.slice(0, 8).map(r => '<li><b>' + escHtml(r.column) + '</b> — ' + r.empty + ' of ' + r.total + ' empty</li>').join('') +
          '</ul><div class="ts-confirm-issue-body">Fill them from the <b>Column Fill</b> panel.</div></div>');
      }
      if (x.build.length) {
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + x.build.length +
          ' value(s) not yet built in PickTrace</div><div class="ts-confirm-issue-body">Listed in the banner above. Create them, then re-download the template.</div></div>');
      }
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner ts-confirm-modal';
      inner.innerHTML = '<h3>Export anyway?</h3><div class="ts-confirm-issues">' + sections.join('') + '</div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" id="dt-conf-cancel">Cancel</button>' +
        '<button class="btn btn-primary" id="dt-conf-ok">Export anyway</button></div>';
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
      const close = v => { overlay.remove(); resolve(v); };
      inner.querySelector('#dt-conf-ok').addEventListener('click', () => close(true));
      inner.querySelector('#dt-conf-cancel').addEventListener('click', () => close(false));
      overlay.addEventListener('click', e => { if (e.target === overlay) close(false); });
    });
  }

  async function guardedDownload(fn) {
    const x = exportBlockers();
    if (hasBlockers(x)) { const ok = await confirmExport(x); if (!ok) return; }
    fn();
  }

  // ═══ File wiring ═══
  function slotLabel(key, kind, text) {
    const el = $('dt-' + key + '-' + kind + '-name');
    if (el) el.textContent = text;
  }

  function place(res, forcedKey) {
    const key = forcedKey || res.key;
    if (!key || !DOMAIN_BY_KEY[key]) return false;
    if (res.kind === 'template') {
      if (!DOMAIN_BY_KEY[key].needsTemplate) return false;
      tpls[key] = res.data;
      slotLabel(key, 'tpl', res.data.fileName + ' · ' + res.data.headers.length + ' cols · ' + res.data.dropdowns.size + ' dropdowns');
    } else {
      srcs[key] = res.data;
      slotLabel(key, 'src', res.data.fileName + ' · ' + res.data.rows.length + ' rows');
    }
    return true;
  }

  function handleSlot(key, kind, file) {
    readAnyFile(file).then(res => {
      if (res.kind !== kind) {
        alert('"' + file.name + '" looks like ' + (res.kind === 'template' ? 'a bulk template' : 'an export') +
          ', not ' + (kind === 'template' ? 'a template' : 'an export') + '. Load it in the matching slot.');
        return;
      }
      if (!place(res, key)) { alert('"' + file.name + '" could not be loaded into the ' + DOMAIN_BY_KEY[key].label + ' slot.'); return; }
      build();
    }).catch(err => alert('Failed to read ' + file.name + ': ' + (err && err.message ? err.message : err)));
  }

  function handleBulk(files) {
    const list = Array.from(files);
    if (!list.length) return;
    Promise.all(list.map(f => readAnyFile(f).then(r => ({ file: f, res: r }), () => ({ file: f, res: null }))))
      .then(results => {
        const unknown = [];
        results.forEach(({ file, res }) => {
          if (!res || !place(res)) unknown.push(file.name);
        });
        build();
        if (unknown.length) alert('Could not identify:\n' + unknown.join('\n') + '\n\nLoad these with their own slot button.');
      });
  }

  function reset() {
    tpls = {}; srcs = {}; built = null; previewKey = null;
    ageFixes = {}; columnFills = {}; smartFix = {}; fillFilter = ''; fillScope = 'all';
    { const f = $('dt-fill-filter'); if (f) f.value = ''; }
    { const s = $('dt-fill-scope'); if (s) s.value = 'all'; }
    { const b = $('dt-build-banner'); if (b) b.style.display = 'none'; }
    DOMAINS.forEach(d => {
      slotLabel(d.key, 'src', 'No export');
      if (d.needsTemplate) slotLabel(d.key, 'tpl', 'No template');
      ['src', 'tpl'].forEach(kind => { const f = $('dt-' + d.key + '-' + kind + '-file'); if (f) f.value = ''; });
      const btn = $('dt-dl-' + d.key); if (btn) btn.disabled = true;
    });
    const bulk = $('dt-bulk-file'); if (bulk) bulk.value = '';
    ['preview', 'smartfix', 'fill', 'fills', 'underage', 'derived', 'fuzzy', 'addrflag',
      'phonebad', 'mismatch', 'remap', 'nooptions', 'autofill', 'required', 'unmapped', 'dropped']
      .forEach(n => { const el = $('dt-section-' + n); if (el) el.style.display = 'none'; });
    $('dt-summary').style.display = 'none';
    $('dt-empty').style.display = '';
    $('dt-dl-all').disabled = true;
  }

  // ═══ Init ═══
  function init() {
    if (initialized) return;
    if (!$('dt-bulk-file')) return; // markup not present yet
    initialized = true;

    $('dt-bulk-file').addEventListener('change', e => { handleBulk(e.target.files); e.target.value = ''; });
    DOMAINS.forEach(d => {
      ['src', 'tpl'].forEach(kind => {
        const el = $('dt-' + d.key + '-' + kind + '-file');
        if (!el) return;
        el.addEventListener('change', e => {
          if (e.target.files[0]) handleSlot(d.key, kind === 'tpl' ? 'template' : 'export', e.target.files[0]);
          e.target.value = '';
        });
      });
      const btn = $('dt-dl-' + d.key);
      if (btn) btn.addEventListener('click', () => guardedDownload(() => download(d.key)));
    });
    $('dt-dl-all').addEventListener('click', () => guardedDownload(downloadAll));
    $('dt-reset').addEventListener('click', reset);
    { const m = $('dt-migrate'); if (m) m.addEventListener('change', e => { migrate = e.target.checked; build(); }); }
    { const t = $('dt-today'); if (t) t.addEventListener('change', e => { setToday = e.target.checked; build(); }); }

    // Under-age corrections: per-row date, or one date applied to everyone listed.
    $('dt-underage-table').addEventListener('change', e => {
      const inp = e.target.closest('.dt-age-row');
      if (!inp) return;
      const v = str(inp.value);
      if (v) ageFixes[inp.dataset.key] = v; else delete ageFixes[inp.dataset.key];
      build();
    });
    $('dt-age-apply').addEventListener('click', () => {
      const v = str($('dt-age-all').value);
      if (!v) { alert('Pick a date to apply.'); return; }
      collect('underAge').forEach(u => { ageFixes[u.key] = v; });
      build();
    });
    $('dt-age-clear').addEventListener('click', () => { ageFixes = {}; build(); });

    // Smart Fixes — one row, or every row whose picker has a value chosen.
    $('dt-smartfix-table').addEventListener('click', e => {
      const btn = e.target.closest('.dt-sf-apply');
      if (!btn) return;
      const sel = btn.closest('tr').querySelector('.dt-sf-pick');
      if (!sel || !sel.value) { alert('Pick a value to write instead.'); return; }
      applySmartFix(sel.dataset.key, sel.dataset.ci, sel.dataset.from, sel.value);
    });
    $('dt-smartfix-all').addEventListener('click', () => {
      const picks = $('dt-smartfix-table').querySelectorAll('.dt-sf-pick');
      let n = 0;
      picks.forEach(sel => {
        if (!sel.value) return;
        stash(smartFix, sel.dataset.key)[sel.dataset.ci + '||' + String(sel.dataset.from).toUpperCase().trim()] = sel.value;
        n++;
      });
      if (!n) { alert('No suggestions to apply — choose values first.'); return; }
      build();
    });

    // Column Fill
    $('dt-fill-table').addEventListener('click', e => {
      const btn = e.target.closest('button');
      if (!btn || !previewKey) return;
      const tr = btn.closest('tr');
      const ci = tr && tr.dataset.ci;
      if (ci == null) return;
      if (btn.classList.contains('dt-fill-clear')) { setFill(ci, '', 'all'); return; }
      const pick = tr.querySelector('.dt-fill-pick'), text = tr.querySelector('.dt-fill-text');
      const val = str(text && text.value) || str(pick && pick.value);
      if (!val) { alert('Pick a dropdown value or type an override first.'); return; }
      if (btn.classList.contains('dt-fill-all')) setFill(ci, val, 'all');
      else if (btn.classList.contains('dt-fill-blank')) setFill(ci, val, 'blank');
    });
    // Filtering only re-renders the panel — no rebuild, so fills stay put.
    $('dt-fill-filter').addEventListener('input', e => {
      fillFilter = str(e.target.value).toLowerCase();
      renderFill();
    });
    $('dt-fill-scope').addEventListener('change', e => { fillScope = e.target.value; renderFill(); });
    $('dt-fill-reset').addEventListener('click', () => {
      if (previewKey) delete columnFills[previewKey];
      build();
    });

    $('dt-build-banner').addEventListener('click', e => {
      const btn = e.target.closest('.dt-build-toggle');
      if (!btn) return;
      const card = btn.closest('.dt-build-col');
      const open = card.classList.toggle('dt-build-open');
      btn.textContent = open ? 'Show less' : 'Show ' + btn.dataset.more + ' more';
    });

    $('dt-preview-pills').addEventListener('click', e => {
      const btn = e.target.closest('.ts-domain-pill');
      if (!btn) return;
      previewKey = btn.dataset.key;
      renderPreview();
    });

    // Click-to-copy, matching the other Bulk Templates panes.
    $('dt-results').addEventListener('click', e => {
      const td = e.target.closest('.cmp-section .data-table td');
      if (!td || e.target.closest('input, button, select')) return;
      const text = (td.textContent || '').trim();
      if (text && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => td.classList.toggle('cmp-copied'), () => {});
      }
    });
  }

  window.dtInit = init;
  // Headless surface — lets the exports → filled-template pipeline be driven
  // without the DOM (see also window.cmpDebug in compare-logic.js).
  window.dtInternals = {
    DOMAINS: DOMAINS,
    readAnyFile: readAnyFile,
    detectExport: detectExport,
    detectTemplate: detectTemplate,
    parseDropdowns: parseDropdowns,
    toYMD: toYMD,
    todayYMD: todayYMD,
    ageFrom: ageFrom,
    MIN_AGE: MIN_AGE,
    suggestFor: suggestFor,
    setState: function (t, s) { tpls = t; srcs = s; },
    setAgeFixes: function (f) { ageFixes = f || {}; },
    setColumnFills: function (f) { columnFills = f || {}; },
    setSmartFix: function (f) { smartFix = f || {}; },
    setFillView: function (filter, scope) { fillFilter = filter || ''; fillScope = scope || 'all'; },
    renderPanel: function (key) {
      built = {};
      DOMAINS.forEach(d => { const m = mapDomain(d); if (m) built[d.key] = m; });
      previewKey = key;
      renderFill();
    },
    setFlags: function (f) {
      if (f && 'migrate' in f) migrate = !!f.migrate;
      if (f && 'setToday' in f) setToday = !!f.setToday;
    },
    mapDomain: mapDomain
  };
  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
