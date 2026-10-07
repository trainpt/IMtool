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
  // Implementation Data Templates split Crop & Variety in two.
  const CROP_KEYS = ['crop', 'croptype', 'cropname'];
  const VARIETY_KEYS = ['variety', 'varietal'];

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
  const UNIT_MARK = new Set(['spc','space','unit','apt','ste','suite','bldg','lot','trlr','rm','pmb','#']);
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
      // "1200 Example Blvd, Ste C Springfield" — that comma closed the street,
      // and the unit after it belongs on the street line, not in the city.
      const t = out.city.split(' ');
      if (t.length > 2 && UNIT_MARK.has(t[0].replace(/[.#]/g, '').toLowerCase())) {
        out.a1 += ', ' + t[0] + ' ' + t[1];
        out.city = t.slice(2).join(' ');
      }
      return out;
    }

    const toks = s.split(' ');
    const bare = t => t.replace(/[.#]/g, '').toLowerCase();
    let cut = -1;
    for (let i = 0; i < toks.length; i++) if (STREET_SUFFIX.has(bare(toks[i]))) cut = i;
    if (cut < 0) {
      for (let i = 0; i < toks.length - 1; i++) if (ES_STREET.has(bare(toks[i]))) { cut = i + 1; break; }
    }
    // No suffix ("1450 N. Maple Apt A Riverton"), but a unit still closes
    // the street line: the marker plus its value, whatever that value is.
    if (cut < 0) {
      for (let i = 1; i < toks.length - 2; i++) if (UNIT_MARK.has(bare(toks[i]))) { cut = i + 1; break; }
    }
    if (cut < 0) { out.a1 = s; out.flag = 'city not identified'; return out; }

    // Carry trailing units onto the street line. A unit marker always takes
    // the token after it as its value — "Apt C", not just "Apt" — so a
    // letter unit can't end up as the start of the city ("C Riverton").
    let end = cut;
    while (end + 1 < toks.length) {
      const nx = toks[end + 1];
      if (UNIT_MARK.has(bare(nx))) end += end + 2 < toks.length ? 2 : 1;
      else if (/^#?\d+[a-z]?$/i.test(nx)) end++;
      else break;
    }
    out.a1 = toks.slice(0, end + 1).join(' ');
    out.city = toks.slice(end + 1).join(' ');
    if (!out.city) out.flag = 'no city';
    return out;
  }

  // "210 Zephyr Lakeview" has no suffix or unit to end the street on, but the
  // same file usually spells that city out on another row ("540 Oak Street
  // Lakeview"). Learn every city the parser did identify, then peel the
  // longest known one off the end of each address it couldn't split.
  function settleCities(parsed) {
    const known = [];
    parsed.forEach(a => {
      const k = a.city.toUpperCase();
      if (!a.flag && k && known.indexOf(k) < 0) known.push(k);
    });
    known.sort((x, y) => y.length - x.length);
    parsed.forEach(a => {
      if (a.flag !== 'city not identified') return;
      const up = a.a1.toUpperCase();
      const hit = known.find(k => up.length > k.length + 1 && up.slice(-(k.length + 1)) === ' ' + k);
      if (!hit) return;
      a.city = a.a1.slice(-hit.length);
      a.a1 = a.a1.slice(0, -hit.length).trim();
      a.flag = '';
    });
    return parsed;
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
    if (['employees', 'sites', 'locations'].indexOf(dom.key) < 0) {
      return { headers: src.headers, rows: src.rows, derived: [], consumed: [], flags: [] };
    }
    const keys = src.headers.map(normLoose);
    const has = k => keys.indexOf(k) >= 0;
    const headers = src.headers.slice();
    const rows = src.rows.map(r => r.slice());
    const derived = [], consumed = [], flags = [];

    // Sites: one "Address" blob → the template's Address1 / City / State / Zip / Country.
    if (dom.key === 'sites' && !has('address1')) {
      const ai = keys.findIndex(k => ROSTER_ADDR_KEYS.indexOf(k) >= 0);
      if (ai >= 0) {
        const base = headers.length;
        headers.push('Address1', 'City', 'State', 'Zip', 'Country');
        const parsed = settleCities(rows.map(r => parseRosterAddress(r[ai])));
        rows.forEach((r, ri) => {
          const a = parsed[ri];
          r[base] = a.a1; r[base + 1] = a.city; r[base + 2] = a.state; r[base + 3] = a.zip; r[base + 4] = a.country;
          if (a.flag && str(r[ai])) flags.push({ domain: dom.label, reason: a.flag, detail: str(r[ai]), row: ri });
        });
        consumed.push(ai);
        derived.push({ domain: dom.label, from: src.headers[ai], into: 'Address1 · City · State · Zip · Country',
          how: 'address parsed', sample: sampleOf(rows, base, base + 4) });
      }
    }

    // Locations: separate Crop Type + Variety → "Crop-Variety", the dropdown's
    // shape. Same join Template Standardize uses; crop alone when no variety.
    if (dom.key === 'locations' && !has('cropvariety')) {
      const ci = keys.findIndex(k => CROP_KEYS.indexOf(k) >= 0);
      const vi = keys.findIndex(k => VARIETY_KEYS.indexOf(k) >= 0);
      if (ci >= 0) {
        const base = headers.length;
        headers.push('Crop & Variety');
        rows.forEach(r => {
          const c = str(r[ci]), v = vi >= 0 ? str(r[vi]) : '';
          r[base] = c && v ? c + '-' + v : c;
        });
        consumed.push(ci);
        if (vi >= 0) consumed.push(vi);
        derived.push({ domain: dom.label, from: src.headers[ci] + (vi >= 0 ? ' + ' + src.headers[vi] : ''),
          into: 'Crop & Variety', how: 'joined as Crop-Variety', sample: sampleOf(rows, base, base) });
      }
    }

    if (dom.key !== 'employees') return { headers: headers, rows: rows, derived: derived, consumed: consumed, flags: flags };

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
        const parsed = settleCities(rows.map(r => parseRosterAddress(r[ai])));
        rows.forEach((r, ri) => {
          const a = parsed[ri];
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
      aliases: {
        'Emergency Phone Number': ['Emergency Contact Phone Number'],
        'Alt ID': ['Payroll Software Employee Number', 'Employee Number']
      },
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
      // Implementation Data Template names. Exact matches win first, so
      // none of these touch a real PickTrace export.
      aliases: {
        'Name': ['Location Name'],
        'Alt ID': ['Location Alt ID'],
        'Acreage': ['Acres'],
        'Clone/Subvariety': ['Clone', 'Subvariety'],
        'Mulch Type': ['Multch Type'],
        'Bed Width, in.': ['Bed Width'],
        'Plant Spacing, in.': ['Plant Spacing'],
        'Row Spacing, in.': ['Row Spacing'],
        'Post Spacing, in.': ['Post Spacing'],
        'Custom Data 1': ['Loc Data 1'],
        'Custom Data 2': ['Loc Data 2'],
        'Custom Data 3': ['Loc Data 3']
      },
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
      aliases: { 'Name*': ['Site', 'Site Name'], 'Alt ID': ['Site Alt ID'] },
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
  let srcOverride = {};// domain → loose template header → source header name (Site* ← Location Group)
  let initialized = false;

  // ═══ Reading ═══
  // One entry point for every upload: works out whether the file is a
  // template or an export. A template resolves to { kind, key, data }; an
  // export to { kind, fileName, sheets: [{ name, key, hidden, data }] } —
  // one sheet for a PickTrace export, one per tab for a client's
  // Implementation Data Template, where the sheet picker chooses.
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
            kind: 'export', fileName: file.name,
            sheets: [{ name: '', key: detectExport(p.headers), hidden: false,
              data: { headers: p.headers, rows: p.rows, fileName: file.name, sheetName: '', headerRow: 1 } }]
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
        // A PickTrace export keeps its records on DATA ENTRY — take that and
        // nothing else. Otherwise every tab is a candidate, except the
        // reference sheet (the employees export's DROP-DOWN INPUTS has 249
        // rows of its own).
        const sheets = dataName
          ? [readSheet(wb, dataName, file.name, true)].filter(Boolean)
          : wb.SheetNames.filter(n => !/drop.?down/i.test(n)).map(n => readSheet(wb, n, file.name)).filter(Boolean);
        if (!sheets.length) { reject(new Error('no sheet with data rows')); return; }
        // A Crews tab is a lookup, not an upload: it travels with every sheet
        // of this workbook so whichever becomes Employees can translate the
        // payroll codes its Crew column carries.
        const crewTab = sheets.find(s => crewLookupFrom(s.data));
        const crews = crewTab ? crewLookupFrom(crewTab.data) : null;
        if (crews) {
          crewTab.lookup = 'crews';
          sheets.forEach(s => { s.data.crewLookup = crews; });
        }
        resolve({ kind: 'export', fileName: file.name, sheets: sheets });
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // A client's Implementation Data Template tab, top to bottom: a title, an
  // instructions paragraph, sometimes a banner row ("Basics", "Plant & Soil"),
  // the real header, a row of field notes, the records, then
  // "*add/delete rows as necessary" and PickTrace's contact block.
  const FOOTER_RE = /^\*?\s*add\s*\/\s*delete rows/i;
  const NOTE_RE = /^(required|optional)\b|^(please|enter|select) |^this (field|tab|is)\b|used (to collect|in reporting)/i;

  // PickTrace exports have the header on row 1, and the domain check finds it
  // there first, so their reading is unchanged. Failing a domain match, the
  // row that looks most like a header wins: many short cells, no sentences.
  function findHeaderRow(aoa) {
    const limit = Math.min(aoa.length, 20);
    for (let r = 0; r < limit; r++) if (detectExport((aoa[r] || []).map(str))) return r;
    let best = 0, bestScore = -Infinity;
    for (let r = 0; r < limit - 1; r++) {
      const cells = (aoa[r] || []).map(str).filter(Boolean);
      if (cells.length < 2) continue;
      let score = cells.length * 2;
      cells.forEach(s => { if (s.length < 40) score += 1; if (s.length > 80) score -= 4; });
      if (score > bestScore) { bestScore = score; best = r; }
    }
    return best;
  }

  // The field-notes row under an Implementation template header: mostly
  // paragraphs, or cells that read as instructions ("Required. …").
  function isNoteRow(row) {
    const cells = (row || []).map(str).filter(Boolean);
    if (!cells.length) return true;
    if (cells.filter(c => c.length > 80).length >= Math.ceil(cells.length / 2)) return true;
    return cells.some(c => NOTE_RE.test(c));
  }

  // One sheet → a candidate export, or null when it holds no records.
  // headerOnTop: a PickTrace DATA ENTRY export, whose header is always row 1.
  function readSheet(wb, name, fileName, headerOnTop) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) return null;
    const top = XLSX.utils.decode_range(ws['!ref']).s.r;
    const lines = [];   // non-empty rows, each with its 1-based sheet row number
    XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true }).forEach((r, i) => {
      if (r.some(c => str(c) !== '')) lines.push({ r: r, n: top + i + 1 });
    });
    if (lines.length < 2) return null;
    const h = headerOnTop ? 0 : findHeaderRow(lines.map(l => l.r));
    const headers = lines[h].r.map(str);
    let body = lines.slice(h + 1);
    // Only a header below row 1 marks a client template; an export's rows
    // pass through untouched.
    if (h > 0) {
      const end = body.findIndex(l => l.r.some(c => FOOTER_RE.test(str(c))));
      if (end >= 0) body = body.slice(0, end);
      while (body.length && isNoteRow(body[0].r)) body.shift();
    }
    if (!body.length) return null;
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    const wsMeta = meta[wb.SheetNames.indexOf(name)];
    return {
      name: name,
      key: detectExport(headers) || keyFromSheetName(name),
      hidden: !!(wsMeta && wsMeta.Hidden),
      data: {
        headers: headers,
        rows: body.map(l => l.r.map(c => (c instanceof Date ? c : str(c)))),
        fileName: fileName, sheetName: name, headerRow: lines[h].n
      }
    };
  }

  // An Implementation Data Template's Crews tab: Crew Leader · Crew Name ·
  // Crew Employer · Cost Code · Other Codes. Employees' Crew column often
  // carries the payroll code ("IRR01") rather than the crew's name
  // ("Irrigation"), and PickTrace crews are created by name — so this maps
  // every code, and every name, to the Crew Name. null when it isn't one.
  function crewLookupFrom(data) {
    const keys = data.headers.map(normLoose);
    const ni = keys.indexOf('crewname');
    if (ni < 0 || !(keys.indexOf('othercodes') >= 0 || keys.indexOf('crewleader') >= 0)) return null;
    const codeCols = keys.map((k, i) => (/^(othercodes?|crewcodes?|payrollcodes?|costcode)$/.test(k) ? i : -1)).filter(i => i >= 0);
    const map = new Map();   // UPPER code or name → Crew Name
    const names = [];
    data.rows.forEach(r => {
      const name = str(r[ni]);
      if (!name) return;
      if (names.indexOf(name) < 0) names.push(name);
      map.set(name.toUpperCase(), name);   // a name always means itself
    });
    data.rows.forEach(r => {
      const name = str(r[ni]);
      if (!name) return;
      codeCols.forEach(ci => str(r[ci]).split(/[,;\/]/).map(str).filter(Boolean).forEach(code => {
        if (!map.has(code.toUpperCase())) map.set(code.toUpperCase(), name);
      }));
    });
    return names.length ? { sheetName: data.sheetName, names: names, map: map } : null;
  }

  // Implementation templates name their tabs plainly; used when the headers
  // alone don't identify the domain.
  function keyFromSheetName(name) {
    const n = str(name).toLowerCase();
    if (/^employees?$/.test(n)) return 'employees';
    if (/^(locations?|blocks?)$/.test(n)) return 'locations';
    if (/^(sites?|ranch(es)?)$/.test(n)) return 'sites';
    if (/^jobs?$/.test(n)) return 'jobs';
    return null;
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
    // Implementation Data Template tabs, which name columns for the client.
    if (h.has('locationname') && (h.has('croptype') || h.has('acreage'))) return 'locations';
    if (h.has('sitename') && (h.has('sitetype') || h.has('address'))) return 'sites';
    if (h.has('jobname') && (h.has('jobcategory') || h.has('acceptspieces'))) return 'jobs';
    return null;
  }

  // A value whose dropdown entry is unambiguous even though the text differs.
  // Returns that entry, or '' to leave the value for Smart Fixes.
  //   Crop & Variety  "Cherries" → the crop's only entry, else its "-Other"
  //   yes / no        "Y", "Yes", "N", "No" → TRUE / FALSE
  //   short codes     "M" → "MALE", "es" → "es-MX" — begins exactly one option
  function autoMatch(header, value, list) {
    const v = str(value), up = v.toUpperCase();
    if (!v) return '';
    if (normLoose(header) === 'cropvariety') {
      if (v.indexOf('-') >= 0) return '';
      const mine = list.filter(o => splitCrop(o).crop.toUpperCase() === up);
      if (mine.length === 1) return mine[0];
      return mine.find(o => splitCrop(o).variety.toUpperCase() === 'OTHER') || '';
    }
    if (/^(Y|YES|N|NO)$/.test(up)) {
      const want = up.charAt(0) === 'Y' ? 'TRUE' : 'FALSE';
      const hit = list.find(o => o.toUpperCase() === want);
      if (hit) return hit;
    }
    if (v.length <= 2) {
      const hits = list.filter(o => o.toUpperCase().indexOf(up) === 0);
      if (hits.length === 1) return hits[0];
    }
    return '';
  }

  // ═══ Mapping ═══
  // pending: loose template header → values this run creates elsewhere (the
  // Sites upload's names, for Locations' Site*). They count as valid
  // dropdown values, so they aren't reported as needing building.
  function mapDomain(dom, pending) {
    const src = srcs[dom.key];
    if (!src) return null;
    const tpl = tpls[dom.key] || null;
    if (dom.needsTemplate && !tpl) return null;

    // Without a template the export's own columns are the schema, kept by
    // position so a repeated header ("GL Code?" twice) keeps its own values.
    const drop = new Set((dom.dropCols || []).map(normLoose));
    const keep = tpl ? null : src.headers.map((h, i) => i).filter(i => str(src.headers[i]) && !drop.has(normLoose(src.headers[i])));
    const headers = tpl ? tpl.headers : keep.map(i => src.headers[i]);
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
    const claimedSrc = new Set(der.consumed.concat(keep || []));
    const colSrc = keep ? keep.slice() : headers.map(h => {
      const k = normLoose(h);
      if (srcIdx.has(k)) { claimedSrc.add(srcIdx.get(k)); return srcIdx.get(k); }
      const al = dom.aliases[h] || dom.aliases[h.replace(/\*\s*$/, '').trim()];
      if (al) for (const a of al) {
        const ak = normLoose(a);
        if (srcIdx.has(ak)) { claimedSrc.add(srcIdx.get(ak)); return srcIdx.get(ak); }
      }
      return -1;
    });

    // A source column you pointed at a template column (Site* ← Location
    // Group) wins over the name match and stops feeding the column it matched
    // by name. The column it displaces stays claimed, so no fuzzy guess picks
    // it up — it is reported as unmapped instead.
    const displaced = new Set();
    const ov = srcOverride[dom.key] || {};
    Object.keys(ov).forEach(hk => {
      const ci = headers.findIndex(h => normLoose(h) === hk);
      const si = eff.headers.findIndex(h => normLoose(h) === normLoose(ov[hk]));
      if (ci < 0 || si < 0) return;
      colSrc.forEach((s, i) => { if (s === si && i !== ci) { colSrc[i] = -1; displaced.add(i); } });
      colSrc[ci] = si;
      claimedSrc.add(si);
    });

    const fuzzy = [];
    headers.forEach((h, ci) => {
      if (colSrc[ci] >= 0 || displaced.has(ci)) return;
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

    // Employees' Crew → the Crews tab's Crew Name ("IRR01" → "Irrigation").
    // Runs before Smart Fixes and the dropdown pass, so the name is what gets
    // matched against PickTrace's crews — or listed as a crew to create.
    const crewMapped = [];
    const crewCol = src.crewLookup ? headers.findIndex(h => normLoose(h) === 'crew') : -1;
    if (dom.key === 'employees' && crewCol >= 0) {
      const tally = new Map();
      rows.forEach(r => {
        const v = r[crewCol];
        const to = v ? src.crewLookup.map.get(v.toUpperCase()) : null;
        if (!to || to === v) return;
        const k = v + '\u0000' + to;
        tally.set(k, (tally.get(k) || 0) + 1);
        r[crewCol] = to;
      });
      tally.forEach((n, k) => {
        const p = k.split('\u0000');
        crewMapped.push({ domain: dom.label, column: headers[crewCol], from: [p[0]], to: p[1], rows: n,
          why: 'Crews tab: code → Crew Name' });
      });
    }

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
        overrides.push({ domain: dom.label, column: headers[sdCol], from: was.length ? was : ['(blank)'], to: today,
          why: 'Locations Start Date = today' });
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
    //
    // Values this run creates elsewhere (pending) join the list: a Site* the
    // Sites upload creates is valid here, not something to build. And a value
    // with exactly one plausible entry (Cherries → Cherries-Other) is matched
    // to it outright rather than queued as a Smart Fix — see autoMatch().
    const mismatches = [], remapped = [], noOptions = [], matched = [], fromPending = [];
    headers.forEach((h, ci) => {
      const own = dropdownFor(dropdowns, h);
      const extra = (pending && pending[normLoose(h)]) || [];
      if (!own && !extra.length) return;
      const list = (own || []).slice();
      const ownUp = new Set(list.map(v => v.toUpperCase()));
      const extraUp = new Set();
      extra.forEach(v => { const k = v.toUpperCase(); if (!ownUp.has(k) && !extraUp.has(k)) { extraUp.add(k); list.push(v); } });

      const caseMap = new Map();
      list.forEach(v => { const k = v.toUpperCase(); if (!caseMap.has(k)) caseMap.set(k, v); });
      const off = [], hits = new Map(), auto = new Map();
      rows.forEach(r => {
        const v = r[ci];
        if (!v) return;
        const k = v.toUpperCase();
        if (caseMap.has(k)) {
          r[ci] = caseMap.get(k);
          if (extraUp.has(k)) hits.set(r[ci], (hits.get(r[ci]) || 0) + 1);
          return;
        }
        const a = list.length ? autoMatch(h, v, list) : '';
        if (a) {
          const m = auto.get(v) || { to: a, n: 0 };
          m.n++;
          auto.set(v, m);
          r[ci] = a;
          return;
        }
        if (off.indexOf(v) < 0) off.push(v);
      });
      hits.forEach((n, v) => fromPending.push({ domain: dom.label, column: h, value: v, count: n }));
      auto.forEach((m, v) => matched.push({ domain: dom.label, column: h, from: [v], to: m.to, rows: m.n,
        why: 'only matching dropdown entry' }));
      // No list of its own: nothing to judge the leftovers against.
      if (!off.length || !own) return;
      if (!own.length) {
        // Enumerate each distinct value with its row count — this list is a
        // build checklist for PickTrace, so it must not be summarised away.
        off.forEach(v => {
          noOptions.push({ domain: dom.label, column: h, value: v, count: rows.filter(r => r[ci] === v).length });
        });
        return;
      }
      if (migrate && own.length === 1 && !extraUp.size) {
        rows.forEach(r => { if (r[ci] && r[ci] !== own[0]) r[ci] = own[0]; });
        remapped.push({ domain: dom.label, column: h, from: off, to: own[0], why: 'only option in this org' });
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
      matched: crewMapped.concat(matched), fromPending: fromPending,
      autofilled: autofilled, required: required, unmapped: unmapped, underAge: underAge,
      derived: der.derived, addrFlags: der.flags, fuzzy: fuzzy, filledCols: filledCols,
      phoneBad: phoneBad, dropdowns: dropdowns, pending: pending || null, key: dom.key, hasTemplate: !!tpl,
      // colSrc indexes these — the source plus any derived virtual columns.
      effHeaders: eff.headers, effRows: eff.rows };
  }

  function describeRow(headers, colSrc, sr) {
    for (let i = 0; i < headers.length; i++) {
      if (colSrc[i] >= 0 && str(sr[colSrc[i]])) return str(sr[colSrc[i]]);
    }
    return '(row)';
  }

  // ═══ Build ═══
  // Sites map first: every site the Sites upload creates is a valid Site*
  // for Locations, so those aren't listed as something to build in PickTrace.
  function buildAll() {
    const out = {};
    const sites = mapDomain(DOMAIN_BY_KEY.sites);
    const nameCi = sites ? sites.headers.findIndex(h => normLoose(h) === 'name') : -1;
    const created = nameCi >= 0 ? sites.rows.map(r => r[nameCi]).filter(Boolean) : [];
    const pending = created.length ? { site: created } : null;
    DOMAINS.forEach(d => {
      const m = d.key === 'sites' ? sites : mapDomain(d, d.key === 'locations' ? pending : null);
      if (m) out[d.key] = m;
    });
    return out;
  }

  function build() {
    built = buildAll();
    if (!previewKey || !built[previewKey]) previewKey = Object.keys(built)[0] || null;
    render();
    if (window.IMDebug) IMDebug.refresh('data-template');
  }

  // ═══ Render ═══
  function render() {
    const keys = Object.keys(built || {});
    $('dt-empty').style.display = keys.length ? 'none' : '';
    renderSummary(); renderPreview(); renderBuildBanner(); renderUnderAge();
    renderSmartFix(); renderFill(); renderSiteSrc();
    renderList('fills', collect('filledCols'));
    renderList('derived', collect('derived'));
    renderList('fuzzy', collect('fuzzy'));
    renderList('addrflag', collect('addrFlags'));
    renderList('phonebad', collect('phoneBad'));
    renderList('mismatch', collect('mismatches'));
    renderList('remap', collect('remapped').concat(collect('matched')));
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
    const fromCrews = collect('matched').filter(m => /^Crews tab/.test(m.why));
    if (fromCrews.length) html += '<span class="cmp-stat">' + plural(fromCrews.length, 'crew code') +
      ' → Crew Name (Crews tab)</span>';
    const auto = collect('matched').length - fromCrews.length;
    if (auto) html += '<span class="cmp-stat">' + plural(auto, 'value') + ' matched to the dropdown</span>';
    // Locations name sites the Sites upload creates — so that file goes first.
    const viaSites = collect('fromPending').length;
    if (viaSites) html += '<span class="cmp-stat">✓ ' + plural(viaSites, 'site') +
      ' created by the Sites upload — upload Sites before Locations</span>';
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

    // The pick list: the template's dropdown plus anything this run creates
    // (Site* lists the sites the Sites upload adds).
    const optionsFor = h => {
      const list = (dropdownFor(dd, h) || []).slice();
      const up = new Set(list.map(o => o.toUpperCase()));
      ((b.pending || {})[normLoose(h)] || []).forEach(o => { if (!up.has(o.toUpperCase())) { up.add(o.toUpperCase()); list.push(o); } });
      return list;
    };
    const cols = b.headers.map((h, ci) => {
      let empty = 0, sample = '';
      b.rows.forEach(r => { if (r[ci] === '') empty++; else if (!sample) sample = r[ci]; });
      return { h: h, ci: ci, empty: empty, sample: sample, req: /\*\s*$/.test(h), list: optionsFor(h) };
    });
    // A column in a field group none of whose columns hold data — Mailing
    // Address Country with no mailing address. Filling it alone makes the
    // partial address PickTrace rejects, so it isn't suggested.
    // Column → how many distinct values the dropdown rejects (off-list, or a
    // list the org hasn't built yet).
    const bad = new Map();
    (b.mismatches || []).forEach(m => bad.set(m.column, (bad.get(m.column) || 0) + m.values.length));
    (b.noOptions || []).forEach(n => bad.set(n.column, (bad.get(n.column) || 0) + 1));
    const groups = DOMAIN_BY_KEY[previewKey].groups || [];
    const deadGroup = c => {
      const g = groups.find(grp => grp.some(x => normLoose(x) === normLoose(c.h)));
      if (!g) return false;
      const set = new Set(g.map(normLoose));
      return cols.every(o => !set.has(normLoose(o.h)) || o.empty === b.rows.length);
    };
    // "Needs attention" (the default) is what's worth acting on: required
    // columns with blanks, empty columns that offer a dropdown to pick from,
    // and anything already filled here so it can be changed. Free-text
    // optional columns (Email, USCIS Number, …) sit under "All columns".
    const shown = cols.filter(c => {
      if (fillFilter && c.h.toLowerCase().indexOf(fillFilter) < 0) return false;
      if (fillScope === 'attn') return !!fills[c.ci] || bad.has(c.h) ||
        (c.empty > 0 && (c.req || (c.list.length > 0 && !deadGroup(c))));
      if (fillScope === 'req') return c.req;
      if (fillScope === 'empty') return c.empty > 0;
      if (fillScope === 'filled') return c.empty === 0;
      return true;
    });
    const need = shown.filter(c => c.req && c.empty > 0);
    // Fully filled, but with values the dropdown won't take (Location Group
    // when the org has none) — here so they can be bulk-set or cleared.
    const offl = shown.filter(c => !(c.req && c.empty > 0) && bad.has(c.h));
    const rest = shown.filter(c => !(c.req && c.empty > 0) && !bad.has(c.h));
    const reqEmpty = cols.filter(c => c.req && c.empty > 0).length;

    $('dt-fill-title').textContent = 'Column Fill — ' + DOMAIN_BY_KEY[previewKey].label +
      ' · ' + shown.length + ' of ' + cols.length + ' columns' +
      (reqEmpty ? ' · ' + reqEmpty + ' required still empty' : '');

    const rowFor = c => {
      const list = c.list;
      const cur = fills[c.ci] ? fills[c.ci].val : '';
      const opts = list.length
        ? '<option value="">— pick a value —</option>' +
          list.map(o => '<option' + (o === cur ? ' selected' : '') + '>' + escHtml(o) + '</option>').join('')
        : '<option value="">(no dropdown — type an override)</option>';
      const status = (bad.has(c.h) ? '<span class="dt-bad">⚠ ' + plural(bad.get(c.h), 'value') + ' not in dropdown</span>' +
          (c.empty ? ' · ' : '') : '') +
        (c.empty === 0
          ? (bad.has(c.h) ? '' : '<span class="dt-ok">✓ all ' + b.rows.length + '</span>')
          : (c.empty === b.rows.length
            ? '<span class="' + (c.req ? 'dt-bad' : 'cmp-warn') + '">' + (c.req ? '⚠ ' : '') + 'all empty</span>'
            : '<span class="cmp-warn">' + c.empty + ' empty</span>'));
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
      html += '<tbody><tr><td colspan="8" class="dt-fill-band">' +
        (fillScope === 'attn' && !fillFilter
          ? 'Nothing needs attention — switch to All columns to bulk-set anything else.'
          : 'No column matches this filter.') + '</td></tr></tbody>';
    }
    if (need.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band dt-fill-band-bad">⚠ Required and still empty (' +
        need.length + ')</td></tr>' + need.map(rowFor).join('') + '</tbody>';
    }
    if (offl.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band dt-fill-band-bad">⚠ Values the dropdown won\'t take — pick one for all rows, Clear the column, or create them in PickTrace (' +
        offl.length + ')</td></tr>' + offl.map(rowFor).join('') + '</tbody>';
    }
    if (rest.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band">' +
        (fillScope === 'attn' ? 'Optional — a dropdown to pick from, or filled here' : need.length ? 'Every other column' : 'Columns') +
        ' (' + rest.length + ')</td></tr>' + rest.map(rowFor).join('') + '</tbody>';
    }
    $('dt-fill-table').innerHTML = html;
  }
  let fillFilter = '';    // Column Fill search box, lowercased
  let fillScope = 'attn'; // attn | all | req | empty | filled

  // ─── Site* source (Locations slot) ───
  // Client templates sometimes carry the site in another column — e.g. a
  // "Site" column that holds block codes and the site names sit in "Location Group".
  // Lists each source column with a sample value so the right one is obvious.
  function renderSiteSrc() {
    const wrap = $('dt-locations-site-wrap'), sel = $('dt-locations-site-src');
    if (!wrap || !sel) return;
    const src = srcs.locations;
    if (!src) { wrap.style.display = 'none'; sel.innerHTML = ''; return; }
    const cur = normLoose((srcOverride.locations || {}).site || '');
    const seen = new Set();
    const sample = i => { for (const r of src.rows) { const v = str(r[i]); if (v) return v; } return ''; };
    sel.innerHTML = '<option value="">Auto — matched by column name</option>' +
      src.headers.map((h, i) => ({ h: str(h), i: i })).filter(o => {
        const k = normLoose(o.h);
        if (!k || seen.has(k)) return false;
        seen.add(k);
        return true;
      }).map(o => {
        const s = sample(o.i);
        return '<option value="' + escHtml(o.h) + '"' + (normLoose(o.h) === cur ? ' selected' : '') + '>' +
          escHtml(o.h + (s ? ' — e.g. ' + s.slice(0, 24) : ' (empty)')) + '</option>';
      }).join('');
    wrap.style.display = '';
  }

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
    remap: { title: 'Rewritten Values', cols: ['Domain', 'Column', 'Was', 'Written as', 'Why'],
      row: r => [r.domain, r.column, r.from.join(' · '), r.to,
        (r.why || '') + (r.rows ? ' · ' + plural(r.rows, 'row') : '')] },
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
    XLSX.writeFile(wb, outName, { cellStyles: true, bookSST: true });
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
      // A source-column choice only survives an export that still has that column.
      const ov = srcOverride[key] || {};
      const have = new Set(res.data.headers.map(normLoose));
      Object.keys(ov).forEach(k => { if (!have.has(normLoose(ov[k]))) delete ov[k]; });
      slotLabel(key, 'src', res.data.fileName + (res.sheetLabel ? ' › ' + res.sheetLabel : '') +
        ' · ' + res.data.rows.length + ' rows');
    }
    return true;
  }

  // A sheet from readAnyFile → place()'s shape. The sheet is named in the
  // slot label only when the workbook offered a choice.
  function asExport(sheet, many) {
    return { kind: 'export', key: sheet.key, data: sheet.data, sheetLabel: many ? sheet.name : '' };
  }

  // ─── Sheet picker ───
  // A client workbook (an Implementation Data Template) carries one tab per
  // domain. From the bulk upload you tick any number of sheets and choose
  // the template each one fills; from a domain's own Export button you pick
  // one. Resolves to [{ sheet, key }], or rejects when cancelled.
  function pickSheets(res, slotKey) {
    return new Promise((resolve, reject) => {
      const prior = $('dt-sheet-picker');
      if (prior) prior.remove();
      const sheets = res.sheets;
      const overlay = document.createElement('div');
      overlay.id = 'dt-sheet-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner dt-sheet-modal';
      const meta = s => s.data.rows.length + ' row' + (s.data.rows.length === 1 ? '' : 's') + ' · ' +
        s.data.headers.filter(Boolean).length + ' cols · header row ' + s.data.headerRow + (s.hidden ? ' · hidden' : '') +
        (s.lookup === 'crews' ? ' · read automatically: turns Employees\' crew codes into Crew Names' : '');
      const tip = s => 'Columns: ' + s.data.headers.filter(Boolean).slice(0, 8).join(', ');
      const close = () => overlay.remove();
      const cancel = () => { close(); reject(new Error('cancelled')); };

      if (slotKey) {
        inner.innerHTML = '<h3>Select sheet</h3>' +
          '<p><code>' + escHtml(res.fileName) + '</code> has ' + sheets.length + ' sheets with data. Pick the one to load as the <b>' +
          escHtml(DOMAIN_BY_KEY[slotKey].label) + '</b> export:</p>' +
          '<div class="cmp-org-list">' + sheets.map((s, i) =>
            '<div class="cmp-org-list-item" data-i="' + i + '" title="' + escHtml(tip(s)) + '">' +
            '<span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
            '<span class="cmp-org-counts">' + escHtml(meta(s) + (s.key === slotKey ? ' · detected' : '')) + '</span></div>').join('') +
          '</div>' +
          '<div class="cmp-export-modal-actions"><button class="btn btn-ghost" data-act="cancel">Cancel</button></div>';
        inner.addEventListener('click', e => {
          const item = e.target.closest('.cmp-org-list-item');
          if (item) { close(); resolve([{ sheet: sheets[+item.dataset.i], key: slotKey }]); }
        });
      } else {
        // Pre-tick what was detected — the first sheet per domain, never a
        // hidden one (those are usually lookup lists).
        const taken = new Set();
        const on = sheets.map(s => {
          if (!s.key || s.hidden || taken.has(s.key)) return false;
          taken.add(s.key);
          return true;
        });
        const opts = k => '<option value="">— skip —</option>' + DOMAINS.map(d =>
          '<option value="' + d.key + '"' + (d.key === k ? ' selected' : '') + '>' + escHtml(d.label) + '</option>').join('');
        inner.innerHTML = '<h3>Select sheets</h3>' +
          '<p><code>' + escHtml(res.fileName) + '</code> has ' + sheets.length + ' sheets with data. Tick each one to load and choose the template it fills &mdash; detected sheets are already ticked.</p>' +
          '<div class="cmp-org-list">' + sheets.map((s, i) =>
            '<div class="cmp-org-list-item dt-sheet-row" title="' + escHtml(tip(s)) + '">' +
            '<label class="dt-sheet-pick"><input type="checkbox" class="dt-sheet-on" data-i="' + i + '"' + (on[i] ? ' checked' : '') + '>' +
            '<span class="dt-sheet-meta"><span class="cmp-org-name">' + escHtml(s.name) + '</span>' +
            '<span class="cmp-org-counts">' + escHtml(meta(s)) + '</span></span></label>' +
            '<select class="input-field dt-sheet-key" data-i="' + i + '">' + opts(s.key) + '</select></div>').join('') +
          '</div>' +
          '<div class="dt-sheet-err" style="display:none;"></div>' +
          '<div class="cmp-export-modal-actions"><button class="btn btn-primary" data-act="load">Load selected</button>' +
          '<button class="btn btn-ghost" data-act="cancel">Cancel</button></div>';
        const boxes = inner.querySelectorAll('.dt-sheet-on');
        const keys = inner.querySelectorAll('.dt-sheet-key');
        const err = inner.querySelector('.dt-sheet-err');
        // Choosing a template ticks the sheet; "skip" unticks it.
        inner.addEventListener('change', e => {
          if (e.target.classList.contains('dt-sheet-key')) boxes[+e.target.dataset.i].checked = !!e.target.value;
          err.style.display = 'none';
        });
        inner.querySelector('[data-act="load"]').addEventListener('click', () => {
          const picks = [], noKey = [], byKey = {};
          boxes.forEach((b, i) => {
            if (!b.checked) return;
            const k = keys[i].value;
            if (!k) { noKey.push(sheets[i].name); return; }
            (byKey[k] = byKey[k] || []).push(sheets[i].name);
            picks.push({ sheet: sheets[i], key: k });
          });
          // Each domain fills one template, so it takes one sheet.
          const dup = Object.keys(byKey).filter(k => byKey[k].length > 1);
          const msg = !picks.length && !noKey.length ? 'Tick at least one sheet, or Cancel.'
            : noKey.length ? 'Choose a template for: ' + noKey.join(', ') + '.'
            : dup.length ? dup.map(k => DOMAIN_BY_KEY[k].label + ' is chosen for more than one sheet (' + byKey[k].join(', ') + ') — keep one.').join(' ')
            : '';
          if (msg) { err.textContent = msg; err.style.display = ''; return; }
          close();
          resolve(picks);
        });
      }
      inner.querySelector('[data-act="cancel"]').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
    });
  }

  function handleSlot(key, kind, file) {
    readAnyFile(file).then(res => {
      if (res.kind !== kind) {
        alert('"' + file.name + '" looks like ' + (res.kind === 'template' ? 'a bulk template' : 'an export') +
          ', not ' + (kind === 'template' ? 'a template' : 'an export') + '. Load it in the matching slot.');
        return;
      }
      if (kind === 'template') {
        if (!place(res, key)) { alert('"' + file.name + '" could not be loaded into the ' + DOMAIN_BY_KEY[key].label + ' slot.'); return; }
        build();
        return;
      }
      const many = res.sheets.length > 1;
      const pick = many ? pickSheets(res, key) : Promise.resolve([{ sheet: res.sheets[0], key: key }]);
      return pick.then(picks => {
        place(asExport(picks[0].sheet, many), key);
        build();
      }, () => {}); // picker cancelled
    }).catch(err => alert('Failed to read ' + file.name + ': ' + (err && err.message ? err.message : err)));
  }

  async function handleBulk(files) {
    const list = Array.from(files);
    if (!list.length) return;
    const results = await Promise.all(list.map(f => readAnyFile(f).then(r => ({ file: f, res: r }), () => ({ file: f, res: null }))));
    const unknown = [], multi = [];
    results.forEach(({ file, res }) => {
      if (!res) { unknown.push(file.name); return; }
      if (res.kind === 'template') { if (!place(res)) unknown.push(file.name); return; }
      if (res.sheets.length > 1) { multi.push(res); return; }
      if (!place(asExport(res.sheets[0], false))) unknown.push(file.name);
    });
    // One picker per multi-sheet workbook, in turn. A cancelled picker loads
    // nothing from that file.
    for (const res of multi) {
      let picks;
      try { picks = await pickSheets(res, null); } catch (e) { continue; }
      picks.forEach(p => place(asExport(p.sheet, true), p.key));
    }
    build();
    if (unknown.length) alert('Could not identify:\n' + unknown.join('\n') + '\n\nLoad these with their own slot button.');
  }

  // ═══ Debug dump ═══
  // This tool builds several domains from several exports in one pass, so the
  // dump is keyed by domain. mapDomain() already records every decision it
  // makes; this mostly repackages that into the shared shape.
  function collectDebug() {
    if (!built || !Object.keys(built).length) return null;
    const keys = Object.keys(built);
    const inputs = [];
    DOMAINS.forEach(d => {
      const s = srcs[d.key], t = tpls[d.key];
      inputs.push(IMDebug.file(d.label + ' — source export', s || null));
      if (d.needsTemplate || t) {
        inputs.push(IMDebug.file(d.label + ' — bulk template', t || null, {
          required: !!d.needsTemplate,
          dropdowns: t && t.dropdowns ? [...t.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : []
        }));
      }
    });

    const asked = [], domains = {};
    keys.forEach(k => {
      const m = built[k];
      const dom = DOMAIN_BY_KEY[k];
      const src = srcs[k];
      const label = dom ? dom.label : k;

      const cols = m.headers.map((h, i) => {
        const si = m.colSrc[i] != null ? m.colSrc[i] : -1;
        const fz = (m.fuzzy || []).find(f => f.column === h);
        return { index: i, header: h, required: /\*\s*$/.test(h), srcIndex: si,
          match: fz ? ('fuzzy guess from "' + fz.from + '" (' + fz.score + '% similar)')
                    : (si >= 0 ? 'exact or alias' : 'unmapped'),
          sample: si >= 0 && m.rows[0] ? m.rows[0][i] : null };
      });

      if ((m.required || []).length) {
        asked.push(IMDebug.ask(k + '/empty-required', 'required', label + ': required columns still blank', {
          count: m.required.length, blocksExport: true,
          items: m.required.map(r => ({ column: r.column, emptyRows: r.empty, totalRows: r.total })) }));
      }
      if ((m.mismatches || []).length) {
        asked.push(IMDebug.ask(k + '/off-list', 'dropdown', label + ': values not in the template dropdown', {
          count: m.mismatches.length, blocksExport: true,
          detail: 'Pick a replacement in Smart Fixes, or create these in PickTrace.',
          items: m.mismatches.map(x => ({ column: x.column, values: x.values, rowCounts: x.counts, allowed: x.allowed })) }));
      }
      if ((m.noOptions || []).length) {
        asked.push(IMDebug.ask(k + '/no-options', 'dropdown', label + ': columns whose dropdown is empty in your org', {
          count: m.noOptions.length, blocksExport: true,
          detail: 'There is nothing to migrate these to — they must be created in PickTrace first.',
          items: m.noOptions.map(x => ({ column: x.column, value: x.value, rows: x.count })) }));
      }
      if ((m.underAge || []).length) {
        asked.push(IMDebug.ask(k + '/under-age', 'data', label + ': people below the minimum working age', {
          count: m.underAge.length, blocksExport: false,
          detail: 'Almost always a typo\'d birth year. Never altered or dropped automatically.',
          items: m.underAge.map(u => ({ name: u.name, dateOfBirth: u.dob, age: u.age,
            correctedTo: ageFixes[u.key] || null })) }));
      }
      if ((m.fuzzy || []).length) {
        asked.push(IMDebug.ask(k + '/fuzzy-columns', 'data', label + ': columns matched by a fuzzy guess', {
          count: m.fuzzy.length, blocksExport: false,
          detail: 'These were not exact or alias matches — check that each guess is right.',
          items: m.fuzzy.map(f => ({ templateColumn: f.column, sourceColumn: f.from, similarity: f.score + '%' })) }));
      }
      if ((m.phoneBad || []).length) {
        asked.push(IMDebug.ask(k + '/bad-phones', 'data', label + ': phone numbers that would not normalize', {
          count: m.phoneBad.length, blocksExport: false,
          items: m.phoneBad.map(p => ({ column: p.column, value: p.value })) }));
      }
      if ((m.addrFlags || []).length) {
        asked.push(IMDebug.ask(k + '/address-flags', 'data', label + ': addresses the parser was unsure about', {
          count: m.addrFlags.length, blocksExport: false, items: m.addrFlags }));
      }
      if ((m.unmapped || []).length) {
        asked.push(IMDebug.ask(k + '/unused-source-columns', 'data', label + ': source columns carrying data that nothing claims', {
          count: m.unmapped.length, blocksExport: false,
          detail: 'Nothing is lost silently — these simply have no destination column.',
          items: m.unmapped.map(u => ({ column: u.column, sampleValues: u.sample })) }));
      }

      const out = IMDebug.output(m.headers, m.rows, {
        note: 'exportedRowIndexes[n] is the index into this domain\'s SOURCE rows behind output row n.' });
      out.exportedRowIndexes = m.rowSrc;

      domains[k] = {
        label: label,
        hasTemplate: !!m.hasTemplate,
        // Source + derived virtual columns, so a derived source shows by name
        // ("Crop & Variety") rather than as a bare index.
        mapping: IMDebug.mapping(cols, m.effHeaders || (src ? src.headers : []), { srcRows: m.effRows || (src ? src.rows : null) }),
        derived: [
          { kind: 'virtual columns', note: 'Columns synthesized from a messy source (name split, address parsed).',
            items: m.derived || [] },
          { kind: 'auto-filled single-option columns',
            note: 'An empty column whose dropdown offers exactly one choice is filled with it — but only on rows whose field group already has data.',
            items: m.autofilled || [] },
          { kind: 'migrated to a single-option list', enabled: migrate,
            note: 'Off-list values rewritten to the only allowed value, plus any blanket overrides (e.g. today\'s Start Date).',
            items: m.remapped || [] },
          { kind: 'matched to a dropdown entry',
            note: 'Off-list values with exactly one plausible entry (Cherries → Cherries-Other, M → MALE, Yes → TRUE), rewritten without a Smart Fix.',
            items: m.matched || [] },
          { kind: 'created by another upload in this run',
            note: 'Values not in the template dropdown that this run creates (Locations Site* ← the Sites upload) — valid once that file is uploaded first.',
            items: m.fromPending || [] },
          { kind: 'date + SSN + phone normalization',
            note: 'Date columns run through toYMD(); SSN and phone columns are reformatted in place.' }
        ],
        fills: (m.filledCols || []).map(f => ({ column: f.column, value: f.value, scope: f.mode, rows: f.rows, kind: 'column fill' }))
          .concat(Object.keys(smartFix[k] || {}).map(sk => {
            const sep = sk.indexOf('||');
            return { column: m.headers[+sk.slice(0, sep)], from: sk.slice(sep + 2), value: smartFix[k][sk], kind: 'smart fix' };
          })),
        edits: { ageFixes: Object.keys(ageFixes).filter(a => a.indexOf(k + '|') === 0)
          .map(a => ({ sourceRowIndex: +a.slice(k.length + 1), correctedDateOfBirth: ageFixes[a] })) },
        dropped: {
          count: (m.dropped || []).length,
          items: (m.dropped || []).slice(0, 500).map(d => ({ reason: d.reason, row: d.detail })),
          dedupedOn: dom && dom.dedupe ? dom.dedupe : null
        },
        output: out,
        provenance: IMDebug.deriveProvenance({
          headers: m.headers, rows: m.rows,
          srcRows: m.rowSrc.map(si => (m.effRows ? m.effRows[si] : src && src.rows ? src.rows[si] : null)),
          colToSrc: m.colSrc.reduce((o, si, i) => { o[i] = si; return o; }, {}),
          fills: (columnFills[k] || {}), smartFixes: (smartFix[k] || {})
        })
      };
    });

    return {
      settings: {
        migrateToSingleOptionLists: migrate,
        stampLocationsStartDateWithToday: setToday,
        sourceColumnOverrides: srcOverride,
        crewLookup: srcs.employees && srcs.employees.crewLookup
          ? { sheet: srcs.employees.crewLookup.sheetName, crewNames: srcs.employees.crewLookup.names,
              codes: [...srcs.employees.crewLookup.map.entries()].map(([k, v]) => ({ code: k, crewName: v })) }
          : null,
        domainsBuilt: keys,
        previewingDomain: previewKey
      },
      inputs: inputs,
      domains: domains,
      asked: asked,
      output: { note: 'This tool writes one file per domain — see domains.<key>.output for each grid.',
        rowCount: keys.reduce((n, k) => n + built[k].rows.length, 0),
        perDomain: keys.map(k => ({ domain: k, rows: built[k].rows.length, columns: built[k].headers.length })) },
      provenance: { note: 'Per-domain — see domains.<key>.provenance.', legend: IMDebug.PROV_LEGEND }
    };
  }

  function reset() {
    tpls = {}; srcs = {}; built = null; previewKey = null;
    ageFixes = {}; columnFills = {}; smartFix = {}; srcOverride = {}; fillFilter = ''; fillScope = 'attn';
    { const f = $('dt-fill-filter'); if (f) f.value = ''; }
    { const s = $('dt-fill-scope'); if (s) s.value = 'attn'; }
    renderSiteSrc();
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
    if (window.IMDebug) IMDebug.refresh('data-template');
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
    if (window.IMDebug) {
      IMDebug.register('data-template', {
        label: 'Data Template Bundle',
        ready: () => !!(built && Object.keys(built).length),
        collect: collectDebug
      });
      IMDebug.wire('dt-debug', 'data-template');
    }
    { const m = $('dt-migrate'); if (m) m.addEventListener('change', e => { migrate = e.target.checked; build(); }); }
    { const t = $('dt-today'); if (t) t.addEventListener('change', e => { setToday = e.target.checked; build(); }); }
    { const s = $('dt-locations-site-src'); if (s) s.addEventListener('change', e => {
      const ov = stash(srcOverride, 'locations');
      if (e.target.value) ov.site = e.target.value; else delete ov.site;
      build();
    }); }

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
      renderFill(); // Column Fill follows the previewed domain
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
    readSheet: readSheet,
    findHeaderRow: findHeaderRow,
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
    setFillView: function (filter, scope) { fillFilter = filter || ''; fillScope = scope || 'attn'; },
    setSrcOverride: function (o) { srcOverride = o || {}; },
    buildAll: buildAll,
    autoMatch: autoMatch,
    renderPanel: function (key) {
      built = buildAll();
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
