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
  // Pack Styles tabs write the weight with its unit ("1 lb", "12oz").
  const WEIGHT_KEYS = ['weight', 'packweight', 'netweight'];
  const WEIGHT_UNITS = { lb: 'LB', lbs: 'LB', pound: 'LB', pounds: 'LB', '#': 'LB',
    oz: 'OZ', ounce: 'OZ', ounces: 'OZ', g: 'G', gr: 'G', gram: 'G', grams: 'G' };

  // "1 lb" → { weight: "1", unit: "LB" }; null when the cell isn't a number
  // followed by one of the template's weight units (LB / OZ / G).
  function splitWeight(v) {
    const m = str(v).match(/^(\d+(?:\.\d+)?|\.\d+)\s*(lbs?|pounds?|#|oz|ounces?|grams?|gr|g)\.?$/i);
    return m ? { weight: String(parseFloat(m[1])), unit: WEIGHT_UNITS[m[2].toLowerCase()] } : null;
  }

  function titleCaseName(v) {
    return str(v).replace(/\s+/g, ' ').toLowerCase()
      .replace(/(^|[\s'-])([a-z])/g, (m, sep, ch) => sep + ch.toUpperCase());
  }

  // Western split, as chosen: first token, last token, everything between is
  // the middle. "JANE Q. DOE" → Jane / Q. / Doe.
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

  // "100 EXAMPLE AVENUE SPRINGFIELD CA, 90000" → street / city / state / zip.
  // Sibling of parseAddress() in sites-standardize.js, extended for the shapes
  // rosters actually contain: PO boxes, SPC/Unit numbers, glued "30ANYTOWN".
  function parseRosterAddress(raw) {
    const out = { a1: '', city: '', state: '', zip: '', country: '', flag: '' };
    let s = str(raw).replace(/\s+/g, ' ');
    if (!s) return out;
    // A trailing country ("…, CA 90000, USA") comes off first, or the ZIP and
    // state aren't at the end and the country passes for the city.
    let m = s.match(/[,\s]\s*(U\.?\s?S\.?(?:\s?A\.?)?|United States(?: of America)?)\s*$/i);
    if (m && m.index > 0) s = s.slice(0, m.index).trim().replace(/,\s*$/, '');

    m = s.match(/[,\s]\s*(\d{5}(?:-\d{4})?)\s*$/);
    if (m) { out.zip = m[1]; s = s.slice(0, m.index).trim(); }
    m = s.match(/(,?)\s*([A-Za-z]{2})\s*$/);
    if (m && (m[1] === ',' || US_STATES.has(m[2].toUpperCase()))) {
      out.state = m[2].toUpperCase();
      s = s.slice(0, s.length - m[0].length).trim();
    }
    out.country = out.state ? 'US' : '';
    s = s.replace(/,\s*$/, '').trim();
    s = s.replace(/(\d)([A-Za-z]{3,})/g, '$1 $2'); // "Unit 30ANYTOWN"

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

  // A numeric column keeps the number and drops an inline unit — "7.25 ac" →
  // 7.25, "2,750 Vines" → 2750 — as Template Standardize does, and an Acreage
  // of 0 is left blank there too. Anything it can't read as one number stays
  // as typed so it shows up rather than vanishing: a note typed into Acres,
  // "4x9", or feet written into an inches column.
  function cleanNumber(v, header) {
    const t = str(v);
    if (!t) return t;
    const nums = t.replace(/,/g, '').match(/-?\d+(?:\.\d+)?/g);
    if (!nums || nums.length !== 1) return t;
    if (/,\s*in\.?$/i.test(header) && /\b(ft|feet|foot)\b|'/i.test(t)) return t;
    const n = parseFloat(nums[0]);
    if (n === 0 && normLoose(header) === 'acreage') return '';
    return String(n);
  }

  function normSsn(v) {
    const s = str(v);
    const d = s.replace(/\D/g, '');
    return d.length === 9 ? d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5) : s;
  }

  // Append virtual columns for whatever the roster packs into one field.
  // Returns the effective source plus a report of what was derived.
  function deriveColumns(dom, src) {
    if (['employees', 'sites', 'locations', 'packstyles'].indexOf(dom.key) < 0) {
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
          const c = tidy(r[ci]), v = vi >= 0 ? tidy(r[vi]) : '';
          r[base] = c && v ? c + '-' + v : c;
        });
        consumed.push(ci);
        if (vi >= 0) consumed.push(vi);
        derived.push({ domain: dom.label, from: src.headers[ci] + (vi >= 0 ? ' + ' + src.headers[vi] : ''),
          into: 'Crop & Variety', how: 'joined as Crop-Variety', sample: sampleOf(rows, base, base) });
      }
    }

    // Pack Styles: "1 lb" in Weight → Weight 1 + Unit LB. The template keeps the
    // number and its unit apart. Rewritten in place, since the source's own
    // Weight column is what the resolver matches; a Unit the row already
    // states wins, and the unit lands in a virtual Unit column when the source
    // has none.
    if (dom.key === 'packstyles') {
      const wi = keys.findIndex(k => WEIGHT_KEYS.indexOf(k) >= 0);
      if (wi >= 0) {
        let ui = keys.indexOf('unit');
        const virtual = ui < 0;
        let n = 0;
        rows.forEach(r => {
          const w = splitWeight(r[wi]);
          if (!w) return;
          if (virtual) { if (ui < 0) { ui = headers.length; headers.push('Unit'); } }
          r[wi] = w.weight;
          if (!str(r[ui])) r[ui] = w.unit;
          n++;
        });
        if (n) derived.push({ domain: dom.label, from: src.headers[wi], into: src.headers[wi] + ' · Unit',
          how: 'weight split from its unit (' + plural(n, 'row') + ')', sample: sampleOf(rows, wi, wi) +
            (ui >= 0 ? ' · ' + sampleOf(rows, ui, ui) : '') });
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

  // ─── Crop & Variety spelling ───
  // A misspelled variety ("Carbernet Franc", "Rousanne", "Petite Verdot") isn't
  // in the dropdown, so it lands on the build list as something to create —
  // a second Cabernet Franc in PickTrace. spellCropVariety() fixes those
  // before the dropdown check. For each crop, then each variety, the right
  // spelling is, in order:
  //   1. the template's own entry  (the org already has it)
  //   2. a known wine grape name   (grape crops only — WINE_GRAPES)
  //   3. the spelling most of your rows use (at least twice as many rows)
  // "Close" means a small edit distance on letters and digits only, scaled to
  // the name's length (none under 5 letters), with the digits identical —
  // "Sample 1234-1" never becomes "Sample 1234-2". Two equally close
  // candidates means no fix. Every fix is listed with an Undo (spellKeep).
  const WINE_GRAPES = ['Agiorgitiko', 'Aglianico', 'Albarino', 'Alicante Bouschet', 'Aligote', 'Alvarinho', 'Arneis',
    'Assyrtiko', 'Barbera', 'Blaufrankisch', 'Bourboulenc', 'Cabernet Franc', 'Cabernet Pfeffer', 'Cabernet Sauvignon',
    'Carignan', 'Carignane', 'Carmenere', 'Chambourcin', 'Charbono', 'Chardonnay', 'Chenin Blanc', 'Cinsault', 'Cinsaut',
    'Clairette Blanche', 'Colombard', 'Counoise', 'Dolcetto', 'Durif', 'Falanghina', 'Fiano', 'French Colombard', 'Frappato',
    'Freisa', 'Furmint', 'Gamay', 'Gewurztraminer', 'Godello', 'Gouveio', 'Graciano', 'Grenache', 'Grenache Blanc',
    'Grenache Gris', 'Grenache Noir', 'Grillo', 'Gros Manseng', 'Gruner Veltliner', 'Lagrein', 'Loureiro', 'Malbec',
    'Malvasia', 'Malvasia Bianca', 'Marsanne', 'Mencia', 'Merlot', 'Mission', 'Montepulciano', 'Mourvedre', 'Muller Thurgau',
    'Muscardin', 'Muscat', 'Muscat Blanc', 'Muscat Canelli', 'Nebbiolo', 'Negroamaro', 'Nerello Mascalese', 'Nero d\'Avola',
    'Nielluccio', 'Norton', 'Orange Muscat', 'Petit Manseng', 'Petit Verdot', 'Petite Sirah', 'Picardan', 'Picpoul',
    'Picpoul Blanc', 'Pinot Blanc', 'Pinot Grigio', 'Pinot Gris', 'Pinot Meunier', 'Pinot Noir', 'Primitivo', 'Refosco',
    'Ribolla Gialla', 'Riesling', 'Rkatsiteli', 'Roussanne', 'Ruby Cabernet', 'Sagrantino', 'Sangiovese', 'Sauvignon Blanc',
    'Sauvignon Gris', 'Schioppettino', 'Semillon', 'Seyval Blanc', 'Shiraz', 'Souzao', 'Syrah', 'Sylvaner', 'Tannat',
    'Tempranillo', 'Teroldego', 'Terret Noir', 'Tinta Amarela', 'Tinta Cao', 'Tinta Roriz', 'Tinto Cao', 'Touriga Franca',
    'Touriga Nacional', 'Trousseau', 'Trousseau Gris', 'Ugni Blanc', 'Vaccarese', 'Valdiguie', 'Verdejo', 'Verdelho',
    'Vermentino', 'Vidal Blanc', 'Viognier', 'Zinfandel', 'Zweigelt'];
  const GRAPE_CROP = /grape|vine|wine|vino|uva/i;
  // Accents, case, spaces and punctuation don't count: "Albariño" = "albarino".
  const spellKey = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '');
  const WINE_KEYS = new Set(WINE_GRAPES.map(spellKey));
  const tidy = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  // Optimal string alignment distance (a swap of two neighbours is one edit),
  // giving up once it passes `max`.
  function editDistance(a, b, max) {
    if (Math.abs(a.length - b.length) > max) return max + 1;
    let p2 = null, p1 = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      let low = i;
      for (let j = 1; j <= b.length; j++) {
        const cost = a[i - 1] === b[j - 1] ? 0 : 1;
        let d = Math.min(p1[j] + 1, cur[j - 1] + 1, p1[j - 1] + cost);
        if (p2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d = Math.min(d, p2[j - 2] + 1);
        cur[j] = d;
        if (d < low) low = d;
      }
      if (low > max) return max + 1;
      p2 = p1; p1 = cur;
    }
    return p1[b.length];
  }
  // How far apart two keys may be: 0 under 5 letters, 1 up to 7, then 2.
  // -1 when they're not a spelling of the same name.
  function spellGap(ka, kb) {
    if (ka === kb) return 0;
    if ((ka.match(/\d+/g) || []).join(',') !== (kb.match(/\d+/g) || []).join(',')) return -1;
    const n = Math.min(ka.length, kb.length), tol = n < 5 ? 0 : n < 8 ? 1 : 2;
    if (!tol) return -1;
    const d = editDistance(ka, kb, tol);
    return d <= tol ? d : -1;
  }
  // The one candidate spelled closest to `value`, or '' when none is close or
  // two different names are equally close. Candidates sharing a key count once.
  function closestSpelling(value, cands) {
    const k = spellKey(value);
    if (!k) return '';
    let best = '', bestKey = '', bestD = 99, tie = false;
    cands.forEach(c => {
      const kc = spellKey(c);
      if (!kc) return;
      const d = spellGap(k, kc);
      if (d < 0) return;
      if (d < bestD) { best = c; bestKey = kc; bestD = d; tie = false; }
      else if (d === bestD && kc !== bestKey) tie = true;
    });
    return tie ? '' : best;
  }
  // rows' Crop & Variety column (ci) against the dropdown `list`. →
  //   fixes  Map from-value → { to, why, rows }   (from is the value as written)
  //   check  [{ from, to, rows, toRows }]          close, but no clear winner
  function spellCropVariety(rows, ci, list) {
    const fixes = new Map(), check = [];
    const listUp = new Set(list.map(v => v.toUpperCase()));
    const tally = new Map();
    rows.forEach(r => { const v = r[ci]; if (v) tally.set(v, (tally.get(v) || 0) + 1); });

    // Crops: the template's spelling, else the one most rows use.
    const tplCrops = new Map(), tplVars = new Map();
    list.forEach(o => {
      const s = splitCrop(o), ck = s.crop.toUpperCase();
      if (!tplCrops.has(ck)) tplCrops.set(ck, s.crop);
      if (!tplVars.has(ck)) tplVars.set(ck, []);
      if (s.variety) tplVars.get(ck).push(s.variety);
    });
    const dataCrops = new Map();
    tally.forEach((n, v) => { const c = tidy(splitCrop(tidy(v)).crop); dataCrops.set(c, (dataCrops.get(c) || 0) + n); });
    const cropFix = new Map();
    const fixCrop = c => {
      if (cropFix.has(c)) return cropFix.get(c);
      let out = c, why = '';
      const up = c.toUpperCase();
      if (tplCrops.has(up)) out = tplCrops.get(up);
      else {
        const plu = s => s.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean).map(w => (w.length > 3 && w.endsWith('s') ? w.slice(0, -1) : w)).join('');
        const byPlural = [...tplCrops.values()].filter(t => plu(t) === plu(c));
        const hit = byPlural.length === 1 ? byPlural[0] : closestSpelling(c, [...tplCrops.values()]);
        if (hit) { out = hit; why = 'crop as in the dropdown'; }
        else {
          const own = dataCrops.get(c) || 0;
          const big = [...dataCrops.keys()].filter(o => o !== c && dataCrops.get(o) >= 2 * own);
          const h2 = closestSpelling(c, big);
          if (h2) { out = h2; why = 'crop as most rows spell it'; }
        }
      }
      const r = { crop: out, why: why };
      cropFix.set(c, r);
      return r;
    };

    // Pass 1: the template, then known grape names.
    const step1 = new Map();   // value → { crop, variety, whys[] }
    tally.forEach((n, v) => {
      if (listUp.has(v.toUpperCase())) return;
      const p = splitCrop(tidy(v)), cf = fixCrop(tidy(p.crop));
      const whys = cf.why ? [cf.why] : [];
      let variety = tidy(p.variety);
      if (variety) {
        const tv = tplVars.get(cf.crop.toUpperCase()) || [];
        const exact = tv.find(x => x.toUpperCase() === variety.toUpperCase());
        let hit = exact || closestSpelling(variety, tv);
        if (hit && !exact) whys.push('matches the dropdown');
        if (!hit && GRAPE_CROP.test(cf.crop) && !WINE_KEYS.has(spellKey(variety))) {
          hit = closestSpelling(variety, WINE_GRAPES);
          if (hit) whys.push('known grape variety');
        }
        if (hit) variety = hit;
      }
      step1.set(v, { crop: cf.crop, variety: variety, whys: whys });
    });

    // Pass 2: the spelling most rows use, per crop, after pass 1.
    const byCrop = new Map();   // CROP → Map(variety → rows)
    tally.forEach((n, v) => {
      const s = step1.get(v) || { crop: splitCrop(v).crop, variety: splitCrop(v).variety };
      const ck = s.crop.toUpperCase();
      if (!byCrop.has(ck)) byCrop.set(ck, new Map());
      const m = byCrop.get(ck);
      if (s.variety) m.set(s.variety, (m.get(s.variety) || 0) + n);
    });
    step1.forEach((s, v) => {
      if (s.variety && !s.whys.some(w => /dropdown|known/.test(w))) {
        const m = byCrop.get(s.crop.toUpperCase()), own = m.get(s.variety) || 0, k = spellKey(s.variety);
        const inTpl = (tplVars.get(s.crop.toUpperCase()) || []).some(x => x.toUpperCase() === s.variety.toUpperCase());
        // Same letters, other casing or spacing: the form most rows use.
        const same = [...m.keys()].filter(o => o !== s.variety && spellKey(o) === k && m.get(o) > own)
          .sort((a, b) => m.get(b) - m.get(a));
        if (same.length) { s.variety = same[0]; s.whys.push('as ' + m.get(same[0]) + ' other rows write it'); }
        else if (!inTpl && !WINE_KEYS.has(k)) {
          const big = [...m.keys()].filter(o => o !== s.variety && m.get(o) >= 2 * own);
          const hit = closestSpelling(s.variety, big);
          if (hit) { s.variety = hit; s.whys.push('as ' + m.get(hit) + ' other rows spell it'); }
          else {
            // Close to a name with about as many rows: worth a look, not a guess.
            const near = closestSpelling(s.variety, [...m.keys()].filter(o => o !== s.variety && spellKey(o) !== k));
            // A tie is listed once, under the name that sorts later.
            if (near && (m.get(near) > own || (m.get(near) === own && near.localeCompare(s.variety) < 0))) {
              check.push({ from: v, to: s.crop + '-' + near, rows: tally.get(v), toRows: m.get(near) });
            }
          }
        }
      }
      const to = s.variety ? s.crop + '-' + s.variety : s.crop;
      if (to === v) return;
      fixes.set(v, { to: to, why: s.whys.length ? s.whys.join(', ') : 'extra spaces removed', rows: tally.get(v) });
    });
    return { fixes: fixes, check: check };
  }

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
        'Site': ['Site Name', 'Ranch Name'],
        'Name': ['Location Name'],
        'Alt ID': ['Location Alt ID'],
        'Acreage': ['Acres'],
        'Plant Count': ['Plant Quantity'],
        'Clone/Subvariety': ['Clone', 'Subvariety'],
        'Mulch Type': ['Multch Type'],
        'Bed Width, in.': ['Bed Width'],
        'Plant Spacing, in.': ['Plant Spacing'],
        'Row Spacing, in.': ['Row Spacing'],
        'Post Spacing, in.': ['Post Spacing'],
        'Custom Data 1': ['Loc Data 1'],
        'Custom Data 2': ['Loc Data 2'],
        'Custom Data 3': ['Loc Data 3'],
        // The Detailed Version's "Producing" (Full Production / Not Full
        // Production / No) — matched to Full / Partial / None in autoMatch.
        'Production Status': ['Producing'],
        'Production Start': ['Producing Start Year', 'Production Start Year']
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
      key: 'packstyles', label: 'Pack Styles', needsTemplate: true,
      archived: 'Archived',
      // Implementation Data Template names. Their "Unit Type (for Piecework
      // Pay)" holds the piecework unit (Box), not the weight unit the
      // template's Unit takes (LB / OZ / G), so it is deliberately not an alias.
      aliases: {
        'Pack Style Name': ['Name', 'Pack Style', 'Pack Name', 'Package Name'],
        'Crop Type': ['Crop', 'Crop Name', 'Commodity'],
        'Count': ['Pack Count'],
        'Pack Type': ['Package Type', 'Container'],
        'Unit': ['Weight Unit', 'UOM']
      },
      dates: [],
      numeric: ['Weight', 'Count'],
      hint: 'Bulk pack styles create template. Its dropdowns validate Crop Type, Method, Pack Type and Unit.'
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
  let exist = {};  // key → [{ fileName, sheetName, headers, rows }]: what the destination org already has
  let cellEdits = {};  // domain → "<rid>|<colIdx>" → value: your edits in the preview
  let removed = {};    // domain → Set of rid: rows you left out
  let dupOk = {};      // domain → Set of rid: possible duplicates you chose to keep
  let siteAddr = { all: '', ranch: {}, site: {} }; // typed site addresses — every site / ranch → / rid →
  let yearFill = { mm: '', dd: '', range: 'first' }; // month + day for year-only dates
  let pvFocus = null;  // { key, ci }: preview narrowed to the rows with that column empty
  let placeNotes = []; // what auto-placement put where, shown on the file strip
  let built = null;// key → { headers, rows, ... }
  let previewKey = null;
  let migrate = true;  // rewrite off-list values to a single-option destination list
  let setToday = true; // stamp Locations' Start Date with today
  let ageFixes = {};   // "<domain>|<source row index>" → corrected YYYY-MM-DD
  let columnFills = {};// domain → colIdx → { val, mode:'all'|'blank' }  (Column Fill panel)
  let smartFix = {};   // domain → "<colIdx>||<UPPER VALUE>" → canonical dropdown value
  let spellKeep = {};  // domain → Set of "<colIdx>||<UPPER VALUE>": spelling fixes you undid (kept as typed)
  let srcOverride = {};// domain → loose template header → source header name (Site* ← Location Group)
  let siteFillOk = {}; // domain → Set of "<first source row>|<site>": Blank Site groups you confirmed copying down
  let activeSection = null;  // section button that's open (see SECTIONS)
  let sectionPicked = false; // true once you click a section — until then the first with something waiting opens
  let filesOpen = null;      // upload slots: null = open until something builds, then folded
  let pvFilter = 'all';      // preview rows: 'all' | 'fix' | 'chg' | 'any' (IMWorkbench.filterBar)
  let pvColsOnly = false;    // preview: only the columns with a problem or a change
  let pvHidden = false;      // preview grid folded down to its file tabs
  let dtSplit = null;        // side by side / stacked (IMWorkbench.split, shared with the single tools)
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

      reader.onload = async e => { try {
        const buf = e.target.result;
        // A plain read: cellStyles stretches !ref over styled-but-empty cells
        // and throws on some PickTrace exports.
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
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
              headers: headers, dropdowns: parseDropdowns(wb),
              restricted: await readRestricted(buf, headers)
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
      } catch (err) { reject(err); } };
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
    // A "(Detailed Version)" tab names Site in the banner row and carries
    // metadata rows under the header — see IMSheetLayout in utils.js.
    const lay = headerOnTop ? null : IMSheetLayout.resolve(lines.map(l => l.r), h);
    const headers = lay ? lay.headers : lines[h].r.map(str);
    let body = lines.slice(h + 1).filter(l => !(lay && lay.isMeta(l.r)));
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
    if (/^pack ?(styles?|sizes?|types?)$/.test(n)) return 'packstyles';
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

  // ─── Which columns PickTrace restricts ───
  // The template puts a list validation on the DATA ENTRY columns that only
  // take listed values. Its DROP-DOWN INPUTS sheet also lists free-text
  // columns (Rootstock, Location Group, Mulch Type…) with the values the org
  // already uses, and those take anything. SheetJS doesn't read validations,
  // so they come from the sheet XML — the same reading as Template
  // Standardize. → Set of loose header keys, or null when the file can't say
  // (then every listed column counts as restricted).
  async function readRestricted(buf, headers) {
    if (!window.IMXlsxTemplate) return null;
    try {
      for (const e of IMXlsxTemplate.zipRead(buf)) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name)) continue;
        const xml = await IMXlsxTemplate.entryText(e);
        if (!/<dataValidation\b/.test(xml) || !/DROP-?DOWN/i.test(xml)) continue;
        const set = new Set();
        (xml.match(/<dataValidation\b[^>]*>/g) || []).forEach(tag => {
          if (!/type="list"/.test(tag)) return;
          ((tag.match(/sqref="([^"]*)"/) || [])[1] || '').split(/\s+/).forEach(ref => {
            const col = (ref.match(/^([A-Z]+)/) || [])[1];
            const h = col ? headers[XLSX.utils.decode_col(col)] : '';
            if (h) set.add(normLoose(h));
          });
        });
        return set.size ? set : null;
      }
    } catch (err) { /* unreadable — every listed column counts as restricted */ }
    return null;
  }
  // Site and Crop & Variety are always checked, whatever the file says.
  const ALWAYS_RESTRICTED = ['site', 'cropvariety'];
  function isFreeText(tpl, header) {
    const k = normLoose(header);
    return !!(tpl && tpl.restricted) && !tpl.restricted.has(k) && ALWAYS_RESTRICTED.indexOf(k) < 0;
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
    if (h.has('packstylename') && h.has('croptype')) return 'packstyles';
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
    if (h.has('packstylename') || (h.has('packtype') && (h.has('packcount') || h.has('croptype') || h.has('weight')))) return 'packstyles';
    return null;
  }

  // A value whose dropdown entry is unambiguous even though the text differs.
  // Returns that entry, or '' to leave the value for Smart Fixes.
  //   Crop & Variety  "Cherries" → the crop's only entry, else its "-Other"
  //   yes / no        "Y", "Yes", "N", "No" → TRUE / FALSE
  //   short codes     "M" → "MALE", "es" → "es-MX" — begins exactly one option
  //   Method          "Conventional" → "CONV" — the one code the value starts with
  //   Pack Type       "Clamshell" → "Clamshells" — singular / plural of one entry
  //   Production      "Full Production" → Full, "Not Full Production" → Partial,
  //     Status        "No" → None (the Detailed Version's Producing column)
  function autoMatch(header, value, list) {
    const v = str(value), up = v.toUpperCase();
    if (!v) return '';
    const hk = normLoose(header);
    if (hk === 'productionstatus') {
      const want = /^not\b|partial/i.test(v) ? 'PARTIAL' : /^full\b/i.test(v) ? 'FULL'
        : /^(no|none|not producing)$/i.test(v) ? 'NONE' : '';
      const hit = want && list.find(o => o.toUpperCase() === want);
      if (hit) return hit;
    }
    if (hk === 'method') {
      const hits = list.filter(o => o.length >= 3 && up.indexOf(o.toUpperCase()) === 0);
      if (hits.length === 1) return hits[0];
    }
    if (hk === 'packtype') {
      const lv = normLoose(v);
      const hits = list.filter(o => {
        const lo = normLoose(o);
        return lo === lv + 's' || lo === lv + 'es' || lv === lo + 's' || lv === lo + 'es';
      });
      if (hits.length === 1) return hits[0];
    }
    if (hk === 'cropvariety') {
      if (v.indexOf('-') >= 0) return '';
      const mine = list.filter(o => splitCrop(o).crop.toUpperCase() === up);
      if (mine.length === 1) return mine[0];
      return mine.find(o => splitCrop(o).variety.toUpperCase() === 'OTHER') || '';
    }
    // "USA" / "United States" → whichever spelling the list uses for the US;
    // "Spanish" → es-MX. Matched as codes on both sides, so a list of "US"
    // and a list of "United States" both work.
    const canon = /country$/.test(hk) ? countryCode : (hk === 'languagepreference' || hk === 'language') ? langCode : null;
    if (canon && canon(v)) {
      const hits = list.filter(o => canon(o) === canon(v));
      if (hits.length === 1) return hits[0];
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

  const COUNTRY_FORMS = {
    US: ['us', 'usa', 'u.s.', 'u.s.a.', 'united states', 'united states of america', 'america', 'estados unidos', 'eeuu', 'ee.uu.'],
    MX: ['mx', 'mex', 'mexico', 'méxico'],
    CA: ['ca', 'can', 'canada']
  };
  const LANG_FORMS = {
    en: ['en', 'en-us', 'eng', 'english', 'ingles', 'inglés'],
    es: ['es', 'es-mx', 'es-us', 'spa', 'spanish', 'espanol', 'español']
  };
  function formCode(forms, v) {
    const s = str(v).toLowerCase().replace(/\s+/g, ' ');
    return Object.keys(forms).find(c => forms[c].indexOf(s) >= 0) || '';
  }
  const countryCode = v => formCode(COUNTRY_FORMS, v);
  const langCode = v => formCode(LANG_FORMS, v);

  // ═══ Mapping ═══
  // ctx (all optional):
  //   pending  loose template header → values this run creates elsewhere (the
  //            Sites file's names, for Locations' Site*). They count as valid
  //            dropdown values, so they aren't reported as needing building.
  //   known    the same, for values the destination org already has (sites
  //            in the existing exports) — valid, and not "created by this run".
  //   extra    Sites only: new site names from Locations, one row each.
  //   ex       this domain's existing-export index (existIndex()).
  function mapDomain(dom, ctx) {
    ctx = ctx || {};
    const pending = ctx.pending || null;
    const extra = ctx.extra || [];
    const ex = ctx.ex || null;
    const tpl = tpls[dom.key] || null;
    // The Sites file can come from Locations alone — no Sites export needed.
    const src = srcs[dom.key] || (extra.length && tpl ? { headers: [], rows: [] } : null);
    if (!src) return null;
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

    // A source header carrying a bracketed note — "Plant Quantity (Tree/Vine
    // Quantity)", "Site (Ranch Name)*", "Alt ID (payroll ID#)" —
    // matched on the words outside the brackets, exactly or by alias. Runs
    // after the exact/alias pass so it never takes a column another template
    // column names outright, and before the fuzzy guess.
    const bareKey = h => normLoose(String(h).replace(/\([^)]*\)/g, ''));
    headers.forEach((h, ci) => {
      if (colSrc[ci] >= 0 || displaced.has(ci)) return;
      const al = dom.aliases[h] || dom.aliases[h.replace(/\*\s*$/, '').trim()] || [];
      const want = [normLoose(h)].concat(al.map(normLoose));
      const si = eff.headers.findIndex((sh, i) => !claimedSrc.has(i) && /\(/.test(str(sh)) && want.indexOf(bareKey(sh)) >= 0);
      if (si >= 0) { colSrc[ci] = si; claimedSrc.add(si); }
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
    const numSet = new Set((dom.numeric || []).map(normLoose));
    const archivedIdx = srcIdx.has(normLoose(dom.archived)) ? srcIdx.get(normLoose(dom.archived)) : -1;
    const dedupeCol = dom.dedupe ? headers.findIndex(h => normLoose(h) === normLoose(dom.dedupe)) : -1;

    const dobCol = headers.findIndex(h => normLoose(h) === 'dateofbirth');
    const ssnCol = headers.findIndex(h => normLoose(h) === 'ssn');
    const phoneCols = (dom.phones || []).map(p => headers.findIndex(h => normLoose(h) === normLoose(p))).filter(i => i >= 0);
    const phoneBad = [];

    // rid: a row's identity across rebuilds — its source row index, or
    // "new:<SITE>" for a site added from Locations. Your edits, removals and
    // per-site addresses are keyed by it.
    let rows = [], rowSrc = [], rid = [];
    const dropped = [];
    eff.rows.forEach((sr, si0) => {
      if (archivedIdx >= 0 && isTrue(sr[archivedIdx])) {
        dropped.push({ domain: dom.label, reason: dom.archived + ' = true', detail: describeRow(headers, colSrc, sr) });
        return;
      }
      const row = headers.map((h, ci) => {
        const si = colSrc[ci];
        if (si < 0) return '';
        const raw = sr[si];
        const k = normLoose(h);
        return dateSet.has(k) ? toYMD(raw) : numSet.has(k) ? cleanNumber(raw, h) : str(raw);
      });
      if (!row.some(v => v !== '')) return;
      // A placeholder row ("n/a" in every cell — the client has none of these)
      // is not data; it's reported rather than exported.
      if (!row.some(v => v !== '' && !/^n\/?a$/i.test(v))) {
        dropped.push({ domain: dom.label, reason: 'every cell is n/a', detail: describeRow(headers, colSrc, sr) });
        return;
      }
      if (ssnCol >= 0 && row[ssnCol]) row[ssnCol] = normSsn(row[ssnCol]);
      phoneCols.forEach(pi => {
        if (!row[pi]) return;
        const p = normPhone(row[pi]);
        if (p.ok) row[pi] = p.value;
        else phoneBad.push({ domain: dom.label, column: headers[pi], value: row[pi] });
      });
      // A site with no Name can't be created. Repeats of a name are handled
      // below (folded when they agree, a collision when they don't).
      if (dedupeCol >= 0 && !row[dedupeCol]) {
        dropped.push({ domain: dom.label, reason: 'no ' + headers[dedupeCol].replace(/\*\s*$/, ''), detail: describeRow(headers, colSrc, sr) });
        return;
      }
      // A corrected birth date you entered in the Under Minimum Age panel wins
      // over the export's, and is keyed to the source row so it survives rebuilds.
      if (dobCol >= 0) {
        const fix = ageFixes[dom.key + '|' + si0];
        if (fix) row[dobCol] = fix;
      }
      rows.push(row);
      rowSrc.push(si0);
      rid.push(String(si0));
    });

    // prov: "<rid>|<ci>" → why the tool (or you) wrote that cell, for writes
    // the source → template rewrite can't explain (preview hover).
    const prov = new Map();
    const nameCi = headers.findIndex(h => normLoose(h) === 'name');
    // New sites named in Locations: a row each, Name only — the rest comes
    // from the addresses, fills and single-option lists like any other site.
    extra.forEach(s => {
      const row = headers.map(() => '');
      if (nameCi >= 0) row[nameCi] = s;
      rows.push(row); rowSrc.push(-1); rid.push('new:' + s.toUpperCase());
      if (nameCi >= 0) prov.set('new:' + s.toUpperCase() + '|' + nameCi, 'New site — named in Locations, not in PickTrace yet');
    });
    // Keep only the rows test() passes, in step across rows / rowSrc / rid.
    const keepWhere = test => {
      const R = [], S = [], I = [];
      rows.forEach((r, i) => { if (test(r, i)) { R.push(r); S.push(rowSrc[i]); I.push(rid[i]); } });
      rows = R; rowSrc = S; rid = I;
    };

    // Rows you left out with × in the preview.
    const removedList = [];
    const rm = removed[dom.key];
    if (rm && rm.size) {
      keepWhere((r, i) => {
        if (!rm.has(rid[i])) return true;
        removedList.push({ domain: dom.label, key: dom.key, rid: rid[i], detail: r.filter(Boolean).slice(0, 3).join(' · ') || '(row)' });
        return false;
      });
    }

    // Your edits in the preview. Applied now, so a renamed row is what the
    // collision check sees, and again after the fills so an edit always wins.
    const ed = cellEdits[dom.key] || {};
    const applyEdits = () => {
      const keys = Object.keys(ed);
      if (!keys.length) return;
      const at = new Map(rid.map((x, i) => [x, i]));
      keys.forEach(k => {
        const p = k.lastIndexOf('|'), i = at.get(k.slice(0, p)), ci = +k.slice(p + 1);
        if (i == null || !(ci >= 0 && ci < headers.length)) return;
        rows[i][ci] = ed[k];
        prov.set(k, 'Your edit');
      });
    };
    applyEdits();

    // Blank Site under a row that has one: the client wrote each site once for
    // its group and left the rows below it empty. Offered under Blank Site —
    // nothing is copied until you confirm the group, which is keyed by its
    // first source row and the site above it — the choice survives rebuilds,
    // and never carries over to a different group after a remap.
    const siteGaps = [];
    const siteCi = dom.key === 'locations' ? headers.findIndex(h => normLoose(h) === 'site') : -1;
    if (siteCi >= 0) {
      const ok = siteFillOk[dom.key] || new Set();
      let last = '', cur = null;
      rows.forEach((r, ri) => {
        if (r[siteCi]) { last = r[siteCi]; cur = null; return; }
        if (!last) return;
        if (!cur) { cur = { domain: dom.label, key: dom.key, start: rowSrc[ri], site: last, rows: [], rids: [], names: [] }; siteGaps.push(cur); }
        cur.rows.push(ri);
        cur.rids.push(rid[ri]);
        if (nameCi >= 0 && r[nameCi] && cur.names.length < 3) cur.names.push(r[nameCi]);
      });
      siteGaps.forEach(g => {
        g.id = g.start + '|' + g.site;
        g.applied = ok.has(g.id);
        if (g.applied) g.rows.forEach(ri => { rows[ri][siteCi] = g.site; });
      });
    }

    // ─── Site addresses ───
    // Typed in the Site addresses section: one for every site, one per ranch
    // (the name before " - "), one per site. A site's own address always
    // wins; the ranch's or everyone's only fills a site the export gave no
    // street. Split into Address1 / City / State / Zip / Country.
    const siteAddrRows = [];
    if (dom.key === 'sites' && nameCi >= 0) {
      const at = {};
      ADDR_PARTS.forEach(k => { at[k] = headers.findIndex(h => normLoose(h) === k); });
      const ranches = ranchMap(rows.map(r => r[nameCi]).filter(Boolean));
      rows.forEach((r, i) => {
        const name = r[nameCi], ranch = ranches.get(name) || name;
        const own = siteAddr.site[rid[i]] || '';
        const fromExport = at.address1 >= 0 && r[at.address1] !== '';
        let text = '', from = '';
        if (own) { text = own; from = 'this site'; }
        else if (!fromExport && siteAddr.ranch[ranch]) { text = siteAddr.ranch[ranch]; from = ranch; }
        else if (!fromExport && siteAddr.all) { text = siteAddr.all; from = 'every site'; }
        if (text) {
          const p = parseSiteAddr(text);
          ADDR_PARTS.forEach(k => {
            const ci = at[k];
            if (ci < 0 || !p[k] || (!own && r[ci])) return;
            r[ci] = p[k];
            prov.set(rid[i] + '|' + ci, 'From the address for ' + from);
          });
        }
        siteAddrRows.push({ rid: rid[i], name: name, ranch: ranch, own: own,
          from: own ? 'site' : fromExport ? 'export' : !text ? '' : from === 'every site' ? 'all' : 'ranch' });
      });
    }

    // ─── Year-only dates ───
    // "2012" in a date column: expanded with the month and day you pick in
    // Year-only dates, keeping each row's year. Start Date is left alone.
    const yearOnly = [];
    const canYear = !!(yearFill.mm && yearFill.dd);
    headers.forEach((h, ci) => {
      const k = normLoose(h);
      if (!dateSet.has(k) || k === 'startdate') return;
      let left = 0, filled = 0;
      rows.forEach((r, i) => {
        if (!isYearOnlyDate(r[ci])) return;
        const v = canYear ? expandYearOnlyDate(r[ci], yearFill.mm, yearFill.dd, yearFill.range) : r[ci];
        if (v === r[ci]) { left++; return; }
        r[ci] = v;
        filled++;
        prov.set(rid[i] + '|' + ci, 'Year-only date — given the month and day you chose');
      });
      if (left || filled) yearOnly.push({ domain: dom.label, key: dom.key, column: h, left: left, filled: filled });
    });

    // ─── Already in PickTrace · repeats · name collisions ───
    // PickTrace keeps one record per name (Sites, Pack Styles, Jobs) or per
    // Site + Name (Locations), and silently drops the rest. So:
    //   - a row the existing export already has is left out;
    //   - repeats that agree on every column fold into one row;
    //   - repeats that disagree, or a Locations name PickTrace already uses
    //     for a different block, are listed under Name collisions.
    const already = [], collisions = [], possible = [];
    const keyCols = (KEY_COLS[dom.key] || []).map(k => headers.findIndex(h => normLoose(h) === k));
    if (keyCols.length && keyCols.every(ci => ci >= 0)) {
      const keyOf = r => keyCols.some(ci => !str(r[ci])) ? '' : keyCols.map(ci => locNorm(r[ci])).join('||');
      const keyLabel = keyCols.map(ci => headers[ci].replace(/\*\s*$/, '')).join(' + ');
      const cvCi = headers.findIndex(h => normLoose(h) === 'cropvariety');
      const byBlock = !!(ex && ex.blockKeys && cvCi >= 0);
      if (ex && ex.keys && ex.keys.size) {
        const counts = new Map();
        rows.forEach(r => { const k = keyOf(r); if (k) counts.set(k, (counts.get(k) || 0) + 1); });
        keepWhere(r => {
          const k = keyOf(r);
          if (!k) return true;
          // Locations without a crop to compare can't tell two blocks of one
          // name apart, so a name repeated in the file is never dropped.
          const hit = byBlock ? ex.blockKeys.has(k + '||' + locNorm(r[cvCi]))
            : ex.keys.has(k) && (dom.key !== 'locations' || counts.get(k) === 1);
          if (!hit) return true;
          already.push({ domain: dom.label, key: dom.key, detail: keyCols.map(ci => r[ci]).join(' · '),
            by: 'Same ' + keyLabel + (byBlock ? ' and Crop & Variety' : '') });
          return false;
        });
      }
      const bucket = () => {
        const m = new Map();
        rows.forEach((r, i) => { const k = keyOf(r); if (!k) return; if (!m.has(k)) m.set(k, []); m.get(k).push(i); });
        return m;
      };
      const gone = new Set();
      bucket().forEach(idx => {
        if (idx.length < 2) return;
        // A blank agrees with anything; two different values don't.
        const agree = headers.every((_, ci) => {
          const vals = new Set();
          idx.forEach(i => { const v = str(rows[i][ci]).toUpperCase(); if (v) vals.add(v); });
          return vals.size <= 1;
        });
        if (!agree) return;
        const keep = rows[idx[0]];
        idx.slice(1).forEach(i => {
          headers.forEach((_, ci) => { if (!keep[ci] && rows[i][ci]) keep[ci] = rows[i][ci]; });
          gone.add(i);
        });
        dropped.push({ domain: dom.label, reason: plural(idx.length, 'row') + ' with the same ' + keyLabel + ' and values — kept as one',
          detail: keyCols.map(ci => keep[ci]).join(' · ') });
      });
      if (gone.size) keepWhere((r, i) => !gone.has(i));
      bucket().forEach((idx, k) => {
        const taken = dom.key === 'locations' && ex && ex.keys && ex.keys.has(k) ? (ex.byName.get(k) || []) : null;
        if (idx.length < 2 && !taken) return;
        collisions.push({ domain: dom.label, key: dom.key, label: keyLabel, keyCols: keyCols,
          value: keyCols.map(ci => rows[idx[0]][ci]).join(' · '), rids: idx.map(i => rid[i]), rows: idx, taken: taken });
      });
    }

    // Employees: the same person twice, by the rules Employee Standardize
    // uses. An Alt ID or SSN match counts only when the names also agree
    // (recycled and test IDs are common); Name + birth date counts unless the
    // SSNs differ. Confident matches against the existing export are left
    // out as Already in PickTrace, against an earlier row as a duplicate; an
    // ID clash with a different name, or a name-only match nothing rules
    // out, is kept and listed under Possible duplicates.
    if (dom.key === 'employees') {
      const at = k => headers.findIndex(h => normLoose(h) === k);
      const fnI = at('firstname'), lnI = at('lastname'), altI = at('altid'), ssnI = at('ssn');
      const ok = dupOk[dom.key] || new Set();
      const db = ex && ex.emp;
      const seenAlt = new Map(), seenSsn = new Map(), seenND = new Map(), seenName = new Map();
      const gone = new Set();
      rows.forEach((r, i) => {
        const g = ci => (ci >= 0 ? str(r[ci]) : '');
        const name = (g(fnI) + ' ' + g(lnI)).trim();
        const nk = normName(name), tok = nameTokens(name);
        const a = g(altI).toUpperCase(), sk = ssnKey(g(ssnI)), dob = g(dobCol);
        const nd = nk && dob ? nk + '|' + dob : '';
        const disp = name + (a ? ' (Alt ID ' + g(altI) + ')' : '');
        const earlier = ' — earlier in this file';
        const cls = (() => {
          if (db && a && db.byAlt.has(a)) {
            const m = db.byAlt.get(a);
            return nameAgree(tok, nameTokens(m.name)) ? { db: 'Alt ID', other: m.display } : { conflict: 'Alt ID — name differs', other: m.display };
          }
          if (db && sk && db.bySsn.has(sk)) {
            const m = db.bySsn.get(sk);
            return nameAgree(tok, nameTokens(m.name)) ? { db: 'SSN', other: m.display } : { conflict: 'SSN — name differs', other: m.display };
          }
          if (db && nd && db.byNameDob.has(nd)) {
            const m = db.byNameDob.get(nd);
            return sk && m.sKey && sk !== m.sKey ? { conflict: 'Name + birth date — SSN differs', other: m.display }
              : { db: 'Name + birth date', other: m.display };
          }
          if (a && seenAlt.has(a)) {
            const s = seenAlt.get(a);
            return nameAgree(tok, nameTokens(s.name)) ? { file: 'Alt ID', other: s.disp } : { conflict: 'Alt ID — name differs', other: s.disp + earlier };
          }
          if (sk && seenSsn.has(sk)) {
            const s = seenSsn.get(sk);
            return nameAgree(tok, nameTokens(s.name)) ? { file: 'SSN', other: s.disp } : { conflict: 'SSN — name differs', other: s.disp + earlier };
          }
          if (nd && seenND.has(nd)) {
            const s = seenND.get(nd);
            return sk && s.sKey && sk !== s.sKey ? { conflict: 'Name + birth date — SSN differs', other: s.disp + earlier }
              : { file: 'Name + birth date', other: s.disp };
          }
          return null;
        })();
        if (cls && cls.db) {
          already.push({ domain: dom.label, key: dom.key, detail: disp, by: cls.db + ' — ' + cls.other });
          gone.add(i);
          return;
        }
        if (cls && cls.file) {
          dropped.push({ domain: dom.label, reason: 'same person as ' + cls.other + earlier + ' (' + cls.file + ')', detail: disp });
          gone.add(i);
          return;
        }
        const me = { disp: disp, name: name, aKey: a, sKey: sk, dob: dob };
        if (a && !seenAlt.has(a)) seenAlt.set(a, me);
        if (sk && !seenSsn.has(sk)) seenSsn.set(sk, me);
        if (nd && !seenND.has(nd)) seenND.set(nd, me);
        let flag = cls && cls.conflict ? { reason: cls.conflict, other: cls.other } : null;
        if (!flag && nk) {
          const ref = db && db.byName.get(nk);
          const f = seenName.get(nk);
          if (ref && !strongConflict(a, sk, dob, ref)) flag = { reason: 'Same name, in PickTrace', other: ref.display };
          else if (f && !strongConflict(a, sk, dob, f)) flag = { reason: 'Same name', other: f.disp + earlier };
        }
        if (nk && !seenName.has(nk)) seenName.set(nk, me);
        if (flag && !ok.has(rid[i])) possible.push({ domain: dom.label, key: dom.key, rid: rid[i], name: disp, reason: flag.reason, other: flag.other });
      });
      if (gone.size) keepWhere((r, i) => !gone.has(i));
    }

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
    applyEdits();

    // Crop & Variety typos → the spelling the dropdown, a known grape name or
    // most rows use (spellCropVariety). Fixes you undid stay as typed but are
    // still listed, so you can redo them.
    const spelled = [], spellCheck = [];
    headers.forEach((h, ci) => {
      if (normLoose(h) !== 'cropvariety' || isFreeText(tpl, h)) return;
      const list = (dropdownFor(dropdowns, h) || []).concat((ctx.known && ctx.known[normLoose(h)]) || []);
      const keep = spellKeep[dom.key] || new Set();
      const res = spellCropVariety(rows, ci, list);
      res.fixes.forEach((f, from) => {
        const kept = keep.has(ci + '||' + from.toUpperCase());
        spelled.push({ domain: dom.label, key: dom.key, ci: ci, column: h, from: from, to: f.to, rows: f.rows, why: f.why, kept: kept });
        if (kept) return;
        rows.forEach(r => { if (r[ci] === from) r[ci] = f.to; });
      });
      res.check.forEach(c => spellCheck.push(Object.assign({ domain: dom.label, key: dom.key, ci: ci, column: h }, c)));
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
    //
    // A column the template lists but doesn't validate (isFreeText) takes any
    // value: its list only snaps casing, and nothing in it is reported.
    const mismatches = [], remapped = [], noOptions = [], matched = [], fromPending = [];
    headers.forEach((h, ci) => {
      const own = dropdownFor(dropdowns, h);
      const extra = (pending && pending[normLoose(h)]) || [];
      const knownVals = (ctx.known && ctx.known[normLoose(h)]) || [];
      if (!own && !extra.length && !knownVals.length) return;
      const free = isFreeText(tpl, h);
      const list = (own || []).slice();
      const ownUp = new Set(list.map(v => v.toUpperCase()));
      const extraUp = new Set();
      extra.forEach(v => { const k = v.toUpperCase(); if (!ownUp.has(k) && !extraUp.has(k)) { extraUp.add(k); list.push(v); } });
      knownVals.forEach(v => { const k = v.toUpperCase(); if (!ownUp.has(k) && !extraUp.has(k)) { ownUp.add(k); list.push(v); } });

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
        if (free) return;
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
      // Crew, Crop & Variety, Crop Type, Site and H2A Contract never migrate: an
      // org with one crew or one crop says nothing about where these rows belong, so an
      // off-list value there is something to create (see the build banner).
      if (migrate && own.length === 1 && !extraUp.size && BUILD_COLUMNS.indexOf(normLoose(h)) < 0) {
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
      if (!list || list.length !== 1 || isFreeText(tpl, h)) return;
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

    // Rows were left out after some findings were recorded: point those
    // findings at the rows' final positions.
    const at = new Map(rid.map((x, i) => [x, i]));
    siteGaps.forEach(g => { g.rows = g.rids.map(x => at.get(x)).filter(i => i != null); });
    possible.forEach(p => { p.row = at.get(p.rid); });

    // Sites still short of a required address part, for Site addresses.
    let addrMissing = 0;
    if (dom.key === 'sites') {
      const req = headers.map((h, ci) => (/\*\s*$/.test(h) && ADDR_PARTS.indexOf(normLoose(h)) >= 0 ? ci : -1)).filter(ci => ci >= 0);
      rows.forEach(r => { if (req.some(ci => r[ci] === '')) addrMissing++; });
    }

    return { headers: headers, colSrc: colSrc, rows: rows, rowSrc: rowSrc, rid: rid, prov: prov, dropped: dropped, siteGaps: siteGaps,
      mismatches: mismatches, remapped: remapped.concat(overrides), noOptions: noOptions,
      matched: crewMapped.concat(spelled.filter(s => !s.kept).map(s => ({ domain: s.domain, column: s.column, from: [s.from], to: s.to,
        rows: s.rows, why: 'Spelling fixed — ' + s.why }))).concat(matched), fromPending: fromPending,
      spelled: spelled, spellCheck: spellCheck,
      autofilled: autofilled, required: required, unmapped: unmapped, underAge: underAge,
      derived: der.derived, addrFlags: der.flags, fuzzy: fuzzy, filledCols: filledCols,
      phoneBad: phoneBad, dropdowns: dropdowns, pending: pending || null, known: ctx.known || null, key: dom.key, hasTemplate: !!tpl,
      removedList: removedList, already: already, collisions: collisions, possible: possible, yearOnly: yearOnly,
      siteAddrRows: siteAddrRows.filter(s => at.has(s.rid)), addrMissing: addrMissing, newSites: extra.length,
      // colSrc indexes these — the source plus any derived virtual columns.
      effHeaders: eff.headers, effRows: eff.rows };
  }

  // ─── Helpers for the checks above ───
  const ADDR_PARTS = ['address1', 'city', 'state', 'zip', 'country'];
  // What makes a record one record in PickTrace.
  const KEY_COLS = { sites: ['name'], locations: ['site', 'name'], packstyles: ['packstylename'], jobs: ['name'] };
  const locNorm = s => String(s == null ? '' : s).toUpperCase().trim().replace(/\s+/g, ' ');

  // A typed site address → { address1, city, state, zip, country }. The Sites
  // tool's parser (it knows Canadian provinces and postal codes) when it's
  // loaded, else this file's roster parser.
  function parseSiteAddr(text) {
    if (window.IMAddress) {
      const p = IMAddress.parse(text);
      return { address1: p.address1, city: p.city, state: p.state, zip: p.zip, country: p.country };
    }
    const p = parseRosterAddress(text);
    return { address1: p.a1, city: p.city, state: p.state, zip: p.zip, country: p.country };
  }

  // Each site's ranch: the name before " - " — or before a bare "-" when
  // another site spells that ranch with " - " ("Example-Block 2" joins
  // "Example - Block 1"); otherwise the whole name. Same rule as Template
  // Standardize. → Map name → ranch.
  function ranchMap(names) {
    const known = new Map();
    names.forEach(s => { const i = s.indexOf(' - '); if (i > 0) { const r = s.slice(0, i).trim(); if (!known.has(r.toUpperCase())) known.set(r.toUpperCase(), r); } });
    const m = new Map();
    names.forEach(s => {
      const i = s.indexOf(' - '), j = s.indexOf('-');
      const head = (i > 0 ? s.slice(0, i) : j > 0 ? s.slice(0, j) : '').trim().toUpperCase();
      m.set(s, known.get(head) || s.trim());
    });
    return m;
  }

  // Employee matching — the rules of employee-standardize.js.
  const normName = s => String(s == null ? '' : s).toUpperCase().trim().replace(/\s+/g, ' ');
  const nameTokens = s => normName(s).split(' ').filter(Boolean);
  // A real SSN: 9 digits, not a placeholder run (000-00-0558), which would
  // otherwise make many people one.
  function ssnKey(s) {
    const d = String(s == null ? '' : s).replace(/\D/g, '');
    if (d.length !== 9 || /^0+$/.test(d) || d.slice(0, 5) === '00000') return '';
    return d;
  }
  // Two names agree when they share two or more tokens — a middle name added
  // or dropped still matches; a recycled ID on a different person doesn't.
  function nameAgree(a, b) {
    if (!a.length || !b.length) return false;
    const B = new Set(b);
    let shared = 0;
    new Set(a).forEach(t => { if (B.has(t)) shared++; });
    return shared >= 2;
  }
  // Same name, but an Alt ID, SSN or birth date that differs: two people.
  function strongConflict(aKey, sKey, dob, ref) {
    return !!((aKey && ref.aKey && aKey !== ref.aKey) || (sKey && ref.sKey && sKey !== ref.sKey) || (dob && ref.dob && dob !== ref.dob));
  }

  // ─── Existing PickTrace exports ───
  // What the destination org already has, per domain. Several files for one
  // domain are read as one.
  //   locations  keys: Site+Name · blockKeys: +Crop & Variety · byName · sites
  //   sites      keys: Name (from a Sites export, plus every Site the
  //              existing Locations export uses)
  //   packstyles / jobs  keys: Name
  //   employees  emp: { byAlt, bySsn, byNameDob, byName }
  //   siteNames  every site PickTrace has (display spelling)
  function existIndex() {
    const out = {};
    const colOf = (headers, keys) => {
      const hk = headers.map(normLoose);
      for (const k of keys) { const i = hk.indexOf(k); if (i >= 0) return i; }
      return -1;
    };
    const siteNames = new Map(); // UPPER → spelling
    const addSite = v => { const s = str(v); if (s && !siteNames.has(locNorm(s))) siteNames.set(locNorm(s), s); };

    (exist.locations || []).forEach(d => {
      const si = colOf(d.headers, ['site', 'sitename', 'ranchname']), ni = colOf(d.headers, ['name', 'locationname', 'block']);
      const cv = colOf(d.headers, ['cropvariety']);
      const ix = out.locations || (out.locations = { keys: new Set(), blockKeys: cv >= 0 ? new Set() : null, byName: new Map() });
      if (cv < 0) ix.blockKeys = null;
      d.rows.forEach(r => {
        if (si >= 0) addSite(r[si]);
        if (si < 0 || ni < 0 || !str(r[si]) || !str(r[ni])) return;
        const k = locNorm(r[si]) + '||' + locNorm(r[ni]);
        ix.keys.add(k);
        if (ix.blockKeys) ix.blockKeys.add(k + '||' + locNorm(r[cv]));
        if (!ix.byName.has(k)) ix.byName.set(k, []);
        ix.byName.get(k).push(cv >= 0 ? str(r[cv]) || '(no crop)' : '');
      });
    });
    (exist.sites || []).forEach(d => {
      const ni = colOf(d.headers, ['site', 'sitename', 'name']);
      if (ni >= 0) d.rows.forEach(r => addSite(r[ni]));
    });
    if (siteNames.size) out.sites = { keys: new Set(siteNames.keys()) };
    [['packstyles', ['packstylename', 'name', 'packstyle']], ['jobs', ['name', 'jobname']]].forEach(([key, cols]) => {
      (exist[key] || []).forEach(d => {
        const ni = colOf(d.headers, cols);
        if (ni < 0) return;
        const ix = out[key] || (out[key] = { keys: new Set() });
        d.rows.forEach(r => { if (str(r[ni])) ix.keys.add(locNorm(r[ni])); });
      });
    });
    (exist.employees || []).forEach(d => {
      const at = k => colOf(d.headers, [k]);
      const fn = at('firstname'), ln = at('lastname'), alt = at('altid'), ssn = at('ssn'), dob = at('dateofbirth'), id = at('employeeid');
      if (fn < 0 || ln < 0) return;
      const ix = (out.employees || (out.employees = { emp: { byAlt: new Map(), bySsn: new Map(), byNameDob: new Map(), byName: new Map() } })).emp;
      d.rows.forEach(r => {
        const g = i => (i >= 0 ? str(r[i]) : '');
        const name = (g(fn) + ' ' + g(ln)).trim();
        if (!name) return;
        const rec = { name: name, display: name + (g(id) ? ' (ID ' + g(id) + ')' : g(alt) ? ' (Alt ID ' + g(alt) + ')' : '') + ' — in PickTrace',
          aKey: g(alt).toUpperCase(), sKey: ssnKey(g(ssn)), dob: toYMD(r[dob >= 0 ? dob : -1]) };
        const nk = normName(name);
        if (rec.aKey && !ix.byAlt.has(rec.aKey)) ix.byAlt.set(rec.aKey, rec);
        if (rec.sKey && !ix.bySsn.has(rec.sKey)) ix.bySsn.set(rec.sKey, rec);
        if (rec.dob && !ix.byNameDob.has(nk + '|' + rec.dob)) ix.byNameDob.set(nk + '|' + rec.dob, rec);
        if (!ix.byName.has(nk)) ix.byName.set(nk, rec);
      });
    });
    out.siteNames = [...siteNames.values()];
    return out;
  }
  function existCount(key) { return (exist[key] || []).reduce((n, d) => n + d.rows.length, 0); }

  function describeRow(headers, colSrc, sr) {
    for (let i = 0; i < headers.length; i++) {
      if (colSrc[i] >= 0 && str(sr[colSrc[i]])) return str(sr[colSrc[i]]);
    }
    return '(row)';
  }

  // ═══ Build ═══
  // Sites map first: every site the Sites file creates is a valid Site* for
  // Locations, and so is every site the existing exports show PickTrace has.
  // A Locations Site that is neither — nor in the Locations template's own
  // Site list — is a new site: with the Sites template loaded it gets a row
  // in the Sites file (no Sites export needed), and Locations is mapped again
  // so it counts as created rather than as something to build.
  function buildAll() {
    const out = {};
    const ex = existIndex();
    const known = ex.siteNames.length ? { site: ex.siteNames } : null;
    const namesOf = m => {
      const ci = m ? m.headers.findIndex(h => normLoose(h) === 'name') : -1;
      return ci >= 0 ? m.rows.map(r => r[ci]).filter(Boolean) : [];
    };
    const pend = list => (list.length ? { site: list } : null);
    let sites = mapDomain(DOMAIN_BY_KEY.sites, { ex: ex.sites });
    let locs = mapDomain(DOMAIN_BY_KEY.locations, { pending: pend(namesOf(sites)), known: known, ex: ex.locations });
    const extra = newSiteNames(locs, namesOf(sites), ex.siteNames);
    if (extra.length && tpls.sites) {
      sites = mapDomain(DOMAIN_BY_KEY.sites, { ex: ex.sites, extra: extra });
      locs = mapDomain(DOMAIN_BY_KEY.locations, { pending: pend(namesOf(sites)), known: known, ex: ex.locations });
    }
    DOMAINS.forEach(d => {
      const m = d.key === 'sites' ? sites : d.key === 'locations' ? locs : mapDomain(d, { ex: ex[d.key] });
      if (m) out[d.key] = m;
    });
    return out;
  }

  // Locations' Site values PickTrace doesn't have and no Sites file row
  // creates — case variants as one, first spelling kept.
  function newSiteNames(locs, inSites, existing) {
    if (!locs || !tpls.locations) return [];
    const ci = locs.headers.findIndex(h => normLoose(h) === 'site');
    if (ci < 0) return [];
    const have = new Set((dropdownFor(tpls.locations.dropdowns, 'Site') || []).concat(inSites, existing).map(locNorm));
    const out = [], seen = new Set();
    locs.rows.forEach(r => {
      const v = str(r[ci]), k = locNorm(v);
      if (!v || have.has(k) || seen.has(k)) return;
      seen.add(k);
      out.push(v);
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
    renderFiles(); renderPreview(); renderBuildBanner(); renderSpelling(); renderUnderAge(); renderSiteGaps(); renderReqSummary();
    renderSmartFix(); renderFill(); renderSiteSrc();
    renderSiteAddr(); renderCollisions(); renderPossible(); renderDates(); renderRemoved();
    renderList('existing', collect('already'));
    renderList('fills', collect('filledCols'));
    renderList('derived', collect('derived'));
    renderList('fuzzy', collect('fuzzy'));
    renderList('addrflag', collect('addrFlags'));
    renderList('phonebad', collect('phoneBad'));
    renderList('remap', collect('remapped').concat(collect('matched')));
    // The banner above it covers Crew, Crop, Site and H2A Contract; this table
    // is every other column the org has no options for yet.
    renderList('nooptions', collect('noOptions').filter(n => BUILD_COLUMNS.indexOf(normLoose(n.column)) < 0));
    renderList('autofill', collect('autofilled'));
    renderList('unmapped', collect('unmapped'));
    renderList('dropped', collect('dropped'));
    renderNav();
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

  // ─── Files ───
  // The upload slots stay open until something builds, then fold into one
  // line naming what's loaded per file type; "Change files" opens them again.
  function renderFiles() {
    const strip = $('dt-files-strip'), slots = $('dt-files');
    const open = filesOpen === null ? !(built && Object.keys(built).length) : filesOpen;
    slots.style.display = open ? '' : 'none';
    if (!DOMAINS.some(d => srcs[d.key] || tpls[d.key] || (exist[d.key] || []).length)) { strip.style.display = 'none'; strip.innerHTML = ''; return; }
    // Each file type: its upload's row count (or what it still needs), how
    // many rows its existing export holds, and ⇄ to swap the two when the
    // auto-placement got them backwards.
    const chip = d => {
      const s = !!srcs[d.key], t = !!tpls[d.key], e = (exist[d.key] || []).length;
      const b = built && built[d.key];
      const state = b ? 'is-ready' : s || t || e ? 'is-partial' : 'is-empty';
      const what = b ? plural(b.rows.length, 'row') + (s ? '' : ' (from Locations)')
        : s ? 'needs its template' : t ? 'needs its export' : e ? 'needs its export' : 'not loaded';
      return '<span class="dt-file-chip ' + state + '"><b>' + escHtml(d.label) + '</b> ' + escHtml(what) +
        (e ? ' <span class="dt-chip-exist" title="' + escHtml((exist[d.key] || []).map(x => x.fileName).join(', ')) + '">· ' +
          existCount(d.key).toLocaleString() + ' in PickTrace</span>' +
          (s && e === 1 ? '<button type="button" class="dt-chip-btn" data-swap="' + d.key + '" title="Swap: build from the existing export, and read the current upload as what PickTrace has">&#8644;</button>' : '') +
          '<button type="button" class="dt-chip-btn" data-unexist="' + d.key + '" title="Stop using the existing export">&times;</button>' : '') +
        '</span>';
    };
    // Locations that name a site the Sites file creates import only after it.
    const viaSites = collect('fromPending').length;
    strip.innerHTML = DOMAINS.map(chip).join('') +
      (viaSites ? '<span class="dt-files-note">Upload <b>Sites</b> to PickTrace before <b>Locations</b> &mdash; ' +
        plural(viaSites, 'site') + ' in Locations come from the Sites file.</span>' : '') +
      placeNotes.map(n => '<span class="dt-files-note">' + escHtml(n) + '</span>').join('') +
      '<button type="button" class="btn btn-ghost btn-sm" id="dt-files-toggle" aria-expanded="' + open + '">' +
      (open ? 'Fold file slots' : 'Change files') + '</button>';
    strip.style.display = '';
  }

  // ─── Preview ───
  // What each cell is: a blocker (required and empty), a value the template
  // won't take, and/or a change the tool made — with the reason, worked out
  // from the same records the Changed values section lists.
  function cellInfo(k) {
    const b = built[k], dom = DOMAIN_BY_KEY[k];
    const req = b.headers.map(h => /\*\s*$/.test(h));
    const off = b.headers.map(() => null), none = b.headers.map(() => null);
    const mark = (into, col, v) => {
      const ci = b.headers.indexOf(col);
      if (ci >= 0) (into[ci] = into[ci] || new Set()).add(String(v).toUpperCase());
    };
    (b.mismatches || []).forEach(m => m.values.forEach(v => mark(off, m.column, v)));
    (b.noOptions || []).forEach(n => mark(none, n.column, n.value));
    const dates = new Set((dom.dates || []).map(normLoose));
    const nums = new Set((dom.numeric || []).map(normLoose));
    const phones = new Set((dom.phones || []).map(normLoose));
    const fills = columnFills[k] || {}, sfx = smartFix[k] || {};
    const siteCi = b.headers.findIndex(h => normLoose(h) === 'site');
    const copied = new Set();
    (b.siteGaps || []).forEach(g => { if (g.applied) g.rows.forEach(ri => copied.add(ri)); });
    const origin = (ri, ci) => {
      const si = b.colSrc[ci];
      return si >= 0 ? str((b.effRows[b.rowSrc[ri]] || [])[si]) : '';
    };
    // Rows in a name collision (on their name columns), possible duplicates
    // (on First / Last Name), and year-only date columns.
    const coll = new Map();
    (b.collisions || []).forEach(c => c.rows.forEach(ri => coll.set(ri, c.keyCols)));
    const dup = new Set((b.possible || []).map(p => p.row).filter(ri => ri != null));
    const nameCols = new Set(b.headers.map((h, ci) => (/^(firstname|lastname)$/.test(normLoose(h)) ? ci : -1)).filter(ci => ci >= 0));
    const yearCols = new Set(b.headers.map((h, ci) => (dates.has(normLoose(h)) && normLoose(h) !== 'startdate' ? ci : -1)).filter(ci => ci >= 0));
    // A rewrite record for this cell: same column, same result, and a "from"
    // that is the source value as typed or as normalized.
    const record = (list, ci, was, now) => {
      const h = b.headers[ci], k2 = normLoose(h);
      const forms = [was, dates.has(k2) ? toYMD(was) : nums.has(k2) ? cleanNumber(was, h) : was, was === '' ? '(blank)' : was];
      return (list || []).find(x => x.column === h && x.to === now && forms.some(f => x.from.indexOf(f) >= 0));
    };
    return {
      // '' | 'req' (required, empty) | 'coll' (shares its name) | 'off' (not
      // in the dropdown) | 'build' (the org has no options for it) | 'dup'
      // (possible duplicate) | 'year' (a date that is only a year)
      problem(ri, ci) {
        const v = b.rows[ri][ci];
        if (req[ci] && v === '') return 'req';
        if (coll.has(ri) && coll.get(ri).indexOf(ci) >= 0) return 'coll';
        if (dup.has(ri) && nameCols.has(ci)) return 'dup';
        if (!v) return '';
        if (yearCols.has(ci) && isYearOnlyDate(v)) return 'year';
        const u = v.toUpperCase();
        return off[ci] && off[ci].has(u) ? 'off' : none[ci] && none[ci].has(u) ? 'build' : '';
      },
      // null, or { was, why } when the cell differs from the source. A change of
      // case only (false → FALSE, matching the dropdown) isn't worth marking.
      change(ri, ci) {
        const now = b.rows[ri][ci], was = origin(ri, ci);
        if (now === was || now.toUpperCase() === was.toUpperCase()) return null;
        const h = normLoose(b.headers[ci]);
        let why = b.colSrc[ci] < 0 ? 'Filled by the tool' : 'Rewritten by the tool';
        let r;
        const pv = b.prov && b.prov.get(b.rid[ri] + '|' + ci);
        if (pv === 'Your edit' && (cellEdits[k] || {})[b.rid[ri] + '|' + ci] === now) why = pv;
        else if (ci === siteCi && copied.has(ri)) why = 'Copied from the row above (you confirmed)';
        else if (fills[ci] && fills[ci].val === now) why = 'Column fill you applied';
        else if (sfx[ci + '||' + was.toUpperCase()] === now) why = 'Smart Fix you applied';
        else if (h === 'dateofbirth' && ageFixes[k + '|' + b.rowSrc[ri]] === now) why = 'Birth date you corrected';
        else if ((r = record(b.remapped, ci, was, now))) why = r.why === 'only option in this org' ? 'Only option in this org (migrated)' : r.why;
        else if ((r = record(b.matched, ci, was, now))) why = r.why === 'only matching dropdown entry' ? 'Only matching dropdown entry' : r.why;
        else if (pv) why = pv;
        else if (!was && (b.autofilled || []).some(a => a.column === b.headers[ci])) why = 'Only option in the dropdown, filled in';
        else if (dates.has(h)) why = 'Date written as YYYY-MM-DD';
        else if (nums.has(h)) why = 'Unit or formatting removed';
        else if (phones.has(h) || h === 'ssn') why = 'Reformatted for PickTrace';
        return { was, why };
      }
    };
  }

  const PROBLEM_TIP = {
    req: 'Required and empty — type it here, or fill the column in Required blanks.',
    coll: 'Another row has this name — PickTrace keeps only one. See Name collisions.',
    off: 'Not in the template\'s dropdown — pick a listed value here or in Not in dropdown.',
    build: 'Your org has no options for this column yet — see Build in PickTrace.',
    dup: 'Possibly the same person as another employee — see Possible duplicates.',
    year: 'Only a year — pick a month and day in Year-only dates.'
  };
  const CELL_CUSTOM = '__dt_custom__';

  // The values a preview cell can pick from: the template's dropdown plus
  // what this run creates or the org already has (Site lists new sites).
  function cellOptions(b, ci) {
    const h = b.headers[ci];
    const list = (dropdownFor(b.dropdowns, h) || []).slice();
    const up = new Set(list.map(o => o.toUpperCase()));
    [b.pending, b.known].forEach(p => ((p || {})[normLoose(h)] || []).forEach(o => {
      if (!up.has(o.toUpperCase())) { up.add(o.toUpperCase()); list.push(o); }
    }));
    return list;
  }

  function renderPreview() {
    const keys = DOMAINS.map(d => d.key).filter(k => built && built[k]);
    const sec = $('dt-section-preview');
    if (!keys.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    sec.classList.toggle('is-collapsed', pvHidden);
    if (keys.indexOf(previewKey) < 0) previewKey = keys[0];
    const infos = {};
    keys.forEach(k => { infos[k] = cellInfo(k); });
    const hasProblem = (k, ri) => built[k].headers.some((_, ci) => infos[k].problem(ri, ci));
    $('dt-preview-pills').innerHTML = keys.map(k => {
      let p = 0;
      built[k].rows.forEach((_, ri) => { if (hasProblem(k, ri)) p++; });
      return '<button type="button" class="btn btn-sm ts-domain-pill' + (k === previewKey ? ' ts-domain-active' : '') +
        '" data-key="' + k + '">' + escHtml(DOMAIN_BY_KEY[k].label) + ' <b>' + built[k].rows.length + '</b>' +
        (p ? ' <span class="dt-pill-warn" title="' + plural(p, 'row') + ' with a problem">⚠ ' + p + '</span>' : '') + '</button>';
    }).join('');

    // Every row, drawn as you scroll (IMWorkbench.rows), filtered by the shared
    // bar: all · needs a fix · changed · either, optionally only the columns
    // with something to see.
    const b = built[previewKey], I = infos[previewKey];
    const all = b.rows.map((_, ri) => ri);
    const fixRow = ri => hasProblem(previewKey, ri), chgRow = ri => b.headers.some((_, ci) => I.change(ri, ci));
    const counts = { all: all.length, fix: 0, chg: 0, any: 0 };
    const flags = all.map(ri => { const f = fixRow(ri), c = chgRow(ri); if (f) counts.fix++; if (c) counts.chg++; if (f || c) counts.any++; return [f, c]; });
    // "Show rows" in Required blanks narrows the grid to the rows missing
    // that column, until you clear it (× on the filter bar).
    if (pvFocus && (pvFocus.key !== previewKey || !b.headers[pvFocus.ci])) pvFocus = null;
    const shown = pvFocus ? all.filter(ri => b.rows[ri][pvFocus.ci] === '')
      : all.filter(ri => IMWorkbench.rowMatch(pvFilter, flags[ri][0], flags[ri][1]));
    let cols = b.headers.map((_, ci) => ci);
    if (pvColsOnly) cols = cols.filter(ci => shown.some(ri => I.problem(ri, ci) || I.change(ri, ci)));
    const head = '<thead><tr><th class="dt-pv-n">#</th>' + cols.map(ci => {
      const h = b.headers[ci];
      const cls = (/\*\s*$/.test(h) ? 'is-req' : '') + (b.colSrc[ci] < 0 ? ' is-nosrc' : '');
      return '<th' + (cls.trim() ? ' class="' + cls.trim() + '"' : '') +
        (b.colSrc[ci] < 0 ? ' title="No source column — filled by the tool or by you, or left empty"' : '') + '>' + escHtml(h) + '</th>';
    }).join('') + '</tr></thead>';
    // Every cell is editable: a dropdown column is a button that becomes its
    // <select> on click, anything else text that becomes an <input> — never
    // a live control per cell (IMWorkbench.pickButton / textCell). × on the
    // row number leaves the row out.
    const hasList = cols.map(ci => cellOptions(b, ci).length > 0);
    const rowHtml = ri => '<tr><td class="dt-pv-n">' + (ri + 1) +
      '<button type="button" class="dt-pv-rm" data-ri="' + ri + '" title="Leave this row out of the export" aria-label="Remove row ' + (ri + 1) + '">&times;</button></td>' +
      cols.map((ci, p) => {
        const v = b.rows[ri][ci];
        const pr = I.problem(ri, ci), c = I.change(ri, ci);
        const cls = 'ts-cell' + (pr ? ' is-' + pr : '') + (c ? ' is-chg' : '');
        const tip = [pr ? PROBLEM_TIP[pr] : '',
          c ? 'Was ' + (c.was === '' ? '(blank)' : '"' + c.was + '"') + ' → now ' + (v === '' ? '(blank)' : '"' + v + '"') + ' · ' + c.why : '']
          .filter(Boolean).join('\n');
        const attrs = ' data-ri="' + ri + '" data-ci="' + ci + '"';
        return '<td class="' + cls + '"' + (tip ? ' title="' + escHtml(tip) + '"' : '') + '>' +
          (hasList[p] ? IMWorkbench.pickButton(v, attrs) : IMWorkbench.textCell(v, attrs)) + '</td>';
      }).join('') + '</tr>';
    const table = $('dt-preview-table');
    IMWorkbench.rows(table, { head: head, ids: shown, row: rowHtml, cols: cols.length + 1,
      keys: [null].concat(cols.map(ci => previewKey + ':' + b.headers[ci])), cell: (ri, p) => b.rows[ri][cols[p - 1]],
      widths: [58].concat(cols.map(ci => IMWorkbench.fitWidth(b.headers[ci], shown.map(ri => b.rows[ri][ci]), 80, 240))),
      empty: pvFocus ? 'Every row has this column filled now.' : pvFilter === 'fix' ? 'No row in this file has a problem.' : 'The tool changed nothing in this file.' });
    IMWorkbench.wirePicks(table, ci => (built && built[previewKey] ? cellOptions(built[previewKey], ci) : []), CELL_CUSTOM);
    IMWorkbench.wireTextCells(table);
    $('dt-pv-filterbar').innerHTML = IMWorkbench.filterBar({ mode: pvFilter, counts: counts, colsOnly: pvColsOnly,
      shownCols: cols.length, totalCols: b.headers.length,
      focus: pvFocus ? 'Rows with ' + b.headers[pvFocus.ci].replace(/\*\s*$/, '') + ' empty (' + shown.length + ')' : '' });
    const dest = b.hasTemplate ? tpls[previewKey].fileName : 'a generated DATA ENTRY sheet';
    $('dt-preview-note').textContent =
      (shown.length !== b.rows.length ? shown.length + ' of ' + b.rows.length + ' rows shown · ' : '') +
      'all ' + plural(b.rows.length, 'row') + ' export into ' + dest;
    const t = $('dt-pv-toggle');
    t.textContent = pvHidden ? 'Show grid' : 'Hide grid';
    t.setAttribute('aria-expanded', String(!pvHidden));
  }

  // ─── Sections ───
  // One button per job, in the order you'd do them. Fix counts are what
  // PickTrace would reject or what needs your yes; Check counts are what the
  // tool did that's worth a look. Each key matches a .dt-panel in index.html.
  // `when` hides a button that can't apply (no file of that type built) and
  // has nothing in it.
  const has = k => !!(built && built[k]);
  const SECTIONS = [
    { key: 'required', group: 'fix', unit: 'column', name: 'Required blanks', desc: 'Fill required (*) columns the export left empty' },
    { key: 'dropdown', group: 'fix', unit: 'value', name: 'Not in dropdown', desc: 'Pick a listed value for anything PickTrace would reject' },
    { key: 'build', group: 'fix', unit: 'value', name: 'Build in PickTrace', desc: 'Values to create in PickTrace, then re-download the template' },
    { key: 'spelling', group: 'fix', unit: 'spelling', name: 'Spelling fixes', desc: 'Misspelled crops and varieties matched to the right name — undo any that are wrong',
      when: () => collect('spelled').length + collect('spellCheck').length > 0 },
    { key: 'siteaddr', group: 'fix', unit: 'site', name: 'Site addresses', desc: 'Street, city, state and zip for each site — for all, per ranch, or one by one',
      when: () => has('sites') },
    { key: 'fillsite', group: 'fix', unit: 'row', name: 'Blank Site', desc: 'Locations with no Site under one that has it — copy it down?',
      when: () => has('locations') },
    { key: 'collisions', group: 'fix', unit: 'row', name: 'Name collisions', desc: 'Rows sharing a name — PickTrace keeps only one',
      when: () => has('sites') || has('locations') || has('packstyles') || has('jobs') },
    { key: 'possible', group: 'fix', unit: 'employee', name: 'Possible duplicates', desc: 'Kept, but they match someone by name or ID — check each one',
      when: () => has('employees') },
    { key: 'dates', group: 'fix', unit: 'cell', name: 'Year-only dates', desc: 'Dates written as just a year — pick a month and day',
      when: () => collect('yearOnly').length > 0 },
    { key: 'underage', group: 'fix', unit: 'employee', name: 'Under age', desc: 'Birth dates that put an employee under ' + MIN_AGE,
      when: () => has('employees') },
    { key: 'changes', group: 'check', unit: 'change', name: 'Changed values', desc: 'What the tool rewrote, filled or split, and why' },
    { key: 'columns', group: 'check', unit: 'column', name: 'Column matching', desc: 'Guessed column pairs, and source columns with no home' },
    { key: 'contact', group: 'check', unit: 'entry', name: 'Addresses & phones', desc: 'Addresses with no city, phones PickTrace won\'t take' },
    { key: 'existing', group: 'check', unit: 'row', name: 'Already in PickTrace', desc: 'Rows the existing export already has — left out',
      when: () => DOMAINS.some(d => (exist[d.key] || []).length) },
    { key: 'dropped', group: 'check', unit: 'row', name: 'Dropped rows', desc: 'Archived, repeated, placeholder and removed rows left out' }
  ];
  const GROUP_LABEL = { fix: 'Fix before upload', check: 'Check' };

  // { count, where } per section — where names the file types it's in.
  function sectionCounts() {
    const tally = (items, n) => {
      const where = [];
      let count = 0;
      items.forEach(x => { count += n ? n(x) : 1; if (x.domain && where.indexOf(x.domain) < 0) where.push(x.domain); });
      return { count, where };
    };
    const bg = buildGroups();
    const build = { count: bg.totalVals + bg.other, where: [...bg.byDomain.keys()] };
    collect('noOptions').forEach(n => { if (build.where.indexOf(n.domain) < 0) build.where.push(n.domain); });
    return {
      required: tally(collect('required')),
      dropdown: tally(collect('mismatches'), m => m.values.length),
      build: build,
      spelling: tally(collect('spelled').concat(collect('spellCheck'))),
      siteaddr: { count: has('sites') ? built.sites.addrMissing : 0, where: has('sites') && built.sites.addrMissing ? ['Sites'] : [] },
      fillsite: tally(collect('siteGaps').filter(g => !g.applied), g => g.rows.length),
      collisions: tally(collect('collisions'), c => c.rows.length),
      possible: tally(collect('possible')),
      dates: tally(collect('yearOnly'), y => y.left),
      underage: tally(collect('underAge')),
      changes: tally(collect('remapped').concat(collect('matched'), collect('autofilled'), collect('filledCols'), collect('derived'))),
      columns: tally(collect('fuzzy').concat(collect('unmapped'))),
      contact: tally(collect('addrFlags').concat(collect('phoneBad'))),
      existing: tally(collect('already')),
      dropped: tally(collect('dropped').concat(collect('removedList')))
    };
  }

  function renderNav() {
    const nav = $('dt-nav'), panels = $('dt-results');
    if (dtSplit) dtSplit.show(!!(built && Object.keys(built).length));
    if (!built || !Object.keys(built).length) { nav.style.display = 'none'; panels.style.display = 'none'; return; }
    const c = sectionCounts();
    const shown = SECTIONS.filter(s => c[s.key].count || !s.when || s.when());
    // Until you pick one, the open section is the first with something waiting.
    if (!sectionPicked || !activeSection || !shown.some(s => s.key === activeSection)) {
      const first = shown.find(s => s.group === 'fix' && c[s.key].count) || shown.find(s => c[s.key].count);
      activeSection = (first || shown[0]).key;
    }
    nav.innerHTML = ['fix', 'check'].map(g =>
      '<div class="dt-nav-group"><div class="dt-nav-label">' + GROUP_LABEL[g] + '</div><div class="dt-nav-row">' +
      shown.filter(s => s.group === g).map(s => {
        const n = c[s.key].count, on = s.key === activeSection;
        return '<button type="button" class="dt-nav-btn is-' + g + (on ? ' is-on' : '') + (n ? '' : ' is-clear') +
          '" data-sec="' + s.key + '" aria-pressed="' + on + '">' +
          '<span class="dt-nav-top"><span class="dt-nav-name">' + escHtml(s.name) + '</span>' +
          '<span class="dt-nav-count' + (n ? '' : ' is-zero') + '">' + (n ? n : '✓ 0') + '</span></span>' +
          '<span class="dt-nav-desc">' + escHtml(s.desc) + '</span>' +
          (n ? '<span class="dt-nav-where">' + escHtml([plural(n, s.unit).replace(/entrys$/, 'entries')].concat(c[s.key].where).join(' · ')) + '</span>' : '') +
          '</button>';
      }).join('') + '</div></div>').join('');
    nav.style.display = '';
    panels.style.display = '';
    // Show the open panel; say so plainly when it holds nothing.
    let filled = false;
    panels.querySelectorAll('.dt-panel').forEach(p => {
      const on = p.dataset.panel === activeSection;
      p.style.display = on ? '' : 'none';
      if (on) filled = [...p.children].some(ch => ch.style.display !== 'none' && (ch.tagName === 'SECTION' || ch.innerHTML.trim()));
    });
    const sec = SECTIONS.find(s => s.key === activeSection);
    const empty = $('dt-panel-empty');
    empty.style.display = filled ? 'none' : '';
    empty.textContent = filled ? '' : '✓ Nothing in ' + sec.name + (sec.group === 'fix' ? ' — nothing here will stop the upload.' : ' — nothing to review.');
  }

  // Bring the section buttons into view: beside the preview when side by
  // side, just under the pinned preview when stacked.
  function revealSections() {
    if (dtSplit) { dtSplit.reveal($('dt-nav')); return; }
    const sc = $('page-compare'), nav = $('dt-nav'), pv = $('dt-section-preview');
    if (!sc || !nav || !sc.scrollTo) return;
    const top = sc.scrollTop + nav.getBoundingClientRect().top - sc.getBoundingClientRect().top - (pv ? pv.offsetHeight : 0) - 8;
    sc.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
  }

  // Required blanks in every file, one chip each — click one to show that
  // file in the preview and in Column Fill.
  function renderReqSummary() {
    const el = $('dt-reqsum');
    if (!el) return;
    const by = new Map();
    collect('required').forEach(q => { if (!by.has(q.domain)) by.set(q.domain, []); by.get(q.domain).push(q); });
    el.innerHTML = [...by.entries()].map(([domain, qs]) => {
      const d = DOMAINS.find(x => x.label === domain);
      return '<button type="button" class="dt-reqsum-chip' + (d && d.key === previewKey ? ' is-on' : '') + '" data-key="' +
        (d ? d.key : '') + '"><b>' + escHtml(domain) + '</b> ' +
        escHtml(qs.map(q => q.column + ' ' + q.empty + '/' + q.total + ' empty').join(' · ')) + '</button>';
    }).join('');
  }

  // Blank Site groups — offered, never applied unasked.
  function renderSiteGaps() {
    const sec = $('dt-section-sitegaps');
    if (!sec) return;
    const gaps = collect('siteGaps');
    if (!gaps.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const pending = gaps.filter(g => !g.applied);
    let rowsN = 0;
    pending.forEach(g => { rowsN += g.rows.length; });
    $('dt-sitegaps-title').textContent = 'Blank Site — ' +
      (pending.length ? plural(rowsN, 'row') + ' in ' + plural(pending.length, 'group') + ' waiting' : 'all copied down');
    $('dt-sitegaps-all').style.display = pending.length > 1 ? '' : 'none';
    const span = rs => rs.length === 1 ? 'Row ' + (rs[0] + 1) : 'Rows ' + (rs[0] + 1) + '–' + (rs[rs.length - 1] + 1);
    $('dt-sitegaps-table').innerHTML = '<thead><tr><th>Preview rows</th><th>Locations</th><th>Site to copy down</th><th></th></tr></thead><tbody>' +
      gaps.map(g => '<tr><td>' + span(g.rows) + ' <span class="text-muted small">(' + plural(g.rows.length, 'row') + ')</span></td>' +
        '<td>' + escHtml(g.names.join(', ') + (g.rows.length > g.names.length ? ', …' : '')) + '</td>' +
        '<td><b>' + escHtml(g.site) + '</b></td><td>' +
        (g.applied
          ? '<span class="dt-ok">✓ Copied</span> <button type="button" class="btn btn-ghost btn-sm" data-act="undo" data-key="' + g.key + '" data-id="' + escHtml(g.id) + '">Undo</button>'
          : '<button type="button" class="btn btn-primary btn-sm" data-act="apply" data-key="' + g.key + '" data-id="' + escHtml(g.id) + '">Copy down</button>') +
        '</td></tr>').join('') + '</tbody>';
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
  // buried the real blockers — it stays in the Needs Building table below.
  const BUILD_COLUMNS = ['h2acontract', 'crew', 'site', 'cropvariety', 'croptype'];
  // Rendered as crop → varieties instead of a flat list.
  const BUILD_NESTED = ['cropvariety'];

  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');

  // "Blackberries-Sample 1234-1" → crop "Blackberries", variety "Sample 1234-1".
  // Split on the FIRST hyphen: crop names here never contain one, but variety
  // names do, so splitting on the last would produce "Blackberries-Sample 1234".
  function splitCrop(v) {
    const i = String(v).indexOf('-');
    return i > 0
      ? { crop: v.slice(0, i).trim(), variety: v.slice(i + 1).trim() }
      : { crop: String(v).trim(), variety: '' };
  }

  // What has to exist in PickTrace before the files will import, grouped
  // domain → column → values. The banner draws it; its button counts it.
  function buildGroups() {
    const all = collect('noOptions');
    const inBuild = col => BUILD_COLUMNS.indexOf(normLoose(col)) >= 0;
    const list = all.filter(n => inBuild(n.column));
    // Off-list values in these columns are just as missing as values in an
    // empty list — the org has crews and crops, just not these ones. The
    // column's current list rides along so Crop & Variety can tell a new
    // crop apart from a new variety under a crop that already exists.
    const allowedBy = new Map();
    let offCount = 0;
    collect('mismatches').forEach(m => {
      if (!inBuild(m.column)) return;
      allowedBy.set(m.domain + '\u0000' + m.column, m.allowed);
      m.values.forEach((v, i) => { list.push({ domain: m.domain, column: m.column, value: v, count: m.counts[i] }); offCount++; });
    });

    // "crew 9" and "Crew 9" are one crew to create, so case variants merge
    // under the first spelling seen.
    const byDomain = new Map();
    let totalVals = 0;
    list.forEach(n => {
      if (!byDomain.has(n.domain)) byDomain.set(n.domain, new Map());
      const cols = byDomain.get(n.domain);
      if (!cols.has(n.column)) cols.set(n.column, []);
      const vs = cols.get(n.column);
      const same = vs.find(v => v.value.toUpperCase() === n.value.toUpperCase());
      if (same) { same.count += n.count; return; }
      vs.push({ value: n.value, count: n.count });
      totalVals++;
    });
    let totalCols = 0;
    byDomain.forEach(cols => { totalCols += cols.size; });
    const other = all.filter(n => !inBuild(n.column)).length;
    return { byDomain, totalVals, totalCols, other, offCount, allowedBy };
  }

  function renderBuildBanner() {
    const el = $('dt-build-banner');
    if (!el) return;
    const { byDomain, totalVals, totalCols, other, offCount, allowedBy } = buildGroups();
    if (!totalVals) { el.style.display = 'none'; el.innerHTML = ''; return; }
    const fixedSpell = collect('spelled').filter(s => !s.kept).length;

    let html = '<div class="dt-build-head">' +
      '<div class="dt-build-title">Build these in PickTrace first</div>' +
      '<div class="dt-build-sub"><b>' + plural(totalVals, 'value') + '</b> across <b>' +
      plural(totalCols, 'column') + '</b> aren\'t in the destination template\'s dropdown, so rows using them are ' +
      'rejected on import. Create them in PickTrace, then re-download the template' +
      (offCount ? ' — or, if one is only spelled differently, map it to the existing value in <b>Smart Fixes</b>' : '') +
      '. The files still download either way.' +
      (fixedSpell ? ' <button type="button" class="dt-spell-link" data-gosec="spelling">' + plural(fixedSpell, 'misspelling') +
        ' matched to the right name, so ' + (fixedSpell === 1 ? 'it isn’t' : 'they aren’t') + ' listed here — see Spelling fixes</button>' : '') +
      (other ?' <span class="dt-build-aside">' + plural(other, 'other value') +
        ' in optional columns are listed below.</span>' : '') +
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
        html += BUILD_NESTED.indexOf(normLoose(col)) >= 0
          ? nestedCard(col, vs, allowedBy.get(domain + '\u0000' + col) || [])
          : flatCard(col, vs);
      });
      html += '</div></section>';
    });

    el.innerHTML = html;
    el.style.display = '';
  }

  // ─── Spelling fixes ───
  // What spellCropVariety() rewrote, each with Undo (keeps it as typed — it
  // then shows up under Build in PickTrace), plus close calls it left alone,
  // each with a one-click "Use …" (a sticky Smart Fix).
  function renderSpelling() {
    const box = $('dt-spell-box');
    if (!box) return;
    const fixes = collect('spelled'), checks = collect('spellCheck');
    if (!fixes.length && !checks.length) { box.style.display = 'none'; box.innerHTML = ''; return; }
    const many = new Set(fixes.concat(checks).map(x => x.domain)).size > 1;
    // Same crop on both sides: show just the varieties.
    const parts = (from, to) => {
      const a = splitCrop(from), b = splitCrop(to);
      const sameCrop = a.crop.toUpperCase() === b.crop.toUpperCase() && a.variety && b.variety;
      return sameCrop ? { crop: b.crop, from: a.variety, to: b.variety } : { crop: '', from: from, to: to };
    };
    const data = (x, act, label, extra) => '<button type="button" class="dt-spell-btn' + (extra || '') + '" data-spell="' + act +
      '" data-key="' + escHtml(x.key) + '" data-ci="' + x.ci + '" data-from="' + escHtml(x.from) + '" data-to="' + escHtml(x.to) + '">' + label + '</button>';
    const done = fixes.filter(f => !f.kept), kept = fixes.filter(f => f.kept);
    let rowsFixed = 0;
    done.forEach(f => { rowsFixed += f.rows; });
    fixes.sort((a, b) => (a.kept ? 1 : 0) - (b.kept ? 1 : 0) || a.to.localeCompare(b.to));
    let html = '<div class="dt-spell-head"><div class="dt-spell-title">' +
      (done.length ? plural(done.length, 'spelling') + ' fixed for you · ' + plural(rowsFixed, 'row') : 'Spelling') + '</div>' +
      '<p class="dt-spell-sub">These looked like typos of a crop or variety the template already lists, a known grape variety, ' +
      'or the spelling most of your rows use — so they’re written that way instead of being created a second time in PickTrace. ' +
      'Undo any that are really a different variety; it then goes back to <b>Build in PickTrace</b>.</p></div>';
    if (fixes.length) {
      html += '<ul class="dt-spell-list">' + fixes.map(f => {
        const p = parts(f.from, f.to);
        return '<li class="dt-spell-row' + (f.kept ? ' is-kept' : '') + '">' +
          '<span class="dt-spell-crop">' + escHtml(p.crop || '') + (many ? ' <span class="dt-spell-dom">' + escHtml(f.domain) + '</span>' : '') + '</span>' +
          '<span class="dt-spell-pair"><span class="dt-spell-from">' + escHtml(p.from) + '</span>' +
          '<span class="dt-spell-arrow" aria-hidden="true">→</span><span class="dt-spell-to">' + escHtml(p.to) + '</span></span>' +
          '<span class="dt-spell-why">' + (f.kept ? 'Kept as typed — to create' : escHtml(f.why.charAt(0).toUpperCase() + f.why.slice(1))) + '</span>' +
          '<span class="dt-spell-rows">' + plural(f.rows, 'row') + '</span>' +
          (f.kept ? data(f, 'redo', 'Fix it') : data(f, 'undo', 'Undo', ' is-quiet')) + '</li>';
      }).join('') + '</ul>';
    }
    if (checks.length) {
      html += '<div class="dt-spell-check"><div class="dt-spell-check-head">Might be the same name — your call (' + checks.length + ')</div>' +
        '<p class="dt-spell-sub">Spelled almost alike, with about as many rows each, so the tool didn’t pick. Left as they are, both get created.</p>' +
        '<ul class="dt-spell-list">' + checks.map(c => {
          const p = parts(c.from, c.to);
          return '<li class="dt-spell-row is-check">' +
            '<span class="dt-spell-crop">' + escHtml(p.crop || '') + '</span>' +
            '<span class="dt-spell-pair"><span class="dt-spell-from">' + escHtml(p.from) + '</span>' +
            '<span class="dt-spell-arrow" aria-hidden="true">≈</span><span class="dt-spell-to">' + escHtml(p.to) + '</span></span>' +
            '<span class="dt-spell-why">' + plural(c.rows, 'row') + ' vs ' + plural(c.toRows, 'row') + '</span><span class="dt-spell-rows"></span>' +
            data(c, 'use', 'Use “' + escHtml(p.to) + '”') + '</li>';
        }).join('') + '</ul></div>';
    }
    box.innerHTML = html;
    box.style.display = '';
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

  // Crop & Variety, split into the two things you create in PickTrace: first
  // the crops the org doesn't have yet, then the varieties under each crop. A
  // crop already in the template's list (`allowed`) only needs its new
  // varieties. Crops order by total rows; varieties within a crop likewise.
  function nestedCard(col, vs, allowed) {
    const known = new Set(allowed.map(a => splitCrop(a).crop.toUpperCase()));
    const crops = new Map();
    vs.forEach(v => {
      const s = splitCrop(v.value);
      const k = s.crop.toUpperCase();
      if (!crops.has(k)) crops.set(k, { name: s.crop, rows: 0, vars: [], isNew: !known.has(k) });
      const g = crops.get(k);
      g.rows += v.count;
      g.vars.push({ name: s.variety || '(no variety)', count: v.count, full: v.value });
    });
    const ordered = [...crops.values()].sort((a, b) => b.rows - a.rows || a.name.localeCompare(b.name));
    const fresh = ordered.filter(g => g.isNew);
    const varieties = n => plural(n, 'variety').replace('varietys', 'varieties');

    const step1 = fresh.length
      ? '<ul class="dt-build-chips">' + fresh.map(g =>
          '<li><span class="dt-build-val">' + escHtml(g.name) + '</span>' +
          '<span class="dt-build-rows">' + plural(g.rows, 'row') + '</span></li>').join('') + '</ul>'
      : '<div class="dt-build-step-none">None — every crop already exists. Only varieties are missing.</div>';
    const step2 = '<div class="dt-build-crops">' +
      ordered.map(g => {
        g.vars.sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
        return '<div class="dt-build-crop">' +
          '<div class="dt-build-crop-head"><span class="dt-build-crop-name">' + escHtml(g.name) +
          (g.isNew ? ' <span class="dt-build-tag dt-build-tag-new">new crop</span>'
                   : ' <span class="dt-build-tag">crop exists</span>') + '</span>' +
          '<span class="dt-build-crop-n">' + g.vars.length + ' · ' + plural(g.rows, 'row') + '</span></div>' +
          '<ul class="dt-build-vals">' +
          g.vars.map(v => '<li title="' + escHtml(v.full) + '">' +
            '<span class="dt-build-val">' + escHtml(v.name) + '</span>' +
            '<span class="dt-build-rows">' + plural(v.count, 'row') + '</span></li>').join('') +
          '</ul></div>';
      }).join('') + '</div>';

    return '<div class="dt-build-col dt-build-col-wide">' +
      '<div class="dt-build-col-head"><span class="dt-build-col-name">' + escHtml(col) + '</span>' +
      '<span class="dt-build-col-n">' + plural(fresh.length, 'new crop') + ' · ' + varieties(vs.length) + '</span></div>' +
      '<div class="dt-build-step"><div class="dt-build-step-head"><span class="dt-build-step-n">1</span>' +
      'Crops to create (' + fresh.length + ')</div>' + step1 + '</div>' +
      '<div class="dt-build-step"><div class="dt-build-step-head"><span class="dt-build-step-n">2</span>' +
      'Varieties to create under each crop (' + vs.length + ')</div>' + step2 + '</div>' +
      '</div>';
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
    const sure = confidentMatch(value, allowed);
    if (sure) return sure;
    const v = normLoose(value);
    let best = '', score = 0;
    allowed.forEach(a => { const s = similarity(v, normLoose(a)); if (s > score) { score = s; best = a; } });
    return score >= 0.6 ? best : '';
  }
  // The one entry that differs only in spacing, punctuation or a plural
  // ("Almonds -Other" → "Almond-Other"), as Template Standardize matches.
  // '' when there's none, or more than one.
  function confidentMatch(value, allowed) {
    const plu = s => String(s == null ? '' : s).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean)
      .map(w => (w.length > 3 && w.endsWith('s') ? w.slice(0, -1) : w)).join('|');
    for (const key of [normLoose, plu]) {
      const k = key(value);
      if (!k) continue;
      const hits = allowed.filter(a => key(a) === k);
      if (hits.length === 1) return hits[0];
    }
    return '';
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
      '<th>Write instead</th><th>Match</th><th></th></tr></thead><tbody>';
    // Each picker is a button until clicked (IMWorkbench.wirePicks) — a full
    // <select> per value listed the column's whole dropdown again and again.
    // A sure match (spacing, punctuation, plural) is "suggested" and goes with
    // Apply all; a closer-looking guess is shown but waits for your Apply.
    const PICK = { kind: 'fix', cls: 'input-field dt-sf-pick', blank: '— leave as-is —', sort: false };
    const items = [];
    list.forEach(m => m.values.forEach((v, vi) => {
      const sure = confidentMatch(v, m.allowed);
      items.push({ m: m, v: v, n: m.counts ? m.counts[vi] : 0, sure: sure, sug: sure || suggestFor(v, m.allowed) });
    }));
    items.sort((a, b) => (b.sure ? 1 : 0) - (a.sure ? 1 : 0) || (b.sug ? 1 : 0) - (a.sug ? 1 : 0) || b.n - a.n);
    items.forEach(x => {
      const m = x.m;
      html += '<tr><td>' + escHtml(m.domain) + '</td><td>' + escHtml(m.column) + '</td>' +
        '<td class="cmp-warn">' + escHtml(x.v) + '</td><td>' + (x.n || '') + '</td>' +
        '<td>' + IMWorkbench.pickButton(x.sug, ' data-key="' + escHtml(m.key) + '" data-ci="' + m.ci +
          '" data-from="' + escHtml(x.v) + '" data-sure="' + (x.sure ? 1 : 0) + '" style="min-width:190px;"', PICK) + '</td>' +
        '<td>' + (x.sure ? '<span class="dt-ok">suggested</span>' : x.sug ? '<span class="cmp-warn">closest guess &mdash; check</span>'
          : '<span class="text-muted">needs a choice</span>') + '</td>' +
        '<td><button class="btn btn-primary btn-sm dt-sf-apply">Apply</button></td></tr>';
    });
    $('dt-smartfix-table').innerHTML = html + '</tbody>';
    IMWorkbench.wirePicks($('dt-smartfix-table'), (ci, b) => {
      const m = (built[b.dataset.key] && built[b.dataset.key].mismatches || []).find(x => x.ci === ci);
      return m ? m.allowed : [];
    }, PICK);
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
    // Only required (*) columns are listed up front. Optional ones stay folded
    // under one toggle row until you open it — or type in the filter, which is
    // a request for a specific column.
    const reqShown = shown.filter(c => c.req);
    const optShown = shown.filter(c => !c.req);
    const optOpen = fillOptOpen || !!fillFilter;
    const need = reqShown.filter(c => c.empty > 0);
    // Fully filled, but with values the dropdown won't take — here so they
    // can be bulk-set or cleared.
    const offl = reqShown.filter(c => c.empty === 0 && bad.has(c.h));
    const rest = reqShown.filter(c => c.empty === 0 && !bad.has(c.h));
    const reqEmpty = cols.filter(c => c.req && c.empty > 0).length;
    const visible = reqShown.length + (optOpen ? optShown.length : 0);

    $('dt-fill-title').textContent = 'Column Fill — ' + DOMAIN_BY_KEY[previewKey].label +
      ' · ' + visible + ' of ' + cols.length + ' columns' +
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
            : '<span class="cmp-warn">' + c.empty + ' empty</span>' +
              ' <button type="button" class="dt-fill-show" title="Show only the rows missing this column in the preview">Show rows</button>'));
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
    if (!reqShown.length) {
      html += '<tbody><tr><td colspan="8" class="dt-fill-band">' +
        (fillScope === 'attn' && !fillFilter
          ? (optShown.length ? 'No required column needs attention.'
                             : 'Nothing needs attention — switch to All columns to bulk-set anything else.')
          : (optShown.length ? 'No required column matches this filter.' : 'No column matches this filter.')) +
        '</td></tr></tbody>';
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
        (fillScope === 'attn' ? 'Required — filled here' : need.length || offl.length ? 'Other required columns' : 'Required columns') +
        ' (' + rest.length + ')</td></tr>' + rest.map(rowFor).join('') + '</tbody>';
    }
    // The fold row says what's inside even while closed, so an optional
    // column the dropdown rejects (Crew, Location Group) isn't out of sight.
    if (optShown.length) {
      const optBad = optShown.filter(c => bad.has(c.h));
      const optFilled = optShown.filter(c => fills[c.ci]).length;
      const notes = [];
      if (optBad.length) notes.push('<span class="dt-bad">⚠ ' + plural(optBad.length, 'column') + ' with values not in the dropdown</span>');
      if (optFilled) notes.push(plural(optFilled, 'column') + ' filled here');
      html += '<tbody><tr><td colspan="8" class="dt-fill-band dt-fill-opt">' +
        '<button type="button" class="dt-fill-opt-btn" aria-expanded="' + optOpen + '"' + (fillFilter ? ' disabled' : '') + '>' +
        '<span class="dt-fill-opt-caret">' + (optOpen ? '▾' : '▸') + '</span>' +
        'Optional columns (' + optShown.length + ')' +
        (notes.length ? '<span class="dt-fill-opt-note">· ' + notes.join(' · ') + '</span>' : '') +
        '<span class="dt-fill-opt-hint">' + (fillFilter ? 'shown while filtering' : optOpen ? 'click to hide' : 'click to edit') + '</span>' +
        '</button></td></tr>' +
        (optOpen ? optBad.concat(optShown.filter(c => !bad.has(c.h))).map(rowFor).join('') : '') +
        '</tbody>';
    }
    $('dt-fill-table').innerHTML = html;
  }
  let fillFilter = '';    // Column Fill search box, lowercased
  let fillScope = 'attn'; // attn | all | req | empty | filled
  let fillOptOpen = false; // optional (no *) columns unfolded in Column Fill

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
    nooptions: { title: 'Optional Columns With No Options Yet', cols: ['Domain', 'Column', 'Value to create', 'Rows'],
      row: n => [n.domain, n.column, n.value, n.count] },
    autofill: { title: 'Auto-Filled Columns', cols: ['Domain', 'Column', 'Value', 'Rows'],
      row: a => [a.domain, a.column, a.value, a.rows + ' of ' + a.total] },
    phonebad: { title: 'Phone Numbers PickTrace Will Reject', cols: ['Domain', 'Column', 'Value'],
      row: p => [p.domain, p.column, p.value] },
    unmapped: { title: 'Unmapped Source Columns', cols: ['Domain', 'Column', 'Sample values'],
      row: u => [u.domain, u.column, u.sample.join(' · ')] },
    dropped: { title: 'Dropped Rows', cols: ['Domain', 'Reason', 'Row'], row: d => [d.domain, d.reason, d.detail] },
    existing: { title: 'Already in PickTrace — left out', cols: ['Domain', 'Row', 'Matched on'], row: d => [d.domain, d.detail, d.by] }
  };

  // ─── Site addresses ───
  // Every site / per ranch / each site, as in Template Standardize's Sites
  // file. Typed as "street, city, ST zip" and split on build. A site whose
  // export carries a street keeps it unless you type one for that site.
  function renderSiteAddr() {
    const sec = $('dt-section-siteaddr');
    if (!sec) return;
    const b = built && built.sites;
    if (!b || !b.siteAddrRows.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    const list = b.siteAddrRows;
    const groups = new Map();
    list.forEach(s => groups.set(s.ranch, (groups.get(s.ranch) || 0) + 1));
    const checked = text => {
      if (!text) return '';
      const p = parseSiteAddr(text);
      if (p.address1 && p.city && p.state && p.zip) {
        return '<span class="ts-addr-ok" title="Address1: ' + escHtml(p.address1) + '">✓ ' + escHtml(p.city + ', ' + p.state + ' ' + p.zip) + '</span>';
      }
      const miss = ['address1', 'city', 'state', 'zip'].filter(k => !p[k]).map(k => (k === 'address1' ? 'street' : k));
      return '<span class="ts-addr-bad">⚠ no ' + miss.join(', ') + ' found</span>';
    };
    const field = (scope, key, val, ph) => '<input type="text" class="input-field ts-addr-in dt-addr-in" data-scope="' + scope + '"' +
      (key != null ? ' data-key="' + escHtml(key) + '"' : '') + ' value="' + escHtml(val) + '" placeholder="' + escHtml(ph) + '">';
    const fromExport = list.filter(s => s.from === 'export').length;
    $('dt-siteaddr-title').textContent = 'Site addresses — ' + (b.addrMissing ? plural(b.addrMissing, 'site') + ' still missing part of one' : 'every site has one');
    let html = '<p class="cmp-slot-hint">Type an address as <b>street, city, ST zip</b> &mdash; it&rsquo;s split into Address1, City, State, Zip and Country. ' +
      'The most specific one wins: a site&rsquo;s own address, then its ranch&rsquo;s, then the one for every site.' +
      (fromExport ? ' ' + plural(fromExport, 'site') + ' already have a street from the export, so only an address typed for that site changes ' + (fromExport === 1 ? 'it' : 'them') + '.' : '') +
      (b.newSites ? ' ' + plural(b.newSites, 'site') + ' came from Locations &mdash; Site values PickTrace doesn&rsquo;t have yet.' : '') + '</p>';
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Every site (' + list.length + ')</div>' +
      '<div class="ts-addr-row">' + field('all', null, siteAddr.all, '123 Main St, City, ST 12345') + checked(siteAddr.all) + '</div></div>';
    // Per ranch only when the names actually group (several sites to a ranch).
    if (groups.size < list.length) {
      html += '<div class="ts-addr-block"><div class="ts-addr-label">Per ranch &mdash; the name before &ldquo; - &rdquo; (' + groups.size + ')</div>' +
        '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Ranch</th><th>Sites</th><th>Address</th></tr></thead><tbody>';
      [...groups.keys()].sort((x, y) => x.localeCompare(y)).forEach(g => {
        const v = siteAddr.ranch[g] || '';
        html += '<tr><td>' + escHtml(g) + '</td><td>' + groups.get(g) + '</td><td><div class="ts-addr-row">' +
          field('ranch', g, v, siteAddr.all ? siteAddr.all + ' (every site)' : 'street, city, ST zip') + checked(v) + '</div></td></tr>';
      });
      html += '</tbody></table></div></div>';
    }
    html += '<div class="ts-addr-block"><div class="ts-addr-label">Each site (' + list.length + ')</div>' +
      '<div class="table-wrap"><table class="data-table ts-addr-table"><thead><tr><th>Site</th><th>Address</th></tr></thead><tbody>';
    const nameCi = b.headers.findIndex(h => normLoose(h) === 'name');
    const partCi = ['address1', 'city', 'state', 'zip'].map(k => b.headers.findIndex(h => normLoose(h) === k));
    list.slice().sort((x, y) => String(x.name).localeCompare(String(y.name))).forEach(s => {
      const ri = b.rid.indexOf(s.rid);
      const cur = ri >= 0 ? partCi.map(ci => (ci >= 0 ? b.rows[ri][ci] : '')) : [];
      const now = cur[0] ? cur[0] + ', ' + [cur[1], [cur[2], cur[3]].filter(Boolean).join(' ')].filter(Boolean).join(', ') : '';
      const ph = s.from === 'export' ? now + ' (from the export)' : s.from === 'ranch' || s.from === 'all' ? now + ' (inherited)' : 'street, city, ST zip';
      html += '<tr><td>' + escHtml(ri >= 0 && nameCi >= 0 ? b.rows[ri][nameCi] : s.name) + '</td><td><div class="ts-addr-row">' +
        field('site', s.rid, s.own, ph) + (s.own ? checked(s.own) : s.from && now ? checked(now) : '') + '</div></td></tr>';
    });
    $('dt-siteaddr-body').innerHTML = html + '</tbody></table></div></div>';
  }

  // ─── Name collisions ───
  // Rows sharing what PickTrace treats as the name — one of them would be
  // dropped on upload. Rename here (an edit to that cell), or leave one out.
  function renderCollisions() {
    const sec = $('dt-section-collisions');
    if (!sec) return;
    const list = collect('collisions');
    if (!list.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    let rowsN = 0;
    list.forEach(c => { rowsN += c.rows.length; });
    $('dt-collisions-title').textContent = 'Name collisions — ' + plural(rowsN, 'row') + ' in ' + plural(list.length, 'group');
    $('dt-collisions-all').style.display = list.some(c => c.rows.length > 1) ? '' : 'none';
    let html = '<thead><tr><th>File</th><th>Shared</th><th>#</th><th>Rename to</th><th></th><th>Differs in</th></tr></thead>';
    list.forEach(c => {
      const b = built[c.key];
      const nameCi = c.keyCols[c.keyCols.length - 1];
      // Columns where the rows of this group disagree — what tells them apart.
      const diff = b.headers.map((h, ci) => {
        const vals = new Set(c.rows.map(ri => str(b.rows[ri][ci]).toUpperCase()).filter(Boolean));
        return vals.size > 1 ? h.replace(/\*\s*$/, '') : '';
      }).filter(Boolean);
      html += '<tbody class="dt-coll-group">';
      c.rows.forEach((ri, n) => {
        const r = b.rows[ri];
        const sample = diff.slice(0, 3).map(h => {
          const ci = b.headers.findIndex(x => x.replace(/\*\s*$/, '') === h);
          return h + ': ' + (r[ci] || '(blank)');
        }).join(' · ');
        html += '<tr>' + (n === 0
          ? '<td rowspan="' + c.rows.length + '">' + escHtml(c.domain) + '</td><td rowspan="' + c.rows.length + '"><b>' + escHtml(c.value) + '</b>' +
            (c.taken ? '<div class="text-muted small">PickTrace already has a block with this ' + escHtml(c.label) +
              (c.taken.filter(Boolean).length ? ' (' + escHtml(c.taken.filter(Boolean).slice(0, 3).join(', ')) + ')' : '') + '</div>' : '') + '</td>'
          : '') +
          '<td>' + (ri + 1) + '</td>' +
          '<td><input type="text" class="input-field dt-coll-name" data-key="' + c.key + '" data-ri="' + ri + '" data-ci="' + nameCi +
            '" value="' + escHtml(r[nameCi]) + '" style="width:200px;"></td>' +
          '<td><button type="button" class="btn btn-ghost btn-sm dt-coll-drop" data-key="' + c.key + '" data-ri="' + ri + '">Leave out</button></td>' +
          '<td class="text-muted small">' + escHtml(sample || (c.taken && c.rows.length === 1 ? 'the existing block' : '')) + '</td></tr>';
      });
      html += '</tbody>';
    });
    $('dt-collisions-table').innerHTML = html;
  }

  // ─── Possible duplicates (Employees) ───
  function renderPossible() {
    const sec = $('dt-section-possible');
    if (!sec) return;
    const list = collect('possible');
    const kept = Object.keys(dupOk).reduce((n, k) => n + dupOk[k].size, 0);
    if (!list.length && !kept) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('dt-possible-title').textContent = list.length ? 'Possible duplicates (' + list.length + ')' : 'Possible duplicates — all checked';
    $('dt-possible-reset').style.display = kept ? '' : 'none';
    $('dt-possible-table').innerHTML = list.length
      ? '<thead><tr><th>#</th><th>Employee</th><th>Why</th><th>Matches</th><th></th></tr></thead><tbody>' +
        list.map(p => '<tr><td>' + (p.row + 1) + '</td><td>' + escHtml(p.name) + '</td><td class="cmp-warn">' + escHtml(p.reason) + '</td>' +
          '<td>' + escHtml(p.other) + '</td><td class="dt-nowrap">' +
          '<button type="button" class="btn btn-primary btn-sm" data-act="keep" data-key="' + p.key + '" data-rid="' + escHtml(p.rid) + '">Keep</button> ' +
          '<button type="button" class="btn btn-ghost btn-sm" data-act="drop" data-key="' + p.key + '" data-rid="' + escHtml(p.rid) + '">Leave out</button></td></tr>').join('') + '</tbody>'
      : '<tbody><tr><td>' + plural(kept, 'employee') + ' you chose to keep. <b>Review again</b> lists them once more.</td></tr></tbody>';
  }

  // ─── Year-only dates ───
  function renderDates() {
    const sec = $('dt-section-dates');
    if (!sec) return;
    const list = collect('yearOnly');
    if (!list.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    let left = 0, filled = 0;
    list.forEach(y => { left += y.left; filled += y.filled; });
    $('dt-dates-title').textContent = 'Year-only dates — ' + (left ? plural(left, 'cell') + ' waiting' : 'all ' + filled + ' given a month and day');
    $('dt-dates-mm').value = yearFill.mm;
    $('dt-dates-dd').value = yearFill.dd;
    $('dt-dates-range').value = yearFill.range;
    $('dt-dates-table').innerHTML = '<thead><tr><th>File</th><th>Column</th><th>Year only</th><th>Filled</th></tr></thead><tbody>' +
      list.map(y => '<tr><td>' + escHtml(y.domain) + '</td><td>' + escHtml(y.column) + '</td><td' + (y.left ? ' class="cmp-warn"' : '') + '>' + y.left +
        '</td><td>' + y.filled + '</td></tr>').join('') + '</tbody>';
  }

  // ─── Rows you left out ───
  function renderRemoved() {
    const sec = $('dt-section-removed');
    if (!sec) return;
    const list = collect('removedList');
    if (!list.length) { sec.style.display = 'none'; return; }
    sec.style.display = '';
    $('dt-removed-title').textContent = 'Left out by you (' + list.length + ')';
    $('dt-removed-table').innerHTML = '<thead><tr><th>File</th><th>Row</th><th></th></tr></thead><tbody>' +
      list.map(x => '<tr><td>' + escHtml(x.domain) + '</td><td>' + escHtml(x.detail) + '</td><td>' +
        '<button type="button" class="btn btn-ghost btn-sm" data-key="' + x.key + '" data-rid="' + escHtml(x.rid) + '">Restore</button></td></tr>').join('') +
      '</tbody>';
  }

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
  // A domain with a template is written by patching the template's own zip
  // (IMXlsxTemplate.build): only DATA ENTRY's rows change, so the dropdown
  // validations, styles and DROP-DOWN INPUTS arrive exactly as PickTrace
  // shipped them — a SheetJS round-trip drops the validations. Jobs have no
  // template, so theirs is a plain generated sheet. → a Blob, or null when
  // there's nothing to write.
  async function fileFor(key) {
    const b = built && built[key];
    if (!b || !b.rows.length) return null;
    if (b.hasTemplate) {
      return IMXlsxTemplate.build({ rawBuffer: tpls[key].rawBuffer, headers: b.headers, rows: b.rows });
    }
    const numeric = new Set((DOMAIN_BY_KEY[key].numeric || []).map(normLoose));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([b.headers]);
    b.rows.forEach((row, ri) => {
      row.forEach((val, c) => {
        if (val == null || val === '') return;
        // Numbers stay numbers where PickTrace expects them; everything else
        // is text so leading-zero IDs survive.
        ws[XLSX.utils.encode_cell({ r: ri + 1, c: c })] = numeric.has(normLoose(b.headers[c])) && /^-?\d+(\.\d+)?$/.test(String(val))
          ? { v: parseFloat(val), t: 'n' } : { v: String(val), t: 's' };
      });
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: b.rows.length, c: b.headers.length - 1 } });
    XLSX.utils.book_append_sheet(wb, ws, 'DATA ENTRY');
    return new Blob([XLSX.write(wb, { type: 'array', bookType: 'xlsx', bookSST: true })],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // "Acme Farms - Pack Styles - 2026-10-07.xlsx" (see IMOutName in utils.js).
  function outName(key) { return IMOutName.fileName(DOMAIN_BY_KEY[key].label); }
  // Upload order: Sites before Locations, whose Site* values they create.
  const EXPORT_ORDER = ['sites', 'locations', 'employees', 'packstyles', 'jobs'];

  function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  async function download(key) {
    try {
      const blob = await fileFor(key);
      if (blob) saveBlob(blob, outName(key));
    } catch (err) {
      alert(DOMAIN_BY_KEY[key].label + ' could not be written: ' + (err && err.message ? err.message : err));
    }
  }

  // Every built file in one .zip, inside a folder of the same name —
  // "Acme Farms - PickTrace Uploads - 2026-10-07/Acme Farms - Sites - …xlsx".
  async function downloadAll() {
    const keys = EXPORT_ORDER.filter(k => built && built[k] && built[k].rows.length);
    const folder = IMOutName.folderName();
    const files = [];
    for (const k of keys) {
      try {
        const blob = await fileFor(k);
        if (blob) files.push({ name: folder + '/' + outName(k), data: new Uint8Array(await blob.arrayBuffer()) });
      } catch (err) {
        alert(DOMAIN_BY_KEY[k].label + ' could not be written, so nothing was downloaded: ' + (err && err.message ? err.message : err));
        return;
      }
    }
    if (files.length) saveBlob(IMXlsxTemplate.bundle(files), folder + '.zip');
  }

  function stash(obj, key) { if (!obj[key]) obj[key] = {}; return obj[key]; }
  function applySmartFix(key, ci, from, to) {
    stash(smartFix, key)[ci + '||' + String(from).toUpperCase().trim()] = to;
    build();
  }
  // A cell you edited (preview or Name collisions), keyed to its row's rid.
  function setEdit(key, ri, ci, val) {
    const b = built && built[key];
    if (!b || !b.rows[ri] || b.rows[ri][ci] === val) return;
    stash(cellEdits, key)[b.rid[ri] + '|' + ci] = val;
    build();
  }
  function leaveOut(key, ris, noBuild) {
    const b = built && built[key];
    if (!b) return;
    const set = removed[key] || (removed[key] = new Set());
    ris.forEach(ri => { if (b.rid[ri] != null) set.add(b.rid[ri]); });
    if (!noBuild) build();
  }
  function setFill(ci, val, mode) {
    if (!previewKey) return;
    stash(columnFills, previewKey)[Number(ci)] = { val: val, mode: mode };
    build();
  }

  // Pre-export confirm, mirroring sites-standardize.js. Off-list values are
  // written exactly as they are, so this is the last place to catch them.
  function exportBlockers() {
    const out = { off: [], req: [], build: [], coll: [] };
    Object.keys(built || {}).forEach(k => {
      const b = built[k];
      if (!b.rows.length) return;
      out.off.push(...b.mismatches);
      out.req.push(...b.required);
      out.build.push(...b.noOptions);
      out.coll.push(...b.collisions);
    });
    return out;
  }
  function hasBlockers(x) { return !!(x.off.length || x.req.length || x.build.length || x.coll.length); }

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
      if (x.coll.length) {
        sections.push('<div class="ts-confirm-issue"><div class="ts-confirm-issue-head">⚠ ' + x.coll.length +
          ' name collision(s)</div><ul class="ts-confirm-list">' +
          x.coll.slice(0, 8).map(c => '<li><b>' + escHtml(c.domain) + '</b>: ' + escHtml(c.value) +
            (c.rows.length > 1 ? ' — ' + c.rows.length + ' rows' : ' — already in PickTrace') + '</li>').join('') +
          '</ul><div class="ts-confirm-issue-body">PickTrace keeps one record per name and silently drops the rest. Rename or leave one out in <b>Name collisions</b>.</div></div>');
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
      setSource(key, res.data, res.sheetLabel);
    }
    return true;
  }

  // The file a domain is built from.
  function setSource(key, data, sheetLabel) {
    if (srcs[key] !== data) forgetRows(key);
    data.sheetLabel = sheetLabel || data.sheetLabel || '';
    srcs[key] = data;
    IMOutName.suggestClient(data.fileName);
    // A source-column choice only survives an export that still has that column.
    const ov = srcOverride[key] || {};
    const have = new Set(data.headers.map(normLoose));
    Object.keys(ov).forEach(k => { if (!have.has(normLoose(ov[k]))) delete ov[k]; });
    slotLabel(key, 'src', data.fileName + (data.sheetLabel ? ' › ' + data.sheetLabel : '') + ' · ' + data.rows.length + ' rows');
  }
  function clearSource(key) {
    delete srcs[key];
    forgetRows(key);
    slotLabel(key, 'src', 'No export');
  }
  // A PickTrace export of what the destination org already has. Several per
  // domain are read as one; loading the same file and sheet again replaces it.
  function addExisting(key, data, sheetLabel) {
    data.sheetLabel = sheetLabel || data.sheetLabel || '';
    const list = exist[key] || (exist[key] = []);
    const same = list.findIndex(d => d.fileName === data.fileName && d.sheetName === data.sheetName);
    if (same >= 0) list[same] = data; else list.push(data);
    existLabel(key);
  }
  function existLabel(key) {
    const list = exist[key] || [];
    slotLabel(key, 'exist', list.length
      ? list.map(d => d.fileName + (d.sheetLabel ? ' › ' + d.sheetLabel : '')).join(', ') + ' · ' + existCount(key) + ' rows'
      : 'None');
  }
  // Edits, removals, kept duplicates and per-site addresses are keyed to the
  // source's rows, so they go when the source does. A site added from
  // Locations keeps its address — its key is its name.
  function forgetRows(key) {
    delete cellEdits[key]; delete removed[key]; delete dupOk[key];
    if (key === 'locations') delete siteFillOk[key];
    if (key === 'sites') Object.keys(siteAddr.site).forEach(k => { if (k.indexOf('new:') !== 0) delete siteAddr.site[k]; });
    if (pvFocus && pvFocus.key === key) pvFocus = null;
  }

  // ─── Auto-placement: the upload vs what PickTrace already has ───
  // A bulk upload can carry, for one file type, both the data to upload and
  // a PickTrace export of what the destination org already has. Per file
  // type:
  //   - one file: it's the upload;
  //   - a client's own sheet (an Implementation Data Template tab, header
  //     below row 1) is the upload, and every PickTrace-shaped export beside
  //     it is what PickTrace already has;
  //   - two PickTrace exports: the one whose Sites, Employers, Crews, Crops…
  //     the destination template lists is what PickTrace has; the other is
  //     the upload. When that's not clear-cut you're asked.
  // A file loaded earlier takes part, so an existing export dropped in later
  // lands beside the upload already there. Swap on the file strip undoes a
  // wrong call.
  const isPickTraceShape = data => (data.headerRow || 1) === 1;
  // The destination org's own lists, compared value by value. Lists that
  // every org shares (Gender, Language, TRUE/FALSE) say nothing about whose
  // export it is, so only these count.
  const OWN_LISTS = ['site', 'employer', 'crew', 'cropvariety', 'croptype', 'sitetype', 'locationgroup', 'h2acontract', 'title', 'packtype'];
  // Share of an export's distinct values the destination template lists —
  // null when there's nothing to compare against.
  function fitScore(key, data) {
    const lists = [];
    const tpl = tpls[key];
    data.headers.forEach((h, ci) => {
      const k = normLoose(h);
      if (!tpl || OWN_LISTS.indexOf(k) < 0) return;
      const l = dropdownFor(tpl.dropdowns, h);
      if (l && l.length) lists.push([ci, l]);
    });
    if (tpls.locations && (key === 'sites' || key === 'locations')) {
      const hk = data.headers.map(normLoose);
      const ci = key === 'sites' ? ['site', 'sitename', 'name'].map(k => hk.indexOf(k)).find(i => i >= 0) : hk.indexOf('site');
      const l = dropdownFor(tpls.locations.dropdowns, 'Site');
      if (ci != null && ci >= 0 && l && l.length && !lists.some(x => x[0] === ci)) lists.push([ci, l]);
    }
    if (!lists.length) return null;
    let hit = 0, tot = 0;
    lists.forEach(([ci, l]) => {
      const set = new Set(l.map(v => v.toUpperCase())), seen = new Set();
      data.rows.forEach(r => {
        const v = str(r[ci]).toUpperCase();
        if (!v || seen.has(v)) return;
        seen.add(v);
        tot++;
        if (set.has(v)) hit++;
      });
    });
    return tot ? hit / tot : null;
  }
  // Two PickTrace exports where one clearly fits the destination org and
  // the other clearly doesn't → { src, ex }; otherwise null (ask).
  function guessRoles(key, cands) {
    if (cands.length !== 2) return null;
    const s = cands.map(c => fitScore(key, c.data));
    if (s[0] == null || s[1] == null || Math.max(s[0], s[1]) < 0.5 || Math.abs(s[0] - s[1]) < 0.3) return null;
    const e = s[0] > s[1] ? 0 : 1;
    return { src: cands[1 - e], ex: [cands[e]] };
  }

  // incoming: [{ key, data, label }] — every export sheet of one upload.
  // → notes for the file strip, naming what went where.
  async function routeExports(incoming) {
    const notes = [];
    const byKey = new Map();
    incoming.forEach(x => { if (!byKey.has(x.key)) byKey.set(x.key, []); byKey.get(x.key).push(x); });
    for (const [key, news] of byKey) {
      const dom = DOMAIN_BY_KEY[key];
      const prior = srcs[key] ? [{ key: key, data: srcs[key], label: srcs[key].sheetLabel, prior: true }] : [];
      const cands = prior.concat(news);
      const clients = cands.filter(c => !isPickTraceShape(c.data));
      let src = null, ex = [], why = '';
      if (cands.length === 1) src = cands[0];
      else if (clients.length) {
        src = clients.filter(c => !c.prior).pop() || clients[0];
        ex = cands.filter(c => isPickTraceShape(c.data));
        why = 'beside the client\'s own ' + dom.label + ' sheet';
      } else {
        const g = guessRoles(key, cands);
        if (g) { src = g.src; ex = g.ex; why = 'its values match the destination template'; }
        else {
          let r;
          try { r = await pickRoles(key, cands); } catch (e) { continue; }
          src = r.src; ex = r.ex;
        }
      }
      if (src) { if (!src.prior) setSource(key, src.data, src.label); }
      else if (prior.length) clearSource(key);
      ex.forEach(c => addExisting(key, c.data, c.label));
      if (ex.length && why) {
        notes.push(dom.label + ': ' + ex.map(c => c.data.fileName).join(', ') + ' read as already in PickTrace (' + why + ')');
      }
    }
    return notes;
  }

  // Which export is which, when the files can't tell: each one is the
  // upload, already in PickTrace, or not used. → { src, ex } or a rejection.
  function pickRoles(key, cands) {
    return new Promise((resolve, reject) => {
      const old = $('dt-role-picker');
      if (old) old.remove();
      const dom = DOMAIN_BY_KEY[key];
      const scores = cands.map(c => fitScore(key, c.data));
      // Best guess: a file already loaded stays the upload; otherwise the
      // one that fits the destination org least.
      let srcI = cands.findIndex(c => c.prior);
      if (srcI < 0) {
        srcI = 0;
        scores.forEach((s, i) => { if (s != null && (scores[srcI] == null || s < scores[srcI])) srcI = i; });
      }
      const overlay = document.createElement('div');
      overlay.id = 'dt-role-picker';
      overlay.className = 'cmp-export-modal';
      overlay.style.display = 'flex';
      const inner = document.createElement('div');
      inner.className = 'cmp-export-modal-inner dt-sheet-modal';
      const opt = (v, l, on) => '<option value="' + v + '"' + (on ? ' selected' : '') + '>' + l + '</option>';
      inner.innerHTML = '<h3>Which ' + escHtml(dom.label) + ' file is which?</h3>' +
        '<p>These are all ' + escHtml(dom.label) + ' exports. Choose the one to <b>build the upload from</b>. A file marked <b>Already in PickTrace</b> ' +
        'is read as what the destination org has, and its rows are left out of the upload.</p>' +
        '<div class="cmp-org-list">' + cands.map((c, i) =>
          '<div class="cmp-org-list-item dt-sheet-row" title="Columns: ' + escHtml(c.data.headers.filter(Boolean).slice(0, 8).join(', ')) + '">' +
          '<span class="dt-sheet-meta"><span class="cmp-org-name">' + escHtml(c.data.fileName + (c.label ? ' › ' + c.label : '')) + '</span>' +
          '<span class="cmp-org-counts">' + escHtml(plural(c.data.rows.length, 'row') + (c.prior ? ' · loaded earlier' : '') +
            (scores[i] != null ? ' · ' + Math.round(scores[i] * 100) + '% of its values are in the destination template' : '')) + '</span></span>' +
          '<select class="input-field dt-role" data-i="' + i + '">' +
          opt('src', 'Build the upload from this', i === srcI) + opt('ex', 'Already in PickTrace', i !== srcI) + opt('skip', 'Don\'t use', false) +
          '</select></div>').join('') + '</div>' +
        '<div class="dt-sheet-err" style="display:none;"></div>' +
        '<div class="cmp-export-modal-actions"><button class="btn btn-primary" data-act="ok">Use these</button>' +
        '<button class="btn btn-ghost" data-act="cancel">Cancel</button></div>';
      const close = () => overlay.remove();
      const cancel = () => { close(); reject(new Error('cancelled')); };
      inner.querySelector('[data-act="ok"]').addEventListener('click', () => {
        const roles = [...inner.querySelectorAll('.dt-role')].map(s => s.value);
        if (roles.filter(r => r === 'src').length > 1) {
          const err = inner.querySelector('.dt-sheet-err');
          err.textContent = 'Only one file can be the upload — mark the others Already in PickTrace or Don\'t use.';
          err.style.display = '';
          return;
        }
        close();
        resolve({ src: cands[roles.indexOf('src')] || null, ex: cands.filter((c, i) => roles[i] === 'ex') });
      });
      inner.querySelector('[data-act="cancel"]').addEventListener('click', cancel);
      overlay.addEventListener('click', e => { if (e.target === overlay) cancel(); });
      overlay.appendChild(inner);
      document.body.appendChild(overlay);
    });
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

  // A slot button says what the file is, so nothing is guessed: Export is
  // the upload, Existing is what PickTrace already has.
  function handleSlot(key, kind, file) {
    readAnyFile(file).then(res => {
      const want = kind === 'existing' ? 'export' : kind;
      if (res.kind !== want) {
        alert('"' + file.name + '" looks like ' + (res.kind === 'template' ? 'a bulk template' : 'an export') +
          ', not ' + (want === 'template' ? 'a template' : 'an export') + '. Load it in the matching slot.');
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
        const s = picks[0].sheet;
        if (kind === 'existing') addExisting(key, s.data, many ? s.name : '');
        else place(asExport(s, many), key);
        placeNotes = [];
        build();
      }, () => {}); // picker cancelled
    }).catch(err => alert('Failed to read ' + file.name + ': ' + (err && err.message ? err.message : err)));
  }

  async function handleBulk(files) {
    const list = Array.from(files);
    if (!list.length) return;
    const results = await Promise.all(list.map(f => readAnyFile(f).then(r => ({ file: f, res: r }), () => ({ file: f, res: null }))));
    const unknown = [], multi = [], incoming = [];
    // Templates first, so the export placement below can compare against
    // the destination org's lists.
    results.forEach(({ file, res }) => {
      if (!res) { unknown.push(file.name); return; }
      if (res.kind === 'template') { if (!place(res)) unknown.push(file.name); return; }
      if (res.sheets.length > 1) { multi.push(res); return; }
      const s = res.sheets[0];
      if (!s.key || !DOMAIN_BY_KEY[s.key]) { unknown.push(file.name); return; }
      incoming.push({ key: s.key, data: s.data, label: '' });
    });
    // One picker per multi-sheet workbook, in turn. A cancelled picker loads
    // nothing from that file.
    for (const res of multi) {
      let picks;
      try { picks = await pickSheets(res, null); } catch (e) { continue; }
      picks.forEach(p => incoming.push({ key: p.key, data: p.sheet.data, label: p.sheet.name }));
    }
    placeNotes = await routeExports(incoming);
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
          dropdowns: t && t.dropdowns ? [...t.dropdowns.entries()].map(([k, v]) => ({ column: k, values: [...v] })) : [],
          validatedColumns: t && t.restricted ? [...t.restricted] : null
        }));
      }
      (exist[d.key] || []).forEach(x => inputs.push(IMDebug.file(d.label + ' — existing PickTrace export', x)));
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
      if ((m.siteGaps || []).length) {
        asked.push(IMDebug.ask(k + '/blank-site', 'data', label + ': blank Site under a row that has one (copied only when confirmed)', {
          count: m.siteGaps.filter(g => !g.applied).length, blocksExport: false,
          items: m.siteGaps.map(g => ({ previewRows: g.rows.map(r => r + 1), site: g.site, copiedDown: !!g.applied })) }));
      }
      if ((m.collisions || []).length) {
        asked.push(IMDebug.ask(k + '/name-collisions', 'data', label + ': rows sharing a name PickTrace keeps only one of', {
          count: m.collisions.length, blocksExport: true,
          items: m.collisions.map(c => ({ shared: c.value, on: c.label, previewRows: c.rows.map(r => r + 1), takenInPickTrace: c.taken })) }));
      }
      if ((m.possible || []).length) {
        asked.push(IMDebug.ask(k + '/possible-duplicates', 'data', label + ': possibly the same person as another employee', {
          count: m.possible.length, blocksExport: false,
          items: m.possible.map(p => ({ previewRow: p.row + 1, employee: p.name, reason: p.reason, matches: p.other })) }));
      }
      if ((m.yearOnly || []).length) {
        asked.push(IMDebug.ask(k + '/year-only-dates', 'data', label + ': dates written as just a year', {
          count: m.yearOnly.reduce((n, y) => n + y.left, 0), blocksExport: false, items: m.yearOnly }));
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
        note: 'exportedRowIndexes[n] is the index into this domain\'s SOURCE rows behind output row n.',
        exportFileName: m.rows.length ? outName(k) : null });
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
          { kind: 'date + number + SSN + phone normalization',
            note: 'Date columns run through toYMD(); numeric columns drop an inline unit via cleanNumber() ("7.25 ac" → 7.25, Acreage 0 → blank); SSN and phone columns are reformatted in place.' }
        ],
        fills: (m.filledCols || []).map(f => ({ column: f.column, value: f.value, scope: f.mode, rows: f.rows, kind: 'column fill' }))
          .concat(Object.keys(smartFix[k] || {}).map(sk => {
            const sep = sk.indexOf('||');
            return { column: m.headers[+sk.slice(0, sep)], from: sk.slice(sep + 2), value: smartFix[k][sk], kind: 'smart fix' };
          })),
        edits: { ageFixes: Object.keys(ageFixes).filter(a => a.indexOf(k + '|') === 0)
          .map(a => ({ sourceRowIndex: +a.slice(k.length + 1), correctedDateOfBirth: ageFixes[a] })),
          cells: Object.keys(cellEdits[k] || {}).map(e => {
            const p = e.lastIndexOf('|');
            return { row: e.slice(0, p), column: m.headers[+e.slice(p + 1)], value: cellEdits[k][e] };
          }),
          removedRows: [...(removed[k] || [])], keptPossibleDuplicates: [...(dupOk[k] || [])] },
        alreadyInPickTrace: { count: (m.already || []).length, items: (m.already || []).slice(0, 500) },
        rowIds: { note: 'rowIds[n] is the source row index behind output row n, or "new:<SITE>" for a site added from Locations.', ids: m.rid },
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
        siteAddresses: siteAddr,
        yearOnlyDateFill: yearFill.mm ? yearFill : null,
        existingExports: DOMAINS.filter(d => (exist[d.key] || []).length)
          .map(d => ({ domain: d.key, files: exist[d.key].map(x => x.fileName + (x.sheetLabel ? ' › ' + x.sheetLabel : '')), rows: existCount(d.key) })),
        autoPlacement: placeNotes,
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
    exist = {}; cellEdits = {}; removed = {}; dupOk = {}; siteAddr = { all: '', ranch: {}, site: {} };
    yearFill = { mm: '', dd: '', range: 'first' }; pvFocus = null; placeNotes = [];
    ageFixes = {}; columnFills = {}; smartFix = {}; spellKeep = {}; srcOverride = {}; fillFilter = ''; fillScope = 'attn'; fillOptOpen = false;
    siteFillOk = {}; activeSection = null; sectionPicked = false; filesOpen = null; pvFilter = 'all'; pvColsOnly = false; pvHidden = false;
    IMOutName.forget();
    { const f = $('dt-fill-filter'); if (f) f.value = ''; }
    { const s = $('dt-fill-scope'); if (s) s.value = 'attn'; }
    renderSiteSrc();
    { const b = $('dt-build-banner'); if (b) b.style.display = 'none'; }
    DOMAINS.forEach(d => {
      slotLabel(d.key, 'src', 'No export');
      if (d.needsTemplate) slotLabel(d.key, 'tpl', 'No template');
      slotLabel(d.key, 'exist', 'None');
      ['src', 'tpl', 'exist'].forEach(kind => { const f = $('dt-' + d.key + '-' + kind + '-file'); if (f) f.value = ''; });
      const btn = $('dt-dl-' + d.key); if (btn) btn.disabled = true;
    });
    const bulk = $('dt-bulk-file'); if (bulk) bulk.value = '';
    ['preview', 'smartfix', 'fill', 'fills', 'underage', 'derived', 'fuzzy', 'addrflag',
      'phonebad', 'remap', 'nooptions', 'autofill', 'unmapped', 'dropped', 'sitegaps',
      'siteaddr', 'collisions', 'possible', 'dates', 'removed', 'existing']
      .forEach(n => { const el = $('dt-section-' + n); if (el) el.style.display = 'none'; });
    ['dt-nav', 'dt-results', 'dt-files-strip'].forEach(id => { const el = $(id); if (el) el.style.display = 'none'; });
    { const f = $('dt-files'); if (f) f.style.display = ''; }
    if (dtSplit) dtSplit.show(false);
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
    // Side by side / stacked — the same layout control as the Template Standardize tools.
    if (window.IMWorkbench) {
      dtSplit = IMWorkbench.split({ preview: $('dt-section-preview'), side: [$('dt-nav'), $('dt-results')],
        grid: $('dt-pv-wrap'), tools: document.querySelector('#dt-section-preview .dt-pv-tools') });
    }
    DOMAINS.forEach(d => {
      ['src', 'tpl', 'exist'].forEach(kind => {
        const el = $('dt-' + d.key + '-' + kind + '-file');
        if (!el) return;
        el.addEventListener('change', e => {
          if (e.target.files[0]) handleSlot(d.key, kind === 'tpl' ? 'template' : kind === 'exist' ? 'existing' : 'export', e.target.files[0]);
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
      // A sure match, or a picker you opened and chose in — never an
      // untouched closest guess.
      picks.forEach(sel => {
        if (!sel.value || (sel.tagName !== 'SELECT' && sel.dataset.sure !== '1')) return;
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
      if (btn.classList.contains('dt-fill-opt-btn')) { fillOptOpen = !fillOptOpen; renderFill(); return; }
      const tr = btn.closest('tr');
      const ci = tr && tr.dataset.ci;
      if (ci == null) return;
      if (btn.classList.contains('dt-fill-show')) {
        pvFocus = { key: previewKey, ci: +ci };
        pvHidden = false;
        IMWorkbench.rowsTop($('dt-preview-table'));
        renderPreview();
        return;
      }
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

    $('dt-spell-box').addEventListener('click', e => {
      const btn = e.target.closest('[data-spell]');
      if (!btn) return;
      const d = btn.dataset, k = d.ci + '||' + String(d.from).toUpperCase();
      if (d.spell === 'use') { applySmartFix(d.key, +d.ci, d.from, d.to); return; }
      const set = spellKeep[d.key] || (spellKeep[d.key] = new Set());
      if (d.spell === 'undo') set.add(k); else set.delete(k);
      build();
    });
    $('dt-build-banner').addEventListener('click', e => {
      const go = e.target.closest('[data-gosec]');
      if (go) { activeSection = go.dataset.gosec; sectionPicked = true; renderNav(); return; }
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
      renderReqSummary();
    });
    $('dt-pv-filterbar').addEventListener('change', e => {
      if (e.target.classList.contains('wb-rowfilter')) { pvFilter = e.target.value; IMWorkbench.rowsTop($('dt-preview-table')); renderPreview(); }
      else if (e.target.classList.contains('wb-colfilter')) { pvColsOnly = e.target.checked; renderPreview(); }
    });
    $('dt-pv-filterbar').addEventListener('click', e => {
      if (!e.target.closest('.wb-focus-clear')) return;
      pvFocus = null;
      IMWorkbench.rowsTop($('dt-preview-table'));
      renderPreview();
    });
    $('dt-pv-toggle').addEventListener('click', () => { pvHidden = !pvHidden; renderPreview(); });

    // Preview edits — a picked or typed value is kept for that row and
    // column across rebuilds; × leaves the row out.
    const pvTable = $('dt-preview-table');
    pvTable.addEventListener('click', e => {
      const rm = e.target.closest('.dt-pv-rm');
      if (!rm || !built || !built[previewKey]) return;
      e.stopPropagation();
      const id = built[previewKey].rid[+rm.dataset.ri];
      if (id == null) return;
      (removed[previewKey] || (removed[previewKey] = new Set())).add(id);
      build();
    });
    pvTable.addEventListener('change', e => {
      const el = e.target;
      if (!el.classList.contains('ts-cell-select') && !el.classList.contains('ts-cell-input')) return;
      const ri = +el.dataset.ri, ci = +el.dataset.ci;
      const b = built && built[previewKey];
      if (!b || !b.rows[ri]) return;
      // "Type custom…" turns the cell into a text box in place.
      if (el.classList.contains('ts-cell-select') && el.value === CELL_CUSTOM) {
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.className = 'ts-cell-input';
        inp.dataset.ri = ri; inp.dataset.ci = ci;
        inp.placeholder = 'Type a value…';
        inp.value = b.rows[ri][ci];
        el.replaceWith(inp);
        inp.focus();
        inp.select();
        return;
      }
      setEdit(previewKey, ri, ci, str(el.value));
    });

    // Sections — a click opens that one and keeps it open across rebuilds.
    $('dt-nav').addEventListener('click', e => {
      const btn = e.target.closest('.dt-nav-btn');
      if (!btn) return;
      activeSection = btn.dataset.sec; sectionPicked = true;
      renderNav();
    });
    $('dt-files-strip').addEventListener('click', e => {
      const swap = e.target.closest('[data-swap]'), un = e.target.closest('[data-unexist]');
      if (swap) {
        const k = swap.dataset.swap, was = srcs[k], ex = (exist[k] || [])[0];
        if (!was || !ex) return;
        exist[k] = [];
        setSource(k, ex, ex.sheetLabel);
        addExisting(k, was, was.sheetLabel);
        placeNotes = [];
        build();
        return;
      }
      if (un) {
        delete exist[un.dataset.unexist];
        existLabel(un.dataset.unexist);
        placeNotes = [];
        build();
        return;
      }
      if (!e.target.closest('#dt-files-toggle')) return;
      filesOpen = $('dt-files').style.display === 'none';
      renderFiles();
    });

    // Site addresses: every site / per ranch / one site — kept on change.
    $('dt-siteaddr-body').addEventListener('change', e => {
      const el = e.target.closest('.dt-addr-in');
      if (!el) return;
      const v = str(el.value), k = el.dataset.key;
      if (el.dataset.scope === 'all') siteAddr.all = v;
      else if (el.dataset.scope === 'ranch') { if (v) siteAddr.ranch[k] = v; else delete siteAddr.ranch[k]; }
      else { if (v) siteAddr.site[k] = v; else delete siteAddr.site[k]; }
      build();
    });

    // Name collisions: a new name is an edit to that cell; Leave out removes
    // the row; "Keep the first of each" removes every later row in a group.
    $('dt-collisions-table').addEventListener('change', e => {
      const inp = e.target.closest('.dt-coll-name');
      if (!inp) return;
      setEdit(inp.dataset.key, +inp.dataset.ri, +inp.dataset.ci, str(inp.value));
    });
    $('dt-collisions-table').addEventListener('click', e => {
      const btn = e.target.closest('.dt-coll-drop');
      if (!btn) return;
      leaveOut(btn.dataset.key, [+btn.dataset.ri]);
    });
    $('dt-collisions-all').addEventListener('click', () => {
      collect('collisions').forEach(c => leaveOut(c.key, c.rows.slice(1), true));
      build();
    });

    // Possible duplicates: Keep stops flagging that row; Leave out removes it.
    $('dt-possible-table').addEventListener('click', e => {
      const btn = e.target.closest('button[data-act]');
      if (!btn) return;
      const k = btn.dataset.key;
      if (btn.dataset.act === 'keep') (dupOk[k] || (dupOk[k] = new Set())).add(btn.dataset.rid);
      else (removed[k] || (removed[k] = new Set())).add(btn.dataset.rid);
      build();
    });
    $('dt-possible-reset').addEventListener('click', () => { dupOk = {}; build(); });

    // Year-only dates: one month + day for every year-only cell.
    $('dt-dates-apply').addEventListener('click', () => {
      const mm = parseInt($('dt-dates-mm').value, 10), dd = parseInt($('dt-dates-dd').value, 10);
      if (!(mm >= 1 && mm <= 12) || !(dd >= 1 && dd <= 31)) { alert('Enter a month (1–12) and a day (1–31).'); return; }
      yearFill = { mm: String(mm), dd: String(dd), range: $('dt-dates-range').value };
      build();
    });
    $('dt-dates-clear').addEventListener('click', () => { yearFill = { mm: '', dd: '', range: 'first' }; build(); });

    // Rows you left out: Restore puts one back.
    $('dt-removed-table').addEventListener('click', e => {
      const btn = e.target.closest('button[data-rid]');
      if (!btn) return;
      const set = removed[btn.dataset.key];
      if (set) set.delete(btn.dataset.rid);
      build();
    });
    $('dt-reqsum').addEventListener('click', e => {
      const chip = e.target.closest('.dt-reqsum-chip');
      if (!chip || !chip.dataset.key) return;
      previewKey = chip.dataset.key;
      renderPreview(); renderFill(); renderReqSummary();
    });
    // Blank Site: copy a group down (or undo it) only on your click.
    $('dt-sitegaps-table').addEventListener('click', e => {
      const btn = e.target.closest('button[data-act]');
      if (!btn) return;
      const set = siteFillOk[btn.dataset.key] || (siteFillOk[btn.dataset.key] = new Set());
      if (btn.dataset.act === 'apply') set.add(btn.dataset.id); else set.delete(btn.dataset.id);
      build();
    });
    $('dt-sitegaps-all').addEventListener('click', () => {
      collect('siteGaps').forEach(g => {
        if (!g.applied) (siteFillOk[g.key] || (siteFillOk[g.key] = new Set())).add(g.id);
      });
      build();
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
    splitWeight: splitWeight,
    setState: function (t, s, e) { tpls = t; srcs = s; exist = e || {}; },
    existIndex: existIndex,
    fitScore: fitScore,
    guessRoles: guessRoles,
    confidentMatch: confidentMatch,
    setAgeFixes: function (f) { ageFixes = f || {}; },
    setColumnFills: function (f) { columnFills = f || {}; },
    setSmartFix: function (f) { smartFix = f || {}; },
    setFillView: function (filter, scope, showOptional) {
      fillFilter = filter || ''; fillScope = scope || 'attn'; fillOptOpen = !!showOptional;
    },
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
