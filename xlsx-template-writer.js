// ═══════════════════════════════════════════════════════════════════════
// Filled-template writer — patches the template zip instead of rebuilding it.
//
// An .xlsx is a zip of XML parts. Reading one into SheetJS and writing it back
// re-generates every part from SheetJS's own model, which silently drops
// anything that model doesn't cover. On PickTrace's bulk templates that costs
// you the <dataValidations> (every dropdown) and most of styles.xml.
//
// Worse, SheetJS's default writer emits string cells as t="str". In OOXML that
// means "cached result of a formula", NOT a literal string — a literal is
// t="s" (shared-string index) or t="inlineStr". Strict readers see t="str"
// with no <f> element and hand back nothing, which is how a perfectly correct
// grid earns "The uploaded file contains an unexpected column header."
//
// So: treat the template as the source of truth and patch it. Every zip entry
// is copied through with its ORIGINAL compressed bytes, and only the DATA ENTRY
// sheet's <sheetData> is rewritten — as inline strings, so sharedStrings.xml,
// the dropdowns, the styles and the header row all survive untouched.
//
// Lifted from locations-migrate.js, which has been shipping this since the
// Legacy migration; that module still carries its own private copy.
//
//   await IMXlsxTemplate.write({
//     rawBuffer,            // ArrayBuffer of the personalized template
//     headers,              // output headers (column count only)
//     rows,                 // [[...]] data rows, written from sheet row 2
//     fileName,             // download name
//     marks                 // optional: Map<rowIdx, Map<colIdx, 'req' | 'off'>> — cells
//                           // to colour for whoever fills the file in (red / amber).
//                           // A marked empty cell is written as a styled blank.
//     notes                 // optional: Map<rowIdx (-1 = header), Map<colIdx, text>> — hover
//                           // comments on those cells (noteAuthor: the name shown)
//   });
// ═══════════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Parse the central directory, keeping each entry's raw (still-compressed)
  // bytes so untouched parts can be re-emitted verbatim.
  function zipRead(buffer) {
    const u8 = new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    const floor = Math.max(0, u8.length - 22 - 65535);
    for (let i = u8.length - 22; i >= floor; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Template is not a valid .xlsx (no zip end record).');
    const count = dv.getUint16(eocd + 10, true);
    let off = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(off, true) !== 0x02014b50) throw new Error('Corrupt zip central directory.');
      const method = dv.getUint16(off + 10, true);
      const mtime = dv.getUint16(off + 12, true);
      const mdate = dv.getUint16(off + 14, true);
      const crc = dv.getUint32(off + 16, true);
      const csize = dv.getUint32(off + 20, true);
      const usize = dv.getUint32(off + 24, true);
      const nlen = dv.getUint16(off + 28, true);
      const elen = dv.getUint16(off + 30, true);
      const clen = dv.getUint16(off + 32, true);
      const lho = dv.getUint32(off + 42, true);
      const name = dec.decode(u8.subarray(off + 46, off + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true);
      const lelen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lelen;
      entries.push({ name, method, crc, usize, mtime, mdate, data: u8.subarray(start, start + csize) });
      off += 46 + nlen + elen + clen;
    }
    return entries;
  }

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  async function entryText(e) {
    return new TextDecoder().decode(e.method === 0 ? e.data : await inflateRaw(e.data));
  }
  // Deflate the replaced part when the platform offers CompressionStream; fall
  // back to stored (method 0), which is equally valid, if it's missing or fails.
  async function replaceEntry(e, text) {
    const data = new TextEncoder().encode(text);
    const crc = crc32(data);
    let out = data, method = 0;
    if (typeof CompressionStream === 'function') {
      try {
        const s = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
        const z = new Uint8Array(await new Response(s).arrayBuffer());
        if (z.length < data.length) { out = z; method = 8; }
      } catch (err) { /* keep stored */ }
    }
    return { name: e.name, method, crc, usize: data.length, mtime: e.mtime, mdate: e.mdate, data: out };
  }

  function zipWrite(entries) {
    const enc = new TextEncoder();
    const parts = [], centrals = [];
    let offset = 0;
    entries.forEach(e => {
      const nb = enc.encode(e.name);
      const lh = new Uint8Array(30 + nb.length);
      const l = new DataView(lh.buffer);
      l.setUint32(0, 0x04034b50, true);
      l.setUint16(4, 20, true);
      l.setUint16(6, e.flags || 0, true);
      l.setUint16(8, e.method, true);
      l.setUint16(10, e.mtime, true);
      l.setUint16(12, e.mdate, true);
      l.setUint32(14, e.crc, true);
      l.setUint32(18, e.data.length, true);
      l.setUint32(22, e.usize, true);
      l.setUint16(26, nb.length, true);
      l.setUint16(28, 0, true);
      lh.set(nb, 30);
      parts.push(lh, e.data);

      const ch = new Uint8Array(46 + nb.length);
      const c = new DataView(ch.buffer);
      c.setUint32(0, 0x02014b50, true);
      c.setUint16(4, 20, true);
      c.setUint16(6, 20, true);
      c.setUint16(8, e.flags || 0, true);
      c.setUint16(10, e.method, true);
      c.setUint16(12, e.mtime, true);
      c.setUint16(14, e.mdate, true);
      c.setUint32(16, e.crc, true);
      c.setUint32(20, e.data.length, true);
      c.setUint32(24, e.usize, true);
      c.setUint16(28, nb.length, true);
      c.setUint32(42, offset, true);
      ch.set(nb, 46);
      centrals.push(ch);
      offset += lh.length + e.data.length;
    });
    let cdSize = 0;
    centrals.forEach(c => { cdSize += c.length; });
    const eocd = new Uint8Array(22);
    const e2 = new DataView(eocd.buffer);
    e2.setUint32(0, 0x06054b50, true);
    e2.setUint16(8, entries.length, true);
    e2.setUint16(10, entries.length, true);
    e2.setUint32(12, cdSize, true);
    e2.setUint32(16, offset, true);
    return new Blob(parts.concat(centrals, [eocd]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const xmlEsc = s => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // Strip characters XML 1.0 forbids outright — free-text source fields
    // occasionally carry stray control bytes that would corrupt the sheet.
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

  function colRef(n) {
    let s = '';
    n += 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  // Locate the DATA ENTRY worksheet part by walking workbook.xml → rels.
  async function findSheetPart(entries, sheetMatch) {
    const byName = {};
    entries.forEach(e => { byName[e.name] = e; });
    if (!byName['xl/workbook.xml'] || !byName['xl/_rels/workbook.xml.rels']) {
      throw new Error('Template is missing its workbook parts.');
    }
    const wbXml = await entryText(byName['xl/workbook.xml']);
    const relsXml = await entryText(byName['xl/_rels/workbook.xml.rels']);
    const sheets = [...wbXml.matchAll(/<sheet\b[^>]*?(?:\/>|><\/sheet>)/g)].map(m => m[0]);
    const want = sheetMatch || /data.?entry/i;
    let rid = null;
    for (const s of sheets) {
      const nm = (s.match(/name="([^"]*)"/) || [])[1] || '';
      if (want.test(nm)) { rid = (s.match(/r:id="([^"]+)"/) || [])[1]; break; }
    }
    if (!rid && sheets.length) rid = (sheets[0].match(/r:id="([^"]+)"/) || [])[1];
    if (!rid) throw new Error('Could not find a DATA ENTRY sheet in the template.');
    const rel = [...relsXml.matchAll(/<Relationship\b[^>]*?(?:\/>|><\/Relationship>)/g)]
      .map(m => m[0]).find(r => (r.match(/Id="([^"]+)"/) || [])[1] === rid);
    if (!rel) throw new Error('Template worksheet relationship is missing.');
    let target = (rel.match(/Target="([^"]+)"/) || [])[1] || '';
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const path = 'xl/' + target;
    if (!byName[path]) throw new Error('Template worksheet part not found: ' + path);
    return byName[path];
  }

  // <sheetData>: the template's own header row verbatim, then our rows as
  // inline strings. Keeping the header byte-for-byte is the point — whatever
  // the uploader matches on, it sees exactly what it shipped.
  function buildSheetData(headerRowXml, rows, colCount, marks, styleOf) {
    const out = ['<sheetData>', headerRowXml];
    rows.forEach((row, ri) => {
      const r = ri + 2; // row 1 is the header
      const rm = marks && marks.get(ri);
      let cells = '';
      for (let c = 0; c < colCount; c++) {
        const v = row[c];
        const st = rm && rm.has(c) && styleOf ? styleOf[rm.get(c)] : null;
        const sa = st != null ? ' s="' + st + '"' : '';
        if (v == null || v === '') {
          if (sa) cells += '<c r="' + colRef(c) + r + '"' + sa + '/>';
          continue;
        }
        cells += '<c r="' + colRef(c) + r + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' +
          xmlEsc(v) + '</t></is></c>';
      }
      if (cells) out.push('<row r="' + r + '">' + cells + '</row>');
    });
    out.push('</sheetData>');
    return out.join('');
  }

  // Two highlight styles for marks: a red and an amber fill, appended to the
  // template's own styles.xml (its fills and cell formats are left as they
  // are; ours go on the end). → { xml, styleOf: { req, off } } or null.
  const MARK_FILL = { req: 'FFFEE2E2', off: 'FFFEF3C7' };
  function addMarkStyles(xml) {
    const fm = xml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
    const xm = xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
    if (!fm || !xm) return null;
    const nFills = (fm[1].match(/<fill[\s>\/]/g) || []).length;
    const nXfs = (xm[1].match(/<xf[\s>\/]/g) || []).length;
    const keys = Object.keys(MARK_FILL), styleOf = {};
    const fills = keys.map(k => '<fill><patternFill patternType="solid"><fgColor rgb="' + MARK_FILL[k] + '"/><bgColor indexed="64"/></patternFill></fill>').join('');
    const xfs = keys.map((k, i) => { styleOf[k] = nXfs + i;
      return '<xf numFmtId="49" fontId="0" fillId="' + (nFills + i) + '" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>'; }).join('');
    // Append to a block and bump its count="…".
    const grow = (block, tag, add, n) => block.replace(new RegExp('</' + tag + '>$'), add + '</' + tag + '>')
      .replace(new RegExp('(<' + tag + '\\b[^>]*\\bcount=")\\d+(")'), '$1' + n + '$2');
    let out = xml.replace(fm[0], grow(fm[0], 'fills', fills, nFills + keys.length));
    const xm2 = out.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
    out = out.replace(xm2[0], grow(xm2[0], 'cellXfs', xfs, nXfs + keys.length));
    return { xml: out, styleOf: styleOf };
  }

  // ─── Cell comments ("notes" in Excel) ───
  // notes: Map<rowIdx, Map<colIdx, text>>; rowIdx -1 is the header row. Each
  // becomes a hover note on that cell. Excel needs three things for a note:
  // xl/commentsN.xml (the text), a VML drawing (the yellow box, hidden until
  // hover) and the sheet wired to both (its .rels + a <legacyDrawing>), plus
  // their content types. The template's own parts are otherwise untouched.
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  function noteList(notes) {
    const list = [];
    notes.forEach((m, ri) => m.forEach((text, ci) => { if (text) list.push({ ref: colRef(ci) + (ri + 2), r: ri + 1, c: ci, text: String(text) }); }));
    return list;
  }
  function commentsXml(list, author) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>' + xmlEsc(author) + '</author></authors><commentList>' +
      list.map(n => '<comment ref="' + n.ref + '" authorId="0"><text><r><rPr><sz val="10"/><rFont val="Calibri"/><family val="2"/></rPr>' +
        '<t xml:space="preserve">' + xmlEsc(n.text) + '</t></r></text></comment>').join('') +
      '</commentList></comments>';
  }
  function vmlXml(list) {
    const shape = (n, i) => {
      const lines = Math.min(12, Math.max(3, Math.ceil(n.text.length / 38) + (n.text.match(/\n/g) || []).length));
      return '<v:shape id="_x0000_s' + (1025 + i) + '" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:230pt;height:' +
        (lines * 13 + 8) + 'pt;z-index:' + (i + 1) + ';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">' +
        '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        '<x:Anchor>' + (n.c + 1) + ', 15, ' + n.r + ', 2, ' + (n.c + 4) + ', 15, ' + (n.r + Math.ceil(lines / 1.3) + 1) + ', 4</x:Anchor>' +
        '<x:AutoFill>False</x:AutoFill><x:Row>' + n.r + '</x:Row><x:Column>' + n.c + '</x:Column></x:ClientData></v:shape>';
    };
    return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
      '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' + list.map(shape).join('') + '</xml>';
  }
  // → { sheetXml, parts: [{ name, text, replaces? }] } or null when the sheet
  // already carries notes (left alone rather than merged).
  async function addComments(entries, sheet, sheetXml, notes, author) {
    const list = noteList(notes);
    if (!list.length || /<legacyDrawing\b/.test(sheetXml)) return null;
    const names = new Set(entries.map(e => e.name));
    let n = 1;
    while (names.has('xl/comments' + n + '.xml') || names.has('xl/drawings/vmlDrawing' + n + '.vml')) n++;
    const cName = 'xl/comments' + n + '.xml', vName = 'xl/drawings/vmlDrawing' + n + '.vml';
    // The sheet's relationships: add ours to the existing file, or start one.
    const dir = sheet.name.slice(0, sheet.name.lastIndexOf('/') + 1), base = sheet.name.slice(dir.length);
    const relsName = dir + '_rels/' + base + '.rels';
    const relsEntry = entries.find(e => e.name === relsName);
    let rels = relsEntry ? await entryText(relsEntry)
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const used = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map(m => m[1]));
    const rid = p => { let i = 1; while (used.has(p + i)) i++; used.add(p + i); return p + i; };
    const cId = rid('rIdNote'), vId = rid('rIdNoteVml');
    rels = rels.replace('</Relationships>',
      '<Relationship Id="' + cId + '" Type="' + NS_REL + '/comments" Target="../comments' + n + '.xml"/>' +
      '<Relationship Id="' + vId + '" Type="' + NS_REL + '/vmlDrawing" Target="../drawings/vmlDrawing' + n + '.vml"/></Relationships>');
    // <legacyDrawing> goes after <drawing> and before these, per the schema.
    let xml = sheetXml;
    if (!/<worksheet\b[^>]*\bxmlns:r=/.test(xml)) xml = xml.replace(/<worksheet\b/, '<worksheet xmlns:r="' + NS_REL + '"');
    const tag = '<legacyDrawing r:id="' + vId + '"/>';
    const after = xml.search(/<(legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/);
    xml = after >= 0 ? xml.slice(0, after) + tag + xml.slice(after) : xml.replace('</worksheet>', tag + '</worksheet>');
    // Content types: VML by extension, the comments part by name.
    const ctEntry = entries.find(e => e.name === '[Content_Types].xml');
    let ct = await entryText(ctEntry);
    if (!/Extension="vml"/i.test(ct)) ct = ct.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>');
    ct = ct.replace('</Types>', '<Override PartName="/' + cName + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>');
    return { sheetXml: xml, parts: [
      { name: '[Content_Types].xml', text: ct, replaces: true },
      { name: relsName, text: rels, replaces: !!relsEntry },
      { name: cName, text: commentsXml(list, author || 'PickTrace') },
      { name: vName, text: vmlXml(list) }
    ] };
  }

  // The filled template as a Blob — write() downloads it; a bundle (one .zip of
  // several files) takes its bytes instead.
  async function build(opts) {
    const o = opts || {};
    if (!o.rawBuffer) throw new Error('No template buffer to write into.');
    const rows = o.rows || [];
    const colCount = (o.headers || []).length || (rows[0] ? rows[0].length : 0);
    const entries = zipRead(o.rawBuffer);
    const sheet = await findSheetPart(entries, o.sheetMatch);
    let xml = await entryText(sheet);

    const sd = xml.match(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>|<sheetData\b[^>]*\/>/);
    if (!sd) throw new Error('Template worksheet has no <sheetData> element.');
    const headerRow = (sd[0].match(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/) || [''])[0];
    if (!headerRow) throw new Error('Template worksheet has no header row to preserve.');
    // Highlights (optional): patch styles.xml once, then point cells at it.
    let styleOf = null, stylesEntry = null, stylesXml = null;
    if (o.marks && o.marks.size) {
      stylesEntry = entries.find(e => e.name === 'xl/styles.xml');
      const patched = stylesEntry ? addMarkStyles(await entryText(stylesEntry)) : null;
      if (patched) { styleOf = patched.styleOf; stylesXml = patched.xml; }
    }
    xml = xml.slice(0, sd.index) + buildSheetData(headerRow, rows, colCount, o.marks, styleOf) +
          xml.slice(sd.index + sd[0].length);
    // Keep <dimension> honest; Excel offers to repair the file if it disagrees.
    xml = xml.replace(/<dimension\b[^>]*\/>|<dimension\b[^>]*>[\s\S]*?<\/dimension>/,
      '<dimension ref="A1:' + colRef(Math.max(0, colCount - 1)) + Math.max(1, rows.length + 1) + '"></dimension>');

    // Notes (optional): hover comments on flagged cells / column headers.
    const extra = [];
    if (o.notes && o.notes.size) {
      const c = await addComments(entries, sheet, xml, o.notes, o.noteAuthor);
      if (c) { xml = c.sheetXml; extra.push(...c.parts); }
    }

    const patchedSheet = await replaceEntry(sheet, xml);
    const patchedStyles = stylesXml ? await replaceEntry(stylesEntry, stylesXml) : null;
    const swap = new Map();
    for (const p of extra) swap.set(p.name, await replaceEntry({ name: p.name, mtime: sheet.mtime, mdate: sheet.mdate }, p.text));
    const out = entries.map(e => e.name === sheet.name ? patchedSheet
      : patchedStyles && e.name === stylesEntry.name ? patchedStyles
      : swap.has(e.name) ? swap.get(e.name) : e);
    extra.forEach(p => { if (!p.replaces) out.push(swap.get(p.name)); });
    return zipWrite(out);
  }

  async function write(opts) {
    const o = opts || {};
    const blob = await build(o);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = o.fileName || 'filled.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return true;
  }

  // A plain .zip of whole files — Download All's one-folder bundle. Each file
  // is already a compressed .xlsx, so it is stored as-is. Names may carry a
  // folder ("Folder/File.xlsx") and accented characters, so they are flagged
  // UTF-8 (bit 11) — without it Windows shows "Crème" as "CrÃ¨me".
  function bundle(files) {
    const now = new Date();
    const mtime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const mdate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const entries = files.map(f => {
      const data = f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
      return { name: f.name, method: 0, crc: crc32(data), usize: data.length, mtime, mdate, data, flags: 0x0800 };
    });
    return new Blob([zipWrite(entries)], { type: 'application/zip' });
  }

  window.IMXlsxTemplate = { write: write, build: build, bundle: bundle, zipRead: zipRead, zipWrite: zipWrite, entryText: entryText };
})();
