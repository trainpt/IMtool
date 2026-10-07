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
//     fileName              // download name
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
      l.setUint16(6, 0, true);
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
      c.setUint16(8, 0, true);
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
  function buildSheetData(headerRowXml, rows, colCount) {
    const out = ['<sheetData>', headerRowXml];
    rows.forEach((row, ri) => {
      const r = ri + 2; // row 1 is the header
      let cells = '';
      for (let c = 0; c < colCount; c++) {
        const v = row[c];
        if (v == null || v === '') continue;
        cells += '<c r="' + colRef(c) + r + '" t="inlineStr"><is><t xml:space="preserve">' +
          xmlEsc(v) + '</t></is></c>';
      }
      if (cells) out.push('<row r="' + r + '">' + cells + '</row>');
    });
    out.push('</sheetData>');
    return out.join('');
  }

  async function write(opts) {
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
    xml = xml.slice(0, sd.index) + buildSheetData(headerRow, rows, colCount) +
          xml.slice(sd.index + sd[0].length);
    // Keep <dimension> honest; Excel offers to repair the file if it disagrees.
    xml = xml.replace(/<dimension\b[^>]*\/>|<dimension\b[^>]*>[\s\S]*?<\/dimension>/,
      '<dimension ref="A1:' + colRef(Math.max(0, colCount - 1)) + Math.max(1, rows.length + 1) + '"></dimension>');

    const patchedSheet = await replaceEntry(sheet, xml);
    const blob = zipWrite(entries.map(e => e.name === sheet.name ? patchedSheet : e));
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

  window.IMXlsxTemplate = { write: write, zipRead: zipRead, zipWrite: zipWrite, entryText: entryText };
})();
