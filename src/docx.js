/*
 * Word (.docx) term sheets, no dependencies (works in Node 18+ and in the browser).
 *   markdownToDocx(md) : the Markdown term sheet as a Word document (bytes)
 *   docxToMarkdown(bytes) : a Word term sheet back to Markdown, so parser.js can read it
 * A .docx is a zip of XML files: we write it with the zip of xlsx.js, and read it with the
 * browser's built-in DecompressionStream (Word compresses the files inside the zip).
 */
(function (root) {
  'use strict';

  const X = typeof module !== 'undefined' && module.exports ? require('./xlsx.js') : root.PFXlsx;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const cells = (line) => line.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());

  const NAVY = '1B2333', ACCENT = '2F5BEA', MUTED = '6B7280', HEAD_BG = 'F6F7F9', LINE = 'D1D5DB';

  // ---------------- Writing ----------------

  // "**bold** and *italic*" -> Word runs.
  function runs(text, base) {
    base = base || '';
    const out = [];
    const re = /\*\*(.+?)\*\*|\*(.+?)\*/g;
    let last = 0, m;
    const run = (t, props) => { if (t) out.push('<w:r>' + (props || base ? '<w:rPr>' + base + (props || '') + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>'); };
    while ((m = re.exec(text))) {
      run(text.slice(last, m.index));
      run(m[1] || m[2], m[1] ? '<w:b/>' : '<w:i/>');
      last = re.lastIndex;
    }
    run(text.slice(last));
    return out.join('');
  }
  const para = (text, style, base) => '<w:p>' + (style ? '<w:pPr><w:pStyle w:val="' + style + '"/></w:pPr>' : '') + runs(text, base) + '</w:p>';

  function table(rows) {
    const border = (side) => '<w:' + side + ' w:val="single" w:sz="4" w:space="0" w:color="' + LINE + '"/>';
    const widths = rows[0].length === 2 ? [2600, 6400] : rows[0].map(() => Math.floor(9000 / rows[0].length));
    const cell = (text, i, head) => '<w:tc><w:tcPr><w:tcW w:w="' + widths[i] + '" w:type="dxa"/>' +
      (head ? '<w:shd w:val="clear" w:color="auto" w:fill="' + HEAD_BG + '"/>' : '') + '</w:tcPr>' +
      '<w:p><w:pPr><w:spacing w:before="40" w:after="40"/>' + (i > 0 && rows[0].length > 2 ? '<w:jc w:val="right"/>' : '') + '</w:pPr>' +
      runs(text, head || i === 0 ? '<w:b/>' : '') + '</w:p></w:tc>';
    return '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>' + ['top', 'bottom', 'insideH'].map(border).join('') +
      '</w:tblBorders><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid>' + widths.map((w) => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' +
      rows.map((r, ri) => '<w:tr>' + r.map((c, i) => cell(c, i, ri === 0)).join('') + '</w:tr>').join('') + '</w:tbl><w:p/>';
  }

  function bodyXml(md) {
    const lines = md.replace(/\r/g, '').split('\n');
    const out = [];
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) { i++; continue; }
      const h = line.match(/^(#{1,3})\s+(.*)$/);
      if (h) { out.push(para(h[2], 'Heading' + h[1].length)); i++; continue; }
      if (line.trim().startsWith('|')) {
        const rows = [];
        while (i < lines.length && lines[i].trim().startsWith('|')) {
          if (!/^\|\s*-/.test(lines[i].trim())) rows.push(cells(lines[i]));
          i++;
        }
        out.push(table(rows));
        continue;
      }
      const li = line.match(/^\s*(?:- |(\d+)\. )(.*)$/);
      if (li) {
        const box = li[2].match(/^\[( |x)\]\s*(.*)$/);
        const mark = li[1] ? li[1] + '.' : box ? (box[1] === 'x' ? '☒' : '☐') : '•';
        out.push('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:ind w:left="360" w:hanging="360"/></w:pPr>' + runs(mark + '\t' + (box ? box[2] : li[2])) + '</w:p>');
        i++;
        continue;
      }
      const italic = /^\*[^*].*\*$/.test(line.trim());
      out.push(italic ? para(line.trim().slice(1, -1), 'Note') : para(line.trim()));
      i++;
    }
    return out.join('').replace(/<w:t xml:space="preserve">([^<]*)\t/g, '<w:t xml:space="preserve">$1</w:t><w:tab/><w:t xml:space="preserve">');
  }

  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:color w:val="1F2937"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="160"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="6" w:color="' + ACCENT + '"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:color w:val="' + NAVY + '"/><w:sz w:val="40"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="' + ACCENT + '"/><w:sz w:val="26"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Note"><w:name w:val="Note"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="' + MUTED + '"/><w:sz w:val="18"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>' +
    '</w:styles>';

  function markdownToDocx(md) {
    const doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
      bodyXml(md) +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1300" w:right="1300" w:bottom="1300" w:left="1300" w:header="600" w:footer="600" w:gutter="0"/></w:sectPr></w:body></w:document>';
    return X.zip([
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
      ['word/document.xml', doc],
      ['word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['word/styles.xml', STYLES],
    ]);
  }

  // ---------------- Reading ----------------

  // Returns the text of one file inside a zip (stored or deflated).
  async function unzipText(bytes, wanted) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = bytes.length - 22;
    while (eocd >= 0 && dv.getUint32(eocd, true) !== 0x06054b50) eocd--;
    if (eocd < 0) throw new Error('not a zip file');
    let p = dv.getUint32(eocd + 16, true);
    const count = dv.getUint16(eocd + 10, true);
    for (let k = 0; k < count; k++) {
      const method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (name !== wanted) continue;
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const data = bytes.subarray(start, start + size);
      if (method === 0) return new TextDecoder().decode(data);
      if (method !== 8) throw new Error('unsupported zip compression');
      const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new TextDecoder().decode(new Uint8Array(await new Response(stream).arrayBuffer()));
    }
    throw new Error(wanted + ' not found');
  }

  // Plain text of a paragraph: all its text runs joined (Word often splits a word across runs).
  const textOf = (xml) => unesc((xml.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>|<w:tab\/>|<w:br\/>/g) || [])
    .map((t) => (t.startsWith('<w:t') && !t.startsWith('<w:tab') ? t.replace(/<[^>]+>/g, '') : ' ')).join('')).trim();

  async function docxToMarkdown(bytes) {
    const xml = await unzipText(bytes, 'word/document.xml');
    const body = xml.slice(xml.indexOf('<w:body'));
    const out = [];
    for (const block of body.match(/<w:tbl>[\s\S]*?<\/w:tbl>|<w:p\/>|<w:p[ >][\s\S]*?<\/w:p>/g) || []) {
      if (block.startsWith('<w:tbl')) {
        const rows = (block.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || []).map((tr) =>
          (tr.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map((tc) => (tc.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map(textOf).join(' ').replace(/\|/g, '/')));
        rows.forEach((r, i) => {
          out.push('| ' + r.join(' | ') + ' |');
          if (i === 0) out.push('|' + r.map(() => ' --- |').join(''));
        });
        out.push('');
        continue;
      }
      const text = textOf(block);
      if (!text) continue;
      const style = (block.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
      const level = (style.match(/^(?:Heading|Titre)(\d)$/i) || [])[1] || (style === 'Title' ? '1' : '');
      out.push((level ? '#'.repeat(Number(level)) + ' ' : '') + text, '');
    }
    return out.join('\n');
  }

  const api = { markdownToDocx, docxToMarkdown };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFDocx = api;
})(typeof window !== 'undefined' ? window : globalThis);
