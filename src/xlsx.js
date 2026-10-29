/*
 * Minimal .xlsx writer, no dependencies (works in Node and in the browser).
 * An .xlsx file is a zip of XML files. This writes only what the model export needs:
 * several sheets, numbers / text / booleans / formulas (with a cached value), cell styles,
 * column widths, frozen panes, defined names and drop-down lists.
 * Files are stored in the zip without compression, which Excel accepts.
 */
(function (root) {
  'use strict';

  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // Column number (1-based) <-> letters: 1 -> A, 27 -> AA.
  function colName(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colNumber(letters) {
    return letters.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
  }

  // Excel stores dates as days since 30 Dec 1899.
  function excelDate(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  class Sheet {
    constructor(name) {
      this.name = name;
      this.rows = new Map(); // row number -> Map(col number -> cell)
      this.widths = [];
      this.freeze = null;
      this.lists = [];
      this.tab = null;
      this.heights = new Map(); // row number -> height in points
    }
    // value: number | string | boolean | { f: 'formula without =', v: cached value } ; style: style id
    set(ref, value, style) {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      const r = Number(m[2]), c = colNumber(m[1]);
      if (!this.rows.has(r)) this.rows.set(r, new Map());
      this.rows.get(r).set(c, { value, style: style || 0 });
      return this;
    }
    width(col, w) { this.widths.push([col, w]); return this; }
    height(row, h) { this.heights.set(row, h); return this; }
    freezeAt(ref) { this.freeze = ref; return this; }
    dropdown(ref, options) { this.lists.push([ref, options]); return this; }

    cellXml(ref, { value, style }) {
      const s = style ? ' s="' + style + '"' : '';
      if (value === null || value === undefined || value === '') return '<c r="' + ref + '"' + s + '/>';
      if (typeof value === 'number') return isFinite(value) ? '<c r="' + ref + '"' + s + '><v>' + value + '</v></c>' : '<c r="' + ref + '"' + s + '/>';
      if (typeof value === 'boolean') return '<c r="' + ref + '"' + s + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
      if (typeof value === 'string') return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(value) + '</t></is></c>';
      // Formula, with an optional cached result so viewers that do not recalculate still show numbers.
      const v = value.v;
      const f = '<f>' + xmlEsc(value.f) + '</f>';
      if (typeof v === 'number' && isFinite(v)) return '<c r="' + ref + '"' + s + '>' + f + '<v>' + v + '</v></c>';
      if (typeof v === 'boolean') return '<c r="' + ref + '"' + s + ' t="b">' + f + '<v>' + (v ? 1 : 0) + '</v></c>';
      if (typeof v === 'string') return '<c r="' + ref + '"' + s + ' t="str">' + f + '<v>' + xmlEsc(v) + '</v></c>';
      return '<c r="' + ref + '"' + s + '>' + f + '</c>';
    }

    toXml() {
      let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
      if (this.tab) x += '<sheetPr><tabColor rgb="FF' + this.tab + '"/></sheetPr>';
      x += '<sheetViews><sheetView workbookViewId="0" showGridLines="0">';
      if (this.freeze) {
        const m = this.freeze.match(/^([A-Z]+)(\d+)$/);
        const xs = colNumber(m[1]) - 1, ys = Number(m[2]) - 1;
        const pane = xs && ys ? 'bottomRight' : xs ? 'topRight' : 'bottomLeft';
        x += '<pane' + (xs ? ' xSplit="' + xs + '"' : '') + (ys ? ' ySplit="' + ys + '"' : '') + ' topLeftCell="' + this.freeze + '" activePane="' + pane + '" state="frozen"/><selection pane="' + pane + '"/>';
      }
      x += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>';
      if (this.widths.length) {
        x += '<cols>' + this.widths.sort((a, b) => a[0] - b[0]).map(([c, w]) => '<col min="' + c + '" max="' + c + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>';
      }
      x += '<sheetData>';
      [...this.rows.keys()].sort((a, b) => a - b).forEach((r) => {
        const cells = this.rows.get(r);
        const ht = this.heights.get(r);
        x += '<row r="' + r + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + [...cells.keys()].sort((a, b) => a - b).map((c) => this.cellXml(colName(c) + r, cells.get(c))).join('') + '</row>';
      });
      x += '</sheetData>';
      if (this.lists.length) {
        x += '<dataValidations count="' + this.lists.length + '">' + this.lists.map(([ref, opts]) =>
          '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="' + ref + '"><formula1>"' + xmlEsc(opts.join(',')) + '"</formula1></dataValidation>').join('') + '</dataValidations>';
      }
      x += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
      return x;
    }
  }

  class Workbook {
    constructor() {
      this.sheets = [];
      this.names = [];
      this.styles = [{}]; // style 0 = default
      this.styleKeys = new Map([['{}', 0]]);
    }
    addSheet(name) { const s = new Sheet(name); this.sheets.push(s); return s; }
    defineName(name, ref) { this.names.push([name, ref]); }

    // Style spec: { bold, italic, color: 'RRGGBB', fill: 'RRGGBB', numFmt: '#,##0', align: 'right', valign: 'center', border: 'top' | 'bottom', size, wrap }
    style(spec) {
      const key = JSON.stringify(spec);
      if (!this.styleKeys.has(key)) { this.styleKeys.set(key, this.styles.length); this.styles.push(spec); }
      return this.styleKeys.get(key);
    }

    stylesXml() {
      const fonts = ['<font><sz val="10"/><name val="Arial"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
      const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], fmts = [];
      const idx = (list, xml) => { let i = list.indexOf(xml); if (i < 0) { list.push(xml); i = list.length - 1; } return i; };
      const xfs = this.styles.map((s) => {
        const font = idx(fonts, '<font>' + (s.bold ? '<b/>' : '') + (s.italic ? '<i/>' : '') + '<sz val="' + (s.size || 10) + '"/>' + (s.color ? '<color rgb="FF' + s.color + '"/>' : '') + '<name val="Arial"/></font>');
        const fill = s.fill ? idx(fills, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + s.fill + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
        const side = (n) => (s.border === n ? '<' + n + ' style="thin"><color rgb="FF9CA3AF"/></' + n + '>' : '<' + n + '/>');
        const border = s.border ? idx(borders, '<border><left/><right/>' + side('top') + side('bottom') + '<diagonal/></border>') : 0;
        const fmt = s.numFmt ? 164 + idx(fmts, s.numFmt) : 0;
        const align = s.align || s.wrap || s.valign ? '<alignment' + (s.align ? ' horizontal="' + s.align + '"' : '') + (s.wrap ? ' wrapText="1" vertical="top"' : s.valign ? ' vertical="' + s.valign + '"' : '') + '/>' : '';
        return '<xf numFmtId="' + fmt + '" fontId="' + font + '" fillId="' + fill + '" borderId="' + border + '" xfId="0"' +
          (fmt ? ' applyNumberFormat="1"' : '') + (font ? ' applyFont="1"' : '') + (fill ? ' applyFill="1"' : '') + (border ? ' applyBorder="1"' : '') +
          (align ? ' applyAlignment="1">' + align + '</xf>' : '/>');
      });
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (fmts.length ? '<numFmts count="' + fmts.length + '">' + fmts.map((f, i) => '<numFmt numFmtId="' + (164 + i) + '" formatCode="' + xmlEsc(f) + '"/>').join('') + '</numFmts>' : '') +
        '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts>' +
        '<fills count="' + fills.length + '">' + fills.join('') + '</fills>' +
        '<borders count="' + borders.length + '">' + borders.join('') + '</borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }

    files() {
      const n = this.sheets.length;
      const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        this.sheets.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>';
      const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
      const wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
        this.sheets.map((s, i) => '<sheet name="' + xmlEsc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets>' +
        (this.names.length ? '<definedNames>' + this.names.map(([nm, ref]) => '<definedName name="' + nm + '">' + xmlEsc(ref) + '</definedName>').join('') + '</definedNames>' : '') +
        '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
      const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        this.sheets.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
        '<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
      const out = [['[Content_Types].xml', ct], ['_rels/.rels', rels], ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', wbRels], ['xl/styles.xml', this.stylesXml()]];
      this.sheets.forEach((s, i) => out.push(['xl/worksheets/sheet' + (i + 1) + '.xml', s.toXml()]));
      return out;
    }

    // Returns the .xlsx file as bytes (Uint8Array).
    toBytes() { return zip(this.files()); }
  }

  // ---- Zip (store method, no compression) ----
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function zip(files) {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    const DOS_TIME = 0, DOS_DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;
    for (const [name, text] of files) {
      const nameBytes = enc.encode(name), data = enc.encode(text), crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
      local.setUint16(8, 0, true); local.setUint16(10, DOS_TIME, true); local.setUint16(12, DOS_DATE, true);
      local.setUint32(14, crc, true); local.setUint32(18, data.length, true); local.setUint32(22, data.length, true);
      local.setUint16(26, nameBytes.length, true); local.setUint16(28, 0, true);
      parts.push(new Uint8Array(local.buffer), nameBytes, data);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, 0, true); cd.setUint16(12, DOS_TIME, true); cd.setUint16(14, DOS_DATE, true);
      cd.setUint32(16, crc, true); cd.setUint32(20, data.length, true); cd.setUint32(24, data.length, true);
      cd.setUint16(28, nameBytes.length, true); cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), nameBytes);
      offset += 30 + nameBytes.length + data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let p = 0;
    for (const b of all) { out.set(b, p); p += b.length; }
    return out;
  }

  const api = { Workbook, colName, colNumber, excelDate, crc32, zip };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFXlsx = api;
})(typeof window !== 'undefined' ? window : globalThis);
