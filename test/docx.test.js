// Word term sheets: what we write must read back to the same deal, including after
// another program re-saves the file with compression.
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const engine = require('../src/engine.js');
const { markdownToDocx, docxToMarkdown } = require('../src/docx.js');
const { parseTermSheet } = require('../src/parser.js');
const { generateDeal } = require('../src/generator.js');
const { termSheet } = require('../src/termsheet.js');
const { crc32 } = require('../src/xlsx.js');

const md = fs.readFileSync(path.join(__dirname, '../docs/coolabah_term_sheet.md'), 'utf8');

test('the Coolabah term sheet survives a trip through Word', async () => {
  const back = await docxToMarkdown(markdownToDocx(md));
  const r = parseTermSheet(back);
  assert.ok(r.ok, 'missing: ' + r.missing.join(', '));
  assert.deepStrictEqual(r.assumed, []);
  assert.strictEqual(r.deal.senior_debt.margin, 0.0175);
  assert.strictEqual(r.deal.meta.name, 'Project Coolabah');
});

test('generated term sheets survive a trip through Word', async () => {
  for (const seed of [1, 7, 4821, 1234, 9999]) {
    const deal = generateDeal(seed);
    const res = engine.run(deal);
    const r = parseTermSheet(await docxToMarkdown(markdownToDocx(termSheet(deal, res))));
    assert.ok(r.ok, 'seed ' + seed + ' missing: ' + r.missing.join(', '));
    assert.ok(Math.abs(engine.run(r.deal).summary.equityIrr - res.summary.equityIrr) < 1e-9, 'seed ' + seed);
  }
});

// Rebuilds a zip with every file compressed (method 8), as Word does when it saves.
function deflatedZip(files) {
  const parts = [], central = [];
  let offset = 0;
  for (const [name, text] of files) {
    const n = Buffer.from(name), raw = Buffer.from(text), data = zlib.deflateRawSync(raw), crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(data.length, 18); local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(n.length, 26);
    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(8, 10);
    cd.writeUInt32LE(crc, 16); cd.writeUInt32LE(data.length, 20); cd.writeUInt32LE(raw.length, 24); cd.writeUInt16LE(n.length, 28); cd.writeUInt32LE(offset, 42);
    parts.push(local, n, data); central.push(cd, n);
    offset += 30 + n.length + data.length;
  }
  const cdBuf = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cdBuf.length, 12); end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...parts, cdBuf, end]));
}

test('a compressed Word file, with words split across runs, is read too', async () => {
  const stored = markdownToDocx(md);
  const doc = await docxToMarkdown(stored); // sanity: the stored version reads
  assert.ok(doc.includes('Margin'));
  // Word often splits a word into several runs when a user edits it: "1.7" + "5%".
  const xml = Buffer.from(stored).toString('utf8');
  const docXml = xml.slice(xml.indexOf('<?xml', xml.indexOf('word/document.xml')), xml.indexOf('</w:document>') + 13)
    .replace('1.75% p.a.', '2.2</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">5% p.a.');
  const bytes = deflatedZip([['word/document.xml', docXml]]);
  const r = parseTermSheet(await docxToMarkdown(bytes));
  assert.ok(r.ok, 'missing: ' + r.missing.join(', '));
  assert.strictEqual(r.deal.senior_debt.margin, 0.0225);
});
