// Term sheet reader: a Model Builder term sheet must give back the deal it describes.
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const engine = require('../src/engine.js');
const { parseTermSheet } = require('../src/parser.js');
const { generateDeal } = require('../src/generator.js');
const { termSheet } = require('../src/termsheet.js');
const coolabah = require('../deal.json');

const SECTIONS = ['timeline', 'construction', 'operations', 'tax', 'senior_debt', 'dsra', 'covenants'];
const md = fs.readFileSync(path.join(__dirname, '../docs/coolabah_term_sheet.md'), 'utf8');

function sameInputs(a, b, label) {
  for (const s of SECTIONS) {
    for (const k of Object.keys(a[s])) {
      const x = a[s][k], y = b[s][k];
      if (typeof x === 'number') assert.ok(Math.abs(x - y) < 1e-9, label + ': ' + s + '.' + k + ' ' + x + ' vs ' + y);
      else assert.deepStrictEqual(y, x, label + ': ' + s + '.' + k);
    }
  }
}

test('the Coolabah term sheet gives back deal.json', () => {
  const r = parseTermSheet(md);
  assert.ok(r.ok, 'missing: ' + r.missing.join(', '));
  assert.deepStrictEqual(r.assumed, []);
  sameInputs(coolabah, r.deal, 'Coolabah');
  assert.strictEqual(r.deal.meta.name, 'Project Coolabah');
  near(engine.run(r.deal).summary.sculptDscr, 1.475, 0.002);
});

test('every input read from the term sheet cites its clause', () => {
  const r = parseTermSheet(md);
  assert.strictEqual(r.deal.sources['senior_debt.margin'], 'Section 3 · Margin: 1.75%');
  for (const s of SECTIONS) for (const k of Object.keys(r.deal[s])) {
    if (k === 'sizing') continue;
    assert.ok(r.deal.sources[s + '.' + k], 'no source for ' + s + '.' + k);
  }
});

test('generated term sheets read back to the same deal and the same model', () => {
  for (let seed = 1; seed <= 100; seed++) {
    const deal = generateDeal(seed);
    const res = engine.run(deal);
    const r = parseTermSheet(termSheet(deal, res));
    assert.ok(r.ok, 'seed ' + seed + ' missing: ' + r.missing.join(', '));
    sameInputs(deal, r.deal, 'seed ' + seed);
    const back = engine.run(r.deal).summary;
    assert.ok(Math.abs(back.equityIrr - res.summary.equityIrr) < 1e-9, 'seed ' + seed + ' IRR');
  }
});

test('an edited term sheet changes the model', () => {
  const edited = md.replace('| Margin | 1.75% p.a.', '| Margin | 2.25% p.a.');
  const r = parseTermSheet(edited);
  assert.strictEqual(r.deal.senior_debt.margin, 0.0225);
  assert.ok(engine.run(r.deal).summary.sculptDscr < 1.475);
});

test('missing core terms are reported, minor ones are assumed and flagged', () => {
  const noMargin = md.split('\n').filter((l) => !l.startsWith('| Margin |')).join('\n');
  const r = parseTermSheet(noMargin);
  assert.strictEqual(r.ok, false);
  assert.ok(r.missing.includes('margin'));

  const noCpi = md.split('\n').filter((l) => !l.startsWith('| CPI |')).join('\n');
  const c = parseTermSheet(noCpi);
  assert.ok(c.ok);
  assert.deepStrictEqual(c.assumed, ['operations.cpi']);
  assert.ok(c.deal.sources['operations.cpi'].startsWith('ASSUMED'));
});

test('a free-form document is not mistaken for a term sheet', () => {
  const r = parseTermSheet('Dear all, please find attached the indicative terms for our wind farm. Margin is 2%.');
  assert.strictEqual(r.ok, false);
  assert.ok(r.missing.length > 10);
});

function near(a, b, tol) { assert.ok(Math.abs(a - b) <= tol, a + ' vs ' + b); }

test('the blank template, once filled in, reads back to the Coolabah deal', () => {
  const tpl = fs.readFileSync(path.join(__dirname, '../docs/term_sheet_template.md'), 'utf8');
  const blank = parseTermSheet(tpl);
  assert.strictEqual(blank.ok, false, 'an unfilled template must not build a model');
  // Each placeholder, in the order it appears, with the Coolabah figure.
  const values = ['Coolabah', '100', 'NSW', '1 July 2027', '1 July 2029', 'Coolabah', '100', 'NSW', 'Coolabah', 'Wattlebank', '20', '60.00', '100', '30 June 2049',
    '98,000,000', '30 June 2047', '140,000', '15', '35', '35', '15',
    '4.00', '1.75', '1.50', '40', 'sculpted', '1.30', '6', '1.10', '1.05', '70',
    '100', '27.0', '99.0', '0.5', '3.20', '2.5', '30', '20'];
  let i = 0;
  const filled = tpl.replace(/\[[^\]]+\]/g, () => values[i++]);
  assert.strictEqual(i, values.length, 'one value per placeholder');
  const r = parseTermSheet(filled);
  assert.ok(r.ok, 'missing: ' + r.missing.join(', '));
  assert.deepStrictEqual(r.assumed, []);
  sameInputs(coolabah, r.deal, 'filled template');
});
