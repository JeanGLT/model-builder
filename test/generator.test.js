// Random deal generator: every seed must give a bankable, internally consistent deal.
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const engine = require('../src/engine.js');
const { generateDeal } = require('../src/generator.js');
const { termSheet } = require('../src/termsheet.js');
const { toHtml } = require('../src/markdown.js');

const SEEDS = Array.from({ length: 300 }, (_, i) => i + 1);
const deals = SEEDS.map((s) => {
  const deal = generateDeal(s);
  return { s, deal, res: engine.run(deal) };
});

test('same seed gives the same deal, different seeds differ', () => {
  assert.deepStrictEqual(generateDeal(4821), generateDeal(4821));
  assert.notDeepStrictEqual(generateDeal(1), generateDeal(2));
  assert.strictEqual(generateDeal(4821).meta.seed, 4821);
});

test('every generated deal runs and passes all integrity checks', () => {
  for (const { s, res } of deals) {
    assert.ok(res.ok, 'seed ' + s + ': ' + res.errors.join('; '));
    for (const c of res.checks) assert.ok(c.ok, 'seed ' + s + ': ' + c.name);
  }
});

test('generated deals are bankable under their own terms', () => {
  for (const { s, deal, res } of deals) {
    const x = res.summary, D = deal.senior_debt;
    assert.ok(x.effGearing <= D.gearing + 1e-9, 'seed ' + s + ' gearing');
    assert.ok(x.sculptDscr >= D.target_dscr - 1e-6, 'seed ' + s + ' sculpted DSCR below target');
    assert.ok(x.minDscrPost >= deal.covenants.lockup_dscr + 0.05 - 1e-6, 'seed ' + s + ' post-tax DSCR headroom');
    assert.strictEqual(x.lockupPeriods, 0, 'seed ' + s + ' lock-up in base case');
    assert.ok(x.equityIrr >= 0.05 && x.equityIrr <= 0.09, 'seed ' + s + ' equity IRR ' + x.equityIrr);
    assert.ok(deal.operations.ppa_price_fc >= 35 && deal.operations.ppa_price_fc <= 130, 'seed ' + s + ' price');
  }
});

test('every model input cites a term sheet clause', () => {
  const paths = ['timeline.financial_close', 'timeline.cod', 'timeline.ppa_end', 'timeline.debt_maturity',
    'construction.capex_total', 'construction.capex_profile', 'operations.capacity_mw', 'operations.capacity_factor_p50',
    'operations.availability', 'operations.degradation', 'operations.ppa_price_fc', 'operations.ppa_indexation',
    'operations.opex_real_fc', 'operations.cpi', 'tax.rate', 'tax.depreciation_years', 'senior_debt.amount',
    'senior_debt.gearing', 'senior_debt.base_rate', 'senior_debt.margin', 'senior_debt.upfront_fee',
    'senior_debt.commitment_fee_pct_margin', 'senior_debt.repayment', 'senior_debt.target_dscr', 'dsra.months_of_ds',
    'covenants.lockup_dscr', 'covenants.default_dscr'];
  for (const { s, deal } of deals.slice(0, 20)) {
    for (const p of paths) assert.ok(deal.sources[p], 'seed ' + s + ' has no source for ' + p);
  }
});

test('the term sheet states the numbers the model uses', () => {
  for (const { s, deal, res } of deals.slice(0, 50)) {
    const ts = termSheet(deal, res);
    const D = deal.senior_debt, O = deal.operations;
    const expect = [
      'AUD ' + (D.amount * 1000).toLocaleString('en-US'),
      'AUD ' + O.ppa_price_fc.toFixed(2) + '/MWh',
      'Margin | ' + (D.margin * 100).toFixed(2) + '%',
      D.target_dscr.toFixed(2) + 'x on P50 CFADS',
      (O.capacity_factor_p50 * 100).toFixed(1) + '%',
      'Deal #' + s,
    ];
    for (const e of expect) assert.ok(ts.includes(e), 'seed ' + s + ' term sheet is missing "' + e + '"');
  }
});

test('the term sheet renders to HTML with its tables', () => {
  const { deal, res } = deals[0];
  const html = toHtml(termSheet(deal, res));
  assert.ok(html.startsWith('<h1>'));
  assert.ok((html.match(/<table>/g) || []).length >= 6);
  assert.ok(!html.includes('|'), 'a table row was not converted');
});
