// Golden tests: Project Coolabah term sheet (SPEC.md section 6).
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const engine = require('../src/engine.js');
// The golden values were recomputed by hand for the v1 scope of the term sheet: no hedging,
// agency fees or maintenance reserve. v1() strips those terms; their own tests are further down.
const full = require('../deal.json');
function v1(d) {
  const x = JSON.parse(JSON.stringify(d));
  delete x.maintenance; delete x.operations.p90_factor; delete x.operations.p99_factor;
  for (const k of ['agency_fee', 'hedge_ratio', 'hedge_years', 'swap_rate']) delete x.senior_debt[k];
  return x;
}
const deal = v1(full);

const r = engine.run(deal);
const s = r.summary;
const near = (actual, expected, tol, label) =>
  assert.ok(Math.abs(actual - expected) <= tol, `${label}: expected ${expected} ±${tol}, got ${actual}`);
const clone = () => JSON.parse(JSON.stringify(deal));

test('model runs without input errors', () => {
  assert.ok(r.ok, r.errors.join('; '));
  assert.strictEqual(r.meta.nC, 4);
  assert.strictEqual(r.meta.nRep, 36);
  assert.strictEqual(r.meta.N, 44);
});

test('A. construction and sources & uses', () => {
  near(s.idc, 4226, 1, 'IDC');
  near(s.commitmentFees, 858, 1, 'Commitment fees');
  near(s.upfrontFee, 1470, 0.001, 'Upfront fee');
  near(s.projectCost, 146554, 2, 'Total project cost');
  near(s.equityAtFc, 48554, 2, 'Equity at FC');
});

test('B. first operating year (FY30)', () => {
  near(s.firstOpYear.energy, 234155, 1, 'Net energy MWh');
  near(s.firstOpYear.revenue, 14760, 1, 'Revenue');
  near(s.firstOpYear.opex, 3362, 1, 'Opex');
  near(s.firstOpYear.ebitda, 11399, 1, 'EBITDA');
});

test('C. sculpting and debt capacity', () => {
  near(s.sculptDscr, 1.475, 0.002, 'Sculpted pre-tax DSCR');
  near(s.llcrCodPre, 1.475, 0.002, 'LLCR at COD');
  near(s.llcrCodPost, 1.417, 0.002, 'LLCR at COD (post-tax)');
  near(s.plcrCod, 1.584, 0.002, 'PLCR at COD');
  near(s.capacityAtTarget, 111206, 10, 'Capacity at 1.30x');
  near(s.firstDebtService, 3863, 2, 'First debt service');
  near(s.firstPrincipal, 1046, 2, 'First principal');
});

test('D. tax', () => {
  near(s.minDscrPost, 1.25, 0.01, 'Min post-tax DSCR');
  assert.strictEqual(s.firstTaxLabel, 'Jun-38');
});

test('E–H. DSRA, covenants, equity', () => {
  near(s.dsraAtCod, 3863, 2, 'DSRA at COD');
  assert.strictEqual(s.lockupPeriods, 0);
  near(s.equityIrr, 0.0515, 0.001, 'Equity IRR');
});

test('all integrity checks pass', () => {
  for (const c of r.checks) assert.ok(c.ok, c.name);
});

test('sizing "lower_of" picks the binding constraint', () => {
  const d = clone();
  d.senior_debt.sizing = 'lower_of';
  const g = engine.run(d);
  assert.strictEqual(g.summary.sizingBinding, 'gearing');
  near(g.summary.amount, 98000, 0.01, 'Gearing-bound debt');

  d.operations.ppa_price_fc = 50; // weaker cash flows: DSCR now binds
  const h = engine.run(d);
  assert.strictEqual(h.summary.sizingBinding, 'DSCR');
  near(h.summary.sculptDscr, 1.30, 1e-6, 'DSCR at target when DSCR binds');
  assert.ok(h.checks.every((c) => c.ok));
});

test('annuity and linear profiles repay the debt', () => {
  for (const rep of ['annuity', 'linear']) {
    const d = clone();
    d.senior_debt.repayment = rep;
    const x = engine.run(d);
    assert.ok(x.checks.find((c) => c.name === 'Debt repaid at maturity').ok, rep);
  }
});

test('bad inputs are rejected with a clear message', () => {
  const d = clone();
  d.construction.capex_profile = [0.2, 0.3, 0.3];
  const x = engine.run(d);
  assert.strictEqual(x.ok, false);
  assert.ok(x.errors.some((e) => e.includes('100%')));
});

test('quarterly periods give the same deal on a finer grid', () => {
  const base = full;
  const semi = engine.run(base);
  const q = engine.run(Object.assign({}, base, { timeline: Object.assign({}, base.timeline, { periodicity: 'quarterly' }) }));
  assert.ok(q.ok, q.errors && q.errors.join('; '));
  assert.strictEqual(q.meta.N, semi.meta.N * 2);
  assert.strictEqual(q.periods[0].label, 'Sep-27');
  assert.ok(q.checks.every((c) => c.ok), q.checks.filter((c) => !c.ok).map((c) => c.name).join(', '));
  // Same capex and the same first-year operations, only cut in quarters.
  assert.ok(Math.abs(q.series.capex.reduce((a, b) => a + b, 0) - 140000) < 1e-6);
  assert.ok(Math.abs(q.summary.firstOpYear.ebitda - semi.summary.firstOpYear.ebitda) < 1e-6);
  // Interest compounds quarterly, so IDC is a little higher; ratios stay close.
  assert.ok(q.summary.idc > semi.summary.idc);
  assert.ok(Math.abs(q.summary.minDscrPost - semi.summary.minDscrPost) < 0.02);
  assert.ok(Math.abs(q.summary.equityIrr - semi.summary.equityIrr) < 0.002);
});

test('random deals pass every check when modelled quarterly', () => {
  const { generateDeal } = require('../src/generator.js');
  for (let seed = 1; seed <= 100; seed++) {
    const d = generateDeal(seed);
    d.timeline.periodicity = 'quarterly';
    const r = engine.run(d);
    assert.ok(r.ok, 'seed ' + seed + ': ' + (r.errors || []).join('; '));
    assert.ok(r.checks.every((c) => c.ok), 'seed ' + seed + ': ' + r.checks.filter((c) => !c.ok).map((c) => c.name).join(', '));
  }
});
