// Golden tests: Project Coolabah term sheet (SPEC.md section 6).
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const engine = require('../src/engine.js');
const deal = require('../deal.json');

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
