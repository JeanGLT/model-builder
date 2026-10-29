// Excel export: structure of the workbook and consistency with the engine.
// Formulas are recalculated independently in test/verify-excel.html (needs a browser).
// Run with: node --test
const test = require('node:test');
const assert = require('node:assert');
const engine = require('../src/engine.js');
const { buildWorkbook, exportXlsx } = require('../src/excel.js');
const { crc32 } = require('../src/xlsx.js');
const { generateDeal } = require('../src/generator.js');
const deal = require('../deal.json');

const res = engine.run(deal);
const wb = buildWorkbook(deal, res);
const sheet = (name) => wb.sheets.find((s) => s.name === name);
const cell = (ws, ref) => {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  const col = m[1].split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const row = ws.rows.get(Number(m[2]));
  return row && row.get(col);
};
const rowOf = (ws, label) => [...ws.rows.entries()].find(([, cells]) => cells.get(1) && cells.get(1).value === label)[0];

test('workbook has the fifteen sheets of a project finance model, in order', () => {
  assert.deepStrictEqual(wb.sheets.map((s) => s.name), ['Cover', 'Dashboard', 'Inputs', 'Scenarios', 'Timing', 'Operations', 'Construction', 'Debt', 'Tax',
    'Reserves', 'Waterfall', 'Ratios', 'Returns', 'Annual', 'Checks']);
});

test('every input is a named cell holding the deal value', () => {
  const names = new Map(wb.names);
  const inputs = sheet('Inputs');
  const expect = { capex_total: 140000, margin: 0.0175, target_dscr: 1.3, ppa_price: 60, cpi: 0.025, debt_amount_input: 98000, repayment: 'sculpted' };
  for (const [name, value] of Object.entries(expect)) {
    const ref = names.get(name);
    assert.ok(ref, 'missing name ' + name);
    assert.strictEqual(cell(inputs, ref.split('!')[1].replace(/\$/g, '')).value, value, name);
  }
  const [, a, b] = names.get('capex_profile').match(/\$B\$(\d+):\$B\$(\d+)/);
  assert.strictEqual(Number(b) - Number(a) + 1, 4, 'one row per construction semester');
});

test('calculation sheets are all formulas, with the engine values cached', () => {
  const N = res.meta.N;
  const check = (label, key, name) => {
    const calc = sheet(name);
    const r = rowOf(calc, label);
    for (let t = 1; t <= N; t++) {
      const c = calc.rows.get(r).get(3 + t);
      assert.ok(c && typeof c.value === 'object' && c.value.f, label + ' period ' + t + ' is not a formula');
      const v = res.series[key][t - 1];
      if (typeof v === 'number') assert.ok(Math.abs(c.value.v - v) < 1e-9, label + ' period ' + t);
    }
  };
  check('Revenue', 'revenue', 'Operations');
  check('Agency fees', 'agency', 'Operations');
  check('Debt service', 'debtService', 'Debt');
  check('Cash tax', 'tax', 'Tax');
  check('DSCR post-tax', 'dscrPost', 'Ratios');
  check('Closing balance', 'debtClose', 'Debt');
  check('DSRA closing balance', 'dsraBal', 'Reserves');
  check('MRA closing balance', 'mraBal', 'Reserves');
  check('5. Distributions to equity', 'distributions', 'Waterfall');
  check('Equity cash flow', 'equityFlows', 'Returns');
});

test('formulas only use functions every spreadsheet app supports', () => {
  const allowed = new Set(['IF', 'AND', 'OR', 'SUM', 'SUMIF', 'SUMPRODUCT', 'MIN', 'MAX', 'ABS', 'INT', 'YEAR', 'MONTH', 'EDATE', 'INDEX', 'COUNT', 'AVERAGE', 'IRR']);
  for (const ws of wb.sheets) {
    for (const cells of ws.rows.values()) {
      for (const c of cells.values()) {
        if (!c.value || typeof c.value !== 'object') continue;
        for (const fn of c.value.f.match(/[A-Z]+(?=\()/g) || []) assert.ok(allowed.has(fn), ws.name + ': ' + fn + ' in ' + c.value.f);
      }
    }
  }
});

test('the .xlsx file is a valid zip with the expected parts', () => {
  const bytes = exportXlsx(deal, res);
  assert.strictEqual(String.fromCharCode(bytes[0], bytes[1]), 'PK');
  const text = Buffer.from(bytes).toString('latin1');
  for (const part of ['[Content_Types].xml', 'xl/workbook.xml', 'xl/styles.xml', 'xl/worksheets/sheet15.xml']) assert.ok(text.includes(part), part);
  assert.ok(text.includes('fullCalcOnLoad="1"'), 'Excel must recalculate on open');
  assert.strictEqual(crc32(new TextEncoder().encode('123456789')), 0xcbf43926); // standard CRC-32 check value
});

test('random deals export too', () => {
  for (const seed of [1, 4821, 9999]) {
    const d = generateDeal(seed);
    const bytes = exportXlsx(d, engine.run(d));
    assert.ok(bytes.length > 100000, 'seed ' + seed);
  }
});

test('a quarterly export has one column per quarter and quarterly formulas', () => {
  const q = Object.assign({}, deal, { timeline: Object.assign({}, deal.timeline, { periodicity: 'quarterly' }) });
  const qr = engine.run(q);
  const qwb = buildWorkbook(q, qr);
  const names = new Map(qwb.names);
  const inputs = qwb.sheets.find((s) => s.name === 'Inputs');
  assert.strictEqual(cell(inputs, names.get('months_per_period').split('!')[1].replace(/\$/g, '')).value, 3);
  const ops = qwb.sheets.find((s) => s.name === 'Operations');
  const r = rowOf(ops, 'Revenue');
  assert.strictEqual(ops.rows.get(r).size, 3 + qr.meta.N, 'label, unit, total and one cell per quarter');
  const ratios = qwb.sheets.find((s) => s.name === 'Ratios');
  const h = ratios.rows.get(rowOf(ratios, '12-month historic DSCR (tested)')).get(3 + 20).value.f;
  assert.strictEqual((h.match(/\+/g) || []).length, 6, '4 quarters of CFADS and of debt service');
});

test('the scenario switch and the sizing constraints are live formulas', () => {
  const names = new Map(wb.names);
  for (const n of ['scenario', 'sc_no', 'sc_energy', 'capacity_p99', 'gearing_cap', 'capacity_at_target', 'mra_target', 'swap_rate', 'cash_multiple'])
    assert.ok(names.has(n), 'missing name ' + n);
  const debt = sheet('Debt');
  const f = debt.rows.get(rowOf(debt, 'Senior debt amount')).get(3).value.f;
  assert.ok(/MIN\(gearing_cap,capacity_at_target,capacity_p99\)/.test(f), f);
  // A downside export carries the downside results and still reconciles.
  const dn = Object.assign({}, deal, { scenario: 'downside' });
  const dr = engine.run(dn);
  const dwb = buildWorkbook(dn, dr);
  const inputs = dwb.sheets.find((s) => s.name === 'Inputs');
  assert.strictEqual(cell(inputs, new Map(dwb.names).get('scenario').split('!')[1].replace(/\$/g, '')).value, 'Downside');
  assert.ok(dr.summary.minDscrPost < res.summary.minDscrPost, 'the downside is worse than the base case');
  assert.strictEqual(dr.summary.amount, res.summary.amount, 'same debt in every case');
});

test('the annual sheet has one column per financial year', () => {
  const a = sheet('Annual');
  const years = new Set(res.periods.map((p) => p.fy));
  assert.strictEqual(a.rows.get(4).size, 2 + years.size);
});
