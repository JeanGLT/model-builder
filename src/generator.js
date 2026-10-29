/*
 * Random deal generator — v1
 * Draws a fictional Australian solar project finance deal from realistic ranges.
 * Same seed -> same deal, so a deal can be shared by its number ("Deal #4821").
 * The debt is sized like a real bank would: lowest of max gearing and the amount
 * sculpted at the target DSCR, rounded down to AUD 0.5m.
 * All names and parties are made up.
 */
(function (root) {
  'use strict';

  const engine = typeof module !== 'undefined' && module.exports ? require('./engine.js') : root.PFEngine;

  // Small seeded random number generator (mulberry32): fast, deterministic, good enough for sampling.
  function rng(seed) {
    let a = seed >>> 0;
    return function () {
      a = (a + 0x6d2b79f5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const NAMES = ['Banksia', 'Brolga', 'Mallee', 'Quandong', 'Bilby', 'Corella', 'Saltbush', 'Ironbark', 'Spinifex',
    'Bluegum', 'Kurrajong', 'Myall', 'Currawong', 'Boobook', 'Gidgee', 'Mulga', 'Paperbark', 'Bimble', 'Yarran', 'Wilga'];
  const SPONSORS = ['Ridgeline', 'Tallowwood', 'Greywater', 'Bluestone', 'Sandhill', 'Redgum', 'Stringybark', 'Glasshouse'];
  const FUNDS = ['I', 'II', 'III', 'IV'];
  const RATINGS = [['BBB-', 'Baa3'], ['BBB', 'Baa2'], ['BBB+', 'Baa1'], ['A-', 'A3']];
  // Capacity factor ranges are for single-axis trackers; fixed-tilt plants get 3 points less.
  const REGIONS = [
    { state: 'NSW', zone: 'Central-West Orana Renewable Energy Zone', cf: [0.25, 0.29], grid: 'NEM' },
    { state: 'NSW', zone: 'New England Renewable Energy Zone', cf: [0.25, 0.28], grid: 'NEM' },
    { state: 'NSW', zone: 'South-West Renewable Energy Zone', cf: [0.26, 0.29], grid: 'NEM' },
    { state: 'QLD', zone: 'Western Downs', cf: [0.27, 0.31], grid: 'NEM' },
    { state: 'QLD', zone: 'North Queensland', cf: [0.28, 0.31], grid: 'NEM' },
    { state: 'VIC', zone: 'Western Victoria Renewable Energy Zone', cf: [0.22, 0.25], grid: 'NEM' },
    { state: 'VIC', zone: 'Murray River Renewable Energy Zone', cf: [0.24, 0.27], grid: 'NEM' },
    { state: 'SA', zone: 'Riverland', cf: [0.25, 0.29], grid: 'NEM' },
    { state: 'WA', zone: 'Mid West', cf: [0.26, 0.30], grid: 'SWIS' },
  ];
  const FC_DATES = ['2027-01-01', '2027-07-01', '2028-01-01', '2028-07-01'];

  // Shapes of the capex drawdown (S-curve), one per construction length in semesters.
  const PROFILES = {
    3: [[0.25, 0.45, 0.30], [0.30, 0.45, 0.25], [0.20, 0.50, 0.30]],
    4: [[0.15, 0.35, 0.35, 0.15], [0.10, 0.30, 0.40, 0.20], [0.20, 0.35, 0.30, 0.15]],
  };

  const iso = (d) => d.toISOString().slice(0, 10);
  // First day of the month, n months after a YYYY-MM-01 date.
  function addMonths(s, n) {
    const [y, m] = s.split('-').map(Number);
    return iso(new Date(Date.UTC(y, m - 1 + n, 1)));
  }
  // Last day of the month before (date + n months): COD + 20 years -> 30 June.
  function endBefore(s, n) {
    const [y, m] = s.split('-').map(Number);
    return iso(new Date(Date.UTC(y, m - 1 + n, 0)));
  }

  function draw(r) {
    const pick = (arr) => arr[Math.floor(r() * arr.length)];
    // Uniform draw between lo and hi on a grid of `step`.
    const between = (lo, hi, step) => {
      const n = Math.round((hi - lo) / step);
      return +(lo + Math.floor(r() * (n + 1)) * step).toFixed(6);
    };

    const name = pick(NAMES);
    const region = pick(REGIONS);
    const tracker = r() < 0.75;
    const mw = between(50, 400, 10);
    const nC = r() < 0.4 ? 3 : 4;
    const fc = pick(FC_DATES);
    const cod = addMonths(fc, 6 * nC);
    const ppaYears = between(15, 20, 1);
    const tailSemesters = between(2, 4, 1);
    const ppaEnd = endBefore(cod, 12 * ppaYears);
    const maturity = endBefore(cod, 12 * ppaYears - 6 * tailSemesters);
    const indexed = r() < 0.7;
    const cf = +(between(region.cf[0], region.cf[1], 0.005) - (tracker ? 0 : 0.03)).toFixed(3);

    const deal = {
      meta: {
        name: 'Project ' + name, currency: 'AUD', units: 'k', generated: true,
        plant: name + ' Solar Farm', state: region.state, zone: region.zone, grid: region.grid,
        technology: tracker ? 'single-axis tracker' : 'fixed-tilt',
        sponsor: pick(SPONSORS) + ' Renewable Partners Fund ' + pick(FUNDS),
        offtaker_rating: pick(RATINGS), ppa_years: ppaYears,
      },
      timeline: { financial_close: fc, cod, ppa_end: ppaEnd, debt_maturity: maturity },
      construction: {
        capex_total: mw * between(1150, 1450, 10),
        capex_profile: pick(PROFILES[nC]).slice(),
      },
      operations: {
        capacity_mw: mw,
        capacity_factor_p50: cf,
        availability: between(0.98, 0.995, 0.005),
        degradation: between(0.004, 0.006, 0.0005),
        ppa_price_fc: 0, // solved below from the sponsor's target IRR
        ppa_indexation: indexed ? 1.0 : 0.0,
        opex_real_fc: mw * between(22, 32, 0.5),
        cpi: between(0.02, 0.03, 0.0025),
      },
      tax: { rate: 0.30, depreciation_years: pick([20, 25]) },
      senior_debt: {
        sizing: 'lower_of',
        amount: 0,
        gearing: between(0.65, 0.80, 0.05),
        base_rate: between(0.035, 0.045, 0.0025),
        margin: between(0.015, 0.025, 0.0005),
        upfront_fee: between(0.01, 0.02, 0.0025),
        commitment_fee_pct_margin: pick([0.35, 0.40, 0.45]),
        repayment: 'sculpted',
        target_dscr: indexed ? between(1.25, 1.35, 0.05) : between(1.30, 1.40, 0.05),
      },
      dsra: { months_of_ds: 6 },
      covenants: { lockup_dscr: pick([1.10, 1.15]), default_dscr: 1.05 },
    };

    // The PPA price is what makes a deal work for the sponsor: pick the lowest price on a
    // AUD 0.50 grid that gives an equity IRR at or above a drawn target.
    const targetIrr = between(0.055, 0.085, 0.0025);
    const irrAt = (price) => {
      deal.operations.ppa_price_fc = price;
      if (!size(deal)) return -1;
      const x = engine.run(deal);
      return x.ok && x.summary.equityIrr !== null ? x.summary.equityIrr : -1;
    };
    let lo = 70, hi = 260; // price steps of AUD 0.50: 35.00 to 130.00/MWh
    if (irrAt(hi / 2) < targetIrr) return null;
    while (hi - lo > 1) {
      const mid = Math.floor((lo + hi) / 2);
      if (irrAt(mid / 2) >= targetIrr) hi = mid; else lo = mid;
    }
    irrAt(hi / 2);
    deal.meta.target_irr = targetIrr;
    deal.sources = sources(deal);
    return deal;
  }

  // Sizes the debt like a lender: lowest of (i) max gearing, (ii) sculpted at the target DSCR on
  // pre-tax P50 CFADS, (iii) the amount keeping the minimum post-tax DSCR at least 0.05x above lock-up.
  // The result is rounded down to AUD 0.5m and stored as a fixed facility amount.
  function size(deal) {
    const D = deal.senior_debt;
    D.sizing = 'lower_of';
    const first = engine.run(deal);
    if (!first.ok) return false;
    let amount = first.summary.amount;
    let binding = first.summary.sizingBinding;
    const floor = deal.covenants.lockup_dscr + 0.05;
    D.sizing = 'fixed';
    const minPost = (a) => { D.amount = a; const x = engine.run(deal); return x.ok ? x.summary.minDscrPost : 0; };
    if (minPost(amount) < floor) {
      let a = 0, b = amount;
      for (let i = 0; i < 16; i++) { const m = (a + b) / 2; if (minPost(m) >= floor) a = m; else b = m; }
      amount = a;
      binding = 'post-tax DSCR';
    }
    D.amount = Math.floor(amount / 500) * 500;
    deal.meta.sizing_binding = binding;
    return D.amount > 0;
  }

  // Clause references, matching the sections written by termsheet.js.
  function sources(d) {
    const O = d.operations, D = d.senior_debt, T = d.timeline;
    const pct = (v, n = 2) => (v * 100).toFixed(n) + '%';
    return {
      'timeline.financial_close': 'Summary: financial close targeted ' + T.financial_close,
      'timeline.cod': 'Summary: COD ' + T.cod,
      'timeline.ppa_end': 'Section 1 · PPA: tenor ' + d.meta.ppa_years + ' years from COD, ending ' + T.ppa_end,
      'timeline.debt_maturity': 'Section 2 · Final Maturity: ' + T.debt_maturity,
      'construction.capex_total': 'Section 2 · Sources and uses: construction capex AUD ' + d.construction.capex_total.toLocaleString('en-US') + 'k',
      'construction.capex_profile': 'Section 2 · Capex drawn ' + d.construction.capex_profile.map((x) => Math.round(x * 100) + '%').join(' / '),
      'operations.capacity_mw': 'Section 6 · Base case: ' + O.capacity_mw + ' MWac',
      'operations.capacity_factor_p50': 'Section 6 · P50 capacity factor ' + pct(O.capacity_factor_p50, 1),
      'operations.availability': 'Section 6 · Availability ' + pct(O.availability, 1),
      'operations.degradation': 'Section 6 · Degradation ' + pct(O.degradation) + ' p.a. from operating year 2',
      'operations.ppa_price_fc': 'Section 1 · PPA: AUD ' + O.ppa_price_fc.toFixed(2) + '/MWh at financial close',
      'operations.ppa_indexation': O.ppa_indexation ? 'Section 1 · PPA: indexed at 100% of CPI annually' : 'Section 1 · PPA: flat price, no indexation',
      'operations.opex_real_fc': 'Section 6 · Operating costs AUD ' + (O.opex_real_fc / 1000).toFixed(3) + 'm p.a. real',
      'operations.cpi': 'Section 6 · CPI ' + pct(O.cpi) + ' p.a., applied annually',
      'tax.rate': 'Section 6 · Corporate tax 30%',
      'tax.depreciation_years': 'Section 6 · Straight-line tax depreciation over ' + d.tax.depreciation_years + ' years',
      'senior_debt.amount': 'Section 2 · Senior Facility AUD ' + (D.amount * 1000).toLocaleString('en-US'),
      'senior_debt.gearing': 'Section 4 · Sizing: maximum ' + pct(D.gearing, 0) + ' of construction capex',
      'senior_debt.base_rate': 'Section 3 · Base case BBSW ' + pct(D.base_rate),
      'senior_debt.margin': 'Section 3 · Margin ' + pct(D.margin) + ' p.a.',
      'senior_debt.upfront_fee': 'Section 3 · Upfront fee ' + pct(D.upfront_fee),
      'senior_debt.commitment_fee_pct_margin': 'Section 3 · Commitment fee ' + pct(D.commitment_fee_pct_margin, 0) + ' of the Margin',
      'senior_debt.repayment': 'Section 4 · Repayment sculpted',
      'senior_debt.target_dscr': 'Section 4 · Target DSCR ' + D.target_dscr.toFixed(2) + 'x on P50 CFADS',
      'dsra.months_of_ds': 'Section 4 · DSRA: next 6 months of scheduled debt service',
      'covenants.lockup_dscr': 'Section 4 · Lock-up < ' + d.covenants.lockup_dscr.toFixed(2) + 'x',
      'covenants.default_dscr': 'Section 4 · Event of Default < ' + d.covenants.default_dscr.toFixed(2) + 'x',
    };
  }

  // Public entry point. Retries with the same random stream until the deal passes every model check,
  // so a seed always gives a bankable deal.
  function generateDeal(seed) {
    const s = Math.max(1, Math.floor(Number(seed)) || 1);
    const r = rng(s);
    for (let attempt = 0; attempt < 50; attempt++) {
      const deal = draw(r);
      if (!deal) continue;
      const res = engine.run(deal);
      if (res.ok && res.checks.every((c) => c.ok) && res.summary.equityIrr !== null) {
        deal.meta.seed = s;
        return deal;
      }
    }
    throw new Error('Could not generate a valid deal for seed ' + s);
  }

  const randomSeed = () => 1 + Math.floor(Math.random() * 9999);

  const api = { generateDeal, randomSeed, rng };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFGenerator = api;
})(typeof window !== 'undefined' ? window : globalThis);
