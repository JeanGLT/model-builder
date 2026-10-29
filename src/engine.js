/*
 * Project finance engine — v1
 * Pure functions: deal JSON in, time series and ratios out. No UI, no LLM.
 * Units: thousands of currency (k). Rates are decimals. Periods are semi-annual (default) or quarterly
 * (deal.timeline.periodicity = 'quarterly').
 * See SPEC.md for the calculation order and the decisions taken where the term sheet is silent.
 */
(function (root) {
  'use strict';

  const MONTH = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function parseDate(s) {
    const [y, m, d] = s.split('-').map(Number);
    return { y, m, d };
  }

  // Months from the start of the financial close month to a date.
  // A date on or after the 28th counts as the end of its month (30 June -> full month).
  function monthsFrom(fc, s) {
    const t = parseDate(s);
    let months = (t.y - fc.y) * 12 + (t.m - fc.m);
    if (t.d >= 28) months += 1;
    return months;
  }

  function sum(a) { return a.reduce((x, y) => x + y, 0); }

  // Months in one model period: 6 (semi-annual, the default) or 3 (quarterly).
  function monthsPerPeriod(deal) { return deal.timeline.periodicity === 'quarterly' ? 3 : 6; }

  function irr(flows) {
    const npv = (r) => flows.reduce((acc, f, i) => acc + f / Math.pow(1 + r, i), 0);
    let lo = -0.99, hi = 1.0;
    if (npv(lo) * npv(hi) > 0) return null;
    for (let i = 0; i < 300; i++) {
      const mid = (lo + hi) / 2;
      if (npv(mid) > 0) lo = mid; else hi = mid;
    }
    return (lo + hi) / 2;
  }

  function validate(deal) {
    const errors = [];
    const p = deal.construction.capex_profile;
    if (Math.abs(sum(p) - 1) > 1e-6) errors.push('Capex profile must add up to 100% (now ' + (sum(p) * 100).toFixed(1) + '%).');
    const fc = parseDate(deal.timeline.financial_close);
    const pm = monthsPerPeriod(deal);
    // Dates and the capex profile are given in semesters, whatever the model periodicity.
    const nS = monthsFrom(fc, deal.timeline.cod) / 6;
    if (!Number.isInteger(nS) || nS <= 0) errors.push('COD must fall a whole number of semesters after financial close.');
    else if (p.length !== nS) errors.push('Capex profile has ' + p.length + ' entries but construction lasts ' + nS + ' semesters.');
    const nM = monthsFrom(fc, deal.timeline.debt_maturity) / 6;
    const nT = monthsFrom(fc, deal.timeline.ppa_end) / 6;
    if (!Number.isInteger(nM) || nM <= nS) errors.push('Debt maturity must be a semester end after COD.');
    if (!Number.isInteger(nT) || nT < nM) errors.push('PPA end must be a semester end on or after debt maturity.');
    const k = deal.dsra.months_of_ds / pm;
    if (!Number.isInteger(k) || k < 0) errors.push('DSRA months must be a multiple of ' + pm + '.');
    return errors;
  }

  function run(deal) {
    const errors = validate(deal);
    if (errors.length) return { ok: false, errors };

    const C = deal.construction, O = deal.operations, T = deal.tax, D = deal.senior_debt;
    const fc = parseDate(deal.timeline.financial_close);
    const pm = monthsPerPeriod(deal);                           // months per period
    const ppy = 12 / pm;                                        // periods per year
    const nC = monthsFrom(fc, deal.timeline.cod) / pm;          // construction periods
    const nM = monthsFrom(fc, deal.timeline.debt_maturity) / pm; // last debt period
    const N = monthsFrom(fc, deal.timeline.ppa_end) / pm;        // last period
    const r = D.base_rate + D.margin;
    const rs = r / ppy;                                         // interest rate per period

    // ---- Timeline -------------------------------------------------------
    const periods = [];
    for (let t = 1; t <= N; t++) {
      const endIdx = fc.m - 1 + pm * t - 1; // 0-based month index of period end
      const year = fc.y + Math.floor(endIdx / 12);
      const month = (endIdx % 12) + 1;
      const fy = month <= 6 ? year : year + 1;
      periods.push({
        t, label: MONTH[month - 1] + '-' + String(year).slice(2), year, month, fy,
        phase: t <= nC ? 'construction' : 'operations',
        opYear: t > nC ? Math.floor((t - nC - 1) / ppy) + 1 : 0,
      });
    }
    const fy1 = periods[0].fy;
    const z = () => new Array(N).fill(0);

    // ---- B. Operations (computed first: sizing depends on CFADS) ---------
    const energy = z(), revenue = z(), opex = z(), ebitda = z(), index = z();
    for (let i = nC; i < N; i++) {
      const p = periods[i];
      const idx = Math.pow(1 + O.cpi, p.fy - fy1);
      const deg = Math.pow(1 - O.degradation, p.opYear - 1);
      const annualE = O.capacity_mw * 8760 * O.capacity_factor_p50 * O.availability * deg;
      const priceIdx = 1 + (idx - 1) * O.ppa_indexation;
      index[i] = idx;
      energy[i] = annualE / ppy;
      revenue[i] = (annualE * O.ppa_price_fc * priceIdx / 1000) / ppy;
      opex[i] = (O.opex_real_fc * idx) / ppy;
      ebitda[i] = revenue[i] - opex[i];
    }
    const cfadsPre = ebitda.slice();

    // ---- Debt sizing ----------------------------------------------------
    const nRep = nM - nC;
    const df = []; // discount factor for each repayment period, from COD
    for (let k = 1; k <= nRep; k++) df.push(Math.pow(1 + rs, -k));
    const pvCfads = sum(df.map((f, k) => f * cfadsPre[nC + k]));
    const capacityAtTarget = pvCfads / D.target_dscr;
    const gearingCap = C.capex_total * D.gearing;
    let amount = D.amount;
    let sizingBinding = 'fixed input';
    if (D.sizing === 'lower_of') {
      amount = Math.min(gearingCap, capacityAtTarget);
      sizingBinding = gearingCap <= capacityAtTarget ? 'gearing' : 'DSCR';
    }
    const effGearing = amount / C.capex_total;

    // ---- A. Construction ------------------------------------------------
    const capex = z(), draw = z(), idc = z(), commit = z(), upfront = z(), equityIn = z();
    const debtOpen = z(), debtClose = z(), interest = z(), principal = z(), debtService = z();
    let bal = 0, undrawn = amount;
    for (let i = 0; i < nC; i++) {
      // The capex profile is per semester: in a quarterly model each quarter takes half its semester's share.
      capex[i] = C.capex_total * C.capex_profile[Math.floor(i * pm / 6)] * pm / 6;
      draw[i] = capex[i] * effGearing;
      debtOpen[i] = bal;
      idc[i] = bal * rs;
      commit[i] = undrawn * D.margin * D.commitment_fee_pct_margin / ppy;
      if (i === 0) upfront[i] = amount * D.upfront_fee;
      equityIn[i] = capex[i] - draw[i] + idc[i] + commit[i] + upfront[i];
      bal += draw[i];
      undrawn -= draw[i];
      debtClose[i] = bal;
    }
    const totalIdc = sum(idc), totalCommit = sum(commit), totalUpfront = sum(upfront);
    const projectCost = C.capex_total + totalIdc + totalCommit + totalUpfront;

    // ---- C. Repayment profile -------------------------------------------
    let sculptDscr = null;
    if (D.repayment === 'sculpted') {
      sculptDscr = pvCfads / amount;
      for (let k = 0; k < nRep; k++) debtService[nC + k] = cfadsPre[nC + k] / sculptDscr;
    } else if (D.repayment === 'annuity') {
      const pmt = amount * rs / (1 - Math.pow(1 + rs, -nRep));
      for (let k = 0; k < nRep; k++) debtService[nC + k] = pmt;
    }
    for (let i = nC; i < nM; i++) {
      debtOpen[i] = bal;
      interest[i] = bal * rs;
      if (D.repayment === 'linear') {
        principal[i] = amount / nRep;
        debtService[i] = principal[i] + interest[i];
      } else {
        principal[i] = debtService[i] - interest[i];
      }
      bal -= principal[i];
      debtClose[i] = bal;
    }
    const balanceAtMaturity = bal;

    // ---- D. Tax ---------------------------------------------------------
    const depreciation = z(), taxable = z(), lossUsed = z(), lossBal = z(), tax = z();
    const depPerPeriod = projectCost / T.depreciation_years / ppy;
    let losses = 0;
    for (let i = nC; i < N; i++) {
      if (i - nC < T.depreciation_years * ppy) depreciation[i] = depPerPeriod;
      taxable[i] = ebitda[i] - depreciation[i] - interest[i];
      if (taxable[i] < 0) {
        losses += -taxable[i];
      } else {
        const use = Math.min(losses, taxable[i]);
        losses -= use;
        lossUsed[i] = use;
        tax[i] = (taxable[i] - use) * T.rate;
      }
      lossBal[i] = losses;
    }
    const cfadsPost = ebitda.map((e, i) => e - tax[i]);

    // ---- G. Ratios ------------------------------------------------------
    const dscrPre = z().map(() => null), dscrPost = z().map(() => null);
    const llcrPre = z().map(() => null), llcrPost = z().map(() => null);
    const hDscr = z().map(() => null);
    for (let i = nC; i < nM; i++) {
      dscrPre[i] = cfadsPre[i] / debtService[i];
      dscrPost[i] = cfadsPost[i] / debtService[i];
      let pvPre = 0, pvPost = 0;
      for (let j = i; j < nM; j++) {
        const f = Math.pow(1 + rs, -(j - i + 1));
        pvPre += cfadsPre[j] * f;
        pvPost += cfadsPost[j] * f;
      }
      llcrPre[i] = pvPre / debtOpen[i];
      llcrPost[i] = pvPost / debtOpen[i];
      // Historic DSCR over the last 12 months (2 semesters or 4 quarters), once 12 months have passed.
      if (i >= nC + ppy - 1) {
        let cf = 0, ds = 0;
        for (let j = i - ppy + 1; j <= i; j++) { cf += cfadsPost[j]; ds += debtService[j]; }
        hDscr[i] = cf / ds;
      }
    }
    let pvPpa = 0;
    for (let j = nC; j < N; j++) pvPpa += cfadsPre[j] * Math.pow(1 + rs, -(j - nC + 1));
    const plcrCod = pvPpa / amount;

    // ---- E/F. DSRA and waterfall ----------------------------------------
    const kDs = deal.dsra.months_of_ds / pm;
    const dsraTarget = z();
    for (let i = nC - 1; i < N; i++) {
      let s = 0;
      for (let j = i + 1; j <= i + kDs && j < N; j++) s += debtService[j];
      dsraTarget[i] = s;
    }
    const dsraFund = z(), dsraTopUp = z(), dsraRelease = z(), dsraDraw = z(), dsraBal = z();
    const cashAfterDs = z(), trapped = z(), trapBal = z(), distributions = z();
    const lockup = z().map(() => false), defaultFlag = z().map(() => false), shortfall = z();
    dsraFund[nC - 1] = dsraTarget[nC - 1];
    dsraBal[nC - 1] = dsraFund[nC - 1];
    let dsra = dsraBal[nC - 1], trap = 0;
    const cov = deal.covenants;
    for (let i = nC; i < N; i++) {
      let cash = cfadsPost[i] - debtService[i];
      cashAfterDs[i] = cash;
      if (cash < 0) {
        const d = Math.min(dsra, -cash);
        dsraDraw[i] = d; dsra -= d; cash += d;
        if (cash < -1e-9) { shortfall[i] = -cash; cash = 0; }
      }
      const gap = dsraTarget[i] - dsra;
      if (gap > 0) { const u = Math.min(gap, cash); dsraTopUp[i] = u; dsra += u; cash -= u; }
      else if (gap < 0) { dsraRelease[i] = -gap; dsra += gap; cash -= gap; }
      dsraBal[i] = dsra;

      if (hDscr[i] !== null) {
        lockup[i] = hDscr[i] < cov.lockup_dscr;
        defaultFlag[i] = hDscr[i] < cov.default_dscr;
      }
      const dsraFull = dsra >= dsraTarget[i] - 1e-6;
      if (lockup[i] || !dsraFull) {
        trapped[i] = cash; trap += cash;
      } else {
        distributions[i] = cash + trap; trap = 0;
      }
      if (i === N - 1) { distributions[i] += trap + dsra; trap = 0; dsraRelease[i] += dsra; dsra = 0; dsraBal[i] = 0; }
      trapBal[i] = trap;
    }

    // ---- H. Equity returns ----------------------------------------------
    const equityFlows = z().map((_, i) => -equityIn[i] - dsraFund[i] + distributions[i]);
    const irrSemi = irr(equityFlows); // per period
    const equityIrr = irrSemi === null ? null : Math.pow(1 + irrSemi, ppy) - 1;

    // ---- Summary & checks ------------------------------------------------
    const debtIdx = [];
    for (let i = nC; i < nM; i++) debtIdx.push(i);
    const minOf = (arr) => Math.min(...debtIdx.map((i) => arr[i]));
    const avgOf = (arr) => sum(debtIdx.map((i) => arr[i])) / debtIdx.length;
    const minIdx = (arr) => debtIdx.reduce((b, i) => (arr[i] < arr[b] ? i : b), debtIdx[0]);
    const firstTax = tax.findIndex((x) => x > 0);
    const totalEquity = sum(equityIn);
    const year1 = (arr) => sum(arr.slice(nC, nC + ppy)); // first 12 months of operations

    const checks = [
      { name: 'Debt fully drawn at COD', ok: Math.abs(sum(draw) - amount) < 0.01, value: sum(draw) },
      { name: 'Debt repaid at maturity', ok: Math.abs(balanceAtMaturity) < 0.01, value: balanceAtMaturity },
      { name: 'Sources equal uses', ok: Math.abs(amount + totalEquity - projectCost) < 0.01, value: amount + totalEquity - projectCost },
      { name: 'Gearing within limit', ok: effGearing <= D.gearing + 1e-9, value: effGearing },
      { name: 'No debt service shortfall', ok: sum(shortfall) < 0.01, value: sum(shortfall) },
      { name: 'No principal re-borrowing', ok: debtIdx.every((i) => principal[i] >= -1e-9), value: Math.min(...debtIdx.map((i) => principal[i])) },
    ];

    return {
      ok: true,
      errors: [],
      meta: { nC, nM, N, nRep, rate: r, pm, ppy, firstOpLabel: periods[nC].label, maturityLabel: periods[nM - 1].label },
      periods,
      series: {
        capex, draw, idc, commit, upfront, equityIn, index,
        energy, revenue, opex, ebitda, cfadsPre, depreciation, taxable, lossUsed, lossBal, tax, cfadsPost,
        debtOpen, interest, principal, debtService, debtClose,
        dscrPre, dscrPost, hDscr, llcrPre, llcrPost,
        dsraTarget, dsraFund, dsraTopUp, dsraRelease, dsraDraw, dsraBal,
        cashAfterDs, trapped, trapBal, distributions, equityFlows, lockup, defaultFlag, shortfall,
      },
      summary: {
        amount, effGearing, sizingBinding,
        capex: C.capex_total, idc: totalIdc, commitmentFees: totalCommit, upfrontFee: totalUpfront,
        projectCost, equityAtFc: totalEquity, dsraAtCod: dsraFund[nC - 1],
        sculptDscr, capacityAtTarget, headroom: capacityAtTarget / amount - 1,
        minDscrPre: minOf(dscrPre), avgDscrPre: avgOf(dscrPre),
        minDscrPost: minOf(dscrPost), avgDscrPost: avgOf(dscrPost), minDscrPostLabel: periods[minIdx(dscrPost)].label,
        llcrCodPre: llcrPre[nC], llcrCodPost: llcrPost[nC], minLlcrPost: minOf(llcrPost), plcrCod,
        firstTaxLabel: firstTax >= 0 ? periods[firstTax].label : null,
        lockupPeriods: lockup.filter(Boolean).length, defaultPeriods: defaultFlag.filter(Boolean).length,
        equityIrr, totalDistributions: sum(distributions),
        firstOpYear: {
          energy: year1(energy), revenue: year1(revenue), opex: year1(opex), ebitda: year1(ebitda),
          margin: year1(ebitda) / year1(revenue),
        },
        firstDebtService: debtService[nC], firstPrincipal: principal[nC],
      },
      checks,
    };
  }

  const api = { run, validate, irr, version: '1.0.0' };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
