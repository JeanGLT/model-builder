/*
 * Excel export — v2
 * Rebuilds the whole model as live Excel formulas, laid out like a bank's project finance model:
 *   Cover · Dashboard · Inputs · Scenarios · Timing · Operations · Construction · Debt · Tax ·
 *   Reserves · Waterfall · Ratios · Returns · Annual · Checks
 * Every calculation sheet has the same period columns (D = period 1), so a formula can point at
 * the same column of another sheet. Change an input (or the scenario) and everything recalculates.
 * Each formula cell also carries the engine's value as a cached result; the Checks sheet compares
 * the two. Calculation order follows engine.js (see SPEC.md section 5).
 */
(function (root) {
  'use strict';

  const node = typeof module !== 'undefined' && module.exports;
  const X = node ? require('./xlsx.js') : root.PFXlsx;
  const E = node ? require('./engine.js') : root.PFEngine;

  const C_ACCENT = '2F5BEA', C_INPUT_BG = 'EEF2FE', C_MUTED = '6B7280', C_HEAD_BG = 'F6F7F9', C_GOOD = '15803D';
  const C_NAVY = '1B2333', C_GOOD_BG = 'DCFCE7', C_CONSTR_BG = 'FFF4E0';
  const F_K = '#,##0;(#,##0);"–"', F_X = '0.000"x"', F_PCT = '0.00%', F_PCT1 = '0.0%';
  const F_DATE = 'dd-mmm-yyyy', F_MONTH = 'mmm-yy', F_INT = '0', F_PRICE = '#,##0.00', F_RATE = '0.0000%';
  const SCENARIO_NAMES = ['Base case (P50)', 'P90 (one-year)', 'P99 (one-year)', 'Downside'];
  const SCENARIO_KEYS = ['base', 'p90', 'p99', 'downside'];

  function buildWorkbook(deal, res) {
    const wb = new X.Workbook();
    const st = {
      h: wb.style({ bold: true, size: 11, color: C_ACCENT }),
      note: wb.style({ italic: true, color: C_MUTED }),
      noteWrap: wb.style({ color: C_MUTED, wrap: true }),
      bold: wb.style({ bold: true }),
      head: wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom' }),
      headR: wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'right' }),
      headC: wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'center' }),
      muted: wb.style({ color: C_MUTED }),
      banner: wb.style({ bold: true, size: 14, color: 'FFFFFF', fill: C_NAVY, valign: 'center' }),
      bannerFill: wb.style({ fill: C_NAVY }),
      group: wb.style({ bold: true, color: C_ACCENT, fill: C_INPUT_BG }),
      ok: wb.style({ bold: true, align: 'center', color: C_GOOD, fill: C_GOOD_BG }),
    };
    const inp = (numFmt) => wb.style({ color: C_ACCENT, fill: C_INPUT_BG, numFmt });
    const calc = (numFmt, bold) => wb.style(bold ? { numFmt, bold: true } : { numFmt });
    // Dark title band across the top of a sheet, with a one-line note under it.
    const banner = (ws, title, note, lastCol) => {
      ws.height(1, 30);
      ws.set('A1', '  ' + title, st.banner);
      for (let c = 2; c <= lastCol; c++) ws.set(X.colName(c) + '1', '', st.bannerFill);
      ws.set('A2', note, st.note);
    };

    const S = res.series, P = res.periods, N = res.meta.N, sum = res.summary;
    const nC = res.meta.nC, nM = res.meta.nM;
    const PM = res.meta.pm, PPY = res.meta.ppy;
    const PERIOD_WORD = PM === 3 ? 'quarterly' : 'semi-annual';
    const FIRST = 4; // column D holds period 1
    const col = (t) => X.colName(FIRST + t - 1);
    const LAST = col(N);
    const T = deal.timeline, Cn = deal.construction, O = deal.operations, D = deal.senior_debt;
    const MR = Object.assign({ cost_real: 0, from_year: 10, to_year: 12, funding_years: 3 }, deal.maintenance || {});
    const DN = Object.assign({ opex_uplift: 0.10, rate_shock: 0.01 }, deal.downside || {});
    const scIndex = Math.max(0, SCENARIO_KEYS.indexOf(deal.scenario || 'base'));
    const sumOf = (a) => a.reduce((x, y) => x + (y || 0), 0);

    // ---------------- Inputs ----------------
    const wsI = wb.addSheet('Inputs');
    wsI.tab = C_ACCENT;
    wsI.width(1, 44).width(2, 18).width(3, 16).width(4, 80);
    banner(wsI, 'Inputs', 'Blue cells are inputs: change them and the whole model recalculates. Column D gives the term sheet clause each input comes from.', 4);
    let r = 4;
    const src = (path) => (deal.sources && deal.sources[path]) || '';
    const section = (title) => { r++; wsI.set('A' + r, title, st.h); r++; };
    const input = (name, label, value, fmt, unit, path, list) => {
      wsI.set('A' + r, label);
      wsI.set('B' + r, value, inp(fmt));
      wsI.set('C' + r, unit || '', st.muted);
      wsI.set('D' + r, /^[a-z_]+\.[a-z_0-9]+$/.test(path || '') ? src(path) : (path || ''), st.muted);
      if (list) wsI.dropdown('B' + r, list);
      wb.defineName(name, 'Inputs!$B$' + r);
      r++;
    };
    wsI.set('A' + r, 'Input', st.head); wsI.set('B' + r, 'Value', st.headR); wsI.set('C' + r, 'Unit', st.head); wsI.set('D' + r, 'Term sheet source', st.head); r++;
    section('Scenario');
    input('scenario', 'Active scenario', SCENARIO_NAMES[scIndex], '', 'drop-down', 'Model choice. The debt is always sized and sculpted on the base case.', SCENARIO_NAMES);
    input('downside_opex', 'Downside: operating cost uplift', DN.opex_uplift, F_PCT1, '', 'Model choice (lender downside case)');
    input('downside_rate', 'Downside: BBSW shock on the unhedged debt', DN.rate_shock, F_PCT, '', 'Model choice (lender downside case)');
    section('Timeline');
    input('months_per_period', 'Months per period (' + PERIOD_WORD + ')', PM, F_INT, 'fixed at export', 'timeline.periodicity');
    input('fc_date', 'Financial close', X.excelDate(T.financial_close), F_DATE, '', 'timeline.financial_close');
    input('cod_date', 'Commercial operation date (COD)', X.excelDate(T.cod), F_DATE, '', 'timeline.cod');
    input('maturity_date', 'Final maturity', X.excelDate(T.debt_maturity), F_DATE, '', 'timeline.debt_maturity');
    input('ppa_end_date', 'PPA expiry', X.excelDate(T.ppa_end), F_DATE, '', 'timeline.ppa_end');
    section('Construction');
    input('capex_total', 'Construction capex', Cn.capex_total, F_K, deal.meta.currency + ' k', 'construction.capex_total');
    const profStart = r;
    Cn.capex_profile.forEach((x, i) => {
      wsI.set('A' + r, 'Capex share, construction semester ' + (i + 1));
      wsI.set('B' + r, x, inp(F_PCT1));
      wsI.set('C' + r, '% of capex', st.muted);
      if (i === 0) wsI.set('D' + r, src('construction.capex_profile'), st.muted);
      r++;
    });
    wb.defineName('capex_profile', 'Inputs!$B$' + profStart + ':$B$' + (r - 1));
    section('Operations');
    input('capacity_mw', 'Capacity', O.capacity_mw, '#,##0.0', 'MWac', 'operations.capacity_mw');
    input('cf_p50', 'P50 capacity factor', O.capacity_factor_p50, F_PCT, '', 'operations.capacity_factor_p50');
    input('p90_factor', 'One-year P90 energy (share of P50)', O.p90_factor || 1, F_PCT1, '', 'operations.p90_factor');
    input('p99_factor', 'One-year P99 energy (share of P50)', O.p99_factor || 1, F_PCT1, '', 'operations.p99_factor');
    input('availability', 'Availability', O.availability, F_PCT, '', 'operations.availability');
    input('degradation', 'Degradation', O.degradation, F_PCT, '% p.a.', 'operations.degradation');
    input('ppa_price', 'PPA price at financial close', O.ppa_price_fc, F_PRICE, deal.meta.currency + '/MWh', 'operations.ppa_price_fc');
    input('ppa_indexation', 'PPA indexation', O.ppa_indexation, '0%', '% of CPI', 'operations.ppa_indexation');
    input('opex_real', 'Operating costs (real, at financial close)', O.opex_real_fc, F_K, 'k p.a.', 'operations.opex_real_fc');
    input('cpi', 'CPI', O.cpi, F_PCT, '% p.a.', 'operations.cpi');
    section('Senior debt');
    input('sizing', 'Debt sizing', D.sizing === 'lower_of' ? 'lower_of' : 'fixed', '', 'fixed / lower_of', 'senior_debt.sizing', ['fixed', 'lower_of']);
    input('debt_amount_input', 'Facility amount (used when sizing = fixed)', D.amount, F_K, deal.meta.currency + ' k', 'senior_debt.amount');
    input('max_gearing', 'Maximum gearing (of capex)', D.gearing, F_PCT1, '', 'senior_debt.gearing');
    input('base_rate', 'Base rate (BBSW)', D.base_rate, F_PCT, '', 'senior_debt.base_rate');
    input('margin', 'Margin', D.margin, F_PCT, '', 'senior_debt.margin');
    input('upfront_fee_pct', 'Upfront fee', D.upfront_fee, F_PCT, '% of facility', 'senior_debt.upfront_fee');
    input('commit_fee_pct', 'Commitment fee', D.commitment_fee_pct_margin, F_PCT1, '% of margin', 'senior_debt.commitment_fee_pct_margin');
    input('agency_fee', 'Agency fees (facility agent and security trustee)', D.agency_fee || 0, F_K, 'k p.a.', 'senior_debt.agency_fee');
    input('repayment', 'Repayment profile', D.repayment, '', 'sculpted / annuity / linear', 'senior_debt.repayment', ['sculpted', 'annuity', 'linear']);
    input('target_dscr', 'Target DSCR (P50)', D.target_dscr, '0.00"x"', '', 'senior_debt.target_dscr');
    section('Interest rate hedging');
    input('hedge_ratio', 'Share of the debt swapped to fixed', D.hedge_ratio || 0, F_PCT1, '', 'senior_debt.hedge_ratio');
    input('hedge_years', 'Hedging tenor from financial close', D.hedge_years || 0, F_INT, 'years', 'senior_debt.hedge_years');
    input('swap_rate', 'Swap rate (replaces BBSW on the hedged share)', D.swap_rate == null ? D.base_rate : D.swap_rate, F_PCT, '', 'senior_debt.swap_rate');
    section('Tax');
    input('tax_rate', 'Corporate tax rate', deal.tax.rate, F_PCT1, '', 'tax.rate');
    input('dep_years', 'Tax depreciation (straight line)', deal.tax.depreciation_years, F_INT, 'years', 'tax.depreciation_years');
    section('Reserves and covenants');
    input('dsra_months', 'DSRA size', deal.dsra.months_of_ds, F_INT, 'months of debt service', 'dsra.months_of_ds');
    input('mra_cost', 'Inverter replacement (MRA), real at financial close', MR.cost_real, F_K, deal.meta.currency + ' k', 'maintenance.cost_real');
    input('mra_from', 'Replacement from operating year', MR.from_year, F_INT, '', 'maintenance.from_year');
    input('mra_to', 'Replacement to operating year', MR.to_year, F_INT, '', 'maintenance.to_year');
    input('mra_years', 'MRA funded over the years before', MR.funding_years, F_INT, 'years', 'maintenance.funding_years');
    input('lockup_dscr', 'Lock-up DSCR (12-month historic)', deal.covenants.lockup_dscr, '0.00"x"', '', 'covenants.lockup_dscr');
    input('default_dscr', 'Event of default DSCR', deal.covenants.default_dscr, '0.00"x"', '', 'covenants.default_dscr');

    section('Derived (formulas, do not edit)');
    const derived = (name, label, f, v, fmt, unit) => {
      wsI.set('A' + r, label); wsI.set('B' + r, { f, v }, calc(fmt)); wsI.set('C' + r, unit || '', st.muted);
      wb.defineName(name, 'Inputs!$B$' + r); r++;
    };
    derived('all_in_rate', 'All-in interest rate, unhedged', 'base_rate+margin', res.meta.rate, F_PCT, 'p.a.');
    derived('periods_per_year', 'Periods per year', '12/months_per_period', PPY, F_INT, '');
    derived('period_rate', 'Unhedged interest rate per period (annuity profile)', 'all_in_rate/periods_per_year', res.meta.rate / PPY, F_RATE, '');
    derived('n_constr', 'Construction periods', '((YEAR(cod_date)-YEAR(fc_date))*12+MONTH(cod_date)-MONTH(fc_date))/months_per_period', nC, F_INT, '');
    derived('fy_first', 'Financial year of period 1 (Jul–Jun)', 'YEAR(EDATE(fc_date,months_per_period)-1)+IF(MONTH(EDATE(fc_date,months_per_period)-1)>6,1,0)', P[0].fy, F_INT, '');
    derived('dsra_periods', 'DSRA size in periods', 'dsra_months/months_per_period', deal.dsra.months_of_ds / PM, F_INT, '');

    // ---------------- Scenarios ----------------
    const wsS = wb.addSheet('Scenarios');
    wsS.tab = C_ACCENT;
    wsS.width(1, 30).width(2, 16).width(3, 16).width(4, 16).width(5, 60);
    banner(wsS, 'Scenarios', 'Lender cases. Pick the active one on the Inputs sheet. Sizing and sculpting always use the base case: the other cases run against that repayment schedule.', 5);
    wsS.set('A4', 'Scenario', st.head); wsS.set('B4', 'Energy (× P50)', st.headR); wsS.set('C4', 'Opex (× base)', st.headR); wsS.set('D4', 'BBSW shock', st.headR); wsS.set('E4', 'What it tests', st.head);
    const scRows = [
      [SCENARIO_NAMES[0], 1, 1, 0, 'The Agreed Base Case: sizes and sculpts the debt.'],
      [SCENARIO_NAMES[1], { f: 'p90_factor', v: O.p90_factor || 1 }, 1, 0, 'A bad year of sun (90% probability of exceeding), every year.'],
      [SCENARIO_NAMES[2], { f: 'p99_factor', v: O.p99_factor || 1 }, 1, 0, 'A very bad year of sun (99% probability of exceeding), every year.'],
      [SCENARIO_NAMES[3], { f: 'p90_factor', v: O.p90_factor || 1 }, { f: '1+downside_opex', v: 1 + DN.opex_uplift }, { f: 'downside_rate', v: DN.rate_shock }, 'P90 energy, higher operating costs and higher rates together.'],
    ];
    scRows.forEach(([name, e, o, rt, what], i) => {
      const row = 5 + i;
      wsS.set('A' + row, name, st.bold);
      wsS.set('B' + row, e, calc(F_PCT1)); wsS.set('C' + row, o, calc(F_PCT1)); wsS.set('D' + row, rt, calc(F_PCT));
      wsS.set('E' + row, what, st.muted);
    });
    const scVals = scRows.map(([, e, o, rt]) => [typeof e === 'object' ? e.v : e, typeof o === 'object' ? o.v : o, typeof rt === 'object' ? rt.v : rt]);
    wsS.set('A10', 'Active scenario', st.h);
    const activeRow = (row, label, name, f, v, fmt) => { wsS.set('A' + row, label); wsS.set('B' + row, { f, v }, calc(fmt, true)); wb.defineName(name, 'Scenarios!$B$' + row); };
    activeRow(11, 'Scenario number', 'sc_no', 'IF(scenario="' + SCENARIO_NAMES[1] + '",2,IF(scenario="' + SCENARIO_NAMES[2] + '",3,IF(scenario="' + SCENARIO_NAMES[3] + '",4,1)))', scIndex + 1, F_INT);
    activeRow(12, 'Energy factor', 'sc_energy', 'INDEX(Scenarios!$B$5:$B$8,sc_no)', scVals[scIndex][0], F_PCT1);
    activeRow(13, 'Operating cost factor', 'sc_opex', 'INDEX(Scenarios!$C$5:$C$8,sc_no)', scVals[scIndex][1], F_PCT1);
    activeRow(14, 'BBSW shock', 'sc_rate', 'INDEX(Scenarios!$D$5:$D$8,sc_no)', scVals[scIndex][2], F_PCT);

    // ---------------- Period sheets ----------------
    // Rows are declared once per sheet with a formula template; row numbers are assigned first so
    // that any row can refer to any other, on any sheet.
    const SHEETS = [];
    let cur = null;
    const sheet = (name, note, tab) => { cur = { name, note, tab, rows: [] }; SHEETS.push(cur); };
    const G = (title) => cur.rows.push({ group: title });
    const L = (key, label, unit, fmt, f, opt) => cur.rows.push(Object.assign({ key, label, unit, fmt, f }, opt || {}));
    const V = (key, label, unit, fmt, f, opt) => cur.rows.push(Object.assign({ key, label, unit, fmt, f, scalar: true }, opt || {}));

    // Base-case and P99 PVs are not engine series: rebuild them for the cached values.
    const pvBack = (cf, rt, last) => { const out = new Array(N).fill(0); let pv = 0; for (let i = last - 1; i >= nC; i--) { pv = (cf[i] + pv) / (1 + rt[i]); out[i] = pv; } return out; };
    const pvBaseS = pvBack(S.sculptCfads, S.baseRate, nM), pvP99S = pvBack(S.cfadsP99, S.baseRate, nM);
    const pvPreS = pvBack(S.cfadsPre, S.rate, nM), pvPostS = pvBack(S.cfadsPost, S.rate, nM), pvPpaS = pvBack(S.cfadsPre, S.rate, N);
    const debtIdx = (i) => i >= nC && i < nM;

    sheet('Timing', 'Period dates and flags used by every other sheet.', C_MUTED);
    L('start', 'Period start', '', F_DATE, (c) => 'EDATE(fc_date,months_per_period*(' + c.t + '-1))', { cache: (i) => X.excelDate(periodStart(P[i], PM)) });
    L('end', 'Period end', '', F_DATE, (c) => 'EDATE(fc_date,months_per_period*' + c.t + ')-1', { cache: (i) => X.excelDate(periodEnd(P[i])) });
    G('Phase flags');
    L('f_con', 'Construction', 'flag', F_INT, (c) => 'IF(' + c('end') + '<cod_date,1,0)', { cache: (i) => (i < nC ? 1 : 0) });
    L('f_cod', 'COD (last construction period)', 'flag', F_INT, (c) => 'IF(AND(' + c('f_con') + '=1,' + c.n('f_con') + '=0),1,0)', { cache: (i) => (i === nC - 1 ? 1 : 0) });
    L('f_op', 'Operations', 'flag', F_INT, (c) => 'IF(AND(' + c('start') + '>=cod_date,' + c('end') + '<=ppa_end_date),1,0)', { cache: (i) => (i >= nC ? 1 : 0) });
    L('f_first', 'First operating period', 'flag', F_INT, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.p('f_op') + '=0),1,0)', { cache: (i) => (i === nC ? 1 : 0) });
    L('f_last', 'Last operating period', 'flag', F_INT, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.n('f_op') + '=0),1,0)', { cache: (i) => (i === N - 1 ? 1 : 0) });
    L('f_debt', 'Debt repayment', 'flag', F_INT, (c) => 'IF(AND(' + c('start') + '>=cod_date,' + c('end') + '<=maturity_date),1,0)', { cache: (i) => (debtIdx(i) ? 1 : 0) });
    G('Calendar');
    L('fy', 'Financial year (Jul–Jun)', 'FY', F_INT, (c) => 'YEAR(' + c('end') + ')+IF(MONTH(' + c('end') + ')>6,1,0)', { cache: (i) => P[i].fy });
    L('f_fyend', 'Financial year end (30 June)', 'flag', F_INT, (c) => 'IF(MONTH(' + c('end') + ')=6,1,0)', { cache: (i) => (P[i].month === 6 ? 1 : 0) });
    L('opyear', 'Operating year', '', F_INT, (c) => 'IF(' + c('f_op') + '=1,INT((' + c.t + '-n_constr-1)/periods_per_year)+1,0)', { cache: (i) => P[i].opYear });
    L('index', 'CPI index (annual steps from FC)', '', '0.0000', (c) => 'IF(' + c('f_op') + '=1,(1+cpi)^(' + c('fy') + '-fy_first),0)', { cache: 'index' });
    L('degf', 'Degradation factor', '', '0.0000', (c) => 'IF(' + c('f_op') + '=1,(1-degradation)^(' + c('opyear') + '-1),0)',
      { cache: (i) => (i >= nC ? Math.pow(1 - O.degradation, P[i].opYear - 1) : 0) });
    L('hedge', 'Share of the debt hedged', '%', F_PCT1, (c) => 'IF(' + c('end') + '<=EDATE(fc_date,12*hedge_years)-1,hedge_ratio,0)', { cache: 'hedgeShare' });

    sheet('Operations', 'Energy, revenue and costs. Scenario rows use the active scenario; base case rows feed the debt sizing.', C_GOOD);
    G('Active scenario');
    L('energyP50', 'Net energy, P50', 'MWh', F_K, (c) => 'IF(' + c('f_op') + '=1,capacity_mw*8760*cf_p50*availability*' + c('degf') + '/periods_per_year,0)', { cache: 'energyP50', total: true });
    L('energy', 'Net energy, active scenario', 'MWh', F_K, (c) => c('energyP50') + '*sc_energy', { cache: 'energy', total: true });
    L('price', 'PPA price', '/MWh', F_PRICE, (c) => 'IF(' + c('f_op') + '=1,ppa_price*(1+(' + c('index') + '-1)*ppa_indexation),0)',
      { cache: (i) => (S.energy[i] ? S.revenue[i] * 1000 / S.energy[i] : 0) });
    L('revenue', 'Revenue', 'k', F_K, (c) => c('energy') + '*' + c('price') + '/1000', { cache: 'revenue', total: true });
    L('opex', 'Operating costs', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,opex_real*' + c('index') + '*sc_opex/periods_per_year,0)', { cache: 'opex', total: true });
    L('ebitda', 'EBITDA', 'k', F_K, (c) => c('revenue') + '-' + c('opex'), { cache: 'ebitda', total: true, bold: true });
    L('agency', 'Agency fees', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,agency_fee/periods_per_year,0)', { cache: 'agency', total: true });
    L('cfadsPre', 'CFADS pre-tax', 'k', F_K, (c) => c('ebitda') + '-' + c('agency'), { cache: 'cfadsPre', total: true, bold: true });
    G('Base case (P50) and P99, for sizing and sculpting');
    L('opexBase', 'Operating costs, base case', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,opex_real*' + c('index') + '/periods_per_year,0)', { cache: 'opexBase', total: true });
    L('sculptCfads', 'CFADS pre-tax, base case P50', 'k', F_K, (c) => c('energyP50') + '*' + c('price') + '/1000-' + c('opexBase') + '-' + c('agency'), { cache: 'sculptCfads', total: true, bold: true });
    L('cfadsP99', 'CFADS pre-tax, one-year P99', 'k', F_K, (c) => c('energyP50') + '*p99_factor*' + c('price') + '/1000-' + c('opexBase') + '-' + c('agency'),
      { cache: (i) => (O.p99_factor ? S.cfadsP99[i] : S.sculptCfads[i]), total: true });

    sheet('Construction', 'Capex, drawdowns and financing costs until COD.', C_MUTED);
    L('capex', 'Capex', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,capex_total*INDEX(capex_profile,INT((' + c.t + '-1)*months_per_period/6)+1)*months_per_period/6,0)', { cache: 'capex', total: true });
    L('draw', 'Debt drawdown', 'k', F_K, (c) => c('capex') + '*eff_gearing', { cache: 'draw', total: true });
    L('idc', 'Interest during construction', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,' + c('debtOpen') + '*' + c('rate') + ',0)', { cache: 'idc', total: true });
    L('commit', 'Commitment fee', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,(debt_amount-' + c('debtOpen') + ')*margin*commit_fee_pct/periods_per_year,0)', { cache: 'commit', total: true });
    L('upfront', 'Upfront fee', 'k', F_K, (c) => 'IF(' + c.t + '=1,debt_amount*upfront_fee_pct,0)', { cache: 'upfront', total: true });
    L('equityIn', 'Equity contribution', 'k', F_K, (c) => c('capex') + '-' + c('draw') + '+' + c('idc') + '+' + c('commit') + '+' + c('upfront'), { cache: 'equityIn', total: true, bold: true });
    V('project_cost', 'Total project cost', 'k', F_K, (R) => 'capex_total+SUM(' + R('idc') + ')+SUM(' + R('commit') + ')+SUM(' + R('upfront') + ')', { cache: sum.projectCost, bold: true });

    sheet('Debt', 'Interest rates with hedging, debt sizing (three constraints), sculpting and the debt schedule.', C_ACCENT);
    G('Interest rates (per period)');
    L('baseRate', 'Rate, base case (sizing and sculpting)', '%', F_RATE, (c) => '(margin+' + c('hedge') + '*swap_rate+(1-' + c('hedge') + ')*base_rate)/periods_per_year', { cache: 'baseRate' });
    L('rate', 'Rate, active scenario', '%', F_RATE, (c) => '(margin+' + c('hedge') + '*swap_rate+(1-' + c('hedge') + ')*(base_rate+sc_rate))/periods_per_year', { cache: 'rate' });
    G('Debt sizing');
    L('pvBase', 'PV of base case CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('sculptCfads') + '+' + c.n('pvBase') + ')/(1+' + c('baseRate') + '),0)', { cache: (i) => pvBaseS[i] });
    L('pvP99', 'PV of P99 CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('cfadsP99') + '+' + c.n('pvP99') + ')/(1+' + c('baseRate') + '),0)', { cache: (i) => (O.p99_factor ? pvP99S[i] : pvBaseS[i]) });
    V('pv_cfads_debt', 'PV of base case CFADS over the debt tenor, at COD', 'k', F_K, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvBase') + ')', { cache: pvBaseS[nC] });
    V('gearing_cap', '(i) Debt at maximum gearing', 'k', F_K, () => 'capex_total*max_gearing', { cache: Cn.capex_total * D.gearing });
    V('capacity_at_target', '(ii) Debt sculpted at target DSCR on P50', 'k', F_K, () => 'pv_cfads_debt/target_dscr', { cache: sum.capacityAtTarget });
    V('capacity_p99', '(iii) Debt sculpted at 1.00x on one-year P99', 'k', F_K, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvP99') + ')', { cache: O.p99_factor ? sum.capacityP99 : pvBaseS[nC] });
    V('debt_amount', 'Senior debt amount', 'k', F_K, () => 'IF(sizing="lower_of",MIN(gearing_cap,capacity_at_target,capacity_p99),debt_amount_input)', { cache: sum.amount, bold: true });
    V('eff_gearing', 'Effective gearing (of capex)', '', F_PCT, () => 'debt_amount/capex_total', { cache: sum.effGearing });
    V('n_rep', 'Number of repayment periods', '', F_INT, (R) => 'SUM(' + R('f_debt') + ')', { cache: res.meta.nRep });
    V('sculpt_dscr', 'Sculpted DSCR (PV of base CFADS ÷ debt)', '', F_X, () => 'pv_cfads_debt/debt_amount', { cache: pvBaseS[nC] / sum.amount });
    V('annuity_pmt', 'Annuity payment per period (if annuity)', 'k', F_K, () => 'debt_amount*period_rate/(1-(1+period_rate)^-n_rep)',
      { cache: sum.amount * res.meta.rate / PPY / (1 - Math.pow(1 + res.meta.rate / PPY, -res.meta.nRep)) });
    G('Debt schedule');
    L('debtOpen', 'Opening balance', 'k', F_K, (c) => (c.first ? '0' : c.p('debtClose')), { cache: 'debtOpen' });
    L('interest', 'Interest (active scenario rate)', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,' + c('debtOpen') + '*' + c('rate') + ',0)', { cache: 'interest', total: true });
    L('principal', 'Principal (fixed by the base case)', 'k', F_K,
      (c) => 'IF(' + c('f_debt') + '=1,IF(repayment="linear",debt_amount/n_rep,IF(repayment="annuity",annuity_pmt-' + c('debtOpen') + '*period_rate,' + c('sculptCfads') + '/sculpt_dscr-' + c('debtOpen') + '*' + c('baseRate') + ')),0)',
      { cache: 'principal', total: true });
    L('debtService', 'Debt service', 'k', F_K, (c) => c('interest') + '+' + c('principal'), { cache: 'debtService', total: true, bold: true });
    L('debtClose', 'Closing balance', 'k', F_K, (c) => c('debtOpen') + '+' + c('draw') + '-' + c('principal'), { cache: 'debtClose' });
    L('debtCloseFy', 'Closing balance at financial year end', 'k', F_K, (c) => 'IF(' + c('f_fyend') + '=1,' + c('debtClose') + ',0)', { cache: (i) => (P[i].month === 6 ? S.debtClose[i] : 0) });

    sheet('Tax', 'Tax depreciation, losses carried forward and cash tax.', C_MUTED);
    L('depreciation', 'Tax depreciation', 'k', F_K, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.t + '-n_constr<=dep_years*periods_per_year),project_cost/dep_years/periods_per_year,0)', { cache: 'depreciation', total: true });
    L('taxable', 'Taxable income', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c('ebitda') + '-' + c('agency') + '-' + c('mraSpend') + '-' + c('depreciation') + '-' + c('interest') + ',0)', { cache: 'taxable', total: true });
    L('lossUsed', 'Tax losses used', 'k', F_K, (c) => 'IF(' + c('taxable') + '>0,MIN(' + c.p('lossBal') + ',' + c('taxable') + '),0)', { cache: 'lossUsed', total: true });
    L('lossBal', 'Tax losses carried forward', 'k', F_K, (c) => c.p('lossBal') + '+MAX(0,-' + c('taxable') + ')-' + c('lossUsed'), { cache: 'lossBal' });
    L('tax', 'Cash tax', 'k', F_K, (c) => 'IF(' + c('taxable') + '>0,(' + c('taxable') + '-' + c('lossUsed') + ')*tax_rate,0)', { cache: 'tax', total: true });
    L('cfadsPost', 'CFADS post-tax', 'k', F_K, (c) => c('cfadsPre') + '-' + c('tax'), { cache: 'cfadsPost', total: true, bold: true });

    sheet('Reserves', 'Debt service reserve (DSRA) and maintenance reserve (MRA).', C_MUTED);
    G('Debt service reserve account (DSRA)');
    L('dsraTarget', 'Target balance (next DSRA months of debt service)', 'k', F_K,
      (c) => 'IF(' + c.t + '>=n_constr,' + Array.from({ length: 12 / PM }, (_, k) => 'IF(dsra_periods>=' + (k + 1) + ',' + c.ahead('debtService', k + 1) + ',0)').join('+') + ',0)', { cache: 'dsraTarget' });
    L('dsraFund', 'Initial funding by equity at COD', 'k', F_K, (c) => 'IF(' + c('f_cod') + '=1,' + c('dsraTarget') + ',0)', { cache: 'dsraFund', total: true });
    L('dsraDraw', 'Drawing to cover a shortfall', 'k', F_K, (c) => 'IF(' + c('cashAfterDs') + '<0,MIN(' + c.p('dsraBal') + ',-' + c('cashAfterDs') + '),0)', { cache: 'dsraDraw', total: true });
    L('dsraTopUp', 'Top-up from cash flow', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,MAX(0,MIN(' + c('dsraTarget') + '-(' + c.p('dsraBal') + '-' + c('dsraDraw') + '),' + c('cash1') + ')),0)', { cache: 'dsraTopUp', total: true });
    L('dsraRel', 'Release of excess over target', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,MAX(0,(' + c.p('dsraBal') + '-' + c('dsraDraw') + ')-' + c('dsraTarget') + '),0)', { total: true });
    L('dsraAfter', 'Balance after top-up and release', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c.p('dsraBal') + '-' + c('dsraDraw') + '+' + c('dsraTopUp') + '-' + c('dsraRel') + ',0)');
    L('dsraFinal', 'Final release at PPA expiry', 'k', F_K, (c) => 'IF(' + c('f_last') + '=1,' + c('dsraAfter') + ',0)', { total: true });
    L('dsraBal', 'DSRA closing balance', 'k', F_K, (c) => 'IF(' + c('f_cod') + '=1,' + c('dsraFund') + ',IF(' + c('f_op') + '=1,' + c('dsraAfter') + '-' + c('dsraFinal') + ',0))', { cache: 'dsraBal', bold: true });
    G('Maintenance reserve account (MRA): inverter replacement');
    L('mraSpend', 'Inverter replacement (nominal)', 'k', F_K,
      (c) => 'IF(AND(' + c('f_op') + '=1,' + c('opyear') + '>=mra_from,' + c('opyear') + '<=mra_to),mra_cost*' + c('index') + '/((mra_to-mra_from+1)*periods_per_year),0)', { cache: 'mraSpend', total: true });
    V('mra_target', 'Total replacement to fund', 'k', F_K, (R) => 'SUM(' + R('mraSpend') + ')', { cache: sumOf(S.mraSpend) });
    L('mraPlan', 'Planned funding', 'k', F_K,
      (c) => 'IF(AND(' + c('f_op') + '=1,mra_years>0,' + c('opyear') + '>=mra_from-mra_years,' + c('opyear') + '<mra_from),mra_target/(mra_years*periods_per_year),0)', { cache: 'mraPlan', total: true });
    L('mraFund', 'Funding from cash flow', 'k', F_K, (c) => 'MIN(' + c('mraPlan') + ',MAX(0,' + c('cashAfterDsra') + '))', { cache: 'mraFund', total: true });
    L('mraFromMra', 'Replacement paid by the MRA', 'k', F_K, (c) => 'MIN(' + c.p('mraBal') + '+' + c('mraFund') + ',' + c('mraSpend') + ')', { total: true });
    L('mraFromCash', 'Replacement paid from cash flow', 'k', F_K, (c) => c('mraSpend') + '-' + c('mraFromMra'), { cache: 'mraFromCash', total: true });
    L('mraAfter', 'Balance before final release', 'k', F_K, (c) => c.p('mraBal') + '+' + c('mraFund') + '-' + c('mraFromMra'));
    L('mraFinal', 'Final release at PPA expiry', 'k', F_K, (c) => 'IF(' + c('f_last') + '=1,' + c('mraAfter') + ',0)', { total: true });
    L('mraBal', 'MRA closing balance', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c('mraAfter') + '-' + c('mraFinal') + ',0)', { cache: 'mraBal', bold: true });

    sheet('Waterfall', 'Cash flow waterfall in the order of the term sheet, down to distributions.', C_GOOD);
    L('wfCfads', '1. CFADS post-tax (after opex, agency fees, tax)', 'k', F_K, (c) => c('cfadsPost'), { cache: 'cfadsPost', total: true, bold: true });
    L('wfDs', '2. Senior debt service', 'k', F_K, (c) => c('debtService'), { cache: 'debtService', total: true });
    L('cashAfterDs', 'Cash after debt service', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c('cfadsPost') + '-' + c('debtService') + ',0)', { cache: 'cashAfterDs', total: true });
    L('shortfall', 'Debt service shortfall after DSRA', 'k', F_K, (c) => 'MAX(0,-' + c('cashAfterDs') + '-' + c('dsraDraw') + ')', { cache: 'shortfall', total: true });
    L('cash1', 'Cash after DSRA drawing', 'k', F_K, (c) => 'MAX(0,' + c('cashAfterDs') + '+' + c('dsraDraw') + ')');
    L('wfDsra', '3. DSRA top-up (−) / release (+)', 'k', F_K, (c) => c('dsraRel') + '-' + c('dsraTopUp'), { total: true });
    L('cashAfterDsra', 'Cash after DSRA', 'k', F_K, (c) => c('cash1') + '-' + c('dsraTopUp') + '+' + c('dsraRel'), { total: true });
    L('wfMra', '4. MRA funding and replacement paid from cash', 'k', F_K, (c) => '-' + c('mraFund') + '-' + c('mraFromCash'), { total: true });
    L('cashAvail', 'Cash available for distribution', 'k', F_K, (c) => c('cashAfterDsra') + '-' + c('mraFund') + '-' + c('mraFromCash'), { total: true, bold: true });
    L('lockup', 'Lock-up flag (12m DSCR below lock-up)', 'flag', F_INT, (c) => 'IF(' + c('hDscr') + '="",0,IF(' + c('hDscr') + '<lockup_dscr,1,0))', { cache: (i) => (S.lockup[i] ? 1 : 0), total: true });
    L('defaultF', 'Default flag (12m DSCR below default level)', 'flag', F_INT, (c) => 'IF(' + c('hDscr') + '="",0,IF(' + c('hDscr') + '<default_dscr,1,0))', { cache: (i) => (S.defaultFlag[i] ? 1 : 0), total: true });
    L('blocked', 'Distribution blocked (lock-up or DSRA not full)', 'flag', F_INT, (c) => 'IF(' + c('f_op') + '=1,IF(OR(' + c('lockup') + '=1,' + c('dsraAfter') + '<' + c('dsraTarget') + '-0.000001),1,0),0)');
    L('trapped', 'Cash trapped this period', 'k', F_K, (c) => 'IF(AND(' + c('blocked') + '=1,' + c('cashAvail') + '>=0),' + c('cashAvail') + ',0)', { cache: 'trapped', total: true });
    L('trapBal', 'Trapped cash balance', 'k', F_K,
      (c) => 'IF(' + c('f_last') + '=1,0,IF(' + c('cashAvail') + '<0,' + c.p('trapBal') + ',IF(' + c('blocked') + '=1,' + c.p('trapBal') + '+' + c('cashAvail') + ',0)))', { cache: 'trapBal' });
    L('distributions', '5. Distributions to equity', 'k', F_K,
      (c) => 'IF(' + c('f_op') + '=1,IF(' + c('cashAvail') + '<0,' + c('cashAvail') + ',IF(' + c('blocked') + '=1,0,' + c('cashAvail') + '+' + c.p('trapBal') + '))+IF(' + c('f_last') + '=1,IF(AND(' + c('blocked') + '=1,' + c('cashAvail') + '>=0),' + c.p('trapBal') + '+' + c('cashAvail') + ',IF(' + c('cashAvail') + '<0,' + c.p('trapBal') + ',0))+' + c('dsraFinal') + '+' + c('mraFinal') + ',0),0)',
      { cache: 'distributions', total: true, bold: true });

    sheet('Ratios', 'Cover ratios of the active scenario: DSCR, 12-month historic DSCR, LLCR, PLCR.', C_ACCENT);
    L('dscrPre', 'DSCR pre-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('cfadsPre') + '/' + c('debtService') + ',"")', { cache: 'dscrPre' });
    L('dscrPost', 'DSCR post-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('cfadsPost') + '/' + c('debtService') + ',"")', { cache: 'dscrPost', bold: true });
    // 12 months = the last 2 semesters or the last 4 quarters (the number of terms is fixed at export).
    const last12 = (c, key) => Array.from({ length: PPY }, (_, k) => c.back(key, k)).join('+');
    L('hDscr', '12-month historic DSCR (tested)', 'x', F_X,
      (c) => 'IF(AND(' + c('f_debt') + '=1,' + c.back('f_debt', PPY - 1) + '=1),(' + last12(c, 'cfadsPost') + ')/(' + last12(c, 'debtService') + '),"")', { cache: 'hDscr' });
    L('pvPre', 'PV of remaining pre-tax CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('cfadsPre') + '+' + c.n('pvPre') + ')/(1+' + c('rate') + '),0)', { cache: (i) => pvPreS[i] });
    L('pvPost', 'PV of remaining post-tax CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('cfadsPost') + '+' + c.n('pvPost') + ')/(1+' + c('rate') + '),0)', { cache: (i) => pvPostS[i] });
    L('llcrPre', 'LLCR pre-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('pvPre') + '/' + c('debtOpen') + ',"")', { cache: 'llcrPre' });
    L('llcrPost', 'LLCR post-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('pvPost') + '/' + c('debtOpen') + ',"")', { cache: 'llcrPost' });
    L('pvPpa', 'PV of remaining pre-tax CFADS to PPA expiry', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,(' + c('cfadsPre') + '+' + c.n('pvPpa') + ')/(1+' + c('rate') + '),0)', { cache: (i) => pvPpaS[i] });
    V('min_dscr_post', 'Minimum DSCR post-tax', 'x', F_X, (R) => 'MIN(' + R('dscrPost') + ')', { cache: sum.minDscrPost, bold: true });
    V('avg_dscr_post', 'Average DSCR post-tax', 'x', F_X, (R) => 'AVERAGE(' + R('dscrPost') + ')', { cache: sum.avgDscrPost });
    V('llcr_cod_pre', 'LLCR at COD, pre-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPre') + ')/debt_amount', { cache: sum.llcrCodPre });
    V('llcr_cod_post', 'LLCR at COD, post-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPost') + ')/debt_amount', { cache: sum.llcrCodPost });
    V('plcr_cod', 'PLCR at COD, pre-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPpa') + ')/debt_amount', { cache: sum.plcrCod });

    sheet('Returns', 'Equity cash flows and returns of the active scenario.', C_GOOD);
    L('retIn', 'Equity contributions during construction', 'k', F_K, (c) => '-' + c('equityIn'), { cache: (i) => -S.equityIn[i], total: true });
    L('retDsra', 'DSRA funding at COD', 'k', F_K, (c) => '-' + c('dsraFund'), { cache: (i) => -S.dsraFund[i], total: true });
    L('retDist', 'Distributions', 'k', F_K, (c) => c('distributions'), { cache: 'distributions', total: true });
    L('equityFlows', 'Equity cash flow', 'k', F_K, (c) => c('retIn') + '+' + c('retDsra') + '+' + c('retDist'), { cache: 'equityFlows', total: true, bold: true });
    let cum = 0;
    const cumS = S.equityFlows.map((v) => (cum += v));
    L('equityCum', 'Cumulative equity cash flow', 'k', F_K, (c) => c.p('equityCum') + '+' + c('equityFlows'), { cache: (i) => cumS[i] });
    V('equity_irr', 'Equity IRR (post-tax, nominal, annualised)', '', F_PCT, (R) => '(1+IRR(' + R('equityFlows') + '))^periods_per_year-1', { cache: sum.equityIrr, bold: true });
    const invested = sumOf(S.equityIn) + sumOf(S.dsraFund);
    V('cash_multiple', 'Cash multiple (distributions ÷ equity invested)', 'x', '0.00"x"', (R) => 'SUM(' + R('retDist') + ')/-(SUM(' + R('retIn') + ')+SUM(' + R('retDsra') + '))', { cache: sum.totalDistributions / invested });

    // Assign row numbers on each sheet: rows 1-6 hold the title and the timeline header.
    const ROW = {};
    for (const sh of SHEETS) {
      let rr = 8;
      for (const row of sh.rows) {
        if (row.group) { row.r = ++rr; rr++; continue; }
        row.r = rr++;
        ROW[row.key] = { sheet: sh.name, r: row.r };
      }
    }
    const rangeOf = (key) => ROW[key].sheet + '!$' + col(1) + '$' + ROW[key].r + ':$' + LAST + '$' + ROW[key].r;
    const R = rangeOf;

    const hdr = wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'right', numFmt: F_INT });
    const hdrCon = wb.style({ bold: true, fill: C_CONSTR_BG, border: 'bottom', align: 'right', numFmt: F_INT });
    const phase = wb.style({ color: C_MUTED, align: 'right', italic: true });
    const dStart = wb.style({ numFmt: F_MONTH, color: C_MUTED, align: 'right' });
    const dEnd = wb.style({ numFmt: F_MONTH, bold: true, align: 'right' });
    const boldTop = wb.style({ bold: true, border: 'top' });

    for (const sh of SHEETS) {
      const ws = wb.addSheet(sh.name);
      ws.tab = sh.tab;
      ws.width(1, 50).width(2, 7).width(3, 14);
      for (let t = 1; t <= N; t++) ws.width(FIRST + t - 1, 11);
      ws.freezeAt('D7');
      banner(ws, sh.name, sh.note + ' ' + (PM === 3 ? 'Quarterly' : 'Semi-annual') + ' periods across, ' + deal.meta.currency + ' thousands unless stated. Shaded columns = construction.', FIRST + N - 1);
      ws.set('A3', 'Phase', st.muted);
      ws.set('A4', 'Period', st.head); ws.set('B4', '', st.head); ws.set('C4', 'Total / value', st.headR);
      ws.set('A5', 'Period start', st.muted); ws.set('A6', 'Period end', st.bold);
      for (let t = 1; t <= N; t++) {
        const con = t <= nC, cl = col(t);
        ws.set(cl + '3', con ? 'Constr.' : t <= nM ? 'Debt' : 'Tail', phase);
        ws.set(cl + '4', t, con ? hdrCon : hdr);
        const pre = sh.name === 'Timing' ? '' : 'Timing!';
        ws.set(cl + '5', { f: pre + cl + ROW.start.r, v: X.excelDate(periodStart(P[t - 1], PM)) }, dStart);
        ws.set(cl + '6', { f: pre + cl + ROW.end.r, v: X.excelDate(periodEnd(P[t - 1])) }, dEnd);
      }
      for (const row of sh.rows) {
        if (row.group) {
          ws.set('A' + row.r, row.group, st.group);
          for (let c = 2; c <= FIRST + N - 1; c++) ws.set(X.colName(c) + row.r, '', st.group);
          continue;
        }
        ws.set('A' + row.r, row.label, row.bold ? boldTop : 0);
        ws.set('B' + row.r, row.unit, st.muted);
        if (row.scalar) {
          ws.set('C' + row.r, { f: row.f(rangeOf), v: row.cache }, calc(row.fmt, true));
          wb.defineName(row.key, sh.name + '!$C$' + row.r);
          continue;
        }
        const style = row.key === 'start' ? dStart : row.key === 'end' ? dEnd : (row.bold ? wb.style({ numFmt: row.fmt, bold: true, border: 'top' }) : calc(row.fmt));
        for (let t = 1; t <= N; t++) {
          const cl = col(t), i = t - 1;
          // A reference to another row in the same period column, prefixed by its sheet if needed.
          const at = (key, c) => (ROW[key].sheet === sh.name ? '' : ROW[key].sheet + '!') + c + ROW[key].r;
          const ref = (key) => at(key, cl);
          ref.t = cl + '$4';
          ref.first = t === 1;
          ref.p = (key) => (t === 1 ? '0' : at(key, col(t - 1)));
          ref.n = (key) => at(key, col(t + 1)); // the column after the last period is empty, i.e. 0
          ref.back = (key, k) => (k === 0 ? ref(key) : t - k < 1 ? '0' : at(key, col(t - k)));
          ref.ahead = (key, k) => at(key, col(t + k)); // columns after the last period are empty, i.e. 0
          let v;
          if (typeof row.cache === 'string') v = S[row.cache][i];
          else if (typeof row.cache === 'function') v = row.cache(i);
          if (v === null) v = '';
          if (typeof v === 'boolean') v = v ? 1 : 0;
          ws.set(cl + row.r, { f: row.f(ref), v }, style);
        }
        if (row.total) {
          const tv = typeof row.cache === 'string' ? sumOf(S[row.cache]) : typeof row.cache === 'function' ? sumOf(Array.from({ length: N }, (_, i) => row.cache(i))) : undefined;
          ws.set('C' + row.r, { f: 'SUM(' + col(1) + row.r + ':' + LAST + row.r + ')', v: tv }, calc(row.fmt, true));
        }
      }
    }

    // ---------------- Annual summary ----------------
    const years = [];
    for (let y = P[0].fy; y <= P[N - 1].fy; y++) years.push(y);
    const wsA = wb.addSheet('Annual');
    wsA.tab = C_GOOD;
    wsA.width(1, 40).width(2, 7);
    years.forEach((_, k) => wsA.width(3 + k, 11));
    wsA.freezeAt('C5');
    banner(wsA, 'Annual summary', 'Financial years ending 30 June. Each figure adds up the periods of its year (SUMIF on the Timing sheet). Active scenario.', 2 + years.length);
    wsA.set('A4', 'Financial year', st.head); wsA.set('B4', '', st.head);
    years.forEach((y, k) => wsA.set(X.colName(3 + k) + '4', y, hdr));
    const fyOf = P.map((p) => p.fy);
    const byYear = (arr) => years.map((y) => sumOf(arr.filter((_, i) => fyOf[i] === y)));
    const seriesFor = { debtCloseFy: S.debtClose.map((v, i) => (P[i].month === 6 ? v : 0)) };
    const annual = [
      ['g', 'Operations'],
      ['Net energy', 'MWh', 'energy'], ['Revenue', 'k', 'revenue'], ['Operating costs', 'k', 'opex'], ['EBITDA', 'k', 'ebitda', 1],
      ['Agency fees', 'k', 'agency'], ['Cash tax', 'k', 'tax'], ['CFADS post-tax', 'k', 'cfadsPost', 1],
      ['g', 'Senior debt'],
      ['Interest', 'k', 'interest'], ['Principal', 'k', 'principal'], ['Debt service', 'k', 'debtService', 1], ['Closing balance (30 June)', 'k', 'debtCloseFy'],
      ['DSCR post-tax (annual)', 'x', 'dscrA', 1],
      ['g', 'Reserves and equity'],
      ['MRA funding', 'k', 'mraFund'], ['Inverter replacement', 'k', 'mraSpend'], ['Distributions to equity', 'k', 'distributions', 1], ['Equity cash flow', 'k', 'equityFlows', 1],
    ];
    const rowA = {};
    let ra = 6;
    for (const [label, unit, key, bold] of annual) {
      if (label === 'g') { wsA.set('A' + ra, unit, st.group); years.forEach((_, k) => wsA.set(X.colName(3 + k) + ra, '', st.group)); ra++; continue; }
      rowA[key] = ra;
      wsA.set('A' + ra, label, bold ? boldTop : 0); wsA.set('B' + ra, unit, st.muted);
      const vals = key === 'dscrA' ? null : byYear(seriesFor[key] || S[key]);
      years.forEach((y, k) => {
        const cl = X.colName(3 + k);
        if (key === 'dscrA') {
          const cf = byYear(S.cfadsPost)[k], ds = byYear(S.debtService)[k];
          wsA.set(cl + ra, { f: 'IF(' + cl + rowA.debtService + '>0,' + cl + rowA.cfadsPost + '/' + cl + rowA.debtService + ',"")', v: ds > 0 ? cf / ds : '' }, wb.style({ numFmt: F_X, bold: true, border: 'top' }));
        } else {
          wsA.set(cl + ra, { f: 'SUMIF(' + R('fy') + ',' + cl + '$4,' + R(key) + ')', v: vals[k] }, bold ? wb.style({ numFmt: F_K, bold: true, border: 'top' }) : calc(F_K));
        }
      });
      ra++;
    }

    // ---------------- Dashboard ----------------
    const wsO = wb.addSheet('Dashboard');
    wsO.tab = C_NAVY;
    wsO.width(1, 44).width(2, 16).width(3, 4).width(4, 40).width(5, 14).width(6, 14).width(7, 14).width(8, 14);
    banner(wsO, 'Dashboard', 'Key results, all linked to the calculation sheets. Active scenario shown in B4.', 8);
    const out = (row, label, f, v, fmt, colL, colV) => { wsO.set((colL || 'A') + row, label); wsO.set((colV || 'B') + row, { f, v }, calc(fmt, false)); };
    wsO.set('A4', 'Active scenario', st.bold); wsO.set('B4', { f: 'scenario', v: SCENARIO_NAMES[scIndex] }, wb.style({ bold: true, color: C_ACCENT }));
    wsO.set('A6', 'Uses', st.head); wsO.set('B6', deal.meta.currency + ' k', st.headR);
    wsO.set('D6', 'Sources', st.head); wsO.set('E6', deal.meta.currency + ' k', st.headR);
    out(7, 'Construction capex', 'capex_total', Cn.capex_total, F_K);
    out(8, 'Interest during construction', 'SUM(' + R('idc') + ')', sum.idc, F_K);
    out(9, 'Upfront fee', 'SUM(' + R('upfront') + ')', sum.upfrontFee, F_K);
    out(10, 'Commitment fee', 'SUM(' + R('commit') + ')', sum.commitmentFees, F_K);
    wsO.set('A11', 'Total uses', st.bold); wsO.set('B11', { f: 'SUM(B7:B10)', v: sum.projectCost }, calc(F_K, true));
    out(7, 'Senior debt', 'debt_amount', sum.amount, F_K, 'D', 'E');
    out(8, 'Sponsor equity at financial close', 'SUM(' + R('equityIn') + ')', sum.equityAtFc, F_K, 'D', 'E');
    wsO.set('D11', 'Total sources', st.bold); wsO.set('E11', { f: 'E7+E8', v: sum.amount + sum.equityAtFc }, calc(F_K, true));
    wsO.set('A12', 'DSRA funded by a further equity contribution at COD:', st.muted);
    wsO.set('B12', { f: 'SUM(' + R('dsraFund') + ')', v: sum.dsraAtCod }, calc(F_K));

    wsO.set('A14', 'Debt sizing: lowest of', st.head); wsO.set('B14', deal.meta.currency + ' k', st.headR);
    out(15, '(i) Maximum gearing × capex', 'gearing_cap', Cn.capex_total * D.gearing, F_K);
    out(16, '(ii) Sculpted at target DSCR on P50', 'capacity_at_target', sum.capacityAtTarget, F_K);
    out(17, '(iii) Sculpted at 1.00x on one-year P99', 'capacity_p99', O.p99_factor ? sum.capacityP99 : pvBaseS[nC], F_K);
    wsO.set('A18', 'Senior debt (' + (D.sizing === 'lower_of' ? 'lowest of the three' : 'fixed facility') + ')', st.bold); wsO.set('B18', { f: 'debt_amount', v: sum.amount }, calc(F_K, true));
    out(19, 'Headroom vs DSCR capacity', 'capacity_at_target/debt_amount-1', sum.headroom, F_PCT1);

    wsO.set('A21', 'Cover ratios (active scenario)', st.head); wsO.set('B21', '', st.headR);
    const ratio = [
      ['Effective gearing (of capex)', 'eff_gearing', sum.effGearing, F_PCT1],
      ['Sculpted DSCR (pre-tax, base case)', 'IF(repayment="sculpted",sculpt_dscr,"n/a")', D.repayment === 'sculpted' ? pvBaseS[nC] / sum.amount : 'n/a', F_X],
      ['Minimum DSCR post-tax', 'min_dscr_post', sum.minDscrPost, F_X],
      ['Average DSCR post-tax', 'avg_dscr_post', sum.avgDscrPost, F_X],
      ['LLCR at COD, post-tax', 'llcr_cod_post', sum.llcrCodPost, F_X],
      ['PLCR at COD, pre-tax', 'plcr_cod', sum.plcrCod, F_X],
      ['Lock-up tests failed', 'SUM(' + R('lockup') + ')', sum.lockupPeriods, F_INT],
      ['Default tests failed', 'SUM(' + R('defaultF') + ')', sum.defaultPeriods, F_INT],
    ];
    ratio.forEach(([l, f, v, fmt], i) => out(22 + i, l, f, v, fmt));
    let ro = 22 + ratio.length + 1;
    wsO.set('A' + ro, 'First operating year', st.head); wsO.set('B' + ro, '', st.headR); ro++;
    const fy = sum.firstOpYear;
    const firstYear = (key) => 'SUMIF(' + R('opyear') + ',1,' + R(key) + ')';
    out(ro++, 'Net energy (MWh)', firstYear('energy'), fy.energy, F_K);
    out(ro++, 'Revenue', firstYear('revenue'), fy.revenue, F_K);
    out(ro++, 'EBITDA', firstYear('ebitda'), fy.ebitda, F_K);
    ro++;
    wsO.set('A' + ro, 'Equity', st.head); wsO.set('B' + ro, '', st.headR); ro++;
    out(ro++, 'Equity IRR (post-tax, nominal)', 'equity_irr', sum.equityIrr, F_PCT);
    out(ro++, 'Cash multiple', 'cash_multiple', sum.totalDistributions / invested, '0.00"x"');
    out(ro++, 'Total distributions', 'SUM(' + R('distributions') + ')', sum.totalDistributions, F_K);

    // Lender cases side by side: values from the web app at export (the live model shows one case at a time).
    wsO.set('D21', 'Lender cases (values at export)', st.head);
    ['Min DSCR', 'Avg DSCR', 'LLCR', 'Equity IRR'].forEach((h, k) => wsO.set(X.colName(5 + k) + '21', h, st.headR));
    SCENARIO_KEYS.forEach((key, k) => {
      const x = E.run(Object.assign({}, deal, { scenario: key }));
      const row = 22 + k;
      wsO.set('D' + row, SCENARIO_NAMES[k], k === scIndex ? st.bold : 0);
      if (!x.ok) return;
      const s = x.summary;
      wsO.set('E' + row, s.minDscrPost, calc(F_X)); wsO.set('F' + row, s.avgDscrPost, calc(F_X));
      wsO.set('G' + row, s.llcrCodPost, calc(F_X)); wsO.set('H' + row, s.equityIrr == null ? '' : s.equityIrr, calc(F_PCT));
    });
    wsO.set('D26', 'Pick a case on the Inputs sheet to see it live everywhere.', st.muted);

    // ---------------- Checks ----------------
    const wsK = wb.addSheet('Checks');
    wsK.tab = C_GOOD;
    wsK.width(1, 56).width(2, 16).width(3, 16).width(4, 14).width(5, 10);
    banner(wsK, 'Checks', 'TRUE means the check passes. Part 2 compares this workbook with the web app at the time of export.', 5);
    wsK.set('A4', 'Model integrity', st.head); wsK.set('B4', 'Result', st.headC);
    const lastCol = (key) => ROW[key].sheet + '!$' + LAST + '$' + ROW[key].r;
    const integrity = [
      ['Debt fully drawn at COD', 'ABS(SUM(' + R('draw') + ')-debt_amount)<0.01'],
      ['Debt repaid at maturity', 'ABS(' + lastCol('debtClose') + ')<0.01'],
      ['Sources equal uses', 'ABS(debt_amount+SUM(' + R('equityIn') + ')-project_cost)<0.01'],
      ['Gearing within limit', 'eff_gearing<=max_gearing+0.000000001'],
      ['No debt service shortfall', 'SUM(' + R('shortfall') + ')<0.01'],
      ['No principal re-borrowing', 'MIN(' + R('principal') + ')>=-0.000000001'],
      ['Capex profile adds up to 100%', 'ABS(SUM(capex_profile)-1)<0.000001'],
      ['One capex share per construction semester', 'COUNT(capex_profile)*6=n_constr*months_per_period'],
      ['Timeline ends at PPA expiry (columns fixed at export)', lastCol('end') + '=ppa_end_date'],
      ['DSRA of 0, ' + (PM === 3 ? '3, 6, 9' : '6') + ' or 12 months (as built)', 'AND(dsra_months>=0,dsra_months<=12,INT(dsra_months/months_per_period)*months_per_period=dsra_months)'],
      ['Hedge ratio between 0% and 100%', 'AND(hedge_ratio>=0,hedge_ratio<=1)'],
      ['P99 ≤ P90 ≤ P50 energy', 'AND(p99_factor<=p90_factor,p90_factor<=1)'],
      ['MRA balance never negative', 'MIN(' + R('mraBal') + ')>=-0.000001'],
      ['Waterfall adds up (CFADS − debt service ± reserves = cash available)', 'ABS(SUM(' + R('cfadsPost') + ')-SUM(' + R('debtService') + ')+SUM(' + R('wfDsra') + ')+SUM(' + R('wfMra') + ')+SUM(' + R('shortfall') + ')+SUM(' + R('dsraDraw') + ')-SUM(' + R('cashAvail') + '))<0.01'],
    ];
    integrity.forEach(([l, f], i) => { wsK.set('A' + (5 + i), l); wsK.set('B' + (5 + i), { f, v: true }, st.ok); });
    let rk = 5 + integrity.length;
    wsK.set('A' + rk, 'All integrity checks pass', st.bold);
    wsK.set('B' + rk, { f: 'AND(B5:B' + (rk - 1) + ')', v: true }, st.ok);
    const integrityAll = 'Checks!$B$' + rk;
    rk += 2;
    wsK.set('A' + rk, 'Reconciliation with the web app', st.head);
    wsK.set('B' + rk, 'Web app', st.headR); wsK.set('C' + rk, 'This workbook', st.headR); wsK.set('D' + rk, 'Difference', st.headR);
    wsK.set('E' + rk, 'Match', st.headC);
    rk++;
    wsK.set('A' + rk, 'App values are frozen at export. If you change an input or the scenario, differences here are expected.', st.note); rk++;
    const recon = [
      ['Senior debt', sum.amount, 'debt_amount', F_K],
      ['Total project cost', sum.projectCost, 'project_cost', F_K],
      ['Interest during construction', sum.idc, 'SUM(' + R('idc') + ')', F_K],
      ['Equity at financial close', sum.equityAtFc, 'SUM(' + R('equityIn') + ')', F_K],
      ['First-year EBITDA', fy.ebitda, firstYear('ebitda'), F_K],
      ['Debt capacity at target DSCR', sum.capacityAtTarget, 'capacity_at_target', F_K],
      ['Debt capacity at 1.00x on P99', O.p99_factor ? sum.capacityP99 : pvBaseS[nC], 'capacity_p99', F_K],
      ['Total agency fees', sum.agencyFees, 'SUM(' + R('agency') + ')', F_K],
      ['Total inverter replacement', sum.maintenance, 'SUM(' + R('mraSpend') + ')', F_K],
      ['Minimum DSCR post-tax', sum.minDscrPost, 'min_dscr_post', F_X],
      ['Average DSCR post-tax', sum.avgDscrPost, 'avg_dscr_post', F_X],
      ['LLCR at COD, pre-tax', sum.llcrCodPre, 'llcr_cod_pre', F_X],
      ['LLCR at COD, post-tax', sum.llcrCodPost, 'llcr_cod_post', F_X],
      ['PLCR at COD', sum.plcrCod, 'plcr_cod', F_X],
      ['DSRA at COD', sum.dsraAtCod, 'SUM(' + R('dsraFund') + ')', F_K],
      ['Total cash tax', sumOf(S.tax), 'SUM(' + R('tax') + ')', F_K],
      ['Total distributions', sum.totalDistributions, 'SUM(' + R('distributions') + ')', F_K],
      ['Equity IRR', sum.equityIrr, 'equity_irr', F_PCT],
    ];
    const reconStart = rk;
    recon.forEach(([l, app, f, fmt]) => {
      wsK.set('A' + rk, l);
      wsK.set('B' + rk, app, calc(fmt));
      wsK.set('C' + rk, { f, v: app }, calc(fmt));
      wsK.set('D' + rk, { f: 'C' + rk + '-B' + rk, v: 0 }, calc(fmt === F_K ? '#,##0.000;(#,##0.000);"–"' : '0.000000'));
      wsK.set('E' + rk, { f: 'ABS(D' + rk + ')<=MAX(0.001,ABS(B' + rk + ')*0.000001)', v: true }, st.ok);
      rk++;
    });
    wsK.set('A' + rk, 'Workbook matches the web app', st.bold);
    wsK.set('E' + rk, { f: 'AND(E' + reconStart + ':E' + (rk - 1) + ')', v: true }, st.ok);
    const reconAll = 'Checks!$E$' + rk;

    // ---------------- Cover (first tab) ----------------
    const wsV = wb.addSheet('Cover');
    wsV.tab = C_NAVY;
    wsV.width(1, 34).width(2, 90);
    banner(wsV, deal.meta.name, 'Project finance model generated from the term sheet by the Model Builder web app · Jean Grollet', 2);
    const cover = [
      ['Deal', deal.meta.name + (deal.meta.seed ? ' (random deal #' + deal.meta.seed + ', fictional)' : '')],
      ['Asset', O.capacity_mw + ' MWac solar' + (deal.meta.state ? ', ' + deal.meta.state + ', Australia' : '')],
      ['Currency and units', deal.meta.currency + ' thousands, ' + PERIOD_WORD + ' periods'],
      ['Timeline', P[0].label + ' to ' + P[N - 1].label + ' (' + N + ' periods; the number of columns is fixed at export)'],
    ];
    cover.forEach(([k, v], i) => { wsV.set('A' + (4 + i), k, st.bold); wsV.set('B' + (4 + i), v); });
    wsV.set('A9', 'Key outputs (live)', st.h);
    const live = [
      ['Active scenario', 'scenario', SCENARIO_NAMES[scIndex], ''],
      ['Senior debt', 'debt_amount', sum.amount, F_K],
      ['Minimum DSCR post-tax', 'min_dscr_post', sum.minDscrPost, F_X],
      ['LLCR at COD, post-tax', 'llcr_cod_post', sum.llcrCodPost, F_X],
      ['Equity IRR', 'equity_irr', sum.equityIrr, F_PCT],
      ['Integrity checks pass', integrityAll, true, ''],
      ['Matches the web app', reconAll, true, ''],
    ];
    live.forEach(([k, f, v, fmt], i) => { wsV.set('A' + (10 + i), k); wsV.set('B' + (10 + i), { f, v }, wb.style({ numFmt: fmt || undefined, bold: true, size: 12, color: fmt || i === 0 ? C_ACCENT : C_GOOD, align: 'left' })); });
    wsV.set('A18', 'How the workbook is built', st.h);
    const guide = [
      ['Dashboard', 'Sources and uses, the three sizing constraints, cover ratios, returns, and the four lender cases side by side.'],
      ['Inputs', 'Every assumption in blue, with the term sheet clause it comes from, and the active scenario. Named cells (e.g. margin, target_dscr) are used in all formulas.'],
      ['Scenarios', 'Base case (P50), P90, P99 and Downside. The debt is sized and sculpted on the base case; the other cases run against that fixed repayment schedule, as lenders do.'],
      ['Timing', 'Period dates, construction / operations / debt flags, financial years, CPI and degradation factors, hedged share.'],
      ['Operations', 'Energy, revenue, operating costs, agency fees and CFADS, for the active scenario and for the base case and P99 used in sizing.'],
      ['Construction', 'Capex, pro rata debt drawdowns, interest during construction, fees and equity.'],
      ['Debt', 'Hedged interest rates, sizing (gearing, DSCR on P50, 1.00x on P99), sculpting and the debt schedule.'],
      ['Tax', 'Straight-line tax depreciation, losses carried forward, cash tax.'],
      ['Reserves', 'DSRA (funded at COD, topped up and released) and MRA (funded ahead of the inverter replacement).'],
      ['Waterfall', 'Cash in the order of the term sheet: CFADS, debt service, DSRA, MRA, lock-up test, distributions.'],
      ['Ratios / Returns', 'DSCR, 12-month historic DSCR, LLCR, PLCR; equity cash flows, IRR and cash multiple.'],
      ['Annual', 'The same model summed by financial year (July–June).'],
      ['Checks', 'Integrity checks and a reconciliation against the web app.'],
      ['Colour code', 'Blue on light blue = input. Black = formula. Nothing is hard-coded on the calculation sheets.'],
      ['Simplifications', 'Interest = balance × annual rate ÷ periods per year (no day count). Swaps replace BBSW on the hedged share (no mark-to-market). No cash sweep or terminal value.'],
    ];
    guide.forEach(([k, v], i) => { wsV.set('A' + (19 + i), k, st.bold); wsV.set('B' + (19 + i), v, st.noteWrap); });

    // Tab order: Cover, Dashboard, Inputs, Scenarios, the period sheets, Annual, Checks.
    const order = ['Cover', 'Dashboard', 'Inputs', 'Scenarios', ...SHEETS.map((s) => s.name), 'Annual', 'Checks'];
    wb.sheets.sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
    return wb;
  }

  function periodEnd(p) {
    const lastDay = new Date(Date.UTC(p.year, p.month, 0)).getUTCDate();
    return p.year + '-' + String(p.month).padStart(2, '0') + '-' + lastDay;
  }
  function periodStart(p, pm) {
    let y = p.year, m = p.month - pm + 1;
    if (m < 1) { m += 12; y -= 1; }
    return y + '-' + String(m).padStart(2, '0') + '-01';
  }

  // Returns the .xlsx file as bytes.
  function exportXlsx(deal, res) { return buildWorkbook(deal, res).toBytes(); }

  const api = { buildWorkbook, exportXlsx, SCENARIO_NAMES };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFExcel = api;
})(typeof window !== 'undefined' ? window : globalThis);
