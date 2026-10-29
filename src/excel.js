/*
 * Excel export — v1
 * Rebuilds the whole model as live Excel formulas: change an input in the Inputs sheet and
 * every line recalculates. Sheets: Cover, Inputs, Calc, Outputs, Checks.
 * Calc follows the same order as engine.js (see SPEC.md section 5). Each formula cell also
 * carries the engine's value as a cached result, and the Checks sheet compares the two.
 */
(function (root) {
  'use strict';

  const X = typeof module !== 'undefined' && module.exports ? require('./xlsx.js') : root.PFXlsx;

  const C_ACCENT = '2F5BEA', C_INPUT_BG = 'EEF2FE', C_MUTED = '6B7280', C_HEAD_BG = 'F6F7F9', C_GOOD = '15803D';
  const C_NAVY = '1B2333', C_GOOD_BG = 'DCFCE7', C_CONSTR_BG = 'FFF4E0';
  const F_K = '#,##0;(#,##0);"–"', F_K1 = '#,##0.0;(#,##0.0);"–"', F_X = '0.000"x"', F_PCT = '0.00%', F_PCT1 = '0.0%';
  const F_DATE = 'dd-mmm-yyyy', F_MONTH = 'mmm-yy', F_INT = '0', F_PRICE = '#,##0.00';

  function buildWorkbook(deal, res) {
    const wb = new X.Workbook();
    const st = {
      title: wb.style({ bold: true, size: 16 }),
      h: wb.style({ bold: true, size: 11, color: C_ACCENT }),
      note: wb.style({ italic: true, color: C_MUTED }),
      noteWrap: wb.style({ color: C_MUTED, wrap: true }),
      bold: wb.style({ bold: true }),
      head: wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom' }),
      headR: wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'right' }),
      muted: wb.style({ color: C_MUTED }),
      good: wb.style({ bold: true, color: C_GOOD }),
      banner: wb.style({ bold: true, size: 14, color: 'FFFFFF', fill: C_NAVY, valign: 'center' }),
      bannerFill: wb.style({ fill: C_NAVY }),
      group: wb.style({ bold: true, color: C_ACCENT, fill: C_INPUT_BG }),
      ok: wb.style({ bold: true, align: 'center', color: C_GOOD, fill: C_GOOD_BG }),
    };
    // Dark title band across the top of a sheet, with a one-line note under it.
    const banner = (ws, title, note, lastCol) => {
      ws.height(1, 30);
      ws.set('A1', '  ' + title, st.banner);
      for (let c = 2; c <= lastCol; c++) ws.set(X.colName(c) + '1', '', st.bannerFill);
      ws.set('A2', note, st.note);
    };
    const inp = (numFmt) => wb.style({ color: C_ACCENT, fill: C_INPUT_BG, numFmt });
    const calc = (numFmt, bold) => wb.style(bold ? { numFmt, bold: true } : { numFmt });

    const S = res.series, P = res.periods, N = res.meta.N, sum = res.summary;
    const FIRST = 4; // column D holds period 1
    const col = (t) => X.colName(FIRST + t - 1);
    const LAST = col(N);
    const PM = res.meta.pm, PPY = res.meta.ppy;
    const PERIOD_WORD = PM === 3 ? 'quarterly' : 'semi-annual';

    // ---------------- Inputs ----------------
    const wsI = wb.addSheet('Inputs');
    wsI.width(1, 40).width(2, 16).width(3, 12).width(4, 80);
    wsI.tab = C_ACCENT;
    banner(wsI, 'Inputs', 'Blue cells are inputs: change them and the whole model recalculates. Column D gives the term sheet clause each input comes from.', 4);
    let r = 4;
    const src = (path) => (deal.sources && deal.sources[path]) || '';
    const section = (ws, title) => { r++; ws.set('A' + r, title, st.h); r++; };
    const input = (name, label, value, fmt, unit, path, list) => {
      wsI.set('A' + r, label);
      wsI.set('B' + r, value, inp(fmt));
      wsI.set('C' + r, unit || '', st.muted);
      wsI.set('D' + r, src(path), st.muted);
      if (list) wsI.dropdown('B' + r, list);
      wb.defineName(name, 'Inputs!$B$' + r);
      r++;
    };
    const T = deal.timeline, Cn = deal.construction, O = deal.operations, D = deal.senior_debt;
    wsI.set('A' + r, 'Input', st.head); wsI.set('B' + r, 'Value', st.headR); wsI.set('C' + r, 'Unit', st.head); wsI.set('D' + r, 'Term sheet source', st.head); r++;
    section(wsI, 'Timeline');
    input('months_per_period', 'Months per period (' + PERIOD_WORD + ')', PM, F_INT, 'fixed at export', 'timeline.periodicity');
    input('fc_date', 'Financial close', X.excelDate(T.financial_close), F_DATE, '', 'timeline.financial_close');
    input('cod_date', 'Commercial operation date (COD)', X.excelDate(T.cod), F_DATE, '', 'timeline.cod');
    input('maturity_date', 'Final maturity', X.excelDate(T.debt_maturity), F_DATE, '', 'timeline.debt_maturity');
    input('ppa_end_date', 'PPA expiry', X.excelDate(T.ppa_end), F_DATE, '', 'timeline.ppa_end');
    section(wsI, 'Construction');
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
    section(wsI, 'Operations');
    input('capacity_mw', 'Capacity', O.capacity_mw, '#,##0.0', 'MWac', 'operations.capacity_mw');
    input('cf_p50', 'P50 capacity factor', O.capacity_factor_p50, F_PCT, '', 'operations.capacity_factor_p50');
    input('availability', 'Availability', O.availability, F_PCT, '', 'operations.availability');
    input('degradation', 'Degradation', O.degradation, F_PCT, '% p.a.', 'operations.degradation');
    input('ppa_price', 'PPA price at financial close', O.ppa_price_fc, F_PRICE, deal.meta.currency + '/MWh', 'operations.ppa_price_fc');
    input('ppa_indexation', 'PPA indexation', O.ppa_indexation, '0%', '% of CPI', 'operations.ppa_indexation');
    input('opex_real', 'Operating costs (real, at financial close)', O.opex_real_fc, F_K, 'k p.a.', 'operations.opex_real_fc');
    input('cpi', 'CPI', O.cpi, F_PCT, '% p.a.', 'operations.cpi');
    section(wsI, 'Senior debt');
    input('sizing', 'Debt sizing', D.sizing === 'lower_of' ? 'lower_of' : 'fixed', '', 'fixed / lower_of', 'senior_debt.sizing', ['fixed', 'lower_of']);
    input('debt_amount_input', 'Facility amount (used when sizing = fixed)', D.amount, F_K, deal.meta.currency + ' k', 'senior_debt.amount');
    input('max_gearing', 'Maximum gearing (of capex)', D.gearing, F_PCT1, '', 'senior_debt.gearing');
    input('base_rate', 'Base rate (BBSW)', D.base_rate, F_PCT, '', 'senior_debt.base_rate');
    input('margin', 'Margin', D.margin, F_PCT, '', 'senior_debt.margin');
    input('upfront_fee_pct', 'Upfront fee', D.upfront_fee, F_PCT, '% of facility', 'senior_debt.upfront_fee');
    input('commit_fee_pct', 'Commitment fee', D.commitment_fee_pct_margin, F_PCT1, '% of margin', 'senior_debt.commitment_fee_pct_margin');
    input('repayment', 'Repayment profile', D.repayment, '', 'sculpted / annuity / linear', 'senior_debt.repayment', ['sculpted', 'annuity', 'linear']);
    input('target_dscr', 'Target DSCR', D.target_dscr, '0.00"x"', '', 'senior_debt.target_dscr');
    section(wsI, 'Tax');
    input('tax_rate', 'Corporate tax rate', deal.tax.rate, F_PCT1, '', 'tax.rate');
    input('dep_years', 'Tax depreciation (straight line)', deal.tax.depreciation_years, F_INT, 'years', 'tax.depreciation_years');
    section(wsI, 'Reserves and covenants');
    input('dsra_months', 'DSRA size', deal.dsra.months_of_ds, F_INT, 'months of debt service', 'dsra.months_of_ds');
    input('lockup_dscr', 'Lock-up DSCR (12-month historic)', deal.covenants.lockup_dscr, '0.00"x"', '', 'covenants.lockup_dscr');
    input('default_dscr', 'Event of default DSCR', deal.covenants.default_dscr, '0.00"x"', '', 'covenants.default_dscr');

    section(wsI, 'Derived (formulas, do not edit)');
    const derived = (name, label, f, v, fmt, unit) => {
      wsI.set('A' + r, label); wsI.set('B' + r, { f, v }, calc(fmt)); wsI.set('C' + r, unit || '', st.muted);
      wb.defineName(name, 'Inputs!$B$' + r); r++;
    };
    derived('all_in_rate', 'All-in interest rate', 'base_rate+margin', res.meta.rate, F_PCT, 'p.a.');
    derived('periods_per_year', 'Periods per year', '12/months_per_period', PPY, F_INT, '');
    derived('period_rate', 'Interest rate per period', 'all_in_rate/periods_per_year', res.meta.rate / PPY, '0.0000%', '');
    derived('n_constr', 'Construction periods', '((YEAR(cod_date)-YEAR(fc_date))*12+MONTH(cod_date)-MONTH(fc_date))/months_per_period', res.meta.nC, F_INT, '');
    derived('fy_first', 'Financial year of period 1 (Jul–Jun)', 'YEAR(EDATE(fc_date,months_per_period)-1)+IF(MONTH(EDATE(fc_date,months_per_period)-1)>6,1,0)', P[0].fy, F_INT, '');
    derived('dsra_periods', 'DSRA size in periods', 'dsra_months/months_per_period', deal.dsra.months_of_ds / PM, F_INT, '');

    // ---------------- Calc ----------------
    // Rows are declared once with a formula template; row numbers are assigned first so that
    // any row can refer to any other (e.g. opening debt = previous closing debt).
    const ROWS = [];
    const G = (title) => ROWS.push({ group: title });
    const L = (key, label, unit, fmt, f, opt) => ROWS.push(Object.assign({ key, label, unit, fmt, f }, opt || {}));
    const V = (key, label, unit, fmt, f, opt) => ROWS.push(Object.assign({ key, label, unit, fmt, f, scalar: true }, opt || {}));

    G('Timeline');
    L('start', 'Period start', '', F_DATE, (c) => 'EDATE(fc_date,months_per_period*(' + c.t + '-1))', { cache: (i) => X.excelDate(periodStart(P[i], PM)) });
    L('end', 'Period end', '', F_DATE, (c) => 'EDATE(fc_date,months_per_period*' + c.t + ')-1', { cache: (i) => X.excelDate(periodEnd(P[i])) });
    L('f_con', 'Construction flag', 'flag', F_INT, (c) => 'IF(' + c('end') + '<cod_date,1,0)', { cache: (i) => (i < res.meta.nC ? 1 : 0) });
    L('f_cod', 'COD flag (last construction period)', 'flag', F_INT, (c) => 'IF(AND(' + c('f_con') + '=1,' + c.n('f_con') + '=0),1,0)', { cache: (i) => (i === res.meta.nC - 1 ? 1 : 0) });
    L('f_op', 'Operations flag', 'flag', F_INT, (c) => 'IF(AND(' + c('start') + '>=cod_date,' + c('end') + '<=ppa_end_date),1,0)', { cache: (i) => (i >= res.meta.nC ? 1 : 0) });
    L('f_first', 'First operating period', 'flag', F_INT, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.p('f_op') + '=0),1,0)', { cache: (i) => (i === res.meta.nC ? 1 : 0) });
    L('f_last', 'Last operating period', 'flag', F_INT, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.n('f_op') + '=0),1,0)', { cache: (i) => (i === N - 1 ? 1 : 0) });
    L('f_debt', 'Debt repayment flag', 'flag', F_INT, (c) => 'IF(AND(' + c('start') + '>=cod_date,' + c('end') + '<=maturity_date),1,0)', { cache: (i) => (i >= res.meta.nC && i < res.meta.nM ? 1 : 0) });
    L('fy', 'Financial year (Jul–Jun)', 'FY', F_INT, (c) => 'YEAR(' + c('end') + ')+IF(MONTH(' + c('end') + ')>6,1,0)', { cache: (i) => P[i].fy });
    L('opyear', 'Operating year', '', F_INT, (c) => 'IF(' + c('f_op') + '=1,INT((' + c.t + '-n_constr-1)/periods_per_year)+1,0)', { cache: (i) => P[i].opYear });
    L('index', 'CPI index (annual steps from FC)', '', '0.0000', (c) => 'IF(' + c('f_op') + '=1,(1+cpi)^(' + c('fy') + '-fy_first),0)', { cache: (i) => S.index[i] });
    L('degf', 'Degradation factor', '', '0.0000', (c) => 'IF(' + c('f_op') + '=1,(1-degradation)^(' + c('opyear') + '-1),0)',
      { cache: (i) => (i >= res.meta.nC ? Math.pow(1 - O.degradation, P[i].opYear - 1) : 0) });

    G('Operations');
    L('energy', 'Net energy', 'MWh', F_K, (c) => 'IF(' + c('f_op') + '=1,capacity_mw*8760*cf_p50*availability*' + c('degf') + '/periods_per_year,0)', { cache: 'energy', total: true });
    L('price', 'PPA price', '/MWh', F_PRICE, (c) => 'IF(' + c('f_op') + '=1,ppa_price*(1+(' + c('index') + '-1)*ppa_indexation),0)',
      { cache: (i) => (S.energy[i] ? S.revenue[i] * 1000 / S.energy[i] : 0) });
    L('revenue', 'Revenue', 'k', F_K, (c) => c('energy') + '*' + c('price') + '/1000', { cache: 'revenue', total: true });
    L('opex', 'Operating costs', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,opex_real*' + c('index') + '/periods_per_year,0)', { cache: 'opex', total: true });
    L('ebitda', 'EBITDA', 'k', F_K, (c) => c('revenue') + '-' + c('opex'), { cache: 'ebitda', total: true, bold: true });
    L('cfadsPre', 'CFADS pre-tax', 'k', F_K, (c) => c('ebitda'), { cache: 'cfadsPre', total: true, bold: true });

    G('Debt sizing');
    L('pvPre', 'PV of remaining pre-tax CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('cfadsPre') + '+' + c.n('pvPre') + ')/(1+period_rate),0)',
      { cache: (i) => (res.series.llcrPre[i] != null ? res.series.llcrPre[i] * S.debtOpen[i] : 0) });
    V('pv_cfads_debt', 'PV of pre-tax CFADS over the debt tenor, at COD', 'k', F_K, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPre') + ')', { cache: sum.capacityAtTarget * D.target_dscr });
    V('capacity_at_target', 'Debt capacity at target DSCR', 'k', F_K, () => 'pv_cfads_debt/target_dscr', { cache: sum.capacityAtTarget });
    V('gearing_cap', 'Debt capacity at maximum gearing', 'k', F_K, () => 'capex_total*max_gearing', { cache: Cn.capex_total * D.gearing });
    V('debt_amount', 'Senior debt amount', 'k', F_K, () => 'IF(sizing="lower_of",MIN(gearing_cap,capacity_at_target),debt_amount_input)', { cache: sum.amount, bold: true });
    V('eff_gearing', 'Effective gearing (of capex)', '', F_PCT, () => 'debt_amount/capex_total', { cache: sum.effGearing });
    V('n_rep', 'Number of repayment periods', '', F_INT, (R) => 'SUM(' + R('f_debt') + ')', { cache: res.meta.nRep });
    V('sculpt_dscr', 'Sculpted DSCR (PV of CFADS ÷ debt)', '', F_X, () => 'pv_cfads_debt/debt_amount', { cache: sum.capacityAtTarget * D.target_dscr / sum.amount });
    V('annuity_pmt', 'Annuity payment per period (if annuity)', 'k', F_K, () => 'debt_amount*period_rate/(1-(1+period_rate)^-n_rep)',
      { cache: sum.amount * res.meta.rate / PPY / (1 - Math.pow(1 + res.meta.rate / PPY, -res.meta.nRep)) });

    G('Construction');
    L('capex', 'Capex', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,capex_total*INDEX(capex_profile,INT((' + c.t + '-1)*months_per_period/6)+1)*months_per_period/6,0)', { cache: 'capex', total: true });
    L('draw', 'Debt drawdown', 'k', F_K, (c) => c('capex') + '*eff_gearing', { cache: 'draw', total: true });
    L('idc', 'Interest during construction', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,' + c('debtOpen') + '*period_rate,0)', { cache: 'idc', total: true });
    L('commit', 'Commitment fee', 'k', F_K, (c) => 'IF(' + c('f_con') + '=1,(debt_amount-' + c('debtOpen') + ')*margin*commit_fee_pct/periods_per_year,0)', { cache: 'commit', total: true });
    L('upfront', 'Upfront fee', 'k', F_K, (c) => 'IF(' + c.t + '=1,debt_amount*upfront_fee_pct,0)', { cache: 'upfront', total: true });
    L('equityIn', 'Equity contribution', 'k', F_K, (c) => c('capex') + '-' + c('draw') + '+' + c('idc') + '+' + c('commit') + '+' + c('upfront'), { cache: 'equityIn', total: true, bold: true });
    V('project_cost', 'Total project cost', 'k', F_K, (R) => 'capex_total+SUM(' + R('idc') + ')+SUM(' + R('commit') + ')+SUM(' + R('upfront') + ')', { cache: sum.projectCost, bold: true });

    G('Senior debt');
    L('debtOpen', 'Opening balance', 'k', F_K, (c) => (c.first ? '0' : c.p('debtClose')), { cache: 'debtOpen' });
    L('interest', 'Interest', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,' + c('debtOpen') + '*period_rate,0)', { cache: 'interest', total: true });
    L('dsSched', 'Scheduled debt service (sculpted or annuity)', 'k', F_K,
      (c) => 'IF(' + c('f_debt') + '=1,IF(repayment="sculpted",' + c('cfadsPre') + '/sculpt_dscr,IF(repayment="annuity",annuity_pmt,0)),0)',
      { cache: (i) => (D.repayment === 'linear' ? 0 : S.debtService[i]) });
    L('principal', 'Principal', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,IF(repayment="linear",debt_amount/n_rep,' + c('dsSched') + '-' + c('interest') + '),0)', { cache: 'principal', total: true });
    L('debtService', 'Debt service', 'k', F_K, (c) => c('interest') + '+' + c('principal'), { cache: 'debtService', total: true, bold: true });
    L('debtClose', 'Closing balance', 'k', F_K, (c) => c('debtOpen') + '+' + c('draw') + '-' + c('principal'), { cache: 'debtClose' });

    G('Tax');
    L('depreciation', 'Tax depreciation', 'k', F_K, (c) => 'IF(AND(' + c('f_op') + '=1,' + c.t + '-n_constr<=dep_years*periods_per_year),project_cost/dep_years/periods_per_year,0)', { cache: 'depreciation', total: true });
    L('taxable', 'Taxable income', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c('ebitda') + '-' + c('depreciation') + '-' + c('interest') + ',0)', { cache: 'taxable', total: true });
    L('lossUsed', 'Tax losses used', 'k', F_K, (c) => 'IF(' + c('taxable') + '>0,MIN(' + c.p('lossBal') + ',' + c('taxable') + '),0)', { cache: 'lossUsed', total: true });
    L('lossBal', 'Tax losses carried forward', 'k', F_K, (c) => c.p('lossBal') + '+MAX(0,-' + c('taxable') + ')-' + c('lossUsed'), { cache: 'lossBal' });
    L('tax', 'Cash tax', 'k', F_K, (c) => 'IF(' + c('taxable') + '>0,(' + c('taxable') + '-' + c('lossUsed') + ')*tax_rate,0)', { cache: 'tax', total: true });
    L('cfadsPost', 'CFADS post-tax', 'k', F_K, (c) => c('ebitda') + '-' + c('tax'), { cache: 'cfadsPost', total: true, bold: true });

    G('Ratios');
    L('dscrPre', 'DSCR pre-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('cfadsPre') + '/' + c('debtService') + ',"")', { cache: 'dscrPre' });
    L('dscrPost', 'DSCR post-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('cfadsPost') + '/' + c('debtService') + ',"")', { cache: 'dscrPost', bold: true });
    // 12 months = the last 2 semesters or the last 4 quarters (the number of terms is fixed at export).
    const last12 = (c, key) => Array.from({ length: PPY }, (_, k) => c.back(key, k)).join('+');
    L('hDscr', '12-month historic DSCR (tested)', 'x', F_X,
      (c) => 'IF(AND(' + c('f_debt') + '=1,' + c.back('f_debt', PPY - 1) + '=1),(' + last12(c, 'cfadsPost') + ')/(' + last12(c, 'debtService') + '),"")',
      { cache: 'hDscr' });
    L('pvPost', 'PV of remaining post-tax CFADS to maturity', 'k', F_K, (c) => 'IF(' + c('f_debt') + '=1,(' + c('cfadsPost') + '+' + c.n('pvPost') + ')/(1+period_rate),0)',
      { cache: (i) => (res.series.llcrPost[i] != null ? res.series.llcrPost[i] * S.debtOpen[i] : 0) });
    L('llcrPre', 'LLCR pre-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('pvPre') + '/' + c('debtOpen') + ',"")', { cache: 'llcrPre' });
    L('llcrPost', 'LLCR post-tax', 'x', F_X, (c) => 'IF(' + c('f_debt') + '=1,' + c('pvPost') + '/' + c('debtOpen') + ',"")', { cache: 'llcrPost' });
    L('pvPpa', 'PV of remaining pre-tax CFADS to PPA expiry', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,(' + c('cfadsPre') + '+' + c.n('pvPpa') + ')/(1+period_rate),0)');
    V('llcr_cod_pre', 'LLCR at COD, pre-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPre') + ')/debt_amount', { cache: sum.llcrCodPre });
    V('llcr_cod_post', 'LLCR at COD, post-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPost') + ')/debt_amount', { cache: sum.llcrCodPost });
    V('plcr_cod', 'PLCR at COD, pre-tax', 'x', F_X, (R) => 'SUMPRODUCT(' + R('f_first') + ',' + R('pvPpa') + ')/debt_amount', { cache: sum.plcrCod });

    G('Debt service reserve account (DSRA)');
    L('dsraTarget', 'Target balance (next DSRA months of debt service)', 'k', F_K,
      (c) => 'IF(' + c.t + '>=n_constr,' + Array.from({ length: 12 / PM }, (_, k) => 'IF(dsra_periods>=' + (k + 1) + ',' + c.ahead('debtService', k + 1) + ',0)').join('+') + ',0)', { cache: 'dsraTarget' });
    L('dsraFund', 'Initial funding by equity at COD', 'k', F_K, (c) => 'IF(' + c('f_cod') + '=1,' + c('dsraTarget') + ',0)', { cache: 'dsraFund', total: true });
    L('cashAfterDs', 'Cash after debt service', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c('cfadsPost') + '-' + c('debtService') + ',0)', { cache: 'cashAfterDs', total: true });
    L('dsraDraw', 'DSRA drawing to cover a shortfall', 'k', F_K, (c) => 'IF(' + c('cashAfterDs') + '<0,MIN(' + c.p('dsraBal') + ',-' + c('cashAfterDs') + '),0)', { cache: 'dsraDraw', total: true });
    L('shortfall', 'Debt service shortfall after DSRA', 'k', F_K, (c) => 'MAX(0,-' + c('cashAfterDs') + '-' + c('dsraDraw') + ')', { cache: 'shortfall', total: true });
    L('cash1', 'Cash after DSRA drawing', 'k', F_K, (c) => 'MAX(0,' + c('cashAfterDs') + '+' + c('dsraDraw') + ')');
    L('dsraTopUp', 'Top-up from cash flow', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,MAX(0,MIN(' + c('dsraTarget') + '-(' + c.p('dsraBal') + '-' + c('dsraDraw') + '),' + c('cash1') + ')),0)', { cache: 'dsraTopUp', total: true });
    L('dsraRel', 'Release of excess over target', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,MAX(0,(' + c.p('dsraBal') + '-' + c('dsraDraw') + ')-' + c('dsraTarget') + '),0)', { total: true });
    L('dsraAfter', 'Balance after top-up and release', 'k', F_K, (c) => 'IF(' + c('f_op') + '=1,' + c.p('dsraBal') + '-' + c('dsraDraw') + '+' + c('dsraTopUp') + '-' + c('dsraRel') + ',0)');
    L('dsraFinal', 'Final release at PPA expiry', 'k', F_K, (c) => 'IF(' + c('f_last') + '=1,' + c('dsraAfter') + ',0)', { total: true });
    L('dsraBal', 'DSRA closing balance', 'k', F_K, (c) => 'IF(' + c('f_cod') + '=1,' + c('dsraFund') + ',IF(' + c('f_op') + '=1,' + c('dsraAfter') + '-' + c('dsraFinal') + ',0))', { cache: 'dsraBal', bold: true });

    G('Cash flow waterfall and equity');
    L('cashAvail', 'Cash available after DSRA', 'k', F_K, (c) => c('cash1') + '-' + c('dsraTopUp') + '+' + c('dsraRel'), { total: true });
    L('lockup', 'Lock-up flag (12m DSCR below lock-up)', 'flag', F_INT, (c) => 'IF(' + c('hDscr') + '="",0,IF(' + c('hDscr') + '<lockup_dscr,1,0))', { cache: (i) => (S.lockup[i] ? 1 : 0), total: true });
    L('defaultF', 'Default flag (12m DSCR below default level)', 'flag', F_INT, (c) => 'IF(' + c('hDscr') + '="",0,IF(' + c('hDscr') + '<default_dscr,1,0))', { cache: (i) => (S.defaultFlag[i] ? 1 : 0), total: true });
    L('blocked', 'Distribution blocked (lock-up or DSRA not full)', 'flag', F_INT, (c) => 'IF(' + c('f_op') + '=1,IF(OR(' + c('lockup') + '=1,' + c('dsraAfter') + '<' + c('dsraTarget') + '-0.000001),1,0),0)');
    L('trapped', 'Cash trapped this period', 'k', F_K, (c) => 'IF(' + c('blocked') + '=1,' + c('cashAvail') + ',0)', { cache: 'trapped', total: true });
    L('trapBal', 'Trapped cash balance', 'k', F_K, (c) => 'IF(' + c('f_last') + '=1,0,IF(' + c('blocked') + '=1,' + c.p('trapBal') + '+' + c('cashAvail') + ',0))', { cache: 'trapBal' });
    L('distributions', 'Distributions to equity', 'k', F_K,
      (c) => 'IF(' + c('f_op') + '=1,IF(' + c('blocked') + '=1,0,' + c('cashAvail') + '+' + c.p('trapBal') + ')+IF(' + c('f_last') + '=1,IF(' + c('blocked') + '=1,' + c.p('trapBal') + '+' + c('cashAvail') + ',0)+' + c('dsraFinal') + ',0),0)',
      { cache: 'distributions', total: true, bold: true });
    L('equityFlows', 'Equity cash flow', 'k', F_K, (c) => '-' + c('equityIn') + '-' + c('dsraFund') + '+' + c('distributions'), { cache: 'equityFlows', total: true, bold: true });
    V('equity_irr', 'Equity IRR (post-tax, nominal, annualised)', '', F_PCT, (R) => '(1+IRR(' + R('equityFlows') + '))^periods_per_year-1', { cache: sum.equityIrr, bold: true });

    // Assign row numbers: rows 1-6 hold the title and the timeline header.
    const ROW = {};
    let rr = 8;
    for (const row of ROWS) {
      if (row.group) { row.r = ++rr; rr++; continue; }
      row.r = rr++;
      ROW[row.key] = row.r;
    }
    ROW.t = 4;

    const wsC = wb.addSheet('Calc');
    wsC.width(1, 46).width(2, 7).width(3, 14);
    for (let t = 1; t <= N; t++) wsC.width(FIRST + t - 1, 11);
    wsC.freezeAt('D7');
    wsC.tab = C_MUTED;
    banner(wsC, 'Calculations', (PM === 3 ? 'Quarterly' : 'Semi-annual') + ' periods across. ' + deal.meta.currency + ' thousands unless stated. Every cell is a formula; inputs come from the Inputs sheet by name. Shaded columns = construction.', FIRST + N - 1);
    wsC.set('A4', 'Period', st.head); wsC.set('B4', '', st.head); wsC.set('C4', 'Total / value', st.headR);
    wsC.set('A5', 'Period start', st.muted); wsC.set('A6', 'Period end', st.bold);
    const boldTop = wb.style({ bold: true, border: 'top' });
    const hdr = wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'right', numFmt: F_INT });
    const dStart = wb.style({ numFmt: F_MONTH, color: C_MUTED, align: 'right' });
    const dEnd = wb.style({ numFmt: F_MONTH, bold: true, align: 'right' });
    const hdrCon = wb.style({ bold: true, fill: C_CONSTR_BG, border: 'bottom', align: 'right', numFmt: F_INT });
    const phase = wb.style({ color: C_MUTED, align: 'right', italic: true });
    wsC.set('A3', 'Phase', st.muted);
    for (let t = 1; t <= N; t++) {
      const con = t <= res.meta.nC;
      wsC.set(col(t) + '3', con ? 'Constr.' : t <= res.meta.nM ? 'Debt' : 'Tail', phase);
      wsC.set(col(t) + '4', t, con ? hdrCon : hdr);
    }

    const rangeOf = (key) => '$' + col(1) + '$' + ROW[key] + ':$' + LAST + '$' + ROW[key];
    for (const row of ROWS) {
      if (row.group) {
        wsC.set('A' + row.r, row.group, st.group);
        for (let c = 2; c <= FIRST + N - 1; c++) wsC.set(X.colName(c) + row.r, '', st.group);
        continue;
      }
      const lab = row.bold ? boldTop : 0;
      wsC.set('A' + row.r, row.label, lab);
      wsC.set('B' + row.r, row.unit, st.muted);
      if (row.scalar) {
        wsC.set('C' + row.r, { f: row.f((key) => 'Calc!' + rangeOf(key)), v: row.cache }, calc(row.fmt, true));
        wb.defineName(row.key, 'Calc!$C$' + row.r);
        continue;
      }
      const style = row.key === 'start' ? dStart : row.key === 'end' ? dEnd : (row.bold ? wb.style({ numFmt: row.fmt, bold: true, border: 'top' }) : calc(row.fmt));
      for (let t = 1; t <= N; t++) {
        const cl = col(t), i = t - 1;
        const ref = (key) => cl + ROW[key];
        ref.t = cl + '$4';
        ref.first = t === 1;
        ref.p = (key) => (t === 1 ? '0' : col(t - 1) + ROW[key]);
        ref.n = (key) => col(t + 1) + ROW[key]; // the column after the last period is empty, i.e. 0
        ref.back = (key, k) => (k === 0 ? cl + ROW[key] : t - k < 1 ? '0' : col(t - k) + ROW[key]);
        ref.ahead = (key, k) => col(t + k) + ROW[key]; // columns after the last period are empty, i.e. 0
        let v;
        if (typeof row.cache === 'string') v = S[row.cache][i];
        else if (typeof row.cache === 'function') v = row.cache(i);
        if (v === null) v = '';
        if (typeof v === 'boolean') v = v ? 1 : 0;
        wsC.set(cl + row.r, { f: row.f(ref), v }, style);
      }
      if (row.total) {
        const tv = typeof row.cache === 'string' ? S[row.cache].reduce((a, b) => a + (b || 0), 0) : undefined;
        wsC.set('C' + row.r, { f: 'SUM(' + col(1) + row.r + ':' + LAST + row.r + ')', v: tv }, calc(row.fmt, true));
      }
    }
    const R = (key) => 'Calc!' + rangeOf(key);

    // ---------------- Outputs ----------------
    const wsO = wb.addSheet('Outputs');
    wsO.width(1, 44).width(2, 16).width(3, 4).width(4, 34).width(5, 16);
    wsO.tab = C_GOOD;
    banner(wsO, 'Outputs', 'Key results, all linked to the Calc sheet.', 5);
    const out = (ws, row, label, f, v, fmt, colL, colV) => { ws.set((colL || 'A') + row, label); ws.set((colV || 'B') + row, { f, v }, calc(fmt, false)); };
    wsO.set('A4', 'Uses', st.head); wsO.set('B4', deal.meta.currency + ' k', st.headR);
    wsO.set('D4', 'Sources', st.head); wsO.set('E4', deal.meta.currency + ' k', st.headR);
    out(wsO, 5, 'Construction capex', 'capex_total', Cn.capex_total, F_K);
    out(wsO, 6, 'Interest during construction', 'SUM(' + R('idc') + ')', sum.idc, F_K);
    out(wsO, 7, 'Upfront fee', 'SUM(' + R('upfront') + ')', sum.upfrontFee, F_K);
    out(wsO, 8, 'Commitment fee', 'SUM(' + R('commit') + ')', sum.commitmentFees, F_K);
    wsO.set('A9', 'Total uses', st.bold); wsO.set('B9', { f: 'SUM(B5:B8)', v: sum.projectCost }, calc(F_K, true));
    out(wsO, 5, 'Senior debt', 'debt_amount', sum.amount, F_K, 'D', 'E');
    out(wsO, 6, 'Sponsor equity at financial close', 'SUM(' + R('equityIn') + ')', sum.equityAtFc, F_K, 'D', 'E');
    wsO.set('D9', 'Total sources', st.bold); wsO.set('E9', { f: 'E5+E6', v: sum.amount + sum.equityAtFc }, calc(F_K, true));
    wsO.set('A10', 'The DSRA is funded by a further equity contribution at COD:', st.muted);
    wsO.set('B10', { f: 'SUM(' + R('dsraFund') + ')', v: sum.dsraAtCod }, calc(F_K));

    wsO.set('A12', 'Debt and cover ratios', st.head); wsO.set('B12', '', st.headR);
    const ratio = [
      ['Senior debt', 'debt_amount', sum.amount, F_K],
      ['Effective gearing (of capex)', 'eff_gearing', sum.effGearing, F_PCT1],
      ['Sculpted DSCR (pre-tax)', 'IF(repayment="sculpted",sculpt_dscr,"n/a")', D.repayment === 'sculpted' ? sum.sculptDscr : 'n/a', F_X],
      ['Minimum DSCR pre-tax', 'MIN(' + R('dscrPre') + ')', sum.minDscrPre, F_X],
      ['Average DSCR pre-tax', 'AVERAGE(' + R('dscrPre') + ')', sum.avgDscrPre, F_X],
      ['Minimum DSCR post-tax', 'MIN(' + R('dscrPost') + ')', sum.minDscrPost, F_X],
      ['Average DSCR post-tax', 'AVERAGE(' + R('dscrPost') + ')', sum.avgDscrPost, F_X],
      ['LLCR at COD, pre-tax', 'llcr_cod_pre', sum.llcrCodPre, F_X],
      ['LLCR at COD, post-tax', 'llcr_cod_post', sum.llcrCodPost, F_X],
      ['PLCR at COD, pre-tax', 'plcr_cod', sum.plcrCod, F_X],
      ['Debt capacity at target DSCR', 'capacity_at_target', sum.capacityAtTarget, F_K],
      ['Headroom vs senior debt', 'capacity_at_target/debt_amount-1', sum.headroom, F_PCT1],
      ['Lock-up tests failed', 'SUM(' + R('lockup') + ')', sum.lockupPeriods, F_INT],
      ['Default tests failed', 'SUM(' + R('defaultF') + ')', sum.defaultPeriods, F_INT],
    ];
    ratio.forEach(([l, f, v, fmt], i) => out(wsO, 13 + i, l, f, v, fmt));
    let ro = 13 + ratio.length + 1;
    wsO.set('A' + ro, 'First operating year', st.head); wsO.set('B' + ro, '', st.headR); ro++;
    const fy = sum.firstOpYear;
    const firstYear = (key) => 'SUMIF(' + R('opyear') + ',1,' + R(key) + ')';
    out(wsO, ro++, 'Net energy (MWh)', firstYear('energy'), fy.energy, F_K);
    out(wsO, ro++, 'Revenue', firstYear('revenue'), fy.revenue, F_K);
    out(wsO, ro++, 'Operating costs', firstYear('opex'), fy.opex, F_K);
    out(wsO, ro++, 'EBITDA', firstYear('ebitda'), fy.ebitda, F_K);
    ro++;
    wsO.set('A' + ro, 'Equity', st.head); wsO.set('B' + ro, '', st.headR); ro++;
    out(wsO, ro++, 'Equity IRR (post-tax, nominal)', 'equity_irr', sum.equityIrr, F_PCT);
    out(wsO, ro++, 'Total distributions', 'SUM(' + R('distributions') + ')', sum.totalDistributions, F_K);

    // ---------------- Checks ----------------
    const wsK = wb.addSheet('Checks');
    wsK.width(1, 52).width(2, 16).width(3, 16).width(4, 14).width(5, 10);
    wsK.tab = C_GOOD;
    banner(wsK, 'Checks', 'TRUE means the check passes. Part 2 compares this workbook with the web app at the time of export.', 5);
    const okStyle = st.ok;
    wsK.set('A4', 'Model integrity', st.head); wsK.set('B4', 'Result', wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'center' }));
    const lastCol = (key) => 'Calc!$' + LAST + '$' + ROW[key];
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
    ];
    integrity.forEach(([l, f], i) => { wsK.set('A' + (5 + i), l); wsK.set('B' + (5 + i), { f, v: true }, okStyle); });
    let rk = 5 + integrity.length;
    wsK.set('A' + rk, 'All integrity checks pass', st.bold);
    wsK.set('B' + rk, { f: 'AND(B5:B' + (rk - 1) + ')', v: true }, wb.style({ bold: true, align: 'center', color: C_GOOD }));
    rk += 2;
    wsK.set('A' + rk, 'Reconciliation with the web app', st.head);
    wsK.set('B' + rk, 'Web app', st.headR); wsK.set('C' + rk, 'This workbook', st.headR); wsK.set('D' + rk, 'Difference', st.headR);
    wsK.set('E' + rk, 'Match', wb.style({ bold: true, fill: C_HEAD_BG, border: 'bottom', align: 'center' }));
    rk++;
    wsK.set('A' + rk, 'App values are frozen at export. If you change an input, differences here are expected.', st.note); rk++;
    const recon = [
      ['Senior debt', sum.amount, 'debt_amount', F_K],
      ['Total project cost', sum.projectCost, 'project_cost', F_K],
      ['Interest during construction', sum.idc, 'SUM(' + R('idc') + ')', F_K],
      ['Equity at financial close', sum.equityAtFc, 'SUM(' + R('equityIn') + ')', F_K],
      ['First-year EBITDA', fy.ebitda, firstYear('ebitda'), F_K],
      ['Debt capacity at target DSCR', sum.capacityAtTarget, 'capacity_at_target', F_K],
      ['Minimum DSCR post-tax', sum.minDscrPost, 'MIN(' + R('dscrPost') + ')', F_X],
      ['Average DSCR post-tax', sum.avgDscrPost, 'AVERAGE(' + R('dscrPost') + ')', F_X],
      ['LLCR at COD, pre-tax', sum.llcrCodPre, 'llcr_cod_pre', F_X],
      ['LLCR at COD, post-tax', sum.llcrCodPost, 'llcr_cod_post', F_X],
      ['PLCR at COD', sum.plcrCod, 'plcr_cod', F_X],
      ['DSRA at COD', sum.dsraAtCod, 'SUM(' + R('dsraFund') + ')', F_K],
      ['Total cash tax', S.tax.reduce((a, b) => a + b, 0), 'SUM(' + R('tax') + ')', F_K],
      ['Total distributions', sum.totalDistributions, 'SUM(' + R('distributions') + ')', F_K],
      ['Equity IRR', sum.equityIrr, 'equity_irr', F_PCT],
    ];
    const reconStart = rk;
    recon.forEach(([l, app, f, fmt]) => {
      wsK.set('A' + rk, l);
      wsK.set('B' + rk, app, calc(fmt));
      wsK.set('C' + rk, { f, v: app }, calc(fmt));
      wsK.set('D' + rk, { f: 'C' + rk + '-B' + rk, v: 0 }, calc(fmt === F_K ? '#,##0.000;(#,##0.000);"–"' : '0.000000'));
      wsK.set('E' + rk, { f: 'ABS(D' + rk + ')<=MAX(0.001,ABS(B' + rk + ')*0.000001)', v: true }, okStyle);
      rk++;
    });
    wsK.set('A' + rk, 'Workbook matches the web app', st.bold);
    wsK.set('E' + rk, { f: 'AND(E' + reconStart + ':E' + (rk - 1) + ')', v: true }, wb.style({ bold: true, align: 'center', color: C_GOOD }));
    const reconAll = 'Checks!$E$' + rk;
    const integrityAll = 'Checks!$B$' + (5 + integrity.length);

    // ---------------- Cover (first tab) ----------------
    const wsV = wb.addSheet('Cover'); // built last (it links to the other sheets), then moved to the front
    wsV.width(1, 34).width(2, 70);
    wsV.tab = C_NAVY;
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
      ['Senior debt', 'debt_amount', sum.amount, F_K],
      ['Minimum DSCR post-tax', 'MIN(' + R('dscrPost') + ')', sum.minDscrPost, F_X],
      ['LLCR at COD, post-tax', 'llcr_cod_post', sum.llcrCodPost, F_X],
      ['Equity IRR', 'equity_irr', sum.equityIrr, F_PCT],
      ['Integrity checks pass', integrityAll, true, ''],
      ['Matches the web app', reconAll, true, ''],
    ];
    live.forEach(([k, f, v, fmt], i) => { wsV.set('A' + (10 + i), k); wsV.set('B' + (10 + i), { f, v }, wb.style({ numFmt: fmt || undefined, bold: true, size: 12, color: fmt ? C_ACCENT : C_GOOD, align: 'left' })); });
    wsV.set('A17', 'How the workbook is built', st.h);
    const guide = [
      ['Inputs', 'Every assumption, in blue, with the term sheet clause it comes from. Named cells (e.g. margin, target_dscr) are used in all formulas.'],
      ['Calc', 'The full ' + PERIOD_WORD + ' model: timeline flags, operations, debt sizing and sculpting, construction funding, debt schedule, tax with losses carried forward, ratios, DSRA and the distribution waterfall.'],
      ['Outputs', 'Sources and uses, DSCR / LLCR / PLCR, debt capacity, first-year figures, equity IRR.'],
      ['Checks', 'Integrity checks (debt drawn and repaid, sources = uses, no shortfall) and a reconciliation against the web app.'],
      ['Colour code', 'Blue on light blue = input. Black = formula. Nothing is hard-coded in Calc.'],
      ['Simplifications (v1)', 'Interest = balance × annual rate / periods per year (no day count). No hedging, agency fees, MRA, cash sweep or terminal value. See the Methodology tab of the web app.'],
    ];
    guide.forEach(([k, v], i) => { wsV.set('A' + (18 + i), k, st.bold); wsV.set('B' + (18 + i), v, st.noteWrap); });
    wb.sheets.unshift(wb.sheets.pop());

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

  const api = { buildWorkbook, exportXlsx };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFExcel = api;
})(typeof window !== 'undefined' ? window : globalThis);
