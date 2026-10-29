/*
 * Term sheet writer — v1
 * Turns a generated deal (plus the engine results, for sources & uses and base case figures)
 * into an indicative term sheet in Markdown, laid out like docs/coolabah_term_sheet.md.
 * Every number the model reads is written in the clause that deal.sources points to.
 */
(function (root) {
  'use strict';

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const longDate = (s) => { const [y, m, d] = s.split('-').map(Number); return d + ' ' + MONTHS[m - 1] + ' ' + y; };
  const nf = (v) => Math.round(v).toLocaleString('en-US');
  const pct = (v, n = 2) => (v * 100).toFixed(n) + '%';
  const m1 = (k) => (k / 1000).toFixed(1) + 'm';
  const m2 = (k) => (k / 1000).toFixed(2) + 'm';

  function termSheet(deal, res) {
    const M = deal.meta, T = deal.timeline, C = deal.construction, O = deal.operations, D = deal.senior_debt;
    const s = res.summary, cov = deal.covenants;
    const nC = C.capex_profile.length;
    const rate = D.base_rate + D.margin;
    const debtYears = (res.meta.nM - res.meta.nC) / 2;
    const tail = (res.meta.N - res.meta.nM) / 2;
    const [sp, mo] = M.offtaker_rating;
    const fy = s.firstOpYear;
    const price1 = fy.revenue * 1000 / fy.energy;
    const indexed = O.ppa_indexation > 0;
    const fp = res.periods[res.meta.nC];
    const firstRepay = (fp.month === 6 ? '30 June ' : '31 December ') + fp.year;
    const L = [];

    L.push('# ' + M.name + ' — Indicative Term Sheet', '');
    L.push('Deal #' + M.seed + ' · generated at random for modelling practice', '');
    L.push('Senior secured, limited-recourse financing of AUD ' + m1(D.amount) + ' for a ' + O.capacity_mw + ' MWac ' + M.technology +
      ' solar farm in ' + M.state + ', Australia, backed by a ' + M.ppa_years + '-year ' + (indexed ? 'CPI-indexed' : 'fixed-price') +
      ' PPA. Financial close targeted ' + longDate(T.financial_close) + '; COD ' + longDate(T.cod) + '.', '');
    L.push('*Fictional transaction generated at random. All parties, names and figures are illustrative. Indicative terms for discussion only; not an offer or commitment to lend. All amounts in AUD unless stated.*', '');

    L.push('## 1. Transaction overview', '');
    L.push('| Term | Description |', '| --- | --- |');
    L.push('| Project | ' + M.plant + ': ' + O.capacity_mw + ' MWac ' + M.technology + ' PV plant, ' + M.zone + ', ' + M.state + '. Connected to the ' + M.grid + '. |');
    L.push('| Borrower | ' + M.plant.replace(' Farm', '') + ' Project Co Pty Ltd, a single-purpose company with no other business or assets. |');
    L.push('| Sponsor | ' + M.sponsor + ' (100% equity). |');
    L.push('| Offtaker | Investment-grade energy retailer (rated ' + sp + ' / ' + mo + '), under a ' + M.ppa_years + '-year PPA. |');
    L.push('| PPA | Pay-as-produced, bundled energy and LGCs. AUD ' + O.ppa_price_fc.toFixed(2) + '/MWh at financial close, ' +
      (indexed ? 'indexed at 100% of CPI annually' : 'flat price, no indexation') + '. Tenor ' + M.ppa_years + ' years from COD, ending ' + longDate(T.ppa_end) + '. |');
    L.push('| EPC Contractor | Tier-1 contractor under a fixed-price, date-certain, lump-sum turnkey contract. |');
    L.push('| Mandated Lead Arrangers | Bank A, Bank B and Bank C (club deal). |');
    L.push('| Purpose | Fund construction capex, interest during construction, financing fees and commitment fees until COD. |', '');

    L.push('## 2. Facilities', '');
    L.push('| Term | Description |', '| --- | --- |');
    L.push('| Senior Facility | AUD ' + nf(D.amount * 1000) + ' Construction Facility, converting at COD into a Term Loan Facility of the same amount. |');
    L.push('| Availability Period | From financial close (' + longDate(T.financial_close) + ') to COD. |');
    L.push('| Drawdowns | Semi-annual, pro rata with equity at a fixed gearing of ' + pct(s.effGearing, 1) + ' of each capex amount. Equity funds all financing costs (upfront fee, IDC, commitment fee). |');
    L.push('| Final Maturity | ' + longDate(T.debt_maturity) + ': ' + debtYears + ' years after scheduled COD, leaving a ' + tail + '-year PPA tail. |');
    L.push('| Currency | AUD. |', '');
    L.push('**Sources and uses at financial close (AUD k)**', '');
    const tot = s.projectCost;
    const p = (v) => pct(v / tot, 1);
    L.push('| Uses | AUD k | % | Sources | AUD k | % |', '| --- | --- | --- | --- | --- | --- |');
    L.push('| Construction capex | ' + nf(C.capex_total) + ' | ' + p(C.capex_total) + ' | Senior Facility | ' + nf(s.amount) + ' | ' + p(s.amount) + ' |');
    L.push('| Interest during construction | ' + nf(s.idc) + ' | ' + p(s.idc) + ' | Sponsor equity | ' + nf(s.equityAtFc) + ' | ' + p(s.equityAtFc) + ' |');
    L.push('| Upfront fee | ' + nf(s.upfrontFee) + ' | ' + p(s.upfrontFee) + ' |  |  |  |');
    L.push('| Commitment fee | ' + nf(s.commitmentFees) + ' | ' + p(s.commitmentFees) + ' |  |  |  |');
    L.push('| **Total uses** | **' + nf(tot) + '** | **100.0%** | **Total sources** | **' + nf(s.amount + s.equityAtFc) + '** | **100.0%** |', '');
    L.push('Capex is drawn ' + C.capex_profile.map((x) => Math.round(x * 100) + '%').join(' / ') + ' over the ' + nC + ' construction semesters.', '');

    L.push('## 3. Pricing and fees', '');
    L.push('All-in cost of debt is ' + pct(rate) + ' p.a. in the base case (BBSW ' + pct(D.base_rate) + ' + ' + pct(D.margin) + ' margin), flat over the life of the loan.', '');
    L.push('| Term | Description |', '| --- | --- |');
    L.push('| Interest Rate | 6-month BBSW plus the Margin, payable semi-annually in arrears on 30 June and 31 December. |');
    L.push('| Margin | ' + pct(D.margin) + ' p.a. during construction and operations. |');
    L.push('| Base case BBSW | ' + pct(D.base_rate) + ' p.a. (lenders\' planning assumption). |');
    L.push('| Upfront Fee | ' + pct(D.upfront_fee) + ' of the Senior Facility (AUD ' + nf(s.upfrontFee) + 'k), payable on financial close. |');
    L.push('| Commitment Fee | ' + pct(D.commitment_fee_pct_margin, 0) + ' of the Margin (' + pct(D.margin * D.commitment_fee_pct_margin) + ' p.a.) on undrawn commitments, payable semi-annually in arrears during the Availability Period. |');
    L.push('| Agency Fee | AUD ' + D.agency_fee + 'k p.a. in total (Facility Agent and Security Trustee), payable while the Senior Facility is outstanding. |');
    L.push('| Interest Rate Hedging | The Borrower must hedge ' + pct(D.hedge_ratio, 0) + ' of projected debt outstanding for ' + D.hedge_years + ' years from financial close, through interest rate swaps with the Lenders. Base case swap rate of ' + pct(D.swap_rate) + ' p.a. |', '');

    L.push('## 4. Repayment, accounts and financial covenants', '');
    L.push('| Term | Description |', '| --- | --- |');
    L.push('| Repayment | ' + (res.meta.nRep) + ' semi-annual instalments from ' + firstRepay + ' to ' + longDate(T.debt_maturity) + ', sculpted to the P50 CFADS profile of the Agreed Base Case. |');
    L.push('| Target DSCR (sculpting) | ' + D.target_dscr.toFixed(2) + 'x on P50 CFADS. |');
    L.push('| Sizing | Senior debt is the lowest of: (i) a maximum of ' + pct(D.gearing, 0) + ' of construction capex; (ii) the amount sculpted at ' + D.target_dscr.toFixed(2) + 'x on pre-tax P50 CFADS; (iii) the amount sculpted at 1.00x on one-year P99 CFADS; (iv) the amount for which the minimum post-tax DSCR of the Agreed Base Case is at least ' + (cov.lockup_dscr + 0.05).toFixed(2) + 'x. Binding constraint: ' + M.sizing_binding + '. |');
    L.push('| Debt Service Reserve Account | Balance equal to the next 6 months of scheduled senior debt service, funded in cash by the Sponsor at COD. |');
    const MR = deal.maintenance;
    L.push('| Maintenance Reserve Account | Funded from cash flow over the ' + MR.funding_years + ' years before the inverter replacement (AUD ' + (MR.cost_real / 1000).toFixed(2) + 'm real, spent in operating years ' + MR.from_year + ' to ' + MR.to_year + '). |', '');
    L.push('**Financial covenants (tested semi-annually from the second calculation date after COD)**', '');
    L.push('| Ratio | Calculation | Level |', '| --- | --- | --- |');
    L.push('| Historic DSCR | CFADS over the last 12 months ÷ senior debt service over the same period | Lock-up < ' + cov.lockup_dscr.toFixed(2) + 'x; Event of Default < ' + cov.default_dscr.toFixed(2) + 'x |');
    L.push('| Maximum gearing | Senior debt ÷ construction capex at financial close | ' + pct(D.gearing, 0) + ' |', '');
    L.push('**Cash flow waterfall (operations)**', '');
    L.push('1. Operating costs and taxes', '2. Agency fees', '3. Senior interest (net of swap payments)', '4. Senior principal', '5. Top-up of the Debt Service Reserve Account', '6. Funding of the Maintenance Reserve Account', '7. Distribution Account, subject to the Distribution Conditions', '');
    L.push('**Distribution Conditions:** no Default outstanding; historic DSCR at or above ' + cov.lockup_dscr.toFixed(2) + 'x; DSRA fully funded.', '');

    L.push('## 5. Security and conditions', '');
    L.push('| Term | Description |', '| --- | --- |');
    L.push('| Security | First-ranking security over all assets of the Borrower, shares in the Borrower, and tripartite deeds with the Offtaker and EPC Contractor. |');
    L.push('| Sponsor Support | Equity commitment letter for AUD ' + m1(s.equityAtFc) + ', backed by a bank letter of credit until fully contributed. |');
    L.push('| Conditions Precedent | Finance and Project documents, satisfactory technical, legal, insurance and model audit reports, and an Agreed Base Case consistent with these terms. |', '');

    L.push('## 6. Agreed Base Case', '');
    L.push('The base case generates EBITDA of AUD ' + m2(fy.ebitda) + ' in the first year of operations, a ' + pct(fy.margin, 0) + ' margin.', '');
    L.push('| Assumption | Base case | First operating year |', '| --- | --- | --- |');
    L.push('| Capacity | ' + O.capacity_mw + ' MWac |  |');
    L.push('| P50 capacity factor | ' + pct(O.capacity_factor_p50, 1) + ' (' + M.technology + ') |  |');
    L.push('| Availability | ' + pct(O.availability, 1) + ' |  |');
    L.push('| Energy yield (one-year) | P90 ' + pct(O.p90_factor, 1) + ' / P99 ' + pct(O.p99_factor, 1) + ' of P50 energy |  |');
    L.push('| Degradation | ' + pct(O.degradation) + ' p.a. from operating year 2 |  |');
    L.push('| Net energy |  | ' + nf(fy.energy) + ' MWh |');
    L.push('| PPA price | AUD ' + O.ppa_price_fc.toFixed(2) + '/MWh at FC' + (indexed ? ', 100% CPI' : ', flat') + ' | AUD ' + price1.toFixed(2) + '/MWh |');
    L.push('| Revenue |  | AUD ' + m2(fy.revenue) + ' |');
    L.push('| Operating costs | AUD ' + (O.opex_real_fc / 1000).toFixed(3) + 'm p.a. real | AUD ' + m2(fy.opex) + ' |');
    L.push('| EBITDA |  | AUD ' + m2(fy.ebitda) + ' |');
    L.push('| CPI | ' + pct(O.cpi) + ' p.a., applied annually |  |');
    L.push('| Corporate tax | 30%, straight-line tax depreciation over ' + deal.tax.depreciation_years + ' years, losses carried forward |  |');

    return L.join('\n') + '\n';
  }

  const api = { termSheet };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFTermSheet = api;
})(typeof window !== 'undefined' ? window : globalThis);
