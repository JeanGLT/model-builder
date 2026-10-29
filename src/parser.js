/*
 * Term sheet reader — v1
 * Reads a term sheet written in the Model Builder layout (Markdown tables, as in
 * docs/coolabah_term_sheet.md or a generated deal) and returns a deal JSON.
 * No AI: each input is found by the label of its table row (e.g. "| Margin | 1.75% p.a. |").
 * Free-form term sheets (PDF, Word, other layouts) go through the Claude prompt instead.
 */
(function (root) {
  'use strict';

  const MONTHS = { january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12 };
  const num = (s) => parseFloat(String(s).replace(/,/g, ''));
  const pct = (s) => num(s) / 100;

  // "1 July 2027" or "2027-07-01" -> "2027-07-01"
  function isoDate(s) {
    let m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[0];
    m = s.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m && MONTHS[m[2].toLowerCase()]) return m[3] + '-' + String(MONTHS[m[2].toLowerCase()]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
    return null;
  }
  const DATE = '(\\d{1,2}\\s+[A-Za-z]+\\s+\\d{4}|\\d{4}-\\d{2}-\\d{2})';

  // Splits the document into table rows, remembering the section ("Section 2") each row sits in.
  function rowsOf(md) {
    const rows = [];
    let section = 'Summary';
    for (const line of md.replace(/\r/g, '').split('\n')) {
      const h = line.match(/^##\s+(\d+)\./);
      if (h) { section = 'Section ' + h[1]; continue; }
      if (!line.trim().startsWith('|') || /^\|\s*-/.test(line.trim())) continue;
      const cells = line.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim().replace(/\*\*/g, ''));
      rows.push({ section, label: cells[0], cells, text: cells.slice(1).join(' | ') });
    }
    return rows;
  }

  function parseTermSheet(md) {
    const rows = rowsOf(md);
    const flat = md.replace(/\*\*/g, '');
    const deal = {
      meta: { name: 'Uploaded deal', currency: 'AUD', units: 'k', uploaded: true },
      timeline: {}, construction: {}, operations: {}, tax: {}, senior_debt: { sizing: 'fixed', repayment: 'sculpted' },
      dsra: {}, covenants: {}, maintenance: {}, sources: {},
    };
    const missing = [], assumed = [];
    const short = (s) => (s.length > 90 ? s.slice(0, 88) + '…' : s);

    // Finds the first row whose label matches, then reads a value from its text with a regex.
    function take(path, labelRe, valueRe, convert, opts) {
      opts = opts || {};
      for (const r of rows) {
        if (!labelRe.test(r.label)) continue;
        const m = (opts.cells ? r.cells.slice(1).join(' | ') : r.text).match(valueRe);
        if (!m) continue;
        const v = convert(m[1], m);
        if (v === null || v === undefined || (typeof v === 'number' && isNaN(v))) continue;
        setPath(path, v);
        deal.sources[path] = r.section + ' · ' + r.label + ': ' + short(m[0]);
        return true;
      }
      return false;
    }
    function fromText(path, re, convert, where) {
      const m = flat.match(re);
      if (!m) return false;
      setPath(path, convert(m[1], m));
      deal.sources[path] = (where || 'Summary') + ': ' + short(m[0].trim());
      return true;
    }
    function setPath(path, v) { const [a, b] = path.split('.'); deal[a][b] = v; }
    function need(path, found, what) { if (!found) missing.push(what); }
    function assume(path, value, why) {
      setPath(path, value);
      deal.sources[path] = 'ASSUMED: ' + why;
      assumed.push(path);
    }

    const title = md.match(/^#\s+(.+)$/m);
    if (title) deal.meta.name = title[1].split(/\s+[—–-]\s+/)[0].trim();
    const seed = md.match(/Deal #(\d+)/);
    if (seed) deal.meta.from_deal = Number(seed[1]);

    // Timeline
    need('timeline.financial_close', fromText('timeline.financial_close', new RegExp('[Ff]inancial close (?:targeted |on )?' + DATE), isoDate), 'financial close date');
    need('timeline.cod', fromText('timeline.cod', new RegExp('COD (?:on |of )?' + DATE), isoDate), 'COD date');
    need('timeline.ppa_end', take('timeline.ppa_end', /^PPA$/i, new RegExp('ending ' + DATE), isoDate), 'PPA end date (PPA row: "ending …")');
    need('timeline.debt_maturity', take('timeline.debt_maturity', /final maturity/i, new RegExp(DATE), isoDate), 'final maturity');

    // Construction
    need('construction.capex_total', take('construction.capex_total', /^construction capex/i, /^([\d,]+(?:\.\d+)?)/, num), 'construction capex (sources and uses table, AUD k)');
    const prof = flat.match(/[Cc]apex is drawn ((?:\d+(?:\.\d+)?%\s*\/\s*)+\d+(?:\.\d+)?%)/);
    if (prof) {
      deal.construction.capex_profile = prof[1].split('/').map((x) => pct(x));
      deal.sources['construction.capex_profile'] = 'Section 2 · Capex drawn ' + prof[1];
    } else missing.push('capex drawdown profile ("Capex is drawn 15% / 35% / …")');

    // Operations
    need('operations.capacity_mw', take('operations.capacity_mw', /^capacity$/i, /([\d,.]+)\s*MWac/, num), 'capacity (MWac)');
    need('operations.capacity_factor_p50', take('operations.capacity_factor_p50', /capacity factor/i, /([\d.]+)%/, pct), 'P50 capacity factor');
    if (!take('operations.availability', /^availability$/i, /([\d.]+)%/, pct)) assume('operations.availability', 0.99, 'availability 99%, market standard');
    if (!take('operations.degradation', /^degradation$/i, /([\d.]+)%/, pct)) assume('operations.degradation', 0.005, 'degradation 0.5% p.a., market standard');
    need('operations.ppa_price_fc', take('operations.ppa_price_fc', /^PPA$/i, /AUD\s*([\d.]+)\/MWh/, num), 'PPA price (PPA row: "AUD …/MWh")');
    if (!take('operations.ppa_indexation', /^PPA$/i, /indexed at (\d+(?:\.\d+)?)% of CPI/i, pct)) {
      if (!take('operations.ppa_indexation', /^PPA$/i, /(flat price|no indexation)/i, () => 0)) assume('operations.ppa_indexation', 1, 'PPA indexed at 100% of CPI');
    }
    need('operations.opex_real_fc', take('operations.opex_real_fc', /^operating costs/i, /AUD\s*([\d.]+)m p\.a\. real/, (v) => num(v) * 1000), 'operating costs ("AUD …m p.a. real")');
    if (!take('operations.cpi', /^CPI$/i, /([\d.]+)%/, pct)) assume('operations.cpi', 0.025, 'CPI 2.5% p.a.');
    if (!take('operations.p90_factor', /energy yield/i, /P90\s*([\d.]+)%/i, pct)) assume('operations.p90_factor', 0.92, 'one-year P90 at 92% of P50 energy');
    if (!take('operations.p99_factor', /energy yield/i, /P99\s*([\d.]+)%/i, pct)) assume('operations.p99_factor', 0.87, 'one-year P99 at 87% of P50 energy');

    // Tax
    if (!take('tax.rate', /corporate tax/i, /^([\d.]+)%/, pct)) assume('tax.rate', 0.30, 'Australian corporate tax rate 30%');
    if (!take('tax.depreciation_years', /corporate tax/i, /over (\d+) years/, num)) assume('tax.depreciation_years', 20, 'straight-line tax depreciation over 20 years');

    // Senior debt
    need('senior_debt.amount', take('senior_debt.amount', /^senior facility$/i, /AUD\s*([\d,]+)/, (v) => num(v) / 1000), 'Senior Facility amount ("AUD 98,000,000")');
    need('senior_debt.gearing', take('senior_debt.gearing', /maximum gearing/i, /(\d+(?:\.\d+)?)%\s*$/, pct), 'maximum gearing');
    need('senior_debt.base_rate', take('senior_debt.base_rate', /base case bbsw/i, /([\d.]+)%/, pct), 'base case BBSW');
    need('senior_debt.margin', take('senior_debt.margin', /^margin$/i, /([\d.]+)%/, pct), 'margin');
    if (!take('senior_debt.upfront_fee', /upfront fee/i, /([\d.]+)% of the (?:Senior )?Facility/i, pct)) assume('senior_debt.upfront_fee', 0.015, 'upfront fee 1.50%');
    if (!take('senior_debt.commitment_fee_pct_margin', /commitment fee/i, /(\d+(?:\.\d+)?)% of the Margin/i, pct)) assume('senior_debt.commitment_fee_pct_margin', 0.40, 'commitment fee 40% of the margin');
    if (!take('senior_debt.agency_fee', /agency fee/i, /AUD\s*([\d,.]+)k p\.a\./i, num)) assume('senior_debt.agency_fee', 0, 'no agency fee stated');
    if (!take('senior_debt.hedge_ratio', /hedging/i, /hedge (?:at least )?([\d.]+)% of/i, pct)) assume('senior_debt.hedge_ratio', 0, 'no interest rate hedging stated');
    if (!take('senior_debt.hedge_years', /hedging/i, /for (?:at least )?(\d+) years/i, num)) assume('senior_debt.hedge_years', 0, 'no hedging tenor stated');
    if (!take('senior_debt.swap_rate', /hedging/i, /swap rate of ([\d.]+)%/i, pct)) assume('senior_debt.swap_rate', deal.senior_debt.base_rate || 0, 'swap rate equal to base case BBSW');
    if (take('senior_debt.repayment', /^repayment$/i, /(sculpted|annuity|linear|straight-line|equal)/i, (v) => ({ 'straight-line': 'linear', equal: 'annuity' }[v.toLowerCase()] || v.toLowerCase())) === false) {
      deal.sources['senior_debt.repayment'] = 'ASSUMED: sculpted repayment';
    }
    need('senior_debt.target_dscr', take('senior_debt.target_dscr', /target dscr/i, /([\d.]+)x/, num), 'target DSCR');

    // Reserves and covenants
    if (!take('dsra.months_of_ds', /debt service reserve/i, /next (\d+) months/, num)) assume('dsra.months_of_ds', 6, 'DSRA of 6 months of debt service');
    if (!take('covenants.lockup_dscr', /historic dscr/i, /Lock-up\s*<\s*([\d.]+)x/i, num)) assume('covenants.lockup_dscr', 1.10, 'lock-up at 1.10x');
    if (!take('maintenance.cost_real', /maintenance reserve/i, /AUD\s*([\d.]+)m real/i, (v) => num(v) * 1000)) assume('maintenance.cost_real', 0, 'no maintenance reserve stated');
    if (!take('maintenance.from_year', /maintenance reserve/i, /years (\d+) to \d+/i, num)) assume('maintenance.from_year', 10, 'replacement from operating year 10');
    if (!take('maintenance.to_year', /maintenance reserve/i, /years \d+ to (\d+)/i, num)) assume('maintenance.to_year', 12, 'replacement to operating year 12');
    if (!take('maintenance.funding_years', /maintenance reserve/i, /over the (\d+) years before/i, num)) assume('maintenance.funding_years', 3, 'funded over 3 years');
    if (!take('covenants.default_dscr', /historic dscr/i, /Default\s*<\s*([\d.]+)x/i, num)) assume('covenants.default_dscr', 1.05, 'event of default at 1.05x');

    return { ok: missing.length === 0, deal, missing, assumed };
  }

  const api = { parseTermSheet };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
