# Term Sheet → Project Finance Model — SPEC v1

Reference deal: **Project Coolabah** (100 MWac solar, NSW, AUD 98.0m senior debt). The term sheet is in `docs/coolabah_term_sheet.md`.

## 1. Goal

A web app that turns a project finance term sheet into a transparent, auditable cash flow model.

Three layers, kept strictly separate:

1. **Inputs (JSON).** Every term and assumption lives in one `deal.json` file. In v1 the user pastes a term sheet into Claude and gets this JSON back; there is no AI call inside the app yet.
2. **Engine (pure TypeScript functions).** Takes `deal.json` and returns every time series and ratio. No UI code, no randomness, no LLM. Fully unit-tested.
3. **UI + export.** An input editor, the results screens, and an Excel export.

The LLM never calculates anything.

## 2. Scope of v1

**In:**
- Semi-annual timeline from financial close to the end of the PPA
- Construction: capex profile, pro rata debt/equity drawdowns at fixed gearing, IDC, commitment fee, upfront fee (all equity-funded)
- Operations: energy, CPI-indexed revenue and opex, EBITDA
- Corporate tax: straight-line depreciation, losses carried forward
- Senior debt: fixed amount, sculpted repayment
- DSRA: cash, funded by equity at COD, target = next period's scheduled debt service
- Ratios: DSCR per period (pre- and post-tax), LLCR, sculpted debt capacity at the target DSCR
- Lock-up test (flag only)
- Equity IRR
- Sources & uses

**Out (v2+):** hedging, agency fees, MRA, DSRA LC, cash sweep, P90/P99 sizing, margin step-ups, multiple tranches, LLM extraction inside the app.

## 3. Conventions

- Units: AUD thousands (`k`). Rates as decimals (0.0575, not 5.75).
- Periods: semi-annual, ending 31 Dec and 30 Jun. Period 1 = 1 Jul 2027 to 31 Dec 2027.
- Construction: periods 1–4. Operations: period 5 (Jul–Dec 2029) onwards.
- Fiscal year = Australian FY (Jul–Jun). FY30 = Jul 2029 to Jun 2030 = first operating year.
- Interest: `balance_start × annual_rate / 2` (simplified; act/365 in v2).
- Indexation: annual steps from FC. Index for FY*n* = (1 + CPI)^(n − 2028). FY30 → 1.025² = 1.050625.
- Annual operating figures are split 50/50 between the two semesters of a FY.
- Degradation: factor = (1 − deg)^(operating_year − 1).

## 4. Input schema (`deal.json`)

```json
{
  "meta": { "name": "Project Coolabah", "currency": "AUD", "units": "k" },
  "timeline": {
    "financial_close": "2027-07-01",
    "cod": "2029-07-01",
    "ppa_end": "2049-06-30",
    "debt_maturity": "2047-06-30",
    "period_months": 6
  },
  "construction": {
    "capex_total": 140000,
    "capex_profile": [0.15, 0.35, 0.35, 0.15]
  },
  "operations": {
    "capacity_mw": 100,
    "capacity_factor_p50": 0.27,
    "availability": 0.99,
    "degradation": 0.005,
    "ppa_price_fc": 60.0,
    "ppa_indexation": 1.0,
    "opex_real_fc": 3200,
    "cpi": 0.025
  },
  "tax": {
    "rate": 0.30,
    "depreciation_years": 20,
    "depreciable_base": "total_project_cost"
  },
  "senior_debt": {
    "amount": 98000,
    "gearing": 0.70,
    "base_rate": 0.04,
    "margin": 0.0175,
    "upfront_fee": 0.015,
    "commitment_fee_pct_margin": 0.40,
    "repayment": "sculpted",
    "target_dscr": 1.30,
    "cfads_basis": "pre_tax",
    "first_repayment": "2029-12-31"
  },
  "dsra": { "months_of_ds": 6, "funding": "equity_at_cod" },
  "covenants": { "lockup_dscr": 1.10, "default_dscr": 1.05 }
}
```

Every field must also carry, in the UI, a link to the term sheet clause it came from (`source: "§3 Margin"`). Store this in a sibling object `sources` with the same keys so the engine schema stays clean.

## 5. Calculation order

**A. Construction (periods 1–4)**

For each period *t*:
1. `capex[t] = capex_total × profile[t]`
2. `debt_draw[t] = capex[t] × gearing`
3. `interest[t] = debt_balance_start[t] × (base + margin) / 2` (IDC)
4. `commitment_fee[t] = undrawn_start[t] × margin × pct / 2`
5. `upfront_fee` in period 1 = `amount × upfront_fee`
6. `equity[t] = capex[t] × (1 − gearing) + interest[t] + commitment_fee[t] + upfront_fee[t]`
7. Check: total drawn = `amount`; error if not.

**B. Operations (period 5 to end of PPA)**

Annual, then halved per semester:
1. `energy = MW × 8760 × CF × availability × degradation_factor` (MWh)
2. `revenue = energy × ppa_price_fc × index / 1000`
3. `opex = opex_real_fc × index`
4. `EBITDA = revenue − opex`
5. `CFADS_pre_tax = EBITDA` (no working capital, no capex in v1)

**C. Debt sculpting (v1 rule, important)**

Debt is a fixed input (gearing constraint binds). Repayment follows the shape of CFADS:
1. Solve the constant DSCR *d* such that `Σ CFADS[t] / d × DF[t] = amount`, where `DF[t] = (1 + r/2)^-t` from COD and `r = base + margin`.
2. `debt_service[t] = CFADS[t] / d`
3. `interest[t] = balance_start × r / 2`; `principal[t] = debt_service[t] − interest[t]`
4. Check: balance at maturity = 0 (tolerance 0.01k).
5. Also output `debt_capacity_at_target = Σ CFADS[t] / target_dscr × DF[t]` (headroom vs. amount).

**D. Tax**

1. `depreciable_base` = total project cost (capex + IDC + fees).
2. Depreciation per semester = base / years / 2, from COD.
3. `taxable = EBITDA − depreciation − senior interest`
4. Losses carried forward; tax = 30% of positive taxable income after using losses.
5. `CFADS_post_tax = EBITDA − tax`

**E. DSRA**

- Target at end of each period = next period's scheduled debt service (0 after maturity).
- Initial funding at COD by equity.
- Top-up from the waterfall (after debt service), excess released to distributions.
- Released in full with the last repayment.

**F. Waterfall (per operating period)**

CFADS_post_tax → senior interest → senior principal → DSRA top-up / release → distributions.

**G. Ratios and covenants**

- `DSCR[t] = CFADS[t] / debt_service[t]`, pre- and post-tax
- Historic 12-month DSCR (sum of last two periods), tested from the second calculation date after COD
- `LLCR[t] = NPV(CFADS to maturity at r) / balance_start[t]`
- Lock-up flag if 12-month DSCR < 1.10x; default flag if < 1.05x

**H. Equity returns**

Equity cash flows = − construction contributions − DSRA funding + distributions (to end of PPA). IRR on semi-annual flows, annualised as `(1 + irr)² − 1`.

## 6. Reference values (golden test)

These were recomputed independently from the term sheet inputs. The engine must match them within the stated tolerance; if it doesn't, the engine is wrong until proven otherwise.

| Output | Expected | Tolerance |
| --- | --- | --- |
| IDC | 4,226 | ±1 |
| Commitment fees | 858 | ±1 |
| Upfront fee | 1,470 | exact |
| Total project cost | 146,554 | ±2 |
| Total equity at FC (excl. DSRA) | 48,554 | ±2 |
| FY30 net energy (MWh) | 234,155 | ±1 |
| FY30 revenue | 14,760 | ±1 |
| FY30 opex | 3,362 | ±1 |
| FY30 EBITDA | 11,399 | ±1 |
| Number of repayment periods | 36 | exact |
| Constant pre-tax DSCR (sculpted, 98m) | 1.475x | ±0.002 |
| LLCR at COD (pre-tax) | 1.475x | ±0.002 |
| Debt capacity at 1.30x (pre-tax) | 111,206 | ±10 |
| First debt service (Dec 2029) | 3,863 | ±2 |
| First principal (Dec 2029) | 1,046 | ±2 |
| Min post-tax DSCR | 1.25x | ±0.01 |
| First tax-paying period | Jan–Jun 2038 | exact |
| DSRA at COD | 3,863 | ±2 |
| Equity IRR (post-tax, nominal) | 5.15% | ±0.1pt |
| LLCR at COD (post-tax) | 1.417x | ±0.002 |
| PLCR at COD (pre-tax) | 1.584x | ±0.002 |

## 7. Points the term sheet leaves open (decisions taken for v1)

1. **Sculpting vs. fixed debt.** The term sheet says "sculpted to 1.30x" but also fixes debt at 98.0m. At 98.0m the sculpted DSCR is 1.475x, not 1.30x; 1.30x would size to c.111m. v1 sculpts 98.0m to the CFADS profile and reports the 1.30x capacity separately.
2. **Pre-tax vs. post-tax CFADS.** Sculpting follows the term sheet (pre-tax). Post-tax DSCR drops to 1.25x once tax kicks in around 2038, below the 1.30x target. Worth showing clearly in the UI: it's the kind of thing a credit committee will ask about.
3. **DSRA not in sources & uses.** The term sheet funds the DSRA at COD but leaves it out of S&U. v1 adds it as a separate equity contribution at COD and shows it in a second S&U line.
4. **Equity IRR looks low (5.15%).** Expected given a AUD 60/MWh PPA and no terminal value after 2049. Not a bug; a sensitivity to price and terminal value is a v2 feature.

## 8. Screens

1. **Inputs.** Form generated from the schema, grouped by section, with the term sheet source next to each field. Import/export `deal.json`.
2. **Dashboard.** Key outputs: S&U, min/avg DSCR (pre/post-tax), LLCR, debt capacity headroom, equity IRR, lock-up flags.
3. **Charts.** CFADS vs. debt service over time, debt balance, DSCR per period with the 1.30x / 1.10x / 1.05x lines.
4. **Cash flow table.** All periods, all lines, scrollable, frozen first column.
5. **Export.** `.xlsx` with an Inputs tab, a Calc tab with live formulas, and a Checks tab.

## 9. Tech stack

**Status of v1 (built 27 Oct 2026):** steps 1–8 of the build order are done in plain JavaScript, no build tool. `src/engine.js` holds the engine, `test/engine.test.js` the golden tests (`node --test`, 10/10 green), and `index.html` the app, built by `node build.js`, which inlines the engine into the page. Moving to the stack below is optional and can happen in Claude Code.

Target stack:
- Vite + TypeScript, plain React
- Engine in `src/engine/` with no React import
- Tests: Vitest, golden test against section 6
- Charts: Recharts
- Excel export: ExcelJS
- Hosting: GitHub Pages (static, no backend in v1)

## 10. Build order for Claude Code

1. Scaffold project, add `deal.json` for Coolabah and the term sheet in `docs/`.
2. Engine part A (construction) + tests for IDC, fees, S&U. Stop when green.
3. Engine part B (operations) + FY30 tests.
4. Engine part C (sculpting) + DSCR, capacity, schedule tests.
5. Engine parts D–H + remaining golden tests.
6. Dashboard and cash flow table.
7. Charts.
8. Inputs editor with sources.
9. Excel export.
10. Deploy to GitHub Pages.

One step per session, tests green before moving on.
