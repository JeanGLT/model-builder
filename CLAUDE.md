# Model Builder (repo: model-builder, local folder: coolabah-model)

Web app: project finance term sheet → semi-annual or quarterly model → Excel with live formulas. Goal: a portfolio piece for Jean's CV, sent to recruiters as a link. Read `SPEC.md` first.

## Decisions (28 Oct 2026)
- v1 = Coolabah example + random deal mode + drop zone (front page layout A).
- Drop zone (option 3): reads Model Builder term sheets (.docx via `src/docx.js`, .md/.txt) and deal JSON locally with `src/parser.js`; PDF / old .doc / free text get a guided Claude prompt step. Automatic Claude reading (serverless backend + API key) comes later.
- Author line: Jean Grollet · jean@grollet.com.
- Random deals: solar only, Australia. Same seed = same deal (`?deal=4821`).
- Everything in English (site, term sheet, Excel).
- Excel = real model with formulas, not values.
- GitHub: public repo https://github.com/JeanGLT/model-builder, live at https://jeanglt.github.io/model-builder/ (GitHub Pages from main, root). Pushing to main updates the live site: ask Jean before pushing.
- Design: "Credit Approved" style (Inter + DM Serif Display, one blue accent #2f5bea, white cards on #f6f7f9). Source: `/Users/jean/Credit-Approved/DESIGN.md`.
- No dependencies. npm install was not used; the .xlsx writer is hand-written (`src/xlsx.js`).

## Decisions (29 Oct 2026)
- Logo D (sun over a rising curve), also the tab icon: `favicon.png` / `favicon.svg`, linked in <head> by build.js (a data-URI icon inside <body> did not show). Headline: "Drop a term sheet. Get a bankable model."
- Assumptions tab (was "Control center"): every input with its clause and status, the only place inputs are edited (no left rail, no Deal JSON tab since 29 Oct). A deal opens on the Term sheet tab, then Assumptions.
- Every "Generate Excel model" button (top right, Assumptions, Cash flow) shows a 1.5 s loading screen, then downloads the .xlsx.
- Start page kept minimal: no Coolabah button (still reachable with `?deal=coolabah`), no "see the output first" block; three cards: term sheet, assumptions, Excel model.
- Periodicity (semi-annual / quarterly) is chosen in the Assumptions tab only, not read from the term sheet (`deal.timeline.periodicity`). Dates and the capex profile stay in semesters; a quarter takes half its semester's capex.
- Term sheet template = `docs/term_sheet_template.md`, blank with [placeholders], downloaded as Word. Dropped unfilled, the app says how many blanks remain. A test fills it with the Coolabah figures and checks it gives back deal.json.
- Excel design: dark title band, coloured tabs, shaded construction columns, light-blue section rows, green check cells.

## Decisions (29 Oct 2026, v2 model)
- Lender cases: base (P50), one-year P90, one-year P99, Downside (P90, opex +10%, BBSW +100 bp). `deal.scenario`, chosen in Assumptions or by clicking a case on Overview.
- The repayment schedule is fixed at financial close: sized and sculpted on the base case; every other case, sensitivity and breakeven runs with `PFEngine.run(deal, { schedule })`.
- Sizing ("lower_of") = lowest of gearing, P50 at target DSCR, one-year P99 at 1.00x.
- New terms: swaps (`hedge_ratio`, `hedge_years`, `swap_rate`), `agency_fee` (k p.a., deducted before CFADS), `maintenance` {cost_real, from_year, to_year, funding_years}. Missing in a term sheet or JSON: ASSUMED neutral values.
- Golden tests stay on the v1 scope (`v1()` strips the new terms in test/engine.test.js); new terms have their own tests.
- Excel v2 = 15 sheets, same period columns on every calculation sheet, cross-sheet references, live scenario switch.
- Link preview: `og.png` (made from `og.svg` with qlmanage + sips) and Open Graph tags written into <head> by build.js.

## Workflow
- `node --test` must stay green (37 tests). `node build.js` after any change in `src/` or `deal.json`, then commit `index.html` too.
- Excel formulas: check with `test/verify-excel.html` served over http (see README).
- Add a golden test before each new feature.

## Pitfalls met
- Without solving the PPA price from a target IRR, random deals had a negative median equity IRR.
- HyperFormula (used only for the check page) has no IRR, OFFSET with named arguments failed, and `SUMPRODUCT((range=1)*range)` failed: replaced by SUMIF and plain references. Keep formulas to common functions.
- Opening the export in Excel by AppleScript hung (Excel sandbox file-access prompt). Jean has other workbooks open: never script `active workbook` or close anything but our own file by name.
- `index.html` is generated: edit `src/app.html`, never `index.html`.
- The verify page caches scripts: bump `?v=` on its <script> tags after changing src/.
- `sips -c` crops around the centre (cropOffset is ignored): draw og.svg on a square canvas with the content in the middle band.
- Reading a .docx: Word compresses the files in the zip; `docx.js` uses the browser's DecompressionStream (also in Node 18+). macOS `textutil` flattens tables when it converts, so it is not a fair test of Word.

## Still to do
- Done 29 Oct: the v2 .xlsx opens in real Excel without a repair prompt; Excel's own recalculation passes all 14 checks, matches the app on all 18 reconciled lines (IRR included) and the scenario switch gives the app's figures. A .docx edited and saved by real Word reads back (margin change picked up).
- How it was tested: `open -a` then AppleScript on the workbook / document by name, wrapped in `perl -e 'alarm 40; exec @ARGV' osascript ...` so it cannot hang.

- Next features: see README "Next steps".
