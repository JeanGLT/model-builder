# Model Builder (coolabah-model)

Web app: project finance term sheet → semi-annual model → Excel with live formulas. Goal: a portfolio piece for Jean's CV, sent to recruiters as a link. Read `SPEC.md` first.

## Decisions (28 Oct 2026)
- v1 = Coolabah example + random deal mode. No in-app term sheet extraction yet (would need a backend holding an API key).
- Random deals: solar only, Australia. Same seed = same deal (`?deal=4821`).
- Everything in English (site, term sheet, Excel).
- Excel = real model with formulas, not values.
- GitHub: public repo https://github.com/JeanGLT/coolabah-model, live at https://jeanglt.github.io/coolabah-model/ (GitHub Pages from main, root). Pushing to main updates the live site: ask Jean before pushing.
- Design: "Credit Approved" style (Inter + DM Serif Display, one blue accent #2f5bea, white cards on #f6f7f9). Source: `/Users/jean/Credit-Approved/DESIGN.md`.
- No dependencies. npm install was not used; the .xlsx writer is hand-written (`src/xlsx.js`).

## Workflow
- `node --test` must stay green (22 tests). `node build.js` after any change in `src/` or `deal.json`, then commit `index.html` too.
- Excel formulas: check with `test/verify-excel.html` served over http (see README).
- Add a golden test before each new feature.

## Pitfalls met
- Without solving the PPA price from a target IRR, random deals had a negative median equity IRR.
- HyperFormula (used only for the check page) has no IRR, OFFSET with named arguments failed, and `SUMPRODUCT((range=1)*range)` failed: replaced by SUMIF and plain references. Keep formulas to common functions.
- Opening the export in Excel by AppleScript hung (Excel sandbox file-access prompt). Jean has other workbooks open: never script `active workbook` or close anything but our own file by name.
- `index.html` is generated: edit `src/app.html`, never `index.html`.

## Still to do
- Jean to open one exported .xlsx in Excel to confirm it opens without a repair prompt.
- Next features: see README "Next steps".
