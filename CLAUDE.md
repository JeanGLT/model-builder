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
- Logo D (sun over a rising curve), also the favicon. Headline: "Drop a term sheet. Get a bankable model."
- Control center = first tab: every input with its clause and status; a drop or a random deal lands there, "Generate model" shows a 0.5 s loading screen then the Overview.
- Periodicity (semi-annual / quarterly) is chosen in the Control center only, not read from the term sheet (`deal.timeline.periodicity`). Dates and the capex profile stay in semesters; a quarter takes half its semester's capex.
- Term sheet template and downloads are Word (.docx). Bottom of the start page: "Download the Excel model" (Coolabah).
- Excel design: dark title band, coloured tabs, shaded construction columns, light-blue section rows, green check cells.

## Workflow
- `node --test` must stay green (34 tests). `node build.js` after any change in `src/` or `deal.json`, then commit `index.html` too.
- Excel formulas: check with `test/verify-excel.html` served over http (see README).
- Add a golden test before each new feature.

## Pitfalls met
- Without solving the PPA price from a target IRR, random deals had a negative median equity IRR.
- HyperFormula (used only for the check page) has no IRR, OFFSET with named arguments failed, and `SUMPRODUCT((range=1)*range)` failed: replaced by SUMIF and plain references. Keep formulas to common functions.
- Opening the export in Excel by AppleScript hung (Excel sandbox file-access prompt). Jean has other workbooks open: never script `active workbook` or close anything but our own file by name.
- `index.html` is generated: edit `src/app.html`, never `index.html`.
- Reading a .docx: Word compresses the files in the zip; `docx.js` uses the browser's DecompressionStream (also in Node 18+). macOS `textutil` flattens tables when it converts, so it is not a fair test of Word.

## Still to do
- Jean to open one exported .xlsx in Excel to confirm it opens without a repair prompt.
- Jean to open a downloaded .docx in Word, edit a number, save and drop it back (tested with generated and re-zipped files, not with real Word).

- Next features: see README "Next steps".
