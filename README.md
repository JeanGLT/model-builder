# Model Builder — term sheet to project finance model

A web app that turns a project finance term sheet into a full semi-annual or quarterly cash flow model: construction funding, operations, sculpted debt, tax, DSRA, covenants, equity IRR, sensitivities and breakevens. The whole model downloads as an Excel workbook with live formulas.

- **Reference deal:** Project Coolabah, a 100 MWac solar farm in NSW with AUD 98.0m of senior debt (`docs/coolabah_term_sheet.md`).
- **Random deals:** one click draws a fictional Australian solar deal from realistic ranges, writes its indicative term sheet and builds the model. Each deal has a number (`?deal=4821`), so the same link always opens the same deal.
- **Drop a term sheet:** the start page reads a Word (.docx) or Markdown term sheet in the Model Builder layout (like the blank downloadable template, once filled in, or any generated deal) directly in the browser, and a deal JSON. For a PDF or a free-form document, it guides the user through one Claude extraction step.
- **Assumptions tab:** after a drop or a random deal, the term sheet opens first, then the Assumptions tab lists every input with its term sheet clause and a status (from the term sheet, assumed, edited). It is the only place inputs are edited. The user picks semi-annual or quarterly periods, then clicks Generate Excel model.
- **Lender cases:** base case (P50), one-year P90, one-year P99 and a Downside (P90, opex +10%, BBSW +100 bp). The debt is sized on the lowest of gearing, DSCR on P50 and 1.00x on P99, sculpted on the base case, and every other case, sensitivity and breakeven runs against that fixed repayment schedule.
- **Lender terms:** interest rate swaps on part of the debt, agency fees, a maintenance reserve (MRA) for the inverter replacement.
- **Excel export (15 sheets):** Cover, Dashboard, Inputs (named cells, with the term sheet clause for each), Scenarios (live switch), Timing, Operations, Construction, Debt, Tax, Reserves, Waterfall, Ratios, Returns, Annual summary and Checks (integrity checks plus a reconciliation against the web app).

Everything runs in the browser. No backend, no install, no dependencies.

**Live app:** https://jeanglt.github.io/model-builder/ (try `?deal=4821`)

## Files

| File | What it does |
| --- | --- |
| `SPEC.md` | What the model does, calculation order, reference values, decisions. Read this first. |
| `docs/term_sheet_template.md` | Blank term sheet (placeholders in square brackets), downloaded as Word from the start page. |
| `deal.json` | Project Coolabah inputs, with the term sheet clause behind each one. |
| `src/engine.js` | The calculation engine: deal JSON in, time series and ratios out. Pure functions, no UI. |
| `src/generator.js` | Random deal generator (seeded, so a deal number always gives the same deal). |
| `src/termsheet.js` | Writes the indicative term sheet of a generated deal, in Markdown. |
| `src/parser.js` | Reads a Model Builder term sheet back into a deal JSON (no AI: each input is found by its table row). |
| `src/markdown.js` | Small Markdown to HTML converter for the term sheets. |
| `src/xlsx.js` | Minimal .xlsx writer (an .xlsx is a zip of XML files). |
| `src/docx.js` | Writes the term sheet as a Word file and reads a Word term sheet back (no library). |
| `src/excel.js` | Rebuilds the model as a 15-sheet Excel workbook of live formulas. |
| `src/app.html` | The interface template. |
| `build.js` | Inlines everything into `index.html` (standalone) and `dist/artifact.html`. |
| `test/*.test.js` | Golden tests, 300 random deals, quarterly periods, term sheet reader, Word round trip, Excel export structure. |
| `test/verify-excel.html` | Recalculates the exported workbook with an independent spreadsheet engine. |

## Commands

```
node --test      # run all the tests (must stay green)
node build.js    # rebuild index.html after any change in src/ or deal.json
```

Open `index.html` in a browser to use the app.

To check the Excel formulas independently, serve the folder and open the check page:

```
python3 -m http.server 8765
```

then open `http://localhost:8765/test/verify-excel.html`. Every row of the reconciliation must say `true` (the IRR line shows `#NAME` there only because that engine has no IRR function; the page checks the IRR separately).

## Publishing on GitHub Pages

1. Make the repository public (GitHub Pages is free only for public repositories).
2. Settings → Pages → Deploy from branch → `main` / root.
3. The app is then at `https://<user>.github.io/<repo>/`. Links like `.../?deal=4821` open a given deal.

## Next steps

1. Cash sweep after 4 trapped periods.
2. Wind farms in the random generator.
3. Automatic reading of PDF / Word term sheets through a small serverless backend holding a Claude API key (the drop zone already routes these files to a guided step).

Add a golden test for every new feature before building it.
