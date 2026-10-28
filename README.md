# Model Builder — term sheet to project finance model

A web app that turns a project finance term sheet into a full semi-annual cash flow model: construction funding, operations, sculpted debt, tax, DSRA, covenants, equity IRR, sensitivities and breakevens. The whole model downloads as an Excel workbook with live formulas.

- **Reference deal:** Project Coolabah, a 100 MWac solar farm in NSW with AUD 98.0m of senior debt (`docs/coolabah_term_sheet.md`).
- **Random deals:** one click draws a fictional Australian solar deal from realistic ranges, writes its indicative term sheet and builds the model. Each deal has a number (`?deal=4821`), so the same link always opens the same deal.
- **Drop a term sheet:** the start page reads a term sheet in the Model Builder layout (Markdown tables, like the Coolabah one or any generated deal) directly in the browser, and a deal JSON. For a PDF, Word or free-form document, it guides the user through one Claude extraction step.
- **Excel export:** Cover, Inputs (named cells, with the term sheet clause for each), Calc (every cell a formula), Outputs and Checks (integrity checks plus a reconciliation against the web app).

Everything runs in the browser. No backend, no install, no dependencies.

**Live app:** https://jeanglt.github.io/model-builder/ (try `?deal=4821`)

## Files

| File | What it does |
| --- | --- |
| `SPEC.md` | What the model does, calculation order, reference values, decisions. Read this first. |
| `deal.json` | Project Coolabah inputs, with the term sheet clause behind each one. |
| `src/engine.js` | The calculation engine: deal JSON in, time series and ratios out. Pure functions, no UI. |
| `src/generator.js` | Random deal generator (seeded, so a deal number always gives the same deal). |
| `src/termsheet.js` | Writes the indicative term sheet of a generated deal, in Markdown. |
| `src/parser.js` | Reads a Model Builder term sheet back into a deal JSON (no AI: each input is found by its table row). |
| `src/markdown.js` | Small Markdown to HTML converter for the term sheets. |
| `src/xlsx.js` | Minimal .xlsx writer (an .xlsx is a zip of XML files). |
| `src/excel.js` | Rebuilds the model as Excel formulas. |
| `src/app.html` | The interface template. |
| `build.js` | Inlines everything into `index.html` (standalone) and `dist/artifact.html`. |
| `test/*.test.js` | Golden tests, 300 random deals, term sheet reader, Excel export structure. |
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

1. P90 / P99 sizing constraint (§4 of the term sheet) and a P90 case in the dashboard.
2. Interest rate hedging (75% for 10 years), agency fees, MRA, cash sweep after 4 trapped periods.
3. Wind farms in the random generator.
4. Automatic reading of PDF / Word term sheets through a small serverless backend holding a Claude API key (the drop zone already routes these files to a guided step).

Add a golden test for every new feature before building it.
