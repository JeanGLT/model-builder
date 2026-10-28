# Coolabah Model Builder

Turns a project finance term sheet (as a deal JSON) into a full semi-annual cash flow model: construction funding, operations, sculpted debt, tax, DSRA, covenants, equity IRR, sensitivities and breakevens.

## Files

- `SPEC.md`: what the model does, calculation order, reference values, open decisions. Read this first.
- `deal.json`: Project Coolabah inputs, with the term sheet clause behind each one.
- `src/engine.js`: the calculation engine (pure functions, no UI).
- `src/app.html`: the interface template.
- `test/engine.test.js`: golden tests against the term sheet.
- `build.js`: inlines engine + deal into `index.html` (standalone) and `dist/artifact.html`.
- `docs/coolabah_term_sheet.md`: the source term sheet.

## Commands

```
node --test      # run the tests (must stay green)
node build.js    # rebuild index.html after any change
```

Open `index.html` in a browser to use the app. No install needed.

## Next steps (for Claude Code)

1. Excel export with live formulas (ExcelJS): Inputs, Calc and Checks tabs.
2. Interest rate hedging (75% for 10 years), agency fees, MRA, cash sweep after 4 trapped periods.
3. P90 / P99 sizing constraints from §4 of the term sheet.
4. In-app term sheet extraction through a small serverless backend holding the API key.
5. Optional: move to Vite + TypeScript.

Add a golden test for every new feature before building it.
