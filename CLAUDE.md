# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install       # install dependencies
npm run dev        # start Vite dev server (add: -- --port <N> --strictPort to pick a specific port)
npm run build       # production build to dist/
npm run preview      # serve the production build locally
```

There is no lint script and no test suite configured in this repo (no ESLint config, no `*.test.*`/`*.spec.*` files, no `npm test`). Verification is build-only: run `npm run build` after changes, and optionally boot `npm run dev` and hit it with curl/a browser to smoke-test.

Deployment is Vercel, triggered by pushing to the connected GitHub repo (see `.vercelignore`); there's no CI config in-repo.

## Architecture

This is a single-page, client-only dashboard — there is no backend. All Excel parsing, aggregation, and state live in the browser.

### Almost everything is in one file

`src/App.jsx` (~3800 lines) contains the entire application: constants/column maps, pure parsing/formatting/aggregation functions, several small presentational components (modals, tooltips, `StatCard`, `MultiSelectDropdown`), and one large `export default function App()` that owns all state and renders the whole page. `src/main.jsx` just mounts `<App />`. Before making a change, `grep`/search within `App.jsx` rather than assuming logic is split across files — it almost never is.

### Two merged data sources, one row shape

Rows come from two independent Excel uploads that get merged into a single `rows` array in state, distinguished by `row.paymentCategory` (`"Card"` | `"UPI"`):

- **Bank/Card offers file** — columns defined in `COLUMN_MAP`, parsed by `parseWorkbookRows`.
- **UPI partner file** — columns defined in `UPI_COLUMN_MAP`, parsed by `parseUpiWorkbookRows`/`buildUpiRow`.

Uploading one file only replaces rows of that category (`setRows(current => [...current.filter(r => r.paymentCategory === otherCategory), ...parsedRows])`), so uploading a new Card file doesn't wipe previously-uploaded UPI data and vice versa.

Every parsed row carries a common set of derived fields regardless of source: `date`, `dateLabel`, `monthKey` (`MM-YYYY`), `fiscalYear` (e.g. `"24-25"`, FY starts April — see `getFiscalYearLabel`/`calendarYearForFiscalMonth`/`FISCAL_MONTH_ORDER`), plus `bankName`, `offerName`, `paymentCategory`, and the numeric fields (`transactionTotal`, `discountAmount`, `bankContribution`/`inoxContribution`, `totalTickets`, `discountedTransactions`, etc.).

### The row-filtering pipeline (read this before adding a new chart/table)

Filters are: `fyFilter` (fiscal years) + `monthFilter` (abbreviated month names, e.g. `"Apr"`) — this pair replaced the old single calendar date-range filter — plus `bankFilter`, `offerFilter` (canonical offer names, not raw), and `paymentCategoryFilter` (`"all" | "card" | "upi"`, shown as Bank/UPI/Both in the UI).

Rather than one `filteredRows`, `App()` maintains several differently-scoped memos because different sections of the dashboard need to ignore different filters:

- `categoryScopedRows` — only `paymentCategoryFilter` applied. Drives the Bank/Offer filter dropdown option lists.
- `filteredRows` — the full filter set (FY+Month+Bank+Offer+Category). The "current selection" used by most KPI cards and tables.
- `bankOfferFilteredRows` — Bank+Offer+Category only, **no date**. Used by the MoM/QoQ/YoY comparison logic (`comparisonKpis`), which computes its own date windows around `latestDataDate`.
- `rankScopeRows` — FY+Month+Category only, **no Bank/Offer**. Used for the global bank revenue ranking (`globalBankRankMap`) so ranks don't shift just because a bank got deselected in the table filter.
- `categoryDateOfferScopedRows` — FY+Month+Offer+Category, **no Bank**. Used for the trend/seasonal/inference charts and panels so selecting fewer banks doesn't collapse the chart's own data.
- `cardRows`/`upiRows` — `filteredRows` split by `paymentCategory`.

When adding a new metric, match it to the correct existing scope rather than writing a fresh ad-hoc filter — picking the wrong one is the most common source of "this chart doesn't match that KPI" bugs in this codebase.

### Offer name canonicalization

Raw offer names vary by channel suffix, card-type prefix, and spacing (e.g. `"HDFC Credit Card - 10% off - Online"` vs `"HDFC Debit Card - 10% off"`). `normalizeOfferChannel` + `OFFER_ALIAS_MAP` collapse these into one canonical string via `canonicalOfferName(offerName)`. `offerFilter` stores canonical names, so every offer-matching comparison in the row-filtering pipeline must call `canonicalOfferName(row.offerName)` — comparing against raw `row.offerName` will silently under-match.

### Aggregation functions are pure

Functions like `computeKpis`, `aggregateBanks`, `aggregateOffers`, `aggregateMonthlySeries`, `aggregateSeasonalByYear`, `aggregateYearlyTotals`, `aggregateChannelRevenue`, `filterGroupFiscal`, `aggregateGroupBankBreakdown` all take a `rows` array (plus sometimes a bank list) and return a derived shape — no closures over component state. Call them with whichever scoped-rows memo above is appropriate; don't reimplement their filtering inline.

### "Universal" (cinema-wide) comparison metrics

Some KPIs (ATP, AVT, Admits) compare bank-side numbers against cinema-wide "universal" totals pulled from optional columns (`universalTransactions`, `admits`, `universalTicketRevenue`, `universalTotalRevenue` — see `OPTIONAL_COLUMNS`) via `getMonthlyReferenceValue`. These columns are optional; the dashboard must keep working when they're absent (`universalATP`/`universalAVT` etc. fall back to `null`, and `StatCard`s render a fallback subtitle instead of the "uplift" line — see `computeUpliftOrContribution`/`UpliftOrContributionLine`).

### Modals

All overlays (`OfferModal`, `BankModal`, `OffersByBankModal` — reused for both Bank and UPI partner breakdowns, `GroupDetailModal` for the Comparison Module) are plain components conditionally rendered at the bottom of `App()`'s JSX based on state, not a routing/portal system.

## Data format (required for uploads to parse correctly)

Column names in the uploaded Excel/CSV must match `COLUMN_MAP` / `UPI_COLUMN_MAP` exactly (case-sensitive) — a mismatch causes that column to be treated as zero/missing rather than erroring loudly, which shows up downstream as `NaN` or an "Unknown Bank"/"Unknown Offer" row. Dates must parse to a year between 2015–2035 (see `parseExcelDate`) or the row's `monthKey`/`fiscalYear` become `"Unknown"` and it drops out of most date-scoped views.
