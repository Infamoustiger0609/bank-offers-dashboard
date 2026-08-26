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

`src/App.jsx` (~4400 lines) contains the entire application: constants/column maps, pure parsing/formatting/aggregation functions, several small presentational components (modals, tooltips, `StatCard`, `MultiSelectDropdown`), and one large `export default function App()` that owns all state and renders the whole page. `src/main.jsx` just mounts `<App />`. Before making a change, `grep`/search within `App.jsx` rather than assuming logic is split across files — it almost never is.

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

Some KPIs (ATP, AVT, SPH, Admits) compare bank-side numbers against cinema-wide "universal" totals pulled from optional columns (`universalTransactions`, `admits`, `universalTicketRevenue`, `universalTotalRevenue` — see `OPTIONAL_COLUMNS`) via `getMonthlyReferenceValue`. These columns are optional; the dashboard must keep working when they're absent (`universalATP`/`universalAVT`/`universalSPH` etc. fall back to `null`, and `StatCard`s render a fallback subtitle instead of the "uplift" line — see `computeUpliftOrContribution`/`UpliftOrContributionLine`). `universalSPH` has no dedicated column of its own — it's derived as Universal Total Revenue minus Universal Ticket Revenue (a stand-in for universal F&B revenue), divided by universal admits, over the same "valid months only" pattern as `universalATP`/`universalAVT`.

### SPH and the Ticket/F&B split

`sph` (Spend Per Head = F&B revenue ÷ admits) and `ticketFnbSplit` are computed from `atpAvtKpis` — the same Bank/UPI/Both-aware source rows used by ATP and AVT (`atpAvtSourceRows`, which branches on `paymentCategoryFilter`). `EMPTY_KPIS`/`computeKpis` carry a `fnbRevenue` field for this.

### Universal background lines on trend charts (separate from the KPI uplift lines above)

Independently of the StatCard-level "uplift vs universal" lines, the Month-wise Bank/UPI Performance section has its own "UNI" toggle (`showUniversal`) that overlays a dashed cinema-wide reference line on the seasonal (Month on Month) and Year on Year charts. Both charts add a second, hidden Recharts axis (`yAxisId="universal"`) so the overlay's scale never distorts the bank-side axis. On the seasonal chart specifically, the universal line is rendered **once per fiscal year** in `seasonalYears` (matching that year's own line color), not a single merged line — summing multiple fiscal years' universal figures together would silently inflate the number, since each year needs its own independent `calendarYearForFiscalMonth` lookup. `attachUniversalToSeasonalPoints`/`attachUniversalToYearlyPoints` compute this as **separate derived arrays** (`seasonalDataWithUniversal`/`yearlyDataWithUniversal`) rather than mutating `seasonalData`/`yearlyData` in place — `seasonalYears` itself derives its year-key set generically from `seasonalData`'s own object keys, so writing a `universalValue` key directly onto those objects would make it show up as a bogus extra "year" series.

### Fiscal-year revenue targets (the "Target" KPI card)

`FISCAL_YEAR_TARGETS` is a hardcoded `{ "24-25": null, "25-26": ..., "26-27": ... }` map (rupee amounts; `null` means no target set for that FY) — update it by hand when targets change, there's no upload path for this. `fiscalYearAchievement` computes achieved-vs-target per FY from the **full unfiltered `rows`** (both Card + UPI, ignoring every UI filter) since this is a company-wide target, not a filtered view. The Target card in the KPI ribbon shows all FYs compactly when `fyFilter` isn't meaningfully narrowed, or one large figure when it's narrowed to a genuine subset.

`TargetDetailModal`'s month-by-month table is the one place in the codebase with non-trivial forecasting logic: each month's pro-rata target is **seasonally weighted**, not a flat straight-line ramp. For FY with target `target_FY`, it looks up the immediately preceding fiscal year (`priorFiscalYearLabel`) and that year's actual monthly revenue (`buildMonthlyRevenueByFY`, filtered by `row.fiscalYear`, not by date range), then `proRataTarget[month] = priorYearActual[month] × (target_FY / priorYearTotalActual)` — i.e. it scales last year's own seasonal shape (heavy months stay heavy) rather than assuming revenue accrues evenly. Falls back to a flat `target_FY / 12` when the prior FY has no data at all (e.g. the first year in the dataset). Months after `latestDataDate` are flagged `isFuture` and render as "—" instead of a red/green ahead-behind tag, since they haven't happened yet. The "Required Run Rate" block for the *current* FY applies the same seasonal-weighting idea to the remaining shortfall, splitting it across remaining months in proportion to the prior year's actual revenue for those same months (falling back to an even split only if the prior year has no data for them).

### Comparison Module (Group A vs Group B)

This used to be an inline page section called "Custom Comparison"; it's now a modal (`showComparisonModal`, opened via the "Comparison Module" button in the header) so it doesn't compete for space with the rest of the dashboard. Each of the two groups is independently filtered by its own bank list plus its own FY + Month selection — `filterGroupFiscal(rows, banks, fiscalYearsSelected, monthsSelected)` uses the same FY/Month model as the main filter bar, not a calendar date range. Each group panel has a clickable "BANK: X / UPI: Y" pill that opens `GroupDetailModal` with a per-bank/partner breakdown (`aggregateGroupBankBreakdown`); that breakdown's "admits" column uses `totalTickets` as a proxy, the same convention used elsewhere in the dashboard since true admits/footfall is cinema-wide, not bank-specific.

### Offer card-type classification (Credit / Debit / Both)

`classifyCardType(rawOfferName)` looks for "credit"/"debit" keywords in the **raw** offer name — checked before any channel-suffix/card-type-prefix stripping happens in `normalizeOfferChannel` — to bucket each unique bank+canonical-offer pair into Credit, Debit, or Both, shown in the Total Offers KPI card. Offers whose raw name mentions neither keyword are folded into **Both** (assumed to apply to both card types), not a separate "Other"/"unspecified" bucket.

### Modals

All overlays (`OfferModal`, `BankModal`, `OffersByBankModal` — reused for both Bank and UPI partner breakdowns, `GroupDetailModal` for the Comparison Module, `TargetDetailModal` for the Target card) are plain components conditionally rendered at the bottom of `App()`'s JSX based on state, not a routing/portal system. `anyModalOpen` (used to hide the sticky filter bar) must be updated whenever a new full-page modal's open-state is added, or the filter bar will visibly float above the modal's backdrop — this has been missed more than once.

Inside the MoM/YoY drill-down panels, `MetricComparisonBox` optionally takes a `bankInfo` prop (`{currentBanks, priorBanks, currentLabel, priorLabel}`) that adds a "{N} banks" pill revealing which banks were new/dropped/common between the two periods being compared (simple `Set` diffing) — omit the prop and the component renders exactly as a plain metric box.

### Excel export

`exportOffersToExcel(offersByEntity, filename)` (via the `xlsx` package, already used for parsing uploads) lets users download the Bank Partners / UPI Partners offer breakdowns — one row per offer per bank/partner, with the offer's active date range and Bank/UPI vs PVR contribution split — as an `.xlsx` file, via small "⬇ Download Excel" buttons on those KPI ribbon cards.

## Data format (required for uploads to parse correctly)

Column names in the uploaded Excel/CSV must match `COLUMN_MAP` / `UPI_COLUMN_MAP` exactly (case-sensitive) — a mismatch causes that column to be treated as zero/missing rather than erroring loudly, which shows up downstream as `NaN` or an "Unknown Bank"/"Unknown Offer" row. Dates must parse to a year between 2015–2035 (see `parseExcelDate`) or the row's `monthKey`/`fiscalYear` become `"Unknown"` and it drops out of most date-scoped views.
