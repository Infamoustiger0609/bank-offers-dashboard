# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install       # install dependencies
npm run dev        # start Vite dev server (add: -- --port <N> --strictPort to pick a specific port)
npm run build       # production build to dist/
npm run preview      # serve the production build locally
```

There is no lint script and no test suite configured in this repo (no ESLint config, no `*.test.*`/`*.spec.*` files, no `npm test`). Verification is build-only: run `npm run build` after changes, and optionally boot `npm run dev` and hit it with curl/a browser to smoke-test. (The build prints a >500 kB chunk warning — that is expected, everything is one bundle.)

Deployment is Vercel, triggered by pushing to the connected GitHub repo (see `.vercelignore`); there's no CI config in-repo. Vercel's filesystem is case-sensitive, so **import paths must match filenames on disk exactly** — this repo has broken on Vercel before from a casing mismatch (see the `Assets`/`assets` fixes in git history).

Stack: React 19, Vite, Tailwind CSS 3 (custom colour tokens in `tailwind.config.js`: `appBg`, `borderSoft`, `textMain`, `textMuted`, `accentBlue`, `accentGreen`, …; fonts Manrope + Space Grotesk via `src/index.css`), Recharts 3, SheetJS (`xlsx`). `vite.config.js` adds `assetsInclude: ["**/*.xlsx"]` so Excel files can be imported as URLs.

Repo quirk: `node_modules/` is currently tracked in git despite being in `.gitignore` (it was committed before the ignore rule). Don't be surprised by it in `git status`/diffs; it should eventually be removed from the index with `git rm -r --cached node_modules`.

## Architecture

This is a single-page, client-only dashboard — there is no backend. All Excel parsing, aggregation, and state live in the browser.

### Almost everything is in one file

`src/App.jsx` (~4,750 lines) contains the entire application: constants/column maps, pure parsing/formatting/aggregation functions, several small presentational components (modals, tooltips, `StatCard`, `MultiSelectDropdown`, `MetricComparisonBox`), and one large `export default function App()` that owns all state and renders the whole page. `src/main.jsx` just mounts `<App />` in `React.StrictMode`; `src/index.css` holds the Tailwind directives, body background gradient and `.scrollbar-thin`. The body is `overflow: hidden` — the page scrolls inside the `App()` root `div`, not the window. Before making a change, `grep`/search within `App.jsx` rather than assuming logic is split across files — it almost never is.

`pvrinox_dashboard.html` in the repo root is just a meta-refresh stub pointing at `dist/index.html`; it is not part of the app.

### Two merged data sources, one row shape

Rows come from two independent Excel uploads that get merged into a single `rows` array in state, distinguished by `row.paymentCategory` (`"Card"` | `"UPI"`):

- **Bank/Card offers file** — columns defined in `COLUMN_MAP`, parsed by `parseWorkbookRows` (via `parseCardArrayBuffer`).
- **UPI partner file** — columns defined in `UPI_COLUMN_MAP`, parsed by `parseUpiWorkbookRows`/`buildUpiRow` (via `parseUpiArrayBuffer`). UPI partner names are normalised by `normalizeUpiPartnerName` (e.g. `"moikwik"`/`"Mobikwik UPI Cashback"` → `MobiKwik`). UPI rows only carry a month, so `resolveUpiMonthDate` infers the year from the previous row when the cell isn't a real date. If the UPI file has no "No TRNX" column, `discountedTransactions` falls back to `totalTickets`; UPI `ticketRevenue` = `transactionTotal`, `fnbRevenue` = 0, and every UPI row is `transactionType: "Online"` with a synthetic offer name `"<Partner> UPI Cashback"`.

Uploading one file only replaces rows of that category (`setRows(current => [...current.filter(r => r.paymentCategory === otherCategory), ...parsedRows])`), so uploading a new Card file doesn't wipe previously-uploaded UPI data and vice versa. The "✕ Remove" buttons likewise remove only their own category.

**Default files auto-load.** On every fresh mount a `useEffect` fetches the two bundled files (`defaultBankDataUrl` → `src/assets/BANK_DATA_APR24-OCT26.xlsx`, `defaultUpiDataUrl` → `src/assets/UPI_DATA_24_APR-26_AUG.xlsx`; display names in `DEFAULT_BANK_FILE_NAME`/`DEFAULT_UPI_FILE_NAME`). A manual upload sets `manualCardUploadRef`/`manualUpiUploadRef`, so a manual upload that starts mid-fetch always wins. Nothing is persisted (no localStorage) — a refresh re-fetches the defaults. **To swap a default file:** drop the new file into `src/assets/`, delete the old one, and update *both* the `import` at the top of `App.jsx` and the matching `DEFAULT_*_FILE_NAME` constant, matching case exactly.

Every parsed row carries a common set of derived fields regardless of source: `date`, `dateLabel`, `monthKey` (`MM-YYYY`), `fiscalYear` (e.g. `"24-25"`, FY starts April — see `getFiscalYearLabel`/`calendarYearForFiscalMonth`/`priorFiscalYearLabel`/`FISCAL_MONTH_ORDER`), plus `bankName`, `offerName`, `paymentCategory`, `offerType` (`inferOfferType`), `transactionType` (Online / Offline - Box Office / Offline - F&B), and the numeric fields (`transactionTotal`, `discountAmount`, `bankContribution`/`inoxContribution`, `amountPaid`, `ticketRevenue`, `fnbRevenue`, `totalTickets`, `discountedTransactions`, `freeTickets`, `convFees`, …).

### The row-filtering pipeline (read this before adding a new chart/table)

Filters are: `fyFilter` (fiscal years) + `monthFilter` (abbreviated month names, e.g. `"Apr"`) — together they replaced the old calendar date-range filter — plus `bankFilter`, `offerFilter` (canonical offer names, not raw), and `paymentCategoryFilter` (`"all" | "card" | "upi"`, shown as Both/Bank/UPI in the UI). Filter state is auto-reset to "everything selected" by `useEffect`s whenever the option lists change (`fiscalYears`, `fiscalMonths`, `banks`, `offers`) — including after any upload or category toggle. The Month dropdown's options are scoped to the selected FYs (`fiscalMonthsForSelectedFy`), and selected months are pruned when FYs change. Quick-period buttons (This Month / Last Month / QTD / FYTD, `applyQuickPeriod`) just set `fyFilter` + `monthFilter` relative to `latestDataDate`; `activeQuickPeriodKey` highlights whichever preset matches.

Rather than one `filteredRows`, `App()` maintains several differently-scoped memos because different sections of the dashboard need to ignore different filters:

- `categoryScopedRows` — only `paymentCategoryFilter` applied. Drives the Bank/Offer filter dropdown option lists (`banks`, `offers`; `offers` is additionally narrowed by `bankFilter`).
- `filteredRows` — the full filter set (FY+Month+Bank+Offer+Category). The "current selection" used by most KPI cards, the scorecard/offer tables, pies, discount split, and modals.
- `bankOfferFilteredRows` — Bank+Offer+Category only, **no date**. Used by the MoM/QoQ/YoY comparison logic (`comparisonKpis`), which computes its own date windows around `latestDataDate` via `getComparisonPeriods`.
- `rankScopeRows` — FY+Month+Category only, **no Bank/Offer**. Used for the global bank revenue ranking (`globalBankRankMap`) so ranks don't shift just because a bank got deselected in the table filter.
- `categoryDateOfferScopedRows` — FY+Month+Offer+Category, **no Bank**. Used for the trend/seasonal/inference charts and panels (and the `universal*ByMonth` maps, `bankActiveRangeMap`) so selecting fewer banks doesn't collapse the chart's own data.
- `cardRows`/`upiRows` — `filteredRows` split by `paymentCategory`. `cardRows` also feeds the offers-by-bank list, the Credit/Debit/Both breakdown and the Discount Budget breakdown.

When adding a new metric, match it to the correct existing scope rather than writing a fresh ad-hoc filter — picking the wrong one is the most common source of "this chart doesn't match that KPI" bugs in this codebase.

Not filtered at all (always the full `rows`): `fiscalYearAchievement`/the Target card, the `admitsData`/`universal*Data` reference maps in `App()`, and the Comparison Module groups (which have their own filters).

### Offer name canonicalization

Raw offer names vary by channel suffix, card-type prefix, and spacing (e.g. `"HDFC Credit Card - 10% off - Online"` vs `"HDFC Debit Card - 10% off"`). `normalizeOfferChannel` (strips trailing channel suffixes like `- Online`/`- Box Office`/`- Digital Platforms`, strips everything up to the last `… Card -` prefix, normalises `%` and `:` spacing; leaves names containing "F&B" untouched) + `OFFER_ALIAS_MAP` (hand-maintained raw→canonical overrides; add new pairs there) collapse these into one canonical string via `canonicalOfferName(offerName)`. `offerFilter` stores canonical names, so every offer-matching comparison in the row-filtering pipeline must call `canonicalOfferName(row.offerName)` — comparing against raw `row.offerName` will silently under-match. (A few places — `offersByBank`, `upiOffersByPartner`, `cardOfferTypeBreakdown`, `aggregateSeasonalByYear`'s offer count — inline `OFFER_ALIAS_MAP[x] || normalizeOfferChannel(x)`, which is equivalent.)

### Aggregation functions are pure

Functions like `computeKpis`, `computeRowTotals`, `aggregateBanks`, `aggregateOffers`, `aggregateMonthlySeries`, `aggregateSeasonalByYear`, `aggregateYearlyTotals`, `aggregateChannelRevenue`, `filterGroupFiscal`, `aggregateGroupBankBreakdown`, `buildDiscountBudgetBreakdown`, `buildBankYearOverYear`, `buildAdjacentMonthInsight`/`buildAggregateAdjacentMonthInsight` all take a `rows` array (plus sometimes a bank list/admits map) and return a derived shape — no closures over component state. Call them with whichever scoped-rows memo above is appropriate; don't reimplement their filtering inline. The exception is `buildPairInsight`, a closure defined inside `App()` (used by the MoM inference panels and the YoY year insight).

### "Universal" (cinema-wide) comparison metrics

Some KPIs (ATP, AVT, SPH, Admits) compare bank-side numbers against cinema-wide "universal" totals pulled from optional columns (`universalTransactions`, `admits` = "Universal Total Admits", `universalTicketRevenue`, `universalTotalRevenue` — see `OPTIONAL_COLUMNS`) via `getMonthlyReferenceValue`, which reads one value per `monthKey` (last row wins) and flags months whose rows disagree (`inconsistent`, surfaced as an amber warning bar). These columns are optional; the dashboard must keep working when they're absent (`universalATP`/`universalAVT`/`universalSPH` etc. fall back to `null`, and `StatCard`s render a fallback subtitle instead of the "uplift" line — see `computeUpliftOrContribution`/`UpliftOrContributionLine`). `universalSPH` has no dedicated column of its own — it's derived as Universal Total Revenue minus Universal Ticket Revenue (a stand-in for universal F&B revenue), divided by universal admits, over the same "valid months only" pattern as `universalATP`/`universalAVT`.

### SPH and the Ticket/F&B split

`sph` (Spend Per Head = F&B revenue ÷ admits) and `ticketFnbSplit` are computed from `atpAvtKpis` — the same Bank/UPI/Both-aware source rows used by ATP and AVT (`atpAvtSourceRows`, which branches on `paymentCategoryFilter`). `EMPTY_KPIS`/`computeKpis` carry a `fnbRevenue` field for this. Note UPI rows have zero F&B revenue, so SPH on UPI/Both is diluted by design. "Admits" at bank level is `totalTickets` (a proxy — true admits are cinema-wide).

### Universal background lines on trend charts (separate from the KPI uplift lines above)

Independently of the StatCard-level "uplift vs universal" lines, the Month-wise Bank/UPI Performance section (modes: Monthly / Month on Month / Year on Year, `trendMode`) has its own "UNI" toggle (`showUniversal`) that overlays a dashed cinema-wide reference line on the seasonal (Month on Month) and Year on Year charts. Both charts add a second, hidden Recharts axis (`yAxisId="universal"`) so the overlay's scale never distorts the bank-side axis. On the seasonal chart specifically, the universal line is rendered **once per fiscal year** in `seasonalYears` (matching that year's own line color), not a single merged line — summing multiple fiscal years' universal figures together would silently inflate the number, since each year needs its own independent `calendarYearForFiscalMonth` lookup. `attachUniversalToSeasonalPoints`/`attachUniversalToYearlyPoints` compute this as **separate derived arrays** (`seasonalDataWithUniversal`/`yearlyDataWithUniversal`) rather than mutating `seasonalData`/`yearlyData` in place — `seasonalYears` itself derives its year-key set generically from `seasonalData`'s own object keys, so writing a `universalValue` key directly onto those objects would make it show up as a bogus extra "year" series. The overlay is only available for the revenue / transactions / admits metrics (`pickUniversalMap`).

Clicking a chart point (or "View Details ▸", `handleViewDetails`) opens an inline drill-down panel: monthly mode → bank-vs-prior-month + YoY + rank; MoM mode → `MetricComparisonBox`es plus per-year-pair "who drove it" insight (`inferencePairs`); YoY mode → `yearInsight`. Headline/notes text comes from the `compute*Note`/`generateHeadlineInsight` helpers.

### Fiscal-year revenue targets (the "Target" KPI card)

`FISCAL_YEAR_TARGETS` is a hardcoded `{ "24-25": null, "25-26": ₹100 Cr, "26-27": ₹150 Cr }` map (rupee amounts; `null` means no target set for that FY) — update it by hand when targets change, there's no upload path for this. `fiscalYearAchievement` computes achieved-vs-target per FY from the **full unfiltered `rows`** (both Card + UPI, ignoring every UI filter) since this is a company-wide target, not a filtered view. The Target card in the info-tile row shows all FYs compactly when `fyFilter` isn't meaningfully narrowed, or one larger figure per FY when it's narrowed to a genuine subset.

`TargetDetailModal`'s month-by-month table is the one place in the codebase with non-trivial forecasting logic: each month's pro-rata target is **seasonally weighted**, not a flat straight-line ramp. For FY with target `target_FY`, it looks up the immediately preceding fiscal year (`priorFiscalYearLabel`) and that year's actual monthly revenue (`buildMonthlyRevenueByFY`, filtered by `row.fiscalYear`, not by date range), then `monthlyProRata[month] = priorYearActual[month] × (target_FY / priorYearTotalActual)` — i.e. it scales last year's own seasonal shape (heavy months stay heavy) rather than assuming revenue accrues evenly. **The displayed "Pro-rata Target" column is the running cumulative sum of those monthly figures** (compared against the cumulative revenue column, so "vs Target" is cumulative ahead/behind). Falls back to a flat `target_FY / 12` per month when the prior FY has no data at all (e.g. the first year in the dataset). Months after `latestDataDate` are flagged `isFuture` and render as "—" instead of a red/green ahead-behind tag, since they haven't happened yet. The "Required Run Rate" block for the *current* FY (the FY of `latestDataDate`) applies the same seasonal-weighting idea to the remaining shortfall, splitting it across remaining months in proportion to the prior year's actual revenue for those same months (falling back to an even split only if the prior year has no data for them).

### Discount Budget (header button → `DiscountBudgetModal` / `DiscountBudgetBankModal`)

Compares each **Card bank's** actual discount spend against a flat monthly budget — UPI partners have no budgets and are excluded (it's built from `cardRows`). `DEFAULT_BANK_MONTHLY_DISCOUNT_BUDGET` (₹15L per bank per month) applies to every bank unless overridden in `BANK_MONTHLY_DISCOUNT_BUDGET_OVERRIDES` (resolved by `getBankMonthlyBudget`). `buildDiscountBudgetBreakdown(cardRows)` groups discount by bank and month; a bank's budget = (number of months in which it had data within the current filter) × monthly budget, and `variance = actual − budget`. Note the sign convention in the UI: variance ≥ 0 is labelled "▲ ahead" in green and negative as "▼ behind" in red (i.e. positive = spent *more* than budget). Banks are sorted by variance ascending. Clicking a bank row opens `DiscountBudgetBankModal` with the month-by-month table. It follows the main filters (FY/Month/Bank/Offer) through `cardRows`.

### Comparison Module (Group A vs Group B)

This used to be an inline page section called "Custom Comparison"; it's now a modal (`showComparisonModal`, opened via the "Comparison Module" button in the header) so it doesn't compete for space with the rest of the dashboard. Each of the two groups is independently filtered by its own bank list plus its own FY + Month selection — `filterGroupFiscal(rows, banks, fiscalYearsSelected, monthsSelected)` uses the same FY/Month model as the main filter bar, not a calendar date range, and ignores the main filters entirely (it runs on the full `rows`). It compares Revenue, Transactions, Discount Amount, Admits and Discount Rate with a B-vs-A delta badge (`renderDeltaBadge`). Each group panel has a clickable "BANK: X / UPI: Y" pill that opens `GroupDetailModal` with a per-bank/partner breakdown (`aggregateGroupBankBreakdown`); that breakdown's "admits" column uses `totalTickets` as a proxy, the same convention used elsewhere in the dashboard since true admits/footfall is cinema-wide, not bank-specific.

### Offer card-type classification (Credit / Debit / Both)

`classifyCardType(rawOfferName)` looks for "credit"/"debit" keywords in the **raw** offer name — checked before any channel-suffix/card-type-prefix stripping happens in `normalizeOfferChannel` — and returns `"credit" | "debit" | "both" | "unspecified"`. `cardOfferTypeBreakdown` then buckets each unique bank+canonical-offer pair (collecting the set of types seen across its raw names) into Credit, Debit, or Both, shown in the Total Offers info tile. Offers whose raw name mentions neither keyword (`"unspecified"`) are folded into **Both** (assumed to apply to both card types) at that call site, not a separate "Other" bucket. The "BMS: 96 / District: 30 (Jul)" text on the Total Offers tile is **hardcoded** — update it by hand.

### Number formatting

All display formatting is in small pure helpers near the top of `App.jsx`. `formatInLakh` (₹, "L") and `formatCountInLakh` (no ₹, "L") show Lakh values (0 decimals at ≥1 L, 2 decimals below that) — **but values with absolute magnitude below 1,00,000 automatically switch to Thousands with one decimal** (`₹45.2 K` / `45.2 K`), so small banks/offers no longer render as "₹0.00 L". `formatInCr` is used for targets/achievements (1 decimal, `NA` for null), `formatInteger` uses `en-IN` grouping, `formatRupee` is 2-decimal rupees (ATP/AVT/SPH), and `formatCompactNumber`/`formatLegendValue` are separate Cr/L/K helpers for legends. Change `formatInLakh`/`formatCountInLakh` together so money and count displays stay consistent.

### Modals

All overlays (`OfferModal`, `BankModal`, `OffersByBankModal` — reused for both Bank and UPI partner breakdowns, `GroupDetailModal` for the Comparison Module, `TargetDetailModal` for the Target card, `DiscountBudgetModal`/`DiscountBudgetBankModal`, and the inline Comparison Module modal) are plain components conditionally rendered at the bottom of `App()`'s JSX based on state, not a routing/portal system. `anyModalOpen` (used to hide the sticky filter bar, which is `z-50`) is currently `Boolean(selectedBank || selectedOffer || showOffersByBank || showOffersByUpi || showTargetDetail)` — it does **not** include `showComparisonModal`, `groupDetailOpen`, `showDiscountBudget` or `discountBudgetBankOpen`, so the filter bar can float over those modals' backdrops (the Comparison and Discount Budget modals are also `z-50`). `anyModalOpen` must be updated whenever a new full-page modal's open-state is added — this has been missed more than once and is currently known-incomplete.

Inside the MoM/YoY drill-down panels, `MetricComparisonBox` optionally takes a `bankInfo` prop (`{currentBanks, priorBanks, currentLabel, priorLabel}`) that adds a "{N} banks" pill revealing which banks were new/dropped/common between the two periods being compared (simple `Set` diffing) — omit the prop and the component renders exactly as a plain metric box.

### Excel export

`exportOffersToExcel(offersByEntity, filename)` (via the `xlsx` package, already used for parsing uploads) lets users download the Bank Partners / UPI Partners offer breakdowns (`bank_offers.xlsx` / `upi_partner_offers.xlsx`) — one row per offer per bank/partner, with the offer's active date range and Bank/UPI vs PVR contribution split — as an `.xlsx` file, via small "⬇ Download Excel" buttons on those info tiles. Both follow the current filters (built from `cardRows`/`upiRows`).

### Branding / favicon

The header uses `src/assets/pvr-inox-logo-transparent.png`. `index.html` currently has **no** `<link rel="icon">` (browser default); a custom bank-logo favicon under `public/` is planned but not yet added — if you add one, put it in `public/` and reference it from `index.html`.

## Data format (required for uploads to parse correctly)

Column names must match `COLUMN_MAP` / `UPI_COLUMN_MAP` (header matching is case-insensitive and whitespace-normalised via `normalizeKey`, but spelling and punctuation must match) — a missing column is treated as zero/missing rather than erroring loudly (missing required columns are listed in an amber banner / UPI error line), which shows up downstream as `NaN`-free zeros or an "Unknown Bank"/"Unknown Offer" row.

- **Card file** (`COLUMN_MAP`): `Card Offers Performance`, `Bank Name`, `No. of Discounted Transactions`, `No. of Free Tickets`, `No. of Tickets`, `Transaction Total (Rs.)`, `Transaction Total Tickets (Rs.)`, `Transaction Total F&B (Rs.)`, `Amount Paid By Customer(Rs.)`, `Discount Amount(Rs.)`, `Discount Contribution Amount Bank (Rs.)`, `Discount Contribution Amount Inox (Rs.)`, `Conv.Fees(Rs.)`, `Date`, `Transaction Type`; optional universal columns `Universal Transactions`, `Universal Total Admits`, `Universal Ticket Revenue`, `Universal Total Revenue`.
- **UPI file** (`UPI_COLUMN_MAP`): `Month`, `UPI Partner`, `Total No. of Tkts`, `Total Amount (Rs.)`, `Overall Discount Amount (Rs.)`, `PVR Discount Amount (Rs.)`; optional `Bank Discount` (else derived as discount − PVR share) and `No TRNX` (else falls back to ticket count).

Only the first sheet of each workbook is read. Dates must parse to a year between 2015–2035 (see `parseExcelDate`) or the row's `monthKey`/`fiscalYear` become `"Unknown"` and it drops out of most date-scoped views.
