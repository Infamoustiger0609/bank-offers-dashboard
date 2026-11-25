# Bank Offers Performance Dashboard

An internal PVR INOX analytics dashboard for tracking how bank card offers and UPI partner cashbacks perform: revenue, transactions, discount cost (bank vs PVR share), ATP/AVT/SPH, and progress against fiscal-year revenue targets. It runs entirely in the browser — Excel files are parsed client-side and nothing is sent to a server.

## Tech stack

- React 19 + Vite
- Tailwind CSS 3
- Recharts (charts)
- SheetJS / `xlsx` (Excel parsing and export)
- Deployed on Vercel

## Getting started

```bash
npm install       # install dependencies
npm run dev       # start the Vite dev server
npm run build     # production build to dist/
npm run preview   # serve the production build locally
```

Add `-- --port <N> --strictPort` to `npm run dev` to choose a port. There is no lint script or test suite; `npm run build` is the verification step.

Deployment happens automatically on Vercel when changes are pushed to the connected GitHub repository. Vercel is case-sensitive, so file names and import paths must match exactly.

## Data

On load the dashboard automatically reads two bundled files from `src/assets/`:

| Source | Default file |
| --- | --- |
| Bank / card offers | `BANK_DATA_APR24-OCT26.xlsx` |
| UPI partners | `UPI_DATA_24_APR-26_AUG.xlsx` |

You can replace either one at any time with the upload buttons in the header (drag-and-drop also works). Uploading a Card file only replaces card data, and uploading a UPI file only replaces UPI data. Uploads are not saved — a page refresh reloads the bundled defaults. To change the bundled defaults, replace the file in `src/assets/` and update the import and file-name constant at the top of `src/App.jsx`.

### Bank / card offers file

First sheet of an `.xlsx`/`.xls`/`.csv`. Column headers must match these names (case-insensitive):

`Card Offers Performance`, `Bank Name`, `No. of Discounted Transactions`, `No. of Free Tickets`, `No. of Tickets`, `Transaction Total (Rs.)`, `Transaction Total Tickets (Rs.)`, `Transaction Total F&B (Rs.)`, `Amount Paid By Customer(Rs.)`, `Discount Amount(Rs.)`, `Discount Contribution Amount Bank (Rs.)`, `Discount Contribution Amount Inox (Rs.)`, `Conv.Fees(Rs.)`, `Date`, `Transaction Type`

Optional cinema-wide ("universal") columns, used for the uplift comparisons and the UNI chart overlay: `Universal Transactions`, `Universal Total Admits`, `Universal Ticket Revenue`, `Universal Total Revenue`.

### UPI partner file

`Month`, `UPI Partner`, `Total No. of Tkts`, `Total Amount (Rs.)`, `Overall Discount Amount (Rs.)`, `PVR Discount Amount (Rs.)`. Optional: `Bank Discount` (otherwise derived from overall minus PVR discount) and `No TRNX` (otherwise ticket count is used).

### Data notes

- Dates must fall between 2015 and 2035, otherwise the row drops out of date-based views.
- Missing columns are treated as zero and reported in a warning banner.
- The fiscal year runs April to March (e.g. `24-25` is Apr 2024 – Mar 2025).

## What's on the dashboard

- **Filters (sticky bar):** Bank / UPI / Both toggle, Fiscal Year, Month, Bank and Offer multi-selects, plus quick periods (This Month, Last Month, QTD, FYTD).
- **Info tiles:** active bank and UPI partner counts, total offers (with a Credit / Debit / Both split), downloadable offer lists per bank and per UPI partner, and fiscal-year targets.
- **Key metrics:** Transactions, Revenue, ATP, AVT, Admits and SPH, each with an optional MoM / QoQ / YoY delta and, where available, an uplift-vs-universal line.
- **Bank / UPI scorecard:** sortable per-partner revenue, transactions and discount, with the Bank vs PVR discount split; click a row for partner details.
- **Charts:** revenue by channel, ticket vs F&B revenue, and the discount contribution split.
- **Offer directory:** every offer by bank/partner with transactions, revenue, discount and discount rate; click for full offer details.
- **Month-wise performance:** monthly trend per bank, month-on-month seasonal comparison across fiscal years, and year-on-year totals, with an optional cinema-wide overlay and click-through drill-downs explaining what changed.
- **Target detail:** achieved vs target per fiscal year, a month-by-month cumulative pro-rata target weighted by last year's seasonality, and the required run rate for the rest of the current year.
- **Comparison Module:** compare two independently filtered groups of banks/UPI partners (Group A vs Group B) on revenue, transactions, discount, admits and discount rate.
- **Discount Budget:** each card bank's actual discount spend against a monthly budget (₹15L per bank per month by default), with a per-bank monthly breakdown. UPI partners are excluded.

Large values are shown in Lakh (L) or Crore (Cr); values below ₹1 L are shown in Thousands (K).

## Project layout

```
src/
  App.jsx       # the whole application (parsing, aggregation, components, page)
  main.jsx      # React entry point
  index.css     # Tailwind directives and base styles
  assets/       # logo and the default bank / UPI Excel files
index.html
tailwind.config.js, postcss.config.js, vite.config.js
```

Everything lives in `src/App.jsx` by design. See `CLAUDE.md` for a detailed architecture guide (filter scopes, offer-name canonicalisation, target logic and so on).
