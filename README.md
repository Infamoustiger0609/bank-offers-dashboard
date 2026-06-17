# Bank Offers Performance Dashboard

A comprehensive analytics dashboard to track and analyze bank offer performance, revenue distribution, discount allocation, and operational insights for PVRINOX.

## 🎯 Overview

The Bank Offers Performance Dashboard provides real-time visibility into:
- **Revenue metrics** - Gross revenue, net revenue, and recovery rates
- **Bank performance** - Revenue, discount costs, and transaction volume by bank
- **Offer analytics** - Individual offer profitability, performance rankings, and date ranges
- **Discount distribution** - Bank vs PVR contribution analysis
- **Monthly trends** - Revenue tracking across selected banks over time

## ✨ Features

### Core Dashboards
- **KPI Cards** - 5 main metrics (Total Transactions, Gross Revenue, Net Revenue, Total Discount, Recovery Rate)
- **Bank Scorecard** - Aggregated performance by bank with scrollable view
- **Discount Split Chart** - Stacked horizontal bar chart showing Bank vs PVR contribution
- **Top Offers Table** - Ranked by revenue with profit calculations and clickable details
- **Offer Duration Tracking** - See how long each offer has been active
- **Monthly Revenue Trend** - Multi-select bank comparison chart

### Advanced Filters
- **Multi-Date Selection** - Filter by one or multiple dates
- **Bank Filter** - View specific bank performance
- **Offer Filter** - Focus on individual offers
- **Group By Toggle** - View all offers or aggregated by bank

### Data Insights
- **Offer Details Modal** - Click any offer to see full breakdown
- **Date Range Visibility** - Track offer active periods
- **Bank Contribution Analysis** - Understand Bank vs PVR cost sharing
- **Profit Calculations** - Revenue minus discount equals profit

## 🚀 Live Dashboard

**Access the dashboard:**
```
https://bank-offers-dashboard.vercel.app
```

## 📊 How to Use

### Step 1: Prepare Your Data
Download the **Bank_Offers_Template.xlsx** file and fill it with your offer data.

**Required columns:**
- Card Offers Performance (offer name)
- No. of Discounted Transactions
- No. of Free Tickets
- No. of Tickets
- Transaction Total (Rs.)
- Transaction Total Tickets (Rs.)
- Transaction Total F&B (Rs.)
- Amount Paid By Customer(Rs.)
- Discount Amount(Rs.)
- Discount Contribution Amount Bank (Rs.)
- Discount Contribution Amount Inox (Rs.)
- Conv.Fees(Rs.)
- Date (format: YYYY-MM-DD)

### Step 2: Upload File
1. Open the dashboard link
2. Click "Upload Excel File" or drag-drop your file
3. Wait for data to load (2-3 seconds)

### Step 3: Explore Data
- **View KPIs** - See top-level metrics
- **Filter by Date/Bank/Offer** - Narrow down analysis
- **Click offers** - See detailed breakdown
- **View trends** - Select multiple banks in monthly chart

## 📁 Data Requirements

### Format
- File type: `.xlsx`, `.xls`, or `.csv`
- Date format: `YYYY-MM-DD` (e.g., 2024-01-15)
- Currency: Indian Rupees (₹)
- Numbers: Enter as digits (not text)

### Column Names (MUST match exactly)
```
Card Offers Performance
No. of Discounted Transactions
No. of Free Tickets
No. of Tickets
Transaction Total (Rs.)
Transaction Total Tickets (Rs.)
Transaction Total F&B (Rs.)
Amount Paid By Customer(Rs.)
Discount Amount(Rs.)
Discount Contribution  Amount Bank (Rs.)
Discount Contribution  Amount Inox (Rs.)
Conv.Fees(Rs.)
Date
```

### Bank Name Recognition
The dashboard automatically extracts bank names from the offer string. 

## 📈 Understanding the Dashboard

### KPI Cards
| Metric | Definition |
|--------|-----------|
| Total Transactions | Count of discounted transactions |
| Gross Revenue | Total transaction amount before discount |
| Net Revenue | Amount actually paid by customers (gross - discount) |
| Total Discount Given | Sum of all discounts offered |
| Recovery Rate | Net Revenue / Gross Revenue (%) |

### Bank Scorecard
Shows per-bank totals across all offers:
- **Total Revenue** - Sum of all offers for that bank
- **Discount Cost** - Bank's total discount contribution
- **Total Transactions** - Volume of discounted transactions
- **Total Offers** - Count of active offers per bank

### Discount Distribution
Stacked bar chart showing:
- **Blue bars** - Bank's discount contribution
- **Green bars** - PVR's discount contribution
- Helps identify which entity bears more cost

### Offer Performance Table
Top 5 offers ranked by revenue:
- **Profit** - Revenue minus discount (key metric)
- **Click to expand** - See full offer details including:
  - Offer type (B1G1, % off, Free Ticket)
  - Bank contribution vs PVR
  - Transaction breakdown

### Total Offers by Bank
Expandable section showing:
- Number of active offers per bank
- Date range (from first to last appearance)
- Identifies long-running vs seasonal offers

### Monthly Trend Chart
- **X-axis** - Date (MM-YYYY format)
- **Y-axis** - Revenue in ₹
- **Multi-select** - Choose 1 or multiple banks
- Independent filters (not affected by main filters)

## 🔧 Technical Stack

- **Frontend Framework** - React 18
- **Styling** - Tailwind CSS
- **Charts** - Recharts
- **Data Processing** - SheetJS (XLSX parsing)
- **Deployment** - Vercel
- **Build Tool** - Vite

## 📦 Installation & Deployment

### Local Development
```bash
# Install dependencies
npm install

# Run development server
npm run dev

# Build for production
npm run build
```

### Deploy on Vercel
1. Push code to GitHub
2. Sign in to https://vercel.com
3. Create "New Project" and select your repo
4. Click "Deploy"
5. Share the live URL with your team

## 👥 Team Usage

### For Operations Head
- Monitor bank partnership health
- Identify top-performing offers
- Track revenue trends
- Make data-driven partnership decisions

### For Analytics Team
- Deep-dive into offer performance
- Analyze discount allocation
- Track seasonal patterns
- Generate performance reports

### For Finance Team
- Monitor discount costs (Bank vs PVR)
- Track recovery rates
- Validate revenue figures
- Analyze profitability by offer

## 🎨 Design & UX

- **Light Theme** - Clean, professional appearance
- **Responsive Layout** - Works on desktop and tablet
- **Color-Coded Insights** - Easy data interpretation
- **Scrollable Tables** - No page overflow
- **Interactive Charts** - Hover for details, click for insights
- **Modal Details** - Comprehensive offer information

## 🔒 Data Security

- **No Backend Required** - All processing happens in your browser
- **No Data Stored** - Files never uploaded to servers
- **Offline Compatible** - Can work without internet (after loading)
- **GDPR Compliant** - Customer data stays private

## ⚠️ Known Limitations

- Works best with data from last 24 months
- Charts handle up to 20 banks clearly
- Single sheet per file (combine multiple sheets before upload)
- Date format must be consistent (YYYY-MM-DD)

## 🐛 Troubleshooting

### "Unknown Bank" appears
- Bank name not recognized in offer string
- Check offer naming convention matches expected format

### Numbers show as "NaN"
- Column name mismatch
- Use exact column names from template

### Chart not displaying
- Empty data after filtering
- Check date format in Excel (YYYY-MM-DD)
- Ensure numeric columns have numbers, not text

### File upload fails
- File format not .xlsx/.xls/.csv
- Corrupted file - re-save and try again
- Large file (>10MB) - split into smaller chunks

## 📞 Support

For issues or feature requests:
1. Check data format matches template
2. Review column names (case-sensitive)
3. Verify date format (YYYY-MM-DD)
4. Clear browser cache and retry

## 📝 Version History

**v1.0** (Current)
- Multi-date selection filter
- Discount split chart (Bank vs PVR)
- Offer duration tracking
- Monthly trend analysis
- Full offer details modal
- Top offers ranking
- Bank scorecard


**Last Updated:** June 20246 
**Dashboard URL:** https://bank-offers-dashboard.vercel.app  
**Template Download:** https://drive.google.com/drive/folders/1aVGXDWyQRN4mlwL3PLSTZUSkd34LJeU_
