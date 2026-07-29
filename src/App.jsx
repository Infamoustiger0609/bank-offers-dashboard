import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import pvrInoxLogo from "./Assets/pvr-inox-logo-transparent.png";
import {
  CartesianGrid,
  Bar,
  BarChart,
  Cell,
  LabelList,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

const COLUMN_MAP = {
  offerName: "Card Offers Performance",
  bankName: "Bank Name",
  discountedTransactions: "No. of Discounted Transactions",
  freeTickets: "No. of Free Tickets",
  totalTickets: "No. of Tickets",
  transactionTotal: "Transaction Total (Rs.)",
  ticketRevenue: "Transaction Total Tickets (Rs.)",
  fnbRevenue: "Transaction Total F&B (Rs.)",
  amountPaid: "Amount Paid By Customer(Rs.)",
  discountAmount: "Discount Amount(Rs.)",
  bankContribution: "Discount Contribution Amount Bank (Rs.)",
  inoxContribution: "Discount Contribution Amount Inox (Rs.)",
  convFees: "Conv.Fees(Rs.)",
  date: "Date",
  transactionType: "Transaction Type",
  universalTransactions: "Universal Transactions",
  admits: "Universal Total Admits",
  universalTicketRevenue: "Universal Ticket Revenue",
  universalTotalRevenue: "Universal Total Revenue",
};

const OPTIONAL_COLUMNS = new Set(["universalTransactions", "admits", "universalTicketRevenue", "universalTotalRevenue"]);

const UPI_COLUMN_MAP = {
  month: "Month",
  bankName: "UPI Partner",
  totalTickets: "Total No. of Tkts",
  transactionTotal: "Total Amount (Rs.)",
  discountAmount: "Overall Discount Amount (Rs.)",
  inoxContribution: "PVR Discount Amount (Rs.)",
  bankContribution: "Bank Discount",
  discountedTransactions: "No TRNX",
};

const UPI_OPTIONAL_COLUMNS = new Set(["bankContribution", "discountedTransactions"]);

const CHANNEL_COLORS = {
  Online: "#2563eb",
  "Offline BO": "#f59e0b",
  "Offline - F&B": "#10b981",
};
const CHANNEL_FALLBACK_COLOR = "#94a3b8";

const TICKET_FNB_COLORS = { Tickets: "#2563eb", "F&B": "#f59e0b" };

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const BANK_COLORS = [
  "#2563eb",
  "#10b981",
  "#f59e0b",
  "#f97316",
  "#8b5cf6",
  "#ef4444",
  "#0ea5e9",
  "#14b8a6",
  "#e11d48",
  "#84cc16",
  "#3b82f6",
  "#06b6d4",
  "#22c55e",
  "#f43f5e",
  "#a855f7",
  "#fb7185",
  "#d97706",
  "#4f46e5",
];

const KPI_COLORS = [
  { accent: "bg-blue-50 border-blue-100", dot: "bg-accentBlue" },
  { accent: "bg-emerald-50 border-emerald-100", dot: "bg-accentGreen" },
  { accent: "bg-amber-50 border-amber-100", dot: "bg-accentAmber" },
  { accent: "bg-orange-50 border-orange-100", dot: "bg-accentOrange" },
  { accent: "bg-violet-50 border-violet-100", dot: "bg-accentPurple" },
  { accent: "bg-teal-50 border-teal-100", dot: "bg-teal-500" },
  { accent: "bg-pink-50 border-pink-100", dot: "bg-pink-500" },
  { accent: "bg-white/90 border-white/60", dot: "bg-indigo-500" },
];

const EMPTY_KPIS = {
  totalTransactions: 0,
  grossRevenue: 0,
  netRevenue: 0,
  totalDiscount: 0,
  recoveryRate: 0,
  ticketRevenue: 0,
  totalTickets: 0,
};

function normalizeKey(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value).replace(/,/g, "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseExcelDate(value) {
  if (!value && value !== 0) return null;
  const finalize = (date) => {
    if (!date || Number.isNaN(date.getTime())) return null;
    const year = date.getFullYear();
    if (year < 2015 || year > 2035) return null;
    return date;
  };
  if (value instanceof Date) return finalize(value);
  if (typeof value === "number") {
    return finalize(new Date(Math.round((value - 25569) * 86400 * 1000)));
  }
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return finalize(parsed);
  const parts = raw.split(/[/-]/).map((part) => Number(part.trim()));
  if (parts.length === 3) {
    const [day, month, year] = parts;
    return finalize(new Date(year < 100 ? 2000 + year : year, (month || 1) - 1, day || 1));
  }
  return null;
}

function resolveUpiMonthDate(rawValue, previousDate) {
  const direct = parseExcelDate(rawValue);
  if (direct) return direct;
  const raw = String(rawValue || "").trim();
  if (!raw || !previousDate) return null;
  const probe = new Date(`${raw} 1, 2000`);
  if (Number.isNaN(probe.getTime())) return null;
  const monthIndex = probe.getMonth();
  let year = previousDate.getFullYear();
  if (monthIndex <= previousDate.getMonth()) year += 1;
  return new Date(year, monthIndex, 1);
}

function normalizeUpiPartnerName(rawName) {
  const name = String(rawName || "").trim();
  const key = name.toLowerCase();
  if (key === "cred") return "CRED";
  if (key === "moikwik" || key === "mobikwik" || key === "mobikwik upi cashback") return "MobiKwik";
  if (key === "paytm" || key === "paytm upi cashback") return "Paytm";
  if (key === "phonepe") return "PhonePe";
  if (key === "airtel cashback upi") return "Airtel";
  return name;
}

function formatCompactNumber(value) {
  const amount = Number(value || 0);
  if (Math.abs(amount) >= 10000000) return `${(amount / 10000000).toFixed(2)} Cr`;
  if (Math.abs(amount) >= 100000) return `${(amount / 100000).toFixed(2)} L`;
  if (Math.abs(amount) >= 1000) return `${(amount / 1000).toFixed(1)}K`;
  return `${Math.round(amount)}`;
}

function formatInLakh(value) {
  const lakhValue = Number(value || 0) / 100000;
  const decimals = Math.abs(lakhValue) >= 1 ? 0 : 2;
  return `₹${lakhValue.toFixed(decimals)} L`;
}

function formatCountInLakh(value) {
  const lakhValue = Number(value || 0) / 100000;
  const decimals = Math.abs(lakhValue) >= 1 ? 0 : 2;
  return `${lakhValue.toFixed(decimals)} L`;
}

function formatInteger(value) {
  return new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(value || 0);
}

function formatDateLabel(date) {
  if (!date) return "Unknown Date";
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric" }).format(date);
}

function formatIsoDate(date) {
  if (!date) return "Unknown";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function exportOffersToExcel(offersByEntity, filename) {
  const rows = [];
  offersByEntity.forEach((entity) => {
    entity.offers.forEach((offer) => {
      rows.push({
        "Bank/Partner": entity.bankName,
        "Offer Name": offer.offerName,
        "Start Date": offer.startLabel,
        "End Date": offer.endLabel,
        "Bank/UPI Contribution (Rs.)": Math.round(offer.bankDiscount * 100) / 100,
        "PVR Contribution (Rs.)": Math.round(offer.pvrDiscount * 100) / 100,
      });
    });
  });
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Offers");
  XLSX.writeFile(workbook, filename);
}

function monthKeyFromDate(date) {
  if (!date) return "Unknown";
  const year = date.getFullYear();
  if (year < 2015 || year > 2035) return "Unknown";
  return `${String(date.getMonth() + 1).padStart(2, "0")}-${year}`;
}

function getFiscalQuarterStart(date) {
  const month = date.getMonth();
  const year = date.getFullYear();
  if (month >= 3 && month <= 5) return new Date(year, 3, 1);
  if (month >= 6 && month <= 8) return new Date(year, 6, 1);
  if (month >= 9 && month <= 11) return new Date(year, 9, 1);
  return new Date(year, 0, 1);
}

function getFiscalYearLabel(date) {
  if (!date) return "Unknown";
  const year = date.getFullYear();
  const month = date.getMonth(); // April = 3
  const fyStartYear = month >= 3 ? year : year - 1;
  return `${String(fyStartYear).slice(-2)}-${String(fyStartYear + 1).slice(-2)}`;
}

function calendarYearForFiscalMonth(fiscalYearLabel, calendarMonthNumber) {
  // calendarMonthNumber is 1-12 (Jan=1 ... Dec=12), matching monthKey's month part
  const fyStartYear = 2000 + Number(fiscalYearLabel.split("-")[0]);
  return calendarMonthNumber >= 4 ? fyStartYear : fyStartYear + 1;
}

const FISCAL_MONTH_ORDER = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"];

const QUICK_PERIODS = [
  { key: "thisMonth", label: "This Month" },
  { key: "lastMonth", label: "Last Month" },
  { key: "qtd", label: "QTD" },
  { key: "fytd", label: "FYTD" },
];

const COMPARISON_MODES = [
  { key: "none", label: "None" },
  { key: "mom", label: "MoM" },
  { key: "qoq", label: "QoQ" },
  { key: "yoy", label: "YoY" },
];

const COMPARISON_LABELS = { mom: "MoM", qoq: "QoQ", yoy: "YoY" };

const TREND_MODES = [
  { key: "monthly", label: "Monthly" },
  { key: "mom", label: "Month on Month" },
  { key: "yoy", label: "Year on Year" },
];

const SEASONAL_METRICS = [
  { key: "revenue", label: "Revenue", field: "transactionTotal" },
  { key: "transactions", label: "Transactions", field: "discountedTransactions" },
  { key: "discount", label: "Discount", field: "discountAmount" },
  { key: "admits", label: "Admits", field: "totalTickets" },
];

function monthRange(year, month) {
  return { start: new Date(year, month, 1), end: new Date(year, month + 1, 0, 23, 59, 59, 999) };
}

function getComparisonPeriods(mode, latestDataDate) {
  if (!latestDataDate) return null;
  const year = latestDataDate.getFullYear();
  const month = latestDataDate.getMonth();

  if (mode === "mom") {
    const current = monthRange(year, month);
    const prior = monthRange(year, month - 1);
    return { currentStart: current.start, currentEnd: current.end, priorStart: prior.start, priorEnd: prior.end };
  }

  if (mode === "qoq") {
    const currentStart = getFiscalQuarterStart(latestDataDate);
    const currentEnd = new Date(currentStart.getFullYear(), currentStart.getMonth() + 3, 0, 23, 59, 59, 999);
    const priorStart = new Date(currentStart.getFullYear(), currentStart.getMonth() - 3, 1);
    const priorEnd = new Date(priorStart.getFullYear(), priorStart.getMonth() + 3, 0, 23, 59, 59, 999);
    return { currentStart, currentEnd, priorStart, priorEnd };
  }

  if (mode === "yoy") {
    const current = monthRange(year, month);
    const prior = monthRange(year - 1, month);
    return { currentStart: current.start, currentEnd: current.end, priorStart: prior.start, priorEnd: prior.end };
  }

  return null;
}

function computeDelta(currentValue, priorValue) {
  if (!priorValue) return null;
  return ((currentValue - priorValue) / priorValue) * 100;
}

function formatRupee(value) {
  return `₹${Number(value || 0).toFixed(2)}`;
}

function computeRowTotals(list) {
  const raw = list.reduce(
    (acc, row) => {
      acc.revenue += row.transactionTotal;
      acc.discount += row.discountAmount;
      acc.transactions += row.discountedTransactions;
      acc.ticketRevenue += row.ticketRevenue;
      return acc;
    },
    { revenue: 0, discount: 0, transactions: 0, ticketRevenue: 0 },
  );
  return {
    revenue: Number(raw.revenue.toFixed(5)),
    discount: Number(raw.discount.toFixed(5)),
    transactions: Number(raw.transactions.toFixed(5)),
    ticketRevenue: Number(raw.ticketRevenue.toFixed(5)),
  };
}

function computeDiscountRateNote(currentRate, priorRate) {
  if (currentRate === null || priorRate === null) return "Not enough data to compare discount rate.";
  const rateDelta = currentRate - priorRate;
  if (rateDelta > 2) return "Discount rate rose — growth came with a higher discount cost.";
  if (rateDelta < -2) return "Efficiency uplift — discount rate improved, meaning offers are converting more cost-effectively.";
  return "Discount rate held steady.";
}

function computeVolumeValueNote(revenueDeltaPercent, txnDeltaPercent) {
  if (revenueDeltaPercent === null || txnDeltaPercent === null) return "Not enough data to compare volume and value.";
  const gap = revenueDeltaPercent - txnDeltaPercent;
  if (gap > 15) return "Strong value uplift — revenue grew faster than transaction count, driven by higher-value transactions.";
  if (gap < -15) return "Transaction volume grew faster than revenue — lower average value per transaction.";
  return "Revenue and transaction volume moved together.";
}

function computeAdmitsNote(admitsDeltaPercent) {
  if (admitsDeltaPercent === null) return "Not enough data to compare admits.";
  if (Math.abs(admitsDeltaPercent) < 3) return "Admits held roughly steady.";
  if (admitsDeltaPercent > 0) return `Healthy uplift in admits — up ${admitsDeltaPercent.toFixed(1)}% versus the prior period.`;
  return `Admits fell ${Math.abs(admitsDeltaPercent).toFixed(1)}% versus the prior period.`;
}

function computeATVNote(atvDeltaPercent) {
  if (atvDeltaPercent === null) return "Not enough data to compare ATV.";
  if (Math.abs(atvDeltaPercent) < 3) return "ATV held roughly steady.";
  if (atvDeltaPercent > 0) return `ATV uplift of ${atvDeltaPercent.toFixed(1)}% — customers spending more per visit.`;
  return `ATV fell ${Math.abs(atvDeltaPercent).toFixed(1)}%.`;
}

function generateHeadlineInsight({ revenueDeltaPercent, admitsDeltaPercent, atvDeltaPercent }) {
  if (revenueDeltaPercent === null) return null;
  const revUp = revenueDeltaPercent > 0;
  const admitsKnown = admitsDeltaPercent !== null;
  const atvKnown = atvDeltaPercent !== null;
  const admitsMovedOpposite = admitsKnown && Math.sign(admitsDeltaPercent) !== Math.sign(revenueDeltaPercent) && Math.abs(admitsDeltaPercent) > 3;
  const admitsMovedSimilarly =
    admitsKnown && Math.sign(admitsDeltaPercent) === Math.sign(revenueDeltaPercent) && Math.abs(admitsDeltaPercent - revenueDeltaPercent) < 15;

  if (admitsKnown && admitsMovedOpposite) {
    return `Revenue ${revUp ? `saw a strong uplift of ${revenueDeltaPercent.toFixed(1)}%` : `fell ${Math.abs(revenueDeltaPercent).toFixed(1)}%`} while admits moved the other way (${
      admitsDeltaPercent > 0 ? "+" : ""
    }${admitsDeltaPercent.toFixed(1)}%) — this performance looks specific to this bank/offer, not a footfall effect.`;
  }
  if (admitsKnown && admitsMovedSimilarly) {
    return `Revenue and admits moved together (${revenueDeltaPercent.toFixed(1)}% vs ${admitsDeltaPercent.toFixed(
      1,
    )}%)${revUp ? " — a broad uplift across footfall and spend" : " — this looks like a broader footfall trend rather than something specific to this bank/offer"}.`;
  }
  if (atvKnown && Math.abs(atvDeltaPercent) > 10 && admitsKnown && Math.abs(admitsDeltaPercent) < 5) {
    return `Admits were roughly flat, but average ticket value ${atvDeltaPercent > 0 ? `rose ${Math.abs(atvDeltaPercent).toFixed(1)}% — a strong per-visitor spend uplift` : `fell ${Math.abs(atvDeltaPercent).toFixed(1)}%`}${
      revUp ? "" : ""
    }.`;
  }
  return `Revenue ${revUp ? `saw an uplift of ${Math.abs(revenueDeltaPercent).toFixed(1)}%` : `fell ${Math.abs(revenueDeltaPercent).toFixed(1)}%`}. Admits/ATV data is incomplete for this period, so it isn't clear whether footfall or spend-per-visitor drove the change.`;
}

const PIE_LABEL_RADIAN = Math.PI / 180;
function renderPieLabel(props) {
  const { cx, cy, midAngle, outerRadius, percent, name, value } = props;
  const radius = outerRadius + 26;
  const x = cx + radius * Math.cos(-midAngle * PIE_LABEL_RADIAN);
  const y = cy + radius * Math.sin(-midAngle * PIE_LABEL_RADIAN);
  return (
    <text x={x} y={y} fill="#1a202c" fontSize={12} fontWeight={600} textAnchor={x > cx ? "start" : "end"} dominantBaseline="central">
      {`${name}: ${formatInLakh(value)} (${(percent * 100).toFixed(0)}%)`}
    </text>
  );
}

function renderClickableDot(dotProps, dataKey, onClick) {
  const value = dotProps.payload[dataKey];
  if (value === undefined || value === null || Number.isNaN(value)) return null;
  return (
    <g key={`dot-${dataKey}-${dotProps.index}`} style={{ cursor: "pointer" }} onClick={onClick}>
      <circle cx={dotProps.cx} cy={dotProps.cy} r={12} fill="transparent" />
      <circle cx={dotProps.cx} cy={dotProps.cy} r={6} fill={dotProps.stroke} stroke="white" strokeWidth={2} />
      <circle cx={dotProps.cx} cy={dotProps.cy} r={9} fill="none" stroke={dotProps.stroke} strokeWidth={1} opacity={0.35} />
    </g>
  );
}

function renderDeltaBadge(value) {
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-bold ${
        value === null
          ? "bg-slate-100 text-slate-500"
          : value > 0
            ? "bg-emerald-50 text-emerald-600"
            : value < 0
              ? "bg-rose-50 text-rose-600"
              : "bg-slate-100 text-slate-500"
      }`}
    >
      {value === null ? "—" : `${value > 0 ? "▲" : value < 0 ? "▼" : "–"} ${Math.abs(value).toFixed(1)}%`}
    </span>
  );
}

function buildBankYearOverYear(bankName, monthKey, allRows) {
  const [month, currentYear] = monthKey.split("-").map(Number);
  const priorYear = currentYear - 1;
  const revenueFor = (year) =>
    allRows
      .filter((r) => r.bankName === bankName && r.monthKey !== "Unknown" && Number(r.monthKey.split("-")[0]) === month && Number(r.monthKey.split("-")[1]) === year)
      .reduce((sum, r) => sum + r.transactionTotal, 0);

  const priorRevenue = revenueFor(priorYear);
  const currentRevenue = revenueFor(currentYear);
  if (!priorRevenue) return null;

  return {
    priorYear,
    currentYear,
    priorRevenue,
    currentRevenue,
    deltaPercent: ((currentRevenue - priorRevenue) / priorRevenue) * 100,
  };
}

function buildAggregateAdjacentMonthInsight(banks, monthKey, allRows, admitsMap) {
  const scopedRows = allRows.filter((row) => row.monthKey !== "Unknown" && (!banks.length || banks.includes(row.bankName)));
  const monthKeysWithData = [...new Set(scopedRows.map((row) => row.monthKey))].sort((a, b) => {
    const [am, ay] = a.split("-").map(Number);
    const [bm, by] = b.split("-").map(Number);
    return new Date(ay, am - 1, 1) - new Date(by, bm - 1, 1);
  });
  const idx = monthKeysWithData.indexOf(monthKey);
  const priorMonthKey = idx > 0 ? monthKeysWithData[idx - 1] : null;
  if (!priorMonthKey) return null;

  const currentRows = scopedRows.filter((r) => r.monthKey === monthKey);
  const priorRows = scopedRows.filter((r) => r.monthKey === priorMonthKey);
  const currentTotals = computeRowTotals(currentRows);
  const priorTotals = computeRowTotals(priorRows);
  if (!priorTotals.revenue) return null;

  const txnDeltaPercent = priorTotals.transactions
    ? ((currentTotals.transactions - priorTotals.transactions) / priorTotals.transactions) * 100
    : null;
  const currentAdmits = admitsMap ? admitsMap.get(monthKey) || 0 : 0;
  const priorAdmits = admitsMap ? admitsMap.get(priorMonthKey) || 0 : 0;
  const admitsDeltaPercent = priorAdmits ? ((currentAdmits - priorAdmits) / priorAdmits) * 100 : null;

  return {
    priorMonthKey,
    currentRevenue: currentTotals.revenue,
    priorRevenue: priorTotals.revenue,
    deltaPercent: ((currentTotals.revenue - priorTotals.revenue) / priorTotals.revenue) * 100,
    currentTransactions: currentTotals.transactions,
    priorTransactions: priorTotals.transactions,
    txnDeltaPercent,
    currentAdmits,
    priorAdmits,
    admitsDeltaPercent,
  };
}

function buildAdjacentMonthInsight(bankName, monthKey, allRows, admitsMap) {
  const bankRows = allRows.filter((row) => row.bankName === bankName && row.monthKey !== "Unknown");
  const monthKeysWithData = [...new Set(bankRows.map((row) => row.monthKey))].sort((a, b) => {
    const [am, ay] = a.split("-").map(Number);
    const [bm, by] = b.split("-").map(Number);
    return new Date(ay, am - 1, 1) - new Date(by, bm - 1, 1);
  });
  const idx = monthKeysWithData.indexOf(monthKey);
  const priorMonthKey = idx > 0 ? monthKeysWithData[idx - 1] : null;

  const currentTotals = computeRowTotals(bankRows.filter((row) => row.monthKey === monthKey));
  const currentDiscountRate = currentTotals.revenue ? (currentTotals.discount / currentTotals.revenue) * 100 : null;
  const currentAdmits = admitsMap.get(monthKey) || 0;
  const currentATV = currentAdmits ? currentTotals.ticketRevenue / currentAdmits : null;

  if (!priorMonthKey) {
    return { hasPrior: false, currentTotals, currentDiscountRate, currentAdmits, currentATV };
  }

  const priorTotals = computeRowTotals(bankRows.filter((row) => row.monthKey === priorMonthKey));
  const priorDiscountRate = priorTotals.revenue ? (priorTotals.discount / priorTotals.revenue) * 100 : null;
  const totalDeltaPercent = priorTotals.revenue ? ((currentTotals.revenue - priorTotals.revenue) / priorTotals.revenue) * 100 : null;
  const txnDeltaPercent = priorTotals.transactions ? ((currentTotals.transactions - priorTotals.transactions) / priorTotals.transactions) * 100 : null;
  const priorAdmits = admitsMap.get(priorMonthKey) || 0;
  const admitsDeltaPercent = priorAdmits ? ((currentAdmits - priorAdmits) / priorAdmits) * 100 : null;
  const priorATV = priorAdmits ? priorTotals.ticketRevenue / priorAdmits : null;
  const atvDeltaPercent = priorATV ? ((currentATV - priorATV) / priorATV) * 100 : null;

  return {
    hasPrior: true,
    priorMonthKey,
    currentTotals,
    priorTotals,
    currentDiscountRate,
    priorDiscountRate,
    totalDeltaPercent,
    txnDeltaPercent,
    discountRateNote: computeDiscountRateNote(currentDiscountRate, priorDiscountRate),
    volumeValueNote: computeVolumeValueNote(totalDeltaPercent, txnDeltaPercent),
    currentAdmits,
    priorAdmits,
    admitsDeltaPercent,
    admitsNote: computeAdmitsNote(admitsDeltaPercent),
    currentATV,
    priorATV,
    atvDeltaPercent,
    atvNote: computeATVNote(atvDeltaPercent),
    headline: generateHeadlineInsight({ revenueDeltaPercent: totalDeltaPercent, admitsDeltaPercent, atvDeltaPercent }),
  };
}

function inferOfferType(offerName) {
  const label = String(offerName || "").toLowerCase();
  if (!label) return "Unknown";
  if (label.includes("b1g1") || label.includes("buy 1 get 1")) return "B1G1";
  if (label.includes("free ticket")) return "Free Ticket";
  if (label.includes("%") || label.includes("percent")) return "Discount %";
  if (label.includes("cashback")) return "Cashback";
  if (label.includes("flat")) return "Flat Discount";
  if (label.includes("off")) return "Discount Offer";
  return "Offer";
}

function buildColumnLookup(headers, columnMap = COLUMN_MAP) {
  const normalizedHeaders = {};
  headers.forEach((header) => {
    normalizedHeaders[normalizeKey(header)] = header;
  });
  return Object.entries(columnMap).reduce((lookup, [field, expected]) => {
    lookup[field] = normalizedHeaders[normalizeKey(expected)] || null;
    return lookup;
  }, {});
}

function parseWorkbookRows(rows) {
  if (!rows.length) {
    return {
      parsedRows: [],
      missingColumns: Object.entries(COLUMN_MAP)
        .filter(([field]) => !OPTIONAL_COLUMNS.has(field))
        .map(([, value]) => value),
    };
  }

  const lookup = buildColumnLookup(Object.keys(rows[0]));
  const missingColumns = Object.entries(lookup)
    .filter(([field, column]) => !column && !OPTIONAL_COLUMNS.has(field))
    .map(([field]) => COLUMN_MAP[field]);

  const parsedRows = rows.map((row, index) => {
    const date = parseExcelDate(lookup.date ? row[lookup.date] : null);
    const offerName = lookup.offerName ? String(row[lookup.offerName] || "").trim() : "";
    const bankName = lookup.bankName ? String(row[lookup.bankName] || "").trim() : "";

    return {
      id: `${bankName || "bank"}-${offerName || "offer"}-${index}`,
      offerName: offerName || "Unknown Offer",
      bankName: bankName || "Unknown Bank",
      paymentCategory: "Card",
      discountedTransactions: parseNumber(lookup.discountedTransactions ? row[lookup.discountedTransactions] : 0),
      freeTickets: parseNumber(lookup.freeTickets ? row[lookup.freeTickets] : 0),
      totalTickets: parseNumber(lookup.totalTickets ? row[lookup.totalTickets] : 0),
      transactionTotal: parseNumber(lookup.transactionTotal ? row[lookup.transactionTotal] : 0),
      ticketRevenue: parseNumber(lookup.ticketRevenue ? row[lookup.ticketRevenue] : 0),
      fnbRevenue: parseNumber(lookup.fnbRevenue ? row[lookup.fnbRevenue] : 0),
      amountPaid: parseNumber(lookup.amountPaid ? row[lookup.amountPaid] : 0),
      discountAmount: parseNumber(lookup.discountAmount ? row[lookup.discountAmount] : 0),
      bankContribution: parseNumber(lookup.bankContribution ? row[lookup.bankContribution] : 0),
      inoxContribution: parseNumber(lookup.inoxContribution ? row[lookup.inoxContribution] : 0),
      convFees: parseNumber(lookup.convFees ? row[lookup.convFees] : 0),
      date,
      dateLabel: formatDateLabel(date),
      monthKey: monthKeyFromDate(date),
      fiscalYear: getFiscalYearLabel(date),
      offerType: inferOfferType(offerName),
      transactionType: lookup.transactionType ? String(row[lookup.transactionType] || "").trim() || "Unknown" : "Unknown",
      universalTransactions:
        lookup.universalTransactions && row[lookup.universalTransactions] !== undefined && row[lookup.universalTransactions] !== ""
          ? Number(row[lookup.universalTransactions])
          : null,
      admits:
        lookup.admits && row[lookup.admits] !== undefined && row[lookup.admits] !== "" ? Number(row[lookup.admits]) : null,
      universalTicketRevenue:
        lookup.universalTicketRevenue && row[lookup.universalTicketRevenue] !== undefined && row[lookup.universalTicketRevenue] !== ""
          ? Number(row[lookup.universalTicketRevenue])
          : null,
      universalTotalRevenue:
        lookup.universalTotalRevenue && row[lookup.universalTotalRevenue] !== undefined && row[lookup.universalTotalRevenue] !== ""
          ? Number(row[lookup.universalTotalRevenue])
          : null,
    };
  });

  return { parsedRows, missingColumns };
}

function buildUpiRow(rawRow, lookup, index, previousDate) {
  const date = resolveUpiMonthDate(lookup.month ? rawRow[lookup.month] : null, previousDate);
  const rawBankName = String(rawRow[lookup.bankName] || "").trim();
  const bankName = normalizeUpiPartnerName(rawBankName);
  const transactionTotal = parseNumber(lookup.transactionTotal ? rawRow[lookup.transactionTotal] : 0);
  const discountAmount = lookup.discountAmount ? parseNumber(rawRow[lookup.discountAmount]) : 0;
  const inoxContribution = lookup.inoxContribution ? parseNumber(rawRow[lookup.inoxContribution]) : 0;
  const bankContribution = lookup.bankContribution
    ? parseNumber(rawRow[lookup.bankContribution])
    : Math.max(0, discountAmount - inoxContribution);
  const totalTickets = lookup.totalTickets ? parseNumber(rawRow[lookup.totalTickets]) : 0;
  const discountedTransactions = lookup.discountedTransactions ? parseNumber(rawRow[lookup.discountedTransactions]) : totalTickets;
  const offerName = `${bankName || "Unknown Bank"} UPI Cashback`;

  return {
    id: `${bankName || "upi"}-${offerName}-${index}`,
    offerName,
    bankName: bankName || "Unknown Bank",
    paymentCategory: "UPI",
    discountedTransactions,
    freeTickets: 0,
    totalTickets,
    transactionTotal,
    ticketRevenue: transactionTotal,
    fnbRevenue: 0,
    amountPaid: transactionTotal - discountAmount,
    discountAmount,
    bankContribution,
    inoxContribution,
    convFees: 0,
    date,
    dateLabel: formatDateLabel(date),
    monthKey: monthKeyFromDate(date),
    fiscalYear: getFiscalYearLabel(date),
    offerType: "UPI Cashback",
    transactionType: "Online",
    universalTransactions: null,
    admits: null,
    universalTicketRevenue: null,
    universalTotalRevenue: null,
  };
}

function parseUpiWorkbookRows(workbook) {
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  if (!rawRows.length) {
    return {
      parsedRows: [],
      missingColumns: Object.entries(UPI_COLUMN_MAP)
        .filter(([field]) => !UPI_OPTIONAL_COLUMNS.has(field))
        .map(([, value]) => value),
    };
  }

  const lookup = buildColumnLookup(Object.keys(rawRows[0]), UPI_COLUMN_MAP);
  const missingColumns = Object.entries(lookup)
    .filter(([field, column]) => !column && !UPI_OPTIONAL_COLUMNS.has(field))
    .map(([field]) => UPI_COLUMN_MAP[field]);

  let previousDate = null;
  const parsedRows = rawRows.map((rawRow, index) => {
    const row = buildUpiRow(rawRow, lookup, index, previousDate);
    if (row.date) previousDate = row.date;
    return row;
  });

  return { parsedRows, missingColumns };
}

function getMonthlyReferenceValue(rows, field) {
  const map = new Map();
  const inconsistent = new Set();
  rows.forEach((row) => {
    if (row[field] === null || row.monthKey === "Unknown") return;
    if (map.has(row.monthKey) && map.get(row.monthKey) !== row[field]) inconsistent.add(row.monthKey);
    map.set(row.monthKey, row[field]);
  });
  return { map, inconsistent };
}

function sumAdmitsForMonths(monthKeys, admitsMap) {
  return [...new Set(monthKeys)].reduce((sum, key) => sum + (admitsMap.get(key) || 0), 0);
}

function formatMonthKeyLabel(monthKey) {
  const [month, year] = monthKey.split("-").map(Number);
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

function computeKpis(rows) {
  return rows.reduce(
    (acc, row) => {
      acc.totalTransactions += row.discountedTransactions;
      acc.grossRevenue += row.transactionTotal;
      acc.netRevenue += row.amountPaid;
      acc.totalDiscount += row.discountAmount;
      acc.ticketRevenue += row.ticketRevenue;
      acc.totalTickets += row.totalTickets || 0;
      return acc;
    },
    { ...EMPTY_KPIS },
  );
}

function filterGroupFiscal(rows, banks, fiscalYearsSelected, monthsSelected) {
  if (!fiscalYearsSelected.length || !monthsSelected.length) return [];
  return rows.filter(
    (r) =>
      r.fiscalYear &&
      fiscalYearsSelected.includes(r.fiscalYear) &&
      r.date &&
      monthsSelected.includes(MONTH_NAMES[r.date.getMonth()]) &&
      (!banks.length || banks.includes(r.bankName)),
  );
}

function aggregateGroupBankBreakdown(groupRows) {
  const grouped = new Map();
  groupRows.forEach((row) => {
    const current = grouped.get(row.bankName) || {
      bankName: row.bankName,
      paymentCategory: row.paymentCategory,
      revenue: 0,
      transactions: 0,
      discount: 0,
      admits: 0,
    };
    current.revenue += row.transactionTotal;
    current.transactions += row.discountedTransactions;
    current.discount += row.discountAmount;
    current.admits += row.totalTickets || 0;
    grouped.set(row.bankName, current);
  });
  return [...grouped.values()].sort((a, b) => b.revenue - a.revenue);
}

function aggregateBanks(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const current = grouped.get(row.bankName) || {
      bankName: row.bankName,
      totalRevenue: 0,
      discountCost: 0,
      totalTransactions: 0,
      paymentCategory: row.paymentCategory,
    };
    current.totalRevenue += row.transactionTotal;
    current.discountCost += row.discountAmount;
    current.totalTransactions += row.discountedTransactions;
    grouped.set(row.bankName, current);
  });
  return [...grouped.values()].sort((a, b) => b.totalRevenue - a.totalRevenue);
}

function normalizeOfferChannel(offerName) {
  if (/F&B/i.test(offerName)) return offerName;

  let name = offerName.trim();

  // Strip trailing channel/sub-channel suffixes (repeat in case of nested suffixes)
  const channelSuffixPattern = /\s*-\s*(Box Office|Digital Platforms|Online|Offline|Refuel)\s*$/i;
  while (channelSuffixPattern.test(name)) {
    name = name.replace(channelSuffixPattern, "").trim();
  }

  // Strip everything up through the last "<...> Card -" prefix (bank name + card type),
  // e.g. "DBS Bank Credit & Debit Card - " or "DBS Debit Card - "
  name = name.replace(/^.*\bCard\b\s*[-–]\s*/i, "").trim();

  // Normalize spacing around % signs so "20 % off" and "20% off" match
  name = name.replace(/(\d+)\s*%/g, "$1%");

  // Normalize spacing around colons so "Offer:Details" and "Offer :  Details" match
  name = name.replace(/:\s*/g, ": ");

  // Collapse any repeated whitespace
  name = name.replace(/\s+/g, " ").trim();

  return name;
}

const OFFER_ALIAS_MAP = {
  "Kotak Cashback+ Credit Card – BOGO Festive Offer - Digital Platforms": "Kotak BOGO Festive Offer",
  "Kotak Solitaire Credit Card – BOGO Festive Offer - Digital Platforms": "Kotak BOGO Festive Offer",
  "Kotak PVR INOX Credit Card 5% Off on Movie Ticket": "Kotak PVR INOX Credit Card 5% Off on Tickets",
  "Buy One Get One": "Buy 1 Get 1",
  "J&K Bank Mastercard DC - 50% off Movie-Srinagar": "J&K Bank Mastercard - 50% off Movie-Srinagar",
  "J&K Bank Mastercard CC - 50% off Movie-Srinagar": "J&K Bank Mastercard - 50% off Movie-Srinagar",
  // Add more raw-name -> canonical-name pairs here as you find them.
};

function canonicalOfferName(offerName) {
  return OFFER_ALIAS_MAP[offerName] || normalizeOfferChannel(offerName);
}

function aggregateOffers(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const canonicalName = canonicalOfferName(row.offerName);
    const key = `${canonicalName}::${row.bankName}`;
    const current = grouped.get(key) || {
      offerName: canonicalName,
      bankName: row.bankName,
      revenue: 0,
      netRevenue: 0,
      discount: 0,
      transactions: 0,
      freeTickets: 0,
      totalTickets: 0,
      ticketRevenue: 0,
      fnbRevenue: 0,
      bankContribution: 0,
      inoxContribution: 0,
      convFees: 0,
      dates: new Set(),
      offerType: row.offerType,
      paymentCategory: row.paymentCategory,
    };
    current.revenue += row.transactionTotal;
    current.netRevenue += row.amountPaid;
    current.discount += row.discountAmount;
    current.transactions += row.discountedTransactions;
    current.freeTickets += row.freeTickets;
    current.totalTickets += row.totalTickets || 0;
    current.ticketRevenue += row.ticketRevenue;
    current.fnbRevenue += row.fnbRevenue;
    current.bankContribution += row.bankContribution;
    current.inoxContribution += row.inoxContribution;
    current.convFees += row.convFees;
    current.dates.add(row.dateLabel);
    grouped.set(key, current);
  });

  return [...grouped.values()]
    .map((offer) => ({
      ...offer,
      dates: [...offer.dates].sort(),
      profit: offer.revenue - offer.inoxContribution,
    }))
    .sort((a, b) => a.bankName.localeCompare(b.bankName) || b.revenue - a.revenue);
}

function aggregateBubbleData(rows, colorMap) {
  return aggregateBanks(rows).map((bank) => ({
    name: bank.bankName,
    discount: bank.discountCost,
    revenue: bank.totalRevenue,
    transactions: bank.totalTransactions,
    fill: colorMap[bank.bankName],
  }));
}

function aggregateChannelRevenue(rows) {
  const grouped = new Map();
  const displayNames = { "Offline - Box Office": "Offline BO" };
  rows.forEach((row) => {
    const rawKey = ["Online", "Offline - Box Office", "Offline - F&B"].includes(row.transactionType) ? row.transactionType : "Offline - Box Office";
    const key = displayNames[rawKey] || rawKey;
    grouped.set(key, (grouped.get(key) || 0) + row.transactionTotal);
  });
  return [...grouped.entries()]
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

function aggregateMonthlySeries(rows, selectedBanks) {
  const monthMap = new Map();
  rows.forEach((row) => {
    if (!selectedBanks.includes(row.bankName)) return;
    if (!monthMap.has(row.monthKey)) monthMap.set(row.monthKey, { monthKey: row.monthKey });
    const current = monthMap.get(row.monthKey);
    current[row.bankName] = (current[row.bankName] || 0) + row.transactionTotal;
  });
  return [...monthMap.values()].sort((a, b) => {
    if (a.monthKey === "Unknown" && b.monthKey === "Unknown") return 0;
    if (a.monthKey === "Unknown") return 1;
    if (b.monthKey === "Unknown") return -1;
    const [am, ay] = a.monthKey.split("-").map(Number);
    const [bm, by] = b.monthKey.split("-").map(Number);
    return new Date(ay, am - 1, 1) - new Date(by, bm - 1, 1);
  });
}

function aggregateSeasonalByYear(rows, selectedBanks, metric = "revenue") {
  const field = SEASONAL_METRICS.find((entry) => entry.key === metric)?.field || "transactionTotal";
  const seasonMap = new Map();
  rows.forEach((row) => {
    if (!selectedBanks.includes(row.bankName)) return;
    if (row.monthKey === "Unknown" || !row.fiscalYear) return;
    const [month] = row.monthKey.split("-").map(Number);
    if (!seasonMap.has(month)) seasonMap.set(month, { month, monthLabel: MONTH_NAMES[month - 1] });
    const current = seasonMap.get(month);
    current[row.fiscalYear] = (current[row.fiscalYear] || 0) + row[field];
  });
  return [...seasonMap.values()].sort((a, b) => {
    const fiscalIndex = (m) => (m >= 4 ? m - 4 : m + 8);
    return fiscalIndex(a.month) - fiscalIndex(b.month);
  });
}

function aggregateYearlyTotals(rows, selectedBanks) {
  const yearMap = new Map();
  rows.forEach((row) => {
    if (!selectedBanks.includes(row.bankName)) return;
    if (!row.fiscalYear || row.fiscalYear === "Unknown") return;
    yearMap.set(row.fiscalYear, (yearMap.get(row.fiscalYear) || 0) + row.transactionTotal);
  });
  return [...yearMap.entries()]
    .map(([year, revenue]) => ({ year, revenue }))
    .sort((a, b) => Number(a.year.split("-")[0]) - Number(b.year.split("-")[0]));
}

function safeRatio(numerator, denominator) {
  if (!denominator) return 0;
  return (numerator / denominator) * 100;
}

function computeUpliftOrContribution(bankValue, universalValue) {
  if (!universalValue || bankValue === null) return null;
  if (bankValue > universalValue) {
    return { type: "uplift", percent: ((bankValue - universalValue) / universalValue) * 100 };
  }
  return { type: "contribution", percent: (bankValue / universalValue) * 100 };
}

function UpliftOrContributionLine({ comparison }) {
  if (!comparison) return null;
  if (comparison.type === "uplift") {
    return <p className="text-xs font-bold text-emerald-600">▲ {comparison.percent.toFixed(0)}% uplift vs universal</p>;
  }
  return <p className="text-xs text-textMuted">Bank is {comparison.percent.toFixed(1)}% of universal</p>;
}

function StatCard({ title, value, subtitle, color, icon, delta, extra }) {
  return (
    <div className={`min-w-0 rounded-3xl border p-3 shadow-soft ${color.accent}`}>
      <div className="flex items-start justify-between gap-3">
        <p className="min-w-0 flex-1 text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">{title}</p>
        <div
          className={`flex h-8 min-w-[50px] flex-shrink-0 items-center justify-center rounded-2xl px-3 ${color.dot} text-[14px] font-semibold text-white`}
          style={{ whiteSpace: "nowrap" }}
        >
          {icon}
        </div>
      </div>
      <p className="mt-1.5 overflow-hidden text-ellipsis whitespace-nowrap text-[18px] font-bold leading-none text-textMain">{value}</p>
      {subtitle ? <p className="mt-1 text-sm font-medium text-slate-500">{subtitle}</p> : null}
      {delta ? (
        <p
          className={`mt-2 inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-bold ${
            delta.value === null
              ? "bg-slate-100 text-slate-500"
              : delta.value > 0
                ? "bg-emerald-50 text-emerald-600"
                : delta.value < 0
                  ? "bg-rose-50 text-rose-600"
                  : "bg-slate-100 text-slate-500"
          }`}
        >
          {delta.value === null
            ? "No prior data"
            : `${delta.value > 0 ? "▲" : delta.value < 0 ? "▼" : "–"} ${Math.abs(delta.value).toFixed(1)}% ${delta.label}`}
        </p>
      ) : null}
      {extra ? <div className="mt-2 space-y-0.5">{extra}</div> : null}
    </div>
  );
}

function MultiSelectDropdown({ options, selected, onToggle, onClear, label, showSelectedNames = false }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const allSelected = options.length > 0 && selected.length === options.length;

  useEffect(() => {
    function handleClick(event) {
      if (!panelRef.current?.contains(event.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  return (
    <div ref={panelRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen((value) => !value)}
        className="flex w-full min-w-[220px] items-center justify-between rounded-2xl border border-borderSoft bg-white px-4 py-3 text-left text-sm font-semibold text-textMain shadow-sm"
      >
        <span className="truncate">
          {selected.length === 0
            ? `No ${label}s selected`
            : allSelected
              ? `All ${label.charAt(0).toUpperCase()}${label.slice(1)}s`
              : showSelectedNames
                ? selected.join(", ")
                : `${selected.length} ${label}${selected.length > 1 ? "s" : ""} selected`}
        </span>
        <span className="text-textMuted">{open ? "▲" : "▼"}</span>
      </button>
      {open ? (
        <div className="absolute right-0 z-[60] mt-2 w-full rounded-2xl border border-borderSoft bg-white p-3 shadow-soft">
          <div className="mb-2 flex items-center justify-between">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">{label}s</p>
            <button type="button" onClick={onClear} className="text-xs font-bold text-accentBlue">
              Clear
            </button>
          </div>
          <div className="max-h-56 space-y-2 overflow-y-auto pr-1 scrollbar-thin">
            <label className="flex cursor-pointer items-center gap-3 rounded-xl px-2 py-2 hover:bg-slate-50">
              <input
                type="checkbox"
                checked={allSelected}
                onChange={(event) => {
                  if (event.target.checked) {
                    options.forEach((option) => {
                      if (!selected.includes(option)) {
                        onToggle(option);
                      }
                    });
                  } else {
                    onClear();
                  }
                }}
                className="h-4 w-4 rounded border-slate-300 text-accentBlue focus:ring-accentBlue"
              />
              <span className="text-sm font-semibold text-textMain">Select All</span>
            </label>
            {options.map((option) => (
              <label key={option} className="flex cursor-pointer items-center gap-3 rounded-xl px-2 py-2 hover:bg-slate-50">
                <input
                  type="checkbox"
                  checked={selected.includes(option)}
                  onChange={() => onToggle(option)}
                  className="h-4 w-4 rounded border-slate-300 text-accentBlue focus:ring-accentBlue"
                />
                <span className="text-sm font-semibold text-textMain">{option}</span>
              </label>
            ))}
          </div>
        </div>
      ) : null}
    </div>
  );
}

function OfferModal({ offer, onClose }) {
  if (!offer) return null;

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 px-4 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-5xl max-h-[85vh] overflow-y-auto rounded-[2rem] border border-white/60 bg-white p-6 shadow-soft scrollbar-thin"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">
              {offer.bankName}
              {offer.paymentCategory === "UPI" ? " (UPI)" : ""}
            </p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">{offer.offerName}</h3>
            <p className="mt-2 inline-flex rounded-full bg-blue-50 px-3 py-1 text-xs font-bold text-accentBlue">{offer.offerType}</p>
          </div>
          <button type="button" onClick={onClose} className="rounded-full border border-borderSoft px-3 py-2 text-sm font-bold text-textMuted hover:bg-slate-50">
            Close
          </button>
        </div>
        <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <MetricBlock title="Revenue" value={formatInLakh(offer.revenue)} />
          <MetricBlock title="Net Revenue" value={formatInLakh(offer.netRevenue)} />
          <MetricBlock title="Discount" value={formatInLakh(offer.discount)} />
          <MetricBlock title="Transactions" value={formatCountInLakh(offer.transactions)} />
          <MetricBlock title="Free Tickets" value={formatCountInLakh(offer.freeTickets)} />
        </div>
        <div className="mt-4 grid gap-4 sm:grid-cols-2">
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Ticket & F&B</p>
            <p className="mt-3 text-sm font-semibold text-textMain">Ticket Revenue: {formatInLakh(offer.ticketRevenue)}</p>
            <p className="mt-2 text-sm font-semibold text-textMain">F&B Revenue: {formatInLakh(offer.fnbRevenue)}</p>
            <p className="mt-2 text-sm font-semibold text-textMain">Total Tickets: {formatCountInLakh(offer.totalTickets)}</p>
            {offer.paymentCategory === "UPI" ? (
              <p className="mt-2 text-xs text-textMuted">
                UPI transaction counts are estimated from ticket count (~57% of source rows had no separate transaction figure).
              </p>
            ) : null}
          </div>
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Contribution Split</p>
            <p className="mt-3 text-sm font-semibold text-textMain">
              Bank: {formatInLakh(offer.bankContribution)} ({offer.discount ? ((offer.bankContribution / offer.discount) * 100).toFixed(0) : 0}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">
              Inox: {formatInLakh(offer.inoxContribution)} ({offer.discount ? ((offer.inoxContribution / offer.discount) * 100).toFixed(0) : 0}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">Conv. Fees: {formatInLakh(offer.convFees)}</p>
          </div>
        </div>
        <div className="mt-4 rounded-2xl border border-borderSoft p-4">
          <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Available Dates</p>
          <p className="mt-3 text-sm font-semibold text-textMain">{offer.dates.join(", ") || "No date data"}</p>
        </div>
      </div>
    </div>
  );
}

function BankModal({ bank, offersEntry, discountEntry, onClose }) {
  if (!bank) return null;

  const today = new Date();
  const offers = offersEntry?.offers || [];
  const activeOffers = offers.filter((offer) => offer.endDate && offer.endDate >= today);
  const closedOffers = offers.filter((offer) => !offer.endDate || offer.endDate < today);

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 px-4 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-5xl max-h-[85vh] overflow-y-auto rounded-[2rem] border border-white/60 bg-white p-6 shadow-soft scrollbar-thin"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Bank / UPI Scorecard</p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">{bank.bankName}</h3>
            <p className="mt-2 inline-flex rounded-full bg-blue-50 px-3 py-1 text-xs font-bold text-accentBlue">{formatInteger(offers.length)} offers</p>
          </div>
          <button type="button" onClick={onClose} className="rounded-full border border-borderSoft px-3 py-2 text-sm font-bold text-textMuted hover:bg-slate-50">
            Close
          </button>
        </div>
        <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <MetricBlock title="Total Revenue" value={formatInLakh(bank.totalRevenue)} />
          <MetricBlock title="Discount Cost" value={formatInLakh(bank.discountCost)} />
          <MetricBlock title="Total Transactions" value={formatCountInLakh(bank.totalTransactions)} />
        </div>
        <div className="mt-4 grid gap-4 lg:grid-cols-2">
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="mb-3 text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Offers</p>
            <div className="max-h-64 overflow-y-auto pr-1 scrollbar-thin">
              <div className="grid gap-4 lg:grid-cols-2">
                <div>
                  <p className="mb-2 text-xs font-bold uppercase tracking-[0.2em] text-emerald-600">Active Offers ({activeOffers.length})</p>
                  {activeOffers.length ? activeOffers.map((offer) => (
                    <div key={offer.offerName} className="mb-2 rounded-xl border border-emerald-100 bg-emerald-50 px-3 py-2 text-sm">
                      <p className="font-semibold text-textMain">{offer.offerName}</p>
                      <p className="text-xs text-textMuted">{offer.startLabel} – {offer.endLabel}</p>
                    </div>
                  )) : <p className="text-xs text-textMuted">No active offers</p>}
                </div>
                <div>
                  <p className="mb-2 text-xs font-bold uppercase tracking-[0.2em] text-rose-500">Closed Offers ({closedOffers.length})</p>
                  {closedOffers.length ? closedOffers.map((offer) => (
                    <div key={offer.offerName} className="mb-2 rounded-xl border border-rose-100 bg-rose-50 px-3 py-2 text-sm">
                      <p className="font-semibold text-textMain">{offer.offerName}</p>
                      <p className="text-xs text-textMuted">{offer.startLabel} – {offer.endLabel}</p>
                    </div>
                  )) : <p className="text-xs text-textMuted">No closed offers</p>}
                </div>
              </div>
            </div>
          </div>
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Discount Split - Bank vs PVR</p>
            <p className="mt-3 text-sm font-semibold text-textMain">
              Bank: {formatInLakh(discountEntry?.bankDiscount || 0)} ({(discountEntry?.bankPercent || 0).toFixed(0)}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">
              PVR: {formatInLakh(discountEntry?.pvrDiscount || 0)} ({(discountEntry?.pvrPercent || 0).toFixed(0)}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">Total: {formatInLakh(discountEntry?.totalDiscount || 0)}</p>
          </div>
        </div>
      </div>
    </div>
  );
}

function OffersByBankModal({
  offersByBank,
  totalOfferCountByBank,
  expandedOfferBank,
  onToggleBank,
  onClose,
  heading = "Total Offers by Each Bank",
  entityNoun = "banks",
}) {
  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 px-4 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-4xl max-h-[85vh] overflow-y-auto rounded-[2rem] border border-white/60 bg-white p-6 shadow-soft scrollbar-thin"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Info Section</p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">{heading}</h3>
            <p className="mt-2 text-sm font-semibold text-textMuted">
              {formatInteger(totalOfferCountByBank)} total offers across {formatInteger(offersByBank.length)} {entityNoun}
            </p>
          </div>
          <button type="button" onClick={onClose} className="rounded-full border border-borderSoft px-3 py-2 text-sm font-bold text-textMuted hover:bg-slate-50">
            Close
          </button>
        </div>

        <div className="mt-4 overflow-hidden rounded-3xl border border-borderSoft">
          <div className="max-h-[420px] overflow-y-scroll scrollbar-thin">
            <div className="divide-y divide-borderSoft">
              {offersByBank.length ? offersByBank.map((entry, index) => {
                const isExpanded = expandedOfferBank === entry.bankName;
                return (
                  <div key={entry.bankName} className={index % 2 === 0 ? "bg-white" : "bg-slate-50/60"}>
                    <button
                      type="button"
                      onClick={() => onToggleBank(entry.bankName)}
                      className="flex w-full items-center justify-between gap-3 px-4 py-4 text-left transition hover:bg-blue-50/60"
                    >
                      <div className="min-w-0">
                        <p className="truncate font-bold text-textMain">{entry.bankName}</p>
                      </div>
                      <div className="flex items-center gap-4">
                        <span className="text-sm font-semibold text-textMain">{formatInteger(entry.offerCount)}</span>
                        <span className="flex h-9 w-9 flex-shrink-0 items-center justify-center rounded-2xl border border-borderSoft bg-white text-sm font-bold text-textMuted">
                          {isExpanded ? "▲" : "▼"}
                        </span>
                      </div>
                    </button>
                    {isExpanded ? (
                      <div className="border-t border-borderSoft bg-white px-4 py-3">
                        <div className="max-h-64 overflow-y-auto rounded-2xl border border-borderSoft">
                          <table className="min-w-full divide-y divide-borderSoft text-sm">
                            <thead className="sticky top-0 bg-slate-50/95 backdrop-blur">
                              <tr>
                                <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Offer Name</th>
                                <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Date Range</th>
                                <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Discount Split</th>
                              </tr>
                            </thead>
                            <tbody>
                              {entry.offers.length ? entry.offers.map((offer, offerIndex) => (
                                <tr key={`${entry.bankName}-${offer.offerName}`} className={offerIndex % 2 === 0 ? "bg-white" : "bg-slate-50/60"}>
                                  <td className="px-4 py-3 font-semibold text-textMain">{offer.offerName}</td>
                                  <td className="px-4 py-3 font-semibold text-textMain">
                                    {offer.startLabel} to {offer.endLabel}
                                  </td>
                                  <td className="px-4 py-3 font-semibold text-textMain">
                                    Bank: {formatInLakh(offer.bankDiscount)}
                                    <br />
                                    PVR: {formatInLakh(offer.pvrDiscount)}
                                  </td>
                                </tr>
                              )) : (
                                <tr>
                                  <td colSpan="3" className="px-4 py-8 text-center font-semibold text-textMuted">
                                    No offers available for this bank.
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    ) : null}
                  </div>
                );
              }) : (
                <div className="px-4 py-10 text-center font-semibold text-textMuted">
                  Upload data to view total offers by {entityNoun.slice(0, -1)}.
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function GroupDetailModal({ group, breakdown, onClose }) {
  if (!group) return null;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div className="max-h-[80vh] w-full max-w-2xl overflow-y-auto rounded-2xl bg-white p-5 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="font-display text-xl font-bold text-textMain">Group {group} — Bank/UPI Breakdown</h3>
          <button type="button" onClick={onClose} className="text-sm font-bold text-textMuted hover:text-textMain">
            ✕ Close
          </button>
        </div>
        <table className="min-w-full text-sm">
          <thead>
            <tr className="text-left text-xs font-bold uppercase tracking-[0.15em] text-textMuted">
              <th className="py-2 pr-4">Bank / Partner</th>
              <th className="py-2 pr-4">Revenue</th>
              <th className="py-2 pr-4">Transactions</th>
              <th className="py-2 pr-4">Discount Amount</th>
              <th className="py-2 pr-4">Admits</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-borderSoft">
            {breakdown.map((entry) => (
              <tr key={entry.bankName}>
                <td className="py-2 pr-4 font-semibold text-textMain">
                  {entry.bankName}
                  {entry.paymentCategory === "UPI" ? " (UPI)" : ""}
                </td>
                <td className="py-2 pr-4 text-textMain">{formatInLakh(entry.revenue)}</td>
                <td className="py-2 pr-4 text-textMain">{formatCountInLakh(entry.transactions)}</td>
                <td className="py-2 pr-4 text-textMain">{formatInLakh(entry.discount)}</td>
                <td className="py-2 pr-4 text-textMain">{formatCountInLakh(entry.admits)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DiscountSplitTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  const data = payload[0].payload;

  return (
    <div
      style={{ borderRadius: "18px", borderColor: "#e2e8f0", background: "#fff", maxWidth: "220px", whiteSpace: "normal" }}
      className="border p-3 shadow-soft"
    >
      <p className="text-sm font-bold text-textMain">{label}</p>
      <p className="mt-1 text-sm font-semibold text-textMain">
        Bank Contribution: {formatInLakh(data.bankDiscount)} ({data.bankPercent.toFixed(1)}%)
      </p>
      <p className="mt-1 text-sm font-semibold text-textMain">
        PVR Contribution: {formatInLakh(data.pvrDiscount)} ({data.pvrPercent.toFixed(1)}%)
      </p>
    </div>
  );
}

function MonthlyTrendTooltip({ active, payload, label, formatValue = formatInLakh }) {
  if (!active || !payload?.length) return null;

  return (
    <div style={{ borderRadius: "18px", borderColor: "#e2e8f0", background: "#fff" }} className="border p-3 shadow-soft">
      <p className="text-sm font-bold text-textMain">{label}</p>
      <div className="scrollbar-thin mt-2 space-y-1 overflow-y-auto pr-1" style={{ maxHeight: "260px" }}>
        {payload.map((entry) => (
          <p key={entry.dataKey} className="flex items-center gap-2 text-sm font-semibold text-textMain">
            <span className="h-2.5 w-2.5 flex-shrink-0 rounded-full" style={{ backgroundColor: entry.color }} />
            <span className="min-w-0 flex-1 truncate">{entry.name}</span>
            <span>{formatValue(entry.value)}</span>
          </p>
        ))}
      </div>
    </div>
  );
}

function MetricBlock({ title, value }) {
  return (
    <div className="rounded-2xl bg-slate-50 p-4">
      <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">{title}</p>
      <p className="mt-2 text-lg font-extrabold text-textMain">{value}</p>
    </div>
  );
}

function MetricRow({ label, current, format, deltaPercent }) {
  const known = deltaPercent !== null && deltaPercent !== undefined && !Number.isNaN(deltaPercent);
  const positive = known && deltaPercent >= 0;
  return (
    <div className="flex items-center justify-between border-b border-borderSoft/60 py-2 last:border-0">
      <span className="text-xs font-semibold text-textMuted">{label}</span>
      <div className="flex items-center gap-2">
        <span className="text-sm font-bold text-textMain">{current === null || current === undefined ? "—" : format(current)}</span>
        {known ? (
          <span className={`inline-flex items-center gap-0.5 text-xs font-bold ${positive ? "text-emerald-600" : "text-rose-500"}`}>
            {positive ? "▲" : "▼"} {Math.abs(deltaPercent).toFixed(1)}%
          </span>
        ) : (
          <span className="text-xs text-textMuted">—</span>
        )}
      </div>
    </div>
  );
}

function MetricComparisonBox({ title, metrics }) {
  return (
    <div className="flex-1 rounded-xl border border-borderSoft bg-white p-3">
      <p className="mb-1 text-xs font-bold uppercase tracking-[0.15em] text-textMuted">{title}</p>
      {metrics.map((m) => (
        <MetricRow key={m.label} {...m} />
      ))}
    </div>
  );
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [missingColumns, setMissingColumns] = useState([]);
  const [dragActive, setDragActive] = useState(false);
  const [upiFileName, setUpiFileName] = useState("");
  const [upiError, setUpiError] = useState("");
  const [upiDragActive, setUpiDragActive] = useState(false);
  const [fyFilter, setFyFilter] = useState([]);
  const [monthFilter, setMonthFilter] = useState([]);
  const [bankFilter, setBankFilter] = useState([]);
  const [offerFilter, setOfferFilter] = useState([]);
  const [paymentCategoryFilter, setPaymentCategoryFilter] = useState("all"); // "all" | "card" | "upi"
  const [selectedOffer, setSelectedOffer] = useState(null);
  const [selectedBank, setSelectedBank] = useState(null);
  const [hoveredBank, setHoveredBank] = useState(null);
  const [showOffersByBank, setShowOffersByBank] = useState(false);
  const [expandedOfferBank, setExpandedOfferBank] = useState(null);
  const [showOffersByUpi, setShowOffersByUpi] = useState(false);
  const [expandedOfferUpiPartner, setExpandedOfferUpiPartner] = useState(null);
  const [showComparisonModal, setShowComparisonModal] = useState(false);
  const [groupDetailOpen, setGroupDetailOpen] = useState(null); // null | "A" | "B"
  const [comparisonMode, setComparisonMode] = useState("none");
  const [trendMode, setTrendMode] = useState("monthly");
  const [seasonalMetric, setSeasonalMetric] = useState("revenue");
  const [selectedSeasonalPoint, setSelectedSeasonalPoint] = useState(null);
  const [selectedMonthlyPoint, setSelectedMonthlyPoint] = useState(null);
  const [selectedYearPoint, setSelectedYearPoint] = useState(null);
  const [bankSortKey, setBankSortKey] = useState(null);
  const [bankSortDir, setBankSortDir] = useState("desc");
  const [offerSortKey, setOfferSortKey] = useState(null);
  const [offerSortDir, setOfferSortDir] = useState("desc");
  const [groupABanks, setGroupABanks] = useState([]);
  const [groupAFy, setGroupAFy] = useState([]);
  const [groupAMonths, setGroupAMonths] = useState([]);
  const [groupBBanks, setGroupBBanks] = useState([]);
  const [groupBFy, setGroupBFy] = useState([]);
  const [groupBMonths, setGroupBMonths] = useState([]);
  const fileInputRef = useRef(null);
  const upiFileInputRef = useRef(null);

  const categoryScopedRows = useMemo(() => {
    if (paymentCategoryFilter === "all") return rows;
    const wantCategory = paymentCategoryFilter === "card" ? "Card" : "UPI";
    return rows.filter((row) => row.paymentCategory === wantCategory);
  }, [rows, paymentCategoryFilter]);

  const banks = useMemo(() => [...new Set(categoryScopedRows.map((row) => row.bankName))].sort(), [categoryScopedRows]);

  const offers = useMemo(() => {
    const bankScopedRows = bankFilter.length ? categoryScopedRows.filter((row) => bankFilter.includes(row.bankName)) : categoryScopedRows;
    return [...new Set(bankScopedRows.map((row) => canonicalOfferName(row.offerName)))].sort();
  }, [categoryScopedRows, bankFilter]);
  const fiscalYears = useMemo(() => {
    const set = new Set(rows.map((r) => r.fiscalYear).filter((fy) => fy && fy !== "Unknown"));
    return [...set].sort((a, b) => Number(a.split("-")[0]) - Number(b.split("-")[0]));
  }, [rows]);

  const fiscalMonths = useMemo(() => {
    const present = new Set(rows.map((r) => (r.date ? MONTH_NAMES[r.date.getMonth()] : null)));
    return FISCAL_MONTH_ORDER.filter((m) => present.has(m));
  }, [rows]);

  const latestDataDate = useMemo(
    () => rows.reduce((max, row) => (row.date && (!max || row.date > max) ? row.date : max), null),
    [rows],
  );

  const quickPeriodFilterSets = useMemo(() => {
    if (!latestDataDate) return {};
    const month = latestDataDate.getMonth();
    const fy = getFiscalYearLabel(latestDataDate);
    const fyStartYear = Number(fy.split("-")[0]) + 2000;
    const result = {};

    result.thisMonth = { fy, months: [MONTH_NAMES[month]] };

    const prior = new Date(latestDataDate.getFullYear(), month - 1, 1);
    result.lastMonth = { fy: getFiscalYearLabel(prior), months: [MONTH_NAMES[prior.getMonth()]] };

    const qtdMonths = [];
    for (let d = new Date(getFiscalQuarterStart(latestDataDate)); d <= latestDataDate; d.setMonth(d.getMonth() + 1)) {
      qtdMonths.push(MONTH_NAMES[d.getMonth()]);
    }
    result.qtd = { fy, months: qtdMonths };

    const fytdMonths = [];
    for (let d = new Date(fyStartYear, 3, 1); d <= latestDataDate; d.setMonth(d.getMonth() + 1)) {
      fytdMonths.push(MONTH_NAMES[d.getMonth()]);
    }
    result.fytd = { fy, months: fytdMonths };

    return result;
  }, [latestDataDate]);

  const activeQuickPeriodKey = useMemo(() => {
    if (!fyFilter.length && !monthFilter.length) return "clear";
    const sortedMonths = [...monthFilter].sort();
    const match = QUICK_PERIODS.find((preset) => {
      const target = quickPeriodFilterSets[preset.key];
      if (!target) return false;
      const sortedTarget = [...target.months].sort();
      return (
        fyFilter.length === 1 &&
        fyFilter[0] === target.fy &&
        sortedTarget.length === sortedMonths.length &&
        sortedTarget.every((m, index) => m === sortedMonths[index])
      );
    });
    return match ? match.key : null;
  }, [fyFilter, monthFilter, quickPeriodFilterSets]);

  const bankColorMap = useMemo(
    () =>
      banks.reduce((acc, bank, index) => {
        acc[bank] = BANK_COLORS[index % BANK_COLORS.length];
        return acc;
      }, {}),
    [banks],
  );

  useEffect(() => {
    setSelectedSeasonalPoint(null);
    setSelectedMonthlyPoint(null);
    setSelectedYearPoint(null);
  }, [seasonalMetric, trendMode, bankFilter]);

  useEffect(() => {
    setFyFilter(fiscalYears);
  }, [fiscalYears]);

  useEffect(() => {
    setMonthFilter(fiscalMonths);
  }, [fiscalMonths]);

  useEffect(() => {
    setBankFilter(banks);
  }, [banks]);

  useEffect(() => {
    setOfferFilter(offers);
  }, [offers]);

  function handleFileChange(file) {
    setFileName(file.name);
    setError("");
    setMissingColumns([]);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target?.result, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        if (!jsonRows.length) {
          setError("The uploaded file is empty.");
          return;
        }
        const { parsedRows, missingColumns: missing } = parseWorkbookRows(jsonRows);
        setRows((current) => [...current.filter((r) => r.paymentCategory === "UPI"), ...parsedRows]);
        setMissingColumns(missing);
        setSelectedOffer(null);
      } catch {
        setError("Unable to parse this file. Please upload a valid Excel sheet with the expected columns.");
      }
    };
    reader.onerror = () => {
      setError("There was a problem reading the file.");
    };
    reader.readAsArrayBuffer(file);
  }

  function handleRemoveFile() {
    setRows((current) => current.filter((r) => r.paymentCategory === "UPI"));
    setFileName("");
    setError("");
    setMissingColumns([]);
    setFyFilter([]);
    setMonthFilter([]);
    setBankFilter([]);
    setOfferFilter([]);
    setSelectedOffer(null);
    setSelectedBank(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  function handleUpiFileChange(file) {
    setUpiFileName(file.name);
    setUpiError("");

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target?.result, { type: "array" });
        const { parsedRows, missingColumns: missing } = parseUpiWorkbookRows(workbook);
        if (!parsedRows.length) {
          setUpiError("The uploaded file is empty.");
          return;
        }
        setRows((current) => [...current.filter((r) => r.paymentCategory !== "UPI"), ...parsedRows]);
        setUpiError(missing.length ? `Missing columns were treated as zero: ${missing.join(", ")}` : "");
        setSelectedOffer(null);
      } catch {
        setUpiError("Unable to parse this file. Please upload a valid Excel sheet with the expected columns.");
      }
    };
    reader.onerror = () => {
      setUpiError("There was a problem reading the file.");
    };
    reader.readAsArrayBuffer(file);
  }

  function handleRemoveUpiFile() {
    setRows((current) => current.filter((r) => r.paymentCategory !== "UPI"));
    setUpiFileName("");
    setUpiError("");
    if (upiFileInputRef.current) upiFileInputRef.current.value = "";
  }

  const filteredRows = useMemo(
    () =>
      rows.filter((row) => {
        const matchesDate = fyFilter.includes(row.fiscalYear) && (row.date ? monthFilter.includes(MONTH_NAMES[row.date.getMonth()]) : false);
        const matchesBank = bankFilter.includes(row.bankName);
        const matchesOffer = offerFilter.includes(canonicalOfferName(row.offerName));
        const matchesPaymentCategory =
          paymentCategoryFilter === "all" || row.paymentCategory === (paymentCategoryFilter === "card" ? "Card" : "UPI");
        return matchesDate && matchesBank && matchesOffer && matchesPaymentCategory;
      }),
    [rows, fyFilter, monthFilter, bankFilter, offerFilter, paymentCategoryFilter],
  );

  const bankOfferFilteredRows = useMemo(
    () =>
      rows.filter((row) => {
        const matchesBank = bankFilter.includes(row.bankName);
        const matchesOffer = offerFilter.includes(canonicalOfferName(row.offerName));
        const matchesPaymentCategory =
          paymentCategoryFilter === "all" || row.paymentCategory === (paymentCategoryFilter === "card" ? "Card" : "UPI");
        return matchesBank && matchesOffer && matchesPaymentCategory;
      }),
    [rows, bankFilter, offerFilter, paymentCategoryFilter],
  );

  const kpis = useMemo(() => {
    const base = computeKpis(filteredRows);
    return { ...base, recoveryRate: safeRatio(base.netRevenue, base.grossRevenue) };
  }, [filteredRows]);

  const cardRows = useMemo(() => filteredRows.filter((r) => r.paymentCategory === "Card"), [filteredRows]);
  const upiRows = useMemo(() => filteredRows.filter((r) => r.paymentCategory === "UPI"), [filteredRows]);

  const activeBanks = useMemo(() => [...new Set(filteredRows.map((r) => r.bankName))], [filteredRows]);
  const activeCardBanks = useMemo(() => activeBanks.filter((b) => cardRows.some((r) => r.bankName === b)), [activeBanks, cardRows]);
  const activeUpiPartners = useMemo(() => activeBanks.filter((b) => upiRows.some((r) => r.bankName === b)), [activeBanks, upiRows]);

  const universalData = useMemo(() => getMonthlyReferenceValue(rows, "universalTransactions"), [rows]);
  const admitsData = useMemo(() => getMonthlyReferenceValue(rows, "admits"), [rows]);
  const universalTicketRevenueData = useMemo(() => getMonthlyReferenceValue(rows, "universalTicketRevenue"), [rows]);
  const universalTotalRevenueData = useMemo(() => getMonthlyReferenceValue(rows, "universalTotalRevenue"), [rows]);

  const activeAdmitsTotal = useMemo(() => {
    const activeMonths = [...new Set(filteredRows.map((row) => row.monthKey).filter((key) => key !== "Unknown"))];
    return activeMonths.reduce((sum, key) => sum + (admitsData.map.get(key) || 0), 0);
  }, [filteredRows, admitsData]);

  const bankAdmits = useMemo(() => filteredRows.reduce((sum, r) => sum + (r.totalTickets || 0), 0), [filteredRows]);

  const atpAvtSourceRows =
    paymentCategoryFilter === "upi" ? upiRows : paymentCategoryFilter === "card" ? cardRows : filteredRows;
  const atpAvtSourceLabel =
    paymentCategoryFilter === "upi" ? "UPI" : paymentCategoryFilter === "card" ? "Bank offers" : "Both";
  const atpAvtKpis = useMemo(() => computeKpis(atpAvtSourceRows), [atpAvtSourceRows]);

  const atp = useMemo(
    () => (atpAvtKpis.totalTickets ? atpAvtKpis.grossRevenue / atpAvtKpis.totalTickets : null),
    [atpAvtKpis],
  );

  const avt = useMemo(
    () => (atpAvtKpis.totalTransactions ? atpAvtKpis.grossRevenue / atpAvtKpis.totalTransactions : null),
    [atpAvtKpis],
  );

  console.log("[ATP DEBUG]", { paymentCategoryFilter, atpAvtSourceRowsLength: atpAvtSourceRows.length, atp, avt });

  const universalATP = useMemo(() => {
    const activeMonths = [...new Set(filteredRows.map((r) => r.monthKey).filter((k) => k !== "Unknown"))];
    const validMonths = activeMonths.filter((key) => universalTicketRevenueData.map.has(key) && admitsData.map.has(key));
    const ticketRev = validMonths.reduce((sum, key) => sum + universalTicketRevenueData.map.get(key), 0);
    const admits = validMonths.reduce((sum, key) => sum + admitsData.map.get(key), 0);
    return admits ? ticketRev / admits : null;
  }, [filteredRows, universalTicketRevenueData, admitsData]);

  const universalAVT = useMemo(() => {
    const activeMonths = [...new Set(filteredRows.map((r) => r.monthKey).filter((k) => k !== "Unknown"))];
    const validMonths = activeMonths.filter((key) => universalTotalRevenueData.map.has(key) && universalData.map.has(key));
    const totalRev = validMonths.reduce((sum, key) => sum + universalTotalRevenueData.map.get(key), 0);
    const universalTxns = validMonths.reduce((sum, key) => sum + universalData.map.get(key), 0);
    return universalTxns ? totalRev / universalTxns : null;
  }, [filteredRows, universalTotalRevenueData, universalData]);

  const avtComparison = useMemo(() => computeUpliftOrContribution(avt, universalAVT), [avt, universalAVT]);
  const atpComparison = useMemo(() => computeUpliftOrContribution(atp, universalATP), [atp, universalATP]);
  const admitsComparison = useMemo(
    () => computeUpliftOrContribution(bankAdmits, activeAdmitsTotal),
    [bankAdmits, activeAdmitsTotal],
  );

  const groupARows = useMemo(
    () => filterGroupFiscal(rows, groupABanks, groupAFy, groupAMonths),
    [rows, groupABanks, groupAFy, groupAMonths],
  );
  const groupBRows = useMemo(
    () => filterGroupFiscal(rows, groupBBanks, groupBFy, groupBMonths),
    [rows, groupBBanks, groupBFy, groupBMonths],
  );
  const groupAKpis = useMemo(() => computeKpis(groupARows), [groupARows]);
  const groupBKpis = useMemo(() => computeKpis(groupBRows), [groupBRows]);
  const groupADiscountRate = safeRatio(groupAKpis.totalDiscount, groupAKpis.grossRevenue);
  const groupBDiscountRate = safeRatio(groupBKpis.totalDiscount, groupBKpis.grossRevenue);
  const groupACardCount = new Set(groupARows.filter((r) => r.paymentCategory === "Card").map((r) => r.bankName)).size;
  const groupAUpiCount = new Set(groupARows.filter((r) => r.paymentCategory === "UPI").map((r) => r.bankName)).size;
  const groupBCardCount = new Set(groupBRows.filter((r) => r.paymentCategory === "Card").map((r) => r.bankName)).size;
  const groupBUpiCount = new Set(groupBRows.filter((r) => r.paymentCategory === "UPI").map((r) => r.bankName)).size;
  const groupABankBreakdown = useMemo(() => aggregateGroupBankBreakdown(groupARows), [groupARows]);
  const groupBBankBreakdown = useMemo(() => aggregateGroupBankBreakdown(groupBRows), [groupBRows]);
  const groupAAdmits = useMemo(() => groupARows.reduce((sum, r) => sum + (r.totalTickets || 0), 0), [groupARows]);
  const groupBAdmits = useMemo(() => groupBRows.reduce((sum, r) => sum + (r.totalTickets || 0), 0), [groupBRows]);

  const comparisonKpis = useMemo(() => {
    if (comparisonMode === "none") return null;
    const periods = getComparisonPeriods(comparisonMode, latestDataDate);
    if (!periods) return null;
    const { currentStart, currentEnd, priorStart, priorEnd } = periods;

    const currentRows = bankOfferFilteredRows.filter((row) => row.date && row.date >= currentStart && row.date <= currentEnd);
    const priorRows = bankOfferFilteredRows.filter((row) => row.date && row.date >= priorStart && row.date <= priorEnd);
    const atpAvtCategory = paymentCategoryFilter === "upi" ? "UPI" : paymentCategoryFilter === "card" ? "Card" : null;
    const currentAtpAvtRows = atpAvtCategory ? currentRows.filter((row) => row.paymentCategory === atpAvtCategory) : currentRows;
    const priorAtpAvtRows = atpAvtCategory ? priorRows.filter((row) => row.paymentCategory === atpAvtCategory) : priorRows;

    const current = computeKpis(currentRows);
    const prior = computeKpis(priorRows);
    const currentAtpAvtKpis = computeKpis(currentAtpAvtRows);
    const priorAtpAvtKpis = computeKpis(priorAtpAvtRows);
    const atpFrom = (k) => (k.totalTickets ? k.grossRevenue / k.totalTickets : null);
    const avtFrom = (k) => (k.totalTransactions ? k.grossRevenue / k.totalTransactions : null);
    const bankAdmitsFrom = (rowsList) => rowsList.reduce((sum, r) => sum + (r.totalTickets || 0), 0);

    return {
      current: {
        ...current,
        recoveryRate: current.grossRevenue ? (current.netRevenue / current.grossRevenue) * 100 : 0,
        atp: atpFrom(currentAtpAvtKpis),
        avt: avtFrom(currentAtpAvtKpis),
        bankAdmits: bankAdmitsFrom(currentRows),
      },
      prior: {
        ...prior,
        recoveryRate: prior.grossRevenue ? (prior.netRevenue / prior.grossRevenue) * 100 : 0,
        atp: atpFrom(priorAtpAvtKpis),
        avt: avtFrom(priorAtpAvtKpis),
        bankAdmits: bankAdmitsFrom(priorRows),
      },
    };
  }, [comparisonMode, latestDataDate, bankOfferFilteredRows, paymentCategoryFilter]);
  const offersByBank = useMemo(() => {
    const grouped = new Map();

    cardRows.forEach((row) => {
      const canonicalName = OFFER_ALIAS_MAP[row.offerName] || normalizeOfferChannel(row.offerName);
      const bankEntry = grouped.get(row.bankName) || {
        bankName: row.bankName,
        offers: new Map(),
      };

      const offerEntry = bankEntry.offers.get(canonicalName) || {
        offerName: canonicalName,
        startDate: row.date,
        endDate: row.date,
        bankDiscount: 0,
        pvrDiscount: 0,
      };

      if (row.date) {
        if (!offerEntry.startDate || row.date < offerEntry.startDate) offerEntry.startDate = row.date;
        if (!offerEntry.endDate || row.date > offerEntry.endDate) offerEntry.endDate = row.date;
      }

      offerEntry.bankDiscount += row.bankContribution;
      offerEntry.pvrDiscount += row.inoxContribution;

      bankEntry.offers.set(canonicalName, offerEntry);
      grouped.set(row.bankName, bankEntry);
    });

    return [...grouped.values()]
      .map((bankEntry) => {
        const offers = [...bankEntry.offers.values()]
          .map((offer) => ({ ...offer, startLabel: formatIsoDate(offer.startDate), endLabel: formatIsoDate(offer.endDate) }))
          .sort((left, right) => (left.startDate?.getTime() || 0) - (right.startDate?.getTime() || 0));
        return { bankName: bankEntry.bankName, offerCount: offers.length, offers };
      })
      .sort((left, right) => right.offerCount - left.offerCount || left.bankName.localeCompare(right.bankName));
  }, [cardRows]);
  const totalOfferCountByBank = useMemo(
    () => offersByBank.reduce((sum, bank) => sum + bank.offerCount, 0),
    [offersByBank],
  );

  const upiOffersByPartner = useMemo(() => {
    const grouped = new Map();

    upiRows.forEach((row) => {
      const canonicalName = OFFER_ALIAS_MAP[row.offerName] || normalizeOfferChannel(row.offerName);
      const partnerEntry = grouped.get(row.bankName) || {
        bankName: row.bankName,
        offers: new Map(),
      };

      const offerEntry = partnerEntry.offers.get(canonicalName) || {
        offerName: canonicalName,
        startDate: row.date,
        endDate: row.date,
        bankDiscount: 0,
        pvrDiscount: 0,
      };

      if (row.date) {
        if (!offerEntry.startDate || row.date < offerEntry.startDate) offerEntry.startDate = row.date;
        if (!offerEntry.endDate || row.date > offerEntry.endDate) offerEntry.endDate = row.date;
      }

      offerEntry.bankDiscount += row.bankContribution;
      offerEntry.pvrDiscount += row.inoxContribution;

      partnerEntry.offers.set(canonicalName, offerEntry);
      grouped.set(row.bankName, partnerEntry);
    });

    return [...grouped.values()]
      .map((partnerEntry) => {
        const offers = [...partnerEntry.offers.values()]
          .map((offer) => ({ ...offer, startLabel: formatIsoDate(offer.startDate), endLabel: formatIsoDate(offer.endDate) }))
          .sort((left, right) => (left.startDate?.getTime() || 0) - (right.startDate?.getTime() || 0));
        return { bankName: partnerEntry.bankName, offerCount: offers.length, offers };
      })
      .sort((left, right) => right.offerCount - left.offerCount || left.bankName.localeCompare(right.bankName));
  }, [upiRows]);
  const totalOfferCountByUpiPartner = useMemo(
    () => upiOffersByPartner.reduce((sum, partner) => sum + partner.offerCount, 0),
    [upiOffersByPartner],
  );

  const totalOfferCount = totalOfferCountByBank + totalOfferCountByUpiPartner;

  const bankRows = useMemo(() => aggregateBanks(filteredRows), [filteredRows]);
  const offerRows = useMemo(() => aggregateOffers(filteredRows), [filteredRows]);

  const rankScopeRows = useMemo(() => {
    return rows.filter((row) => {
      const matchesDate = fyFilter.includes(row.fiscalYear) && (row.date ? monthFilter.includes(MONTH_NAMES[row.date.getMonth()]) : false);
      const matchesCategory =
        paymentCategoryFilter === "all" || row.paymentCategory === (paymentCategoryFilter === "card" ? "Card" : "UPI");
      return matchesDate && matchesCategory;
    });
  }, [rows, fyFilter, monthFilter, paymentCategoryFilter]);

  const globalBankRows = useMemo(() => aggregateBanks(rankScopeRows), [rankScopeRows]);
  const globalBankRankMap = useMemo(() => {
    const sorted = [...globalBankRows].sort((a, b) => b.totalRevenue - a.totalRevenue);
    const map = new Map();
    sorted.forEach((b, index) => map.set(b.bankName, index + 1));
    return map;
  }, [globalBankRows]);

  const sortedBankRows = useMemo(() => {
    if (!bankSortKey) return bankRows;
    const sorted = [...bankRows].sort((a, b) => a[bankSortKey] - b[bankSortKey]);
    return bankSortDir === "desc" ? sorted.reverse() : sorted;
  }, [bankRows, bankSortKey, bankSortDir]);

  const sortedOfferRows = useMemo(() => {
    if (!offerSortKey) return offerRows;
    const sorted = [...offerRows].sort((a, b) => a[offerSortKey] - b[offerSortKey]);
    return offerSortDir === "desc" ? sorted.reverse() : sorted;
  }, [offerRows, offerSortKey, offerSortDir]);
  const categoryDateOfferScopedRows = useMemo(() => {
    return rows.filter((row) => {
      const matchesDate = fyFilter.includes(row.fiscalYear) && (row.date ? monthFilter.includes(MONTH_NAMES[row.date.getMonth()]) : false);
      const matchesOffer = offerFilter.includes(canonicalOfferName(row.offerName));
      const matchesCategory =
        paymentCategoryFilter === "all" || row.paymentCategory === (paymentCategoryFilter === "card" ? "Card" : "UPI");
      return matchesDate && matchesOffer && matchesCategory;
    });
  }, [rows, fyFilter, monthFilter, offerFilter, paymentCategoryFilter]);

  const bankActiveRangeMap = useMemo(() => {
    const map = new Map();
    categoryDateOfferScopedRows.forEach((row) => {
      if (!row.date) return;
      const existing = map.get(row.bankName);
      if (!existing) {
        map.set(row.bankName, { start: row.date, end: row.date });
      } else {
        if (row.date < existing.start) existing.start = row.date;
        if (row.date > existing.end) existing.end = row.date;
      }
    });
    return map;
  }, [categoryDateOfferScopedRows]);

  const monthlySeries = useMemo(
    () => aggregateMonthlySeries(categoryDateOfferScopedRows, bankFilter),
    [categoryDateOfferScopedRows, bankFilter],
  );
  const seasonalData = useMemo(
    () => aggregateSeasonalByYear(categoryDateOfferScopedRows, bankFilter, seasonalMetric),
    [categoryDateOfferScopedRows, bankFilter, seasonalMetric],
  );
  const seasonalYears = useMemo(() => {
    const years = new Set();
    seasonalData.forEach((entry) => {
      Object.keys(entry).forEach((key) => {
        if (key !== "month" && key !== "monthLabel") years.add(key);
      });
    });
    return [...years].sort();
  }, [seasonalData]);
  const yearlyData = useMemo(
    () => aggregateYearlyTotals(categoryDateOfferScopedRows, bankFilter),
    [categoryDateOfferScopedRows, bankFilter],
  );
  const discountData = useMemo(() => {
    const grouped = new Map();

    filteredRows.forEach((row) => {
      const current = grouped.get(row.bankName) || {
        bankName: row.bankName,
        bankDiscount: 0,
        pvrDiscount: 0,
        totalDiscount: 0,
      };

      current.bankDiscount += row.bankContribution;
      current.pvrDiscount += row.inoxContribution;
      current.totalDiscount += row.bankContribution + row.inoxContribution;
      grouped.set(row.bankName, current);
    });

    grouped.forEach((current) => {
      current.bankPercent = current.totalDiscount ? (current.bankDiscount / current.totalDiscount) * 100 : 0;
      current.pvrPercent = current.totalDiscount ? (current.pvrDiscount / current.totalDiscount) * 100 : 0;
    });

    const sorted = [...grouped.values()].sort((left, right) => right.totalDiscount - left.totalDiscount);
    return sorted;
  }, [filteredRows]);
  const discountByBankName = useMemo(() => new Map(discountData.map((entry) => [entry.bankName, entry])), [discountData]);
  const channelData = useMemo(() => aggregateChannelRevenue(filteredRows), [filteredRows]);
  const ticketVsFnbData = useMemo(() => {
    const ticket = filteredRows.reduce((sum, r) => sum + r.ticketRevenue, 0);
    const fnb = filteredRows.reduce((sum, r) => sum + r.fnbRevenue, 0);
    return [{ name: "Tickets", value: ticket }, { name: "F&B", value: fnb }];
  }, [filteredRows]);
  const overallSplit = useMemo(() => {
    const totalBank = discountData.reduce((sum, d) => sum + d.bankDiscount, 0);
    const totalPvr = discountData.reduce((sum, d) => sum + d.pvrDiscount, 0);
    const total = totalBank + totalPvr;
    return {
      totalBank,
      totalPvr,
      bankPercent: total ? (totalBank / total) * 100 : 0,
      pvrPercent: total ? (totalPvr / total) * 100 : 0,
    };
  }, [discountData]);

  function formatLegendValue(value) {
    const amount = Number(value || 0);
    if (Math.abs(amount) >= 10000000) return `${(amount / 10000000).toFixed(1)} Cr`;
    if (Math.abs(amount) >= 100000) return `${(amount / 100000).toFixed(1)} L`;
    return formatInLakh(amount);
  }

  function toggleSelectedBank(bankName) {
    setSelectedBank((current) => (current === bankName ? null : bankName));
  }

  function toggleExpandedOfferBank(bankName) {
    setExpandedOfferBank((current) => (current === bankName ? null : bankName));
  }

  function toggleExpandedOfferUpiPartner(partnerName) {
    setExpandedOfferUpiPartner((current) => (current === partnerName ? null : partnerName));
  }

  function applyQuickPeriod(key, latestDate) {
    if (!latestDate) return;
    const fy = getFiscalYearLabel(latestDate);
    const monthName = MONTH_NAMES[latestDate.getMonth()];
    const fyStartYear = Number(fy.split("-")[0]) + 2000;

    if (key === "thisMonth") {
      setFyFilter([fy]);
      setMonthFilter([monthName]);
    } else if (key === "lastMonth") {
      const prior = new Date(latestDate.getFullYear(), latestDate.getMonth() - 1, 1);
      setFyFilter([getFiscalYearLabel(prior)]);
      setMonthFilter([MONTH_NAMES[prior.getMonth()]]);
    } else if (key === "qtd") {
      const quarterStart = getFiscalQuarterStart(latestDate);
      const months = [];
      for (let d = new Date(quarterStart); d <= latestDate; d.setMonth(d.getMonth() + 1)) {
        months.push(MONTH_NAMES[d.getMonth()]);
      }
      setFyFilter([fy]);
      setMonthFilter(months);
    } else if (key === "fytd") {
      const fyStart = new Date(fyStartYear, 3, 1);
      const months = [];
      for (let d = new Date(fyStart); d <= latestDate; d.setMonth(d.getMonth() + 1)) {
        months.push(MONTH_NAMES[d.getMonth()]);
      }
      setFyFilter([fy]);
      setMonthFilter(months);
    }
  }

  const selectedBankRow = useMemo(() => bankRows.find((bank) => bank.bankName === selectedBank) || null, [bankRows, selectedBank]);
  const selectedBankOffersEntry = useMemo(() => offersByBank.find((entry) => entry.bankName === selectedBank) || null, [offersByBank, selectedBank]);
  const selectedBankDiscountEntry = useMemo(() => discountData.find((entry) => entry.bankName === selectedBank) || null, [discountData, selectedBank]);
  const anyModalOpen = Boolean(selectedBank || selectedOffer || showOffersByBank || showOffersByUpi);

  function buildPairInsight(priorYear, currentYear, relevantRowsForYear, admitsMap) {
    const currentRows = relevantRowsForYear(currentYear);
    const priorRows = relevantRowsForYear(priorYear);

    const currentTotals = computeRowTotals(currentRows);
    const priorTotals = computeRowTotals(priorRows);

    const currentTotal = currentTotals.revenue;
    const priorTotal = priorTotals.revenue;
    const totalDeltaAbs = currentTotal - priorTotal;
    const totalDeltaPercent = priorTotal ? (totalDeltaAbs / priorTotal) * 100 : null;

    const txnDeltaPercent = priorTotals.transactions ? ((currentTotals.transactions - priorTotals.transactions) / priorTotals.transactions) * 100 : null;

    const currentDiscountRate = currentTotals.revenue ? (currentTotals.discount / currentTotals.revenue) * 100 : null;
    const priorDiscountRate = priorTotals.revenue ? (priorTotals.discount / priorTotals.revenue) * 100 : null;

    const discountRateNote = computeDiscountRateNote(currentDiscountRate, priorDiscountRate);
    const volumeValueNote = computeVolumeValueNote(totalDeltaPercent, txnDeltaPercent);

    const currentAdmits = sumAdmitsForMonths(currentRows.map((r) => r.monthKey), admitsMap);
    const priorAdmits = sumAdmitsForMonths(priorRows.map((r) => r.monthKey), admitsMap);
    const admitsDeltaPercent = priorAdmits ? ((currentAdmits - priorAdmits) / priorAdmits) * 100 : null;
    const admitsNote = computeAdmitsNote(admitsDeltaPercent);

    const currentATV = currentAdmits ? currentTotals.ticketRevenue / currentAdmits : null;
    const priorATV = priorAdmits ? priorTotals.ticketRevenue / priorAdmits : null;
    const atvDeltaPercent = priorATV ? ((currentATV - priorATV) / priorATV) * 100 : null;
    const atvNote = computeATVNote(atvDeltaPercent);

    const headline = generateHeadlineInsight({ revenueDeltaPercent: totalDeltaPercent, admitsDeltaPercent, atvDeltaPercent });

    const allBanksInPair = new Set([...priorRows.map((r) => r.bankName), ...currentRows.map((r) => r.bankName)]);

    const bankBreakdown = [...allBanksInPair]
      .map((bankName) => {
        const priorRevenue = Number(
          priorRows
            .filter((r) => r.bankName === bankName)
            .reduce((sum, r) => sum + r.transactionTotal, 0)
            .toFixed(5),
        );
        const currentRevenue = Number(
          currentRows
            .filter((r) => r.bankName === bankName)
            .reduce((sum, r) => sum + r.transactionTotal, 0)
            .toFixed(5),
        );
        const deltaAbs = currentRevenue - priorRevenue;
        const contributionPercent = totalDeltaAbs ? (deltaAbs / totalDeltaAbs) * 100 : null;
        return { bankName, priorRevenue, currentRevenue, deltaAbs, contributionPercent };
      })
      .sort((a, b) => Math.abs(b.deltaAbs) - Math.abs(a.deltaAbs));

    const topBank = bankBreakdown.length ? bankBreakdown[0].bankName : null;

    return {
      currentTotal,
      priorTotal,
      totalDeltaPercent,
      currentTransactions: currentTotals.transactions,
      priorTransactions: priorTotals.transactions,
      txnDeltaPercent,
      currentDiscountRate,
      priorDiscountRate,
      discountRateNote,
      volumeValueNote,
      currentAdmits,
      priorAdmits,
      admitsDeltaPercent,
      admitsNote,
      currentATV,
      priorATV,
      atvDeltaPercent,
      atvNote,
      headline,
      bankBreakdown,
      topBank,
    };
  }

  const inferencePairs = useMemo(() => {
    if (!selectedSeasonalPoint) return [];
    const { month } = selectedSeasonalPoint;

    const relevantRowsForMonth = (fiscalYearLabel) => {
      const calYear = calendarYearForFiscalMonth(fiscalYearLabel, month);
      return categoryDateOfferScopedRows.filter(
        (r) =>
          r.monthKey !== "Unknown" &&
          Number(r.monthKey.split("-")[0]) === month &&
          Number(r.monthKey.split("-")[1]) === calYear &&
          (!bankFilter.length || bankFilter.includes(r.bankName)),
      );
    };

    const yearsWithData = seasonalYears
      .filter((fy) => relevantRowsForMonth(fy).length > 0)
      .sort((a, b) => Number(a.split("-")[0]) - Number(b.split("-")[0]));

    const pairs = [];
    for (let i = 1; i < yearsWithData.length; i++) {
      const priorYear = yearsWithData[i - 1];
      const currentYear = yearsWithData[i];
      pairs.push({ priorYear, currentYear, ...buildPairInsight(priorYear, currentYear, relevantRowsForMonth, admitsData.map) });
    }
    return pairs;
  }, [selectedSeasonalPoint, seasonalYears, categoryDateOfferScopedRows, bankFilter, admitsData]);

  const monthlyInsight = useMemo(() => {
    if (!selectedMonthlyPoint) return null;
    return buildAdjacentMonthInsight(selectedMonthlyPoint.bankName, selectedMonthlyPoint.monthKey, categoryDateOfferScopedRows, admitsData.map);
  }, [selectedMonthlyPoint, categoryDateOfferScopedRows, admitsData]);

  const monthlyYoY = useMemo(() => {
    if (!selectedMonthlyPoint) return null;
    return buildBankYearOverYear(selectedMonthlyPoint.bankName, selectedMonthlyPoint.monthKey, categoryDateOfferScopedRows);
  }, [selectedMonthlyPoint, categoryDateOfferScopedRows]);

  const bankRank = useMemo(() => {
    if (!selectedMonthlyPoint) return null;
    const sorted = [...bankRows].sort((a, b) => b.totalRevenue - a.totalRevenue);
    const rank = sorted.findIndex((b) => b.bankName === selectedMonthlyPoint.bankName) + 1;
    return rank > 0 ? rank : null;
  }, [bankRows, selectedMonthlyPoint]);

  const momAdjacentMonthInsight = useMemo(() => {
    if (!selectedSeasonalPoint) return null;
    const { month } = selectedSeasonalPoint;
    const relevantRowsForMonth = (fiscalYearLabel) => {
      const calYear = calendarYearForFiscalMonth(fiscalYearLabel, month);
      return categoryDateOfferScopedRows.filter(
        (r) =>
          r.monthKey !== "Unknown" &&
          Number(r.monthKey.split("-")[0]) === month &&
          Number(r.monthKey.split("-")[1]) === calYear &&
          (!bankFilter.length || bankFilter.includes(r.bankName)),
      );
    };
    const yearsWithData = seasonalYears
      .filter((fy) => relevantRowsForMonth(fy).length > 0)
      .sort((a, b) => Number(a.split("-")[0]) - Number(b.split("-")[0]));
    if (!yearsWithData.length) return null;
    const currentYear = yearsWithData[yearsWithData.length - 1];
    const calYear = calendarYearForFiscalMonth(currentYear, month);
    const monthKey = `${String(month).padStart(2, "0")}-${calYear}`;
    return buildAggregateAdjacentMonthInsight(bankFilter, monthKey, categoryDateOfferScopedRows, admitsData.map);
  }, [selectedSeasonalPoint, seasonalYears, categoryDateOfferScopedRows, bankFilter, admitsData]);

  const yearInsight = useMemo(() => {
    if (!selectedYearPoint) return null;
    const { year } = selectedYearPoint;

    const relevantRowsForYear = (y) =>
      categoryDateOfferScopedRows.filter((r) => r.fiscalYear === y && (!bankFilter.length || bankFilter.includes(r.bankName)));

    const candidateYears = [
      ...new Set(categoryDateOfferScopedRows.filter((r) => r.fiscalYear && r.fiscalYear !== "Unknown").map((r) => r.fiscalYear)),
    ].sort((a, b) => Number(a.split("-")[0]) - Number(b.split("-")[0]));
    const yearsWithData = candidateYears.filter((y) => relevantRowsForYear(y).length > 0);
    const idx = yearsWithData.indexOf(year);
    const priorYear = idx > 0 ? yearsWithData[idx - 1] : null;

    if (!priorYear) {
      const currentRows = relevantRowsForYear(year);
      const currentTotals = computeRowTotals(currentRows);
      const currentDiscountRate = currentTotals.revenue ? (currentTotals.discount / currentTotals.revenue) * 100 : null;
      const currentAdmits = sumAdmitsForMonths(currentRows.map((r) => r.monthKey), admitsData.map);
      const currentATV = currentAdmits ? currentTotals.ticketRevenue / currentAdmits : null;
      return { year, hasPrior: false, currentTotals, currentDiscountRate, currentAdmits, currentATV };
    }

    return { year, priorYear, hasPrior: true, ...buildPairInsight(priorYear, year, relevantRowsForYear, admitsData.map) };
  }, [selectedYearPoint, categoryDateOfferScopedRows, bankFilter, admitsData]);

  function handleViewDetails() {
    if (trendMode === "monthly") {
      if (selectedMonthlyPoint || !monthlySeries.length || !bankFilter.length) return;
      const latestRow = monthlySeries[monthlySeries.length - 1];
      const bankName = bankFilter.find((bank) => typeof latestRow[bank] === "number") || bankFilter[0];
      setSelectedMonthlyPoint({ bankName, monthKey: latestRow.monthKey, monthLabel: formatMonthKeyLabel(latestRow.monthKey) });
    } else if (trendMode === "mom") {
      if (selectedSeasonalPoint || !seasonalData.length) return;
      const latestRow = seasonalData[seasonalData.length - 1];
      setSelectedSeasonalPoint({ month: latestRow.month, monthLabel: latestRow.monthLabel });
    } else if (trendMode === "yoy") {
      if (selectedYearPoint || !yearlyData.length) return;
      const latestRow = yearlyData[yearlyData.length - 1];
      setSelectedYearPoint({ year: latestRow.year });
    }
  }

  return (
    <div className="h-full overflow-hidden bg-appBg text-textMain">
      <div className="mx-auto flex h-full max-w-[1800px] flex-col gap-5 overflow-y-auto px-4 py-4 scrollbar-thin sm:px-6 lg:px-8">
        {/* Row 1 — Brand + Quick Period + Bank Offer Penetration */}
        <div className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-white/60 bg-white/80 px-4 py-2.5 shadow-soft backdrop-blur">
          <div className="flex items-center gap-3">
            <img src={pvrInoxLogo} alt="PVR INOX" className="h-7 w-auto" />
            <h1 className="font-display text-base font-bold text-textMain sm:text-lg">BANK OFFERS PERFORMANCE</h1>
            <button
              type="button"
              onClick={() => setShowComparisonModal(true)}
              className="flex items-center gap-1.5 rounded-full border-2 border-accentBlue bg-accentBlue px-4 py-1.5 text-xs font-extrabold uppercase tracking-wide text-white shadow-sm transition hover:bg-blue-700"
            >
              ⇄ Comparison Module
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {QUICK_PERIODS.map((preset) => (
              <button
                key={preset.key}
                type="button"
                disabled={!latestDataDate}
                onClick={() => applyQuickPeriod(preset.key, latestDataDate)}
                className={`rounded-full border px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide transition disabled:cursor-not-allowed disabled:opacity-40 ${
                  activeQuickPeriodKey === preset.key
                    ? "border-accentBlue bg-accentBlue text-white"
                    : "border-borderSoft bg-white text-textMuted hover:bg-slate-50"
                }`}
              >
                {preset.label}
              </button>
            ))}
            <button
              type="button"
              onClick={() => {
                setFyFilter([]);
                setMonthFilter([]);
              }}
              className={`rounded-full border px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide transition ${
                activeQuickPeriodKey === "clear"
                  ? "border-accentBlue bg-accentBlue text-white"
                  : "border-borderSoft bg-white text-textMuted hover:bg-slate-50"
              }`}
            >
              Clear
            </button>
          </div>

          <div className="ml-auto flex items-center gap-2">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(event) => {
                const [file] = [...(event.target.files || [])];
                if (file) handleFileChange(file);
              }}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              onDragOver={(event) => {
                event.preventDefault();
                setDragActive(true);
              }}
              onDragLeave={() => setDragActive(false)}
              onDrop={(event) => {
                event.preventDefault();
                setDragActive(false);
                const [file] = [...(event.dataTransfer.files || [])];
                if (file) handleFileChange(file);
              }}
              title={fileName || "Upload Excel Performance File"}
              className={`flex items-center gap-2 rounded-full border px-3 py-1.5 text-xs font-semibold transition ${
                dragActive ? "border-accentBlue bg-blue-50 text-accentBlue" : "border-borderSoft bg-white text-textMain hover:bg-slate-50"
              }`}
            >
              📎 <span className="max-w-[140px] truncate">{fileName || "Upload File"}</span>
            </button>
            {fileName ? (
              <button
                type="button"
                onClick={handleRemoveFile}
                className="rounded-full border border-rose-200 bg-rose-50 px-2 py-1 text-[10px] font-bold text-rose-600 hover:bg-rose-100"
              >
                ✕ Remove
              </button>
            ) : null}

            <input
              ref={upiFileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(event) => {
                const [file] = [...(event.target.files || [])];
                if (file) handleUpiFileChange(file);
              }}
            />
            <button
              type="button"
              onClick={() => upiFileInputRef.current?.click()}
              onDragOver={(event) => {
                event.preventDefault();
                setUpiDragActive(true);
              }}
              onDragLeave={() => setUpiDragActive(false)}
              onDrop={(event) => {
                event.preventDefault();
                setUpiDragActive(false);
                const [file] = [...(event.dataTransfer.files || [])];
                if (file) handleUpiFileChange(file);
              }}
              title={upiFileName || "Upload UPI Partner File"}
              className={`flex items-center gap-2 rounded-full border px-3 py-1.5 text-xs font-semibold transition ${
                upiDragActive ? "border-accentBlue bg-blue-50 text-accentBlue" : "border-borderSoft bg-white text-textMain hover:bg-slate-50"
              }`}
            >
              📎 <span className="max-w-[140px] truncate">{upiFileName || "Upload UPI Data"}</span>
            </button>
            {upiFileName ? (
              <button
                type="button"
                onClick={handleRemoveUpiFile}
                className="rounded-full border border-rose-200 bg-rose-50 px-2 py-1 text-[10px] font-bold text-rose-600 hover:bg-rose-100"
              >
                ✕ Remove
              </button>
            ) : null}
          </div>
        </div>
        {error ? <p className="text-right text-[10px] font-semibold text-rose-600">{error}</p> : null}
        {upiError ? <p className="text-right text-[10px] font-semibold text-rose-600">{upiError}</p> : null}

        {missingColumns.length || universalData.inconsistent.size > 0 ? (
          <div className="rounded-2xl border border-amber-200 bg-amber-50 px-4 py-2 text-xs font-semibold text-amber-700">
            {missingColumns.length ? <p>Missing columns were handled as zero values: {missingColumns.join(", ")}</p> : null}
            {universalData.inconsistent.size > 0 ? (
              <p>Inconsistent universal totals found for: {[...universalData.inconsistent].map(formatMonthKeyLabel).join(", ")}</p>
            ) : null}
          </div>
        ) : null}

        {/* Row 2 (was Row 3) — sticky, centered Date/Bank/Offer filter bar */}
        <div
          className={
            anyModalOpen
              ? "hidden"
              : "sticky top-0 z-50 -mx-4 flex justify-center border-b border-borderSoft bg-white/95 px-4 py-2 shadow-sm backdrop-blur sm:-mx-6 sm:px-6 lg:-mx-8 lg:px-8"
          }
        >
          <div className="flex w-full items-center gap-3">
            <div className="flex flex-shrink-0 rounded-full border border-borderSoft bg-white p-1 shadow-sm">
              {[
                { key: "card", label: "Bank" },
                { key: "upi", label: "UPI" },
                { key: "all", label: "Both" },
              ].map((opt) => (
                <button
                  key={opt.key}
                  type="button"
                  onClick={() => setPaymentCategoryFilter(opt.key)}
                  className={`rounded-full px-3 py-1.5 text-xs font-bold uppercase tracking-wide transition ${
                    paymentCategoryFilter === opt.key ? "bg-accentBlue text-white" : "text-textMuted hover:bg-slate-50"
                  }`}
                >
                  {opt.label}
                </button>
              ))}
            </div>
            <div className="flex flex-1 flex-wrap items-center justify-center gap-3">
              <MultiSelectDropdown
                options={fiscalYears}
                selected={fyFilter}
                onToggle={(fy) => setFyFilter((current) => (current.includes(fy) ? current.filter((item) => item !== fy) : [...current, fy]))}
                onClear={() => setFyFilter([])}
                label="FY"
                showSelectedNames
              />
              <MultiSelectDropdown
                options={fiscalMonths}
                selected={monthFilter}
                onToggle={(month) => setMonthFilter((current) => (current.includes(month) ? current.filter((item) => item !== month) : [...current, month]))}
                onClear={() => setMonthFilter([])}
                label="month"
                showSelectedNames
              />
              <MultiSelectDropdown
                options={banks}
                selected={bankFilter}
                onToggle={(bank) => {
                  setBankFilter((current) => (current.includes(bank) ? current.filter((item) => item !== bank) : [...current, bank]));
                }}
                onClear={() => {
                  console.log("[Clear audit] Bank filter cleared, new value:", []);
                  setBankFilter([]);
                }}
                label="bank"
              />
              <MultiSelectDropdown
                options={offers}
                selected={offerFilter}
                onToggle={(offer) => {
                  setOfferFilter((current) => (current.includes(offer) ? current.filter((item) => item !== offer) : [...current, offer]));
                }}
                onClear={() => {
                  console.log("[Clear audit] Offer filter cleared, new value:", []);
                  setOfferFilter([]);
                }}
                label="offer"
              />
            </div>
          </div>
        </div>

        {/* Row 3 (was Row 2) — Total Banks / Total Offers / Info Section, 4 equal compact boxes */}
        <div className="grid grid-cols-4 gap-3">
          <div className="flex h-[90px] flex-col justify-center rounded-2xl border border-white/60 bg-white/90 p-3 shadow-soft">
            <p className="text-[10px] font-bold uppercase tracking-[0.2em] text-slate-500">Total Banks / UPI</p>
            <p className="mt-1 text-2xl font-extrabold text-textMain">
              {formatInteger(activeCardBanks.length)} / {formatInteger(activeUpiPartners.length)}
            </p>
          </div>
          <div className="flex h-[90px] flex-col justify-center rounded-2xl border border-white/60 bg-white/90 p-3 shadow-soft">
            <p className="text-[10px] font-bold uppercase tracking-[0.2em] text-slate-500">Total Offers</p>
            <p className="mt-1 text-2xl font-extrabold text-textMain">{formatInteger(totalOfferCount)}</p>
            <p className="text-xs text-textMuted">
              {formatInteger(totalOfferCountByBank)} Card / {formatInteger(totalOfferCountByUpiPartner)} UPI
            </p>
          </div>
          <div
            role="button"
            tabIndex={0}
            onClick={() => setShowOffersByUpi(true)}
            onKeyDown={(event) => {
              if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                setShowOffersByUpi(true);
              }
            }}
            className="flex h-[90px] cursor-pointer items-center justify-between gap-2 rounded-2xl border border-white/60 bg-white/90 p-3 text-left shadow-soft"
          >
            <div className="min-w-0">
              <p className="mt-1 truncate text-sm font-bold text-textMain">UPI Partners</p>
              <p className="text-xs text-textMuted">
                {formatInteger(totalOfferCountByUpiPartner)} offers across {formatInteger(upiOffersByPartner.length)} partners
              </p>
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  exportOffersToExcel(upiOffersByPartner, "upi_partner_offers.xlsx");
                }}
                className="mt-1 rounded-full border border-borderSoft bg-white px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-textMuted hover:bg-slate-50"
              >
                ⬇ Download Excel
              </button>
            </div>
            <span className="flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-xl border border-borderSoft bg-slate-50 text-sm font-bold text-textMuted">
              →
            </span>
          </div>
          <div
            role="button"
            tabIndex={0}
            onClick={() => setShowOffersByBank(true)}
            onKeyDown={(event) => {
              if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                setShowOffersByBank(true);
              }
            }}
            className="flex h-[90px] cursor-pointer items-center justify-between gap-2 rounded-2xl border border-white/60 bg-white/90 p-3 text-left shadow-soft"
          >
            <div className="min-w-0">
              <p className="mt-1 truncate text-sm font-bold text-textMain">Bank Partners</p>
              <p className="text-xs text-textMuted">
                {formatInteger(totalOfferCountByBank)} offers across {formatInteger(offersByBank.length)} banks
              </p>
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  exportOffersToExcel(offersByBank, "bank_offers.xlsx");
                }}
                className="mt-1 rounded-full border border-borderSoft bg-white px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-textMuted hover:bg-slate-50"
              >
                ⬇ Download Excel
              </button>
            </div>
            <span className="flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-xl border border-borderSoft bg-slate-50 text-sm font-bold text-textMuted">
              →
            </span>
          </div>
        </div>

        {/* Row 4 — Key Metrics (kept at full size, most generous space) */}
        <section className="flex flex-col gap-4">
          <div className="flex items-center justify-between gap-3">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Key Metrics</p>
            <div className="flex items-center gap-3">
              <div className="inline-flex rounded-full border border-borderSoft bg-white p-1 shadow-sm">
                {COMPARISON_MODES.map((mode) => (
                  <button
                    key={mode.key}
                    type="button"
                    onClick={() => setComparisonMode(mode.key)}
                    className={`rounded-full px-3 py-1.5 text-xs font-bold uppercase tracking-wide transition ${
                      comparisonMode === mode.key ? "bg-accentBlue text-white" : "text-textMuted hover:bg-slate-50"
                    }`}
                  >
                    {mode.label}
                  </button>
                ))}
              </div>
              <p className="text-xs text-textMuted">Current vs previous</p>
            </div>
          </div>
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
            <StatCard
              title="Total Transactions"
              value={formatCountInLakh(kpis.totalTransactions)}
              subtitle={
                <>
                  {formatCountInLakh(cardRows.reduce((s, r) => s + r.discountedTransactions, 0))} Card
                  <br />
                  {formatCountInLakh(upiRows.reduce((s, r) => s + r.discountedTransactions, 0))} UPI
                </>
              }
              color={KPI_COLORS[0]}
              icon="TX"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.totalTransactions, comparisonKpis.prior.totalTransactions), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
            <StatCard
              title="Revenue"
              value={formatInLakh(kpis.grossRevenue)}
              subtitle={
                <>
                  {formatInLakh(cardRows.reduce((s, r) => s + r.transactionTotal, 0))} Card
                  <br />
                  {formatInLakh(upiRows.reduce((s, r) => s + r.transactionTotal, 0))} UPI
                </>
              }
              color={KPI_COLORS[0]}
              icon="GR"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.grossRevenue, comparisonKpis.prior.grossRevenue), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
            <StatCard
              title="ATP"
              value={atp !== null ? formatRupee(atp) : "—"}
              subtitle={
                universalATP !== null ? (
                  <>
                    {atpAvtSourceLabel}: {formatRupee(atp)}
                    <br />
                    Universal: {formatRupee(universalATP)}
                  </>
                ) : (
                  "Add 'Universal Ticket Revenue' column to compare vs. universal"
                )
              }
              color={KPI_COLORS[0]}
              icon="TP"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.atp, comparisonKpis.prior.atp), label: COMPARISON_LABELS[comparisonMode] } : undefined}
              extra={<UpliftOrContributionLine comparison={atpComparison} />}
            />
            <StatCard
              title="AVT"
              value={avt !== null ? formatRupee(avt) : "—"}
              subtitle={
                universalAVT !== null ? (
                  <>
                    {atpAvtSourceLabel}: {formatRupee(avt)}
                    <br />
                    Universal: {formatRupee(universalAVT)}
                  </>
                ) : (
                  "Add 'Universal Total Revenue' column to compare vs. universal"
                )
              }
              color={KPI_COLORS[0]}
              icon="AT"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.avt, comparisonKpis.prior.avt), label: COMPARISON_LABELS[comparisonMode] } : undefined}
              extra={<UpliftOrContributionLine comparison={avtComparison} />}
            />
            <StatCard
              title="Admits"
              value={`${formatCountInLakh(activeAdmitsTotal)} / ${formatCountInLakh(bankAdmits)}`}
              subtitle={
                <>
                  Universal / Bank
                  <br />
                  <UpliftOrContributionLine comparison={admitsComparison} />
                </>
              }
              color={KPI_COLORS[0]}
              icon="AD"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.bankAdmits, comparisonKpis.prior.bankAdmits), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
          </div>
        </section>

        <section className="grid gap-5 xl:grid-cols-[1.65fr_1fr]">
          <div className="flex h-[488px] flex-col overflow-hidden rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <div className="flex items-center justify-between gap-3">
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Bank / UPI Partner Scorecard</p>
                <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Bank / UPI Partner Performance Overview</h2>
              </div>
              <div className="flex items-center gap-3">
                {bankSortKey ? (
                  <button
                    type="button"
                    onClick={() => setBankSortKey(null)}
                    className="text-xs font-bold text-accentBlue hover:underline"
                  >
                    Reset to default order
                  </button>
                ) : null}
                <p className="rounded-full bg-slate-100 px-3 py-1 text-xs font-bold text-slate-600">{formatInteger(bankRows.length)} banks</p>
              </div>
            </div>
            <div className="mt-4 h-[380px] overflow-hidden rounded-3xl border border-borderSoft">
              <div className="h-full overflow-y-scroll overflow-x-auto scrollbar-thin">
                <table className="min-w-full table-fixed divide-y divide-borderSoft text-sm">
                  <thead className="sticky top-0 bg-slate-50/95 backdrop-blur">
                    <tr>
                      <th className="w-[18%] px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">
                        Bank/UPI
                        <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                          ({bankRows.length} banks)
                        </span>
                      </th>
                      <th
                        className="w-[14%] cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                        onClick={() => {
                          const key = "totalTransactions";
                          if (bankSortKey === key) setBankSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                          else {
                            setBankSortKey(key);
                            setBankSortDir("desc");
                          }
                        }}
                      >
                        Total Txns {bankSortKey === "totalTransactions" ? (bankSortDir === "desc" ? "▼" : "▲") : ""}
                        <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                          ({formatCountInLakh(kpis.totalTransactions)})
                        </span>
                      </th>
                      <th
                        className="w-[20%] cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                        onClick={() => {
                          const key = "totalRevenue";
                          if (bankSortKey === key) setBankSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                          else {
                            setBankSortKey(key);
                            setBankSortDir("desc");
                          }
                        }}
                      >
                        Total Revenue {bankSortKey === "totalRevenue" ? (bankSortDir === "desc" ? "▼" : "▲") : ""}
                        <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                          ({formatInLakh(kpis.grossRevenue)})
                        </span>
                      </th>
                      <th
                        className="w-[16%] cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                        onClick={() => {
                          const key = "discountCost";
                          if (bankSortKey === key) setBankSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                          else {
                            setBankSortKey(key);
                            setBankSortDir("desc");
                          }
                        }}
                      >
                        Total Discount {bankSortKey === "discountCost" ? (bankSortDir === "desc" ? "▼" : "▲") : ""}
                        <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                          ({formatInLakh(kpis.totalDiscount)})
                        </span>
                      </th>
                      <th className="w-[32%] px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Discount Split (Bank / PVR)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedBankRows.length ? sortedBankRows.map((bank, index) => {
                      const split = discountByBankName.get(bank.bankName);
                      return (
                        <tr key={bank.bankName} className={`${index % 2 === 0 ? "bg-white" : "bg-slate-50/60"} cursor-pointer transition hover:bg-blue-50/70`} onClick={() => toggleSelectedBank(bank.bankName)}>
                          <td className="px-4 py-3 font-bold text-textMain">
                            {bank.bankName}
                            {bank.paymentCategory === "UPI" ? " (UPI)" : ""}
                          </td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            {formatCountInLakh(bank.totalTransactions)}
                            <span className="ml-1 text-xs font-bold text-textMuted">
                              ({kpis.totalTransactions ? ((bank.totalTransactions / kpis.totalTransactions) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            <span className="whitespace-nowrap">{formatInLakh(bank.totalRevenue)}</span>
                            <span className="mt-0.5 block text-xs font-bold text-textMuted">
                              ({kpis.grossRevenue ? ((bank.totalRevenue / kpis.grossRevenue) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            <span className="whitespace-nowrap">{formatInLakh(bank.discountCost)}</span>
                            <span className="mt-0.5 block text-xs font-bold text-textMuted">
                              ({kpis.totalDiscount ? ((bank.discountCost / kpis.totalDiscount) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 text-xs font-semibold text-textMain">
                            {split ? (
                              <div className="flex flex-col gap-0.5">
                                <span className="text-accentBlue">Bank: {formatInLakh(split.bankDiscount)} ({split.bankPercent.toFixed(0)}%)</span>
                                <span className="text-accentGreen">PVR: {formatInLakh(split.pvrDiscount)} ({split.pvrPercent.toFixed(0)}%)</span>
                              </div>
                            ) : "—"}
                          </td>
                        </tr>
                      );
                    }) : (
                      <tr>
                        <td colSpan="5" className="px-4 py-10 text-center font-semibold text-textMuted">Upload data to view the bank scorecard.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div className="flex h-[488px] flex-col justify-start overflow-hidden rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <p className="text-sm font-bold uppercase tracking-[0.24em] text-textMuted">Discount Split By Partners</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Overall Bank / UPI vs PVR Contribution</h2>
            <p className="mt-1 text-sm font-semibold text-textMuted">
              Total Discount: {formatInLakh(overallSplit.totalBank + overallSplit.totalPvr)}
            </p>
            <div className="mt-6 flex flex-col gap-6">
              <div>
                <div className="flex items-center justify-between text-sm font-bold text-textMain">
                  <span className="text-accentBlue">Bank: {formatInLakh(overallSplit.totalBank)} ({overallSplit.bankPercent.toFixed(0)}%)</span>
                  <span className="text-accentGreen">PVR: {formatInLakh(overallSplit.totalPvr)} ({overallSplit.pvrPercent.toFixed(0)}%)</span>
                </div>
                <div className="mt-3 flex h-8 w-full overflow-hidden rounded-full bg-slate-100">
                  <div className="h-full bg-accentBlue transition-all" style={{ width: `${overallSplit.bankPercent}%` }} />
                  <div className="h-full bg-accentGreen transition-all" style={{ width: `${overallSplit.pvrPercent}%` }} />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="rounded-2xl bg-blue-50 p-4">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-accentBlue">Bank/UPI Contribution</p>
                  <p className="mt-2 text-2xl font-bold text-textMain">{formatInLakh(overallSplit.totalBank)}</p>
                  <p className="mt-1 text-sm font-semibold text-textMuted">{overallSplit.bankPercent.toFixed(1)}% of total discount</p>
                </div>
                <div className="rounded-2xl bg-emerald-50 p-4">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-accentGreen">PVR Contribution</p>
                  <p className="mt-2 text-2xl font-bold text-textMain">{formatInLakh(overallSplit.totalPvr)}</p>
                  <p className="mt-1 text-sm font-semibold text-textMuted">{overallSplit.pvrPercent.toFixed(1)}% of total discount</p>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="grid gap-5 xl:grid-cols-2">
          <div className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Channel Mix</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Revenue by Channel</h2>
            <div className="relative mt-4 h-[340px] overflow-hidden rounded-3xl bg-appBg p-3">
              <span className="absolute right-3 top-3 z-10 rounded-full bg-slate-100 px-3 py-1 text-xs font-bold text-textMuted">
                Total: {formatInLakh(channelData.reduce((sum, c) => sum + c.value, 0))}
              </span>
              {channelData.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart margin={{ top: 30, right: 100, bottom: 30, left: 100 }}>
                    <Pie
                      data={channelData}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      outerRadius={75}
                      label={renderPieLabel}
                      labelLine={{ stroke: "#94a3b8", strokeWidth: 1 }}
                    >
                      {channelData.map((entry) => (
                        <Cell key={entry.name} fill={CHANNEL_COLORS[entry.name] || CHANNEL_FALLBACK_COLOR} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => formatInLakh(value)} contentStyle={{ borderRadius: "18px", borderColor: "#e2e8f0" }} />
                    <Legend />
                  </PieChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Upload data to view revenue by channel.
                </div>
              )}
            </div>
          </div>

          <div className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Revenue Split</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Ticket vs F&B Revenue</h2>
            <div className="relative mt-4 h-[340px] overflow-hidden rounded-3xl bg-appBg p-3">
              <span className="absolute right-3 top-3 z-10 rounded-full bg-slate-100 px-3 py-1 text-xs font-bold text-textMuted">
                Total: {formatInLakh(ticketVsFnbData.reduce((sum, c) => sum + c.value, 0))}
              </span>
              {filteredRows.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart margin={{ top: 30, right: 100, bottom: 30, left: 100 }}>
                    <Pie
                      data={ticketVsFnbData}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      outerRadius={75}
                      label={renderPieLabel}
                      labelLine={{ stroke: "#94a3b8", strokeWidth: 1 }}
                    >
                      {ticketVsFnbData.map((entry) => (
                        <Cell key={entry.name} fill={TICKET_FNB_COLORS[entry.name]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => formatInLakh(value)} contentStyle={{ borderRadius: "18px", borderColor: "#e2e8f0" }} />
                    <Legend />
                  </PieChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Upload data to view ticket vs F&B revenue.
                </div>
              )}
            </div>
          </div>
        </section>

        <section className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Offer Directory</p>
              <h2 className="mt-1 font-display text-2xl font-bold text-textMain">All offers grouped by Bank / UPI Partners</h2>
            </div>
            {offerSortKey ? (
              <button
                type="button"
                onClick={() => setOfferSortKey(null)}
                className="text-xs font-bold text-accentBlue hover:underline"
              >
                Reset to default order
              </button>
            ) : null}
          </div>
          <div className="mt-4 overflow-hidden rounded-3xl border border-borderSoft">
            <div className="h-[320px] overflow-y-scroll scrollbar-thin">
              <table className="min-w-full divide-y divide-borderSoft text-sm">
                <thead className="sticky top-0 bg-slate-50/95 backdrop-blur">
                  <tr>
                    <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">
                      Bank/UPI
                      <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                        ({banks.length} banks)
                      </span>
                    </th>
                    <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">
                      Offer Name
                      <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                        ({offerRows.length} offers)
                      </span>
                    </th>
                    <th
                      className="cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                      onClick={() => {
                        const key = "transactions";
                        if (offerSortKey === key) setOfferSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                        else {
                          setOfferSortKey(key);
                          setOfferSortDir("desc");
                        }
                      }}
                    >
                      Transactions {offerSortKey === "transactions" ? (offerSortDir === "desc" ? "▼" : "▲") : ""}
                      <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                        ({formatCountInLakh(kpis.totalTransactions)})
                      </span>
                    </th>
                    <th
                      className="cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                      onClick={() => {
                        const key = "revenue";
                        if (offerSortKey === key) setOfferSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                        else {
                          setOfferSortKey(key);
                          setOfferSortDir("desc");
                        }
                      }}
                    >
                      Revenue {offerSortKey === "revenue" ? (offerSortDir === "desc" ? "▼" : "▲") : ""}
                      <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                        ({formatInLakh(kpis.grossRevenue)})
                      </span>
                    </th>
                    <th
                      className="cursor-pointer select-none px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted hover:text-textMain"
                      onClick={() => {
                        const key = "discount";
                        if (offerSortKey === key) setOfferSortDir((dir) => (dir === "desc" ? "asc" : "desc"));
                        else {
                          setOfferSortKey(key);
                          setOfferSortDir("desc");
                        }
                      }}
                    >
                      Discount {offerSortKey === "discount" ? (offerSortDir === "desc" ? "▼" : "▲") : ""}
                      <span className="mt-0.5 block text-[10px] font-semibold normal-case tracking-normal text-textMuted/70">
                        ({formatInLakh(kpis.totalDiscount)})
                      </span>
                    </th>
                    <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Discount %</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedOfferRows.length ? sortedOfferRows.map((offer, index) => (
                    <tr key={`${offer.offerName}-${offer.bankName}`} className={`${index % 2 === 0 ? "bg-white" : "bg-slate-50/60"} cursor-pointer transition hover:bg-blue-50/70`} onClick={() => setSelectedOffer(offer)}>
                      <td className="px-4 py-3 font-semibold text-textMain">
                        {offer.bankName}
                        {offer.paymentCategory === "UPI" ? " (UPI)" : ""}
                      </td>
                      <td className="px-4 py-3 font-bold text-textMain">{offer.offerName}</td>
                      <td className="px-4 py-3 font-semibold text-textMain">{formatCountInLakh(offer.transactions)}</td>
                      <td className="px-4 py-3 font-semibold text-textMain">
                        {formatInLakh(offer.revenue)}
                        <span className="ml-1 text-xs font-bold text-textMuted">
                          ({kpis.grossRevenue ? ((offer.revenue / kpis.grossRevenue) * 100).toFixed(2) : "0.00"}%)
                        </span>
                      </td>
                      <td className="px-4 py-3 font-semibold text-textMain">
                        {formatInLakh(offer.discount)}
                        <span className="ml-1 text-xs font-bold text-textMuted">
                          ({kpis.totalDiscount ? ((offer.discount / kpis.totalDiscount) * 100).toFixed(2) : "0.00"}%)
                        </span>
                      </td>
                      <td className="px-4 py-3 font-bold text-amber-600">{offer.revenue ? ((offer.discount / offer.revenue) * 100).toFixed(1) : "0.0"}%</td>
                    </tr>
                  )) : (
                    <tr>
                      <td colSpan="6" className="px-4 py-10 text-center font-semibold text-textMuted">Upload data to view offer performance.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </section>

        <section className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Month-wise Bank/UPI Performance</p>
              <h2 className="mt-1 font-display text-2xl font-bold text-textMain">
                {trendMode === "monthly"
                  ? "Monthly revenue trend by selected banks / UPI Partners"
                  : trendMode === "mom"
                    ? "Seasonal comparison across years for selected banks"
                    : "Yearly revenue totals for selected banks"}
              </h2>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <div className="inline-flex rounded-full border border-borderSoft bg-white p-1 shadow-sm">
                {TREND_MODES.map((mode) => (
                  <button
                    key={mode.key}
                    type="button"
                    onClick={() => setTrendMode(mode.key)}
                    className={`rounded-full px-3 py-1.5 text-xs font-bold uppercase tracking-wide transition ${
                      trendMode === mode.key ? "bg-accentBlue text-white" : "text-textMuted hover:bg-slate-50"
                    }`}
                  >
                    {mode.label}
                  </button>
                ))}
              </div>
              <button
                type="button"
                onClick={handleViewDetails}
                className="rounded-full border border-borderSoft bg-white px-3 py-1.5 text-xs font-bold uppercase tracking-wide text-textMuted transition hover:bg-slate-50"
              >
                View Details ▸
              </button>
              <div className="rounded-full border border-borderSoft bg-white px-3 py-1.5 text-xs font-bold text-textMuted shadow-sm">
                {formatInteger(totalOfferCountByBank)} Bank / {formatInteger(totalOfferCountByUpiPartner)} UPI offers
              </div>
              {trendMode === "mom" ? (
                <div className="inline-flex rounded-full border border-borderSoft bg-white p-1 shadow-sm">
                  {SEASONAL_METRICS.map((metric) => (
                    <button
                      key={metric.key}
                      type="button"
                      onClick={() => setSeasonalMetric(metric.key)}
                      className={`rounded-full px-3 py-1.5 text-xs font-bold uppercase tracking-wide transition ${
                        seasonalMetric === metric.key ? "bg-accentBlue text-white" : "text-textMuted hover:bg-slate-50"
                      }`}
                    >
                      {metric.label}
                    </button>
                  ))}
                </div>
              ) : null}
            </div>
          </div>
          <div className="mt-4 h-[320px] overflow-hidden rounded-3xl border border-borderSoft bg-white">
            {trendMode === "monthly" ? (
              monthlySeries.length && bankFilter.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={monthlySeries} margin={{ top: 10, right: 20, left: 0, bottom: 10 }}>
                    <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                    <XAxis dataKey="monthKey" stroke="#718096" />
                    <YAxis stroke="#718096" tickFormatter={formatInLakh} />
                    <Tooltip allowEscapeViewBox={{ x: false, y: true }} content={<MonthlyTrendTooltip />} />
                    <Legend />
                    {bankFilter.map((bank) => (
                      <Line
                        key={bank}
                        type="monotone"
                        dataKey={bank}
                        stroke={bankColorMap[bank]}
                        strokeWidth={3}
                        dot={(dotProps) =>
                          renderClickableDot(dotProps, bank, () =>
                            setSelectedMonthlyPoint({
                              bankName: bank,
                              monthKey: dotProps.payload.monthKey,
                              monthLabel: formatMonthKeyLabel(dotProps.payload.monthKey),
                            }),
                          )
                        }
                        activeDot={{ r: 6 }}
                        connectNulls
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Upload data and select one or more banks to view the month-wise revenue chart.
                </div>
              )
            ) : trendMode === "mom" ? (
              seasonalData.length && bankFilter.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={seasonalData} margin={{ top: 10, right: 20, left: 10, bottom: 10 }}>
                    <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                    <XAxis dataKey="monthLabel" stroke="#718096" />
                    <YAxis
                      stroke="#718096"
                      tickFormatter={seasonalMetric === "transactions" || seasonalMetric === "admits" ? formatCountInLakh : formatInLakh}
                      label={{
                        value:
                          seasonalMetric === "revenue"
                            ? "Revenue (₹ L)"
                            : seasonalMetric === "transactions"
                              ? "Transactions"
                              : seasonalMetric === "admits"
                                ? "Admits"
                                : "Discount (₹ L)",
                        angle: -90,
                        position: "insideLeft",
                      }}
                    />
                    <Tooltip
                      allowEscapeViewBox={{ x: false, y: true }}
                      content={<MonthlyTrendTooltip formatValue={seasonalMetric === "transactions" || seasonalMetric === "admits" ? formatCountInLakh : formatInLakh} />}
                    />
                    <Legend />
                    {seasonalYears.map((year, index) => (
                      <Line
                        key={year}
                        type="monotone"
                        dataKey={year}
                        stroke={BANK_COLORS[index % BANK_COLORS.length]}
                        strokeWidth={3}
                        dot={(dotProps) =>
                          renderClickableDot(dotProps, year, () =>
                            setSelectedSeasonalPoint({ month: dotProps.payload.month, monthLabel: dotProps.payload.monthLabel }),
                          )
                        }
                        activeDot={{ r: 6 }}
                        connectNulls
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Select at least one bank to view the month-on-month seasonality chart.
                </div>
              )
            ) : yearlyData.length && bankFilter.length ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={yearlyData} margin={{ top: 24, right: 20, left: 0, bottom: 10 }}>
                  <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                  <XAxis dataKey="year" stroke="#718096" />
                  <YAxis stroke="#718096" tickFormatter={formatInLakh} />
                  <Tooltip formatter={(value) => formatInLakh(value)} contentStyle={{ borderRadius: "18px", borderColor: "#e2e8f0" }} />
                  <Legend />
                  <Bar
                    dataKey="revenue"
                    name="Revenue"
                    fill="#2563eb"
                    radius={[8, 8, 0, 0]}
                    cursor="pointer"
                    activeBar={{ fill: "#1d4ed8" }}
                    onClick={(data) => setSelectedYearPoint({ year: data.year })}
                  >
                    <LabelList dataKey="revenue" position="top" formatter={(value) => formatInLakh(value)} style={{ fontSize: 12, fontWeight: 700, fill: "#1a202c" }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                Select at least one bank to view the year-on-year revenue chart.
              </div>
            )}
          </div>
          <p className="mt-2 flex items-center gap-1.5 text-xs font-semibold text-textMuted">
            <span className="inline-block h-2 w-2 rounded-full bg-accentBlue"></span>
            {trendMode === "monthly"
              ? "Click any point on the chart, or use View Details, to see what changed since the previous month"
              : trendMode === "mom"
                ? "Click any point on the chart, or use View Details, to see what drove that month's change"
                : "Click any bar on the chart, or use View Details, to see what drove that year's change"}
          </p>
          {trendMode === "mom" && selectedSeasonalPoint ? (
            <div className="mt-4 rounded-2xl border border-borderSoft bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between">
                <p className="text-sm font-bold text-textMain">{selectedSeasonalPoint.monthLabel}</p>
                <button type="button" onClick={() => setSelectedSeasonalPoint(null)} className="text-xs font-bold text-textMuted hover:text-textMain">
                  ✕ Close
                </button>
              </div>

              <div className="mb-3 flex flex-col gap-3 sm:flex-row">
                <MetricComparisonBox
                  title="Last Month vs Current"
                  metrics={[
                    { label: "Revenue", current: momAdjacentMonthInsight?.currentRevenue, format: formatInLakh, deltaPercent: momAdjacentMonthInsight?.deltaPercent ?? null },
                    { label: "Admits", current: momAdjacentMonthInsight?.currentAdmits, format: formatCountInLakh, deltaPercent: momAdjacentMonthInsight?.admitsDeltaPercent ?? null },
                    { label: "Transactions", current: momAdjacentMonthInsight?.currentTransactions, format: formatCountInLakh, deltaPercent: momAdjacentMonthInsight?.txnDeltaPercent ?? null },
                  ]}
                />
                <MetricComparisonBox
                  title="Last Year vs Current"
                  metrics={
                    inferencePairs.length
                      ? (() => {
                          const latest = inferencePairs[inferencePairs.length - 1];
                          return [
                            { label: "Revenue", current: latest.currentTotal, format: formatInLakh, deltaPercent: latest.totalDeltaPercent },
                            { label: "Admits", current: latest.currentAdmits, format: formatCountInLakh, deltaPercent: latest.admitsDeltaPercent },
                            { label: "Transactions", current: latest.currentTransactions, format: formatCountInLakh, deltaPercent: latest.txnDeltaPercent },
                          ];
                        })()
                      : [
                          { label: "Revenue", current: null, format: formatInLakh, deltaPercent: null },
                          { label: "Admits", current: null, format: formatCountInLakh, deltaPercent: null },
                          { label: "Transactions", current: null, format: formatCountInLakh, deltaPercent: null },
                        ]
                  }
                />
              </div>

              {inferencePairs.length === 0 ? (
                <p className="text-sm text-textMuted">Not enough historical data for this month to compare.</p>
              ) : (
                inferencePairs.map((pair, idx) => (
                  <div key={`${pair.priorYear}-${pair.currentYear}`} className={idx > 0 ? "mt-4 border-t border-borderSoft pt-4" : ""}>
                    <p className="mb-2 text-sm font-bold text-textMain">
                      {pair.priorYear} → {pair.currentYear}
                    </p>
                    {pair.headline ? (
                      <div className="mb-2 rounded-xl bg-blue-50 p-3">
                        <p className="text-sm font-semibold text-blue-900">{pair.headline}</p>
                      </div>
                    ) : null}
                    <p className="mb-3 text-xs text-textMuted">
                      Shows the size of the year-over-year change, not a reason for it — the data has no field that explains why.
                    </p>
                    <p className="mb-3 text-sm text-textMain">
                      Overall Revenue{" "}
                      {pair.totalDeltaPercent === null
                        ? "change unavailable"
                        : `${pair.totalDeltaPercent > 0 ? "up" : "down"} ${Math.abs(pair.totalDeltaPercent).toFixed(1)}%`}{" "}
                      ({formatInLakh(pair.priorTotal)} → {formatInLakh(pair.currentTotal)})
                    </p>

                    <p className="mb-2 text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Who drove it (Revenue)</p>
                    <div className="mb-3 space-y-1.5">
                      {pair.bankBreakdown.slice(0, 5).map((b) => {
                        const range = bankActiveRangeMap.get(b.bankName);
                        return (
                          <div key={b.bankName} className="flex items-center gap-3 text-sm">
                            <div className="w-40 flex-shrink-0">
                              <span className="text-textMain">
                                Rank #{globalBankRankMap.get(b.bankName) || "—"} - {b.bankName}
                              </span>
                              {range ? (
                                <p className="text-[11px] text-textMuted">
                                  {formatIsoDate(range.start)} – {formatIsoDate(range.end)}
                                </p>
                              ) : null}
                            </div>
                            <span className={b.deltaAbs >= 0 ? "font-semibold text-emerald-600" : "font-semibold text-rose-500"}>
                              {b.deltaAbs >= 0 ? "+" : ""}
                              {formatInLakh(b.deltaAbs)} Revenue ({b.contributionPercent === null ? "—" : `${b.contributionPercent.toFixed(0)}%`} of revenue change)
                            </span>
                          </div>
                        );
                      })}
                    </div>

                    <div className="mb-3 grid grid-cols-2 gap-3">
                      <div className="rounded-xl bg-white p-3">
                        <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Discount rate this month</p>
                        <p className="text-sm text-textMain">
                          {pair.priorDiscountRate === null ? "—" : `${pair.priorDiscountRate.toFixed(1)}%`} →{" "}
                          {pair.currentDiscountRate === null ? "—" : `${pair.currentDiscountRate.toFixed(1)}%`}
                        </p>
                        <p className="mt-1 text-xs text-textMuted">{pair.discountRateNote}</p>
                      </div>
                      <div className="rounded-xl bg-white p-3">
                        <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Volume vs value</p>
                        <p className="text-sm text-textMain">
                          Txns {pair.txnDeltaPercent === null ? "—" : `${pair.txnDeltaPercent > 0 ? "+" : ""}${pair.txnDeltaPercent.toFixed(1)}%`} · Revenue{" "}
                          {pair.totalDeltaPercent === null ? "—" : `${pair.totalDeltaPercent > 0 ? "+" : ""}${pair.totalDeltaPercent.toFixed(1)}%`}
                        </p>
                        <p className="mt-1 text-xs text-textMuted">{pair.volumeValueNote}</p>
                      </div>
                      <div className="rounded-xl bg-white p-3">
                        <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Admits</p>
                        <p className="text-sm text-textMain">
                          {formatCountInLakh(pair.priorAdmits)} → {formatCountInLakh(pair.currentAdmits)}
                        </p>
                        <p className="mt-1 text-xs text-textMuted">{pair.admitsNote}</p>
                      </div>
                      <div className="rounded-xl bg-white p-3">
                        <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">ATV</p>
                        <p className="text-sm text-textMain">
                          {pair.priorATV === null ? "—" : formatRupee(pair.priorATV)} → {pair.currentATV === null ? "—" : formatRupee(pair.currentATV)}
                        </p>
                        <p className="mt-1 text-xs text-textMuted">{pair.atvNote}</p>
                      </div>
                    </div>

                    <p className="text-xs italic text-textMuted">
                      Worth checking: was there a specific offer or partnership change behind {pair.topBank ? pair.topBank : "this shift"} in{" "}
                      {selectedSeasonalPoint.monthLabel.split(" ")[0]} {pair.currentYear}?
                    </p>
                  </div>
                ))
              )}
            </div>
          ) : null}

          {trendMode === "monthly" && selectedMonthlyPoint ? (
            <div className="mt-4 rounded-2xl border border-borderSoft bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between">
                <p className="text-sm font-bold text-textMain">
                  {selectedMonthlyPoint.bankName} — {selectedMonthlyPoint.monthLabel}
                </p>
                <button type="button" onClick={() => setSelectedMonthlyPoint(null)} className="text-xs font-bold text-textMuted hover:text-textMain">
                  ✕ Close
                </button>
              </div>

              <div className="mb-3 space-y-1 rounded-xl bg-blue-50 p-3">
                <p className="text-sm font-semibold text-blue-900">
                  Last Year vs Current Year:{" "}
                  {monthlyYoY
                    ? `${formatInLakh(monthlyYoY.priorRevenue)} → ${formatInLakh(monthlyYoY.currentRevenue)} (${
                        monthlyYoY.deltaPercent >= 0 ? "▲" : "▼"
                      } ${Math.abs(monthlyYoY.deltaPercent).toFixed(1)}%)`
                    : "—"}
                </p>
                <p className="text-sm font-semibold text-blue-900">
                  Current Month vs Last Month:{" "}
                  {monthlyInsight.hasPrior && monthlyInsight.totalDeltaPercent !== null
                    ? `${formatInLakh(monthlyInsight.priorTotals.revenue)} → ${formatInLakh(monthlyInsight.currentTotals.revenue)} (${
                        monthlyInsight.totalDeltaPercent >= 0 ? "▲" : "▼"
                      } ${Math.abs(monthlyInsight.totalDeltaPercent).toFixed(1)}%)`
                    : "—"}
                </p>
                <p className="text-sm font-semibold text-blue-900">
                  Rank: {bankRank ? `#${bankRank} of ${bankRows.length} banks by revenue` : "—"}
                </p>
              </div>

              {monthlyInsight.hasPrior ? (
                <>
                  {monthlyInsight.headline ? (
                    <div className="mb-2 rounded-xl bg-blue-50 p-3">
                      <p className="text-sm font-semibold text-blue-900">{monthlyInsight.headline}</p>
                    </div>
                  ) : null}
                  <p className="mb-3 text-xs text-textMuted">
                    Shows the size of the change, not a reason for it — the data has no field that explains why.
                  </p>
                  <p className="mb-3 text-sm text-textMain">
                    Overall{" "}
                    {monthlyInsight.totalDeltaPercent === null
                      ? "change unavailable"
                      : `${monthlyInsight.totalDeltaPercent > 0 ? "up" : "down"} ${Math.abs(monthlyInsight.totalDeltaPercent).toFixed(1)}%`}{" "}
                    ({formatInLakh(monthlyInsight.priorTotals.revenue)} → {formatInLakh(monthlyInsight.currentTotals.revenue)})
                  </p>
                  <div className="grid grid-cols-2 gap-3">
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Discount rate</p>
                      <p className="text-sm text-textMain">
                        {monthlyInsight.priorDiscountRate === null ? "—" : `${monthlyInsight.priorDiscountRate.toFixed(1)}%`} →{" "}
                        {monthlyInsight.currentDiscountRate === null ? "—" : `${monthlyInsight.currentDiscountRate.toFixed(1)}%`}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{monthlyInsight.discountRateNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Volume vs value</p>
                      <p className="text-sm text-textMain">
                        Txns{" "}
                        {monthlyInsight.txnDeltaPercent === null
                          ? "—"
                          : `${monthlyInsight.txnDeltaPercent > 0 ? "+" : ""}${monthlyInsight.txnDeltaPercent.toFixed(1)}%`}{" "}
                        · Revenue{" "}
                        {monthlyInsight.totalDeltaPercent === null
                          ? "—"
                          : `${monthlyInsight.totalDeltaPercent > 0 ? "+" : ""}${monthlyInsight.totalDeltaPercent.toFixed(1)}%`}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{monthlyInsight.volumeValueNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Admits</p>
                      <p className="text-sm text-textMain">
                        {formatCountInLakh(monthlyInsight.priorAdmits)} → {formatCountInLakh(monthlyInsight.currentAdmits)}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{monthlyInsight.admitsNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">ATV</p>
                      <p className="text-sm text-textMain">
                        {monthlyInsight.priorATV === null ? "—" : formatRupee(monthlyInsight.priorATV)} →{" "}
                        {monthlyInsight.currentATV === null ? "—" : formatRupee(monthlyInsight.currentATV)}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{monthlyInsight.atvNote}</p>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  <p className="mb-1 text-xs font-semibold text-amber-700">No prior month to compare — showing this month's figures.</p>
                  <p className="text-sm text-textMain">
                    Revenue: {formatInLakh(monthlyInsight.currentTotals.revenue)} · Transactions:{" "}
                    {formatCountInLakh(monthlyInsight.currentTotals.transactions)} · Discount: {formatInLakh(monthlyInsight.currentTotals.discount)} ·
                    Discount rate: {monthlyInsight.currentDiscountRate === null ? "—" : `${monthlyInsight.currentDiscountRate.toFixed(1)}%`} · Admits:{" "}
                    {formatCountInLakh(monthlyInsight.currentAdmits)} · ATV:{" "}
                    {monthlyInsight.currentATV === null ? "—" : formatRupee(monthlyInsight.currentATV)}
                  </p>
                </>
              )}
            </div>
          ) : null}

          {trendMode === "yoy" && selectedYearPoint ? (
            <div className="mt-4 rounded-2xl border border-borderSoft bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between">
                <p className="text-sm font-bold text-textMain">{selectedYearPoint.year}</p>
                <button type="button" onClick={() => setSelectedYearPoint(null)} className="text-xs font-bold text-textMuted hover:text-textMain">
                  ✕ Close
                </button>
              </div>

              <div className="mb-3">
                <MetricComparisonBox
                  title="Last Year vs Current Year"
                  metrics={[
                    { label: "Revenue", current: yearInsight.hasPrior ? yearInsight.currentTotal : null, format: formatInLakh, deltaPercent: yearInsight.hasPrior ? yearInsight.totalDeltaPercent : null },
                    { label: "Admits", current: yearInsight.hasPrior ? yearInsight.currentAdmits : null, format: formatCountInLakh, deltaPercent: yearInsight.hasPrior ? yearInsight.admitsDeltaPercent : null },
                    { label: "Transactions", current: yearInsight.hasPrior ? yearInsight.currentTransactions : null, format: formatCountInLakh, deltaPercent: yearInsight.hasPrior ? yearInsight.txnDeltaPercent : null },
                  ]}
                />
              </div>

              {yearInsight.hasPrior ? (
                <>
                  {yearInsight.headline ? (
                    <div className="mb-2 rounded-xl bg-blue-50 p-3">
                      <p className="text-sm font-semibold text-blue-900">{yearInsight.headline}</p>
                    </div>
                  ) : null}
                  <p className="mb-3 text-xs text-textMuted">
                    Shows the size of the change, not a reason for it — the data has no field that explains why.
                  </p>
                  <p className="mb-3 text-sm text-textMain">
                    Overall Revenue{" "}
                    {yearInsight.totalDeltaPercent === null
                      ? "change unavailable"
                      : `${yearInsight.totalDeltaPercent > 0 ? "up" : "down"} ${Math.abs(yearInsight.totalDeltaPercent).toFixed(1)}%`}{" "}
                    ({formatInLakh(yearInsight.priorTotal)} → {formatInLakh(yearInsight.currentTotal)})
                  </p>

                  <p className="mb-2 text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Who drove it (Revenue)</p>
                  <div className="mb-3 space-y-1.5">
                    {yearInsight.bankBreakdown.slice(0, 5).map((b) => {
                      const range = bankActiveRangeMap.get(b.bankName);
                      return (
                        <div key={b.bankName} className="flex items-center gap-3 text-sm">
                          <div className="w-40 flex-shrink-0">
                            <span className="text-textMain">
                              Rank #{globalBankRankMap.get(b.bankName) || "—"} - {b.bankName}
                            </span>
                            {range ? (
                              <p className="text-[11px] text-textMuted">
                                {formatIsoDate(range.start)} – {formatIsoDate(range.end)}
                              </p>
                            ) : null}
                          </div>
                          <span className={b.deltaAbs >= 0 ? "font-semibold text-emerald-600" : "font-semibold text-rose-500"}>
                            {b.deltaAbs >= 0 ? "+" : ""}
                            {formatInLakh(b.deltaAbs)} Revenue ({b.contributionPercent === null ? "—" : `${b.contributionPercent.toFixed(0)}%`} of revenue change)
                          </span>
                        </div>
                      );
                    })}
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Discount rate</p>
                      <p className="text-sm text-textMain">
                        {yearInsight.priorDiscountRate === null ? "—" : `${yearInsight.priorDiscountRate.toFixed(1)}%`} →{" "}
                        {yearInsight.currentDiscountRate === null ? "—" : `${yearInsight.currentDiscountRate.toFixed(1)}%`}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{yearInsight.discountRateNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Volume vs value</p>
                      <p className="text-sm text-textMain">
                        Txns{" "}
                        {yearInsight.txnDeltaPercent === null ? "—" : `${yearInsight.txnDeltaPercent > 0 ? "+" : ""}${yearInsight.txnDeltaPercent.toFixed(1)}%`}{" "}
                        · Revenue{" "}
                        {yearInsight.totalDeltaPercent === null
                          ? "—"
                          : `${yearInsight.totalDeltaPercent > 0 ? "+" : ""}${yearInsight.totalDeltaPercent.toFixed(1)}%`}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{yearInsight.volumeValueNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">Admits</p>
                      <p className="text-sm text-textMain">
                        {formatCountInLakh(yearInsight.priorAdmits)} → {formatCountInLakh(yearInsight.currentAdmits)}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{yearInsight.admitsNote}</p>
                    </div>
                    <div className="rounded-xl bg-white p-3">
                      <p className="text-xs font-bold uppercase tracking-[0.15em] text-textMuted">ATV</p>
                      <p className="text-sm text-textMain">
                        {yearInsight.priorATV === null ? "—" : formatRupee(yearInsight.priorATV)} →{" "}
                        {yearInsight.currentATV === null ? "—" : formatRupee(yearInsight.currentATV)}
                      </p>
                      <p className="mt-1 text-xs text-textMuted">{yearInsight.atvNote}</p>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  <p className="mb-1 text-xs font-semibold text-amber-700">No prior year to compare — showing this year's figures.</p>
                  <p className="text-sm text-textMain">
                    Revenue: {formatInLakh(yearInsight.currentTotals.revenue)} · Transactions: {formatCountInLakh(yearInsight.currentTotals.transactions)}{" "}
                    · Discount: {formatInLakh(yearInsight.currentTotals.discount)} · Discount rate:{" "}
                    {yearInsight.currentDiscountRate === null ? "—" : `${yearInsight.currentDiscountRate.toFixed(1)}%`} · Admits:{" "}
                    {formatCountInLakh(yearInsight.currentAdmits)} · ATV:{" "}
                    {yearInsight.currentATV === null ? "—" : formatRupee(yearInsight.currentATV)}
                  </p>
                </>
              )}
            </div>
          ) : null}
        </section>
      </div>

      {showComparisonModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={() => setShowComparisonModal(false)}>
          <div className="max-h-[90vh] w-full max-w-4xl overflow-y-auto rounded-[2rem] bg-white p-6 shadow-2xl" onClick={(e) => e.stopPropagation()}>
            <div className="mb-2 flex items-center justify-between">
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Custom Comparison</p>
                <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Bank/UPI Group A vs Bank/UPI Group B</h2>
              </div>
              <button type="button" onClick={() => setShowComparisonModal(false)} className="text-sm font-bold text-textMuted hover:text-textMain">
                ✕ Close
              </button>
            </div>
            <p className="mt-1 text-sm text-textMuted">
              Independent of the main Date/Bank/Offer filters above — each group has its own banks and date range.
            </p>

            <div className="mt-4 grid gap-4 md:grid-cols-2">
              <div className="rounded-2xl border border-borderSoft bg-slate-50 p-4">
                <div className="flex items-center justify-between">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Group A</p>
                  <button
                    type="button"
                    onClick={() => setGroupDetailOpen("A")}
                    className="rounded-xl bg-white px-2.5 py-1 text-right text-[11px] font-bold leading-tight text-textMuted shadow-sm hover:bg-slate-100"
                  >
                    BANK: {groupACardCount}
                    <br />
                    UPI: {groupAUpiCount}
                  </button>
                </div>
                <div className="mt-2">
                  <MultiSelectDropdown
                    options={banks}
                    selected={groupABanks}
                    onToggle={(bank) =>
                      setGroupABanks((current) => (current.includes(bank) ? current.filter((item) => item !== bank) : [...current, bank]))
                    }
                    onClear={() => setGroupABanks([])}
                    label="bank"
                  />
                </div>
                <div className="mt-3 flex flex-col gap-2">
                  <MultiSelectDropdown
                    options={fiscalYears}
                    selected={groupAFy}
                    onToggle={(fy) => setGroupAFy((current) => (current.includes(fy) ? current.filter((i) => i !== fy) : [...current, fy]))}
                    onClear={() => setGroupAFy([])}
                    label="FY"
                    showSelectedNames
                  />
                  <MultiSelectDropdown
                    options={fiscalMonths}
                    selected={groupAMonths}
                    onToggle={(m) => setGroupAMonths((current) => (current.includes(m) ? current.filter((i) => i !== m) : [...current, m]))}
                    onClear={() => setGroupAMonths([])}
                    label="month"
                    showSelectedNames
                  />
                </div>
              </div>

              <div className="rounded-2xl border border-borderSoft bg-slate-50 p-4">
                <div className="flex items-center justify-between">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Group B</p>
                  <button
                    type="button"
                    onClick={() => setGroupDetailOpen("B")}
                    className="rounded-xl bg-white px-2.5 py-1 text-right text-[11px] font-bold leading-tight text-textMuted shadow-sm hover:bg-slate-100"
                  >
                    BANK: {groupBCardCount}
                    <br />
                    UPI: {groupBUpiCount}
                  </button>
                </div>
                <div className="mt-2">
                  <MultiSelectDropdown
                    options={banks}
                    selected={groupBBanks}
                    onToggle={(bank) =>
                      setGroupBBanks((current) => (current.includes(bank) ? current.filter((item) => item !== bank) : [...current, bank]))
                    }
                    onClear={() => setGroupBBanks([])}
                    label="bank"
                  />
                </div>
                <div className="mt-3 flex flex-col gap-2">
                  <MultiSelectDropdown
                    options={fiscalYears}
                    selected={groupBFy}
                    onToggle={(fy) => setGroupBFy((current) => (current.includes(fy) ? current.filter((i) => i !== fy) : [...current, fy]))}
                    onClear={() => setGroupBFy([])}
                    label="FY"
                    showSelectedNames
                  />
                  <MultiSelectDropdown
                    options={fiscalMonths}
                    selected={groupBMonths}
                    onToggle={(m) => setGroupBMonths((current) => (current.includes(m) ? current.filter((i) => i !== m) : [...current, m]))}
                    onClear={() => setGroupBMonths([])}
                    label="month"
                    showSelectedNames
                  />
                </div>
              </div>
            </div>

            <div className="mt-5">
              {!groupAFy.length || !groupAMonths.length || !groupBFy.length || !groupBMonths.length ? (
                <p className="text-sm text-textMuted">Select FY and Month for both groups to compare.</p>
              ) : !groupARows.length || !groupBRows.length ? (
                <p className="text-sm text-textMuted">No data for this selection.</p>
              ) : (
                <div className="overflow-x-auto">
                  <table className="min-w-full text-sm">
                    <thead>
                      <tr className="text-left text-xs font-bold uppercase tracking-[0.15em] text-textMuted">
                        <th className="py-2 pr-4">Metric</th>
                        <th className="py-2 pr-4">Group A</th>
                        <th className="py-2 pr-4">Group B</th>
                        <th className="py-2">Delta (B vs A)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-borderSoft">
                      <tr>
                        <td className="py-2 pr-4 font-semibold text-textMain">Revenue</td>
                        <td className="py-2 pr-4 text-textMain">{formatInLakh(groupAKpis.grossRevenue)}</td>
                        <td className="py-2 pr-4 text-textMain">{formatInLakh(groupBKpis.grossRevenue)}</td>
                        <td className="py-2">{renderDeltaBadge(computeDelta(groupBKpis.grossRevenue, groupAKpis.grossRevenue))}</td>
                      </tr>
                      <tr>
                        <td className="py-2 pr-4 font-semibold text-textMain">Transactions</td>
                        <td className="py-2 pr-4 text-textMain">{formatCountInLakh(groupAKpis.totalTransactions)}</td>
                        <td className="py-2 pr-4 text-textMain">{formatCountInLakh(groupBKpis.totalTransactions)}</td>
                        <td className="py-2">{renderDeltaBadge(computeDelta(groupBKpis.totalTransactions, groupAKpis.totalTransactions))}</td>
                      </tr>
                      <tr>
                        <td className="py-2 pr-4 font-semibold text-textMain">Discount Amount</td>
                        <td className="py-2 pr-4 text-textMain">{formatInLakh(groupAKpis.totalDiscount)}</td>
                        <td className="py-2 pr-4 text-textMain">{formatInLakh(groupBKpis.totalDiscount)}</td>
                        <td className="py-2">{renderDeltaBadge(computeDelta(groupBKpis.totalDiscount, groupAKpis.totalDiscount))}</td>
                      </tr>
                      <tr>
                        <td className="py-2 pr-4 font-semibold text-textMain">Admits</td>
                        <td className="py-2 pr-4 text-textMain">{formatCountInLakh(groupAAdmits)}</td>
                        <td className="py-2 pr-4 text-textMain">{formatCountInLakh(groupBAdmits)}</td>
                        <td className="py-2">{renderDeltaBadge(computeDelta(groupBAdmits, groupAAdmits))}</td>
                      </tr>
                      <tr>
                        <td className="py-2 pr-4 font-semibold text-textMain">Discount Rate</td>
                        <td className="py-2 pr-4 text-textMain">{groupADiscountRate.toFixed(1)}%</td>
                        <td className="py-2 pr-4 text-textMain">{groupBDiscountRate.toFixed(1)}%</td>
                        <td className="py-2">{renderDeltaBadge(computeDelta(groupBDiscountRate, groupADiscountRate))}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      ) : null}

      <GroupDetailModal
        group={groupDetailOpen}
        breakdown={groupDetailOpen === "A" ? groupABankBreakdown : groupDetailOpen === "B" ? groupBBankBreakdown : []}
        onClose={() => setGroupDetailOpen(null)}
      />

      <OfferModal offer={selectedOffer} onClose={() => setSelectedOffer(null)} />
      <BankModal
        bank={selectedBankRow}
        offersEntry={selectedBankOffersEntry}
        discountEntry={selectedBankDiscountEntry}
        onClose={() => setSelectedBank(null)}
      />
      {showOffersByBank ? (
        <OffersByBankModal
          offersByBank={offersByBank}
          totalOfferCountByBank={totalOfferCountByBank}
          expandedOfferBank={expandedOfferBank}
          onToggleBank={toggleExpandedOfferBank}
          onClose={() => setShowOffersByBank(false)}
        />
      ) : null}
      {showOffersByUpi ? (
        <OffersByBankModal
          offersByBank={upiOffersByPartner}
          totalOfferCountByBank={totalOfferCountByUpiPartner}
          expandedOfferBank={expandedOfferUpiPartner}
          onToggleBank={toggleExpandedOfferUpiPartner}
          onClose={() => setShowOffersByUpi(false)}
          heading="Total Offers by Each UPI Partner"
          entityNoun="partners"
        />
      ) : null}
    </div>
  );
}
