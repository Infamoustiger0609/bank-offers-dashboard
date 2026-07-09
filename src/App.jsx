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
};

const OPTIONAL_COLUMNS = new Set(["universalTransactions"]);

const CHANNEL_COLORS = {
  Online: "#2563eb",
  "Offline - Box Office": "#f59e0b",
  "Offline - F&B": "#10b981",
};
const CHANNEL_FALLBACK_COLOR = "#94a3b8";

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
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  const parts = raw.split(/[/-]/).map((part) => Number(part.trim()));
  if (parts.length === 3) {
    const [day, month, year] = parts;
    const date = new Date(year < 100 ? 2000 + year : year, (month || 1) - 1, day || 1);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  return null;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function formatCompactCurrency(value) {
  const amount = Number(value || 0);
  if (Math.abs(amount) >= 10000000) return `₹${(amount / 10000000).toFixed(1)} Cr`;
  if (Math.abs(amount) >= 100000) return `₹${(amount / 100000).toFixed(1)} L`;
  if (Math.abs(amount) >= 1000) return `₹${(amount / 1000).toFixed(1)}K`;
  return `₹${Math.round(amount)}`;
}

function formatCompactNumber(value) {
  const amount = Number(value || 0);
  if (Math.abs(amount) >= 10000000) return `${(amount / 10000000).toFixed(2)} Cr`;
  if (Math.abs(amount) >= 100000) return `${(amount / 100000).toFixed(2)} L`;
  if (Math.abs(amount) >= 1000) return `${(amount / 1000).toFixed(1)}K`;
  return `${Math.round(amount)}`;
}

function formatInCrore(value) {
  return `${(Number(value || 0) / 10000000).toFixed(2)} Cr`;
}

function formatInLakh(value) {
  return `₹${(Number(value || 0) / 100000).toFixed(2)} L`;
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

function monthKeyFromDate(date) {
  if (!date) return "Unknown";
  return `${String(date.getMonth() + 1).padStart(2, "0")}-${date.getFullYear()}`;
}

function getFiscalYearStart(date) {
  const year = date.getMonth() >= 3 ? date.getFullYear() : date.getFullYear() - 1;
  return new Date(year, 3, 1);
}

function getFiscalQuarterStart(date) {
  const month = date.getMonth();
  const year = date.getFullYear();
  if (month >= 3 && month <= 5) return new Date(year, 3, 1);
  if (month >= 6 && month <= 8) return new Date(year, 6, 1);
  if (month >= 9 && month <= 11) return new Date(year, 9, 1);
  return new Date(year, 0, 1);
}

const QUICK_PERIODS = [
  { key: "thisMonth", label: "This Month" },
  { key: "lastMonth", label: "Last Month" },
  { key: "qtd", label: "QTD" },
  { key: "fytd", label: "FYTD" },
];

function getQuickPeriodRange(key, latestDate) {
  if (!latestDate) return null;
  const year = latestDate.getFullYear();
  const month = latestDate.getMonth();
  switch (key) {
    case "thisMonth":
      return { rangeStart: new Date(year, month, 1), rangeEnd: latestDate };
    case "lastMonth":
      return { rangeStart: new Date(year, month - 1, 1), rangeEnd: new Date(year, month, 0, 23, 59, 59, 999) };
    case "qtd":
      return { rangeStart: getFiscalQuarterStart(latestDate), rangeEnd: latestDate };
    case "fytd":
      return { rangeStart: getFiscalYearStart(latestDate), rangeEnd: latestDate };
    default:
      return null;
  }
}

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

function buildColumnLookup(headers) {
  const normalizedHeaders = {};
  headers.forEach((header) => {
    normalizedHeaders[normalizeKey(header)] = header;
  });
  return Object.entries(COLUMN_MAP).reduce((lookup, [field, expected]) => {
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
      offerType: inferOfferType(offerName),
      transactionType: lookup.transactionType ? String(row[lookup.transactionType] || "").trim() || "Unknown" : "Unknown",
      universalTransactions:
        lookup.universalTransactions && row[lookup.universalTransactions] !== undefined && row[lookup.universalTransactions] !== ""
          ? Number(row[lookup.universalTransactions])
          : null,
    };
  });

  return { parsedRows, missingColumns };
}

function getUniversalTransactionsByMonth(rows) {
  const map = new Map();
  const inconsistent = new Set();
  rows.forEach((row) => {
    if (row.universalTransactions === null || row.monthKey === "Unknown") return;
    if (map.has(row.monthKey) && map.get(row.monthKey) !== row.universalTransactions) {
      inconsistent.add(row.monthKey);
    }
    map.set(row.monthKey, row.universalTransactions);
  });
  return { map, inconsistent };
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
      return acc;
    },
    { ...EMPTY_KPIS },
  );
}

function aggregateBanks(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const current = grouped.get(row.bankName) || {
      bankName: row.bankName,
      totalRevenue: 0,
      discountCost: 0,
      totalTransactions: 0,
    };
    current.totalRevenue += row.transactionTotal;
    current.discountCost += row.discountAmount;
    current.totalTransactions += row.discountedTransactions;
    grouped.set(row.bankName, current);
  });
  return [...grouped.values()].sort((a, b) => b.totalRevenue - a.totalRevenue);
}

function aggregateOffers(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = `${row.offerName}::${row.bankName}`;
    const current = grouped.get(key) || {
      offerName: row.offerName,
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
    };
    current.revenue += row.transactionTotal;
    current.netRevenue += row.amountPaid;
    current.discount += row.discountAmount;
    current.transactions += row.discountedTransactions;
    current.freeTickets += row.freeTickets;
    current.totalTickets += row.totalTickets;
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
  rows.forEach((row) => {
    const key = row.transactionType || "Unknown";
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

function aggregateSeasonalByYear(rows, selectedBanks) {
  const seasonMap = new Map();
  rows.forEach((row) => {
    if (!selectedBanks.includes(row.bankName)) return;
    if (row.monthKey === "Unknown") return;
    const [month, year] = row.monthKey.split("-").map(Number);
    if (!seasonMap.has(month)) seasonMap.set(month, { month, monthLabel: MONTH_NAMES[month - 1] });
    const current = seasonMap.get(month);
    current[year] = (current[year] || 0) + row.transactionTotal;
  });
  return [...seasonMap.values()].sort((a, b) => a.month - b.month);
}

function aggregateYearlyTotals(rows, selectedBanks) {
  const yearMap = new Map();
  rows.forEach((row) => {
    if (!selectedBanks.includes(row.bankName)) return;
    if (row.monthKey === "Unknown") return;
    const [, year] = row.monthKey.split("-").map(Number);
    yearMap.set(year, (yearMap.get(year) || 0) + row.transactionTotal);
  });
  return [...yearMap.entries()].map(([year, revenue]) => ({ year: String(year), revenue })).sort((a, b) => Number(a.year) - Number(b.year));
}

function safeRatio(numerator, denominator) {
  if (!denominator) return 0;
  return (numerator / denominator) * 100;
}

function StatCard({ title, value, subtitle, color, icon, delta, extra }) {
  return (
    <div className={`min-w-0 rounded-3xl border p-5 shadow-soft ${color.accent}`}>
      <div className="flex items-start justify-between gap-3">
        <p className="min-w-0 flex-1 text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">{title}</p>
        <div
          className={`flex h-10 min-w-[50px] flex-shrink-0 items-center justify-center rounded-2xl px-3 ${color.dot} text-[14px] font-semibold text-white`}
          style={{ whiteSpace: "nowrap" }}
        >
          {icon}
        </div>
      </div>
      <p className="mt-3 overflow-hidden text-ellipsis whitespace-nowrap text-[24px] font-bold leading-none text-textMain">{value}</p>
      {subtitle ? <p className="mt-2 text-sm font-medium text-slate-500">{subtitle}</p> : null}
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

function UploadPanel({ onFileChange, dragActive, setDragActive, fileName, error, onRemoveFile, inputRef }) {
  function handleDrop(event) {
    event.preventDefault();
    setDragActive(false);
    const [file] = [...(event.dataTransfer.files || [])];
    if (file) onFileChange(file);
  }

  return (
    <div className="rounded-3xl border border-dashed border-borderSoft bg-white/90 p-4 shadow-soft">
      <div
        className={`flex cursor-pointer flex-col items-center justify-center rounded-3xl border-2 border-dashed px-6 py-6 text-center transition ${dragActive ? "border-accentBlue bg-blue-50" : "border-slate-200 bg-slate-50/70 hover:border-accentGreen hover:bg-emerald-50"}`}
        onClick={() => inputRef.current?.click()}
        onDragOver={(event) => {
          event.preventDefault();
          setDragActive(true);
        }}
        onDragLeave={() => setDragActive(false)}
        onDrop={handleDrop}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          className="hidden"
          onChange={(event) => {
            const [file] = [...(event.target.files || [])];
            if (file) onFileChange(file);
          }}
        />
        <p className="font-display text-lg font-bold text-textMain">Upload Excel Performance File</p>
        <p className="mt-2 text-sm text-textMuted">Drag and drop or click to upload `.xlsx`, `.xls`, or `.csv`</p>
        {fileName ? (
          <div className="mt-3 flex items-center gap-2">
            <p className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600">{fileName}</p>
            <button
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                onRemoveFile();
              }}
              className="rounded-full border border-rose-200 bg-rose-50 px-3 py-1 text-xs font-bold text-rose-600 hover:bg-rose-100"
            >
              ✕ Remove
            </button>
          </div>
        ) : (
          <p className="mt-3 rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600">No file selected</p>
        )}
      </div>
      {error ? <p className="mt-3 text-sm font-semibold text-rose-600">{error}</p> : null}
    </div>
  );
}

function MultiSelectDropdown({ options, selected, onToggle, onClear, label }) {
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
        <span>
          {selected.length === 0
            ? `No ${label}s selected`
            : allSelected
              ? `All ${label.charAt(0).toUpperCase()}${label.slice(1)}s`
              : `${selected.length} ${label}${selected.length > 1 ? "s" : ""} selected`}
        </span>
        <span className="text-textMuted">{open ? "▲" : "▼"}</span>
      </button>
      {open ? (
        <div className="absolute right-0 z-20 mt-2 w-full rounded-2xl border border-borderSoft bg-white p-3 shadow-soft">
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

function DateMultiSelectDropdown({ options, selected, onToggle, onClear, onApplyRange }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const [rangeStart, setRangeStart] = useState("");
  const [rangeEnd, setRangeEnd] = useState("");
  const allSelected = options.length > 0 && selected.length === options.length;

  useEffect(() => {
    function handleClick(event) {
      if (!panelRef.current?.contains(event.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  const summary =
    selected.length === 0
      ? "No dates selected"
      : allSelected
        ? "All Dates"
        : `${selected.length} date${selected.length > 1 ? "s" : ""} selected`;

  return (
    <div ref={panelRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen((value) => !value)}
        className="flex w-full min-w-[220px] items-center justify-between rounded-2xl border border-borderSoft bg-white px-4 py-3 text-left text-sm font-semibold text-textMain shadow-sm"
      >
        <span>{summary}</span>
        <span className="text-textMuted">{open ? "▲" : "▼"}</span>
      </button>
      {open ? (
        <div className="absolute left-0 z-20 mt-2 w-full rounded-2xl border border-borderSoft bg-white p-3 shadow-soft">
          <div className="mb-2 flex items-center justify-between">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Dates</p>
            <button type="button" onClick={onClear} className="text-xs font-bold text-accentBlue">
              Clear
            </button>
          </div>
          <div className="mb-3 border-b border-borderSoft pb-3">
            <p className="mb-2 text-xs font-bold uppercase tracking-[0.2em] text-textMuted">Custom Range</p>
            <div className="flex flex-col gap-2">
              <div>
                <label className="mb-1 block text-[10px] font-bold uppercase text-textMuted">From</label>
                <input type="month" value={rangeStart} onChange={(e) => setRangeStart(e.target.value)} className="w-full rounded-lg border border-borderSoft px-2 py-1.5 text-xs" />
              </div>
              <div>
                <label className="mb-1 block text-[10px] font-bold uppercase text-textMuted">To</label>
                <input type="month" value={rangeEnd} onChange={(e) => setRangeEnd(e.target.value)} className="w-full rounded-lg border border-borderSoft px-2 py-1.5 text-xs" />
              </div>
            </div>
            <button
              type="button"
              disabled={!rangeStart || !rangeEnd}
              onClick={() => onApplyRange(rangeStart, rangeEnd)}
              className="mt-2 w-full rounded-lg bg-accentBlue py-1.5 text-xs font-bold text-white disabled:opacity-40"
            >
              Apply Range
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
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">{offer.bankName}</p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">{offer.offerName}</h3>
            <p className="mt-2 inline-flex rounded-full bg-blue-50 px-3 py-1 text-xs font-bold text-accentBlue">{offer.offerType}</p>
          </div>
          <button type="button" onClick={onClose} className="rounded-full border border-borderSoft px-3 py-2 text-sm font-bold text-textMuted hover:bg-slate-50">
            Close
          </button>
        </div>
        <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <MetricBlock title="Revenue" value={formatCurrency(offer.revenue)} />
          <MetricBlock title="Net Revenue" value={formatCurrency(offer.netRevenue)} />
          <MetricBlock title="Discount" value={formatCurrency(offer.discount)} />
          <MetricBlock title="Transactions" value={formatInteger(offer.transactions)} />
          <MetricBlock title="Free Tickets" value={formatInteger(offer.freeTickets)} />
        </div>
        <div className="mt-4 grid gap-4 sm:grid-cols-2">
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Ticket & F&B</p>
            <p className="mt-3 text-sm font-semibold text-textMain">Ticket Revenue: {formatCurrency(offer.ticketRevenue)}</p>
            <p className="mt-2 text-sm font-semibold text-textMain">F&B Revenue: {formatCurrency(offer.fnbRevenue)}</p>
            <p className="mt-2 text-sm font-semibold text-textMain">Total Tickets: {formatInteger(offer.totalTickets)}</p>
          </div>
          <div className="rounded-2xl border border-borderSoft p-4">
            <p className="text-xs font-bold uppercase tracking-[0.18em] text-textMuted">Contribution Split</p>
            <p className="mt-3 text-sm font-semibold text-textMain">
              Bank: {formatCurrency(offer.bankContribution)} ({offer.discount ? ((offer.bankContribution / offer.discount) * 100).toFixed(0) : 0}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">
              Inox: {formatCurrency(offer.inoxContribution)} ({offer.discount ? ((offer.inoxContribution / offer.discount) * 100).toFixed(0) : 0}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">Conv. Fees: {formatCurrency(offer.convFees)}</p>
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
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Bank Scorecard</p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">{bank.bankName}</h3>
            <p className="mt-2 inline-flex rounded-full bg-blue-50 px-3 py-1 text-xs font-bold text-accentBlue">{formatInteger(offers.length)} offers</p>
          </div>
          <button type="button" onClick={onClose} className="rounded-full border border-borderSoft px-3 py-2 text-sm font-bold text-textMuted hover:bg-slate-50">
            Close
          </button>
        </div>
        <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <MetricBlock title="Total Revenue" value={formatCurrency(bank.totalRevenue)} />
          <MetricBlock title="Discount Cost" value={formatCurrency(bank.discountCost)} />
          <MetricBlock title="Total Transactions" value={formatInteger(bank.totalTransactions)} />
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
              Bank: {formatCurrency(discountEntry?.bankDiscount || 0)} ({(discountEntry?.bankPercent || 0).toFixed(0)}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">
              PVR: {formatCurrency(discountEntry?.pvrDiscount || 0)} ({(discountEntry?.pvrPercent || 0).toFixed(0)}%)
            </p>
            <p className="mt-2 text-sm font-semibold text-textMain">Total: {formatCurrency(discountEntry?.totalDiscount || 0)}</p>
          </div>
        </div>
      </div>
    </div>
  );
}

function OffersByBankModal({ offersByBank, totalOfferCountByBank, expandedOfferBank, onToggleBank, onClose }) {
  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 px-4 backdrop-blur-sm" onClick={onClose}>
      <div
        className="w-full max-w-4xl max-h-[85vh] overflow-y-auto rounded-[2rem] border border-white/60 bg-white p-6 shadow-soft scrollbar-thin"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Info Section</p>
            <h3 className="mt-2 font-display text-2xl font-bold text-textMain">Total Offers by Each Bank</h3>
            <p className="mt-2 text-sm font-semibold text-textMuted">
              {formatInteger(totalOfferCountByBank)} total offers across {formatInteger(offersByBank.length)} banks
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
                                    Bank: {formatCurrency(offer.bankDiscount)}
                                    <br />
                                    PVR: {formatCurrency(offer.pvrDiscount)}
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
                  Upload data to view total offers by bank.
                </div>
              )}
            </div>
          </div>
        </div>
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
        Bank Contribution: {formatCurrency(data.bankDiscount)} ({data.bankPercent.toFixed(1)}%)
      </p>
      <p className="mt-1 text-sm font-semibold text-textMain">
        PVR Contribution: {formatCurrency(data.pvrDiscount)} ({data.pvrPercent.toFixed(1)}%)
      </p>
    </div>
  );
}

function MonthlyTrendTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;

  return (
    <div style={{ borderRadius: "18px", borderColor: "#e2e8f0", background: "#fff" }} className="border p-3 shadow-soft">
      <p className="text-sm font-bold text-textMain">{label}</p>
      <div className="scrollbar-thin mt-2 space-y-1 overflow-y-auto pr-1" style={{ maxHeight: "260px" }}>
        {payload.map((entry) => (
          <p key={entry.dataKey} className="flex items-center gap-2 text-sm font-semibold text-textMain">
            <span className="h-2.5 w-2.5 flex-shrink-0 rounded-full" style={{ backgroundColor: entry.color }} />
            <span className="min-w-0 flex-1 truncate">{entry.name}</span>
            <span>{formatCurrency(entry.value)}</span>
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

export default function App() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [missingColumns, setMissingColumns] = useState([]);
  const [dragActive, setDragActive] = useState(false);
  const [dateFilter, setDateFilter] = useState([]);
  const [bankFilter, setBankFilter] = useState([]);
  const [offerFilter, setOfferFilter] = useState([]);
  const [selectedOffer, setSelectedOffer] = useState(null);
  const [selectedChartBanks, setSelectedChartBanks] = useState([]);
  const [selectedBank, setSelectedBank] = useState(null);
  const [hoveredBank, setHoveredBank] = useState(null);
  const [showOffersByBank, setShowOffersByBank] = useState(false);
  const [expandedOfferBank, setExpandedOfferBank] = useState(null);
  const [comparisonMode, setComparisonMode] = useState("none");
  const [trendMode, setTrendMode] = useState("monthly");
  const [bankSortKey, setBankSortKey] = useState(null);
  const [bankSortDir, setBankSortDir] = useState("desc");
  const [offerSortKey, setOfferSortKey] = useState(null);
  const [offerSortDir, setOfferSortDir] = useState("desc");
  const fileInputRef = useRef(null);

  const banks = useMemo(() => [...new Set(rows.map((row) => row.bankName))].sort(), [rows]);
  const offers = useMemo(() => [...new Set(rows.map((row) => row.offerName))].sort(), [rows]);
  const dates = useMemo(
    () =>
      [...new Set(rows.map((row) => row.dateLabel))].sort((left, right) => {
        const leftDate = rows.find((row) => row.dateLabel === left)?.date?.getTime() || 0;
        const rightDate = rows.find((row) => row.dateLabel === right)?.date?.getTime() || 0;
        return rightDate - leftDate;
      }),
    [rows],
  );

  const latestDataDate = useMemo(
    () => rows.reduce((max, row) => (row.date && (!max || row.date > max) ? row.date : max), null),
    [rows],
  );

  const quickPeriodLabelSets = useMemo(() => {
    if (!latestDataDate) return {};
    return QUICK_PERIODS.reduce((acc, preset) => {
      const range = getQuickPeriodRange(preset.key, latestDataDate);
      const labels = new Set();
      rows.forEach((row) => {
        if (row.date && row.date >= range.rangeStart && row.date <= range.rangeEnd) {
          labels.add(row.dateLabel);
        }
      });
      acc[preset.key] = [...labels].sort();
      return acc;
    }, {});
  }, [latestDataDate, rows]);

  const activeQuickPeriodKey = useMemo(() => {
    if (!dateFilter.length) return "clear";
    const sortedFilter = [...dateFilter].sort();
    const match = QUICK_PERIODS.find((preset) => {
      const labels = quickPeriodLabelSets[preset.key] || [];
      return labels.length === sortedFilter.length && labels.every((label, index) => label === sortedFilter[index]);
    });
    return match ? match.key : null;
  }, [dateFilter, quickPeriodLabelSets]);

  const bankColorMap = useMemo(
    () =>
      banks.reduce((acc, bank, index) => {
        acc[bank] = BANK_COLORS[index % BANK_COLORS.length];
        return acc;
      }, {}),
    [banks],
  );

  useEffect(() => {
    setSelectedChartBanks((current) => current.filter((bank) => banks.includes(bank)));
  }, [banks]);

  useEffect(() => {
    setDateFilter(dates);
  }, [dates]);

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
          setRows([]);
          setError("The uploaded file is empty.");
          return;
        }
        const { parsedRows, missingColumns: missing } = parseWorkbookRows(jsonRows);
        setRows(parsedRows);
        setMissingColumns(missing);
        setSelectedOffer(null);
      } catch {
        setRows([]);
        setError("Unable to parse this file. Please upload a valid Excel sheet with the expected columns.");
      }
    };
    reader.onerror = () => {
      setRows([]);
      setError("There was a problem reading the file.");
    };
    reader.readAsArrayBuffer(file);
  }

  function handleRemoveFile() {
    setRows([]);
    setFileName("");
    setError("");
    setMissingColumns([]);
    setDateFilter([]);
    setBankFilter([]);
    setOfferFilter([]);
    setSelectedOffer(null);
    setSelectedBank(null);
    setSelectedChartBanks([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  function applyDateRange(startStr, endStr) {
    const [startYear, startMonth] = startStr.split("-").map(Number);
    const [endYear, endMonth] = endStr.split("-").map(Number);
    const start = new Date(startYear, startMonth - 1, 1);
    const end = new Date(endYear, endMonth, 0, 23, 59, 59, 999);
    const matching = [...new Set(rows.filter((row) => row.date && row.date >= start && row.date <= end).map((row) => row.dateLabel))];
    setDateFilter(matching);
  }

  const filteredRows = useMemo(
    () =>
      rows.filter((row) => {
        const matchesDate = dateFilter.includes(row.dateLabel);
        const matchesBank = bankFilter.includes(row.bankName);
        const matchesOffer = offerFilter.includes(row.offerName);
        return matchesDate && matchesBank && matchesOffer;
      }),
    [rows, dateFilter, bankFilter, offerFilter],
  );

  const bankOfferFilteredRows = useMemo(
    () =>
      rows.filter((row) => {
        const matchesBank = bankFilter.includes(row.bankName);
        const matchesOffer = offerFilter.includes(row.offerName);
        return matchesBank && matchesOffer;
      }),
    [rows, bankFilter, offerFilter],
  );

  const kpis = useMemo(() => {
    const base = computeKpis(filteredRows);
    return { ...base, recoveryRate: safeRatio(base.netRevenue, base.grossRevenue) };
  }, [filteredRows]);

  const universalData = useMemo(() => getUniversalTransactionsByMonth(rows), [rows]);

  const penetrationData = useMemo(() => {
    const activeMonths = [...new Set(filteredRows.map((row) => row.monthKey).filter((key) => key !== "Unknown"))];
    const missingMonths = activeMonths.filter((key) => !universalData.map.has(key));
    const universalTotal = activeMonths.reduce((sum, key) => sum + (universalData.map.get(key) || 0), 0);
    const percent = universalTotal ? (kpis.totalTransactions / universalTotal) * 100 : null;
    return { percent, universalTotal, missingMonths, hasAnyData: universalData.map.size > 0 };
  }, [filteredRows, universalData, kpis.totalTransactions]);

  const activeDateRangeLabel = useMemo(() => {
    if (!filteredRows.length) return "No data";
    if (!dateFilter.length) {
      const minDate = filteredRows.reduce((min, row) => (row.date && (!min || row.date < min) ? row.date : min), null);
      const maxDate = filteredRows.reduce((max, row) => (row.date && (!max || row.date > max) ? row.date : max), null);
      return `All Dates (${formatDateLabel(minDate)} – ${formatDateLabel(maxDate)})`;
    }
    const selectedDates = filteredRows.map((row) => row.date).filter(Boolean).sort((a, b) => a - b);
    const minDate = selectedDates[0];
    const maxDate = selectedDates[selectedDates.length - 1];
    return minDate.getTime() === maxDate.getTime() ? formatDateLabel(minDate) : `${formatDateLabel(minDate)} – ${formatDateLabel(maxDate)}`;
  }, [filteredRows, dateFilter]);

  const comparisonKpis = useMemo(() => {
    if (comparisonMode === "none") return null;
    const periods = getComparisonPeriods(comparisonMode, latestDataDate);
    if (!periods) return null;
    const { currentStart, currentEnd, priorStart, priorEnd } = periods;

    const currentRows = bankOfferFilteredRows.filter((row) => row.date && row.date >= currentStart && row.date <= currentEnd);
    const priorRows = bankOfferFilteredRows.filter((row) => row.date && row.date >= priorStart && row.date <= priorEnd);

    const current = computeKpis(currentRows);
    const prior = computeKpis(priorRows);

    return {
      current: { ...current, recoveryRate: current.grossRevenue ? (current.netRevenue / current.grossRevenue) * 100 : 0 },
      prior: { ...prior, recoveryRate: prior.grossRevenue ? (prior.netRevenue / prior.grossRevenue) * 100 : 0 },
    };
  }, [comparisonMode, latestDataDate, bankOfferFilteredRows]);
  const extraKpis = useMemo(
    () => ({
      totalBanks: new Set(filteredRows.map((row) => row.bankName)).size,
      totalOffers: new Set(filteredRows.map((row) => row.offerName)).size,
    }),
    [filteredRows],
  );
  const offersByBank = useMemo(() => {
    const grouped = new Map();

    filteredRows.forEach((row) => {
      const bankEntry = grouped.get(row.bankName) || {
        bankName: row.bankName,
        offers: new Map(),
      };

      const offerEntry = bankEntry.offers.get(row.offerName) || {
        offerName: row.offerName,
        startDate: row.date,
        endDate: row.date,
        bankDiscount: 0,
        pvrDiscount: 0,
      };

      if (row.date) {
        if (!offerEntry.startDate || row.date < offerEntry.startDate) {
          offerEntry.startDate = row.date;
        }
        if (!offerEntry.endDate || row.date > offerEntry.endDate) {
          offerEntry.endDate = row.date;
        }
      }

      offerEntry.bankDiscount += row.bankContribution;
      offerEntry.pvrDiscount += row.inoxContribution;

      bankEntry.offers.set(row.offerName, offerEntry);
      grouped.set(row.bankName, bankEntry);
    });

    return [...grouped.values()]
      .map((bankEntry) => {
        const offers = [...bankEntry.offers.values()]
          .map((offer) => ({
            ...offer,
            startLabel: formatIsoDate(offer.startDate),
            endLabel: formatIsoDate(offer.endDate),
          }))
          .sort((left, right) => (left.startDate?.getTime() || 0) - (right.startDate?.getTime() || 0));

        return {
          bankName: bankEntry.bankName,
          offerCount: offers.length,
          offers,
        };
      })
      .sort((left, right) => right.offerCount - left.offerCount || left.bankName.localeCompare(right.bankName));
  }, [filteredRows]);
  const totalOfferCountByBank = useMemo(
    () => offersByBank.reduce((sum, bank) => sum + bank.offerCount, 0),
    [offersByBank],
  );

  const bankRows = useMemo(() => aggregateBanks(filteredRows), [filteredRows]);
  const offerRows = useMemo(() => aggregateOffers(filteredRows), [filteredRows]);

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
  const monthlySeries = useMemo(() => aggregateMonthlySeries(rows, selectedChartBanks), [rows, selectedChartBanks]);
  const seasonalData = useMemo(() => aggregateSeasonalByYear(rows, selectedChartBanks), [rows, selectedChartBanks]);
  const seasonalYears = useMemo(() => Object.keys(seasonalData[0] || {}).filter((key) => key !== "month" && key !== "monthLabel"), [seasonalData]);
  const yearlyData = useMemo(() => aggregateYearlyTotals(rows, selectedChartBanks), [rows, selectedChartBanks]);
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
    return formatCurrency(amount);
  }

  function toggleSelectedBank(bankName) {
    setSelectedBank((current) => (current === bankName ? null : bankName));
  }

  function toggleExpandedOfferBank(bankName) {
    setExpandedOfferBank((current) => (current === bankName ? null : bankName));
  }

  function applyQuickPeriod(rangeStart, rangeEnd) {
    const labels = new Set();
    rows.forEach((row) => {
      if (row.date && row.date >= rangeStart && row.date <= rangeEnd) {
        labels.add(row.dateLabel);
      }
    });
    setDateFilter([...labels]);
  }

  const selectedBankRow = useMemo(() => bankRows.find((bank) => bank.bankName === selectedBank) || null, [bankRows, selectedBank]);
  const selectedBankOffersEntry = useMemo(() => offersByBank.find((entry) => entry.bankName === selectedBank) || null, [offersByBank, selectedBank]);
  const selectedBankDiscountEntry = useMemo(() => discountData.find((entry) => entry.bankName === selectedBank) || null, [discountData, selectedBank]);

  return (
    <div className="h-full overflow-hidden bg-appBg text-textMain">
      <div className="mx-auto flex h-full max-w-[1800px] flex-col gap-5 overflow-y-auto px-4 py-4 scrollbar-thin sm:px-6 lg:px-8">
        <header className="grid gap-4 lg:grid-cols-[1.4fr_1fr]">
          <div className="rounded-[2rem] border border-white/60 bg-white/80 p-6 shadow-soft backdrop-blur">
            <img src={pvrInoxLogo} alt="PVR INOX" className="mb-3 h-9 w-auto" />
            <p className="text-xs font-bold uppercase tracking-[0.32em] text-accentBlue">Revenue Intelligence</p>
            <h1 className="mt-3 font-display text-3xl font-bold text-textMain sm:text-4xl">BANK OFFERS PERFORMANCE</h1>
            <p className="mt-3 max-w-2xl text-sm font-medium leading-6 text-textMuted">

            </p>
            <div className="mt-4">
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Quick Period</p>
              <div className="mt-2 flex flex-wrap gap-2">
                {QUICK_PERIODS.map((preset) => (
                  <button
                    key={preset.key}
                    type="button"
                    disabled={!latestDataDate}
                    onClick={() => {
                      const range = getQuickPeriodRange(preset.key, latestDataDate);
                      if (range) applyQuickPeriod(range.rangeStart, range.rangeEnd);
                    }}
                    className={`rounded-full border px-4 py-2 text-xs font-bold uppercase tracking-wide transition disabled:cursor-not-allowed disabled:opacity-40 ${
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
                  onClick={() => setDateFilter([])}
                  className={`rounded-full border px-4 py-2 text-xs font-bold uppercase tracking-wide transition ${
                    activeQuickPeriodKey === "clear"
                      ? "border-accentBlue bg-accentBlue text-white"
                      : "border-borderSoft bg-white text-textMuted hover:bg-slate-50"
                  }`}
                >
                  Clear
                </button>
              </div>
            </div>
            {penetrationData.hasAnyData ? (
              <div className="mt-4 rounded-2xl border border-borderSoft bg-white/70 p-4">
                <div className="flex items-start justify-between gap-3">
                  <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Bank Offer Penetration</p>
                  <p className="text-xs font-extrabold text-textMain">{activeDateRangeLabel}</p>
                </div>
                {penetrationData.missingMonths.length ? (
                  <p className="mt-2 text-sm font-semibold text-amber-700">
                    Universal transaction data missing for: {penetrationData.missingMonths.map(formatMonthKeyLabel).join(", ")}
                  </p>
                ) : penetrationData.percent !== null ? (
                  <>
                    <p className="mt-2 text-[28px] font-bold leading-none text-textMain">{penetrationData.percent.toFixed(1)}%</p>
                    <p className="mt-2 text-sm font-medium text-textMuted">
                      {formatInCrore(kpis.totalTransactions)} partnered / {formatInCrore(penetrationData.universalTotal)} PVR INOX Total Transaction
                    </p>
                  </>
                ) : (
                  <p className="mt-2 text-sm font-semibold text-textMuted">No transactions in the selected period.</p>
                )}
                {universalData.inconsistent.size > 0 ? (
                  <p className="mt-2 text-xs font-semibold text-rose-600">
                    Inconsistent universal totals found for: {[...universalData.inconsistent].map(formatMonthKeyLabel).join(", ")}
                  </p>
                ) : null}
              </div>
            ) : (
              <p className="mt-4 text-xs font-medium text-textMuted">
                Add a "Universal Transactions" column to your file to see penetration %.
              </p>
            )}
            {missingColumns.length ? (
              <div className="mt-4 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm font-semibold text-amber-700">
                Missing columns were handled as zero values: {missingColumns.join(", ")}
              </div>
            ) : null}
          </div>
          <div className="grid gap-4">
            <UploadPanel
              onFileChange={handleFileChange}
              dragActive={dragActive}
              setDragActive={setDragActive}
              fileName={fileName}
              error={error}
              onRemoveFile={handleRemoveFile}
              inputRef={fileInputRef}
            />
            <div className="rounded-[2rem] border border-white/60 bg-white/90 p-4 shadow-soft">
              <div className="flex flex-wrap items-end gap-3">
                <DateMultiSelectDropdown
                  options={dates}
                  selected={dateFilter}
                  onToggle={(date) => {
                    setDateFilter((current) => (current.includes(date) ? current.filter((item) => item !== date) : [...current, date]));
                  }}
                  onClear={() => {
                    console.log("[Clear audit] Date filter cleared, new value:", []);
                    setDateFilter([]);
                  }}
                  onApplyRange={applyDateRange}
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
        </header>

        <section className="grid gap-5 lg:grid-cols-[1fr_1fr] items-start">
          <div className="grid h-[160px] gap-4 md:grid-cols-2">
            <StatCard title="Total Banks" value={formatInteger(extraKpis.totalBanks)} subtitle="Unique bank partners" color={KPI_COLORS[7]} icon="TB" />
            <StatCard title="Total Offers" value={formatInteger(extraKpis.totalOffers)} subtitle="Active offers" color={KPI_COLORS[7]} icon="TO" />
          </div>

          <div className="h-[160px] rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <button
              type="button"
              onClick={() => setShowOffersByBank(true)}
              className="flex w-full items-center justify-between gap-3 text-left"
            >
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Info Section</p>
                <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Total Offers by Each Bank</h2>
                <p className="mt-2 text-sm font-semibold text-textMuted">
                  {formatInteger(totalOfferCountByBank)} total offers across {formatInteger(offersByBank.length)} banks
                </p>
              </div>
              <span className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-2xl border border-borderSoft bg-slate-50 text-lg font-bold text-textMuted">
                →
              </span>
            </button>
          </div>
        </section>

        <section className="flex flex-col gap-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Key Metrics</p>
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
          </div>
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
            <StatCard
              title="Total Transactions"
              value={formatInCrore(kpis.totalTransactions)}
              color={KPI_COLORS[0]}
              icon="TX"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.totalTransactions, comparisonKpis.prior.totalTransactions), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
            <StatCard
              title="Gross Revenue"
              value={formatCompactCurrency(kpis.grossRevenue)}
              color={KPI_COLORS[0]}
              icon="GR"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.grossRevenue, comparisonKpis.prior.grossRevenue), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
            <StatCard
              title="Net Revenue"
              value={formatCompactCurrency(kpis.netRevenue)}
              color={KPI_COLORS[0]}
              icon="NR"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.netRevenue, comparisonKpis.prior.netRevenue), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
            <StatCard
              title="Total Discount Given"
              value={formatCompactCurrency(kpis.totalDiscount)}
              color={KPI_COLORS[0]}
              icon="DG"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.totalDiscount, comparisonKpis.prior.totalDiscount), label: COMPARISON_LABELS[comparisonMode] } : undefined}
              extra={
                <>
                  <p className="text-xs font-semibold text-slate-500">
                    Bank: {formatCurrency(overallSplit.totalBank)} ({overallSplit.bankPercent.toFixed(0)}%)
                  </p>
                  <p className="text-xs font-semibold text-slate-500">
                    PVR: {formatCurrency(overallSplit.totalPvr)} ({overallSplit.pvrPercent.toFixed(0)}%)
                  </p>
                </>
              }
            />
            <StatCard
              title="Total Discount Rate"
              value={`${kpis.recoveryRate.toFixed(1)}%`}
              color={KPI_COLORS[0]}
              icon="RR"
              delta={comparisonKpis ? { value: computeDelta(comparisonKpis.current.recoveryRate, comparisonKpis.prior.recoveryRate), label: COMPARISON_LABELS[comparisonMode] } : undefined}
            />
          </div>
        </section>

        <section className="grid gap-5 xl:grid-cols-[1.65fr_1fr]">
          <div className="flex h-[620px] flex-col overflow-hidden rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <div className="flex items-center justify-between gap-3">
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Bank Scorecard</p>
                <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Bank Performance Overview</h2>
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
            <div className="mt-4 h-[512px] overflow-hidden rounded-3xl border border-borderSoft">
              <div className="h-full overflow-y-scroll overflow-x-auto scrollbar-thin">
                <table className="min-w-full table-fixed divide-y divide-borderSoft text-sm">
                  <thead className="sticky top-0 bg-slate-50/95 backdrop-blur">
                    <tr>
                      <th className="w-[18%] px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">
                        Bank
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
                          ({formatInCrore(kpis.totalTransactions)})
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
                          ({formatInCrore(kpis.grossRevenue)})
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
                          ({formatInCrore(kpis.totalDiscount)})
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
                          <td className="px-4 py-3 font-bold text-textMain">{bank.bankName}</td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            {formatInteger(bank.totalTransactions)}
                            <span className="ml-1 text-xs font-bold text-textMuted">
                              ({kpis.totalTransactions ? ((bank.totalTransactions / kpis.totalTransactions) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            <span className="whitespace-nowrap">{formatInCrore(bank.totalRevenue)}</span>
                            <span className="mt-0.5 block text-xs font-bold text-textMuted">
                              ({kpis.grossRevenue ? ((bank.totalRevenue / kpis.grossRevenue) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 font-semibold text-textMain">
                            <span className="whitespace-nowrap">{formatInCrore(bank.discountCost)}</span>
                            <span className="mt-0.5 block text-xs font-bold text-textMuted">
                              ({kpis.totalDiscount ? ((bank.discountCost / kpis.totalDiscount) * 100).toFixed(1) : "0.0"}%)
                            </span>
                          </td>
                          <td className="px-4 py-3 text-xs font-semibold text-textMain">
                            {split ? (
                              <div className="flex flex-col gap-0.5">
                                <span className="text-accentBlue">Bank: {formatInCrore(split.bankDiscount)} ({split.bankPercent.toFixed(0)}%)</span>
                                <span className="text-accentGreen">PVR: {formatInCrore(split.pvrDiscount)} ({split.pvrPercent.toFixed(0)}%)</span>
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

          <div className="flex h-[620px] flex-col overflow-hidden rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Discount Split</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Discount Distribution - Bank vs PVR</h2>
            <div className="mt-4 h-[512px] overflow-hidden rounded-3xl bg-appBg p-3">
              {discountData.length ? (
                <ResponsiveContainer width="100%" height={512}>
                  <BarChart
                    data={discountData}
                    layout="vertical"
                    margin={{ top: 10, right: 24, left: 0, bottom: 10 }}
                    barCategoryGap="26%"
                    barGap={6}
                  >
                    <CartesianGrid stroke="#cbd5e1" strokeDasharray="4 4" />
                    <XAxis type="number" stroke="#718096" tickFormatter={formatCompactCurrency} allowDecimals={false} tickCount={6} />
                    <YAxis
                      type="category"
                      dataKey="bankName"
                      stroke="#718096"
                      width={120}
                      tick={{ fontSize: 11 }}
                    />
                    <Tooltip allowEscapeViewBox={{ x: false, y: false }} content={<DiscountSplitTooltip />} />
                    <Legend wrapperStyle={{ paddingTop: "20px" }} />
                    <Bar dataKey="bankDiscount" stackId="a" fill="#2563eb" name="Bank Contribution" minPointSize={4} barSize={14}>
                      <LabelList dataKey="bankPercent" position="inside" formatter={(v) => `${v.toFixed(0)}%`} style={{ fontSize: 10, fill: "#fff", fontWeight: 600 }} />
                    </Bar>
                    <Bar dataKey="pvrDiscount" stackId="a" fill="#10b981" name="PVR Contribution" minPointSize={4} barSize={14}>
                      <LabelList dataKey="pvrPercent" position="inside" formatter={(v) => `${v.toFixed(0)}%`} style={{ fontSize: 10, fill: "#fff", fontWeight: 600 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Upload data to view discount split by bank.
                </div>
              )}
            </div>
          </div>
        </section>

        <section className="grid gap-5 xl:grid-cols-[1.65fr_1fr]">
          <div className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Channel Mix</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Revenue by Channel</h2>
            <div className="relative mt-4 h-[340px] overflow-hidden rounded-3xl bg-appBg p-3">
              <span className="absolute right-3 top-3 z-10 rounded-full bg-slate-100 px-3 py-1 text-xs font-bold text-textMuted">
                Total: {formatInCrore(channelData.reduce((sum, c) => sum + c.value, 0))}
              </span>
              {channelData.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={channelData}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      outerRadius={110}
                      label={({ name, value, percent }) => `${name}: ${formatCompactCurrency(value)} (${(percent * 100).toFixed(0)}%)`}
                    >
                      {channelData.map((entry) => (
                        <Cell key={entry.name} fill={CHANNEL_COLORS[entry.name] || CHANNEL_FALLBACK_COLOR} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => formatCurrency(value)} contentStyle={{ borderRadius: "18px", borderColor: "#e2e8f0" }} />
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
            <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Discount Split</p>
            <h2 className="mt-1 font-display text-2xl font-bold text-textMain">Overall Bank vs PVR Contribution</h2>
            <div className="mt-6 flex flex-col gap-6">
              <div>
                <div className="flex items-center justify-between text-sm font-bold text-textMain">
                  <span className="text-accentBlue">Bank: {formatCurrency(overallSplit.totalBank)} ({overallSplit.bankPercent.toFixed(0)}%)</span>
                  <span className="text-accentGreen">PVR: {formatCurrency(overallSplit.totalPvr)} ({overallSplit.pvrPercent.toFixed(0)}%)</span>
                </div>
                <div className="mt-3 flex h-8 w-full overflow-hidden rounded-full bg-slate-100">
                  <div className="h-full bg-accentBlue transition-all" style={{ width: `${overallSplit.bankPercent}%` }} />
                  <div className="h-full bg-accentGreen transition-all" style={{ width: `${overallSplit.pvrPercent}%` }} />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="rounded-2xl bg-blue-50 p-4">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-accentBlue">Bank Contribution</p>
                  <p className="mt-2 text-2xl font-bold text-textMain">{formatCompactCurrency(overallSplit.totalBank)}</p>
                  <p className="mt-1 text-sm font-semibold text-textMuted">{overallSplit.bankPercent.toFixed(1)}% of total discount</p>
                </div>
                <div className="rounded-2xl bg-emerald-50 p-4">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-accentGreen">PVR Contribution</p>
                  <p className="mt-2 text-2xl font-bold text-textMain">{formatCompactCurrency(overallSplit.totalPvr)}</p>
                  <p className="mt-1 text-sm font-semibold text-textMuted">{overallSplit.pvrPercent.toFixed(1)}% of total discount</p>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="rounded-[2rem] border border-white/60 bg-white/90 p-5 shadow-soft">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Offer Directory</p>
              <h2 className="mt-1 font-display text-2xl font-bold text-textMain">All offers grouped by bank</h2>
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
                      Bank
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
                        ({formatInCrore(kpis.totalTransactions)})
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
                        ({formatInCrore(kpis.grossRevenue)})
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
                        ({formatInCrore(kpis.totalDiscount)})
                      </span>
                    </th>
                    <th className="px-4 py-3 text-left font-bold uppercase tracking-[0.18em] text-textMuted">Discount %</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedOfferRows.length ? sortedOfferRows.map((offer, index) => (
                    <tr key={`${offer.offerName}-${offer.bankName}`} className={`${index % 2 === 0 ? "bg-white" : "bg-slate-50/60"} cursor-pointer transition hover:bg-blue-50/70`} onClick={() => setSelectedOffer(offer)}>
                      <td className="px-4 py-3 font-semibold text-textMain">{offer.bankName}</td>
                      <td className="px-4 py-3 font-bold text-textMain">{offer.offerName}</td>
                      <td className="px-4 py-3 font-semibold text-textMain">{formatInteger(offer.transactions)}</td>
                      <td className="px-4 py-3 font-semibold text-textMain">
                        {formatInCrore(offer.revenue)}
                        <span className="ml-1 text-xs font-bold text-textMuted">
                          ({kpis.grossRevenue ? ((offer.revenue / kpis.grossRevenue) * 100).toFixed(2) : "0.00"}%)
                        </span>
                      </td>
                      <td className="px-4 py-3 font-semibold text-textMain">
                        {formatInCrore(offer.discount)}
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
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Month-wise Bank Performance</p>
              <h2 className="mt-1 font-display text-2xl font-bold text-textMain">
                {trendMode === "monthly"
                  ? "Monthly revenue trend by selected banks"
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
              <span className="text-xs font-bold uppercase tracking-[0.24em] text-textMuted">Independent Filter</span>
              <MultiSelectDropdown
                options={banks}
                selected={selectedChartBanks}
                onToggle={(bank) => {
                  setSelectedChartBanks((current) => (current.includes(bank) ? current.filter((item) => item !== bank) : [...current, bank]));
                }}
                onClear={() => setSelectedChartBanks([])}
                label="bank"
              />
            </div>
          </div>
          <div className="mt-4 h-[320px] overflow-hidden rounded-3xl border border-borderSoft bg-white">
            {trendMode === "monthly" ? (
              monthlySeries.length && selectedChartBanks.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={monthlySeries} margin={{ top: 10, right: 20, left: 0, bottom: 10 }}>
                    <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                    <XAxis dataKey="monthKey" stroke="#718096" />
                    <YAxis stroke="#718096" tickFormatter={formatCompactCurrency} />
                    <Tooltip allowEscapeViewBox={{ x: false, y: true }} content={<MonthlyTrendTooltip />} />
                    <Legend />
                    {selectedChartBanks.map((bank) => (
                      <Line key={bank} type="monotone" dataKey={bank} stroke={bankColorMap[bank]} strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 6 }} connectNulls />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Upload data and select one or more banks to view the month-wise revenue chart.
                </div>
              )
            ) : trendMode === "mom" ? (
              seasonalData.length && selectedChartBanks.length ? (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={seasonalData} margin={{ top: 10, right: 20, left: 0, bottom: 10 }}>
                    <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                    <XAxis dataKey="monthLabel" stroke="#718096" />
                    <YAxis stroke="#718096" tickFormatter={formatCompactCurrency} />
                    <Tooltip allowEscapeViewBox={{ x: false, y: true }} content={<MonthlyTrendTooltip />} />
                    <Legend />
                    {seasonalYears.map((year, index) => (
                      <Line key={year} type="monotone" dataKey={year} stroke={BANK_COLORS[index % BANK_COLORS.length]} strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 6 }} connectNulls />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                  Select at least one bank to view the month-on-month seasonality chart.
                </div>
              )
            ) : yearlyData.length && selectedChartBanks.length ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={yearlyData} margin={{ top: 24, right: 20, left: 0, bottom: 10 }}>
                  <CartesianGrid stroke="#e2e8f0" strokeDasharray="4 4" />
                  <XAxis dataKey="year" stroke="#718096" />
                  <YAxis stroke="#718096" tickFormatter={formatCompactCurrency} />
                  <Tooltip formatter={(value) => formatCurrency(value)} contentStyle={{ borderRadius: "18px", borderColor: "#e2e8f0" }} />
                  <Legend />
                  <Bar dataKey="revenue" name="Revenue" fill="#2563eb" radius={[8, 8, 0, 0]}>
                    <LabelList dataKey="revenue" position="top" formatter={(value) => formatCompactCurrency(value)} style={{ fontSize: 12, fontWeight: 700, fill: "#1a202c" }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="flex h-full items-center justify-center rounded-3xl border border-dashed border-borderSoft bg-slate-50 text-center text-sm font-semibold text-textMuted">
                Select at least one bank to view the year-on-year revenue chart.
              </div>
            )}
          </div>
        </section>
      </div>

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
    </div>
  );
}
