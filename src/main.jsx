import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  ArrowDownToLine,
  Check,
  ChevronRight,
  CircleAlert,
  FileCheck2,
  FileSpreadsheet,
  FolderOpen,
  Layers3,
  Loader2,
  Play,
  RefreshCw,
  Settings2,
  UploadCloud,
} from "lucide-react";
import "./styles.css";

const CATEGORIES = ["CME", "CMS", "LMA", "YLC", "Others"];
const INVOICE_CATEGORY_ORDER = ["CMS", "CME", "LMA", "YLC", "Others"];
const DISPLAY_TABS = ["All", "Matched", "Amount mismatch", "Aggregator only", "API only"];

function normalize(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return XLSX.SSF.format("yyyy-mm-dd hh:mm:ss", value);
  let text = String(value).trim();
  if (/^\d+\.0$/.test(text)) text = text.slice(0, -2);
  return text;
}

function parseDateString(dateStr) {
  if (!dateStr) return null;
  let match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{2}):(\d{2}):(\d{2}))?/);
  if (match) {
    const [_, y, m, d, hh = 0, mm = 0, ss = 0] = match.map(Number);
    return new Date(y, m - 1, d, hh, mm, ss);
  }
  match = dateStr.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s+(\d{2}):(\d{2}):(\d{2}))?/);
  if (match) {
    const [_, d, m, y, hh = 0, mm = 0, ss = 0] = match.map(Number);
    return new Date(y, m - 1, d, hh, mm, ss);
  }
  const parsed = new Date(dateStr);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function toNumber(value) {
  if (value === undefined || value === null || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function workbookRows(workbook, sheetName, headerRowIndex) {
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true,
    range: headerRowIndex,
  });
}

function sheetMatrix(workbook, sheetName) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
}

function scoreHeaderRow(row) {
  const cells = row.map((cell) => normalize(cell).toLowerCase()).filter(Boolean);
  const keywords = ["ref", "amount", "date", "order", "invoice", "category", "dept", "total"];
  const hits = keywords.reduce((count, keyword) => {
    return count + (cells.some((cell) => cell.includes(keyword)) ? 1 : 0);
  }, 0);
  return cells.length + hits * 6;
}

function detectHeaderRow(workbook, sheetName) {
  const matrix = sheetMatrix(workbook, sheetName).slice(0, 15);
  let bestIndex = 0;
  let bestScore = -1;
  matrix.forEach((row, index) => {
    const score = scoreHeaderRow(row);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestIndex;
}

async function readWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  return {
    id: `${file.name}-${file.lastModified}-${file.size}`,
    name: file.name,
    size: file.size,
    workbook,
    sheetNames: workbook.SheetNames,
  };
}

function inferCategory(fileName, row = {}) {
  const haystack = `${fileName} ${Object.values(row).join(" ")}`.toUpperCase();
  if (haystack.includes("CMS") || haystack.includes("MAT")) return "CMS";
  if (haystack.includes("CME")) return "CME";
  if (haystack.includes("LMA")) return "LMA";
  if (haystack.includes("YLC")) return "YLC";
  return "Others";
}

function rowObject(headers, values) {
  const object = {};
  headers.forEach((header, index) => {
    const key = normalize(header);
    if (!key || object[key] === undefined) object[key] = values[index] ?? "";
  });
  return object;
}

function trimTrailingBlankCells(values) {
  const trimmed = [...values];
  while (trimmed.length && !normalize(trimmed[trimmed.length - 1])) trimmed.pop();
  return trimmed;
}

function parseAggregatorFiles(files) {
  const rows = [];
  const sources = [];

  files.forEach((file) => {
    file.sheetNames.forEach((sheetName) => {
      const headerRow = detectHeaderRow(file.workbook, sheetName);
      const data = workbookRows(file.workbook, sheetName, headerRow).filter((row) => {
        return Object.values(row).some((value) => normalize(value));
      });
      const lowerFile = file.name.toLowerCase();
      const lowerSheet = sheetName.toLowerCase();
      const isHdfc = lowerFile.includes("hdfc") || data.some((row) => row["CCAvenue Ref#"]);
      const isBillDesk = lowerFile.includes("billdesk") || lowerSheet.includes("aimaothers") || data.some((row) => row["PGI Ref. No."]);

      data.forEach((row, index) => {
        if (isHdfc && lowerSheet.includes("mat")) {
          rows.push({
            id: `agg-${file.id}-${sheetName}-${index}`,
            source: "HDFC",
            sourceFile: file.name,
            sheetName,
            sourceRow: index + headerRow + 2,
            category: "CMS",
            primaryKey: normalize(row["CCAvenue Ref#"]),
            secondaryKey: normalize(row["Order No"]),
            date: normalize(row["Order Datetime"]),
            grossAmount: toNumber(row["Order Amount"] || row["Gross Amount"]),
            charges: "",
            gst: "",
            netAmount: "",
            settlementDate: normalize(row["Order Stlmt Date"]),
            statusText: normalize(row["Order Status"]),
            matchAttempts: [
              ["referenceNo", normalize(row["CCAvenue Ref#"]), "HDFC CCAvenue Ref# = API Reference_No"],
              ["transactionNo", normalize(row["Order No"]), "HDFC Order No = API Transaction_No"],
            ],
          });
          return;
        }

        if (isHdfc) {
          rows.push({
            id: `agg-${file.id}-${sheetName}-${index}`,
            source: "HDFC",
            sourceFile: file.name,
            sheetName,
            sourceRow: index + headerRow + 2,
            category: "CME",
            primaryKey: normalize(row["Order No"]),
            secondaryKey: normalize(row["CCAvenue Ref#"]),
            date: normalize(row["Order Datetime"]),
            grossAmount: toNumber(row["Order Amount"] || row["Gross Amount"]),
            charges: "",
            gst: "",
            netAmount: "",
            settlementDate: normalize(row["Order Stlmt Date"]),
            statusText: normalize(row["Order Status"]),
            matchAttempts: [
              ["referenceNo", normalize(row["Order No"]), "HDFC Order No = API Reference_No"],
              ["transactionNo", normalize(row["CCAvenue Ref#"]), "HDFC CCAvenue Ref# = API Transaction_No"],
            ],
          });
          return;
        }

        if (isBillDesk) {
          const category = normalize(row.Dept) || inferCategory(file.name, row);
          rows.push({
            id: `agg-${file.id}-${sheetName}-${index}`,
            source: "BillDesk",
            sourceFile: file.name,
            sheetName,
            sourceRow: index + headerRow + 2,
            category,
            primaryKey: normalize(row["Ref. 1"]),
            secondaryKey: normalize(row["PGI Ref. No."]),
            date: normalize(row["Date of Txn"]),
            grossAmount: toNumber(row["|"] || row.Amount || row["Gross Amount"]),
            charges: toNumber(row.Charges),
            gst: toNumber(row.GST),
            netAmount: toNumber(row["Net Amount"]),
            settlementDate: normalize(row["Settlement Date"]),
            statusText: normalize(row.STATUS),
            matchAttempts: [
              ["referenceNo", normalize(row["Ref. 1"]), "BillDesk Ref. 1 = API Reference_No"],
              ["transactionNo", normalize(row["PGI Ref. No."]), "BillDesk PGI Ref. No. = API Transaction_No"],
              ["referenceNo", normalize(row["PGI Ref. No."]), "BillDesk PGI Ref. No. = API Reference_No"],
              ["transactionNo", normalize(row["Ref. 1"]), "BillDesk Ref. 1 = API Transaction_No"],
            ],
          });
        }
      });

      sources.push({
        file: file.name,
        sheet: sheetName,
        headerRow: headerRow + 1,
        rows: data.length,
        type: isHdfc ? "HDFC" : isBillDesk ? "BillDesk" : "Unknown",
      });
    });
  });

  return { rows, sources };
}

function parsePiFiles(files, categoryOverrides) {
  const rows = [];
  const sources = [];

  files.forEach((file) => {
    file.sheetNames.forEach((sheetName) => {
      const headerRow = detectHeaderRow(file.workbook, sheetName);
      const matrix = sheetMatrix(file.workbook, sheetName);
      const headers = trimTrailingBlankCells(matrix[headerRow] || []);
      const data = matrix
        .slice(headerRow + 1)
        .map((values, offset) => ({ values, offset, row: rowObject(headers, values) }))
        .filter(({ values }) => values.some((value) => normalize(value)));
      const firstRow = data[0] || {};
      const override = categoryOverrides[file.id];
      const category = override || inferCategory(file.name, firstRow.row || {});
      data.forEach(({ row, values, offset }) => {
        rows.push({
          id: `pi-${file.id}-${sheetName}-${offset}`,
          category,
          sourceFile: file.name,
          sheetName,
          sourceRow: offset + headerRow + 2,
          referenceNo: normalize(row.Reference_No),
          transactionNo: normalize(row.Transaction_No),
          invoiceNo: normalize(row["Invoice No"]),
          date: normalize(row.Transaction_Date),
          candidate: normalize(row["Name of the Candidate"]),
          totalAmount: toNumber(row["Total Amt"]),
          incomeAmount: toNumber(row["Income Amount"]),
          costCentre: normalize(row["Cost Centre"] || row["Project Code"]),
          tallyHeaders: headers,
          tallyValues: headers.map((_, index) => values[index] ?? ""),
        });
      });
      sources.push({
        file: file.name,
        sheet: sheetName,
        category,
        headerRow: headerRow + 1,
        rows: data.length,
      });
    });
  });

  return { rows, sources };
}

function queueLookup(piRows) {
  const lookup = new Map();
  piRows.forEach((row, index) => {
    [
      ["referenceNo", row.referenceNo],
      ["transactionNo", row.transactionNo],
    ].forEach(([field, value]) => {
      if (!value) return;
      const key = `${row.category}::${field}::${value}`;
      if (!lookup.has(key)) lookup.set(key, []);
      lookup.get(key).push(index);
    });
  });
  return lookup;
}

function reconcileRows(aggregatorRows, piRows, startDateStr, endDateStr) {
  let filteredAgg = aggregatorRows;
  let filteredPi = piRows;

  if (startDateStr || endDateStr) {
    let startLimit = null;
    let endLimit = null;

    if (startDateStr) {
      const [y, m, d] = startDateStr.split('-').map(Number);
      startLimit = new Date(y, m - 1, d, 0, 0, 0, 0);
    }
    if (endDateStr) {
      const [y, m, d] = endDateStr.split('-').map(Number);
      endLimit = new Date(y, m - 1, d, 23, 59, 59, 999);
    }

    const filterFn = (row) => {
      if (!row.date) return false;
      const d = parseDateString(row.date);
      if (!d) return false;
      if (startLimit && d < startLimit) return false;
      if (endLimit && d > endLimit) return false;
      return true;
    };

    filteredAgg = aggregatorRows.filter(filterFn);
    filteredPi = piRows.filter(filterFn);
  }

  const lookup = queueLookup(filteredPi);
  const usedPi = new Set();
  const matched = [];
  const aggregatorOnly = [];

  filteredAgg.forEach((agg) => {
    const candidates = [];
    agg.matchAttempts.forEach(([field, value, rule]) => {
      if (!value) return;
      const key = `${agg.category}::${field}::${value}`;
      (lookup.get(key) || []).forEach((piIndex) => {
        if (!usedPi.has(piIndex)) {
          const pi = filteredPi[piIndex];
          candidates.push({
            piIndex,
            pi,
            rule,
            amountDelta: Math.abs(agg.grossAmount - pi.totalAmount),
          });
        }
      });
    });

    if (!candidates.length) {
      aggregatorOnly.push({
        status: "Aggregator only",
        category: agg.category,
        aggregatorSource: agg.source,
        aggregatorFile: agg.sourceFile,
        aggregatorSheet: agg.sheetName,
        aggregatorRow: agg.sourceRow,
        aggregatorKey: agg.primaryKey,
        aggregatorAltKey: agg.secondaryKey,
        aggregatorDate: agg.date,
        aggregatorAmount: agg.grossAmount,
        charges: agg.charges,
        gst: agg.gst,
        netAmount: agg.netAmount,
        settlementDate: agg.settlementDate,
        gatewayStatus: agg.statusText,
        remarks: "Not found in API data",
      });
      return;
    }

    candidates.sort((a, b) => a.amountDelta - b.amountDelta);
    const selected = candidates[0];
    usedPi.add(selected.piIndex);
    const pi = selected.pi;
    const difference = Number((agg.grossAmount - pi.totalAmount).toFixed(2));
    matched.push({
      status: Math.abs(difference) <= 0.01 ? "Matched" : "Amount mismatch",
      category: agg.category,
      matchRule: selected.rule,
      amountDifference: difference,
      aggregatorSource: agg.source,
      aggregatorFile: agg.sourceFile,
      aggregatorSheet: agg.sheetName,
      aggregatorRow: agg.sourceRow,
      aggregatorKey: agg.primaryKey,
      aggregatorAltKey: agg.secondaryKey,
      aggregatorDate: agg.date,
      aggregatorAmount: agg.grossAmount,
      charges: agg.charges,
      gst: agg.gst,
      netAmount: agg.netAmount,
      settlementDate: agg.settlementDate,
      gatewayStatus: agg.statusText,
      piFile: pi.sourceFile,
      piSheet: pi.sheetName,
      piRow: pi.sourceRow,
      piReferenceNo: pi.referenceNo,
      piTransactionNo: pi.transactionNo,
      piDate: pi.date,
      piAmount: pi.totalAmount,
      candidate: pi.candidate,
      invoiceNo: pi.invoiceNo,
      costCentre: pi.costCentre,
      tallyHeaders: pi.tallyHeaders,
      tallyValues: pi.tallyValues,
      remarks: Math.abs(difference) <= 0.01 ? "" : "Gross amount differs from API total",
    });
  });

  const piOnly = filteredPi
    .filter((_, index) => !usedPi.has(index))
    .map((pi) => ({
      status: "API only",
      category: pi.category,
      piFile: pi.sourceFile,
      piSheet: pi.sheetName,
      piRow: pi.sourceRow,
      piReferenceNo: pi.referenceNo,
      piTransactionNo: pi.transactionNo,
      piDate: pi.date,
      piAmount: pi.totalAmount,
      candidate: pi.candidate,
      invoiceNo: pi.invoiceNo,
      costCentre: pi.costCentre,
      remarks: "Not found in aggregator data",
    }));

  const all = [...matched, ...aggregatorOnly, ...piOnly];
  const summary = CATEGORIES.slice(0, 4).map((category) => {
    const aggForCategory = filteredAgg.filter((row) => row.category === category);
    const piForCategory = filteredPi.filter((row) => row.category === category);
    const matchedForCategory = matched.filter((row) => row.category === category);
    const mismatches = matchedForCategory.filter((row) => row.status === "Amount mismatch");
    const aggOnly = aggregatorOnly.filter((row) => row.category === category);
    const piMissing = piOnly.filter((row) => row.category === category);
    return {
      category,
      aggregatorRecords: aggForCategory.length,
      piRecords: piForCategory.length,
      matchedRecords: matchedForCategory.length,
      amountMismatches: mismatches.length,
      aggregatorOnly: aggOnly.length,
      piOnly: piMissing.length,
      aggregatorGrossTotal: sum(aggForCategory, "grossAmount"),
      piTotal: sum(piForCategory, "totalAmount"),
      matchedAggGross: sum(matchedForCategory, "aggregatorAmount"),
      matchedPiTotal: sum(matchedForCategory, "piAmount"),
      unmatchedAggGross: sum(aggOnly, "aggregatorAmount"),
      unmatchedPiTotal: sum(piMissing, "piAmount"),
    };
  });

  return {
    all,
    matched,
    amountMismatches: matched.filter((row) => row.status === "Amount mismatch"),
    aggregatorOnly,
    piOnly,
    summary,
  };
}

function sum(rows, field) {
  return Number(rows.reduce((total, row) => total + toNumber(row[field]), 0).toFixed(2));
}

function money(value) {
  return new Intl.NumberFormat("en-IN", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 2,
  }).format(toNumber(value));
}

function integer(value) {
  return new Intl.NumberFormat("en-IN").format(value || 0);
}

function exportTallyOnlyReport(result) {
  const workbook = XLSX.utils.book_new();
  const invoiceRows = buildInvoiceSheetRows(result.matched.filter((row) => row.status === "Matched"));
  const sheet = Array.isArray(invoiceRows[0]) ? XLSX.utils.aoa_to_sheet(invoiceRows) : XLSX.utils.json_to_sheet(invoiceRows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Invoice");
  XLSX.writeFile(workbook, `tally_ready_invoice_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function exportReport(result, metadata) {
  const workbook = XLSX.utils.book_new();
  const invoiceRows = buildInvoiceSheetRows(result.matched.filter((row) => row.status === "Matched"));
  const summaryRows = result.summary.map((row) => ({
    Category: row.category,
    "Aggregator Records": row.aggregatorRecords,
    "API Records": row.piRecords,
    "Matched Records": row.matchedRecords,
    "Amount Mismatches": row.amountMismatches,
    "Aggregator Only": row.aggregatorOnly,
    "API Only": row.piOnly,
    "Aggregator Gross Total": row.aggregatorGrossTotal,
    "API Total": row.piTotal,
    "Matched Agg Gross": row.matchedAggGross,
    "Matched API Total": row.matchedPiTotal,
    "Unmatched Agg Gross": row.unmatchedAggGross,
    "Unmatched API Total": row.unmatchedPiTotal,
  }));

  const rules = [
    ["Generated", new Date().toLocaleString()],
    ["Aggregator files", metadata.aggregatorFiles.map((file) => file.name).join(", ")],
    ["API files", metadata.piFiles.map((file) => file.name).join(", ")],
    ["Date Filter Applied", metadata.startDate || metadata.endDate ? `${metadata.startDate || 'Beginning'} to ${metadata.endDate || 'End'}` : "None"],
    ["HDFC MAT", "CCAvenue Ref# = CMS API Reference_No; Order No = API Transaction_No"],
    ["HDFC Others", "Order No = CME API Reference_No; CCAvenue Ref# = API Transaction_No"],
    ["BillDesk", "Ref. 1 = API Reference_No; PGI Ref. No. = API Transaction_No"],
    ["Amount basis", "HDFC Order Amount and BillDesk gross amount column compared to API Total Amt"],
  ];

  const sheets = [
    ["Invoice", invoiceRows],
    ["Summary", summaryRows],
    ["Matching Rules", rules],
    ["Matched", auditRows(result.matched)],
    ["Amount Mismatch", auditRows(result.amountMismatches)],
    ["Aggregator Only", result.aggregatorOnly],
    ["API Only", auditRows(result.piOnly)],
    ["All Results", auditRows(result.all)],
  ];

  sheets.forEach(([name, rows]) => {
    const sheet = Array.isArray(rows[0]) ? XLSX.utils.aoa_to_sheet(rows) : XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  });

  XLSX.writeFile(workbook, `reconciled_tally_report_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function auditRows(rows) {
  return rows.map(({ tallyHeaders, tallyValues, ...row }) => ({
    Status: row.status,
    Category: row.category,
    "Match Rule": row.matchRule || "",
    "Amount Difference": row.amountDifference ?? "",
    "Aggregator Source": row.aggregatorSource || "",
    "Aggregator File": row.aggregatorFile || "",
    "Aggregator Sheet": row.aggregatorSheet || "",
    "Aggregator Row": row.aggregatorRow || "",
    "Aggregator Ref": row.aggregatorKey || "",
    "Aggregator Alt Ref": row.aggregatorAltKey || "",
    "Aggregator Date": row.aggregatorDate || "",
    "Aggregator Amount": row.aggregatorAmount ?? "",
    Charges: row.charges ?? "",
    GST: row.gst ?? "",
    "Net Amount": row.netAmount ?? "",
    "Settlement Date": row.settlementDate || "",
    "Gateway Status": row.gatewayStatus || "",
    "API File": row.piFile || "",
    "API Sheet": row.piSheet || "",
    "API Row": row.piRow || "",
    "API Reference_No": row.piReferenceNo || "",
    "API Transaction_No": row.piTransactionNo || "",
    "API Date": row.piDate || "",
    "API Amount": row.piAmount ?? "",
    Candidate: row.candidate || "",
    "Invoice No": row.invoiceNo || "",
    "Cost Centre": row.costCentre || "",
    Remarks: row.remarks || "",
  }));
}

function buildInvoiceSheetRows(matchedRows) {
  const rows = [];
  const grouped = new Map();
  matchedRows.forEach((row) => {
    const headers = row.tallyHeaders || [];
    if (!headers.length) return;
    const signature = headers.map((header) => normalize(header)).join("¦");
    const key = `${row.category}¦${signature}`;
    if (!grouped.has(key)) grouped.set(key, { category: row.category, headers, rows: [] });
    grouped.get(key).rows.push(row.tallyValues || []);
  });

  Array.from(grouped.values())
    .sort((a, b) => {
      const aIndex = INVOICE_CATEGORY_ORDER.includes(a.category) ? INVOICE_CATEGORY_ORDER.indexOf(a.category) : INVOICE_CATEGORY_ORDER.length;
      const bIndex = INVOICE_CATEGORY_ORDER.includes(b.category) ? INVOICE_CATEGORY_ORDER.indexOf(b.category) : INVOICE_CATEGORY_ORDER.length;
      return aIndex - bIndex || a.category.localeCompare(b.category);
    })
    .forEach((group, index) => {
      if (index > 0) rows.push([]);
      rows.push(group.headers);
      group.rows.forEach((values) => rows.push(group.headers.map((_, cellIndex) => values[cellIndex] ?? "")));
    });

  return rows.length ? rows : [["No reconciled invoice rows"]];
}

function App() {
  const [aggregatorFiles, setAggregatorFiles] = useState([]);
  const [piFiles, setPiFiles] = useState([]);
  const [categoryOverrides, setCategoryOverrides] = useState({});
  const [activeTab, setActiveTab] = useState("All");
  const [activeCategory, setActiveCategory] = useState("CME");
  const [busy, setBusy] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [selectedRow, setSelectedRow] = useState(null);

  const parsed = useMemo(() => {
    const aggregators = parseAggregatorFiles(aggregatorFiles);
    const pi = parsePiFiles(piFiles, categoryOverrides);
    return { aggregators, pi };
  }, [aggregatorFiles, piFiles, categoryOverrides]);

  const visibleRows = useMemo(() => {
    if (!result) return [];
    const rowsByTab = {
      All: result.all,
      Matched: result.matched.filter((row) => row.status === "Matched"),
      "Amount mismatch": result.amountMismatches,
      "Aggregator only": result.aggregatorOnly,
      "API only": result.piOnly,
    };
    return rowsByTab[activeTab] || [];
  }, [activeTab, result]);

  const totals = useMemo(() => {
    if (!result) return { matched: 0, mismatch: 0, aggregatorOnly: 0, piOnly: 0, totalAmount: 0 };
    return {
      matched: result.matched.filter((row) => row.status === "Matched").length,
      mismatch: result.amountMismatches.length,
      aggregatorOnly: result.aggregatorOnly.length,
      piOnly: result.piOnly.length,
      totalAmount: sum(result.matched, "aggregatorAmount"),
    };
  }, [result]);

  async function loadFiles(event, target) {
    setError("");
    setBusy(target);
    try {
      const loaded = await Promise.all(Array.from(event.target.files || []).map(readWorkbookFile));
      if (target === "aggregator") setAggregatorFiles((current) => [...current, ...loaded]);
      if (target === "pi") {
        setPiFiles((current) => [...current, ...loaded]);
        const nextOverrides = {};
        loaded.forEach((file) => {
          nextOverrides[file.id] = inferCategory(file.name);
        });
        setCategoryOverrides((current) => ({ ...current, ...nextOverrides }));
      }
      setResult(null);
    } catch (err) {
      setError(err.message || "Unable to read workbook.");
    } finally {
      setBusy("");
      event.target.value = "";
    }
  }

  function runReconciliation() {
    setError("");
    if (!parsed.aggregators.rows.length) {
      setError("Upload at least one aggregator workbook.");
      return;
    }
    if (!parsed.pi.rows.length) {
      setError("Upload at least one API invoice workbook.");
      return;
    }
    setResult(reconcileRows(parsed.aggregators.rows, parsed.pi.rows, startDate, endDate));
    setActiveTab("All");
  }

  function resetAll() {
    setAggregatorFiles([]);
    setPiFiles([]);
    setCategoryOverrides({});
    setStartDate("");
    setEndDate("");
    setSelectedRow(null);
    setResult(null);
    setError("");
  }

  return (
    <div className="app-shell">
      <aside className="workflow">
        {[
          ["Upload", "Source workbooks", UploadCloud],
          ["Map", "Category rules", Layers3],
          ["Reconcile", "Match and tally", RefreshCw],
          ["Export", "Excel report", ArrowDownToLine],
        ].map(([label, note, Icon], index) => (
          <div className="step" key={label}>
            <div className="step-index">{index + 1}</div>
            <Icon size={20} />
            <div>
              <strong>{label}</strong>
              <span>{note}</span>
            </div>
          </div>
        ))}
      </aside>

      <main className="workspace">
        <header className="topbar">
          <div>
            <h1>Payment Gateway API Reconciliation</h1>
            <p>Upload aggregator sheets, assign API invoice files by category, then generate the final tally-ready reconciled report.</p>
          </div>
          <button className="ghost-button" onClick={resetAll}>
            <RefreshCw size={16} />
            Reset
          </button>
        </header>

        {error && (
          <div className="error-line">
            <CircleAlert size={18} />
            {error}
          </div>
        )}

        <section className="grid upload-grid">
          <UploadPanel
            title="Aggregator Sheets"
            note="HDFC, BillDesk, or similar payment gateway workbooks"
            files={aggregatorFiles}
            busy={busy === "aggregator"}
            onUpload={(event) => loadFiles(event, "aggregator")}
            onRemove={(id) => {
              setAggregatorFiles((current) => current.filter((file) => file.id !== id));
              setResult(null);
            }}
          />
          <UploadPanel
            title="API Invoice Sheets"
            note="CME, CMS, LMA, YLC, and other invoice exports"
            files={piFiles}
            busy={busy === "pi"}
            onUpload={(event) => loadFiles(event, "pi")}
            onRemove={(id) => {
              setPiFiles((current) => current.filter((file) => file.id !== id));
              setCategoryOverrides((current) => {
                const next = { ...current };
                delete next[id];
                return next;
              });
              setResult(null);
            }}
            categoryOverrides={categoryOverrides}
            onCategoryChange={(id, category) => {
              setCategoryOverrides((current) => ({ ...current, [id]: category }));
              setResult(null);
            }}
          />
        </section>

        <section className="grid middle-grid">
          <div className="panel mapping-panel">
            <div className="panel-header">
              <div>
                <h2>Category Mapping</h2>
                <p>Review how API sheets are classified before reconciliation.</p>
              </div>
              <button
                className="ghost-button"
                onClick={() => {
                  const nextOverrides = {};
                  piFiles.forEach((file) => {
                    nextOverrides[file.id] = inferCategory(file.name);
                  });
                  setCategoryOverrides(nextOverrides);
                  setResult(null);
                }}
              >
                <Settings2 size={16} />
                Default Rules
              </button>
            </div>

            <div className="category-tabs">
              {CATEGORIES.map((category) => (
                <button
                  key={category}
                  className={category === activeCategory ? "active" : ""}
                  onClick={() => setActiveCategory(category)}
                >
                  {category}
                </button>
              ))}
            </div>

            <div className="mapping-table">
              <div className="mapping-row head">
                <span>API category</span>
                <span>Aggregator source</span>
                <span>Primary match</span>
                <span>Amount basis</span>
              </div>
              {rulesForCategory(activeCategory).map((row) => (
                <div className="mapping-row" key={row.join("-")}>
                  <span>{row[0]}</span>
                  <span>{row[1]}</span>
                  <span>{row[2]}</span>
                  <span>{row[3]}</span>
                </div>
              ))}
            </div>

            <SourcePreview aggregators={parsed.aggregators} pi={parsed.pi} />

            <div className="date-filter-section">
              <div className="section-header">
                <h3>Date Range Filter</h3>
                <small className="muted">Only reconcile records falling within this range (optional)</small>
              </div>
              <div className="date-inputs">
                <div className="input-group">
                  <label htmlFor="startDate">Start Date</label>
                  <input
                    type="date"
                    id="startDate"
                    value={startDate}
                    onChange={(e) => {
                      setStartDate(e.target.value);
                      setResult(null);
                    }}
                  />
                </div>
                <div className="input-group">
                  <label htmlFor="endDate">End Date</label>
                  <input
                    type="date"
                    id="endDate"
                    value={endDate}
                    onChange={(e) => {
                      setEndDate(e.target.value);
                      setResult(null);
                    }}
                  />
                </div>
                {(startDate || endDate) && (
                  <button className="clear-date-button" onClick={() => {
                    setStartDate("");
                    setEndDate("");
                    setResult(null);
                  }}>
                    Clear Date Filter
                  </button>
                )}
              </div>
            </div>

            <div className="runbar">
              <button className="primary-button" onClick={runReconciliation}>
                <Play size={17} />
                Run Reconciliation
              </button>
              <span>{integer(parsed.aggregators.rows.length)} aggregator rows</span>
              <ChevronRight size={16} />
              <span>{integer(parsed.pi.rows.length)} API rows</span>
            </div>
          </div>

          <SummaryPanel result={result} totals={totals} onCardClick={setActiveTab} activeTab={activeTab} />
        </section>

        <section className="panel result-panel">
          <div className="result-toolbar">
            <div className="result-tabs">
              {DISPLAY_TABS.map((tab) => (
                <button key={tab} className={activeTab === tab ? "active" : ""} onClick={() => setActiveTab(tab)}>
                  {tab}
                </button>
              ))}
            </div>
            <button
              className="export-button"
              disabled={!result}
              onClick={() => exportReport(result, { aggregatorFiles, piFiles, startDate, endDate })}
            >
              <FileSpreadsheet size={17} />
              Export to Excel
            </button>
            <button
              className="export-button tally-only-button"
              disabled={!result}
              onClick={() => exportTallyOnlyReport(result)}
            >
              <ArrowDownToLine size={17} />
              Export Tally-Ready Only
            </button>
          </div>
          <ResultsTable rows={visibleRows} onRowClick={setSelectedRow} />
        </section>
      </main>
      {selectedRow && (
        <DetailModal row={selectedRow} onClose={() => setSelectedRow(null)} />
      )}
    </div>
  );
}

function UploadPanel({ title, note, files, busy, onUpload, onRemove, categoryOverrides, onCategoryChange }) {
  return (
    <div className="panel upload-panel">
      <div className="panel-header">
        <div>
          <h2>{title}</h2>
          <p>{note}</p>
        </div>
        <span className="count-pill">{files.length} files</span>
      </div>
      <label className="drop-zone">
        {busy ? <Loader2 className="spin" size={22} /> : <FolderOpen size={22} />}
        <span>Drop files here or click to browse</span>
        <small>.xlsx and .xls files</small>
        <input type="file" multiple accept=".xlsx,.xls" onChange={onUpload} />
      </label>
      <div className="file-list">
        {files.map((file) => (
          <div className="file-row" key={file.id}>
            <FileCheck2 size={20} />
            <div>
              <strong>{file.name}</strong>
              <span>
                {(file.size / 1024 / 1024).toFixed(2)} MB · {file.sheetNames.length} tab{file.sheetNames.length === 1 ? "" : "s"}
              </span>
            </div>
            {categoryOverrides && (
              <select value={categoryOverrides[file.id] || "Others"} onChange={(event) => onCategoryChange(file.id, event.target.value)}>
                {CATEGORIES.map((category) => (
                  <option key={category}>{category}</option>
                ))}
              </select>
            )}
            <button className="icon-button" onClick={() => onRemove(file.id)} aria-label={`Remove ${file.name}`}>
              ×
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function rulesForCategory(category) {
  if (category === "CMS") {
    return [
      ["CMS", "HDFC MAT", "CCAvenue Ref# → Reference_No", "Order Amount = Total Amt"],
      ["CMS", "BillDesk", "Ref. 1 / PGI Ref. No.", "Gross amount = Total Amt"],
    ];
  }
  if (category === "CME") {
    return [
      ["CME", "HDFC Others", "Order No → Reference_No", "Order Amount = Total Amt"],
      ["CME", "BillDesk", "Ref. 1 / PGI Ref. No.", "Gross amount = Total Amt"],
    ];
  }
  if (category === "LMA" || category === "YLC") {
    return [[category, "BillDesk", "Ref. 1 / PGI Ref. No.", "Gross amount = Total Amt"]];
  }
  return [["Others", "BillDesk or unclassified", "Manual review", "Shown as exception"]];
}

function SourcePreview({ aggregators, pi }) {
  const aggRows = aggregators.sources.slice(0, 4);
  const piRows = pi.sources.slice(0, 4);
  return (
    <div className="source-preview">
      <div>
        <strong>Aggregator tabs</strong>
        {aggRows.length ? aggRows.map((row) => <span key={`${row.file}-${row.sheet}`}>{row.type} · {row.sheet} · {integer(row.rows)}</span>) : <span>No aggregator tabs loaded</span>}
      </div>
      <div>
        <strong>API tabs</strong>
        {piRows.length ? piRows.map((row) => <span key={`${row.file}-${row.sheet}`}>{row.category} · {row.sheet} · {integer(row.rows)}</span>) : <span>No API tabs loaded</span>}
      </div>
    </div>
  );
}

function SummaryPanel({ result, totals, onCardClick, activeTab }) {
  const cards = [
    ["Matched", totals.matched, "Transactions reconciled", "good", Check],
    ["Amount mismatch", totals.mismatch, "Reference matched, amount differs", "warn", CircleAlert],
    ["Aggregator only", totals.aggregatorOnly, "Not found in API", "info", FileSpreadsheet],
    ["API only", totals.piOnly, "Not found in aggregator", "violet", FileSpreadsheet],
  ];
  return (
    <aside className="panel summary-panel">
      <div className="panel-header">
        <div>
          <h2>Reconciliation Summary</h2>
          <p>{result ? "Latest run is ready for review." : "Run reconciliation after uploading files."}</p>
        </div>
      </div>
      {cards.map(([title, value, label, tone, Icon]) => {
        const isActive = activeTab === title;
        return (
          <div 
            className={`metric ${tone} clickable ${isActive ? "active-metric" : ""}`} 
            key={title}
            onClick={() => onCardClick(title)}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                onCardClick(title);
              }
            }}
          >
            <Icon size={22} />
            <div>
              <strong>{title}</strong>
              <span>{label}</span>
            </div>
            <b>{integer(value)}</b>
          </div>
        );
      })}
      <div className="total-box">
        <span>Matched gross amount</span>
        <strong>₹ {money(totals.totalAmount)}</strong>
      </div>
    </aside>
  );
}

function ResultsTable({ rows, onRowClick }) {
  const preview = rows.slice(0, 150);
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Status</th>
            <th>Category</th>
            <th>Aggregator Ref</th>
            <th>API Ref</th>
            <th>Aggregator Amount</th>
            <th>API Amount</th>
            <th>Difference</th>
            <th>Candidate / Remarks</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {preview.length ? (
            preview.map((row, index) => (
              <tr 
                key={`${row.status}-${index}`}
                className="clickable-row"
                onClick={() => onRowClick && onRowClick(row)}
              >
                <td>
                  <span className={`status ${statusClass(row.status)}`}>{row.status}</span>
                </td>
                <td>{row.category}</td>
                <td>{row.aggregatorKey || "—"}</td>
                <td>{row.piReferenceNo || "—"}</td>
                <td>{row.aggregatorAmount !== undefined ? `₹ ${money(row.aggregatorAmount)}` : "—"}</td>
                <td>{row.piAmount !== undefined ? `₹ ${money(row.piAmount)}` : "—"}</td>
                <td>{row.amountDifference !== undefined ? `₹ ${money(row.amountDifference)}` : "—"}</td>
                <td>{row.candidate || row.remarks || "—"}</td>
                <td className="row-action-cell">
                  <span className="details-link">Details →</span>
                </td>
              </tr>
            ))
          ) : (
            <tr>
              <td colSpan="9" className="empty-state">
                Upload files and run reconciliation to preview the report.
              </td>
            </tr>
          )}
        </tbody>
      </table>
      {rows.length > preview.length && <div className="table-note">Showing first {preview.length} of {integer(rows.length)} rows. Export includes all rows.</div>}
    </div>
  );
}

function DetailModal({ row, onClose }) {
  if (!row) return null;

  let title = "";
  let message = "";
  let tone = "info";

  if (row.status === "Matched") {
    title = "Reconciliation Successful";
    message = `This record was successfully matched. The aggregator payment matches the API invoice total of ₹${money(row.piAmount)} using the rule: "${row.matchRule}".`;
    tone = "good";
  } else if (row.status === "Amount mismatch") {
    title = "Amount Mismatch Detected";
    const delta = row.amountDifference;
    message = `The payment and API invoice references matched successfully using the rule: "${row.matchRule}", but the amounts do not match. Aggregator amount is ₹${money(row.aggregatorAmount)} and API amount is ₹${money(row.piAmount)}. The difference is ₹${money(delta)}.`;
    tone = "warn";
  } else if (row.status === "Aggregator only") {
    title = "Payment Found in Aggregator Only";
    message = `This payment was found in the gateway reports (${row.aggregatorSource}) but does not exist in the API invoice sheets. This suggests the invoice record might be missing, unexported, or classified under a different category.`;
    tone = "info";
  } else if (row.status === "API only") {
    title = "Invoice Found in API Only";
    message = `This invoice was found in the API exports but does not exist in any aggregator reports. This suggests that the customer has either not paid yet, or the payment is still pending gateway settlement.`;
    tone = "violet";
  }

  const displayVal = (val) => val !== undefined && val !== "" ? val : "—";

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <header className="modal-header">
          <div>
            <span className={`status-badge ${statusClass(row.status)}`}>{row.status}</span>
            <h2>{title}</h2>
          </div>
          <button className="close-button" onClick={onClose} aria-label="Close details modal">×</button>
        </header>

        <section className={`modal-message-banner ${tone}`}>
          <p>{message}</p>
        </section>

        <div className="modal-comparison-grid">
          <div className="comparison-card">
            <h3>Aggregator Transaction Details</h3>
            {row.status === "API only" ? (
              <div className="empty-side-state">
                <span>Not found in aggregator data</span>
                <p>No gateway records matched this API invoice.</p>
              </div>
            ) : (
              <div className="card-fields">
                <div className="field-row">
                  <span className="field-label">Source Gate:</span>
                  <span className="field-value"><strong>{displayVal(row.aggregatorSource)}</strong></span>
                </div>
                <div className="field-row">
                  <span className="field-label">Source File:</span>
                  <span className="field-value">{displayVal(row.aggregatorFile)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Sheet &amp; Row:</span>
                  <span className="field-value">{displayVal(row.aggregatorSheet)} · Row {displayVal(row.aggregatorRow)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Primary Key (Ref):</span>
                  <span className="field-value highlight">{displayVal(row.aggregatorKey)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Secondary Key (Alt):</span>
                  <span className="field-value">{displayVal(row.aggregatorAltKey)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Transaction Date:</span>
                  <span className="field-value">{displayVal(row.aggregatorDate)}</span>
                </div>
                <div className="field-row amount-row">
                  <span className="field-label">Gross Amount:</span>
                  <span className="field-value">₹ {money(row.aggregatorAmount)}</span>
                </div>
                {row.charges !== "" && (
                  <div className="field-row">
                    <span className="field-label">Charges &amp; GST:</span>
                    <span className="field-value">₹ {money(row.charges)} + ₹ {money(row.gst)}</span>
                  </div>
                )}
                {row.netAmount !== "" && (
                  <div className="field-row">
                    <span className="field-label">Net Amount:</span>
                    <span className="field-value">₹ {money(row.netAmount)}</span>
                  </div>
                )}
                <div className="field-row">
                  <span className="field-label">Gateway Status:</span>
                  <span className="field-value">{displayVal(row.gatewayStatus)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Settlement Date:</span>
                  <span className="field-value">{displayVal(row.settlementDate)}</span>
                </div>
              </div>
            )}
          </div>

          <div className="comparison-card">
            <h3>API Invoice Details</h3>
            {row.status === "Aggregator only" ? (
              <div className="empty-side-state">
                <span>Not found in API data</span>
                <p>No invoice record was found matching this transaction reference.</p>
              </div>
            ) : (
              <div className="card-fields">
                <div className="field-row">
                  <span className="field-label">API Category:</span>
                  <span className="field-value"><strong>{displayVal(row.category)}</strong></span>
                </div>
                <div className="field-row">
                  <span className="field-label">Source File:</span>
                  <span className="field-value">{displayVal(row.piFile)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Sheet &amp; Row:</span>
                  <span className="field-value">{displayVal(row.piSheet)} · Row {displayVal(row.piRow)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Candidate Name:</span>
                  <span className="field-value">{displayVal(row.candidate)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Reference No:</span>
                  <span className="field-value highlight">{displayVal(row.piReferenceNo)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Transaction No:</span>
                  <span className="field-value">{displayVal(row.piTransactionNo)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Invoice No:</span>
                  <span className="field-value">{displayVal(row.invoiceNo)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Transaction Date:</span>
                  <span className="field-value">{displayVal(row.piDate)}</span>
                </div>
                <div className="field-row amount-row">
                  <span className="field-label">API Total Amount:</span>
                  <span className="field-value">₹ {money(row.piAmount)}</span>
                </div>
                <div className="field-row">
                  <span className="field-label">Cost Centre:</span>
                  <span className="field-value">{displayVal(row.costCentre)}</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {row.status !== "API only" && row.status !== "Aggregator only" && (
          <div className="modal-matching-trace">
            <h3>Match Execution Details</h3>
            <div className="trace-details">
              <p>
                <strong>Match Rule Used:</strong> {row.matchRule || "—"}
              </p>
              <p>
                <strong>Remarks:</strong> {row.remarks || "No discrepancies found."}
              </p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function statusClass(status) {
  if (status === "Matched") return "matched";
  if (status === "Amount mismatch") return "mismatch";
  if (status === "Aggregator only") return "agg-only";
  return "pi-only";
}

createRoot(document.getElementById("root")).render(<App />);
