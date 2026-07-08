import React, { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "@e965/xlsx";
import {
  BarChart3,
  CheckCircle2,
  ChevronDown,
  CloudUpload,
  FileSpreadsheet,
  FileText,
  Lightbulb,
  LineChart,
  LockKeyhole,
  Maximize2,
  MoreVertical,
  PieChart,
  RefreshCw,
  Sigma,
  Sun,
  Table2,
  Upload,
  X,
} from "lucide-react";
import { Bar, Line, Pie } from "react-chartjs-2";
import {
  ArcElement,
  BarElement,
  CategoryScale,
  Chart as ChartJS,
  Filler,
  Legend,
  LineElement,
  LinearScale,
  PointElement,
  Tooltip,
} from "chart.js";

ChartJS.register(
  ArcElement,
  BarElement,
  CategoryScale,
  Filler,
  Legend,
  LineElement,
  LinearScale,
  PointElement,
  Tooltip
);

const MAX_FILE_SIZE = 10 * 1024 * 1024;
const MAX_ROWS = 10000;
const MAX_COLUMNS = 80;
const MAX_CELLS = 500000;
const FILTER_PAGE_SIZE = 100;
const TYPE_OPTIONS = ["text", "number", "date", "currency", "percent"];
const TYPE_LABELS = {
  text: "Text",
  number: "Number",
  date: "Date",
  currency: "Currency",
  percent: "Percent",
};
const DATE_BUCKETS = {
  raw: "Original date",
  day: "Day",
  week: "Week",
  month: "Month",
  quarter: "Quarter",
  year: "Year",
};
const AGGREGATIONS = {
  sum: "Sum",
  count: "Count",
  average: "Average",
};
const COLORS = [
  "#1f6feb",
  "#14a7a5",
  "#7c5bd6",
  "#22a06b",
  "#f59e0b",
  "#e11d48",
  "#0ea5e9",
  "#64748b",
];
const SAMPLE_COLUMNS = ["Region", "Category", "Sub-Category", "Sales", "Quantity", "Order Date"];
const SAMPLE_ROWS = [
  { Region: "East", Category: "Technology", "Sub-Category": "Phones", Sales: "152805", Quantity: "640", "Order Date": "2026-01-15" },
  { Region: "East", Category: "Office Supplies", "Sub-Category": "Binders", Sales: "108728", Quantity: "830", "Order Date": "2026-02-10" },
  { Region: "East", Category: "Furniture", "Sub-Category": "Chairs", Sales: "71654", Quantity: "220", "Order Date": "2026-03-04" },
  { Region: "West", Category: "Technology", "Sub-Category": "Accessories", Sales: "183656", Quantity: "710", "Order Date": "2026-01-28" },
  { Region: "West", Category: "Office Supplies", "Sub-Category": "Paper", Sales: "126438", Quantity: "930", "Order Date": "2026-02-16" },
  { Region: "West", Category: "Furniture", "Sub-Category": "Tables", Sales: "94721", Quantity: "260", "Order Date": "2026-03-21" },
  { Region: "Central", Category: "Technology", "Sub-Category": "Machines", Sales: "101734", Quantity: "410", "Order Date": "2026-01-09" },
  { Region: "Central", Category: "Office Supplies", "Sub-Category": "Storage", Sales: "81104", Quantity: "510", "Order Date": "2026-02-25" },
  { Region: "Central", Category: "Furniture", "Sub-Category": "Bookcases", Sales: "63812", Quantity: "185", "Order Date": "2026-04-11" },
  { Region: "South", Category: "Technology", "Sub-Category": "Copiers", Sales: "90432", Quantity: "340", "Order Date": "2026-01-20" },
  { Region: "South", Category: "Office Supplies", "Sub-Category": "Labels", Sales: "69735", Quantity: "470", "Order Date": "2026-02-03" },
  { Region: "South", Category: "Furniture", "Sub-Category": "Furnishings", Sales: "49263", Quantity: "170", "Order Date": "2026-03-14" },
];
const TRUNCATE_LIMIT = 28;
const CHART_LABEL_LIMIT = 18;
const LEGEND_LABEL_LIMIT = 22;

function truncateLabel(value, limit = TRUNCATE_LIMIT) {
  const text = String(value ?? "");
  if (text.length <= limit) return text;
  return `${text.slice(0, Math.max(1, limit - 3))}...`;
}

function cleanHeader(header, index) {
  const value = String(header || "").trim();
  if (/^__EMPTY(?:_\d+)?$/i.test(value)) return `Column ${index + 1}`;
  return value || `Column ${index + 1}`;
}

function makeUniqueHeader(header, existingHeaders) {
  const base = header || "Column";
  let candidate = base;
  let suffix = 2;
  while (existingHeaders.has(candidate.toLowerCase())) {
    candidate = `${base} ${suffix}`;
    suffix += 1;
  }
  existingHeaders.add(candidate.toLowerCase());
  return candidate;
}

function normalizeRows(rows) {
  if (!rows.length) return { columns: [], rows: [] };
  const headerMap = new Map();
  const existingHeaders = new Set();
  const columns = [];

  rows.forEach((row) => {
    Object.keys(row).forEach((originalKey) => {
      if (headerMap.has(originalKey)) return;
      const cleaned = cleanHeader(originalKey, columns.length);
      const normalizedKey = makeUniqueHeader(cleaned, existingHeaders);
      headerMap.set(originalKey, normalizedKey);
      columns.push(normalizedKey);
    });
  });

  return {
    columns,
    rows: rows.map((row) => {
      const next = {};
      columns.forEach((column) => {
        next[column] = "";
      });
      Object.entries(row).forEach(([originalKey, value]) => {
        const normalizedKey = headerMap.get(originalKey);
        if (normalizedKey) next[normalizedKey] = value ?? "";
      });
      return next;
    }),
  };
}

function isBlank(value) {
  const text = String(value ?? "").trim();
  return text === "" || ["-", "Unknown", "N/A", "null"].includes(text);
}

function parseNumberish(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const raw = String(value).trim();
  if (!raw || raw === "-") return null;
  const negative = /^\(.*\)$/.test(raw);
  const percent = /%$/.test(raw);
  const unitMatch = raw.match(/[\u4e07\u842c\u4ebf\u5104]$/);
  const unitMultiplier = unitMatch ? (/[\u4e07\u842c]/.test(unitMatch[0]) ? 10000 : 100000000) : 1;
  let cleaned = raw
    .replace(/[()]/g, "")
    .replace(/[%\u4e07\u842c\u4ebf\u5104]/g, "")
    .replace(/[\s,$\u00a5\uffe5\u20ac\u00a3]/g, "");

  const lastDot = cleaned.lastIndexOf(".");
  const lastComma = cleaned.lastIndexOf(",");
  if (lastDot >= 0 && lastComma >= 0) {
    cleaned =
      lastComma > lastDot
        ? cleaned.replace(/\./g, "").replace(",", ".")
        : cleaned.replace(/,/g, "");
  } else if (lastComma >= 0) {
    const parts = cleaned.split(",");
    const decimalLike = parts.length === 2 && parts[1].length > 0 && parts[1].length <= 2;
    cleaned = decimalLike ? parts.join(".") : parts.join("");
  }

  if (!/^[-+]?\d*\.?\d+$/.test(cleaned)) return null;
  const number = Number(cleaned) * unitMultiplier;
  if (!Number.isFinite(number)) return null;
  const signed = negative ? -number : number;
  return percent ? signed / 100 : signed;
}

function makeUtcDate(year, month, day) {
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  if (
    date.getUTCFullYear() !== Number(year) ||
    date.getUTCMonth() !== Number(month) - 1 ||
    date.getUTCDate() !== Number(day)
  ) {
    return null;
  }
  return date;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return new Date(Date.UTC(date.y, date.m - 1, date.d));
  }
  if (value == null || value === "") return null;

  const raw = String(value).trim();
  if (!raw || raw === "-") return null;

  const monthNames = {
    jan: 0,
    january: 0,
    feb: 1,
    february: 1,
    mar: 2,
    march: 2,
    apr: 3,
    april: 3,
    may: 4,
    jun: 5,
    june: 5,
    jul: 6,
    july: 6,
    aug: 7,
    august: 7,
    sep: 8,
    sept: 8,
    september: 8,
    oct: 9,
    october: 9,
    nov: 10,
    november: 10,
    dec: 11,
    december: 11,
  };
  const yearFirst = raw.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?$/);
  if (yearFirst) {
    const [, year, month, day = "1"] = yearFirst;
    return makeUtcDate(year, month, day);
  }

  const monthFirst = raw.match(/^([A-Za-z]+)[-\s/.](\d{4})$/);
  if (monthFirst) {
    const [, monthName, year] = monthFirst;
    const month = monthNames[monthName.toLowerCase()];
    if (month != null) return makeUtcDate(year, month + 1, 1);
  }

  const dayMonthYear = raw.match(/^(\d{1,2})[-\s/.]([A-Za-z]+)[-\s/.](\d{4})$/);
  if (dayMonthYear) {
    const [, day, monthName, year] = dayMonthYear;
    const month = monthNames[monthName.toLowerCase()];
    if (month != null) return makeUtcDate(year, month + 1, day);
  }

  return null;
}

function getIsoWeek(date) {
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNumber = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - dayNumber);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((target - yearStart) / 86400000 + 1) / 7);
  return { year: target.getUTCFullYear(), week };
}

function formatDatePart(date, bucket = "month") {
  if (!date) return "Unrecognized date";
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  if (bucket === "raw" || bucket === "day") return `${year}-${month}-${day}`;
  if (bucket === "week") {
    const iso = getIsoWeek(date);
    return `${iso.year}-W${String(iso.week).padStart(2, "0")}`;
  }
  if (bucket === "quarter") return `${year}-Q${Math.floor(date.getUTCMonth() / 3) + 1}`;
  if (bucket === "year") return String(year);
  return `${year}-${month}`;
}

function toDimensionValue(row, field, type, dateBucket = "month") {
  const value = row[field];
  if (isBlank(value)) return "(blank)";
  if (type === "date") return formatDatePart(parseDateValue(value), dateBucket);
  if (type === "currency" || type === "number" || type === "percent") {
    const parsed = parseNumberish(value);
    return parsed == null ? String(value).trim() : String(parsed);
  }
  return String(value).trim();
}

function toMeasureValue(row, field, type) {
  if (isBlank(row[field])) return null;
  if (type === "date") return parseDateValue(row[field])?.getTime() ?? null;
  if (type === "currency" || type === "number" || type === "percent") return parseNumberish(row[field]);
  return parseNumberish(row[field]);
}

function inferTypes(rows, columns) {
  const sample = rows.slice(0, 250);
  return columns.reduce((acc, column) => {
    const values = sample.map((row) => row[column]).filter((value) => !isBlank(value));
    const total = values.length || 1;
    let dateHits = 0;
    let numberHits = 0;
    let currencyHits = 0;

    values.forEach((value) => {
      const asString = String(value);
      if (parseDateValue(value)) dateHits += 1;
      if (parseNumberish(value) != null) numberHits += 1;
      if (/[\u00a5\uffe5\u20ac\u00a3$,]/.test(asString) && parseNumberish(value) != null) currencyHits += 1;
    });

    if (dateHits / total >= 0.6 && values.length >= 2) acc[column] = "date";
    else if (currencyHits / total >= 0.35) acc[column] = "currency";
    else if (numberHits / total >= 0.7) acc[column] = "number";
    else acc[column] = "text";
    return acc;
  }, {});
}

function uniqueValues(rows, field, type, dateBucket = "month") {
  const set = new Set();
  rows.forEach((row) => {
    set.add(toDimensionValue(row, field, type, dateBucket));
  });
  return Array.from(set).sort((a, b) => String(a).localeCompare(String(b)));
}

function buildPivot({ rows, columnTypes, rowFields, columnField, valueField, aggregation, filters, dateBuckets = {} }) {
  if (!rows.length || !valueField || (!rowFields.length && !columnField)) {
    return { columns: [], records: [], rowHeaders: [], filteredCount: 0, usableValueCount: 0 };
  }

  const effectiveRowFields = rowFields.length ? rowFields : columnField ? [columnField] : [];
  const effectiveColumnField = rowFields.length ? columnField : "";
  const activeFilters = Object.entries(filters).filter(([, selected]) => selected instanceof Set);
  const filteredRows = rows.filter((row) =>
    activeFilters.every(([field, selected]) =>
      selected.has(toDimensionValue(row, field, columnTypes[field], dateBuckets[field]))
    )
  );

  const buckets = new Map();
  const columnKeys = new Set();
  let usableValueCount = 0;

  filteredRows.forEach((row) => {
    const rowParts = effectiveRowFields.map((field) => toDimensionValue(row, field, columnTypes[field], dateBuckets[field]));
    const rowKey = rowParts.join(" | ") || "Total";
    const colKey = effectiveColumnField ? toDimensionValue(row, effectiveColumnField, columnTypes[effectiveColumnField], dateBuckets[effectiveColumnField]) : "Total";
    const bucketKey = `${rowKey}|||${colKey}`;
    const parsed = toMeasureValue(row, valueField, columnTypes[valueField]);
    const bucket = buckets.get(bucketKey) || {
      rowKey,
      rowParts,
      colKey,
      sum: 0,
      count: 0,
    };

    if (aggregation === "count") {
      if (!isBlank(row[valueField])) {
        bucket.count += 1;
        usableValueCount += 1;
      }
    } else if (parsed != null) {
      bucket.sum += parsed;
      bucket.count += 1;
      usableValueCount += 1;
    }

    buckets.set(bucketKey, bucket);
    columnKeys.add(colKey);
  });

  const columns = columnKeys.size ? Array.from(columnKeys) : ["Total"];
  const recordsByRow = new Map();
  buckets.forEach((bucket) => {
    if (!recordsByRow.has(bucket.rowKey)) {
      recordsByRow.set(bucket.rowKey, {
        rowKey: bucket.rowKey,
        rowParts: bucket.rowParts,
        values: {},
      });
    }
    const value =
      aggregation === "average"
        ? bucket.count
          ? bucket.sum / bucket.count
          : null
        : aggregation === "count"
          ? bucket.count
          : bucket.count
            ? bucket.sum
            : null;
    recordsByRow.get(bucket.rowKey).values[bucket.colKey] = value;
  });

  const records = Array.from(recordsByRow.keys()).map((key) => recordsByRow.get(key));
  return {
    columns,
    records,
    rowHeaders: effectiveRowFields.length ? effectiveRowFields : ["Total"],
    filteredCount: filteredRows.length,
    usableValueCount,
  };
}

function formatMetric(value, type, aggregation) {
  if (value == null || Number.isNaN(value)) return "";
  if (aggregation === "count") return Math.round(value).toLocaleString();
  if (type === "percent") {
    return value.toLocaleString(undefined, { style: "percent", maximumFractionDigits: 2 });
  }
  if (type === "currency") {
    return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function sanitizeForCsv(value) {
  const text = String(value ?? "");
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function makeCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          const value = sanitizeForCsv(cell);
          return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
        })
        .join(",")
    )
    .join("\n");
}

function sanitizeForSpreadsheetText(value) {
  return sanitizeForCsv(value);
}

function sumMetricValues(values) {
  let hasNumber = false;
  const sum = values.reduce((total, value) => {
    const number = Number(value);
    if (!Number.isFinite(number)) return total;
    hasNumber = true;
    return total + number;
  }, 0);
  return hasNumber ? sum : null;
}

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function autoFitColumns(worksheet, minimumWidth = 12, maximumWidth = 36) {
  worksheet.columns.forEach((column) => {
    let width = minimumWidth;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const text = cell.value == null ? "" : String(cell.value);
      width = Math.max(width, Math.min(maximumWidth, text.length + 2));
    });
    column.width = width;
  });
}

function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        dynamicTyping: false,
        complete: (result) => {
          if (result.errors?.length) {
            reject(new Error(result.errors[0].message));
            return;
          }
          resolve(normalizeRows(result.data));
        },
        error: (error) => reject(error),
      });
      return;
    }

    if (ext === "xlsx") {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const workbook = XLSX.read(event.target.result, { type: "array", cellDates: true });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
          resolve(normalizeRows(data));
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error("Could not read this file."));
      reader.readAsArrayBuffer(file);
      return;
    }

    reject(new Error("Only CSV and XLSX files are supported."));
  });
}

function getInitialTypes() {
  return inferTypes(SAMPLE_ROWS, SAMPLE_COLUMNS);
}

export default function App() {
  const [fileName, setFileName] = useState("sales_sample.csv");
  const [rows, setRows] = useState(SAMPLE_ROWS);
  const [columns, setColumns] = useState(SAMPLE_COLUMNS);
  const [columnTypes, setColumnTypes] = useState(getInitialTypes);
  const [rowFields, setRowFields] = useState(["Region"]);
  const [columnField, setColumnField] = useState("Category");
  const [valueField, setValueField] = useState("Sales");
  const [aggregation, setAggregation] = useState("sum");
  const [filterFields, setFilterFields] = useState([]);
  const [filters, setFilters] = useState({});
  const [filterSearches, setFilterSearches] = useState({});
  const [dateBuckets, setDateBuckets] = useState({ "Order Date": "month" });
  const [chartOverride, setChartOverride] = useState("auto");
  const [chartIntent, setChartIntent] = useState("compare");
  const [activeView, setActiveView] = useState("chart");
  const [isFullChartOpen, setIsFullChartOpen] = useState(false);
  const [isParsing, setIsParsing] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState("");
  const chartRef = useRef(null);
  const fileInputRef = useRef(null);
  const toolRef = useRef(null);

  const numericFields = columns.filter((column) => ["number", "currency", "percent"].includes(columnTypes[column]));
  const dimensionFields = columns.filter((column) => column !== valueField);

  const pivot = useMemo(
    () =>
      buildPivot({
        rows,
        columnTypes,
        rowFields,
        columnField,
        valueField,
        aggregation,
        filters,
        dateBuckets,
      }),
    [rows, columnTypes, rowFields, columnField, valueField, aggregation, filters, dateBuckets]
  );

  const recommendedChart = useMemo(() => {
    if (chartOverride !== "auto") return chartOverride;
    const hasDate = rowFields.some((field) => columnTypes[field] === "date");
    if (hasDate) return "line";
    if (chartIntent === "share" && pivot.records.length > 1 && pivot.records.length <= 12 && pivot.columns.length === 1) {
      return "pie";
    }
    return "bar";
  }, [chartIntent, chartOverride, columnTypes, pivot.columns.length, pivot.records.length, rowFields]);

  const chartData = useMemo(() => {
    const labels = pivot.records.length
      ? pivot.records.map((record) => truncateLabel(record.rowKey, CHART_LABEL_LIMIT))
      : pivot.columns.map((column) => truncateLabel(column, CHART_LABEL_LIMIT));

    if (recommendedChart === "pie") {
      const data = pivot.records.length
        ? pivot.records.map((record) => record.values[pivot.columns[0]] ?? null)
        : pivot.columns.map(() => 0);
      return {
        labels,
        datasets: [
          {
            data,
            backgroundColor: COLORS,
            borderColor: "#ffffff",
            borderWidth: 2,
          },
        ],
      };
    }

    return {
      labels,
      datasets: pivot.columns.map((column, index) => ({
        label: truncateLabel(column, LEGEND_LABEL_LIMIT),
        fullLabel: column,
        data: pivot.records.map((record) => record.values[column] ?? null),
        backgroundColor: COLORS[index % COLORS.length],
        borderColor: COLORS[index % COLORS.length],
        borderRadius: recommendedChart === "bar" ? 2 : 0,
        borderWidth: 2,
        tension: 0.32,
        fill: false,
      })),
    };
  }, [pivot, recommendedChart]);

  const fullChartData = useMemo(() => {
    const labels = pivot.records.length
      ? pivot.records.map((record) => truncateLabel(record.rowKey, CHART_LABEL_LIMIT))
      : pivot.columns.map((column) => truncateLabel(column, CHART_LABEL_LIMIT));

    if (recommendedChart === "pie") {
      const data = pivot.records.length
        ? pivot.records.map((record) => record.values[pivot.columns[0]] ?? null)
        : pivot.columns.map(() => 0);
      return {
        labels,
        datasets: [
          {
            data,
            backgroundColor: COLORS,
            borderColor: "#ffffff",
            borderWidth: 2,
          },
        ],
      };
    }

    return {
      labels,
      datasets: pivot.columns.map((column, index) => ({
        label: truncateLabel(column, LEGEND_LABEL_LIMIT),
        fullLabel: column,
        data: pivot.records.map((record) => record.values[column] ?? null),
        backgroundColor: COLORS[index % COLORS.length],
        borderColor: COLORS[index % COLORS.length],
        borderRadius: recommendedChart === "bar" ? 2 : 0,
        borderWidth: 2,
        tension: 0.32,
        fill: false,
      })),
    };
  }, [pivot, recommendedChart]);

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: recommendedChart === "pie" ? "bottom" : "right",
        labels: { boxWidth: 10, boxHeight: 10, color: "#334155", font: { size: 12, weight: 600 } },
      },
      tooltip: {
        callbacks: {
          label: (context) =>
            `${context.dataset.fullLabel || context.dataset.label || "Value"}: ${formatMetric(
              context.parsed.y ?? context.parsed,
              columnTypes[valueField],
              aggregation
            )}`,
          title: (items) => {
            const index = items[0]?.dataIndex ?? 0;
            return pivot.records[index]?.rowKey || items[0]?.label || "";
          },
        },
      },
    },
    scales:
      recommendedChart === "pie"
        ? {}
        : {
            x: {
              grid: { display: false },
              ticks: {
                color: "#475569",
                maxRotation: 0,
                minRotation: 0,
                callback: function callback(value) {
                  return truncateLabel(this.getLabelForValue(value), CHART_LABEL_LIMIT);
                },
              },
            },
            y: {
              beginAtZero: true,
              border: { display: false },
              grid: { color: "#e5eaf1" },
              ticks: { color: "#475569", callback: (value) => `${Number(value) / 1000}K` },
            },
          },
  };

  const fullChartOptions = {
    ...chartOptions,
    plugins: {
      ...chartOptions.plugins,
      legend: {
        ...chartOptions.plugins.legend,
        position: recommendedChart === "pie" ? "bottom" : "right",
      },
    },
    scales:
      recommendedChart === "pie"
        ? {}
        : {
            ...chartOptions.scales,
            x: {
              ...chartOptions.scales.x,
              ticks: {
                ...chartOptions.scales.x.ticks,
                autoSkip: false,
                maxRotation: 0,
                minRotation: 0,
              },
            },
          },
  };

  async function handleFile(file) {
    setError("");
    if (!file) return;
    if (file.size > MAX_FILE_SIZE) {
      setError("The file is larger than 10MB. Please reduce the data range first.");
      return;
    }

    try {
      setIsParsing(true);
      const parsed = await parseFile(file);
      if (!parsed.rows.length) throw new Error("No data rows were detected.");
      if (parsed.rows.length > MAX_ROWS) {
        setError(`This file has ${parsed.rows.length.toLocaleString()} rows. The browser tool supports up to 10,000 rows.`);
        return;
      }
      if (parsed.columns.length > MAX_COLUMNS) {
        setError(`This file has ${parsed.columns.length.toLocaleString()} columns. The browser tool supports up to ${MAX_COLUMNS} columns.`);
        return;
      }
      if (parsed.rows.length * parsed.columns.length > MAX_CELLS) {
        setError(
          `This file has ${(parsed.rows.length * parsed.columns.length).toLocaleString()} cells. Please reduce it below ${MAX_CELLS.toLocaleString()} cells.`
        );
        return;
      }

      const inferred = inferTypes(parsed.rows, parsed.columns);
      const firstNumeric =
        parsed.columns.find((column) => ["number", "currency", "percent"].includes(inferred[column])) || parsed.columns[0];
      const firstDimension = parsed.columns.find((column) => column !== firstNumeric) || parsed.columns[0];
      const secondDimension = parsed.columns.find((column) => column !== firstNumeric && column !== firstDimension) || "";
      const nextDateBuckets = parsed.columns.reduce((acc, column) => {
        if (inferred[column] === "date") acc[column] = "month";
        return acc;
      }, {});

      setFileName(file.name);
      setRows(parsed.rows);
      setColumns(parsed.columns);
      setColumnTypes(inferred);
      setDateBuckets(nextDateBuckets);
      setValueField(firstNumeric);
      setRowFields(firstDimension ? [firstDimension] : []);
      setColumnField(secondDimension);
      setFilterFields([]);
      setFilters({});
      setFilterSearches({});
      setAggregation("sum");
      setChartOverride("auto");
      setActiveView("chart");
    } catch (parseError) {
      setError(parseError.message || "File parsing failed.");
    } finally {
      setIsParsing(false);
    }
  }

  function loadSampleData() {
    const inferred = getInitialTypes();
    setFileName("sales_sample.csv");
    setRows(SAMPLE_ROWS);
    setColumns(SAMPLE_COLUMNS);
    setColumnTypes(inferred);
    setDateBuckets({ "Order Date": "month" });
    setValueField("Sales");
    setRowFields(["Region"]);
    setColumnField("Category");
    setFilterFields([]);
    setFilters({});
    setFilterSearches({});
    setAggregation("sum");
    setChartOverride("auto");
    setChartIntent("compare");
    setActiveView("chart");
    setError("");
    toolRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function reconcileFieldRoles(nextValueField) {
    setRowFields((current) => current.filter((field) => field !== nextValueField && columns.includes(field)));
    setColumnField((current) => (current && current !== nextValueField && columns.includes(current) ? current : ""));
    setFilterFields((current) => {
      const next = current.filter((field) => field !== nextValueField && columns.includes(field));
      setFilters((previous) =>
        next.reduce((acc, field) => {
          acc[field] = previous[field] ?? null;
          return acc;
        }, {})
      );
      setFilterSearches((previous) =>
        next.reduce((acc, field) => {
          acc[field] = previous[field] ?? "";
          return acc;
        }, {})
      );
      return next;
    });
  }

  function handleValueFieldChange(nextValueField) {
    setValueField(nextValueField);
    reconcileFieldRoles(nextValueField);
  }

  function toggleFilterField(field) {
    setFilterFields((current) => {
      const exists = current.includes(field);
      const next = exists ? current.filter((item) => item !== field) : [...current, field];
      setFilters((previous) => {
        const copy = { ...previous };
        if (exists) delete copy[field];
        else copy[field] = null;
        return copy;
      });
      setFilterSearches((previous) => {
        const copy = { ...previous };
        if (exists) delete copy[field];
        else copy[field] = "";
        return copy;
      });
      return next;
    });
  }

  function toggleFilterValue(field, value, values) {
    setFilters((current) => {
      const selected = current[field] instanceof Set ? new Set(current[field]) : new Set(values);
      if (selected.has(value)) selected.delete(value);
      else selected.add(value);
      return { ...current, [field]: selected };
    });
  }

  function setAllFilterValues(field) {
    setFilters((current) => ({ ...current, [field]: null }));
  }

  function clearAllFilterValues(field) {
    setFilters((current) => ({ ...current, [field]: new Set() }));
  }

  function getPivotMatrix() {
    const header = [...(pivot.rowHeaders.length ? pivot.rowHeaders : ["Row"]), ...pivot.columns];
    const body = pivot.records.map((record) => [
      ...(record.rowParts.length ? record.rowParts : [record.rowKey]),
      ...pivot.columns.map((column) => record.values[column] ?? ""),
    ]);
    return [header, ...body];
  }

  function exportPivotCsv() {
    downloadBlob("pivot-table.csv", makeCsv(getPivotMatrix()), "text/csv;charset=utf-8");
  }

  function exportPng() {
    const chart = chartRef.current;
    if (!chart) return;
    const url = chart.toBase64Image("image/png", 1);
    const link = document.createElement("a");
    link.href = url;
    link.download = "pivot-chart.png";
    link.click();
  }

  async function exportWorkbook() {
    setIsExporting(true);
    setError("");
    try {
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Privacy-first Pivot Chart Maker";
      workbook.created = new Date();
      workbook.modified = new Date();

      const sourceSheet = workbook.addWorksheet("Sheet1 Raw Data");
      sourceSheet.addRow(columns.map(sanitizeForSpreadsheetText));
      rows.forEach((row) => {
        sourceSheet.addRow(columns.map((column) => sanitizeForSpreadsheetText(row[column])));
      });
      sourceSheet.views = [{ state: "frozen", ySplit: 1 }];
      sourceSheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
      sourceSheet.getRow(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2563EB" },
      };
      sourceSheet.getRow(1).alignment = { vertical: "middle" };
      autoFitColumns(sourceSheet);

      const pivotSheet = workbook.addWorksheet("Sheet2 Pivot Chart");
      pivotSheet.getCell("A1").value = "Privacy-first Pivot Chart Maker";
      pivotSheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FF172033" } };
      pivotSheet.mergeCells("A1:F1");

      const configRows = [
        ["Rows", pivot.rowHeaders.join(", ") || "-"],
        ["Columns", rowFields.length ? columnField || "-" : "-"],
        ["Values", valueField || "-"],
        ["Aggregation", AGGREGATIONS[aggregation]],
        ["Filtered rows", `${pivot.filteredCount} / ${rows.length}`],
      ];
      configRows.forEach((row, index) => {
        const excelRow = pivotSheet.getRow(index + 3);
        excelRow.values = row.map(sanitizeForSpreadsheetText);
        excelRow.getCell(1).font = { bold: true };
      });

      const chart = chartRef.current;
      if (chart) {
        const imageId = workbook.addImage({
          base64: chart.toBase64Image("image/png", 1),
          extension: "png",
        });
        pivotSheet.addImage(imageId, {
          tl: { col: 0, row: 9 },
          ext: { width: 820, height: 360 },
        });
      }

      const matrix = getPivotMatrix();
      const tableStartRow = 33;
      pivotSheet.getCell(`A${tableStartRow - 2}`).value = "Pivot Table";
      pivotSheet.getCell(`A${tableStartRow - 2}`).font = { bold: true, size: 14 };
      matrix.forEach((row, index) => {
        pivotSheet.getRow(tableStartRow + index).values = row.map((cell, cellIndex) =>
          index === 0 || cellIndex < pivot.rowHeaders.length ? sanitizeForSpreadsheetText(cell) : cell
        );
      });

      const headerRow = pivotSheet.getRow(tableStartRow);
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2563EB" },
      };

      pivotSheet.eachRow((row, rowNumber) => {
        if (rowNumber > tableStartRow) {
          row.eachCell((cell, columnNumber) => {
            if (columnNumber > (pivot.rowHeaders.length || 1)) {
              cell.numFmt = "#,##0.00";
            }
          });
        }
      });
      autoFitColumns(pivotSheet);

      const buffer = await workbook.xlsx.writeBuffer();
      downloadBlob(
        "pivot-chart-maker-export.xlsx",
        buffer,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
    } catch (exportError) {
      setError(exportError.message || "XLSX export failed. Please reduce the data size and try again.");
    } finally {
      setIsExporting(false);
    }
  }

  const hasPivotConfiguration = rows.length > 0 && valueField && (rowFields.length > 0 || columnField);
  const hasUsableValues = aggregation === "count" ? pivot.filteredCount > 0 : pivot.usableValueCount > 0;
  const canRender = hasPivotConfiguration && pivot.records.length > 0 && hasUsableValues;
  const ChartComponent = recommendedChart === "line" ? Line : recommendedChart === "pie" ? Pie : Bar;
  const chartTitle = `${AGGREGATIONS[aggregation]} of ${valueField} by ${(pivot.rowHeaders[0] || "Rows")}${
    rowFields.length && columnField ? ` and ${columnField}` : ""
  }`;
  const visibleRecords = pivot.records.slice(0, 10);
  const grandTotals = pivot.columns.reduce((acc, column) => {
    acc[column] = sumMetricValues(pivot.records.map((record) => record.values[column]));
    return acc;
  }, {});
  const grandTotal = sumMetricValues(Object.values(grandTotals));
  const emptyStateMessage =
    pivot.filteredCount === 0
      ? "No rows match the current filters."
      : !hasUsableValues
        ? `Scanned ${pivot.filteredCount.toLocaleString()} rows, but 0 rows can be used for ${AGGREGATIONS[aggregation].toLowerCase()} of "${valueField}".`
        : "Choose at least one row or column field to build a pivot.";

  function clearFilters() {
    setFilters((current) =>
      Object.keys(current).reduce((acc, field) => {
        acc[field] = null;
        return acc;
      }, {})
    );
  }

  function resetConfiguration() {
    const firstDimension = columns.find((column) => column !== valueField) || "";
    const secondDimension = columns.find((column) => column !== valueField && column !== firstDimension) || "";
    setRowFields(firstDimension ? [firstDimension] : []);
    setColumnField(secondDimension);
    clearFilters();
    setChartOverride("auto");
  }

  const renderPivotTable = ({ full = false } = {}) => {
    const tableRecords = full ? pivot.records : visibleRecords;
    return (
      <table className="pivot-table">
        <thead>
          <tr>
            {(pivot.rowHeaders.length ? pivot.rowHeaders : ["Row"]).map((header) => (
              <th className="row-label-column" key={header} title={header}>
                {truncateLabel(header)}
              </th>
            ))}
            {pivot.columns.map((column) => (
              <th key={column} title={column}>
                {truncateLabel(column)}
              </th>
            ))}
            <th>Grand Total</th>
          </tr>
        </thead>
        <tbody>
          {tableRecords.map((record) => {
            const rowTotal = sumMetricValues(pivot.columns.map((column) => record.values[column]));
            return (
              <tr key={record.rowKey}>
                {(record.rowParts.length ? record.rowParts : [record.rowKey]).map((part, index) => (
                  <td className="truncate-cell row-label-column" key={`${record.rowKey}-${index}`} title={part}>
                    {truncateLabel(part)}
                  </td>
                ))}
                {pivot.columns.map((column) => (
                  <td key={column} className="metric-cell">
                    {formatMetric(record.values[column], columnTypes[valueField], aggregation)}
                  </td>
                ))}
                <td className="metric-cell strong-cell">{formatMetric(rowTotal, columnTypes[valueField], aggregation)}</td>
              </tr>
            );
          })}
          <tr className="grand-row">
            <td title="Grand Total">Grand Total</td>
            {pivot.columns.map((column) => (
              <td key={column} className="metric-cell">
                {formatMetric(grandTotals[column], columnTypes[valueField], aggregation)}
              </td>
            ))}
            <td className="metric-cell">{formatMetric(grandTotal, columnTypes[valueField], aggregation)}</td>
          </tr>
        </tbody>
      </table>
    );
  };

  return (
    <main className="site-shell">
      <header className="top-nav">
        <a className="brand" href="#top" aria-label="Privacy-first Pivot Chart Maker">
          <span className="brand-mark">
            <BarChart3 size={25} />
          </span>
          <span>Privacy-first Pivot Chart Maker</span>
        </a>
        <nav className="nav-links" aria-label="Primary navigation">
          <a href="#features">Features</a>
          <a href="#how-it-works">How It Works</a>
          <a href="#privacy">Privacy</a>
          <a href="#faq">FAQ</a>
          <a href="#contact">Contact</a>
        </nav>
        <div className="nav-actions">
          <button className="icon-button" type="button" aria-label="Toggle theme">
            <Sun size={18} />
          </button>
          <button className="primary-button compact" type="button" onClick={() => fileInputRef.current?.click()}>
            Try the Tool
          </button>
        </div>
      </header>

      <section className="hero-section" id="top">
        <h1>Build Pivot Charts Directly in Your Browser</h1>
        <p>
          Your data never leaves your device. Upload, pivot, visualize, and export all locally, privately, and securely.
        </p>
        <div className="hero-actions">
          <button className="primary-button" type="button" onClick={() => fileInputRef.current?.click()}>
            <CloudUpload size={20} />
            Try the Tool
          </button>
          <button className="secondary-button" type="button" onClick={loadSampleData}>
            <FileText size={20} />
            Use Sample Data
          </button>
        </div>

        
      </section>

      <section className="tool-frame" ref={toolRef} aria-label="Pivot chart maker">
        <div className="tool-titlebar">
          <div className="tool-brand">
            <BarChart3 size={18} />
            <strong>Privacy-first Pivot Chart Maker</strong>
          </div>
          <div className="view-tabs" role="tablist" aria-label="Output views">
            <button className={activeView === "chart" ? "active" : ""} type="button" onClick={() => setActiveView("chart")}>
              <BarChart3 size={15} />
              Chart
            </button>
            <button className={activeView === "table" ? "active" : ""} type="button" onClick={() => setActiveView("table")}>
              <Table2 size={15} />
              Pivot Table
            </button>
            <button className={activeView === "preview" ? "active" : ""} type="button" onClick={() => setActiveView("preview")}>
              <FileSpreadsheet size={15} />
              Data Preview
            </button>
          </div>
          <button className="reset-button" type="button" onClick={loadSampleData}>
            <RefreshCw size={14} />
            Reset
          </button>
        </div>

        <div className="workspace-grid">
          <aside className="control-panel">
            <section className="panel-step">
              <h2>1. Upload your data</h2>
              <label className="upload-zone">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".csv,.xlsx"
                  onChange={(event) => handleFile(event.target.files?.[0])}
                />
                <Upload size={27} />
                <strong>Drag and drop CSV or XLSX here</strong>
                <span>or click to browse</span>
                <small>Max 10MB, 10,000 rows, {MAX_COLUMNS} columns, {MAX_CELLS.toLocaleString()} cells.</small>
              </label>
              <div className="file-chip">
                <FileSpreadsheet size={18} />
                <strong>{fileName || "No file selected"}</strong>
                <span>{isParsing ? "Parsing..." : `${rows.length.toLocaleString()} rows`}</span>
                <CheckCircle2 size={16} />
              </div>
              {error ? <div className="error-box">{error}</div> : null}
            </section>

            <section className="panel-step">
              <h2>2. Fields & types</h2>
              <div className="field-table">
                <div className="field-table-head">
                  <span>Field name</span>
                  <span>Detected type</span>
                  <span>Override</span>
                </div>
                {columns.map((column) => (
                  <label className="field-row" key={column}>
                    <span title={column}>{column}</span>
                    <span className={`type-token ${columnTypes[column]}`}>
                      {columnTypes[column] === "number" || columnTypes[column] === "currency" ? "123" : columnTypes[column] === "date" ? "Cal" : "Abc"}
                      <em>{TYPE_LABELS[columnTypes[column]]}</em>
                    </span>
                    <span className="field-controls">
                      <select
                        value={columnTypes[column]}
                        onChange={(event) => {
                          const nextType = event.target.value;
                          setColumnTypes((current) => ({ ...current, [column]: nextType }));
                          if (nextType === "date") {
                            setDateBuckets((current) => ({ ...current, [column]: current[column] || "month" }));
                          }
                        }}
                      >
                        {TYPE_OPTIONS.map((type) => (
                          <option value={type} key={type}>
                            {TYPE_LABELS[type]}
                          </option>
                        ))}
                      </select>
                      {columnTypes[column] === "date" ? (
                        <select
                          value={dateBuckets[column] || "month"}
                          aria-label={`${column} date grouping`}
                          onChange={(event) => setDateBuckets((current) => ({ ...current, [column]: event.target.value }))}
                        >
                          {Object.entries(DATE_BUCKETS).map(([value, label]) => (
                            <option value={value} key={value}>
                              {label}
                            </option>
                          ))}
                        </select>
                      ) : null}
                    </span>
                  </label>
                ))}
              </div>
            </section>

            <section className="panel-step">
              <h2>3. Pivot configuration</h2>
              <div className="config-grid">
                <label>
                  Row field
                  <select value={rowFields[0] || ""} onChange={(event) => setRowFields(event.target.value ? [event.target.value] : [])}>
                    <option value="">None</option>
                    {dimensionFields.map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  Column field
                  <select value={columnField} onChange={(event) => setColumnField(event.target.value)}>
                    <option value="">None</option>
                    {dimensionFields.map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  Value field
                  <select value={valueField} onChange={(event) => handleValueFieldChange(event.target.value)}>
                    {columns.map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                  </select>
                  {numericFields.length === 0 ? <small>No obvious number column detected.</small> : null}
                </label>
                <label>
                  Aggregation
                  <select value={aggregation} onChange={(event) => setAggregation(event.target.value)}>
                    {Object.entries(AGGREGATIONS).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
            </section>

            <section className="panel-step">
              <h2>4. Filters (optional)</h2>
              <div className="filter-pill-row">
                {filterFields.map((field) => {
                  const values = uniqueValues(rows, field, columnTypes[field], dateBuckets[field]);
                  const selected = filters[field];
                  const selectedCount = selected instanceof Set ? selected.size : values.length;
                  return (
                    <button className="filter-pill" type="button" key={field} onClick={() => toggleFilterField(field)}>
                      {field}: {selectedCount === values.length ? "All" : `${selectedCount}/${values.length}`}
                      <X size={13} />
                    </button>
                  );
                })}
                <select value="" onChange={(event) => event.target.value && toggleFilterField(event.target.value)}>
                  <option value="">+ Add filter</option>
                  {dimensionFields
                    .filter((field) => !filterFields.includes(field))
                    .map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                </select>
              </div>
            </section>

            <section className="panel-step">
              <h2>5. Chart settings</h2>
              <div className="chart-config">
                <label>
                  Chart intent
                  <span className="segmented">
                    <button className={chartIntent === "compare" ? "active" : ""} type="button" onClick={() => setChartIntent("compare")}>
                      Compare
                    </button>
                    <button className={chartIntent === "share" ? "active" : ""} type="button" onClick={() => setChartIntent("share")}>
                      Share
                    </button>
                  </span>
                </label>
                <label>
                  Recommended chart
                  <select value={chartOverride} onChange={(event) => setChartOverride(event.target.value)}>
                    <option value="auto">Auto recommendation</option>
                    <option value="bar">Clustered Column (Bar)</option>
                    <option value="line">Line</option>
                    <option value="pie">Pie</option>
                  </select>
                </label>
              </div>
            </section>
          </aside>

          <section className="output-panel">
            <div className="chart-toolbar">
              <div>
                <h2>{chartTitle}</h2>
                <span>
                  {pivot.filteredCount.toLocaleString()} of {rows.length.toLocaleString()} rows included
                </span>
              </div>
              <div className="chart-mode-buttons">
                <button className={recommendedChart === "bar" ? "active" : ""} type="button" onClick={() => setChartOverride("bar")} aria-label="Bar chart">
                  <BarChart3 size={20} />
                </button>
                <button className={recommendedChart === "line" ? "active" : ""} type="button" onClick={() => setChartOverride("line")} aria-label="Line chart">
                  <LineChart size={20} />
                </button>
                <button className={recommendedChart === "pie" ? "active" : ""} type="button" onClick={() => setChartOverride("pie")} aria-label="Pie chart">
                  <PieChart size={20} />
                </button>
                <button type="button" aria-label="More options">
                  <MoreVertical size={20} />
                </button>
              </div>
            </div>

            {canRender && activeView === "chart" ? (
              <button className="full-chart-fab" type="button" onClick={() => setIsFullChartOpen(true)} aria-label="View full chart" title="View full chart">
                <Maximize2 size={19} />
              </button>
            ) : null}

            {filterFields.length ? (
              <div className="filters-row">
                {filterFields.map((field) => {
                  const values = uniqueValues(rows, field, columnTypes[field], dateBuckets[field]);
                  const selected = filters[field];
                  const selectedCount = selected instanceof Set ? selected.size : values.length;
                  const query = filterSearches[field] || "";
                  const visibleValues = values
                    .filter((value) => String(value).toLowerCase().includes(query.trim().toLowerCase()))
                    .slice(0, FILTER_PAGE_SIZE);
                  const hiddenCount = Math.max(
                    0,
                    values.filter((value) => String(value).toLowerCase().includes(query.trim().toLowerCase())).length -
                      visibleValues.length
                  );
                  return (
                    <div className="filter-block" key={field}>
                      <div className="filter-heading">
                        <strong>{field}</strong>
                        <span>
                          {selectedCount} / {values.length} selected
                        </span>
                      </div>
                      <div className="filter-tools">
                        <button type="button" onClick={() => setAllFilterValues(field)}>
                          All
                        </button>
                        <button type="button" onClick={() => clearAllFilterValues(field)}>
                          None
                        </button>
                      </div>
                      {values.length > FILTER_PAGE_SIZE || query ? (
                        <input
                          className="filter-search"
                          type="search"
                          value={query}
                          placeholder="Search values"
                          onChange={(event) =>
                            setFilterSearches((current) => ({ ...current, [field]: event.target.value }))
                          }
                        />
                      ) : null}
                      <div className="filter-values">
                        {visibleValues.map((value) => (
                          <label className="filter-chip" key={value}>
                            <input
                              type="checkbox"
                              checked={selected instanceof Set ? selected.has(value) : true}
                              onChange={() => toggleFilterValue(field, value, values)}
                            />
                            <span>{value}</span>
                          </label>
                        ))}
                      </div>
                      {hiddenCount ? <small>{hiddenCount.toLocaleString()} more values. Search to narrow the list.</small> : null}
                    </div>
                  );
                })}
              </div>
            ) : null}

            {canRender && activeView !== "preview" ? (
              <div className={activeView === "table" ? "chart-wrap hidden" : "chart-wrap"}>
                <ChartComponent ref={chartRef} data={chartData} options={chartOptions} />
              </div>
            ) : null}

            {activeView === "preview" ? (
              <div className="table-wrap preview-table">
                <table>
                  <thead>
                    <tr>
                      {columns.map((column) => (
                        <th key={column}>{column}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.slice(0, 10).map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {columns.map((column) => (
                          <td key={column}>{row[column]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : canRender ? (
              <div className="table-wrap">
                <div className="table-title">
                  <strong>Pivot Table</strong>
                  <span>Top 10 shown</span>
                </div>
                {renderPivotTable()}
              </div>
            ) : (
              <div className="empty-state">
                <strong>{emptyStateMessage}</strong>
                <span>
                  {pivot.filteredCount.toLocaleString()} of {rows.length.toLocaleString()} rows included,
                  {` ${pivot.usableValueCount.toLocaleString()} usable values found.`}
                </span>
                <div>
                  <button type="button" onClick={clearFilters}>
                    Clear filters
                  </button>
                  <button type="button" onClick={resetConfiguration}>
                    Restore defaults
                  </button>
                </div>
              </div>
            )}

            <div className="export-actions">
              <button className="export-button excel" type="button" onClick={exportWorkbook} disabled={!canRender || isExporting}>
                <FileSpreadsheet size={20} />
                {isExporting ? "Exporting..." : "Export XLSX"}
              </button>
              <button className="export-button png" type="button" onClick={exportPng} disabled={!canRender || isExporting}>
                <BarChart3 size={20} />
                Export PNG
              </button>
              <button className="export-button csv" type="button" onClick={exportPivotCsv} disabled={!canRender || isExporting}>
                <FileText size={20} />
                Export CSV
                <ChevronDown size={16} />
              </button>
            </div>
          </section>
        </div>
      </section>

      {isFullChartOpen ? (
        <div className="modal-backdrop" role="presentation" onClick={() => setIsFullChartOpen(false)}>
          <section className="full-chart-modal" role="dialog" aria-modal="true" aria-labelledby="full-chart-title" onClick={(event) => event.stopPropagation()}>
            <header className="modal-header">
              <div>
                <h2 id="full-chart-title">{chartTitle}</h2>
                <span>
                  {pivot.filteredCount.toLocaleString()} of {rows.length.toLocaleString()} rows included
                </span>
              </div>
              <button className="modal-close" type="button" onClick={() => setIsFullChartOpen(false)} aria-label="Close full chart">
                <X size={20} />
              </button>
            </header>
            <div className="modal-chart-wrap">
              <ChartComponent data={fullChartData} options={fullChartOptions} />
            </div>
          </section>
        </div>
      ) : null}

      <section className="workflow-strip" id="how-it-works" aria-label="How Privacy-first Pivot Chart Maker works">
        <article>
          <span className="step-number">1</span>
          <span className="step-icon">
            <Upload size={28} />
          </span>
          <div>
            <h2>Upload CSV or XLSX</h2>
            <p>Select a spreadsheet file and parse it locally in your browser.</p>
          </div>
        </article>
        <span className="workflow-arrow" aria-hidden="true">/</span>
        <article>
          <span className="step-number">2</span>
          <span className="step-icon">
            <Table2 size={28} />
          </span>
          <div>
            <h2>Build a pivot table</h2>
            <p>Choose row fields, column fields, value fields, aggregation, and filters.</p>
          </div>
        </article>
        <span className="workflow-arrow" aria-hidden="true">/</span>
        <article>
          <span className="step-number">3</span>
          <span className="step-icon">
            <BarChart3 size={28} />
          </span>
          <div>
            <h2>Export charts</h2>
            <p>Download pivot tables as CSV or XLSX, or export the chart as PNG.</p>
          </div>
        </article>
      </section>

      <section className="feature-grid" id="features">
        <article>
          <span className="feature-icon green-bg">
            <LockKeyhole size={31} />
          </span>
          <div>
            <h2>Local processing</h2>
            <p>All parsing, aggregation, and export happen in your browser.</p>
            <strong>Your data never leaves your device.</strong>
          </div>
        </article>
        <article>
          <span className="feature-icon purple-bg">
            <Sigma size={31} />
          </span>
          <div>
            <h2>Flexible aggregations</h2>
            <p>Sum, Average, Count, Min, Max, Distinct Count, and more.</p>
            <span>Full control over your insights.</span>
          </div>
        </article>
        <article>
          <span className="feature-icon blue-bg">
            <Lightbulb size={31} />
          </span>
          <div>
            <h2>Smart chart recommendation</h2>
            <p>We suggest the best chart for your data and intent.</p>
            <span>Bar, Line, and Pie charts supported.</span>
          </div>
        </article>
        <article id="changelog">
          <span className="feature-icon amber-bg">
            <FileText size={31} />
          </span>
          <div>
            <h2>Export ready</h2>
            <p>Export pivot tables or charts to XLSX, PNG, or CSV.</p>
            <span>Ready to use anywhere.</span>
          </div>
        </article>
      </section>

      <section className="seo-section" id="faq" aria-labelledby="faq-title">
        <div className="seo-section-heading">
          <h2 id="faq-title">FAQ</h2>
          <p>Quick answers for people comparing spreadsheet tools, pivot table builders, and local data analysis workflows.</p>
        </div>
        <div className="faq-grid">
          <article>
            <h3>Is this pivot chart maker free?</h3>
            <p>Yes. The current browser tool is free to use for CSV and XLSX files within the listed size limits.</p>
          </article>
          <article>
            <h3>Does my spreadsheet get uploaded to a server?</h3>
            <p>No. File parsing, pivot aggregation, chart rendering, and export happen locally in your browser.</p>
          </article>
          <article>
            <h3>Which file formats are supported?</h3>
            <p>The tool supports CSV and XLSX files. You can export pivot results as CSV, XLSX, or PNG.</p>
          </article>
          <article>
            <h3>Can I create pivot tables without Excel?</h3>
            <p>Yes. You can upload a spreadsheet, choose fields, aggregate values, filter categories, and export the result without opening Excel.</p>
          </article>
          <article>
            <h3>What happens if my data has messy headers?</h3>
            <p>The parser normalizes blank, repeated, spaced, and generated headers while preserving the original row values.</p>
          </article>
          <article>
            <h3>Why are there file size limits?</h3>
            <p>The tool runs inside your browser, so limits help avoid freezing low-memory devices when parsing or exporting spreadsheets.</p>
          </article>
        </div>
      </section>

      <section className="seo-section split" id="privacy" aria-labelledby="privacy-title">
        <div>
          <h2 id="privacy-title">Privacy</h2>
          <p>
            Privacy-first Pivot Chart Maker is designed for local spreadsheet analysis. Your CSV or XLSX file is read by
            browser APIs on your device, and the app does not intentionally send spreadsheet contents to a remote server.
          </p>
        </div>
        <div>
          <h3>What the app stores</h3>
          <p>
            The app keeps uploaded rows in browser memory while you work. Refreshing the page clears the current session.
            Exported files are generated on your device.
          </p>
        </div>
      </section>

      <section className="seo-section split contact-section" id="contact" aria-labelledby="contact-title">
        <div>
          <h2 id="contact-title">Contact</h2>
          <p>Questions, bug reports, and privacy concerns are welcome.</p>
        </div>
        <a href="mailto:cloudhu2000@gmail.com">cloudhu2000@gmail.com</a>
      </section>
    </main>
  );
}
