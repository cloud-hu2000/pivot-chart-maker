import React, { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "@e965/xlsx";
import {
  BarChart3,
  CalendarDays,
  Download,
  FileSpreadsheet,
  Filter,
  PieChart,
  RefreshCw,
  Table2,
  Upload,
} from "lucide-react";
import {
  Bar,
  Line,
  Pie,
} from "react-chartjs-2";
import {
  BarElement,
  CategoryScale,
  Chart as ChartJS,
  Filler,
  Legend,
  LineElement,
  LinearScale,
  PointElement,
  ArcElement,
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
const TYPE_OPTIONS = ["text", "number", "date", "currency"];
const TYPE_LABELS = {
  text: "文本",
  number: "数字",
  date: "日期",
  currency: "货币",
};
const AGGREGATIONS = {
  sum: "求和",
  count: "计数",
  average: "平均值",
};
const COLORS = [
  "#2563eb",
  "#16a34a",
  "#dc2626",
  "#ca8a04",
  "#7c3aed",
  "#0891b2",
  "#db2777",
  "#4b5563",
];

function cleanHeader(header, index) {
  const value = String(header || "").trim();
  return value || `Column ${index + 1}`;
}

function normalizeRows(rows) {
  if (!rows.length) return { columns: [], rows: [] };
  const columns = Array.from(
    rows.reduce((set, row) => {
      Object.keys(row).forEach((key) => set.add(cleanHeader(key, set.size)));
      return set;
    }, new Set())
  );

  return {
    columns,
    rows: rows.map((row) => {
      const next = {};
      columns.forEach((column) => {
        next[column] = row[column] ?? "";
      });
      return next;
    }),
  };
}

function isBlank(value) {
  return value == null || String(value).trim() === "" || ["-", "未知", "N/A", "null"].includes(String(value).trim());
}

function parseNumberish(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const raw = String(value).trim();
  if (!raw || raw === "-") return null;
  const negative = /^\(.*\)$/.test(raw);
  const cleaned = raw
    .replace(/[,$¥￥€£\s]/g, "")
    .replace(/元|人民币|rmb|RMB|%/g, "")
    .replace(/[()]/g, "");
  if (!/^[-+]?\d*\.?\d+$/.test(cleaned)) return null;
  const number = Number(cleaned);
  if (!Number.isFinite(number)) return null;
  return negative ? -number : number;
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
  const yearFirst = raw.match(/^(\d{4})[-/.年](\d{1,2})(?:[-/.月](\d{1,2}))?(?:日)?$/);
  if (yearFirst) {
    const [, year, month, day = "1"] = yearFirst;
    return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  }

  const monthFirst = raw.match(/^([A-Za-z]+)[-\s/.](\d{4})$/);
  if (monthFirst) {
    const [, monthName, year] = monthFirst;
    const month = monthNames[monthName.toLowerCase()];
    if (month != null) return new Date(Date.UTC(Number(year), month, 1));
  }

  const dayMonthYear = raw.match(/^(\d{1,2})[-\s/.]([A-Za-z]+)[-\s/.](\d{4})$/);
  if (dayMonthYear) {
    const [, day, monthName, year] = dayMonthYear;
    const month = monthNames[monthName.toLowerCase()];
    if (month != null) return new Date(Date.UTC(Number(year), month, Number(day)));
  }

  return null;
}

function formatDatePart(date) {
  if (!date) return "Invalid date";
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function toDimensionValue(row, field, type) {
  const value = row[field];
  if (isBlank(value)) return "(空值)";
  if (type === "date") return formatDatePart(parseDateValue(value));
  if (type === "currency" || type === "number") {
    const parsed = parseNumberish(value);
    return parsed == null ? String(value).trim() : String(parsed);
  }
  return String(value).trim();
}

function toMeasureValue(row, field, type) {
  if (isBlank(row[field])) return null;
  if (type === "date") return parseDateValue(row[field])?.getTime() ?? null;
  if (type === "currency" || type === "number") return parseNumberish(row[field]);
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
      if (/[¥￥$€£元,]/.test(asString) && parseNumberish(value) != null) currencyHits += 1;
    });

    if (dateHits / total >= 0.6 && values.length >= 2) acc[column] = "date";
    else if (currencyHits / total >= 0.35) acc[column] = "currency";
    else if (numberHits / total >= 0.7) acc[column] = "number";
    else acc[column] = "text";
    return acc;
  }, {});
}

function uniqueValues(rows, field, type, limit = 100) {
  const set = new Set();
  rows.forEach((row) => {
    if (set.size < limit) set.add(toDimensionValue(row, field, type));
  });
  return Array.from(set).sort((a, b) => String(a).localeCompare(String(b), "zh-Hans-CN"));
}

function sortSmart(values) {
  return [...values].sort((a, b) => {
    const dateA = /^\d{4}-\d{2}$/.test(a);
    const dateB = /^\d{4}-\d{2}$/.test(b);
    if (dateA && dateB) return a.localeCompare(b);
    const numA = Number(a);
    const numB = Number(b);
    if (Number.isFinite(numA) && Number.isFinite(numB)) return numA - numB;
    return String(a).localeCompare(String(b), "zh-Hans-CN");
  });
}

function buildPivot({ rows, columnTypes, rowFields, columnField, valueField, aggregation, filters }) {
  if (!rows.length || !valueField || (!rowFields.length && !columnField)) {
    return { columns: [], records: [], rowHeaders: [], filteredCount: 0 };
  }

  const activeFilters = Object.entries(filters).filter(([, selected]) => selected?.size);
  const filteredRows = rows.filter((row) =>
    activeFilters.every(([field, selected]) =>
      selected.has(toDimensionValue(row, field, columnTypes[field]))
    )
  );

  const buckets = new Map();
  const columnKeys = new Set();

  filteredRows.forEach((row) => {
    const rowParts = rowFields.map((field) => toDimensionValue(row, field, columnTypes[field]));
    const rowKey = rowParts.join(" | ") || "Total";
    const colKey = columnField ? toDimensionValue(row, columnField, columnTypes[columnField]) : "Total";
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
      if (!isBlank(row[valueField])) bucket.count += 1;
    } else if (parsed != null) {
      bucket.sum += parsed;
      bucket.count += 1;
    }

    buckets.set(bucketKey, bucket);
    columnKeys.add(colKey);
  });

  const columns = sortSmart(columnKeys.size ? Array.from(columnKeys) : ["Total"]);
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
          : bucket.sum;
    recordsByRow.get(bucket.rowKey).values[bucket.colKey] = value;
  });

  const records = sortSmart(Array.from(recordsByRow.keys())).map((key) => recordsByRow.get(key));
  return {
    columns,
    records,
    rowHeaders: rowFields.length ? rowFields : [columnField || "Total"],
    filteredCount: filteredRows.length,
  };
}

function formatMetric(value, type, aggregation) {
  if (value == null || Number.isNaN(value)) return "";
  if (aggregation === "count") return Math.round(value).toLocaleString();
  if (type === "currency") {
    return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function makeCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          const value = cell == null ? "" : String(cell);
          return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
        })
        .join(",")
    )
    .join("\n");
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
      reader.onerror = () => reject(new Error("无法读取文件"));
      reader.readAsArrayBuffer(file);
      return;
    }

    reject(new Error("仅支持 CSV / XLSX 文件"));
  });
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState([]);
  const [columns, setColumns] = useState([]);
  const [columnTypes, setColumnTypes] = useState({});
  const [rowFields, setRowFields] = useState([]);
  const [columnField, setColumnField] = useState("");
  const [valueField, setValueField] = useState("");
  const [aggregation, setAggregation] = useState("sum");
  const [filterFields, setFilterFields] = useState([]);
  const [filters, setFilters] = useState({});
  const [chartOverride, setChartOverride] = useState("auto");
  const [chartIntent, setChartIntent] = useState("compare");
  const [error, setError] = useState("");
  const chartRef = useRef(null);

  const numericFields = columns.filter((column) => ["number", "currency"].includes(columnTypes[column]));
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
      }),
    [rows, columnTypes, rowFields, columnField, valueField, aggregation, filters]
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
      ? pivot.records.map((record) => record.rowKey)
      : pivot.columns;

    if (recommendedChart === "pie") {
      const data = pivot.records.length
        ? pivot.records.map((record) => record.values[pivot.columns[0]] || 0)
        : pivot.columns.map(() => 0);
      return {
        labels,
        datasets: [
          {
            data,
            backgroundColor: COLORS,
            borderWidth: 1,
          },
        ],
      };
    }

    return {
      labels,
      datasets: pivot.columns.map((column, index) => ({
        label: column,
        data: pivot.records.map((record) => record.values[column] || 0),
        backgroundColor: COLORS[index % COLORS.length],
        borderColor: COLORS[index % COLORS.length],
        borderWidth: 2,
        tension: 0.3,
        fill: false,
      })),
    };
  }, [pivot, recommendedChart]);

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "bottom",
        labels: { boxWidth: 12, font: { size: 12 } },
      },
      tooltip: {
        callbacks: {
          label: (context) => `${context.dataset.label || "值"}: ${formatMetric(context.parsed.y ?? context.parsed, columnTypes[valueField], aggregation)}`,
        },
      },
    },
    scales:
      recommendedChart === "pie"
        ? {}
        : {
            x: { ticks: { maxRotation: 40, minRotation: 0 } },
            y: { beginAtZero: true },
          },
  };

  async function handleFile(file) {
    setError("");
    if (!file) return;
    if (file.size > MAX_FILE_SIZE) {
      setError("文件超过 10MB，请先缩小数据范围。");
      return;
    }

    try {
      const parsed = await parseFile(file);
      if (!parsed.rows.length) throw new Error("没有识别到数据行。");
      if (parsed.rows.length > MAX_ROWS) {
        setError(`当前文件有 ${parsed.rows.length.toLocaleString()} 行，MVP 版本最多支持 10,000 行。`);
        return;
      }

      const inferred = inferTypes(parsed.rows, parsed.columns);
      const firstNumeric = parsed.columns.find((column) => ["number", "currency"].includes(inferred[column])) || parsed.columns[0];
      const firstDimension = parsed.columns.find((column) => column !== firstNumeric) || parsed.columns[0];

      setFileName(file.name);
      setRows(parsed.rows);
      setColumns(parsed.columns);
      setColumnTypes(inferred);
      setValueField(firstNumeric);
      setRowFields(firstDimension ? [firstDimension] : []);
      setColumnField("");
      setFilterFields([]);
      setFilters({});
      setAggregation("sum");
      setChartOverride("auto");
    } catch (parseError) {
      setError(parseError.message || "文件解析失败。");
    }
  }

  function toggleRowField(field) {
    setRowFields((current) =>
      current.includes(field) ? current.filter((item) => item !== field) : [...current, field]
    );
  }

  function toggleFilterField(field) {
    setFilterFields((current) => {
      const exists = current.includes(field);
      const next = exists ? current.filter((item) => item !== field) : [...current, field];
      setFilters((previous) => {
        const copy = { ...previous };
        if (exists) delete copy[field];
        else copy[field] = new Set(uniqueValues(rows, field, columnTypes[field]).slice(0, 20));
        return copy;
      });
      return next;
    });
  }

  function toggleFilterValue(field, value) {
    setFilters((current) => {
      const selected = new Set(current[field] || []);
      if (selected.has(value)) selected.delete(value);
      else selected.add(value);
      return { ...current, [field]: selected };
    });
  }

  function getPivotMatrix() {
    const header = [...(rowFields.length ? rowFields : ["Row"]), ...pivot.columns];
    const body = pivot.records.map((record) => [
      ...(rowFields.length ? record.rowParts : [record.rowKey]),
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
    const ExcelJS = (await import("exceljs")).default;
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Privacy-first Pivot Chart Maker";
    workbook.created = new Date();
    workbook.modified = new Date();

    const sourceSheet = workbook.addWorksheet("Sheet1 Raw Data");
    sourceSheet.addRow(columns);
    rows.forEach((row) => {
      sourceSheet.addRow(columns.map((column) => row[column]));
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
      ["Rows", rowFields.join(", ") || "-"],
      ["Columns", columnField || "-"],
      ["Values", valueField || "-"],
      ["Aggregation", AGGREGATIONS[aggregation]],
      ["Filtered rows", `${pivot.filteredCount} / ${rows.length}`],
    ];
    configRows.forEach((row, index) => {
      const excelRow = pivotSheet.getRow(index + 3);
      excelRow.values = row;
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
      pivotSheet.getRow(tableStartRow + index).values = row;
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
          if (columnNumber > (rowFields.length || 1)) {
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
  }

  const canRender = rows.length > 0 && valueField && (rowFields.length > 0 || columnField);
  const ChartComponent = recommendedChart === "line" ? Line : recommendedChart === "pie" ? Pie : Bar;

  return (
    <main className="app-shell">
      <section className="workspace-header">
        <div>
          <p className="eyebrow">Privacy-first Pivot Chart Maker</p>
          <h1>上传文件不离开浏览器，一分钟生成可筛选数据图表。</h1>
        </div>
        <div className="limit-strip">
          <span>CSV / XLSX</span>
          <span>最大 10MB</span>
          <span>最多 10,000 行</span>
        </div>
      </section>

      <section className="workspace-grid">
        <aside className="control-panel">
          <label className="upload-zone">
            <input
              type="file"
              accept=".csv,.xlsx"
              onChange={(event) => handleFile(event.target.files?.[0])}
            />
            <Upload size={28} />
            <strong>{fileName || "上传 CSV / XLSX"}</strong>
            <span>所有解析、聚合和导出都在当前浏览器完成。</span>
          </label>

          {error ? <div className="error-box">{error}</div> : null}

          {columns.length ? (
            <>
              <div className="panel-section">
                <div className="section-title">
                  <FileSpreadsheet size={17} />
                  <span>字段类型</span>
                </div>
                <div className="type-grid">
                  {columns.map((column) => (
                    <label className="field-type" key={column}>
                      <span title={column}>{column}</span>
                      <select
                        value={columnTypes[column]}
                        onChange={(event) =>
                          setColumnTypes((current) => ({ ...current, [column]: event.target.value }))
                        }
                      >
                        {TYPE_OPTIONS.map((type) => (
                          <option value={type} key={type}>
                            {TYPE_LABELS[type]}
                          </option>
                        ))}
                      </select>
                    </label>
                  ))}
                </div>
              </div>

              <div className="panel-section">
                <div className="section-title">
                  <Table2 size={17} />
                  <span>透视表设置</span>
                </div>

                <div className="field-group">
                  <span>行字段</span>
                  <div className="choice-list">
                    {dimensionFields.map((field) => (
                      <label className="checkbox-line" key={field}>
                        <input
                          type="checkbox"
                          checked={rowFields.includes(field)}
                          onChange={() => toggleRowField(field)}
                        />
                        <span>{field}</span>
                      </label>
                    ))}
                  </div>
                </div>

                <label className="field-group">
                  <span>列字段</span>
                  <select value={columnField} onChange={(event) => setColumnField(event.target.value)}>
                    <option value="">不使用列字段</option>
                    {dimensionFields.map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="field-group">
                  <span>数值字段</span>
                  <select value={valueField} onChange={(event) => setValueField(event.target.value)}>
                    {columns.map((field) => (
                      <option key={field} value={field}>
                        {field}
                      </option>
                    ))}
                  </select>
                  {numericFields.length === 0 ? <small>没有明显数字列，可在字段类型里手动改。</small> : null}
                </label>

                <label className="field-group">
                  <span>聚合方式</span>
                  <select value={aggregation} onChange={(event) => setAggregation(event.target.value)}>
                    {Object.entries(AGGREGATIONS).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              <div className="panel-section">
                <div className="section-title">
                  <Filter size={17} />
                  <span>筛选字段</span>
                </div>
                <div className="choice-list compact">
                  {dimensionFields.map((field) => (
                    <label className="checkbox-line" key={field}>
                      <input
                        type="checkbox"
                        checked={filterFields.includes(field)}
                        onChange={() => toggleFilterField(field)}
                      />
                      <span>{field}</span>
                    </label>
                  ))}
                </div>
              </div>
            </>
          ) : null}
        </aside>

        <section className="output-panel">
          {canRender ? (
            <>
              <div className="output-toolbar">
                <div>
                  <div className="status-line">
                    <RefreshCw size={15} />
                    <span>{pivot.filteredCount.toLocaleString()} / {rows.length.toLocaleString()} 行参与计算</span>
                  </div>
                  <strong>{AGGREGATIONS[aggregation]}：{valueField}</strong>
                </div>
                <div className="toolbar-actions">
                  <button type="button" onClick={exportWorkbook} title="导出 Excel 工作簿">
                    <Download size={16} />
                    XLSX
                  </button>
                  <button type="button" onClick={exportPng} title="导出图表 PNG">
                    <Download size={16} />
                    PNG
                  </button>
                  <button type="button" onClick={exportPivotCsv} title="导出透视表 CSV">
                    <Download size={16} />
                    CSV
                  </button>
                </div>
              </div>

              {filterFields.length ? (
                <div className="filters-row">
                  {filterFields.map((field) => {
                    const values = uniqueValues(rows, field, columnTypes[field]);
                    return (
                      <div className="filter-block" key={field}>
                        <strong>{field}</strong>
                        <div className="filter-values">
                          {values.map((value) => (
                            <label className="filter-chip" key={value}>
                              <input
                                type="checkbox"
                                checked={filters[field]?.has(value) || false}
                                onChange={() => toggleFilterValue(field, value)}
                              />
                              <span>{value}</span>
                            </label>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              ) : null}

              <div className="chart-settings">
                <div className="recommendation">
                  {recommendedChart === "line" ? <CalendarDays size={16} /> : recommendedChart === "pie" ? <PieChart size={16} /> : <BarChart3 size={16} />}
                  <span>
                    推荐图表：
                    {recommendedChart === "line" ? "折线图（检测到时间字段）" : recommendedChart === "pie" ? "饼图（占比结构）" : "柱状图（分类维度）"}
                  </span>
                </div>
                <label>
                  图表意图
                  <select value={chartIntent} onChange={(event) => setChartIntent(event.target.value)}>
                    <option value="compare">分类对比</option>
                    <option value="share">占比结构</option>
                  </select>
                </label>
                <label>
                  图表类型
                  <select value={chartOverride} onChange={(event) => setChartOverride(event.target.value)}>
                    <option value="auto">自动推荐</option>
                    <option value="bar">柱状图</option>
                    <option value="line">折线图</option>
                    <option value="pie">饼图</option>
                  </select>
                </label>
              </div>

              <div className="chart-wrap">
                <ChartComponent ref={chartRef} data={chartData} options={chartOptions} />
              </div>

              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {(rowFields.length ? rowFields : ["Row"]).map((header) => (
                        <th key={header}>{header}</th>
                      ))}
                      {pivot.columns.map((column) => (
                        <th key={column}>{column}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {pivot.records.map((record) => (
                      <tr key={record.rowKey}>
                        {(rowFields.length ? record.rowParts : [record.rowKey]).map((part, index) => (
                          <td key={`${record.rowKey}-${index}`}>{part}</td>
                        ))}
                        {pivot.columns.map((column) => (
                          <td key={column} className="metric-cell">
                            {formatMetric(record.values[column], columnTypes[valueField], aggregation)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div className="empty-state">
              <FileSpreadsheet size={42} />
              <h2>上传数据后开始生成透视图表</h2>
              <p>支持脏日期、货币文本和混合字段的初步识别；识别不准时可以手动改列类型。</p>
            </div>
          )}
        </section>
      </section>
    </main>
  );
}
